import asyncio
import html
import io
import json
import logging
import os
import re
import tempfile
from datetime import datetime, timedelta, timezone
from datetime import time as dtime
from zoneinfo import ZoneInfo

from openpyxl import Workbook
from openpyxl.styles import Font
from telegram import InlineKeyboardButton as B
from telegram import InlineKeyboardMarkup as M
from telegram import ReplyKeyboardMarkup, Update
from telegram.error import BadRequest, Forbidden, RetryAfter, TelegramError
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

import config
import db

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
log = logging.getLogger("bot")
TZ = ZoneInfo(config.TZ)

BTN_ORDER, BTN_TRACK, BTN_FAQ = "📝 ثبت سفارش", "🔎 پیگیری سفارش", "❓ سوالات متداول"
BTN_MY, BTN_RATES = "📋 سفارش‌های من", "💹 نرخ کارمزد"
BTN_FIN = "💰 وضعیت مالی من"
BTN_ADMIN, BTN_CANCEL = "🛠 مدیریت", "🔙 انصراف"
PAGE = 8
TOP_N = 3  # تعداد سفارش‌های نمایش‌داده‌شده به‌طور پیش‌فرض در لیست مدیر
DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩٬", "01234567890123456789,")


# ======================= helpers =======================
def fa2en(s: str) -> str:
    return s.translate(DIGITS)


def esc(s) -> str:
    return html.escape(str(s))


def role_of(uid: int):
    if uid in config.ADMIN_IDS:
        return "owner"
    if uid in config.OPERATOR_IDS:
        return "operator"
    if uid in config.VIEWER_IDS:
        return "viewer"
    return None


def is_admin(uid: int) -> bool:
    """هر کسی که یکی از نقش‌های مدیریتی را دارد."""
    return role_of(uid) is not None


def can(uid: int, perm: str) -> bool:
    r = role_of(uid)
    return bool(r) and perm in config.ROLE_PERMS[r]


def all_staff() -> set:
    return set(config.ADMIN_IDS) | config.OPERATOR_IDS | config.VIEWER_IDS


def staff_ids(perm: str) -> list:
    return sorted(i for i in all_staff() if can(i, perm))


def fmt_time(iso: str) -> str:
    return datetime.fromisoformat(iso).astimezone(TZ).strftime("%Y-%m-%d %H:%M")


def get_field(key):
    return next(f for f in config.FORM_FIELDS if f["key"] == key)


def main_kb(uid):
    if is_admin(uid):  # پنل ادمین جدا از پنل مشتری است؛ دکمه‌های پنل همیشه پایین صفحه‌اند
        return admin_reply_kb(uid)
    rows = [[BTN_ORDER, BTN_TRACK], [BTN_MY, BTN_FIN], [BTN_RATES, BTN_FAQ]]
    if is_admin(uid):
        rows.append([BTN_ADMIN])
    return ReplyKeyboardMarkup(rows, resize_keyboard=True)


def validate(f, raw):
    """برمی‌گرداند (مقدار نرمال‌شده, پیام خطا)"""
    required = f.get("required", True)
    if not required and raw == "-":
        return "", None
    kind = f["kind"]
    if kind == "phone":
        v = re.sub(r"[\s\-]", "", fa2en(raw))
        if not re.fullmatch(r"(\+98|0098|0)?9\d{9}", v):
            return None, "شماره موبایل معتبر نیست. نمونه: 09123456789"
        return "0" + v[-10:], None
    if kind == "int":
        v = fa2en(raw).replace(",", "").replace(" ", "")
        if not v.isdigit() or int(v) <= 0:
            return None, "لطفاً فقط عدد (بزرگ‌تر از صفر) وارد کنید."
        return str(int(v)), None
    if kind == "choice":
        if raw not in f["choices"]:
            return None, "لطفاً یکی از گزینه‌های پایین را انتخاب کنید."
        return raw, None
    if not raw or len(raw) > 300:
        return None, "متن خالی یا بیش از حد طولانی است."
    return raw, None


SECRET_MASK = "••••••"


def field_text(f, v, reveal=False):
    """مقدار فیلد برای نمایش؛ فیلدهای رمز فقط برای مجازها نمایش داده می‌شود."""
    if not v:
        return "—"
    if f["kind"] == "secret" and not reveal:
        return SECRET_MASK
    return v


def render_order(o, admin=False, reveal=False):
    data = json.loads(o["data"])
    lines = [
        f"🔖 کد رهگیری: <code>{esc(o['code'])}</code>",
        f"وضعیت: {config.STATUSES.get(o['status'], o['status'])}",
        f"زمان ثبت: {fmt_time(o['created_at'])}",
        "",
    ]
    for f in config.FORM_FIELDS:
        lines.append(f"{esc(f['label'])}: {esc(field_text(f, data.get(f['key']), reveal))}")
    if admin:
        uname = f"@{o['username']}" if o["username"] else "—"
        if o["status"] == "done":
            fee = db.order_fees(o)
            bal = db.customer_summary(o)["balance"]
            lbl = "بدهکار" if bal > 0 else "بستانکار" if bal < 0 else "تسویه"
            lines += ["",
                      f"🧾 کارمزد این سفارش: {fee['commission']:,} ریال "
                      f"({fee['rate']:g}٪ از {fee['trade_amount']:,})",
                      f"💰 حساب کل مشتری: {abs(bal):,} ریال ({lbl})"]
        lines += ["", f"👤 تلگرام: {esc(uname)} | id: <code>{o['user_id']}</code>"]
        n_cust = len(db.customer_orders(o))
        if n_cust > 1:
            lines.append(f"👥 تعداد سفارش‌های این مشتری: {n_cust}")
        sims = db.find_similar(o)
        if sims:
            lines.append("")
            for r, level in sims[:5]:
                tag = "🔴 سفارش تکراری" if level == "exact" else "🟠 سفارش مشابه"
                lines.append(f"{tag}: <code>{esc(r['code'])}</code> ({fmt_time(r['created_at'])})")
        npend = db.count_pending_receipts(user_id=o["user_id"])
        if npend:
            lines.append(f"🧾 رسید پرداخت کل حساب در انتظار بررسی: {npend}")
        notes = db.list_notes(o["id"])
        if notes:
            total = db.count_notes(o["id"])
            lines += ["", f"📝 یادداشت‌های داخلی ({total}):"]
            for n in notes:
                lines.append(f"• {fmt_time(n['at'])}: {esc(n['text'])}")
    return "\n".join(lines)


CLOSED_MSG = {
    "closed": "⛔ پذیرش سفارش در حال حاضر متوقف است. لطفاً بعداً مراجعه کنید.",
    "full": "⛔ ظرفیت پذیرش سفارش تکمیل شده است. لطفاً بعداً مراجعه کنید.",
}


def intake_text() -> str:
    s = db.intake_state()
    if s["reason"] == "closed":
        status = "⛔ بسته (دستی)"
    elif s["reason"] == "full":
        status = "⛔ ظرفیت تکمیل شده"
    else:
        status = "✅ باز"
    cap = "نامحدود" if s["capacity"] == 0 else str(s["capacity"])
    rem = "—" if s["remaining"] is None else str(s["remaining"])
    return (
        "⚙️ <b>پذیرش و ظرفیت</b>\n\n"
        f"وضعیت: {status}\n"
        f"ظرفیت: {cap}\n"
        f"ثبت‌شده تاکنون: {s['used']}\n"
        f"باقی‌مانده: {rem}\n"
        f"نرخ کارمزد پیش‌فرض: {db.get_commission_percent():g}٪\n\n"
        "ℹ️ ظرفیت یعنی سقف کل سفارش‌ها؛ شمارش ریست نمی‌شود و سفارش‌های قبلی هم حساب می‌شوند. "
        "سفارش‌های خطادار در ظرفیت حساب نمی‌شوند."
    )


def intake_kb():
    s = db.intake_state()
    toggle = "⏸ توقف پذیرش" if s["accepting"] else "▶️ شروع پذیرش"
    return M(
        [
            [B(toggle, callback_data="a:captog")],
            [B("🔢 تنظیم ظرفیت", callback_data="a:capset")],
            [B("🧾 نرخ کارمزد", callback_data="a:rate")],
            [B("🏠 پنل", callback_data="a:home")],
        ]
    )


async def show(q, text, kb=None):
    try:
        await q.edit_message_text(text, reply_markup=kb, parse_mode="HTML")
    except BadRequest as e:
        if "not modified" not in str(e).lower():
            raise


# ======================= customer side =======================
async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    u = update.effective_user
    if is_admin(u.id):
        await update.message.reply_text(
            f"👋 سلام! به پنل مدیریت خوش آمدید.\nنقش شما: {config.ROLE_LABELS[role_of(u.id)]}",
            reply_markup=main_kb(u.id),
        )
        return
    db.upsert_user(u.id, u.username)
    await update.message.reply_text(
        "سلام 👋 به ربات ثبت سفارش خوش آمدید.\nاز منوی پایین یکی را انتخاب کنید.",
        reply_markup=main_kb(update.effective_user.id),
    )


async def cancel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    await update.message.reply_text("لغو شد.", reply_markup=main_kb(update.effective_user.id))


async def ask_field(msg, ctx):
    f = config.FORM_FIELDS[ctx.user_data["idx"]]
    text = f["label"]
    if not f.get("required", True):
        text += "\n(اختیاری؛ برای رد شدن «-» بفرستید)"
    rows = [list(f["choices"])] if f["kind"] == "choice" else []
    rows.append([BTN_CANCEL])
    await msg.reply_text(text, reply_markup=ReplyKeyboardMarkup(rows, resize_keyboard=True))


async def start_form(msg, ctx, uid):
    if is_admin(uid):
        await msg.reply_text("مدیرها نیازی به ثبت سفارش ندارند.", reply_markup=main_kb(uid))
        return
    reason = db.intake_state()["reason"]
    if reason:
        await msg.reply_text(CLOSED_MSG[reason], reply_markup=main_kb(uid))
        return
    if db.count_orders_today(uid) >= config.MAX_ORDERS_PER_DAY:
        await msg.reply_text("تعداد سفارش‌های امروز شما به سقف مجاز رسیده است.")
        return
    ctx.user_data.clear()
    ctx.user_data.update(state=("form",), idx=0, form={})
    await msg.reply_text(
        "⚠️. پیام حاوی «کلمه عبور» بعد از دریافت از چت پاک می‌شود."
    )
    await ask_field(msg, ctx)


def confirm_kb():
    return M(
        [
            [B("✅ تأیید و ثبت", callback_data="ok")],
            [B("✏️ اصلاح یک مورد", callback_data="fe"), B("🔁 شروع دوباره", callback_data="redo")],
        ]
    )


async def send_confirm(msg, ctx):
    ctx.user_data["state"] = ("confirm",)
    lines = ["لطفاً اطلاعات را بررسی کنید:", ""]
    for fld in config.FORM_FIELDS:
        lines.append(f"{esc(fld['label'])}: {esc(field_text(fld, ctx.user_data['form'].get(fld['key'])))}")
    await msg.reply_text("\n".join(lines), reply_markup=confirm_kb(), parse_mode="HTML")


async def _wipe(msg):
    """پیام حاوی رمز را از چت مشتری پاک می‌کند (اگر تلگرام اجازه بدهد)."""
    try:
        await msg.delete()
    except TelegramError:
        pass


async def form_step(update, ctx, raw):
    i = ctx.user_data["idx"]
    f = config.FORM_FIELDS[i]
    val, err = validate(f, raw)
    if f["kind"] == "secret":
        await _wipe(update.message)
    if err:
        await update.message.reply_text(err)
        return
    ctx.user_data["form"][f["key"]] = val
    i += 1
    if i < len(config.FORM_FIELDS):
        ctx.user_data["idx"] = i
        await ask_field(update.message, ctx)
        return
    await send_confirm(update.message, ctx)


MY_PAGE = 8
TOP_N = 3  # تعداد سفارش‌های نمایش‌داده‌شده به‌طور پیش‌فرض در لیست مدیر


def my_orders_view(uid, page):
    """فهرست سفارش‌های خود مشتری با کد رهگیری (فقط سفارش‌های همین uid)."""
    rows, total = db.list_user_orders(uid, MY_PAGE, page * MY_PAGE)
    if not total:
        return "هنوز سفارشی ثبت نکرده‌اید.", None
    lines = [f"📋 <b>سفارش‌های شما</b> ({total})", ""]
    kb = []
    for r in rows:
        sym = json.loads(r["data"]).get("symbol", "")
        st = config.STATUSES.get(r["status"], r["status"])
        lines.append(f"🔖 <code>{esc(r['code'])}</code> | {esc(sym)} | {st}")
        kb.append([B(f"{r['code']} | {sym}", callback_data=f"mo:{r['id']}")])
    nav = []
    if page > 0:
        nav.append(B("◀️ قبلی", callback_data=f"my:{page - 1}"))
    if (page + 1) * MY_PAGE < total:
        nav.append(B("بعدی ▶️", callback_data=f"my:{page + 1}"))
    if nav:
        kb.append(nav)
    lines += ["", "برای دیدن جزئیات هر سفارش، روی دکمه‌ی آن بزنید."]
    return "\n".join(lines), M(kb)


def fin_view(uid):
    """وضعیت مالی مشتری: جمع بدهی/پرداخت/مانده، فهرست سفارش‌ها و پیگیری رسیدها."""
    s = db.user_finance(uid)
    recs = db.list_user_receipts(uid, 10)
    if not s["items"] and not recs:
        return "هنوز سفارش یا رسیدی ثبت نشده است.", None
    bal = s["balance"]
    lbl = "بدهکار" if bal > 0 else "بستانکار" if bal < 0 else "تسویه"
    lines = [
        "💰 <b>وضعیت مالی شما</b>",
        "",
        f"جمع بدهی (کارمزدها): {s['debt']:,} ریال",
        f"جمع دریافتی‌ها: {s['paid']:,} ریال",
        f"مانده حساب: {abs(bal):,} ریال ({lbl})",
    ]
    kb = []
    if s["items"]:
        lines += ["", "<b>سفارش‌ها</b>"]
        for r, f in s["items"][:15]:
            sym = json.loads(r["data"]).get("symbol", "")
            fee_txt = f"{f['commission']:,} ریال" if r["status"] == "done" else "هنوز مشخص نشده"
            lines.append(
                f"🔖 <code>{esc(r['code'])}</code> | {esc(sym)}\n"
                f"   کارمزد: {fee_txt}"
            )
    pending = db.count_pending_receipts(user_id=uid)
    if bal > 0 and pending == 0:
        kb.append([B(f"🧾 ارسال رسید پرداخت کل بدهی ({bal:,} ریال)", callback_data="rc:all")])
    if recs:
        lines += ["", "<b>پیگیری پرداخت کل حساب</b>"]
        for r in recs:
            amt = f" | {r['amount']:,} ریال" if r["status"] == "approved" else ""
            lines.append(f"#{r['id']} | {fmt_time(r['created_at'])} | {RECEIPT_CUSTOMER[r['status']]}{amt}")
    if pending:
        lines += ["", "⏳ یک رسید شما در انتظار بررسی است؛ پس از اعلام نتیجه می‌توانید رسید دیگری ارسال کنید."]
    kb.append([B("🔄 به‌روزرسانی", callback_data="fin")])
    return "\n".join(lines), M(kb)


async def menu_text(update, ctx, raw):
    uid = update.effective_user.id
    if is_admin(uid):  # ادمین فقط پنل مدیریت را می‌بیند، نه منوی مشتری
        await update.message.reply_text("🛠 پنل مدیریت", reply_markup=main_kb(uid))
        return
    if raw == BTN_ORDER:
        await start_form(update.message, ctx, uid)
    elif raw == BTN_TRACK:
        ctx.user_data["state"] = ("track",)
        await update.message.reply_text(
            "کد رهگیری را وارد کنید (مثلاً SR-AB12CD):",
            reply_markup=ReplyKeyboardMarkup([[BTN_CANCEL]], resize_keyboard=True),
        )
    elif raw == BTN_MY:
        text, kb = my_orders_view(uid, 0)
        await update.message.reply_text(text, reply_markup=kb, parse_mode="HTML")
    elif raw == BTN_FIN:
        text, kb = fin_view(uid)
        await update.message.reply_text(text, reply_markup=kb, parse_mode="HTML")
    elif raw == BTN_RATES:
        t = db.get_rates_text()
        await update.message.reply_text(
            "💹 نرخ کارمزد سهام‌ها\n\n" + t if t else "نرخ‌ها هنوز اعلام نشده است؛ لطفاً کانال را ببینید.",
            reply_markup=main_kb(uid),
        )
    elif raw == BTN_FAQ:
        rows = db.faq_list()
        if not rows:
            await update.message.reply_text("فعلاً سؤالی ثبت نشده است.")
            return
        kb = M([[B(r["question"], callback_data=f"q:{r['id']}")] for r in rows])
        await update.message.reply_text("سؤال خود را انتخاب کنید:", reply_markup=kb)
    else:
        await update.message.reply_text("از منوی پایین استفاده کنید.", reply_markup=main_kb(uid))


async def on_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    raw = update.message.text.strip()
    if raw == BTN_CANCEL:
        return await cancel(update, ctx)
    if is_admin(uid) and raw in ADMIN_BTN_DATA:  # دکمه‌های پنل پایین صفحه
        if ctx.user_data.get("state") is not None:  # عملیات نیمه‌کاره لغو و دکمه‌های پنل برگردانده شود
            await update.message.reply_text("عملیات قبلی لغو شد.", reply_markup=main_kb(uid))
        ctx.user_data.clear()
        return await admin_button(update, ctx, raw)
    st = ctx.user_data.get("state")
    if st is None:
        return await menu_text(update, ctx, raw)
    kind = st[0]
    if kind == "form":
        return await form_step(update, ctx, raw)
    if kind == "form_edit":
        f = config.FORM_FIELDS[st[1]]
        val, err = validate(f, raw)
        if f["kind"] == "secret":
            await _wipe(update.message)
        if err:
            return await update.message.reply_text(err)
        ctx.user_data["form"][f["key"]] = val
        return await send_confirm(update.message, ctx)
    if kind == "track":
        o = db.get_order_by_code(fa2en(raw).upper().replace(" ", ""))
        ctx.user_data.clear()
        # مشتری فقط سفارش خودش را می‌بیند
        if o and o["user_id"] == uid:
            await update.message.reply_text(
                render_order(o), parse_mode="HTML", reply_markup=main_kb(uid)
            )
        else:
            await update.message.reply_text("سفارشی با این کد برای شما پیدا نشد.", reply_markup=main_kb(uid))
        return
    if kind == "receipt":
        if len(raw) < 4 or len(raw) > 300:
            return await update.message.reply_text(
                "عکس رسید را بفرستید، یا شماره پیگیری بانکی را بنویسید (۴ تا ۳۰۰ کاراکتر)."
            )
        return await submit_receipt(update, ctx, "text", text=raw)
    if kind.startswith("a_") and is_admin(uid):
        return await admin_text(update, ctx, raw, st)
    await cancel(update, ctx)


# ======================= admin side =======================
ADMIN_ITEMS = [
    ("view", "📋 سفارش‌ها", "a:l:top:0"),
    ("view", "👥 کاربران و حساب‌ها", "a:users:0"),
    ("view", "🔍 جستجو", "a:search"),
    ("export", "📊 خروجی اکسل", "a:x"),
    ("faq", "❓ مدیریت سوالات", "a:faq"),
    ("rates", "💹 نرخ کارمزد سهام‌ها", "a:rt"),
    ("cap", "⚙️ پذیرش و ظرفیت", "a:cap"),
    ("broadcast", "📣 اطلاعیه گروهی", "a:bc"),
    ("backup", "💾 بکاپ و بازیابی", "a:bk"),
]
ADMIN_BTN_DATA = {t: d for _, t, d in ADMIN_ITEMS}  # متن دکمه ← داده‌ی عملیات


def admin_kb(uid):
    btns = [B(t, callback_data=d) for perm, t, d in ADMIN_ITEMS if can(uid, perm)]
    return M([btns[i : i + 2] for i in range(0, len(btns), 2)])


def admin_reply_kb(uid):
    """دکمه‌های پنل مدیریت، ثابت پایین صفحه (بر اساس سطح دسترسی)."""
    btns = [t for perm, t, d in ADMIN_ITEMS if can(uid, perm)]
    rows = [btns[i : i + 2] for i in range(0, len(btns), 2)]
    return ReplyKeyboardMarkup(rows, resize_keyboard=True)


class _MsgQuery:
    """شبیه‌ساز CallbackQuery تا دکمه‌های پایین صفحه از همان منطق پنل استفاده کنند."""

    def __init__(self, update, data):
        self.message = update.message
        self.from_user = update.effective_user
        self.data = data

    async def answer(self, *a, **k):
        pass

    async def edit_message_text(self, text, reply_markup=None, parse_mode=None):
        await self.message.reply_text(text, reply_markup=reply_markup, parse_mode=parse_mode)

    async def edit_message_reply_markup(self, reply_markup=None):
        pass


async def admin_button(update, ctx, raw):
    q = _MsgQuery(update, ADMIN_BTN_DATA[raw])
    await admin_cb(q, ctx, q.data[2:].split(":"))


def order_button(o):
    data = json.loads(o["data"])
    label = f"{o['code']} | {data.get('symbol', '')} | {config.STATUSES[o['status']].split()[0]}"
    return [B(label, callback_data=f"a:o:{o['id']}")]


def order_kb(o, uid):
    rows = []
    if can(uid, "status"):
        status_row = [
            B(label, callback_data=f"a:s:{o['id']}:{key}")
            for key, label in config.STATUSES.items()
            if key != o["status"]
        ]
        rows += [status_row[i : i + 2] for i in range(0, len(status_row), 2)]
    act_row = []
    if can(uid, "edit"):
        act_row.append(B("✏️ ویرایش اطلاعات", callback_data=f"a:e:{o['id']}"))
    if can(uid, "note"):
        act_row.append(B("📝 یادداشت", callback_data=f"a:n:{o['id']}"))
    if act_row:
        rows.append(act_row)
    if can(uid, "pay") and db.count_pending_receipts(user_id=o["user_id"]):
        rows.append([B("🧾 بررسی رسید کل حساب", callback_data=f"a:rcl:{o['user_id']}")])
    info_row = [B("👤 سابقه مشتری", callback_data=f"a:hist:{o['id']}")]
    if can(uid, "pay"):
        info_row.insert(0, B("💳 پرداخت و فاکتور", callback_data=f"a:pay:{o['id']}"))
    rows.append(info_row)
    rows.append([B("⬅️ لیست", callback_data="a:l:top:0"), B("🏠 پنل", callback_data="a:home")])
    return M(rows)


def pay_kb(o):
    oid = o["id"]
    rows = []
    if o["status"] == "done":
        rows.append([B("🧾 ارسال صورتحساب کل مشتری", callback_data=f"a:inv:{oid}")])
    rows.append([B("💰 ثبت مبلغ دریافتی", callback_data=f"a:pa:{oid}")])
    rows.append([B("⬅️ سفارش", callback_data=f"a:o:{oid}")])
    return M(rows)


async def notify_admins(ctx, o):
    kb = M([[B("مشاهده / مدیریت", callback_data=f"a:o:{o['id']}")]])
    for aid in staff_ids("status"):
        try:
            await ctx.bot.send_message(
                aid, "📥 سفارش جدید\n\n" + render_order(o, admin=True, reveal=can(aid, "secret")), parse_mode="HTML", reply_markup=kb
            )
        except TelegramError:
            log.exception("notify admin %s failed", aid)


async def broadcast_admins(ctx, text):
    """هشدارهای سیستمی (بکاپ، تکمیل ظرفیت) فقط برای مالک‌ها."""
    for aid in staff_ids("backup"):
        try:
            await ctx.bot.send_message(aid, text)
        except TelegramError:
            log.exception("broadcast to admin %s failed", aid)


async def run_broadcast(ctx, admin_id, text):
    """ارسال اطلاعیه به همه‌ی کاربران فعال، با فاصله بین پیام‌ها (محدودیت تلگرام)."""
    sent = failed = blocked = 0
    try:
        for tid in db.broadcast_targets(exclude=all_staff()):
            try:
                await ctx.bot.send_message(tid, text)
                sent += 1
            except RetryAfter as e:
                await asyncio.sleep(e.retry_after + 1)
                try:
                    await ctx.bot.send_message(tid, text)
                    sent += 1
                except TelegramError:
                    failed += 1
            except Forbidden:
                db.mark_blocked(tid)
                blocked += 1
            except TelegramError:
                failed += 1
            await asyncio.sleep(0.05)
    except Exception:
        log.exception("broadcast crashed")
    finally:
        ctx.bot_data["bc_running"] = False
    log.info("broadcast by %s: sent=%s blocked=%s failed=%s", admin_id, sent, blocked, failed)
    try:
        await ctx.bot.send_message(
            admin_id,
            f"📣 ارسال اطلاعیه تمام شد.\n✅ ارسال‌شده: {sent}\n🚫 بلاک‌کرده: {blocked}\n⚠️ ناموفق: {failed}",
        )
    except TelegramError:
        log.exception("broadcast report failed")


RECEIPT_LABELS = {"approved": "تأیید", "rejected": "رد"}
RECEIPT_CUSTOMER = {"pending": "⏳ در انتظار بررسی", "approved": "✅ تأیید شد", "rejected": "❌ رد شد"}


async def notify_receipt(ctx, rid, only=None):
    """ارسال رسید پرداخت کل حساب برای مالک‌ها."""
    r = db.get_receipt(rid)
    o = db.get_order(r["order_id"])
    d = json.loads(o["data"])
    summary = db.user_finance(r["user_id"])
    lines = [
        f"🧾 <b>رسید پرداخت کل حساب #{rid}</b>",
        f"مشتری: {esc(d.get('full_name') or '—')} | {esc(d.get('phone') or '—')}",
        f"جمع بدهی: {summary['debt']:,} | پرداختی قبلی: {summary['paid']:,} | مانده کل: {summary['balance']:,} ریال",
    ]
    if r["text"]:
        lines.append(f"شماره پیگیری / توضیح: {esc(r['text'][:200])}")
    for dup in db.find_duplicate_receipts(r):
        lines.append(
            f"🔴 رسید مشابه قبلاً در سیستم ثبت شده است "
            f"(#{dup['id']}، {RECEIPT_LABELS.get(dup['status'], 'در انتظار')})"
        )
    text = "\n".join(lines)
    kb = M(
        [[B("✅ تأیید پرداخت", callback_data=f"a:rca:{rid}"), B("❌ رد رسید", callback_data=f"a:rcr:{rid}")]]
    )
    for aid in only or staff_ids("pay"):
        try:
            if r["kind"] == "photo":
                await ctx.bot.send_photo(aid, r["file_id"], caption=text, parse_mode="HTML", reply_markup=kb)
            elif r["kind"] == "document":
                await ctx.bot.send_document(aid, r["file_id"], caption=text, parse_mode="HTML", reply_markup=kb)
            else:
                await ctx.bot.send_message(aid, text, parse_mode="HTML", reply_markup=kb)
        except TelegramError:
            log.exception("notify receipt %s to %s failed", rid, aid)


async def submit_receipt(update, ctx, kind, file_id=None, file_uid=None, text=None):
    uid = update.effective_user.id
    st = ctx.user_data.get("state")
    if not st or st[0] != "receipt" or st[1] != uid:
        ctx.user_data.clear()
        return await update.message.reply_text("درخواست ارسال رسید معتبر نیست؛ از وضعیت مالی اقدام کنید.", reply_markup=main_kb(uid))
    summary = db.user_finance(uid)
    if summary["balance"] <= 0:
        ctx.user_data.clear()
        return await update.message.reply_text("در حال حاضر بدهی پرداخت‌نشده‌ای ندارید.", reply_markup=main_kb(uid))
    if db.count_pending_receipts(user_id=uid):
        ctx.user_data.clear()
        return await update.message.reply_text("یک رسید شما در انتظار بررسی است؛ لطفاً تا اعلام نتیجه صبر کنید.", reply_markup=main_kb(uid))
    completed = [r for r, _ in summary["items"] if r["status"] == "done"]
    if not completed:
        ctx.user_data.clear()
        return await update.message.reply_text("سفارش تکمیل‌شده‌ای برای ثبت پرداخت وجود ندارد.", reply_markup=main_kb(uid))
    ctx.user_data.clear()
    try:
        rid = db.add_receipt(completed[0]["id"], uid, kind, file_id, file_uid, text)
    except ValueError:
        return await update.message.reply_text(
            "یک رسید شما در انتظار بررسی است؛ لطفاً تا اعلام نتیجه صبر کنید.",
            reply_markup=main_kb(uid),
        )
    await update.message.reply_text(
        "✅ رسید پرداخت کل بدهی شما دریافت شد و برای بررسی ارسال شد. پس از تأیید مدیر، مانده حساب به‌روزرسانی می‌شود.",
        reply_markup=main_kb(uid),
    )
    await notify_receipt(ctx, rid)


async def finish_approval(msg, ctx, rid, amount, uid):
    res = db.approve_receipt(rid, amount, uid)
    if not res:
        return await msg.reply_text("این رسید قبلاً بررسی شده است.")
    label = "بدهکار" if res["balance"] > 0 else "بستانکار" if res["balance"] < 0 else "تسویه"
    await msg.reply_text(
        f"✅ رسید پرداخت کل حساب #{rid} تأیید شد.\n"
        f"مبلغ ثبت‌شده: {amount:,} ریال\n"
        f"جمع پرداختی‌های مشتری: {res['paid']:,} ریال\n"
        f"مانده کل حساب: {abs(res['balance']):,} ریال ({label})"
    )
    try:
        await ctx.bot.send_message(
            res["user_id"],
            f"✅ پرداخت کل بدهی شما تأیید شد.\nمبلغ ثبت‌شده: {amount:,} ریال\n"
            f"مانده حساب: {abs(res['balance']):,} ریال ({label})",
        )
    except TelegramError:
        pass


ACT_PERM = {
    "home": "view", "l": "view", "o": "view", "search": "view", "hist": "view",
    "users": "view", "usr": "view", "uhist": "view",
    "s": "status", "e": "edit", "ef": "edit", "n": "note",
    "pay": "pay", "pa": "pay", "inv": "pay",
    "rca": "pay", "rcf": "pay", "rcm": "pay", "rcr": "pay", "rcl": "pay",
    "x": "export", "xp": "export", "xe": "export",
    "bk": "backup", "bkn": "backup", "bkr": "backup", "bkc": "backup", "bky": "backup",
    "cap": "cap", "captog": "cap", "capset": "cap", "rate": "cap",
    "faq": "faq", "fq": "faq", "fqadd": "faq", "fqq": "faq", "fqe": "faq", "fqd": "faq", "fqdy": "faq",
    "bc": "broadcast", "bcy": "broadcast", "bcc": "broadcast",
    "rt": "rates", "rte": "rates", "rtn": "rates",
}  # fmt: skip
STATE_PERM = {
    "a_search": "view", "a_edit": "edit", "a_note": "note", "a_paid": "pay", "a_rc_amount": "pay",
    "a_symrate": "rates",
    "a_cap_set": "cap", "a_rate": "cap", "a_faq_add_q": "faq", "a_faq_add_a": "faq",
    "a_faq_edit": "faq", "a_restore": "backup", "a_restore_confirm": "backup",
    "a_bc": "broadcast", "a_bc_confirm": "broadcast", "a_rates": "rates",
}  # fmt: skip
NO_ACCESS = "⛔ دسترسی شما برای این بخش کافی نیست."


def backup_filename() -> str:
    return f"backup-{datetime.now(TZ):%Y%m%d-%H%M}.db"


async def auto_backup(ctx: ContextTypes.DEFAULT_TYPE):
    """بکاپ شبانه؛ برای همه‌ی مدیرها ارسال می‌شود."""
    try:
        data = db.make_backup().getvalue()
    except Exception:
        log.exception("auto backup failed")
        await broadcast_admins(ctx, "⚠️ بکاپ خودکار شبانه ناموفق بود. لطفاً بکاپ دستی بگیرید و لاگ را بررسی کنید.")
        return
    name = backup_filename()
    for aid in staff_ids("backup"):
        try:
            await ctx.bot.send_document(
                aid, io.BytesIO(data), filename=name, caption="💾 بکاپ خودکار شبانه — این فایل را نگه دارید."
            )
        except TelegramError:
            log.exception("send backup to admin %s failed", aid)


PERIODS = {0: "همه‌ی زمان‌ها", 1: "امروز", 7: "۷ روز اخیر", 30: "۳۰ روز اخیر"}


def since_iso(days: int):
    """days=0 یعنی همه؛ 1 یعنی امروز (از نیمه‌شب به وقت TZ)؛ 7 یعنی ۷ روز اخیر."""
    if not days:
        return None
    start = datetime.now(TZ).replace(hour=0, minute=0, second=0, microsecond=0) - timedelta(days=days - 1)
    return start.astimezone(timezone.utc).isoformat(timespec="seconds")


def _put(ws, row, bold=False):
    ws.append(row)
    for c in ws[ws.max_row]:
        if isinstance(c.value, str) and c.value.startswith("="):
            c.data_type = "s"  # متنِ مشتری هرگز به‌عنوان فرمول اجرا نشود
        if bold:
            c.font = Font(bold=True)


def build_xlsx(status=None, days=0) -> io.BytesIO:
    orders = db.list_orders_filtered(status, since_iso(days))
    rate = db.get_commission_percent()
    wb = Workbook()
    ws = wb.active
    ws.title = "orders"
    ws.sheet_view.rightToLeft = True
    header = ["کد رهگیری", "وضعیت", "زمان ثبت", "آخرین تغییر", "یوزرنیم", "آیدی تلگرام"]
    header += [f["label"] for f in config.FORM_FIELDS]
    header += ["کارمزد / بدهی (ریال)", "مبلغ دریافتی (ریال)", "مانده (ریال)"]
    _put(ws, header, bold=True)

    by_status, by_symbol, by_day = {}, {}, {}
    live = {"n": 0, "debt": 0, "paid": 0}
    for o in orders:
        data = json.loads(o["data"])
        fee = db.order_fees(o, rate)
        vals = []
        for f in config.FORM_FIELDS:
            v = data.get(f["key"], "")
            if f["kind"] == "secret":
                v = SECRET_MASK if v else ""
            vals.append(int(v) if f["kind"] == "int" and str(v).isdigit() else v)
        _put(
            ws,
            [
                o["code"],
                config.STATUSES.get(o["status"], o["status"]),
                fmt_time(o["created_at"]),
                fmt_time(o["updated_at"]),
                o["username"] or "",
                o["user_id"],
            ]
            + vals
            + [fee["commission"], fee["paid"], fee["balance"]],
        )
        s_ = by_status.setdefault(o["status"], [0, 0])
        s_[0] += 1
        s_[1] += fee["commission"]
        if o["status"] == config.INACTIVE_STATUS:
            continue
        live["n"] += 1
        live["debt"] += fee["commission"]
        live["paid"] += fee["paid"]
        sym = (data.get("symbol") or "—").strip() or "—"
        e = by_symbol.setdefault(sym, [0, 0, 0])
        e[0] += 1
        if data.get("side") == "خرید":
            e[1] += 1
        elif data.get("side") == "فروش":
            e[2] += 1
        d = datetime.fromisoformat(o["created_at"]).astimezone(TZ).strftime("%Y-%m-%d")
        e = by_day.setdefault(d, [0, 0])
        e[0] += 1
        e[1] += fee["commission"]

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 18

    rp = wb.create_sheet("report")
    rp.sheet_view.rightToLeft = True
    label = config.STATUSES.get(status, "همه‌ی وضعیت‌ها") if status else "همه‌ی وضعیت‌ها"
    _put(rp, ["فیلتر", f"بازه: {PERIODS.get(days, days)}", f"وضعیت: {label}"], True)
    rp.append([])
    _put(rp, ["خلاصه (بدون سفارش‌های خطادار)", "مقدار"], True)
    for k, v in [
        ("تعداد سفارش", live["n"]),
        ("جمع بدهی / کارمزد (ریال)", live["debt"]),
        ("جمع مبلغ دریافتی (ریال)", live["paid"]),
        ("مانده (ریال)", live["debt"] - live["paid"]),
    ]:
        _put(rp, [k, v])
    rp.append([])
    _put(rp, ["بر اساس وضعیت", "تعداد", "کارمزد (ریال)"], True)
    for key, (n, c) in by_status.items():
        _put(rp, [config.STATUSES.get(key, key), n, c])
    rp.append([])
    _put(rp, ["بر اساس نماد", "تعداد", "تعداد خرید", "تعداد فروش"], True)
    for sym, (n, b_, s2) in sorted(by_symbol.items(), key=lambda kv: -kv[1][0]):
        _put(rp, [sym, n, b_, s2])
    rp.append([])
    _put(rp, ["بر اساس روز", "تعداد", "کارمزد (ریال)"], True)
    for d, (n, c) in sorted(by_day.items()):
        _put(rp, [d, n, c])
    for col in rp.columns:
        rp.column_dimensions[col[0].column_letter].width = 26

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


async def admin_cb(q, ctx, p):
    uid = q.from_user.id
    act = p[0]
    need = ACT_PERM.get(act)
    if not need or not can(uid, need):
        await q.message.reply_text(NO_ACCESS)
        return

    if act == "home":
        ctx.user_data.clear()
        await show(q, "پنل مدیریت", admin_kb(uid))

    elif act == "l":  # a:l:<status|all>:<page>
        status, page = p[1], int(p[2])
        top = status == "top"  # نمای پیش‌فرض: فقط جدیدترین‌ها؛ با «همه» کل فهرست
        if top:
            rows, total = db.list_orders(None, TOP_N, 0)
        else:
            rows, total = db.list_orders(None if status == "all" else status, PAGE, page * PAGE)
        kb = [order_button(o) for o in rows]
        nav = []
        if not top and page > 0:
            nav.append(B("◀️ قبلی", callback_data=f"a:l:{status}:{page - 1}"))
        if not top and (page + 1) * PAGE < total:
            nav.append(B("بعدی ▶️", callback_data=f"a:l:{status}:{page + 1}"))
        if nav:
            kb.append(nav)
        kb.append([B("همه", callback_data="a:l:all:0")])
        for key, label in config.STATUSES.items():
            kb.append([B(label, callback_data=f"a:l:{key}:0")])
        kb.append([B("🏠 پنل", callback_data="a:home")])
        if top:
            head = f"📋 آخرین سفارش‌ها ({len(rows)} از {total}) — برای دیدن همه، «همه» را بزنید"
        else:
            title = "همه" if status == "all" else config.STATUSES[status]
            head = f"📋 سفارش‌ها — {title} ({total})"
        await show(q, head, M(kb))

    elif act == "users":
        page = max(0, int(p[1]) if len(p) > 1 else 0)
        rows, total = db.list_users(PAGE, page * PAGE, exclude=all_staff())
        kb = []
        for user in rows:
            finance = db.user_finance(user["user_id"])
            bal = finance["balance"]
            label = "بدهکار" if bal > 0 else "بستانکار" if bal < 0 else "تسویه"
            username = f"@{user['username']}" if user["username"] else f"ID {user['user_id']}"
            pending = db.count_pending_receipts(user_id=user["user_id"])
            receipt_mark = " 🧾" if pending else ""
            kb.append([B(
                f"{username} | {abs(bal):,} ریال ({label}){receipt_mark}",
                callback_data=f"a:usr:{user['user_id']}",
            )])
        nav = []
        if page > 0:
            nav.append(B("◀️ قبلی", callback_data=f"a:users:{page - 1}"))
        if (page + 1) * PAGE < total:
            nav.append(B("بعدی ▶️", callback_data=f"a:users:{page + 1}"))
        if nav:
            kb.append(nav)
        kb.append([B("🏠 پنل مدیریت", callback_data="a:home")])
        await show(q, f"👥 <b>کاربران و وضعیت حساب‌ها</b>\nتعداد کل: {total}\nصفحه {page + 1}", M(kb))

    elif act == "usr":
        user_id = int(p[1])
        user = db.get_user(user_id)
        if not user:
            return await q.message.reply_text("کاربر پیدا نشد.")
        finance = db.user_finance(user_id)
        rows = [r for r, _ in finance["items"]]
        orders = db.customer_orders(rows[0]) if rows else []
        if not rows:
            # کاربری که هنوز سفارش ندارد
            orders = []
        profile = json.loads(orders[0]["data"]) if orders else {}
        name = profile.get("full_name") or "—"
        phone = profile.get("phone") or "—"
        balance = finance["balance"]
        status = "بدهکار" if balance > 0 else "بستانکار" if balance < 0 else "تسویه"
        pending = db.count_pending_receipts(user_id=user_id)
        receipts = db.list_user_receipts(user_id, 5)
        lines = [
            "👤 <b>پرونده مالی مشتری</b>",
            f"نام: {esc(name)}",
            f"موبایل: {esc(phone)}",
            f"شناسه تلگرام: <code>{user_id}</code>",
            f"نام کاربری: @{esc(user['username'] or '—')}",
            "",
            f"تعداد سفارش‌ها: {len(orders)}",
            f"جمع بدهی: <b>{finance['debt']:,} ریال</b>",
            f"جمع پرداختی: <b>{finance['paid']:,} ریال</b>",
            f"مانده: <b>{abs(balance):,} ریال ({status})</b>",
            f"رسید در انتظار بررسی: {pending}",
        ]
        if receipts:
            lines += ["", "<b>آخرین پرداخت‌ها و رسیدها</b>"]
            for receipt in receipts:
                amount = f" | {receipt['amount']:,} ریال" if receipt["amount"] is not None else ""
                lines.append(
                    f"#{receipt['id']} | {fmt_time(receipt['created_at'])} | "
                    f"{RECEIPT_LABELS.get(receipt['status'], receipt['status'])}{amount}"
                )
        kb = []
        if pending:
            kb.append([B("🧾 بررسی رسیدهای پرداخت", callback_data=f"a:rcl:{user_id}")])
        if orders:
            kb.append([B("📋 مشاهده سفارش‌های مشتری", callback_data=f"a:uhist:{user_id}")])
            done = next((r for r in orders if r["status"] == "done"), None)
            if done and can(uid, "pay"):
                kb.append([B("📨 ارسال صورتحساب کل", callback_data=f"a:inv:{done['id']}")])
        kb.append([B("⬅️ فهرست کاربران", callback_data="a:users:0")])
        await show(q, "\\n".join(lines), M(kb))

    elif act == "uhist":
        user_id = int(p[1])
        user = db.get_user(user_id)
        if not user:
            return await q.message.reply_text("کاربر پیدا نشد.")
        account_ids = db.account_user_ids(user_id)
        orders = []
        for account_id in account_ids:
            found, _ = db.list_orders(None, 500, 0, user_id=account_id)
            orders.extend(found)
        orders.sort(key=lambda r: r["id"], reverse=True)
        kb = [order_button(row) for row in orders[:30]]
        kb.append([B("⬅️ پرونده مالی", callback_data=f"a:usr:{user_id}")])
        await show(q, f"📋 سفارش‌های مشتری ({len(orders)})", M(kb))

    elif act == "o":
        o = db.get_order(int(p[1]))
        if not o:
            return await show(q, "سفارش پیدا نشد.", admin_kb(uid))
        await show(q, render_order(o, admin=True, reveal=can(uid, "secret")), order_kb(o, uid))

    elif act == "s":  # a:s:<id>:<status>
        oid, status = int(p[1]), p[2]
        if status not in config.STATUSES:
            return
        o = db.get_order(oid)
        if not o:
            return await show(q, "سفارش پیدا نشد.", admin_kb(uid))
        if status == "done":
            amount = db.order_total(o)
            if amount <= 0:
                return await q.message.reply_text(
                    "⚠️ مبلغ سفارش از اطلاعات ثبت‌شده قابل محاسبه نیست. مقدار «قدرت خرید» یا «تعداد و قیمت» سفارش را بررسی کنید."
                )
            res = db.complete_order(oid, amount, uid)
            if not res:
                return await q.message.reply_text(
                    "⚠️ نرخ کارمزد این نماد ثبت نشده است. از «💹 نرخ کارمزد سهام‌ها ← ➕ نرخ نماد» نرخ را ثبت کنید و دوباره وضعیت را تغییر دهید."
                )
            o = db.get_order(oid)
        else:
            db.set_status(oid, status, uid)
            o = db.get_order(oid)
        await show(q, render_order(o, admin=True, reveal=can(uid, "secret")), order_kb(o, uid))
        try:
            note = f"وضعیت سفارش {o['code']} تغییر کرد: {config.STATUSES[status]}"
            if status == "error":
                note += "\n\nلطفاً نام کاربری و کلمه عبور را بررسی کنید و سفارش را دوباره ثبت کنید."
            await ctx.bot.send_message(o["user_id"], note)
        except TelegramError:
            pass  # مشتری ربات را بلاک کرده یا چت بسته است

    elif act == "e":
        oid = int(p[1])
        kb = [[B(f["label"], callback_data=f"a:ef:{oid}:{f['key']}")] for f in config.FORM_FIELDS]
        kb.append([B("⬅️ بازگشت", callback_data=f"a:o:{oid}")])
        o = db.get_order(oid)
        if not o:
            return await show(q, "سفارش پیدا نشد.", admin_kb(uid))
        data = json.loads(o["data"])
        cur = "\n".join(
            f"{esc(f['label'])}: {esc(field_text(f, data.get(f['key']), can(uid, 'secret')))}"
            for f in config.FORM_FIELDS
        )
        await show(q, f"کدام فیلد ویرایش شود؟\n\n<b>مقادیر فعلی:</b>\n{cur}", M(kb))

    elif act == "ef":
        oid, key = int(p[1]), p[2]
        ctx.user_data["state"] = ("a_edit", oid, key)
        f = get_field(key)
        o = db.get_order(oid)
        cur = field_text(f, json.loads(o["data"]).get(key) if o else "", can(uid, "secret"))
        await q.message.reply_text(
            f"مقدار فعلی «{f['label']}»:\n{cur}\n\nمقدار جدید را بفرستید:",
            reply_markup=ReplyKeyboardMarkup([[BTN_CANCEL]], resize_keyboard=True),
        )

    elif act == "n":
        ctx.user_data["state"] = ("a_note", int(p[1]))
        await q.message.reply_text(
            "متن یادداشت داخلی را بفرستید (مشتری آن را نمی‌بیند):",
            reply_markup=ReplyKeyboardMarkup([[BTN_CANCEL]], resize_keyboard=True),
        )

    elif act == "search":
        ctx.user_data["state"] = ("a_search",)
        await q.message.reply_text(
            "کد رهگیری، نام، شماره موبایل، نماد یا یوزرنیم را بفرستید:",
            reply_markup=ReplyKeyboardMarkup([[BTN_CANCEL]], resize_keyboard=True),
        )

    elif act == "x":
        kb = M(
            [
                [B("📅 همه‌ی زمان‌ها", callback_data="a:xp:0")],
                [
                    B("امروز", callback_data="a:xp:1"),
                    B("۷ روز اخیر", callback_data="a:xp:7"),
                    B("۳۰ روز اخیر", callback_data="a:xp:30"),
                ],
                [B("🏠 پنل", callback_data="a:home")],
            ]
        )
        await show(q, "📊 <b>خروجی اکسل</b>\n\nبازه‌ی زمانی را انتخاب کنید:", kb)

    elif act == "xp":  # a:xp:<days>
        days = int(p[1])
        if days not in PERIODS:
            return
        rows = [[B("همه‌ی وضعیت‌ها", callback_data=f"a:xe:{days}:all")]]
        sts = list(config.STATUSES.items())
        for i in range(0, len(sts), 2):
            rows.append([B(lb, callback_data=f"a:xe:{days}:{k}") for k, lb in sts[i : i + 2]])
        rows.append([B("⬅️ بازگشت", callback_data="a:x")])
        await show(q, f"بازه: {PERIODS[days]}\nوضعیت سفارش‌ها را انتخاب کنید:", M(rows))

    elif act == "xe":  # a:xe:<days>:<status|all>
        days, status = int(p[1]), p[2]
        if days not in PERIODS or (status != "all" and status not in config.STATUSES):
            return
        await q.message.reply_document(
            document=build_xlsx(None if status == "all" else status, days),
            filename=f"orders-{datetime.now(TZ):%Y%m%d-%H%M}.xlsx",
            caption="📊 برگه‌ی «orders» لیست سفارش‌ها و برگه‌ی «report» گزارش خلاصه است.",
        )

    # ---- customer history ----
    elif act == "hist":
        o = db.get_order(int(p[1]))
        if not o:
            return await show(q, "سفارش پیدا نشد.", admin_kb(uid))
        d = json.loads(o["data"])
        rows = db.customer_orders(o)
        s = db.customer_summary(o)
        n_err = sum(1 for r in rows if r["status"] == "error")
        bal = s["balance"]
        bal_lbl = "بدهکار" if bal > 0 else "بستانکار" if bal < 0 else "تسویه"
        text = (
            "👤 <b>پرونده‌ی مشتری</b>\n\n"
            f"نام: {esc(d.get('full_name') or '—')}\n"
            f"موبایل: {esc(d.get('phone') or '—')}\n\n"
            f"تعداد کل سفارش‌ها: {len(rows)}\n"
            f"🆕 ثبت‌شده: {len(rows) - n_err}\n"
            f"⚠️ دارای خطا: {n_err}\n\n"
            "💰 <b>حساب مالی</b> (بدون سفارش‌های خطادار)\n"
            f"جمع بدهی (کارمزدها): {s['debt']:,} ریال\n"
            f"جمع پرداختی‌ها: {s['paid']:,} ریال\n"
            f"مانده حساب: {abs(bal):,} ریال ({bal_lbl})\n\n"
            f"اولین سفارش: {fmt_time(rows[-1]['created_at'])}"
        )
        kb = [order_button(r) for r in rows[:15]]
        kb.append([B("⬅️ بازگشت به سفارش", callback_data=f"a:o:{o['id']}")])
        await show(q, text, M(kb))

    # ---- payment & commission ----
    elif act == "pay":
        o = db.get_order(int(p[1]))
        if not o:
            return await show(q, "سفارش پیدا نشد.", admin_kb(uid))
        await show(q, render_order(o, admin=True, reveal=can(uid, "secret")), pay_kb(o))

    elif act == "inv":
        oid = int(p[1])
        o = db.get_order(oid)
        if not o or o["status"] != "done":
            return await q.message.reply_text("برای این سفارش هنوز صورتحساب قابل ارسال نیست.")

        # صورتحساب تجمیعی حساب تلگرام مشتری؛ پرداخت‌های تأییدشده‌ی دفتر کل هم لحاظ می‌شوند.
        finance = db.user_finance(o["user_id"])
        completed_count = sum(1 for row, _ in finance["items"] if row["status"] == "done")
        if not completed_count:
            return await q.message.reply_text("برای این مشتری سفارش تکمیل‌شده‌ای جهت صدور صورتحساب وجود ندارد.")

        total_debt = finance["debt"]
        total_paid = finance["paid"]
        balance = finance["balance"]
        balance_label = "بدهکار" if balance > 0 else "بستانکار" if balance < 0 else "تسویه"
        invoice = (
            "🧾 <b>صورتحساب کل مشتری</b>\n"
            f"تعداد سفارش‌های انجام‌شده: <b>{completed_count}</b>\n"
            f"جمع کل بدهی (کارمزدها): <b>{total_debt:,} ریال</b>\n"
            f"جمع پرداختی‌ها: {total_paid:,} ریال\n"
            "━━━━━━━━━━━━━━\n"
            f"<b>مانده بدهی مشتری: {abs(balance):,} ریال ({balance_label})</b>"
        )
        try:
            await ctx.bot.send_message(o["user_id"], invoice, parse_mode="HTML")
            await q.message.reply_text("صورتحساب کل مشتری با موفقیت ارسال شد.")
        except TelegramError:
            await q.message.reply_text("ارسال صورتحساب ناموفق بود؛ احتمالاً مشتری ربات را مسدود کرده است.")

    elif act == "pa":
        ctx.user_data["state"] = ("a_paid", int(p[1]))
        await q.message.reply_text(
            "مبلغ دریافتی از مشتری را به ریال و فقط با عدد بفرستید (۰ = هیچ):",
            reply_markup=ReplyKeyboardMarkup([[BTN_CANCEL]], resize_keyboard=True),
        )

    elif act == "rate":
        ctx.user_data["state"] = ("a_rate",)
        await q.message.reply_text(
            f"نرخ فعلی کارمزد: {db.get_commission_percent():g}٪\n"
            "نرخ پیش‌فرض را به درصد بفرستید (مثلاً 0.5). این نرخ زمانی استفاده می‌شود "
            "که برای نماد سفارش نرخ اختصاصی ثبت نشده باشد.",
            reply_markup=ReplyKeyboardMarkup([[BTN_CANCEL]], resize_keyboard=True),
        )

    # ---- payment receipts ----
    elif act in ("rca", "rcf", "rcm", "rcr"):
        rid = int(p[1])
        r = db.get_receipt(rid)
        if not r:
            return await q.message.reply_text("رسید پیدا نشد.")
        if r["status"] != "pending":
            return await q.message.reply_text(f"این رسید قبلاً {RECEIPT_LABELS[r['status']]} شده است.")
        o = db.get_order(r["order_id"])
        summary = db.user_finance(r["user_id"])
        if act == "rca":
            rows = []
            if summary["balance"] > 0:
                rows.append([B(f"✅ ثبت کل مانده ({summary['balance']:,} ریال)", callback_data=f"a:rcf:{rid}")])
            rows.append([B("✏️ مبلغ دیگر", callback_data=f"a:rcm:{rid}")])
            await q.message.reply_text(
                f"تأیید رسید کل حساب #{rid}\n"
                f"جمع بدهی: {summary['debt']:,} | پرداختی قبلی: {summary['paid']:,} | مانده کل: {summary['balance']:,} ریال\n\n"
                "چه مبلغی از این رسید ثبت شود؟",
                reply_markup=M(rows),
            )
        elif act == "rcf":
            if summary["balance"] <= 0:
                return await q.message.reply_text("مانده کل حساب صفر یا منفی است؛ «مبلغ دیگر» را بزنید.")
            try:
                await q.edit_message_reply_markup(None)
            except TelegramError:
                pass
            await finish_approval(q.message, ctx, rid, summary["balance"], uid)
        elif act == "rcm":
            ctx.user_data["state"] = ("a_rc_amount", rid)
            await q.message.reply_text(
                f"مبلغ دریافتی رسید #{rid} را به ریال و فقط با عدد بفرستید:",
                reply_markup=ReplyKeyboardMarkup([[BTN_CANCEL]], resize_keyboard=True),
            )
        else:  # rcr
            ok, _ = db.reject_receipt(rid, uid)
            try:
                await q.edit_message_reply_markup(None)
            except TelegramError:
                pass
            if not ok:
                return await q.message.reply_text("این رسید قبلاً بررسی شده است.")
            await q.message.reply_text(f"❌ رسید #{rid} رد شد.")
            try:
                await ctx.bot.send_message(
                    r["user_id"],
                    "❌ رسید پرداخت کل حساب شما پذیرفته نشد. لطفاً پس از بررسی، رسید صحیح را از بخش «وضعیت مالی» دوباره ارسال کنید.",
                )
            except TelegramError:
                pass

    elif act == "rcl":
        pend = db.list_pending_receipts(user_id=int(p[1]))
        if not pend:
            return await q.message.reply_text("رسید پرداخت در انتظاری برای این مشتری نیست.")
        for r in pend:
            await notify_receipt(ctx, r["id"], only=[uid])

    # ---- commission rates text (shown to customers) ----
    elif act == "rt":
        cur = db.get_rates_text()
        sr = "\n".join(f"• {esc(sym)}: {rate:g}٪" for sym, rate in db.list_symbol_rates()) or "هنوز نرخ نمادی ثبت نشده است."
        kb = M([
            [B("➕ ثبت نرخ نماد", callback_data="a:rtn")],
            [B("✏️ ویرایش متن نرخ‌های نمایشی", callback_data="a:rte")],
            [B("🏠 پنل", callback_data="a:home")],
        ])
        await show(
            q,
            "💹 <b>نرخ کارمزد نمادها</b> (برای محاسبه خودکار فاکتور)\n\n" + sr
            + "\n\n<b>متنی که مشتری می‌بیند:</b>\n" + (esc(cur) if cur else "هنوز ثبت نشده."),
            kb,
        )

    elif act == "rtn":
        ctx.user_data["state"] = ("a_symrate",)
        await q.message.reply_text(
            "نام نماد و درصد کارمزد را بفرستید. مثال: فولاد 2",
            reply_markup=ReplyKeyboardMarkup([[BTN_CANCEL]], resize_keyboard=True),
        )

    elif act == "rte":
        ctx.user_data["state"] = ("a_rates",)
        await q.message.reply_text(
            "متن کامل نرخ‌ها را بفرستید (جایگزین متن قبلی می‌شود):",
            reply_markup=ReplyKeyboardMarkup([[BTN_CANCEL]], resize_keyboard=True),
        )

    # ---- broadcast ----
    elif act == "bc":
        total, active = db.user_counts(all_staff())
        ctx.user_data["state"] = ("a_bc",)
        await q.message.reply_text(
            "📣 متن اطلاعیه را بفرستید.\n"
            f"به {active} کاربر ارسال می‌شود ({total - active} نفر ربات را بلاک کرده‌اند).\n"
            "قبل از ارسال، پیش‌نمایش را می‌بینید و تأیید می‌گیرم.",
            reply_markup=ReplyKeyboardMarkup([[BTN_CANCEL]], resize_keyboard=True),
        )

    elif act == "bcc":
        ctx.user_data.clear()
        await show(q, "ارسال اطلاعیه لغو شد.", admin_kb(uid))

    elif act == "bcy":
        st = ctx.user_data.get("state")
        if not st or st[0] != "a_bc_confirm":
            return await show(q, "درخواست منقضی شده؛ دوباره از «اطلاعیه گروهی» شروع کنید.", admin_kb(uid))
        if ctx.bot_data.get("bc_running"):
            return await show(q, "یک ارسال دیگر در حال انجام است؛ کمی بعد دوباره تلاش کنید.", admin_kb(uid))
        text = st[1]
        ctx.user_data.clear()
        ctx.bot_data["bc_running"] = True
        ctx.application.create_task(run_broadcast(ctx, uid, text))
        await show(q, "⏳ ارسال شروع شد. پس از پایان، گزارش را برایتان می‌فرستم.", admin_kb(uid))

    # ---- backup / restore ----
    elif act == "bk":
        ctx.user_data.clear()
        kb = M(
            [
                [B("📥 دریافت بکاپ همین حالا", callback_data="a:bkn")],
                [B("♻️ بازیابی از فایل بکاپ", callback_data="a:bkr")],
                [B("🏠 پنل", callback_data="a:home")],
            ]
        )
        await show(
            q,
            "💾 <b>بکاپ و بازیابی</b>\n\n"
            f"بکاپ خودکار هر شب ساعت {esc(config.BACKUP_TIME)} برای مدیرها ارسال می‌شود.\n"
            "فایل‌های بکاپ حاوی اطلاعات مشتریان است؛ آن‌ها را جای امن نگه دارید.",
            kb,
        )

    elif act == "bkn":
        await q.message.reply_document(document=db.make_backup(), filename=backup_filename())

    elif act == "bkr":
        ctx.user_data["state"] = ("a_restore",)
        await q.message.reply_text(
            "⚠️ فایل بکاپ (.db) را بفرستید.\n"
            "قبل از جایگزینی، خلاصه‌ی فایل را می‌بینید و تأیید می‌گیرم. "
            "(حداکثر حجم ۲۰ مگابایت)",
            reply_markup=ReplyKeyboardMarkup([[BTN_CANCEL]], resize_keyboard=True),
        )

    elif act == "bkc":
        st = ctx.user_data.get("state")
        if st and st[0] == "a_restore_confirm":
            try:
                os.remove(st[1])
            except OSError:
                pass
        ctx.user_data.clear()
        await show(q, "بازیابی لغو شد. اطلاعات فعلی دست‌نخورده ماند.", admin_kb(uid))

    elif act == "bky":
        st = ctx.user_data.get("state")
        if not st or st[0] != "a_restore_confirm":
            return await show(q, "درخواست منقضی شده؛ دوباره از «بکاپ و بازیابی» شروع کنید.", admin_kb(uid))
        path = st[1]
        ctx.user_data.clear()
        try:
            _, err = db.validate_backup(path)
            if err:
                return await show(q, f"❌ {esc(err)}", admin_kb(uid))
            safety = db.restore_backup(path)
        except Exception:
            log.exception("restore failed")
            return await show(q, "❌ بازیابی ناموفق بود؛ اطلاعات فعلی تغییری نکرد.", admin_kb(uid))
        finally:
            try:
                os.remove(path)
            except OSError:
                pass
        log.warning("database restored by admin %s; safety copy: %s", uid, safety)
        await show(
            q,
            "✅ بازیابی انجام شد.\n"
            f"از وضعیت قبلی یک کپی ایمنی کنار دیتابیس ذخیره شد:\n<code>{esc(os.path.basename(safety))}</code>",
            admin_kb(uid),
        )

    # ---- intake / capacity ----
    elif act == "cap":
        await show(q, intake_text(), intake_kb())

    elif act == "captog":
        db.set_accepting(not db.intake_state()["accepting"])
        await show(q, intake_text(), intake_kb())

    elif act == "capset":
        ctx.user_data["state"] = ("a_cap_set",)
        await q.message.reply_text(
            "ظرفیت جدید را به صورت عدد بفرستید (۰ = بدون محدودیت).\n"
            "ظرفیت، سقف کل سفارش‌هاست و سفارش‌های قبلی (به‌جز خطادارها) هم شمرده می‌شوند.",
            reply_markup=ReplyKeyboardMarkup([[BTN_CANCEL]], resize_keyboard=True),
        )

    # ---- FAQ management ----
    elif act == "faq":
        rows = db.faq_list()
        kb = [[B(r["question"], callback_data=f"a:fq:{r['id']}")] for r in rows]
        kb.append([B("➕ افزودن سؤال", callback_data="a:fqadd")])
        kb.append([B("🏠 پنل", callback_data="a:home")])
        await show(q, "مدیریت سوالات متداول", M(kb))

    elif act == "fq":
        r = db.faq_get(int(p[1]))
        if not r:
            return await show(q, "پیدا نشد.", admin_kb(uid))
        kb = M(
            [
                [B("✏️ سؤال", callback_data=f"a:fqq:{r['id']}"), B("✏️ پاسخ", callback_data=f"a:fqe:{r['id']}")],
                [B("🗑 حذف", callback_data=f"a:fqd:{r['id']}")],
                [B("⬅️ بازگشت", callback_data="a:faq")],
            ]
        )
        await show(q, f"<b>{esc(r['question'])}</b>\n\n{esc(r['answer'])}", kb)

    elif act == "fqadd":
        ctx.user_data["state"] = ("a_faq_add_q",)
        await q.message.reply_text(
            "متن سؤال را بفرستید:", reply_markup=ReplyKeyboardMarkup([[BTN_CANCEL]], resize_keyboard=True)
        )

    elif act in ("fqq", "fqe"):
        col = "question" if act == "fqq" else "answer"
        ctx.user_data["state"] = ("a_faq_edit", int(p[1]), col)
        await q.message.reply_text(
            "متن جدید را بفرستید:", reply_markup=ReplyKeyboardMarkup([[BTN_CANCEL]], resize_keyboard=True)
        )

    elif act == "fqd":
        kb = M(
            [[B("✅ بله، حذف شود", callback_data=f"a:fqdy:{p[1]}"), B("خیر", callback_data=f"a:fq:{p[1]}")]]
        )
        await show(q, "این سؤال حذف شود؟", kb)

    elif act == "fqdy":
        db.faq_delete(int(p[1]))
        q.data = "a:faq"
        await admin_cb(q, ctx, ["faq"])


async def admin_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    if is_admin(uid):
        await update.message.reply_text("🛠 پنل مدیریت", reply_markup=main_kb(uid))


async def admin_text(update, ctx, raw, st):
    uid = update.effective_user.id
    kind = st[0]
    msg = update.message
    need = STATE_PERM.get(kind)
    if need and not can(uid, need):
        ctx.user_data.clear()
        return await msg.reply_text(NO_ACCESS, reply_markup=main_kb(uid))

    if kind == "a_search":
        ctx.user_data.clear()
        rows = db.search_orders(fa2en(raw).strip())
        if not rows:
            return await msg.reply_text("موردی پیدا نشد.", reply_markup=main_kb(uid))
        await msg.reply_text("نتایج:", reply_markup=main_kb(uid))
        await msg.reply_text(f"{len(rows)} مورد", reply_markup=M([order_button(o) for o in rows]))

    elif kind == "a_edit":
        _, oid, key = st
        val, err = validate(get_field(key), raw)
        if get_field(key)["kind"] == "secret":
            await _wipe(msg)
        if err:
            return await msg.reply_text(err)
        db.set_field(oid, key, val, uid)
        ctx.user_data.clear()
        await msg.reply_text("ذخیره شد ✅", reply_markup=main_kb(uid))
        await msg.reply_text("بازگشت به سفارش:", reply_markup=M([[B("مشاهده سفارش", callback_data=f"a:o:{oid}")]]))

    elif kind == "a_bc_confirm":
        await msg.reply_text("برای ارسال یا لغو، دکمه‌های زیر پیش‌نمایش را بزنید، یا «🔙 انصراف».")

    elif kind in ("a_restore", "a_restore_confirm"):
        await msg.reply_text("لطفاً فایل بکاپ (.db) را به صورت فایل (Document) بفرستید، یا «🔙 انصراف» را بزنید.")

    elif kind == "a_note":
        if not raw or len(raw) > 400:
            return await msg.reply_text("یادداشت خالی یا بیش از ۴۰۰ کاراکتر است.")
        oid = st[1]
        db.add_note(oid, uid, raw)
        ctx.user_data.clear()
        await msg.reply_text("یادداشت ذخیره شد ✅", reply_markup=main_kb(uid))
        await msg.reply_text(
            "بازگشت به سفارش:", reply_markup=M([[B("مشاهده سفارش", callback_data=f"a:o:{oid}")]])
        )

    elif kind == "a_cap_set":
        v = fa2en(raw).replace(",", "").strip()
        if not v.isdigit() or int(v) > 100000:
            return await msg.reply_text("لطفاً یک عدد صحیح (۰ تا ۱۰۰٬۰۰۰) بفرستید.")
        db.set_capacity(int(v))
        ctx.user_data.clear()
        await msg.reply_text("ذخیره شد ✅", reply_markup=main_kb(uid))
        await msg.reply_text(
            intake_text(), parse_mode="HTML", reply_markup=intake_kb()
        )

    elif kind == "a_faq_add_q":
        ctx.user_data["state"] = ("a_faq_add_a", raw)
        await msg.reply_text("حالا متن پاسخ را بفرستید:")

    elif kind == "a_faq_add_a":
        db.faq_add(st[1], raw)
        ctx.user_data.clear()
        await msg.reply_text("سؤال اضافه شد ✅", reply_markup=main_kb(uid))

    elif kind == "a_faq_edit":
        db.faq_update(st[1], st[2], raw)
        ctx.user_data.clear()
        await msg.reply_text("ذخیره شد ✅", reply_markup=main_kb(uid))

    elif kind == "a_paid":
        v = fa2en(raw).replace(",", "").replace(" ", "")
        if not v.isdigit() or int(v) > 10**13:
            return await msg.reply_text("لطفاً فقط عدد (ریال، بدون حروف) بفرستید.")
        oid = st[1]
        db.set_paid_amount(oid, int(v), uid)
        ctx.user_data.clear()
        await msg.reply_text("مبلغ دریافتی ثبت شد ✅", reply_markup=main_kb(uid))
        await msg.reply_text(
            "بازگشت:", reply_markup=M([[B("💳 پرداخت و فاکتور", callback_data=f"a:pay:{oid}")]])
        )

    elif kind == "a_symrate":
        parts = fa2en(raw).replace("٪", "").replace("%", "").replace("٫", ".").split()
        try:
            pct, sym = float(parts[-1]), " ".join(parts[:-1]).strip()
        except (ValueError, IndexError):
            pct, sym = -1, ""
        if not sym or not 0 < pct <= 100:
            return await msg.reply_text("فرمت صحیح: نام نماد و درصد. مثال: فولاد 2")
        db.set_symbol_rate(sym, pct)
        ctx.user_data.clear()
        await msg.reply_text(f"نرخ کارمزد {esc(sym)} با مقدار {pct:g}٪ ذخیره شد ✅", parse_mode="HTML", reply_markup=main_kb(uid))

    elif kind == "a_rate":
        v = fa2en(raw).replace("٪", "").replace("%", "").replace("٫", ".").replace("/", ".").strip()
        try:
            pct = float(v)
        except ValueError:
            pct = -1
        if not 0 <= pct <= 100:
            return await msg.reply_text("لطفاً عددی بین ۰ تا ۱۰۰ بفرستید (مثلاً 0.5).")
        db.set_commission_percent(pct)
        ctx.user_data.clear()
        await msg.reply_text("ذخیره شد ✅", reply_markup=main_kb(uid))
        await msg.reply_text(intake_text(), parse_mode="HTML", reply_markup=intake_kb())

    elif kind == "a_rc_amount":
        v = fa2en(raw).replace(",", "").replace(" ", "")
        if not v.isdigit() or not 0 < int(v) <= 10**13:
            return await msg.reply_text("لطفاً فقط عدد (ریال، بزرگ‌تر از صفر) بفرستید.")
        ctx.user_data.clear()
        await msg.reply_text("در حال ثبت…", reply_markup=main_kb(uid))
        await finish_approval(msg, ctx, st[1], int(v), uid)

    elif kind == "a_rates":
        if not raw or len(raw) > 3500:
            return await msg.reply_text("متن خالی یا بیش از ۳۵۰۰ کاراکتر است.")
        db.set_rates_text(raw)
        ctx.user_data.clear()
        await msg.reply_text("ذخیره شد ✅ مشتری‌ها از «💹 نرخ کارمزد» همین متن را می‌بینند.", reply_markup=main_kb(uid))

    elif kind == "a_bc":
        if not raw or len(raw) > 3500:
            return await msg.reply_text("متن خالی یا بیش از ۳۵۰۰ کاراکتر است.")
        ctx.user_data["state"] = ("a_bc_confirm", raw)
        _, active = db.user_counts(all_staff())
        kb = M(
            [[B(f"✅ ارسال به {active} نفر", callback_data="a:bcy"), B("❌ لغو", callback_data="a:bcc")]]
        )
        await msg.reply_text("پیش‌نمایش اطلاعیه:\n\n" + raw, reply_markup=kb)


async def on_document(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    st = ctx.user_data.get("state")
    if st and st[0] == "receipt":
        doc = update.message.document
        if not ((doc.mime_type or "").startswith("image/") or doc.mime_type == "application/pdf"):
            return await update.message.reply_text("فقط عکس یا فایل PDF رسید قابل قبول است.")
        if doc.file_size and doc.file_size > 10 * 1024 * 1024:
            return await update.message.reply_text("حجم فایل بیش از ۱۰ مگابایت است.")
        cap = (update.message.caption or "").strip()[:300] or None
        return await submit_receipt(update, ctx, "document", doc.file_id, doc.file_unique_id, cap)
    if not can(uid, "backup") or not st or st[0] != "a_restore":
        return
    doc = update.message.document
    if doc.file_size and doc.file_size > 20 * 1024 * 1024:
        return await update.message.reply_text("حجم فایل بیش از ۲۰ مگابایت است و قابل دریافت نیست.")
    fd, path = tempfile.mkstemp(suffix=".db")
    os.close(fd)
    try:
        tg_file = await doc.get_file()
        await tg_file.download_to_drive(path)
        info, err = db.validate_backup(path)
    except TelegramError:
        log.exception("download backup failed")
        os.remove(path)
        return await update.message.reply_text("دریافت فایل ناموفق بود؛ دوباره تلاش کنید.")
    if err:
        os.remove(path)
        return await update.message.reply_text(f"❌ {err}\nفایل دیگری بفرستید یا «🔙 انصراف» را بزنید.")
    ctx.user_data["state"] = ("a_restore_confirm", path)
    last = fmt_time(info["last_order"]) if info["last_order"] else "—"
    kb = M([[B("✅ بله، جایگزین شود", callback_data="a:bky"), B("❌ انصراف", callback_data="a:bkc")]])
    await update.message.reply_text(
        "📦 خلاصه‌ی فایل بکاپ:\n"
        f"تعداد سفارش‌ها: {info['orders']}\n"
        f"تعداد سوالات متداول: {info['faq']}\n"
        f"آخرین سفارش: {last}\n\n"
        "⚠️ با تأیید، همه‌ی اطلاعات فعلی با این فایل جایگزین می‌شود "
        "(از وضعیت فعلی یک کپی ایمنی گرفته می‌شود).",
        reply_markup=kb,
    )


async def on_photo(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    st = ctx.user_data.get("state")
    if not st or st[0] != "receipt":
        if not is_admin(uid):
            await update.message.reply_text(
                "برای ارسال رسید کل بدهی، از بخش «💰 وضعیت مالی من» اقدام کنید."
            )
        return
    ph = update.message.photo[-1]
    cap = (update.message.caption or "").strip()[:300] or None
    await submit_receipt(update, ctx, "photo", ph.file_id, ph.file_unique_id, cap)


# ======================= callbacks (shared) =======================
async def on_cb(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    d, uid = q.data, q.from_user.id

    if d.startswith("a:"):
        if not is_admin(uid):
            return
        return await admin_cb(q, ctx, d[2:].split(":"))

    if d.startswith("q:"):
        r = db.faq_get(int(d[2:]))
        if r:
            await q.message.reply_text(f"❓ {r['question']}\n\n{r['answer']}")
        return

    if d.startswith("my:"):
        text, kb = my_orders_view(uid, max(int(d[3:]), 0))
        return await show(q, text, kb)

    if d.startswith("mo:"):
        o = db.get_order(int(d[3:]))
        # مشتری فقط سفارش خودش را می‌بیند
        if o and o["user_id"] == uid:
            await q.message.reply_text(render_order(o), parse_mode="HTML")
        return

    if d == "fin":
        if is_admin(uid):
            return
        text, kb = fin_view(uid)
        return await show(q, text, kb)

    if d == "rc:all":
        if is_admin(uid):
            return
        summary = db.user_finance(uid)
        if summary["balance"] <= 0:
            return await q.message.reply_text("در حال حاضر بدهی پرداخت‌نشده‌ای ندارید.")
        if db.count_pending_receipts(user_id=uid):
            return await q.message.reply_text("یک رسید شما در انتظار بررسی است؛ لطفاً تا اعلام نتیجه صبر کنید.")
        ctx.user_data.clear()
        ctx.user_data["state"] = ("receipt", uid)
        return await q.message.reply_text(
            f"🧾 ارسال رسید پرداخت کل حساب\nمبلغ بدهی فعلی: {summary['balance']:,} ریال\n\n"
            "عکس یا فایل (PDF) رسید پرداخت کل بدهی را بفرستید؛ "
            "می‌توانید شماره پیگیری را در توضیح (کپشن) بنویسید.\n"
            "اگر عکس ندارید، فقط شماره پیگیری بانکی را به‌صورت متن بفرستید.",
            reply_markup=ReplyKeyboardMarkup([[BTN_CANCEL]], resize_keyboard=True),
        )

    if d == "fe":
        if ctx.user_data.get("state") != ("confirm",):
            return await q.edit_message_reply_markup(None)
        rows = [
            [B(f["label"], callback_data=f"fef:{i}")] for i, f in enumerate(config.FORM_FIELDS)
        ]
        rows.append([B("⬅️ بازگشت", callback_data="fb")])
        return await q.edit_message_reply_markup(M(rows))

    if d == "fb":
        if ctx.user_data.get("state") != ("confirm",):
            return await q.edit_message_reply_markup(None)
        return await q.edit_message_reply_markup(confirm_kb())

    if d.startswith("fef:"):
        i = int(d[4:])
        if ctx.user_data.get("state") != ("confirm",) or not 0 <= i < len(config.FORM_FIELDS):
            return await q.edit_message_reply_markup(None)
        ctx.user_data.update(state=("form_edit", i), idx=i)
        await q.edit_message_reply_markup(None)
        return await ask_field(q.message, ctx)

    if d == "redo":
        await q.edit_message_reply_markup(None)
        return await start_form(q.message, ctx, uid)

    if d == "ok":
        if ctx.user_data.get("state") != ("confirm",) or not ctx.user_data.get("form"):
            return await q.edit_message_reply_markup(None)
        data = ctx.user_data["form"]
        oid, code = db.create_order(uid, q.from_user.username, data)
        ctx.user_data.clear()
        await q.edit_message_reply_markup(None)
        if oid is None:  # code == reason: "closed" | "full"
            await q.message.reply_text(CLOSED_MSG[code], reply_markup=main_kb(uid))
            return
        await q.message.reply_text(
            f"✅ سفارش شما دریافت شد.\n\nکد رهگیری: {code}\n"
            "برای پیگیری از منوی «پیگیری سفارش» استفاده کنید.",
            reply_markup=main_kb(uid),
        )
        await notify_admins(ctx, db.get_order(oid))
        st = db.intake_state()
        if st["capacity"] and st["remaining"] == 0:
            await broadcast_admins(
                ctx,
                f"🔔 ظرفیت پذیرش ({st['capacity']} سفارش) تکمیل شد و سفارش‌گیری متوقف شد.\n"
                "برای ادامه از «مدیریت ← پذیرش و ظرفیت» ظرفیت جدید تنظیم کنید.",
            )


def main():
    db.init_db()
    builder = Application.builder().token(config.BOT_TOKEN)
    if config.PROXY_URL:
        builder = builder.proxy(config.PROXY_URL).get_updates_proxy(config.PROXY_URL)
    app = builder.build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("cancel", cancel))
    app.add_handler(CommandHandler("admin", admin_cmd))
    app.add_handler(CallbackQueryHandler(on_cb))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_text))
    app.add_handler(MessageHandler(filters.Document.ALL, on_document))
    app.add_handler(MessageHandler(filters.PHOTO, on_photo))
    if app.job_queue is None:
        log.warning("job-queue غیرفعال است؛ بکاپ خودکار اجرا نمی‌شود. نصب: pip install 'python-telegram-bot[job-queue]'")
    else:
        hh, mm = (int(x) for x in config.BACKUP_TIME.split(":"))
        app.job_queue.run_daily(auto_backup, time=dtime(hh, mm, tzinfo=TZ), name="auto_backup")
    log.info("bot started")
    app.run_polling()


if __name__ == "__main__":
    main()