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
logging.getLogger("httpx").setLevel(logging.WARNING)
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


def fmt_num(v) -> str:
    """عدد با جداکننده‌ی سه‌رقمی: 100000000 ← 100,000,000"""
    try:
        return f"{int(str(v).replace(',', '').strip()):,}"
    except (TypeError, ValueError):
        return str(v)


def money(n) -> str:
    """مبلغ به تومان با جداکننده‌ی سه‌رقمی: 2000000 ← 2,000,000 تومان"""
    return f"{int(n):,} تومان"


FA_DIGITS = str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹")


def fa_num(n) -> str:
    return str(n).translate(FA_DIGITS)


def get_field(key):
    return next(f for f in config.FORM_FIELDS if f["key"] == key)


def main_kb(uid):
    if is_admin(uid):  # پنل ادمین جدا از پنل مشتری است؛ دکمه‌های پنل همیشه پایین صفحه‌اند
        return admin_reply_kb(uid)
    rows = [[BTN_ORDER, BTN_TRACK], [BTN_MY, BTN_FIN], [BTN_RATES, BTN_FAQ]]
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
    if f["kind"] == "int":
        return fmt_num(v)
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
                      f"🧾 کارمزد این سفارش: {fee['commission']:,} تومان "
                      f"({fee['rate']:g}٪ از {fee['trade_amount']:,})",
                      f"💰 حساب کل مشتری: {abs(bal):,} تومان ({lbl})"]
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


DEFAULT_PREFIX = "↩️ همان قبلی:"


def default_for(ctx, f):
    """مقدار قبلی/ذخیره‌شده‌ی یک فیلد برای پیشنهاد (رمز هرگز پیشنهاد نمی‌شود)."""
    if f["kind"] == "secret":
        return None
    ud = ctx.user_data
    return (ud.get("form") or {}).get(f["key"]) or (ud.get("defaults") or {}).get(f["key"]) or None


def use_default(ctx, f, raw):
    if raw.startswith(DEFAULT_PREFIX):
        return default_for(ctx, f) or raw
    return raw


def next_idx(ctx, i):
    skip = ctx.user_data.get("skip") or set()
    while i < len(config.FORM_FIELDS) and config.FORM_FIELDS[i]["key"] in skip:
        i += 1
    return i


async def ask_field(msg, ctx):
    f = config.FORM_FIELDS[ctx.user_data["idx"]]
    text = f["label"]
    if not f.get("required", True):
        text += "\n(اختیاری؛ برای رد شدن «-» بفرستید)"
    rows = []
    default = default_for(ctx, f)
    if default:
        shown = field_text(f, default)
        text += f"\nمقدار قبلی: {shown}"
        if f["kind"] != "choice":
            rows.append([f"{DEFAULT_PREFIX} {shown[:40]}"])
    if f["kind"] == "choice":
        rows.append(list(f["choices"]))
    rows.append([BTN_CANCEL])
    await msg.reply_text(text, reply_markup=ReplyKeyboardMarkup(rows, resize_keyboard=True))


def order_block_msg(uid):
    """اگر ثبت سفارش ممکن نیست، متن دلیل؛ وگرنه None."""
    if is_admin(uid):
        return "مدیرها نیازی به ثبت سفارش ندارند."
    reason = db.intake_state()["reason"]
    if reason:
        return CLOSED_MSG[reason]
    if db.count_orders_today(uid) >= config.MAX_ORDERS_PER_DAY:
        return "تعداد سفارش‌های امروز شما به سقف مجاز رسیده است."
    return None


async def begin_form(msg, ctx, form=None, skip=(), defaults=None):
    """شروع فرم؛ form مقادیر از پیش‌پرشده، skip کلیدهایی که دوباره پرسیده نمی‌شوند، defaults مقادیر پیشنهادی."""
    ctx.user_data.clear()
    ctx.user_data.update(
        state=("form",), idx=0, form=dict(form or {}), skip=set(skip), defaults=dict(defaults or {})
    )
    await msg.reply_text("⚠️ پیام حاوی «کلمه عبور» بعد از دریافت از چت پاک می‌شود.")
    i = next_idx(ctx, 0)
    if i >= len(config.FORM_FIELDS):
        return await send_confirm(msg, ctx)
    ctx.user_data["idx"] = i
    await ask_field(msg, ctx)


async def start_form(msg, ctx, uid):
    blocked = order_block_msg(uid)
    if blocked:
        await msg.reply_text(blocked, reply_markup=main_kb(uid))
        return
    ctx.user_data.clear()
    profile = db.get_profile(uid)
    if not profile:
        return await begin_form(msg, ctx)
    lines = ["📌 <b>اطلاعات ذخیره‌شده‌ی شما</b>", ""]
    for f in config.FORM_FIELDS:
        if f.get("remember") and profile.get(f["key"]):
            lines.append(f"{esc(f['label'])}: {esc(field_text(f, profile[f['key']]))}")
    lines += ["", "می‌توانید با همین اطلاعات سفارش بدهید، آن‌ها را ویرایش کنید یا فرم را خالی پر کنید."]
    kb = M(
        [
            [B("✅ استفاده از اطلاعات ذخیره‌شده", callback_data="pf:use")],
            [B("✏️ ویرایش اطلاعات ذخیره‌شده", callback_data="pf:edit")],
            [B("🆕 ورود اطلاعات جدید", callback_data="pf:new"), B("🗑 حذف اطلاعات", callback_data="pf:del")],
        ]
    )
    await msg.reply_text("\n".join(lines), reply_markup=kb, parse_mode="HTML")


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
    raw = use_default(ctx, f, raw)
    val, err = validate(f, raw)
    if f["kind"] == "secret":
        await _wipe(update.message)
    if err:
        await update.message.reply_text(err)
        return
    ctx.user_data["form"][f["key"]] = val
    i = next_idx(ctx, i + 1)
    if i < len(config.FORM_FIELDS):
        ctx.user_data["idx"] = i
        await ask_field(update.message, ctx)
        return
    await send_confirm(update.message, ctx)


MY_PAGE = 8


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


async def send_fin(msg, uid):
    """وضعیت مالی مشتری: فقط صورتحساب‌های پرداخت‌نشده‌ای که مدیر فرستاده؛ زیر هر کدام دکمه‌ی ارسال رسید."""
    invs = db.list_unpaid_invoices(uid)
    if not invs:
        await msg.reply_text("✅ شما بدهی یا صورتحساب پرداخت‌نشده‌ای ندارید.", reply_markup=main_kb(uid))
        return
    total = sum(i["amount"] - i["paid_amount"] for i in invs)
    await msg.reply_text(
        "💰 <b>وضعیت مالی شما</b>\n"
        f"صورتحساب‌های پرداخت‌نشده: {len(invs)}\n"
        f"جمع مبلغ قابل پرداخت: {money(total)}",
        parse_mode="HTML",
        reply_markup=main_kb(uid),
    )
    for inv in invs[:10]:
        rem = inv["amount"] - inv["paid_amount"]
        lines = [f"🧾 <b>صورتحساب #{inv['id']}</b> | {fmt_time(inv['created_at'])}", "", esc(inv["text"]), "",
                 f"مبلغ قابل پرداخت: {money(rem)}"]
        if inv["paid_amount"]:
            lines.append(f"پرداخت‌شده تا اینجا: {money(inv['paid_amount'])}")
        last = db.last_invoice_receipt(inv["id"])
        kb = None
        if last and last["status"] == "pending":
            lines += ["", "⏳ رسید شما در انتظار بررسی مدیر است."]
        else:
            if last and last["status"] == "rejected":
                lines += ["", "❌ آخرین رسید شما رد شد؛ لطفاً رسید صحیح را ارسال کنید."]
            kb = M([[B("🧾 ارسال رسید پرداخت", callback_data=f"rc:{inv['id']}")]])
        await msg.reply_text("\n".join(lines), parse_mode="HTML", reply_markup=kb)
    if len(invs) > 10:
        await msg.reply_text(f"… و {len(invs) - 10} صورتحساب دیگر. پس از تسویه‌ی موارد بالا نمایش داده می‌شوند.")


def rates_lines():
    """فهرست نرخ نمادها؛ همین فهرست هم به مشتری و هم به مدیر نشان داده می‌شود."""
    lines = [f"• {sym}: {rate:g}٪" for sym, rate in db.list_symbol_rates()]
    default = db.get_commission_percent()
    if default > 0:
        if lines:
            lines.append("")
            lines.append(f"سایر نمادها: {default:g}٪")
        else:
            lines.append(f"نرخ کارمزد همه‌ی نمادها: {default:g}٪")
    return lines


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
        await send_fin(update.message, uid)
    elif raw == BTN_RATES:
        lines = rates_lines()
        await update.message.reply_text(
            "💹 نرخ کارمزد سهام‌ها\n\n" + "\n".join(lines) if lines else "نرخ‌ها هنوز اعلام نشده است؛ لطفاً کانال را ببینید.",
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
        raw = use_default(ctx, f, raw)
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
        rows.append([B("🧾 صدور صورتحساب مشتری", callback_data=f"a:inv:{oid}")])
    rows.append([B("💰 ثبت مبلغ دریافتی", callback_data=f"a:pa:{oid}")])
    rows.append([B("⬅️ سفارش", callback_data=f"a:o:{oid}")])
    return M(rows)


def invoice_default_text(name, count, amount):
    return config.INVOICE_TEMPLATE.format(
        name=name or "مشتری گرامی", count=fa_num(count), amount=f"{amount:,}"
    )


def invoice_preview(inv):
    """پیش‌نمایش صورتحساب برای مالک (متن ساده؛ همان‌طور که به مشتری می‌رسد)."""
    text = (
        "🧾 پیش‌نمایش صورتحساب — به همین صورت برای مشتری ارسال می‌شود:\n"
        "━━━━━━━━━━━━━━\n"
        f"{inv['text']}\n"
        "━━━━━━━━━━━━━━\n"
        f"مبلغ قابل پرداخت (ثبت در سیستم): {money(inv['amount'])}"
    )
    if inv["amount"] < inv["default"]:
        text += f"\nتخفیف نسبت به مبلغ پیشنهادی: {money(inv['default'] - inv['amount'])}"
    kb = M(
        [
            [B("✅ تأیید و ارسال برای مشتری", callback_data="a:invy")],
            [B("✏️ ویرایش متن", callback_data="a:invt"), B("💲 تغییر مبلغ / تخفیف", callback_data="a:inva")],
            [B("❌ انصراف", callback_data="a:invc")],
        ]
    )
    return text, kb


async def start_invoice(msg, ctx, user_id):
    d = db.invoice_draft(user_id)
    if not d["count"] or d["amount"] <= 0:
        return await msg.reply_text(
            "چیزی برای صدور صورتحساب نیست: این مشتری سفارش انجام‌شده‌ی بدون صورتحساب و با بدهی باز ندارد."
        )
    name = db.account_last_data(user_id).get("full_name")
    inv = {
        "user_id": user_id, "count": d["count"], "default": d["amount"], "amount": d["amount"],
        "name": name, "edited": False,
        "text": invoice_default_text(name, d["count"], d["amount"]),
    }
    ctx.user_data.pop("state", None)
    ctx.user_data["inv"] = inv
    text, kb = invoice_preview(inv)
    await msg.reply_text(text, reply_markup=kb)


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
        f"🧾 <b>رسید پرداخت #{rid}</b>",
        f"مشتری: {esc(d.get('full_name') or '—')} | {esc(d.get('phone') or '—')}",
        f"جمع بدهی: {summary['debt']:,} | پرداختی قبلی: {summary['paid']:,} | مانده کل: {summary['balance']:,} تومان",
    ]
    inv = db.get_invoice(r["invoice_id"]) if r["invoice_id"] else None
    if inv:
        lines.append(
            f"🧾 صورتحساب #{inv['id']} | مبلغ: {money(inv['amount'])} | "
            f"مانده‌ی صورتحساب: {money(inv['amount'] - inv['paid_amount'])}"
        )
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
    if not st or st[0] != "receipt" or st[1] != uid or len(st) < 3:
        ctx.user_data.clear()
        return await update.message.reply_text(
            "درخواست ارسال رسید معتبر نیست؛ از «💰 وضعیت مالی من» اقدام کنید.", reply_markup=main_kb(uid)
        )
    inv = db.get_invoice(st[2])
    ctx.user_data.clear()
    if not inv or inv["status"] != "unpaid" or uid not in db.account_user_ids(inv["user_id"]):
        return await update.message.reply_text("این صورتحساب دیگر پرداخت‌نشده نیست.", reply_markup=main_kb(uid))
    try:
        rid = db.add_receipt(db.invoice_anchor_order(inv), uid, kind, file_id, file_uid, text, invoice_id=inv["id"])
    except ValueError:
        return await update.message.reply_text(
            "یک رسید برای این صورتحساب در انتظار بررسی است؛ لطفاً تا اعلام نتیجه صبر کنید.",
            reply_markup=main_kb(uid),
        )
    await update.message.reply_text(
        f"✅ رسید پرداخت صورتحساب #{inv['id']} دریافت شد و برای بررسی ارسال شد. "
        "پس از تأیید مدیر، به شما اطلاع داده می‌شود.",
        reply_markup=main_kb(uid),
    )
    await notify_receipt(ctx, rid)


async def finish_approval(msg, ctx, rid, amount, uid):
    res = db.approve_receipt(rid, amount, uid)
    if not res:
        return await msg.reply_text("این رسید قبلاً بررسی شده است.")
    label = "بدهکار" if res["balance"] > 0 else "بستانکار" if res["balance"] < 0 else "تسویه"
    extra = []
    for inv_id, _take, full in res["allocations"]:
        inv = db.get_invoice(inv_id)
        if full:
            extra.append(f"🧾 صورتحساب #{inv_id} تسویه شد ✅")
        else:
            extra.append(f"🧾 صورتحساب #{inv_id}: مانده {money(inv['amount'] - inv['paid_amount'])}")
    await msg.reply_text(
        "\n".join(
            [f"✅ رسید #{rid} تأیید شد.", f"مبلغ ثبت‌شده: {money(amount)}"]
            + extra
            + [f"مانده کل حساب: {abs(res['balance']):,} تومان ({label})"]
        )
    )
    try:
        await ctx.bot.send_message(
            res["user_id"],
            "\n".join(["✅ پرداخت شما تأیید شد.", f"مبلغ ثبت‌شده: {money(amount)}"] + extra),
        )
    except TelegramError:
        pass


ACT_PERM = {
    "home": "view", "l": "view", "o": "view", "search": "view", "hist": "view",
    "users": "view", "usr": "view", "uhist": "view", "upay": "pay",
    "s": "status", "e": "edit", "ef": "edit", "n": "note",
    "pay": "pay", "pa": "pay", "inv": "pay",
    "rca": "pay", "rcf": "pay", "rcm": "pay", "rcr": "pay", "rcl": "pay",
    "x": "export", "xp": "export", "xe": "export",
    "bk": "backup", "bkn": "backup", "bkr": "backup", "bkc": "backup", "bky": "backup",
    "cap": "cap", "captog": "cap", "capset": "cap", "rate": "cap",
    "faq": "faq", "fq": "faq", "fqadd": "faq", "fqq": "faq", "fqe": "faq", "fqd": "faq", "fqdy": "faq",
    "bc": "broadcast", "bcy": "broadcast", "bcc": "broadcast",
    "rt": "rates", "rtn": "rates",
    "invy": "pay", "invt": "pay", "inva": "pay", "invc": "pay",
}  # fmt: skip
STATE_PERM = {
    "a_search": "view", "a_edit": "edit", "a_note": "note", "a_paid": "pay", "a_user_paid": "pay", "a_rc_amount": "pay",
    "a_symrate": "rates", "a_inv_text": "pay", "a_inv_amt": "pay",
    "a_cap_set": "cap", "a_rate": "cap", "a_faq_add_q": "faq", "a_faq_add_a": "faq",
    "a_faq_edit": "faq", "a_restore": "backup", "a_restore_confirm": "backup",
    "a_bc": "broadcast", "a_bc_confirm": "broadcast",
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
    header += ["کارمزد / بدهی (تومان)", "مبلغ دریافتی (تومان)", "مانده (تومان)"]
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

    for row in ws.iter_rows(min_row=2):  # جداکننده‌ی سه‌رقمی برای مبلغ‌ها (ستون آیدی تلگرام مستثناست)
        for c in row:
            if isinstance(c.value, int) and c.column != 6:
                c.number_format = "#,##0"
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
        ("جمع بدهی / کارمزد (تومان)", live["debt"]),
        ("جمع مبلغ دریافتی (تومان)", live["paid"]),
        ("مانده (تومان)", live["debt"] - live["paid"]),
    ]:
        _put(rp, [k, v])
    rp.append([])
    _put(rp, ["بر اساس وضعیت", "تعداد", "کارمزد (تومان)"], True)
    for key, (n, c) in by_status.items():
        _put(rp, [config.STATUSES.get(key, key), n, c])
    rp.append([])
    _put(rp, ["بر اساس نماد", "تعداد", "تعداد خرید", "تعداد فروش"], True)
    for sym, (n, b_, s2) in sorted(by_symbol.items(), key=lambda kv: -kv[1][0]):
        _put(rp, [sym, n, b_, s2])
    rp.append([])
    _put(rp, ["بر اساس روز", "تعداد", "کارمزد (تومان)"], True)
    for d, (n, c) in sorted(by_day.items()):
        _put(rp, [d, n, c])
    for row in rp.iter_rows():
        for c in row:
            if isinstance(c.value, int):
                c.number_format = "#,##0"
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
        filter_key = p[1] if len(p) > 1 and p[1] in ("all", "debtors", "pending") else "all"
        page = max(0, int(p[2]) if len(p) > 2 else (int(p[1]) if len(p) > 1 and p[1].isdigit() else 0))
        all_rows, _ = db.list_users(100000, 0, exclude=all_staff())
        eligible = []
        for user in all_rows:
            finance = db.user_finance(user["user_id"])
            pending = db.count_pending_receipts(user_id=user["user_id"])
            if filter_key == "debtors" and finance["balance"] <= 0:
                continue
            if filter_key == "pending" and not pending:
                continue
            eligible.append((user, finance, pending))
        total = len(eligible)
        rows = eligible[page * PAGE:(page + 1) * PAGE]
        kb = [
            [B("همه کاربران", callback_data="a:users:all:0"),
             B("فقط بدهکارها", callback_data="a:users:debtors:0")],
            [B("رسیدهای در انتظار", callback_data="a:users:pending:0")],
        ]
        for user, finance, pending in rows:
            bal = finance["balance"]
            label = "بدهکار" if bal > 0 else "بستانکار" if bal < 0 else "تسویه"
            username = f"@{user['username']}" if user["username"] else f"ID {user['user_id']}"
            receipt_mark = " 🧾" if pending else ""
            kb.append([B(
                f"{username} | {abs(bal):,} تومان ({label}){receipt_mark}",
                callback_data=f"a:usr:{user['user_id']}",
            )])
        nav = []
        if page > 0:
            nav.append(B("◀️ قبلی", callback_data=f"a:users:{filter_key}:{page - 1}"))
        if (page + 1) * PAGE < total:
            nav.append(B("بعدی ▶️", callback_data=f"a:users:{filter_key}:{page + 1}"))
        if nav:
            kb.append(nav)
        kb.append([B("🏠 پنل مدیریت", callback_data="a:home")])
        title = {"all": "همه کاربران", "debtors": "فقط بدهکارها", "pending": "رسیدهای در انتظار"}[filter_key]
        await show(q, f"👥 <b>{title}</b>\nتعداد: {total}\nصفحه {page + 1}", M(kb))

    elif act == "usr":
        user_id = int(p[1])
        user = db.get_user(user_id)
        if not user:
            return await q.message.reply_text("کاربر پیدا نشد.")
        finance = db.user_finance(user_id)
        rows = [r for r, _ in finance["items"]]
        account_ids = db.account_user_ids(user_id)
        orders = []
        for account_id in account_ids:
            found, _ = db.list_user_orders(account_id, 500, 0)
            orders.extend(found)
        orders.sort(key=lambda r: r["id"], reverse=True)
        profile = json.loads(orders[0]["data"]) if orders else {}
        name = profile.get("full_name") or "—"
        phone = profile.get("phone") or "—"
        balance = finance["balance"]
        status = "بدهکار" if balance > 0 else "بستانکار" if balance < 0 else "تسویه"
        pending = db.count_pending_receipts(user_id=user_id)
        receipts = db.list_user_receipts(user_id, 5)
        account_payments = db.list_account_payments(user_id, 5)
        invs = db.list_unpaid_invoices(user_id)
        lines = [
            "👤 <b>پرونده مالی مشتری</b>",
            f"نام: {esc(name)}",
            f"موبایل: {esc(phone)}",
            f"شناسه تلگرام: <code>{user_id}</code>",
            f"نام کاربری: {('@' + esc(user['username'])) if user['username'] else '—'}",
            f"وضعیت دسترسی: {'مسدود' if user['blocked'] else 'فعال'}",
            f"عضویت: {fmt_time(user['first_seen'])}",
            "",
            f"تعداد سفارش‌ها: {len(orders)}",
            f"جمع بدهی: <b>{finance['debt']:,} تومان</b>",
            f"جمع پرداختی: <b>{finance['paid']:,} تومان</b>",
            f"مانده: <b>{abs(balance):,} تومان ({status})</b>",
            f"رسید در انتظار بررسی: {pending}",
            f"صورتحساب‌های پرداخت‌نشده: {len(invs)}"
            + (f" | جمع مانده: {money(sum(i['amount'] - i['paid_amount'] for i in invs))}" if invs else ""),
        ]
        if receipts or account_payments:
            lines += ["", "<b>آخرین پرداخت‌ها و رسیدها</b>"]
            for receipt in receipts:
                if receipt["status"] != "approved":
                    lines.append(
                        f"رسید #{receipt['id']} | {fmt_time(receipt['created_at'])} | "
                        f"{RECEIPT_LABELS.get(receipt['status'], receipt['status'])}"
                    )
            for payment in account_payments:
                if payment["source"] == "manual" and str(payment.get("note") or "").startswith("تخفیف"):
                    source = "تخفیف صورتحساب"
                else:
                    source = "پرداخت دستی مدیر" if payment["source"] == "manual" else "پرداخت تأییدشده"
                lines.append(
                    f"{source} #{payment['id']} | {fmt_time(payment['created_at'])} | "
                    f"{payment['amount']:,} تومان"
                )
        kb = []
        if pending:
            kb.append([B("🧾 بررسی رسیدهای پرداخت", callback_data=f"a:rcl:{user_id}")])
        if orders:
            kb.append([B("📋 مشاهده سفارش‌های مشتری", callback_data=f"a:uhist:{user_id}")])
            done = next((r for r in orders if r["status"] == "done"), None)
            if done and can(uid, "pay"):
                kb.append([B("🧾 صدور صورتحساب", callback_data=f"a:inv:{done['id']}")])
        if can(uid, "pay"):
            kb.append([B("➕ ثبت پرداخت دستی برای کل حساب", callback_data=f"a:upay:{user_id}")])
        kb.append([B("⬅️ فهرست کاربران", callback_data="a:users:all:0")])
        await show(q, "\n".join(lines), M(kb))

    elif act == "uhist":
        user_id = int(p[1])
        user = db.get_user(user_id)
        if not user:
            return await q.message.reply_text("کاربر پیدا نشد.")
        account_ids = db.account_user_ids(user_id)
        orders = []
        for account_id in account_ids:
            found, _ = db.list_user_orders(account_id, 500, 0)
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
            f"جمع بدهی (کارمزدها): {s['debt']:,} تومان\n"
            f"جمع پرداختی‌ها: {s['paid']:,} تومان\n"
            f"مانده حساب: {abs(bal):,} تومان ({bal_lbl})\n\n"
            f"اولین سفارش: {fmt_time(rows[-1]['created_at'])}"
        )
        kb = [order_button(r) for r in rows[:15]]
        kb.append([B("⬅️ بازگشت به سفارش", callback_data=f"a:o:{o['id']}")])
        await show(q, text, M(kb))

    # ---- payment & commission ----
    elif act == "upay":
        user_id = int(p[1])
        finance = db.user_finance(user_id)
        ctx.user_data["state"] = ("a_user_paid", user_id)
        await q.message.reply_text(
            f"ثبت پرداخت دستی برای کل حساب مشتری\n"
            f"مانده فعلی: {abs(finance['balance']):,} تومان\n"
            "مبلغ پرداخت‌شده را به تومان و فقط با عدد بفرستید:",
            reply_markup=ReplyKeyboardMarkup([[BTN_CANCEL]], resize_keyboard=True),
        )

    elif act == "pay":
        o = db.get_order(int(p[1]))
        if not o:
            return await show(q, "سفارش پیدا نشد.", admin_kb(uid))
        await show(q, render_order(o, admin=True, reveal=can(uid, "secret")), pay_kb(o))

    elif act == "inv":  # a:inv:<order_id> → پیش‌نویس صورتحساب
        o = db.get_order(int(p[1]))
        if not o:
            return await q.message.reply_text("سفارش پیدا نشد.")
        await start_invoice(q.message, ctx, o["user_id"])

    elif act in ("invy", "invt", "inva", "invc"):
        inv = ctx.user_data.get("inv")
        if not inv:
            return await q.message.reply_text("درخواست منقضی شده؛ دوباره «🧾 صدور صورتحساب» را بزنید.")
        if act == "invc":
            ctx.user_data.pop("inv", None)
            ctx.user_data.pop("state", None)
            try:
                await q.edit_message_reply_markup(None)
            except TelegramError:
                pass
            return await q.message.reply_text("صدور صورتحساب لغو شد.", reply_markup=main_kb(uid))
        if act == "invt":
            ctx.user_data["state"] = ("a_inv_text",)
            await q.message.reply_text(
                "متن جدید صورتحساب را بفرستید. متن فعلی (برای کپی و ویرایش):\n\n" + inv["text"],
                reply_markup=ReplyKeyboardMarkup([[BTN_CANCEL]], resize_keyboard=True),
            )
        elif act == "inva":
            ctx.user_data["state"] = ("a_inv_amt",)
            await q.message.reply_text(
                "مبلغ نهایی صورتحساب را به «تومان» و فقط با عدد بفرستید.\n"
                f"حداکثر {inv['default']:,} تومان (مبلغ پیشنهادی)؛ عدد کمتر یعنی تخفیف.",
                reply_markup=ReplyKeyboardMarkup([[BTN_CANCEL]], resize_keyboard=True),
            )
        else:  # invy
            res = db.create_invoice(inv["user_id"], inv["text"], inv["amount"], inv["default"], uid)
            if not res:
                ctx.user_data.pop("inv", None)
                return await q.message.reply_text(
                    "وضعیت حساب مشتری از زمان تهیه‌ی پیش‌نویس تغییر کرده است؛ لطفاً صدور صورتحساب را دوباره شروع کنید."
                )
            ctx.user_data.pop("inv", None)
            try:
                await q.edit_message_reply_markup(None)
            except TelegramError:
                pass
            try:
                await ctx.bot.send_message(
                    inv["user_id"],
                    f"{inv['text']}\n\n🧾 شماره صورتحساب: {res['id']}",
                    reply_markup=M([[B("🧾 ارسال رسید پرداخت", callback_data=f"rc:{res['id']}")]]),
                )
                note = f"✅ صورتحساب #{res['id']} برای مشتری ارسال شد."
            except TelegramError:
                note = (
                    f"⚠️ صورتحساب #{res['id']} ثبت شد ولی ارسال پیام ناموفق بود (احتمالاً مشتری ربات را مسدود کرده). "
                    "مشتری در صورت بازگشت، آن را در «وضعیت مالی من» می‌بیند."
                )
            if res["discount"]:
                note += f"\nتخفیف ثبت‌شده در حساب مشتری: {money(res['discount'])}"
            await q.message.reply_text(note, reply_markup=main_kb(uid))

    elif act == "pa":
        ctx.user_data["state"] = ("a_paid", int(p[1]))
        await q.message.reply_text(
            "مبلغ دریافتی از مشتری را به تومان و فقط با عدد بفرستید (۰ = هیچ):",
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
        inv = db.get_invoice(r["invoice_id"]) if r["invoice_id"] else None
        default_amt = inv["amount"] - inv["paid_amount"] if inv else summary["balance"]
        if act == "rca":
            rows = []
            if default_amt > 0:
                rows.append([B(f"✅ ثبت مبلغ صورتحساب ({default_amt:,} تومان)", callback_data=f"a:rcf:{rid}")])
            rows.append([B("✏️ مبلغ دیگر", callback_data=f"a:rcm:{rid}")])
            await q.message.reply_text(
                f"تأیید رسید #{rid}" + (f" — صورتحساب #{inv['id']}" if inv else " (کل حساب)") + "\n"
                f"جمع بدهی: {summary['debt']:,} | پرداختی قبلی: {summary['paid']:,} | مانده کل: {summary['balance']:,} تومان\n\n"
                "چه مبلغی از این رسید ثبت شود؟",
                reply_markup=M(rows),
            )
        elif act == "rcf":
            if default_amt <= 0:
                return await q.message.reply_text("مبلغی برای ثبت باقی نمانده؛ «مبلغ دیگر» را بزنید.")
            try:
                await q.edit_message_reply_markup(None)
            except TelegramError:
                pass
            await finish_approval(q.message, ctx, rid, default_amt, uid)
        elif act == "rcm":
            ctx.user_data["state"] = ("a_rc_amount", rid)
            await q.message.reply_text(
                f"مبلغ دریافتی رسید #{rid} را به تومان و فقط با عدد بفرستید:",
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
                    "❌ رسید پرداخت شما پذیرفته نشد. لطفاً رسید صحیح را از بخش «💰 وضعیت مالی من» (زیر همان صورتحساب) دوباره ارسال کنید.",
                )
            except TelegramError:
                pass

    elif act == "rcl":
        pend = db.list_pending_receipts(user_id=int(p[1]))
        if not pend:
            return await q.message.reply_text("رسید پرداخت در انتظاری برای این مشتری نیست.")
        for r in pend:
            await notify_receipt(ctx, r["id"], only=[uid])

    # ---- commission rates (همان فهرستی که مشتری می‌بیند) ----
    elif act == "rt":
        lines = rates_lines()
        kb = M([[B("➕ ثبت / ویرایش / حذف نرخ نماد", callback_data="a:rtn")], [B("🏠 پنل", callback_data="a:home")]])
        await show(
            q,
            "💹 <b>نرخ کارمزد سهام‌ها</b>\n(دقیقاً همین فهرست به مشتری نمایش داده می‌شود و کارمزد سفارش‌ها هم از آن محاسبه می‌شود)\n\n"
            + (esc("\n".join(lines)) if lines else "هنوز نرخی ثبت نشده است."),
            kb,
        )

    elif act == "rtn":
        ctx.user_data["state"] = ("a_symrate",)
        await q.message.reply_text(
            "نام نماد و درصد کارمزد را بفرستید. مثال: فولاد 2\n"
            "برای حذف نرخ یک نماد: فولاد حذف",
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

    elif kind == "a_user_paid":
        v = fa2en(raw).replace(",", "").replace(" ", "")
        if not v.isdigit() or not 0 < int(v) <= 10**13:
            return await msg.reply_text("لطفاً مبلغ مثبت و معتبر را به تومان وارد کنید.")
        user_id = st[1]
        db.add_manual_customer_payment(user_id, int(v), uid, "ثبت دستی توسط مدیر")
        ctx.user_data.clear()
        finance = db.user_finance(user_id)
        bal = finance["balance"]
        label = "بدهکار" if bal > 0 else "بستانکار" if bal < 0 else "تسویه"
        await msg.reply_text(
            f"پرداخت دستی {int(v):,} تومان ثبت شد.\n"
            f"جمع پرداختی‌ها: {finance['paid']:,} تومان\n"
            f"مانده کل: {abs(bal):,} تومان ({label})",
            reply_markup=main_kb(uid),
        )
        await msg.reply_text(
            "پرونده مالی مشتری:",
            reply_markup=M([[B("مشاهده پرونده", callback_data=f"a:usr:{user_id}")]]),
        )

    elif kind == "a_paid":
        v = fa2en(raw).replace(",", "").replace(" ", "")
        if not v.isdigit() or int(v) > 10**13:
            return await msg.reply_text("لطفاً فقط عدد (تومان، بدون حروف) بفرستید.")
        oid = st[1]
        db.set_paid_amount(oid, int(v), uid)
        ctx.user_data.clear()
        await msg.reply_text("مبلغ دریافتی ثبت شد ✅", reply_markup=main_kb(uid))
        await msg.reply_text(
            "بازگشت:", reply_markup=M([[B("💳 پرداخت و فاکتور", callback_data=f"a:pay:{oid}")]])
        )

    elif kind == "a_symrate":
        parts = fa2en(raw).replace("٪", "").replace("%", "").replace("٫", ".").split()
        if len(parts) >= 2 and parts[-1] in ("حذف", "-"):
            ok = db.delete_symbol_rate(" ".join(parts[:-1]))
            ctx.user_data.clear()
            return await msg.reply_text(
                "نرخ این نماد حذف شد ✅" if ok else "برای این نماد نرخی ثبت نشده بود.", reply_markup=main_kb(uid)
            )
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
            return await msg.reply_text("لطفاً فقط عدد (تومان، بزرگ‌تر از صفر) بفرستید.")
        ctx.user_data.clear()
        await msg.reply_text("در حال ثبت…", reply_markup=main_kb(uid))
        await finish_approval(msg, ctx, st[1], int(v), uid)

    elif kind == "a_inv_text":
        inv = ctx.user_data.get("inv")
        if not inv:
            ctx.user_data.clear()
            return await msg.reply_text("درخواست منقضی شده؛ دوباره صدور صورتحساب را شروع کنید.", reply_markup=main_kb(uid))
        if not raw or len(raw) > 3000:
            return await msg.reply_text("متن خالی یا بیش از ۳۰۰۰ کاراکتر است.")
        inv["text"], inv["edited"] = raw, True
        ctx.user_data.pop("state", None)
        await msg.reply_text("متن ثبت شد ✅", reply_markup=main_kb(uid))
        text, kb = invoice_preview(inv)
        await msg.reply_text(text, reply_markup=kb)

    elif kind == "a_inv_amt":
        inv = ctx.user_data.get("inv")
        if not inv:
            ctx.user_data.clear()
            return await msg.reply_text("درخواست منقضی شده؛ دوباره صدور صورتحساب را شروع کنید.", reply_markup=main_kb(uid))
        v = fa2en(raw).replace(",", "").replace(" ", "")
        if not v.isdigit() or not 0 < int(v) <= inv["default"]:
            return await msg.reply_text(f"عدد معتبر بین ۱ و {inv['default']:,} (تومان) بفرستید.")
        inv["amount"] = int(v)
        note = "مبلغ ثبت شد ✅"
        if inv["edited"]:
            note += "\nمتن را قبلاً دستی ویرایش کرده‌اید و تغییر نکرد؛ در صورت نیاز «ویرایش متن» را بزنید."
        else:
            inv["text"] = invoice_default_text(inv["name"], inv["count"], inv["amount"])
        ctx.user_data.pop("state", None)
        await msg.reply_text(note, reply_markup=main_kb(uid))
        text, kb = invoice_preview(inv)
        await msg.reply_text(text, reply_markup=kb)

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
            await q.message.reply_text(
                render_order(o),
                parse_mode="HTML",
                reply_markup=M([[B("🔁 ثبت دوباره‌ی این سفارش", callback_data=f"pf:re:{o['id']}")]]),
            )
        return

    if d.startswith("pf:"):
        if is_admin(uid):
            return
        parts = d.split(":")
        act = parts[1]
        try:
            await q.edit_message_reply_markup(None)
        except TelegramError:
            pass
        if act == "del":
            db.delete_profile(uid)
            return await q.message.reply_text("اطلاعات ذخیره‌شده حذف شد.", reply_markup=main_kb(uid))
        blocked = order_block_msg(uid)
        if blocked:
            return await q.message.reply_text(blocked, reply_markup=main_kb(uid))
        profile = db.get_profile(uid)
        if act == "use":
            return await begin_form(q.message, ctx, form=profile, skip=profile.keys(), defaults=profile)
        if act == "edit":
            return await begin_form(q.message, ctx, defaults=profile)
        if act == "new":
            return await begin_form(q.message, ctx)
        if act == "re" and len(parts) == 3 and parts[2].isdigit():
            o = db.get_order(int(parts[2]))
            if not o or o["user_id"] != uid:
                return
            old = json.loads(o["data"])
            data = {f["key"]: old[f["key"]] for f in config.FORM_FIELDS if f["kind"] != "secret" and old.get(f["key"])}
            return await begin_form(q.message, ctx, form=data, skip=data.keys(), defaults=data)
        return

    if d == "fin":
        if is_admin(uid):
            return
        return await send_fin(q.message, uid)

    if d.startswith("rc:"):
        if is_admin(uid):
            return
        key = d[3:]
        inv = db.get_invoice(int(key)) if key.isdigit() else None
        if not inv:
            return await q.message.reply_text("برای ارسال رسید، از «💰 وضعیت مالی من» و زیر صورتحساب اقدام کنید.")
        if inv["status"] != "unpaid" or uid not in db.account_user_ids(inv["user_id"]):
            return await q.message.reply_text("این صورتحساب دیگر پرداخت‌نشده نیست.")
        if db.count_pending_receipts(invoice_id=inv["id"]):
            return await q.message.reply_text("یک رسید برای این صورتحساب در انتظار بررسی است؛ لطفاً تا اعلام نتیجه صبر کنید.")
        ctx.user_data.clear()
        ctx.user_data["state"] = ("receipt", uid, inv["id"])
        return await q.message.reply_text(
            f"🧾 ارسال رسید پرداخت — صورتحساب #{inv['id']}\n"
            f"مبلغ قابل پرداخت: {money(inv['amount'] - inv['paid_amount'])}\n\n"
            "عکس یا فایل (PDF) رسید پرداخت را بفرستید؛ می‌توانید شماره پیگیری را در توضیح (کپشن) بنویسید.\n"
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
        prof = {f["key"]: data[f["key"]] for f in config.FORM_FIELDS if f.get("remember") and data.get(f["key"])}
        if prof:
            db.save_profile(uid, prof)
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
