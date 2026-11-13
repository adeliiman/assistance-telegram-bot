import glob
import io
import json
import os
import secrets
import sqlite3
import tempfile
from pathlib import Path
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone

import config

ALPHABET = "23456789ABCDEFGHJKLMNPQRSTUVWXYZ"  # بدون حروف/اعداد مشابه (O/0, I/1)


@contextmanager
def conn():
    c = sqlite3.connect(config.DB_PATH)
    c.row_factory = sqlite3.Row
    try:
        yield c
        c.commit()
    finally:
        c.close()


def now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def init_db():
    with conn() as c:
        c.executescript(
            """
            CREATE TABLE IF NOT EXISTS orders(
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              code TEXT UNIQUE NOT NULL,
              user_id INTEGER NOT NULL,
              username TEXT,
              status TEXT NOT NULL DEFAULT 'new',
              data TEXT NOT NULL,
              created_at TEXT NOT NULL,
              updated_at TEXT NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_orders_user ON orders(user_id);
            CREATE INDEX IF NOT EXISTS idx_orders_status ON orders(status);
            CREATE TABLE IF NOT EXISTS order_log(
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              order_id INTEGER NOT NULL,
              actor_id INTEGER,
              action TEXT NOT NULL,
              detail TEXT,
              at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS order_notes(
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              order_id INTEGER NOT NULL,
              admin_id INTEGER NOT NULL,
              text TEXT NOT NULL,
              at TEXT NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_notes_order ON order_notes(order_id);
            CREATE TABLE IF NOT EXISTS settings(
              key TEXT PRIMARY KEY,
              value TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS faq(
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              question TEXT NOT NULL,
              answer TEXT NOT NULL,
              position INTEGER NOT NULL DEFAULT 0
            );
            CREATE TABLE IF NOT EXISTS users(
              user_id INTEGER PRIMARY KEY,
              username TEXT,
              first_seen TEXT NOT NULL,
              blocked INTEGER NOT NULL DEFAULT 0
            );
            CREATE TABLE IF NOT EXISTS receipts(
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              order_id INTEGER NOT NULL,
              user_id INTEGER NOT NULL,
              kind TEXT NOT NULL,            -- photo | document | text
              file_id TEXT,
              file_uid TEXT,
              text TEXT,
              status TEXT NOT NULL DEFAULT 'pending',  -- pending | approved | rejected
              amount INTEGER,
              created_at TEXT NOT NULL,
              decided_at TEXT,
              decided_by INTEGER
            );
            CREATE INDEX IF NOT EXISTS idx_receipts_order ON receipts(order_id);
            CREATE TABLE IF NOT EXISTS customer_payments(
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              receipt_id INTEGER NOT NULL UNIQUE,
              user_id INTEGER NOT NULL,
              amount INTEGER NOT NULL,
              created_at TEXT NOT NULL,
              actor_id INTEGER NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_customer_payments_user ON customer_payments(user_id);
            """
        )
        # مهاجرت ساختار: ستون‌های پرداخت و کارمزد (روی دیتابیس‌ها و بکاپ‌های قدیمی هم کار می‌کند)
        cols = {r["name"] for r in c.execute("PRAGMA table_info(orders)")}
        if "paid_amount" not in cols:
            c.execute("ALTER TABLE orders ADD COLUMN paid_amount INTEGER NOT NULL DEFAULT 0")
        if "commission" not in cols:  # NULL = محاسبه‌ی خودکار از نرخ کارمزد
            c.execute("ALTER TABLE orders ADD COLUMN commission INTEGER")
        if "trade_amount" not in cols:
            c.execute("ALTER TABLE orders ADD COLUMN trade_amount INTEGER")
        if "fee_rate" not in cols:
            c.execute("ALTER TABLE orders ADD COLUMN fee_rate REAL")
        # هماهنگ‌سازی با وضعیت‌های جدید (روی دیتابیس‌ها و بکاپ‌های قدیمی هم کار می‌کند)
        c.execute("UPDATE orders SET status='new' WHERE status IN ('in_progress')")
        c.execute("UPDATE orders SET status='error' WHERE status='canceled'")
        # مشتری‌های قبلی (که سفارش داده‌اند) برای اطلاعیه‌ی گروهی وارد جدول کاربران می‌شوند
        c.execute(
            "INSERT OR IGNORE INTO users(user_id, username, first_seen) "
            "SELECT user_id, MAX(username), MIN(created_at) FROM orders GROUP BY user_id"
        )
        if c.execute("SELECT COUNT(*) FROM faq").fetchone()[0] == 0:
            for i, (q, a) in enumerate(config.DEFAULT_FAQ):
                c.execute("INSERT INTO faq(question, answer, position) VALUES(?,?,?)", (q, a, i))


# ---------------- orders ----------------
def _log(c, order_id, actor_id, action, detail=""):
    c.execute(
        "INSERT INTO order_log(order_id, actor_id, action, detail, at) VALUES(?,?,?,?,?)",
        (order_id, actor_id, action, detail, now()),
    )

def symbol_rate(symbol):
    with conn() as c:
        v = _get(c, "symrate:" + symbol.strip(), None)
    return float(v) if v is not None else None


def set_symbol_rate(symbol, percent: float):
    with conn() as c:
        _set(c, "symrate:" + symbol.strip(), percent)


def list_symbol_rates():
    with conn() as c:
        return [(r["key"][8:], float(r["value"])) for r in
                c.execute("SELECT key, value FROM settings WHERE key LIKE 'symrate:%' ORDER BY key")]


def complete_order(order_id, amount: int, actor_id):
    """تکمیل سفارش و محاسبه خودکار کارمزد از نرخ همان نماد یا نرخ پیش‌فرض."""
    if amount <= 0:
        return None
    with conn() as c:
        c.execute("BEGIN IMMEDIATE")
        o = c.execute("SELECT * FROM orders WHERE id=?", (order_id,)).fetchone()
        if not o:
            return None
        sym = json.loads(o["data"]).get("symbol", "").strip()
        v = _get(c, "symrate:" + sym, None)
        rate = float(v) if v is not None else float(
            _get(c, "commission_percent", str(config.DEFAULT_COMMISSION_PERCENT))
        )
        if rate <= 0:
            return None
        fee = round(amount * rate / 100)
        c.execute(
            "UPDATE orders SET status='done', trade_amount=?, fee_rate=?, commission=?, updated_at=? WHERE id=?",
            (amount, rate, fee, now(), order_id),
        )
        _log(c, order_id, actor_id, "done", f"amount={amount} rate={rate} fee={fee}")
        return {
            "code": o["code"], "user_id": o["user_id"], "symbol": sym,
            "amount": amount, "rate": rate, "fee": fee,
        }

# ---------------- settings / capacity ----------------
def _get(c, key, default):
    r = c.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return r["value"] if r else default


def _set(c, key, value):
    c.execute(
        "INSERT INTO settings(key, value) VALUES(?,?)"
        " ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (key, str(value)),
    )


def _intake(c) -> dict:
    """وضعیت پذیرش سفارش. capacity=0 یعنی نامحدود.
    سفارش‌های خطادار در شمارش ظرفیت حساب نمی‌شوند."""
    accepting = _get(c, "accepting", "1") == "1"
    capacity = int(_get(c, "capacity", "0"))
    used = c.execute(
        "SELECT COUNT(*) FROM orders WHERE status!='error'"
    ).fetchone()[0]
    remaining = None if capacity == 0 else max(capacity - used, 0)
    if not accepting:
        reason = "closed"
    elif remaining == 0:
        reason = "full"
    else:
        reason = None
    return {
        "accepting": accepting,
        "capacity": capacity,
        "used": used,
        "remaining": remaining,
        "reason": reason,  # None | "closed" | "full"
    }


def intake_state() -> dict:
    with conn() as c:
        return _intake(c)


def set_accepting(flag: bool):
    with conn() as c:
        _set(c, "accepting", "1" if flag else "0")


def set_capacity(n: int):
    """سقف کل سفارش‌ها (به‌جز لغوشده‌ها). شمارش ریست نمی‌شود و همه‌ی سفارش‌های
    قبلی هم حساب می‌شوند. n=0 یعنی بدون محدودیت."""
    with conn() as c:
        _set(c, "capacity", n)


# ---------------- users (برای اطلاعیه‌ی گروهی) ----------------
def _upsert_user(c, user_id, username):
    c.execute(
        "INSERT INTO users(user_id, username, first_seen, blocked) VALUES(?,?,?,0)"
        " ON CONFLICT(user_id) DO UPDATE SET"
        " username=COALESCE(excluded.username, users.username), blocked=0",
        (user_id, username, now()),
    )


def upsert_user(user_id, username=None):
    with conn() as c:
        _upsert_user(c, user_id, username)


def broadcast_targets(exclude=()) -> list:
    ex = set(exclude)
    with conn() as c:
        ids = [r[0] for r in c.execute("SELECT user_id FROM users WHERE blocked=0 ORDER BY user_id")]
    return [i for i in ids if i not in ex]


def mark_blocked(user_id):
    with conn() as c:
        c.execute("UPDATE users SET blocked=1 WHERE user_id=?", (user_id,))


def user_counts(exclude=()):
    """(کل کاربران, کاربران فعال)؛ شناسه‌های exclude (مدیرها) حساب نمی‌شوند."""
    ex = set(exclude)
    with conn() as c:
        rows = [r for r in c.execute("SELECT user_id, blocked FROM users") if r[0] not in ex]
    return len(rows), sum(1 for r in rows if not r[1])


# ---------------- orders ----------------
def create_order(user_id, username, data: dict):
    """برمی‌گرداند (order_id, code) یا (None, reason) اگر پذیرش بسته/پر باشد.
    بررسی ظرفیت و درج سفارش در یک تراکنش قفل‌شده انجام می‌شود تا
    ثبت هم‌زمان چند کاربر از ظرفیت عبور نکند."""
    with conn() as c:
        c.execute("BEGIN IMMEDIATE")
        reason = _intake(c)["reason"]
        if reason:
            return None, reason
        while True:
            code = "SR-" + "".join(secrets.choice(ALPHABET) for _ in range(6))
            if not c.execute("SELECT 1 FROM orders WHERE code=?", (code,)).fetchone():
                break
        t = now()
        cur = c.execute(
            "INSERT INTO orders(code, user_id, username, status, data, created_at, updated_at)"
            " VALUES(?,?,?,?,?,?,?)",
            (code, user_id, username, "new", json.dumps(data, ensure_ascii=False), t, t),
        )
        _log(c, cur.lastrowid, user_id, "created")
        _upsert_user(c, user_id, username)
        return cur.lastrowid, code


def count_orders_today(user_id) -> int:
    since = (datetime.now(timezone.utc) - timedelta(days=1)).isoformat(timespec="seconds")
    with conn() as c:
        return c.execute(
            "SELECT COUNT(*) FROM orders WHERE user_id=? AND created_at>=?", (user_id, since)
        ).fetchone()[0]


def get_order(order_id):
    with conn() as c:
        return c.execute("SELECT * FROM orders WHERE id=?", (order_id,)).fetchone()


def get_order_by_code(code):
    with conn() as c:
        return c.execute("SELECT * FROM orders WHERE code=?", (code,)).fetchone()


def list_orders(status=None, limit=8, offset=0):
    where, args = ("WHERE status=?", [status]) if status else ("", [])
    with conn() as c:
        total = c.execute(f"SELECT COUNT(*) FROM orders {where}", args).fetchone()[0]
        rows = c.execute(
            f"SELECT * FROM orders {where} ORDER BY id DESC LIMIT ? OFFSET ?", args + [limit, offset]
        ).fetchall()
    return rows, total


def list_user_orders(user_id, limit=8, offset=0):
    """سفارش‌های یک مشتری (جدیدترین اول) و تعداد کل."""
    with conn() as c:
        total = c.execute("SELECT COUNT(*) FROM orders WHERE user_id=?", (user_id,)).fetchone()[0]
        rows = c.execute(
            "SELECT * FROM orders WHERE user_id=? ORDER BY id DESC LIMIT ? OFFSET ?",
            (user_id, limit, offset),
        ).fetchall()
    return rows, total


def list_orders_filtered(status=None, since=None):
    """برای خروجی اکسل. since = رشته‌ی ISO به وقت UTC (یا None)."""
    q, args = "SELECT * FROM orders WHERE 1=1", []
    if status:
        q += " AND status=?"
        args.append(status)
    if since:
        q += " AND created_at>=?"
        args.append(since)
    with conn() as c:
        return c.execute(q + " ORDER BY id", args).fetchall()


def search_orders(q, limit=20):
    like = f"%{q}%"
    with conn() as c:
        return c.execute(
            "SELECT * FROM orders WHERE code LIKE ? OR data LIKE ? OR username LIKE ?"
            " ORDER BY id DESC LIMIT ?",
            (like, like, like, limit),
        ).fetchall()


def all_orders():
    with conn() as c:
        return c.execute("SELECT * FROM orders ORDER BY id").fetchall()


def set_status(order_id, status, actor_id):
    with conn() as c:
        c.execute("UPDATE orders SET status=?, updated_at=? WHERE id=?", (status, now(), order_id))
        if status != "done":
            c.execute("UPDATE orders SET commission=NULL, trade_amount=NULL, fee_rate=NULL WHERE id=?", (order_id,))
        _log(c, order_id, actor_id, "status", status)


def set_field(order_id, key, value, actor_id):
    with conn() as c:
        row = c.execute("SELECT data FROM orders WHERE id=?", (order_id,)).fetchone()
        data = json.loads(row["data"])
        old = data.get(key, "")
        data[key] = value
        c.execute(
            "UPDATE orders SET data=?, updated_at=? WHERE id=?",
            (json.dumps(data, ensure_ascii=False), now(), order_id),
        )
        _log(c, order_id, actor_id, "edit", f"{key}: {old!r} -> {value!r}")


# ---------------- payment & commission ----------------
def _to_int(v) -> int:
    try:
        return int(str(v).replace(",", "").strip())
    except (TypeError, ValueError):
        return 0


def order_total(o) -> int:
    """مبلغ سفارش را از اطلاعات ثبت‌شده استخراج می‌کند (همه مبالغ ریال هستند).

    اگر تعداد و قیمت موجود باشند، ارزش معامله از حاصل‌ضرب آن‌ها محاسبه می‌شود؛
    در نسخه‌هایی که فیلد قیمت ندارند، قدرت خرید ثبت‌شده مبنا قرار می‌گیرد.
    """
    d = json.loads(o["data"])
    quantity = _to_int(d.get("quantity"))
    price = _to_int(d.get("price"))
    if quantity > 0 and price > 0:
        return quantity * price
    return _to_int(d.get("buying_power"))


def get_commission_percent() -> float:
    with conn() as c:
        try:
            return float(_get(c, "commission_percent", str(config.DEFAULT_COMMISSION_PERCENT)))
        except ValueError:
            return 0.0


def set_commission_percent(percent: float):
    with conn() as c:
        _set(c, "commission_percent", percent)


def order_fees(o, rate=None) -> dict:
    commission = o["commission"] or 0
    paid = o["paid_amount"] or 0
    return {
        "commission": commission,
        "trade_amount": o["trade_amount"] or 0,
        "rate": o["fee_rate"] or 0,
        "paid": paid,
        "balance": commission - paid,
    }


def set_paid_amount(order_id, amount: int, actor_id):
    with conn() as c:
        c.execute("UPDATE orders SET paid_amount=?, updated_at=? WHERE id=?", (amount, now(), order_id))
        _log(c, order_id, actor_id, "paid_amount", str(amount))




# ---------------- payment receipts ----------------
def _account_user_ids(c, user_id):
    """شناسه‌های تلگرام متعلق به یک مشتری؛ تطبیق بر اساس آیدی یا شماره موبایل."""
    anchor = c.execute(
        "SELECT * FROM orders WHERE user_id=? ORDER BY id DESC LIMIT 1", (user_id,)
    ).fetchone()
    if not anchor:
        return [user_id]
    phone = json.loads(anchor["data"]).get("phone", "")
    rows = c.execute(
        "SELECT DISTINCT user_id FROM orders WHERE user_id=? OR (?!='' AND data LIKE ?)",
        (user_id, phone, f'%"phone": "{phone}"%'),
    ).fetchall()
    return sorted({row["user_id"] for row in rows} | {user_id})


def add_receipt(order_id, user_id, kind, file_id=None, file_uid=None, text=None) -> int:
    """ثبت یک رسید برای کل حساب؛ هم‌زمان بیش از یک رسید معلق برای مشتری مجاز نیست."""
    with conn() as c:
        c.execute("BEGIN IMMEDIATE")
        account_ids = _account_user_ids(c, user_id)
        marks = ",".join("?" for _ in account_ids)
        pending = c.execute(
            f"SELECT 1 FROM receipts WHERE user_id IN ({marks}) AND status='pending' LIMIT 1",
            account_ids,
        ).fetchone()
        if pending:
            raise ValueError("customer already has a pending receipt")
        cur = c.execute(
            "INSERT INTO receipts(order_id, user_id, kind, file_id, file_uid, text, created_at)"
            " VALUES(?,?,?,?,?,?,?)",
            (order_id, user_id, kind, file_id, file_uid, text, now()),
        )
        _log(c, order_id, user_id, "receipt", f"#{cur.lastrowid} account-wide {kind}")
        return cur.lastrowid


def get_receipt(rid):
    with conn() as c:
        return c.execute("SELECT * FROM receipts WHERE id=?", (rid,)).fetchone()


def count_pending_receipts(user_id=None, order_id=None) -> int:
    q, args = "SELECT COUNT(*) FROM receipts WHERE status='pending'", []
    if user_id is not None:
        with conn() as c:
            account_ids = _account_user_ids(c, user_id)
        marks = ",".join("?" for _ in account_ids)
        q += f" AND user_id IN ({marks})"
        args.extend(account_ids)
    if order_id is not None:
        q += " AND order_id=?"
        args.append(order_id)
    with conn() as c:
        return c.execute(q, args).fetchone()[0]


def list_pending_receipts(user_id):
    """رسیدهای در انتظار بررسی برای کل حساب مشتری."""
    with conn() as c:
        account_ids = _account_user_ids(c, user_id)
        marks = ",".join("?" for _ in account_ids)
        return c.execute(
            f"SELECT * FROM receipts WHERE user_id IN ({marks}) AND status='pending' ORDER BY id",
            account_ids,
        ).fetchall()


def find_duplicate_receipts(rec, limit=3):
    """رسیدهای قبلی (غیر ردشده) با همان فایل یا همان شماره پیگیری؛ فقط برای هشدار به مالک."""
    conds, args = [], [rec["id"]]
    if rec["file_uid"]:
        conds.append("r.file_uid=?")
        args.append(rec["file_uid"])
    t = (rec["text"] or "").strip().lower()
    if len(t) >= 6:
        conds.append("LOWER(TRIM(r.text))=?")
        args.append(t)
    if not conds:
        return []
    with conn() as c:
        return c.execute(
            "SELECT r.*, o.code AS code FROM receipts r JOIN orders o ON o.id=r.order_id"
            " WHERE r.id!=? AND r.status!='rejected' AND (" + " OR ".join(conds) + ")"
            " ORDER BY r.id DESC LIMIT ?",
            args + [limit],
        ).fetchall()


def approve_receipt(rid, amount: int, actor_id):
    """ثبت تأیید و پرداخت در دفتر کل مشتری، نه روی یک سفارش خاص."""
    with conn() as c:
        c.execute("BEGIN IMMEDIATE")
        cur = c.execute(
            "UPDATE receipts SET status='approved', amount=?, decided_at=?, decided_by=?"
            " WHERE id=? AND status='pending'",
            (amount, now(), actor_id, rid),
        )
        if cur.rowcount == 0:
            return None
        r = c.execute("SELECT * FROM receipts WHERE id=?", (rid,)).fetchone()
        o = c.execute("SELECT * FROM orders WHERE id=?", (r["order_id"],)).fetchone()
        c.execute(
            "INSERT INTO customer_payments(receipt_id,user_id,amount,created_at,actor_id) VALUES(?,?,?,?,?)",
            (rid, r["user_id"], amount, now(), actor_id),
        )
        _log(c, o["id"], actor_id, "receipt_approved", f"#{rid} customer payment +{amount}")
        account_ids = _account_user_ids(c, r["user_id"])
        marks = ",".join("?" for _ in account_ids)
        sums = c.execute(
            f"SELECT COALESCE(SUM(commission),0) debt, COALESCE(SUM(paid_amount),0) legacy_paid"
            f" FROM orders WHERE user_id IN ({marks}) AND status!='error'", account_ids
        ).fetchone()
        ledger = c.execute(
            f"SELECT COALESCE(SUM(amount),0) FROM customer_payments WHERE user_id IN ({marks})",
            account_ids,
        ).fetchone()[0]
        paid = sums["legacy_paid"] + ledger
        return {"order_id":o["id"],"code":o["code"],"user_id":r["user_id"],
                "paid":paid,"commission":sums["debt"],"balance":sums["debt"]-paid}


def reject_receipt(rid, actor_id):
    """برمی‌گرداند (موفق؟, ردیف رسید)."""
    with conn() as c:
        r = c.execute("SELECT * FROM receipts WHERE id=?", (rid,)).fetchone()
        if not r:
            return False, None
        cur = c.execute(
            "UPDATE receipts SET status='rejected', decided_at=?, decided_by=? WHERE id=? AND status='pending'",
            (now(), actor_id, rid),
        )
        if cur.rowcount:
            _log(c, r["order_id"], actor_id, "receipt_rejected", f"#{rid}")
        return cur.rowcount > 0, r


def list_user_receipts(user_id, limit=10):
    """رسیدهای یک مشتری (جدیدترین اول) همراه کد سفارش."""
    with conn() as c:
        account_ids = _account_user_ids(c, user_id)
        marks = ",".join("?" for _ in account_ids)
        return c.execute(
            "SELECT r.*, o.code AS code FROM receipts r JOIN orders o ON o.id=r.order_id"
            f" WHERE r.user_id IN ({marks}) ORDER BY r.id DESC LIMIT ?",
            account_ids + [limit],
        ).fetchall()


def list_order_receipts(order_id, limit=5):
    with conn() as c:
        return c.execute(
            "SELECT * FROM receipts WHERE order_id=? ORDER BY id DESC LIMIT ?", (order_id, limit)
        ).fetchall()


def user_finance(user_id) -> dict:
    """حساب مالی خودِ مشتری (فقط با همین آیدی تلگرام؛ سفارش‌های خطادار حساب نمی‌شوند)."""
    rate = get_commission_percent()
    with conn() as c:
        account_ids = _account_user_ids(c, user_id)
        marks = ",".join("?" for _ in account_ids)
        rows = c.execute(
            f"SELECT * FROM orders WHERE user_id IN ({marks}) AND status!='error' ORDER BY id DESC",
            account_ids,
        ).fetchall()
        ledger_paid = c.execute(
            f"SELECT COALESCE(SUM(amount),0) FROM customer_payments WHERE user_id IN ({marks})",
            account_ids,
        ).fetchone()[0]
    items = [(r, order_fees(r, rate)) for r in rows]
    debt = sum(f["commission"] for _, f in items)
    legacy_paid = sum(f["paid"] for _, f in items)
    paid = legacy_paid + ledger_paid
    return {"items": items, "debt": debt, "paid": paid, "balance": debt - paid}


# ---------------- customer ledger / rates text ----------------
def customer_summary(order) -> dict:
    """حساب یکجای مشتری: جمع بدهی (کارمزدها)، پرداختی‌ها و مانده؛ سفارش‌های خطادار حساب نمی‌شوند."""
    rate = get_commission_percent()
    rows = [r for r in customer_orders(order) if r["status"] != "error"]
    debt = sum(order_fees(r, rate)["commission"] for r in rows)
    legacy_paid = sum(order_fees(r, rate)["paid"] for r in rows)
    ids = sorted({r["user_id"] for r in rows})
    ledger_paid = 0
    if ids:
        marks = ",".join("?" for _ in ids)
        with conn() as c:
            ledger_paid = c.execute(
                f"SELECT COALESCE(SUM(amount),0) FROM customer_payments WHERE user_id IN ({marks})", ids
            ).fetchone()[0]
    paid = legacy_paid + ledger_paid
    return {"debt": debt, "paid": paid, "balance": debt - paid}


def get_rates_text() -> str:
    with conn() as c:
        return _get(c, "rates_text", "")


def set_rates_text(text: str):
    with conn() as c:
        _set(c, "rates_text", text)


# ---------------- customer history ----------------
def customer_orders(order):
    """همه‌ی سفارش‌های همین مشتری (همان آیدی تلگرام یا همان شماره موبایل)، جدیدترین اول."""
    phone = json.loads(order["data"]).get("phone", "")
    with conn() as c:
        return c.execute(
            "SELECT * FROM orders WHERE user_id=? OR (?!='' AND data LIKE ?) ORDER BY id DESC",
            (order["user_id"], phone, f'%"phone": "{phone}"%'),
        ).fetchall()


# ---------------- internal notes ----------------
def add_note(order_id, admin_id, text):
    with conn() as c:
        c.execute(
            "INSERT INTO order_notes(order_id, admin_id, text, at) VALUES(?,?,?,?)",
            (order_id, admin_id, text, now()),
        )
        _log(c, order_id, admin_id, "note", text[:100])


def list_notes(order_id, limit=5):
    """آخرین یادداشت‌ها؛ جدیدترین آخر لیست است."""
    with conn() as c:
        rows = c.execute(
            "SELECT * FROM order_notes WHERE order_id=? ORDER BY id DESC LIMIT ?",
            (order_id, limit),
        ).fetchall()
    return rows[::-1]


def count_notes(order_id) -> int:
    with conn() as c:
        return c.execute(
            "SELECT COUNT(*) FROM order_notes WHERE order_id=?", (order_id,)
        ).fetchone()[0]


# ---------------- duplicate / similar detection ----------------
DUP_WINDOW_HOURS = 24


def find_similar(order):
    """سفارش‌های مشابه (غیر خطادار؛ از DUP_WINDOW_HOURS ساعت قبل از این سفارش تا حالا، یعنی سفارش‌های بعدی هم دیده می‌شوند).
    هم‌صاحب = همان آیدی تلگرام یا همان شماره موبایل.
    برمی‌گرداند [(row, level)] که level='exact' (نماد، نوع، تعداد، قیمت یکسان)
    یا 'similar' (نماد و نوع یکسان) است."""
    data = json.loads(order["data"])
    phone = data.get("phone", "")
    since = (
        datetime.fromisoformat(order["created_at"]) - timedelta(hours=DUP_WINDOW_HOURS)
    ).isoformat(timespec="seconds")
    with conn() as c:
        rows = c.execute(
            "SELECT * FROM orders WHERE id!=? AND status!='error' AND created_at>=?"
            " AND (user_id=? OR (?!='' AND data LIKE ?)) ORDER BY id DESC",
            (order["id"], since, order["user_id"], phone, f'%"phone": "{phone}"%'),
        ).fetchall()
    out = []
    for r in rows:
        d = json.loads(r["data"])
        if d.get("symbol", "").strip() != data.get("symbol", "").strip():
            continue
        if d.get("side") != data.get("side"):
            continue
        same = d.get("quantity") == data.get("quantity") and d.get("price") == data.get("price")
        out.append((r, "exact" if same else "similar"))
    return out


# ---------------- backup / restore ----------------
REQUIRED_TABLES = {"orders", "order_log", "faq"}
REQUIRED_ORDER_COLS = {"id", "code", "user_id", "status", "data", "created_at", "updated_at"}
KEEP_SAFETY_COPIES = 3


def make_backup() -> io.BytesIO:
    """یک کپی هماهنگ (consistent) از دیتابیس زنده با API رسمی SQLite."""
    fd, tmp = tempfile.mkstemp(suffix=".db")
    os.close(fd)
    try:
        src = sqlite3.connect(config.DB_PATH)
        dst = sqlite3.connect(tmp)
        try:
            src.backup(dst)
        finally:
            dst.close()
            src.close()
        with open(tmp, "rb") as f:
            return io.BytesIO(f.read())
    finally:
        os.remove(tmp)


def validate_backup(path):
    """برمی‌گرداند (خلاصه‌ی محتوا, پیام خطا)."""
    try:
        c = sqlite3.connect(Path(path).resolve().as_uri() + "?mode=ro", uri=True)
        try:
            if c.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
                return None, "فایل آسیب دیده است."
            tables = {r[0] for r in c.execute("SELECT name FROM sqlite_master WHERE type='table'")}
            if not REQUIRED_TABLES <= tables:
                return None, "این فایل، بکاپ این ربات نیست."
            cols = {r[1] for r in c.execute("PRAGMA table_info(orders)")}
            if not REQUIRED_ORDER_COLS <= cols:
                return None, "ساختار جدول سفارش‌ها در این فایل معتبر نیست."
            return {
                "orders": c.execute("SELECT COUNT(*) FROM orders").fetchone()[0],
                "faq": c.execute("SELECT COUNT(*) FROM faq").fetchone()[0],
                "last_order": c.execute("SELECT MAX(created_at) FROM orders").fetchone()[0],
            }, None
        finally:
            c.close()
    except sqlite3.DatabaseError:
        return None, "فایل ارسالی دیتابیس معتبر نیست."


def restore_backup(path) -> str:
    """قبل از جایگزینی، از وضعیت فعلی یک کپی ایمنی کنار دیتابیس ذخیره می‌کند
    و مسیر آن را برمی‌گرداند. جدول‌ها/ستون‌های جدیدتر که در بکاپ قدیمی نیستند ساخته می‌شوند."""
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    safety = f"{config.DB_PATH}.before-restore-{ts}"
    with open(safety, "wb") as f:
        f.write(make_backup().getbuffer())
    src = sqlite3.connect(path)
    dst = sqlite3.connect(config.DB_PATH)
    try:
        src.backup(dst)
    finally:
        dst.close()
        src.close()
    init_db()
    for old in sorted(glob.glob(f"{config.DB_PATH}.before-restore-*"))[:-KEEP_SAFETY_COPIES]:
        os.remove(old)
    return safety


# ---------------- faq ----------------
def faq_list():
    with conn() as c:
        return c.execute("SELECT * FROM faq ORDER BY position, id").fetchall()


def faq_get(faq_id):
    with conn() as c:
        return c.execute("SELECT * FROM faq WHERE id=?", (faq_id,)).fetchone()


def faq_add(question, answer):
    with conn() as c:
        pos = c.execute("SELECT COALESCE(MAX(position), -1) + 1 FROM faq").fetchone()[0]
        c.execute("INSERT INTO faq(question, answer, position) VALUES(?,?,?)", (question, answer, pos))


def faq_update(faq_id, column, value):
    assert column in ("question", "answer")
    with conn() as c:
        c.execute(f"UPDATE faq SET {column}=? WHERE id=?", (value, faq_id))


def faq_delete(faq_id):
    with conn() as c:
        c.execute("DELETE FROM faq WHERE id=?", (faq_id,))