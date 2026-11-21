# -*- coding: utf-8 -*-
"""Firma vakilining buyurtma so'rovlari (10-talab, dizayn/BUYURTMA-SOROV.md): ma'lumotlar qatlami.

Bot (Telegram) va panel shu funksiyalarni chaqiradi; bu modul faqat sqlite bilan ishlaydi.
Holatlar: so'rov — "kutmoqda" | "javob" (admin javob yubordi) | "rad"; dori qatori — "kutmoqda" | "qabul" | "ozgardi" | "rad".
"""
import datetime

SCHEMA = """
CREATE TABLE IF NOT EXISTS order_req(
    id INTEGER PRIMARY KEY AUTOINCREMENT, chat_id INTEGER NOT NULL, firm TEXT, name TEXT, created_at TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'kutmoqda', resolved_by INTEGER, resolved_name TEXT, resolved_at TEXT);
CREATE TABLE IF NOT EXISTS order_req_items(
    id INTEGER PRIMARY KEY AUTOINCREMENT, req_id INTEGER NOT NULL, pos INTEGER NOT NULL, name TEXT NOT NULL,
    producer TEXT, qty_asked REAL NOT NULL, qty_ok REAL, status TEXT NOT NULL DEFAULT 'kutmoqda', note TEXT);
CREATE INDEX IF NOT EXISTS order_req_items_req ON order_req_items(req_id);
CREATE INDEX IF NOT EXISTS order_req_chat ON order_req(chat_id, created_at);
"""

OPEN, ANSWERED, REJECTED = "kutmoqda", "javob", "rad"
I_WAIT, I_OK, I_CHANGED, I_REJ = "kutmoqda", "qabul", "ozgardi", "rad"


def _now():
    return datetime.datetime.now().isoformat(timespec="seconds")


def init(db):
    db.executescript(SCHEMA)
    db.commit()


def create(db, chat_id, firm, name, items, ts=None):
    """items: [{"name", "producer", "qty"}] (qty > 0). Yangi so'rov raqami (id) qaytadi."""
    items = [i for i in items if float(i.get("qty") or 0) > 0]
    if not items:
        raise ValueError("Savatda dori yo'q")
    cur = db.execute("INSERT INTO order_req(chat_id, firm, name, created_at) VALUES(?,?,?,?)",
                     (chat_id, firm, name, ts or _now()))
    rid = cur.lastrowid
    db.executemany("INSERT INTO order_req_items(req_id, pos, name, producer, qty_asked) VALUES(?,?,?,?,?)",
                   [(rid, n, i["name"], i.get("producer") or "", float(i["qty"])) for n, i in enumerate(items, 1)])
    db.commit()
    return rid


def get(db, rid):
    r = db.execute("SELECT * FROM order_req WHERE id=?", (rid,)).fetchone()
    if not r:
        return None, []
    return r, db.execute("SELECT * FROM order_req_items WHERE req_id=? ORDER BY pos", (rid,)).fetchall()


def decide_item(db, rid, pos, qty=None, reject=False):
    """Admin bitta dori bo'yicha qaror: qty — tasdiqlangan son (None — so'ralgancha), reject — olib tashlash.
    So'rov allaqachon javob berilgan bo'lsa False."""
    r, items = get(db, rid)
    if not r or r["status"] != OPEN:
        return False
    it = next((i for i in items if i["pos"] == pos), None)
    if not it:
        return False
    if reject:
        db.execute("UPDATE order_req_items SET qty_ok=0, status=? WHERE id=?", (I_REJ, it["id"]))
    else:
        q = it["qty_asked"] if qty is None else max(0.0, float(qty))
        st = I_REJ if q <= 0 else (I_OK if abs(q - it["qty_asked"]) < 1e-9 else I_CHANGED)
        db.execute("UPDATE order_req_items SET qty_ok=?, status=? WHERE id=?", (q, st, it["id"]))
    db.commit()
    return True


def accept_all(db, rid):
    r, items = get(db, rid)
    if not r or r["status"] != OPEN:
        return False
    for it in items:
        if it["status"] == I_WAIT:
            db.execute("UPDATE order_req_items SET qty_ok=qty_asked, status=? WHERE id=?", (I_OK, it["id"]))
    db.commit()
    return True


def all_decided(items):
    return bool(items) and all(i["status"] != I_WAIT for i in items)


def finish(db, rid, by_id, by_name, reject=False):
    """Javobni yuborish (yoki butun so'rovni rad etish). Birinchi javob bergan admin hal qiladi:
    so'rov allaqachon yopilgan bo'lsa False (ikkinchi admin)."""
    r, items = get(db, rid)
    if not r or r["status"] != OPEN:
        return False
    if reject:
        db.execute("UPDATE order_req_items SET qty_ok=0, status=? WHERE req_id=?", (I_REJ, rid))
    elif not all_decided(items):
        return False
    cur = db.execute("UPDATE order_req SET status=?, resolved_by=?, resolved_name=?, resolved_at=? WHERE id=? AND status=?",
                     (REJECTED if reject else ANSWERED, by_id, by_name, _now(), rid, OPEN))
    db.commit()
    return cur.rowcount == 1


def totals(items):
    asked = sum(i["qty_asked"] for i in items)
    ok = sum((i["qty_ok"] or 0) for i in items if i["status"] in (I_OK, I_CHANGED))
    changed = sum(1 for i in items if i["status"] == I_CHANGED)
    return asked, ok, changed


def mine(db, chat_id, limit=20):
    return db.execute("SELECT * FROM order_req WHERE chat_id=? ORDER BY id DESC LIMIT ?", (chat_id, limit)).fetchall()


def open_requests(db):
    return db.execute("SELECT * FROM order_req WHERE status=? ORDER BY id", (OPEN,)).fetchall()


def recent(db, days=7, now=None):
    since = ((now or datetime.datetime.now()) - datetime.timedelta(days=days)).isoformat(timespec="seconds")
    return db.execute("SELECT * FROM order_req WHERE created_at>=? ORDER BY id DESC", (since,)).fetchall()


def approved_rows(db, since=None, firm=None):
    """Excel uchun: tasdiqlangan qatorlar (firma, sana, dori, so'ralgan, tasdiqlangan)."""
    q = ("SELECT r.id, r.firm, r.name AS who, r.resolved_at, i.name, i.producer, i.qty_asked, i.qty_ok, i.status "
         "FROM order_req r JOIN order_req_items i ON i.req_id=r.id WHERE r.status=? AND i.status IN (?,?)")
    a = [ANSWERED, I_OK, I_CHANGED]
    if since:
        q += " AND r.resolved_at>=?"
        a.append(since)
    if firm:
        q += " AND r.firm=?"
        a.append(firm)
    return db.execute(q + " ORDER BY r.id, i.pos", a).fetchall()
