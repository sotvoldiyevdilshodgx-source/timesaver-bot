# -*- coding: utf-8 -*-
"""
TimeSaver server paneli (brauzerda ochiladigan oyna).

- Bot bilan BIR jarayonda ishlaydi (timesaver_bot.py ishga tushganda o'zi ochiladi).
- Faqat shu kompyuterdan ochiladi: http://127.0.0.1:8080  (config.ini [panel]).
- SQL Serverga ulanmaydi: bot.db va Excel fayl bilan ishlaydi.
"""
import datetime
import glob
import html
import json
import logging
import os
import re
import threading
import time
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import abc_tahlil as ABC
import timesaver_bot as B
import buyurtma as BY
import fom_sql as FS
import uzkir as UK

PANEL_DIR = B.APP_DIR  # panel.html shu yerda (eski nom, testlar uchun)
LOOPBACK = {"127.0.0.1", "localhost", "::1"}
MAX_UPLOAD = 300 * 1024 * 1024
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# bosh sahifa kartalari: kunlik surat (14 kunlik trend chizig'i uchun)
STAT_KEYS = ("srok", "old", "users", "pending")
DAILY_STATS_SQL = "CREATE TABLE IF NOT EXISTS daily_stats(sana TEXT PRIMARY KEY, srok INT, old INT, users INT, pending INT)"
TREND_DAYS = 14

# ko'chirish ABC: «bajarildi» / «bajarilmadi» (bekor — hisobga kirmaydi)
KOCH_DONE = ("berildi", "olindi")
KOCH_NOT_DONE = ("yangi", "yopildi", "rad")
ABC_PERIODS = {"bugun": (0, 1), "kecha": (1, 1), "7": (0, 7)}  # (necha kun oldin tugaydi, necha kun)

# avtomatik (har soatlik) yangilash: daqiqa, 0 — o'chiq (AVTO-ANIQLASH.md §6)
AUTO_REFRESH_EVERY = (0, 30, 60, 120)
AUTO_REFRESH_DEFAULTS = {"auto_refresh_every": "0", "auto_refresh_from": "08:00", "auto_refresh_to": "20:00"}
HM_RE = re.compile(r"^([01]?\d|2[0-3]):[0-5]\d$")

# buyurtma so'rovi qatorining holati (Excel uchun, kirillda)
OREQ_ITEM_LBL = {BY.I_WAIT: "Кутмоқда", BY.I_OK: "Қабул қилинди", BY.I_CHANGED: "Ўзгартирилди", BY.I_REJ: "Рад этилди"}


def cyr_keys(obj, keys):
    """Faqat ko'rsatilgan kalitlardagi tizim matnlarini kirillga o'tkazadi (KPI sarlavhalari, yorliqlar).
    Dori / firma / odam nomlari bu kalitlarda bo'lmaydi, shuning uchun ular o'zgarmaydi."""
    return UK.deep(obj, keys=tuple(keys))


def read_db_ini(path=None):
    """db.ini (Excel makrosi o'qiydigan SQL ulanish): kalit=qiymat, bo'limsiz. Kalitlar kichik harfda qaytadi."""
    out = {}
    try:
        text = FS.read_text_ansi(path or B.DB_INI)
    except OSError:
        return out
    for ln in text.splitlines():
        k, sep, v = ln.partition("=")
        if sep and not k.strip().startswith((";", "#", "[")):
            out[k.strip().lower()] = v.strip()
    return out


def write_db_ini(server, database, uid, pwd, path=None):
    """Makros (VBA IniValue) bilan mos: Server/Database/Uid/Pwd. Uid bo'sh — Windows autentifikatsiyasi."""
    path = path or B.DB_INI
    os.makedirs(os.path.dirname(path), exist_ok=True)
    lines = ["; TimeSaver: SQL ulanish (Excel makrosi va bot uchun)", "Server=%s" % server, "Database=%s" % database]
    if uid:
        lines += ["Uid=%s" % uid, "Pwd=%s" % pwd]
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="mbcs", errors="replace", newline="\r\n") as f:  # VBA Line Input — Windows ANSI
        f.write("\n".join(lines) + "\n")
    os.replace(tmp, path)


def sql_conn_str(server, database, uid, pwd):
    """Makros va fom_sql bilan bir xil drayver ({SQL Server}). Uid bo'sh — Windows autentifikatsiyasi."""
    cs = "Driver={SQL Server};Server=%s;Database=%s;" % (server, database)
    return cs + (("Uid=%s;Pwd=%s;" % (FS.odbc_quote(uid), FS.odbc_quote(pwd))) if uid else "Trusted_Connection=yes;")


def _pyodbc():
    try:
        import pyodbc
        return pyodbc
    except ImportError:
        return None


def sql_test(server, database, uid, pwd):
    """SQL Serverga ulanishni sinash (pyodbc bo'lsa). {ok, error}"""
    pyodbc = _pyodbc()
    if pyodbc is None:
        return {"ok": True, "error": "", "skipped": True}
    cs = sql_conn_str(server, database, uid, pwd)
    try:
        con = pyodbc.connect(cs, timeout=8)
        try:
            con.cursor().execute("SELECT 1").fetchone()
        finally:
            con.close()
        return {"ok": True, "error": ""}
    except Exception as e:  # noqa
        return {"ok": False, "error": str(e)[:300]}


# ---------- sozlash ustasi: SQL serverni avtomatik topish ----------
SQL_INSTANCES_KEY = r"SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL"
SQL_LOGIN_TIMEOUT = 4  # soniya: har bir nomzod server uchun


def sql_local_instances(winreg=None):
    """Shu kompyuterga o'rnatilgan SQL Server nusxalari (reestrdan, faqat o'qiladi): MSSQLSERVER → «.», boshqalari → «.\\NOM»."""
    if winreg is None:
        try:
            import winreg
        except ImportError:  # Windows emas
            return []
    out = []
    for view in (getattr(winreg, "KEY_WOW64_64KEY", 0), getattr(winreg, "KEY_WOW64_32KEY", 0)):
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, SQL_INSTANCES_KEY, 0, winreg.KEY_READ | view) as k:
                i = 0
                while True:
                    try:
                        name = winreg.EnumValue(k, i)[0]
                    except OSError:
                        break
                    out.append("." if str(name).upper() == "MSSQLSERVER" else ".\\" + str(name))
                    i += 1
        except OSError:
            continue
    return out


def sql_candidates(ini_server="", instances=(), computer=None):
    """Sinab ko'riladigan serverlar tartibi: db.ini dagi server, reestrdagi nusxalar, keyin odatiy mahalliy nomlar.
    Takrorlar (katta-kichik harfdan qat'i nazar) olib tashlanadi."""
    if computer is None:
        computer = os.environ.get("COMPUTERNAME", "")
    seq = [ini_server, *instances, ".", "localhost", "(local)", ".\\SQLEXPRESS"]
    if computer:
        seq += [computer, computer + "\\SQLEXPRESS"]
    out, seen = [], set()
    for s in seq:
        s = str(s or "").strip()
        if s and s.lower() not in seen:
            seen.add(s.lower())
            out.append(s)
    return out


def sql_databases(server, uid="", pwd="", odbc=None, timeout=SQL_LOGIN_TIMEOUT):
    """Serverga ulanib, foydalanuvchi bazalari ro'yxatini oladi (tizim bazalari — master, tempdb, model, msdb — kirmaydi).
    (ro'yxat, None) yoki (None, xato matni)."""
    odbc = odbc or _pyodbc()
    if odbc is None:
        return None, "pyodbc o'rnatilmagan"
    try:
        con = odbc.connect(sql_conn_str(server, "master", uid, pwd), timeout=timeout)
        try:
            rows = con.cursor().execute("SELECT name FROM sys.databases WHERE database_id > 4 ORDER BY name").fetchall()
        finally:
            con.close()
        return [str(r[0]) for r in rows], None
    except Exception as e:  # noqa
        return None, str(e)[:300]


def sql_detect(ini=None, instances=None, odbc=None, computer=None):
    """Mahalliy SQL Serverni topadi: nomzodlarni parallel sinaydi, tartibdagi birinchi ishlagani tanlanadi.
    db.ini da login bo'lsa — shu login bilan, aks holda Windows orqali. pyodbc bo'lmasa faqat nomzodlar qaytadi."""
    ini = read_db_ini() if ini is None else ini
    instances = sql_local_instances() if instances is None else instances
    cands = sql_candidates(ini.get("server", ""), instances, computer)
    uid, pwd = ini.get("uid", ""), ini.get("pwd", "")
    base = {"candidates": cands, "auth": "login" if uid else "windows", "uid": uid,
            "database": ini.get("database", ""), "instances": list(instances)}
    odbc = odbc or _pyodbc()
    if odbc is None:
        return {"ok": False, "pyodbc": False, "server": cands[0] if (ini.get("server") or instances) else "", "databases": [],
                "error": "Bu kompyuterda SQL tekshiruvi o'rnatilmagan. Server nomini qo'lda kiriting.", **base}
    import concurrent.futures as cf
    ex = cf.ThreadPoolExecutor(max_workers=min(8, len(cands)))
    futs = [ex.submit(sql_databases, s, uid, pwd, odbc) for s in cands]
    ex.shutdown(wait=False)  # qolgan sinovlar fonda tugaydi; javob tartibdagi birinchi ishlagan server bilan darhol qaytadi
    deadline, tried = time.time() + SQL_LOGIN_TIMEOUT * 3 + 5, []
    for s, f in zip(cands, futs):
        try:
            dbs, err = f.result(timeout=max(0.0, deadline - time.time()))
        except cf.TimeoutError:
            dbs, err = None, "javob kelmadi"
        if dbs is not None:
            return {"ok": True, "pyodbc": True, "server": s, "databases": dbs, "tried": tried, **base}
        tried.append({"server": s, "detail": err})  # «detail» — ODBC matni, kirillga o'girilmaydi
    return {"ok": False, "pyodbc": True, "server": "", "databases": [], "tried": tried,
            "error": "SQL server topilmadi. Server nomini qo'lda kiriting.", **base}


# ---------- sozlash ustasi: birinchi /start bosgan odam — admin ----------
SETUP_TOKEN_AT = "setup_token_at"   # bot.db settings: ustada token kiritilgan vaqt (unix soniya)
SETUP_STARTS = "setup_starts"       # ustada /start bosganlar (JSON): qo'shimcha adminlar ro'yxati uchun
SETUP_AUTO_ADMIN = "setup_auto_admin"


def _start_text(m):
    t = (m.get("text") or "").strip()
    return t == "/start" or t.startswith("/start ") or t.startswith("/start@")


def setup_starts(bot):
    try:
        v = json.loads(bot.setting(SETUP_STARTS) or "[]")
        return v if isinstance(v, list) else []
    except ValueError:
        return []


def auto_admin(bot, m):
    """Sozlash ustasi: botga /start yozgan BIRINCHI odam avtomatik admin bo'ladi. Faqat:
    shaxsiy chat, ustasi tugamagan, adminlar 0 ta va /start token ustada kiritilgandan KEYIN yozilgan (xabar sanasi).
    Shartlardan biri bajarilmasa — hech narsa qilmaydi (odam faqat «qo'shimcha adminlar» ro'yxatiga yoziladi; uni
    paneldan qo'lda admin qilish mumkin). True — shu xabar bilan admin qilindi."""
    try:
        chat, frm = m.get("chat") or {}, m.get("from") or {}
        if chat.get("type") != "private" or not _start_text(m) or frm.get("is_bot") or frm.get("id") != chat.get("id"):
            return False
        if bot.setup_done:
            return False
        cid = int(chat["id"])
        name = " ".join(x for x in (frm.get("first_name"), frm.get("last_name")) if x)[:60]
        date = int(m.get("date") or 0)
        if cid not in bot.admins:  # ustadagi nomzodlar ro'yxati (oxirgi 20 ta)
            lst = [x for x in setup_starts(bot) if x.get("id") != cid]
            lst.insert(0, {"id": cid, "name": name, "username": frm.get("username") or "", "date": date})
            bot.set_setting(SETUP_STARTS, json.dumps(lst[:20], ensure_ascii=False))
        try:
            at = int(float(bot.setting(SETUP_TOKEN_AT) or 0))
        except ValueError:
            at = 0
        if bot.admins or not at or date < at:
            return False
        bot.set_admins({cid})
        bot.set_setting(SETUP_AUTO_ADMIN, str(cid))
        logging.warning("Sozlash ustasi: birinchi /start — %s (%d) avtomatik admin qilindi", name or "?", cid)
        try:
            bot.tg.send(cid, "✅ Siz <b>%s</b> botida admin qilindingiz. Sozlashni kompyuterdagi panelda davom ettiring." % B.esc(bot.app_name))
        except Exception:  # noqa
            pass
        return True
    except Exception:  # noqa
        logging.exception("Sozlash ustasi: avto-admin xatosi")
        return False


def install_auto_admin(bot):
    """bot.on_message ni o'raydi (bot faylini o'zgartirmasdan): avval auto_admin, keyin odatiy ishlov. Admin bo'lgan odam
    shu /start bilan darhol admin menyusini oladi. Bir bot uchun bir marta."""
    if getattr(bot, "_auto_admin_hook", False):
        return
    orig = bot.on_message

    def on_message(m):
        auto_admin(bot, m)
        return orig(m)

    bot.on_message = on_message
    bot._auto_admin_hook = True


class Api:
    def __init__(self, bot, cfg):
        self.bot = bot
        self.cfg = cfg
        self.bc = {"running": False, "kind": "", "total": 0, "ok": 0, "skip": 0, "fail": 0, "t1": None}
        self.bc_lock = threading.Lock()
        self._snap_last = None  # oxirgi yozilgan kunlik surat (har 10 soniyada bir xil qatorni qayta yozmaslik uchun)
        self._me = {}  # bot haqida (getMe): ustada ism va @username ko'rsatish uchun
        self._me_loading = False
        install_auto_admin(bot)
        with bot.dblock:  # panel ochilganda: kunlik surat va buyurtma so'rovlari jadvallari
            bot.db.execute(DAILY_STATS_SQL)
            bot.db.commit()
            BY.init(bot.db)
        if bot.setting("dash_jamlangan") != "1":
            # «Срок яқин» / «N+ кун турган» endi qatorlar emas, dori bo'yicha jamlangan son — eski surat bilan solishtirilmasin
            with bot.dblock:
                bot.db.execute("UPDATE daily_stats SET srok=NULL, old=NULL")
                bot.db.commit()
            bot.set_setting("dash_jamlangan", "1")

    # ---------- umumiy ----------
    def state(self):
        b, d = self.bot, self.bot.data
        d.reload()
        with b.dblock:
            cnt = {s: b.db.execute("SELECT COUNT(*) FROM clients WHERE status=?", (s,)).fetchone()[0]
                   for s in ("active", "pending", "blocked")}
            inv = b.db.execute("SELECT COUNT(*) FROM invites").fetchone()[0]
            accreq = b.db.execute("SELECT COUNT(*) FROM acc_requests WHERE status='pending'").fetchone()[0]
            oreq = b.db.execute("SELECT COUNT(*) FROM order_req WHERE status=?", (BY.OPEN,)).fetchone()[0]
        when = d.when()
        rs = dict(b.rstate)
        rs["steps"] = {k: list(v) for k, v in (rs.get("steps") or {}).items()}
        rs["elapsed"] = int((rs.get("t1") or time.time()) - rs["t0"]) if rs.get("t0") else 0
        out = {
            "alive": time.time() - b.last_poll < 90,
            "conflict": time.time() - getattr(b, "conflict_at", 0) < 120,  # 409: shu token bilan ikkinchi bot nusxasi
            "when": B.dt_text(when) if when else "-",
            "when_hm": when.strftime("%H:%M") if when else "--:--",
            "when_date": when.strftime("%d.%m.%Y") if when else "",
            "age_hours": round(d.age_hours(), 1) if when else None,
            "stale_hours": float(b.cfg.get("auto", "stale_hours", fallback="72")),
            "error": d.error,
            # bosh sahifa: dori bo'yicha jamlangan (barcha filiallar) — Srok va N+ kun (N — old_days, odatda 90)
            "srok": len(self.dash_rows("s")), "srok_days": b.srok_days,
            "old": len(self.dash_rows("o")), "old_days": self.old_days(),
            "acc_pending": accreq, "oreq_pending": oreq, "users": cnt["active"], "pending": cnt["pending"], "blocked": cnt["blocked"], "invited": inv,
            "refresh": rs,
            "bc": dict(self.bc),
            "auto_on": b.auto_on, "auto_time": b.auto_time.strftime("%H:%M"), "last_auto": b.setting("last_auto", "—"),
            "excel": d.path or d.file_pattern(),
            "today": datetime.date.today().strftime("%d.%m.%Y"),
            "token_ok": b.token_ok(), "no_admins": not b.admins,
            "maint_on": b.maint_on, "maint_since": b.setting("maint_since", "—"), "maint_waiting": len(b.maint_waiting()),
            "koch": self._koch_brief(),
        }
        trend = self._trend({k: out[k] for k in STAT_KEYS})
        if trend:
            out["trend"] = trend
        return out

    def _trend(self, vals, days=TREND_DAYS):
        """Bugungi sonlarni daily_stats ga yozadi (upsert) va oxirgi `days` kunlik qatorni qaytaradi:
        {srok: [..], old: [..], users: [..], pending: [..]}. Bo'sh kunlar oldingi qiymat bilan to'ldiriladi;
        bazada 2 kundan kam surat bo'lsa None (panel trend chizig'ini ko'rsatmaydi)."""
        b = self.bot
        today = B.now().date()
        start = today - datetime.timedelta(days=days - 1)
        row = (today.isoformat(),) + tuple(int(vals[k] or 0) for k in STAT_KEYS)
        with b.dblock:
            if row != self._snap_last:
                b.db.execute("INSERT OR REPLACE INTO daily_stats(sana, %s) VALUES(?,?,?,?,?)" % ", ".join(STAT_KEYS), row)
                b.db.commit()
                self._snap_last = row
            rows = b.db.execute("SELECT * FROM daily_stats WHERE sana>=? AND sana<=? ORDER BY sana",
                                (start.isoformat(), today.isoformat())).fetchall()
            prev = b.db.execute("SELECT * FROM daily_stats WHERE sana<? ORDER BY sana DESC LIMIT 1", (start.isoformat(),)).fetchone()
        if len(rows) + (1 if prev else 0) < 2:
            return None
        by, cur = {r["sana"]: r for r in rows}, prev
        out = {k: [] for k in STAT_KEYS}
        for i in range(days):
            cur = by.get((start + datetime.timedelta(days=i)).isoformat(), cur)
            if cur is None:  # tarix boshlanmagan kunlar
                continue
            for k in STAT_KEYS:
                if cur[k] is not None:  # NULL — hisob usuli o'zgargan kungacha (dash_jamlangan) — trendga kirmaydi
                    out[k].append(int(cur[k]))
        return out

    def _koch_brief(self):
        """Bosh sahifa uchun: bugungi bajarilish % va kechikkanlar."""
        b = self.bot
        if not b.koch_on or not b.koch_has_tasks():
            return {"has": False, "on": b.koch_on}
        s = b.koch_summary()
        kpi = s["kpi"].values()
        jami, ol = sum(v["jami"] for v in kpi), sum(v["olindi"] for v in kpi)
        z = s["kpi"]["shoshilinch"]
        return {"has": True, "on": True, "pct": int(round(100.0 * ol / jami)) if jami else 0, "late": sum(v["kechikkan"] for v in kpi),
                "z_done": z["olindi"], "z_total": z["jami"]}

    # ---------- hisobni tasdiqlash so'rovlari ----------
    def acc_requests(self):
        b = self.bot
        b.data.reload()
        return b.acc_pending()

    def acc_request_action(self, p):
        try:
            rid = int(p.get("id"))
        except (TypeError, ValueError):
            return {"ok": False, "error": "Noto'g'ri so'rov"}
        approve = p.get("action") == "approve"
        res = self.bot.acc_resolve(rid, approve)
        if res is None:
            return {"ok": False, "error": "Bu so'rov allaqachon hal qilingan."}
        return {"ok": True, "msg": "Tasdiqlandi: mijoz hisobni ko'ra boshlaydi" if approve else "Rad etildi"}

    def tags(self):
        self.bot.data.reload()
        return sorted(self.bot.data.all_tags())

    def accounts(self, q):
        self.bot.data.reload()
        q = B.norm(q)
        names = self.bot.data.hisob_names()
        return [n for n in names if not q or q in B.norm(n)][:30]

    # ---------- foydalanuvchilar ----------
    def _user(self, r):
        b = self.bot
        return {"id": r["chat_id"], "name": r["name"] or "", "phone": r["phone"] or "", "username": r["username"] or "",
                "status": r["status"], "role": r["role"], "created": (r["created"] or "")[:16].replace("T", " "),
                "firms": b.firms_of(r["chat_id"]), "accounts": b.accounts_of(r["chat_id"]),
                "last_seen": (r["last_seen"] or "")[:16].replace("T", " "), "left": bool(r["left_at"]),
                "filial": B.norm(r["filial"]) if r["role"] == "filial" else ""}

    def users(self, status="", q=""):
        b = self.bot
        out = []
        with b.dblock:
            rows = b.db.execute("SELECT * FROM clients ORDER BY CASE status WHEN 'pending' THEN 0 WHEN 'active' THEN 1 ELSE 2 END, created DESC").fetchall()
            for r in rows:
                out.append(self._user(r))
            for r in b.db.execute("SELECT * FROM invites ORDER BY created DESC").fetchall():
                out.append({"id": "inv:" + r["phone9"], "name": r["name"] or "", "phone": r["phone"] or "", "username": "",
                            "status": "invited", "role": r["role"], "created": (r["created"] or "")[:16].replace("T", " "),
                            "firms": [x for x in (r["firms"] or "").split(";") if x],
                            "accounts": [x for x in (r["accounts"] or "").split(";") if x],
                            "filial": (r["filial"] or "") if r["role"] == "filial" else ""})
        ql = B.norm(q)
        qd = B.digits(q)
        res = []
        for u in out:
            if status in ("staff", "filial"):
                if u["role"] != status or u["status"] == "invited":
                    continue
            elif status and u["status"] != status:
                continue
            if ql:
                hay = B.norm(" ".join([u["name"], u["phone"], u["username"], u["filial"]] + u["firms"]))
                if ql not in hay and not (qd and qd in B.digits(u["phone"])):
                    continue
            if u["status"] == "pending":
                u["suggest"] = b.suggest_tags(u["phone"])
            res.append(u)
        return res

    @staticmethod
    def _list(v):
        if isinstance(v, str):
            v = re.split(r"[;\n|]", v)
        return [B.norm(x) for x in (v or []) if str(x).strip()]

    def _role(self, p):
        """(rol, filial yoki None, xato yoki None). Filial faqat Ko'chirish varag'idagi ro'yxatdan."""
        role = p.get("role") if p.get("role") in ("staff", "filial") else "firm"
        if role != "filial":
            return role, None, None
        fil = B.norm(p.get("filial"))
        self.bot.data.reload()
        if fil not in {B.norm(x) for x in self.bot.data.koch_filials()}:
            return role, None, "Filialni ro'yxatdan tanlang."
        return role, fil, None

    def filials(self):
        self.bot.data.reload()
        return self.bot.data.koch_filials()

    def user_save(self, p):
        """Mavjud mijozni saqlash (ism, rol, firmalar, hisoblar) va ixtiyoriy ravishda faollashtirish."""
        b = self.bot
        cid = int(p["id"])
        c = b.client(cid)
        if not c:
            raise ValueError("Mijoz topilmadi")
        role, fil, err = self._role(p)
        if err:
            return {"ok": False, "error": err}
        firms, accs = self._list(p.get("firms")), self._list(p.get("accounts"))
        if role == "filial":
            firms = []
        with b.dblock:
            b.db.execute("UPDATE clients SET name=? WHERE chat_id=?", ((p.get("name") or c["name"] or "").strip(), cid))
            b.db.commit()
        if role != c["role"] or B.norm(c["filial"]) != (fil or ""):
            b.set_role(cid, role, fil)
        with b.dblock:
            b.db.execute("DELETE FROM client_accounts WHERE chat_id=?", (cid,))
            for a in accs:
                b.db.execute("INSERT OR IGNORE INTO client_accounts VALUES(?,?)", (cid, a))
            b.db.commit()
        b.set_firms(cid, firms)
        c = b.client(cid)
        missing = [t for t in firms if t not in set(b.data.all_tags())]
        msg = "Saqlandi"
        if p.get("approve") and c["status"] != "active":
            if not b.can_use(c):
                return {"ok": False, "error": "Faollashtirish uchun firma belgisini kiriting yoki «Xodim» / «Filial xodimi» ni tanlang."}
            b.activate(cid)
            msg = "Tasdiqlandi, mijozga xabar yuborildi"
        return {"ok": True, "msg": msg, "missing": missing}

    def user_action(self, p):
        b = self.bot
        act = p.get("action")
        uid = str(p.get("id"))
        if uid.startswith("inv:"):
            if act == "delete":
                with b.dblock:
                    b.db.execute("DELETE FROM invites WHERE phone9=?", (uid[4:],))
                    b.db.commit()
                return {"ok": True, "msg": "Taklif o'chirildi"}
            return {"ok": False, "error": "Bu amal taklif uchun mavjud emas"}
        cid = int(uid)
        c = b.client(cid)
        if not c:
            return {"ok": False, "error": "Mijoz topilmadi"}
        if act == "block":
            b.set_status(cid, "blocked")
            return {"ok": True, "msg": "Bloklandi"}
        if act == "unblock":
            if not b.can_use(c):
                return {"ok": False, "error": "Avval firma belgisini kiriting."}
            b.activate(cid)
            return {"ok": True, "msg": "Faollashtirildi"}
        if act == "delete":
            b.delete_client(cid)
            return {"ok": True, "msg": "O'chirildi"}
        return {"ok": False, "error": "Noma'lum amal"}

    def user_add(self, p):
        """Qo'lda qo'shish: telefon raqami bo'yicha taklif. Mijoz botga kirib kontaktini yuborganda avtomatik qabul qilinadi."""
        b = self.bot
        phone = (p.get("phone") or "").strip()
        d9 = B.digits(phone)[-9:]
        if len(d9) < 9:
            return {"ok": False, "error": "Telefon raqamini to'liq kiriting (masalan +998 90 123 45 67)."}
        role, fil, err = self._role(p)
        if err:
            return {"ok": False, "error": err}
        firms, accs = self._list(p.get("firms")), self._list(p.get("accounts"))
        if role == "filial":
            firms = []
        elif role != "staff" and not firms:
            return {"ok": False, "error": "Firma belgisini kiriting yoki «Xodim» / «Filial xodimi» ni tanlang."}
        name = (p.get("name") or "").strip()
        with b.dblock:
            row = None
            for r in b.db.execute("SELECT * FROM clients").fetchall():
                if B.digits(r["phone"])[-9:] == d9:
                    row = r
                    break
        if row:  # mijoz allaqachon botga kirgan: to'g'ridan-to'g'ri yangilab faollashtiramiz
            res = self.user_save({"id": row["chat_id"], "name": name or row["name"], "role": role, "firms": firms,
                                  "accounts": accs, "approve": True, "filial": fil})
            if res.get("ok"):
                res["msg"] = "Mijoz botda bor edi: ma'lumotlari yangilandi va faollashtirildi"
            return res
        with b.dblock:
            b.db.execute("INSERT OR REPLACE INTO invites(phone9, phone, name, role, firms, accounts, created, filial) "
                         "VALUES(?,?,?,?,?,?,?,?)",
                         (d9, phone, name, role, ";".join(firms), ";".join(accs), B.now().isoformat(timespec="seconds"), fil))
            b.db.commit()
        missing = [t for t in firms if t not in set(b.data.all_tags())]
        return {"ok": True, "msg": "Qo'shildi. Foydalanuvchi botga kirib kontaktini yuborgach, avtomatik qabul qilinadi.",
                "missing": missing}

    # ---------- hisobotlar ----------
    def report_counts(self):
        b = self.bot
        b.data.reload()
        d = b.data
        return {"srok": len(d.srok_rows(None, b.srok_days)), "old": len(d.old_rows(None)), "need": len(b.need_rows(None)),
                "stock": len(d.stock_rows(None)), "narx": len(d.narx), "acc": len(d.hisob), "srok_days": b.srok_days}

    def report_file(self, kind, firm=""):
        b = self.bot
        b.data.reload()
        tags = {B.norm(firm)} if firm else None
        res = b.prepare(kind, tags, None)
        if not res:
            return None
        return res[0], res[1]

    # ---------- xabar yuborish ----------
    def bc_preview(self, kind):
        b = self.bot
        b.data.reload()
        withdata = nodata = 0
        for c in b.audience("all"):
            ok = b.prepare(kind, b.tags_for(c), b.accs_for_client(c))
            if ok:
                withdata += 1
            else:
                nodata += 1
        return {"with": withdata, "without": nodata, "when": B.dt_text(b.data.when())}

    def bc_audience_counts(self):
        b = self.bot
        return {x: len(b.audience(x)) for x in ("all", "firm", "staff")}

    def bc_send(self, p):
        b = self.bot
        kind = p.get("kind")
        text = (p.get("text") or "").strip()
        who = p.get("who") or "all"
        if kind == "text" and not text:
            return {"ok": False, "error": "Matn bo'sh."}
        if kind != "text" and kind not in b.KINDS:
            return {"ok": False, "error": "Noma'lum hisobot turi."}
        with b.rlock:  # Telegramdagi ommaviy yuborish bilan bir vaqtda ketmasin
            if b.bc_running or self.bc["running"]:
                return {"ok": False, "error": "Yuborish allaqachon ketmoqda."}
            b.bc_running = True
            self.bc = {"running": True, "kind": kind, "total": len(b.audience(who)), "ok": 0, "skip": 0, "fail": 0,
                       "gone": 0, "t1": None}
        threading.Thread(target=self._bc_job, args=(kind, B.simple_markup(text), who), daemon=True).start()
        return {"ok": True}

    def _bc_job(self, kind, text, who):
        b = self.bot
        try:
            for c in b.audience(who):
                self.bc[b.deliver(c, kind, text)] += 1
                time.sleep(0.05)
        except Exception:  # noqa
            logging.exception("Panel: ommaviy yuborish xatosi")
        finally:
            self.bc["running"] = False
            self.bc["t1"] = time.time()
            b.bc_running = False
            logging.info("Panel xabari tugadi (%s): %s", kind, {k: self.bc[k] for k in ("ok", "skip", "gone", "fail")})

    def send_one(self, p):
        """Bitta foydalanuvchiga: matn va/yoki bir nechta hisobot (har biri faqat uning firmasi bo'yicha)."""
        b = self.bot
        try:
            cid = int(p.get("id"))
        except (TypeError, ValueError):
            return {"ok": False, "error": "Foydalanuvchi tanlanmagan."}
        c = b.client(cid)
        if not c:
            return {"ok": False, "error": "Foydalanuvchi topilmadi."}
        if c["status"] != "active":
            return {"ok": False, "error": "Foydalanuvchi faol emas (holati: %s). Avval tasdiqlang." % c["status"]}
        if not b.can_use(c):
            return {"ok": False, "error": "Foydalanuvchiga firma belgisi biriktirilmagan."}
        kinds = [k for k in (p.get("kinds") or []) if k in b.KINDS]
        text = (p.get("text") or "").strip()
        if not kinds and not text:
            return {"ok": False, "error": "Matn yozing yoki kamida bitta hisobotni tanlang."}
        res = []
        if text:
            try:
                b.tg.send(cid, B.simple_markup(text))
                res.append({"what": "Matn", "ok": True, "msg": "yuborildi"})
            except Exception as e:  # noqa
                if b.mark_gone(cid, e):
                    return {"ok": False, "error": "Foydalanuvchi botni bloklagan (tark etgan), yuborib bo'lmaydi."}
                res.append({"what": "Matn", "ok": False, "msg": "yuborib bo'lmadi: %s" % e})
        b.data.reload()
        accs = b.accs_for_client(c)
        for k in kinds:
            label = re.sub(r"^\W+\s*", "", b.KINDS[k])
            if k == "acc" and accs is not None and not b.data.account_rows(accs):
                res.append({"what": label, "ok": False, "msg": "hisob tashkiloti topilmadi (firma telefoni mos kelmadi)"})
                continue
            try:
                ok = b.send_report(cid, k, b.tags_for(c), accs, quiet=True)
                res.append({"what": label, "ok": bool(ok), "msg": "yuborildi" if ok else "ma'lumot yo'q (yuborilmadi)"})
            except Exception as e:  # noqa
                logging.exception("Panel: bitta odamga yuborib bo'lmadi: %s", cid)
                res.append({"what": label, "ok": False, "msg": "xato: %s" % e})
            time.sleep(0.05)
        return {"ok": all(r["ok"] for r in res), "results": res, "name": c["name"] or c["phone"]}

    # ---------- sozlash ustasi (birinchi ishga tushish) ----------
    def _bot_me(self):
        """Bot ismi (getMe) fon oqimida bir marta olinadi: dastur qayta ishga tushganda @username bo'sh bo'ladi."""
        b = self.bot
        if self._me or self._me_loading or not b.token_ok():
            return
        self._me_loading = True

        def run():
            try:
                me = b.tg.call("getMe", timeout=10)
                if isinstance(me, dict):
                    self._me = me
                    b.username = b.username or me.get("username", "")
            except Exception:  # noqa
                pass
            finally:
                self._me_loading = False
        threading.Thread(target=run, daemon=True, name="getme").start()

    def setup_get(self):
        b, d = self.bot, self.bot.data
        d.reload()
        ini = read_db_ini()
        self._bot_me()
        with b.dblock:
            rows = b.db.execute("SELECT chat_id, name, username, phone, created FROM clients ORDER BY created DESC LIMIT 12").fetchall()
        starts = setup_starts(b)
        known = {x.get("id"): x for x in starts}
        # ustada /start bosganlar (kontakt yubormagan bo'lsa ham) birinchi, keyin kontakt yuborgan mijozlar
        cands = [{"id": x["id"], "name": x.get("name") or "", "username": x.get("username") or "",
                  "created": datetime.datetime.fromtimestamp(x["date"]).strftime("%Y-%m-%d %H:%M") if x.get("date") else ""}
                 for x in starts if isinstance(x.get("id"), int)]
        cands += [{"id": r["chat_id"], "name": r["name"] or r["phone"] or "", "username": r["username"] or "",
                   "created": (r["created"] or "")[:16].replace("T", " ")} for r in rows if r["chat_id"] not in known]
        cands = [c for c in cands if c["id"] not in b.admins][:8]
        names = {r["chat_id"]: r["name"] or r["phone"] or "" for r in rows}
        names.update({x.get("id"): x.get("name") or "" for x in starts})
        try:
            token_at = int(float(b.setting(SETUP_TOKEN_AT) or 0))
        except ValueError:
            token_at = 0
        path = d.find_path()
        rs = b.rstate
        return {"done": b.setup_done, "brand": B.BRAND, "version": B.VERSION, "name": b.app_get("name"),
                "sklad_name": b.cfg.get("excel", "sklad_name", fallback=""), "data_dir": B.BASE_DIR,
                "token_ok": b.token_ok(), "username": b.username, "bot_title": self._me.get("first_name", ""),
                "connected": time.time() - b.last_poll < 90,
                "admins": sorted(b.admins), "admins_info": [{"id": a, "name": names.get(a, "")} for a in sorted(b.admins)],
                "auto_admin": str(b.setting(SETUP_AUTO_ADMIN) or ""), "candidates": cands,
                # «Botni ochish → /start» kutilmoqda: token ustada kiritilgan, admin hali yo'q, ustasi tugamagan
                "waiting_admin": bool(token_at and not b.admins and not b.setup_done and b.token_ok()),
                "excel": {"path": path or d.file_pattern(), "found": bool(path), "error": d.error if path else None,
                          "when": B.dt_text(d.when()) if path and d.when() else "",
                          "rows": {"Sale": len(d.sale), "Srok": len(d.srok), "30+ kun": len(d.old), "Narx": len(d.narx),
                                   "Hisob": len(d.hisob), "Ko'chirish": len(d.koch)} if path else {}},
                "sql": {"ini": B.DB_INI, "exists": os.path.exists(B.DB_INI), "server": ini.get("server", ""),
                        "database": ini.get("database", ""), "uid": ini.get("uid", ""), "has_pwd": bool(ini.get("pwd"))},
                "refresh": {"running": bool(rs.get("running")), "ok": rs.get("ok"), "msg": rs.get("msg", ""),
                            "steps": {k: list(v) for k, v in (rs.get("steps") or {}).items()}}}

    def setup_token(self, p):
        """Ustada token: tekshiradi (getMe), saqlaydi va «birinchi /start — admin» oynasini ochadi (vaqt belgisi).
        Xato turi (kind): format — shakli noto'g'ri, rejected — Telegram qabul qilmadi, network — internet yo'q."""
        b = self.bot
        token = str(p.get("token") or "").strip()
        if not token:
            return {"ok": False, "kind": "format", "error": "Tokenni kiriting."}
        try:
            uname = b.set_token(token)
        except ValueError as e:
            s = str(e)
            if "shakli" in s:
                return {"ok": False, "kind": "format", "error": "Token noto'g'ri ko'rinishda. @BotFather bergan to'liq "
                                                                 "tokenni nusxalang (masalan 123456789:AAH…)."}
            if "qabul qilmadi" in s:
                return {"ok": False, "kind": "rejected", "error": "Telegram bu tokenni qabul qilmadi: u eskirgan yoki xato. "
                                                                   "@BotFather → /mybots → botingiz → «API Token» dan yangisini oling."}
            return {"ok": False, "kind": "network", "error": "Telegramga ulanib bo'lmadi. Internetni tekshirib, qayta urinib ko'ring."}
        b.set_setting(SETUP_TOKEN_AT, str(int(time.time())))
        self._me = {}
        try:
            me = b.tg.call("getMe", timeout=10)
            if isinstance(me, dict):
                self._me = me
        except Exception:  # noqa
            pass
        return {"ok": True, "username": uname, "title": self._me.get("first_name", "") or uname,
                "msg": "Bot ulandi: @%s" % uname}

    def setup_sql_detect(self, p=None):
        return sql_detect()

    def setup_sql_dbs(self, p):
        """Ko'rsatilgan serverdagi bazalar ro'yxati (ulanish tekshiruvi ham). Parol bo'sh va login db.ini dagi bilan bir xil —
        saqlangan parol ishlatiladi (parol panelga qaytmaydi)."""
        server = str(p.get("server") or "").strip()
        if not server:
            return {"ok": False, "error": "SQL server nomini kiriting."}
        old = read_db_ini()
        uid = str(p.get("uid") or "").strip()
        pwd = str(p.get("pwd") or "") or (old.get("pwd", "") if uid and uid == old.get("uid") else "")
        if _pyodbc() is None:
            return {"ok": False, "pyodbc": False, "databases": [],
                    "error": "Bu kompyuterda SQL tekshiruvi o'rnatilmagan. Baza nomini qo'lda kiriting."}
        dbs, err = sql_databases(server, uid, pwd)
        if dbs is None:
            return {"ok": False, "pyodbc": True, "databases": [], "detail": err,
                    "error": "SQL serverga ulanib bo'lmadi. Server nomi, login va parolni tekshiring."}
        return {"ok": True, "pyodbc": True, "databases": dbs, "msg": "Ulandi: %d ta baza" % len(dbs)}

    def setup_set(self, p):
        b = self.bot
        if "name" in p:
            name = re.sub(r"\s+", " ", str(p["name"])).strip()[:60]
            if not name:
                return {"ok": False, "error": "Dorixona nomini kiriting."}
            b.app_set("name", name)
        if "sklad_name" in p:
            v = B.norm(p["sklad_name"])[:60]
            B.write_ini("excel", "sklad_name", v)
            b.cfg["excel"]["sklad_name"] = v
            b.data.sklad_name = v or B.norm("ДОРИХОНА СКЛАД")
        if "sql" in p:
            s = p["sql"] or {}
            server, db = str(s.get("server", "")).strip(), str(s.get("database", "")).strip()
            if not server or not db:
                return {"ok": False, "error": "SQL server va baza nomini kiriting."}
            old = read_db_ini()
            uid = str(s.get("uid", "")).strip()
            pwd = str(s.get("pwd", "")) or (old.get("pwd", "") if uid and uid == old.get("uid") else "")
            try:
                write_db_ini(server, db, uid, pwd)
            except OSError as e:
                return {"ok": False, "error": "db.ini ga yozib bo'lmadi (%s). Dasturni administrator sifatida bir marta oching." % e}
            res = sql_test(server, db, uid, pwd) if s.get("test") else None
            if res and not res["ok"]:
                return {"ok": False, "detail": res["error"],  # ODBC matni alohida: kirillga o'girilib buzilmasin
                        "error": "Sozlama saqlandi, lekin SQL ga ulanib bo'lmadi. Server, baza nomi va loginni tekshiring."}
            return {"ok": True, "msg": "SQL sozlamasi saqlandi" + (" va ulanish tekshirildi ✅" if res else "")}
        if "excel_path" in p:
            path = str(p["excel_path"]).strip().strip('"')
            if not os.path.isfile(path):
                return {"ok": False, "error": "Fayl topilmadi: %s" % UK.mark(path)}
            return self._use_excel(path)
        if p.get("finish"):
            if not b.token_ok():
                return {"ok": False, "error": "Avval bot tokenini kiriting."}
            if not b.admins:
                return {"ok": False, "error": "Kamida bitta admin qo'shing."}
            if not b.data.find_path():
                return {"ok": False, "error": "Avval Excel faylni yuklang."}
            b.app_set("setup_done", "yes")
            return {"ok": True, "msg": "Sozlash tugadi. %s ishga tayyor!" % B.BRAND}
        return {"ok": True, "msg": "Saqlandi"}

    def setup_upload(self, fname, body):
        """Brauzerdan yuklangan Excel: ma'lumot papkasidagi baza\\ ga saqlanadi va bot shu fayldan o'qiy boshlaydi."""
        base = os.path.basename(str(fname or "")).strip()
        if not re.search(r"\.(xlsm|xlsx)$", base, re.I):
            return {"ok": False, "error": "Faqat .xlsm yoki .xlsx fayl yuklang (makrosli baza fayli — .xlsm)."}
        if len(body) < 1000:
            return {"ok": False, "error": "Fayl bo'sh yoki buzilgan."}
        base = re.sub(r'[\\/:*?"<>|]+', "_", base)
        folder = os.path.join(B.BASE_DIR, "baza")
        os.makedirs(folder, exist_ok=True)
        dest = os.path.join(folder, base)
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        # Avval vaqtinchalik nom bilan saqlab, tekshiramiz (kengaytma o'zgarmaydi — openpyxl shuni talab qiladi).
        # O'qilmasa — ishlab turgan fayl joyida qoladi; o'qilsa — eskisi .bak ga o'tadi va yangisi uning o'rnini oladi.
        tmp = os.path.join(folder, "_yuklash_%s_%s" % (stamp, base))
        with open(tmp, "wb") as f:
            f.write(body)
        b = self.bot
        old = b.cfg["excel"].get("file", "")
        b.cfg["excel"]["file"] = tmp
        ok = b.data.reload(force=True)
        err = b.data.error or "Excelni o'qib bo'lmadi"
        b.cfg["excel"]["file"] = old
        if not ok:
            b.data.reload(force=True)
            try:
                os.remove(tmp)
            except OSError:
                pass
            return {"ok": False, "error": "Faylni o'qib bo'lmadi: u buzilgan yoki kerakli varaqlar yo'q. Ishlab turgan fayl "
                                          "o'zgarmadi.", "detail": err}
        if os.path.exists(dest):  # eski faylni saqlab qo'yamiz
            os.replace(dest, dest + ".%s.bak" % stamp)
        os.replace(tmp, dest)
        self._prune_bak(dest)
        return self._use_excel(dest)

    @staticmethod
    def _prune_bak(dest, keep=3):
        """Har yuklashda to'liq nusxa qolardi (300 MB gacha): faqat oxirgi `keep` ta zaxira saqlanadi."""
        baks = sorted(glob.glob(glob.escape(dest) + ".*.bak"))
        for p in baks[:-keep]:
            try:
                os.remove(p)
            except OSError:
                pass

    def _use_excel(self, path):
        """Yangi faylni sinab o'qiydi; o'qilmasa eski fayl qoladi."""
        b = self.bot
        old = b.cfg["excel"].get("file", "")
        b.cfg["excel"]["file"] = path
        if not b.data.reload(force=True):
            err = b.data.error or "Excelni o'qib bo'lmadi"
            b.cfg["excel"]["file"] = old
            b.data.reload(force=True)
            return {"ok": False, "error": "Faylni o'qib bo'lmadi: u buzilgan yoki kerakli varaqlar yo'q. Ishlab turgan fayl "
                                          "o'zgarmadi.", "detail": err}
        B.write_ini("excel", "file", path)
        d = b.data
        return {"ok": True, "msg": "Excel ulandi: %s · sotuv %d, srok %d, narx %d qator" % (
            UK.mark(os.path.basename(path)), len(d.sale), len(d.srok), len(d.narx))}  # fayl nomi kirillga o'girilmaydi

    # ---------- sozlamalar, jurnal ----------
    def settings_get(self):
        b = self.bot
        tok = b.cfg["bot"].get("token", "").strip()
        ok = b.token_ok()
        return {"auto_on": b.auto_on, "auto_time": b.auto_time.strftime("%H:%M"), "srok_days": b.srok_days,
                "last_auto": b.setting("last_auto", "—"), "excel": b.data.path or b.data.file_pattern(),
                "token_ok": ok, "token": (tok[:6] + "…" + tok[-4:]) if ok and len(tok) > 12 else "",
                "username": b.username, "connected": time.time() - b.last_poll < 90,
                "admins": sorted(b.admins), "cover": b.cover,
                "weekly_old": "old" in b.weekly_kinds, "weekly_need": "need" in b.weekly_kinds,
                "weekly_day": b.weekly_day, "weekdays": b.WEEKDAYS, "last_weekly": b.setting("last_weekly", "—"),
                "last_backup": b.setting("last_backup", "—"),
                "maint_on": b.maint_on, "maint_text": self._to_markup(b.maint_text),
                "koch": {k: b.kset(k) for k in b.KOCH_DEFAULTS},
                "alerts": [{"key": k, "name": n, "on": b.alert_on(k)} for k, n in b.ALERTS],
                "hisob_internal": ", ".join(b.internal_patterns()),
                "auto_refresh_time": b.auto_refresh_time.strftime("%H:%M") if b.auto_refresh_time else "",
                "last_refresh_ok": b.setting("last_refresh_ok_date", "—"),
                "sale_days": b.sale_days, "period_days": b.data.period_days, **self._auto_refresh()}

    def _auto_refresh(self):
        """Har soatlik avtomatik yangilash sozlamalari (bot rejalashtiruvchisi keyin o'qiydi)."""
        b = self.bot
        v = {k: b.setting(k, d) for k, d in AUTO_REFRESH_DEFAULTS.items()}
        try:
            every = int(v["auto_refresh_every"])
        except (TypeError, ValueError):
            every = 0
        return {"auto_refresh_every": every if every in AUTO_REFRESH_EVERY else 0,
                "auto_refresh_from": v["auto_refresh_from"] if HM_RE.match(v["auto_refresh_from"] or "") else "08:00",
                "auto_refresh_to": v["auto_refresh_to"] if HM_RE.match(v["auto_refresh_to"] or "") else "20:00"}

    def settings_view(self):
        """GET /api/settings: settings_get + tizim yorliqlari kirillda (hafta kunlari, ogohlantirish nomlari)."""
        s = self.settings_get()
        s["weekdays"] = [UK.lat2cyr(x) for x in s["weekdays"]]
        s["alerts"] = cyr_keys(s["alerts"], ("name",))
        return s

    @staticmethod
    def _to_markup(t):
        """HTML → panel matni (**qalin**, __kursiv__), tahrirlash uchun."""
        t = re.sub(r"</?b>", "**", t or "")
        t = re.sub(r"</?i>", "__", t)
        return html.unescape(re.sub(r"<[^>]+>", "", t))

    def settings_set(self, p):
        b = self.bot
        out = {"ok": True, "msg": "Saqlandi"}
        if "auto_on" in p:
            b.set_setting("auto_on", "1" if p["auto_on"] else "0")
        if "srok_days" in p:
            n = int(p["srok_days"])
            if not 1 <= n <= 3650:
                return {"ok": False, "error": "Srok kuni 1 dan 3650 gacha bo'lishi kerak."}
            b.set_setting("srok_days", n)
        for k in ("weekly_old", "weekly_need"):
            if k in p:
                b.set_setting(k, "1" if p[k] else "0")
        if "weekly_day" in p:
            n = int(p["weekly_day"])
            if not 0 <= n <= 6:
                return {"ok": False, "error": "Hafta kuni noto'g'ri."}
            b.set_setting("weekly_day", n)
        for k, _ in b.ALERTS:
            if k in p:
                b.set_setting(k, "1" if p[k] else "0")
        if "hisob_internal" in p:
            v = ", ".join(x.strip() for x in str(p["hisob_internal"]).split(",") if x.strip())
            b.set_setting("hisob_internal", v)
            out["msg"] = "Ichki hisoblar ro'yxati saqlandi"
        if "sale_days" in p:
            try:
                n = int(p["sale_days"] or 0)
            except (TypeError, ValueError):
                return {"ok": False, "error": "Kunlar sonini kiriting (1–365) yoki 0 (Exceldagi davr)."}
            if not 0 <= n <= 365:
                return {"ok": False, "error": "Sotuv davri 1 dan 365 kungacha bo'lishi kerak (0 — Exceldagi davr)."}
            b.set_setting("sale_days", n)
            out["msg"] = ("Sotuv davri: %d kun. Keyingi «Yangilash» dan keyin kuchga kiradi" % n if n
                          else "Sotuv davri Excel'dagi sanalardan olinadi")
        if "auto_refresh_time" in p:
            v = str(p["auto_refresh_time"] or "").strip()
            if v and not re.match(r"^([01]?\d|2[0-3]):[0-5]\d$", v):
                return {"ok": False, "error": "Vaqtni HH:MM ko'rinishida kiriting (masalan 08:30) yoki bo'sh qoldiring."}
            b.set_setting("auto_refresh_time", v)
            out["msg"] = "Avto-yangilash: %s" % (v or "o'chirildi")
        if any(k in p for k in AUTO_REFRESH_DEFAULTS):  # har soatlik avtomatik yangilash
            cur = self._auto_refresh()
            try:
                every = int(p.get("auto_refresh_every", cur["auto_refresh_every"]) or 0)
            except (TypeError, ValueError):
                every = -1
            if every not in AUTO_REFRESH_EVERY:
                return {"ok": False, "error": "Oraliq: 30, 60 yoki 120 daqiqa (0 — o'chiq)."}
            fr = str(p.get("auto_refresh_from", cur["auto_refresh_from"]) or "").strip()
            to = str(p.get("auto_refresh_to", cur["auto_refresh_to"]) or "").strip()
            if not HM_RE.match(fr) or not HM_RE.match(to):
                return {"ok": False, "error": "Ish vaqtini HH:MM ko'rinishida kiriting (masalan 08:00)."}
            fr, to = b.parse_hm(fr).strftime("%H:%M"), b.parse_hm(to).strftime("%H:%M")
            if fr >= to:
                return {"ok": False, "error": "Ish vaqti boshlanishi tugashidan oldin bo'lishi kerak."}
            b.set_setting("auto_refresh_every", every)
            b.set_setting("auto_refresh_from", fr)
            b.set_setting("auto_refresh_to", to)
            out["msg"] = ("Avtomatik yangilash: har %d daqiqada, %s–%s" % (every, fr, to)) if every else "Avtomatik yangilash o'chirildi"
        for k in b.KOCH_DEFAULTS:  # ko'chirish nazorati sozlamalari
            if k not in p:
                continue
            v = p[k]
            if k == "koch_task_time" or k.startswith("koch_deadline_"):
                v = str(v or "").strip()
                if v and not b.parse_hm(v):
                    return {"ok": False, "error": "Vaqtni HH:MM ko'rinishida kiriting (masalan 10:00)."}
                if not v and k == "koch_task_time":
                    return {"ok": False, "error": "Bu vaqt bo'sh bo'lishi mumkin emas."}
                v = b.parse_hm(v).strftime("%H:%M") if v else ""
            elif k in b.KOCH_NUMS:
                name, lo, hi = b.KOCH_NUMS[k]
                if not str(v).strip().isdigit() or not lo <= int(v) <= hi:
                    return {"ok": False, "error": "%s: %d dan %d gacha." % (name, lo, hi)}
                v = str(int(v))
            else:
                v = "1" if v in (True, 1, "1", "yes", "true") else "0"
            b.set_setting(k, v)
        if "maint_text" in p:
            t = str(p["maint_text"]).strip()
            b.set_setting("maint_text", B.simple_markup(t) if t else "")
            out["msg"] = "Matn saqlandi"
        if "maint_on" in p:
            n = b.set_maint(bool(p["maint_on"]), by="panel")
            out["msg"] = ("🛠 Texnik rejim yoqildi: mijozlar botdan foydalana olmaydi" if p["maint_on"]
                          else "✅ Texnik rejim o'chirildi. Kutganlarga xabar yuborilmoqda: %d ta" % n)
        if p.get("backup_now"):
            out["msg"] = "Zaxira nusxa olindi: %s" % b.backup_db()
        if p.get("token"):
            try:
                uname = b.set_token(str(p["token"]))
            except ValueError as e:
                return {"ok": False, "error": str(e)}
            out["msg"] = "Token saqlandi: @%s. Bot yangi token bilan ishlay boshlaydi." % uname
        if "admin_add" in p:
            raw = str(p["admin_add"]).strip()
            if not re.fullmatch(r"\d{5,15}", raw):
                return {"ok": False, "error": "Telegram ID faqat raqamlardan iborat bo'ladi (masalan 1112223334). ID ni bilish uchun botga /myid yozing."}
            new = int(raw)
            if new in b.admins:
                return {"ok": False, "error": "Bu ID allaqachon admin."}
            b.set_admins(b.admins | {new})
            out["msg"] = "Admin qo'shildi: %d" % new
            if b.token_ok():
                try:
                    b.tg.send(new, "✅ Siz <b>%s</b> botida admin qilindingiz. Panelni ochish: /admin" % B.esc(b.app_name),
                              b.home_kb(admin=True))
                except Exception:  # noqa
                    out["msg"] += ". Unga xabar yuborib bo'lmadi: u avval botga /start yozishi kerak."
        if "admin_del" in p:
            rid = int(p["admin_del"])
            if rid not in b.admins:
                return {"ok": False, "error": "Bunday admin yo'q."}
            if len(b.admins) <= 1:
                return {"ok": False, "error": "Oxirgi adminni o'chirib bo'lmaydi."}
            b.set_admins(b.admins - {rid})
            out["msg"] = "Admin o'chirildi: %d" % rid
        return out

    # ---------- buyurtma (yetkazib beruvchilar) ----------
    @staticmethod
    def _cover(v):
        try:
            n = int(v)
        except (TypeError, ValueError):
            n = 30
        return min(max(n, 1), 365)

    def orders(self, cover):
        b = self.bot
        # standart: zaxira = sotuv davri → buyurtma = ehtiyoj (sotuv − qoldiq)
        cover = self._cover(cover) if str(cover or "").strip() else self._cover(max(1, b.data.period_days))
        sup = b.order_suppliers(cover)
        return {"cover": cover, "period": b.data.period_days, "when": B.dt_text(b.data.when()), "suppliers": sup}

    def order_excel(self, p):
        firm = (p.get("firm") or "").strip()
        items = [i for i in (p.get("items") or []) if B.to_float(i.get("qty")) > 0]
        if not firm or not items:
            return None
        safe = re.sub(r"[^\w\-]+", "_", firm, flags=re.U).strip("_")[:40] or "firma"
        return B.build_order(firm, items), "Buyurtma_%s_%s.xlsx" % (safe, datetime.date.today().isoformat())

    def order_save(self, p):
        firm = (p.get("firm") or "").strip()
        if not firm:
            return {"ok": False, "error": "Yetkazib beruvchi tanlanmagan."}
        oid, n, total = self.bot.save_order(firm, self._cover(p.get("cover")), p.get("items") or [])
        if not n:
            return {"ok": False, "error": "Buyurtmada tovar yo'q (hamma miqdorlar 0)."}
        return {"ok": True, "msg": "Saqlandi: %d ta tovar · %s so'm" % (n, B.money(total))}

    def order_contact(self, p):
        firm = (p.get("firm") or "").strip()
        user = (p.get("username") or "").strip().lstrip("@")
        if not firm:
            return {"ok": False, "error": "Yetkazib beruvchi tanlanmagan."}
        if user and not re.fullmatch(r"[A-Za-z0-9_]{4,32}", user):
            return {"ok": False, "error": "Username faqat lotin harf, raqam va _ dan iborat bo'ladi (masalan grandpharm_menejer)."}
        phone = (p.get("phone") or "").strip()
        if phone and len(B.digits(phone)) < 9:
            return {"ok": False, "error": "Telefon raqami to'liq emas."}
        self.bot.set_contact(firm, phone, user)
        return {"ok": True, "msg": "Kontakt saqlandi"}

    def order_blacklist(self, p):
        firm = (p.get("firm") or "").strip()
        on = bool(p.get("on"))
        reason = (p.get("reason") or "").strip()
        if not firm:
            return {"ok": False, "error": "Yetkazib beruvchi tanlanmagan."}
        if on and not reason:
            return {"ok": False, "error": "Sababini yozing."}
        if p.get("item"):
            self.bot.set_item_blacklist(firm, p["item"], on, reason)
            return {"ok": True, "msg": "Tovar yashirildi" if on else "Tovar qaytarildi"}
        self.bot.set_blacklist(firm, on, reason)
        return {"ok": True, "msg": "Qora ro'yxatga qo'shildi" if on else "Qora ro'yxatdan olindi"}

    # ---------- ko'chirish nazorati (dashboard) ----------
    @staticmethod
    def _date(v):
        try:
            return datetime.date.fromisoformat(str(v)).isoformat()
        except (TypeError, ValueError):
            return B.now().date().isoformat()

    def koch_today(self, date=None):
        b = self.bot
        s = b.koch_summary(self._date(date) if date else None)
        s["now"] = B.now().strftime("%H:%M")
        return s

    def koch_tasks(self, q):
        cat = q.get("cat") or {"zudlik": "shoshilinch"}.get(q.get("turi"), q.get("turi")) or None
        return self.bot.koch_task_list(self._date(q.get("date")), q.get("filial") or None, cat if cat in B.KOCH_CAT else None,
                                       q.get("status") or None, q.get("q") or None)

    def koch_task(self, tid):
        try:
            d = self.bot.koch_task_detail(int(tid))
        except (TypeError, ValueError):
            d = None
        return d or {"ok": False, "error": "Topshiriq topilmadi"}

    def koch_history(self, days):
        try:
            days = 30 if int(days) >= 30 else 7
        except (TypeError, ValueError):
            days = 7
        return self.bot.koch_history(days)

    def koch_excel(self, q):
        d = self._date(q.get("date"))
        rows = self.koch_tasks(q)
        if not rows:
            return None
        return B.build_koch_nazorat(self.bot, rows, d), "Kochirish_nazorat_%s.xlsx" % d

    def koch_remind(self, p):
        b = self.bot
        f = p.get("filial") or "all"
        if isinstance(p.get("filials"), list):  # «Кўчириш · ABC»: bajarilmagan topshiriqlari bor beruvchi filiallar
            targets = sorted({B.norm(x) for x in p["filials"] if str(x).strip()})
            if not targets:
                return {"ok": False, "error": "Eslatma uchun filial yo'q."}
            n = sum(b.remind_filial(x, "🔔 <b>Admin eslatmasi: ko'chirish</b>") for x in targets)
            return {"ok": True, "msg": "Eslatma yuborildi: %d filial, %d xodim" % (len(targets), n)}
        late = [r["filial"] for r in b.koch_summary()["filials"] if r["late"]]
        targets = late if f == "all" else [B.norm(f)]
        n = sum(b.remind_filial(x, "🔔 <b>Admin eslatmasi: ko'chirish</b>") for x in targets)
        return {"ok": True, "msg": "Eslatma yuborildi: %d filial, %d xodim" % (len(targets), n)}

    def koch_regen(self, p):
        b = self.bot
        b.data.reload(force=True)
        if b.data.error:
            return {"ok": False, "error": b.data.error}
        try:
            cnt = b.koch_regen_manual(by=(None, "admin (panel)"))
        except RuntimeError as e:
            return {"ok": False, "error": str(e)}
        return {"ok": True, "msg": "Qayta shakllantirildi: yangi %d · yangilangan %d · bekor %d · saqlangan %d" % (
            cnt["new"], cnt["updated"], cnt["cancelled"], cnt["kept"])}

    def koch_set_status(self, p):
        try:
            tid = int(p.get("id"))
        except (TypeError, ValueError):
            return {"ok": False, "error": "Noto'g'ri so'rov"}
        err = self.bot.koch_set_status(tid, p.get("status"), p.get("note") or "", (None, "admin (panel)"))
        return {"ok": False, "error": err} if err else {"ok": True, "msg": "Holat o'zgartirildi"}

    # ---------- ko'chirish · ABC ----------
    def koch_abc(self, period="bugun"):
        """GET /api/koch/abc?period=bugun|kecha|7 — ko'chirish topshiriqlari pul bo'yicha A / B / C.

        Bitta kun uchun natija (bot.koch_abc_summary(sana) yoki shu moduldagi _koch_abc_day(sana) — shakli bir xil):
          classes:  {"A"|"B"|"C": {jami, bajarildi, bajarilmadi, sum_jami, sum_bajarildi, sum_bajarilmadi}}
          filials:  {beruvchi filial: {"A": [bajarildi, jami], "B": [...], "C": [...]}}
          done:     [item, ...]  — bajarildi = berildi | olindi
          not_done: [item, ...]  — bajarilmadi = yangi | yopildi | rad (bekor — hisobga kirmaydi)
          item = {cls, name, src, dst, qty, sum, status, time} (+ ixtiyoriy: id, sana, producer, reason, late_min)
        «7 kun» — har kun alohida tasniflanadi (o'sha kunning tarmoq bo'yicha jami pulidan), keyin yig'iladi.
        Javobga qo'shiladi: period, dates, has, jami, bajarildi, sum_jami, sum_bajarildi, classes[*].share (pul ulushi, %)."""
        period = period if period in ABC_PERIODS else "bugun"
        back, n = ABC_PERIODS[period]
        end = B.now().date() - datetime.timedelta(days=back)
        dates = [(end - datetime.timedelta(days=i)).isoformat() for i in range(n - 1, -1, -1)]
        fn = getattr(self.bot, "koch_abc_summary", None)
        days = []
        for d in dates:
            s = self._abc_norm(fn(sana=d), d) if callable(fn) else self._koch_abc_day(d)
            for it in (s.get("done") or []) + (s.get("not_done") or []):
                it.setdefault("sana", d)
            days.append(s)
        out = self._abc_merge(days)
        out.update(period=period, dates=dates, today=B.now().date().isoformat())
        return out

    def _abc_norm(self, s, sana):
        """bot.koch_abc_summary natijasini panel shakliga keltiradi. Bot shakli: filials — ro'yxat, matrix —
        {filial: {"A": {jami, bajarildi, ...}, ..., "total": {...}}}, qatorda given_at / received_at / qty_given.
        Panel shaklidagi natija (filials — dict) o'zgarishsiz o'tadi; yetishmagan time / late_min qo'shiladi."""
        s = dict(s or {})
        if not isinstance(s.get("filials"), dict):
            mx = s.get("matrix") or {}
            s["filials"] = {f: {c: [int((v.get(c) or {}).get("bajarildi") or 0), int((v.get(c) or {}).get("jami") or 0)]
                                for c in ABC.CLASSES} for f, v in mx.items()}
        b, n, rows = self.bot, B.now(), None
        info = getattr(b, "task_info", None)
        for it in (s.get("done") or []) + (s.get("not_done") or []):
            st = it.get("status")
            if "time" not in it:
                tm = it.get("received_at") if st == "olindi" else it.get("given_at") if st in ("berildi", "rad") else ""
                it["time"] = (tm or "")[-5:] if tm else ""
            if st in KOCH_DONE and it.get("qty_given") is not None:
                it["qty"] = it["qty_given"]  # pul (sum) — reja bo'yicha, ko'rsatiladigan son — berilgani
            if st == "yangi" and "late_min" not in it and callable(info) and it.get("id") is not None:
                if rows is None:
                    rows = {t["id"]: t for t in b.koch_tasks(sana)}
                t = rows.get(it["id"])
                try:
                    i = info(t, n) if t is not None else {}
                    if i.get("late"):
                        it["late_min"] = int(i.get("late_min") or 0)
                except Exception:  # noqa — kechikish faqat qo'shimcha belgi
                    pass
        return s

    def _koch_abc_day(self, sana):
        """bot.koch_abc_summary bo'lmaguncha: bot.koch_tasks(sana) + abc_tahlil (pul = reja soni × kirim narxi)."""
        b = self.bot
        tasks = [t for t in b.koch_tasks(sana) if t["status"] != "bekor"]
        keys = set(tasks[0].keys()) if tasks else set()
        col = lambda t, k: t[k] if k in keys else None
        price = lambda t: float(t["prih"] or 0)
        plan_val = lambda t: float(t["qty_plan"] or 0) * price(t) or float(col(t, "summa") or 0)
        cls_of = ABC.classify(tasks, plan_val, lambda t: t["id"])
        info = getattr(b, "task_info", None)
        n = B.now()
        classes = {c: dict.fromkeys(("jami", "bajarildi", "bajarilmadi", "sum_jami", "sum_bajarildi", "sum_bajarilmadi"), 0)
                   for c in ABC.CLASSES}
        filials, done, not_done = {}, [], []
        for t in tasks:
            cls, ok = cls_of[t["id"]], t["status"] in KOCH_DONE
            plan = float(t["qty_plan"] or 0)
            qty = float(t["qty_given"]) if ok and t["qty_given"] is not None else plan
            s = plan_val(t)  # pul — reja bo'yicha (bot.koch_abc_summary bilan bir xil), ko'rsatiladigan son — berilgani
            tm = (t["received_at"] if t["status"] == "olindi" else t["given_at"]) or ""
            it = {"id": t["id"], "sana": sana, "cls": cls, "name": t["name"], "producer": col(t, "producer") or "",
                  "src": t["src"], "dst": t["dst"], "qty": qty, "sum": round(s, 2), "status": t["status"],
                  "time": tm[11:16] if t["status"] in ("berildi", "olindi", "rad") else "",
                  "reason": (col(t, "reason") or "") if t["status"] == "rad" else ""}
            if t["status"] == "yangi" and callable(info):
                try:
                    i = info(t, n)
                    if i.get("late"):
                        it["late_min"] = int(i.get("late_min") or 0)
                except Exception:  # noqa — kechikish faqat qo'shimcha belgi
                    pass
            c = classes[cls]
            c["jami"] += 1
            c["sum_jami"] += it["sum"]
            c["bajarildi" if ok else "bajarilmadi"] += 1
            c["sum_bajarildi" if ok else "sum_bajarilmadi"] += it["sum"]
            cell = filials.setdefault(t["src"], {k: [0, 0] for k in ABC.CLASSES})[cls]
            cell[1] += 1
            cell[0] += 1 if ok else 0
            (done if ok else not_done).append(it)
        return {"sana": sana, "classes": classes, "filials": filials, "done": done, "not_done": not_done}

    @staticmethod
    def _abc_merge(days):
        """Bir nechta kunlik natijani bittaga yig'adi (bitta kun bo'lsa ham shu yerdan o'tadi: jami va ulush qo'shiladi)."""
        nums = ("jami", "bajarildi", "bajarilmadi", "sum_jami", "sum_bajarildi", "sum_bajarilmadi")
        classes = {c: dict.fromkeys(nums, 0) for c in ABC.CLASSES}
        filials, done, not_done = {}, [], []
        for s in days:
            for c in ABC.CLASSES:
                src = (s.get("classes") or {}).get(c) or {}
                for k in nums:
                    classes[c][k] += src.get(k) or 0
            for f, v in (s.get("filials") or {}).items():
                dst = filials.setdefault(f, {k: [0, 0] for k in ABC.CLASSES})
                for c in ABC.CLASSES:
                    a = (v or {}).get(c) or [0, 0]
                    dst[c][0] += a[0]
                    dst[c][1] += a[1]
            done += s.get("done") or []
            not_done += s.get("not_done") or []
        order = {c: i for i, c in enumerate(ABC.CLASSES)}
        key = lambda it: (order.get(it.get("cls"), 9), -(it.get("sum") or 0), it.get("name") or "")
        total = sum(classes[c]["sum_jami"] for c in ABC.CLASSES)
        for c in ABC.CLASSES:
            classes[c]["share"] = int(round(100.0 * classes[c]["sum_jami"] / total)) if total else 0
            for k in ("sum_jami", "sum_bajarildi", "sum_bajarilmadi"):
                classes[c][k] = round(classes[c][k], 2)
        jami = sum(classes[c]["jami"] for c in ABC.CLASSES)
        return {"has": jami > 0, "classes": classes, "filials": filials, "done": sorted(done, key=key),
                "not_done": sorted(not_done, key=key), "jami": jami,
                "bajarildi": sum(classes[c]["bajarildi"] for c in ABC.CLASSES), "sum_jami": round(total, 2),
                "sum_bajarildi": round(sum(classes[c]["sum_bajarildi"] for c in ABC.CLASSES), 2)}

    # ---- Bosh sahifa: Srok / 90+ kun · A / B / C (6-talab, panel tomoni; bot tomoni — filial xodimi fil_abc) ----
    STOCK_ABC_TOP = 20
    OLD_DAYS = 90  # bosh sahifadagi «N+ кун турган» chegarasi (config.ini [panel] old_days)

    def old_days(self):
        try:
            n = int(float(self.cfg.get("panel", "old_days", fallback=str(self.OLD_DAYS))))
        except (ValueError, TypeError):
            n = self.OLD_DAYS
        return max(1, n)

    @staticmethod
    def _fil_name(f):
        return re.sub(r"\s+", " ", str(f or "")).strip()

    def dash_rows(self, kind):
        """Bosh sahifa (butun tarmoq): har dori BITTA qator, qoldiq — barcha filiallar va skladdagi jami.
        's' — Srok: muddatiga srok_days kundan kam qolgan partiyasi bor dorilar; qoldiq/pul — Srok varag'idagi BARCHA
             partiyalar (filiallar + sklad) yig'indisi, sana — eng yaqini.
        'o' — N+ kun: biror filialda N kundan ortiq turgan dorilar (N — old_days, odatda 90); kun — eng uzog'i,
             qoldiq — Sale varag'idagi jami (30 kundan kam turgan filiallardagisi ham), narx — 30+ kun varag'idan."""
        b, d = self.bot, self.bot.data
        n_old = self.old_days()
        ck = (kind, id(d.srok), id(d.old), id(d.sale), len(d.srok), len(d.old), len(d.sale), d.mtime, b.srok_days, n_old,
              datetime.date.today())
        cache = getattr(self, "_dash_cache", None)
        if cache is None:
            cache = self._dash_cache = {}
        if cache.get(kind, (None,))[0] == ck:
            return cache[kind][1]
        key = lambda r: (B.norm(r["name"]), B.norm(r["producer"]))
        out = []
        if kind == "s":
            tot = {}
            for r in d.srok:
                t = tot.setdefault(key(r), {"ost": 0.0, "pul": 0.0, "fil": {}})
                t["ost"] += r["ost"] or 0
                t["pul"] += (r["ost"] or 0) * (r["prih"] or 0)
                if (r["ost"] or 0) > 0 and B.fil_key(r["filial"]):
                    t["fil"].setdefault(B.fil_key(r["filial"]), self._fil_name(r["filial"]))
            g = {}
            for r in d.srok_rows(None, b.srok_days):
                k = key(r)
                if k not in g or (r["kun"], r["date"] or datetime.date.max) < (g[k]["kun"], g[k]["date"] or datetime.date.max):
                    g[k] = r
            for k, r in g.items():
                t = tot[k]
                fils = [t["fil"][x] for x in sorted(t["fil"])]
                out.append(dict(r, ost=t["ost"], prih=(t["pul"] / t["ost"]) if t["ost"] else (r["prih"] or 0),
                                summa=t["pul"], filials=fils, filial=self._fil_short(fils)))
            out.sort(key=lambda r: (r["kun"], B.norm(r["name"])))
        else:
            # Sale sarlavhasidagi nom («СКЛАД») o'rniga Srok / 30+ kun varag'idagi yozilishi («ДОРИХОНА СКЛАД»)
            known = {}
            for r in d.srok:
                known.setdefault(B.fil_key(r["filial"]), self._fil_name(r["filial"]))
            heads = [known.get(B.fil_key(h), h) if h else h for h in (d.sale_heads or [])]
            sale = {}
            for r in d.sale:
                sale.setdefault(key(r), r)
            g = {}
            for r in d.old:
                t = g.setdefault(key(r), {"r": r, "kun": 0, "ost": 0.0, "summa": 0.0, "fil": {}})
                t["kun"] = max(t["kun"], r["kun"] or 0)
                t["ost"] += r["ost"] or 0
                t["summa"] += r["summa"] or ((r["ost"] or 0) * (r["prih"] or 0))
                if B.fil_key(r["filial"]):
                    t["fil"].setdefault(B.fil_key(r["filial"]), self._fil_name(r["filial"]))
            for k, t in g.items():
                if t["kun"] < n_old:
                    continue
                s = sale.get(k)
                fil = dict(t["fil"])
                if s:
                    for i, col in enumerate((0, 2, 4, 6)):
                        if i < len(heads) and heads[i] and s["v"][col] > 0:
                            fil.setdefault(B.fil_key(heads[i]) or B.norm(heads[i]), self._fil_name(heads[i]))
                ost = max(s["stock"] if s else 0.0, t["ost"])
                narx = (t["summa"] / t["ost"]) if t["ost"] else (t["r"]["prih"] or 0)
                fils = [fil[x] for x in sorted(fil)]
                out.append(dict(t["r"], kun=t["kun"], ost=ost, prih=narx, summa=ost * narx, filials=fils,
                                filial=self._fil_short(fils)))
            out.sort(key=lambda r: (-r["kun"], B.norm(r["name"])))
        cache[kind] = (ck, out)
        return out

    @staticmethod
    def _fil_short(fils):
        """Jamlangan qator ostidagi yozuv: 1–2 filial — nomlari, ko'p bo'lsa — «N филиалда»."""
        return " · ".join(fils) if len(fils) <= 2 else "%d филиалда" % len(fils)

    def _stock_abc(self, kind, filial=""):
        """(qatorlar, guruhlar, xulosa, pul funksiyasi, filiallar, tanlangan filial). kind: 's' — Srok, 'o' — N+ kun.
        filial bo'sh — butun tarmoq, dori bo'yicha jamlangan (dash_rows); aks holda shu filialning qatorlari
        (fil_key bo'yicha aniq, boshqa filial aralashmaydi)."""
        b, d = self.bot, self.bot.data
        d.reload()
        if kind == "s":
            rows = d.srok_rows(None, b.srok_days)
        else:
            n_old = self.old_days()
            rows = [r for r in d.old_rows(None) if (r["kun"] or 0) >= n_old]
        names = {}
        for r in rows:
            k = B.fil_key(r["filial"])
            if k:
                names.setdefault(k, self._fil_name(r["filial"]))
        fk = B.fil_key(filial) if filial else ""
        if fk and fk in names:
            rows = [r for r in rows if B.fil_key(r["filial"]) == fk]
        else:
            fk = ""
            rows = self.dash_rows(kind)
        val = b.abc_value(kind)
        groups = ABC.split(rows, val)
        return rows, groups, ABC.summary(groups, val), val, [names[k] for k in sorted(names)], (names.get(fk, "") if fk else "")

    def stock_abc(self, kind="s", filial=""):
        """GET /api/stock/abc?kind=s|o&filial= — {kind, days, filial, filials, jami, sum, classes{A,B,C: n, sum, share},
        rows{A,B,C: birinchi 20 ta (pul ↓)}}. Qolganlari — /api/stock/abc/excel."""
        kind = kind if kind in ("s", "o") else "s"
        rows, groups, sm, val, fils, fil = self._stock_abc(kind, filial)

        def item(r):
            dt = r.get("date")
            return {"name": r["name"], "producer": r.get("producer") or "", "filial": re.sub(r"\s+", " ", str(r["filial"] or "")).strip(),
                    "date": dt.strftime("%d.%m.%Y") if dt else "", "kun": r["kun"], "ost": r["ost"] or 0, "sum": round(val(r))}
        return {"kind": kind, "days": self.bot.srok_days, "old_days": self.old_days(), "jamlangan": not fil,
                "filial": fil, "filials": fils, "top": self.STOCK_ABC_TOP,
                "jami": len(rows), "sum": round(sum(val(r) for r in rows)),
                "classes": {k: {"n": v["n"], "sum": round(v["sum"]), "share": v["share"]} for k, v in sm.items()},
                "rows": {k: [item(r) for r in groups[k][:self.STOCK_ABC_TOP]] for k in groups}}

    def stock_abc_excel(self, kind="s", cls="", filial=""):
        """Excel: cls bo'sh — hamma sinflar (A, B, C tartibida), aks holda bitta sinfning hammasi. Ma'lumot bo'lmasa None."""
        kind = kind if kind in ("s", "o") else "s"
        rows, groups, sm, val, fils, fil = self._stock_abc(kind, filial)
        cls = cls if cls in groups else ""
        part = [(k, r) for k in ("A", "B", "C") if not cls or k == cls for r in groups[k]]
        if not part:
            return None
        # butun tarmoq: qator — dori (jamlangan), Filial ustunida dori bor barcha filiallar; filial tanlansa — o'sha filial
        fcol, qcol, pcol = (("Filial", "Qoldiq", "Kirim narxi") if fil else ("Filiallar", "Jami qoldiq", "O'rtacha kirim narxi"))
        fname = (lambda r: r["filial"]) if fil else (lambda r: ", ".join(r.get("filials") or []))
        if kind == "s":
            heads = ["ABC", fcol, "Dori", "Ishlab chiqaruvchi", "Srok" if fil else "Eng yaqin srok", "Qolgan kun", qcol, pcol, "Summa"]
            data = [[k, fname(r), r["name"], r["producer"], r.get("date"), r["kun"], r["ost"], r["prih"], val(r)] for k, r in part]
            widths, dcols, ncols = [6, 22 if fil else 40, 46, 26, 12, 10, 10, 12, 14], (4,), (5, 6, 7, 8)
        else:
            heads = ["ABC", fcol, "Dori", "Ishlab chiqaruvchi", "Turgan kun" if fil else "Eng uzoq turgan kun", qcol, pcol, "Summa"]
            data = [[k, fname(r), r["name"], r["producer"], r["kun"], r["ost"], r["prih"], val(r)] for k, r in part]
            widths, dcols, ncols = [6, 22 if fil else 40, 46, 26, 10, 10, 12, 14], (), (4, 5, 6, 7)
        name = "Srok" if kind == "s" else "%d+ kun" % self.old_days()
        content = B.xlsx_bytes("%s %s" % (name, cls or "ABC"), heads, data, widths, datecols=dcols, numcols=ncols)
        short = re.sub(r"[^\w]+", "_", B.fil_key(fil), flags=re.U).strip("_")[:20] if fil else "tarmoq"
        return content, "%s_%s_%s_%s.xlsx" % ("Srok" if kind == "s" else "%dkun" % self.old_days(), cls or "ABC", short or "filial",
                                               datetime.date.today().isoformat())

    def koch_abc_excel(self, period):
        s = self.koch_abc(period)
        if not s["has"]:
            return None
        st = {"yangi": "янги", "berildi": "берилди", "olindi": "олинди", "rad": "рад", "yopildi": "ёпилди"}
        rows = [[it.get("sana", ""), it["cls"], "Бажарилди" if ok else "Бажарилмади", it["name"], it.get("producer", ""), it["src"],
                 it["dst"], it["qty"], it["sum"], st.get(it["status"], it["status"]), it.get("time") or "", it.get("reason") or ""]
                for ok, lst in ((True, s["done"]), (False, s["not_done"])) for it in lst]
        content = B.xlsx_bytes("Ko'chirish ABC", ["Sana", "Guruh", "Natija", "Dori", "Ishlab chiqaruvchi", "Beruvchi", "Oluvchi",
                                                  "Soni", "Summa", "Holat", "Vaqt", "Sabab"], rows,
                               [12, 8, 14, 44, 22, 22, 22, 9, 14, 11, 8, 26], numcols=(7, 8))
        name = s["dates"][0] if len(s["dates"]) == 1 else "%s_%s" % (s["dates"][0], s["dates"][-1])
        return content, "Kochirish_ABC_%s.xlsx" % name

    # ---------- buyurtma so'rovlari (firma vakili → admin, BUYURTMA-SOROV.md) ----------
    @staticmethod
    def _ts(v):
        return (v or "")[:16].replace("T", " ")

    def _oreq(self, r, items):
        asked, ok, changed = BY.totals(items)
        return {"id": r["id"], "firm": r["firm"] or "", "name": r["name"] or "", "chat_id": r["chat_id"],
                "created_at": self._ts(r["created_at"]), "status": r["status"], "resolved_name": r["resolved_name"] or "",
                "resolved_at": self._ts(r["resolved_at"]), "n": len(items), "asked": asked, "ok": ok, "changed": changed,
                "rejected": sum(1 for i in items if i["status"] == BY.I_REJ),
                "decided": sum(1 for i in items if i["status"] != BY.I_WAIT)}

    def oreq_list(self, status=""):
        """GET /api/orders?status=kutmoqda|hal|all — so'rovlar ro'yxati (yangilari tepada) va filtr sonlari."""
        b = self.bot
        with b.dblock:
            if status == "kutmoqda":
                rows = b.db.execute("SELECT * FROM order_req WHERE status=? ORDER BY id DESC", (BY.OPEN,)).fetchall()
            elif status == "hal":
                rows = b.db.execute("SELECT * FROM order_req WHERE status<>? ORDER BY id DESC LIMIT 300", (BY.OPEN,)).fetchall()
            else:
                rows = b.db.execute("SELECT * FROM order_req ORDER BY id DESC LIMIT 300").fetchall()
            out = [self._oreq(r, BY.get(b.db, r["id"])[1]) for r in rows]
            n_open = b.db.execute("SELECT COUNT(*) FROM order_req WHERE status=?", (BY.OPEN,)).fetchone()[0]
            n_all = b.db.execute("SELECT COUNT(*) FROM order_req").fetchone()[0]
        return {"items": out, "counts": {"kutmoqda": n_open, "hal": n_all - n_open, "all": n_all}}

    def oreq_get(self, rid):
        b = self.bot
        try:
            rid = int(rid)
        except (TypeError, ValueError):
            return {"ok": False, "error": "So'rov topilmadi"}
        with b.dblock:
            r, items = BY.get(b.db, rid)
        if not r:
            return {"ok": False, "error": "So'rov topilmadi"}
        return {"ok": True, "req": self._oreq(r, items), "all_decided": BY.all_decided(items),
                "items": [{"pos": i["pos"], "name": i["name"], "producer": i["producer"] or "", "qty_asked": i["qty_asked"],
                           "qty_ok": i["qty_ok"], "status": i["status"]} for i in items]}

    @staticmethod
    def _rid_pos(p, need_pos=False):
        try:
            return int(p.get("id")), (int(p.get("pos")) if need_pos else None)
        except (TypeError, ValueError):
            return None, None

    def oreq_decide(self, p):
        """POST /api/orders/decide {id, pos, qty | reject}: bitta dori bo'yicha qaror (qty yo'q — so'ralgancha)."""
        rid, pos = self._rid_pos(p, True)
        if rid is None:
            return {"ok": False, "error": "Noto'g'ri so'rov"}
        qty = None
        if not p.get("reject") and p.get("qty") not in (None, ""):
            v = p["qty"]  # faqat butun son 0..ORD_MAX (NaN / inf / kasr / bool emas)
            if isinstance(v, float) and v == v and v not in (float("inf"), float("-inf")) and v.is_integer():
                v = int(v)
            qty = v if isinstance(v, int) and not isinstance(v, bool) else B.Bot.ord_int(v) if isinstance(v, str) else None
            if qty is None or not 0 <= qty <= B.Bot.ORD_MAX:
                return {"ok": False, "error": "Butun son kiriting: 0 dan %d gacha." % B.Bot.ORD_MAX}
        with self.bot.dblock:
            ok = BY.decide_item(self.bot.db, rid, pos, qty, bool(p.get("reject")))
        if not ok:
            return {"ok": False, "error": "So'rov allaqachon hal qilingan yoki dori topilmadi."}
        return dict(self.oreq_get(rid), msg="Saqlandi")

    def oreq_accept_all(self, p):
        rid, _ = self._rid_pos(p)
        if rid is None:
            return {"ok": False, "error": "Noto'g'ri so'rov"}
        with self.bot.dblock:
            ok = BY.accept_all(self.bot.db, rid)
        if not ok:
            return {"ok": False, "error": "So'rov allaqachon hal qilingan."}
        return dict(self.oreq_get(rid), msg="Hammasi so'ralgancha qabul qilindi")

    def oreq_finish(self, p):
        """POST /api/orders/finish {id, reject?}: javobni yuborish (yoki butun so'rovni rad etish)."""
        b = self.bot
        rid, _ = self._rid_pos(p)
        if rid is None:
            return {"ok": False, "error": "Noto'g'ri so'rov"}
        reject = bool(p.get("reject"))
        with b.dblock:
            r, items = BY.get(b.db, rid)
            if not r:
                return {"ok": False, "error": "So'rov topilmadi"}
            if r["status"] != BY.OPEN:
                return {"ok": False, "error": "Bu so'rovga allaqachon javob berilgan (%s)." % (r["resolved_name"] or "boshqa admin")}
            if not reject and not BY.all_decided(items):
                return {"ok": False, "error": "Avval har bir dori bo'yicha qaror qiling (✅ Ha, son yoki ❌ olib tashlash)."}
            ok = BY.finish(b.db, rid, None, "Админ (панел)", reject)
        if not ok:
            return {"ok": False, "error": "Javobni saqlab bo'lmadi: so'rov o'zgargan. Sahifa yangilandi."}
        msg = "So'rov rad etildi" if reject else "Javob yuborildi"
        notify = getattr(b, "order_notify", None)
        if callable(notify):
            try:
                notify(rid)
            except Exception:  # noqa
                logging.exception("Panel: buyurtma so'rovi #%d javobini yuborib bo'lmadi", rid)
                msg += ", lekin vakilga Telegram xabarini yuborib bo'lmadi"
        else:
            # TODO(bot): bot.order_notify(rid) — firma vakiliga Telegram'da javob (BUYURTMA-SOROV.md, «Firma vakili» 4-band).
            # Bot muhandisi qo'shadi; ungacha javob faqat bazada saqlanadi va vakil «📋 Сўровларим» da ko'radi.
            msg += " (vakilga Telegram xabari hali ulanmagan)"
        return dict(self.oreq_get(rid), msg=msg)

    def oreq_excel(self, days=7, firm=None):
        """GET /api/orders/excel?days=7[&firm=] — tasdiqlangan qatorlar (.xlsx, sarlavhalar kirillda)."""
        try:
            days = min(max(int(days), 1), 366)
        except (TypeError, ValueError):
            days = 7
        since = (datetime.datetime.now() - datetime.timedelta(days=days)).isoformat(timespec="seconds")
        with self.bot.dblock:
            rows = BY.approved_rows(self.bot.db, since, firm or None)
        if not rows:
            return None
        data = [[r["id"], r["firm"] or "", r["who"] or "", self._ts(r["resolved_at"]), r["name"], r["producer"] or "",
                 r["qty_asked"], r["qty_ok"], OREQ_ITEM_LBL.get(r["status"], r["status"])] for r in rows]
        content = B.xlsx_bytes("Buyurtma so'rovlari", ["№", "Firma", "Vakil", "Javob vaqti", "Dori", "Ishlab chiqaruvchi",
                                                       "So'ralgan", "Tasdiqlangan", "Holat"], data,
                               [7, 26, 22, 17, 44, 22, 11, 13, 15], numcols=(6, 7))
        return content, "Buyurtma_sorovlari_%dkun_%s.xlsx" % (days, datetime.date.today().isoformat())

    # ---------- statistika, mijoz ko'rinishi ----------
    def stats(self, days):
        try:
            days = min(max(int(days), 1), 180)
        except (TypeError, ValueError):
            days = 7
        return self.bot.stats(days)

    def user_preview(self, uid):
        """Mijoz botda nimani ko'radi (oddiy matn ko'rinishida). Mijozga hech narsa yuborilmaydi."""
        b = self.bot
        c = b.client(int(uid))
        if not c:
            return {"ok": False, "error": "Mijoz topilmadi"}
        if c["role"] != "staff" and not b.can_use(c):
            return {"ok": False, "error": "Firma biriktirilmagan: mijoz hali hech narsa ko'rmaydi."}
        plain = lambda t: html.unescape(re.sub(r"<[^>]+>", "", UK.lat2cyr(t or "")))
        panel, kb = b.client_panel(c, save_sig=False)
        buttons = [[x["text"] for x in row] for row in kb["keyboard"]]
        if c["role"] == "filial":
            f = B.norm(c["filial"])
            sec = {k["title"]: plain(b.filial_view(f, k["key"])) for k in B.KOCH_CATS}
            sec.update({"Keladigan": plain(b.filial_view(f, "kin")), "Bugungi holat": plain(b.filial_view(f, "kday"))})
            return {"ok": True, "name": c["name"] or c["phone"], **UK.deep({"panel": plain(panel), "buttons": buttons, "sections": {UK.lat2cyr(k): v for k, v in sec.items()}})}
        tags, label = b.eff(c)
        sec = {}
        for k, name in (("srok", "Srok"), ("old", "30+ kun"), ("need", "Buyurtma"), ("stock", "Qoldiq")):
            sec[name] = plain(b.section_text(k, tags, label))
        acc = b.hisob_text(tags, None, label) if c["role"] == "staff" else b.hisob_text(tags, b.accounts_of(c["chat_id"]), auto=False)
        sec["Hisob"] = plain(acc) if acc else "Hisob biriktirilmagan."
        buttons = [[x["text"] for x in row] for row in kb["keyboard"]]
        return {"ok": True, "name": c["name"] or c["phone"], **UK.deep({"panel": plain(panel), "buttons": buttons, "sections": {UK.lat2cyr(k): v for k, v in sec.items()}})}

    def log_tail(self, n=200):
        path = os.path.join(B.BASE_DIR, "bot.log")
        try:
            with open(path, "rb") as f:
                f.seek(0, 2)
                size = f.tell()
                f.seek(max(0, size - 120_000))
                raw = f.read().decode("utf-8", "replace")
        except OSError:
            return []
        lines = [ln for ln in raw.splitlines() if re.match(r"^\d{4}-\d\d-\d\d ", ln)]
        return lines[-n:][::-1]


def make_handler(api, password):
    html_path = os.path.join(B.APP_DIR, "panel.html")

    class H(BaseHTTPRequestHandler):
        server_version = "AbuPanel"

        def log_message(self, fmt, *args):  # jim
            pass

        # ---- yordamchilar ----
        def _host_ok(self):
            host = (self.headers.get("Host") or "").rsplit(":", 1)[0].strip("[]").lower()
            return host in LOOPBACK or host == self.server.bind_host

        def _auth_ok(self):
            if not password:
                return True
            import base64
            h = self.headers.get("Authorization", "")
            if h.startswith("Basic "):
                try:
                    user, _, pw = base64.b64decode(h[6:]).decode("utf-8").partition(":")
                    return pw == password
                except Exception:  # noqa
                    return False
            return False

        def _send(self, code, body, ctype="application/json; charset=utf-8", extra=None):
            if isinstance(body, (dict, list)):
                # server xabarlari kirillda; ma'lumot belgilari (esc) olib tashlanadi
                body = json.dumps(UK.deep(body, keys=("msg", "error", "warn", "hint")), ensure_ascii=False)
                body = UK.strip_marks(body).encode("utf-8")
            elif isinstance(body, str):
                body = body.encode("utf-8")
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Content-Type-Options", "nosniff")
            for k, v in (extra or {}).items():
                self.send_header(k, v)
            self.end_headers()
            self.wfile.write(body)

        def _guard(self):
            if not self._host_ok():
                self._send(403, {"ok": False, "error": "Ruxsat yo'q"})
                return False
            if not self._auth_ok():
                self._send(401, "Parol kerak", "text/plain; charset=utf-8", {"WWW-Authenticate": 'Basic realm="TimeSaver"'})
                return False
            return True

        # ---- GET ----
        def do_GET(self):
            if not self._guard():
                return
            u = urllib.parse.urlparse(self.path)
            q = {k: v[0] for k, v in urllib.parse.parse_qs(u.query).items()}
            try:
                if u.path in ("/", "/index.html"):
                    with open(html_path, "rb") as f:
                        return self._send(200, f.read(), "text/html; charset=utf-8")
                if u.path in ("/logo.svg", "/favicon.ico", "/logo-128.png"):
                    name = {"/favicon.ico": "timesaver.ico"}.get(u.path, u.path[1:])
                    ctype = {"svg": "image/svg+xml", "ico": "image/x-icon", "png": "image/png"}[name.rsplit(".", 1)[1]]
                    try:
                        with open(os.path.join(B.APP_DIR, "branding", name), "rb") as f:
                            return self._send(200, f.read(), ctype)
                    except OSError:
                        return self._send(404, {"ok": False, "error": "topilmadi"})
                if u.path == "/api/state":
                    return self._send(200, api.state())
                if u.path == "/api/setup":
                    return self._send(200, api.setup_get())
                if u.path == "/api/users":
                    return self._send(200, api.users(q.get("status", ""), q.get("q", "")))
                if u.path == "/api/acc_requests":
                    return self._send(200, api.acc_requests())
                if u.path == "/api/tags":
                    return self._send(200, api.tags())
                if u.path == "/api/filials":
                    return self._send(200, api.filials())
                if u.path == "/api/accounts":
                    return self._send(200, api.accounts(q.get("q", "")))
                if u.path == "/api/reports":
                    return self._send(200, api.report_counts())
                if u.path == "/api/report":
                    kind = q.get("kind", "")
                    if kind not in api.bot.KINDS:
                        return self._send(404, {"ok": False, "error": "Noma'lum hisobot"})
                    res = api.report_file(kind, q.get("firm", ""))
                    if not res:
                        return self._send(404, {"ok": False, "error": "Bu bo'lim bo'yicha ma'lumot topilmadi"})
                    content, fname = res
                    return self._send(200, content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                      {"Content-Disposition": 'attachment; filename="%s"' % fname})
                if u.path == "/api/bc/audience":
                    return self._send(200, api.bc_audience_counts())
                if u.path == "/api/bc/preview":
                    return self._send(200, api.bc_preview(q.get("kind", "srok")))
                if u.path == "/api/settings":
                    return self._send(200, api.settings_view())
                if u.path == "/api/log":
                    return self._send(200, api.log_tail())
                if u.path == "/api/orders":
                    # ?status=… — firma vakillarining buyurtma so'rovlari; ?cover=… (eski) — yetkazib beruvchilar ro'yxati
                    if "status" in urllib.parse.parse_qs(u.query, keep_blank_values=True):
                        return self._send(200, api.oreq_list(q.get("status", "")))
                    return self._send(200, api.orders(q.get("cover", "")))
                if u.path == "/api/orders/excel":
                    res = api.oreq_excel(q.get("days", "7"), q.get("firm"))
                    if not res:
                        return self._send(404, {"ok": False, "error": "Bu davrda tasdiqlangan buyurtma so'rovi yo'q"})
                    return self._send(200, res[0], XLSX, {"Content-Disposition": "attachment; filename=\"%s\"" % res[1]})
                m = re.fullmatch(r"/api/orders/(\d+)", u.path)
                if m:
                    return self._send(200, api.oreq_get(m.group(1)))
                if u.path == "/api/orders/history":
                    return self._send(200, api.bot.orders_history(q.get("firm") or None))
                if u.path == "/api/stats":
                    return self._send(200, cyr_keys(api.stats(q.get("days", "7")), ("label",)))
                if u.path == "/api/koch/today":
                    return self._send(200, cyr_keys(api.koch_today(q.get("date")), ("title",)))
                if u.path == "/api/koch/abc":
                    s = api.koch_abc(q.get("period", "bugun"))
                    rr = set(getattr(api.bot, "RAD_REASONS", ()))  # tizim tanlovlari («Tovar topilmadi») — kirillda
                    for it in s["not_done"]:
                        if it.get("reason") in rr:
                            it["reason"] = UK.lat2cyr(it["reason"])
                    return self._send(200, s)
                if u.path == "/api/stock/abc":
                    return self._send(200, api.stock_abc(q.get("kind", "s"), q.get("filial", "")))
                if u.path == "/api/stock/abc/excel":
                    res = api.stock_abc_excel(q.get("kind", "s"), q.get("cls", ""), q.get("filial", ""))
                    if not res:
                        return self._send(404, {"ok": False, "error": "Bu bo'lim bo'yicha dori yo'q"})
                    return self._send(200, res[0], XLSX, {"Content-Disposition": "attachment; filename*=UTF-8''%s" % urllib.parse.quote(res[1])})
                if u.path == "/api/koch/abc/excel":
                    res = api.koch_abc_excel(q.get("period", "bugun"))
                    if not res:
                        return self._send(404, {"ok": False, "error": "Bu davrda ko'chirish topshirig'i yo'q"})
                    return self._send(200, res[0], XLSX, {"Content-Disposition": "attachment; filename=\"%s\"" % res[1]})
                if u.path == "/api/koch/tasks":
                    return self._send(200, api.koch_tasks(q))
                if u.path == "/api/koch/task":
                    return self._send(200, api.koch_task(q.get("id")))
                if u.path == "/api/koch/history":
                    return self._send(200, api.koch_history(q.get("days", "7")))
                if u.path == "/api/koch/excel":
                    res = api.koch_excel(q)
                    if not res:
                        return self._send(404, {"ok": False, "error": "Bu tanlov bo'yicha topshiriq yo'q"})
                    return self._send(200, res[0], "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                      {"Content-Disposition": "attachment; filename=\"%s\"" % res[1]})
                if u.path == "/api/sklad":
                    return self._send(200, api.bot.sklad_summary())
                if u.path == "/api/sklad/excel":
                    s = api.bot.sklad_summary()
                    if not s["items"] and not s["excluded"]:
                        return self._send(404, {"ok": False, "error": "Sklad bo'yicha taqsimot yo'q"})
                    return self._send(200, B.build_sklad(s), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                      {"Content-Disposition": "attachment; filename=\"Sklad_taqsimot_%s.xlsx\"" % datetime.date.today().isoformat()})
                if u.path == "/api/fom":
                    s = api.bot.fom_summary()
                    if isinstance(s, dict) and isinstance(s.get("cats"), list):  # toifa sarlavhalari (🔴 SHOSHILINCH …)
                        s["cats"] = [UK.lat2cyr(x) if isinstance(x, str) else x for x in s["cats"]]
                    return self._send(200, s)
                if u.path == "/api/kirim":
                    return self._send(200, api.bot.kirim_summary(q.get("sana", "")))
                if u.path == "/api/fom/file":
                    try:
                        res = api.bot.fom_file(int(q.get("id", "0")))
                    except ValueError:
                        res = None
                    if not res:
                        return self._send(404, {"ok": False, "error": "Bron topilmadi"})
                    return self._send(200, res[0], "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                      {"Content-Disposition": "attachment; filename*=UTF-8''%s" % urllib.parse.quote(res[1])})
                if u.path == "/api/user/preview":
                    return self._send(200, api.user_preview(q.get("id", "0")))
                return self._send(404, {"ok": False, "error": "topilmadi"})
            except Exception as e:  # noqa
                logging.exception("Panel GET xatosi: %s", self.path)
                return self._send(500, {"ok": False, "error": str(e)})

        # ---- POST ----
        def do_POST(self):
            if not self._guard():
                return
            if self.headers.get("X-Requested-With") != "panel":
                # so'rov tanasini o'qib tashlaymiz: o'qilmagan ma'lumot bilan yopilgan soket Windows'da RST beradi va
                # mijoz 403 o'rniga «ulanish uzildi» oladi
                try:
                    n = int(self.headers.get("Content-Length") or 0)
                    if 0 < n <= 1 << 20:
                        self.rfile.read(n)
                except Exception:  # noqa
                    pass
                return self._send(403, {"ok": False, "error": "Ruxsat yo'q"})
            if urllib.parse.urlparse(self.path).path == "/api/setup/upload":  # xom fayl, JSON emas
                try:
                    n = int(self.headers.get("Content-Length") or 0)
                    if not 0 < n <= MAX_UPLOAD:
                        return self._send(200, {"ok": False, "error": "Fayl hajmi 300 MB dan oshmasligi kerak."})
                    body = self.rfile.read(n)
                    return self._send(200, api.setup_upload(urllib.parse.unquote(self.headers.get("X-File-Name", "")), body))
                except Exception as e:  # noqa
                    logging.exception("Excel yuklashda xato")
                    return self._send(200, {"ok": False, "error": str(e)})
            try:
                n = int(self.headers.get("Content-Length") or 0)
                p = json.loads(self.rfile.read(n).decode("utf-8") or "{}") if n else {}
            except Exception:  # noqa
                return self._send(400, {"ok": False, "error": "Noto'g'ri so'rov"})
            path = urllib.parse.urlparse(self.path).path
            try:
                if path == "/api/setup":
                    return self._send(200, api.setup_set(p))
                if path == "/api/setup/token":
                    return self._send(200, api.setup_token(p))
                if path == "/api/setup/sql_detect":
                    return self._send(200, api.setup_sql_detect(p))
                if path == "/api/setup/sql_dbs":
                    return self._send(200, api.setup_sql_dbs(p))
                if path == "/api/refresh":
                    if api.bot.start_refresh(None):
                        return self._send(200, {"ok": True})
                    return self._send(200, {"ok": False, "error": "Yangilanish allaqachon ketmoqda."})
                if path == "/api/user/save":
                    return self._send(200, api.user_save(p))
                if path == "/api/user/action":
                    return self._send(200, api.user_action(p))
                if path == "/api/user/add":
                    return self._send(200, api.user_add(p))
                if path == "/api/acc_request":
                    return self._send(200, api.acc_request_action(p))
                if path == "/api/send_one":
                    return self._send(200, api.send_one(p))
                if path == "/api/bc/send":
                    return self._send(200, api.bc_send(p))
                if path == "/api/settings":
                    return self._send(200, api.settings_set(p))
                if path == "/api/order/excel":
                    res = api.order_excel(p)
                    if not res:
                        return self._send(200, {"ok": False, "error": "Buyurtmada tovar yo'q."})
                    content, fname = res
                    return self._send(200, content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                      {"Content-Disposition": "attachment; filename*=UTF-8''%s" % urllib.parse.quote(fname)})
                if path == "/api/order/save":
                    return self._send(200, api.order_save(p))
                if path == "/api/order/contact":
                    return self._send(200, api.order_contact(p))
                if path == "/api/order/blacklist":
                    return self._send(200, api.order_blacklist(p))
                if path == "/api/koch/remind":
                    return self._send(200, api.koch_remind(p))
                if path == "/api/koch/regen":
                    return self._send(200, api.koch_regen(p))
                if path == "/api/koch/set_status":
                    return self._send(200, api.koch_set_status(p))
                if path == "/api/orders/decide":
                    return self._send(200, api.oreq_decide(p))
                if path == "/api/orders/accept_all":
                    return self._send(200, api.oreq_accept_all(p))
                if path == "/api/orders/finish":
                    return self._send(200, api.oreq_finish(p))
                if path == "/api/fom/cancel":
                    try:
                        err = api.bot.fom_cancel(int(p.get("id")))
                    except (TypeError, ValueError):
                        err = "Noto'g'ri so'rov"
                    return self._send(200, {"ok": False, "error": err} if err else {"ok": True, "msg": "Bron bekor qilindi, dorilar ro'yxatga qaytdi"})
                if path == "/api/kirim/create":
                    made = api.bot.kirim_create(str(p.get("sana", "")), p.get("alloc") or {}, str(p.get("sig", "")))
                    return self._send(200, {"ok": True, "ids": [b["id"] for b in made],
                                            "msg": "FOM bron yaratildi: " + ", ".join("#%d → %s (FOM #%s)" % (b["id"], b["dst"], b["doc_no"]) for b in made)})
                if path == "/api/fom/check":
                    return self._send(200, api.bot.fom_check(str(p.get("which", "")), str(p.get("sig", ""))))
                if path == "/api/fom/create":
                    made = api.bot.fom_create(str(p.get("which", "")), str(p.get("sig", "")))
                    if not made:
                        return self._send(200, {"ok": False, "error": "Ro'yxat o'zgargan yoki bron allaqachon yaratilgan. Sahifa yangilandi."})
                    return self._send(200, {"ok": True, "ids": [b["id"] for b in made],
                                            "msg": "Bron yaratildi: " + ", ".join("#%d → %s" % (b["id"], b["dst"]) for b in made)})
                return self._send(404, {"ok": False, "error": "topilmadi"})
            except Exception as e:  # noqa
                logging.exception("Panel POST xatosi: %s", path)
                return self._send(200, {"ok": False, "error": str(e)})

    return H


class PanelServer(ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = True
    bind_host = "127.0.0.1"


def start(bot, cfg):
    sec = cfg["panel"] if cfg.has_section("panel") else {}
    if str(sec.get("enabled", "yes")).lower() not in ("yes", "1", "true", "ha"):
        logging.info("Server paneli o'chirilgan (config.ini [panel] enabled=no)")
        return None
    host = str(sec.get("host", "127.0.0.1")).strip() or "127.0.0.1"
    port = int(sec.get("port", "8080"))
    password = str(sec.get("password", "")).strip()
    if host not in LOOPBACK and not password:
        logging.warning("Panel: tarmoqqa ochish uchun [panel] password kerak. Xavfsizlik uchun 127.0.0.1 ishlatiladi.")
        host = "127.0.0.1"
    api = Api(bot, cfg)
    srv, err = None, None
    for pt in range(port, port + 11):
        try:
            srv = PanelServer((host, pt), make_handler(api, password))
            port = pt
            break
        except OSError as e:
            err = e
    if srv is None:
        raise OSError("Panel uchun bo'sh port topilmadi (%d..%d): %s" % (port, port + 10, err))
    srv.bind_host = host.lower()
    bot.panel_port = port
    try:
        with open(os.path.join(B.BASE_DIR, "panel_port.txt"), "w") as f:
            f.write(str(port))
    except OSError:
        pass
    threading.Thread(target=srv.serve_forever, daemon=True, name="panel").start()
    logging.info("Server paneli: http://%s:%d", "127.0.0.1" if host in LOOPBACK else host, port)
    return srv
