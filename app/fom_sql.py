# -*- coding: utf-8 -*-
"""FOM (SQL Server 2005, FOMSKLADSKLAD) ga QORALAMA bron yozish: перемещение (INVOICE.DOCTYPE 15, STATE 2).

FOM'ning o'zi qoralama qanday yozsa, shunday yozadi (FOM yaratgan qoralamalar 48125/48126/48132 bo'yicha):
  INVOICE    — DOCTYPE 15, NUMBER 0, ORG = filial, STATE 2, SMENA = foydalanuvchining ochiq smenasi
  INVOICELN  — partiya (INCOMELN) bo'yicha: KOL, PRICE = PRICE1 = PRICEROZ1, NDS, STATE 2
Qoldiq (INCOMELN.SALE) va hujjat summasini FOM triggerlari o'zi hisoblaydi; skladda yetmasa, inventarizatsiya
yoki yopilgan davr bo'lsa — FOM triggeri rad etadi va hammasi bekor bo'ladi (bitta tranzaksiya).
Raqam berish va filialga jo'natish — FOM'da odam «o'tkazganda».

Himoya: baza nomi «_TEST» bilan tugamasa, config.ini [fom] live = yes bo'lmaguncha hech narsa yozilmaydi.
"""
import datetime
import logging
import re

DOCTYPE_MOVE = 15      # ПЕРЕМЕЩЕНИЕ В АПТЕКУ
STATE_DRAFT = 2        # qoralama (sarlavha va qatorlar)
INC_POSTED = 3         # o'tkazilgan kirim partiyasi
DB_INI = r"C:\ProgramData\BazaReport\db.ini"
QUERY_TIMEOUT = 60     # soniya: FOM so'rovi / qulf kutish chegarasi


def odbc_quote(v):
    """ODBC ulanish satri qiymati: ; yoki } bo'lsa ham to'g'ri o'qilsin (PWD={ab;c}})."""
    return "{%s}" % str(v).replace("}", "}}")
DEFAULT_ORG = {"ДОРИХОНА ЧИНОЗ": 2, "ДОРИХОНА УРГУТ": 7, "ДОРИХОНА ЭРКИН": 9}


def key(s):
    return re.sub(r"\s+", " ", str(s if s is not None else "")).strip().upper()


def num(v):
    return float(v or 0)


class FomError(Exception):
    pass


class TSql:
    """SQL Server: yangi hujjat ID si o'sha batch ichida SCOPE_IDENTITY() bilan (triggerlar boshqa jadvalga yozadi)."""

    @staticmethod
    def computed(cur, table):
        """Jadvalning hisoblanadigan ustunlari (ularga yozib bo'lmaydi; FOM o'zi hisoblaydi)."""
        cur.execute("SELECT name FROM sys.computed_columns WHERE object_id = OBJECT_ID(?)", ("dbo." + table,))
        return {str(r[0]).upper() for r in cur.fetchall()}

    @staticmethod
    def insert_id(cur, sql, params):
        cur.execute("SET NOCOUNT ON; " + sql + "; SELECT CAST(SCOPE_IDENTITY() AS int)", params)
        while cur.description is None:
            if not cur.nextset():
                raise FomError("Yangi hujjat raqami olinmadi")
        row = cur.fetchone()
        if not row or row[0] is None:
            raise FomError("Yangi hujjat raqami olinmadi")
        return int(row[0])


def read_text_ansi(path):
    """db.ini matni. Panel uni Windows ANSI (mbcs) da yozadi — makros (VBA Line Input) shunday o'qiydi; eski fayllar
    UTF-8 bo'lishi mumkin. Avval UTF-8 (qat'iy), bo'lmasa ANSI: kirill/lotin bo'lmagan parol buzilmasin."""
    with open(path, "rb") as f:
        raw = f.read()
    try:
        return raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        try:
            return raw.decode("mbcs")
        except LookupError:  # Windows emas
            return raw.decode("cp1251", errors="replace")


def read_flat_ini(path):
    """db.ini: kalit=qiymat (bo'limsiz), Excel makrosi bilan bir xil o'qiladi."""
    out = {}
    try:
        text = read_text_ansi(path)
    except OSError:
        return out
    for line in text.splitlines():
        if "=" in line:
            k, v = line.split("=", 1)
            out.setdefault(k.strip().lower(), v.strip())
    return out


class FomSql:
    def __init__(self, cfg, connect=None, dialect=TSql):
        sec = cfg["fom"] if cfg.has_section("fom") else {}
        g = lambda k, d="": str(sec.get(k, d) or d).strip()  # noqa: E731
        self.mode = g("mode", "fayl").lower()
        self.database = g("database", "FOMSKLADSKLAD_TEST")
        self.live = g("live", "no").lower() in ("yes", "ha", "1", "true")
        try:
            self.fom_user = int(g("fom_user", "1"))
        except ValueError:
            self.fom_user = 1
        self.server, self.uid, self.pwd = g("server"), g("user"), g("password")
        self.ini_path = g("db_ini", DB_INI)
        try:   # sinov chegarasi: bitta bronga ko'pi bilan N ta dori (0 = cheklovsiz)
            self.limit = max(0, int(g("limit", "0")))
        except ValueError:
            self.limit = 0
        self.org = {key(k): v for k, v in DEFAULT_ORG.items()}
        for part in g("org").split(","):           # org = ДОРИХОНА ЧИНОЗ:2, ДОРИХОНА УРГУТ:7
            if ":" in part:
                n, i = part.rsplit(":", 1)
                if i.strip().isdigit():
                    self.org[key(n)] = int(i)
        self._connect = connect
        self.d = dialect

    @property
    def enabled(self):
        return self.mode == "sql"

    @property
    def is_test(self):
        return self.database.upper().endswith("_TEST")

    # ---------- ulanish ----------
    def connect(self):
        if not self.database:
            raise FomError("config.ini [fom] database bo'sh")
        if not self.is_test and not self.live:
            raise FomError("Ishchi bazaga (%s) yozish o'chiq. Avval _TEST bazada sinang; keyin config.ini [fom] live = yes"
                           % self.database)
        if self._connect:
            return self._connect()
        try:
            import pyodbc
        except ImportError:
            raise FomError("pyodbc o'rnatilmagan: install.bat ni ishga tushiring (python -m pip install pyodbc)")
        ini = read_flat_ini(self.ini_path)
        server = self.server or ini.get("server") or "SERVER01"
        uid, pwd = (self.uid, self.pwd) if self.uid else (ini.get("uid", ""), ini.get("pwd", ""))
        cs = "DRIVER={SQL Server};SERVER=%s;DATABASE=%s;" % (server, self.database)
        cs += ("UID=%s;PWD=%s;" % (odbc_quote(uid), odbc_quote(pwd))) if uid else "Trusted_Connection=yes;"
        try:
            con = pyodbc.connect(cs, autocommit=False, timeout=15)
        except Exception as e:  # noqa
            raise FomError("FOM bazasiga ulanib bo'lmadi (%s / %s): %s" % (server, self.database, e))
        # timeout=15 faqat ulanish uchun. So'rov ham cheksiz kutmasin: FOM'da jadval band bo'lsa (operator tranzaksiyasi)
        # bot bazasining qulfi ushlanib, butun bot qotib qolardi. Endi xato bilan qaytadi va hammasi bekor qilinadi.
        con.timeout = QUERY_TIMEOUT
        try:
            c = con.cursor()
            c.execute("SET LOCK_TIMEOUT %d" % (QUERY_TIMEOUT * 1000))
            c.close()
        except Exception:  # noqa
            logging.warning("FOM: LOCK_TIMEOUT o'rnatilmadi", exc_info=True)
        return con

    # ---------- tekshiruvlar ----------
    def _checks(self, cur, dsts):
        cur.execute("SELECT INVENT, BLOCKDATA FROM SETUPS")
        r = cur.fetchone()
        if r and r[0]:
            raise FomError("FOM'da inventarizatsiya ketmoqda — bron yozilmaydi")
        if r and r[1]:
            bd = r[1] if isinstance(r[1], datetime.datetime) else datetime.datetime.fromisoformat(str(r[1]))
            if bd.date() >= datetime.date.today():
                raise FomError("FOM'da davr yopilgan (%s gacha)" % bd.date())
        cur.execute("SELECT MAX(ID) FROM SMENA WHERE [USER] = ? AND STATE = 9", (self.fom_user,))
        r = cur.fetchone()
        if not r or r[0] is None:
            raise FomError("FOM foydalanuvchisi #%d da ochiq smena yo'q" % self.fom_user)
        smena = int(r[0])
        orgs = {}
        for d in dsts:
            oid = self.org.get(key(d))
            if not oid:
                raise FomError("«%s» uchun FOM ORG raqami yo'q (config.ini [fom] org)" % d)
            cur.execute("SELECT NAME FROM ORG WHERE ID = ?", (oid,))
            r = cur.fetchone()
            if not r or key(r[0]) != key(d):
                raise FomError("FOM ORG #%d nomi «%s», kutilgan «%s» — xavfsizlik uchun to'xtatildi"
                               % (oid, r[0] if r else "?", d))
            orgs[d] = oid
        return smena, orgs

    def _stock(self, cur):
        """{(nom, ishlab chiqaruvchi): [partiya, ...]} — muddati yaqinroq birinchi (FEFO)."""
        cur.execute(
            "SELECT L.ID, G.NAME, G.PRNAME, L.KOL, L.SALE, L.ROBOTKOL, L.ROBOTSALE, L.EXPDATA, "
            "L.PRICEROZ1, L.PRICEROZ1NDS, L.PRICEROZ1NONDS "
            "FROM INCOMELN L JOIN GOOD G ON G.ID = L.GOOD "
            "WHERE L.STATE = ? AND L.KOL > L.SALE AND L.EXPDATA > ?",
            (INC_POSTED, datetime.datetime.combine(datetime.date.today(), datetime.time())))
        out = {}
        for r in cur.fetchall():
            kol, sale, rk, rs = num(r[3]), num(r[4]), num(r[5]), num(r[6])
            avail = min(kol - sale, (kol - rk) - (sale - rs))
            if avail <= 0:
                continue
            price, nds = num(r[8]), num(r[9])
            nonds = num(r[10]) if r[10] is not None else round(price - nds, 2)
            out.setdefault((key(r[1]), key(r[2])), []).append(
                {"id": int(r[0]), "avail": avail, "exp": str(r[7])[:10], "price": price, "nds": nds, "nonds": nonds})
        for v in out.values():
            v.sort(key=lambda p: (p["exp"], p["id"]))
        return out

    def _plan(self, cur, groups):
        stock = self._stock(cur)
        plans = []
        for _, d, rows in groups:
            lines, short, taken = [], [], 0
            for r in rows:
                parts = stock.get((key(r["name"]), key(r["producer"])))
                need = num(r["qty"])
                if self.limit and taken >= self.limit:
                    short.append({"task_id": r["id"], "name": r["name"], "need": need, "got": 0,
                                  "why": "sinov chegarasi (limit = %d)" % self.limit})
                    continue
                if not parts:
                    short.append({"task_id": r["id"], "name": r["name"], "need": need, "got": 0, "why": "FOM'da qoldiq yo'q"})
                    continue
                got = 0.0
                for p in parts:
                    if need - got <= 0:
                        break
                    take = min(p["avail"], need - got)
                    if take <= 0:
                        continue
                    p["avail"] -= take
                    got += take
                    lines.append({"task_id": r["id"], "incomeln": p["id"], "kol": take, "price": p["price"],
                                  "nds": p["nds"], "nonds": p["nonds"], "exp": p["exp"]})
                if got > 0:
                    taken += 1
                if got < need:
                    short.append({"task_id": r["id"], "name": r["name"], "need": need, "got": got, "why": "skladda yetmaydi"})
            plans.append({"dst": d, "lines": lines, "short": short})
        return plans

    # ---------- ochiq API ----------
    def preview(self, groups):
        """Yozmasdan: FOM'da nechta qator yoziladi, nimasi yetmaydi."""
        con = self.connect()
        try:
            cur = con.cursor()
            self._checks(cur, [g[1] for g in groups])
            plans = self._plan(cur, groups)
            con.rollback()
            return plans
        except FomError:
            raise
        except Exception as e:  # noqa — masalan, loginda huquq yo'q
            raise FomError("FOM'ni tekshirib bo'lmadi: %s" % re.sub(r"\s+", " ", str(e))[:300])
        finally:
            con.close()

    def write(self, groups, labels):
        """groups: [(i, dst, rows)], labels: {dst: matn}. Hammasi yoki hech biri.
        Qaytaradi: [{"dst", "invoice", "lines", "short"}]."""
        return self._write([g[1] for g in groups], lambda cur: self._plan(cur, groups), labels)

    # ---------- bugun kelgan yuk (kirim) ----------
    def _arrivals(self, cur, day):
        """Shu kuni skladga kelgan, hali tarqatilmagan (KOL > SALE), muddati o'tmagan partiyalar."""
        d0 = datetime.datetime.combine(day, datetime.time())
        today0 = datetime.datetime.combine(datetime.date.today(), datetime.time())
        cur.execute(
            "SELECT L.ID, G.NAME, G.PRNAME, L.KOL, L.SALE, L.ROBOTKOL, L.ROBOTSALE, L.EXPDATA, "
            "L.PRICEROZ1, L.PRICEROZ1NDS, L.PRICEROZ1NONDS, O.NAME "
            "FROM INCOMELN L JOIN GOOD G ON G.ID = L.GOOD LEFT JOIN ORG O ON O.ID = L.ORG "
            "WHERE L.STATE = ? AND L.DATA >= ? AND L.DATA < ? AND L.KOL > L.SALE AND L.EXPDATA > ?",
            (INC_POSTED, d0, d0 + datetime.timedelta(days=1), today0))
        out = []
        for r in cur.fetchall():
            kol, sale, rk, rs = num(r[3]), num(r[4]), num(r[5]), num(r[6])
            avail = min(kol - sale, (kol - rk) - (sale - rs))
            if avail <= 0:
                continue
            price, nds = num(r[8]), num(r[9])
            out.append({"id": int(r[0]), "name": str(r[1] or "").strip(), "producer": str(r[2] or "").strip(),
                        "avail": avail, "exp": str(r[7])[:10], "price": price, "nds": nds,
                        "nonds": num(r[10]) if r[10] is not None else round(price - nds, 2),
                        "supplier": str(r[11] or "").strip()})
        out.sort(key=lambda p: (p["exp"], p["id"]))
        return out

    def arrivals(self, day):
        con = self.connect()
        try:
            return self._arrivals(con.cursor(), day)
        except FomError:
            raise
        except Exception as e:  # noqa
            raise FomError("FOM'dan kirimni o'qib bo'lmadi: %s" % re.sub(r"\s+", " ", str(e))[:300])
        finally:
            con.close()

    def write_kirim(self, day, alloc, labels):
        """alloc: {dorixona: [{"name", "producer", "qty", "task_id"}]} — shu kungi kirim partiyalaridan (muddati yaqini
        birinchi). Partiyada so'ralgancha qolmagan bo'lsa (shu orada sotildi / boshqa bron) — hech narsa yozilmaydi."""
        def build(cur):
            parts = {}
            for p in self._arrivals(cur, day):
                parts.setdefault((key(p["name"]), key(p["producer"])), []).append(p)
            plans = []
            for dst, items in alloc.items():
                lines = []
                for it in items:
                    need, got = num(it["qty"]), 0.0
                    for p in parts.get((key(it["name"]), key(it["producer"])), []):
                        if need - got <= 0:
                            break
                        take = min(p["avail"], need - got)
                        if take <= 0:
                            continue
                        p["avail"] -= take
                        got += take
                        lines.append({"task_id": it["task_id"], "incomeln": p["id"], "kol": take, "price": p["price"],
                                      "nds": p["nds"], "nonds": p["nonds"], "exp": p["exp"]})
                    if got < need - 1e-9:
                        raise FomError("«%s»: kirimda %s qolgan, taqsimotda %s — sahifani yangilang"
                                       % (it["name"], ("%g" % got), ("%g" % need)))
                plans.append({"dst": dst, "lines": lines, "short": []})
            return plans
        return self._write(list(alloc), build, labels)

    def _write(self, dsts, build, labels):
        con = self.connect()
        try:
            cur = con.cursor()
            smena, orgs = self._checks(cur, dsts)
            plans = build(cur)
            skip_h, skip_l = self.d.computed(cur, "INVOICE"), self.d.computed(cur, "INVOICELN")
            for p in plans:
                if not p["lines"]:
                    p["invoice"] = None
                    continue
                label = labels.get(p["dst"], "")
                head = [("DOCTYPE", DOCTYPE_MOVE), ("NUMBER", 0), ("ORG", orgs[p["dst"]]), ("PURPOSE", label[:100]),
                        ("SMENA", smena), ("STATE", STATE_DRAFT), ("USERS", self.fom_user), ("SKLAD", 0),
                        ("COMMENTS", label[:200])]   # PAYMENT — FOM standarti (100), FOM qoralamalaridagi kabi
                head = [c for c in head if c[0] not in skip_h]
                inv = self.d.insert_id(cur, "INSERT INTO INVOICE(%s) VALUES(%s)" % (
                    ", ".join(c[0] for c in head), ", ".join("?" * len(head))), tuple(c[1] for c in head))
                cols = [c for c in ("INVOICE", "INCOMELN", "KOL", "PRICE", "PRICE1", "STATE", "USERS", "FIRSTID", "NDS",
                                    "PRICENONDS1", "PRICENONDS") if c not in skip_l]
                rows = []
                for ln in p["lines"]:
                    v = {"INVOICE": inv, "INCOMELN": ln["incomeln"], "KOL": ln["kol"], "PRICE": ln["price"],
                         "PRICE1": ln["price"], "STATE": STATE_DRAFT, "USERS": self.fom_user, "FIRSTID": 0,
                         "NDS": ln["nds"], "PRICENONDS1": ln["nonds"], "PRICENONDS": ln["nonds"]}
                    rows.append(tuple(v[c] for c in cols))
                cur.executemany("INSERT INTO INVOICELN(%s) VALUES(%s)" % (", ".join(cols), ", ".join("?" * len(cols))), rows)
                p["invoice"] = inv
            con.commit()
            return plans
        except FomError:
            con.rollback()
            raise
        except Exception as e:  # noqa — FOM triggeri xatosi (yetmaydi / inventarizatsiya / davr yopilgan)
            try:
                con.rollback()
            except Exception:  # noqa
                pass
            logging.exception("FOM bron yozilmadi")
            raise FomError("FOM bronni qabul qilmadi, hech narsa yozilmadi: %s" % re.sub(r"\s+", " ", str(e))[:700])
        finally:
            try:
                con.close()
            except Exception:  # noqa
                pass
