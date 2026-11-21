# -*- coding: utf-8 -*-
"""2-bosqich: FOM'ga qoralama bron (INVOICE DOCTYPE 15 / STATE 2 + INVOICELN). Haqiqiy SQL Server o'rniga SQLite'dagi
FOM jadvallari nusxasi; SALE ni oshiruvchi va «skladda yetmaydi» triggeri ham bor. Tekshiriladi: FEFO, yetmaganlar,
xavfsizlik (_TEST / live), inventarizatsiya, ORG nomi, smena, FOM rad etsa — ikki tomonda ham hech narsa yozilmasligi."""
import datetime
import io
import json
import os
import sqlite3
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from fixtures import ADMIN, CH  # noqa: E402
from openpyxl import load_workbook  # noqa: E402
from test_sklad import Case  # noqa: E402
import fom_sql  # noqa: E402
import timesaver_panel as P  # noqa: E402

D = lambda n: (datetime.date.today() + datetime.timedelta(days=n)).isoformat() + " 00:00:00"  # noqa: E731

SCHEMA = """
CREATE TABLE GOOD(ID INTEGER PRIMARY KEY, NAME TEXT, PRNAME TEXT);
CREATE TABLE INCOMELN(ID INTEGER PRIMARY KEY, GOOD INT, KOL REAL, SALE REAL DEFAULT 0, ROBOTKOL REAL, ROBOTSALE REAL,
    EXPDATA TEXT, STATE INT, PRICEROZ1 REAL, PRICEROZ1NDS REAL, PRICEROZ1NONDS REAL, DATA TEXT, ORG INT);
CREATE TABLE ORG(ID INTEGER PRIMARY KEY, NAME TEXT);
CREATE TABLE SMENA(ID INTEGER PRIMARY KEY, [USER] INT, STATE INT);
CREATE TABLE SETUPS(INVENT INT, BLOCKDATA TEXT);
CREATE TABLE INVOICE(ID INTEGER PRIMARY KEY AUTOINCREMENT, DOCTYPE INT, NUMBER INT, ORG INT, PURPOSE TEXT, SMENA INT,
    STATE INT, USERS INT, SKLAD INT, PAYMENT INT DEFAULT 100, COMMENTS TEXT);
CREATE TABLE INVOICELN(ID INTEGER PRIMARY KEY AUTOINCREMENT, INVOICE INT, INCOMELN INT, KOL REAL, PRICE REAL, PRICE1 REAL,
    STATE INT, USERS INT, FIRSTID INT, NDS REAL, PRICENONDS1 REAL, PRICENONDS REAL);
CREATE TRIGGER INVOICELN_INSERT AFTER INSERT ON INVOICELN BEGIN
    UPDATE INCOMELN SET SALE = SALE + NEW.KOL WHERE ID = NEW.INCOMELN;
    SELECT RAISE(ABORT, 'У вас нет такого количества на складе!') WHERE EXISTS(SELECT 1 FROM INCOMELN WHERE SALE > KOL);
END;
INSERT INTO ORG VALUES(1, 'ДОРИХОНА СКЛАД'), (2, 'ДОРИХОНА ЧИНОЗ'), (7, 'ДОРИХОНА УРГУТ'), (9, 'ДОРИХОНА  ЭРКИН');
INSERT INTO SMENA VALUES(3, 3, 9), (8, 1, 9), (16, 49, 9), (7, 1, 10);
INSERT INTO SETUPS VALUES(0, '2000-01-01 00:00:00');
INSERT INTO GOOD VALUES(10, 'СКЛ ДОРИ 1', 'РОССИЯ'), (11, 'СКЛ  ДОРИ 1', 'РОССИЯ'), (20, 'СКЛ ДОРИ 2', 'ИНДИЯ'),
    (30, 'СКЛ АРАЛАШ', 'ИНДИЯ'), (40, 'СКЛ АССОРТ', 'ГЕРМАНИЯ');
"""


class Lite:
    computed_cols = {}

    @staticmethod
    def insert_id(cur, sql, params):
        cur.execute(sql, params)
        return cur.lastrowid

    @classmethod
    def computed(cls, cur, table):
        return set(cls.computed_cols.get(table, ()))


class FomCase(Case):
    def setUp(self):
        super().setUp()
        self.bot.koch_build()
        fd, self.fpath = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        con = sqlite3.connect(self.fpath)
        con.executescript(SCHEMA)
        con.executemany("INSERT INTO INCOMELN(ID, GOOD, KOL, SALE, ROBOTKOL, ROBOTSALE, EXPDATA, STATE, PRICEROZ1, PRICEROZ1NDS, PRICEROZ1NONDS) "
                        "VALUES(?,?,?,?,?,?,?,?,?,?,?)", [
                            (100, 10, 5, 2, None, None, D(400), 3, 1120, 120, 1000),   # ДОРИ 1: 3 ta, muddati uzoqroq
                            (101, 11, 50, 0, 0, 0, D(60), 3, 1120, 120, 1000),         # dublikat tovar, muddati yaqin
                            (102, 10, 99, 0, None, None, D(-1), 3, 1120, 120, 1000),   # muddati o'tgan — olinmaydi
                            (103, 10, 99, 0, None, None, D(500), 12, 1120, 120, 1000), # STATE 12 — olinmaydi
                            (104, 20, 9, 0, None, None, D(90), 3, 2240, 240, 2000),
                            (105, 30, 1, 0, None, None, D(200), 3, 900, 0, None)])     # АРАЛАШ: 1 ta (kerak 2)
        con.commit()
        con.close()
        self.bot.cfg.add_section("fom")
        self.bot.cfg.set("fom", "mode", "sql")
        self.bot.cfg.set("fom", "database", "FOMSKLADSKLAD_TEST")
        self.use()

    def tearDown(self):
        super().tearDown()
        os.remove(self.fpath)

    def use(self):
        self.bot._fom_sql = fom_sql.FomSql(self.bot.cfg, connect=lambda: sqlite3.connect(self.fpath), dialect=Lite)

    def q(self, sql, *a):
        con = sqlite3.connect(self.fpath)
        try:
            return con.execute(sql, a).fetchall()
        finally:
            con.close()

    def fom(self, sql, *a):
        con = sqlite3.connect(self.fpath)
        con.execute(sql, a)
        con.commit()
        con.close()

    def grp(self, dst):
        return [g for g in self.bot.fom_summary()["groups"] if g["dst"] == dst][0]

    def create(self, dst):
        g = self.grp(dst)
        return self.bot.fom_create(str(g["i"]), g["sig"])

    def local(self):
        return self.bot.db.execute("SELECT * FROM fom_boron ORDER BY id").fetchall()


class Write(FomCase):
    def test_qoralama_fefo_va_yetmaganlar(self):
        b = self.create(CH)[0]
        inv = self.q("SELECT ID, DOCTYPE, NUMBER, ORG, PURPOSE, SMENA, STATE, USERS, SKLAD, PAYMENT FROM INVOICE")
        self.assertEqual(len(inv), 1)
        self.assertEqual(inv[0][1:], (15, 0, 2, "BOT bron #%d" % b["id"], 8, 2, 1, 0, 100), "PAYMENT — FOM standarti")
        lines = sorted(self.q("SELECT INCOMELN, KOL, PRICE, PRICE1, STATE, NDS, PRICENONDS FROM INVOICELN"))
        # ДОРИ 1 (5 ta): avval muddati yaqin 101 (dublikat tovar) dan 5; АРАЛАШ: 105 dan 1 (yana 1 yetmaydi)
        self.assertEqual(lines, [(101, 5.0, 1120.0, 1120.0, 2, 120.0, 1000.0), (105, 1.0, 900.0, 900.0, 2, 0.0, 900.0)])
        self.assertEqual(self.q("SELECT SALE FROM INCOMELN WHERE ID IN (101, 105) ORDER BY ID"), [(5.0,), (1.0,)])
        loc = self.local()[0]
        self.assertEqual((loc["status"], loc["doc_no"], loc["mode"]), ("qoralama", str(inv[0][0]), "sql"))
        short = {x["name"]: (x["got"], x["why"]) for x in json.loads(loc["short"])}
        self.assertEqual(short, {"СКЛ АРАЛАШ": (1.0, "skladda yetmaydi"), "СКЛ АССОРТ": (0, "FOM'da qoldiq yo'q")})
        self.assertIn("FOM qoralama #%s" % inv[0][0], self.tg.texts(ADMIN)[-1])
        ws = load_workbook(io.BytesIO(self.bot.fom_file(b["id"])[0])).active
        self.assertEqual([c.value for c in ws[1]][-2:], ["FOM'га ёзилди", "Изоҳ"])
        got = {r[1]: (r[6], r[7]) for r in ws.iter_rows(min_row=2, values_only=True)}
        self.assertEqual(got["СКЛ ДОРИ 1"], (5, None))
        self.assertEqual(got["СКЛ АССОРТ"], (0, "FOM'da qoldiq yo'q"))
        m = self.bot.fom_summary()["made"][0]
        self.assertEqual((m["status"], m["doc"], m["short"]), ("qoralama", str(inv[0][0]), 2))
        self.assertIn("FOM", self.bot.fom_cancel(b["id"]), "FOM qoralamasi panelda bekor qilinmaydi")

    def test_hisoblanadigan_ustunlarga_yozilmaydi(self):
        # FOM'da PRICENONDS1 / PRICENONDS — hisoblanadigan ustun (SQL Server xatosi 271)
        Lite.computed_cols = {"INVOICELN": ("PRICENONDS1", "PRICENONDS")}
        try:
            con = sqlite3.connect(self.fpath)
            con.execute("CREATE TRIGGER NOCOMP BEFORE INSERT ON INVOICELN WHEN NEW.PRICENONDS IS NOT NULL OR NEW.PRICENONDS1 IS NOT NULL "
                        "BEGIN SELECT RAISE(ABORT, 'Столбец PRICENONDS изменить нельзя (271)'); END")
            con.commit()
            con.close()
            self.create(CH)
            self.assertEqual(len(self.q("SELECT * FROM INVOICELN")), 2)
            self.assertEqual(self.q("SELECT COUNT(*) FROM INVOICELN WHERE PRICENONDS IS NOT NULL"), [(0,)])
        finally:
            Lite.computed_cols = {}

    def test_sinov_chegarasi(self):
        self.bot.cfg.set("fom", "limit", "1")
        self.use()
        b = self.create(CH)[0]
        self.assertEqual(len({r[0] for r in self.q("SELECT INCOMELN FROM INVOICELN")}), 1, "faqat 1 ta dori yozildi")
        why = [x["why"] for x in json.loads(b["short"])]
        self.assertIn("sinov chegarasi (limit = 1)", why)

    def test_fefo_bolib_olish(self):
        self.fom("UPDATE INCOMELN SET KOL = 2 WHERE ID = 101")   # yaqin partiyada 2 ta qoldi → qolgan 3 ta 100 dan
        self.create(CH)
        self.assertEqual(self.q("SELECT INCOMELN, KOL FROM INVOICELN WHERE INCOMELN IN (100, 101) ORDER BY ID"),
                         [(101, 2.0), (100, 3.0)])

    def test_hammasi_va_org_raqamlari(self):
        s = self.bot.fom_summary()
        self.bot.fom_create("a", s["all_sig"])
        orgs = sorted(r[0] for r in self.q("SELECT ORG FROM INVOICE"))
        self.assertEqual(orgs, [2, 7, 9], "ЧИНОЗ 2, УРГУТ 7, ЭРКИН 9 (ORG nomidagi ikki probel)")

    def test_tekshiruv_yozmaydi(self):
        g = self.grp(CH)
        r = self.bot.fom_check(str(g["i"]), g["sig"])
        self.assertTrue(r["ok"], r)
        self.assertEqual((r["plans"][0]["lines"], r["plans"][0]["kol"], len(r["plans"][0]["short"])), (2, 6.0, 2))
        self.assertEqual(self.q("SELECT COUNT(*) FROM INVOICE"), [(0,)])
        self.assertEqual(self.q("SELECT SALE FROM INCOMELN WHERE ID = 101"), [(0.0,)])
        self.assertEqual(self.local(), [])


class Safety(FomCase):
    def nothing(self):
        self.assertEqual(self.q("SELECT COUNT(*) FROM INVOICE"), [(0,)])
        self.assertEqual(self.q("SELECT COUNT(*) FROM INVOICELN"), [(0,)])
        self.assertEqual(self.local(), [], "bot bazasiga ham yozilmaydi")
        self.assertIn(CH, [g["dst"] for g in self.bot.fom_summary()["groups"]], "dorilar ro'yxatda qoladi")

    def fails(self, part):
        with self.assertRaises(fom_sql.FomError) as e:
            self.create(CH)
        self.assertIn(part, str(e.exception))
        self.nothing()

    def test_ishchi_baza_live_siz_yopiq(self):
        self.bot.cfg.set("fom", "database", "FOMSKLADSKLAD")
        self.use()
        self.fails("live = yes")
        self.bot.cfg.set("fom", "live", "yes")
        self.use()
        self.create(CH)
        self.assertEqual(len(self.q("SELECT * FROM INVOICE")), 1)

    def test_inventarizatsiya(self):
        self.fom("UPDATE SETUPS SET INVENT = 1")
        self.fails("inventarizatsiya")

    def test_davr_yopilgan(self):
        self.fom("UPDATE SETUPS SET BLOCKDATA = ?", D(1))
        self.fails("davr yopilgan")

    def test_smena_yoq(self):
        self.fom("DELETE FROM SMENA WHERE [USER] = 1 AND STATE = 9")
        self.fails("smena")

    def test_org_nomi_mos_emas(self):
        self.fom("UPDATE ORG SET NAME = 'БОШҚА АПТЕКА' WHERE ID = 2")
        self.fails("xavfsizlik")

    def test_fom_rad_etsa_hammasi_bekor(self):
        # FOM triggeri rad etadi (masalan, kassa shu payt sotib yubordi): ikkinchi dorixona yozilayotganda
        con = sqlite3.connect(self.fpath)
        con.execute("CREATE TRIGGER RAD AFTER INSERT ON INVOICELN WHEN NEW.INCOMELN = 104 BEGIN "
                    "SELECT RAISE(ABORT, 'У вас нет такого количества на складе!'); END")
        con.commit()
        con.close()
        with self.assertRaises(fom_sql.FomError) as e:
            self.bot.fom_create("a", self.bot.fom_summary()["all_sig"])
        self.assertIn("hech narsa yozilmadi", str(e.exception))
        self.nothing()
        self.assertEqual(self.q("SELECT SALE FROM INCOMELN WHERE ID = 101"), [(0.0,)], "FOM qoldig'i qaytdi")

    def test_pyodbc_yoq_yoki_ulanmasa(self):
        self.bot._fom_sql = fom_sql.FomSql(self.bot.cfg)   # haqiqiy ulanish: bu yerda server yo'q
        with self.assertRaises(fom_sql.FomError):
            self.create(CH)
        self.assertEqual(self.local(), [])

    def test_fayl_rejimi_ozgarmagan(self):
        self.bot.cfg.set("fom", "mode", "fayl")
        self.use()
        b = self.create(CH)[0]
        self.assertEqual((b["mode"], b["status"]), ("fayl", "fayl"))
        self.assertEqual(self.q("SELECT COUNT(*) FROM INVOICE"), [(0,)])


class Ini(unittest.TestCase):
    def test_db_ini(self):
        fd, p = tempfile.mkstemp(suffix=".ini")
        os.write(fd, "Server = SRV1\r\nDatabase=FOMBAZA\r\nUid=botuser\r\nPwd=a=b\r\n".encode())
        os.close(fd)
        try:
            self.assertEqual(fom_sql.read_flat_ini(p), {"server": "SRV1", "database": "FOMBAZA", "uid": "botuser", "pwd": "a=b"})
        finally:
            os.remove(p)


class PanelCheck(FomCase):
    def test_api(self):
        self.assertTrue(P.Api(self.bot, self.bot.cfg))
        s = self.bot.fom_summary()
        self.assertEqual((s["mode"], s["db"]), ("sql", "FOMSKLADSKLAD_TEST"))
        r = self.bot.fom_check("a", "notogri")
        self.assertFalse(r["ok"])


if __name__ == "__main__":
    unittest.main()
