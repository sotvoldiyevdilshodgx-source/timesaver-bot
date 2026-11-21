# -*- coding: utf-8 -*-
"""Soddalashtirilgan sozlash ustasi (4 qadam): token tekshiruvi (/api/setup/token), birinchi /start — avtomatik admin
(xavfsizlik shartlari), SQL serverni avtomatik topish (pyodbc yo'q / soxta pyodbc), bazalar ro'yxati va HTTP yo'llari."""
import json
import os
import sys
import threading
import time
import unittest
import urllib.error
import urllib.request
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fixtures import B, BotCase, add_client  # noqa: E402

import timesaver_panel as P  # noqa: E402

TOKEN = "1000000001:TEST-token-namuna-haqiqiy-emas-0000"
NEW_ADMIN, OTHER = 1112223334, 7777000111


def start_msg(cid, date, text="/start admin", chat_type="private", from_id=None, name="Akmal"):
    return {"message": {"message_id": 1, "date": int(date), "text": text, "chat": {"id": cid, "type": chat_type},
                        "from": {"id": from_id or cid, "first_name": name, "username": name.lower()}}}


class ProbeOK:
    """B.TG o'rniga (set_token ichidagi getMe tekshiruvi)."""

    def __init__(self, token, base=""):
        pass

    def call(self, method, params=None, timeout=60):
        return {"id": 1, "is_bot": True, "username": "shifo_farm_bot", "first_name": "Shifo Farm"}


class ProbeRejected(ProbeOK):
    def call(self, method, params=None, timeout=60):
        raise B.TGError('HTTP 401: {"ok":false,"error_code":401,"description":"Unauthorized"}', 401)


class ProbeOffline(ProbeOK):
    def call(self, method, params=None, timeout=60):
        raise urllib.error.URLError("getaddrinfo failed")


class FakeCursor:
    def __init__(self, rows):
        self.rows, self.sql = rows, None

    def execute(self, sql):
        self.sql = sql
        return self

    def fetchall(self):
        return [(r,) for r in self.rows]


class FakeOdbc:
    """pyodbc o'rniga: faqat `ok` serverlarga ulanadi; ulanish satrlari yozib olinadi."""

    def __init__(self, ok, dbs=("FOMBAZA", "FOMSKLAD")):
        self.ok, self.dbs, self.seen = dict(ok) if isinstance(ok, dict) else {s: list(dbs) for s in ok}, dbs, []
        self.lock = threading.Lock()

    def connect(self, cs, timeout=0):
        with self.lock:
            self.seen.append(cs)
        server = cs.split("Server=", 1)[1].split(";", 1)[0]
        if server not in self.ok:
            raise RuntimeError("[08001] SQL Server does not exist or access denied")
        rows = self.ok[server]

        class Con:
            def cursor(self_):
                return FakeCursor(rows)

            def close(self_):
                pass
        return Con()


class FakeWinreg:
    HKEY_LOCAL_MACHINE, KEY_READ, KEY_WOW64_64KEY, KEY_WOW64_32KEY = 1, 0x20019, 0x100, 0x200

    def __init__(self, names, fail=False):
        self.names, self.fail = names, fail

    def OpenKey(self, root, path, res, access):
        if self.fail or access & self.KEY_WOW64_32KEY:  # 32-bit ko'rinishida yo'q
            raise OSError("not found")
        names = self.names

        class K:
            def __enter__(s):
                return s

            def __exit__(s, *a):
                return False
        k = K()
        k.names = names
        return k

    def EnumValue(self, k, i):
        if i >= len(k.names):
            raise OSError("no more")
        return (k.names[i], "MSSQL16." + k.names[i], 1)


class SqlAniqlash(unittest.TestCase):
    def test_nomzodlar_tartibi_va_takrorsiz(self):
        c = P.sql_candidates("SRV01", [".\\SQLEXPRESS", "."], computer="PC1")
        self.assertEqual(c, ["SRV01", ".\\SQLEXPRESS", ".", "localhost", "(local)", "PC1", "PC1\\SQLEXPRESS"])
        c = P.sql_candidates("", [".\\sqlexpress"], computer="")
        self.assertEqual(c, [".\\sqlexpress", ".", "localhost", "(local)"], "katta-kichik harf takrori yo'q, kompyuter nomi yo'q")

    def test_reestrdagi_nusxalar(self):
        self.assertEqual(P.sql_local_instances(FakeWinreg(["MSSQLSERVER", "SQLEXPRESS"])), [".", ".\\SQLEXPRESS"])
        self.assertEqual(P.sql_local_instances(FakeWinreg([], fail=True)), [])

    def test_pyodbc_yoq(self):
        with mock.patch.object(P, "_pyodbc", lambda: None):
            r = P.sql_detect(ini={"server": "SRV01", "database": "FOMBAZA"}, instances=[], computer="PC1")
        self.assertFalse(r["ok"])
        self.assertFalse(r["pyodbc"])
        self.assertEqual(r["server"], "SRV01", "db.ini dagi server taklif qilinadi")
        self.assertEqual(r["candidates"][0], "SRV01")
        self.assertEqual(r["database"], "FOMBAZA")
        with mock.patch.object(P, "_pyodbc", lambda: None):
            self.assertEqual(P.sql_detect(ini={}, instances=[], computer="")["server"], "")

    def test_birinchi_ishlagani_tanlanadi(self):
        odbc = FakeOdbc({"(local)": ["FOMBAZA", "FOMSKLAD"], "localhost": ["X"]})
        r = P.sql_detect(ini={}, instances=[".\\SQLEXPRESS"], odbc=odbc, computer="PC1")
        self.assertTrue(r["ok"], r)
        self.assertEqual(r["server"], "localhost", "tartibdagi birinchi ishlagan server")
        self.assertEqual(r["databases"], ["X"])
        self.assertEqual(r["auth"], "windows")
        self.assertEqual([t["server"] for t in r["tried"]], [".\\SQLEXPRESS", "."])
        self.assertTrue(all("Trusted_Connection=yes" in cs and "Database=master" in cs for cs in odbc.seen))

    def test_login_db_ini_dan(self):
        odbc = FakeOdbc(["SRV01"])
        r = P.sql_detect(ini={"server": "SRV01", "uid": "bot", "pwd": "p;w}d"}, instances=[], odbc=odbc, computer="")
        self.assertTrue(r["ok"])
        self.assertEqual((r["server"], r["auth"], r["uid"]), ("SRV01", "login", "bot"))
        self.assertNotIn("pwd", r, "parol javobga qo'shilmaydi")
        self.assertIn("Uid={bot};Pwd={p;w}}d};", odbc.seen[0])

    def test_hech_biri_ishlamadi(self):
        r = P.sql_detect(ini={}, instances=[], odbc=FakeOdbc([]), computer="")
        self.assertFalse(r["ok"])
        self.assertTrue(r["pyodbc"])
        self.assertEqual(len(r["tried"]), 4)
        self.assertIn("08001", r["tried"][0]["detail"])

    def test_bazalar_royxati(self):
        dbs, err = P.sql_databases("SRV", odbc=FakeOdbc(["SRV"], dbs=["A", "B"]))
        self.assertEqual((dbs, err), (["A", "B"], None))
        dbs, err = P.sql_databases("NOPE", odbc=FakeOdbc(["SRV"]))
        self.assertIsNone(dbs)
        self.assertIn("08001", err)


class UstaCase(BotCase):
    clients = False

    def setUp(self):
        super().setUp()
        self._base, self._ini = B.BASE_DIR, B.DB_INI
        B.BASE_DIR = self.work
        B.DB_INI = os.path.join(self.work, "BazaReport", "db.ini")
        # yangi o'rnatish: token va admin yo'q, ustasi tugamagan
        self.bot.cfg.read_dict({"app": {"setup_done": "no"}, "bot": {"token": "", "admin_ids": ""}})
        self.bot.admins = set()
        self.api = P.Api(self.bot, self.bot.cfg)

    def tearDown(self):
        B.BASE_DIR, B.DB_INI = self._base, self._ini
        super().tearDown()

    def set_token(self, probe=ProbeOK):
        with mock.patch.object(B, "TG", probe):
            return self.api.setup_token({"token": TOKEN})


class Token(UstaCase):
    def test_token_togri(self):
        r = self.set_token()
        self.assertTrue(r["ok"], r)
        self.assertEqual(r["username"], "shifo_farm_bot")
        self.assertTrue(self.bot.token_ok())
        at = int(self.bot.setting(P.SETUP_TOKEN_AT))
        self.assertLessEqual(abs(at - time.time()), 5)
        with open(os.path.join(self.work, "config.ini"), encoding="utf-8") as f:
            self.assertIn(TOKEN, f.read())
        g = self.api.setup_get()
        self.assertTrue(g["waiting_admin"], "token kiritildi, admin yo'q — /start kutilmoqda")
        self.assertEqual(g["username"], "shifo_farm_bot")

    def test_token_xato_turlari(self):
        r = self.api.setup_token({"token": "12345:abc"})
        self.assertEqual((r["ok"], r["kind"]), (False, "format"))
        self.assertEqual(self.api.setup_token({"token": " "})["kind"], "format")
        r = self.set_token(ProbeRejected)
        self.assertEqual((r["ok"], r["kind"]), (False, "rejected"))
        r = self.set_token(ProbeOffline)
        self.assertEqual((r["ok"], r["kind"]), (False, "network"))
        self.assertFalse(self.bot.token_ok(), "xato token saqlanmaydi")
        self.assertIsNone(self.bot.setting(P.SETUP_TOKEN_AT), "xato tokenda /start oynasi ochilmaydi")
        self.assertFalse(self.api.setup_get()["waiting_admin"])


class AvtoAdmin(UstaCase):
    def test_birinchi_start_admin_boladi(self):
        self.set_token()
        self.bot.handle(start_msg(NEW_ADMIN, time.time()))
        self.assertEqual(self.bot.admins, {NEW_ADMIN})
        self.assertEqual(self.bot.setting(P.SETUP_AUTO_ADMIN), str(NEW_ADMIN))
        texts = self.tg.texts(NEW_ADMIN)
        self.assertTrue(any("admin qilindingiz" in t for t in texts), texts)
        self.assertGreaterEqual(len(texts), 2, "shu /start bilan admin menyusi ham ochiladi")
        with open(os.path.join(self.work, "config.ini"), encoding="utf-8") as f:
            self.assertIn("admin_ids = %d" % NEW_ADMIN, f.read())
        g = self.api.setup_get()
        self.assertEqual(g["admins_info"], [{"id": NEW_ADMIN, "name": "Akmal"}])
        self.assertNotIn(NEW_ADMIN, [c["id"] for c in g["candidates"]], "admin nomzodlar ro'yxatida turmaydi")
        self.assertFalse(g["waiting_admin"])
        # ikkinchi odam — admin bo'lmaydi, faqat ro'yxatga tushadi
        self.bot.handle(start_msg(OTHER, time.time(), name="Begona"))
        self.assertEqual(self.bot.admins, {NEW_ADMIN})
        self.assertIn(OTHER, [c["id"] for c in self.api.setup_get()["candidates"]])

    def test_tokendan_oldingi_start_admin_bolmaydi(self):
        self.set_token()
        at = int(self.bot.setting(P.SETUP_TOKEN_AT))
        self.bot.handle(start_msg(OTHER, at - 60, name="Eski"))  # Telegram navbatida qolgan eski xabar
        self.assertEqual(self.bot.admins, set())
        c = self.api.setup_get()["candidates"]
        self.assertEqual(c[0]["id"], OTHER, "qo'lda admin qilish uchun ro'yxatda turadi")
        self.assertEqual(c[0]["name"], "Eski")

    def test_admin_bor_bolsa_yoq(self):
        self.set_token()
        self.bot.set_admins({111222})
        self.bot.handle(start_msg(NEW_ADMIN, time.time()))
        self.assertEqual(self.bot.admins, {111222})

    def test_ustasi_tugagan_bolsa_yoq(self):
        self.set_token()
        self.bot.app_set("setup_done", "yes")
        self.bot.handle(start_msg(NEW_ADMIN, time.time()))
        self.assertEqual(self.bot.admins, set())

    def test_token_ustada_kiritilmagan_bolsa_yoq(self):
        self.bot.cfg["bot"]["token"] = TOKEN  # masalan «Sozlamalar» dan yoki config.ini da — vaqt belgisi yo'q
        self.bot.handle(start_msg(NEW_ADMIN, time.time()))
        self.assertEqual(self.bot.admins, set())
        self.assertFalse(self.api.setup_get()["waiting_admin"])

    def test_guruh_va_boshqa_matnlar_yoq(self):
        self.set_token()
        now = time.time()
        self.bot.handle(start_msg(-100500, now, chat_type="group", from_id=NEW_ADMIN))
        self.bot.handle(start_msg(NEW_ADMIN, now, text="salom"))
        self.bot.handle(start_msg(NEW_ADMIN, now, text="/myid"))
        self.assertEqual(self.bot.admins, set())
        self.assertFalse(P.auto_admin(self.bot, start_msg(NEW_ADMIN, now, from_id=OTHER)["message"]),
                         "yuboruvchi va chat bir xil bo'lishi kerak")
        m = start_msg(NEW_ADMIN, now)["message"]
        m["from"]["is_bot"] = True
        self.assertFalse(P.auto_admin(self.bot, m))
        self.assertEqual(self.bot.admins, set())

    def test_ilgak_bir_marta(self):
        P.Api(self.bot, self.bot.cfg)  # ikkinchi Api — on_message qayta o'ralmaydi
        self.set_token()
        with mock.patch.object(P, "auto_admin", wraps=P.auto_admin) as aa:
            self.bot.handle(start_msg(NEW_ADMIN, time.time()))
        self.assertEqual(aa.call_count, 1)
        self.assertEqual(self.bot.admins, {NEW_ADMIN})

    def test_eski_nomzodlar_ham(self):
        add_client(self.bot, 777001, "firm", status="pending", name="Ali")
        self.set_token()
        self.bot.handle(start_msg(OTHER, int(self.bot.setting(P.SETUP_TOKEN_AT)) - 5, name="Vali"))
        ids = [c["id"] for c in self.api.setup_get()["candidates"]]
        self.assertEqual(ids, [OTHER, 777001])


class SqlEndpoint(UstaCase):
    def test_bazalar(self):
        self.assertFalse(self.api.setup_sql_dbs({"server": " "})["ok"])
        with mock.patch.object(P, "_pyodbc", lambda: None):
            r = self.api.setup_sql_dbs({"server": "SRV"})
        self.assertEqual((r["ok"], r["pyodbc"]), (False, False))
        odbc = FakeOdbc(["SRV"], dbs=["FOMBAZA"])
        with mock.patch.object(P, "_pyodbc", lambda: odbc):
            r = self.api.setup_sql_dbs({"server": "SRV"})
            self.assertEqual((r["ok"], r["databases"]), (True, ["FOMBAZA"]))
            r = self.api.setup_sql_dbs({"server": "NOPE"})
        self.assertFalse(r["ok"])
        self.assertIn("08001", r["detail"])
        self.assertNotIn("08001", r["error"], "ODBC matni alohida (kirillga o'girilmaydi)")

    def test_saqlangan_parol(self):
        self.api.setup_set({"sql": {"server": "SRV", "database": "DB", "uid": "bot", "pwd": "Maxfiy1"}})
        odbc = FakeOdbc(["SRV"])
        with mock.patch.object(P, "_pyodbc", lambda: odbc):
            self.assertTrue(self.api.setup_sql_dbs({"server": "SRV", "uid": "bot", "pwd": ""})["ok"])
            self.api.setup_sql_dbs({"server": "SRV", "uid": "boshqa", "pwd": ""})
        self.assertIn("Pwd={Maxfiy1}", odbc.seen[0], "login bir xil — saqlangan parol")
        self.assertIn("Pwd={}", odbc.seen[1], "boshqa login — saqlangan parol berilmaydi")

    def test_detect_db_ini_bilan(self):
        self.api.setup_set({"sql": {"server": "SRV9", "database": "FOMBAZA", "uid": "", "pwd": ""}})
        odbc = FakeOdbc(["SRV9"])
        with mock.patch.object(P, "_pyodbc", lambda: odbc), mock.patch.object(P, "sql_local_instances", lambda: []):
            r = self.api.setup_sql_detect({})
        self.assertTrue(r["ok"])
        self.assertEqual((r["server"], r["database"]), ("SRV9", "FOMBAZA"))

    def test_setup_get_yangilash_qadamlari(self):
        self.bot.rstate = {"running": False, "ok": True, "msg": "ok", "steps": {"Sale": ("ok", "27"), "Srok": ("err", "x")}}
        self.assertEqual(self.api.setup_get()["refresh"]["steps"], {"Sale": ["ok", "27"], "Srok": ["err", "x"]})


class Http(UstaCase):
    def setUp(self):
        super().setUp()
        self.srv = P.PanelServer(("127.0.0.1", 0), P.make_handler(self.api, ""))
        threading.Thread(target=self.srv.serve_forever, daemon=True).start()
        self.url = "http://127.0.0.1:%d" % self.srv.server_address[1]
        self.op = urllib.request.build_opener(urllib.request.ProxyHandler({}))

    def tearDown(self):
        self.srv.shutdown()
        self.srv.server_close()
        super().tearDown()

    def post(self, path, body, csrf=True):
        req = urllib.request.Request(self.url + path, data=json.dumps(body).encode(), method="POST",
                                     headers={"Content-Type": "application/json", **({"X-Requested-With": "panel"} if csrf else {})})
        with self.op.open(req, timeout=10) as r:
            return json.loads(r.read().decode("utf-8"))

    def test_csrf(self):
        for path in ("/api/setup/token", "/api/setup/sql_detect", "/api/setup/sql_dbs"):
            with self.assertRaises(urllib.error.HTTPError) as e:
                self.post(path, {"token": TOKEN}, csrf=False)
            self.assertEqual(e.exception.code, 403, path)
        self.assertFalse(self.bot.token_ok())

    def test_token_http(self):
        with mock.patch.object(B, "TG", ProbeOK):
            r = self.post("/api/setup/token", {"token": TOKEN})
        self.assertTrue(r["ok"], r)
        self.assertIn("@shifo_farm_bot", r["msg"])
        r = self.post("/api/setup/token", {"token": "x"})
        self.assertEqual(r["kind"], "format")
        self.assertRegex(r["error"], "[А-Яа-яЎўҚқҒғҲҳ]", "xabar kirillda")

    def test_sql_http(self):
        with mock.patch.object(P, "sql_detect", lambda: {"ok": True, "server": ".", "databases": ["A"], "candidates": ["."]}):
            r = self.post("/api/setup/sql_detect", {})
        self.assertEqual(r["server"], ".")
        with mock.patch.object(P, "_pyodbc", lambda: FakeOdbc(["SRV"], dbs=["A"])):
            r = self.post("/api/setup/sql_dbs", {"server": "SRV"})
        self.assertEqual(r["databases"], ["A"])


class Html(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with open(os.path.join(P.PANEL_DIR, "panel.html"), encoding="utf-8") as f:
            cls.html = f.read()

    def test_tort_qadam(self):
        self.assertIn('const WSTEPS = ["Бот", "Маълумот", "Синов", "Тайёр"];', self.html)
        for s in ("/api/setup/token", "/api/setup/sql_detect", "/api/setup/sql_dbs", "/api/setup/upload", "?start=admin",
                  "https://t.me/BotFather", "Токенни қандай оламан?", "Ботни очиш", "Қўлда киритиш", "Логин/парол билан",
                  "Синовни бошлаш", "Панелни очиш"):
            self.assertIn(s, self.html, s)


if __name__ == "__main__":
    unittest.main()
