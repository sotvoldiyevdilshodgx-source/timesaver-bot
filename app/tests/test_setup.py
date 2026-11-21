# -*- coding: utf-8 -*-
"""TimeSaver: sozlash ustasi (nom, SQL db.ini, Excel yuklash, tugatish), ma'lumot papkasi, logo va panel."""
import json
import os
import shutil
import sys
import threading
import unittest
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fixtures import B, BotCase, make_xlsx  # noqa: E402

import timesaver_panel as P  # noqa: E402


class Setup(BotCase):
    clients = False

    def setUp(self):
        super().setUp()
        self._base, self._ini = B.BASE_DIR, B.DB_INI
        B.BASE_DIR = self.work
        B.DB_INI = os.path.join(self.work, "BazaReport", "db.ini")
        self.api = P.Api(self.bot, self.bot.cfg)

    def tearDown(self):
        B.BASE_DIR, B.DB_INI = self._base, self._ini
        super().tearDown()

    def test_eski_ornatish_tugagan_hisoblanadi(self):
        self.assertTrue(self.bot.setup_done, "token bor, [app] yo'q — eski o'rnatish")
        self.bot.cfg.read_dict({"app": {"setup_done": "no"}})
        self.assertFalse(self.bot.setup_done)

    def test_nom(self):
        self.assertEqual(self.bot.app_name, "TimeSaver")
        r = self.api.setup_set({"name": "  Shifo   Farm ", "sklad_name": "markaziy sklad"})
        self.assertTrue(r["ok"], r)
        self.assertEqual(self.bot.app_name, "Shifo Farm")
        self.assertEqual(self.bot.data.sklad_name, B.norm("markaziy sklad"))
        with open(os.path.join(self.work, "config.ini"), encoding="utf-8") as f:
            text = f.read()
        self.assertIn("name = Shifo Farm", text)
        self.assertFalse(self.api.setup_set({"name": "  "})["ok"])

    def test_sql_db_ini_makros_formatida(self):
        r = self.api.setup_set({"sql": {"server": "SRV1", "database": "FOMBAZA", "uid": "bot", "pwd": "Parol123"}})
        self.assertTrue(r["ok"], r)
        with open(B.DB_INI, "rb") as f:
            raw = f.read()
        self.assertIn(b"Server=SRV1\r\nDatabase=FOMBAZA\r\nUid=bot\r\nPwd=Parol123\r\n", raw)
        ini = P.read_db_ini()
        self.assertEqual((ini["server"], ini["uid"], ini["pwd"]), ("SRV1", "bot", "Parol123"))
        # parol bo'sh qoldirilsa — eskisi saqlanadi
        self.api.setup_set({"sql": {"server": "SRV2", "database": "FOMBAZA", "uid": "bot", "pwd": ""}})
        self.assertEqual(P.read_db_ini()["pwd"], "Parol123")
        g = self.api.setup_get()["sql"]
        self.assertTrue(g["has_pwd"])
        self.assertNotIn("pwd", g, "parol panelga qaytmaydi")
        self.assertFalse(self.api.setup_set({"sql": {"server": "", "database": "X"}})["ok"])

    def test_windows_autentifikatsiya(self):
        self.api.setup_set({"sql": {"server": "SRV", "database": "DB", "uid": "", "pwd": ""}})
        self.assertNotIn("uid", P.read_db_ini())

    def test_excel_yuklash(self):
        src = make_xlsx(os.path.join(self.work, "manba.xlsx"), "new", None)
        with open(src, "rb") as f:
            body = f.read()
        r = self.api.setup_upload("baza_yangi.xlsx", body)
        self.assertTrue(r["ok"], r)
        dest = os.path.join(self.work, "baza", "baza_yangi.xlsx")
        self.assertEqual(self.bot.data.find_path(), dest)
        self.assertTrue(self.bot.data.sale)
        # qayta yuklansa eskisi .bak bo'lib qoladi
        self.api.setup_upload("baza_yangi.xlsx", body)
        self.assertTrue([n for n in os.listdir(os.path.join(self.work, "baza")) if n.endswith(".bak")])
        self.assertFalse(self.api.setup_upload("virus.exe", body)["ok"])
        self.assertFalse(self.api.setup_upload("..\\..\\x.xlsx", b"short")["ok"])

    def test_yuklash_yoldan_tashqariga_chiqmaydi(self):
        src = make_xlsx(os.path.join(self.work, "m.xlsx"), "new", None)
        with open(src, "rb") as f:
            r = self.api.setup_upload("..\\..\\evil.xlsx", f.read())
        self.assertTrue(r["ok"], r)
        self.assertTrue(os.path.exists(os.path.join(self.work, "baza", "evil.xlsx")))

    def test_buzilgan_excel(self):
        before = self.bot.data.find_path()
        r = self.api.setup_upload("buzuq.xlsm", b"x" * 5000)
        self.assertFalse(r["ok"])
        self.assertEqual(self.bot.data.find_path(), before, "buzuq fayl ulanmaydi, eski fayl qoladi")
        self.assertTrue(self.bot.data.sale)

    def test_tugatish_tekshiradi(self):
        self.bot.cfg["bot"]["token"] = ""
        self.assertFalse(self.api.setup_set({"finish": True})["ok"])
        self.bot.cfg["bot"]["token"] = "1:x"
        self.bot.admins = set()
        self.assertIn("admin", self.api.setup_set({"finish": True})["error"])
        self.bot.admins = {1}
        r = self.api.setup_set({"finish": True})
        self.assertTrue(r["ok"], r)
        self.assertTrue(self.bot.setup_done)

    def test_nomzodlar(self):
        from fixtures import add_client
        add_client(self.bot, 777001, "firm", status="pending", name="Ali")
        c = self.api.setup_get()["candidates"]
        self.assertEqual([x["id"] for x in c], [777001])


class Http(BotCase):
    clients = False

    def setUp(self):
        super().setUp()
        self._base = B.BASE_DIR
        B.BASE_DIR = self.work
        self.srv = P.PanelServer(("127.0.0.1", 0), P.make_handler(P.Api(self.bot, self.bot.cfg), ""))
        threading.Thread(target=self.srv.serve_forever, daemon=True).start()
        self.url = "http://127.0.0.1:%d" % self.srv.server_address[1]
        self.op = urllib.request.build_opener(urllib.request.ProxyHandler({}))

    def tearDown(self):
        self.srv.shutdown()
        self.srv.server_close()
        B.BASE_DIR = self._base
        super().tearDown()

    def test_logo_va_panel(self):
        with self.op.open(self.url + "/", timeout=10) as r:
            html = r.read().decode("utf-8")
        self.assertIn("TimeSaver", html)
        self.assertIn('id="sideNav"', html)
        if os.path.exists(os.path.join(B.APP_DIR, "branding", "logo.svg")):
            with self.op.open(self.url + "/logo.svg", timeout=10) as r:
                self.assertEqual(r.headers["Content-Type"], "image/svg+xml")

    def test_yuklash_http(self):
        src = make_xlsx(os.path.join(self.work, "m.xlsx"), "new", None)
        with open(src, "rb") as f:
            body = f.read()
        req = urllib.request.Request(self.url + "/api/setup/upload", data=body, method="POST",
                                     headers={"X-Requested-With": "panel", "X-File-Name": "baza%20yangi.xlsx"})
        with self.op.open(req, timeout=20) as r:
            res = json.loads(r.read().decode("utf-8"))
        self.assertTrue(res["ok"], res)
        self.assertTrue(os.path.exists(os.path.join(self.work, "baza", "baza yangi.xlsx")))

    def test_yuklash_csrf(self):
        req = urllib.request.Request(self.url + "/api/setup/upload", data=b"x" * 2000, method="POST")
        with self.assertRaises(urllib.error.HTTPError) as e:
            self.op.open(req, timeout=10)
        self.assertEqual(e.exception.code, 403)

    def test_setup_get(self):
        with self.op.open(self.url + "/api/setup", timeout=10) as r:
            s = json.loads(r.read().decode("utf-8"))
        self.assertEqual(s["brand"], "TimeSaver")
        self.assertTrue(s["excel"]["found"])


class DataDir(unittest.TestCase):
    def test_env(self):
        import importlib
        import tempfile
        d = tempfile.mkdtemp(prefix="ts_home_")
        old = os.environ.get("TIMESAVER_HOME")
        os.environ["TIMESAVER_HOME"] = d
        try:
            self.assertEqual(B._data_dir(), d)
        finally:
            if old is None:
                os.environ.pop("TIMESAVER_HOME", None)
            else:
                os.environ["TIMESAVER_HOME"] = old
            shutil.rmtree(d, ignore_errors=True)
        importlib.invalidate_caches()

    def test_config_namunadan(self):
        import tempfile
        d = tempfile.mkdtemp(prefix="ts_home_")
        base = B.BASE_DIR
        try:
            B.BASE_DIR = d
            path = B.ensure_config()
            if os.path.exists(os.path.join(B.APP_DIR, "config.example.ini")):
                cfg = B.load_config(path)
                self.assertEqual(cfg.get("panel", "port"), "8765")
                self.assertEqual(cfg.get("bot", "token"), "")
        finally:
            B.BASE_DIR = base
            shutil.rmtree(d, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
