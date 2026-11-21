# -*- coding: utf-8 -*-
"""B7: server paneli — ko'chirish API (toifa bo'yicha KPI, cat parametri, Excel), bosh sahifa kartasi, sozlamalar,
sahifada topshiriqlar ro'yxati yo'qligi. HTTP — 127.0.0.1 dagi vaqtinchalik port."""
import base64
import io
import json
import os
import re
import sys
import threading
import unittest
import urllib.error
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fixtures import B, BotCase, CH, EK, F_CH, F_EK
import timesaver_panel as P
from openpyxl import load_workbook


class Api(BotCase):
    def setUp(self):
        super().setUp()
        self.api = P.Api(self.bot, self.bot.cfg)

    def test_today_kpi_toifalar(self):
        self.bot.koch_build()
        s = self.api.koch_today()
        self.assertEqual(s["cats"], ["shoshilinch", "assortiment", "oddiy"])
        self.assertEqual({c: s["kpi"][c]["jami"] for c in s["cats"]}, {"shoshilinch": 9, "assortiment": 4, "oddiy": 6})
        k = s["kpi"]["shoshilinch"]
        self.assertEqual((k["title"], k["icon"], k["deadline"], k["seconds_left"]), ("SHOSHILINCH", "🔴", "10:00", 7200))
        self.assertEqual(s["kpi"]["oddiy"]["deadline"], "18:00")
        self.assertEqual(s["now"], "08:00")

    def test_bosh_sahifa_kartasi(self):
        self.assertEqual(self.api.state()["koch"], {"has": False, "on": True})
        self.bot.koch_build()
        tid = self.bot.db.execute("SELECT id FROM koch_tasks WHERE name='АВИТЕН'").fetchone()[0]
        self.bot.koch_give(F_EK, self.bot.client(F_EK), tid)
        self.bot.koch_receive(F_CH, self.bot.client(F_CH), tid)
        self.clock.at(10, 30)
        k = self.api.state()["koch"]
        self.assertEqual((k["pct"], k["late"]), (5, 12), "1 / 19 olindi; 8 SHOSHILINCH + 4 ASSORTIMENT kechikkan")

    def test_tasks_cat_param(self):
        self.bot.koch_build()
        self.assertEqual(len(self.api.koch_tasks({"cat": "assortiment"})), 4)
        self.assertEqual(len(self.api.koch_tasks({"turi": "zudlik"})), 9, "eski parametr: zudlik → shoshilinch")
        self.assertEqual(len(self.api.koch_tasks({"cat": "nimadir"})), 19, "noma'lum toifa — filtrsiz")
        self.assertEqual(len(self.api.koch_tasks({"filial": CH, "cat": "oddiy"})), 6)
        t = self.api.koch_tasks({"cat": "oddiy"})[0]
        self.assertEqual((t["cat"], t["turi"]), ("oddiy", "oddiy"))

    def test_excel(self):
        self.bot.koch_build()
        content, fname = self.api.koch_excel({})
        self.assertRegex(fname, r"^Kochirish_nazorat_\d{4}-\d\d-\d\d\.xlsx$")
        wb = load_workbook(io.BytesIO(content))
        self.assertEqual(wb.sheetnames, ["Topshiriqlar", "Filiallar"])
        ws = wb["Topshiriqlar"]
        self.assertEqual(ws["B1"].value, "Toifa")
        self.assertEqual({ws.cell(r, 2).value for r in range(2, ws.max_row + 1)}, {"SHOSHILINCH", "ASSORTIMENT", "ODDIY"})

    def test_task_va_holat(self):
        self.bot.koch_build()
        tid = self.bot.db.execute("SELECT id FROM koch_tasks WHERE name='ОДДИЙ ЧАР'").fetchone()[0]
        self.assertEqual(self.api.koch_task(str(tid))["cat"], "oddiy")
        self.assertFalse(self.api.koch_task("x").get("ok", True))
        self.assertTrue(self.api.koch_set_status({"id": tid, "status": "bekor", "note": "test"})["ok"])
        self.assertFalse(self.api.koch_set_status({"id": tid, "status": "bekor", "note": ""})["ok"])

    def test_remind_regen(self):
        self.bot.koch_build()
        self.clock.at(10, 30)
        r = self.api.koch_remind({"filial": "all"})
        self.assertTrue(r["ok"])
        self.assertIn("2 filial", r["msg"])
        self.assertTrue(self.api.koch_regen({})["ok"])

    def test_history(self):
        self.bot.koch_build()
        self.clock.at(11, 0)
        h = self.api.koch_history("30")
        self.assertEqual(h["days"], 30)
        self.assertEqual(h["rank"][0]["filial"], CH)

    def test_sozlamalar(self):
        s = self.api.settings_get()
        for k in ("koch_deadline_shoshilinch", "koch_deadline_assortiment", "koch_deadline_oddiy", "koch_srok_warn_days",
                  "koch_assort_min_days", "koch_page_size"):
            self.assertIn(k, s["koch"])
        self.assertNotIn("koch_deadline_zudlik", s["koch"])
        self.assertEqual((s["koch"]["koch_assort_min_days"], s["koch"]["koch_page_size"]), ("60", "10"))
        self.assertTrue(self.api.settings_set({"koch_deadline_assortiment": "11:15"})["ok"])
        self.assertEqual(self.bot.kset("koch_deadline_assortiment"), "11:15")
        self.assertTrue(self.api.settings_set({"koch_deadline_shoshilinch": ""})["ok"], "bo'sh — muddatsiz")
        self.assertIsNone(self.bot.koch_deadline("shoshilinch"))
        self.assertFalse(self.api.settings_set({"koch_task_time": ""})["ok"])
        self.assertFalse(self.api.settings_set({"koch_page_size": "50"})["ok"])
        self.assertFalse(self.api.settings_set({"koch_srok_warn_days": "abc"})["ok"])
        self.assertTrue(self.api.settings_set({"koch_assort_min_days": "45"})["ok"])
        self.assertEqual(self.bot.kint("koch_assort_min_days"), 45)

    def test_mijoz_korinishi_filial(self):
        r = self.api.user_preview(F_CH)
        self.assertEqual(list(r["sections"])[:3], ["ШОШИЛИНЧ", "АССОРТИМЕНТ", "ОДДИЙ"])
        self.assertEqual(r["buttons"][0], ["🔴 ШОШИЛИНЧ · 0"])


class Html(unittest.TestCase):
    """panel.html: topshiriqlar jadvali, drawer va /api/koch/tasks so'rovi yo'q; KPI toifalar bo'yicha."""

    @classmethod
    def setUpClass(cls):
        with open(os.path.join(P.PANEL_DIR, "panel.html"), encoding="utf-8") as f:
            cls.html = f.read()
        cls.koch = cls.html[cls.html.index("KO'CHIRISH NAZORATI (dashboard)"):cls.html.index("async function saveMaintText")]

    def test_royxat_yoq(self):
        for s in ("/api/koch/tasks", "loadTasks", "kOpen", 'id="ktasks"', 'id="klist"', "kSetStatus"):
            self.assertNotIn(s, self.html, s)

    def test_kpi_toifalar(self):
        self.assertIn("s.cats.map(c => kpiRow(s.kpi[c]", self.koch)
        self.assertNotIn("zudlik", self.koch.lower())

    def test_sozlamalar_toifalar(self):
        for k in ("koch_deadline_shoshilinch", "koch_deadline_assortiment", "koch_deadline_oddiy", "koch_page_size"):
            self.assertIn('"%s"' % k, self.html)
        self.assertNotIn("koch_deadline_zudlik", self.html)
        self.assertNotIn("koch_rem_report_on", self.html)

    def test_rangli_toifa_tugmasi_yoq(self):
        self.assertNotIn('"danger"', self.html)


class Http(BotCase):
    """Haqiqiy HTTP: 127.0.0.1, _guard (Host / parol), HTML va JSON."""

    def setUp(self):
        super().setUp()
        self.bot.koch_build()
        api = P.Api(self.bot, self.bot.cfg)
        self.srv = P.PanelServer(("127.0.0.1", 0), P.make_handler(api, "sir"))
        threading.Thread(target=self.srv.serve_forever, daemon=True).start()
        self.base = "http://127.0.0.1:%d" % self.srv.server_address[1]

    def tearDown(self):
        self.srv.shutdown()
        self.srv.server_close()
        super().tearDown()

    def get(self, path, auth=True, host=None):
        req = urllib.request.Request(self.base + path)
        if auth:
            req.add_header("Authorization", "Basic " + base64.b64encode(b"admin:sir").decode())
        if host:
            req.add_header("Host", host)
        try:
            with urllib.request.urlopen(req, timeout=10) as r:
                return r.status, r.read()
        except urllib.error.HTTPError as e:
            return e.code, e.read()

    def test_guard(self):
        self.assertEqual(self.get("/api/koch/today", auth=False)[0], 401)
        self.assertEqual(self.get("/api/koch/today", host="evil.example.com")[0], 403)

    def test_today_va_html(self):
        code, body = self.get("/api/koch/today")
        self.assertEqual(code, 200)
        self.assertEqual(json.loads(body)["kpi"]["assortiment"]["jami"], 4)
        code, body = self.get("/")
        self.assertEqual(code, 200)
        self.assertNotIn(b"/api/koch/tasks", body)
        code, body = self.get("/api/koch/tasks?cat=oddiy")
        self.assertEqual((code, len(json.loads(body))), (200, 6), "endpoint qoladi (Excel / kelajak uchun)")


if __name__ == "__main__":
    unittest.main()
