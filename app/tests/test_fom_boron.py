# -*- coding: utf-8 -*-
"""Server paneli: Sklad > «📤 FOM bron» — dorixona bo'yicha bron fayli, ikki marta yaratmaslik, ro'yxat o'zgarsa himoya.
Sklad xodimining Telegram tugmalari o'zgarmaydi. FOM (SQL Server) ga hech narsa yozilmaydi."""
import base64
import io
import json
import os
import sys
import threading
import unittest
import urllib.error
import urllib.parse
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fixtures import ADMIN, CH, EK, UCH, cb, inline_data, kb_texts  # noqa: E402
from openpyxl import load_workbook  # noqa: E402
from test_sklad import F_SK, SK, Case  # noqa: E402
import timesaver_panel as P  # noqa: E402


class Fom(Case):
    def setUp(self):
        super().setUp()
        self.bot.koch_build()

    def bron(self):
        return self.bot.db.execute("SELECT * FROM fom_boron ORDER BY id").fetchall()

    def grp(self, dst):
        return [g for g in self.bot.fom_summary()["groups"] if g["dst"] == dst][0]

    def test_telegram_tugmasi_yoq(self):
        self.assertNotIn("📤 FOM bron", kb_texts(self.bot.client_panel(self.c)[1]))
        self.assertIn("🏭 Taqsimot", kb_texts(self.bot.client_panel(self.c)[1]))

    def test_barcha_toifalar_muddati_otgansiz(self):
        s = self.bot.fom_summary()
        self.assertTrue(s["has"])
        self.assertEqual(s["cats"], ["🔴 SHOSHILINCH", "🟡 ASSORTIMENT", "🟢 ODDIY"])
        names = {i["name"] for g in s["groups"] for i in g["items"]}
        self.assertNotIn("СКЛ ЭСКИ", names, "muddati o'tgan kirmaydi")
        for g in s["groups"]:
            self.assertEqual(sum(g["by_cat"].values()), g["n"])
        exp = {}
        for t in self.bot.koch_tasks(self.bot._today(), src=SK, statuses=("yangi", "berildi")):
            why = self.bot.koch_exclude_reason(self.bot._task_row(t)) or ""
            if why != "Muddati o'tgan":
                exp[t["dst"]] = exp.get(t["dst"], 0) + 1
        self.assertEqual({g["dst"]: g["n"] for g in s["groups"]}, exp, "sklad topshirig'ining hammasi (toifadan qat'i nazar)")

    def test_config_faqat_shoshilinch(self):
        self.bot.cfg.add_section("fom")
        self.bot.cfg.set("fom", "cats", "shoshilinch")
        s = self.bot.fom_summary()
        self.assertEqual(s["cats"], ["🔴 SHOSHILINCH"])
        self.assertEqual(sorted(g["dst"] for g in s["groups"]), sorted([CH, UCH]))
        self.assertNotIn(EK, [g["dst"] for g in s["groups"]], "EK da faqat muddati o'tgan SHOSHILINCH va ODDIY bor")

    def test_yaratish_va_fayl(self):
        g = self.grp(CH)
        made = self.bot.fom_create(str(g["i"]), g["sig"])
        self.assertEqual(len(made), 1)
        b = self.bron()[0]
        self.assertEqual((b["src"], b["dst"], b["cats"], b["status"], b["created_by"]), (SK, CH, "assortiment,shoshilinch", "fayl", "admin (panel)"))
        content, fname = self.bot.fom_file(b["id"])
        self.assertTrue(fname.startswith("Bron_%d_" % b["id"]))
        ws = load_workbook(io.BytesIO(content)).active
        self.assertEqual([c.value for c in ws[1]], ["№", "Дори", "Ишлаб чиқарувчи", "Сони", "Муддат", "Тоифа"], "narx yo'q")
        self.assertEqual(sorted(r[1] for r in ws.iter_rows(min_row=2, values_only=True)), ["СКЛ АРАЛАШ", "СКЛ АССОРТ", "СКЛ ДОРИ 1"])
        self.assertIn("FOM bron yaratildi", self.tg.texts(ADMIN)[-1])
        self.assertNotIn(CH, [x["dst"] for x in self.bot.fom_summary()["groups"]])
        self.assertEqual([m["dst"] for m in self.bot.fom_summary()["made"]], [CH])

    def test_ikki_marta_yaratilmaydi(self):
        s = self.bot.fom_summary()
        self.assertEqual(len(self.bot.fom_create("a", s["all_sig"])), 3)
        self.assertIsNone(self.bot.fom_create("a", s["all_sig"]))
        self.assertEqual(len(self.bron()), 3)
        self.assertEqual(self.bot.fom_summary()["groups"], [])

    def test_royxat_ozgarsa_yaratilmaydi(self):
        sig = self.bot.fom_summary()["all_sig"]
        with self.bot.dblock:
            self.bot.db.execute("UPDATE koch_tasks SET qty_plan=qty_plan+1 WHERE src=? AND name=?", (SK, "СКЛ ДОРИ 1"))
            self.bot.db.commit()
        self.assertIsNone(self.bot.fom_create("a", sig))
        self.assertEqual(self.bron(), [])

    def test_bekor_qilish(self):
        g = self.grp(CH)
        b = self.bot.fom_create(str(g["i"]), g["sig"])[0]
        self.assertNotIn(CH, [x["dst"] for x in self.bot.fom_summary()["groups"]])
        self.assertIsNone(self.bot.fom_cancel(b["id"]))
        self.assertIn("bekor qilindi", self.tg.texts(ADMIN)[-1])
        s = self.bot.fom_summary()
        self.assertIn(CH, [x["dst"] for x in s["groups"]], "dorilar ro'yxatga qaytdi")
        self.assertEqual(s["made"], [])
        self.assertIsNotNone(self.bot.fom_cancel(b["id"]), "ikkinchi marta bekor qilinmaydi")
        g = self.grp(CH)
        self.assertEqual(len(self.bot.fom_create(str(g["i"]), g["sig"])), 1, "qayta bron qilish mumkin")
        with self.bot.dblock:
            self.bot.db.execute("UPDATE fom_boron SET status='qoralama' WHERE dst=? AND status='fayl'", (CH,))
            self.bot.db.commit()
        b2 = self.bot.fom_today()[0]
        self.assertIn("FOM", self.bot.fom_cancel(b2["id"]), "FOM'ga yozilgani bekor qilinmaydi")

    def test_koch_tasks_ozgarmaydi(self):
        before = self.bot.db.execute("SELECT id, status, qty_plan FROM koch_tasks ORDER BY id").fetchall()
        self.bot.fom_create("a", self.bot.fom_summary()["all_sig"])
        after = self.bot.db.execute("SELECT id, status, qty_plan FROM koch_tasks ORDER BY id").fetchall()
        self.assertEqual([tuple(r) for r in before], [tuple(r) for r in after])

    def test_berilgani_ham_kiradi(self):
        i = self.bot.koch_route_dsts(SK, "shoshilinch").index(CH)
        give = [d for d in inline_data(self.bot.koch_list_text(self.c, "shoshilinch")[1]) if d.startswith("kg:shoshilinch:%d:" % i)][0]
        self.bot.handle(cb(F_SK, give.replace("kg:", "kgy:", 1)))
        self.assertIn(CH, [g["dst"] for g in self.bot.fom_summary()["groups"]])

    def test_topshiriq_yoq(self):
        self.bot.set_setting("koch_on", "0")
        s = self.bot.fom_summary()
        self.assertEqual((s["has"], s["groups"]), (False, []))


class Http(Case):
    def setUp(self):
        super().setUp()
        self.bot.koch_build()
        self.srv = P.PanelServer(("127.0.0.1", 0), P.make_handler(P.Api(self.bot, self.bot.cfg), "sir"))
        threading.Thread(target=self.srv.serve_forever, daemon=True).start()
        self.base = "http://127.0.0.1:%d" % self.srv.server_address[1]

    def tearDown(self):
        self.srv.shutdown()
        self.srv.server_close()
        super().tearDown()

    def req(self, path, body=None, xrw=True):
        r = urllib.request.Request(self.base + path, data=None if body is None else json.dumps(body).encode())
        r.add_header("Authorization", "Basic " + base64.b64encode(b"admin:sir").decode())
        if body is not None:
            r.add_header("Content-Type", "application/json")
            if xrw:
                r.add_header("X-Requested-With", "panel")
        try:
            with urllib.request.urlopen(r, timeout=10) as x:
                return x.status, dict(x.headers), x.read()
        except urllib.error.HTTPError as e:
            return e.code, dict(e.headers), e.read()

    def test_api_toliq_oqim(self):
        code, _, body = self.req("/api/fom")
        s = json.loads(body)
        self.assertEqual((code, len(s["groups"])), (200, 3))
        self.assertEqual(self.req("/api/fom/create", {"which": "a", "sig": s["all_sig"]}, xrw=False)[0], 403, "CSRF himoya")
        code, _, body = self.req("/api/fom/create", {"which": "a", "sig": s["all_sig"]})
        r = json.loads(body)
        self.assertTrue(r["ok"], r)
        self.assertEqual(len(r["ids"]), 3)
        r2 = json.loads(self.req("/api/fom/create", {"which": "a", "sig": s["all_sig"]})[2])
        self.assertFalse(r2["ok"])
        code, h, body = self.req("/api/fom/file?id=%d" % r["ids"][0])
        self.assertEqual(code, 200)
        self.assertIn("Bron_%d_" % r["ids"][0], urllib.parse.unquote(h["Content-Disposition"]))
        load_workbook(io.BytesIO(body))
        self.assertEqual(self.req("/api/fom/file?id=999")[0], 404)
        r3 = json.loads(self.req("/api/fom/cancel", {"id": r["ids"][0]})[2])
        self.assertTrue(r3["ok"], r3)
        self.assertEqual(len(json.loads(self.req("/api/fom")[2])["groups"]), 1)
        self.assertFalse(json.loads(self.req("/api/fom/cancel", {"id": "x"})[2])["ok"])
        self.assertEqual(self.req("/api/fom/file?id=abc")[0], 404)

    def test_html_bloki(self):
        with open(os.path.join(P.PANEL_DIR, "panel.html"), encoding="utf-8") as f:
            html = f.read()
        for s in ('id="fombox"', "fomLoad()", "/api/fom/create", "/api/fom/file?id=", "/api/fom/cancel"):
            self.assertIn(s, html)


if __name__ == "__main__":
    unittest.main()
