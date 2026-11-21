# -*- coding: utf-8 -*-
"""Sklad: Skladdan varag'ini o'qish, muddat himoyasi, sklad xodimi paneli, dorixona qabuli, server paneli, admin ko'rinishi."""
import datetime
import io
import json
import os
import sys
import threading
import unittest
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fixtures import B, BotCase, CH, UCH, EK, ADMIN, F_CH, F_EK, TODAY, add_client, cb, day, inline_data, kb_texts, msg
from openpyxl import load_workbook

import timesaver_panel as P  # noqa: E402

SK = "ДОРИХОНА СКЛАД"
F_SK = 6100
HEAD = ["Qayerga", "Dori", "Ishlab chiqaruvchi", "Skladda qoldiq", "Sotuv", "Filialda qoldiq", "Sotilmayotgan filial(lar)",
        "Ulardagi qoldiq", "Soni", "Turi", "Sabab", "Kirim narxi", "Summa (kirim)"]
ROWS = [(CH, "СКЛ ДОРИ 1", "РОССИЯ", 20, 5, 0, None, 0, 5, "SHOSHILINCH", "Tugagan, sotuv bor", 1000, 5000),
        (UCH, "СКЛ ДОРИ 1", "РОССИЯ", 20, 3, 0, None, 0, 3, "SHOSHILINCH", "Tugagan, sotuv bor", 1000, 3000),
        (EK, "СКЛ ДОРИ 2", "ИНДИЯ", 10, 4, 1, None, 0, 3, "ODDIY", "Skladda bor, oluvchida yetmaydi", 2000, 6000),
        (CH, "СКЛ АССОРТ", "ИНДИЯ", 40, 0, 0, None, 0, 1, "ASSORTIMENT", "Filialda yo'q, boshqa filialda sotilyapti", 500, 500),
        (EK, "СКЛ ЭСКИ", "ИНДИЯ", 6, 2, 0, None, 0, 2, "SHOSHILINCH", "Tugagan, sotuv bor", 700, 1400),   # faqat muddati o'tgan
        (CH, "СКЛ АРАЛАШ", "ИНДИЯ", 9, 2, 0, None, 0, 2, "SHOSHILINCH", "Tugagan, sotuv bor", 900, 1800)]  # o'tgan + yaroqli
SROK = [(SK, "СКЛ ДОРИ 1", "РОССИЯ", "A", day(200), 200, 20, 1000), (SK, "СКЛ ДОРИ 2", "ИНДИЯ", "A", day(100), 100, 10, 2000),
        (SK, "СКЛ АССОРТ", "ИНДИЯ", "A", day(300), 300, 40, 500), (SK, "СКЛ ЭСКИ", "ИНДИЯ", "A", day(-3), -3, 6, 700),
        (SK, "СКЛ АРАЛАШ", "ИНДИЯ", "A", day(-1), -1, 4, 900), (SK, "СКЛ АРАЛАШ", "ИНДИЯ", "A", day(150), 150, 5, 900)]


def add_sklad(bot, rows=ROWS, srok=SROK):
    path = bot.data.find_path()
    wb = load_workbook(path)
    s = wb["Srok"]
    for r in srok:
        s.append(list(r))
    w = wb.create_sheet("Skladdan")
    w.append(HEAD)
    for r in rows:
        w.append(list(r))
    wb.save(path)
    bot.data.reload(force=True)


class Case(BotCase):
    def setUp(self):
        super().setUp()
        add_sklad(self.bot)
        add_client(self.bot, F_SK, "filial", filial=SK, name="Omborchi")
        self.c = self.bot.client(F_SK)

    def tasks(self, **kw):
        q = "SELECT * FROM koch_tasks WHERE src=?" + "".join(" AND %s=?" % k for k in kw)
        return self.bot.db.execute(q, (SK,) + tuple(kw.values())).fetchall()


class Read(Case):
    def row(self, name, dst=None):
        return [r for r in self.bot.data.koch if r["name"] == name and (dst is None or norm_(r["dst"]) == dst)][0]

    def test_qatorlar_oqildi(self):
        r = self.row("СКЛ ДОРИ 1", CH)
        self.assertEqual((r["src"], r["qty"], r["src_stock"], r["dst_sales"], r["summa"], r["cat"]), (SK, 5, 20, 5, 5000, "shoshilinch"))
        self.assertEqual(r["date"], (TODAY + datetime.timedelta(days=200)))
        self.assertEqual(self.row("СКЛ ДОРИ 2")["cat"], "oddiy")
        self.assertEqual(self.row("СКЛ АССОРТ")["cat"], "assortiment")

    def test_skladdan_varogi_bolmasa_ishlaydi(self):
        self.assertTrue([r for r in self.bot.data.koch if r["src"] == SK])
        wb_path = self.bot.data.find_path()
        wb = load_workbook(wb_path)
        del wb["Skladdan"]
        wb.save(wb_path)
        self.bot.data.reload(force=True)
        self.assertIsNone(self.bot.data.error)
        self.assertEqual([r for r in self.bot.data.koch if r["src"] == SK], [])

    def test_sklad_filiallar_royxatida(self):
        self.assertIn(SK, self.bot.data.koch_filials())
        self.assertEqual(self.bot.find_filial("sklad")[0], SK)

    def test_faqat_muddati_otgani_topshiriqqa_kirmaydi(self):
        self.bot.koch_build()
        names = {t["name"] for t in self.tasks()}
        self.assertNotIn("СКЛ ЭСКИ", names)
        self.assertEqual(self.bot.koch_exclude_reason(self.row("СКЛ ЭСКИ")), "Muddati o'tgan")

    def test_aralash_partiyada_yaroqli_muddat(self):
        self.assertEqual(self.row("СКЛ АРАЛАШ")["date"], TODAY + datetime.timedelta(days=150))
        self.bot.koch_build()
        self.assertEqual(len(self.tasks(name="СКЛ АРАЛАШ")), 1)

    def test_limit_skladga_tasir_qilmaydi(self):
        self.bot.set_setting("koch_assort_limit", 1)
        extra = [(CH, "СКЛ АС %d" % i, "ИНДИЯ", 9, 0, 0, None, 0, 1, "ASSORTIMENT", "Filialda yo'q, boshqa filialda sotilyapti", 500, 500) for i in range(5)]
        srok = [(SK, "СКЛ АС %d" % i, "ИНДИЯ", "A", day(300), 300, 9, 500) for i in range(5)]
        path = self.bot.data.find_path()
        wb = load_workbook(path)
        for r in extra:
            wb["Skladdan"].append(list(r))
        for r in srok:
            wb["Srok"].append(list(r))
        wb.save(path)
        self.bot.data.reload(force=True)
        self.bot.koch_build()
        self.assertEqual(len(self.tasks(cat="assortiment")), 6)
        self.assertEqual(self.bot.koch_view(SK, "assortiment")[1], [], "skladda navbat yo'q")


def norm_(s):
    return B.norm(s)


class Panel(Case):
    def setUp(self):
        super().setUp()
        self.bot.koch_build()

    def test_sklad_paneli(self):
        text, kb = self.bot.client_panel(self.c)
        self.assertIn("Sklad paneli", text)
        self.assertIn("Dorixonalarga:", text)
        self.assertNotIn("Sizga keladi", text)
        self.assertFalse([t for t in kb_texts(kb) if "Keladigan" in t])
        self.assertTrue([t for t in kb_texts(kb) if t.startswith("🔴 SHOSHILINCH · 3")])

    def test_filial_paneli_ozgarmadi(self):
        text, kb = self.bot.client_panel(self.bot.client(F_CH))
        self.assertIn("Sizga keladi", text)
        self.assertTrue([t for t in kb_texts(kb) if "Keladigan" in t])

    def test_royxat_dorixona_boyicha(self):
        text, kb = self.bot.koch_list_text(self.c, "shoshilinch")
        self.assertIn("Dorixonada tugagan", text)
        self.assertIn("skladda 20", text)
        self.assertEqual(len([d for d in inline_data(kb) if d.startswith("kg:")]), 2, "ЧИНОЗ va УРГУТ")
        self.assertIn("СКЛ ЭСКИ", text, "muddati o'tgani alohida ko'rsatiladi")

    def test_ommaviy_berish_va_dorixona_qabuli(self):
        i = self.bot.koch_route_dsts(SK, "shoshilinch").index(CH)
        btn = [d for d in inline_data(self.bot.koch_list_text(self.c, "shoshilinch")[1]) if d.startswith("kg:shoshilinch:%d:" % i)][0]
        self.bot.handle(cb(F_SK, btn.replace("kg:", "kgy:", 1)))
        self.assertEqual({t["status"] for t in self.tasks(dst=CH, cat="shoshilinch")}, {"berildi"})
        self.assertIn(SK + " dan keladi", self.tg.texts(F_CH)[-1])
        text, kb = self.bot.incoming_list(CH)
        v = [d for d in inline_data(kb) if d.startswith("kv:")]
        self.assertEqual(len(v), 1)
        self.assertIn("← %s" % SK, text)
        self.bot.handle(cb(F_CH, v[0].replace("kv:", "kvy:", 1)))
        self.assertEqual({t["status"] for t in self.tasks(dst=CH, cat="shoshilinch")}, {"olindi"})
        self.assertIn("qabul qildi", self.tg.texts(F_SK)[-1])

    def test_excel(self):
        self.bot.handle(cb(F_SK, "kx:shoshilinch"))
        wb = load_workbook(io.BytesIO(self.tg.to(F_SK, "doc")[-1][3]))
        self.assertIn("СКЛ ЭСКИ", [r[1] for r in wb["Ko'chirilmaydi"].iter_rows(min_row=2, values_only=True)])

    def test_boshqa_filial_ro_yxatida_yoq(self):
        self.assertNotIn("СКЛ ДОРИ 1", self.bot.koch_list_text(self.bot.client(F_EK), "shoshilinch")[0])


def sheet_text(wb):
    return " ".join(str(v) for ws in wb.worksheets for row in ws.iter_rows(values_only=True) for v in row if v is not None)


class Taqsimot(Case):
    def setUp(self):
        super().setUp()
        self.bot.koch_build()

    def press(self):
        self.bot.handle(msg(F_SK, "🏭 Taqsimot"))  # eski pastki tugma → 🏭 Taqsimot darajasi (yangi panel, Excel tugmasi bilan)
        self.assertIn("kt:x", inline_data(self.tg.last(F_SK)[3]))
        self.bot.handle(cb(F_SK, "kt:x"))
        return self.tg.to(F_SK, "doc")

    def book(self):
        return load_workbook(io.BytesIO(self.press()[-1][3]))

    def test_tugma_faqat_skladda(self):
        self.assertIn("🏭 Taqsimot", kb_texts(self.bot.client_panel(self.c)[1]))
        self.assertNotIn("🏭 Taqsimot", kb_texts(self.bot.client_panel(self.bot.client(F_CH))[1]))

    def test_dorixona_xodimi_ololmaydi(self):
        self.bot.handle(msg(F_CH, "🏭 Taqsimot"))
        self.bot.handle(cb(F_CH, "kt:x"))
        self.bot.handle(cb(F_CH, "kt"))
        self.assertEqual(self.tg.to(F_CH, "doc"), [])
        self.assertFalse([t for t in self.tg.texts(F_CH) if "Sklad taqsimoti" in t], "dorixona xodimiga taqsimot ko'rinmaydi")

    def test_varaqlar_va_jadval(self):
        docs = self.press()
        self.assertTrue(docs[-1][2].startswith("Sklad_taqsimot_"))
        wb = load_workbook(io.BytesIO(docs[-1][3]))
        self.assertEqual(wb.sheetnames, ["Taqsimot"] + sorted([CH, EK, UCH]) + ["Berilmaydi"])
        ws = wb["Taqsimot"]
        head = [c.value for c in ws[1]]
        self.assertEqual(head[:5], ["Dori", "Ishlab chiqaruvchi", "Skladda qoldiq", "Eng yaqin muddat", "Toifa"])
        self.assertEqual(head[-2:], ["Jami beriladi", "Skladda qoladi"])
        r = [row for row in ws.iter_rows(min_row=2, values_only=True) if row[0] == "СКЛ ДОРИ 1"][0]
        self.assertEqual((r[2], r[4], r[head.index(CH)], r[head.index(UCH)], r[-2], r[-1]), (20, "SHOSHILINCH", 5, 3, 8, 12))

    def test_toifa_ranglari(self):
        ws = self.book()["Taqsimot"]
        head = [c.value for c in ws[1]]
        rows = {r[0].value: r for r in ws.iter_rows(min_row=2)}
        color = lambda c: c.fill.fgColor.rgb[-6:]
        self.assertEqual(color(rows["СКЛ ДОРИ 1"][4]), "F8C9C9")
        self.assertEqual(color(rows["СКЛ ДОРИ 1"][head.index(CH)]), "F8C9C9")
        self.assertEqual(color(rows["СКЛ ДОРИ 2"][head.index(EK)]), "C6E9D1")
        self.assertEqual(color(rows["СКЛ АССОРТ"][head.index(CH)]), "FCE8A6")

    def test_aralash_toifa_hammasi_yoziladi(self):
        path = self.bot.data.find_path()
        wb = load_workbook(path)
        wb["Skladdan"].append([EK, "СКЛ ДОРИ 1", "РОССИЯ", 20, 1, 5, None, 0, 1, "ODDIY", "Skladda bor, oluvchida yetmaydi", 1000, 1000])
        wb.save(path)
        self.bot.data.reload(force=True)
        self.bot.koch_build()
        ws = self.book()["Taqsimot"]
        head = [c.value for c in ws[1]]
        r = [row for row in ws.iter_rows(min_row=2) if row[0].value == "СКЛ ДОРИ 1"][0]
        self.assertEqual(r[4].value, "SHOSHILINCH + ODDIY")
        self.assertEqual(r[head.index(EK)].fill.fgColor.rgb[-6:], "C6E9D1", "bu dorixonada ODDIY")

    def test_narx_yoq(self):
        wb = self.book()
        for ws in wb.worksheets:
            head = [c.value for c in ws[1]]
            self.assertFalse({"Summa", "Kirim narxi"} & set(head), ws.title)
        self.assertNotIn("7300", sheet_text(wb))

    def test_xabarda_summa_yoq(self):
        caption = self.press()[-1][4]
        self.assertIn("Sklad taqsimoti", caption)
        self.assertIn("🔴", caption)
        self.assertNotIn("so'm", caption)

    def test_toifa_excellarida_narx_yoq(self):
        for cat in ("shoshilinch", "assortiment", "oddiy"):
            self.bot.handle(cb(F_SK, "kx:%s" % cat))
            doc = self.tg.to(F_SK, "doc")[-1]
            wb = load_workbook(io.BytesIO(doc[3]))
            for ws in wb.worksheets:
                head = [c.value for c in ws[1]]
                self.assertFalse({"Summa", "Kirim narxi"} & set(head), (cat, ws.title))
                self.assertEqual(ws.cell(1, ws.max_column).value, head[-1])
            self.assertNotIn("so'm", doc[4])
            ws = wb.worksheets[0]
            total = [c.value for c in ws[ws.max_row]]
            self.assertEqual(total[0], "JAMI")
            self.assertEqual(len(total), len([c.value for c in ws[1]]), "JAMI ustunlari siljimagan")

    def test_toifa_matnlarida_summa_yoq(self):
        text, kb = self.bot.koch_list_text(self.c, "shoshilinch")
        self.assertNotIn("so'm", text)
        btn = [d for d in inline_data(kb) if d.startswith("kg:")][0]
        _, cat, i, sig = btn.split(":")
        self.assertNotIn("so'm", self.bot.koch_route_card(SK, cat, int(i), sig)[0])
        self.assertNotIn("Summa", self.bot.koch_item_card(self.c, self.tasks()[0]["id"])[0])

    def test_dorixona_narxlari_ozgarmadi(self):
        rows = [r for r in self.bot.data.koch if B.norm(r["src"]) == SK]
        wb = load_workbook(io.BytesIO(B.build_koch_cat_xlsx(rows, [], "shoshilinch", CH)))
        self.assertIn("Kirim narxi", [c.value for c in wb.worksheets[0][1]])
        self.assertEqual(self.bot.koch_sum(CH, [{"summa": 1500}]), " · <b>1 500</b> so'm")

    def test_admin_exceli_summali(self):
        self.bot.handle(cb(ADMIN, "a:skx"))
        wb = load_workbook(io.BytesIO(self.tg.to(ADMIN, "doc")[-1][3]))
        self.assertIn("Summa", [c.value for c in wb[CH][1]])
        self.assertIn("Toifa", [c.value for c in wb["Taqsimot"][1]])

    def test_taqsimot_yoq(self):
        self.bot.db.execute("DELETE FROM koch_tasks")
        self.bot.db.commit()
        self.bot.data.koch = []
        self.bot.data.reload = lambda *a, **k: None
        self.bot.handle(msg(F_SK, "🏭 Taqsimot"))
        self.assertEqual(self.tg.to(F_SK, "doc"), [])
        self.assertIn("Taqsimot yo'q", " ".join(self.tg.texts(F_SK)))


class Summary(Case):
    def test_reja_topshiriqsiz(self):
        self.bot.db.execute("DELETE FROM koch_tasks")
        self.bot.db.commit()
        s = self.bot.sklad_summary()
        self.assertFalse(s["has"])
        self.assertEqual(s["kpi"]["items"], 4)
        self.assertEqual(s["kpi"]["n"], 5)
        self.assertEqual(s["items"][0]["to"][CH]["status"], "reja")

    def test_taqsimot(self):
        self.bot.koch_build()
        s = self.bot.sklad_summary()
        self.assertTrue(s["has"])
        self.assertEqual(s["dst"], sorted([CH, UCH, EK]))
        i = [x for x in s["items"] if x["name"] == "СКЛ ДОРИ 1"][0]
        self.assertEqual((i["stock"], i["total"], i["left"]), (20, 8, 12))
        self.assertEqual((i["to"][CH]["qty"], i["to"][UCH]["qty"]), (5, 3))
        self.assertEqual([r["name"] for r in s["excluded"]], ["СКЛ ЭСКИ"])
        g = {x["filial"]: x for x in s["dorixona"]}
        self.assertEqual((g[CH]["n"], g[CH]["qty"], g[CH]["summa"]), (3, 8, 7300))
        self.assertEqual(s["items"][0]["rank"], 0, "SHOSHILINCH birinchi")

    def test_holat_yangilanadi(self):
        self.bot.koch_build()
        t = self.tasks(dst=CH, cat="assortiment")[0]
        self.bot.db.execute("UPDATE koch_tasks SET status='olindi' WHERE id=?", (t["id"],))
        self.bot.db.commit()
        k = self.bot.sklad_summary()["kpi"]
        self.assertEqual((k["given"], k["taken"], k["left"]), (1, 1, 4))

    def test_excel_varaqlari(self):
        self.bot.koch_build()
        wb = load_workbook(io.BytesIO(B.build_sklad(self.bot.sklad_summary())))
        self.assertEqual(wb.sheetnames, ["Taqsimot"] + sorted([CH, EK, UCH]) + ["Berilmaydi"])
        head = [c.value for c in wb["Taqsimot"][1]]
        self.assertIn(CH, head)

    def test_admin(self):
        self.bot.koch_build()
        self.bot.handle(msg(ADMIN, "/start"))
        self.assertIn("a:sk", inline_data(self.tg.last(ADMIN, "Foydalanuvchilar")[3]))
        self.bot.handle(cb(ADMIN, "a:sk"))
        t = self.tg.texts(ADMIN)[-1]
        self.assertIn("Sklad taqsimoti", t)
        self.assertIn("Muddati sababli berilmaydi", t)
        self.bot.handle(cb(ADMIN, "a:skx"))
        self.assertTrue(self.tg.to(ADMIN, "doc")[-1][2].startswith("Sklad_taqsimot_"))

    def test_mijoz_admin_tugmasini_bosa_olmaydi(self):
        self.bot.handle(cb(F_CH, "a:sk"))
        self.bot.handle(cb(F_CH, "a:skx"))
        self.assertEqual(self.tg.to(F_CH, "doc"), [])

    def test_bosh_taqsimot(self):
        self.bot.data.koch = []
        self.assertIn("Taqsimot yo'q", self.bot.sklad_admin_text())


class Api(Case):
    def setUp(self):
        super().setUp()
        self.bot.koch_build()
        self.srv = P.PanelServer(("127.0.0.1", 0), P.make_handler(P.Api(self.bot, self.bot.cfg), ""))
        self.srv.bind_host = "127.0.0.1"
        threading.Thread(target=self.srv.serve_forever, daemon=True).start()
        self.url = "http://127.0.0.1:%d" % self.srv.server_address[1]
        self.op = urllib.request.build_opener(urllib.request.ProxyHandler({}))

    def tearDown(self):
        self.srv.shutdown()
        self.srv.server_close()
        super().tearDown()

    def test_api(self):
        with self.op.open(self.url + "/api/sklad", timeout=10) as r:
            s = json.loads(r.read().decode("utf-8"))
        self.assertEqual(s["sklad"], SK)
        self.assertEqual(s["kpi"]["items"], 4)

    def test_excel(self):
        with self.op.open(self.url + "/api/sklad/excel", timeout=10) as r:
            self.assertIn("Sklad_taqsimot_", r.headers["Content-Disposition"])
            self.assertTrue(load_workbook(io.BytesIO(r.read())).sheetnames)


if __name__ == "__main__":
    unittest.main()
