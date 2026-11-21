# -*- coding: utf-8 -*-
"""Ommaviy ko'chirish: ASSORTIMENT kunlik limiti, yo'nalish bo'yicha berish / olish, ro'yxat himoyasi, kun oxiri yopish."""
import datetime
import io
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fixtures import B, BotCase, CH, UCH, EK, F_CH, F_EK, F_UCH, TODAY, cb, inline_data
from openpyxl import load_workbook

YESTERDAY = (TODAY - datetime.timedelta(days=1)).isoformat()


class Case(BotCase):
    def setUp(self):
        super().setUp()
        self.bot.koch_build()
        self.c = self.bot.client(F_CH)

    def tasks(self, **kw):
        q = "SELECT * FROM koch_tasks" + (" WHERE " + " AND ".join("%s=?" % k for k in kw) if kw else "")
        return self.bot.db.execute(q, tuple(kw.values())).fetchall()

    def press(self, cid, data):
        self.bot.handle(cb(cid, data))

    def give_btn(self, cat="shoshilinch", dst=EK):
        """Ro'yxatdagi shu yo'nalishning ommaviy tugmasi."""
        i = self.bot.koch_route_dsts(CH, cat).index(dst)
        return [d for d in inline_data(self.bot.koch_list_text(self.c, cat)[1]) if d.startswith("kg:%s:%d:" % (cat, i))][0]


class Limit(Case):
    def build(self, lim):
        self.bot.db.execute("DELETE FROM koch_tasks")
        self.bot.db.commit()
        self.bot.set_setting("koch_assort_limit", lim)
        self.bot.koch_build()
        return sorted(t["name"] for t in self.tasks(cat="assortiment"))

    def test_eng_kop_sotilgani_qoladi(self):
        self.assertEqual(self.build(2), ["АССОРТ 03", "АССОРТ 07"], "beruvchida sotuvi 7 va 6 — qoldi")

    def test_nol_cheksiz(self):
        self.assertEqual(len(self.build(0)), 4)

    def test_faqat_assortiment_cheklanadi(self):
        self.build(1)
        self.assertEqual(len(self.tasks(cat="shoshilinch")), 9)
        self.assertEqual(len(self.tasks(cat="oddiy")), 6)

    def test_navbat_royxatda_va_excelda(self):
        self.build(2)
        text = self.bot.koch_list_text(self.c, "assortiment")[0]
        self.assertIn("Yana 2 ta dori navbatda", text)
        self.assertIn("Jami: <b>2</b> ta dori", text)
        self.bot.filial_excel(F_CH, self.c, "assortiment")
        wb = load_workbook(io.BytesIO(self.tg.to(F_CH, "doc")[-1][3]))
        why = [r[-1] for r in wb["Ko'chirilmaydi"].iter_rows(min_row=2, values_only=True)]
        self.assertEqual(why.count(self.bot.WHY_QUEUE), 2)

    def test_bajarilgani_limitga_kiradi_boshqasi_tushmaydi(self):
        self.build(2)
        self.assertTrue(self.bot.koch_give(F_CH, self.c, self.tasks(name="АССОРТ 03")[0]["id"]))
        self.bot.koch_build(regen=True)
        active = sorted(t["name"] for t in self.tasks(cat="assortiment") if t["status"] != "bekor")
        self.assertEqual(active, ["АССОРТ 03", "АССОРТ 07"], "berilgani limitda, yangi qo'shilmadi")

    def test_joy_boshatilsa_navbatdagi_kiradi(self):
        self.build(2)
        self.bot.data.koch = [r for r in self.bot.data.koch if r["name"] != "АССОРТ 07"]
        cnt = self.bot.koch_build(regen=True)
        self.assertEqual((cnt["new"], cnt["cancelled"]), (1, 1))
        self.assertEqual(sorted(t["name"] for t in self.tasks(status="yangi", cat="assortiment")), ["АССОРТ 03", "АССОРТ 06"],
                         "navbatda eng ko'p sotilgani (06 — sotuv 4, 04 — sotuv 3)")

    def test_regen_hozirgi_royxatni_almashtirmaydi(self):
        self.build(2)
        for r in self.bot.data.koch:
            if r["name"] == "АССОРТ 04":
                r["src_sales"] = 99  # sotuvi oshdi, lekin kunning o'rtasida ro'yxat almashmaydi
        self.bot.koch_build(regen=True)
        self.assertEqual(sorted(t["name"] for t in self.tasks(status="yangi", cat="assortiment")), ["АССОРТ 03", "АССОРТ 07"])

    def test_limit_kamaysa_regen_qisqartiradi(self):
        self.build(0)
        self.bot.set_setting("koch_assort_limit", 1)
        self.assertEqual(self.bot.koch_build(regen=True)["cancelled"], 3)


class Give(Case):
    def test_royxatda_yonalish_tugmalari(self):
        text, kb = self.bot.koch_list_text(self.c, "shoshilinch")
        self.assertIn("Yo'nalishlar:", text)
        self.assertEqual(len([d for d in inline_data(kb) if d.startswith("kg:")]), 2)

    def test_tasdiq_soralmasdan_berilmaydi(self):
        self.press(F_CH, self.give_btn())
        self.assertIn("Ommaviy berish", self.tg.texts(F_CH)[-1])
        self.assertEqual(self.tasks(status="berildi"), [])

    def test_hammasi_berildi_bitta_xabar(self):
        n = len(self.tasks(src=CH, dst=EK, cat="shoshilinch", status="yangi"))
        self.assertGreater(n, 1)
        self.press(F_CH, self.give_btn().replace("kg:", "kgy:", 1))
        rows = self.tasks(src=CH, dst=EK, cat="shoshilinch")
        self.assertEqual({t["status"] for t in rows}, {"berildi"})
        self.assertTrue(all(t["qty_given"] == t["qty_plan"] and t["given_by"] == F_CH for t in rows))
        self.assertEqual(len(self.tg.to(F_EK)), 1, "oluvchiga bitta umumiy xabar")
        self.assertIn("<b>%d</b> ta dori" % n, self.tg.texts(F_EK)[0])
        ev = self.bot.db.execute("SELECT COUNT(*) FROM koch_events WHERE action='berildi' AND note='ommaviy'").fetchone()[0]
        self.assertEqual(ev, n, "har bir dori tarixda bor")
        self.assertEqual(self.tasks(src=CH, dst=UCH, status="berildi"), [], "boshqa yo'nalishga tegmadi")

    def test_ikki_marta_bosish(self):
        btn = self.give_btn().replace("kg:", "kgy:", 1)
        self.press(F_CH, btn)
        before = len(self.tg.to(F_EK))
        self.press(F_CH, btn)
        self.assertEqual(len(self.tg.to(F_EK)), before, "takroriy xabar ketmadi")

    def test_royxat_ozgarsa_belgilanmaydi(self):
        btn = self.give_btn().replace("kg:", "kgy:", 1)
        t = self.tasks(src=CH, dst=EK, cat="shoshilinch", status="yangi")[0]
        self.bot.data.koch = [dict(r, qty=r["qty"] + 1) if r["name"] == t["name"] else r for r in self.bot.data.koch]
        self.bot.koch_build(regen=True)
        self.press(F_CH, btn)
        self.assertEqual(self.tasks(status="berildi"), [])
        self.assertIn("Ro'yxat o'zgardi", self.tg.texts(F_CH)[-1])

    def test_farq_belgilangach_qolgani_beriladi(self):
        btn = self.give_btn().replace("kg:", "kgy:", 1)
        first = self.tasks(src=CH, dst=EK, cat="shoshilinch", status="yangi")[0]["id"]
        self.assertTrue(self.bot.koch_refuse(F_CH, self.c, first, "Tovar topilmadi"))
        self.press(F_CH, btn)
        self.assertEqual(self.tasks(status="berildi"), [], "eski tugma — ro'yxat o'zgargan")
        self.press(F_CH, self.give_btn().replace("kg:", "kgy:", 1))
        st = {t["id"]: t["status"] for t in self.tasks(src=CH, dst=EK, cat="shoshilinch")}
        self.assertEqual(st.pop(first), "rad")
        self.assertEqual(set(st.values()), {"berildi"})

    def test_boshqa_filial_bera_olmaydi(self):
        self.press(F_EK, self.give_btn().replace("kg:", "kgy:", 1))
        self.assertEqual(self.tasks(status="berildi"), [])

    def test_firma_va_begona_bosa_olmaydi(self):
        self.press(5002, self.give_btn().replace("kg:", "kgy:", 1))
        self.assertEqual(self.tasks(status="berildi"), [])

    def test_yonalish_filtri(self):
        i = self.bot.koch_route_dsts(CH, "shoshilinch").index(EK)
        text, kb = self.bot.koch_list_text(self.c, "shoshilinch", 0, i)
        n = len(self.tasks(src=CH, dst=EK, cat="shoshilinch", status="yangi"))
        self.assertIn("Jami: <b>%d</b> ta dori" % n, text)
        self.assertNotIn("→ %s" % UCH, text)
        data = inline_data(kb)
        self.assertTrue(any(d.startswith("kgx:") for d in data))
        self.assertTrue(all(d.endswith(":%d" % i) for d in data if d.startswith("ki:")), "kartochkadan shu ro'yxatga qaytadi")

    def test_yonalish_exceli(self):
        i = self.bot.koch_route_dsts(CH, "shoshilinch").index(EK)
        self.press(F_CH, "kgx:shoshilinch:%d" % i)
        wb = load_workbook(io.BytesIO(self.tg.to(F_CH, "doc")[-1][3]))
        got = {r[0] for r in wb.worksheets[0].iter_rows(min_row=2, values_only=True) if r[0] != "JAMI"}
        self.assertEqual(got, {EK})


class Receive(Case):
    def setUp(self):
        super().setUp()
        self.press(F_CH, self.give_btn().replace("kg:", "kgy:", 1))
        self.tg.sent.clear()
        self.ek = self.bot.client(F_EK)

    def btn(self):
        return [d for d in inline_data(self.bot.incoming_list(EK)[1]) if d.startswith("kv:")][0]

    def test_tasdiq_soralmasdan_olinmaydi(self):
        self.press(F_EK, self.btn())
        self.assertIn("Ommaviy qabul", self.tg.texts(F_EK)[-1])
        self.assertEqual(self.tasks(status="olindi"), [])

    def test_hammasi_olindi_bitta_xabar(self):
        n = len(self.tasks(status="berildi", dst=EK))
        self.press(F_EK, self.btn().replace("kv:", "kvy:", 1))
        rows = self.tasks(dst=EK, status="olindi")
        self.assertEqual(len(rows), n)
        self.assertTrue(all(t["qty_received"] == t["qty_given"] and t["received_by"] == F_EK for t in rows))
        self.assertEqual(len(self.tg.to(F_CH)), 1, "beruvchiga bitta xabar")

    def test_kam_kelgani_belgilangach(self):
        t = self.tasks(status="berildi", dst=EK)[0]
        btn = self.btn().replace("kv:", "kvy:", 1)
        self.assertTrue(self.bot.koch_receive(F_EK, self.ek, t["id"], max(1, t["qty_given"] - 1)))
        self.press(F_EK, btn)
        self.assertIn("Ro'yxat o'zgardi", self.tg.texts(F_EK)[-1])
        self.assertEqual(len(self.tasks(status="berildi", dst=EK)) + 1, len(self.tasks(dst=EK, cat="shoshilinch", src=CH)))

    def test_oluvchi_filtri(self):
        text, kb = self.bot.fil_in_screen(EK, 0, 0)  # panel ekrani: «◀ Orqaga» — hamma yo'nalishlarga (kl:0)
        self.assertIn(("kl:0"), inline_data(kb))
        self.assertTrue(any(d.startswith("kv:0:") for d in inline_data(kb)))

    def test_boshqa_filial_ololmaydi(self):
        self.press(F_UCH, self.btn().replace("kv:", "kvy:", 1))
        self.assertEqual(self.tasks(status="olindi"), [])


class Close(Case):
    def age(self, t, sana=YESTERDAY):
        self.bot.db.execute("UPDATE koch_tasks SET sana=?, created_at=? WHERE id=?", (sana, sana + "T08:30:00", t["id"]))
        self.bot.db.commit()

    def test_yangi_yopiladi_berildi_qoladi(self):
        a, b = self.tasks(src=CH, status="yangi")[:2]
        self.assertTrue(self.bot.koch_give(F_CH, self.c, b["id"]))
        self.age(a)
        self.age(b)
        self.assertEqual(self.bot.koch_close_old(), 1)
        self.assertEqual(self.tasks(id=a["id"])[0]["status"], "yopildi")
        self.assertEqual(self.tasks(id=b["id"])[0]["status"], "berildi")

    def test_bugungi_tegilmaydi(self):
        self.assertEqual(self.bot.koch_close_old(), 0)

    def test_kechikish_osmaydi(self):
        t = self.tasks(cat="shoshilinch", status="yangi")[0]
        self.age(t)
        self.bot.koch_close_old()
        t = self.tasks(id=t["id"])[0]
        i1 = self.bot.task_info(t, B.now())
        i2 = self.bot.task_info(t, B.now() + datetime.timedelta(days=30))
        self.assertEqual(i1["late_min"], i2["late_min"])
        self.assertEqual((i1["late_done"], i1["guilty_filial"], i1["late"]), (True, t["src"], False))

    def test_tarixda_bajarilmagan_korinadi(self):
        t = self.tasks(cat="shoshilinch", status="yangi")[0]
        self.age(t)
        self.bot.koch_close_old()
        h = self.bot.koch_history(7)
        self.assertEqual({r["filial"] for r in h["rank"]}, {t["src"]})

    def test_yopilgani_berib_bolmaydi(self):
        t = self.tasks(src=CH, status="yangi")[0]
        self.age(t)
        self.bot.koch_close_old()
        self.assertIsNone(self.bot.koch_give(F_CH, self.c, t["id"]))

    def test_yangi_kun_topshirigi_bilan_yopiladi(self):
        t = self.tasks(status="yangi")[0]
        self.age(t)
        self.bot.db.execute("DELETE FROM koch_tasks WHERE sana=?", (TODAY.isoformat(),))
        self.bot.db.commit()
        self.bot.set_setting("koch_snap_date", YESTERDAY)
        self.clock.at(23, 59)
        when = datetime.datetime.combine(TODAY, datetime.time(9, 0))
        self.bot.data.when = lambda: when
        self.assertTrue(self.bot.koch_try_snapshot())
        self.assertEqual(self.tasks(id=t["id"])[0]["status"], "yopildi")


if __name__ == "__main__":
    unittest.main()
