# -*- coding: utf-8 -*-
"""Avto-aniqlash tuzatishlari (dizayn/AVTO-ANIQLASH.md): aniqlash qayta shakllantirishdan oldin, yolg'iz beruvchidagi kamayish
holatni o'zgartirmaydi, «Berdim» paytidagi oluvchi qoldig'i va 3 kunlik oyna, har oluvchi o'zi bo'yicha, yig'ma ogohlantirish."""
import datetime
import os
import sys
import threading
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fixtures import B, ADMIN, CH, UCH, EK, F_CH, TODAY  # noqa: E402
import test_avto_aniqlash_bot as T  # noqa: E402  (Base — testsiz yordamchi sinf)
import avto_aniqlash as AA  # noqa: E402

at = T.at
ADMIN2 = 9002


class Attribute(unittest.TestCase):
    def test_bitta_topshiriq_decide_bilan_bir_xil(self):
        for plan in (1, 3, 8):
            for down in (-2, 0, 1, 2, 3, 5, 6, 7, 8, 10):
                for up in (-1, 0, 1, 2, 3, 5, 6, 7, 8, 12):
                    st, q = AA.decide(plan, down, 0, 0, up)
                    st2, q2, _ = AA.attribute([{"plan": plan, "down": down, "up": up}])[0]
                    self.assertEqual((st, round(q, 6)), (st2, round(q2, 6)), (plan, down, up))

    def test_berildi_faqat_ishora(self):
        self.assertEqual(AA.HINTS, ("berildi",))
        self.assertEqual(AA.attribute([{"plan": 2, "down": 2, "up": 0}])[0][0], "berildi")

    def test_har_oluvchi_ozi(self):
        r = AA.attribute([{"plan": 1, "down": 1, "up": 1}, {"plan": 1, "down": 1, "up": 0}])
        self.assertEqual([x[0] for x in r], ["olindi", None], "qo'shni ko'paymadi — unga tegilmaydi")
        r = AA.attribute([{"plan": 6, "down": 6, "up": 6}, {"plan": 2, "down": 6, "up": 2}])
        self.assertEqual([x[0] for x in r], ["olindi", "tekshiring"], "−6 faqat kattasiga yetadi")
        r = AA.attribute([{"plan": 6, "down": 8, "up": 6}, {"plan": 2, "down": 8, "up": 2}])
        self.assertEqual([(x[0], x[1], x[2]) for x in r], [("olindi", 6, 6), ("olindi", 2, 2)])


class AvvalAniqlash(T.Base):
    """1: qayta shakllantirish ko'chirilgan qatorni bekor qilishidan oldin aniqlanadi; reja kamaytirilmaydi."""

    def setUp(self):
        super().setUp()
        self.stock("ОДДИЙ ЧАР", ch=20, uch=0)  # ЧИНОЗ → УРГУТ, reja 8
        self.stock("АВИТЕН", ek=10, ch=0)      # ЭРКИН → ЧИНОЗ, reja 3
        self.bot.data.when = lambda: at(8, 30)
        self.clock.at(8, 31)
        self.bot.koch_build()

    def later(self, h):
        t = at(h)
        self.bot.data.when = lambda: t
        self.clock.at(h, 1)

    def excel(self, drop=(), qty=None):
        rows = []
        for r in self.bot.data.koch:
            if r["name"] in drop:
                continue
            r = dict(r)
            if qty and r["name"] in qty:
                r["qty"] = qty[r["name"]]
            rows.append(r)
        self.bot.data.koch = rows

    def test_sinxron_yangilash(self):
        self.stock("ОДДИЙ ЧАР", ch=12, uch=8)  # ko'chirildi, tugma bosilmadi → makros ro'yxatidan chiqdi
        self.stock("АВИТЕН", ek=9, ch=1)       # qisman (−1 / +1 < 80%): makros rejani 2 ga tushirdi
        self.excel(drop=("ОДДИЙ ЧАР",), qty={"АВИТЕН": 2})
        self.later(11)
        self.assertTrue(self.bot.koch_sync_due())
        self.assertIsNotNone(self.bot.koch_try_snapshot())
        t = self.task("ОДДИЙ ЧАР")
        self.assertEqual((t["status"], t["auto_by"], t["qty_received"]), ("olindi", "tizim", 8), "bekor emas")
        av = self.task("АВИТЕН")
        self.assertEqual((av["status"], av["qty_plan"]), ("yangi", 3), "ko'chirish izi bor — reja kamaytirilmadi")
        self.excel(drop=("АВИТЕН",))
        self.later(12)
        self.bot.koch_try_snapshot()
        self.assertEqual(self.task("АВИТЕН")["status"], "yangi", "ro'yxatdan chiqdi, lekin bekor emas — aniqlashga qoladi")

    def test_izsiz_qator_baribir_bekor(self):
        self.excel(drop=("ОДДИЙ ЧАР",), qty={"АВИТЕН": 2})
        self.later(11)
        self.bot.koch_try_snapshot()
        self.assertEqual(self.task("ОДДИЙ ЧАР")["status"], "bekor")
        self.assertEqual(self.task("АВИТЕН")["qty_plan"], 2)

    def test_tekshiring_belgili_reja_kamaytirilmaydi(self):
        self.stock("АВИТЕН", ch=3)  # faqat oluvchida ko'paydi → «tekshiring»
        self.bot.koch_auto_detect()
        self.excel(qty={"АВИТЕН": 1})
        self.later(11)
        self.bot.koch_try_snapshot()
        self.assertEqual(self.task("АВИТЕН")["qty_plan"], 3)

    def test_panel_qayta_shakllantirish(self):
        self.stock("ОДДИЙ ЧАР", ch=12, uch=8)
        self.excel(drop=("ОДДИЙ ЧАР",))
        self.later(11)
        self.bot.koch_regen_manual(by=(None, "admin (panel)"))
        self.assertEqual(self.task("ОДДИЙ ЧАР")["status"], "olindi")

    def test_ikki_oqim_bir_ogohlantirish(self):
        self.stock("АВИТЕН", ch=3)
        th = [threading.Thread(target=self.bot.koch_auto_detect) for _ in range(4)]
        for t in th:
            t.start()
        for t in th:
            t.join(10)
        self.assertEqual(len(self.admin_warns("Tekshiring")), 1)


class Sotuv(T.Base):
    """2: beruvchidagi oddiy sotuv «berildi (tizim)» qilmaydi."""

    def test_sotuv(self):
        self.add_task("ОДДИЙ ЧАР", CH, UCH, 2, 20, 0)
        self.stock("ОДДИЙ ЧАР", ch=18, uch=0)  # ЧИНОЗ da 2 ta sotildi
        n0 = len(self.tg.texts(ADMIN))
        for h in range(9, 14):
            self.bot.data.when = lambda h=h: at(h)
            self.clock.at(h, 1)
            self.bot.koch_auto_detect()
        t = self.task("ОДДИЙ ЧАР")
        self.assertEqual((t["status"], t["given_by"], t["auto_by"], t["auto_warn"]), ("yangi", None, None, None))
        self.assertEqual(len(self.tg.texts(ADMIN)), n0)


class Berdim(T.Base):
    """2: «Berdim» paytidagi oluvchi qoldig'i (dst_stock1) va 3 kunlik oyna."""

    def test_oluvchi_berdimdan_oldin_kopaygan(self):
        self.stock("ОДДИЙ ЧАР", ch=20, uch=0)
        self.bot.koch_build()
        self.stock("ОДДИЙ ЧАР", uch=7)  # «Berdim» dan oldin yetkazib beruvchidan keldi
        self.clock.at(9, 0)
        self.assertTrue(self.bot.koch_give(F_CH, self.bot.client(F_CH), self.task("ОДДИЙ ЧАР")["id"]))
        self.assertEqual(self.task("ОДДИЙ ЧАР")["dst_stock1"], 7)
        self.stock("ОДДИЙ ЧАР", ch=12)  # beruvchida −8, oluvchida «Berdim» dan beri o'zgarmadi
        self.assertEqual(self.bot.koch_auto_detect()["olindi"], 0)
        self.assertEqual(self.task("ОДДИЙ ЧАР")["status"], "berildi")
        self.stock("ОДДИЙ ЧАР", uch=15)
        self.assertEqual(self.bot.koch_auto_detect()["olindi"], 1)
        t = self.task("ОДДИЙ ЧАР")
        self.assertEqual((t["status"], t["qty_received"], t["auto_note"]), ("olindi", 8, "ЧИНОЗ −8 · УРГУТ +8"))

    def given(self, name, src, dst, plan, s0, d0, days):
        tid = self.add_task(name, src, dst, plan, s0, d0, status="berildi")
        g = (at(9) - datetime.timedelta(days=days)).isoformat(timespec="seconds")
        with self.bot.dblock:
            self.bot.db.execute("UPDATE koch_tasks SET sana=?, given_at=?, given_by=?, qty_given=? WHERE id=?",
                                (g[:10], g, F_CH, plan, tid))
            self.bot.db.commit()
        return tid

    def test_3_kunlik_oyna(self):
        old = self.given("ОДДИЙ 04", CH, EK, 4, 10, 0, days=4)
        new = self.given("ОДДИЙ ЧАР", CH, UCH, 8, 20, 0, days=2)
        self.stock("ОДДИЙ 04", ch=6, ek=4)
        self.stock("ОДДИЙ ЧАР", ch=12, uch=8)
        self.clock.at(9, 30)
        self.bot.koch_auto_detect()
        st = {r["id"]: r["status"] for r in self.bot.db.execute("SELECT id, status FROM koch_tasks")}
        self.assertEqual((st[old], st[new]), ("berildi", "olindi"), "«Berdim» dan 4 kun o'tgan — avto-yopilmaydi")


class HarOluvchi(T.Base):
    """3: bitta beruvchidan bir xil dori bir nechta oluvchiga."""

    def test_bittasi_kochdi(self):
        a = self.add_task("ОДДИЙ ЧАР", CH, UCH, 1, 20, 0)
        b = self.add_task("ОДДИЙ ЧАР", CH, EK, 1, 20, 0)
        self.stock("ОДДИЙ ЧАР", ch=19, uch=1, ek=0)
        res = self.bot.koch_auto_detect()
        self.assertEqual((res["olindi"], res["tekshiring"]), (1, 0))
        st = {r["id"]: (r["status"], r["auto_warn"]) for r in self.bot.db.execute("SELECT * FROM koch_tasks")}
        self.assertEqual(st[a], ("olindi", None))
        self.assertEqual(st[b], ("yangi", None), "oluvchisi ko'paymagan — ogohlantirilmaydi")
        self.assertFalse(self.admin_warns("Tekshiring"))

    def test_kamayish_faqat_kattasiga_yetadi(self):
        a = self.add_task("ОДДИЙ 04", CH, EK, 6, 20, 0)
        b = self.add_task("ОДДИЙ 04", CH, UCH, 2, 20, 0)
        self.stock("ОДДИЙ 04", ch=14, ek=6, uch=2)
        res = self.bot.koch_auto_detect()
        st = {r["id"]: (r["status"], r["auto_warn"]) for r in self.bot.db.execute("SELECT * FROM koch_tasks")}
        self.assertEqual((st[a], st[b]), (("olindi", None), ("yangi", "t")))
        self.assertEqual((res["olindi"], res["tekshiring"]), (1, 1))


class Yigma(T.Base):
    """5: bir yangilashdagi «Tekshiring» lar — har adminga bitta xabar."""

    def test_bitta_xabar_har_adminga(self):
        self.bot.admins = {ADMIN, ADMIN2}
        self.stock("ОДДИЙ ЧАР", ch=20, uch=0)
        self.stock("АВИТЕН", ek=10, ch=0)
        self.bot.koch_build()
        self.stock("ОДДИЙ ЧАР", uch=8)
        self.stock("АВИТЕН", ch=3)
        n1, n2 = len(self.tg.texts(ADMIN)), len(self.tg.texts(ADMIN2))
        self.assertEqual(self.bot.koch_auto_detect()["tekshiring"], 2)
        new1, new2 = self.tg.texts(ADMIN)[n1:], self.tg.texts(ADMIN2)[n2:]
        self.assertEqual((len(new1), len(new2)), (1, 1))
        self.assertIn("ko'chirish: <b>2</b> ta", new1[0])
        self.assertIn("ОДДИЙ ЧАР", new1[0])
        self.assertIn("АВИТЕН", new1[0])

    def test_15_tadan_kop(self):
        text = self.bot.koch_warn_digest(["dori %d" % i for i in range(1, 21)])
        self.assertIn("ko'chirish: <b>20</b> ta", text)
        self.assertIn("15. dori 15", text)
        self.assertNotIn("16. dori 16", text)
        self.assertIn("yana <b>5</b> ta", text)


if __name__ == "__main__":
    unittest.main()
