# -*- coding: utf-8 -*-
"""B6: eslatmalar — muddatdan oldin, muddat o'tganda (SHOSHILINCH / ASSORTIMENT), kun oxiri hisoboti, texnik rejim."""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fixtures import B, BotCase, CH, UCH, EK, ADMIN, F_CH, F_EK, F_UCH, cb, inline_data


class Eslatma(BotCase):
    def setUp(self):
        super().setUp()
        self.bot.koch_build()

    def tick(self, h, m=0):
        self.clock.at(h, m)
        if self.bot.koch_reminders_due():
            self.bot.koch_reminders()

    def tid(self, name):
        return self.bot.db.execute("SELECT id FROM koch_tasks WHERE name=?", (name,)).fetchone()[0]

    def test_muddatdan_oldin(self):
        self.tick(9, 29)
        self.assertFalse(self.tg.sent)
        self.tick(9, 30)
        t = self.tg.last(F_CH, "qoldi")
        self.assertIn("⏰ <b>30 daqiqa qoldi</b>: 11 ta ko'chirish bajarilmagan (SHOSHILINCH / ASSORTIMENT)", t[2],
                      "ODDIY kirmaydi: 7 + 4")
        self.assertIn("kp:shoshilinch:0", inline_data(t[3]))
        self.assertIn("2 ta ko'chirish bajarilmagan", self.tg.last(F_EK, "qoldi")[2])
        self.assertFalse(self.tg.to(F_UCH), "ochiq topshirig'i yo'q filialga xabar yo'q")
        self.assertFalse(self.tg.to(ADMIN))
        n = len(self.tg.sent)
        self.tick(9, 45)
        self.assertEqual(len(self.tg.sent), n, "kuniga bir marta")

    def test_muddat_otdi(self):
        self.tick(9, 30)
        self.bot.koch_give(F_EK, self.bot.client(F_EK), self.tid("АВИТЕН"))  # endi ЧИНОЗ «Oldim» bosishi kerak
        self.tick(10, 1)
        self.assertIn("⛔ <b>Muddat o'tdi</b>: 12 ta ko'chirish bajarilmagan", self.tg.last(F_CH, "Muddat o'tdi")[2],
                      "11 berish + 1 qabul qilish")
        self.assertIn("1 ta ko'chirish bajarilmagan", self.tg.last(F_EK, "Muddat o'tdi")[2])
        adm = self.tg.texts(ADMIN)
        self.assertEqual(len(adm), 1, "adminlarga bitta xabar")
        self.assertIn("⛔ <b>SHOSHILINCH / ASSORTIMENT muddati (10:00) o'tdi</b>", adm[0])
        self.assertIn("🔴 SHOSHILINCH: 9 ta", adm[0])
        self.assertRegex(adm[0], "🔴 %s · \\d+ / \\d+ / <b>12</b>" % CH, "aybdorlar jadvali")

    def test_toifa_muddatlari_alohida(self):
        self.bot.set_setting("koch_deadline_assortiment", "11:00")
        self.tick(9, 30)
        self.assertIn("7 ta ko'chirish bajarilmagan (SHOSHILINCH)", self.tg.last(F_CH, "qoldi")[2])
        self.tick(10, 30)
        self.assertIn("4 ta ko'chirish bajarilmagan (ASSORTIMENT)", self.tg.last(F_CH, "qoldi")[2])
        self.tick(11, 0)
        self.assertIn("⛔ <b>ASSORTIMENT muddati (11:00) o'tdi</b>", self.tg.texts(ADMIN)[-1])

    def test_kun_oxiri_hisoboti(self):
        self.bot.koch_refuse(F_CH, self.bot.client(F_CH), self.tid("ШОШ ДОРИ 01"), "Tovar topilmadi")
        for h in (9, 10, 17):
            self.tick(h, 59)
        n = len(self.tg.texts(ADMIN))
        self.tick(18, 0)
        rep = self.tg.texts(ADMIN)[n:]
        self.assertEqual(len(rep), 1)
        for s in ("📊 <b>Ko'chirish · kunlik hisobot</b>", "🔴 SHOSHILINCH: bajarilish <b>0%</b> (0/9)", "🟢 ODDIY: bajarilish",
                  "❌ <b>Rad</b>: 1 ta", "Tovar topilmadi", "⛔ <b>Kechikkan / kech bajarilgan</b>"):
            self.assertIn(s, rep[0])
        self.tick(19, 0)
        self.assertEqual(len(self.tg.texts(ADMIN)), n + 1)

    def test_hisobot_ochirilgan(self):
        self.bot.set_setting("alert_daily_report", "0")
        self.tick(18, 30)
        self.assertFalse([t for t in self.tg.texts(ADMIN) if "kunlik hisobot" in t])

    def test_topshiriq_yaratilmagan(self):
        self.bot.db.execute("DELETE FROM koch_tasks")
        self.bot.db.commit()
        self.tick(10, 0)
        self.assertIn("topshirig'i yaratilmadi", self.tg.texts(ADMIN)[-1])
        self.assertFalse(self.tg.to(F_CH))

    def test_ochirilgan_eslatmalar(self):
        self.bot.set_setting("koch_rem_before_on", "0")
        self.tick(9, 30)
        self.assertFalse(self.tg.sent)
        self.bot.set_setting("koch_on", "0")
        self.clock.at(10, 0)
        self.bot.koch_reminders()
        self.assertFalse(self.tg.sent)

    def test_texnik_rejimda_yuborilmaydi(self):
        self.bot.set_setting("maint_on", "1")
        called = []
        self.bot.koch_reminders = lambda: called.append(1)
        self.bot.bg = lambda fn, name=None: fn()
        self.clock.at(10, 0)
        self.bot.run_scheduled()
        self.assertFalse(called)
        self.bot.set_setting("maint_on", "0")
        self.bot.run_scheduled()
        self.assertTrue(called)

    def test_birinchi_kun(self):
        self.bot.set_setting("koch_installed", B.now().date().isoformat())
        self.bot.set_setting("koch_rem1_date", B.now().date().isoformat())
        self.clock.at(10, 0)
        self.assertFalse(self.bot.koch_reminders_due(), "o'rnatilgan kuni — ertangi kundan")

    def test_admin_kochirish_eslatma(self):
        self.clock.at(10, 30)
        self.bot.handle(cb(ADMIN, "a:tr"))
        data = inline_data(self.tg.sent[-1][3])
        self.assertIn("a:krm:all", data)
        self.bot.handle(cb(ADMIN, "a:krm:all"))
        self.assertIn("🔔 Eslatma yuborildi: 2 filial", self.tg.texts(ADMIN)[-1], "natija — panelning o'zida (izoh)")
        self.assertIn("🔔 <b>Admin eslatmasi: ko'chirish</b>", self.tg.last(F_CH, "Admin eslatmasi")[2])


if __name__ == "__main__":
    unittest.main()
