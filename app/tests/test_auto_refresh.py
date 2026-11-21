# -*- coding: utf-8 -*-
"""B9: avtomatik yangilash (soxta makros; haqiqiy Excel ishga tushmaydi), shundan keyin kunlik topshiriq."""
import datetime
import os
import sys
import time
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fixtures import B, BotCase, ADMIN, FIRM, F_CH, TODAY, msg
import timesaver_panel as P

OK_STATUS = "START\nOK Sale|5\nOK Srok|4\nOK 30+ kun|0\nOK Narx|3\nOK Hisob|5\nOK Ko'chirish|27\nDONE\n"
FAIL_STATUS = "START\nFAIL [Microsoft][ODBC SQL Server Driver] SQL Server does not exist\n"


class AvtoYangilash(BotCase):
    def setUp(self):
        super().setUp()
        self.status = os.path.join(self.work, "refresh_status.txt")
        self.calls, self.ok = [], True
        self.bot.runner = self.fake_runner
        self.yesterday = datetime.datetime.combine(TODAY - datetime.timedelta(days=1), datetime.time(18, 0))
        self.when = self.yesterday
        self.bot.data.when = lambda: self.when
        self.bot.set_setting("auto_on", "1")

    def fake_runner(self, vbs, xlsm, timeout):
        self.calls.append(1)
        with open(self.status, "w", encoding="utf-8") as f:
            f.write(OK_STATUS if self.ok else FAIL_STATUS)
        if self.ok:  # makros Excelni yangiladi
            self.when = B.now()
        return (0, "") if self.ok else (1, "")

    def wait(self, attr="refreshing"):
        for _ in range(200):
            if not getattr(self.bot, attr):
                return
            time.sleep(0.02)
        self.fail("kutish tugamadi: " + attr)

    def test_vaqt_va_shartlar(self):
        self.assertEqual(self.bot.auto_refresh_time, datetime.time(8, 30))
        self.assertFalse(self.bot.auto_refresh_due(), "08:00 — hali emas")
        self.clock.at(9, 5)
        self.bot.set_setting("maint_on", "1")
        self.bot.run_scheduled()
        self.assertFalse(self.calls or self.bot.refreshing, "texnik rejimda ishlamaydi")
        self.bot.set_setting("maint_on", "0")
        self.bot.set_setting("auto_refresh_time", "")
        self.assertFalse(self.bot.auto_refresh_due(), "bo'sh — o'chiq")
        self.bot.set_setting("auto_refresh_time", "08:30")
        self.when = datetime.datetime.combine(TODAY, datetime.time(7, 50))
        self.assertFalse(self.bot.auto_refresh_due(), "bugun qo'lda yangilangan — qayta emas")

    def test_muvaffaqiyatli_va_topshiriq(self):
        self.clock.at(9, 5)
        self.bot.run_scheduled()
        self.assertTrue(self.bot.refreshing)
        self.bot.run_scheduled()
        self.assertFalse(self.bot.auto_running, "avto-yuborish yangilash tugashini kutadi")
        self.wait()
        for _ in range(200):  # topshiriq yangilash oqimida, refreshing=False dan keyin yaratiladi
            if self.tg.last(F_CH, "topshirig'i tayyor"):
                break
            time.sleep(0.02)
        self.assertEqual(len(self.calls), 1)
        self.assertEqual(self.bot.setting("last_refresh_ok_date"), TODAY.isoformat())
        self.assertFalse([t for t in self.tg.texts(ADMIN) if "Avto-yangilash bajarilmadi" in t])
        self.assertTrue(self.bot.koch_has_tasks(), "yangilashdan keyin kunlik topshiriq shakllandi")
        self.assertTrue(self.tg.last(F_CH, "topshirig'i tayyor"), "filialga xabar")
        self.bot.run_scheduled()
        self.wait("auto_running")
        self.assertEqual(self.bot.setting("last_auto_date"), TODAY.isoformat(), "keyingi siklda avto-yuborish")
        self.bot.run_scheduled()
        self.assertEqual(len(self.calls), 1, "shu kuni qayta yangilanmaydi")

    def test_kun_davomida_obnovit_topshiriqni_yangilaydi(self):
        self.clock.at(9, 5)
        self.bot.run_scheduled()
        self.wait()
        for _ in range(200):  # topshiriq va sinxron belgisi yangilash oqimida yoziladi
            if self.bot.setting("koch_sync_stamp"):
                break
            time.sleep(0.02)
        s1 = self.bot.setting("koch_sync_stamp")
        self.assertTrue(s1)
        self.clock.at(14, 20)
        self.bot.data.koch = [r for r in self.bot.data.koch if r["name"] != "ШОШ ДОРИ 00"]
        self.bot.data.reload = lambda force=False: True  # test: varaq yuqorida o'zgartirildi
        self.assertTrue(self.bot.start_refresh(None))  # admin «Обновить» bosdi
        self.wait()
        for _ in range(200):
            if self.bot.setting("koch_sync_stamp") != s1:
                break
            time.sleep(0.02)
        self.assertGreater(self.bot.setting("koch_sync_stamp"), s1)
        st = self.bot.db.execute("SELECT status FROM koch_tasks WHERE name='ШОШ ДОРИ 00'").fetchone()[0]
        self.assertEqual(st, "bekor", "ro'yxatdan chiqqan dori topshiriqdan ham chiqdi")

    def test_muvaffaqiyatsiz(self):
        self.ok = False
        self.when = B.now() - datetime.timedelta(days=4)
        self.clock.at(9, 5)
        self.bot.run_scheduled()
        self.wait()
        t = [x for x in self.tg.texts(ADMIN) if "Avto-yangilash bajarilmadi" in x]
        self.assertTrue(t and "SQL Server does not exist" in t[0], "adminga sababi bilan")
        self.assertFalse(self.bot.koch_has_tasks(), "muvaffaqiyatsiz yangilashdan topshiriq yo'q")
        self.bot.run_scheduled()
        self.assertEqual(len(self.calls), 1, "o'sha kuni qayta urinmaydi")
        self.wait("auto_running")
        self.assertIn("to'xtatildi (ma'lumot eski)", self.bot.setting("last_auto") or "")
        self.assertFalse(self.tg.to(FIRM, "doc"))

    def test_sozlamalar(self):
        self.bot.state[ADMIN] = ("set_artime", None)
        self.bot.handle(msg(ADMIN, "7:45"))
        self.assertEqual(self.bot.auto_refresh_time, datetime.time(7, 45))
        self.bot.state[ADMIN] = ("set_artime", None)
        self.bot.handle(msg(ADMIN, "-"))
        self.assertIsNone(self.bot.auto_refresh_time)
        api = P.Api(self.bot, self.bot.cfg)
        self.assertTrue(api.settings_set({"auto_refresh_time": "08:30"})["ok"])
        self.assertEqual(api.settings_get()["auto_refresh_time"], "08:30")
        self.assertFalse(api.settings_set({"auto_refresh_time": "25:00"})["ok"])


if __name__ == "__main__":
    unittest.main()
