# -*- coding: utf-8 -*-
"""B8: admin ogohlantirishlari (yangi kalitlar va migratsiya), ichki hisoblar."""
import io
import os
import sqlite3
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fixtures import B, BotCase, ADMIN, FIRM, TAG_A, cb, inline_data, close_bot
import timesaver_panel as P
from openpyxl import load_workbook

KEYS = ["alert_expired", "alert_loss_price", "alert_bad_dates", "alert_internal", "alert_daily_report"]


class Ogohlantirish(BotCase):
    def test_kalitlar(self):
        self.assertEqual([k for k, _ in self.bot.ALERTS], KEYS)

    def test_hammasi_yoqilgan(self):
        parts, sheets = self.bot.admin_alerts()
        txt = "\n".join(parts)
        self.assertIn("Muddati o'tgan, qoldig'i bor</b>: <b>2</b> partiya · 50 000", txt)
        self.assertIn("Sotuv narxi ≤ kirim narxi</b>: <b>2</b> ta", txt)
        self.assertIn("kirim 2 500 · sotuv 2 400", txt)
        self.assertIn("<b>1</b> ta — bazada to'g'rilash kerak", txt)
        self.assertIn("ИНВЕНТАРИЗАЦИЯ", txt)
        self.assertNotIn("ГАММА", txt)
        self.assertEqual([s[0] for s in sheets], ["Muddati o'tgan", "Narx", "Xato muddat", "Ichki hisoblar"])

    def test_bitta_xabar_va_excel(self):
        self.bot.daily_admin()
        msgs, docs = self.tg.to(ADMIN, "send"), self.tg.to(ADMIN, "doc")
        self.assertEqual((len(msgs), len(docs)), (1, 1))
        self.assertLessEqual(len(msgs[0][2]), 4096)
        wb = load_workbook(io.BytesIO(docs[0][3]))
        self.assertEqual(wb["Muddati o'tgan"].max_row, 3)

    def test_ochirish_telegram(self):
        self.bot.handle(cb(ADMIN, "a:set:alerts"))
        data = inline_data(self.tg.sent[-1][3])
        for k in KEYS:
            self.assertIn("a:set:al:" + k, data)
        self.bot.handle(cb(ADMIN, "a:set:al:alert_loss_price"))
        self.assertFalse(self.bot.alert_on("alert_loss_price"))
        self.assertFalse(any("Sotuv narxi" in p for p in self.bot.admin_alerts()[0]))
        self.bot.handle(cb(ADMIN, "a:set:al:alert_bad_dates"))
        self.assertFalse(any("Xato muddatlar" in p for p in self.bot.admin_alerts()[0]))

    def test_hammasi_ochiq(self):
        for k in KEYS:
            self.bot.set_setting(k, "0")
        self.bot.daily_admin()
        self.assertFalse(self.tg.to(ADMIN))

    def test_panel(self):
        api = P.Api(self.bot, self.bot.cfg)
        s = api.settings_get()
        self.assertEqual([a["key"] for a in s["alerts"]], KEYS)
        self.assertTrue(api.settings_set({"alert_bad_dates": False})["ok"])
        self.assertFalse(self.bot.alert_on("alert_bad_dates"))
        self.assertTrue(api.settings_set({"alert_daily_report": False})["ok"])
        self.assertFalse(self.bot.alert_on("alert_daily_report"))

    def test_ichki_hisoblar(self):
        t = self.bot.hisob_text(None, None)
        self.assertIn("Tashkilotlar: <b>2</b>", t)
        self.assertIn("🏦 Ichki hisoblar", t)
        with self.bot.dblock:
            self.bot.db.execute("INSERT INTO client_accounts VALUES(?,?)", (FIRM, "ДОРИХОНА ЧИНОЗ"))
            self.bot.db.commit()
        rows, net = self.bot.hisob_net({TAG_A}, self.bot.accounts_of(FIRM), auto=False)
        self.assertTrue(all(r["org"] != "ДОРИХОНА ЧИНОЗ" for r in rows), "mijoz panelida ichki hisob ko'rinmaydi")
        self.bot.set_setting("hisob_internal", "")
        self.assertTrue(self.bot.is_internal("КУП ЧИККАН  ДОРИЛАР ЧИНОЗ"))
        self.assertFalse(self.bot.is_internal("КУП ЧИКАН ДОРИЛАР СКЛАД"))


class Migratsiya(BotCase):
    clients = False

    def test_eski_kalitlar_kochadi(self):
        path = self.bot.db_path
        close_bot(self.bot)
        con = sqlite3.connect(path)
        con.executemany("INSERT OR REPLACE INTO settings VALUES(?,?)", [("alert_price", "0"), ("alert_baddate", "0")])
        con.execute("DELETE FROM settings WHERE key IN ('alert_loss_price', 'alert_bad_dates')")
        con.commit()
        con.close()
        self.bot = B.Bot(self.bot.cfg, tg=self.tg, db_path=path)
        self.assertFalse(self.bot.alert_on("alert_loss_price"), "alert_price=0 → alert_loss_price=0")
        self.assertFalse(self.bot.alert_on("alert_bad_dates"))
        self.assertTrue(self.bot.alert_on("alert_expired"))
        self.assertEqual(self.bot.setting("alert_price"), "0", "eski kalit o'chirilmaydi")


if __name__ == "__main__":
    unittest.main()
