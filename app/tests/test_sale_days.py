# -*- coding: utf-8 -*-
"""Sotuv davri (kun) — panelda belgilanadi va «Обновить» da refresh.vbs ga uzatiladi (DateBegin = bugun - kun)."""
import os
import sys
import time
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fixtures import B, BotCase  # noqa: E402
import timesaver_panel as P  # noqa: E402

OK_STATUS = "START\nOK Sale|5\nOK Srok|4\nOK 30+ kun|0\nOK Narx|3\nOK Hisob|5\nOK Ko'chirish|27\nDONE\n"


class SotuvDavri(BotCase):
    def setUp(self):
        super().setUp()
        self.status = os.path.join(self.work, "refresh_status.txt")
        self.calls = []
        self.api = P.Api(self.bot, self.bot.cfg)

    def runner(self, vbs, xlsm, timeout, days=None):
        self.calls.append(days)
        with open(self.status, "w", encoding="utf-8") as f:
            f.write(OK_STATUS)
        return 0, ""

    def run_refresh(self):
        self.bot.start_refresh()
        for _ in range(250):
            if not self.bot.refreshing:
                return
            time.sleep(0.02)
        self.fail("yangilash tugamadi")

    def test_sozlama(self):
        s = self.api.settings_get()
        self.assertEqual(s["sale_days"], 0)
        self.assertEqual(s["period_days"], self.bot.data.period_days)
        r = self.api.settings_set({"sale_days": 45})
        self.assertTrue(r["ok"], r)
        self.assertIn("45 kun", r["msg"])
        self.assertEqual(self.bot.sale_days, 45)
        self.assertFalse(self.api.settings_set({"sale_days": 400})["ok"])
        self.assertFalse(self.api.settings_set({"sale_days": "abc"})["ok"])
        self.assertTrue(self.api.settings_set({"sale_days": 0})["ok"])
        self.assertEqual(self.bot.sale_days, 0)

    def test_runner_ga_uzatiladi(self):
        self.runner.__func__.takes_days = True
        self.bot.runner = self.runner
        self.bot.set_setting("sale_days", 60)
        self.run_refresh()
        self.assertEqual(self.calls, [60])

    def test_nol_bolsa_uzatilmaydi(self):
        calls = []
        self.bot.runner = lambda vbs, xlsm, timeout: (calls.append(1), open(self.status, "w").write(OK_STATUS), (0, ""))[2]
        self.run_refresh()
        self.assertEqual(calls, [1], "eski runner (3 argument) ishlayveradi")

    def test_default_runner_argumentlari(self):
        with mock.patch.object(B.subprocess, "run") as run:
            run.return_value = mock.Mock(returncode=0, stdout=b"", stderr=b"")
            B.default_runner("r.vbs", "b.xlsm", 10, days=30)
            self.assertEqual(run.call_args[0][0][-2:], ["b.xlsm", "30"])
            B.default_runner("r.vbs", "b.xlsm", 10)
            self.assertEqual(run.call_args[0][0][-1], "b.xlsm")


if __name__ == "__main__":
    unittest.main()


class BuyurtmaEhtiyoj(BotCase):
    def test_standart_zaxira_sotuv_davri(self):
        api = P.Api(self.bot, self.bot.cfg)
        r = api.orders("")
        self.assertEqual(r["cover"], self.bot.data.period_days, "standart: zaxira = sotuv davri")
        period = max(1, self.bot.data.period_days)
        for s in r["suppliers"]:
            for it in s["items"]:
                exp = -(-(it["sales"] / period * period - it["stock"]) // 1)
                self.assertEqual(it["need"], int(exp), "buyurtma = sotuv − qoldiq (yuqoriga yaxlitlab)")
        self.assertEqual(api.orders("30")["cover"], 30, "tugma bilan boshqa zaxira tanlash mumkin")
