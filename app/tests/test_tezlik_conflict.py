# -*- coding: utf-8 -*-
"""2026-11-20: Sale varag'idagi takror qatorlar qo'shiladi; norm/fil_key keshi; «409 Conflict» (shu token bilan ikkinchi
bot nusxasi) aniqlanadi va panelda ko'rsatiladi."""
import os
import sys
import time
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fixtures import B, BotCase
import timesaver_panel as P


class TakrorSale(unittest.TestCase):
    def row(self, name, firm, v):
        return {"name": name, "producer": "ИНДИЯ", "mnn": "", "firm": firm, "v": list(v),
                "stock": v[0] + v[2] + v[4] + v[6], "sales": v[1] + v[3] + v[5]}

    def test_qoshiladi(self):
        rows = [self.row("ПЕКТОСЕН  ТАБ  N30", "900000001", [2.5, 0, 0, 0, 0, 0, 0]),
                self.row("ПЕКТОСЕН ТАБ N30", "", [4, 0, 0, 0, 0, 1.5, 0])]
        out = B.Data.merge_sale(rows)
        self.assertEqual(len(out), 1)
        self.assertEqual((out[0]["stock"], out[0]["sales"], out[0]["firm"]), (6.5, 1.5, "900000001"))
        self.assertEqual(out[0]["v"], [6.5, 0, 0, 0, 0, 1.5, 0])
        self.assertEqual(rows[0]["v"][0], 2.5, "kirish qatorlari o'zgarmaydi")

    def test_ikki_xil_firma_alohida(self):
        rows = [self.row("ДОРИ", "F1", [1, 0, 0, 0, 0, 0, 0]), self.row("ДОРИ", "F2", [2, 0, 0, 0, 0, 0, 0])]
        self.assertEqual(len(B.Data.merge_sale(rows)), 2)


class NormKesh(unittest.TestCase):
    def test_natija_ozgarmagan(self):
        self.assertEqual(B.norm("  дорихона  чиноз "), "ДОРИХОНА ЧИНОЗ")
        self.assertEqual(B.norm(None), "")
        self.assertEqual(B.norm(5), "5")
        self.assertEqual(B.norm("a’b"), "A'B")
        self.assertEqual(B.fil_key("ДОРИХОНА  ЧИНОЗ"), "ЧИНОЗ")
        self.assertEqual(B.fil_key("ДОРИХОНА ЧИНОЗ"), "ЧИНОЗ")


class Conflict(BotCase):
    def test_409_aniqlanadi(self):
        bot = self.bot
        bot.cfg["bot"]["token"] = "123:ABC"
        bot.run_scheduled = lambda: None
        calls = []

        def poll(offset, timeout=30):
            calls.append(1)
            if len(calls) == 1:
                raise B.TGError("HTTP 409: Conflict: terminated by other getUpdates request", 409)
            raise SystemExit

        bot.poll_once = poll
        with mock.patch.object(B.time, "sleep", lambda s: None), self.assertRaises(SystemExit):
            bot.run_forever()
        self.assertGreater(bot.conflict_at, time.time() - 5)
        self.assertTrue(P.Api(bot, bot.cfg).state()["conflict"])
        bot.conflict_at = time.time() - 600
        self.assertFalse(P.Api(bot, bot.cfg).state()["conflict"], "eski 409 — ogohlantirish yo'qoladi")


if __name__ == "__main__":
    unittest.main()
