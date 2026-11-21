# -*- coding: utf-8 -*-
"""Hisobotlar > Srok Excel: «Firma tel» ustuni (Srok varag'ining D ustuni = FOM GOOD.EDIZM = Лист1 «Фирма Тел»)."""
import io
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fixtures import TAG_A, TAG_B, make_bot  # noqa: E402
from openpyxl import load_workbook  # noqa: E402


class SrokFirma(unittest.TestCase):
    def setUp(self):
        self.bot, self.tg, self.work = make_bot("new")

    def test_firma_tel_ustuni(self):
        res = self.bot.prepare("srok", None)
        self.assertIsNotNone(res)
        ws = load_workbook(io.BytesIO(res[0])).active
        head = [c.value for c in ws[1]]
        self.assertEqual(head, ["Филиал", "Товар", "Ишлаб чиқарувчи", "Фирма тел", "Срок", "Қолган кун", "Қолдиқ"])
        firms = {r[3] for r in ws.iter_rows(min_row=2, values_only=True)}
        self.assertTrue(firms & {TAG_A, TAG_B}, firms)
        self.assertTrue(hasattr(ws.cell(2, 5).value, "year"), "Srok ustuni sana bo'lib qoladi")


if __name__ == "__main__":
    unittest.main()
