# -*- coding: utf-8 -*-
"""B2: Data — Ko'chirish varag'ini o'qish (yangi / avvalgi / eski format), toifa, yordamchilar."""
import collections
import datetime
import os
import shutil
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fixtures import B, CH, UCH, EK, EK2, TODAY, make_bot, close_bot


class Format(unittest.TestCase):
    def load(self, koch):
        bot, tg, work = make_bot(koch, extra_cfg=None)
        self.addCleanup(shutil.rmtree, work, True)
        close_bot(bot)
        return bot.data

    def test_yangi_format(self):
        d = self.load("new")
        self.assertEqual(len(d.koch), 27)
        self.assertTrue(d.koch_new)
        c = collections.Counter(r["cat"] for r in d.koch)
        self.assertEqual(c, {"shoshilinch": 11, "assortiment": 8, "oddiy": 7, None: 1})
        r = d.koch[0]
        for k in ("src", "dst", "name", "producer", "qty", "dst_sales", "dst_stock", "prih", "summa", "date", "turi", "sabab",
                  "src_sales", "src_stock", "firm", "cat", "days_idle"):
            self.assertIn(k, r)
        self.assertEqual(r["turi"], "SHOSHILINCH")
        self.assertEqual(sorted(x["days_idle"] for x in d.koch if x["cat"] == "oddiy"), [33, 45, 50, 60, 95, 120, 200])

    def test_avvalgi_makros_zudlik(self):
        d = self.load("zudlik")
        c = collections.Counter(r["cat"] for r in d.koch)
        self.assertEqual(c, {"shoshilinch": 11, "assortiment": 8, "oddiy": 7, None: 1}, "Zudlik + Sabab → to'g'ri toifa")

    def test_eski_format_K_N_yoq(self):
        d = self.load("old")
        self.assertEqual(len(d.koch), 27)
        self.assertFalse(d.koch_new)
        self.assertTrue(all(r["cat"] == "oddiy" and r["sabab"] == "" for r in d.koch), "orqaga moslik: hammasi ODDIY")

    def test_varaq_yoq(self):
        d = self.load(None)
        self.assertEqual((d.koch, d.koch_filials(), d.error), ([], [], None))


class Toifa(unittest.TestCase):
    def test_koch_category_of(self):
        f = B.koch_category_of
        self.assertEqual(f("SHOSHILINCH", ""), ("shoshilinch", 0))
        self.assertEqual(f("assortiment", ""), ("assortiment", 0))
        self.assertEqual(f("ODDIY", "45 kun turibdi, oluvchida yetmaydi"), ("oddiy", 45))
        self.assertEqual(f("", "Tugagan, sotuv bor"), ("shoshilinch", 0))
        self.assertEqual(f("Zudlik", "Filialda yo\u2019q, boshqa filialda sotilyapti"), ("assortiment", 0), "apostrof turi farqi")
        self.assertEqual(f(None, "12 kun turibdi, oluvchida yetmaydi"), ("oddiy", 12))
        self.assertEqual(f("Zudlik", ""), ("shoshilinch", 0), "orqaga moslik: ZUDLIK")
        self.assertEqual(f("Oddiy", ""), ("oddiy", 0))
        self.assertEqual(f("", ""), ("oddiy", 0))
        self.assertEqual(f("BOSHQA", "nimadir"), (None, 0))

    def test_norm(self):
        self.assertEqual(B.norm(EK2), B.norm(EK))
        self.assertEqual(B.norm("DOCTOR\u2019S"), B.norm("doctor's"))
        self.assertEqual(B.norm("O\u02bcZ"), B.norm("o`z"))


class Yordamchilar(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.bot, cls.tg, cls.work = make_bot("new")
        cls.d = cls.bot.data

    @classmethod
    def tearDownClass(cls):
        close_bot(cls.bot)
        shutil.rmtree(cls.work, ignore_errors=True)

    def test_koch_filials(self):
        self.assertEqual(self.d.koch_filials(), [UCH, CH, EK], "3 ta, «ДОРИХОНА  ЭРКИН» bitta (asl yozilish, probel birlashgan)")

    def test_koch_out_in(self):
        self.assertEqual(len(self.d.koch_out(EK)), 8, "2 probelli A ustuni ham topiladi")
        self.assertEqual(len(self.d.koch_out("дорихона  эркин", "oddiy")), 6)
        self.assertEqual(len(self.d.koch_out(CH, "assortiment")), 8)
        self.assertTrue(all(r["cat"] for r in self.d.koch_out(CH)), "toifasizlar ro'yxatga kirmaydi")
        self.assertEqual({B.norm(r["src"]) for r in self.d.koch_in(CH)}, {EK})
        self.assertEqual(self.d.koch_for, self.d.koch_out, "eski nom saqlangan")

    def test_koch_date_state(self):
        s = B.Data.koch_date_state
        self.assertEqual(s(None), "unknown")
        self.assertEqual(s(datetime.date(2300, 1, 1)), "unknown")
        self.assertEqual(s(TODAY - datetime.timedelta(days=1)), "expired")
        self.assertEqual(s(TODAY + datetime.timedelta(days=3)), "warn")
        self.assertEqual(s(TODAY + datetime.timedelta(days=45)), "ok")
        self.assertEqual(s(TODAY + datetime.timedelta(days=45), warn_days=60), "warn")

    def test_manfiy_sotuv_0(self):
        self.assertTrue(all(r["src_sales"] >= 0 and r["dst_sales"] >= 0 for r in self.d.koch))
        self.assertTrue(all(x >= 0 for r in self.d.sale for x in (r["v"][1], r["v"][3], r["v"][5])))
        self.assertEqual(self.d.sale[0]["v"][1], 0.0, "Лист1 dagi −1 sotuv → 0")


@unittest.skipUnless(os.path.exists(r"E:\sale3\baza_yangi\baza_yangi.xlsm"), "haqiqiy Excel yo'q")
class HaqiqiyFayl(unittest.TestCase):
    def test_toifalar_yigindisi(self):
        """§12-savol 5: shoshilinch + assortiment + oddiy + noma'lum = jami (haqiqiy fayl nusxasi bo'yicha)."""
        import configparser
        cfg = configparser.ConfigParser(interpolation=None)
        cfg.read_dict({"excel": {"file": r"E:\sale3\baza_yangi\baza_yangi.xlsm", "sheet_srok": "Srok", "sheet_old": "30+ kun",
                                 "sheet_narx": "Narx", "sheet_hisob": "Hisob", "sale_first_row": "8"}})
        d = B.Data(cfg)
        self.assertTrue(d.reload(), d.error)
        c = collections.Counter(r["cat"] or "noma'lum" for r in d.koch)
        print("\n    Haqiqiy fayl: %s · jami %d" % (dict(c), len(d.koch)))
        self.assertEqual(sum(c.values()), len(d.koch))
        self.assertEqual(c.get("noma'lum", 0), 0, [r["sabab"] for r in d.koch if not r["cat"]][:10])


if __name__ == "__main__":
    unittest.main()
