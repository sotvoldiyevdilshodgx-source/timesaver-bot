# -*- coding: utf-8 -*-
"""Yangi talablar uchun mantiq modullari: A/B/C tahlili, ko'chirishni avtomatik aniqlash, buyurtma so'rovlari."""
import os
import sqlite3
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import abc_tahlil as ABC  # noqa: E402
import avto_aniqlash as AV  # noqa: E402
import buyurtma as BY  # noqa: E402


class AbcTahlil(unittest.TestCase):
    def test_80_15_5(self):
        rows = [("a", 50), ("b", 20), ("c", 10), ("d", 8), ("e", 5), ("f", 4), ("g", 2), ("h", 1)]  # jami 100
        g = ABC.split(rows, lambda r: r[1])
        self.assertEqual([r[0] for r in g["A"]], ["a", "b", "c"], "0–80%")
        self.assertEqual([r[0] for r in g["B"]], ["d", "e", "f"], "80–95% (d 80 dan boshlanadi)")
        self.assertEqual([r[0] for r in g["C"]], ["g", "h"])
        s = ABC.summary(g, lambda r: r[1])
        self.assertEqual((s["A"]["n"], s["A"]["sum"], s["A"]["share"]), (3, 80, 80))
        self.assertEqual(s["B"]["share"] + s["C"]["share"], 20)

    def test_bitta_juda_qimmat(self):
        g = ABC.split([("x", 900), ("y", 50), ("z", 50)], lambda r: r[1])
        self.assertEqual([r[0] for r in g["A"]], ["x"], "chegarani kesib o'tsa ham A")
        self.assertEqual([r[0] for r in g["B"]], ["y"])
        self.assertEqual([r[0] for r in g["C"]], ["z"])

    def test_nol_va_bosh(self):
        self.assertEqual(ABC.split([], lambda r: r), {"A": [], "B": [], "C": []})
        g = ABC.split([("p", 0), ("q", 10)], lambda r: r[1])
        self.assertEqual([r[0] for r in g["A"]], ["q"])
        self.assertEqual([r[0] for r in g["C"]], ["p"], "pulsiz — C")
        self.assertEqual(ABC.classify([("p", 0), ("q", 10)], lambda r: r[1], lambda r: r[0]), {"q": "A", "p": "C"})


class AvtoAniqlash(unittest.TestCase):
    def test_holatlar(self):
        self.assertEqual(AV.decide(12, 36, 24, 4, 16), ("olindi", 12))
        self.assertEqual(AV.decide(12, 36, 24, 4, 4), ("berildi", 12))
        self.assertEqual(AV.decide(12, 36, 36, 4, 16), ("tekshiring", 12))
        self.assertEqual(AV.decide(12, 36, 35, 4, 5), (None, 0))

    def test_80_foiz_chegara(self):
        self.assertEqual(AV.decide(10, 20, 12, 0, 8)[0], "olindi", "8/10 = 80% — yetarli")
        self.assertEqual(AV.decide(10, 20, 13, 0, 8)[0], "tekshiring", "beruvchida 7 — kam")
        self.assertEqual(AV.decide(10, 20, 10, 0, 7)[0], "berildi", "oluvchida 7 — kam (sotib yuborgan bo'lishi mumkin)")

    def test_malumot_yoq(self):
        self.assertEqual(AV.decide(5, None, 1, 0, 5), (None, 0))
        self.assertEqual(AV.decide(0, 10, 0, 0, 10), (None, 0))

    def test_ulush(self):
        self.assertEqual(AV.share(15, {1: 10, 2: 20}), {1: 5.0, 2: 10.0})


class Buyurtma(unittest.TestCase):
    def setUp(self):
        self.db = sqlite3.connect(":memory:")
        self.db.row_factory = sqlite3.Row
        BY.init(self.db)
        self.items = [{"name": "Парацетамол 500мг №10", "producer": "X", "qty": 20},
                      {"name": "Линекс №16", "producer": "Y", "qty": 25},
                      {"name": "Цитрамон П №10", "producer": "Z", "qty": 30}]

    def test_toliq_oqim(self):
        rid = BY.create(self.db, 501, "АЛЬФА ФАРМ", "Анвар", self.items)
        self.assertEqual([r["id"] for r in BY.open_requests(self.db)], [rid])
        self.assertTrue(BY.decide_item(self.db, rid, 1, qty=10))
        self.assertFalse(BY.finish(self.db, rid, 9001, "Admin"), "hammasi hal bo'lmagan")
        self.assertTrue(BY.accept_all(self.db, rid))
        r, items = BY.get(self.db, rid)
        self.assertEqual([(i["qty_ok"], i["status"]) for i in items], [(10, "ozgardi"), (25, "qabul"), (30, "qabul")])
        self.assertEqual(BY.totals(items), (75, 65, 1))
        self.assertTrue(BY.finish(self.db, rid, 9001, "Admin"))
        self.assertFalse(BY.finish(self.db, rid, 9002, "Admin 2"), "ikkinchi admin hal qila olmaydi")
        self.assertFalse(BY.decide_item(self.db, rid, 2, qty=1), "yopilgan so'rov o'zgarmaydi")
        self.assertEqual(BY.get(self.db, rid)[0]["status"], "javob")
        self.assertEqual(len(BY.approved_rows(self.db)), 3)
        self.assertEqual([r["id"] for r in BY.mine(self.db, 501)], [rid])

    def test_rad(self):
        rid = BY.create(self.db, 501, "АЛЬФА ФАРМ", "Анвар", self.items)
        self.assertTrue(BY.decide_item(self.db, rid, 3, reject=True))
        self.assertTrue(BY.finish(self.db, rid, 9001, "Admin", reject=True))
        r, items = BY.get(self.db, rid)
        self.assertEqual(r["status"], "rad")
        self.assertTrue(all(i["status"] == "rad" for i in items))
        self.assertEqual(BY.approved_rows(self.db), [])

    def test_bosh_savat_va_nol(self):
        with self.assertRaises(ValueError):
            BY.create(self.db, 1, "F", "N", [{"name": "A", "qty": 0}])
        rid = BY.create(self.db, 1, "F", "N", [{"name": "A", "qty": 5}])
        BY.decide_item(self.db, rid, 1, qty=0)
        self.assertEqual(BY.get(self.db, rid)[1][0]["status"], "rad", "0 ta — rad")


if __name__ == "__main__":
    unittest.main()
