# -*- coding: utf-8 -*-
"""Sklad: bugun kelgan yuk (FOM kirim partiyalari) → filiallarga sotuv ulushi bo'yicha → FOM qoralama bron."""
import datetime
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fixtures import ADMIN, CH, EK, UCH  # noqa: E402
from test_fom_sql import D, FomCase  # noqa: E402
import fom_sql  # noqa: E402
import timesaver_bot as B  # noqa: E402

T = lambda n=0: (datetime.date.today() + datetime.timedelta(days=n)).isoformat() + " 00:00:00"  # noqa: E731


class Kirim(FomCase):
    def setUp(self):
        super().setUp()
        self.fom("INSERT INTO GOOD VALUES(50, 'КИРИМ ДОРИ', 'ГЕРМАНИЯ'), (51, 'КИРИМ ЯНГИ', 'ИНДИЯ')")
        self.fom("INSERT INTO ORG VALUES(3183, 'ГРАНД ФАРМ')")
        con_rows = [(200, 50, 6, D(100), T(0)), (201, 50, 4, D(300), T(0)),   # bugun: 10 ta (muddati yaqini 6 ta)
                    (202, 50, 9, D(300), T(-1)),                              # kecha keldi
                    (203, 51, 5, D(200), T(0)),                               # bugun, Sale da yo'q
                    (204, 50, 3, D(-2), T(0))]                                # muddati o'tgan
        for i, g, k, e, d in con_rows:
            self.fom("INSERT INTO INCOMELN(ID, GOOD, KOL, SALE, EXPDATA, STATE, PRICEROZ1, PRICEROZ1NDS, PRICEROZ1NONDS, DATA, ORG) "
                     "VALUES(?,?,?,0,?,3,1120,120,1000,?,3183)", i, g, k, e, d)
        self.bot.cfg.set("fom", "kirim_qadam", "1")   # eski testlar: 1 talab; 10 talab — test_qadam_10
        # sotuv: ЧИНОЗ 6, УРГУТ 3, ЭРКИН 1
        self.bot.data.sale.append({"name": "КИРИМ ДОРИ", "producer": "ГЕРМАНИЯ", "mnn": "", "firm": "", "v": [1, 6, 0, 3, 2, 1, 50],
                                   "stock": 53, "sales": 10})

    def items(self, sana=None):
        s = self.bot.kirim_summary(sana)
        self.assertTrue(s["ok"], s)
        return s, {i["name"]: i for i in s["items"]}

    def test_royxat_va_sotuv_ulushi(self):
        s, it = self.items()
        self.assertEqual(sorted(it), ["КИРИМ ДОРИ", "КИРИМ ЯНГИ"], "kechagi va muddati o'tgani kirmaydi")
        k = it["КИРИМ ДОРИ"]
        self.assertEqual((k["avail"], k["exp"], k["supplier"], k["in_sale"]), (10.0, D(100)[:10], "ГРАНД ФАРМ", True))
        # sotuv 6/3/1, qoldiq 1/0/2 → ehtiyoj 5/3/0; 10 ta keldi → 5/3/0, ЭРКИН da dori bor (assortiment kerak emas)
        self.assertEqual(k["need"], {"ДОРИХОНА ЧИНОЗ": 5, "ДОРИХОНА УРГУТ": 3, "ДОРИХОНА ЭРКИН": 0})
        self.assertEqual(k["alloc"], {"ДОРИХОНА ЧИНОЗ": 5, "ДОРИХОНА УРГУТ": 3, "ДОРИХОНА ЭРКИН": 0})
        self.assertEqual(k["stock"]["ДОРИХОНА ЧИНОЗ"], 1)
        self.assertEqual(it["КИРИМ ЯНГИ"]["alloc"], {"ДОРИХОНА ЧИНОЗ": 1, "ДОРИХОНА УРГУТ": 1, "ДОРИХОНА ЭРКИН": 1},
                         "yangi dori, hech qayerda yo'q — har aptekaga 1 dona assortiment, 2 tasi skladda")
        self.assertFalse(it["КИРИМ ЯНГИ"]["in_sale"])

    def test_kechagi_sana(self):
        s, it = self.items((datetime.date.today() - datetime.timedelta(days=1)).isoformat())
        self.assertEqual(it["КИРИМ ДОРИ"]["avail"], 9.0)

    def test_zaxira(self):
        self.bot.cfg.set("fom", "kirim_zaxira", "20")
        s, it = self.items()
        self.assertEqual(it["КИРИМ ДОРИ"]["alloc"], {"ДОРИХОНА ЧИНОЗ": 5, "ДОРИХОНА УРГУТ": 3, "ДОРИХОНА ЭРКИН": 0}, "8 ta = ehtiyoj")

    def test_qadam_10_va_kop_kelsa(self):
        self.bot.cfg.remove_option("fom", "kirim_qadam")   # standart: 10
        self.fom("UPDATE INCOMELN SET KOL = 40 WHERE ID = 201")   # bugun: 6 + 40 = 46 ta
        self.bot.data.sale[-1]["v"] = [3, 26, 0, 13, 50, 1, 0]   # ehtiyoj: 23 / 13 / 0 (ЭРКИН da 50 bor)
        s, it = self.items()
        self.assertEqual(s["qadam"], 10)
        self.assertEqual(it["КИРИМ ДОРИ"]["alloc"], {"ДОРИХОНА ЧИНОЗ": 20, "ДОРИХОНА УРГУТ": 10, "ДОРИХОНА ЭРКИН": 0},
                         "23 → 20, 13 → 10; qolgan 16 ta skladda")

    def test_qoida(self):
        A = B.Bot.kirim_alloc
        # foydalanuvchi misoli: 5 ta; ЧИНОЗ ehtiyoj 2, УРГУТ sotilmagan (qoldiq 0), ЭРКИН ehtiyoj 2 → 2 / 1 / 2
        self.assertEqual(A(5, [2, 0, 2], [0, 0, 1]), [2, 1, 2])
        self.assertEqual(A(5, [2, 0, 2], [0, 4, 1]), [2, 0, 2], "УРГУТ da dori bor — assortiment berilmaydi, 1 ta skladda")
        self.assertEqual(A(3, [4, 0, 2], [0, 0, 0]), [2, 0, 1], "yetmasa — ehtiyoj nisbatida")
        self.assertEqual(A(100, [23, 11, 7], [0, 0, 0]), [20, 10, 7], "≥10 → 10 ga yaxlitlash")
        self.assertEqual(A(30, [40, 20, 0], [5, 5, 0]), [20, 10, 0], "30 ta: 20/10 (ehtiyoj nisbatida)")
        self.assertEqual(A(0, [5, 5, 5], [0, 0, 0]), [0, 0, 0])
        self.assertEqual(A(4, [0, 0, 0], [0, 0, 3]), [1, 1, 0], "ehtiyoj yo'q: faqat assortiment")
        self.assertEqual(A(7, [2.3, 0, 0], [1, 1, 1], 1), [3, 0, 0], "ehtiyoj yuqoriga yaxlitlanadi")

    def test_ulush(self):
        self.assertEqual(B.Bot._share(10, [6, 3, 1]), [6, 3, 1])
        self.assertEqual(B.Bot._share(7, [1, 1, 1]), [3, 2, 2])
        self.assertEqual(B.Bot._share(5, [0, 0, 0]), [0, 0, 0])
        self.assertEqual(sum(B.Bot._share(13, [2.5, 7, 0.3])), 13)

    def test_bron_yaratish(self):
        s, it = self.items()
        k = [i for i, x in enumerate(s["items"]) if x["name"] == "КИРИМ ДОРИ"][0]
        made = self.bot.kirim_create(s["sana"], {CH: {k: 6}, UCH: {k: 3}, EK: {k: 1}}, s["sig"])
        self.assertEqual(len(made), 3)
        inv = self.q("SELECT ID, ORG, STATE, NUMBER, PURPOSE FROM INVOICE ORDER BY ID")
        self.assertEqual(sorted(r[1] for r in inv), [2, 7, 9])
        self.assertTrue(all(r[2] == 2 and r[3] == 0 and r[4].startswith("BOT kirim #") for r in inv))
        lines = sorted(self.q("SELECT i.ORG, l.INCOMELN, l.KOL FROM INVOICELN l JOIN INVOICE i ON i.ID = l.INVOICE"))
        self.assertEqual(lines, [(2, 200, 6.0), (7, 201, 3.0), (9, 201, 1.0)], "muddati yaqin partiya birinchi")
        self.assertEqual(self.q("SELECT SALE FROM INCOMELN WHERE ID IN (200, 201) ORDER BY ID"), [(6.0,), (4.0,)])
        b = self.local()
        self.assertEqual({(x["dst"], x["cats"], x["status"]) for x in b},
                         {(CH, "kirim", "qoralama"), (UCH, "kirim", "qoralama"), (EK, "kirim", "qoralama")})
        self.assertIn("Kirim taqsimoti", self.tg.texts(ADMIN)[-1])
        s2, it2 = self.items()
        self.assertNotIn("КИРИМ ДОРИ", it2, "tarqatilgani ro'yxatdan chiqdi (FOM'da SALE oshdi)")
        self.assertTrue(self.bot.fom_file(b[0]["id"]))
        self.assertTrue([m for m in self.bot.fom_summary()["made"] if m["kirim"]])

    def test_bitta_filial_keyin_qolgani(self):
        self.fom("UPDATE INCOMELN SET KOL = 40 WHERE ID = 201")   # 46 ta; ehtiyoj 5/3/0 (qadam 1)
        s, it = self.items()
        k = [i for i, x in enumerate(s["items"]) if x["name"] == "КИРИМ ДОРИ"][0]
        self.assertEqual(it["КИРИМ ДОРИ"]["alloc"], {CH: 5, UCH: 3, EK: 0})
        self.bot.kirim_create(s["sana"], {CH: {k: 5}}, s["sig"])          # faqat ЧИНОЗ
        s2, it2 = self.items()
        self.assertEqual(it2["КИРИМ ДОРИ"]["avail"], 41.0)
        self.assertEqual(it2["КИРИМ ДОРИ"]["given"][CH], 5)
        self.assertEqual(it2["КИРИМ ДОРИ"]["alloc"], {CH: 0, UCH: 3, EK: 0}, "ЧИНОЗ ga ikkinchi marta taklif qilinmaydi")
        b = self.local()[0]
        with self.bot.dblock:   # bekor qilingan bron hisobga olinmaydi
            self.bot.db.execute("UPDATE fom_boron SET status='bekor' WHERE id=?", (b["id"],))
            self.bot.db.commit()
        self.assertEqual(self.items()[1]["КИРИМ ДОРИ"]["given"][CH], 0)

    def test_yaxlitlash_qoldigi_qayta_taklif_qilinmaydi(self):
        self.bot.cfg.remove_option("fom", "kirim_qadam")   # 10
        self.fom("UPDATE INCOMELN SET KOL = 40 WHERE ID = 201")   # 46 ta
        self.bot.data.sale[-1]["v"] = [3, 30, 0, 13, 50, 1, 0]     # ehtiyoj 27 / 13 / 0
        s, it = self.items()
        k = [i for i, x in enumerate(s["items"]) if x["name"] == "КИРИМ ДОРИ"][0]
        self.assertEqual(it["КИРИМ ДОРИ"]["alloc"], {CH: 20, UCH: 10, EK: 0})
        self.bot.kirim_create(s["sana"], {CH: {k: 20}}, s["sig"])
        s2, it2 = self.items()
        self.assertEqual(it2["КИРИМ ДОРИ"]["alloc"], {CH: 0, UCH: 10, EK: 0}, "ЧИНОЗ ga qolgan 7 ta qayta taklif qilinmaydi")

    def test_kam_kelsa_ulushlar_barqaror(self):
        self.bot.data.sale[-1]["v"] = [0, 6, 0, 4, 0, 10, 0]   # ehtiyoj 6/4/10 = 20, keldi 10
        s, it = self.items()
        k = [i for i, x in enumerate(s["items"]) if x["name"] == "КИРИМ ДОРИ"][0]
        a1 = it["КИРИМ ДОРИ"]["alloc"]
        self.assertEqual(a1, {CH: 3, UCH: 2, EK: 5})
        self.bot.kirim_create(s["sana"], {EK: {k: 5}}, s["sig"])
        self.assertEqual(self.items()[1]["КИРИМ ДОРИ"]["alloc"], {CH: 3, UCH: 2, EK: 0}, "qolganlar ulushi o'zgarmaydi")

    def test_kelganidan_kop_mumkin_emas(self):
        s, it = self.items()
        k = [i for i, x in enumerate(s["items"]) if x["name"] == "КИРИМ ДОРИ"][0]
        with self.assertRaises(fom_sql.FomError):
            self.bot.kirim_create(s["sana"], {CH: {k: 8}, UCH: {k: 3}}, s["sig"])
        self.assertEqual(self.q("SELECT COUNT(*) FROM INVOICE"), [(0,)])
        self.assertEqual(self.local(), [])

    def test_royxat_ozgarsa(self):
        s, it = self.items()
        self.fom("UPDATE INCOMELN SET SALE = 1 WHERE ID = 200")   # shu orada kassada sotildi
        with self.assertRaises(fom_sql.FomError) as e:
            self.bot.kirim_create(s["sana"], {CH: {0: 1}}, s["sig"])
        self.assertIn("o'zgargan", str(e.exception))
        self.assertEqual(self.local(), [])

    def test_fayl_rejimida_yopiq(self):
        self.bot.cfg.set("fom", "mode", "fayl")
        self.use()
        s = self.bot.kirim_summary()
        self.assertFalse(s["ok"])
        self.assertIn("mode = sql", s["error"])


if __name__ == "__main__":
    unittest.main()
