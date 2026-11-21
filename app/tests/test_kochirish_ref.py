# -*- coding: utf-8 -*-
"""B1: Ko'chirish qoidalari (tools/kochirish_ref.py). Miqdor tuzatishi (TUZATISH_kochirish_miqdor_prompt.md):
beruvchi hech qachon 0 da qolmaydi, oluvchiga faqat kerakli miqdor (KOCH_COVER = 1 davr sotuvi).
Har holatda ikkala qonun alohida tekshiriladi. python -m unittest tests.test_kochirish_ref"""
import collections
import datetime
import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
import kochirish_ref as K  # noqa: E402

CH, UCH, EK = "ДОРИХОНА ЧИНОЗ", "ДОРИХОНА УРГУТ", "ДОРИХОНА ЭРКИН"
FIL = [CH, UCH, EK]
A, B, C = CH, UCH, EK          # alifbo bo'yicha: УРГУТ (B) < ЧИНОЗ (A) < ЭРКИН (C)
TODAY = datetime.date(2026, 10, 29)
NEUTRAL = (1, 0)               # na oluvchi, na beruvchi (q=1 — tegilmaydi; sotuv 0 — ehtiyoj yo'q)


def run(q, s, k=0, srok=(), old=(), price=10000, davr=10, srok_fil=None, log=None):
    """q, s: {filial: qiymat}. srok: [(filial, qolgan_kun, qoldiq)]. old: [(filial, necha_kun)]."""
    sale = [{"name": "АВИТЕН", "producer": "РОССИЯ", "q": [q.get(f, 0) for f in FIL], "s": [s.get(f, 0) for f in FIL], "k": k}]
    sr = [{"filial": srok_fil or f, "name": "АВИТЕН", "producer": "РОССИЯ", "date": TODAY + datetime.timedelta(days=d),
           "kun": d, "ost": o, "prih": price} for f, d, o in srok]
    od = [{"filial": f, "name": "АВИТЕН", "producer": "РОССИЯ", "kun": n} for f, n in old]
    narx = [{"name": "АВИТЕН", "producer": "РОССИЯ", "prih": price}]
    return K.build(sale, FIL, davr, sr, od, narx, today=TODAY, log=log)


def case(a, b, c=NEUTRAL, **kw):
    """a, b, c: (qoldiq, sotuv) — A, B, C filiallari uchun."""
    q = {A: a[0], B: b[0], C: c[0]}
    s = {A: a[1], B: b[1], C: c[1]}
    return run(q, s, **kw), q, s


class Qonunlar(unittest.TestCase):
    def laws(self, koch, q, s):
        """1-qonun: beruvchida kamida MIN_QOLDIQ qoladi. 2-qonun: oluvchiga ehtiyojidan ko'p emas."""
        given, got = collections.Counter(), collections.Counter()
        for r in koch:
            given[r["src"]] += r["qty"]
            got[r["dst"]] += r["qty"]
        for f, n in given.items():
            self.assertGreaterEqual(int(q[f]) - n, K.MIN_QOLDIQ, "1-qonun: %s da %s − %s" % (f, q[f], n))
        for f, n in got.items():
            sv = max(0, s[f])
            need = K.ASSORT_SONI if int(q[f]) == 0 and sv == 0 else max(0, K.vceil(sv * K.KOCH_COVER) - int(q[f]))
            self.assertLessEqual(n, need, "2-qonun: %s ga %s (ehtiyoj %s)" % (f, n, need))

    def check(self, a, b, c=NEUTRAL, expect=(), left=None, **kw):
        (koch, _, _), q, s = case(a, b, c, **kw)
        self.laws(koch, q, s)
        self.assertEqual([(r["src"], r["dst"], r["qty"]) for r in koch], list(expect))
        if left:
            for f, n in left.items():
                self.assertEqual(q[f] - sum(r["qty"] for r in koch if r["src"] == f), n, "%s da qoladi" % f)
        return koch

    def test_01_siprofloksatsin(self):
        r = self.check((2, 0), (0, 1), expect=[(A, B, 1)], left={A: 1})
        self.assertEqual(r[0]["turi"], "SHOSHILINCH")

    def test_02_bittasi_bor_tegilmaydi(self):
        self.check((1, 0), (0, 3), expect=[])

    def test_03_bint(self):
        self.check((300, 0), (0, 50), expect=[(A, B, 50)], left={A: 250})

    def test_04_bisnova(self):
        self.check((3, 0), (0, 1), expect=[(A, B, 1)], left={A: 2})

    def test_05(self):
        self.check((5, 0), (0, 2), expect=[(A, B, 2)], left={A: 3})

    def test_06_beruvchi_ham_sotadi(self):
        self.check((10, 4), (0, 3), expect=[(A, B, 3)], left={A: 7})

    def test_07_beruvchiga_ozi_yetadi(self):
        self.check((5, 5), (0, 2), expect=[])

    def test_08_assortiment_istisno(self):
        r = self.check((6, 6), (0, 0), expect=[(A, B, 1)], left={A: 5})
        self.assertEqual(r[0]["turi"], "ASSORTIMENT")

    def test_09_assortiment_bittasi_bor(self):
        self.check((1, 3), (0, 0), expect=[])

    def test_10_oddiy(self):
        r = self.check((10, 0), (2, 3), expect=[(A, B, 1)], old=[(A, 45)], price=60000)
        self.assertEqual((r[0]["turi"], r[0]["sabab"]), ("ODDIY", "45 kun turibdi, oluvchida yetmaydi"))

    def test_11_ikki_oluvchi(self):
        self.check((300, 0), (0, 50), (0, 30), expect=[(A, B, 50), (A, C, 30)], left={A: 220})

    def test_12_beruvchida_kam_teng_ehtiyoj(self):
        self.check((3, 0), (0, 2), (0, 2), expect=[(A, B, 2)], left={A: 1})

    def test_13_kasr_sotuv(self):
        self.check((4, 0), (0, 0.3333), expect=[(A, B, 1)], left={A: 3})

    def test_14_manfiy_sotuv(self):
        r = self.check((2, -1), (0, 1), expect=[(A, B, 1)], left={A: 1})
        self.assertEqual(r[0]["src_sales"], 0, "beruvchi sotuvi −1 → 0")
        self.check((5, 0), (0, -2), expect=[])   # oluvchi sotuvi −2 → 0, hech qayerda sotuv yo'q


class Qoidalar(unittest.TestCase):
    def test_davr_miqdorga_tasir_qilmaydi(self):
        a = run({A: 300, C: 1}, {B: 50}, davr=10)[0]
        b = run({A: 300, C: 1}, {B: 50}, davr=30)[0]
        self.assertEqual([r["qty"] for r in a], [r["qty"] for r in b])
        self.assertEqual(a[0]["qty"], 50, "10 kunlik davrda ham 50 (eski formulada 150)")

    def test_kasr_qoldiq_berilmaydi(self):
        self.assertEqual(run({A: 1.9, C: 1}, {B: 2})[0], [], "Int(1,9) = 1 — tegilmaydi")

    def test_muddati_otgan(self):
        koch, _, _ = run({A: 5, C: 1}, {B: 2}, srok=[(A, -3, 5)])
        self.assertFalse([x for x in koch if x["src"] == A])
        koch, _, _ = run({A: 5, C: 1}, {B: 2}, srok=[(A, -3, 2), (A, 100, 3)])
        self.assertEqual([x["qty"] for x in koch if x["src"] == A], [2], "yaroqli 3 ta: 1 tasi qoladi, 2 tasi beriladi")

    def test_oddiy_kichik_summa(self):
        self.assertEqual(run({A: 10, B: 2, C: 1}, {B: 3}, old=[(A, 45)], price=10000)[0], [])

    def test_oddiy_emas_12_kun(self):
        self.assertEqual(run({A: 10, B: 2, C: 1}, {B: 3}, old=[(A, 12)], price=60000)[0], [])

    def test_sklad_ehtiyojcha(self):
        koch, sk, _ = run({A: 5, C: 1}, {B: 2}, k=40)
        self.assertEqual(koch, [], "skladda bor — filialdan ko'chirilmaydi")
        self.assertEqual([(x["dst"], x["qty"], x["turi"]) for x in sk], [(B, 2, "SHOSHILINCH")], "40 emas, ehtiyojcha 2")
        _, sk, _ = run({C: 1}, {B: 3}, k=1)
        self.assertEqual([x["qty"] for x in sk], [1], "sklad oxirgi qadog'ini ham beradi (MIN_QOLDIQ yo'q)")

    def test_beruvchilar_tartibi(self):
        koch, _, _ = run({A: 10, C: 4}, {B: 3})
        self.assertEqual([(x["src"], x["qty"]) for x in koch], [(A, 3)], "ko'p bera oladigani birinchi")
        koch, _, _ = run({A: 10, C: 10}, {B: 2}, old=[(C, 60), (A, 40)])
        self.assertEqual(koch[0]["src"], C, "teng bo'lsa — ko'proq turgani")
        koch, _, _ = run({A: 3, C: 3}, {B: 5})
        self.assertEqual(sorted((x["src"], x["qty"]) for x in koch), [(A, 2), (C, 2)], "ikkalasida 1 tadan qoladi")

    def test_bitta_beruvchi_hamma_toifaga(self):
        # A: 10 ta (srok 10 kun) · B: SHOSHILINCH 2 · C: ODDIY ehtiyoj 3 − 2 = 1
        koch, _, msg = run({A: 10, C: 2}, {B: 2, C: 3}, old=[(A, 45)], srok=[(A, 10, 10)], price=60000)
        self.assertEqual([(x["dst"], x["turi"], x["qty"]) for x in koch], [(B, "SHOSHILINCH", 2), (C, "ODDIY", 1)])
        self.assertTrue(msg.startswith("SHOSHILINCH: 1 · ASSORTIMENT: 0 · ODDIY: 1 · jami"), msg)
        self.assertTrue(msg.endswith("Skladdan: 0"), msg)

    def test_himoya_logi_bosh(self):
        log = []
        run({A: 300, B: 0, C: 0}, {B: 50, C: 30}, log=log)
        run({A: 3}, {B: 2, C: 2}, log=log)
        self.assertEqual(log, [], "to'g'ri mantiqda himoya hech qachon ishlamaydi")

    def test_min_qoldiq_1_dan_kam_bolsa_1(self):
        old = K.MIN_QOLDIQ
        try:
            K.MIN_QOLDIQ = 0
            self.assertEqual([x["qty"] for x in run({A: 2, C: 1}, {B: 5})[0]], [1])
        finally:
            K.MIN_QOLDIQ = old

    def test_turi_va_sabab(self):
        for koch in (run({A: 10, C: 1}, {B: 3})[0], run({A: 6}, {A: 6})[0],
                     run({A: 10, B: 2, C: 1}, {B: 3}, old=[(A, 45)], price=60000)[0]):
            for r in koch:
                self.assertIn(r["turi"], ("SHOSHILINCH", "ASSORTIMENT", "ODDIY"))
                self.assertRegex(r["sabab"], r"^(Tugagan, sotuv bor|Filialda yo'q, boshqa filialda sotilyapti|\d+ kun turibdi, oluvchida yetmaydi)$")

    def test_norm(self):
        self.assertEqual(K.norm("ДОРИХОНА  ЭРКИН"), K.norm("дорихона ЭРКИН "))
        self.assertEqual(K.norm("O’ZBEK"), K.norm("O'ZBEK"))
        self.assertEqual(K.norm("OʼZBEK"), K.norm("o`zbek"))

    def test_srok_2_probelli_filial(self):
        koch, _, _ = run({A: 5, C: 1}, {B: 2}, srok=[(A, 5, 5)], srok_fil="ДОРИХОНА  ЧИНОЗ")
        self.assertEqual([x["qty"] for x in koch if x["src"] == A], [2])

    def test_vceil(self):
        self.assertEqual([K.vceil(x) for x in (0, 0.3333, 1, 1.0000001, 2.5)], [0, 1, 1, 1, 3])


if __name__ == "__main__":
    unittest.main()
