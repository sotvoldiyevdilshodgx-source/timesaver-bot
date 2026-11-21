# -*- coding: utf-8 -*-
"""Yangi bo'limlar: (6) filial xodimi ⏰ Srok / 📦 30+ kun — A / B / C, 20 ta + qolgani Excel, filiallar ajratilgan;
(7) 📊 Sotuv va qoldiq — dori bloklari, saralash, sahifalash, firma ajratilgan; (8) ko'chirish A / B / C — belgilar,
tartib, bosh menyu qatori, «📋 Bugungi holat», koch_abc_summary."""
import io
import re
import unittest

from fixtures import (ADMIN, CH, EK, EK2, F_CH, F_EK, F_UCH, FIRM, STAFF, TAG_A, TAG_B, UCH, B, BotCase, add_client, cb,
                      day, inline_data)
from openpyxl import load_workbook

import abc_tahlil as ABC
from uzkir import strip_marks

SK_ID = 6010


def add_rows(bot, sheet, rows, first=5):
    """Test Exceliga qator qo'shadi (mavjudlaridan keyin) va qayta yuklaydi."""
    path = bot.data.find_path()
    wb = load_workbook(path)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    start = max(ws.max_row + 1, first)
    for i, r in enumerate(rows):
        for j, v in enumerate(r, 1):
            ws.cell(start + i, j, v)
    wb.save(path)
    bot.data.reload(force=True)


def buttons(markup):
    return [b for row in (markup or {}).get("inline_keyboard", []) for b in row]


class Press:
    def press(self, cid, data):
        """Tugma panel xabarida bosildi; oxirgi panel (matn, tugmalar) qaytadi."""
        self.bot.handle(cb(cid, data, self.bot.panel_mid(cid)))
        last = [s for s in self.tg.sent if s[1] == cid and s[0] in ("send", "edit")]
        return last[-1][2], last[-1][3]


# =====================================================================================================================
# 6) ⏰ Srok / 📦 30+ kun — A / B / C (filial xodimi)
# =====================================================================================================================
SROK_CH = [(CH, "СРОК %02d" % i, "ИНДИЯ", TAG_A, day(5 + 2 * i), 5 + 2 * i, 1, 20000 - 100 * i) for i in range(30)]


class SrokAbc(Press, BotCase):
    def setUp(self):
        super().setUp()
        add_rows(self.bot, "Srok", SROK_CH + [
            ("ЧИНОЗ", "ҚИСҚА НОМ", "ИНДИЯ", TAG_A, day(40), 40, 1, 15000),          # qisqa nom — o'sha filial
            (EK2, "ЭКС СРОК 1", "ИНДИЯ", TAG_B, day(20), 20, 3, 7000),                # 2 probel — EK
            (EK, "ЭКС СРОК 2", "ИНДИЯ", TAG_B, day(60), 60, 1, 3000),
            ("ДОРИХОНА ЧИНОЗ 2", "БЕГОНА ДОРИ", "ИНДИЯ", TAG_A, day(10), 10, 9, 99000),  # boshqa filial (nomi o'xshash)
            ("ДОРИХОНА СКЛАД", "СКЛАД ДОРИ", "ИНДИЯ", TAG_A, day(15), 15, 2, 5000)])
        add_rows(self.bot, "30+ kun", [
            (CH, "ЭСКИ 1", "ИНДИЯ", "ТАЪМИНОТЧИ", TAG_A, 45, 3, 1000, 50000),
            (CH, "ЭСКИ 2", "ИНДИЯ", "ТАЪМИНОТЧИ", TAG_A, 120, 2, 4000, 0),          # summa yo'q → 2 × 4000
            (CH, "ЭСКИ 3", "ИНДИЯ", "ТАЪМИНОТЧИ", TAG_A, 31, 1, 500, 700),
            (UCH, "УЧ ЭСКИ", "ИНДИЯ", "ТАЪМИНОТЧИ", TAG_B, 60, 5, 1000, 5000)])
        self.c = self.bot.client(F_CH)

    def ch_names(self):
        return {r[1] for r in SROK_CH} | {"ҚИСҚА НОМ", "ШОШ ДОРИ 00"}

    def test_bosh_menyu_qatori(self):
        rows = self.bot.filial_root_rows(self.c)
        self.assertIn([("⏰ Srok · 32", "kc:s"), ("📦 30+ kun · 3", "kc:o")], rows)
        self.assertIn([("⏰ Srok · 3", "kc:s"), ("📦 30+ kun · 0", "kc:o")], self.bot.filial_root_rows(self.bot.client(F_EK)))
        self.assertIn([("⏰ Srok · 0", "kc:s"), ("📦 30+ kun · 1", "kc:o")], self.bot.filial_root_rows(self.bot.client(F_UCH)))
        text = self.bot.filial_home_text(self.c)
        self.assertIn("⏰ Srok · A: <b>25</b> ta · ", text)
        self.assertIn("📦 30+ kun · A: <b>", text)

    def test_umumiy_ekran(self):
        text, kb = self.press(F_CH, "kc:s")
        self.assertTrue(text.startswith("🏠 › ⏰ Srok\n\n⏰ <b>Srok</b> · 🏬 %s" % CH))
        self.assertIn("Jami: <b>32</b> ta · <b>", text)
        self.assertIn("\n<b>A</b> · 25 ta · ", text)
        self.assertIn("\n<b>B</b> · 5 ta · ", text)
        self.assertIn("\n<b>C</b> · 2 ta · ", text)
        rows, groups, sm, val = self.bot.fil_abc(CH, "s")
        self.assertEqual({k: sm[k]["n"] for k in "ABC"}, {"A": 25, "B": 5, "C": 2})
        self.assertIn("\n<b>A</b> · 25 ta · %s · %d%%" % (B.mln(sm["A"]["sum"], False), sm["A"]["share"]), text)
        self.assertEqual(sum(sm[k]["share"] for k in "ABC"), 100)
        b = buttons(kb)
        self.assertEqual([x["text"] for x in b[:3]], ["A · 25 ta", "B · 5 ta", "C · 2 ta"])
        self.assertEqual(b[0].get("style"), "primary")
        data = inline_data(kb)
        self.assertEqual(data[:3], ["kc:s:A", "kc:s:B", "kc:s:C"])
        self.assertTrue(data[3].startswith("kc:s:x:"))
        self.assertEqual(data[-2:], ["kh", "kh"], "◀ Orqaga · 🏠 Menyu")
        self.assertTrue(all(len(d.encode()) < 64 for d in data))

    def test_sinf_royxati_20_ta_va_excel(self):
        text, kb = self.press(F_CH, "kc:s:A")
        self.assertTrue(text.startswith("🏠 › ⏰ Srok › A\n\n⏰ <b>Srok · A</b> · 🏬 %s\nPul bo'yicha · 25 tadan 20 tasi" % CH))
        nums = re.findall(r"\n(\d+)\. <b>", text)
        self.assertEqual(nums, [str(i) for i in range(1, 21)], "faqat 20 ta")
        # pul bo'yicha kamayish: СРОК 00 (20 000) birinchi, СРОК 19 (18 100) yigirmanchi
        self.assertIn("\n1. <b>СРОК 00</b> 🔴\n1 ta · %s (5 kun) · <b>20 000</b>" % day(5).strftime("%d.%m.%Y"), text)
        self.assertIn("\n20. <b>СРОК 19</b>\n1 ta · %s (43 kun) · <b>18 100</b>" % day(43).strftime("%d.%m.%Y"), text)
        self.assertNotIn("СРОК 20", text)
        self.assertIn("\n13. <b>СРОК 12</b> 🔴", text, "29 kun — 🔴")
        self.assertIn("\n14. <b>СРОК 13</b>\n", text, "31 kun — belgisiz")
        self.assertIn("20 / 25 ko'rsatildi", text)
        self.assertIn("Qolgan 5 tasi", text)
        self.assertIn("🔴 — muddatiga 30 kundan kam qolgan", text)
        b = buttons(kb)
        self.assertEqual(b[0]["text"], "📥 Qolgan 5 tasini Excel'da olish")
        self.assertEqual([(x["text"], x["callback_data"]) for x in b[-2:]], [("◀ A/B/C", "kc:s"), ("🏠 Menyu", "kh")])
        self.assertFalse(self.tg.to(F_CH, "doc"))
        before = len(self.tg.sent)
        self.press(F_CH, b[0]["callback_data"])
        docs = self.tg.to(F_CH, "doc")
        self.assertEqual(len(docs), 1, "Excel — alohida xabar")
        self.assertTrue(docs[0][2].startswith("Srok_A_ЧИНОЗ_") and docs[0][2].endswith(".xlsx"))
        self.assertIn("Faylda: 25 ta", docs[0][4])
        self.assertIn("21–25 tasi — faqat shu faylda", docs[0][4])
        after = self.tg.sent[before:]
        self.assertEqual([s[0] for s in after], ["doc", "send"], "fayl, keyin panel uning ostiga")
        self.assertIn("📎 Excel yuborildi", after[1][2])
        ws = load_workbook(io.BytesIO(docs[0][3])).active
        head = [c.value for c in ws[1]]
        self.assertEqual(head, ["ABC", "Дори", "Ишлаб чиқарувчи", "Срок", "Қолган кун", "Қолдиқ", "Кирим нархи", "Сумма"])
        self.assertNotIn("Фирма", " ".join(head), "filial xodimi firma ko'rmaydi")
        data = list(ws.iter_rows(min_row=2, values_only=True))
        self.assertEqual(len(data), 25, "sinfning hammasi")
        self.assertEqual([r[1] for r in data], ["СРОК %02d" % i for i in range(25)], "xabardagi tartib")
        self.assertEqual({r[0] for r in data}, {"A"})
        self.assertEqual(data[0][-1], 20000)

    def test_kichik_sinf_excel_tugmasiz(self):
        text, kb = self.press(F_CH, "kc:s:C")
        self.assertIn("2 / 2 ko'rsatildi", text)
        self.assertIn("\n1. <b>ҚИСҚА НОМ</b>", text, "qisqa filial nomi ham shu filial")
        self.assertIn("\n2. <b>ШОШ ДОРИ 00</b> 🔴\n4 ta · %s (−5 kun) · <b>10 000</b>" % day(-5).strftime("%d.%m.%Y"), text)
        self.assertEqual(inline_data(kb), ["kc:s", "kh"])

    def test_hammasi_excel(self):
        _, kb = self.press(F_CH, "kc:s")
        x = [d for d in inline_data(kb) if d.startswith("kc:s:x:")][0]
        self.press(F_CH, x)
        doc = self.tg.to(F_CH, "doc")[-1]
        data = list(load_workbook(io.BytesIO(doc[3])).active.iter_rows(min_row=2, values_only=True))
        self.assertEqual(len(data), 32)
        self.assertEqual([r[0] for r in data], ["A"] * 25 + ["B"] * 5 + ["C"] * 2)
        self.assertEqual({r[1] for r in data}, self.ch_names())

    def test_eskirgan_tugma_fayl_yubormaydi(self):
        _, kb = self.press(F_CH, "kc:s:A")
        x = inline_data(kb)[0]
        add_rows(self.bot, "Srok", [(CH, "ЯНГИ ҚИММАТ", "ИНДИЯ", TAG_A, day(50), 50, 1, 900000)])
        text, _ = self.press(F_CH, x)
        self.assertFalse(self.tg.to(F_CH, "doc"))
        self.assertIn("Ma'lumot yangilandi", text)
        self.assertIn("ЯНГИ ҚИММАТ", text, "yangi ro'yxat ko'rsatildi")
        self.press(F_CH, "kc:s:A:x:00000000")
        self.assertFalse(self.tg.to(F_CH, "doc"))

    def test_30_kun(self):
        text, kb = self.press(F_CH, "kc:o")
        self.assertTrue(text.startswith("🏠 › 📦 30+ kun\n\n📦 <b>30+ kun</b> · 🏬 %s" % CH))
        self.assertIn("Jami: <b>3</b> ta", text)
        text, _ = self.press(F_CH, "kc:o:A")
        self.assertIn("\n1. <b>ЭСКИ 1</b>\n3 ta · 45 kun turibdi · <b>50 000</b>", text, "summa ustuni")
        text, _ = self.press(F_CH, "kc:o:B")
        self.assertIn("\n1. <b>ЭСКИ 2</b>\n2 ta · 120 kun turibdi · <b>8 000</b>", text, "summa yo'q — qoldiq × narx")
        self.assertNotIn("УЧ ЭСКИ", text)

    # ---- xavfsizlik: filial faqat o'z ma'lumotini ko'radi ----
    def test_filiallar_ajratilgan(self):
        other = {"ЭКС СРОК 1", "ЭКС СРОК 2", "АВИТЕН", "БЕГОНА ДОРИ", "СКЛАД ДОРИ", "УЧ ЭСКИ", "ШОШ ДОРИ 01"}
        seen = []
        for data in ("kc:s", "kc:s:A", "kc:s:B", "kc:s:C", "kc:o", "kc:o:A", "kc:o:B", "kc:o:C"):
            seen.append(self.press(F_CH, data)[0])
        seen.append(self.bot.filial_home_text(self.c))
        for t in seen:
            for n in other:
                self.assertNotIn(n, t, "ЧИНОЗ boshqa filial dorisini ko'rmaydi")
        for data in [d for d in inline_data(self.press(F_CH, "kc:s")[1]) if ":x:" in d] + [
                d for d in inline_data(self.press(F_CH, "kc:s:A")[1]) if ":x:" in d]:
            self.press(F_CH, data)
        for doc in self.tg.to(F_CH, "doc"):
            names = {r[1] for r in load_workbook(io.BytesIO(doc[3])).active.iter_rows(min_row=2, values_only=True)}
            self.assertFalse(names & other)
        # EK: faqat o'zi (2 probelli nom ham)
        ek = " ".join(self.press(F_EK, d)[0] for d in ("kc:s", "kc:s:A", "kc:s:B", "kc:s:C"))
        for n in ("ЭКС СРОК 1", "ЭКС СРОК 2", "АВИТЕН"):
            self.assertIn(n, ek)
        for n in self.ch_names() | {"БЕГОНА ДОРИ"}:
            self.assertNotIn(n, ek)
        # UCH: Srok yo'q (200 kunlik qator chegaradan tashqari), 30+ kun — faqat o'zi
        self.assertIn("✅ Bunday dori yo'q.", self.press(F_UCH, "kc:s")[0])
        uch = self.press(F_UCH, "kc:o:A")[0]
        self.assertIn("УЧ ЭСКИ", uch)
        self.assertNotIn("ЭСКИ 1", uch)

    def test_boshqa_rollar_kc_ishlatolmaydi(self):
        for cid in (FIRM, STAFF, ADMIN):
            n = len(self.tg.sent)
            self.bot.handle(cb(cid, "kc:s:A"))
            self.bot.handle(cb(cid, "kc:s:x:00000000"))
            self.assertEqual(len(self.tg.sent), n, "filial xodimi emas — javob yo'q")
        add_client(self.bot, 7777, "filial", status="pending", filial=CH)
        n = len(self.tg.sent)
        self.bot.handle(cb(7777, "kc:s"))
        self.assertEqual(len(self.tg.sent), n, "faol bo'lmagan filial xodimi")

    def test_sklad_xodimiga_summa_korsatilmaydi(self):
        add_client(self.bot, SK_ID, "filial", filial="ДОРИХОНА СКЛАД")
        text, _ = self.press(SK_ID, "kc:s")
        self.assertIn("Jami: <b>1</b> ta\n", text)
        self.assertNotIn("so'm", text)
        self.assertNotIn("mln", text)
        text, kb = self.press(SK_ID, "kc:s:A")
        self.assertIn("СКЛАД ДОРИ", text)
        self.assertNotIn("5 000", text)

    def test_eski_matn_tugmasi(self):
        self.bot.handle({"message": {"chat": {"id": F_CH, "type": "private"}, "from": {"id": F_CH}, "text": "⏰ Srok · 32"}})
        self.assertTrue(self.tg.sent[-1][2].startswith("🏠 › ⏰ Srok"))


# =====================================================================================================================
# 7) 📊 Sotuv va qoldiq
# =====================================================================================================================
SALE_EXTRA = [  # nomi, ishlab chiqaruvchi, mnn, firma, CH qoldiq, CH sotuv, UCH qoldiq, UCH sotuv, EK qoldiq, EK sotuv, sklad
    ("ТЕЗ ТУГАЙДИ", "ИНДИЯ", "", TAG_A, 0, 30, 1, 20, 2, 10, 0),
    ("КЎП СОТИЛАДИ", "ИНДИЯ", "", TAG_A, 50, 40, 50, 30, 50, 30, 100),
    ("СОТУВСИЗ", "ИНДИЯ", "", TAG_A, 4, 0, 0, 0, 0, 0, 2),
    ("БЕТА ТЕЗ", "ИНДИЯ", "", TAG_B, 0, 50, 0, 50, 0, 50, 0)]


class SotuvQoldiq(Press, BotCase):
    def setUp(self):
        super().setUp()
        add_rows(self.bot, None, SALE_EXTRA, first=8)

    def blocks(self, text):
        return re.findall(r"💊 <b>(.+?)</b>", text)

    def pre(self, text, name):
        m = re.search(r"💊 <b>%s</b>[^\n]*\n<pre>(.*?)</pre>\n<i>(.*?)</i>" % re.escape(name), text, re.S)
        return m.group(1), m.group(2)

    def firm_names(self, tag):
        return {r["name"] for r in self.bot.data.sale if B.norm(r["firm"]) == tag}

    def all_pages(self, cid, sort, pre="c:"):
        names, page, seen = [], 0, set()
        while page not in seen:
            seen.add(page)
            text, kb = self.press(cid, "%ssq:%s:%d" % (pre, sort, page))
            names += self.blocks(text)
            nxt = [b["callback_data"] for b in buttons(kb) if b["text"] == "▶"]
            if not nxt:
                break
            page = int(nxt[0].rsplit(":", 1)[1])
        return names

    def test_firma_vakili_ekrani(self):
        self.assertIn(("📊 Sotuv va qoldiq", "c:stock"), [b for r in self.bot.client_root_rows(self.bot.client(FIRM)) for b in r])
        text, kb = self.press(FIRM, "c:stock")
        n = len(self.firm_names(TAG_A))
        self.assertTrue(text.startswith("🏠 › 📊 Sotuv va qoldiq\n\n📊 <b>Sotuv va qoldiq</b> · %s\nOxirgi 10 kun · <b>%d</b> xil dori" % (
            TAG_A, n)))
        self.assertEqual(self.blocks(text)[0], "КЎП СОТИЛАДИ", "🔥 Sotuv: eng ko'p sotilgan birinchi")
        self.assertEqual(len(self.blocks(text)), 5)
        self.assertIn("\n1–5 / %d · 🕒 " % n, text)
        b = buttons(kb)
        self.assertEqual([(x["text"], x["callback_data"], x.get("style")) for x in b[:3]],
                         [("🔥 Sotuv", "c:sq:s:0", "primary"), ("⚠️ Tugayapti", "c:sq:t:0", None), ("🔤 А–Я", "c:sq:a:0", None)])
        pages = (n - 1) // 5 + 1
        self.assertEqual([(x["text"], x["callback_data"]) for x in b[3:6]],
                         [("◀", "c:sq:s:%d" % (pages - 1)), ("1 / %d" % pages, "c:sq:s:0"), ("▶", "c:sq:s:1")])
        data = inline_data(kb)
        self.assertIn("c:find", data)
        self.assertIn("x:stock", data)
        self.assertEqual(data[-2:], ["c:home", "c:home"])
        self.assertNotIn("a:fm:stock", data, "firma vakili firma tanlamaydi")
        self.assertTrue(all(len(d.encode()) < 64 for d in data))

    def test_blok_jadvali(self):
        text, _ = self.press(FIRM, "c:stock")
        pre, tail = self.pre(text, "КЎП СОТИЛАДИ")
        lines = pre.split("\n")
        self.assertEqual([ln.split()[0] for ln in lines], ["Сотилди", "Чиноз", "Ургут", "Эркин", "Склад", "Жами"],
                         "Sale sarlavhasidagi filial nomlari (kirill), Sklad, Jami")
        self.assertEqual(lines[0].split(), ["Сотилди", "Бор"])
        self.assertEqual(lines[1].split(), ["Чиноз", "40", "50"])
        self.assertEqual(lines[4].split(), ["Склад", "—", "100"])
        self.assertEqual(lines[5].split(), ["Жами", "100", "250"])
        self.assertEqual(len({len(ln) for ln in lines}), 1, "ustunlar tekis (bir xil kenglik)")
        self.assertFalse(re.search("[\U0001F300-\U0001FAFF☀-➿]", pre), "jadval ichida emoji yo'q")
        self.assertEqual(tail, "≈ 25 kunga yetadi")
        self.assertNotIn("КЎП СОТИЛАДИ</b> ⚠️", text)

    def test_tugayapti(self):
        text, kb = self.press(FIRM, "c:sq:t:0")
        self.assertTrue(text.startswith("🏠 › 📊 Sotuv va qoldiq › ⚠️ Tugayapti\n\n⚠️ <b>Tugayapti</b> · %s\n"
                                        "7 kundan kam yetadigan dorilar: <b>1</b> ta" % TAG_A))
        self.assertEqual(self.blocks(text), ["ТЕЗ ТУГАЙДИ"], "BETA firmasining dorisi ko'rinmaydi")
        self.assertIn("💊 <b>ТЕЗ ТУГАЙДИ</b> ⚠️", text)
        pre, tail = self.pre(text, "ТЕЗ ТУГАЙДИ")
        self.assertEqual(tail, "≈ 0 kunga yetadi · buyurtma bering")
        lines = pre.split("\n")
        self.assertTrue(all(ln.endswith(" !") for ln in lines[1:4]), "filialda 0 yoki 3 kundan kamga — «!»")
        self.assertFalse(lines[4].endswith("!"))
        self.assertEqual([b.get("style") for b in buttons(kb)[:3]], [None, "primary", None])
        self.assertNotIn("◀", [b["text"] for b in buttons(kb)], "bitta sahifa — sahifalagich yo'q")

    def test_sotuvsiz_dori(self):
        names = self.all_pages(FIRM, "a")
        self.assertIn("СОТУВСИЗ", names)
        page = names.index("СОТУВСИЗ") // 5
        text, _ = self.press(FIRM, "c:sq:a:%d" % page)
        self.assertEqual(self.pre(text, "СОТУВСИЗ")[1], "≈ — kunga yetadi")

    def test_alifbo_va_sahifalash(self):
        names = self.all_pages(FIRM, "a")
        want = sorted(self.firm_names(TAG_A), key=B.norm)
        self.assertEqual(names, want, "А–Я, har sahifada 5 tadan, hammasi bir marta")
        pages = (len(want) - 1) // 5 + 1
        text, _ = self.press(FIRM, "c:sq:a:-1")
        self.assertEqual(self.blocks(text), want[(pages - 1) * 5:], "◀ birinchi sahifada — oxirgisiga")
        self.assertIn("%d–%d / %d" % ((pages - 1) * 5 + 1, len(want), len(want)), text)

    # ---- xavfsizlik: firma vakili faqat o'z firmasi ----
    def test_firma_ajratilgan(self):
        mine, other = self.firm_names(TAG_A), self.firm_names(TAG_B)
        for sort in ("s", "t", "a"):
            names = set(self.all_pages(FIRM, sort))
            self.assertTrue(names <= mine, sort)
            self.assertFalse(names & other, sort)
        self.assertEqual(set(self.all_pages(FIRM, "s")), mine)

    def test_xodim_avval_firma_tanlaydi(self):
        text, kb = self.press(STAFF, "c:stock")
        self.assertTrue(text.startswith("🏠 › 📊 Sotuv va qoldiq › 🏢 Firma\n\n📊 <b>Sotuv va qoldiq</b> · 👔 Xodim"))
        self.assertEqual(self.bot.state[STAFF][0], "firm_pick")
        data = inline_data(kb)
        self.assertIn("c:sqa", data)
        fp = [d for d in data if d.startswith("fp:")]
        self.assertEqual(len(fp), 2, "eng ko'p sotuvchi firmalar")
        sales = lambda tag: sum(r["sales"] for r in self.bot.data.sale if B.norm(r["firm"]) == tag)
        top = self.bot.fsearch[STAFF]
        self.assertEqual(top, sorted([TAG_A, TAG_B], key=lambda t: -sales(t)), "sotuv bo'yicha")
        first, second = top
        text, kb = self.press(STAFF, fp[0])
        self.assertIn("📊 <b>Sotuv va qoldiq</b> · %s" % first, text)
        self.assertIn("a:fm:stock", inline_data(kb))
        self.assertEqual(set(self.all_pages(STAFF, "s")), self.firm_names(first))
        self.assertFalse(set(self.all_pages(STAFF, "a")) & self.firm_names(second))
        # 🌐 Hamma firmalar
        self.bot.fmode.pop(STAFF, None)
        self.press(STAFF, "c:stock")
        text, _ = self.press(STAFF, "c:sqa")
        self.assertIn("📊 <b>Sotuv va qoldiq</b> · Hamma firmalar", text)
        self.assertEqual(set(self.all_pages(STAFF, "s")), self.firm_names(TAG_A) | self.firm_names(TAG_B))
        text, _ = self.press(STAFF, "c:stock")
        self.assertNotIn("🏢 Firma\n", text.split("\n\n")[0], "qayta so'ralmaydi")

    def test_xodim_firma_nomini_yozadi(self):
        self.press(STAFF, "c:stock")
        self.bot.handle({"message": {"chat": {"id": STAFF, "type": "private"}, "from": {"id": STAFF}, "text": "АЛЬФА"}})
        text = self.tg.sent[-1][2]
        self.assertIn("📊 <b>Sotuv va qoldiq</b> · %s" % TAG_A, text, "yagona topilgan firma — shu bo'limga qaytadi")

    def test_admin(self):
        rows = self.bot.admin_panel(ADMIN)[1]
        self.assertIn("a:sec:stock", inline_data(rows))
        text, kb = self.press(ADMIN, "a:sec:stock")
        self.assertIn("👔 Admin", text)
        self.assertIn("a:sqa", inline_data(kb))
        text, kb = self.press(ADMIN, "a:sqa")
        self.assertIn("Hamma firmalar", text)
        self.assertEqual(inline_data(kb)[:3], ["a:sq:s:0", "a:sq:t:0", "a:sq:a:0"])
        self.assertIn("a:find", inline_data(kb))
        self.assertEqual(inline_data(kb)[-2:], ["a:home", "a:home"])
        text, _ = self.press(ADMIN, "a:sq:t:0")
        self.assertEqual(self.blocks(text), ["БЕТА ТЕЗ", "ТЕЗ ТУГАЙДИ"])
        self.press(ADMIN, "a:sq:s:0")
        self.press(ADMIN, "x:stock")
        self.assertTrue(self.tg.to(ADMIN, "doc")[-1][2].startswith("Qoldiq_"))
        self.assertIn("Dori nomini yozing", self.press(ADMIN, "a:find")[0])

    def test_filial_xodimi_korolmaydi(self):
        for d in ("c:stock", "c:sq:s:0", "c:sqa"):
            n = len(self.tg.sent)
            self.bot.handle(cb(F_CH, d))
            self.assertEqual(len(self.tg.sent), n)


# =====================================================================================================================
# 8) Ko'chirish A / B / C
# =====================================================================================================================
class KochAbcBot(Press, BotCase):
    def setUp(self):
        super().setUp()
        self.bot.koch_build()
        self.c = self.bot.client(F_CH)
        self.today = B.now().date().isoformat()

    def tid(self, name):
        return self.bot.db.execute("SELECT id FROM koch_tasks WHERE name=?", (name,)).fetchone()[0]

    def expected(self):
        tasks = [t for t in self.bot.koch_tasks(self.today) if t["status"] != "bekor"]
        return ABC.classify(tasks, lambda t: (t["qty_plan"] or 0) * (t["prih"] or 0), lambda t: t["id"]), {t["id"]: t for t in tasks}

    def test_sinflar_butun_tarmoq(self):
        exp, tasks = self.expected()
        self.assertEqual(self.bot.koch_abc(), exp)
        self.assertEqual(exp[self.tid("АВИТЕН")], "A", "3 × 20 000 — eng qimmatlaridan")
        self.assertEqual(set(exp.values()), {"A", "B", "C"})
        self.assertIs(self.bot.koch_abc(), self.bot.koch_abc(), "kesh")
        self.bot.db.execute("UPDATE koch_tasks SET qty_plan=500 WHERE name='ОДДИЙ ЧАР'")
        self.bot.db.commit()
        self.assertEqual(self.bot.koch_abc()[self.tid("ОДДИЙ ЧАР")], "A", "imzo o'zgardi — qayta hisoblandi")
        self.bot.db.execute("UPDATE koch_tasks SET status='bekor' WHERE name='ОДДИЙ ЧАР'")
        self.bot.db.commit()
        self.assertNotIn(self.tid("ОДДИЙ ЧАР"), self.bot.koch_abc(), "bekor — hisobga kirmaydi")

    def test_royxatda_belgi_va_tartib(self):
        exp, tasks = self.expected()
        for cat in ("shoshilinch", "assortiment", "oddiy"):
            text, kb = self.bot.koch_list_text(self.c, cat)
            ids = [int(d.split(":")[1]) for d in inline_data(kb) if d.startswith("ki:")]
            marks = re.findall(r"\n\d+\. <code>([ABC])</code> ", text)
            self.assertEqual(marks, [exp[i] for i in ids], "belgi = sinf, raqam = tugma")
            groups = {}
            for i in ids:
                groups.setdefault(B.norm(tasks[i]["dst"]), []).append(i)
            for g in groups.values():
                key = [(B.Bot.ABC_RANK[exp[i]], -tasks[i]["qty_plan"] * tasks[i]["prih"]) for i in g]
                self.assertEqual(key, sorted(key), "guruh ichida: A birinchi, keyin pul ↓")
        # yo'nalish ro'yxati (bitta oluvchi) ham shu tartibda
        text, kb = self.bot.koch_list_text(self.c, "shoshilinch", 0, 0)
        self.assertIn("<code>", text)

    def test_kartochkada_belgi(self):
        text = strip_marks(self.bot.koch_item_card(self.bot.client(F_EK), self.tid("АВИТЕН"))[0])
        self.assertTrue(text.startswith("🔴 <b>SHOSHILINCH</b> · #%d · <code>A</code>\n💊 <b>АВИТЕН</b>" % self.tid("АВИТЕН")))
        text, _ = self.press(F_EK, "ki:%d:0" % self.tid("АВИТЕН"))
        self.assertIn("· <code>A</code>", text)

    def test_bosh_menyu_A_qoldi(self):
        exp, tasks = self.expected()
        a = [i for i, t in tasks.items() if exp[i] == "A" and t["src"] == CH]
        self.assertTrue(a)
        s = sum(tasks[i]["qty_plan"] * tasks[i]["prih"] for i in a)
        self.assertIn("💰 A: <b>%d</b> ta qoldi · %s" % (len(a), B.mln(s)), self.bot.filial_home_text(self.c))
        self.assertTrue(self.bot.koch_give(F_CH, self.c, a[0]))
        self.assertIn("💰 A: <b>%d</b> ta qoldi" % (len(a) - 1), self.bot.filial_home_text(self.c))

    def test_bugungi_holat(self):
        exp, tasks = self.expected()
        mine = {i: t for i, t in tasks.items() if t["src"] == CH}
        a = sorted((i for i in mine if exp[i] == "A"), key=lambda i: -mine[i]["qty_plan"] * mine[i]["prih"])
        self.bot.koch_give(F_CH, self.c, a[0])
        b = [i for i in mine if exp[i] == "B"][0]
        self.bot.koch_refuse(F_CH, self.c, b, "Tovar topilmadi")
        t = self.bot.day_text(CH)
        self.assertIn("Jami: <b>%d</b> ta · bajarildi <b>1</b> · qoldi <b>%d</b>" % (len(mine), len(mine) - 1), t)
        pre = re.search(r"<pre>(.*?)</pre>", t, re.S).group(1).split("\n")
        self.assertEqual(pre[0].split(), ["Бажарилди", "Қолди"])
        cnt = lambda k: sum(1 for i in mine if exp[i] == k)
        self.assertEqual(pre[1].split(), ["A", "(80%)", "1", "та", str(cnt("A") - 1), "та"])
        self.assertEqual(pre[2].split(), ["B", "(15%)", "0", "та", str(cnt("B")), "та"], "rad — bajarilmagan")
        self.assertEqual(pre[4].split(), ["Жами", "1", "та", str(len(mine) - 1), "та"])
        self.assertEqual(len({len(x) for x in pre}), 1, "tekis")
        self.assertIn("❌ <b>Bajarilmaganlar (A):</b>\n1. <b>%s</b>" % mine[a[1]]["name"], t)
        self.assertIn("🔴 <b>SHOSHILINCH</b>\n📤 beradi:", t, "toifalar kesimi saqlangan")
        text, kb = self.press(F_CH, "kd")
        self.assertTrue(text.startswith("🏠 › 📋 Bugungi holat"))
        self.assertEqual(inline_data(kb)[:2], ["kd:n:0", "kd:d:0"])
        self.assertEqual([x["text"] for x in buttons(kb)[:2]], ["❌ Bajarilmaganlar · %d" % (len(mine) - 1), "✅ Bajarilganlar · 1"])
        text, kb = self.press(F_CH, "kd:d:0")
        self.assertTrue(text.startswith("🏠 › 📋 Bugungi holat › ✅ Bajarilganlar"))
        self.assertIn("\n\n<b>A</b> · 1 ta\n1. <code>A</code> <b>%s</b>" % mine[a[0]]["name"], text)
        self.assertEqual(inline_data(kb), ["kd", "kh"])
        text, kb = self.press(F_CH, "kd:n:0")
        heads = re.findall(r"\n\n<b>([ABC])</b> · (\d+) ta", text)
        self.assertEqual(heads[0], ("A", str(cnt("A") - 1)))
        self.assertEqual([h[0] for h in heads], sorted(h[0] for h in heads), "A, B, C tartibida")
        self.assertIn("1-sahifa / 2", text, "11 ta — 10 tadan")
        self.assertIn("kd:n:1", inline_data(kb))
        text, _ = self.press(F_CH, "kd:n:1")
        self.assertIn("\n11. <code>", text)
        rad = [x for x in self.bot.koch_abc_summary(filial=CH)["not_done"] if x["id"] == b][0]
        self.assertEqual((rad["status"], rad["reason"], rad["cls"]), ("rad", "Tovar topilmadi", "B"))

    def test_summary_shakli(self):
        exp, tasks = self.expected()
        s = self.bot.koch_abc_summary()
        self.assertEqual(set(s) >= {"sana", "filial", "has", "limits", "total", "classes", "filials", "matrix", "done", "not_done"},
                         True)
        keys = {"jami", "bajarildi", "bajarilmadi", "sum_jami", "sum_bajarildi", "sum_bajarilmadi"}
        self.assertEqual(set(s["total"]), keys)
        for k in "ABC":
            self.assertEqual(set(s["classes"][k]), keys | {"share"})
            self.assertEqual(s["classes"][k]["jami"], sum(1 for v in exp.values() if v == k))
        self.assertEqual(s["total"]["jami"], len(tasks))
        self.assertEqual(sum(v[c][1] for v in s["filials"].values() for c in "ABC"), len(tasks))
        self.assertEqual(sorted(s["matrix"]), sorted({t["src"] for t in tasks.values()}))
        item = s["not_done"][0]
        for k in ("id", "cls", "name", "src", "dst", "qty", "sum", "status", "time", "given_at", "received_at", "cat"):
            self.assertIn(k, item)
        order = [(B.Bot.ABC_RANK[i["cls"]], -i["sum"]) for i in s["not_done"]]
        self.assertEqual(order, sorted(order))
        f = self.bot.koch_abc_summary(filial=EK2)  # 2 probel bilan ham
        self.assertEqual(f["filial"], EK)
        self.assertTrue(all(i["src"] == EK for i in f["done"] + f["not_done"]))
        self.assertEqual(f["total"]["jami"], sum(1 for t in tasks.values() if t["src"] == EK))
        self.assertEqual(f["classes"]["A"]["jami"], sum(1 for i, t in tasks.items() if t["src"] == EK and exp[i] == "A"),
                         "sinf butun tarmoqdan, son — shu filial")

    def test_excel_rejimi_belgilar(self):
        self.bot.set_setting("koch_on", "0")
        text, _ = self.bot.koch_list_text(self.c, "shoshilinch")
        marks = re.findall(r"\n\d+\. <code>([ABC])</code> ", text)
        self.assertEqual(len(marks), 7)
        self.assertEqual(marks[0], "A")


if __name__ == "__main__":
    unittest.main()
