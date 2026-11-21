# -*- coding: utf-8 -*-
"""B4: filial xodimi paneli — 3 toifa, rangli klaviatura, filtr/saralash, ro'yxat, sahifalash, Excel, Keladigan."""
import io
import json
import os
import re
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fixtures import B, BotCase, CH, UCH, EK, STAFF, F_CH, F_EK, F_UCH, msg, cb, kb_buttons, inline_data
from openpyxl import load_workbook


def names(rows):
    return [r["name"] for r in rows]


class Panel(BotCase):
    def setUp(self):
        super().setUp()
        self.bot.koch_build()
        self.c = self.bot.client(F_CH)

    def tid(self, name):
        return self.bot.db.execute("SELECT id FROM koch_tasks WHERE name=?", (name,)).fetchone()[0]

    # ---- B4.1 klaviatura ----
    def test_klaviatura_style_va_sonlar(self):
        kb = self.bot.filial_kb(self.c)
        rows = [[b["text"] for b in row] for row in kb["keyboard"]]
        self.assertEqual(rows, [["🔴 SHOSHILINCH · 7"], ["🟡 ASSORTIMENT · 4", "🟢 ODDIY · 1"],
                                ["📥 Keladigan · 6", "📋 Bugungi holat"], ["⏰ Srok · 1", "📦 30+ kun · 0"], ["🏠 Menyu"]])
        b = kb_buttons(kb)
        self.assertEqual(b[0]["style"], "danger")
        self.assertNotIn("style", b[1], "ASSORTIMENT: style yo'q (sariq rang yo'q)")
        self.assertEqual(b[2]["style"], "success")
        self.assertTrue(all("style" not in x for x in b[3:]))
        self.assertTrue(kb["is_persistent"])
        js = json.dumps(kb, ensure_ascii=False)
        self.assertNotIn("null", js, "style=None kaliti yuborilmaydi")
        for cat, n in (("shoshilinch", 7), ("assortiment", 4), ("oddiy", 1)):
            self.assertIn("Jami: <b>%d</b> ta dori" % n, self.bot.koch_list_text(self.c, cat)[0], "tugmadagi N = Jami")

    def test_n0_bolsa_ham_tugma(self):
        kb = self.bot.filial_kb(self.bot.client(F_UCH))
        self.assertEqual([b["text"] for b in kb_buttons(kb)][:3], ["🔴 SHOSHILINCH · 0", "🟡 ASSORTIMENT · 0", "🟢 ODDIY · 0"])

    def test_xush_kelibsiz(self):
        text, _ = self.bot.filial_panel(self.c)
        self.assertIn("🏬 <b>%s</b>" % CH, text)
        for line in ("🔴 SHOSHILINCH — berish kerak: <b>7</b>", "🟡 ASSORTIMENT — berish kerak: <b>4</b>",
                     "🟢 ODDIY — berish kerak: <b>1</b>", "📥 Sizga keladi: <b>6</b>",
                     "⏰ SHOSHILINCH muddati 10:00 · 2 soat 0 daqiqa qoldi", "🕒 Ma'lumot:"):
            self.assertIn(line, text)

    def test_tugma_matni_handler(self):
        for t in ("🔴 SHOSHILINCH · 99", "shoshilinch", "SHOSHILINCH · 0", "🚨 Zudlik · 3"):
            self.bot.handle(msg(F_CH, t))  # eski pastki tugma (yozilgan matn) → shu daraja yangi panel bo'lib pastga tushadi
            last = self.tg.last(F_CH)
            self.assertEqual(last[0], "send")
            self.assertTrue(last[2].startswith("🏠 › 🔴 SHOSHILINCH\n\n🔴 <b>SHOSHILINCH</b> · 🏬 " + CH), (t, last[2][:80]))
        self.bot.handle(msg(F_CH, "🟡 assortiment · 4"))
        self.assertTrue(any("🟡 <b>ASSORTIMENT</b>" in x for x in self.tg.texts(F_CH)[-2:]))
        self.bot.handle(msg(STAFF, "SHOSHILINCH"))
        self.assertFalse([x for x in self.tg.texts(STAFF) if "<b>SHOSHILINCH</b> · 🏬" in x], "boshqa rol — avvalgi xatti-harakat")

    # ---- B4.3 filtr va saralash ----
    def test_shoshilinch_filtr_saralash(self):
        rows, exc = self.bot.koch_view(CH, "shoshilinch")
        self.assertEqual(names(rows), ["ШОШ ДОРИ 00", "ШОШ ДОРИ 06", "ШОШ ДОРИ 01", "ШОШ ДОРИ 02", "ШОШ ДОРИ 07",
                                       "ШОШ ДОРИ 04", "ШОШ ДОРИ 05"], "muddat ↑, noma'lum oxirida, summa ↓, nom")
        self.assertEqual([(r["name"], why) for r, why in exc], [("ШОШ ДОРИ 03", "Muddati o'tgan")])
        self.assertTrue(all(B.norm(r["src"]) == CH for r in rows), "faqat A = o'z filiali")
        self.assertNotIn("КАСР", names(rows), "E < 1 chiqmaydi")

    def test_assortiment_filtr_saralash(self):
        rows, exc = self.bot.koch_view(CH, "assortiment")
        self.assertEqual(names(rows), ["АССОРТ 03", "АССОРТ 07", "АССОРТ 06", "АССОРТ 04"], "beruvchidagi sotuv ↓")
        self.assertEqual(sorted((r["name"], why[:12]) for r, why in exc),
                         [("АССОРТ 00", "Muddati 60 k"), ("АССОРТ 01", "Muddati 60 k"), ("АССОРТ 02", "Muddati 60 k"),
                          ("АССОРТ 05", "Muddati o'tg")])
        self.bot.set_setting("koch_assort_min_days", "30")
        self.assertEqual(len(self.bot.koch_view(CH, "assortiment")[1]), 2, "sozlama ishlaydi: 3 kunlik + o'tgan")

    def test_oddiy_filtr_saralash(self):
        rows, exc = self.bot.koch_view(EK, "oddiy")
        self.assertEqual(names(rows), ["ОДДИЙ 05", "ОДДИЙ 01", "ОДДИЙ 03", "ОДДИЙ 00", "ОДДИЙ 04"], "turgan kun ↓")
        self.assertEqual([r["name"] for r, _ in exc], ["ОДДИЙ 02"])

    # ---- B4.4 ro'yxat xabari ----
    def test_royxat_matni(self):
        text, kb = self.bot.koch_list_text(self.c, "shoshilinch")
        self.assertTrue(text.startswith("🔴 <b>SHOSHILINCH</b> · 🏬 %s\nTugagan, sotuv bor" % CH))
        self.assertIn("⏰ Muddat 10:00 gacha · 2 soat 0 daqiqa qoldi", text)
        self.assertIn("<b>→ %s</b> · 4 ta" % EK, text)
        self.assertIn("<b>→ %s</b> · 3 ta" % UCH, text)
        # oluvchi guruhi ichida: A birinchi, keyin pul (reja × kirim narxi) bo'yicha — 02 (30 000), 07 (18 000), 00 (12 500), 05
        self.assertRegex(text, r"\n1\. <code>A</code> ШОШ ДОРИ 02 · <b>1</b> qadoq · srok \d\d\.\d\d\.\d{4} \(u yerda sotuv 4\)")
        self.assertRegex(text, r"\n3\. <code>B</code> ШОШ ДОРИ 00 · <b>5</b> qadoq · srok \d\d\.\d\d\.\d{4} ⚠️ \(u yerda sotuv 2\)")
        self.assertIn("srok ?", text)
        self.assertIn("⛔ Muddati o'tgan (ko'chirilmaydi)</b> · 1 ta", text)
        nums = [int(x) for x in re.findall(r"\n(\d+)\. ", text)]
        self.assertEqual(sorted(nums), list(range(1, 8)), "umumiy tartib raqami")
        data = inline_data(kb)
        self.assertEqual(len([d for d in data if d.startswith("ki:")]), 7)
        self.assertIn("kx:shoshilinch", data)
        self.assertTrue(all(len(d.encode()) <= 64 for d in data))
        self.clock.at(10, 25)
        self.assertIn("⛔ Muddat 10:00 · 25 daqiqa kechikdi", self.bot.koch_list_text(self.c, "shoshilinch")[0])

    def test_qator_qoshimchalari(self):
        a = self.bot.koch_list_text(self.c, "assortiment")[0]
        self.assertIn("(sizda sotuv 7)", a)
        self.assertNotIn("⚠️", a, "ASSORTIMENT: ⚠️ yo'q")
        self.assertIn("ℹ️ Muddati 60 kundan kam bo'lgan 3 ta dori kiritilmadi (Excelda).", a)
        o = self.bot.koch_list_text(self.bot.client(F_EK), "oddiy")[0]
        self.assertIn("· 200 kun turibdi · (u yerda sotuv 3)", o)
        self.assertIn("ОДДИЙ 01 · <b>5</b> qadoq · srok ", o)
        self.assertRegex(o, r"ОДДИЙ 01 .* ⚠️ · 95 kun", "ODDIY: < 30 kun ⚠️")

    def test_kasr_vergul(self):
        self.assertEqual(self.bot.fq(0.5), "0,5")
        self.assertEqual(self.bot.fq(3), "3")

    def test_bosh_royxat(self):
        text, kb = self.bot.koch_list_text(self.bot.client(F_UCH), "shoshilinch")
        self.assertTrue(text.startswith("🔴 <b>SHOSHILINCH</b> · 🏬 %s\n\n✅ Hozircha shoshilinch ko'chirish yo'q." % UCH))
        self.assertIsNone(kb)

    def test_sahifalash(self):
        self.bot.set_setting("koch_page_size", "5")
        text, kb = self.bot.koch_list_text(self.c, "shoshilinch", 0)
        data = inline_data(kb)
        # guruhlangan tartib: avval EK guruhi (birinchi qator EK ga), ichida A / B / C va pul bo'yicha; keyin UCH
        self.assertEqual([d for d in data if d.startswith("ki:")], ["ki:%d:0" % self.tid(n) for n in (
            "ШОШ ДОРИ 02", "ШОШ ДОРИ 07", "ШОШ ДОРИ 00", "ШОШ ДОРИ 05", "ШОШ ДОРИ 01")])
        self.assertIn("kp:shoshilinch:1", data)
        self.assertNotIn("◀", [b["text"] for row in kb["inline_keyboard"] for b in row])
        self.assertIn("1-sahifa / 2", text)
        self.bot.handle(cb(F_CH, "kp:shoshilinch:1"))
        t2 = self.tg.texts(F_CH)[-1]
        self.assertIn("\n6. ", t2)
        self.assertIn("\n7. ", t2)
        self.assertNotIn("\n1. ", t2)
        self.assertIn("kp:shoshilinch:0", inline_data(self.tg.sent[-1][3]))

    def test_4096_dan_oshmaydi(self):
        rows = [(CH, EK, "ДОРИ %03d %s" % (i, "Ж" * 60), "Р", 1, 1, 0, 100, 100, None, "SHOSHILINCH", "Tugagan", 0, 2)
                for i in range(300)]
        bot, tg, work = __import__("fixtures").make_bot("new", rows)
        self.addCleanup(__import__("shutil").rmtree, work, True)
        __import__("fixtures").add_client(bot, F_CH, "filial", filial=CH)
        bot.set_setting("koch_on", "0")
        text, _ = bot.koch_list_text(bot.client(F_CH), "shoshilinch")
        self.assertLessEqual(len(text), 4096)
        __import__("fixtures").close_bot(bot)

    # ---- B4.5 Excel ----
    def excel(self, cid, cat):
        self.bot.handle(cb(cid, "kx:" + cat))
        docs = self.tg.to(cid, "doc")
        self.assertTrue(docs, "Excel yuborilmadi")
        return docs[-1][2], load_workbook(io.BytesIO(docs[-1][3]))

    def test_excel_assortiment(self):
        fname, wb = self.excel(F_CH, "assortiment")
        self.assertRegex(fname, r"^ASSORTIMENT_ЧИНОЗ_\d{4}-\d\d-\d\d\.xlsx$")
        self.assertEqual(wb.sheetnames, [CH, "Ko'chirilmaydi"])
        ws = wb[CH]
        head = [c.value for c in ws[1]]
        self.assertEqual(head, ["Qayerga", "Dori", "Ishlab chiqaruvchi", "Ko'chirish soni", "Oluvchida sotuv", "Oluvchida qoldiq",
                                "Beruvchida sotuv", "Beruvchida qoldiq", "Kirim narxi", "Summa", "Eng yaqin muddat", "Qolgan kun",
                                "Sabab", "Holat", "ABC"])
        self.assertEqual(ws["A1"].fill.fgColor.rgb[-6:], "F5C33B", "sarlavha — toifa rangi")
        self.assertEqual([ws.cell(r, 2).value for r in range(2, 6)], ["АССОРТ 03", "АССОРТ 07", "АССОРТ 06", "АССОРТ 04"])
        self.assertEqual(ws.cell(6, 1).value, "JAMI")
        self.assertEqual(ws.cell(2, 14).value, "yangi")
        ex = wb["Ko'chirilmaydi"]
        self.assertEqual(ex.max_row, 5)
        self.assertEqual(ex.cell(1, ex.max_column).value, "Nega chiqarildi")
        reasons = {ex.cell(r, 2).value: ex.cell(r, ex.max_column).value for r in range(2, 6)}
        self.assertEqual(reasons["АССОРТ 05"], "Muddati o'tgan")
        self.assertTrue(reasons["АССОРТ 00"].startswith("Muddati 60 kundan kam"))
        row05 = [r for r in range(2, 6) if ex.cell(r, 2).value == "АССОРТ 05"][0]
        self.assertEqual(ex.cell(row05, 1).fill.fgColor.rgb[-6:], "F4B6B6", "muddati o'tgan — qizil")

    def test_excel_oddiy_turgan_kun(self):
        _, wb = self.excel(F_EK, "oddiy")
        ws = wb[EK]
        head = [c.value for c in ws[1]]
        self.assertEqual(head[4], "Turgan kun")
        self.assertEqual(ws["A1"].fill.fgColor.rgb[-6:], "30A46C")
        by = {ws.cell(r, 2).value: r for r in range(2, ws.max_row)}
        self.assertEqual(ws.cell(by["ОДДИЙ 05"], 5).value, 200)
        self.assertEqual(ws.cell(by["ОДДИЙ 05"], 1).fill.fgColor.rgb[-6:], "FBE0DC", "turgan kun ≥ 90 — och qizil")
        self.assertEqual(ws.cell(by["ОДДИЙ 01"], 1).fill.fgColor.rgb[-6:], "FFF2B3", "qolgan kun < 30 — sariq")
        self.assertEqual(ws.cell(by["ОДДИЙ 04"], 1).fill.fill_type, None)

    def test_excel_faqat_filial_xodimi(self):
        for cid in (STAFF, 9001):
            self.bot.handle(cb(cid, "kx:shoshilinch"))
            self.assertFalse(self.tg.to(cid, "doc"))
        self.bot.handle(cb(F_CH, "kx:boshqa"))
        self.assertFalse(self.tg.to(F_CH, "doc"))
        self.bot.handle(cb(F_CH, "x:kz"))  # eski xabardagi tugma
        self.assertTrue(self.tg.to(F_CH, "doc")[-1][2].startswith("SHOSHILINCH_"))

    # ---- B4.6 Keladigan, Bugungi holat ----
    def test_keladigan(self):
        text, kb = self.bot.incoming_list(CH)
        self.assertIn("Yo'lda (berilgan): <b>0</b> · kutilmoqda: <b>6</b>", text)
        self.assertIn("• 🔴 АВИТЕН", text)
        self.assertIn("• 🟢 ОДДИЙ 05", text)
        self.assertIsNone(kb)
        self.bot.handle(cb(F_EK, "kb:%d" % self.tid("АВИТЕН")))
        text, kb = self.bot.incoming_list(CH)
        self.assertIn("📦 Berildi — 📥 Oldim kutilmoqda</b> (1/1):\n1. 🔴 АВИТЕН · <b>3</b> qadoq ← " + EK, text)
        self.assertEqual([d for d in inline_data(kb) if not d.startswith("kv:")], ["ki:%d:0" % self.tid("АВИТЕН")])
        self.bot.handle(msg(F_CH, "📥 Keladigan · 6"))
        self.assertTrue(any("📥 <b>Keladigan</b>" in t for t in self.tg.texts(F_CH)))

    def test_bugungi_holat(self):
        t = self.bot.day_text(CH)
        for s in ("🔴 <b>SHOSHILINCH</b>\n📤 beradi: jami 7 · berdi 0 · qoldi 7\n📥 oladi: jami 1 · oldi 0 · qoldi 1",
                  "🟡 <b>ASSORTIMENT</b>\n📤 beradi: jami 4", "🟢 <b>ODDIY</b>\n📤 beradi: jami 1 · berdi 0 · qoldi 1\n📥 oladi: jami 5",
                  "✅ Kechikkan yo'q"):
            self.assertIn(s, t)
        self.clock.at(10, 30)
        self.assertIn("⛔ siz tomondan kechikkan: <b>7</b>", self.bot.day_text(CH))

    # ---- holatlar ----
    def test_topshiriq_shakllanmagan(self):
        self.bot.db.execute("DELETE FROM koch_tasks")
        self.bot.db.commit()
        self.assertIn("hali shakllanmagan", self.bot.koch_list_text(self.c, "oddiy")[0])
        self.assertIn("hali shakllanmagan", self.bot.filial_panel(self.c)[0])

    def test_nazorat_ochiq_excel_rejimi(self):
        self.bot.set_setting("koch_on", "0")
        self.bot.db.execute("DELETE FROM koch_tasks")
        self.bot.db.commit()
        rows, _ = self.bot.koch_view(CH, "shoshilinch")
        self.assertEqual(len(rows), 7, "nazorat o'chiq: Excel qatorlari, xuddi shu filtr")
        text, kb = self.bot.koch_list_text(self.c, "shoshilinch")
        self.assertNotIn("ki:", " ".join(inline_data(kb)), "topshiriq yo'q — kartochka tugmasi yo'q")
        self.assertIn("kx:shoshilinch", inline_data(kb))

    def test_eski_xabar_tugmalari(self):
        self.bot.handle(cb(F_CH, "kn:z:0"))
        self.assertIn("🔴 <b>SHOSHILINCH</b>", self.tg.texts(F_CH)[-1])
        self.bot.handle(cb(F_CH, "kn:o:0"))
        self.assertIn("🟢 <b>ODDIY</b>", self.tg.texts(F_CH)[-1])

    def test_admin_mijoz_korinishi(self):
        self.bot.handle(cb(9001, "v2:%d:assortiment" % F_CH))
        self.assertIn("🟡 <b>ASSORTIMENT</b>", self.tg.texts(9001)[-1])
        self.assertIn("v2:%d:shoshilinch" % F_CH, inline_data(self.tg.sent[-1][3]))
        self.assertFalse(self.tg.to(F_CH), "mijozga hech narsa yuborilmaydi")


class StyleFallback(unittest.TestCase):
    """reply_kb style + Telegram 400 bo'lsa rangsiz qayta yuborish."""

    def test_reply_kb(self):
        kb = B.reply_kb([[{"text": "A", "style": "danger"}, {"text": "B", "style": None}], ["C"]])
        self.assertEqual(kb["keyboard"], [[{"text": "A", "style": "danger"}, {"text": "B"}], [{"text": "C"}]])

    def test_400_dan_keyin_rangsiz(self):
        tg = B.TG("1:x")
        sent = []

        def fake_open(req, timeout):
            p = json.loads(req.data.decode("utf-8"))
            sent.append(p)
            if "style" in json.dumps(p):
                raise B.TGError("HTTP 400: Bad Request: can't parse keyboard button", 400)
            return {"message_id": 1}

        tg._open = fake_open
        kb = B.reply_kb([[{"text": "🔴 SHOSHILINCH · 1", "style": "danger"}]])
        self.assertEqual(tg.send(1, "salom", kb), {"message_id": 1})
        self.assertEqual(len(sent), 2)
        self.assertEqual(sent[1]["reply_markup"]["keyboard"], [[{"text": "🔴 ШОШИЛИНЧ · 1"}]], "Telegramga kirillda")
        self.assertTrue(tg.no_style)
        tg.send(1, "yana", kb)
        self.assertEqual(len(sent), 3, "keyingi safar darhol rangsiz")

    def test_boshqa_400_qayta_yuborilmaydi(self):
        tg = B.TG("1:x")
        n = []

        def fake_open(req, timeout):
            n.append(1)
            raise B.TGError("HTTP 400: Bad Request: chat not found", 400)

        tg._open = fake_open
        with self.assertRaises(B.TGError):
            tg.send(1, "x")  # klaviaturasiz
        self.assertEqual(len(n), 1)


if __name__ == "__main__":
    unittest.main()
