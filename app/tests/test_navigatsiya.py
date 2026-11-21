# -*- coding: utf-8 -*-
"""Ichma-ich navigatsiya (dizayn/bot/NAVIGATSIYA.md): bitta «tirik» panel xabari, yo'l qatori, [◀ Orqaga] [🏠 Menyu],
pastki klaviatura yo'q (faqat kontakt so'rash), Excel / yozilgan matndan keyin panel pastga qayta tushadi."""
import json
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fixtures import (B, BotCase, ADMIN, STAFF, FIRM, F_CH, CH, TAG_A, add_client, cb, msg, inline_data)  # noqa: E402
import uzkir as UK  # noqa: E402


def btn_texts(markup):
    return [b["text"] for row in (markup or {}).get("inline_keyboard", []) for b in row]


def last_row(markup):
    return [(b["text"], b["callback_data"]) for b in markup["inline_keyboard"][-1]]


class Base(BotCase):
    def open(self, cid, text="/start"):
        """Yangi panel (yozilgan matn) → panel xabarining id si."""
        self.bot.handle(msg(cid, text))
        return self.bot.panel_mid(cid)

    def press(self, cid, data, mid=None):
        """Panel xabaridagi tugmani bosish (mid — hozirgi panel)."""
        self.bot.handle(cb(cid, data, self.bot.panel_mid(cid) if mid is None else mid))
        return self.tg.sent[-1]

    def sends(self, cid):
        return self.tg.to(cid, "send")


class TelegramDelete(unittest.TestCase):
    def test_delete_xato_kotarmaydi(self):
        tg = B.TG("1:x")
        calls = []

        def fake_open(req, timeout):
            calls.append(json.loads(req.data.decode("utf-8")))
            raise B.TGError("HTTP 400: Bad Request: message can't be deleted", 400)

        tg._open = fake_open
        self.assertFalse(tg.delete(5, 77))
        self.assertEqual(calls[0], {"chat_id": 5, "message_id": 77})
        self.assertTrue(tg.url.endswith("/"), "deleteMessage usuli")

    def test_inline_style_va_rangsiz_qayta(self):
        kb = B.inline_kb([[("🔴 A", "x", "danger"), ("B", "y", None)], [("C", "z")]])
        self.assertEqual(kb["inline_keyboard"][0], [{"text": "🔴 A", "callback_data": "x", "style": "danger"},
                                                    {"text": "B", "callback_data": "y"}])
        tg = B.TG("1:x")
        sent = []

        def fake_open(req, timeout):
            p = json.loads(req.data.decode("utf-8"))
            sent.append(p)
            if "style" in json.dumps(p):
                raise B.TGError("HTTP 400: Bad Request: can't parse inline keyboard button", 400)
            return {"message_id": 1}

        tg._open = fake_open
        tg.send(1, "x", kb)
        self.assertEqual(len(sent), 2)
        self.assertNotIn("style", json.dumps(sent[1]))


class BittaPanel(Base):
    def test_tugma_bosilsa_shu_xabar_ozgaradi(self):
        mid = self.open(FIRM)
        n_send = len(self.sends(FIRM))
        kind, cid, text, kb = self.press(FIRM, "c:srok")
        self.assertEqual(kind, "edit", "yangi xabar emas — panel joyida o'zgardi")
        self.assertEqual(self.tg.calls[-1], ("editMessageText", {"chat_id": FIRM, "message_id": mid}))
        self.assertTrue(text.startswith("🏠 › ⏰ Srok\n\n"), text[:40])
        self.assertIn("x:srok", inline_data(kb))
        self.assertEqual(last_row(kb), [("◀ Orqaga", "c:home"), ("🏠 Menyu", "c:home")])
        kind, _, text, kb = self.press(FIRM, "c:home")
        self.assertEqual(kind, "edit")
        self.assertTrue(text.startswith("🏠 <b>Bosh menyu</b>"), "bosh menyuda yo'l qatori va navigatsiya yo'q")
        self.assertNotIn("◀ Orqaga", btn_texts(kb))
        self.assertEqual(len(self.sends(FIRM)), n_send, "chatga yangi xabar tushmadi")
        self.assertEqual(self.tg.deleted, [])
        self.assertEqual(self.bot.panel_mid(FIRM), mid)

    def test_panel_id_saqlanadi(self):
        mid = self.open(FIRM)
        self.assertEqual(self.bot.setting("panel_%d" % FIRM), str(mid))
        self.bot._panels.clear()  # bot qayta ishga tushdi
        self.assertEqual(self.press(FIRM, "c:old", mid)[0], "edit")

    def test_eski_xabardagi_tugma_panelni_pastga_tushiradi(self):
        mid = self.open(FIRM)
        self.bot.handle(cb(FIRM, "c:srok", 55))  # eski xabar (panel emas)
        self.assertEqual(self.tg.sent[-1][0], "send")
        self.assertIn((FIRM, mid), self.tg.deleted, "eski panel o'chirildi")
        self.assertNotEqual(self.bot.panel_mid(FIRM), mid)

    def test_excel_dan_keyin_panel_fayl_ostida(self):
        mid = self.open(FIRM)
        self.press(FIRM, "c:srok")
        self.press(FIRM, "x:srok")
        kinds = [s[0] for s in self.tg.to(FIRM)]
        self.assertEqual(kinds[-2:], ["doc", "send"], "avval fayl, keyin panel")
        self.assertTrue(self.tg.to(FIRM, "doc")[-1][4].startswith("⏰ <b>Srogi"), "fayl izohi o'zgarmagan")
        text = self.tg.sent[-1][2]
        self.assertTrue(text.startswith("🏠 › ⏰ Srok\n\n📎 Excel yuborildi"), text[:60])
        self.assertIn((FIRM, mid), self.tg.deleted)
        self.assertEqual(self.bot.panel_mid(FIRM), self.tg.mids[-1][1])

    def test_yozilgan_matn_yangi_panel(self):
        first = self.open(FIRM)
        second = self.open(FIRM, "ШОШ ДОРИ")
        kind, _, text, kb = self.tg.sent[-1]
        self.assertEqual(kind, "send")
        self.assertTrue(text.startswith("🏠 › 🔍 Dori qidirish › «ШОШ ДОРИ»\n\n🔍"), text[:60])
        self.assertEqual(last_row(kb), [("🔍 Boshqa dori", "c:find"), ("🏠 Menyu", "c:home")])
        self.assertEqual(self.tg.deleted, [(FIRM, first)], "chatda bitta panel")
        self.assertNotEqual(first, second)

    def test_qidiruv_tugmasi_soraydi(self):
        self.open(FIRM)
        kind, _, text, kb = self.press(FIRM, "c:find")
        self.assertEqual(kind, "edit")
        self.assertIn("Dori nomini yozing", text)
        self.assertEqual(last_row(kb), [("◀ Orqaga", "c:home"), ("🏠 Menyu", "c:home")])

    def test_hisob_biriktirilmagan_panelda(self):
        add_client(self.bot, 7400, "firm", firms=["ГАММА 900000000"])
        self.open(7400)
        kind, _, text, kb = self.press(7400, "c:acc")
        self.assertEqual(kind, "edit")
        self.assertIn("🏠 › 💰 Hisob", text)
        self.assertIn("adminga yuborildi", text)
        self.assertTrue(self.tg.last(ADMIN, "Hisob so'rovi"), "adminga bildirishnoma")
        self.assertIn("a:home", inline_data(self.tg.last(ADMIN, "Hisob so'rovi")[3]))


class PastkiKlaviatura(Base):
    def test_remove_keyboard_bir_marta(self):
        self.open(FIRM)
        rm = [s for s in self.sends(FIRM) if s[3] == {"remove_keyboard": True}]
        self.assertEqual(len(rm), 1)
        self.assertLess(self.sends(FIRM).index(rm[0]), len(self.sends(FIRM)) - 1, "panel undan keyin")
        self.open(FIRM)
        self.press(FIRM, "c:srok", 12345)  # eski xabardan — panel qayta yuboriladi
        self.assertEqual(len([s for s in self.sends(FIRM) if s[3] == {"remove_keyboard": True}]), 1, "faqat birinchi marta")
        self.assertEqual(self.bot.setting("kbclean_%d" % FIRM), "1")
        self.bot._kbclean.clear()  # qayta ishga tushgandan keyin ham
        self.open(FIRM)
        self.assertEqual(len([s for s in self.sends(FIRM) if s[3] == {"remove_keyboard": True}]), 1)

    def test_pastki_klaviatura_yuborilmaydi(self):
        self.bot.koch_build()
        for cid, data in ((FIRM, ("c:srok", "c:acc", "c:find", "c:home")), (STAFF, ("a:fm", "c:need", "c:home")),
                          (F_CH, ("kp:shoshilinch:0", "kl:0", "kd", "kh")), (ADMIN, ("a:sec:srok", "a:usr", "a:bc", "a:set", "a:home"))):
            self.open(cid)
            for d in data:
                self.press(cid, d)
        self.bot.handle(msg(7300, "/start"))  # yangi foydalanuvchi: yagona pastki klaviatura — kontakt
        reply = [s for s in self.tg.sent if s[0] == "send" and s[3] and "keyboard" in s[3]]
        self.assertEqual(len(reply), 1)
        self.assertEqual(reply[0][1], 7300)
        self.assertTrue(reply[0][3]["keyboard"][0][0]["request_contact"])

    def test_yangi_foydalanuvchi_ikkinchi_marta_tozalanmaydi(self):
        self.bot.handle(msg(7301, None, contact={"user_id": 7301, "phone_number": "998935550011"}))
        self.assertEqual(self.tg.sent[-1][3], {"remove_keyboard": True}, "kontakt klaviaturasi olindi")
        self.bot.set_firms(7301, [TAG_A])
        self.bot.activate(7301)
        self.bot.handle(cb(7301, "c:home", self.tg.mids[-1][1]))  # «🏠 Menyuni ochish»
        self.assertEqual(len([s for s in self.tg.to(7301, "send") if s[3] == {"remove_keyboard": True}]), 1)


class EskiTugmalar(Base):
    def setUp(self):
        super().setUp()
        UK.KB_BACK.clear()  # bot qayta ishga tushgan: eski kirill tugmalar xotirada yo'q

    def test_mijoz(self):
        self.bot.handle(msg(FIRM, "⏰ Срок · 3"))
        self.assertTrue(self.tg.sent[-1][2].startswith("🏠 › ⏰ Srok"))
        self.bot.handle(msg(FIRM, "💰 Ҳисоб · 500 000 кредит"))
        self.assertTrue(self.tg.sent[-1][2].startswith("🏠 › 💰 Hisob"))
        self.bot.handle(msg(FIRM, "📦 30+ kun · 2"))  # lotin ham
        self.assertTrue(self.tg.sent[-1][2].startswith("🏠 › 📦 30+ kun"))
        self.bot.handle(msg(FIRM, "🏠 Меню"))
        self.assertTrue(self.tg.sent[-1][2].startswith("🏠 <b>Bosh menyu</b>"))

    def test_xodim_firma_tugmasi(self):
        self.bot.handle(msg(STAFF, "🏢 Фирма: ҳаммаси"))
        self.assertTrue(self.tg.sent[-1][2].startswith("🏠 › 🏢 Firma"))
        self.assertEqual(self.bot.state[STAFF][0], "firm_pick")

    def test_filial(self):
        self.bot.koch_build()
        self.bot.handle(msg(F_CH, "🔴 ШОШИЛИНЧ · 7"))
        self.assertTrue(self.tg.sent[-1][2].startswith("🏠 › 🔴 SHOSHILINCH\n\n🔴 <b>SHOSHILINCH</b>"))
        self.bot.handle(msg(F_CH, "📋 Бугунги ҳолат"))
        self.assertTrue(self.tg.sent[-1][2].startswith("🏠 › 📋 Bugungi holat"))
        self.bot.handle(msg(F_CH, "🏠 Меню"))
        self.assertTrue(self.tg.sent[-1][2].startswith("🏠 <b>Bosh menyu</b>\n🏬"))

    def test_admin(self):
        self.bot.handle(msg(ADMIN, "👥 Фойдаланувчилар"))
        self.assertTrue(self.tg.sent[-1][2].startswith("🏠 › 👥 Foydalanuvchilar"))
        self.bot.handle(msg(ADMIN, "⏰ Срок"))
        self.assertTrue(self.tg.sent[-1][2].startswith("🏠 › ⏰ Srok"))
        self.bot.handle(msg(ADMIN, "🏢 Фирма: ҳаммаси"))
        self.assertTrue(self.tg.sent[-1][2].startswith("🏠 › 🏢 Firma"))
        self.assertEqual(self.tg.sent[-1][0], "send", "yozilgan matn — yangi panel")


class BoshMenyular(Base):
    def test_firma_vakili(self):
        data = inline_data(self.bot.client_home(self.bot.client(FIRM))[1])
        self.assertEqual(data, ["c:acc", "c:srok", "c:old", "c:need", "c:stock", "c:find"])
        self.bot.db.execute("INSERT INTO client_accounts VALUES(?,?)", (FIRM, "ООО АЛЬФА 901112233"))  # tasdiqlangan hisob
        self.bot.db.commit()
        self.open(FIRM)
        texts = btn_texts(self.tg.sent[-1][3])
        self.assertTrue([t for t in texts if t.startswith("⏰ Srok · ")], "son tugmada")
        self.assertTrue([t for t in texts if t.startswith("💰 Hisob · ") and "kredit" in t])

    def test_xodim(self):
        data = inline_data(self.bot.client_home(self.bot.client(STAFF))[1])
        self.assertEqual(data, ["c:acc", "c:srok", "c:old", "c:need", "c:stock", "c:find", "a:fm"])
        self.assertEqual(btn_texts(self.bot.client_home(self.bot.client(STAFF))[1])[-1], "🏢 Firma: hammasi")

    def test_filial(self):
        self.bot.koch_build()
        kb = self.bot.client_home(self.bot.client(F_CH))[1]
        self.assertEqual(inline_data(kb), ["kp:shoshilinch:0", "kp:assortiment:0", "kp:oddiy:0", "kl:0", "kd", "kc:s", "kc:o", "kh"])
        self.assertEqual(kb["inline_keyboard"][0][0], {"text": "🔴 SHOSHILINCH · 7", "callback_data": "kp:shoshilinch:0",
                                                       "style": "danger"})
        self.assertEqual(btn_texts(kb)[-1], "🔄 Yangilash")

    def test_sklad(self):
        from test_sklad import SK, add_sklad
        add_sklad(self.bot)
        add_client(self.bot, 6100, "filial", filial=SK)
        self.bot.koch_build()
        data = inline_data(self.bot.client_home(self.bot.client(6100))[1])
        self.assertIn("kt", data)
        self.assertNotIn("kl:0", data)
        self.open(6100)
        kind, _, text, kb = self.press(6100, "kt")
        self.assertEqual(kind, "edit")
        self.assertTrue(text.startswith("🏠 › 🏭 Taqsimot\n\n🏭 <b>Sklad taqsimoti</b>"))
        self.assertIn("kt:x", inline_data(kb))
        self.press(6100, "kt:x")
        self.assertTrue(self.tg.to(6100, "doc")[-1][2].startswith("Sklad_taqsimot_"))
        self.assertEqual(self.tg.sent[-1][0], "send", "panel fayl ostida")

    def test_admin(self):
        add_client(self.bot, 7200, "firm", status="pending")
        kb = self.bot.admin_panel(ADMIN)[1]
        data = inline_data(kb)
        self.assertEqual(data[0], "ul:pending:0", "kutayotgan so'rovlar — birinchi qator")
        for d in ("a:sec:srok", "a:sec:old", "a:sec:need", "a:sec:acc", "a:fm", "a:rep", "a:tr", "a:usr", "a:bc", "a:sk",
                  "a:stat", "a:log", "a:set", "a:ref", "a:mt"):
            self.assertIn(d, data)
        texts = btn_texts(kb)
        self.assertIn("🏢 Firma: hammasi", texts)
        self.assertTrue([t for t in texts if t.startswith("⏰ Srok · ")])
        self.open(ADMIN)
        self.assertEqual(len([s for s in self.sends(ADMIN) if s[3] and "inline_keyboard" in s[3]]), 1, "bitta panel xabari")

    def test_bolim_ichma_ich(self):
        self.bot.koch_build()
        self.open(F_CH)
        kind, _, text, kb = self.press(F_CH, "kp:shoshilinch:0")
        self.assertEqual(kind, "edit")
        self.assertEqual(last_row(kb), [("◀ Orqaga", "kh"), ("🏠 Menyu", "kh")])
        route = [d for d in inline_data(kb) if d.startswith("kg:")][0]
        kind, _, text, kb = self.press(F_CH, route)
        self.assertEqual(kind, "edit")
        self.assertRegex(text, r"^🏠 › 🔴 SHOSHILINCH › → ДОРИХОНА")
        self.assertEqual(last_row(kb), [("◀ Orqaga", "kp:shoshilinch:0"), ("🏠 Menyu", "kh")])
        tid = [d for d in inline_data(self.press(F_CH, "kp:shoshilinch:0")[3]) if d.startswith("ki:")][0]
        kind, _, text, kb = self.press(F_CH, tid)
        self.assertRegex(text, r"^🏠 › 🔴 SHOSHILINCH › #\d+\n\n")
        self.assertEqual(last_row(kb), [("◀ Orqaga", "kp:shoshilinch:0"), ("🏠 Menyu", "kh")])
        self.assertEqual(len(self.sends(F_CH)), 2, "remove_keyboard + bitta panel; qolgani — tahrir")

    def test_filial_son_kiritish_yangi_panel(self):
        self.bot.koch_build()
        tid = self.bot.db.execute("SELECT id FROM koch_tasks WHERE src=? AND status='yangi' AND qty_plan>1", (CH,)).fetchone()[0]
        mid = self.open(F_CH)
        kind, _, text, kb = self.press(F_CH, "kq:%d" % tid)
        self.assertEqual(kind, "edit")
        self.assertIn("Nechta berdingiz", text)
        self.assertEqual(last_row(kb), [("◀ Orqaga", "ki:%d:0" % tid), ("🏠 Menyu", "kh")])
        self.bot.handle(msg(F_CH, "1"))
        kind, _, text, _ = self.tg.sent[-1]
        self.assertEqual(kind, "send")
        self.assertIn("✅ Berildi:", text)
        self.assertIn((F_CH, mid), self.tg.deleted)


class AdminFirmaFiltri(Base):
    def test_ichma_ich_firma_tanlash(self):
        self.open(ADMIN)
        kind, _, text, kb = self.press(ADMIN, "a:sec:srok")
        self.assertEqual(kind, "edit")
        self.assertIn("a:fm:srok", inline_data(kb))
        kind, _, text, kb = self.press(ADMIN, "a:fm:srok")
        self.assertEqual(kind, "edit")
        self.assertTrue(text.startswith("🏠 › ⏰ Srok › 🏢 Firma\n\n"), text[:50])
        self.assertEqual(last_row(kb), [("◀ Orqaga", "a:sec:srok"), ("🏠 Menyu", "a:home")])
        self.assertEqual(self.bot.state[ADMIN][0], "firm_pick")
        prompt = self.bot.panel_mid(ADMIN)
        self.bot.handle(msg(ADMIN, "А"))  # ikkala firma belgisida bor: ro'yxat — yangi panel
        kind, _, text, kb = self.tg.sent[-1]
        self.assertEqual(kind, "send")
        self.assertIn((ADMIN, prompt), self.tg.deleted)
        self.assertIn("Topildi: <b>2</b>", text)
        self.assertEqual(last_row(kb), [("◀ Orqaga", "a:fm:srok"), ("🏠 Menyu", "a:home")])
        i = [B.norm(t) for t in self.bot.fsearch[ADMIN]].index(TAG_A)
        kind, _, text, kb = self.press(ADMIN, "fp:%d:%s" % (i, self.bot._ck(self.bot.fsearch[ADMIN][i])))
        self.assertEqual(kind, "edit")
        self.assertEqual(self.bot.fm(ADMIN), ({TAG_A}, TAG_A))
        self.assertTrue(text.startswith("🏠 › ⏰ Srok\n\n✅ <b>Firma tanlandi:</b>"), "tanlangach — kelgan bo'limga")
        self.assertIn("🏢 <b>%s</b>" % TAG_A, text)
        self.assertIn("🏢 Firma: " + self.bot.short(TAG_A, 16), btn_texts(kb))
        self.assertNotIn(ADMIN, self.bot.state)
        kind, _, text, kb = self.press(ADMIN, "a:fm:srok")
        self.assertIn("fp:off", inline_data(kb))
        kind, _, text, kb = self.press(ADMIN, "fp:off")
        self.assertEqual(self.bot.fm(ADMIN), (None, None))
        self.assertIn("hamma firmalar", text)
        self.assertTrue(text.startswith("🏠 › ⏰ Srok"))

    def test_bitta_topilsa_darhol_tanlanadi_va_hisobotlar(self):
        self.open(ADMIN)
        self.press(ADMIN, "a:rep")
        kind, _, text, kb = self.press(ADMIN, "a:fm:rep")
        self.assertTrue(text.startswith("🏠 › 📊 Hisobotlar › 🏢 Firma"))
        self.bot.handle(msg(ADMIN, "АЛЬФА"))
        text = self.tg.sent[-1][2]
        self.assertTrue(text.startswith("🏠 › 📊 Hisobotlar\n\n✅ <b>Firma tanlandi:</b>"), text[:60])
        self.assertIn("Excel fayl: 🏢 <b>%s</b>" % TAG_A, text)
        self.press(ADMIN, "a:r:srok")
        self.assertTrue(self.tg.to(ADMIN, "doc"))

    def test_xodim_firma_filtri(self):
        self.open(STAFF)
        kind, _, text, kb = self.press(STAFF, "c:old")
        self.assertIn("a:fm:old", inline_data(kb))
        kind, _, text, kb = self.press(STAFF, "a:fm:old")
        self.assertTrue(text.startswith("🏠 › 📦 30+ kun › 🏢 Firma"))
        self.assertEqual(last_row(kb), [("◀ Orqaga", "c:old"), ("🏠 Menyu", "c:home")])
        self.bot.handle(msg(STAFF, "АЛЬФА"))
        self.assertTrue(self.tg.sent[-1][2].startswith("🏠 › 📦 30+ kun\n\n✅ <b>Firma tanlandi:</b>"))
        self.assertEqual(self.bot.fm(STAFF)[1], TAG_A)

    def test_boshqa_rollar_admin_tugmalariga_kira_olmaydi(self):
        for cid in (FIRM, F_CH):
            for d in ("a:sec:srok", "a:fm:srok", "a:fm", "fp:off", "a:usr", "a:rep", "a:r:srok", "uc:%d" % STAFF, "a:home"):
                self.bot.handle(cb(cid, d))
        self.assertFalse(self.tg.to(FIRM))
        self.assertFalse(self.tg.to(F_CH))
        self.bot.handle(cb(STAFF, "a:sec:srok"))
        self.bot.handle(cb(STAFF, "a:usr"))
        self.assertFalse(self.tg.to(STAFF), "xodim — faqat firma filtri, admin bo'limlari yo'q")


class Bildirishnomalar(Base):
    def test_yangi_sorov_va_natija(self):
        self.open(ADMIN)
        panel = self.bot.panel_mid(ADMIN)
        self.bot.handle(msg(7500, None, contact={"user_id": 7500, "phone_number": "998901112233"}))
        push = [m for c, m in self.tg.mids if c == ADMIN][-1]
        self.assertIn("a:home", inline_data(self.tg.last(ADMIN, "Yangi so'rov")[3]))
        self.bot.handle(cb(ADMIN, "ap:7500", push))
        self.assertEqual(self.bot.client(7500)["status"], "active")
        edits = [s for s in self.tg.to(ADMIN, "edit") if "Qabul qilindi" in s[2]]
        self.assertTrue(edits and edits[-1][3] is None, "bildirishnoma bir qatorli natijaga aylandi")
        self.assertEqual(self.tg.sent[-1][0], "send", "panel pastga tushdi")
        self.assertIn((ADMIN, panel), self.tg.deleted)
        push_c = self.tg.last(7500, "qabul qilindingiz")
        self.assertEqual(inline_data(push_c[3]), ["c:home"], "mijozga: bitta «🏠 Menyuni ochish»")

    def test_ertalabki_topshiriq_ochish(self):
        self.bot.koch_build()
        self.bot._koch_day_start()
        push = self.tg.last(F_CH, "topshirig'i tayyor")
        self.assertEqual(inline_data(push[3]), ["kh"])
        pmid = [m for c, m in self.tg.mids if c == F_CH][-1]
        self.bot.handle(cb(F_CH, "kh", pmid))
        self.assertEqual(self.tg.sent[-1][0], "send", "panel bildirishnoma ostida yangidan")
        self.assertTrue(self.tg.sent[-1][2].startswith("🏠 <b>Bosh menyu</b>"))
        self.assertFalse([s for s in self.tg.to(F_CH, "edit")], "bildirishnoma o'zgarmadi")

    def test_eslatma_admin(self):
        self.bot._admins_send("⛔ test")
        self.assertEqual(inline_data(self.tg.last(ADMIN, "test")[3]), ["a:home"])


if __name__ == "__main__":
    unittest.main()
