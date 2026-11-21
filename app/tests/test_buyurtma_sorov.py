# -*- coding: utf-8 -*-
"""10-talab (dizayn/BUYURTMA-SOROV.md): firma vakili ro'yxat → savat → «📤 Adminga yuborish»; har adminga alohida karta,
dori bo'yicha qaror shu kartaning o'zida, birinchi javob bergan hal qiladi, vakilga natija (push), kartalar qisqaradi."""
import io
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fixtures import B, BotCase, ADMIN, STAFF, FIRM, F_CH, TAG_A, TAG_B, add_client, cb, msg, inline_data  # noqa: E402
import timesaver_panel as P  # noqa: E402
import buyurtma as BY  # noqa: E402
from openpyxl import load_workbook  # noqa: E402

ADMIN2 = 9002
FIRM_B = 5003


def btn_texts(markup):
    """Tugma matnlari; ma'lumot belgilari (uzkir.mark — haqiqiy botda TG._cyr oladi) olib tashlanadi."""
    return [B.UK.strip_marks(b["text"]) for row in (markup or {}).get("inline_keyboard", []) for b in row]


def set_sale(bot, name, v):
    """Sale qatori (filiallar bo'yicha qoldiq / sotuv): v = [ch_q, ch_s, uch_q, uch_s, ek_q, ek_s, sklad_q]."""
    r = next(r for r in bot.data.sale if r["name"] == name)
    r["v"] = [float(x) for x in v]
    r["stock"] = v[0] + v[2] + v[4] + v[6]
    r["sales"] = v[1] + v[3] + v[5]
    return r


class Base(BotCase):
    def setUp(self):
        super().setUp()
        sale = self.bot.data.sale
        a = [r for r in sale if B.norm(r["firm"]) == B.norm(TAG_A)]
        b = [r for r in sale if B.norm(r["firm"]) == B.norm(TAG_B)]
        self.zero, self.low = a[0]["name"], a[1]["name"]
        set_sale(self.bot, self.zero, [0, 21, 0, 15, 0, 10, 0])     # tugagan: qoldiq 0, 46 ta sotilgan → taklif 46
        set_sale(self.bot, self.low, [3, 60, 3, 60, 2, 36, 0])      # qoldiq 8, sotuv 156 → ≈0 kun, taklif 148
        self.other = b[0]["name"]
        set_sale(self.bot, self.other, [0, 30, 0, 0, 0, 0, 0])      # boshqa firmaning tugagan dorisi
        add_client(self.bot, FIRM_B, "firm", firms=[TAG_B], name="Vakil B")

    def open(self, cid, text="/start"):
        self.bot.handle(msg(cid, text))
        return self.bot.panel_mid(cid)

    def press(self, cid, data, mid=None):
        self.bot.handle(cb(cid, data, self.bot.panel_mid(cid) if mid is None else mid))
        return self.tg.sent[-1]

    def toggle(self, cid, i):
        rows = self.bot.oreq_low(self.bot.tags_for(self.bot.client(cid)))
        return self.press(cid, "c:b:t:%d:%s:0" % (i, self.bot._ck(rows[i]["key"])))

    def send_request(self):
        self.open(FIRM)
        self.press(FIRM, "c:need")
        self.toggle(FIRM, 0)
        self.toggle(FIRM, 1)
        self.press(FIRM, "c:b:send")
        return BY.open_requests(self.bot.db)[-1]["id"]

    def card_mid(self, rid, admin=ADMIN):
        r = self.bot.db.execute("SELECT message_id FROM order_req_msgs WHERE req_id=? AND chat_id=?", (rid, admin)).fetchone()
        return r[0] if r else None


class FirmaVakili(Base):
    def test_royxat_tugagan_birinchi_va_faqat_oz_firmasi(self):
        rows = self.bot.oreq_low(self.bot.tags_for(self.bot.client(FIRM)))
        self.assertEqual([r["name"] for r in rows], [self.zero, self.low], "avval qoldiq 0, keyin ≤ 7 kun")
        self.assertEqual([r["sug"] for r in rows], [46, 148], "taklif = sotuv × cover − qoldiq (yuqoriga)")
        self.assertNotIn(self.other, [r["name"] for r in rows], "boshqa firmaning dorisi ko'rinmaydi")
        texts = btn_texts(self.bot.client_home(self.bot.client(FIRM))[1])
        self.assertIn("🧾 Buyurtma · 2", texts)
        self.open(FIRM)
        kind, _, text, kb = self.press(FIRM, "c:need")
        self.assertEqual(kind, "edit", "panelning o'zi o'zgardi")
        self.assertTrue(text.startswith("🏠 › 🧾 Buyurtma\n\n🧾 <b>Buyurtma so'rovi</b>"), text[:60])
        self.assertIn("qoldiq 0 · tugagan · 10 kunda 46", text)
        self.assertNotIn(self.other, text)
        data = inline_data(kb)
        self.assertIn("x:need", data, "eski «Buyurtma tavsiyasi» Exceli saqlanadi")
        self.assertIn("c:b:c", data)
        self.assertIn("c:b:m", data)
        self.assertEqual(data[-2:], ["c:home", "c:home"])

    def test_eski_hisobot_exceli(self):
        self.open(FIRM)
        self.press(FIRM, "c:need")
        self.press(FIRM, "x:need")
        self.assertTrue(self.tg.to(FIRM, "doc")[-1][2].startswith("Buyurtma_"))

    def test_savat_toggle_va_qayta_tushganda_saqlanadi(self):
        self.open(FIRM)
        self.press(FIRM, "c:need")
        _, _, text, kb = self.toggle(FIRM, 0)
        self.assertEqual(self.bot.cart(FIRM), [{"name": self.zero, "producer": self.bot.oreq_low({B.norm(TAG_A)})[0]["producer"],
                                                "qty": 46}])
        self.assertIn("✓ <b>%s</b>" % self.zero, text)
        self.assertIn("🛒 Savat · 1", btn_texts(kb))
        self.open(FIRM)  # /start — panel pastga qayta tushdi
        self.bot._panels.clear()
        self.assertEqual(len(self.bot.cart(FIRM)), 1, "savat bot.db da (settings) saqlanadi")
        self.toggle(FIRM, 0)
        self.assertEqual(self.bot.cart(FIRM), [], "yana bosilsa — savatdan chiqadi")

    def test_eskirgan_tugma(self):
        self.open(FIRM)
        _, _, text, _ = self.press(FIRM, "c:b:t:0:ffffff:0")
        self.assertIn("Ro'yxat yangilandi", text)
        self.assertEqual(self.bot.cart(FIRM), [])

    def test_savat_son_ozgartirish(self):
        self.open(FIRM)
        self.press(FIRM, "c:need")
        self.toggle(FIRM, 0)
        self.toggle(FIRM, 1)
        kind, _, text, kb = self.press(FIRM, "c:b:c")
        self.assertEqual(kind, "edit")
        self.assertTrue(text.startswith("🏠 › 🧾 Buyurtma › 🛒 Savat"))
        self.assertIn("<pre>", text)
        self.assertIn("▸ Tanlandi: <b>%s</b>" % self.zero, text)
        for d in ("c:b:q:0:-5", "c:b:q:0:1", "c:b:w:0", "c:b:r:0", "c:b:send", "c:b:clr"):
            self.assertIn(d, inline_data(kb))
        self.press(FIRM, "c:b:q:0:5")
        self.press(FIRM, "c:b:q:0:-1")
        self.assertEqual(self.bot.cart(FIRM)[0]["qty"], 50)
        self.press(FIRM, "c:b:c:1")
        self.press(FIRM, "c:b:q:1:-5")
        self.assertEqual(self.bot.cart(FIRM)[1]["qty"], 143)
        self.press(FIRM, "c:b:q:0:7")  # ruxsat etilmagan qadam
        self.assertEqual(self.bot.cart(FIRM)[0]["qty"], 50)
        # yozilgan son
        _, _, text, _ = self.press(FIRM, "c:b:w:0")
        self.assertIn("Nechta kerak?", text)
        self.bot.handle(msg(FIRM, "abc"))
        self.assertIn("Butun son yozing", self.tg.sent[-1][2])
        self.assertEqual(self.bot.cart(FIRM)[0]["qty"], 50)
        self.bot.handle(msg(FIRM, "30"))
        self.assertEqual(self.bot.cart(FIRM)[0]["qty"], 30)
        self.assertEqual(self.tg.sent[-1][0], "send", "yozilgan matnga javob — yangi panel")
        self.assertIn("✅ %s — 30 ta" % self.zero, self.tg.sent[-1][2])
        # minimal 1
        for _ in range(8):
            self.press(FIRM, "c:b:q:0:-5")
        self.assertEqual(self.bot.cart(FIRM)[0]["qty"], 1)
        self.press(FIRM, "c:b:r:0")
        self.assertEqual([i["name"] for i in self.bot.cart(FIRM)], [self.low])
        self.press(FIRM, "c:b:clr")
        self.assertEqual(self.bot.cart(FIRM), [])
        _, _, text, _ = self.tg.sent[-1]
        self.assertIn("Savat bo'sh", text)

    def test_yuborish_va_admin_kartasi(self):
        rid = self.send_request()
        r, items = BY.get(self.bot.db, rid)
        self.assertEqual((r["chat_id"], r["firm"], r["status"]), (FIRM, B.norm(TAG_A), "kutmoqda"))
        self.assertEqual([(i["name"], i["qty_asked"]) for i in items], [(self.zero, 46), (self.low, 148)])
        self.assertEqual(self.bot.cart(FIRM), [], "savat bo'shadi")
        kind, _, text, kb = [s for s in self.tg.sent if s[1] == FIRM][-1]
        self.assertEqual(kind, "edit")
        self.assertIn("⏳ <b>So'rov №%d yuborildi</b>" % rid, text)
        self.assertIn("2 ta dori · jami <b>194</b> ta", text)
        self.assertIn("c:b:m", inline_data(kb))
        card = [s for s in self.tg.to(ADMIN, "send") if "Buyurtma so'rovi №%d" % rid in s[2]]
        self.assertEqual(len(card), 1, "adminga alohida karta (push)")
        self.assertIsNotNone(self.card_mid(rid))
        data = inline_data(card[0][3])
        self.assertEqual(data, ["oq:%d:all" % rid, "oq:%d:i:1" % rid, "oq:%d:i:2" % rid, "oq:%d:rad" % rid, "a:home"])
        self.assertIn("1 · %s · 46 ta" % self.bot.short(self.zero, 22), btn_texts(card[0][3]))
        self.assertIn("qoldiq 0 · tugagan", card[0][2])

    def test_xavfsizlik(self):
        # savatga (eski) boshqa firmaning dorisi tushib qolgan bo'lsa ham — yuborilmaydi
        self.bot.cart_save(FIRM, [{"name": self.other, "producer": "", "qty": 5}])
        self.open(FIRM)
        _, _, text, _ = self.press(FIRM, "c:b:send")
        self.assertIn("yuboriladigan dori yo'q", text)
        self.assertEqual(BY.open_requests(self.bot.db), [])
        # xodim, filial xodimi — buyurtma so'rovi yo'q; xodimda «Buyurtma» — eski bo'lim
        for cid in (STAFF, F_CH):
            n = len(self.tg.sent)
            self.bot.handle(cb(cid, "c:b:l:0"))
            self.bot.handle(cb(cid, "c:b:send"))
            self.assertEqual(len(self.tg.sent), n, "javob yo'q")
        self.open(STAFF)
        _, _, text, _ = self.press(STAFF, "c:need")
        self.assertIn("Buyurtma tavsiyasi", text)
        # firma vakili admin tugmalarini bosa olmaydi
        rid = self.send_request()
        for d in ("oq:%d:all" % rid, "oq:%d:send" % rid, "oq:%d:rady" % rid, "oq:l"):
            self.bot.handle(cb(FIRM, d))
        r, items = BY.get(self.bot.db, rid)
        self.assertEqual(r["status"], "kutmoqda")
        self.assertTrue(all(i["status"] == "kutmoqda" for i in items))


class Admin(Base):
    def test_dori_qarori_yarim_va_yozilgan_son(self):
        rid = self.send_request()
        mid = self.card_mid(rid)
        self.bot.handle(cb(ADMIN, "oq:%d:i:1" % rid, mid))
        kind, _, text, kb = self.tg.sent[-1]
        self.assertEqual(kind, "edit")
        self.assertEqual(self.tg.calls[-1], ("editMessageText", {"chat_id": ADMIN, "message_id": mid}), "shu karta o'zgardi")
        self.assertIn("so'raldi <b>46</b> ta. Rozimisiz?", text)
        self.assertIn("<pre>", text, "filiallar bo'yicha qoldiq / sotuv")
        self.assertEqual(btn_texts(kb)[:5], ["✅ Ha, 46 ta", "½ · 23", "¾ · 35", "−5", "+5"])
        self.assertIn("oq:%d:w:1" % rid, inline_data(kb))
        self.assertIn("oq:%d:x:1" % rid, inline_data(kb))
        self.bot.handle(cb(ADMIN, "oq:%d:i:1:41" % rid, mid))  # −5
        self.assertEqual(btn_texts(self.tg.sent[-1][3])[0], "✅ Ha, 41 ta")
        self.bot.handle(cb(ADMIN, "oq:%d:s:1:23" % rid, mid))  # ½
        _, items = BY.get(self.bot.db, rid)
        self.assertEqual((items[0]["status"], items[0]["qty_ok"]), ("ozgardi", 23))
        text = self.tg.sent[-1][2]
        self.assertIn("46 → 23 ta ✏️", text, "kartaga qaytdi")
        # yozilgan son
        self.bot.handle(cb(ADMIN, "oq:%d:w:2" % rid, mid))
        self.assertIn("Nechta tasdiqlaysiz?", self.tg.sent[-1][2])
        self.bot.handle(msg(ADMIN, "100"))
        _, items = BY.get(self.bot.db, rid)
        self.assertEqual((items[1]["status"], items[1]["qty_ok"]), ("ozgardi", 100))
        self.assertEqual(self.tg.calls[-1], ("editMessageText", {"chat_id": ADMIN, "message_id": mid}), "karta joyida")
        self.assertIn("oq:%d:send" % rid, inline_data(self.tg.sent[-1][3]), "hammasi hal — 📤 Javobni yuborish")
        self.bot.handle(cb(ADMIN, "oq:%d:send" % rid, mid))
        r, _ = BY.get(self.bot.db, rid)
        self.assertEqual((r["status"], r["resolved_by"]), ("javob", ADMIN))
        push = self.tg.last(FIRM, "So'rov №%d bo'yicha javob" % rid)
        self.assertTrue(push, "vakilga natija")
        self.assertIn("46 → 23", push[2])
        self.assertIn("148 → 100", push[2])
        self.assertIn("46 emas, <b>23</b> ta qabul qilindi", push[2])
        self.assertEqual(self.bot.panel_mid(FIRM), [m for c, m in self.tg.mids if c == FIRM][-1], "push — vakilning yangi paneli")
        self.assertIn("c:b:m", inline_data(push[3]))
        # karta bir qatorga qisqardi («Siz»), panel pastga tushdi
        collapsed = [s for s in self.tg.sent if s[0] == "edit" and s[1] == ADMIN and s[2].startswith("✅ №%d" % rid)]
        self.assertTrue(collapsed and collapsed[-1][3] is None)
        self.assertIn("Siz", collapsed[-1][2])
        self.assertIn("2 ta o'zgartirildi", collapsed[-1][2])
        self.assertIn("Javob yuborildi", self.tg.sent[-1][2])
        self.assertEqual(self.tg.sent[-1][0], "send")

    def test_olib_tashlash_va_hammasini_qabul(self):
        rid = self.send_request()
        mid = self.card_mid(rid)
        self.bot.handle(cb(ADMIN, "oq:%d:x:2" % rid, mid))
        self.bot.handle(cb(ADMIN, "oq:%d:all" % rid, mid))
        _, items = BY.get(self.bot.db, rid)
        self.assertEqual([(i["status"], i["qty_ok"]) for i in items], [("qabul", 46), ("rad", 0)])
        self.bot.handle(cb(ADMIN, "oq:%d:send" % rid, mid))
        push = self.tg.last(FIRM, "bo'yicha javob")
        self.assertIn("qabul qilinmadi", push[2])

    def test_yozilgan_son_notogri(self):
        rid = self.send_request()
        mid = self.card_mid(rid)
        self.bot.handle(cb(ADMIN, "oq:%d:w:1" % rid, mid))
        self.bot.handle(msg(ADMIN, "2.5"))
        self.assertIn("Butun son yozing", self.tg.sent[-1][2])
        self.assertEqual(self.bot.state[ADMIN][0], "oreq_qty", "son kutilmoqda")
        self.bot.handle(msg(ADMIN, "0"))
        _, items = BY.get(self.bot.db, rid)
        self.assertEqual(items[0]["status"], "rad", "0 — olib tashlash")

    def test_rad_etish(self):
        rid = self.send_request()
        mid = self.card_mid(rid)
        self.bot.handle(cb(ADMIN, "oq:%d:rad" % rid, mid))
        self.assertIn("rad etilsinmi", self.tg.sent[-1][2])
        self.assertEqual(BY.get(self.bot.db, rid)[0]["status"], "kutmoqda", "tasdiqsiz rad etilmaydi")
        self.bot.handle(cb(ADMIN, "oq:%d:rady" % rid, mid))
        self.assertEqual(BY.get(self.bot.db, rid)[0]["status"], "rad")
        self.assertTrue(self.tg.last(FIRM, "So'rov №%d rad etildi" % rid))

    def test_birinchi_admin_yutadi_va_kartalar_qisqaradi(self):
        self.bot.admins = {ADMIN, ADMIN2}
        rid = self.send_request()
        m1, m2 = self.card_mid(rid, ADMIN), self.card_mid(rid, ADMIN2)
        self.assertTrue(m1 and m2 and m1 != m2, "har adminga alohida karta")
        self.bot.handle(cb(ADMIN, "oq:%d:all" % rid, m1))
        self.bot.handle({"callback_query": {"id": "1", "from": {"id": ADMIN, "first_name": "Sardor"}, "data": "oq:%d:send" % rid,
                                            "message": {"message_id": m1}}})
        r, _ = BY.get(self.bot.db, rid)
        self.assertEqual((r["resolved_by"], r["resolved_name"]), (ADMIN, "Sardor"))
        e2 = [s for s in self.tg.sent if s[0] == "edit" and s[1] == ADMIN2]
        self.assertTrue(e2 and e2[-1][2].startswith("✅ №%d" % rid) and e2[-1][3] is None, "ikkinchi admin kartasi qisqardi")
        self.assertIn("Sardor", e2[-1][2])
        self.assertNotIn("Siz", e2[-1][2])
        # ikkinchi admin eski kartadagi tugmani bosdi — hech narsa o'zgarmaydi, karta qisqa holda
        n_push = len([s for s in self.tg.sent if s[1] == FIRM and "bo'yicha javob" in s[2]])
        for d in ("oq:%d:i:1" % rid, "oq:%d:s:1:5" % rid, "oq:%d:send" % rid, "oq:%d:rady" % rid):
            self.bot.handle(cb(ADMIN2, d, m2))
            self.assertTrue(self.tg.sent[-1][2].startswith("✅ №%d" % rid), d)
        r, items = BY.get(self.bot.db, rid)
        self.assertEqual((r["status"], items[0]["qty_ok"]), ("javob", 46))
        self.assertEqual(len([s for s in self.tg.sent if s[1] == FIRM and "bo'yicha javob" in s[2]]), n_push, "vakilga bitta javob")

    def test_order_notify_panel_orqali(self):
        rid = self.send_request()
        api = P.Api(self.bot, self.bot.cfg)
        self.assertTrue(api.oreq_accept_all({"id": rid})["ok"])
        res = api.oreq_finish({"id": rid})
        self.assertEqual((res["ok"], res["msg"]), (True, "Javob yuborildi"))
        self.assertTrue(self.tg.last(FIRM, "So'rov №%d bo'yicha javob" % rid))
        e = [s for s in self.tg.sent if s[0] == "edit" and s[1] == ADMIN and s[2].startswith("✅ №%d" % rid)]
        self.assertTrue(e, "admin kartasi qisqardi")
        self.assertIn("Админ (панел)", e[-1][2])
        self.assertFalse(self.bot.order_notify(rid), "ikkinchi marta yuborilmaydi")
        self.assertFalse(self.bot.order_notify(999))

    def test_admin_royxati_va_excel(self):
        rid = self.send_request()
        kb = self.bot.admin_panel(ADMIN)[1]
        self.assertIn("🧾 Buyurtma so'rovlari · 1", btn_texts(kb))
        self.assertIn("🧾 Buyurtma so'rovlari: <b>1</b> ta kutmoqda", self.bot.admin_panel(ADMIN)[0])
        self.open(ADMIN)
        kind, _, text, kb = self.press(ADMIN, "oq:l")
        self.assertEqual(kind, "edit")
        self.assertTrue(text.startswith("🏠 › 🧾 Buyurtma so'rovlari"))
        self.assertIn("№%d · %s · 2 ta dori" % (rid, B.norm(TAG_A)), text)
        self.assertIn("oq:%d" % rid, inline_data(kb))
        kind, _, text, kb = self.press(ADMIN, "oq:%d" % rid)
        self.assertTrue(text.startswith("🏠 › 🧾 Buyurtma so'rovlari › №%d" % rid), "panelda — yo'l qatori bilan")
        self.assertEqual(inline_data(kb)[-2:], ["oq:l", "a:home"])
        self.press(ADMIN, "oq:lx")
        self.assertIn("tasdiqlangan so'rov yo'q", self.tg.sent[-1][2])
        self.press(ADMIN, "oq:%d:all" % rid)
        self.press(ADMIN, "oq:%d:send" % rid)
        self.assertTrue(self.tg.sent[-1][2].startswith("🏠 › 🧾 Buyurtma so'rovlari"), "panelda — ro'yxatga qaytadi")
        self.press(ADMIN, "oq:lx")
        fname, content = self.tg.to(ADMIN, "doc")[-1][2:4]
        self.assertTrue(fname.startswith("Buyurtma_sorovlari_7kun_"))
        ws = load_workbook(io.BytesIO(content)).active
        self.assertEqual(ws.max_row, 3)

    def test_sorovlarim_va_excel(self):
        rid = self.send_request()
        mid = self.card_mid(rid)
        self.bot.handle(cb(ADMIN, "oq:%d:s:1:10" % rid, mid))
        self.bot.handle(cb(ADMIN, "oq:%d:s:2:148" % rid, mid))
        self.bot.handle(cb(ADMIN, "oq:%d:send" % rid, mid))
        # boshqa firmaning so'rovi — vakil B
        BY.create(self.bot.db, FIRM_B, B.norm(TAG_B), "Vakil B", [{"name": self.other, "qty": 3}])
        kind, _, text, kb = self.press(FIRM, "c:b:m")
        self.assertEqual(kind, "edit", "push'dagi tugma — u endi panel")
        self.assertTrue(text.startswith("🏠 › 🧾 Buyurtma › 📋 So'rovlarim"))
        self.assertIn("<b>1</b> ta so'rov", text)
        self.assertIn("✏️", text)
        self.press(FIRM, "c:b:mx")
        ws = load_workbook(io.BytesIO(self.tg.to(FIRM, "doc")[-1][3])).active
        names = [ws.cell(r, 3).value for r in range(2, ws.max_row + 1)]
        self.assertEqual(names, [self.zero, self.low], "faqat o'z so'rovlari")


if __name__ == "__main__":
    unittest.main()
