# -*- coding: utf-8 -*-
"""Buyurtma so'rovi tuzatishlari: Telegram tugmalari chegarasi (savat ≤ 40, sahifalash), tugmadagi ma'lumot kirillga
o'girilmaydi, «nan» / «inf» son emas, o'qish va qaror orasida yopilgan so'rov."""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fixtures import B, ADMIN, FIRM, cb, msg, inline_data  # noqa: E402
import test_buyurtma_sorov as T  # noqa: E402  (Base — testsiz yordamchi sinf)
import timesaver_panel as P  # noqa: E402
import buyurtma as BY  # noqa: E402

btn_texts = T.btn_texts


def n_buttons(markup):
    return sum(len(row) for row in (markup or {}).get("inline_keyboard", []))


class Base(T.Base):
    def big_request(self, n=45, name="DORI %02d"):
        with self.bot.dblock:
            return BY.create(self.bot.db, FIRM, "FIRMA", "Vakil", [{"name": name % i, "qty": i} for i in range(1, n + 1)],
                             ts=B.now().isoformat(timespec="seconds"))


class Tugmalar(Base):
    def test_savat_40_tadan_oshmaydi(self):
        self.bot.cart_save(FIRM, [{"name": "X %d" % i, "producer": "", "qty": 1} for i in range(self.bot.ORD_CART_MAX)])
        self.open(FIRM)
        self.press(FIRM, "c:need")
        _, _, text, _ = self.toggle(FIRM, 0)
        self.assertIn("Savat to'ldi", text)
        self.assertEqual(len(self.bot.cart(FIRM)), 40)

    def test_savat_sahifalanadi(self):
        self.bot.cart_save(FIRM, [{"name": "X %d" % i, "producer": "", "qty": 1} for i in range(25)])
        self.open(FIRM)
        _, _, _, kb = self.press(FIRM, "c:b:c:0")
        data = inline_data(kb)
        self.assertIn("c:b:c:9", data)
        self.assertNotIn("c:b:c:11", data, "faqat 1-sahifa (1–10)")
        self.assertIn("c:b:c:10", data, "▶ — keyingi sahifaning birinchisi")
        self.assertIn("1 / 3", btn_texts(kb))
        _, _, text, kb = self.press(FIRM, "c:b:c:12")
        self.assertIn("2 / 3", btn_texts(kb))
        self.assertIn("c:b:c:19", inline_data(kb))
        self.assertIn("c:b:q:12:1", inline_data(kb), "tanlangan dori amallari")
        self.assertLess(n_buttons(kb), 30)

    def test_admin_kartasi_sahifalanadi(self):
        rid = self.big_request(45)
        self.bot.ord_push_admins(rid)
        card = self.tg.last(ADMIN, "№%d" % rid)
        self.assertLess(n_buttons(card[3]), 20, "45 ta dori — tugmalar sahifalangan")
        self.assertLess(len(card[2]), 4096)
        data = inline_data(card[3])
        self.assertIn("oq:%d:i:10" % rid, data)
        self.assertNotIn("oq:%d:i:11" % rid, data)
        self.assertIn("oq:%d:p:1" % rid, data)
        mid = self.card_mid(rid)
        self.bot.handle(cb(ADMIN, "oq:%d:p:1" % rid, mid))
        text, kb = self.tg.sent[-1][2], self.tg.sent[-1][3]
        self.assertIn("Dorilar 11–20 / 45", text)
        self.assertIn("oq:%d:i:15" % rid, inline_data(kb))
        self.bot.handle(cb(ADMIN, "oq:%d:i:15" % rid, mid))
        self.assertIn("oq:%d:p:1" % rid, inline_data(self.tg.sent[-1][3]), "◀ Orqaga — shu sahifaga")
        self.bot.handle(cb(ADMIN, "oq:%d:s:15:7" % rid, mid))
        self.assertIn("oq:%d:i:15" % rid, inline_data(self.tg.sent[-1][3]), "qarordan keyin — o'sha sahifa")
        self.assertEqual(BY.get(self.bot.db, rid)[1][14]["qty_ok"], 7)

    def test_push_zaxira(self):
        rid = self.big_request(3)
        orig, calls = self.tg.send, []

        def send(cid, text, markup=None):
            calls.append(markup)
            if len(calls) == 1:
                raise RuntimeError("Bad Request: reply markup is too long")
            return orig(cid, text, markup)
        self.tg.send = send
        self.bot.ord_push_admins(rid)
        self.assertEqual(inline_data(calls[-1]), ["oq:%d" % rid])
        self.assertTrue(self.card_mid(rid), "karta xabari saqlandi")


class Lotin(Base):
    def test_tugmadagi_dori_nomi_kirillga_ogirilmaydi(self):
        rid = self.big_request(2, name="Ceftriaxone %d")
        text, kb = self.bot.ord_card(rid)
        p = B.TG._cyr({"text": text, "reply_markup": kb})
        labels = [b["text"] for row in p["reply_markup"]["inline_keyboard"] for b in row]  # Telegramga ketadigan aynan matn
        self.assertTrue(any("Ceftriaxone 1" in t for t in labels), labels)
        self.assertFalse(any("Цефтриахоне" in t for t in labels), labels)
        self.assertIn("Ҳаммасини қабул қилиш", labels[0], "tizim matni kirillga o'tdi")
        self.assertFalse(any(B.UK.DATA_OPEN in t or B.UK.DATA_CLOSE in t for t in labels), "belgilar olib tashlandi")
        self.assertIn("Ceftriaxone 1", p["text"])

    def test_savat_va_royxat_tugmalari(self):
        self.bot.cart_save(FIRM, [{"name": "Ceftriaxone", "producer": "", "qty": 1}])
        self.open(FIRM)
        _, _, _, kb = self.press(FIRM, "c:b:c:0")
        labels = btn_texts(B.TG._cyr({"text": "x", "reply_markup": kb})["reply_markup"])
        self.assertIn("▸ 1 · Ceftriaxone", labels)


class Son(Base):
    def test_vakil_nan_inf(self):
        self.open(FIRM)
        self.press(FIRM, "c:need")
        self.toggle(FIRM, 0)
        for bad in ("nan", "inf", "-inf", "1e3", "2.5", "-1", "0", "100001", "NaN"):
            self.press(FIRM, "c:b:w:0")
            self.bot.handle(msg(FIRM, bad))
            self.assertIn("Butun son yozing", self.tg.sent[-1][2], bad)
        self.assertEqual(self.bot.cart(FIRM)[0]["qty"], 46)
        self.bot.handle(msg(FIRM, "1 200"))
        self.assertEqual(self.bot.cart(FIRM)[0]["qty"], 1200)

    def test_admin_nan_inf(self):
        rid = self.send_request()
        mid = self.card_mid(rid)
        self.bot.handle(cb(ADMIN, "oq:%d:w:1" % rid, mid))
        for bad in ("nan", "inf", "2.5", "-3"):
            self.bot.handle(msg(ADMIN, bad))
            self.assertIn("Butun son yozing", self.tg.sent[-1][2], bad)
            self.assertEqual(self.bot.state[ADMIN][0], "oreq_qty")
        self.bot.handle(cb(ADMIN, "oq:%d:s:1:nan" % rid, mid))  # soxta tugma ma'lumoti — e'tiborsiz
        self.assertEqual(BY.get(self.bot.db, rid)[1][0]["status"], "kutmoqda")

    def test_panel_faqat_butun(self):
        rid = self.send_request()
        api = P.Api(self.bot, self.bot.cfg)
        for bad in (float("nan"), float("inf"), "nan", "inf", 2.5, "2.5", -1, 100001, True, "1e3"):
            self.assertFalse(api.oreq_decide({"id": rid, "pos": 1, "qty": bad})["ok"], bad)
        self.assertEqual(BY.get(self.bot.db, rid)[1][0]["status"], "kutmoqda")
        self.assertTrue(api.oreq_decide({"id": rid, "pos": 1, "qty": 7.0})["ok"])
        self.assertTrue(api.oreq_decide({"id": rid, "pos": 2, "qty": "0"})["ok"])
        items = BY.get(self.bot.db, rid)[1]
        self.assertEqual([(i["qty_ok"], i["status"]) for i in items], [(7, "ozgardi"), (0, "rad")])


class Yopilgan(Base):
    """8: o'qish va qaror orasida boshqa admin / panel yopdi — ochiq karta qayta chizilmaydi."""

    def racing(self, name):
        orig = getattr(BY, name)

        def fn(db, rid, *a, **k):
            BY.finish(db, rid, 777, "Panel", reject=True)
            return orig(db, rid, *a, **k)
        setattr(BY, name, fn)
        return lambda: setattr(BY, name, orig)

    def test_qaror(self):
        for name, data in (("decide_item", "oq:%d:s:1:5"), ("decide_item", "oq:%d:x:1"), ("accept_all", "oq:%d:all")):
            rid = self.send_request()
            mid = self.card_mid(rid)
            restore = self.racing(name)
            try:
                self.bot.handle(cb(ADMIN, data % rid, mid))
            finally:
                restore()
            self.assertTrue(self.tg.sent[-1][2].startswith("❌ №%d" % rid), (data, self.tg.sent[-1][2]))
            self.assertIsNone(self.tg.sent[-1][3])

    def test_yozilgan_son(self):
        rid = self.send_request()
        mid = self.card_mid(rid)
        self.bot.handle(cb(ADMIN, "oq:%d:w:1" % rid, mid))
        restore = self.racing("decide_item")
        try:
            self.bot.handle(msg(ADMIN, "5"))
        finally:
            restore()
        self.assertTrue(self.tg.sent[-1][2].startswith("❌ №%d" % rid), self.tg.sent[-1][2])


if __name__ == "__main__":
    unittest.main()
