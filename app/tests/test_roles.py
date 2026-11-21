# -*- coding: utf-8 -*-
"""B3: rollar va filial xodimini aniqlash (invite, Telegram kartochka, server paneli), staff panelida ko'chirish yo'q."""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fixtures import (B, BotCase, CH, UCH, EK, ADMIN, STAFF, FIRM, F_CH, msg, cb, kb_texts, inline_data)
import timesaver_panel as P


class Rollar(BotCase):
    def test_staff_panelida_kochirish_yoq(self):
        kb = kb_texts(self.bot.client_kb(self.bot.client(STAFF)))
        self.assertFalse([t for t in kb if "Ko'chirish" in t or "SHOSHILINCH" in t], kb)
        self.assertTrue(kb[-1].startswith("🏢 Firma"))
        for data in ("x:koch", "kx:shoshilinch", "ki:1"):
            self.bot.handle(cb(STAFF, data))
        self.assertFalse(self.tg.to(STAFF, "doc"))

    def test_can_use_va_tags(self):
        c = self.bot.client(F_CH)
        self.assertTrue(self.bot.is_filial(c) and self.bot.can_use(c))
        self.assertEqual(self.bot.tags_for(c), set(), "filial xodimi firma ma'lumotini ko'rmaydi")

    def test_invite_filial(self):
        r = self.bot.add_user("+998901234567 filial ЧИНОЗ")
        inv = self.bot.db.execute("SELECT * FROM invites WHERE phone9='901234567'").fetchone()
        self.assertEqual((inv["role"], inv["filial"]), ("filial", CH), r)
        self.assertIn("aniq emas", self.bot.add_user("+998901112299 filial ДОРИХОНА"))
        self.assertIn("topilmadi", self.bot.add_user("+998901112299 filial ТОШКЕНТ"))
        self.bot.handle(msg(7001, None, contact={"user_id": 7001, "phone_number": "998901234567"}))
        c = self.bot.client(7001)
        self.assertEqual((c["status"], c["role"], c["filial"]), ("active", "filial", CH))
        self.assertTrue(any("🏬" in t for t in self.tg.texts(7001)), "kontakt kelganda filial paneli")

    def test_telegram_kartochka(self):
        self.bot.handle(cb(ADMIN, "uc:%d" % FIRM))
        self.assertIn("fl:%d" % FIRM, inline_data(self.tg.to(ADMIN, "send")[-1][3]))
        self.bot.handle(cb(ADMIN, "fl:%d" % FIRM))
        i = [B.norm(x) for x in self.bot.filsearch[ADMIN]].index(UCH)
        self.bot.handle(cb(ADMIN, "fs:%d:%d:%s" % (FIRM, i, self.bot._ck(self.bot.filsearch[ADMIN][i]))))
        self.assertEqual((self.bot.client(FIRM)["role"], self.bot.client(FIRM)["filial"]), ("filial", UCH))
        self.bot.handle(cb(ADMIN, "uc:%d" % FIRM))
        card = self.tg.texts(ADMIN)[-1]
        data = inline_data(self.tg.to(ADMIN, "send")[-1][3])
        self.assertIn("filial xodimi · " + UCH, card)
        self.assertTrue({"rf:%d" % FIRM, "rs:%d" % FIRM} <= set(data))
        self.bot.handle(cb(ADMIN, "rs:%d" % FIRM))
        self.assertEqual((self.bot.client(FIRM)["role"], self.bot.client(FIRM)["filial"]), ("staff", None), "filial tozalandi")

    def test_royxatlarda(self):
        c = self.bot.client(F_CH)
        self.assertEqual(self.bot.role_label(c), "filial · " + CH)
        self.assertIn(CH, self.bot._user_btn(c)[0])
        self.assertTrue(any(r["chat_id"] == F_CH for r in self.bot.find_users("ЧИНОЗ")))

    def test_server_paneli(self):
        api = P.Api(self.bot, self.bot.cfg)
        self.assertEqual([B.norm(x) for x in api.filials()], [UCH, CH, EK])
        self.assertFalse(api.user_save({"id": STAFF, "role": "filial", "filial": "ТОШКЕНТ"})["ok"])
        self.assertTrue(api.user_save({"id": STAFF, "role": "filial", "filial": "дорихона  ургут"})["ok"])
        self.assertEqual(self.bot.client(STAFF)["filial"], UCH)
        self.assertTrue([u for u in api.users("filial") if u["id"] == STAFF])
        self.assertTrue(api.user_save({"id": STAFF, "role": "staff"})["ok"])
        self.assertIsNone(self.bot.client(STAFF)["filial"])
        r = api.user_add({"phone": "+998 93 555 44 33", "role": "filial", "filial": EK})
        self.assertTrue(r["ok"])
        self.assertEqual(self.bot.db.execute("SELECT filial FROM invites WHERE phone9='935554433'").fetchone()[0], EK)

    def test_filial_boshqa_bolimlarni_ololmaydi(self):
        for data in ("c:srok", "x:srok", "x:acc", "x:koch", "a:fm", "fp:0"):
            self.bot.handle(cb(F_CH, data))
        self.assertFalse(self.tg.to(F_CH, "doc"))
        self.assertFalse(self.tg.texts(F_CH))


if __name__ == "__main__":
    unittest.main()
