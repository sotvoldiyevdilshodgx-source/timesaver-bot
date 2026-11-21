# -*- coding: utf-8 -*-
"""Kirill alifbosi: lotin → kirill (uzkir), Telegramga chiqishda o'girish, ma'lumot (nomlar) himoyasi,
kirill pastki tugmani bot tanishi."""
import json
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fixtures import B, BotCase, ADMIN, msg  # noqa: E402
import uzkir as UK  # noqa: E402


class Lat2Cyr(unittest.TestCase):
    def test_sozlar(self):
        cases = {
            "Xush kelibsiz!": "Хуш келибсиз!",
            "Ma'lumot yangilangan": "Маълумот янгиланган",
            "Ko'chirish · Yo'nalishlar": "Кўчириш · Йўналишлар",
            "SHOSHILINCH · ASSORTIMENT · ODDIY": "ШОШИЛИНЧ · АССОРТИМЕНТ · ОДДИЙ",
            "Foydalanuvchilar · Buyurtma · Qoldiq · Hisob": "Фойдаланувчилар · Буюртма · Қолдиқ · Ҳисоб",
            "e'lon · yetadi · so'm · mln": "эълон · етади · сўм · млн",
            "Bugungi holat · Menyu · Texnik rejim": "Бугунги ҳолат · Меню · Техник режим",
            "eng ko'p · tanlang · sozlang": "энг кўп · танланг · созланг",
            "yo · tog' · bog'": "ё · тоғ · боғ",
        }
        for a, b in cases.items():
            self.assertEqual(UK.lat2cyr(a), b)

    def test_himoya(self):
        self.assertEqual(UK.lat2cyr("Excel'da yuklab olish"), "Excel'да юклаб олиш")
        self.assertEqual(UK.lat2cyr("TimeSaver · /cancel · @BotFather"), "TimeSaver · /cancel · @BotFather")
        self.assertEqual(UK.lat2cyr('<b>Salom</b> <a href="https://t.me/x">havola</a>'),
                         '<b>Салом</b> <a href="https://t.me/x">ҳавола</a>')
        self.assertEqual(UK.lat2cyr("Dori: " + B.esc("Nurofen <forte>")), "Дори: Nurofen &lt;forte&gt;",
                         "ma'lumot o'zgarmaydi, belgilar olinadi")

    def test_fayl_yoli_ozgarmaydi(self):
        self.assertEqual(UK.lat2cyr(r"Sozlama fayli topilmadi: C:\ProgramData\BazaReport\db.ini"),
                         r"Созлама файли топилмади: C:\ProgramData\BazaReport\db.ini")
        self.assertEqual(UK.lat2cyr(r"Excel: C:\Users\user\Telegram Desktop\TimeSaver\baza_yangi.xlsm topildi"),
                         r"Excel: C:\Users\user\Telegram Desktop\TimeSaver\baza_yangi.xlsm топилди")
        self.assertEqual(UK.lat2cyr("config.ini va refresh.vbs yo'q"), "config.ini ва refresh.vbs йўқ")
        self.assertEqual(UK.lat2cyr(r"Tarmoq: \\server\baza papka"), r"Тармоқ: \\server\baza папка")

    def test_bosh_va_bosh_bolmagan(self):
        self.assertEqual(UK.lat2cyr(""), "")
        self.assertIsNone(UK.lat2cyr(None))


class TelegramChegarasi(unittest.TestCase):
    def tg(self):
        tg = B.TG("1:x")
        sent = []

        def fake_open(req, timeout):
            sent.append(json.loads(req.data.decode("utf-8")))
            return {"message_id": 1}

        tg._open = fake_open
        return tg, sent

    def test_matn_va_tugmalar_kirillda(self):
        tg, sent = self.tg()
        tg.send(1, "🔴 <b>SHOSHILINCH</b> · " + B.esc("Nurofen"), B.inline_kb([[("Orqaga", "a:home")]]))
        p = sent[0]
        self.assertEqual(p["text"], "🔴 <b>ШОШИЛИНЧ</b> · Nurofen")
        self.assertEqual(p["reply_markup"]["inline_keyboard"][0][0], {"text": "Орқага", "callback_data": "a:home"})

    def test_pastki_tugma_eslab_qolinadi(self):
        tg, sent = self.tg()
        tg.send(1, "menyu", B.reply_kb([["⏰ Srok · 3"]]))
        self.assertEqual(sent[0]["reply_markup"]["keyboard"][0][0]["text"], "⏰ Срок · 3")
        self.assertEqual(UK.KB_BACK["⏰ Срок · 3"], "⏰ Srok · 3")


class KirillTugmaBosildi(BotCase):
    def test_admin_kirill_tugma(self):
        UK.KB_BACK["👥 Фойдаланувчилар"] = "👥 Foydalanuvchilar"
        self.bot.handle(msg(ADMIN, "👥 Фойдаланувчилар"))
        self.assertTrue(self.tg.texts(ADMIN), "bot kirill tugmani tanidi va javob berdi")


if __name__ == "__main__":
    unittest.main()
