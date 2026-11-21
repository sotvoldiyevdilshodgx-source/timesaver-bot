# -*- coding: utf-8 -*-
"""2026-11-17 tekshiruvidagi tuzatishlar: eski tugmalar, bloklangan foydalanuvchi, yo'ldagi tovar, bo'sh varaq,
qayta shakllantirish, fayl yo'li, SQL ulanish satri va db.ini kodlashi."""
import datetime
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fixtures import B, BotCase, CH, UCH, EK, ADMIN, FIRM, F_CH, F_EK, F_UCH, TODAY, add_client, cb, msg, inline_data
from openpyxl import load_workbook
import fom_sql

WORKER = {CH: F_CH, EK: F_EK, UCH: F_UCH}


class EskiTugma(BotCase):
    def accounts(self):
        return [r[0] for r in self.bot.db.execute("SELECT account FROM client_accounts WHERE chat_id=?", (FIRM,))]

    def test_eski_qidiruv_tugmasi_boshqa_hisobni_biriktirmaydi(self):
        self.bot.search[ADMIN] = ["ALFA PHARM", "ALFA MED"]
        old = "at:%d:0:%s" % (FIRM, self.bot._ck("ALFA PHARM"))
        self.bot.search[ADMIN] = ["BETA MED"]  # admin keyin boshqa so'z bilan qidirdi
        self.bot.handle(cb(ADMIN, old))
        self.assertEqual(self.accounts(), [])
        self.assertIn("eskirgan", self.tg.texts(ADMIN)[-1])
        self.bot.handle(cb(ADMIN, "at:%d:0:%s" % (FIRM, self.bot._ck("BETA MED"))))
        self.assertEqual(self.accounts(), ["BETA MED"])

    def test_eski_matn_korinishi_yangi_matnni_yubormaydi(self):
        self.bot.pending_text[ADMIN] = "ikkinchi matn"
        self.bot.handle(cb(ADMIN, "a:tx:all:%s" % self.bot._ck("birinchi matn")))
        self.assertIn("eskirgan", self.tg.texts(ADMIN)[-1])
        self.assertEqual(self.bot.pending_text.get(ADMIN), "ikkinchi matn", "matn saqlanib qoladi")


class Bloklangan(BotCase):
    def test_firma_belgisi_tahriri_blokdan_chiqarmaydi(self):
        add_client(self.bot, 7777, "firm", status="blocked")
        self.bot.state[ADMIN] = ("firms", 7777)
        self.bot.handle(msg(ADMIN, "АЛЬФА 901112233"))
        self.assertEqual(self.bot.client(7777)["status"], "blocked")
        self.assertFalse(self.tg.texts(7777), "bloklangan foydalanuvchiga «qabul qilindingiz» bormaydi")

    def test_kutayotgan_foydalanuvchi_faollashadi(self):
        add_client(self.bot, 7778, "firm", status="pending")
        self.bot.state[ADMIN] = ("firms", 7778)
        self.bot.handle(msg(ADMIN, "АЛЬФА 901112233"))
        self.assertEqual(self.bot.client(7778)["status"], "active")


class YoldagiTovar(BotCase):
    def setUp(self):
        super().setUp()
        self.bot.koch_build()

    def test_uch_kun_oldin_berilgani_qabul_qilinadi(self):
        t = self.bot.db.execute("SELECT * FROM koch_tasks WHERE src=? AND status='yangi'", (CH,)).fetchone()
        self.assertTrue(self.bot.koch_give(F_CH, self.bot.client(F_CH), t["id"]))
        old = (TODAY - datetime.timedelta(days=3)).isoformat()
        self.bot.db.execute("UPDATE koch_tasks SET sana=? WHERE id=?", (old, t["id"]))
        self.bot.db.commit()
        dst = WORKER[t["dst"]]
        c = self.bot.client(dst)
        self.assertIn(t["id"], [x["id"] for _, ts in self.bot.koch_in_routes(t["dst"]) for x in ts], "Keladigan ro'yxatida")
        self.assertGreaterEqual(self.bot.filial_counts(t["dst"])["kin"], 1)
        self.assertTrue(self.bot.koch_receive(dst, c, t["id"]))
        self.assertEqual(self.bot.db.execute("SELECT status FROM koch_tasks WHERE id=?", (t["id"],)).fetchone()[0], "olindi")


class QaytaShakllantirish(BotCase):
    def test_kunning_birinchisi_otgan_kunni_yopadi(self):
        self.bot.koch_build()
        y = (TODAY - datetime.timedelta(days=1)).isoformat()
        t = self.bot.db.execute("SELECT * FROM koch_tasks WHERE status='yangi'").fetchone()
        self.bot.db.execute("UPDATE koch_tasks SET sana=? WHERE id=?", (y, t["id"]))
        self.bot.db.commit()
        self.bot.set_setting("koch_snap_date", y)  # bugungi ertalabki shakllantirish o'xshamagan
        self.bot.koch_regen_manual(by=(None, "admin"))
        self.assertEqual(self.bot.db.execute("SELECT status FROM koch_tasks WHERE id=?", (t["id"],)).fetchone()[0], "yopildi")
        self.assertEqual(self.bot.setting("koch_snap_date"), TODAY.isoformat())


class BoshVaraq(BotCase):
    def test_bosh_kochirish_varagi_butun_excelni_buzmaydi(self):
        path = self.bot.data.path
        wb = load_workbook(path)
        ws = [w for w in wb.worksheets if "chirish" in w.title][0]
        title = ws.title
        wb.remove(ws)
        wb.create_sheet(title)
        wb.save(path)
        self.assertTrue(self.bot.data.reload(force=True), self.bot.data.error)
        self.assertEqual(self.bot.data.koch, [])
        self.assertTrue(self.bot.data.sale)


class Yordamchi(unittest.TestCase):
    def test_qavsli_fayl_nomi(self):
        d = tempfile.mkdtemp()
        p = os.path.join(d, "baza[1].xlsm")
        open(p, "wb").close()
        self.assertEqual(B.newest_file(p), p)

    def test_odbc_quote(self):
        self.assertEqual(fom_sql.odbc_quote("ab;c}d"), "{ab;c}}d}")

    def test_db_ini_ansi_va_utf8(self):
        d = tempfile.mkdtemp()
        p = os.path.join(d, "db.ini")
        for enc in ("cp1251", "utf-8"):
            with open(p, "w", encoding=enc) as f:
                f.write("Server=S\nUid=sa\nPwd=Пароль1\n")
            self.assertEqual(fom_sql.read_flat_ini(p)["pwd"], "Пароль1", enc)


if __name__ == "__main__":
    unittest.main()
