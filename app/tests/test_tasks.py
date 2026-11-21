# -*- coding: utf-8 -*-
"""B5: kunlik topshiriq (snapshot), migratsiya, toifa muddatlari, kartochka, Berdim / Oldim / Rad / Kam keldi."""
import collections
import datetime
import os
import shutil
import sqlite3
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fixtures import B, BotCase, CH, UCH, EK, ADMIN, STAFF, F_CH, F_EK, F_UCH, TODAY, msg, cb, inline_data, make_bot, close_bot


class Snapshot(BotCase):
    def test_yaratish_va_chiqarishlar(self):
        cnt = self.bot.koch_build()
        self.assertEqual(cnt, {"new": 19, "updated": 0, "cancelled": 0, "kept": 0})
        rows = self.bot.db.execute("SELECT * FROM koch_tasks").fetchall()
        self.assertEqual(collections.Counter(r["cat"] for r in rows), {"shoshilinch": 9, "assortiment": 4, "oddiy": 6})
        got = {r["name"] for r in rows}
        for n in ("КАСР", "НОМАЪЛУМ", "ШОШ ДОРИ 03", "АССОРТ 00", "АССОРТ 02", "АССОРТ 05", "ОДДИЙ 02"):
            self.assertNotIn(n, got, "E < 1, toifasiz, muddati o'tgan, ASSORTIMENT < 60 kun — topshiriqqa kirmaydi")
        r = [x for x in rows if x["name"] == "ОДДИЙ 05"][0]
        self.assertEqual((r["src"], r["dst"], r["days_idle"], r["status"], r["turi"]), (EK, CH, 200, "yangi", "ODDIY"))
        self.assertEqual((r["dst_sales"], r["src_stock"]), (3, 5))

    def test_takror_ozgartirmaydi(self):
        self.bot.koch_build()
        self.assertEqual(self.bot.koch_build(), {"new": 0, "updated": 0, "cancelled": 0, "kept": 19})

    def test_qayta_shakllantirish(self):
        self.bot.koch_build()
        tid = lambda n: self.bot.db.execute("SELECT id FROM koch_tasks WHERE name=?", (n,)).fetchone()[0]
        self.assertTrue(self.bot.koch_give(F_CH, self.bot.client(F_CH), tid("ШОШ ДОРИ 00")))
        self.bot.data.koch = [r for r in self.bot.data.koch if r["name"] not in ("ШОШ ДОРИ 00", "ШОШ ДОРИ 01")]
        cnt = self.bot.koch_build(regen=True, by=(None, "admin"))
        self.assertEqual((cnt["cancelled"], cnt["updated"], cnt["new"], cnt["kept"]), (1, 0, 0, 17), "o'zgarmaganlar — saqlandi")
        self.assertFalse(self.bot.db.execute("SELECT 1 FROM koch_events WHERE action='regen:yangilandi'").fetchone(),
                         "o'zgarmagan qator tarixga yozilmaydi")
        st = dict(self.bot.db.execute("SELECT name, status FROM koch_tasks WHERE name IN ('ШОШ ДОРИ 00','ШОШ ДОРИ 01')").fetchall())
        self.assertEqual(st, {"ШОШ ДОРИ 00": "berildi", "ШОШ ДОРИ 01": "bekor"}, "berildi saqlanadi, ro'yxatdan chiqqan yangi — bekor")

    def test_eski_makros_bilan_yaratilmaydi(self):
        bot, tg, work = make_bot("old")
        self.addCleanup(shutil.rmtree, work, True)
        self.addCleanup(close_bot, bot)
        self.clock.at(23, 59)  # Excel vaqti (Srok!B2) — hozir; topshiriq vaqti o'tgan
        self.assertFalse(bot.koch_snapshot_due(), "Turi ustuni yo'q eski varaq — topshiriq yaratilmaydi")


class AvtoSinxron(BotCase):
    """Kun davomidagi har «Обновить» (bot yoki Excelda qo'lda) dan keyin bugungi topshiriq Excelga moslashadi."""

    def setUp(self):
        super().setUp()
        t0 = B.datetime.datetime.combine(TODAY, B.datetime.time(8, 30))  # ertalabki avto-yangilash
        self.bot.data.when = lambda: t0
        self.bot.koch_build()
        self.orig = list(self.bot.data.koch)

    def refresh(self, h, m, rows):
        """Excel yangilandi: «Ko'chirish» varag'i yangi, ichidagi vaqt — h:m."""
        self.clock.at(h, m)
        self.bot.data.koch = rows
        t = B.datetime.datetime.combine(TODAY, B.datetime.time(h, m))
        self.bot.data.when = lambda: t
        return self.bot.koch_try_snapshot()

    def row(self, name):
        return self.bot.db.execute("SELECT * FROM koch_tasks WHERE name=?", (name,)).fetchone()

    def excel(self, drop=(), qty=None, add=None):
        out = []
        for r in self.orig:
            if r["name"] in drop:
                continue
            r = dict(r)
            if qty and r["name"] in qty:
                r["qty"] = qty[r["name"]]
            out.append(r)
        if add:
            out.append(dict(out[0], name=add))
        return out

    def test_yangilanmagan_bolsa_tegmaydi(self):
        self.assertFalse(self.bot.koch_sync_due())
        self.assertIsNone(self.bot.koch_try_snapshot())

    def test_obnovitdan_keyin_moslashadi(self):
        give = self.row("ШОШ ДОРИ 01")["id"]
        self.bot.koch_give(F_CH, self.bot.client(F_CH), give)
        cnt = self.refresh(11, 0, self.excel(drop=("ШОШ ДОРИ 00", "ШОШ ДОРИ 01"), qty={"ШОШ ДОРИ 02": 3}, add="ЯНГИ ДОРИ"))
        self.assertEqual((cnt["new"], cnt["updated"], cnt["cancelled"]), (1, 1, 1))
        self.assertEqual(self.row("ШОШ ДОРИ 00")["status"], "bekor")
        self.assertEqual(self.row("ШОШ ДОРИ 01")["status"], "berildi", "berilgan saqlanadi")
        self.assertEqual(self.row("ШОШ ДОРИ 02")["qty_plan"], 3)
        self.assertEqual(self.row("ЯНГИ ДОРИ")["status"], "yangi")
        ev = [e[0] for e in self.bot.db.execute("SELECT action FROM koch_events WHERE action LIKE 'regen:%'")]
        self.assertEqual(sorted(ev), ["regen:bekor", "regen:yangilandi"], "faqat haqiqiy o'zgarishlar tarixda")
        self.assertIsNone(self.bot.koch_try_snapshot(), "shu Excel bilan takror sinxron yo'q")
        self.assertIn("Jami: <b>6</b> ta dori", self.bot.koch_list_text(self.bot.client(F_CH), "shoshilinch")[0],
                      "ro'yxat (faqat «yangi»): 7 − 2 + 1 yangi dori")

    def test_muddatdan_keyin_qoshilgan_kechikkan_emas(self):
        self.refresh(11, 0, self.excel(drop=("ШОШ ДОРИ 00",), add="ЯНГИ ДОРИ"))
        self.clock.at(12, 0)
        self.assertFalse(self.bot.task_info(self.row("ЯНГИ ДОРИ"))["late"], "10:00 dan keyin qo'shilgan")
        self.assertTrue(self.bot.task_info(self.row("ШОШ ДОРИ 06"))["late"], "ertalabki — kechikkan")
        self.refresh(13, 0, self.excel())  # ШОШ ДОРИ 00 qaytdi
        r = self.row("ШОШ ДОРИ 00")
        self.assertEqual(r["status"], "yangi")
        self.assertFalse(self.bot.task_info(r)["late"], "qaytgan topshiriq ham kechikkan emas")
        self.bot.koch_give(F_CH, self.bot.client(F_CH), r["id"])
        self.bot.koch_receive(F_EK, self.bot.client(F_EK), r["id"])
        self.assertFalse(self.bot.task_info(self.row("ШОШ ДОРИ 00"))["late_done"])

    def test_kechagi_excel_sinxron_qilmaydi(self):
        self.bot.data.koch = self.excel(drop=("ШОШ ДОРИ 00",))
        y = B.datetime.datetime.combine(TODAY - B.datetime.timedelta(days=1), B.datetime.time(18, 0))
        self.bot.data.when = lambda: y
        self.assertFalse(self.bot.koch_sync_due())

    def test_nazorat_ochiq_yoki_yangilanmoqda(self):
        self.bot.data.koch = self.excel(drop=("ШОШ ДОРИ 00",))
        t = B.datetime.datetime.combine(TODAY, B.datetime.time(11, 0))
        self.bot.data.when = lambda: t
        self.bot.refreshing = True
        self.assertFalse(self.bot.koch_sync_due(), "yangilash tugashini kutadi")
        self.bot.refreshing = False
        self.bot.set_setting("koch_on", "0")
        self.assertFalse(self.bot.koch_sync_due())
        self.bot.set_setting("koch_on", "1")
        self.assertTrue(self.bot.koch_sync_due())


class BoribKelish(BotCase):
    """Oxirgi 14 kunda B → A olingan tovar bugun A → B taklif qilinmaydi (topshiriqda ham, ro'yxatda ham)."""

    def setUp(self):
        super().setUp()
        self.bot.koch_build()
        tid = self.bot.db.execute("SELECT id FROM koch_tasks WHERE name='АВИТЕН'").fetchone()[0]   # ЭРКИН → ЧИНОЗ
        self.bot.koch_give(F_EK, self.bot.client(F_EK), tid)
        self.bot.koch_receive(F_CH, self.bot.client(F_CH), tid)
        av = [r for r in self.bot.data.koch if r["name"] == "АВИТЕН"][0]
        self.back = dict(av, src=CH, dst=EK)            # bugun Excel teskari yo'nalishni taklif qildi
        self.bot.data.koch = self.bot.data.koch + [self.back]

    def test_sabab(self):
        self.assertEqual(self.bot.koch_exclude_reason(self.back), "14 kun ichida shu filialdan kelgan")
        other = dict(self.back, dst=UCH)
        self.assertIsNone(self.bot.koch_exclude_reason(other), "boshqa filialga — mumkin")

    def test_royxat_va_excel(self):
        rows, exc = self.bot.koch_view(CH, "shoshilinch")
        self.assertNotIn(("АВИТЕН", EK), [(r["name"], B.norm(r["dst"])) for r in rows])
        self.assertIn(("АВИТЕН", "14 kun ichida shu filialdan kelgan"), [(r["name"], w) for r, w in exc])
        self.assertIn("14 kun ichida shu filialdan kelgan 1 ta dori", self.bot.koch_list_text(self.bot.client(F_CH), "shoshilinch")[0])

    def test_topshiriqqa_kirmaydi(self):
        self.bot.koch_build(regen=True)
        self.assertFalse(self.bot.db.execute("SELECT 1 FROM koch_tasks WHERE name='АВИТЕН' AND src=? AND dst=?", (CH, EK)).fetchone())

    def test_muddat_va_ochirish(self):
        self.clock.at(9, 0, TODAY + B.datetime.timedelta(days=15))
        self.assertIsNone(self.bot.koch_exclude_reason(self.back), "15 kundan keyin — yana mumkin")
        self.clock.at(9, 0, TODAY + B.datetime.timedelta(days=3))
        self.assertTrue(self.bot.koch_exclude_reason(self.back))
        self.bot.set_setting("koch_no_return_days", "0")
        self.assertIsNone(self.bot.koch_exclude_reason(self.back), "0 — filtr o'chiq")


class Migratsiya(unittest.TestCase):
    def test_eski_baza(self):
        work = tempfile.mkdtemp(prefix="abu_mig_")
        self.addCleanup(shutil.rmtree, work, True)
        db = os.path.join(work, "bot.db")
        con = sqlite3.connect(db)
        con.executescript("""
            CREATE TABLE settings(key TEXT PRIMARY KEY, value TEXT);
            INSERT INTO settings VALUES('koch_deadline_zudlik', '09:15'), ('koch_rem_report_on', '0');
            CREATE TABLE koch_tasks(
                id INTEGER PRIMARY KEY AUTOINCREMENT, sana TEXT, src TEXT, dst TEXT, name TEXT, producer TEXT, firm TEXT,
                turi TEXT, sabab TEXT, qty_plan REAL, prih REAL, summa REAL, srok TEXT,
                status TEXT NOT NULL DEFAULT 'yangi', qty_given REAL, qty_received REAL, given_by INTEGER, given_at TEXT,
                received_by INTEGER, received_at TEXT, note TEXT, created_at TEXT, kkey TEXT, UNIQUE(sana, kkey));
            INSERT INTO koch_tasks(sana, src, dst, name, turi, sabab, qty_plan, status, kkey) VALUES
                ('2026-10-29', 'A', 'B', 'D1', 'zudlik', 'Tugagan, sotuv bor', 1, 'olindi', 'k1'),
                ('2026-10-29', 'A', 'B', 'D2', 'zudlik', 'Filialda yo''q, boshqa filialda sotilyapti', 1, 'yangi', 'k2'),
                ('2026-10-29', 'A', 'B', 'D3', 'oddiy', '45 kun turibdi, oluvchida yetmaydi', 2, 'berildi', 'k3');
        """)
        con.commit()
        con.close()
        bot, tg, w2 = make_bot("new")
        self.addCleanup(shutil.rmtree, w2, True)
        close_bot(bot)
        bot = B.Bot(bot.cfg, tg=tg, db_path=db)
        self.addCleanup(close_bot, bot)
        rows = {r["name"]: r for r in bot.db.execute("SELECT * FROM koch_tasks")}
        self.assertEqual({n: (r["cat"], r["days_idle"], r["status"]) for n, r in rows.items()},
                         {"D1": ("shoshilinch", 0, "olindi"), "D2": ("assortiment", 0, "yangi"), "D3": ("oddiy", 45, "berildi")},
                         "mavjud qatorlar saqlanadi, toifa turi/sabab dan")
        self.assertEqual((bot.kset("koch_deadline_shoshilinch"), bot.kset("koch_deadline_assortiment")), ("09:15", "09:15"))
        self.assertEqual(bot.setting("koch_deadline_zudlik"), "09:15", "eski kalit o'chirilmaydi")
        self.assertEqual(bot.setting("alert_daily_report"), "0")
        B.Bot(bot.cfg, tg=tg, db_path=db).db.close()  # takror ishga tushirish xatosiz


class Muddat(BotCase):
    def setUp(self):
        super().setUp()
        self.bot.koch_build()

    def t(self, name):
        return self.bot.db.execute("SELECT * FROM koch_tasks WHERE name=?", (name,)).fetchone()

    def test_toifa_muddatlari(self):
        sh, a, o = self.t("ШОШ ДОРИ 00"), self.t("АССОРТ 03"), self.t("ОДДИЙ ЧАР")
        self.assertEqual([self.bot.koch_deadline(x["cat"]).strftime("%H:%M") for x in (sh, a, o)], ["10:00", "10:00", "18:00"])
        self.clock.at(10, 30)
        i = self.bot.task_info(sh)
        self.assertEqual((i["late"], i["guilty"], i["guilty_filial"], i["late_min"]), (True, "src", CH, 30))
        self.assertTrue(self.bot.task_info(a)["late"])
        self.assertFalse(self.bot.task_info(o)["late"])
        self.bot.set_setting("koch_deadline_assortiment", "11:00")
        self.assertFalse(self.bot.task_info(a)["late"], "ASSORTIMENT muddati alohida sozlanadi")
        self.bot.set_setting("koch_deadline_oddiy", "")
        self.clock.at(23, 0)
        self.assertFalse(self.bot.task_info(o)["late"], "ODDIY muddatsiz")

    def test_aybdor_oluvchi_va_kech_bajarilgan(self):
        sh = self.t("АВИТЕН")
        c_ek, c_ch = self.bot.client(F_EK), self.bot.client(F_CH)
        self.clock.at(9, 0)
        self.bot.koch_give(F_EK, c_ek, sh["id"])
        self.clock.at(10, 20)
        i = self.bot.task_info(self.t("АВИТЕН"))
        self.assertEqual((i["late"], i["guilty_filial"]), (True, CH), "berildi → aybdor oluvchi")
        self.assertTrue(self.bot.koch_receive(F_CH, c_ch, sh["id"]), "kechikkan ham bajariladi")
        i = self.bot.task_info(self.t("АВИТЕН"))
        self.assertEqual((i["late"], i["late_done"], i["late_min"], i["guilty_filial"]), (False, True, 20, CH))


class Kartochka(BotCase):
    def setUp(self):
        super().setUp()
        self.bot.koch_build()

    def tid(self, name):
        return self.bot.db.execute("SELECT id FROM koch_tasks WHERE name=?", (name,)).fetchone()[0]

    def row(self, tid):
        return self.bot.db.execute("SELECT * FROM koch_tasks WHERE id=?", (tid,)).fetchone()

    def test_kartochka_va_ruxsat(self):
        tid = self.tid("АВИТЕН")
        self.bot.handle(cb(F_EK, "ki:%d:0" % tid))
        text, kb = self.tg.sent[-1][2], self.tg.sent[-1][3]
        self.assertTrue(text.startswith("🏠 › 🔴 SHOSHILINCH › #%d\n\n" % tid), "yo'l qatori")
        for s in ("🔴 <b>SHOSHILINCH</b> · #%d" % tid, "💊 <b>АВИТЕН</b>", "🏬 %s → <b>%s</b>" % (EK, CH),
                  "Soni: <b>3</b> qadoq", "Sabab: Tugagan, sotuv bor", "Holat: <b>⏳ yangi</b>", "⏰ Muddat 10:00 gacha"):
            self.assertIn(s, text)
        self.assertEqual(inline_data(kb), ["kb:%d" % tid, "kq:%d" % tid, "kr:%d" % tid, "kp:shoshilinch:0", "kh"],
                         "oxirgi qator: ◀ Orqaga (ro'yxatga) · 🏠 Menyu")
        self.bot.handle(cb(F_CH, "ki:%d:0" % tid))
        self.assertEqual(inline_data(self.tg.sent[-1][3]), ["kl:0", "kh"], "oluvchi yangi holatda faqat ko'radi")
        self.bot.handle(cb(F_UCH, "ki:%d:0" % tid))
        self.assertIn("tegishli emas", self.tg.texts(F_UCH)[-1])
        for cid in (STAFF, ADMIN):
            n = len(self.tg.sent)
            self.bot.handle(cb(cid, "ki:%d:0" % tid))
            self.bot.handle(cb(cid, "kb:%d" % tid))
            self.assertEqual(len(self.tg.sent), n, "boshqa rol — jim rad")
        self.assertEqual(self.row(tid)["status"], "yangi")

    def test_toliq_oqim_berdim_oldim(self):
        tid = self.tid("АВИТЕН")
        self.bot.handle(cb(F_CH, "kb:%d" % tid))
        self.assertEqual(self.row(tid)["status"], "yangi", "oluvchi «Berdim» bosa olmaydi")
        self.bot.handle(cb(F_EK, "kb:%d" % tid))
        r = self.row(tid)
        self.assertEqual((r["status"], r["qty_given"], r["given_by"]), ("berildi", 3, F_EK))
        note = self.tg.last(F_CH, "dan keladi")
        self.assertIn("📦 <b>%s dan keladi</b>: АВИТЕН · <b>3</b> qadoq" % EK, note[2])
        self.assertEqual(inline_data(note[3]), ["ka:%d" % tid, "kk:%d" % tid])
        self.assertIn("Holat: <b>📦 berildi</b>", self.tg.texts(F_EK)[-1] if "Holat" in self.tg.texts(F_EK)[-1] else
                      [x for x in self.tg.texts(F_EK) if "Holat" in x][-1])
        self.bot.handle(cb(F_EK, "kb:%d" % tid))  # ikkinchi marta
        self.assertEqual(self.bot.db.execute("SELECT COUNT(*) FROM koch_events WHERE task_id=? AND action='berildi'",
                                             (tid,)).fetchone()[0], 1, "idempotent")
        self.bot.handle(cb(F_EK, "ka:%d" % tid))
        self.assertEqual(self.row(tid)["status"], "berildi", "beruvchi «Oldim» bosa olmaydi")
        self.bot.handle(cb(F_CH, "ka:%d" % tid))
        r = self.row(tid)
        self.assertEqual((r["status"], r["qty_received"], r["received_by"]), ("olindi", 3, F_CH))
        self.assertIn("✅ <b>%s qabul qildi</b>: АВИТЕН · 3 qadoq" % CH, self.tg.last(F_EK, "qabul qildi")[2])
        self.bot.handle(cb(F_CH, "ka:%d" % tid))
        self.assertEqual(self.bot.db.execute("SELECT COUNT(*) FROM koch_events WHERE task_id=?", (tid,)).fetchone()[0], 2)

    def test_ikki_xodim_bir_vaqtda(self):
        from fixtures import add_client
        add_client(self.bot, 6099, "filial", filial=EK)
        tid = self.tid("АВИТЕН")
        a = self.bot.koch_give(F_EK, self.bot.client(F_EK), tid)
        b = self.bot.koch_give(6099, self.bot.client(6099), tid)
        self.assertEqual((a, b), (True, None))

    def test_boshqa_son(self):
        tid = self.tid("ШОШ ДОРИ 00")  # reja 5
        self.bot.handle(cb(F_CH, "kq:%d" % tid))
        self.assertIn("Nechta berdingiz", self.tg.texts(F_CH)[-1])
        self.bot.handle(msg(F_CH, "9"))
        self.assertIn("1 dan 5 gacha", self.tg.texts(F_CH)[-1])
        self.bot.handle(msg(F_CH, "2,5"))
        self.assertIn("1 dan 5 gacha", self.tg.texts(F_CH)[-1], "butun son")
        self.bot.handle(msg(F_CH, "2"))
        r = self.row(tid)
        self.assertEqual((r["status"], r["qty_given"]), ("berildi", 2))
        self.assertIn("(reja 5)", self.tg.last(F_EK, "dan keladi")[2])
        self.bot.handle(cb(F_EK, "kk:%d" % tid))
        self.bot.handle(msg(F_EK, "1"))
        r = self.row(tid)
        self.assertEqual((r["status"], r["qty_received"]), ("olindi", 1))
        self.assertIn("kam keldi", self.tg.last(F_CH, "qabul qildi")[2])
        self.assertTrue(self.bot.task_info(r)["farq"])

    def test_toifa_tugmasi_holatni_bekor_qiladi(self):
        tid = self.tid("ШОШ ДОРИ 00")
        self.bot.handle(cb(F_CH, "kq:%d" % tid))
        self.bot.handle(msg(F_CH, "🟢 ODDIY · 1"))
        self.assertTrue(any("🟢 <b>ODDIY</b>" in x for x in self.tg.texts(F_CH)[-2:]))
        self.bot.handle(msg(F_CH, "2"))
        self.assertEqual(self.row(tid)["status"], "yangi")

    def test_bera_olmayman(self):
        tid = self.tid("ШОШ ДОРИ 01")
        self.bot.handle(cb(F_CH, "kr:%d" % tid))
        self.assertEqual(inline_data(self.tg.sent[-1][3]), ["krr:%d:0" % tid, "krr:%d:1" % tid, "krr:%d:2" % tid,
                                                            "ki:%d:0" % tid, "kh"], "sabablar + ◀ Orqaga (kartochkaga) · 🏠 Menyu")
        self.bot.handle(cb(F_CH, "krr:%d:0" % tid))
        r = self.row(tid)
        self.assertEqual((r["status"], r["reason"]), ("rad", "Tovar topilmadi"))
        self.assertIn("bera olmaydi", self.tg.last(F_UCH, "bera olmaydi")[2])
        tid2 = self.tid("ШОШ ДОРИ 02")
        self.bot.handle(cb(F_CH, "krr:%d:2" % tid2))
        self.bot.handle(msg(F_CH, "Omborda singan"))
        self.assertEqual((self.row(tid2)["status"], self.row(tid2)["reason"]), ("rad", "Omborda singan"))
        self.clock.at(11, 0)
        self.assertFalse(self.bot.task_info(self.row(tid2))["late"], "rad — kechikkan emas")

    def test_cancel(self):
        tid = self.tid("ШОШ ДОРИ 02")
        self.bot.handle(cb(F_CH, "krr:%d:2" % tid))
        self.bot.handle(msg(F_CH, "/cancel"))
        self.bot.handle(msg(F_CH, "Omborda singan"))
        self.assertEqual(self.row(tid)["status"], "yangi", "/cancel — o'zgarmadi")

    def test_farq_filtri_va_regen_audit(self):
        tid = self.tid("ШОШ ДОРИ 00")
        self.bot.koch_give(F_CH, self.bot.client(F_CH), tid, 3)
        lst = self.bot.koch_task_list(status="farq")
        self.assertEqual([t["name"] for t in lst], ["ШОШ ДОРИ 00"])
        self.bot.data.koch = [r for r in self.bot.data.koch if r["name"] != "ШОШ ДОРИ 01"]
        self.bot.koch_build(regen=True, by=(None, "admin"))
        ev = self.bot.koch_task_detail(self.tid("ШОШ ДОРИ 01"))["events"]
        self.assertEqual(ev[-1]["action"], "regen:bekor")

    def test_sana_tekshiruvi(self):
        tid = self.tid("ШОШ ДОРИ 00")
        self.clock.at(9, 0, TODAY + datetime.timedelta(days=1))
        self.assertTrue(self.bot.koch_give(F_CH, self.bot.client(F_CH), tid), "kechagi topshiriq — bajariladi")
        tid2 = self.tid("ШОШ ДОРИ 01")
        self.clock.at(9, 0, TODAY + datetime.timedelta(days=2))
        self.assertIsNone(self.bot.koch_give(F_CH, self.bot.client(F_CH), tid2), "2 kun oldingi — yo'q")

    def test_admin_holat_ozgartirish(self):
        tid = self.tid("ОДДИЙ ЧАР")
        self.assertEqual(self.bot.koch_set_status(tid, "bekor", ""), "Izoh yozing")
        self.assertIsNone(self.bot.koch_set_status(tid, "bekor", "noto'g'ri qator"))
        self.assertEqual(self.row(tid)["status"], "bekor")
        d = self.bot.koch_task_detail(tid)
        self.assertEqual((d["cat"], d["events"][-1]["action"]), ("oddiy", "admin:bekor"))


if __name__ == "__main__":
    unittest.main()
