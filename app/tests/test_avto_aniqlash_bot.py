# -*- coding: utf-8 -*-
"""9-talab (dizayn/AVTO-ANIQLASH.md): ko'chirishni qoldiq o'zgarishidan aniqlash, har soatlik avto-yangilash;
11: A / B / C belgisi «📥 Keladigan» da va ko'chirish Excellarida."""
import datetime
import io
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fixtures import B, BotCase, ADMIN, CH, UCH, EK, F_CH, F_EK, F_UCH, TODAY, cb  # noqa: E402
from uzkir import strip_marks  # noqa: E402
from openpyxl import load_workbook  # noqa: E402

COL = {CH: 0, UCH: 2, EK: 4}  # fixture: Sale 6-qatori E6 = ЧИНОЗ, G6 = УРГУТ, I6 = ЭРКИН, K6 = СКЛАД


def at(h, m=0):
    return datetime.datetime.combine(TODAY, datetime.time(h, m))


class Base(BotCase):
    def sale(self, name):
        return next(r for r in self.bot.data.sale if r["name"] == name)

    def stock(self, name, **by):
        """stock(«ДОРИ», ch=20, uch=0) — Sale varag'idagi filial qoldig'i (hozirgi)."""
        r = self.sale(name)
        for k, v in by.items():
            r["v"][COL[{"ch": CH, "uch": UCH, "ek": EK}[k]]] = float(v)
        r["stock"] = r["v"][0] + r["v"][2] + r["v"][4] + r["v"][6]

    def task(self, name):
        return self.bot.db.execute("SELECT * FROM koch_tasks WHERE name=? ORDER BY id DESC", (name,)).fetchone()

    def admin_warns(self, part):
        return [t for t in self.tg.texts(ADMIN) if part in t]

    def add_task(self, name, src, dst, plan, s0, d0, status="yangi", producer="ВЕНГРИЯ"):
        with self.bot.dblock:
            cur = self.bot.db.execute(
                "INSERT INTO koch_tasks(sana, src, dst, name, producer, cat, qty_plan, prih, summa, status, created_at, kkey, "
                "src_stock0, dst_stock0) VALUES(?,?,?,?,?,'oddiy',?,1000,?,?,?,?,?,?)",
                (TODAY.isoformat(), B.norm(src), B.norm(dst), name, producer, plan, plan * 1000, status, at(8).isoformat(),
                 "%s|%s|%s" % (src, dst, name), s0, d0))
            self.bot.db.commit()
        return cur.lastrowid


class StockSnapshot(Base):
    def test_yaratilganda_saqlanadi_regen_ozgartirmaydi(self):
        self.stock("ОДДИЙ ЧАР", ch=20, uch=0)
        self.bot.koch_build()
        t = self.task("ОДДИЙ ЧАР")
        self.assertEqual((t["src_stock0"], t["dst_stock0"]), (20, 0), "Sale varag'idan (beruvchi / oluvchi ustuni)")
        self.assertEqual(t["src_stock"], 10, "Ko'chirish varag'idagi qiymat alohida")
        self.stock("ОДДИЙ ЧАР", ch=25, uch=0)  # ko'chirish izi yo'q (beruvchida ko'paydi) — reja yangilanadi
        row = next(r for r in self.bot.data.koch if r["name"] == "ОДДИЙ ЧАР")
        row["qty"], row["src_stock"], row["dst_stock"] = 6, 4, 9
        self.bot.koch_build(regen=True, by=(None, "admin"))
        t = self.task("ОДДИЙ ЧАР")
        self.assertEqual((t["qty_plan"], t["src_stock"], t["dst_stock"]), (6, 4, 9), "qayta shakllantirish yangiladi")
        self.assertEqual((t["src_stock0"], t["dst_stock0"]), (20, 0), "surat o'zgarmadi")

    def test_eski_qatorlar_migratsiyasi(self):
        self.bot.koch_build()
        with self.bot.dblock:
            self.bot.db.execute("UPDATE koch_tasks SET src_stock0=NULL, dst_stock0=NULL")
            self.bot.db.commit()
        self.bot._init_db()
        bad = self.bot.db.execute("SELECT COUNT(*) FROM koch_tasks WHERE src_stock0 IS NOT src_stock OR dst_stock0 IS NOT dst_stock")
        self.assertEqual(bad.fetchone()[0], 0, "mavjud qatorlar: hozirgi qiymatdan nusxa")
        cols = {r["name"] for r in self.bot.db.execute("PRAGMA table_info(koch_tasks)")}
        self.assertTrue({"src_stock0", "dst_stock0", "auto_by", "auto_note"} <= cols)

    def test_filial_ustuni(self):
        self.assertEqual([self.bot.stock_col(x) for x in (CH, UCH, EK, "ДОРИХОНА  ЭРКИН", "ДОРИХОНА СКЛАД", "НОМАЪЛУМ")],
                         [0, 2, 4, 4, 6, None])


class Detect(Base):
    def setUp(self):
        super().setUp()
        self.stock("ОДДИЙ ЧАР", ch=20, uch=0)   # ЧИНОЗ → УРГУТ, reja 8
        self.stock("АВИТЕН", ek=10, ch=0)       # ЭРКИН → ЧИНОЗ, reja 3
        self.bot.koch_build()

    def test_olindi_tizim(self):
        self.clock.at(10, 5)
        self.stock("ОДДИЙ ЧАР", ch=12, uch=8)
        res = self.bot.koch_auto_detect()
        self.assertEqual(res["olindi"], 1)
        t = self.task("ОДДИЙ ЧАР")
        self.assertEqual((t["status"], t["qty_given"], t["qty_received"], t["auto_by"]), ("olindi", 8, 8, "tizim"))
        self.assertEqual((t["given_at"][11:16], t["received_at"][11:16]), ("10:05", "10:05"))
        self.assertEqual(t["auto_note"], "ЧИНОЗ −8 · УРГУТ +8")
        ev = self.bot.db.execute("SELECT * FROM koch_events WHERE task_id=?", (t["id"],)).fetchall()
        self.assertEqual([(e["action"], e["user_name"], e["qty"]) for e in ev], [("tizim:olindi", "tizim", 8)])
        self.assertEqual(self.bot.koch_by(t), "tizim")
        item = [i for i in self.bot.koch_abc_summary()["done"] if i["id"] == t["id"]][0]
        self.assertEqual((item["by"], item["auto_note"]), ("tizim", "ЧИНОЗ −8 · УРГУТ +8"))
        self.assertEqual(self.bot.koch_task_dict(t)["by"], "tizim")
        card = strip_marks(self.bot.koch_item_card(self.bot.client(F_UCH), t["id"])[0])
        self.assertIn("Holat: <b>✅ olindi</b> 10:05 · 🤖 tizim", card)
        self.assertIn("🤖 Qoldiq: ЧИНОЗ −8 · УРГУТ +8", card)
        self.assertEqual(self.bot.koch_auto_detect()["olindi"], 0, "qayta aniqlanmaydi")

    def test_faqat_beruvchida_kamaydi_holat_ozgarmaydi_keyin_olindi(self):
        self.stock("ОДДИЙ ЧАР", ch=13)  # −7 (≥ 80%), oluvchida o'zgarmadi — sotuv bo'lishi mumkin
        n0 = len(self.tg.texts(ADMIN))
        res = self.bot.koch_auto_detect()
        self.assertEqual((res["ishora"], res["olindi"]), (1, 0))
        self.assertNotIn("berildi", res)
        t = self.task("ОДДИЙ ЧАР")
        self.assertEqual((t["status"], t["qty_given"], t["auto_by"], t["given_by"]), ("yangi", None, None, None),
                         "yolg'iz kamayish «berildi» qilmaydi")
        self.assertEqual(len(self.tg.texts(ADMIN)), n0, "adminga xabar yo'q")
        self.assertIn("Yo'lda (berilgan): <b>0</b>", self.bot.incoming_list(UCH)[0], "oluvchida yo'lda emas, kutilmoqda")
        self.stock("ОДДИЙ ЧАР", uch=7)
        self.assertEqual(self.bot.koch_auto_detect()["olindi"], 1)
        t = self.task("ОДДИЙ ЧАР")
        self.assertEqual((t["status"], t["qty_given"], t["qty_received"]), ("olindi", 7, 7))

    def test_kam_ozgarish_hech_narsa(self):
        self.stock("ОДДИЙ ЧАР", ch=14, uch=1)  # −6 < 8 × 0,8
        self.assertEqual(sum(self.bot.koch_auto_detect().values()), 0)
        self.assertEqual(self.task("ОДДИЙ ЧАР")["status"], "yangi")

    def test_tekshiring_bir_marta(self):
        self.stock("АВИТЕН", ch=3)  # faqat oluvchida ko'paydi
        self.assertEqual(self.bot.koch_auto_detect()["tekshiring"], 1)
        self.assertEqual(self.task("АВИТЕН")["status"], "yangi", "holat o'zgarmaydi")
        self.assertEqual(len(self.admin_warns("Tekshiring")), 1)
        self.assertIn("АВИТЕН", self.admin_warns("Tekshiring")[0])
        self.assertEqual(self.bot.koch_auto_detect()["tekshiring"], 0)
        self.assertEqual(len(self.admin_warns("Tekshiring")), 1, "bir topshiriq — bitta ogohlantirish")

    def test_qolda_holat_pasaytirilmaydi(self):
        give, rad = self.task("ОДДИЙ ЧАР"), self.task("ШОШ ДОРИ 01")
        self.assertTrue(self.bot.koch_refuse(F_CH, self.bot.client(F_CH), rad["id"], "Tovar topilmadi"))
        self.assertTrue(self.bot.koch_give(F_CH, self.bot.client(F_CH), give["id"], 5))
        self.stock("ОДДИЙ ЧАР", ch=15)          # faqat beruvchida −5 → «berildi» (o'zi shunday)
        self.stock("ШОШ ДОРИ 01", ch=0, uch=9)  # rad etilgan: dalil bo'lsa ham tegilmaydi
        self.bot.koch_auto_detect()
        g = self.task("ОДДИЙ ЧАР")
        self.assertEqual((g["status"], g["qty_given"], g["auto_by"], g["given_by"]), ("berildi", 5, None, F_CH))
        self.assertEqual(self.bot.koch_by(g), "Xodim Chinoz", "manba — xodim")
        self.assertEqual(self.task("ШОШ ДОРИ 01")["status"], "rad")
        # admin qo'lda «yangi» qilgan — tizim qayta tegmaydi
        av = self.task("АВИТЕН")
        self.assertIsNone(self.bot.koch_set_status(av["id"], "yangi", "qayta tekshiramiz"))
        self.stock("АВИТЕН", ek=7, ch=3)
        self.bot.koch_auto_detect()
        self.assertEqual(self.task("АВИТЕН")["status"], "yangi")
        # olindi (qo'lda) — umuman ko'rilmaydi
        self.stock("ОДДИЙ ЧАР", uch=5)
        self.assertTrue(self.bot.koch_receive(F_UCH, self.bot.client(F_UCH), g["id"]))
        self.bot.koch_auto_detect()
        g = self.task("ОДДИЙ ЧАР")
        self.assertEqual((g["status"], g["received_by"], g["auto_by"]), ("olindi", F_UCH, None))

    def test_berdim_qoldiq_ozgarmadi_2_yangilashdan_keyin(self):
        when = {"t": at(8, 5)}
        self.bot.data.when = lambda: when["t"]
        self.clock.at(8, 10)
        t = self.task("ОДДИЙ ЧАР")
        self.assertTrue(self.bot.koch_give(F_CH, self.bot.client(F_CH), t["id"]))
        self.bot.koch_auto_detect()  # ma'lumot «Berdim» dan oldingi — sanalmaydi
        self.assertFalse(self.task("ОДДИЙ ЧАР")["auto_idle"])
        when["t"] = at(9)
        self.clock.at(9, 2)
        self.bot.koch_auto_detect()
        self.bot.koch_auto_detect()  # o'sha ma'lumot — ikkinchi marta sanalmaydi
        self.assertEqual(self.task("ОДДИЙ ЧАР")["auto_idle"], 1)
        self.assertFalse(self.admin_warns("Berdim» dedi"))
        when["t"] = at(10)
        self.assertEqual(self.bot.koch_auto_detect()["berdim_warn"], 1)
        w = self.admin_warns("Berdim» dedi")
        self.assertEqual(len(w), 1)
        self.assertTrue(w[0].startswith("⚠ <b>Tekshiring</b> · ko'chirish: <b>1</b> ta"), w[0])
        self.assertIn("%s «Berdim» dedi (08:10), qoldiq o'zgarmadi" % CH, w[0])
        self.assertIn("ОДДИЙ ЧАР", w[0])
        when["t"] = at(11)
        self.bot.koch_auto_detect()
        self.assertEqual(len(self.admin_warns("Berdim» dedi")), 1, "bir marta")
        self.assertEqual(self.task("ОДДИЙ ЧАР")["status"], "berildi")

    def test_berdim_beruvchida_kamaygan_ogohlantirmaydi(self):
        when = {"t": at(8, 5)}
        self.bot.data.when = lambda: when["t"]
        self.clock.at(8, 10)
        self.assertTrue(self.bot.koch_give(F_CH, self.bot.client(F_CH), self.task("ОДДИЙ ЧАР")["id"]))
        self.stock("ОДДИЙ ЧАР", ch=12)  # beruvchida kamaydi (ishora), oluvchi hali kiritmagan
        for h in (9, 10, 11, 12):
            when["t"] = at(h)
            self.clock.at(h, 2)
            self.bot.koch_auto_detect()
        self.assertFalse(self.admin_warns("Berdim» dedi"), "beruvchida qoldiq o'zgargan")
        self.assertEqual(self.task("ОДДИЙ ЧАР")["status"], "berildi")

    def test_nazorat_ochiq(self):
        self.bot.set_setting("koch_on", "0")
        self.stock("ОДДИЙ ЧАР", ch=12, uch=8)
        self.assertEqual(sum(self.bot.koch_auto_detect().values()), 0)


class Share(Base):
    def test_bitta_oluvchi_ikki_beruvchi_ikki_marta_sanalmaydi(self):
        a = self.add_task("ОДДИЙ 04", CH, EK, 4, 10, 0)
        b = self.add_task("ОДДИЙ 04", UCH, EK, 4, 10, 0)
        self.stock("ОДДИЙ 04", ch=6, uch=6, ek=4)  # oluvchiga hammasi bo'lib +4 keldi (ikkala topshiriqqa emas)
        self.bot.koch_auto_detect()
        st = {t["id"]: (t["status"], t["qty_given"]) for t in self.bot.db.execute("SELECT * FROM koch_tasks WHERE id IN (?,?)", (a, b))}
        self.assertEqual(st, {a: ("yangi", None), b: ("yangi", None)}, "+4 ikkiga bo'lindi (2 + 2 < 80%) — olindi emas; "
                         "yolg'iz beruvchidagi kamayish holatni o'zgartirmaydi")

    def test_bitta_beruvchi_reja_bo_yicha_bolinadi(self):
        a = self.add_task("ОДДИЙ 04", CH, EK, 6, 20, 0)
        b = self.add_task("ОДДИЙ 04", CH, UCH, 2, 20, 0)
        self.stock("ОДДИЙ 04", ch=12, ek=6, uch=2)
        self.bot.koch_auto_detect()
        st = {t["id"]: (t["status"], t["qty_received"]) for t in self.bot.db.execute("SELECT * FROM koch_tasks WHERE id IN (?,?)", (a, b))}
        self.assertEqual(st, {a: ("olindi", 6), b: ("olindi", 2)})
        self.assertEqual(self.bot.db.execute("SELECT auto_note FROM koch_tasks WHERE id=?", (a,)).fetchone()[0],
                         "ЧИНОЗ −6 · ЭРКИН +6", "−8 reja bo'yicha (6 : 2) bo'lindi")


OK_STATUS = "START\nOK Sale|5\nOK Srok|4\nOK 30+ kun|0\nOK Narx|3\nOK Hisob|5\nOK Ko'chirish|27\nDONE\n"
FAIL_STATUS = "START\nFAIL [Microsoft][ODBC SQL Server Driver] SQL Server does not exist\n"


class Hourly(BotCase):
    sale, stock, task = Base.sale, Base.stock, Base.task

    def setUp(self):
        super().setUp()
        self.status = os.path.join(self.work, "refresh_status.txt")
        self.calls, self.ok = [], True
        self.bot.runner = self.fake_runner
        self.when = datetime.datetime.combine(TODAY - datetime.timedelta(days=1), datetime.time(18, 0))
        self.bot.data.when = lambda: self.when
        self.bot.set_setting("auto_on", "0")
        for k, v in (("auto_refresh_every", 60), ("auto_refresh_from", "08:00"), ("auto_refresh_to", "20:00")):
            self.bot.set_setting(k, v)
        self.detects, orig = [], self.bot.koch_auto_detect
        self.bot.koch_auto_detect = lambda: self.detects.append(B.now()) or orig()
        self.threads, bg = [], self.bot.bg
        self.bot.bg = lambda fn, *a, name="ish": self.threads.append(bg(fn, *a, name=name)) or self.threads[-1]

    def fake_runner(self, vbs, xlsm, timeout):
        self.calls.append(B.now())
        with open(self.status, "w", encoding="utf-8") as f:
            f.write(OK_STATUS if self.ok else FAIL_STATUS)
        if self.ok:
            self.when = B.now()
        return (0, "") if self.ok else (1, "")

    def cycle(self, h, m=0):
        self.clock.at(h, m)
        self.bot.run_scheduled()
        for _ in range(3):  # yangilash → (topshiriq, aniqlash) va avto-ishlar — hammasi fon oqimlarida
            for t in list(self.threads):
                t.join(10)
        self.assertFalse(self.bot.refreshing or self.bot.auto_running)
        return len(self.calls)

    def test_oraliq_ish_vaqti_va_ertalabki_bilan_toqnashmaydi(self):
        self.assertEqual(self.cycle(8, 10), 0, "ertalabki avto-yangilash (08:30) dan oldin har soatlik kutadi")
        self.assertEqual(self.cycle(8, 40), 1, "ertalabki avto-yangilash")
        self.assertEqual(self.bot.setting("last_auto_refresh_try"), TODAY.isoformat())
        self.assertIsNone(self.bot.setting("last_hourly_refresh"), "bu ertalabki edi, har soatlik emas")
        self.assertTrue(self.bot.koch_has_tasks(), "kunlik topshiriq shakllandi")
        self.assertEqual(self.cycle(9, 0), 1, "oraliq (60 daq.) o'tmagan")
        self.assertEqual(self.cycle(9, 41), 2, "har soatlik")
        self.assertTrue(self.bot.setting("last_hourly_refresh").startswith(TODAY.isoformat() + "T09:41"))
        self.assertEqual(self.cycle(10, 20), 2)
        self.assertEqual(len([t for t in self.tg.texts(F_CH) if "topshirig'i tayyor" in t]), 1, "ertalabki topshiriq xabari bir marta")
        self.assertEqual(len(self.detects), 2, "har muvaffaqiyatli yangilashdan keyin avto-aniqlash")
        self.assertFalse([t for t in self.tg.texts(ADMIN) if "Avto-yangilash" in t or "Yangilandi" in t or "bajarilmadi" in t],
                         "avto / har soatlik yangilash — chatga xabar yo'q")
        self.assertEqual(self.cycle(10, 45), 3)
        self.assertEqual(self.cycle(20, 30), 3, "ish vaqtidan tashqari")
        self.bot.set_setting("auto_refresh_every", 0)
        self.assertEqual(self.cycle(19, 59), 3, "o'chiq")

    def test_muvaffaqiyatsiz_har_soatlik_jim(self):
        self.bot.set_setting("auto_refresh_time", "")  # ertalabki o'chiq: har soatlik darhol
        self.ok = False
        self.assertEqual(self.cycle(11, 0), 1)
        self.assertFalse([t for t in self.tg.texts(ADMIN) if "bajarilmadi" in t], "faqat jurnal")
        self.assertEqual(self.detects, [], "muvaffaqiyatsiz yangilashdan keyin aniqlanmaydi")
        self.assertEqual(self.cycle(11, 30), 1)
        self.ok = True
        self.assertEqual(self.cycle(12, 1), 2)
        self.assertEqual(len(self.detects), 1)

    def test_yangilashdan_keyin_avval_aniqlash_keyin_qayta_shakllantirish(self):
        """Ko'chirilgan qator makros ro'yxatidan chiqadi: qayta shakllantirish uni «bekor» qilishidan oldin aniqlanadi.
        refreshing aniqlash va qayta shakllantirish tugaguncha ushlab turiladi."""
        self.bot.set_setting("auto_refresh_time", "")
        self.bot.data.reload = lambda force=False: True  # Sale / Ko'chirish xotirada (quyida «Excel» o'zgartiradi)
        self.stock("ОДДИЙ ЧАР", ch=20, uch=0)
        self.when = at(8, 30)
        self.clock.at(8, 31)
        self.bot.koch_build()
        orig_runner, seen = self.fake_runner, []

        def excel(vbs, xlsm, timeout):
            self.stock("ОДДИЙ ЧАР", ch=12, uch=8)  # tugma bosilmadi, lekin ko'chirildi
            self.bot.data.koch = [r for r in self.bot.data.koch if r["name"] != "ОДДИЙ ЧАР"]
            return orig_runner(vbs, xlsm, timeout)
        self.bot.runner = excel
        build = self.bot.koch_build
        self.bot.koch_build = lambda *a, **k: seen.append(("build", self.bot.refreshing, len(self.detects))) or build(*a, **k)
        self.clock.at(11, 0)
        self.assertTrue(self.bot.start_refresh(None))
        for t in list(self.threads):
            t.join(10)
        t = self.task("ОДДИЙ ЧАР")
        self.assertEqual((t["status"], t["auto_by"], t["qty_received"]), ("olindi", "tizim", 8), "bekor emas — aniqlandi")
        self.assertEqual(seen, [("build", True, 1)], "qayta shakllantirish aniqlashdan keyin, refreshing hali True")
        self.assertFalse(self.bot.refreshing)

    def test_qolda_yangilash_oraliqni_boshlaydi(self):
        self.bot.set_setting("auto_refresh_time", "")
        self.clock.at(13, 0)
        self.assertTrue(self.bot.start_refresh(None))  # admin «Yangilash»
        for t in list(self.threads):
            t.join(10)
        self.assertEqual(self.cycle(13, 30), 1, "qo'lda yangilashdan 30 daqiqa o'tdi — hali emas")
        self.assertEqual(self.cycle(14, 0), 2)


class ABC(Base):
    def setUp(self):
        super().setUp()
        self.bot.koch_build()
        self.cls = self.bot.koch_abc()

    def test_keladigan_abc(self):
        av = self.task("АВИТЕН")
        k = self.cls[av["id"]]
        text = self.bot.incoming_list(CH)[0]
        self.assertIn("• 🔴 АВИТЕН · 3 qadoq ← %s · <code>%s</code>" % (EK, k), text, "kutilmoqda")
        self.bot.handle(cb(F_EK, "kb:%d" % av["id"]))
        text = self.bot.incoming_list(CH)[0]
        self.assertIn("1. 🔴 АВИТЕН · <b>3</b> qadoq ← %s · <code>%s</code>" % (EK, k), text, "berilgan")

    def test_filial_exceli_abc(self):
        self.bot.handle(cb(F_CH, "kx:shoshilinch"))
        wb = load_workbook(io.BytesIO(self.tg.to(F_CH, "doc")[-1][3]))
        ws = wb[CH]
        head = [c.value for c in ws[1]]
        self.assertEqual(head[-1], "ABC")
        name_col, abc_col = head.index("Dori") + 1, len(head)
        got = {ws.cell(r, name_col).value: ws.cell(r, abc_col).value for r in range(2, ws.max_row)}
        self.assertTrue(got)
        for name, k in got.items():
            self.assertEqual(k, self.cls[self.task(name)["id"]], name)
        self.assertEqual(ws.cell(ws.max_row, 1).value, "JAMI")
        self.assertNotIn("ABC", [c.value for c in wb["Ko'chirilmaydi"][1]])

    def test_nazorat_exceli_abc_va_manba(self):
        self.stock("ОДДИЙ ЧАР", ch=self.task("ОДДИЙ ЧАР")["src_stock0"] - 8, uch=self.task("ОДДИЙ ЧАР")["dst_stock0"] + 8)
        self.bot.koch_auto_detect()
        self.bot.handle(cb(ADMIN, "a:trx"))
        ws = load_workbook(io.BytesIO(self.tg.to(ADMIN, "doc")[-1][3]))["Topshiriqlar"]
        head = [c.value for c in ws[1]]
        self.assertEqual(head[-3:], ["Manba", "Isbot", "ABC"])
        rows = {r[head.index("Dori")]: r for r in ws.iter_rows(min_row=2, values_only=True)}
        r = rows["ОДДИЙ ЧАР"]
        self.assertEqual((r[head.index("Manba")], r[head.index("Holat")]), ("tizim", "olindi"))
        self.assertEqual(r[-1], self.cls[self.task("ОДДИЙ ЧАР")["id"]])
        self.assertTrue(all(row[-1] in ("A", "B", "C") for row in rows.values()))


if __name__ == "__main__":
    unittest.main()
