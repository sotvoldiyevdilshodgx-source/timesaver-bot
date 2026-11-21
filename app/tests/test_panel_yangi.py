# -*- coding: utf-8 -*-
"""Server paneli: kunlik surat (trend), «Кўчириш · ABC», buyurtma so'rovlari, avtomatik yangilash sozlamalari,
tizim yorliqlari kirillda. Vaqtinchalik baza va 127.0.0.1 dagi vaqtinchalik port."""
import base64
import datetime
import io
import json
import os
import sys
import threading
import unittest
import urllib.error
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fixtures import B, BotCase, CH, EK, F_CH, F_EK, FIRM, TODAY
import timesaver_panel as P
import buyurtma as BY
from openpyxl import load_workbook


def snap(bot, sana, srok, old=1, users=2, pending=0):
    with bot.dblock:
        bot.db.execute("INSERT OR REPLACE INTO daily_stats VALUES(?,?,?,?,?)", (sana, srok, old, users, pending))
        bot.db.commit()


class Trend(BotCase):
    def setUp(self):
        super().setUp()
        self.api = P.Api(self.bot, self.bot.cfg)

    def test_jadval_va_bugungi_surat(self):
        s = self.api.state()
        self.assertNotIn("trend", s, "faqat bugungi surat — trend chizig'i yo'q")
        r = self.bot.db.execute("SELECT * FROM daily_stats").fetchall()
        self.assertEqual([(x["sana"], x["srok"], x["users"]) for x in r], [(TODAY.isoformat(), s["srok"], s["users"])])
        self.api.state()
        self.assertEqual(self.bot.db.execute("SELECT COUNT(*) FROM daily_stats").fetchone()[0], 1, "kuniga bitta qator")

    def test_14_kun_boshliqlar_toldiriladi(self):
        d = lambda n: (TODAY - datetime.timedelta(days=n)).isoformat()
        snap(self.bot, d(20), 50)   # oynadan oldingi oxirgi qiymat — boshidan olib boriladi
        snap(self.bot, d(10), 40)
        snap(self.bot, d(5), 30)
        s = self.api.state()
        t = s["trend"]
        self.assertEqual(set(t), set(P.STAT_KEYS))
        self.assertTrue(all(len(v) == 14 for v in t.values()))
        self.assertEqual(t["srok"][:3], [50, 50, 50])
        self.assertEqual(t["srok"][3], 40)              # 13 - 10 = 3-indeks
        self.assertEqual(t["srok"][8:13], [30] * 5)
        self.assertEqual(t["srok"][-1], s["srok"])
        self.assertEqual(t["users"][-1], s["users"])

    def test_tarix_qisqa(self):
        snap(self.bot, (TODAY - datetime.timedelta(days=2)).isoformat(), 9)
        t = self.api.state()["trend"]
        self.assertEqual(len(t["srok"]), 3, "tarix boshlanmagan kunlar qo'shilmaydi")
        self.assertEqual(t["srok"][:2], [9, 9])


class KochAbc(BotCase):
    def setUp(self):
        super().setUp()
        self.api = P.Api(self.bot, self.bot.cfg)
        self.bot.koch_build()

    def tid(self, name):
        return self.bot.db.execute("SELECT id FROM koch_tasks WHERE name=?", (name,)).fetchone()[0]

    def test_shakl_va_sinflar(self):
        t1 = self.tid("АВИТЕН")       # 3 × 20000 = 60 000 — eng qimmatlaridan
        self.bot.koch_give(F_EK, self.bot.client(F_EK), t1)
        s = self.api.koch_abc("bugun")
        self.assertEqual((s["period"], s["dates"], s["has"]), ("bugun", [TODAY.isoformat()], True))
        self.assertEqual(set(s["classes"]), {"A", "B", "C"})
        tasks = [t for t in self.bot.koch_tasks(TODAY.isoformat()) if t["status"] != "bekor"]
        self.assertEqual(s["jami"], len(tasks))
        self.assertEqual(sum(c["jami"] for c in s["classes"].values()), len(tasks))
        self.assertEqual(s["bajarildi"], 1)
        self.assertEqual(sum(c["share"] for c in s["classes"].values()), 100)
        self.assertGreaterEqual(s["classes"]["A"]["share"], 75)
        d = s["done"][0]
        self.assertEqual((d["name"], d["cls"], d["status"], d["src"], d["dst"], d["qty"], d["sum"]),
                         ("АВИТЕН", "A", "berildi", self.bot.db.execute("SELECT src FROM koch_tasks WHERE id=?", (t1,)).fetchone()[0],
                          CH, 3, 60000))
        self.assertEqual(d["time"], "08:00")
        for it in s["not_done"]:
            self.assertIn(it["status"], P.KOCH_NOT_DONE)
            self.assertEqual(set(("cls", "name", "src", "dst", "qty", "sum", "status", "time")) - set(it), set())
        # sinf ichida pul bo'yicha kamayish tartibi
        a = [it["sum"] for it in s["not_done"] if it["cls"] == "A"]
        self.assertEqual(a, sorted(a, reverse=True))
        # matritsa: beruvchi filial × sinf → [bajarildi, jami]
        tot = sum(v[c][1] for v in s["filials"].values() for c in "ABC")
        self.assertEqual(tot, len(tasks))
        self.assertEqual(sum(v[c][0] for v in s["filials"].values() for c in "ABC"), 1)

    def test_rad_va_bekor(self):
        t = self.tid("ШОШ ДОРИ 01")
        self.bot.koch_refuse(F_CH, self.bot.client(F_CH), t, "Tovar topilmadi")
        b = self.tid("ОДДИЙ ЧАР")
        self.api.koch_set_status({"id": b, "status": "bekor", "note": "test"})
        s = self.api.koch_abc("bugun")
        names = [i["name"] for i in s["done"] + s["not_done"]]
        self.assertNotIn("ОДДИЙ ЧАР", names, "bekor — hisobga kirmaydi")
        r = next(i for i in s["not_done"] if i["name"] == "ШОШ ДОРИ 01")
        self.assertEqual((r["status"], r["reason"]), ("rad", "Tovar topilmadi"))

    def test_kechikish(self):
        self.clock.at(11, 0)
        s = self.api.koch_abc("bugun")
        late = [i for i in s["not_done"] if i.get("late_min")]
        self.assertTrue(late)
        self.assertEqual(late[0]["late_min"], 60, "SHOSHILINCH muddati 10:00")

    def test_kecha_va_7_kun(self):
        s = self.api.koch_abc("kecha")
        self.assertEqual((s["has"], s["jami"]), (False, 0))
        self.assertEqual(s["dates"], [(TODAY - datetime.timedelta(days=1)).isoformat()])
        s7 = self.api.koch_abc("7")
        self.assertEqual(len(s7["dates"]), 7)
        self.assertEqual(s7["jami"], self.api.koch_abc("bugun")["jami"])
        self.assertEqual(self.api.koch_abc("nimadir")["period"], "bugun")

    def test_bot_metodi_bolsa_ishlatiladi(self):
        calls = []

        def fake(sana=None, filial=None):
            calls.append(sana)
            return {"classes": {"A": {"jami": 2, "bajarildi": 1, "bajarilmadi": 1, "sum_jami": 900, "sum_bajarildi": 600,
                                      "sum_bajarilmadi": 300}},
                    "filials": {CH: {"A": [1, 2]}},
                    "done": [{"cls": "A", "name": "X", "src": CH, "dst": EK, "qty": 1, "sum": 600, "status": "olindi", "time": "09:00"}],
                    "not_done": [{"cls": "A", "name": "Y", "src": CH, "dst": EK, "qty": 1, "sum": 300, "status": "yangi", "time": ""}]}

        self.bot.koch_abc_summary = fake
        try:
            s = self.api.koch_abc("7")
        finally:
            del self.bot.koch_abc_summary
        self.assertEqual(len(calls), 7, "7 kun — har kun alohida")
        self.assertEqual(s["classes"]["A"]["jami"], 14)
        self.assertEqual(s["classes"]["B"]["jami"], 0)
        self.assertEqual(s["filials"][CH]["A"], [7, 14])
        self.assertEqual(s["filials"][CH]["C"], [0, 0])
        self.assertEqual(len(s["done"]), 7)
        self.assertEqual(s["classes"]["A"]["share"], 100)
        self.assertEqual(s["done"][0]["sana"], calls[0])

    def test_zaxira_hisob_bot_bilan_bir_xil(self):
        """bot.koch_abc_summary bo'lmasa panel o'zi hisoblaydi — natija bir xil bo'lishi kerak."""
        t1 = self.tid("АВИТЕН")
        self.bot.koch_give(F_EK, self.bot.client(F_EK), t1)
        self.clock.at(10, 30)
        via_bot = self.api.koch_abc("bugun")
        self.bot.koch_abc_summary = None   # metod yo'qdek
        try:
            local = self.api.koch_abc("bugun")
        finally:
            del self.bot.koch_abc_summary
        if not callable(getattr(self.bot, "koch_abc_summary", None)):
            self.skipTest("bot.koch_abc_summary hali yo'q")
        for c in "ABC":
            for k in ("jami", "bajarildi", "bajarilmadi", "share"):
                self.assertEqual(via_bot["classes"][c][k], local["classes"][c][k], (c, k))
        self.assertEqual(via_bot["filials"], local["filials"])
        strip = lambda lst: [(i["name"], i["cls"], i["status"], i["qty"], round(i["sum"]), i["time"], i.get("late_min")) for i in lst]
        self.assertEqual(sorted(strip(via_bot["not_done"])), sorted(strip(local["not_done"])))
        self.assertEqual(strip(via_bot["done"]), strip(local["done"]))

    def test_excel(self):
        content, fname = self.api.koch_abc_excel("bugun")
        self.assertEqual(fname, "Kochirish_ABC_%s.xlsx" % TODAY.isoformat())
        ws = load_workbook(io.BytesIO(content)).active
        self.assertEqual(ws.title, "Кўчириш ABC")
        self.assertEqual([c.value for c in ws[1]][:4], ["Сана", "Гуруҳ", "Натижа", "Дори"])
        self.assertEqual(ws.max_row - 1, self.api.koch_abc("bugun")["jami"])
        self.assertIsNone(self.api.koch_abc_excel("kecha"))

    def test_eslatma_filiallar(self):
        r = self.api.koch_remind({"filials": [CH, CH]})
        self.assertTrue(r["ok"])
        self.assertIn("1 filial", r["msg"])
        self.assertFalse(self.api.koch_remind({"filials": []})["ok"])


class Orders(BotCase):
    def setUp(self):
        super().setUp()
        self.api = P.Api(self.bot, self.bot.cfg)
        self.rid = BY.create(self.bot.db, FIRM, "АЛЬФА ФАРМ", "Анвар Каримов",
                             [{"name": "ПАРАЦЕТАМОЛ 500МГ №10", "producer": "UZ", "qty": 20},
                              {"name": "АНАЛЬГИН №10", "qty": 15}, {"name": "НОШПА", "qty": 5}])

    def test_jadvallar_init(self):
        names = {r[0] for r in self.bot.db.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        self.assertTrue({"order_req", "order_req_items", "daily_stats"} <= names)

    def test_royxat_filtr(self):
        r = self.api.oreq_list("kutmoqda")
        self.assertEqual(r["counts"], {"kutmoqda": 1, "hal": 0, "all": 1})
        it = r["items"][0]
        self.assertEqual((it["id"], it["firm"], it["name"], it["n"], it["asked"], it["ok"], it["status"]),
                         (self.rid, "АЛЬФА ФАРМ", "Анвар Каримов", 3, 40, 0, "kutmoqda"))
        self.assertEqual(self.api.oreq_list("hal")["items"], [])
        self.assertEqual(len(self.api.oreq_list("all")["items"]), 1)

    def test_qaror_va_javob(self):
        self.assertFalse(self.api.oreq_finish({"id": self.rid})["ok"], "hamma dori hal qilinmagan")
        r = self.api.oreq_decide({"id": self.rid, "pos": 1, "qty": 10})
        self.assertTrue(r["ok"])
        self.assertEqual([(i["status"], i["qty_ok"]) for i in r["items"]][0], ("ozgardi", 10))
        self.assertFalse(self.api.oreq_decide({"id": self.rid, "pos": 1, "qty": "abc"})["ok"])
        self.assertFalse(self.api.oreq_decide({"id": self.rid, "pos": 9})["ok"])
        r = self.api.oreq_decide({"id": self.rid, "pos": 3, "reject": True})
        self.assertEqual(r["items"][2]["status"], "rad")
        r = self.api.oreq_decide({"id": self.rid, "pos": 2})
        self.assertEqual(r["items"][1]["status"], "qabul")
        self.assertTrue(r["all_decided"])
        r = self.api.oreq_finish({"id": self.rid})
        self.assertTrue(r["ok"])
        self.assertEqual(r["msg"], "Javob yuborildi", "bot.order_notify ulangan: vakilga Telegram xabari ketdi")
        self.assertTrue([t for t in self.tg.texts(FIRM) if "So'rov №%d bo'yicha javob" % self.rid in t], "vakilga javob (push)")
        self.assertEqual((r["req"]["status"], r["req"]["ok"], r["req"]["changed"]), ("javob", 25, 1))
        self.assertFalse(self.api.oreq_finish({"id": self.rid})["ok"], "ikkinchi javob — rad")
        self.assertFalse(self.api.oreq_decide({"id": self.rid, "pos": 1, "qty": 5})["ok"])
        self.assertEqual(self.api.oreq_list("hal")["counts"]["hal"], 1)

    def test_hammasini_qabul_va_bildirish(self):
        sent = []
        self.bot.order_notify = sent.append
        try:
            self.assertTrue(self.api.oreq_accept_all({"id": self.rid})["ok"])
            r = self.api.oreq_finish({"id": self.rid})
        finally:
            del self.bot.order_notify
        self.assertEqual((r["ok"], sent), (True, [self.rid]))
        self.assertEqual(r["req"]["ok"], 40)

    def test_rad_etish(self):
        r = self.api.oreq_finish({"id": self.rid, "reject": True})
        self.assertTrue(r["ok"])
        self.assertEqual(r["req"]["status"], "rad")
        self.assertIsNone(self.api.oreq_excel(7), "rad etilganlar Excel'ga kirmaydi")

    def test_excel(self):
        self.api.oreq_decide({"id": self.rid, "pos": 1, "qty": 12})
        self.api.oreq_decide({"id": self.rid, "pos": 2, "reject": True})
        self.api.oreq_decide({"id": self.rid, "pos": 3})
        self.api.oreq_finish({"id": self.rid})
        content, fname = self.api.oreq_excel("7")
        self.assertRegex(fname, r"^Buyurtma_sorovlari_7kun_\d{4}-\d\d-\d\d\.xlsx$")
        ws = load_workbook(io.BytesIO(content)).active
        self.assertEqual([c.value for c in ws[1]], ["№", "Фирма", "Вакил", "Жавоб вақти", "Дори", "Ишлаб чиқарувчи",
                                                    "Сўралган", "Тасдиқланган", "Ҳолат"])
        rows = [[c.value for c in r] for r in ws.iter_rows(min_row=2)]
        self.assertEqual([(r[4], r[6], r[7], r[8]) for r in rows],
                         [("ПАРАЦЕТАМОЛ 500МГ №10", 20, 12, "Ўзгартирилди"), ("НОШПА", 5, 5, "Қабул қилинди")])
        self.assertEqual(rows[0][1], "АЛЬФА ФАРМ")

    def test_get(self):
        self.assertFalse(self.api.oreq_get("x")["ok"])
        self.assertFalse(self.api.oreq_get(999)["ok"])
        r = self.api.oreq_get(self.rid)
        self.assertEqual(([i["pos"] for i in r["items"]], r["all_decided"]), ([1, 2, 3], False))


class Settings(BotCase):
    def setUp(self):
        super().setUp()
        self.api = P.Api(self.bot, self.bot.cfg)

    def test_avto_yangilash(self):
        s = self.api.settings_get()
        self.assertEqual((s["auto_refresh_every"], s["auto_refresh_from"], s["auto_refresh_to"]), (0, "08:00", "20:00"))
        self.assertTrue(self.api.settings_set({"auto_refresh_every": 60})["ok"])
        self.assertEqual(self.bot.setting("auto_refresh_every"), "60")
        r = self.api.settings_set({"auto_refresh_from": "7:30", "auto_refresh_to": "21:00"})
        self.assertTrue(r["ok"])
        self.assertEqual((self.bot.setting("auto_refresh_from"), self.bot.setting("auto_refresh_to")), ("07:30", "21:00"))
        self.assertEqual(self.bot.setting("auto_refresh_every"), "60", "oraliq o'zgarmaydi")
        self.assertFalse(self.api.settings_set({"auto_refresh_every": 45})["ok"])
        self.assertFalse(self.api.settings_set({"auto_refresh_from": "25:00"})["ok"])
        self.assertFalse(self.api.settings_set({"auto_refresh_from": "21:00", "auto_refresh_to": "08:00"})["ok"])
        self.assertTrue(self.api.settings_set({"auto_refresh_every": 0})["ok"])
        self.assertEqual(self.api.settings_get()["auto_refresh_every"], 0)

    def test_yorliqlar_kirillda(self):
        s = self.api.settings_view()
        self.assertEqual(s["weekdays"][0], "Душанба")
        self.assertTrue(all(not any("a" <= ch.lower() <= "z" for ch in a["name"]) for a in s["alerts"]), s["alerts"])
        self.assertEqual(self.api.settings_get()["weekdays"][0], "Dushanba", "ichki qiymat o'zgarmaydi")


class StockAbc(BotCase):
    """Bosh sahifa: Срок / 30+ кун · A / B / C (6-talab, panel tomoni)."""
    def setUp(self):
        super().setUp()
        self.api = P.Api(self.bot, self.bot.cfg)

    def test_shakl(self):
        for kind in ("s", "o"):
            s = self.api.stock_abc(kind)
            rows = self.api.dash_rows(kind)  # butun tarmoq — dori bo'yicha jamlangan
            self.assertTrue(s["jamlangan"])
            self.assertEqual(s["kind"], kind)
            self.assertEqual(s["jami"], len(rows))
            self.assertEqual(sum(s["classes"][c]["n"] for c in "ABC"), len(rows))
            self.assertEqual(s["filial"], "")
            for c in "ABC":
                part = s["rows"][c]
                self.assertLessEqual(len(part), P.Api.STOCK_ABC_TOP)
                self.assertEqual([r["sum"] for r in part], sorted([r["sum"] for r in part], reverse=True), "pul bo'yicha kamayadi")
            if rows:
                self.assertTrue(s["rows"]["A"], "A bo'sh emas")

    def test_filial(self):
        s = self.api.stock_abc("s")
        self.assertTrue(s["filials"])
        f = s["filials"][0]
        one = self.api.stock_abc("s", f)
        self.assertEqual(one["filial"], f)
        self.assertTrue(all(B.fil_key(r["filial"]) == B.fil_key(f) for c in "ABC" for r in one["rows"][c]))
        self.assertLessEqual(one["jami"], s["jami"])
        self.assertEqual(self.api.stock_abc("s", "YO'Q FILIAL")["filial"], "", "noma'lum filial — butun tarmoq")

    def test_excel(self):
        s = self.api.stock_abc("s")
        content, name = self.api.stock_abc_excel("s", "")
        ws = load_workbook(io.BytesIO(content)).active
        self.assertEqual(ws.max_row - 1, s["jami"])
        self.assertEqual(ws.cell(1, 1).value, "ABC")
        self.assertTrue(name.startswith("Srok_ABC_"))
        c = next(k for k in "ABC" if s["classes"][k]["n"])
        content, name = self.api.stock_abc_excel("s", c)
        ws = load_workbook(io.BytesIO(content)).active
        self.assertEqual(ws.max_row - 1, s["classes"][c]["n"])
        self.assertTrue(all(ws.cell(i, 1).value == c for i in range(2, ws.max_row + 1)))
        self.assertEqual(self.api.stock_abc_excel("o", "") is None, not self.api.dash_rows("o"), "bo'sh bo'lsa — fayl yo'q")

    def _jam_data(self):
        """ТЕНОВИР: 30+ kun varag'ida faqat B filial (5 dona, 114 kun); Sale bo'yicha jami 55 (A 10, B 5, sklad 40).
        ДОРИ2: hamma joyda 30–60 kun — 90+ ro'yxatiga kirmaydi."""
        d = self.bot.data
        today = datetime.date.today()
        d.sale_heads = ["ФИЛ А", "ФИЛ Б", "", "СКЛАД"]
        d.sale = [{"name": "ТЕНОВИР", "producer": "ИНДИЯ", "mnn": "", "firm": "F1", "v": [10, 1, 5, 1, 0, 0, 40],
                   "stock": 55, "sales": 2},
                  {"name": "ДОРИ2", "producer": "X", "mnn": "", "firm": "F1", "v": [3, 0, 0, 0, 0, 0, 0], "stock": 3, "sales": 0}]
        d.old = [{"filial": "ФИЛ Б", "name": "ТЕНОВИР", "producer": "ИНДИЯ", "supplier": "", "firm": "F1", "kun": 114,
                  "ost": 2, "prih": 250000, "summa": 500000},
                 {"filial": "ФИЛ Б", "name": "ТЕНОВИР", "producer": "ИНДИЯ", "supplier": "", "firm": "F1", "kun": 38,
                  "ost": 3, "prih": 250000, "summa": 750000},
                 {"filial": "ФИЛ А", "name": "ДОРИ2", "producer": "X", "supplier": "", "firm": "F1", "kun": 60,
                  "ost": 3, "prih": 1000, "summa": 3000}]
        srok = lambda f, days, ost, prih: {"filial": f, "name": "ТЕНОВИР", "producer": "ИНДИЯ", "firm": "F1",
                                           "date": today + datetime.timedelta(days=days), "kun": days, "ost": ost, "prih": prih}
        d.srok = [srok("ФИЛ Б", 11, 2, 250000), srok("ФИЛ Б", 41, 3, 250000), srok("ДОРИХОНА СКЛАД", 41, 40, 190000),
                  srok("ФИЛ А", 400, 10, 190000)]
        self.bot.data.reload = lambda force=False: True

    def test_jamlangan(self):
        self._jam_data()
        o = self.api.dash_rows("o")
        self.assertEqual([(r["name"], r["kun"], r["ost"]) for r in o], [("ТЕНОВИР", 114, 55)],
                         "90+ kun: biror filialda 114 kun → jami qoldiq barcha filiallardan; ДОРИ2 (60 kun) kirmaydi")
        self.assertEqual(o[0]["filials"], ["ДОРИХОНА СКЛАД", "ФИЛ А", "ФИЛ Б"])
        self.assertEqual(round(o[0]["summa"]), 55 * 250000)
        s = self.api.dash_rows("s")
        self.assertEqual([(r["name"], r["kun"], r["ost"]) for r in s], [("ТЕНОВИР", 11, 55)],
                         "Srok: bitta qator, eng yaqin muddat, qoldiq — barcha partiyalar")
        self.assertEqual(round(s[0]["summa"]), 5 * 250000 + 50 * 190000)
        st = self.api.state()
        self.assertEqual((st["srok"], st["old"], st["old_days"]), (1, 1, 90))
        self.assertEqual(self.api.stock_abc("o", "ФИЛ Б")["jami"], 1, "filial tanlansa — faqat o'sha filialning 90+ qatori")
        if not self.bot.cfg.has_section("panel"):
            self.bot.cfg.add_section("panel")
        self.bot.cfg["panel"]["old_days"] = "30"
        self.assertEqual(len(self.api.dash_rows("o")), 2, "chegara config.ini [panel] old_days dan")


class Http(BotCase):
    def setUp(self):
        super().setUp()
        self.bot.koch_build()
        self.api = P.Api(self.bot, self.bot.cfg)
        self.rid = BY.create(self.bot.db, FIRM, "АЛЬФА ФАРМ", "Анвар", [{"name": "ДОРИ 1", "qty": 4}, {"name": "ДОРИ 2", "qty": 2}])
        self.srv = P.PanelServer(("127.0.0.1", 0), P.make_handler(self.api, ""))
        threading.Thread(target=self.srv.serve_forever, daemon=True).start()
        self.base = "http://127.0.0.1:%d" % self.srv.server_address[1]

    def tearDown(self):
        self.srv.shutdown()
        self.srv.server_close()
        super().tearDown()

    def req(self, path, body=None, csrf=True):
        data = None if body is None else json.dumps(body).encode()
        r = urllib.request.Request(self.base + path, data=data)
        if body is not None:
            r.add_header("Content-Type", "application/json")
            if csrf:
                r.add_header("X-Requested-With", "panel")
        try:
            with urllib.request.urlopen(r, timeout=10) as resp:
                return resp.status, resp.headers.get("Content-Type", ""), resp.read()
        except urllib.error.HTTPError as e:
            return e.code, e.headers.get("Content-Type", ""), e.read()

    def j(self, path, body=None):
        return json.loads(self.req(path, body)[2])

    def test_orders_endpoints(self):
        self.assertEqual(self.j("/api/orders?status=kutmoqda")["items"][0]["id"], self.rid)
        self.assertIn("items", self.j("/api/orders?status="), "bo'sh status — hammasi")
        self.assertIn("suppliers", self.j("/api/orders?cover="), "eski yetkazib beruvchilar ro'yxati ishlaydi")
        self.assertEqual(self.j("/api/orders/%d" % self.rid)["req"]["n"], 2)
        self.assertEqual(self.req("/api/orders/decide", {"id": self.rid, "pos": 1, "qty": 3}, csrf=False)[0], 403)
        r = self.j("/api/orders/decide", {"id": self.rid, "pos": 1, "qty": 3})
        self.assertEqual((r["ok"], r["msg"]), (True, "Сақланди"))
        self.assertTrue(self.j("/api/orders/accept_all", {"id": self.rid})["ok"])
        r = self.j("/api/orders/finish", {"id": self.rid})
        self.assertTrue(r["ok"])
        self.assertTrue(r["msg"].startswith("Жавоб юборилди"), r["msg"])
        code, ctype, body = self.req("/api/orders/excel?days=7")
        self.assertEqual(code, 200)
        self.assertIn("spreadsheetml", ctype)
        code, _, body = self.req("/api/orders/excel?days=7&firm=YOQ")
        self.assertEqual(code, 404)

    def test_stock_abc_endpoints(self):
        s = self.j("/api/stock/abc?kind=s")
        self.assertIn("classes", s)
        st, ct, body = self.req("/api/stock/abc/excel?kind=s&cls=")
        self.assertEqual((st, ct), (200, P.XLSX))
        self.assertEqual(self.j("/api/stock/abc?kind=zzz")["kind"], "s")

    def test_koch_abc_endpoints(self):
        s = self.j("/api/koch/abc?period=7")
        self.assertEqual((s["period"], len(s["dates"]), s["has"]), ("7", 7, True))
        code, ctype, _ = self.req("/api/koch/abc/excel?period=bugun")
        self.assertEqual((code, "spreadsheetml" in ctype), (200, True))
        self.assertEqual(self.req("/api/koch/abc/excel?period=kecha")[0], 404)

    def test_kirill_yorliqlar(self):
        k = self.j("/api/koch/today")["kpi"]
        self.assertEqual([k[c]["title"] for c in ("shoshilinch", "assortiment", "oddiy")], ["ШОШИЛИНЧ", "АССОРТИМЕНТ", "ОДДИЙ"])
        s = self.j("/api/settings")
        self.assertEqual(s["weekdays"][4], "Жума")
        self.assertIn("auto_refresh_every", s)
        self.assertTrue(self.j("/api/settings", {"auto_refresh_every": 120})["ok"])
        self.assertEqual(self.j("/api/settings")["auto_refresh_every"], 120)

    def test_state_trend(self):
        snap(self.bot, (TODAY - datetime.timedelta(days=1)).isoformat(), 0)
        s = self.j("/api/state")
        self.assertEqual(len(s["trend"]["srok"]), 2)


class Html(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with open(os.path.join(P.PANEL_DIR, "panel.html"), encoding="utf-8") as f:
            cls.html = f.read()

    def test_yangi_bolimlar(self):
        for s in ("/api/koch/abc?period=", "/api/koch/abc/excel?period=", "/api/orders?status=", "/api/orders/decide",
                  "/api/orders/accept_all", "/api/orders/finish", "/api/orders/excel?days=", "auto_refresh_every",
                  "auto_refresh_from", "auto_refresh_to", "Бажарилганлар", "Бажарилмаганлар", "Буюртма сўровлари",
                  "Автоматик янгилаш"):
            self.assertTrue(s in self.html, s)


if __name__ == "__main__":
    unittest.main()
