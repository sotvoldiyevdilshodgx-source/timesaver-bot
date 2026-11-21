# -*- coding: utf-8 -*-
"""Testlar uchun umumiy fixture'lar (unittest): soxta Telegram, vaqtinchalik bot.db, joriy vaqtni almashtirish,
test xlsx yasovchi. Haqiqiy bot.db, config.ini, Excel va Telegramga TEGMAYDI."""
import configparser
import datetime
import logging
import os
import shutil
import sys
import tempfile
import unittest

BOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BOT_DIR)
import timesaver_bot as B  # noqa: E402
from uzkir import strip_marks  # noqa: E402
from openpyxl import Workbook  # noqa: E402

CH, UCH, EK = "ДОРИХОНА ЧИНОЗ", "ДОРИХОНА УРГУТ", "ДОРИХОНА ЭРКИН"
EK2 = "ДОРИХОНА  ЭРКИН"          # Srok / 30+ kun / Ko'chirish dagi kabi 2 probel bilan
TAG_A, TAG_B = "АЛЬФА 901112233", "БЕТА 977778899"
ADMIN, STAFF, FIRM, F_CH, F_EK, F_UCH, F_CH2 = 9001, 5001, 5002, 6001, 6002, 6004, 6003
TODAY = datetime.date.today()
_REAL_NOW = B.now
logging.disable(logging.WARNING)  # kutilgan ogohlantirishlar (toifasiz qator va h.k.) test chiqishini to'ldirmasin


def day(n):
    """Bugundan n kun keyingi sana (datetime, Excel uchun)."""
    return datetime.datetime.combine(TODAY + datetime.timedelta(days=n), datetime.time())


FAR = datetime.datetime(2300, 1, 1)
SH, AS, OD = "Tugagan, sotuv bor", "Filialda yo'q, boshqa filialda sotilyapti", "%d kun turibdi, oluvchida yetmaydi"


def koch_rows():
    """Ko'chirish varag'i (yangi format, 14 ustun). 3 filial, uchala toifa, turli muddatlar. 27 qator."""
    rows = []
    # 🔴 SHOSHILINCH: ЧИНОЗ dan (srok 3 / 45 / 90 / o'tgan / 2300 / yo'q), ЭРКИН (2 probel) dan 2 ta
    for i, (dst, d, q, price) in enumerate([(EK, day(3), 5, 2500), (UCH, day(45), 2, 5400), (EK, day(90), 1, 30000),
                                            (EK, day(-2), 3, 1500), (UCH, FAR, 2, 4000), (EK, None, 1, 8000),
                                            (UCH, day(10), 4, 2000), (EK, day(200), 2, 9000)]):
        rows.append((CH, dst, "ШОШ ДОРИ %02d" % i, "РОССИЯ", q, 2 + i % 3, 0, price, q * price, d, "SHOSHILINCH", SH, 0, q + 3))
    rows.append((EK2, CH, "АВИТЕН", "РОССИЯ", 3, 4, 0, 20000, 60000, day(60), "SHOSHILINCH", SH, 0, 9))
    rows.append((EK2, UCH, "ШОШ ЭКС 01", "ИНДИЯ", 2, 1, 0, 7000, 14000, day(5), "SHOSHILINCH", SH, 0, 4))
    # 🟡 ASSORTIMENT: ЧИНОЗ dan (srok 3 / 45 / 59 / 61 / 90 / o'tgan / 2300), beruvchidagi sotuv har xil
    for i, (dst, d, ss) in enumerate([(EK, day(3), 5), (EK, day(45), 1), (UCH, day(59), 2), (UCH, day(61), 7),
                                       (EK, day(90), 3), (EK, day(-5), 9), (UCH, FAR, 4), (EK, day(300), 6)]):
        rows.append((CH, dst, "АССОРТ %02d" % i, "ГЕРМАНИЯ", 1, 0, 0, 10000 + i, 10000 + i, d, "ASSORTIMENT", AS, ss, 5))
    # 🟢 ODDIY: ЭРКИН (2 probel) dan ЧИНОЗ ga, turgan kun har xil; bittasi ЧИНОЗ dan
    for i, (days, d, q) in enumerate([(45, day(100), 7), (95, day(25), 5), (120, day(-1), 6), (60, FAR, 4),
                                       (33, day(400), 9), (200, day(70), 3)]):
        rows.append((EK2, CH, "ОДДИЙ %02d" % i, "ВЕНГРИЯ", q, 3, 1, 9000, q * 9000, d, "ODDIY", OD % days, 0, q + 2))
    rows.append((CH, UCH, "ОДДИЙ ЧАР", "ВЕНГРИЯ", 8, 5, 1, 7000, 56000, day(150), "ODDIY", OD % 50, 0, 10))
    rows.append((CH, EK, "НОМАЪЛУМ", "X", 1, 1, 0, 100, 100, day(20), "BOSHQA", "nimadir", 0, 1))   # toifasiz
    rows.append((CH, EK, "КАСР", "X", 0.4, 1, 0, 100, 40, day(20), "SHOSHILINCH", SH, 0, 1))       # E < 1
    return rows


def make_xlsx(path, koch="new", rows=None):
    """koch: 'new' (SHOSHILINCH/ASSORTIMENT/ODDIY), 'zudlik' (avvalgi makros: Zudlik/Oddiy), 'old' (10 ustun), None."""
    rows = rows if rows is not None else koch_rows()
    wb = Workbook()
    ws = wb.active
    ws.title = "Лист1"
    ws["C4"], ws["C5"] = datetime.datetime.combine(TODAY - datetime.timedelta(days=10), datetime.time()), day(0)
    ws["E6"], ws["G6"], ws["I6"], ws["K6"] = CH, UCH, EK, "СКЛАД"
    names = sorted({(r[2], r[3]) for r in rows})
    for i, (n, p) in enumerate(names, 8):
        ws.cell(i, 1, n); ws.cell(i, 2, p); ws.cell(i, 4, TAG_A if i % 2 else TAG_B)
        for j in range(5, 11):
            ws.cell(i, j, 3)
    ws.cell(8, 6, -1)   # manfiy sotuv (qaytarilgan tovar)
    srok = wb.create_sheet("Srok")
    srok["B2"] = datetime.datetime.now()
    for i, r in enumerate([
        (CH, "ШОШ ДОРИ 00", "РОССИЯ", TAG_A, day(-5), -5, 4, 2500, 3000, 10000),     # muddati o'tgan
        (EK2, "АВИТЕН", "РОССИЯ", TAG_B, day(-1), -1, 2, 20000, 25000, 40000),        # muddati o'tgan
        (CH, "АССОРТ 00", "ГЕРМАНИЯ", TAG_B, datetime.datetime(TODAY.year + 15, 1, 1), 5000, 1, 9000, 9900, 9000),
        (UCH, "ШОШ ДОРИ 01", "РОССИЯ", TAG_A, day(200), 200, 5, 5400, 6000, 27000)], 5):
        for j, v in enumerate(r, 1):
            srok.cell(i, j, v)
    wb.create_sheet("30+ kun").append([])
    narx = wb.create_sheet("Narx")
    for i, r in enumerate([("ШОШ ДОРИ 00", "РОССИЯ", TAG_A, 2500, 2400, None), ("ШОШ ДОРИ 01", "РОССИЯ", TAG_A, 5400, 5400, None),
                           ("АССОРТ 00", "ГЕРМАНИЯ", TAG_B, 9000, 9900, None)], 5):
        for j, v in enumerate(r, 1):
            narx.cell(i, j, v)
    hisob = wb.create_sheet("Hisob")
    for i, r in enumerate([("ООО АЛЬФА 901112233", 0, 500000, -500000), ("ДОРИХОНА ЧИНОЗ", 900000, 0, 900000),
                           ("ИНВЕНТАРИЗАЦИЯ", 300000, 0, 300000), ("КУП ЧИККАН ТОВАР", 0, 70000, -70000),
                           ("ООО ГАММА", 120000, 0, 120000)], 5):
        for j, v in enumerate(r, 1):
            hisob.cell(i, j, v)
    if koch:
        k = wb.create_sheet("Ko'chirish")
        n = 10 if koch == "old" else 14
        k.append(["Qayerdan", "Qayerga", "Dori", "Ishlab chiqaruvchi", "Ko'chirish soni", "Oluvchida sotuv", "Oluvchida qoldiq",
                  "Kirim narxi", "Summa (kirim)", "Eng yaqin muddat", "Turi", "Sabab", "Beruvchida sotuv", "Beruvchida qoldiq"][:n])
        for r in rows:
            r = list(r)
            if koch == "zudlik":
                r[10] = {"SHOSHILINCH": "Zudlik", "ASSORTIMENT": "Zudlik", "ODDIY": "Oddiy"}.get(r[10], r[10])
            k.append(r[:n])
    wb.save(path)
    return path


class FakeTG:
    """Telegram o'rniga: yuborilganlarni yozib oladi."""

    def __init__(self):
        self.sent, self.calls, self.mid, self.deleted = [], [], 100, []
        self.mids = []  # (cid, message_id) — yuborilgan xabarlar tartibida (send ↔ message_id)

    def send(self, cid, text, markup=None):
        self.mid += 1
        self.sent.append(("send", cid, strip_marks(text), markup))
        self.mids.append((cid, self.mid))
        return {"message_id": self.mid}

    def edit(self, cid, mid, text, markup=None):
        self.sent.append(("edit", cid, strip_marks(text), markup))
        self.calls.append(("editMessageText", {"chat_id": cid, "message_id": mid}))
        return {}

    def delete(self, cid, mid):
        self.deleted.append((cid, mid))
        return True

    def send_document(self, cid, fname, content, caption=None):
        self.sent.append(("doc", cid, fname, content, strip_marks(caption)))
        return {}

    def call(self, method, params=None, timeout=60):
        self.calls.append((method, params))
        return True

    def set_token(self, t):
        pass

    def to(self, cid, kind=None):
        return [s for s in self.sent if s[1] == cid and (kind is None or s[0] == kind)]

    def texts(self, cid):
        return [s[2] for s in self.sent if s[1] == cid and s[0] in ("send", "edit")]

    def last(self, cid, contains=""):
        m = [s for s in self.sent if s[1] == cid and s[0] in ("send", "edit") and contains in s[2]]
        return m[-1] if m else None


class Clock:
    """Joriy vaqtni almashtirish: B.now() shu soatni qaytaradi."""

    def __init__(self, h=8, m=0, d=None):
        self.t = datetime.datetime.combine(d or TODAY, datetime.time(h, m))

    def at(self, h, m=0, d=None):
        self.t = datetime.datetime.combine(d or self.t.date(), datetime.time(h, m))
        return self

    def __call__(self):
        return self.t


def make_bot(koch="new", rows=None, extra_cfg=None):
    work = tempfile.mkdtemp(prefix="abu_test_")
    xlsx = make_xlsx(os.path.join(work, "baza.xlsx"), koch, rows)
    cfg = configparser.ConfigParser(interpolation=None)
    d = {"bot": {"token": "1:x", "admin_ids": str(ADMIN)},
         "excel": {"file": xlsx, "sheet_sale": "", "sheet_srok": "Srok", "sheet_old": "30+ kun", "sheet_narx": "Narx",
                   "sheet_hisob": "Hisob", "sale_first_row": "8"},
         "auto": {"enabled": "no", "time": "09:00", "stale_hours": "72"}}
    for sec, vals in (extra_cfg or {}).items():
        d.setdefault(sec, {}).update(vals)
    cfg.read_dict(d)
    tg = FakeTG()
    bot = B.Bot(cfg, tg=tg, db_path=os.path.join(work, "bot.db"))
    bot.sleep = lambda s: None
    bot.data.reload()
    return bot, tg, work


def add_client(bot, cid, role="firm", status="active", firms=(), filial=None, name=None):
    with bot.dblock:
        bot.db.execute("INSERT INTO clients(chat_id,phone,name,status,role,created) VALUES(?,?,?,?,?,?)",
                       (cid, "+99890%07d" % cid, name or "U%d" % cid, status, role, "2026-10-28T10:00:00"))
        bot.db.commit()
    if firms:
        bot.set_firms(cid, list(firms))
    if filial is not None:
        with bot.dblock:
            bot.db.execute("UPDATE clients SET filial=? WHERE chat_id=?", (B.norm(filial), cid))
            bot.db.commit()


def msg(cid, text, **kw):
    m = {"chat": {"id": cid, "type": "private"}, "from": {"id": cid, "first_name": "T"}, "text": text}
    m.update(kw)
    return {"message": m}


def cb(cid, data, mid=1):
    """Tugma bosildi. mid — tugma turgan xabar (standart 1 — eski xabar; panel xabarida: bot.panel_mid(cid))."""
    return {"callback_query": {"id": "1", "from": {"id": cid}, "data": data, "message": {"message_id": mid}}}


def kb_buttons(markup):
    """Reply klaviatura tugmalari (dict ko'rinishida)."""
    return [b if isinstance(b, dict) else {"text": b} for row in (markup or {}).get("keyboard", []) for b in row]


def kb_texts(markup):
    return [b["text"] for b in kb_buttons(markup)]


def inline_data(markup):
    return [b["callback_data"] for row in (markup or {}).get("inline_keyboard", []) for b in row]


TEXT_FUNCS = ("filial_panel", "koch_list_text", "incoming_list", "client_panel", "section_text", "day_text", "filial_view",
              "hisob_text", "koch_card", "koch_in_card", "admin_panel", "sklad_admin_text", "koch_task_text", "task_card")


def unmark_texts(bot):
    """Testda bevosita chaqirilgan matn funksiyalari: esc() qo'ygan ko'rinmas ma'lumot belgilari olib tashlanadi
    (haqiqiy botda buni Telegram qatlami — TG.call qiladi)."""
    def wrap(fn):
        def inner(*a, **kw):
            r = fn(*a, **kw)
            if isinstance(r, str):
                return strip_marks(r)
            if isinstance(r, tuple) and r and isinstance(r[0], str):
                return (strip_marks(r[0]),) + r[1:]
            return r
        return inner
    for name in TEXT_FUNCS:
        if hasattr(bot, name):
            setattr(bot, name, wrap(getattr(bot, name)))


class BotCase(unittest.TestCase):
    """Har testda yangi bot (vaqtinchalik baza) va soxta vaqt. Standart: uchta filial xodimi, staff, firma."""
    koch = "new"
    clients = True

    def setUp(self):
        self.clock = Clock(8, 0)
        B.now = self.clock
        self.bot, self.tg, self.work = make_bot(self.koch)
        unmark_texts(self.bot)
        if self.clients:
            add_client(self.bot, STAFF, "staff")
            add_client(self.bot, FIRM, "firm", firms=[TAG_A])
            add_client(self.bot, F_CH, "filial", filial=CH, name="Xodim Chinoz")
            add_client(self.bot, F_EK, "filial", filial=EK, name="Xodim Erkin")
            add_client(self.bot, F_UCH, "filial", filial=UCH, name="Xodim Urgut")

    def tearDown(self):
        B.now = _REAL_NOW
        close_bot(self.bot)
        shutil.rmtree(self.work, ignore_errors=True)


def close_bot(bot):
    """Bot ochgan sqlite ulanishini yopadi (ResourceWarning bo'lmasin)."""
    try:
        bot.db.close()
    except Exception:  # noqa
        pass
