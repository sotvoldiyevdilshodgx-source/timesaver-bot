# -*- coding: utf-8 -*-
"""
TimeSaver — dorixonalar uchun Telegram-bot (server paneli bilan).

- Ma'lumot manbai: FAQAT Excel fayl (baza_yangi.xlsm). Bot SQL Serverga ulanmaydi.
- Mijoz (firma) tasdiqlangach, o'z firmasining tovarlari bo'yicha: srok, 30+ kun, buyurtma, hisob, qoldiq/sotuv.
- Admin paneli: hisobotlar, foydalanuvchilar, ommaviy yuborish, sozlamalar va "Yangilash" (Excel makrosini ishga tushiradi).
- Srogi N kundan kam tovarlar har kuni belgilangan vaqtda firmalarga avtomatik yuboriladi.

Faqat standart kutubxona + openpyxl.
"""
import configparser
import datetime
import functools
import glob
import hashlib
import html
import io
import json
import logging
import math
import logging.handlers
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid

import abc_tahlil as ABC_T
import avto_aniqlash as AA
import buyurtma as BY
import uzkir as UK
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

FROZEN = bool(getattr(sys, "frozen", False))
# APP_DIR — dastur fayllari (panel.html, refresh.vbs, logo); o'rnatilgan exe da faqat o'qish uchun.
APP_DIR = getattr(sys, "_MEIPASS", None) or os.path.dirname(os.path.abspath(__file__))


def _data_dir():
    """BASE_DIR — o'zgaradigan fayllar: config.ini, bot.db, bot.log, zaxiralar.
    TIMESAVER_HOME > o'rnatilgan exe: %ProgramData%\\TimeSaver > manba kod papkasi (eski usul)."""
    d = os.environ.get("TIMESAVER_HOME", "").strip()
    if not d and FROZEN:
        d = os.path.join(os.environ.get("PROGRAMDATA") or r"C:\ProgramData", "TimeSaver")
    d = d or os.path.dirname(os.path.abspath(__file__))
    os.makedirs(d, exist_ok=True)
    return d


BASE_DIR = _data_dir()
BRAND = "TimeSaver"
VERSION = "1.0.0"
DB_INI = r"C:\ProgramData\BazaReport\db.ini"   # Excel makrosi va fom_sql shu fayldan SQL ulanishini o'qiydi
SHEET_ORDER = ["Sale", "Srok", "30+ kun", "Narx", "Hisob"]


# ----------------------------------------------------------------------------
# Yordamchi funksiyalar
# ----------------------------------------------------------------------------
def load_config(path=None):
    cp = configparser.ConfigParser(interpolation=None)
    path = path or os.path.join(BASE_DIR, "config.ini")
    with open(path, encoding="utf-8-sig") as f:
        cp.read_file(f)
    return cp


def write_ini(section, key, value, path=None):
    """config.ini ichida bitta qiymatni o'zgartiradi (izohlar va boshqa qatorlar saqlanadi)."""
    path = path or os.path.join(BASE_DIR, "config.ini")
    try:
        with open(path, encoding="utf-8-sig") as f:
            lines = f.read().splitlines()
    except FileNotFoundError:
        lines = []
    sec_re = re.compile(r"^\s*\[(.+?)\]\s*$")
    key_re = re.compile(r"^\s*%s\s*[=:]" % re.escape(key), re.I)
    start = end = None
    for i, ln in enumerate(lines):
        m = sec_re.match(ln)
        if m:
            if start is not None and end is None:
                end = i
            if m.group(1).strip().lower() == section.lower():
                start = i
    if start is not None and end is None:
        end = len(lines)
    new = "%s = %s" % (key, value)
    if start is None:
        lines += ["", "[%s]" % section, new]
    else:
        for i in range(start + 1, end):
            if key_re.match(lines[i]):
                lines[i] = new
                break
        else:
            pos = end
            while pos > start + 1 and not lines[pos - 1].strip():
                pos -= 1
            lines.insert(pos, new)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8", newline="") as f:
        f.write("\r\n".join(lines) + "\r\n")
    os.replace(tmp, path)


_APOS = str.maketrans({"ʼ": "'", "’": "'", "‘": "'", "`": "'"})


_WS = re.compile(r"\s+")


@functools.lru_cache(maxsize=262144)
def _norm(s):
    return _WS.sub(" ", s).strip().upper().translate(_APOS)


def norm(s):
    """Solishtirish uchun: ortiqcha probellar birlashadi, katta harf, apostroflar (ʼ ’ ‘ `) bittasiga.
    Natija keshlanadi: bir xil nomlar (dori, filial, firma) har ekranda yuz minglab marta solishtiriladi."""
    return _norm(s if isinstance(s, str) else str(s if s is not None else ""))


# Ko'chirish toifalari — yagona manba (tartib = ko'rsatish va saralash tartibi)
KOCH_CATS = [
    {"key": "shoshilinch", "turi": "SHOSHILINCH", "title": "SHOSHILINCH", "icon": "🔴", "style": "danger",
     "hint": "Tugagan, sotuv bor — boshqa filialga zudlik bilan berish kerak", "xls_fill": "E5484D"},
    {"key": "assortiment", "turi": "ASSORTIMENT", "title": "ASSORTIMENT", "icon": "🟡", "style": None,
     "hint": "Filialda yo'q, boshqa filialda sotilyapti — namuna sifatida bering", "xls_fill": "F5C33B"},
    {"key": "oddiy", "turi": "ODDIY", "title": "ODDIY", "icon": "🟢", "style": "success",
     "hint": "Sizda uzoq turibdi, boshqa filialda yetmayapti — o'sha yerga bering", "xls_fill": "30A46C"},
]
KOCH_CAT = {c["key"]: c for c in KOCH_CATS}
KOCH_CAT_ORDER = {c["key"]: i for i, c in enumerate(KOCH_CATS)}
_IDLE_RE = re.compile(r"^\s*(\d+)\s*kun\s+turibdi", re.I)


def koch_category_of(turi, sabab):
    """(toifa kaliti yoki None, days_idle). Yagona funksiya: tugmadagi son ham, ro'yxat ham shundan.
    1) K ∈ {SHOSHILINCH, ASSORTIMENT, ODDIY}; 2) L matni; 3) orqaga moslik (K = ZUDLIK / ODDIY / bo'sh)."""
    k, s = norm(turi), norm(sabab)
    m = _IDLE_RE.match(str(sabab or ""))
    idle = int(m.group(1)) if m else 0
    for c in KOCH_CATS:
        if k == c["turi"]:
            return c["key"], idle
    if s.startswith("TUGAGAN"):
        return "shoshilinch", 0
    if s.startswith("FILIALDA YO'Q"):
        return "assortiment", 0
    if m:
        return "oddiy", idle
    if k == "ZUDLIK":
        return "shoshilinch", 0
    if k in ("ODDIY", ""):
        return "oddiy", idle
    return None, 0


def digits(s):
    return re.sub(r"\D", "", str(s or ""))


_FIL_PREFIX = ("ДОРИХОНА ", "DORIXONA ", "DORIXONA ")


@functools.lru_cache(maxsize=4096)
def _fil_key(t):
    t = t.replace("Ё", "Е")
    for p in _FIL_PREFIX:
        if t.startswith(p) and len(t) > len(p):
            return t[len(p):].strip()
    return t


def fil_key(s):
    """Filial nomini solishtirish kaliti: norm + Ё=Е + umumiy old so'z («ДОРИХОНА») olinadi. Shunday qilib Srok / 30+ kun
    varag'idagi «ДОРИХОНА  ЧИНОЗ» (2 probel), Ko'chirish varag'idagi «ДОРИХОНА ЧИНОЗ» va qisqa «ЧИНОЗ» bitta filial.
    Qism bo'yicha (substring) solishtirilmaydi: boshqa filialning ma'lumoti hech qachon aralashmaydi."""
    return _fil_key(norm(s))


CYR_LAT = dict(zip("абвгдеёжзийклмнопрстуфхцчшщъыьэюяўқғҳ",
                   ["a", "b", "v", "g", "d", "e", "e", "j", "z", "i", "y", "k", "l", "m", "n", "o", "p", "r", "s", "t",
                    "u", "f", "x", "ts", "ch", "sh", "sh", "", "i", "", "e", "yu", "ya", "o", "k", "g", "x"]))


@functools.lru_cache(maxsize=200000)
def canon(s):
    """Qidiruv uchun yagona ko'rinish: kirill va lotin bir xil bo'ladi («paracetamol» = «ПАРАЦЕТАМОЛ», «no-shpa» = «НО-ШПА»)."""
    t = "".join(CYR_LAT.get(ch, ch) for ch in str(s or "").lower())
    t = t.replace("o'", "o").replace("g'", "g").replace("ʻ", "").replace("‘", "").replace("'", "")
    for a, b in (("shch", "1"), ("sh", "1"), ("ch", "2"), ("zh", "j"), ("kh", "x"), ("ph", "f"), ("ts", "s")):
        t = t.replace(a, b)
    t = re.sub(r"c(?=[eiy])", "s", t)
    t = t.replace("ck", "k").replace("c", "k").replace("h", "x").replace("w", "v").replace("q", "k")
    t = t.replace("1", "sh").replace("2", "ch")
    t = re.sub(r"[^a-z0-9]+", " ", t)
    return re.sub(r"(.)\1+", r"\1", t).strip()


def esc_t(s):
    """Tizim matni uchun HTML-xavfsiz (belgilanmaydi — Telegramga chiqishda kirillga o'tadi): Excel makrosining lotin
    izohlari («sabab»), voqea nomlari va h.k."""
    return html.escape(str(s if s is not None else ""), quote=False)


def esc(s):
    """HTML uchun xavfsiz qiladi va ma'lumot deb belgilaydi: Telegramga chiqishda kirillga o'girilmaydi (uzkir)."""
    return UK.mark(html.escape(str(s if s is not None else ""), quote=False))


def bl(s):
    """Tugma matnidagi ma'lumot (dori, firma, filial, odam nomi): uzkir.markup kirillga o'girmaydi. Tugma matni HTML emas —
    escape qilinmaydi."""
    return UK.mark(str(s if s is not None else ""))


def to_float(v):
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(" ", "").replace("\xa0", "").replace(",", "."))
    except ValueError:
        return 0.0


def to_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    return None


def money(v):
    return "{:,.0f}".format(v).replace(",", " ")


def mln(v, unit=True):
    """Katta summa: «42,6 mln so'm» (1 mln dan kam — «740 000 so'm»). unit=False — «so'm» siz."""
    t = ("%.1f" % (v / 1e6)).replace(".", ",") + " mln" if abs(v) >= 1e6 else money(v)
    return t + (" so'm" if unit else "")


def qty(v):
    return ("%.2f" % v).rstrip("0").rstrip(".") if v != int(v) else str(int(v))


def newest_file(pattern):
    if pattern and os.path.isfile(pattern):  # aniq yo'l: nomida [ ] bo'lsa ham glob uni naqsh deb o'qimasin
        return pattern
    files = [f for f in glob.glob(pattern) if not os.path.basename(f).startswith("~$")]

    def mt(f):  # Excel saqlayotganda fayl bir lahza yo'qolishi mumkin
        try:
            return os.path.getmtime(f)
        except OSError:
            return -1
    return max(files, key=mt) if files else None


def now():
    return datetime.datetime.now()


def dt_text(ts):
    return ts.strftime("%d.%m.%Y %H:%M") if ts else "-"


# ----------------------------------------------------------------------------
# Excel dan ma'lumot o'qish
# ----------------------------------------------------------------------------
class Data:
    """baza_yangi.xlsm ni o'qiydi. Fayl o'zgarganda qayta yuklaydi. SQL bilan aloqasi yo'q."""

    def __init__(self, cfg):
        self.sec = cfg["excel"]
        self.lock = threading.Lock()
        self.path = None
        self.mtime = None
        self.stamp = None  # oxirgi yangilanish vaqti
        self.error = None
        self.period_days = 30
        self.sale, self.srok, self.old, self.narx, self.hisob = [], [], [], [], []
        self.sale_heads = []  # Sale varag'i 6-qatori: E, G, I, K — filiallar va sklad nomlari (v[0], v[2], v[4], v[6])
        self.koch = []  # Ko'chirish varag'i (filiallararo ko'chirish ro'yxati)
        self.koch_new = False  # varaq yangi makros formatidami (K = Turi)
        self.sklad_name = norm(self.sec.get("sklad_name", "") or "ДОРИХОНА СКЛАД")  # Srok varag'idagi sklad nomi

    def file_pattern(self):
        return self.sec.get("file", "").strip()

    def find_path(self):
        return newest_file(self.file_pattern())

    def _sheet(self, wb, key, default_idx=None):
        want = norm(self.sec.get("sheet_" + key, "")) or norm({"sale": "", "srok": "Srok", "old": "30+ kun",
                                                              "narx": "Narx", "hisob": "Hisob",
                                                              "koch": "Ko'chirish", "sklad": "Skladdan"}[key])
        for ws in wb.worksheets:
            if want and norm(ws.title) == want:
                return ws
        if key == "koch":  # makros nomni ‘ ` ’ kabi boshqa tutuq belgisi bilan yozgan bo'lsa ham topamiz
            bare = lambda s: re.sub(r"[\W_]+", "", norm(s))
            for ws in wb.worksheets:
                if bare(ws.title) == bare(want):
                    return ws
        if default_idx is not None and len(wb.worksheets) > default_idx:
            return wb.worksheets[default_idx]
        return None

    @staticmethod
    def _rows(ws, first, ncols):
        if ws is None:
            return
        for r in ws.iter_rows(min_row=first, max_col=ncols, values_only=True):
            r = list(r) + [None] * (ncols - len(r))
            if r[0] in (None, "") and r[1] in (None, ""):
                continue
            yield r

    def reload(self, force=False):
        with self.lock:
            path = self.find_path()
            if not path:
                self.error = "Excel fayl topilmadi: " + self.file_pattern()
                return False
            try:
                mtime = os.path.getmtime(path)
            except OSError as e:  # Excel saqlayotgan paytda
                self.error = "Excel fayl hozir band: %s" % e
                return False
            if not force and path == self.path and mtime == self.mtime:
                return True
            tmp = os.path.join(tempfile.gettempdir(), "timesaver_cache_%d%s" % (os.getpid(), os.path.splitext(path)[1] or ".xlsx"))
            try:
                shutil.copy2(path, tmp)  # Excel ochiq bo'lsa ham o'qiy olamiz
                wb = load_workbook(tmp, read_only=True, data_only=True)
                try:
                    self._parse(wb)
                finally:
                    wb.close()
                self.path, self.mtime, self.error = path, mtime, None
                logging.info("Excel yuklandi: %s (sale %d, srok %d, 30+ %d, narx %d, hisob %d, ko'chirish %d)", path,
                             len(self.sale), len(self.srok), len(self.old), len(self.narx), len(self.hisob), len(self.koch))
                return True
            except Exception as e:  # noqa
                logging.exception("Excel o'qishda xato")
                self.error = "Excelni o'qib bo'lmadi: %s" % e
                return False
            finally:
                try:
                    os.remove(tmp)
                except OSError:
                    pass

    @staticmethod
    def merge_sale(rows):
        """Bir xil dori (nom + ishlab chiqaruvchi, bo'shliqlar hisobga olinmaydi) Sale varag'ida bir necha qatorda kelsa
        (FOM'da firma bir partiyada yozilgan, boshqasida bo'sh; nomda ortiqcha bo'shliq) — bitta qatorga qo'shiladi.
        Aks holda bot faqat birinchi qatorni olib, qoldiqni kam ko'rsatardi (FOM'da 6, botda 5). Firma — bo'sh bo'lmagani;
        ikki xil (bo'sh bo'lmagan) firma bo'lsa qatorlar alohida qoladi (har firma vakili o'zinikini ko'radi)."""
        out, first = [], {}
        for r in rows:
            k = (norm(r["name"]), norm(r["producer"]))
            m = first.get(k)
            if m is not None and (not m["firm"] or not r["firm"] or norm(m["firm"]) == norm(r["firm"])):
                m["v"] = [a + b for a, b in zip(m["v"], r["v"])]
                m["stock"] += r["stock"]
                m["sales"] += r["sales"]
                m["firm"] = m["firm"] or r["firm"]
                m["mnn"] = m["mnn"] or r["mnn"]
                continue
            r = dict(r, v=list(r["v"]))
            first.setdefault(k, r)
            out.append(r)
        if len(out) != len(rows):
            logging.info("Sale: %d ta takror qator qo'shildi (bir xil dori bir necha qatorda)", len(rows) - len(out))
        return out

    def _parse(self, wb):
        ws = self._sheet(wb, "sale", 0)
        sale = []
        for r in self._rows(ws, int(self.sec.get("sale_first_row", "8")), 12):
            v = [to_float(x) for x in r[4:11]]
            for i in (1, 3, 5):  # manfiy sotuv (qaytarilgan tovar) = 0
                v[i] = max(0.0, v[i])
            sale.append({"name": str(r[0]).strip(), "producer": str(r[1] or "").strip(), "mnn": str(r[2] or "").strip(),
                         "firm": str(r[3] or "").strip(), "v": v,
                         "stock": v[0] + v[2] + v[4] + v[6], "sales": v[1] + v[3] + v[5]})
        sale = self.merge_sale(sale)
        period = 30
        heads = []
        try:
            h = list(ws.iter_rows(min_row=6, max_row=6, min_col=5, max_col=11, values_only=True))
            h = list(h[0]) + [None] * 7 if h else [None] * 7
            heads = [str(h[i] or "").strip() for i in (0, 2, 4, 6)]
        except Exception:  # noqa
            pass
        try:
            c = list(ws.iter_rows(min_row=4, max_row=5, min_col=3, max_col=3, values_only=True))
            a, b = to_date(c[0][0]), to_date(c[1][0])
            if a and b and (b - a).days > 0:
                period = (b - a).days
        except Exception:  # noqa
            pass
        srok = []
        for r in self._rows(self._sheet(wb, "srok"), 5, 10):
            srok.append({"filial": str(r[0] or "").strip(), "name": str(r[1]).strip(), "producer": str(r[2] or "").strip(),
                         "firm": str(r[3] or "").strip(), "date": to_date(r[4]), "kun": to_float(r[5]),
                         "ost": to_float(r[6]), "prih": to_float(r[7])})
        old = []
        for r in self._rows(self._sheet(wb, "old"), 5, 9):
            old.append({"filial": str(r[0] or "").strip(), "name": str(r[1]).strip(), "producer": str(r[2] or "").strip(),
                        "supplier": str(r[3] or "").strip(), "firm": str(r[4] or "").strip(), "kun": to_float(r[5]),
                        "ost": to_float(r[6]), "prih": to_float(r[7]), "summa": to_float(r[8])})
        narx = []
        for r in self._rows(self._sheet(wb, "narx"), 5, 6):
            narx.append({"name": str(r[0]).strip(), "producer": str(r[1] or "").strip(), "firm": str(r[2] or "").strip(),
                         "prih": to_float(r[3]), "roz": to_float(r[4]), "sana": to_date(r[5])})
        hisob = []
        for r in self._rows(self._sheet(wb, "hisob"), 5, 4):
            hisob.append({"org": str(r[0]).strip(), "debit": to_float(r[1]), "credit": to_float(r[2]),
                          "itogo": to_float(r[3])})
        # Ko'chirish varag'i (makros yaratadi, 1-qator sarlavha). Firma Sale varag'idan (tovar + ishlab chiqaruvchi) olinadi.
        # K Turi (SHOSHILINCH / ASSORTIMENT / ODDIY; avvalgi makros: Zudlik / Oddiy), L Sabab, M/N beruvchida sotuv/qoldiq.
        # Toifa — koch_category_of(K, L); eski faylda (K–N yo'q) hammasi ODDIY. Manfiy sotuv = 0.
        firm_of = {}
        for r in sale:
            firm_of.setdefault((norm(r["name"]), norm(r["producer"])), r["firm"])
        koch = []
        kws = self._sheet(wb, "koch")
        # Bo'sh varaq (ko'chirish yo'q kun): iter_rows [] qaytaradi — butun Excel yuklanishini buzmasin
        head = next(iter(kws.iter_rows(min_row=1, max_row=1, max_col=14, values_only=True)), ()) if kws is not None else ()
        self.koch_new = len(head) > 10 and norm(head[10]) == "TURI"  # yangi makros (Turi ustuni bor)
        unknown = 0
        for r in self._rows(kws, 2, 14):
            name, producer = str(r[2] or "").strip(), str(r[3] or "").strip()
            if not name:
                continue
            sabab = str(r[11] or "").strip()
            cat, idle = koch_category_of(r[10], sabab)
            if cat is None:
                unknown += 1
            koch.append({"src": str(r[0] or "").strip(), "dst": str(r[1] or "").strip(), "name": name,
                         "producer": producer, "qty": to_float(r[4]), "dst_sales": max(0.0, to_float(r[5])),
                         "dst_stock": to_float(r[6]), "prih": to_float(r[7]), "summa": to_float(r[8]),
                         "date": to_date(r[9]), "turi": str(r[10] or "").strip(), "sabab": sabab,
                         "src_sales": max(0.0, to_float(r[12])), "src_stock": to_float(r[13]),
                         "firm": firm_of.get((norm(name), norm(producer)), ""), "cat": cat, "days_idle": idle})
        # Skladdan varag'i (makros): sklad → dorixona. Beruvchisi sklad bo'lgan oddiy ko'chirish qatori sifatida qo'shiladi.
        # Makros muddatni tekshirmaydi: muddat Srok varag'idan olinadi, faqat muddati o'tgan qoldiq bo'lsa qator chiqariladi.
        near, gone = {}, {}
        today = now().date()
        for r in srok:
            if norm(r["filial"]) == self.sklad_name and r["ost"] > 0 and r["date"]:
                k = (norm(r["name"]), norm(r["producer"]))
                if r["date"] >= today:
                    near[k] = min(near.get(k, r["date"]), r["date"])
                else:
                    gone[k] = max(gone.get(k, r["date"]), r["date"])
        for r in self._rows(self._sheet(wb, "sklad"), 2, 13):
            dst, name, producer = str(r[0] or "").strip(), str(r[1] or "").strip(), str(r[2] or "").strip()
            if not dst or not name:
                continue
            sabab = str(r[10] or "").strip()
            cat, idle = koch_category_of(r[9], sabab)
            if cat is None:
                unknown += 1
            k = (norm(name), norm(producer))
            koch.append({"src": self.sklad_name, "dst": dst, "name": name, "producer": producer, "qty": to_float(r[8]),
                         "dst_sales": max(0.0, to_float(r[4])), "dst_stock": to_float(r[5]), "prih": to_float(r[11]),
                         "summa": to_float(r[12]), "date": near.get(k) or gone.get(k), "turi": str(r[9] or "").strip(), "sabab": sabab,
                         "src_sales": 0.0, "src_stock": to_float(r[3]), "firm": firm_of.get(k, ""), "cat": cat, "days_idle": idle})
        if unknown:
            logging.warning("Ko'chirish varag'ida toifasi aniqlanmagan %d qator (Turi/Sabab mos emas) — ro'yxatlarga kirmaydi", unknown)
        stamp = None
        for key in ("srok", "old", "narx", "hisob"):
            w = self._sheet(wb, key)
            if w is not None:
                try:
                    v = list(w.iter_rows(min_row=2, max_row=2, min_col=2, max_col=2, values_only=True))[0][0]
                    if isinstance(v, datetime.datetime):
                        stamp = v
                        break
                except Exception:  # noqa
                    pass
        self.sale, self.srok, self.old, self.narx, self.hisob = sale, srok, old, narx, hisob
        self.koch = koch
        self.sale_heads = heads
        self.period_days = period
        self.stamp = stamp

    # ---- so'rovlar ----
    def when(self):
        """Ma'lumot qachon yangilangan (Excel ichidagi vaqt, bo'lmasa fayl vaqti)."""
        if self.stamp:
            return self.stamp
        return datetime.datetime.fromtimestamp(self.mtime) if self.mtime else None

    def age_hours(self):
        w = self.when()
        return (now() - w).total_seconds() / 3600 if w else 1e9

    def all_tags(self):
        tags = set()
        for rows in (self.sale, self.srok, self.old, self.narx):
            tags.update(norm(r["firm"]) for r in rows if r["firm"])
        return sorted(tags)

    @staticmethod
    def _by_tags(rows, tags):
        if tags is None:
            return list(rows)
        return [r for r in rows if norm(r["firm"]) in tags]

    def srok_rows(self, tags, days):
        today = datetime.date.today()
        out = []
        for r in self._by_tags(self.srok, tags):
            kun = (r["date"] - today).days if r["date"] else r["kun"]
            if kun <= days:
                out.append(dict(r, kun=kun))
        out.sort(key=lambda r: (r["kun"], norm(r["name"])))
        return out

    def old_rows(self, tags):
        out = self._by_tags(self.old, tags)
        return sorted(out, key=lambda r: (-r["kun"], norm(r["name"])))

    def stock_rows(self, tags):
        return sorted(self._by_tags(self.sale, tags), key=lambda r: norm(r["name"]))

    def need_rows(self, tags, cover):
        out = []
        for r in self._by_tags(self.sale, tags):
            need = r["sales"] * cover - r["stock"]
            if r["sales"] > 0 and need > 0:
                out.append(dict(r, need=need))
        return sorted(out, key=lambda r: norm(r["name"]))

    def koch_rows(self, tags):
        return sorted(self._by_tags(self.koch, tags), key=lambda r: (-r["summa"], norm(r["name"])))

    def koch_filials(self):
        """Ko'chirish varag'idagi (A ∪ B) takrorlanmas filiallar, asl yozilishi bilan (norm bo'yicha bittadan)."""
        seen = {}
        for r in self.koch:
            for x in (r["src"], r["dst"]):
                if norm(x):
                    seen.setdefault(norm(x), re.sub(r"\s+", " ", x).strip())
        return [seen[k] for k in sorted(seen)]

    def koch_out(self, filial, cat=None):
        """Shu filialdan (A ustun) chiqadigan qatorlar. cat: 'shoshilinch' | 'assortiment' | 'oddiy' | None (hammasi)."""
        f = norm(filial)
        return [r for r in self.koch if r["cat"] and norm(r["src"]) == f and (cat is None or r["cat"] == cat)]

    def koch_in(self, filial, cat=None):
        """Shu filialga (B ustun) keladigan qatorlar."""
        f = norm(filial)
        return [r for r in self.koch if r["cat"] and norm(r["dst"]) == f and (cat is None or r["cat"] == cat)]

    koch_for, koch_incoming = koch_out, koch_in  # eski nomlar (orqaga moslik)

    @staticmethod
    def koch_date_state(date, warn_days=30, today=None):
        """'ok' | 'warn' (qolgan kun < warn_days) | 'expired' | 'unknown' (sana yo'q yoki yili bugun + 10 dan katta)."""
        today = today or now().date()
        if not date or date.year > today.year + 10:
            return "unknown"
        left = (date - today).days
        return "expired" if left < 0 else ("warn" if left < warn_days else "ok")

    def account_rows(self, names):
        names = {norm(n) for n in names}
        return [r for r in self.hisob if norm(r["org"]) in names]

    def hisob_names(self):
        return [r["org"] for r in self.hisob]


# ----------------------------------------------------------------------------
# Excel hisobotlarni tayyorlash
# ----------------------------------------------------------------------------
HEAD_FILL = PatternFill("solid", fgColor="2F7D4F")
HEAD_FONT = Font(bold=True, color="FFFFFF")


def xlsx_bytes(title, headers, rows, widths, datecols=(), numcols=()):
    wb = Workbook()
    ws = wb.active
    ws.title = UK.lat2cyr(title)[:31]
    ws.append([UK.lat2cyr(h) if isinstance(h, str) else h for h in headers])
    for r in rows:
        ws.append(r)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for c in ws[1]:
        c.fill, c.font = HEAD_FILL, HEAD_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=2):
        for i in datecols:
            row[i].number_format = "DD.MM.YYYY"
        for i in numcols:
            row[i].number_format = "#,##0.##"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def build_srok(rows):
    data = [[r["filial"], r["name"], r["producer"], r["firm"], r["date"], r["kun"], r["ost"]] for r in rows]
    return xlsx_bytes("Srok", ["Filial", "Tovar", "Ishlab chiqaruvchi", "Firma tel", "Srok", "Qolgan kun", "Qoldiq"], data,
                      [22, 46, 26, 14, 13, 11, 11], datecols=(4,), numcols=(5, 6))


def build_old(rows):
    data = [[r["filial"], r["name"], r["producer"], r["kun"], r["ost"]] for r in rows]
    return xlsx_bytes("30+ kun", ["Filial", "Tovar", "Ishlab chiqaruvchi", "Necha kun", "Qoldiq"], data,
                      [22, 46, 26, 11, 11], numcols=(3, 4))


def build_old_admin(rows):
    data = [[r["filial"], r["name"], r["producer"], r["supplier"], r["firm"], r["kun"], r["ost"], r["summa"]]
            for r in rows]
    return xlsx_bytes("30+ kun", ["Filial", "Tovar", "Ishlab chiqaruvchi", "Ta'minotchi", "Firma", "Necha kun",
                                  "Qoldiq", "Summa (kirim)"], data, [22, 46, 26, 28, 24, 10, 10, 15],
                      numcols=(5, 6, 7))


def build_abc(kind, rows, cls_of, value, prices=True, title=None):
    """Filial xodimi: ⏰ Srok ('s') / 📦 30+ kun ('o') — A / B / C ustuni bilan, berilgan tartibda (pul bo'yicha). Firma
    ma'lumoti yo'q (filial xodimi firma ko'rmaydi); prices=False — narx va summasiz (sklad xodimi)."""
    if kind == "s":
        heads = ["ABC", "Dori", "Ishlab chiqaruvchi", "Srok", "Qolgan kun", "Qoldiq"]
        data = [[cls_of(r), r["name"], r["producer"], r["date"], r["kun"], r["ost"]] for r in rows]
        widths, datecols = [6, 46, 26, 12, 10, 10], (3,)
    else:
        heads = ["ABC", "Dori", "Ishlab chiqaruvchi", "Turgan kun", "Qoldiq"]
        data = [[cls_of(r), r["name"], r["producer"], r["kun"], r["ost"]] for r in rows]
        widths, datecols = [6, 46, 26, 10, 10], ()
    if prices:
        heads += ["Kirim narxi", "Summa"]
        widths += [12, 14]
        for d, r in zip(data, rows):
            d += [r["prih"], value(r)]
    n = len(heads)
    return xlsx_bytes(title or ("Srok" if kind == "s" else "30+ kun"), heads, data, widths, datecols=datecols,
                      numcols=tuple(i for i in range(3, n) if i not in datecols))


def build_stock(rows, days):
    data = [[r["name"], r["producer"], r["v"][0], r["v"][1], r["v"][2], r["v"][3], r["v"][4], r["v"][5], r["v"][6],
             r["stock"], r["sales"]] for r in rows]
    heads = ["Tovar", "Ishlab chiqaruvchi", "Chinoz qoldiq", "Chinoz sotuv", "Urgut qoldiq", "Urgut sotuv",
             "Erkin qoldiq", "Erkin sotuv", "Sklad qoldiq", "Jami qoldiq", "Jami sotuv (%d kun)" % days]
    return xlsx_bytes("Qoldiq va sotuv", heads, data, [46, 26] + [11] * 9, numcols=tuple(range(2, 11)))


def build_need(rows, days):
    data = [[r["name"], r["producer"], r["v"][0], r["v"][2], r["v"][4], r["v"][6], r["stock"], r["sales"], r["need"]]
            for r in rows]
    heads = ["Tovar", "Ishlab chiqaruvchi", "Chinoz", "Urgut", "Erkin", "Sklad", "Jami qoldiq",
             "Sotuv (%d kun)" % days, "Buyurtma kerak"]
    return xlsx_bytes("Buyurtma", heads, data, [46, 26, 10, 11, 11, 10, 12, 14, 15], numcols=tuple(range(2, 9)))


def build_koch(rows, cls_of=None):
    """cls_of(qator) → 'A' | 'B' | 'C' | "" (ko'chirish A / B / C ustuni; None — ustunsiz, eski ko'rinish)."""
    data = [[r["src"], r["dst"], r["name"], r["producer"], r["firm"], r["qty"], r["dst_sales"], r["dst_stock"], r["prih"],
             r["summa"], r["date"]] for r in rows]
    data.append(["JAMI", "", "", "", "", sum(r["qty"] for r in rows), None, None, None, sum(r["summa"] for r in rows), None])
    heads = ["Qayerdan", "Qayerga", "Tovar", "Ishlab chiqaruvchi", "Firma", "Ko'chirish soni", "Oluvchida sotuv", "Oluvchida qoldiq",
             "Kirim narxi", "Summa (kirim)", "Eng yaqin srok"]
    widths, dcols, ncols = [20, 20, 46, 26, 24, 11, 11, 11, 13, 15, 13], (10,), (5, 6, 7, 8, 9)
    if cls_of:
        data = [d + [cls_of(r)] for r, d in zip(rows, data)] + [data[-1] + [""]]
        heads, widths = heads + ["ABC"], widths + [6]
    return xlsx_bytes("Ko'chirish", heads, data, widths, datecols=dcols, numcols=ncols)


def build_koch_filial(rows, filial):
    """Filial xodimi uchun Ko'chirish Exceli (bitta tur): varaq nomi = filial."""
    data = [[r["dst"], r["name"], r["producer"], r["qty"], r["dst_sales"], r["dst_stock"], r["prih"], r["summa"], r["date"],
             r.get("sabab", "")] for r in rows]
    data.append(["JAMI", "", "", sum(r["qty"] for r in rows), None, None, None, sum(r["summa"] for r in rows), None, ""])
    title = re.sub(r"[\\/?*\[\]:]", " ", filial).strip()[:31] or "Ko'chirish"
    return xlsx_bytes(title, ["Qayerga", "Dori", "Ishlab chiqaruvchi", "Soni", "Oluvchida sotuv", "Oluvchida qoldiq",
                              "Kirim narxi", "Summa", "Eng yaqin srok", "Sabab"], data,
                      [22, 46, 26, 9, 11, 11, 13, 15, 13, 40], datecols=(8,), numcols=(3, 4, 5, 6, 7))


FILL_WARN = PatternFill("solid", fgColor="FFF2B3")
FILL_EXPIRED = PatternFill("solid", fgColor="F4B6B6")
FILL_IDLE = PatternFill("solid", fgColor="FBE0DC")
KOCH_ST_XLS = {"yangi": "yangi", "berildi": "berildi", "olindi": "olindi", "rad": "rad", "bekor": "bekor", "yopildi": "yopildi"}


def build_koch_cat_xlsx(rows, excluded, cat, filial, warn_days=30, with_status=False, today=None, prices=True, cls_of=None):
    """B4.5: filial xodimi Exceli — bitta toifa. 1-varaq (nomi — filial): ro'yxat + JAMI; 2-varaq «Ko'chirilmaydi»:
    muddat sababli chiqarilganlar. Sarlavha foni — toifa rangi; qatorlar: < 30 kun — sariq, o'tgan — qizil,
    ODDIY'da turgan kun ≥ 90 — och qizil. prices=False: kirim narxi va summa ustunlarisiz (sklad xodimi uchun).
    cls_of(qator) → 'A' | 'B' | 'C': ro'yxat varag'ida oxirgi ustun «ABC» (None — ustunsiz)."""
    k = KOCH_CAT[cat]
    today = today or now().date()
    idle = cat == "oddiy"
    heads = (["Qayerga", "Dori", "Ishlab chiqaruvchi", "Ko'chirish soni"] + (["Turgan kun"] if idle else []) +
             ["Oluvchida sotuv", "Oluvchida qoldiq", "Beruvchida sotuv", "Beruvchida qoldiq"] + (["Kirim narxi", "Summa"] if prices else []) +
             ["Eng yaqin muddat", "Qolgan kun", "Sabab"])
    widths = [22, 46, 26, 10] + ([9] if idle else []) + [10, 10, 10, 10] + ([12, 14] if prices else []) + [12, 9, 40]

    def line(r):
        d = r["date"] if Data.koch_date_state(r["date"], warn_days, today) != "unknown" else None
        return ([norm(r["dst"]), r["name"], r["producer"], r["qty"]] + ([r.get("days_idle") or 0] if idle else []) +
                [r["dst_sales"], r["dst_stock"], r["src_sales"], r["src_stock"]] + ([r["prih"], r["summa"]] if prices else []) +
                [d, (d - today).days if d else None, r.get("sabab") or ""])

    def fill(ws, row_no, r):
        st = Data.koch_date_state(r["date"], warn_days, today)
        f = FILL_EXPIRED if st == "expired" else FILL_WARN if st == "warn" else FILL_IDLE if (
            idle and (r.get("days_idle") or 0) >= 90) else None
        if f:
            for c in ws[row_no]:
                c.fill = f

    wb = Workbook()
    hfill = PatternFill("solid", fgColor=k["xls_fill"])
    sheets = [(re.sub(r"[\\/?*\[\]:]", " ", filial).strip()[:31] or k["title"], rows, with_status, None),
              ("Ko'chirilmaydi", [r for r, _ in excluded], False, [why for _, why in excluded])]
    wb.remove(wb.active)
    for title, data, status, why in sheets:
        ws = wb.create_sheet(title)
        abc = cls_of is not None and why is None
        h = heads + (["Holat"] if status else []) + (["Nega chiqarildi"] if why is not None else []) + (["ABC"] if abc else [])
        ws.append(h)
        for i, r in enumerate(data):
            ws.append(line(r) + ([KOCH_ST_XLS.get(r.get("status"), "")] if status else []) + ([why[i]] if why is not None else []) +
                      ([cls_of(r)] if abc else []))
            fill(ws, i + 2, r)
        if why is None:
            tot = ["JAMI", "", "", sum(r["qty"] for r in data)] + ([None] if idle else []) + [None] * 4 + (
                [None, sum(r["summa"] for r in data)] if prices else []) + [None, None, ""] + ([""] if status else []) + (
                [""] if abc else [])
            ws.append(tot)
            for c in ws[ws.max_row]:
                c.font = Font(bold=True)
        for i, w in enumerate(widths + ([12] if status else []) + ([30] if why is not None else []) + ([6] if abc else []), 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        for c in ws[1]:
            c.fill, c.font = hfill, HEAD_FONT
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        dcol = h.index("Eng yaqin muddat")
        for row in ws.iter_rows(min_row=2):
            row[dcol].number_format = "DD.MM.YYYY"
            for j in range(3, dcol):
                row[j].number_format = "#,##0.##"
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(h)), max(1, ws.max_row - (1 if why is None else 0)))
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def xlsx_multi(sheets):
    """Bir nechta varaqli Excel (xlsx_bytes uslubida). sheets: [(nom, sarlavhalar, qatorlar, kengliklar, datecols, numcols[, bezak(ws)])]."""
    wb = Workbook()
    wb.remove(wb.active)
    for title, headers, rows, widths, datecols, numcols, *paint in sheets:
        ws = wb.create_sheet(title[:31])
        ws.append(headers)
        for r in rows:
            ws.append(r)
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        for c in ws[1]:
            c.fill, c.font = HEAD_FILL, HEAD_FONT
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for row in ws.iter_rows(min_row=2):
            for i in datecols:
                row[i].number_format = "DD.MM.YYYY"
            for i in numcols:
                row[i].number_format = "#,##0.##"
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for p in paint:
            p(ws)
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


KOCH_TASK_HEADS = ["Sana", "Toifa", "Qayerdan", "Qayerga", "Dori", "Ishlab chiqaruvchi", "Reja", "Berildi", "Olindi", "Summa",
                   "Srok", "Holat", "Kechikish (daq.)", "Aybdor", "Farq", "Berdi", "Berilgan", "Oldi", "Olingan", "Izoh", "Manba", "Isbot",
                   "ABC"]
KOCH_TASK_WIDTHS = [11, 8, 20, 20, 40, 24, 7, 8, 8, 13, 11, 18, 10, 20, 6, 18, 9, 18, 9, 30, 10, 30, 6]
KOCH_TASK_NUMS = (6, 7, 8, 9)


def koch_task_rows(tasks):
    """koch_task_dict lar → Excel qatorlari."""
    st = KOCH_ST_XLS
    return [[t["sana"], KOCH_CAT.get(t["cat"], KOCH_CAT["oddiy"])["title"], t["src"], t["dst"], t["name"], t["producer"], t["qty_plan"],
             t["qty_given"], t["qty_received"], t["summa"], t["srok"],
             st.get(t["status"], t["status"]) + (" (kechikkan)" if t["late"] else " (kech)" if t["late_done"] else ""),
             t["late_min"] or None, t["guilty"], "⚠" if t["farq"] else "", t["given_by"], t["given_at"], t["received_by"],
             t["received_at"], t["note"], "tizim" if t.get("by") == "tizim" else "", t.get("auto_note") or "", t.get("cls") or ""]
            for t in tasks]


def build_koch_tasks(bot, rows):
    """Ko'chirish topshiriqlari (holati bilan) — filial xodimi Exceli."""
    abc = {}
    tasks = [bot.koch_task_dict(t, cls=abc.setdefault(t["sana"], bot.koch_abc(t["sana"]))) for t in rows]
    return xlsx_bytes("Ko'chirish", KOCH_TASK_HEADS, koch_task_rows(tasks), KOCH_TASK_WIDTHS, numcols=KOCH_TASK_NUMS)


def build_koch_nazorat(bot, tasks, sana):
    """Admin / panel: Kochirish_nazorat_<sana>.xlsx — topshiriqlar (filtr bo'yicha) va filiallar xulosasi."""
    s = bot.koch_summary(sana)
    fil = [[r["filial"], r["give"], r["given"], r["give_left"], r["take"], r["taken"], r["take_left"], r["late"], r["rad"],
            r["pct"], {"red": "kechikkan", "yellow": "jarayonda", "green": "bajarildi"}[r["state"]], r["last"]] for r in s["filials"]]
    return xlsx_multi([
        ("Topshiriqlar", KOCH_TASK_HEADS, koch_task_rows(tasks), KOCH_TASK_WIDTHS, (), KOCH_TASK_NUMS),
        ("Filiallar", ["Filial", "Beradi jami", "Berdi", "Berish qoldi", "Oladi jami", "Oldi", "Olish qoldi", "Kechikkan", "Rad",
                       "Bajarilish %", "Holat", "Oxirgi harakat"], fil, [24, 10, 8, 10, 10, 8, 10, 10, 6, 11, 12, 12], (), ())])


SKLAD_CAT_FILL = {"shoshilinch": PatternFill("solid", fgColor="F8C9C9"), "assortiment": PatternFill("solid", fgColor="FCE8A6"),
                  "oddiy": PatternFill("solid", fgColor="C6E9D1")}


def build_sklad(s, prices=True):
    """Sklad taqsimoti: umumiy jadval (dori × dorixona, toifa matni va ranglari), har dorixonaga yig'ish ro'yxati va berilmaydiganlar.
    s: sklad_summary(). prices=False: summasiz (sklad xodimi uchun)."""
    dates = lambda v: datetime.date.fromisoformat(v) if v else None
    key = lambda c: c if c in KOCH_CAT else "oddiy"
    cats = lambda i: sorted({key(x["cat"]) for x in i["to"].values()}, key=KOCH_CAT_ORDER.get)

    def paint_all(ws):  # «Toifa» — eng shoshilinchi rangida, son katagi — o'sha dorixonadagi toifa rangida
        for n, i in enumerate(s["items"], 2):
            ws.cell(n, 5).fill = SKLAD_CAT_FILL[cats(i)[0]]
            for j, d in enumerate(s["dst"], 6):
                if d in i["to"]:
                    ws.cell(n, j).fill = SKLAD_CAT_FILL[key(i["to"][d]["cat"])]

    def paint_one(d):
        def paint(ws):
            for n, i in enumerate([i for i in s["items"] if d in i["to"]], 2):
                ws.cell(n, 4).fill = SKLAD_CAT_FILL[key(i["to"][d]["cat"])]
        return paint

    sheets = [("Taqsimot", ["Dori", "Ishlab chiqaruvchi", "Skladda qoldiq", "Eng yaqin muddat", "Toifa"] + s["dst"] + ["Jami beriladi", "Skladda qoladi"],
               [[i["name"], i["producer"], i["stock"], dates(i["srok"]), " + ".join(KOCH_CAT[c]["title"] for c in cats(i))] +
                [(i["to"].get(d) or {}).get("qty") for d in s["dst"]] + [i["total"], i["left"]] for i in s["items"]],
               [46, 26, 11, 12, 34] + [13] * len(s["dst"]) + [11, 11], (3,), (), paint_all)]
    for d in s["dst"]:
        rows = [[i["name"], i["producer"], i["to"][d]["qty"], KOCH_CAT[key(i["to"][d]["cat"])]["title"], dates(i["srok"])] +
                ([i["to"][d]["summa"]] if prices else []) + [i["to"][d]["status"]] for i in s["items"] if d in i["to"]]
        sheets.append((re.sub(r"[\\/?*\[\]:]", " ", d), ["Dori", "Ishlab chiqaruvchi", "Soni", "Toifa", "Eng yaqin muddat"] +
                       (["Summa"] if prices else []) + ["Holat"], rows, [46, 26, 9, 14, 12] + ([14] if prices else []) + [10], (4,),
                       (5,) if prices else (), paint_one(d)))
    if s["excluded"]:
        sheets.append(("Berilmaydi", ["Qayerga", "Dori", "Ishlab chiqaruvchi", "Soni", "Muddat", "Sabab"],
                       [[r["dst"], r["name"], r["producer"], r["qty"], dates(r["srok"]), r["why"]] for r in s["excluded"]],
                       [22, 46, 26, 9, 12, 34], (4,), ()))
    return xlsx_multi(sheets)


def build_fom_boron(b, items, fom=None):
    """FOM bron fayli (sklad → dorixona): narxsiz. fom: {task_id: (FOM'ga yozilgan soni, izoh)} — SQL rejimida."""
    head = ["№", "Dori", "Ishlab chiqaruvchi", "Soni", "Muddat", "Toifa"]
    rows = []
    for i, it in enumerate(items, 1):
        row = [i, it["name"], it["producer"], it["qty"], datetime.date.fromisoformat(it["srok"]) if it["srok"] else None,
               KOCH_CAT.get(it["cat"], {}).get("title", it["cat"] or "")]
        if fom is not None:
            row += list(fom.get(it["task_id"], (0, "FOM'da qoldiq yo'q")))
        rows.append(tuple(row))
    if fom is not None:
        head += ["FOM'ga yozildi", "Izoh"]
    return xlsx_bytes("Bron %d" % b["id"], head, rows, [6, 46, 28, 10, 13, 16, 15, 24][:len(head)],
                      datecols=(4,), numcols=(3, 6) if fom is not None else (3,))


def build_narx(rows):
    data = [[r["name"], r["producer"], r["firm"], r["prih"], r["roz"], r["sana"]] for r in rows]
    return xlsx_bytes("Narx", ["Tovar", "Ishlab chiqaruvchi", "Firma", "Kirim narxi", "Sotuv narxi", "Oxirgi kirim"],
                      data, [46, 26, 24, 13, 13, 13], datecols=(5,), numcols=(3, 4))


def build_order(firm, items):
    """Yetkazib beruvchiga buyurtma varag'i (panel → Buyurtma → Excel)."""
    data = [[i["name"], i.get("producer", ""), to_float(i.get("stock")), to_float(i.get("sales")), to_float(i["qty"]),
             to_float(i.get("price")), to_float(i["qty"]) * to_float(i.get("price"))] for i in items]
    data.append(["JAMI", "", None, None, sum(r[4] for r in data), None, sum(r[6] for r in data)])
    content = xlsx_bytes("Buyurtma", ["Tovar", "Ishlab chiqaruvchi", "Qoldiq", "Sotuv", "Buyurtma (dona)", "Kirim narxi",
                                      "Summa"], data, [46, 26, 10, 10, 15, 13, 15], numcols=(2, 3, 4, 5, 6))
    wb = load_workbook(io.BytesIO(content))
    ws = wb.active
    ws.insert_rows(1, 2)
    ws["A1"] = "%s · buyurtma · %s" % (firm, datetime.date.today().strftime("%d.%m.%Y"))
    ws["A1"].font = Font(bold=True, size=13)
    for c in ws[ws.max_row]:
        c.font = Font(bold=True)
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = "A3:G%d" % (ws.max_row - 1)
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def build_hisob(rows):
    data = [[r["org"], r["debit"], r["credit"], r["debit"] - r["credit"]] for r in rows]
    return xlsx_bytes("Hisob", ["Tashkilot", "Debet", "Kredit", "Qoldiq (debet - kredit)"], data, [52, 16, 16, 20],
                      numcols=(1, 2, 3))


# ----------------------------------------------------------------------------
# Telegram API (urllib)
# ----------------------------------------------------------------------------
class TGError(Exception):
    def __init__(self, msg, code=None, retry_after=None):
        super().__init__(msg)
        self.code = code
        self.retry_after = retry_after

    @property
    def gone(self):
        """Foydalanuvchi botni bloklagan / akkaunti o'chgan: unga yuborib bo'lmaydi."""
        s = str(self).lower()
        return self.code == 403 or (self.code == 400 and "chat not found" in s)


def cut_html(text, limit=4000):
    """Telegram chegarasi uchun qisqartiradi: yarim qolgan teg yoki &entity; ni olib tashlaydi."""
    if len(text) <= limit:
        return text
    t = text[:limit]
    t = re.sub(r"<[^>]*$", "", t)
    t = re.sub(r"&[#\w]*$", "", t)
    return t + "…"


class TG:
    RETRIES = 3

    def __init__(self, token, base="https://api.telegram.org"):
        self.base = base.rstrip("/")
        self.url = "%s/bot%s/" % (self.base, token)
        host = urllib.parse.urlparse(self.base).hostname
        handlers = [urllib.request.ProxyHandler({})] if host in ("127.0.0.1", "localhost") else []
        self.opener = urllib.request.build_opener(*handlers)
        self.sleep = time.sleep

    def set_token(self, token):
        self.url = "%s/bot%s/" % (self.base, token)

    def _open_once(self, req, timeout):
        try:
            with self.opener.open(req, timeout=timeout) as r:
                res = json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", "replace")
            retry = None
            try:
                retry = json.loads(body).get("parameters", {}).get("retry_after")
            except ValueError:
                pass
            raise TGError("HTTP %s: %s" % (e.code, body[:300]), e.code, retry)
        if not res.get("ok"):
            raise TGError(str(res), res.get("error_code"))
        return res["result"]

    def _open(self, req, timeout):
        """429 (Too Many Requests) da kutib qayta yuboradi. Ulanish o'rnatilmagan bo'lsa (URLError) ham qayta
        urinadi: bunda so'rov Telegramga yetib bormagan, demak xabar ikki marta ketmaydi."""
        for i in range(self.RETRIES):
            last = i == self.RETRIES - 1
            try:
                return self._open_once(req, timeout)
            except TGError as e:
                if e.code != 429 or last:
                    raise
                self.sleep(min(int(e.retry_after or 3), 60) + 1)
            except urllib.error.URLError:
                if last:
                    raise
                self.sleep(2 * (i + 1))

    no_style = False  # Telegram tugma rangini (KeyboardButton.style) rad etgan bo'lsa — keyin rangsiz yuboriladi

    @staticmethod
    def _strip_style(params):
        rm = params.get("reply_markup") or {}
        key = "inline_keyboard" if "inline_keyboard" in rm else "keyboard"
        kb = rm.get(key)
        if not kb or not any(isinstance(b, dict) and "style" in b for row in kb for b in row):
            return None
        p = dict(params)
        p["reply_markup"] = dict(rm, **{key: [[{k: v for k, v in b.items() if k != "style"} if isinstance(b, dict) else b
                                               for b in row] for row in kb]})
        return p

    @staticmethod
    def _cyr(params):
        """Foydalanuvchiga ko'rinadigan matn kirillga (uzkir). Pastki klaviatura tugmalari eslab qolinadi: foydalanuvchi
        kirill tugmani bossa, bot uni koddagi asl (lotin) nomi bilan taniydi."""
        if not params or not any(k in params for k in ("text", "caption", "reply_markup", "commands")):
            return params
        p = dict(params)
        for k in ("text", "caption"):
            if isinstance(p.get(k), str):
                p[k] = UK.lat2cyr(p[k])
        rm = p.get("reply_markup")
        if isinstance(rm, dict):
            for row in rm.get("keyboard") or []:
                for b in row:
                    t = b.get("text") if isinstance(b, dict) else b
                    if isinstance(t, str):
                        UK.KB_BACK[UK.lat2cyr(t)] = UK.strip_marks(t)
            p["reply_markup"] = UK.markup(rm)
        if isinstance(p.get("commands"), list):
            p["commands"] = [dict(c, description=UK.lat2cyr(c.get("description", ""))) for c in p["commands"]]
        return p

    def call(self, method, params=None, timeout=60):
        params = self._cyr(params or {})
        plain = self._strip_style(params)
        if plain is not None and self.no_style:
            params = plain
        data = json.dumps(params).encode("utf-8")
        req = urllib.request.Request(self.url + method, data, {"Content-Type": "application/json"})
        try:
            return self._open(req, timeout)
        except TGError as e:
            if plain is None or self.no_style or e.code != 400 or "parse entities" in str(e) or "not modified" in str(e):
                raise
            logging.warning("Telegram tugma rangini (style) qabul qilmadi, rangsiz qayta yuborildi: %s", str(e)[:200])
            self.no_style = True
            return self.call(method, plain, timeout)

    def send_document(self, chat_id, filename, content, caption=None):
        boundary = uuid.uuid4().hex
        parts = []

        def field(name, value):
            parts.append(('--%s\r\nContent-Disposition: form-data; name="%s"\r\n\r\n%s\r\n'
                          % (boundary, name, value)).encode("utf-8"))

        field("chat_id", chat_id)
        if caption:
            field("caption", UK.lat2cyr(caption))
            field("parse_mode", "HTML")
        parts.append(('--%s\r\nContent-Disposition: form-data; name="document"; filename="%s"\r\n'
                      'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n'
                      % (boundary, filename)).encode("utf-8") + content + b"\r\n")
        parts.append(("--%s--\r\n" % boundary).encode("utf-8"))
        req = urllib.request.Request(self.url + "sendDocument", b"".join(parts),
                                     {"Content-Type": "multipart/form-data; boundary=" + boundary})
        return self._open(req, 120)

    def _html_call(self, method, p):
        """HTML belgilashda xato bo'lsa (masalan yopilmagan teg), oddiy matn bilan qayta yuboradi."""
        try:
            return self.call(method, p)
        except TGError as e:
            if "parse entities" not in str(e):
                raise
            p = dict(p)
            p.pop("parse_mode")
            p["text"] = html.unescape(re.sub(r"<[^>]+>", "", p["text"]))
            return self.call(method, p)

    def send(self, chat_id, text, markup=None):
        p = {"chat_id": chat_id, "text": cut_html(text), "parse_mode": "HTML", "disable_web_page_preview": True}
        if markup:
            p["reply_markup"] = markup
        return self._html_call("sendMessage", p)

    def edit(self, chat_id, message_id, text, markup=None):
        p = {"chat_id": chat_id, "message_id": message_id, "text": cut_html(text), "parse_mode": "HTML",
             "disable_web_page_preview": True}
        if markup:
            p["reply_markup"] = markup
        return self._html_call("editMessageText", p)

    def delete(self, chat_id, message_id):
        """deleteMessage. Hech qachon xato ko'tarmaydi (xabar allaqachon o'chgan, 48 soatdan eski va h.k.)."""
        try:
            return self.call("deleteMessage", {"chat_id": chat_id, "message_id": message_id}, timeout=15)
        except Exception as e:  # noqa
            logging.debug("Xabarni o'chirib bo'lmadi (%s, %s): %s", chat_id, message_id, e)
            return False


def reply_kb(rows, **kw):
    """Tugma: matn yoki {"text", "style"} (style=None — kalit yuborilmaydi)."""
    d = {"keyboard": [[{"text": t} if isinstance(t, str) else {k: v for k, v in t.items() if v is not None} for t in row]
                      for row in rows], "resize_keyboard": True}
    d.update(kw)
    return d


# Avvalgi versiyaning pastki tugmalari: foydalanuvchida eski klaviatura (kirill) qolgan bo'lishi mumkin. Bot qayta ishga
# tushgach KB_BACK bo'sh — shuning uchun matn boshidagi kirill nom koddagi lotin nomga almashtiriladi (sonlar o'zgarmaydi).
LEGACY_BUTTONS = ("⏰ Srok", "📦 30+ kun", "🧾 Buyurtma", "💰 Hisob", "📊 Qoldiq", "🏢 Firma", "👥 Foydalanuvchilar",
                  "📢 Xabar yuborish", "🔄 Yangilash", "⚙️ Admin panel", "🏠 Menyu", "📥 Keladigan", "📋 Bugungi holat",
                  "🏭 Taqsimot", "🚨 Zudlik") + tuple("%s %s" % (k["icon"], k["title"]) for k in KOCH_CATS)


def legacy_button(text):
    """Eski pastki tugma matni (kirill yoki lotin) → koddagi lotin nomi; boshqa matn o'zgarmaydi."""
    if text in UK.KB_BACK:
        return UK.KB_BACK[text]
    for lat in LEGACY_BUTTONS:
        cyr = UK.lat2cyr(lat)
        if cyr != lat and text.startswith(cyr):
            return lat + text[len(cyr):]
    return text


def inline_kb(rows):
    """Tugma: (matn, callback) yoki (matn, callback, style) — style=None bo'lsa kalit yuborilmaydi."""
    return {"inline_keyboard": [[dict({"text": b[0], "callback_data": b[1]}, **({"style": b[2]} if len(b) > 2 and b[2] else {}))
                                 for b in row] for row in rows]}


ENTITY_TAGS = {"bold": "b", "italic": "i", "underline": "u", "strikethrough": "s", "spoiler": "tg-spoiler",
               "code": "code", "pre": "pre", "blockquote": "blockquote"}


def entities_to_html(text, entities):
    """Telegramda yozilgan formatlangan matnni (qalin, kursiv, havola...) HTML ga aylantiradi.
    Telegram offsetlari UTF-16 birliklarida beriladi."""
    text = text or ""
    if not entities:
        return esc(text)
    raw = text.encode("utf-16-le")
    n = len(raw) // 2
    opens, closes = {}, {}
    for e in sorted(entities, key=lambda e: (e["offset"], -e["length"])):
        a, b = e["offset"], e["offset"] + e["length"]
        if e["type"] == "text_link":
            o, c = '<a href="%s">' % html.escape(e.get("url", ""), quote=True), "</a>"
        elif e["type"] in ENTITY_TAGS:
            tag = ENTITY_TAGS[e["type"]]
            o, c = "<%s>" % tag, "</%s>" % tag
        else:
            continue
        opens.setdefault(a, []).append(o)
        closes.setdefault(b, []).insert(0, c)
    cuts = sorted({0, n} | set(opens) | set(closes))
    out = []
    for a, b in zip(cuts, cuts[1:] + [None]):
        out += closes.get(a, []) + opens.get(a, [])
        if b is not None:
            out.append(esc(raw[a * 2:b * 2].decode("utf-16-le", "replace")))
    return "".join(out)


def simple_markup(text):
    """Paneldagi matn uchun: **qalin**, __kursiv__ → HTML. Qolgani xavfsiz (escape) qilinadi."""
    t = esc(text or "")
    t = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", t, flags=re.S)
    t = re.sub(r"__(.+?)__", r"<i>\1</i>", t, flags=re.S)
    return t


# ----------------------------------------------------------------------------
# Excel makrosini ishga tushirish ("Yangilash")
# ----------------------------------------------------------------------------
def default_runner(vbs_path, xlsm_path, timeout, days=0):
    """cscript refresh.vbs <xlsm> [kunlar] ni ishga tushiradi. (qaytish kodi, matn) qaytaradi.
    kunlar > 0 — sotuv davri (DateBegin = bugun - kunlar); 0 — Exceldagi davr uzunligi saqlanadi."""
    kw = {}
    if os.name == "nt":  # qora cmd oynasi ko'rinmasin (Excel o'zi ham yashirin ochiladi)
        kw["creationflags"] = 0x08000000  # CREATE_NO_WINDOW
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        si.wShowWindow = 0  # SW_HIDE
        kw["startupinfo"] = si
    args = ["cscript", "//nologo", vbs_path, xlsm_path] + ([str(int(days))] if days and int(days) > 0 else [])
    before = _automation_excels(kw) if os.name == "nt" else {}
    try:
        p = subprocess.run(args, capture_output=True, timeout=timeout, **kw)
    except subprocess.TimeoutExpired:
        # subprocess faqat cscript ni o'ldiradi; u COM orqali ochgan yashirin EXCEL.EXE faylni band qilib qoladi va
        # keyingi har bir «Yangilash» «fayl band» bilan tugaydi. Shu yangilash ochgan Excel(lar)ni yopamiz.
        killed = _kill_new_excels(before, kw) if os.name == "nt" else []
        logging.error("Yangilash %d soniyada tugamadi; yopilgan Excel: %s", timeout, killed or "yo'q")
        return -1, "Excel %d soniyada javob bermadi (makros xabar oynasida to'xtagan bo'lishi mumkin)%s" % (
            timeout, "; yashirin Excel yopildi" if killed else "")
    out = (p.stdout or b"").decode("cp866", "replace") + (p.stderr or b"").decode("cp866", "replace")
    return p.returncode, out.strip()


def _automation_excels(kw):
    """{pid: buyruq qatori} — COM orqali (/automation) ochilgan EXCEL.EXE lar. Xato bo'lsa {}."""
    ps = ("Get-CimInstance Win32_Process -Filter \"Name='EXCEL.EXE'\" | "
          "ForEach-Object { [string]$_.ProcessId + [char]9 + [string]$_.CommandLine }")
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                           capture_output=True, timeout=60, **kw)
    except Exception:  # noqa
        return {}
    out = {}
    for ln in (r.stdout or b"").decode("utf-8", "replace").splitlines():
        pid, _, cmd = ln.partition("\t")
        if pid.strip().isdigit() and "automation" in cmd.lower():
            out[int(pid)] = cmd
    return out


def _kill_new_excels(before, kw):
    """Yangilashdan oldin bo'lmagan /automation Excel jarayonlarini yopadi (foydalanuvchi ochgan Excelga tegmaydi)."""
    killed = []
    for pid in set(_automation_excels(kw)) - set(before):
        try:
            subprocess.run(["taskkill", "/F", "/T", "/PID", str(pid)], capture_output=True, timeout=30, **kw)
            killed.append(pid)
        except Exception:  # noqa
            logging.warning("Excel (PID %s) yopilmadi", pid, exc_info=True)
    return killed


def parse_status(path):
    """refresh_status.txt dan qadamlar holati: ({varaq: ('ok'|'err', matn)}, tugadimi, xato)."""
    steps, done, fail = {}, False, None
    try:
        with open(path, "rb") as f:
            raw = f.read()
        try:
            content = raw.decode("utf-8-sig")
        except UnicodeDecodeError:
            content = raw.decode("cp1251", "replace")
        if True:
            for line in content.splitlines():
                line = line.strip()
                if line.startswith("OK "):
                    name, _, n = line[3:].rpartition("|")
                    steps[name.strip()] = ("ok", n.strip())
                elif line.startswith("ERR "):
                    name, _, msg = line[4:].partition("|")
                    steps[name.strip()] = ("err", msg.strip())
                elif line.startswith("DONE"):
                    done = True
                elif line.startswith("FAIL"):
                    fail = line[4:].strip() or "noma'lum xato"
    except OSError:
        pass
    return steps, done, fail


# ----------------------------------------------------------------------------
# Bot
# ----------------------------------------------------------------------------
class Bot:
    def __init__(self, cfg, tg=None, db_path=None, runner=None):
        self.cfg = cfg
        self.tg = tg or TG(cfg["bot"]["token"].strip(), cfg["bot"].get("api_base", "https://api.telegram.org"))
        self.admins = {int(x) for x in re.findall(r"-?\d+", cfg["bot"].get("admin_ids", ""))}
        self.data = Data(cfg)
        self.db_path = db_path or os.path.join(BASE_DIR, "bot.db")
        self._local = threading.local()
        self._shared_db = None  # ":memory:" (testlar) uchun bitta umumiy ulanish
        self.dblock = threading.RLock()
        self.state = {}  # admin_id -> (kind, payload)
        self.pending_text = {}
        self.search = {}
        self.fsearch = {}
        self.fmode = {}  # admin_id -> (tags, firma nomi)
        self.filsearch = {}  # admin_id -> filiallar ro'yxati (filial xodimi tanlash)
        self._panels = {}  # chat_id -> panel xabarining message_id (settings: panel_<id>)
        self._panel_last = {}  # chat_id -> (matn, tugmalar): panelning oxirgi holati (Excel'dan keyin qayta tushirish uchun)
        self._kbclean = set()  # eski pastki klaviaturasi olib tashlangan chatlar (settings: kbclean_<id>)
        self.fret = {}  # chat_id -> firma tanlangach qaytiladigan bo'lim (srok/old/need/acc/stock/rep yoki None — bosh menyu)
        self.sqall = set()  # xodim/admin «📊 Sotuv va qoldiq» da «🌐 Hamma firmalar» ni tanlagan (firma so'ralmaydi)
        self._abc_cache = {}  # ko'chirish ABC: sana -> (imzo, {task_id: sinf})
        self._anames = {}  # admin_id -> Telegram ismi (buyurtma so'roviga kim javob berdi)
        self.runner = runner or default_runner
        self.refreshing = False
        self._snap_lock = threading.Lock()
        self._detect_lock = threading.Lock()  # avto-aniqlash bir vaqtda ikki oqimda ishlamaydi (ogohlantirish ikki marta ketmaydi)
        self.bc_running = False
        self.auto_running = False
        self.rstate = {"running": False, "steps": {}, "ok": None, "msg": "", "t0": 0, "t1": None}
        self.last_poll = 0.0
        self.conflict_at = 0.0  # oxirgi «409 Conflict»: shu token bilan boshqa joyda ham bot ishlayapti
        self._conflict_logged = 0.0
        self.username = ""
        self.rlock = threading.Lock()
        self.sleep = time.sleep
        self._init_db()
        if self.setting("koch_installed") is None:
            # mavjud bazaga birinchi o'rnatilgan kuni (kun o'rtasida) eslatma / hisobot yubormaymiz — ertangi kundan boshlanadi
            today = now().date().isoformat()
            self.set_setting("koch_installed", today)
            if self.db.execute("SELECT COUNT(*) FROM clients").fetchone()[0]:
                for k in ("koch_rem1_date", "koch_rem2_date", "koch_rem3_date"):
                    self.set_setting(k, today)

    # ---------------- token va adminlar ----------------
    def token_ok(self):
        t = self.cfg["bot"].get("token", "").strip()
        return bool(t) and not t.startswith("BU_YERGA")

    def set_token(self, token):
        """Tokenni Telegramda tekshiradi (getMe), keyin ishlayotgan botga va config.ini ga yozadi. @username qaytaradi."""
        token = token.strip()
        if not re.match(r"^\d{6,}:[A-Za-z0-9_-]{30,}$", token):
            raise ValueError("Token shakli noto'g'ri. @BotFather bergan to'liq tokenni kiriting (masalan 123456789:AAH...).")
        probe = TG(token, self.cfg["bot"].get("api_base", "https://api.telegram.org"))
        try:
            me = probe.call("getMe", timeout=15)
        except TGError as e:
            raise ValueError("Telegram tokenni qabul qilmadi: %s" % e)
        except Exception as e:  # noqa
            raise ValueError("Telegramga ulanib bo'lmadi (internetni tekshiring): %s" % e)
        self.tg.set_token(token)
        self.cfg["bot"]["token"] = token
        write_ini("bot", "token", token)
        self.username = me.get("username", "")
        logging.info("Bot tokeni almashtirildi: @%s", self.username)
        return self.username

    def set_admins(self, ids):
        ids = sorted({int(x) for x in ids})
        self.admins = set(ids)
        value = ", ".join(str(i) for i in ids)
        self.cfg["bot"]["admin_ids"] = value
        write_ini("bot", "admin_ids", value)

    # ---------------- mahsulot: nom va sozlash ustasi ----------------
    def app_get(self, key, default=""):
        return (self.cfg.get("app", key, fallback=default) if self.cfg.has_section("app") else default).strip()

    def app_set(self, key, value):
        if not self.cfg.has_section("app"):
            self.cfg.add_section("app")
        self.cfg["app"][key] = str(value)
        write_ini("app", key, value)

    @property
    def app_name(self):
        """Dorixona nomi (sozlash ustasida kiritiladi), bo'lmasa mahsulot nomi."""
        return self.app_get("name") or BRAND

    @property
    def setup_done(self):
        """Sozlash ustasi tugaganmi. [app] setup_done yo'q bo'lsa (eski o'rnatish) — token bor bo'lsa tugagan hisoblanadi."""
        v = self.app_get("setup_done")
        return v.lower() in ("yes", "1", "true", "ha") if v else self.token_ok()

    # ---------------- sozlamalar / baza ----------------
    def _connect(self):
        con = sqlite3.connect(self.db_path, check_same_thread=False, timeout=30)
        con.row_factory = sqlite3.Row
        return con

    @property
    def db(self):
        """Har bir oqim (bot, panel so'rovlari, fon ishlari) o'z ulanishini oladi. Avval bitta ulanish hamma oqimlarda
        ishlatilgani uchun «bad parameter or other API misuse» xatosi chiqardi. Yozishlar baribir dblock bilan."""
        if self.db_path == ":memory:":
            if self._shared_db is None:
                self._shared_db = self._connect()
            return self._shared_db
        con = getattr(self._local, "con", None)
        if con is None:
            con = self._local.con = self._connect()
        return con

    def _init_db(self):
        if self.db_path != ":memory:":
            self.db.execute("PRAGMA journal_mode=WAL")
        self.db.executescript("""
            CREATE TABLE IF NOT EXISTS clients(
                chat_id INTEGER PRIMARY KEY, phone TEXT, name TEXT, username TEXT,
                status TEXT NOT NULL DEFAULT 'pending', role TEXT NOT NULL DEFAULT 'firm', created TEXT);
            CREATE TABLE IF NOT EXISTS client_firms(chat_id INTEGER, firm TEXT, PRIMARY KEY(chat_id, firm));
            CREATE TABLE IF NOT EXISTS client_accounts(chat_id INTEGER, account TEXT, PRIMARY KEY(chat_id, account));
            CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT);
            CREATE TABLE IF NOT EXISTS invites(
                phone9 TEXT PRIMARY KEY, phone TEXT, name TEXT, role TEXT NOT NULL DEFAULT 'firm',
                firms TEXT, accounts TEXT, created TEXT);
            CREATE TABLE IF NOT EXISTS acc_requests(
                id INTEGER PRIMARY KEY AUTOINCREMENT, chat_id INTEGER, account TEXT,
                status TEXT NOT NULL DEFAULT 'pending', created TEXT, UNIQUE(chat_id, account));
            CREATE TABLE IF NOT EXISTS activity(
                id INTEGER PRIMARY KEY AUTOINCREMENT, chat_id INTEGER, ts TEXT, action TEXT);
            CREATE INDEX IF NOT EXISTS activity_ts ON activity(ts);
            CREATE TABLE IF NOT EXISTS supplier_blacklist(firm TEXT PRIMARY KEY, reason TEXT, created TEXT);
            CREATE TABLE IF NOT EXISTS item_blacklist(
                firm TEXT, item TEXT, reason TEXT, created TEXT, PRIMARY KEY(firm, item));
            CREATE TABLE IF NOT EXISTS supplier_contacts(firm TEXT PRIMARY KEY, phone TEXT, username TEXT, updated TEXT);
            CREATE TABLE IF NOT EXISTS orders(
                id INTEGER PRIMARY KEY AUTOINCREMENT, firm TEXT, created TEXT, cover INTEGER, items TEXT,
                total REAL, n INTEGER);
            CREATE TABLE IF NOT EXISTS transfers(
                id INTEGER PRIMARY KEY AUTOINCREMENT, created TEXT, src TEXT, dst TEXT, name TEXT, producer TEXT,
                qty REAL, turi TEXT, status TEXT NOT NULL DEFAULT 'sent', sent_by INTEGER, sent_at TEXT,
                recv_by INTEGER, recv_at TEXT, prih REAL, summa REAL, kkey TEXT);
            CREATE INDEX IF NOT EXISTS transfers_kkey ON transfers(kkey, created);
            CREATE INDEX IF NOT EXISTS transfers_dst ON transfers(dst, status);
            CREATE TABLE IF NOT EXISTS koch_tasks(
                id INTEGER PRIMARY KEY AUTOINCREMENT, sana TEXT, src TEXT, dst TEXT, name TEXT, producer TEXT, firm TEXT,
                turi TEXT, sabab TEXT, qty_plan REAL, prih REAL, summa REAL, srok TEXT,
                status TEXT NOT NULL DEFAULT 'yangi', qty_given REAL, qty_received REAL, given_by INTEGER, given_at TEXT,
                received_by INTEGER, received_at TEXT, note TEXT, created_at TEXT, kkey TEXT, UNIQUE(sana, kkey));
            CREATE INDEX IF NOT EXISTS koch_tasks_src ON koch_tasks(sana, src, status);
            CREATE INDEX IF NOT EXISTS koch_tasks_dst ON koch_tasks(sana, dst, status);
            CREATE TABLE IF NOT EXISTS koch_events(
                id INTEGER PRIMARY KEY AUTOINCREMENT, task_id INTEGER, action TEXT, user_id INTEGER, user_name TEXT,
                qty REAL, note TEXT, ts TEXT);
            CREATE INDEX IF NOT EXISTS koch_events_task ON koch_events(task_id);
            CREATE TABLE IF NOT EXISTS fom_boron(
                id INTEGER PRIMARY KEY AUTOINCREMENT, sana TEXT, src TEXT, dst TEXT, cats TEXT, n INTEGER, qty REAL,
                created_by TEXT, created_at TEXT, mode TEXT, status TEXT, doc_no TEXT, error TEXT);
            CREATE INDEX IF NOT EXISTS fom_boron_sana ON fom_boron(sana, src);
            CREATE TABLE IF NOT EXISTS fom_boron_items(
                boron_id INTEGER, task_id INTEGER, name TEXT, producer TEXT, cat TEXT, qty REAL, srok TEXT,
                PRIMARY KEY(boron_id, task_id));
            CREATE INDEX IF NOT EXISTS fom_boron_items_task ON fom_boron_items(task_id);
            CREATE TABLE IF NOT EXISTS fom_boron_lines(
                boron_id INTEGER, task_id INTEGER, incomeln INTEGER, kol REAL, price REAL, srok TEXT);
            CREATE INDEX IF NOT EXISTS fom_boron_lines_b ON fom_boron_lines(boron_id);
        """)
        if "short" not in {r["name"] for r in self.db.execute("PRAGMA table_info(fom_boron)")}:
            self.db.execute("ALTER TABLE fom_boron ADD COLUMN short TEXT")
        # eski bazaga yangi ustunlar (mavjud ma'lumot saqlanadi)
        cols = {r["name"] for r in self.db.execute("PRAGMA table_info(clients)")}
        for col in ("last_seen", "left_at", "filial"):
            if col not in cols:
                self.db.execute("ALTER TABLE clients ADD COLUMN %s TEXT" % col)
        if "filial" not in {r["name"] for r in self.db.execute("PRAGMA table_info(invites)")}:
            self.db.execute("ALTER TABLE invites ADD COLUMN filial TEXT")
        # ko'chirish: 3 toifa (cat) va kartochka maydonlari; eski topshiriqlarga toifa turi/sabab dan yoziladi
        cols = {r["name"] for r in self.db.execute("PRAGMA table_info(koch_tasks)")}
        for col, typ in (("cat", "TEXT"), ("days_idle", "INTEGER"), ("reason", "TEXT"), ("dst_sales", "REAL"),
                         ("dst_stock", "REAL"), ("src_sales", "REAL"), ("src_stock", "REAL")):
            if col not in cols:
                self.db.execute("ALTER TABLE koch_tasks ADD COLUMN %s %s" % (col, typ))
        for r in self.db.execute("SELECT id, turi, sabab FROM koch_tasks WHERE cat IS NULL").fetchall():
            cat, idle = koch_category_of(r["turi"], r["sabab"])
            self.db.execute("UPDATE koch_tasks SET cat=?, days_idle=? WHERE id=?", (cat or "oddiy", idle, r["id"]))
        self.db.execute("CREATE INDEX IF NOT EXISTS koch_tasks_cat ON koch_tasks(sana, cat, status)")
        # avto-aniqlash (dizayn/AVTO-ANIQLASH.md): yaratilgandagi qoldiq surati (qayta shakllantirishda o'zgarmaydi), manba va isbot;
        # auto_idle / auto_seen — «Berdim» dan keyin qoldiq o'zgarmagan yangilashlar soni, auto_warn — adminga yuborilgan ogohlantirishlar;
        # dst_stock1 — «Berdim» bosilgandagi oluvchi qoldig'i (oluvchidagi ko'payish shundan o'lchanadi)
        cols = {r["name"] for r in self.db.execute("PRAGMA table_info(koch_tasks)")}
        for col, typ in (("src_stock0", "REAL"), ("dst_stock0", "REAL"), ("auto_by", "TEXT"), ("auto_note", "TEXT"),
                         ("auto_idle", "INTEGER"), ("auto_seen", "TEXT"), ("auto_warn", "TEXT"), ("dst_stock1", "REAL")):
            if col not in cols:
                self.db.execute("ALTER TABLE koch_tasks ADD COLUMN %s %s" % (col, typ))
        self.db.execute("UPDATE koch_tasks SET src_stock0=src_stock WHERE src_stock0 IS NULL AND src_stock IS NOT NULL")
        self.db.execute("UPDATE koch_tasks SET dst_stock0=dst_stock WHERE dst_stock0 IS NULL AND dst_stock IS NOT NULL")
        # buyurtma so'rovlari (dizayn/BUYURTMA-SOROV.md): jadvallar buyurtma.py da; adminlarga ketgan kartalar — har adminda bittadan
        BY.init(self.db)
        self.db.execute("CREATE TABLE IF NOT EXISTS order_req_msgs(req_id INTEGER, chat_id INTEGER, message_id INTEGER, "
                        "PRIMARY KEY(req_id, chat_id))")
        # sozlamalar: eski kalitlar qiymati yangisiga ko'chiriladi (eskisi o'chirilmaydi)
        for old, new in (("koch_deadline_zudlik", "koch_deadline_shoshilinch"), ("koch_deadline_zudlik", "koch_deadline_assortiment"),
                         ("koch_rem_report_on", "alert_daily_report"), ("alert_price", "alert_loss_price"),
                         ("alert_baddate", "alert_bad_dates")):
            self.db.execute("INSERT OR IGNORE INTO settings SELECT ?, value FROM settings WHERE key=?", (new, old))
        self.db.commit()

    def setting(self, key, default=None):
        r = self.db.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
        return r["value"] if r else default

    def set_setting(self, key, value):
        with self.dblock:
            self.db.execute("INSERT OR REPLACE INTO settings VALUES(?,?)", (key, str(value)))
            self.db.commit()

    @property
    def srok_days(self):
        return int(self.setting("srok_days", self.cfg.get("auto", "srok_days", fallback="90")))

    @property
    def auto_on(self):
        return self.setting("auto_on", "1" if self.cfg.get("auto", "enabled", fallback="yes").lower() in
                            ("yes", "1", "true", "ha") else "0") == "1"

    @property
    def auto_time(self):
        h, m = self.cfg.get("auto", "time", fallback="09:00").split(":")
        return datetime.time(int(h), int(m))

    @property
    def sale_days(self):
        """Sotuv davri (kun), panelda belgilanadi. 0 — Excel'dagi davr (DateBegin..DateEnd) uzunligi saqlanadi."""
        try:
            n = int(self.setting("sale_days", "0") or 0)
        except ValueError:
            return 0
        return n if 1 <= n <= 365 else 0

    @property
    def cover(self):
        return float(self.cfg.get("order", "cover", fallback="1"))

    def client(self, cid):
        return self.db.execute("SELECT * FROM clients WHERE chat_id=?", (cid,)).fetchone()

    def firms_of(self, cid):
        return [r["firm"] for r in self.db.execute("SELECT firm FROM client_firms WHERE chat_id=? ORDER BY firm", (cid,))]

    def accounts_of(self, cid):
        return [r["account"] for r in self.db.execute("SELECT account FROM client_accounts WHERE chat_id=? ORDER BY account", (cid,))]

    def set_firms(self, cid, firms):
        with self.dblock:
            self.db.execute("DELETE FROM client_firms WHERE chat_id=?", (cid,))
            for f in firms:
                self.db.execute("INSERT OR IGNORE INTO client_firms VALUES(?,?)", (cid, norm(f)))
            self.db.commit()

    def set_status(self, cid, status):
        with self.dblock:
            self.db.execute("UPDATE clients SET status=? WHERE chat_id=?", (status, cid))
            self.db.commit()

    def tags_for(self, c):
        """None = hamma tovar (xodim), aks holda firma belgilari to'plami. Filial xodimi firma ma'lumotini ko'rmaydi."""
        if c["role"] == "staff":
            return None
        if c["role"] == "filial":
            return set()
        return {norm(f) for f in self.firms_of(c["chat_id"])}

    def active_clients(self):
        """Faol va botni tark etmagan (bloklamagan) foydalanuvchilar."""
        return self.db.execute("SELECT * FROM clients WHERE status='active' AND left_at IS NULL").fetchall()

    def can_use(self, c):
        return c["role"] == "staff" or self.is_filial(c) or bool(self.firms_of(c["chat_id"]))

    # ---------------- filial xodimi (ko'chirish) ----------------
    @staticmethod
    def is_filial(c):
        """role='filial' va filial biriktirilgan. Filial HAR DOIM bazadan olinadi (foydalanuvchi matnidan emas)."""
        return bool(c) and c["role"] == "filial" and bool(norm(c["filial"]))

    @property
    def sklad(self):
        return self.data.sklad_name

    def is_sklad(self, filial):
        """Sklad xodimi — filial xodimi, filiali sklad: panel, tasdiqlash va eslatmalar umumiy."""
        return norm(filial) == self.sklad

    def set_role(self, cid, role, filial=None):
        """Rolni o'zgartiradi: firm | staff | filial. Filial faqat role='filial' da saqlanadi, aks holda tozalanadi."""
        with self.dblock:
            self.db.execute("UPDATE clients SET role=?, filial=? WHERE chat_id=?",
                            (role, norm(filial) if role == "filial" and filial else None, cid))
            self.db.commit()
        self.fmode.pop(cid, None)
        self.state.pop(cid, None)

    def role_label(self, c):
        """Ro'yxatlar uchun: «xodim», «filial · <nom>» yoki firma belgilari."""
        if c["role"] == "staff":
            return "xodim"
        if c["role"] == "filial":
            return "filial · %s" % (norm(c["filial"]) or "—")
        return ", ".join(self.firms_of(c["chat_id"])) or "—"

    def find_filial(self, text):
        """Ko'chirish varag'idagi filiallardan qidiradi (norm, qism bo'yicha). (topilgan yoki None, nomzodlar)."""
        q = norm(text)
        names = [norm(n) for n in self.data.koch_filials()]
        if q in ("SKLAD", "СКЛАД") and self.sklad in names:
            return self.sklad, [self.sklad]
        if q in names:
            return q, [q]
        found = [n for n in names if q and q in n]
        return (found[0] if len(found) == 1 else None), found

    # ---------------- faollik, o'chirish, yetkazib bo'lmaganlar ----------------
    def touch(self, cid, action):
        """Faollik statistikasi uchun: kim, qachon, qaysi bo'lim. Botga qaytgan odamning «tark etgan» belgisi olinadi."""
        ts = now().isoformat(timespec="seconds")
        with self.dblock:
            self.db.execute("INSERT INTO activity(chat_id, ts, action) VALUES(?,?,?)", (cid, ts, action))
            self.db.execute("UPDATE clients SET last_seen=?, left_at=NULL WHERE chat_id=?", (ts, cid))
            self.db.commit()

    def mark_gone(self, cid, err):
        """Foydalanuvchi botni bloklagan bo'lsa, keyingi yuborishlarda u o'tkazib yuboriladi."""
        if not (isinstance(err, TGError) and err.gone):
            return False
        with self.dblock:
            self.db.execute("UPDATE clients SET left_at=? WHERE chat_id=? AND left_at IS NULL",
                            (now().isoformat(timespec="seconds"), cid))
            self.db.commit()
        logging.info("Foydalanuvchi botni tark etgan (bloklagan): %s", cid)
        return True

    def delete_client(self, cid):
        """Mijozni va unga tegishli hamma yozuvlarni o'chiradi."""
        with self.dblock:
            for t in ("clients", "client_firms", "client_accounts", "acc_requests"):
                self.db.execute("DELETE FROM %s WHERE chat_id=?" % t, (cid,))
            self.db.execute("DELETE FROM settings WHERE key IN (?, ?, ?, ?)", ("kbsig_%d" % cid, "acc_req_%d" % cid,
                                                                             "panel_%d" % cid, "kbclean_%d" % cid))
            self.db.commit()
        for d in (self.state, self.fmode, self.fsearch, self.search, self.pending_text, self._panels, self._panel_last, self.fret):
            d.pop(cid, None)
        self._kbclean.discard(cid)
        self.sqall.discard(cid)

    # ---------------- texnik rejim ----------------
    MAINT_TEXT = ("🛠 <b>Botda texnik ishlar olib borilmoqda.</b>\n"
                  "Tez orada qayta ishga tushadi. Ishlar tugagach, sizga shu yerga xabar beramiz.")

    @property
    def maint_on(self):
        return self.setting("maint_on", "0") == "1"

    @property
    def maint_text(self):
        return self.setting("maint_text") or self.MAINT_TEXT

    def maint_waiting(self):
        try:
            return json.loads(self.setting("maint_waiting", "[]"))
        except ValueError:
            return []

    def _maint_wait_add(self, cid):
        with self.dblock:
            w = self.maint_waiting()
            if cid not in w:
                self.set_setting("maint_waiting", json.dumps(w + [cid]))

    def set_maint(self, on, text=None, by=None):
        """Texnik rejim: mijozlar botdan foydalana olmaydi (adminlar ishlayveradi), avto-yuborish to'xtaydi.
        O'chirilganda shu vaqt ichida yozganlarga «bot yana ishlayapti» xabari va paneli yuboriladi."""
        if text is not None:
            self.set_setting("maint_text", text.strip())
        if on == self.maint_on:
            return 0
        if on:
            self.set_setting("maint_waiting", "[]")
            self.set_setting("maint_since", dt_text(now()))
            self.set_setting("maint_on", "1")
            logging.info("Texnik rejim YOQILDI (%s)", by or "panel")
            return 0
        self.set_setting("maint_on", "0")
        waiting = self.maint_waiting()
        self.set_setting("maint_waiting", "[]")
        logging.info("Texnik rejim o'chirildi (%s). Kutganlar: %d", by or "panel", len(waiting))
        self.bg(self._maint_notify, waiting, name="maint")
        return len(waiting)

    def _maint_notify(self, ids):
        for cid in ids:
            c = self.client(cid)
            if not c:
                continue
            try:
                if c["status"] == "active" and self.can_use(c):
                    self.panel_resend(cid, *self.client_home(c), note="✅ <b>Texnik ishlar tugadi, bot yana ishlayapti.</b>")
                else:
                    self.tg.send(cid, "✅ Texnik ishlar tugadi, bot yana ishlayapti.")
            except Exception as e:  # noqa
                if not self.mark_gone(cid, e):
                    logging.exception("Texnik rejimdan keyin xabar yuborib bo'lmadi: %s", cid)
            self.sleep(0.05)

    def bg(self, fn, *args, name="ish"):
        """Uzoq davom etadigan ishni (ommaviy yuborish) fonda bajaradi: bot shu vaqtda ham javob beradi."""
        def run():
            try:
                fn(*args)
            except Exception:  # noqa
                logging.exception("Fon ishida xato: %s", name)
            finally:  # shu oqim ochgan baza ulanishi yopiladi
                con = getattr(self._local, "con", None)
                if con is not None:
                    self._local.con = None
                    con.close()
        t = threading.Thread(target=run, daemon=True, name=name)
        t.start()
        return t

    # ---------------- hisobotlar ----------------
    KINDS = {"srok": "⏰ Srok", "old": "📦 30+ kun", "need": "🧾 Buyurtma", "stock": "📊 Qoldiq va sotuv",
             "acc": "💰 Hisob", "narx": "💲 Narxlar"}

    def prepare(self, kind, tags, accounts=None):
        """(bytes, fayl nomi, qatorlar soni, sarlavha) yoki None."""
        d = self.data
        today = datetime.date.today().isoformat()
        if kind == "srok":
            rows = d.srok_rows(tags, self.srok_days)
            if not rows:
                return None
            exp = sum(1 for r in rows if r["kun"] < 0)
            title = "⏰ <b>Srogi %d kundan kam tovarlar</b>\nPartiyalar: %d" % (self.srok_days, len(rows))
            if exp:
                title += " (shundan muddati o'tgan: %d)" % exp
            return build_srok(rows), "Srok_%s.xlsx" % today, len(rows), title
        if kind == "old":
            rows = d.old_rows(tags)
            if not rows:
                return None
            content = build_old(rows) if tags is not None else build_old_admin(rows)
            return content, "30kun_%s.xlsx" % today, len(rows), "📦 <b>30 kundan ortiq turgan tovarlar</b>\nQatorlar: %d" % len(rows)
        if kind == "need":
            rows = self.need_rows(tags)
            if not rows:
                return None
            return (build_need(rows, d.period_days), "Buyurtma_%s.xlsx" % today, len(rows),
                    "🧾 <b>Buyurtma tavsiyasi</b>\nTovarlar: %d" % len(rows))
        if kind == "stock":
            rows = d.stock_rows(tags)
            if not rows:
                return None
            return (build_stock(rows, d.period_days), "Qoldiq_%s.xlsx" % today, len(rows),
                    "📊 <b>Qoldiq va sotuv</b> (%d kun)\nTovarlar: %d" % (d.period_days, len(rows)))
        if kind == "narx":
            rows = sorted(d._by_tags(d.narx, tags), key=lambda r: norm(r["name"]))
            if not rows:
                return None
            return build_narx(rows), "Narx_%s.xlsx" % today, len(rows), "💲 <b>Narxlar</b>\nTovarlar: %d" % len(rows)
        if kind == "koch":
            rows = d.koch_rows(tags)
            if not rows:
                return None
            xm = self.koch_abc_excel()
            return (build_koch(rows, lambda r: xm.get(self.koch_key(r), "")), "Kochirish_%s.xlsx" % today, len(rows),
                    "🔁 <b>Filiallararo ko'chirish</b>\nQatorlar: %d · %s qadoq · %s so'm" % (
                        len(rows), qty(sum(r["qty"] for r in rows)), money(sum(r["summa"] for r in rows))))
        if kind == "acc":
            rows = d.hisob if accounts is None else d.account_rows(accounts)
            if not rows:
                return None
            return build_hisob(rows), "Hisob_%s.xlsx" % today, len(rows), "💰 <b>Hisob (debet / kredit)</b>"
        return None

    def send_report(self, chat, kind, tags, accounts=None, quiet=False, mid=None, panel=False):
        """Excel hisobot. panel=True (foydalanuvchi tugma bosdi): fayl alohida xabar, panel uning ostiga qayta tushadi;
        ma'lumot yo'q bo'lsa — panelning o'zida izoh. panel=False (ommaviy / yakka yuborish): oddiy xabar."""
        self.data.reload()
        if not self.data.path:
            if panel:
                self.panel_note(chat, mid, "⚠ Ma'lumot fayli topilmadi. Keyinroq urinib ko'ring.")
            else:
                self.tg.send(chat, "Ma'lumot fayli topilmadi. Keyinroq urinib ko'ring.")
            return False
        res = self.prepare(kind, tags, accounts)
        if not res:
            if panel:
                self.panel_note(chat, mid, "➖ Bu bo'lim bo'yicha ma'lumot topilmadi.")
            elif not quiet:
                self.tg.send(chat, "Bu bo'lim bo'yicha ma'lumot topilmadi.")
            return False
        content, fname, n, title = res
        caption = cut_html("%s\n🕒 Ma'lumot: %s" % (title, dt_text(self.data.when())), 1000)
        if panel:
            self.panel_doc(chat, fname, content, caption)
        else:
            self.tg.send_document(chat, fname, content, caption)
        return True

    # ---------------- panellar ----------------
    # Bitta «tirik» panel xabari (dizayn/bot/NAVIGATSIYA.md). Har rolning bosh menyusi — *_root_rows(): yangi bo'lim
    # qo'shish uchun shu ro'yxatga qator qo'shing. Bo'lim ekrani — screen() (yo'l qatori + matn + tugmalar + navigatsiya).
    def client_root_rows(self, c):
        """Firma vakili / xodim bosh menyusi: [[(matn, callback)]]. Sonlar tugmalarda. Xodimda oxirgi qator — firma filtri."""
        d = self.data
        d.reload()
        tags, _ = self.eff(c)
        n_srok = len(d.srok_rows(tags, self.srok_days))
        n_old = len(d.old_rows(tags))
        n_need = len(self.oreq_low(tags)) if c["role"] == "firm" else len(self.need_rows(tags))
        hisob = "💰 Hisob"
        if c["role"] != "staff":
            rows, net = self.hisob_net(tags, self.accounts_of(c["chat_id"]), auto=False)
        elif tags is not None:
            rows, net = self.hisob_net(tags, None)
        else:
            rows, net = [], 0
        if rows and net:
            hisob = "💰 Hisob · %s %s" % (money(abs(net)), "kredit" if net < 0 else "debet")
        out = [[(hisob, "c:acc")], [("⏰ Srok · %d" % n_srok, "c:srok"), ("📦 30+ kun · %d" % n_old, "c:old")],
               [("🧾 Buyurtma · %d" % n_need, "c:need"), ("📊 Sotuv va qoldiq", "c:stock")], [("🔍 Dori qidirish", "c:find")]]
        if c["role"] == "staff":
            out.append([(self.firm_btn(c["chat_id"]), "a:fm")])
        return out

    def client_home_text(self, c):
        d = self.data
        d.reload()
        tags, flabel = self.eff(c)
        label = (flabel or "Xodim (hamma firmalar)") if c["role"] == "staff" else ", ".join(sorted(tags))[:80]
        n_srok = len(d.srok_rows(tags, self.srok_days))
        n_old = len(d.old_rows(tags))
        n_need = len(self.oreq_low(tags)) if c["role"] == "firm" else len(self.need_rows(tags))
        text = "🏠 <b>Bosh menyu</b>\n🏢 <b>%s</b>\n\n⏰ Srogi yaqin: <b>%d</b> ta\n📦 30+ kun: <b>%d</b> ta\n🧾 Buyurtma kerak: <b>%d</b> ta" % (
            esc(label), n_srok, n_old, n_need)
        if c["role"] != "staff":
            rows, net = self.hisob_net(tags, self.accounts_of(c["chat_id"]), auto=False)
            if rows and net:
                text += "\n💰 Hisob: <b>%s</b> %s" % (money(abs(net)), "kredit" if net < 0 else "debet")
        return text + ("\n\n🕒 Ma'lumot: %s\nBo'limni tanlang 👇\n"
                       "🔍 Dori nomini yozsangiz, qoldig'i, narxi va srokini ko'rsataman.") % dt_text(d.when())

    def client_panel(self, c, save_sig=True):
        """(matn, tugmalar pastki klaviatura ko'rinishida) — server paneli «mijoz ko'rinishi» va admin ko'rinishi uchun.
        Botda panel inline yuboriladi: client_home(). save_sig — eski chaqiruvlar uchun (endi ishlatilmaydi)."""
        if c["role"] == "filial":
            return self.filial_panel(c, save_sig)
        return self.client_home_text(c), self.client_kb(c)

    def client_home(self, c):
        """Mijoz bosh menyusi (inline panel)."""
        if c["role"] == "filial":
            return self.filial_home(c)
        return self.client_home_text(c), inline_kb(self.client_root_rows(c))

    def home_kb(self, admin=False):
        """Bitta tugma: panelni ochish (server paneli yangi adminga xabar yuborganda)."""
        return inline_kb([[("⚙️ Admin panel", "a:home")]] if admin else [[("🏠 Menyu", "c:home")]])

    def admin_root_rows(self, cid, n):
        """Admin bosh menyusi: [[(matn, callback)]]. n — matndagi sonlar (pending, srok, old). Yangi bo'lim — shu yerga."""
        out = []
        if n["pending"]:
            out.append([("⏳ Kutayotgan so'rovlar · %d" % n["pending"], "ul:pending:0")])
        out.append([("🧾 Buyurtma so'rovlari · %d" % n.get("oreq", 0), "oq:l")])
        out += [[("⏰ Srok · %d" % n["srok"], "a:sec:srok"), ("📦 30+ kun · %d" % n["old"], "a:sec:old")],
                [("🧾 Buyurtma", "a:sec:need"), ("💰 Hisob", "a:sec:acc")],
                [("📊 Sotuv va qoldiq", "a:sec:stock"), (self.firm_btn(cid), "a:fm")],
                [("📊 Hisobotlar (Excel)", "a:rep"), ("🔁 Ko'chirishlar", "a:tr")],
                [("👥 Foydalanuvchilar", "a:usr"), ("📢 Xabar yuborish", "a:bc")],
                [("🏭 Sklad", "a:sk"), ("📈 Statistika", "a:stat"), ("📜 Jurnal", "a:log")],
                [("⚙️ Sozlamalar", "a:set"), ("🔄 Yangilash", "a:ref")],
                [("✅ Texnik rejimni o'chirish" if self.maint_on else "🛠 Texnik rejim (botni uzish)", "a:mt")]]
        return out

    def admin_panel(self, cid=None):
        d = self.data
        d.reload()
        n_users = self.db.execute("SELECT COUNT(*) FROM clients WHERE status='active'").fetchone()[0]
        n = {"pending": self.db.execute("SELECT COUNT(*) FROM clients WHERE status='pending'").fetchone()[0],
             "srok": len(d.srok_rows(None, self.srok_days)), "old": len(d.old),
             "oreq": self.db.execute("SELECT COUNT(*) FROM order_req WHERE status=?", (BY.OPEN,)).fetchone()[0]}
        text = ("<b>⚙️ Admin panel</b>\n\n"
                "⏰ Srogi yaqin (%d kun): <b>%d</b> ta\n"
                "📦 30+ kun turgan: <b>%d</b> ta\n"
                "👥 Foydalanuvchilar: <b>%d</b> · ⏳ kutmoqda: <b>%d</b>\n%s\n"
                "🕒 Ma'lumot yangilangan: <b>%s</b>") % (
            self.srok_days, n["srok"], n["old"], n_users, n["pending"],
            ("🧾 Buyurtma so'rovlari: <b>%d</b> ta kutmoqda\n" % n["oreq"]) if n["oreq"] else "", dt_text(d.when()))
        if d.error:
            text += "\n⚠ " + esc(d.error)
        elif d.age_hours() > float(self.cfg.get("auto", "stale_hours", fallback="72")):
            text += "\n⚠ Ma'lumot eskirgan, «Yangilash» ni bosing."
        if self.maint_on:
            text = ("🛠 <b>TEXNIK REJIM YOQILGAN</b> (%s dan)\nMijozlar botdan foydalana olmaydi, avto-yuborish to'xtatilgan. "
                    "Kutayotganlar: %d\n━━━━━━━━━━\n") % (self.setting("maint_since", "—"), len(self.maint_waiting())) + text
        tags, label = self.fm(cid)
        if tags is not None:
            text += "\n🏢 Firma: <b>%s</b> (bo'limlar shu firma bo'yicha)" % esc(label)
        text += "\n\n🔍 Dori nomini yozsangiz, qoldig'i, narxi va srokini ko'rsataman."
        if cid:
            text += "\nID: <code>%d</code>" % cid
        return text, inline_kb(self.admin_root_rows(cid, n))

    def maint_action(self, cid, mid, sub):
        """Telegramdan texnik rejimni boshqarish. Yoqishdan oldin tasdiq so'raladi."""
        cr = ("🛠 Texnik rejim",)
        if sub == "on":
            self.set_maint(True, by="tg:%d" % cid)
            return self.panel_show(cid, mid, *self.screen(cr, "🛠 <b>Texnik rejim yoqildi.</b>\n\nMijozlarga shu xabar ko'rinadi:\n\n%s\n\n"
                                   "Siz (admin) botdan odatdagidek foydalanasiz. Tugatgach «✅ O'chirish» ni bosing." % self.maint_text,
                                   [[("✅ O'chirish", "a:mt:off")]], back="a:home", home="a:home"))
        if sub == "off":
            n = self.set_maint(False, by="tg:%d" % cid)
            return self.panel_show(cid, mid, *self.screen(cr, "✅ <b>Texnik rejim o'chirildi, bot ishlayapti.</b>\nShu vaqtda yozgan %d "
                                                          "kishiga xabar yuborilmoqda." % n, back="a:home", home="a:home"))
        if sub == "text":
            return self.ask(cid, mid, cr + ("✏️ Matn",), "Mijozlarga ko'rinadigan yangi matnni yozing 👇\nBekor qilish: /cancel",
                            ("maint_text", None), back="a:home")  # a:mt emas: yoqilgan rejimda u o'chirib qo'yadi
        if self.maint_on:
            return self.maint_action(cid, mid, "off")
        return self.panel_show(cid, mid, *self.screen(cr, "🛠 <b>Texnik rejimni yoqasizmi?</b>\n\nMijozlar botga yozsa, shu javobni oladi:\n\n%s\n\n"
                               "Adminlar ishlayveradi. Avto-yuborish to'xtaydi. O'chirilgach, kutganlarga xabar boradi." % self.maint_text,
                               [[("🛠 Ha, yoqish", "a:mt:on"), ("✏️ Matnni o'zgartirish", "a:mt:text")]], back="a:home", home="a:home"))

    def show(self, cid, mid, text, markup):
        """Xabarni joyida o'zgartiradi (bildirishnomaning o'zida natija), bo'lmasa yangi xabar. Panel uchun — panel_show."""
        if mid:
            try:
                return self.tg.edit(cid, mid, text, markup)
            except TGError as e:
                if "not modified" in str(e):
                    return None
        return self.tg.send(cid, text, markup)

    # ---------------- panel xabari: bitta «tirik» xabar ----------------
    # Tugma bosilsa — shu panel xabari joyida o'zgaradi (panel_show). Matn yozilsa, Excel yoki boshqa alohida xabar
    # chiqsa — panel pastga qayta yuboriladi, eskisi o'chiriladi (panel_resend / panel_again). Panel xabarining id si
    # xotirada va settings da (panel_<chat>) saqlanadi: bot qayta ishga tushsa ham bitta panel qoladi.
    @staticmethod
    def crumb(*parts):
        """Yo'l qatori: «🏠 › ⏰ Srok › A». Ma'lumot (nomlar) bo'laklarini chaqiruvchi esc() qiladi."""
        return " › ".join(("🏠",) + tuple(p for p in parts if p))

    @staticmethod
    def nav_row(back, home, back_text="◀ Orqaga"):
        return [(back_text, back), ("🏠 Menyu", home)]

    def screen(self, crumbs, body, rows=None, back=None, home=None):
        """Bo'lim ekrani → (matn, inline tugmalar): yo'l qatori + matn + qo'shimcha qatorlar + [◀ Orqaga] [🏠 Menyu].
        rows: [[(matn, callback)]] yoki tayyor inline_kb; back=None — navigatsiya qatorisiz (masalan, jarayon)."""
        kb = [list(r) for r in rows.get("inline_keyboard", [])] if isinstance(rows, dict) else inline_kb(rows or [])["inline_keyboard"]
        if back:
            kb += inline_kb([self.nav_row(back, home or back)])["inline_keyboard"]
        return cut_html("%s\n\n%s" % (self.crumb(*crumbs), body)), ({"inline_keyboard": kb} if kb else None)

    def ask(self, cid, mid, crumbs, prompt, state, back, rows=None):
        """Matn kiritishni so'raydigan ekran (son, qidiruv, xabar matni): holat saqlanadi, javob — yangi panel."""
        self.state[cid] = state
        return self.panel_show(cid, mid, *self.screen(crumbs, prompt, rows, back=back, home=self.home_cb(cid)))

    @staticmethod
    def with_note(text, note):
        """Izoh («✅ Berildi», «📎 Excel yuborildi») — yo'l qatoridan keyin (bosh menyuda — boshida)."""
        if not note:
            return text
        if text.startswith("🏠 ›"):
            head, _, rest = text.partition("\n\n")
            return "%s\n\n%s\n\n%s" % (head, note, rest)
        return "%s\n\n%s" % (note, text)

    def home_cb(self, cid):
        """Bosh menyu tugmasining callback'i (rolga qarab)."""
        if cid in self.admins:
            return "a:home"
        return "kh" if self.is_filial(self.client(cid)) else "c:home"

    def home(self, cid):
        """Rolning bosh menyusi (matn, inline tugmalar)."""
        if cid in self.admins:
            return self.admin_panel(cid)
        return self.client_home(self.client(cid))

    def panel_mid(self, cid):
        if cid not in self._panels:
            v = self.setting("panel_%d" % cid) or ""
            self._panels[cid] = int(v) if v.isdigit() else None
        return self._panels[cid]

    def is_panel(self, cid, mid):
        """mid — shu chatning panel xabarimi. Eski xabar yoki bildirishnoma — yo'q (ular panelga aylantirilmaydi)."""
        return bool(mid) and self.panel_mid(cid) == mid

    def _panel_keep(self, cid, mid, text, kb):
        if mid != self.panel_mid(cid):
            self._panels[cid] = mid
            self.set_setting("panel_%d" % cid, mid or "")
        self._panel_last[cid] = (text, kb)

    def kb_clean(self, cid):
        if cid not in self._kbclean and self.setting("kbclean_%d" % cid) == "1":
            self._kbclean.add(cid)
        return cid in self._kbclean

    def kb_cleaned(self, cid):
        """Eski pastki klaviatura olib tashlangan deb belgilaydi (masalan, remove_keyboard bilan javob ketganda)."""
        if not self.kb_clean(cid):
            self._kbclean.add(cid)
            self.set_setting("kbclean_%d" % cid, "1")

    def kb_cleanup(self, cid):
        """Avvalgi versiyaning pastki klaviaturasi bo'lgan foydalanuvchiga — bir marta remove_keyboard xabari."""
        if self.kb_clean(cid):
            return
        self.kb_cleaned(cid)
        try:
            self.tg.send(cid, "🔄 Menyu yangilandi: bo'limlar endi xabar ostidagi tugmalarda.", {"remove_keyboard": True})
        except Exception as e:  # noqa
            if not self.mark_gone(cid, e):
                logging.warning("Pastki klaviaturani olib tashlab bo'lmadi (%s): %s", cid, e)

    def panel_show(self, cid, mid, text, kb, note=None):
        """Tugma bosildi: panel xabari (mid) joyida o'zgaradi (editMessageText). Bosilgan xabar panel bo'lmasa (bildirishnoma,
        eski xabar) yoki mid yo'q bo'lsa — panel pastga qayta yuboriladi: chatda doim bitta panel."""
        if self.is_panel(cid, mid) and self.kb_clean(cid):
            try:
                r = self.tg.edit(cid, mid, self.with_note(text, note), kb)
            except TGError as e:
                if "not modified" not in str(e):  # xabar o'chirilgan / juda eski — yangisi pastga
                    return self.panel_resend(cid, text, kb, note)
                r = None
            self._panel_keep(cid, mid, text, kb)
            return r
        return self.panel_resend(cid, text, kb, note)

    def panel_resend(self, cid, text, kb, note=None):
        """Yangi panel pastga (yozilgan matnga javob, Excel / bildirishnomadan keyin): yangisi yuboriladi, eskisi o'chiriladi."""
        self.kb_cleanup(cid)
        old = self.panel_mid(cid)
        r = self.tg.send(cid, self.with_note(text, note), kb)
        mid = r.get("message_id") if isinstance(r, dict) else None
        if old and old != mid:
            self.tg.delete(cid, old)
        self._panel_keep(cid, mid, text, kb)
        return r

    def panel_last(self, cid):
        return self._panel_last.get(cid) or self.home(cid)

    def panel_again(self, cid, note=None):
        """Alohida xabardan (Excel fayl) keyin: panelning oxirgi holati shu xabar ostiga qayta tushadi."""
        return self.panel_resend(cid, *self.panel_last(cid), note=note)

    def panel_note(self, cid, mid, note):
        """Yangi xabar chiqmadi (ma'lumot yo'q, kuting...): panelning joriy holati izoh bilan."""
        return self.panel_show(cid, mid, *self.panel_last(cid), note=note)

    def panel_doc(self, cid, fname, content, caption):
        """Excel — alohida xabar (hujjat, izohi o'zgarmaydi), keyin panel fayl ostiga qayta tushadi."""
        r = self.tg.send_document(cid, fname, content, caption)
        self.panel_again(cid, "📎 Excel yuborildi (yuqorida).")
        return r

    def push_done(self, cid, mid, note):
        """Bildirishnomadagi amal bajarildi (yangi so'rov, hisob so'rovi): o'sha xabar bir qatorli natijaga aylanadi va
        None qaytadi (panel pastga tushadi). Panelning o'zida bosilgan bo'lsa — mid (panel joyida o'zgaradi)."""
        if mid and not self.is_panel(cid, mid):
            try:
                self.tg.edit(cid, mid, note, None)
            except Exception:  # noqa
                pass
            return None
        return mid

    # ---------------- yangilanishlarni qayta ishlash ----------------
    def handle(self, upd):
        if "callback_query" in upd:
            return self.on_callback(upd["callback_query"])
        if "message" in upd:
            return self.on_message(upd["message"])

    def on_message(self, m):
        cid = m["chat"]["id"]
        if m["chat"].get("type") != "private":
            return
        text = (m.get("text") or "").strip()
        text = legacy_button(text)  # eski kirill pastki tugma → koddagi nomi
        if text == "/myid":
            return self.tg.send(cid, "Sizning Telegram ID: <b>%d</b>" % cid)
        if self.maint_on and cid not in self.admins:
            self._maint_wait_add(cid)
            return self.tg.send(cid, self.maint_text)
        if cid in self.admins:
            return self.admin_message(cid, m, text)
        return self.client_message(cid, m, text)

    # ---------------- mijoz ----------------
    def client_message(self, cid, m, text):
        c = self.client(cid)
        contact = m.get("contact")
        if contact:
            if contact.get("user_id") != m["from"]["id"]:
                return self.tg.send(cid, "Iltimos, faqat o'z kontaktingizni yuboring (tugma orqali).")
            if not c:
                name = " ".join(x for x in (m["from"].get("first_name"), m["from"].get("last_name")) if x)
                with self.dblock:
                    self.db.execute("INSERT INTO clients(chat_id,phone,name,username,status,role,created) VALUES(?,?,?,?,?,?,?)",
                                    (cid, contact["phone_number"], name, m["from"].get("username") or "", "pending", "firm",
                                     now().isoformat(timespec="seconds")))
                    self.db.commit()
                if self.apply_invite(cid):
                    return None
                self.notify_admins_new(cid)
            elif c["status"] == "active" and self.can_use(c):  # allaqachon qabul qilingan: panelni qaytaramiz
                return self.panel_resend(cid, *self.client_home(c))
            elif c["status"] == "blocked":
                return self.tg.send(cid, "Kontaktingiz bloklangan, iltimos adminga murojaat qiling.")
            self.kb_cleaned(cid)  # «📱 Kontaktni yuborish» klaviaturasi shu xabar bilan olinadi
            return self.tg.send(cid, "Kontaktingiz ko'rib chiqilmoqda, iltimos kuting.", {"remove_keyboard": True})
        if not c:  # yagona pastki klaviatura: Telegram kontaktni faqat shu tugma orqali beradi
            return self.tg.send(cid, "Assalomu alaykum! Ro'yxatdan o'tish uchun kontaktingizni yuboring.",
                                reply_kb([[{"text": "📱 Kontaktni yuborish", "request_contact": True}]], one_time_keyboard=True))
        if c["status"] == "pending":
            return self.tg.send(cid, "Kontaktingiz ko'rib chiqilmoqda, iltimos kuting.")
        if c["status"] == "blocked":
            return self.tg.send(cid, "Kontaktingiz bloklangan, iltimos adminga murojaat qiling.")
        if not self.can_use(c):
            return self.tg.send(cid, "Sizga hali %s biriktirilmagan. Iltimos adminga murojaat qiling." % (
                "filial" if c["role"] == "filial" else "firma"))
        if c["role"] == "filial":  # filial xodimi: faqat ko'chirish bo'limlari
            return self.filial_message(cid, c, text)
        if c["role"] == "staff":
            if text.startswith("🏢 Firma"):  # eski pastki tugma: firma tanlash darajasi
                return self.firm_prompt(cid, None)
            st = self.state.get(cid)
            if st and st[0] == "firm_pick" and text and not text.startswith("/") and not self.client_button(text):
                return self.firm_pick_input(cid, text)
        st = self.state.get(cid)
        if st and st[0] == "ord_qty" and c["role"] == "firm" and text and not text.startswith("/"):
            return self.ord_qty_input(cid, c, st, text)
        if text == "/cancel":
            self.state.pop(cid, None)
        if text in ("🏠 Menyu", "⚙️ Admin panel"):  # eski «🏠 Menyu» tugmasi
            self.state.pop(cid, None)
            return self.panel_resend(cid, *self.client_home(c))
        kind = self.client_button(text)
        if kind:  # eski pastki tugma (kirill/lotin): shu bo'lim yangi panel bo'lib ochiladi
            self.state.pop(cid, None)
            self.touch(cid, kind)
            return self.client_section(cid, c, kind)
        if len(text) >= 3 and not text.startswith("/"):  # dori nomi: qidiruv — javob yangi panel
            self.state.pop(cid, None)
            self.touch(cid, "search")
            tags, label = self.eff(c)
            return self.panel_resend(cid, *self.screen(("🔍 Dori qidirish", "«%s»" % esc(self.short(text.strip(), 24))),
                                                       self.search_text(text, tags, label, firm_view=c["role"] != "staff"),
                                                       [self.nav_row("c:find", "c:home", "🔍 Boshqa dori")]))
        self.touch(cid, "start")
        self.panel_resend(cid, *self.client_home(c))

    def search_prompt(self, cid, mid, c):
        """🔍 Dori qidirish: panel dori nomini so'raydi (keyin yozilgan matn — qidiruv, natija yangi panel)."""
        self.state.pop(cid, None)
        label = self.fm(cid)[1] if c["role"] == "staff" else ", ".join(sorted(self.tags_for(c)))[:60]
        body = ("🔍 <b>Dori qidirish</b>%s\nDori nomini yozing 👇 (kamida 3 harf)\n"
                "Qoldiq (filiallar bo'yicha), narx va eng yaqin srokni ko'rsataman.") % (" · 🏢 %s" % esc(label) if label else "")
        return self.panel_show(cid, mid, *self.screen(("🔍 Dori qidirish",), body, back="c:home", home="c:home"))

    def apply_invite(self, cid):
        """Admin oldindan qo'shgan telefon raqami bilan kelgan mijozni avtomatik qabul qiladi."""
        c = self.client(cid)
        d9 = digits(c["phone"])[-9:] if c else ""
        inv = self.db.execute("SELECT * FROM invites WHERE phone9=?", (d9,)).fetchone() if len(d9) == 9 else None
        if not inv:
            return False
        with self.dblock:
            if inv["name"]:
                self.db.execute("UPDATE clients SET name=? WHERE chat_id=?", (inv["name"], cid))
            self.db.execute("UPDATE clients SET role=?, filial=? WHERE chat_id=?",
                            (inv["role"], norm(inv["filial"]) if inv["role"] == "filial" and inv["filial"] else None, cid))
            self.db.execute("DELETE FROM invites WHERE phone9=?", (d9,))
            self.db.commit()
        self.set_firms(cid, [x for x in (inv["firms"] or "").split(";") if x.strip()])
        with self.dblock:
            for a in (inv["accounts"] or "").split(";"):
                if a.strip():
                    self.db.execute("INSERT OR IGNORE INTO client_accounts VALUES(?,?)", (cid, norm(a)))
            self.db.commit()
        c = self.client(cid)
        if not self.can_use(c):
            return False  # firma ham, xodim ham emas: admin qo'lda ko'rib chiqadi
        self.activate(cid)
        for a in self.admins:
            try:
                self.tg.send(a, "✅ <b>Avto-qabul</b> (oldindan qo'shilgan)\n%s · %s" % (esc(c["name"]), esc(c["phone"])),
                             inline_kb([[("👤 Kartochka", "uc:%d" % cid), ("🏠 Menyu", "a:home")]]))
            except Exception:  # noqa
                pass
        return True

    def request_account(self, cid, c):
        """Hisobi biriktirilmagan mijoz «Hisob» ni bosdi: adminlarga biriktirish tugmasi. Mijozga ko'rsatiladigan matn qaytadi."""
        msg = "💰 Sizga hisob hali biriktirilmagan. So'rovingiz adminga yuborildi, biriktirilgach shu yerdan ko'rasiz."
        last = float(self.setting("acc_req_%d" % cid, "0"))
        if time.time() - last < 3600:  # bir soatda bir martadan ko'p admin'ni bezovta qilmaymiz
            return msg
        self.set_setting("acc_req_%d" % cid, time.time())
        for a in self.admins:
            try:
                self.tg.send(a, "💰 <b>Hisob so'rovi</b>\n%s · %s\nFirma: %s" % (
                    esc(c["name"]), esc(c["phone"]), esc(", ".join(self.firms_of(cid)) or "—")),
                    inline_kb([[("💰 Biriktirish", "ac:%d" % cid)], [("🏠 Menyu", "a:home")]]))
            except Exception:  # noqa
                logging.exception("Adminga xabar yuborib bo'lmadi: %s", a)
        return msg

    def notify_admins_new(self, cid):
        c = self.client(cid)
        sug = self.suggest_tags(c["phone"])
        text = "🆕 <b>Yangi so'rov</b>\nIsm: %s\nTelefon: %s\nID: %d" % (esc(c["name"]), esc(c["phone"]), cid)
        rows = []
        if sug:
            text += "\nTaklif etilgan firma: " + esc(", ".join(sug))
            rows.append([(("✅ Qabul (%s)" % ", ".join(sug))[:60], "ap:%d" % cid)])
        rows.append([("✅ Qabul qilish", "ok:%d" % cid), ("👔 Xodim", "st:%d" % cid)])
        rows.append([("🏬 Filial xodimi", "fl:%d" % cid), ("⛔ Blok", "bl:%d" % cid)])
        rows.append([("🏠 Menyu", "a:home")])
        for a in self.admins:
            try:
                self.tg.send(a, text, inline_kb(rows))
            except Exception:
                logging.exception("Adminga xabar yuborib bo'lmadi: %s", a)

    def suggest_tags(self, phone):
        d9 = digits(phone)[-9:]
        if len(d9) < 9:
            return []
        self.data.reload()
        return [t for t in self.data.all_tags() if len(digits(t)) >= 9 and digits(t)[-9:] == d9]

    def activate(self, target, notify=True):
        self.set_status(target, "active")
        if notify:
            try:  # bildirishnoma (alohida xabar) + bitta tugma: bosilganda panel shu xabar ostida ochiladi
                c = self.client(target)
                if self.is_filial(c):
                    who = "🏬 <b>%s</b> · filial xodimi" % esc(norm(c["filial"]))
                elif c["role"] == "staff":
                    who = "👔 <b>Xodim</b> (hamma firmalar)"
                else:
                    who = "🏢 Firma: <b>%s</b>" % esc(", ".join(self.firms_of(target)) or "—")
                self.tg.send(target, "✅ <b>Siz qabul qilindingiz!</b>\n%s\n\nHamma bo'lim bitta menyuda 👇" % who,
                             inline_kb([[("🏠 Menyuni ochish", "kh" if self.is_filial(c) else "c:home")]]))
                self.ensure_acc_requests(target)
            except Exception:
                logging.exception("Mijozga xabar yuborib bo'lmadi: %s", target)

    def on_client_callback(self, cid, mid, data):
        """Firma vakili / xodim panelining tugmalari (c:<bo'lim>, c:home, c:find)."""
        c = self.client(cid)
        if not c or c["status"] != "active" or not self.can_use(c) or c["role"] == "filial":
            return
        self.state.pop(cid, None)
        act = data.split(":", 1)[1]
        if act in ("srok", "old", "need", "stock", "acc"):
            self.touch(cid, act)
            return self.client_section(cid, c, act, mid)
        if act.startswith("sq"):  # 📊 Sotuv va qoldiq: c:sq:<tartib>:<sahifa> | c:sqa (xodim: hamma firmalar)
            return self.sq_action(cid, mid, act.split(":"))
        if act == "b" or act.startswith("b:"):  # 🧾 Buyurtma so'rovi (faqat firma vakili)
            return self.on_order_client(cid, c, mid, act.split(":")[1:])
        if act == "find":
            return self.search_prompt(cid, mid, c)
        if act == "home":
            return self.panel_show(cid, mid, *self.client_home(c))

    # ---------------- admin ----------------
    def admin_message(self, cid, m, text):
        st = self.state.get(cid)
        if text == "/cancel":
            self.state.pop(cid, None)
            return self.send_admin_panel(cid)
        key = self.admin_button(text)
        if key:  # eski pastki menyu tugmasi: kutilayotgan kiritishni bekor qiladi, shu daraja yangi panel bo'lib ochiladi
            self.state.pop(cid, None)
            return self.admin_press(cid, key)
        if st and text and not text.startswith("/"):
            return self.admin_state_input(cid, st, text, m)
        if text in ("/start", "/menu", "/admin", "/help", ""):
            return self.send_admin_panel(cid)
        if text in ("/clients",):
            return self.users_list(cid, None, "all", 0)
        if text in ("/pending",):
            return self.users_list(cid, None, "pending", 0)
        if text in ("/stat", "/stats"):
            return self.admin_menu_action(cid, None, ["stat"])
        if text in ("/kochirish", "/koch"):
            return self.koch_admin_action(cid, None, "tr", [])
        mm = re.match(r"^/c(-?\d+)$", text)
        if mm:
            return self.client_card(cid, int(mm.group(1)))
        if len(text) >= 3 and not text.startswith("/"):  # dori qidirish (tanlangan firma bo'yicha) — javob yangi panel
            tags, label = self.fm(cid)
            return self.panel_resend(cid, *self.screen(("🔍 «%s»" % esc(self.short(text.strip(), 24)),),
                                                       self.search_text(text, tags, label), back="a:home", home="a:home"))
        return self.send_admin_panel(cid)

    def send_admin_panel(self, cid, mid=None, note=None):
        """Admin bosh menyusi: tugma bosilganda (mid) joyida, aks holda yangi panel pastga."""
        return self.panel_show(cid, mid, *self.admin_panel(cid), note=note)

    def a_screen(self, cid, mid, crumbs, body, rows=None, back="a:home", note=None):
        """Admin bo'limi ekrani (panelning o'zida)."""
        return self.panel_show(cid, mid, *self.screen(crumbs, body, rows, back=back, home="a:home"), note=note)

    def on_callback(self, cq):
        cid = cq["from"]["id"]
        data = cq.get("data", "")
        mid = (cq.get("message") or {}).get("message_id")
        if cid in self.admins and cq["from"].get("first_name"):
            self._anames[cid] = " ".join(x for x in (cq["from"].get("first_name"), cq["from"].get("last_name")) if x)[:40]
        maint = self.maint_on and cid not in self.admins
        try:
            p = {"callback_query_id": cq["id"]}
            if maint:  # tugma bosilganda oyna chiqadi: texnik ishlar
                p.update(text=html.unescape(re.sub(r"<[^>]+>", "", self.maint_text))[:190], show_alert=True)
            self.tg.call("answerCallbackQuery", p)
        except Exception:  # noqa
            pass
        if maint:
            self._maint_wait_add(cid)
            return
        if data.startswith("c:"):
            return self.on_client_callback(cid, mid, data)
        if data.startswith("x:"):
            return self.on_excel_callback(cid, data[2:], mid)
        if data.split(":", 1)[0] in self.FIL_CB:  # filial xodimi tugmalari (admin/staff/firma uchun emas)
            return self.on_filial_callback(cid, mid, data)
        if cid not in self.admins and (data == "a:fm" or data.startswith("a:fm:") or data.startswith("fp:")):
            c = self.client(cid)
            if c and c["role"] == "staff" and c["status"] == "active":  # xodim: firma filtri (ichma-ich daraja)
                if data.startswith("a:fm"):
                    return self.firm_prompt(cid, mid, (data.split(":") + [None])[2])
                return self.firm_pick_cb(cid, mid, data.split(":") + [""])
            return
        if cid not in self.admins:
            return
        return self.on_admin_callback(cid, mid, data)

    def on_admin_callback(self, cid, mid, data):
        parts = data.split(":")
        act = parts[0]
        if act not in ("at", "fp") and not (act == "ul" and parts[-1] == "p"):
            self.state.pop(cid, None)  # boshqa daraja ochildi: kutilayotgan matn kiritish bekor (kerak bo'lsa, ekran qayta so'raydi)
        if act == "a":
            return self.admin_menu_action(cid, mid, parts[1:])
        if act == "oq":  # 🧾 buyurtma so'rovlari (kartalar)
            try:
                return self.on_order_admin(cid, mid, parts)
            except (ValueError, IndexError):  # buzilgan / eski tugma
                return None
        if act == "pg":
            return self.list_clients(cid, int(parts[1]), None if parts[2] == "all" else parts[2], mid)
        if act in ("ak", "ar"):
            res = self.acc_resolve(int(parts[1]), act == "ak")
            note = {True: "✅ Tasdiqlandi" if act == "ak" else "❌ Rad etildi", None: "Bu so'rov allaqachon hal qilingan."}[res]
            return self.acc_requests_list(cid, self.push_done(cid, mid, note), note)
        if act == "ul" and parts[1] == "accreq":
            return self.acc_requests_list(cid, mid)
        if act == "ul":
            return self.users_list(cid, mid, parts[1], int(parts[2]), pick=len(parts) > 3 and parts[3] == "p")
        if act == "uc":
            return self.client_card(cid, int(parts[1]), mid)
        if act == "us":
            return self.ask(cid, mid, ("👥 Foydalanuvchilar", "🔍 Qidirish"),
                            "🔍 Ism, telefon, firma belgisi yoki ID yozing 👇\nBekor qilish: /cancel", ("user_search", None), back="a:usr")
        if act == "ua":
            return self.ask(cid, mid, ("👥 Foydalanuvchilar", "➕ Qo'shish"),
                            "➕ <b>Foydalanuvchi qo'shish</b>\nTelefon va firma belgisini yozing:\n"
                            "<code>+998901234567 АБДУВОХИД ТОШКЕНТ 94-5560909</code>\n"
                            "Xodim bo'lsa: <code>+998901234567 xodim</code>\n"
                            "Filial xodimi bo'lsa: <code>+998901234567 filial ЧИНОЗ</code>\n"
                            "Foydalanuvchi botga kirib kontaktini yuborganda avtomatik qabul qilinadi.\nBekor qilish: /cancel",
                            ("add_user", None), back="a:usr")
        if act == "ud":
            with self.dblock:
                self.db.execute("DELETE FROM invites WHERE phone9=?", (parts[1],))
                self.db.commit()
            return self.users_list(cid, mid, "inv", 0)
        if act == "fp":
            return self.firm_pick_cb(cid, mid, parts + [""])
        if act == "m1":
            return self.one_menu(cid, int(parts[1]), mid)
        if act == "m2":
            return self.one_send(cid, int(parts[1]), parts[2], mid)
        if act == "at":
            target = int(parts[1])
            name = self._pick(self.search.get(cid, []), parts[2], (parts + [""])[3])
            if not name:  # boshqa qidiruvdagi eski tugma — boshqa tashkilot hisobi biriktirilib qolmasin
                return self.panel_note(cid, mid, "⚠ Ro'yxat eskirgan: oxirgi qidiruv natijasidagi tugmani bosing.")
            with self.dblock:
                self.db.execute("INSERT OR IGNORE INTO client_accounts VALUES(?,?)", (target, norm(name)))
                self.db.commit()
            return self.client_card(cid, target, mid, "💰 Qo'shildi: %s\nYana qo'shish uchun boshqa so'z yozing, tugatish: /cancel" % esc(name))
        target = int(parts[1])
        c = self.client(target)
        if not c:
            return self.panel_note(cid, mid, "Mijoz topilmadi.")
        nm = esc(c["name"])
        uc = "uc:%d" % target
        card = lambda note: self.client_card(cid, target, self.push_done(cid, mid, note), note)
        if act == "ap":
            sug = self.suggest_tags(c["phone"])
            if not sug:
                return self.client_card(cid, target, mid, "Taklif topilmadi, «Qabul qilish» tugmasini ishlating.")
            self.set_firms(target, sug)
            self.activate(target)
            note = "✅ Qabul qilindi: %s → %s" % (nm, esc(", ".join(sug)))
            return self.after_firm(cid, target, note, self.push_done(cid, mid, note))
        if act == "ok":
            if self.can_use(c):
                self.activate(target)
                return card("✅ Faollashtirildi: %s" % nm)
            return self.ask(cid, mid, ("👥", nm, "🏢 Firma"), "%s uchun firma belgisini yozing 👇\nBir nechta bo'lsa, ; bilan ajrating.\n"
                            "Bekor qilish: /cancel" % nm, ("firms", target), back=uc)
        if act == "st":
            new = "firm" if c["role"] == "staff" else "staff"
            self.set_role(target, new)
            if new == "staff" and c["status"] != "blocked":
                self.activate(target)
            return card("%s: %s" % (nm, "xodim sifatida faollashtirildi" if new == "staff" else "oddiy firma (firma belgisini biriktiring)"))
        if act in ("rs", "rf"):  # filial xodimini boshqa rolga qaytarish (filial tozalanadi)
            self.set_role(target, "staff" if act == "rs" else "firm")
            c = self.client(target)
            if act == "rs" or self.can_use(c):
                if c["status"] != "blocked":
                    self.activate(target)
                return card("%s: %s" % (nm, "xodim qilindi" if act == "rs" else "firma qilindi"))
            return self.ask(cid, mid, ("👥", nm, "🏢 Firma"), "%s firma qilindi. Firma belgisini yozing 👇\nBir nechta bo'lsa, ; bilan ajrating.\n"
                            "Bekor qilish: /cancel" % nm, ("firms", target), back=uc)
        if act == "fl":  # filial xodimi qilish: filiallar ro'yxati
            self.data.reload()
            names = self.data.koch_filials()
            if not names:
                return self.client_card(cid, target, mid, "⚠ Ko'chirish varag'ida filial topilmadi. Avval «🔄 Yangilash» ni bosing.")
            self.filsearch[cid] = names
            return self.a_screen(cid, mid, ("👥", nm, "🏬 Filial"), "🏬 <b>%s</b> qaysi filial xodimi?\nHozirgi: %s" % (
                nm, esc(norm(c["filial"]) or "—")), [[(n[:55], "fs:%d:%d:%s" % (target, i, self._ck(n)))] for i, n in enumerate(names)],
                back=uc)
        if act == "fs":
            name = self._pick(self.filsearch.get(cid, []), parts[2], (parts + ["", ""])[3])
            if not name:
                return self.panel_note(cid, mid, "⚠ Ro'yxat eskirgan, qaytadan tanlang.")
            self.set_role(target, "filial", name)
            if c["status"] != "blocked":  # bloklangan foydalanuvchi rol o'zgarishi bilan ochilib ketmasin
                self.activate(target)
            return card("✅ %s endi <b>filial xodimi · %s</b>" % (nm, esc(name)))
        if act == "fm":
            return self.ask(cid, mid, ("👥", nm, "🏢 Firma"), "%s uchun firma belgisini yozing 👇 (hozirgi: %s).\nBir nechta bo'lsa, ; bilan "
                            "ajrating.\nBekor qilish: /cancel" % (nm, esc(", ".join(self.firms_of(target)) or "—")), ("firms", target), back=uc)
        if act == "ac":
            return self.ask(cid, mid, ("👥", nm, "💰 Hisob"), "Hisob varag'idagi tashkilot nomining bir qismini yozing 👇 (qidiruv).\n"
                            "Bekor qilish: /cancel", ("accsearch", target), back=uc)
        if act == "sk":
            return self.client_card(cid, target, mid, "Hisob biriktirilmadi. Keyin kartadan qo'shishingiz mumkin: /c%d" % target)
        if act == "bl":
            self.set_status(target, "blocked")
            return card("⛔ Bloklandi: %s" % nm)
        if act == "del":
            return self.a_screen(cid, mid, ("👥", nm, "🗑 O'chirish"), "🗑 <b>%s</b> o'chirilsinmi? Bu qaytarilmaydi." % nm,
                                 [[("Ha, o'chirish", "dy:%d" % target), ("Bekor qilish", uc)]], back=uc)
        if act == "dy":
            self.delete_client(target)
            return self.users_menu(cid, mid, "🗑 O'chirildi: %s" % nm)
        if act == "vw":
            return self.view_as(cid, target, None, mid)
        if act == "v2":
            return self.view_as(cid, target, parts[2], mid)

    def after_firm(self, cid, target, msg, mid=None):
        c = self.client(target)
        return self.a_screen(cid, mid, ("👥", esc(c["name"] if c else str(target)), "💰 Hisob"),
                             msg + "\n\nHisob varag'idagi tashkilotni ham biriktirasizmi?",
                             [[("💰 Biriktirish", "ac:%d" % target), ("O'tkazib yuborish", "sk:%d" % target)]], back="uc:%d" % target)

    def reports_menu(self, cid, mid, note=None):
        """📊 Hisobotlar (Excel): tur tanlash; «🏢 Firma» — shu ichidagi qadam (fp → qaytib shu ekranga)."""
        tags, label = self.fm(cid)
        kb = [[("⏰ Srok (%d kun)" % self.srok_days, "a:r:srok"), ("📦 30+ kun", "a:r:old")],
              [("🧾 Buyurtma", "a:r:need"), ("📊 Qoldiq va sotuv", "a:r:stock")],
              [("💲 Narxlar", "a:r:narx"), ("💰 Hisob", "a:r:acc")], [(self.firm_btn(cid), "a:fm:rep")]]
        return self.a_screen(cid, mid, ("📊 Hisobotlar",), "<b>📊 Hisobotlar</b>\nExcel fayl: %s" % (
            "🏢 <b>%s</b>" % esc(label) if tags else "hamma firmalar"), kb, note=note)

    def admin_menu_action(self, cid, mid, p):
        a = p[0] if p else "home"
        if a == "home":
            self.state.pop(cid, None)
            return self.send_admin_panel(cid, mid)
        if a == "sec":  # bosh menyudagi bo'lim: ⏰ Srok / 📦 30+ kun / 🧾 Buyurtma / 💰 Hisob / 📊 Sotuv va qoldiq
            if len(p) > 1 and p[1] == "stock":
                return self.sq_screen(cid, mid)
            return self.admin_section(cid, p[1], mid) if len(p) > 1 and p[1] in ("srok", "old", "need", "acc") else None
        if a in ("sq", "sqa"):  # 📊 Sotuv va qoldiq: a:sq:<tartib>:<sahifa> | a:sqa (hamma firmalar)
            return self.sq_action(cid, mid, p)
        if a == "find":
            return self.a_screen(cid, mid, ("🔍 Dori qidirish",), "🔍 <b>Dori qidirish</b>%s\nDori nomini yozing 👇 (kamida 3 harf)" % (
                " · 🏢 %s" % esc(self.fm(cid)[1]) if self.fm(cid)[0] is not None else ""))
        if a == "rep":
            return self.reports_menu(cid, mid)
        if a == "r":
            return self.on_excel_callback(cid, p[1], mid)
        if a == "usr":
            return self.users_menu(cid, mid)
        if a == "bc":
            return self.bc_menu(cid, mid)
        if a == "log":
            return self.a_screen(cid, mid, ("📜 Jurnal",), self.log_text())
        if a == "stat":
            return self.a_screen(cid, mid, ("📈 Statistika",), self.stats_text())
        if a == "mt":
            return self.maint_action(cid, mid, p[1] if len(p) > 1 else "")
        if a in ("tr", "trx", "krm"):  # ko'chirish nazorati: bugungi holat, Excel, kechikkanlarga eslatma
            return self.koch_admin_action(cid, mid, a, p[1:])
        if a == "sk":
            return self.a_screen(cid, mid, ("🏭 Sklad",), self.sklad_admin_text(), [[("🔄 Yangilash", "a:sk"), ("📎 Excel", "a:skx")]])
        if a == "skx":
            s = self.sklad_summary()
            if not s["items"] and not s["excluded"]:
                return self.panel_note(cid, mid, "➖ Sklad bo'yicha taqsimot yo'q.")
            return self.panel_doc(cid, "Sklad_taqsimot_%s.xlsx" % self._today(), build_sklad(s), "🏭 <b>Sklad taqsimoti</b> · %s" % self._today())
        if a == "fm":  # firma filtri: a:fm (bosh menyu) | a:fm:<bo'lim> — tanlangach shu bo'limga qaytadi
            return self.firm_prompt(cid, mid, p[1] if len(p) > 1 else None)
        if a == "bq":
            kind = p[1]
            if kind == "text":
                return self.ask(cid, mid, ("📢 Xabar yuborish", "✉️ Matn"), "✉️ Ommaviy xabar matnini bitta xabar qilib yozing 👇\n"
                                "Keyin mijozlar ko'radigan ko'rinishini ko'rasiz va kimga yuborishni tanlaysiz.\nBekor qilish: /cancel",
                                ("bc_text", None), back="a:bc")
            if kind == "one":
                self.state[cid] = ("bc_user", None)
                return self.users_list(cid, mid, "active", 0, pick=True)
            return self.bc_confirm(cid, mid, kind)
        if a == "go":
            return self.bc_send_report(cid, p[1], mid)
        if a == "tx":
            return self.bc_send_text(cid, p[1], (p + ["", ""])[2], mid)
        if a == "set":
            return self.settings_action(cid, mid, p[1:])
        if a == "ref":
            return self.start_refresh(cid, mid=mid)

    # ---- sozlamalar ----
    def settings_action(self, cid, mid, p, note=None):
        sub = p[0] if p else ""
        cr = ("⚙️ Sozlamalar",)
        if sub == "auto":
            self.set_setting("auto_on", "0" if self.auto_on else "1")
        elif sub == "days":
            return self.ask(cid, mid, cr + ("⏰ Srok kuni",), "Srok necha kundan kam bo'lganda ogohlantirilsin? Raqam yozing 👇 "
                            "(hozir: %d).\nBekor qilish: /cancel" % self.srok_days, ("set_days", None), back="a:set")
        elif sub in ("wold", "wneed"):
            k = "weekly_" + sub[1:]
            self.set_setting(k, "0" if self.setting(k, "0") == "1" else "1")
        elif sub == "wday":
            self.set_setting("weekly_day", (self.weekly_day + 1) % 7)
        elif sub == "backup":
            try:
                note = "💾 Zaxira nusxa olindi: <code>%s</code>" % esc(self.backup_db())
            except Exception as e:  # noqa
                note = "❌ Zaxira olinmadi: %s" % esc(e)
        elif sub == "al" and len(p) > 1 and p[1] in dict(self.ALERTS):
            self.set_setting(p[1], "0" if self.alert_on(p[1]) else "1")
            return self.settings_action(cid, mid, ["alerts"])
        elif sub == "alerts":
            kb = [[("%s %s" % ("✅" if self.alert_on(k) else "⛔", name), "a:set:al:%s" % k)] for k, name in self.ALERTS]
            return self.a_screen(cid, mid, cr + ("🔔 Ogohlantirishlar",), "🔔 <b>Kunlik ogohlantirishlar</b>\nHar kuni %s da adminga "
                                 "bitta xabar + Excel. Bosib yoqing / o'chiring:" % self.auto_time.strftime("%H:%M"), kb, back="a:set")
        elif sub == "artime":
            return self.ask(cid, mid, cr + ("🔄 Avto-yangilash",), "Har kuni soat nechada Excel avtomatik yangilansin? HH:MM yozing 👇 "
                            "(masalan 08:30).\nO'chirish uchun: <code>-</code>\nBekor qilish: /cancel", ("set_artime", None), back="a:set")
        elif sub in ("kd", "kdz", "kdo"):  # ko'chirish muddati (toifa bo'yicha; kdz/kdo — eski xabarlardagi tugmalar)
            cat = {"kdz": "shoshilinch", "kdo": "oddiy"}.get(sub) or (p[1] if len(p) > 1 else "")
            if cat not in KOCH_CAT:
                return None
            k = KOCH_CAT[cat]
            return self.ask(cid, mid, cr + ("🔁 Ko'chirish", "%s %s" % (k["icon"], k["title"])),
                            "%s %s muddati (hozir: %s). HH:MM yozing 👇 (masalan 10:00), muddatsiz uchun «-».\nBekor qilish: /cancel" % (
                                k["icon"], k["title"], self.kset("koch_deadline_" + cat) or "muddatsiz"), ("set_kd", cat), back="a:set:koch")
        elif sub == "kn" and len(p) > 1 and p[1] in self.KOCH_NUMS:
            name, lo, hi = self.KOCH_NUMS[p[1]]
            return self.ask(cid, mid, cr + ("🔁 Ko'chirish",), "%s (hozir: %s). Son yozing 👇 (%d–%d).\nBekor qilish: /cancel" % (
                name, self.kset(p[1]), lo, hi), ("set_kn", p[1]), back="a:set:koch")
        elif sub == "koch":
            dl = lambda c: self.kset("koch_deadline_" + c) or "muddatsiz"
            kb = ([[("%s %s: %s" % (k["icon"], k["title"], dl(k["key"])), "a:set:kd:%s" % k["key"])] for k in KOCH_CATS] +
                  [[("%s: %s" % (name, self.kset(key)), "a:set:kn:%s" % key)] for key, (name, lo, hi) in self.KOCH_NUMS.items()])
            return self.a_screen(cid, mid, cr + ("🔁 Ko'chirish",), "🔁 <b>Ko'chirish sozlamalari</b>\nTopshiriq: %s dan keyingi "
                                 "«Yangilash». Bosib o'zgartiring:" % self.kset("koch_task_time"), kb, back="a:set", note=note)
        d = self.data
        wk = self.weekly_kinds
        text = ("<b>⚙️ Sozlamalar</b>\n\n"
                "Avto-yuborish (srok): <b>%s</b>, har kuni <b>%s</b>\n"
                "Srok chegarasi: <b>%d kun</b>\n"
                "Oxirgi avto-yuborish: %s\n\n"
                "📅 Haftalik avto-yuborish: <b>%s</b>, har <b>%s</b> %s\n"
                "Oxirgi haftalik: %s\n\n"
                "💾 bot.db zaxira nusxasi: har kuni (oxirgisi: %s)\n"
                "🔁 Ko'chirish nazorati: <b>" + ("yoqilgan" if self.koch_on else "o'chirilgan") + "</b> · muddatlar: " +
                esc(" · ".join("%s %s" % (k["icon"], self.kset("koch_deadline_" + k["key"]) or "muddatsiz") for k in KOCH_CATS)) +
                "\n\n"
                "Excel: %s\nYangilangan: %s\n"
                "Adminlar: %s") % (
            "yoqilgan ✅" if self.auto_on else "o'chirilgan ⛔", self.auto_time.strftime("%H:%M"), self.srok_days,
            self.setting("last_auto", "—"),
            ", ".join(self.KINDS[k] for k in wk) if wk else "o'chirilgan ⛔", self.WEEKDAYS[self.weekly_day].lower(),
            self.auto_time.strftime("%H:%M"), self.setting("last_weekly", "—"), self.setting("last_backup", "—"),
            esc(os.path.basename(d.path or d.file_pattern())), dt_text(d.when()),
            ", ".join(str(x) for x in sorted(self.admins)))
        on = lambda k: "✅" if k in wk else "⛔"
        kb = [[("Avto-yuborish: %s" % ("o'chirish" if self.auto_on else "yoqish"), "a:set:auto")],
              [("Srok kunini o'zgartirish", "a:set:days")],
              [("%s Haftalik 30+ kun" % on("old"), "a:set:wold"), ("%s Haftalik buyurtma" % on("need"), "a:set:wneed")],
              [("📅 Kun: %s" % self.WEEKDAYS[self.weekly_day], "a:set:wday"), ("💾 Zaxira olish", "a:set:backup")],
              [("🔁 Ko'chirish sozlamalari", "a:set:koch")],
              [("🔔 Ogohlantirishlar", "a:set:alerts")],
              [("🔄 Avto-yangilash: %s" % (self.auto_refresh_time.strftime("%H:%M") if self.auto_refresh_time else "o'chiq"),
                "a:set:artime")]]
        return self.a_screen(cid, mid, cr, text, kb, note=note)

    def admin_state_input(self, cid, st, text, m=None):
        """Admin yozgan matn (holat bo'yicha). Javob — yangi panel (eski panel o'chiriladi)."""
        kind, target = st
        # ommaviy/yakka matn: admin Telegramda qo'ygan formatlash (qalin, kursiv, havola) saqlanadi
        fmt = entities_to_html((m or {}).get("text") or text, (m or {}).get("entities")) if m else esc(text)
        cr = ("⚙️ Sozlamalar",)
        if kind == "firms":
            tags = [norm(x) for x in re.split(r"[;\n|]", text) if x.strip()]
            known = set(self.data.all_tags())
            miss = [t for t in tags if t not in known]
            self.set_firms(target, tags)
            c = self.client(target)
            self.state.pop(cid, None)
            if c and c["status"] == "pending":  # bloklangan foydalanuvchi firma belgisi tahrirlanganda ochilib ketmasin
                self.activate(target)
            msg = "Firmalar saqlandi: %s" % esc(", ".join(tags))
            if miss:
                msg += "\n⚠ Ma'lumotda topilmadi: %s" % esc(", ".join(miss))
            return self.after_firm(cid, target, msg)
        if kind == "accsearch":
            self.data.reload()
            q = norm(text)
            found = [n for n in self.data.hisob_names() if q in norm(n)][:10]
            c = self.client(target)
            crumbs = ("👥", esc(c["name"] if c else str(target)), "💰 Hisob")
            if not found:
                return self.ask(cid, None, crumbs, "Hech narsa topilmadi. Boshqa so'z yozing yoki /cancel.", st, back="uc:%d" % target)
            self.search[cid] = found
            return self.a_screen(cid, None, crumbs, "Qo'shish uchun tanlang:",
                                 [[(n[:55], "at:%d:%d:%s" % (target, i, self._ck(n)))] for i, n in enumerate(found)], back="uc:%d" % target)
        if kind == "oreq_qty":
            return self.ord_qty_admin_input(cid, st, text)
        if kind == "set_days":
            if not text.strip().isdigit() or not (1 <= int(text) <= 3650):
                return self.ask(cid, None, cr + ("⏰ Srok kuni",), "Faqat raqam yozing (masalan 90) yoki /cancel.", st, back="a:set")
            self.set_setting("srok_days", int(text))
            self.state.pop(cid, None)
            return self.settings_action(cid, None, [], note="✅ Saqlandi")
        if kind == "firm_pick":
            return self.firm_pick_input(cid, text)
        if kind == "set_artime":
            v = text.strip()
            if v in ("-", "0", "yo'q", "off"):
                v = ""
            elif not re.match(r"^([01]?\d|2[0-3]):[0-5]\d$", v):
                return self.ask(cid, None, cr + ("🔄 Avto-yangilash",), "Vaqtni HH:MM ko'rinishida yozing (masalan 08:30), o'chirish uchun "
                                "«-», yoki /cancel.", st, back="a:set")
            self.set_setting("auto_refresh_time", v)
            self.state.pop(cid, None)
            return self.settings_action(cid, None, [], note="✅ Saqlandi")
        if kind == "set_kd":
            v = text.strip()
            if v in ("-", "0", "yo'q"):
                v = ""
            elif not self.parse_hm(v):
                return self.ask(cid, None, cr + ("🔁 Ko'chirish",), "Vaqtni HH:MM ko'rinishida yozing (masalan 10:00), muddatsiz uchun «-», "
                                "yoki /cancel.", st, back="a:set:koch")
            else:
                v = self.parse_hm(v).strftime("%H:%M")
            self.set_setting("koch_deadline_" + target, v)
            self.state.pop(cid, None)
            return self.settings_action(cid, None, ["koch"], note="✅ Saqlandi")
        if kind == "set_kn":
            name, lo, hi = self.KOCH_NUMS[target]
            if not text.strip().isdigit() or not (lo <= int(text) <= hi):
                return self.ask(cid, None, cr + ("🔁 Ko'chirish",), "Son yozing (%d–%d) yoki /cancel." % (lo, hi), st, back="a:set:koch")
            self.set_setting(target, int(text))
            self.state.pop(cid, None)
            return self.settings_action(cid, None, ["koch"], note="✅ Saqlandi")
        if kind == "maint_text":
            self.state.pop(cid, None)
            self.set_setting("maint_text", fmt)
            return self.maint_action(cid, None, "")
        if kind == "user_search":
            found = self.find_users(text)
            if not found:
                return self.ask(cid, None, ("👥 Foydalanuvchilar", "🔍 Qidirish"), "Hech kim topilmadi. Boshqa so'z yozing yoki /cancel.",
                                st, back="a:usr")
            self.state.pop(cid, None)
            return self.a_screen(cid, None, ("👥 Foydalanuvchilar", "🔍 «%s»" % esc(self.short(text, 20))),
                                 "Topildi: <b>%d</b> ta" % len(found), [[self._user_btn(r)] for r in found], back="a:usr")
        if kind == "add_user":
            self.state.pop(cid, None)
            return self.a_screen(cid, None, ("👥 Foydalanuvchilar", "➕ Qo'shish"), self.add_user(text),
                                 [[("➕ Yana qo'shish", "ua")]], back="a:usr")
        if kind == "bc_user":
            found = self.find_users(text)
            if not found:
                return self.users_list(cid, None, "active", 0, pick=True, note="Hech kim topilmadi. Boshqa so'z yozing yoki /cancel.")
            self.state.pop(cid, None)
            return self.a_screen(cid, None, ("📢 Xabar yuborish", "👤 Bitta odamga"), "Kimga yuboramiz? Tanlang:",
                                 [[self._user_btn(r, pick=True)] for r in found], back="a:bq:one")
        if kind == "one_text":
            c = self.client(target)
            self.state.pop(cid, None)
            if not c:
                return self.send_admin_panel(cid, note="Mijoz topilmadi.")
            try:
                self.tg.send(target, fmt)
            except Exception as e:  # noqa
                self.mark_gone(target, e)
                return self.one_menu(cid, target, None, "❌ Yuborilmadi: %s" % esc(e))
            return self.one_menu(cid, target, None, "✅ Yuborildi: %s" % esc(c["name"]))
        if kind == "bc_text":
            self.pending_text[cid] = fmt
            self.state.pop(cid, None)
            n_all, n_firm, n_staff = (len(self.audience(x)) for x in ("all", "firm", "staff"))
            k = self._ck(fmt)  # eski ko'rinishdagi tugma keyingi yozilgan matnni yuborib yubormasin
            kb = [[("📣 Hammaga (%d)" % n_all, "a:tx:all:" + k)],
                  [("🏢 Firmalarga (%d)" % n_firm, "a:tx:firm:" + k), ("👔 Xodimlarga (%d)" % n_staff, "a:tx:staff:" + k)],
                  [("✏️ Qayta yozish", "a:bq:text"), ("Bekor qilish", "a:bc")]]
            return self.a_screen(cid, None, ("📢 Xabar yuborish", "Ko'rib chiqish"),
                                 "<b>Xabar ko'rinishi:</b>\n\n%s\n\nKimga yuboramiz? Tugma bosilishi bilan yuboriladi." % fmt, kb, back="a:bc")

    def audience(self, who):
        out = []
        for c in self.active_clients():
            if not self.can_use(c):
                continue
            if who == "firm" and c["role"] != "firm":  # filial xodimlari «Firmalarga» ga kirmaydi
                continue
            if who == "staff" and c["role"] != "staff":
                continue
            out.append(c)
        return out

    # ---- ommaviy yuborish ----
    def bc_confirm(self, cid, mid, kind):
        self.data.reload()
        withdata, nodata = 0, 0
        for c in self.audience("all"):
            ok = self.prepare(kind, self.tags_for(c), self.accs_for_client(c))
            if ok:
                withdata += 1
            else:
                nodata += 1
        text = ("<b>%s</b> yuboriladi.\n\n✅ Ma'lumoti bor: <b>%d</b> ta\n➖ Ma'lumoti yo'q (yuborilmaydi): <b>%d</b> ta\n\n🕒 Ma'lumot: %s\n\nDavom etasizmi?"
                % (self.KINDS[kind], withdata, nodata, dt_text(self.data.when())))
        return self.a_screen(cid, mid, ("📢 Xabar yuborish", self.KINDS[kind]), text,
                             [[("✅ Yuborish", "a:go:%s" % kind), ("Bekor qilish", "a:bc")]], back="a:bc")

    def bc_send_report(self, cid, kind, mid=None):
        return self.start_broadcast(cid, kind, "all", mid=mid)

    def bc_send_text(self, cid, who, ck=None, mid=None):
        text = self.pending_text.get(cid)
        if not text or (ck is not None and ck != self._ck(text)):
            return self.a_screen(cid, mid, ("📢 Xabar yuborish",), "Matn topilmadi yoki eskirgan (keyin boshqa matn yozilgan). "
                                 "Qaytadan boshlang.", [[("✉️ Matn yozish", "a:bq:text")]], back="a:bc")
        self.pending_text.pop(cid, None)
        return self.start_broadcast(cid, "text", who, text, mid)

    def start_broadcast(self, cid, kind, who, text=None, mid=None):
        """Ommaviy yuborishni fonda boshlaydi. Bir vaqtda faqat bittasi. Jarayon va natija — panelning o'zida (mid)."""
        with self.rlock:
            if self.bc_running:
                if cid:
                    self.panel_note(cid, mid, "📢 Boshqa yuborish hali tugamadi, iltimos kuting.")
                return False
            self.bc_running = True
        self.bg(self._broadcast_job, cid, kind, who, text, mid, name="broadcast")
        return True

    def deliver(self, c, kind, text=None):
        """Bitta foydalanuvchiga yuborish. 'ok' | 'skip' (ma'lumot yo'q) | 'gone' (botni tark etgan) | 'fail'."""
        cc = c["chat_id"]
        try:
            if kind == "text":
                self.tg.send(cc, text)
                return "ok"
            return "ok" if self.send_report(cc, kind, self.tags_for(c), self.accs_for_client(c), quiet=True) else "skip"
        except Exception as e:  # noqa
            if self.mark_gone(cc, e):
                return "gone"
            logging.exception("Yuborib bo'lmadi: %s", cc)
            return "fail"

    def _broadcast_job(self, cid, kind, who, text, mid=None):
        try:
            people = self.audience(who)
            cnt = {"ok": 0, "skip": 0, "gone": 0, "fail": 0}
            label = "✉️ Matn" if kind == "text" else self.KINDS.get(kind, kind)
            crumb = self.crumb("📢 Xabar yuborish", "Natija") + "\n\n"

            def status(final=False):
                done = sum(cnt.values())
                head = "✅ <b>Yuborish tugadi</b>" if final else "📢 <b>Yuborilmoqda…</b>"
                bar = "▰" * int(10 * done / max(1, len(people))) + "▱" * (10 - int(10 * done / max(1, len(people))))
                t = "%s%s · %s\n%s %d/%d\n\n✅ Yuborildi: <b>%d</b>" % (crumb, head, label, bar, done, len(people), cnt["ok"])
                if kind != "text":
                    t += "\n➖ Ma'lumot yo'q: %d" % cnt["skip"]
                if cnt["gone"]:
                    t += "\n🚫 Botni tark etgan: %d" % cnt["gone"]
                return t + "\n❌ Xato: %d" % cnt["fail"]

            if cid:  # jarayon panelning o'zida (tugmasiz — ikki marta ishga tushirib bo'lmaydi)
                self.panel_show(cid, mid, status(), None)
                mid = self.panel_mid(cid)
            last = time.time()
            for c in people:
                cnt[self.deliver(c, kind, text)] += 1
                if mid and time.time() - last > 3:  # jarayonni har 3 soniyada yangilaymiz
                    try:
                        self.tg.edit(cid, mid, status())
                    except Exception:  # noqa
                        pass
                    last = time.time()
                self.sleep(0.05)
            logging.info("Ommaviy yuborish tugadi (%s): %s", kind, cnt)
            if cid:
                kb = inline_kb([[("📢 Yana xabar", "a:bc"), ("🏠 Menyu", "a:home")]])
                try:
                    self.tg.edit(cid, mid, status(True), kb)
                    self._panel_keep(cid, mid, status(True), kb)
                except Exception:  # noqa
                    self.panel_resend(cid, status(True), kb)
            return cnt
        finally:
            self.bc_running = False

    # ---- foydalanuvchilar ro'yxati ----
    def list_clients(self, cid, page, only=None, mid=None):
        """Eski «pg:» tugmalari (/clients ro'yxati)."""
        q, args = "SELECT * FROM clients", []
        if only:
            q += " WHERE status=?"
            args.append(only)
        q += " ORDER BY created DESC"
        rows = self.db.execute(q, args).fetchall()
        per = 20
        part = rows[page * per:(page + 1) * per]
        cr = ("👥 Foydalanuvchilar", "📋 Mijozlar")
        if not part:
            return self.a_screen(cid, mid, cr, "Ro'yxat bo'sh.", back="a:usr")
        icon = {"active": "✅", "pending": "⏳", "blocked": "⛔"}
        lines = []
        for r in part:
            firms = self.role_label(r)
            lines.append("/c%d %s %s (%s) · %s" % (r["chat_id"], icon.get(r["status"], "?"), esc(r["name"]), esc(r["phone"]), esc(firms)))
        text = "<b>Mijozlar</b> (%d-%d / %d):\n\n%s" % (page * per + 1, page * per + len(part), len(rows), "\n".join(lines))
        rows_kb = []
        if (page + 1) * per < len(rows):
            rows_kb.append([("Keyingisi ▶", "pg:%d:%s" % (page + 1, only or "all"))])
        return self.a_screen(cid, mid, cr, text, rows_kb, back="a:usr")

    def client_card(self, cid, target, mid=None, note=None):
        c = self.client(target)
        if not c:
            return self.users_menu(cid, mid, "Mijoz topilmadi.")
        auto = [r["account"] for r in self.db.execute("SELECT account FROM acc_requests WHERE chat_id=? AND status='pending'", (target,))]
        role = {"staff": "xodim", "filial": "🏬 filial xodimi · %s" % (esc(norm(c["filial"])) or "— (filial tanlanmagan)")}.get(c["role"], "firma")
        holat = {"active": "✅ faol", "pending": "⏳ kutmoqda", "blocked": "⛔ bloklangan"}.get(c["status"], c["status"])
        text = ("<b>Mijoz</b>\nIsm: %s\nTelefon: %s\nID: %d\nHolat: %s\nRol: %s\nFirmalar: %s\nHisoblar: %s%s\n"
                "Oxirgi faollik: %s%s") % (
            esc(c["name"]), esc(c["phone"]), target, holat, role,
            esc(", ".join(self.firms_of(target)) or "—"), esc(", ".join(self.accounts_of(target)) or "—"),
            ("\n⏳ Tasdiq kutayotgan hisob: " + esc(", ".join(auto))) if auto else "",
            (c["last_seen"] or "—")[:16].replace("T", " "),
            ("\n🚫 Botni tark etgan (bloklagan): " + c["left_at"][:10]) if c["left_at"] else "")
        row1 = [("Firma belgisi", "fm:%d" % target), ("Hisob qo'shish", "ac:%d" % target)]
        if c["role"] == "filial":
            role_btns = [("🏢 Firma qilish", "rf:%d" % target), ("👔 Xodim qilish", "rs:%d" % target)]
        else:
            role_btns = [("👔 Xodim" if c["role"] != "staff" else "Firma qilish", "st:%d" % target)]
        row2 = [("⛔ Blok", "bl:%d" % target) if c["status"] == "active" else ("✅ Faollashtirish", "ok:%d" % target),
                ("🗑 O'chirish", "del:%d" % target)]
        row3 = role_btns + [("🏬 Filial xodimi" if c["role"] != "filial" else "🏬 Filialni o'zgartirish", "fl:%d" % target)]
        row4 = [("✉️ Xabar / hisobot yuborish", "m1:%d" % target), ("👁 Mijoz ko'rinishi", "vw:%d" % target)]
        return self.a_screen(cid, mid, ("👥 Foydalanuvchilar", esc(c["name"] or c["phone"] or str(target))), text,
                             [row1, row2, row3, row4], back="a:usr", note=note)

    # ================= YANGI PANELLAR: matnli javoblar, firma rejimi, hisob =================
    @staticmethod
    def smoney(v):
        return ("−" if v < 0 else "") + money(abs(v))

    @staticmethod
    def _ck(s):
        """Tugma uchun qisqa nazorat belgisi (6 belgi): eski xabardagi tugma yangi ro'yxatdagi boshqa nomni tanlamasin."""
        return hashlib.md5(norm(s).encode("utf-8")).hexdigest()[:6]

    def _pick(self, names, i, ck):
        """names[i] — faqat tugma ko'rsatilgandagi nom bilan bir xil bo'lsa; aks holda None (ro'yxat eskirgan)."""
        try:
            i = int(i)
        except (TypeError, ValueError):
            return None
        return names[i] if 0 <= i < len(names) and ck == self._ck(names[i]) else None

    @staticmethod
    def short(s, n=18):
        s = str(s or "")
        return s if len(s) <= n else s[:n - 1].rstrip() + "…"

    @staticmethod
    def kun_text(k):
        return ("−%d kun" % abs(k)) if k < 0 else "%d kun" % k

    def foot(self):
        return "\n\n🕒 Ma'lumot: %s" % dt_text(self.data.when())

    def head(self, title, label):
        t = "<b>%s</b>" % title
        if label:
            t += "\n🏢 <b>%s</b>" % esc(label)
        return t

    # ---- ichki hisoblar (umumiy qarzga kirmaydi) ----
    INTERNAL_DEFAULT = "ДОРИХОНА*, ИНВЕНТАРИЗАЦИЯ, КУП ЧИККАН*"

    def internal_patterns(self):
        """config.ini [hisob] internal (vergul bilan, * = boshlanishi). Panel sozlamasi (bot.db) bo'lsa, u ustun."""
        raw = self.setting("hisob_internal") or self.cfg.get("hisob", "internal", fallback=self.INTERNAL_DEFAULT)
        return [norm(x) for x in raw.split(",") if x.strip()]

    def is_internal(self, org):
        o = norm(org)
        for p in self.internal_patterns():
            if (p.endswith("*") and o.startswith(p[:-1].strip())) or o == p:
                return True
        return False

    # ---- hisob: tashkilot ↔ firma ----
    def accounts_for(self, tags, manual=(), auto=True):
        """auto=True (admin): firma belgisidagi telefon (oxirgi 9 raqam) Hisob varag'idagi tashkilot nomida bo'lsa, shu
        tashkilot ham olinadi. auto=False (mijoz): faqat admin tasdiqlagan (qo'lda biriktirilgan) tashkilotlar."""
        names = {norm(a) for a in (manual or ())}
        if not auto:
            return sorted(names)
        tags = {norm(t) for t in (tags or ())}
        d9s = {digits(t)[-9:] for t in tags if len(digits(t)) >= 9}
        for r in self.data.hisob:
            o = norm(r["org"])
            if o in names:
                continue
            dd = digits(o)
            if o in tags or any(x in dd for x in d9s):
                names.add(o)
        return sorted(names)

    def stock_cost(self, tags):
        """Firmaning hozirgi tovar qoldig'i kirim narxida (Srok varag'i: qoldiq × kirim narxi)."""
        return sum(r["ost"] * r["prih"] for r in self.data._by_tags(self.data.srok, tags))

    def accs_for_client(self, c):
        if c["role"] == "staff":
            return None
        return self.accounts_for(None, self.accounts_of(c["chat_id"]), auto=False)

    def hisob_net(self, tags, manual, auto=True):
        rows = [r for r in self.data.account_rows(self.accounts_for(tags, manual, auto)) if not self.is_internal(r["org"])]
        return rows, sum(r["debit"] - r["credit"] for r in rows)

    def hisob_text(self, tags, manual, label=None, auto=True):
        """Chatdagi Hisob javobi. Firma bo'yicha hisob topilmasa None qaytaradi."""
        d = self.data
        if tags is None:  # xodim/admin: umumiy ko'rinish (ichki hisoblar alohida)
            inner = [r for r in d.hisob if abs(r["debit"] - r["credit"]) >= 1 and self.is_internal(r["org"])]
            rows = [r for r in d.hisob if abs(r["debit"] - r["credit"]) >= 1 and not self.is_internal(r["org"])]
            deb = sum(r["debit"] for r in rows)
            cre = sum(r["credit"] for r in rows)
            top_c = sorted((r for r in rows if r["credit"] >= 1), key=lambda r: -r["credit"])[:5]
            top_d = sorted((r for r in rows if r["debit"] >= 1), key=lambda r: -r["debit"])[:5]
            t = self.head("💰 Hisob (umumiy)", label) + "\nTashkilotlar: <b>%d</b>\nBizga qarz (debet): <b>%s</b>\nBiz qarzmiz (kredit): <b>%s</b>" % (
                len(rows), money(deb), money(cre))
            if top_c:
                t += "\n\n<b>Eng katta qarzimiz:</b>\n" + "\n".join(
                    "%d. %s · %s" % (i, esc(self.short(r["org"], 34)), money(r["credit"])) for i, r in enumerate(top_c, 1))
            if top_d:
                t += "\n\n<b>Eng katta qarzdorlar:</b>\n" + "\n".join(
                    "%d. %s · %s" % (i, esc(self.short(r["org"], 34)), money(r["debit"])) for i, r in enumerate(top_d, 1))
            if inner:
                t += "\n\n<b>🏦 Ichki hisoblar</b> (yuqoridagi summalarga kirmaydi):\n" + "\n".join(
                    "• %s · %s" % (esc(self.short(r["org"], 34)), self.smoney(r["debit"] - r["credit"])) for r in inner[:10])
            return t + self.foot()
        rows, net = self.hisob_net(tags, manual, auto)
        if not rows:
            return None
        t = self.head("💰 Hisob", label)
        if len(rows) > 1:
            t += "\n" + "\n".join("• %s: %s" % (esc(self.short(r["org"], 30)), ("−" if r["debit"] - r["credit"] < 0 else "") +
                                                money(abs(r["debit"] - r["credit"]))) for r in rows)
        else:
            t += "\n" + esc(rows[0]["org"])
        if net < 0:
            qarz = -net
            s = self.stock_cost(tags)
            t += ("\n\nQarzimiz (Остаток на конец): <b>%s</b>\n"
                  "Tovar qoldig'i (kirim narxida): <b>%s</b>\n"
                  "   ↳ <i>qarzdan ayirib olish mumkin</i>\n"
                  "Sotilgan tovar uchun to'lov: <b>%s</b>") % (money(qarz), money(s), self.smoney(qarz - s))
        elif net > 0:
            t += "\n\nQarzdor (bizga qarz): <b>%s</b>" % money(net)
        else:
            t += "\n\nHisob nolga teng."
        return t + self.foot()

    # ---- filiallarni jamlash (chat matni uchun) ----
    @staticmethod
    def merge_srok(rows):
        g = {}
        for r in rows:
            k = (norm(r["name"]), norm(r["producer"]), r["date"], norm(r["firm"]))
            if k not in g:
                g[k] = dict(r, summa=0.0, ost=0.0)
            g[k]["ost"] += r["ost"]
            g[k]["summa"] += r["ost"] * r["prih"]
        return sorted(g.values(), key=lambda r: (r["kun"], norm(r["name"])))

    @staticmethod
    def merge_old(rows):
        g = {}
        for r in rows:
            k = (norm(r["name"]), norm(r["producer"]), norm(r["firm"]))
            if k not in g:
                g[k] = dict(r, ost=0.0, summa=0.0, kun=0)
            g[k]["ost"] += r["ost"]
            g[k]["summa"] += r["summa"]
            g[k]["kun"] = max(g[k]["kun"], r["kun"])
        return sorted(g.values(), key=lambda r: (-r["kun"], norm(r["name"])))

    TOP = 20

    def firm_line(self, r, tags):
        """Hamma firmalar ko'rinishida (xodim/admin) har qatorga firma nomi qo'shiladi."""
        return ("\n   🏢 " + esc(self.short(r["firm"], 40))) if tags is None and r.get("firm") else ""

    # ---- boshqa bo'limlarning matnli javobi ----
    def section_text(self, kind, tags, label=None):
        d = self.data
        d.reload()
        top = self.TOP
        more = lambda n: ("\n… yana %d ta (Excelda)" % (n - top)) if n > top else ""
        if kind == "srok":
            rows = self.merge_srok(d.srok_rows(tags, self.srok_days))
            h = self.head("⏰ Srogi %d kundan kam tovarlar" % self.srok_days, label)
            if not rows:
                return h + "\n\n✅ Bunday tovar yo'q." + self.foot()
            exp = sum(1 for r in rows if r["kun"] < 0)
            lines = ["%d. %s%s · %s · %s dona%s" % (i, "⚠️ " if r["kun"] < 0 else "", esc(self.short(r["name"], 40)),
                                                    self.kun_text(r["kun"]), qty(r["ost"]), self.firm_line(r, tags))
                     for i, r in enumerate(rows[:top], 1)]
            t = h + "\nPartiyalar: <b>%d</b>" % len(rows)
            if exp:
                t += " · muddati o'tgan: <b>%d</b>" % exp
            t += "\nJami summa: <b>%s</b> (kirim narxida)" % money(sum(r["summa"] for r in rows))
            t += "\n\n<b>Eng yaqinlari:</b>\n" + "\n".join(lines) + more(len(rows))
            return t + self.foot()
        if kind == "old":
            rows = self.merge_old(d.old_rows(tags))
            h = self.head("📦 30 kundan ortiq turgan tovarlar", label)
            if not rows:
                return h + "\n\n✅ Bunday tovar yo'q." + self.foot()
            lines = ["%d. %s · %d kun · %s dona%s" % (i, esc(self.short(r["name"], 40)), r["kun"], qty(r["ost"]), self.firm_line(r, tags))
                     for i, r in enumerate(rows[:top], 1)]
            t = h + "\nTovarlar: <b>%d</b>\nJami summa: <b>%s</b> (kirim narxida)" % (len(rows), money(sum(r["summa"] for r in rows)))
            t += "\n\n<b>Eng uzoq turganlari:</b>\n" + "\n".join(lines) + more(len(rows))
            return t + self.foot()
        if kind == "need":
            rows = sorted(self.need_rows(tags), key=lambda r: -r["need"])
            h = self.head("🧾 Buyurtma tavsiyasi", label)
            if not rows:
                return h + "\n\n✅ Hozircha buyurtma kerak emas." + self.foot()
            lines = ["%d. %s · kerak <b>%s</b>%s" % (i, esc(self.short(r["name"], 40)), qty(r["need"]), self.firm_line(r, tags))
                     for i, r in enumerate(rows[:top], 1)]
            t = h + "\nTovarlar: <b>%d</b> · jami yetishmaydi: <b>%s</b> dona" % (len(rows), qty(sum(r["need"] for r in rows)))
            t += "\n\n<b>Eng ko'p kerakli:</b>\n" + "\n".join(lines) + more(len(rows))
            return t + self.foot()
        if kind == "koch":
            rows = d.koch_rows(tags)
            h = self.head("🔁 Filiallararo ko'chirish", label)
            if not rows:
                return h + "\n\n✅ Hozircha ko'chirish kerak bo'lgan tovar yo'q." + self.foot()
            routes = {}
            for r in rows:
                g = routes.setdefault((r["src"], r["dst"]), [0, 0.0, 0.0])
                g[0] += 1
                g[1] += r["qty"]
                g[2] += r["summa"]
            t = h + "\nQatorlar: <b>%d</b> · <b>%s</b> qadoq\nJami summa: <b>%s</b> (kirim narxida)" % (
                len(rows), qty(sum(r["qty"] for r in rows)), money(sum(r["summa"] for r in rows)))
            t += "\n\n<b>Yo'nalishlar:</b>\n" + "\n".join(
                "• %s → %s · %d ta · %s qadoq · %s" % (esc(self.short(a, 20)), esc(self.short(b, 20)), g[0], qty(g[1]), money(g[2]))
                for (a, b), g in sorted(routes.items(), key=lambda x: -x[1][2]))
            t += "\n\n<b>Eng qimmatlari:</b>\n" + "\n".join(
                "%d. %s · %s → %s · <b>%s</b> qadoq · %s%s" % (i, esc(self.short(r["name"], 36)), esc(self.short(r["src"], 16)),
                                                          esc(self.short(r["dst"], 16)), qty(r["qty"]), money(r["summa"]),
                                                          self.firm_line(r, tags))
                for i, r in enumerate(rows[:top], 1)) + more(len(rows))
            return t + self.foot()
        if kind == "stock":
            rows = d.stock_rows(tags)
            h = self.head("📊 Qoldiq va sotuv (%d kun)" % d.period_days, label)
            if not rows:
                return h + "\n\nMa'lumot topilmadi." + self.foot()
            best = sorted(rows, key=lambda r: -r["sales"])[:top]
            t = h + "\nTovarlar: <b>%d</b>\nJami qoldiq: <b>%s</b> · jami sotuv: <b>%s</b>" % (
                len(rows), qty(sum(r["stock"] for r in rows)), qty(sum(r["sales"] for r in rows)))
            t += "\n\n<b>Eng ko'p sotilgan:</b>\n" + "\n".join(
                "%d. %s · sotuv %s · qoldiq %s%s" % (i, esc(self.short(r["name"], 36)), qty(r["sales"]), qty(r["stock"]), self.firm_line(r, tags))
                for i, r in enumerate(best, 1))
            return t + self.foot()
        return "Noma'lum bo'lim."

    # ---- admin / xodim: firma filtri va bo'limlar ----
    def fm(self, cid):
        """(tags yoki None, nomi yoki None) — admin/xodimning tanlangan firmasi."""
        v = self.fmode.get(cid)
        return (set(v[0]), v[1]) if v else (None, None)

    def eff(self, c):
        """(tags, label): oddiy firma uchun o'z firmalari; xodim uchun tanlangan firma yoki None (hamma)."""
        if c["role"] == "staff":
            return self.fm(c["chat_id"])
        return self.tags_for(c), None

    def firm_btn(self, cid):
        tags, label = self.fm(cid)
        return "🏢 Firma: hammasi" if tags is None else "🏢 Firma: " + self.short(label, 16)

    def admin_kb(self, cid):
        """Avvalgi pastki klaviatura (endi xabarlarga biriktirilmaydi; tugmalari admin_button orqali taniladi)."""
        return reply_kb([["⏰ Srok", "📦 30+ kun"], ["🧾 Buyurtma", "💰 Hisob"],
                         ["👥 Foydalanuvchilar", "📢 Xabar yuborish"], [self.firm_btn(cid), "🔄 Yangilash"]],
                        is_persistent=True)

    @staticmethod
    def admin_button(text):
        if text in ("⏰ Srok", "📦 30+ kun", "🧾 Buyurtma", "💰 Hisob", "👥 Foydalanuvchilar", "📢 Xabar yuborish"):
            return {"⏰ Srok": "srok", "📦 30+ kun": "old", "🧾 Buyurtma": "need", "💰 Hisob": "acc",
                    "👥 Foydalanuvchilar": "users", "📢 Xabar yuborish": "bc"}[text]
        if text == "🔄 Yangilash":
            return "refresh"
        if text.startswith("🏢 Firma"):
            return "firm"
        if text.startswith("📊 Sotuv"):
            return "stock"
        if text in ("⚙️ Admin panel", "🏠 Menyu"):
            return "home"
        return None

    def admin_press(self, cid, key, mid=None):
        """Eski pastki tugma (yoki shu nomli tugma) → panelning mos darajasi."""
        if key == "home":
            return self.send_admin_panel(cid, mid)
        if key in ("srok", "old", "need", "acc"):
            return self.admin_section(cid, key, mid)
        if key == "stock":
            return self.sq_screen(cid, mid)
        if key == "users":
            return self.users_menu(cid, mid)
        if key == "bc":
            return self.bc_menu(cid, mid)
        if key == "refresh":
            return self.start_refresh(cid, mid=mid)
        if key == "firm":
            return self.firm_prompt(cid, mid)

    SECTION = {"srok": "⏰ Srok", "old": "📦 30+ kun", "need": "🧾 Buyurtma", "stock": "📊 Sotuv va qoldiq", "acc": "💰 Hisob",
               "rep": "📊 Hisobotlar"}

    def admin_section(self, cid, kind, mid=None, note=None):
        """⏰ Srok / 📦 30+ kun / 🧾 Buyurtma / 💰 Hisob: xulosa + [📎 Excel] [🏢 Firma: …] + navigatsiya."""
        tags, label = self.fm(cid)
        rows = [[("📎 Excel fayl", "x:" + kind)], [(self.firm_btn(cid), "a:fm:" + kind)]]
        if kind == "acc":
            text = self.hisob_text(tags, None, label)
            if text is None:
                text = ("💰 <b>%s</b> uchun hisob topilmadi.\nFirma belgisida telefon raqami bo'lsa, tashkilot avtomatik bog'lanadi. "
                        "Aks holda Foydalanuvchilar → Hisob qo'shish orqali biriktiring." % esc(label))
                rows = rows[1:]
        else:
            text = self.section_text(kind, tags, label)
        return self.a_screen(cid, mid, (self.SECTION[kind],), text, rows, note=note)

    def firm_prompt(self, cid, mid=None, ret=None):
        """🏢 Firma tanlash — ichma-ich daraja: firma nomi/telefon/dori yoziladi (matn), natija tugmalarida tanlanadi.
        ret — tanlangach qaytiladigan bo'lim (srok/old/need/acc/stock/rep), None — bosh menyu."""
        ret = ret if ret in self.SECTION else None
        self.fret[cid] = ret
        self.state[cid] = ("firm_pick", None)
        tags, label = self.fm(cid)
        body = ("🏢 <b>Firma tanlash</b>\nHozir: <b>%s</b>\n\nFirma belgisi, telefon, nomi yoki firmaning bitta dorisi nomini "
                "yozing 👇\n<i>Tanlangach, hamma bo'limlar faqat shu firma bo'yicha ishlaydi.</i>\nBekor qilish: /cancel") % (
            esc(label) if tags is not None else "hamma firmalar")
        rows = [[("🌐 Hamma firmalar", "fp:off")]] if tags is not None else []
        return self.panel_show(cid, mid, *self.screen(self.firm_crumbs(ret), body, rows, back=self.firm_back(cid, ret),
                                                      home=self.home_cb(cid)))

    def firm_crumbs(self, ret):
        return ((self.SECTION[ret],) if ret else ()) + ("🏢 Firma",)

    def firm_back(self, cid, ret):
        """Firma tanlashdan «◀ Orqaga»: kelgan bo'limga."""
        admin = cid in self.admins
        if not ret:
            return "a:home" if admin else "c:home"
        if ret == "rep":
            return "a:rep"
        return ("a:sec:" if admin else "c:") + ret

    def firm_pick_input(self, cid, text):
        q, qd = norm(text), digits(text)
        found = [t for t in self.data.all_tags() if q in t or (len(qd) >= 4 and qd in digits(t))][:10]
        hint = {}
        if len(q) >= 3:
            # dori nomi bo'yicha: shu dori qaysi firmaniki bo'lsa, o'sha firma topiladi
            for rows in (self.data.sale, self.data.srok, self.data.old, self.data.narx):
                for r in rows:
                    if r["firm"] and q in norm(r.get("name") or ""):
                        t = norm(r["firm"])
                        hint.setdefault(t, r.get("name") or "")
            for t in sorted(hint):
                if t not in found and len(found) < 10:
                    found.append(t)
        ret = self.fret.get(cid)
        crumbs, home = self.firm_crumbs(ret), self.home_cb(cid)
        if not found:
            return self.panel_resend(cid, *self.screen(crumbs, "Hech narsa topilmadi. Boshqa so'z yozing yoki /cancel.",
                                                       back=self.firm_back(cid, ret), home=home))
        if len(found) == 1:
            return self.firm_set(cid, found[0])
        self.fsearch[cid] = found
        rows = [[(bl(self.short(t + ((" · " + hint[t]) if t in hint else ""), 55)), "fp:%d:%s" % (i, self._ck(t)))] for i, t in enumerate(found)]
        if self.fm(cid)[0] is not None:
            rows.append([("🌐 Hamma firmalar", "fp:off")])
        return self.panel_resend(cid, *self.screen(crumbs + ("«%s»" % esc(self.short(text.strip(), 20)),),
                                                   "Topildi: <b>%d</b> ta. Tanlang (yoki boshqa so'z yozing):" % len(found), rows,
                                                   back="a:fm:%s" % ret if ret else "a:fm", home=home))

    def firm_pick_cb(self, cid, mid, p):
        """fp:<i>:<ck> — qidiruv natijasidan firma; fp:off — hamma firmalar."""
        if p[1] == "off":
            self.state.pop(cid, None)
            self.fmode.pop(cid, None)
            if self.fret.get(cid) == "stock":  # 📊 Sotuv va qoldiq: «hamma firmalar» tanlandi — qayta so'ralmaydi
                self.sqall.add(cid)
            return self.firm_return(cid, mid, "🌐 Firma filtri o'chirildi: hamma firmalar bo'yicha.")
        name = self._pick(self.fsearch.get(cid, []), p[1], p[2])
        if not name:
            return self.panel_note(cid, mid, "⚠ Ro'yxat eskirgan, qaytadan qidiring.")
        return self.firm_set(cid, name, mid)

    def firm_set(self, cid, tag, mid=None):
        self.state.pop(cid, None)
        self.fmode[cid] = ({norm(tag)}, tag)
        return self.firm_return(cid, mid, "✅ <b>Firma tanlandi:</b> 🏢 %s\nEndi hamma bo'limlar faqat shu firma bo'yicha." % esc(tag))

    def firm_return(self, cid, mid, note):
        """Firma tanlangach — kelgan bo'limga (yoki bosh menyuga), izoh bilan."""
        ret = self.fret.pop(cid, None)
        if cid in self.admins:
            if ret == "rep":
                return self.reports_menu(cid, mid, note)
            if ret in ("srok", "old", "need", "acc"):
                return self.admin_section(cid, ret, mid, note)
            if ret == "stock":
                return self.sq_screen(cid, mid, note=note)
            return self.send_admin_panel(cid, mid, note)
        c = self.client(cid)
        if ret in ("srok", "old", "need", "stock", "acc"):
            return self.client_section(cid, c, ret, mid, note)
        return self.panel_show(cid, mid, *self.client_home(c), note=note)

    # ---- mijoz paneli ----
    def client_kb(self, c):
        """Bosh menyu tugmalari pastki klaviatura ko'rinishida (server paneli «mijoz ko'rinishi», eski testlar).
        Telegramga endi yuborilmaydi — bot inline panel (client_home) ishlatadi."""
        if c["role"] == "filial":
            return self.filial_kb(c)
        return reply_kb([[t for t, _ in row] for row in self.client_root_rows(c)], is_persistent=True)

    def sync_kb(self, cid, c):
        """Avval sonlar o'zgarsa pastki klaviaturani qayta yuborardi. Pastki klaviatura yo'q: sonlar panel ochilganda
        (🏠 Menyu / 🔄 Yangilash) yangilanadi, alohida xabar yuborilmaydi."""
        return None

    @staticmethod
    def client_button(text):
        for key, kind in (("💰 Hisob", "acc"), ("⏰ Srok", "srok"), ("📦 30+", "old"), ("🧾 Buyurtma", "need"), ("📊 Qoldiq", "stock"),
                          ("📊 Sotuv", "stock")):
            if text.startswith(key):
                return kind
        return None

    def client_section(self, cid, c, kind, mid=None, note=None):
        """Mijoz bo'limi panelning o'zida: yo'l qatori + xulosa + [📎 Excel] (+ xodimga firma filtri) + navigatsiya."""
        staff = c["role"] == "staff"
        if kind == "koch" and not staff:  # ko'chirish faqat xodimlar uchun
            return self.panel_show(cid, mid, *self.client_home(c))
        if kind == "stock":  # 📊 Sotuv va qoldiq: dori bloklari, saralash, sahifalash
            return self.sq_screen(cid, mid, note=note)
        if kind == "need" and c["role"] == "firm":  # 🧾 Buyurtma so'rovi (eski hisobot — «📄 Buyurtma tavsiyasi (Excel)»)
            return self.ord_list_screen(cid, c, mid, note=note)
        tags, label = self.eff(c)
        rows = [[("📎 Excel fayl", "x:" + kind)]]
        if kind == "acc":
            if staff:
                text = self.hisob_text(tags, None, label)
                if text is None:
                    text, rows = ("💰 <b>%s</b> uchun hisob topilmadi.\nFirma belgisida telefon raqami bo'lsa, tashkilot avtomatik "
                                  "bog'lanadi." % esc(label)), []
            else:
                text = self.hisob_text(tags, self.accounts_of(cid), auto=False)
                if text is None:
                    rows = []
                    if self.ensure_acc_requests(cid):
                        text = "💰 Hisobingiz <b>tasdiqlash uchun adminga yuborildi</b>.\nTasdiqlangach shu yerdan ko'rasiz."
                    else:
                        text = self.request_account(cid, c)
        else:
            text = self.section_text(kind, tags, label)
        if staff:
            rows.append([(self.firm_btn(cid), "a:fm:" + kind)])
        return self.panel_show(cid, mid, *self.screen((self.SECTION.get(kind, kind),), text, rows, back="c:home", home="c:home"),
                               note=note)

    # ---- 📊 Sotuv va qoldiq: har dori alohida blok (filiallar bo'yicha sotildi / bor), saralash va sahifalash ----
    SQ_SORTS = (("s", "🔥 Sotuv"), ("t", "⚠️ Tugayapti"), ("a", "🔤 А–Я"))
    SQ_PER = 5    # sahifadagi dori
    SQ_LOW = 7    # «≈ N kunga yetadi» shundan kam — ⚠️ va «buyurtma bering»
    SQ_MARK = 3   # filialda qoldiq 0 yoki shuncha kundan kamga yetsa — «!» (jadval ichida emoji yo'q)
    SQ_DEFAULT_NAMES = ("Chinoz", "Urgut", "Erkin", "Sklad")

    def branch_names(self):
        """Sale varag'idagi filiallar (v[0]/v[1], v[2]/v[3], v[4]/v[5]) va sklad (v[6]) nomlari — 6-qator sarlavhasidan
        («ДОРИХОНА ЧИНОЗ» → «Чиноз»), bo'lmasa botdagi odatiy nomlar (kirillda). Jadval kirill matn bo'yicha tekislanadi."""
        heads = self.data.sale_heads or []
        out = []
        for i, dflt in enumerate(self.SQ_DEFAULT_NAMES):
            k = fil_key(heads[i]) if i < len(heads) and heads[i] else ""
            out.append(self.short(k.capitalize(), 11) if k else UK.lat2cyr(dflt))
        return out

    def sq_days(self, r):
        """Qoldiq necha kunga yetadi: qoldiq ÷ (sotuv / davr). Sotuv yo'q — None."""
        return r["stock"] / (r["sales"] / float(self.data.period_days or 30)) if r["sales"] > 0 else None

    def sq_rows(self, tags, sort):
        """Saralangan dorilar. 's' — sotuv ↓, 't' — tugayapti (≈ SQ_LOW kundan kam, eng tezi birinchi), 'a' — А–Я."""
        rows = self.data.stock_rows(tags)
        if sort == "t":
            rows = [r for r in rows if self.sq_days(r) is not None and self.sq_days(r) < self.SQ_LOW]
            return sorted(rows, key=lambda r: (self.sq_days(r), -r["sales"], norm(r["name"])))
        if sort == "a":
            return rows
        return sorted(rows, key=lambda r: (-r["sales"], norm(r["name"])))

    def sq_block(self, r):
        """💊 <b>Nomi</b> + <pre> jadval (Sotildi / Bor har filialda, Sklad, Jami) + «≈ N kunga yetadi»."""
        names = self.branch_names()
        cy = UK.lat2cyr
        jami = cy("Jami")
        w = max(len(x) for x in names + [jami]) + 1
        v, p = r["v"], float(self.data.period_days or 30)
        lines = [" " * w + cy("Sotildi").rjust(8) + cy("Bor").rjust(6)]
        for n, (si, bi) in zip(names[:3], ((1, 0), (3, 2), (5, 4))):
            s, b = v[si], v[bi]
            low = s > 0 and (b <= 0 or b * p / s < self.SQ_MARK)
            lines.append(n.ljust(w) + qty(s).rjust(8) + qty(b).rjust(6) + (" !" if low else ""))
        lines.append(names[3].ljust(w) + "—".rjust(8) + qty(v[6]).rjust(6))
        lines.append(jami.ljust(w) + qty(r["sales"]).rjust(8) + qty(r["stock"]).rjust(6))
        days = self.sq_days(r)
        low = days is not None and days < self.SQ_LOW
        if days is None:
            tail = "≈ — kunga yetadi"
        else:
            tail = "≈ %s kunga yetadi%s" % ("365+" if days > 365 else int(days), " · buyurtma bering" if low else "")
        return "💊 <b>%s</b>%s\n<pre>%s</pre>\n<i>%s</i>" % (esc(r["name"]), " ⚠️" if low else "", esc("\n".join(lines)), tail)

    def sq_text(self, tags, label, sort="s", page=0):
        """(matn, sahifa, sahifalar soni). tags — firma(lar) (None — hamma), label — sarlavhadagi nom."""
        d = self.data
        d.reload()
        sort = sort if sort in dict(self.SQ_SORTS) else "s"
        rows = self.sq_rows(tags, sort)
        allrows = rows if sort != "t" else d.stock_rows(tags)
        icon, title = ("⚠️", "Tugayapti") if sort == "t" else ("📊", "Sotuv va qoldiq")
        h = "%s <b>%s</b>%s" % (icon, title, (" · %s" % esc(label)) if label else "")
        if sort == "t":
            h += "\n%d kundan kam yetadigan dorilar: <b>%d</b> ta" % (self.SQ_LOW, len(rows))
        else:
            h += "\nOxirgi %d kun · <b>%d</b> xil dori\nSotildi: <b>%s</b> ta · Bor: <b>%s</b> ta\nTartib: %s" % (
                d.period_days, len(allrows), qty(sum(r["sales"] for r in allrows)), qty(sum(r["stock"] for r in allrows)),
                "eng ko'p sotilgan (🔥 Sotuv)" if sort == "s" else "alifbo bo'yicha (🔤 А–Я)")
        if not rows:
            body = "✅ %d kundan kam yetadigan dori yo'q." % self.SQ_LOW if sort == "t" else "Ma'lumot topilmadi."
            return "%s\n\n%s%s" % (h, body, self.foot()), 0, 1
        per = self.SQ_PER
        pages = (len(rows) - 1) // per + 1
        page = page % pages
        part = rows[page * per:(page + 1) * per]
        return "%s\n\n%s\n\n%d–%d / %d · 🕒 %s" % (h, "\n".join(self.sq_block(r) for r in part), page * per + 1,
                                                  page * per + len(part), len(rows), dt_text(d.when())), page, pages

    def sq_who(self, cid):
        """(prefiks, tags, label, xodimmi/adminmi) yoki None (bo'lim bu foydalanuvchiga yo'q)."""
        if cid in self.admins:
            tags, label = self.fm(cid)
            return "a:", tags, label, True
        c = self.client(cid)
        if not c or c["status"] != "active" or not self.can_use(c) or c["role"] == "filial":
            return None
        if c["role"] == "staff":
            tags, label = self.fm(cid)
            return "c:", tags, label, True
        tags = self.tags_for(c)
        return "c:", tags, ", ".join(sorted(tags))[:60], False

    def sq_screen(self, cid, mid=None, sort="s", page=0, note=None):
        """📊 Sotuv va qoldiq. Firma vakili — faqat o'z firmasi (tags). Xodim / admin — avval firma (yoki «🌐 Hamma firmalar»)."""
        who = self.sq_who(cid)
        if not who:
            return None
        pre, tags, label, staff = who
        home = pre + "home"
        if staff and tags is None and cid not in self.sqall:
            return self.sq_firm_pick(cid, mid, pre, note)
        text, page, pages = self.sq_text(tags, label if tags is not None else ("Hamma firmalar" if staff else label), sort, page)
        sort = sort if sort in dict(self.SQ_SORTS) else "s"
        cb = lambda s_, p_: "%ssq:%s:%d" % (pre, s_, p_)
        rows = [[(t, cb(k, 0), "primary" if k == sort else None) for k, t in self.SQ_SORTS]]
        if pages > 1:
            rows.append([("◀", cb(sort, (page - 1) % pages)), ("%d / %d" % (page + 1, pages), cb(sort, page)),
                         ("▶", cb(sort, (page + 1) % pages))])
        rows.append([("🔍 Dori qidirish", pre + "find"), ("📥 Excel", "x:stock")])
        if staff:
            rows.append([(self.firm_btn(cid), "a:fm:stock")])
        crumbs = (self.SECTION["stock"],) + (("⚠️ Tugayapti",) if sort == "t" else ())
        return self.panel_show(cid, mid, *self.screen(crumbs, text, rows, back=home, home=home), note=note)

    def sq_firm_pick(self, cid, mid, pre, note=None):
        """Xodim / admin: qaysi firma? Eng ko'p sotuvchi firmalar tugmada, nomini yozish ham mumkin (firm_pick holati)."""
        self.fret[cid] = "stock"
        self.state[cid] = ("firm_pick", None)
        sales, cnt = {}, {}
        for r in self.data.sale:
            t = norm(r["firm"])
            if t:
                sales[t] = sales.get(t, 0) + r["sales"]
                cnt[t] = cnt.get(t, 0) + 1
        top = sorted(sales, key=lambda t: (-sales[t], t))[:6]
        self.fsearch[cid] = top
        btns = [(bl(self.short(t, 22)) + " · %d" % cnt[t], "fp:%d:%s" % (i, self._ck(t))) for i, t in enumerate(top)]
        rows = [btns[i:i + 2] for i in range(0, len(btns), 2)] + [[("🌐 Hamma firmalar", pre + "sqa")]]
        body = ("📊 <b>Sotuv va qoldiq</b> · 👔 %s\nQaysi firma dorilarini ko'ramiz?\nEng ko'p sotuvchi firmalar 👇\n"
                "Yoki firma nomini yozing.") % ("Admin" if cid in self.admins else "Xodim")
        home = pre + "home"
        return self.panel_show(cid, mid, *self.screen((self.SECTION["stock"], "🏢 Firma"), body, rows, back=home, home=home),
                               note=note)

    def sq_action(self, cid, mid, p):
        """p: ['sq', <tartib>, <sahifa>] | ['sqa'] (hamma firmalar)."""
        self.state.pop(cid, None)
        if p[0] == "sqa":
            if self.sq_who(cid) and self.sq_who(cid)[3]:
                self.sqall.add(cid)
                self.fmode.pop(cid, None)
            return self.sq_screen(cid, mid)
        try:
            page = int(p[2]) if len(p) > 2 else 0
        except ValueError:
            page = 0
        if cid not in self.admins:
            self.touch(cid, "stock")
        return self.sq_screen(cid, mid, p[1] if len(p) > 1 else "s", page)

    def on_excel_callback(self, cid, kind, mid=None):
        """«📎 Excel fayl» tugmasi (admin: tanlangan firma bo'yicha, mijoz: o'z firmasi). Fayl — alohida xabar, panel pastga."""
        if kind not in ("srok", "old", "need", "stock", "acc", "narx", "koch", "kz", "ko"):
            return
        if cid in self.admins:
            if kind in ("kz", "ko"):  # filial Excellari faqat filial xodimi uchun
                return
            tags, _ = self.fm(cid)
            accs = self.accounts_for(tags) if (kind == "acc" and tags is not None) else None
            return self.send_report(cid, kind, tags, accs, mid=mid, panel=True)
        c = self.client(cid)
        if not c or c["status"] != "active" or not self.can_use(c):
            return
        if c["role"] == "filial":  # filial xodimi: faqat o'z filialining toifa Exceli (x:kz / x:ko — eski xabarlardagi tugmalar)
            if kind in ("kz", "ko"):
                cat = "shoshilinch" if kind == "kz" else "oddiy"
                self.touch(cid, "excel:" + cat)
                return self.filial_excel(cid, c, cat, mid=mid)
            return
        if kind in ("koch", "kz", "ko"):  # firma mijozi va xodim ko'chirish Excelini olmaydi
            return
        self.touch(cid, "excel:" + kind)
        if c["role"] == "staff":
            tags, _ = self.fm(cid)
            return self.send_report(cid, kind, tags, self.accounts_for(tags) if (kind == "acc" and tags is not None) else None,
                                    mid=mid, panel=True)
        return self.send_report(cid, kind, self.tags_for(c), self.accs_for_client(c), mid=mid, panel=True)

    # ================= 🧾 buyurtma so'rovi: firma vakili → admin (dizayn/BUYURTMA-SOROV.md) =================
    # Firma vakili (c:b:*): ro'yxat (tugagan / tugayotgan) → savat (bot.db settings: ocart_<chat>, panel qayta tushsa ham saqlanadi)
    # → «📤 Adminga yuborish» (buyurtma.create). Adminlar (oq:*): har biriga alohida karta (order_req_msgs da message_id),
    # dori bo'yicha qaror shu kartaning o'zida; «📤 Javobni yuborish» — birinchi javob bergan hal qiladi (buyurtma.finish),
    # order_notify — vakilga natija va hamma adminlarda karta bir qatorga qisqaradi. Server paneli ham order_notify ni chaqiradi.
    ORD_DAYS = 7        # shuncha kundan kamga yetadigan dorilar ro'yxatda (qoldiq 0 — birinchi)
    ORD_PER = 5         # ro'yxat sahifasidagi dori
    ORD_MAX = 100000    # savatdagi bitta dori soni chegarasi
    ORD_CART_MAX = 40   # savatdagi dorilar soni chegarasi (Telegram tugmalari ~100 tadan oshmasin)
    ORD_BTN_PER = 10    # savat / admin kartasida bir sahifadagi dori tugmalari (◀ ▶)
    ORD_INT = re.compile(r"^(0|[1-9]\d{0,5})$")  # yozilgan son: faqat butun (nan / inf / 2.5 / -1 emas)
    ORD_STEPS = (("−5", -5), ("−1", -1), ("+1", 1), ("+5", 5))
    ORD_ITEM_LBL = {BY.I_WAIT: "kutmoqda", BY.I_OK: "qabul", BY.I_CHANGED: "o'zgartirildi", BY.I_REJ: "rad"}

    @staticmethod
    def ord_key(name, producer):
        return "%s|%s" % (norm(name), norm(producer))

    def ord_label(self, c):
        return ", ".join(sorted(self.tags_for(c) or ()))[:80]

    def oreq_low(self, tags):
        """Firma vakili ko'radigan ro'yxat: o'z firmasining sotuvi bor, jami qoldig'i 0 yoki ≤ ORD_DAYS kunga yetadigan dorilari
        (qora ro'yxat — need_rows dagi kabi chiqariladi). Qatorga: "days" (qoldiq necha kunga yetadi, tugagan — 0), "sug" (taklif:
        «Buyurtma» bo'limidagi need = sotuv × cover − qoldiq, yuqoriga yaxlitlangan, kamida 1), "key". Avval tugaganlar."""
        if not tags:
            return []
        bl, ibl = self.blacklist(), self.item_blacklist()
        period = float(self.data.period_days or 30)
        out = []
        for r in self.data._by_tags(self.data.sale, tags):
            firm = norm(r["firm"])
            if r["sales"] <= 0 or firm in bl or (firm, norm(r["name"])) in ibl:
                continue
            days = r["stock"] / (r["sales"] / period) if r["stock"] > 0 else 0
            if r["stock"] > 0 and days > self.ORD_DAYS:
                continue
            sug = max(1, int(math.ceil(r["sales"] * self.cover - r["stock"] - 1e-9)))
            out.append(dict(r, days=days, sug=sug, key=self.ord_key(r["name"], r["producer"])))
        out.sort(key=lambda r: (r["stock"] > 0, r["days"], -r["sales"], norm(r["name"])))
        return out

    def ord_basis(self, r):
        """«qoldiq 0 · tugagan · 30 kunda 46» | «qoldiq 8 · ≈2 kun · 30 kunda 156» (r — Sale qatori)."""
        period = self.data.period_days or 30
        if r["stock"] <= 0:
            head = "qoldiq 0 · tugagan"
        else:
            days = r["stock"] / (r["sales"] / float(period)) if r["sales"] > 0 else None
            head = "qoldiq %s · %s" % (qty(r["stock"]), ("≈%d kun" % int(days)) if days is not None else "sotuv yo'q")
        return "%s · %d kunda %s" % (head, period, qty(r["sales"]))

    def cart(self, cid):
        try:
            v = json.loads(self.setting("ocart_%d" % cid) or "[]")
        except ValueError:
            v = []
        return [i for i in v if isinstance(i, dict) and i.get("name") and to_float(i.get("qty")) > 0] if isinstance(v, list) else []

    def cart_save(self, cid, items):
        self.set_setting("ocart_%d" % cid, json.dumps(items, ensure_ascii=False))

    def ord_list_screen(self, cid, c, mid=None, page=0, note=None):
        """🏠 › 🧾 Buyurtma: tugagan / tugayotgan dorilar, har biri — savatga qo'shish / olish tugmasi."""
        self.data.reload()
        rows = self.oreq_low(self.tags_for(c))
        cart = self.cart(cid)
        inc = {self.ord_key(i["name"], i.get("producer")): i for i in cart}
        h = "🧾 <b>Buyurtma so'rovi</b> · %s" % esc(self.ord_label(c))
        kb = []
        if not rows:
            body = h + "\n\n✅ Tugagan yoki tugayotgan dori yo'q."
        else:
            per = self.ORD_PER
            pages = (len(rows) - 1) // per + 1
            page = page % pages
            part = rows[page * per:(page + 1) * per]
            lines = []
            for n, r in enumerate(part, page * per + 1):
                it = inc.get(r["key"])
                lines.append("%d. %s<b>%s</b> · %s ta\n<i>%s</i>" % (n, "✓ " if it else "", esc(r["name"]), qty(it["qty"] if it else r["sug"]),
                                                                  self.ord_basis(r)))
                kb.append([(("✓ " if it else "➕ ") + bl(self.short(r["name"], 32)), "c:b:t:%d:%s:%d" % (n - 1, self._ck(r["key"]), page))])
            body = "%s\nTugagan va tugayotgan: <b>%d</b> ta dori. Nom yonida — taklif soni.\n\n%s" % (h, len(rows), "\n".join(lines))
            if pages > 1:
                kb.append([("◀", "c:b:l:%d" % ((page - 1) % pages)), ("%d / %d" % (page + 1, pages), "c:b:l:%d" % page),
                           ("▶", "c:b:l:%d" % ((page + 1) % pages))])
            body += "\n\n%d–%d / %d · 🕒 %s" % (page * per + 1, page * per + len(part), len(rows), dt_text(self.data.when()))
        body += ("\n🛒 Savatda: <b>%d</b> ta dori · <b>%s</b> ta" % (len(cart), qty(sum(to_float(i["qty"]) for i in cart)))
                 if cart else "\n🛒 Savat bo'sh — dorini bosing, taklif soni bilan savatga tushadi.")
        kb.append([("🛒 Savat · %d" % len(cart), "c:b:c"), ("📋 So'rovlarim", "c:b:m")])
        kb.append([("📄 Buyurtma tavsiyasi (Excel)", "x:need")])
        return self.panel_show(cid, mid, *self.screen(("🧾 Buyurtma",), body, kb, back="c:home", home="c:home"), note=note)

    def ord_toggle(self, cid, c, mid, i, ck, page):
        rows = self.oreq_low(self.tags_for(c))
        r = rows[i] if 0 <= i < len(rows) else None
        if not r or self._ck(r["key"]) != ck:
            return self.ord_list_screen(cid, c, mid, page, note="⚠ Ro'yxat yangilandi, qaytadan tanlang.")
        cart = self.cart(cid)
        rest = [x for x in cart if self.ord_key(x["name"], x.get("producer")) != r["key"]]
        if len(rest) == len(cart):
            if len(cart) >= self.ORD_CART_MAX:
                return self.ord_list_screen(cid, c, mid, page, note="⚠ Savat to'ldi: ko'pi bilan %d ta dori. Avval «📤 Adminga yuborish» "
                                            "ni bosing yoki savatdan olib tashlang." % self.ORD_CART_MAX)
            rest.append({"name": r["name"], "producer": r["producer"], "qty": r["sug"]})
        self.cart_save(cid, rest)
        return self.ord_list_screen(cid, c, mid, page)

    def ord_cart_screen(self, cid, c, mid=None, sel=0, note=None):
        """🛒 Savat: dori tanlanadi (▸), soni −5 / −1 / +1 / +5 yoki yozib o'zgartiriladi."""
        cart = self.cart(cid)
        if not cart:
            body = "🛒 <b>Savat bo'sh</b>\nRo'yxatdan dorini bosing — taklif soni bilan savatga tushadi."
            kb = [[("➕ Dori qo'shish", "c:b:l:0")]]
        else:
            sel = min(max(sel, 0), len(cart) - 1)
            cy = UK.lat2cyr
            w = min(24, max(len(self.short(i["name"], 24)) for i in cart))
            lines = ["%s%d %s %s" % ("▸" if n == sel else " ", n + 1, self.short(i["name"], 24).ljust(w), (qty(to_float(i["qty"])) + " " +
                                                                                                           cy("ta")).rjust(9))
                     for n, i in enumerate(cart)]
            lines.append("  " + cy("Jami").ljust(w + len(str(len(cart)))) + (qty(sum(to_float(i["qty"]) for i in cart)) + " " + cy("ta")).rjust(10))
            it = cart[sel]
            r = self.sale_index().get((norm(it["name"]), norm(it.get("producer"))))
            body = ("🛒 <b>Savat</b> · %d ta dori\nDorini tanlang, keyin sonini o'zgartiring.\n<pre>%s</pre>\n▸ Tanlandi: <b>%s</b>%s" % (
                len(cart), esc("\n".join(lines)), esc(it["name"]), ("\n<i>%s</i>" % self.ord_basis(r)) if r else ""))
            per = self.ORD_BTN_PER  # tugmalar sahifalanadi (tanlangan dori turgan sahifa): Telegram ~100 tugmadan ko'pini olmaydi
            page, pages = sel // per, (len(cart) - 1) // per + 1
            btns = [(("▸ " if n == sel else "") + "%d · %s" % (n + 1, bl(self.short(i["name"], 12))), "c:b:c:%d" % n)
                    for n, i in enumerate(cart) if n // per == page]
            kb = [btns[k:k + 3] for k in range(0, len(btns), 3)]
            if pages > 1:
                kb.append([("◀", "c:b:c:%d" % (((page - 1) % pages) * per)), ("%d / %d" % (page + 1, pages), "c:b:c:%d" % sel),
                           ("▶", "c:b:c:%d" % (((page + 1) % pages) * per))])
            kb.append([(t, "c:b:q:%d:%d" % (sel, d)) for t, d in self.ORD_STEPS])
            kb.append([("✏️ Son yozish", "c:b:w:%d" % sel), ("🗑 Olib tashlash", "c:b:r:%d" % sel)])
            kb.append([("📤 Adminga yuborish", "c:b:send", "primary")])
            kb.append([("➕ Dori qo'shish", "c:b:l:0"), ("🗑 Tozalash", "c:b:clr")])
        return self.panel_show(cid, mid, *self.screen(("🧾 Buyurtma", "🛒 Savat"), body, kb, back="c:need", home="c:home"), note=note)

    def ord_set_qty(self, cid, pos, n=None, delta=0):
        """Savatdagi pos-dori soni: n (yozilgan) yoki delta. Yangi son yoki None (pos yo'q)."""
        cart = self.cart(cid)
        if not 0 <= pos < len(cart):
            return None
        v = to_float(cart[pos]["qty"]) + delta if n is None else n
        cart[pos]["qty"] = int(min(max(1, v), self.ORD_MAX))
        self.cart_save(cid, cart)
        return cart[pos]["qty"]

    @classmethod
    def ord_int(cls, text):
        """Yozilgan son → int yoki None. Faqat butun raqamlar (bo'shliqlar olinadi): «nan», «inf», «2.5», «-1», «1e3» — None."""
        s = re.sub(r"[\s  ]", "", str(text or ""))
        return int(s) if cls.ORD_INT.match(s) else None

    def ord_qty_input(self, cid, c, st, text):
        pos = st[1]
        cart = self.cart(cid)
        if not 0 <= pos < len(cart):
            self.state.pop(cid, None)
            return self.ord_cart_screen(cid, c, None)
        n = self.ord_int(text)
        if n is None or not 1 <= n <= self.ORD_MAX:
            return self.ask(cid, None, ("🧾 Buyurtma", "🛒 Savat", "✏️ Son"), "Butun son yozing: 1 dan %d gacha. Bekor qilish: /cancel" %
                            self.ORD_MAX, st, back="c:b:c:%d" % pos)
        self.state.pop(cid, None)
        self.ord_set_qty(cid, pos, n)
        return self.ord_cart_screen(cid, c, None, pos, note="✅ %s — %s ta" % (esc(cart[pos]["name"]), qty(n)))

    def ord_send(self, cid, c, mid):
        """📤 Adminga yuborish: faqat o'z firmasining dorilari (savat saqlangandan keyin firma o'zgargan bo'lsa ham)."""
        own = {(norm(r["name"]), norm(r["producer"])) for r in self.data._by_tags(self.data.sale, self.tags_for(c))}
        items = [{"name": i["name"], "producer": i.get("producer") or "", "qty": int(to_float(i["qty"]))} for i in self.cart(cid)
                 if (norm(i["name"]), norm(i.get("producer"))) in own]
        if not items:
            self.cart_save(cid, [])
            return self.ord_cart_screen(cid, c, mid, note="⚠ Savatda yuboriladigan dori yo'q.")
        with self.dblock:
            rid = BY.create(self.db, cid, self.ord_label(c), c["name"] or c["phone"] or str(cid), items,
                            ts=now().isoformat(timespec="seconds"))
        self.cart_save(cid, [])
        self.touch(cid, "order")
        logging.info("Buyurtma so'rovi #%d: %s · %d ta dori", rid, self.ord_label(c), len(items))
        try:
            self.ord_push_admins(rid)
        except Exception:  # noqa
            logging.exception("Buyurtma so'rovi #%d: adminlarga karta yuborib bo'lmadi", rid)
        w = min(26, max(len(self.short(i["name"], 26)) for i in items))
        lines = "\n".join(self.short(i["name"], 26).ljust(w) + (qty(i["qty"]) + " " + UK.lat2cyr("ta")).rjust(9) for i in items)
        body = ("⏳ <b>So'rov №%d yuborildi</b> · %s\n%d ta dori · jami <b>%s</b> ta\n<pre>%s</pre>\nAdmin javob berganda shu yerga xabar "
                "keladi.\nHolati: 📋 So'rovlarim.") % (rid, esc(self.ord_label(c)), len(items), qty(sum(i["qty"] for i in items)), esc(lines))
        return self.panel_show(cid, mid, *self.screen(("🧾 Buyurtma", "📤 Yuborildi"), body,
                                                      [[("📋 So'rovlarim", "c:b:m")], [("🏠 Menyu", "c:home")]]))

    @staticmethod
    def ord_sym(r, items):
        if r["status"] == BY.OPEN:
            return "⏳"
        asked, ok, changed = BY.totals(items)
        if r["status"] == BY.REJECTED or not ok:
            return "❌"
        return "✏️" if changed or any(i["status"] == BY.I_REJ for i in items) else "✅"

    def ord_mine_screen(self, cid, c, mid=None, note=None):
        """📋 So'rovlarim: oxirgi 30 kun, yangisi yuqorida, holati bilan."""
        since = (now() - datetime.timedelta(days=30)).isoformat(timespec="seconds")
        with self.dblock:
            reqs = [(r, BY.get(self.db, r["id"])[1]) for r in BY.mine(self.db, cid, 50) if (r["created_at"] or "") >= since]
        h = "📋 <b>So'rovlarim</b> · %s" % esc(self.ord_label(c))
        if not reqs:
            body = h + "\n\nOxirgi 30 kunda so'rov yo'q."
        else:
            cy = UK.lat2cyr
            lines = ["№".ljust(5) + cy("Sana").ljust(6) + cy("Dori").rjust(5) + cy("So'raldi").rjust(9) + cy("Qabul").rjust(7)]
            for r, items in reqs:
                asked, ok, _ = BY.totals(items)
                ca = r["created_at"] or ""
                lines.append(str(r["id"]).ljust(5) + ("%s.%s" % (ca[8:10], ca[5:7])).ljust(6) + str(len(items)).rjust(5) +
                             qty(asked).rjust(9) + ("—" if r["status"] == BY.OPEN else qty(ok)).rjust(7) + " " + self.ord_sym(r, items))
            body = "%s\nOxirgi 30 kun · <b>%d</b> ta so'rov\n<pre>%s</pre>\n<i>⏳ kutmoqda · ✅ qabul · ✏️ o'zgartirildi · ❌ rad</i>" % (
                h, len(reqs), esc("\n".join(lines)))
        kb = [[("📥 Excel'da olish", "c:b:mx")]] if reqs else []
        return self.panel_show(cid, mid, *self.screen(("🧾 Buyurtma", "📋 So'rovlarim"), body + self.foot(), kb, back="c:need",
                                                      home="c:home"), note=note)

    def ord_mine_excel(self, cid, mid):
        """Faqat shu chatning so'rovlari (dori bo'yicha qatorlar)."""
        with self.dblock:
            reqs = [(r, BY.get(self.db, r["id"])[1]) for r in BY.mine(self.db, cid, 500)]
        rows = [[r["id"], (r["created_at"] or "")[:16].replace("T", " "), i["name"], i["producer"] or "", i["qty_asked"],
                 i["qty_ok"] if r["status"] != BY.OPEN else None, UK.lat2cyr(self.ORD_ITEM_LBL.get(i["status"], i["status"])),
                 UK.lat2cyr({BY.OPEN: "kutmoqda", BY.ANSWERED: "javob berildi", BY.REJECTED: "rad etildi"}.get(r["status"], r["status"]))]
                for r, items in reqs for i in items]
        if not rows:
            return self.panel_note(cid, mid, "➖ So'rov yo'q.")
        content = xlsx_bytes("So'rovlarim", ["№", "Sana", "Dori", "Ishlab chiqaruvchi", "So'raldi", "Tasdiqlandi", "Holat", "So'rov holati"],
                             rows, [7, 17, 44, 24, 10, 12, 15, 16], numcols=(4, 5))
        self.touch(cid, "excel:orders")
        return self.panel_doc(cid, "Sorovlarim_%s.xlsx" % self._today(), content, "📋 <b>So'rovlarim</b> · %d ta qator" % len(rows))

    def on_order_client(self, cid, c, mid, p):
        """c:b:<amal>:… — faqat faol firma vakili (o'z firmasi; savat — o'z chatida)."""
        if c["role"] != "firm":
            return None
        sub = p[0] if p else "l"
        num = lambda k, d=0: int(p[k]) if len(p) > k and re.match(r"^-?\d+$", p[k]) else d
        if sub == "l":
            return self.ord_list_screen(cid, c, mid, num(1))
        if sub == "t":
            return self.ord_toggle(cid, c, mid, num(1, -1), p[2] if len(p) > 2 else "", num(3))
        if sub == "c":
            return self.ord_cart_screen(cid, c, mid, num(1))
        if sub == "q":
            if num(2) not in [d for _, d in self.ORD_STEPS]:
                return None
            self.ord_set_qty(cid, num(1, -1), delta=num(2))
            return self.ord_cart_screen(cid, c, mid, num(1))
        if sub == "w":
            cart, pos = self.cart(cid), num(1, -1)
            if not 0 <= pos < len(cart):
                return self.ord_cart_screen(cid, c, mid)
            return self.ask(cid, mid, ("🧾 Buyurtma", "🛒 Savat", "✏️ Son"), "✏️ <b>%s</b>\nHozir: %s ta. Nechta kerak? Son yozing 👇 (1–%d).\n"
                            "Bekor qilish: /cancel" % (esc(cart[pos]["name"]), qty(to_float(cart[pos]["qty"])), self.ORD_MAX),
                            ("ord_qty", pos), back="c:b:c:%d" % pos)
        if sub == "r":
            cart, pos = self.cart(cid), num(1, -1)
            if 0 <= pos < len(cart):
                gone = cart.pop(pos)
                self.cart_save(cid, cart)
                return self.ord_cart_screen(cid, c, mid, pos, note="🗑 Olib tashlandi: %s" % esc(gone["name"]))
            return self.ord_cart_screen(cid, c, mid)
        if sub == "clr":
            self.cart_save(cid, [])
            return self.ord_cart_screen(cid, c, mid, note="🗑 Savat tozalandi.")
        if sub == "send":
            return self.ord_send(cid, c, mid)
        if sub == "m":
            self.touch(cid, "orders")
            return self.ord_mine_screen(cid, c, mid)
        if sub == "mx":
            return self.ord_mine_excel(cid, mid)
        return None

    # ---- admin: karta ----
    def admin_name(self, cid):
        return self._anames.get(cid) or "Admin"

    def ord_closed_line(self, r, items, admin=None):
        """Javob berilgan so'rov — bir qator: «✅ №27 · FIRMA · 3 ta dori · 75 → 65 ta · 1 ta o'zgartirildi · Siz, 11:05»."""
        asked, ok, changed = BY.totals(items)
        who = "Siz" if admin is not None and r["resolved_by"] == admin else esc(r["resolved_name"] or "admin")
        tm = (r["resolved_at"] or "")[11:16]
        if r["status"] == BY.REJECTED:
            return "❌ №%d · %s · %d ta dori · rad etildi · %s, %s" % (r["id"], esc(r["firm"] or ""), len(items), who, tm)
        return "✅ №%d · %s · %d ta dori · %s → %s ta%s · %s, %s" % (r["id"], esc(r["firm"] or ""), len(items), qty(asked), qty(ok),
                                                                 (" · %d ta o'zgartirildi" % changed) if changed else "", who, tm)

    def ord_page(self, pos):
        """pos-dori turgan sahifa (admin kartasi)."""
        return max(0, (int(pos or 1) - 1) // self.ORD_BTN_PER)

    def ord_card(self, rid, panel=False, page=0):
        """So'rov kartasi (umumiy ko'rinish) → (matn, tugmalar). panel=True — admin panelida (yo'l qatori, ◀ Orqaga).
        Dorilar ORD_BTN_PER tadan sahifalanadi (◀ ▶): matnda ham, tugmalarda ham faqat shu sahifadagilar."""
        r, items = BY.get(self.db, rid)
        idx = self.sale_index()
        decided = sum(1 for i in items if i["status"] != BY.I_WAIT)
        per = self.ORD_BTN_PER
        pages = max(1, (len(items) - 1) // per + 1)
        page = min(max(0, page), pages - 1)
        part = items[page * per:(page + 1) * per]
        lines = []
        for i in part:
            a, ok = qty(i["qty_asked"]), qty(i["qty_ok"] or 0)
            tail = {BY.I_WAIT: "%s ta" % a, BY.I_OK: "%s ta ✅" % a, BY.I_CHANGED: "%s → %s ta ✏️" % (a, ok),
                    BY.I_REJ: "%s ta ❌" % a}.get(i["status"], a)
            s = idx.get((norm(i["name"]), norm(i["producer"])))
            lines.append("%d. <b>%s</b> — %s%s" % (i["pos"], esc(i["name"]), tail, ("\n<i>%s</i>" % self.ord_basis(s)) if s else ""))
        asked, ok, changed = BY.totals(items)
        if BY.all_decided(items):
            tail = "Jami: <b>%s</b> ta (%s dan) · %d ta o'zgartirildi\nHammasi hal bo'ldi — javobni yuboring." % (qty(ok), qty(asked), changed)
        else:
            tail = "Jami: <b>%s</b> ta · %d ta dori\nDorini bosing — sonini o'zgartirish yoki olib tashlash." % (qty(asked), len(items))
        text = "🧾 <b>Buyurtma so'rovi №%d</b> · %s\n👤 %s · %s\n\n%s\n\n%s" % (
            rid, esc(r["firm"] or ""), esc(r["name"] or ""), ("hal qilindi %d / %d" % (decided, len(items))) if decided else "firma vakili",
            ("Dorilar %d–%d / %d:\n" % (page * per + 1, page * per + len(part), len(items)) if pages > 1 else "") + "\n".join(lines), tail)
        kb = [[("📤 Javobni yuborish", "oq:%d:send" % rid, "primary")]] if BY.all_decided(items) else [
            [("✅ Hammasini qabul qilish", "oq:%d:all" % rid, "success")]]
        for i in part:
            done = i["status"] in (BY.I_OK, BY.I_CHANGED)
            mark = "✓ " if done else "✕ " if i["status"] == BY.I_REJ else ""
            kb.append([("%s%d · %s · %s ta" % (mark, i["pos"], bl(self.short(i["name"], 22)), qty(i["qty_ok"] if done else i["qty_asked"])),
                        "oq:%d:i:%d" % (rid, i["pos"]))])
        if pages > 1:
            kb.append([("◀", "oq:%d:p:%d" % (rid, (page - 1) % pages)), ("%d / %d" % (page + 1, pages), "oq:%d:p:%d" % (rid, page)),
                       ("▶", "oq:%d:p:%d" % (rid, (page + 1) % pages))])
        kb.append([("❌ Rad etish", "oq:%d:rad" % rid, "danger")])
        if panel:
            return self.screen(("🧾 Buyurtma so'rovlari", "№%d" % rid), text, kb, back="oq:l", home="a:home")
        kb.append([("🏠 Menyuga qaytish", "a:home")])
        return cut_html(text), inline_kb(kb)

    def ord_item_card(self, rid, pos, v=None, panel=False):
        """Bitta dori: «… so'raldi 20 ta. Rozimisiz?» + filiallar bo'yicha qoldiq / sotuv. v — taklif etilayotgan son (−5 / +5)."""
        r, items = BY.get(self.db, rid)
        it = next((i for i in items if i["pos"] == pos), None)
        if not it:
            return None
        asked = it["qty_asked"]
        if v is None:
            v = it["qty_ok"] if it["status"] in (BY.I_OK, BY.I_CHANGED) else asked
        v = min(max(1, v), self.ORD_MAX)
        s = self.sale_index().get((norm(it["name"]), norm(it["producer"])))
        text = "🧾 №%d · %s › %d / %d\n<b>%s</b> · so'raldi <b>%s</b> ta. Rozimisiz?" % (
            rid, esc(r["firm"] or ""), pos, len(items), esc(it["name"]), qty(asked))
        if it["status"] != BY.I_WAIT:
            text += "\nHozir: %s" % ("olib tashlangan ❌" if it["status"] == BY.I_REJ else "%s ta ✓" % qty(it["qty_ok"] or 0))
        if s:
            cy, names = UK.lat2cyr, self.branch_names()
            w = max(len(x) for x in names + [cy("Jami"), cy("Filial")]) + 1
            lines = [cy("Filial").ljust(w) + cy("Qoldiq").rjust(7) + cy("Sotuv").rjust(7)]
            for n, (si, bi) in zip(names[:3], ((1, 0), (3, 2), (5, 4))):
                lines.append(n.ljust(w) + qty(s["v"][bi]).rjust(7) + qty(s["v"][si]).rjust(7))
            lines.append(names[3].ljust(w) + qty(s["v"][6]).rjust(7) + "—".rjust(7))
            lines.append(cy("Jami").ljust(w) + qty(s["stock"]).rjust(7) + qty(s["sales"]).rjust(7))
            text += "\n<i>%s</i>\n<pre>%s</pre>" % (self.ord_basis(s), esc("\n".join(lines)))
            daily = s["sales"] / float(self.data.period_days or 30)
            if daily > 0:
                text += "\n%s ta ≈ %d kunga yetadi" % (qty(v), int((v + max(0, s["stock"])) / daily))
        half, three = max(1, int(math.ceil(asked / 2.0))), max(1, int(math.ceil(asked * 3 / 4.0)))
        cbq = lambda n: "oq:%d:s:%d:%s" % (rid, pos, qty(n))
        kb = [[("✅ Ha, %s ta" % qty(v), cbq(v), "success")],
              [("½ · %s" % qty(half), cbq(half)), ("¾ · %s" % qty(three), cbq(three)),
               ("−5", "oq:%d:i:%d:%s" % (rid, pos, qty(max(1, v - 5)))), ("+5", "oq:%d:i:%d:%s" % (rid, pos, qty(min(self.ORD_MAX, v + 5))))],
              [("✏️ Sonini yozish", "oq:%d:w:%d" % (rid, pos))], [("❌ Olib tashlash", "oq:%d:x:%d" % (rid, pos), "danger")]]
        back = ("oq:%d:p:%d" % (rid, self.ord_page(pos))) if len(items) > self.ORD_BTN_PER else "oq:%d" % rid
        if panel:
            return self.screen(("🧾 Buyurtma so'rovlari", "№%d" % rid, "%d / %d" % (pos, len(items))), text, kb,
                               back=back, home="a:home")
        kb.append([("◀ Orqaga", back)])
        return cut_html(text), inline_kb(kb)

    def ord_view(self, cid, mid, scr, note=None, resend=False):
        """Karta panelda ochilgan bo'lsa — panel; bildirishnoma (alohida karta) bo'lsa — shu xabarning o'zi o'zgaradi."""
        if resend:
            return self.panel_resend(cid, scr[0], scr[1], note)
        if self.is_panel(cid, mid):
            return self.panel_show(cid, mid, scr[0], scr[1], note)
        return self.show(cid, mid, self.with_note(scr[0], note), scr[1])

    def ord_push_admins(self, rid):
        """Yangi so'rov: har adminga alohida karta (NAVIGATSIYA.md, 5-qoida); message_id lar — kartani keyin qisqartirish uchun."""
        text, kb = self.ord_card(rid)  # sahifalangan: tugmalar soni dorilar soniga bog'liq emas
        for a in self.admins:
            try:
                r = self.tg.send(a, text, kb)
            except Exception:  # noqa
                logging.exception("Adminga buyurtma so'rovi kartasini yuborib bo'lmadi: %s", a)
                try:  # zaxira: qisqa xabar, karta tugma orqali ochiladi
                    r = self.tg.send(a, "🧾 <b>Yangi buyurtma so'rovi №%d</b>\nKartani ochish uchun bosing." % rid,
                                     inline_kb([[("🧾 Ochish", "oq:%d" % rid)]]))
                except Exception:  # noqa
                    logging.exception("Adminga buyurtma so'rovi xabarini yuborib bo'lmadi: %s", a)
                    continue
            m = r.get("message_id") if isinstance(r, dict) else None
            if m:
                with self.dblock:
                    self.db.execute("INSERT OR REPLACE INTO order_req_msgs VALUES(?,?,?)", (rid, a, m))
                    self.db.commit()

    def ord_admin_list(self, cid, mid=None, note=None):
        """🏠 › 🧾 Buyurtma so'rovlari: kutayotganlar (tugma — karta) va oxirgi 7 kunda hal qilinganlar."""
        with self.dblock:
            opened = [(r, BY.get(self.db, r["id"])[1]) for r in BY.open_requests(self.db)]
            done = [(r, BY.get(self.db, r["id"])[1]) for r in BY.recent(self.db, 7, now()) if r["status"] != BY.OPEN]
        ts = lambda v: "%s.%s %s" % ((v or "")[8:10], (v or "")[5:7], (v or "")[11:16])
        text = "🧾 <b>Buyurtma so'rovlari</b> · oxirgi 7 kun\nKutayotganini bosing — karta ochiladi."
        text += "\n\n⏳ <b>Kutmoqda</b> · %d" % len(opened) + "".join(
            "\n№%d · %s · %d ta dori\n<i>%s ta · %s</i>" % (r["id"], esc(self.short(r["firm"], 30)), len(items), qty(BY.totals(items)[0]),
                                                         ts(r["created_at"])) for r, items in opened[:15])
        if not opened:
            text += "\nYangi so'rov yo'q."
        if done:
            text += "\n\n<b>Hal qilindi</b> · %d" % len(done) + "".join(
                "\n%s №%d · %s · %s → %s ta\n<i>%s, %s</i>" % (self.ord_sym(r, items), r["id"], esc(self.short(r["firm"], 30)),
                                                            qty(BY.totals(items)[0]), qty(BY.totals(items)[1]),
                                                            "Siz" if r["resolved_by"] == cid else esc(r["resolved_name"] or "admin"),
                                                            ts(r["resolved_at"])) for r, items in done[:10])
        kb = [[("⏳ №%d · %s" % (r["id"], bl(self.short(r["firm"], 26))), "oq:%d" % r["id"])] for r, _ in opened[:15]]
        kb.append([("📥 Tasdiqlanganlar Excel", "oq:lx")])
        return self.a_screen(cid, mid, ("🧾 Buyurtma so'rovlari",), text + self.foot(), kb, note=note)

    def ord_admin_excel(self, cid, mid, days=7):
        since = (now() - datetime.timedelta(days=days)).isoformat(timespec="seconds")
        with self.dblock:
            rows = BY.approved_rows(self.db, since)
        if not rows:
            return self.panel_note(cid, mid, "➖ Oxirgi %d kunda tasdiqlangan so'rov yo'q." % days)
        data = [[r["id"], r["firm"] or "", r["who"] or "", (r["resolved_at"] or "")[:16].replace("T", " "), r["name"], r["producer"] or "",
                 r["qty_asked"], r["qty_ok"], UK.lat2cyr(self.ORD_ITEM_LBL.get(r["status"], r["status"]))] for r in rows]
        content = xlsx_bytes("Buyurtma so'rovlari", ["№", "Firma", "Vakil", "Javob vaqti", "Dori", "Ishlab chiqaruvchi", "So'ralgan",
                                                     "Tasdiqlangan", "Holat"], data, [7, 26, 22, 17, 44, 22, 11, 13, 15], numcols=(6, 7))
        return self.panel_doc(cid, "Buyurtma_sorovlari_%dkun_%s.xlsx" % (days, self._today()), content,
                              "🧾 <b>Tasdiqlangan buyurtma so'rovlari</b> · oxirgi %d kun · %d qator" % (days, len(data)))

    def on_order_admin(self, cid, mid, p):
        """oq:l | oq:lx | oq:<id>[:i:<pos>[:<son>] | :s:<pos>:<son> | :x:<pos> | :w:<pos> | :all | :rad | :rady | :send]. Faqat admin."""
        if len(p) < 2:
            return None
        if p[1] == "l":
            return self.ord_admin_list(cid, mid)
        if p[1] == "lx":
            return self.ord_admin_excel(cid, mid)
        rid = int(p[1])
        sub = p[2] if len(p) > 2 else ""
        panel = self.is_panel(cid, mid)
        with self.dblock:
            r, items = BY.get(self.db, rid)
        if not r:
            return self.panel_note(cid, mid, "So'rov topilmadi.")
        if r["status"] != BY.OPEN:  # boshqa admin (yoki panel) allaqachon javob bergan: karta qisqaradi
            if panel:
                return self.ord_admin_list(cid, mid, "ℹ️ " + self.ord_closed_line(r, items, cid))
            return self.show(cid, mid, self.ord_closed_line(r, items, cid), None)
        pos = int(p[3]) if len(p) > 3 and p[3].isdigit() else None
        note = None
        page = self.ord_page(pos) if sub in ("s", "x") and pos is not None else 0
        if sub == "p":  # kartaning sahifasi (◀ ▶)
            return self.ord_view(cid, mid, self.ord_card(rid, panel, pos or 0))
        if sub == "i" and pos is not None:
            scr = self.ord_item_card(rid, pos, self.ord_int(p[4]) if len(p) > 4 else None, panel)
            return self.ord_view(cid, mid, scr) if scr else None
        if sub == "w" and pos is not None:
            it = next((i for i in items if i["pos"] == pos), None)
            if not it:
                return None
            self.state[cid] = ("oreq_qty", (rid, pos, mid))
            body = ("🧾 №%d · %s › %d / %d\n✏️ <b>%s</b> · so'raldi <b>%s</b> ta\nNechta tasdiqlaysiz? Son yozing 👇 (0 — olib tashlash).\n"
                    "Bekor qilish: /cancel") % (rid, esc(r["firm"] or ""), pos, len(items), esc(it["name"]), qty(it["qty_asked"]))
            kb = [[("◀ Orqaga", "oq:%d:i:%d" % (rid, pos))]]
            return self.ord_view(cid, mid, self.screen(("🧾 Buyurtma so'rovlari", "№%d" % rid, "✏️"), body, kb, back="oq:%d:i:%d" % (rid, pos),
                                                       home="a:home") if panel else (body, inline_kb(kb)))
        if sub in ("s", "x", "all"):
            n = self.ord_int(p[4]) if sub == "s" and len(p) > 4 else None
            if (sub != "all" and pos is None) or (sub == "s" and n is None):
                return None
            with self.dblock:
                if sub == "all":
                    ok = BY.accept_all(self.db, rid)
                else:
                    ok = BY.decide_item(self.db, rid, pos, min(n, self.ORD_MAX)) if sub == "s" else BY.decide_item(
                        self.db, rid, pos, reject=True)
            if not ok:  # o'qish va qaror orasida boshqa admin / panel yopgan bo'lishi mumkin: ochiq karta qayta chizilmaydi
                with self.dblock:
                    r, items = BY.get(self.db, rid)
                if not r or r["status"] != BY.OPEN:
                    return self.on_order_admin(cid, mid, p[:2])
                note = "⚠ Dori topilmadi."
        elif sub == "rad":
            body = "🧾 №%d · %s\n❌ <b>Butun so'rov rad etilsinmi?</b>\n%s · %d ta dori. Firma vakiliga «rad etildi» xabari boradi." % (
                rid, esc(r["firm"] or ""), esc(r["name"] or ""), len(items))
            kb = [[("❌ Ha, rad etish", "oq:%d:rady" % rid, "danger")], [("◀ Orqaga", "oq:%d" % rid)]]
            return self.ord_view(cid, mid, self.screen(("🧾 Buyurtma so'rovlari", "№%d" % rid, "❌"), body, kb[:1], back="oq:%d" % rid,
                                                       home="a:home") if panel else (body, inline_kb(kb)))
        elif sub in ("send", "rady"):
            with self.dblock:
                ok = BY.finish(self.db, rid, cid, self.admin_name(cid), reject=sub == "rady")
            if not ok:
                with self.dblock:
                    r, items = BY.get(self.db, rid)
                if r["status"] != BY.OPEN:
                    return self.on_order_admin(cid, mid, p[:2])
                note = "⚠ Avval har dori bo'yicha qaror qiling (✅ Ha, son yoki ❌ olib tashlash)."
            else:
                self.order_notify(rid, by=cid)
                done = "📨 <b>Javob yuborildi</b>\n%s (%s) oldi" if sub == "send" else "❌ <b>So'rov rad etildi</b>\n%s (%s) ga xabar ketdi"
                done %= (esc(r["name"] or ""), esc(r["firm"] or ""))
                if panel:
                    return self.ord_admin_list(cid, mid, done)
                return self.send_admin_panel(cid, None, done)  # karta bir qatorga qisqardi, panel pastga tushadi
        return self.ord_view(cid, mid, self.ord_card(rid, panel, page), note)

    def ord_qty_admin_input(self, cid, st, text):
        """Admin yozgan son (karta «✏️ Sonini yozish»)."""
        rid, pos, mid = st[1]
        n = self.ord_int(text)
        with self.dblock:
            r, items = BY.get(self.db, rid)
        if not r or r["status"] != BY.OPEN:
            self.state.pop(cid, None)
            return self.on_order_admin(cid, mid, ["oq", str(rid)])
        if n is None or not 0 <= n <= self.ORD_MAX:
            self.state[cid] = st
            return self.ord_view(cid, mid, self.ord_card(rid, self.is_panel(cid, mid), self.ord_page(pos)),
                                 "⚠ Butun son yozing: 0 dan %d gacha (0 — olib tashlash). Yoki /cancel." % self.ORD_MAX)
        self.state.pop(cid, None)
        with self.dblock:
            ok = BY.decide_item(self.db, rid, pos, n, reject=n == 0)
        if not ok:  # shu orada boshqa admin / panel yopgan: yopilgan holat (bir qator)
            return self.on_order_admin(cid, mid, ["oq", str(rid)])
        panel = self.is_panel(cid, mid)
        return self.ord_view(cid, mid, self.ord_card(rid, panel, self.ord_page(pos)), "✅ Saqlandi: %s ta" % qty(n), resend=panel)

    def order_notify(self, rid, by=None):
        """So'rovga javob berildi (bot yoki server paneli): firma vakiliga natija (push — yangi panel bo'ladi), hamma adminlarda
        karta bir qatorga qisqaradi. by — javob bergan admin (uning kartasida «Siz»). Har so'rov uchun bir marta."""
        with self.dblock:
            r, items = BY.get(self.db, rid)
            if not r or r["status"] == BY.OPEN or self.setting("onotify_%d" % rid):
                return False
            self.db.execute("INSERT OR REPLACE INTO settings VALUES(?,?)", ("onotify_%d" % rid, now().isoformat(timespec="seconds")))
            self.db.commit()
            msgs = self.db.execute("SELECT chat_id, message_id FROM order_req_msgs WHERE req_id=?", (rid,)).fetchall()
        asked, ok, changed = BY.totals(items)
        when = "%s %s" % ("bugun" if (r["resolved_at"] or "")[:10] == self._today() else (r["resolved_at"] or "")[:10],
                          (r["resolved_at"] or "")[11:16])
        if r["status"] == BY.REJECTED:
            text = "❌ <b>So'rov №%d rad etildi</b> · %s\nAdmin ko'rib chiqdi · %s\n%d ta dori · %s ta so'ralgan edi." % (
                rid, esc(r["firm"] or ""), when, len(items), qty(asked))
        else:
            cy = UK.lat2cyr
            w = min(22, max(len(self.short(i["name"], 22)) for i in items))
            sym = {BY.I_OK: "✅", BY.I_CHANGED: "✏️", BY.I_REJ: "❌"}
            lines = [" " * w + cy("So'raldi").rjust(9) + "   " + cy("Qabul")]
            for i in items:
                lines.append(self.short(i["name"], 22).ljust(w) + qty(i["qty_asked"]).rjust(9) + " → " +
                             qty(i["qty_ok"] or 0).ljust(5) + " " + sym.get(i["status"], ""))
            lines.append(cy("Jami").ljust(w) + qty(asked).rjust(9) + " → " + qty(ok))
            text = "✅ <b>So'rov №%d bo'yicha javob</b> · %s\nAdmin ko'rib chiqdi · %s\n<pre>%s</pre>" % (
                rid, esc(r["firm"] or ""), when, esc("\n".join(lines)))
            for i in items:
                if i["status"] == BY.I_CHANGED:
                    text += "\n✏️ %s: %s emas, <b>%s</b> ta qabul qilindi." % (esc(self.short(i["name"], 30)), qty(i["qty_asked"]),
                                                                            qty(i["qty_ok"] or 0))
                elif i["status"] == BY.I_REJ:
                    text += "\n❌ %s: qabul qilinmadi." % esc(self.short(i["name"], 30))
            text += "\nHamma so'rov — 📋 So'rovlarim va Excel'da."
        c = self.client(r["chat_id"])
        if c and c["status"] == "active":
            try:
                self.panel_resend(r["chat_id"], cut_html(text), inline_kb([[("📋 So'rovlarim", "c:b:m"), ("🏠 Menyu", "c:home")]]))
            except Exception as e:  # noqa
                if not self.mark_gone(r["chat_id"], e):
                    logging.exception("Buyurtma so'rovi #%d: vakilga javob yuborib bo'lmadi", rid)
        for m in msgs:
            try:
                self.tg.edit(m["chat_id"], m["message_id"], self.ord_closed_line(r, items, m["chat_id"]), None)
            except Exception:  # noqa
                pass
        logging.info("Buyurtma so'rovi #%d javobi: %s · %s → %s (%s)", rid, r["status"], qty(asked), qty(ok), r["resolved_name"])
        return True

    # ================= filial xodimi paneli: ko'chirish nazorati («Berdim / Oldim») =================
    # Toifalar, matnlar, ranglar — faqat KOCH_CATS dan (modul boshida). Ro'yxat / kartochka / Excel — umumiy funksiyalar.
    ST_NAME = {"yangi": "⏳ yangi", "berildi": "📦 berildi", "olindi": "✅ olindi", "rad": "❌ rad", "bekor": "🚫 bekor",
               "yopildi": "⌛ yopildi"}
    RAD_REASONS = ("Tovar topilmadi", "Muddati yaqin, buzilgan", "Boshqa")
    WHY_QUEUE = "Navbatda: kunlik limitdan tashqari"
    # filial callback prefikslari (boshqa prefikslar bilan to'qnashmaydi; kn/kxr — eski xabarlardagi tugmalar uchun)
    # kg/kgy/kgx — yo'nalish bo'yicha ommaviy berish, kv/kvy — ommaviy qabul qilish
    FIL_CB = ("ki", "kp", "kx", "kb", "kq", "kr", "krr", "kxr", "ka", "kk", "kl", "kd", "kh", "kn", "kg", "kgy", "kgx", "kv", "kvy",
              "kt", "kc")  # kt — 🏭 Taqsimot (sklad xodimi), kc — ⏰ Srok / 📦 30+ kun (A / B / C)
    EMPTY_TEXT = {"shoshilinch": "shoshilinch ko'chirish", "assortiment": "assortiment uchun ko'chirish",
                  "oddiy": "oddiy ko'chirish"}
    SKLAD_HINT = {"shoshilinch": "Dorixonada tugagan, sotuv bor — birinchi navbatda bering",
                  "assortiment": "Dorixonada yo'q, boshqa dorixonada sotilyapti — namuna sifatida bering",
                  "oddiy": "Dorixonada yetmayapti — sotuvicha to'ldiring"}

    @staticmethod
    def filial_button(text):
        """Pastki tugma → ('cat', kalit) | 'kin' | 'ksk' | 'kday' | 'home' | None. Belgi va «· N» dan, harf katta-kichikligidan qat'i nazar."""
        t = norm(re.sub(r"·.*$", "", text or ""))
        for c in KOCH_CATS:
            if c["title"] in t:
                return ("cat", c["key"])
        if "ZUDLIK" in t:  # eng eski klaviatura (🚨 Zudlik)
            return ("cat", "shoshilinch")
        if t.endswith("KELADIGAN"):
            return "kin"
        if t.endswith("TAQSIMOT"):
            return "ksk"
        if "BUGUNGI HOLAT" in t:
            return "kday"
        if t.startswith("⏰ SROK"):
            return "ksrok"
        if t.startswith("📦 30+"):
            return "kold"
        if t.endswith("MENYU"):
            return "home"
        return None

    @staticmethod
    def koch_key(r):
        """Ko'chirish qatorining kaliti: beruvchi + oluvchi + tovar."""
        return "|".join(norm(r[k]) for k in ("src", "dst", "name", "producer"))

    def koch_pending(self, filial, cat):
        """Exceldagi ro'yxat (toifa bo'yicha)."""
        return self.data.koch_out(filial, cat)

    def filial_users(self, filial):
        """Shu filialning faol xodimlari (xabar yuborish uchun)."""
        return [c for c in self.active_clients() if self.is_filial(c) and norm(c["filial"]) == norm(filial)]

    def _notify_filial(self, filial, text, markup=None, skip=None):
        n = 0
        for c in self.filial_users(filial):
            if c["chat_id"] == skip:
                continue
            try:
                self.tg.send(c["chat_id"], cut_html(text), markup)
                n += 1
            except Exception as e:  # noqa
                if not self.mark_gone(c["chat_id"], e):
                    logging.exception("Filial xodimiga xabar yuborib bo'lmadi: %s", c["chat_id"])
            self.sleep(0.05)
        return n

    # ---- sozlamalar (§11) ----
    KOCH_DEFAULTS = {"koch_on": "1", "koch_task_time": "08:30", "koch_deadline_shoshilinch": "10:00",
                     "koch_deadline_assortiment": "10:00", "koch_deadline_oddiy": "18:00", "koch_remind_before": "30",
                     "koch_srok_warn_days": "30", "koch_assort_min_days": "60", "koch_page_size": "10",
                     "koch_rem_start_on": "1", "koch_rem_before_on": "1", "koch_rem_overdue_on": "1",
                     "koch_no_return_days": "14", "koch_assort_limit": "30"}

    KOCH_NUMS = {"koch_remind_before": ("⏰ Eslatma (daqiqa oldin)", 0, 240), "koch_srok_warn_days": ("⚠️ Srok ogohlantirish (kun)", 1, 365),
                 "koch_assort_min_days": ("🟡 ASSORTIMENT minimal muddat (kun)", 0, 365),
                 "koch_assort_limit": ("🟡 ASSORTIMENT kunlik limit (har filialga, 0 — cheksiz)", 0, 1000),
                 "koch_page_size": ("📄 Sahifadagi dori", 5, 10),
                 "koch_no_return_days": ("↩️ Qaytarib yubormaslik (kun, 0 — o'chiq)", 0, 90)}

    def kset(self, key):
        return self.setting(key, self.KOCH_DEFAULTS[key])

    def kint(self, key):
        try:
            return int(self.kset(key))
        except (TypeError, ValueError):
            return int(self.KOCH_DEFAULTS[key])

    @staticmethod
    def parse_hm(v):
        m = re.match(r"^\s*(\d{1,2}):(\d{2})\s*$", str(v or ""))
        if not m or int(m.group(1)) > 23 or int(m.group(2)) > 59:
            return None
        return datetime.time(int(m.group(1)), int(m.group(2)))

    @property
    def koch_on(self):
        return str(self.kset("koch_on")).strip().lower() in ("1", "yes", "true", "ha")

    def koch_time(self, key):
        return self.parse_hm(self.kset(key))

    def koch_deadline(self, cat, sana=None):
        """Toifa muddati (datetime) yoki None (muddatsiz)."""
        t = self.koch_time("koch_deadline_" + (cat if cat in KOCH_CAT else "oddiy"))
        if t is None:
            return None
        d = datetime.date.fromisoformat(sana) if sana else now().date()
        return datetime.datetime.combine(d, t)

    @staticmethod
    def _today():
        return now().date().isoformat()

    @staticmethod
    def _yesterday():
        return (now().date() - datetime.timedelta(days=1)).isoformat()

    KOCH_RECV_DAYS = 14  # «berildi» (yo'ldagi) tovarni shuncha kun ichida «📥 Oldim» qilish mumkin

    def _recv_dates(self):
        """Yo'ldagi tovar qabul qilinadigan kunlar: bugun, kecha, ... (koch_close_old «berildi» ni yopmaydi)."""
        d = now().date()
        return [(d - datetime.timedelta(days=i)).isoformat() for i in range(self.KOCH_RECV_DAYS + 1)]

    @staticmethod
    def _dt(s):
        try:
            return datetime.datetime.fromisoformat(s) if s else None
        except ValueError:
            return None

    @staticmethod
    def _d(s):
        if isinstance(s, datetime.date):
            return s
        try:
            return datetime.date.fromisoformat(s) if s else None
        except ValueError:
            return None

    @staticmethod
    def fq(v):
        """Son: 0,5 ko'rinishida (vergul bilan)."""
        return qty(v or 0).replace(".", ",")

    # ---- B4.3: toifa bo'yicha filtr (muddat bo'yicha chiqarish) va saralash ----
    def koch_exclude_reason(self, row, today=None):
        """Qator ro'yxatga kirmasligi sababi yoki None. (E < 1, toifasiz, muddati o'tgan, ASSORTIMENT: < 60 kun,
        oxirgi koch_no_return_days kunda teskari yo'nalishda olingan — borib-kelish)."""
        if not row.get("cat") or (row.get("qty") or 0) < 1:
            return "Ko'chirish soni 1 dan kam" if row.get("cat") else "Toifa aniqlanmagan"
        st = Data.koch_date_state(self._d(row.get("date")), self.kint("koch_srok_warn_days"), today)
        if st == "expired":
            return "Muddati o'tgan"
        if row["cat"] == "assortiment" and st != "unknown":
            left = (self._d(row["date"]) - (today or now().date())).days
            if left < self.kint("koch_assort_min_days"):
                return "Muddati %d kundan kam (%d kun)" % (self.kint("koch_assort_min_days"), left)
        if self.koch_key(row) in self._koch_returned(today):
            return "%d kun ichida shu filialdan kelgan" % self.kint("koch_no_return_days")
        return None

    def _koch_returned(self, today=None):
        """Borib-kelishga qarshi: oxirgi N kunda B → A «olindi» bo'lgan tovarlar uchun A → B kalitlari to'plami.
        Kesh (60 soniya; yangi «olindi» / admin holati o'zgarganda tozalanadi)."""
        days, d = self.kint("koch_no_return_days"), today or now().date()
        if days <= 0:
            return frozenset()
        c = getattr(self, "_ret_cache", None)
        if c and c[0] == (d, days) and (now() - c[1]).total_seconds() < 60:
            return c[2]
        since = (d - datetime.timedelta(days=days)).isoformat()
        keys = frozenset("|".join(norm(x) for x in (t["dst"], t["src"], t["name"], t["producer"])) for t in self.db.execute(
            "SELECT src, dst, name, producer FROM koch_tasks WHERE status='olindi' AND received_at >= ?", (since,)))
        self._ret_cache = ((d, days), now(), keys)
        return keys

    def koch_sort(self, rows, cat):
        far = datetime.date.max
        dkey = lambda r: (self._d(r["date"]) if Data.koch_date_state(self._d(r["date"])) != "unknown" else None) or far
        if cat == "assortiment":
            key = lambda r: (-(r.get("src_sales") or 0), dkey(r), norm(r["name"]))
        elif cat == "oddiy":
            key = lambda r: (-(r.get("days_idle") or 0), -(r.get("summa") or 0), norm(r["name"]))
        else:
            key = lambda r: (dkey(r), -(r.get("summa") or 0), norm(r["name"]))
        return sorted(rows, key=key)

    # ---- ko'chirish A / B / C (pul = reja soni × kirim narxi, kunning hamma topshiriqlari, butun tarmoq bo'yicha) ----
    ABC_RANK = {"A": 0, "B": 1, "C": 2}

    @staticmethod
    def _kval(q, prih, summa=None):
        """Ko'chirish ABC puli: reja soni × kirim narxi (kirim narxi yo'q bo'lsa — Exceldagi summa)."""
        q, prih = q or 0, prih or 0
        return q * prih if prih else (summa or 0)

    def koch_abc(self, sana=None):
        """{task_id: 'A'|'B'|'C'} — shu kunning «bekor» dan boshqa hamma topshiriqlari (hamma filiallar) bo'yicha.
        Joyida hisoblanadi, kesh — (sana, topshiriqlar imzosi): reja / narx / ro'yxat o'zgarsa qayta hisoblanadi.
        Holat (berildi, olindi, rad) sinfni o'zgartirmaydi."""
        sana = sana or self._today()
        rows = self.db.execute("SELECT id, qty_plan, prih, summa FROM koch_tasks WHERE sana=? AND status!='bekor' ORDER BY id",
                               (sana,)).fetchall()
        sig = hashlib.md5(";".join("%d:%r:%r:%r" % (r["id"], r["qty_plan"], r["prih"], r["summa"]) for r in rows)
                          .encode("utf-8")).hexdigest()
        hit = self._abc_cache.get(sana)
        if hit and hit[0] == sig:
            return hit[1]
        res = ABC_T.classify(rows, lambda r: self._kval(r["qty_plan"], r["prih"], r["summa"]), lambda r: r["id"])
        if len(self._abc_cache) > 8:
            self._abc_cache.clear()
        self._abc_cache[sana] = (sig, res)
        return res

    def koch_abc_excel(self):
        """Nazorat o'chiq (topshiriq yo'q, ro'yxat Exceldan): {koch_key: sinf} — ro'yxatga kiradigan hamma qatorlar bo'yicha."""
        d = self.data
        key = (d.path, d.mtime, len(d.koch), self._today())
        hit = self._abc_cache.get("excel")
        if hit and hit[0] == key:
            return hit[1]
        rows = [r for r in d.koch if norm(r["src"]) and norm(r["dst"]) and not self.koch_exclude_reason(r)]
        res = ABC_T.classify(rows, lambda r: self._kval(r["qty"], r["prih"], r["summa"]), self.koch_key)
        self._abc_cache["excel"] = (key, res)
        return res

    def koch_cls(self, rows=(), sana=None):
        """Ro'yxat qatori → 'A' | 'B' | 'C' (topshiriq — id bo'yicha, Excel qatori — kalit bo'yicha)."""
        tm = self.koch_abc(sana) if any(r.get("id") for r in rows) else {}
        xm = []

        def cls(r):
            if r.get("id"):
                return tm.get(r["id"], "C")
            if not xm:
                xm.append(self.koch_abc_excel())
            return xm[0].get(self.koch_key(r), "C")
        return cls

    def koch_abc_summary(self, sana=None, filial=None):
        """Ko'chirish A / B / C xulosasi — bot («📋 Bugungi holat», bosh menyu) va server paneli (Api.koch_abc) uchun.
        Sinf — shu kunning hamma topshiriqlari («bekor» dan tashqari) bo'yicha BUTUN TARMOQDA (koch_abc, pul = reja soni ×
        kirim narxi). filial berilsa — sonlar, matritsa va ro'yxatlar faqat shu filial BERUVCHI (src) bo'lgan topshiriqlar
        bo'yicha (sinf baribir butun tarmoqdan). Bajarildi — «berildi» | «olindi»; bajarilmadi — «yangi» | «rad» | «yopildi».

        Qaytadi:
          {"sana": "YYYY-MM-DD", "filial": "ДОРИХОНА ЧИНОЗ" | None, "has": bool (shu kun topshiriq bormi),
           "limits":  {"A": 80, "B": 15, "C": 5},
           "total":   {"jami", "bajarildi", "bajarilmadi", "sum_jami", "sum_bajarildi", "sum_bajarilmadi"},
           "classes": {"A" | "B" | "C": {...shu 6 kalit..., "share": jami pulning % i (butun son)}},
           "filials": {beruvchi filial: {"A": [bajarildi, jami], "B": [...], "C": [...]}},
           "matrix":  {beruvchi filial: {"A": {6 kalit}, "B": {...}, "C": {...}, "total": {6 kalit}}},
           "done": [item, ...], "not_done": [item, ...]}      # A, B, C tartibida, sinf ichida pul ↓
          item = {"id", "sana", "cls", "cat", "name", "producer", "src", "dst", "by" ('tizim' | xodim ismi | ""),
                  "auto_note" (avto-aniqlash isboti: «ЧИНОЗ −12 · УРГУТ +12»),
                  "qty"  — bajarilganda berilgan son (qty_given), aks holda reja (qty_plan),
                  "qty_plan", "qty_given", "qty_received",
                  "sum"  — qty × kirim narxi (narx yo'q bo'lsa Exceldagi summa, son ulushida),
                  "status", "time" (berildi / olindi / rad vaqti HH:MM, aks holda ""), "given_at", "received_at" (HH:MM),
                  "reason" (faqat rad), "late_min" (faqat kechikkan «yangi»)}
          Summalar (sum_*) — itemlarning sum i yig'indisi: bajarilgan qismi haqiqatda berilgan son bo'yicha."""
        sana = sana or self._today()
        f = norm(filial) if filial else None
        cls = self.koch_abc(sana)
        n = now()
        tasks = [t for t in self.koch_tasks(sana) if t["status"] != "bekor"]
        keys = ("jami", "bajarildi", "bajarilmadi", "sum_jami", "sum_bajarildi", "sum_bajarilmadi")
        blank = lambda: dict.fromkeys(keys, 0)

        def add(b, ok, v):
            b["jami"] += 1
            b["sum_jami"] += v
            b["bajarildi" if ok else "bajarilmadi"] += 1
            b["sum_bajarildi" if ok else "sum_bajarilmadi"] += v

        total, classes, matrix, cells, done, not_done = blank(), {k: blank() for k in ABC_T.CLASSES}, {}, {}, [], []
        for t in tasks:
            if f and t["src"] != f:
                continue
            k = cls.get(t["id"], "C")
            ok = t["status"] in ("berildi", "olindi")
            plan = t["qty_plan"] or 0
            q = t["qty_given"] if ok and t["qty_given"] is not None else plan
            v = q * t["prih"] if t["prih"] else (t["summa"] or 0) * (q / plan if plan else 1)
            m = matrix.setdefault(t["src"], dict({c_: blank() for c_ in ABC_T.CLASSES}, total=blank()))
            add(m[k], ok, v)
            add(m["total"], ok, v)
            add(classes[k], ok, v)
            add(total, ok, v)
            cell = cells.setdefault(t["src"], {c_: [0, 0] for c_ in ABC_T.CLASSES})[k]
            cell[0] += 1 if ok else 0
            cell[1] += 1
            tm = (t["received_at"] if t["status"] == "olindi" else t["given_at"]) or ""
            it = {"id": t["id"], "sana": sana, "cls": k, "cat": t["cat"], "name": t["name"], "producer": t["producer"] or "",
                  "src": t["src"], "dst": t["dst"], "qty": q, "qty_plan": plan, "qty_given": t["qty_given"],
                  "qty_received": t["qty_received"], "sum": round(v, 2), "status": t["status"],
                  "time": tm[11:16] if t["status"] in ("berildi", "olindi", "rad") else "",
                  "given_at": (t["given_at"] or "")[11:16], "received_at": (t["received_at"] or "")[11:16],
                  "reason": (t["reason"] or "") if t["status"] == "rad" else "",
                  "by": self.koch_by(t), "auto_note": t["auto_note"] or ""}
            if t["status"] == "yangi":
                i = self.task_info(t, n)
                if i["late"]:
                    it["late_min"] = int(i["late_min"] or 0)
            (done if ok else not_done).append(it)
        for b in classes.values():
            b["share"] = int(round(100.0 * b["sum_jami"] / total["sum_jami"])) if total["sum_jami"] else 0
        order = lambda i: (self.ABC_RANK[i["cls"]], -i["sum"], norm(i["name"]), i["id"])
        return {"sana": sana, "filial": f, "has": bool(tasks) or self.koch_has_tasks(sana), "limits": dict(ABC_T.SHARE),
                "total": total, "classes": classes, "filials": cells, "matrix": matrix,
                "done": sorted(done, key=order), "not_done": sorted(not_done, key=order)}

    def _task_row(self, t):
        """koch_tasks qatori → ro'yxat qatori (Excel qatori bilan bir xil maydonlar)."""
        return {"id": t["id"], "src": t["src"], "dst": t["dst"], "name": t["name"], "producer": t["producer"], "cat": t["cat"],
                "qty": t["qty_plan"], "summa": t["summa"] or 0, "prih": t["prih"] or 0, "date": self._d(t["srok"]),
                "sabab": t["sabab"] or "", "days_idle": t["days_idle"] or 0, "dst_sales": t["dst_sales"] or 0,
                "dst_stock": t["dst_stock"] or 0, "src_sales": t["src_sales"] or 0, "src_stock": t["src_stock"] or 0,
                "status": t["status"], "qty_given": t["qty_given"], "qty_received": t["qty_received"]}

    def koch_view(self, f, cat):
        """(ro'yxat qatorlari, chiqarilganlar [(qator, sabab)]). Topshiriq bo'lsa — uning «yangi» lari, aks holda Excel."""
        f = norm(f)
        excel = self.data.koch_out(f, cat)
        excluded = [(r, why) for r in excel for why in [self.koch_exclude_reason(r)]
                    if why and (why.startswith("Muddati") or why.endswith("filialdan kelgan"))]
        if self.koch_on and self.koch_has_tasks():
            rows = [self._task_row(t) for t in self.koch_tasks(self._today(), src=f, cat=cat, statuses=("yangi",))]
            rows = [r for r in rows for why in [self.koch_exclude_reason(r) or ""]
                    if why != "Muddati o'tgan" and not why.endswith("filialdan kelgan")]
        elif self.koch_on:
            rows = []
        else:
            rows = [r for r in excel if not self.koch_exclude_reason(r)]
        lim = self.kint("koch_assort_limit")
        if cat == "assortiment" and self.koch_on and self.koch_has_tasks() and lim > 0 and not self.is_sklad(f):
            # limit to'lgan bo'lsa: qolganlar bugungi ro'yxatda yo'q, Excelda ko'rinadi va keyingi kunlarda navbat keladi
            active = {t["kkey"] for t in self.koch_tasks(self._today(), src=f, cat=cat) if t["status"] != "bekor"}
            if len(active) >= lim:
                excluded += [(r, self.WHY_QUEUE) for r in excel if not self.koch_exclude_reason(r) and self.koch_key(r) not in active]
        return self.koch_sort(rows, cat), excluded

    # ---- yo'nalish bo'yicha ommaviy tasdiqlash ----
    @staticmethod
    def koch_sig(pairs):
        """Ro'yxat imzosi (id, son): xodim ko'rgan ro'yxat tasdiq paytida ham o'zgarmaganini tekshiradi."""
        raw = ",".join("%d:%s" % (i, qty(q or 0)) for i, q in sorted(pairs))
        return hashlib.md5(raw.encode("ascii")).hexdigest()[:10]

    def koch_route_dsts(self, f, cat):
        """Bugungi oluvchilar (barqaror tartib: tugmadagi raqam shu ro'yxatdagi o'rin)."""
        return sorted({t["dst"] for t in self.koch_tasks(self._today(), src=norm(f), cat=cat)})

    def koch_routes(self, f, cat, rows=None):
        """[(raqam, oluvchi, qatorlar)] — filial hozir ko'rib turgan «yangi» qatorlar, oluvchi bo'yicha."""
        if not self.koch_on:
            return []
        rows = self.koch_view(f, cat)[0] if rows is None else rows
        if not rows or not all(r.get("id") for r in rows):
            return []
        out = []
        for i, d in enumerate(self.koch_route_dsts(f, cat)):
            part = [r for r in rows if norm(r["dst"]) == d]
            if part:
                out.append((i, d, part))
        return out

    def koch_route(self, f, cat, i, sig):
        """(oluvchi, qatorlar) yoki None: yo'nalish yo'q yoki ro'yxat xodim ko'rgandan keyin o'zgargan."""
        for n, d, rows in self.koch_routes(f, cat):
            if n == i and self.koch_sig((r["id"], r["qty"]) for r in rows) == sig:
                return d, rows
        return None

    def koch_sum(self, f, rows, bold=True):
        """« · N so'm». Summa kirim narxida, shuning uchun sklad xodimiga ko'rsatilmaydi."""
        if self.is_sklad(f):
            return ""
        return (" · <b>%s</b> so'm" if bold else " · %s so'm") % money(sum(r["summa"] for r in rows))

    def koch_route_card(self, f, cat, i, sig):
        """Ommaviy berishdan oldingi tasdiq: nima va qancha ketayotgani ko'rsatiladi."""
        route = self.koch_route(f, cat, i, sig)
        if not route:
            return None
        d, rows = route
        k = KOCH_CAT[cat]
        text = ("✅ <b>Ommaviy berish</b>\n%s %s\n🏬 %s → <b>%s</b>\n<b>%d</b> ta dori · <b>%s</b> qadoq%s\n\n"
                "Ro'yxatdagi hammasini berdingizmi?\nBerilmagan yoki boshqa sonda berilgan dori bo'lsa, avval "
                "«✏️ Farqni belgilash» ni bosing.") % (k["icon"], k["title"], esc(norm(f)), esc(d), len(rows),
                                                      qty(sum(r["qty"] for r in rows)), self.koch_sum(f, rows))
        return text, inline_kb([[("✅ Ha, hammasini berdim", "kgy:%s:%d:%s" % (cat, i, sig))],
                                [("✏️ Farqni belgilash", "kp:%s:0:%d" % (cat, i)), ("📎 Excel", "kgx:%s:%d" % (cat, i))],
                                self.nav_row("kp:%s:0" % cat, "kh")])

    def koch_give_route(self, cid, c, cat, i, sig):
        """Yo'nalishdagi hamma «yangi» — berildi. Hammasi yoki hech biri; ro'yxat o'zgargan bo'lsa None. Oluvchiga bitta xabar."""
        f = norm(c["filial"])
        ts = now().isoformat(timespec="seconds")
        with self.dblock:
            route = self.koch_route(f, cat, i, sig)
            if not route:
                return None
            d, rows = route
            idx = self.sale_index()
            try:
                for r in rows:
                    cur = self.db.execute("UPDATE koch_tasks SET status='berildi', qty_given=qty_plan, given_by=?, given_at=?, "
                                          "dst_stock1=? WHERE id=? AND status='yangi' AND src=?",
                                          (cid, ts, self.koch_dst_now(r["id"], idx), r["id"], f))
                    if cur.rowcount != 1:
                        self.db.rollback()
                        return None
                    self._event(r["id"], "berildi", cid, self._actor(cid), r["qty"], "ommaviy", ts)
                self.db.commit()
            except Exception:
                self.db.rollback()
                raise
        k = KOCH_CAT[cat]
        self._notify_filial(d, "📦 <b>%s dan keladi</b>: <b>%d</b> ta dori · <b>%s</b> qadoq\n%s %s · kelganda «📥 Keladigan» dan tasdiqlang." % (
            esc(f), len(rows), qty(sum(r["qty"] for r in rows)), k["icon"], k["title"]), inline_kb([[("📥 Keladigan", "kl:0")]]))
        return d, rows

    def koch_in_routes(self, f):
        """[(beruvchi, «berildi» topshiriqlar)] — oxirgi KOCH_RECV_DAYS kun, oluvchi filial uchun."""
        groups = {}
        for d in self._recv_dates():
            for t in self.koch_tasks(d, dst=norm(f), statuses=("berildi",)):
                groups.setdefault(t["src"], []).append(t)
        return sorted(groups.items())

    def koch_in_route(self, f, i, sig):
        """(beruvchi, topshiriqlar) yoki None: yo'nalish yo'q yoki ro'yxat xodim ko'rgandan keyin o'zgargan."""
        routes = self.koch_in_routes(f)
        if not 0 <= i < len(routes):
            return None
        src, tasks = routes[i]
        return (src, tasks) if self.koch_sig((t["id"], t["qty_given"]) for t in tasks) == sig else None

    def koch_in_card(self, f, i, sig):
        """Ommaviy qabuldan oldingi tasdiq."""
        route = self.koch_in_route(f, i, sig)
        if not route:
            return None
        src, tasks = route
        text = ("📥 <b>Ommaviy qabul</b>\n🏬 %s ← <b>%s</b>\n<b>%d</b> ta dori · <b>%s</b> qadoq · <b>%s</b> so'm\n\n"
                "Hammasi keldimi?\nKam kelgan dori bo'lsa, avval «⚠ Farqni belgilash» ni bosing.") % (
            esc(norm(f)), esc(src), len(tasks), qty(sum(t["qty_given"] for t in tasks)),
            money(sum(t["qty_given"] * (t["prih"] or 0) for t in tasks)))
        return text, inline_kb([[("✅ Ha, hammasini oldim", "kvy:%d:%s" % (i, sig))],
                                [("⚠ Farqni belgilash", "kl:0:%d" % i)], self.nav_row("kl:0", "kh")])

    def koch_receive_route(self, cid, c, i, sig):
        """Beruvchidan kelgan hamma «berildi» — olindi. Hammasi yoki hech biri; ro'yxat o'zgargan bo'lsa None."""
        f = norm(c["filial"])
        ts = now().isoformat(timespec="seconds")
        with self.dblock:
            route = self.koch_in_route(f, i, sig)
            if not route:
                return None
            src, tasks = route
            try:
                for t in tasks:
                    cur = self.db.execute("UPDATE koch_tasks SET status='olindi', qty_received=qty_given, received_by=?, received_at=? "
                                          "WHERE id=? AND status='berildi' AND dst=?", (cid, ts, t["id"], f))
                    if cur.rowcount != 1:
                        self.db.rollback()
                        return None
                    self._event(t["id"], "olindi", cid, self._actor(cid), t["qty_given"], "ommaviy", ts)
                self.db.commit()
            except Exception:
                self.db.rollback()
                raise
        self._ret_cache = None  # borib-kelish filtri yangilanadi
        self._notify_filial(src, "✅ <b>%s qabul qildi</b>: <b>%d</b> ta dori · <b>%s</b> qadoq" % (
            esc(f), len(tasks), qty(sum(t["qty_given"] for t in tasks))))
        return src, tasks

    def koch_close_old(self):
        """Yangi kun topshirig'i yaratilgach: o'tgan kunlarning bajarilmagan «yangi» lari yopiladi. «berildi» lar ochiq qoladi
        (tovar yo'lda). Yopilgan topshiriq beruvchining bajarilmagan ishi sifatida statistikada qoladi."""
        with self.dblock:
            cur = self.db.execute("UPDATE koch_tasks SET status='yopildi', note='kun oxiri' WHERE sana<? AND status='yangi'", (self._today(),))
            self.db.commit()
        if cur.rowcount:
            logging.info("Ko'chirish: o'tgan kunlarning bajarilmagan %d topshirig'i yopildi", cur.rowcount)
        return cur.rowcount

    def _srok_txt(self, r, cat):
        d = self._d(r.get("date"))
        st = Data.koch_date_state(d, self.kint("koch_srok_warn_days"))
        if st == "unknown":
            return "srok ?"
        s = "srok " + d.strftime("%d.%m.%Y")
        if st == "warn" and cat != "assortiment":
            s += " ⚠️"
        return s

    def _row_extra(self, r, cat):
        if self.is_sklad(r.get("src")):
            return "(u yerda sotuv %s · skladda %s)" % (self.fq(r.get("dst_sales")), self.fq(r.get("src_stock")))
        if cat == "assortiment":
            return "(sizda sotuv %s)" % self.fq(r.get("src_sales"))
        if cat == "oddiy":
            return "· %d kun turibdi · (u yerda sotuv %s)" % (r.get("days_idle") or 0, self.fq(r.get("dst_sales")))
        return "(u yerda sotuv %s)" % self.fq(r.get("dst_sales"))

    def deadline_line(self, cat, short=False):
        """⏰ SHOSHILINCH muddati 10:00 · 45 daqiqa qoldi  |  short: ⏰ Muddat 10:00 gacha · 45 daqiqa qoldi."""
        c = KOCH_CAT[cat]
        dl = self.koch_deadline(cat)
        if not dl:
            return "⏰ %s: muddatsiz" % c["title"]
        left = (dl - now()).total_seconds() / 60
        if left >= 0:
            return ("⏰ Muddat %s gacha · %s qoldi" if short else "⏰ " + c["title"] + " muddati %s · %s qoldi") % (
                dl.strftime("%H:%M"), self.fmt_min(left))
        return ("⛔ Muddat %s · %s kechikdi" if short else "⛔ " + c["title"] + " muddati %s · %s kechikdi") % (
            dl.strftime("%H:%M"), self.fmt_min(left))

    def koch_list_text(self, c, cat, page=0, di=None):
        """B4.4: toifa ro'yxati (HTML) + inline tugmalar. Qayerga bo'yicha guruhlangan, umumiy raqam, sahifalab.
        di: faqat shu yo'nalish (koch_route_dsts dagi raqam) — farqni belgilash uchun."""
        f = norm(c["filial"])
        k = KOCH_CAT[cat]
        h = "%s <b>%s</b> · 🏬 %s" % (k["icon"], k["title"], esc(f))
        if self.koch_on and not self.koch_has_tasks():
            return h + "\n\nBugungi topshiriq hali shakllanmagan. %s dan keyingi «Yangilash» tugagach paydo bo'ladi." % (
                self.kset("koch_task_time")), None
        rows, excluded = self.koch_view(f, cat)
        routes = self.koch_routes(f, cat, rows)
        route = next((x for x in routes if x[0] == di), None) if di is not None else None
        tail = ":%d" % route[0] if route else ""
        if route:
            rows = route[2]
            excluded = [(r, why) for r, why in excluded if norm(r["dst"]) == route[1]]
        if not rows:
            text = h + "\n\n✅ Hozircha %s yo'q." % self.EMPTY_TEXT[cat]
        else:
            per = max(1, self.kint("koch_page_size"))
            order, groups = [], {}
            for r in rows:
                g = norm(r["dst"])
                if g not in groups:
                    order.append(g)
                groups.setdefault(g, []).append(r)
            cls = self.koch_cls(rows)
            for g in order:  # oluvchi guruhi ichida: A birinchi, keyin pul bo'yicha (guruhlar tartibi — toifa saralashidan)
                groups[g].sort(key=lambda r: (self.ABC_RANK[cls(r)], -self._kval(r["qty"], r["prih"], r["summa"])))
            flat = [(g, r) for g in order for r in groups[g]]
            pages = (len(flat) - 1) // per + 1
            page = min(max(page, 0), pages - 1)
            text = h + "\n%s\nJami: <b>%d</b> ta dori · <b>%s</b> qadoq%s" % (
                self.SKLAD_HINT[cat] if self.is_sklad(f) else k["hint"], len(rows), qty(sum(r["qty"] for r in rows)),
                self.koch_sum(f, rows))
            if self.koch_on:
                text += "\n" + self.deadline_line(cat, short=True)
            if routes and not route:
                text += "\n\n<b>Yo'nalishlar:</b>\n" + "\n".join("→ %s · %d ta · %s qadoq%s" % (
                    esc(d), len(rr), qty(sum(r["qty"] for r in rr)), self.koch_sum(f, rr, bold=False)) for _, d, rr in routes)
            cur = None
            for i, (g, r) in enumerate(flat[page * per:(page + 1) * per], page * per + 1):
                if g != cur:
                    text += "\n\n<b>→ %s</b> · %d ta" % (esc(g), len(groups[g]))
                    cur = g
                text += "\n%d. <code>%s</code> %s · <b>%s</b> qadoq · %s %s" % (
                    i, cls(r), esc(self.short(r["name"], 38)), qty(r["qty"]), self._srok_txt(r, cat), self._row_extra(r, cat))
            if pages > 1:
                text += "\n\n<i>%d-sahifa / %d</i>" % (page + 1, pages)
        back = sum(1 for r, why in excluded if why.endswith("filialdan kelgan"))
        if back:
            text += "\n\nℹ️ %d kun ichida shu filialdan kelgan %d ta dori qaytarib yuborilmaydi (Excelda)." % (
                self.kint("koch_no_return_days"), back)
        if cat == "shoshilinch":
            exp = [r for r, why in excluded if why == "Muddati o'tgan"]
            if exp:
                text += "\n\n<b>⛔ Muddati o'tgan (ko'chirilmaydi)</b> · %d ta\n" % len(exp) + "\n".join(
                    "• %s → %s · %s" % (esc(self.short(r["name"], 34)), esc(self.short(norm(r["dst"]), 18)),
                                        self._d(r["date"]).strftime("%d.%m.%Y")) for r in exp[:5]) + (
                    "\n… yana %d ta (Excelda)" % (len(exp) - 5) if len(exp) > 5 else "")
        elif excluded:
            short = [why for r, why in excluded if why.startswith("Muddati") and why != "Muddati o'tgan"]
            if cat == "assortiment" and short:
                text += "\n\nℹ️ Muddati %d kundan kam bo'lgan %d ta dori kiritilmadi (Excelda)." % (
                    self.kint("koch_assort_min_days"), len(short))
            nexp = sum(1 for r, why in excluded if why == "Muddati o'tgan")
            if nexp:
                text += "\nℹ️ Muddati o'tgan %d ta dori kiritilmadi (Excelda)." % nexp
        nq = sum(1 for r, why in excluded if why == self.WHY_QUEUE)
        if nq:
            text += "\n\n📋 Yana %d ta dori navbatda: kunlik limit %d ta, keyingi kunlarda ro'yxatga kiradi (Excelda)." % (
                nq, self.kint("koch_assort_limit"))
        text += self.foot()
        kb = []
        for n, d, rr in ([route] if route else routes):  # ommaviy berish: yo'nalish bo'yicha bitta tugma
            sig = self.koch_sig((r["id"], r["qty"]) for r in rr)
            kb.append([(("✅ Hammasini berdim · %d ta" % len(rr)) if route else "✅ → %s · %d ta" % (bl(self.short(d, 26)), len(rr)),
                        "kg:%s:%d:%s" % (cat, n, sig))])
        if rows:
            ids = [r for g, r in flat[page * per:(page + 1) * per]]
            if all(r.get("id") for r in ids):
                btns = [(str(i), "ki:%d:%d%s" % (r["id"], page, tail)) for i, r in enumerate(ids, page * per + 1)]
                kb += [btns[i:i + 5] for i in range(0, len(btns), 5)]
            nav = []
            if page > 0:
                nav.append(("◀", "kp:%s:%d%s" % (cat, page - 1, tail)))
            if page < pages - 1:
                nav.append(("▶", "kp:%s:%d%s" % (cat, page + 1, tail)))
            kb.append(nav + [("📎 Excel", ("kgx:%s:%d" % (cat, route[0])) if route else "kx:%s" % cat)])
        elif excluded:
            kb.append([("📎 Excel", "kx:%s" % cat)])
        return cut_html(text), (inline_kb(kb) if kb else None)

    def koch_item_card(self, c, tid, page=0, di=None):
        """B5.3: dori kartochkasi va rolga mos tugmalar. Faqat shu topshiriqning beruvchi yoki oluvchi filiali."""
        f = norm(c["filial"])
        t = self.db.execute("SELECT * FROM koch_tasks WHERE id=?", (tid,)).fetchone()
        if not t or f not in (t["src"], t["dst"]):
            return "Topshiriq topilmadi yoki sizga tegishli emas.", None
        k = KOCH_CAT.get(t["cat"], KOCH_CAT["oddiy"])
        r = self._task_row(t)
        abc = self.koch_abc(t["sana"]).get(t["id"]) if t["status"] != "bekor" else None
        text = "%s <b>%s</b> · #%d%s\n💊 <b>%s</b>\n%s\n🏬 %s → <b>%s</b>\nSoni: <b>%s</b> qadoq%s\n%s · %s\nSabab: %s" % (
            k["icon"], k["title"], t["id"], (" · <code>%s</code>" % abc) if abc else "", esc(t["name"]), esc(t["producer"] or ""), esc(t["src"]), esc(t["dst"]),
            qty(t["qty_plan"]), "" if self.is_sklad(f) else " · Summa: %s so'm" % money(t["summa"] or 0), self._srok_txt(r, t["cat"]),
            self._row_extra(r, t["cat"]), esc_t(t["sabab"] or "—"))
        text += "\nHolat: <b>%s</b>" % self.ST_NAME.get(t["status"], t["status"])
        tm = (t["received_at"] if t["status"] == "olindi" else t["given_at"] if t["status"] in ("berildi", "rad") else "") or ""
        if tm[11:16]:
            text += " %s" % tm[11:16]
        text += self.koch_by_text(t)
        if t["qty_given"] is not None:
            text += " · berildi %s" % qty(t["qty_given"])
        if t["qty_received"] is not None:
            text += " · olindi %s" % qty(t["qty_received"])
        if t["reason"]:
            text += "\nSabab (rad): %s" % esc(t["reason"])
        if t["auto_note"]:
            text += "\n🤖 Qoldiq: %s" % esc(t["auto_note"])
        i = self.task_info(t)
        if i["late"]:
            text += "\n⛔ Kechikkan: %s · aybdor: %s" % (self.fmt_min(i["late_min"]), esc(i["guilty_filial"]))
        elif t["status"] in ("yangi", "berildi") and t["sana"] == self._today():
            text += "\n" + self.deadline_line(t["cat"], short=True)
        ev = self.db.execute("SELECT * FROM koch_events WHERE task_id=? ORDER BY id DESC LIMIT 5", (tid,)).fetchall()
        if ev:
            text += "\n\n<b>Tarix:</b>\n" + "\n".join("%s · %s · %s%s" % (
                (e["ts"] or "")[11:16], esc_t(e["action"]), esc(e["user_name"] or ""),
                (" · %s qadoq" % qty(e["qty"])) if e["qty"] is not None else "") for e in reversed(ev))
        kb = []
        if t["src"] == f and t["status"] == "yangi" and t["sana"] in (self._today(), self._yesterday()):
            kb.append([("✅ Berdim", "kb:%d" % tid), ("✏️ Boshqa son", "kq:%d" % tid)])
            kb.append([("❌ Bera olmayman", "kr:%d" % tid)])
        if t["dst"] == f and t["status"] == "berildi" and t["sana"] in self._recv_dates():
            kb.append([("📥 Oldim", "ka:%d" % tid), ("⚠ Kam keldi", "kk:%d" % tid)])
        kb.append(self.nav_row("kp:%s:%d%s" % (t["cat"], page, "" if di is None else ":%d" % di) if t["src"] == f else "kl:0", "kh"))
        return text, inline_kb(kb)

    # ---- 1) kunlik topshiriq (snapshot) ----
    def koch_has_tasks(self, sana=None):
        return bool(self.db.execute("SELECT 1 FROM koch_tasks WHERE sana=? LIMIT 1", (sana or self._today(),)).fetchone())

    def koch_snapshot_due(self, in_refresh=False):
        """koch_task_time dan keyingi birinchi muvaffaqiyatli yangilash (bot yoki Excelda qo'lda — Excel ichidagi vaqt).
        in_refresh=True — yangilash oqimining o'zi (Excel tugagan, refreshing hali tushirilmagan) chaqirmoqda."""
        if not self.koch_on or (self.refreshing and not in_refresh):
            return False
        t, n = self.koch_time("koch_task_time"), now()
        if t is None or n.time() < t or self.setting("koch_snap_date") == n.date().isoformat():
            return False
        self.data.reload()
        w = self.data.when()
        # eski makros varag'idan (Turi ustunsiz) topshiriq yaratilmaydi: yangi Kochirish.bas bilan «Yangilash» kerak
        return bool(w) and not self.data.error and self.data.koch_new and w >= datetime.datetime.combine(n.date(), t)

    def koch_sync_due(self, in_refresh=False):
        """Bugungi topshiriq bor va Excel shundan keyin yana yangilangan (bot yoki Excelda qo'lda «Yangilash»):
        Excel ichidagi vaqt oxirgi shakllantirishdagidan yangi."""
        d = self._today()
        if not self.koch_on or (self.refreshing and not in_refresh) or self.setting("koch_snap_date") != d or not self.koch_has_tasks(d):
            return False
        self.data.reload()
        w = self.data.when()
        if not w or self.data.error or not self.data.koch_new or w.date().isoformat() != d:
            return False
        last = self.setting("koch_sync_stamp")
        return not last or w.isoformat(timespec="seconds") > last

    @staticmethod
    def _same(a, b):
        if isinstance(a, (int, float)) and isinstance(b, (int, float)):
            return abs(a - b) < 1e-6
        return (a if a not in ("", None) else None) == (b if b not in ("", None) else None)

    def koch_limit(self, rows, old):
        """ASSORTIMENT: har beruvchi filialga kunlik limit. Bajarilganlar (berildi / olindi / rad) limitga kiradi. Bo'sh joylar
        avval hozirgi «yangi» lar bilan to'ldiriladi (ro'yxat kun davomida bekorga almashmasin), keyin beruvchida eng ko'p
        sotilayotganlar bilan. rows: {kkey: qator}, old: bugungi {kkey: topshiriq}."""
        lim = self.kint("koch_assort_limit")
        if lim <= 0:
            return rows
        used, groups = {}, {}
        for t in old.values():
            if t["cat"] == "assortiment" and t["status"] in ("berildi", "olindi", "rad"):
                used[t["src"]] = used.get(t["src"], 0) + 1
        for k, r in rows.items():
            t = old.get(k)
            if r["cat"] != "assortiment" or self.is_sklad(r["src"]) or (t is not None and t["status"] in ("berildi", "olindi", "rad")):
                continue  # sklad limitsiz: limit filiallar o'rtasidagi ko'chirish uchun
            if t is not None and t["status"] == "bekor" and t["note"] != "regen":
                continue  # admin bekor qilgan: limitda joy egallamaydi
            g = groups.setdefault(norm(r["src"]), ([], []))
            g[0 if t is not None and t["status"] == "yangi" else 1].append(r)
        out = dict(rows)
        for src, (current, fresh) in groups.items():
            order = self.koch_sort(current, "assortiment") + self.koch_sort(fresh, "assortiment")
            for r in order[max(0, lim - used.get(src, 0)):]:
                del out[self.koch_key(r)]
        return out

    def koch_build(self, sana=None, regen=False, by=None):
        """Bugungi topshiriqni "Ko'chirish" varag'idan yaratadi (B4.3 chiqarishlari qo'llangan: E < 1, toifasiz, muddati o'tgan,
        ASSORTIMENT < 60 kun). regen=True: o'zgargan «yangi» lar yangilanadi, ro'yxatdan chiqqanlari «bekor» (note='regen'),
        qaytganlari yana «yangi»; «berildi» / «olindi» / «rad» lar saqlanadi. Tarixga faqat haqiqiy o'zgarish yoziladi."""
        sana = sana or self._today()
        day0 = datetime.date.fromisoformat(sana)
        ts = now().isoformat(timespec="seconds")
        rows = {}
        for r in self.data.koch:
            if norm(r["src"]) and norm(r["dst"]) and not self.koch_exclude_reason(r, day0):
                rows.setdefault(self.koch_key(r), r)
        cnt = {"new": 0, "updated": 0, "cancelled": 0, "kept": 0}
        uid, uname = by if by else (None, "tizim")
        sale = self.sale_index()
        cols = ("src", "dst", "name", "producer", "firm", "turi", "cat", "sabab", "days_idle", "qty_plan", "prih", "summa", "srok",
                "dst_sales", "dst_stock", "src_sales", "src_stock")
        with self.dblock:
            old = {t["kkey"]: t for t in self.db.execute("SELECT * FROM koch_tasks WHERE sana=?", (sana,))}
            rows = self.koch_limit(rows, old)
            for k, r in rows.items():
                vals = (norm(r["src"]), norm(r["dst"]), r["name"], r["producer"], r.get("firm") or "", r.get("turi") or "",
                        r["cat"], r.get("sabab") or "", r.get("days_idle") or 0, r["qty"], r["prih"], r["summa"],
                        r["date"].isoformat() if r["date"] else None, r["dst_sales"], r["dst_stock"], r["src_sales"],
                        r["src_stock"])
                t = old.get(k)
                if t is None:
                    # boshlang'ich qoldiq surati (avto-aniqlash uchun): Sale varag'idan, topilmasa — Ko'chirish varag'idagi qiymat
                    s0, d0 = (self.stock_now(r["name"], r["producer"], f, sale) for f in (r["src"], r["dst"]))
                    self.db.execute("INSERT INTO koch_tasks(sana, %s, status, created_at, kkey, src_stock0, dst_stock0) "
                                    "VALUES(?,%s,'yangi',?,?,?,?)" % (", ".join(cols), ",".join("?" * len(cols))),
                                    (sana,) + vals + (ts, k, r["src_stock"] if s0 is None else s0, r["dst_stock"] if d0 is None else d0))
                    cnt["new"] += 1
                elif regen and t["status"] == "bekor" and t["note"] == "regen":
                    # ro'yxatga qaytdi: yangi topshiriq kabi (created_at = hozir — muddatdan keyin bo'lsa kechikkan emas)
                    self.db.execute("UPDATE koch_tasks SET %s, status='yangi', note=NULL, created_at=? WHERE id=?" % ", ".join(
                        "%s=?" % c_ for c_ in cols), vals + (ts, t["id"]))
                    self._event(t["id"], "regen:qaytdi", uid, uname, r["qty"], None, ts)
                    cnt["new"] += 1
                elif regen and t["status"] == "yangi" and not all(self._same(t[c_], v) for c_, v in zip(cols, vals)):
                    if (r["qty"] or 0) < (t["qty_plan"] or 0) and (t["auto_warn"] or self.koch_moved(t, sale)):
                        # qoldiq ko'chirish izini ko'rsatmoqda (yoki tizim belgilagan): reja kamaytirilmaydi — avto-aniqlash hal qiladi
                        vals = vals[:9] + (t["qty_plan"], vals[10], t["summa"]) + vals[12:]
                        if all(self._same(t[c_], v) for c_, v in zip(cols, vals)):
                            cnt["kept"] += 1
                            continue
                    self.db.execute("UPDATE koch_tasks SET %s WHERE id=? AND status='yangi'" % ", ".join(
                        "%s=?" % c_ for c_ in cols), vals + (t["id"],))
                    if not self._same(t["qty_plan"], vals[9]):
                        self._event(t["id"], "regen:yangilandi", uid, uname, vals[9], "reja %s edi" % qty(t["qty_plan"]), ts)
                    cnt["updated"] += 1
                else:
                    cnt["kept"] += 1
            if regen:
                for k, t in old.items():
                    if k not in rows and t["status"] == "yangi":
                        if self.koch_moved(t, sale):  # ro'yxatdan chiqishi ko'chirish bo'lgani uchun bo'lishi mumkin: avto-aniqlashga qoladi
                            cnt["kept"] += 1
                            continue
                        self.db.execute("UPDATE koch_tasks SET status='bekor', note='regen' WHERE id=? AND status='yangi'", (t["id"],))
                        self._event(t["id"], "regen:bekor", uid, uname, None, "Qayta shakllantirishda ro'yxatda yo'q", ts)
                        cnt["cancelled"] += 1
            self.db.commit()
        self.set_setting("koch_snap_date", sana)
        w = self.data.when()
        if w:
            self.set_setting("koch_sync_stamp", w.isoformat(timespec="seconds"))
        logging.info("Ko'chirish topshirig'i %s (%s): %s", sana, "qayta shakllantirildi" if regen else "yaratildi", cnt)
        return cnt

    def koch_try_snapshot(self, in_refresh=False, detected=False):
        """Vaqti kelgan bo'lsa bugungi topshiriqni yaratadi va filial xodimlariga xabar beradi. Asosiy sikl va yangilash
        oqimi bir vaqtda chaqirsa — bittasi bajaradi (xabar ikki marta ketmaydi). Yaratish / qayta shakllantirishdan OLDIN
        avto-aniqlash ishlaydi (detected=True — chaqiruvchi shu ma'lumot bilan allaqachon aniqlagan): aks holda makros
        ro'yxatidan chiqqan (ko'chirilgan) qatorlar «bekor» bo'lib, aniqlashga yetib bormaydi."""
        if not self._snap_lock.acquire(blocking=False):
            return None
        try:
            return self._koch_snapshot(in_refresh, detected)
        finally:
            self._snap_lock.release()

    def _koch_detect_first(self):
        try:
            self.koch_auto_detect()
        except Exception:  # noqa
            logging.exception("Ko'chirishni avtomatik aniqlab bo'lmadi")

    def _koch_snapshot(self, in_refresh=False, detected=False):
        if not self.koch_snapshot_due(in_refresh):
            if self.koch_sync_due(in_refresh):  # kun davomidagi har «Yangilash» dan keyin topshiriq Excelga moslashadi
                if not detected:
                    self._koch_detect_first()
                return self.koch_build(regen=True, by=(None, "tizim (Yangilash)"))
            return None
        if not detected:
            self._koch_detect_first()
        cnt = self.koch_build()
        self._koch_day_start()
        return cnt

    def koch_regen_manual(self, by):
        """Panel «Qayta shakllantirish». Bugun hali topshiriq yaratilmagan bo'lsa (ertalabki yangilash o'xshamagan),
        bu kunning birinchi shakllantirishi: o'tgan kunlar yopiladi va filiallarga ertalabki xabar ketadi."""
        if not self._snap_lock.acquire(timeout=120):
            raise RuntimeError("Topshiriq hozir shakllantirilmoqda, birozdan keyin qayta urinib ko'ring")
        try:
            first = self.setting("koch_snap_date") != self._today()
            self._koch_detect_first()  # avval aniqlash: ko'chirilgan qatorlar «bekor» bo'lib ketmasin
            cnt = self.koch_build(regen=True, by=by)
            if first:
                self._koch_day_start()
            return cnt
        finally:
            self._snap_lock.release()

    def _koch_day_start(self):
        """Kunning birinchi topshirig'idan keyin: o'tgan kunlarning «yangi» lari yopiladi, filiallarga xabar."""
        self.koch_close_old()
        if not self.maint_on and self.kset("koch_rem_start_on") == "1":
            d = self._today()
            for f in sorted({t["src"] for t in self.koch_tasks(d)} | {t["dst"] for t in self.koch_tasks(d)}):
                lines = "\n".join("%s %s — berish: <b>%d</b>" % (k["icon"], k["title"], len(self.koch_tasks(d, src=f, cat=k["key"])))
                                  for k in KOCH_CATS)
                self._notify_filial(f, "📋 <b>Bugungi ko'chirish topshirig'i tayyor</b>\n🏬 %s\n\n%s\n📥 Sizga keladi: <b>%d</b>\n%s" % (
                    esc(f), lines, len(self.koch_tasks(d, dst=f)), self.deadline_line("shoshilinch")),
                    inline_kb([[("📂 Ochish", "kh")]]))  # bildirishnoma: bosilganda panel shu xabar ostida ochiladi

    def koch_tasks(self, sana, src=None, dst=None, cat=None, statuses=None, turi=None):
        cat = cat or {"zudlik": "shoshilinch"}.get(turi, turi)
        q, a = "SELECT * FROM koch_tasks WHERE sana=?", [sana]
        for col, v in (("src", src), ("dst", dst)):
            if v:
                q += " AND %s=?" % col
                a.append(norm(v))
        if cat:
            q += " AND cat=?"
            a.append(cat)
        if statuses:
            q += " AND status IN (%s)" % ",".join("?" * len(statuses))
            a += list(statuses)
        order = "CASE cat " + " ".join("WHEN '%s' THEN %d" % (c["key"], i) for i, c in enumerate(KOCH_CATS)) + " ELSE 9 END"
        return self.db.execute(q + " ORDER BY %s, srok IS NULL, srok, summa DESC, id" % order, a).fetchall()

    def _event(self, tid, action, uid, uname, qty_, note, ts=None):
        self.db.execute("INSERT INTO koch_events(task_id, action, user_id, user_name, qty, note, ts) VALUES(?,?,?,?,?,?,?)",
                        (tid, action, uid, uname, qty_, note, ts or now().isoformat(timespec="seconds")))

    # ---- 2) muddat va kechikish (B5.2) ----
    def task_info(self, t, n=None):
        """kechikkanmi, kim aybdor, kech bajarilganmi, miqdor farqi."""
        n = n or now()
        dl = self.koch_deadline(t["cat"], t["sana"])
        created = self._dt(t["created_at"])
        if dl and created and created > dl:  # muddatdan keyin («Yangilash» bilan) qo'shilgan — shu kuni kechikkan hisoblanmaydi
            dl = None
        st = t["status"]
        info = {"deadline": dl, "late": False, "late_done": False, "guilty": None, "guilty_filial": None, "late_min": 0,
                "farq": (t["qty_given"] is not None and t["qty_given"] < t["qty_plan"]) or
                        (t["qty_received"] is not None and t["qty_given"] is not None and t["qty_received"] < t["qty_given"])}
        if dl and st in ("yangi", "berildi") and n > dl:
            info.update(late=True, guilty="src" if st == "yangi" else "dst", late_min=int((n - dl).total_seconds() // 60))
        elif dl and st == "olindi":
            rec, giv = self._dt(t["received_at"]), self._dt(t["given_at"])
            if rec and rec > dl:
                info.update(late_done=True, late_min=int((rec - dl).total_seconds() // 60),
                            guilty="src" if giv and giv > dl else "dst")
        elif dl and st == "yopildi":  # o'sha kuni bajarilmagan: kechikish kun oxirigacha hisoblanadi, vaqt o'tgani sari o'smaydi
            end = datetime.datetime.combine(dl.date(), datetime.time(23, 59))
            info.update(late_done=True, guilty="src", late_min=int((end - dl).total_seconds() // 60))
        if info["guilty"]:
            info["guilty_filial"] = t[info["guilty"]]
        return info

    @staticmethod
    def fmt_min(m):
        m = int(abs(m))
        return ("%d soat %d daqiqa" % (m // 60, m % 60)) if m >= 60 else "%d daqiqa" % m

    # ---- B4.1: klaviatura, xush kelibsiz, xabarlar ----
    def filial_counts(self, f):
        """{toifa: N, 'kin': N4} — tugmadagi son = ro'yxatdagi «Jami»."""
        out = {k["key"]: len(self.koch_view(f, k["key"])[0]) for k in KOCH_CATS}
        if self.koch_on and self.koch_has_tasks():  # = «📥 Keladigan»: bugungi kutilayotganlar + yo'ldagi (oldingi kunlardan ham)
            out["kin"] = (len(self.koch_tasks(self._today(), dst=f, statuses=("yangi",)))
                          + sum(len(tasks) for _, tasks in self.koch_in_routes(f)))
        elif self.koch_on:
            out["kin"] = 0
        else:
            out["kin"] = len(self.data.koch_in(f))
        return out

    def filial_root_rows(self, c):
        """Filial (sklad) xodimi bosh menyusi: [[(matn, callback, style)]]. Sonlar tugmalarda. Yangi bo'lim — shu yerga qator."""
        self.data.reload()
        n = self.filial_counts(norm(c["filial"]))
        btn = lambda k: ("%s %s · %d" % (KOCH_CAT[k]["icon"], KOCH_CAT[k]["title"], n[k]), "kp:%s:0" % k, KOCH_CAT[k]["style"])
        day = [("🏭 Taqsimot", "kt") if self.is_sklad(c["filial"]) else ("📥 Keladigan · %d" % n["kin"], "kl:0"), ("📋 Bugungi holat", "kd")]
        f = norm(c["filial"])
        abc = [("⏰ Srok · %d" % len(self.fil_abc_rows(f, "s")), "kc:s"), ("📦 30+ kun · %d" % len(self.fil_abc_rows(f, "o")), "kc:o")]
        return [[btn("shoshilinch")], [btn("assortiment"), btn("oddiy")], day, abc]

    def filial_kb(self, c):
        """Bosh menyu tugmalari pastki klaviatura ko'rinishida (server paneli «mijoz ko'rinishi»). Telegramga yuborilmaydi."""
        rows = [[{"text": b[0], "style": b[2] if len(b) > 2 else None} for b in row] for row in self.filial_root_rows(c)]
        return reply_kb(rows + [["🏠 Menyu"]], is_persistent=True)

    def filial_home(self, c):
        """Filial xodimi bosh menyusi (inline panel)."""
        return self.filial_home_text(c), inline_kb(self.filial_root_rows(c) + [[("🔄 Yangilash", "kh")]])

    def filial_panel(self, c, save_sig=True):
        """(matn, tugmalar pastki klaviatura ko'rinishida) — server paneli va admin ko'rinishi uchun; botda — filial_home."""
        return self.filial_home_text(c), self.filial_kb(c)

    def filial_home_text(self, c):
        d = self.data
        d.reload()
        f = norm(c["filial"])
        n = self.filial_counts(f)
        sk = self.is_sklad(f)
        text = ("🏭 <b>Sklad paneli</b>\n<b>%s</b>\n\n" if sk else "🏠 <b>Bosh menyu</b>\n🏬 <b>%s</b> · filial\n\n") % esc(f)
        if self.koch_on and not self.koch_has_tasks():
            text += "📋 Bugungi ko'chirish topshirig'i hali shakllanmagan (%s dan keyingi «Yangilash» tugagach paydo bo'ladi)." % (
                self.kset("koch_task_time"))
        else:
            text += "\n".join("%s %s — berish kerak: <b>%d</b> ta" % (k["icon"], k["title"], n[k["key"]]) for k in KOCH_CATS)
            if sk:
                left = {}
                for t in self.koch_tasks(self._today(), src=f, statuses=("yangi",)):
                    left[t["dst"]] = left.get(t["dst"], 0) + 1
                if left:
                    text += "\n\n<b>Dorixonalarga:</b>\n" + "\n".join("→ %s · %d ta" % (esc(d), left[d]) for d in sorted(left))
            else:
                text += "\n📥 Sizga keladi: <b>%d</b> ta" % n["kin"]
            if self.koch_on and self.koch_has_tasks():
                a = self.koch_abc_summary(filial=f)["classes"]["A"]
                if a["jami"]:  # beradigan A lar (eng qimmat) — bajarilmagani
                    text += "\n💰 A: <b>%d</b> ta qoldi%s" % (a["bajarilmadi"], "" if sk else " · %s" % mln(a["sum_bajarilmadi"]))
            if self.koch_on:
                text += "\n\n" + self.deadline_line("shoshilinch")
        abc = []
        for kind in ("s", "o"):
            rows, _, sm, _ = self.fil_abc(f, kind)
            if rows and sm["A"]["n"]:
                abc.append("%s · A: <b>%d</b> ta%s" % (self.ABC_KINDS[kind], sm["A"]["n"], "" if sk else " · %s" % mln(sm["A"]["sum"])))
        if abc:
            text += "\n\n" + "\n".join(abc)
        return text + "\n\n🕒 Ma'lumot: %s\nBo'limni tanlang 👇" % dt_text(d.when())

    def filial_message(self, cid, c, text):
        """Filial xodimining xabari. Filial nomi bazadan (c['filial']) olinadi. Toifa tugmasi — qidiruv va fallback'dan oldin.
        Eski pastki tugma (kirill/lotin) — shu bo'lim yangi panel bo'lib ochiladi."""
        st = self.state.get(cid)
        kind = self.filial_button(text)
        if st and st[0] in ("koch_qty", "koch_rad", "koch_kam"):
            self.state.pop(cid, None)
            if text != "/cancel" and not kind:
                return self.filial_state_input(cid, c, st, text)
        if isinstance(kind, tuple):
            self.touch(cid, kind[1])
            return self.panel_resend(cid, *self.fil_list_screen(c, kind[1]))
        if kind in ("kin", "kday"):
            self.touch(cid, kind)
            f = norm(c["filial"])
            return self.panel_resend(cid, *(self.fil_in_screen(f) if kind == "kin" else self.fil_day_screen(f)))
        if kind == "ksk" and self.is_sklad(c["filial"]):
            self.touch(cid, kind)
            return self.panel_resend(cid, *self.fil_sklad_screen())
        if kind in ("ksrok", "kold"):
            self.touch(cid, "abc:" + kind[1])
            return self.panel_resend(cid, *self.fil_abc_screen(c, kind[1]))
        self.touch(cid, "start")
        return self.panel_resend(cid, *self.filial_home(c))

    # ---- filial panelining ekranlari: yo'l qatori + navigatsiya (matn funksiyalari o'zgarmaydi) ----
    def fil_screen(self, crumbs, text, kb=None, back="kh"):
        return self.screen(crumbs, text, kb, back=back, home="kh")

    @staticmethod
    def cat_crumb(cat):
        return "%s %s" % (KOCH_CAT[cat]["icon"], KOCH_CAT[cat]["title"])

    def fil_list_screen(self, c, cat, page=0, di=None):
        """🔴/🟡/🟢 ro'yxat (di — bitta yo'nalish: «🏠 › 🔴 SHOSHILINCH › → ЭРКИН»)."""
        text, kb = self.koch_list_text(c, cat, page, di)
        crumbs = (self.cat_crumb(cat),)
        dsts = self.koch_route_dsts(c["filial"], cat) if di is not None else []
        if di is not None and 0 <= di < len(dsts):
            return self.fil_screen(crumbs + ("→ " + esc(self.short(dsts[di], 24)),), text, kb, back="kp:%s:0" % cat)
        return self.fil_screen(crumbs, text, kb)

    def fil_in_screen(self, f, page=0, si=None):
        """📥 Keladigan (si — bitta beruvchi)."""
        text, kb = self.incoming_list(f, page, si)
        routes = self.koch_in_routes(f) if si is not None else []
        if si is not None and 0 <= si < len(routes):
            return self.fil_screen(("📥 Keladigan", "← " + esc(self.short(routes[si][0], 24))), text, kb, back="kl:0")
        return self.fil_screen(("📥 Keladigan",), text, kb)

    def fil_day_screen(self, f):
        """📋 Bugungi holat + [❌ Bajarilmaganlar · N] [✅ Bajarilganlar · N] (A / B / C bo'yicha ro'yxatlar)."""
        rows = []
        if self.koch_on and self.koch_has_tasks():
            t = self.koch_abc_summary(filial=f)["total"]
            if t["jami"]:
                rows = [[("❌ Bajarilmaganlar · %d" % t["bajarilmadi"], "kd:n:0"), ("✅ Bajarilganlar · %d" % t["bajarildi"], "kd:d:0")]]
        return self.fil_screen(("📋 Bugungi holat",), self.day_text(f), rows)

    def fil_sklad_screen(self):
        """🏭 Taqsimot (sklad xodimi): xulosa narxlarsiz + [📎 Excel]."""
        s = self.sklad_summary()
        rows = [[("📎 Excel", "kt:x")]] if (s["items"] or s["excluded"]) else []
        return self.fil_screen(("🏭 Taqsimot",), self.sklad_admin_text(s, prices=False), rows)

    # ---- ⏰ Srok / 📦 30+ kun: A / B / C (filial xodimi — FAQAT o'z filiali; filial bazadan, callback'dan emas) ----
    ABC_KINDS = {"s": "⏰ Srok", "o": "📦 30+ kun"}
    ABC_TOP = 20
    ABC_HINT = {"A": "eng qimmat — birinchi navbatda", "B": "o'rtacha narxdagilar", "C": "ko'p, lekin arzon"}

    def fil_abc_rows(self, f, kind):
        """Shu filialning qatorlari: 's' — Srok (qolgan kun ≤ srok_days), 'o' — 30+ kun. Filial fil_key bo'yicha aniq
        solishtiriladi («ДОРИХОНА  ЧИНОЗ» = «ДОРИХОНА ЧИНОЗ» = «ЧИНОЗ»), boshqa filial qatori hech qachon kirmaydi."""
        k = fil_key(f)
        if not k:
            return []
        d = self.data
        rows = d.srok_rows(None, self.srok_days) if kind == "s" else d.old_rows(None)
        return [r for r in rows if fil_key(r["filial"]) == k]

    @staticmethod
    def abc_value(kind):
        """Pul: Srok — qoldiq × kirim narxi; 30+ kun — summa ustuni (bo'lmasa qoldiq × kirim narxi)."""
        if kind == "s":
            return lambda r: (r["ost"] or 0) * (r["prih"] or 0)
        return lambda r: r["summa"] if r.get("summa") else (r["ost"] or 0) * (r["prih"] or 0)

    def fil_abc(self, f, kind):
        """(qatorlar, {sinf: [qatorlar, pul ↓]}, {sinf: {n, sum, share}}, pul funksiyasi)."""
        rows = self.fil_abc_rows(f, kind)
        val = self.abc_value(kind)
        groups = ABC_T.split(rows, val)
        return rows, groups, ABC_T.summary(groups, val), val

    @staticmethod
    def abc_sig(rows):
        """Ro'yxat imzosi: tugma ko'rsatilgandan keyin ma'lumot o'zgargan bo'lsa (Yangilash), eski tugma ishlamaydi."""
        raw = ";".join("%s|%s|%s|%s|%s|%s|%s" % (norm(r["name"]), norm(r["producer"]), r.get("date"), r["kun"], r["ost"], r["prih"],
                                                 r.get("summa")) for r in rows)
        return hashlib.md5(raw.encode("utf-8")).hexdigest()[:8]

    def fil_abc_screen(self, c, kind):
        """Srok / 30+ kun: jami va A / B / C (soni · pul · ulush) + [A · N ta] [B · N ta] [C · N ta] [📥 Hammasi Excel'da]."""
        self.data.reload()
        f = norm(c["filial"])
        sk = self.is_sklad(f)
        rows, groups, sm, val = self.fil_abc(f, kind)
        title = self.ABC_KINDS[kind]
        icon, name = title.split(" ", 1)
        text = "%s <b>%s</b> · 🏬 %s\n%s" % (icon, name, esc(f), ("Muddati %d kundan kam qolgan dorilar" % self.srok_days)
                                           if kind == "s" else "30 kundan ortiq turgan dorilar")
        if not rows:
            return self.fil_screen((title,), text + "\n\n✅ Bunday dori yo'q." + self.foot())
        text += "\nJami: <b>%d</b> ta%s\n" % (len(rows), "" if sk else " · <b>%s</b>" % mln(sum(val(r) for r in rows)))
        for k in ABC_T.CLASSES:
            v = sm[k]
            text += "\n<b>%s</b> · %d ta%s · %d%%\n<i>%s</i>" % (k, v["n"], "" if sk else " · %s" % mln(v["sum"], False), v["share"],
                                                            self.ABC_HINT[k])
        if sm["A"]["n"]:
            text += "\n\nA dan boshlang: %d ta dori pulning %d%% ini ushlab turibdi." % (sm["A"]["n"], sm["A"]["share"])
        text += "\nPul = qoldiq × narx" if kind == "s" else "\nPul = summa (kirim narxida)"
        kb = [[("%s · %d ta" % (k, sm[k]["n"]), "kc:%s:%s" % (kind, k), "primary" if k == "A" else None)
               for k in ABC_T.CLASSES if sm[k]["n"]],
              [("📥 Hammasi Excel'da", "kc:%s:x:%s" % (kind, self.abc_sig(rows)), "success")]]
        return self.fil_screen((title,), text + self.foot(), kb)

    def fil_abc_list(self, c, kind, cls):
        """Bitta sinf: pul bo'yicha birinchi ABC_TOP tasi (2 qatorli band); qolgani — «📥 Qolgan K tasini Excel'da olish»."""
        self.data.reload()
        f = norm(c["filial"])
        sk = self.is_sklad(f)
        rows, groups, sm, val = self.fil_abc(f, kind)
        part = groups.get(cls, [])
        title = self.ABC_KINDS[kind]
        icon, name = title.split(" ", 1)
        text = "%s <b>%s · %s</b> · 🏬 %s" % (icon, name, cls, esc(f))
        kb = []
        if not part:
            text += "\n\nBu guruhda dori yo'q."
        else:
            top, rest = part[:self.ABC_TOP], part[self.ABC_TOP:]
            text += "\nPul bo'yicha · %d tadan %d tasi" % (len(part), len(top))
            red = False
            for i, r in enumerate(top, 1):
                if kind == "s":
                    hot = r["kun"] < 30
                    red = red or hot
                    when = self.kun_text(r["kun"])
                    info = "%s ta · %s" % (qty(r["ost"]), ("%s (%s)" % (r["date"].strftime("%d.%m.%Y"), when)) if r["date"] else when)
                else:
                    hot = False
                    info = "%s ta · %d kun turibdi" % (qty(r["ost"]), r["kun"])
                text += "\n%d. <b>%s</b>%s\n%s%s" % (i, esc(self.short(r["name"], 40)), " 🔴" if hot else "", info,
                                                    "" if sk else " · <b>%s</b>" % money(val(r)))
            text += "\n\n%d / %d ko'rsatildi%s" % (len(top), len(part), "" if sk else " · %s" % mln(sum(val(r) for r in top)))
            if rest:
                text += "\nQolgan %d tasi%s — Excel'da" % (len(rest), "" if sk else " (%s)" % mln(sum(val(r) for r in rest)))
                kb.append([("📥 Qolgan %d tasini Excel'da olish" % len(rest), "kc:%s:%s:x:%s" % (kind, cls, self.abc_sig(part)),
                            "success")])
            if red:
                text += "\n🔴 — muddatiga 30 kundan kam qolgan"
        kb.append([("◀ A/B/C", "kc:%s" % kind), ("🏠 Menyu", "kh")])
        return self.screen((title, cls), cut_html(text + self.foot()), kb)

    def fil_abc_excel(self, cid, c, mid, kind, cls, sig):
        """Excel — alohida xabar (NAVIGATSIYA 5-qoida), panel fayl ostiga. cls=None — hamma sinflar (A, B, C tartibida).
        Ro'yxat tugma ko'rsatilgandan keyin o'zgargan bo'lsa (sig), fayl yuborilmaydi — yangi ro'yxat ko'rsatiladi."""
        self.data.reload()
        f = norm(c["filial"])
        rows, groups, sm, val = self.fil_abc(f, kind)
        part = groups.get(cls, []) if cls else [r for k in ABC_T.CLASSES for r in groups[k]]
        if self.abc_sig(part if cls else rows) != sig:
            scr = self.fil_abc_list(c, kind, cls) if cls else self.fil_abc_screen(c, kind)
            return self.panel_show(cid, mid, *scr, note="⚠ <b>Ma'lumot yangilandi</b>, fayl yuborilmadi. Ro'yxatni tekshirib, qaytadan bosing.")
        if not part:
            return self.panel_note(cid, mid, "➖ Bu bo'lim bo'yicha ma'lumot topilmadi.")
        sinf = {id(r): k for k in ABC_T.CLASSES for r in groups[k]}
        name = "Srok" if kind == "s" else "30+ kun"
        content = build_abc(kind, part, lambda r: sinf[id(r)], val, prices=not self.is_sklad(f),
                            title="%s %s" % (name, cls) if cls else name)
        short = re.sub(r"[^\w]+", "_", f.split()[-1] if f.split() else f, flags=re.U).strip("_")[:20] or "filial"
        fname = "%s_%s_%s_%s.xlsx" % ("Srok" if kind == "s" else "30kun", cls or "ABC", short, self._today())
        cap = "%s · %s · %s\nFaylda: %d ta%s" % (self.ABC_KINDS[kind], cls or "A / B / C", esc(f), len(part),
                                                 "" if self.is_sklad(f) else " · %s" % mln(sum(val(r) for r in part)))
        if cls and len(part) > self.ABC_TOP:
            cap += "\n%d–%d tasi — faqat shu faylda" % (self.ABC_TOP + 1, len(part))
        self.touch(cid, "excel:abc:" + kind)
        return self.panel_doc(cid, fname, content, cut_html("%s\n🕒 Ma'lumot: %s" % (cap, dt_text(self.data.when())), 1000))

    def fil_card_screen(self, c, tid, page=0, di=None):
        """Dori kartochkasi: «🏠 › 🔴 SHOSHILINCH › #12» (oluvchida — «📥 Keladigan › #12»)."""
        text, kb = self.koch_item_card(c, tid, page, di)
        t = self.db.execute("SELECT src, cat FROM koch_tasks WHERE id=?", (tid,)).fetchone()
        if kb is None or not t:
            return self.fil_screen(("📥 Keladigan",), text)
        first = self.cat_crumb(t["cat"] if t["cat"] in KOCH_CAT else "oddiy") if t["src"] == norm(c["filial"]) else "📥 Keladigan"
        return self.screen((first, "#%d" % tid), text, kb)

    def filial_view(self, f, kind):
        """Admin «mijoz ko'rinishi» uchun matn (kind: toifa kaliti | kz/ko (eski) | kin | kday)."""
        kind = {"kz": "shoshilinch", "ko": "oddiy"}.get(kind, kind)
        c = {"filial": f, "chat_id": 0}
        if kind in KOCH_CAT:
            return self.koch_list_text(c, kind, 0)[0]
        if kind == "kin":
            return self.incoming_list(norm(f), 0)[0]
        return self.day_text(norm(f))

    def incoming_list(self, f, page=0, si=None):
        """📥 Keladigan: avval berilganlar (📥 Oldim kutilmoqda), keyin hali berilmaganlar (faqat ko'rish).
        si: faqat shu beruvchi (koch_in_routes dagi raqam) — farqni belgilash uchun."""
        f = norm(f)
        h = "📥 <b>Keladigan</b> · 🏬 %s" % esc(f)
        icon = lambda cat: KOCH_CAT.get(cat, KOCH_CAT["oddiy"])["icon"]
        if not self.koch_on:
            rows = self.data.koch_in(f)
            return h + "\n\n" + ("\n".join("%d. %s %s ← %s · %s qadoq" % (i, icon(r["cat"]), esc(self.short(r["name"], 34)),
                                                                           esc(norm(r["src"])), qty(r["qty"]))
                                           for i, r in enumerate(rows[:20], 1)) or "Hozircha keladigan tovar yo'q.") + self.foot(), None
        routes = self.koch_in_routes(f)
        one = si if si is not None and 0 <= si < len(routes) else None
        given = routes[one][1] if one is not None else [t for _, tasks in routes for t in tasks]
        waiting = self.koch_tasks(self._today(), dst=f, statuses=("yangi",))
        h += "\nYo'lda (berilgan): <b>%d</b> · kutilmoqda: <b>%d</b>" % (len(given), len(waiting))
        tail = "" if one is None else ":%d" % one
        kb, text = [], h
        abc = {}
        cls = lambda t: abc.setdefault(t["sana"], self.koch_abc(t["sana"])).get(t["id"], "C")
        for n, (src, tasks) in enumerate(routes):  # ommaviy qabul: beruvchi bo'yicha bitta tugma
            if one is None or n == one:
                sig = self.koch_sig((t["id"], t["qty_given"]) for t in tasks)
                kb.append([(("📥 Hammasini oldim · %d ta" % len(tasks)) if one is not None else "📥 ← %s · %d ta" % (bl(self.short(src, 26)), len(tasks)),
                            "kv:%d:%s" % (n, sig))])
        per = max(1, self.kint("koch_page_size"))
        if given:
            pages = (len(given) - 1) // per + 1
            page = min(max(page, 0), pages - 1)
            part = given[page * per:(page + 1) * per]
            lines = ["%d. %s %s · <b>%s</b> qadoq ← %s%s · <code>%s</code>%s" % (
                i, icon(t["cat"]), esc(self.short(t["name"], 34)), qty(t["qty_given"]), esc(self.short(t["src"], 22)),
                (" <i>(reja %s)</i>" % qty(t["qty_plan"])) if t["qty_given"] < t["qty_plan"] else "", cls(t),
                " · 🤖 tizim" if t["auto_by"] == "tizim" and not t["given_by"] else "")
                     for i, t in enumerate(part, page * per + 1)]
            text += "\n\n<b>📦 Berildi — 📥 Oldim kutilmoqda</b> (%d/%d):\n%s" % (page + 1, pages, "\n".join(lines))
            btns = [(str(i), "ki:%d:0" % t["id"]) for i, t in enumerate(part, page * per + 1)]
            kb += [btns[i:i + 5] for i in range(0, len(btns), 5)]
            nav = []
            if page > 0:
                nav.append(("◀", "kl:%d%s" % (page - 1, tail)))
            if page < pages - 1:
                nav.append(("▶", "kl:%d%s" % (page + 1, tail)))
            if nav:
                kb.append(nav)
        if waiting:
            text += "\n\n<b>⏳ Kutilmoqda</b> (hali berilmagan):\n" + "\n".join(
                "• %s %s · %s qadoq ← %s · <code>%s</code>" % (icon(t["cat"]), esc(self.short(t["name"], 34)), qty(t["qty_plan"]),
                                                               esc(self.short(t["src"], 22)), cls(t)) for t in waiting[:10]) + (
                ("\n… yana %d ta" % (len(waiting) - 10)) if len(waiting) > 10 else "")
        if not given and not waiting:
            text += "\n\nHozircha sizga keladigan tovar yo'q."
        return cut_html(text), (inline_kb(kb) if kb else None)

    def incoming_text(self, filial):
        return self.incoming_list(norm(filial))[0]

    def day_abc_text(self, f, s=None):
        """📋 Bugungi holat boshi: A / B / C × Bajarildi / Qoldi (<pre> jadval, shu filial beradiganlar) + bajarilmagan A lar."""
        s = s or self.koch_abc_summary(filial=f)
        tot, sk = s["total"], self.is_sklad(f)
        if not tot["jami"]:
            return "Ko'chirish natijasi · A / B / C\nBugun sizdan beriladigan topshiriq yo'q."
        cy = UK.lat2cyr
        ta = lambda n: cy("%d ta" % n)
        lines = [" " * 9 + cy("Bajarildi").rjust(10) + cy("Qoldi").rjust(8)]
        for k in ABC_T.CLASSES:
            v = s["classes"][k]
            lines.append(("%s %s" % (k, ("(%d%%)" % ABC_T.SHARE[k]).rjust(5))).ljust(9) + ta(v["bajarildi"]).rjust(10) +
                         ta(v["bajarilmadi"]).rjust(8))
        lines.append(cy("Jami").ljust(9) + ta(tot["bajarildi"]).rjust(10) + ta(tot["bajarilmadi"]).rjust(8))
        text = ("Ko'chirish natijasi · A / B / C\nJami: <b>%d</b> ta · bajarildi <b>%d</b> · qoldi <b>%d</b>\n<pre>%s</pre>" % (
            tot["jami"], tot["bajarildi"], tot["bajarilmadi"], esc("\n".join(lines))))
        left = [i for i in s["not_done"] if i["cls"] == "A"]
        if left:
            text += "\n❌ <b>Bajarilmaganlar (A):</b>\n" + "\n".join("%d. <b>%s</b> %s%s\n→ %s · %s ta%s" % (
                n, esc(self.short(i["name"], 36)), KOCH_CAT.get(i["cat"], KOCH_CAT["oddiy"])["icon"],
                " · %s" % self.ST_NAME[i["status"]] if i["status"] != "yangi" else "", esc(self.short(i["dst"], 24)), qty(i["qty"]),
                "" if sk else " · %s" % money(i["sum"])) for n, i in enumerate(left[:5], 1))
            if len(left) > 5:
                text += "\n… yana %d ta" % (len(left) - 5)
            if not sk:
                text += "\nQolgan A: <b>%s</b>" % mln(s["classes"]["A"]["sum_bajarilmadi"])
        return text

    def fil_day_list(self, f, which, page=0):
        """❌ Bajarilmaganlar (which='n') / ✅ Bajarilganlar ('d') — bugun shu filial beradiganlar, A / B / C guruhlab, sahifalab."""
        f = norm(f)
        done = which == "d"
        title = "✅ Bajarilganlar" if done else "❌ Bajarilmaganlar"
        s = self.koch_abc_summary(filial=f)
        items = s["done"] if done else s["not_done"]
        sk = self.is_sklad(f)
        text = "📋 <b>Bugungi holat</b> · %s · 🏬 %s\nJami: <b>%d</b> ta%s" % (
            title, esc(f), len(items), "" if sk or not items else " · <b>%s</b>" % mln(sum(i["sum"] for i in items)))
        kb = []
        if not items:
            text += "\n\n" + ("Hozircha bajarilgan topshiriq yo'q." if done else "✅ Hammasi bajarildi.")
        else:
            per = 10
            pages = (len(items) - 1) // per + 1
            page = min(max(page, 0), pages - 1)
            cur = None
            for n, i in enumerate(items[page * per:(page + 1) * per], page * per + 1):
                if i["cls"] != cur:
                    cur = i["cls"]
                    text += "\n\n<b>%s</b> · %d ta" % (cur, sum(1 for x in items if x["cls"] == cur))
                q = i["qty_given"] if done and i["qty_given"] is not None else i["qty"]
                text += "\n%d. <code>%s</code> <b>%s</b> %s\n→ %s · %s ta%s · %s%s" % (
                    n, cur, esc(self.short(i["name"], 36)), KOCH_CAT.get(i["cat"], KOCH_CAT["oddiy"])["icon"],
                    esc(self.short(i["dst"], 24)), qty(q), "" if sk else " · %s" % money(i["sum"]), self.ST_NAME[i["status"]],
                    ((" " + i["given_at"]) if done and i["given_at"] else "") + (" · 🤖 tizim" if done and i.get("by") == "tizim" else ""))
            if pages > 1:
                text += "\n\n<i>%d-sahifa / %d</i>" % (page + 1, pages)
                nav = []
                if page > 0:
                    nav.append(("◀", "kd:%s:%d" % (which, page - 1)))
                if page < pages - 1:
                    nav.append(("▶", "kd:%s:%d" % (which, page + 1)))
                kb.append(nav)
        return self.screen(("📋 Bugungi holat", title), cut_html(text + self.foot()), kb, back="kd", home="kh")

    def day_text(self, f):
        """📋 Bugungi holat: A / B / C (bajarildi / qoldi), toifalar kesimida beradi / oladi, kechikkanlar, muddatgacha vaqt."""
        f = norm(f)
        h = "📋 <b>Bugungi holat</b> · 🏬 %s" % esc(f)
        if not self.koch_on or not self.koch_has_tasks():
            return h + "\n\nBugungi topshiriq hali shakllanmagan."
        d = self._today()
        out = [h, "\n" + self.day_abc_text(f)]
        late_all = 0
        for k in KOCH_CATS:
            give = [t for t in self.koch_tasks(d, src=f, cat=k["key"]) if t["status"] != "bekor"]
            take = [t for t in self.koch_tasks(d, dst=f, cat=k["key"]) if t["status"] != "bekor"]
            late = sum(1 for t in self.koch_tasks(d, cat=k["key"], statuses=("yangi", "berildi"))
                       if self.task_info(t)["late"] and self.task_info(t)["guilty_filial"] == f)
            late_all += late
            g_done = sum(1 for t in give if t["status"] in ("berildi", "olindi"))
            g_rad = sum(1 for t in give if t["status"] == "rad")
            t_done = sum(1 for t in take if t["status"] == "olindi")
            t_left = sum(1 for t in take if t["status"] in ("yangi", "berildi"))
            out.append("\n%s <b>%s</b>\n📤 beradi: jami %d · berdi %d · qoldi %d%s\n📥 oladi: jami %d · oldi %d · qoldi %d%s\n%s" % (
                k["icon"], k["title"], len(give), g_done, len(give) - g_done - g_rad, (" · rad %d" % g_rad) if g_rad else "",
                len(take), t_done, t_left, ("\n⛔ siz tomondan kechikkan: <b>%d</b>" % late) if late else "",
                self.deadline_line(k["key"], short=True)))
        out.append("\n" + ("⛔ Jami kechikkan: <b>%d</b>" % late_all if late_all else "✅ Kechikkan yo'q"))
        return "\n".join(out)

    # ---- B4.5: Excel (toifa bo'yicha) ----
    def filial_excel(self, cid, c, cat, di=None, mid=None):
        """Faqat shu filial va toifa. 1-varaq — ro'yxat (bo'yalgan), 2-varaq «Ko'chirilmaydi» — chiqarilganlar sababi bilan.
        di: faqat shu yo'nalish (ommaviy berishdan oldin tovar yig'ish uchun). Fayl — alohida xabar, panel uning ostiga."""
        self.data.reload()
        f = norm(c["filial"])
        k = KOCH_CAT[cat]
        rows, excluded = self.koch_view(f, cat)
        dsts = self.koch_route_dsts(f, cat)
        if di is not None and 0 <= di < len(dsts):
            rows = [r for r in rows if norm(r["dst"]) == dsts[di]]
            excluded = [(r, why) for r, why in excluded if norm(r["dst"]) == dsts[di]]
        if not rows and not excluded:
            return self.panel_note(cid, mid, "➖ Bu bo'lim bo'yicha ma'lumot topilmadi.")
        short = re.sub(r"[^\w]+", "_", f.split()[-1] if f.split() else f, flags=re.U).strip("_")[:20] or "filial"
        fname = "%s_%s_%s.xlsx" % (k["title"], short, datetime.date.today().isoformat())
        content = build_koch_cat_xlsx(rows, excluded, cat, f, warn_days=self.kint("koch_srok_warn_days"),
                                      with_status=self.koch_on and self.koch_has_tasks(), prices=not self.is_sklad(f),
                                      cls_of=self.koch_cls(rows))
        self.panel_doc(cid, fname, content, cut_html("%s <b>%s</b> · %s\nRo'yxatda: %d ta%s\n🕒 Ma'lumot: %s" % (
            k["icon"], k["title"], esc(f), len(rows), self.koch_sum(f, rows, bold=False), dt_text(self.data.when())), 1000))

    # ---- B5.3: amallar: Berdim / Boshqa son / Bera olmayman / Oldim / Kam keldi ----
    def _actor(self, cid):
        c = self.client(cid)
        return (c["name"] or c["phone"] or str(cid)) if c else str(cid)

    def koch_give(self, cid, c, tid, n=None):
        """✅ Berdim (n=None: reja bo'yicha) yoki ✏️ Boshqa son. Faqat beruvchi filial (bazadan), bugungi/kechagi, «yangi»."""
        f = norm(c["filial"])
        t = self.db.execute("SELECT * FROM koch_tasks WHERE id=?", (tid,)).fetchone()
        if not t or t["src"] != f or t["sana"] not in (self._today(), self._yesterday()) or t["status"] != "yangi":
            return None
        n = t["qty_plan"] if n is None else n
        if not (0 < n <= t["qty_plan"]):
            return None
        ts = now().isoformat(timespec="seconds")
        d1 = self.koch_dst_now(tid)  # oluvchining «Berdim» paytidagi qoldig'i (yuklangan Sale ma'lumotidan)
        with self.dblock:
            cur = self.db.execute("UPDATE koch_tasks SET status='berildi', qty_given=?, given_by=?, given_at=?, dst_stock1=? "
                                  "WHERE id=? AND status='yangi' AND src=?", (n, cid, ts, d1, tid, f))
            if cur.rowcount != 1:
                self.db.rollback()
                return None
            self._event(tid, "berildi", cid, self._actor(cid), n, None if n == t["qty_plan"] else "reja %s" % qty(t["qty_plan"]), ts)
            self.db.commit()
        k = KOCH_CAT.get(t["cat"], KOCH_CAT["oddiy"])
        self._notify_filial(t["dst"], "📦 <b>%s dan keladi</b>: %s · <b>%s</b> qadoq%s\n%s %s · kelganda «📥 Oldim» ni bosing." % (
            esc(t["src"]), esc(t["name"]), qty(n), (" <i>(reja %s)</i>" % qty(t["qty_plan"])) if n < t["qty_plan"] else "",
            k["icon"], k["title"]), inline_kb([[("📥 Oldim", "ka:%d" % tid), ("⚠ Kam keldi", "kk:%d" % tid)]]))
        return True

    def koch_refuse(self, cid, c, tid, reason):
        f = norm(c["filial"])
        t = self.db.execute("SELECT * FROM koch_tasks WHERE id=?", (tid,)).fetchone()
        if not t or t["src"] != f or t["sana"] not in (self._today(), self._yesterday()) or t["status"] != "yangi":
            return None
        ts = now().isoformat(timespec="seconds")
        reason = (reason or "").strip()[:200] or "—"
        with self.dblock:
            cur = self.db.execute("UPDATE koch_tasks SET status='rad', reason=?, given_by=?, given_at=? WHERE id=? AND status='yangi' "
                                  "AND src=?", (reason, cid, ts, tid, f))
            if cur.rowcount != 1:
                self.db.rollback()
                return None
            self._event(tid, "rad", cid, self._actor(cid), None, reason, ts)
            self.db.commit()
        self._notify_filial(t["dst"], "❌ <b>%s bera olmaydi</b>: %s · %s qadoq\nSabab: %s" % (
            esc(t["src"]), esc(t["name"]), qty(t["qty_plan"]), esc(reason)))
        return True

    def koch_receive(self, cid, c, tid, n=None):
        """📥 Oldim (n=None: berilgan son) yoki ⚠ Kam keldi. Faqat oluvchi filial (bazadan), «berildi»."""
        f = norm(c["filial"])
        t = self.db.execute("SELECT * FROM koch_tasks WHERE id=?", (tid,)).fetchone()
        if not t or t["dst"] != f or t["sana"] not in self._recv_dates() or t["status"] != "berildi":
            return None
        n = t["qty_given"] if n is None else n
        if not (0 < n <= t["qty_given"]):
            return None
        ts = now().isoformat(timespec="seconds")
        with self.dblock:
            cur = self.db.execute("UPDATE koch_tasks SET status='olindi', qty_received=?, received_by=?, received_at=? "
                                  "WHERE id=? AND status='berildi' AND dst=?", (n, cid, ts, tid, f))
            if cur.rowcount != 1:
                self.db.rollback()
                return None
            self._event(tid, "olindi", cid, self._actor(cid), n, None if n == t["qty_given"] else "kam keldi", ts)
            self.db.commit()
        self._ret_cache = None  # borib-kelish filtri yangilanadi
        self._notify_filial(t["src"], "✅ <b>%s qabul qildi</b>: %s · %s qadoq%s" % (
            esc(t["dst"]), esc(t["name"]), qty(n), (" ⚠ <b>kam keldi</b> (berilgan %s)" % qty(t["qty_given"])) if n < t["qty_given"] else ""))
        return True

    def filial_state_input(self, cid, c, st, text):
        """Yozilgan son / sabab. Javob — yangi panel (dori kartochkasi izoh bilan); eski panel o'chiriladi."""
        kind, tid = st[0], st[1]
        t = self.db.execute("SELECT * FROM koch_tasks WHERE id=?", (tid,)).fetchone()
        if not t:
            return self.panel_resend(cid, *self.filial_home(c), note="Topshiriq topilmadi.")
        busy = "Bu topshiriq allaqachon bajarilgan yoki sizga tegishli emas."
        if kind == "koch_rad":
            ok = self.koch_refuse(cid, c, tid, text)
            return self.panel_resend(cid, *self.fil_card_screen(c, tid), note=(
                "❌ Belgilandi: bera olmaysiz — %s" % esc(t["name"])) if ok else busy)
        try:
            n = float(text.replace(",", ".").strip())
        except ValueError:
            n = -1
        limit = t["qty_plan"] if kind == "koch_qty" else (t["qty_given"] or 0)
        if not (0 < n <= limit) or n != int(n):
            return self.ask(cid, None, (self.cat_crumb(t["cat"]) if kind == "koch_qty" else "📥 Keladigan", "#%d" % tid),
                            "Butun son yozing: 1 dan %s gacha. Bekor qilish: /cancel" % qty(limit), st, back="ki:%d:0" % tid)
        ok = self.koch_give(cid, c, tid, n) if kind == "koch_qty" else self.koch_receive(cid, c, tid, n)
        if not ok:
            return self.panel_resend(cid, *self.fil_card_screen(c, tid), note=busy)
        return self.panel_resend(cid, *self.fil_card_screen(c, tid), note=(
            "✅ Berildi: %s · %s qadoq" if kind == "koch_qty" else "📥 Olindi: %s · %s qadoq") % (esc(t["name"]), qty(n)))

    def on_filial_callback(self, cid, mid, data):
        """Faqat faol filial xodimi. Filial callback dan emas, bazadan olinadi; id, sana va holat tekshiriladi (jim rad).
        Har tugma panelning o'zini o'zgartiradi (panel_show); Excel — alohida xabar, panel uning ostiga."""
        c = self.client(cid)
        if not c or c["status"] != "active" or not self.is_filial(c):
            return
        f = norm(c["filial"])
        p = data.split(":")
        busy = "Bu topshiriq allaqachon bajarilgan yoki sizga tegishli emas."
        show = lambda scr, note=None: self.panel_show(cid, mid, scr[0], scr[1], note)
        self.state.pop(cid, None)  # boshqa tugma: kutilayotgan son / sabab kiritish bekor
        try:
            act = p[0]
            if act == "kh":
                return show(self.filial_home(c))
            if act == "kd" and len(p) > 1 and p[1] in ("n", "d"):  # bajarilmaganlar / bajarilganlar (A / B / C)
                return show(self.fil_day_list(f, p[1], int(p[2]) if len(p) > 2 else 0))
            if act == "kc":  # ⏰ Srok / 📦 30+ kun: A / B / C — faqat o'z filiali (filial callback dan emas, bazadan)
                kind = p[1] if len(p) > 1 else ""
                if kind not in self.ABC_KINDS:
                    return None
                if len(p) == 2:
                    self.touch(cid, "abc:" + kind)
                    return show(self.fil_abc_screen(c, kind))
                if p[2] == "x":
                    return self.fil_abc_excel(cid, c, mid, kind, None, p[3] if len(p) > 3 else "")
                if p[2] in ABC_T.CLASSES:
                    if len(p) > 3 and p[3] == "x":
                        return self.fil_abc_excel(cid, c, mid, kind, p[2], p[4] if len(p) > 4 else "")
                    return show(self.fil_abc_list(c, kind, p[2]))
                return None
            if act == "kd" or (act == "kn" and p[1] == "d"):
                self.touch(cid, "kday")
                return show(self.fil_day_screen(f))
            if act == "kt":  # 🏭 Taqsimot — faqat sklad xodimi
                if not self.is_sklad(f):
                    return None
                if len(p) > 1 and p[1] == "x":
                    s = self.sklad_summary()
                    if not s["items"] and not s["excluded"]:
                        return self.panel_note(cid, mid, "➖ Taqsimot yo'q.")
                    self.touch(cid, "excel:ksk")
                    return self.panel_doc(cid, "Sklad_taqsimot_%s.xlsx" % self._today(), build_sklad(s, prices=False),
                                          cut_html(self.sklad_admin_text(s, prices=False), 1000))
                self.touch(cid, "ksk")
                return show(self.fil_sklad_screen())
            if act in ("kp", "kn"):  # kn — eski «topshiriq tayyor» xabaridagi tugmalar
                cat = {"z": "shoshilinch", "o": "oddiy"}.get(p[1], p[1])
                if cat not in KOCH_CAT:
                    return None
                return show(self.fil_list_screen(c, cat, int(p[2]) if len(p) > 2 else 0, int(p[3]) if len(p) > 3 else None))
            if act == "kx":
                if p[1] in KOCH_CAT:
                    self.touch(cid, "excel:" + p[1])
                    return self.filial_excel(cid, c, p[1], mid=mid)
                return None
            if act == "kgx":
                if p[1] in KOCH_CAT:
                    self.touch(cid, "excel:" + p[1])
                    return self.filial_excel(cid, c, p[1], int(p[2]), mid=mid)
                return None
            if act == "kl":
                self.touch(cid, "kin")
                return show(self.fil_in_screen(f, int(p[1]) if len(p) > 1 else 0, int(p[2]) if len(p) > 2 else None))
            changed = "⚠ <b>Ro'yxat o'zgardi</b>, hech narsa belgilanmadi. Yangi ro'yxatni tekshirib, qaytadan bosing."
            if act in ("kg", "kgy"):
                cat, i, sig = p[1], int(p[2]), p[3]
                if cat not in KOCH_CAT:
                    return None
                if act == "kg":
                    card = self.koch_route_card(f, cat, i, sig)
                else:
                    self.touch(cid, "kb")
                    res = self.koch_give_route(cid, c, cat, i, sig)
                    if res:
                        return show(self.fil_screen((self.cat_crumb(cat), "→ " + esc(self.short(res[0], 24))),
                                                    "✅ <b>Berildi</b> → %s\n%d ta dori · %s qadoq\nOluvchiga xabar yuborildi." % (
                                                        esc(res[0]), len(res[1]), qty(sum(r["qty"] for r in res[1]))),
                                                    back="kp:%s:0" % cat))
                    card = None
                if card:
                    dsts = self.koch_route_dsts(f, cat)
                    dst = dsts[i] if 0 <= i < len(dsts) else ""
                    return show(self.screen((self.cat_crumb(cat), "→ " + esc(self.short(dst, 24))), card[0], card[1]))
                return show(self.fil_list_screen(c, cat), changed)
            if act in ("kv", "kvy"):
                i, sig = int(p[1]), p[2]
                if act == "kv":
                    card = self.koch_in_card(f, i, sig)
                else:
                    self.touch(cid, "ka")
                    res = self.koch_receive_route(cid, c, i, sig)
                    if res:
                        return show(self.fil_screen(("📥 Keladigan", "← " + esc(self.short(res[0], 24))),
                                                    "📥 <b>Olindi</b> ← %s\n%d ta dori · %s qadoq" % (
                                                        esc(res[0]), len(res[1]), qty(sum(t["qty_given"] for t in res[1]))), back="kl:0"))
                    card = None
                if card:
                    return show(self.screen(("📥 Keladigan", "📥 Qabul"), card[0], card[1]))
                return show(self.fil_in_screen(f), changed)
            tid = int(p[1])
            t = self.db.execute("SELECT * FROM koch_tasks WHERE id=?", (tid,)).fetchone()
            if not t or f not in (t["src"], t["dst"]):
                return self.panel_note(cid, mid, busy)
            crumbs = (self.cat_crumb(t["cat"] if t["cat"] in KOCH_CAT else "oddiy") if t["src"] == f else "📥 Keladigan", "#%d" % tid)
            if act == "ki":
                return show(self.fil_card_screen(c, tid, int(p[2]) if len(p) > 2 else 0, int(p[3]) if len(p) > 3 else None))
            if act == "kb":
                self.touch(cid, "kb")
                ok = self.koch_give(cid, c, tid)
                return show(self.fil_card_screen(c, tid), ("✅ Berildi: %s" % esc(t["name"])) if ok else busy)
            if act == "kq":
                if t["src"] != f or t["status"] != "yangi":
                    return self.panel_note(cid, mid, busy)
                return self.ask(cid, mid, crumbs + ("✏️ Boshqa son",), "✏️ <b>%s</b>\nReja: %s qadoq. Nechta berdingiz? Son yozing 👇 (1–%s).\n"
                                "Bekor qilish: /cancel" % (esc(t["name"]), qty(t["qty_plan"]), qty(t["qty_plan"])),
                                ("koch_qty", tid), back="ki:%d:0" % tid)
            if act == "kr":
                if t["src"] != f or t["status"] != "yangi":
                    return self.panel_note(cid, mid, busy)
                return show(self.fil_screen(crumbs + ("❌",), "❌ <b>%s</b> — nima uchun bera olmaysiz?" % esc(t["name"]),
                                            [[(r, "krr:%d:%d" % (tid, i))] for i, r in enumerate(self.RAD_REASONS)], back="ki:%d:0" % tid))
            if act in ("krr", "kxr"):
                i = int(p[2])
                if i == len(self.RAD_REASONS) - 1:  # Boshqa: matn so'raladi
                    if t["src"] != f or t["status"] != "yangi":
                        return self.panel_note(cid, mid, busy)
                    return self.ask(cid, mid, crumbs + ("❌",), "❌ <b>%s</b>\nSababini yozing 👇 Bekor qilish: /cancel" % esc(t["name"]),
                                    ("koch_rad", tid), back="ki:%d:0" % tid)
                self.touch(cid, "kx")
                ok = self.koch_refuse(cid, c, tid, self.RAD_REASONS[i])
                return show(self.fil_card_screen(c, tid), ("❌ Belgilandi: %s — %s" % (esc(t["name"]), self.RAD_REASONS[i])) if ok else busy)
            if act == "ka":
                self.touch(cid, "ka")
                ok = self.koch_receive(cid, c, tid)
                if len(p) == 2 and not self.is_panel(cid, mid):  # bildirishnomadagi tugma: o'sha xabar natijaga aylanadi
                    return self.show(cid, mid, ("📥 Olindi: %s · %s qadoq" % (esc(t["name"]), qty(t["qty_given"] or 0))) if ok else busy,
                                     inline_kb([[("🏠 Menyu", "kh")]]))
                return show(self.fil_card_screen(c, tid), ("📥 Olindi: %s" % esc(t["name"])) if ok else busy)
            if act == "kk":
                if t["dst"] != f or t["status"] != "berildi":
                    return self.panel_note(cid, mid, busy)
                return self.ask(cid, mid, crumbs + ("⚠ Kam keldi",), "⚠ <b>%s</b>\nBerilgan: %s qadoq. Nechta keldi? Son yozing 👇 (1–%s).\n"
                                "Bekor qilish: /cancel" % (esc(t["name"]), qty(t["qty_given"]), qty(t["qty_given"])),
                                ("koch_kam", tid), back="ki:%d:0" % tid)
        except (ValueError, KeyError, IndexError):
            return None

    # ---- B6: eslatmalar ----
    URGENT = ("shoshilinch", "assortiment")

    def filial_open(self, sana, cat=None, cats=None):
        """{filial: [ochiq topshiriqlar]} — filial hozir javobgar bo'lganlar (beruvchi «yangi», oluvchi «berildi»)."""
        out = {}
        for t in self.koch_tasks(sana, cat=cat, statuses=("yangi", "berildi")):
            if cats and t["cat"] not in cats:
                continue
            out.setdefault(t["src"] if t["status"] == "yangi" else t["dst"], []).append(t)
        return out

    def remind_filial(self, f, head=None):
        """Filial xodimlariga botda eslatma (ochiq topshiriqlari bilan). Yuborilgan odamlar soni."""
        rows = self.filial_open(self._today()).get(norm(f), [])
        if not rows:
            return 0
        give = sum(1 for t in rows if t["src"] == norm(f))
        text = (head or "🔔 <b>Eslatma: ko'chirish</b>") + "\n🏬 %s\n\n📤 Berish kerak: <b>%d</b> ta\n📥 Qabul qilish kerak: <b>%d</b> ta\n%s" % (
            esc(norm(f)), give, len(rows) - give, self.deadline_line("shoshilinch"))
        return self._notify_filial(f, text, self._remind_kb())

    @staticmethod
    def _remind_kb():
        return inline_kb([[("%s %s" % (k["icon"], k["title"]), "kp:%s:0" % k["key"]) for k in KOCH_CATS],
                          [("📥 Keladigan", "kl:0"), ("🏠 Menyu", "kh")]])

    def _urgent_deadlines(self):
        """{muddat vaqti: [toifalar]} — SHOSHILINCH va ASSORTIMENT (muddati bir xil bo'lsa — bitta eslatma)."""
        out = {}
        for cat in self.URGENT:
            dl = self.koch_deadline(cat)
            if dl:
                out.setdefault(dl, []).append(cat)
        return out

    def _koch_first_day(self, d):
        """Birinchi o'rnatilgan kuni (kun o'rtasida) eslatma / hisobot yo'q — ertangi kundan."""
        return self.setting("koch_installed") == d and self.setting("koch_rem1_date") == d

    def koch_reminders_due(self):
        n, d = now(), self._today()
        if self._koch_first_day(d):
            return False
        rb = datetime.timedelta(minutes=self.kint("koch_remind_before"))
        for dl in self._urgent_deadlines():
            hm = dl.strftime("%H%M")
            if (n >= dl - rb and self.setting("koch_rem1_%s" % hm) != d) or (n >= dl and self.setting("koch_rem2_%s" % hm) != d):
                return True
        rep = self.koch_deadline("oddiy") or datetime.datetime.combine(n.date(), datetime.time(20, 0))
        return n >= rep and self.setting("koch_rem3_date") != d

    def koch_reminders(self):
        """Muddatdan oldin (09:30), muddat o'tganda (10:00) — SHOSHILINCH/ASSORTIMENT; kun oxiri hisoboti (ODDIY muddatidan keyin).
        Har biri kuniga bir marta."""
        if not self.koch_on:
            return
        n, d = now(), self._today()
        rb = datetime.timedelta(minutes=self.kint("koch_remind_before"))
        if self._koch_first_day(d):
            return
        names = lambda cats: " / ".join(KOCH_CAT[c]["title"] for c in cats)
        for dl, cats in sorted(self._urgent_deadlines().items()):
            hm = dl.strftime("%H%M")
            if n >= dl - rb and self.setting("koch_rem1_%s" % hm) != d:
                self.set_setting("koch_rem1_%s" % hm, d)
                if n < dl and self.kset("koch_rem_before_on") == "1":
                    for f, rows in self.filial_open(d, cats=cats).items():
                        self._notify_filial(f, "⏰ <b>%s qoldi</b>: %d ta ko'chirish bajarilmagan (%s)\n🏬 %s" % (
                            self.fmt_min((dl - n).total_seconds() / 60), len(rows), names(cats), esc(f)), self._remind_kb())
            if n >= dl and self.setting("koch_rem2_%s" % hm) != d:
                self.set_setting("koch_rem2_%s" % hm, d)
                if self.kset("koch_rem_overdue_on") == "1":
                    if not self.koch_has_tasks(d):
                        self._admins_send("⚠ <b>Bugungi ko'chirish topshirig'i yaratilmadi</b>: %s dan keyin muvaffaqiyatli "
                                          "«Yangilash» bo'lmadi. Panel → 🔁 Ko'chirish → «Qayta shakllantirish»." % self.kset("koch_task_time"))
                    else:
                        late = self.filial_open(d, cats=cats)
                        for f, rows in late.items():
                            self._notify_filial(f, "⛔ <b>Muddat o'tdi</b>: %d ta ko'chirish bajarilmagan (%s)\n🏬 %s" % (
                                len(rows), names(cats), esc(f)), self._remind_kb())
                        self._admins_send(self.koch_admin_text(d, head="⛔ <b>%s muddati (%s) o'tdi</b>" % (
                            names(cats), dl.strftime("%H:%M"))) + "\n\n📊 Batafsil: server paneli → 🔁 Ko'chirish (dashboard)")
        rep = self.koch_deadline("oddiy") or datetime.datetime.combine(n.date(), datetime.time(20, 0))
        if n >= rep and self.setting("koch_rem3_date") != d:
            self.set_setting("koch_rem3_date", d)
            if self.alert_on("alert_daily_report") and self.koch_has_tasks(d):
                self._admins_send(self.koch_day_report(d))

    def _admins_send(self, text, markup=None):
        """Adminlarga bildirishnoma (alohida xabar), ostida «🏠 Menyu» — bosilsa panel shu xabar ostiga tushadi."""
        markup = markup or inline_kb([[("🏠 Menyu", "a:home")]])
        for a in self.admins:
            try:
                self.tg.send(a, cut_html(text), markup)
            except Exception:  # noqa
                logging.exception("Adminga xabar yuborib bo'lmadi: %s", a)

    def koch_admin_text(self, sana=None, head=None):
        """Filiallar bo'yicha jadval (Telegram): toifalar KPI, beradi / oladi / kechikkan (aybdorlar)."""
        s = self.koch_summary(sana or self._today())
        h = head or "🔁 <b>Ko'chirish · bugun</b>"
        if not s["has"]:
            return h + "\n\nTopshiriq hali shakllanmagan."
        t = h
        for k in KOCH_CATS:
            v = s["kpi"][k["key"]]
            t += "\n%s %s: %d ta · berildi %d · olindi %d · kechikkan %d · <b>%d%%</b>" % (
                k["icon"], k["title"], v["jami"], v["berildi"], v["olindi"], v["kechikkan"], v["pct"])
        t += "\n" + self.deadline_line("shoshilinch")
        t += "\n\n<b>Filiallar</b> (beradi qoldi / oladi qoldi / kechikkan — aybdor):\n" + "\n".join(
            "%s %s · %d / %d / <b>%d</b>" % ({"red": "🔴", "yellow": "🟡", "green": "🟢"}[r["state"]], esc(self.short(r["filial"], 24)),
                                            r["give_left"], r["take_left"], r["late"]) for r in s["filials"])
        return t

    def koch_day_report(self, sana):
        s = self.koch_summary(sana)
        rows = [t for t in self.koch_tasks(sana) if t["status"] != "bekor"]
        infos = [(t, self.task_info(t)) for t in rows]
        late = [(t, i) for t, i in infos if i["late"] or i["late_done"]]
        rad = [t for t in rows if t["status"] == "rad"]
        farq = [(t, i) for t, i in infos if i["farq"]]
        t = "📊 <b>Ko'chirish · kunlik hisobot</b> · %s\n" % datetime.date.fromisoformat(sana).strftime("%d.%m.%Y")
        for k in KOCH_CATS:
            v = s["kpi"][k["key"]]
            t += "\n%s %s: bajarilish <b>%d%%</b> (%d/%d)" % (k["icon"], k["title"], v["pct"], v["olindi"], v["jami"])
        if late:
            by = {}
            for tt, i in late:
                by[i["guilty_filial"]] = by.get(i["guilty_filial"], 0) + 1
            t += "\n\n⛔ <b>Kechikkan / kech bajarilgan</b>: %d ta\n" % len(late) + "\n".join(
                "• %s · %d ta" % (esc(f), n) for f, n in sorted(by.items(), key=lambda x: -x[1]))
        if rad:
            t += "\n\n❌ <b>Rad</b>: %d ta\n" % len(rad) + "\n".join(
                "• %s → %s · %s · <i>%s</i>" % (esc(self.short(x["src"], 18)), esc(self.short(x["dst"], 18)), esc(self.short(x["name"], 28)),
                                               esc(x["reason"] or x["note"] or "")) for x in rad[:15])
        if farq:
            t += "\n\n⚠ <b>Miqdor farqi</b>: %d ta\n" % len(farq) + "\n".join(
                "• %s · reja %s · berildi %s · olindi %s" % (esc(self.short(x["name"], 28)), qty(x["qty_plan"]), qty(x["qty_given"] or 0),
                                                           qty(x["qty_received"]) if x["qty_received"] is not None else "—")
                for x, _ in farq[:15])
        return t

    def koch_admin_action(self, cid, mid, a, p):
        """Admin Telegram: /kochirish, «🔁 Ko'chirish» (a:tr), Excel (a:trx), 🔔 eslatma (a:krm:<i> | a:krm:all)."""
        d = self._today()
        if a == "trx":
            rows = self.koch_task_list(d)
            if not rows:
                return self.panel_note(cid, mid, "➖ Bugungi topshiriq yo'q.")
            return self.panel_doc(cid, "Kochirish_nazorat_%s.xlsx" % d, build_koch_nazorat(self, rows, d),
                                  "🔁 <b>Ko'chirish nazorati</b> · %s" % d)
        late = [r["filial"] for r in self.koch_summary(d)["filials"] if r["late"]]
        note = None
        if a == "krm":
            targets = late if p and p[0] == "all" else [late[int(p[0])]] if p and p[0].isdigit() and int(p[0]) < len(late) else []
            n = sum(self.remind_filial(f, "🔔 <b>Admin eslatmasi: ko'chirish</b>") for f in targets)
            note = "🔔 Eslatma yuborildi: %d filial, %d xodim." % (len(targets), n)
            late = [r["filial"] for r in self.koch_summary(d)["filials"] if r["late"]]
        kb = [[("🔔 %s" % bl(self.short(f, 28)), "a:krm:%d" % i)] for i, f in enumerate(late[:8])]
        if late:
            kb.append([("🔔 Hamma kechikkanlarga", "a:krm:all")])
        kb.append([("🔄 Yangilash", "a:tr"), ("📎 Excel", "a:trx")])
        return self.a_screen(cid, mid, ("🔁 Ko'chirishlar",), self.koch_admin_text(d), kb, note=note)

    # ---- B7: dashboard ma'lumoti (server paneli) ----
    def koch_summary(self, sana=None):
        """KPI (har toifa), filiallar jadvali, yo'nalishlar matritsasi."""
        sana = sana or self._today()
        n = now()
        tasks = [t for t in self.koch_tasks(sana) if t["status"] != "bekor"]
        infos = {t["id"]: self.task_info(t, n) for t in tasks}
        kpi = {}
        for k in KOCH_CATS:
            rows = [t for t in tasks if t["cat"] == k["key"]]
            dl = self.koch_deadline(k["key"], sana)
            jami = len(rows)
            ol = sum(1 for t in rows if t["status"] == "olindi")
            kpi[k["key"]] = {"jami": jami, "berildi": sum(1 for t in rows if t["status"] in ("berildi", "olindi")), "olindi": ol,
                             "rad": sum(1 for t in rows if t["status"] == "rad"),
                             "kechikkan": sum(1 for t in rows if infos[t["id"]]["late"]),
                             "kech_bajarilgan": sum(1 for t in rows if infos[t["id"]]["late_done"]),
                             "farq": sum(1 for t in rows if infos[t["id"]]["farq"]),
                             "pct": int(round(100.0 * ol / jami)) if jami else 0, "title": k["title"], "icon": k["icon"],
                             "deadline": dl.strftime("%H:%M") if dl else "", "seconds_left": int((dl - n).total_seconds()) if dl else None}
        fil = {}
        blank = lambda: {"give": 0, "given": 0, "give_left": 0, "take": 0, "taken": 0, "take_left": 0, "late": 0, "rad": 0, "last": ""}
        for t in tasks:
            s, d = fil.setdefault(t["src"], blank()), fil.setdefault(t["dst"], blank())
            s["give"] += 1
            d["take"] += 1
            if t["status"] in ("berildi", "olindi"):
                s["given"] += 1
            if t["status"] == "yangi":
                s["give_left"] += 1
            if t["status"] == "rad":
                s["rad"] += 1
            if t["status"] == "olindi":
                d["taken"] += 1
            if t["status"] in ("yangi", "berildi"):
                d["take_left"] += 1
            for who, ts in ((s, t["given_at"]), (d, t["received_at"])):
                if ts and ts > who["last"]:
                    who["last"] = ts
            g = infos[t["id"]]["guilty_filial"]
            if g and infos[t["id"]]["late"]:
                fil.setdefault(g, blank())["late"] += 1
        frows = []
        for f, v in fil.items():
            involved = [t for t in tasks if f in (t["src"], t["dst"])]
            done = sum(1 for t in involved if t["status"] == "olindi")
            open_ = v["give_left"] + v["take_left"]
            state = "red" if v["late"] else ("green" if not open_ else "yellow")
            frows.append(dict(v, filial=f, pct=int(round(100.0 * done / len(involved))) if involved else 0, state=state,
                              last=v["last"][11:16] if v["last"] else ""))
        frows.sort(key=lambda r: (-r["late"], {"red": 0, "yellow": 1, "green": 2}[r["state"]], r["filial"]))
        names = sorted(fil)
        cells = {}
        for t in tasks:
            c = cells.setdefault(t["src"], {}).setdefault(t["dst"], [0, 0])
            c[1] += 1
            if t["status"] == "olindi":
                c[0] += 1
        return {"sana": sana, "has": bool(tasks) or self.koch_has_tasks(sana), "kpi": kpi, "cats": [k["key"] for k in KOCH_CATS],
                "filials": frows, "matrix": {"filials": names, "cells": cells}, "task_time": self.kset("koch_task_time"),
                "on": self.koch_on}

    def koch_task_dict(self, t, n=None, cls=None):
        """cls — koch_abc(sana) natijasi (ro'yxatda bir marta hisoblanadi), None — shu topshiriq kuni bo'yicha."""
        i = self.task_info(t, n)
        name_of = lambda cid: self._actor(cid) if cid else ""
        return {"id": t["id"], "sana": t["sana"], "src": t["src"], "dst": t["dst"], "name": t["name"], "producer": t["producer"],
                "firm": t["firm"], "cat": t["cat"], "turi": t["cat"], "sabab": t["sabab"], "days_idle": t["days_idle"] or 0,
                "qty_plan": t["qty_plan"], "qty_given": t["qty_given"], "qty_received": t["qty_received"], "prih": t["prih"],
                "summa": t["summa"] or 0, "srok": t["srok"] or "", "status": t["status"], "reason": t["reason"] or "",
                "note": t["note"] or "", "given_by": name_of(t["given_by"]), "given_at": (t["given_at"] or "")[11:16],
                "received_by": name_of(t["received_by"]), "received_at": (t["received_at"] or "")[11:16], "late": i["late"],
                "late_done": i["late_done"], "late_min": i["late_min"], "guilty": i["guilty_filial"] or "", "farq": i["farq"],
                "by": self.koch_by(t), "auto_by": t["auto_by"] or "", "auto_note": t["auto_note"] or "",
                "cls": (self.koch_abc(t["sana"]) if cls is None else cls).get(t["id"], "C") if t["status"] != "bekor" else ""}

    def koch_task_list(self, sana=None, filial=None, cat=None, status=None, q=None, limit=5000, turi=None):
        n = now()
        cat = cat or {"zudlik": "shoshilinch"}.get(turi, turi)
        rows = self.koch_tasks(sana or self._today(), cat=cat or None)
        cls = self.koch_abc(sana or self._today())
        fl, qq = norm(filial), canon(q) if q else ""
        out = []
        for t in rows:
            if fl and fl not in (t["src"], t["dst"]):
                continue
            if qq and qq not in canon(t["name"]):
                continue
            d = self.koch_task_dict(t, n, cls)
            if status == "kechikkan" and not (d["late"] or d["late_done"]):
                continue
            if status == "farq" and not d["farq"]:
                continue
            if status in self.ST_NAME and t["status"] != status:
                continue
            out.append(d)
            if len(out) >= limit:
                break
        return out

    def koch_task_detail(self, tid):
        t = self.db.execute("SELECT * FROM koch_tasks WHERE id=?", (tid,)).fetchone()
        if not t:
            return None
        ev = [{"ts": (e["ts"] or "").replace("T", " ")[:16], "action": e["action"], "user": e["user_name"] or "", "qty": e["qty"],
               "note": e["note"] or ""} for e in self.db.execute("SELECT * FROM koch_events WHERE task_id=? ORDER BY id", (tid,))]
        return dict(self.koch_task_dict(t), events=ev)

    def koch_set_status(self, tid, status, note, by=(None, "admin")):
        """Admin: holatni qo'lda o'zgartirish yoki bekor qilish (izoh majburiy, koch_events ga yoziladi)."""
        if status not in self.ST_NAME:
            return "Noma'lum holat"
        if not (note or "").strip():
            return "Izoh yozing"
        t = self.db.execute("SELECT * FROM koch_tasks WHERE id=?", (tid,)).fetchone()
        if not t:
            return "Topshiriq topilmadi"
        ts = now().isoformat(timespec="seconds")
        sets = {"status": status, "note": note.strip()[:200]}
        if status in ("berildi", "olindi") and t["qty_given"] is None:
            sets.update(qty_given=t["qty_plan"], given_at=t["given_at"] or ts)
        if status == "olindi" and t["qty_received"] is None:
            sets.update(qty_received=t["qty_given"] if t["qty_given"] is not None else t["qty_plan"], received_at=ts)
        if status == "yangi":
            sets.update(qty_given=None, given_at=None, given_by=None, qty_received=None, received_at=None, received_by=None)
        if status == "berildi":
            sets.update(qty_received=None, received_at=None, received_by=None)
        sets["auto_by"] = None  # admin qo'ydi: manba endi tizim emas (avto-aniqlash bu topshiriqqa boshqa tegmaydi)
        with self.dblock:
            self.db.execute("UPDATE koch_tasks SET %s WHERE id=?" % ", ".join("%s=?" % k for k in sets), list(sets.values()) + [tid])
            self._event(tid, "admin:" + status, by[0], by[1], None, note.strip()[:200], ts)
            self.db.commit()
        self._ret_cache = None
        logging.info("Ko'chirish #%d: admin holatni %s ga o'zgartirdi (%s)", tid, status, note)
        return None

    # ================= ko'chirishni qoldiqdan avtomatik aniqlash (dizayn/AVTO-ANIQLASH.md) =================
    # Topshiriq yaratilganda src_stock0 / dst_stock0 (Sale varag'idagi qoldiq) saqlanadi. Har muvaffaqiyatli yangilashdan keyin
    # (qo'lda, ertalabki, har soatlik) ochiq topshiriqlar hozirgi qoldiq bilan solishtiriladi (avto_aniqlash.attribute).
    # Tizim faqat «olindi» qo'yadi (yangi / berildi → olindi); admin qo'lda o'zgartirgan topshiriqqa tizim tegmaydi.
    AUTO_IDLE_WARN = 2  # «Berdim» dan keyin shuncha yangilashda qoldiq o'zgarmasa — adminga «⚠ Tekshiring»

    def sale_index(self):
        """{(nom, ishlab chiqaruvchi): Sale qatori} — boshqa joylardagi kabi birinchi qator."""
        idx = {}
        for r in self.data.sale:
            idx.setdefault((norm(r["name"]), norm(r["producer"])), r)
        return idx

    def stock_col(self, filial):
        """Sale varag'idagi qoldiq ustuni (v indeksi) shu filial uchun: 6-qator sarlavhalari («Sotuv va qoldiq» dagi kabi —
        Data.sale_heads, fil_key bo'yicha), sklad — v[6]. Topilmasa None (bu filial bo'yicha aniqlanmaydi)."""
        k = fil_key(filial)
        if not k:
            return None
        for i, h in enumerate(self.data.sale_heads or []):
            if h and fil_key(h) == k:
                return (0, 2, 4, 6)[i]
        return 6 if self.is_sklad(filial) else None

    def stock_now(self, name, producer, filial, idx=None):
        """Hozirgi qoldiq (Sale varag'i) yoki None (dori yoki filial ustuni topilmadi)."""
        col = self.stock_col(filial)
        r = (idx if idx is not None else self.sale_index()).get((norm(name), norm(producer)))
        return None if col is None or r is None else r["v"][col]

    def koch_dst_now(self, tid, idx=None):
        """«Berdim» bosilganda: oluvchi filialdagi hozirgi qoldiq (yuklangan Sale ma'lumoti) yoki None — dst_stock1 uchun."""
        t = self.db.execute("SELECT name, producer, dst FROM koch_tasks WHERE id=?", (tid,)).fetchone()
        if not t or not self.data.sale:
            return None
        return self.stock_now(t["name"], t["producer"], t["dst"], idx)

    def koch_moved(self, t, idx=None):
        """Surat olinganidan beri beruvchida qoldiq kamaydi VA oluvchida ko'paydi (ko'chirish bo'lgan bo'lishi mumkin).
        Qayta shakllantirish bunday «yangi» ni bekor qilmaydi va rejasini kamaytirmaydi — avto-aniqlashga qoldiradi."""
        if t["src_stock0"] is None or t["dst_stock0"] is None:
            return False
        sn = self.stock_now(t["name"], t["producer"], t["src"], idx)
        dn = self.stock_now(t["name"], t["producer"], t["dst"], idx)
        return sn is not None and dn is not None and sn < t["src_stock0"] - 1e-9 and dn > t["dst_stock0"] + 1e-9

    def koch_by(self, t):
        """Holatni kim qo'ydi: «tizim» (avto-aniqlash), xodim ismi yoki "" (admin / noma'lum)."""
        st = t["status"]
        uid = t["received_by"] if st == "olindi" else t["given_by"] if st in ("berildi", "rad") else None
        if uid:
            return self._actor(uid)
        return "tizim" if st in ("berildi", "olindi") and t["auto_by"] == "tizim" else ""

    def koch_by_text(self, t):
        """« · 🤖 tizim» | « · 👤 Aziz» | ""."""
        by = self.koch_by(t)
        return "" if not by else " · 🤖 tizim" if by == "tizim" else " · 👤 %s" % esc(by)

    @staticmethod
    def _q2(v):
        v = round(float(v), 2)
        return float(int(v)) if abs(v - int(v)) < 1e-9 else v

    KOCH_AUTO_DAYS = 3  # «berildi» topshiriq «Berdim» (given_at) dan shuncha kun ichida avto-aniqlanadi (keyin — yetkazib beruvchi kirimi)
    AUTO_DIGEST_MAX = 15  # adminga bitta yig'ma «Tekshiring» xabarida ko'rsatiladigan qatorlar

    def koch_auto_detect(self):
        """Ochiq topshiriqlar (bugungi «yangi», «Berdim» dan KOCH_AUTO_DAYS kun ichidagi «berildi») bo'yicha qoldiq o'zgarishini
        tekshiradi (dizayn/AVTO-ANIQLASH.md). Faqat «ikkala tomonda ham o'zgardi» holatni o'zgartiradi (→ «olindi»); yolg'iz
        beruvchidagi kamayish (sotuv bo'lishi mumkin) — faqat ishora. «berildi» da oluvchidagi ko'payish «Berdim» bosilgandagi
        qoldiqdan (dst_stock1, bo'lmasa dst_stock0) o'lchanadi. Bitta oluvchiga bir nechta topshiriq — oluvchidagi ko'payish reja
        bo'yicha bo'linadi (avto_aniqlash.share); bitta beruvchidan bir nechta oluvchiga — har oluvchi o'zi bo'yicha
        (avto_aniqlash.attribute). Ogohlantirishlar har adminga bitta yig'ma xabarda. Natija: {"olindi", "ishora",
        "tekshiring", "berdim_warn"} — sonlar."""
        with self._detect_lock:  # ikki oqim bir vaqtda: ikkinchisi birinchisi yozgan belgilarni ko'radi
            return self._koch_auto_detect()

    def _koch_auto_detect(self):
        res = dict.fromkeys(("olindi", "ishora", "tekshiring", "berdim_warn"), 0)
        if not self.koch_on or not self.data.sale:
            return res
        idx = self.sale_index()
        n0 = now()
        today = n0.date()
        since = (today - datetime.timedelta(days=self.KOCH_RECV_DAYS)).isoformat()
        given_since = (n0 - datetime.timedelta(days=self.KOCH_AUTO_DAYS)).isoformat(timespec="seconds")
        w = self.data.when()
        stamp = w.isoformat(timespec="seconds") if w else ""
        tasks = self.db.execute("SELECT * FROM koch_tasks WHERE (status='yangi' AND sana=?) OR (status='berildi' AND sana>=?) "
                                "ORDER BY id", (today.isoformat(), since)).fetchall()
        manual = {r[0] for r in self.db.execute("SELECT DISTINCT task_id FROM koch_events WHERE action LIKE 'admin:%'")}
        rows = []
        for t in tasks:
            if t["id"] in manual or t["src_stock0"] is None or t["dst_stock0"] is None:
                continue
            if t["status"] == "berildi" and (t["given_at"] or "") < given_since:
                continue  # «Berdim» dan 3 kundan ko'p o'tdi: oluvchidagi ko'payish endi yetkazib beruvchidan bo'lishi mumkin
            sn = self.stock_now(t["name"], t["producer"], t["src"], idx)
            dn = self.stock_now(t["name"], t["producer"], t["dst"], idx)
            if sn is None or dn is None:
                continue
            berildi = t["status"] == "berildi"
            plan = t["qty_given"] if berildi and t["qty_given"] else t["qty_plan"]
            d0 = t["dst_stock1"] if berildi and t["dst_stock1"] is not None else t["dst_stock0"]
            rows.append({"t": t, "plan": plan or 0, "down": t["src_stock0"] - sn, "up": dn - d0})
        drug = lambda x: (norm(x["t"]["name"]), norm(x["t"]["producer"]))
        groups = {}  # bitta oluvchiga bir nechta topshiriq: oluvchidagi ko'payish reja bo'yicha bo'linadi
        for x in rows:
            groups.setdefault(drug(x) + (x["t"]["dst"],), []).append(x)
        for g in groups.values():
            if len(g) > 1:
                sh = AA.share(max(0.0, max(x["up"] for x in g)), {i: x["plan"] for i, x in enumerate(g)})
                for i, x in enumerate(g):
                    x["up"] = sh[i]
        groups = {}  # bitta beruvchidan: har oluvchi o'zi bo'yicha (avto_aniqlash.attribute)
        for x in rows:
            groups.setdefault(drug(x) + (x["t"]["src"],), []).append(x)
        for g in groups.values():
            for x, (st, q, d) in zip(g, AA.attribute(g)):
                x["st"], x["q"], x["down"] = st, q, d
        ts = n0.isoformat(timespec="seconds")
        warns = []
        with self.dblock:
            for x in rows:
                t = x["t"]
                st, q = x["st"], self._q2(x["q"])
                note = "%s −%s · %s +%s" % (fil_key(t["src"]), qty(self._q2(max(0, x["down"]))), fil_key(t["dst"]),
                                            qty(self._q2(max(0, x["up"]))))
                flags = t["auto_warn"] or ""
                if st == "olindi":
                    given = t["qty_given"] if t["status"] == "berildi" and t["qty_given"] is not None else q
                    cur = self.db.execute(
                        "UPDATE koch_tasks SET status='olindi', qty_given=?, qty_received=?, given_at=COALESCE(given_at, ?), "
                        "received_at=COALESCE(received_at, ?), auto_by='tizim', auto_note=? WHERE id=? AND status=?",
                        (given, min(q, given), ts, ts, note, t["id"], t["status"]))
                    if cur.rowcount == 1:
                        self._event(t["id"], "tizim:olindi", None, "tizim", min(q, given), note, ts)
                        res["olindi"] += 1
                    continue
                if st in AA.HINTS:  # faqat beruvchida kamaydi: sotuv bo'lishi mumkin — holat o'zgarmaydi
                    res["ishora"] += 1
                    continue
                if st == "tekshiring" and "t" not in flags:
                    flags += "t"
                    self.db.execute("UPDATE koch_tasks SET auto_warn=?, auto_note=? WHERE id=?", (flags, note, t["id"]))
                    self._event(t["id"], "tizim:tekshiring", None, "tizim", q, note, ts)
                    down = self._q2(max(0, x["down"]))
                    warns.append("💊 %s · %s → %s · reja %s qadoq\nOluvchida +%s, beruvchida %s — boshqa joydan kelgan bo'lishi mumkin." % (
                        esc(t["name"]), esc(t["src"]), esc(t["dst"]), qty(t["qty_plan"]), qty(q),
                        ("−%s" % qty(down)) if down > 0 else "qoldiq o'zgarmadi"))
                    res["tekshiring"] += 1
                # «Berdim» bosilgan (xodim), lekin beruvchida ham, oluvchida ham qoldiq o'zgarmagan: yangilashlar sanaladi
                if (st is None and t["status"] == "berildi" and t["given_by"] and t["auto_by"] is None and stamp
                        and stamp > (t["given_at"] or "") and stamp != (t["auto_seen"] or "")):
                    n = (t["auto_idle"] or 0) + 1
                    if n >= self.AUTO_IDLE_WARN and "b" not in flags:
                        flags += "b"
                        warns.append("💊 %s · %s → %s · %s qadoq\n%s «Berdim» dedi (%s), qoldiq o'zgarmadi — %d ta yangilashdan keyin "
                                     "ham beruvchida kamaymadi." % (esc(t["name"]), esc(t["src"]), esc(t["dst"]),
                                                                    qty(t["qty_given"] or 0), esc(t["src"]),
                                                                    (t["given_at"] or "")[11:16], n))
                        self._event(t["id"], "tizim:qoldiq_ozgarmadi", None, "tizim", None, note, ts)
                        res["berdim_warn"] += 1
                    self.db.execute("UPDATE koch_tasks SET auto_idle=?, auto_seen=?, auto_warn=? WHERE id=?",
                                    (n, stamp, flags, t["id"]))
            self.db.commit()
        if res["olindi"]:
            self._ret_cache = None
        if warns:  # bitta yig'ma xabar (har adminga), har topshiriq uchun bir marta
            self._admins_send(self.koch_warn_digest(warns, n0))
        if any(res.values()):
            logging.info("Ko'chirish avto-aniqlash: %s", res)
        return res

    def koch_warn_digest(self, items, when=None):
        """«⚠ Tekshiring» — shu yangilashdagi hamma ogohlantirish bitta xabarda (ko'pi bilan AUTO_DIGEST_MAX qator + «yana N ta»)."""
        head = "⚠ <b>Tekshiring</b> · ko'chirish: <b>%d</b> ta%s" % (
            len(items), (" · %s" % when.strftime("%H:%M")) if when else "")
        body = ["%d. %s" % (i, s) for i, s in enumerate(items[:self.AUTO_DIGEST_MAX], 1)]
        if len(items) > self.AUTO_DIGEST_MAX:
            body.append("… yana <b>%d</b> ta — to'liq ro'yxat: «🔁 Ko'chirishlar» (Excel)." % (len(items) - self.AUTO_DIGEST_MAX))
        return head + "\n\n" + "\n\n".join(body)

    def koch_history(self, days=7):
        """Har filial uchun kunlik «o'z vaqtida bajarilish %», o'rtacha kechikish, kechikkan kunlar, rad soni, reyting."""
        end = now().date()
        dates = [(end - datetime.timedelta(days=i)).isoformat() for i in range(days - 1, -1, -1)]
        series, agg = {}, {}
        for d in dates:
            tasks = [t for t in self.koch_tasks(d) if t["status"] != "bekor"]
            if not tasks:
                continue
            per = {}
            for t in tasks:
                i = self.task_info(t)
                for f in {t["src"], t["dst"]}:
                    p = per.setdefault(f, {"n": 0, "bad": 0})
                    p["n"] += 1
                g = i["guilty_filial"]
                if g and (i["late"] or i["late_done"]):
                    per.setdefault(g, {"n": 0, "bad": 0})["bad"] += 1
                    a = agg.setdefault(g, {"late_min": [], "late_days": set(), "rad": 0})
                    a["late_min"].append(i["late_min"])
                    a["late_days"].add(d)
                if t["status"] == "rad":
                    agg.setdefault(t["src"], {"late_min": [], "late_days": set(), "rad": 0})["rad"] += 1
            for f, p in per.items():
                series.setdefault(f, {})[d] = int(round(100.0 * max(0, p["n"] - p["bad"]) / p["n"])) if p["n"] else 100
        rows = []
        for f in sorted(series):
            a = agg.get(f, {"late_min": [], "late_days": set(), "rad": 0})
            vals = list(series[f].values())
            rows.append({"filial": f, "avg_pct": int(round(sum(vals) / len(vals))) if vals else 0,
                         "avg_late_min": int(round(sum(a["late_min"]) / len(a["late_min"]))) if a["late_min"] else 0,
                         "late_days": len(a["late_days"]), "rad": a["rad"]})
        rank = sorted(rows, key=lambda r: (-r["late_days"], -r["avg_late_min"]))
        return {"days": days, "dates": dates, "series": {f: [series[f].get(d) for d in dates] for f in series}, "rows": rows,
                "rank": [r for r in rank if r["late_days"]]}

    # ================= sklad taqsimoti (server paneli va admin) =================
    def sklad_summary(self):
        """Sklad → dorixonalar: dori × dorixona jadvali, dorixonalar xulosasi, muddat sababli berilmaydiganlar.
        Bugungi topshiriq bo'lsa — undan (holati bilan), aks holda Excel rejasidan."""
        d = self.data
        d.reload()
        sk, today = self.sklad, self._today()
        plan = [r for r in d.koch if norm(r["src"]) == sk]
        tasks = [t for t in self.koch_tasks(today, src=sk) if t["status"] != "bekor"] if self.koch_on else []
        if tasks:
            rows = [dict(self._task_row(t), srok=t["srok"]) for t in tasks]
        else:
            rows = [dict(r, status="reja", srok=r["date"].isoformat() if r["date"] else None) for r in plan
                    if not self.koch_exclude_reason(r)]
        excluded = []
        for r in plan:
            why = self.koch_exclude_reason(r)
            if why and why.startswith("Muddati"):
                excluded.append({"dst": norm(r["dst"]), "name": r["name"], "producer": r["producer"], "qty": r["qty"], "why": why,
                                 "srok": r["date"].isoformat() if r["date"] else None})
        items, dorixona = {}, {}
        for r in rows:
            dst = norm(r["dst"])
            i = items.setdefault((norm(r["name"]), norm(r["producer"])), {
                "name": r["name"], "producer": r["producer"], "stock": r.get("src_stock") or 0, "srok": r["srok"], "to": {}, "total": 0.0,
                "rank": 9})
            i["to"][dst] = {"qty": r["qty"], "status": r["status"], "cat": r["cat"], "summa": r["summa"]}
            i["total"] += r["qty"]
            i["rank"] = min(i["rank"], KOCH_CAT_ORDER.get(r["cat"], 9))
            g = dorixona.setdefault(dst, {"filial": dst, "n": 0, "qty": 0.0, "summa": 0.0, "given": 0, "taken": 0, "rad": 0, "left": 0,
                                          "cats": {k["key"]: 0 for k in KOCH_CATS}})
            g["n"] += 1
            g["qty"] += r["qty"]
            g["summa"] += r["summa"]
            g["cats"][r["cat"]] = g["cats"].get(r["cat"], 0) + 1
            g["given"] += r["status"] in ("berildi", "olindi")
            g["taken"] += r["status"] == "olindi"
            g["rad"] += r["status"] == "rad"
            g["left"] += r["status"] in ("yangi", "reja")
        for i in items.values():
            i["left"] = max(0.0, i["stock"] - i["total"])
        out = sorted(items.values(), key=lambda i: (i["rank"], -i["total"], norm(i["name"])))
        fil = [dorixona[k] for k in sorted(dorixona)]
        n = sum(g["n"] for g in fil)
        return {"sklad": sk, "has": bool(tasks), "on": self.koch_on, "dst": [g["filial"] for g in fil], "dorixona": fil, "items": out,
                "excluded": excluded, "when": dt_text(d.when()),
                "kpi": {"items": len(out), "n": n, "qty": sum(g["qty"] for g in fil), "summa": sum(g["summa"] for g in fil),
                        "given": sum(g["given"] for g in fil), "taken": sum(g["taken"] for g in fil), "rad": sum(g["rad"] for g in fil),
                        "left": sum(g["left"] for g in fil), "pct": int(round(100.0 * sum(g["taken"] for g in fil) / n)) if n else 0}}

    def sklad_admin_text(self, s=None, prices=True):
        s = s or self.sklad_summary()
        t = "🏭 <b>Sklad taqsimoti</b>"
        if not s["dorixona"] and not s["excluded"]:
            return t + "\n\nTaqsimot yo'q: Excelda «Skladdan» varag'i bo'sh yoki hali yangilanmagan." + self.foot()
        k = s["kpi"]
        so_m = lambda v: " · <b>%s</b> so'm" % money(v) if prices else ""
        t += "\n<b>%d</b> xil dori · <b>%s</b> qadoq%s" % (k["items"], qty(k["qty"]), so_m(k["summa"]))
        if s["has"]:
            t += "\nBerildi: <b>%d</b> · olindi: <b>%d</b> · qoldi: <b>%d</b> (%d%%)" % (k["given"], k["taken"], k["left"], k["pct"])
        else:
            t += "\n<i>Reja: bugungi topshiriq hali shakllanmagan</i>"
        for g in s["dorixona"]:
            t += "\n\n→ <b>%s</b>: %d ta · %s qadoq%s\n   %s" % (
                esc(g["filial"]), g["n"], qty(g["qty"]), so_m(g["summa"]),
                " · ".join("%s %d" % (c["icon"], g["cats"].get(c["key"], 0)) for c in KOCH_CATS))
            if s["has"]:
                t += " · qoldi %d" % g["left"]
        if s["excluded"]:
            t += "\n\n⛔ <b>Muddati sababli berilmaydi</b>: %d qator (Excelda)" % len(s["excluded"])
        return t + self.foot()

    # ================= sklad: FOM bron (перемещение) — server panelidan =================
    # Admin server panelida bugungi sklad topshirig'idan (standart: barcha toifalar) har dorixona uchun bron yaratadi.
    # Hozircha «fayl» rejimi: bron Excel fayl bo'lib yuklab olinadi (FOM'ga kiritish uchun); FOM bazasiga yozilmaydi.
    # Bir topshiriq qatori faqat bitta bronga kiradi. Sklad xodimining Telegram tugmalari o'zgarmaydi.
    @property
    def fom_cats(self):
        allc = ",".join(k["key"] for k in KOCH_CATS)
        raw = self.cfg.get("fom", "cats", fallback=allc) if self.cfg.has_section("fom") else allc
        cats = [c.strip().lower() for c in raw.split(",") if c.strip().lower() in KOCH_CAT]
        return [k["key"] for k in KOCH_CATS if k["key"] in cats] or [k["key"] for k in KOCH_CATS]

    def fom_groups(self):
        """[(raqam, dorixona, qatorlar)] — bugungi sklad topshirig'idan hali bron qilinmagan «yangi»/«berildi» qatorlar."""
        sk = self.sklad
        if not self.koch_on or not self.koch_has_tasks():
            return []
        done = {r[0] for r in self.db.execute(
            "SELECT i.task_id FROM fom_boron_items i JOIN fom_boron b ON b.id=i.boron_id WHERE b.sana=? AND b.status NOT IN ('xato','bekor')",
            (self._today(),))}
        rows = []
        for cat in self.fom_cats:
            for t in self.koch_tasks(self._today(), src=sk, cat=cat, statuses=("yangi", "berildi")):
                if t["id"] in done:
                    continue
                r = self._task_row(t)
                why = self.koch_exclude_reason(r) or ""
                if why == "Muddati o'tgan" or why.endswith("filialdan kelgan"):
                    continue
                rows.append(r)
        out = []
        for i, d in enumerate(sorted({t["dst"] for t in self.koch_tasks(self._today(), src=sk)})):
            part = [r for r in rows if norm(r["dst"]) == d]
            if part:
                out.append((i, d, part))
        return out

    def fom_sig(self, groups):
        return self.koch_sig((r["id"], r["qty"]) for g in groups for r in g[2])

    def fom_pick(self, which, sig):
        """which: dorixona raqami yoki 'a' (hammasi). Ro'yxat admin ko'rgandan keyin o'zgargan bo'lsa None."""
        groups = self.fom_groups()
        if str(which) != "a":
            groups = [g for g in groups if str(g[0]) == str(which)]
        if not groups or self.fom_sig(groups) != sig:
            return None
        return groups

    def fom_today(self):
        return self.db.execute("SELECT * FROM fom_boron WHERE sana=? AND src=? AND status NOT IN ('xato','bekor') ORDER BY id",
                               (self._today(), self.sklad)).fetchall()

    def _fom_orphan_alert(self, plans, what):
        """FOM'ga yozilib bo'lgan (commit), lekin bot bazasiga yozilmagan: qayta bosilsa FOM'da dublikat bo'lardi.
        Adminlarga qoralama raqamlari bilan ogohlantirish."""
        docs = [str(p["invoice"]) for p in (plans or []) if p.get("invoice")]
        if not docs:
            return
        logging.error("FOM %s: FOM'ga yozildi (qoralama %s), lekin bot bazasiga yozilmadi", what, ", ".join(docs))
        text = ("⚠️ <b>FOM %s: diqqat</b>\nFOM'da qoralama(lar) yaratildi: <b>%s</b>, lekin bot bazasiga yozilmadi.\n"
                "Qayta bosishdan oldin FOM'da shu qoralamalarni tekshiring (dublikat bo'lmasin).") % (esc(what), esc(", ".join(docs)))
        for a in self.admins:
            try:
                self.tg.send(a, text)
            except Exception:  # noqa
                pass

    def fom_create(self, which, sig, by="admin (panel)"):
        """Bron(lar)ni yaratadi: hammasi yoki hech biri. Ro'yxat o'zgargan / allaqachon yaratilgan bo'lsa None."""
        ts = now().isoformat(timespec="seconds")
        made, fom = [], self.fom_sql
        sqlmode = fom.enabled
        plans = None
        with self.dblock:
            groups = self.fom_pick(which, sig)
            if not groups:
                return None
            try:
                bids = {}
                for _, d, rows in groups:
                    cur = self.db.execute(
                        "INSERT INTO fom_boron(sana, src, dst, cats, n, qty, created_by, created_at, mode, status) "
                        "VALUES(?,?,?,?,?,?,?,?,?,?)",
                        (self._today(), self.sklad, d, ",".join(sorted({r["cat"] for r in rows})), len(rows),
                         sum(r["qty"] for r in rows), by, ts, "sql" if sqlmode else "fayl", "fayl"))
                    bid = cur.lastrowid
                    self.db.executemany(
                        "INSERT INTO fom_boron_items(boron_id, task_id, name, producer, cat, qty, srok) VALUES(?,?,?,?,?,?,?)",
                        [(bid, r["id"], r["name"], r["producer"], r["cat"], r["qty"], r["date"].isoformat() if r["date"] else None)
                         for r in rows])
                    bids[d] = bid
                    made.append(bid)
                if sqlmode:
                    # avval FOM (bitta tranzaksiya); u rad etsa — bot bazasiga ham hech narsa yozilmaydi
                    plans = fom.write(groups, {d: "BOT bron #%d" % b for d, b in bids.items()})
                    for p in plans:
                        bid = bids[p["dst"]]
                        self.db.executemany(
                            "INSERT INTO fom_boron_lines(boron_id, task_id, incomeln, kol, price, srok) VALUES(?,?,?,?,?,?)",
                            [(bid, ln["task_id"], ln["incomeln"], ln["kol"], ln["price"], ln["exp"]) for ln in p["lines"]])
                        self.db.execute("UPDATE fom_boron SET status=?, doc_no=?, short=? WHERE id=?",
                                        ("qoralama" if p["invoice"] else "fayl", str(p["invoice"] or ""),
                                         json.dumps(p["short"], ensure_ascii=False), bid))
                self.db.commit()
            except Exception:
                self.db.rollback()
                self._fom_orphan_alert(plans, "bron")
                raise
        made = [self.db.execute("SELECT * FROM fom_boron WHERE id=?", (b,)).fetchone() for b in made]
        text = "📤 <b>FOM bron yaratildi</b> · %s\n%s\n%s" % (esc(by), esc(self.sklad), "\n".join(
            "#%d → %s · %d ta dori · %s qadoq%s" % (b["id"], esc(b["dst"]), b["n"], qty(b["qty"]), self._fom_tail(b)) for b in made))
        for a in self.admins:
            try:
                self.tg.send(a, text)
            except Exception:  # noqa
                logging.exception("Adminga xabar yuborib bo'lmadi: %s", a)
        return made

    @property
    def fom_sql(self):
        """FOM SQL yozuvchisi (config.ini [fom] mode = sql bo'lsa yoqiq)."""
        import fom_sql
        if getattr(self, "_fom_sql", None) is None:
            self._fom_sql = fom_sql.FomSql(self.cfg)
        return self._fom_sql

    def _fom_tail(self, b):
        if (b["mode"] or "") != "sql":
            return ""
        short = json.loads(b["short"] or "[]")
        t = " · FOM qoralama #%s" % b["doc_no"] if b["doc_no"] else " · FOM'ga yozilmadi (qoldiq yo'q)"
        return t + (" · yetmadi: %d ta" % len(short) if short else "")

    def fom_check(self, which, sig):
        """Yozmasdan oldin: FOM'da nima yoziladi, nima yetmaydi. {"ok", "plans"|"error"}."""
        groups = self.fom_pick(which, sig)
        if not groups:
            return {"ok": False, "error": "Ro'yxat o'zgargan yoki bron allaqachon yaratilgan"}
        if not self.fom_sql.enabled:
            return {"ok": True, "mode": "fayl", "plans": []}
        import fom_sql
        try:
            plans = self.fom_sql.preview(groups)
        except fom_sql.FomError as e:
            return {"ok": False, "error": str(e)}
        return {"ok": True, "mode": "sql", "db": self.fom_sql.database, "plans": [
            {"dst": p["dst"], "lines": len(p["lines"]), "kol": sum(ln["kol"] for ln in p["lines"]),
             "short": [{"name": x["name"], "need": x["need"], "got": x["got"], "why": x["why"]} for x in p["short"]]}
            for p in plans]}

    def fom_cancel(self, bid, by="admin (panel)"):
        """Bugungi bronni bekor qiladi: dorilari yana ro'yxatga qaytadi. FOM'ga yozilgan bron bekor qilinmaydi."""
        with self.dblock:
            b = self.db.execute("SELECT * FROM fom_boron WHERE id=?", (bid,)).fetchone()
            if not b or b["sana"] != self._today() or b["status"] in ("xato", "bekor"):
                return "Bron topilmadi yoki allaqachon bekor qilingan"
            if b["status"] != "fayl":
                return "Bu bron FOM'ga yozilgan: uni FOM dasturida o'chiring"
            self.db.execute("UPDATE fom_boron SET status='bekor', error=? WHERE id=?",
                            ("bekor: %s %s" % (by, now().strftime("%H:%M")), bid))
            self.db.commit()
        for a in self.admins:
            try:
                self.tg.send(a, "✖ <b>FOM bron bekor qilindi</b> · %s\n#%d → %s · %d ta dori" % (esc(by), b["id"], esc(b["dst"]), b["n"]))
            except Exception:  # noqa
                logging.exception("Adminga xabar yuborib bo'lmadi: %s", a)
        return None

    def fom_items(self, bid):
        return self.db.execute("SELECT * FROM fom_boron_items WHERE boron_id=? ORDER BY rowid", (bid,)).fetchall()

    def fom_file(self, bid):
        """(xlsx, fayl nomi) yoki None."""
        b = self.db.execute("SELECT * FROM fom_boron WHERE id=?", (bid,)).fetchone()
        if not b:
            return None
        fom = None
        if (b["mode"] or "") == "sql":
            fom = {}
            for ln in self.db.execute("SELECT task_id, SUM(kol) k FROM fom_boron_lines WHERE boron_id=? GROUP BY task_id", (bid,)):
                fom[ln["task_id"]] = (ln["k"], "")
            for x in json.loads(b["short"] or "[]"):
                fom[x["task_id"]] = (x["got"], x["why"])
        return build_fom_boron(b, self.fom_items(bid), fom), "Bron_%d_%s_%s.xlsx" % (
            b["id"], b["sana"], re.sub(r"[^\w-]+", "_", b["dst"]).strip("_"))

    def fom_summary(self):
        """Server paneli uchun: bron qilinmagan dorixonalar va bugun yaratilgan bronlar."""
        groups = self.fom_groups() if self.koch_on else []
        return {
            "has": bool(self.koch_on and self.koch_has_tasks()),
            "cats": [KOCH_CAT[k]["icon"] + " " + KOCH_CAT[k]["title"] for k in self.fom_cats],
            "cat_keys": self.fom_cats,
            "icons": {k: KOCH_CAT[k]["icon"] for k in self.fom_cats},
            "groups": [{"i": i, "dst": d, "n": len(rows), "qty": sum(r["qty"] for r in rows), "sig": self.fom_sig([(i, d, rows)]),
                        "by_cat": {k: sum(1 for r in rows if r["cat"] == k) for k in self.fom_cats},
                        "items": [{"name": r["name"], "producer": r["producer"], "qty": r["qty"], "cat": r["cat"],
                                   "srok": r["date"].isoformat() if r["date"] else ""} for r in rows]}
                       for i, d, rows in groups],
            "all_sig": self.fom_sig(groups) if groups else "",
            "made": [{"id": b["id"], "dst": b["dst"], "n": b["n"], "qty": b["qty"], "time": (b["created_at"] or "")[11:16],
                      "by": b["created_by"] or "", "status": b["status"], "doc": b["doc_no"] or "",
                      "short": len(json.loads(b["short"] or "[]")), "kirim": (b["cats"] or "") == "kirim"}
                     for b in self.fom_today()],
            "mode": "sql" if self.fom_sql.enabled else "fayl",
            "db": self.fom_sql.database if self.fom_sql.enabled else "",
        }

    # ================= sklad: bugun kelgan yuk (kirim) → filiallarga sotuv ulushi bo'yicha → FOM bron =================
    # Kirim — FOM'dan (shu kuni o'tkazilgan partiyalar, hali tarqatilmagan qismi). Sotuv — Excel «Sale» varag'idan
    # (filiallar: ЧИНОЗ v[1], УРГУТ v[3], ЭРКИН v[5]). Taqsimot = sotuv ulushi (eng katta qoldiq usuli, butun son).
    KIRIM_DST = (("ДОРИХОНА ЧИНОЗ", 0), ("ДОРИХОНА УРГУТ", 2), ("ДОРИХОНА ЭРКИН", 4))   # (filial, Sale v[] qoldiq indeksi)

    @staticmethod
    def _share(total, weights):
        """total (butun) ni weights ulushida butun sonlarga bo'ladi (eng katta qoldiq usuli)."""
        s = sum(weights)
        if total <= 0 or s <= 0:
            return [0] * len(weights)
        raw = [total * w / s for w in weights]
        out = [int(x) for x in raw]
        for i in sorted(range(len(raw)), key=lambda i: (-round(raw[i] - out[i], 9), -weights[i], i))[:total - sum(out)]:
            out[i] += 1
        return out

    @classmethod
    def kirim_alloc(cls, total, needs, stocks, step=10):
        """Kelgan «total» ni aptekalarga bo'lish (butun sonlar):
        1) ehtiyoj bo'yicha: yetsa — har biriga ehtiyoji, yetmasa — ehtiyoj nisbatida;
        2) step (10) va undan katta miqdor step ga pastga yaxlitlanadi (11 → 10, 23 → 20), kichigi o'zgarmaydi;
        3) qoldiqdan: shu dori umuman yo'q (qoldiq 0) va hech narsa olmagan aptekaga 1 dona (assortiment);
        4) qolgani skladda."""
        total = int(total)
        needs = [max(0, int(math.ceil(n - 1e-9))) for n in needs]
        if total <= 0:
            return [0] * len(needs)
        out = list(needs) if total >= sum(needs) else cls._share(total, needs)
        step = max(1, int(step))
        out = [a if a < step else a // step * step for a in out]
        left = total - sum(out)
        for i in range(len(out)):
            if left <= 0:
                break
            if out[i] == 0 and stocks[i] <= 0:
                out[i] = 1
                left -= 1
        return out

    @property
    def kirim_qadam(self):
        try:
            return max(1, int(float(self.cfg.get("fom", "kirim_qadam", fallback="10"))))
        except ValueError:
            return 10

    @property
    def kirim_zaxira(self):
        try:
            return min(100.0, max(0.0, float(self.cfg.get("fom", "kirim_zaxira", fallback="0"))))
        except ValueError:
            return 0.0

    def _kirim_day(self, sana):
        try:
            return datetime.date.fromisoformat(str(sana)) if sana else now().date()
        except ValueError:
            return now().date()

    def kirim_items(self, day):
        """[{name, producer, supplier, avail, exp, sales{}, stock{}, alloc{}}] — kirimdagi har bir dori."""
        import fom_sql
        parts = self.fom_sql.arrivals(day)
        sale = {}
        for r in self.data.sale:
            sale.setdefault((norm(r["name"]), norm(r["producer"])), r)
        items = {}
        for p in parts:
            k = (fom_sql.key(p["name"]), fom_sql.key(p["producer"]))
            it = items.get(k)
            if it is None:
                it = items[k] = {"name": p["name"], "producer": p["producer"], "suppliers": [], "avail": 0.0, "exp": p["exp"]}
            it["avail"] += p["avail"]
            if p["supplier"] and p["supplier"] not in it["suppliers"]:
                it["suppliers"].append(p["supplier"])
        # bugun kirimdan allaqachon bron qilingani (bekor qilinmagan): ehtiyojdan ayriladi — ikki marta berilmaydi
        given = {}
        for g in self.db.execute(
                "SELECT b.dst, i.name, i.producer, SUM(i.qty) q FROM fom_boron_items i JOIN fom_boron b ON b.id=i.boron_id "
                "WHERE b.sana=? AND b.cats='kirim' AND b.status NOT IN ('xato','bekor') GROUP BY b.dst, i.name, i.producer",
                (self._today(),)):
            k = (norm(g["dst"]), norm(g["name"]), norm(g["producer"]))
            given[k] = given.get(k, 0.0) + (g["q"] or 0)
        z = self.kirim_zaxira
        out = []
        for it in sorted(items.values(), key=lambda x: norm(x["name"])):
            r = sale.get((norm(it["name"]), norm(it["producer"])))
            v = r["v"] if r else [0.0] * 7
            sales = [max(0.0, v[i + 1]) for _, i in self.KIRIM_DST]
            gv = [given.get((norm(d), norm(it["name"]), norm(it["producer"])), 0.0) for d, _ in self.KIRIM_DST]
            stocks = [v[i] + gv[n] for n, (_, i) in enumerate(self.KIRIM_DST)]
            needs = [max(0.0, sales[n] * self.cover - stocks[n]) for n in range(len(sales))]
            for n in range(len(gv)):   # bugun shu dori bo'yicha bron olgan filial — qayta taklif qilinmaydi (assortiment ham)
                if gv[n] > 0:
                    needs[n], stocks[n] = 0.0, max(stocks[n], 1.0)
            total = int(it["avail"] * (100.0 - z) / 100.0 + 1e-9)
            alloc = self.kirim_alloc(total, needs, stocks, self.kirim_qadam)
            out.append({"name": it["name"], "producer": it["producer"], "supplier": ", ".join(it["suppliers"]),
                        "avail": it["avail"], "exp": it["exp"], "in_sale": bool(r),
                        "sales": {d: sales[n] for n, (d, _) in enumerate(self.KIRIM_DST)},
                        "stock": {d: v[i] for d, i in self.KIRIM_DST},
                        "given": {d: gv[n] for n, (d, _) in enumerate(self.KIRIM_DST)},
                        "need": {d: int(math.ceil(needs[n] - 1e-9)) for n, (d, _) in enumerate(self.KIRIM_DST)},
                        "alloc": {d: alloc[n] for n, (d, _) in enumerate(self.KIRIM_DST)}})
        return out

    def _kirim_sig(self, items):
        raw = "|".join("%s/%s=%s" % (norm(i["name"]), norm(i["producer"]), qty(i["avail"])) for i in items)
        return hashlib.md5(raw.encode("utf-8")).hexdigest()[:10]

    def kirim_summary(self, sana=None):
        import fom_sql
        day = self._kirim_day(sana)
        base = {"sana": day.isoformat(), "dst": [d for d, _ in self.KIRIM_DST], "zaxira": self.kirim_zaxira,
                "qadam": self.kirim_qadam,
                "period": self.data.period_days}
        if not self.fom_sql.enabled:
            return dict(base, ok=False, error="Bugungi kirim FOM'dan o'qiladi: config.ini [fom] mode = sql kerak")
        try:
            items = self.kirim_items(day)
        except fom_sql.FomError as e:
            return dict(base, ok=False, error=str(e))
        return dict(base, ok=True, db=self.fom_sql.database, items=items, sig=self._kirim_sig(items))

    def kirim_create(self, sana, alloc, sig, by="admin (panel)"):
        """alloc: {dorixona: {dori raqami (items ro'yxatidagi): soni}}. Har dorixona uchun bitta FOM qoralama bron."""
        import fom_sql
        day = self._kirim_day(sana)
        if not self.fom_sql.enabled:
            raise fom_sql.FomError("config.ini [fom] mode = sql kerak")
        items = self.kirim_items(day)
        if self._kirim_sig(items) != sig:
            raise fom_sql.FomError("Kirim ro'yxati o'zgargan (shu orada sotildi yoki bron qilindi) — sahifani yangilang")
        dsts = [d for d, _ in self.KIRIM_DST]
        used = [0.0] * len(items)
        want = {}
        for d, m in (alloc or {}).items():
            if d not in dsts:
                raise fom_sql.FomError("Noma'lum dorixona: %s" % d)
            for k, q in (m or {}).items():
                k, q = int(k), float(q or 0)
                if q < 0 or k < 0 or k >= len(items) or q != int(q):
                    raise fom_sql.FomError("Noto'g'ri miqdor")
                if q > 0:
                    used[k] += q
                    want.setdefault(d, []).append({"name": items[k]["name"], "producer": items[k]["producer"],
                                                   "qty": q, "task_id": -(k + 1), "exp": items[k]["exp"]})
        for k, u in enumerate(used):
            if u > items[k]["avail"] + 1e-9:
                raise fom_sql.FomError("«%s»: kelgan %s, taqsimlangan %s" % (items[k]["name"], qty(items[k]["avail"]), qty(u)))
        if not want:
            raise fom_sql.FomError("Taqsimotda dori yo'q")
        ts = now().isoformat(timespec="seconds")
        made, plans = [], None
        with self.dblock:
            # qayta tekshiruv qulf ichida: ikki marta bosish / ikki oyna bir xil taqsimotni ikki marta yozmasin
            if self._kirim_sig(self.kirim_items(day)) != sig:
                raise fom_sql.FomError("Kirim ro'yxati o'zgargan (shu orada bron qilindi) — sahifani yangilang")
            try:
                bids = {}
                for d in dsts:
                    rows = want.get(d)
                    if not rows:
                        continue
                    cur = self.db.execute(
                        "INSERT INTO fom_boron(sana, src, dst, cats, n, qty, created_by, created_at, mode, status) "
                        "VALUES(?,?,?,?,?,?,?,?,?,?)",
                        (self._today(), self.sklad, d, "kirim", len(rows), sum(r["qty"] for r in rows), by, ts, "sql", "fayl"))
                    bids[d] = cur.lastrowid
                    self.db.executemany(
                        "INSERT INTO fom_boron_items(boron_id, task_id, name, producer, cat, qty, srok) VALUES(?,?,?,?,?,?,?)",
                        [(cur.lastrowid, r["task_id"], r["name"], r["producer"], "kirim", r["qty"], r["exp"]) for r in rows])
                    made.append(cur.lastrowid)
                plans = self.fom_sql.write_kirim(day, {d: want[d] for d in bids}, {d: "BOT kirim #%d" % b for d, b in bids.items()})
                for p in plans:
                    bid = bids[p["dst"]]
                    self.db.executemany(
                        "INSERT INTO fom_boron_lines(boron_id, task_id, incomeln, kol, price, srok) VALUES(?,?,?,?,?,?)",
                        [(bid, ln["task_id"], ln["incomeln"], ln["kol"], ln["price"], ln["exp"]) for ln in p["lines"]])
                    self.db.execute("UPDATE fom_boron SET status='qoralama', doc_no=?, short='[]' WHERE id=?",
                                    (str(p["invoice"] or ""), bid))
                self.db.commit()
            except Exception:
                self.db.rollback()
                self._fom_orphan_alert(plans, "kirim")
                raise
        made = [self.db.execute("SELECT * FROM fom_boron WHERE id=?", (b,)).fetchone() for b in made]
        text = "📦 <b>Kirim taqsimoti → FOM bron</b> · %s\n%s · %s\n%s" % (esc(by), esc(self.sklad), day.isoformat(), "\n".join(
            "#%d → %s · %d ta dori · %s qadoq%s" % (b["id"], esc(b["dst"]), b["n"], qty(b["qty"]), self._fom_tail(b)) for b in made))
        for a in self.admins:
            try:
                self.tg.send(a, text)
            except Exception:  # noqa
                logging.exception("Adminga xabar yuborib bo'lmadi: %s", a)
        return made

    # ================= hisobni tasdiqlash so'rovlari =================
    def acc_candidates(self, c):
        """Firma telefoni bo'yicha topilgan, hali tasdiqlanmagan tashkilotlar."""
        if c["role"] == "staff":
            return []
        approved = {norm(a) for a in self.accounts_of(c["chat_id"])}
        return [a for a in self.accounts_for(self.tags_for(c), (), auto=True) if a not in approved]

    def ensure_acc_requests(self, cid, notify=True):
        """Nomzod tashkilotlar uchun so'rov yaratadi (adminga xabar yuboradi). Kutilayotgan so'rovlar sonini qaytaradi."""
        c = self.client(cid)
        if not c:
            return 0
        self.data.reload()
        created = []
        for acc in self.acc_candidates(c):
            with self.dblock:
                cur = self.db.execute("INSERT OR IGNORE INTO acc_requests(chat_id,account,status,created) VALUES(?,?,'pending',?)",
                                      (cid, acc, now().isoformat(timespec="seconds")))
                self.db.commit()
            if cur.rowcount:
                created.append(cur.lastrowid)
        if notify:
            for rid in created:
                self.notify_acc_request(rid)
        return self.db.execute("SELECT COUNT(*) FROM acc_requests WHERE chat_id=? AND status='pending'", (cid,)).fetchone()[0]

    def acc_request_info(self, rid):
        r = self.db.execute("SELECT * FROM acc_requests WHERE id=?", (rid,)).fetchone()
        if not r:
            return None
        c = self.client(r["chat_id"])
        rows = self.data.account_rows([r["account"]])
        net = sum(x["debit"] - x["credit"] for x in rows)
        return {"id": rid, "chat_id": r["chat_id"], "status": r["status"], "account": r["account"],
                "name": (c["name"] if c else "") or "", "phone": (c["phone"] if c else "") or "",
                "firms": self.firms_of(r["chat_id"]), "net": net, "found": bool(rows), "created": (r["created"] or "")[:16].replace("T", " ")}

    def acc_pending(self):
        ids = [r["id"] for r in self.db.execute("SELECT id FROM acc_requests WHERE status='pending' ORDER BY id")]
        return [x for x in (self.acc_request_info(i) for i in ids) if x]

    def notify_acc_request(self, rid):
        i = self.acc_request_info(rid)
        if not i:
            return
        bal = ("%s %s" % (money(abs(i["net"])), "kredit" if i["net"] < 0 else "debet")) if i["found"] and i["net"] else "0"
        text = ("💰 <b>Hisobni tasdiqlash so'rovi</b>\nMijoz: %s · %s\nFirma: %s\nTashkilot (hisob): %s\nQoldiq: %s\n\n"
                "Tasdiqlansa, mijoz shu hisobni ko'ra boshlaydi. Server panelida ham tasdiqlash mumkin.") % (
            esc(i["name"]), esc(i["phone"]), esc(", ".join(i["firms"]) or "—"), esc(i["account"]), bal)
        for a in self.admins:
            try:
                self.tg.send(a, text, inline_kb([[("✅ Tasdiqlash", "ak:%d" % rid), ("❌ Rad etish", "ar:%d" % rid)], [("🏠 Menyu", "a:home")]]))
            except Exception:  # noqa
                logging.exception("Adminga so'rov yuborib bo'lmadi: %s", a)

    def acc_resolve(self, rid, approve):
        """True/False: bajarildi. None: so'rov topilmadi yoki allaqachon hal qilingan."""
        r = self.db.execute("SELECT * FROM acc_requests WHERE id=?", (rid,)).fetchone()
        if not r or r["status"] != "pending":
            return None
        with self.dblock:
            self.db.execute("UPDATE acc_requests SET status=? WHERE id=?", ("approved" if approve else "rejected", rid))
            if approve:
                self.db.execute("INSERT OR IGNORE INTO client_accounts VALUES(?,?)", (r["chat_id"], norm(r["account"])))
            self.db.commit()
        if approve:
            c = self.client(r["chat_id"])
            if c and c["status"] == "active":
                try:
                    self.tg.send(r["chat_id"], "✅ <b>Hisobingiz tasdiqlandi.</b> «💰 Hisob» tugmasini bosing.",
                                 inline_kb([[("💰 Hisob", "c:acc"), ("🏠 Menyu", "c:home")]]))
                except Exception:  # noqa
                    logging.exception("Mijozga xabar yuborib bo'lmadi")
        return True

    def acc_requests_list(self, cid, mid=None, note=None):
        rows = self.acc_pending()
        cr = ("👥 Foydalanuvchilar", "💰 Hisob so'rovlari")
        if not rows:
            return self.a_screen(cid, mid, cr, "💰 Tasdiqlanishi kerak bo'lgan hisob so'rovlari yo'q.", back="a:usr", note=note)
        lines, kb = [], []
        for n, i in enumerate(rows[:8], 1):
            bal = ("%s %s" % (money(abs(i["net"])), "kredit" if i["net"] < 0 else "debet")) if i["found"] and i["net"] else "0"
            lines.append("%d. %s · %s\n   🏢 %s\n   💰 %s: %s" % (n, esc(i["name"]), esc(i["phone"]), esc(", ".join(i["firms"]) or "—"), esc(self.short(i["account"], 40)), bal))
            kb.append([("✅ %d" % n, "ak:%d" % i["id"]), ("❌ %d" % n, "ar:%d" % i["id"])])
        return self.a_screen(cid, mid, cr, "<b>💰 Hisob tasdiqlash so'rovlari</b> (%d)\n\n%s" % (len(rows), "\n".join(lines)), kb,
                             back="a:usr", note=note)

    # ================= foydalanuvchilar =================
    def user_counts(self):
        q = lambda sql: self.db.execute(sql).fetchone()[0]
        return {"pending": q("SELECT COUNT(*) FROM clients WHERE status='pending'"),
                "active": q("SELECT COUNT(*) FROM clients WHERE status='active'"),
                "staff": q("SELECT COUNT(*) FROM clients WHERE role='staff' AND status='active'"),
                "filial": q("SELECT COUNT(*) FROM clients WHERE role='filial' AND status='active'"),
                "blocked": q("SELECT COUNT(*) FROM clients WHERE status='blocked'"),
                "inv": q("SELECT COUNT(*) FROM invites"), "all": q("SELECT COUNT(*) FROM clients"),
                "accreq": q("SELECT COUNT(*) FROM acc_requests WHERE status='pending'")}

    def users_menu(self, cid, mid=None, note=None):
        n = self.user_counts()
        text = "<b>👥 Foydalanuvchilar</b>\nFaol: <b>%d</b> · kutmoqda: <b>%d</b> · bloklangan: <b>%d</b>\nTakliflar (hali kirmagan): <b>%d</b>" % (
            n["active"], n["pending"], n["blocked"], n["inv"])
        kb = inline_kb([[("⏳ Kutmoqda (%d)" % n["pending"], "ul:pending:0"), ("✅ Faol (%d)" % n["active"], "ul:active:0")],
                        [("👔 Xodimlar (%d)" % n["staff"], "ul:staff:0"), ("⛔ Bloklangan (%d)" % n["blocked"], "ul:blocked:0")],
                        [("🏬 Filial xodimlari (%d)" % n["filial"], "ul:filial:0")],
                        [("📨 Takliflar (%d)" % n["inv"], "ul:inv:0"), ("🔍 Qidirish", "us")],
                        [("➕ Foydalanuvchi qo'shish", "ua"), ("📋 Hammasi (%d)" % n["all"], "ul:all:0")],
                        [("💰 Hisob so'rovlari (%d)" % n["accreq"], "ul:accreq:0")]])
        return self.a_screen(cid, mid, ("👥 Foydalanuvchilar",), text, kb, note=note)

    USER_Q = {"pending": "WHERE status='pending'", "active": "WHERE status='active'", "blocked": "WHERE status='blocked'",
              "staff": "WHERE role='staff' AND status='active'", "filial": "WHERE role='filial' AND status='active'", "all": ""}

    def _user_btn(self, r, pick=False):
        icon = {"active": "✅", "pending": "⏳", "blocked": "⛔"}.get(r["status"], "•")
        firms = self.role_label(r)
        firms = self.short(firms, 30 if r["role"] == "filial" else 20)
        label = "%s %s · %s" % (icon, bl(self.short(r["name"] or r["phone"], 18)), firms if r["role"] in ("staff", "filial") else bl(firms))
        return (label, ("m1:%d" if pick else "uc:%d") % r["chat_id"])

    def users_list(self, cid, mid, f, page, pick=False, note=None):
        per = 8
        tail = ":p" if pick else ""
        back = "a:bc" if pick else "a:usr"
        names = {"pending": "⏳ Kutmoqda", "active": "✅ Faol", "blocked": "⛔ Bloklangan", "staff": "👔 Xodimlar",
                 "filial": "🏬 Filial xodimlari", "all": "📋 Hammasi", "inv": "📨 Takliflar"}
        cr = ("📢 Xabar yuborish", "👤 Bitta odamga") if pick else ("👥 Foydalanuvchilar", names.get(f, "Ro'yxat"))
        if f == "inv":
            rows = self.db.execute("SELECT * FROM invites ORDER BY created DESC").fetchall()
            part = rows[page * per:(page + 1) * per]
            if not part:
                return self.a_screen(cid, mid, cr, "📨 Takliflar ro'yxati bo'sh.", back=back, note=note)
            inv_role = lambda r: ("xodim" if r["role"] == "staff" else "filial · " + (r["filial"] or "—") if r["role"] == "filial"
                                  else (r["firms"] or "—").replace(";", ", "))
            lines = ["📨 %s · %s · %s" % (esc(r["phone"]), esc(r["name"] or "—"), esc(inv_role(r))) for r in part]
            kb = [[("🗑 %s" % r["phone"], "ud:%s" % r["phone9"])] for r in part]
            nav = []
            if page > 0:
                nav.append(("◀", "ul:inv:%d" % (page - 1)))
            if (page + 1) * per < len(rows):
                nav.append(("▶", "ul:inv:%d" % (page + 1)))
            return self.a_screen(cid, mid, cr, "<b>📨 Takliflar</b> (%d)\nBotga hali kirmagan, kirganda avto-qabul qilinadi:\n\n%s" % (
                len(rows), "\n".join(lines)), kb + ([nav] if nav else []), back=back, note=note)
        rows = self.db.execute("SELECT * FROM clients %s ORDER BY created DESC" % self.USER_Q.get(f, "")).fetchall()
        part = rows[page * per:(page + 1) * per]
        if not part:
            return self.a_screen(cid, mid, cr, "Ro'yxat bo'sh.", back=back, note=note)
        kb = [[self._user_btn(r, pick)] for r in part]
        nav = []
        if page > 0:
            nav.append(("◀", "ul:%s:%d%s" % (f, page - 1, tail)))
        if (page + 1) * per < len(rows):
            nav.append(("▶", "ul:%s:%d%s" % (f, page + 1, tail)))
        title = ("👤 <b>Kimga yuboramiz?</b>\nRo'yxatdan tanlang yoki ism, telefon, firma, ID yozing 👇 (qidiruv).\nBekor qilish: /cancel\n"
                 if pick else "<b>%s</b>\n" % names.get(f, "Ro'yxat")) + "%d-%d / %d" % (page * per + 1, page * per + len(part), len(rows))
        return self.a_screen(cid, mid, cr, title, kb + ([nav] if nav else []), back=back, note=note)

    def find_users(self, q):
        qn, qd = norm(q), digits(q)
        out = []
        for r in self.db.execute("SELECT * FROM clients ORDER BY created DESC").fetchall():
            hay = " ".join([norm(r["name"]), norm(r["username"]), norm(" ".join(self.firms_of(r["chat_id"]))), norm(r["filial"]),
                            str(r["chat_id"])])
            if qn in hay or (len(qd) >= 4 and qd in digits(r["phone"])):
                out.append(r)
        return out[:10]

    def add_user(self, text):
        m = re.match(r"^\s*(\+?[\d\s()\-]{9,}?)\s*(?=[^\d\s()\-+]|$)(.*)$", text, re.S)
        if not m:
            return "❌ Telefon raqami topilmadi. Namuna: <code>+998901234567 АБДУВОХИД ТОШКЕНТ 94-5560909</code>"
        phone, rest = m.group(1).strip(), m.group(2).strip()
        d9 = digits(phone)[-9:]
        if len(d9) < 9:
            return "❌ Telefon raqami to'liq emas."
        staff = rest.lower() in ("xodim", "staff", "ходим")
        fil = None
        mf = re.match(r"^(?:filial|филиал)\b\s*(.*)$", rest, re.I | re.S)
        if mf:  # +998901234567 filial ЧИНОЗ
            self.data.reload()
            fil, cand = self.find_filial(mf.group(1))
            if not fil:
                names = self.data.koch_filials()
                if not names:
                    return "❌ Ko'chirish varag'ida filial topilmadi. Avval «🔄 Yangilash» ni bosing."
                return ("❌ Filial %s: «%s».\nMavjud filiallar: %s" % (
                    "aniq emas (bir nechtasi mos keldi)" if cand else "topilmadi", esc(mf.group(1).strip()),
                    esc(", ".join(cand or names))))
        firms = [] if (staff or fil) else [norm(x) for x in re.split(r"[;\n|]", rest) if x.strip()]
        if not staff and not fil and not firms:
            return "❌ Telefondan keyin firma belgisini yozing (yoki «xodim», yoki «filial ЧИНОЗ»)."
        role = "staff" if staff else "filial" if fil else "firm"
        known = set(self.data.all_tags())
        warn = ("\n⚠ Ma'lumotda topilmadi: %s" % esc(", ".join(t for t in firms if t not in known))) if any(t not in known for t in firms) else ""
        label = "xodim" if staff else ("filial xodimi · " + fil) if fil else ", ".join(firms)
        for r in self.db.execute("SELECT * FROM clients").fetchall():
            if digits(r["phone"])[-9:] == d9:
                self.set_role(r["chat_id"], role, fil)
                self.set_firms(r["chat_id"], firms)
                self.activate(r["chat_id"])
                return "✅ Mijoz botda bor edi: yangilandi va faollashtirildi (%s → %s)%s" % (esc(r["name"]), esc(label), warn)
        with self.dblock:
            self.db.execute("INSERT OR REPLACE INTO invites(phone9, phone, name, role, firms, accounts, created, filial) "
                            "VALUES(?,?,?,?,?,?,?,?)",
                            (d9, phone, "", role, ";".join(firms), "", now().isoformat(timespec="seconds"), fil))
            self.db.commit()
        return "✅ Taklif qo'shildi: %s → %s%s\nFoydalanuvchi botga kirib kontaktini yuborganda avtomatik qabul qilinadi." % (
            esc(phone), esc(label), warn)

    # ================= xabar yuborish: ommaviy va yakka =================
    def bc_menu(self, cid, mid=None):
        n = len(self.audience("all"))
        text = ("<b>📢 Xabar yuborish</b>\n\n<b>Ommaviy</b> — hamma faol foydalanuvchilarga (<b>%d</b> ta). "
                "Har firma faqat o'z ma'lumotini oladi.\n<b>Yakka</b> — bitta tanlangan odamga (matn yoki hisobot).") % n
        kb = inline_kb([[("⏰ Srok", "a:bq:srok"), ("📦 30+ kun", "a:bq:old")],
                        [("🧾 Buyurtma", "a:bq:need"), ("💰 Hisob", "a:bq:acc")],
                        [("✉️ Ommaviy matn", "a:bq:text")],
                        [("👤 Bitta odamga", "a:bq:one")]])
        return self.a_screen(cid, mid, ("📢 Xabar yuborish",), text, kb)

    def one_menu(self, cid, target, mid=None, note=None):
        c = self.client(target)
        if not c:
            return self.users_menu(cid, mid, "Mijoz topilmadi.")
        firms = "xodim" if c["role"] == "staff" else (", ".join(self.firms_of(target)) or "—")
        kb = inline_kb([[("✉️ Matn yozish", "m2:%d:text")],
                        [("⏰ Srok", "m2:%d:srok"), ("📦 30+ kun", "m2:%d:old")],
                        [("🧾 Buyurtma", "m2:%d:need"), ("💰 Hisob", "m2:%d:acc")],
                        [("📊 Qoldiq", "m2:%d:stock"), ("💲 Narxlar", "m2:%d:narx")],
                        self.nav_row("a:bq:one", "a:home", "◀ Boshqa odam")])
        for row in kb["inline_keyboard"]:
            for b in row:
                if "%d" in b["callback_data"]:
                    b["callback_data"] = b["callback_data"] % target
        return self.a_screen(cid, mid, ("📢 Xabar yuborish", esc(c["name"] or c["phone"] or str(target))),
                             "👤 <b>%s</b>\n%s · %s\nNimani yuboramiz?" % (esc(c["name"]), esc(c["phone"]), esc(firms)), kb, back=None, note=note)

    def one_send(self, cid, target, kind, mid=None):
        c = self.client(target)
        if not c:
            return self.users_menu(cid, mid, "Mijoz topilmadi.")
        if kind == "text":
            return self.ask(cid, mid, ("📢 Xabar yuborish", esc(c["name"] or str(target)), "✉️ Matn"),
                            "✉️ <b>%s</b> ga yuboriladigan matnni yozing 👇\nBekor qilish: /cancel" % esc(c["name"]), ("one_text", target),
                            back="m1:%d" % target)
        if c["status"] != "active" or not self.can_use(c):
            return self.one_menu(cid, target, mid, "⚠ %s faol emas yoki firma biriktirilmagan, hisobot yuborib bo'lmaydi." % esc(c["name"]))
        accs = self.accs_for_client(c) if kind == "acc" else None
        try:
            ok = self.send_report(target, kind, self.tags_for(c), accs, quiet=True)
        except Exception as e:  # noqa
            if self.mark_gone(target, e):
                return self.one_menu(cid, target, mid, "🚫 %s botni bloklagan (tark etgan), yuborib bo'lmaydi." % esc(c["name"]))
            logging.exception("Yakka yuborish xatosi")
            return self.one_menu(cid, target, mid, "❌ Yuborilmadi: %s" % esc(e))
        return self.one_menu(cid, target, mid, ("✅ Yuborildi: %s → %s" if ok else "➖ Ma'lumot yo'q: %s → %s (hech narsa yuborilmadi)") % (
            esc(c["name"]), self.KINDS.get(kind, kind)))

    def log_text(self, n=25):
        path = os.path.join(BASE_DIR, "bot.log")
        try:
            with open(path, "rb") as f:
                f.seek(0, 2)
                f.seek(max(0, f.tell() - 60000))
                raw = f.read().decode("utf-8", "replace")
        except OSError:
            return "Jurnal fayli topilmadi."
        lines = [ln for ln in raw.splitlines() if re.match(r"^\d{4}-\d\d-\d\d ", ln)][-n:]
        return "<b>📜 Jurnal</b> (oxirgi %d)\n<pre>%s</pre>" % (len(lines), esc("\n".join(l[:150] for l in lines))[-3300:])

    # ================= qora ro'yxat va buyurtma (yetkazib beruvchilar) =================
    def blacklist(self):
        """{firma: (sabab, sana)} — buyurtma berilmaydigan yetkazib beruvchilar."""
        return {r["firm"]: (r["reason"], r["created"]) for r in self.db.execute("SELECT * FROM supplier_blacklist")}

    def item_blacklist(self):
        """{(firma, tovar): (sabab, sana)} — buyurtma qilinmaydigan alohida tovarlar."""
        return {(r["firm"], r["item"]): (r["reason"], r["created"]) for r in self.db.execute("SELECT * FROM item_blacklist")}

    def need_rows(self, tags):
        """Buyurtma tavsiyasi (bot va hisobotlar uchun): qora ro'yxatdagi firma va tovarlar chiqarib tashlanadi."""
        bl, ibl = self.blacklist(), self.item_blacklist()
        return [r for r in self.data.need_rows(tags, self.cover)
                if norm(r["firm"]) not in bl and (norm(r["firm"]), norm(r["name"])) not in ibl]

    def price_map(self):
        """Narx varag'idan kirim narxi: (tovar, ishlab chiqaruvchi) bo'yicha, bo'lmasa faqat tovar nomi bo'yicha."""
        d = self.data
        key = (d.path, d.mtime)
        if getattr(self, "_pm_key", None) != key:
            full, by_name = {}, {}
            for r in sorted(d.narx, key=lambda r: r["sana"] or datetime.date.min):  # eng oxirgi kirim ustun
                if r["prih"]:
                    full[(norm(r["name"]), norm(r["producer"]))] = r["prih"]
                    by_name[norm(r["name"])] = r["prih"]
            self._pm, self._pm_key = (full, by_name), key
        return self._pm

    def price_of(self, name, producer):
        full, by_name = self.price_map()
        return full.get((norm(name), norm(producer))) or by_name.get(norm(name)) or 0.0

    @staticmethod
    def guess_phone(firm):
        d9 = digits(firm)[-9:]
        return "+998" + d9 if len(d9) == 9 else ""

    def order_suppliers(self, cover_days=30):
        """Panelning «Buyurtma» bo'limi: har bir yetkazib beruvchi (firma belgisi) bo'yicha nima buyurtma qilish kerak.
        buyurtma = kunlik sotuv × zaxira kunlari − jami qoldiq (4 filial + sklad)."""
        d = self.data
        d.reload()
        period = max(1, d.period_days)
        bl, ibl = self.blacklist(), self.item_blacklist()
        contacts = {r["firm"]: r for r in self.db.execute("SELECT * FROM supplier_contacts")}
        last = {r["firm"]: r["created"] for r in self.db.execute("SELECT firm, MAX(created) AS created FROM orders GROUP BY firm")}
        groups = {}
        for r in d.sale:
            firm = norm(r["firm"])
            if not firm or r["sales"] <= 0:
                continue
            daily = r["sales"] / period
            need = int(-(-(daily * cover_days - r["stock"]) // 1))  # yuqoriga yaxlitlash
            if need <= 0:
                continue
            price = self.price_of(r["name"], r["producer"])
            item = {"name": r["name"], "producer": r["producer"], "v": [r["v"][0], r["v"][2], r["v"][4], r["v"][6]],
                    "stock": r["stock"], "sales": r["sales"], "days": int(r["stock"] / daily) if r["stock"] > 0 else 0,
                    "need": need, "price": price}
            g = groups.setdefault(firm, {"items": [], "hidden": []})
            if (firm, norm(r["name"])) in ibl:
                reason, when = ibl[(firm, norm(r["name"]))]
                g["hidden"].append(dict(item, reason=reason or "", when=(when or "")[:10]))
            else:
                g["items"].append(item)
        out = []
        for firm, g in groups.items():
            items = sorted(g["items"], key=lambda x: (x["stock"] > 0, x["days"], norm(x["name"])))
            if not items and not g["hidden"]:
                continue
            zero = sum(1 for x in items if x["stock"] <= 0)
            level = "r" if zero else ("a" if any(x["days"] < 7 for x in items) else "g")
            ct = contacts.get(firm)
            out.append({"firm": firm, "items": items, "hidden": g["hidden"], "n": len(items), "zero": zero, "level": level,
                        "total": sum(x["need"] * x["price"] for x in items),
                        "last_order": (last.get(firm) or "")[:16].replace("T", " "),
                        "phone": (ct["phone"] if ct and ct["phone"] else self.guess_phone(firm)),
                        "username": (ct["username"] if ct else "") or "",
                        "blacklisted": firm in bl, "bl_reason": (bl.get(firm) or ("", ""))[0] or "",
                        "bl_when": ((bl.get(firm) or ("", ""))[1] or "")[:10]})
        rank = {"r": 0, "a": 1, "g": 2}
        out.sort(key=lambda s: (s["blacklisted"], not s["items"], rank[s["level"]], -s["total"]))
        return out

    def set_blacklist(self, firm, on, reason=""):
        firm = norm(firm)
        with self.dblock:
            if on:
                self.db.execute("INSERT OR REPLACE INTO supplier_blacklist VALUES(?,?,?)",
                                (firm, reason.strip(), now().isoformat(timespec="seconds")))
            else:
                self.db.execute("DELETE FROM supplier_blacklist WHERE firm=?", (firm,))
            self.db.commit()
        logging.info("Qora ro'yxat: %s %s (%s)", firm, "qo'shildi" if on else "olib tashlandi", reason)

    def set_item_blacklist(self, firm, item, on, reason=""):
        firm, item = norm(firm), norm(item)
        with self.dblock:
            if on:
                self.db.execute("INSERT OR REPLACE INTO item_blacklist VALUES(?,?,?,?)",
                                (firm, item, reason.strip(), now().isoformat(timespec="seconds")))
            else:
                self.db.execute("DELETE FROM item_blacklist WHERE firm=? AND item=?", (firm, item))
            self.db.commit()

    def set_contact(self, firm, phone, username):
        username = (username or "").strip().lstrip("@")
        with self.dblock:
            self.db.execute("INSERT OR REPLACE INTO supplier_contacts VALUES(?,?,?,?)",
                            (norm(firm), (phone or "").strip(), username, now().isoformat(timespec="seconds")))
            self.db.commit()

    def save_order(self, firm, cover, items):
        """«Buyurtma berildi»: tarix uchun yoziladi (keyin «oxirgi buyurtma» sanasi ko'rinadi)."""
        items = [i for i in items if to_float(i.get("qty")) > 0]
        total = sum(to_float(i.get("qty")) * to_float(i.get("price")) for i in items)
        with self.dblock:
            cur = self.db.execute("INSERT INTO orders(firm, created, cover, items, total, n) VALUES(?,?,?,?,?,?)",
                                  (norm(firm), now().isoformat(timespec="seconds"), int(cover),
                                   json.dumps(items, ensure_ascii=False), total, len(items)))
            self.db.commit()
        logging.info("Buyurtma berildi: %s · %d ta tovar · %s", norm(firm), len(items), money(total))
        return cur.lastrowid, len(items), total

    def orders_history(self, firm=None, limit=30):
        q, a = "SELECT * FROM orders", []
        if firm:
            q, a = q + " WHERE firm=?", [norm(firm)]
        rows = self.db.execute(q + " ORDER BY id DESC LIMIT ?", a + [limit]).fetchall()
        return [{"id": r["id"], "firm": r["firm"], "created": (r["created"] or "")[:16].replace("T", " "),
                 "n": r["n"], "total": r["total"], "cover": r["cover"]} for r in rows]

    # ================= dori qidirish =================
    def search_text(self, q, tags, label=None, firm_view=False):
        """Dori nomining bir qismi bo'yicha: qoldiq (filiallar), sotuv, eng yaqin srok va narx.
        firm_view=True (oddiy firma): faqat kirim narxi ko'rsatiladi."""
        d = self.data
        d.reload()
        words = canon(q).split()
        match = lambda name: bool(words) and all(w in canon(name) for w in words)
        found = {}
        for r in d._by_tags(d.sale, tags):
            if match(r["name"]):
                k = (norm(r["name"]), norm(r["producer"]))
                g = found.setdefault(k, {"name": r["name"], "producer": r["producer"], "firms": set(),
                                         "v": [0.0] * 4, "stock": 0.0, "sales": 0.0})
                g["firms"].add(r["firm"])
                for i, j in enumerate((0, 2, 4, 6)):
                    g["v"][i] += r["v"][j]
                g["stock"] += r["stock"]
                g["sales"] += r["sales"]
        if not found:  # sotuvda yo'q, lekin narx varag'ida bor
            for r in d._by_tags(d.narx, tags):
                if match(r["name"]):
                    found.setdefault((norm(r["name"]), norm(r["producer"])),
                                     {"name": r["name"], "producer": r["producer"], "firms": {r["firm"]},
                                      "v": None, "stock": 0.0, "sales": 0.0})
        h = "🔍 <b>«%s»</b>" % esc(q.strip())
        if label:
            h += " · 🏢 %s" % esc(label)
        if not found:
            return h + "\n\nHech narsa topilmadi. Dori nomining bir qismini yozing (kamida 3 harf)."
        srok = {}
        for r in d._by_tags(d.srok, tags):
            k = (norm(r["name"]), norm(r["producer"]))
            if k in found and r["date"] and r["ost"] > 0:
                s = srok.setdefault(k, {})
                s[r["date"]] = s.get(r["date"], 0) + r["ost"]
        narx = {}
        for r in sorted(d._by_tags(d.narx, tags), key=lambda r: r["sana"] or datetime.date.min):
            k = (norm(r["name"]), norm(r["producer"]))
            if k in found:
                narx[k] = r
        items = sorted(found.items(), key=lambda kv: (-kv[1]["sales"], norm(kv[1]["name"])))
        today = datetime.date.today()
        blocks = []
        for k, g in items[:10]:
            t = "💊 <b>%s</b>%s" % (esc(g["name"]), (" · " + esc(self.short(g["producer"], 30))) if g["producer"] else "")
            if g["v"] is not None:
                t += "\n📦 Qoldiq: <b>%s</b> (Chinoz %s · Urgut %s · Erkin %s · Sklad %s)" % (
                    qty(g["stock"]), qty(g["v"][0]), qty(g["v"][1]), qty(g["v"][2]), qty(g["v"][3]))
                t += "\n📈 Sotuv (%d kun): <b>%s</b>" % (d.period_days, qty(g["sales"]))
            if k in srok:
                dt = min(srok[k])
                kun = (dt - today).days
                t += "\n⏰ Eng yaqin srok: %s (%s, %s dona)" % (dt.strftime("%d.%m.%Y"), self.kun_text(kun), qty(srok[k][dt]))
            if k in narx:
                n = narx[k]
                t += "\n💲 Kirim: %s" % money(n["prih"]) if firm_view else "\n💲 Kirim: %s · Sotuv: %s" % (money(n["prih"]), money(n["roz"]))
            firms = sorted(f for f in g["firms"] if f)
            if tags is None and firms:
                t += "\n🏢 " + esc(", ".join(firms)[:80])
            blocks.append(t)
        more = ("\n\n… yana %d ta. Aniqroq yozing." % (len(items) - 10)) if len(items) > 10 else ""
        return "%s · topildi: <b>%d</b>\n\n%s%s%s" % (h, len(items), "\n\n".join(blocks), more, self.foot())

    # ================= faollik statistikasi =================
    ACTION_NAMES = {"srok": "⏰ Srok", "old": "📦 30+ kun", "need": "🧾 Buyurtma", "stock": "📊 Sotuv va qoldiq", "acc": "💰 Hisob",
                    "search": "🔍 Qidiruv", "start": "🏠 Menyu", "narx": "💲 Narx", "koch": "🔁 Ko'chirish",
                    "kz": "🚨 Zudlik", "ko": "🔁 Oddiy", "kin": "📥 Keladigan", "kday": "📋 Bugungi holat", "ksk": "🏭 Taqsimot",
                    "kb": "✅ Berdim", "kx": "❌ Bera olmayman", "ka": "📥 Oldim",
                    "abc:s": "⏰ Srok (filial, ABC)", "abc:o": "📦 30+ kun (filial, ABC)",
                    **{k["key"]: "%s %s" % (k["icon"], k["title"]) for k in KOCH_CATS}}

    def action_name(self, a):
        if a.startswith("excel:"):
            return "📎 Excel (%s)" % self.ACTION_NAMES.get(a[6:], a[6:])
        return self.ACTION_NAMES.get(a, a)

    def stats(self, days=7):
        since = (now() - datetime.timedelta(days=days)).isoformat(timespec="seconds")
        since30 = (now() - datetime.timedelta(days=30)).isoformat(timespec="seconds")
        today = now().date().isoformat()
        one = lambda sql, *a: self.db.execute(sql, a).fetchone()[0]
        acts = self.db.execute("SELECT action, COUNT(*) AS n FROM activity WHERE ts>=? GROUP BY action ORDER BY n DESC",
                               (since,)).fetchall()
        top = self.db.execute("""SELECT c.chat_id, c.name, c.phone, c.role, c.last_seen, COUNT(a.id) AS n
                                 FROM activity a JOIN clients c ON c.chat_id=a.chat_id WHERE a.ts>=?
                                 GROUP BY c.chat_id ORDER BY n DESC LIMIT 10""", (since,)).fetchall()
        sleeping = self.db.execute("""SELECT chat_id, name, phone, last_seen FROM clients WHERE status='active'
                                      AND left_at IS NULL AND (last_seen IS NULL OR last_seen<?)
                                      ORDER BY last_seen""", (since30,)).fetchall()
        daily = self.db.execute("""SELECT substr(ts,1,10) AS d, COUNT(DISTINCT chat_id) AS u, COUNT(*) AS n FROM activity
                                   WHERE ts>=? GROUP BY d ORDER BY d""", (since,)).fetchall()
        return {
            "days": days,
            "active_users": one("SELECT COUNT(*) FROM clients WHERE status='active' AND left_at IS NULL"),
            "left": one("SELECT COUNT(*) FROM clients WHERE left_at IS NOT NULL"),
            "today": one("SELECT COUNT(DISTINCT chat_id) FROM activity WHERE ts>=?", today),
            "week": one("SELECT COUNT(DISTINCT chat_id) FROM activity WHERE ts>=?", since),
            "month": one("SELECT COUNT(DISTINCT chat_id) FROM activity WHERE ts>=?", since30),
            "actions": [{"action": r["action"], "label": self.action_name(r["action"]), "n": r["n"]} for r in acts],
            "top": [{"id": r["chat_id"], "name": r["name"] or r["phone"] or "", "phone": r["phone"] or "",
                     "role": r["role"], "n": r["n"], "last_seen": (r["last_seen"] or "")[:16].replace("T", " ")} for r in top],
            "sleeping": [{"id": r["chat_id"], "name": r["name"] or r["phone"] or "", "phone": r["phone"] or "",
                          "last_seen": (r["last_seen"] or "")[:16].replace("T", " ")} for r in sleeping],
            "daily": [{"date": r["d"], "users": r["u"], "n": r["n"]} for r in daily],
        }

    def stats_text(self):
        s = self.stats(7)
        t = ("<b>📈 Faollik statistikasi</b>\n\n👥 Faol foydalanuvchilar: <b>%d</b>\n"
             "Botga kirgan: bugun <b>%d</b> · 7 kunda <b>%d</b> · 30 kunda <b>%d</b>") % (
            s["active_users"], s["today"], s["week"], s["month"])
        if s["left"]:
            t += "\n🚫 Botni tark etgan: <b>%d</b>" % s["left"]
        if s["actions"]:
            t += "\n\n<b>Bo'limlar (7 kun):</b>\n" + "\n".join("%s · %d" % (a["label"], a["n"]) for a in s["actions"][:8])
        if s["top"]:
            t += "\n\n<b>Eng faollar (7 kun):</b>\n" + "\n".join(
                "%d. %s · %d marta · /c%d" % (i, esc(self.short(u["name"], 24)), u["n"], u["id"]) for i, u in enumerate(s["top"], 1))
        if s["sleeping"]:
            t += "\n\n😴 30 kundan beri kirmagan: <b>%d</b> ta" % len(s["sleeping"])
        return t

    # ================= admin: mijoz ko'rinishi =================
    def view_as(self, cid, target, kind=None, mid=None):
        """Admin mijozning panelini u ko'rgandek ko'radi (mijozga hech narsa yuborilmaydi)."""
        c = self.client(target)
        if not c:
            return self.users_menu(cid, mid, "Mijoz topilmadi.")
        if c["role"] != "staff" and not self.can_use(c):
            return self.client_card(cid, target, mid, "Bu mijozga firma biriktirilmagan: u hali hech narsa ko'rmaydi.")
        head = "👁 <b>Mijoz ko'rinishi:</b> %s\n<i>(mijozga hech narsa yuborilmaydi)</i>\n━━━━━━━━━━\n" % esc(c["name"] or c["phone"])
        cr = ("👥 Foydalanuvchilar", esc(c["name"] or c["phone"] or str(target)), "👁 Ko'rinish")
        if c["role"] == "filial":  # filial xodimi: faqat ko'chirish bo'limlari
            f = norm(c["filial"])
            if kind is None:
                text = self.filial_panel(c, save_sig=False)[0]
            elif kind in KOCH_CAT or kind in ("kz", "ko", "kin", "kday"):
                text = self.filial_view(f, kind)
            else:
                text = "Bu bo'lim filial xodimiga ko'rinmaydi."
            kb = [[("%s %s" % (k["icon"], k["title"]), "v2:%d:%s" % (target, k["key"])) for k in KOCH_CATS],
                  [("📥 Keladigan", "v2:%d:kin" % target), ("📋 Bugungi holat", "v2:%d:kday" % target)],
                  [("🏠 Mijoz menyusi", "vw:%d" % target)]]
            return self.a_screen(cid, mid, cr, head + text, kb, back="uc:%d" % target)
        if kind is None:
            text, _ = self.client_panel(c, save_sig=False)
        elif kind == "acc":
            tags, label = self.eff(c)
            if c["role"] == "staff":
                text = self.hisob_text(tags, None, label)
            else:
                text = self.hisob_text(tags, self.accounts_of(target), auto=False)
            text = text or "💰 Hisob biriktirilmagan: mijoz «Hisob» ni bossa, tasdiqlash so'rovi yaratiladi."
        elif kind == "stock":
            tags, label = self.eff(c)
            text = self.sq_text(tags, label if tags is not None else "Hamma firmalar")[0]
        else:
            tags, label = self.eff(c)
            text = self.section_text(kind, tags, label)
        kb = [[("⏰ Srok", "v2:%d:srok" % target), ("📦 30+", "v2:%d:old" % target), ("🧾 Buyurtma", "v2:%d:need" % target)],
              [("📊 Sotuv va qoldiq", "v2:%d:stock" % target), ("💰 Hisob", "v2:%d:acc" % target), ("🏠 Mijoz menyusi", "vw:%d" % target)]]
        return self.a_screen(cid, mid, cr, head + text, kb, back="uc:%d" % target)

    # ---------------- Excelni yangilash ("Yangilash") ----------------
    def refresh_paths(self):
        path = self.data.find_path() or self.data.file_pattern()
        status = os.path.join(os.path.dirname(path), self.cfg.get("refresh", "status_file", fallback="refresh_status.txt"))
        vbs = self.cfg.get("refresh", "vbs", fallback="refresh.vbs")
        if not os.path.isabs(vbs):
            # O'rnatilgan dasturda faqat dastur papkasidan (Program Files — faqat admin yoza oladi). Ma'lumot papkasiga
            # (%ProgramData%\TimeSaver) har bir foydalanuvchi yoza oladi: u yerga qo'yilgan skript bot nomidan ishga tushardi.
            # Manba koddan ishlaganda (eski usul) — avval ma'lumot papkasidagi moslashtirilgan nusxa.
            dirs = (APP_DIR,) if FROZEN else (BASE_DIR, APP_DIR)
            vbs = next((p for p in (os.path.join(d, vbs) for d in dirs) if os.path.exists(p)), os.path.join(APP_DIR, vbs))
        return path, status, vbs

    def start_refresh(self, cid=None, auto=False, mid=None, hourly=False):
        """cid=None: Telegramsiz (server panelidan) ishga tushirish; holat self.rstate da.
        auto=True: kunlik avto-yangilash — muvaffaqiyatsiz bo'lsa adminlarga sababi bilan xabar boradi.
        hourly=True (auto bilan): har soatlik avto-yangilash — chatga hech narsa yozilmaydi, faqat jurnal.
        Telegramdan (cid): jarayon va natija panel xabarining o'zida (mid)."""
        with self.rlock:
            if self.refreshing:
                if cid:
                    self.panel_note(cid, mid, "🔄 Yangilanish allaqachon ketmoqda, iltimos kuting.")
                return False
            self.refreshing = True
            self.rstate = {"running": True, "steps": {}, "ok": None, "msg": "", "t0": time.time(), "t1": None}
        self.set_setting("last_refresh_start", now().isoformat(timespec="seconds"))  # har soatlik oraliq shundan sanaladi
        self.bg(self._refresh_job, cid, auto, mid, hourly, name="refresh")
        return True

    # ---------------- avtomatik yangilash ----------------
    @property
    def auto_refresh_time(self):
        """Har kuni avto-yangilash vaqti (standart 08:30). Bo'sh = o'chirilgan."""
        v = self.setting("auto_refresh_time", "08:30").strip()
        m = re.match(r"^(\d{1,2}):(\d{2})$", v)
        if not m or int(m.group(1)) > 23 or int(m.group(2)) > 59:
            return None
        return datetime.time(int(m.group(1)), int(m.group(2)))

    def auto_refresh_due(self):
        t = self.auto_refresh_time
        n = now()
        today = n.date().isoformat()
        if t is None or n.time() < t or self.refreshing:
            return False
        if self.setting("last_refresh_ok_date") == today or self.setting("last_auto_refresh_try") == today:
            return False  # bugun muvaffaqiyatli yangilangan yoki avto-urinish bo'lgan (xato bo'lsa ham qayta urinmaymiz)
        self.data.reload()
        w = self.data.when()
        return not (w and w.date() == n.date())  # Excel ichidagi vaqt: bugun (qo'lda ham) yangilangan bo'lsa — kerak emas

    def start_auto_refresh(self):
        self.set_setting("last_auto_refresh_try", now().date().isoformat())
        logging.info("Avto-yangilash boshlandi")
        return self.start_refresh(None, auto=True)

    # ---- har soatlik avto-yangilash (panel sozlamasi: auto_refresh_every 0/30/60/120 daqiqa, auto_refresh_from..to) ----
    HOURLY_EVERY = (30, 60, 120)

    def hourly_conf(self):
        """(oraliq daqiqa yoki 0, boshlanish time, tugash time)."""
        try:
            every = int(self.setting("auto_refresh_every", "0") or 0)
        except ValueError:
            every = 0
        fr = self.parse_hm(self.setting("auto_refresh_from", "08:00")) or datetime.time(8, 0)
        to = self.parse_hm(self.setting("auto_refresh_to", "20:00")) or datetime.time(20, 0)
        return (every if every in self.HOURLY_EVERY else 0), fr, to

    def hourly_refresh_due(self):
        """Yoqilgan, ish vaqti ichida, yangilash ketmayapti va oxirgi yangilash (har qanday: qo'lda, ertalabki, har soatlik)
        boshlanganidan oraliq o'tgan. Ertalabki avto-yangilash vaqti kelmaguncha (bugun hali urinilmagan bo'lsa) kutadi —
        ertalabki yangilash va kunlik topshiriq vaqti o'zgarmaydi."""
        every, fr, to = self.hourly_conf()
        n = now()
        if not every or self.refreshing or not (fr <= n.time() <= to):
            return False
        t = self.auto_refresh_time
        if t is not None and n.time() < t and self.setting("last_auto_refresh_try") != n.date().isoformat():
            return False
        last = self._dt(self.setting("last_refresh_start") or "")
        return last is None or (n - last).total_seconds() >= every * 60

    def start_hourly_refresh(self):
        self.set_setting("last_hourly_refresh", now().isoformat(timespec="seconds"))
        logging.info("Har soatlik avto-yangilash boshlandi")
        return self.start_refresh(None, auto=True, hourly=True)

    @staticmethod
    def progress_text(steps, running=True, extra=""):
        lines, waiting_marked = [], False
        for name in SHEET_ORDER:
            st = steps.get(name)
            if st and st[0] == "ok":
                lines.append("✅ %s <i>%s</i>" % (name, ("· %s qator" % st[1]) if st[1] else ""))
            elif st and st[0] == "err":
                lines.append("❌ %s <i>%s</i>" % (name, esc(st[1])[:80]))
            elif running and not waiting_marked:
                lines.append("⏳ %s <i>jarayonda…</i>" % name)
                waiting_marked = True
            else:
                lines.append("◻️ %s <i>kutmoqda</i>" % name)
        head = "<b>🔄 Ma'lumot yangilanmoqda</b>" if running else "<b>🔄 Yangilash</b>"
        done = sum(1 for s in steps.values() if s[0] == "ok")
        bar = "▰" * done + "▱" * (len(SHEET_ORDER) - done)
        return "%s\n%s\n\n%s\n\n%s" % (head, bar, "\n".join(lines), extra)

    def _refresh_job(self, cid, auto=False, mid=None, hourly=False):
        rs = self.rstate
        crumb = self.crumb("🔄 Yangilash") + "\n\n"
        home = inline_kb([[("🏠 Menyu", "a:home")]])

        def say(text, kb=None):  # panel xabarining o'zi (jarayonda tugmasiz — ikki marta ishga tushirib bo'lmaydi)
            if not cid:
                return None
            self.panel_show(cid, mid, crumb + text, kb)
            return {"message_id": self.panel_mid(cid)}

        def early(msg):
            rs.update(running=False, ok=False, msg=msg, t1=time.time())
            say("❌ " + esc(msg), home)

        try:
            path, status, vbs = self.refresh_paths()
            timeout = int(self.cfg.get("refresh", "timeout_sec", fallback="900"))
            if not os.path.exists(path):
                return early("Excel fayl topilmadi: %s" % path)
            try:
                with open(path, "ab"):
                    pass
            except OSError:
                return early("Excel fayl band (kimdir ochib qo'ygan). Faylni yoping va qaytadan «Yangilash» ni bosing.")
            try:
                os.remove(status)
            except OSError:
                pass
            extra = "Taxminan 1–2 daqiqa. Shu vaqt ichida qayta bosmang."
            msg = say(self.progress_text({}, True, extra))
            mid = msg.get("message_id") if isinstance(msg, dict) else None
            result = {}

            def work():
                try:
                    if self.sale_days and self.runner is default_runner:
                        result["rc"], result["out"] = self.runner(vbs, path, timeout, days=self.sale_days)
                    elif self.sale_days and getattr(self.runner, "takes_days", False):
                        result["rc"], result["out"] = self.runner(vbs, path, timeout, days=self.sale_days)
                    else:
                        result["rc"], result["out"] = self.runner(vbs, path, timeout)
                except Exception as e:  # noqa
                    result["rc"], result["out"] = -1, str(e)

            t = threading.Thread(target=work, daemon=True)
            t.start()
            last, t0 = None, time.time()
            while t.is_alive():
                steps, done, fail = parse_status(status)
                rs["steps"] = dict(steps)
                txt = self.progress_text(steps, True, extra)
                if txt != last and mid:
                    try:
                        self.tg.edit(cid, mid, crumb + txt)
                        last = txt
                    except Exception:  # noqa
                        pass
                t.join(2)
                if time.time() - t0 > timeout + 30:
                    break
            steps, done, fail = parse_status(status)
            rs["steps"] = dict(steps)
            ok = result.get("rc") == 0 and done and not fail and all(s[0] == "ok" for s in steps.values()) and steps
            self.data.reload(force=True)
            if ok:
                final = "<b>✅ Yangilandi</b>\n%s\n\n🕒 %s · davomiyligi %d soniya" % (
                    self.progress_text(steps, False, "").split("\n\n", 1)[1].strip(), dt_text(self.data.when()), time.time() - t0)
                rs.update(ok=True, msg="Yangilandi · %d soniya" % (time.time() - t0))
            else:
                err = fail or result.get("out") or "Excel javob bermadi"
                bad = ["%s: %s" % (k, v[1]) for k, v in steps.items() if v[0] == "err"]
                final = "<b>❌ Yangilash to'liq bajarilmadi</b>\n%s\n\n%s" % (
                    esc("; ".join(bad) or err)[:500],
                    "Boshqa varaqlar eski ma'lumot bilan qoldi. Server kompyuterida Excel va SQL ulanishini tekshiring.")
                rs.update(ok=False, msg=("; ".join(bad) or err)[:500])
            if cid:
                try:
                    if not mid:
                        raise TGError("panel yo'q")
                    self.tg.edit(cid, mid, crumb + final, home)
                    self._panel_keep(cid, mid, crumb + final, home)
                except Exception:  # noqa
                    self.panel_resend(cid, crumb + final, home)
        except Exception as e:  # noqa
            logging.exception("Yangilash xatosi")
            rs.update(ok=False, msg="Yangilash xatosi: %s" % e)
            try:
                say("❌ Yangilash xatosi: %s" % esc(e), home)
            except Exception:  # noqa
                pass
        finally:
            rs.update(running=False, t1=time.time())
            try:
                if rs.get("ok"):
                    self.set_setting("last_refresh_ok_date", now().date().isoformat())
                elif auto and hourly:  # har soatlik: chatga yozilmaydi (keyingi soatda yana urinadi)
                    logging.warning("Har soatlik avto-yangilash bajarilmadi: %s", rs.get("msg"))
                elif auto:  # avto-yangilash muvaffaqiyatsiz: adminlarga sababi (avto-yuborish eski ma'lumotni stale_hours bilan to'xtatadi)
                    logging.warning("Avto-yangilash bajarilmadi: %s", rs.get("msg"))
                    for a in self.admins:
                        try:
                            self.tg.send(a, "❌ <b>Avto-yangilash bajarilmadi</b> (%s)\nSabab: %s\n\nQo'lda «🔄 Yangilash» ni bosib ko'ring." % (
                                now().strftime("%H:%M"), esc(rs.get("msg") or "noma'lum")))
                        except Exception:  # noqa
                            logging.exception("Adminga xabar yuborib bo'lmadi: %s", a)
            finally:
                if not rs.get("ok"):
                    self.refreshing = False
        if rs.get("ok"):
            # refreshing yangilashdan keyingi ishlar tugaguncha ushlab turiladi: ustma-ust yangilash / asosiy sikl ikkinchi marta
            # aniqlamaydi (ogohlantirish ikki marta ketmaydi)
            try:
                try:  # avval ko'chirishni qoldiq o'zgarishidan aniqlash (AVTO-ANIQLASH.md) — eski ro'yxat bo'yicha
                    self.koch_auto_detect()
                except Exception:  # noqa
                    logging.exception("Ko'chirishni avtomatik aniqlab bo'lmadi")
                try:  # keyin bugungi ko'chirish topshirig'i (vaqti kelgan bo'lsa) yoki qayta shakllantirish
                    self.koch_try_snapshot(in_refresh=True, detected=True)
                except Exception:  # noqa
                    logging.exception("Ko'chirish topshirig'ini yaratib bo'lmadi")
                if hourly:
                    logging.info("Har soatlik avto-yangilash tugadi: %s", rs.get("msg"))
            finally:
                self.refreshing = False

    # ---------------- avtomatik srok yuborish ----------------
    def auto_due(self):
        if not self.auto_on:
            return False
        n = now()
        return n.time() >= self.auto_time and self.setting("last_auto_date") != n.date().isoformat()

    WEEKDAYS = ["Dushanba", "Seshanba", "Chorshanba", "Payshanba", "Juma", "Shanba", "Yakshanba"]

    @property
    def weekly_kinds(self):
        """Haftalik avto-yuborish: qaysi hisobotlar yoqilgan (30+ kun, buyurtma)."""
        return [k for k in ("old", "need") if self.setting("weekly_%s" % k, "0") == "1"]

    @property
    def weekly_day(self):
        return int(self.setting("weekly_day", "0"))  # 0 = dushanba

    def weekly_due(self):
        n = now()
        return (bool(self.weekly_kinds) and n.weekday() == self.weekly_day and n.time() >= self.auto_time
                and self.setting("last_weekly_date") != n.date().isoformat())

    def _auto_send(self, kinds, title, last_key):
        """Avto-yuborish (kunlik srok yoki haftalik): ma'lumot eski bo'lsa to'xtaydi va adminga xabar beradi."""
        self.data.reload()
        stale = float(self.cfg.get("auto", "stale_hours", fallback="72"))
        if self.data.age_hours() > stale or not self.data.path:
            self.set_setting(last_key, "%s — to'xtatildi (ma'lumot eski)" % dt_text(now()))
            for a in self.admins:
                try:
                    self.tg.send(a, "⚠ %s to'xtatildi: ma'lumot %s da yangilangan (%d soatdan eski). «Yangilash» ni bosing, keyin Xabar yuborish."
                                 % (title, dt_text(self.data.when()), int(stale)))
                except Exception:  # noqa
                    pass
            return
        lines = []
        for kind in kinds:
            cnt = {"ok": 0, "skip": 0, "gone": 0, "fail": 0}
            for c in self.audience("all"):
                cnt[self.deliver(c, kind)] += 1
                self.sleep(0.05)
            lines.append("%s: ✅ %d · ➖ %d · ❌ %d%s" % (self.KINDS[kind], cnt["ok"], cnt["skip"], cnt["fail"],
                                                      (" · 🚫 %d" % cnt["gone"]) if cnt["gone"] else ""))
        self.set_setting(last_key, "%s — %s" % (dt_text(now()), "; ".join(lines)))
        for a in self.admins:
            try:
                self.tg.send(a, "<b>%s</b>\n%s\n\n<i>✅ yuborildi · ➖ ma'lumot yo'q · ❌ xato</i>" % (title, "\n".join(lines)))
            except Exception:  # noqa
                pass

    def auto_srok(self):
        self.set_setting("last_auto_date", now().date().isoformat())
        self._auto_send(["srok"], "⏰ Srok avto-yuborish (%d kun)" % self.srok_days, "last_auto")

    def auto_weekly(self):
        self.set_setting("last_weekly_date", now().date().isoformat())
        self._auto_send(self.weekly_kinds, "📅 Haftalik avto-yuborish", "last_weekly")

    # ---------------- bot.db zaxira nusxasi ----------------
    BACKUP_KEEP = 14

    def backup_db(self):
        """bot.db ni backups/ papkasiga nusxalaydi (SQLite backup: bot ishlab turganda ham xavfsiz). Oxirgi 14 kun saqlanadi."""
        folder = os.path.join(os.path.dirname(os.path.abspath(self.db_path)), "backups")
        os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, "bot_%s.db" % now().strftime("%Y-%m-%d"))
        dst = sqlite3.connect(path)
        try:
            with self.dblock:
                self.db.backup(dst)
        finally:
            dst.close()
        for old in sorted(glob.glob(os.path.join(folder, "bot_*.db")))[:-self.BACKUP_KEEP]:
            try:
                os.remove(old)
            except OSError:
                pass
        with self.dblock:  # 180 kundan eski faollik yozuvlari kerak emas
            self.db.execute("DELETE FROM activity WHERE ts < ?", ((now() - datetime.timedelta(days=180)).isoformat(),))
            self.db.commit()
        self.set_setting("last_backup", dt_text(now()))
        self.set_setting("last_backup_date", now().date().isoformat())
        logging.info("bot.db zaxira nusxasi: %s", path)
        return path

    def backup_due(self):
        return self.db_path != ":memory:" and self.setting("last_backup_date") != now().date().isoformat()

    # ---------------- asosiy sikl ----------------
    # ---------------- kunlik ogohlantirishlar (admin) ----------------
    # alert_daily_report — ko'chirish kunlik hisoboti (ODDIY muddatidan keyin, alohida xabar); qolganlari — kunlik bitta xabarda
    ALERTS = (("alert_expired", "⛔ Muddati o'tgan, qoldig'i bor partiyalar"), ("alert_loss_price", "💸 Sotuv narxi ≤ kirim narxi"),
              ("alert_bad_dates", "📅 Xato muddatlar"), ("alert_internal", "🏦 Ichki hisoblar"),
              ("alert_daily_report", "📊 Ko'chirish: kunlik hisobot"))

    def alert_on(self, key):
        return self.setting(key, "1") == "1"

    def admin_alerts(self):
        """(matn bo'limlari, Excel varaqlari). Faqat yoqilgan va ma'lumoti bor ogohlantirishlar."""
        d = self.data
        d.reload()
        today = datetime.date.today()
        parts, sheets = [], []
        line_srok = lambda i, r: "%d. %s · %s · %s · %s dona · %s" % (
            i, esc(self.short(r["name"], 32)), esc(self.short(norm(r["filial"]), 18)),
            r["date"].strftime("%d.%m.%Y") if r["date"] else "—", qty(r["ost"]), money(r["ost"] * r["prih"]))
        if self.alert_on("alert_expired"):
            rows = sorted([r for r in d.srok if r["ost"] > 0 and ((r["date"] - today).days if r["date"] else r["kun"]) < 0],
                          key=lambda r: -r["ost"] * r["prih"])
            if rows:
                parts.append("⛔ <b>Muddati o'tgan, qoldig'i bor</b>: <b>%d</b> partiya · %s so'm\n%s" % (
                    len(rows), money(sum(r["ost"] * r["prih"] for r in rows)),
                    "\n".join(line_srok(i, r) for i, r in enumerate(rows[:10], 1))))
                sheets.append(("Muddati o'tgan", ["Filial", "Tovar", "Ishlab chiqaruvchi", "Firma", "Srok", "Qoldiq", "Kirim narxi",
                                                  "Summa"],
                               [[r["filial"], r["name"], r["producer"], r["firm"], r["date"], r["ost"], r["prih"], r["ost"] * r["prih"]]
                                for r in rows], [22, 46, 26, 22, 12, 10, 13, 15], (4,), (5, 6, 7)))
        if self.alert_on("alert_loss_price"):
            rows = sorted([r for r in d.narx if r["prih"] > 0 and r["roz"] <= r["prih"]], key=lambda r: r["roz"] - r["prih"])
            if rows:
                parts.append("💸 <b>Sotuv narxi ≤ kirim narxi</b>: <b>%d</b> ta\n%s" % (len(rows), "\n".join(
                    "%d. %s · kirim %s · sotuv %s" % (i, esc(self.short(r["name"], 34)), money(r["prih"]), money(r["roz"]))
                    for i, r in enumerate(rows[:10], 1))))
                sheets.append(("Narx", ["Tovar", "Ishlab chiqaruvchi", "Firma", "Kirim narxi", "Sotuv narxi", "Farq"],
                               [[r["name"], r["producer"], r["firm"], r["prih"], r["roz"], r["roz"] - r["prih"]] for r in rows],
                               [46, 26, 22, 13, 13, 13], (), (3, 4, 5)))
        if self.alert_on("alert_bad_dates"):
            rows = sorted([r for r in d.srok if r["date"] and r["date"].year > today.year + 10], key=lambda r: r["date"], reverse=True)
            if rows:
                parts.append("📅 <b>Xato muddatlar</b> (yil > %d): <b>%d</b> ta — bazada to'g'rilash kerak\n%s" % (
                    today.year + 10, len(rows), "\n".join(line_srok(i, r) for i, r in enumerate(rows[:10], 1))))
                sheets.append(("Xato muddat", ["Filial", "Tovar", "Ishlab chiqaruvchi", "Firma", "Srok", "Qoldiq"],
                               [[r["filial"], r["name"], r["producer"], r["firm"], r["date"], r["ost"]] for r in rows],
                               [22, 46, 26, 22, 12, 10], (4,), (5,)))
        if self.alert_on("alert_internal"):
            rows = [r for r in d.hisob if self.is_internal(r["org"])]
            if rows:
                parts.append("🏦 <b>Ichki hisoblar</b> (umumiy qarzga kirmaydi): <b>%d</b> ta\n%s" % (len(rows), "\n".join(
                    "• %s · %s" % (esc(self.short(r["org"], 34)), self.smoney(r["debit"] - r["credit"])) for r in rows[:10])))
                sheets.append(("Ichki hisoblar", ["Tashkilot", "Debet", "Kredit", "Qoldiq"],
                               [[r["org"], r["debit"], r["credit"], r["debit"] - r["credit"]] for r in rows], [46, 18, 18, 18], (),
                               (1, 2, 3)))
        return parts, sheets

    def daily_admin_due(self):
        n = now()
        return n.time() >= self.auto_time and self.setting("last_admin_daily") != n.date().isoformat()

    def daily_admin(self):
        """Har kuni avto-yuborish vaqtida adminga BITTA xabar + Excel: ogohlantirishlar va qabul qilinmagan ko'chirishlar."""
        self.set_setting("last_admin_daily", now().date().isoformat())
        parts, sheets = self.admin_alerts()
        if not parts:
            return
        head = "🔔 <b>Kunlik ogohlantirishlar</b> · %s\n🕒 Ma'lumot: %s" % (now().strftime("%d.%m.%Y"), dt_text(self.data.when()))
        if self.data.age_hours() > float(self.cfg.get("auto", "stale_hours", fallback="72")):
            head += "\n⚠ Ma'lumot eskirgan, «Yangilash» ni bosing."
        text = cut_html(head + "\n\n" + "\n\n".join(parts), 4000)
        content = xlsx_multi(sheets) if sheets else None
        for a in self.admins:
            try:
                self.tg.send(a, text, inline_kb([[("🔁 Ko'chirishlar", "a:tr"), ("⚙️ Sozlamalar", "a:set")], [("🏠 Menyu", "a:home")]]))
                if content:
                    self.tg.send_document(a, "Ogohlantirishlar_%s.xlsx" % now().date().isoformat(), content,
                                          "🔔 Kunlik ogohlantirishlar (to'liq ro'yxat)")
            except Exception:  # noqa
                logging.exception("Adminga kunlik xabar yuborib bo'lmadi: %s", a)

    def run_scheduled(self):
        """Kunlik zaxira, kunlik srok va haftalik avto-yuborish. Yuborishlar fonda: bot shu vaqtda ham javob beradi."""
        if self.backup_due():
            try:
                self.backup_db()
            except Exception:  # noqa
                logging.exception("bot.db zaxira nusxasini olib bo'lmadi")
                self.set_setting("last_backup_date", now().date().isoformat())  # bugun qayta urinmaymiz
        try:  # kunlik ko'chirish topshirig'i (koch_task_time dan keyingi birinchi muvaffaqiyatli yangilashdan)
            self.koch_try_snapshot()
        except Exception:  # noqa
            logging.exception("Ko'chirish topshirig'ini yaratib bo'lmadi")
        if self.auto_running or self.maint_on:  # texnik rejimda avto-yuborish, avto-yangilash va eslatmalar kutadi
            return
        if self.auto_refresh_due():
            self.start_auto_refresh()
            return
        if self.hourly_refresh_due():
            self.start_hourly_refresh()
            return
        if self.refreshing:  # avto-yuborish yangilash tugashini kutadi
            return
        jobs = []
        if self.auto_due():
            jobs.append(self.auto_srok)
        if self.weekly_due():
            jobs.append(self.auto_weekly)
        if self.daily_admin_due():
            jobs.append(self.daily_admin)
        if self.koch_on and self.koch_reminders_due():
            jobs.append(self.koch_reminders)
        if jobs:
            self.auto_running = True

            def run():
                try:
                    for j in jobs:
                        j()
                finally:
                    self.auto_running = False
            self.bg(run, name="auto")

    def poll_once(self, offset, timeout=30):
        self.last_poll = time.time()
        ups = self.tg.call("getUpdates", {"offset": offset, "timeout": timeout,
                                          "allowed_updates": ["message", "callback_query"]}, timeout=timeout + 15)
        self.last_poll = time.time()
        for u in ups:
            offset = u["update_id"] + 1
            try:
                self.handle(u)
            except Exception:
                logging.exception("Yangilanishni qayta ishlashda xato: %s", u.get("update_id"))
        return offset

    def run_forever(self):
        offset = None
        logging.info("Bot ishga tushdi. Adminlar: %s", sorted(self.admins))
        while True:
            try:
                if not self.token_ok():
                    time.sleep(3)  # token server panelidan kiritilishini kutadi
                    continue
                self.run_scheduled()
                offset = self.poll_once(offset, timeout=20)
            except TGError as e:
                if e.code == 409:  # shu token bilan boshqa nusxa ham getUpdates qilyapti: tugmalar ikkiga bo'linib, bot qotadi
                    self.conflict_at = time.time()
                    if self.conflict_at - self._conflict_logged > 600:
                        self._conflict_logged = self.conflict_at
                        logging.error("DIQQAT: shu bot tokeni bilan boshqa joyda ham bot ishlayapti (409 Conflict). Bot sekin "
                                      "ishlaydi va xabarlarning bir qismi boshqa nusxaga ketadi. Ikkinchi nusxani to'xtating "
                                      "yoki har bir aptekaga alohida bot (@BotFather) yarating.")
                    time.sleep(5)
                    continue
                logging.warning("Telegramga ulanishda xato: %s. 5 soniyadan keyin qayta uriniladi.", e)
                time.sleep(5)
            except Exception as e:  # noqa
                logging.warning("Telegramga ulanishda xato: %s. 5 soniyadan keyin qayta uriniladi.", e)
                time.sleep(5)


def setup_logging():
    h = logging.handlers.RotatingFileHandler(os.path.join(BASE_DIR, "bot.log"), maxBytes=1_000_000, backupCount=3, encoding="utf-8")
    handlers = [h] + ([logging.StreamHandler(sys.stdout)] if sys.stdout else [])  # pythonw da konsol yo'q
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s", handlers=handlers)


def ensure_config():
    """Birinchi ishga tushish: ma'lumot papkasida config.ini bo'lmasa, namunadan yaratiladi (token va Excel — sozlash ustasida)."""
    path = os.path.join(BASE_DIR, "config.ini")
    sample = os.path.join(APP_DIR, "config.example.ini")
    if not os.path.exists(path) and os.path.exists(sample):
        shutil.copyfile(sample, path)
        logging.info("Yangi o'rnatish: config.ini namunadan yaratildi (%s)", path)
    return path


def create_bot():
    """Logni sozlaydi, config ni o'qiydi, botni yaratadi va server panelini ishga tushiradi."""
    setup_logging()
    ensure_config()
    try:
        cfg = load_config()
    except FileNotFoundError:
        raise SystemExit("config.ini topilmadi: %s\nEski bot papkasidagi config.ini ni shu papkaga nusxalang." % os.path.join(BASE_DIR, "config.ini"))
    bot = Bot(cfg)
    if not bot.token_ok():
        logging.warning("Bot tokeni kiritilmagan. Server panelida Sozlamalar bo'limida kiriting.")
    if not bot.admins:
        logging.warning("admin_ids bo'sh. Server panelida Sozlamalar bo'limida admin ID qo'shing.")
    try:
        import timesaver_panel
        timesaver_panel.start(bot, cfg)
    except Exception:  # noqa
        logging.exception("Server paneli ishga tushmadi (bot ishlashda davom etadi)")
    return bot


def main():
    try:
        bot = create_bot()
    except SystemExit as e:
        print(e)
        sys.exit(1)
    bot.run_forever()


if __name__ == "__main__":
    main()
