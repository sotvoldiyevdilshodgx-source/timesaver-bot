# -*- coding: utf-8 -*-
"""O'zbek lotin → kirill (foydalanuvchiga ko'rinadigan matn uchun).

TimeSaver'da hamma interfeys o'zbek kirill alifbosida. Koddagi eski lotin matnlar Telegram va panelga chiqishda
shu funksiya orqali kirillga o'tadi (bitta chegara: TG.call, panel JSON «msg»/«error», Excel sarlavhalari).

Himoyalanadi (o'zgarmaydi):
- HTML teglari va &entity; lar;
- ma'lumot: DATA_OPEN … DATA_CLOSE orasidagi matn (esc() shunday o'raydi: dori/foydalanuvchi nomlari, admin yozgan matn);
  belgilar natijadan olib tashlanadi;
- havolalar, @username, /buyruq, e-pochta;
- PROTECT ro'yxatidagi so'zlar (TimeSaver, Excel, SQL ...).
"""
import re

DATA_OPEN, DATA_CLOSE = chr(0xE000), chr(0xE001)  # ko'rinmas belgilar: ma'lumot chegarasi
KB_BACK = {}  # kirill tugma matni -> koddagi asl matn (TG.call to'ldiradi)

PROTECT = {
    "timesaver", "excel", "abc", "sql", "id", "telegram", "fom", "windows", "botfather", "html", "ok", "pdf",
    "xlsx", "xlsm", "xls", "vba", "url", "api", "http", "https", "www", "ctrl", "enter", "esc", "tab", "f5",
    "a", "b", "c", "webview", "webview2", "pywebview", "edge", "chrome", "android", "ios", "json", "csv", "utf",
    "admin_ids", "token", "cover", "status", "timesaver",
}

_APOS = "'ʼ’‘`"
_VOW = set("aeiouAEIOU")
_SINGLE = {
    "a": "а", "b": "б", "d": "д", "f": "ф", "g": "г", "h": "ҳ", "i": "и", "j": "ж", "k": "к", "l": "л", "m": "м",
    "n": "н", "o": "о", "p": "п", "q": "қ", "r": "р", "s": "с", "t": "т", "u": "у", "v": "в", "x": "х", "y": "й",
    "z": "з", "c": "ц", "w": "в",
}
# himoyalanadigan bo'laklar: ma'lumot, teg, entity, havola, @nik, /buyruq, e-pochta
_SEG = re.compile(
    "(" + re.escape(DATA_OPEN) + ".*?" + re.escape(DATA_CLOSE) + r"|<[^>]*>|&[#\w]+;|https?://\S+|www\.\S+"
    r"|[\w.+-]+@[\w-]+\.[\w.]+|(?<![\w/])@\w+|(?<![\w/:])/[a-z_]+\b"
    # Windows yo'li: C:\ProgramData\BazaReport\db.ini, \\server\share (probelli qism keyingi «\» gacha davom etadi)
    r"|(?<![\w])[A-Za-z]:\\[^\s<>\"']*(?: [^\s<>\"'\\]+\\[^\s<>\"']*)*|\\\\[^\s<>\"']+"
    # fayl nomi: baza_yangi.xlsm, config.ini, refresh.vbs
    r"|\b[\w-]+\.(?:xlsm|xlsx|xls|ini|vbs|exe|log|db|txt|csv|json|html|py)\b)", re.S)
_WORD = re.compile(r"[A-Za-z][A-Za-z'ʼ’‘`]*(?:[OoGg]['ʼ’‘`]|[A-Za-z])|[A-Za-z]")  # oxiri o' / g' bo'lishi mumkin


def _word(w):
    low = w.lower()
    base = re.split("[" + _APOS + "]", low, 1)[0]
    if low in PROTECT:
        return w
    if base in PROTECT and base != low:  # Excel'da → Excel'да
        cut = len(base)
        return w[:cut] + w[cut] + _word(w[cut + 1:]) if len(w) > cut + 1 else w
    out = []
    i, n = 0, len(w)
    while i < n:
        ch = w[i]
        lo = ch.lower()
        nxt = w[i + 1].lower() if i + 1 < n else ""
        nxt2 = w[i + 2] if i + 2 < n else ""
        if lo in ("o", "g") and nxt and nxt in _APOS:
            out.append(("ў" if lo == "o" else "ғ", ch.isupper()))
            i += 2
            continue
        if lo == "s" and nxt == "h":
            out.append(("ш", ch.isupper()))
            i += 2
            continue
        if lo == "c" and nxt == "h":
            out.append(("ч", ch.isupper()))
            i += 2
            continue
        if lo == "y" and nxt in ("o", "u", "a", "e") and not (nxt == "o" and nxt2 and nxt2 in _APOS):
            out.append(({"o": "ё", "u": "ю", "a": "я", "e": "е"}[nxt], ch.isupper()))
            i += 2
            continue
        if lo == "e":
            prev = w[i - 1] if i else ""
            out.append(("э" if i == 0 or prev in _VOW else "е", ch.isupper()))
            i += 1
            continue
        if ch in _APOS:
            out.append(("ъ", False))
            i += 1
            continue
        out.append((_SINGLE.get(lo, ch), ch.isupper()))
        i += 1
    upper_all = len(w) > 1 and w.upper() == w
    res = "".join(c.upper() if (up or upper_all) else c for c, up in out)
    return res


def _plain(s):
    return _WORD.sub(lambda m: _word(m.group(0)), s)


def lat2cyr(text):
    """Matndagi o'zbek lotin so'zlarni kirillga o'tkazadi; himoyalangan bo'laklar o'zgarmaydi; ma'lumot belgilari olinadi."""
    if not text or not isinstance(text, str):
        return text
    out, pos = [], 0
    for m in _SEG.finditer(text):
        out.append(_plain(text[pos:m.start()]))
        seg = m.group(0)
        if seg.startswith(DATA_OPEN):
            seg = seg[1:-1]
        out.append(seg)
        pos = m.end()
    out.append(_plain(text[pos:]))
    return strip_marks("".join(out))


def strip_marks(text):
    if not isinstance(text, str):
        return text
    return text.replace(DATA_OPEN, "").replace(DATA_CLOSE, "")


def mark(s):
    """Ma'lumotni o'raydi: lat2cyr uni o'zgartirmaydi (dori, firma, odam nomlari, admin yozgan matn)."""
    return DATA_OPEN + s + DATA_CLOSE


def deep(obj, keys=None):
    """dict/list ichidagi satrlarni o'tkazadi. keys berilsa — faqat shu kalitlardagi satrlar."""
    if isinstance(obj, str):
        return lat2cyr(obj) if keys is None else obj
    if isinstance(obj, list):
        return [deep(x, keys) for x in obj]
    if isinstance(obj, dict):
        return {k: (lat2cyr(v) if isinstance(v, str) and (keys is None or k in keys) else deep(v, keys))
                for k, v in obj.items()}
    return obj


def markup(m):
    """Telegram reply_markup: tugma matnlari kirillga (callback_data, url o'zgarmaydi)."""
    if not isinstance(m, dict):
        return m
    out = dict(m)
    for k in ("inline_keyboard", "keyboard"):
        if k in m:
            out[k] = [[(dict(b, text=lat2cyr(b["text"])) if isinstance(b, dict) and "text" in b
                        else lat2cyr(b) if isinstance(b, str) else b) for b in row] for row in m[k]]
    if "input_field_placeholder" in m:
        out["input_field_placeholder"] = lat2cyr(m["input_field_placeholder"])
    return out
