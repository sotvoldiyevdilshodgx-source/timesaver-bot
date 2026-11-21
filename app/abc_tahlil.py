# -*- coding: utf-8 -*-
"""A / B / C tahlili (pul bo'yicha): Srok, 30+ kun va ko'chirish uchun bitta qoida.

Qatorlar pul (qiymat) bo'yicha kamayish tartibida saralanadi. A — jami pulning birinchi 80% i (eng qimmatlari),
B — keyingi 15%, C — qolgan 5%. Qator chegarani kesib o'tsa, u boshlangan guruhga kiradi: shuning uchun eng qimmat
bitta dori yolg'iz jami pulning 90% i bo'lsa ham A da bo'ladi, A hech qachon bo'sh qolmaydi (agar pul bo'lsa).
Qiymati 0 (yoki manfiy) qatorlar C ga tushadi.
"""

LIMITS = (("A", 0.80), ("B", 0.95), ("C", 1.0))
CLASSES = ("A", "B", "C")
SHARE = {"A": 80, "B": 15, "C": 5}


def split(rows, value):
    """rows — ro'yxat, value(row) — pul. {"A": [...], "B": [...], "C": [...]} qaytaradi, har biri pul bo'yicha kamayish."""
    items = sorted(rows, key=lambda r: -max(0.0, float(value(r) or 0)))
    total = sum(max(0.0, float(value(r) or 0)) for r in items)
    out = {k: [] for k in CLASSES}
    if total <= 0:
        out["C"] = list(items)
        return out
    acc = 0.0
    for r in items:
        v = max(0.0, float(value(r) or 0))
        start = acc / total
        acc += v
        if v <= 0:
            out["C"].append(r)
            continue
        for cls, lim in LIMITS:
            if start < lim - 1e-12:
                out[cls].append(r)
                break
    return out


def classify(rows, value, key):
    """{key(row): "A"|"B"|"C"} — har qatorning sinfi (masalan, topshiriq id → sinf)."""
    res = {}
    for cls, lst in split(rows, value).items():
        for r in lst:
            res[key(r)] = cls
    return res


def summary(groups, value):
    """{"A": {"n", "sum", "share"}, ...} — share: jami pulga nisbatan foiz (butun son)."""
    total = sum(max(0.0, float(value(r) or 0)) for lst in groups.values() for r in lst)
    out = {}
    for cls in CLASSES:
        s = sum(max(0.0, float(value(r) or 0)) for r in groups.get(cls, []))
        out[cls] = {"n": len(groups.get(cls, [])), "sum": s, "share": int(round(100.0 * s / total)) if total else 0}
    return out
