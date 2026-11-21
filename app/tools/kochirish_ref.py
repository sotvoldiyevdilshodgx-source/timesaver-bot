# -*- coding: utf-8 -*-
"""
Ko'chirish mantig'ining Python nusxasi (etalon). Excel makrosi vba/Kochirish.bas shu qoidalarning AYNAN o'zini bajaradi:
ikkala kodda qadamlar bir xil raqamlangan ([1] ... [9]), qatorma-qator solishtirish uchun.

Ishlatish:
    python tools/kochirish_ref.py                 -> config.ini dagi Excel bo'yicha natija (makros nima chiqarishini oldindan ko'rish)
    python tools/kochirish_ref.py fayl.xlsm       -> boshqa fayl bo'yicha

Qoidalar (sotuv manfiy bo'lsa 0; q = Int(yaroqli qoldiq); ceil(x) = -Int(-x)):
    ehtiyoj(ol)   = max(0, ceil(s(ol) * KOCH_COVER) - q(ol))        oluvchiga FAQAT kerakli miqdor (davr sotuvi)
    qoldiradi(b)  = max(MIN_QOLDIQ, ceil(s(b) * KOCH_COVER))        beruvchi hech qachon 0 da qolmaydi
    bera_oladi(b) = q(b) - qoldiradi(b)
    SHOSHILINCH: q(ol)=0, s(ol)>0                       -> min(bera_oladi, ehtiyoj)
    ASSORTIMENT: q(ol)=0, s(ol)=0, boshqa filialda sotuv bor -> min(ASSORT_SONI, q(b) - MIN_QOLDIQ)  (beruvchi sotuvisiz)
    ODDIY      : s(b)=0 va "30+ kun" >= TURGAN_KUN, s(ol)>0, q(ol)>0 -> min(bera_oladi, ehtiyoj); summa < MIN_SUMMA_ODDIY — yo'q
    Taqsimlash: bera_oladi va ehtiyoj har yozilgan qatordan keyin kamayadi; SHOSHILINCH → ASSORTIMENT → ODDIY,
    oluvchilar ehtiyoj ↓ (teng — filial nomi), beruvchilar bera_oladi ↓ (teng — "30+ kun" ↓, filial nomi).
    Himoya: beruvchida qoladigan >= MIN_QOLDIQ va berilgan <= ehtiyoj (ASSORTIMENT: <= ASSORT_SONI), aks holda yozilmaydi.
    SKLADda qoldiq bo'lsa: filiallardan ko'chirilmaydi, "Skladdan" ga ehtiyojcha (sklad Int(K) gacha; MIN_QOLDIQ yo'q).
"""
import datetime
import math
import os
import re
import sys

# [0] Sozlamalar (Kochirish.bas boshidagi Private Const lar bilan bir xil)
KOCH_COVER = 1               # oluvchiga necha davrlik sotuv (1 = davr sotuvicha)
MIN_QOLDIQ = 1               # beruvchida kamida shuncha qoladi (< 1 bo'lsa 1)
TURGAN_KUN = 30
MIN_SUMMA_ODDIY = 50000
MIN_SUMMA_SHOSH = 0          # SHOSHILINCH va ASSORTIMENT uchun
ASSORT_SONI = 1
SROK_MIN_KUN = 0

Z1, Z2, OD = 1, 2, 3          # SHOSHILINCH, ASSORTIMENT, ODDIY (tartib = varaqda saralash tartibi)
TURI = {Z1: "SHOSHILINCH", Z2: "ASSORTIMENT", OD: "ODDIY"}
SABAB_Z1 = "Tugagan, sotuv bor"
SABAB_Z2 = "Filialda yo'q, boshqa filialda sotilyapti"
SABAB_OD = "%d kun turibdi, oluvchida yetmaydi"
SABAB_SKLAD_OD = "Skladda bor, oluvchida yetmaydi"
APOS = {"ʼ": "'", "’": "'", "‘": "'", "`": "'"}


def norm(s):
    """VBA: UCase(WorksheetFunction.Trim(CStr(v))) + apostroflar bittasiga — ortiqcha probellar birlashadi, katta harf."""
    t = re.sub(r"\s+", " ", str(s if s is not None else "")).strip().upper()
    for a, b in APOS.items():
        t = t.replace(a, b)
    return t


def num(v):
    """VBA Num(): son bo'lmasa 0."""
    if isinstance(v, bool) or v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(" ", "").replace(",", "."))
    except ValueError:
        return 0.0


def vint(x):
    """VBA Int(): pastga yaxlitlash (manfiyda ham)."""
    return math.floor(x)


def r6(x):
    """VBA Round(x, 6): suzuvchi nuqta xatosini yo'qotadi (VBA 0,15*30-3,5 ni 0,99999.. deb hisoblaydi, Python 1,0)."""
    return round(x, 6)


def vceil(x):
    """VBA -Int(-Round(x, 6)): yuqoriga yaxlitlash (0,3333 -> 1; 2,0000001 -> 2)."""
    return -vint(-r6(x))


def as_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    return None


def build(sale, filials, davr_kun, srok, old, narx, today=None, sklad_name="СКЛАД", log=None):
    """
    sale    : [{"name","producer","q":[3],"s":[3],"k": sklad qoldig'i, "l": srok (Лист1 L, ixtiyoriy)}]
    filials : [E6, G6, I6] filial nomlari
    srok    : [{"filial","name","producer","date","kun","ost","prih"}]   (Srok varag'i)
    old     : [{"filial","name","producer","kun"}]                       ("30+ kun" varag'i)
    narx    : [{"name","producer","prih"}]                               (Narx varag'i)
    log     : ro'yxat berilsa — himoya tekshiruvidan o'tmagan qatorlar (VBA: yashirin "KochLog" varag'i)
    Natija  : (kochirish qatorlari, skladdan qatorlari, xabar). davr_kun miqdorga ta'sir qilmaydi (KOCH_COVER — davrlar).
    """
    today = today or datetime.date.today()
    fil = [norm(f) for f in filials]
    nf = len(fil)
    mq = max(1, MIN_QOLDIQ)
    log = log if log is not None else []

    # [1] Srok: har filial+tovar uchun yaroqli qoldiq, eng yaqin yaroqli muddat, kirim narxi
    found, valid, near, sprice = set(), {}, {}, {}
    for r in srok:
        f = norm(r["filial"])
        if not f or not norm(r["name"]):
            continue
        fk = f + "|" + norm(r["name"]) + "|" + norm(r["producer"])
        ost = num(r["ost"])
        if ost <= 0:
            continue
        found.add(fk)
        d = as_date(r.get("date"))
        kun = (d - today).days if d else num(r.get("kun"))
        if kun >= 0 and kun >= SROK_MIN_KUN:          # muddati o'tgan partiya doim chiqariladi
            valid[fk] = valid.get(fk, 0.0) + ost
            if d and (fk not in near or d < near[fk]):
                near[fk] = d
        if fk not in sprice and num(r.get("prih")) > 0:
            sprice[fk] = num(r["prih"])

    # [2] "30+ kun": filial+tovar necha kun turibdi (eng kattasi)
    turgan = {}
    for r in old:
        fk = norm(r["filial"]) + "|" + norm(r["name"]) + "|" + norm(r["producer"])
        turgan[fk] = max(turgan.get(fk, 0.0), num(r["kun"]))

    # [3] Narx: birinchi uchragan kirim narxi
    price_of = {}
    for r in narx:
        k = norm(r["name"]) + "|" + norm(r["producer"])
        if norm(r["name"]) and k not in price_of:
            price_of[k] = num(r["prih"])

    koch, sklad = [], []
    for p in sale:
        nm = norm(p["name"])
        if not nm:
            continue
        k = nm + "|" + norm(p["producer"])

        # [4] har filial: qoldiq q = Int(yaroqli), sotuv, bera oladi (SHOSHILINCH/ODDIY va ASSORTIMENT uchun alohida)
        raw, q, s, cap, capz2, tk, given = ([0.0] * nf for _ in range(7))
        any_sale = False
        for i in range(nf):
            raw[i], s[i] = num(p["q"][i]), max(0.0, num(p["s"][i]))   # manfiy sotuv (qaytarilgan tovar) = 0
            if s[i] > 0:
                any_sale = True
            fk = fil[i] + "|" + k
            v = min(valid.get(fk, 0.0), raw[i]) if fk in found else raw[i]
            q[i] = max(0, vint(r6(v)))                                # kasr qadoq berilmaydi
            # beruvchi hech qachon 0 da qolmaydi: o'zida max(MIN_QOLDIQ, o'z davr sotuvi) qoladi
            cap[i] = max(0, q[i] - max(mq, vceil(s[i] * KOCH_COVER)))
            # ISTISNO (ASSORTIMENT): beruvchining o'z sotuvi hisobga olinmaydi, faqat MIN_QOLDIQ o'zida qoladi
            capz2[i] = max(0, q[i] - mq)
            tk[i] = turgan.get(fk, 0.0)
        price = price_of.get(k, 0.0)
        if price <= 0:
            price = max([sprice.get(fil[i] + "|" + k, 0.0) for i in range(nf)] + [0.0])

        # [5] oluvchilar: turi va ehtiyoj (faqat kerakli miqdor: davr sotuvi − qoldig'i)
        recv = []
        for j in range(nf):
            eh = max(0, vceil(s[j] * KOCH_COVER) - q[j])
            if q[j] == 0 and s[j] > 0:
                recv.append((j, Z1, eh))
            elif q[j] == 0 and s[j] == 0 and any_sale:
                recv.append((j, Z2, ASSORT_SONI))
            elif q[j] > 0 and s[j] > 0 and eh >= 1:
                recv.append((j, OD, eh))
        # [6] tartib: SHOSHILINCH → ASSORTIMENT → ODDIY, ichida ehtiyoji kattasi, teng bo'lsa filial nomi (alifbo)
        recv.sort(key=lambda x: (x[1], -x[2], fil[x[0]], x[0]))

        skq = num(p.get("k"))
        if skq > 0:
            # [7] SKLAD: filiallardan ko'chirilmaydi, sklad Int(K) gacha ehtiyojcha beradi (MIN_QOLDIQ va "30+ kun" yo'q)
            skcap = vint(skq)
            donors = [i for i in range(nf) if raw[i] > 0 and s[i] <= 0]
            for j, tp, need in recv:
                if skcap < 1:
                    break
                n = vint(min(need, skcap))
                summa = n * price
                if (tp == OD and summa < MIN_SUMMA_ODDIY) or (tp != OD and summa < MIN_SUMMA_SHOSH):
                    continue
                sklad.append({"dst": fil[j], "name": p["name"], "producer": p["producer"], "sklad": skq,
                              "dst_sales": s[j], "dst_stock": raw[j], "donors": ", ".join(fil[i] for i in donors),
                              "donor_qty": sum(raw[i] for i in donors), "qty": n, "rank": tp, "turi": TURI[tp],
                              "sabab": {Z1: SABAB_Z1, Z2: SABAB_Z2, OD: SABAB_SKLAD_OD}[tp], "prih": price, "summa": summa})
                skcap -= n
            continue

        # [8] filiallararo: har oluvchiga beruvchilar — ko'p bera oladigani birinchi (teng: ko'proq turgani, filial nomi)
        for j, tp, need in recv:
            rem = need
            cand = []
            for g in range(nf):
                if g == j:
                    continue
                avail = (capz2[g] if tp == Z2 else cap[g]) - given[g]
                if avail < 1:
                    continue
                if tp == OD and not (s[g] <= 0 and tk[g] >= TURGAN_KUN):
                    continue
                cand.append((g, avail))
            cand.sort(key=lambda x: (-x[1], -tk[x[0]], fil[x[0]], x[0]))
            for g, avail in cand:
                n = vint(min(rem, avail))
                if n < 1:
                    continue
                summa = n * price
                if (tp == OD and summa < MIN_SUMMA_ODDIY) or (tp != OD and summa < MIN_SUMMA_SHOSH):
                    continue
                left = q[g] - given[g] - n
                if left < mq or n > need:     # himoya: beruvchi 0 da qolmasin, oluvchiga ehtiyojdan ko'p emas
                    log.append({"src": fil[g], "dst": fil[j], "name": p["name"], "producer": p["producer"], "qty": n,
                                "turi": TURI[tp], "qoldiq": q[g], "qoladi": left, "ehtiyoj": need})
                    continue
                koch.append({"src": fil[g], "dst": fil[j], "name": p["name"], "producer": p["producer"], "qty": n,
                             "dst_sales": s[j], "dst_stock": raw[j], "prih": price, "summa": summa,
                             "date": near.get(fil[g] + "|" + k), "rank": tp, "turi": TURI[tp],
                             "sabab": {Z1: SABAB_Z1, Z2: SABAB_Z2, OD: SABAB_OD % vint(tk[g])}[tp],
                             "src_sales": s[g], "src_stock": raw[g], "src_valid": q[g]})
                given[g] += n
                rem -= n
                if rem < 1:
                    break

    # [9] tartib: Turi (SHOSHILINCH → ASSORTIMENT → ODDIY), eng yaqin muddat (bo'shi oxirida), summa kattasi
    koch.sort(key=lambda r: (r["rank"], r["date"] or datetime.date.max, -r["summa"]))
    sklad.sort(key=lambda r: (r["rank"], -r["summa"]))
    cnt = {t: sum(1 for r in koch if r["rank"] == t) for t in (Z1, Z2, OD)}
    msg = "SHOSHILINCH: %d · ASSORTIMENT: %d · ODDIY: %d · jami %s so'm · Skladdan: %d" % (
        cnt[Z1], cnt[Z2], cnt[OD], "{:,.0f}".format(sum(r["summa"] for r in koch)).replace(",", " "), len(sklad))
    return koch, sklad, msg


# ---------------------------------------------------------------- Excel dan o'qish (oldindan ko'rish uchun)
def load_xlsx(path):
    from openpyxl import load_workbook
    import shutil
    import tempfile
    tmp = os.path.join(tempfile.gettempdir(), "koch_ref_%d%s" % (os.getpid(), os.path.splitext(path)[1]))
    shutil.copy2(path, tmp)
    wb = load_workbook(tmp, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        head = list(ws.iter_rows(min_row=4, max_row=6, max_col=11, values_only=True))
        d1, d2 = as_date(head[0][2]), as_date(head[1][2])
        davr = (d2 - d1).days if d1 and d2 else 30
        filials = [head[2][4], head[2][6], head[2][8]]
        sklad_name = head[2][10] or "СКЛАД"
        sale = []
        for r in ws.iter_rows(min_row=8, max_col=12, values_only=True):
            if r[0] in (None, ""):
                continue
            sale.append({"name": r[0], "producer": r[1], "q": [r[4], r[6], r[8]], "s": [r[5], r[7], r[9]],
                         "k": r[10], "l": r[11]})
        get = lambda name: wb[name] if name in wb.sheetnames else None
        srok = [{"filial": r[0], "name": r[1], "producer": r[2], "date": r[4], "kun": r[5], "ost": r[6], "prih": r[7]}
                for r in (get("Srok").iter_rows(min_row=5, max_col=10, values_only=True) if get("Srok") else []) if r[0]]
        old = [{"filial": r[0], "name": r[1], "producer": r[2], "kun": r[5]}
               for r in (get("30+ kun").iter_rows(min_row=5, max_col=9, values_only=True) if get("30+ kun") else []) if r[0]]
        narx = [{"name": r[0], "producer": r[1], "prih": r[3]}
                for r in (get("Narx").iter_rows(min_row=5, max_col=6, values_only=True) if get("Narx") else []) if r[0]]
        return dict(sale=sale, filials=filials, davr_kun=davr, srok=srok, old=old, narx=narx, sklad_name=sklad_name)
    finally:
        wb.close()
        os.remove(tmp)


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else None
    if not path:
        import configparser
        cp = configparser.ConfigParser(interpolation=None)
        cp.read(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "config.ini"), encoding="utf-8-sig")
        path = cp["excel"]["file"].strip()
    data = load_xlsx(path)
    koch, sklad, msg = build(**data)
    print(msg)
    print("davr: %s kun · filiallar: %s" % (data["davr_kun"], ", ".join(norm(f) for f in data["filials"])))
    for r in koch[:25]:
        print("  %-6s %-18s -> %-18s %4d x %-40s %12s  %s" % (r["turi"], r["src"][:18], r["dst"][:18], r["qty"],
                                                            str(r["name"])[:40], "{:,.0f}".format(r["summa"]), r["sabab"]))


if __name__ == "__main__":
    main()
