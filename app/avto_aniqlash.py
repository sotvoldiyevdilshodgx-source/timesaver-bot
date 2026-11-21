# -*- coding: utf-8 -*-
"""Ko'chirishni qoldiq o'zgarishidan avtomatik aniqlash (9-talab, dizayn/AVTO-ANIQLASH.md).

Topshiriq yaratilganda beruvchi (src) va oluvchi (dst) filialdagi qoldiq surati saqlanadi. Har Excel yangilanishida
hozirgi qoldiq bilan solishtiriladi:
- beruvchida kamaydi VA oluvchida ko'paydi (ikkalasi ham rejaning >= 80%)  → "olindi"  (son = min(kamaydi, ko'paydi, reja))
- faqat beruvchida kamaydi                                                → "berildi" — faqat ISHORA (HINTS): bot holatni
  o'zgartirmaydi. Beruvchidagi kamayishni oddiy sotuvdan ajratib bo'lmaydi; topshiriq oluvchida ko'payish bilan birga
  «olindi» bo'lguncha (yoki xodim «Berdim» bosguncha) «yangi» qoladi. Ishora faqat «Berdim» dan keyin «qoldiq o'zgarmadi»
  ogohlantirishini to'xtatadi (beruvchida o'zgarish bor).
- faqat oluvchida ko'paydi                                                 → "tekshiring" (boshqa joydan kelgan bo'lishi mumkin)
- hech biri                                                                → None

Bir beruvchidan bir xil dori bir nechta oluvchiga (attribute): har topshiriq o'z oluvchisining ko'payishi bo'yicha baholanadi;
beruvchidagi jami kamayish faqat oluvchisi ko'paygan topshiriqlar rejalari yig'indisi bilan solishtiriladi.
"""

THRESHOLD = 0.8
HINTS = ("berildi",)  # holatga qo'llanmaydigan natijalar
_EPS = 1e-9


def decide(plan, src0, src_now, dst0, dst_now, threshold=THRESHOLD):
    """(holat yoki None, son). plan — reja bo'yicha son; qoldiqlar None bo'lsa — ma'lumot yo'q, qaror yo'q.
    "berildi" — faqat ishora (HINTS), holat o'zgarmaydi."""
    if not plan or plan <= 0 or None in (src0, src_now, dst0, dst_now):
        return None, 0
    need = plan * threshold - _EPS
    down = float(src0) - float(src_now)
    up = float(dst_now) - float(dst0)
    gave, got = down >= need, up >= need
    if gave and got:
        return "olindi", min(down, up, float(plan))
    if gave:
        return "berildi", min(down, float(plan))
    if got:
        return "tekshiring", min(up, float(plan))
    return None, 0


def attribute(items, threshold=THRESHOLD):
    """Bir beruvchi + bir dori bo'yicha topshiriqlar (bir yoki bir nechta oluvchi). items: [{"plan", "down", "up"}] —
    down: beruvchidagi kamayish (guruhda eng kattasi olinadi), up: shu topshiriq oluvchisining ko'payishi.
    Qaytaradi: [(holat yoki None, son, beruvchidan shu topshiriqqa ulush)] — items tartibida.

    - Oluvchisi rejaning >= 80% iga ko'paygan topshiriqlar («ko'paygan»lar) «olindi» bo'ladi, agar beruvchidagi jami kamayish
      ularning rejalari yig'indisining >= 80% i bo'lsa. Yetmasa — rejasi kattasidan boshlab sig'ganicha «olindi», qolgan
      ko'paygan topshiriqlar «tekshiring» (oluvchiga boshqa joydan kelgan bo'lishi mumkin).
    - Oluvchisi ko'paymagan topshiriq qo'shnisi tufayli ogohlantirilmaydi; unga faqat «olindi» lardan ortgan kamayish
      tegishli (yetarli bo'lsa — "berildi" ishorasi).
    Bitta topshiriq uchun natija decide() bilan bir xil."""
    n = len(items)
    if not n:
        return []
    plan = [max(0.0, float(x.get("plan") or 0)) for x in items]
    up = [float(x.get("up") or 0) for x in items]
    total = max([0.0] + [float(x.get("down") or 0) for x in items])
    rose = [i for i in range(n) if plan[i] > 0 and up[i] >= plan[i] * threshold - _EPS]
    taken, need = [], 0.0
    for i in sorted(rose, key=lambda i: (-plan[i], i)):
        if total >= (need + plan[i]) * threshold - _EPS:
            taken.append(i)
            need += plan[i]
    others = [i for i in range(n) if i not in taken]
    d = [0.0] * n
    if taken:
        used = total if not others else min(total, need)
        for i, v in share(used, {i: plan[i] for i in taken}).items():
            d[i] = v
        rest = total - used
    else:
        rest = total
    if others:
        sh = share(rest, {i: plan[i] for i in others})
        if not any(plan[i] > 0 for i in others):
            sh = {i: rest / len(others) for i in others}
        for i, v in sh.items():
            d[i] = v
    out = []
    for i in range(n):
        p = plan[i]
        if p <= 0:
            out.append((None, 0, d[i]))
        elif i in taken:
            out.append(("olindi", min(d[i], up[i], p), d[i]))
        elif i in rose:
            out.append(("tekshiring", min(up[i], p), d[i]))
        elif d[i] >= p * threshold - _EPS:
            out.append(("berildi", min(d[i], p), d[i]))
        else:
            out.append((None, 0, d[i]))
    return out


def share(changes, plans):
    """Bir xil dori bir nechta topshiriqda bo'lsa: umumiy o'zgarishni reja soniga mutanosib bo'ladi.
    changes — umumiy o'zgarish (son), plans — {id: reja}. {id: ulush} qaytaradi."""
    total = sum(max(0.0, float(p)) for p in plans.values())
    if total <= 0:
        return {k: 0.0 for k in plans}
    return {k: float(changes) * max(0.0, float(p)) / total for k, p in plans.items()}
