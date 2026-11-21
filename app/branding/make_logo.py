# -*- coding: utf-8 -*-
"""TimeSaver logosi: dumaloq tun ko'ki fon, laym rangli «vaqt halqasi», shaftoli rangli T harfi.
Bitta geometriyadan: logo.svg, timesaver.ico (16..256), logo-*.png, tray holat ikonkalari, Inno Setup rasmlari.
Ishlatish: python make_logo.py  (branding papkasiga yozadi)."""
import math
import os

from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
NAVY, LIME, PEACH, CREAM = "#1F2A6B", "#D4F06A", "#FF9E7A", "#FFF4E8"
R = 85.0                      # logo radiusi (birlik koordinatalarda, markaz 0,0)
RING_R, RING_W = 52.0, 10.0   # halqa
RING_FROM, RING_TO = -90.0, 180.0   # soat strelkasi bo'yicha: 12 dan 9 gacha (270°)
DOT = (-37.0, -37.0, 6.0)
BAR = (-30.0, -27.0, 30.0, -9.0, 9.0)    # T ko'ndalang: x0, y0, x1, y1, radius
STEM = (-9.0, -27.0, 9.0, 29.0, 9.0)     # T ustuni
STATUS = {"ok": "#3CCB7F", "busy": "#F5B544", "bad": "#F0524A", "maint": "#5B8DEF"}


def rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4)) + (255,)


def svg():
    a0, a1 = math.radians(RING_FROM), math.radians(RING_TO)
    x0, y0 = RING_R * math.cos(a0), RING_R * math.sin(a0)
    x1, y1 = RING_R * math.cos(a1), RING_R * math.sin(a1)
    return f"""<svg xmlns="http://www.w3.org/2000/svg" viewBox="-85 -85 170 170" width="512" height="512">
  <title>TimeSaver</title>
  <circle r="85" fill="{NAVY}"/>
  <path d="M{x0:.2f} {y0:.2f} A{RING_R} {RING_R} 0 1 1 {x1:.2f} {y1:.2f}" fill="none" stroke="{LIME}" stroke-width="{RING_W}" stroke-linecap="round"/>
  <circle cx="{DOT[0]}" cy="{DOT[1]}" r="{DOT[2]}" fill="{PEACH}"/>
  <rect x="{BAR[0]}" y="{BAR[1]}" width="{BAR[2] - BAR[0]}" height="{BAR[3] - BAR[1]}" rx="{BAR[4]}" fill="{PEACH}"/>
  <rect x="{STEM[0]}" y="{STEM[1]}" width="{STEM[2] - STEM[0]}" height="{STEM[3] - STEM[1]}" rx="{STEM[4]}" fill="{PEACH}"/>
</svg>
"""


def draw(size, status=None, k=8):
    """PIL bilan chizish (8x katta chizib, kichraytiriladi — silliq qirralar uchun)."""
    S = size * k
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    s = S / (2 * R)                       # birlik -> piksel
    P = lambda x, y: (S / 2 + x * s, S / 2 + y * s)
    d.ellipse([0, 0, S - 1, S - 1], fill=rgb(NAVY))
    o = RING_R + RING_W / 2               # PIL qalinlikni chegaradan ichkariga chizadi
    box = [*P(-o, -o), *P(o, o)]
    d.arc(box, RING_FROM, RING_TO, fill=rgb(LIME), width=int(RING_W * s))
    for ang in (RING_FROM, RING_TO):      # yumaloq uchlar
        cx, cy = P(RING_R * math.cos(math.radians(ang)), RING_R * math.sin(math.radians(ang)))
        rr = RING_W * s / 2
        d.ellipse([cx - rr, cy - rr, cx + rr, cy + rr], fill=rgb(LIME))
    cx, cy = P(DOT[0], DOT[1])
    d.ellipse([cx - DOT[2] * s, cy - DOT[2] * s, cx + DOT[2] * s, cy + DOT[2] * s], fill=rgb(PEACH))
    for x0, y0, x1, y1, rad in (BAR, STEM):
        d.rounded_rectangle([*P(x0, y0), *P(x1, y1)], radius=rad * s, fill=rgb(PEACH))
    if status:                            # tray: o'ng pastda holat nuqtasi
        r = S * 0.17
        cx = cy = S - r - S * 0.02
        d.ellipse([cx - r - S * 0.035, cy - r - S * 0.035, cx + r + S * 0.035, cy + r + S * 0.035], fill=rgb(CREAM))
        d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=rgb(STATUS[status]))
    return img.resize((size, size), Image.LANCZOS)


def wizard_images():
    """Inno Setup: chap katta rasm (164x314, 2x: 328x628) va o'ng yuqori kichik rasm (55x58, 2x: 110x116)."""
    for w, h, name in ((164, 314, "wizard.bmp"), (328, 628, "wizard-2x.bmp")):
        img = Image.new("RGB", (w, h), rgb(NAVY)[:3])
        d = ImageDraw.Draw(img)
        for i in range(0, h, max(4, h // 40)):     # nozik gorizontal chiziqlar (laym, juda xira)
            d.line([(0, i), (w, i)], fill=(36, 50, 120))
        logo = draw(int(w * 0.62))
        img.paste(logo, ((w - logo.width) // 2, int(h * 0.22)), logo)
        img.save(os.path.join(HERE, name))
    for w, h, name in ((55, 58, "wizard-small.bmp"), (110, 116, "wizard-small-2x.bmp")):
        img = Image.new("RGB", (w, h), (255, 255, 255))
        logo = draw(int(min(w, h) * 0.9))
        img.paste(logo, ((w - logo.width) // 2, (h - logo.height) // 2), logo)
        img.save(os.path.join(HERE, name))


def main():
    with open(os.path.join(HERE, "logo.svg"), "w", encoding="utf-8") as f:
        f.write(svg())
    for n in (64, 128, 256, 512):
        draw(n).save(os.path.join(HERE, "logo-%d.png" % n))
    sizes = [16, 20, 24, 32, 40, 48, 64, 128, 256]
    big = draw(256)
    big.save(os.path.join(HERE, "timesaver.ico"), sizes=[(n, n) for n in sizes],
             append_images=[draw(n) for n in sizes[:-1]])
    for k in STATUS:
        draw(64, k).save(os.path.join(HERE, "tray-%s.png" % k))
    wizard_images()
    print("tayyor:", sorted(os.listdir(HERE)))


if __name__ == "__main__":
    main()
