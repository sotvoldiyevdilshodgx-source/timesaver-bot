# -*- coding: utf-8 -*-
"""
TimeSaver: bot + server paneli + Windows soat yonidagi (tray) ikonka.

Ishga tushirish:  start_tray.bat  (konsol oynasisiz, pythonw orqali)
Ikonka rangi:  yashil = bot ishlayapti | to'q sariq = Excel yangilanmoqda | qizil = Telegramga ulanmagan |
               ko'k = texnik rejim (mijozlar uchun o'chiq).
Ikonkani bosing (o'ng tugma): Панелни очиш, Янгилаш (Excel), Журнал, Чиқиш.
Panel alohida oynada ochiladi (timesaver_window.py, pywebview); exe: TimeSaver.exe --window.
Kerak: pip install pystray pillow pywebview
"""
import logging
import os
import socket
import sys
import threading
import time

if __name__ == "__main__" and "--window" in sys.argv[1:]:
    # TimeSaver.exe --window: faqat panel oynasi (alohida jarayon, pywebview asosiy oqimda ishlaydi)
    import timesaver_window
    sys.exit(timesaver_window.main([a for a in sys.argv[1:] if a != "--window"]))

from PIL import Image, ImageDraw

import timesaver_bot as B
import timesaver_window as W

COLORS = {"ok": (47, 158, 95), "busy": (232, 160, 48), "bad": (214, 69, 54), "maint": (70, 120, 200)}
LOCK_PORT = 47241  # TimeSaver (47231) bilan bir kompyuterda ishlay oladi


def make_icon(kind="ok", size=64):
    """TimeSaver logosi + o'ng pastda holat nuqtasi (branding/tray-<kind>.png). Fayl bo'lmasa — oddiy chizilgan ikonka."""
    try:
        return Image.open(os.path.join(B.APP_DIR, "branding", "tray-%s.png" % kind)).convert("RGBA").resize((size, size), Image.LANCZOS)
    except OSError:
        return _drawn_icon(kind, size)


def _drawn_icon(kind="ok", size=64):
    """Zaxira: yumaloq to'rtburchak ustida oq «A» harfi."""
    k = 8
    S = size * k
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([S * 0.04, S * 0.04, S * 0.96, S * 0.96], radius=S * 0.24, fill=COLORS[kind] + (255,))
    white = (255, 255, 255, 255)
    t = S * 0.115  # chiziq qalinligi
    apex, bl, br = (S * 0.5, S * 0.20), (S * 0.24, S * 0.80), (S * 0.76, S * 0.80)
    d.line([bl, apex, br], fill=white, width=int(t), joint="curve")
    for x, y in (bl, br):
        d.ellipse([x - t / 2, y - t / 2, x + t / 2, y + t / 2], fill=white)
    d.ellipse([apex[0] - t / 2, apex[1] - t / 2, apex[0] + t / 2, apex[1] + t / 2], fill=white)
    yb = S * 0.62  # ko'ndalang chiziq
    xl = bl[0] + (apex[0] - bl[0]) * (bl[1] - yb) / (bl[1] - apex[1])
    xr = br[0] - (br[0] - apex[0]) * (br[1] - yb) / (br[1] - apex[1])
    d.line([(xl, yb), (xr, yb)], fill=white, width=int(t * 0.9))
    return img.resize((size, size), Image.LANCZOS)


def state_of(bot):
    """(holat, ikonka sarlavhasi)"""
    d = bot.data
    when = B.dt_text(d.when()) if d.when() else "-"
    if bot.rstate.get("running"):
        done = sum(1 for s in bot.rstate.get("steps", {}).values() if s[0] == "ok")
        return "busy", B.BRAND + ": янгиланмоқда (%d/%d)" % (done, len(B.SHEET_ORDER))
    if time.time() - bot.last_poll > 90:
        return "bad", B.BRAND + ": Телеграмга уланмаган"
    if bot.maint_on:
        return "maint", B.BRAND + ": ТЕХНИК РЕЖИМ (мижозлар учун ўчиқ)"
    return "ok", B.BRAND + ": ишлаяпти\nМаълумот: %s" % when


def single_instance():
    s = socket.socket()
    try:
        s.bind(("127.0.0.1", LOCK_PORT))
    except OSError:
        return None
    return s


class Tray:
    def __init__(self, bot, cfg, pystray):
        self.bot, self.cfg, self.ps = bot, cfg, pystray
        self.port = int(cfg.get("panel", "port", fallback="8080")) if cfg.has_section("panel") else 8080
        self.icon = None
        self.kind = None
        self.was_running = False

    # --- menyu amallari ---
    def open_panel(self, icon=None, item=None):
        W.open_window(getattr(self.bot, "panel_port", self.port))

    def refresh(self, icon=None, item=None):
        if not self.bot.start_refresh(None):
            self.notify("Янгиланиш аллақачон кетмоқда.")

    def toggle_maint(self, icon=None, item=None):
        on = not self.bot.maint_on
        n = self.bot.set_maint(on, by="tray")
        self.notify("🛠 Техник режим ёқилди: мижозлар ботдан фойдалана олмайди" if on
                    else "✅ Техник режим ўчирилди. Кутганларга хабар: %d та" % n)

    def open_log(self, icon=None, item=None):
        try:
            os.startfile(os.path.join(B.BASE_DIR, "bot.log"))  # noqa (faqat Windows)
        except Exception:  # noqa
            pass

    def quit(self, icon=None, item=None):
        W.close_window()  # panel serveri to'xtaydi — ochiq oyna bo'sh qolmasin
        if self.icon:
            self.icon.stop()
        os._exit(0)

    def notify(self, text):
        try:
            if self.icon:
                self.icon.notify(text, B.BRAND)
        except Exception:  # noqa
            pass

    # --- holatni kuzatish ---
    def watch(self):
        while True:
            try:
                kind, title = state_of(self.bot)
                if self.icon:
                    if kind != self.kind:
                        self.icon.icon = make_icon(kind)
                        self.kind = kind
                    self.icon.title = title[:127]
                running = bool(self.bot.rstate.get("running"))
                if self.was_running and not running:
                    ok = self.bot.rstate.get("ok")
                    self.notify("✅ Маълумот янгиланди" if ok else "❌ Янгилаш бажарилмади: %s" % self.bot.rstate.get("msg", "")[:150])
                self.was_running = running
            except Exception:  # noqa
                logging.exception("Tray kuzatuvi xatosi")
            time.sleep(2)

    def run(self):
        ps = self.ps
        menu = ps.Menu(
            ps.MenuItem("Панелни очиш", self.open_panel, default=True),
            ps.MenuItem("Янгилаш (Excel)", self.refresh),
            ps.MenuItem("Техник режим (ботни узиш)", self.toggle_maint, checked=lambda item: self.bot.maint_on),
            ps.MenuItem("Журнални очиш", self.open_log),
            ps.Menu.SEPARATOR,
            ps.MenuItem("Чиқиш", self.quit),
        )
        self.icon = ps.Icon(B.BRAND, make_icon("bad"), B.BRAND, menu)
        threading.Thread(target=self.watch, daemon=True).start()
        self.icon.run()


def main():
    lock = single_instance()
    if lock is None:
        # Ikkinchi nusxa Telegramda 409 xatosini beradi: yangisini ochmaymiz, ishlayotganining panel oynasini ochamiz.
        try:
            with open(os.path.join(B.BASE_DIR, "panel_port.txt")) as f:
                port = int(f.read().strip())
            W.open_window(port)
            return
        except (OSError, ValueError):
            pass
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, B.BRAND + " аллақачон ишлаяпти (соат ёнидаги иконкага қаранг).", B.BRAND, 0x40)
        except Exception:  # noqa
            pass
        return
    try:
        import pystray
    except ImportError:
        print("pystray ўрнатилмаган. Бажаринг:  python -m pip install pystray pillow")
        return
    try:
        bot = B.create_bot()
    except BaseException as e:  # noqa  (pythonw da xato ko'rinmay qolmasin)
        msg = str(e) if isinstance(e, SystemExit) else "Бот ишга тушмади: %s\n\nБатафсил: startup_error.txt" % e
        try:
            import traceback
            with open(os.path.join(B.BASE_DIR, "startup_error.txt"), "w", encoding="utf-8") as f:
                f.write(traceback.format_exc())
        except Exception:  # noqa
            pass
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, msg, B.BRAND, 0x10)
        except Exception:  # noqa
            print(msg)
        return
    if not hasattr(bot, "panel_port"):
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, "Сервер панели ишга тушмади (порт банд бўлиши мумкин). Батафсил: bot.log", B.BRAND, 0x30)
        except Exception:  # noqa
            pass
    threading.Thread(target=bot.run_forever, daemon=True, name="bot").start()
    tray = Tray(bot, bot.cfg, pystray)
    if not bot.setup_done and hasattr(bot, "panel_port"):  # birinchi ishga tushish: sozlash ustasi panel oynasida ochiladi
        threading.Timer(1.5, tray.open_panel).start()
    tray.run()


if __name__ == "__main__":
    main()
