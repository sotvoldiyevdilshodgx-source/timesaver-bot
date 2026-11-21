# -*- coding: utf-8 -*-
"""
TimeSaver: server paneli — alohida ish stoli oynasi (pywebview + Edge WebView2), brauzer emas.

Tray (soat yonidagi ikonka) bu oynani ALOHIDA JARAYONDA ochadi (pywebview asosiy oqimni talab qiladi):
  o'rnatilgan exe:  TimeSaver.exe --window --port 8765
  manba koddan:     pythonw timesaver_window.py --port 8765

- Faqat BITTA oyna: oyna jarayoni 127.0.0.1:WINDOW_PORT ni band qiladi. Ikkinchi «ochish» so'rovi shu portga
  ulanadi -> mavjud oyna tiklanadi (minimallashtirilgan bo'lsa) va oldinga chiqariladi, yangi oyna ochilmaydi.
- ✕ faqat oynani yopadi: tray va bot ishlayveradi.
- Panel paroli ([panel] password) bo'lsa: oyna uni config.ini dan o'zi o'qiydi va WebView2 ning
  BasicAuthenticationRequested hodisasida faqat panel manziliga beradi (buyruq qatorida ham, URL da ham parol yo'q).
- Excel yuklab olish: «Saqlash» oynasi (Yuklanmalar papkasi). Fayl yuborish (<input type=file>) — Windows oynasi.
- WebView2 / pywebview ishlamasa — panel oddiy brauzerda ochiladi, sababi window.log ga yoziladi.
Kerak: pip install pywebview  (pythonnet bilan birga o'rnatiladi)
"""
import ctypes
import logging
import logging.handlers
import os
import socket
import subprocess
import sys
import threading
import webbrowser

TITLE = "TimeSaver"
SIZE = (1360, 860)
MIN_SIZE = (1100, 700)
WINDOW_PORT = 47242   # tray qulfi 47241; bu — oynaning «bitta nusxa» qulfi va «oldinga chiq» kanali
MAGIC = "TIMESAVER-WINDOW"
FROZEN = bool(getattr(sys, "frozen", False))
APP_DIR = getattr(sys, "_MEIPASS", None) or os.path.dirname(os.path.abspath(__file__))

_u32 = ctypes.WinDLL("user32", use_last_error=True) if os.name == "nt" else None
if _u32:
    from ctypes import wintypes
    for _n, _a in (("IsIconic", [wintypes.HWND]), ("IsWindow", [wintypes.HWND]), ("SetForegroundWindow", [wintypes.HWND]),
                   ("ShowWindow", [wintypes.HWND, ctypes.c_int]), ("AllowSetForegroundWindow", [wintypes.DWORD])):
        getattr(_u32, _n).argtypes = _a
        getattr(_u32, _n).restype = wintypes.BOOL


def panel_url(port):
    return "http://127.0.0.1:%d/" % int(port)


def _origin(url):
    """http://127.0.0.1:8765/#koch -> http://127.0.0.1:8765"""
    p = url.split("#", 1)[0].lower()
    i = p.find("/", p.find("//") + 2)
    return p if i < 0 else p[:i]


def _bring_to_front(hwnd):
    if not (_u32 and hwnd and _u32.IsWindow(hwnd)):
        return
    _u32.ShowWindow(hwnd, 9 if _u32.IsIconic(hwnd) else 5)  # SW_RESTORE (maksimal holati saqlanadi) | SW_SHOW
    _u32.SetForegroundWindow(hwnd)


def _readline(sock, limit=4096):
    buf = b""
    while not buf.endswith(b"\n") and len(buf) < limit:
        ch = sock.recv(1024)
        if not ch:
            break
        buf += ch
    return buf.decode("utf-8", "replace").strip()


# ----------------------------------------------------------------------------
# Tray tomoni: oynani ochish / oldinga chiqarish / yopish
# ----------------------------------------------------------------------------
def _ask(cmd, timeout=1.5):
    """Ishlayotgan oynaga buyruq yuboradi. (pid, hwnd) yoki None (oyna yo'q)."""
    try:
        s = socket.create_connection(("127.0.0.1", WINDOW_PORT), timeout=timeout)
    except OSError:
        return None
    with s:
        try:
            hello = _readline(s).split()
            if len(hello) != 3 or hello[0] != MAGIC:
                return None  # portni boshqa dastur egallagan
            pid, hwnd = int(hello[1]), int(hello[2])
        except (OSError, ValueError):
            return None
        try:
            if _u32:
                _u32.AllowSetForegroundWindow(pid)  # oyna jarayoniga o'zini oldinga chiqarishga ruxsat
            s.sendall((cmd + "\n").encode("utf-8"))
            _readline(s)
        except OSError:
            pass
        return pid, hwnd


def focus_existing(url=None):
    """Oyna ochiq bo'lsa — oldinga chiqaradi va True qaytaradi."""
    r = _ask("show " + (url or ""))
    if r is None:
        return False
    _bring_to_front(r[1])
    return True


def close_window():
    """Ochiq oynani yopadi (tray «Чиқиш» da — panel serveri to'xtaydi, oyna bo'sh qolmasin)."""
    _ask("close", timeout=0.5)


def _spawn(port):
    if FROZEN:
        cmd = [sys.executable, "--window", "--port", str(int(port))]
    else:
        exe = sys.executable
        w = os.path.join(os.path.dirname(exe), "pythonw.exe")
        if os.path.basename(exe).lower() == "python.exe" and os.path.isfile(w):
            exe = w  # konsol oynasi chiqmasin
        cmd = [exe, os.path.abspath(__file__), "--port", str(int(port))]
    env = dict(os.environ)
    env["PYINSTALLER_RESET_ENVIRONMENT"] = "1"  # exe o'zini mustaqil nusxa sifatida ishga tushiradi
    flags = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0) | getattr(subprocess, "CREATE_NO_WINDOW", 0)
    subprocess.Popen(cmd, env=env, close_fds=True, creationflags=flags,
                     stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def open_window(port):
    """Panel oynasini ochadi: ochiq bo'lsa — oldinga chiqaradi; ishga tushmasa — brauzerda ochadi."""
    url = panel_url(port)
    if focus_existing(url):
        return True
    try:
        _spawn(port)
        return True
    except Exception:  # noqa
        logging.exception("Panel oynasi ishga tushmadi — panel brauzerda ochiladi")
        webbrowser.open(url)
        return False


# ----------------------------------------------------------------------------
# Oyna jarayoni
# ----------------------------------------------------------------------------
class _State:
    def __init__(self, url):
        self.url = url
        self.origin = _origin(url)
        self.hwnd = 0
        self.window = None
        self.password = ""
        self.auth_tries = 0
        self.reload_password()

    def reload_password(self):
        self.auth_tries = 0
        try:
            import timesaver_bot as B
            cfg = B.load_config()
            self.password = str(cfg.get("panel", "password", fallback="")).strip() if cfg.has_section("panel") else ""
        except Exception as e:  # noqa
            logging.warning("config.ini o'qilmadi (panel paroli): %s", e)
            self.password = ""

    def show(self, url):
        self.reload_password()
        if url and _origin(url) != self.origin and self.window:
            self.url, self.origin = url, _origin(url)  # tray boshqa portda qayta ishga tushgan
            self.window.load_url(url)
        _bring_to_front(self.hwnd)


def _serve(lock, st):
    while True:
        try:
            conn, _ = lock.accept()
        except OSError:
            return
        threading.Thread(target=_handle, args=(conn, st), daemon=True).start()


def _handle(conn, st):
    with conn:
        try:
            conn.settimeout(3)
            conn.sendall(("%s %d %d\n" % (MAGIC, os.getpid(), st.hwnd)).encode("ascii"))
            line = _readline(conn)
            if line == "close":
                conn.sendall(b"ok\n")
                if st.window:
                    st.window.destroy()
            elif line.startswith("show"):
                st.show(line[4:].strip())
                conn.sendall(b"ok\n")
        except Exception:  # noqa
            logging.exception("Oyna buyrug'i xatosi")


def _bind():
    s = socket.socket()
    try:
        s.bind(("127.0.0.1", WINDOW_PORT))
        s.listen(4)
    except OSError:
        s.close()
        return None
    return s


def _setup_log(base_dir):
    try:
        h = logging.handlers.RotatingFileHandler(os.path.join(base_dir, "window.log"), maxBytes=300_000, backupCount=1, encoding="utf-8")
        logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s", handlers=[h])
    except OSError:
        logging.basicConfig(level=logging.INFO)


def _port_from(argv, base_dir):
    for i, a in enumerate(argv):
        if a == "--port" and i + 1 < len(argv):
            try:
                return int(argv[i + 1])
            except ValueError:
                pass
    try:
        with open(os.path.join(base_dir, "panel_port.txt")) as f:
            return int(f.read().strip())
    except (OSError, ValueError):
        pass
    try:
        import timesaver_bot as B
        return int(B.load_config().get("panel", "port", fallback="8080"))
    except Exception:  # noqa
        return 8080


def _patch_webview2(st):
    """pywebview ning EdgeChrome.on_webview_ready iga ulanamiz (UI oqimida, birinchi so'rovdan oldin)."""
    from webview.platforms import edgechromium as ec
    orig = ec.EdgeChrome.on_webview_ready

    def on_auth(sender, args):
        try:
            uri = str(args.Uri).lower()
            if not st.password or not (uri == st.origin or uri.startswith(st.origin + "/")):
                return  # boshqa manzil yoki parol noma'lum -> WebView2 ning o'z oynasi
            st.auth_tries += 1
            if st.auth_tries > 3:
                return  # parol mos kelmadi -> WebView2 ning o'z parol oynasi
            args.Response.UserName = TITLE
            args.Response.Password = st.password
        except Exception:  # noqa
            logging.exception("Panel paroli berilmadi")

    def on_external(sender, args):
        try:
            uri = str(args.Uri)
            if uri.lower().startswith("tg:"):  # Telegram havolasi — so'rovsiz Telegram ilovasida
                args.Cancel = True
                os.startfile(uri)  # noqa (faqat Windows)
        except Exception as e:  # noqa
            logging.warning("Tashqi havola ochilmadi: %s", e)

    def on_ready(self, sender, args):
        if not args.IsSuccess:
            logging.error("WebView2 ishga tushmadi: %s — panel brauzerda ochiladi", args.InitializationException)
            webbrowser.open(st.url)
            threading.Timer(0.2, lambda: st.window and st.window.destroy()).start()
            return orig(self, sender, args)
        core = sender.CoreWebView2
        try:
            core.BasicAuthenticationRequested += on_auth
        except Exception:  # noqa
            logging.exception("BasicAuthenticationRequested yo'q (eski WebView2) — parol so'raladi")
        try:
            core.LaunchingExternalUriScheme += on_external
        except Exception:  # noqa
            pass
        orig(self, sender, args)
        try:
            s = core.Settings
            s.AreDefaultContextMenusEnabled = True     # o'ng tugma: nusxa olish / qo'yish (bot tokeni va h.k.)
            s.AreBrowserAcceleratorKeysEnabled = True  # F5, Ctrl+F; DevTools o'chiqligicha qoladi
        except Exception:  # noqa
            pass
        try:
            st.hwnd = int(self.form.Handle.ToInt64())
        except Exception:  # noqa
            logging.exception("Oyna HWND olinmadi")

    ec.EdgeChrome.on_webview_ready = on_ready


def run_window(st, storage):
    import webview
    from webview.platforms import winforms  # renderer shu yerda tanlanadi
    if getattr(winforms, "renderer", "") != "edgechromium":
        raise RuntimeError("Edge WebView2 topilmadi (renderer=%s)" % getattr(winforms, "renderer", "?"))
    _patch_webview2(st)
    webview.settings["ALLOW_DOWNLOADS"] = True
    webview.settings["OPEN_EXTERNAL_LINKS_IN_BROWSER"] = True
    st.window = webview.create_window(TITLE, st.url, width=SIZE[0], height=SIZE[1], min_size=MIN_SIZE,
                                      text_select=True, zoomable=True)
    icon = os.path.join(APP_DIR, "branding", "timesaver.ico")
    os.makedirs(storage, exist_ok=True)
    webview.start(private_mode=False, storage_path=storage, icon=icon if os.path.isfile(icon) else None)


def main(argv=None):
    argv = sys.argv[1:] if argv is None else list(argv)
    import timesaver_bot as B
    _setup_log(B.BASE_DIR)
    url = panel_url(_port_from(argv, B.BASE_DIR))
    lock = _bind()
    if lock is None:
        if focus_existing(url):
            return 0
        logging.warning("Oyna porti %d band, lekin javob yo'q — oyna baribir ochiladi", WINDOW_PORT)
    if not FROZEN:
        try:  # manba koddan: vazifalar panelida «Python» emas, TimeSaver ikonkasi bilan alohida tugma
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("TimeSaver.Panel")
        except Exception:  # noqa
            pass
    st = _State(url)
    if lock is not None:
        threading.Thread(target=_serve, args=(lock, st), daemon=True, name="window-lock").start()
    storage = os.path.join(os.environ.get("LOCALAPPDATA") or B.BASE_DIR, "TimeSaver", "WebView2")
    try:
        run_window(st, storage)
    except Exception:  # noqa
        logging.exception("Panel oynasi ochilmadi — panel brauzerda ochiladi")
        webbrowser.open(url)
        return 1
    finally:
        if lock is not None:
            lock.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
