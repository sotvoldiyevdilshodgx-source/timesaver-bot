# -*- coding: utf-8 -*-
"""Ko'chirish dashboardini ko'rish uchun namunaviy panel (vaqtinchalik baza, soxta Telegram, test Exceli):
    python tests/serve_demo.py [port] [yangi]
Haqiqiy bot.db / config.ini / db.ini / Excel / Telegram / panel_port.txt ga tegmaydi. Panel: http://127.0.0.1:8096/#koch
«yangi» — birinchi o'rnatish (sozlash ustasi): token, admin va Excel yo'q. Soxta Telegram: har qanday to'g'ri shakldagi
token qabul qilinadi (999999… bilan boshlangani — «Telegram qabul qilmadi»), token kiritilgach 6 soniyadan keyin
«Akmal» botga /start yozgandek bo'ladi. «Yangilash» sinovi Excelsiz, soxta makros bilan o'tadi."""
import datetime
import os
import sys
import threading
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fixtures import B, CH, UCH, EK, F_CH, F_EK, F_UCH, make_bot, add_client  # noqa: E402
import timesaver_panel as P  # noqa: E402

port = int(sys.argv[1]) if len(sys.argv) > 1 else 8096
fresh = "yangi" in sys.argv[2:]
bot, tg, work = make_bot("new")
B.BASE_DIR = work  # config.ini / baza\ yuklash — vaqtinchalik papkada
B.DB_INI = os.path.join(work, "BazaReport", "db.ini")
with open(os.path.join(work, "config.ini"), "w", encoding="utf-8") as f:
    f.write("[app]\nname =\nsetup_done = %s\n[bot]\ntoken = %s\nadmin_ids = %s\n[excel]\nfile = %s\n" % (
        "no" if fresh else "yes", "" if fresh else "1:x", "" if fresh else "9001", bot.cfg["excel"]["file"]))
staff = {CH: F_CH, EK: F_EK, UCH: F_UCH}
for f, cid in staff.items():
    add_client(bot, cid, "filial", filial=f, name="Xodim " + f.split()[-1].title())
today = datetime.date.today()
real_now = B.now
at = lambda d, h, m: (lambda: datetime.datetime.combine(d, datetime.time(h, m)))
for back in range(6, -1, -1):  # 7 kunlik tarix, bugun — qisman bajarilgan
    d = today - datetime.timedelta(days=back)
    B.now = at(d, 8, 45)
    bot.koch_build(d.isoformat())
    for k, t in enumerate(bot.koch_tasks(d.isoformat())):
        c, r = bot.client(staff[t["src"]]), bot.client(staff[t["dst"]])
        if back == 0 and k % 3:
            continue
        B.now = at(d, 10, 25) if (back + k) % 4 == 0 else at(d, 9, 20)
        if k == 1 and back % 2:
            bot.koch_refuse(c["chat_id"], c, t["id"], "Tovar topilmadi")
            continue
        bot.koch_give(c["chat_id"], c, t["id"], t["qty_plan"] if k else max(1, t["qty_plan"] - 1))
        if back or k == 0:
            bot.koch_receive(r["chat_id"], r, t["id"])
B.now = real_now

ME = {"id": 7000001, "is_bot": True, "username": "shifo_farm_bot", "first_name": "Шифо Фарм"}


class DemoTG:
    """Token tekshiruvi (getMe) uchun soxta Telegram."""

    def __init__(self, token, base=""):
        self.token = token

    def call(self, method, params=None, timeout=60):
        if self.token.startswith("999999"):
            raise B.TGError('HTTP 401: {"ok":false,"error_code":401,"description":"Unauthorized"}', 401)
        return dict(ME)


def demo_runner(vbs, path, timeout):
    """Soxta makros: refresh_status.txt ga varaqlarni birma-bir yozadi."""
    status = os.path.join(os.path.dirname(path), bot.cfg.get("refresh", "status_file", fallback="refresh_status.txt"))
    d = bot.data
    with open(status, "w", encoding="utf-8") as f:
        for name, n in (("Sale", len(d.sale)), ("Srok", len(d.srok)), ("30+ kun", len(d.old)), ("Narx", len(d.narx)),
                        ("Hisob", len(d.hisob))):
            time.sleep(1.2)
            f.write("OK %s|%d\n" % (name, n))
            f.flush()
        f.write("DONE\n")
    return 0, ""


if fresh:
    B.TG = DemoTG
    tg.call = lambda method, params=None, timeout=60: dict(ME) if method == "getMe" else True
    bot.cfg.read_dict({"app": {"name": "", "setup_done": "no"}, "bot": {"token": "", "admin_ids": ""}})
    bot.admins = set()
    xlsx = bot.cfg["excel"]["file"]
    bot.cfg["excel"]["file"] = os.path.join(work, "baza", "baza_yangi.xlsm")  # hali yuklanmagan
    bot.data.reload(force=True)
    bot.runner = demo_runner

    def fake_start():  # token ustada kiritilgach «Akmal» botga /start yozadi
        while True:
            time.sleep(1)
            if bot.setting(P.SETUP_TOKEN_AT) and not bot.admins:
                time.sleep(6)
                bot.handle({"message": {"message_id": 1, "date": int(time.time()), "text": "/start admin",
                                        "chat": {"id": 1112223334, "type": "private"},
                                        "from": {"id": 1112223334, "first_name": "Акмал", "last_name": "Каримов",
                                                 "username": "akmal_k"}}})
                return
    threading.Thread(target=fake_start, daemon=True).start()
    print("Yangi o'rnatish rejimi. Excel fayl yo'li (2-qadam, «Ёки компьютердаги файл йўли»): %s" % xlsx, flush=True)

srv = P.PanelServer(("127.0.0.1", port), P.make_handler(P.Api(bot, bot.cfg), ""))
print("Demo panel: http://127.0.0.1:%d/#%s  (vaqtinchalik papka: %s)" % (port, "setup" if fresh else "koch", work), flush=True)
try:
    srv.serve_forever()
finally:
    srv.server_close()
