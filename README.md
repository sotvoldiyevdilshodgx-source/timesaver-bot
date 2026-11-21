# TimeSaver

Dorixonalar tarmog'i uchun Windows dasturi: **Telegram-bot** va **server paneli**.
Hisob dasturining (FOM, SQL Server) ma'lumotini Excel makrosi orqali oladi va firmalarga, filiallarga va adminga
kerakli hisobotlarni botda ko'rsatadi.

## Imkoniyatlar

**Telegram-bot** (bitta panel xabari, ichma-ich inline menyular):
- ⏰ **Srok**: muddati yaqin dorilar, firmalarga har kuni avtomatik yuborish
- 📦 **30+ kun**: uzoq turib qolgan tovarlar
- 🧾 **Buyurtma**: sotuv va qoldiq asosida tavsiya, firma vakilidan buyurtma so'rovi
- 💰 **Hisob**: debet / kredit
- 📊 **Sotuv va qoldiq**: filiallar bo'yicha, dori qidirish (kirill / lotin)
- 🔁 **Filiallararo ko'chirish**: Shoshilinch / Assortiment / Oddiy toifalari, topshiriq, nazorat va eslatmalar
- 🏭 **Sklad taqsimoti**, FOM'ga qoralama bron yozish
- 👥 Rollar: admin, xodim, firma vakili, filial xodimi

**Server paneli** (WebView2 oynasi):
- bosh sahifa: yangilash holati, Srok va 90+ kun (dori bo'yicha jamlangan) ABC tahlili, bugungi ko'chirish
- foydalanuvchilar, xabar yuborish, hisobotlar (Excel), statistika, jurnal
- 6 qadamli sozlash ustasi (bot tokeni, SQL baza, Excel fayl)

## Tuzilishi

```
app/
  timesaver_bot.py      bot: ma'lumot (Excel), menyular, ko'chirish, rejalashtirilgan ishlar
  timesaver_panel.py    server paneli (HTTP API) + panel.html
  timesaver_window.py   panel oynasi (pywebview)
  timesaver_tray.py     tray belgisi va ishga tushirish
  fom_sql.py            FOM bazasiga qoralama yozish
  vba/                  Excel makroslari (Module1.bas, Kochirish.bas)
  tests/                unittest testlari
installer/              PyInstaller spec va Inno Setup skripti
build.ps1               setup.exe yig'ish
dizayn/                 dizayn topshiriqlari va maketlar
```

## Ishga tushirish (ishlab chiquvchi uchun)

Python 3.11+ kerak.

```bash
pip install openpyxl pystray pillow pyodbc pywebview
```

```bash
cd app && python -m unittest discover -s tests -p "test_*.py"
```

`setup.exe` yig'ish (Windows, Inno Setup 6 o'rnatilgan bo'lishi kerak):

```powershell
.\build.ps1
```

Sozlamalar `config.ini` faylida saqlanadi (namuna: `app/config.example.ini`). Bot tokeni, SQL parol va
mijoz ma'lumotlari repoga kirmaydi.

> Kod va testlardagi dorixona, filial, server nomlari va raqamlar namunaviy (haqiqiy emas).
