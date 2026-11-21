# TimeSaver Telegram-bot paneli — dizayn topshirig'i

Bot dorixonalar tarmog'i uchun. Uch xil foydalanuvchi bor: **mijoz (firma vakili)**, **filial xodimi**, **admin**.
Vazifa: botdagi panellar (xabarlar, pastki tugmalar, xabar ostidagi tugmalar) uchun yangi, zamonaviy, toza
dizaynni Telegram ko'rinishidagi telefon maketlarida chizish. Server paneli uchun tanlangan dizayn — «Тун»
(to'q ko'k #1F2A6B, laym #D4F06A, shaftoli #FF9E7A); botda shu ruhni emoji va rangli tugmalar orqali saqlang.

## Shablon (MAJBURIY)
`dizayn/bot/_shablon.html` faylini to'liq nusxa qiling (CSS, mavzu almashtirgich, logo-avatar, komponentlar)
va faqat ekranlarni (`<div class="screen">…`) qo'shing. CSS'ni o'zgartirmang; juda zarur bo'lsa, oxiriga
bir nechta qo'shimcha qoida yozish mumkin. Komponentlar:
- `.msg` — botning xabari, `.msg.out` — foydalanuvchining xabari (tugma bosilishi ham shunday ko'rinadi), `.t` — vaqt.
- `.wrap` > `.msg` + `.ikb` — xabar va uning ostidagi tugmalar (inline). `.ib.primary/.success/.danger` — rangli tugma.
- `.rkb` — pastki doimiy klaviatura (`.kb`, rangli: `.kb.danger/.success/.primary`). `.input` — klaviatura o'rniga matn maydoni.
- `.doc` — Excel fayl xabari. `.q` — ajratilgan blok (blockquote). `.day` — sana belgisi.
Telegram'da mumkin bo'lmagan narsa chizmang: xabar ichida faqat matn (qalin, kursiv, `code`, blockquote, havola),
emoji va yangi qatorlar. Rangli tugmalar mumkin (Bot API style: danger / success / primary).

## Matn uslubi (hamma rol uchun bir xil)
- HAMMA matn — o'zbek tili, KIRILL alifbosi. Lotin faqat: TimeSaver, Excel, ABC, A/B/C.
- Xabar boshi: `emoji <b>Сарлавҳа</b> · контекст`, keyingi qatorda kulrang (`.muted`) qisqa izoh.
- Raqamlar qalin. Pul: `8 740 000 сўм`, katta summa: `42,6 млн сўм`. Sana: `12.01.2027`. Miqdor: `46 та`.
- Ro'yxat bandi 2 qator: `1. <b>Аугментин 625мг №14</b>` / `<span class="muted">46 та · 12.01.2027 (56 кун)</span> · <b>8 740 000</b>`.
- Oxirgi qator: `🕒 Маълумот: бугун 08:42`.
- Xabar ostidagi tugmalar: qatorda ko'pi bilan 3 ta; asosiy amal rangli; navigatsiya qatori oxirida: `◀ Орқага` · `🏠 Меню`.
- Pastki klaviatura tugmalarida son bo'lsa: `⏰ Срок · 18`.
- Qisqa, mobil ekranga sig'adigan qatorlar. Ortiqcha emoji yo'q: har bo'limning bitta doimiy emojisi bor:
  ⏰ Срок · 📦 30+ кун · 🧾 Буюртма · 💰 Ҳисоб · 📊 Қолдиқ · 🔴 ШОШИЛИНЧ · 🟡 АССОРТИМЕНТ · 🟢 ОДДИЙ ·
  📥 Келадиган · 📋 Бугунги ҳолат · 🏠 Меню · ⚙️ Админ панел · 🔄 Янгилаш · 👥 Фойдаланувчилар · 📢 Хабар юбориш · 🏢 Фирма

## YANGI funksiya (filial xodimi uchun, albatta chizilsin)
Filial menyusiga «⏰ Срок» va «📦 30+ кун» tugmalari qo'shiladi. Bosilganda keyingi xabar: A / B / C bo'limlari.
A — jami pulning 80% ini tashkil qiladigan eng qimmat dorilar, B — keyingi 15%, C — qolgan 5%. Pul = qoldiq × narx.
Bo'lim tanlanganda: pul bo'yicha kamayish tartibida **20 tasi** xabarda, qolgani **Excel fayl** bo'lib yuboriladi
(«📥 Қолган 6 тасини Excel'да олиш» tugmasi). 30+ kun bo'limi ham xuddi shu mantiqda.
Namuna ma'lumot (ДОРИХОНА ЧИНОЗ): Срок — A 26 та · 42,6 млн, B 31 та · 8,0 млн, C 77 та · 2,7 млн;
30+ кун — A 38 та · 61,2 млн, B 64 та · 11,5 млн, C 240 та · 3,8 млн.
Dorilar: Аугментин 625мг №14, Эссенциале форте №30, Нурофен 200мг №24, Линекс №16, Цитрамон П №10,
Флемоксин солютаб 500мг №20, Но-шпа 40мг №24, Мезим форте №20, Терафлю №10, Арбидол 100мг №20.

## Fayl
- Fayl boshida `<title>TimeSaver бот — <rol></title>`, `.top h1` da rol nomi, pastda `.notes` da 3–4 qatorli izoh (kirillda).
- Ekranlar 360×740 telefon ichida; matn telefondan toshib chiqmasin (chat pastdan yuqoriga to'ladi, eskisi yuqorida kesiladi — bu normal).
- Yorug' va qorong'i mavzuda ham tekshiring (yuqoridagi «Кун / Тун» tugmasi).
