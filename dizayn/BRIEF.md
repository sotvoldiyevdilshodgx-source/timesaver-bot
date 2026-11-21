# TimeSaver — dizayn varianti uchun umumiy topshiriq

## Mahsulot
TimeSaver — dorixonalar tarmog'i uchun Windows dasturi: Telegram-bot + server paneli. Admin (dorixona egasi/menejeri)
panelda: Excel bazani yangilaydi («Обновить»), filiallar o'rtasida dori ko'chirishni nazorat qiladi, foydalanuvchilarni
tasdiqlaydi, hisobotlar yuboradi. Panel endi brauzerda EMAS, alohida Windows oynasida (WebView2) ochiladi.

## Har bir variant nima ko'rsatishi kerak (bitta statik HTML mockup)
1. Windows oynasi ramkasi (~1280×800 ko'rinishda, sahifa markazida): o'z sarlavha satri — logo + «TimeSaver» + oyna tugmalari (—, ▢, ✕).
   Brauzer manzil satri yo'q.
2. Chap menyu (yoki variant g'oyasiga mos navigatsiya), guruhlari va bandlari aynan shu matnlar bilan:
   - (guruhsiz) Бош саҳифа · Кўчириш · Склад · Буюртма
   - Мижозлар: Фойдаланувчилар · Хабар юбориш · Ҳисоботлар
   - Тизим: Статистика · Журнал · Созламалар · Созлаш устаси
   Faol band: Бош саҳифа. «Фойдаланувчилар» yonida qizil nishon «3», «Кўчириш» yonida «82%».
3. Bosh sahifa:
   - Salomlashuv: «Хайрли кун, Намуна Фарм»
   - Yangilanish holati: «Маълумот янгиланди · бугун 08:42», asosiy tugma «Янгилаш» (Обновить), ikkinchi tugma «Техник режим».
   - 4 ta ko'rsatkich: «Срок яқин (90 кун)» 128 · «30+ кун турган» 342 · «Фойдаланувчилар» 57 · «Тасдиқ кутмоқда» 3
   - «Бугунги кўчириш» — 3 filial: ДОРИХОНА ЧИНОЗ 92% (0 кечикди), ДОРИХОНА УРГУТ 76% (2 кечикди), ДОРИХОНА ЭРКИН 64% (4 кечикди)
   - YANGI funksiya ko'rinishi «Срок · ABC таҳлил» kartasi: A / B / C tablari (A faol). A — jami pulning 80% i.
     Jadval (pul bo'yicha kamayish): # · Дори · Муддат · Қолдиқ · Сумма
       1 Аугментин 625мг №14 · 12.01.2027 · 46 · 8 740 000 сўм
       2 Эссенциале форте №30 · 28.12.2026 · 38 · 6 270 000 сўм
       3 Нурофен 200мг №24 · 03.02.2027 · 120 · 4 080 000 сўм
       4 Линекс №16 · 19.01.2027 · 54 · 3 510 000 сўм
       5 Цитрамон П №10 · 07.12.2026 · 300 · 1 650 000 сўм
     Ostida: «Яна 15 та · Excel'да юклаб олиш» tugmasi. Tab nishonlari: A 20 та · 42,6 млн, B 31 та · 8,0 млн, C 77 та · 2,7 млн.
4. Yorug' va qorong'i mavzu: ikkalasi ham chiroyli bo'lsin; o'ng yuqorida almashtirgich (kichik JS, localStorage try/catch bilan).

## Qat'iy talablar
- BUTUN matn — o'zbek tili, KIRILL alifbosi (yuqoridagi matnlar namuna). Lotin harflari faqat «TimeSaver», «Excel», «ABC».
- Bitta o'z-o'zidan ishlaydigan .html fayl: inline CSS/JS. Tashqi manba faqat Google Fonts (kirillni qo'llaydigan shrift:
  masalan Golos Text, Onest, Manrope, Inter, Unbounded, Rubik, Nunito, IBM Plex Sans) — boshqa CDN yo'q, rasm yo'q.
- Ikonkalar — inline SVG (o'zingiz chizing, oddiy chiziqli). Emoji ishlatmang.
- Logo (aynan shu SVG, o'zgartirmang, istalgan o'lchamda):
  <svg viewBox="-85 -85 170 170"><circle r="85" fill="#1F2A6B"/><path d="M0.00 -52.00 A52.0 52.0 0 1 1 -52.00 0.00" fill="none" stroke="#D4F06A" stroke-width="10.0" stroke-linecap="round"/><circle cx="-37.0" cy="-37.0" r="6.0" fill="#FF9E7A"/><rect x="-30.0" y="-27.0" width="60.0" height="18.0" rx="9.0" fill="#FF9E7A"/><rect x="-9.0" y="-27.0" width="18.0" height="56.0" rx="9.0" fill="#FF9E7A"/></svg>
  Brend ranglari: to'q ko'k #1F2A6B, laym #D4F06A, shaftoli #FF9E7A (variant g'oyasiga qarab erkin foydalaning).
- Haqiqiy dasturga o'xshasin: aniq tekislash, o'qiladigan kontrast (WCAG AA), raqamlar tabular-nums, hover holatlari.
- Fayl boshida <title>: «TimeSaver — Вариант N: <nom>».
- Sahifa pastida (oyna tashqarisida) 2–3 qatorli izoh, kirillda: variant g'oyasi, kimga mos, kuchli tomoni.
