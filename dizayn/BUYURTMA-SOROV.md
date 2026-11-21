# 10-talab: firma vakilining buyurtma so'rovi → admin tasdig'i

Foydalanuvchi: «firma panelidagi Буюртма tugmasi endilikda firma xodimlari tugagan dorilarini ko'rib adminga so'rov yuborsin,
masalan "Парацетамол 20 та буюртмага розимисиз?" — admin to'g'ri bo'lsa "Ҳа", 20 ta ko'p bo'lsa sonini 10 taga tushirsin».
Tanlov (17.11.2026): **bitta so'rovda bir nechta dori**; tasdiqdan keyin **xabar + ro'yxat + Excel** (FOM'ga yozilmaydi).

## Firma vakili (bot, ichma-ich navigatsiya)
🏠 › 🧾 Буюртма:
1. Ro'yxat: firmaning **tugagan va tugayotgan** dorilari (avval jami qoldiq 0, keyin ≤ 7 kunga yetadiganlar),
   har biri: nomi · qoldiq (filiallar) · 30 kunlik sotuv · **taklif soni** (= sotuv × cover − qoldiq, yaxlitlangan).
   Har dori — tugma; bosilganda shu dori «savat»ga taklif soni bilan qo'shiladi (✓ belgisi).
2. 🛒 Савата (N та): har dori qatorida son; tugmalar [−5] [−1] [+1] [+5] yoki son yozish; [🗑 Олиб ташлаш].
3. [📤 Админга юбориш] → «⏳ Сўров №27 юборилди · 3 та дори · жами 75 та». Panel bosh menyuga qaytadi.
4. Admin javob berganda push: «✅ Сўров №27 · Парацетамол 500мг №10 — 10 та қабул қилинди (20 дан)…».
5. 🧾 Буюртма › 📋 Сўровларим: holat bilan ro'yxat (⏳ кутмоқда / ✅ қабул / ✏️ ўзгартирилди / ❌ рад).

## Admin (bot)
1. Push karta: «🧾 Буюртма сўрови №27 · АЛЬФА ФАРМ · Анвар Каримов», dorilar: `1. Парацетамол 500мг №10 — 20 та`
   (yonida qisqa asos: «қолдиқ 0 · 30 кунда 46 та сотилди»). Tugmalar:
   [✅ Ҳаммасини қабул қилиш] / [1 · Парацетамол 20] [2 · …] (har dori alohida) / [❌ Рад этиш].
2. Dori bosilganda shu xabar o'zgaradi: «Парацетамол 500мг №10 · сўралди 20 та. Розимисиз?»
   [✅ Ҳа, 20 та] / [10] [15] [−5] [+5] / [✏️ Сонини ёзиш] / [❌ Бу дорини олиб ташлаш] / [◀ Орқага].
3. Hamma dori hal bo'lgach: [📤 Жавобни юбориш] → firma vakiliga xabar, karta bir qatorga qisqaradi:
   «✅ №27 · 3 та дори · 2 та ўзгартирилди · Сиз, 11:05».
4. Bir nechta admin bo'lsa, birinchi javob bergani hal qiladi; boshqalarda karta qisqaradi.

## Server paneli
«Буюртма» bo'limida «Сўровлар» jadvali: №, firma, vakil, dorilar soni, so'ralgan/tasdiqlangan son, holat, vaqt;
qatorni ochib xuddi botdagidek tasdiqlash/son o'zgartirish; tasdiqlanganlarni Excel'ga yuklab olish (kun/hafta/firma bo'yicha).

## Saqlash
Yangi jadvallar: `order_req(id, chat_id, firm, created_at, status, resolved_by, resolved_at)` va
`order_req_items(req_id, name, producer, qty_asked, qty_ok, status)`.
