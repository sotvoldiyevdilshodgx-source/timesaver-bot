# 9-talab: ko'chirishni qoldiqdan avtomatik aniqlash

Foydalanuvchi: «tizim filial dorini berdimi yoki yo'q — beruvchi filialda qoldiq kamayadi, oluvchida ko'payadi — shundan aniqlab olsin».
Tanlov (17.11.2026): **avtomatik + tugma zaxira sifatida**; tekshiruv **har 1 soatda** (ish vaqti 08:00–20:00).
Tuzatish (18.11.2026): yolg'iz beruvchidagi kamayish endi holatni o'zgartirmaydi; «Бердим» paytidagi oluvchi qoldig'i va 3 kunlik oyna;
har oluvchi o'zi bo'yicha; ogohlantirishlar bitta yig'ma xabarda; aniqlash qayta shakllantirishdan oldin.

## Qanday ishlaydi
1. **Boshlang'ich surat.** Topshiriq yaratilganda beruvchi va oluvchidagi qoldiq saqlanadi
   (`src_stock0`, `dst_stock0` — yangi ustunlar; mavjud `src_stock/dst_stock` «Qayta shakllantirish»da yangilanib ketadi, shuning uchun alohida).
   Xodim «✅ Бердим» ni bosganda oluvchining **o'sha paytdagi** qoldig'i ham saqlanadi (`dst_stock1`, botda yuklangan Sale ma'lumotidan).
2. **Har yangilashda** (har soat avtomatik «Обновить» yoki qo'lda) Sale varag'idan shu dori (nom + ishlab chiqaruvchi) bo'yicha
   har filialdagi hozirgi qoldiq olinadi va topshiriqdagi suratga solishtiriladi:
   - `kamaydi = src_stock0 − src_hozir` (beruvchida);
   - `ko'paydi = dst_hozir − dst_stock0` (oluvchida; «berildi» topshiriqda — `dst_hozir − dst_stock1`, u bo'lmasa `dst_stock0`).
   **Tartib:** avval aniqlash, keyin topshiriqni yaratish / qayta shakllantirish — bot yangilashidan keyin ham, Excelda qo'lda
   yangilangandan keyingi sinxronlashda ham, panel «Qayta shakllantirish» da ham. Aks holda makros ko'chirilgan qatorni ro'yxatdan
   chiqaradi va u aniqlanmasdan «bekor» bo'lib ketadi.
3. **Qaror** (`q` — reja bo'yicha son, «berildi» da — berilgan son):
   | Beruvchida | Oluvchida | Natija |
   |---|---|---|
   | kamaydi ≥ q×0,8 | ko'paydi ≥ q×0,8 | ✅ **олинди (тизим)** — son = min(kamaydi, ko'paydi, q) |
   | kamaydi ≥ q×0,8 | o'zgarmadi | holat **o'zgarmaydi** (faqat ishora: «Бердим» dan keyingi «қолдиқ ўзгармади» ogohlantirishi chiqmaydi) |
   | o'zgarmadi | ko'paydi ≥ q×0,8 | ⚠ **текширинг** — boshqa joydan kelgan bo'lishi mumkin (adminga) |
   | o'zgarmadi | o'zgarmadi | holat o'zgarmaydi |
   Beruvchidagi kamayishni oddiy sotuvdan ajratib bo'lmaydi, shuning uchun tizim faqat **ikkala tomonda ham o'zgarish** bo'lsa
   holatni o'zgartiradi (→ «olindi»). «Берилди (тизим)» endi qo'yilmaydi; topshiriq «yangi» qoladi va keyingi yangilashlarda
   oluvchida ko'payish ko'rinsa «olindi» bo'ladi.
4. **«Берилди» topshiriqlar** (xodim «Бердим» bosgan) faqat **«Бердим» dan 3 kun ichida** avtomatik yopiladi (`KOCH_AUTO_DAYS`).
   Keyin oluvchidagi ko'payish yetkazib beruvchidan bo'lishi mumkin — «📥 Олдим» tugmasi qoladi (14 kun).
5. **Bir xil dori bir nechta topshiriqda.**
   - Bitta beruvchidan bir nechta oluvchiga: har topshiriq **o'z oluvchisi** bo'yicha baholanadi. Topshiriq «olindi» bo'ladi, agar
     o'z oluvchisida ko'payish ≥ rejaning 80% i VA beruvchidagi jami kamayish ≥ oluvchisi ko'paygan topshiriqlar rejalari
     yig'indisining 80% i bo'lsa. Kamayish hammasiga yetmasa — rejasi kattasidan boshlab sig'ganicha «olindi», qolgan ko'paygan
     topshiriqlar «текширинг». Oluvchisi ko'paymagan topshiriq qo'shnisi tufayli ogohlantirilmaydi.
   - Bitta oluvchiga bir nechta beruvchidan: oluvchidagi ko'payish topshiriqlar orasida reja soniga mutanosib bo'linadi.
6. **Qayta shakllantirish** «yangi» topshiriqni bekor qilmaydi va rejasini kamaytirmaydi, agar surat olinganidan beri beruvchida
   kamayib, oluvchida ko'paygan bo'lsa (ko'chirish izi) yoki tizim uni «текширинг» deb belgilagan bo'lsa — buni aniqlash hal qiladi.
7. **Tugmalar zaxira.** «✅ Бердим» / «📥 Олдим» qoladi. Tizim aniqlaganda xodimga bosish shart emas.
   Xodim bosgan, lekin qoldiq 2 ta yangilashdan keyin ham o'zgarmagan bo'lsa → adminga «⚠ Текширинг: … Бердим деди, қолдиқ ўзгармади».
8. **Ogohlantirishlar yig'ma.** Bir yangilashdagi hamma «⚠ Текширинг» lar har adminga **bitta xabarda**: ko'pi bilan 15 qator,
   qolgani «… яна N та». Har topshiriq uchun bir marta. Yangilashdan keyingi ishlar (aniqlash + qayta shakllantirish) tugaguncha
   yangilash «ketmoqda» hisoblanadi — ustma-ust ishga tushib, ogohlantirishni ikki marta yubormaydi.
9. **Ko'rinishi.** Panel va botda holat yonida manba: «олинди 10:05 · 🤖 тизим» yoki «олинди 09:14 · 👤 Азиз».
   Panel qatorida isbot: «ЧИНОЗ −12 · УРГУТ +12».
10. **Har soatlik yangilash** Sozlamalarda yoqiladi/o'chiriladi (oraliq: 30 / 60 / 120 daqiqa, ish vaqti chegarasi).
    Yangilash ketayotganda bot javob berishda davom etadi; ertalabki topshiriq vaqti o'zgarmaydi.

## Cheklovlar
- Aniqlik Excel (SQL) dagi qoldiq qanchalik tez yangilanishiga bog'liq: FOM'da hujjat o'tkazilmaguncha qoldiq o'zgarmaydi.
- Oluvchi o'sha soatda shu dorini sotib yuborsa, ko'payish kichik ko'rinadi — 80% chegara shuni hisobga oladi.
- Beruvchi hujjatni o'tkazgan, oluvchi hali kiritmagan bo'lsa, topshiriq oluvchi kiritguncha «yangi» ko'rinadi (sotuvdan ajratib
  bo'lmaydi); xodim «✅ Бердим» ni bossa — «берилди».
- `dst_stock1` botda yuklangan oxirgi ma'lumotdan olinadi: «Бердим» dan oldin kirgan, lekin hali yangilanmagan kirim keyingi
  yangilashda ko'payish bo'lib ko'rinishi mumkin (3 kunlik oyna va beruvchidagi kamayish sharti xavfni kamaytiradi).
