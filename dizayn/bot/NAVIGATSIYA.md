# Bot navigatsiyasi: ichma-ich, bitta «tirik» xabar (foydalanuvchi talabi, 17.11.2026)

Foydalanuvchi: «tugmalar chatga chiqmasin, ichma-ich kirib ketsin».

## Qoidalar
1. **Pastki klaviatura (reply keyboard) YO'Q.** Tugma bosilganda chatga «⏰ Срок · 134» kabi xabar tushmaydi.
   Pastda faqat oddiy matn maydoni (`.input`) va chapda Telegram'ning «☰ Меню» tugmasi (bot buyruqlari: /start).
2. **Bitta panel xabari.** /start → bot bitta xabar yuboradi (Бош меню). Har bir tugma bosilganda **shu xabarning
   o'zi** o'zgaradi (editMessageText): matn ham, ostidagi tugmalar ham. Chatda yangi xabar paydo bo'lmaydi.
3. **Ichma-ich daraxt.** Har xabar boshida «yo'l» (breadcrumb) qatori: `🏠 › ⏰ Срок › A` (`.crumb`).
   Har darajaning oxirgi qatorida: `◀ Орқага` · `🏠 Меню`. Chuqurlik 3–4 darajadan oshmasin.
4. **Foydalanuvchining xabarlari** faqat: `/start`, qidiruv uchun yozilgan matn (dori nomi), son kiritish (masalan, «Бошқа сон»).
5. **Alohida xabar faqat zarur bo'lganda:** Excel fayl (hujjat), bot o'zi yuboradigan bildirishnomalar (ertalabki topshiriq,
   yangi so'rov adminga, eslatma). Bunday xabar ostida «🏠 Менюга қайтиш» tugmasi bo'ladi; panel xabari esa yangidan pastga
   tushadi (eski panel o'chiriladi), shunda chatda doim bitta panel bo'ladi.
6. Maketda o'tishni ko'rsatish: ketma-ket telefonlar = bitta xabarning holatlari. Bosilgan tugma `.ib.pressed` (laym chegara)
   bilan belgilanadi, telefon ustidagi izohda «… bosildi → shu xabar o'zgardi» deb yoziladi.

## Qo'shimcha CSS (shablon oxiriga qo'shing)
.crumb{font-size:12.5px;color:var(--meta);margin-bottom:4px;display:block}
.ib.pressed{box-shadow:inset 0 0 0 2px #D4F06A}
.menu-btn{background:var(--ib-primary);color:#fff;border-radius:8px;padding:5px 10px;font-size:13px;font-weight:600}
.tree{max-width:1400px;margin:0 auto 6px;padding:0 16px}
.tree pre{background:var(--head);color:var(--head-tx);border:1px solid var(--sep);border-radius:12px;padding:14px 16px;font-size:13px;line-height:1.55;overflow:auto;margin:0}
