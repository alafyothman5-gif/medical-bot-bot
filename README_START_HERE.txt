MedMCQ / Medical Bot — GitHub Ready

افتح أولاً RUN_THIS_ON_SERVER.txt وانسخ الأمر الواحد الموجود فيه على السيرفر بعد رفع المشروع إلى GitHub.

الملفات داخل هذا ZIP:
- bot.py
- worker.py
- queue_jobs.py
- sms_webhook.py
- payment_watchdog.py
- requirements.txt
- runtime.txt
- README_START_HERE.txt
- RUN_THIS_ON_SERVER.txt

الجديد في هذه النسخة:
1) واجهة Dashboard جديدة في /start:
   - خطة الطالب
   - تاريخ انتهاء الاشتراك
   - استخدام بنك السنوات السابقة اليوم
   - الجامعة والمادة الحالية
   - أزرار واضحة للتدريب، البنك، الملخص، OSCE، التقدم، الاشتراك، الدعم، ولوحة الأدمن.

2) بنك السنوات السابقة أصبح Premium مع تجربة مجانية:
   - Free: يسمح له بـ 20 سؤال يوميًا فقط من بنك السنوات السابقة.
   - Premium و Premium Plus: بدون حد.
   - العداد مستقل في SQLite داخل user_daily_usage حسب user_id + day + kind=past_paper_question.
   - العد يتم عند عرض السؤال، وليس فقط دخول الصفحة.

3) صفحة تدريب السنوات السابقة صارت مرتبة:
   - عدد الأسئلة المتاحة
   - محاولات الطالب
   - الصحيحة والخاطئة والدقة
   - أزرار: اختبار سريع 20، تدريب عشوائي، حل بالترتيب، أسئلتي الخاطئة، تحليلي.

4) صفحة الاشتراك صارت متجر واضح:
   - Premium شهري 20 د.ل / 30 يوم
   - Premium فصل كامل 45 د.ل / 4 شهور
   - Premium Plus 65 د.ل / 4 شهور
   - توضّح أن بنك السنوات السابقة كامل ضمن Premium.

5) الملخصات صارت منظمة:
   - عنوان واضح
   - الفكرة العامة
   - أهم التعريفات
   - النقاط الامتحانية الذهبية
   - التشخيص/Management إن وجد
   - أخطاء شائعة
   - خلاصة وأسئلة مراجعة.
   - الرسائل الطويلة تنقسم لأجزاء آمنة لتجنب حدود تيليجرام.

6) زر 🧠 ماذا أدرس الآن؟:
   - يعتمد على أضعف قسم حسب دقة الطالب وأخطائه ومحاولاته.
   - يعطي خطة: راجع الأخطاء، اختبر سريع، ثم انتقل عند 80%.

لم يتم كسر هذه الميزات:
- الدفع شبه التلقائي بالمدار.
- SMS webhook على port 8088.
- token الافتراضي: MadarPay2026.
- sender المقبول: Almadar.
- رقم التحويل: 0918874659.
- منع تكرار SMS لمدة 14 يوم.
- حفظ تحويلات المدار غير المرتبطة لمدة 24 ساعة، وربطها تلقائيًا إذا فتح الطالب طلب اشتراك بعدها.
- heartbeat لهاتف الدفع.
- /payment_status
- /pending_payments
- /my_plan
- /set_plan
- /remove_premium

رفع JSON جاهز:
- زر الأدمن: 📤 رفع ملف JSON جاهز
- يختار الأدمن الجامعة والمادة.
- يرسل ملف JSON.
- يتم التخزين مباشرة في جدول past_papers، وهو نفس البنك الذي يقرأ منه الطالب.
- لا يدخل في staging الخام.
- رسالة النجاح تعرض القسم، الجديد المضاف، وإجمالي أسئلة القسم.

أشكال JSON المقبولة:
1) قائمة مباشرة: [ {...}, {...} ]
2) قاموس يحتوي أحد المفاتيح:
   - questions
   - items
   - data
   - mcqs

حقول السؤال المقبولة:
- question أو text أو q أو stem أو prompt
- options أو choices أو answers
- answer أو correct_answer أو correct أو correct_option أو key

تنبيه مهم:
- البوت لا يخترع إجابات عند رفع JSON جاهز.
- أي سؤال بدون إجابة صريحة لا يتم حفظه.

متغيرات .env المهمة للدفع:
MADAR_PAYMENT_NUMBER=0918874659
SMS_WEBHOOK_TOKEN=MadarPay2026
SMS_WEBHOOK_PORT=8088
ALLOWED_SMS_SENDERS=Almadar
SMS_DUPLICATE_TTL_DAYS=14
PAYMENT_PHONE_HEARTBEAT_FILE=data/payment_phone_heartbeat.json
PAYMENT_PHONE_OFFLINE_AFTER_SECONDS=1800
PAYMENT_WATCHDOG_INTERVAL_SECONDS=300
PAYMENT_PENDING_ALERT_AFTER_SECONDS=1800

لا تضع مفاتيح API داخل GitHub. ضعها في .env على السيرفر فقط.
