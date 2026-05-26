MedMCQ — GitHub Ready Build
Payment Safety Mode مضاف بدون كسر نظام المدار الحالي.

الملفات الأساسية:
- bot.py
- worker.py
- queue_jobs.py
- sms_webhook.py
- payment_watchdog.py
- payment_safety_test.py
- requirements.txt
- runtime.txt
- RUN_THIS_ON_SERVER.txt

ما تم تثبيته في هذه النسخة:
1) sms_webhook.py ما زال يعمل على port 8088.
2) token الافتراضي: MadarPay2026.
3) sender المقبول افتراضيًا: Almadar.
4) رقم التحويل الافتراضي: 0918874659.
5) /sms و /heartbeat مدعومان كما هما.
6) payment_watchdog.py باقٍ لمراقبة heartbeat.
7) /payment_status و /pending_payments باقية.
8) /payment_audit مضاف لعرض آخر 20 SMS من المدار.
9) التفعيل التلقائي صار Atomic Transaction:
   - order لازم pending
   - SMS لازم unclaimed
   - الرقم مطابق
   - المبلغ مطابق
   - SMS داخل الوقت المسموح
   - يتم تفعيل الاشتراك وتغيير order و SMS في نفس transaction
10) أي حالة غير مؤكدة لا تفعّل الطالب ولا ترفضه نهائيًا؛ تبقى pending للأدمن.
11) كل SMS من Almadar تُحفظ في جدول unclaimed_madar_sms:
   id, sms_hash, sender, raw_text, parsed_phone, parsed_amount, received_at, status, claimed_order_id, claimed_user_id, claimed_at, note
12) duplicate SMS يمنع التفعيل لمدة 14 يوم، وبعدها لا يمنع التجديد.
13) التحويل قبل الطلب مدعوم: SMS تبقى unclaimed، وعند إنشاء الطلب خلال 24 ساعة يتم ربطها تلقائيًا.
14) لا يمكن فتح أكثر من طلب pending لنفس الطالب. يظهر زر إلغاء الطلب الحالي.
15) رفع JSON الجاهز لبنك الأسئلة لم يتم كسره.

متغيرات .env المهمة:
TELEGRAM_BOT_TOKEN=ضع_توكن_البوت
MADAR_PAYMENT_NUMBER=0918874659
SMS_WEBHOOK_TOKEN=MadarPay2026
SMS_WEBHOOK_PORT=8088
ALLOWED_SMS_SENDERS=Almadar
SMS_DUPLICATE_TTL_DAYS=14
PAYMENT_PHONE_HEARTBEAT_FILE=data/payment_phone_heartbeat.json
PAYMENT_PHONE_OFFLINE_AFTER_SECONDS=1800
PAYMENT_WATCHDOG_INTERVAL_SECONDS=300
PAYMENT_PENDING_ALERT_AFTER_SECONDS=1800

اختبار الأمان:
بعد تثبيت requirements يمكن تشغيل:
python3 payment_safety_test.py

النتيجة المتوقعة:
RESULT: PASS

للتشغيل على السيرفر:
افتح RUN_THIS_ON_SERVER.txt وانسخ الأمر الواحد كاملًا.
