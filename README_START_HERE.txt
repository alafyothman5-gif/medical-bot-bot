MedMCQ / Medical Bot — GitHub Ready

افتح أولاً ملف RUN_THIS_ON_SERVER.txt وانسخ الأمر الواحد الموجود فيه على السيرفر بعد رفع المشروع إلى GitHub.

الملفات الأساسية:
- bot.py: بوت تيليجرام الرئيسي + لوحة الأدمن + الاشتراكات + Past Papers.
- worker.py: عامل الخلفية للملخصات وبنوك الأسئلة عبر Redis.
- queue_jobs.py: طابور Redis.
- sms_webhook.py: استقبال رسائل المدار من MacroDroid على /sms و /heartbeat.
- payment_watchdog.py: مراقبة هاتف الدفع والتنبيه لو توقف heartbeat.

أهم إصلاح في هذه النسخة:
- زر الأدمن 📤 رفع ملف JSON جاهز أصبح يطلب JSON فقط.
- لا يدخل وضع staging الخام.
- يحفظ الأسئلة مباشرة في جدول past_papers.
- نفس الجدول هو الذي يقرأ منه تدريب عشوائي وحل بالترتيب.
- يدعم الأقسام مثل benghazi / obgyn و ajdabiya / obgyn.
- يوحّد obgyn و Ob/Gyn و ob-gyn و ob_gyn.
- لا يخترع إجابات: أي سؤال بدون answer/correct/key/correct_option لا يتم حفظه.

أشكال JSON المقبولة:
1) قائمة مباشرة:
[
  {"question":"...", "options":{"A":"...","B":"...","C":"...","D":"..."}, "correct":"A"}
]

2) قاموس يحتوي أحد المفاتيح:
- questions
- items
- data
- mcqs

حقول السؤال المقبولة:
- question أو text أو q أو stem أو prompt
- options أو choices أو answers
- answer أو correct_answer أو correct أو correct_option أو key

الاشتراكات:
- Free
- Premium شهري: 20 د.ل لمدة 30 يوم
- Premium فصل كامل / 4 شهور: 45 د.ل لمدة 120 يوم
- Premium Plus فصل كامل / 4 شهور: 65 د.ل لمدة 120 يوم

أوامر الأدمن المهمة:
- /ai_policy
- /set_plan USER_ID premium 30
- /set_plan USER_ID premium 120
- /set_plan USER_ID premium_plus 120
- /remove_premium USER_ID
- /payment_status
- /pending_payments
- /past_stats

إعداد MacroDroid:
SMS payment:
http://SERVER_IP:8088/sms?token=MadarPay2026&sender={sms_number}&text={sms_message}

Payment heartbeat كل 10 دقائق:
http://SERVER_IP:8088/heartbeat?token=MadarPay2026&device=madar_phone&sender=Almadar

متغيرات .env المدعومة للدفع:
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
