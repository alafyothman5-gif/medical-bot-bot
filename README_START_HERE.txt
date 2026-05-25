MedMCQ Bot - SAFE AI ROUTER VERSION

ما الجديد:
- الطلاب العاديون: كاش أولاً، ثم Google، ثم OpenRouter Free، ثم Groq Free.
- OpenRouter المدفوع لا يعمل للطلاب العاديين افتراضياً.
- OpenRouter المدفوع يعمل للأدمن فقط.
- Premium مدعوم لاحقاً عبر PREMIUM_ENABLED=true و /set_plan USER_ID premium.
- أمر أدمن جديد: /ai_policy يعرض حالة المفاتيح والسياسة.
- أمر أدمن جديد: /set_plan USER_ID premium أو free.

إعدادات .env المقترحة:

# Google مفاتيح متعددة، افصل بينها بفاصلة أو سطر جديد
GOOGLE_API_KEYS=KEY1,KEY2,KEY3

# OpenRouter مجاني: يستعمل موديل مجاني فقط
OPENROUTER_FREE_KEYS=YOUR_OPENROUTER_KEY
OPENROUTER_FREE_MODEL=openrouter/free

# Groq مجاني اختياري
GROQ_API_KEYS=GROQ_KEY1,GROQ_KEY2
GROQ_MODEL=llama-3.1-8b-instant

# OpenRouter المدفوع للأدمن فقط
OPENROUTER_PAID_KEYS=YOUR_PAID_OPENROUTER_KEY
GEMINI_MODEL=google/gemini-2.5-flash-lite

# حماية التكلفة
FREE_PROVIDER_MODE=1
ALLOW_PAID_FALLBACK_FOR_USERS=0
ALLOW_PAID_FALLBACK_FOR_ADMIN=1
PREMIUM_ENABLED=0
ALLOW_PAID_FALLBACK_FOR_PREMIUM=1

طريقة التأكد بعد التشغيل:
/ai_policy

ترتيب الاستخدام:
1) الكاش والمخزون بدون API
2) Google direct
3) OpenRouter Free
4) Groq Free
5) OpenRouter Paid فقط للأدمن أو Premium عند تفعيله
