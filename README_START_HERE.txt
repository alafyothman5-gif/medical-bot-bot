MedMCQ Bot - Premium + Support + Quick Past Papers Final Ready

التغييرات المدمجة:
- أسعار الاشتراك: Premium شهري 20 د.ل، Premium فصل كامل 45 د.ل، Premium Plus فصل كامل 65 د.ل.
- صفحة الاشتراكات تعرض الخطة الحالية + User ID + رقم المدار من .env.
- الطالب يختار الخطة ثم تظهر له رسالة طلب اشتراك جاهزة فيها User ID ورقم المدار وأمر التفعيل للأدمن.
- صفحة الدعم الجديدة: تواصل مع الدعم / الإبلاغ عن مشكلة / رجوع.
- حذف معرف الدعم القديم Othma2003 من الرسائل.
- أوامر الاشتراك: /my_plan و /set_plan و /remove_premium.
- حدود يومية مختلفة لـ Free و Premium و Premium Plus.
- Quick Past Papers: زر ⚡ جمع سنوات سابقة سريع + /finish_past.
- Free router: Google -> OpenRouter Free -> Groq -> Cloudflare، والـ Paid محمي للأدمن/بريميوم حسب الإعداد.

مهم بعد الرفع للسيرفر:
1) أضف رقم المدار في .env:
MADAR_PAYMENT_NUMBER=09XXXXXXXX

2) ثم شغل:
cd /root/medical_bot && git fetch origin main && git reset --hard origin/main && pip install -r requirements.txt --break-system-packages && pkill -9 -f "bot.py|worker.py|venv/bin/python" 2>/dev/null || true; rm -f nohup.out worker.log; nohup python3 bot.py > nohup.out 2>&1 & nohup python3 worker.py > worker.log 2>&1 & sleep 8; tail -60 nohup.out; tail -40 worker.log

أوامر التفعيل:
/set_plan USER_ID premium 30
/set_plan USER_ID premium 120
/set_plan USER_ID premium_plus 120
/remove_premium USER_ID
