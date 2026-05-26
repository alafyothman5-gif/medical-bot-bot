MedMCQ Bot - Premium + SMS Auto Payment + Quick Past Papers

التغييرات المدمجة:
- أسعار الاشتراك: Premium شهري 20 د.ل، Premium فصل كامل 45 د.ل، Premium Plus فصل كامل 65 د.ل.
- الطالب يختار الخطة ثم يكتب رقم المدار الذي سيحوّل منه: 091xxxxxxx أو 093xxxxxxx.
- البوت ينشئ طلب دفع وينتظر رسالة المدار الرسمية التي تصل إلى هاتف الأدمن.
- ملف sms_webhook.py يستقبل رسالة المدار من الهاتف عبر MacroDroid/SMS Forwarder.
- البوت يطابق: رقم المحوّل + المبلغ + وقت الطلب، ثم يفعّل الاشتراك تلقائيًا.
- يدعم 091 و093 وصيغ 21891 و21893.
- إذا الرسالة غير مفهومة أو لا يوجد طلب مطابق، لا يرفض الطالب تلقائيًا؛ يرسلها للمراجعة.
- صفحة الدعم والاشتراكات جاهزة، ورقم المدار الافتراضي: 0918874659.
- Quick Past Papers: زر ⚡ جمع سنوات سابقة سريع + /finish_past.
- Free router: Google -> OpenRouter Free -> Groq -> Cloudflare، والـ Paid محمي للأدمن/بريميوم حسب الإعداد.

بعد رفع الملفات إلى GitHub شغّل هذا الأمر على السيرفر:

cd /root/medical_bot && \
git fetch origin main && \
git reset --hard origin/main && \
python3 << 'PY'
from pathlib import Path
import secrets
p = Path('.env')
s = p.read_text(encoding='utf-8', errors='ignore') if p.exists() else ''
updates = {
    'MADAR_PAYMENT_NUMBER': '0918874659',
    'SMS_WEBHOOK_TOKEN': secrets.token_urlsafe(24),
    'SMS_WEBHOOK_PORT': '8088',
}
lines = []
seen = set()
for line in s.splitlines():
    if '=' in line:
        k = line.split('=', 1)[0].strip()
        if k in updates:
            lines.append(f'{k}={updates[k]}')
            seen.add(k)
        else:
            lines.append(line)
    else:
        lines.append(line)
for k, v in updates.items():
    if k not in seen:
        lines.append(f'{k}={v}')
p.write_text('\n'.join(lines) + '\n', encoding='utf-8')
print('✅ تم ضبط MADAR_PAYMENT_NUMBER و SMS_WEBHOOK_TOKEN')
print('SMS_WEBHOOK_TOKEN=' + updates['SMS_WEBHOOK_TOKEN'])
PY

ثم شغّل البوت + الوركر + مستقبل رسائل المدار:

cd /root/medical_bot && \
pip install -r requirements.txt --break-system-packages && \
pkill -9 -f "bot.py|worker.py|sms_webhook.py|venv/bin/python" 2>/dev/null || true; \
rm -f nohup.out worker.log sms_webhook.log; \
nohup python3 bot.py > nohup.out 2>&1 & \
nohup python3 worker.py > worker.log 2>&1 & \
nohup python3 sms_webhook.py > sms_webhook.log 2>&1 & \
sleep 8; \
echo "=== bot log ==="; tail -40 nohup.out; \
echo "=== worker log ==="; tail -30 worker.log; \
echo "=== sms webhook log ==="; tail -30 sms_webhook.log

اختبار سريع لمستقبل الرسائل من السيرفر:

cd /root/medical_bot
TOKEN=$(grep '^SMS_WEBHOOK_TOKEN=' .env | cut -d= -f2-)
curl "http://127.0.0.1:8088/sms?token=$TOKEN&text=المشترك الكريم، لقد تم تحويل 2 د.ل الي رصيدك من الرقم 218914852966، رصيدك الحالي 3.71 د.ل"

إعداد MacroDroid على الهاتف الذي فيه شريحة 0918874659:
1) ثبّت MacroDroid.
2) Add Macro.
3) Trigger: SMS Received.
4) اجعل المرسل Any أو المدار، والمحتوى يحتوي كلمة: تحويل.
5) Action: HTTP Request أو Webhook.
6) Method: POST أو GET.
7) URL يكون بهذا الشكل:
   http://SERVER_IP:8088/sms?token=SMS_WEBHOOK_TOKEN
8) إذا كان POST: Body/Form field اسمه text وقيمته نص الرسالة SMS Message.
   إذا كان GET: أضف في الرابط &text=[sms_message] حسب المتغير الذي يعرضه MacroDroid.
9) Save.

مهم:
- SERVER_IP هو IPv4 الخاص بالسيرفر في DigitalOcean.
- SMS_WEBHOOK_TOKEN تأخذه من .env أو من نتيجة الأمر الأول.
- لو لم يتفعل الاشتراك تلقائيًا، افتح sms_webhook.log لمعرفة هل وصلت الرسالة أم لا.
