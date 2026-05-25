نسخة ECONOMY_PRELOAD_FIXED

الموجود في هذه النسخة:
- Google Gemini مباشر: المفتاح الأول ثم الثاني، وبعدها OpenRouter كاحتياط.
- حدود الخطة المجانية:
  * ملخص: مرة كل 24 ساعة
  * Basic: مرة كل 24 ساعة
  * Cases: مرة كل 48 ساعة
  * Challenge: مرة كل 48 ساعة
  * OSCE: مرة كل أسبوع
- الإدمن مستثنى من الحدود.
- الكاش لا يستهلك API ولا يستهلك حد المستخدم.
- أقصى حجم ملف: 30MB.
- أقصى نص للذكاء: 25000 حرف، مع توزيع من بداية/وسط/نهاية المحاضرة.
- زر في لوحة الأدمن: 📚 تخزين محاضرات الفصل.

طريقة التشغيل:
cd /root/medical_bot && \
git fetch origin main && \
git reset --hard origin/main && \
pip install -r requirements.txt --break-system-packages && \
pkill -9 -f "bot.py|worker.py|venv/bin/python" 2>/dev/null || true; \
rm -f nohup.out worker.log; \
nohup python3 bot.py > nohup.out 2>&1 & \
nohup python3 worker.py > worker.log 2>&1 & \
sleep 8; \
tail -60 nohup.out; \
tail -60 worker.log
