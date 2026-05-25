# medical_bot_GOOGLE_FIRST_GIF_QUEUE_STABLE

هذه نسخة مستقرة مبنية على النسخة التي كانت شغالة عندك، وفيها:

- bot.py: واجهة البوت واستقبال الملفات والأزرار.
- worker.py: يعالج الملخصات وتوليد بنك الأسئلة في الخلفية.
- queue_jobs.py: طابور Redis/Valkey.
- requirements.txt

الجديد في هذه النسخة:
1) Google/Gemini مباشر أولاً باستخدام GOOGLE_API_KEYS.
2) لو مفاتيح Google فشلت، يستخدم OpenRouter كاحتياط فقط.
3) تدوير تلقائي بين مفاتيح Google حتى يتوزع الاستهلاك.
4) GIF ترحيب عند /start إذا وضعت WELCOME_GIF في .env.

مهم:
- لا ترفع .env إلى GitHub.
- لا تشغل /root/bot.py.
- شغل فقط /root/medical_bot/bot.py و /root/medical_bot/worker.py.
- لا ترفع __pycache__.

إعداد .env المطلوب:

TELEGRAM_BOT_TOKEN=توكن_تيليجرام
GOOGLE_API_KEYS=المفتاح_الأول,المفتاح_الثاني
OPENROUTER_API_KEY=مفتاح_أوبن_راوتر_احتياطي
GEMINI_DIRECT_MODEL=gemini-2.5-flash-lite
GEMINI_MODEL=google/gemini-2.5-flash-lite
USE_QUEUE=1

لإضافة GIF الترحيب:
ضع واحد من هذه الخيارات في .env:

WELCOME_GIF=Telegram_file_id_or_direct_gif_url

أو ضع ملف اسمه welcome.gif داخل /root/medical_bot ثم أضف:

WELCOME_GIF=welcome.gif

أمر التشغيل بعد رفع الملفات إلى GitHub ثم سحبها على السيرفر:

cd /root/medical_bot
git fetch origin main
git reset --hard origin/main
sudo systemctl stop medicalbot.service medical-bot.service telegram-bot.service bot.service 2>/dev/null || true
sudo systemctl disable medicalbot.service medical-bot.service telegram-bot.service bot.service 2>/dev/null || true
pkill -9 -f "bot.py|worker.py|venv/bin/python|/root/bot.py" 2>/dev/null || true
sudo systemctl enable valkey-server 2>/dev/null || true
sudo systemctl start valkey-server 2>/dev/null || true
python3 -m pip install -r requirements.txt --break-system-packages
rm -f nohup.out worker.log
nohup python3 bot.py > nohup.out 2>&1 &
nohup python3 worker.py > worker.log 2>&1 &
sleep 8
ps aux | grep -E "bot.py|worker.py|python" | grep -v grep
tail -80 nohup.out
tail -80 worker.log

للتأكد أنه يستخدم Google أولاً:

cd /root/medical_bot
tail -120 nohup.out
tail -120 worker.log

ابحث عن:
AI provider: Google direct

إذا ظهرت:
AI provider: OpenRouter fallback
معناه مفاتيح Google فشلت مؤقتاً أو الموديل غير متاح، واستخدم OpenRouter كاحتياط.


PAST PAPERS DEDUP UPDATE
------------------------
- /upload_final now ignores duplicated past-paper questions automatically.
- It reports:
  ✅ new saved
  ♻️ duplicates skipped
  ⚠️ invalid questions
- Duplicate detection tolerates spaces, punctuation, Arabic letter variants, and option order.
- Admin can use /past_stats to see counts per university/subject.
