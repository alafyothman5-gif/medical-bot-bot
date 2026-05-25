# medical_bot_QUEUE_STABLE

هذه نسخة مستقرة مبنية على النسخة التي أرسلتها، مع إعادة تفعيل Redis/Worker بطريقة أنظف.

الملفات:
- bot.py: البوت الأساسي.
- worker.py: يعالج الملخصات وتوليد بنك الأسئلة في الخلفية.
- queue_jobs.py: طابور Redis/Valkey.
- requirements.txt

مهم:
- لا ترفع .env لأي مكان.
- لا تشغل /root/bot.py.
- شغل فقط /root/medical_bot/bot.py و /root/medical_bot/worker.py.
- إذا تريد تعطيل الطابور مؤقتاً ضع USE_QUEUE=0 في .env أو شغل bot.py فقط.

أمر التشغيل بعد رفع الملفات إلى /root/medical_bot:

cd /root/medical_bot
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
tail -60 nohup.out
tail -60 worker.log
