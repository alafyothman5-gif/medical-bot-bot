نسخة Premium + Quick Past Papers

الميزات:
- Premium شهري 15 د.ل
- Premium فصل كامل 40 د.ل / 120 يوم
- Premium Plus فصل كامل 60 د.ل / 120 يوم
- /my_plan لعرض خطة الطالب ورقم User ID
- /set_plan USER_ID premium 30 لتفعيل شهر
- /set_plan USER_ID premium 120 لتفعيل فصل كامل
- /set_plan USER_ID premium_plus 120 لتفعيل Premium Plus
- /remove_premium USER_ID لإرجاع الطالب Free
- إزالة معرف الدعم الشخصي من رسائل الدعم
- زر جديد في لوحة الأدمن: ⚡ جمع سنوات سابقة سريع
- /finish_past لتصدير ZIP مراجعة أسئلة السنوات السابقة بسرعة

التشغيل:
cd /root/medical_bot && \
git fetch origin main && \
git reset --hard origin/main && \
pip install -r requirements.txt --break-system-packages && \
pkill -9 -f "bot.py|worker.py|venv/bin/python" 2>/dev/null || true; \
rm -f nohup.out worker.log; \
nohup python3 bot.py > nohup.out 2>&1 & \
nohup python3 worker.py > worker.log 2>&1 & \
sleep 8; \
tail -60 nohup.out; tail -40 worker.log
