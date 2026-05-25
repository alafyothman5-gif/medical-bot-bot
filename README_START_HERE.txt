نسخة إصلاح تجميع أسئلة السنوات السابقة

الجديد:
- عند التجميع، البوت لا يحفظ النص المستخرج فقط.
- يحفظ أيضاً الملفات الأصلية التي أرسلها الأدمن داخل data/admin_staging_files.
- عند الضغط على "أعطيني المخزون ZIP" أو استخدام /stage_dump، يخرج ZIP يحتوي:
  README.txt
  metadata.json
  uni/subject/raw_questions.txt
  uni/subject/original_files/...

طريقة الاستخدام:
1) /stage_start ajdabiya obgyn أو من لوحة الأدمن.
2) أرسل صور / PDF / Word / PPT للأسئلة.
3) اضغط "أعطيني المخزون ZIP" أو اكتب /stage_dump.
4) أرسل ZIP الناتج إلى ChatGPT ليقرأ original_files إذا كان OCR مشوه.
5) ارفع JSON النهائي للبوت عبر /upload_final uni subject.

مهم:
- لا ترفع .env إلى GitHub.
- لا تشغل أكثر من نسخة bot.py.
- worker.py والqueue_jobs.py مطلوبين للنسخة الحالية.
