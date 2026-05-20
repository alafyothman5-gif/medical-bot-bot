#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Smart Medical Assistant Bot
- Primary AI: Groq (Llama 3 70B) – fast, high limits, free
- Fallback AI: Google Gemini 1.5 Flash – free, reliable backup
- Queue-based request management, LRU cache, Pickle persistence
- Strict prompts, low temperature, JSON validation for accuracy
"""

import asyncio
import base64
import copy
import hashlib
import io
import json
import logging
import os
import pickle
import random
import re
import threading
import time
import uuid
from collections import OrderedDict, defaultdict
from datetime import datetime
from typing import Any, DefaultDict, Dict, List, Optional, Tuple

# AI libraries
import google.generativeai as genai
from groq import Groq

# Telegram
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, BotCommand
from telegram.error import BadRequest as _TelegramBadRequest
from telegram.ext import (
    ApplicationBuilder,
    PicklePersistence,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    filters,
    ContextTypes,
)

# Document processing
from pypdf import PdfReader
from docx import Document

# ----------------------------------------------------------------------
# LOGGING
# ----------------------------------------------------------------------
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

# ----------------------------------------------------------------------
# ENVIRONMENT VARIABLES
# ----------------------------------------------------------------------
TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]

# Groq (primary)
GROQ_API_KEY = os.environ.get("GROQ_API_KEY")
GROQ_AVAILABLE = bool(GROQ_API_KEY)
if GROQ_AVAILABLE:
    groq_client = Groq(api_key=GROQ_API_KEY)
    logger.info("Groq client initialized (primary AI)")

# Google Gemini (fallback)
GOOGLE_API_KEY = os.environ["GOOGLE_API_KEY"]
genai.configure(api_key=GOOGLE_API_KEY)
gemini_model = genai.GenerativeModel('gemini-1.5-flash')
logger.info("Gemini client initialized (fallback AI)")

# ----------------------------------------------------------------------
# PERFORMANCE & CONCURRENCY
# ----------------------------------------------------------------------
_MAX_CONCURRENT = int(os.environ.get("MAX_CONCURRENT_WORKERS", "10"))
_API_SEMAPHORE = asyncio.Semaphore(_MAX_CONCURRENT)
_REQUEST_QUEUE = asyncio.Queue(maxsize=500)
_QUEUE_WORKERS = []

# Result cache (LRU)
_RESULT_CACHE_SIZE = int(os.environ.get("RESULT_CACHE_SIZE", "512"))
_RESULT_CACHE: OrderedDict = OrderedDict()
_RESULT_CACHE_LOCK = threading.Lock()

def _cache_get(key: tuple) -> Any:
    with _RESULT_CACHE_LOCK:
        val = _RESULT_CACHE.get(key)
        if val is None:
            return None
        _RESULT_CACHE.move_to_end(key)
    return copy.deepcopy(val)

def _cache_put(key: tuple, value: Any) -> None:
    snapshot = copy.deepcopy(value)
    with _RESULT_CACHE_LOCK:
        if key in _RESULT_CACHE:
            _RESULT_CACHE.move_to_end(key)
            _RESULT_CACHE[key] = snapshot
            return
        _RESULT_CACHE[key] = snapshot
        while len(_RESULT_CACHE) > _RESULT_CACHE_SIZE:
            _RESULT_CACHE.popitem(last=False)

def _md5_bytes(data: bytes) -> str:
    return hashlib.md5(data).hexdigest()

# ----------------------------------------------------------------------
# ADMIN GATE
# ----------------------------------------------------------------------
ADMIN_USER_IDS: list[int] = [1692255414]   # Replace with your Telegram ID

def _is_admin(user_id: int) -> bool:
    return user_id in ADMIN_USER_IDS

logger.info(f"Admin gate: {len(ADMIN_USER_IDS)} admin(s)")

# ----------------------------------------------------------------------
# SAMPLE MEDICAL TEXT (for self-test)
# ----------------------------------------------------------------------
SAMPLE_MEDICAL_TEXT = (
    "The heart is a four-chambered muscular organ that pumps blood through the circulatory "
    "system. The right atrium receives deoxygenated blood from the body via the superior and "
    "inferior vena cava, passes it to the right ventricle, which pumps it through the pulmonary "
    "artery to the lungs. Oxygenated blood returns through the pulmonary veins to the left atrium, "
    "flows into the left ventricle, and is ejected via the aorta to the systemic circulation.\n\n"
    "The sinoatrial (SA) node, located in the right atrium, is the primary pacemaker of the heart "
    "and normally fires at 60–100 beats per minute in resting adults. The atrioventricular (AV) "
    "node delays conduction to allow ventricular filling. Coronary arteries supply the myocardium "
    "with blood; rupture of an atherosclerotic plaque with subsequent thrombus formation is the "
    "most common cause of acute myocardial infarction.\n\n"
    "Common ECG findings in acute STEMI include ST-segment elevation in two or more contiguous "
    "leads. First-line acute management includes aspirin, oxygen if hypoxic, nitrates for pain, "
    "and reperfusion via primary PCI within 90 minutes when available."
)

# ----------------------------------------------------------------------
# INTERNATIONALISATION (UI translations)
# ----------------------------------------------------------------------
SUPPORTED_LANGS = ("en", "ar")
TRANSLATIONS = {
    "en": {
        "choose_lang": "🌐 <b>Choose your language</b>\n<i>اختر لغتك المفضلة</i>",
        "lang_set": "✅ Language set to English.",
        "btn_lang_en": "🇬🇧  English",
        "btn_lang_ar": "🇸🇦  العربية",
        "welcome": (
            "Welcome to your Smart Medical Assistant! 🩺⚡️\n\n"
            "Upload any medical handout (PDF, image, DOCX, TXT) and I'll generate high‑yield MCQs, summaries, or clinical detective cases!\n\n"
            "How it works 🎮\n"
            "1️⃣ Choose a mode: Quiz / Summary / Clinical Detective.\n"
            "2️⃣ For Quiz: pick style (Basic / Cases / Challenging), then number of questions.\n"
            "3️⃣ Tap A–E to answer — instant feedback.\n"
            "4️⃣ Review mistakes with /review.\n\n"
            "Ready? Send a file or use the menu below 👇"
        ),
        "help": (
            "📋 How to use the bot\n"
            "─────────────────────\n\n"
            "1. Send a medical PDF, image, DOCX, or TXT.\n"
            "2. Choose mode: Quiz / Summary / Clinical Detective.\n"
            "3. For Quiz: pick style (Basic/Cases/Challenging), then number of questions.\n"
            "4. Tap A–E to answer — instant feedback.\n"
            "5. Review wrong answers with /review.\n\n"
            "Commands\n"
            "  /stats     your quiz statistics\n"
            "  /review    replay wrong answers from last quiz\n"
            "  /stop      stop current quiz\n"
            "  /language  switch interface language\n\n"
            "Note: Questions and explanations are always in English."
        ),
        "diff_easy": "🟢  Easy",
        "diff_medium": "🟡  Medium",
        "diff_hard": "🔴  Hard",
        "diff_nightmare": "💀  USMLE Nightmare",
        "diff_easy_label": "🟢 Easy",
        "diff_medium_label": "🟡 Medium",
        "diff_hard_label": "🔴 Hard",
        "diff_nightmare_label": "💀 USMLE Nightmare",
        "pdf_received": "📄 File received — extracting text…",
        "pdf_ready": "✅ *File ready* — choose difficulty:",
        "pdf_no_text": "⚠️ Could not extract text from this file.\nPlease ensure it contains selectable text.",
        "pdf_failed": "❌ Failed to read the file.\n\n{err}",
        "pdf_only": "⚠️ Please send a medical PDF.",
        "no_pdf": "⚠️ No file found. Please send a file first.",
        "stale_button": "⚠️ This button is outdated. Send a new file to start a fresh quiz.",
        "diff_selected": "✅ *{diff}* — how many questions?",
        "generating": "⏳ *Generating {count} high-quality medical questions*\n_{diff}_\n\nPlease wait a few seconds.",
        "ready": "✅ *Quiz ready!*\n\n*Questions:* {count}  ·  *Difficulty:* {diff}\n\nGood luck 🩺",
        "no_questions": "❌ Couldn't generate questions. Please try a different PDF.",
        "json_error": "❌ Couldn't parse the generated questions. Please try again.",
        "mcq_error": "❌ An error occurred while generating questions.\n\n{err}",
        "meta_question": "Question {i} / {n}",
        "correct": "✅ *Correct*",
        "wrong": "❌ *Wrong* — you chose *{chosen}*",
        "answer_label": "*Answer:*",
        "explanation": "📖 *Explanation:*",
        "send_retry": "⚠️ Network hiccup — retrying…",
        "send_failed_alive": "⚠️ The next question failed to send. Your quiz is still active — tap /review or /stats, or send a new PDF.",
        "answer_error_recovery": "⚠️ An internal error occurred. The quiz was resumed.",
        "plain_final": "Quiz Complete!\nScore: {score}/{total} ({pct}%)",
        "plain_review_header": "Review {i}/{n}",
        "plain_answered": "You answered: {chosen}",
        "plain_correct_line": "Correct: {letter}. {opt}",
        "plain_correct": "Correct!",
        "plain_wrong": "Wrong! You chose {chosen}",
        "plain_explanation": "Explanation: {text}",
        "complete_title": "🎓 *Quiz Complete*",
        "lbl_difficulty": "*Difficulty:*",
        "lbl_score": "*Score:*",
        "lbl_grade": "*Grade:*",
        "wrong_count": "❌ {n} wrong answer{plural} — tap below to review.",
        "btn_review": "📋 Review Mistakes",
        "btn_stats": "📊 My Stats",
        "btn_new_quiz": "➕ New Quiz",
        "grade_outstanding": "🏆 Outstanding",
        "grade_excellent": "⭐ Excellent",
        "grade_good": "👍 Good",
        "grade_satisfactory": "📚 Satisfactory",
        "grade_review": "⚠️ Needs review",
        "grade_study": "❌ Significant study needed",
        "btn_review_next": "Next ▶",
        "btn_review_exit": "✖ Done",
        "review_header": "📋 *Review {i} / {n}*",
        "review_your": "*Your answer:*",
        "review_correct_short": "✅ correct",
        "review_wrong_short": "❌ wrong",
        "review_timeout_short": "⏰ no answer",
        "review_complete": "✅ *Review complete*\n\nSend a new PDF to start another quiz.",
        "review_ended": "Review ended. Send a new PDF to start another quiz.",
        "review_none": "✅ No wrong answers to review.",
        "review_no_active": "No active review.",
        "review_no_wrong_inline": "No wrong answers to review.",
        "stats_header": "📊 *Your Statistics*",
        "stats_none": "📊 *Your Statistics*\n\nNo quizzes completed yet.\nSend a medical PDF to get started.",
        "stats_quizzes": "🏁 *Quizzes:*",
        "stats_correct": "✅ *Correct:*",
        "stats_accuracy": "🎯 *Accuracy:*",
        "stats_best": "🏆 *Best score:*",
        "stats_fav": "❤️ *Favourite:*",
        "stopped": "🛑 Quiz stopped. Send a new PDF to start again.",
        "no_active": "No active quiz. Send a PDF to start.",
        "no_session_btn": "⚠️ This quiz has expired. Send a new PDF to start a fresh quiz.",
        "expired_btn": "⚠️ This quiz is expired. Please continue with the latest quiz.",
        "quiz_in_progress": "A quiz is in progress — use the A–E buttons to answer.\nUse /stop to end the quiz.",
        "send_pdf_hint": "Send me a medical PDF to start an interactive quiz.\nUse /help for instructions.",
        "unhandled_error": "⚠️ An unexpected error occurred. Please try again or use /stop to reset.",
        "choose_mode": "📂 Choose a mode:\n\n💡 *Tip:* Select Quiz for MCQs, Summary for notes, or Detective for a clinical case.",
        "btn_mode_quiz": "📝 Quiz",
        "btn_mode_summary": "📚 Summary",
        "btn_mode_detective": "🕵️‍♂️ Clinical Detective",
        "mode_quiz_prompt": "📝 *Quiz mode* — send your file to begin\\.",
        "mode_detective_prompt": (
            "🕵️‍♂️ *Clinical Detective Mode*\n\n"
            "You'll get a mysterious case and up to 3 diagnostic moves "
            "before the final verdict.\n\n"
            "Send a medical file (PDF, image, DOCX, TXT) — OR — type a "
            "topic or specialty (e.g. \"pediatric cardiology\", "
            "\"diabetic ketoacidosis\") to base the case on."
        ),
        "detective_opening": "🕵️‍♂️ *New case opened…*\n_Building vignette — a few seconds._",
        "detective_thinking": "🕵️‍♂️ _Analysing your move…_",
        "detective_round_header": "🕵️‍♂️ *Round {round}/{max}*",
        "detective_final_header": "🏁 *Final verdict*",
        "detective_finished": "✅ Case closed. Use /start to play a new one.",
        "detective_error": "❌ Something went wrong running the case.\n\n{err}",
        "detective_no_input": "⚠️ Send a file or type a topic to start the case.",
        "detective_busy": "⏳ Hold on — still processing your previous move.",
        "mode_summary_prompt": "📚 *Summary mode* — send your file to begin\\.",
        "choose_summary_style": "📋 Choose a summary style:",
        "btn_style_disease": "🩺 Disease Profile",
        "btn_style_highyield": "🧠 High-Yield & Mnemonics",
        "btn_style_osce": "📋 OSCE / Clinical Approach",
        "btn_back": "⬅️ Back",
        "style_selected_disease": "🩺 *Disease Profile* — send your file to begin\\.",
        "style_selected_highyield": "🧠 *High‑Yield & Mnemonics* — send your file to begin\\.",
        "style_selected_osce": "📋 *OSCE / Clinical Approach* — send your file to begin\\.",
        "summary_generating": "⏳ *Generating summary…*\n\nThis may take a moment\\.",
        "summary_header": "📚 *Lecture Summary*\n\n",
        "summary_failed": "❌ Failed to generate summary.\n\n{err}",
        "file_received": "📄 File received — extracting text…",
        "file_no_text": "⚠️ Could not extract any text from this file.",
        "file_unsupported": (
            "⚠️ Unsupported file type. Please send a PDF, "
            "image (JPG/PNG), Word doc (.docx), or text file (.txt)."
        ),
        "support_text": "شكراً لأنك تفكر في دعم المشروع! 🌟\nبإمكانك دعم استمرارية البوت عبر:",
        "support_madar": "📱 ادعمني (رصيد مدار)",
        "support_details": (
            "لدعم استمرار البوت عبر رصيد مدار:\n\n"
            "1. قم بتحويل الرصيد إلى الرقم: 0918874659\n"
            "2. بعد التحويل، أرسل لي صورة الإيصال أو رقم هاتفك في رسالة خاصة هنا:\n"
            "@Othma2003\n\n"
            "شكراً لدعمك يا دكتور! 💙"
        ),
    },
    "ar": {
        "choose_lang": "🌐 <b>اختر لغتك المفضلة</b>\n<i>Choose your language</i>",
        "lang_set": "✅ تم تعيين اللغة إلى العربية.",
        "btn_lang_en": "🇬🇧  English",
        "btn_lang_ar": "🇸🇦  العربية",
        "welcome": (
            "مرحباً بك في مساعدك الطبي الذكي! 🩺⚡️\n\n"
            "ارفع أي ملف طبي (PDF، صورة، DOCX، TXT) وسأقوم بتحويله إلى أسئلة، ملخصات، أو حالات تحقيق طبي فورية!\n\n"
            "كيف يعمل 🎮\n"
            "1️⃣ اختر الوضع: اختبار / ملخص / محقق طبي.\n"
            "2️⃣ في الاختبار: اختر نمط الأسئلة (أساسي / حالات / تحدي)، ثم عدد الأسئلة.\n"
            "3️⃣ اضغط A–E للإجابة — تغذية راجعة فورية.\n"
            "4️⃣ راجع أخطاءك عبر /review.\n\n"
            "جاهز؟ أرسل ملفاً أو استخدم القائمة 👇"
        ),
        "help": (
            "📋 طريقة الاستخدام\n"
            "─────────────────────\n\n"
            "1. أرسل ملف PDF طبي، صورة، DOCX، أو TXT.\n"
            "2. اختر الوضع: اختبار / ملخص / محقق طبي.\n"
            "3. للاختبار: اختر نمط الأسئلة (أساسي / حالات / تحدي)، ثم عدد الأسئلة.\n"
            "4. اضغط A–E للإجابة — تغذية راجعة فورية.\n"
            "5. راجع الإجابات الخاطئة عبر /review.\n\n"
            "الأوامر\n"
            "  /stats     إحصائياتك\n"
            "  /review    مراجعة الأخطاء\n"
            "  /stop      إيقاف الاختبار\n"
            "  /language  تغيير اللغة\n\n"
            "ملاحظة: الأسئلة والشروحات باللغة الإنجليزية فقط."
        ),
        "diff_easy": "🟢  سهل",
        "diff_medium": "🟡  متوسط",
        "diff_hard": "🔴  صعب",
        "diff_nightmare": "💀  كابوس USMLE",
        "diff_easy_label": "🟢 سهل",
        "diff_medium_label": "🟡 متوسط",
        "diff_hard_label": "🔴 صعب",
        "diff_nightmare_label": "💀 كابوس USMLE",
        "pdf_received": "📄 تم استلام الملف — جاري استخراج النص…",
        "pdf_ready": "✅ *الملف جاهز* — اختر مستوى الصعوبة:",
        "pdf_no_text": "⚠️ تعذّر استخراج النص من هذا الملف.\nتأكد من أن الملف يحتوي على نص قابل للتحديد.",
        "pdf_failed": "❌ فشل في قراءة الملف.\n\n{err}",
        "pdf_only": "⚠️ من فضلك أرسل ملف PDF طبي.",
        "no_pdf": "⚠️ لم يُعثر على ملف. أرسل ملفاً أولاً.",
        "stale_button": "⚠️ هذا الزر قديم. أرسل ملفاً جديداً لبدء اختبار جديد.",
        "diff_selected": "✅ *{diff}* — كم عدد الأسئلة؟",
        "generating": "⏳ *جاري إنشاء {count} سؤال طبي عالي الجودة*\n_{diff}_\n\nيرجى الانتظار بضع ثوانٍ.",
        "ready": "✅ *الاختبار جاهز!*\n\n*الأسئلة:* {count}  ·  *الصعوبة:* {diff}\n\nبالتوفيق 🩺",
        "no_questions": "❌ تعذّر إنشاء الأسئلة. من فضلك جرّب ملفاً آخر.",
        "json_error": "❌ تعذّر تحليل الأسئلة. من فضلك حاول مرة أخرى.",
        "mcq_error": "❌ حدث خطأ أثناء إنشاء الأسئلة.\n\n{err}",
        "meta_question": "سؤال {i} / {n}",
        "correct": "✅ *إجابة صحيحة*",
        "wrong": "❌ *إجابة خاطئة* — اخترت *{chosen}*",
        "answer_label": "*الإجابة الصحيحة:*",
        "explanation": "📖 *الشرح:*",
        "send_retry": "⚠️ مشكلة شبكة — جاري إعادة المحاولة…",
        "send_failed_alive": "⚠️ فشل إرسال السؤال التالي. اختبارك لا يزال نشطاً — اضغط /review أو /stats أو أرسل ملف PDF جديد.",
        "answer_error_recovery": "⚠️ حدث خطأ داخلي. تم استئناف الاختبار.",
        "plain_final": "اكتمل الاختبار!\nالنتيجة: {score}/{total} ({pct}%)",
        "plain_review_header": "مراجعة {i}/{n}",
        "plain_answered": "إجابتك: {chosen}",
        "plain_correct_line": "الإجابة الصحيحة: {letter}. {opt}",
        "plain_correct": "إجابة صحيحة!",
        "plain_wrong": "إجابة خاطئة! اخترت {chosen}",
        "plain_explanation": "الشرح: {text}",
        "complete_title": "🎓 *اكتمل الاختبار*",
        "lbl_difficulty": "*الصعوبة:*",
        "lbl_score": "*النتيجة:*",
        "lbl_grade": "*التقدير:*",
        "wrong_count": "❌ {n} إجابة خاطئة — اضغط أدناه للمراجعة.",
        "btn_review": "📋 مراجعة الأخطاء",
        "btn_stats": "📊 إحصائياتي",
        "btn_new_quiz": "➕ اختبار جديد",
        "grade_outstanding": "🏆 ممتاز جداً",
        "grade_excellent": "⭐ ممتاز",
        "grade_good": "👍 جيد",
        "grade_satisfactory": "📚 مقبول",
        "grade_review": "⚠️ يحتاج مراجعة",
        "grade_study": "❌ يحتاج دراسة مكثفة",
        "btn_review_next": "التالي ▶",
        "btn_review_exit": "✖ إنهاء",
        "review_header": "📋 *مراجعة {i} / {n}*",
        "review_your": "*إجابتك:*",
        "review_correct_short": "✅ صحيحة",
        "review_wrong_short": "❌ خاطئة",
        "review_timeout_short": "⏰ بدون إجابة",
        "review_complete": "✅ *اكتملت المراجعة*\n\nأرسل ملف PDF جديد لبدء اختبار آخر.",
        "review_ended": "تم إنهاء المراجعة. أرسل ملف PDF جديد لبدء اختبار آخر.",
        "review_none": "✅ لا توجد إجابات خاطئة للمراجعة.",
        "review_no_active": "لا توجد مراجعة نشطة.",
        "review_no_wrong_inline": "لا توجد إجابات خاطئة للمراجعة.",
        "stats_header": "📊 *إحصائياتك*",
        "stats_none": "📊 *إحصائياتك*\n\nلم تكمل أي اختبار بعد.\nأرسل ملف PDF طبي للبدء.",
        "stats_quizzes": "🏁 *الاختبارات:*",
        "stats_correct": "✅ *الصحيحة:*",
        "stats_accuracy": "🎯 *الدقة:*",
        "stats_best": "🏆 *أفضل نتيجة:*",
        "stats_fav": "❤️ *المفضّلة:*",
        "stopped": "🛑 تم إيقاف الاختبار. أرسل ملف PDF جديد لبدء اختبار آخر.",
        "no_active": "لا يوجد اختبار نشط. أرسل ملف PDF للبدء.",
        "no_session_btn": "⚠️ انتهت صلاحية هذا الاختبار. أرسل ملف PDF جديد لبدء اختبار جديد.",
        "expired_btn": "⚠️ انتهت صلاحية هذا الاختبار. من فضلك تابع مع الاختبار الأحدث.",
        "quiz_in_progress": "يوجد اختبار قيد التشغيل — استخدم أزرار A–E للإجابة.\nاستخدم /stop لإنهاء الاختبار.",
        "send_pdf_hint": "أرسل لي ملف PDF طبي لبدء اختبار تفاعلي.\nاستخدم /help للتعليمات.",
        "unhandled_error": "⚠️ حدث خطأ غير متوقع. من فضلك حاول مرة أخرى أو استخدم /stop لإعادة التعيين.",
        "choose_mode": "📂 اختر وضع التشغيل:\n\n💡 *نصيحة:* اختر اختباراً، ملخصاً، أو المحقق الطبي.",
        "btn_mode_quiz": "📝 اختبار",
        "btn_mode_summary": "📚 ملخص",
        "btn_mode_detective": "🕵️‍♂️ المحقق الطبي",
        "mode_quiz_prompt": "📝 *وضع الاختبار* — أرسل ملفك للبدء\\.",
        "mode_detective_prompt": (
            "🕵️‍♂️ *وضع المحقق الطبي*\n\n"
            "ستحصل على حالة سريرية غامضة وحتى 3 خطوات تشخيصية قبل الحكم النهائي.\n\n"
            "أرسل ملفاً طبياً (PDF أو صورة أو DOCX أو TXT) — أو — اكتب موضوعاً "
            "أو تخصصاً (مثل \"طب أطفال القلب\"، \"الحماض الكيتوني السكري\") "
            "لبناء الحالة عليه."
        ),
        "detective_opening": "🕵️‍♂️ *تم فتح حالة جديدة…*\n_جاري بناء العرض السريري — بضع ثوانٍ._",
        "detective_thinking": "🕵️‍♂️ _جاري تحليل خطوتك…_",
        "detective_round_header": "🕵️‍♂️ *الجولة {round}/{max}*",
        "detective_final_header": "🏁 *الحكم النهائي*",
        "detective_finished": "✅ تم إغلاق الحالة. استخدم /start لبدء حالة جديدة.",
        "detective_error": "❌ حدث خطأ أثناء تشغيل الحالة.\n\n{err}",
        "detective_no_input": "⚠️ أرسل ملفاً أو اكتب موضوعاً لبدء الحالة.",
        "detective_busy": "⏳ من فضلك انتظر — ما زلنا نعالج خطوتك السابقة.",
        "mode_summary_prompt": "📚 *وضع الملخص* — أرسل ملفك للبدء\\.",
        "choose_summary_style": "📋 اختر نمط الملخص:",
        "btn_style_disease": "🩺 ملف المرض",
        "btn_style_highyield": "🧠 النقاط الحيوية والتذكيرات",
        "btn_style_osce": "📋 نهج OSCE السريري",
        "btn_back": "⬅️ رجوع",
        "style_selected_disease": "🩺 *ملف المرض* — أرسل ملفك للبدء\\.",
        "style_selected_highyield": "🧠 *النقاط الحيوية* — أرسل ملفك للبدء\\.",
        "style_selected_osce": "📋 *نهج OSCE* — أرسل ملفك للبدء\\.",
        "summary_generating": "⏳ *جاري إنشاء الملخص…*\n\nقد يستغرق هذا لحظة\\.",
        "summary_header": "📚 *ملخص المحاضرة*\n\n",
        "summary_failed": "❌ فشل إنشاء الملخص.\n\n{err}",
        "file_received": "📄 تم استلام الملف — جاري استخراج النص…",
        "file_no_text": "⚠️ تعذّر استخراج أي نص من هذا الملف.",
        "file_unsupported": (
            "⚠️ نوع الملف غير مدعوم. أرسل ملف PDF أو صورة (JPG/PNG) "
            "أو مستند Word (.docx) أو ملفاً نصياً (.txt)."
        ),
        "support_text": "شكراً لأنك تفكر في دعم المشروع! 🌟\nبإمكانك دعم استمرارية البوت عبر:",
        "support_madar": "📱 ادعمني (رصيد مدار)",
        "support_details": (
            "لدعم استمرار البوت عبر رصيد مدار:\n\n"
            "1. قم بتحويل الرصيد إلى الرقم: 0918874659\n"
            "2. بعد التحويل، أرسل لي صورة الإيصال أو رقم هاتفك في رسالة خاصة هنا:\n"
            "@Othma2003\n\n"
            "شكراً لدعمك يا دكتور! 💙"
        ),
    },
}

def get_lang(user_data: dict) -> str:
    lang = user_data.get("lang", "en")
    return lang if lang in SUPPORTED_LANGS else "en"

def t(key: str, lang: str = "en", **kwargs) -> str:
    bundle = TRANSLATIONS.get(lang, TRANSLATIONS["en"])
    text = bundle.get(key, TRANSLATIONS["en"].get(key, key))
    if kwargs:
        try:
            return text.format(**kwargs)
        except Exception:
            return text
    return text

def diff_label(difficulty: str, lang: str) -> str:
    return t(f"diff_{difficulty}_label", lang)

# ----------------------------------------------------------------------
# DIFFICULTY SYSTEM PROMPTS (strict, anti-hallucination)
# ----------------------------------------------------------------------
DIFFICULTY_CONFIGS = {
    "easy": {
        "system_prompt": (
            "You are a medical educator writing introductory-level MCQs for first- and second-year medical students.\n"
            "You MUST base every question ONLY on the provided source text. Do NOT use external knowledge.\n"
            "If the source text lacks information for a proper question, respond with an empty JSON array.\n\n"
            "VIGNETTE: brief clinical vignette with age, sex, chief complaint, and 1-2 key findings.\n"
            "DISTRACTORS: 5 options, all belonging to same clinical category, no filler.\n"
            "EXPLANATION: 2–3 sentences, state why correct answer is right and name the most tempting wrong answer.\n"
            "All output MUST be English. Output valid JSON array only: [{\"question\":\"...\",\"options\":{\"A\":\"...\",...},\"correct\":\"A\",\"explanation\":\"...\"}]\n"
        )
    },
    "medium": {
        "system_prompt": (
            "You are a medical professor writing USMLE Step 1-style MCQs.\n"
            "Base questions ONLY on the provided text. Do NOT invent facts.\n\n"
            "VIGNETTE: full clinical vignette with demographics, history, exam, and at least one lab/imaging.\n"
            "Requires integration of 2-3 clues. Stem asks for most likely diagnosis, next step, mechanism, or initial investigation.\n"
            "DISTRACTORS: 5 highly plausible options targeting common errors.\n"
            "EXPLANATION: 3–4 sentences, include reasoning chain, mechanism, and why two tempting distractors are wrong.\n"
            "Output English JSON only.\n"
        )
    },
    "hard": {
        "system_prompt": (
            "You are a senior USMLE examiner writing Step 2/3 style MCQs.\n"
            "Use ONLY the provided source text.\n\n"
            "VIGNETTE: detailed, with vitals, positive/negative findings, multiple investigations, and deliberate complexity.\n"
            "Stem targets higher-order reasoning (management of complications, atypical presentation, etc.).\n"
            "DISTRACTORS: all 5 closely related, each defensible in a different scenario.\n"
            "EXPLANATION: 4–5 sentences, include correct reasoning, mechanism, why two dangerous distractors are wrong, and a clinical pearl.\n"
            "Output English JSON only.\n"
        )
    },
    "nightmare": {
        "system_prompt": (
            "You are the most rigorous medical board examiner. Write elite trap questions for USMLE Step 3 / subspecialty.\n"
            "Use ONLY the provided source text. Never hallucinate.\n\n"
            "VIGNETTE: extremely detailed, with red herrings, subtle clues requiring synthesis.\n"
            "DISTRACTORS: each appears correct at first glance, exploiting cognitive biases.\n"
            "EXPLANATION: 5–6 sentences, unmask the red herring, state correct reasoning, explain mechanism, address the most dangerous distractor, and give a high-yield pearl.\n"
            "Output English JSON only.\n"
        )
    },
}

# ----------------------------------------------------------------------
# HELPER: AI CALL WITH FALLBACK (Groq primary, Gemini secondary)
# ----------------------------------------------------------------------
async def _call_ai_with_fallback(system_prompt: str, user_prompt: str, temperature: float = 0.2, max_tokens: int = 4096, retries: int = 2) -> str:
    """Call Groq, fallback to Gemini on failure. Raises exception if both fail."""
    last_exception = None
    # Try Groq if available
    if GROQ_AVAILABLE:
        for attempt in range(retries):
            try:
                response = await asyncio.to_thread(
                    groq_client.chat.completions.create,
                    model="llama3-70b-8192",
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    temperature=temperature,
                    max_tokens=max_tokens
                )
                return response.choices[0].message.content.strip()
            except Exception as e:
                last_exception = e
                logger.warning(f"Groq attempt {attempt+1} failed: {e}")
                if attempt < retries - 1:
                    await asyncio.sleep(0.5 * (2 ** attempt))  # exponential backoff
    # Fallback to Gemini
    try:
        full_prompt = f"{system_prompt}\n\n{user_prompt}"
        response = await asyncio.to_thread(
            gemini_model.generate_content,
            full_prompt,
            generation_config=genai.types.GenerationConfig(
                temperature=temperature,
                max_output_tokens=max_tokens
            )
        )
        return response.text.strip()
    except Exception as e:
        last_exception = e
        logger.error(f"Gemini fallback also failed: {e}")
    raise Exception(f"All AI providers failed: {last_exception}")

# ----------------------------------------------------------------------
# MCQ GENERATION (with chunking, parallel via queue, strict JSON validation)
# ----------------------------------------------------------------------
_MCQ_BATCH_SIZE = 5
_BATCH_PERSPECTIVES = [
    "Emphasize diagnostic reasoning and clinical recognition.",
    "Emphasize pathophysiology and underlying mechanisms.",
    "Emphasize management, treatment, and next-best-step decisions.",
    "Emphasize complications, prognosis, and atypical presentations.",
    "Emphasize pharmacology, drug mechanisms, and side-effects.",
]
STYLE_MODIFIERS = {
    "basic": (
        "OVERRIDE: BASIC RECALL MODE. No vignettes. Ultra-short direct factual questions. "
        "Each question stem ≤15 words, each option 1-2 words. Output exactly 5 options A-E. "
        "Explanations 1-2 sentences. All in English."
    ),
    "cases": (
        "OVERRIDE: CLINICAL CASES. Present patient scenarios (age, gender, symptoms, vitals). "
        "Ask for most likely diagnosis, next step, or initial test. No definitions."
    ),
    "challenging": (
        "OVERRIDE: CHALLENGING / TWO-STEP. Highly complex USMLE-style. Require deduction of hidden diagnosis "
        "then answer about mechanism or rare exception."
    ),
}

async def _generate_mcq_chunk(
    text: str,
    count: int,
    difficulty: str,
    batch_idx: int,
    total_batches: int,
    style: str | None = None,
) -> list[dict]:
    config = DIFFICULTY_CONFIGS.get(difficulty, DIFFICULTY_CONFIGS["hard"])
    perspective = _BATCH_PERSPECTIVES[batch_idx % len(_BATCH_PERSPECTIVES)]
    system_prompt = config["system_prompt"]
    if style and style in STYLE_MODIFIERS:
        system_prompt = system_prompt + "\n\n" + STYLE_MODIFIERS[style]

    if total_batches > 1:
        user_prompt = (
            f"Generate exactly {count} MCQs at the requested difficulty.\n"
            f"Batch {batch_idx+1} of {total_batches}. Focus: {perspective}\n"
            "Do NOT duplicate questions across batches.\n\n"
            f"SOURCE TEXT:\n{text}"
        )
    else:
        user_prompt = f"Generate exactly {count} MCQs.\n\nSOURCE TEXT:\n{text}"

    # Try up to 2 times to get valid JSON
    for attempt in range(2):
        raw = await _call_ai_with_fallback(system_prompt, user_prompt, temperature=0.2, max_tokens=4096)
        # Extract JSON array
        json_match = re.search(r"\[.*\]", raw, re.DOTALL)
        if not json_match:
            logger.warning(f"No JSON array found, attempt {attempt+1}")
            continue
        try:
            parsed = json.loads(json_match.group(0))
            if isinstance(parsed, list) and all(isinstance(q, dict) for q in parsed):
                return parsed
        except json.JSONDecodeError as e:
            logger.warning(f"JSON decode error, attempt {attempt+1}: {e}")
    return []  # fallback empty

def shuffle_question_options(q: dict) -> dict:
    """Randomize A-E order while preserving correct answer."""
    letters = ["A", "B", "C", "D", "E"]
    if "options" not in q or "correct" not in q:
        return q
    correct_letter = str(q["correct"]).strip().upper()
    if correct_letter not in q["options"]:
        return q
    items = [(q["options"][L], L == correct_letter) for L in letters if L in q["options"]]
    if len(items) != 5:
        return q
    random.shuffle(items)
    new_options = {}
    new_correct = correct_letter
    for i, (text, is_correct) in enumerate(items):
        L = letters[i]
        new_options[L] = text
        if is_correct:
            new_correct = L
    q["options"] = new_options
    q["correct"] = new_correct
    return q

async def generate_mcqs(
    text: str,
    count: int = 20,
    difficulty: str = "hard",
    style: str | None = None,
) -> list[dict]:
    MAX_CHARS = 40000
    if len(text) > MAX_CHARS:
        text = text[:MAX_CHARS]
    if count <= _MCQ_BATCH_SIZE:
        questions = await _generate_mcq_chunk(text, count, difficulty, 0, 1, style)
    else:
        sizes = []
        remaining = count
        while remaining > 0:
            sizes.append(min(_MCQ_BATCH_SIZE, remaining))
            remaining -= _MCQ_BATCH_SIZE
        tasks = [_generate_mcq_chunk(text, sz, difficulty, i, len(sizes), style) for i, sz in enumerate(sizes)]
        results = await asyncio.gather(*tasks)
        questions = []
        for res in results:
            questions.extend(res)
        if len(questions) > count:
            questions = questions[:count]
    for q in questions:
        shuffle_question_options(q)
    return questions

# ----------------------------------------------------------------------
# SUMMARY GENERATION
# ----------------------------------------------------------------------
_SUMMARY_STYLE_PROMPTS = {
    "disease": (
        "You are an expert medical professor. Create a high-yield lecture summary using ONLY the provided text.\n"
        "Structure: *🌟 High-Yield Overview*, *🦠 Etiology & Pathophysiology*, *🩺 Clinical Presentation*, *🔬 Diagnosis & Investigations*, *💊 Management & Treatment*, *💡 Key Pearls*.\n"
        "Use Telegram Markdown: *bold* for headings and key terms, - for bullets. One idea per bullet. No external knowledge.\n"
        "All content in English.\n"
    ),
    "highyield": (
        "You are a medical educator. Produce High-Yield board exam notes from the provided text only.\n"
        "Use *bold* for high-yield facts, emojis for topic headings, - for bullets. Include a 🧠 *Mnemonic:* where possible.\n"
        "End with *⚡ Quick Recall* (5-7 ultra-condensed bullets). English only.\n"
    ),
    "osce": (
        "You are a clinical educator. Write an OSCE-style clinical approach summary based ONLY on the provided text.\n"
        "Headings: *📋 History Taking*, *🩺 Physical Examination*, *🔬 Investigations*, *💊 Management*, *⚠️ Red Flags & Complications*, *🗣️ Patient Communication*.\n"
        "Use bullet points. English only.\n"
    ),
}

async def generate_summary(text: str, style: str = "disease") -> str:
    MAX_CHARS = 40000
    if len(text) > MAX_CHARS:
        text = text[:MAX_CHARS]
    system_prompt = _SUMMARY_STYLE_PROMPTS.get(style, _SUMMARY_STYLE_PROMPTS["disease"])
    user_prompt = f"Summarize the following lecture material:\n\n{text}"
    return await _call_ai_with_fallback(system_prompt, user_prompt, temperature=0.2, max_tokens=2048)

# ----------------------------------------------------------------------
# CLINICAL DETECTIVE MODE
# ----------------------------------------------------------------------
DETECTIVE_MAX_ROUNDS = 3

def _detective_system_prompt(round_num: int, max_rounds: int) -> str:
    is_final = round_num >= max_rounds
    base = (
        "You are a master clinical simulator running a 'Clinical Detective' case. "
        "Use the provided background text as the only source of medical facts. "
        "Do not invent conditions or treatments not mentioned.\n"
        "Respond in English. Be concise (≤180 words).\n"
    )
    if is_final:
        base += (
            "FINAL ROUND: Your reply MUST include:\n"
            "1. Realistic results of the student's move.\n"
            "2. The CORRECT diagnosis (label 'Diagnosis:').\n"
            "3. Brief evaluation of the student's reasoning.\n"
            "4. 2-3 high-yield teaching pearls.\n"
            "5. Last line exactly: 'Score: X/10'.\n"
            "Do NOT ask another question."
        )
    else:
        base += (
            "MID-CASE TURN: Your reply MUST:\n"
            "1. Give realistic results for the student's move.\n"
            "2. Drop one subtle clinical clue without revealing the diagnosis.\n"
            "3. End with exactly: 'What is your next move?'"
        )
    return base

async def generate_detective_opening(context_text: str) -> str:
    context_text = (context_text or "").strip()[:8000]
    system_prompt = (
        "You are a master clinical simulator. Build a mysterious, incomplete medical case vignette based ONLY on the provided text.\n"
        "Present: 1) Chief complaint, 2) Patient demographics + brief history, 3) Vital signs (HR, BP, RR, SpO2, Temp).\n"
        "Do NOT reveal the diagnosis. End with exactly: 'What is your next move?'\n"
        "English only, ≤200 words."
    )
    user_prompt = f"CONTEXT:\n{context_text}"
    return await _call_ai_with_fallback(system_prompt, user_prompt, temperature=0.5, max_tokens=600)

async def generate_detective_turn(context_text: str, history: list, round_num: int, max_rounds: int = 3) -> str:
    context_text = (context_text or "").strip()[:6000]
    system_prompt = _detective_system_prompt(round_num, max_rounds)
    conversation = f"{system_prompt}\n\nBackground: {context_text}\n"
    for m in history:
        role = m.get("role", "").upper()
        content = m.get("content", "")
        conversation += f"{role}: {content}\n"
    return await _call_ai_with_fallback(system_prompt, conversation, temperature=0.5, max_tokens=800)

# ----------------------------------------------------------------------
# UTILITIES: PDF / IMAGE / DOCX / TXT extraction (unchanged from original)
# ----------------------------------------------------------------------
_SCANNED_PDF_MIN_CHARS = 50
_SCANNED_PDF_OCR_PAGES = 5
_SCANNED_PDF_RENDER_DPI = 200

def _render_scanned_pdf_pages(pdf_bytes: bytes) -> list[bytes]:
    import fitz
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        page_count = min(len(doc), _SCANNED_PDF_OCR_PAGES)
        if page_count == 0:
            return []
        zoom = _SCANNED_PDF_RENDER_DPI / 72.0
        matrix = fitz.Matrix(zoom, zoom)
        out = []
        for idx in range(page_count):
            try:
                pixmap = doc.load_page(idx).get_pixmap(matrix=matrix, alpha=False)
                out.append(pixmap.tobytes("jpeg"))
            except Exception as e:
                logger.warning(f"Render page {idx+1} failed: {e}")
                out.append(b"")
        return out
    finally:
        doc.close()

def _ocr_scanned_pdf(pdf_bytes: bytes) -> str:
    pages_jpeg = _render_scanned_pdf_pages(pdf_bytes)
    if not pages_jpeg:
        return ""
    chunks = []
    for idx, jpeg in enumerate(pages_jpeg):
        if not jpeg:
            continue
        try:
            page_text = extract_text_from_image(jpeg, "image/jpeg")
            if page_text:
                chunks.append(page_text)
        except Exception as e:
            logger.warning(f"OCR page {idx+1} failed: {e}")
    return "\n\n".join(chunks)

def extract_text_from_image(image_bytes: bytes, mime_type: str = "image/jpeg") -> str:
    base64_image = base64.b64encode(image_bytes).decode("utf-8")
    # Use Groq/Gemini via the fallback caller for OCR
    system_prompt = "You are an OCR engine. Extract all text from this image exactly as written. Do not summarize. Output only the raw text."
    user_prompt = f"Image data: data:image/jpeg;base64,{base64_image}"
    # We need synchronous call here; we'll use a helper that runs async and then sync wrapper
    # For simplicity, call Gemini directly synchronously (this function is already sync)
    full_prompt = f"{system_prompt}\n\n{user_prompt}"
    response = gemini_model.generate_content(
        full_prompt,
        generation_config=genai.types.GenerationConfig(temperature=0.0, max_output_tokens=2000)
    )
    return response.text.strip()

def extract_text_from_pdf(pdf_bytes: bytes) -> str:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    pages = []
    for page in reader.pages:
        s = page.extract_text()
        if s:
            pages.append(s.strip())
    text = "\n\n".join(pages)
    if len(text.strip()) >= _SCANNED_PDF_MIN_CHARS:
        return text
    logger.info("Text layer too short, falling back to OCR")
    try:
        ocr_text = _ocr_scanned_pdf(pdf_bytes)
        return ocr_text if ocr_text.strip() else text
    except Exception as e:
        logger.error(f"OCR fallback failed: {e}")
        return text

async def extract_text_from_pdf_async(pdf_bytes: bytes) -> str:
    return await asyncio.to_thread(extract_text_from_pdf, pdf_bytes)

def extract_text_from_docx(docx_bytes: bytes) -> str:
    doc = Document(io.BytesIO(docx_bytes))
    parts = [para.text for para in doc.paragraphs if para.text.strip()]
    return "\n\n".join(parts)

def extract_text_from_txt(txt_bytes: bytes) -> str:
    for enc in ("utf-8", "utf-16", "latin-1"):
        try:
            return txt_bytes.decode(enc)
        except UnicodeDecodeError:
            continue
    return txt_bytes.decode("utf-8", errors="replace")

# ----------------------------------------------------------------------
# STATS (simple quiz statistics)
# ----------------------------------------------------------------------
def get_stats(user_data: dict) -> dict:
    if "stats" not in user_data:
        user_data["stats"] = {
            "quizzes_completed": 0,
            "total_correct": 0,
            "total_questions": 0,
            "highest_pct": 0,
            "difficulty_counts": {"easy": 0, "medium": 0, "hard": 0, "nightmare": 0},
        }
    return user_data["stats"]

def update_stats(user_data: dict, score: int, total: int, difficulty: str) -> None:
    stats = get_stats(user_data)
    pct = round(score / total * 100) if total > 0 else 0
    stats["quizzes_completed"] += 1
    stats["total_correct"] += score
    stats["total_questions"] += total
    if pct > stats["highest_pct"]:
        stats["highest_pct"] = pct
    dc = stats.setdefault("difficulty_counts", {})
    dc[difficulty] = dc.get(difficulty, 0) + 1

def format_stats_text(user_data: dict, name: str | None = None) -> str:
    lang = get_lang(user_data)
    stats = get_stats(user_data)
    total_q = stats["total_questions"]
    total_c = stats["total_correct"]
    accuracy = round(total_c / total_q * 100) if total_q > 0 else 0
    display_name = (name or "—").strip() or "—"
    best = stats.get("highest_pct", 0)
    fav = max(stats.get("difficulty_counts", {}).items(), key=lambda x: x[1])[0] if stats.get("difficulty_counts") else "—"
    return (
        f"{t('stats_header', lang)}\n\n"
        f"{t('stats_user', lang, name=display_name)}\n"
        f"{t('stats_quizzes', lang)} {stats['quizzes_completed']}\n"
        f"{t('stats_correct', lang)} {total_c} / {total_q}\n"
        f"{t('stats_accuracy', lang)} {accuracy}%\n"
        f"{t('stats_best', lang)} {best}%\n"
        f"{t('stats_fav', lang)} {fav}"
    )

# ----------------------------------------------------------------------
# TELEGRAM HELPERS (keyboards, send functions)
# ----------------------------------------------------------------------
def build_language_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("🇬🇧 English", callback_data="lang:en"),
        InlineKeyboardButton("🇸🇦 العربية", callback_data="lang:ar"),
    ]])

def build_mode_keyboard(lang: str) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton(t("btn_mode_quiz", lang), callback_data="mode:quiz")],
        [InlineKeyboardButton(t("btn_mode_summary", lang), callback_data="mode:summary")],
        [InlineKeyboardButton(t("btn_mode_detective", lang), callback_data="mode:detective")],
    ])

def build_summary_style_keyboard(lang: str) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton(t("btn_style_disease", lang), callback_data="style:disease")],
        [InlineKeyboardButton(t("btn_style_highyield", lang), callback_data="style:highyield")],
        [InlineKeyboardButton(t("btn_style_osce", lang), callback_data="style:osce")],
        [InlineKeyboardButton(t("btn_back", lang), callback_data="setup:back")],
    ])

def build_question_style_keyboard(setup_id: str) -> InlineKeyboardMarkup:
    def cb(style: str) -> str:
        return f"qstyle:{setup_id}:{style}"
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📝 Basic MCQs", callback_data=cb("basic"))],
        [InlineKeyboardButton("🩺 Clinical Cases", callback_data=cb("cases"))],
        [InlineKeyboardButton("🧠 Challenging", callback_data=cb("challenging"))],
    ])

def build_difficulty_keyboard(lang: str, setup_id: str) -> InlineKeyboardMarkup:
    def cb(level: str) -> str:
        return f"difficulty:{setup_id}:{level}"
    return InlineKeyboardMarkup([
        [InlineKeyboardButton(t("diff_easy", lang), callback_data=cb("easy"))],
        [InlineKeyboardButton(t("diff_medium", lang), callback_data=cb("medium"))],
        [InlineKeyboardButton(t("diff_hard", lang), callback_data=cb("hard"))],
        [InlineKeyboardButton(t("diff_nightmare", lang), callback_data=cb("nightmare"))],
    ])

def build_count_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("5", callback_data="count:5"),
        InlineKeyboardButton("10", callback_data="count:10"),
        InlineKeyboardButton("20", callback_data="count:20"),
        InlineKeyboardButton("35", callback_data="count:35"),
        InlineKeyboardButton("50", callback_data="count:50"),
    ]])

def build_question_keyboard(quiz_id: str, question_idx: int) -> InlineKeyboardMarkup:
    def cb(letter: str) -> str:
        return f"answer:{quiz_id}:{question_idx}:{letter}"
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("  A  ", callback_data=cb("A")),
        InlineKeyboardButton("  B  ", callback_data=cb("B")),
        InlineKeyboardButton("  C  ", callback_data=cb("C")),
        InlineKeyboardButton("  D  ", callback_data=cb("D")),
        InlineKeyboardButton("  E  ", callback_data=cb("E")),
    ]])

def build_post_quiz_keyboard(has_wrong: bool, lang: str) -> InlineKeyboardMarkup:
    rows = []
    if has_wrong:
        rows.append([InlineKeyboardButton(t("btn_review", lang), callback_data="rv:start")])
    rows.append([InlineKeyboardButton(t("btn_stats", lang), callback_data="rv:stats")])
    rows.append([InlineKeyboardButton(t("btn_new_quiz", lang), callback_data="rv:new")])
    return InlineKeyboardMarkup(rows)

def build_stats_keyboard(lang: str) -> InlineKeyboardMarkup:
    # Empty keyboard as no nickname/XP features
    return InlineKeyboardMarkup([])

def build_review_nav_keyboard(current: int, total: int, lang: str) -> InlineKeyboardMarkup:
    buttons = []
    if current < total - 1:
        buttons.append(InlineKeyboardButton(t("btn_review_next", lang), callback_data="rv:next"))
    buttons.append(InlineKeyboardButton(t("btn_review_exit", lang), callback_data="rv:exit"))
    return InlineKeyboardMarkup([buttons])

def build_leaderboard_keyboard(lang: str) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([[InlineKeyboardButton(t("btn_back", lang), callback_data="menu:back")]])

def progress_bar(current: int, total: int, width: int = 12) -> str:
    filled = round(current / total * width) if total > 0 else 0
    bar = "█" * filled + "░" * (width - filled)
    pct = round(current / total * 100) if total > 0 else 0
    return f"{bar}  {pct}%"

def escape_md(text: object) -> str:
    text = "" if text is None else str(text)
    for ch in ("*", "_", "`", "["):
        text = text.replace(ch, f"\\{ch}")
    return text

def format_question(q: dict, index: int, total: int, difficulty: str, lang: str) -> str:
    opts = q.get("options") or {}
    bar = progress_bar(index, total)
    def opt(key: str) -> str:
        return escape_md(opts.get(key) or f"[{key}?]")
    return (
        f"`{bar}`\n"
        f"*{t('meta_question', lang, i=index, n=total)}* ·   {diff_label(difficulty, lang)}\n"
        f"\n{escape_md(q.get('question') or '')}\n\n"
        f"*A.* {opt('A')}\n"
        f"*B.* {opt('B')}\n"
        f"*C.* {opt('C')}\n"
        f"*D.* {opt('D')}\n"
        f"*E.* {opt('E')}"
    )

def format_plain_question(q: dict, index: int, total: int) -> str:
    opts = q.get("options") or {}
    def opt(key: str) -> str:
        return str(opts.get(key) or f"[{key}?]")
    return (
        f"Q{index}/{total}\n\n"
        f"{q.get('question') or ''}\n\n"
        f"A. {opt('A')}\nB. {opt('B')}\nC. {opt('C')}\nD. {opt('D')}\nE. {opt('E')}"
    )

def format_review_question(entry: dict, index: int, total: int, lang: str) -> str:
    q = entry["question"]
    chosen = entry["chosen"]
    correct = q["correct"].upper()
    opts = q["options"]
    if chosen == "—":
        your_line = f"{t('review_your', lang)} —   {t('review_timeout_short', lang)}"
    elif chosen == correct:
        your_line = f"{t('review_your', lang)} *{chosen}* {t('review_correct_short', lang)}"
    else:
        your_line = f"{t('review_your', lang)} *{chosen}* {t('review_wrong_short', lang)}"
    return (
        f"{t('review_header', lang, i=index, n=total)}\n\n"
        f"{escape_md(q['question'])}\n\n"
        f"*A.* {escape_md(opts['A'])}\n*B.* {escape_md(opts['B'])}\n*C.* {escape_md(opts['C'])}\n*D.* {escape_md(opts['D'])}\n*E.* {escape_md(opts['E'])}\n\n"
        f"{your_line}\n"
        f"{t('answer_label', lang)} *{correct}.* {escape_md(opts[correct])}\n\n"
        f"{t('explanation', lang)} {escape_md(q['explanation'])}"
    )

async def _safe_send(bot, chat_id: int, text: str, **kwargs) -> bool:
    for attempt in range(1, 3):
        try:
            await bot.send_message(chat_id=chat_id, text=text, **kwargs)
            return True
        except Exception as e:
            logger.warning(f"send_message attempt {attempt} failed: {e}")
            if attempt == 1:
                await asyncio.sleep(1.0)
    return False

async def _send_question(chat_id: int, user_data: dict, bot) -> bool:
    session = user_data.get("session")
    if not session:
        return False
    questions = session["questions"]
    idx = session["current"]
    lang = get_lang(user_data)
    if idx >= len(questions):
        return await _send_final_score(chat_id, user_data, bot)
    q = questions[idx]
    quiz_id = session["quiz_id"]
    keyboard = build_question_keyboard(quiz_id, idx)
    md_text = format_question(q, idx+1, len(questions), session["difficulty"], lang)
    sent = await _safe_send(bot, chat_id, md_text, parse_mode="Markdown", reply_markup=keyboard)
    if not sent:
        plain = format_plain_question(q, idx+1, len(questions))
        sent = await _safe_send(bot, chat_id, plain, reply_markup=keyboard)
    if not sent:
        logger.error(f"Failed to send Q{idx+1}/{len(questions)}")
        await _safe_send(bot, chat_id, t("send_failed_alive", lang))
        return False
    return True

async def _send_final_score(chat_id: int, user_data: dict, bot) -> bool:
    session = user_data.get("session")
    if not session:
        return False
    lang = get_lang(user_data)
    score = session["score"]
    questions = session["questions"]
    total = len(questions)
    pct = round(score / total * 100) if total > 0 else 0
    diff = session["difficulty"]
    wrong_answers = session.get("wrong_answers", [])
    wrong_count = len(wrong_answers)
    if pct >= 90:
        grade_key = "grade_outstanding"
    elif pct >= 80:
        grade_key = "grade_excellent"
    elif pct >= 70:
        grade_key = "grade_good"
    elif pct >= 60:
        grade_key = "grade_satisfactory"
    elif pct >= 40:
        grade_key = "grade_review"
    else:
        grade_key = "grade_study"
    bar = progress_bar(total, total)
    msg = (
        f"{t('complete_title', lang)}\n\n`{bar}`\n\n"
        f"{t('lbl_score', lang)}        `{score} / {total}`   ({pct}%)\n"
        f"{t('lbl_grade', lang)}        {t(grade_key, lang)}\n"
        f"{t('lbl_difficulty', lang)}   {diff_label(diff, lang)}"
    )
    if wrong_count > 0:
        plural = "" if wrong_count == 1 else ("s" if lang == "en" else "")
        msg += f"\n\n{t('wrong_count', lang, n=wrong_count, plural=plural)}"
    keyboard = build_post_quiz_keyboard(has_wrong=bool(wrong_answers), lang=lang)
    sent = await _safe_send(bot, chat_id, msg, parse_mode="Markdown", reply_markup=keyboard)
    if not sent:
        sent = await _safe_send(bot, chat_id, t("plain_final", lang, score=score, total=total, pct=pct), reply_markup=keyboard)
    if not sent:
        logger.error("Final score delivery failed")
        return False
    update_stats(user_data, score, total, diff)
    user_data.pop("session", None)
    user_data["review_pool"] = wrong_answers if wrong_answers else []
    return True

# ----------------------------------------------------------------------
# QUEUE WORKER (background processing)
# ----------------------------------------------------------------------
async def _queue_worker():
    while True:
        try:
            coro, future = await _REQUEST_QUEUE.get()
            async with _API_SEMAPHORE:
                result = await coro
            future.set_result(result)
        except Exception as e:
            future.set_exception(e)
        finally:
            _REQUEST_QUEUE.task_done()

def _start_queue_workers(count: int = 5):
    for _ in range(count):
        asyncio.create_task(_queue_worker())

async def enqueue_request(coro):
    loop = asyncio.get_running_loop()
    future = loop.create_future()
    await _REQUEST_QUEUE.put((coro, future))
    return await future

# ----------------------------------------------------------------------
# TELEGRAM HANDLERS (commands, callbacks, messages)
# ----------------------------------------------------------------------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.pop("awaiting_nickname", None)
    await update.message.reply_text(
        "Welcome to your Smart Medical Assistant! 🩺\nPlease choose your language:\n\n"
        "أهلاً بك في مساعدك الطبي الذكي! 🩺\nيرجى اختيار لغتك:",
        reply_markup=build_language_keyboard()
    )

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    lang = get_lang(context.user_data)
    await update.message.reply_text(t("help", lang), parse_mode=None)

async def stop_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ud = context.user_data
    lang = get_lang(ud)
    had_state = any(k in ud for k in ("session", "setup_id", "pending_text", "review", "detective"))
    for k in ("session", "setup_id", "pending_text", "pending_difficulty", "pending_style",
              "selected_style", "pending_count", "pending_mode", "pending_summary_style",
              "review", "review_pool", "detective"):
        ud.pop(k, None)
    await update.message.reply_text(t("stopped" if had_state else "no_active", lang))

async def stats_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    name = (user.first_name or user.full_name) if user else None
    lang = get_lang(context.user_data)
    await update.message.reply_text(
        format_stats_text(context.user_data, name=name),
        parse_mode="Markdown",
        reply_markup=build_stats_keyboard(lang),
    )

async def review_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    lang = get_lang(context.user_data)
    pool = context.user_data.get("review_pool", [])
    if not pool:
        await update.message.reply_text(t("review_none", lang))
        return
    context.user_data["review"] = {"pool": pool, "current": 0}
    await send_review_question(update.effective_chat.id, context)

async def send_review_question(chat_id: int, context: ContextTypes.DEFAULT_TYPE):
    lang = get_lang(context.user_data)
    review = context.user_data.get("review")
    if not review:
        await context.bot.send_message(chat_id=chat_id, text=t("review_no_active", lang))
        return
    pool = review["pool"]
    current = review["current"]
    if current >= len(pool):
        context.user_data.pop("review", None)
        await context.bot.send_message(chat_id=chat_id, text=t("review_complete", lang), parse_mode="Markdown")
        return
    entry = pool[current]
    text = format_review_question(entry, current+1, len(pool), lang)
    keyboard = build_review_nav_keyboard(current, len(pool), lang)
    await _safe_send(context.bot, chat_id, text, parse_mode="Markdown", reply_markup=keyboard)

async def language_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    lang = get_lang(context.user_data)
    await update.message.reply_text(t("choose_lang", lang), reply_markup=build_language_keyboard())

async def support_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    lang = get_lang(context.user_data)
    keyboard = [[InlineKeyboardButton(t("support_madar", lang), callback_data="support_madar")]]
    await update.message.reply_text(t("support_text", lang), reply_markup=InlineKeyboardMarkup(keyboard))

async def button_click(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    lang = get_lang(context.user_data)
    if query.data == "support_madar":
        await query.edit_message_text(t("support_details", lang))

async def testbot_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if not _is_admin(user.id):
        await update.message.reply_text("Unauthorized.")
        return
    await update.message.reply_text("🧪 Self-test running... (simplified)")
    # Basic test: check if AI responds
    try:
        test_response = await _call_ai_with_fallback("You are a helpful assistant.", "Say 'OK'", temperature=0.0, max_tokens=10)
        if "OK" in test_response:
            await update.message.reply_text("✅ AI test passed.")
        else:
            await update.message.reply_text("⚠️ AI test returned unexpected output.")
    except Exception as e:
        await update.message.reply_text(f"❌ AI test failed: {e}")

async def handle_setlang_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    parts = query.data.split(":")
    if len(parts) != 2 or parts[1] not in SUPPORTED_LANGS:
        return
    lang = parts[1]
    context.user_data["lang"] = lang
    await query.message.delete()
    user = update.effective_user
    name_html = user.mention_html() if user else ""
    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text=t("welcome", lang, name=name_html),
        parse_mode="HTML",
    )
    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text=t("choose_mode", lang),
        parse_mode="Markdown",
        reply_markup=build_mode_keyboard(lang),
    )

async def handle_lang_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    parts = query.data.split(":")
    if len(parts) != 2 or parts[1] not in SUPPORTED_LANGS:
        return
    lang = parts[1]
    context.user_data["lang"] = lang
    await query.edit_message_text(t("lang_set", lang))
    user = update.effective_user
    name_html = user.mention_html() if user else ""
    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text=t("welcome", lang, name=name_html),
        parse_mode="HTML",
    )
    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text=t("choose_mode", lang),
        parse_mode="Markdown",
        reply_markup=build_mode_keyboard(lang),
    )

async def handle_mode_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    parts = query.data.split(":")
    if len(parts) != 2 or parts[1] not in ("quiz", "summary", "detective"):
        return
    mode = parts[1]
    lang = get_lang(context.user_data)
    context.user_data.pop("detective", None)
    context.user_data["pending_mode"] = mode
    await query.edit_message_reply_markup()
    if mode == "summary":
        await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text=t("choose_summary_style", lang),
            reply_markup=build_summary_style_keyboard(lang),
        )
    elif mode == "detective":
        await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text=t("mode_detective_prompt", lang),
            parse_mode="Markdown",
        )
    else:
        await context.bot.send_message(
            chat_id=update.effective_chat.id,
            text=t("mode_quiz_prompt", lang),
            parse_mode="Markdown",
        )

async def handle_setup_back_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    lang = get_lang(context.user_data)
    context.user_data.pop("pending_mode", None)
    await query.edit_message_text(
        text=t("choose_mode", lang),
        parse_mode="Markdown",
        reply_markup=build_mode_keyboard(lang),
    )

async def handle_summary_style_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    parts = query.data.split(":")
    if len(parts) != 2 or parts[1] not in ("disease", "highyield", "osce"):
        return
    style = parts[1]
    lang = get_lang(context.user_data)
    context.user_data["pending_summary_style"] = style
    await query.edit_message_reply_markup()
    await context.bot.send_message(
        chat_id=update.effective_chat.id,
        text=t(f"style_selected_{style}", lang),
        parse_mode="Markdown",
    )

async def handle_qstyle_selection(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    parts = query.data.split(":")
    if len(parts) != 3:
        return
    _, setup_id, style = parts
    session_setup_id = context.user_data.get("setup_id")
    has_pending_text = bool(context.user_data.get("pending_text"))
    if session_setup_id and session_setup_id != setup_id:
        await query.edit_message_text(t("stale_button", get_lang(context.user_data)))
        return
    if not has_pending_text:
        await query.edit_message_reply_markup()
        return
    if style not in ("basic", "cases", "challenging"):
        return
    context.user_data["selected_style"] = style
    context.user_data["pending_style"] = style
    context.user_data["pending_difficulty"] = {"basic": "easy", "cases": "medium", "challenging": "hard"}[style]
    await query.edit_message_text(
        "كم عدد الأسئلة؟ / How many questions?",
        reply_markup=build_count_keyboard(),
    )

async def handle_difficulty_selection(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    parts = query.data.split(":")
    if len(parts) != 3:
        return
    _, setup_id, difficulty = parts
    lang = get_lang(context.user_data)
    session_setup_id = context.user_data.get("setup_id")
    has_pending_text = bool(context.user_data.get("pending_text"))
    if session_setup_id and session_setup_id != setup_id:
        await query.edit_message_text(t("stale_button", lang))
        return
    if not has_pending_text:
        await query.edit_message_reply_markup()
        return
    if difficulty not in ("easy", "medium", "hard", "nightmare"):
        return
    context.user_data["pending_difficulty"] = difficulty
    await query.edit_message_text(
        t("diff_selected", lang, diff=diff_label(difficulty, lang)),
        parse_mode="Markdown",
        reply_markup=build_count_keyboard(),
    )

async def handle_count_selection(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    parts = query.data.split(":")
    if len(parts) != 2:
        return
    try:
        count = int(parts[1])
    except ValueError:
        return
    if count not in (5, 10, 20, 35, 50):
        return
    lang = get_lang(context.user_data)
    text = context.user_data.get("pending_text")
    difficulty = context.user_data.get("pending_difficulty")
    style = context.user_data.get("pending_style")
    if not text or not difficulty:
        await query.edit_message_text(t("stale_button", lang))
        return
    await query.edit_message_text(
        t("generating", lang, count=count, diff=diff_label(difficulty, lang)),
        parse_mode="Markdown",
    )
    try:
        file_md5 = context.user_data.get("file_md5")
        cache_key = (file_md5, "quiz", difficulty, count, style) if file_md5 else None
        questions = _cache_get(cache_key) if cache_key else None
        if questions is None:
            questions = await generate_mcqs(text, count, difficulty, style)
            if cache_key and questions:
                _cache_put(cache_key, questions)
        if not questions:
            await query.edit_message_text(t("no_questions", lang), reply_markup=build_count_keyboard())
            return
        quiz_id = str(uuid.uuid4())[:8]
        context.user_data["session"] = {
            "quiz_id": quiz_id,
            "questions": questions,
            "current": 0,
            "score": 0,
            "difficulty": difficulty,
            "wrong_answers": [],
        }
        await query.edit_message_text(
            t("ready", lang, count=len(questions), diff=diff_label(difficulty, lang)),
            parse_mode="Markdown",
        )
        await _send_question(update.effective_chat.id, context.user_data, context.bot)
        # Clear setup data after first question sent
        for k in ("pending_text", "pending_difficulty", "pending_style", "selected_style", "pending_count", "setup_id"):
            context.user_data.pop(k, None)
    except Exception as e:
        logger.exception("MCQ generation error")
        await query.edit_message_text(t("mcq_error", lang, err=str(e)), reply_markup=build_count_keyboard())

async def handle_review_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    action = query.data.split(":")[1]
    lang = get_lang(context.user_data)
    chat_id = update.effective_chat.id
    if action == "start":
        pool = context.user_data.get("review_pool", [])
        if not pool:
            await query.edit_message_text(t("review_no_wrong_inline", lang))
            return
        context.user_data["review"] = {"pool": pool, "current": 0}
        await query.edit_message_reply_markup()
        await send_review_question(chat_id, context)
    elif action == "next":
        review = context.user_data.get("review")
        if not review:
            await query.edit_message_text(t("review_no_active", lang))
            return
        review["current"] += 1
        await query.edit_message_reply_markup()
        await send_review_question(chat_id, context)
    elif action == "exit":
        context.user_data.pop("review", None)
        await query.edit_message_reply_markup()
        await context.bot.send_message(chat_id=chat_id, text=t("review_ended", lang))
    elif action == "stats":
        text = format_stats_text(context.user_data)
        await query.edit_message_reply_markup()
        await context.bot.send_message(chat_id=chat_id, text=text, parse_mode="Markdown", reply_markup=build_stats_keyboard(lang))
    elif action == "new":
        await query.edit_message_reply_markup()
        await context.bot.send_message(chat_id=chat_id, text=t("send_pdf_hint", lang))

async def handle_answer(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    parts = query.data.split(":")
    if len(parts) != 4:
        return
    _, cb_quiz_id, cb_idx_str, chosen = parts
    try:
        cb_idx = int(cb_idx_str)
    except ValueError:
        return
    chat_id = update.effective_chat.id
    lang = get_lang(context.user_data)
    session = context.user_data.get("session")
    if not session or session.get("quiz_id") != cb_quiz_id or session.get("current") != cb_idx:
        await query.edit_message_reply_markup()
        return
    questions = session["questions"]
    idx = session["current"]
    q = questions[idx]
    correct = q["correct"].upper()
    opts = q["options"]
    is_correct = (chosen == correct)
    if is_correct:
        session["score"] += 1
        result_line = t("correct", lang)
    else:
        result_line = t("wrong", lang, chosen=chosen)
        session.setdefault("wrong_answers", []).append({"question": q, "chosen": chosen})
    session["current"] += 1
    next_idx = session["current"]
    feedback = (
        f"{result_line}\n"
        f"{t('answer_label', lang)} *{correct}.* {escape_md(opts.get(correct, correct))}\n\n"
        f"{t('explanation', lang)} {escape_md(q.get('explanation', ''))}"
    )
    original_text = format_question(q, idx+1, len(questions), session["difficulty"], lang)
    full_text = original_text + "\n\n" + feedback
    try:
        await query.edit_message_text(text=full_text, parse_mode="Markdown", reply_markup=None)
    except Exception:
        await query.edit_message_reply_markup()
        plain_head = t("plain_correct", lang) if is_correct else t("plain_wrong", lang, chosen=chosen)
        plain = f"{plain_head}\n\n{t('plain_correct_line', lang, letter=correct, opt=opts.get(correct, correct))}\n\n{t('plain_explanation', lang, text=q.get('explanation', ''))}"
        await _safe_send(context.bot, chat_id, plain)
    await asyncio.sleep(0.5)
    if next_idx >= len(questions):
        await _send_final_score(chat_id, context.user_data, context.bot)
    else:
        await _send_question(chat_id, context.user_data, context.bot)

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message = update.message
    lang = get_lang(context.user_data)
    pending_mode = context.user_data.get("pending_mode", "quiz")
    # Clear previous state
    for k in ("session", "setup_id", "pending_text", "pending_difficulty", "pending_style",
              "selected_style", "pending_count", "review", "review_pool", "detective"):
        context.user_data.pop(k, None)
    # Determine file type
    if message.photo:
        file_type = "image"
        tg_file = await message.photo[-1].get_file()
        raw_bytes = bytes(await tg_file.download_as_bytearray())
    elif message.document:
        doc = message.document
        mime = (doc.mime_type or "").lower()
        name = (doc.file_name or "").lower()
        if mime == "application/pdf" or name.endswith(".pdf"):
            file_type = "pdf"
        elif mime in ("image/jpeg", "image/jpg") or name.endswith((".jpg", ".jpeg")):
            file_type = "image"
        elif mime == "image/png" or name.endswith(".png"):
            file_type = "image"
        elif "wordprocessingml" in mime or name.endswith(".docx"):
            file_type = "docx"
        elif mime == "text/plain" or name.endswith(".txt"):
            file_type = "txt"
        else:
            await message.reply_text(t("file_unsupported", lang))
            return
        tg_file = await doc.get_file()
        raw_bytes = bytes(await tg_file.download_as_bytearray())
    else:
        await message.reply_text(t("file_unsupported", lang))
        return
    msg = await message.reply_text(t("file_received", lang))
    try:
        # Extract text
        if file_type == "pdf":
            text = await extract_text_from_pdf_async(raw_bytes)
        elif file_type == "image":
            text = await asyncio.to_thread(extract_text_from_image, raw_bytes, "image/jpeg")
        elif file_type == "docx":
            text = await asyncio.to_thread(extract_text_from_docx, raw_bytes)
        else:  # txt
            text = extract_text_from_txt(raw_bytes)
        if not text.strip():
            await msg.edit_text(t("file_no_text", lang))
            return
        context.user_data["file_md5"] = _md5_bytes(raw_bytes)
        if pending_mode == "summary":
            style = context.user_data.get("pending_summary_style", "disease")
            await msg.edit_text(t("summary_generating", lang), parse_mode="Markdown")
            summary = await generate_summary(text, style)
            await msg.delete()
            await message.reply_text(t("summary_header", lang) + summary, parse_mode="Markdown")
        elif pending_mode == "detective":
            await msg.edit_text(t("detective_opening", lang), parse_mode="Markdown")
            opening = await generate_detective_opening(text)
            context.user_data["detective"] = {
                "context": text[:8000],
                "history": [{"role": "assistant", "content": opening}],
                "round": 0,
                "max_rounds": DETECTIVE_MAX_ROUNDS,
                "finished": False,
                "busy": False,
            }
            await msg.delete()
            await message.reply_text(opening)
        else:  # quiz mode
            setup_id = str(uuid.uuid4())[:8]
            context.user_data["setup_id"] = setup_id
            context.user_data["pending_text"] = text
            await msg.edit_text(
                "كيف تفضل أن يكون نمط الأسئلة؟ / Choose the question style:",
                reply_markup=build_question_style_keyboard(setup_id),
            )
    except Exception as e:
        logger.exception("File handling error")
        await msg.edit_text(t("pdf_failed", lang, err=e))

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    lang = get_lang(context.user_data)
    detective = context.user_data.get("detective")
    pending_mode = context.user_data.get("pending_mode")
    if detective and isinstance(detective, dict) and not detective.get("finished"):
        # Process detective move
        move = update.message.text.strip()
        if not move:
            return
        if detective.get("busy"):
            await update.message.reply_text(t("detective_busy", lang))
            return
        detective["busy"] = True
        try:
            next_round = detective.get("round", 0) + 1
            history = detective.get("history", [])
            history.append({"role": "user", "content": move})
            reply = await generate_detective_turn(
                detective.get("context", ""), history, next_round, detective.get("max_rounds", 3)
            )
            history.append({"role": "assistant", "content": reply})
            detective["history"] = history
            detective["round"] = next_round
            detective["busy"] = False
            if next_round >= detective.get("max_rounds", 3):
                detective["finished"] = True
                context.user_data.pop("detective", None)
                context.user_data.pop("pending_mode", None)
                await update.message.reply_text(reply)
                await update.message.reply_text(t("detective_finished", lang))
            else:
                await update.message.reply_text(reply)
        except Exception as e:
            logger.exception("Detective turn error")
            detective["busy"] = False
            await update.message.reply_text(t("detective_error", lang, err=e))
        return
    if pending_mode == "detective" and not detective:
        topic = update.message.text.strip()
        if not topic:
            await update.message.reply_text(t("detective_no_input", lang))
            return
        msg = await update.message.reply_text(t("detective_opening", lang), parse_mode="Markdown")
        opening = await generate_detective_opening(topic)
        context.user_data["detective"] = {
            "context": topic[:8000],
            "history": [{"role": "assistant", "content": opening}],
            "round": 0,
            "max_rounds": DETECTIVE_MAX_ROUNDS,
            "finished": False,
            "busy": False,
        }
        await msg.delete()
        await update.message.reply_text(opening)
        return
    if context.user_data.get("session"):
        await update.message.reply_text(t("quiz_in_progress", lang))
    else:
        await update.message.reply_text(t("send_pdf_hint", lang))

# ----------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------
BOT_COMMANDS = [
    BotCommand("start", "🏠 Start bot & language"),
    BotCommand("help", "ℹ️ Help & instructions"),
    BotCommand("stop", "🛑 End current quiz"),
    BotCommand("stats", "📊 View performance stats"),
    BotCommand("review", "❌ Review past mistakes"),
    BotCommand("language", "🌐 Change interface language"),
    BotCommand("support", "📱 Support development"),
    BotCommand("testbot", "🔧 Admin: run self-test"),
]

BOT_COMMANDS_AR = [
    BotCommand("start", "🏠 البداية واختيار اللغة"),
    BotCommand("help", "ℹ️ المساعدة والتعليمات"),
    BotCommand("stop", "🛑 إيقاف الاختبار الحالي"),
    BotCommand("stats", "📊 عرض إحصائيات الأداء"),
    BotCommand("review", "❌ مراجعة الأخطاء السابقة"),
    BotCommand("language", "🌐 تغيير لغة الواجهة"),
    BotCommand("support", "📱 دعم استمرار البوت"),
    BotCommand("testbot", "🔧 المشرف: اختبار ذاتي"),
]

async def post_init(application):
    await application.bot.set_my_commands(BOT_COMMANDS)
    await application.bot.set_my_commands(BOT_COMMANDS_AR, language_code="ar")
    logger.info("Bot ready with Groq+Gemini, Queue, Pickle persistence")

async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    logger.error("Unhandled exception", exc_info=context.error)
    if isinstance(update, Update) and update.effective_message:
        lang = "en"
        try:
            if context.user_data:
                lang = get_lang(context.user_data)
        except Exception:
            pass
        await update.effective_message.reply_text(t("unhandled_error", lang))

def main():
    _start_queue_workers(count=5)
    persistence = PicklePersistence(filepath="bot_data.pkl")
    app = ApplicationBuilder().token(TOKEN).persistence(persistence).concurrent_updates(True).post_init(post_init).build()
    app.add_error_handler(error_handler)
    # Command handlers
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("stop", stop_command))
    app.add_handler(CommandHandler("stats", stats_command))
    app.add_handler(CommandHandler("review", review_command))
    app.add_handler(CommandHandler("language", language_command))
    app.add_handler(CommandHandler("support", support_command))
    app.add_handler(CommandHandler("testbot", testbot_command))
    # Callback handlers
    app.add_handler(CallbackQueryHandler(handle_setlang_callback, pattern=r"^setlang:"))
    app.add_handler(CallbackQueryHandler(handle_lang_callback, pattern=r"^lang:"))
    app.add_handler(CallbackQueryHandler(handle_mode_callback, pattern=r"^mode:"))
    app.add_handler(CallbackQueryHandler(handle_setup_back_callback, pattern=r"^setup:back$"))
    app.add_handler(CallbackQueryHandler(handle_summary_style_callback, pattern=r"^style:"))
    app.add_handler(CallbackQueryHandler(handle_qstyle_selection, pattern=r"^qstyle:"))
    app.add_handler(CallbackQueryHandler(handle_difficulty_selection, pattern=r"^difficulty:"))
    app.add_handler(CallbackQueryHandler(handle_count_selection, pattern=r"^count:"))
    app.add_handler(CallbackQueryHandler(button_click, pattern=r"^support_madar$"))
    app.add_handler(CallbackQueryHandler(handle_answer, pattern=r"^answer:"))
    app.add_handler(CallbackQueryHandler(handle_review_callback, pattern=r"^rv:"))
    # Message handlers
    class SupportedFileFilter(filters.MessageFilter):
        def filter(self, message):
            doc = message.document
            if not doc:
                return False
            mime = (doc.mime_type or "").lower()
            name = (doc.file_name or "").lower()
            return (mime in ("application/pdf", "image/jpeg", "image/jpg", "image/png",
                             "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "text/plain") or
                    name.endswith((".pdf", ".jpg", ".jpeg", ".png", ".docx", ".txt")))
    app.add_handler(MessageHandler(filters.PHOTO, handle_file))
    app.add_handler(MessageHandler(SupportedFileFilter(), handle_file))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    logger.info("Starting bot with Groq+Gemini, queue system, and enhanced accuracy.")
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
