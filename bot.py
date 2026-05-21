#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Smart Medical Assistant Bot - OpenRouter Paid Economical Edition
- PPTX, chunking, auto-cleanup, back navigation, admin logs
- Daily usage limits (Basic 3, Cases 2, Challenge 1, Summary 3)
- Improved JSON extraction, plain summary, exact question count prompt
- Summary max_output_tokens raised to 8192 (fix Disease/OSCE blank)
- Disease key disease_v2, summary empty check & error feedback
- File-based persistence for /workspace
- Concurrent generation locks to prevent duplicate API calls
- Fixed: Medical safety filters, 4-option MCQs, long summary chunking, and hard tier prompts
"""

import asyncio
import hashlib
import io
import json
import logging
import os
import random
import re
import sqlite3
import time
import uuid
import urllib.request
import urllib.error
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple

from docx import Document
from pptx import Presentation
from pypdf import PdfReader
from telegram import BotCommand, InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import (
    Application,
    ApplicationBuilder,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    PicklePersistence,
    filters,
)

# =============================================================================
# LOGGING
# =============================================================================
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger("medical_bot_openrouter")

# =============================================================================
# ENVIRONMENT
# =============================================================================
TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
if not TOKEN:
    raise ValueError("TELEGRAM_BOT_TOKEN environment variable not set")
OPENROUTER_API_KEY = os.environ.get("OPENROUTER_API_KEY")
if not OPENROUTER_API_KEY:
    raise ValueError("OPENROUTER_API_KEY environment variable not set")
GEMINI_MODEL_NAME = os.environ.get("GEMINI_MODEL", "google/gemini-2.5-flash-lite")
DATABASE_PATH = os.environ.get(
    "DATABASE_PATH",
    "/workspace/bot_cache.sqlite3"
)
MAX_FILE_SIZE_BYTES = 30 * 1024 * 1024
MAX_TEXT_CHARS_FOR_AI = 20000
MAX_CONCURRENT_GEMINI = int(os.environ.get("MAX_CONCURRENT_GEMINI", "2"))
COOLDOWN_SECONDS = int(os.environ.get("COOLDOWN_SECONDS", "20"))

ADMIN_USER_IDS = {1692255414}
ADMIN_LOG_CHAT_ID = int(os.environ.get("ADMIN_LOG_CHAT_ID", "-1003987051773"))

_GEMINI_SEMAPHORE = asyncio.Semaphore(MAX_CONCURRENT_GEMINI)

# Generation locks to prevent duplicate expensive API calls
_generation_locks: Dict[str, asyncio.Lock] = {}

def get_generation_lock(key: str) -> asyncio.Lock:
    if key not in _generation_locks:
        _generation_locks[key] = asyncio.Lock()
    return _generation_locks[key]


# =============================================================================
# TRANSLATIONS
# =============================================================================
SUPPORTED_LANGS = ("en", "ar")

TRANSLATIONS = {
    "en": {
        "choose_lang": "🌐 Choose your language\nاختر لغتك",
        "welcome": "Welcome to your Smart Medical Assistant 🩺\n\nSend a medical PDF, DOCX, PPTX, or TXT file.",
        "help": "Send a medical file up to 30MB. Choose Quiz or Summary. Commands: /start /stop /stats /review /language /support",
        "choose_mode": "📂 Choose mode:",
        "mode_quiz": "📝 Quiz",
        "mode_summary": "📚 Summary",
        "mode_quiz_prompt": "📝 Quiz mode selected. Send your file.",
        "mode_summary_prompt": "📚 Summary mode selected. Choose style, then send your file.",
        "summary_style_prompt": "📋 Choose summary style:",
        "style_disease": "🩺 Disease Profile",
        "style_highyield": "🧠 High-Yield",
        "style_osce": "📋 OSCE",
        "file_received": "📄 File received — extracting text...",
        "file_too_large": "⚠️ File is larger than 30MB. Please send a smaller file.",
        "file_unsupported": "⚠️ Unsupported file type. Send PDF, DOCX, PPTX, or TXT.",
        "file_no_text": "⚠️ I could not extract readable text from this file.",
        "file_empty": "⚠️ File is empty or contains only images (no readable text).",
        "file_ready_quiz": "✅ File ready. Choose question level:",
        "file_ready_summary": "✅ File ready. Preparing summary...",
        "cached_file": "♻️ Same file found in cache. No extra AI cost.",
        "level_basic": "🟢 Basic",
        "level_cases": "🩺 Cases",
        "level_challenge": "🧠 Challenge",
        "basic_count_prompt": "How many Basic questions do you want?",
        "generating_pool": "⏳ Preparing question bank. This happens once per file...",
        "generating_summary": "⏳ Preparing summary. This happens once per file/style...",
        "not_enough": "⚠️ This lecture only has {n} available questions for this level. Starting with available questions.",
        "quiz_ready": "✅ Quiz ready: {n} questions.",
        "no_questions": "❌ No questions could be generated from this file.",
        "summary_header": "📚 Summary\n\n",
        "stale": "⚠️ This button is outdated. Send /start or upload the file again.",
        "cooldown": "⚠️ Please wait {wait} seconds before sending another request.",
        "busy": "⚠️ You already have a request being processed. Please wait.",
        "q_meta": "Question {i}/{n}",
        "correct": "✅ Correct",
        "wrong": "❌ Wrong — you chose {chosen}",
        "answer": "Answer:",
        "explanation": "Explanation:",
        "complete": "🎓 Quiz Complete\n\nScore: {score}/{total} ({pct}%)",
        "review": "📋 Review Mistakes",
        "stats": "📊 My Stats",
        "new_quiz": "➕ New Quiz",
        "review_none": "✅ No wrong answers to review.",
        "review_done": "✅ Review complete.",
        "stopped": "🛑 Stopped.",
        "no_active": "No active quiz.",
        "quiz_in_progress": "A quiz is in progress. Use /stop to end it.",
        "send_file_hint": "Send a medical file to start.",
        "stats_text": "📊 Your Statistics\n\nQuizzes: {quizzes}\nCorrect: {correct}/{total}\nAccuracy: {accuracy}%\nBest score: {best}%",
        "error": "⚠️ Error: {err}",
        "summary_failed": "⚠️ Failed to generate summary. Please try again or use another file.",
    },
    "ar": {
        "choose_lang": "🌐 اختر لغتك",
        "welcome": "مرحباً بك في المساعد الطبي الذكي 🩺\n\nأرسل ملف PDF أو DOCX أو PPTX أو TXT.",
        "help": "أرسل ملفاً طبياً حتى 30MB. اختر اختبار أو ملخص. الأوامر: /start /stop /stats /review /language /support",
        "choose_mode": "📂 اختر الوضع:",
        "mode_quiz": "📝 اختبار",
        "mode_summary": "📚 ملخص",
        "mode_quiz_prompt": "📝 تم اختيار الاختبار. أرسل الملف.",
        "mode_summary_prompt": "📚 تم اختيار الملخص. اختر نوع الملخص ثم أرسل الملف.",
        "summary_style_prompt": "📋 اختر نوع الملخص:",
        "style_disease": "🩺 ملف المرض",
        "style_highyield": "🧠 نقاط مهمة",
        "style_osce": "📋 OSCE",
        "file_received": "📄 تم استلام الملف — جاري استخراج النص...",
        "file_too_large": "⚠️ الملف أكبر من 30MB. أرسل ملفاً أصغر.",
        "file_unsupported": "⚠️ نوع الملف غير مدعوم. أرسل PDF, DOCX, PPTX أو TXT فقط.",
        "file_no_text": "⚠️ لم أستطع استخراج نص مقروء من الملف.",
        "file_empty": "⚠️ الملف فارغ أو يحتوي على صور فقط (لا يوجد نص مقروء).",
        "file_ready_quiz": "✅ الملف جاهز. اختر مستوى الأسئلة:",
        "file_ready_summary": "✅ الملف جاهز. جاري تجهيز الملخص...",
        "cached_file": "♻️ الملف موجود مسبقاً في الكاش. لن نستهلك الذكاء الاصطناعي من جديد.",
        "level_basic": "🟢 Basic مباشر",
        "level_cases": "🩺 Cases حالات",
        "level_challenge": "🧠 Challenge تحدي",
        "basic_count_prompt": "كم سؤال Basic تريد؟",
        "generating_pool": "⏳ جاري تجهيز بنك الأسئلة. يحدث هذا مرة واحدة فقط لكل ملف...",
        "generating_summary": "⏳ جاري تجهيز الملخص. يحدث هذا مرة واحدة فقط لكل ملف/نوع...",
        "not_enough": "⚠️ هذه المحاضرة فيها {n} سؤال فقط لهذا المستوى. سأبدأ بالموجود.",
        "quiz_ready": "✅ الاختبار جاهز: {n} سؤال.",
        "no_questions": "❌ لم أستطع توليد أسئلة من هذا الملف.",
        "summary_header": "📚 الملخص\n\n",
        "stale": "⚠️ هذا الزر قديم. أرسل /start أو ارفع الملف من جديد.",
        "cooldown": "⚠️ انتظر {wait} ثانية قبل إرسال طلب جديد.",
        "busy": "⚠️ لديك طلب قيد المعالجة. الرجاء الانتظار.",
        "q_meta": "السؤال {i}/{n}",
        "correct": "✅ صحيح",
        "wrong": "❌ خطأ — اخترت {chosen}",
        "answer": "الإجابة:",
        "explanation": "الشرح:",
        "complete": "🎓 انتهى الاختبار\n\nالنتيجة: {score}/{total} ({pct}%)",
        "review": "📋 مراجعة الأخطاء",
        "stats": "📊 إحصائياتي",
        "new_quiz": "➕ اختبار جديد",
        "review_none": "✅ لا توجد أخطاء للمراجعة.",
        "review_done": "✅ انتهت المراجعة.",
        "stopped": "🛑 تم الإيقاف.",
        "no_active": "لا يوجد اختبار نشط.",
        "quiz_in_progress": "يوجد اختبار نشط. استخدم /stop لإنهائه.",
        "send_file_hint": "أرسل ملفاً طبياً للبدء.",
        "stats_text": "📊 إحصائياتك\n\nالاختبارات: {quizzes}\nالصحيح: {correct}/{total}\nالدقة: {accuracy}%\nأفضل نتيجة: {best}%",
        "error": "⚠️ خطأ: {err}",
        "summary_failed": "⚠️ فشل توليد الملخص. حاول مرة أخرى أو استخدم ملفاً آخر.",
    },
}


def get_lang(user_data: dict) -> str:
    lang = user_data.get("lang", "ar")
    return lang if lang in SUPPORTED_LANGS else "ar"


def t(user_data: dict, key: str, **kwargs) -> str:
    lang = get_lang(user_data)
    text = TRANSLATIONS.get(lang, TRANSLATIONS["ar"]).get(key, key)
    try:
        return text.format(**kwargs)
    except Exception:
        return text


# =============================================================================
# SQLITE CACHE
# =============================================================================
_db_lock = asyncio.Lock()

def _db_connect() -> sqlite3.Connection:
    conn = sqlite3.connect(DATABASE_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db() -> None:
    conn = _db_connect()
    cur = conn.cursor()
    cur.execute("PRAGMA journal_mode=WAL")
    cur.execute("PRAGMA synchronous=NORMAL")
    cur.execute("""
        CREATE TABLE IF NOT EXISTS files (
            file_hash TEXT PRIMARY KEY,
            file_name TEXT,
            file_size INTEGER,
            mime_type TEXT,
            extracted_text TEXT,
            created_at TEXT
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS question_banks (
            file_hash TEXT NOT NULL,
            level TEXT NOT NULL,
            questions_json TEXT NOT NULL,
            created_at TEXT,
            PRIMARY KEY (file_hash, level)
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS summaries (
            file_hash TEXT NOT NULL,
            style TEXT NOT NULL,
            summary_text TEXT NOT NULL,
            created_at TEXT,
            PRIMARY KEY (file_hash, style)
        )
    """)
    conn.commit()
    conn.close()

async def db_get_file(file_hash: str) -> Optional[dict]:
    async with _db_lock:
        def work():
            conn = _db_connect()
            row = conn.execute("SELECT * FROM files WHERE file_hash=?", (file_hash,)).fetchone()
            conn.close()
            return dict(row) if row else None
        return await asyncio.to_thread(work)

async def db_save_file(file_hash: str, file_name: str, file_size: int, mime_type: str, extracted_text: str) -> None:
    async with _db_lock:
        def work():
            conn = _db_connect()
            conn.execute(
                "INSERT OR REPLACE INTO files(file_hash, file_name, file_size, mime_type, extracted_text, created_at) VALUES(?,?,?,?,?,?)",
                (file_hash, file_name, file_size, mime_type, extracted_text, datetime.utcnow().isoformat())
            )
            conn.commit()
            conn.close()
        await asyncio.to_thread(work)

async def db_get_questions(file_hash: str, level: str) -> Optional[List[dict]]:
    async with _db_lock:
        def work():
            conn = _db_connect()
            row = conn.execute("SELECT questions_json FROM question_banks WHERE file_hash=? AND level=?", (file_hash, level)).fetchone()
            conn.close()
            if not row:
                return None
            try:
                return json.loads(row["questions_json"])
            except:
                return None
        return await asyncio.to_thread(work)

async def db_save_questions(file_hash: str, level: str, questions: List[dict]) -> None:
    async with _db_lock:
        def work():
            conn = _db_connect()
            conn.execute(
                "INSERT OR REPLACE INTO question_banks(file_hash, level, questions_json, created_at) VALUES(?,?,?,?)",
                (file_hash, level, json.dumps(questions, ensure_ascii=False), datetime.utcnow().isoformat())
            )
            conn.commit()
            conn.close()
        await asyncio.to_thread(work)

async def db_get_summary(file_hash: str, style: str) -> Optional[str]:
    async with _db_lock:
        def work():
            conn = _db_connect()
            row = conn.execute("SELECT summary_text FROM summaries WHERE file_hash=? AND style=?", (file_hash, style)).fetchone()
            conn.close()
            return row["summary_text"] if row else None
        return await asyncio.to_thread(work)

async def db_save_summary(file_hash: str, style: str, summary: str) -> None:
    async with _db_lock:
        def work():
            conn = _db_connect()
            conn.execute(
                "INSERT OR REPLACE INTO summaries(file_hash, style, summary_text, created_at) VALUES(?,?,?,?)",
                (file_hash, style, summary, datetime.utcnow().isoformat())
            )
            conn.commit()
            conn.close()
        await asyncio.to_thread(work)


# =============================================================================
# AUTO CLEANUP
# =============================================================================
async def cleanup_old_cache(days: int = 30) -> None:
    cutoff = (datetime.utcnow() - timedelta(days=days)).isoformat()

    async with _db_lock:
        def work():
            conn = _db_connect()
            cur = conn.cursor()

            old_files = cur.execute(
                "SELECT file_hash FROM files WHERE created_at < ?",
                (cutoff,)
            ).fetchall()

            old_hashes = [row["file_hash"] for row in old_files]

            for file_hash in old_hashes:
                cur.execute("DELETE FROM question_banks WHERE file_hash=?", (file_hash,))
                cur.execute("DELETE FROM summaries WHERE file_hash=?", (file_hash,))
                cur.execute("DELETE FROM files WHERE file_hash=?", (file_hash,))

            cur.execute("DELETE FROM question_banks WHERE created_at < ?", (cutoff,))
            cur.execute("DELETE FROM summaries WHERE created_at < ?", (cutoff,))

            conn.commit()
            conn.close()

            return len(old_hashes)

        deleted = await asyncio.to_thread(work)

    logger.info("Cleanup done. Deleted old files: %s", deleted)


# =============================================================================
# HELPERS
# =============================================================================
user_last_request: Dict[int, float] = {}

def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()

def escape_md(text: object) -> str:
    value = "" if text is None else str(text)
    for ch in ("*", "_", "`", "["):
        value = value.replace(ch, f"\\{ch}")
    return value

def progress_bar(current: int, total: int, width: int = 12) -> str:
    if total <= 0:
        return "░" * width + " 0%"
    filled = round(current / total * width)
    pct = round(current / total * 100)
    return "█" * filled + "░" * (width - filled) + f" {pct}%"

async def cooldown_guard(update: Update, user_data: dict, user_id: int) -> bool:
    now = time.time()
    last = user_last_request.get(user_id, 0)
    if now - last < COOLDOWN_SECONDS:
        wait = int(COOLDOWN_SECONDS - (now - last))
        await update.effective_message.reply_text(t(user_data, "cooldown", wait=wait))
        return True
    user_last_request[user_id] = now
    return False

def get_stats(user_data: dict) -> dict:
    if "stats" not in user_data:
        user_data["stats"] = {"quizzes": 0, "correct": 0, "total": 0, "best": 0}
    return user_data["stats"]

def update_stats(user_data: dict, score: int, total: int) -> None:
    stats = get_stats(user_data)
    pct = round(score / total * 100) if total else 0
    stats["quizzes"] += 1
    stats["correct"] += score
    stats["total"] += total
    if pct > stats["best"]:
        stats["best"] = pct

# =============================================================================
# DAILY USAGE LIMITS
# =============================================================================
def today_key() -> str:
    return datetime.utcnow().strftime("%Y-%m-%d")

def usage_get(user_data: dict) -> dict:
    day = today_key()
    usage = user_data.get("daily_usage")

    if not usage or usage.get("day") != day:
        usage = {
            "day": day,
            "basic": 0,
            "cases": 0,
            "challenge": 0,
            "summary": 0,
        }
        user_data["daily_usage"] = usage

    return usage

def usage_limit_for(kind: str) -> int:
    limits = {
        "basic": 3,
        "cases": 2,
        "challenge": 1,
        "summary": 3,
    }
    return limits.get(kind, 0)

def is_admin_user(user_id: Optional[int]) -> bool:
    return bool(user_id and user_id in ADMIN_USER_IDS)

async def check_daily_limit(update: Update, context: ContextTypes.DEFAULT_TYPE, kind: str) -> bool:
    user = update.effective_user

    if is_admin_user(user.id if user else None):
        return True

    usage = usage_get(context.user_data)
    limit = usage_limit_for(kind)

    if usage.get(kind, 0) >= limit:
        await update.effective_message.reply_text(
            f"⚠️ وصلت للحد اليومي لهذا النوع.\n\n"
            f"الحد اليومي:\n"
            f"Basic: 3\n"
            f"Cases: 2\n"
            f"Challenge: 1\n"
            f"Summary: 3\n\n"
            f"يرجع العداد غداً."
        )
        return False

    return True

def increment_daily_usage(user_data: dict, kind: str) -> None:
    usage = usage_get(user_data)
    usage[kind] = usage.get(kind, 0) + 1


# =============================================================================
# OPENROUTER API
# =============================================================================
async def call_gemini(prompt: str, temperature: float = 0.2, max_output_tokens: int = 4096) -> str:
    async with _GEMINI_SEMAPHORE:
        def work():
            payload = {
                "model": GEMINI_MODEL_NAME,
                "messages": [{"role": "user", "content": prompt}],
                "temperature": temperature,
                "max_tokens": max_output_tokens,
            }
            req = urllib.request.Request(
                "https://openrouter.ai/api/v1/chat/completions",
                data=json.dumps(payload).encode("utf-8"),
                headers={
                    "Authorization": f"Bearer {OPENROUTER_API_KEY}",
                    "Content-Type": "application/json",
                    "HTTP-Referer": "https://telegram.org",
                    "X-Title": "Medical Telegram Bot",
                },
                method="POST",
            )
            try:
                print(f"--> [OpenRouter] Sending request. Tokens: {max_output_tokens}")
                with urllib.request.urlopen(req, timeout=120) as response:
                    raw_data = response.read().decode("utf-8")
                    data = json.loads(raw_data)
                    
                    if "choices" not in data:
                        print(f"❌ [OpenRouter] Error Format Received: {raw_data}")
                        return ""
                        
                    print("<-- [OpenRouter] Response received successfully!")
                    return data["choices"][0]["message"]["content"].strip()
                    
            except urllib.error.HTTPError as e:
                error_body = e.read().decode('utf-8')
                print(f"❌ [OpenRouter] HTTP ERROR {e.code}: {error_body}")
                raise Exception(f"HTTP {e.code}: {error_body[:200]}")
            except Exception as e:
                print(f"❌ [OpenRouter] UNKNOWN ERROR: {str(e)}")
                raise e
        return await asyncio.to_thread(work)


# =============================================================================
# TEXT EXTRACTION
# =============================================================================
def classify_document(mime: str, name: str) -> str:
    name = name.lower()
    if name.endswith('.pdf'):
        return "pdf"
    if name.endswith(('.docx', '.doc')):
        return "docx"
    if name.endswith(('.pptx', '.ppt')):
        return "pptx"
    if name.endswith('.txt'):
        return "txt"
    return "unsupported"

def extract_text(file_type: str, raw_bytes: bytes) -> str:
    f = io.BytesIO(raw_bytes)
    try:
        if file_type == "pdf":
            return "\n".join([page.extract_text() for page in PdfReader(f).pages if page.extract_text()])
        if file_type == "docx":
            return "\n".join([p.text for p in Document(f).paragraphs if p.text.strip()])
        if file_type == "pptx":
            prs = Presentation(f)
            parts = []
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        parts.append(shape.text)
            return "\n".join(parts)
        if file_type == "txt":
            return raw_bytes.decode('utf-8', errors='ignore')
    except Exception as e:
        logger.error(f"Extraction error for {file_type}: {e}")
        return ""
    return ""


# =============================================================================
# PROMPTS & GENERATION
# =============================================================================
LEVEL_LIMITS = {"basic": 50, "cases": 5, "challenge": 2}
BASIC_COUNTS = (5, 10, 20, 30, 40)

QUESTION_SCHEMA = """ Return valid JSON array only. No markdown. Each item must be: {
    "question": "...",
    "options": {"A":"...", "B":"...", "C":"...", "D":"...", "E":"..."},
    "correct": "A",
    "explanation": "..."
}
Exactly 5 options A-E. correct is one of A-E. Base only on source text."""

ACADEMIC_PREFIX = "Educational medical simulation. Do not provide real medical advice. "

LEVEL_PROMPTS = {
    "basic": f"{ACADEMIC_PREFIX}Generate direct medical recall MCQs. Straightforward factual questions. No patient vignettes. {QUESTION_SCHEMA}",
    "cases": f"{ACADEMIC_PREFIX}Generate USMLE Step 2 CK level CLINICAL VIGNETTES. Write detailed patient scenarios including age, gender, chief complaint, vital signs, relevant physical exam, and lab findings. The question must ask for the 'most likely diagnosis', 'next best step in management', or 'best initial test'. Distractors must be highly plausible differential diagnoses. {QUESTION_SCHEMA}",
    "challenge": f"{ACADEMIC_PREFIX}Generate EXTREMELY HARD, multi-step clinical reasoning MCQs (USMLE Step 1/Step 2 CK Hard tier). The student must first diagnose an atypical or complex presentation from a vignette, and then answer a 2nd or 3rd order question (e.g., underlying pathophysiology, specific pharmacological mechanism of action, embryological origin, or specific genetic mutation). All 5 options must be highly homogeneous and tricky, requiring deep analytical thinking to differentiate. {QUESTION_SCHEMA}",
}

SUMMARY_PROMPTS = {
    "disease_v2": f"{ACADEMIC_PREFIX}High-yield medical summary with headings: Overview, Etiology, Clinical, Diagnosis, Management, Complications, Pearls.",
    "highyield": f"{ACADEMIC_PREFIX}Concise high-yield notes, bullet points. End with Quick Recall.",
    "osce": f"{ACADEMIC_PREFIX}OSCE-style approach: History, Examination, Investigations, Management, Counseling, Red Flags.",
}


def extract_json_array(raw: str) -> Optional[list]:
    """Improved JSON extraction: strips markdown code fences, uses correct regex."""
    raw = (raw or "").strip()

    try:
        parsed = json.loads(raw)
        return parsed if isinstance(parsed, list) else None
    except Exception:
        pass

    raw = raw.replace("```json", "").replace("```", "").strip()

    match = re.search(r"\[.*\]", raw, flags=re.DOTALL)
    if match:
        try:
            parsed = json.loads(match.group(0))
            return parsed if isinstance(parsed, list) else None
        except Exception:
            return None

    return None


def normalize_questions(items: list) -> List[dict]:
    clean = []
    for q in items:
        if not isinstance(q, dict): continue
        question = str(q.get("question", "")).strip()
        options = q.get("options")
        correct = str(q.get("correct", "")).strip().upper()
        explanation = str(q.get("explanation", "")).strip()
        
        if not question or not isinstance(options, dict) or correct not in ["A", "B", "C", "D", "E"]: continue
        
        norm_opts = {}
        ok = True
        for letter in ["A", "B", "C", "D"]:
            val = str(options.get(letter, "")).strip()
            if not val:
                ok = False
                break
            norm_opts[letter] = val
            
        if not ok: continue
        
        e_val = str(options.get("E", "")).strip()
        if not e_val:
            norm_opts["E"] = "None of the above"
        else:
            norm_opts["E"] = e_val
            
        if correct not in norm_opts: continue
        clean.append({"question": question, "options": norm_opts, "correct": correct, "explanation": explanation or "Based on lecture text."})
    return clean

def shuffle_options(q: dict) -> dict:
    letters = list("ABCDE")
    opts = q.get("options", {})
    correct = q.get("correct", "").upper()
    if correct not in letters or any(l not in opts for l in letters): return q
    pairs = [(opts[l], l == correct) for l in letters]
    random.shuffle(pairs)
    new_opts = {}
    new_correct = "A"
    for idx, (text, is_correct) in enumerate(pairs):
        letter = letters[idx]
        new_opts[letter] = text
        if is_correct: new_correct = letter
    q = dict(q)
    q["options"] = new_opts
    q["correct"] = new_correct
    return q

def split_text(text: str, chunk_size: int = 10000) -> List[str]:
    return [text[i:i+chunk_size] for i in range(0, len(text), chunk_size)]


async def generate_question_bank(text: str, level: str) -> List[dict]:
    chunks = split_text(text, chunk_size=10000)
    all_questions = []
    target = LEVEL_LIMITS[level]

    for chunk in chunks:
        prompt = (
            f"{LEVEL_PROMPTS[level]}\n"
            f"Generate up to {target} questions only.\n"
            f"Return JSON array only. No markdown. No explanation outside JSON.\n\n"
            f"SOURCE TEXT:\n{chunk}"
        )

        raw = await call_gemini(prompt, temperature=0.15, max_output_tokens=8192)
        parsed = extract_json_array(raw)

        if parsed:
            all_questions.extend(normalize_questions(parsed))

        if len(all_questions) >= target:
            break

    return all_questions[:target]


async def get_or_create_question_bank(file_hash: str, text: str, level: str) -> List[dict]:
    """Return cached questions or generate them under a lock to avoid duplicate API calls."""
    lock = get_generation_lock(f"questions:{file_hash}:{level}")

    async with lock:
        cached = await db_get_questions(file_hash, level)

        if cached is not None:
            return cached

        questions = await generate_question_bank(text, level)

        if questions:
            await db_save_questions(file_hash, level, questions)

        return questions


async def get_or_create_summary(file_hash: str, text: str, style: str) -> str:
    """Return cached summary or generate it under a lock, with increased token limit."""
    lock = get_generation_lock(f"summary:{file_hash}:{style}")

    async with lock:
        cached = await db_get_summary(file_hash, style)

        if cached:
            return cached

        source = text[:MAX_TEXT_CHARS_FOR_AI]

        prompt = (
            f"{SUMMARY_PROMPTS.get(style, SUMMARY_PROMPTS['disease_v2'])}"
            f"\nSOURCE:\n{source}"
        )

        try:
            summary = await call_gemini(
                prompt,
                temperature=0.2,
                max_output_tokens=8192
            )
        except Exception as e:
            logger.error(
                f"Summary generation error for style {style}: {e}"
            )
            return ""

        if summary and summary.strip():
            await db_save_summary(file_hash, style, summary)

        return summary


# =============================================================================
# KEYBOARDS
# =============================================================================
def lang_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🇸🇦 العربية", callback_data="lang:ar"),
         InlineKeyboardButton("🇬🇧 English", callback_data="lang:en")]
    ])

def mode_keyboard(user_data: dict) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton(t(user_data, "mode_quiz"), callback_data="mode:quiz")],
        [InlineKeyboardButton(t(user_data, "mode_summary"), callback_data="mode:summary")],
        [InlineKeyboardButton("📞 الدعم", callback_data="back:support")]
    ])

def summary_style_keyboard(user_data: dict) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton(t(user_data, "style_disease"), callback_data="style:disease_v2")],
        [InlineKeyboardButton(t(user_data, "style_highyield"), callback_data="style:highyield")],
        [InlineKeyboardButton(t(user_data, "style_osce"), callback_data="style:osce")],
        [InlineKeyboardButton("⬅️ رجوع", callback_data="back:mode")]
    ])

def level_keyboard(setup_id: str, user_data: dict) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton(t(user_data, "level_basic"), callback_data=f"level:{setup_id}:basic")],
        [InlineKeyboardButton(t(user_data, "level_cases"), callback_data=f"level:{setup_id}:cases")],
        [InlineKeyboardButton(t(user_data, "level_challenge"), callback_data=f"level:{setup_id}:challenge")],
        [InlineKeyboardButton("⬅️ رجوع", callback_data="back:mode")]
    ])

def basic_count_keyboard(setup_id: str) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton("5", callback_data=f"count:{setup_id}:5"),
            InlineKeyboardButton("10", callback_data=f"count:{setup_id}:10"),
            InlineKeyboardButton("20", callback_data=f"count:{setup_id}:20"),
        ],
        [
            InlineKeyboardButton("30", callback_data=f"count:{setup_id}:30"),
            InlineKeyboardButton("40", callback_data=f"count:{setup_id}:40"),
        ],
        [InlineKeyboardButton("⬅️ رجوع", callback_data=f"back:levels:{setup_id}")]
    ])

def answer_keyboard(quiz_id: str, idx: int) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([[InlineKeyboardButton(f" {l} ", callback_data=f"ans:{quiz_id}:{idx}:{l}") for l in "ABCDE"]])

def post_quiz_keyboard(user_data: dict, has_wrong: bool) -> InlineKeyboardMarkup:
    rows = []
    if has_wrong:
        rows.append([InlineKeyboardButton(t(user_data, "review"), callback_data="review:start")])
    rows.append([InlineKeyboardButton(t(user_data, "stats"), callback_data="review:stats")])
    rows.append([InlineKeyboardButton(t(user_data, "new_quiz"), callback_data="review:new")])
    return InlineKeyboardMarkup(rows)

def review_keyboard(current: int, total: int) -> InlineKeyboardMarkup:
    buttons = []
    if current < total - 1:
        buttons.append(InlineKeyboardButton("Next ▶", callback_data="review:next"))
    buttons.append(InlineKeyboardButton("Done ✖", callback_data="review:exit"))
    return InlineKeyboardMarkup([buttons])


# =============================================================================
# FORMATTING
# =============================================================================
def format_question(q: dict, idx: int, total: int, user_data: dict) -> str:
    opts = q.get("options", {})
    return f"{progress_bar(idx, total)}\n{t(user_data, 'q_meta', i=idx, n=total)}\n\n{escape_md(q.get('question', ''))}\n\nA. {escape_md(opts.get('A', ''))}\nB. {escape_md(opts.get('B', ''))}\nC. {escape_md(opts.get('C', ''))}\nD. {escape_md(opts.get('D', ''))}\nE. {escape_md(opts.get('E', ''))}"

def format_review(entry: dict, idx: int, total: int, user_data: dict) -> str:
    q = entry["question"]
    chosen = entry["chosen"]
    correct = q["correct"]
    opts = q["options"]
    return f"📋 Review {idx}/{total}\n\n{escape_md(q['question'])}\n\nA. {escape_md(opts['A'])}\nB. {escape_md(opts['B'])}\nC. {escape_md(opts['C'])}\nD. {escape_md(opts['D'])}\nE. {escape_md(opts['E'])}\n\nYour answer: {chosen}\nCorrect answer: {correct}. {escape_md(opts[correct])}\n\n{t(user_data, 'explanation')} {escape_md(q.get('explanation', ''))}"


# =============================================================================
# SENDERS
# =============================================================================
async def safe_send(bot, chat_id: int, text: str, **kwargs) -> bool:
    for attempt in range(2):
        try:
            await bot.send_message(chat_id=chat_id, text=text, **kwargs)
            return True
        except Exception as e:
            logger.warning("send failed: %s", e)
            await asyncio.sleep(0.5)
    return False

async def send_current_question(chat_id: int, context: ContextTypes.DEFAULT_TYPE) -> None:
    session = context.user_data.get("session")
    if not session: return
    idx = session["current"]
    questions = session["questions"]
    if idx >= len(questions):
        await send_final_score(chat_id, context)
        return
    text = format_question(questions[idx], idx+1, len(questions), context.user_data)
    await safe_send(context.bot, chat_id, text, parse_mode="Markdown", reply_markup=answer_keyboard(session["quiz_id"], idx))

async def send_final_score(chat_id: int, context: ContextTypes.DEFAULT_TYPE) -> None:
    session = context.user_data.get("session")
    if not session: return
    total = len(session["questions"])
    score = session["score"]
    pct = round(score/total*100) if total else 0
    wrong = session.get("wrong", [])
    update_stats(context.user_data, score, total)
    text = t(context.user_data, "complete", score=score, total=total, pct=pct)
    await context.bot.send_message(chat_id=chat_id, text=text, reply_markup=post_quiz_keyboard(context.user_data, bool(wrong)))
    context.user_data["review_pool"] = wrong
    context.user_data.pop("session", None)


# =============================================================================
# COMMANDS HANDLERS
# =============================================================================
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(t(context.user_data, "choose_lang"), reply_markup=lang_keyboard())

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(t(context.user_data, "help"))

async def language_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(t(context.user_data, "choose_lang"), reply_markup=lang_keyboard())

async def stop_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    had = any(k in context.user_data for k in ("session", "pending_file", "review"))
    for k in ("session", "pending_file", "pending_mode", "pending_summary_style", "review", "review_pool", "busy"):
        context.user_data.pop(k, None)
    await update.message.reply_text(t(context.user_data, "stopped" if had else "no_active"))

async def stats_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    stats = get_stats(context.user_data)
    total = stats["total"]
    accuracy = round(stats["correct"]/total*100) if total else 0
    await update.message.reply_text(t(context.user_data, "stats_text", quizzes=stats["quizzes"], correct=stats["correct"], total=total, accuracy=accuracy, best=stats["best"]))

async def review_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    pool = context.user_data.get("review_pool", [])
    if not pool:
        await update.message.reply_text(t(context.user_data, "review_none"))
        return
    context.user_data["review"] = {"pool": pool, "current": 0}
    await send_review(update.effective_chat.id, context)

async def support_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("💳 دعم عبر رصيد المدار", callback_data="support:balance")],
        [InlineKeyboardButton("💬 تواصل مع الدعم", callback_data="support:contact")],
        [InlineKeyboardButton("📢 أبلغ عن مشكلة", callback_data="support:problem")],
        [InlineKeyboardButton("⬅️ رجوع", callback_data="back:mode")]
    ])

    await update.message.reply_text(
        "📞 مركز الدعم\n\n"
        "اختر نوع الدعم الذي تحتاجه:",
        reply_markup=keyboard
    )

async def testbot_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if not user or user.id not in ADMIN_USER_IDS:
        await update.message.reply_text("Unauthorized.")
        return
    try:
        reply = await call_gemini("Reply with exactly: OK", temperature=0.0, max_output_tokens=10)
        await update.message.reply_text(f"✅ OpenRouter response: {reply[:50]}")
    except Exception as e:
        await update.message.reply_text(f"❌ OpenRouter failed: {e}")


# =============================================================================
# CALLBACK HANDLERS
# =============================================================================
async def lang_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    _, lang = query.data.split(":", 1)
    if lang not in SUPPORTED_LANGS: return
    context.user_data["lang"] = lang
    await query.edit_message_text(t(context.user_data, "welcome"))
    await context.bot.send_message(chat_id=update.effective_chat.id, text=t(context.user_data, "choose_mode"), reply_markup=mode_keyboard(context.user_data))

async def mode_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    _, mode = query.data.split(":", 1)
    if mode not in ("quiz", "summary"): return
    context.user_data["pending_mode"] = mode
    await query.edit_message_reply_markup(reply_markup=None)
    if mode == "summary":
        await context.bot.send_message(chat_id=update.effective_chat.id, text=t(context.user_data, "summary_style_prompt"), reply_markup=summary_style_keyboard(context.user_data))
    else:
        await context.bot.send_message(chat_id=update.effective_chat.id, text=t(context.user_data, "mode_quiz_prompt"))

async def summary_style_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    _, style = query.data.split(":", 1)
    if style not in SUMMARY_PROMPTS: return
    context.user_data["pending_summary_style"] = style
    context.user_data["pending_mode"] = "summary"
    await query.edit_message_reply_markup(reply_markup=None)
    await context.bot.send_message(chat_id=update.effective_chat.id, text=t(context.user_data, "mode_summary_prompt"))

async def level_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    parts = query.data.split(":")
    if len(parts) != 3: return
    _, setup_id, level = parts
    pending = context.user_data.get("pending_file")
    if not pending or pending.get("setup_id") != setup_id:
        await query.edit_message_text(t(context.user_data, "stale"))
        return
    if level not in LEVEL_LIMITS: return

    if not await check_daily_limit(update, context, level):
        return

    pending["level"] = level
    if level == "basic":
        await query.edit_message_text(t(context.user_data, "basic_count_prompt"), reply_markup=basic_count_keyboard(setup_id))
    else:
        await query.edit_message_text(t(context.user_data, "generating_pool"))
        await start_quiz_from_bank(update, context, setup_id, level, LEVEL_LIMITS[level])

async def count_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    parts = query.data.split(":")
    if len(parts) != 3: return
    _, setup_id, count_s = parts
    try:
        count = int(count_s)
    except ValueError:
        return
    if count not in BASIC_COUNTS: return
    pending = context.user_data.get("pending_file")
    if not pending or pending.get("setup_id") != setup_id:
        await query.edit_message_text(t(context.user_data, "stale"))
        return
    await query.edit_message_text(t(context.user_data, "generating_pool"))
    await start_quiz_from_bank(update, context, setup_id, "basic", count)

async def start_quiz_from_bank(update: Update, context: ContextTypes.DEFAULT_TYPE, setup_id: str, level: str, requested_count: int):
    query = update.callback_query
    pending = context.user_data.get("pending_file")
    if not pending or pending.get("setup_id") != setup_id:
        await query.edit_message_text(t(context.user_data, "stale"))
        return
    try:
        questions = await get_or_create_question_bank(pending["file_hash"], pending["text"], level)
        if not questions:
            await query.edit_message_text(t(context.user_data, "no_questions"))
            return
        pool = [shuffle_options(q) for q in questions]
        random.shuffle(pool)
        selected = pool[:min(requested_count, len(pool))]
        if len(selected) < requested_count:
            await context.bot.send_message(chat_id=update.effective_chat.id, text=t(context.user_data, "not_enough", n=len(selected)))
        quiz_id = str(uuid.uuid4())[:8]
        context.user_data["session"] = {"quiz_id": quiz_id, "questions": selected, "current": 0, "score": 0, "wrong": [], "level": level}
        context.user_data.pop("pending_file", None)
        await query.edit_message_text(t(context.user_data, "quiz_ready", n=len(selected)))
        await send_current_question(update.effective_chat.id, context)
        increment_daily_usage(context.user_data, level)
    except Exception as e:
        logger.exception("quiz start failed")
        await query.edit_message_text(t(context.user_data, "error", err=str(e)))

async def answer_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    parts = query.data.split(":")
    if len(parts) != 4: return
    _, quiz_id, idx_s, chosen = parts
    try:
        idx = int(idx_s)
    except ValueError:
        return
    session = context.user_data.get("session")
    if not session or session.get("quiz_id") != quiz_id or session.get("current") != idx:
        await query.edit_message_reply_markup(reply_markup=None)
        return
    questions = session["questions"]
    q = questions[idx]
    correct = q["correct"]
    opts = q["options"]
    is_correct = chosen == correct
    if is_correct:
        session["score"] += 1
        result = t(context.user_data, "correct")
    else:
        result = t(context.user_data, "wrong", chosen=chosen)
        session["wrong"].append({"question": q, "chosen": chosen})
    feedback = format_question(q, idx+1, len(questions), context.user_data) + "\n\n" + f"{escape_md(result)}\n{t(context.user_data, 'answer')} {correct}. {escape_md(opts.get(correct, ''))}\n\n{t(context.user_data, 'explanation')} {escape_md(q.get('explanation', ''))}"
    try:
        await query.edit_message_text(feedback, parse_mode="Markdown")
    except:
        await query.edit_message_reply_markup(reply_markup=None)
        await context.bot.send_message(update.effective_chat.id, feedback.replace("*", ""))
    session["current"] += 1
    await asyncio.sleep(0.4)
    if session["current"] >= len(questions):
        await send_final_score(update.effective_chat.id, context)
    else:
        await send_current_question(update.effective_chat.id, context)

async def review_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    _, action = query.data.split(":", 1)
    if action == "start":
        pool = context.user_data.get("review_pool", [])
        if not pool:
            await query.edit_message_text(t(context.user_data, "review_none"))
            return
        context.user_data["review"] = {"pool": pool, "current": 0}
        await query.edit_message_reply_markup(reply_markup=None)
        await send_review(update.effective_chat.id, context)
    elif action == "next":
        review = context.user_data.get("review")
        if review:
            review["current"] += 1
            await query.edit_message_reply_markup(reply_markup=None)
            await send_review(update.effective_chat.id, context)
    elif action == "exit":
        context.user_data.pop("review", None)
        await query.edit_message_reply_markup(reply_markup=None)
        await context.bot.send_message(update.effective_chat.id, t(context.user_data, "review_done"))
    elif action == "stats":
        await query.edit_message_reply_markup(reply_markup=None)
        stats = get_stats(context.user_data)
        total = stats["total"]
        accuracy = round(stats["correct"]/total*100) if total else 0
        await context.bot.send_message(update.effective_chat.id, t(context.user_data, "stats_text", quizzes=stats["quizzes"], correct=stats["correct"], total=total, accuracy=accuracy, best=stats["best"]))
    elif action == "new":
        await query.edit_message_reply_markup(reply_markup=None)
        await context.bot.send_message(update.effective_chat.id, t(context.user_data, "choose_mode"), reply_markup=mode_keyboard(context.user_data))

async def send_review(chat_id: int, context: ContextTypes.DEFAULT_TYPE):
    review = context.user_data.get("review")
    if not review: return
    pool = review["pool"]
    current = review["current"]
    if current >= len(pool):
        context.user_data.pop("review", None)
        await context.bot.send_message(chat_id, t(context.user_data, "review_done"))
        return
    await context.bot.send_message(chat_id=chat_id, text=format_review(pool[current], current+1, len(pool), context.user_data), parse_mode="Markdown", reply_markup=review_keyboard(current, len(pool)))


# =============================================================================
# BACK NAVIGATION & SUPPORT CALLBACKS
# =============================================================================
async def back_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    parts = query.data.split(":")
    target = parts[1] if len(parts) > 1 else "mode"

    if target == "mode":
        await query.edit_message_text(
            t(context.user_data, "choose_mode"),
            reply_markup=mode_keyboard(context.user_data)
        )

    elif target == "support":
        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("💳 دعم عبر رصيد المدار", callback_data="support:balance")],
            [InlineKeyboardButton("💬 تواصل مع الدعم", callback_data="support:contact")],
            [InlineKeyboardButton("📢 أبلغ عن مشكلة", callback_data="support:problem")],
            [InlineKeyboardButton("⬅️ رجوع", callback_data="back:mode")]
        ])

        await query.edit_message_text(
            "📞 مركز الدعم\n\n"
            "اختر نوع الدعم الذي تحتاجه:",
            reply_markup=keyboard
        )

    elif target == "levels":
        if len(parts) != 3:
            return
        setup_id = parts[2]
        pending = context.user_data.get("pending_file")
        if not pending or pending.get("setup_id") != setup_id:
            await query.edit_message_text(t(context.user_data, "stale"))
            return
        await query.edit_message_text(
            t(context.user_data, "file_ready_quiz"),
            reply_markup=level_keyboard(setup_id, context.user_data)
        )

async def support_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == "support:balance":
        await query.edit_message_text(
            "💳 دعم عبر رصيد المدار\n\n"
            "رقم المدار:\n"
            "0918874659\n\n"
            "بعد إرسال الرصيد، اضغط تواصل مع الدعم واكتب اسمك أو رقم العملية.",
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("💬 تواصل مع الدعم", callback_data="support:contact")],
                [InlineKeyboardButton("⬅️ رجوع", callback_data="back:support")]
            ])
        )

    elif query.data == "support:contact":
        context.user_data["awaiting_support_message"] = "contact"
        await query.edit_message_text(
            "💬 اكتب رسالتك الآن.\n\n"
            "مثال:\n"
            "اسمي أحمد، أرسلت رصيد، وأريد تفعيل الخدمة."
        )

    elif query.data == "support:problem":
        context.user_data["awaiting_support_message"] = "problem"
        await query.edit_message_text(
            "📢 اكتب المشكلة التي تواجهك الآن.\n\n"
            "يفضل تكتب:\n"
            "- نوع الملف\n"
            "- ماذا ضغطت\n"
            "- ما الخطأ الذي ظهر لك"
        )


# =============================================================================
# ADMIN LOG
# =============================================================================
async def send_admin_log(context: ContextTypes.DEFAULT_TYPE, text: str):
    try:
        await context.bot.send_message(
            chat_id=ADMIN_LOG_CHAT_ID,
            text=text
        )
    except Exception as e:
        logger.warning("Admin log failed: %s", e)


# =============================================================================
# FILE HANDLING
# =============================================================================
async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if user and await cooldown_guard(update, context.user_data, user.id):
        return
    if context.user_data.get("busy"):
        await update.message.reply_text(t(context.user_data, "busy"))
        return
    context.user_data["busy"] = True
    try:
        doc = update.message.document
        if not doc:
            await update.message.reply_text("⚠️ أرسل ملفاً (PDF, DOCX, PPTX, TXT) وليس صورة.")
            return

        await send_admin_log(
            context,
            f"📥 ملف جديد\n"
            f"User: {user.id if user else 'unknown'}\n"
            f"Name: {user.full_name if user else 'unknown'}\n"
            f"File: {doc.file_name if doc else 'unknown'}\n"
            f"Size: {doc.file_size if doc else 'unknown'} bytes"
        )

        file_type = classify_document(doc.mime_type or "", doc.file_name or "")
        if file_type == "unsupported":
            await update.message.reply_text(t(context.user_data, "file_unsupported"))
            return

        if doc.file_size and doc.file_size > MAX_FILE_SIZE_BYTES:
            await update.message.reply_text(t(context.user_data, "file_too_large"))
            return

        status = await update.message.reply_text(t(context.user_data, "file_received"))
        tg_file = await doc.get_file()
        raw_bytes = bytes(await tg_file.download_as_bytearray())
        file_hash = sha256_bytes(raw_bytes)

        cached_file = await db_get_file(file_hash)
        if cached_file:
            text = cached_file["extracted_text"]
            await status.edit_text(t(context.user_data, "cached_file"))
            await asyncio.sleep(0.4)
        else:
            text = extract_text(file_type, raw_bytes)
            if not text or len(text.strip()) < 100:
                await status.edit_text(t(context.user_data, "file_empty"))
                return
            await db_save_file(file_hash, doc.file_name or "document", len(raw_bytes), doc.mime_type or "", text)

        mode = context.user_data.get("pending_mode", "quiz")
        if mode == "summary":
            if not await check_daily_limit(update, context, "summary"):
                return

            style = context.user_data.get("pending_summary_style", "disease_v2")
            await status.edit_text(t(context.user_data, "generating_summary"))
            summary = await get_or_create_summary(file_hash, text, style)
            await status.delete()
            
            if summary:
                full_text = t(context.user_data, "summary_header") + summary
                
                # تقسيم النص لتجاوز حد رسائل تيليجرام (4096 حرف)
                chunks = [full_text[i:i+4000] for i in range(0, len(full_text), 4000)]
                for chunk in chunks:
                    await update.message.reply_text(chunk)
                    
                increment_daily_usage(context.user_data, "summary")
            else:
                await update.message.reply_text(t(context.user_data, "summary_failed"))
            return

        setup_id = str(uuid.uuid4())[:8]
        context.user_data["pending_mode"] = "quiz"
        context.user_data["pending_file"] = {
            "setup_id": setup_id,
            "file_hash": file_hash,
            "file_name": doc.file_name or "document",
            "text": text
        }
        await status.edit_text(t(context.user_data, "file_ready_quiz"), reply_markup=level_keyboard(setup_id, context.user_data))

    except Exception as e:
        logger.exception("handle_file error")
        await update.message.reply_text("❌ حدث خطأ أثناء معالجة الملف. تأكد أن الملف ليس تالفاً.")
    finally:
        context.user_data["busy"] = False


# =============================================================================
# TEXT HANDLER (supports admin forwarding)
# =============================================================================
async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    text = update.message.text

    support_type = context.user_data.get("awaiting_support_message")
    if support_type:
        context.user_data.pop("awaiting_support_message", None)

        kind = "💬 تواصل مع الدعم" if support_type == "contact" else "📢 بلاغ مشكلة"

        await send_admin_log(
            context,
            f"{kind}\n\n"
            f"User ID: {user.id if user else 'unknown'}\n"
            f"Name: {user.full_name if user else 'unknown'}\n"
            f"Username: @{user.username if user and user.username else 'لا يوجد'}\n\n"
            f"Message:\n{text}"
        )

        await update.message.reply_text(
            "✅ تم إرسال رسالتك للدعم.\n"
            "سيتم التواصل معك قريباً إذا احتاج الأمر."
        )
        return

    if context.user_data.get("session"):
        await update.message.reply_text(t(context.user_data, "quiz_in_progress"))
        return

    await update.message.reply_text(t(context.user_data, "send_file_hint"))


async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    logger.error("Unhandled exception", exc_info=context.error)
    if isinstance(update, Update) and update.effective_message:
        try:
            await update.effective_message.reply_text(t(context.user_data or {}, "error", err=str(context.error)))
        except: pass


# =============================================================================
# MAIN
# =============================================================================
BOT_COMMANDS = [
    BotCommand("start", "Start bot"),
    BotCommand("help", "Help"),
    BotCommand("stop", "Stop active quiz"),
    BotCommand("stats", "Show stats"),
    BotCommand("review", "Review mistakes"),
    BotCommand("language", "Change language"),
    BotCommand("support", "الدعم"),
    BotCommand("testbot", "Admin test"),
]

async def post_init(application: Application):
    init_db()
    await cleanup_old_cache(days=30)
    await application.bot.set_my_commands(BOT_COMMANDS)
    logger.info(
        "Bot ready. Model=%s DB=%s MaxFile=%sMB",
        GEMINI_MODEL_NAME,
        DATABASE_PATH,
        MAX_FILE_SIZE_BYTES // 1024 // 1024
    )

def main():
    persistence = PicklePersistence(
        filepath="/workspace/bot_user_data.pkl"
    )
    application = (
        ApplicationBuilder()
        .token(TOKEN)
        .persistence(persistence)
        .concurrent_updates(True)
        .post_init(post_init)
        .build()
    )
    application.add_error_handler(error_handler)
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(CommandHandler("stop", stop_command))
    application.add_handler(CommandHandler("stats", stats_command))
    application.add_handler(CommandHandler("review", review_command))
    application.add_handler(CommandHandler("language", language_command))
    application.add_handler(CommandHandler("support", support_command))
    application.add_handler(CommandHandler("testbot", testbot_command))
    application.add_handler(CallbackQueryHandler(lang_callback, pattern=r"^lang:"))
    application.add_handler(CallbackQueryHandler(mode_callback, pattern=r"^mode:"))
    application.add_handler(CallbackQueryHandler(summary_style_callback, pattern=r"^style:"))
    application.add_handler(CallbackQueryHandler(level_callback, pattern=r"^level:"))
    application.add_handler(CallbackQueryHandler(count_callback, pattern=r"^count:"))
    application.add_handler(CallbackQueryHandler(answer_callback, pattern=r"^ans:"))
    application.add_handler(CallbackQueryHandler(review_callback, pattern=r"^review:"))
    application.add_handler(CallbackQueryHandler(support_callback, pattern=r"^support:"))
    application.add_handler(CallbackQueryHandler(back_callback, pattern=r"^back:"))

    application.add_handler(MessageHandler(filters.Document.ALL, handle_file))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    logger.info("Starting bot polling...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
