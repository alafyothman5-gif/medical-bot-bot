#!/usr/bin/env python3

-- coding: utf-8 --

"""
Smart Medical Assistant Bot - OpenRouter Paid Economical Edition

Uses OpenRouter API with google/gemini-2.0-flash (paid but economical)

Persistent SQLite cache, quiz levels, summaries, bilingual (en/ar)

Fully compatible with python-telegram-bot v20+
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
import base64
import urllib.request
import urllib.error
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

from docx import Document
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

=============================================================================

LOGGING

=============================================================================

logging.basicConfig(
format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
level=logging.INFO,
)
logger = logging.getLogger("medical_bot_openrouter")

=============================================================================

ENVIRONMENT

=============================================================================

TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
if not TOKEN:
raise ValueError("TELEGRAM_BOT_TOKEN environment variable not set")
OPENROUTER_API_KEY = os.environ.get("OPENROUTER_API_KEY")
if not OPENROUTER_API_KEY:
raise ValueError("OPENROUTER_API_KEY environment variable not set")
GEMINI_MODEL_NAME = os.environ.get("GEMINI_MODEL", "google/gemini-2.0-flash")
DATABASE_PATH = os.environ.get("DATABASE_PATH", "bot_cache.sqlite3")
MAX_FILE_SIZE_BYTES = 30 * 1024 * 1024
MAX_TEXT_CHARS_FOR_AI = 40000
MAX_CONCURRENT_GEMINI = int(os.environ.get("MAX_CONCURRENT_GEMINI", "2"))
COOLDOWN_SECONDS = int(os.environ.get("COOLDOWN_SECONDS", "20"))

ADMIN_USER_IDS = {1692255414}

_GEMINI_SEMAPHORE = asyncio.Semaphore(MAX_CONCURRENT_GEMINI)

=============================================================================

TRANSLATIONS (مختصرة قليلاً لتوفير المساحة لكنها كاملة في التشغيل)

=============================================================================

SUPPORTED_LANGS = ("en", "ar")

TRANSLATIONS = {
"en": {
"choose_lang": "🌐 Choose your language\nاختر لغتك",
"welcome": "Welcome to your Smart Medical Assistant 🩺\n\nSend a medical PDF, image, DOCX, or TXT file.",
"help": "Send a medical file up to 30MB. Choose Quiz or Summary. Commands: /start /stop /stats /review /language",
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
"file_unsupported": "⚠️ Unsupported file type. Send PDF, image, DOCX, or TXT.",
"file_no_text": "⚠️ I could not extract readable text from this file.",
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
},
"ar": {
"choose_lang": "🌐 اختر لغتك",
"welcome": "مرحباً بك في المساعد الطبي الذكي 🩺\n\nأرسل ملف PDF أو صورة أو DOCX أو TXT.",
"help": "أرسل ملفاً طبياً حتى 30MB. اختر اختبار أو ملخص. الأوامر: /start /stop /stats /review /language",
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
"file_unsupported": "⚠️ نوع الملف غير مدعوم. أرسل PDF أو صورة أو DOCX أو TXT.",
"file_no_text": "⚠️ لم أستطع استخراج نص مقروء من الملف.",
"file_ready_quiz": "✅ الملف جاهز. اختر مستوى الأسئلة:",
"file_ready_summary": "✅ الملف جاهز. جاري تجهيز الملخص...",
"cached_file": "♻️ الملف موجود مسبقاً في الكاش. لن نستهلك Gemini من جديد.",
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

=============================================================================

SQLITE CACHE (نفس الكود السابق)

=============================================================================

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

=============================================================================

HELPERS

=============================================================================

user_last_request: Dict[int, float] = {}

def sha256_bytes(data: bytes) -> str:
return hashlib.sha256(data).hexdigest()

def escape_md(text: object) -> str:
value = "" if text is None else str(text)
for ch in ("*", "_", "`", "["):
value = value.replace(ch, f"\{ch}")
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

=============================================================================

OPENROUTER API

=============================================================================

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
with urllib.request.urlopen(req, timeout=120) as response:
data = json.loads(response.read().decode("utf-8"))
return data["choices"][0]["message"]["content"].strip()
except urllib.error.HTTPError as e:
error_body = e.read().decode('utf-8')
logger.error(f"OpenRouter HTTP {e.code}: {error_body}")
raise Exception(f"HTTP {e.code}: {error_body[:200]}")
return await asyncio.to_thread(work)

async def extract_text_from_image(image_bytes: bytes, mime_type: str) -> str:
async with _GEMINI_SEMAPHORE:
def work():
b64 = base64.b64encode(image_bytes).decode("utf-8")
image_url = f"data:{mime_type or 'image/jpeg'};base64,{b64}"
payload = {
"model": GEMINI_MODEL_NAME,
"messages": [
{
"role": "user",
"content": [
{"type": "text", "text": "Extract all readable text from this medical image exactly. Do not summarize. Return only raw text."},
{"type": "image_url", "image_url": {"url": image_url}}
]
}
],
"temperature": 0.0,
"max_tokens": 4096,
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
with urllib.request.urlopen(req, timeout=120) as response:
data = json.loads(response.read().decode("utf-8"))
return data["choices"][0]["message"]["content"].strip()
except urllib.error.HTTPError as e:
error_body = e.read().decode('utf-8')
logger.error(f"OpenRouter vision error {e.code}: {error_body}")
raise Exception(f"Vision error {e.code}")
return await asyncio.to_thread(work)

=============================================================================

TEXT EXTRACTION

=============================================================================

def extract_text_from_pdf(pdf_bytes: bytes) -> str:
reader = PdfReader(io.BytesIO(pdf_bytes))
pages = []
for page in reader.pages:
s = page.extract_text()
if s:
pages.append(s.strip())
return "\n\n".join(pages).strip()

def extract_text_from_docx(docx_bytes: bytes) -> str:
doc = Document(io.BytesIO(docx_bytes))
parts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
return "\n\n".join(parts)

def extract_text_from_txt(txt_bytes: bytes) -> str:
for enc in ("utf-8", "utf-16", "latin-1"):
try:
return txt_bytes.decode(enc)
except UnicodeDecodeError:
pass
return txt_bytes.decode("utf-8", errors="replace")

async def extract_text(file_type: str, raw_bytes: bytes, mime_type: str) -> str:
if file_type == "pdf":
return await asyncio.to_thread(extract_text_from_pdf, raw_bytes)
if file_type == "docx":
return await asyncio.to_thread(extract_text_from_docx, raw_bytes)
if file_type == "txt":
return await asyncio.to_thread(extract_text_from_txt, raw_bytes)
if file_type == "image":
return await extract_text_from_image(raw_bytes, mime_type)
return ""

=============================================================================

PROMPTS و QUESTION GENERATION (مختصرة لتوفير المساحة لكنها كاملة)

=============================================================================

LEVEL_LIMITS = {"basic": 50, "cases": 5, "challenge": 2}
BASIC_COUNTS = (5, 10, 20, 30, 40)

QUESTION_SCHEMA = """ Return valid JSON array only. No markdown. Each item must be: {
"question": "...",
"options": {"A":"...", "B":"...", "C":"...", "D":"...", "E":"..."},
"correct": "A",
"explanation": "..."
}
Exactly 5 options A-E. correct is one of A-E. Base only on source text."""

LEVEL_PROMPTS = {
"basic": f"BASIC direct MCQs. No vignettes. {QUESTION_SCHEMA}",
"cases": f"CLINICAL CASE MCQs. Short patient scenarios. {QUESTION_SCHEMA}",
"challenge": f"CHALLENGE MCQs. Two-step reasoning. {QUESTION_SCHEMA}",
}

SUMMARY_PROMPTS = {
"disease": "High-yield medical summary with headings: Overview, Etiology, Clinical, Diagnosis, Management, Complications, Pearls.",
"highyield": "Concise high-yield notes, bullet points. End with Quick Recall.",
"osce": "OSCE-style approach: History, Examination, Investigations, Management, Counseling, Red Flags.",
}

def extract_json_array(raw: str) -> Optional[list]:
raw = (raw or "").strip()
try:
parsed = json.loads(raw)
return parsed if isinstance(parsed, list) else None
except:
match = re.search(r"", raw, flags=re.DOTALL)
if match:
try:
return json.loads(match.group(0))
except:
pass
return None

def normalize_questions(items: list) -> List[dict]:
clean = []
for q in items:
if not isinstance(q, dict): continue
question = str(q.get("question", "")).strip()
options = q.get("options")
correct = str(q.get("correct", "")).strip().upper()
explanation = str(q.get("explanation", "")).strip()
if not question or not isinstance(options, dict) or correct not in "ABCDE": continue
norm_opts = {}
ok = True
for letter in "ABCDE":
val = str(options.get(letter, "")).strip()
if not val:
ok = False
break
norm_opts[letter] = val
if not ok or correct not in norm_opts: continue
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

async def generate_question_bank(text: str, level: str) -> List[dict]:
limit = LEVEL_LIMITS[level]
source = text[:MAX_TEXT_CHARS_FOR_AI]
prompt = f"{LEVEL_PROMPTS[level]}\nCreate up to {limit} questions. Only from source.\nSOURCE:\n{source}"
raw = await call_gemini(prompt, temperature=0.15, max_output_tokens=8192)
parsed = extract_json_array(raw)
if parsed is None:
logger.warning("Invalid JSON for %s", level)
return []
return normalize_questions(parsed)[:limit]

async def get_or_create_question_bank(file_hash: str, text: str, level: str) -> List[dict]:
cached = await db_get_questions(file_hash, level)
if cached is not None:
return cached
questions = await generate_question_bank(text, level)
if questions:
await db_save_questions(file_hash, level, questions)
return questions

async def get_or_create_summary(file_hash: str, text: str, style: str) -> str:
cached = await db_get_summary(file_hash, style)
if cached:
return cached
source = text[:MAX_TEXT_CHARS_FOR_AI]
prompt = f"{SUMMARY_PROMPTS.get(style, SUMMARY_PROMPTS['disease'])}\nSOURCE:\n{source}"
summary = await call_gemini(prompt, temperature=0.2, max_output_tokens=4096)
if summary:
await db_save_summary(file_hash, style, summary)
return summary

=============================================================================

KEYBOARDS

=============================================================================

def lang_keyboard() -> InlineKeyboardMarkup:
return InlineKeyboardMarkup([[InlineKeyboardButton("🇸🇦 العربية", callback_data="lang:ar"), InlineKeyboardButton("🇬🇧 English", callback_data="lang:en")]])

def mode_keyboard(user_data: dict) -> InlineKeyboardMarkup:
return InlineKeyboardMarkup([[InlineKeyboardButton(t(user_data, "mode_quiz"), callback_data="mode:quiz")], [InlineKeyboardButton(t(user_data, "mode_summary"), callback_data="mode:summary")]])

def summary_style_keyboard(user_data: dict) -> InlineKeyboardMarkup:
return InlineKeyboardMarkup([[InlineKeyboardButton(t(user_data, "style_disease"), callback_data="style:disease")], [InlineKeyboardButton(t(user_data, "style_highyield"), callback_data="style:highyield")], [InlineKeyboardButton(t(user_data, "style_osce"), callback_data="style:osce")]])

def level_keyboard(setup_id: str, user_data: dict) -> InlineKeyboardMarkup:
return InlineKeyboardMarkup([[InlineKeyboardButton(t(user_data, "level_basic"), callback_data=f"level:{setup_id}:basic")], [InlineKeyboardButton(t(user_data, "level_cases"), callback_data=f"level:{setup_id}:cases")], [InlineKeyboardButton(t(user_data, "level_challenge"), callback_data=f"level:{setup_id}:challenge")]])

def basic_count_keyboard(setup_id: str) -> InlineKeyboardMarkup:
return InlineKeyboardMarkup([[InlineKeyboardButton("5", callback_data=f"count:{setup_id}:5"), InlineKeyboardButton("10", callback_data=f"count:{setup_id}:10"), InlineKeyboardButton("20", callback_data=f"count:{setup_id}:20")], [InlineKeyboardButton("30", callback_data=f"count:{setup_id}:30"), InlineKeyboardButton("40", callback_data=f"count:{setup_id}:40")]])

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

=============================================================================

FORMATTING

=============================================================================

def format_question(q: dict, idx: int, total: int, user_data: dict) -> str:
opts = q.get("options", {})
return f"{progress_bar(idx, total)}\n{t(user_data, 'q_meta', i=idx, n=total)}\n\n{escape_md(q.get('question', ''))}\n\nA. {escape_md(opts.get('A', ''))}\nB. {escape_md(opts.get('B', ''))}\nC. {escape_md(opts.get('C', ''))}\nD. {escape_md(opts.get('D', ''))}\nE. {escape_md(opts.get('E', ''))}"

def format_review(entry: dict, idx: int, total: int, user_data: dict) -> str:
q = entry["question"]
chosen = entry["chosen"]
correct = q["correct"]
opts = q["options"]
return f"📋 Review {idx}/{total}\n\n{escape_md(q['question'])}\n\nA. {escape_md(opts['A'])}\nB. {escape_md(opts['B'])}\nC. {escape_md(opts['C'])}\nD. {escape_md(opts['D'])}\nE. {escape_md(opts['E'])}\n\nYour answer: {chosen}\nCorrect answer: {correct}. {escape_md(opts[correct])}\n\n{t(user_data, 'explanation')} {escape_md(q.get('explanation', ''))}"

=============================================================================

SENDERS

=============================================================================

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

=============================================================================

COMMANDS HANDLERS

=============================================================================

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

=============================================================================

CALLBACKS

=============================================================================

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

=============================================================================

FILE HANDLING

=============================================================================

def classify_document(mime: str, name: str) -> Optional[str]:
mime = (mime or "").lower()
name = (name or "").lower()
if mime == "application/pdf" or name.endswith(".pdf"): return "pdf"
if mime in ("image/jpeg","image/jpg","image/png","image/webp") or name.endswith((".jpg",".jpeg",".png",".webp")): return "image"
if "wordprocessingml" in mime or name.endswith(".docx"): return "docx"
if mime == "text/plain" or name.endswith(".txt"): return "txt"
return None

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
user = update.effective_user
if user and await cooldown_guard(update, context.user_data, user.id): return
if context.user_data.get("busy"):
await update.message.reply_text(t(context.user_data, "busy"))
return
context.user_data["busy"] = True
try:
message = update.message
mode = context.user_data.get("pending_mode", "quiz")
file_name = "photo.jpg"
mime_type = "image/jpeg"
raw_bytes = b""
file_type = ""
if message.photo:
tg_file = await message.photo[-1].get_file()
raw_bytes = bytes(await tg_file.download_as_bytearray())
file_type = "image"
file_name = "telegram_photo.jpg"
mime_type = "image/jpeg"
elif message.document:
doc = message.document
file_name = doc.file_name or "file"
mime_type = doc.mime_type or "application/octet-stream"
if doc.file_size and doc.file_size > MAX_FILE_SIZE_BYTES:
await message.reply_text(t(context.user_data, "file_too_large"))
return
file_type = classify_document(mime_type, file_name) or ""
if not file_type:
await message.reply_text(t(context.user_data, "file_unsupported"))
return
tg_file = await doc.get_file()
raw_bytes = bytes(await tg_file.download_as_bytearray())
else:
await message.reply_text(t(context.user_data, "file_unsupported"))
return
if len(raw_bytes) > MAX_FILE_SIZE_BYTES:
await message.reply_text(t(context.user_data, "file_too_large"))
return
status = await message.reply_text(t(context.user_data, "file_received"))
file_hash = sha256_bytes(raw_bytes)
cached_file = await db_get_file(file_hash)
if cached_file:
text = cached_file["extracted_text"]
await status.edit_text(t(context.user_data, "cached_file"))
await asyncio.sleep(0.4)
else:
text = await extract_text(file_type, raw_bytes, mime_type)
if not text.strip() or len(text.strip()) < 30:
await status.edit_text(t(context.user_data, "file_no_text"))
return
await db_save_file(file_hash, file_name, len(raw_bytes), mime_type, text)
if mode == "summary":
style = context.user_data.get("pending_summary_style", "disease")
await status.edit_text(t(context.user_data, "generating_summary"))
summary = await get_or_create_summary(file_hash, text, style)
await status.delete()
await message.reply_text(t(context.user_data, "summary_header") + summary, parse_mode="Markdown")
return
setup_id = str(uuid.uuid4())[:8]
context.user_data["pending_mode"] = "quiz"
context.user_data["pending_file"] = {"setup_id": setup_id, "file_hash": file_hash, "file_name": file_name, "text": text}
await status.edit_text(t(context.user_data, "file_ready_quiz"), reply_markup=level_keyboard(setup_id, context.user_data))
except Exception as e:
logger.exception("file handling failed")
await update.message.reply_text(t(context.user_data, "error", err=str(e)))
finally:
context.user_data["busy"] = False

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
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

=============================================================================

MAIN

=============================================================================

BOT_COMMANDS = [
BotCommand("start", "Start bot"),
BotCommand("help", "Help"),
BotCommand("stop", "Stop active quiz"),
BotCommand("stats", "Show stats"),
BotCommand("review", "Review mistakes"),
BotCommand("language", "Change language"),
BotCommand("testbot", "Admin test"),
]

async def post_init(application: Application):
init_db()
await application.bot.set_my_commands(BOT_COMMANDS)
logger.info("Bot ready. Model=%s DB=%s MaxFile=%sMB", GEMINI_MODEL_NAME, DATABASE_PATH, MAX_FILE_SIZE_BYTES // 1024 // 1024)

def main():
persistence = PicklePersistence(filepath="bot_user_data.pkl")
application = (
ApplicationBuilder()
.token(TOKEN)
.persistence(persistence)
.concurrent_updates(True)
.post_init(post_init)
.build()
)
# Add handlers
application.add_error_handler(error_handler)
application.add_handler(CommandHandler("start", start))
application.add_handler(CommandHandler("help", help_command))
application.add_handler(CommandHandler("stop", stop_command))
application.add_handler(CommandHandler("stats", stats_command))
application.add_handler(CommandHandler("review", review_command))
application.add_handler(CommandHandler("language", language_command))
application.add_handler(CommandHandler("testbot", testbot_command))
application.add_handler(CallbackQueryHandler(lang_callback, pattern=r"^lang:"))
application.add_handler(CallbackQueryHandler(mode_callback, pattern=r"^mode:"))
application.add_handler(CallbackQueryHandler(summary_style_callback, pattern=r"^style:"))
application.add_handler(CallbackQueryHandler(level_callback, pattern=r"^level:"))
application.add_handler(CallbackQueryHandler(count_callback, pattern=r"^count:"))
application.add_handler(CallbackQueryHandler(answer_callback, pattern=r"^ans:"))
application.add_handler(CallbackQueryHandler(review_callback, pattern=r"^review:"))
# File filters
class SupportedFileFilter(filters.MessageFilter):
def filter(self, message):
if message.photo: return True
doc = message.document
if not doc: return False
return classify_document(doc.mime_type or "", doc.file_name or "") is not None
application.add_handler(MessageHandler(filters.PHOTO, handle_file))
application.add_handler(MessageHandler(SupportedFileFilter(), handle_file))
application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
logger.info("Starting bot polling...")
application.run_polling(allowed_updates=Update.ALL_TYPES)

if name == "main":
main() 
