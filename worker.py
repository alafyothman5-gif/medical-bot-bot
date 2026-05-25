#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import asyncio
import logging
import os

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

from telegram import Bot, InlineKeyboardButton, InlineKeyboardMarkup

from queue_jobs import pop_job, get_job, set_job_status, queue_size

logging.basicConfig(
    format="%(asctime)s - worker - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger("medical_bot_worker")

TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
if not TOKEN:
    raise RuntimeError("TELEGRAM_BOT_TOKEN is missing from environment")


async def _send_long(bot: Bot, chat_id: int, text: str, prefix: str = "") -> int:
    text = (prefix or "") + (text or "")
    chunks = [text[i:i + 3800] for i in range(0, len(text), 3800)] or ["⚠️ لا يوجد محتوى لإرساله."]
    for chunk in chunks:
        await bot.send_message(chat_id=chat_id, text=chunk)
    return len(chunks)


async def process_job(bot: Bot, job: dict):
    job_id = job["job_id"]
    chat_id = int(job["chat_id"])
    user_id = int(job.get("user_id") or 0)
    job_type = job.get("job_type", "")
    payload = job.get("payload", {}) if isinstance(job.get("payload"), dict) else {}

    set_job_status(job_id, "processing")

    if job_type == "test":
        await asyncio.sleep(1)
        await bot.send_message(chat_id=chat_id, text=f"✅ Worker شغال.\nJob ID: {job_id}\nQueue left: {queue_size()}")
        set_job_status(job_id, "done", result={"ok": True})
        return

    if job_type == "summary":
        from bot import get_or_create_summary, db_add_xp

        file_hash = payload.get("file_hash", "")
        text = payload.get("text", "")
        style = payload.get("style", "disease_v2")

        await bot.send_message(chat_id=chat_id, text=f"⏳ بدأ تجهيز الملخص في الخلفية...\nJob ID: {job_id}")
        summary = await get_or_create_summary(file_hash, text, style)

        if not summary or not str(summary).strip():
            await bot.send_message(chat_id=chat_id, text="⚠️ فشل توليد الملخص. جرّب ملفاً أوضح.")
            set_job_status(job_id, "failed", error="empty summary")
            return

        chunks = await _send_long(bot, chat_id, summary, prefix="📚 الملخص\n\n")

        if user_id:
            try:
                p = await db_add_xp(user_id, 15, "Summary generated", summary_done=True)
                await bot.send_message(chat_id=chat_id, text=f"⭐ +15 XP | 🎖 Level {p.get('level', 1)} | 🔥 Streak {p.get('streak', 0)} يوم")
            except Exception:
                logger.exception("XP update failed")

        set_job_status(job_id, "done", result={"ok": True, "chunks": chunks})
        return

    if job_type == "question_bank":
        from bot import get_or_create_question_bank, normalize_questions

        file_hash = payload.get("file_hash", "")
        text = payload.get("text", "")
        level = payload.get("level", "basic")
        setup_id = payload.get("setup_id", "")

        await bot.send_message(
            chat_id=chat_id,
            text=f"⏳ بدأ تجهيز بنك الأسئلة في الخلفية...\n📚 المستوى: {level}\nJob ID: {job_id}",
        )

        questions = normalize_questions(await get_or_create_question_bank(file_hash, text, level))
        logger.info("question_bank done level=%s count=%s setup_id=%s", level, len(questions), setup_id)

        if not questions:
            await bot.send_message(chat_id=chat_id, text="⚠️ لم أستطع توليد أسئلة من هذا الملف. جرّب ملفاً أوضح.")
            set_job_status(job_id, "failed", error="no questions")
            return

        if setup_id:
            await bot.send_message(
                chat_id=chat_id,
                text=(
                    "✅ بنك الأسئلة جاهز.\n"
                    f"📚 المستوى: {level}\n"
                    f"📝 عدد الأسئلة المتاحة: {len(questions)}\n\n"
                    "اضغط الزر لبدء الاختبار مباشرة:"
                ),
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("▶️ ابدأ الاختبار الآن", callback_data=f"level:{setup_id}:{level}")]
                ]),
            )
        else:
            await bot.send_message(
                chat_id=chat_id,
                text=f"✅ بنك الأسئلة جاهز.\n📚 المستوى: {level}\n📝 عدد الأسئلة المتاحة: {len(questions)}\nارجع واضغط نفس المستوى مرة ثانية.",
            )

        set_job_status(job_id, "done", result={"ok": True, "questions": len(questions)})
        return

    await bot.send_message(chat_id=chat_id, text=f"⚠️ نوع مهمة غير معروف: {job_type}")
    set_job_status(job_id, "failed", error=f"unknown job_type: {job_type}")


async def main():
    bot = Bot(TOKEN)
    logger.info("Worker started")
    while True:
        job_id = pop_job(timeout=5)
        if not job_id:
            await asyncio.sleep(1)
            continue
        job = get_job(job_id)
        if not job:
            logger.warning("Job not found: %s", job_id)
            continue
        try:
            logger.info("Processing job %s type=%s", job_id, job.get("job_type"))
            await process_job(bot, job)
        except Exception as e:
            logger.exception("Job failed: %s", job_id)
            set_job_status(job_id, "failed", error=str(e))


if __name__ == "__main__":
    asyncio.run(main())
