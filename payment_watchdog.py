#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import asyncio
import json
import os
import time
from pathlib import Path

import aiohttp

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

import bot

HEARTBEAT_FILE = os.environ.get("PAYMENT_PHONE_HEARTBEAT_FILE", "data/payment_phone_heartbeat.json")
STATE_FILE = os.environ.get("PAYMENT_WATCHDOG_STATE_FILE", "data/payment_watchdog_state.json")
OFFLINE_AFTER_SECONDS = int(os.environ.get("PAYMENT_PHONE_OFFLINE_AFTER_SECONDS", "1800"))
CHECK_EVERY_SECONDS = int(os.environ.get("PAYMENT_WATCHDOG_INTERVAL_SECONDS", "300"))
PENDING_ALERT_AFTER_SECONDS = int(os.environ.get("PAYMENT_PENDING_ALERT_AFTER_SECONDS", "1800"))


def _load_json(path, default):
    try:
        if not os.path.exists(path):
            return default
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except Exception:
        return default


def _save_json(path, data):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


async def send_admin(text: str):
    if not bot.TOKEN or not bot.ADMIN_LOG_CHAT_ID:
        return
    url = f"https://api.telegram.org/bot{bot.TOKEN}/sendMessage"
    payload = {"chat_id": int(bot.ADMIN_LOG_CHAT_ID), "text": text}
    try:
        async with aiohttp.ClientSession() as session:
            async with session.post(url, json=payload, timeout=20) as resp:
                await resp.text()
    except Exception:
        pass


def human_age(seconds):
    seconds = int(max(0, seconds))
    if seconds < 60:
        return f"{seconds} ثانية"
    if seconds < 3600:
        return f"{seconds//60} دقيقة"
    if seconds < 86400:
        return f"{seconds//3600} ساعة"
    return f"{seconds//86400} يوم"


def get_pending_orders():
    now = time.time()
    conn = bot._db_connect()
    rows = conn.execute(
        "SELECT * FROM payment_orders WHERE status='pending' ORDER BY created_at DESC LIMIT 20"
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows if now - float(r["created_at"]) >= PENDING_ALERT_AFTER_SECONDS]


async def main():
    bot.init_db()
    Path("data").mkdir(exist_ok=True)
    print("Payment watchdog started", flush=True)

    while True:
        state = _load_json(STATE_FILE, {})
        now = time.time()

        hb = _load_json(HEARTBEAT_FILE, {})
        hb_ts = float(hb.get("ts") or 0)

        if not hb_ts:
            if now - float(state.get("last_no_heartbeat_alert", 0)) > 3600:
                await send_admin(
                    "⚠️ لا توجد أي إشارة من هاتف الدفع حتى الآن.\n\n"
                    "افتح MacroDroid وتأكد من تشغيل Heartbeat."
                )
                state["last_no_heartbeat_alert"] = now
        else:
            age = now - hb_ts
            if age > OFFLINE_AFTER_SECONDS:
                if now - float(state.get("last_offline_alert", 0)) > 3600:
                    await send_admin(
                        "⚠️ تنبيه: هاتف الدفع لا يرسل إشارات.\n\n"
                        f"آخر إشارة كانت منذ: {human_age(age)}\n"
                        "قد يكون الهاتف طافي، النت فاصل، أو MacroDroid توقف."
                    )
                    state["last_offline_alert"] = now

        pending = get_pending_orders()
        if pending and now - float(state.get("last_pending_alert", 0)) > 1800:
            lines = ["📋 توجد طلبات دفع معلقة تحتاج متابعة:\n"]
            for r in pending[:5]:
                lines.append(
                    f"- #{r['id']} | User {r['user_id']} | {r['plan_title']} | "
                    f"{int(float(r['amount']))} د.ل | {r['payer_phone']} | منذ {human_age(now - float(r['created_at']))}"
                )
            lines.append("\nاكتب /pending_payments لعرضها وتفعيلها بزر واحد.")
            await send_admin("\n".join(lines))
            state["last_pending_alert"] = now

        _save_json(STATE_FILE, state)
        await asyncio.sleep(CHECK_EVERY_SECONDS)


if __name__ == "__main__":
    asyncio.run(main())
