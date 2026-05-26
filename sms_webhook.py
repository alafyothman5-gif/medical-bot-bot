#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Almadar SMS Webhook — Payment Safety Mode
Receives official Almadar SMS text forwarded from MacroDroid.
Safety rules:
- only configured sender(s), default Almadar
- every accepted Almadar SMS is saved in unclaimed_madar_sms ledger
- duplicate SMS is blocked for SMS_DUPLICATE_TTL_DAYS
- auto activation is atomic: order pending + SMS unclaimed + phone + amount + time window
- uncertain cases remain pending for admin review
"""

import json
import os
import sys
import time
from pathlib import Path
from typing import Optional

import aiohttp
from aiohttp import web

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

import bot

HOST = os.environ.get("SMS_WEBHOOK_HOST", "0.0.0.0")
PORT = int(os.environ.get("SMS_WEBHOOK_PORT", "8088"))
TOKEN = os.environ.get("SMS_WEBHOOK_TOKEN", "MadarPay2026").strip()
HEARTBEAT_FILE = os.environ.get("PAYMENT_PHONE_HEARTBEAT_FILE", "data/payment_phone_heartbeat.json")


def _auth_ok(request: web.Request) -> bool:
    token = request.query.get("token", "").strip()
    if not token:
        auth = request.headers.get("Authorization", "")
        if auth.lower().startswith("bearer "):
            token = auth.split(" ", 1)[1].strip()
    if not token:
        token = request.headers.get("X-SMS-Token", "").strip()
    return bool(token and token == TOKEN)


async def _send_telegram_message(chat_id: int, text: str, reply_markup: Optional[dict] = None) -> None:
    if not bot.TOKEN:
        return
    payload = {"chat_id": int(chat_id), "text": text}
    if reply_markup:
        payload["reply_markup"] = reply_markup
    url = f"https://api.telegram.org/bot{bot.TOKEN}/sendMessage"
    try:
        async with aiohttp.ClientSession() as session:
            async with session.post(url, json=payload, timeout=20) as resp:
                await resp.text()
    except Exception:
        pass


async def _send_admin(text: str, reply_markup: Optional[dict] = None) -> None:
    await _send_telegram_message(bot.ADMIN_LOG_CHAT_ID, text, reply_markup=reply_markup)


async def _read_text_from_request(request: web.Request) -> str:
    if request.method == "GET":
        return request.query.get("text", "") or request.query.get("sms", "") or request.query.get("message", "")
    ctype = (request.headers.get("Content-Type") or "").lower()
    if "application/json" in ctype:
        data = await request.json()
        if isinstance(data, dict):
            return str(data.get("text") or data.get("sms") or data.get("message") or data.get("body") or "")
        return str(data)
    if "application/x-www-form-urlencoded" in ctype or "multipart/form-data" in ctype:
        data = await request.post()
        return str(data.get("text") or data.get("sms") or data.get("message") or data.get("body") or "")
    raw = await request.text()
    return raw.strip()


def _normalize_sms_sender(value: str) -> str:
    return "".join(str(value or "").strip().lower().split())


def _sender_allowed(sender: str) -> bool:
    raw = os.environ.get("ALLOWED_SMS_SENDERS", "Almadar").strip()
    if not raw:
        return True
    allowed = [_normalize_sms_sender(x) for x in raw.split(",") if x.strip()]
    return _normalize_sms_sender(sender) in allowed


def _write_payment_heartbeat(device: str = "madar_phone", sender: str = "") -> dict:
    Path(HEARTBEAT_FILE).parent.mkdir(parents=True, exist_ok=True)
    data = {"ok": True, "device": device or "madar_phone", "sender": sender or "", "ts": time.time()}
    Path(HEARTBEAT_FILE).write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return data


async def heartbeat_handler(request: web.Request) -> web.Response:
    if not _auth_ok(request):
        return web.json_response({"ok": False, "error": "unauthorized"}, status=401)
    device = request.query.get("device", "madar_phone")
    sender = request.query.get("sender", "")
    data = _write_payment_heartbeat(device=device, sender=sender)
    return web.json_response(data)


async def sms_handler(request: web.Request) -> web.Response:
    if not _auth_ok(request):
        return web.json_response({"ok": False, "error": "unauthorized"}, status=401)

    sms_text = (await _read_text_from_request(request)).strip()
    if not sms_text:
        return web.json_response({"ok": False, "error": "empty_sms"}, status=400)

    sender = (
        request.query.get("sender")
        or request.query.get("from")
        or request.headers.get("X-SMS-Sender")
        or ""
    ).strip()

    if not _sender_allowed(sender):
        await _send_admin(
            "🚫 تم رفض SMS لأنها ليست من مرسل مسموح.\n\n"
            f"المرسل: {sender or 'غير معروف'}\n\n"
            f"النص:\n{sms_text[:1000]}"
        )
        return web.json_response({"ok": True, "rejected": True, "reason": "sender_not_allowed", "sender": sender})

    parsed = bot.extract_madar_transfer_sms(sms_text)
    phone = parsed.get("phone") or ""
    amount = parsed.get("amount")
    print(f"SMS_CHECK sender={sender!r} phone={phone!r} amount={amount!r}", flush=True)

    if not phone or amount is None:
        ledger = await bot.db_record_incoming_madar_sms(
            sender=sender or "Almadar",
            sms_text=sms_text,
            phone=phone,
            amount=amount,
            status="ignored",
            note="could not parse phone or amount",
        )
        await _send_admin(
            "⚠️ وصلت رسالة مدار لكن لم أستطع قراءة الرقم/المبلغ.\n\n"
            f"Ledger SMS ID: {ledger.get('sms_id')}\n"
            f"النص:\n{sms_text[:1000]}"
        )
        return web.json_response({"ok": True, "parsed": False, "needs_review": True, "sms_id": ledger.get("sms_id")})

    ledger = await bot.db_record_incoming_madar_sms(
        sender=sender or "Almadar",
        sms_text=sms_text,
        phone=phone,
        amount=float(amount),
        status="unclaimed",
        note="received from sms_webhook",
    )
    sms_id = int(ledger.get("sms_id") or 0)

    if ledger.get("duplicate"):
        return web.json_response({"ok": True, "duplicate": True, "sms_id": sms_id})

    result = await bot.db_atomic_approve_matching_sms(
        phone,
        float(amount),
        sms_id,
        note="atomic auto approved by Almadar SMS webhook",
    )

    if result.get("ok"):
        order = result["order"]
        await _send_telegram_message(
            int(order["user_id"]),
            "✅ تم تأكيد تحويل المدار تلقائيًا وتفعيل اشتراكك.\n\n"
            f"الخطة: {order['plan_title']}\n"
            f"المدة: {order['days']} يوم\n"
            f"المبلغ: {int(float(order['amount']))} د.ل"
        )
        await _send_admin(
            "✅ تفعيل تلقائي آمن من رسالة المدار\n\n"
            f"Ledger SMS ID: {sms_id}\n"
            f"Order: {order['order_code']}\n"
            f"User ID: {order['user_id']}\n"
            f"الخطة: {order['plan_title']}\n"
            f"المبلغ: {int(float(order['amount']))} د.ل\n"
            f"رقم المحوّل: {order['payer_phone']}"
        )
        return web.json_response({"ok": True, "auto_approved": True, "order_id": order["id"], "sms_id": sms_id})

    status = result.get("status")
    if status == "ambiguous":
        await _send_admin(
            "⚠️ رسالة مدار مطابقة لأكثر من طلب، تحتاج مراجعة.\n\n"
            f"Ledger SMS ID: {sms_id}\n"
            f"الرقم: {phone}\nالمبلغ: {amount}\n\nالنص:\n{sms_text[:1000]}"
        )
        return web.json_response({"ok": True, "ambiguous": True, "sms_id": sms_id})

    # No exact pending order: keep SMS as unclaimed for 24h+ so a later order can claim it.
    await _send_admin(
        "⚠️ رسالة مدار بدون طلب دفع مطابق.\n\n"
        f"Ledger SMS ID: {sms_id}\n"
        f"الرقم المقروء: {phone}\n"
        f"المبلغ المقروء: {amount} د.ل\n"
        f"الحالة: {status}\n\n"
        f"النص:\n{sms_text[:1000]}"
    )
    return web.json_response({"ok": True, "matched": False, "phone": phone, "amount": amount, "sms_id": sms_id, "status": status})


async def health(request: web.Request) -> web.Response:
    return web.json_response({"ok": True, "service": "almadar_sms_webhook", "mode": "payment_safety"})


def main() -> None:
    bot.init_db()
    if not TOKEN:
        print("WARNING: SMS_WEBHOOK_TOKEN is empty. Set it in .env before public use.", file=sys.stderr)
    app = web.Application()
    app.router.add_get("/", health)
    app.router.add_get("/heartbeat", heartbeat_handler)
    app.router.add_get("/sms", sms_handler)
    app.router.add_post("/sms", sms_handler)
    print(f"Starting SMS webhook on http://{HOST}:{PORT}/sms")
    web.run_app(app, host=HOST, port=PORT)


if __name__ == "__main__":
    main()
