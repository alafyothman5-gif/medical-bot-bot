#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Payment Safety Mode self-test for MedMCQ.
Runs on a temporary SQLite database and does not touch production data.
"""

import asyncio
import os
import sqlite3
import tempfile
import time
from pathlib import Path

TMP_DIR = tempfile.mkdtemp(prefix="medmcq_payment_safety_")
os.environ["DATABASE_PATH"] = str(Path(TMP_DIR) / "payment_safety.sqlite3")
os.environ.setdefault("SMS_DUPLICATE_TTL_DAYS", "14")
os.environ.setdefault("PAYMENT_UNCLAIMED_MATCH_SECONDS", "86400")

import bot  # noqa: E402


def sms_text(phone: str, amount: float, tag: str) -> str:
    return f"Almadar: تم استلام تحويل بقيمة {int(amount)} د.ل من الرقم {phone}. ref {tag}. رصيدك الحالي 100 د.ل"


async def record_sms(phone: str, amount: float, tag: str, raw: str | None = None):
    raw_text = raw or sms_text(phone, amount, tag)
    parsed = bot.extract_madar_transfer_sms(raw_text)
    return await bot.db_record_incoming_madar_sms(
        sender="Almadar",
        sms_text=raw_text,
        phone=parsed.get("phone") or phone,
        amount=parsed.get("amount") if parsed.get("amount") is not None else amount,
        status="unclaimed",
        note="payment_safety_test",
    )


async def create_order(user_id: int, phone: str, amount: float, plan: str = "premium", days: int = 30):
    return await bot.db_create_payment_order(
        user_id=user_id,
        plan=plan,
        days=days,
        plan_title=f"TEST {plan}",
        amount=float(amount),
        payer_phone=phone,
    )


async def user_plan(user_id: int) -> str:
    return await bot.db_get_user_plan(user_id)


def fetch_sql(sql: str, params=()):
    conn = sqlite3.connect(os.environ["DATABASE_PATH"])
    conn.row_factory = sqlite3.Row
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


async def main() -> int:
    checks: list[tuple[str, bool, str]] = []

    def check(name: str, cond: bool, detail: str = ""):
        checks.append((name, bool(cond), detail))

    bot.init_db()

    # 1) Order before SMS: exact match must auto-approve atomically.
    order1 = await create_order(101, "0911111111", 20)
    led1 = await record_sms("0911111111", 20, "order_before_sms")
    res1 = await bot.db_atomic_approve_matching_sms("0911111111", 20, led1["sms_id"], note="test order before sms")
    check("order_before_sms_auto_approved", res1.get("ok") and await user_plan(101) == "premium", str(res1))

    # 2) SMS before order: unclaimed SMS must be claimed when the order is created.
    led2 = await record_sms("0912222222", 45, "sms_before_order")
    order2 = await create_order(102, "0912222222", 45, plan="premium", days=120)
    auto2 = await bot.try_auto_approve_order_from_stored_sms(order2.get("order"))
    check("sms_before_order_auto_linked", auto2 and await user_plan(102) == "premium", str(led2))

    # 3) Duplicate SMS must not activate a new user/order.
    dup_raw = sms_text("0911111111", 20, "order_before_sms")
    order3 = await create_order(103, "0911111111", 20)
    led3 = await record_sms("0911111111", 20, "duplicate_sms", raw=dup_raw)
    if not led3.get("duplicate"):
        await bot.db_atomic_approve_matching_sms("0911111111", 20, led3["sms_id"], note="should not happen")
    check("duplicate_sms_blocked", led3.get("duplicate") and await user_plan(103) == "free", str(led3))

    # 4) Wrong phone must not activate the pending order.
    await create_order(104, "0914444444", 20)
    led4 = await record_sms("0919999999", 20, "wrong_phone")
    res4 = await bot.db_atomic_approve_matching_sms("0919999999", 20, led4["sms_id"], note="wrong phone")
    check("wrong_phone_not_approved", (not res4.get("ok")) and await user_plan(104) == "free", str(res4))

    # 5) Wrong amount must not activate the pending order.
    await create_order(105, "0915555555", 45)
    led5 = await record_sms("0915555555", 20, "wrong_amount")
    res5 = await bot.db_atomic_approve_matching_sms("0915555555", 20, led5["sms_id"], note="wrong amount")
    check("wrong_amount_not_approved", (not res5.get("ok")) and await user_plan(105) == "free", str(res5))

    # 6) Multiple pending order protection for same user.
    first = await create_order(106, "0916666666", 20)
    second = await create_order(106, "0916666667", 45)
    check("one_pending_order_per_user", first.get("ok") and second.get("open_pending"), str(second))

    # 7) Bulk 50 orders + 50 SMS exact matches.
    approved = 0
    for i in range(50):
        uid = 1000 + i
        phone = f"09177{i:05d}"[:10]
        amount = [20.0, 45.0, 65.0][i % 3]
        plan = "premium_plus" if amount == 65.0 else "premium"
        days = 120 if amount in {45.0, 65.0} else 30
        made = await create_order(uid, phone, amount, plan=plan, days=days)
        if not made.get("ok"):
            continue
        led = await record_sms(phone, amount, f"bulk_{i}")
        res = await bot.db_atomic_approve_matching_sms(phone, amount, led["sms_id"], note="bulk safety test")
        if res.get("ok"):
            approved += 1
    check("bulk_50_exact_matches", approved == 50, f"approved={approved}")

    duplicate_claims = fetch_sql(
        "SELECT claimed_order_id, COUNT(*) AS c FROM unclaimed_madar_sms WHERE claimed_order_id > 0 GROUP BY claimed_order_id HAVING c > 1"
    )
    check("no_order_claimed_by_two_sms", len(duplicate_claims) == 0, str(duplicate_claims))

    duplicate_users = fetch_sql(
        "SELECT claimed_user_id, COUNT(*) AS c FROM unclaimed_madar_sms WHERE claimed_user_id > 0 GROUP BY claimed_user_id HAVING c > 1"
    )
    # A user should not have multiple claimed SMS in this test.
    check("no_user_double_claimed_in_test", len(duplicate_users) == 0, str(duplicate_users))

    bad = [x for x in checks if not x[1]]
    print("Payment Safety Mode test report")
    print("Database:", os.environ["DATABASE_PATH"])
    print("Total checks:", len(checks))
    for name, ok, detail in checks:
        print(("PASS" if ok else "FAIL"), "-", name, detail if not ok else "")
    print("RESULT:", "PASS" if not bad else "FAIL")
    return 0 if not bad else 1


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))
