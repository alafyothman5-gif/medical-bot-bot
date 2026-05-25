import json
import os
import time
import uuid
from typing import Any, Dict, Optional

import redis

REDIS_URL = os.environ.get("REDIS_URL", "redis://127.0.0.1:6379/0")
QUEUE_NAME = os.environ.get("QUEUE_NAME", "medical_bot:jobs")

r = redis.Redis.from_url(REDIS_URL, decode_responses=True)


def create_job(job_type: str, chat_id: int, user_id: int, payload: Dict[str, Any]) -> str:
    job_id = str(uuid.uuid4())
    job = {
        "job_id": job_id,
        "job_type": job_type,
        "chat_id": int(chat_id),
        "user_id": int(user_id or 0),
        "payload": payload or {},
        "status": "queued",
        "created_at": time.time(),
    }
    r.set(f"medical_bot:job:{job_id}", json.dumps(job, ensure_ascii=False), ex=60 * 60 * 24)
    r.rpush(QUEUE_NAME, job_id)
    return job_id


def get_job(job_id: str) -> Optional[Dict[str, Any]]:
    raw = r.get(f"medical_bot:job:{job_id}")
    return json.loads(raw) if raw else None


def set_job_status(job_id: str, status: str, result: Optional[Dict[str, Any]] = None, error: str = "") -> None:
    job = get_job(job_id)
    if not job:
        return
    job["status"] = status
    job["updated_at"] = time.time()
    if result is not None:
        job["result"] = result
    if error:
        job["error"] = str(error)
    r.set(f"medical_bot:job:{job_id}", json.dumps(job, ensure_ascii=False), ex=60 * 60 * 24)


def pop_job(timeout: int = 5) -> Optional[str]:
    item = r.blpop(QUEUE_NAME, timeout=timeout)
    return item[1] if item else None


def queue_size() -> int:
    return int(r.llen(QUEUE_NAME))


def save_pending_quiz(user_id: int, setup_id: str, pending: Dict[str, Any]) -> None:
    if not user_id or not setup_id or not isinstance(pending, dict):
        return
    r.set(
        f"medical_bot:pending_quiz:{int(user_id)}:{setup_id}",
        json.dumps(pending, ensure_ascii=False),
        ex=60 * 60 * 24,
    )


def load_pending_quiz(user_id: int, setup_id: str) -> Optional[Dict[str, Any]]:
    raw = r.get(f"medical_bot:pending_quiz:{int(user_id)}:{setup_id}")
    if not raw:
        return None
    obj = json.loads(raw)
    return obj if isinstance(obj, dict) else None
