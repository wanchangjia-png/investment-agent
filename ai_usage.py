#!/usr/bin/env python3
"""AI 调用日志：每次 LLM 请求记一条 JSONL，用于查看消耗、延迟和失败率。"""
import json
import os
import threading
from datetime import datetime, timedelta

from settings import DATA_DIR, now

LOG_DIR = DATA_DIR / "logs"
LOG_PATH = LOG_DIR / "ai_usage.jsonl"
MAX_BYTES = 5 * 1024 * 1024
_LOCK = threading.Lock()

# 粗略估算用（美元 / 百万 token），价格变动时改这里即可
PRICES = {
    "deepseek-chat": {"input": 0.27, "cache_hit": 0.07, "output": 1.10},
    "deepseek-reasoner": {"input": 0.55, "cache_hit": 0.14, "output": 2.19},
    "gpt-4o-mini": {"input": 0.15, "cache_hit": 0.15, "output": 0.60},
}
DEFAULT_PRICE = {"input": 0.0, "cache_hit": 0.0, "output": 0.0}


def log_call(kind, provider, model, started_at, usage=None, error=None):
    """记录一次调用；usage 用 OpenAI 兼容格式，缺失时按 0 处理。"""
    usage = usage or {}
    try:
        finished = now()
        started = datetime.fromtimestamp(started_at, finished.tzinfo) if started_at else finished
        latency_ms = int((finished - started).total_seconds() * 1000)
        prompt_tokens = int(usage.get("prompt_tokens") or 0)
        completion_tokens = int(usage.get("completion_tokens") or 0)
        total_tokens = int(usage.get("total_tokens") or prompt_tokens + completion_tokens)
        cache_hit = int(usage.get("prompt_cache_hit_tokens") or usage.get("cache_hit_tokens") or 0)
        record = {
            "ts": finished.isoformat(timespec="seconds"),
            "kind": kind,
            "provider": provider,
            "model": model,
            "latency_ms": latency_ms,
            "prompt_tokens": prompt_tokens,
            "completion_tokens": completion_tokens,
            "total_tokens": total_tokens,
            "cache_hit_tokens": cache_hit,
            "error": (str(error)[:300] if error else None),
        }
        line = json.dumps(record, ensure_ascii=False) + "\n"
    except Exception:
        return None

    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        with _LOCK:
            if LOG_PATH.exists() and LOG_PATH.stat().st_size > MAX_BYTES:
                os.replace(LOG_PATH, LOG_PATH.with_suffix(".jsonl.1"))
            with LOG_PATH.open("a", encoding="utf-8") as f:
                f.write(line)
        return record
    except OSError:
        return None


def _cost(record):
    price = PRICES.get(record.get("model", ""), DEFAULT_PRICE)
    hit = record.get("cache_hit_tokens", 0)
    prompt = max(record.get("prompt_tokens", 0) - hit, 0)
    return (prompt * price["input"]
            + hit * price["cache_hit"]
            + record.get("completion_tokens", 0) * price["output"]) / 1_000_000


def _empty_bucket():
    return {"calls": 0, "errors": 0, "prompt_tokens": 0, "completion_tokens": 0,
            "total_tokens": 0, "cache_hit_tokens": 0, "latency_ms": 0,
            "cost_usd": 0.0, "last_at": None}


def _add(bucket, record):
    bucket["calls"] += 1
    if record.get("error"):
        bucket["errors"] += 1
    for key in ("prompt_tokens", "completion_tokens", "total_tokens",
                "cache_hit_tokens", "latency_ms"):
        bucket[key] += record.get(key, 0)
    bucket["cost_usd"] += _cost(record)
    bucket["last_at"] = record.get("ts")


def _finish(bucket):
    calls = bucket["calls"] or 1
    bucket["avg_latency_ms"] = round(bucket["latency_ms"] / calls)
    bucket["cost_usd"] = round(bucket["cost_usd"], 6)
    bucket.pop("latency_ms", None)
    return bucket


def summarize(days=30):
    """按天回溯，汇总总消耗、各用途/各模型的调用次数和延迟。"""
    days = max(1, min(int(days or 30), 365))
    cutoff = now() - timedelta(days=days)
    totals = _empty_bucket()
    by_kind, by_model = {}, {}
    recent = []

    try:
        with LOG_PATH.open("r", encoding="utf-8") as f:
            for line in f:
                try:
                    record = json.loads(line)
                    ts = datetime.fromisoformat(record.get("ts", ""))
                except (ValueError, TypeError, json.JSONDecodeError):
                    continue
                if ts < cutoff:
                    continue
                _add(totals, record)
                _add(by_kind.setdefault(record.get("kind", "other"), _empty_bucket()), record)
                _add(by_model.setdefault(record.get("model", "unknown"), _empty_bucket()), record)
                recent.append(record)
    except FileNotFoundError:
        pass
    except OSError:
        pass

    for bucket in [totals] + list(by_kind.values()) + list(by_model.values()):
        _finish(bucket)
    totals["days"] = days
    recent.sort(key=lambda r: r.get("ts", ""), reverse=True)
    return {"totals": totals,
            "by_kind": by_kind,
            "by_model": by_model,
            "recent": recent[:30]}
