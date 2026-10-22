import time

import ai_usage


def _use_temp_log(monkeypatch, tmp_path):
    monkeypatch.setattr(ai_usage, "LOG_DIR", tmp_path)
    monkeypatch.setattr(ai_usage, "LOG_PATH", tmp_path / "ai_usage.jsonl")


def test_log_and_summarize(monkeypatch, tmp_path):
    _use_temp_log(monkeypatch, tmp_path)
    ai_usage.log_call("chat", "deepseek", "deepseek-chat", time.time() - 1.0,
                      {"prompt_tokens": 1000, "completion_tokens": 200,
                       "total_tokens": 1200, "prompt_cache_hit_tokens": 400})
    ai_usage.log_call("briefing", "deepseek", "deepseek-chat", time.time() - 0.5,
                      None, "HTTP 500: boom")

    summary = ai_usage.summarize(30)
    totals = summary["totals"]
    assert totals["calls"] == 2
    assert totals["errors"] == 1
    assert totals["total_tokens"] == 1200
    assert totals["cost_usd"] > 0
    assert set(summary["by_kind"]) == {"chat", "briefing"}
    assert len(summary["recent"]) == 2


def test_summarize_without_log(monkeypatch, tmp_path):
    _use_temp_log(monkeypatch, tmp_path)
    summary = ai_usage.summarize(7)
    assert summary["totals"]["calls"] == 0
    assert summary["recent"] == []
    assert summary["totals"]["days"] == 7
