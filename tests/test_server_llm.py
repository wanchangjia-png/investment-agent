import json

import ai_usage
import server


class FakeResponse:
    def __init__(self, payload=None, chunks=None):
        self.payload = payload
        self.chunks = list(chunks or [])

    def read(self, size=-1):
        if self.chunks:
            return self.chunks.pop(0)
        return self.payload or b""

    def close(self):
        pass


def _use_temp_log(monkeypatch, tmp_path):
    monkeypatch.setattr(ai_usage, "LOG_DIR", tmp_path)
    monkeypatch.setattr(ai_usage, "LOG_PATH", tmp_path / "ai_usage.jsonl")


def _records(tmp_path):
    return [json.loads(line) for line in (tmp_path / "ai_usage.jsonl").read_text().splitlines()]


def test_call_llm_text_logs_usage(monkeypatch, tmp_path):
    _use_temp_log(monkeypatch, tmp_path)
    monkeypatch.setitem(server.LLM_CONFIG, "api_key", "test-key")
    payload = {
        "choices": [{"message": {"content": "  你好  "}}],
        "usage": {"prompt_tokens": 100, "completion_tokens": 23, "total_tokens": 123},
    }
    monkeypatch.setattr(server.urllib.request, "urlopen",
                        lambda *a, **k: FakeResponse(json.dumps(payload).encode()))

    assert server._call_llm_text("system", "user") == "你好"
    record = _records(tmp_path)[-1]
    assert record["kind"] == "briefing"
    assert record["total_tokens"] == 123
    assert record["error"] is None


def test_call_llm_text_logs_error(monkeypatch, tmp_path):
    _use_temp_log(monkeypatch, tmp_path)
    monkeypatch.setitem(server.LLM_CONFIG, "api_key", "test-key")

    def boom(*a, **k):
        raise RuntimeError("network down")

    monkeypatch.setattr(server.urllib.request, "urlopen", boom)
    assert server._call_llm_text("system", "user") is None
    record = _records(tmp_path)[-1]
    assert "network down" in record["error"]


def test_chat_stream_logs_usage(monkeypatch, tmp_path):
    _use_temp_log(monkeypatch, tmp_path)
    monkeypatch.setitem(server.LLM_CONFIG, "api_key", "test-key")
    chunks = [
        b'data: {"choices":[{"delta":{"content":"hi"}}]}\n\n',
        b'data: {"choices":[],"usage":{"prompt_tokens":10,"completion_tokens":2,"total_tokens":12}}\n\n',
        b"data: [DONE]\n\n",
    ]
    monkeypatch.setattr(server.urllib.request, "urlopen",
                        lambda *a, **k: FakeResponse(chunks=chunks))

    out = list(server.chat_stream([{"role": "user", "content": "hello"}], search_enabled=False))
    assert any('"hi"' in item for item in out)
    record = _records(tmp_path)[-1]
    assert record["kind"] == "chat"
    assert record["total_tokens"] == 12
    assert record["error"] is None
