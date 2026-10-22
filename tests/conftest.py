import os
import shutil
import sys
import tempfile
import uuid
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

# 必须在导入 settings / portfolio_agent 之前改路径，测试只动临时副本
_TMP = Path(tempfile.mkdtemp(prefix="portfolio_test_"))
_BASE = _TMP / "portfolio.xlsx"
shutil.copy2(ROOT / "portfolio.xlsx", _BASE)
os.environ["PORTFOLIO_XLSX"] = str(_BASE)
os.environ["PORTFOLIO_BACKUP_DIR"] = str(_TMP / "backups")


@pytest.fixture
def excel(monkeypatch):
    """每个用例一份独立副本，避免用例之间互相影响。"""
    import portfolio_agent as agent
    import settings

    path = _TMP / f"case_{uuid.uuid4().hex}.xlsx"
    shutil.copy2(_BASE, path)
    monkeypatch.setattr(settings, "EXCEL_PATH", path)
    monkeypatch.setattr(agent, "EXCEL_PATH", path)
    monkeypatch.setattr(settings, "_fee_cfg", None, raising=False)
    monkeypatch.setattr(settings, "_fee_cfg_mtime", 0.0, raising=False)
    yield path
    if path.exists():
        path.unlink()
