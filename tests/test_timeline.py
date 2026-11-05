import datetime as dt
from collections import defaultdict

import pytest

import timeline as tl


def _weekdays(start, end):
    out, d = [], dt.date.fromisoformat(start)
    stop = dt.date.fromisoformat(end)
    while d <= stop:
        if d.weekday() < 5:
            out.append(d.isoformat())
        d += dt.timedelta(days=1)
    return out


CAL = _weekdays("2025-12-01", "2026-03-10")


def _series(prices, days=None):
    """prices: {symbol: {date: close}} / {symbol: float}"""
    out = {}
    for symbol, spec in prices.items():
        if isinstance(spec, dict):
            out[symbol] = [[d, c] for d, c in sorted(spec.items())]
        else:
            out[symbol] = [[d, spec] for d in (days or CAL)]
    return out


def _txn(row, date, direction, code, name, qty, price, amount, fee=0.0, post=None):
    return {
        "row": row, "date": date, "direction": direction, "category": "股票",
        "name": name, "code": code, "qty": qty, "price": price,
        "amount": amount, "fee": fee, "ledger_qty": post,
    }


@pytest.fixture
def patched(monkeypatch):
    """把时间线的外部依赖换成可控数据。"""
    state = {"txns": [], "actions": [], "history": [], "flows": [], "holdings": []}

    monkeypatch.setattr(tl, "TIMELINE_START", "2026-01")
    monkeypatch.setattr(tl, "_today", lambda: "2026-03-10")
    monkeypatch.setattr(tl.agent, "load_transactions", lambda: state["txns"])
    monkeypatch.setattr(tl.agent, "load_corporate_actions", lambda: state["actions"])
    monkeypatch.setattr(tl.agent, "load_history", lambda: state["history"])
    monkeypatch.setattr(tl.agent, "load_capital_flows", lambda: state["flows"])
    monkeypatch.setattr(tl.agent, "load_portfolio", lambda: state["holdings"])
    monkeypatch.setattr(tl, "SECTOR_ETFS", [
        ("etf-wine", "酒ETF", "白酒/食品饮料"),
        ("etf-ev", "新能源车ETF", "新能源车/锂电"),
        ("etf-broad", "沪深300ETF", "宽基/海外"),
    ])
    return state


def test_collect_txns_opening_from_first_balance():
    txns = [
        _txn(2, "2026-01-05", "卖出", "000858", "五粮液", 100, 10, 1000, 5, post=200),
        _txn(3, "2026-02-05", "买入", "000858", "五粮液", 100, 12, 1200, 5, post=300),
    ]
    events, opening, meta = tl._collect_txns(txns)
    assert opening["000858"] == 300  # 卖后剩 200，卖前就是 300
    assert [e[1] for e in events["000858"]] == [-100, 100]
    assert meta["000858"]["name"] == "五粮液"


def test_month_pnl_includes_fees_and_dividends(patched):
    patched["txns"] = [
        _txn(2, "2026-01-05", "买入", "000858", "五粮液", 100, 10, 1000, 5, post=100),
        _txn(3, "2026-01-20", "卖出", "000858", "五粮液", 100, 12, 1200, 6, post=0),
    ]
    patched["actions"] = [{"type": "现金分红", "code": "000858", "date": "2026-01-15", "amount": 20}]
    series = _series({"sh000300": 100, "sz000858": 10})
    data = tl._build_all(series)

    jan = next(m for y in data["years"] for m in y["months"] if m["month"] == "2026-01")
    assert jan["pnl"] == pytest.approx(189 + 20)
    held_days = [d for d in CAL if "2026-01-05" <= d <= "2026-01-19"]
    month_days = [d for d in CAL if d.startswith("2026-01")]
    assert jan["avg_stock_value"] == pytest.approx(100 * 10 * len(held_days) / len(month_days))
    assert jan["return_pct"] == pytest.approx(jan["pnl"] / jan["avg_stock_value"] * 100, abs=0.01)
    assert jan["capital"]["fees"] == pytest.approx(11)
    assert jan["capital"]["dividends"] == pytest.approx(20)


def test_no_trade_month_uses_price_change(patched):
    patched["txns"] = [
        _txn(2, "2025-12-10", "买入", "002594", "比亚迪", 100, 20, 2000, 5, post=100),
    ]
    series = _series({"sh000300": 100, "sz002594": {**{d: 21.0 for d in CAL if d < "2026-02-01"},
                                                     **{d: 22.0 for d in CAL if d >= "2026-02-01"}}})
    data = tl._build_all(series)

    feb = next(m for y in data["years"] for m in y["months"] if m["month"] == "2026-02")
    assert feb["pnl"] == pytest.approx(100 * 22 - 100 * 21)
    assert feb["turnover"] == 0
    assert feb["main_sector"]["sector"] == "新能源车/锂电"
    assert feb["main_sector"]["turnover"] == 0
    assert feb["return_pct"] == pytest.approx(100 / 2200 * 100, abs=0.01)


def test_sector_rank_contributors_and_opportunity(patched):
    patched["txns"] = [
        _txn(2, "2026-01-05", "买入", "000858", "五粮液", 100, 10, 1000, 5, post=100),
        _txn(3, "2026-01-08", "买入", "002594", "比亚迪", 100, 20, 2000, 5, post=100),
        _txn(4, "2026-01-20", "卖出", "002594", "比亚迪", 100, 18, 1800, 6, post=0),
    ]
    patched["holdings"] = [{"类别": "股票", "账户": "A股主账户", "代码": "000858", "数量": 100}]
    series = _series({
        "sh000300": 100,
        "sz000858": {**{d: 10.0 for d in CAL if d < "2026-01-10"},
                     **{d: 11.0 for d in CAL if "2026-01-10" <= d < "2026-02-01"},
                     **{d: 12.0 for d in CAL if d >= "2026-02-01"}},
        "sz002594": 20,
        "etf-wine": {"2025-12-31": 1.0, "2026-01-30": 1.10},
        "etf-ev": {"2025-12-31": 1.0, "2026-01-30": 1.05},
        "etf-broad": {"2025-12-31": 1.0, "2026-01-30": 1.50},
    })
    data = tl._build_all(series)
    jan = next(m for y in data["years"] for m in y["months"] if m["month"] == "2026-01")

    assert jan["main_sector"]["sector"] == "新能源车/锂电"  # 成交额 2000+1800 > 1000
    assert [c["name"] for c in jan["drags"]] == ["比亚迪"]
    assert jan["contributors"][0]["name"] == "五粮液"
    opp = jan["opportunity"]
    assert opp["best_sector"] == "白酒/食品饮料"  # 宽基不参与评选
    assert opp["best_return_pct"] == pytest.approx(10.0)
    assert opp["coverage"] == 2
    assert opp["gap_pct"] == pytest.approx(10.0 - jan["return_pct"], abs=0.01)
    assert not any("宽基" in d["sector"] for d in data["warnings"])


def test_late_listed_etf_excluded_from_month(patched):
    patched["txns"] = []
    series = _series({
        "sh000300": 100,
        "etf-wine": {"2025-12-31": 1.0, "2026-01-30": 1.1},
        "etf-ev": {"2026-02-10": 1.0, "2026-02-27": 1.2},
        "etf-broad": {"2025-12-31": 1.0, "2026-01-30": 1.5},
    })
    data = tl._build_all(series)
    jan = next(m for y in data["years"] for m in y["months"] if m["month"] == "2026-01")
    assert jan["opportunity"]["coverage"] == 1
    assert jan["opportunity"]["best_sector"] == "白酒/食品饮料"


def test_year_best_etf_2024_uses_november_baseline(patched):
    series = _series({
        "etf-wine": {"2023-12-31": 0.5, "2024-10-31": 1.0, "2024-12-31": 1.1},
    })
    best = tl._year_best_etf(series, "2024", "2026-03-10")
    assert best["return_pct"] == pytest.approx(10.0)


def test_nav_returns_exclude_capital_flows(patched):
    patched["history"] = [
        {"日期": "2025-12-31", "总资产": 100000},
        {"日期": "2026-01-30", "总资产": 110000},
        {"日期": "2026-02-27", "总资产": 104000},
    ]
    patched["flows"] = [
        {"date": "2026-01-10 10:00", "type": "存入", "amount": 5000},
        {"date": "2026-01-20 10:00", "type": "取出", "amount": 2000},
        {"date": "2026-02-05 10:00", "type": "取出", "amount": 4000},
    ]
    months = tl._nav_returns()
    assert months["2026-01"] == pytest.approx((110000 - 100000 - 3000) / 100000 * 100, abs=0.01)
    assert months["2026-02"] == pytest.approx((104000 - 110000 + 4000) / 110000 * 100, abs=0.01)

    years = tl._nav_year_returns()
    assert years["2026"]["start"] == "2025-12-31"
    assert years["2026"]["full_year"] is True
    assert years["2026"]["return_pct"] == pytest.approx((104000 - 100000 + 1000) / 100000 * 100, abs=0.01)


def test_build_timeline_result_cache(patched, monkeypatch, excel):
    calls = {"build": 0,
             "series": 0}

    def fake_series(force=False):
        calls["series"] += 1
        return _series({"sh000300": 100})

    def fake_build(series):
        calls["build"] += 1
        return {"years": [], "warnings": [], "generated_at": "x",
                "range": {"start": "2026-01", "end": "2026-01", "partial": False}, "meta": {}}

    monkeypatch.setattr(tl, "load_series", fake_series)
    monkeypatch.setattr(tl, "_build_all", fake_build)

    first = tl.build_timeline()
    second = tl.build_timeline()
    assert calls["build"] == 1
    assert first.get("cache", {}).get("hit") is False
    assert second.get("cache", {}).get("hit") is True

    tl.build_timeline(force=True)
    assert calls["build"] == 2


def test_missing_price_is_reported_not_zeroed(patched):
    patched["txns"] = [
        _txn(2, "2026-01-05", "买入", "000858", "五粮液", 100, 10, 1000, 5, post=100),
    ]
    series = _series({"sh000300": 100, "sz000001": 10})  # 000858 完全没有行情
    data = tl._build_all(series)
    jan = next(m for y in data["years"] for m in y["months"] if m["month"] == "2026-01")
    assert jan["pnl"] == 0  # 不按 0 元估值
    assert any("缺当月行情" in w for w in data["warnings"])
