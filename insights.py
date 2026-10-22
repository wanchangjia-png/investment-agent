#!/usr/bin/env python3
"""持仓深度分析：基准对比、时间加权收益、回撤、归因、行业暴露、风险与提醒。"""
import json
import math
import threading
import urllib.request
from collections import defaultdict
from datetime import datetime

import openpyxl

import portfolio_agent as agent
from settings import BEIJING_TZ
BENCHMARK_CODE = "sh000300"
BENCHMARK_NAME = "沪深300"

# ============ 万万的交易纪律（2026-10-22 按她的习惯定制）============
DISCIPLINE = {
    "cash_floor_pct": 20.0,      # 现金长期不低于总资产 20%
    "single_cap_pct": 20.0,      # 单只股票不超过 20%
    "trade_size_pct": 5.0,       # 单笔买入不超过 5%（约 1 万）
    "take_profit_tiers": [20.0, 30.0],  # 成本 +20% / +30% 各卖 1/3
    "red_line_pct": -20.0,       # 浮亏 20% 触发减半红线
    "gold_tiers": [0.0, 10.0],   # 黄金：回本卖出 1/3，+10% 再卖 1/3
}

_index_cache = {"at": None, "data": None}
_index_lock = threading.Lock()


def _num(v):
    return float(v) if isinstance(v, (int, float)) else 0.0


# ============ 基准行情 ============
def fetch_index_history(days=240):
    """沪深300 日线收盘价 {日期: 收盘}，10 分钟内存缓存"""
    now = datetime.now()
    with _index_lock:
        if _index_cache["data"] and _index_cache["at"] and (now - _index_cache["at"]).total_seconds() < 600:
            return _index_cache["data"]
    url = (f"https://web.ifzq.gtimg.cn/appstock/app/fqkline/get"
           f"?param={BENCHMARK_CODE},day,,,{days},qfq")
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        raw = json.loads(urllib.request.urlopen(req, timeout=10).read().decode("utf-8"))
        node = raw["data"][BENCHMARK_CODE]
        rows = node.get("qfqday") or node.get("day") or []
        data = {r[0]: float(r[2]) for r in rows if len(r) >= 3}
    except Exception as e:
        print(f"⚠️ 沪深300行情获取失败: {e}")
        data = {}
    with _index_lock:
        _index_cache["at"] = now
        _index_cache["data"] = data
    return data


# ============ 收益与基准 ============
def _flows_by_snapshot(snapshot_dates):
    """把出入金归到第一个不早于它日期的净值快照（快照是收盘后补录的，通常已含当日出入金）"""
    flows = defaultdict(float)
    for f in agent.load_capital_flows():
        date = str(f["date"])[:10]
        amt = _num(f["amount"])
        signed = amt if f["type"] == "存入" else -amt
        target = next((sd for sd in snapshot_dates if sd >= date), None)
        if target:
            flows[target] += signed
    return flows


def performance():
    """基于净值快照的时间加权收益，对比沪深300"""
    history = agent.load_history()
    if len(history) < 2:
        return {"points": [], "metrics": {}, "error": "净值历史不足两天"}

    flows = _flows_by_snapshot([r["日期"] for r in history])
    bench = fetch_index_history()

    points = []
    idx = 1.0
    prev_value = _num(history[0]["总资产"])
    prev_snapshot_date = history[0]["日期"]
    bench_start = bench.get(history[0]["日期"])
    peak, max_dd = idx, 0.0
    daily_returns = []

    for rec in history:
        date = rec["日期"]
        value = _num(rec["总资产"])
        if date != history[0]["日期"] and prev_value > 0:
            r = (value - flows.get(date, 0.0)) / prev_value - 1
            if abs(r) < 0.5:  # 过滤明显错误快照
                idx *= (1 + r)
                prev_date = datetime.strptime(prev_snapshot_date, "%Y-%m-%d")
                gap = max((datetime.strptime(date, "%Y-%m-%d") - prev_date).days, 1)
                daily_returns.append(r / math.sqrt(gap))
            peak = max(peak, idx)
            if peak > 0:
                max_dd = min(max_dd, idx / peak - 1)
        prev_value = value
        prev_snapshot_date = date

        b_close = bench.get(date)
        b_idx = None
        if b_close and bench_start:
            bench_idx = b_close / bench_start
            points.append({
                "date": date,
                "account": round(idx * 100, 3),
                "benchmark": round(bench_idx * 100, 3),
                "total": value,
                "benchmark_nav": b_close,
            })
    # 若历史日期与交易日错位，补最后一个交易日的基准点
    bench_total_return = (points[-1]["benchmark"] / 100 - 1) if points else 0
    account_total_return = idx - 1
    if len(daily_returns) >= 2:
        mean = sum(daily_returns) / len(daily_returns)
        var = sum((x - mean) ** 2 for x in daily_returns) / (len(daily_returns) - 1)
        vol = math.sqrt(var) * math.sqrt(252)
        sharpe = (mean * 252 - 0.015) / (math.sqrt(var) * math.sqrt(252)) if var > 0 else 0
        win_rate = sum(1 for x in daily_returns if x > 0) / len(daily_returns)
    else:
        vol = sharpe = win_rate = 0

    start_date = history[0]["日期"]
    end_date = history[-1]["日期"]
    try:
        span_days = (datetime.strptime(end_date, "%Y-%m-%d") - datetime.strptime(start_date, "%Y-%m-%d")).days
    except Exception:
        span_days = 0

    metrics = {
        "start_date": start_date,
        "end_date": end_date,
        "span_days": span_days,
        "trading_days": len(daily_returns),
        "account_return": round(account_total_return * 100, 2),
        "benchmark_return": round(bench_total_return * 100, 2),
        "excess_return": round((account_total_return - bench_total_return) * 100, 2),
        "max_drawdown": round(max_dd * 100, 2),
        "volatility": round(vol * 100, 2),
        "sharpe": round(sharpe, 2),
        "win_rate": round(win_rate * 100, 1),
        "benchmark_name": BENCHMARK_NAME,
        "note": "账户收益按净值快照扣除出入金后的时间加权口径近似计算",
    }
    return {"points": points, "metrics": metrics}


# ============ 收益归因 ============
def attribution():
    holdings = agent.load_portfolio()
    realized = defaultdict(float)
    for t in agent.load_transactions():
        if t["direction"] == "卖出" and t["realized"] is not None:
            realized[t["name"]] += t["realized"]
    # 交易流水覆盖不到的老仓 / 账户二，用「已实现盈亏」补
    for r in agent.get_legacy_realized_records():
        realized[r["name"]] += _num(r["pnl"])

    dividends = defaultdict(float)
    for a in agent.load_corporate_actions():
        if a["type"] in agent.CASH_ACTION_TYPES:
            dividends[a["name"] or "利息/调账"] += _num(a["amount"])

    items = {}
    active = set()
    for h in holdings:
        if h["类别"] != "股票":
            continue
        name = h["名称"]
        items.setdefault(name, {"name": name, "sector": agent.SECTOR_MAP.get(name, "其他"),
                                "unrealized": 0.0, "realized": 0.0, "dividend": 0.0,
                                "market_value": 0.0})
        items[name]["unrealized"] += _num(h["盈亏"])
        items[name]["market_value"] += _num(h["市值"])
        if _num(h["数量"]) > 0:
            active.add(name)
    for name, pnl in realized.items():
        items.setdefault(name, {"name": name, "sector": agent.SECTOR_MAP.get(name, "其他"),
                                "unrealized": 0.0, "realized": 0.0, "dividend": 0.0,
                                "market_value": 0.0})
        items[name]["realized"] += pnl
    for name, amt in dividends.items():
        items.setdefault(name, {"name": name, "sector": "现金",
                                "unrealized": 0.0, "realized": 0.0, "dividend": 0.0,
                                "market_value": 0.0})
        items[name]["dividend"] += amt

    gold_unrealized = sum(_num(h["盈亏"]) for h in holdings if h["类别"] == "黄金")
    rows = []
    for it in items.values():
        total = it["unrealized"] + it["realized"] + it["dividend"]
        if abs(total) < 0.01:
            continue
        rows.append({
            "name": it["name"],
            "sector": it["sector"],
            "unrealized": round(it["unrealized"], 2),
            "realized": round(it["realized"], 2),
            "dividend": round(it["dividend"], 2),
            "pnl": round(total, 2),
            "market_value": round(it["market_value"], 2),
            "cleared": (it["name"] not in active and it["market_value"] == 0
                        and abs(it["realized"]) > 0.01),
        })
    if abs(gold_unrealized) >= 0.01:
        rows.append({"name": "黄金", "sector": "黄金", "unrealized": round(gold_unrealized, 2),
                     "realized": 0.0, "dividend": 0.0, "pnl": round(gold_unrealized, 2),
                     "market_value": round(sum(_num(h["市值"]) for h in holdings if h["类别"] == "黄金"), 2),
                     "cleared": False})

    total = sum(r["pnl"] for r in rows)
    denom = abs(total) if abs(total) > 1 else 1
    for r in rows:
        r["share"] = round(r["pnl"] / denom * 100, 1)
    rows.sort(key=lambda r: r["pnl"])
    return {"items": rows, "total": round(total, 2)}


def opportunity_cost():
    """机会成本：每只持仓从首次买入至今，与同期沪深300比较"""
    holdings = [h for h in agent.load_portfolio()
                if h["类别"] == "股票" and _num(h["数量"]) > 0]
    if not holdings:
        return {"items": [], "summary": {}}

    first_buy = {}
    for t in sorted(agent.load_transactions(), key=lambda x: x["date"]):
        if t["direction"] == "买入" and t["name"] not in first_buy and t["price"]:
            first_buy[t["name"]] = {"date": t["date"][:10], "price": t["price"]}

    index = fetch_index_history(240)
    dates = sorted(index)
    if not dates:
        return {"items": [], "summary": {"error": "沪深300行情不可用"}}

    bench_now = index[dates[-1]]
    items = []
    for h in holdings:
        fb = first_buy.get(h["名称"])
        price, qty, cost = _num(h["现价"]), _num(h["数量"]), _num(h["成本价"])
        if not fb or not price:
            continue
        bench_start_date = next((d for d in dates if d >= fb["date"]), None)
        if not bench_start_date or not index[bench_start_date]:
            continue
        stock_ret = (price / fb["price"] - 1) * 100
        bench_ret = (bench_now / index[bench_start_date] - 1) * 100
        mv = price * qty
        gap_pct = stock_ret - bench_ret
        items.append({
            "name": h["名称"],
            "start": fb["date"],
            "days": (datetime.now() - datetime.strptime(fb["date"], "%Y-%m-%d")).days,
            "first_price": round(fb["price"], 3),
            "price": round(price, 3),
            "cost": round(cost, 3),
            "market_value": round(mv, 2),
            "stock_ret": round(stock_ret, 2),
            "bench_ret": round(bench_ret, 2),
            "gap_pct": round(gap_pct, 2),
            "opportunity": round(mv * gap_pct / 100, 0),
        })

    if not items:
        return {"items": [], "summary": {}}
    items.sort(key=lambda x: x["opportunity"])
    total_mv = sum(i["market_value"] for i in items)
    weighted_stock = sum(i["stock_ret"] * i["market_value"] for i in items) / total_mv if total_mv else 0
    weighted_bench = sum(i["bench_ret"] * i["market_value"] for i in items) / total_mv if total_mv else 0
    return {
        "items": items,
        "summary": {
            "stock_ret": round(weighted_stock, 2),
            "bench_ret": round(weighted_bench, 2),
            "gap_pct": round(weighted_stock - weighted_bench, 2),
            "opportunity": round(sum(i["opportunity"] for i in items), 0),
            "since": min(i["start"] for i in items),
            "benchmark": BENCHMARK_NAME,
        },
    }


# ============ 行业与集中度 ============
def exposure():
    holdings = agent.load_portfolio()
    total_value = sum(_num(h["市值"]) for h in holdings)
    stock_holdings = [h for h in holdings if h["类别"] == "股票" and _num(h["数量"]) > 0]
    stock_value = sum(_num(h["市值"]) for h in stock_holdings)

    sectors = defaultdict(float)
    for h in stock_holdings:
        sectors[agent.SECTOR_MAP.get(h["名称"], "其他")] += _num(h["市值"])
    sector_rows = [
        {"sector": s, "value": round(v, 2),
         "pct_stock": round(v / stock_value * 100, 1) if stock_value else 0,
         "pct_total": round(v / total_value * 100, 1) if total_value else 0}
        for s, v in sectors.items()
    ]
    sector_rows.sort(key=lambda x: -x["value"])

    positions = sorted(
        [{"name": h["名称"], "value": _num(h["市值"]), "qty": int(_num(h["数量"])),
          "price": round(_num(h["现价"]), 3), "cost": round(_num(h["成本价"]), 3),
          "account": h["账户"], "sector": agent.SECTOR_MAP.get(h["名称"], "其他"),
          "pct_total": round(_num(h["市值"]) / total_value * 100, 1) if total_value else 0,
          "return_pct": round(_num(h["收益率"]) * 100, 2)} for h in stock_holdings],
        key=lambda x: -x["value"],
    )
    top1 = positions[0]["pct_total"] if positions else 0
    top3 = sum(p["pct_total"] for p in positions[:3])
    hhi = sum((p["pct_total"] / 100) ** 2 for p in positions)
    cash = sum(_num(h["市值"]) for h in holdings if h["类别"] == "现金")
    return {
        "total_value": round(total_value, 2),
        "stock_value": round(stock_value, 2),
        "cash": round(cash, 2),
        "cash_pct": round(cash / total_value * 100, 1) if total_value else 0,
        "top1_pct": round(top1, 1),
        "top3_pct": round(top3, 1),
        "hhi": round(hhi, 3),
        "positions": positions,
        "sectors": sector_rows,
    }


# ============ 可执行的数字与步骤 ============
def _lots(shares):
    """A 股一手 100 股，卖出股数向下取整到整手"""
    shares = int(shares or 0)
    if shares < 100:
        return 0
    return int(shares // 100) * 100


def _sell_shares(p, amount):
    """卖出 amount 元该持仓需要多少股"""
    if not p.get("price"):
        return 0
    return min(_lots(amount / p["price"]), _lots(p.get("qty", 0)))


def nav_trend(perf=None, window=20):
    """用净值快照判断「企稳」：最近一段的低点后连续 3 个净值日不创新低，且回升 ≥2%。

    用最近 window 个净值日找低点，而不是历史最低点——否则几个月前的老低点会让
    「企稳」永远成立，掩盖当下还在创新低的事实。
    """
    perf = perf or performance()
    pts = perf.get("points") or []
    if len(pts) < 2:
        return {"defined": False, "stabilized": False}
    values = [(p["date"], p["account"]) for p in pts]
    peak_i = max(range(len(values)), key=lambda i: values[i][1])
    peak_date, peak_idx = values[peak_i]
    cur_date, cur_idx = values[-1]

    recent_start = max(0, len(values) - window)
    recent = values[recent_start:]
    low_i = min(range(len(recent)), key=lambda i: recent[i][1])
    trough_date, trough_idx = recent[low_i]
    days_since_low = len(recent) - 1 - low_i
    rebound = (cur_idx / trough_idx - 1) * 100 if trough_idx else 0
    return {
        "defined": True,
        "stabilized": days_since_low >= 3 and rebound >= 2,
        "peak_date": peak_date,
        "peak_index": round(peak_idx, 2),
        "trough_date": trough_date,
        "trough_index": round(trough_idx, 2),
        "current_date": cur_date,
        "current_index": round(cur_idx, 2),
        "days_since_low": days_since_low,
        "rebound_pct": round(rebound, 2),
        "drawdown_pct": round((cur_idx / peak_idx - 1) * 100, 2) if peak_idx else 0,
    }


def _stabilize_status(t):
    if not t.get("defined"):
        return "净值快照不足，暂时无法判断企稳。"
    if t["stabilized"]:
        return (f"现在已满足企稳条件：{t['trough_date']} 见低点后已过 {t['days_since_low']} 个净值日没创新低，"
                f"并且回升 {t['rebound_pct']:+.1f}%（≥2%）。")
    if t["days_since_low"] == 0:
        return (f"最新净值日（{t['trough_date']}）就是这轮低点，还没有「连续 3 个净值日不创新低」，"
                f"也没有「回升 2%」，所以现在没企稳，继续等，等今天收盘后的净值再看。")
    return (f"现在还没满足：{t['trough_date']} 见低点后只过了 {t['days_since_low']} 个净值日"
            f"（要连续 3 个不创新低），净值较低点回升 {t['rebound_pct']:+.1f}%（要 ≥2%）。"
            f"所以现在要继续等，不抄底。")


# ============ 风险规则 ============
def risks(perf=None, attr=None, exp=None):
    perf = perf or performance()
    exp = exp or exposure()
    rows = []

    def add(level, title, detail, action, metric="", steps=None):
        rows.append({"level": level, "title": title, "detail": detail, "action": action,
                     "metric": metric, "steps": steps or []})

    m = perf.get("metrics", {})
    total = exp.get("total_value", 0)
    trend = nav_trend(perf)
    dd = m.get("max_drawdown", 0)

    if dd <= -12 or dd <= -8:
        level = "高" if dd <= -12 else "中"
        budget = total * 0.05
        steps = [
            "暂停买入：不补仓、不定投、不因为跌得多就加仓；已经持有的先不动也可以，重点是别把现金买进去。",
            "「净值企稳」的标准：从这轮低点开始，连续 3 个净值日（每个交易日收盘后记录的净值快照）都不创新低，"
            "并且最新净值比低点高 2% 以上。两个条件同时满足才算企稳。",
            _stabilize_status(trend),
            f"企稳之后：先拿不超过总资产 5%（约 {budget:,.0f} 元）试买，分 2-3 次、每次只买 1 只，"
            "每次买完至少等 3 个净值日再决定下一步。没企稳就一直等。",
        ]
        detail = (f"历史最大回撤 {dd:.1f}%（{trend.get('peak_date', '高点')} 起算），"
                  f"当前距高点 {trend.get('drawdown_pct', 0):+.1f}%。"
                  f"最近 20 个净值日的低点在 {trend.get('trough_date', '—')}，"
                  f"最新 {trend.get('current_date', '—')} 较低点 {trend.get('rebound_pct', 0):+.1f}%。")
        add(level, "回撤超过 12%" if level == "高" else "回撤偏大", detail,
            "暂停买入；净值连续 3 个净值日不创新低且回升 2% 后，再用总资产 5% 分 2-3 次试买。",
            f"{dd:.1f}%", steps)

    if m.get("excess_return", 0) <= -10:
        worst = [p for p in sorted(exp["positions"], key=lambda x: x["return_pct"]) if p["return_pct"] < 0][:2]
        steps = ["先把跑输的标的列成「减仓清单」，不是今天就卖："]
        for p in worst:
            if p["cost"] > p["price"]:
                rise = (p["cost"] / p["price"] - 1) * 100
                steps.append(
                    f"{p['name']}：浮亏 {p['return_pct']:.1f}%，现价 {p['price']:.2f}，成本 {p['cost']:.2f}。"
                    f"回到成本价附近（还要涨 {rise:.1f}%）先减 1/3 ≈ {_lots(p['qty'] / 3)} 股；"
                    f"不到这个价不主动割，继续拿着。")
        steps.append(
            "减出来的钱先放现金，等账户净值企稳（连续 3 个净值日不创新低 + 回升 2%）后，"
            f"再考虑换成沪深300 这类宽基 ETF（第一次不超过总资产 5%，约 {total * 0.05:,.0f} 元）；"
            "不要马上买回同类个股。")
        add("中", "明显跑输沪深300",
            f"同期账户 {m['account_return']:+.1f}%，沪深300 {m['benchmark_return']:+.1f}%，落后 {abs(m['excess_return']):.1f} 个百分点。",
            "把跑输的标的列成减仓清单，反弹到成本附近减 1/3，钱先放现金。", f"{m['excess_return']:+.1f}pp", steps)

    if exp["top1_pct"] > 20 and exp["positions"]:
        p = exp["positions"][0]
        target_value = total * 0.15
        over = p["value"] - target_value
        sell = _sell_shares(p, over)
        if sell:
            half = _lots(sell / 2) or sell
            steps = [
                f"{p['name']} 现价 {p['price']:.2f}，持有 {p['qty']} 股，市值 {p['value']:,.0f} 元，"
                f"占总资产 {p['pct_total']:.1f}%（上限 20%，建议压到 15%）。",
                f"压到 15%（{target_value:,.0f} 元）需要卖出约 {over:,.0f} 元，按现价约 {sell} 股。",
                f"分 2 次卖：找反弹的一天先卖 {half} 股（约 {half * p['price']:,.0f} 元）；"
                f"隔至少 1 个交易日，再卖剩下的 {sell - half} 股；不要在大跌当天一次性砸出。",
                "卖出的钱先放现金，不马上换股，等净值企稳后再配置。",
            ]
            action = f"{p['name']} 分 2 次减 {sell} 股（约 {over:,.0f} 元），把占比从 {p['pct_total']:.1f}% 降到 15%。"
        else:
            steps = [f"{p['name']} 占总资产 {p['pct_total']:.1f}%，超出部分不足一手，暂不需要卖。"]
            action = f"{p['name']} 占比 {p['pct_total']:.1f}%，超出部分不足一手，先不操作。"
        add("高", "单票集中度偏高", f"第一大持仓 {p['name']} 占总资产 {p['pct_total']:.1f}%，超过 20% 上限。",
            action, f"{exp['top1_pct']:.1f}%", steps)
    elif exp["top3_pct"] > 45 and len(exp["positions"]) >= 3:
        top3 = exp["positions"][:3]
        over = sum(p["value"] for p in top3) - total * 0.40
        p = top3[0]
        sell = _sell_shares(p, over)
        names = "、".join(f"{x['name']} {x['pct_total']:.1f}%" for x in top3)
        steps = [
            f"前三大持仓：{names}，合计 {exp['top3_pct']:.1f}%（目标 ≤40%）。",
            (f"最直接的办法是从里面最大的 {p['name']} 减：卖 {sell} 股（约 {sell * p['price']:,.0f} 元），"
             f"分 2 次，每次 {(sell // 2 // 100) * 100 or sell} 股，间隔至少 1 个交易日。") if sell else
            "超出金额不足一手，暂时不用卖，新增资金不要再买这三只。",
            "卖出的钱先放现金；以后新增资金优先买和现有行业不同的宽基或弱相关行业。",
        ]
        add("中", "头部持仓过重", f"前三大持仓合计占总资产 {exp['top3_pct']:.1f}%。",
            f"把前三大压到 40% 以下：{p['name']} 减约 {sell} 股，分 2 次卖。", f"{exp['top3_pct']:.1f}%", steps)

    top_sector = exp["sectors"][0] if exp["sectors"] else None
    if top_sector and top_sector["pct_total"] > 30:
        target_value = total * 0.25
        over = top_sector["value"] - target_value
        members = [p for p in exp["positions"] if p["sector"] == top_sector["sector"]]
        p = max(members, key=lambda x: x["value"]) if members else None
        sell = _sell_shares(p, over) if p else 0
        steps = [
            f"{top_sector['sector']} 合计 {top_sector['value']:,.0f} 元，占总资产 {top_sector['pct_total']:.1f}%（目标 ≤25%）。"
            f"需要减少约 {over:,.0f} 元。",
            (f"从该行业里最大的 {p['name']} 减 {sell} 股（约 {sell * p['price']:,.0f} 元），分 2 次卖。"
             if sell else "超出金额不足一手，暂时不卖，新增资金不要再买这个行业。"),
            "减出来的钱先放现金；新增资金优先买和该行业弱相关的品种（例如宽基指数），不要继续加同一行业。",
        ]
        add("高", "行业过度集中", f"{top_sector['sector']} 占总资产 {top_sector['pct_total']:.1f}%，单一行业风险过高。",
            f"把 {top_sector['sector']} 从 {top_sector['pct_total']:.1f}% 降到 25%：{p['name'] if p else ''} 减约 {sell} 股。",
            f"{top_sector['pct_total']:.1f}%", steps)
    elif top_sector and top_sector["pct_stock"] > 45:
        members = [p for p in exp["positions"] if p["sector"] == top_sector["sector"]]
        others = [s["sector"] for s in exp["sectors"][1:4]]
        steps = [
            f"{top_sector['sector']} 占股票仓位 {top_sector['pct_stock']:.1f}%，股票部分过于集中。",
            f"以后新增的股票资金不要再买这个行业；现在持有 {len(members)} 只："
            + "、".join(f"{p['name']}（{p['return_pct']:+.1f}%）" for p in members) + "。",
            f"要补行业，优先从和它弱相关的方向里选 1-2 个（现在已有：{'、'.join(others) or '—'}），"
            f"单只新买入不超过总资产 5%（约 {total * 0.05:,.0f} 元）。",
        ]
        add("中", "股票仓位行业集中", f"{top_sector['sector']} 占股票仓位 {top_sector['pct_stock']:.1f}%。",
            "新增资金不买该行业，补 1-2 个弱相关行业。", f"{top_sector['pct_stock']:.1f}%", steps)

    for p in exp["positions"]:
        if p["return_pct"] <= -20 and p["price"] > 0:
            level = "高" if p["return_pct"] <= -30 else "中"
            q1, q2 = _lots(p["qty"] / 3), _lots(p["qty"] / 3)
            p1, p2 = p["price"] * 1.10, p["price"] * 1.20
            stop = p["price"] * 0.92
            steps = [
                f"{p['name']}：现价 {p['price']:.2f}，成本 {p['cost']:.2f}，浮亏 {p['return_pct']:.1f}%，"
                f"持有 {p['qty']} 股，市值 {p['value']:,.0f} 元。",
                f"第一档：反弹到 {p1:.2f}（较现价 +10%）卖 1/3 ≈ {q1} 股，约 {q1 * p1:,.0f} 元。",
                f"第二档：反弹到 {p2:.2f}（较现价 +20%）再卖 1/3 ≈ {q2} 股，约 {q2 * p2:,.0f} 元。",
                f"防守线：如果从现价再跌 8%（到 {stop:.2f}），且当天没有明确利好，就把剩余仓位减半，"
                "避免亏损继续扩大。",
                "两档都没到就持有，不补仓摊薄；不要因为「已经亏很多」就一把梭补。",
            ]
            add(level, f"{p['name']} 深度亏损" if level == "高" else f"{p['name']} 亏损扩大",
                f"浮亏 {p['return_pct']:.1f}%，占总资产 {p['pct_total']:.1f}%。",
                f"{p['name']}：反弹到 {p1:.2f} 减 1/3（{q1} 股），到 {p2:.2f} 再减 1/3；跌破 {stop:.2f} 剩余减半。",
                f"{p['return_pct']:.1f}%", steps)

    if exp["cash_pct"] < 10:
        need = total * 0.10 - exp["cash"]
        weakest = min([p for p in exp["positions"] if p["qty"] > 0] or [None],
                      key=lambda x: x["return_pct"] if x else 0)
        sell = _sell_shares(weakest, need) if weakest else 0
        steps = [
            f"现金 {exp['cash']:,.0f} 元，占总资产 {exp['cash_pct']:.1f}%（底线 5%，建议 10%）。",
            f"把现金补到 10%（{total * 0.10:,.0f} 元）还需要 {max(need, 0):,.0f} 元。",
            (f"从浮亏最大的 {weakest['name']} 卖 {sell} 股（约 {sell * weakest['price']:,.0f} 元）来补；"
             "如果它已经深亏、不想割，就等反弹到 +10% 再卖。") if sell else
            "持仓都不足一手，暂时卖不出，先用后续收入/分红补现金。",
            "补出来的现金先放着，不要马上买回。",
        ]
        add("高" if exp["cash_pct"] < 5 else "中",
            "现金缓冲不足" if exp["cash_pct"] < 5 else "现金偏少",
            f"现金占总资产 {exp['cash_pct']:.1f}%，低于{'5% 的应急底线' if exp['cash_pct'] < 5 else '建议的 10%'}。",
            f"把现金补到 10%：还需约 {max(need, 0):,.0f} 元，优先从 {weakest['name'] if weakest else '浮亏最大持仓'} 减仓。",
            f"{exp['cash_pct']:.1f}%", steps)
    elif exp["cash_pct"] > 25:
        batch = total * 0.03
        steps = [
            f"现金 {exp['cash']:,.0f} 元，占总资产 {exp['cash_pct']:.1f}%，超过 25% 的闲置线。",
            (f"分 3 批配置宽基：每次约 {batch:,.0f} 元（总资产 3%），间隔至少 5 个交易日，"
             "优先沪深300 / 中证500 这类宽基，不买行业主题。")
            if trend.get("stabilized") else
            "先别买：现在净值还没企稳（3 个净值日不创新低 + 回升 2%），现金先放着，等信号出现再分 3 批买。",
            "也可以先放货币基金/国债逆回购，别让钱完全闲置。",
        ]
        add("低", "现金闲置偏多", f"现金占总资产 {exp['cash_pct']:.1f}%，超过 25%。",
            "净值企稳后分 3 批买宽基，每次 3% 总资产。", f"{exp['cash_pct']:.1f}%", steps)

    stock_count = len([p for p in exp["positions"] if p["value"] > 0])
    if stock_count < 4:
        sectors = [s["sector"] for s in exp["sectors"]]
        steps = [
            f"现在只有 {stock_count} 只有效持仓，至少补到 5 只、3-5 个行业。",
            f"单只新买入不超过总资产 15%（约 {total * 0.15:,.0f} 元），第一次先买 5%（约 {total * 0.05:,.0f} 元）。",
            f"避开已有行业：{'、'.join(sectors) or '—'}；优先选弱相关的宽基或行业。",
            "等净值企稳后再买，现在先选好标的、列好价格，不动手。",
        ]
        add("中", "持仓过于集中", f"仅有 {stock_count} 只有效持仓。",
            f"补到 5 只、3-5 个行业，单只不超过 15%。", f"{stock_count} 只", steps)
    elif stock_count > 10:
        small = [p for p in exp["positions"] if p["pct_total"] < 2]
        steps = [
            f"持有 {stock_count} 只，超过 10 只，精力会被摊薄。",
            f"先把占总资产不到 2%（{total * 0.02:,.0f} 元）的小仓位列出来："
            + ("、".join(f"{p['name']}（{p['pct_total']:.1f}%）" for p in small) if small else "暂无") + "。",
            "优先清理其中浮亏且长期跑输的，一次卖 1 只，卖完的钱先放现金。",
            "保留 6-8 只有把握的，单只不超过总资产 20%。",
        ]
        add("低", "持仓过于分散", f"持有 {stock_count} 只股票，管理精力被摊薄。",
            "清理 2% 以下的小仓位，保留 6-8 只。", f"{stock_count} 只", steps)

    order = {"高": 0, "中": 1, "低": 2}
    rows.sort(key=lambda r: order.get(r["level"], 3))
    return rows


def actions(risk_rows):
    """把风险转换成可勾选的行动清单"""
    out = []
    for i, r in enumerate(risk_rows, 1):
        if r["level"] == "低":
            continue
        out.append({
            "priority": "今日" if r["level"] == "高" else "本周",
            "level": r["level"],
            "action": r["action"],
            "reason": r["title"],
        })
    if not out:
        out.append({"priority": "维持", "level": "低", "action": "维持现有仓位，按计划复盘", "reason": "当前无高风险项"})
    return out[:6]


# ============ 今日提醒 ============
def alerts(stocks):
    out = []
    for s in stocks:
        qty = s.get("qty") or 0
        price = s.get("price") or 0
        cost = s.get("cost") or 0
        today_pnl = s.get("today_pnl") or 0
        if qty > 0 and price > 0:
            prev_close = price - today_pnl / qty
            if prev_close > 0:
                chg = (price / prev_close - 1) * 100
                if abs(chg) >= 3:
                    out.append({"type": "异动", "name": s["name"],
                                "text": f"今日 {chg:+.1f}%（{today_pnl:+,.0f} 元）",
                                "level": "高" if abs(chg) >= 5 else "中"})
        if qty > 0 and cost > 0 and price > 0:
            gap = (price / cost - 1) * 100
            if -5 <= gap < 0:
                out.append({"type": "接近回本", "name": s["name"],
                            "text": (f"距成本 {abs(gap):.1f}%，回本价 {cost:.2f}；"
                                     f"到价可先减 1/3（{_lots(qty / 3)} 股，约 {_lots(qty / 3) * cost:,.0f} 元），"
                                     "剩下的再看。"),
                            "level": "低"})
    order = {"高": 0, "中": 1, "低": 2}
    out.sort(key=lambda x: order.get(x["level"], 3))
    return out


def recovery():
    """离回本还有多远：账户整体 + 每只亏损持仓"""
    holdings = agent.load_portfolio()
    total = sum(_num(h["市值"]) for h in holdings)
    net_capital = _num(agent.get_net_capital()[0])
    need_amount = max(net_capital - total, 0)
    positions = []
    for h in holdings:
        if h["类别"] != "股票" or _num(h["数量"]) <= 0:
            continue
        cost, price, qty = _num(h["成本价"]), _num(h["现价"]), _num(h["数量"])
        if cost <= 0 or price <= 0 or price >= cost:
            continue
        positions.append({
            "name": h["名称"],
            "price": round(price, 2),
            "cost": round(cost, 2),
            "qty": int(qty),
            "rise_needed": round((cost / price - 1) * 100, 1),
            "loss": round((price - cost) * qty, 2),
            "value": round(price * qty, 2),
            "account": h["账户"],
        })
    positions.sort(key=lambda x: x["loss"])
    return {
        "total_value": round(total, 2),
        "net_capital": round(net_capital, 2),
        "need_amount": round(need_amount, 2),
        "need_pct": round(need_amount / total * 100, 2) if total else 0,
        "positions": positions,
    }


def health(perf=None, exp=None):
    """0-100 健康分，扣分项可解释"""
    perf = perf or performance()
    exp = exp or exposure()
    m = perf.get("metrics", {})
    score = 100.0
    deductions = []

    def cut(points, reason):
        nonlocal score
        score -= points
        deductions.append({"points": points, "reason": reason})

    dd = m.get("max_drawdown", 0)
    if dd <= -12:
        cut(20, f"最大回撤 {dd:.1f}%")
    elif dd <= -8:
        cut(12, f"最大回撤 {dd:.1f}%")
    elif dd <= -5:
        cut(6, f"最大回撤 {dd:.1f}%")

    top1 = exp.get("top1_pct", 0)
    if top1 > 25:
        cut(18, f"第一大持仓占 {top1:.1f}%")
    elif top1 > 20:
        cut(10, f"第一大持仓占 {top1:.1f}%")
    elif top1 > 15:
        cut(5, f"第一大持仓占 {top1:.1f}%")

    top3 = exp.get("top3_pct", 0)
    if top3 > 50:
        cut(10, f"前三大持仓占 {top3:.1f}%")
    elif top3 > 45:
        cut(5, f"前三大持仓占 {top3:.1f}%")

    top_sector = (exp.get("sectors") or [{}])[0]
    if top_sector.get("pct_total", 0) > 35:
        cut(12, f"{top_sector['sector']} 占 {top_sector['pct_total']:.1f}%")
    elif top_sector.get("pct_total", 0) > 30:
        cut(6, f"{top_sector['sector']} 占 {top_sector['pct_total']:.1f}%")

    cash = exp.get("cash_pct", 0)
    if cash < 5:
        cut(15, f"现金仅 {cash:.1f}%")
    elif cash < 10:
        cut(8, f"现金仅 {cash:.1f}%")
    elif cash > 30:
        cut(4, f"现金闲置 {cash:.1f}%")

    deep = [p for p in exp.get("positions", []) if p.get("return_pct", 0) <= -20]
    if deep:
        cut(min(18, 6 * len(deep)), f"{len(deep)} 只持仓浮亏超 20%")

    count = len([p for p in exp.get("positions", []) if p.get("value", 0) > 0])
    if count < 4:
        cut(8, f"仅持有 {count} 只股票")

    score = max(0, min(100, round(score)))
    level = "健康" if score >= 80 else "一般" if score >= 65 else "偏弱" if score >= 50 else "高风险"
    deductions.sort(key=lambda x: -x["points"])
    return {"score": score, "level": level, "deductions": deductions[:4]}


def verdict(snap):
    """把数据压成一句话：今天该不该动手 + 为什么 + 具体怎么做"""
    alerts = snap.get("alerts") or []
    risk_rows = snap.get("risks") or []
    perf = snap.get("performance") or {}
    exp = snap.get("exposure") or {}
    rec = snap.get("recovery") or {}
    trend = nav_trend(perf)

    high_risks = [r for r in risk_rows if r["level"] == "高"]
    high_alerts = [a for a in alerts if a["level"] == "高"]

    if high_risks or high_alerts:
        headline = "今天：不买，先控风险"
        tone = "warning"
    elif risk_rows:
        headline = "今天：不新开仓，处理现有持仓"
        tone = "caution"
    else:
        headline = "今天：不用动手，持有观察"
        tone = "calm"

    reasons = []
    for a in high_alerts[:2]:
        reasons.append(f"{a['name']} {a['text']}")
    for r in high_risks[:2]:
        reasons.append(r["detail"])
    if len(reasons) < 3 and rec.get("need_pct", 0) > 0:
        reasons.append(f"账户回本需整体上涨 {rec['need_pct']:.1f}%")
    if len(reasons) < 3 and exp.get("cash_pct"):
        reasons.append(f"现金占总资产 {exp['cash_pct']:.1f}%")

    triggers = []
    if trend.get("defined"):
        if trend["stabilized"]:
            triggers.append(
                f"「净值企稳」已满足：{trend['trough_date']} 见低点，已连续 {trend['days_since_low']} 个净值日不创新低，"
                f"回升 {trend['rebound_pct']:+.1f}%。可以开始用不超过总资产 5%"
                f"（约 {exp.get('total_value', 0) * 0.05:,.0f} 元）分 2-3 次试买。")
        else:
            triggers.append(
                f"「净值企稳」没满足：{trend['trough_date']} 见低点后只有 {trend['days_since_low']} 个净值日"
                f"（要 3 个），较低点回升 {trend['rebound_pct']:+.1f}%（要 ≥2%）。"
                "现在做什么：不买、不抄底；下一个净值日再看这两个数字。")

    for p in rec.get("positions", [])[:2]:
        if p["rise_needed"] >= 15 and p["price"] > 0:
            q1 = _lots(p["qty"] / 3)
            triggers.append(
                f"{p['name']} 现价 {p['price']:.2f}，成本 {p['cost']:.2f}（回本要涨 {p['rise_needed']:.1f}%）。"
                f"反弹到 {p['price'] * 1.1:.2f}（+10%）先卖 {q1} 股（约 {q1 * p['price'] * 1.1:,.0f} 元）；"
                f"到 {p['price'] * 1.2:.2f}（+20%）再卖 {q1} 股。今天不补仓。")

    if exp.get("cash_pct", 0) > 20:
        triggers.append(
            f"现金 {exp['cash']:,.0f} 元（占总资产 {exp['cash_pct']:.1f}%）：现在保留不动，不抄底。"
            f"低于 10%（约 {exp.get('total_value', 0) * 0.10:,.0f} 元）时才需要从持仓减仓补现金。")
    if not triggers:
        triggers.append("没有触发条件，今天不需要操作。")

    return {"headline": headline, "tone": tone, "reasons": reasons[:3], "triggers": triggers[:3]}


def glossary(perf=None, exp=None):
    """把建议里出现的术语翻译成能直接照做的标准"""
    trend = nav_trend(perf)
    total = (exp or {}).get("total_value", 0)
    return [
        {"term": "净值企稳",
         "meaning": ("从这轮低点开始，连续 3 个净值日（每个交易日收盘后自动记录的净值快照）"
                     "都不创新低，并且最新净值比低点高 2% 以上，两个条件同时满足才算。"
                     + _stabilize_status(trend))},
        {"term": "净值日",
         "meaning": "指盘后记录一条净值快照的日子，正常一个交易日一条；周末和节假日没有。"},
        {"term": "最大回撤",
         "meaning": "从历史最高点跌到最低点的幅度，衡量最坏情况下账户缩水多少。"},
        {"term": "回本需涨",
         "meaning": "（净投入本金 − 当前总资产）÷ 当前总资产。比如还差 10%，账户整体涨 10% 才回本。"},
        {"term": "1 手",
         "meaning": "A 股 1 手 = 100 股，买入必须是整手；所以补仓金额会按「手数 × 100 × 现价」计算。"},
        {"term": "1/3 仓位",
         "meaning": "按持有的股数算，不是按金额猜。例如持有 900 股，1/3 就是卖 300 股；不足 100 股按 100 股算。"},
        {"term": "试买额度 5%",
         "meaning": f"按总资产算。现在总资产 {total:,.0f} 元，5% 约为 {total * 0.05:,.0f} 元，"
                    "分成 2-3 次买，不是一次买完。"},
        {"term": "摊薄成本",
         "meaning": ("卖出盈利会摊低剩余持仓的成本价，所以可能出现负数；负成本表示这只已经全部回本还有赚，"
                     "不是数据错误。")},
    ]


def market_snapshot():
    """沪深300 近期走势，给 AI 判断大盘环境用"""
    bench = fetch_index_history(240)
    if not bench:
        return {"defined": False}
    dates = sorted(bench)
    closes = [bench[d] for d in dates]
    last_date, last = dates[-1], closes[-1]

    def ret(n):
        if len(closes) <= n:
            return None
        return round((last / closes[-1 - n] - 1) * 100, 2)

    win20 = closes[-20:] if len(closes) >= 20 else closes
    win60 = closes[-60:] if len(closes) >= 60 else closes
    high20, low20 = max(win20), min(win20)
    up_streak = down_streak = 0
    for i in range(len(closes) - 1, 0, -1):
        if closes[i] > closes[i - 1]:
            up_streak += 1
        else:
            break
    for i in range(len(closes) - 1, 0, -1):
        if closes[i] < closes[i - 1]:
            down_streak += 1
        else:
            break
    return {
        "defined": True,
        "name": BENCHMARK_NAME,
        "date": last_date,
        "close": round(last, 2),
        "ret_5d": ret(5),
        "ret_20d": ret(20),
        "ret_60d": ret(60),
        "high_20d": round(high20, 2),
        "low_20d": round(low20, 2),
        "off_high_20d": round((last / high20 - 1) * 100, 2) if high20 else 0,
        "above_low_20d": round((last / low20 - 1) * 100, 2) if low20 else 0,
        "high_60d": round(max(win60), 2),
        "low_60d": round(min(win60), 2),
        "up_streak": up_streak,
        "down_streak": down_streak,
    }


def _theme(p):
    sector = (p.get("sector") or "").split("（")[0]
    name = p.get("name") or ""
    if "半导体" in name or "半导体" in sector or "芯片" in sector:
        return "半导体"
    if "新能源车" in sector:
        return "新能源车"
    if "消费电子" in sector:
        return "消费电子"
    return sector or "其他"


def structure_advice(exp=None):
    """直接回答：股票数量多不多、哪些持仓重复"""
    exp = exp or exposure()
    positions = [p for p in exp.get("positions", []) if p["value"] > 0]
    n = len(positions)
    total = exp.get("total_value", 0)
    if n == 0:
        return {"count": 0, "verdict": "偏少", "headline": "目前没有股票持仓。", "points": []}

    if 5 <= n <= 8:
        verdict = "合适"
    elif n > 8:
        verdict = "偏多"
    else:
        verdict = "偏少"

    points = []
    small = [p for p in positions if p["pct_total"] < 3.0]
    if small:
        names = "、".join(f"{p['name']}（{p['pct_total']:.1f}%）" for p in small)
        points.append(
            f"有 {len(small)} 只占比低于 3%：{names}，合计 {sum(p['value'] for p in small):,.0f} 元、"
            f"{sum(p['pct_total'] for p in small):.1f}%。这类小仓位对收益贡献有限，却要花精力盯，"
            "建议合并或清掉。")

    groups = defaultdict(list)
    for p in positions:
        groups[_theme(p)].append(p)
    overlaps = []
    for theme, ps in groups.items():
        if len(ps) < 2:
            continue
        pct = sum(p["pct_total"] for p in ps)
        keep = max(ps, key=lambda p: p["value"])
        merge_out = min(ps, key=lambda p: p["value"])
        overlaps.append({"theme": theme, "names": [p["name"] for p in ps],
                         "pct_total": round(pct, 1), "keep": keep["name"],
                         "merge_out": merge_out["name"],
                         "merge_value": round(merge_out["value"], 0)})
        if pct >= 12:
            points.append(
                f"{theme} 有 {len(ps)} 只（{'、'.join(p['name'] for p in ps)}），合计占总资产 {pct:.1f}%，"
                f"暴露偏重。以后加仓只能选其中一只；如果只想留一只，优先保留 {keep['name']}，"
                f"卖出 {merge_out['name']}（约 {merge_out['value']:,.0f} 元）。")
        else:
            points.append(
                f"{theme} 有 {len(ps)} 只（{'、'.join(p['name'] for p in ps)}），合计 {pct:.1f}%，"
                f"方向重复但金额不大；二选一保留即可，卖出较小的 {merge_out['name']}"
                f"（约 {merge_out['value']:,.0f} 元），别两边都补。")

    if not small and not overlaps:
        points.append("没有低于 3% 的碎仓，也没有明显的同行业重复，结构是干净的。")

    if verdict == "偏多":
        headline = f"当前 {n} 只，偏多。按你的资金量（约 {total / 10000:.1f} 万），5-8 只足够覆盖，建议先清理小仓位。"
    elif verdict == "偏少":
        headline = f"当前 {n} 只，偏少。按你的资金量（约 {total / 10000:.1f} 万），5-8 只、3-5 个行业更稳。"
    else:
        headline = (f"当前 {n} 只，数量合适（5-8 只为宜），不用为了凑数再买；"
                    f"下一步是把结构理干净。")

    return {
        "count": n,
        "verdict": verdict,
        "ideal_range": "5-8 只",
        "headline": headline,
        "points": points[:4],
        "small_positions": [p["name"] for p in small],
        "overlaps": overlaps,
    }


def _gold_position():
    for h in agent.load_portfolio():
        if h["类别"] == "黄金" and _num(h["数量"]) > 0:
            return {
                "name": h["名称"], "grams": round(_num(h["数量"]), 2),
                "cost": round(_num(h["成本价"]), 2), "price": round(_num(h["现价"]), 2),
                "value": round(_num(h["市值"]), 2), "pnl": round(_num(h["盈亏"]), 2),
            }
    return None


def discipline(perf=None, exp=None):
    """把她的交易习惯固化成可检查的纪律：止盈档位、卖出红线、补仓门槛"""
    perf = perf or performance()
    exp = exp or exposure()
    total = exp.get("total_value", 0)
    trend = nav_trend(perf)

    take_profit = []
    red_lines = []
    for p in exp.get("positions", []):
        if p["price"] <= 0 or p["cost"] <= 0 or p["qty"] <= 0:
            continue
        ret = p["return_pct"]
        hit = [t for t in DISCIPLINE["take_profit_tiers"] if ret >= t]
        if hit:
            tier = max(hit)
            qty = _lots(p["qty"] * len(hit) / 3)
            take_profit.append({
                "name": p["name"], "return_pct": ret, "tier": tier,
                "target_price": round(p["cost"] * (1 + tier / 100), 2),
                "qty": qty, "amount": round(qty * p["price"], 0),
                "message": (f"{p['name']}：已到 +{tier:.0f}% 止盈档，按纪律累计卖 {len(hit)}/3"
                            f"（{qty} 股，约 {qty * p['price']:,.0f} 元）"),
            })
        if ret <= DISCIPLINE["red_line_pct"]:
            qty = _lots(p["qty"] / 2)
            red_lines.append({
                "name": p["name"], "return_pct": ret,
                "cut_price": round(p["cost"] * (1 + DISCIPLINE["red_line_pct"] / 100), 2),
                "qty": qty, "amount": round(qty * p["price"], 0),
                "message": (f"{p['name']}：浮亏 {ret:.1f}%，触发 {DISCIPLINE['red_line_pct']:.0f}% 红线；"
                            f"按纪律减半 {qty} 股（约 {qty * p['price']:,.0f} 元）"),
            })

    gold = _gold_position()
    gold_tiers = []
    if gold and gold["cost"] > 0 and gold["price"] > 0:
        for tier in DISCIPLINE["gold_tiers"]:
            target = gold["cost"] * (1 + tier / 100)
            gold_tiers.append({
                "label": "回本" if tier == 0 else f"+{tier:.0f}%",
                "price": round(target, 2),
                "distance_pct": round((target / gold["price"] - 1) * 100, 1),
                "reached": gold["price"] >= target,
                "grams": round(gold["grams"] / 3, 2),
                "amount": round(gold["grams"] / 3 * target, 0),
            })

    gates = [
        {"key": "cash", "label": f"现金 ≥{DISCIPLINE['cash_floor_pct']:.0f}%",
         "ok": exp.get("cash_pct", 0) >= DISCIPLINE["cash_floor_pct"],
         "current": f"现在 {exp.get('cash_pct', 0):.1f}%"},
        {"key": "single", "label": f"单票 ≤{DISCIPLINE['single_cap_pct']:.0f}%",
         "ok": exp.get("top1_pct", 0) <= DISCIPLINE["single_cap_pct"],
         "current": f"最大 {exp.get('top1_pct', 0):.1f}%"},
        {"key": "trade_size", "label": f"单笔 ≤{DISCIPLINE['trade_size_pct']:.0f}%",
         "ok": None, "current": f"约 {total * DISCIPLINE['trade_size_pct'] / 100:,.0f} 元"},
        {"key": "nav", "label": "净值企稳",
         "ok": trend.get("stabilized", False),
         "current": (f"已连续 {trend.get('days_since_low', 0)} 日不创新低、回升 {trend.get('rebound_pct', 0):+.1f}%"
                     if trend.get("defined") else "快照不足")},
        {"key": "fundamental", "label": "基本面/消息没变坏",
         "ok": None, "current": "需要你手动确认"},
    ]
    return {
        "config": DISCIPLINE,
        "gates": gates,
        "take_profit": take_profit,
        "red_lines": red_lines,
        "gold": gold,
        "gold_tiers": gold_tiers,
        "nav": trend,
        "suggested_max": round(min(total * DISCIPLINE["trade_size_pct"] / 100,
                                   max(exp.get("cash", 0) - total * DISCIPLINE["cash_floor_pct"] / 100, 0)), 0),
    }


def add_check(name, amount=None, price=None, fundamentals_ok=False, lots=None):
    """补仓前算账：按手数（1 手 = 100 股）计算，五道门槛 + 补后成本/仓位/现金/跌幅损失"""
    holdings = agent.load_portfolio()
    total = sum(_num(h["市值"]) for h in holdings)
    cash = sum(_num(h["市值"]) for h in holdings if h["类别"] == "现金")
    stock = next((h for h in holdings if h["类别"] == "股票" and h["名称"] == name), None)
    if not stock:
        raise ValueError("没找到这只股票的持仓")
    qty_old = _num(stock["数量"])
    cost_old = _num(stock["成本价"])
    price = float(price or _num(stock["现价"]))
    if price <= 0:
        raise ValueError("现价无效，请先刷新行情")

    if lots is not None:
        qty = int(lots) * 100
        if qty <= 0:
            raise ValueError("补仓手数至少 1 手（100 股）")
    else:
        amount = float(amount or 0)
        if amount <= 0:
            raise ValueError("请输入补仓手数")
        qty = _lots(amount / price)
        if qty <= 0:
            raise ValueError(f"金额不足一手（100 股约需 {price * 100:,.0f} 元）")
    actual = qty * price
    if actual > cash:
        raise ValueError(f"现金只有 {cash:,.0f} 元，不够这笔补仓（需 {actual:,.0f} 元）")
    new_qty = qty_old + qty
    new_cost = (qty_old * cost_old + actual) / new_qty
    new_value = new_qty * price
    cash_after = cash - actual
    cash_after_pct = cash_after / total * 100 if total else 0
    single_pct = new_value / total * 100 if total else 0

    trend = nav_trend()
    gates = [
        {"key": "cash", "label": "现金 ≥20%",
         "ok": cash_after_pct >= DISCIPLINE["cash_floor_pct"],
         "current": f"补后 {cash_after_pct:.1f}%",
         "detail": f"补完剩 {cash_after:,.0f} 元，底线 {total * 0.20:,.0f} 元"},
        {"key": "single", "label": "单票 ≤20%",
         "ok": single_pct <= DISCIPLINE["single_cap_pct"],
         "current": f"补后 {single_pct:.1f}%",
         "detail": f"补完市值 {new_value:,.0f} 元，上限 {total * 0.20:,.0f} 元"},
        {"key": "trade_size", "label": "单笔 ≤5%",
         "ok": actual <= total * DISCIPLINE["trade_size_pct"] / 100 + 1,
         "current": f"这笔 {actual:,.0f} 元",
         "detail": f"单笔上限 {total * 0.05:,.0f} 元"},
        {"key": "nav", "label": "净值企稳",
         "ok": bool(trend.get("stabilized")),
         "current": ("已企稳" if trend.get("stabilized") else "没企稳"),
         "detail": _stabilize_status(trend)},
        {"key": "fundamental", "label": "基本面/消息没变坏",
         "ok": bool(fundamentals_ok),
         "current": "已确认" if fundamentals_ok else "未确认",
         "detail": ("你已勾选确认，请对自己的判断负责" if fundamentals_ok
                    else "需要你确认：公司业绩、行业逻辑、近期消息没有变坏")},
    ]

    def drop_loss(drop):
        return round((price * (1 - drop / 100) - new_cost) * new_qty, 0)

    rise_before = (cost_old / price - 1) * 100 if price < cost_old else 0
    rise_after = (new_cost / price - 1) * 100 if price < new_cost else 0
    failed = [g for g in gates if not g["ok"]]
    return {
        "name": name,
        "allowed": not failed,
        "gates": gates,
        "failed": [g["label"] for g in failed],
        "price": round(price, 3),
        "lots": qty // 100,
        "qty": qty,
        "amount": round(actual, 0),
        "qty_old": int(qty_old),
        "qty_new": int(new_qty),
        "cost_old": round(cost_old, 3),
        "cost_new": round(new_cost, 3),
        "cost_change": round(new_cost - cost_old, 3),
        "value_new": round(new_value, 0),
        "single_pct": round(single_pct, 1),
        "cash_after": round(cash_after, 0),
        "cash_after_pct": round(cash_after_pct, 1),
        "drop_10_loss": drop_loss(10),
        "drop_20_loss": drop_loss(20),
        "rise_before": round(rise_before, 1),
        "rise_after": round(rise_after, 1),
        "fundamentals_ok": bool(fundamentals_ok),
        "nav": trend,
    }


def snapshot(stocks=None):
    perf = performance()
    attr = attribution()
    exp = exposure()
    risk_rows = risks(perf=perf, attr=attr, exp=exp)
    snap = {
        "time": datetime.now(BEIJING_TZ).strftime("%Y-%m-%d %H:%M:%S"),
        "performance": perf,
        "attribution": attr,
        "exposure": exp,
        "risks": risk_rows,
        "actions": actions(risk_rows),
        "alerts": alerts(stocks or []),
        "health": health(perf, exp),
        "recovery": recovery(),
        "glossary": glossary(perf, exp),
        "discipline": discipline(perf, exp),
        "market": market_snapshot(),
        "structure": structure_advice(exp),
        "opportunity": opportunity_cost(),
    }
    snap["verdict"] = verdict(snap)
    return snap


# ============ 决策记录 ============
DECISION_HEADERS = ["日期", "类型", "标的", "参考价", "目标价", "理由", "状态", "复盘"]


def _decision_sheet(wb):
    if "决策记录" in wb.sheetnames:
        return wb["决策记录"]
    ws = wb.create_sheet("决策记录")
    ws.append(DECISION_HEADERS)
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 10
    ws.column_dimensions["F"].width = 40
    ws.column_dimensions["G"].width = 10
    ws.column_dimensions["H"].width = 40
    return ws


def load_decisions():
    try:
        wb = openpyxl.load_workbook(agent.EXCEL_PATH)
        ws = _decision_sheet(wb)
        rows = []
        for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not row[0]:
                continue
            rows.append({
                "row": i,
                "date": str(row[0]),
                "type": row[1] or "",
                "name": row[2] or "",
                "price": _num(row[3]),
                "target": _num(row[4]),
                "reason": row[5] or "",
                "status": row[6] or "待执行",
                "review": row[7] or "",
            })
        wb.close()
        rows.reverse()
        return rows
    except Exception as e:
        print(f"⚠️ 读取决策记录失败: {e}")
        return []


def add_decision(data):
    try:
        wb = openpyxl.load_workbook(agent.EXCEL_PATH)
        ws = _decision_sheet(wb)
        ws.append([
            datetime.now(BEIJING_TZ).strftime("%Y-%m-%d %H:%M"),
            data.get("type", "买入"),
            data.get("name", ""),
            _num(data.get("price")) or None,
            _num(data.get("target")) or None,
            data.get("reason", ""),
            data.get("status") or "待执行",
            data.get("review", ""),
        ])
        agent.save_workbook(wb)
        wb.close()
        return True
    except Exception as e:
        print(f"⚠️ 新增决策记录失败: {e}")
        return False


def update_decision(row, status=None, review=None):
    try:
        wb = openpyxl.load_workbook(agent.EXCEL_PATH)
        ws = _decision_sheet(wb)
        if status is not None:
            ws.cell(row=row, column=7, value=status)
        if review is not None:
            ws.cell(row=row, column=8, value=review)
        agent.save_workbook(wb)
        wb.close()
        return True
    except Exception as e:
        print(f"⚠️ 更新决策记录失败: {e}")
        return False


def delete_decision(row):
    try:
        wb = openpyxl.load_workbook(agent.EXCEL_PATH)
        ws = _decision_sheet(wb)
        ws.delete_rows(row)
        agent.save_workbook(wb)
        wb.close()
        return True
    except Exception as e:
        print(f"⚠️ 删除决策记录失败: {e}")
        return False
