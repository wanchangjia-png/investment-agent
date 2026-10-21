#!/usr/bin/env python3
"""持仓深度分析：基准对比、时间加权收益、回撤、归因、行业暴露、风险与提醒。"""
import json
import math
import threading
import urllib.request
from collections import defaultdict
from datetime import datetime, timedelta, timezone

import openpyxl

import portfolio_agent as agent

BEIJING_TZ = timezone(timedelta(hours=8))
BENCHMARK_CODE = "sh000300"
BENCHMARK_NAME = "沪深300"

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
    for r in agent.get_realized_pnl_records():
        realized[r["name"]] += _num(r["pnl"])

    items = {}
    for h in holdings:
        if h["类别"] != "股票":
            continue
        name = h["名称"]
        items.setdefault(name, {"name": name, "sector": agent.SECTOR_MAP.get(name, "其他"),
                                "unrealized": 0.0, "realized": 0.0, "market_value": 0.0})
        items[name]["unrealized"] += _num(h["盈亏"])
        items[name]["market_value"] += _num(h["市值"])
    for name, pnl in realized.items():
        items.setdefault(name, {"name": name, "sector": agent.SECTOR_MAP.get(name, "其他"),
                                "unrealized": 0.0, "realized": 0.0, "market_value": 0.0})
        items[name]["realized"] += pnl

    gold_unrealized = sum(_num(h["盈亏"]) for h in holdings if h["类别"] == "黄金")
    rows = []
    for it in items.values():
        total = it["unrealized"] + it["realized"]
        if abs(total) < 0.01:
            continue
        rows.append({
            "name": it["name"],
            "sector": it["sector"],
            "unrealized": round(it["unrealized"], 2),
            "realized": round(it["realized"], 2),
            "pnl": round(total, 2),
            "market_value": round(it["market_value"], 2),
        })
    if abs(gold_unrealized) >= 0.01:
        rows.append({"name": "黄金", "sector": "黄金", "unrealized": round(gold_unrealized, 2),
                     "realized": 0.0, "pnl": round(gold_unrealized, 2),
                     "market_value": round(sum(_num(h["市值"]) for h in holdings if h["类别"] == "黄金"), 2)})

    total = sum(r["pnl"] for r in rows)
    denom = abs(total) if abs(total) > 1 else 1
    for r in rows:
        r["share"] = round(r["pnl"] / denom * 100, 1)
    rows.sort(key=lambda r: r["pnl"])
    return {"items": rows, "total": round(total, 2)}


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
        [{"name": h["名称"], "value": _num(h["市值"]), "pct_total": round(_num(h["市值"]) / total_value * 100, 1) if total_value else 0,
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


# ============ 风险规则 ============
def risks(perf=None, attr=None, exp=None):
    perf = perf or performance()
    exp = exp or exposure()
    rows = []

    def add(level, title, detail, action, metric=""):
        rows.append({"level": level, "title": title, "detail": detail, "action": action, "metric": metric})

    m = perf.get("metrics", {})
    if m.get("max_drawdown", 0) <= -12:
        add("高", "回撤超过 12%", f"自净值高点最大回撤 {m['max_drawdown']:.1f}%，已进入深度回撤区间。",
            "先控制仓位，暂停加仓，等净值企稳再恢复。", f"{m['max_drawdown']:.1f}%")
    elif m.get("max_drawdown", 0) <= -8:
        add("中", "回撤偏大", f"最大回撤 {m['max_drawdown']:.1f}%，超过 8% 的警戒线。",
            "检查亏损最大的两只持仓，设定反弹减仓价位。", f"{m['max_drawdown']:.1f}%")

    if m.get("excess_return", 0) <= -10:
        add("中", "明显跑输沪深300", f"同期账户 {m['account_return']:+.1f}%，沪深300 {m['benchmark_return']:+.1f}%，落后 {abs(m['excess_return']):.1f} 个百分点。",
            "对照基准持仓，考虑用宽基 ETF 替代持续跑输的个股。", f"{m['excess_return']:+.1f}pp")

    if exp["top1_pct"] > 20:
        add("高", "单票集中度偏高", f"第一大持仓占总资产 {exp['top1_pct']:.1f}%，超过 20% 上限。",
            "分批减到 15% 以内，避免单只股票决定整体盈亏。", f"{exp['top1_pct']:.1f}%")
    elif exp["top3_pct"] > 45:
        add("中", "头部持仓过重", f"前三大持仓合计占总资产 {exp['top3_pct']:.1f}%。",
            "把前三大持仓合计压到 40% 以下，增加行业分散。", f"{exp['top3_pct']:.1f}%")

    top_sector = exp["sectors"][0] if exp["sectors"] else None
    if top_sector and top_sector["pct_total"] > 30:
        add("高", "行业过度集中", f"{top_sector['sector']} 占总资产 {top_sector['pct_total']:.1f}%，单一行业风险过高。",
            "新增不同行业资产，把单一行业占比降到 25% 以下。", f"{top_sector['pct_total']:.1f}%")
    elif top_sector and top_sector["pct_stock"] > 45:
        add("中", "股票仓位行业集中", f"{top_sector['sector']} 占股票仓位 {top_sector['pct_stock']:.1f}%。",
            "股票部分补充 1-2 个弱相关行业。", f"{top_sector['pct_stock']:.1f}%")

    for p in exp["positions"]:
        if p["return_pct"] <= -30:
            add("高", f"{p['name']} 深度亏损", f"浮亏 {p['return_pct']:.1f}%，占总资产 {p['pct_total']:.1f}%。",
                "设定 +10% / +20% 两档反弹减仓价，不被动等回本。", f"{p['return_pct']:.1f}%")
        elif p["return_pct"] <= -20:
            add("中", f"{p['name']} 亏损扩大", f"浮亏 {p['return_pct']:.1f}%，占总资产 {p['pct_total']:.1f}%。",
                "检查基本面是否变化，设定止损或反弹减仓计划。", f"{p['return_pct']:.1f}%")

    if exp["cash_pct"] < 5:
        add("高", "现金缓冲不足", f"现金仅占总资产 {exp['cash_pct']:.1f}%，低于 5% 的应急底线。",
            "至少留出 10% 现金，避免被动卖出。", f"{exp['cash_pct']:.1f}%")
    elif exp["cash_pct"] < 10:
        add("中", "现金偏少", f"现金占总资产 {exp['cash_pct']:.1f}%，低于建议的 10%。",
            "后续反弹减仓时不要立刻回补，把现金补到 10%。", f"{exp['cash_pct']:.1f}%")
    elif exp["cash_pct"] > 25:
        add("低", "现金闲置偏多", f"现金占总资产 {exp['cash_pct']:.1f}%，超过 25%。",
            "分批配置宽基 ETF 或货币基金，别让资金长期空转。", f"{exp['cash_pct']:.1f}%")

    stock_count = len([p for p in exp["positions"] if p["value"] > 0])
    if stock_count < 4:
        add("中", "持仓过于集中", f"仅有 {stock_count} 只有效持仓。",
            "分散到 5-8 只、3-5 个行业，单只不超过 20%。", f"{stock_count} 只")
    elif stock_count > 10:
        add("低", "持仓过于分散", f"持有 {stock_count} 只股票，管理精力被摊薄。",
            "保留最有把握的 6-8 只，清理小仓位。", f"{stock_count} 只")

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
                            "text": f"距成本 {abs(gap):.1f}%，回本价 {cost:.2f}",
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
    """把数据压成一句话：今天该不该动手 + 为什么 + 触发条件"""
    alerts = snap.get("alerts") or []
    risk_rows = snap.get("risks") or []
    perf_m = (snap.get("performance") or {}).get("metrics", {})
    exp = snap.get("exposure") or {}
    rec = snap.get("recovery") or {}

    high_risks = [r for r in risk_rows if r["level"] == "高"]
    high_alerts = [a for a in alerts if a["level"] == "高"]

    if high_risks or high_alerts:
        headline = "今天：不加仓，先控风险"
        tone = "warning"
    elif risk_rows:
        headline = "今天：不开新仓，处理现有持仓"
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
    for a in high_alerts[:1] + [x for x in alerts if x["level"] == "中"][:1]:
        triggers.append(f"{a['name']}：{a['text']}，先看有没有基本面消息，不做情绪化补仓。")
    for p in rec.get("positions", [])[:2]:
        if p["rise_needed"] >= 15:
            triggers.append(
                f"{p['name']} 现价 {p['price']:.2f}，回本需涨 {p['rise_needed']:.1f}% 到 {p['cost']:.2f}；"
                f"反弹 10%（{p['price']*1.1:.2f}）可先减 1/3。"
            )
    if exp.get("cash_pct", 0) > 20:
        triggers.append(f"现金 {exp['cash']:,.0f} 元（{exp['cash_pct']:.1f}%）：回撤未收窄前不动，不抄底。")
    if not triggers:
        triggers.append("没有触发条件，今天不需要操作。")

    return {"headline": headline, "tone": tone, "reasons": reasons[:3], "triggers": triggers[:3]}


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
            "待执行",
            "",
        ])
        wb.save(agent.EXCEL_PATH)
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
        wb.save(agent.EXCEL_PATH)
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
        wb.save(agent.EXCEL_PATH)
        wb.close()
        return True
    except Exception as e:
        print(f"⚠️ 删除决策记录失败: {e}")
        return False
