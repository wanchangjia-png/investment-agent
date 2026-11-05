#!/usr/bin/env python3
"""按月回顾：主炒板块、个股贡献、月度收益与机会成本（行业 ETF 最优解）。

口径：
- 只统计交易流水中的「股票」类（含 ETF、可转债）；黄金不计入。
- 月度盈亏 = 期末市值 - 期初市值 - 买入 + 卖出 - 费用 + 分红/红利税。
- 股票用不复权收盘价，ETF 基准用复权收盘价。
- 收益率分母为当月日均持仓市值；2026-07 起另给净值口径的账户收益。
"""
import bisect
import json
import os
import threading
import time
import urllib.request
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor

import portfolio_agent as agent
import settings

TIMELINE_START = "2024-11"
CACHE_VERSION = 3        # 行情缓存格式
RESULT_VERSION = 8       # 计算结果格式（改动算法时递增）
PRICE_CACHE_NAME = "timeline_price_cache.json"
RESULT_CACHE_NAME = "timeline_cache.json"

# 交易流水代码 → 板块（覆盖 2024-10 以来的全部 91 个标的）
STOCK_SECTOR_MAP = {
    "512480": "半导体/芯片", "159327": "半导体设备",
    "002475": "消费电子",
    "000158": "AI/算力", "000628": "AI/算力", "002362": "AI/算力", "159819": "AI/算力",
    "300598": "软件/计算机", "300559": "软件/计算机", "601360": "软件/计算机",
    "603528": "软件/计算机", "603990": "医药/医疗",
    "002017": "通信", "002104": "通信", "002123": "通信", "300130": "通信", "603220": "通信",
    "000156": "传媒/游戏", "000892": "传媒/游戏", "002174": "传媒/游戏", "002181": "传媒/游戏",
    "002558": "传媒/游戏", "002602": "传媒/游戏", "002624": "传媒/游戏", "300058": "传媒/游戏",
    "300251": "传媒/游戏", "300459": "传媒/游戏", "600880": "传媒/游戏", "600986": "传媒/游戏",
    "002878": "传媒/游戏",
    "002050": "新能源车/锂电", "002466": "新能源车/锂电", "002594": "新能源车/锂电",
    "002805": "新能源车/锂电", "002850": "新能源车/锂电", "601689": "新能源车/锂电",
    "603659": "新能源车/锂电", "605378": "新能源车/锂电",
    "002865": "光伏/储能", "603806": "光伏/储能", "600089": "光伏/储能", "002006": "光伏/储能",
    "000768": "军工", "002297": "军工", "300065": "军工", "603678": "军工",
    "000858": "白酒/食品饮料", "002568": "白酒/食品饮料", "002582": "白酒/食品饮料",
    "600887": "白酒/食品饮料", "605499": "白酒/食品饮料",
    "000534": "医药/医疗", "600276": "医药/医疗", "603222": "医药/医疗",
    "000987": "券商/金融", "600901": "券商/金融", "600906": "券商/金融", "512000": "券商/金融",
    "600036": "银行/保险", "601838": "银行/保险",
    "002254": "钢铁/化工",
    "000601": "煤炭/石油/能源", "600900": "煤炭/石油/能源", "600688": "煤炭/石油/能源",
    "600938": "煤炭/石油/能源", "603619": "煤炭/石油/能源",
    "000619": "房地产/基建", "000560": "房地产/基建", "002146": "房地产/基建",
    "600383": "房地产/基建",
    "600660": "汽车/零部件", "600699": "汽车/零部件",
    "300024": "机器人/高端制造", "603666": "机器人/高端制造", "301023": "机器人/高端制造",
    "002526": "机器人/高端制造",
    "002697": "消费/商贸", "601933": "消费/商贸", "605136": "消费/商贸",
    "603877": "消费/商贸", "605299": "消费/商贸", "300651": "消费/商贸",
    "002769": "消费/商贸", "603657": "消费/商贸",
    "113708": "可转债/其他",
    "159205": "宽基/海外", "510760": "宽基/海外", "510980": "宽基/海外",
    "588000": "宽基/海外", "159941": "宽基/海外",
}

# 每个板块一个代表 ETF；宽基/可转债不参与「最强板块」评选
SECTOR_ETFS = [
    ("sh512480", "半导体ETF", "半导体/芯片"),
    ("sz159516", "半导体设备ETF", "半导体设备"),
    ("sz159732", "消费电子ETF", "消费电子"),
    ("sz159819", "人工智能ETF", "AI/算力"),
    ("sz159998", "计算机ETF", "软件/计算机"),
    ("sh515880", "通信ETF", "通信"),
    ("sh512980", "传媒ETF", "传媒/游戏"),
    ("sh515030", "新能源车ETF", "新能源车/锂电"),
    ("sh515790", "光伏ETF", "光伏/储能"),
    ("sh512660", "军工ETF", "军工"),
    ("sh512690", "酒ETF", "白酒/食品饮料"),
    ("sh512010", "医药ETF", "医药/医疗"),
    ("sh512000", "券商ETF", "券商/金融"),
    ("sh512800", "银行ETF", "银行/保险"),
    ("sh512400", "有色金属ETF", "有色/资源"),
    ("sz159930", "能源ETF", "煤炭/石油/能源"),
    ("sz159870", "化工ETF", "钢铁/化工"),
    ("sh512200", "房地产ETF", "房地产/基建"),
    ("sh516110", "汽车ETF", "汽车/零部件"),
    ("sh562500", "机器人ETF", "机器人/高端制造"),
    ("sz159928", "消费ETF", "消费/商贸"),
]
UNRANKED_SECTORS = {"宽基/海外", "可转债/其他", "其他/未分类"}

_cache_lock = threading.RLock()


# ============ 通用 ============
def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _today():
    return settings.now().strftime("%Y-%m-%d")


def _exchange_prefix(code):
    code = str(code)
    if code.startswith(("110", "111", "113")):
        return "sh" + code
    if code.startswith(("6", "5", "9")):
        return "sh" + code
    if code.startswith(("0", "1", "2", "3")):
        return "sz" + code
    return "sh" + code


def _atomic_write_json(path, payload):
    tmp = str(path) + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False)
    os.replace(tmp, str(path))


def _cache_paths():
    base = settings.EXCEL_PATH.parent
    return base / PRICE_CACHE_NAME, base / RESULT_CACHE_NAME


def _month_list(start, end):
    y, m = int(start[:4]), int(start[5:7])
    ey, em = int(end[:4]), int(end[5:7])
    out = []
    while (y, m) <= (ey, em):
        out.append(f"{y:04d}-{m:02d}")
        m += 1
        if m > 12:
            y, m = y + 1, 1
    return out


def _month_bounds(month):
    y, m = int(month[:4]), int(month[5:7])
    start = f"{y:04d}-{m:02d}-01"
    nxt = f"{y + 1:04d}-01-01" if m == 12 else f"{y:04d}-{m + 1:02d}-01"
    return start, nxt


# ============ 行情 ============
def fetch_series(symbol, adjust="", days=800):
    """腾讯日线，返回 [[日期, 收盘], ...]；失败返回 None。"""
    url = (f"https://web.ifzq.gtimg.cn/appstock/app/fqkline/get"
           f"?param={symbol},day,,,{days},{adjust}")
    for attempt in range(3):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            raw = urllib.request.urlopen(req, timeout=10).read().decode("utf-8")
            node = json.loads(raw)["data"][symbol]
            rows = node.get("qfqday") or node.get("day") or []
            return [[r[0], float(r[2])] for r in rows if len(r) >= 3]
        except Exception:
            time.sleep(0.6 * (attempt + 1))
    return None


def _load_price_cache():
    path, _ = _cache_paths()
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if data.get("version") == CACHE_VERSION:
            return data
    except Exception:
        pass
    return {"version": CACHE_VERSION, "date": "", "series": {}}


def load_series(force=False):
    """返回 {symbol: [[日期, 收盘], ...]}，股票不复权、ETF 复权，按日缓存。"""
    today = _today()
    with _cache_lock:
        cache = _load_price_cache()
    existing = cache.get("series", {})
    if not force and cache.get("date") == today and existing:
        return existing

    stock_codes = set()
    for t in agent.load_transactions():
        if t.get("category") == "股票" and t.get("code"):
            stock_codes.add(str(t["code"]))
    tasks = [(_exchange_prefix(code), "") for code in sorted(stock_codes)]
    tasks += [(symbol, "qfq") for symbol, _, _ in SECTOR_ETFS]
    tasks.append(("sh000300", ""))

    def work(task):
        symbol, adjust = task
        return symbol, fetch_series(symbol, adjust)

    fetched = dict(existing)
    failures = []
    with ThreadPoolExecutor(max_workers=6) as ex:
        for symbol, rows in ex.map(work, tasks):
            if rows:
                fetched[symbol] = rows
            else:
                failures.append(symbol)

    # 全部失败时保留旧缓存，避免一次网络抖动把行情缓存清空
    if fetched and not (len(failures) == len(tasks) and existing):
        with _cache_lock:
            _atomic_write_json(_cache_paths()[0], {
                "version": CACHE_VERSION,
                "date": today if not failures else cache.get("date", ""),
                "series": fetched,
            })
    return fetched


def _date_index(rows):
    return [r[0] for r in rows], [r[1] for r in rows]


def _asof(dates, closes, date):
    i = bisect.bisect_right(dates, date)
    return closes[i - 1] if i > 0 else None


# ============ 流水与持仓 ============
def _collect_txns(txns):
    """返回 (events, opening, meta)。

    events[code] = [(date, signed_qty, 流水行号, 结存数量)]
    opening[code] = 流水窗口开始前的期初数量（由首条结存反推）
    """
    by_code = defaultdict(list)
    meta = {}
    for t in txns:
        if t.get("category") != "股票" or not t.get("code"):
            continue
        if t.get("direction") not in ("买入", "卖出"):
            continue
        code = str(t["code"])
        signed = _num(t.get("qty")) * (1 if t["direction"] == "买入" else -1)
        by_code[code].append((str(t["date"])[:10], signed, int(t.get("row") or 0), t))
        meta.setdefault(code, {"name": t.get("name") or code,
                               "account": t.get("account") or ""})
    events, opening = {}, {}
    for code, rows in by_code.items():
        rows.sort(key=lambda x: (x[0], x[2]))
        first = rows[0][3]
        post = first.get("ledger_qty")
        # 流水自带的结存数量是权威值；个别行结存为 0 会反推出负数，按 0 处理
        opening[code] = max(0.0, _num(post) - rows[0][1]) if post is not None else 0.0
        events[code] = [(d, q, r, t.get("ledger_qty")) for d, q, r, t in rows]
    return events, opening, meta


def _collect_dividends():
    out = defaultdict(list)
    for a in agent.load_corporate_actions():
        if a.get("type") not in ("现金分红", "红利税") or not a.get("code"):
            continue
        out[str(a["code"])].append((str(a["date"])[:10], _num(a.get("amount"))))
    return out


def _nav_returns():
    """{month: 净值口径月收益率}，用快照首尾并扣除当月出入金。"""
    history = sorted(agent.load_history(), key=lambda r: r["日期"])
    flows = [(str(f["date"])[:10], (1 if f["type"] == "存入" else -1) * _num(f["amount"]))
             for f in agent.load_capital_flows()]
    out = {}
    for month in sorted({r["日期"][:7] for r in history}):
        start, nxt = _month_bounds(month)
        prev = [r for r in history if r["日期"] < start]
        inside = [r for r in history if start <= r["日期"] < nxt]
        if not prev or not inside:
            continue
        base = _num(prev[-1]["总资产"])
        end = _num(inside[-1]["总资产"])
        net_flow = sum(amt for d, amt in flows if start <= d < nxt)
        if base > 0:
            out[month] = round((end - base - net_flow) / base * 100, 2)
    return out


def _nav_year_returns():
    """{year: 净值口径区间收益}，覆盖范围为快照首日到最近一天。"""
    history = sorted(agent.load_history(), key=lambda r: r["日期"])
    flows = [(str(f["date"])[:10], (1 if f["type"] == "存入" else -1) * _num(f["amount"]))
             for f in agent.load_capital_flows()]
    by_year = defaultdict(list)
    for r in history:
        by_year[str(r["日期"])[:4]].append(r)
    out = {}
    for year, rows in by_year.items():
        earlier = [r for r in history if str(r["日期"])[:10] < f"{year}-01-01"]
        if not earlier and len(rows) < 2:
            continue
        base_row = earlier[-1] if earlier else rows[0]
        start_date, end_date = str(base_row["日期"])[:10], str(rows[-1]["日期"])[:10]
        base, end = _num(base_row["总资产"]), _num(rows[-1]["总资产"])
        net_flow = sum(amt for d, amt in flows if start_date < d <= end_date)
        if base > 0:
            out[year] = {
                "return_pct": round((end - base - net_flow) / base * 100, 2),
                "start": start_date, "end": end_date,
                "full_year": bool(earlier),
            }
    return out


# ============ 计算 ============
def _index_and_calendar(series, today):
    rows = series.get("sh000300") or next((v for v in series.values() if v), [])
    dates, closes = _date_index(rows)
    calendar = [d for d in dates if d <= today]
    return (dates, closes), calendar


def _etf_monthly_returns(series, months, calendar):
    rows_by_symbol = {s: _date_index(rows) for s, rows in series.items() if rows}
    out = {}
    for month in months:
        start, nxt = _month_bounds(month)
        days = [d for d in calendar if start <= d < nxt]
        if not days:
            continue
        prev_days = [d for d in calendar if d < days[0]]
        if not prev_days:
            continue
        prev_day, last_day = prev_days[-1], days[-1]
        ranked = {}
        for symbol, name, sector in SECTOR_ETFS:
            if sector in UNRANKED_SECTORS or symbol not in rows_by_symbol:
                continue
            dates, closes = rows_by_symbol[symbol]
            p0 = _asof(dates, closes, prev_day)
            p1 = _asof(dates, closes, last_day)
            if not p0 or not p1:
                continue
            ranked[sector] = {"sector": sector, "etf": symbol, "name": name,
                              "return_pct": round((p1 / p0 - 1) * 100, 2)}
        out[month] = ranked
    return out


def _year_best_etf(series, year, today):
    year = str(year)
    base = "2024-10-31" if year == "2024" else f"{int(year) - 1}-12-31"
    end = min(f"{year}-12-31", today)
    best = None
    for symbol, name, sector in SECTOR_ETFS:
        if sector in UNRANKED_SECTORS or not series.get(symbol):
            continue
        dates, closes = _date_index(series[symbol])
        p0 = _asof(dates, closes, base)
        if p0 is None:
            p0 = next((c for d, c in zip(dates, closes) if d >= f"{year}-01-01"), None)
        p1 = _asof(dates, closes, end)
        if not p0 or not p1:
            continue
        ret = (p1 / p0 - 1) * 100
        if best is None or ret > best["return_pct"]:
            best = {"sector": sector, "etf": symbol, "name": name,
                    "return_pct": round(ret, 2)}
    return best


def _build_all(series):
    txns = agent.load_transactions()
    events, opening, meta = _collect_txns(txns)
    dividends = _collect_dividends()
    nav_returns = _nav_returns()
    nav_years = _nav_year_returns()
    today = _today()
    last_month = today[:7]
    months = _month_list(TIMELINE_START, last_month)

    (idx_dates, _idx_closes), calendar = _index_and_calendar(series, today)
    if not calendar:
        raise RuntimeError("行情日历获取失败")
    price_idx = {}
    for code in meta:
        rows = series.get(_exchange_prefix(code))
        if rows:
            price_idx[code] = _date_index(rows)
    etf_returns = _etf_monthly_returns(series, months, calendar)

    txns_by_month = defaultdict(list)
    for t in txns:
        if t.get("category") == "股票" and t.get("code") and t.get("direction") in ("买入", "卖出"):
            txns_by_month[str(t["date"])[:7]].append(t)

    shares = {code: q for code, q in opening.items() if abs(q) > 1e-9}
    sorted_events = {code: sorted(rows, key=lambda x: (x[0], x[2]))
                     for code, rows in events.items()}
    events_by_day = defaultdict(list)
    for code, rows in sorted_events.items():
        for d, qty, _row, _post in rows:
            events_by_day[d].append((code, qty))

    first_start, _ = _month_bounds(months[0])
    for code, rows in sorted_events.items():
        for d, signed, _row, _post in rows:
            if d >= first_start:
                break
            shares[code] = shares.get(code, 0.0) + signed

    warnings = []
    missing_prices = set()
    month_rows = []
    for month in months:
        start, nxt = _month_bounds(month)
        days = [d for d in calendar if start <= d < nxt]
        if not days:
            continue
        is_partial = month == last_month
        first_day, last_day = days[0], days[-1]
        prev_days = [d for d in idx_dates if d < first_day]
        prev_day = prev_days[-1] if prev_days else None

        q0 = dict(shares)
        buy_by_code, sell_by_code, fee_by_code = defaultdict(float), defaultdict(float), defaultdict(float)
        div_by_code = defaultdict(float)
        month_txns = txns_by_month.get(month, [])
        for t in month_txns:
            code = str(t["code"])
            amt, fee = _num(t.get("amount")), _num(t.get("fee"))
            if t["direction"] == "买入":
                buy_by_code[code] += amt
            else:
                sell_by_code[code] += amt
            fee_by_code[code] += fee
        for code, rows in dividends.items():
            for d, amt in rows:
                if start <= d < nxt:
                    div_by_code[code] += amt

        mv_sum = defaultdict(float)
        for day in days:
            for code, signed in events_by_day.get(day, []):
                shares[code] = shares.get(code, 0.0) + signed
            for code, qty in shares.items():
                if abs(qty) < 1e-9:
                    continue
                idx = price_idx.get(code)
                if not idx:
                    continue
                px = _asof(idx[0], idx[1], day)
                if px is not None:
                    mv_sum[code] += qty * px
        day_count = len(days)
        avg_mv_total = sum(mv_sum.values()) / day_count if day_count else 0.0
        q1 = dict(shares)

        pnl, ret = {}, {}
        codes = set(q0) | set(q1) | set(buy_by_code) | set(sell_by_code)
        for code in codes:
            qty0, qty1 = q0.get(code, 0.0), q1.get(code, 0.0)
            if abs(qty0) < 1e-9 and abs(qty1) < 1e-9 and not buy_by_code.get(code) and not sell_by_code.get(code):
                continue
            idx = price_idx.get(code)
            p0 = _asof(idx[0], idx[1], prev_day) if (idx and prev_day) else None
            p1 = _asof(idx[0], idx[1], last_day) if idx else None
            if abs(qty0) > 1e-9 and p0 is None:
                missing_prices.add(f"{meta.get(code, {}).get('name', code)}（{month}）")
                continue
            if abs(qty1) > 1e-9 and p1 is None:
                missing_prices.add(f"{meta.get(code, {}).get('name', code)}（{month}）")
                continue
            pnl[code] = (qty1 * (p1 or 0.0) - qty0 * (p0 or 0.0)
                         - buy_by_code.get(code, 0.0) + sell_by_code.get(code, 0.0)
                         - fee_by_code.get(code, 0.0) + div_by_code.get(code, 0.0))
            avg_code_mv = mv_sum.get(code, 0.0) / day_count if day_count else 0.0
            ret[code] = round(pnl[code] / avg_code_mv * 100, 2) if avg_code_mv > 0 else None

        total_pnl = round(sum(pnl.values()), 2)
        my_return = round(total_pnl / avg_mv_total * 100, 2) if avg_mv_total > 0 else None

        sector_pnl, sector_turnover, sector_end_mv = defaultdict(float), defaultdict(float), defaultdict(float)
        for code, v in pnl.items():
            sector_pnl[STOCK_SECTOR_MAP.get(code, "其他/未分类")] += v
        for code in set(buy_by_code) | set(sell_by_code):
            sector_turnover[STOCK_SECTOR_MAP.get(code, "其他/未分类")] += \
                buy_by_code.get(code, 0.0) + sell_by_code.get(code, 0.0)
        for code, qty in q1.items():
            if abs(qty) < 1e-9:
                continue
            idx = price_idx.get(code)
            px = _asof(idx[0], idx[1], last_day) if idx else None
            if px is not None:
                sector_end_mv[STOCK_SECTOR_MAP.get(code, "其他/未分类")] += qty * px

        sectors = [{
            "sector": name,
            "pnl": round(sector_pnl.get(name, 0.0), 2),
            "turnover": round(sector_turnover.get(name, 0.0), 2),
            "value": round(sector_end_mv.get(name, 0.0), 2),
        } for name in set(sector_pnl) | set(sector_turnover) | set(sector_end_mv)]
        by_turnover = sorted(sectors, key=lambda s: s["turnover"], reverse=True)
        main_sector = next((s for s in by_turnover if s["turnover"] > 0), None)
        if main_sector is None:
            by_value = sorted(sectors, key=lambda s: s["value"], reverse=True)
            if by_value and by_value[0]["value"] > 0:
                main_sector = dict(by_value[0], turnover=0.0)

        stocks = [{
            "code": code,
            "name": meta.get(code, {}).get("name", code),
            "sector": STOCK_SECTOR_MAP.get(code, "其他/未分类"),
            "pnl": round(v, 2),
            "return_pct": ret.get(code),
        } for code, v in pnl.items() if abs(v) >= 0.005]
        contributors = sorted([s for s in stocks if s["pnl"] > 0], key=lambda s: s["pnl"], reverse=True)[:5]
        drags = sorted([s for s in stocks if s["pnl"] < 0], key=lambda s: s["pnl"])[:5]

        ranked = sorted(etf_returns.get(month, {}).values(), key=lambda x: x["return_pct"], reverse=True)
        opportunity = None
        if ranked:
            best = ranked[0]
            gap = round(best["return_pct"] - my_return, 2) if my_return is not None else None
            cost = (round(avg_mv_total * (best["return_pct"] - my_return) / 100, 2)
                    if my_return is not None else None)
            opportunity = {
                "best_sector": best["sector"],
                "best_etf": f"{best['etf']} {best['name']}",
                "best_return_pct": best["return_pct"],
                "gap_pct": gap,
                "cost_rmb": cost,
                "top3": ranked[:3],
                "coverage": len(ranked),
            }

        month_rows.append({
            "month": month,
            "label": f"{month[:4]}年{int(month[5:7])}月",
            "partial": is_partial,
            "pnl": total_pnl,
            "return_pct": my_return,
            "return_basis": "holdings",
            "nav_return_pct": nav_returns.get(month),
            "nav_available": month in nav_returns,
            "avg_stock_value": round(avg_mv_total, 2),
            "turnover": round(sum(buy_by_code.values()) + sum(sell_by_code.values()), 2),
            "capital": {
                "buys": round(sum(buy_by_code.values()), 2),
                "sells": round(sum(sell_by_code.values()), 2),
                "fees": round(sum(fee_by_code.values()), 2),
                "dividends": round(sum(div_by_code.values()), 2),
                "trades": len(month_txns),
            },
            "main_sector": main_sector,
            "sectors": by_turnover[:8],
            "contributors": contributors,
            "drags": drags,
            "opportunity": opportunity,
        })

    try:
        live = {(h.get("代码") or ""): _num(h.get("数量"))
                for h in agent.load_portfolio()
                if h.get("类别") == "股票" and h.get("账户") == "A股主账户"}
        if any(abs(qty - shares.get(code, 0.0)) > 1e-6 for code, qty in live.items()):
            warnings.append("当前持仓与流水结存有差异（本月成交可能还没进对账单明细，T+1），"
                            "本月收益暂以流水为准，下次导入对账单后自动补齐。")
    except Exception:
        pass
    if missing_prices:
        warnings.append(f"{len(missing_prices)} 个标的缺当月行情，已从收益中剔除："
                        + "、".join(sorted(missing_prices)[:5]))

    years = []
    for year in sorted({m["month"][:4] for m in month_rows}, reverse=True):
        months_in_year = [m for m in month_rows if m["month"][:4] == year]
        year_pnl = round(sum(m["pnl"] for m in months_in_year), 2)
        compound, has_ret = 1.0, False
        for m in months_in_year:
            if m["return_pct"] is not None:
                has_ret = True
                compound *= (1 + m["return_pct"] / 100)
        best = _year_best_etf(series, year, today)
        nav = nav_years.get(year)
        years.append({
            "year": int(year),
            "summary": {
                "pnl": year_pnl,
                "return_pct": round((compound - 1) * 100, 2) if has_ret else None,
                "best_sector": best["sector"] if best else None,
                "best_return_pct": best["return_pct"] if best else None,
                "months": len(months_in_year),
                "nav_return_pct": nav["return_pct"] if nav else None,
                "nav_start": nav["start"] if nav else None,
                "nav_full_year": nav["full_year"] if nav else False,
            },
            "months": list(reversed(months_in_year)),
        })

    return {
        "generated_at": settings.now().strftime("%Y-%m-%d %H:%M:%S"),
        "range": {"start": TIMELINE_START, "end": last_month,
                  "partial": last_month == today[:7]},
        "years": years,
        "warnings": warnings,
        "meta": {"etf_count": len(SECTOR_ETFS), "ledger_sig": _ledger_sig()},
    }


def _ledger_sig():
    try:
        st = settings.EXCEL_PATH.stat()
        return [st.st_mtime_ns, st.st_size]
    except OSError:
        return None


def build_timeline(force=False):
    """对外入口：签名不变时直接读结果缓存。"""
    _, result_path = _cache_paths()
    sig = _ledger_sig()
    today = _today()
    if not force:
        try:
            with open(result_path, "r", encoding="utf-8") as f:
                cached = json.load(f)
            if (cached.get("version") == RESULT_VERSION
                    and cached.get("ledger_sig") == sig
                    and cached.get("prices_date") == today):
                data = cached.get("data") or {}
                data["cache"] = {"hit": True, "prices_date": today}
                return data
        except Exception:
            pass

    with _cache_lock:
        # force 只跳过结果缓存；行情仍按自然日缓存，避免一次刷新等上百个请求
        series = load_series(force=False)
        data = _build_all(series)
        _atomic_write_json(result_path, {
            "version": RESULT_VERSION, "ledger_sig": sig,
            "prices_date": today, "data": data,
        })
    data["cache"] = {"hit": False, "prices_date": today}
    return data


if __name__ == "__main__":
    import sys
    t0 = time.time()
    result = build_timeline(force="--force" in sys.argv)
    print(f"耗时 {time.time() - t0:.1f}s  月份 {sum(len(y['months']) for y in result['years'])}")
    for y in result["years"][:2]:
        print(y["year"], y["summary"])
        for m in y["months"][:3]:
            opp = m["opportunity"] or {}
            print(m["month"], m["pnl"], m["return_pct"], m["nav_return_pct"],
                  (m["main_sector"] or {}).get("sector"),
                  opp.get("best_sector"), opp.get("gap_pct"))
    if result.get("warnings"):
        print("warnings:", result["warnings"])
