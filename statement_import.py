#!/usr/bin/env python3
"""东方财富证券「股票明细对账单」解析与导入。

PDF 文本提取会丢失列间距，这里按字符坐标还原表格。
"""
import re
import json
from datetime import datetime

import PyPDF2

# 列区间（按 PDF 坐标，左闭右开）
COLS = [
    ("date", 20, 62),
    ("type", 62, 132),
    ("code", 132, 175),
    ("name", 175, 262),
    ("qty", 262, 290),
    ("price", 290, 338),
    ("amount", 338, 396),
    ("fee", 396, 442),
    ("stamp", 442, 481),
    ("transfer", 481, 516),
    ("balance", 516, 620),
]

SUMMARY_COLS = [
    ("market", 35, 112),
    ("code", 112, 190),
    ("name", 190, 272),
    ("qty", 272, 352),
    ("price", 352, 422),
    ("cost", 422, 502),
    ("value", 502, 620),
]

TRADE_TYPES = {"证券买入", "证券卖出"}
DIV_TYPES = {"红利入账", "股息红利差异扣税"}


def _num(text):
    s = str(text or "").strip().replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _lines(page, page_no):
    items = []

    def visitor(text, cm, tm, font_dict, font_size):
        t = (text or "").strip()
        if t:
            items.append({"page": page_no, "y": round(tm[5], 1), "x": round(tm[4], 1), "text": t})

    page.extract_text(visitor_text=visitor)
    items.sort(key=lambda i: (-i["y"], i["x"]))
    grouped = []
    for it in items:
        if grouped and abs(grouped[-1]["y"] - it["y"]) <= 1.2:
            grouped[-1]["items"].append(it)
        else:
            grouped.append({"page": page_no, "y": it["y"], "items": [it]})
    return grouped


def _cells(line, cols):
    out = {name: "" for name, _, _ in cols}
    for it in line["items"]:
        x = it["x"]
        for name, lo, hi in cols:
            if lo <= x < hi:
                prev = out[name]
                out[name] = (prev + " " + it["text"]).strip() if prev else it["text"]
                break
    return out


def parse_pdf(path, password):
    reader = PyPDF2.PdfReader(path)
    if reader.is_encrypted:
        if not reader.decrypt(password):
            raise ValueError("PDF 密码错误")
    meta = {"query_range": "", "printed_at": "", "customer": "", "account": ""}
    summary, trades = [], []
    for page_no, page in enumerate(reader.pages, 1):
        for line in _lines(page, page_no):
            text = " ".join(i["text"] for i in line["items"])
            if text.startswith("查询区间"):
                meta["query_range"] = text.split("：", 1)[-1]
            elif text.startswith("打印日期"):
                meta["printed_at"] = text.split("：", 1)[-1]
            elif "客户姓名" in text:
                meta["customer"] = text.split("客户姓名：", 1)[-1].split()[0]
            elif text.startswith("资金账号"):
                meta["account"] = text.split("：", 1)[-1].split()[0]
            elif text.startswith("总资产(RMB)"):
                for it in line["items"]:
                    if it["x"] > 90 and re.match(r"^[\d.]+$", it["text"]):
                        meta["total_assets"] = _num(it["text"])
                        break
            cells = _cells(line, COLS)
            if cells["date"].isdigit() and len(cells["date"]) == 8:
                rtype = cells["type"]
                qty = _num(cells["qty"])
                price = _num(cells["price"])
                amount = _num(cells["amount"])
                rec = {
                    "date": f"{cells['date'][:4]}-{cells['date'][4:6]}-{cells['date'][6:]}",
                    "type": rtype, "code": cells["code"].strip(), "name": cells["name"].strip(),
                    "qty": qty, "price": price, "amount": amount,
                    "fee": _num(cells["fee"]), "stamp": _num(cells["stamp"]),
                    "transfer": _num(cells["transfer"]), "balance": _num(cells["balance"]),
                    "page": page_no,
                }
                trades.append(rec)
                continue
            sc = _cells(line, SUMMARY_COLS)
            m = re.match(r"^(\d{6})\s*(.*)$", sc["code"].strip())
            if m and (sc["name"].strip() or m.group(2).strip()):
                name = (sc["name"].strip() or m.group(2)).replace(" ", "")
                summary.append({
                    "market": sc["market"].strip(), "code": m.group(1),
                    "name": name, "qty": _num(sc["qty"]),
                    "price": _num(sc["price"]), "cost": _num(sc["cost"]),
                    "value": _num(sc["value"]),
                })
    # 去重（同一行可能被重复解析）
    seen, uniq = set(), []
    for t in trades:
        key = (t["date"], t["type"], t["code"], t["name"], t["qty"], t["price"], t["amount"], t["balance"])
        if key not in seen:
            seen.add(key)
            uniq.append(t)
    trades = sorted(uniq, key=lambda x: (x["date"], x["page"]), reverse=True)
    return {"meta": meta, "summary": summary, "trades": trades}


def opening_positions(trades, summary):
    """用期末持仓 - 区间买卖，反推 7/21 收盘持仓"""
    qty = {s["code"]: s["qty"] for s in summary}
    names = {s["code"]: s["name"] for s in summary}
    for t in trades:
        code = t["code"]
        if t["type"] == "证券买入":
            qty[code] = qty.get(code, 0) - t["qty"]
            names.setdefault(code, t["name"])
        elif t["type"] == "证券卖出":
            qty[code] = qty.get(code, 0) + t["qty"]
            names.setdefault(code, t["name"])
        elif t["type"] == "新股入账":
            qty[code] = qty.get(code, 0) - t["qty"]
            names.setdefault(code, t["name"])
    return {c: {"name": names.get(c, ""), "qty": q} for c, q in qty.items() if abs(q) > 1e-9}


# ---------- 导入 Excel ----------

# 7/21 收盘持仓的真实成本（无历史流水时的起点假设，来自持仓表与陈述成本反推）
OPENING_COST = {
    "002050": 38.9149, "002475": 62.026, "159327": 1.1655, "601689": 56.5577,
    "002466": 65.404, "603678": 84.246, "000858": 258.743,
}
APP_NAMES = {
    "159327": "半导体设备", "512480": "半导体ETF", "588000": "科创50",
    "605499": "东鹏饮料",
}
SEC_TYPES = {"证券买入", "证券卖出", "新股入账"}
STMT_SOURCE = "对账单 2026/07/22-2026/10/22"


def _app_name(code, name):
    return APP_NAMES.get(code) or name.replace(" ", "")


def build_ledger(data):
    """按日期滚动真实成本，返回可直接写入「交易流水」的行。"""
    opening = opening_positions(data["trades"], data["summary"])
    state, last_cost = {}, {}
    for code, info in opening.items():
        q = info["qty"]
        c = OPENING_COST.get(code, 0.0)
        state[code] = [q, c]
        last_cost[code] = c

    trades = [t for t in data["trades"] if t["type"] in SEC_TYPES]
    trades.sort(key=lambda t: (t["date"], t["page"]))
    rows, warnings = [], []
    for t in trades:
        code, name = t["code"], _app_name(t["code"], t["name"])
        qty, price = t["qty"], t["price"]
        q, cost = state.get(code, [0.0, OPENING_COST.get(code, 0.0)])
        if code not in state:
            state[code] = [q, cost]
            last_cost[code] = cost
        gross = round(qty * price, 2)
        fee = round(t["fee"] + t["stamp"] + t["transfer"], 2)
        note = "对账单导入"

        if t["type"] == "证券买入":
            direction, realized = "买入", None
            if q < 0:  # 回补此前的异常卖空
                nq = q + qty
                if nq < 0:
                    nc = cost
                elif nq == 0:
                    nc = 0.0
                else:
                    nc = round((gross + fee) / nq, 3)
            else:
                nq = q + qty
                nc = round((q * cost + gross + fee) / nq, 3) if nq else 0.0
        elif t["type"] == "新股入账":
            direction, realized = "买入", None
            gross = 0.0
            note = "对账单导入｜新股入账（可转债按面值 100 元/张计成本）"
            nq = q + qty
            nc = round((q * cost + 1000.0) / nq, 3) if nq else 0.0
        else:
            direction = "卖出"
            basis = last_cost.get(code, cost) if q <= 0 else cost
            realized = round((price - basis) * qty - fee, 2)
            nq = q - qty
            if q <= 0:
                warnings.append(f"{t['date']} {name} 卖出 {qty:g} 股时结存不足，对账单疑似漏记买入")
                note = "对账单导入（⚠️ 卖出前结存不足，疑似漏记买入）"
                nc = basis
            else:
                nc = cost if nq > 0 else 0.0

        last_cost[code] = cost if q > 0 else last_cost.get(code, cost)
        state[code] = [nq, nc]
        rows.append([
            t["date"], direction, "股票", name, code, qty, price, gross,
            "A股主账户", note, fee, realized, round(cost, 3), round(nq, 4), round(nc, 3),
        ])
    return rows, warnings, state


LEDGER_HEADERS = ["日期", "方向", "资产类别", "资产名称", "代码", "成交数量",
                  "成交价格", "成交金额", "账户", "备注", "费用", "实现盈亏",
                  "成本价(复盘)", "结存数量", "结存成本"]


def _upsert_transactions(ws, rows):
    # 旧版占位行是 A2:J2 合并单元格，不清掉会把首条流水写坏
    for rng in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rng))
    # 旧表头只有 A–J，补齐复盘列
    for i, header in enumerate(LEDGER_HEADERS, 1):
        ws.cell(1, i, header)
    # 清掉上次同一来源的导入，保证可重复执行
    for r in range(ws.max_row, 1, -1):
        note = str(ws.cell(r, 10).value or "")
        first = ws.cell(r, 1).value
        malformed = (ws.cell(r, 2).value is None and ws.cell(r, 14).value is not None)
        if (note.startswith("对账单导入") or "新股入账" in note
                or (isinstance(first, str) and first.startswith("（")) or malformed):
            ws.delete_rows(r)
    for row in rows:
        ws.append(row)


def _write_corporate_actions(wb, data):
    name = "公司行为"
    if name in wb.sheetnames:
        ws = wb[name]
        for r in range(ws.max_row, 1, -1):
            if ws.cell(r, 9).value == STMT_SOURCE:
                ws.delete_rows(r)
    else:
        ws = wb.create_sheet(name)
        ws.append(["日期", "代码", "名称", "类型", "金额", "数量", "每股金额", "说明", "来源"])
        for col, width in {"A": 12, "B": 10, "C": 14, "D": 14, "E": 12,
                           "F": 10, "G": 10, "H": 28, "I": 26}.items():
            ws.column_dimensions[col].width = width
    type_map = {
        "红利入账": "现金分红", "股息红利差异扣税": "红利税",
        "利息归本": "利息", "OTC现金宝调账入": "现金调账",
    }
    rows = []
    for t in data["trades"]:
        if t["type"] not in type_map:
            continue
        rows.append([
            t["date"], t["code"], _app_name(t["code"], t["name"]),
            type_map[t["type"]], t["amount"], t["qty"], t["price"],
            "对账单导入", STMT_SOURCE,
        ])
    rows.sort(key=lambda r: r[0])
    for row in rows:
        ws.append(row)
    return len(rows)


def _update_holdings(wb, data):
    ws = wb["持仓表"]
    summary = {s["code"]: s for s in data["summary"]}
    cash_total = round(data["meta"].get("total_assets", 0) - sum(s["value"] for s in data["summary"]), 2)

    def row_of(code):
        for r in range(2, ws.max_row + 1):
            if str(ws.cell(r, 4).value or "") == code and ws.cell(r, 2).value == "A股主账户":
                return r
        return None

    for code, s in summary.items():
        r = row_of(code)
        if r is None:
            continue
        ws.cell(r, 3, _app_name(code, s["name"]))
        ws.cell(r, 5, s["qty"])
        ws.cell(r, 6, s["cost"])
        ws.cell(r, 7, s["price"])
        pnl = round((s["price"] - s["cost"]) * s["qty"], 2) if s["cost"] else 0.0
        ws.cell(r, 8, s["value"])
        ws.cell(r, 9, pnl)
        ws.cell(r, 10, round(pnl / (s["cost"] * s["qty"]), 4) if s["cost"] and s["qty"] else 0)
        if code == "002475":
            ws.cell(r, 12, "对账单：已清仓（10/19 卖出最后 100 股）")

    # 账户一现金：总资产 - 证券市值
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 1).value == "现金" and ws.cell(r, 2).value == "A股主账户":
            ws.cell(r, 5, cash_total)
            ws.cell(r, 8, cash_total)
            ws.cell(r, 12, "对账单 10/22：资金可用 50,268.80 + 余额 0.10（44.06 中含今日分红 43.96 未计总资产）")
            break

    # 仓位占比仍按账户一股票市值合计
    stock_total = sum(s["value"] for s in data["summary"])
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 1).value == "股票" and ws.cell(r, 2).value == "A股主账户":
            val = ws.cell(r, 8).value or 0
            ws.cell(r, 11, round(val / stock_total, 4) if stock_total else 0)

    # 合计行
    mv = pnl = 0.0
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 1).value == "合计":
            break
        mv += ws.cell(r, 8).value or 0
        pnl += ws.cell(r, 9).value or 0
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 1).value == "合计":
            ws.cell(r, 8, round(mv, 2))
            ws.cell(r, 9, round(pnl, 2))
            break
    return cash_total, stock_total


def import_statement(pdf_path, password, excel_path="portfolio.xlsx", dry_run=False):
    import openpyxl
    data = parse_pdf(pdf_path, password)
    ledger, warnings, state = build_ledger(data)
    if dry_run:
        return {"meta": data["meta"], "ledger": ledger, "warnings": warnings, "state": state}

    wb = openpyxl.load_workbook(excel_path)
    ws = wb["交易流水"]
    _upsert_transactions(ws, ledger)
    actions = _write_corporate_actions(wb, data)
    cash_total, stock_total = _update_holdings(wb, data)
    wb.save(excel_path)
    wb.close()
    return {
        "meta": data["meta"],
        "ledger_rows": len(ledger), "corporate_actions": actions,
        "cash": cash_total, "stock_value": stock_total,
        "warnings": warnings, "final": state,
    }


if __name__ == "__main__":
    import sys
    if len(sys.argv) >= 4 and sys.argv[1] == "import":
        result = import_statement(sys.argv[2], sys.argv[3], sys.argv[4] if len(sys.argv) > 4 else "portfolio.xlsx")
        print(json.dumps({k: v for k, v in result.items() if k != "final"}, ensure_ascii=False, indent=2))
        for code, (q, c) in result["final"].items():
            if abs(q) > 1e-9:
                print(" ", code, q, round(c, 3))
    else:
        data = parse_pdf(sys.argv[1], sys.argv[2])
        print(json.dumps(data["meta"], ensure_ascii=False, indent=2))
        print("summary rows:", len(data["summary"]))
        for s in data["summary"]:
            print(" ", s)
        print("trades:", len(data["trades"]))
        print("opening:", opening_positions(data["trades"], data["summary"]))
