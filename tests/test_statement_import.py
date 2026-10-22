import openpyxl

import statement_import as si


def test_query_range():
    assert si._query_range({"query_range": "2026/09/22-2026/10/22"}) == (
        "2026-09-22", "2026-10-22")
    assert si._query_range({}) == (None, None)


def test_merge_opening_prefers_ledger():
    reverse = {"002475": {"name": "立讯精密", "qty": 200}}
    ledger = {"002475": [100, 53.15]}
    out = si._merge_opening(reverse, ledger, {"002475": 105.509})
    assert out["002475"] == {"qty": 100, "cost": 53.15}


def test_merge_opening_falls_back_to_statement_cost():
    out = si._merge_opening({"510760": {"name": "上证指数ETF", "qty": 7500}}, {},
                            {"510760": 1.267})
    assert out["510760"] == {"qty": 7500, "cost": 1.267}


def test_upsert_keeps_rows_outside_range():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "交易流水"
    ws.append(si.LEDGER_HEADERS)
    old = ["2026-07-24", "买入", "股票", "三花智控", "002050", 100, 35, 3500,
           "A股主账户", "对账单导入", 5, None, 35, 100, 35]
    window = ["2026-09-24", "买入", "股票", "立讯精密", "002475", 100, 52.08, 5208,
              "A股主账户", "对账单导入", 5, None, 0, 100, 52.13]
    stale = window[:]
    stale[5] = 999
    ws.append(old)
    ws.append(stale)

    si._upsert_transactions(ws, [window], "2026-09-22", "2026-10-22")

    rows = list(ws.iter_rows(min_row=2, values_only=True))
    assert len(rows) == 2
    assert rows[0][0] == "2026-07-24"
    assert rows[1][0] == "2026-09-24"
    assert rows[1][5] == 100


def test_ledger_opening_state(tmp_path):
    path = tmp_path / "book.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "交易流水"
    ws.append(si.LEDGER_HEADERS)
    ws.append(["2026-09-17", "卖出", "股票", "三花智控", "002050", 100, 34.5, 3450,
               "A股主账户", "对账单导入", 6, -320, 37.6, 800, 37.638])
    ws.append(["2026-10-19", "卖出", "股票", "三花智控", "002050", 100, 31.92, 3192,
               "A股主账户", "对账单导入", 6, -578, 37.638, 700, 37.638])
    wb.save(path)

    state = si._ledger_opening_state(str(path), "2026-09-22")
    assert state["002050"] == [800, 37.638]
