#!/usr/bin/env python3
"""生成独立的 HTML 投资看板（Chart.js），在浏览器直接打开。"""
import json
from datetime import datetime
from pathlib import Path
import portfolio_agent as agent

DATA_DIR = Path(__file__).parent
OUTPUT_PATH = DATA_DIR / "dashboard.html"


def _load_data():
    holdings = agent.load_portfolio()
    total_value, total_pnl, accounts = agent.calculate(holdings)
    history = agent.load_history()
    by_type = {}
    for h in holdings:
        t = h["类别"]
        by_type[t] = by_type.get(t, 0) + h["市值"]
    return holdings, total_value, total_pnl, accounts, history, by_type


def _stock_rows(holdings):
    rows = ""
    for h in holdings:
        if h["类别"] != "股票":
            continue
        pnl = h["盈亏"] or 0
        pnl_cls = "up" if pnl > 0 else "down" if pnl < 0 else ""
        ret = h["收益率"] or 0
        ret_cls = "up" if ret > 0 else "down" if ret < 0 else ""
        alloc = h["仓位占比"] * 100 if h["仓位占比"] else 0
        rows += f"""<tr>
<td>{h['名称']}</td>
<td>{h['账户']}</td>
<td>{f'{h["数量"]:.0f}' if h.get('数量') else '-'}</td>
<td>{f'{h["现价"]:.2f}' if h.get('现价') else '-'}</td>
<td>{f'{h["成本价"]:.2f}' if h.get('成本价') else '-'}</td>
<td class="num">{h['市值']:,.0f}</td>
<td class="num {pnl_cls}">{pnl:+,.0f}</td>
<td class="num {ret_cls}">{ret*100:+.1f}%</td>
<td class="num">{alloc:.1f}%</td>
</tr>"""
    return rows


def _account_cards(accounts):
    cards = ""
    for acct, data in sorted(accounts.items()):
        pnl = data["盈亏"]
        cards += f"""<div class="account-card">
<div class="account-name">{acct}</div>
<div class="account-value">{data['市值']:,.0f}</div>
<div class="account-pnl {'up' if pnl > 0 else 'down'}">{pnl:+,.0f}</div>
</div>"""
    return cards


def _history_rows(history):
    rows = ""
    for r in history:
        rows += f"""<tr>
<td>{r['日期']}</td>
<td class="num">{r['总资产']:,.0f}</td>
<td class="num">{r['累计盈亏']:+,.0f}</td>
</tr>"""
    return rows


def generate():
    holdings, tv, tpnl, accounts, history, by_type = _load_data()
    cost_base = tv - tpnl
    ret_total = tpnl / cost_base * 100 if cost_base > 0 else 0

    # 序列化 JS 数据
    pie_labels = json.dumps(list(by_type.keys()), ensure_ascii=False)
    pie_data = json.dumps(list(by_type.values()))
    pie_colors = json.dumps(["#2F5496", "#FFC000", "#70AD47"])
    hist_dates = json.dumps([r["日期"] for r in history], ensure_ascii=False)
    hist_values = json.dumps([r["总资产"] for r in history])
    hist_pnl = json.dumps([r["累计盈亏"] for r in history])

    stock_count = sum(1 for h in holdings if h["类别"] == "股票" and h.get("数量"))

    # 读取模板
    template_path = DATA_DIR / "dashboard_template.html"
    html = template_path.read_text(encoding="utf-8")

    html = html.replace("{{NOW}}", datetime.now().strftime("%Y-%m-%d %H:%M"))
    html = html.replace("{{TOTAL_VALUE}}", f"{tv:,.0f}")
    html = html.replace("{{TOTAL_PNL}}", f"{tpnl:+,.0f}")
    html = html.replace("{{TOTAL_PNL_CLASS}}", "down" if tpnl < 0 else "up")
    html = html.replace("{{RET_TOTAL}}", f"{ret_total:+.2f}%")
    html = html.replace("{{RET_TOTAL_CLASS}}", "down" if ret_total < 0 else "up")
    html = html.replace("{{RET_TOTAL_LABEL}}", "亏损" if ret_total < 0 else "盈利")
    html = html.replace("{{STOCK_COUNT}}", str(stock_count))
    html = html.replace("{{ACCOUNT_CARDS}}", _account_cards(accounts))
    html = html.replace("{{PIE_LABELS}}", pie_labels)
    html = html.replace("{{PIE_DATA}}", pie_data)
    html = html.replace("{{PIE_COLORS}}", pie_colors)
    html = html.replace("{{HIST_DATES}}", hist_dates)
    html = html.replace("{{HIST_VALUES}}", hist_values)
    html = html.replace("{{HIST_PNL}}", hist_pnl)
    html = html.replace("{{STOCK_ROWS}}", _stock_rows(holdings))

    # 注意替换顺序: HISTORY_BLOCK 要先于 HISTORY_ROWS
    if history:
        block = """<div class="chart-container" style="height:240px;">
    <canvas id="historyChart"></canvas>
</div>
<table style="margin-top:16px;">
    <thead><tr><th>日期</th><th class="num">总资产</th><th class="num">累计盈亏</th></tr></thead>
    <tbody>__HISTORY_ROWS__</tbody>
</table>"""
        html = html.replace("{{HISTORY_BLOCK}}", block)
        html = html.replace("__HISTORY_ROWS__", _history_rows(history))
        js_block = """new Chart(document.getElementById('historyChart'), {
type: 'line',
data: { labels: __HIST_DATES__, datasets: [{ label: '累计盈亏', data: __HIST_PNL__, borderColor: '#CC0000',
backgroundColor: 'rgba(204, 0, 0, 0.1)', fill: true, tension: 0.3, pointRadius: 4,
pointBackgroundColor: '#CC0000' }] },
options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } },
scales: { y: { ticks: { callback: function(v) { return v.toLocaleString(); } } } } }
});"""
        html = html.replace("{{HISTORY_JS}}", js_block)
        html = html.replace("__HIST_DATES__", hist_dates)
        html = html.replace("__HIST_PNL__", hist_pnl)
    else:
        html = html.replace("{{HISTORY_BLOCK}}", '<p style="color:#888;">暂无历史记录。运行 python3 portfolio_agent.py snapshot 添加。</p>')
        html = html.replace("{{HISTORY_JS}}", "")

    OUTPUT_PATH.write_text(html, encoding="utf-8")
    print(f"✅ 看板已生成: {OUTPUT_PATH}")
    print(f"   在浏览器中打开查看: file://{OUTPUT_PATH}")


if __name__ == "__main__":
    generate()
