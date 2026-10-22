#!/usr/bin/env python3
"""
投资追踪 Agent - Web 服务器
启动后浏览器打开 http://localhost:5000 即可使用

支持对话 AI（DeepSeek / OpenAI），联网搜索。
"""
import json
import os
import sys
import re
import time
import threading
import urllib.request
import urllib.parse
from datetime import datetime
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from pathlib import Path

# 确保能找到同目录模块
sys.path.insert(0, str(Path(__file__).parent))
import portfolio_agent as agent
import insights
import monitor

HOST = os.environ.get("HOST", "0.0.0.0")
PORT = int(os.environ.get("PORT", os.environ.get("RAILWAY_PORT", 5050)))
DATA_DIR = Path(__file__).parent
HTML_PATH = DATA_DIR / "webapp.html"

# ============ LLM 配置 ============
# 通过环境变量设置:
#   LLM_PROVIDER=deepseek|openai
#   LLM_API_KEY=sk-xxxxx
#   LLM_MODEL=deepseek-chat|gpt-4o
# 也可以在网页 Agent 的设置页面配置

LLM_CONFIG = {
    "provider": os.environ.get("LLM_PROVIDER", "deepseek"),
    "api_key": os.environ.get("LLM_API_KEY", ""),
    "model": os.environ.get("LLM_MODEL", "deepseek-chat"),
}

LLM_ENDPOINTS = {
    "deepseek": "https://api.deepseek.com/v1/chat/completions",
    "openai": "https://api.openai.com/v1/chat/completions",
}


def _read_config_sheet():
    """从 Excel 配置表读取 LLM 设置"""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(agent.EXCEL_PATH)
        if "配置" in wb.sheetnames:
            ws = wb["配置"]
            cfg = {}
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0] and row[1]:
                    cfg[str(row[0]).strip()] = str(row[1]).strip()
            wb.close()
            # 如果 Excel 里有 API Key，覆盖环境变量
            if "api_key" in cfg and cfg["api_key"]:
                LLM_CONFIG["api_key"] = cfg["api_key"]
            if "provider" in cfg and cfg["provider"]:
                LLM_CONFIG["provider"] = cfg["provider"]
            if "model" in cfg and cfg["model"]:
                LLM_CONFIG["model"] = cfg["model"]
    except Exception:
        pass


def save_config_sheet(key, value):
    """保存配置到 Excel 配置表"""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(agent.EXCEL_PATH)
        if "配置" not in wb.sheetnames:
            wb.create_sheet("配置")
        ws = wb["配置"]
        # 检查是否已有该键
        found = False
        for row in ws.iter_rows(min_row=2, max_col=2, values_only=False):
            if row[0].value and str(row[0].value).strip() == key:
                row[1].value = value
                found = True
                break
        if not found:
            next_row = (ws.max_row or 1) + 1
            ws.cell(row=next_row, column=1, value=key)
            ws.cell(row=next_row, column=2, value=value)
        agent.save_workbook(wb)
        wb.close()
        return True
    except Exception:
        return False


_read_config_sheet()

def _portfolio_summary():
    """生成简短持仓概况字符串"""
    h = agent.load_portfolio()
    tv, tp, accts = agent.calculate(h)
    cash = sum(x["市值"] for x in h if x["类别"] == "现金")
    stocks_str = ", ".join(f"{x['名称']}{x['数量']}股" for x in h if x["类别"] == "股票")
    return f"总资产{tv:,.0f}元 | 现金{cash:,.0f}元 | 累计盈亏{tp:+,.0f}元 | 持仓: {stocks_str}"


def _build_system_prompt():
    """构建包含持仓上下文的 system prompt"""
    from datetime import timezone, timedelta
    now = datetime.now(timezone(timedelta(hours=8)))
    date_str = now.strftime("%Y年%m月%d日 %A")

    holdings = agent.load_portfolio()
    total_value, total_pnl, accounts = agent.calculate(holdings)
    risks = agent.risk_analysis(holdings)

    # 计算股票占大盘子的比例
    def total_alloc(mv):
        return mv / total_value * 100 if total_value > 0 else 0

    # 加载记忆
    memories = agent.load_memories()
    mem_text = ""
    if memories:
        mem_text = "\n".join(f"- {m}" for m in memories[-20:])

    prompt = f"""你是万万的专属投资理财助手，名叫「万万的投资 Agent」。

当前日期：{date_str}（北京时间）

## 你的角色
- 你是专业的个人理财顾问，擅长 A 股投资分析、仓位管理、资产配置
- 回答要专业、简洁、有实操性
- 对风险要明确提示，不做虚假承诺

## 万万当前的持仓

总资产: {total_value:,.0f} 元
累计盈亏: {total_pnl:+,.0f} 元
总收益率: {total_pnl/(total_value-total_pnl)*100:.2f}%"""

    # 股票持仓（仓位占比 = 占总资产比例）
    stock_text = "\n\n### 股票持仓：\n"
    for h in holdings:
        if h["类别"] == "股票" and (h["数量"] or 0) > 0 and h["账户"] != "账户二（短线）":
            ret = h["收益率"] * 100 if h["收益率"] else 0
            mv = h["市值"] or 0
            alloc = total_alloc(mv)
            stock_text += f"- {h['名称']}：{h['数量']}股 @ {h['现价']:.2f}元，市值 {mv:,.0f}元，盈亏 {h['盈亏']:+,.0f}元 ({ret:+.1f}%)，占总资产 {alloc:.1f}%\n"
    prompt += stock_text

    # 短线账户持仓（单独列出）
    short_text = ""
    for h in holdings:
        if h["类别"] == "股票" and (h["数量"] or 0) > 0 and h["账户"] == "账户二（短线）":
            if not short_text:
                short_text = "\n### 短线账户持仓：\n"
            ret = h["收益率"] * 100 if h["收益率"] else 0
            mv = h["市值"] or 0
            short_text += f"- {h['名称']}：{h['数量']}股 @ {h['现价']:.2f}元，市值 {mv:,.0f}元，盈亏 {h['盈亏']:+,.0f}元 ({ret:+.1f}%)\n"
    if short_text:
        prompt += short_text

    # 账户概况
    acc_text = "\n### 各账户：\n"
    for acct, data in sorted(accounts.items()):
        acc_text += f"- {acct}：资产 {data['市值']:,.0f}元，盈亏 {data['盈亏']:+,.0f}元\n"
    prompt += acc_text

    # 资产配置
    by_type = {}
    for h in holdings:
        t = h["类别"]
        by_type[t] = by_type.get(t, 0) + h["市值"]
    prompt += f"\n### 资产配置：\n"
    for t, v in sorted(by_type.items(), key=lambda x: -x[1]):
        prompt += f"- {t}：{v:,.0f}元 ({v/total_value*100:.1f}%)\n"

    from collections import defaultdict
    sector_groups = defaultdict(list)
    for h2 in holdings:
        if h2["类别"] == "股票" and (h2["数量"] or 0) > 0 and h2["名称"] in agent.SECTOR_MAP:
            sector_groups[agent.SECTOR_MAP[h2["名称"]]].append(h2)
    if sector_groups:
        prompt += "\n### 板块分布：\n"
        for sector, stocks in sorted(sector_groups.items(), key=lambda x: -sum(s["市值"] for s in x[1])):
            names = ", ".join(f"{s['名称']}({s['市值']:,.0f}元)" for s in stocks)
            total_sector = sum(s["市值"] for s in stocks)
            prompt += f"- {sector}：{names}（占总资产 {total_alloc(total_sector):.1f}%）\n"

    # 当前风险提示
    if risks:
        prompt += "\n### 当前诊断的风险：\n"
        for r in risks[:5]:
            prompt += f"- [{r['级别']}] {r['问题']}\n"

    # 近期操作记录
    try:
        edits = agent.load_edit_log()
        if edits:
            prompt += "\n### 近期操作记录：\n"
            for e in edits[-8:]:
                prompt += f"- {e['time']}：{e['detail']}\n"
    except Exception:
        pass

    # 近期出入金
    try:
        flows = agent.load_capital_flows()
        if flows:
            recent = flows[-5:] if len(flows) > 5 else flows
            prompt += "\n### 近期出入金：\n"
            for f in recent:
                sign = "+" if f["type"] == "存入" else "-"
                note = f"（{f['note']}）" if f.get("note") else ""
                prompt += f"- {f['date']} {f['type']} {sign}{f['amount']:,.0f}元{note}\n"
    except Exception:
        pass

    # 净值趋势（最近几天）
    try:
        history = agent.load_history()
        if history:
            recent_h = history[-5:] if len(history) > 5 else history
            prompt += "\n### 近期净值变化：\n"
            for h in recent_h:
                daily = h.get("当日盈亏", 0) or 0
                sign = "+" if daily >= 0 else ""
                prompt += f"- {h['日期']}：总资产 {h['总资产']:,.0f}元，当日盈亏 {sign}{daily:,.0f}元\n"
    except Exception:
        pass

    # AI 记忆
    if mem_text:
        prompt += "\n### 记忆（来自之前的对话）：\n"
        prompt += mem_text + "\n"

    prompt += """

## 重要规则（必须遵守）
1. **A 股最小交易单位是 1手 = 100股**，所有买卖建议的股数必须为 100 的整数倍
2. **仓位百分比以总资产为分母**（不是以股票市值作分母），单票占总资产不超过 25%
3. **给出具体数字**：建议操作时，明确说"卖 X 手（XX 股）"而不是"减仓一部分"
4. **现金为王**：至少保留总资产 10% 的现金
5. **参考操作记录**：回答时可参考用户的近期买卖、出入金操作，分析其操作效果并给出判断

## 回答风格
- 短期建议：具体操作层面（减仓、加仓、止损线）
- 中期建议：配置调整、行业布局
- 长期建议：资产框架、财务规划
- 需要联网信息时，使用 <search>关键词</search> 标明

## 注意
- 你只提供分析建议，不构成投资指令
- 所有操作建议需要结合万万自己的判断
- 数据来自万万手动维护的持仓表，可能有延迟

## 记忆功能
- 你有长期记忆能力，之前的对话要点会显示在"记忆"中
- 如果你想记住用户的重要偏好、决定或结论，在回答末尾加上 <save_memory>内容</save_memory>
- 每次对话开始时你会看到之前保存的记忆"""
    return prompt


def _bing_parse(html, max_results):
    """解析 Bing 搜索结果页（li.b_algo 块），优先取 h2 里的真实标题"""
    results = []
    blocks = re.findall(r'<li class="b_algo".*?</li>', html, re.DOTALL)
    for block in blocks[:max_results]:
        m = re.search(r'<h2[^>]*>\s*<a[^>]*href="(https?://[^"]+)"[^>]*>(.*?)</a>', block, re.DOTALL)
        if not m:
            m = re.search(r'<a[^>]*href="(https?://[^"]+)"[^>]*>(.*?)</a>', block, re.DOTALL)
        if not m:
            continue
        title = re.sub(r'<[^>]+>', '', m.group(2)).strip()
        pm = re.search(r'<p[^>]*>(.*?)</p>', block, re.DOTALL)
        snippet = re.sub(r'<[^>]+>', '', pm.group(1)).strip() if pm else ""
        if title:
            results.append({"title": title, "url": m.group(1), "snippet": snippet})
    return results


def web_search(query, max_results=5):
    """联网搜索，返回搜索结果列表。优先 Bing（国内可达），DuckDuckGo 兜底。"""
    ua = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36"
    results = []

    # 1) Bing 国内可直连
    for attempt in range(2):
        try:
            req = urllib.request.Request(
                f"https://cn.bing.com/search?q={urllib.parse.quote(query)}&setlang=zh-hans",
                headers={"User-Agent": ua},
            )
            resp = urllib.request.urlopen(req, timeout=8)
            html = resp.read().decode("utf-8", errors="ignore")
            results = _bing_parse(html, max_results)
            if results:
                return results
        except Exception as e:
            print(f"⚠️ Bing 搜索尝试 {attempt+1} 失败: {e}")

    # 2) DuckDuckGo 兜底（墙外环境可用）
    for attempt in range(2):
        try:
            if attempt == 0:
                data = urllib.parse.urlencode({"q": query}).encode()
                req = urllib.request.Request(
                    "https://html.duckduckgo.com/html/",
                    data=data,
                    headers={"User-Agent": ua, "Content-Type": "application/x-www-form-urlencoded"},
                )
            else:
                encoded = urllib.parse.quote(query)
                req = urllib.request.Request(
                    f"https://lite.duckduckgo.com/lite/?q={encoded}",
                    headers={"User-Agent": ua},
                )
            resp = urllib.request.urlopen(req, timeout=8)
            html = resp.read().decode("utf-8", errors="ignore")

            if attempt == 0:
                # 解析 HTML 模式：找 result__a（标题链接）和 result__snippet
                titles_links = re.findall(r'<a[^>]*class="result__a"[^>]*href="([^"]*)"[^>]*>(.*?)</a>', html, re.DOTALL)
                snippets = re.findall(r'<a[^>]*class="result__snippet"[^>]*>(.*?)</a>', html, re.DOTALL)
                for i in range(min(len(titles_links), max_results)):
                    url = titles_links[i][0]
                    title = re.sub(r'<[^>]+>', '', titles_links[i][1]).strip()
                    snippet = re.sub(r'<[^>]+>', '', snippets[i]).strip() if i < len(snippets) else ""
                    if url and not url.startswith("http"):
                        url = "https://" + url
                    results.append({"title": title, "url": url, "snippet": snippet})
            else:
                # 解析 Lite 模式
                links = re.findall(r'<a[^>]*href="([^"]*)"[^>]*class="result-link"[^>]*>(.*?)</a>', html, re.DOTALL)
                snippets = re.findall(r'<td[^>]*class="result-snippet"[^>]*>(.*?)</td>', html, re.DOTALL)
                for i in range(min(len(links), max_results)):
                    url = links[i][0]
                    title = re.sub(r'<[^>]+>', '', links[i][1]).strip()
                    snippet = re.sub(r'<[^>]+>', '', snippets[i]).strip() if i < len(snippets) else ""
                    if url and not url.startswith("http"):
                        url = "https://" + url
                    results.append({"title": title, "url": url, "snippet": snippet})

            if results:
                break
        except Exception as e:
            print(f"⚠️ DuckDuckGo 搜索尝试 {attempt+1} 失败: {e}")
            continue

    return results


_NEWS_JUNK = (
    "quote.eastmoney.com", "baike.baidu.com", "xueqiu.com/hq",
    "q.10jqka.com.cn/", "realstock/company", "finance.sina.com.cn/stock/",
)


def _news_ok(title, url, snippet):
    if not title or "http" in title.lower() or len(title) < 6:
        return False
    if len(snippet or "") < 20:
        return False
    for junk in _NEWS_JUNK:
        if junk in (url or ""):
            return False
    return True


def finance_news(limit=20, pages=2, keywords=None):
    """新浪财经 7x24 快讯（真实新闻流），可按持仓关键词优先排序"""
    ua = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36"
    items, seen = [], set()
    for page in range(1, pages + 1):
        try:
            url = (f"https://zhibo.sina.com.cn/api/zhibo/feed?zhibo_id=152"
                   f"&page={page}&page_size=50&tag_id=0")
            req = urllib.request.Request(url, headers={
                "User-Agent": ua, "Referer": "https://finance.sina.com.cn"})
            raw = urllib.request.urlopen(req, timeout=10).read().decode("utf-8", "ignore")
            feed = json.loads(raw)["result"]["data"]["feed"]["list"]
        except Exception as e:
            print(f"⚠️ 新浪快讯第 {page} 页失败: {e}")
            continue
        for it in feed:
            text = re.sub(r"<[^>]+>", "", it.get("rich_text") or it.get("text") or "").strip()
            if not text or text in seen:
                continue
            seen.add(text)
            m = re.match(r"^【(.+?)】(.*)$", text, re.S)
            title = (m.group(1) if m else text[:40]).strip()
            snippet = (m.group(2) if m else text)[:140].strip() or text[:140]
            items.append({
                "title": title, "snippet": snippet,
                "url": (it.get("docurl") or "https://finance.sina.com.cn/7x24/").strip(),
                "time": (it.get("create_time") or "")[5:16],
            })
    kws = [k for k in (keywords or []) if k]
    if kws:
        hit = [i for i in items if any(k in i["title"] + i["snippet"] for k in kws)]
        rest = [i for i in items if i not in hit]
        items = hit + rest
    return items[:limit]


def chat_stream(messages, search_enabled=True):
    """流式调用 LLM API，逐 block 产出文本"""
    api_key = LLM_CONFIG["api_key"]
    if not api_key:
        yield json.dumps({"error": "未配置 API Key，请在设置页面填写"}, ensure_ascii=False)
        return

    provider = LLM_CONFIG["provider"]
    model = LLM_CONFIG["model"]
    endpoint = LLM_ENDPOINTS.get(provider, LLM_ENDPOINTS["deepseek"])
    if not endpoint:
        yield json.dumps({"error": f"未知 provider: {provider}"}, ensure_ascii=False)
        return

    # 构建消息列表
    system_prompt = _build_system_prompt()
    full_messages = [{"role": "system", "content": system_prompt}]

    # 检查最后一条消息是否需要联网搜索
    last_msg = messages[-1]["content"] if messages else ""
    search_keywords = ["搜索", "查一下", "联网", "最新", "今天", "行情", "新闻", "股票", "基金", "涨", "跌"]
    should_search = search_enabled and any(kw in last_msg for kw in search_keywords)

    search_results = None
    if should_search:
        yield json.dumps({"type": "search_status", "content": "🔍 正在联网搜索..."}, ensure_ascii=False)

        # 异步搜索：后台线程 + 心跳保活，防止 Railway 超时断开 SSE
        result_box = []
        def _do_search():
            result_box.append(web_search(last_msg))
        t = threading.Thread(target=_do_search, daemon=True)
        t.start()

        phases = ["⏳ 搜索中", "⏳ 搜索中.", "⏳ 搜索中..", "⏳ 搜索中..."]
        idx = 0
        max_wait = 35
        start = time.time()
        while t.is_alive():
            if time.time() - start > max_wait:
                break
            t.join(timeout=2.5)
            idx = (idx + 1) % len(phases)
            yield json.dumps({"type": "search_heartbeat", "content": phases[idx]}, ensure_ascii=False)

        if result_box:
            search_results = result_box[0]

        if search_results:
            context = "\n\n### 联网搜索结果：\n"
            for i, r in enumerate(search_results, 1):
                context += f"{i}. [{r['title']}]({r['url']})\n   {r['snippet']}\n"
            full_messages[0]["content"] += context
            yield json.dumps({"type": "search_status", "content": f"✅ 搜索到 {len(search_results)} 条结果，正在分析..."}, ensure_ascii=False)
        else:
            yield json.dumps({"type": "search_status", "content": "⚠️ 搜索结果为空，直接分析"}, ensure_ascii=False)

    full_messages.extend(messages)

    # 准备请求
    req_body = json.dumps({
        "model": model,
        "messages": full_messages,
        "stream": True,
        "max_tokens": 4096,
        "temperature": 0.7,
    }).encode("utf-8")

    req = urllib.request.Request(
        endpoint,
        data=req_body,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
            "Accept": "text/event-stream",
        },
        method="POST",
    )

    try:
        resp = urllib.request.urlopen(req, timeout=60)
        buffer = ""
        while True:
            chunk = resp.read(1024)
            if not chunk:
                break
            buffer += chunk.decode("utf-8", errors="ignore")
            while "\n" in buffer:
                line, buffer = buffer.split("\n", 1)
                line = line.strip()
                if not line or line.startswith(":"):
                    continue
                if line == "data: [DONE]":
                    return
                if line.startswith("data: "):
                    try:
                        data = json.loads(line[6:])
                        delta = data.get("choices", [{}])[0].get("delta", {})
                        content = delta.get("content", "")
                        if content:
                            yield json.dumps({"type": "text", "content": content}, ensure_ascii=False)
                    except json.JSONDecodeError:
                        continue
    except urllib.error.HTTPError as e:
        err_body = e.read().decode("utf-8", errors="ignore")
        yield json.dumps({"type": "error", "content": f"API 请求失败 (HTTP {e.code}): {err_body[:200]}"}, ensure_ascii=False)
    except Exception as e:
        yield json.dumps({"type": "error", "content": f"请求异常: {str(e)}"}, ensure_ascii=False)


# ============ API 处理 ============

def _fetch_prev_closes(force=False):
    """昨收价：优先用刷新时落盘的缓存，避免每次打开页面都请求腾讯"""
    import urllib.request
    holdings = agent.load_portfolio()
    codes = {}
    for h in holdings:
        if h["类别"] == "股票" and h.get("代码"):
            codes[h["名称"]] = h["代码"]
    for name, code in agent.STOCK_CODES.items():
        if name not in codes:
            codes[name] = code
    if not codes:
        return {}
    cached = agent.load_prev_closes()
    if not force and cached and all(name in cached for name in codes):
        return cached
    tdx_codes = []
    code_to_name = {}
    for name, code in codes.items():
        for variant in agent._code_variants(code):
            tdx_codes.append(variant)
            code_to_name[variant] = name
    url = "http://qt.gtimg.cn/q=" + ",".join(tdx_codes)
    prev_closes = {}
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        resp = urllib.request.urlopen(req, timeout=8)
        text = resp.read().decode("gbk")
        for line in text.strip().split(";"):
            line = line.strip()
            if not line or "=" not in line:
                continue
            parts = line.split("~")
            if len(parts) >= 5:
                ticker = parts[0].split("=")[0].strip("v_")
                name_from_api = parts[1] if len(parts) > 1 else ""
                if ticker in code_to_name and name_from_api and name_from_api != "?":
                    name = code_to_name[ticker]
                    try:
                        pc = float(parts[4]) if parts[4] else 0
                        if pc > 0 and name not in prev_closes:
                            prev_closes[name] = pc
                    except ValueError:
                        pass
    except Exception:
        pass
    merged = {**cached, **prev_closes}
    agent.save_prev_closes(merged)
    return merged


LAST_REFRESH_FILE = DATA_DIR / ".last_refresh"


def _load_last_refresh():
    try:
        return LAST_REFRESH_FILE.read_text(encoding="utf-8").strip() or None
    except Exception:
        return None


def _save_last_refresh():
    try:
        LAST_REFRESH_FILE.write_text(datetime.now().strftime("%Y-%m-%d %H:%M:%S"), encoding="utf-8")
    except Exception:
        pass


def api_data():
    """返回持仓数据 JSON"""
    holdings = agent.load_portfolio()
    total_value, total_pnl, accounts = agent.calculate(holdings)
    history = agent.load_history()
    prev_closes = _fetch_prev_closes()
    yesterday_qty = agent.load_yesterday_qty_map()  # 名称 → 昨日持股数

    stocks = []
    for h in holdings:
        if h["类别"] == "股票":
            qty = h["数量"] or 0
            if qty == 0:
                continue  # 已清仓的不显示在持仓中
            price = h["现价"] or 0
            prev_close = prev_closes.get(h["名称"], 0)
            # 今日新增的股票（昨日持股数=0），用成本价代替昨收价计算今日收益
            yesterday_q = yesterday_qty.get(h["名称"], None)
            is_new_today = yesterday_q is not None and yesterday_q == 0
            if is_new_today:
                today_pnl = round((price - h["成本价"]) * qty, 2) if qty > 0 else 0
            else:
                today_pnl = round((price - prev_close) * qty, 2) if prev_close > 0 and qty > 0 else 0
            stocks.append({
                "name": h["名称"],
                "code": h["代码"],
                "account": h["账户"],
                "price": price,
                "cost": h["成本价"],
                "qty": qty,
                "market_value": h["市值"],
                "pnl": h["盈亏"],
                "return_pct": round(h["收益率"] * 100, 2) if h["收益率"] else 0,
                "alloc_pct": round(h["仓位占比"] * 100, 1) if h["仓位占比"] else 0,
                "today_pnl": today_pnl,
            })

    by_type = {}
    for h in holdings:
        t = h["类别"]
        by_type[t] = by_type.get(t, 0) + h["市值"]

    # 完整原始数据（包含黄金、现金等，供编辑弹窗使用）
    raw_holdings = []
    for h in holdings:
        raw_holdings.append({
            "type": h["类别"],
            "account": h["账户"],
            "name": h["名称"],
            "code": h.get("代码", ""),
            "qty": h["数量"],
            "cost": h["成本价"],
            "market_value": h.get("市值", 0),
            "note": h.get("备注", ""),
        })

    net_capital, flow_count = agent.get_net_capital()
    true_pnl = round(total_value - net_capital, 2) if net_capital > 0 else total_pnl

    # 当日出入金净额（用于调整今日盈亏）
    # 上一个净值快照已包含该日全部交易，因此只统计严格晚于该日的出入金，
    # 避免快照日当天收盘后发生的黄金卖出等取现被误算成今日亏损
    today = datetime.now().strftime("%Y-%m-%d")
    history = agent.load_history()
    prev_date = today
    for r in history:
        if r["日期"] != today:
            prev_date = r["日期"]
    all_flows = agent.load_capital_flows()
    today_net_flow = sum(
        f["amount"] if f["type"] == "存入" else -f["amount"]
        for f in all_flows if str(f["date"])[:10] > prev_date
    )

    realized_pnl = agent.get_realized_pnl_total()
    ledger_realized = agent.get_ledger_realized_total()
    dividend_total = agent.get_corporate_action_total()
    unrealized_pnl = round(sum(
        (h.get("盈亏") or 0) for h in holdings if h["类别"] in ("股票", "黄金")), 2)
    gold_sold_total = agent.get_gold_sold_total()

    # 已清仓股票（持股数=0）
    cleared_stocks = agent.get_cleared_positions()

    return {
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "updated_at": _load_last_refresh(),
        "total_value": total_value,
        "total_pnl": total_pnl,
        "unrealized_pnl": unrealized_pnl,
        "realized_pnl": realized_pnl,
        "ledger_realized": ledger_realized,
        "dividend_total": dividend_total,
        "gold_sold_total": gold_sold_total,
        "total_return_pct": round(total_pnl / (total_value - total_pnl) * 100, 2) if (total_value - total_pnl) > 0 else 0,
        "net_capital": net_capital,
        "true_pnl": true_pnl,
        "today_net_flow": today_net_flow,
        "stocks": stocks,
        "cleared_stocks": cleared_stocks,
        "accounts": {k: {"value": v["市值"], "pnl": v["盈亏"], "stock_value": v.get("股票市值", 0)} for k, v in accounts.items()},
        "allocation": {k: v for k, v in sorted(by_type.items(), key=lambda x: -x[1])},
        "history": [{"date": r["日期"], "weekday": agent._weekday_cn(r["日期"]), "total": r["总资产"], "pnl": r["累计盈亏"], "daily": r["当日盈亏"]} for r in history],
        "holdings_raw": raw_holdings,
    }


def api_refresh():
    """刷新行情并返回更新后的数据"""
    holdings = agent.load_portfolio()
    prices, prev_closes, failed = agent.fetch_prices()
    if prev_closes:
        agent.save_prev_closes({**agent.load_prev_closes(), **prev_closes})
    if prices:
        holdings = agent.update_holdings_with_prices(holdings, prices, prev_closes)
        agent.save_prices(holdings)
        _save_last_refresh()
    # 记录净值（只在15:00收盘后记录）
    total_value, total_pnl, accounts = agent.calculate(holdings)
    detail = {k: v["市值"] for k, v in accounts.items()}
    snapshot_recorded = agent.save_networth_snapshot(total_value, total_pnl, detail)

    # 每次刷新都同步 Excel 到 GitHub，防止部署丢数据
    try:
        agent.sync_portfolio_to_github()
    except Exception:
        pass
    # 更新 HTML 看板
    try:
        from report_generator import generate as gen_html
        gen_html()
    except Exception:
        pass
    return {
        "success": True,
        "updated": len(prices),
        "failed": failed,
        "snapshot_recorded": snapshot_recorded,
        "data": api_data(),
    }


def api_advice():
    """返回投资建议"""
    holdings = agent.load_portfolio()
    risks = agent.risk_analysis(holdings)
    advice_text = agent.generate_advice(holdings, risks)
    agent.save_advice(advice_text)
    risks_out = [{"level": r["级别"], "issue": r["问题"], "suggestion": r["建议"]} for r in risks]
    return {
        "risks": risks_out,
        "text": advice_text,
    }


def api_breakeven():
    """返回回本分析"""
    holdings = agent.load_portfolio()
    cash = sum(h["市值"] for h in holdings if h["类别"] == "现金")
    return {"stocks": agent.breakeven_analysis(holdings, cash)}


def api_kline():
    """返回持仓股票的日K线数据"""
    try:
        return {"data": agent.fetch_kline()}
    except Exception as e:
        return {"error": str(e), "data": {}}


def api_breakeven_realistic(data):
    """AI分析某只股票基于市场行情的真实回本预估"""
    name = data.get("name", "")
    sector = data.get("sector", "")
    cost = data.get("cost", 0)
    price = data.get("price", 0)
    rise_needed = (cost / price - 1) * 100 if price > 0 else 0
    if not name or not sector:
        return {"error": "参数不全"}

    api_key = LLM_CONFIG["api_key"]
    if not api_key:
        return {"scenarios": None, "error": "未配置 API Key"}

    provider = LLM_CONFIG["provider"]
    model = LLM_CONFIG["model"]
    endpoint = LLM_ENDPOINTS.get(provider, LLM_ENDPOINTS["deepseek"])

    # 搜索市场消息
    queries = [f"{name} 股票 走势 分析 {datetime.now().strftime('%Y-%m')}", f"{sector} 板块 行情 展望"]
    news = []
    for q in queries:
        for r in web_search(q, 4):
            news.append(f"- [{r['title']}] {r['snippet']}")
    context = "\n".join(news[:8]) or "无搜索结果"

    prompt = f"""你是A股技术分析师。分析以下股票的合理回本预期。

股票：{name}
板块：{sector}
成本价：{cost}元
现价：{price}元（需涨 {rise_needed:.1f}% 回本）

## 最新市场信息
{context}

## 请基于以上市场信息，输出三个场景的合理月涨幅预估（纯JSON）：
{{{{
  "optimistic": {{"monthly_rise": 数字, "reason": "简析"}},
  "moderate": {{"monthly_rise": 数字, "reason": "简析"}},
  "conservative": {{"monthly_rise": 数字, "reason": "简析"}}
}}}}

要求：
- 每月涨幅基于真实市场情况（板块热度、个股消息等）
- optimistic 3-8%，moderate 1-4%，conservative 0.3-1.5%
- reason 用一句话说明判断依据，必须用中文
- 输出纯JSON，不要其他文字"""

    messages = [
        {"role": "system", "content": "你必须用中文回复，输出纯JSON，不要任何其他文字。"},
        {"role": "user", "content": prompt},
    ]

    try:
        req_body = json.dumps({"model": model, "messages": messages, "stream": False, "max_tokens": 1024, "temperature": 0.3}).encode()
        req = urllib.request.Request(endpoint, data=req_body, headers={"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"}, method="POST")
        resp = urllib.request.urlopen(req, timeout=60)
        result = json.loads(resp.read().decode("utf-8"))
        content = result["choices"][0]["message"]["content"]
        m = re.search(r'\{.*\}', content, re.DOTALL)
        scenarios = json.loads(m.group()) if m else json.loads(content)

        # 计算回本天数
        out = {}
        for k, v in scenarios.items():
            mr = v["monthly_rise"]
            if mr > 0 and rise_needed > 0:
                days = round(rise_needed / (mr / 22))
            else:
                days = 999
            label_map = {"optimistic":"乐观","moderate":"中性","conservative":"保守"}
            out[k] = {"label": label_map.get(k, k), "days": days, "monthly_rise": mr, "reason": v.get("reason", "")}
        return {"scenarios": out, "rise_needed": round(rise_needed, 1), "name": name}
    except Exception as e:
        return {"error": str(e), "scenarios": None}


def api_add_position_advice(data):
    """AI 分析某只股票是否应该加仓，基于市场消息和行情"""
    stock_name = data.get("name", "")
    sector = data.get("sector", "")
    qty = data.get("qty", 0)
    cost_price = data.get("cost", 0)
    current_price = data.get("price", 0)
    cash = data.get("cash", 0)

    api_key = LLM_CONFIG["api_key"]
    if not api_key:
        return {"error": "未配置 API Key，请在设置页面填写"}

    provider = LLM_CONFIG["provider"]
    model = LLM_CONFIG["model"]
    endpoint = LLM_ENDPOINTS.get(provider, LLM_ENDPOINTS["deepseek"])

    # 搜索股票 + 板块消息
    search_queries = [
        f"{stock_name} 股票 最新消息 {datetime.now().strftime('%Y-%m')}",
        f"{sector} 板块 最新行情 分析",
    ]
    news_items = []
    for q in search_queries:
        results = web_search(q, max_results=4)
        for r in results:
            news_items.append(f"- [{r['title']}] {r['snippet']}")

    search_context = "\n".join(news_items[:8]) if news_items else "未搜索到相关消息"

    rise_needed = (cost_price / current_price - 1) * 100 if current_price > 0 else 0

    portfolio_summary = _portfolio_summary()

    prompt = f"""你是专业的A股投资顾问。请分析以下持仓是否应该加仓。

## 万万整体持仓
{portfolio_summary}

## 持仓信息
- 股票：{stock_name}
- 所属板块：{sector}
- 持有数量：{qty} 股
- 成本均价：{cost_price} 元
- 当前价格：{current_price} 元
- 需涨 {rise_needed:.1f}% 才能回本
- 可用现金：{cash:,.0f} 元

## 最新市场消息
{search_context}

## 任务
基于以上持仓数据和最新的市场消息、板块行情，给出加仓建议。
以 JSON 格式输出（不要其他文字）：
{{{{
  "should_add": true/false,
  "timing": "立即加仓/分批次加仓/观望等回调/不建议加仓",
  "suggested_lots": 建议几手（0表示不加，1手=100股），
  "reason": "判断理由（结合市场消息具体说明）",
  "risk_tip": "风险提示"
}}}}"""

    messages = [
        {"role": "system", "content": "你是严谨的A股投资顾问，总是基于最新市场信息做判断。必须用中文，输出纯 JSON，不要任何其他文字。"},
        {"role": "user", "content": prompt},
    ]

    req_body = json.dumps({
        "model": model,
        "messages": messages,
        "stream": False,
        "max_tokens": 1024,
        "temperature": 0.3,
    }).encode("utf-8")

    try:
        req = urllib.request.Request(
            endpoint, data=req_body,
            headers={"Content-Type": "application/json",
                     "Authorization": f"Bearer {api_key}"},
            method="POST",
        )
        resp = urllib.request.urlopen(req, timeout=60)
        result = json.loads(resp.read().decode("utf-8"))
        content = result["choices"][0]["message"]["content"]

        # 提取 JSON
        json_match = re.search(r'\{.*\}', content, re.DOTALL)
        advice = json.loads(json_match.group()) if json_match else json.loads(content)
        advice["stock_name"] = stock_name
        advice["sector"] = sector
        return advice
    except Exception as e:
        return {"error": f"AI 分析失败: {str(e)}", "stock_name": stock_name}


# ============ 深度诊断 / 晨报 / 决策 ============

def _today_stocks():
    """当前持仓 + 今日涨跌，用于生成提醒"""
    holdings = agent.load_portfolio()
    prev_closes = _fetch_prev_closes()
    yesterday_qty = agent.load_yesterday_qty_map()
    out = []
    for h in holdings:
        if h["类别"] != "股票":
            continue
        qty = h["数量"] or 0
        if qty == 0:
            continue
        price = h["现价"] or 0
        prev_close = prev_closes.get(h["名称"], 0)
        yq = yesterday_qty.get(h["名称"])
        if yq is not None and yq == 0:
            today_pnl = round((price - h["成本价"]) * qty, 2)
        else:
            today_pnl = round((price - prev_close) * qty, 2) if prev_close > 0 else 0
        out.append({
            "name": h["名称"], "price": price, "cost": h["成本价"], "qty": qty,
            "return_pct": round((h["收益率"] or 0) * 100, 2),
            "today_pnl": today_pnl,
        })
    return out


def api_insights():
    holdings = agent.load_portfolio()
    total_value, total_pnl, _ = agent.calculate(holdings)
    snap = insights.snapshot(_today_stocks())
    snap["portfolio"] = {
        "total_value": round(total_value, 2),
        "total_pnl": round(total_pnl, 2),
        "true_pnl": round(total_value - agent.get_net_capital()[0], 2),
        "net_capital": round(agent.get_net_capital()[0], 2),
    }
    return snap


BRIEFING_HEADERS = ["日期", "生成时间", "内容", "来源"]


def _briefing_sheet(wb):
    if "晨报" in wb.sheetnames:
        return wb["晨报"]
    ws = wb.create_sheet("晨报")
    ws.append(BRIEFING_HEADERS)
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 100
    ws.column_dimensions["D"].width = 10
    return ws


def _load_briefing(day):
    try:
        from openpyxl import load_workbook
        wb = load_workbook(agent.EXCEL_PATH)
        if "晨报" not in wb.sheetnames:
            wb.close()
            return None
        ws = wb["晨报"]
        found = None
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] and str(row[0]) == day:
                found = {"date": str(row[0]), "time": str(row[1] or ""),
                         "content": row[2] or "", "source": row[3] or ""}
        wb.close()
        return found
    except Exception as e:
        print(f"⚠️ 读取晨报失败: {e}")
        return None


def _save_briefing(day, content, source):
    try:
        from openpyxl import load_workbook
        wb = load_workbook(agent.EXCEL_PATH)
        ws = _briefing_sheet(wb)
        ws.append([day, datetime.now().strftime("%Y-%m-%d %H:%M"), content, source])
        wb.save(agent.EXCEL_PATH)
        wb.close()
        return True
    except Exception as e:
        print(f"⚠️ 保存晨报失败: {e}")
        return False


# ============ AI 市场判断（企稳与否 + 今天能不能买）============
MARKET_SHEET = "市场判断"
_market_cache = {"at": None, "data": None}


def _market_sheet(wb):
    if MARKET_SHEET in wb.sheetnames:
        return wb[MARKET_SHEET]
    ws = wb.create_sheet(MARKET_SHEET)
    ws.append(["时间", "日期", "结论", "置信度", "内容JSON", "来源JSON"])
    return ws


def _cache_fresh(at, max_age_min):
    if not at:
        return False
    try:
        return (datetime.now() - at).total_seconds() < max_age_min * 60
    except Exception:
        return False


def _load_market_view(max_age_min=30):
    if _market_cache["data"] and _cache_fresh(_market_cache["at"], max_age_min):
        return _market_cache["data"]
    try:
        from openpyxl import load_workbook
        wb = load_workbook(agent.EXCEL_PATH, data_only=True)
        if MARKET_SHEET not in wb.sheetnames:
            wb.close()
            return None
        ws = wb[MARKET_SHEET]
        last = None
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and row[4]:
                last = row
        wb.close()
        if not last:
            return None
        at = datetime.strptime(str(last[0]), "%Y-%m-%d %H:%M:%S")
        if not _cache_fresh(at, max_age_min):
            return None
        data = json.loads(last[4])
        if "stop_loss" not in data:
            return None
        data["sources"] = json.loads(last[5]) if last[5] else []
        _market_cache["at"] = at
        _market_cache["data"] = data
        return data
    except Exception as e:
        print(f"⚠️ 读取市场判断缓存失败: {e}")
        return None


def _save_market_view(data):
    try:
        from openpyxl import load_workbook
        wb = load_workbook(agent.EXCEL_PATH)
        ws = _market_sheet(wb)
        ws.append([
            data.get("as_of", ""), data.get("as_of", "")[:10],
            data.get("stance", ""), data.get("confidence", ""),
            json.dumps({k: v for k, v in data.items() if k != "sources"}, ensure_ascii=False),
            json.dumps(data.get("sources", []), ensure_ascii=False),
        ])
        wb.save(agent.EXCEL_PATH)
        wb.close()
        _market_cache["at"] = datetime.now()
        _market_cache["data"] = data
        return True
    except Exception as e:
        print(f"⚠️ 保存市场判断失败: {e}")
        return False


def _rule_market_view(snap, note=""):
    disc = snap.get("discipline") or {}
    nav = (disc.get("nav") or {})
    ms = snap.get("market") or {}
    stabilized = bool(nav.get("stabilized"))
    stance = "已企稳" if stabilized else "未企稳"
    summary = (
        f"净值{stance}：低点后 {nav.get('days_since_low', 0)} 个净值日不创新低、"
        f"回升 {nav.get('rebound_pct', 0):+.1f}%；沪深300 20 日 {ms.get('ret_20d', 0):+.1f}%。"
    )
    red = disc.get("red_lines") or []
    if red:
        stop_loss = "该止损"
        stop_note = "；".join(r.get("message", "") for r in red) + "。触线就按纪律执行，别拖。"
    else:
        stop_loss = "暂不止损"
        stop_note = "没有持仓触发 -20% 红线，按纪律继续持有。"
    return {
        "stance": stance, "confidence": "低", "summary": summary,
        "market_note": (f"沪深300 最新 {ms.get('close', '—')}，20 日 {ms.get('ret_20d', 0):+.1f}%，"
                        f"距 20 日高点 {ms.get('off_high_20d', 0):+.1f}%。"),
        "account_note": nav.get("stabilized") and "账户净值已满足企稳条件。" or
                        "账户净值尚未满足「连续 3 个净值日不创新低 + 回升 2%」。",
        "reason": note or "未配置 AI Key，只能按大盘和净值规则判断，无法结合新闻。",
        "what_changes": "连续 3 个净值日不创新低且回升 2%，同时大盘不再创新低。",
        "buy_permission": "继续暂停",
        "stop_loss": stop_loss, "stop_loss_note": stop_note,
        "evidence": [], "sources": [],
        "rule": {"stance": stance, "days_since_low": nav.get("days_since_low", 0),
                 "rebound_pct": nav.get("rebound_pct", 0), "trough_date": nav.get("trough_date", "")},
        "as_of": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "source": "规则",
    }


def _ai_market_view():
    api_key = LLM_CONFIG["api_key"] or os.environ.get("DEEPSEEK_API_KEY", "")
    if not api_key:
        return _rule_market_view(insights.snapshot(_today_stocks()))

    snap = insights.snapshot(_today_stocks())
    ms = snap.get("market") or {}
    disc = snap.get("discipline") or {}
    nav = disc.get("nav") or {}
    exp = snap.get("exposure") or {}
    perf_m = (snap.get("performance") or {}).get("metrics", {})
    attr = (snap.get("attribution") or {}).get("items", [])[:5]
    rec = snap.get("recovery") or {}

    holdings_lines = []
    for p in exp.get("positions") or []:
        if not p.get("qty"):
            continue
        ret = p.get("return_pct", 0)
        holdings_lines.append(
            f"- {p['name']}：{p['qty']} 股，成本 {p['cost']:.3f}，现价 {p['price']:.3f}，"
            f"{'浮盈' if ret >= 0 else '浮亏'} {ret:+.1f}%，占总资产 {p['pct_total']:.1f}%")

    today = datetime.now().strftime("%Y-%m-%d")
    names = [p.get("name", "") for p in (exp.get("positions") or [])]
    sector_words = [(s.get("sector") or "").split("（")[0] for s in (exp.get("sectors") or [])[:3]]
    keywords = list(dict.fromkeys([x for x in names + sector_words + ["沪深300", "A股", "大盘"] if x]))

    sources, news_items, seen = [], [], set()
    for it in finance_news(limit=16, keywords=keywords):
        if it["title"] in seen:
            continue
        seen.add(it["title"])
        news_items.append(f"- [{it['time']}] {it['title']}：{it['snippet'][:100]}")
        sources.append({"title": it["title"], "url": it["url"], "snippet": it["snippet"],
                        "time": it["time"]})
        if len(news_items) >= 10:
            break

    sector_q = " ".join(sector_words[:2])
    for q in [f"A股 大盘 走势 分析 {today}", f"{sector_q} 板块 最新消息"]:
        if len(news_items) >= 12:
            break
        for r in web_search(q, max_results=3):
            title = (r.get("title") or "").strip()
            snippet = (r.get("snippet") or "").strip()[:120]
            if title in seen or not _news_ok(title, r.get("url"), snippet):
                continue
            seen.add(title)
            news_items.append(f"- {title}：{snippet}")
            sources.append({"title": title, "url": r.get("url", ""), "snippet": snippet})
            if len(news_items) >= 12:
                break
    news_context = "\n".join(news_items) if news_items else "（本次没有搜到有效新闻）"

    gates_text = "；".join(
        f"{g['label']}：{'通过' if g['ok'] is True else '未通过' if g['ok'] is False else '待确认'}"
        for g in disc.get("gates", []))

    context = f"""## 大盘（沪深300）
- 最新日期/收盘：{ms.get('date', '—')} / {ms.get('close', '—')}
- 近 5 日 {ms.get('ret_5d', 0):+.2f}%；近 20 日 {ms.get('ret_20d', 0):+.2f}%；近 60 日 {ms.get('ret_60d', 0):+.2f}%
- 距 20 日高点 {ms.get('off_high_20d', 0):+.2f}%；高于 20 日低点 {ms.get('above_low_20d', 0):+.2f}%
- 当前连跌 {ms.get('down_streak', 0)} 天 / 连涨 {ms.get('up_streak', 0)} 天

## 账户净值（时间加权，已剔除出入金）
- 机械规则判断：{'已企稳' if nav.get('stabilized') else '未企稳'}
- 最近低点 {nav.get('trough_date', '—')}，低点后 {nav.get('days_since_low', 0)} 个净值日，较低点回升 {nav.get('rebound_pct', 0):+.2f}%
- 区间账户收益 {perf_m.get('account_return', 0):+.2f}%，沪深300 {perf_m.get('benchmark_return', 0):+.2f}%，超额 {perf_m.get('excess_return', 0):+.2f}pp；最大回撤 {perf_m.get('max_drawdown', 0):+.2f}%
- 整体回本需上涨 {rec.get('need_pct', 0):.1f}%

## 账户与纪律
- 总资产 {exp.get('total_value', 0):,.0f} 元；现金 {exp.get('cash', 0):,.0f} 元（{exp.get('cash_pct', 0):.1f}%，底线 20%）
- 前三大持仓合计 {exp.get('top3_pct', 0):.1f}%；最大单票 {exp.get('top1_pct', 0):.1f}%
- 补仓五关：{gates_text}
- 已触发的止损红线：{'；'.join(r.get('message', '') for r in disc.get('red_lines') or []) or '无'}
- 止盈提醒：{'；'.join(t.get('message', '') for t in disc.get('take_profit') or []) or '无'}

## 持仓明细（判断该不该止损用）
{chr(10).join(holdings_lines) or '- 无股票持仓'}

## 盈亏拖累
{chr(10).join(f"- {x['name']}：{x['pnl']:+,.0f} 元（{x['share']:+.0f}%）" for x in attr) or '- 无'}

## 最新新闻（只能引用这里，不能编造）
{news_context}
"""

    system = """你是万万的 A 股投资顾问。你必须给出明确判断，不要和稀泥。
硬性规则：
1. 只能引用上面出现的数字和新闻标题/摘要，禁止编造新闻、政策、支撑位、压力位、点位预测。
2. 严格区分「市场企稳」和「账户净值企稳」，不要混为一谈。
3. 必须遵守她的纪律：现金低于 20% 时，buy_permission 最多只能是「继续暂停」；补仓五关未过，不能建议买入。
4. 止损必须给明确结论：浮亏达到 -20% 的票，按她的纪律必须建议减半，并给出股数和金额，不能因为「长期看好」回避；没触线的票不要主动劝割，除非新闻有明确个股利空，且要指出是哪条。
5. 新闻不足以判断就直说「新闻不足」，并降低 confidence，主要依据指数和净值。
6. 输出纯 JSON，不要任何其他文字、不要 markdown。"""
    prompt = f"""{context}

## 任务
结合大盘、账户净值、持仓明细和新闻，判断：
1) 市场是否企稳；2) 账户是否企稳；3) 今天能不能恢复买入或补仓；4) 该不该止损或减仓。

输出 JSON：
{{
  "stance": "已企稳" 或 "未企稳" 或 "观察中",
  "confidence": "高" 或 "中" 或 "低",
  "summary": "一句话结论，不超过 50 字",
  "market_note": "大盘环境 1-2 句，必须带数字",
  "account_note": "账户净值 1-2 句，必须带机械规则数字",
  "reason": "综合理由 2-3 句，有新闻就引用新闻标题",
  "what_changes": "看到什么会改判，要具体可观察",
  "buy_permission": "可以恢复买入" 或 "只能小幅试买" 或 "继续暂停",
  "stop_loss": "该止损" 或 "暂不止损" 或 "观察",
  "stop_loss_note": "具体到哪只票、卖多少股、约多少钱；暂不止损就说清为什么，不超过 80 字",
  "evidence": ["引用的新闻标题，必须来自上面的列表，没有就空数组"]
}}"""

    raw = _call_llm_text(system, prompt, max_tokens=700, temperature=0.2)
    if not raw:
        return _rule_market_view(snap, note="AI 调用失败，已退回规则判断。")
    try:
        m = re.search(r"\{.*\}", raw, re.DOTALL)
        data = json.loads(m.group()) if m else json.loads(raw)
    except Exception as e:
        print(f"⚠️ 市场判断 JSON 解析失败: {e}\n{raw[:300]}")
        return _rule_market_view(snap, note="AI 返回格式异常，已退回规则判断。")

    if not data.get("stop_loss"):
        data["stop_loss"] = "该止损" if disc.get("red_lines") else "暂不止损"
    if not data.get("stop_loss_note"):
        data["stop_loss_note"] = (
            "；".join(r.get("message", "") for r in disc.get("red_lines") or [])
            or "没有持仓触发 -20% 红线，按纪律继续持有。")
    data["evidence"] = [x for x in (data.get("evidence") or []) if isinstance(x, str)][:4]
    blob = " ".join(str(data.get(k) or "") for k in
                    ("summary", "market_note", "account_note", "reason", "what_changes"))
    blob += " " + " ".join(data["evidence"])
    picked = [s for s in sources if s["title"] in blob or s["title"][:12] in blob]
    data["sources"] = (picked or sources)[:5]
    data["rule"] = {
        "stance": "已企稳" if nav.get("stabilized") else "未企稳",
        "days_since_low": nav.get("days_since_low", 0),
        "rebound_pct": nav.get("rebound_pct", 0),
        "trough_date": nav.get("trough_date", ""),
    }
    data["gates"] = disc.get("gates", [])
    data["as_of"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    data["source"] = "AI"
    data["market"] = ms
    return data


def api_market_view(force=False):
    if not force:
        cached = _load_market_view(30)
        if cached:
            return {"success": True, "cached": True, **cached}
    data = _ai_market_view()
    _save_market_view(data)
    return {"success": True, "cached": False, **data}


def _rule_briefing(snap):
    """不依赖 AI 的兜底晨报：只说结论、原因和触发条件"""
    v = snap.get("verdict") or {}
    h = snap.get("health") or {}
    rec = snap.get("recovery") or {}
    lines = [v.get("headline", "今天：持有观察"), ""]
    if v.get("reasons"):
        lines.append("为什么：")
        for r in v["reasons"]:
            lines.append(f"- {r}")
    if v.get("triggers"):
        lines.append("")
        lines.append("怎么做：")
        for t in v["triggers"]:
            lines.append(f"- {t}")
    disc = snap.get("discipline") or {}
    rules = (disc.get("take_profit") or []) + (disc.get("red_lines") or [])
    if rules:
        lines.append("")
        lines.append("纪律提醒：")
        for r in rules:
            lines.append(f"- {r['message']}")
    lines.append("")
    lines.append(f"账户健康分 {h.get('score', 0)}/100（{h.get('level', '')}）；"
                 f"整体回本需上涨 {rec.get('need_pct', 0):.1f}%。")
    return "\n".join(lines)


def _call_llm_text(system_prompt, user_prompt, max_tokens=900, temperature=0.4):
    api_key = LLM_CONFIG["api_key"]
    if not api_key:
        return None
    endpoint = LLM_ENDPOINTS.get(LLM_CONFIG["provider"], LLM_ENDPOINTS["deepseek"])
    body = json.dumps({
        "model": LLM_CONFIG["model"],
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        "stream": False,
        "max_tokens": max_tokens,
        "temperature": temperature,
    }).encode("utf-8")
    req = urllib.request.Request(
        endpoint, data=body,
        headers={"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"},
        method="POST",
    )
    try:
        resp = urllib.request.urlopen(req, timeout=60)
        data = json.loads(resp.read().decode("utf-8"))
        return data["choices"][0]["message"]["content"].strip()
    except Exception as e:
        print(f"⚠️ 晨报 AI 生成失败: {e}")
        return None


def _ai_briefing(snap):
    p = snap.get("performance", {}).get("metrics", {})
    pf = snap.get("portfolio", {})
    exp = snap.get("exposure", {})
    attr = snap.get("attribution", {}).get("items", [])
    holdings = agent.load_portfolio()
    stocks = [h for h in holdings if h["类别"] == "股票" and (h["数量"] or 0) > 0]

    lines = []
    for h in stocks:
        ret = (h["收益率"] or 0) * 100
        state = "浮盈" if ret >= 0 else "浮亏"
        lines.append(
            f"- {h['名称']} {h['数量']}股 现价{h['现价']:.2f} 成本{h['成本价']:.2f} "
            f"{state}{ret:+.1f}% 占总资产{h['市值'] / pf.get('total_value', 1) * 100:.1f}%"
        )
    v = snap.get("verdict") or {}
    h = snap.get("health") or {}
    rec = snap.get("recovery") or {}
    disc = snap.get("discipline") or {}
    context = f"""## 系统结论（必须遵守，不要改变方向）
{v.get('headline', '')}
原因：{('；'.join(v.get('reasons', []))) or '无'}
系统给出的触发条件：{('；'.join(v.get('triggers', []))) or '无'}

## 账户
总资产 {pf.get('total_value', 0):,.0f} 元；累计盈亏 {pf.get('total_pnl', 0):+,.0f} 元；净投入本金 {pf.get('net_capital', 0):,.0f} 元。
现金占总资产 {exp.get('cash_pct', 0):.1f}%；前三大持仓合计 {exp.get('top3_pct', 0):.1f}%。
健康分 {h.get('score', 0)}/100（{h.get('level', '')}），扣分项：{('；'.join(d['reason'] for d in h.get('deductions', []))) or '无'}。
整体回本需上涨 {rec.get('need_pct', 0):.1f}%（还差 {rec.get('need_amount', 0):,.0f} 元）。
区间（{p.get('start_date', '')} 起）账户收益 {p.get('account_return', 0):+.1f}%，沪深300 {p.get('benchmark_return', 0):+.1f}%，超额 {p.get('excess_return', 0):+.1f}pp；最大回撤 {p.get('max_drawdown', 0):.1f}%。

## 持仓
{chr(10).join(lines)}

## 盈亏归因（负数是拖累）
{chr(10).join(f"- {x['name']}：{x['pnl']:+,.0f} 元，占亏损 {x['share']:.0f}%" for x in attr[:5])}

## 今日异动
{chr(10).join(f"- {a['name']}：{a['text']}" for a in (snap.get('alerts') or [])[:5]) or '- 无异动'}

## 万的交易纪律（必须遵守）
- 现金不低于总资产 20%；单只股票不超过 20%；单笔买入（含补仓）不超过 5%
- 补仓五关：现金≥20%、单票≤20%、单笔≤5%、净值企稳、基本面/消息没变坏，少一关都不能劝补
- 止盈：成本 +20%、+30% 各卖 1/3；卖出红线：浮亏 ≥20% 减半
- 黄金：回到成本卖 1/3，+10% 再卖 1/3

## 当前纪律提醒
{chr(10).join(f"- {x['message']}" for x in (disc.get('take_profit') or []) + (disc.get('red_lines') or [])) or '- 未触发止盈或红线'}
"""
    system = """你是万万的 A 股投资助手。把数据翻译成人话，不要当分析师表演。
硬性要求：
1. 只输出 3 条，每条一行，以「-」开头，总字数不超过 220 字
2. 第一条必须复述系统结论；第二条如果有「当前纪律提醒」，必须原样引用其中一条（保留股票名、股数、金额、价格），没有就写「今天没有新增纪律提醒」；第三条最多再提一只票
3. 只能引用上面明确给出的数字。禁止出现上下文里没有的股票名、价格、股数、涨跌幅；禁止编造支撑位、压力位、技术形态、宏观消息、公司新闻
4. 注意涨跌方向：浮盈是正数（如 +33.6%），浮亏是负数，不要把浮盈写成浮亏
5. 补仓五关没过时，绝不能建议补仓；不要说「适当减仓」，要给具体股数和金额
6. 不承诺收益，不用「大概率」「必涨」这类词；不确定的数字宁可不写
不要标题，不要表格，不要结尾总结。"""
    return _call_llm_text(system, context, max_tokens=400, temperature=0.3)


def api_briefing(force=False):
    day = datetime.now().strftime("%Y-%m-%d")
    cached = _load_briefing(day)
    if cached and cached.get("content") and not force:
        return {"briefing": cached, "cached": True, "snapshot": None}
    snap = api_insights()
    rule_text = _rule_briefing(snap)
    ai_text = _ai_briefing(snap)
    if ai_text:
        _save_briefing(day, ai_text, "AI")
        return {"briefing": {"date": day, "time": datetime.now().strftime("%Y-%m-%d %H:%M"),
                             "content": ai_text, "source": "AI"},
                "cached": False, "snapshot": snap}
    _save_briefing(day, rule_text, "规则")
    return {"briefing": {"date": day, "time": datetime.now().strftime("%Y-%m-%d %H:%M"),
                         "content": rule_text, "source": "规则"},
            "cached": False, "snapshot": snap}


def _invalidate_ai_cache():
    """持仓变化后，旧的 AI 市场判断和当日晨报必须作废"""
    from openpyxl import load_workbook
    wb = load_workbook(agent.EXCEL_PATH)
    changed = False
    if "市场判断" in wb.sheetnames:
        ws = wb["市场判断"]
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)
            changed = True
    if "晨报" in wb.sheetnames:
        ws = wb["晨报"]
        today = datetime.now().strftime("%Y-%m-%d")
        for r in range(ws.max_row, 1, -1):
            if str(ws.cell(r, 1).value)[:10] == today:
                ws.delete_rows(r)
                changed = True
    if changed:
        wb.save(agent.EXCEL_PATH)
    wb.close()


def api_import_statement(data):
    """导入券商对账单：前端把 PDF 转成 base64 传上来，落临时文件后解析"""
    import base64
    import tempfile
    from statement_import import import_statement

    raw = str(data.get("content_b64") or "")
    if not raw:
        raise ValueError("请选择对账单 PDF")
    try:
        blob = base64.b64decode(raw.split(",", 1)[-1])
    except Exception:
        raise ValueError("PDF 数据损坏，请重新选择文件")
    if not blob.startswith(b"%PDF"):
        raise ValueError("不是有效的 PDF 文件")
    password = str(data.get("password") or "")
    if not password:
        raise ValueError("请填写 PDF 密码")
    dry = bool(data.get("dry_run"))
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(blob)
            tmp_path = tmp.name
        result = import_statement(tmp_path, password, str(agent.EXCEL_PATH), dry_run=dry)
    finally:
        if tmp_path:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass
    if not dry:
        _invalidate_ai_cache()
    return result


def api_decisions():
    rows = insights.load_decisions()
    price_map = {h["名称"]: h["现价"] for h in agent.load_portfolio()}
    for r in rows:
        cur = price_map.get(r["name"])
        r["current_price"] = cur
        if cur and r["price"]:
            r["since_pct"] = round((cur / r["price"] - 1) * 100, 2)
        else:
            r["since_pct"] = None
    return {"decisions": rows}


# ============ HTTP 服务器 ============

class PortfolioHandler(BaseHTTPRequestHandler):
    def _send_json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_html(self, html):
        body = html.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        self.end_headers()
        self.wfile.write(body)

    def _read_html(self):
        if not HTML_PATH.exists():
            return "<html><body><h2>webapp.html 未找到</h2></body></html>"
        return HTML_PATH.read_text(encoding="utf-8")

    def _send_static(self, name):
        allow = {"chart.umd.min.js": "application/javascript; charset=utf-8",
                 "chartjs-plugin-datalabels.min.js": "application/javascript; charset=utf-8"}
        ctype = allow.get(name)
        if not ctype:
            self._send_json({"error": "not found"}, 404)
            return
        path = DATA_DIR / "static" / name
        if not path.exists():
            self._send_json({"error": "not found"}, 404)
            return
        body = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-cache")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = urllib.parse.urlparse(self.path).path
        if path == "/" or path == "/index.html":
            self._send_html(self._read_html())

        elif path.startswith("/static/"):
            self._send_static(path[len("/static/"):])

        elif path == "/api/data":
            try:
                self._send_json(api_data())
            except Exception as e:
                self._send_json({"error": str(e)}, 500)

        elif path == "/api/refresh":
            try:
                self._send_json(api_refresh())
            except Exception as e:
                self._send_json({"error": str(e)}, 500)

        elif path == "/api/advice":
            try:
                self._send_json(api_advice())
            except Exception as e:
                self._send_json({"error": str(e)}, 500)

        elif path == "/api/breakeven":
            try:
                self._send_json(api_breakeven())
            except Exception as e:
                self._send_json({"error": str(e)}, 500)

        elif path == "/api/kline":
            try:
                self._send_json(api_kline())
            except Exception as e:
                self._send_json({"error": str(e)}, 500)

        elif path == "/api/capital-flows":
            try:
                flows = agent.load_capital_flows()
                net, cnt = agent.get_net_capital()
                self._send_json({"flows": flows, "net_capital": net})
            except Exception as e:
                self._send_json({"error": str(e)}, 500)

        elif path == "/api/config":
            self._send_json({
                "provider": LLM_CONFIG["provider"],
                "model": LLM_CONFIG["model"],
                "has_key": bool(LLM_CONFIG["api_key"]),
                "key_prefix": LLM_CONFIG["api_key"][:8] + "..." if LLM_CONFIG["api_key"] else "",
                "fees": agent.load_trade_fee_config(),
            })

        elif path == "/api/calendar":
            try:
                force = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query).get("force", ["0"])[0] in ("1", "true")
                self._send_json(monitor.company_events(force))
            except Exception as e:
                self._send_json({"error": str(e)}, 500)

        elif path == "/api/announcements":
            try:
                force = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query).get("force", ["0"])[0] in ("1", "true")
                self._send_json(monitor.announcements(force))
            except Exception as e:
                self._send_json({"error": str(e)}, 500)

        elif path == "/api/sell-plans":
            try:
                self._send_json(monitor.sell_plan_overview())
            except Exception as e:
                self._send_json({"error": str(e)}, 500)

        elif path == "/api/quiet-days":
            try:
                self._send_json(monitor.quiet_overview())
            except Exception as e:
                self._send_json({"error": str(e)}, 500)

        elif path == "/api/realized-pnl":
            try:
                records = agent.get_realized_pnl_records()
                ledger_total = agent.get_ledger_realized_total()
                self._send_json({
                    "records": records,
                    "total": round(sum(r["pnl"] for r in records), 2),
                    "ledger_total": ledger_total,
                    "grand_total": agent.get_realized_pnl_total(),
                })
            except Exception as e:
                self._send_json({"error": str(e)}, 500)

        elif path == "/api/edit-history":
            try:
                logs = agent.load_edit_log()
                self._send_json({"logs": logs})
            except Exception as e:
                self._send_json({"error": str(e)}, 500)

        elif path == "/api/transactions":
            try:
                self._send_json(agent.transaction_review())
            except Exception as e:
                self._send_json({"error": str(e)}, 500)

        elif path == "/api/corporate-actions":
            try:
                self._send_json({
                    "actions": agent.load_corporate_actions(),
                    "total": agent.get_corporate_action_total(),
                })
            except Exception as e:
                self._send_json({"error": str(e)}, 500)

        elif path == "/api/insights":
            try:
                self._send_json(api_insights())
            except Exception as e:
                self._send_json({"error": str(e)}, 500)

        elif path == "/api/briefing":
            try:
                self._send_json(api_briefing())
            except Exception as e:
                self._send_json({"error": str(e)}, 500)

        elif path == "/api/decisions":
            try:
                self._send_json(api_decisions())
            except Exception as e:
                self._send_json({"error": str(e)}, 500)

        elif path == "/api/market-view":
            try:
                self._send_json(api_market_view())
            except Exception as e:
                self._send_json({"error": str(e)}, 500)

        else:
            self._send_json({"error": "not found"}, 404)

    def do_POST(self):
        path = urllib.parse.urlparse(self.path).path
        content_length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(content_length) if content_length > 0 else b"{}"
        try:
            data = json.loads(body) if body else {}
        except json.JSONDecodeError:
            self._send_json({"success": False, "error": "请求体不是合法 JSON"}, 400)
            return

        # 写操作前自动留快照；删除/撤销强制备份（不受 5 分钟节流影响）
        _WRITE_PATHS = {
            "/api/holdings-save", "/api/realized-pnl", "/api/capital-flow",
            "/api/transactions", "/api/decisions", "/api/save-config", "/api/refresh",
            "/api/corporate-actions", "/api/import-statement", "/api/calendar",
            "/api/sell-plans", "/api/quiet-days",
        }
        is_write = (path in _WRITE_PATHS
                    and not (path == "/api/transactions" and data.get("action") == "preview"))
        destructive = (data.get("action") in ("delete", "undo")
                       or path in ("/api/holdings-save", "/api/import-statement"))

        # 多线程下写 Excel 串行化：读-改-写期间不让第二个请求插进来
        if is_write:
            with agent.EXCEL_LOCK:
                agent.auto_backup(path.strip("/").replace("/", "_"), force=destructive)
                self._dispatch_post(path, data)
        else:
            self._dispatch_post(path, data)

    def _dispatch_post(self, path, data):
        if path == "/api/chat":
            self.send_response(200)
            self.send_header("Content-Type", "text/event-stream; charset=utf-8")
            self.send_header("Cache-Control", "no-cache")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()

            messages = data.get("messages", [])
            search_enabled = data.get("search", True)
            full_response = ""

            for chunk in chat_stream(messages, search_enabled):
                try:
                    # 累积文本，提取记忆
                    try:
                        cdata = json.loads(chunk) if chunk.startswith("{") else {}
                        if cdata.get("type") == "text":
                            full_response += cdata.get("content", "")
                    except Exception:
                        pass
                    self.wfile.write(f"data: {chunk}\n\n".encode("utf-8"))
                    self.wfile.flush()
                except BrokenPipeError:
                    break
            try:
                self.wfile.write(b"data: [DONE]\n\n")
                self.wfile.flush()
            except BrokenPipeError:
                pass

            # 保存记忆
            mem_match = re.findall(r'<save_memory>(.*?)</save_memory>', full_response, re.DOTALL)
            for mem_content in mem_match:
                mem_content = mem_content.strip()
                if mem_content:
                    try:
                        agent.save_memory(mem_content)
                        print(f"🧠 保存记忆: {mem_content[:50]}...")
                    except Exception:
                        pass

        elif path == "/api/save-config":
            key = data.get("key", "")
            value = data.get("value", "")
            if key == "api_key":
                LLM_CONFIG["api_key"] = value
                save_config_sheet("api_key", value)
            elif key == "provider":
                LLM_CONFIG["provider"] = value
                save_config_sheet("provider", value)
            elif key == "model":
                LLM_CONFIG["model"] = value
                save_config_sheet("model", value)
            elif key in agent.TRADE_FEE_DEFAULTS:
                try:
                    fee = float(value)
                    if fee < 0:
                        raise ValueError
                except (TypeError, ValueError):
                    self._send_json({"success": False, "error": "费率必须是非负数字"})
                    return
                agent.update_trade_fee_config({key: fee})
            self._send_json({"success": True})

        elif path == "/api/capital-flow":
            try:
                action = data.get("action", "create")
                if action == "update":
                    row = data.get("row")
                    amount = float(data.get("amount", 0))
                    flow_type = data.get("type", "存入")
                    note = data.get("note", "")
                    if not row:
                        self._send_json({"success": False, "error": "缺少行号"})
                        return
                    if amount <= 0:
                        self._send_json({"success": False, "error": "金额必须大于0"})
                        return
                    ok = agent.update_capital_flow(int(row), amount, flow_type, note)
                    if ok:
                        agent.sync_portfolio_to_github()
                    self._send_json({"success": ok})
                elif action == "delete":
                    if not data.get("confirm"):
                        self._send_json({"success": False, "error": "删除需要二次确认"}, 400)
                        return
                    row = data.get("row")
                    if not row:
                        self._send_json({"success": False, "error": "缺少行号"})
                        return
                    ok = agent.delete_capital_flow(int(row))
                    if ok:
                        agent.sync_portfolio_to_github()
                    self._send_json({"success": ok})
                else:
                    amount = float(data.get("amount", 0))
                    flow_type = data.get("type", "存入")
                    note = data.get("note", "")
                    if amount <= 0:
                        self._send_json({"success": False, "error": "金额必须大于0"})
                        return
                    ok = agent.add_capital_flow(amount, flow_type, note)
                    agent.sync_portfolio_to_github()
                    self._send_json({"success": ok})
            except Exception as e:
                self._send_json({"success": False, "error": str(e)}, 500)

        elif path == "/api/capital-flows":
            try:
                flows = agent.load_capital_flows()
                net, cnt = agent.get_net_capital()
                self._send_json({"flows": flows, "net_capital": net})
            except Exception as e:
                self._send_json({"error": str(e)}, 500)

        elif path == "/api/add-position-advice":
            try:
                result = api_add_position_advice(data)
                self._send_json(result)
            except Exception as e:
                self._send_json({"error": str(e)}, 500)

        elif path == "/api/holdings-save":
            try:
                holdings_data = data.get("holdings", [])
                result = agent.save_holdings(holdings_data)
                prices, prev_closes, failed = agent.fetch_prices()
                if prices:
                    result = agent.update_holdings_with_prices(result, prices, prev_closes)
                    agent.save_prices(result)
                # 同步到 GitHub，防止 Railway 部署时数据丢失
                agent.sync_portfolio_to_github()
                self._send_json({"success": True})
            except Exception as e:
                self._send_json({"success": False, "error": str(e)}, 500)

        elif path == "/api/breakeven-realistic":
            try:
                result = api_breakeven_realistic(data)
                self._send_json(result)
            except Exception as e:
                self._send_json({"error": str(e)}, 500)

        elif path == "/api/realized-pnl":
            try:
                action = data.get("action", "create")
                if action == "update":
                    row = data.get("row")
                    name = data.get("name", "").strip()
                    amount = data.get("amount")
                    if not row:
                        self._send_json({"success": False, "error": "缺少行号"})
                        return
                    ok = agent.update_realized_pnl(
                        int(row), name=name or None, amount=amount,
                        date=data.get("date") or None,
                        qty=data.get("qty") if data.get("qty") not in (None, "") else None,
                        price=data.get("price") if data.get("price") not in (None, "") else None,
                        cost=data.get("cost") if data.get("cost") not in (None, "") else None,
                    )
                    if ok:
                        agent.sync_portfolio_to_github()
                    self._send_json({"success": ok})
                elif action == "delete":
                    if not data.get("confirm"):
                        self._send_json({"success": False, "error": "删除需要二次确认"}, 400)
                        return
                    row = data.get("row")
                    if not row:
                        self._send_json({"success": False, "error": "缺少行号"})
                        return
                    ok = agent.delete_realized_pnl(int(row))
                    if ok:
                        agent.sync_portfolio_to_github()
                    self._send_json({"success": ok})
                else:
                    name = data.get("name", "").strip()
                    amount = float(data.get("amount", 0))
                    qty = data.get("qty")
                    price = data.get("price")
                    cost = data.get("cost")
                    if not name:
                        self._send_json({"success": False, "error": "请选择股票"})
                        return
                    if amount == 0 and not (qty and price and cost):
                        self._send_json({"success": False, "error": "请输入盈亏金额，或填写股数/卖出价/成本价"})
                        return
                    agent.record_realized_pnl(
                        name,
                        qty=float(qty or 0),
                        price=float(price or 0),
                        cost=float(cost or 0),
                        amount=amount if amount != 0 else None,
                        date=data.get("date") or None,
                    )
                    agent.sync_portfolio_to_github()
                    self._send_json({"success": True})
            except Exception as e:
                self._send_json({"success": False, "error": str(e)}, 500)

        elif path == "/api/transactions":
            try:
                action = data.get("action")
                if action == "preview":
                    self._send_json({"success": True, "preview": agent.preview_transaction(data)})
                elif action == "undo":
                    if not data.get("confirm"):
                        self._send_json({"success": False, "error": "撤销需要二次确认"}, 400)
                        return
                    result = agent.undo_last_transaction()
                    result.update(agent.transaction_review())
                    agent.sync_portfolio_to_github()
                    self._send_json(result)
                else:
                    result = agent.record_transaction(data)
                    result.update(agent.transaction_review())
                    agent.sync_portfolio_to_github()
                    self._send_json(result)
            except Exception as e:
                self._send_json({"success": False, "error": str(e)}, 400)

        elif path == "/api/corporate-actions":
            try:
                action = data.get("action", "create")
                if action == "delete":
                    if not data.get("confirm"):
                        self._send_json({"success": False, "error": "删除需要二次确认"}, 400)
                        return
                    ok = agent.delete_corporate_action(int(data["row"]))
                else:
                    if not str(data.get("name", "")).strip():
                        self._send_json({"success": False, "error": "请填写名称"})
                        return
                    amount = float(data.get("amount", 0) or 0)
                    if amount == 0:
                        self._send_json({"success": False, "error": "请填写金额（红利税填负数）"})
                        return
                    ok = agent.add_corporate_action(
                        data.get("date") or datetime.now().strftime("%Y-%m-%d"),
                        str(data.get("code", "")).strip(), str(data.get("name", "")).strip(),
                        data.get("type") or "现金分红", amount, data.get("note", ""))
                self._send_json({
                    "success": ok,
                    "actions": agent.load_corporate_actions(),
                    "total": agent.get_corporate_action_total(),
                })
            except Exception as e:
                self._send_json({"success": False, "error": str(e)}, 400)

        elif path == "/api/calendar":
            try:
                action = data.get("action", "create")
                if action == "delete":
                    if not data.get("confirm"):
                        self._send_json({"success": False, "error": "删除需要二次确认"}, 400)
                        return
                    monitor.delete_manual_event(int(data["row"]))
                else:
                    monitor.add_manual_event(data)
                self._send_json({"success": True, **monitor.company_events()})
            except Exception as e:
                self._send_json({"success": False, "error": str(e)}, 400)

        elif path == "/api/sell-plans":
            try:
                action = data.get("action", "create")
                if action == "delete":
                    if not data.get("confirm"):
                        self._send_json({"success": False, "error": "删除需要二次确认"}, 400)
                        return
                    monitor.delete_sell_plan(int(data["row"]))
                elif action == "update":
                    monitor.update_sell_plan(data)
                else:
                    if not str(data.get("name", "")).strip():
                        self._send_json({"success": False, "error": "请填写标的名称"})
                        return
                    monitor.add_sell_plan(data)
                self._send_json({"success": True, **monitor.sell_plan_overview()})
            except Exception as e:
                self._send_json({"success": False, "error": str(e)}, 400)

        elif path == "/api/quiet-days":
            try:
                if data.get("action") == "delete":
                    if not data.get("confirm"):
                        self._send_json({"success": False, "error": "删除需要二次确认"}, 400)
                        return
                    monitor.delete_quiet_day(int(data["row"]))
                else:
                    monitor.add_quiet_day(data.get("date"), data.get("reason", ""))
                self._send_json({"success": True, **monitor.quiet_overview()})
            except Exception as e:
                self._send_json({"success": False, "error": str(e)}, 400)

        elif path == "/api/import-statement":
            try:
                result = api_import_statement(data)
                result = {k: v for k, v in result.items() if k not in ("state", "final")}
                self._send_json({"success": True, **result})
            except Exception as e:
                self._send_json({"success": False, "error": str(e)}, 400)

        elif path == "/api/add-check":
            try:
                name = str(data.get("name", "")).strip()
                if not name:
                    self._send_json({"success": False, "error": "请选择要补仓的股票"})
                    return
                result = insights.add_check(
                    name,
                    amount=data.get("amount"),
                    price=data.get("price"),
                    fundamentals_ok=bool(data.get("fundamentals_ok")),
                    lots=data.get("lots"),
                )
                self._send_json({"success": True, **result})
            except Exception as e:
                self._send_json({"success": False, "error": str(e)}, 400)

        elif path == "/api/briefing":
            try:
                self._send_json(api_briefing(force=bool(data.get("force"))))
            except Exception as e:
                self._send_json({"error": str(e)}, 500)

        elif path == "/api/market-view":
            try:
                self._send_json(api_market_view(force=bool(data.get("force"))))
            except Exception as e:
                self._send_json({"error": str(e)}, 500)

        elif path == "/api/decisions":
            try:
                action = data.get("action", "create")
                if action == "update":
                    ok = insights.update_decision(int(data["row"]),
                                                  status=data.get("status"),
                                                  review=data.get("review"))
                elif action == "delete":
                    if not data.get("confirm"):
                        self._send_json({"success": False, "error": "删除需要二次确认"}, 400)
                        return
                    ok = insights.delete_decision(int(data["row"]))
                else:
                    if not str(data.get("name", "")).strip():
                        self._send_json({"success": False, "error": "请填写标的"})
                        return
                    ok = insights.add_decision(data)
                if ok:
                    agent.sync_portfolio_to_github()
                self._send_json({"success": ok, **api_decisions()})
            except Exception as e:
                self._send_json({"success": False, "error": str(e)}, 500)

        else:
            self._send_json({"error": "not found"}, 404)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")
        self.end_headers()

    def log_message(self, format, *args):
        print(f"[{datetime.now().strftime('%H:%M:%S')}] {args[0]} {args[1]} {args[2]}")


def main():
    ThreadingHTTPServer.allow_reuse_address = True
    ThreadingHTTPServer.daemon_threads = True
    removed = agent.cleanup_old_files()
    if removed:
        print(f"🧹 已清理 {removed} 个过期日志/备份")
    server = ThreadingHTTPServer((HOST, PORT), PortfolioHandler)
    print(f"\n{'='*50}")
    print(f"  📊 投资追踪 Agent 已启动！")
    print(f"  {'='*50}")
    print(f"  浏览器打开:")
    print(f"  → http://{HOST}:{PORT}")
    print(f"  {'='*50}")
    print(f"  按 Ctrl+C 停止服务器\n")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n🛑 服务器已停止")
        server.server_close()


def _pull_latest_from_github():
    """启动时从 GitHub 拉取最新的 portfolio.xlsx，防止部署导致数据丢失"""
    import urllib.request
    token = os.environ.get("GH_TOKEN")
    if not token:
        return
    repo = "wanchangjia-png/investment-agent"
    url = f"https://api.github.com/repos/{repo}/contents/portfolio.xlsx"
    try:
        req = urllib.request.Request(url, headers={
            "Authorization": f"Bearer {token}",
            "User-Agent": "investment-agent",
            "Accept": "application/vnd.github.v3+json",
        })
        resp = urllib.request.urlopen(req, timeout=15)
        import json, base64
        data = json.loads(resp.read().decode("utf-8"))
        content = base64.b64decode(data["content"])
        with open(agent.EXCEL_PATH, "wb") as f:
            f.write(content)
        print("📥 已从 GitHub 拉取最新 portfolio.xlsx")
    except Exception as e:
        print(f"⚠️ 从 GitHub 拉取失败（首次部署或 token 问题）: {e}")


if __name__ == "__main__":
    _pull_latest_from_github()
    main()
