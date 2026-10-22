import portfolio_agent as agent


def test_realized_total_consistency(excel):
    ledger = agent.get_ledger_realized_total()
    legacy = sum(r["pnl"] for r in agent.get_legacy_realized_records())
    total = agent.get_realized_pnl_total()
    if ledger or legacy:
        assert total == round(ledger + legacy, 2)
    else:
        assert isinstance(total, float)


def test_corporate_action_add_delete(excel):
    before = agent.get_corporate_action_total()
    agent.add_corporate_action("2099-04-01", "000000", "测试股份", "现金分红", 88.8, "pytest")
    assert agent.get_corporate_action_total() == round(before + 88.8, 2)

    record = next(a for a in agent.load_corporate_actions() if a["note"] == "pytest")
    agent.delete_corporate_action(record["row"])
    assert agent.get_corporate_action_total() == before


def test_calculate_includes_realized(excel):
    holdings = agent.load_portfolio()
    total_value, total_pnl, accounts = agent.calculate(holdings)
    assert total_value >= 0
    assert accounts
    expected = sum((h.get("市值") or 0) for h in holdings
                   if h["类别"] in ("股票", "现金", "黄金"))
    assert round(total_value, 2) == round(expected, 2)
