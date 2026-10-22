import monitor


def test_manual_calendar_event(excel):
    detail = "pytest 手动事件"
    monitor.add_manual_event({"date": "2099-01-05", "name": "测试股份",
                              "type": "其他", "detail": detail})
    events = [e for e in monitor.load_manual_events() if e["detail"] == detail]
    assert len(events) == 1
    assert events[0]["date"] == "2099-01-05"

    monitor.delete_manual_event(events[0]["row"])
    assert not [e for e in monitor.load_manual_events() if e["detail"] == detail]


def test_sell_plan_lifecycle(excel):
    name = "pytest卖出标的"
    monitor.add_sell_plan({"date": "2099-02-01", "name": name, "amount": 12345.6,
                           "purpose": "再买入", "target": "测试ETF", "note": "pytest"})
    plans = [p for p in monitor.load_sell_plans() if p["name"] == name]
    assert len(plans) == 1
    plan = plans[0]
    assert plan["amount"] == 12345.6
    assert plan["status"] == "待执行"

    monitor.update_sell_plan({"row": plan["row"], "status": "已执行", "note": "已完成"})
    plan = [p for p in monitor.load_sell_plans() if p["name"] == name][0]
    assert plan["status"] == "已执行"
    assert plan["note"] == "已完成"

    monitor.delete_sell_plan(plan["row"])
    assert not [p for p in monitor.load_sell_plans() if p["name"] == name]


def test_quiet_day_upsert_and_delete(excel):
    day = "2099-03-08"
    monitor.add_quiet_day(day, "测试不动")
    rows = [q for q in monitor.load_quiet_days() if q["date"] == day]
    assert len(rows) == 1
    assert rows[0]["reason"] == "测试不动"

    monitor.add_quiet_day(day, "改个理由")
    rows = [q for q in monitor.load_quiet_days() if q["date"] == day]
    assert len(rows) == 1
    assert rows[0]["reason"] == "改个理由"

    monitor.delete_quiet_day(rows[0]["row"])
    assert not [q for q in monitor.load_quiet_days() if q["date"] == day]


def test_sell_plan_overview_shape(excel):
    overview = monitor.sell_plan_overview()
    assert set(overview) >= {"plans", "unplanned", "cash", "planned_rebuy", "purposes"}
    assert isinstance(overview["cash"], float)
