# -*- coding: utf-8 -*-
"""
阿苏私人银行 2.3
Google Sheet 云端存储版 + AI 决策中枢升级版

核心升级：
1. 登录权限：爸爸 / 妈妈 / 阿苏
2. pending_requests 待审批队列（爸爸、妈妈均可批准）
3. rules 风控规则表
4. backups 自动备份与恢复
5. score_history 信用分历史
6. scorecard 信用评分卡
7. rewards / badges 奖励与徽章系统
8. 儿童首页
9. weekly_report 周报
10. AI 消费分类与情景规划

运行：
    streamlit run asu_money2.py

Google Sheet：
    主数据页：state
    A1: key
    B1: value
    A2: bank_data
    B2: 完整 JSON

自动生成展示页：
    summary
    transactions
    budgets
    goals
    merchants
    settings
    pending_requests
    score_history
    weekly_report

Streamlit Secrets：
    SHEET_NAME = "asu_bank"

    [gcp_service_account]
    type = "service_account"
    project_id = "xxx"
    private_key_id = "xxx"
    private_key = "-----BEGIN PRIVATE KEY-----\\nxxx\\n-----END PRIVATE KEY-----\\n"
    client_email = "xxx@xxx.iam.gserviceaccount.com"
    client_id = "xxx"
    auth_uri = "https://accounts.google.com/o/oauth2/auth"
    token_uri = "https://oauth2.googleapis.com/token"
    auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
    client_x509_cert_url = "xxx"

可选：
    DEEPSEEK_API_KEY = "xxx"
    DEEPSEEK_MODEL = "deepseek-v4-flash"

原则：
    Python 负责全部数字计算和审批结论。
    DeepSeek 只负责分类辅助、中文解释、建议话术，不负责编造数字。
"""

from __future__ import annotations

import copy
import json
import os
import re
import uuid
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

try:
    import requests
except Exception:
    requests = None

try:
    import gspread
    from google.oauth2.service_account import Credentials
except Exception:
    gspread = None
    Credentials = None


APP_NAME = "阿苏私人银行 2.3"
SCHEMA_VERSION = "asu-bank-2-3"
DATA_PATH = Path("asu_money2_data.json")
SESSION_KEY = "asu_bank_2_1_state"

STATE_WS = "state"
STATE_KEY = "bank_data"

GSHEET_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# ============================================================
# 1. 基础数据
# ============================================================

def empty_data() -> Dict[str, Any]:
    return {
        "schema_version": SCHEMA_VERSION,
        "settings": {
            "owner": "阿苏",
            "currency": "USD",
            "start_cash": 0.0,
            "start_savings": 0.0,
            "base_score": 0,
            "cash_floor": 0.0,
            "loan_asset_limit": 0.45,
            "hard_loan_asset_limit": 0.65,
            "approval_threshold": 15.0,
            "auto_approve_safe_income": True,
            "created_at": datetime.now().isoformat(timespec="seconds"),
        },
        "rules": {
            "budget_warning_line": 0.90,
            "budget_reject_line": 1.20,
            "score_reject_line": 650,
            "approval_threshold": 15.0,
            "loan_asset_soft_limit": 0.45,
            "loan_asset_hard_limit": 0.65,
            "cash_floor_default": 0.0,
            "weekly_no_impulse_reward": 5,
            "monthly_budget_reward": 10
        },
        "budgets": {},
        "goals": [],
        "merchants": [],
        "transactions": [],
        "pending_requests": [],
        "score_history": [],
        "weekly_reports": [],
        "backups": [],
        "badges": [],
        "rewards": [],
    }


def uid() -> str:
    return str(uuid.uuid4())


def now_str() -> str:
    return datetime.now().isoformat(timespec="seconds")


def today_str() -> str:
    return date.today().isoformat()


def month_str(d: Optional[date] = None) -> str:
    return (d or date.today()).strftime("%Y-%m")


def week_start(d: Optional[date] = None) -> date:
    x = d or date.today()
    return x - timedelta(days=x.weekday())


def fnum(x: Any, default: float = 0.0) -> float:
    try:
        if x is None or x == "":
            return default
        return float(x)
    except Exception:
        return default


def inum(x: Any, default: int = 0) -> int:
    try:
        if x is None or x == "":
            return default
        return int(x)
    except Exception:
        return default


def clamp(x: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, x))


def money(x: Any, currency: str = "USD") -> str:
    v = fnum(x)
    if currency == "USD":
        return f"${v:,.2f}"
    return f"{currency} {v:,.2f}"


def percent(x: Any) -> str:
    return f"{fnum(x) * 100:.0f}%"


def tx_month(tx: Dict[str, Any]) -> str:
    return str(tx.get("date", ""))[:7]


def score_text(score: Optional[int]) -> str:
    return "未建立" if score is None else str(score)


def pretty_value(key: str, value: Any) -> str:
    if value is None:
        return "未建立"
    if isinstance(value, (int, float)):
        if ("率" in key or "占比" in key or "使用率" in key or "上限" in key) and abs(float(value)) <= 3:
            return percent(value)
        if "信用分" in key:
            return f"{value:.0f}"
        return money(value)
    return str(value)


def decision_class(result: str) -> str:
    if result in {"批准", "通过", "开放", "已批准", "已入账"}:
        return "ok"
    if result in {"延迟", "限额通过", "暂缓开放", "观察", "待家长审批", "待审批", "未建立"}:
        return "warn"
    return "bad"


def current_operator() -> str:
    return st.session_state.get("current_operator", "爸爸")


def can_approve() -> bool:
    return current_operator() in {"爸爸", "妈妈"}


def approval_hint() -> str:
    if can_approve():
        return f"当前操作人：{current_operator()}，拥有审批权限。"
    return f"当前操作人：{current_operator()}，只能提交申请，不能批准入账。"


def can_admin() -> bool:
    return current_operator() == "爸爸"


def configured_passwords() -> bool:
    try:
        return any(k in st.secrets for k in ["DAD_PASSWORD", "MOM_PASSWORD", "ASU_PASSWORD"])
    except Exception:
        return False


def require_login() -> None:
    """如果配置了密码，则强制登录；否则沿用侧边栏身份选择。"""
    if not configured_passwords():
        return

    if st.session_state.get("authenticated"):
        st.session_state["current_operator"] = st.session_state.get("authenticated_role", "阿苏")
        return

    st.title("阿苏私人银行登录")
    role = st.selectbox("选择身份", ["爸爸", "妈妈", "阿苏"])
    password = st.text_input("密码", type="password")
    key_map = {"爸爸": "DAD_PASSWORD", "妈妈": "MOM_PASSWORD", "阿苏": "ASU_PASSWORD"}
    expected = secret_value(key_map[role], "")

    if st.button("进入", type="primary"):
        if expected and password == expected:
            st.session_state["authenticated"] = True
            st.session_state["authenticated_role"] = role
            st.session_state["current_operator"] = role
            st.rerun()
        else:
            st.error("密码错误，或该身份未配置密码。")
    st.stop()


def logout() -> None:
    for k in ["authenticated", "authenticated_role"]:
        st.session_state.pop(k, None)
    st.rerun()


def rule_value(data: Dict[str, Any], key: str, default: float) -> float:
    return fnum((data.get("rules") or {}).get(key), default)


def has_activity(data: Dict[str, Any]) -> bool:
    s = data.get("settings", {})
    return (
        fnum(s.get("start_cash")) != 0
        or fnum(s.get("start_savings")) != 0
        or bool(data.get("transactions"))
        or bool(data.get("budgets"))
        or bool(data.get("goals"))
        or bool(data.get("merchants"))
        or bool(data.get("pending_requests"))
    )


def category_options(data: Dict[str, Any]) -> List[str]:
    return sorted(set(list(data.get("budgets", {}).keys()) + ["零花钱", "家庭贷款", "其他"])) or ["其他"]


def spending_categories(data: Dict[str, Any]) -> List[str]:
    cats = list(data.get("budgets", {}).keys())
    return cats if cats else ["其他"]


# ============================================================
# 2. 数据标准化
# ============================================================

def normalize_data(raw: Dict[str, Any]) -> Dict[str, Any]:
    if not isinstance(raw, dict):
        return empty_data()

    base = empty_data()

    settings = base["settings"]
    if isinstance(raw.get("settings"), dict):
        settings.update(raw["settings"])

    settings["owner"] = settings.get("owner") or "阿苏"
    settings["currency"] = settings.get("currency") or "USD"
    settings["start_cash"] = fnum(settings.get("start_cash"))
    settings["start_savings"] = fnum(settings.get("start_savings"))
    settings["base_score"] = inum(settings.get("base_score"), 0)
    settings["cash_floor"] = fnum(settings.get("cash_floor"))
    settings["loan_asset_limit"] = fnum(settings.get("loan_asset_limit"), 0.45)
    settings["hard_loan_asset_limit"] = fnum(settings.get("hard_loan_asset_limit"), 0.65)
    settings["approval_threshold"] = fnum(settings.get("approval_threshold"), 15.0)
    settings["auto_approve_safe_income"] = bool(settings.get("auto_approve_safe_income", True))
    base["settings"] = settings

    # 风控规则，支持从 Google Sheet 展示页查看；主数据仍以 state JSON 为准
    rules = dict(base.get("rules", {}))
    if isinstance(raw.get("rules"), dict):
        for k, v in raw["rules"].items():
            rules[str(k)] = fnum(v, rules.get(str(k), 0.0))
    base["rules"] = rules

    budgets = {}
    for k, v in (raw.get("budgets") or {}).items():
        name = str(k).strip()
        if name:
            budgets[name] = fnum(v)
    base["budgets"] = budgets

    goals = []
    for g in raw.get("goals") or []:
        if isinstance(g, dict):
            goals.append({
                "id": g.get("id") or uid(),
                "name": g.get("name") or "未命名目标",
                "target": fnum(g.get("target")),
                "current": fnum(g.get("current")),
                "deadline": g.get("deadline") or "",
                "category": g.get("category") or "其他",
            })
    base["goals"] = goals

    merchants = []
    for m in raw.get("merchants") or []:
        if isinstance(m, dict):
            merchants.append({
                "id": m.get("id") or uid(),
                "name": m.get("name") or "未命名商户",
                "category": m.get("category") or "其他",
                "discount": fnum(m.get("discount")),
                "required_score": inum(m.get("required_score"), 0),
                "category_budget_cap": fnum(m.get("category_budget_cap"), 1.0),
                "note": m.get("note") or "",
            })
    base["merchants"] = merchants

    txs = []
    for tx in raw.get("transactions") or []:
        if isinstance(tx, dict):
            txs.append({
                "id": tx.get("id") or uid(),
                "date": tx.get("date") or today_str(),
                "type": tx.get("type") or "消费",
                "amount": fnum(tx.get("amount")),
                "category": tx.get("category") or "其他",
                "party": tx.get("party") or tx.get("counterparty") or "",
                "memo": tx.get("memo") or "",
                "expected_repayment": fnum(tx.get("expected_repayment")),
                "principal_repaid": fnum(tx.get("principal_repaid")),
                "interest_received": fnum(tx.get("interest_received")),
                "due_date": tx.get("due_date") or "",
            })
    base["transactions"] = txs

    pending = []
    for r in raw.get("pending_requests") or []:
        if isinstance(r, dict):
            pending.append({
                "id": r.get("id") or uid(),
                "created_at": r.get("created_at") or now_str(),
                "request_text": r.get("request_text") or "",
                "request_type": r.get("request_type") or "消费",
                "amount": fnum(r.get("amount")),
                "category": r.get("category") or "其他",
                "merchant": r.get("merchant") or "",
                "applicant": r.get("applicant") or "阿苏",
                "ai_category": r.get("ai_category") or r.get("category") or "其他",
                "necessity": r.get("necessity") or "未知",
                "impulse_level": r.get("impulse_level") or "未知",
                "python_decision": r.get("python_decision") or "观察",
                "parent_status": r.get("parent_status") or "待审批",
                "final_status": r.get("final_status") or "未入账",
                "decision_json": r.get("decision_json") or {},
                "pending_tx": r.get("pending_tx") or None,
                "parent_note": r.get("parent_note") or "",
                "approved_by": r.get("approved_by") or "",
                "approved_at": r.get("approved_at") or "",
                "booked_at": r.get("booked_at") or "",
            })
    base["pending_requests"] = pending

    score_history = []
    for h in raw.get("score_history") or []:
        if isinstance(h, dict):
            score_history.append({
                "time": h.get("time") or now_str(),
                "score": h.get("score"),
                "change": h.get("change"),
                "event": h.get("event") or "",
                "reason": h.get("reason") or "",
            })
    base["score_history"] = score_history

    weekly_reports = []
    for w in raw.get("weekly_reports") or []:
        if isinstance(w, dict):
            weekly_reports.append({
                "week_start": w.get("week_start") or "",
                "generated_at": w.get("generated_at") or now_str(),
                "content": w.get("content") or "",
                "metrics": w.get("metrics") or {},
            })
    base["weekly_reports"] = weekly_reports

    backups = []
    for b in raw.get("backups") or []:
        if isinstance(b, dict):
            backups.append({
                "id": b.get("id") or uid(),
                "time": b.get("time") or now_str(),
                "operator": b.get("operator") or "未知",
                "event": b.get("event") or "备份",
                "reason": b.get("reason") or "",
                "snapshot_json": b.get("snapshot_json") or "{}",
            })
    base["backups"] = backups[-40:]

    badges = []
    for badge in raw.get("badges") or []:
        if isinstance(badge, dict):
            badges.append({
                "id": badge.get("id") or uid(),
                "name": badge.get("name") or "未命名徽章",
                "earned_at": badge.get("earned_at") or now_str(),
                "description": badge.get("description") or "",
            })
    base["badges"] = badges

    rewards = []
    for reward in raw.get("rewards") or []:
        if isinstance(reward, dict):
            rewards.append({
                "id": reward.get("id") or uid(),
                "title": reward.get("title") or "未命名奖励",
                "created_at": reward.get("created_at") or now_str(),
                "created_by": reward.get("created_by") or "",
                "status": reward.get("status") or "可用",
                "note": reward.get("note") or "",
            })
    base["rewards"] = rewards

    base["schema_version"] = SCHEMA_VERSION
    return base


# ============================================================
# 3. Google Sheet 存储
# ============================================================

def gsheet_enabled() -> bool:
    if gspread is None or Credentials is None:
        return False
    try:
        return "gcp_service_account" in st.secrets and "SHEET_NAME" in st.secrets
    except Exception:
        return False


def get_gsheet_client():
    if not gsheet_enabled():
        raise RuntimeError("Google Sheet 未配置或依赖未安装")
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=GSHEET_SCOPES)
    return gspread.authorize(creds)


def get_spreadsheet():
    return get_gsheet_client().open(str(st.secrets["SHEET_NAME"]))


def get_or_create_ws(spreadsheet, title: str, rows: int = 100, cols: int = 20):
    try:
        return spreadsheet.worksheet(title)
    except Exception:
        return spreadsheet.add_worksheet(title=title, rows=rows, cols=cols)


def col_letters(n: int) -> str:
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def set_table(ws, header: List[str], rows: List[List[Any]]) -> None:
    ws.clear()
    values = [header] + rows
    if values and header:
        end_col = col_letters(len(header))
        ws.update(f"A1:{end_col}{len(values)}", values, value_input_option="RAW")


def sync_readable_sheets(data: Dict[str, Any]) -> None:
    if not gsheet_enabled():
        return

    data = normalize_data(data)
    sh = get_spreadsheet()
    metrics = calc_financials(data)
    ccy = metrics["currency"]

    summary_rows = [
        ["更新时间", now_str(), "程序自动写入"],
        ["总资产", metrics["total_assets"], "现金 + 储蓄 + 应收贷款本金"],
        ["净资产", metrics["net_assets"], "总资产 - 负债"],
        ["现金余额", metrics["cash"], "可立即使用资金"],
        ["储蓄余额", metrics["savings"], "储蓄账户余额"],
        ["应收贷款本金", metrics["receivables"], "已放出但未收回本金"],
        ["负债余额", metrics["liabilities"], "未偿还借入资金"],
        ["信用分", score_text(metrics["credit_score"]), "家庭内部风控评分"],
        ["本月收入", metrics["month_income"], month_str()],
        ["本月消费", metrics["month_expense"], month_str()],
        ["本月支出收入比", metrics["spend_income_ratio"], "支出 / 收入"],
        ["贷款资产占比", metrics["loan_asset_ratio"], "应收贷款本金 / 总资产"],
        ["待审批数量", len([r for r in data.get("pending_requests", []) if r.get("parent_status") == "待审批"]), "pending_requests"],
        ["已获徽章", len(data.get("badges", [])), "badges"],
        ["可用奖励", len([r for r in data.get("rewards", []) if r.get("status") == "可用"]), "rewards"],
        ["备份数量", len(data.get("backups", [])), "backups"],
    ]
    set_table(get_or_create_ws(sh, "summary"), ["指标", "数值", "说明"], summary_rows)

    tx_rows = []
    for tx in sorted(data.get("transactions", []), key=lambda x: str(x.get("date", "")), reverse=True):
        tx_rows.append([
            tx.get("date", ""),
            tx.get("type", ""),
            tx.get("amount", 0),
            tx.get("category", ""),
            tx.get("party", ""),
            tx.get("principal_repaid", 0),
            tx.get("interest_received", 0),
            tx.get("expected_repayment", 0),
            tx.get("due_date", ""),
            tx.get("memo", ""),
        ])
    set_table(
        get_or_create_ws(sh, "transactions", rows=max(100, len(tx_rows) + 5), cols=10),
        ["日期", "类型", "金额", "分类", "对象/商户", "本金回收", "利息收入", "预计回款", "到期日", "备注"],
        tx_rows,
    )

    budget_rows = []
    usage = metrics["budget_usage"]
    for cat, limit in data.get("budgets", {}).items():
        u = usage.get(cat, {"spent": 0, "ratio": 0})
        budget_rows.append([cat, limit, u["spent"], u["ratio"]])
    set_table(get_or_create_ws(sh, "budgets"), ["分类", "月度预算", "本月已花", "使用率"], budget_rows)

    goal_rows = []
    for g in metrics["goals"]:
        goal_rows.append([g["name"], g["category"], g["target"], g["current"], g["remaining"], g["progress"], g.get("deadline", ""), g.get("days_left", "")])
    set_table(get_or_create_ws(sh, "goals"), ["目标", "分类", "目标金额", "当前金额", "剩余金额", "完成度", "截止日", "剩余天数"], goal_rows)

    merchant_rows = []
    for m in data.get("merchants", []):
        res = evaluate_merchant_access(data, m)
        merchant_rows.append([m["name"], m["category"], m["discount"], m["required_score"], m["category_budget_cap"], res["result"], "；".join(res["failed"]), m.get("note", "")])
    set_table(get_or_create_ws(sh, "merchants"), ["商户", "分类", "折扣", "最低信用分", "品类预算上限", "状态", "失败原因", "说明"], merchant_rows)

    s = data.get("settings", {})
    setting_rows = [[k, v] for k, v in s.items()]
    set_table(get_or_create_ws(sh, "settings"), ["设置项", "值"], setting_rows)

    pending_rows = []
    for r in sorted(data.get("pending_requests", []), key=lambda x: str(x.get("created_at", "")), reverse=True):
        pending_rows.append([
            r.get("created_at", ""),
            r.get("applicant", ""),
            r.get("request_type", ""),
            r.get("amount", 0),
            r.get("category", ""),
            r.get("merchant", ""),
            r.get("ai_category", ""),
            r.get("necessity", ""),
            r.get("impulse_level", ""),
            r.get("python_decision", ""),
            r.get("parent_status", ""),
            r.get("final_status", ""),
            r.get("approved_by", ""),
            r.get("request_text", ""),
            r.get("parent_note", ""),
        ])
    set_table(
        get_or_create_ws(sh, "pending_requests", rows=max(100, len(pending_rows) + 5), cols=15),
        ["申请时间", "申请人", "类型", "金额", "Python分类", "商户", "AI分类", "必要性", "冲动等级", "系统结论", "家长状态", "最终状态", "审批人", "原始申请", "家长备注"],
        pending_rows,
    )

    score_rows = []
    for h in sorted(data.get("score_history", []), key=lambda x: str(x.get("time", "")), reverse=True):
        score_rows.append([h.get("time", ""), h.get("score", ""), h.get("change", ""), h.get("event", ""), h.get("reason", "")])
    set_table(get_or_create_ws(sh, "score_history"), ["时间", "信用分", "变化", "事件", "原因"], score_rows)

    report_rows = []
    for w in sorted(data.get("weekly_reports", []), key=lambda x: str(x.get("week_start", "")), reverse=True):
        m = w.get("metrics", {})
        report_rows.append([
            w.get("week_start", ""),
            w.get("generated_at", ""),
            m.get("income", ""),
            m.get("expense", ""),
            m.get("savings", ""),
            m.get("score", ""),
            w.get("content", ""),
        ])
    set_table(get_or_create_ws(sh, "weekly_report"), ["周起始日", "生成时间", "收入", "消费", "储蓄", "信用分", "周报内容"], report_rows)

    rule_rows = [[k, v] for k, v in data.get("rules", {}).items()]
    set_table(get_or_create_ws(sh, "rules"), ["规则", "当前值"], rule_rows)

    backup_rows = []
    for b in sorted(data.get("backups", []), key=lambda x: str(x.get("time", "")), reverse=True):
        backup_rows.append([b.get("id", ""), b.get("time", ""), b.get("operator", ""), b.get("event", ""), b.get("reason", "")])
    set_table(get_or_create_ws(sh, "backups"), ["备份ID", "时间", "操作人", "事件", "原因"], backup_rows)

    badge_rows = []
    for badge in sorted(data.get("badges", []), key=lambda x: str(x.get("earned_at", "")), reverse=True):
        badge_rows.append([badge.get("name", ""), badge.get("earned_at", ""), badge.get("description", "")])
    set_table(get_or_create_ws(sh, "badges"), ["徽章", "获得时间", "说明"], badge_rows)

    reward_rows = []
    for reward in sorted(data.get("rewards", []), key=lambda x: str(x.get("created_at", "")), reverse=True):
        reward_rows.append([reward.get("title", ""), reward.get("created_at", ""), reward.get("created_by", ""), reward.get("status", ""), reward.get("note", "")])
    set_table(get_or_create_ws(sh, "rewards"), ["奖励", "创建时间", "创建人", "状态", "说明"], reward_rows)


def load_data_from_gsheet() -> Dict[str, Any]:
    sh = get_spreadsheet()
    ws = get_or_create_ws(sh, STATE_WS, rows=10, cols=2)
    if ws.row_values(1)[:2] != ["key", "value"]:
        ws.update("A1:B1", [["key", "value"]], value_input_option="RAW")

    rows = ws.get_all_records()
    for row in rows:
        if row.get("key") == STATE_KEY:
            text = row.get("value") or ""
            if not str(text).strip():
                data = empty_data()
                save_data_to_gsheet(data)
                return data
            return normalize_data(json.loads(text))

    data = empty_data()
    save_data_to_gsheet(data)
    return data


def save_data_to_gsheet(data: Dict[str, Any]) -> None:
    data = normalize_data(data)
    sh = get_spreadsheet()
    ws = get_or_create_ws(sh, STATE_WS, rows=10, cols=2)
    if ws.row_values(1)[:2] != ["key", "value"]:
        ws.update("A1:B1", [["key", "value"]], value_input_option="RAW")

    text = json.dumps(data, ensure_ascii=False)
    rows = ws.get_all_records()
    target = None
    for i, row in enumerate(rows, start=2):
        if row.get("key") == STATE_KEY:
            target = i
            break

    if target is None:
        ws.append_row([STATE_KEY, text], value_input_option="RAW")
    else:
        ws.update_cell(target, 2, text)

    sync_readable_sheets(data)


def save_data_local(data: Dict[str, Any]) -> None:
    DATA_PATH.write_text(json.dumps(normalize_data(data), ensure_ascii=False, indent=2), encoding="utf-8")


def load_data_local() -> Dict[str, Any]:
    if not DATA_PATH.exists():
        data = empty_data()
        save_data_local(data)
        return data
    try:
        raw = json.loads(DATA_PATH.read_text(encoding="utf-8"))
        return normalize_data(raw)
    except Exception:
        data = empty_data()
        save_data_local(data)
        return data


def load_data() -> Dict[str, Any]:
    if gsheet_enabled():
        try:
            st.session_state["storage_backend"] = "Google Sheet"
            return load_data_from_gsheet()
        except Exception as exc:
            st.session_state["storage_backend"] = f"本地 JSON 兜底：Google Sheet 读取失败：{exc}"
            return load_data_local()
    st.session_state["storage_backend"] = "本地 JSON"
    return load_data_local()


def save_data(data: Dict[str, Any]) -> None:
    data = normalize_data(data)
    if gsheet_enabled():
        try:
            save_data_to_gsheet(data)
            st.session_state["storage_backend"] = "Google Sheet"
            return
        except Exception as exc:
            st.session_state["storage_backend"] = f"本地 JSON 兜底：Google Sheet 写入失败：{exc}"
    save_data_local(data)


def get_data() -> Dict[str, Any]:
    if SESSION_KEY not in st.session_state:
        st.session_state[SESSION_KEY] = load_data()
    return st.session_state[SESSION_KEY]


def strip_snapshot(data: Dict[str, Any]) -> Dict[str, Any]:
    snap = normalize_data(data)
    # 避免备份里套备份，控制体积
    snap["backups"] = []
    return snap


def add_badge(data: Dict[str, Any], name: str, description: str) -> None:
    existing = {b.get("name") for b in data.get("badges", [])}
    if name not in existing:
        data.setdefault("badges", []).append({
            "id": uid(),
            "name": name,
            "earned_at": now_str(),
            "description": description,
        })


def evaluate_badges(data: Dict[str, Any]) -> None:
    metrics = calc_financials(data)
    if any(tx.get("type") == "收入" for tx in data.get("transactions", [])):
        add_badge(data, "第一笔收入", "已经建立现金流记录。")
    if len(data.get("budgets", {})) >= 3:
        add_badge(data, "预算建筑师", "已经建立至少三个预算分类。")
    if metrics.get("credit_score") is not None:
        add_badge(data, "信用分已建立", "风控系统已经可以追踪信用分。")
    if metrics.get("savings", 0) > 0 or any(fnum(g.get("current")) > 0 for g in data.get("goals", [])):
        add_badge(data, "储蓄启动", "已经开始建立储蓄或目标资金。")
    if data.get("pending_requests"):
        add_badge(data, "合规申请人", "已经使用待审批流程。")
    usage = metrics.get("budget_usage", {})
    if usage and all((r.get("limit", 0) <= 0 or r.get("ratio", 0) <= 1.0) for r in usage.values()):
        add_badge(data, "预算守纪律", "当前预算没有超支。")


def commit(data: Dict[str, Any], event: str = "保存数据", reason: str = "") -> None:
    old_state = normalize_data(st.session_state.get(SESSION_KEY, data)) if SESSION_KEY in st.session_state else normalize_data(data)
    old = calc_financials(old_state).get("credit_score")
    new_data = normalize_data(data)
    backup = {
        "id": uid(),
        "time": now_str(),
        "operator": current_operator(),
        "event": event,
        "reason": reason or "",
        "snapshot_json": json.dumps(strip_snapshot(old_state), ensure_ascii=False),
    }
    new_data["backups"] = (old_state.get("backups", []) + [backup])[-40:]
    new_score = calc_financials(new_data).get("credit_score")
    if new_score is not None:
        change = None if old is None else new_score - old
        new_data.setdefault("score_history", []).append({
            "time": now_str(),
            "score": new_score,
            "change": change,
            "event": event,
            "reason": reason or event,
        })
        new_data["score_history"] = new_data["score_history"][-200:]
    evaluate_badges(new_data)
    st.session_state[SESSION_KEY] = new_data
    save_data(new_data)


def reload_data() -> None:
    st.session_state.pop(SESSION_KEY, None)


def reset_empty() -> None:
    data = empty_data()
    st.session_state[SESSION_KEY] = data
    save_data(data)


def make_tx(
    tx_type: str,
    amount: float,
    category: str,
    party: str = "",
    memo: str = "",
    tx_date: Optional[str] = None,
    expected_repayment: float = 0.0,
    principal_repaid: float = 0.0,
    interest_received: float = 0.0,
    due_date: str = "",
) -> Dict[str, Any]:
    return {
        "id": uid(),
        "date": tx_date or today_str(),
        "type": tx_type,
        "amount": float(amount),
        "category": category,
        "party": party,
        "memo": memo,
        "expected_repayment": float(expected_repayment),
        "principal_repaid": float(principal_repaid),
        "interest_received": float(interest_received),
        "due_date": due_date,
    }


# ============================================================
# 4. 账务计算
# ============================================================

def build_loan_book(transactions: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    loans: List[Dict[str, Any]] = []
    for tx in sorted(transactions, key=lambda x: (str(x.get("date", "")), str(x.get("id", "")))):
        if tx.get("type") == "放贷":
            principal = fnum(tx.get("amount"))
            if principal > 0:
                loans.append({
                    "loan_id": tx.get("id"),
                    "date": tx.get("date"),
                    "borrower": tx.get("party") or "未填写",
                    "principal": principal,
                    "repaid": 0.0,
                    "remaining": principal,
                    "expected_repayment": fnum(tx.get("expected_repayment")),
                    "due_date": tx.get("due_date") or "",
                    "memo": tx.get("memo") or "",
                })
        elif tx.get("type") == "还款":
            borrower = tx.get("party") or ""
            principal_repaid = fnum(tx.get("principal_repaid")) or fnum(tx.get("amount"))
            candidates = [l for l in loans if l["remaining"] > 0 and (not borrower or l["borrower"] == borrower)]
            if not candidates:
                candidates = [l for l in loans if l["remaining"] > 0]
            for loan in candidates:
                if principal_repaid <= 0:
                    break
                applied = min(loan["remaining"], principal_repaid)
                loan["remaining"] -= applied
                loan["repaid"] += applied
                principal_repaid -= applied
    return loans


def calc_budget_usage(data: Dict[str, Any], month: Optional[str] = None) -> Dict[str, Dict[str, float]]:
    month = month or month_str()
    usage = {cat: {"limit": fnum(limit), "spent": 0.0, "ratio": 0.0} for cat, limit in data.get("budgets", {}).items()}
    for tx in data.get("transactions", []):
        if tx.get("type") == "消费" and tx_month(tx) == month:
            cat = tx.get("category") or "其他"
            usage.setdefault(cat, {"limit": 0.0, "spent": 0.0, "ratio": 0.0})
            usage[cat]["spent"] += fnum(tx.get("amount"))
    for row in usage.values():
        row["ratio"] = row["spent"] / row["limit"] if row["limit"] > 0 else 0.0
    return usage


def calc_goals(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    out = []
    now = date.today()
    for g in data.get("goals", []):
        target = fnum(g.get("target"))
        current = fnum(g.get("current"))
        progress = current / target if target > 0 else 0.0
        days_left = None
        if g.get("deadline"):
            try:
                days_left = (date.fromisoformat(g["deadline"]) - now).days
            except Exception:
                days_left = None
        out.append({**g, "progress": progress, "remaining": max(0.0, target - current), "days_left": days_left})
    return out


def calc_credit_score(
    data: Dict[str, Any],
    cash: float,
    savings: float,
    receivables: float,
    liabilities: float,
    total_assets: float,
    month_income: float,
    month_expense: float,
    budget_usage: Dict[str, Dict[str, float]],
    overdue: float,
) -> Tuple[Optional[int], List[str]]:
    s = data.get("settings", {})
    base = inum(s.get("base_score"), 0)
    if base <= 0:
        return None, ["账户尚未建立信用分。请在左侧设置基础信用分。"]

    score = float(base)
    factors = []
    floor = fnum(s.get("cash_floor"))

    if floor > 0:
        if cash < 0:
            score -= 45
            factors.append("现金余额为负：-45")
        elif cash < floor:
            score -= 16
            factors.append(f"现金低于安全线 {money(floor)}：-16")
        elif cash >= floor * 2:
            score += 6
            factors.append("现金缓冲较充足：+6")

    savings_ratio = savings / total_assets if total_assets > 0 else 0
    if total_assets > 0:
        if savings_ratio >= 0.30:
            score += 8
            factors.append("储蓄占总资产 30% 以上：+8")
        elif savings_ratio < 0.10:
            score -= 8
            factors.append("储蓄占总资产低于 10%：-8")

    if month_income <= 0 and month_expense > 0:
        score -= 12
        factors.append("本月有消费但没有收入记录：-12")
    elif month_income > 0:
        ratio = month_expense / month_income
        if ratio > 1.0:
            score -= 35
            factors.append("本月支出超过收入：-35")
        elif ratio > 0.85:
            score -= 18
            factors.append("本月支出达到收入 85% 以上：-18")
        elif ratio > 0.65:
            score -= 8
            factors.append("本月支出达到收入 65% 以上：-8")
        elif ratio <= 0.45:
            score += 5
            factors.append("本月支出控制在收入 45% 以下：+5")

    loan_ratio = receivables / total_assets if total_assets > 0 else 0
    if loan_ratio > 0.65:
        score -= 30
        factors.append("应收贷款本金占总资产超过 65%：-30")
    elif loan_ratio > 0.50:
        score -= 20
        factors.append("应收贷款本金占总资产超过 50%：-20")
    elif loan_ratio > 0.35:
        score -= 9
        factors.append("应收贷款本金占总资产超过 35%：-9")

    debt_ratio = liabilities / total_assets if total_assets > 0 else 0
    if debt_ratio > 0.50:
        score -= 45
        factors.append("负债占总资产超过 50%：-45")
    elif debt_ratio > 0.25:
        score -= 25
        factors.append("负债占总资产超过 25%：-25")
    elif liabilities > 0:
        score -= 7
        factors.append("存在未偿还负债：-7")

    overspent = sum(1 for r in budget_usage.values() if r["limit"] > 0 and r["ratio"] > 1)
    warning = sum(1 for r in budget_usage.values() if r["limit"] > 0 and 0.85 < r["ratio"] <= 1)

    if overspent:
        p = min(32, overspent * 10)
        score -= p
        factors.append(f"{overspent} 个预算品类已超支：-{p}")
    if warning:
        p = min(15, warning * 4)
        score -= p
        factors.append(f"{warning} 个预算品类接近上限：-{p}")
    if overdue > 0:
        score -= 24
        factors.append(f"存在逾期未收回应收本金 {money(overdue)}：-24")

    if not factors:
        factors.append("当前没有明显加分或扣分因素。")
    return int(round(clamp(score, 300, 850))), factors


def calc_financials(data: Dict[str, Any], month: Optional[str] = None) -> Dict[str, Any]:
    month = month or month_str()
    s = data.get("settings", {})
    ccy = s.get("currency", "USD")
    cash = fnum(s.get("start_cash"))
    savings = fnum(s.get("start_savings"))
    liabilities = 0.0
    month_income = 0.0
    month_expense = 0.0
    total_income = 0.0
    total_expense = 0.0

    for tx in sorted(data.get("transactions", []), key=lambda x: str(x.get("date", ""))):
        typ = tx.get("type")
        amount = fnum(tx.get("amount"))
        is_month = tx_month(tx) == month
        if typ == "收入":
            cash += amount
            total_income += amount
            if is_month:
                month_income += amount
        elif typ == "消费":
            cash -= amount
            total_expense += amount
            if is_month:
                month_expense += amount
        elif typ == "转入储蓄":
            cash -= amount
            savings += amount
        elif typ == "储蓄取出":
            cash += amount
            savings -= amount
        elif typ == "放贷":
            cash -= amount
        elif typ == "还款":
            principal = fnum(tx.get("principal_repaid")) or amount
            interest = fnum(tx.get("interest_received"))
            cash += principal + interest
            total_income += interest
            if is_month:
                month_income += interest
        elif typ == "借入":
            cash += amount
            liabilities += amount
        elif typ == "偿还负债":
            cash -= amount
            liabilities -= amount

    liabilities = max(0.0, liabilities)
    loans = build_loan_book(data.get("transactions", []))
    receivables = sum(fnum(l.get("remaining")) for l in loans)

    overdue = 0.0
    now = date.today()
    for loan in loans:
        due = loan.get("due_date") or ""
        remaining = fnum(loan.get("remaining"))
        if remaining > 0 and due:
            try:
                if date.fromisoformat(due) < now:
                    overdue += remaining
            except Exception:
                pass

    total_assets = cash + savings + receivables
    net_assets = total_assets - liabilities
    usage = calc_budget_usage(data, month)
    goals = calc_goals(data)
    score, factors = calc_credit_score(data, cash, savings, receivables, liabilities, total_assets, month_income, month_expense, usage, overdue)

    spend_income_ratio = month_expense / month_income if month_income > 0 else (1.0 if month_expense > 0 else 0.0)
    loan_asset_ratio = receivables / total_assets if total_assets > 0 else 0.0
    debt_asset_ratio = liabilities / total_assets if total_assets > 0 else 0.0
    savings_asset_ratio = savings / total_assets if total_assets > 0 else 0.0

    return {
        "currency": ccy,
        "cash": cash,
        "savings": savings,
        "receivables": receivables,
        "liabilities": liabilities,
        "total_assets": total_assets,
        "net_assets": net_assets,
        "month_income": month_income,
        "month_expense": month_expense,
        "total_income": total_income,
        "total_expense": total_expense,
        "spend_income_ratio": spend_income_ratio,
        "loan_asset_ratio": loan_asset_ratio,
        "debt_asset_ratio": debt_asset_ratio,
        "savings_asset_ratio": savings_asset_ratio,
        "budget_usage": usage,
        "goals": goals,
        "loans": loans,
        "overdue_principal": overdue,
        "credit_score": score,
        "score_factors": factors,
    }


# ============================================================
# 5. 审批、AI分类、情景规划
# ============================================================

def simulate_transaction(data: Dict[str, Any], tx: Dict[str, Any]) -> Dict[str, Any]:
    before = calc_financials(data)
    after_data = copy.deepcopy(data)
    after_data.setdefault("transactions", []).append(tx)
    after = calc_financials(after_data)
    before_score = before["credit_score"]
    after_score = after["credit_score"]
    delta_score = None if before_score is None or after_score is None else after_score - before_score
    return {
        "before": before,
        "after": after,
        "delta": {
            "cash": after["cash"] - before["cash"],
            "receivables": after["receivables"] - before["receivables"],
            "total_assets": after["total_assets"] - before["total_assets"],
            "net_assets": after["net_assets"] - before["net_assets"],
            "credit_score": delta_score,
        },
        "tx": tx,
    }


def usage_for(metrics: Dict[str, Any], cat: str) -> Dict[str, float]:
    return metrics.get("budget_usage", {}).get(cat, {"limit": 0.0, "spent": 0.0, "ratio": 0.0})


def evaluate_purchase_decision(data: Dict[str, Any], amount: float, category: str, desc: str, merchant: str = "") -> Dict[str, Any]:
    tx = make_tx("消费", amount, category, party=merchant, memo=desc)
    sim = simulate_transaction(data, tx)
    before, after, delta = sim["before"], sim["after"], sim["delta"]
    budget = usage_for(after, category)
    floor = fnum(data.get("settings", {}).get("cash_floor"))
    threshold = rule_value(data, "approval_threshold", fnum(data.get("settings", {}).get("approval_threshold"), 15.0))

    result = "批准"
    reasons: List[str] = []

    if amount <= 0:
        result = "拒绝"
        reasons.append("购买金额必须大于 0。")
    if after["cash"] < 0:
        result = "拒绝"
        reasons.append("购买后现金余额为负，触发硬性拒绝。")
    if budget["limit"] > 0 and budget["ratio"] > rule_value(data, "budget_reject_line", 1.20):
        result = "拒绝"
        reasons.append(f"{category}预算使用率将达到 {percent(budget['ratio'])}，超过 120% 红线。")
    if after["credit_score"] is not None and after["credit_score"] < rule_value(data, "score_reject_line", 650):
        result = "拒绝"
        reasons.append("交易后信用分低于 650，进入高风险区。")

    if result != "拒绝":
        flags = []
        if amount >= threshold:
            flags.append(f"金额达到家长审批线 {money(threshold)}。")
        if floor > 0 and after["cash"] < floor:
            flags.append(f"购买后现金余额低于安全线 {money(floor)}。")
        if budget["limit"] > 0 and budget["ratio"] >= rule_value(data, "budget_warning_line", 0.90):
            flags.append(f"{category}预算使用率将达到 {percent(budget['ratio'])}，接近或超过上限。")
        if delta["credit_score"] is not None and delta["credit_score"] <= -8:
            flags.append(f"信用分预计下降 {abs(delta['credit_score'])} 分。")
        if after["month_income"] > 0 and after["spend_income_ratio"] >= 0.85:
            flags.append(f"本月支出/收入比将达到 {percent(after['spend_income_ratio'])}。")
        if flags:
            result = "待家长审批" if amount >= threshold else "延迟"
            reasons.extend(flags)

    if result == "批准":
        reasons.append("现金余额、预算使用率、信用分变化均未触发硬性风控限制。")

    action = "可以购买，但建议先提交审批队列再入账。"
    if result == "拒绝":
        action = "不要购买。先补现金、建预算，或等收入到账后再申请。"
    elif result == "延迟":
        action = "建议延迟到下一笔收入到账后再买，或者把金额拆成两周预算。"
    elif result == "待家长审批":
        action = "进入待审批队列，由家长最终确认是否入账。"

    return {
        "kind": "消费审批",
        "result": result,
        "reasons": reasons,
        "action": action,
        "metrics": {
            "购买金额": amount,
            "消费分类": category,
            "商户": merchant or "未填写",
            "当前现金余额": before["cash"],
            "交易后现金余额": after["cash"],
            "当前信用分": before["credit_score"],
            "交易后信用分": after["credit_score"],
            "信用分变化": delta["credit_score"],
            "品类预算上限": budget["limit"],
            "交易后品类已花": budget["spent"],
            "交易后品类预算使用率": budget["ratio"],
            "交易后本月支出收入比": after["spend_income_ratio"],
        },
        "pending_tx": tx,
        "simulation": sim,
    }


def evaluate_loan_decision(data: Dict[str, Any], principal: float, borrower: str, expected: float, due: str, memo: str = "") -> Dict[str, Any]:
    tx = make_tx("放贷", principal, "家庭贷款", party=borrower, memo=memo, expected_repayment=expected, due_date=due)
    sim = simulate_transaction(data, tx)
    before, after, delta = sim["before"], sim["after"], sim["delta"]
    s = data.get("settings", {})
    floor = fnum(s.get("cash_floor"))
    soft = rule_value(data, "loan_asset_soft_limit", fnum(s.get("loan_asset_limit"), 0.45))
    hard = rule_value(data, "loan_asset_hard_limit", fnum(s.get("hard_loan_asset_limit"), 0.65))

    result = "通过"
    reasons: List[str] = []
    max_by_cash = max(0.0, before["cash"] - floor) if floor > 0 else max(0.0, before["cash"])
    safe_capacity = max(0.0, soft * max(before["total_assets"], 1.0) - before["receivables"])
    suggested = max(0.0, min(principal, max_by_cash, safe_capacity))

    if principal <= 0:
        result = "拒绝"
        reasons.append("放贷本金必须大于 0。")
    if after["cash"] < 0:
        result = "拒绝"
        reasons.append("放贷后现金余额为负，流动性不足。")
    if after["total_assets"] > 0 and after["loan_asset_ratio"] > hard:
        result = "拒绝"
        reasons.append(f"放贷后应收贷款本金占总资产 {percent(after['loan_asset_ratio'])}，超过 {percent(hard)} 红线。")
    if after["credit_score"] is not None and after["credit_score"] < rule_value(data, "score_reject_line", 650):
        result = "拒绝"
        reasons.append("放贷后信用分低于 650。")

    if result != "拒绝":
        flags = []
        if floor > 0 and after["cash"] < floor:
            flags.append(f"放贷后现金低于安全线 {money(floor)}。")
        if after["total_assets"] > 0 and after["loan_asset_ratio"] > soft:
            flags.append(f"放贷后应收贷款本金占总资产 {percent(after['loan_asset_ratio'])}，超过建议线 {percent(soft)}。")
        if principal > suggested and suggested > 0:
            flags.append(f"建议最高放贷金额为 {money(suggested)}。")
        if not due:
            flags.append("没有填写预计还款日，回款纪律不足。")
        if flags:
            result = "限额通过"
            reasons.extend(flags)

    if result == "通过":
        reasons.append("放贷后现金缓冲、贷款资产占比、信用分均未触发硬性风控限制。")

    action = "可以放贷，但建议进入家长审批队列。"
    if result == "拒绝":
        action = "本次不建议放贷。先增加现金余额，或降低放贷金额。"
    elif result == "限额通过":
        action = f"建议限额放贷，最高不超过 {money(suggested)}；必须写清还款日。"

    return {
        "kind": "贷款审批",
        "result": result,
        "reasons": reasons,
        "action": action,
        "metrics": {
            "拟借出本金": principal,
            "借款人": borrower or "未填写",
            "预计回款": expected,
            "当前现金余额": before["cash"],
            "放贷后现金余额": after["cash"],
            "当前应收贷款本金": before["receivables"],
            "放贷后应收贷款本金": after["receivables"],
            "放贷后贷款资产占比": after["loan_asset_ratio"],
            "当前信用分": before["credit_score"],
            "放贷后信用分": after["credit_score"],
            "信用分变化": delta["credit_score"],
            "建议最高放贷金额": suggested,
        },
        "pending_tx": tx,
        "simulation": sim,
    }


def estimate_score_after_transaction(data: Dict[str, Any], tx: Dict[str, Any]) -> Dict[str, Any]:
    sim = simulate_transaction(data, tx)
    return {
        "current_score": sim["before"]["credit_score"],
        "after_score": sim["after"]["credit_score"],
        "change": sim["delta"]["credit_score"],
    }


def secret_value(key: str, default: str = "") -> str:
    try:
        if key in st.secrets:
            return str(st.secrets[key])
    except Exception:
        pass
    return os.getenv(key, default)


def local_category(text: str) -> Dict[str, str]:
    s = (text or "").lower()
    category = "其他"
    rules = {
        "游戏": ["minecraft", "robux", "游戏", "道具", "皮肤", "steam", "switch"],
        "甜品": ["甜品", "奶茶", "冰淇淋", "蛋糕", "糖", "饮料"],
        "玩具": ["玩具", "lego", "乐高", "手办", "娃娃"],
        "学习": ["书", "学习", "课程", "文具", "作业", "训练"],
        "宠物": ["猫", "猫粮", "猫砂", "宠物", "罐头"],
        "餐饮": ["饭", "餐", "披萨", "pizza", "汉堡", "麦当劳"],
    }
    for cat, keys in rules.items():
        if any(k in s for k in keys):
            category = cat
            break
    necessity = "半必要" if category in {"学习", "宠物", "餐饮"} else "非必要"
    impulse = "高" if category in {"游戏", "甜品", "玩具"} else "中"
    merchant = "自然语言输入"
    return {"category": category, "necessity": necessity, "impulse_level": impulse, "merchant": merchant, "note": "本地关键词分类"}


def ai_classify_purchase(text: str) -> Dict[str, str]:
    api_key = secret_value("DEEPSEEK_API_KEY")
    model = secret_value("DEEPSEEK_MODEL", "deepseek-v4-flash")
    if not api_key or requests is None:
        return local_category(text)

    prompt = f"""
请把下面的儿童消费/贷款请求分类，必须只输出 JSON，不要解释。

可用分类：
游戏、甜品、玩具、学习、宠物、餐饮、家庭贷款、其他

字段：
category: 分类
necessity: 必要 / 半必要 / 非必要
impulse_level: 低 / 中 / 高
merchant: 猜测商户，没有则写 自然语言输入
note: 10到20字简短说明

用户输入：
{text}
""".strip()

    try:
        resp = requests.post(
            "https://api.deepseek.com/chat/completions",
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            json={
                "model": model,
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0,
                "max_tokens": 200,
            },
            timeout=12,
        )
        resp.raise_for_status()
        content = resp.json()["choices"][0]["message"]["content"]
        m = re.search(r"\{.*\}", content, re.S)
        if not m:
            return local_category(text)
        out = json.loads(m.group(0))
        return {
            "category": str(out.get("category") or "其他"),
            "necessity": str(out.get("necessity") or "未知"),
            "impulse_level": str(out.get("impulse_level") or "未知"),
            "merchant": str(out.get("merchant") or "自然语言输入"),
            "note": str(out.get("note") or ""),
        }
    except Exception:
        return local_category(text)


def scenario_planning(data: Dict[str, Any], amount: float, category: str, desc: str, merchant: str = "") -> List[Dict[str, Any]]:
    scenarios = []

    today_decision = evaluate_purchase_decision(data, amount, category, desc, merchant)
    scenarios.append({
        "方案": "今天购买",
        "结论": today_decision["result"],
        "交易后现金": today_decision["metrics"]["交易后现金余额"],
        "交易后信用分": today_decision["metrics"]["交易后信用分"],
        "信用分变化": today_decision["metrics"]["信用分变化"],
        "说明": today_decision["action"],
    })

    next_data = copy.deepcopy(data)
    next_data.setdefault("transactions", []).append(make_tx("收入", amount, "零花钱", memo="情景模拟：下周收入到账"))
    next_decision = evaluate_purchase_decision(next_data, amount, category, desc + "（下周购买）", merchant)
    scenarios.append({
        "方案": "下周收入后购买",
        "结论": next_decision["result"],
        "交易后现金": next_decision["metrics"]["交易后现金余额"],
        "交易后信用分": next_decision["metrics"]["交易后信用分"],
        "信用分变化": next_decision["metrics"]["信用分变化"],
        "说明": "先等同等金额收入到账，再购买，现金压力明显下降。",
    })

    half = round(amount / 2, 2)
    split_decision = evaluate_purchase_decision(data, half, category, desc + "（分两周第一笔）", merchant)
    scenarios.append({
        "方案": "分两周购买",
        "结论": split_decision["result"],
        "交易后现金": split_decision["metrics"]["交易后现金余额"],
        "交易后信用分": split_decision["metrics"]["交易后信用分"],
        "信用分变化": split_decision["metrics"]["信用分变化"],
        "说明": f"先支出 {money(half)}，剩余下周再处理，预算冲击最低。",
    })

    return scenarios


def extract_amounts(text: str) -> List[float]:
    nums: List[float] = []
    for p in [r"\$\s*([0-9]+(?:\.[0-9]+)?)", r"([0-9]+(?:\.[0-9]+)?)\s*(?:美元|刀|块|元)"]:
        for m in re.finditer(p, text or "", flags=re.I):
            nums.append(float(m.group(1)))
    if not nums:
        for m in re.finditer(r"([0-9]+(?:\.[0-9]+)?)", text or ""):
            nums.append(float(m.group(1)))
    return nums


def infer_borrower(text: str) -> str:
    m = re.search(r"借给(.+?)(?:\$|[0-9]|，|,|。|$)", text or "")
    if m:
        name = re.sub(r"\s+", "", m.group(1))
        return name[:12] if name else "未填写"
    for name in ["爸爸", "妈妈", "法法", "同学", "朋友"]:
        if name in (text or ""):
            return name
    return "未填写"


def parse_natural_request(data: Dict[str, Any], text: str) -> Dict[str, Any]:
    amounts = extract_amounts(text)
    if any(x in text for x in ["借给", "放贷", "贷款给"]):
        principal = amounts[0] if amounts else 0.0
        expected = amounts[1] if len(amounts) >= 2 else principal
        return evaluate_loan_decision(data, principal, infer_borrower(text), expected, "", text)
    amount = amounts[0] if amounts else 0.0
    cls = ai_classify_purchase(text)
    return evaluate_purchase_decision(data, amount, cls["category"], text, cls["merchant"])


# ============================================================
# 6. 风险、计划、商户、周报
# ============================================================

def build_risk_radar(data: Dict[str, Any]) -> Dict[str, List[Dict[str, str]]]:
    m = calc_financials(data)
    ccy = m["currency"]
    risks = {"高风险": [], "中风险": [], "低风险": []}
    if not has_activity(data):
        risks["低风险"].append({"title": "账户为空", "detail": "暂无交易、暂无预算、暂无目标。请先建立收入、预算或信用分。"})
        return risks

    floor = fnum(data.get("settings", {}).get("cash_floor"))
    if m["cash"] < 0:
        risks["高风险"].append({"title": "现金余额为负", "detail": f"当前现金 {money(m['cash'], ccy)}，应暂停所有非必要消费。"})
    elif floor > 0 and m["cash"] < floor:
        risks["中风险"].append({"title": "现金缓冲偏低", "detail": f"当前现金 {money(m['cash'], ccy)}，低于安全线 {money(floor, ccy)}。"})
    else:
        risks["低风险"].append({"title": "现金余额未触发警报", "detail": f"当前现金 {money(m['cash'], ccy)}。"})

    if m["month_income"] > 0:
        if m["spend_income_ratio"] > 1:
            risks["高风险"].append({"title": "本月支出超过收入", "detail": f"支出/收入比 {percent(m['spend_income_ratio'])}。"})
        elif m["spend_income_ratio"] > 0.85:
            risks["中风险"].append({"title": "本月支出接近收入上限", "detail": f"支出/收入比 {percent(m['spend_income_ratio'])}。"})
        else:
            risks["低风险"].append({"title": "支出收入比可控", "detail": f"支出/收入比 {percent(m['spend_income_ratio'])}。"})

    if m["loan_asset_ratio"] > 0.60:
        risks["高风险"].append({"title": "贷款资产占比过高", "detail": f"应收贷款本金占总资产 {percent(m['loan_asset_ratio'])}，流动性弱。"})
    elif m["loan_asset_ratio"] > 0.45:
        risks["中风险"].append({"title": "贷款资产占比偏高", "detail": f"应收贷款本金占总资产 {percent(m['loan_asset_ratio'])}，建议收回部分本金。"})
    else:
        risks["低风险"].append({"title": "贷款资产占比未触发警报", "detail": f"应收贷款本金占总资产 {percent(m['loan_asset_ratio'])}。"})

    if m["liabilities"] > 0:
        if m["debt_asset_ratio"] > 0.35:
            risks["高风险"].append({"title": "负债偏高", "detail": f"负债占总资产 {percent(m['debt_asset_ratio'])}。"})
        else:
            risks["中风险"].append({"title": "存在未偿还负债", "detail": f"负债余额 {money(m['liabilities'], ccy)}。"})
    else:
        risks["低风险"].append({"title": "无未偿还负债", "detail": "净资产未被负债侵蚀。"})

    for cat, row in m["budget_usage"].items():
        if row["limit"] > 0 and row["ratio"] > 1:
            risks["高风险"].append({"title": f"{cat}预算超支", "detail": f"使用率 {percent(row['ratio'])}。"})
        elif row["limit"] > 0 and row["ratio"] > 0.85:
            risks["中风险"].append({"title": f"{cat}预算接近上限", "detail": f"使用率 {percent(row['ratio'])}。"})

    if m["overdue_principal"] > 0:
        risks["高风险"].append({"title": "存在逾期贷款", "detail": f"逾期本金 {money(m['overdue_principal'], ccy)}，暂停新增放贷。"})
    return risks


def build_weekly_plan(data: Dict[str, Any]) -> List[str]:
    if not has_activity(data):
        return ["先建立第一笔收入或现金余额。", "新增至少 3 个预算分类，例如：游戏、甜品、学习。", "如果要使用信用分审批，请在左侧设置基础信用分。"]

    m = calc_financials(data)
    actions: List[str] = []
    overspent = [cat for cat, r in m["budget_usage"].items() if r["limit"] > 0 and r["ratio"] > 1]
    near = [cat for cat, r in m["budget_usage"].items() if r["limit"] > 0 and 0.85 < r["ratio"] <= 1]
    if overspent:
        actions.append(f"本周暂停 {', '.join(overspent)} 类非必要消费。")
    elif near:
        actions.append(f"本周 {', '.join(near)} 类消费必须先审批。")
    if m["loan_asset_ratio"] > 0.45:
        actions.append("优先收回应收贷款本金，目标至少收回 $10。")
    floor = fnum(data.get("settings", {}).get("cash_floor"))
    if floor > 0 and m["cash"] < floor:
        actions.append("下一笔收入先补现金缓冲，不立刻消费。")
    if m["savings_asset_ratio"] < 0.25 and m["month_income"] > 0:
        actions.append("下一笔收入的 30% 转入储蓄。")
    if m["liabilities"] > 0:
        actions.append("优先偿还负债，不新增借入资金。")
    return actions[:5] or ["维持当前消费节奏，但非必要消费继续走审批。", "下一笔收入至少 20% 转入储蓄。"]


def evaluate_merchant_access(data: Dict[str, Any], merchant: Dict[str, Any]) -> Dict[str, Any]:
    m = calc_financials(data)
    cat = merchant.get("category") or "其他"
    usage = usage_for(m, cat)
    required = inum(merchant.get("required_score"), 0)
    cap = fnum(merchant.get("category_budget_cap"), 1.0)
    score = m["credit_score"]

    checks = {
        "信用分达标": score is not None and score >= required,
        "现金余额为正": m["cash"] > 0,
        "本月支出收入比不高于 90%": m["spend_income_ratio"] <= 0.90,
        "该品类预算未过高": usage["ratio"] <= cap if usage["limit"] > 0 else True,
    }
    opened = all(checks.values())
    failed = [k for k, ok in checks.items() if not ok]
    if opened:
        result = "开放"
        message = f"折扣开放：{fnum(merchant.get('discount')) * 100:.0f}%"
    elif checks["信用分达标"]:
        result = "暂缓开放"
        message = "信用分达标，但现金或预算条件未通过。"
    else:
        result = "关闭"
        message = "信用分未达标或尚未建立。"
    return {
        "result": result,
        "message": message,
        "failed": failed,
        "checks": checks,
        "metrics": {
            "当前信用分": score,
            "要求信用分": required,
            "现金余额": m["cash"],
            "本月支出收入比": m["spend_income_ratio"],
            "该品类预算使用率": usage["ratio"],
            "品类开放上限": cap,
        },
    }


def generate_weekly_report(data: Dict[str, Any]) -> Dict[str, Any]:
    start = week_start()
    end = start + timedelta(days=7)
    rows = []
    for tx in data.get("transactions", []):
        try:
            d = date.fromisoformat(str(tx.get("date")))
            if start <= d < end:
                rows.append(tx)
        except Exception:
            pass

    income = sum(fnum(tx.get("amount")) for tx in rows if tx.get("type") == "收入")
    expense = sum(fnum(tx.get("amount")) for tx in rows if tx.get("type") == "消费")
    savings_move = sum(fnum(tx.get("amount")) for tx in rows if tx.get("type") == "转入储蓄")
    m = calc_financials(data)
    top_cat = "无"
    if rows:
        spend_by_cat: Dict[str, float] = {}
        for tx in rows:
            if tx.get("type") == "消费":
                spend_by_cat[tx.get("category") or "其他"] = spend_by_cat.get(tx.get("category") or "其他", 0) + fnum(tx.get("amount"))
        if spend_by_cat:
            top_cat = max(spend_by_cat, key=spend_by_cat.get)

    actions = build_weekly_plan(data)
    content = "\n".join([
        f"阿苏私人银行周报：{start.isoformat()} 至 {(end - timedelta(days=1)).isoformat()}",
        f"本周收入：{money(income)}",
        f"本周消费：{money(expense)}",
        f"本周转入储蓄：{money(savings_move)}",
        f"最大消费分类：{top_cat}",
        f"当前信用分：{score_text(m['credit_score'])}",
        "下周行动：",
        *[f"{i+1}. {a}" for i, a in enumerate(actions)],
    ])

    return {
        "week_start": start.isoformat(),
        "generated_at": now_str(),
        "content": content,
        "metrics": {
            "income": income,
            "expense": expense,
            "savings": savings_move,
            "score": score_text(m["credit_score"]),
            "top_category": top_cat,
        },
    }


# ============================================================
# 7. DeepSeek 报告
# ============================================================

def local_report(decision: Dict[str, Any]) -> str:
    lines = [f"### {decision.get('kind', '风控报告')}", f"**审批结论：{decision.get('result')}**", "", "#### 关键硬指标"]
    for k, v in decision.get("metrics", {}).items():
        lines.append(f"- {k}：{pretty_value(k, v)}")
    lines += ["", "#### 风控原因"]
    for r in decision.get("reasons", []):
        lines.append(f"- {r}")
    lines += ["", "#### 操作指令", decision.get("action", "")]
    return "\n".join(lines)


def deepseek_decision_report(decision: Dict[str, Any]) -> str:
    api_key = secret_value("DEEPSEEK_API_KEY")
    model = secret_value("DEEPSEEK_MODEL", "deepseek-v4-flash")
    if not api_key or requests is None:
        return local_report(decision)

    payload = {
        "kind": decision.get("kind"),
        "result": decision.get("result"),
        "metrics": decision.get("metrics"),
        "reasons": decision.get("reasons"),
        "action": decision.get("action"),
    }
    system_prompt = "你是阿苏私人银行2.1的中文私人银行客户经理。只能解释Python已计算的硬指标。禁止编造数字，禁止改变审批结论。"
    try:
        resp = requests.post(
            "https://api.deepseek.com/chat/completions",
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            json={
                "model": model,
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": json.dumps(payload, ensure_ascii=False, indent=2)},
                ],
                "temperature": 0.2,
                "max_tokens": 900,
            },
            timeout=20,
        )
        resp.raise_for_status()
        return resp.json()["choices"][0]["message"]["content"]
    except Exception as exc:
        return local_report(decision) + f"\n\n> DeepSeek 调用失败，已切换本地报告：{exc}"


# ============================================================
# 8. UI
# ============================================================

def inject_css() -> None:
    st.markdown(
        """
        <style>
        .main .block-container {max-width:1280px; padding-top:1.2rem; padding-bottom:2rem;}
        .hero {background:linear-gradient(135deg,#003C71 0%,#005EB8 50%,#0A74DA 100%); color:white; padding:26px 30px; border-radius:24px; margin-bottom:18px; box-shadow:0 16px 38px rgba(0,62,130,.25);}
        .hero h1 {margin:0; font-size:34px;}
        .hero p {margin:8px 0 0; opacity:.93; font-size:16px;}
        .card {border:1px solid #D6E4F5; border-radius:18px; background:white; padding:18px; box-shadow:0 8px 22px rgba(15,23,42,.06); min-height:112px;}
        .label {color:#6B7280; font-size:13px; margin-bottom:6px;}
        .value {font-size:28px; font-weight:780; color:#111827; line-height:1.12;}
        .sub {color:#6B7280; font-size:12px; margin-top:6px;}
        .decision {border:1px solid #D6E4F5; border-radius:18px; background:white; padding:18px; margin:10px 0; box-shadow:0 8px 22px rgba(15,23,42,.05);}
        .ok {border-left:7px solid #166534;}
        .warn {border-left:7px solid #B45309;}
        .bad {border-left:7px solid #B91C1C;}
        .pill {display:inline-block; padding:4px 10px; border-radius:999px; background:#EAF3FF; color:#003C71; border:1px solid #B9D7F6; font-size:12px; font-weight:700; margin-bottom:8px;}
        .small {color:#6B7280; font-size:13px;}
        div[data-testid="stMetric"] {border:1px solid #D6E4F5; border-radius:16px; padding:12px 14px; background:white; box-shadow:0 6px 18px rgba(15,23,42,.05);}
        .stTabs [data-baseweb="tab-list"] {gap:5px;}
        .stTabs [data-baseweb="tab"] {border-radius:999px; background:#F2F6FB; padding:8px 13px;}
        .stTabs [aria-selected="true"] {background:#005EB8 !important; color:white !important;}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_hero(data: Dict[str, Any]) -> None:
    owner = data.get("settings", {}).get("owner", "阿苏")
    st.markdown(
        f"""
        <div class="hero">
            <h1>{APP_NAME}</h1>
            <p>{owner} 的家庭金融审批系统：密码权限 · AI分类 · 家长审批 · 备份恢复 · 奖励徽章 · 自动周报</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.caption(f"当前存储：{st.session_state.get('storage_backend', '未知')}")


def metric_card(label: str, value: str, sub: str = "") -> None:
    st.markdown(
        f"""
        <div class="card">
            <div class="label">{label}</div>
            <div class="value">{value}</div>
            <div class="sub">{sub}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_decision(decision: Dict[str, Any], data: Dict[str, Any], key: str, allow_queue: bool = True) -> None:
    css = decision_class(decision.get("result", "观察"))
    st.markdown(
        f"""
        <div class="decision {css}">
            <div class="pill">{decision.get('kind', '审批')}</div>
            <h3 style="margin: 4px 0 8px;">审批结果：{decision.get('result')}</h3>
            <p class="small">{decision.get('action', '')}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if decision.get("reasons"):
        st.markdown("#### 风控原因")
        for r in decision["reasons"]:
            st.write(f"- {r}")

    if decision.get("metrics"):
        st.dataframe(pd.DataFrame([{"指标": k, "值": pretty_value(k, v)} for k, v in decision["metrics"].items()]), use_container_width=True, hide_index=True)

    with st.expander("AI 客户经理报告", expanded=True):
        st.markdown(deepseek_decision_report(decision))

    tx = decision.get("pending_tx")
    if tx:
        c1, c2 = st.columns(2)
        with c1:
            disabled = decision.get("result") == "拒绝" or not can_approve()
            label = "家长直接入账（爸爸/妈妈）" if can_approve() else "无直接入账权限"
            if st.button(label, disabled=disabled, type="primary", key=f"{key}_book"):
                data.setdefault("transactions", []).append(tx)
                commit(data, event=f"{current_operator()}直接入账", reason=decision.get("result", ""))
                st.success(f"{current_operator()}已直接入账并同步存储。")
                st.rerun()
            if not can_approve():
                st.caption("阿苏账号只能提交到待审批队列。")
        with c2:
            if allow_queue and st.button("提交到待审批队列", disabled=decision.get("result") == "拒绝", key=f"{key}_queue"):
                req = {
                    "id": uid(),
                    "created_at": now_str(),
                    "request_text": tx.get("memo", ""),
                    "request_type": tx.get("type", "消费"),
                    "amount": tx.get("amount", 0),
                    "category": tx.get("category", "其他"),
                    "merchant": tx.get("party", ""),
                    "applicant": current_operator(),
                    "ai_category": tx.get("category", "其他"),
                    "necessity": "见AI分类",
                    "impulse_level": "见AI分类",
                    "python_decision": decision.get("result", ""),
                    "parent_status": "待审批",
                    "final_status": "未入账",
                    "decision_json": decision,
                    "pending_tx": tx,
                    "parent_note": "",
                    "approved_by": "",
                    "approved_at": "",
                    "booked_at": "",
                }
                data.setdefault("pending_requests", []).append(req)
                commit(data, event="提交待审批", reason=tx.get("memo", ""))
                st.success("已提交到待审批队列。")
                st.rerun()


def page_home(data: Dict[str, Any]) -> None:
    m = calc_financials(data)
    ccy = m["currency"]

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        metric_card("总资产", money(m["total_assets"], ccy), "现金 + 储蓄 + 应收贷款本金")
    with c2:
        metric_card("净资产", money(m["net_assets"], ccy), "总资产 - 负债")
    with c3:
        metric_card("现金余额", money(m["cash"], ccy), "可立即使用资金")
    with c4:
        metric_card("信用分", score_text(m["credit_score"]), "家庭内部风控评分，非 FICO")

    c5, c6, c7, c8 = st.columns(4)
    with c5:
        metric_card("待审批", str(len([r for r in data.get("pending_requests", []) if r.get("parent_status") == "待审批"])), "需要家长处理")
    with c6:
        metric_card("应收贷款本金", money(m["receivables"], ccy), f"贷款资产占比 {percent(m['loan_asset_ratio'])}")
    with c7:
        metric_card("本月消费", money(m["month_expense"], ccy), f"收入 {money(m['month_income'])}")
    with c8:
        metric_card("本月支出/收入", percent(m["spend_income_ratio"]), "支出纪律")

    st.divider()
    left, right = st.columns([1.05, 1])
    with left:
        st.subheader("AI 本周行动计划")
        for i, a in enumerate(build_weekly_plan(data), 1):
            st.write(f"**{i}.** {a}")

        st.subheader("信用分因子")
        for f in m["score_factors"]:
            st.write(f"- {f}")

    with right:
        st.subheader("本月预算使用率")
        rows = []
        for cat, row in m["budget_usage"].items():
            rows.append({"分类": cat, "已花": row["spent"], "预算": row["limit"], "使用率": row["ratio"]})
        df = pd.DataFrame(rows)
        if df.empty:
            st.info("暂无预算。请到“预算管理”新增。")
        else:
            st.bar_chart(df.set_index("分类")[["使用率"]])
            show = df.copy()
            show["已花"] = show["已花"].map(money)
            show["预算"] = show["预算"].map(money)
            show["使用率"] = show["使用率"].map(percent)
            st.dataframe(show, use_container_width=True, hide_index=True)


def page_add_transaction(data: Dict[str, Any]) -> None:
    st.subheader("新增交易")
    tx_type = st.selectbox("交易类型", ["收入", "消费", "转入储蓄", "储蓄取出", "放贷", "还款", "借入", "偿还负债"], index=0)
    with st.form("add_tx_form"):
        d = st.date_input("日期", value=date.today())
        amount = st.number_input("金额", min_value=0.0, step=1.0, format="%.2f")
        category = st.selectbox("分类", category_options(data))
        party = st.text_input("对象 / 商户 / 借款人", "")
        memo = st.text_area("备注", "")

        expected = 0.0
        principal_repaid = 0.0
        interest_received = 0.0
        due = ""

        if tx_type == "放贷":
            expected = st.number_input("预计回款总额", min_value=0.0, step=1.0, format="%.2f")
            due = st.date_input("预计还款日", value=date.today() + timedelta(days=7)).isoformat()
        if tx_type == "还款":
            principal_repaid = st.number_input("本金回收", min_value=0.0, step=1.0, format="%.2f")
            interest_received = st.number_input("利息收入", min_value=0.0, step=0.5, format="%.2f")
            amount = principal_repaid + interest_received

        submitted = st.form_submit_button("保存交易", type="primary")

    if submitted:
        tx = make_tx(tx_type, amount, category, party=party, memo=memo, tx_date=d.isoformat(), expected_repayment=expected, principal_repaid=principal_repaid, interest_received=interest_received, due_date=due)
        data.setdefault("transactions", []).append(tx)
        commit(data, event=f"新增交易：{tx_type}", reason=memo)
        st.success("交易已保存并同步存储。")
        st.rerun()


def page_ai_center(data: Dict[str, Any]) -> None:
    st.subheader("AI 决策中心")
    st.caption("自动识别分类，给出审批结论和三种情景方案。")

    text = st.text_area("输入请求", value="我想买 $18 的 Minecraft 道具", height=120)
    applicant = st.text_input("申请人", value="阿苏")

    if st.button("AI 分类 + 风控审批", type="primary"):
        amounts = extract_amounts(text)
        is_loan = any(x in text for x in ["借给", "放贷", "贷款给"])
        if is_loan:
            decision = parse_natural_request(data, text)
            st.session_state["latest_ai_decision"] = decision
            st.session_state["latest_ai_classification"] = {"category": "家庭贷款", "necessity": "金融行为", "impulse_level": "中", "merchant": infer_borrower(text), "note": "贷款申请"}
            st.session_state["latest_scenarios"] = []
        else:
            cls = ai_classify_purchase(text)
            amount = amounts[0] if amounts else 0.0
            decision = evaluate_purchase_decision(data, amount, cls["category"], text, cls["merchant"])
            decision["pending_tx"]["memo"] = text
            st.session_state["latest_ai_decision"] = decision
            st.session_state["latest_ai_classification"] = cls
            st.session_state["latest_scenarios"] = scenario_planning(data, amount, cls["category"], text, cls["merchant"])

    if "latest_ai_classification" in st.session_state:
        st.markdown("#### AI 自动分类")
        st.dataframe(pd.DataFrame([st.session_state["latest_ai_classification"]]), use_container_width=True, hide_index=True)

    if st.session_state.get("latest_scenarios"):
        st.markdown("#### 情景规划")
        show = []
        for s in st.session_state["latest_scenarios"]:
            show.append({k: pretty_value(k, v) if isinstance(v, (int, float)) or v is None else v for k, v in s.items()})
        st.dataframe(pd.DataFrame(show), use_container_width=True, hide_index=True)

    if "latest_ai_decision" in st.session_state:
        render_decision(st.session_state["latest_ai_decision"], data, "ai_center")


def page_purchase(data: Dict[str, Any]) -> None:
    st.subheader("消费审批")
    with st.form("purchase_form"):
        c1, c2, c3 = st.columns(3)
        with c1:
            amount = st.number_input("购买金额", value=18.0, min_value=0.0, step=1.0, format="%.2f")
        with c2:
            category = st.selectbox("消费分类", spending_categories(data))
        with c3:
            merchant = st.text_input("商户", "Minecraft 商店")
        desc = st.text_area("购买说明", "Minecraft 道具")
        submitted = st.form_submit_button("模拟消费审批", type="primary")

    if submitted:
        st.session_state["purchase_decision"] = evaluate_purchase_decision(data, amount, category, desc, merchant)
        st.session_state["purchase_scenarios"] = scenario_planning(data, amount, category, desc, merchant)

    if "purchase_scenarios" in st.session_state:
        st.markdown("#### 情景规划")
        show = []
        for s in st.session_state["purchase_scenarios"]:
            show.append({k: pretty_value(k, v) if isinstance(v, (int, float)) or v is None else v for k, v in s.items()})
        st.dataframe(pd.DataFrame(show), use_container_width=True, hide_index=True)

    if "purchase_decision" in st.session_state:
        render_decision(st.session_state["purchase_decision"], data, "purchase")


def page_loan(data: Dict[str, Any]) -> None:
    st.subheader("贷款审批")
    with st.form("loan_form"):
        c1, c2, c3 = st.columns(3)
        with c1:
            principal = st.number_input("拟借出本金", value=20.0, min_value=0.0, step=1.0, format="%.2f")
        with c2:
            borrower = st.text_input("借款人", "爸爸")
        with c3:
            expected = st.number_input("预计回款总额", value=22.0, min_value=0.0, step=1.0, format="%.2f")
        due = st.date_input("预计还款日", value=date.today() + timedelta(days=7))
        memo = st.text_area("备注", "家庭临时周转")
        submitted = st.form_submit_button("模拟贷款审批", type="primary")

    if submitted:
        st.session_state["loan_decision"] = evaluate_loan_decision(data, principal, borrower, expected, due.isoformat(), memo)

    if "loan_decision" in st.session_state:
        render_decision(st.session_state["loan_decision"], data, "loan")


def build_scorecard(data: Dict[str, Any]) -> pd.DataFrame:
    m = calc_financials(data)
    rows = []
    cash_floor = fnum(data.get("settings", {}).get("cash_floor"))
    cash_score = 200 if cash_floor <= 0 or m["cash"] >= cash_floor * 2 else 140 if m["cash"] >= cash_floor else 80 if m["cash"] >= 0 else 20
    rows.append(["现金纪律", cash_score, 200, "现金余额与安全线"])

    usage = m.get("budget_usage", {})
    if usage:
        max_ratio = max((r.get("ratio", 0) for r in usage.values()), default=0)
        budget_score = 200 if max_ratio <= 0.75 else 160 if max_ratio <= 0.90 else 110 if max_ratio <= 1.0 else 50
    else:
        budget_score = 80
    rows.append(["预算纪律", budget_score, 200, "预算使用率与超支情况"])

    savings_score = 150 if m["savings_asset_ratio"] >= 0.30 else 110 if m["savings_asset_ratio"] >= 0.15 else 70 if m["savings"] > 0 else 30
    rows.append(["储蓄纪律", savings_score, 150, "储蓄占总资产比例"])

    loan_ratio = m["loan_asset_ratio"]
    loan_score = 150 if loan_ratio <= 0.25 else 110 if loan_ratio <= 0.45 else 70 if loan_ratio <= 0.65 else 20
    rows.append(["贷款纪律", loan_score, 150, "应收贷款占总资产比例"])

    debt_ratio = m["debt_asset_ratio"]
    debt_score = 100 if debt_ratio == 0 else 80 if debt_ratio <= 0.15 else 50 if debt_ratio <= 0.35 else 10
    rows.append(["负债纪律", debt_score, 100, "负债占总资产比例"])

    stability_score = 100 if len(data.get("score_history", [])) >= 5 else 70 if len(data.get("transactions", [])) >= 3 else 40
    rows.append(["稳定性", stability_score, 100, "记录连续性与系统使用频率"])

    df = pd.DataFrame(rows, columns=["项目", "得分", "满分", "说明"])
    df["完成度"] = df["得分"] / df["满分"]
    return df


def page_child_dashboard(data: Dict[str, Any]) -> None:
    st.subheader("儿童首页")
    st.caption("阿苏常用入口：申请购买、看信用分、看奖励、看审批结果。")
    m = calc_financials(data)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("我的现金", money(m["cash"]))
    c2.metric("我的信用分", score_text(m["credit_score"]))
    c3.metric("待审批", len([r for r in data.get("pending_requests", []) if r.get("parent_status") == "待审批"]))
    c4.metric("徽章", len(data.get("badges", [])))

    st.markdown("#### 我要买东西")
    text = st.text_area("把想买的东西写在这里", value="我想买 $5 的甜品", key="child_request_text")
    if st.button("提交购买申请", type="primary"):
        cls = ai_classify_purchase(text)
        amounts = extract_amounts(text)
        amount = amounts[0] if amounts else 0.0
        decision = evaluate_purchase_decision(data, amount, cls["category"], text, cls["merchant"])
        tx = decision.get("pending_tx")
        req = {
            "id": uid(),
            "created_at": now_str(),
            "request_text": text,
            "request_type": "消费",
            "amount": amount,
            "category": cls["category"],
            "merchant": cls["merchant"],
            "applicant": current_operator(),
            "ai_category": cls["category"],
            "necessity": cls["necessity"],
            "impulse_level": cls["impulse_level"],
            "python_decision": decision.get("result", ""),
            "parent_status": "待审批",
            "final_status": "未入账",
            "decision_json": decision,
            "pending_tx": tx,
            "parent_note": "",
            "approved_by": "",
            "approved_at": "",
            "booked_at": "",
        }
        data.setdefault("pending_requests", []).append(req)
        commit(data, event="儿童首页提交申请", reason=text)
        st.success("申请已提交给爸爸/妈妈审批。")
        st.rerun()

    st.markdown("#### 我的徽章")
    if not data.get("badges"):
        st.info("还没有徽章。先记录收入、建立预算或提交一次合规申请。")
    else:
        st.dataframe(pd.DataFrame(data.get("badges", []))[['name','earned_at','description']], use_container_width=True, hide_index=True)

    st.markdown("#### 我的待审批")
    mine = [r for r in data.get("pending_requests", []) if r.get("applicant") == current_operator()]
    if not mine:
        st.info("暂无申请。")
    else:
        st.dataframe(pd.DataFrame(mine)[["created_at","amount","category","parent_status","final_status","request_text"]], use_container_width=True, hide_index=True)


def page_rules(data: Dict[str, Any]) -> None:
    st.subheader("风控规则")
    st.caption("这些规则写入 state 主数据，并同步到 Google Sheet 的 rules 展示页。")
    rules = data.setdefault("rules", empty_data()["rules"])
    if not can_admin():
        st.warning("只有爸爸管理员可以修改风控规则。")
        st.dataframe(pd.DataFrame([{"规则": k, "当前值": v} for k, v in rules.items()]), use_container_width=True, hide_index=True)
        return
    with st.form("rules_form"):
        new_rules = {}
        labels = {
            "budget_warning_line": "预算预警线",
            "budget_reject_line": "预算拒绝线",
            "score_reject_line": "信用分拒绝线",
            "approval_threshold": "家长审批金额线",
            "loan_asset_soft_limit": "贷款资产建议线",
            "loan_asset_hard_limit": "贷款资产硬红线",
            "cash_floor_default": "默认现金安全线",
            "weekly_no_impulse_reward": "周度无冲动消费奖励分",
            "monthly_budget_reward": "月度预算纪律奖励分",
        }
        for k, default in empty_data()["rules"].items():
            new_rules[k] = st.number_input(labels.get(k, k), value=fnum(rules.get(k, default)), step=0.05 if "line" in k or "limit" in k else 1.0, format="%.2f")
        submitted = st.form_submit_button("保存规则", type="primary")
    if submitted:
        data["rules"] = new_rules
        commit(data, event="保存风控规则", reason="rules updated")
        st.success("规则已保存。")
        st.rerun()


def page_backups(data: Dict[str, Any]) -> None:
    st.subheader("备份与恢复")
    st.caption("每次保存都会自动备份上一个版本。只保留最近 40 次。")
    backups = data.get("backups", [])
    if not backups:
        st.info("暂无备份。")
        return
    df = pd.DataFrame([{k: b.get(k) for k in ["id","time","operator","event","reason"]} for b in backups]).sort_values("time", ascending=False)
    st.dataframe(df, use_container_width=True, hide_index=True)
    if not can_admin():
        st.warning("只有爸爸管理员可以恢复备份。")
        return
    options = [f"{b.get('time')} · {b.get('operator')} · {b.get('event')} · {b.get('id')}" for b in sorted(backups, key=lambda x: str(x.get('time','')), reverse=True)]
    selected = st.selectbox("选择要恢复的备份", options)
    selected_id = selected.split(" · ")[-1]
    if st.button("恢复到这个版本", type="primary"):
        backup = next((b for b in backups if b.get("id") == selected_id), None)
        if backup:
            restored = normalize_data(json.loads(backup.get("snapshot_json") or "{}"))
            restored["backups"] = backups
            commit(restored, event="恢复备份", reason=selected_id)
            st.success("已恢复备份。")
            st.rerun()


def page_rewards(data: Dict[str, Any]) -> None:
    st.subheader("奖励与徽章")
    evaluate_badges(data)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### 徽章")
        if not data.get("badges"):
            st.info("暂无徽章。")
        else:
            st.dataframe(pd.DataFrame(data.get("badges", []))[["name","earned_at","description"]], use_container_width=True, hide_index=True)
    with c2:
        st.markdown("#### 奖励券")
        if not data.get("rewards"):
            st.info("暂无奖励券。")
        else:
            st.dataframe(pd.DataFrame(data.get("rewards", []))[["title","created_at","created_by","status","note"]], use_container_width=True, hide_index=True)
    st.divider()
    if can_approve():
        with st.form("reward_form"):
            title = st.text_input("奖励名称", "周末小额自由消费券")
            note = st.text_input("说明", "用于奖励预算纪律或储蓄进度")
            submitted = st.form_submit_button("发放奖励", type="primary")
        if submitted and title.strip():
            data.setdefault("rewards", []).append({"id": uid(), "title": title.strip(), "created_at": now_str(), "created_by": current_operator(), "status": "可用", "note": note})
            commit(data, event="发放奖励", reason=title)
            st.success("奖励已发放。")
            st.rerun()


def page_pending(data: Dict[str, Any]) -> None:
    st.subheader("待审批队列")
    st.caption(approval_hint())

    pending = data.get("pending_requests", [])
    if not pending:
        st.info("暂无待审批申请。")
        return

    for r in sorted(pending, key=lambda x: str(x.get("created_at", "")), reverse=True):
        css = decision_class(r.get("parent_status", "待审批"))
        approved_by = r.get("approved_by") or "尚未审批"
        st.markdown(
            f"""
            <div class="decision {css}">
                <div class="pill">{r.get('parent_status')}</div>
                <h3>{r.get('request_type')}：{money(r.get('amount'))} · {r.get('category')}</h3>
                <p class="small">{r.get('request_text')}</p>
                <p class="small">申请人：{r.get('applicant', '')}；审批人：{approved_by}</p>
                <p class="small">系统结论：{r.get('python_decision')}；AI分类：{r.get('ai_category')}；必要性：{r.get('necessity')}；冲动等级：{r.get('impulse_level')}</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        c1, c2, c3 = st.columns([1, 1, 2])
        if r.get("parent_status") == "待审批":
            if can_approve():
                with c1:
                    if st.button(f"{current_operator()}批准并入账", key=f"approve_{r['id']}", type="primary"):
                        tx = r.get("pending_tx")
                        if tx:
                            data.setdefault("transactions", []).append(tx)
                        r["parent_status"] = "已批准"
                        r["final_status"] = "已入账"
                        r["approved_by"] = current_operator()
                        r["approved_at"] = now_str()
                        r["booked_at"] = now_str()
                        commit(data, event=f"{current_operator()}批准入账", reason=r.get("request_text", ""))
                        st.success(f"{current_operator()}已批准并入账。")
                        st.rerun()
                with c2:
                    if st.button(f"{current_operator()}拒绝", key=f"reject_{r['id']}"):
                        r["parent_status"] = "已拒绝"
                        r["final_status"] = "未入账"
                        r["approved_by"] = current_operator()
                        r["approved_at"] = now_str()
                        commit(data, event=f"{current_operator()}拒绝申请", reason=r.get("request_text", ""))
                        st.warning(f"{current_operator()}已拒绝。")
                        st.rerun()
                with c3:
                    note = st.text_input("家长备注", value=r.get("parent_note", ""), key=f"note_{r['id']}")
                    if st.button("保存备注", key=f"save_note_{r['id']}"):
                        r["parent_note"] = note
                        r["approved_by"] = current_operator()
                        commit(data, event=f"{current_operator()}保存审批备注", reason=note)
                        st.success("备注已保存。")
                        st.rerun()
            else:
                st.info("当前操作人没有审批权限。请切换为爸爸或妈妈。")


def page_risk_radar(data: Dict[str, Any]) -> None:
    st.subheader("风险雷达")
    risks = build_risk_radar(data)
    cols = st.columns(3)
    for col, level in zip(cols, ["高风险", "中风险", "低风险"]):
        with col:
            st.markdown(f"### {level}")
            if not risks[level]:
                st.info("暂无")
            for item in risks[level]:
                css = "bad" if level == "高风险" else "warn" if level == "中风险" else "ok"
                st.markdown(f"<div class='decision {css}'><b>{item['title']}</b><p class='small'>{item['detail']}</p></div>", unsafe_allow_html=True)
    st.divider()
    st.subheader("本周行动指令")
    for i, a in enumerate(build_weekly_plan(data), 1):
        st.write(f"**{i}.** {a}")


def transactions_df(data: Dict[str, Any]) -> pd.DataFrame:
    rows = []
    for tx in data.get("transactions", []):
        rows.append({
            "日期": tx.get("date"),
            "类型": tx.get("type"),
            "金额": fnum(tx.get("amount")),
            "分类": tx.get("category"),
            "对象/商户": tx.get("party"),
            "本金回收": fnum(tx.get("principal_repaid")),
            "利息收入": fnum(tx.get("interest_received")),
            "预计回款": fnum(tx.get("expected_repayment")),
            "到期日": tx.get("due_date"),
            "备注": tx.get("memo"),
            "id": tx.get("id"),
        })
    if not rows:
        return pd.DataFrame()
    return pd.DataFrame(rows).sort_values("日期", ascending=False)


def page_monthly_statement(data: Dict[str, Any]) -> None:
    st.subheader("月度账单")
    df = transactions_df(data)
    if df.empty:
        st.info("暂无交易。")
        return
    months = sorted(df["日期"].astype(str).str[:7].unique(), reverse=True)
    selected = st.selectbox("选择月份", months)
    m = calc_financials(data, selected)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("本月收入", money(m["month_income"]))
    c2.metric("本月消费", money(m["month_expense"]))
    c3.metric("支出/收入比", percent(m["spend_income_ratio"]))
    c4.metric("信用分", score_text(m["credit_score"]))
    month_df = df[df["日期"].astype(str).str[:7] == selected].copy()
    st.dataframe(month_df.drop(columns=["id"], errors="ignore"), use_container_width=True, hide_index=True)


def page_budget(data: Dict[str, Any]) -> None:
    st.subheader("预算管理")
    with st.form("budget_form"):
        new_budgets = {}
        if data.get("budgets"):
            for cat, limit in data["budgets"].items():
                new_budgets[cat] = st.number_input(f"{cat} 月度预算", value=fnum(limit), min_value=0.0, step=1.0, format="%.2f", key=f"budget_{cat}")
        else:
            st.info("暂无预算分类。请新增。")
        st.markdown("#### 新增分类")
        c1, c2 = st.columns([2, 1])
        with c1:
            new_cat = st.text_input("分类名称", "")
        with c2:
            new_limit = st.number_input("预算金额", min_value=0.0, step=1.0, format="%.2f")
        submitted = st.form_submit_button("保存预算", type="primary")
    if submitted:
        if new_cat.strip():
            new_budgets[new_cat.strip()] = new_limit
        data["budgets"] = new_budgets
        commit(data, event="保存预算", reason="预算更新")
        st.success("预算已保存。")
        st.rerun()
    m = calc_financials(data)
    rows = [{"分类": cat, "已花": money(row["spent"]), "预算": money(row["limit"]), "使用率": percent(row["ratio"])} for cat, row in m["budget_usage"].items()]
    if rows:
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


def page_goals(data: Dict[str, Any]) -> None:
    st.subheader("储蓄目标")
    goals = calc_goals(data)
    if not goals:
        st.info("暂无储蓄目标。")
    for g in goals:
        st.markdown(f"#### {g['name']}")
        st.progress(min(1.0, g["progress"]))
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("目标", money(g["target"]))
        c2.metric("当前", money(g["current"]))
        c3.metric("完成度", percent(g["progress"]))
        c4.metric("剩余", money(g["remaining"]))
        if g["days_left"] is not None:
            st.caption(f"截止日：{g['deadline']}；剩余 {g['days_left']} 天")
        st.divider()

    with st.form("goal_form"):
        st.markdown("#### 新增储蓄目标")
        name = st.text_input("目标名称", "")
        c1, c2, c3 = st.columns(3)
        with c1:
            target = st.number_input("目标金额", min_value=0.0, step=5.0, format="%.2f")
        with c2:
            current = st.number_input("当前金额", min_value=0.0, step=5.0, format="%.2f")
        with c3:
            category = st.selectbox("关联分类", spending_categories(data))
        deadline = st.date_input("截止日", value=date.today() + timedelta(days=90))
        submitted = st.form_submit_button("新增目标", type="primary")
    if submitted and name.strip():
        data.setdefault("goals", []).append({"id": uid(), "name": name.strip(), "target": target, "current": current, "deadline": deadline.isoformat(), "category": category})
        commit(data, event="新增储蓄目标", reason=name)
        st.success("目标已新增。")
        st.rerun()


def page_score(data: Dict[str, Any]) -> None:
    st.subheader("信用分")
    m = calc_financials(data)
    st.metric("当前信用分", score_text(m["credit_score"]))
    st.caption("家庭内部风控评分，不是 FICO，不是银行真实征信。")
    st.markdown("#### 当前因子")
    for f in m["score_factors"]:
        st.write(f"- {f}")
    st.divider()
    st.markdown("#### 信用评分卡")
    card = build_scorecard(data)
    st.dataframe(card, use_container_width=True, hide_index=True)
    st.bar_chart(card.set_index("项目")[["完成度"]])
    st.divider()
    st.markdown("#### 信用分历史")
    hist = pd.DataFrame(data.get("score_history", []))
    if hist.empty:
        st.info("暂无信用分历史。")
    else:
        st.dataframe(hist.sort_values("time", ascending=False), use_container_width=True, hide_index=True)


def page_merchants(data: Dict[str, Any]) -> None:
    st.subheader("商户权益")
    if not data.get("merchants"):
        st.info("暂无商户权益。")
    for merchant in data.get("merchants", []):
        r = evaluate_merchant_access(data, merchant)
        css = decision_class(r["result"])
        failed = "无" if not r["failed"] else "；".join(r["failed"])
        st.markdown(
            f"<div class='decision {css}'><div class='pill'>{merchant.get('category')}</div><h3>{merchant.get('name')}</h3><p><b>{r['message']}</b></p><p class='small'>暂缓/失败原因：{failed}</p><p class='small'>{merchant.get('note','')}</p></div>",
            unsafe_allow_html=True,
        )
    st.divider()
    with st.form("merchant_form"):
        st.markdown("#### 新增商户")
        name = st.text_input("商户名称", "")
        category = st.selectbox("商户分类", spending_categories(data))
        c1, c2, c3 = st.columns(3)
        with c1:
            required = st.number_input("最低信用分", min_value=0, max_value=850, value=0)
        with c2:
            discount = st.number_input("折扣比例", min_value=0.0, max_value=0.9, value=0.08, step=0.01)
        with c3:
            cap = st.number_input("品类预算开放上限", min_value=0.0, max_value=2.0, value=0.90, step=0.05)
        note = st.text_input("说明", "")
        submitted = st.form_submit_button("新增商户", type="primary")
    if submitted and name.strip():
        data.setdefault("merchants", []).append({"id": uid(), "name": name.strip(), "category": category, "discount": discount, "required_score": int(required), "category_budget_cap": cap, "note": note})
        commit(data, event="新增商户", reason=name)
        st.success("商户已新增。")
        st.rerun()


def page_weekly(data: Dict[str, Any]) -> None:
    st.subheader("自动周报")
    if st.button("生成本周周报", type="primary"):
        report = generate_weekly_report(data)
        data.setdefault("weekly_reports", []).append(report)
        data["weekly_reports"] = data["weekly_reports"][-52:]
        commit(data, event="生成周报", reason=report["week_start"])
        st.success("周报已生成并同步。")
        st.rerun()
    if not data.get("weekly_reports"):
        st.info("暂无周报。")
    for r in sorted(data.get("weekly_reports", []), key=lambda x: x.get("week_start", ""), reverse=True):
        with st.expander(f"{r.get('week_start')} 周报", expanded=False):
            st.text(r.get("content", ""))


def page_transactions(data: Dict[str, Any]) -> None:
    st.subheader("交易流水")
    df = transactions_df(data)
    if df.empty:
        st.info("暂无交易。")
    else:
        st.dataframe(df.drop(columns=["id"], errors="ignore"), use_container_width=True, hide_index=True)
        csv_text = df.drop(columns=["id"], errors="ignore").to_csv(index=False, encoding="utf-8-sig")
        st.download_button("下载交易 CSV", csv_text.encode("utf-8-sig"), "asu_money2_transactions.csv", "text/csv")
    json_text = json.dumps(normalize_data(data), ensure_ascii=False, indent=2)
    st.download_button("下载完整 JSON 备份", json_text.encode("utf-8"), "asu_money2_backup.json", "application/json")
    st.divider()
    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("重新从存储读取"):
            reload_data()
            st.rerun()
    with c2:
        if st.button("手动刷新 Google Sheet 展示页"):
            try:
                sync_readable_sheets(data)
                st.success("展示页已刷新。")
            except Exception as e:
                st.error(f"刷新失败：{e}")
    with c3:
        confirm = st.checkbox("确认清空所有数据")
        if st.button("清空为空库", disabled=not confirm):
            reset_empty()
            st.success("已清空。")
            st.rerun()


def render_sidebar(data: Dict[str, Any]) -> None:
    st.sidebar.title("系统设置")

    if configured_passwords():
        st.sidebar.success(f"已登录：{current_operator()}")
        if st.sidebar.button("退出登录"):
            logout()
    else:
        operators = ["爸爸", "妈妈", "阿苏"]
        current = st.session_state.get("current_operator", "爸爸")
        idx = operators.index(current) if current in operators else 0
        st.sidebar.selectbox(
            "当前操作人",
            operators,
            index=idx,
            key="current_operator",
            help="未配置密码时可手动切换；正式使用建议配置 DAD_PASSWORD / MOM_PASSWORD / ASU_PASSWORD。"
        )
        st.sidebar.warning("未启用密码登录，当前身份可手动切换。")

    if can_admin():
        st.sidebar.success("爸爸管理员：可改设置、规则、备份恢复")
    elif can_approve():
        st.sidebar.success(f"{current_operator()}拥有审批权限")
    else:
        st.sidebar.info("阿苏只能提交申请，不能批准入账")

    s = data.setdefault("settings", {})
    if not can_admin():
        st.sidebar.caption("系统设置仅爸爸管理员可修改。")
        return
    with st.sidebar.form("settings_form"):
        owner = st.text_input("账户名称", s.get("owner", "阿苏"))
        start_cash = st.number_input("初始现金", value=fnum(s.get("start_cash")), step=1.0, format="%.2f")
        start_savings = st.number_input("初始储蓄", value=fnum(s.get("start_savings")), step=1.0, format="%.2f")
        base_score = st.number_input("基础信用分", min_value=0, max_value=850, value=inum(s.get("base_score"), 0))
        cash_floor = st.number_input("最低现金安全线", value=fnum(s.get("cash_floor")), step=1.0, format="%.2f")
        approval_threshold = st.number_input("家长审批金额线", value=fnum(s.get("approval_threshold"), 15.0), min_value=0.0, step=1.0, format="%.2f")
        loan_limit = st.number_input("贷款资产建议上限", value=fnum(s.get("loan_asset_limit"), 0.45), min_value=0.0, max_value=1.0, step=0.05)
        hard_loan_limit = st.number_input("贷款资产硬红线", value=fnum(s.get("hard_loan_asset_limit"), 0.65), min_value=0.0, max_value=1.0, step=0.05)
        submitted = st.form_submit_button("保存设置", type="primary")
    if submitted:
        s["owner"] = owner
        s["start_cash"] = start_cash
        s["start_savings"] = start_savings
        s["base_score"] = int(base_score)
        s["cash_floor"] = cash_floor
        s["approval_threshold"] = approval_threshold
        s["loan_asset_limit"] = loan_limit
        s["hard_loan_asset_limit"] = hard_loan_limit
        commit(data, event="保存设置", reason="系统设置更新")
        st.sidebar.success("设置已保存。")
        st.rerun()

    st.sidebar.divider()
    st.sidebar.caption("连接状态")
    if gsheet_enabled():
        st.sidebar.success(f"Google Sheet：{st.secrets.get('SHEET_NAME', '')}")
    else:
        st.sidebar.warning("Google Sheet 未启用，使用本地 JSON")
    if secret_value("DEEPSEEK_API_KEY"):
        st.sidebar.success(f"DeepSeek：{secret_value('DEEPSEEK_MODEL', 'deepseek-v4-flash')}")
    else:
        st.sidebar.info("DeepSeek 未配置，使用本地报告")


def main() -> None:
    st.set_page_config(page_title=APP_NAME, page_icon="🏦", layout="wide", initial_sidebar_state="expanded")
    inject_css()
    require_login()
    data = get_data()
    render_sidebar(data)
    render_hero(data)

    tabs = st.tabs([
        "1 儿童首页",
        "2 首页总览",
        "3 新增交易",
        "4 AI 决策中心",
        "5 消费审批",
        "6 贷款审批",
        "7 待审批队列",
        "8 风险雷达",
        "9 月度账单",
        "10 预算管理",
        "11 储蓄目标",
        "12 信用分",
        "13 商户权益",
        "14 周报",
        "15 奖励徽章",
        "16 风控规则",
        "17 备份恢复",
        "18 交易流水",
    ])

    with tabs[0]:
        page_child_dashboard(data)
    with tabs[1]:
        page_home(data)
    with tabs[2]:
        page_add_transaction(data)
    with tabs[3]:
        page_ai_center(data)
    with tabs[4]:
        page_purchase(data)
    with tabs[5]:
        page_loan(data)
    with tabs[6]:
        page_pending(data)
    with tabs[7]:
        page_risk_radar(data)
    with tabs[8]:
        page_monthly_statement(data)
    with tabs[9]:
        page_budget(data)
    with tabs[10]:
        page_goals(data)
    with tabs[11]:
        page_score(data)
    with tabs[12]:
        page_merchants(data)
    with tabs[13]:
        page_weekly(data)
    with tabs[14]:
        page_rewards(data)
    with tabs[15]:
        page_rules(data)
    with tabs[16]:
        page_backups(data)
    with tabs[17]:
        page_transactions(data)


if __name__ == "__main__":
    main()
