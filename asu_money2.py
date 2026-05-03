# -*- coding: utf-8 -*-
"""
阿苏私人银行 3.0
2.4 教育闭环版 + 3.0 家庭金融操作系统合并版

核心结构：
    1. 首页仪表盘
    2. 儿童模式
    3. 家长驾驶舱
    4. 审批与交易
    5. 成长与课程
    6. 报表与数据
    7. 系统设置

Google Sheet：
    state = 主数据库
    其他页签 = 自动生成展示页

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

登录可选：
    DAD_PASSWORD = "xxx"
    MOM_PASSWORD = "xxx"
    ASU_PASSWORD = "xxx"
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


APP_NAME = "阿苏私人银行 3.0"
SCHEMA_VERSION = "asu-bank-3-0"
DATA_PATH = Path("asu_money2_data.json")
SESSION_KEY = "asu_bank_3_0_state"
STATE_WS = "state"
STATE_KEY = "bank_data"

GSHEET_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# ============================================================
# 1. 基础数据
# ============================================================

def now_str() -> str:
    return datetime.now().isoformat(timespec="seconds")


def today_str() -> str:
    return date.today().isoformat()


def month_str(d: Optional[date] = None) -> str:
    return (d or date.today()).strftime("%Y-%m")


def week_start(d: Optional[date] = None) -> date:
    x = d or date.today()
    return x - timedelta(days=x.weekday())


def uid() -> str:
    return str(uuid.uuid4())


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
            "approval_threshold": 15.0,
            "loan_asset_limit": 0.45,
            "hard_loan_asset_limit": 0.65,
            "cooldown_game_hours": 72,
            "cooldown_treat_hours": 24,
            "cooldown_toy_hours": 168,
            "cooldown_default_hours": 48,
            "savings_rate_weekly": 0.0015,
            "loan_rate_weekly": 0.02,
            "late_fee_weekly": 0.05,
            "created_at": now_str(),
        },
        "rules": {
            "budget_warning_line": 0.90,
            "budget_reject_line": 1.20,
            "credit_reject_line": 650,
            "approval_threshold": 15.0,
            "loan_asset_soft_line": 0.45,
            "loan_asset_hard_line": 0.65,
            "cooldown_reward_points": 3,
            "review_reward_points": 2,
            "task_default_reward": 1.0,
        },
        "budgets": {},
        "goals": [],
        "merchants": [],
        "transactions": [],
        "pending_requests": [],
        "score_history": [],
        "weekly_reports": [],
        "monthly_cfo_reports": [],
        "wishlist": [],
        "post_purchase_reviews": [],
        "tasks": [],
        "badges": [],
        "rewards": [],
        "backups": [],
        "accounts": {
            "cash": {"name": "现金账户", "target_ratio": 0.40},
            "savings": {"name": "储蓄账户", "target_ratio": 0.30},
            "goal_fund": {"name": "目标基金", "target_ratio": 0.20},
            "charity": {"name": "慈善账户", "target_ratio": 0.05},
            "risk_fund": {"name": "风险投资账户", "target_ratio": 0.05},
        },
        "curriculum_progress": {},
    }


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


def decision_class(result: str) -> str:
    if result in {"批准", "通过", "开放", "已批准", "已入账", "完成", "已发放"}:
        return "ok"
    if result in {"延迟", "限额通过", "暂缓开放", "观察", "待家长审批", "待审批", "未建立", "冷静期"}:
        return "warn"
    return "bad"


def pretty_value(key: str, value: Any) -> str:
    if value is None:
        return "未建立"
    if isinstance(value, (int, float)):
        if ("率" in key or "占比" in key or "使用率" in key or "上限" in key or "比例" in key) and abs(float(value)) <= 3:
            return percent(value)
        if "信用分" in key or key in {"score"}:
            return f"{value:.0f}"
        return money(value)
    return str(value)


def secret_value(key: str, default: str = "") -> str:
    try:
        if key in st.secrets:
            return str(st.secrets[key])
    except Exception:
        pass
    return os.getenv(key, default)



def has_activity(data: Dict[str, Any]) -> bool:
    """判断账户是否已经有真实活动。"""
    s = data.get("settings", {})
    return (
        fnum(s.get("start_cash")) != 0
        or fnum(s.get("start_savings")) != 0
        or bool(data.get("transactions"))
        or bool(data.get("budgets"))
        or bool(data.get("goals"))
        or bool(data.get("merchants"))
        or bool(data.get("pending_requests"))
        or bool(data.get("wishlist"))
        or bool(data.get("tasks"))
        or bool(data.get("post_purchase_reviews"))
    )


def category_options(data: Dict[str, Any]) -> List[str]:
    return sorted(set(list(data.get("budgets", {}).keys()) + ["零花钱", "家庭贷款", "游戏", "甜品", "学习", "宠物", "玩具", "餐饮", "其他"]))


def spending_categories(data: Dict[str, Any]) -> List[str]:
    cats = sorted(set(list(data.get("budgets", {}).keys()) + ["游戏", "甜品", "学习", "宠物", "玩具", "餐饮", "其他"]))
    return cats or ["其他"]


def is_parent() -> bool:
    return st.session_state.get("role") in {"爸爸", "妈妈"}


def is_admin() -> bool:
    return st.session_state.get("role") == "爸爸"


def current_actor() -> str:
    return st.session_state.get("role", "阿苏")


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
    for k in ["start_cash", "start_savings", "base_score", "cash_floor", "approval_threshold", "loan_asset_limit", "hard_loan_asset_limit", "cooldown_game_hours", "cooldown_treat_hours", "cooldown_toy_hours", "cooldown_default_hours", "savings_rate_weekly", "loan_rate_weekly", "late_fee_weekly"]:
        if k in settings:
            settings[k] = fnum(settings[k]) if "hours" not in k and k != "base_score" else inum(settings[k], int(fnum(settings[k])))
    settings["base_score"] = inum(settings.get("base_score"), 0)
    settings["owner"] = settings.get("owner") or "阿苏"
    settings["currency"] = settings.get("currency") or "USD"
    base["settings"] = settings

    rules = base["rules"]
    if isinstance(raw.get("rules"), dict):
        rules.update(raw["rules"])
    for k, v in list(rules.items()):
        rules[k] = fnum(v)
    base["rules"] = rules

    budgets = {}
    for k, v in (raw.get("budgets") or {}).items():
        name = str(k).strip()
        if name:
            budgets[name] = fnum(v)
    base["budgets"] = budgets

    def list_of_dicts(name: str) -> List[Dict[str, Any]]:
        return [x for x in (raw.get(name) or []) if isinstance(x, dict)]

    goals = []
    for g in list_of_dicts("goals"):
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
    for m in list_of_dicts("merchants"):
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
    for tx in list_of_dicts("transactions"):
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
            "account": tx.get("account") or "cash",
        })
    base["transactions"] = txs

    pending = []
    for r in list_of_dicts("pending_requests"):
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
            "approver": r.get("approver") or "",
            "approved_at": r.get("approved_at") or "",
            "booked_at": r.get("booked_at") or "",
        })
    base["pending_requests"] = pending

    wishlist = []
    for w in list_of_dicts("wishlist"):
        wishlist.append({
            "id": w.get("id") or uid(),
            "created_at": w.get("created_at") or now_str(),
            "item": w.get("item") or "",
            "amount": fnum(w.get("amount")),
            "category": w.get("category") or "其他",
            "merchant": w.get("merchant") or "",
            "cooldown_until": w.get("cooldown_until") or now_str(),
            "status": w.get("status") or "冷静期",
            "applicant": w.get("applicant") or "阿苏",
            "final_decision": w.get("final_decision") or "",
        })
    base["wishlist"] = wishlist

    reviews = []
    for r in list_of_dicts("post_purchase_reviews"):
        reviews.append({
            "id": r.get("id") or uid(),
            "transaction_id": r.get("transaction_id") or "",
            "review_time": r.get("review_time") or now_str(),
            "happiness_score": inum(r.get("happiness_score"), 0),
            "still_using": r.get("still_using") or "",
            "would_buy_again": r.get("would_buy_again") or "",
            "lesson": r.get("lesson") or "",
            "reward_given": bool(r.get("reward_given", False)),
        })
    base["post_purchase_reviews"] = reviews

    tasks = []
    for t in list_of_dicts("tasks"):
        tasks.append({
            "id": t.get("id") or uid(),
            "created_at": t.get("created_at") or now_str(),
            "task_name": t.get("task_name") or "",
            "reward_amount": fnum(t.get("reward_amount")),
            "category": t.get("category") or "任务收入",
            "status": t.get("status") or "待完成",
            "submitted_at": t.get("submitted_at") or "",
            "approved_by": t.get("approved_by") or "",
            "booked_tx_id": t.get("booked_tx_id") or "",
        })
    base["tasks"] = tasks

    simple_lists = ["score_history", "weekly_reports", "monthly_cfo_reports", "badges", "rewards", "backups"]
    for name in simple_lists:
        base[name] = list_of_dicts(name)

    if isinstance(raw.get("accounts"), dict):
        for k, v in raw["accounts"].items():
            if isinstance(v, dict):
                base["accounts"][k] = {
                    "name": v.get("name") or k,
                    "target_ratio": fnum(v.get("target_ratio"), base["accounts"].get(k, {}).get("target_ratio", 0.0)),
                }

    if isinstance(raw.get("curriculum_progress"), dict):
        base["curriculum_progress"].update(raw["curriculum_progress"])

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
    creds = Credentials.from_service_account_info(dict(st.secrets["gcp_service_account"]), scopes=GSHEET_SCOPES)
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
    if header:
        ws.update(f"A1:{col_letters(len(header))}{len(values)}", values, value_input_option="RAW")


def sync_readable_sheets(data: Dict[str, Any]) -> None:
    if not gsheet_enabled():
        return

    data = normalize_data(data)
    sh = get_spreadsheet()
    metrics = calc_financials(data)

    set_table(get_or_create_ws(sh, "summary"), ["指标", "数值", "说明"], [
        ["更新时间", now_str(), "程序自动写入"],
        ["总资产", metrics["total_assets"], "现金 + 储蓄 + 应收贷款本金"],
        ["净资产", metrics["net_assets"], "总资产 - 负债"],
        ["现金余额", metrics["cash"], "可立即使用资金"],
        ["储蓄余额", metrics["savings"], "储蓄账户余额"],
        ["应收贷款本金", metrics["receivables"], "未收回本金"],
        ["负债余额", metrics["liabilities"], "未偿还借入资金"],
        ["信用分", score_text(metrics["credit_score"]), "家庭内部风控评分"],
        ["本月收入", metrics["month_income"], month_str()],
        ["本月消费", metrics["month_expense"], month_str()],
        ["待审批数量", len([r for r in data.get("pending_requests", []) if r.get("parent_status") == "待审批"]), "pending_requests"],
        ["愿望清单数量", len([w for w in data.get("wishlist", []) if w.get("status") in {"冷静期", "可申请"}]), "wishlist"],
    ])

    tx_rows = []
    for tx in sorted(data.get("transactions", []), key=lambda x: str(x.get("date", "")), reverse=True):
        tx_rows.append([tx.get("date", ""), tx.get("type", ""), tx.get("amount", 0), tx.get("category", ""), tx.get("party", ""), tx.get("account", ""), tx.get("principal_repaid", 0), tx.get("interest_received", 0), tx.get("expected_repayment", 0), tx.get("due_date", ""), tx.get("memo", "")])
    set_table(get_or_create_ws(sh, "transactions", rows=max(100, len(tx_rows) + 5), cols=11), ["日期", "类型", "金额", "分类", "对象/商户", "账户", "本金回收", "利息收入", "预计回款", "到期日", "备注"], tx_rows)

    budget_rows = []
    usage = metrics["budget_usage"]
    for cat, limit in data.get("budgets", {}).items():
        u = usage.get(cat, {"spent": 0, "ratio": 0})
        budget_rows.append([cat, limit, u["spent"], u["ratio"]])
    set_table(get_or_create_ws(sh, "budgets"), ["分类", "月度预算", "本月已花", "使用率"], budget_rows)

    goal_rows = [[g["name"], g["category"], g["target"], g["current"], g["remaining"], g["progress"], g.get("deadline", ""), g.get("days_left", "")] for g in metrics["goals"]]
    set_table(get_or_create_ws(sh, "goals"), ["目标", "分类", "目标金额", "当前金额", "剩余金额", "完成度", "截止日", "剩余天数"], goal_rows)

    merchant_rows = []
    for m in data.get("merchants", []):
        res = evaluate_merchant_access(data, m)
        merchant_rows.append([m["name"], m["category"], m["discount"], m["required_score"], m["category_budget_cap"], res["result"], "；".join(res["failed"]), m.get("note", "")])
    set_table(get_or_create_ws(sh, "merchants"), ["商户", "分类", "折扣", "最低信用分", "品类预算上限", "状态", "失败原因", "说明"], merchant_rows)

    pending_rows = []
    for r in sorted(data.get("pending_requests", []), key=lambda x: str(x.get("created_at", "")), reverse=True):
        pending_rows.append([r.get("created_at", ""), r.get("applicant", ""), r.get("request_type", ""), r.get("amount", 0), r.get("category", ""), r.get("merchant", ""), r.get("ai_category", ""), r.get("necessity", ""), r.get("impulse_level", ""), r.get("python_decision", ""), r.get("parent_status", ""), r.get("final_status", ""), r.get("approver", ""), r.get("request_text", ""), r.get("parent_note", "")])
    set_table(get_or_create_ws(sh, "pending_requests", rows=max(100, len(pending_rows)+5), cols=15), ["申请时间", "申请人", "类型", "金额", "Python分类", "商户", "AI分类", "必要性", "冲动等级", "系统结论", "家长状态", "最终状态", "审批人", "原始申请", "家长备注"], pending_rows)

    set_table(get_or_create_ws(sh, "wishlist"), ["创建时间", "物品", "金额", "分类", "商户", "冷静期结束", "状态", "申请人", "最终决定"], [[w.get("created_at", ""), w.get("item", ""), w.get("amount", 0), w.get("category", ""), w.get("merchant", ""), w.get("cooldown_until", ""), w.get("status", ""), w.get("applicant", ""), w.get("final_decision", "")] for w in data.get("wishlist", [])])

    set_table(get_or_create_ws(sh, "post_purchase_reviews"), ["复盘时间", "交易ID", "满意度", "是否还在用", "还会再买吗", "学到什么", "奖励已发"], [[r.get("review_time", ""), r.get("transaction_id", ""), r.get("happiness_score", ""), r.get("still_using", ""), r.get("would_buy_again", ""), r.get("lesson", ""), r.get("reward_given", False)] for r in data.get("post_purchase_reviews", [])])

    set_table(get_or_create_ws(sh, "tasks"), ["创建时间", "任务", "奖励金额", "分类", "状态", "提交时间", "批准人", "入账交易ID"], [[t.get("created_at", ""), t.get("task_name", ""), t.get("reward_amount", 0), t.get("category", ""), t.get("status", ""), t.get("submitted_at", ""), t.get("approved_by", ""), t.get("booked_tx_id", "")] for t in data.get("tasks", [])])

    set_table(get_or_create_ws(sh, "score_history"), ["时间", "信用分", "变化", "事件", "原因"], [[h.get("time", ""), h.get("score", ""), h.get("change", ""), h.get("event", ""), h.get("reason", "")] for h in data.get("score_history", [])])

    set_table(get_or_create_ws(sh, "weekly_report"), ["周起始日", "生成时间", "收入", "消费", "储蓄", "信用分", "周报内容"], [[w.get("week_start", ""), w.get("generated_at", ""), (w.get("metrics") or {}).get("income", ""), (w.get("metrics") or {}).get("expense", ""), (w.get("metrics") or {}).get("savings", ""), (w.get("metrics") or {}).get("score", ""), w.get("content", "")] for w in data.get("weekly_reports", [])])

    set_table(get_or_create_ws(sh, "monthly_cfo_report"), ["月份", "生成时间", "报告内容"], [[r.get("month", ""), r.get("generated_at", ""), r.get("content", "")] for r in data.get("monthly_cfo_reports", [])])

    set_table(get_or_create_ws(sh, "badges"), ["时间", "徽章", "原因", "授予人"], [[b.get("time", ""), b.get("badge", ""), b.get("reason", ""), b.get("granted_by", "")] for b in data.get("badges", [])])

    set_table(get_or_create_ws(sh, "rewards"), ["时间", "奖励", "金额/点数", "状态", "授予人", "说明"], [[r.get("time", ""), r.get("reward", ""), r.get("value", ""), r.get("status", ""), r.get("granted_by", ""), r.get("note", "")] for r in data.get("rewards", [])])

    set_table(get_or_create_ws(sh, "rules"), ["规则", "值"], [[k, v] for k, v in data.get("rules", {}).items()])
    set_table(get_or_create_ws(sh, "settings"), ["设置项", "值"], [[k, v] for k, v in data.get("settings", {}).items()])
    set_table(get_or_create_ws(sh, "accounts"), ["账户", "名称", "目标比例"], [[k, v.get("name", ""), v.get("target_ratio", 0)] for k, v in data.get("accounts", {}).items()])
    set_table(get_or_create_ws(sh, "backups"), ["时间", "操作人", "事件", "原因", "JSON长度"], [[b.get("time", ""), b.get("actor", ""), b.get("event", ""), b.get("reason", ""), len(b.get("snapshot", ""))] for b in data.get("backups", [])[-100:]])


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
        return normalize_data(json.loads(DATA_PATH.read_text(encoding="utf-8")))
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


def add_backup(data: Dict[str, Any], event: str, reason: str) -> Dict[str, Any]:
    d = normalize_data(data)
    snapshot = json.dumps(d, ensure_ascii=False)
    d.setdefault("backups", []).append({
        "time": now_str(),
        "actor": current_actor(),
        "event": event,
        "reason": reason,
        "snapshot": snapshot,
    })
    d["backups"] = d["backups"][-100:]
    return d


def commit(data: Dict[str, Any], event: str = "保存数据", reason: str = "") -> None:
    old_score = calc_financials(st.session_state.get(SESSION_KEY, data)).get("credit_score") if SESSION_KEY in st.session_state else None
    new_data = add_backup(normalize_data(data), event, reason)
    new_score = calc_financials(new_data).get("credit_score")
    if new_score is not None:
        new_data.setdefault("score_history", []).append({
            "time": now_str(),
            "score": new_score,
            "change": None if old_score is None else new_score - old_score,
            "event": event,
            "reason": reason or event,
        })
        new_data["score_history"] = new_data["score_history"][-300:]
    st.session_state[SESSION_KEY] = new_data
    save_data(new_data)


def reload_data() -> None:
    st.session_state.pop(SESSION_KEY, None)


def reset_empty() -> None:
    data = empty_data()
    st.session_state[SESSION_KEY] = data
    save_data(data)


def make_tx(tx_type: str, amount: float, category: str, party: str = "", memo: str = "", tx_date: Optional[str] = None, expected_repayment: float = 0.0, principal_repaid: float = 0.0, interest_received: float = 0.0, due_date: str = "", account: str = "cash") -> Dict[str, Any]:
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
        "account": account,
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


def calc_account_balances(data: Dict[str, Any]) -> Dict[str, float]:
    balances = {k: 0.0 for k in data.get("accounts", {}).keys()}
    balances.setdefault("cash", fnum(data.get("settings", {}).get("start_cash")))
    balances.setdefault("savings", fnum(data.get("settings", {}).get("start_savings")))
    balances["cash"] = fnum(data.get("settings", {}).get("start_cash"))
    balances["savings"] = fnum(data.get("settings", {}).get("start_savings"))

    for tx in data.get("transactions", []):
        account = tx.get("account") or "cash"
        balances.setdefault(account, 0.0)
        typ = tx.get("type")
        amount = fnum(tx.get("amount"))
        if typ == "收入":
            balances[account] += amount
        elif typ == "消费":
            balances[account] -= amount
        elif typ == "转入储蓄":
            balances["cash"] -= amount
            balances["savings"] += amount
        elif typ == "储蓄取出":
            balances["savings"] -= amount
            balances["cash"] += amount
        elif typ == "放贷":
            balances["cash"] -= amount
        elif typ == "还款":
            balances["cash"] += (fnum(tx.get("principal_repaid")) or amount) + fnum(tx.get("interest_received"))
        elif typ == "借入":
            balances["cash"] += amount
        elif typ == "偿还负债":
            balances["cash"] -= amount
    return balances


def calc_credit_score(data: Dict[str, Any], cash: float, savings: float, receivables: float, liabilities: float, total_assets: float, month_income: float, month_expense: float, budget_usage: Dict[str, Dict[str, float]], overdue: float) -> Tuple[Optional[int], List[str], Dict[str, int]]:
    s = data.get("settings", {})
    base = inum(s.get("base_score"), 0)
    if base <= 0:
        return None, ["账户尚未建立信用分。请在设置里建立基础信用分。"], {}

    score = float(base)
    factors = []
    card = {
        "现金纪律": 180,
        "预算纪律": 180,
        "储蓄纪律": 140,
        "贷款纪律": 130,
        "负债纪律": 100,
        "稳定性": 90,
    }

    floor = fnum(s.get("cash_floor"))
    if floor > 0:
        if cash < 0:
            score -= 45
            card["现金纪律"] -= 60
            factors.append("现金余额为负：-45")
        elif cash < floor:
            score -= 16
            card["现金纪律"] -= 25
            factors.append(f"现金低于安全线 {money(floor)}：-16")
        elif cash >= floor * 2:
            score += 6
            card["现金纪律"] += 10
            factors.append("现金缓冲较充足：+6")

    savings_ratio = savings / total_assets if total_assets > 0 else 0
    if total_assets > 0:
        if savings_ratio >= 0.30:
            score += 8
            card["储蓄纪律"] += 15
            factors.append("储蓄占总资产 30% 以上：+8")
        elif savings_ratio < 0.10:
            score -= 8
            card["储蓄纪律"] -= 25
            factors.append("储蓄占总资产低于 10%：-8")

    if month_income <= 0 and month_expense > 0:
        score -= 12
        card["稳定性"] -= 20
        factors.append("本月有消费但没有收入记录：-12")
    elif month_income > 0:
        ratio = month_expense / month_income
        if ratio > 1.0:
            score -= 35
            card["预算纪律"] -= 60
            factors.append("本月支出超过收入：-35")
        elif ratio > 0.85:
            score -= 18
            card["预算纪律"] -= 35
            factors.append("本月支出达到收入 85% 以上：-18")
        elif ratio > 0.65:
            score -= 8
            card["预算纪律"] -= 15
            factors.append("本月支出达到收入 65% 以上：-8")
        elif ratio <= 0.45:
            score += 5
            card["预算纪律"] += 10
            factors.append("本月支出控制在收入 45% 以下：+5")

    loan_ratio = receivables / total_assets if total_assets > 0 else 0
    if loan_ratio > 0.65:
        score -= 30
        card["贷款纪律"] -= 60
        factors.append("应收贷款本金占总资产超过 65%：-30")
    elif loan_ratio > 0.50:
        score -= 20
        card["贷款纪律"] -= 40
        factors.append("应收贷款本金占总资产超过 50%：-20")
    elif loan_ratio > 0.35:
        score -= 9
        card["贷款纪律"] -= 15
        factors.append("应收贷款本金占总资产超过 35%：-9")

    debt_ratio = liabilities / total_assets if total_assets > 0 else 0
    if debt_ratio > 0.50:
        score -= 45
        card["负债纪律"] -= 70
        factors.append("负债占总资产超过 50%：-45")
    elif debt_ratio > 0.25:
        score -= 25
        card["负债纪律"] -= 40
        factors.append("负债占总资产超过 25%：-25")
    elif liabilities > 0:
        score -= 7
        card["负债纪律"] -= 10
        factors.append("存在未偿还负债：-7")

    overspent = sum(1 for r in budget_usage.values() if r["limit"] > 0 and r["ratio"] > 1)
    warning = sum(1 for r in budget_usage.values() if r["limit"] > 0 and 0.85 < r["ratio"] <= 1)
    if overspent:
        p = min(32, overspent * 10)
        score -= p
        card["预算纪律"] -= min(50, overspent * 20)
        factors.append(f"{overspent} 个预算品类已超支：-{p}")
    if warning:
        p = min(15, warning * 4)
        score -= p
        card["预算纪律"] -= min(20, warning * 8)
        factors.append(f"{warning} 个预算品类接近上限：-{p}")
    if overdue > 0:
        score -= 24
        card["贷款纪律"] -= 40
        factors.append(f"存在逾期未收回应收本金 {money(overdue)}：-24")

    for k in card:
        card[k] = int(clamp(card[k], 0, 200))
    if not factors:
        factors.append("当前没有明显加分或扣分因素。")
    return int(round(clamp(score, 300, 850))), factors, card


def calc_financials(data: Dict[str, Any], month: Optional[str] = None) -> Dict[str, Any]:
    month = month or month_str()
    s = data.get("settings", {})
    ccy = s.get("currency", "USD")
    balances = calc_account_balances(data)
    cash = balances.get("cash", 0.0)
    savings = balances.get("savings", 0.0)
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
            total_income += amount
            if is_month:
                month_income += amount
        elif typ == "消费":
            total_expense += amount
            if is_month:
                month_expense += amount
        elif typ == "还款":
            interest = fnum(tx.get("interest_received"))
            total_income += interest
            if is_month:
                month_income += interest
        elif typ == "借入":
            liabilities += amount
        elif typ == "偿还负债":
            liabilities -= amount

    liabilities = max(0.0, liabilities)
    loans = build_loan_book(data.get("transactions", []))
    receivables = sum(fnum(l.get("remaining")) for l in loans)

    overdue = 0.0
    today = date.today()
    for loan in loans:
        due = loan.get("due_date") or ""
        rem = fnum(loan.get("remaining"))
        if rem > 0 and due:
            try:
                if date.fromisoformat(due) < today:
                    overdue += rem
            except Exception:
                pass

    total_assets = cash + savings + receivables + sum(v for k, v in balances.items() if k not in {"cash", "savings"})
    net_assets = total_assets - liabilities
    usage = calc_budget_usage(data, month)
    goals = calc_goals(data)
    score, factors, score_card = calc_credit_score(data, cash, savings, receivables, liabilities, total_assets, month_income, month_expense, usage, overdue)

    spend_income_ratio = month_expense / month_income if month_income > 0 else (1.0 if month_expense > 0 else 0.0)
    loan_asset_ratio = receivables / total_assets if total_assets > 0 else 0.0
    debt_asset_ratio = liabilities / total_assets if total_assets > 0 else 0.0
    savings_asset_ratio = savings / total_assets if total_assets > 0 else 0.0

    return {
        "currency": ccy,
        "cash": cash,
        "savings": savings,
        "balances": balances,
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
        "score_card": score_card,
    }


# ============================================================
# 5. AI 分类、审批、报告
# ============================================================

def local_category(text: str) -> Dict[str, str]:
    s = (text or "").lower()
    rules = {
        "游戏": ["minecraft", "robux", "游戏", "道具", "皮肤", "steam", "switch"],
        "甜品": ["甜品", "奶茶", "冰淇淋", "蛋糕", "糖", "饮料"],
        "玩具": ["玩具", "lego", "乐高", "手办", "娃娃"],
        "学习": ["书", "学习", "课程", "文具", "作业", "训练"],
        "宠物": ["猫", "猫粮", "猫砂", "宠物", "罐头"],
        "餐饮": ["饭", "餐", "披萨", "pizza", "汉堡", "麦当劳"],
    }
    category = "其他"
    for cat, keys in rules.items():
        if any(k in s for k in keys):
            category = cat
            break
    necessity = "半必要" if category in {"学习", "宠物", "餐饮"} else "非必要"
    impulse = "高" if category in {"游戏", "甜品", "玩具"} else "中"
    return {"category": category, "necessity": necessity, "impulse_level": impulse, "merchant": "自然语言输入", "note": "本地关键词分类"}


def ai_classify_purchase(text: str) -> Dict[str, str]:
    api_key = secret_value("DEEPSEEK_API_KEY")
    model = secret_value("DEEPSEEK_MODEL", "deepseek-v4-flash")
    if not api_key or requests is None:
        return local_category(text)
    prompt = f"""
请把下面的儿童消费/任务/贷款请求分类，只输出 JSON。
分类只能从以下选择：
游戏、甜品、玩具、学习、宠物、餐饮、家庭贷款、任务收入、其他

字段：
category
necessity: 必要 / 半必要 / 非必要
impulse_level: 低 / 中 / 高
merchant
note

用户输入：
{text}
""".strip()
    try:
        resp = requests.post(
            "https://api.deepseek.com/chat/completions",
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            json={"model": model, "messages": [{"role": "user", "content": prompt}], "temperature": 0, "max_tokens": 220},
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


def extract_amounts(text: str) -> List[float]:
    nums: List[float] = []
    for p in [r"\$\s*([0-9]+(?:\.[0-9]+)?)", r"([0-9]+(?:\.[0-9]+)?)\s*(?:美元|刀|块|元)"]:
        for m in re.finditer(p, text or "", flags=re.I):
            nums.append(float(m.group(1)))
    if not nums:
        for m in re.finditer(r"([0-9]+(?:\.[0-9]+)?)", text or ""):
            nums.append(float(m.group(1)))
    return nums


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
    rules = data.get("rules", {})
    floor = fnum(data.get("settings", {}).get("cash_floor"))
    threshold = fnum(rules.get("approval_threshold"), fnum(data.get("settings", {}).get("approval_threshold"), 15.0))
    warning_line = fnum(rules.get("budget_warning_line"), 0.90)
    reject_line = fnum(rules.get("budget_reject_line"), 1.20)
    credit_reject = fnum(rules.get("credit_reject_line"), 650)

    result = "批准"
    reasons: List[str] = []

    if amount <= 0:
        result = "拒绝"
        reasons.append("购买金额必须大于 0。")
    if after["cash"] < 0:
        result = "拒绝"
        reasons.append("购买后现金余额为负，触发硬性拒绝。")
    if budget["limit"] > 0 and budget["ratio"] > reject_line:
        result = "拒绝"
        reasons.append(f"{category}预算使用率将达到 {percent(budget['ratio'])}，超过拒绝线 {percent(reject_line)}。")
    if after["credit_score"] is not None and after["credit_score"] < credit_reject:
        result = "拒绝"
        reasons.append(f"交易后信用分低于 {credit_reject:.0f}。")

    if result != "拒绝":
        flags = []
        if amount >= threshold:
            flags.append(f"金额达到家长审批线 {money(threshold)}。")
        if floor > 0 and after["cash"] < floor:
            flags.append(f"购买后现金余额低于安全线 {money(floor)}。")
        if budget["limit"] > 0 and budget["ratio"] >= warning_line:
            flags.append(f"{category}预算使用率将达到 {percent(budget['ratio'])}，接近或超过预警线。")
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
        action = "建议延迟到下一笔收入到账后再买，或者加入愿望清单冷静期。"
    elif result == "待家长审批":
        action = "进入待审批队列，由家长最终确认。"

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


def scenario_planning(data: Dict[str, Any], amount: float, category: str, desc: str, merchant: str = "") -> List[Dict[str, Any]]:
    today_decision = evaluate_purchase_decision(data, amount, category, desc, merchant)
    next_data = copy.deepcopy(data)
    next_data.setdefault("transactions", []).append(make_tx("收入", amount, "零花钱", memo="情景模拟：下周收入到账"))
    next_decision = evaluate_purchase_decision(next_data, amount, category, desc + "（下周购买）", merchant)
    half = round(amount / 2, 2)
    split_decision = evaluate_purchase_decision(data, half, category, desc + "（分两周第一笔）", merchant)
    return [
        {"方案": "今天购买", "结论": today_decision["result"], "交易后现金": today_decision["metrics"]["交易后现金余额"], "交易后信用分": today_decision["metrics"]["交易后信用分"], "信用分变化": today_decision["metrics"]["信用分变化"], "说明": today_decision["action"]},
        {"方案": "下周收入后购买", "结论": next_decision["result"], "交易后现金": next_decision["metrics"]["交易后现金余额"], "交易后信用分": next_decision["metrics"]["交易后信用分"], "信用分变化": next_decision["metrics"]["信用分变化"], "说明": "先等同等金额收入到账，再购买，现金压力明显下降。"},
        {"方案": "分两周购买", "结论": split_decision["result"], "交易后现金": split_decision["metrics"]["交易后现金余额"], "交易后信用分": split_decision["metrics"]["交易后信用分"], "信用分变化": split_decision["metrics"]["信用分变化"], "说明": f"先支出 {money(half)}，剩余下周处理，预算冲击最低。"},
    ]


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
    payload = {"kind": decision.get("kind"), "result": decision.get("result"), "metrics": decision.get("metrics"), "reasons": decision.get("reasons"), "action": decision.get("action")}
    prompt = "你是阿苏私人银行3.0的中文私人银行客户经理。只能解释Python已计算的硬指标。禁止编造数字，禁止改变审批结论。"
    try:
        resp = requests.post(
            "https://api.deepseek.com/chat/completions",
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            json={"model": model, "messages": [{"role": "system", "content": prompt}, {"role": "user", "content": json.dumps(payload, ensure_ascii=False, indent=2)}], "temperature": 0.2, "max_tokens": 900},
            timeout=20,
        )
        resp.raise_for_status()
        return resp.json()["choices"][0]["message"]["content"]
    except Exception as exc:
        return local_report(decision) + f"\n\n> DeepSeek 调用失败，已切换本地报告：{exc}"


# ============================================================
# 6. 运营逻辑：任务、愿望、复盘、报告
# ============================================================

def cooldown_hours_for(data: Dict[str, Any], category: str) -> int:
    s = data.get("settings", {})
    if category == "游戏":
        return inum(s.get("cooldown_game_hours"), 72)
    if category == "甜品":
        return inum(s.get("cooldown_treat_hours"), 24)
    if category == "玩具":
        return inum(s.get("cooldown_toy_hours"), 168)
    return inum(s.get("cooldown_default_hours"), 48)


def add_wishlist_item(data: Dict[str, Any], item: str, amount: float, category: str, merchant: str) -> None:
    until = datetime.now() + timedelta(hours=cooldown_hours_for(data, category))
    data.setdefault("wishlist", []).append({
        "id": uid(),
        "created_at": now_str(),
        "item": item,
        "amount": amount,
        "category": category,
        "merchant": merchant,
        "cooldown_until": until.isoformat(timespec="seconds"),
        "status": "冷静期",
        "applicant": current_actor(),
        "final_decision": "",
    })


def update_wishlist_status(data: Dict[str, Any]) -> None:
    now = datetime.now()
    for w in data.get("wishlist", []):
        if w.get("status") == "冷静期":
            try:
                if datetime.fromisoformat(w.get("cooldown_until")) <= now:
                    w["status"] = "可申请"
            except Exception:
                pass


def grant_badge(data: Dict[str, Any], badge: str, reason: str) -> None:
    data.setdefault("badges", []).append({"time": now_str(), "badge": badge, "reason": reason, "granted_by": current_actor()})


def grant_reward(data: Dict[str, Any], reward: str, value: Any, note: str) -> None:
    data.setdefault("rewards", []).append({"time": now_str(), "reward": reward, "value": value, "status": "未使用", "granted_by": current_actor(), "note": note})


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
    spend_by_cat: Dict[str, float] = {}
    for tx in rows:
        if tx.get("type") == "消费":
            cat = tx.get("category") or "其他"
            spend_by_cat[cat] = spend_by_cat.get(cat, 0) + fnum(tx.get("amount"))
    top_cat = max(spend_by_cat, key=spend_by_cat.get) if spend_by_cat else "无"
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
    return {"week_start": start.isoformat(), "generated_at": now_str(), "content": content, "metrics": {"income": income, "expense": expense, "savings": savings_move, "score": score_text(m["credit_score"]), "top_category": top_cat}}


def generate_monthly_cfo_report(data: Dict[str, Any]) -> Dict[str, Any]:
    m = calc_financials(data)
    month = month_str()
    analytics = behavior_analytics(data)
    content = "\n".join([
        f"阿苏家庭 CFO 月报：{month}",
        "",
        "一、现金流",
        f"本月收入：{money(m['month_income'])}",
        f"本月消费：{money(m['month_expense'])}",
        f"支出收入比：{percent(m['spend_income_ratio'])}",
        "",
        "二、资产结构",
        f"总资产：{money(m['total_assets'])}",
        f"现金：{money(m['cash'])}",
        f"储蓄：{money(m['savings'])}",
        f"应收贷款本金：{money(m['receivables'])}",
        "",
        "三、行为变化",
        f"冲动消费申请数：{analytics['impulse_requests']}",
        f"审批通过率：{percent(analytics['approval_rate'])}",
        f"复盘平均满意度：{analytics['avg_review_score']:.1f}/10",
        "",
        "四、父母干预建议",
        *[f"{i+1}. {x}" for i, x in enumerate(parent_intervention_suggestions(data))],
    ])
    return {"month": month, "generated_at": now_str(), "content": content}


def build_weekly_plan(data: Dict[str, Any]) -> List[str]:
    if not has_activity(data):
        return ["先建立第一笔收入或现金余额。", "新增至少 3 个预算分类。", "建立基础信用分。"]
    m = calc_financials(data)
    actions: List[str] = []
    overspent = [cat for cat, r in m["budget_usage"].items() if r["limit"] > 0 and r["ratio"] > 1]
    near = [cat for cat, r in m["budget_usage"].items() if r["limit"] > 0 and 0.85 < r["ratio"] <= 1]
    if overspent:
        actions.append(f"本周暂停 {', '.join(overspent)} 类非必要消费。")
    elif near:
        actions.append(f"本周 {', '.join(near)} 类消费必须先审批。")
    if m["loan_asset_ratio"] > 0.45:
        actions.append("优先收回应收贷款本金。")
    if m["savings_asset_ratio"] < 0.25 and m["month_income"] > 0:
        actions.append("下一笔收入的 30% 转入储蓄。")
    if len([w for w in data.get("wishlist", []) if w.get("status") == "冷静期"]) > 0:
        actions.append("保持愿望清单冷静期，不急于审批。")
    return actions[:5] or ["维持当前节奏，非必要消费继续走审批。"]


def behavior_analytics(data: Dict[str, Any]) -> Dict[str, Any]:
    pending = data.get("pending_requests", [])
    approved = [r for r in pending if r.get("parent_status") == "已批准"]
    rejected = [r for r in pending if r.get("parent_status") == "已拒绝"]
    impulse = [r for r in pending if r.get("impulse_level") == "高"]
    reviews = data.get("post_purchase_reviews", [])
    avg = sum(inum(r.get("happiness_score"), 0) for r in reviews) / len(reviews) if reviews else 0.0
    total_decided = len(approved) + len(rejected)
    return {
        "total_requests": len(pending),
        "approved": len(approved),
        "rejected": len(rejected),
        "approval_rate": len(approved) / total_decided if total_decided else 0.0,
        "impulse_requests": len(impulse),
        "avg_review_score": avg,
        "wishlist_active": len([w for w in data.get("wishlist", []) if w.get("status") in {"冷静期", "可申请"}]),
    }


def parent_intervention_suggestions(data: Dict[str, Any]) -> List[str]:
    a = behavior_analytics(data)
    m = calc_financials(data)
    suggestions = []
    if a["impulse_requests"] >= 3:
        suggestions.append("高冲动申请偏多，建议把游戏和甜品冷静期延长。")
    if a["avg_review_score"] and a["avg_review_score"] < 6:
        suggestions.append("消费后满意度偏低，建议强化复盘后再给同类预算。")
    if m["spend_income_ratio"] > 0.85 and m["month_income"] > 0:
        suggestions.append("本月支出收入比偏高，建议下月降低自由消费预算。")
    if m["savings_asset_ratio"] < 0.2 and m["total_assets"] > 0:
        suggestions.append("储蓄占比不足，建议收入到账后自动划拨储蓄。")
    if not suggestions:
        suggestions.append("当前行为稳定，建议以奖励和复盘为主，减少强干预。")
    return suggestions


def financial_word_card() -> Dict[str, str]:
    cards = [
        {"英文": "Budget", "中文": "预算", "解释": "提前规定一类消费最多能花多少钱。"},
        {"英文": "Liquidity", "中文": "流动性", "解释": "你马上能用的钱。借出去的钱还是资产，但流动性差。"},
        {"英文": "Credit Score", "中文": "信用分", "解释": "系统根据你的消费、储蓄、还款纪律给出的行为评分。"},
        {"英文": "Opportunity Cost", "中文": "机会成本", "解释": "买了这个东西，就少了买别的东西的机会。"},
        {"英文": "Interest", "中文": "利息", "解释": "别人使用你的钱，需要额外支付的费用。"},
        {"英文": "Principal", "中文": "本金", "解释": "最开始借出去或存进去的钱。"},
        {"英文": "Debt", "中文": "债务", "解释": "你欠别人的钱。"},
        {"英文": "Asset", "中文": "资产", "解释": "属于你的、有价值的东西。"},
    ]
    idx = date.today().toordinal() % len(cards)
    return cards[idx]


def evaluate_merchant_access(data: Dict[str, Any], merchant: Dict[str, Any]) -> Dict[str, Any]:
    m = calc_financials(data)
    cat = merchant.get("category") or "其他"
    usage = m.get("budget_usage", {}).get(cat, {"limit": 0, "ratio": 0})
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
    return {"result": result, "message": message, "failed": failed, "checks": checks, "metrics": {}}


# ============================================================
# 7. 登录
# ============================================================

def passwords_enabled() -> bool:
    return bool(secret_value("DAD_PASSWORD") and secret_value("MOM_PASSWORD") and secret_value("ASU_PASSWORD"))


def login_gate() -> None:
    if not passwords_enabled():
        if "role" not in st.session_state:
            st.session_state["role"] = "爸爸"
        st.sidebar.warning("未启用密码登录，使用手动身份选择。")
        st.session_state["role"] = st.sidebar.selectbox("当前操作人", ["爸爸", "妈妈", "阿苏"], index=["爸爸", "妈妈", "阿苏"].index(st.session_state["role"]))
        return

    if st.session_state.get("authenticated"):
        st.sidebar.success(f"已登录：{st.session_state.get('role')}")
        if st.sidebar.button("退出登录"):
            st.session_state.clear()
            st.rerun()
        return

    st.title("阿苏私人银行")
    role = st.selectbox("选择身份", ["爸爸", "妈妈", "阿苏"])
    pw = st.text_input("密码", type="password")
    if st.button("登录", type="primary"):
        key = {"爸爸": "DAD_PASSWORD", "妈妈": "MOM_PASSWORD", "阿苏": "ASU_PASSWORD"}[role]
        if pw == secret_value(key):
            st.session_state["authenticated"] = True
            st.session_state["role"] = role
            st.rerun()
        else:
            st.error("密码错误")
    st.stop()


# ============================================================
# 8. UI 样式与组件
# ============================================================

def inject_css() -> None:
    st.markdown("""
    <style>
    .main .block-container {max-width:1280px; padding-top:1rem; padding-bottom:2rem;}
    .hero {
        background: radial-gradient(circle at 20% 20%, #38bdf8 0%, transparent 25%),
                    linear-gradient(135deg, #002B5B 0%, #005EB8 50%, #00A3E0 100%);
        color:white; padding:30px 34px; border-radius:28px; margin-bottom:20px;
        box-shadow:0 18px 42px rgba(0,62,130,.26);
    }
    .hero h1 {margin:0; font-size:40px; letter-spacing:.3px;}
    .hero p {margin:10px 0 0; opacity:.95; font-size:17px;}
    .hero-grid {display:grid; grid-template-columns:repeat(4,1fr); gap:14px; margin-top:18px;}
    .hero-mini {background:rgba(255,255,255,.14); border:1px solid rgba(255,255,255,.25); border-radius:18px; padding:12px;}
    .hero-mini div:first-child {font-size:12px; opacity:.85;}
    .hero-mini div:last-child {font-size:22px; font-weight:800;}
    .card {border:1px solid #D6E4F5; border-radius:20px; background:white; padding:18px; box-shadow:0 8px 22px rgba(15,23,42,.06); min-height:112px;}
    .soft-card {border:1px solid #E5E7EB; border-radius:22px; background:#FFFFFF; padding:20px; box-shadow:0 10px 28px rgba(15,23,42,.06);}
    .label {color:#6B7280; font-size:13px; margin-bottom:6px;}
    .value {font-size:28px; font-weight:800; color:#111827; line-height:1.12;}
    .sub {color:#6B7280; font-size:12px; margin-top:6px;}
    .decision {border:1px solid #D6E4F5; border-radius:18px; background:white; padding:18px; margin:10px 0; box-shadow:0 8px 22px rgba(15,23,42,.05);}
    .ok {border-left:7px solid #166534;}
    .warn {border-left:7px solid #B45309;}
    .bad {border-left:7px solid #B91C1C;}
    .pill {display:inline-block; padding:4px 10px; border-radius:999px; background:#EAF3FF; color:#003C71; border:1px solid #B9D7F6; font-size:12px; font-weight:700; margin-bottom:8px;}
    .small {color:#6B7280; font-size:13px;}
    div[data-testid="stMetric"] {border:1px solid #D6E4F5; border-radius:16px; padding:12px 14px; background:white; box-shadow:0 6px 18px rgba(15,23,42,.05);}
    </style>
    """, unsafe_allow_html=True)


def render_hero(data: Dict[str, Any]) -> None:
    m = calc_financials(data)
    st.markdown(f"""
    <div class="hero">
      <h1>{APP_NAME}</h1>
      <p>家庭金融操作系统：愿望清单 · 冷静期 · 任务收入 · 消费复盘 · AI CFO 月报 · 儿童金融课程</p>
      <div class="hero-grid">
        <div class="hero-mini"><div>总资产</div><div>{money(m['total_assets'])}</div></div>
        <div class="hero-mini"><div>现金</div><div>{money(m['cash'])}</div></div>
        <div class="hero-mini"><div>信用分</div><div>{score_text(m['credit_score'])}</div></div>
        <div class="hero-mini"><div>待审批</div><div>{len([r for r in data.get('pending_requests', []) if r.get('parent_status') == '待审批'])}</div></div>
      </div>
    </div>
    """, unsafe_allow_html=True)
    st.caption(f"当前身份：{current_actor()}；当前存储：{st.session_state.get('storage_backend', '未知')}")


def metric_card(label: str, value: str, sub: str = "") -> None:
    st.markdown(f"<div class='card'><div class='label'>{label}</div><div class='value'>{value}</div><div class='sub'>{sub}</div></div>", unsafe_allow_html=True)


def render_decision(decision: Dict[str, Any], data: Dict[str, Any], key: str, allow_queue: bool = True) -> None:
    css = decision_class(decision.get("result", "观察"))
    st.markdown(f"<div class='decision {css}'><div class='pill'>{decision.get('kind','审批')}</div><h3>审批结果：{decision.get('result')}</h3><p class='small'>{decision.get('action','')}</p></div>", unsafe_allow_html=True)
    if decision.get("reasons"):
        st.markdown("#### 风控原因")
        for r in decision["reasons"]:
            st.write(f"- {r}")
    if decision.get("metrics"):
        st.dataframe(pd.DataFrame([{"指标": k, "值": pretty_value(k, v)} for k, v in decision["metrics"].items()]), use_container_width=True, hide_index=True)
    with st.expander("AI 客户经理报告", expanded=True):
        st.markdown(local_report(decision))

    tx = decision.get("pending_tx")
    if tx:
        c1, c2 = st.columns(2)
        with c1:
            if st.button("直接入账", disabled=(not is_parent() or decision.get("result") == "拒绝"), type="primary", key=f"{key}_book"):
                data.setdefault("transactions", []).append(tx)
                commit(data, event="直接入账", reason=tx.get("memo", ""))
                st.success("已入账。")
                st.rerun()
        with c2:
            if allow_queue and st.button("提交待审批", disabled=decision.get("result") == "拒绝", key=f"{key}_queue"):
                req = {
                    "id": uid(), "created_at": now_str(), "request_text": tx.get("memo", ""), "request_type": tx.get("type", "消费"),
                    "amount": tx.get("amount", 0), "category": tx.get("category", "其他"), "merchant": tx.get("party", ""),
                    "applicant": current_actor(), "ai_category": tx.get("category", "其他"), "necessity": "见AI分类",
                    "impulse_level": "见AI分类", "python_decision": decision.get("result", ""), "parent_status": "待审批",
                    "final_status": "未入账", "decision_json": decision, "pending_tx": tx, "parent_note": "", "approver": "",
                    "approved_at": "", "booked_at": "",
                }
                data.setdefault("pending_requests", []).append(req)
                commit(data, event="提交待审批", reason=tx.get("memo", ""))
                st.success("已提交待审批。")
                st.rerun()


# ============================================================
# 9. 页面
# ============================================================

def page_home(data: Dict[str, Any]) -> None:
    m = calc_financials(data)
    render_hero(data)
    c1, c2, c3, c4 = st.columns(4)
    with c1: metric_card("净资产", money(m["net_assets"]), "总资产 - 负债")
    with c2: metric_card("储蓄余额", money(m["savings"]), f"储蓄占比 {percent(m['savings_asset_ratio'])}")
    with c3: metric_card("本月消费", money(m["month_expense"]), f"收入 {money(m['month_income'])}")
    with c4: metric_card("愿望清单", str(len([w for w in data.get("wishlist", []) if w.get("status") in {"冷静期", "可申请"}])), "训练延迟满足")

    st.divider()
    left, mid, right = st.columns([1,1,1])
    with left:
        st.subheader("今日金融词卡")
        card = financial_word_card()
        st.markdown(f"<div class='soft-card'><h3>{card['英文']}（{card['中文']}）</h3><p>{card['解释']}</p></div>", unsafe_allow_html=True)
    with mid:
        st.subheader("本周行动")
        for i, a in enumerate(build_weekly_plan(data), 1):
            st.write(f"**{i}.** {a}")
    with right:
        st.subheader("父母干预建议")
        for i, a in enumerate(parent_intervention_suggestions(data), 1):
            st.write(f"**{i}.** {a}")


def page_child(data: Dict[str, Any]) -> None:
    st.subheader("儿童模式")
    tab1, tab2, tab3, tab4 = st.tabs(["我要买东西", "我的愿望清单", "任务换收入", "我的奖励"])

    with tab1:
        text = st.text_area("你想买什么？", value="我想买 $18 的 Minecraft 道具")
        if st.button("提交给 AI 审批", type="primary"):
            amounts = extract_amounts(text)
            cls = ai_classify_purchase(text)
            amount = amounts[0] if amounts else 0.0
            decision = evaluate_purchase_decision(data, amount, cls["category"], text, cls["merchant"])
            st.session_state["child_decision"] = decision
            st.session_state["child_cls"] = cls
            st.session_state["child_scenarios"] = scenario_planning(data, amount, cls["category"], text, cls["merchant"])
        if "child_cls" in st.session_state:
            st.markdown("#### AI 分类")
            st.dataframe(pd.DataFrame([st.session_state["child_cls"]]), use_container_width=True, hide_index=True)
        if "child_scenarios" in st.session_state:
            st.markdown("#### 三种方案")
            st.dataframe(pd.DataFrame(st.session_state["child_scenarios"]), use_container_width=True, hide_index=True)
        if "child_decision" in st.session_state:
            render_decision(st.session_state["child_decision"], data, "child_decision")

    with tab2:
        update_wishlist_status(data)
        with st.form("wishlist_form"):
            item = st.text_input("愿望", "")
            amount = st.number_input("估计金额", min_value=0.0, step=1.0, format="%.2f")
            category = st.selectbox("分类", spending_categories(data))
            merchant = st.text_input("商户", "")
            submitted = st.form_submit_button("加入愿望清单", type="primary")
        if submitted and item:
            add_wishlist_item(data, item, amount, category, merchant)
            commit(data, event="加入愿望清单", reason=item)
            st.success("已加入愿望清单。冷静期结束后再决定。")
            st.rerun()
        st.dataframe(pd.DataFrame(data.get("wishlist", [])), use_container_width=True, hide_index=True)

    with tab3:
        st.markdown("#### 可做任务")
        tasks = data.get("tasks", [])
        if not tasks:
            st.info("暂无任务。请让爸爸妈妈在家长驾驶舱发布任务。")
        for t in tasks:
            st.markdown(f"**{t.get('task_name')}** · 奖励 {money(t.get('reward_amount'))} · 状态：{t.get('status')}")
            if t.get("status") == "待完成" and st.button("我完成了，提交确认", key=f"submit_task_{t['id']}"):
                t["status"] = "待家长确认"
                t["submitted_at"] = now_str()
                commit(data, event="提交任务", reason=t.get("task_name", ""))
                st.success("已提交给家长确认。")
                st.rerun()

    with tab4:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### 徽章")
            st.dataframe(pd.DataFrame(data.get("badges", [])), use_container_width=True, hide_index=True)
        with c2:
            st.markdown("#### 奖励券")
            st.dataframe(pd.DataFrame(data.get("rewards", [])), use_container_width=True, hide_index=True)


def page_parent(data: Dict[str, Any]) -> None:
    st.subheader("家长驾驶舱")
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["审批中心", "发布任务", "奖励与徽章", "行为分析", "备份恢复"])

    with tab1:
        pending = data.get("pending_requests", [])
        if not pending:
            st.info("暂无待审批申请。")
        for r in sorted(pending, key=lambda x: str(x.get("created_at", "")), reverse=True):
            css = decision_class(r.get("parent_status", "待审批"))
            st.markdown(f"<div class='decision {css}'><div class='pill'>{r.get('parent_status')}</div><h3>{r.get('request_type')}：{money(r.get('amount'))} · {r.get('category')}</h3><p class='small'>{r.get('request_text')}</p><p class='small'>系统结论：{r.get('python_decision')}；申请人：{r.get('applicant')}</p></div>", unsafe_allow_html=True)
            if r.get("parent_status") == "待审批":
                c1, c2, c3 = st.columns([1,1,2])
                with c1:
                    if st.button("批准并入账", key=f"approve_{r['id']}", type="primary", disabled=not is_parent()):
                        tx = r.get("pending_tx")
                        if tx:
                            data.setdefault("transactions", []).append(tx)
                        r["parent_status"] = "已批准"
                        r["final_status"] = "已入账"
                        r["approver"] = current_actor()
                        r["approved_at"] = now_str()
                        r["booked_at"] = now_str()
                        commit(data, event="家长批准入账", reason=r.get("request_text", ""))
                        st.success("已批准并入账。")
                        st.rerun()
                with c2:
                    if st.button("拒绝", key=f"reject_{r['id']}", disabled=not is_parent()):
                        r["parent_status"] = "已拒绝"
                        r["final_status"] = "未入账"
                        r["approver"] = current_actor()
                        r["approved_at"] = now_str()
                        commit(data, event="家长拒绝申请", reason=r.get("request_text", ""))
                        st.warning("已拒绝。")
                        st.rerun()
                with c3:
                    note = st.text_input("备注", value=r.get("parent_note", ""), key=f"note_{r['id']}")
                    if st.button("保存备注", key=f"save_note_{r['id']}", disabled=not is_parent()):
                        r["parent_note"] = note
                        commit(data, event="保存审批备注", reason=note)
                        st.success("已保存。")
                        st.rerun()

    with tab2:
        with st.form("task_form"):
            task_name = st.text_input("任务名称", "阅读 30 分钟")
            reward_amount = st.number_input("奖励金额", min_value=0.0, value=fnum(data.get("rules", {}).get("task_default_reward"), 1.0), step=0.5)
            category = st.text_input("收入分类", "任务收入")
            submitted = st.form_submit_button("发布任务", type="primary", disabled=not is_parent())
        if submitted:
            data.setdefault("tasks", []).append({"id": uid(), "created_at": now_str(), "task_name": task_name, "reward_amount": reward_amount, "category": category, "status": "待完成", "submitted_at": "", "approved_by": "", "booked_tx_id": ""})
            commit(data, event="发布任务", reason=task_name)
            st.success("任务已发布。")
            st.rerun()
        for t in data.get("tasks", []):
            st.write(f"{t.get('task_name')} · {money(t.get('reward_amount'))} · {t.get('status')}")
            if t.get("status") == "待家长确认" and st.button("确认并发放收入", key=f"confirm_task_{t['id']}", disabled=not is_parent()):
                tx = make_tx("收入", fnum(t.get("reward_amount")), t.get("category", "任务收入"), party="家庭任务市场", memo=t.get("task_name", ""))
                data.setdefault("transactions", []).append(tx)
                t["status"] = "已发放"
                t["approved_by"] = current_actor()
                t["booked_tx_id"] = tx["id"]
                grant_badge(data, "任务执行者", f"完成任务：{t.get('task_name')}")
                commit(data, event="确认任务收入", reason=t.get("task_name", ""))
                st.success("已发放收入。")
                st.rerun()

    with tab3:
        c1, c2 = st.columns(2)
        with c1:
            with st.form("badge_form"):
                badge = st.text_input("徽章名称", "理性消费者")
                reason = st.text_input("授予原因", "完成冷静期后放弃冲动消费")
                ok = st.form_submit_button("授予徽章", disabled=not is_parent())
            if ok:
                grant_badge(data, badge, reason)
                commit(data, event="授予徽章", reason=badge)
                st.success("已授予徽章。")
                st.rerun()
        with c2:
            with st.form("reward_form"):
                reward = st.text_input("奖励名称", "甜品券")
                value = st.text_input("价值", "$3")
                note = st.text_input("说明", "本周表现良好")
                ok = st.form_submit_button("发放奖励", disabled=not is_parent())
            if ok:
                grant_reward(data, reward, value, note)
                commit(data, event="发放奖励", reason=reward)
                st.success("已发放奖励。")
                st.rerun()

    with tab4:
        a = behavior_analytics(data)
        st.dataframe(pd.DataFrame([a]), use_container_width=True, hide_index=True)
        st.markdown("#### 父母干预建议")
        for i, s in enumerate(parent_intervention_suggestions(data), 1):
            st.write(f"**{i}.** {s}")

    with tab5:
        if not is_admin():
            st.info("只有爸爸管理员可以恢复备份。")
        backups = data.get("backups", [])
        if not backups:
            st.info("暂无备份。")
        else:
            labels = [f"{i}: {b.get('time')} · {b.get('event')} · {b.get('reason')}" for i, b in enumerate(backups)]
            idx = st.selectbox("选择备份", list(range(len(labels))), format_func=lambda i: labels[i])
            if st.button("恢复到此备份", disabled=not is_admin()):
                try:
                    restored = normalize_data(json.loads(backups[idx]["snapshot"]))
                    commit(restored, event="恢复备份", reason=labels[idx])
                    st.success("已恢复。")
                    st.rerun()
                except Exception as e:
                    st.error(f"恢复失败：{e}")


def page_approval_and_transactions(data: Dict[str, Any]) -> None:
    st.subheader("审批与交易")
    tab1, tab2, tab3, tab4 = st.tabs(["AI 决策中心", "新增交易", "贷款审批", "交易流水"])
    with tab1:
        text = st.text_area("输入请求", value="我想买 $18 的 Minecraft 道具")
        if st.button("AI 分类 + 风控审批", type="primary"):
            amounts = extract_amounts(text)
            cls = ai_classify_purchase(text)
            amount = amounts[0] if amounts else 0.0
            decision = evaluate_purchase_decision(data, amount, cls["category"], text, cls["merchant"])
            st.session_state["ai_cls"] = cls
            st.session_state["ai_decision"] = decision
            st.session_state["ai_scenarios"] = scenario_planning(data, amount, cls["category"], text, cls["merchant"])
        if "ai_cls" in st.session_state:
            st.markdown("#### AI 分类")
            st.dataframe(pd.DataFrame([st.session_state["ai_cls"]]), use_container_width=True, hide_index=True)
        if "ai_scenarios" in st.session_state:
            st.markdown("#### 情景规划")
            st.dataframe(pd.DataFrame(st.session_state["ai_scenarios"]), use_container_width=True, hide_index=True)
        if "ai_decision" in st.session_state:
            render_decision(st.session_state["ai_decision"], data, "ai_decision")
    with tab2:
        tx_type = st.selectbox("交易类型", ["收入", "消费", "转入储蓄", "储蓄取出", "放贷", "还款", "借入", "偿还负债"], index=0)
        with st.form("add_tx_form"):
            d = st.date_input("日期", value=date.today())
            amount = st.number_input("金额", min_value=0.0, step=1.0, format="%.2f")
            category = st.selectbox("分类", category_options(data))
            party = st.text_input("对象 / 商户 / 借款人", "")
            account = st.selectbox("账户", list(data.get("accounts", {}).keys()))
            memo = st.text_area("备注", "")
            expected = 0.0; principal_repaid = 0.0; interest_received = 0.0; due = ""
            if tx_type == "放贷":
                expected = st.number_input("预计回款总额", min_value=0.0, step=1.0, format="%.2f")
                due = st.date_input("预计还款日", value=date.today() + timedelta(days=7)).isoformat()
            if tx_type == "还款":
                principal_repaid = st.number_input("本金回收", min_value=0.0, step=1.0, format="%.2f")
                interest_received = st.number_input("利息收入", min_value=0.0, step=0.5, format="%.2f")
                amount = principal_repaid + interest_received
            ok = st.form_submit_button("保存交易", type="primary", disabled=(current_actor()=="阿苏"))
        if ok:
            tx = make_tx(tx_type, amount, category, party=party, memo=memo, tx_date=d.isoformat(), expected_repayment=expected, principal_repaid=principal_repaid, interest_received=interest_received, due_date=due, account=account)
            data.setdefault("transactions", []).append(tx)
            commit(data, event=f"新增交易：{tx_type}", reason=memo)
            st.success("交易已保存。")
            st.rerun()
    with tab3:
        st.info("贷款审批可在 AI 决策中心输入：借给爸爸 $20，预计一周后还 $22。")
    with tab4:
        df = transactions_df(data)
        if df.empty:
            st.info("暂无交易。")
        else:
            st.dataframe(df.drop(columns=["id"], errors="ignore"), use_container_width=True, hide_index=True)


def transactions_df(data: Dict[str, Any]) -> pd.DataFrame:
    rows = []
    for tx in data.get("transactions", []):
        rows.append({"日期": tx.get("date"), "类型": tx.get("type"), "金额": fnum(tx.get("amount")), "分类": tx.get("category"), "对象/商户": tx.get("party"), "账户": tx.get("account"), "备注": tx.get("memo"), "id": tx.get("id")})
    return pd.DataFrame(rows).sort_values("日期", ascending=False) if rows else pd.DataFrame()


def page_growth(data: Dict[str, Any]) -> None:
    st.subheader("成长与课程")
    tab1, tab2, tab3, tab4 = st.tabs(["消费后复盘", "金融词卡", "儿童课程", "信用评分卡"])
    with tab1:
        df = transactions_df(data)
        purchase_df = df[df["类型"] == "消费"] if not df.empty else pd.DataFrame()
        if purchase_df.empty:
            st.info("暂无可复盘消费。")
        else:
            tx_id = st.selectbox("选择消费交易", list(purchase_df["id"]), format_func=lambda x: str(purchase_df[purchase_df["id"] == x].iloc[0]["备注"]))
            with st.form("review_form"):
                score = st.slider("买完后的满意度", 1, 10, 7)
                still = st.selectbox("现在还在用吗", ["是", "否", "偶尔"])
                again = st.selectbox("如果重新选择，还会买吗", ["会", "不会", "不确定"])
                lesson = st.text_area("这次学到什么")
                ok = st.form_submit_button("保存复盘")
            if ok:
                data.setdefault("post_purchase_reviews", []).append({"id": uid(), "transaction_id": tx_id, "review_time": now_str(), "happiness_score": score, "still_using": still, "would_buy_again": again, "lesson": lesson, "reward_given": True})
                grant_reward(data, "复盘奖励", data.get("rules", {}).get("review_reward_points", 2), "完成消费后复盘")
                commit(data, event="消费后复盘", reason=lesson)
                st.success("复盘已保存。")
                st.rerun()
    with tab2:
        card = financial_word_card()
        st.markdown(f"<div class='soft-card'><h2>{card['英文']}（{card['中文']}）</h2><p>{card['解释']}</p></div>", unsafe_allow_html=True)
    with tab3:
        curriculum = [
            ("第一课：预算 Budget", "学习如何给游戏、甜品、学习用品设定预算。"),
            ("第二课：流动性 Liquidity", "理解为什么借出去的钱还是资产，但不能马上花。"),
            ("第三课：信用 Credit", "理解按时还款、预算纪律如何影响信用分。"),
            ("第四课：机会成本 Opportunity Cost", "买了一个东西，就失去买另一个东西的机会。"),
            ("第五课：资产配置 Asset Allocation", "把钱分到现金、储蓄、目标基金和慈善账户。"),
        ]
        progress = data.setdefault("curriculum_progress", {})
        for title, desc in curriculum:
            done = bool(progress.get(title, False))
            st.markdown(f"**{title}** · {'已完成' if done else '未完成'}")
            st.caption(desc)
            if st.button("标记完成", key=f"cur_{title}"):
                progress[title] = True
                grant_badge(data, "金融学习者", f"完成：{title}")
                commit(data, event="完成课程", reason=title)
                st.success("已完成课程。")
                st.rerun()
    with tab4:
        m = calc_financials(data)
        card = m.get("score_card", {})
        if not card:
            st.info("信用分尚未建立。")
        else:
            st.dataframe(pd.DataFrame([{"项目": k, "分数": v, "满分": 200 if k != "负债纪律" else 100} for k, v in card.items()]), use_container_width=True, hide_index=True)


def page_reports(data: Dict[str, Any]) -> None:
    st.subheader("报表与数据")
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["家庭金融仪表盘", "资产配置", "周报", "AI CFO 月报", "数据备份"])
    with tab1:
        a = behavior_analytics(data)
        st.dataframe(pd.DataFrame([a]), use_container_width=True, hide_index=True)
        st.markdown("#### 父母干预建议")
        for i, s in enumerate(parent_intervention_suggestions(data), 1):
            st.write(f"**{i}.** {s}")
    with tab2:
        m = calc_financials(data)
        balances = m["balances"]
        total = sum(v for v in balances.values() if v > 0)
        rows = []
        for k, conf in data.get("accounts", {}).items():
            bal = balances.get(k, 0.0)
            target = fnum(conf.get("target_ratio"))
            current = bal / total if total > 0 else 0
            rows.append({"账户": conf.get("name", k), "余额": money(bal), "当前比例": percent(current), "目标比例": percent(target), "偏离": percent(current - target)})
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    with tab3:
        if st.button("生成本周周报", type="primary"):
            report = generate_weekly_report(data)
            data.setdefault("weekly_reports", []).append(report)
            data["weekly_reports"] = data["weekly_reports"][-52:]
            commit(data, event="生成周报", reason=report["week_start"])
            st.success("周报已生成。")
            st.rerun()
        for r in sorted(data.get("weekly_reports", []), key=lambda x: x.get("week_start", ""), reverse=True):
            with st.expander(f"{r.get('week_start')} 周报"):
                st.text(r.get("content", ""))
    with tab4:
        if st.button("生成 AI CFO 月报", type="primary"):
            report = generate_monthly_cfo_report(data)
            data.setdefault("monthly_cfo_reports", []).append(report)
            data["monthly_cfo_reports"] = data["monthly_cfo_reports"][-36:]
            commit(data, event="生成 AI CFO 月报", reason=report["month"])
            st.success("月报已生成。")
            st.rerun()
        for r in sorted(data.get("monthly_cfo_reports", []), key=lambda x: x.get("month", ""), reverse=True):
            with st.expander(f"{r.get('month')} 月报"):
                st.text(r.get("content", ""))
    with tab5:
        st.download_button("下载完整 JSON 备份", json.dumps(normalize_data(data), ensure_ascii=False, indent=2).encode("utf-8"), "asu_bank_backup.json", "application/json")
        if st.button("手动刷新 Google Sheet 展示页"):
            try:
                sync_readable_sheets(data)
                st.success("已刷新。")
            except Exception as e:
                st.error(f"刷新失败：{e}")


def page_settings(data: Dict[str, Any]) -> None:
    st.subheader("系统设置")
    tab1, tab2, tab3 = st.tabs(["账户设置", "风控规则", "商户与预算"])
    with tab1:
        s = data.setdefault("settings", {})
        with st.form("settings_form"):
            owner = st.text_input("账户名称", s.get("owner", "阿苏"))
            start_cash = st.number_input("初始现金", value=fnum(s.get("start_cash")), step=1.0, format="%.2f")
            start_savings = st.number_input("初始储蓄", value=fnum(s.get("start_savings")), step=1.0, format="%.2f")
            base_score = st.number_input("基础信用分", min_value=0, max_value=850, value=inum(s.get("base_score"), 0))
            cash_floor = st.number_input("最低现金安全线", value=fnum(s.get("cash_floor")), step=1.0, format="%.2f")
            ok = st.form_submit_button("保存设置", type="primary", disabled=not is_admin())
        if ok:
            s["owner"] = owner; s["start_cash"] = start_cash; s["start_savings"] = start_savings; s["base_score"] = int(base_score); s["cash_floor"] = cash_floor
            commit(data, event="保存设置", reason="账户设置")
            st.success("已保存。")
            st.rerun()
    with tab2:
        rules = data.setdefault("rules", {})
        with st.form("rules_form"):
            new_rules = {}
            for k, v in rules.items():
                new_rules[k] = st.number_input(k, value=fnum(v), step=0.01, format="%.4f")
            ok = st.form_submit_button("保存规则", type="primary", disabled=not is_admin())
        if ok:
            data["rules"] = new_rules
            data["settings"]["approval_threshold"] = fnum(new_rules.get("approval_threshold"), data["settings"].get("approval_threshold", 15))
            commit(data, event="保存风控规则", reason="规则更新")
            st.success("规则已保存。")
            st.rerun()
    with tab3:
        st.markdown("#### 预算")
        with st.form("budget_form"):
            new_budgets = {}
            for cat, limit in data.get("budgets", {}).items():
                new_budgets[cat] = st.number_input(f"{cat} 月度预算", value=fnum(limit), min_value=0.0, step=1.0, format="%.2f", key=f"budget_{cat}")
            new_cat = st.text_input("新增分类")
            new_limit = st.number_input("新增预算", min_value=0.0, step=1.0, format="%.2f")
            ok = st.form_submit_button("保存预算", disabled=not is_parent())
        if ok:
            if new_cat.strip():
                new_budgets[new_cat.strip()] = new_limit
            data["budgets"] = new_budgets
            commit(data, event="保存预算", reason="预算更新")
            st.success("预算已保存。")
            st.rerun()


def main() -> None:
    st.set_page_config(page_title=APP_NAME, page_icon="🏦", layout="wide", initial_sidebar_state="expanded")
    inject_css()
    login_gate()
    data = get_data()
    update_wishlist_status(data)

    section = st.sidebar.radio(
        "导航",
        ["首页仪表盘", "儿童模式", "家长驾驶舱", "审批与交易", "成长与课程", "报表与数据", "系统设置"],
        index=0,
    )
    st.sidebar.caption(f"存储：{st.session_state.get('storage_backend', '未知')}")

    if section == "首页仪表盘":
        page_home(data)
    elif section == "儿童模式":
        render_hero(data)
        page_child(data)
    elif section == "家长驾驶舱":
        render_hero(data)
        page_parent(data)
    elif section == "审批与交易":
        render_hero(data)
        page_approval_and_transactions(data)
    elif section == "成长与课程":
        render_hero(data)
        page_growth(data)
    elif section == "报表与数据":
        render_hero(data)
        page_reports(data)
    elif section == "系统设置":
        render_hero(data)
        page_settings(data)


if __name__ == "__main__":
    main()
