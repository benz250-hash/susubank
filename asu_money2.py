# -*- coding: utf-8 -*-
"""
阿苏私人银行 2.0
Google Sheet 主数据 + 可读报表页签版

存储设计：
1. state 是唯一主数据源，A2/B2 存完整 JSON。
2. transactions / budgets / goals / merchants / settings / summary 是自动生成的可读报表。
3. 程序只读取 state，不读取展示页，避免手工改表造成数据错乱。
4. Google Sheet 优先；失败时退回本地 asu_money2_data.json。
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

APP_NAME = "阿苏私人银行 2.0"
SCHEMA_VERSION = "asu-bank-gsheet-readable-v1"
DATA_PATH = Path("asu_money2_data.json")
SESSION_KEY = "asu_bank_state_readable_v1"
STATE_WORKSHEET = "state"
STATE_KEY = "bank_data"

GSHEET_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

REPORT_SHEETS = [
    "summary",
    "transactions",
    "budgets",
    "goals",
    "merchants",
    "settings",
]


# ============================================================
# 1. 空库结构
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
            "created_at": datetime.now().isoformat(timespec="seconds"),
        },
        "budgets": {},
        "goals": [],
        "merchants": [],
        "transactions": [],
    }


# ============================================================
# 2. 通用工具
# ============================================================

def uid() -> str:
    return str(uuid.uuid4())


def today_str() -> str:
    return date.today().isoformat()


def month_str(d: Optional[date] = None) -> str:
    return (d or date.today()).strftime("%Y-%m")


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
    value = fnum(x)
    if currency == "USD":
        return f"${value:,.2f}"
    return f"{currency} {value:,.2f}"


def percent(x: Any) -> str:
    return f"{fnum(x) * 100:.0f}%"


def tx_month(tx: Dict[str, Any]) -> str:
    return str(tx.get("date", ""))[:7]


def score_text(score: Optional[int]) -> str:
    return "未建立" if score is None else str(score)


def has_activity(data: Dict[str, Any]) -> bool:
    s = data.get("settings", {})
    return (
        fnum(s.get("start_cash")) != 0
        or fnum(s.get("start_savings")) != 0
        or bool(data.get("transactions"))
        or bool(data.get("budgets"))
        or bool(data.get("goals"))
        or bool(data.get("merchants"))
    )


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
    if result in {"批准", "通过", "开放"}:
        return "ok"
    if result in {"延迟", "限额通过", "暂缓开放", "观察", "未建立"}:
        return "warn"
    return "bad"


def category_options(data: Dict[str, Any]) -> List[str]:
    cats = list(data.get("budgets", {}).keys())
    return sorted(set(cats + ["零花钱", "家庭贷款", "其他"])) or ["其他"]


def spending_categories(data: Dict[str, Any]) -> List[str]:
    cats = list(data.get("budgets", {}).keys())
    return cats if cats else ["其他"]


# ============================================================
# 3. 数据标准化
# ============================================================

def normalize_data(raw: Dict[str, Any]) -> Dict[str, Any]:
    if not isinstance(raw, dict):
        return empty_data()

    base = empty_data()

    settings = base["settings"]
    incoming = raw.get("settings", {})
    if isinstance(incoming, dict):
        settings.update(incoming)
    settings["owner"] = settings.get("owner") or "阿苏"
    settings["currency"] = settings.get("currency") or "USD"
    settings["start_cash"] = fnum(settings.get("start_cash"))
    settings["start_savings"] = fnum(settings.get("start_savings"))
    settings["base_score"] = inum(settings.get("base_score"), 0)
    settings["cash_floor"] = fnum(settings.get("cash_floor"))
    settings["loan_asset_limit"] = fnum(settings.get("loan_asset_limit"), 0.45)
    settings["hard_loan_asset_limit"] = fnum(settings.get("hard_loan_asset_limit"), 0.65)
    base["settings"] = settings

    budgets: Dict[str, float] = {}
    for k, v in (raw.get("budgets") or {}).items():
        name = str(k).strip()
        if name:
            if isinstance(v, dict):
                budgets[name] = fnum(v.get("limit"))
            else:
                budgets[name] = fnum(v)
    base["budgets"] = budgets

    goals: List[Dict[str, Any]] = []
    for g in raw.get("goals") or []:
        if not isinstance(g, dict):
            continue
        goals.append({
            "id": g.get("id") or uid(),
            "name": g.get("name") or "未命名目标",
            "target": fnum(g.get("target")),
            "current": fnum(g.get("current")),
            "deadline": g.get("deadline") or "",
            "category": g.get("category") or "其他",
        })
    base["goals"] = goals

    merchants: List[Dict[str, Any]] = []
    for m in raw.get("merchants") or []:
        if not isinstance(m, dict):
            continue
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

    transactions: List[Dict[str, Any]] = []
    for tx in raw.get("transactions") or []:
        if not isinstance(tx, dict):
            continue
        transactions.append({
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
    base["transactions"] = transactions

    base["schema_version"] = SCHEMA_VERSION
    return base


# ============================================================
# 4. Google Sheet 存储层
# ============================================================

def gsheet_enabled() -> bool:
    if gspread is None or Credentials is None:
        return False
    try:
        return "gcp_service_account" in st.secrets and "SHEET_NAME" in st.secrets
    except Exception:
        return False


def get_secret_value(key: str, default: str = "") -> str:
    try:
        if key in st.secrets:
            return str(st.secrets[key])
    except Exception:
        pass
    return os.getenv(key, default)


def get_gsheet_client():
    if not gsheet_enabled():
        raise RuntimeError("Google Sheet 未配置，或 gspread / google-auth 未安装。")
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=GSHEET_SCOPES)
    return gspread.authorize(creds)


def get_spreadsheet():
    client = get_gsheet_client()
    return client.open(str(st.secrets["SHEET_NAME"]))


def get_or_create_worksheet(spreadsheet, title: str, rows: int = 100, cols: int = 20):
    try:
        return spreadsheet.worksheet(title)
    except Exception:
        return spreadsheet.add_worksheet(title=title, rows=rows, cols=cols)


def replace_worksheet(ws, rows: List[List[Any]]) -> None:
    ws.clear()
    if rows:
        ws.update("A1", rows, value_input_option="RAW")


def get_state_worksheet():
    spreadsheet = get_spreadsheet()
    ws = get_or_create_worksheet(spreadsheet, STATE_WORKSHEET, rows=10, cols=2)
    try:
        header = ws.row_values(1)
        if header[:2] != ["key", "value"]:
            ws.update("A1:B1", [["key", "value"]], value_input_option="RAW")
    except Exception:
        ws.update("A1:B1", [["key", "value"]], value_input_option="RAW")
    return ws


def load_data_from_gsheet() -> Dict[str, Any]:
    ws = get_state_worksheet()
    rows = ws.get_all_records()
    for row in rows:
        if row.get("key") == STATE_KEY:
            raw_text = row.get("value") or ""
            if not str(raw_text).strip():
                data = empty_data()
                save_data_to_gsheet(data)
                return data
            return normalize_data(json.loads(raw_text))

    data = empty_data()
    save_data_to_gsheet(data)
    return data


def save_data_to_gsheet(data: Dict[str, Any]) -> None:
    data = normalize_data(data)
    ws = get_state_worksheet()
    json_text = json.dumps(data, ensure_ascii=False)

    rows = ws.get_all_records()
    target_row = None
    for i, row in enumerate(rows, start=2):
        if row.get("key") == STATE_KEY:
            target_row = i
            break

    if target_row is None:
        ws.append_row([STATE_KEY, json_text], value_input_option="RAW")
    else:
        ws.update_cell(target_row, 2, json_text)

    refresh_report_sheets(data)


def save_data_local(data: Dict[str, Any]) -> None:
    DATA_PATH.write_text(json.dumps(normalize_data(data), ensure_ascii=False, indent=2), encoding="utf-8")


def load_data_local() -> Dict[str, Any]:
    if not DATA_PATH.exists():
        data = empty_data()
        save_data_local(data)
        return data
    try:
        raw = json.loads(DATA_PATH.read_text(encoding="utf-8"))
        if raw.get("schema_version") != SCHEMA_VERSION:
            data = empty_data()
            save_data_local(data)
            return data
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


def commit(data: Dict[str, Any]) -> None:
    st.session_state[SESSION_KEY] = normalize_data(data)
    save_data(st.session_state[SESSION_KEY])


def reload_from_storage() -> None:
    st.session_state.pop(SESSION_KEY, None)


def reset_empty() -> None:
    data = empty_data()
    st.session_state[SESSION_KEY] = data
    save_data(data)


# ============================================================
# 5. 报表页签展开
# ============================================================

def refresh_report_sheets(data: Dict[str, Any]) -> None:
    """把主数据 state 展开成多个可读页签。程序不从这些页签读数据。"""
    if not gsheet_enabled():
        return

    spreadsheet = get_spreadsheet()
    data = normalize_data(data)
    metrics = calc_financials(data)
    ccy = metrics["currency"]

    # summary
    summary_rows = [
        ["指标", "数值", "说明"],
        ["总资产", metrics["total_assets"], "现金 + 储蓄 + 应收贷款本金"],
        ["净资产", metrics["net_assets"], "总资产 - 负债"],
        ["现金余额", metrics["cash"], "可立即使用资金"],
        ["储蓄余额", metrics["savings"], "储蓄账户余额"],
        ["应收贷款本金", metrics["receivables"], "已放出但未收回本金"],
        ["负债余额", metrics["liabilities"], "未偿还借入资金"],
        ["信用分", "未建立" if metrics["credit_score"] is None else metrics["credit_score"], "家庭内部风控评分"],
        ["本月收入", metrics["month_income"], month_str()],
        ["本月消费", metrics["month_expense"], month_str()],
        ["本月支出收入比", metrics["spend_income_ratio"], "支出 / 收入"],
        ["贷款资产占比", metrics["loan_asset_ratio"], "应收贷款本金 / 总资产"],
        ["负债占比", metrics["debt_asset_ratio"], "负债 / 总资产"],
        ["储蓄占比", metrics["savings_asset_ratio"], "储蓄 / 总资产"],
        ["最后同步时间", datetime.now().isoformat(timespec="seconds"), "程序自动写入"],
    ]
    replace_worksheet(get_or_create_worksheet(spreadsheet, "summary", 50, 5), summary_rows)

    # transactions
    tx_rows = [["日期", "类型", "金额", "分类", "对象/商户", "本金回收", "利息收入", "预计回款", "到期日", "备注", "id"]]
    for tx in sorted(data.get("transactions", []), key=lambda x: (str(x.get("date", "")), str(x.get("id", ""))), reverse=True):
        tx_rows.append([
            tx.get("date", ""),
            tx.get("type", ""),
            fnum(tx.get("amount")),
            tx.get("category", ""),
            tx.get("party", ""),
            fnum(tx.get("principal_repaid")),
            fnum(tx.get("interest_received")),
            fnum(tx.get("expected_repayment")),
            tx.get("due_date", ""),
            tx.get("memo", ""),
            tx.get("id", ""),
        ])
    replace_worksheet(get_or_create_worksheet(spreadsheet, "transactions", max(100, len(tx_rows) + 20), 12), tx_rows)

    # budgets
    budget_rows = [["分类", "月度预算", "本月已花", "使用率"]]
    for cat, limit in sorted(data.get("budgets", {}).items()):
        row = metrics["budget_usage"].get(cat, {"spent": 0.0, "ratio": 0.0})
        budget_rows.append([cat, fnum(limit), row["spent"], row["ratio"]])
    replace_worksheet(get_or_create_worksheet(spreadsheet, "budgets", max(30, len(budget_rows) + 10), 6), budget_rows)

    # goals
    goal_rows = [["目标", "分类", "目标金额", "当前金额", "剩余金额", "完成度", "截止日", "剩余天数", "id"]]
    for g in metrics["goals"]:
        goal_rows.append([
            g.get("name", ""),
            g.get("category", ""),
            fnum(g.get("target")),
            fnum(g.get("current")),
            fnum(g.get("remaining")),
            fnum(g.get("progress")),
            g.get("deadline", ""),
            "" if g.get("days_left") is None else g.get("days_left"),
            g.get("id", ""),
        ])
    replace_worksheet(get_or_create_worksheet(spreadsheet, "goals", max(30, len(goal_rows) + 10), 10), goal_rows)

    # merchants
    merchant_rows = [["商户", "分类", "折扣", "最低信用分", "品类预算开放上限", "当前状态", "失败条件", "说明", "id"]]
    for m in data.get("merchants", []):
        access = evaluate_merchant_access(data, m)
        merchant_rows.append([
            m.get("name", ""),
            m.get("category", ""),
            fnum(m.get("discount")),
            inum(m.get("required_score")),
            fnum(m.get("category_budget_cap")),
            access["result"],
            "；".join(access["failed"]),
            m.get("note", ""),
            m.get("id", ""),
        ])
    replace_worksheet(get_or_create_worksheet(spreadsheet, "merchants", max(30, len(merchant_rows) + 10), 10), merchant_rows)

    # settings
    s = data.get("settings", {})
    settings_rows = [["字段", "值", "说明"]]
    labels = {
        "owner": "账户名称",
        "currency": "币种",
        "start_cash": "初始现金",
        "start_savings": "初始储蓄",
        "base_score": "基础信用分，0 表示未建立",
        "cash_floor": "最低现金安全线",
        "loan_asset_limit": "贷款资产建议上限",
        "hard_loan_asset_limit": "贷款资产硬红线",
        "created_at": "账户创建时间",
    }
    for key in ["owner", "currency", "start_cash", "start_savings", "base_score", "cash_floor", "loan_asset_limit", "hard_loan_asset_limit", "created_at"]:
        settings_rows.append([key, s.get(key, ""), labels.get(key, "")])
    replace_worksheet(get_or_create_worksheet(spreadsheet, "settings", 30, 5), settings_rows)


# ============================================================
# 6. 交易对象与账务引擎
# ============================================================

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


def build_loan_book(transactions: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    loans: List[Dict[str, Any]] = []
    sorted_txs = sorted(transactions, key=lambda x: (str(x.get("date", "")), str(x.get("id", ""))))

    for tx in sorted_txs:
        typ = tx.get("type")
        if typ == "放贷":
            principal = fnum(tx.get("amount"))
            if principal <= 0:
                continue
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
        elif typ == "还款":
            borrower = tx.get("party") or ""
            principal_repaid = fnum(tx.get("principal_repaid"))
            if principal_repaid <= 0:
                principal_repaid = fnum(tx.get("amount"))

            candidates = [loan for loan in loans if loan["remaining"] > 0 and (not borrower or loan["borrower"] == borrower)]
            if not candidates:
                candidates = [loan for loan in loans if loan["remaining"] > 0]

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
    usage: Dict[str, Dict[str, float]] = {
        cat: {"limit": fnum(limit), "spent": 0.0, "ratio": 0.0}
        for cat, limit in data.get("budgets", {}).items()
    }

    for tx in data.get("transactions", []):
        if tx.get("type") != "消费":
            continue
        if tx_month(tx) != month:
            continue
        cat = tx.get("category") or "其他"
        usage.setdefault(cat, {"limit": 0.0, "spent": 0.0, "ratio": 0.0})
        usage[cat]["spent"] += fnum(tx.get("amount"))

    for row in usage.values():
        row["ratio"] = row["spent"] / row["limit"] if row["limit"] > 0 else 0.0
    return usage


def calc_goal_status(data: Dict[str, Any]) -> List[Dict[str, Any]]:
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
    overdue_principal: float,
) -> Tuple[Optional[int], List[str]]:
    settings = data.get("settings", {})
    base_score = inum(settings.get("base_score"), 0)
    if base_score <= 0:
        return None, ["账户尚未建立信用分。请在左侧设置基础信用分。"]

    score = float(base_score)
    factors: List[str] = []
    cash_floor = fnum(settings.get("cash_floor"))

    if cash_floor > 0:
        if cash < 0:
            score -= 45
            factors.append("现金余额为负：-45")
        elif cash < cash_floor:
            score -= 16
            factors.append(f"现金低于安全线 {money(cash_floor)}：-16")
        elif cash >= cash_floor * 2:
            score += 6
            factors.append("现金缓冲较充足：+6")

    savings_ratio = savings / total_assets if total_assets > 0 else 0.0
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
        spend_ratio = month_expense / month_income
        if spend_ratio > 1.0:
            score -= 35
            factors.append("本月支出超过收入：-35")
        elif spend_ratio > 0.85:
            score -= 18
            factors.append("本月支出达到收入 85% 以上：-18")
        elif spend_ratio > 0.65:
            score -= 8
            factors.append("本月支出达到收入 65% 以上：-8")
        elif spend_ratio <= 0.45:
            score += 5
            factors.append("本月支出控制在收入 45% 以下：+5")

    loan_ratio = receivables / total_assets if total_assets > 0 else 0.0
    if loan_ratio > 0.65:
        score -= 30
        factors.append("应收贷款本金占总资产超过 65%：-30")
    elif loan_ratio > 0.50:
        score -= 20
        factors.append("应收贷款本金占总资产超过 50%：-20")
    elif loan_ratio > 0.35:
        score -= 9
        factors.append("应收贷款本金占总资产超过 35%：-9")

    debt_ratio = liabilities / total_assets if total_assets > 0 else 0.0
    if debt_ratio > 0.50:
        score -= 45
        factors.append("负债占总资产超过 50%：-45")
    elif debt_ratio > 0.25:
        score -= 25
        factors.append("负债占总资产超过 25%：-25")
    elif liabilities > 0:
        score -= 7
        factors.append("存在未偿还负债：-7")

    overspent = sum(1 for row in budget_usage.values() if row["limit"] > 0 and row["ratio"] > 1.0)
    warning = sum(1 for row in budget_usage.values() if row["limit"] > 0 and 0.85 < row["ratio"] <= 1.0)
    if overspent:
        penalty = min(32, overspent * 10)
        score -= penalty
        factors.append(f"{overspent} 个预算品类已超支：-{penalty}")
    if warning:
        penalty = min(15, warning * 4)
        score -= penalty
        factors.append(f"{warning} 个预算品类接近上限：-{penalty}")
    if overdue_principal > 0:
        score -= 24
        factors.append(f"存在逾期未收回应收本金 {money(overdue_principal)}：-24")
    if not factors:
        factors.append("当前没有明显加分或扣分因素。")
    return int(round(clamp(score, 300, 850))), factors


def calc_financials(data: Dict[str, Any], month: Optional[str] = None) -> Dict[str, Any]:
    month = month or month_str()
    settings = data.get("settings", {})
    currency = settings.get("currency", "USD")
    cash = fnum(settings.get("start_cash"))
    savings = fnum(settings.get("start_savings"))
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
    receivables = sum(fnum(loan.get("remaining")) for loan in loans)

    overdue_principal = 0.0
    today = date.today()
    for loan in loans:
        due = loan.get("due_date") or ""
        remaining = fnum(loan.get("remaining"))
        if remaining <= 0 or not due:
            continue
        try:
            if date.fromisoformat(due) < today:
                overdue_principal += remaining
        except Exception:
            pass

    total_assets = cash + savings + receivables
    net_assets = total_assets - liabilities
    budget_usage = calc_budget_usage(data, month)
    goals = calc_goal_status(data)
    credit_score, score_factors = calc_credit_score(
        data, cash, savings, receivables, liabilities, total_assets,
        month_income, month_expense, budget_usage, overdue_principal
    )

    spend_income_ratio = month_expense / month_income if month_income > 0 else (1.0 if month_expense > 0 else 0.0)
    loan_asset_ratio = receivables / total_assets if total_assets > 0 else 0.0
    debt_asset_ratio = liabilities / total_assets if total_assets > 0 else 0.0
    savings_asset_ratio = savings / total_assets if total_assets > 0 else 0.0

    return {
        "currency": currency,
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
        "budget_usage": budget_usage,
        "goals": goals,
        "loans": loans,
        "overdue_principal": overdue_principal,
        "credit_score": credit_score,
        "score_factors": score_factors,
    }


# ============================================================
# 7. 审批、风险、AI
# ============================================================

def simulate_transaction(data: Dict[str, Any], tx: Dict[str, Any]) -> Dict[str, Any]:
    before = calc_financials(data)
    after_data = copy.deepcopy(data)
    after_data.setdefault("transactions", []).append(tx)
    after = calc_financials(after_data)
    score_delta = None if before["credit_score"] is None or after["credit_score"] is None else after["credit_score"] - before["credit_score"]
    return {
        "before": before,
        "after": after,
        "delta": {
            "cash": after["cash"] - before["cash"],
            "savings": after["savings"] - before["savings"],
            "receivables": after["receivables"] - before["receivables"],
            "liabilities": after["liabilities"] - before["liabilities"],
            "total_assets": after["total_assets"] - before["total_assets"],
            "net_assets": after["net_assets"] - before["net_assets"],
            "credit_score": score_delta,
        },
        "tx": tx,
    }


def usage_for(metrics: Dict[str, Any], category: str) -> Dict[str, float]:
    return metrics.get("budget_usage", {}).get(category, {"limit": 0.0, "spent": 0.0, "ratio": 0.0})


def evaluate_purchase_decision(data: Dict[str, Any], amount: float, category: str, description: str, merchant: str = "") -> Dict[str, Any]:
    tx = make_tx("消费", amount, category, party=merchant, memo=description)
    sim = simulate_transaction(data, tx)
    before, after, delta = sim["before"], sim["after"], sim["delta"]
    budget = usage_for(after, category)
    cash_floor = fnum(data.get("settings", {}).get("cash_floor"))
    result = "批准"
    reasons: List[str] = []

    if amount <= 0:
        result = "拒绝"
        reasons.append("购买金额必须大于 0。")
    if after["cash"] < 0:
        result = "拒绝"
        reasons.append("购买后现金余额为负，触发硬性拒绝。")
    if budget["limit"] > 0 and budget["ratio"] > 1.20:
        result = "拒绝"
        reasons.append(f"{category}预算使用率将达到 {percent(budget['ratio'])}，超过 120% 红线。")
    if after["credit_score"] is not None and after["credit_score"] < 650:
        result = "拒绝"
        reasons.append("交易后信用分低于 650，进入高风险区。")

    if result != "拒绝":
        flags = []
        if cash_floor > 0 and after["cash"] < cash_floor:
            flags.append(f"购买后现金余额低于安全线 {money(cash_floor)}。")
        if budget["limit"] > 0 and budget["ratio"] >= 0.90:
            flags.append(f"{category}预算使用率将达到 {percent(budget['ratio'])}，接近或超过上限。")
        if delta["credit_score"] is not None and delta["credit_score"] <= -8:
            flags.append(f"信用分预计下降 {abs(delta['credit_score'])} 分。")
        if after["month_income"] > 0 and after["spend_income_ratio"] >= 0.85:
            flags.append(f"本月支出/收入比将达到 {percent(after['spend_income_ratio'])}。")
        if flags:
            result = "延迟"
            reasons.extend(flags)

    if result == "批准":
        reasons.append("现金余额、预算使用率、信用分变化均未触发硬性风控限制。")
    action = "不要购买。先补现金、建预算，或等收入到账后再申请。" if result == "拒绝" else "建议延迟到下一笔收入到账后再买，或者把金额拆成两周预算。" if result == "延迟" else "可以购买，但需要正式入账；非必要消费不要动用储蓄目标资金。"

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
    }


def evaluate_loan_decision(data: Dict[str, Any], principal: float, borrower: str, expected_repayment: float, due_date: str, memo: str = "") -> Dict[str, Any]:
    tx = make_tx("放贷", principal, "家庭贷款", party=borrower, memo=memo, expected_repayment=expected_repayment, due_date=due_date)
    sim = simulate_transaction(data, tx)
    before, after, delta = sim["before"], sim["after"], sim["delta"]
    settings = data.get("settings", {})
    cash_floor = fnum(settings.get("cash_floor"))
    soft_limit = fnum(settings.get("loan_asset_limit"), 0.45)
    hard_limit = fnum(settings.get("hard_loan_asset_limit"), 0.65)
    result = "通过"
    reasons: List[str] = []
    max_by_cash = max(0.0, before["cash"] - cash_floor) if cash_floor > 0 else max(0.0, before["cash"])
    safe_capacity = max(0.0, soft_limit * max(before["total_assets"], 1.0) - before["receivables"])
    suggested_limit = max(0.0, min(principal, max_by_cash, safe_capacity))

    if principal <= 0:
        result = "拒绝"
        reasons.append("放贷本金必须大于 0。")
    if after["cash"] < 0:
        result = "拒绝"
        reasons.append("放贷后现金余额为负，流动性不足。")
    if after["total_assets"] > 0 and after["loan_asset_ratio"] > hard_limit:
        result = "拒绝"
        reasons.append(f"放贷后应收贷款本金占总资产 {percent(after['loan_asset_ratio'])}，超过 {percent(hard_limit)} 红线。")
    if after["credit_score"] is not None and after["credit_score"] < 650:
        result = "拒绝"
        reasons.append("放贷后信用分低于 650。")

    if result != "拒绝":
        flags = []
        if cash_floor > 0 and after["cash"] < cash_floor:
            flags.append(f"放贷后现金低于安全线 {money(cash_floor)}。")
        if after["total_assets"] > 0 and after["loan_asset_ratio"] > soft_limit:
            flags.append(f"放贷后应收贷款本金占总资产 {percent(after['loan_asset_ratio'])}，超过建议线 {percent(soft_limit)}。")
        if principal > suggested_limit and suggested_limit > 0:
            flags.append(f"建议最高放贷金额为 {money(suggested_limit)}。")
        if not due_date:
            flags.append("没有填写预计还款日，回款纪律不足。")
        if flags:
            result = "限额通过"
            reasons.extend(flags)

    if result == "通过":
        reasons.append("放贷后现金缓冲、贷款资产占比、信用分均未触发硬性风控限制。")
    action = "本次不建议放贷。先增加现金余额，或降低放贷金额。" if result == "拒绝" else f"建议限额放贷，最高不超过 {money(suggested_limit)}；必须写清还款日。" if result == "限额通过" else "可以放贷，但回款本金到账后应优先补充现金余额。"

    return {
        "kind": "贷款审批",
        "result": result,
        "reasons": reasons,
        "action": action,
        "metrics": {
            "拟借出本金": principal,
            "借款人": borrower or "未填写",
            "预计回款": expected_repayment,
            "当前现金余额": before["cash"],
            "放贷后现金余额": after["cash"],
            "当前应收贷款本金": before["receivables"],
            "放贷后应收贷款本金": after["receivables"],
            "放贷后贷款资产占比": after["loan_asset_ratio"],
            "当前信用分": before["credit_score"],
            "放贷后信用分": after["credit_score"],
            "信用分变化": delta["credit_score"],
            "建议最高放贷金额": suggested_limit,
        },
        "pending_tx": tx,
    }


def estimate_score_after_transaction(data: Dict[str, Any], tx: Dict[str, Any]) -> Dict[str, Any]:
    sim = simulate_transaction(data, tx)
    return {"current_score": sim["before"]["credit_score"], "after_score": sim["after"]["credit_score"], "change": sim["delta"]["credit_score"]}


def build_risk_radar(data: Dict[str, Any]) -> Dict[str, List[Dict[str, str]]]:
    m = calc_financials(data)
    risks = {"高风险": [], "中风险": [], "低风险": []}
    if not has_activity(data):
        risks["低风险"].append({"title": "账户为空", "detail": "暂无交易、暂无预算、暂无目标。请先建立收入、预算或信用分。"})
        return risks

    cash_floor = fnum(data.get("settings", {}).get("cash_floor"))
    if m["cash"] < 0:
        risks["高风险"].append({"title": "现金余额为负", "detail": f"当前现金 {money(m['cash'])}，应暂停所有非必要消费。"})
    elif cash_floor > 0 and m["cash"] < cash_floor:
        risks["中风险"].append({"title": "现金缓冲偏低", "detail": f"当前现金 {money(m['cash'])}，低于安全线 {money(cash_floor)}。"})
    else:
        risks["低风险"].append({"title": "现金余额未触发警报", "detail": f"当前现金 {money(m['cash'])}。"})

    if m["month_income"] > 0:
        if m["spend_income_ratio"] > 1.0:
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
            risks["中风险"].append({"title": "存在未偿还负债", "detail": f"负债余额 {money(m['liabilities'])}。"})
    else:
        risks["低风险"].append({"title": "无未偿还负债", "detail": "净资产未被负债侵蚀。"})

    for cat, row in m["budget_usage"].items():
        if row["limit"] <= 0:
            continue
        if row["ratio"] > 1.0:
            risks["高风险"].append({"title": f"{cat}预算超支", "detail": f"使用率 {percent(row['ratio'])}。"})
        elif row["ratio"] > 0.85:
            risks["中风险"].append({"title": f"{cat}预算接近上限", "detail": f"使用率 {percent(row['ratio'])}。"})

    if m["overdue_principal"] > 0:
        risks["高风险"].append({"title": "存在逾期贷款", "detail": f"逾期本金 {money(m['overdue_principal'])}，暂停新增放贷。"})
    return risks


def build_weekly_plan(data: Dict[str, Any]) -> List[str]:
    if not has_activity(data):
        return ["先建立第一笔收入或现金余额。", "新增至少 3 个预算分类，例如：游戏、甜品、学习。", "如果要使用信用分审批，请在左侧设置基础信用分。"]
    m = calc_financials(data)
    actions: List[str] = []
    overspent = [cat for cat, row in m["budget_usage"].items() if row["limit"] > 0 and row["ratio"] > 1.0]
    near = [cat for cat, row in m["budget_usage"].items() if row["limit"] > 0 and 0.85 < row["ratio"] <= 1.0]
    if overspent:
        actions.append(f"本周暂停 {', '.join(overspent)} 类非必要消费。")
    elif near:
        actions.append(f"本周 {', '.join(near)} 类消费必须先审批。")
    if m["loan_asset_ratio"] > 0.45:
        actions.append("优先收回应收贷款本金，目标至少收回 $10。")
    cash_floor = fnum(data.get("settings", {}).get("cash_floor"))
    if cash_floor > 0 and m["cash"] < cash_floor:
        actions.append("下一笔收入先补现金缓冲，不立刻消费。")
    if m["savings_asset_ratio"] < 0.25 and m["month_income"] > 0:
        actions.append("下一笔收入的 30% 转入储蓄。")
    if m["liabilities"] > 0:
        actions.append("优先偿还负债，不新增借入资金。")
    return actions[:5] if actions else ["维持当前消费节奏，但非必要消费继续走审批。", "下一笔收入至少 20% 转入储蓄。", "有人还款后，先补现金余额，再考虑消费。"]


def evaluate_merchant_access(data: Dict[str, Any], merchant: Dict[str, Any]) -> Dict[str, Any]:
    m = calc_financials(data)
    category = merchant.get("category") or "其他"
    usage = usage_for(m, category)
    required_score = inum(merchant.get("required_score"), 0)
    category_cap = fnum(merchant.get("category_budget_cap"), 1.0)
    current_score = m["credit_score"]
    checks = {
        "信用分达标": current_score is not None and current_score >= required_score,
        "现金余额为正": m["cash"] > 0,
        "本月支出收入比不高于 90%": m["spend_income_ratio"] <= 0.90,
        "该品类预算未过高": usage["ratio"] <= category_cap if usage["limit"] > 0 else True,
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
    return {"result": result, "message": message, "failed": failed, "checks": checks, "metrics": {"当前信用分": current_score, "要求信用分": required_score, "现金余额": m["cash"], "本月支出收入比": m["spend_income_ratio"], "该品类预算使用率": usage["ratio"], "品类开放上限": category_cap}}


def local_report(decision: Dict[str, Any]) -> str:
    lines = [f"### {decision.get('kind', '风控报告')}", f"**审批结论：{decision.get('result')}**", "", "#### 关键硬指标"]
    for k, v in decision.get("metrics", {}).items():
        lines.append(f"- {k}：{pretty_value(k, v)}")
    lines += ["", "#### 风控原因"]
    for reason in decision.get("reasons", []):
        lines.append(f"- {reason}")
    lines += ["", "#### 操作指令", decision.get("action", "")]
    return "\n".join(lines)


def deepseek_decision_report(decision: Dict[str, Any]) -> str:
    api_key = get_secret_value("DEEPSEEK_API_KEY")
    model = get_secret_value("DEEPSEEK_MODEL", "deepseek-v4-flash")
    if not api_key or requests is None:
        return local_report(decision)
    payload = {"kind": decision.get("kind"), "result": decision.get("result"), "metrics": decision.get("metrics"), "reasons": decision.get("reasons"), "action": decision.get("action")}
    system_prompt = "你是阿苏私人银行2.0的中文私人银行客户经理。只能解释Python已计算的硬指标，禁止编造数字，禁止改变审批结论。输出：审批结论、关键数字、风控原因、操作指令。语气克制清楚。"
    try:
        resp = requests.post(
            "https://api.deepseek.com/chat/completions",
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            json={"model": model, "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": json.dumps(payload, ensure_ascii=False, indent=2)}], "temperature": 0.2, "max_tokens": 900},
            timeout=20,
        )
        resp.raise_for_status()
        return resp.json()["choices"][0]["message"]["content"]
    except Exception as exc:
        return local_report(decision) + f"\n\n> DeepSeek 调用失败，已切换本地报告：{exc}"


# ============================================================
# 8. 自然语言入口
# ============================================================

def extract_amounts(text: str) -> List[float]:
    if not text:
        return []
    patterns = [r"\$\s*([0-9]+(?:\.[0-9]+)?)", r"([0-9]+(?:\.[0-9]+)?)\s*(?:美元|刀|块|元)"]
    nums: List[float] = []
    for p in patterns:
        for m in re.finditer(p, text, flags=re.I):
            nums.append(float(m.group(1)))
    if not nums:
        for m in re.finditer(r"([0-9]+(?:\.[0-9]+)?)", text):
            nums.append(float(m.group(1)))
    return nums


def infer_category(text: str) -> str:
    s = text.lower()
    rules = {
        "游戏": ["minecraft", "robux", "游戏", "道具", "皮肤", "steam", "switch"],
        "甜品": ["甜品", "奶茶", "冰淇淋", "蛋糕", "糖", "饮料"],
        "玩具": ["玩具", "lego", "乐高", "手办", "娃娃"],
        "学习": ["书", "学习", "课程", "文具", "作业", "训练"],
        "宠物": ["猫", "猫粮", "猫砂", "宠物", "罐头"],
        "餐饮": ["饭", "餐", "披萨", "pizza", "汉堡", "麦当劳"],
    }
    for category, keys in rules.items():
        if any(k in s for k in keys):
            return category
    return "其他"


def infer_borrower(text: str) -> str:
    m = re.search(r"借给(.+?)(?:\$|[0-9]|，|,|。|$)", text)
    if m:
        name = re.sub(r"\s+", "", m.group(1))
        return name[:12] if name else "未填写"
    for name in ["爸爸", "妈妈", "法法", "同学", "朋友"]:
        if name in text:
            return name
    return "未填写"


def parse_natural_request(data: Dict[str, Any], text: str) -> Dict[str, Any]:
    amounts = extract_amounts(text)
    is_loan = any(x in text for x in ["借给", "放贷", "贷款给"])
    is_purchase = any(x in text for x in ["想买", "购买", "买", "消费", "花"])
    if is_loan:
        principal = amounts[0] if amounts else 0.0
        expected = amounts[1] if len(amounts) >= 2 else principal
        return evaluate_loan_decision(data, principal, infer_borrower(text), expected, "", text)
    if is_purchase or amounts:
        amount = amounts[0] if amounts else 0.0
        return evaluate_purchase_decision(data, amount, infer_category(text), text, "自然语言输入")
    return {"kind": "AI 决策中心", "result": "观察", "reasons": ["没有识别出明确金额或明确动作。"], "action": "请按格式输入：我想买 $18 的 Minecraft 道具；或：借给爸爸 $20，预计一周后还 $22。", "metrics": {}, "pending_tx": None}


# ============================================================
# 9. UI
# ============================================================

def inject_css() -> None:
    st.markdown(
        """
        <style>
        .main .block-container { max-width: 1280px; padding-top: 1.2rem; padding-bottom: 2rem; }
        .hero { background: linear-gradient(135deg, #003C71 0%, #005EB8 50%, #0A74DA 100%); color: white; padding: 26px 30px; border-radius: 24px; margin-bottom: 18px; box-shadow: 0 16px 38px rgba(0, 62, 130, 0.25); }
        .hero h1 { margin: 0; font-size: 34px; letter-spacing: 0.3px; }
        .hero p { margin: 8px 0 0; opacity: 0.93; font-size: 16px; }
        .card { border: 1px solid #D6E4F5; border-radius: 18px; background: white; padding: 18px; box-shadow: 0 8px 22px rgba(15, 23, 42, 0.06); min-height: 112px; }
        .label { color: #6B7280; font-size: 13px; margin-bottom: 6px; }
        .value { font-size: 28px; font-weight: 780; color: #111827; line-height: 1.12; }
        .sub { color: #6B7280; font-size: 12px; margin-top: 6px; }
        .decision { border: 1px solid #D6E4F5; border-radius: 18px; background: white; padding: 18px; margin: 10px 0; box-shadow: 0 8px 22px rgba(15, 23, 42, 0.05); }
        .ok { border-left: 7px solid #166534; } .warn { border-left: 7px solid #B45309; } .bad { border-left: 7px solid #B91C1C; }
        .pill { display: inline-block; padding: 4px 10px; border-radius: 999px; background: #EAF3FF; color: #003C71; border: 1px solid #B9D7F6; font-size: 12px; font-weight: 700; margin-bottom: 8px; }
        .small { color: #6B7280; font-size: 13px; }
        div[data-testid="stMetric"] { border: 1px solid #D6E4F5; border-radius: 16px; padding: 12px 14px; background: white; box-shadow: 0 6px 18px rgba(15,23,42,0.05); }
        .stTabs [data-baseweb="tab-list"] { gap: 5px; }
        .stTabs [data-baseweb="tab"] { border-radius: 999px; background: #F2F6FB; padding: 8px 13px; }
        .stTabs [aria-selected="true"] { background: #005EB8 !important; color: white !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_hero(data: Dict[str, Any]) -> None:
    owner = data.get("settings", {}).get("owner", "阿苏")
    st.markdown(f"""<div class="hero"><h1>{APP_NAME}</h1><p>{owner} 的家庭版私人银行：交易前审批 · 交易后模拟 · 风险雷达 · Google Sheet 可读报表</p></div>""", unsafe_allow_html=True)
    st.caption(f"当前存储：{st.session_state.get('storage_backend', '未知')}；主数据页：state；展示页：summary / transactions / budgets / goals / merchants / settings")


def metric_card(label: str, value: str, sub: str = "") -> None:
    st.markdown(f"""<div class="card"><div class="label">{label}</div><div class="value">{value}</div><div class="sub">{sub}</div></div>""", unsafe_allow_html=True)


def render_decision(decision: Dict[str, Any], data: Dict[str, Any], key: str) -> None:
    css = decision_class(decision.get("result", "观察"))
    st.markdown(f"""<div class="decision {css}"><div class="pill">{decision.get('kind', '审批')}</div><h3 style="margin: 4px 0 8px;">审批结果：{decision.get('result')}</h3><p class="small">{decision.get('action', '')}</p></div>""", unsafe_allow_html=True)
    if decision.get("reasons"):
        st.markdown("#### 风控原因")
        for reason in decision["reasons"]:
            st.write(f"- {reason}")
    if decision.get("metrics"):
        st.dataframe(pd.DataFrame([{"指标": k, "值": pretty_value(k, v)} for k, v in decision["metrics"].items()]), use_container_width=True, hide_index=True)
    with st.expander("AI 客户经理报告", expanded=True):
        st.markdown(deepseek_decision_report(decision))
    tx = decision.get("pending_tx")
    if tx:
        disabled = decision.get("result") == "拒绝"
        if st.button("确认入账" if not disabled else "拒绝结果不可入账", disabled=disabled, type="primary", key=key):
            data.setdefault("transactions", []).append(tx)
            commit(data)
            st.success("已入账并同步到 state 和展示页。")
            st.rerun()


def transactions_df(data: Dict[str, Any]) -> pd.DataFrame:
    rows = []
    for tx in data.get("transactions", []):
        rows.append({"日期": tx.get("date"), "类型": tx.get("type"), "金额": fnum(tx.get("amount")), "分类": tx.get("category"), "对象/商户": tx.get("party"), "本金回收": fnum(tx.get("principal_repaid")), "利息收入": fnum(tx.get("interest_received")), "预计回款": fnum(tx.get("expected_repayment")), "到期日": tx.get("due_date"), "备注": tx.get("memo"), "id": tx.get("id")})
    if not rows:
        return pd.DataFrame()
    return pd.DataFrame(rows).sort_values("日期", ascending=False)


# ============================================================
# 10. 页面
# ============================================================

def page_home(data: Dict[str, Any]) -> None:
    m = calc_financials(data)
    ccy = m["currency"]
    c1, c2, c3, c4 = st.columns(4)
    with c1: metric_card("总资产", money(m["total_assets"], ccy), "现金 + 储蓄 + 应收贷款本金")
    with c2: metric_card("净资产", money(m["net_assets"], ccy), "总资产 - 负债")
    with c3: metric_card("现金余额", money(m["cash"], ccy), "可立即使用资金")
    with c4: metric_card("信用分", score_text(m["credit_score"]), "家庭内部风控评分，非 FICO")
    c5, c6, c7, c8 = st.columns(4)
    with c5: metric_card("储蓄余额", money(m["savings"], ccy), f"储蓄占比 {percent(m['savings_asset_ratio'])}")
    with c6: metric_card("应收贷款本金", money(m["receivables"], ccy), f"贷款资产占比 {percent(m['loan_asset_ratio'])}")
    with c7: metric_card("负债余额", money(m["liabilities"], ccy), f"负债占比 {percent(m['debt_asset_ratio'])}")
    with c8: metric_card("本月支出/收入", percent(m["spend_income_ratio"]), f"收入 {money(m['month_income'])} / 支出 {money(m['month_expense'])}")
    st.divider()
    left, right = st.columns([1.05, 1])
    with left:
        st.subheader("AI 本周行动计划")
        for i, action in enumerate(build_weekly_plan(data), 1):
            st.write(f"**{i}.** {action}")
        st.subheader("信用分因子")
        for factor in m["score_factors"]:
            st.write(f"- {factor}")
    with right:
        st.subheader("本月预算使用率")
        rows = [{"分类": cat, "已花": row["spent"], "预算": row["limit"], "使用率": row["ratio"]} for cat, row in m["budget_usage"].items()]
        df = pd.DataFrame(rows)
        if df.empty:
            st.info("暂无预算。请先到预算管理新增预算分类。")
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
        expected_repayment = 0.0
        principal_repaid = 0.0
        interest_received = 0.0
        due_date = ""
        if tx_type == "放贷":
            expected_repayment = st.number_input("预计回款总额", min_value=0.0, step=1.0, format="%.2f")
            due_date = st.date_input("预计还款日", value=date.today() + timedelta(days=7)).isoformat()
        if tx_type == "还款":
            principal_repaid = st.number_input("本金回收", min_value=0.0, step=1.0, format="%.2f")
            interest_received = st.number_input("利息收入", min_value=0.0, step=0.5, format="%.2f")
            amount = principal_repaid + interest_received
        submitted = st.form_submit_button("保存交易", type="primary")
    if submitted:
        data.setdefault("transactions", []).append(make_tx(tx_type, amount, category, party=party, memo=memo, tx_date=d.isoformat(), expected_repayment=expected_repayment, principal_repaid=principal_repaid, interest_received=interest_received, due_date=due_date))
        commit(data)
        st.success("交易已保存，state 和展示页已同步。")
        st.rerun()


def page_ai_center(data: Dict[str, Any]) -> None:
    st.subheader("AI 决策中心")
    text = st.text_area("输入请求", value="我想买 $18 的 Minecraft 道具", height=120, placeholder="例如：借给爸爸 $20，预计一周后还 $22")
    if st.button("运行审批", type="primary"):
        st.session_state["latest_ai_decision"] = parse_natural_request(data, text)
    if "latest_ai_decision" in st.session_state:
        render_decision(st.session_state["latest_ai_decision"], data, "confirm_ai_decision")


def page_purchase(data: Dict[str, Any]) -> None:
    st.subheader("消费审批")
    with st.form("purchase_form"):
        c1, c2, c3 = st.columns(3)
        with c1: amount = st.number_input("购买金额", value=18.0, min_value=0.0, step=1.0, format="%.2f")
        with c2: category = st.selectbox("消费分类", spending_categories(data))
        with c3: merchant = st.text_input("商户", "Minecraft 商店")
        desc = st.text_area("购买说明", "Minecraft 道具")
        submitted = st.form_submit_button("模拟消费审批", type="primary")
    if submitted:
        st.session_state["purchase_decision"] = evaluate_purchase_decision(data, amount, category, desc, merchant)
    if "purchase_decision" in st.session_state:
        render_decision(st.session_state["purchase_decision"], data, "confirm_purchase")


def page_loan(data: Dict[str, Any]) -> None:
    st.subheader("贷款审批")
    with st.form("loan_form"):
        c1, c2, c3 = st.columns(3)
        with c1: principal = st.number_input("拟借出本金", value=20.0, min_value=0.0, step=1.0, format="%.2f")
        with c2: borrower = st.text_input("借款人", "爸爸")
        with c3: expected = st.number_input("预计回款总额", value=22.0, min_value=0.0, step=1.0, format="%.2f")
        due = st.date_input("预计还款日", value=date.today() + timedelta(days=7))
        memo = st.text_area("备注", "家庭临时周转")
        submitted = st.form_submit_button("模拟贷款审批", type="primary")
    if submitted:
        st.session_state["loan_decision"] = evaluate_loan_decision(data, principal, borrower, expected, due.isoformat(), memo)
    if "loan_decision" in st.session_state:
        render_decision(st.session_state["loan_decision"], data, "confirm_loan")


def page_risk_radar(data: Dict[str, Any]) -> None:
    st.subheader("风险雷达")
    risks = build_risk_radar(data)
    cols = st.columns(3)
    for col, level in zip(cols, ["高风险", "中风险", "低风险"]):
        with col:
            st.markdown(f"### {level}")
            if not risks[level]: st.info("暂无")
            for item in risks[level]:
                css = "bad" if level == "高风险" else "warn" if level == "中风险" else "ok"
                st.markdown(f"""<div class="decision {css}"><b>{item['title']}</b><p class="small">{item['detail']}</p></div>""", unsafe_allow_html=True)
    st.divider()
    st.subheader("本周行动指令")
    for i, action in enumerate(build_weekly_plan(data), 1):
        st.write(f"**{i}.** {action}")


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
    expense = month_df[month_df["类型"] == "消费"]
    if not expense.empty:
        st.subheader("消费分类")
        st.bar_chart(expense.groupby("分类")["金额"].sum().sort_values(ascending=False))
    st.subheader("明细")
    st.dataframe(month_df.drop(columns=["id"], errors="ignore"), use_container_width=True, hide_index=True)


def page_budget(data: Dict[str, Any]) -> None:
    st.subheader("预算管理")
    with st.form("budget_form"):
        new_budgets = {}
        if data.get("budgets"):
            for cat, limit in data["budgets"].items():
                new_budgets[cat] = st.number_input(f"{cat} 月度预算", value=fnum(limit), min_value=0.0, step=1.0, format="%.2f", key=f"budget_{cat}")
        else:
            st.info("暂无预算分类。请在下面新增。")
        st.markdown("#### 新增分类")
        c1, c2 = st.columns([2, 1])
        with c1: new_cat = st.text_input("分类名称", "")
        with c2: new_limit = st.number_input("预算金额", min_value=0.0, step=1.0, format="%.2f")
        submitted = st.form_submit_button("保存预算", type="primary")
    if submitted:
        if new_cat.strip(): new_budgets[new_cat.strip()] = new_limit
        data["budgets"] = new_budgets
        commit(data)
        st.success("预算已保存，展示页 budgets 已同步。")
        st.rerun()
    rows = [{"分类": cat, "已花": money(row["spent"]), "预算": money(row["limit"]), "使用率": percent(row["ratio"])} for cat, row in calc_financials(data)["budget_usage"].items()]
    if rows: st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


def page_goals(data: Dict[str, Any]) -> None:
    st.subheader("储蓄目标")
    goals = calc_goal_status(data)
    if not goals: st.info("暂无储蓄目标。")
    for goal in goals:
        st.markdown(f"#### {goal['name']}")
        st.progress(min(1.0, goal["progress"]))
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("目标", money(goal["target"]))
        c2.metric("当前", money(goal["current"]))
        c3.metric("完成度", percent(goal["progress"]))
        c4.metric("剩余", money(goal["remaining"]))
        if goal["days_left"] is not None: st.caption(f"截止日：{goal['deadline']}；剩余 {goal['days_left']} 天")
        st.divider()
    with st.form("goal_form"):
        st.markdown("#### 新增储蓄目标")
        name = st.text_input("目标名称", "")
        c1, c2, c3 = st.columns(3)
        with c1: target = st.number_input("目标金额", min_value=0.0, step=5.0, format="%.2f")
        with c2: current = st.number_input("当前金额", min_value=0.0, step=5.0, format="%.2f")
        with c3: category = st.selectbox("关联分类", spending_categories(data))
        deadline = st.date_input("截止日", value=date.today() + timedelta(days=90))
        submitted = st.form_submit_button("新增目标", type="primary")
    if submitted and name.strip():
        data.setdefault("goals", []).append({"id": uid(), "name": name.strip(), "target": target, "current": current, "deadline": deadline.isoformat(), "category": category})
        commit(data)
        st.success("目标已新增，展示页 goals 已同步。")
        st.rerun()


def page_score(data: Dict[str, Any]) -> None:
    st.subheader("信用分")
    m = calc_financials(data)
    st.metric("当前信用分", score_text(m["credit_score"]))
    st.caption("家庭内部风控评分，不是 FICO，不是银行真实征信。")
    st.markdown("#### 当前因子")
    for factor in m["score_factors"]: st.write(f"- {factor}")
    st.divider(); st.markdown("#### 交易前模拟")
    with st.form("score_form"):
        typ = st.selectbox("交易类型", ["消费", "收入", "放贷", "转入储蓄", "借入", "偿还负债"])
        amount = st.number_input("金额", value=10.0, min_value=0.0, step=1.0, format="%.2f")
        category = st.selectbox("分类", category_options(data))
        submitted = st.form_submit_button("模拟", type="primary")
    if submitted:
        score = estimate_score_after_transaction(data, make_tx(typ, amount, category, memo="信用分模拟"))
        c1, c2, c3 = st.columns(3)
        c1.metric("当前信用分", score_text(score["current_score"]))
        c2.metric("交易后信用分", score_text(score["after_score"]))
        c3.metric("变化", "未建立" if score["change"] is None else score["change"])


def page_merchants(data: Dict[str, Any]) -> None:
    st.subheader("商户权益")
    if not data.get("merchants"): st.info("暂无商户权益。")
    for merchant in data.get("merchants", []):
        result = evaluate_merchant_access(data, merchant)
        css = decision_class(result["result"])
        failed = "无" if not result["failed"] else "；".join(result["failed"])
        st.markdown(f"""<div class="decision {css}"><div class="pill">{merchant.get('category')}</div><h3>{merchant.get('name')}</h3><p><b>{result['message']}</b></p><p class="small">暂缓/失败原因：{failed}</p><p class="small">{merchant.get('note', '')}</p></div>""", unsafe_allow_html=True)
        with st.expander(f"查看 {merchant.get('name')} 风控条件"):
            st.dataframe(pd.DataFrame([{"条件": k, "是否通过": "通过" if v else "未通过"} for k, v in result["checks"].items()]), use_container_width=True, hide_index=True)
    st.divider()
    with st.form("merchant_form"):
        st.markdown("#### 新增商户")
        name = st.text_input("商户名称", "")
        category = st.selectbox("商户分类", spending_categories(data))
        c1, c2, c3 = st.columns(3)
        with c1: required = st.number_input("最低信用分", min_value=0, max_value=850, value=0)
        with c2: discount = st.number_input("折扣比例", min_value=0.0, max_value=0.9, value=0.08, step=0.01)
        with c3: cap = st.number_input("品类预算开放上限", min_value=0.0, max_value=2.0, value=0.90, step=0.05)
        note = st.text_input("说明", "")
        submitted = st.form_submit_button("新增商户", type="primary")
    if submitted and name.strip():
        data.setdefault("merchants", []).append({"id": uid(), "name": name.strip(), "category": category, "discount": discount, "required_score": int(required), "category_budget_cap": cap, "note": note})
        commit(data)
        st.success("商户已新增，展示页 merchants 已同步。")
        st.rerun()


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
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        if st.button("重新从存储读取"):
            reload_from_storage(); st.rerun()
    with c2:
        if st.button("手动刷新展示页"):
            refresh_report_sheets(data); st.success("展示页已刷新。")
    with c3:
        if st.button("删除最后一笔交易"):
            if data.get("transactions"):
                data["transactions"].pop(); commit(data); st.success("已删除最后一笔交易并同步。") ; st.rerun()
    with c4:
        confirm = st.checkbox("确认清空所有数据")
        if st.button("清空为空库", disabled=not confirm):
            reset_empty(); st.success("已清空并同步为空库。") ; st.rerun()


def render_sidebar(data: Dict[str, Any]) -> None:
    st.sidebar.title("系统设置")
    settings = data.setdefault("settings", {})
    with st.sidebar.form("settings_form"):
        owner = st.text_input("账户名称", settings.get("owner", "阿苏"))
        start_cash = st.number_input("初始现金", value=fnum(settings.get("start_cash")), step=1.0, format="%.2f")
        start_savings = st.number_input("初始储蓄", value=fnum(settings.get("start_savings")), step=1.0, format="%.2f")
        base_score = st.number_input("基础信用分", min_value=0, max_value=850, value=inum(settings.get("base_score"), 0))
        cash_floor = st.number_input("最低现金安全线", value=fnum(settings.get("cash_floor")), step=1.0, format="%.2f")
        loan_limit = st.number_input("贷款资产建议上限", value=fnum(settings.get("loan_asset_limit"), 0.45), min_value=0.0, max_value=1.0, step=0.05)
        hard_loan_limit = st.number_input("贷款资产硬红线", value=fnum(settings.get("hard_loan_asset_limit"), 0.65), min_value=0.0, max_value=1.0, step=0.05)
        submitted = st.form_submit_button("保存设置", type="primary")
    if submitted:
        settings["owner"] = owner
        settings["start_cash"] = start_cash
        settings["start_savings"] = start_savings
        settings["base_score"] = int(base_score)
        settings["cash_floor"] = cash_floor
        settings["loan_asset_limit"] = loan_limit
        settings["hard_loan_asset_limit"] = hard_loan_limit
        commit(data)
        st.sidebar.success("设置已保存并同步。")
        st.rerun()
    st.sidebar.divider()
    st.sidebar.caption("连接状态")
    if gsheet_enabled(): st.sidebar.success(f"Google Sheet：{st.secrets.get('SHEET_NAME', '')}")
    else: st.sidebar.warning("Google Sheet 未启用，使用本地 JSON")
    if get_secret_value("DEEPSEEK_API_KEY"): st.sidebar.success(f"DeepSeek：{get_secret_value('DEEPSEEK_MODEL', 'deepseek-v4-flash')}")
    else: st.sidebar.info("DeepSeek 未配置，使用本地报告")


# ============================================================
# 11. 主程序
# ============================================================

def main() -> None:
    st.set_page_config(page_title=APP_NAME, page_icon="🏦", layout="wide", initial_sidebar_state="expanded")
    inject_css()
    data = get_data()
    render_sidebar(data)
    render_hero(data)
    tabs = st.tabs(["1 首页总览", "2 新增交易", "3 AI 决策中心", "4 消费审批", "5 贷款审批", "6 风险雷达", "7 月度账单", "8 预算管理", "9 储蓄目标", "10 信用分", "11 商户权益", "12 交易流水"])
    with tabs[0]: page_home(data)
    with tabs[1]: page_add_transaction(data)
    with tabs[2]: page_ai_center(data)
    with tabs[3]: page_purchase(data)
    with tabs[4]: page_loan(data)
    with tabs[5]: page_risk_radar(data)
    with tabs[6]: page_monthly_statement(data)
    with tabs[7]: page_budget(data)
    with tabs[8]: page_goals(data)
    with tabs[9]: page_score(data)
    with tabs[10]: page_merchants(data)
    with tabs[11]: page_transactions(data)


if __name__ == "__main__":
    main()
