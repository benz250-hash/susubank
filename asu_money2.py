# -*- coding: utf-8 -*-
"""
阿苏私人银行 3.0 增强版

重点：
1. 阿苏首页像 Chase 儿童银行：资产、净资产、今日可安全花、信用分、任务、申请。
2. 爸爸妈妈权限相同：审批、赏金任务、设置、备份恢复。
3. 阿苏拥有记账权限：可以在“私人银行后台 → 新增交易”自己录入、修正、删除单笔交易。
4. Google Sheet 提速：日常只读写 state 主数据库；展示页手动刷新。
5. Python 负责账务和审批硬规则；AI 只负责分类和中文解释。

运行：
    streamlit run asu_money3.py

requirements.txt：
    streamlit
    pandas
    requests
    gspread
    google-auth
    plotly
    openpyxl
    numpy
    matplotlib
"""

from __future__ import annotations

import copy
import json
import os
import re
import uuid
from datetime import date, datetime, timedelta
from io import BytesIO
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

try:
    import plotly.express as px
    import plotly.graph_objects as go
except Exception:
    px = None
    go = None

try:
    import numpy as np
except Exception:
    np = None

try:
    import openpyxl  # noqa: F401
except Exception:
    openpyxl = None


APP_NAME = "阿苏私人银行 3.0 完整版"
SCHEMA_VERSION = "asu-bank-3-0-enhanced-bookkeeping"
DATA_PATH = Path("asu_money3_data.json")
OLD_DATA_PATH = Path("asu_money2_data.json")
SESSION_KEY = "asu_bank_3_0_state"
STATE_WS = "state"
STATE_KEY = "bank_data"
CHUNK_SIZE = 45000

PARENTS = {"爸爸", "妈妈"}
CHILD = "阿苏"
OPERATORS = ["爸爸", "妈妈", "阿苏"]

GSHEET_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

TX_TYPES = ["收入", "消费", "转入储蓄", "储蓄取出", "放贷", "还款", "借入", "偿还负债"]
DEFAULT_CATEGORIES = ["零花钱", "游戏", "甜品", "玩具", "学习", "宠物", "餐饮", "交通", "礼物", "赏金任务", "家庭贷款", "其他"]


# ============================================================
# 1. 基础工具
# ============================================================

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
        return int(float(x))
    except Exception:
        return default


def clamp(x: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, x))


def money(x: Any, currency: str = "USD") -> str:
    v = fnum(x)
    sign = "-" if v < 0 else ""
    v = abs(v)
    if currency == "USD":
        return f"{sign}${v:,.2f}"
    if currency == "CNY":
        return f"{sign}¥{v:,.2f}"
    return f"{sign}{currency} {v:,.2f}"


def percent(x: Any) -> str:
    return f"{fnum(x) * 100:.0f}%"


def tx_month(tx: Dict[str, Any]) -> str:
    return str(tx.get("date", ""))[:7]


def score_text(score: Optional[int]) -> str:
    return "未建立" if score is None else str(score)


def safe_date(s: Any) -> Optional[date]:
    try:
        return date.fromisoformat(str(s)[:10])
    except Exception:
        return None


def current_operator() -> str:
    return st.session_state.get("current_operator", "爸爸")


def can_parent() -> bool:
    return current_operator() in PARENTS


def can_bookkeep() -> bool:
    # 新规则：阿苏也有完整记账权限，可以新增、修正、删除单笔交易。
    return current_operator() in {"爸爸", "妈妈", "阿苏"}


def role_label() -> str:
    if can_parent():
        return "家长权限"
    return "阿苏记账权限"


def decision_class(result: str) -> str:
    if result in {"批准", "通过", "开放", "已批准", "已入账", "已完成", "已支付", "可购买"}:
        return "ok"
    if result in {"延迟", "限额通过", "暂缓开放", "观察", "待家长审批", "待审批", "未建立", "已提交", "已领取"}:
        return "warn"
    return "bad"


def pretty_value(key: str, value: Any, currency: str = "USD") -> str:
    if value is None:
        return "未建立"
    if isinstance(value, (int, float)):
        if any(word in key for word in ["率", "占比", "使用率", "上限", "比例"]):
            return percent(value)
        if "信用分" in key:
            return f"{value:.0f}"
        return money(value, currency)
    return str(value)


def secret_value(key: str, default: str = "") -> str:
    try:
        if key in st.secrets:
            return str(st.secrets[key])
    except Exception:
        pass
    return os.getenv(key, default)


# ============================================================
# 2. 数据模型
# ============================================================

def empty_data() -> Dict[str, Any]:
    return {
        "schema_version": SCHEMA_VERSION,
        "settings": {
            "owner": "阿苏",
            "currency": "USD",
            "start_cash": 0.0,
            "start_savings": 0.0,
            "base_score": 720,
            "cash_floor": 20.0,
            "approval_threshold": 15.0,
            "sheet_report_sync_on_save": False,
            "created_at": now_str(),
        },
        "rules": {
            "budget_warning_line": 0.90,
            "budget_reject_line": 1.20,
            "score_reject_line": 650,
            "loan_asset_soft_limit": 0.45,
            "loan_asset_hard_limit": 0.65,
            "spend_income_warning_line": 0.85,
        },
        "budgets": {"游戏": 20.0, "甜品": 15.0, "学习": 30.0},
        "goals": [],
        "merchants": [],
        "transactions": [],
        "pending_requests": [],
        "score_history": [],
        "weekly_reports": [],
        "backups": [],
        "badges": [],
        "rewards": [],
        "bounties": [],
        "audit_log": [],
    }


def normalize_tx(tx: Dict[str, Any]) -> Dict[str, Any]:
    return {
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
        "loan_id": tx.get("loan_id") or "",
        "bounty_id": tx.get("bounty_id") or "",
        "created_by": tx.get("created_by") or current_operator(),
        "created_at": tx.get("created_at") or now_str(),
        "updated_by": tx.get("updated_by") or "",
        "updated_at": tx.get("updated_at") or "",
    }


def normalize_data(raw: Any) -> Dict[str, Any]:
    if not isinstance(raw, dict):
        return empty_data()
    base = empty_data()

    if isinstance(raw.get("settings"), dict):
        base["settings"].update(raw["settings"])
    s = base["settings"]
    s["owner"] = s.get("owner") or "阿苏"
    s["currency"] = s.get("currency") or "USD"
    s["start_cash"] = fnum(s.get("start_cash"))
    s["start_savings"] = fnum(s.get("start_savings"))
    s["base_score"] = inum(s.get("base_score"), 720)
    s["cash_floor"] = fnum(s.get("cash_floor"), 20.0)
    s["approval_threshold"] = fnum(s.get("approval_threshold"), 15.0)
    s["sheet_report_sync_on_save"] = bool(s.get("sheet_report_sync_on_save", False))

    if isinstance(raw.get("rules"), dict):
        for k, v in raw["rules"].items():
            base["rules"][str(k)] = fnum(v, base["rules"].get(str(k), 0.0))

    budgets: Dict[str, float] = {}
    for k, v in (raw.get("budgets") or {}).items():
        name = str(k).strip()
        if name:
            budgets[name] = fnum(v)
    base["budgets"] = budgets

    goals = []
    for g in raw.get("goals") or []:
        if isinstance(g, dict):
            target = fnum(g.get("target"))
            current = fnum(g.get("current"))
            status = g.get("status") or ("已完成" if target > 0 and current >= target else "进行中")
            goals.append({
                "id": g.get("id") or uid(),
                "name": g.get("name") or "未命名目标",
                "target": target,
                "current": current,
                "deadline": g.get("deadline") or "",
                "category": g.get("category") or "长期储蓄",
                "note": g.get("note") or "",
                "status": status,
                "archived_at": g.get("archived_at") or "",
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

    base["transactions"] = [normalize_tx(tx) for tx in raw.get("transactions") or [] if isinstance(tx, dict)]

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

    bounties = []
    for b in raw.get("bounties") or []:
        if isinstance(b, dict):
            bounties.append({
                "id": b.get("id") or uid(),
                "title": b.get("title") or "未命名任务",
                "description": b.get("description") or "",
                "reward_amount": fnum(b.get("reward_amount")),
                "reward_points": inum(b.get("reward_points"), 0),
                "category": b.get("category") or "家庭任务",
                "difficulty": b.get("difficulty") or "普通",
                "deadline": b.get("deadline") or "",
                "status": b.get("status") or "开放",
                "created_by": b.get("created_by") or "",
                "created_at": b.get("created_at") or now_str(),
                "assigned_to": b.get("assigned_to") or "",
                "claimed_at": b.get("claimed_at") or "",
                "submitted_at": b.get("submitted_at") or "",
                "submission_note": b.get("submission_note") or "",
                "reviewed_by": b.get("reviewed_by") or "",
                "reviewed_at": b.get("reviewed_at") or "",
                "parent_note": b.get("parent_note") or "",
                "paid_tx_id": b.get("paid_tx_id") or "",
            })
    base["bounties"] = bounties

    for name in ["score_history", "weekly_reports", "backups", "badges", "rewards", "audit_log"]:
        if isinstance(raw.get(name), list):
            base[name] = raw[name]

    base["score_history"] = base.get("score_history", [])[-300:]
    base["backups"] = base.get("backups", [])[-30:]
    base["audit_log"] = base.get("audit_log", [])[-500:]
    base["schema_version"] = SCHEMA_VERSION
    return base


def strip_snapshot(data: Dict[str, Any]) -> Dict[str, Any]:
    snap = normalize_data(data)
    snap["backups"] = []
    return snap


def add_audit(data: Dict[str, Any], event: str, detail: str = "") -> None:
    data.setdefault("audit_log", []).append({"time": now_str(), "operator": current_operator(), "event": event, "detail": detail})
    data["audit_log"] = data["audit_log"][-500:]


# ============================================================
# 3. 登录和存储
# ============================================================

def configured_passwords() -> bool:
    try:
        return any(k in st.secrets for k in ["DAD_PASSWORD", "MOM_PASSWORD", "ASU_PASSWORD"])
    except Exception:
        return False


def require_login() -> None:
    if not configured_passwords():
        return
    if st.session_state.get("authenticated"):
        st.session_state["current_operator"] = st.session_state.get("authenticated_role", CHILD)
        return
    st.markdown("<div class='login-card'>", unsafe_allow_html=True)
    st.title("阿苏私人银行")
    st.caption("请选择身份并输入密码。阿苏有记账权限；爸爸妈妈有审批和设置权限。")
    role = st.selectbox("选择身份", OPERATORS)
    password = st.text_input("密码", type="password")
    key_map = {"爸爸": "DAD_PASSWORD", "妈妈": "MOM_PASSWORD", "阿苏": "ASU_PASSWORD"}
    expected = secret_value(key_map[role], "")
    if st.button("进入私人银行", type="primary", use_container_width=True):
        if expected and password == expected:
            st.session_state["authenticated"] = True
            st.session_state["authenticated_role"] = role
            st.session_state["current_operator"] = role
            st.rerun()
        else:
            st.error("密码错误，或该身份未配置密码。")
    st.markdown("</div>", unsafe_allow_html=True)
    st.stop()


def logout() -> None:
    for k in ["authenticated", "authenticated_role"]:
        st.session_state.pop(k, None)
    st.rerun()


def gsheet_enabled() -> bool:
    if gspread is None or Credentials is None:
        return False
    try:
        return "gcp_service_account" in st.secrets and "SHEET_NAME" in st.secrets
    except Exception:
        return False


@st.cache_resource(show_spinner=False)
def get_gsheet_client_cached():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=GSHEET_SCOPES)
    return gspread.authorize(creds)


@st.cache_resource(show_spinner=False)
def get_spreadsheet_cached(sheet_name: str):
    return get_gsheet_client_cached().open(sheet_name)


def get_spreadsheet():
    if not gsheet_enabled():
        raise RuntimeError("Google Sheet 未配置或依赖未安装。")
    return get_spreadsheet_cached(str(st.secrets["SHEET_NAME"]))


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


def encode_state_rows(data: Dict[str, Any]) -> List[List[str]]:
    text = json.dumps(normalize_data(data), ensure_ascii=False, separators=(",", ":"))
    chunks = [text[i:i + CHUNK_SIZE] for i in range(0, len(text), CHUNK_SIZE)] or [""]
    rows = [["key", "value"], ["schema_version", SCHEMA_VERSION], ["bank_data_chunks", str(len(chunks))]]
    rows += [[f"bank_data_{i:03d}", chunk] for i, chunk in enumerate(chunks)]
    return rows


def decode_state_rows(records: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    by_key = {str(r.get("key")): str(r.get("value") or "") for r in records}
    if STATE_KEY in by_key and by_key[STATE_KEY].strip():
        return normalize_data(json.loads(by_key[STATE_KEY]))
    if "bank_data_chunks" in by_key:
        n = inum(by_key.get("bank_data_chunks"), 0)
        text = "".join(by_key.get(f"bank_data_{i:03d}", "") for i in range(n))
        if text.strip():
            return normalize_data(json.loads(text))
    return None


def load_data_from_gsheet() -> Dict[str, Any]:
    sh = get_spreadsheet()
    ws = get_or_create_ws(sh, STATE_WS, rows=20, cols=2)
    records = ws.get_all_records()
    decoded = decode_state_rows(records)
    if decoded is None:
        data = empty_data()
        save_data_to_gsheet(data)
        return data
    return decoded


def save_data_to_gsheet(data: Dict[str, Any]) -> None:
    sh = get_spreadsheet()
    ws = get_or_create_ws(sh, STATE_WS, rows=20, cols=2)
    rows = encode_state_rows(data)
    ws.clear()
    ws.update(f"A1:B{len(rows)}", rows, value_input_option="RAW")


def save_data_local(data: Dict[str, Any]) -> None:
    DATA_PATH.write_text(json.dumps(normalize_data(data), ensure_ascii=False, indent=2), encoding="utf-8")


def load_data_local() -> Dict[str, Any]:
    source = DATA_PATH
    if not DATA_PATH.exists() and OLD_DATA_PATH.exists():
        source = OLD_DATA_PATH
    if not source.exists():
        data = empty_data()
        save_data_local(data)
        return data
    try:
        data = normalize_data(json.loads(source.read_text(encoding="utf-8")))
        if source == OLD_DATA_PATH:
            save_data_local(data)
        return data
    except Exception:
        data = empty_data()
        save_data_local(data)
        return data


def load_data() -> Dict[str, Any]:
    if gsheet_enabled():
        try:
            st.session_state["storage_backend"] = "Google Sheet state 快速存储"
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
            st.session_state["storage_backend"] = "Google Sheet state 快速存储"
            if data.get("settings", {}).get("sheet_report_sync_on_save"):
                sync_readable_sheets(data)
            return
        except Exception as exc:
            st.session_state["storage_backend"] = f"本地 JSON 兜底：Google Sheet 写入失败：{exc}"
    save_data_local(data)


def get_data() -> Dict[str, Any]:
    if SESSION_KEY not in st.session_state:
        st.session_state[SESSION_KEY] = load_data()
    return st.session_state[SESSION_KEY]


def reload_data() -> None:
    st.session_state.pop(SESSION_KEY, None)


def reset_empty() -> None:
    data = empty_data()
    st.session_state[SESSION_KEY] = data
    save_data(data)


# ============================================================
# 4. 账务计算
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
    loan_id: str = "",
    bounty_id: str = "",
) -> Dict[str, Any]:
    return normalize_tx({
        "id": uid(), "date": tx_date or today_str(), "type": tx_type, "amount": amount,
        "category": category, "party": party, "memo": memo,
        "expected_repayment": expected_repayment, "principal_repaid": principal_repaid,
        "interest_received": interest_received, "due_date": due_date,
        "loan_id": loan_id, "bounty_id": bounty_id,
        "created_by": current_operator(), "created_at": now_str(),
    })


def build_loan_book(transactions: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    loans: List[Dict[str, Any]] = []
    by_id: Dict[str, Dict[str, Any]] = {}
    for tx in sorted(transactions, key=lambda x: (str(x.get("date", "")), str(x.get("created_at", "")), str(x.get("id", "")))):
        typ = tx.get("type")
        if typ == "放贷":
            principal = fnum(tx.get("amount"))
            if principal <= 0:
                continue
            loan_id = tx.get("loan_id") or tx.get("id") or uid()
            loan = {
                "loan_id": loan_id,
                "tx_id": tx.get("id"),
                "date": tx.get("date"),
                "borrower": tx.get("party") or "未填写",
                "principal": principal,
                "repaid": 0.0,
                "remaining": principal,
                "expected_repayment": fnum(tx.get("expected_repayment")),
                "expected_interest": max(0.0, fnum(tx.get("expected_repayment")) - principal),
                "interest_received": 0.0,
                "due_date": tx.get("due_date") or "",
                "memo": tx.get("memo") or "",
                "status": "未还",
            }
            loans.append(loan)
            by_id[loan_id] = loan
            by_id[str(tx.get("id"))] = loan
        elif typ == "还款":
            amount = fnum(tx.get("principal_repaid")) or fnum(tx.get("amount"))
            interest = fnum(tx.get("interest_received"))
            target = by_id.get(str(tx.get("loan_id"))) if tx.get("loan_id") else None
            if target:
                candidates = [target]
            else:
                borrower = tx.get("party") or ""
                candidates = [l for l in loans if l["remaining"] > 0 and (not borrower or l["borrower"] == borrower)]
                if not candidates:
                    candidates = [l for l in loans if l["remaining"] > 0]
            for loan in candidates:
                if amount <= 0:
                    break
                applied = min(loan["remaining"], amount)
                loan["remaining"] -= applied
                loan["repaid"] += applied
                loan["interest_received"] += interest if applied > 0 else 0.0
                amount -= applied

    now = date.today()
    for loan in loans:
        if loan["remaining"] <= 0.0001:
            loan["remaining"] = 0.0
            loan["status"] = "已结清"
        elif loan.get("due_date"):
            d = safe_date(loan["due_date"])
            if d and d < now:
                loan["status"] = "逾期"
            elif loan["repaid"] > 0:
                loan["status"] = "部分还款"
        elif loan["repaid"] > 0:
            loan["status"] = "部分还款"
    return loans


def calc_budget_usage(data: Dict[str, Any], month: Optional[str] = None) -> Dict[str, Dict[str, float]]:
    month = month or month_str()
    usage = {cat: {"limit": fnum(limit), "spent": 0.0, "ratio": 0.0, "remaining": fnum(limit)} for cat, limit in data.get("budgets", {}).items()}
    for tx in data.get("transactions", []):
        if tx.get("type") == "消费" and tx_month(tx) == month:
            cat = tx.get("category") or "其他"
            usage.setdefault(cat, {"limit": 0.0, "spent": 0.0, "ratio": 0.0, "remaining": 0.0})
            usage[cat]["spent"] += fnum(tx.get("amount"))
    for row in usage.values():
        row["ratio"] = row["spent"] / row["limit"] if row["limit"] > 0 else 0.0
        row["remaining"] = max(0.0, row["limit"] - row["spent"]) if row["limit"] > 0 else 0.0
    return usage


def calc_goals(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    out = []
    today = date.today()
    for g in data.get("goals", []):
        target = fnum(g.get("target"))
        current = fnum(g.get("current"))
        progress = current / target if target > 0 else 0.0
        days_left = None
        if g.get("deadline"):
            d = safe_date(g["deadline"])
            if d:
                days_left = (d - today).days
        status = g.get("status") or "进行中"
        if status == "进行中" and target > 0 and current >= target:
            status = "已完成"
        out.append({**g, "status": status, "progress": progress, "remaining": max(0.0, target - current), "days_left": days_left})
    return out


def calc_credit_score(data: Dict[str, Any], cash: float, savings: float, receivables: float, liabilities: float, total_assets: float, month_income: float, month_expense: float, budget_usage: Dict[str, Dict[str, float]], overdue: float) -> Tuple[Optional[int], List[str]]:
    s = data.get("settings", {})
    base = inum(s.get("base_score"), 0)
    if base <= 0:
        return None, ["账户尚未建立信用分。请在设置里填写基础信用分。"]
    score = float(base)
    factors: List[str] = []
    floor = fnum(s.get("cash_floor"))

    if cash < 0:
        score -= 55; factors.append("现金余额为负：-55")
    elif floor > 0 and cash < floor:
        score -= 18; factors.append(f"现金低于安全线 {money(floor)}：-18")
    elif floor > 0 and cash >= floor * 2:
        score += 6; factors.append("现金缓冲较充足：+6")

    if total_assets > 0:
        savings_ratio = savings / total_assets
        if savings_ratio >= 0.35:
            score += 10; factors.append("储蓄占总资产 35% 以上：+10")
        elif savings_ratio < 0.10:
            score -= 8; factors.append("储蓄占总资产低于 10%：-8")

    if month_income <= 0 and month_expense > 0:
        score -= 12; factors.append("本月有消费但没有收入记录：-12")
    elif month_income > 0:
        ratio = month_expense / month_income
        if ratio > 1.0:
            score -= 35; factors.append("本月支出超过收入：-35")
        elif ratio > 0.85:
            score -= 18; factors.append("本月支出达到收入 85% 以上：-18")
        elif ratio > 0.65:
            score -= 8; factors.append("本月支出达到收入 65% 以上：-8")
        elif ratio <= 0.45:
            score += 5; factors.append("本月支出控制在收入 45% 以下：+5")

    loan_ratio = receivables / total_assets if total_assets > 0 else 0.0
    if loan_ratio > 0.65:
        score -= 30; factors.append("应收贷款本金占总资产超过 65%：-30")
    elif loan_ratio > 0.50:
        score -= 20; factors.append("应收贷款本金占总资产超过 50%：-20")
    elif loan_ratio > 0.35:
        score -= 9; factors.append("应收贷款本金占总资产超过 35%：-9")

    debt_ratio = liabilities / total_assets if total_assets > 0 else 0.0
    if debt_ratio > 0.50:
        score -= 45; factors.append("负债占总资产超过 50%：-45")
    elif debt_ratio > 0.25:
        score -= 25; factors.append("负债占总资产超过 25%：-25")
    elif liabilities > 0:
        score -= 7; factors.append("存在未偿还负债：-7")

    overspent = sum(1 for r in budget_usage.values() if r["limit"] > 0 and r["ratio"] > 1)
    near = sum(1 for r in budget_usage.values() if r["limit"] > 0 and 0.85 < r["ratio"] <= 1)
    if overspent:
        p = min(32, overspent * 10); score -= p; factors.append(f"{overspent} 个预算品类已超支：-{p}")
    if near:
        p = min(15, near * 4); score -= p; factors.append(f"{near} 个预算品类接近上限：-{p}")
    if overdue > 0:
        score -= 24; factors.append(f"存在逾期未收回应收本金 {money(overdue)}：-24")

    completed_bounties = sum(1 for b in data.get("bounties", []) if b.get("status") in {"已完成", "已支付"})
    if completed_bounties >= 3:
        bonus = min(12, completed_bounties * 2)
        score += bonus; factors.append(f"完成赏金任务 {completed_bounties} 个：+{bonus}")
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
    month_income = month_expense = total_income = total_expense = 0.0

    for tx in sorted(data.get("transactions", []), key=lambda x: (str(x.get("date", "")), str(x.get("created_at", "")), str(x.get("id", "")))):
        typ = tx.get("type")
        amount = fnum(tx.get("amount"))
        is_month = tx_month(tx) == month
        if typ == "收入":
            cash += amount; total_income += amount
            if is_month: month_income += amount
        elif typ == "消费":
            cash -= amount; total_expense += amount
            if is_month: month_expense += amount
        elif typ == "转入储蓄":
            cash -= amount; savings += amount
        elif typ == "储蓄取出":
            cash += amount; savings -= amount
        elif typ == "放贷":
            cash -= amount
        elif typ == "还款":
            principal = fnum(tx.get("principal_repaid")) or amount
            interest = fnum(tx.get("interest_received"))
            cash += principal + interest; total_income += interest
            if is_month: month_income += interest
        elif typ == "借入":
            cash += amount; liabilities += amount
        elif typ == "偿还负债":
            cash -= amount; liabilities -= amount

    liabilities = max(0.0, liabilities)
    loans = build_loan_book(data.get("transactions", []))
    receivables = sum(fnum(l.get("remaining")) for l in loans)
    overdue = sum(fnum(l.get("remaining")) for l in loans if l.get("status") == "逾期")
    total_assets = cash + savings + receivables
    net_assets = total_assets - liabilities
    usage = calc_budget_usage(data, month)
    goals = calc_goals(data)
    score, factors = calc_credit_score(data, cash, savings, receivables, liabilities, total_assets, month_income, month_expense, usage, overdue)
    spend_income_ratio = month_expense / month_income if month_income > 0 else (1.0 if month_expense > 0 else 0.0)
    loan_asset_ratio = receivables / total_assets if total_assets > 0 else 0.0
    debt_asset_ratio = liabilities / total_assets if total_assets > 0 else 0.0
    savings_asset_ratio = savings / total_assets if total_assets > 0 else 0.0
    safe_spend_cash = max(0.0, cash - fnum(s.get("cash_floor")))
    budget_remaining = sum(r["remaining"] for r in usage.values() if r["limit"] > 0)
    safe_to_spend = min(safe_spend_cash, budget_remaining) if budget_remaining > 0 else safe_spend_cash

    return {
        "currency": ccy, "cash": cash, "savings": savings, "receivables": receivables,
        "liabilities": liabilities, "total_assets": total_assets, "net_assets": net_assets,
        "month_income": month_income, "month_expense": month_expense,
        "total_income": total_income, "total_expense": total_expense,
        "spend_income_ratio": spend_income_ratio, "loan_asset_ratio": loan_asset_ratio,
        "debt_asset_ratio": debt_asset_ratio, "savings_asset_ratio": savings_asset_ratio,
        "safe_to_spend": safe_to_spend, "safe_spend_cash": safe_spend_cash,
        "budget_remaining": budget_remaining, "budget_usage": usage, "goals": goals,
        "loans": loans, "overdue_principal": overdue,
        "credit_score": score, "score_factors": factors,
    }


def has_activity(data: Dict[str, Any]) -> bool:
    s = data.get("settings", {})
    return bool(fnum(s.get("start_cash")) or fnum(s.get("start_savings")) or data.get("transactions") or data.get("budgets") or data.get("goals") or data.get("pending_requests") or data.get("bounties"))


def category_options(data: Dict[str, Any]) -> List[str]:
    return sorted(set(DEFAULT_CATEGORIES + list(data.get("budgets", {}).keys())))


def is_goal_active(g: Dict[str, Any]) -> bool:
    return (g.get("status") or "进行中") not in {"已完成", "已归档", "已删除"}


def active_goals_from_metrics(metrics: Dict[str, Any]) -> List[Dict[str, Any]]:
    return [g for g in metrics.get("goals", []) if is_goal_active(g)]


def archived_goals_from_metrics(metrics: Dict[str, Any]) -> List[Dict[str, Any]]:
    return [g for g in metrics.get("goals", []) if not is_goal_active(g)]


def is_bounty_active(b: Dict[str, Any]) -> bool:
    return (b.get("status") or "开放") not in {"已支付", "已取消", "已归档", "已删除"}


def active_bounties(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    return [b for b in data.get("bounties", []) if is_bounty_active(b)]


def historical_bounties(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    return [b for b in data.get("bounties", []) if not is_bounty_active(b)]


def add_badge(data: Dict[str, Any], name: str, description: str) -> None:
    existing = {b.get("name") for b in data.get("badges", []) if isinstance(b, dict)}
    if name not in existing:
        data.setdefault("badges", []).append({"id": uid(), "name": name, "earned_at": now_str(), "description": description})


def evaluate_badges(data: Dict[str, Any]) -> None:
    m = calc_financials(data)
    if any(tx.get("type") == "收入" for tx in data.get("transactions", [])):
        add_badge(data, "第一笔收入", "已经建立现金流记录。")
    if len(data.get("budgets", {})) >= 3:
        add_badge(data, "预算建筑师", "已经建立至少三个预算分类。")
    if m.get("credit_score") is not None:
        add_badge(data, "信用分已建立", "风控系统已经可以追踪信用分。")
    if data.get("pending_requests"):
        add_badge(data, "合规申请人", "已经使用待审批流程。")
    if any(tx.get("created_by") == CHILD for tx in data.get("transactions", [])):
        add_badge(data, "自主记账员", "阿苏已经开始自己维护交易流水。")
    if any(b.get("status") in {"已完成", "已支付"} for b in data.get("bounties", [])):
        add_badge(data, "赏金猎人", "已经完成至少一个家庭赏金任务。")


def commit(data: Dict[str, Any], event: str = "保存数据", reason: str = "") -> None:
    old_state = normalize_data(st.session_state.get(SESSION_KEY, data))
    old_score = calc_financials(old_state).get("credit_score")
    new_data = normalize_data(data)
    backup = {"id": uid(), "time": now_str(), "operator": current_operator(), "event": event, "reason": reason or "", "snapshot_json": json.dumps(strip_snapshot(old_state), ensure_ascii=False)}
    new_data["backups"] = (old_state.get("backups", []) + [backup])[-30:]
    add_audit(new_data, event, reason)
    new_score = calc_financials(new_data).get("credit_score")
    if new_score is not None:
        change = None if old_score is None else new_score - old_score
        new_data.setdefault("score_history", []).append({"time": now_str(), "score": new_score, "change": change, "event": event, "reason": reason or event})
        new_data["score_history"] = new_data["score_history"][-300:]
    evaluate_badges(new_data)
    st.session_state[SESSION_KEY] = new_data
    save_data(new_data)


# ============================================================
# 5. 审批、AI、计划
# ============================================================

def rule_value(data: Dict[str, Any], key: str, default: float) -> float:
    return fnum((data.get("rules") or {}).get(key), default)


def simulate_transaction(data: Dict[str, Any], tx: Dict[str, Any]) -> Dict[str, Any]:
    before = calc_financials(data)
    after_data = copy.deepcopy(data)
    after_data.setdefault("transactions", []).append(tx)
    after = calc_financials(after_data)
    before_score = before["credit_score"]
    after_score = after["credit_score"]
    delta_score = None if before_score is None or after_score is None else after_score - before_score
    return {"before": before, "after": after, "delta": {"cash": after["cash"] - before["cash"], "receivables": after["receivables"] - before["receivables"], "total_assets": after["total_assets"] - before["total_assets"], "net_assets": after["net_assets"] - before["net_assets"], "credit_score": delta_score}, "tx": tx}


def usage_for(metrics: Dict[str, Any], cat: str) -> Dict[str, float]:
    return metrics.get("budget_usage", {}).get(cat, {"limit": 0.0, "spent": 0.0, "ratio": 0.0, "remaining": 0.0})


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


def local_category(text: str) -> Dict[str, str]:
    s = (text or "").lower()
    rules = {
        "游戏": ["minecraft", "robux", "游戏", "道具", "皮肤", "steam", "switch", "roblox"],
        "甜品": ["甜品", "奶茶", "冰淇淋", "蛋糕", "糖", "饮料", "boba"],
        "玩具": ["玩具", "lego", "乐高", "手办", "娃娃", "卡牌"],
        "学习": ["书", "学习", "课程", "文具", "作业", "训练", "练习册"],
        "宠物": ["猫", "猫粮", "猫砂", "宠物", "罐头"],
        "餐饮": ["饭", "餐", "披萨", "pizza", "汉堡", "麦当劳", "lunch"],
        "交通": ["车", "uber", "公交", "地铁", "停车"],
    }
    category = "其他"
    for cat, keys in rules.items():
        if any(k in s for k in keys):
            category = cat; break
    necessity = "半必要" if category in {"学习", "宠物", "餐饮", "交通"} else "非必要"
    impulse = "高" if category in {"游戏", "甜品", "玩具"} else "中"
    return {"category": category, "necessity": necessity, "impulse_level": impulse, "merchant": "自然语言输入", "note": "本地关键词分类"}


def ai_classify_purchase(text: str) -> Dict[str, str]:
    api_key = secret_value("DEEPSEEK_API_KEY")
    model = secret_value("DEEPSEEK_MODEL", "deepseek-chat")
    if not api_key or requests is None:
        return local_category(text)
    prompt = f"""
请把下面的儿童消费/贷款请求分类，必须只输出 JSON，不要解释。
可用分类：游戏、甜品、玩具、学习、宠物、餐饮、交通、礼物、家庭贷款、其他
字段：category, necessity, impulse_level, merchant, note
用户输入：{text}
""".strip()
    try:
        resp = requests.post(
            "https://api.deepseek.com/chat/completions",
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            json={"model": model, "messages": [{"role": "user", "content": prompt}], "temperature": 0, "max_tokens": 200},
            timeout=12,
        )
        resp.raise_for_status()
        content = resp.json()["choices"][0]["message"]["content"]
        m = re.search(r"\{.*\}", content, re.S)
        if not m:
            return local_category(text)
        out = json.loads(m.group(0))
        return {"category": str(out.get("category") or "其他"), "necessity": str(out.get("necessity") or "未知"), "impulse_level": str(out.get("impulse_level") or "未知"), "merchant": str(out.get("merchant") or "自然语言输入"), "note": str(out.get("note") or "")}
    except Exception:
        return local_category(text)


def evaluate_purchase_decision(data: Dict[str, Any], amount: float, category: str, desc: str, merchant: str = "") -> Dict[str, Any]:
    tx = make_tx("消费", amount, category, party=merchant, memo=desc)
    sim = simulate_transaction(data, tx)
    before, after, delta = sim["before"], sim["after"], sim["delta"]
    budget = usage_for(after, category)
    floor = fnum(data.get("settings", {}).get("cash_floor"))
    threshold = fnum(data.get("settings", {}).get("approval_threshold"), 15.0)
    result = "批准"
    reasons: List[str] = []
    if amount <= 0:
        result = "拒绝"; reasons.append("购买金额必须大于 0。")
    if after["cash"] < 0:
        result = "拒绝"; reasons.append("购买后现金余额为负，触发硬性拒绝。")
    if budget["limit"] > 0 and budget["ratio"] > rule_value(data, "budget_reject_line", 1.20):
        result = "拒绝"; reasons.append(f"{category}预算使用率将达到 {percent(budget['ratio'])}，超过预算红线。")
    if after["credit_score"] is not None and after["credit_score"] < rule_value(data, "score_reject_line", 650):
        result = "拒绝"; reasons.append("交易后信用分低于风控线。")
    if result != "拒绝":
        flags = []
        if amount >= threshold: flags.append(f"金额达到家长审批线 {money(threshold)}。")
        if floor > 0 and after["cash"] < floor: flags.append(f"购买后现金余额低于安全线 {money(floor)}。")
        if budget["limit"] > 0 and budget["ratio"] >= rule_value(data, "budget_warning_line", 0.90): flags.append(f"{category}预算使用率将达到 {percent(budget['ratio'])}。")
        if delta["credit_score"] is not None and delta["credit_score"] <= -8: flags.append(f"信用分预计下降 {abs(delta['credit_score'])} 分。")
        if after["month_income"] > 0 and after["spend_income_ratio"] >= rule_value(data, "spend_income_warning_line", 0.85): flags.append(f"本月支出/收入比将达到 {percent(after['spend_income_ratio'])}。")
        if flags:
            result = "待家长审批" if amount >= threshold else "延迟"
            reasons.extend(flags)
    if result == "批准":
        reasons.append("现金余额、预算使用率、信用分变化均未触发硬性限制。")
    if result == "拒绝": action = "不要购买。先补现金、建预算，或等收入到账后再申请。"
    elif result == "延迟": action = "建议延迟到下一笔收入到账后再买，或者降低金额。"
    elif result == "待家长审批": action = "进入待审批队列，由爸爸或妈妈确认是否入账。"
    else: action = "可购买；家长仍可复核。"
    return {"kind": "消费审批", "result": result, "reasons": reasons, "action": action, "metrics": {"购买金额": amount, "消费分类": category, "商户": merchant or "未填写", "当前现金余额": before["cash"], "交易后现金余额": after["cash"], "当前信用分": before["credit_score"], "交易后信用分": after["credit_score"], "信用分变化": delta["credit_score"], "品类预算上限": budget["limit"], "交易后品类已花": budget["spent"], "交易后品类预算使用率": budget["ratio"], "交易后本月支出收入比": after["spend_income_ratio"]}, "pending_tx": tx, "simulation": sim}


def evaluate_loan_decision(data: Dict[str, Any], principal: float, borrower: str, expected: float, due: str, memo: str = "") -> Dict[str, Any]:
    loan_id = uid()
    tx = make_tx("放贷", principal, "家庭贷款", party=borrower, memo=memo, expected_repayment=expected, due_date=due, loan_id=loan_id)
    sim = simulate_transaction(data, tx)
    before, after, delta = sim["before"], sim["after"], sim["delta"]
    floor = fnum(data.get("settings", {}).get("cash_floor"))
    soft = rule_value(data, "loan_asset_soft_limit", 0.45)
    hard = rule_value(data, "loan_asset_hard_limit", 0.65)
    max_by_cash = max(0.0, before["cash"] - floor) if floor > 0 else max(0.0, before["cash"])
    safe_capacity = max(0.0, soft * max(before["total_assets"], 1.0) - before["receivables"])
    suggested = max(0.0, min(principal, max_by_cash, safe_capacity))
    result = "通过"; reasons: List[str] = []
    if principal <= 0: result = "拒绝"; reasons.append("放贷本金必须大于 0。")
    if after["cash"] < 0: result = "拒绝"; reasons.append("放贷后现金余额为负，流动性不足。")
    if after["total_assets"] > 0 and after["loan_asset_ratio"] > hard: result = "拒绝"; reasons.append(f"放贷后贷款资产占比 {percent(after['loan_asset_ratio'])}，超过红线 {percent(hard)}。")
    if after["credit_score"] is not None and after["credit_score"] < rule_value(data, "score_reject_line", 650): result = "拒绝"; reasons.append("放贷后信用分低于风控线。")
    if result != "拒绝":
        flags = []
        if floor > 0 and after["cash"] < floor: flags.append(f"放贷后现金低于安全线 {money(floor)}。")
        if after["total_assets"] > 0 and after["loan_asset_ratio"] > soft: flags.append(f"放贷后贷款资产占比 {percent(after['loan_asset_ratio'])}，超过建议线 {percent(soft)}。")
        if principal > suggested and suggested > 0: flags.append(f"建议最高放贷金额为 {money(suggested)}。")
        if not due: flags.append("没有填写预计还款日，回款纪律不足。")
        if flags: result = "限额通过"; reasons.extend(flags)
    if result == "通过": reasons.append("放贷后现金缓冲、贷款资产占比、信用分均未触发硬性限制。")
    action = "建议进入家长审批队列。"
    if result == "拒绝": action = "本次不建议放贷。先增加现金余额，或降低放贷金额。"
    elif result == "限额通过": action = f"建议限额放贷，最高不超过 {money(suggested)}，并写清还款日。"
    return {"kind": "贷款审批", "result": result, "reasons": reasons, "action": action, "metrics": {"拟借出本金": principal, "借款人": borrower or "未填写", "预计回款": expected, "当前现金余额": before["cash"], "放贷后现金余额": after["cash"], "当前应收贷款本金": before["receivables"], "放贷后应收贷款本金": after["receivables"], "放贷后贷款资产占比": after["loan_asset_ratio"], "当前信用分": before["credit_score"], "放贷后信用分": after["credit_score"], "信用分变化": delta["credit_score"], "建议最高放贷金额": suggested}, "pending_tx": tx, "simulation": sim}


def parse_natural_request(data: Dict[str, Any], text: str) -> Dict[str, Any]:
    amounts = extract_amounts(text)
    if any(x in text for x in ["借给", "放贷", "贷款给"]):
        principal = amounts[0] if amounts else 0.0
        expected = amounts[1] if len(amounts) >= 2 else principal
        return evaluate_loan_decision(data, principal, infer_borrower(text), expected, "", text)
    amount = amounts[0] if amounts else 0.0
    cls = ai_classify_purchase(text)
    decision = evaluate_purchase_decision(data, amount, cls["category"], text, cls["merchant"])
    decision["ai"] = cls
    return decision


def scenario_planning(data: Dict[str, Any], amount: float, category: str, desc: str, merchant: str = "") -> List[Dict[str, Any]]:
    scenarios = []
    today_decision = evaluate_purchase_decision(data, amount, category, desc, merchant)
    scenarios.append({"方案": "今天购买", "结论": today_decision["result"], "交易后现金": today_decision["metrics"]["交易后现金余额"], "交易后信用分": today_decision["metrics"]["交易后信用分"], "信用分变化": today_decision["metrics"]["信用分变化"], "说明": today_decision["action"]})
    next_data = copy.deepcopy(data)
    next_data.setdefault("transactions", []).append(make_tx("收入", amount, "零花钱", memo="情景模拟：下周收入到账"))
    next_decision = evaluate_purchase_decision(next_data, amount, category, desc + "（下周购买）", merchant)
    scenarios.append({"方案": "下周收入后购买", "结论": next_decision["result"], "交易后现金": next_decision["metrics"]["交易后现金余额"], "交易后信用分": next_decision["metrics"]["交易后信用分"], "信用分变化": next_decision["metrics"]["信用分变化"], "说明": "先等同等金额收入到账，再购买，现金压力下降。"})
    half = round(amount / 2, 2)
    split_decision = evaluate_purchase_decision(data, half, category, desc + "（分两周第一笔）", merchant)
    scenarios.append({"方案": "分两周购买", "结论": split_decision["result"], "交易后现金": split_decision["metrics"]["交易后现金余额"], "交易后信用分": split_decision["metrics"]["交易后信用分"], "信用分变化": split_decision["metrics"]["信用分变化"], "说明": f"先支出 {money(half)}，剩余下周再处理，预算冲击更小。"})
    return scenarios


def submit_pending_request(data: Dict[str, Any], decision: Dict[str, Any], request_text: str, applicant: str = CHILD) -> None:
    cls = decision.get("ai") or ai_classify_purchase(request_text)
    kind = decision.get("kind", "消费审批")
    amount = fnum(decision.get("metrics", {}).get("购买金额") or decision.get("metrics", {}).get("拟借出本金"))
    category = str(decision.get("metrics", {}).get("消费分类") or ("家庭贷款" if kind == "贷款审批" else cls.get("category", "其他")))
    data.setdefault("pending_requests", []).append({"id": uid(), "created_at": now_str(), "request_text": request_text, "request_type": kind.replace("审批", ""), "amount": amount, "category": category, "merchant": str(decision.get("metrics", {}).get("商户") or ""), "applicant": applicant, "ai_category": cls.get("category", "其他"), "necessity": cls.get("necessity", "未知"), "impulse_level": cls.get("impulse_level", "未知"), "python_decision": decision.get("result", "观察"), "parent_status": "待审批", "final_status": "未入账", "decision_json": {k: v for k, v in decision.items() if k != "simulation"}, "pending_tx": decision.get("pending_tx"), "parent_note": "", "approved_by": "", "approved_at": "", "booked_at": ""})


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
    model = secret_value("DEEPSEEK_MODEL", "deepseek-chat")
    if not api_key or requests is None:
        return local_report(decision)
    payload = {"kind": decision.get("kind"), "result": decision.get("result"), "metrics": decision.get("metrics"), "reasons": decision.get("reasons"), "action": decision.get("action")}
    system_prompt = "你是阿苏私人银行的中文儿童金融教练。只能解释Python已计算的硬指标，禁止编造数字，禁止改变审批结论。"
    try:
        resp = requests.post("https://api.deepseek.com/chat/completions", headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}, json={"model": model, "messages": [{"role": "system", "content": system_prompt}, {"role": "user", "content": json.dumps(payload, ensure_ascii=False)}], "temperature": 0.2, "max_tokens": 800}, timeout=15)
        resp.raise_for_status()
        return resp.json()["choices"][0]["message"]["content"]
    except Exception:
        return local_report(decision)


def build_risk_radar(data: Dict[str, Any]) -> Dict[str, List[Dict[str, str]]]:
    m = calc_financials(data); ccy = m["currency"]
    risks = {"高风险": [], "中风险": [], "低风险": []}
    if not has_activity(data):
        risks["低风险"].append({"title": "账户为空", "detail": "先建立第一笔收入、预算或目标。"}); return risks
    floor = fnum(data.get("settings", {}).get("cash_floor"))
    if m["cash"] < 0: risks["高风险"].append({"title": "现金余额为负", "detail": f"当前现金 {money(m['cash'], ccy)}，暂停非必要消费。"})
    elif floor > 0 and m["cash"] < floor: risks["中风险"].append({"title": "现金缓冲偏低", "detail": f"当前现金 {money(m['cash'], ccy)}，低于安全线 {money(floor, ccy)}。"})
    else: risks["低风险"].append({"title": "现金余额安全", "detail": f"当前现金 {money(m['cash'], ccy)}。"})
    if m["month_income"] > 0:
        if m["spend_income_ratio"] > 1: risks["高风险"].append({"title": "本月支出超过收入", "detail": f"支出/收入比 {percent(m['spend_income_ratio'])}。"})
        elif m["spend_income_ratio"] > 0.85: risks["中风险"].append({"title": "支出收入比偏高", "detail": f"支出/收入比 {percent(m['spend_income_ratio'])}。"})
        else: risks["低风险"].append({"title": "支出收入比可控", "detail": f"支出/收入比 {percent(m['spend_income_ratio'])}。"})
    if m["loan_asset_ratio"] > 0.60: risks["高风险"].append({"title": "贷款资产占比过高", "detail": f"应收贷款占总资产 {percent(m['loan_asset_ratio'])}。"})
    elif m["loan_asset_ratio"] > 0.45: risks["中风险"].append({"title": "贷款资产占比偏高", "detail": f"应收贷款占总资产 {percent(m['loan_asset_ratio'])}。"})
    else: risks["低风险"].append({"title": "贷款占比可控", "detail": f"应收贷款占总资产 {percent(m['loan_asset_ratio'])}。"})
    if m["liabilities"] > 0: risks["中风险"].append({"title": "存在未偿还负债", "detail": f"负债余额 {money(m['liabilities'], ccy)}。"})
    else: risks["低风险"].append({"title": "无负债", "detail": "净资产没有被负债侵蚀。"})
    for cat, row in m["budget_usage"].items():
        if row["limit"] > 0 and row["ratio"] > 1: risks["高风险"].append({"title": f"{cat}预算超支", "detail": f"使用率 {percent(row['ratio'])}。"})
        elif row["limit"] > 0 and row["ratio"] > 0.85: risks["中风险"].append({"title": f"{cat}预算接近上限", "detail": f"使用率 {percent(row['ratio'])}。"})
    if m["overdue_principal"] > 0: risks["高风险"].append({"title": "存在逾期贷款", "detail": f"逾期本金 {money(m['overdue_principal'], ccy)}。"})
    return risks


def build_weekly_plan(data: Dict[str, Any]) -> List[str]:
    if not has_activity(data):
        return ["先建立第一笔收入或现金余额。", "新增至少 3 个预算分类。", "设置基础信用分和现金安全线。"]
    m = calc_financials(data)
    actions: List[str] = []
    overspent = [cat for cat, r in m["budget_usage"].items() if r["limit"] > 0 and r["ratio"] > 1]
    near = [cat for cat, r in m["budget_usage"].items() if r["limit"] > 0 and 0.85 < r["ratio"] <= 1]
    if overspent: actions.append(f"本周暂停 {', '.join(overspent)} 类非必要消费。")
    elif near: actions.append(f"本周 {', '.join(near)} 类消费必须先审批。")
    if m["loan_asset_ratio"] > 0.45: actions.append("优先收回应收贷款本金，暂停新增放贷。")
    floor = fnum(data.get("settings", {}).get("cash_floor"))
    if floor > 0 and m["cash"] < floor: actions.append("下一笔收入先补现金安全线。")
    if m["savings_asset_ratio"] < 0.25 and m["month_income"] > 0: actions.append("下一笔收入至少 30% 转入储蓄。")
    if m["liabilities"] > 0: actions.append("优先偿还负债，不新增借入资金。")
    if [b for b in data.get("bounties", []) if b.get("status") == "开放"]: actions.append("选择一个赏金任务，用劳动换收入。")
    return actions[:5] or ["维持当前消费节奏。", "下一笔收入至少 20% 转入储蓄。"]


def evaluate_merchant_access(data: Dict[str, Any], merchant: Dict[str, Any]) -> Dict[str, Any]:
    m = calc_financials(data); cat = merchant.get("category") or "其他"; usage = usage_for(m, cat)
    required = inum(merchant.get("required_score"), 0); cap = fnum(merchant.get("category_budget_cap"), 1.0); score = m["credit_score"]
    checks = {"信用分达标": score is not None and score >= required, "现金余额为正": m["cash"] > 0, "本月支出收入比不高于90%": m["spend_income_ratio"] <= 0.90, "该品类预算未过高": usage["ratio"] <= cap if usage["limit"] > 0 else True}
    opened = all(checks.values()); failed = [k for k, ok in checks.items() if not ok]
    if opened: result = "开放"; message = f"折扣开放：{fnum(merchant.get('discount')) * 100:.0f}%"
    elif checks["信用分达标"]: result = "暂缓开放"; message = "信用分达标，但现金或预算条件未通过。"
    else: result = "关闭"; message = "信用分未达标或尚未建立。"
    return {"result": result, "message": message, "failed": failed, "checks": checks}


def generate_weekly_report(data: Dict[str, Any]) -> Dict[str, Any]:
    start = week_start(); end = start + timedelta(days=7)
    rows = []
    for tx in data.get("transactions", []):
        d = safe_date(tx.get("date"))
        if d and start <= d < end: rows.append(tx)
    income = sum(fnum(tx.get("amount")) for tx in rows if tx.get("type") == "收入")
    expense = sum(fnum(tx.get("amount")) for tx in rows if tx.get("type") == "消费")
    savings_move = sum(fnum(tx.get("amount")) for tx in rows if tx.get("type") == "转入储蓄")
    completed_bounties = len([b for b in data.get("bounties", []) if b.get("status") in {"已完成", "已支付"} and b.get("reviewed_at", "")[:10] >= start.isoformat()])
    m = calc_financials(data)
    spend_by_cat: Dict[str, float] = {}
    for tx in rows:
        if tx.get("type") == "消费": spend_by_cat[tx.get("category") or "其他"] = spend_by_cat.get(tx.get("category") or "其他", 0.0) + fnum(tx.get("amount"))
    top_cat = max(spend_by_cat, key=spend_by_cat.get) if spend_by_cat else "无"
    actions = build_weekly_plan(data)
    content = "\n".join([f"阿苏私人银行周报：{start.isoformat()} 至 {(end - timedelta(days=1)).isoformat()}", f"本周收入：{money(income)}", f"本周消费：{money(expense)}", f"本周转入储蓄：{money(savings_move)}", f"本周完成赏金任务：{completed_bounties} 个", f"最大消费分类：{top_cat}", f"当前信用分：{score_text(m['credit_score'])}", "下周行动：", *[f"{i + 1}. {a}" for i, a in enumerate(actions)]])
    return {"week_start": start.isoformat(), "generated_at": now_str(), "content": content, "metrics": {"income": income, "expense": expense, "savings": savings_move, "score": score_text(m["credit_score"]), "top_category": top_cat, "completed_bounties": completed_bounties}}


# ============================================================
# 6. Google Sheet 展示页和 Excel
# ============================================================

def _sheet_a1_range(title: str, rows: int, cols: int) -> str:
    safe_title = str(title).replace("'", "''")
    return f"'{safe_title}'!A1:{col_letters(cols)}{rows}"


def _pad_sheet_values(values: List[List[Any]], rows: int, cols: int) -> List[List[Any]]:
    out: List[List[Any]] = []
    for row in values[:rows]:
        fixed = list(row[:cols])
        fixed += [""] * (cols - len(fixed))
        out.append(fixed)
    while len(out) < rows:
        out.append([""] * cols)
    return out


def _build_sheet_tables(data: Dict[str, Any], scope: str = "core") -> Dict[str, Dict[str, Any]]:
    """生成 Google Sheet 展示页数据。

    scope="core"：只刷常用核心表，避免 Sheets API 429。
    scope="full"：刷全部展示表，低频使用。
    """
    data = normalize_data(data)
    m = calc_financials(data)

    tables: Dict[str, Dict[str, Any]] = {}
    def add(name: str, header: List[str], rows: List[List[Any]], min_rows: int = 80) -> None:
        tables[name] = {"header": header, "rows": rows, "min_rows": min_rows, "cols": len(header)}

    summary_rows = [
        ["更新时间", now_str(), "手动刷新展示页"],
        ["总资产", m["total_assets"], "现金 + 储蓄 + 应收贷款本金"],
        ["净资产", m["net_assets"], "总资产 - 负债"],
        ["现金余额", m["cash"], "可立即使用资金"],
        ["储蓄余额", m["savings"], "储蓄账户余额"],
        ["应收贷款本金", m["receivables"], "未收回本金"],
        ["负债余额", m["liabilities"], "未偿还借入资金"],
        ["今日可安全花", m["safe_to_spend"], "现金安全线和预算约束后的可用额"],
        ["信用分", score_text(m["credit_score"]), "家庭内部风控评分"],
        ["本月收入", m["month_income"], month_str()],
        ["本月消费", m["month_expense"], month_str()],
        ["待审批数量", len([r for r in data.get("pending_requests", []) if r.get("parent_status") == "待审批"]), "pending_requests"],
        ["开放赏金任务", len([b for b in data.get("bounties", []) if b.get("status") == "开放"]), "bounties"],
    ]
    add("summary", ["指标", "数值", "说明"], summary_rows, min_rows=40)

    tx_rows = [[tx.get("date", ""), tx.get("type", ""), tx.get("amount", 0), tx.get("category", ""), tx.get("party", ""), tx.get("memo", ""), tx.get("created_by", ""), tx.get("loan_id", ""), tx.get("bounty_id", ""), tx.get("id", "")] for tx in sorted(data.get("transactions", []), key=lambda x: str(x.get("date", "")), reverse=True)]
    add("transactions", ["日期", "类型", "金额", "分类", "对象/商户", "备注", "创建人", "贷款ID", "赏金ID", "交易ID"], tx_rows, min_rows=300)

    budget_rows = [[cat, row["limit"], row["spent"], row["remaining"], row["ratio"]] for cat, row in m["budget_usage"].items()]
    add("budgets", ["分类", "预算", "已花", "剩余", "使用率"], budget_rows, min_rows=80)

    pending_rows = [[r.get("created_at", ""), r.get("applicant", ""), r.get("request_type", ""), r.get("amount", 0), r.get("category", ""), r.get("python_decision", ""), r.get("parent_status", ""), r.get("final_status", ""), r.get("approved_by", ""), r.get("request_text", ""), r.get("parent_note", "")] for r in sorted(data.get("pending_requests", []), key=lambda x: str(x.get("created_at", "")), reverse=True)]
    add("pending_requests", ["申请时间", "申请人", "类型", "金额", "分类", "系统结论", "家长状态", "最终状态", "审批人", "原始申请", "家长备注"], pending_rows, min_rows=200)

    bounty_rows = [[b.get("title", ""), b.get("status", ""), b.get("reward_amount", 0), b.get("reward_points", 0), b.get("difficulty", ""), b.get("deadline", ""), b.get("created_by", ""), b.get("assigned_to", ""), b.get("submission_note", ""), b.get("reviewed_by", ""), b.get("parent_note", "")] for b in sorted(data.get("bounties", []), key=lambda x: str(x.get("created_at", "")), reverse=True)]
    add("bounties", ["任务", "状态", "赏金", "积分", "难度", "截止日", "发布人", "领取人", "提交说明", "审核人", "家长备注"], bounty_rows, min_rows=200)

    loan_rows = [[l.get("loan_id"), l.get("date"), l.get("borrower"), l.get("principal"), l.get("repaid"), l.get("remaining"), l.get("expected_repayment"), l.get("expected_interest"), l.get("interest_received"), l.get("due_date"), l.get("status"), l.get("memo")] for l in m["loans"]]
    add("loan_book", ["贷款ID", "日期", "借款人", "本金", "已还本金", "剩余本金", "预计回款", "预计利息", "已收利息", "到期日", "状态", "备注"], loan_rows, min_rows=120)

    if scope != "full":
        return tables

    goal_rows = [[g["name"], g["category"], g["target"], g["current"], g["remaining"], g["progress"], g.get("deadline", ""), g.get("days_left", ""), g.get("status", "进行中")] for g in m["goals"]]
    add("goals", ["目标", "分类", "目标金额", "当前金额", "剩余金额", "完成度", "截止日", "剩余天数", "状态"], goal_rows, min_rows=120)

    score_rows = [[h.get("time", ""), h.get("score", ""), h.get("change", ""), h.get("event", ""), h.get("reason", "")] for h in sorted(data.get("score_history", []), key=lambda x: str(x.get("time", "")), reverse=True)]
    add("score_history", ["时间", "信用分", "变化", "事件", "原因"], score_rows, min_rows=300)

    merchant_rows = []
    for mer in data.get("merchants", []):
        res = evaluate_merchant_access(data, mer)
        merchant_rows.append([mer.get("name", ""), mer.get("category", ""), mer.get("discount", 0), mer.get("required_score", 0), mer.get("category_budget_cap", 1.0), res.get("result", ""), "；".join(res.get("failed", [])), mer.get("note", "")])
    add("merchants", ["商户", "分类", "折扣", "最低信用分", "品类预算上限", "状态", "失败原因", "说明"], merchant_rows, min_rows=120)

    add("settings", ["设置项", "值"], [[k, v] for k, v in data.get("settings", {}).items()], min_rows=80)
    add("rules", ["规则", "当前值"], [[k, v] for k, v in data.get("rules", {}).items()], min_rows=80)

    report_rows = []
    for w in sorted(data.get("weekly_reports", []), key=lambda x: str(x.get("week_start", "")), reverse=True):
        wm = w.get("metrics", {}) or {}
        report_rows.append([w.get("week_start", ""), w.get("generated_at", ""), wm.get("income", ""), wm.get("expense", ""), wm.get("savings", ""), wm.get("score", ""), wm.get("top_category", ""), w.get("content", "")])
    add("weekly_report", ["周起始日", "生成时间", "收入", "消费", "储蓄", "信用分", "最大消费分类", "周报内容"], report_rows, min_rows=120)

    add("backups", ["备份ID", "时间", "操作人", "事件", "原因"], [[b.get("id", ""), b.get("time", ""), b.get("operator", ""), b.get("event", ""), b.get("reason", "")] for b in sorted(data.get("backups", []), key=lambda x: str(x.get("time", "")), reverse=True)], min_rows=120)
    add("badges", ["徽章", "获得时间", "说明"], [[b.get("name", ""), b.get("earned_at", ""), b.get("description", "")] for b in sorted(data.get("badges", []), key=lambda x: str(x.get("earned_at", "")), reverse=True)], min_rows=120)
    add("rewards", ["奖励", "创建时间", "创建人", "状态", "说明"], [[r.get("title", ""), r.get("created_at", ""), r.get("created_by", ""), r.get("status", ""), r.get("note", "")] for r in sorted(data.get("rewards", []), key=lambda x: str(x.get("created_at", "")), reverse=True)], min_rows=120)
    add("audit_log", ["时间", "操作人", "事件", "详情"], [[a.get("time", ""), a.get("operator", ""), a.get("event", ""), a.get("detail", "")] for a in sorted(data.get("audit_log", []), key=lambda x: str(x.get("time", "")), reverse=True)], min_rows=300)
    return tables


def sync_readable_sheets(data: Dict[str, Any], scope: str = "core") -> Dict[str, Any]:
    """低配额 Google Sheet 展示页刷新。

    兼容 gspread：不再调用 Spreadsheet.batch_clear()，因为部分 gspread 版本没有这个方法。
    做法：用 values_batch_update 一次性写入各展示页，并把目标区域补空白，顺便覆盖旧内容。
    已有 worksheet 情况下，核心刷新通常只需要 1 个 values 批量写请求。
    """
    if not gsheet_enabled():
        raise RuntimeError("Google Sheet 未配置。")
    scope = "full" if scope == "full" else "core"
    data = normalize_data(data)
    sh = get_spreadsheet()
    tables = _build_sheet_tables(data, scope=scope)

    existing_titles = {ws.title for ws in sh.worksheets()}
    for title, table in tables.items():
        if title not in existing_titles:
            # 新建 worksheet 仍会消耗写请求。首次初始化如果遇到 429，等一分钟后重试。
            sh.add_worksheet(title=title, rows=max(table["min_rows"], len(table["rows"]) + 10), cols=max(table["cols"], 2))

    updates: List[Dict[str, Any]] = []
    for title, table in tables.items():
        header = table["header"]
        rows = table["rows"]
        cols = table["cols"]
        # 写满固定区域，用空白覆盖旧数据，避免再单独 clear。
        row_count = max(table["min_rows"], len(rows) + 1)
        values = _pad_sheet_values([header] + rows, row_count, cols)
        updates.append({
            "range": _sheet_a1_range(title, row_count, cols),
            "values": values,
        })

    if updates:
        if hasattr(sh, "values_batch_update"):
            sh.values_batch_update({"valueInputOption": "RAW", "data": updates})
        else:
            # 极老 gspread 兜底：逐表 update。会更慢，但不会调用不存在的 batch_clear。
            for item in updates:
                title = item["range"].split("!", 1)[0].strip("'").replace("''", "'")
                ws = sh.worksheet(title)
                ws.update(item["range"].split("!", 1)[1], item["values"], value_input_option="RAW")

    return {"scope": scope, "sheet_count": len(tables), "updated_at": now_str()}

def build_excel_export(data: Dict[str, Any]) -> Optional[bytes]:
    if openpyxl is None:
        return None
    normalized = normalize_data(data); metrics = calc_financials(normalized)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame([{"指标": k, "数值": v} for k, v in {"总资产": metrics["total_assets"], "净资产": metrics["net_assets"], "现金": metrics["cash"], "储蓄": metrics["savings"], "应收贷款本金": metrics["receivables"], "负债": metrics["liabilities"], "信用分": metrics["credit_score"], "本月收入": metrics["month_income"], "本月消费": metrics["month_expense"]}.items()]).to_excel(writer, sheet_name="summary", index=False)
        pd.DataFrame(normalized.get("transactions", [])).to_excel(writer, sheet_name="transactions", index=False)
        pd.DataFrame(normalized.get("pending_requests", [])).to_excel(writer, sheet_name="pending", index=False)
        pd.DataFrame(normalized.get("bounties", [])).to_excel(writer, sheet_name="bounties", index=False)
        pd.DataFrame(normalized.get("goals", [])).to_excel(writer, sheet_name="goals", index=False)
        pd.DataFrame(normalized.get("score_history", [])).to_excel(writer, sheet_name="score_history", index=False)
        pd.DataFrame(normalized.get("audit_log", [])).to_excel(writer, sheet_name="audit_log", index=False)
        pd.DataFrame(metrics.get("loans", [])).to_excel(writer, sheet_name="loans", index=False)
        pd.DataFrame([{"分类": k, **v} for k, v in metrics.get("budget_usage", {}).items()]).to_excel(writer, sheet_name="budgets", index=False)
        pd.DataFrame(normalized.get("weekly_reports", [])).to_excel(writer, sheet_name="weekly_reports", index=False)
    output.seek(0)
    return output.getvalue()


# ============================================================
# 7. UI 组件和图表
# ============================================================

def inject_css() -> None:
    st.markdown(
        """
        <style>
        .block-container { padding-top: 1.1rem; max-width: 1380px; }
        .bank-hero { background: linear-gradient(135deg, #08245c 0%, #0b3d91 55%, #174ea6 100%); border-radius: 26px; padding: 28px 32px; color: white; box-shadow: 0 18px 40px rgba(11,61,145,.28); margin-bottom: 18px; }
        .bank-hero h1 { margin: 0; font-size: 2.25rem; letter-spacing: -.02em; }
        .bank-hero p { opacity: .88; margin: 8px 0 0; }
        .money-big { font-size: 2.8rem; font-weight: 800; line-height: 1.05; margin-top: 14px; }
        .metric-card { background: white; border: 1px solid #e5e7eb; border-radius: 22px; padding: 18px 20px; box-shadow: 0 8px 24px rgba(15,23,42,.06); min-height: 124px; }
        .metric-title { color: #6b7280; font-size: .9rem; margin-bottom: 8px; }
        .metric-value { color: #111827; font-size: 1.6rem; font-weight: 780; }
        .metric-note { color: #6b7280; font-size: .82rem; margin-top: 6px; }
        .decision { border-radius: 18px; padding: 16px 18px; margin: 12px 0; border: 1px solid #e5e7eb; background: white; }
        .decision.ok { border-left: 8px solid #15803d; }
        .decision.warn { border-left: 8px solid #d97706; }
        .decision.bad { border-left: 8px solid #b91c1c; }
        .pill { display:inline-block; border-radius:999px; padding:4px 10px; font-size:.78rem; font-weight:700; background:#edf2ff; color:#0b3d91; margin-bottom:8px; }
        .bounty { border-radius: 18px; padding: 16px 18px; border: 1px solid #dbeafe; background: linear-gradient(180deg,#ffffff,#f8fbff); margin-bottom: 12px; }
        .small { color:#6b7280; font-size:.88rem; }
        .login-card { max-width: 520px; margin: 8vh auto; background: white; border:1px solid #e5e7eb; border-radius:24px; padding:32px; box-shadow:0 20px 50px rgba(15,23,42,.12); }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_hero(data: Dict[str, Any], metrics: Dict[str, Any]) -> None:
    ccy = metrics["currency"]
    st.markdown(f"""
        <div class="bank-hero">
          <div class="pill">{APP_NAME} · {role_label()}</div>
          <h1>{data.get('settings', {}).get('owner', '阿苏')}的私人银行</h1>
          <p>主线：资产、信用、预算、审批、任务收入。阿苏可自己记账；爸爸妈妈负责审批和规则。</p>
          <div class="money-big">{money(metrics['total_assets'], ccy)}</div>
          <p>总资产 = 现金 + 储蓄 + 应收贷款本金；净资产 = 总资产 - 负债。</p>
        </div>
    """, unsafe_allow_html=True)


def metric_card(title: str, value: str, note: str = "") -> None:
    st.markdown(f"<div class='metric-card'><div class='metric-title'>{title}</div><div class='metric-value'>{value}</div><div class='metric-note'>{note}</div></div>", unsafe_allow_html=True)


def dataframe_or_empty(rows: List[Dict[str, Any]], columns: Optional[List[str]] = None, empty_text: str = "暂无数据。") -> None:
    if not rows:
        st.info(empty_text); return
    df = pd.DataFrame(rows)
    if columns:
        df = df[[c for c in columns if c in df.columns]]
    st.dataframe(df, use_container_width=True, hide_index=True)


def plotly_enabled() -> bool:
    return px is not None and go is not None


def next_chart_key(prefix: str = "plotly") -> str:
    """Avoid StreamlitDuplicateElementId when identical Plotly figures appear in different tabs."""
    st.session_state["_asu_plotly_chart_counter"] = st.session_state.get("_asu_plotly_chart_counter", 0) + 1
    return f"{prefix}_{st.session_state['_asu_plotly_chart_counter']}"


def render_asset_mix_chart(metrics: Dict[str, Any]) -> None:
    rows = [{"资产类别": "现金", "金额": max(0.0, fnum(metrics.get("cash")))}, {"资产类别": "储蓄", "金额": max(0.0, fnum(metrics.get("savings")))}, {"资产类别": "应收贷款本金", "金额": max(0.0, fnum(metrics.get("receivables")))}]
    df = pd.DataFrame([r for r in rows if r["金额"] > 0])
    if df.empty: st.info("暂无资产结构图。"); return
    if plotly_enabled():
        fig = px.pie(df, names="资产类别", values="金额", title="资产结构")
        fig.update_layout(margin=dict(l=10, r=10, t=50, b=10), height=320)
        st.plotly_chart(fig, use_container_width=True, key=next_chart_key("plotly"))
    else:
        st.bar_chart(df.set_index("资产类别"))


def render_budget_usage_chart(metrics: Dict[str, Any]) -> None:
    rows = []
    for cat, row in metrics.get("budget_usage", {}).items():
        limit = fnum(row.get("limit"))
        if limit <= 0: continue
        spent = fnum(row.get("spent"))
        rows.append({"分类": cat, "已花": spent, "剩余预算": max(0.0, limit - spent)})
    df = pd.DataFrame(rows)
    if df.empty: st.info("暂无预算图。"); return
    if plotly_enabled():
        fig = px.bar(df, x="分类", y=["已花", "剩余预算"], barmode="stack", title="预算使用情况")
        fig.update_layout(margin=dict(l=10, r=10, t=50, b=10), height=340)
        st.plotly_chart(fig, use_container_width=True, key=next_chart_key("plotly"))
    else:
        st.bar_chart(df.set_index("分类"))


def render_credit_score_chart(score_history: List[Dict[str, Any]]) -> None:
    df = pd.DataFrame(score_history)
    if df.empty or "score" not in df.columns: st.info("暂无信用分走势。"); return
    df["time"] = pd.to_datetime(df.get("time"), errors="coerce")
    df["score"] = pd.to_numeric(df.get("score"), errors="coerce")
    df = df.dropna(subset=["time", "score"]).sort_values("time")
    if df.empty: st.info("暂无可绘制的信用分走势。"); return
    if np is not None and len(df) >= 3:
        df["三次移动平均"] = df["score"].rolling(3, min_periods=1).mean()
    if plotly_enabled():
        y_cols = ["score"] + (["三次移动平均"] if "三次移动平均" in df.columns else [])
        fig = px.line(df, x="time", y=y_cols, markers=True, title="信用分走势")
        fig.update_layout(margin=dict(l=10, r=10, t=50, b=10), height=340, yaxis_title="信用分", xaxis_title="时间")
        st.plotly_chart(fig, use_container_width=True, key=next_chart_key("plotly"))
    else:
        st.line_chart(df.set_index("time")[["score"]])


def render_monthly_spending_chart(rows: List[Dict[str, Any]], currency: str) -> None:
    by_cat: Dict[str, float] = {}
    for tx in rows:
        if tx.get("type") == "消费": by_cat[tx.get("category") or "其他"] = by_cat.get(tx.get("category") or "其他", 0.0) + fnum(tx.get("amount"))
    if not by_cat: st.info("该月暂无消费分类图。"); return
    df = pd.DataFrame({"分类": list(by_cat.keys()), "金额": list(by_cat.values())}).sort_values("金额", ascending=False)
    if plotly_enabled():
        fig = px.bar(df, x="分类", y="金额", title=f"本月消费分类（{currency}）")
        fig.update_layout(margin=dict(l=10, r=10, t=50, b=10), height=340)
        st.plotly_chart(fig, use_container_width=True, key=next_chart_key("plotly"))
    else:
        st.bar_chart(df.set_index("分类"))


def render_monthly_cashflow_trend(data: Dict[str, Any]) -> None:
    rows = []
    for tx in data.get("transactions", []):
        mth = tx_month(tx)
        if not mth: continue
        typ = tx.get("type"); amount = fnum(tx.get("amount"))
        if typ == "收入": rows.append({"月份": mth, "项目": "收入", "金额": amount})
        elif typ == "消费": rows.append({"月份": mth, "项目": "消费", "金额": amount})
        elif typ == "还款": rows.append({"月份": mth, "项目": "利息收入", "金额": fnum(tx.get("interest_received"))})
    df = pd.DataFrame(rows)
    if df.empty: st.info("暂无月度现金流趋势。"); return
    grouped = df.groupby(["月份", "项目"], as_index=False)["金额"].sum().sort_values("月份")
    if plotly_enabled():
        fig = px.bar(grouped, x="月份", y="金额", color="项目", barmode="group", title="月度现金流趋势")
        fig.update_layout(margin=dict(l=10, r=10, t=50, b=10), height=340)
        st.plotly_chart(fig, use_container_width=True, key=next_chart_key("plotly"))
    else:
        st.bar_chart(grouped.pivot(index="月份", columns="项目", values="金额").fillna(0))


def render_decision(decision: Dict[str, Any]) -> None:
    css = decision_class(decision.get("result", "观察"))
    st.markdown(f"<div class='decision {css}'><div class='pill'>{decision.get('kind')}</div><h3>结论：{decision.get('result')}</h3><p class='small'>{decision.get('action','')}</p></div>", unsafe_allow_html=True)
    c1, c2 = st.columns([1, 1])
    with c1:
        st.markdown("#### 风控原因")
        for r in decision.get("reasons", []): st.write(f"- {r}")
    with c2:
        st.markdown("#### 关键指标")
        metric_rows = [{"指标": k, "数值": pretty_value(k, v)} for k, v in decision.get("metrics", {}).items()]
        st.dataframe(pd.DataFrame(metric_rows), use_container_width=True, hide_index=True)


def render_sidebar(data: Dict[str, Any]) -> None:
    st.sidebar.title("私人银行控制台")
    if configured_passwords():
        st.sidebar.success(f"已登录：{current_operator()}")
        if st.sidebar.button("退出登录"): logout()
    else:
        current = st.session_state.get("current_operator", "爸爸")
        idx = OPERATORS.index(current) if current in OPERATORS else 0
        st.sidebar.selectbox("当前操作人", OPERATORS, index=idx, key="current_operator", help="正式部署建议配置 DAD_PASSWORD / MOM_PASSWORD / ASU_PASSWORD。")
        st.sidebar.warning("未启用密码登录，当前身份可手动切换。")
    st.sidebar.caption(f"存储：{st.session_state.get('storage_backend', '未读取')}")
    if can_parent(): st.sidebar.success("爸爸妈妈权限相同：审批、后台、设置、赏金任务。")
    else: st.sidebar.info("阿苏可以提交申请、领取任务、提交完成，也可以自己新增/修正交易。")
    m = calc_financials(data)
    st.sidebar.metric("总资产", money(m["total_assets"], m["currency"]))
    st.sidebar.metric("信用分", score_text(m["credit_score"]))
    st.sidebar.metric("今日可安全花", money(m["safe_to_spend"], m["currency"]))


# ============================================================
# 8. 页面：阿苏首页
# ============================================================

def page_asu_home(data: Dict[str, Any]) -> None:
    m = calc_financials(data); ccy = m["currency"]
    render_hero(data, m)
    c1, c2, c3, c4 = st.columns(4)
    with c1: metric_card("净资产", money(m["net_assets"], ccy), "总资产 - 负债")
    with c2: metric_card("现金", money(m["cash"], ccy), "可立即使用，但要守安全线")
    with c3: metric_card("今日可安全花", money(m["safe_to_spend"], ccy), "现金安全线 + 预算剩余额")
    with c4: metric_card("信用分", score_text(m["credit_score"]), "越稳，权限越大")

    st.markdown("### 我要申请")
    with st.form("asu_quick_request"):
        text = st.text_input("申请内容", placeholder="例如：我想买一个 $12 的 Minecraft 皮肤 / 我想借给爸爸 $20 下周还 $22")
        submitted = st.form_submit_button("评估并提交申请", type="primary", use_container_width=True)
    if submitted:
        decision = parse_natural_request(data, text)
        submit_pending_request(data, decision, text, applicant=current_operator())
        commit(data, event="阿苏提交申请", reason=text)
        st.success("申请已进入爸爸/妈妈审批队列。")
        st.rerun()

    st.markdown("### 快速记一笔")
    st.caption("阿苏拥有记账权限。这里是简化入口；完整入口在“私人银行后台 → 新增交易”。")
    with st.form("asu_fast_tx"):
        c1, c2, c3 = st.columns(3)
        with c1: tx_type = st.selectbox("类型", ["收入", "消费", "转入储蓄", "储蓄取出"], key="asu_fast_type")
        with c2: amount = st.number_input("金额", min_value=0.0, value=5.0, step=1.0, format="%.2f", key="asu_fast_amount")
        with c3: category = st.selectbox("分类", category_options(data), key="asu_fast_category")
        memo = st.text_input("备注", key="asu_fast_memo")
        ok = st.form_submit_button("保存这笔交易", type="primary")
    if ok:
        data.setdefault("transactions", []).append(make_tx(tx_type, amount, category, memo=memo))
        commit(data, event="阿苏快速记账", reason=f"{tx_type} {money(amount)} {category}")
        st.success("已保存。")
        st.rerun()

    st.markdown("### 目标和任务")
    left, right = st.columns([1, 1])
    with left:
        st.markdown("#### 我的储蓄目标")
        shown_goals = active_goals_from_metrics(m)
        if not shown_goals: st.info("还没有进行中的储蓄目标。已完成目标会记录在后台历史里。")
        for g in shown_goals:
            st.write(f"**{g['name']}** · {money(g['current'], ccy)} / {money(g['target'], ccy)}")
            st.progress(min(1.0, fnum(g["progress"])))
            st.caption(f"还差 {money(g['remaining'], ccy)}；截止日：{g.get('deadline') or '未设置'}")
    with right:
        st.markdown("#### 可领取赏金任务")
        open_bounties = [b for b in active_bounties(data) if b.get("status") == "开放"]
        if not open_bounties: st.info("暂无开放任务。")
        for b in open_bounties[:6]:
            st.markdown(f"<div class='bounty'><h4>{b.get('title')}</h4><p class='small'>{b.get('description')}</p><div class='pill'>赏金 {money(b.get('reward_amount'), ccy)} · {b.get('reward_points')}分 · {b.get('difficulty')}</div></div>", unsafe_allow_html=True)
            if st.button("领取这个任务", key=f"claim_{b['id']}", disabled=current_operator() != CHILD):
                b["status"] = "已领取"; b["assigned_to"] = current_operator(); b["claimed_at"] = now_str()
                commit(data, event="领取赏金任务", reason=b.get("title", "")); st.rerun()

    st.markdown("### 我的任务进度")
    my_bounties = [b for b in active_bounties(data) if b.get("assigned_to") == current_operator() or (current_operator() == CHILD and b.get("assigned_to") == CHILD)]
    if not my_bounties: st.info("你还没有领取任务。")
    for b in my_bounties:
        with st.expander(f"{b.get('title')} · {b.get('status')} · {money(b.get('reward_amount'), ccy)}", expanded=b.get("status") in {"已领取", "已退回"}):
            st.write(b.get("description") or "无说明")
            if b.get("status") in {"已领取", "已退回"} and current_operator() == CHILD:
                note = st.text_area("完成说明", value=b.get("submission_note", ""), key=f"submit_note_{b['id']}")
                if st.button("提交完成，等待家长审核", key=f"submit_bounty_{b['id']}"):
                    b["status"] = "已提交"; b["submission_note"] = note; b["submitted_at"] = now_str()
                    commit(data, event="提交赏金任务", reason=b.get("title", "")); st.rerun()

    st.markdown("### 我的审批状态")
    mine = [r for r in data.get("pending_requests", []) if r.get("applicant") in {current_operator(), CHILD}]
    rows = [{"时间": r.get("created_at"), "内容": r.get("request_text"), "金额": money(r.get("amount"), ccy), "系统结论": r.get("python_decision"), "家长状态": r.get("parent_status"), "最终状态": r.get("final_status"), "审批人": r.get("approved_by")} for r in sorted(mine, key=lambda x: str(x.get("created_at", "")), reverse=True)[:10]]
    dataframe_or_empty(rows, empty_text="暂无申请记录。")


# ============================================================
# 9. 页面：家长工作台
# ============================================================

def approve_pending_request(data: Dict[str, Any], req: Dict[str, Any], note: str = "") -> None:
    tx = req.get("pending_tx")
    if tx:
        tx = normalize_tx(tx); tx["created_by"] = current_operator(); tx["created_at"] = now_str()
        data.setdefault("transactions", []).append(tx)
        req["final_status"] = "已入账"; req["booked_at"] = now_str()
    else:
        req["final_status"] = "已批准未入账"
    req["parent_status"] = "已批准"; req["approved_by"] = current_operator(); req["approved_at"] = now_str(); req["parent_note"] = note


def reject_pending_request(req: Dict[str, Any], note: str = "") -> None:
    req["parent_status"] = "已拒绝"; req["final_status"] = "不入账"; req["approved_by"] = current_operator(); req["approved_at"] = now_str(); req["parent_note"] = note


def page_parent_workspace(data: Dict[str, Any]) -> None:
    st.title("家长工作台")
    if not can_parent(): st.warning("当前身份不是爸爸/妈妈，只能查看部分信息。")
    tab1, tab2, tab3, tab4 = st.tabs(["审批中心", "发布赏金任务", "任务审核", "AI评估器"])
    with tab1:
        st.subheader("待审批申请")
        pending = [r for r in data.get("pending_requests", []) if r.get("parent_status") == "待审批"]
        if not pending: st.info("暂无待审批申请。")
        for r in sorted(pending, key=lambda x: str(x.get("created_at", "")), reverse=True):
            css = decision_class(r.get("python_decision", "观察"))
            st.markdown(f"<div class='decision {css}'><div class='pill'>{r.get('python_decision')}</div><h3>{r.get('request_type')}：{money(r.get('amount'))} · {r.get('category')}</h3><p class='small'>{r.get('request_text')}</p><p class='small'>申请人：{r.get('applicant')}；AI分类：{r.get('ai_category')}；必要性：{r.get('necessity')}；冲动等级：{r.get('impulse_level')}</p></div>", unsafe_allow_html=True)
            note = st.text_input("家长备注", key=f"parent_note_{r['id']}")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("批准并入账", key=f"approve_{r['id']}", disabled=not can_parent(), type="primary"):
                    approve_pending_request(data, r, note); commit(data, event="批准申请", reason=r.get("request_text", "")); st.rerun()
            with c2:
                if st.button("拒绝", key=f"reject_{r['id']}", disabled=not can_parent()):
                    reject_pending_request(r, note); commit(data, event="拒绝申请", reason=r.get("request_text", "")); st.rerun()
    with tab2:
        st.subheader("发布赏金任务")
        with st.form("bounty_form"):
            title = st.text_input("任务标题", "整理书桌并拍照提交")
            desc = st.text_area("任务说明", "把书桌清理干净，书本分类，垃圾扔掉，最后拍照或文字说明。")
            c1, c2, c3 = st.columns(3)
            with c1: reward_amount = st.number_input("赏金金额", min_value=0.0, value=5.0, step=1.0, format="%.2f")
            with c2: reward_points = st.number_input("信用积分", min_value=0, value=5, step=1)
            with c3: difficulty = st.selectbox("难度", ["简单", "普通", "困难"])
            c4, c5 = st.columns(2)
            with c4: category = st.text_input("任务分类", "家庭任务")
            with c5: deadline = st.date_input("截止日", value=date.today() + timedelta(days=7)).isoformat()
            submit = st.form_submit_button("发布任务", type="primary", disabled=not can_parent())
        if submit:
            data.setdefault("bounties", []).append({"id": uid(), "title": title.strip() or "未命名任务", "description": desc.strip(), "reward_amount": reward_amount, "reward_points": int(reward_points), "category": category.strip() or "家庭任务", "difficulty": difficulty, "deadline": deadline, "status": "开放", "created_by": current_operator(), "created_at": now_str(), "assigned_to": "", "claimed_at": "", "submitted_at": "", "submission_note": "", "reviewed_by": "", "reviewed_at": "", "parent_note": "", "paid_tx_id": ""})
            commit(data, event="发布赏金任务", reason=title); st.success("赏金任务已发布。"); st.rerun()
        rows = [{"标题": b.get("title"), "状态": b.get("status"), "赏金": money(b.get("reward_amount")), "积分": b.get("reward_points"), "截止日": b.get("deadline"), "发布人": b.get("created_by"), "领取人": b.get("assigned_to")} for b in sorted(active_bounties(data), key=lambda x: str(x.get("created_at", "")), reverse=True)]
        dataframe_or_empty(rows, empty_text="暂无当前赏金任务。已支付/已取消任务已进入历史。")
        if historical_bounties(data):
            with st.expander(f"已完成/已取消任务历史（{len(historical_bounties(data))}）", expanded=False):
                history_rows = [{"标题": b.get("title"), "状态": b.get("status"), "赏金": money(b.get("reward_amount")), "领取人": b.get("assigned_to"), "审核人": b.get("reviewed_by"), "审核时间": b.get("reviewed_at")} for b in sorted(historical_bounties(data), key=lambda x: str(x.get("reviewed_at") or x.get("created_at") or ""), reverse=True)]
                dataframe_or_empty(history_rows)
        st.divider()
        st.markdown("#### 修改/删除当前任务")
        active_for_edit = sorted(active_bounties(data), key=lambda x: str(x.get("created_at", "")), reverse=True)
        if not active_for_edit:
            st.caption("暂无可修改任务。")
        for b in active_for_edit:
            bid = b.get("id") or uid()
            with st.expander(f"管理：{b.get('title')} · {b.get('status')}", expanded=False):
                with st.form(f"parent_edit_bounty_{bid}"):
                    c1, c2, c3 = st.columns(3)
                    title2 = c1.text_input("任务标题", b.get("title", ""), key=f"parent_edit_title_{bid}")
                    amount2 = c2.number_input("赏金金额", min_value=0.0, value=fnum(b.get("reward_amount")), step=1.0, format="%.2f", key=f"parent_edit_amount_{bid}")
                    points2 = c3.number_input("信用积分", min_value=0, value=inum(b.get("reward_points"), 0), step=1, key=f"parent_edit_points_{bid}")
                    desc2 = st.text_area("任务说明", b.get("description", ""), key=f"parent_edit_desc_{bid}")
                    category2 = c1.text_input("任务分类", b.get("category", "家庭任务"), key=f"parent_edit_category_{bid}")
                    difficulty2 = c2.selectbox("难度", ["简单", "普通", "困难"], index=["简单", "普通", "困难"].index(b.get("difficulty")) if b.get("difficulty") in ["简单", "普通", "困难"] else 1, key=f"parent_edit_difficulty_{bid}")
                    old_deadline2 = safe_date(b.get("deadline")) or date.today() + timedelta(days=7)
                    deadline2 = c3.date_input("截止日", value=old_deadline2, key=f"parent_edit_deadline_{bid}").isoformat()
                    saved2 = st.form_submit_button("保存修改", disabled=not can_parent())
                if saved2:
                    b.update({"title": title2.strip() or "未命名任务", "description": desc2.strip(), "reward_amount": amount2, "reward_points": int(points2), "category": category2.strip() or "家庭任务", "difficulty": difficulty2, "deadline": deadline2})
                    commit(data, event="修改赏金任务", reason=b.get("title", ""))
                    st.rerun()
                c1, c2 = st.columns(2)
                if c1.button("取消并转入历史", key=f"parent_cancel_bounty_{bid}", disabled=not can_parent()):
                    b["status"] = "已取消"; b["reviewed_by"] = current_operator(); b["reviewed_at"] = now_str()
                    commit(data, event="取消赏金任务", reason=b.get("title", ""))
                    st.rerun()
                if c2.button("删除未完成任务", key=f"parent_delete_bounty_{bid}", disabled=not can_parent()):
                    data["bounties"] = [x for x in data.get("bounties", []) if x.get("id") != bid]
                    commit(data, event="删除赏金任务", reason=b.get("title", ""))
                    st.rerun()
    with tab3:
        st.subheader("任务审核")
        submitted = [b for b in data.get("bounties", []) if b.get("status") == "已提交"]
        if not submitted: st.info("暂无待审核任务。")
        for b in submitted:
            st.markdown(f"<div class='bounty'><h4>{b.get('title')}</h4><p class='small'>{b.get('description')}</p><p>提交说明：{b.get('submission_note') or '未填写'}</p><div class='pill'>赏金 {money(b.get('reward_amount'))} · 积分 {b.get('reward_points')}</div></div>", unsafe_allow_html=True)
            note = st.text_input("审核备注", key=f"review_note_{b['id']}")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("通过并发放赏金", key=f"approve_bounty_{b['id']}", disabled=not can_parent(), type="primary"):
                    b["status"] = "已支付"; b["reviewed_by"] = current_operator(); b["reviewed_at"] = now_str(); b["parent_note"] = note
                    tx = make_tx("收入", fnum(b.get("reward_amount")), "赏金任务", party=current_operator(), memo=f"赏金任务：{b.get('title')}", bounty_id=b.get("id"))
                    b["paid_tx_id"] = tx["id"]; data.setdefault("transactions", []).append(tx)
                    commit(data, event="发放赏金", reason=b.get("title", "")); st.rerun()
            with c2:
                if st.button("退回修改", key=f"return_bounty_{b['id']}", disabled=not can_parent()):
                    b["status"] = "已退回"; b["reviewed_by"] = current_operator(); b["reviewed_at"] = now_str(); b["parent_note"] = note
                    commit(data, event="退回赏金任务", reason=b.get("title", "")); st.rerun()
    with tab4:
        st.subheader("AI 决策评估器")
        text = st.text_input("输入一条申请", "我想买一个 $12 的 Minecraft 皮肤", key="parent_ai_eval_text")
        if st.button("生成评估", type="primary"):
            st.session_state["parent_ai_eval_decision"] = parse_natural_request(data, text)
            st.session_state["parent_ai_eval_source_text"] = text
        decision = st.session_state.get("parent_ai_eval_decision")
        if decision:
            render_decision(decision)
            source_text = st.session_state.get("parent_ai_eval_source_text", text)
            amount = fnum(decision.get("metrics", {}).get("购买金额")); cat = str(decision.get("metrics", {}).get("消费分类") or "其他")
            if amount > 0 and decision.get("kind") == "消费审批": st.dataframe(pd.DataFrame(scenario_planning(data, amount, cat, source_text, str(decision.get("metrics", {}).get("商户") or ""))), use_container_width=True, hide_index=True)
            st.markdown(deepseek_decision_report(decision))
            if st.button("把这条评估加入待审批", key="add_eval_pending"):
                submit_pending_request(data, decision, source_text, applicant=CHILD); commit(data, event="家长代提交评估申请", reason=source_text)
                st.session_state.pop("parent_ai_eval_decision", None); st.session_state.pop("parent_ai_eval_source_text", None); st.rerun()


# ============================================================
# 10. 页面：私人银行后台
# ============================================================

def page_bank_backend(data: Dict[str, Any]) -> None:
    st.title("私人银行后台")
    if not can_parent():
        st.info("阿苏当前可使用“新增交易”栏目进行自主记账；预算、设置、审批等敏感操作仍由爸爸/妈妈处理。")
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs(["总览", "新增交易", "贷款台账", "预算/目标", "月度账单", "风险雷达", "商户权益", "周报"])
    with tab1:
        m = calc_financials(data); render_hero(data, m)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("净资产", money(m["net_assets"], m["currency"])); c2.metric("现金", money(m["cash"], m["currency"])); c3.metric("储蓄", money(m["savings"], m["currency"])); c4.metric("应收贷款", money(m["receivables"], m["currency"]))
        c5, c6, c7, c8 = st.columns(4)
        c5.metric("本月收入", money(m["month_income"], m["currency"])); c6.metric("本月消费", money(m["month_expense"], m["currency"])); c7.metric("支出收入比", percent(m["spend_income_ratio"])); c8.metric("信用分", score_text(m["credit_score"]))
        st.markdown("#### 资产与现金流图表")
        left, right = st.columns(2)
        with left: render_asset_mix_chart(m)
        with right: render_monthly_cashflow_trend(data)
        if data.get("score_history"):
            st.markdown("#### 信用分走势"); render_credit_score_chart(data.get("score_history", []))
        st.markdown("#### 最近交易")
        rows = [{"日期": tx.get("date"), "类型": tx.get("type"), "金额": money(tx.get("amount"), m["currency"]), "分类": tx.get("category"), "对象": tx.get("party"), "创建人": tx.get("created_by"), "备注": tx.get("memo")} for tx in sorted(data.get("transactions", []), key=lambda x: str(x.get("date", "")), reverse=True)[:20]]
        dataframe_or_empty(rows, empty_text="暂无交易。")
    with tab2:
        st.subheader("新增交易")
        st.caption("阿苏、爸爸、妈妈都可以记账。删除和修正会写入审计日志，并自动备份。")
        with st.form("tx_form"):
            c1, c2, c3 = st.columns(3)
            with c1:
                tx_date = st.date_input("日期", value=date.today()).isoformat(); tx_type = st.selectbox("类型", TX_TYPES)
            with c2:
                amount = st.number_input("金额", min_value=0.0, value=5.0, step=1.0, format="%.2f"); category = st.selectbox("分类", category_options(data))
            with c3:
                party = st.text_input("对象/商户/借款人", ""); due_date = st.date_input("到期日/还款日", value=date.today() + timedelta(days=7)).isoformat()
            memo = st.text_input("备注", "")
            c4, c5, c6 = st.columns(3)
            with c4: expected = st.number_input("预计回款", min_value=0.0, value=0.0, step=1.0, format="%.2f")
            with c5: principal_repaid = st.number_input("还款本金", min_value=0.0, value=0.0, step=1.0, format="%.2f")
            with c6: interest_received = st.number_input("收到利息", min_value=0.0, value=0.0, step=1.0, format="%.2f")
            submitted = st.form_submit_button("保存交易", type="primary", disabled=not can_bookkeep())
        if submitted:
            tx = make_tx(tx_type, amount, category, party=party, memo=memo, tx_date=tx_date, expected_repayment=expected, principal_repaid=principal_repaid, interest_received=interest_received, due_date=due_date if tx_type == "放贷" else "")
            data.setdefault("transactions", []).append(tx)
            commit(data, event="新增交易", reason=f"{current_operator()} 记账：{tx_type} {money(amount)} {category}")
            st.success("交易已保存。"); st.rerun()
        st.markdown("#### 交易流水与单笔删除")
        m = calc_financials(data)
        rows = [{"日期": tx.get("date"), "类型": tx.get("type"), "金额": money(tx.get("amount"), m["currency"]), "分类": tx.get("category"), "对象": tx.get("party"), "创建人": tx.get("created_by"), "备注": tx.get("memo"), "id": tx.get("id")} for tx in sorted(data.get("transactions", []), key=lambda x: str(x.get("date", "")), reverse=True)]
        dataframe_or_empty([{k: v for k, v in r.items() if k != "id"} for r in rows], empty_text="暂无交易。")
        if rows:
            options = [f"{r['日期']} · {r['类型']} · {r['金额']} · {r['分类']} · {r['备注']} · {r['id']}" for r in rows]
            selected = st.selectbox("选择要删除的单笔交易", options, key="delete_tx_select")
            selected_id = selected.split(" · ")[-1]
            confirm = st.checkbox("确认删除这笔交易", key="delete_tx_confirm")
            if st.button("删除选中交易", disabled=(not confirm or not can_bookkeep())):
                before_len = len(data.get("transactions", []))
                data["transactions"] = [tx for tx in data.get("transactions", []) if tx.get("id") != selected_id]
                if len(data["transactions"]) < before_len:
                    commit(data, event="删除单笔交易", reason=selected)
                    st.success("已删除。"); st.rerun()
                else:
                    st.error("未找到该交易。")
    with tab3:
        st.subheader("贷款台账")
        m = calc_financials(data)
        rows = [{"贷款ID": l.get("loan_id"), "日期": l.get("date"), "借款人": l.get("borrower"), "本金": money(l.get("principal"), m["currency"]), "已还本金": money(l.get("repaid"), m["currency"]), "剩余本金": money(l.get("remaining"), m["currency"]), "预计回款": money(l.get("expected_repayment"), m["currency"]), "已收利息": money(l.get("interest_received"), m["currency"]), "到期日": l.get("due_date"), "状态": l.get("status")} for l in m["loans"]]
        dataframe_or_empty(rows, empty_text="暂无贷款。")
        st.divider(); st.markdown("#### 贷款审批模拟")
        with st.form("loan_eval_form"):
            c1, c2, c3, c4 = st.columns(4)
            with c1: principal = st.number_input("拟借出本金", min_value=0.0, value=10.0, step=1.0, format="%.2f")
            with c2: borrower = st.text_input("借款人", "爸爸")
            with c3: expected = st.number_input("预计回款", min_value=0.0, value=11.0, step=1.0, format="%.2f")
            with c4: due = st.date_input("预计还款日", value=date.today() + timedelta(days=7)).isoformat()
            memo = st.text_input("说明", "家庭短期借款")
            submit = st.form_submit_button("评估贷款")
        if submit:
            st.session_state["loan_eval_decision"] = evaluate_loan_decision(data, principal, borrower, expected, due, memo)
            st.session_state["loan_eval_text"] = f"放贷给{borrower} {money(principal)}，预计回款 {money(expected)}，到期 {due}"
        loan_decision = st.session_state.get("loan_eval_decision")
        if loan_decision:
            render_decision(loan_decision)
            if st.button("加入待审批队列", key="loan_add_pending"):
                submit_pending_request(data, loan_decision, st.session_state.get("loan_eval_text", "贷款审批"), applicant=CHILD)
                commit(data, event="提交贷款审批", reason=str(loan_decision.get("metrics", {}).get("借款人") or "贷款审批"))
                st.session_state.pop("loan_eval_decision", None); st.session_state.pop("loan_eval_text", None); st.rerun()
    with tab4:
        st.subheader("预算管理")
        m = calc_financials(data)
        usage_rows = [{"分类": cat, "预算": money(row["limit"], m["currency"]), "已花": money(row["spent"], m["currency"]), "剩余": money(row["remaining"], m["currency"]), "使用率": percent(row["ratio"])} for cat, row in m["budget_usage"].items()]
        dataframe_or_empty(usage_rows, empty_text="暂无预算。"); render_budget_usage_chart(m)
        with st.form("budget_form"):
            c1, c2 = st.columns(2)
            with c1: cat = st.text_input("预算分类", "游戏")
            with c2: limit = st.number_input("月度预算", min_value=0.0, value=20.0, step=1.0, format="%.2f")
            submit = st.form_submit_button("保存预算", type="primary", disabled=not can_parent())
        if submit:
            data.setdefault("budgets", {})[cat.strip() or "其他"] = limit
            commit(data, event="保存预算", reason=f"{cat}={limit}"); st.rerun()
        st.divider(); page_goals_full(data)
    with tab5:
        st.subheader("月度账单")
        selected_month = st.text_input("月份", month_str())
        rows = [tx for tx in data.get("transactions", []) if tx_month(tx) == selected_month]
        m2 = calc_financials(data, selected_month)
        c1, c2, c3 = st.columns(3)
        c1.metric("月收入", money(m2["month_income"], m2["currency"])); c2.metric("月消费", money(m2["month_expense"], m2["currency"])); c3.metric("支出收入比", percent(m2["spend_income_ratio"]))
        render_monthly_spending_chart(rows, m2["currency"])
        dataframe_or_empty([{"日期": tx.get("date"), "类型": tx.get("type"), "金额": money(tx.get("amount"), m2["currency"]), "分类": tx.get("category"), "对象": tx.get("party"), "创建人": tx.get("created_by"), "备注": tx.get("memo")} for tx in rows], empty_text="该月暂无交易。")
    with tab6:
        st.subheader("风险雷达")
        risks = build_risk_radar(data)
        for level in ["高风险", "中风险", "低风险"]:
            st.markdown(f"#### {level}")
            if not risks[level]: st.caption("无。")
            for r in risks[level]:
                css = "bad" if level == "高风险" else "warn" if level == "中风险" else "ok"
                st.markdown(f"<div class='decision {css}'><strong>{r['title']}</strong><p class='small'>{r['detail']}</p></div>", unsafe_allow_html=True)
        st.markdown("#### 本周行动")
        for i, a in enumerate(build_weekly_plan(data), 1): st.write(f"{i}. {a}")
    with tab7:
        st.subheader("商户权益")
        if not data.get("merchants"): st.info("暂无商户。")
        for mer in data.get("merchants", []):
            res = evaluate_merchant_access(data, mer); css = decision_class(res["result"])
            st.markdown(f"<div class='decision {css}'><div class='pill'>{res['result']}</div><h3>{mer.get('name')} · {mer.get('category')}</h3><p class='small'>{res['message']}</p><p class='small'>失败条件：{'、'.join(res['failed']) if res['failed'] else '无'}</p></div>", unsafe_allow_html=True)
        with st.form("merchant_form"):
            c1, c2, c3 = st.columns(3)
            with c1: name = st.text_input("商户名称", "书店奖励"); category = st.selectbox("分类", category_options(data), key="merchant_cat")
            with c2: discount = st.number_input("折扣比例", min_value=0.0, max_value=1.0, value=0.10, step=0.05); required_score = st.number_input("最低信用分", min_value=0, max_value=850, value=720)
            with c3: cap = st.number_input("品类预算开放上限", min_value=0.0, max_value=2.0, value=0.90, step=0.05); note = st.text_input("说明", "信用好时开放")
            submit = st.form_submit_button("新增商户权益", disabled=not can_parent())
        if submit:
            data.setdefault("merchants", []).append({"id": uid(), "name": name, "category": category, "discount": discount, "required_score": required_score, "category_budget_cap": cap, "note": note})
            commit(data, event="新增商户权益", reason=name); st.rerun()
    with tab8:
        st.subheader("周报")
        if st.button("生成本周周报", type="primary", disabled=not can_parent()):
            report = generate_weekly_report(data); data.setdefault("weekly_reports", []).append(report); data["weekly_reports"] = data["weekly_reports"][-52:]
            commit(data, event="生成周报", reason=report["week_start"]); st.rerun()
        if data.get("weekly_reports"):
            latest = data["weekly_reports"][-1]; st.text_area("最新周报", latest.get("content", ""), height=260)
        rows = [{"周起始日": w.get("week_start"), "生成时间": w.get("generated_at"), "收入": money((w.get("metrics") or {}).get("income")), "消费": money((w.get("metrics") or {}).get("expense")), "信用分": (w.get("metrics") or {}).get("score"), "最大消费分类": (w.get("metrics") or {}).get("top_category")} for w in sorted(data.get("weekly_reports", []), key=lambda x: x.get("week_start", ""), reverse=True)]
        dataframe_or_empty(rows, empty_text="暂无周报。")


# ============================================================
# 11. 页面：系统设置
# ============================================================

def page_settings(data: Dict[str, Any]) -> None:
    st.title("系统设置")
    if not can_parent(): st.warning("系统设置需要爸爸或妈妈权限。")
    tab1, tab2, tab3, tab4 = st.tabs(["账户设置", "风控规则", "Google Sheet/备份", "数据导出"])
    with tab1:
        s = data.setdefault("settings", {})
        with st.form("settings_form"):
            c1, c2 = st.columns(2)
            with c1:
                owner = st.text_input("账户名称", s.get("owner", "阿苏")); currency = st.selectbox("币种", ["USD", "CNY"], index=0 if s.get("currency", "USD") == "USD" else 1)
                start_cash = st.number_input("初始现金", value=fnum(s.get("start_cash")), step=1.0, format="%.2f"); start_savings = st.number_input("初始储蓄", value=fnum(s.get("start_savings")), step=1.0, format="%.2f")
            with c2:
                base_score = st.number_input("基础信用分", min_value=0, max_value=850, value=inum(s.get("base_score"), 720)); cash_floor = st.number_input("现金安全线", value=fnum(s.get("cash_floor")), step=1.0, format="%.2f")
                approval_threshold = st.number_input("家长审批金额线", value=fnum(s.get("approval_threshold")), step=1.0, format="%.2f"); sheet_sync = st.checkbox("保存时自动刷新 Google Sheet 展示页（不建议，较慢）", value=bool(s.get("sheet_report_sync_on_save", False)))
            submit = st.form_submit_button("保存账户设置", type="primary", disabled=not can_parent())
        if submit:
            s.update({"owner": owner, "currency": currency, "start_cash": start_cash, "start_savings": start_savings, "base_score": base_score, "cash_floor": cash_floor, "approval_threshold": approval_threshold, "sheet_report_sync_on_save": sheet_sync})
            commit(data, event="保存账户设置", reason="settings"); st.rerun()
    with tab2:
        rules = data.setdefault("rules", {})
        with st.form("rules_form"):
            c1, c2, c3 = st.columns(3)
            with c1: budget_warning = st.number_input("预算预警线", min_value=0.0, max_value=2.0, value=fnum(rules.get("budget_warning_line"), 0.90), step=0.05); budget_reject = st.number_input("预算拒绝线", min_value=0.0, max_value=3.0, value=fnum(rules.get("budget_reject_line"), 1.20), step=0.05)
            with c2: score_reject = st.number_input("信用分拒绝线", min_value=300, max_value=850, value=inum(rules.get("score_reject_line"), 650)); spend_warning = st.number_input("支出收入比预警线", min_value=0.0, max_value=2.0, value=fnum(rules.get("spend_income_warning_line"), 0.85), step=0.05)
            with c3: loan_soft = st.number_input("贷款资产建议线", min_value=0.0, max_value=2.0, value=fnum(rules.get("loan_asset_soft_limit"), 0.45), step=0.05); loan_hard = st.number_input("贷款资产红线", min_value=0.0, max_value=2.0, value=fnum(rules.get("loan_asset_hard_limit"), 0.65), step=0.05)
            submit = st.form_submit_button("保存风控规则", type="primary", disabled=not can_parent())
        if submit:
            rules.update({"budget_warning_line": budget_warning, "budget_reject_line": budget_reject, "score_reject_line": score_reject, "spend_income_warning_line": spend_warning, "loan_asset_soft_limit": loan_soft, "loan_asset_hard_limit": loan_hard})
            commit(data, event="保存风控规则", reason="rules"); st.rerun()
    with tab3:
        st.subheader("Google Sheet 和备份")
        st.write(f"当前存储：{st.session_state.get('storage_backend', '未知')}")
        c1, c2, c3 = st.columns(3)
        with c1:
            if st.button("重新从存储读取"):
                reload_data(); st.rerun()
        with c2:
            if st.button("刷新核心展示页", disabled=not can_parent()):
                try: result = sync_readable_sheets(data, scope="core"); st.success(f"核心展示页已刷新：{result['sheet_count']} 张表。")
                except Exception as e: st.error(f"刷新失败：{e}")
            if st.button("完整刷新展示页（低频）", disabled=not can_parent(), help="会刷新更多 worksheet。若遇到 429，请等待 1 分钟后再试。"):
                try: result = sync_readable_sheets(data, scope="full"); st.success(f"完整展示页已刷新：{result['sheet_count']} 张表。")
                except Exception as e: st.error(f"刷新失败：{e}")
        with c3:
            confirm = st.checkbox("确认清空所有数据")
            if st.button("清空为空库", disabled=(not confirm or not can_parent())):
                reset_empty(); st.success("已清空。"); st.rerun()
        st.divider(); st.markdown("#### 备份恢复")
        backups = data.get("backups", [])
        if not backups: st.info("暂无备份。")
        else:
            rows = [{"备份ID": b.get("id"), "时间": b.get("time"), "操作人": b.get("operator"), "事件": b.get("event"), "原因": b.get("reason")} for b in sorted(backups, key=lambda x: x.get("time", ""), reverse=True)]
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
            selected = st.selectbox("选择要恢复的备份", [b.get("id") for b in reversed(backups)])
            if st.button("恢复选中备份", disabled=not can_parent()):
                backup = next((b for b in backups if b.get("id") == selected), None)
                if backup:
                    restored = normalize_data(json.loads(backup.get("snapshot_json") or "{}")); restored["backups"] = backups
                    commit(restored, event="恢复备份", reason=selected); st.success("已恢复。"); st.rerun()
    with tab4:
        st.subheader("数据导出")
        json_text = json.dumps(normalize_data(data), ensure_ascii=False, indent=2)
        st.download_button("下载完整 JSON 备份", json_text.encode("utf-8"), "asu_money3_backup.json", "application/json")
        excel_bytes = build_excel_export(data)
        if excel_bytes is not None:
            st.download_button("下载 Excel 总账包", excel_bytes, "asu_money3_private_bank.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        tx_df = pd.DataFrame(data.get("transactions", []))
        if not tx_df.empty: st.download_button("下载交易 CSV", tx_df.to_csv(index=False).encode("utf-8-sig"), "asu_money3_transactions.csv", "text/csv")
        st.markdown("#### 审计日志")
        rows = [{"时间": a.get("time"), "操作人": a.get("operator"), "事件": a.get("event"), "详情": a.get("detail")} for a in sorted(data.get("audit_log", []), key=lambda x: x.get("time", ""), reverse=True)[:100]]
        dataframe_or_empty(rows, empty_text="暂无审计日志。")




# ============================================================
# 12. 完整功能面板：保留 2.x 全部功能，并加入 3.0 新模块
# ============================================================

def transactions_df(data: Dict[str, Any]) -> pd.DataFrame:
    rows = []
    for tx in data.get("transactions", []):
        rows.append({
            "日期": tx.get("date"),
            "类型": tx.get("type"),
            "金额": fnum(tx.get("amount")),
            "分类": tx.get("category"),
            "对象/商户/借款人": tx.get("party"),
            "本金回收": fnum(tx.get("principal_repaid")),
            "利息收入": fnum(tx.get("interest_received")),
            "预计回款": fnum(tx.get("expected_repayment")),
            "到期日": tx.get("due_date"),
            "创建人": tx.get("created_by"),
            "备注": tx.get("memo"),
            "id": tx.get("id"),
        })
    if not rows:
        return pd.DataFrame()
    return pd.DataFrame(rows).sort_values(["日期", "id"], ascending=[False, False])


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
    bounty_score = 100 if len([b for b in data.get("bounties", []) if b.get("status") in {"已完成", "已支付"}]) >= 3 else 70 if data.get("bounties") else 40
    rows.append(["任务收入", bounty_score, 100, "赏金任务完成情况"])
    df = pd.DataFrame(rows, columns=["项目", "得分", "满分", "说明"])
    df["完成度"] = df["得分"] / df["满分"]
    return df


def page_dashboard_full(data: Dict[str, Any]) -> None:
    m = calc_financials(data); ccy = m["currency"]
    render_hero(data, m)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("总资产", money(m["total_assets"], ccy))
    c2.metric("净资产", money(m["net_assets"], ccy))
    c3.metric("现金", money(m["cash"], ccy))
    c4.metric("信用分", score_text(m["credit_score"]))
    c5, c6, c7, c8 = st.columns(4)
    c5.metric("应收贷款", money(m["receivables"], ccy))
    c6.metric("今日可安全花", money(m["safe_to_spend"], ccy))
    c7.metric("本月消费", money(m["month_expense"], ccy))
    c8.metric("支出/收入", percent(m["spend_income_ratio"]))
    left, right = st.columns([1, 1])
    with left:
        st.markdown("#### AI 本周行动计划")
        for i, a in enumerate(build_weekly_plan(data), 1):
            st.write(f"**{i}.** {a}")
        st.markdown("#### 信用分因子")
        for f in m["score_factors"]:
            st.write(f"- {f}")
    with right:
        render_asset_mix_chart(m)
        render_budget_usage_chart(m)
    st.markdown("#### 最近交易")
    df = transactions_df(data)
    if df.empty:
        st.info("暂无交易。")
    else:
        show = df.head(12).copy()
        show["金额"] = show["金额"].map(lambda x: money(x, ccy))
        st.dataframe(show.drop(columns=["id"], errors="ignore"), use_container_width=True, hide_index=True)


def page_add_transaction_full(data: Dict[str, Any]) -> None:
    st.subheader("新增交易")
    st.caption("阿苏、爸爸、妈妈都可以记账；审批和系统设置仍由爸爸/妈妈管理。")
    with st.form("full_add_tx_form"):
        c1, c2, c3 = st.columns(3)
        with c1:
            d = st.date_input("日期", value=date.today(), key="full_tx_date")
            tx_type = st.selectbox("类型", TX_TYPES, key="full_tx_type")
        with c2:
            amount = st.number_input("金额", min_value=0.0, value=5.0, step=1.0, format="%.2f", key="full_tx_amount")
            category = st.selectbox("分类", category_options(data), key="full_tx_category")
        with c3:
            party = st.text_input("对象/商户/借款人", "", key="full_tx_party")
            due = st.date_input("到期日/还款日", value=date.today() + timedelta(days=7), key="full_tx_due").isoformat()
        memo = st.text_input("备注", "", key="full_tx_memo")
        c4, c5, c6 = st.columns(3)
        with c4:
            expected = st.number_input("预计回款", min_value=0.0, value=0.0, step=1.0, format="%.2f", key="full_tx_expected")
        with c5:
            principal_repaid = st.number_input("还款本金", min_value=0.0, value=0.0, step=1.0, format="%.2f", key="full_tx_principal")
        with c6:
            interest_received = st.number_input("收到利息", min_value=0.0, value=0.0, step=1.0, format="%.2f", key="full_tx_interest")
        submitted = st.form_submit_button("保存交易", type="primary", disabled=not can_bookkeep())
    if submitted:
        if tx_type == "还款":
            amount = principal_repaid + interest_received if principal_repaid + interest_received > 0 else amount
        tx = make_tx(tx_type, amount, category, party=party, memo=memo, tx_date=d.isoformat(), expected_repayment=expected, principal_repaid=principal_repaid, interest_received=interest_received, due_date=due if tx_type == "放贷" else "")
        data.setdefault("transactions", []).append(tx)
        commit(data, event="新增交易", reason=f"{tx_type} {money(amount)} {category}")
        st.success("交易已保存。")
        st.rerun()


def page_ai_center_full(data: Dict[str, Any]) -> None:
    st.subheader("AI 决策中心")
    st.caption("自然语言输入 → AI 分类 → Python 硬规则审批 → 情景规划 → 可加入待审批。")
    text = st.text_area("输入请求", value="我想买 $18 的 Minecraft 道具", height=120, key="full_ai_text")
    if st.button("AI 分类 + 风控审批", type="primary", key="full_ai_eval"):
        decision = parse_natural_request(data, text)
        st.session_state["full_ai_decision"] = decision
        st.session_state["full_ai_text_saved"] = text
    decision = st.session_state.get("full_ai_decision")
    if decision:
        render_decision(decision)
        if decision.get("kind") == "消费审批":
            st.markdown("#### 情景规划")
            amount = fnum(decision.get("metrics", {}).get("购买金额"))
            cat = str(decision.get("metrics", {}).get("消费分类") or "其他")
            merchant = str(decision.get("metrics", {}).get("商户") or "")
            scenarios = scenario_planning(data, amount, cat, st.session_state.get("full_ai_text_saved", text), merchant)
            st.dataframe(pd.DataFrame(scenarios), use_container_width=True, hide_index=True)
        with st.expander("AI 客户经理报告", expanded=True):
            st.markdown(deepseek_decision_report(decision))
        if st.button("提交到待审批队列", key="full_ai_to_pending", disabled=decision.get("result") == "拒绝"):
            submit_pending_request(data, decision, st.session_state.get("full_ai_text_saved", text), applicant=current_operator())
            commit(data, event="提交待审批", reason=st.session_state.get("full_ai_text_saved", text))
            st.session_state.pop("full_ai_decision", None)
            st.rerun()


def page_purchase_full(data: Dict[str, Any]) -> None:
    st.subheader("消费审批")
    with st.form("full_purchase_form"):
        c1, c2, c3 = st.columns(3)
        amount = c1.number_input("购买金额", value=18.0, min_value=0.0, step=1.0, format="%.2f", key="full_purchase_amount")
        category = c2.selectbox("消费分类", category_options(data), key="full_purchase_cat")
        merchant = c3.text_input("商户", "Minecraft 商店", key="full_purchase_merchant")
        desc = st.text_area("购买说明", "Minecraft 道具", key="full_purchase_desc")
        submitted = st.form_submit_button("模拟消费审批", type="primary")
    if submitted:
        st.session_state["full_purchase_decision"] = evaluate_purchase_decision(data, amount, category, desc, merchant)
        st.session_state["full_purchase_text"] = desc
    decision = st.session_state.get("full_purchase_decision")
    if decision:
        render_decision(decision)
        st.markdown("#### 情景规划")
        scenarios = scenario_planning(data, fnum(decision["metrics"].get("购买金额")), str(decision["metrics"].get("消费分类") or "其他"), st.session_state.get("full_purchase_text", desc), str(decision["metrics"].get("商户") or ""))
        st.dataframe(pd.DataFrame(scenarios), use_container_width=True, hide_index=True)
        c1, c2 = st.columns(2)
        if c1.button("直接入账", key="full_purchase_book", disabled=(decision.get("result") == "拒绝" or not can_bookkeep())):
            data.setdefault("transactions", []).append(normalize_tx(decision["pending_tx"]))
            commit(data, event="消费直接入账", reason=decision.get("result", ""))
            st.rerun()
        if c2.button("提交到待审批", key="full_purchase_pending", disabled=decision.get("result") == "拒绝"):
            submit_pending_request(data, decision, st.session_state.get("full_purchase_text", desc), applicant=current_operator())
            commit(data, event="消费提交待审批", reason=st.session_state.get("full_purchase_text", desc))
            st.rerun()


def page_loan_full(data: Dict[str, Any]) -> None:
    st.subheader("贷款审批")
    with st.form("full_loan_form"):
        c1, c2, c3 = st.columns(3)
        principal = c1.number_input("拟借出本金", value=20.0, min_value=0.0, step=1.0, format="%.2f", key="full_loan_principal")
        borrower = c2.text_input("借款人", "爸爸", key="full_loan_borrower")
        expected = c3.number_input("预计回款总额", value=22.0, min_value=0.0, step=1.0, format="%.2f", key="full_loan_expected")
        due = st.date_input("预计还款日", value=date.today() + timedelta(days=7), key="full_loan_due")
        memo = st.text_area("备注", "家庭临时周转", key="full_loan_memo")
        submitted = st.form_submit_button("模拟贷款审批", type="primary")
    if submitted:
        st.session_state["full_loan_decision"] = evaluate_loan_decision(data, principal, borrower, expected, due.isoformat(), memo)
        st.session_state["full_loan_text"] = f"放贷给{borrower} {money(principal)}，预计回款 {money(expected)}，到期 {due.isoformat()}"
    decision = st.session_state.get("full_loan_decision")
    if decision:
        render_decision(decision)
        c1, c2 = st.columns(2)
        if c1.button("直接入账", key="full_loan_book", disabled=(decision.get("result") == "拒绝" or not can_bookkeep())):
            data.setdefault("transactions", []).append(normalize_tx(decision["pending_tx"]))
            commit(data, event="贷款直接入账", reason=decision.get("result", ""))
            st.rerun()
        if c2.button("提交到待审批", key="full_loan_pending", disabled=decision.get("result") == "拒绝"):
            submit_pending_request(data, decision, st.session_state.get("full_loan_text", "贷款审批"), applicant=current_operator())
            commit(data, event="贷款提交待审批", reason=st.session_state.get("full_loan_text", "贷款审批"))
            st.rerun()


def page_pending_full(data: Dict[str, Any]) -> None:
    st.subheader("待审批队列")
    pending = data.get("pending_requests", [])
    if not pending:
        st.info("暂无申请。")
        return
    status_filter = st.selectbox("状态筛选", ["全部", "待审批", "已批准", "已拒绝"], key="full_pending_filter")
    items = [r for r in pending if status_filter == "全部" or r.get("parent_status") == status_filter]
    for r in sorted(items, key=lambda x: str(x.get("created_at", "")), reverse=True):
        css = decision_class(r.get("parent_status", "待审批"))
        st.markdown(f"<div class='decision {css}'><div class='pill'>{r.get('parent_status')}</div><h3>{r.get('request_type')}：{money(r.get('amount'))} · {r.get('category')}</h3><p class='small'>{r.get('request_text')}</p><p class='small'>申请人：{r.get('applicant')}；系统结论：{r.get('python_decision')}；审批人：{r.get('approved_by') or '未审批'}</p></div>", unsafe_allow_html=True)
        if r.get("parent_status") == "待审批":
            note = st.text_input("家长备注", value=r.get("parent_note", ""), key=f"full_pending_note_{r['id']}")
            c1, c2 = st.columns(2)
            if c1.button("批准并入账", key=f"full_pending_approve_{r['id']}", disabled=not can_parent(), type="primary"):
                approve_pending_request(data, r, note)
                commit(data, event="批准申请", reason=r.get("request_text", ""))
                st.rerun()
            if c2.button("拒绝", key=f"full_pending_reject_{r['id']}", disabled=not can_parent()):
                reject_pending_request(r, note)
                commit(data, event="拒绝申请", reason=r.get("request_text", ""))
                st.rerun()


def page_risk_full(data: Dict[str, Any]) -> None:
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


def page_monthly_full(data: Dict[str, Any]) -> None:
    st.subheader("月度账单")
    df = transactions_df(data)
    if df.empty:
        st.info("暂无交易。")
        return
    months = sorted(df["日期"].astype(str).str[:7].unique(), reverse=True)
    selected = st.selectbox("选择月份", months, key="full_month_select")
    m = calc_financials(data, selected)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("本月收入", money(m["month_income"], m["currency"]))
    c2.metric("本月消费", money(m["month_expense"], m["currency"]))
    c3.metric("支出/收入比", percent(m["spend_income_ratio"]))
    c4.metric("信用分", score_text(m["credit_score"]))
    rows = [tx for tx in data.get("transactions", []) if tx_month(tx) == selected]
    render_monthly_spending_chart(rows, m["currency"])
    month_df = df[df["日期"].astype(str).str[:7] == selected].copy()
    if not month_df.empty:
        month_df["金额"] = month_df["金额"].map(lambda x: money(x, m["currency"]))
        st.dataframe(month_df.drop(columns=["id"], errors="ignore"), use_container_width=True, hide_index=True)


def page_budget_full(data: Dict[str, Any]) -> None:
    st.subheader("预算管理")
    m = calc_financials(data)
    rows = [{"分类": cat, "预算": money(row["limit"], m["currency"]), "已花": money(row["spent"], m["currency"]), "剩余": money(row.get("remaining"), m["currency"]), "使用率": percent(row["ratio"])} for cat, row in m["budget_usage"].items()]
    dataframe_or_empty(rows, empty_text="暂无预算。")
    render_budget_usage_chart(m)
    with st.form("full_budget_form"):
        c1, c2 = st.columns(2)
        cat = c1.text_input("分类名称", "游戏", key="full_budget_cat")
        limit = c2.number_input("月度预算", min_value=0.0, value=20.0, step=1.0, format="%.2f", key="full_budget_limit")
        submit = st.form_submit_button("保存/更新预算", type="primary", disabled=not can_parent())
    if submit:
        data.setdefault("budgets", {})[cat.strip() or "其他"] = limit
        commit(data, event="保存预算", reason=f"{cat}={limit}")
        st.rerun()
    if data.get("budgets") and can_parent():
        del_cat = st.selectbox("删除预算分类", list(data.get("budgets", {}).keys()), key="full_budget_del_cat")
        if st.button("删除该预算分类", key="full_budget_delete"):
            data.get("budgets", {}).pop(del_cat, None)
            commit(data, event="删除预算", reason=del_cat)
            st.rerun()


def page_goals_full(data: Dict[str, Any]) -> None:
    st.subheader("储蓄目标")
    st.caption("家长可以修改、归档或删除目标；已完成/已归档目标不会在阿苏首页继续显示，但仍保留在历史记录里。")
    m = calc_financials(data)
    active = active_goals_from_metrics(m)
    archived = archived_goals_from_metrics(m)

    st.markdown("#### 进行中的目标")
    if not active:
        st.info("暂无进行中的储蓄目标。")
    for g in active:
        gid = g.get("id") or uid()
        with st.expander(f"{g.get('name')} · {money(g.get('current'), m['currency'])} / {money(g.get('target'), m['currency'])}", expanded=False):
            st.progress(min(1.0, fnum(g.get("progress"))))
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("目标", money(g.get("target"), m["currency"]))
            c2.metric("当前", money(g.get("current"), m["currency"]))
            c3.metric("完成度", percent(g.get("progress")))
            c4.metric("剩余", money(g.get("remaining"), m["currency"]))
            st.caption(f"截止日：{g.get('deadline') or '未设置'}；分类：{g.get('category')}；状态：{g.get('status')}")

            with st.form(f"edit_goal_{gid}"):
                ec1, ec2, ec3 = st.columns(3)
                new_name = ec1.text_input("目标名称", g.get("name", ""), key=f"goal_name_{gid}")
                new_target = ec2.number_input("目标金额", min_value=0.0, value=fnum(g.get("target")), step=5.0, format="%.2f", key=f"goal_target_{gid}")
                new_current = ec3.number_input("当前金额", min_value=0.0, value=fnum(g.get("current")), step=5.0, format="%.2f", key=f"goal_current_{gid}")
                new_category = ec1.text_input("分类", g.get("category", "长期储蓄"), key=f"goal_category_{gid}")
                old_deadline = safe_date(g.get("deadline")) or date.today() + timedelta(days=90)
                new_deadline = ec2.date_input("截止日", value=old_deadline, key=f"goal_deadline_{gid}").isoformat()
                new_note = ec3.text_input("说明", g.get("note", ""), key=f"goal_note_{gid}")
                save = st.form_submit_button("保存修改", disabled=not can_parent())
            if save:
                g.update({"name": new_name.strip() or "未命名目标", "target": new_target, "current": new_current, "category": new_category.strip() or "长期储蓄", "deadline": new_deadline, "note": new_note, "status": "已完成" if new_target > 0 and new_current >= new_target else "进行中"})
                commit(data, event="修改储蓄目标", reason=g.get("name", ""))
                st.rerun()

            b1, b2, b3 = st.columns(3)
            if b1.button("标记完成并归档", key=f"goal_archive_{gid}", disabled=not can_parent()):
                g["status"] = "已完成"
                g["archived_at"] = now_str()
                commit(data, event="完成储蓄目标", reason=g.get("name", ""))
                st.rerun()
            if b2.button("仅归档", key=f"goal_only_archive_{gid}", disabled=not can_parent()):
                g["status"] = "已归档"
                g["archived_at"] = now_str()
                commit(data, event="归档储蓄目标", reason=g.get("name", ""))
                st.rerun()
            if b3.button("删除目标", key=f"goal_delete_{gid}", disabled=not can_parent()):
                data["goals"] = [x for x in data.get("goals", []) if x.get("id") != gid]
                commit(data, event="删除储蓄目标", reason=g.get("name", ""))
                st.rerun()

    st.divider()
    st.markdown("#### 新增储蓄目标")
    with st.form("full_goal_form"):
        c1, c2, c3 = st.columns(3)
        name = c1.text_input("目标名称", "猫咪基金", key="full_goal_name")
        target = c2.number_input("目标金额", min_value=0.0, value=100.0, step=5.0, format="%.2f", key="full_goal_target")
        current = c3.number_input("当前金额", min_value=0.0, value=0.0, step=5.0, format="%.2f", key="full_goal_current")
        category = c1.text_input("分类", "长期储蓄", key="full_goal_category")
        deadline = c2.date_input("截止日", value=date.today() + timedelta(days=90), key="full_goal_deadline").isoformat()
        note = c3.text_input("说明", "", key="full_goal_note")
        submit = st.form_submit_button("新增储蓄目标", disabled=not can_parent())
    if submit:
        data.setdefault("goals", []).append({"id": uid(), "name": name.strip() or "未命名目标", "target": target, "current": current, "deadline": deadline, "category": category, "note": note, "status": "已完成" if target > 0 and current >= target else "进行中", "archived_at": ""})
        commit(data, event="新增储蓄目标", reason=name)
        st.rerun()

    if archived:
        st.divider()
        with st.expander(f"已完成/已归档目标历史（{len(archived)}）", expanded=False):
            rows = [{"目标": g.get("name"), "状态": g.get("status"), "目标金额": money(g.get("target"), m["currency"]), "当前金额": money(g.get("current"), m["currency"]), "分类": g.get("category"), "截止日": g.get("deadline"), "归档时间": g.get("archived_at")} for g in archived]
            dataframe_or_empty(rows)
            if can_parent():
                options = [f"{g.get('name')} · {g.get('id')}" for g in archived]
                selected = st.selectbox("恢复一个历史目标", options, key="restore_goal_select")
                selected_id = selected.split(" · ")[-1]
                if st.button("恢复为进行中", key="restore_goal_btn"):
                    for g in data.get("goals", []):
                        if g.get("id") == selected_id:
                            g["status"] = "进行中"
                            g["archived_at"] = ""
                            commit(data, event="恢复储蓄目标", reason=g.get("name", ""))
                            st.rerun()


def page_score_full(data: Dict[str, Any]) -> None:
    st.subheader("信用分")
    m = calc_financials(data)
    st.metric("当前信用分", score_text(m["credit_score"]))
    st.caption("家庭内部风控评分，不是 FICO，不是银行真实征信。")
    st.markdown("#### 当前因子")
    for f in m["score_factors"]:
        st.write(f"- {f}")
    st.markdown("#### 信用评分卡")
    card = build_scorecard(data)
    st.dataframe(card, use_container_width=True, hide_index=True)
    if plotly_enabled():
        fig = px.bar(card, x="项目", y="完成度", title="信用评分卡完成度")
        st.plotly_chart(fig, use_container_width=True, key=next_chart_key("plotly"))
    else:
        st.bar_chart(card.set_index("项目")[["完成度"]])
    st.markdown("#### 信用分历史")
    render_credit_score_chart(data.get("score_history", []))
    dataframe_or_empty(sorted(data.get("score_history", []), key=lambda x: str(x.get("time", "")), reverse=True), empty_text="暂无信用分历史。")


def page_merchants_full(data: Dict[str, Any]) -> None:
    st.subheader("商户权益")
    if not data.get("merchants"):
        st.info("暂无商户权益。")
    for merchant in data.get("merchants", []):
        r = evaluate_merchant_access(data, merchant)
        css = decision_class(r["result"])
        failed = "无" if not r["failed"] else "；".join(r["failed"])
        st.markdown(f"<div class='decision {css}'><div class='pill'>{merchant.get('category')}</div><h3>{merchant.get('name')}</h3><p><b>{r['message']}</b></p><p class='small'>暂缓/失败原因：{failed}</p><p class='small'>{merchant.get('note','')}</p></div>", unsafe_allow_html=True)
    with st.form("full_merchant_form"):
        c1, c2, c3 = st.columns(3)
        name = c1.text_input("商户名称", "书店奖励", key="full_merchant_name")
        category = c1.selectbox("商户分类", category_options(data), key="full_merchant_category")
        required = c2.number_input("最低信用分", min_value=0, max_value=850, value=720, key="full_merchant_score")
        discount = c2.number_input("折扣比例", min_value=0.0, max_value=0.9, value=0.08, step=0.01, key="full_merchant_discount")
        cap = c3.number_input("品类预算开放上限", min_value=0.0, max_value=2.0, value=0.90, step=0.05, key="full_merchant_cap")
        note = c3.text_input("说明", "", key="full_merchant_note")
        submit = st.form_submit_button("新增商户", type="primary", disabled=not can_parent())
    if submit:
        data.setdefault("merchants", []).append({"id": uid(), "name": name.strip() or "未命名商户", "category": category, "discount": discount, "required_score": int(required), "category_budget_cap": cap, "note": note})
        commit(data, event="新增商户", reason=name)
        st.rerun()


def page_weekly_full(data: Dict[str, Any]) -> None:
    st.subheader("自动周报")
    if st.button("生成本周周报", type="primary", disabled=not can_parent(), key="full_weekly_generate"):
        report = generate_weekly_report(data)
        data.setdefault("weekly_reports", []).append(report)
        data["weekly_reports"] = data["weekly_reports"][-52:]
        commit(data, event="生成周报", reason=report["week_start"])
        st.rerun()
    if not data.get("weekly_reports"):
        st.info("暂无周报。")
    for r in sorted(data.get("weekly_reports", []), key=lambda x: x.get("week_start", ""), reverse=True):
        with st.expander(f"{r.get('week_start')} 周报", expanded=False):
            st.text(r.get("content", ""))


def page_rewards_full(data: Dict[str, Any]) -> None:
    st.subheader("奖励与徽章")
    evaluate_badges(data)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### 徽章")
        rows = [{"徽章": b.get("name"), "获得时间": b.get("earned_at"), "说明": b.get("description")} for b in data.get("badges", [])]
        dataframe_or_empty(rows, empty_text="暂无徽章。")
    with c2:
        st.markdown("#### 奖励券")
        rows = [{"奖励": r.get("title"), "创建时间": r.get("created_at"), "创建人": r.get("created_by"), "状态": r.get("status"), "说明": r.get("note")} for r in data.get("rewards", [])]
        dataframe_or_empty(rows, empty_text="暂无奖励券。")
    st.divider()
    with st.form("full_reward_form"):
        title = st.text_input("奖励名称", "周末小额自由消费券", key="full_reward_title")
        note = st.text_input("说明", "用于奖励预算纪律或储蓄进度", key="full_reward_note")
        submitted = st.form_submit_button("发放奖励", type="primary", disabled=not can_parent())
    if submitted and title.strip():
        data.setdefault("rewards", []).append({"id": uid(), "title": title.strip(), "created_at": now_str(), "created_by": current_operator(), "status": "可用", "note": note})
        commit(data, event="发放奖励", reason=title)
        st.rerun()
    available = [r for r in data.get("rewards", []) if r.get("status") == "可用"]
    if available and can_parent():
        opt = st.selectbox("标记奖励券", [f"{r.get('title')} · {r.get('id')}" for r in available], key="full_reward_mark")
        selected_id = opt.split(" · ")[-1]
        c1, c2 = st.columns(2)
        if c1.button("标记为已使用", key="full_reward_used"):
            for r in data.get("rewards", []):
                if r.get("id") == selected_id:
                    r["status"] = "已使用"
            commit(data, event="使用奖励券", reason=selected_id)
            st.rerun()
        if c2.button("作废奖励券", key="full_reward_void"):
            for r in data.get("rewards", []):
                if r.get("id") == selected_id:
                    r["status"] = "已作废"
            commit(data, event="作废奖励券", reason=selected_id)
            st.rerun()


def page_rules_full(data: Dict[str, Any]) -> None:
    st.subheader("风控规则")
    rules = data.setdefault("rules", empty_data()["rules"])
    rows = [{"规则": k, "当前值": v} for k, v in rules.items()]
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    if not can_parent():
        st.warning("只有爸爸/妈妈可以修改风控规则。")
        return
    with st.form("full_rules_form"):
        new_rules = {}
        labels = {
            "budget_warning_line": "预算预警线",
            "budget_reject_line": "预算拒绝线",
            "score_reject_line": "信用分拒绝线",
            "loan_asset_soft_limit": "贷款资产建议线",
            "loan_asset_hard_limit": "贷款资产红线",
            "debt_asset_warning_line": "负债资产预警线",
            "spend_income_warning_line": "支出收入比预警线",
            "weekly_no_impulse_reward": "周度无冲动消费奖励分",
            "monthly_budget_reward": "月度预算纪律奖励分",
        }
        for k, default in empty_data()["rules"].items():
            new_rules[k] = st.number_input(labels.get(k, k), value=fnum(rules.get(k, default)), step=0.05 if "line" in k or "limit" in k else 1.0, format="%.2f", key=f"full_rule_{k}")
        submitted = st.form_submit_button("保存规则", type="primary")
    if submitted:
        data["rules"] = new_rules
        commit(data, event="保存风控规则", reason="rules updated")
        st.rerun()


def page_backups_full(data: Dict[str, Any]) -> None:
    st.subheader("备份与恢复")
    backups = data.get("backups", [])
    if not backups:
        st.info("暂无备份。")
        return
    df = pd.DataFrame([{k: b.get(k) for k in ["id", "time", "operator", "event", "reason"]} for b in backups]).sort_values("time", ascending=False)
    st.dataframe(df, use_container_width=True, hide_index=True)
    if not can_parent():
        st.warning("只有爸爸/妈妈可以恢复备份。")
        return
    options = [f"{b.get('time')} · {b.get('operator')} · {b.get('event')} · {b.get('id')}" for b in sorted(backups, key=lambda x: str(x.get('time','')), reverse=True)]
    selected = st.selectbox("选择要恢复的备份", options, key="full_backup_select")
    selected_id = selected.split(" · ")[-1]
    if st.button("恢复到这个版本", type="primary", key="full_backup_restore"):
        backup = next((b for b in backups if b.get("id") == selected_id), None)
        if backup:
            restored = normalize_data(json.loads(backup.get("snapshot_json") or "{}"))
            restored["backups"] = backups
            commit(restored, event="恢复备份", reason=selected_id)
            st.rerun()


def page_transactions_full(data: Dict[str, Any]) -> None:
    st.subheader("交易流水")
    df = transactions_df(data)
    if df.empty:
        st.info("暂无交易。")
    else:
        show = df.copy()
        ccy = calc_financials(data)["currency"]
        show["金额"] = show["金额"].map(lambda x: money(x, ccy))
        st.dataframe(show.drop(columns=["id"], errors="ignore"), use_container_width=True, hide_index=True)
        st.download_button("下载交易 CSV", df.drop(columns=["id"], errors="ignore").to_csv(index=False).encode("utf-8-sig"), "asu_money3_transactions.csv", "text/csv", key="full_tx_csv")
        st.divider()
        st.markdown("#### 删除单笔交易")
        if can_bookkeep():
            options = [f"{row['日期']} · {row['类型']} · {money(row['金额'], ccy)} · {row['分类']} · {row['id']}" for _, row in df.iterrows()]
            selected = st.selectbox("选择交易", options, key="full_tx_delete_select")
            selected_id = selected.split(" · ")[-1]
            confirm = st.checkbox("确认删除这笔交易", key="full_tx_delete_confirm")
            if st.button("删除单笔交易", disabled=not confirm, key="full_tx_delete_btn"):
                before = len(data.get("transactions", []))
                data["transactions"] = [tx for tx in data.get("transactions", []) if tx.get("id") != selected_id]
                if len(data["transactions"]) < before:
                    commit(data, event="删除单笔交易", reason=selected_id)
                    st.success("已删除。")
                    st.rerun()
        else:
            st.info("当前身份没有记账权限。")
    json_text = json.dumps(normalize_data(data), ensure_ascii=False, indent=2)
    st.download_button("下载完整 JSON 备份", json_text.encode("utf-8"), "asu_money3_backup.json", "application/json", key="full_json_backup")


def page_bounties_full(data: Dict[str, Any]) -> None:
    st.subheader("赏金任务")
    st.caption("默认只显示未完成任务；已支付、已取消、已归档的任务会进入历史记录，不再打扰主流程。")
    tab1, tab2, tab3, tab4 = st.tabs(["当前任务", "发布任务", "审核/发放", "历史记录"])

    with tab1:
        active = sorted(active_bounties(data), key=lambda x: str(x.get("created_at", "")), reverse=True)
        if not active:
            st.info("暂无当前任务。已完成任务在历史记录里。")
        for b in active:
            bid = b.get("id") or uid()
            with st.expander(f"{b.get('title')} · {b.get('status')} · {money(b.get('reward_amount'))}", expanded=False):
                st.write(b.get("description") or "无说明")
                st.caption(f"积分：{b.get('reward_points')}；难度：{b.get('difficulty')}；截止日：{b.get('deadline') or '未设置'}；领取人：{b.get('assigned_to') or '未领取'}")
                if can_parent():
                    with st.form(f"edit_bounty_{bid}"):
                        c1, c2, c3 = st.columns(3)
                        title = c1.text_input("任务标题", b.get("title", ""), key=f"edit_bounty_title_{bid}")
                        reward_amount = c2.number_input("赏金金额", min_value=0.0, value=fnum(b.get("reward_amount")), step=1.0, format="%.2f", key=f"edit_bounty_amount_{bid}")
                        reward_points = c3.number_input("信用积分", min_value=0, value=inum(b.get("reward_points"), 0), step=1, key=f"edit_bounty_points_{bid}")
                        desc = st.text_area("任务说明", b.get("description", ""), key=f"edit_bounty_desc_{bid}")
                        category = c1.text_input("任务分类", b.get("category", "家庭任务"), key=f"edit_bounty_category_{bid}")
                        difficulty = c2.selectbox("难度", ["简单", "普通", "困难"], index=["简单", "普通", "困难"].index(b.get("difficulty")) if b.get("difficulty") in ["简单", "普通", "困难"] else 1, key=f"edit_bounty_difficulty_{bid}")
                        old_deadline = safe_date(b.get("deadline")) or date.today() + timedelta(days=7)
                        deadline = c3.date_input("截止日", value=old_deadline, key=f"edit_bounty_deadline_{bid}").isoformat()
                        saved = st.form_submit_button("保存修改")
                    if saved:
                        b.update({"title": title.strip() or "未命名任务", "description": desc.strip(), "reward_amount": reward_amount, "reward_points": int(reward_points), "category": category.strip() or "家庭任务", "difficulty": difficulty, "deadline": deadline})
                        commit(data, event="修改赏金任务", reason=b.get("title", ""))
                        st.rerun()
                    c1, c2 = st.columns(2)
                    if c1.button("取消并归档", key=f"cancel_bounty_{bid}"):
                        b["status"] = "已取消"
                        b["reviewed_by"] = current_operator()
                        b["reviewed_at"] = now_str()
                        commit(data, event="取消赏金任务", reason=b.get("title", ""))
                        st.rerun()
                    if c2.button("删除未完成任务", key=f"delete_bounty_{bid}"):
                        data["bounties"] = [x for x in data.get("bounties", []) if x.get("id") != bid]
                        commit(data, event="删除赏金任务", reason=b.get("title", ""))
                        st.rerun()

    with tab2:
        with st.form("full_bounty_form"):
            title = st.text_input("任务标题", "整理书桌并拍照提交", key="full_bounty_title")
            desc = st.text_area("任务说明", "把书桌清理干净，书本分类，垃圾扔掉，最后拍照或文字说明。", key="full_bounty_desc")
            c1, c2, c3 = st.columns(3)
            reward_amount = c1.number_input("赏金金额", min_value=0.0, value=5.0, step=1.0, format="%.2f", key="full_bounty_amount")
            reward_points = c2.number_input("信用积分", min_value=0, value=5, step=1, key="full_bounty_points")
            difficulty = c3.selectbox("难度", ["简单", "普通", "困难"], key="full_bounty_difficulty")
            category = c1.text_input("任务分类", "家庭任务", key="full_bounty_category")
            deadline = c2.date_input("截止日", value=date.today() + timedelta(days=7), key="full_bounty_deadline").isoformat()
            submit = st.form_submit_button("发布任务", type="primary", disabled=not can_parent())
        if submit:
            data.setdefault("bounties", []).append({"id": uid(), "title": title.strip() or "未命名任务", "description": desc.strip(), "reward_amount": reward_amount, "reward_points": int(reward_points), "category": category.strip() or "家庭任务", "difficulty": difficulty, "deadline": deadline, "status": "开放", "created_by": current_operator(), "created_at": now_str(), "assigned_to": "", "claimed_at": "", "submitted_at": "", "submission_note": "", "reviewed_by": "", "reviewed_at": "", "parent_note": "", "paid_tx_id": ""})
            commit(data, event="发布赏金任务", reason=title)
            st.rerun()

    with tab3:
        submitted = [b for b in data.get("bounties", []) if b.get("status") == "已提交"]
        if not submitted:
            st.info("暂无待审核任务。")
        for b in submitted:
            bid = b.get("id") or uid()
            st.markdown(f"<div class='bounty'><h4>{b.get('title')}</h4><p class='small'>{b.get('description')}</p><p>提交说明：{b.get('submission_note') or '未填写'}</p><div class='pill'>赏金 {money(b.get('reward_amount'))} · 积分 {b.get('reward_points')}</div></div>", unsafe_allow_html=True)
            note = st.text_input("审核备注", key=f"full_bounty_review_note_{bid}")
            c1, c2 = st.columns(2)
            if c1.button("通过并发放赏金", key=f"full_bounty_pay_{bid}", disabled=not can_parent(), type="primary"):
                b["status"] = "已支付"; b["reviewed_by"] = current_operator(); b["reviewed_at"] = now_str(); b["parent_note"] = note
                tx = make_tx("收入", fnum(b.get("reward_amount")), "赏金任务", party=current_operator(), memo=f"赏金任务：{b.get('title')}", bounty_id=b.get("id"))
                b["paid_tx_id"] = tx["id"]
                data.setdefault("transactions", []).append(tx)
                commit(data, event="发放赏金", reason=b.get("title", ""))
                st.rerun()
            if c2.button("退回修改", key=f"full_bounty_return_{bid}", disabled=not can_parent()):
                b["status"] = "已退回"; b["reviewed_by"] = current_operator(); b["reviewed_at"] = now_str(); b["parent_note"] = note
                commit(data, event="退回赏金任务", reason=b.get("title", ""))
                st.rerun()

    with tab4:
        history = sorted(historical_bounties(data), key=lambda x: str(x.get("reviewed_at") or x.get("created_at") or ""), reverse=True)
        rows = [{"标题": b.get("title"), "状态": b.get("status"), "赏金": money(b.get("reward_amount")), "积分": b.get("reward_points"), "截止日": b.get("deadline"), "发布人": b.get("created_by"), "领取人": b.get("assigned_to"), "审核人": b.get("reviewed_by"), "审核时间": b.get("reviewed_at")} for b in history]
        dataframe_or_empty(rows, empty_text="暂无历史任务。")
        st.caption("已支付任务已经转成收入交易，默认不建议删除；如需纠错，请到交易流水删除对应收入，再重发任务。")


def page_google_sheet_full(data: Dict[str, Any]) -> None:
    st.subheader("Google Sheet 导入/导出")
    st.write(f"当前存储：{st.session_state.get('storage_backend', '未知')}")
    st.caption("主数据从 state 工作表导入/保存；summary、transactions 等是展示页，手动刷新。")
    c1, c2, c3 = st.columns(3)
    if c1.button("重新从 Google Sheet / 本地读取", key="full_reload_storage"):
        reload_data(); st.rerun()
    if c2.button("刷新核心展示页", key="full_sync_sheets", disabled=not can_parent()):
        try:
            result = sync_readable_sheets(data, scope="core")
            st.success(f"核心展示页已刷新：{result['sheet_count']} 张表。")
        except Exception as e:
            st.error(f"刷新失败：{e}")
    if c2.button("完整刷新展示页（低频）", key="full_sync_sheets_all", disabled=not can_parent(), help="会刷新更多 worksheet。若遇到 429，请等待 1 分钟后再试。"):
        try:
            result = sync_readable_sheets(data, scope="full")
            st.success(f"完整展示页已刷新：{result['sheet_count']} 张表。")
        except Exception as e:
            st.error(f"刷新失败：{e}")
    if c3.button("保存当前主数据到存储", key="full_force_save", disabled=not can_parent()):
        save_data(data)
        st.success("已保存当前 state。")
    st.markdown("#### JSON 导入")
    uploaded = st.file_uploader("上传 asu_money3_backup.json", type=["json"], key="full_json_import")
    if uploaded is not None:
        try:
            raw = json.loads(uploaded.read().decode("utf-8"))
            imported = normalize_data(raw)
            st.success("JSON 可读取。确认后会覆盖当前数据。")
            if st.button("确认导入并覆盖", key="full_json_import_confirm", disabled=not can_parent()):
                commit(imported, event="导入 JSON 备份", reason="uploaded json")
                st.rerun()
        except Exception as e:
            st.error(f"JSON 读取失败：{e}")


def page_data_export_full(data: Dict[str, Any]) -> None:
    st.subheader("数据导出")
    json_text = json.dumps(normalize_data(data), ensure_ascii=False, indent=2)
    st.download_button("下载完整 JSON 备份", json_text.encode("utf-8"), "asu_money3_backup.json", "application/json", key="full_export_json")
    excel_bytes = build_excel_export(data)
    if excel_bytes is not None:
        st.download_button("下载 Excel 总账包", excel_bytes, "asu_money3_private_bank.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="full_export_excel")
    else:
        st.info("Excel 导出需要 openpyxl。")
    df = transactions_df(data)
    if not df.empty:
        st.download_button("下载交易 CSV", df.to_csv(index=False).encode("utf-8-sig"), "asu_money3_transactions.csv", "text/csv", key="full_export_tx_csv")
    if data.get("bounties"):
        bdf = pd.DataFrame(data.get("bounties", []))
        st.download_button("下载赏金任务 CSV", bdf.to_csv(index=False).encode("utf-8-sig"), "asu_money3_bounties.csv", "text/csv", key="full_export_bounty_csv")


def page_audit_full(data: Dict[str, Any]) -> None:
    st.subheader("审计日志")
    rows = [{"时间": a.get("time"), "操作人": a.get("operator"), "事件": a.get("event"), "详情": a.get("detail")} for a in sorted(data.get("audit_log", []), key=lambda x: x.get("time", ""), reverse=True)]
    dataframe_or_empty(rows, empty_text="暂无审计日志。")


def page_full_control_panel(data: Dict[str, Any]) -> None:
    # Legacy debug panel retained in code but no longer exposed in sidebar.
    st.title("完整功能面板")
    st.caption("保留原版银行功能：输入、输出、评估、审批、账务、报表、风控、备份、Google Sheet、奖励、赏金任务。")
    tabs = st.tabs([
        "1 儿童首页", "2 首页总览", "3 新增交易", "4 AI决策中心", "5 消费审批", "6 贷款审批",
        "7 待审批队列", "8 风险雷达", "9 月度账单", "10 预算管理", "11 储蓄目标", "12 信用分",
        "13 商户权益", "14 周报", "15 奖励徽章", "16 风控规则", "17 备份恢复", "18 交易流水",
        "19 赏金任务", "20 Google Sheet", "21 数据导出", "22 审计日志"
    ])
    with tabs[0]: page_asu_home(data)
    with tabs[1]: page_dashboard_full(data)
    with tabs[2]: page_add_transaction_full(data)
    with tabs[3]: page_ai_center_full(data)
    with tabs[4]: page_purchase_full(data)
    with tabs[5]: page_loan_full(data)
    with tabs[6]: page_pending_full(data)
    with tabs[7]: page_risk_full(data)
    with tabs[8]: page_monthly_full(data)
    with tabs[9]: page_budget_full(data)
    with tabs[10]: page_goals_full(data)
    with tabs[11]: page_score_full(data)
    with tabs[12]: page_merchants_full(data)
    with tabs[13]: page_weekly_full(data)
    with tabs[14]: page_rewards_full(data)
    with tabs[15]: page_rules_full(data)
    with tabs[16]: page_backups_full(data)
    with tabs[17]: page_transactions_full(data)
    with tabs[18]: page_bounties_full(data)
    with tabs[19]: page_google_sheet_full(data)
    with tabs[20]: page_data_export_full(data)
    with tabs[21]: page_audit_full(data)


# ============================================================
# 12. 主程序
# ============================================================

def main() -> None:
    st.set_page_config(page_title=APP_NAME, page_icon="🏦", layout="wide")
    inject_css(); require_login()
    data = get_data(); render_sidebar(data)
    # Reset per-run chart counter so Plotly widgets get deterministic unique IDs.
    st.session_state["_asu_plotly_chart_counter"] = 0

    mode = st.sidebar.radio(
        "页面",
        ["阿苏首页", "家长工作台", "私人银行后台", "系统设置"],
        index=0,
    )
    if mode == "阿苏首页": page_asu_home(data)
    elif mode == "家长工作台": page_parent_workspace(data)
    elif mode == "私人银行后台": page_bank_backend(data)
    else: page_settings(data)


if __name__ == "__main__":
    main()
