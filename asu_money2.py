# -*- coding: utf-8 -*-
"""
阿苏私人银行 3.0
儿童银行体验版 + Google Sheet 提速版 + 父母同权 + 赏金任务系统

运行：
    streamlit run asu_money3.py

基于 requirements.txt：
    streamlit
    pandas
    requests
    gspread
    google-auth

增强依赖（已安装时自动启用）：
    plotly      # 交互图表
    openpyxl    # Excel 导出
    numpy       # 趋势计算辅助
    matplotlib  # 预留，不作为主图表库

Streamlit Secrets 示例：
    SHEET_NAME = "asu_bank"
    DAD_PASSWORD = "可选"
    MOM_PASSWORD = "可选"
    ASU_PASSWORD = "可选"
    DEEPSEEK_API_KEY = "可选"
    DEEPSEEK_MODEL = "deepseek-chat"

    [gcp_service_account]
    type = "service_account"
    project_id = "xxx"
    private_key_id = "xxx"
    private_key = "-----BEGIN PRIVATE KEY-----\nxxx\n-----END PRIVATE KEY-----\n"
    client_email = "xxx@xxx.iam.gserviceaccount.com"
    client_id = "xxx"
    auth_uri = "https://accounts.google.com/o/oauth2/auth"
    token_uri = "https://oauth2.googleapis.com/token"
    auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
    client_x509_cert_url = "xxx"

3.0 设计原则：
1. 阿苏首页是主线：总资产、净资产、今日可安全花、信用分、目标、任务、申请。
2. 爸爸妈妈权限完全相同：审批、设置、赏金任务、后台交易都可操作。
3. Google Sheet 提速：日常保存只写 state，不自动刷新全部展示页；展示页手动刷新。
4. Python 负责硬规则和账务计算；AI 只负责分类解释和话术，不改变审批结论。
5. 私人银行业务不丢：消费审批、贷款审批、预算、目标、风险雷达、月报、周报、商户权益、备份。
"""

from __future__ import annotations

import copy
import json
import os
import re
import uuid
from io import BytesIO
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

try:
    import numpy as np
except Exception:
    np = None

try:
    import plotly.express as px
    import plotly.graph_objects as go
except Exception:
    px = None
    go = None

try:
    import openpyxl  # noqa: F401
except Exception:
    openpyxl = None

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


APP_NAME = "阿苏私人银行 3.0 增强版"
SCHEMA_VERSION = "asu-bank-3-0"
DATA_PATH = Path("asu_money3_data.json")
OLD_DATA_PATH = Path("asu_money2_data.json")
SESSION_KEY = "asu_bank_3_0_state"
STATE_WS = "state"
STATE_KEY = "bank_data"
CHUNK_SIZE = 45000

GSHEET_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

PARENTS = {"爸爸", "妈妈"}
CHILD = "阿苏"
OPERATORS = ["爸爸", "妈妈", "阿苏"]

DEFAULT_CATEGORIES = ["零花钱", "游戏", "甜品", "玩具", "学习", "宠物", "餐饮", "交通", "礼物", "家庭贷款", "赏金任务", "其他"]
TX_TYPES = ["收入", "消费", "转入储蓄", "储蓄取出", "放贷", "还款", "借入", "偿还负债"]


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
    return f"{sign}{currency} {v:,.2f}"


def percent(x: Any) -> str:
    return f"{fnum(x) * 100:.0f}%"


def tx_month(tx: Dict[str, Any]) -> str:
    return str(tx.get("date", ""))[:7]


def score_text(score: Optional[int]) -> str:
    return "未建立" if score is None else str(score)


def current_operator() -> str:
    return st.session_state.get("current_operator", "爸爸")


def can_parent() -> bool:
    return current_operator() in PARENTS


def can_child() -> bool:
    return current_operator() == CHILD


def role_label() -> str:
    return "家长权限" if can_parent() else "阿苏模式"


def decision_class(result: str) -> str:
    if result in {"批准", "通过", "开放", "已批准", "已入账", "已完成", "已支付", "可购买"}:
        return "ok"
    if result in {"延迟", "限额通过", "暂缓开放", "观察", "待家长审批", "待审批", "未建立", "已提交"}:
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


def safe_date(s: str) -> Optional[date]:
    try:
        return date.fromisoformat(str(s)[:10])
    except Exception:
        return None


# ============================================================
# 2. 数据模型和标准化
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
            "auto_approve_safe_income": True,
            "sheet_report_sync_on_save": False,
            "created_at": now_str(),
        },
        "rules": {
            "budget_warning_line": 0.90,
            "budget_reject_line": 1.20,
            "score_reject_line": 650,
            "loan_asset_soft_limit": 0.45,
            "loan_asset_hard_limit": 0.65,
            "debt_asset_warning_line": 0.25,
            "spend_income_warning_line": 0.85,
            "weekly_no_impulse_reward": 5.0,
            "monthly_budget_reward": 10.0,
        },
        "budgets": {
            "游戏": 20.0,
            "甜品": 15.0,
            "学习": 30.0,
        },
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
        "loan_id": tx.get("loan_id") or "",
        "bounty_id": tx.get("bounty_id") or "",
        "expected_repayment": fnum(tx.get("expected_repayment")),
        "principal_repaid": fnum(tx.get("principal_repaid")),
        "interest_received": fnum(tx.get("interest_received")),
        "due_date": tx.get("due_date") or "",
        "created_by": tx.get("created_by") or "",
        "created_at": tx.get("created_at") or now_str(),
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
    s["auto_approve_safe_income"] = bool(s.get("auto_approve_safe_income", True))
    s["sheet_report_sync_on_save"] = bool(s.get("sheet_report_sync_on_save", False))

    if isinstance(raw.get("rules"), dict):
        for k, v in raw["rules"].items():
            if k in base["rules"]:
                base["rules"][k] = fnum(v, base["rules"][k])
            else:
                base["rules"][str(k)] = fnum(v)

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
                "category": g.get("category") or "长期储蓄",
                "note": g.get("note") or "",
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

    # 控制历史长度，避免 Google Sheet state 过大。
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
    data.setdefault("audit_log", []).append({
        "time": now_str(),
        "operator": current_operator(),
        "event": event,
        "detail": detail,
    })
    data["audit_log"] = data["audit_log"][-500:]


# ============================================================
# 3. 登录和权限
# ============================================================

def secret_value(key: str, default: str = "") -> str:
    try:
        if key in st.secrets:
            return str(st.secrets[key])
    except Exception:
        pass
    return os.getenv(key, default)


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
    st.caption("请选择身份并输入密码。爸爸和妈妈拥有相同家长权限。")
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


# ============================================================
# 4. Google Sheet 和本地存储
# ============================================================

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
    # 兼容 2.x：单 cell bank_data。
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
    values = ws.get_all_records()
    decoded = decode_state_rows(values)
    if decoded is None:
        data = empty_data()
        save_data_to_gsheet(data)
        return data
    return decoded


def save_data_to_gsheet(data: Dict[str, Any]) -> None:
    # 3.0 提速点：日常保存只写 state，不刷新 summary/transactions 等展示页。
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
                # 默认关闭。只有用户明确开启时，才同步全部可读展示页。
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


def commit(data: Dict[str, Any], event: str = "保存数据", reason: str = "") -> None:
    old_state = normalize_data(st.session_state.get(SESSION_KEY, data))
    old_score = calc_financials(old_state).get("credit_score")
    new_data = normalize_data(data)

    backup = {
        "id": uid(),
        "time": now_str(),
        "operator": current_operator(),
        "event": event,
        "reason": reason or "",
        "snapshot_json": json.dumps(strip_snapshot(old_state), ensure_ascii=False),
    }
    new_data["backups"] = (old_state.get("backups", []) + [backup])[-30:]
    add_audit(new_data, event, reason)

    new_score = calc_financials(new_data).get("credit_score")
    if new_score is not None:
        change = None if old_score is None else new_score - old_score
        new_data.setdefault("score_history", []).append({
            "time": now_str(),
            "score": new_score,
            "change": change,
            "event": event,
            "reason": reason or event,
        })
        new_data["score_history"] = new_data["score_history"][-300:]

    evaluate_badges(new_data)
    st.session_state[SESSION_KEY] = new_data
    save_data(new_data)


# ============================================================
# 5. 账务计算
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
        "loan_id": loan_id,
        "bounty_id": bounty_id,
        "created_by": current_operator(),
        "created_at": now_str(),
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
            target = None
            if tx.get("loan_id"):
                target = by_id.get(str(tx.get("loan_id")))
            candidates: List[Dict[str, Any]] = []
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
        return None, ["账户尚未建立信用分。请在设置里填写基础信用分。"]

    score = float(base)
    factors: List[str] = []
    floor = fnum(s.get("cash_floor"))

    if cash < 0:
        score -= 55
        factors.append("现金余额为负：-55")
    elif floor > 0 and cash < floor:
        score -= 18
        factors.append(f"现金低于安全线 {money(floor)}：-18")
    elif floor > 0 and cash >= floor * 2:
        score += 6
        factors.append("现金缓冲较充足：+6")

    if total_assets > 0:
        savings_ratio = savings / total_assets
        if savings_ratio >= 0.35:
            score += 10
            factors.append("储蓄占总资产 35% 以上：+10")
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

    overspent = sum(1 for r in budget_usage.values() if r["limit"] > 0 and r["ratio"] > 1)
    near = sum(1 for r in budget_usage.values() if r["limit"] > 0 and 0.85 < r["ratio"] <= 1)
    if overspent:
        p = min(32, overspent * 10)
        score -= p
        factors.append(f"{overspent} 个预算品类已超支：-{p}")
    if near:
        p = min(15, near * 4)
        score -= p
        factors.append(f"{near} 个预算品类接近上限：-{p}")
    if overdue > 0:
        score -= 24
        factors.append(f"存在逾期未收回应收本金 {money(overdue)}：-24")

    approved_bounties = sum(1 for b in data.get("bounties", []) if b.get("status") in {"已完成", "已支付"})
    if approved_bounties >= 3:
        score += min(12, approved_bounties * 2)
        factors.append(f"完成赏金任务 {approved_bounties} 个：+{min(12, approved_bounties * 2)}")

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

    for tx in sorted(data.get("transactions", []), key=lambda x: (str(x.get("date", "")), str(x.get("created_at", "")), str(x.get("id", "")))):
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
        "safe_to_spend": safe_to_spend,
        "safe_spend_cash": safe_spend_cash,
        "budget_remaining": budget_remaining,
        "budget_usage": usage,
        "goals": goals,
        "loans": loans,
        "overdue_principal": overdue,
        "credit_score": score,
        "score_factors": factors,
    }


def has_activity(data: Dict[str, Any]) -> bool:
    s = data.get("settings", {})
    return bool(
        fnum(s.get("start_cash"))
        or fnum(s.get("start_savings"))
        or data.get("transactions")
        or data.get("budgets")
        or data.get("goals")
        or data.get("pending_requests")
        or data.get("bounties")
    )


def category_options(data: Dict[str, Any]) -> List[str]:
    return sorted(set(DEFAULT_CATEGORIES + list(data.get("budgets", {}).keys())))


# ============================================================
# 6. AI 分类、审批和情景评估
# ============================================================

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
        "甜品": ["甜品", "奶茶", "冰淇淋", "蛋糕", "糖", "饮料", "boba", "milk tea"],
        "玩具": ["玩具", "lego", "乐高", "手办", "娃娃", "卡牌"],
        "学习": ["书", "学习", "课程", "文具", "作业", "训练", "练习册", "notebook"],
        "宠物": ["猫", "猫粮", "猫砂", "宠物", "罐头", "rawz"],
        "餐饮": ["饭", "餐", "披萨", "pizza", "汉堡", "麦当劳", "burger", "lunch"],
        "交通": ["车", "uber", "公交", "地铁", "汽油", "停车"],
    }
    category = "其他"
    for cat, keys in rules.items():
        if any(k in s for k in keys):
            category = cat
            break
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

可用分类：
游戏、甜品、玩具、学习、宠物、餐饮、交通、礼物、家庭贷款、其他

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
            "savings": after["savings"] - before["savings"],
            "receivables": after["receivables"] - before["receivables"],
            "total_assets": after["total_assets"] - before["total_assets"],
            "net_assets": after["net_assets"] - before["net_assets"],
            "credit_score": delta_score,
        },
        "tx": tx,
    }


def usage_for(metrics: Dict[str, Any], cat: str) -> Dict[str, float]:
    return metrics.get("budget_usage", {}).get(cat, {"limit": 0.0, "spent": 0.0, "ratio": 0.0, "remaining": 0.0})


def rule_value(data: Dict[str, Any], key: str, default: float) -> float:
    return fnum((data.get("rules") or {}).get(key), default)


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
        result = "拒绝"
        reasons.append("购买金额必须大于 0。")
    if after["cash"] < 0:
        result = "拒绝"
        reasons.append("购买后现金余额为负，触发硬性拒绝。")
    if budget["limit"] > 0 and budget["ratio"] > rule_value(data, "budget_reject_line", 1.20):
        result = "拒绝"
        reasons.append(f"{category}预算使用率将达到 {percent(budget['ratio'])}，超过预算红线。")
    if after["credit_score"] is not None and after["credit_score"] < rule_value(data, "score_reject_line", 650):
        result = "拒绝"
        reasons.append("交易后信用分低于风控线。")

    if result != "拒绝":
        flags = []
        if amount >= threshold:
            flags.append(f"金额达到家长审批线 {money(threshold)}。")
        if floor > 0 and after["cash"] < floor:
            flags.append(f"购买后现金余额低于安全线 {money(floor)}。")
        if budget["limit"] > 0 and budget["ratio"] >= rule_value(data, "budget_warning_line", 0.90):
            flags.append(f"{category}预算使用率将达到 {percent(budget['ratio'])}。")
        if delta["credit_score"] is not None and delta["credit_score"] <= -8:
            flags.append(f"信用分预计下降 {abs(delta['credit_score'])} 分。")
        if after["month_income"] > 0 and after["spend_income_ratio"] >= rule_value(data, "spend_income_warning_line", 0.85):
            flags.append(f"本月支出/收入比将达到 {percent(after['spend_income_ratio'])}。")
        if flags:
            result = "待家长审批" if amount >= threshold else "延迟"
            reasons.extend(flags)

    if result == "批准":
        reasons.append("现金余额、预算使用率、信用分变化均未触发硬性限制。")

    if result == "拒绝":
        action = "不要购买。先补现金、建预算，或等收入到账后再申请。"
    elif result == "延迟":
        action = "建议延迟到下一笔收入到账后再买，或者降低金额。"
    elif result == "待家长审批":
        action = "进入待审批队列，由爸爸或妈妈确认是否入账。"
    else:
        action = "可购买；如果阿苏提交申请，家长仍可复核。"

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
    result = "通过"
    reasons: List[str] = []

    if principal <= 0:
        result = "拒绝"
        reasons.append("放贷本金必须大于 0。")
    if after["cash"] < 0:
        result = "拒绝"
        reasons.append("放贷后现金余额为负，流动性不足。")
    if after["total_assets"] > 0 and after["loan_asset_ratio"] > hard:
        result = "拒绝"
        reasons.append(f"放贷后贷款资产占比 {percent(after['loan_asset_ratio'])}，超过红线 {percent(hard)}。")
    if after["credit_score"] is not None and after["credit_score"] < rule_value(data, "score_reject_line", 650):
        result = "拒绝"
        reasons.append("放贷后信用分低于风控线。")

    if result != "拒绝":
        flags = []
        if floor > 0 and after["cash"] < floor:
            flags.append(f"放贷后现金低于安全线 {money(floor)}。")
        if after["total_assets"] > 0 and after["loan_asset_ratio"] > soft:
            flags.append(f"放贷后贷款资产占比 {percent(after['loan_asset_ratio'])}，超过建议线 {percent(soft)}。")
        if principal > suggested and suggested > 0:
            flags.append(f"建议最高放贷金额为 {money(suggested)}。")
        if not due:
            flags.append("没有填写预计还款日，回款纪律不足。")
        if flags:
            result = "限额通过"
            reasons.extend(flags)

    if result == "通过":
        reasons.append("放贷后现金缓冲、贷款资产占比、信用分均未触发硬性限制。")
    action = "建议进入家长审批队列。"
    if result == "拒绝":
        action = "本次不建议放贷。先增加现金余额，或降低放贷金额。"
    elif result == "限额通过":
        action = f"建议限额放贷，最高不超过 {money(suggested)}，并写清还款日。"

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
        "说明": "先等同等金额收入到账，再购买，现金压力下降。",
    })
    half = round(amount / 2, 2)
    split_decision = evaluate_purchase_decision(data, half, category, desc + "（分两周第一笔）", merchant)
    scenarios.append({
        "方案": "分两周购买",
        "结论": split_decision["result"],
        "交易后现金": split_decision["metrics"]["交易后现金余额"],
        "交易后信用分": split_decision["metrics"]["交易后信用分"],
        "信用分变化": split_decision["metrics"]["信用分变化"],
        "说明": f"先支出 {money(half)}，剩余下周再处理，预算冲击更小。",
    })
    return scenarios


def submit_pending_request(data: Dict[str, Any], decision: Dict[str, Any], request_text: str, applicant: str = CHILD) -> None:
    cls = decision.get("ai") or ai_classify_purchase(request_text)
    data.setdefault("pending_requests", []).append({
        "id": uid(),
        "created_at": now_str(),
        "request_text": request_text,
        "request_type": decision.get("kind", "消费审批").replace("审批", ""),
        "amount": fnum(decision.get("metrics", {}).get("购买金额") or decision.get("metrics", {}).get("拟借出本金")),
        "category": str(decision.get("metrics", {}).get("消费分类") or "家庭贷款" if decision.get("kind") == "贷款审批" else cls.get("category", "其他")),
        "merchant": str(decision.get("metrics", {}).get("商户") or ""),
        "applicant": applicant,
        "ai_category": cls.get("category", "其他"),
        "necessity": cls.get("necessity", "未知"),
        "impulse_level": cls.get("impulse_level", "未知"),
        "python_decision": decision.get("result", "观察"),
        "parent_status": "待审批",
        "final_status": "未入账",
        "decision_json": {k: v for k, v in decision.items() if k != "simulation"},
        "pending_tx": decision.get("pending_tx"),
        "parent_note": "",
        "approved_by": "",
        "approved_at": "",
        "booked_at": "",
    })


def local_report(decision: Dict[str, Any]) -> str:
    ccy = "USD"
    lines = [f"### {decision.get('kind', '风控报告')}", f"**审批结论：{decision.get('result')}**", "", "#### 关键硬指标"]
    for k, v in decision.get("metrics", {}).items():
        lines.append(f"- {k}：{pretty_value(k, v, ccy)}")
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
    payload = {
        "kind": decision.get("kind"),
        "result": decision.get("result"),
        "metrics": decision.get("metrics"),
        "reasons": decision.get("reasons"),
        "action": decision.get("action"),
    }
    system_prompt = "你是阿苏私人银行的中文儿童金融教练。只能解释Python已计算的硬指标，禁止编造数字，禁止改变审批结论。"
    try:
        resp = requests.post(
            "https://api.deepseek.com/chat/completions",
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            json={
                "model": model,
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": json.dumps(payload, ensure_ascii=False)},
                ],
                "temperature": 0.2,
                "max_tokens": 800,
            },
            timeout=15,
        )
        resp.raise_for_status()
        return resp.json()["choices"][0]["message"]["content"]
    except Exception:
        return local_report(decision)


# ============================================================
# 7. 风险、徽章、周报、商户
# ============================================================

def add_badge(data: Dict[str, Any], name: str, description: str) -> None:
    existing = {b.get("name") for b in data.get("badges", []) if isinstance(b, dict)}
    if name not in existing:
        data.setdefault("badges", []).append({
            "id": uid(),
            "name": name,
            "earned_at": now_str(),
            "description": description,
        })


def evaluate_badges(data: Dict[str, Any]) -> None:
    m = calc_financials(data)
    if any(tx.get("type") == "收入" for tx in data.get("transactions", [])):
        add_badge(data, "第一笔收入", "已经建立现金流记录。")
    if len(data.get("budgets", {})) >= 3:
        add_badge(data, "预算建筑师", "已经建立至少三个预算分类。")
    if m.get("credit_score") is not None:
        add_badge(data, "信用分已建立", "风控系统已经可以追踪信用分。")
    if m.get("savings", 0) > 0 or any(fnum(g.get("current")) > 0 for g in data.get("goals", [])):
        add_badge(data, "储蓄启动", "已经开始建立储蓄或目标资金。")
    if data.get("pending_requests"):
        add_badge(data, "合规申请人", "已经使用待审批流程。")
    if any(b.get("status") in {"已完成", "已支付"} for b in data.get("bounties", [])):
        add_badge(data, "赏金猎人", "已经完成至少一个家庭赏金任务。")
    usage = m.get("budget_usage", {})
    if usage and all((r.get("limit", 0) <= 0 or r.get("ratio", 0) <= 1.0) for r in usage.values()):
        add_badge(data, "预算守纪律", "当前预算没有超支。")


def build_risk_radar(data: Dict[str, Any]) -> Dict[str, List[Dict[str, str]]]:
    m = calc_financials(data)
    ccy = m["currency"]
    risks = {"高风险": [], "中风险": [], "低风险": []}
    if not has_activity(data):
        risks["低风险"].append({"title": "账户为空", "detail": "先建立第一笔收入、预算或目标。"})
        return risks
    floor = fnum(data.get("settings", {}).get("cash_floor"))
    if m["cash"] < 0:
        risks["高风险"].append({"title": "现金余额为负", "detail": f"当前现金 {money(m['cash'], ccy)}，暂停非必要消费。"})
    elif floor > 0 and m["cash"] < floor:
        risks["中风险"].append({"title": "现金缓冲偏低", "detail": f"当前现金 {money(m['cash'], ccy)}，低于安全线 {money(floor, ccy)}。"})
    else:
        risks["低风险"].append({"title": "现金余额安全", "detail": f"当前现金 {money(m['cash'], ccy)}。"})

    if m["month_income"] > 0:
        if m["spend_income_ratio"] > 1:
            risks["高风险"].append({"title": "本月支出超过收入", "detail": f"支出/收入比 {percent(m['spend_income_ratio'])}。"})
        elif m["spend_income_ratio"] > rule_value(data, "spend_income_warning_line", 0.85):
            risks["中风险"].append({"title": "支出收入比偏高", "detail": f"支出/收入比 {percent(m['spend_income_ratio'])}。"})
        else:
            risks["低风险"].append({"title": "支出收入比可控", "detail": f"支出/收入比 {percent(m['spend_income_ratio'])}。"})

    if m["loan_asset_ratio"] > 0.60:
        risks["高风险"].append({"title": "贷款资产占比过高", "detail": f"应收贷款占总资产 {percent(m['loan_asset_ratio'])}。"})
    elif m["loan_asset_ratio"] > 0.45:
        risks["中风险"].append({"title": "贷款资产占比偏高", "detail": f"应收贷款占总资产 {percent(m['loan_asset_ratio'])}。"})
    else:
        risks["低风险"].append({"title": "贷款占比可控", "detail": f"应收贷款占总资产 {percent(m['loan_asset_ratio'])}。"})

    if m["liabilities"] > 0:
        risks["中风险"].append({"title": "存在未偿还负债", "detail": f"负债余额 {money(m['liabilities'], ccy)}。"})
    else:
        risks["低风险"].append({"title": "无负债", "detail": "净资产没有被负债侵蚀。"})

    for cat, row in m["budget_usage"].items():
        if row["limit"] > 0 and row["ratio"] > 1:
            risks["高风险"].append({"title": f"{cat}预算超支", "detail": f"使用率 {percent(row['ratio'])}。"})
        elif row["limit"] > 0 and row["ratio"] > 0.85:
            risks["中风险"].append({"title": f"{cat}预算接近上限", "detail": f"使用率 {percent(row['ratio'])}。"})
    if m["overdue_principal"] > 0:
        risks["高风险"].append({"title": "存在逾期贷款", "detail": f"逾期本金 {money(m['overdue_principal'], ccy)}。"})
    return risks


def build_weekly_plan(data: Dict[str, Any]) -> List[str]:
    if not has_activity(data):
        return ["先建立第一笔收入或现金余额。", "新增至少 3 个预算分类。", "设置基础信用分和现金安全线。"]
    m = calc_financials(data)
    actions: List[str] = []
    overspent = [cat for cat, r in m["budget_usage"].items() if r["limit"] > 0 and r["ratio"] > 1]
    near = [cat for cat, r in m["budget_usage"].items() if r["limit"] > 0 and 0.85 < r["ratio"] <= 1]
    if overspent:
        actions.append(f"本周暂停 {', '.join(overspent)} 类非必要消费。")
    elif near:
        actions.append(f"本周 {', '.join(near)} 类消费必须先审批。")
    if m["loan_asset_ratio"] > 0.45:
        actions.append("优先收回应收贷款本金，暂停新增放贷。")
    floor = fnum(data.get("settings", {}).get("cash_floor"))
    if floor > 0 and m["cash"] < floor:
        actions.append("下一笔收入先补现金安全线。")
    if m["savings_asset_ratio"] < 0.25 and m["month_income"] > 0:
        actions.append("下一笔收入至少 30% 转入储蓄。")
    if m["liabilities"] > 0:
        actions.append("优先偿还负债，不新增借入资金。")
    open_bounties = [b for b in data.get("bounties", []) if b.get("status") == "开放"]
    if open_bounties:
        actions.append("选择一个赏金任务，用劳动换收入，而不是只靠消费申请。")
    return actions[:5] or ["维持当前消费节奏。", "下一笔收入至少 20% 转入储蓄。"]


def generate_weekly_report(data: Dict[str, Any]) -> Dict[str, Any]:
    start = week_start()
    end = start + timedelta(days=7)
    rows = []
    for tx in data.get("transactions", []):
        d = safe_date(tx.get("date"))
        if d and start <= d < end:
            rows.append(tx)
    income = sum(fnum(tx.get("amount")) for tx in rows if tx.get("type") == "收入")
    expense = sum(fnum(tx.get("amount")) for tx in rows if tx.get("type") == "消费")
    savings_move = sum(fnum(tx.get("amount")) for tx in rows if tx.get("type") == "转入储蓄")
    completed_bounties = len([b for b in data.get("bounties", []) if b.get("status") in {"已完成", "已支付"} and b.get("reviewed_at", "")[:10] >= start.isoformat()])
    m = calc_financials(data)
    spend_by_cat: Dict[str, float] = {}
    for tx in rows:
        if tx.get("type") == "消费":
            cat = tx.get("category") or "其他"
            spend_by_cat[cat] = spend_by_cat.get(cat, 0.0) + fnum(tx.get("amount"))
    top_cat = max(spend_by_cat, key=spend_by_cat.get) if spend_by_cat else "无"
    actions = build_weekly_plan(data)
    content = "\n".join([
        f"阿苏私人银行周报：{start.isoformat()} 至 {(end - timedelta(days=1)).isoformat()}",
        f"本周收入：{money(income)}",
        f"本周消费：{money(expense)}",
        f"本周转入储蓄：{money(savings_move)}",
        f"本周完成赏金任务：{completed_bounties} 个",
        f"最大消费分类：{top_cat}",
        f"当前信用分：{score_text(m['credit_score'])}",
        "下周行动：",
        *[f"{i + 1}. {a}" for i, a in enumerate(actions)],
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
            "completed_bounties": completed_bounties,
        },
    }


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
        "本月支出收入比不高于90%": m["spend_income_ratio"] <= 0.90,
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
    return {"result": result, "message": message, "failed": failed, "checks": checks, "metrics": {"当前信用分": score, "要求信用分": required, "现金余额": m["cash"], "本月支出收入比": m["spend_income_ratio"], "该品类预算使用率": usage["ratio"], "品类开放上限": cap}}


# ============================================================
# 8. Google Sheet 展示页：手动刷新
# ============================================================

def sync_readable_sheets(data: Dict[str, Any]) -> None:
    if not gsheet_enabled():
        raise RuntimeError("Google Sheet 未配置。")
    data = normalize_data(data)
    sh = get_spreadsheet()
    m = calc_financials(data)

    summary_rows = [
        ["更新时间", now_str(), "手动刷新展示页"],
        ["总资产", m["total_assets"], "现金 + 储蓄 + 应收贷款本金"],
        ["净资产", m["net_assets"], "总资产 - 负债"],
        ["现金余额", m["cash"], "可立即使用资金"],
        ["储蓄余额", m["savings"], "储蓄账户余额"],
        ["应收贷款本金", m["receivables"], "放出但未收回本金"],
        ["负债余额", m["liabilities"], "未偿还借入资金"],
        ["今日可安全花", m["safe_to_spend"], "现金安全线和预算约束后的可用额"],
        ["信用分", score_text(m["credit_score"]), "家庭内部风控评分"],
        ["本月收入", m["month_income"], month_str()],
        ["本月消费", m["month_expense"], month_str()],
        ["支出收入比", m["spend_income_ratio"], "本月消费 / 本月收入"],
        ["待审批数量", len([r for r in data.get("pending_requests", []) if r.get("parent_status") == "待审批"]), "pending_requests"],
        ["开放赏金任务", len([b for b in data.get("bounties", []) if b.get("status") == "开放"]), "bounties"],
        ["待审核赏金任务", len([b for b in data.get("bounties", []) if b.get("status") == "已提交"]), "bounties"],
    ]
    set_table(get_or_create_ws(sh, "summary"), ["指标", "数值", "说明"], summary_rows)

    tx_rows = []
    for tx in sorted(data.get("transactions", []), key=lambda x: str(x.get("date", "")), reverse=True):
        tx_rows.append([tx.get("date", ""), tx.get("type", ""), tx.get("amount", 0), tx.get("category", ""), tx.get("party", ""), tx.get("memo", ""), tx.get("loan_id", ""), tx.get("bounty_id", "")])
    set_table(get_or_create_ws(sh, "transactions", rows=max(100, len(tx_rows) + 5), cols=8), ["日期", "类型", "金额", "分类", "对象/商户", "备注", "贷款ID", "赏金ID"], tx_rows)

    budget_rows = [[cat, row["limit"], row["spent"], row["remaining"], row["ratio"]] for cat, row in m["budget_usage"].items()]
    set_table(get_or_create_ws(sh, "budgets"), ["分类", "预算", "已花", "剩余", "使用率"], budget_rows)

    goal_rows = [[g["name"], g["category"], g["target"], g["current"], g["remaining"], g["progress"], g.get("deadline", ""), g.get("days_left", "")] for g in m["goals"]]
    set_table(get_or_create_ws(sh, "goals"), ["目标", "分类", "目标金额", "当前金额", "剩余金额", "完成度", "截止日", "剩余天数"], goal_rows)

    pending_rows = []
    for r in sorted(data.get("pending_requests", []), key=lambda x: str(x.get("created_at", "")), reverse=True):
        pending_rows.append([r.get("created_at", ""), r.get("applicant", ""), r.get("request_type", ""), r.get("amount", 0), r.get("category", ""), r.get("python_decision", ""), r.get("parent_status", ""), r.get("final_status", ""), r.get("approved_by", ""), r.get("request_text", ""), r.get("parent_note", "")])
    set_table(get_or_create_ws(sh, "pending_requests", rows=max(100, len(pending_rows) + 5), cols=11), ["申请时间", "申请人", "类型", "金额", "分类", "系统结论", "家长状态", "最终状态", "审批人", "原始申请", "家长备注"], pending_rows)

    bounty_rows = []
    for b in sorted(data.get("bounties", []), key=lambda x: str(x.get("created_at", "")), reverse=True):
        bounty_rows.append([b.get("title", ""), b.get("status", ""), b.get("reward_amount", 0), b.get("reward_points", 0), b.get("difficulty", ""), b.get("deadline", ""), b.get("created_by", ""), b.get("assigned_to", ""), b.get("submission_note", ""), b.get("reviewed_by", ""), b.get("parent_note", "")])
    set_table(get_or_create_ws(sh, "bounties", rows=max(100, len(bounty_rows) + 5), cols=11), ["任务", "状态", "赏金", "积分", "难度", "截止日", "发布人", "领取人", "提交说明", "审核人", "家长备注"], bounty_rows)

    loan_rows = [[l.get("loan_id"), l.get("date"), l.get("borrower"), l.get("principal"), l.get("repaid"), l.get("remaining"), l.get("expected_repayment"), l.get("expected_interest"), l.get("interest_received"), l.get("due_date"), l.get("status"), l.get("memo")] for l in m["loans"]]
    set_table(get_or_create_ws(sh, "loan_book"), ["贷款ID", "日期", "借款人", "本金", "已还本金", "剩余本金", "预计回款", "预计利息", "已收利息", "到期日", "状态", "备注"], loan_rows)

    score_rows = [[h.get("time", ""), h.get("score", ""), h.get("change", ""), h.get("event", ""), h.get("reason", "")] for h in sorted(data.get("score_history", []), key=lambda x: str(x.get("time", "")), reverse=True)]
    set_table(get_or_create_ws(sh, "score_history"), ["时间", "信用分", "变化", "事件", "原因"], score_rows)

    audit_rows = [[a.get("time", ""), a.get("operator", ""), a.get("event", ""), a.get("detail", "")] for a in sorted(data.get("audit_log", []), key=lambda x: str(x.get("time", "")), reverse=True)]
    set_table(get_or_create_ws(sh, "audit_log"), ["时间", "操作人", "事件", "详情"], audit_rows)


# ============================================================
# 9. UI 样式和组件
# ============================================================

def inject_css() -> None:
    st.markdown(
        """
        <style>
        :root { --bank-blue:#0b3d91; --bank-blue-2:#072b63; --ink:#111827; --muted:#6b7280; --soft:#f3f6fb; --line:#e5e7eb; }
        .block-container { padding-top: 1.1rem; max-width: 1380px; }
        .bank-hero { background: linear-gradient(135deg, #08245c 0%, #0b3d91 50%, #174ea6 100%); border-radius: 26px; padding: 28px 32px; color: white; box-shadow: 0 18px 40px rgba(11,61,145,.28); margin-bottom: 18px; }
        .bank-hero h1 { margin: 0; font-size: 2.25rem; letter-spacing: -.02em; }
        .bank-hero p { opacity: .88; margin: 8px 0 0; }
        .money-big { font-size: 2.8rem; font-weight: 800; line-height: 1.05; margin-top: 14px; }
        .subtle { color: var(--muted); font-size: .92rem; }
        .metric-card { background: white; border: 1px solid var(--line); border-radius: 22px; padding: 18px 20px; box-shadow: 0 8px 24px rgba(15,23,42,.06); min-height: 124px; }
        .metric-title { color: var(--muted); font-size: .9rem; margin-bottom: 8px; }
        .metric-value { color: var(--ink); font-size: 1.6rem; font-weight: 780; }
        .metric-note { color: var(--muted); font-size: .82rem; margin-top: 6px; }
        .chase-panel { background: #f7f9fc; border: 1px solid #e7eef8; border-radius: 24px; padding: 20px; margin: 14px 0; }
        .decision { border-radius: 18px; padding: 16px 18px; margin: 12px 0; border: 1px solid var(--line); background: white; }
        .decision.ok { border-left: 8px solid #15803d; }
        .decision.warn { border-left: 8px solid #d97706; }
        .decision.bad { border-left: 8px solid #b91c1c; }
        .pill { display:inline-block; border-radius:999px; padding:4px 10px; font-size:.78rem; font-weight:700; background:#edf2ff; color:#0b3d91; margin-bottom:8px; }
        .bounty { border-radius: 18px; padding: 16px 18px; border: 1px solid #dbeafe; background: linear-gradient(180deg,#ffffff,#f8fbff); margin-bottom: 12px; }
        .bounty h4 { margin: 0 0 6px; }
        .small { color: var(--muted); font-size:.88rem; }
        .login-card { max-width: 520px; margin: 8vh auto; background: white; border:1px solid #e5e7eb; border-radius:24px; padding:32px; box-shadow:0 20px 50px rgba(15,23,42,.12); }
        div[data-testid="stMetricValue"] { font-size: 1.7rem; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_hero(data: Dict[str, Any], metrics: Dict[str, Any]) -> None:
    ccy = metrics["currency"]
    st.markdown(
        f"""
        <div class="bank-hero">
          <div class="pill">{APP_NAME} · {role_label()}</div>
          <h1>{data.get('settings', {}).get('owner', '阿苏')}的私人银行</h1>
          <p>主线：资产、信用、预算、审批、任务收入。爸爸妈妈权限相同，阿苏走申请和任务流程。</p>
          <div class="money-big">{money(metrics['total_assets'], ccy)}</div>
          <p>总资产 = 现金 + 储蓄 + 应收贷款本金；净资产 = 总资产 - 负债。</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def metric_card(title: str, value: str, note: str = "") -> None:
    st.markdown(f"<div class='metric-card'><div class='metric-title'>{title}</div><div class='metric-value'>{value}</div><div class='metric-note'>{note}</div></div>", unsafe_allow_html=True)


def dataframe_or_empty(rows: List[Dict[str, Any]], columns: Optional[List[str]] = None, empty_text: str = "暂无数据。") -> None:
    if not rows:
        st.info(empty_text)
        return
    df = pd.DataFrame(rows)
    if columns:
        df = df[[c for c in columns if c in df.columns]]
    st.dataframe(df, use_container_width=True, hide_index=True)




def plotly_enabled() -> bool:
    return px is not None and go is not None


def render_asset_mix_chart(metrics: Dict[str, Any]) -> None:
    """资产结构图：现金、储蓄、应收贷款。"""
    rows = [
        {"资产类别": "现金", "金额": max(0.0, fnum(metrics.get("cash")))},
        {"资产类别": "储蓄", "金额": max(0.0, fnum(metrics.get("savings")))},
        {"资产类别": "应收贷款本金", "金额": max(0.0, fnum(metrics.get("receivables")))},
    ]
    df = pd.DataFrame([r for r in rows if r["金额"] > 0])
    if df.empty:
        st.info("暂无资产结构图。先建立现金、储蓄或贷款记录。")
        return
    if plotly_enabled():
        fig = px.pie(df, names="资产类别", values="金额", title="资产结构")
        fig.update_layout(margin=dict(l=10, r=10, t=50, b=10), height=320)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.bar_chart(df.set_index("资产类别"))


def render_credit_score_chart(score_history: List[Dict[str, Any]]) -> None:
    df = pd.DataFrame(score_history)
    if df.empty or "score" not in df.columns:
        st.info("暂无信用分走势。")
        return
    df["time"] = pd.to_datetime(df.get("time"), errors="coerce")
    df["score"] = pd.to_numeric(df.get("score"), errors="coerce")
    df = df.dropna(subset=["time", "score"]).sort_values("time")
    if df.empty:
        st.info("暂无可绘制的信用分走势。")
        return
    if np is not None and len(df) >= 3:
        df["三次移动平均"] = df["score"].rolling(3, min_periods=1).mean()
    if plotly_enabled():
        y_cols = ["score"] + (["三次移动平均"] if "三次移动平均" in df.columns else [])
        fig = px.line(df, x="time", y=y_cols, markers=True, title="信用分走势")
        fig.update_layout(margin=dict(l=10, r=10, t=50, b=10), height=340, yaxis_title="信用分", xaxis_title="时间")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.line_chart(df.set_index("time")[["score"]])


def render_budget_usage_chart(metrics: Dict[str, Any]) -> None:
    rows = []
    for cat, row in metrics.get("budget_usage", {}).items():
        limit = fnum(row.get("limit"))
        if limit <= 0:
            continue
        spent = fnum(row.get("spent"))
        rows.append({"分类": cat, "已花": spent, "剩余预算": max(0.0, limit - spent)})
    df = pd.DataFrame(rows)
    if df.empty:
        st.info("暂无预算图。")
        return
    if plotly_enabled():
        fig = px.bar(df, x="分类", y=["已花", "剩余预算"], barmode="stack", title="预算使用情况")
        fig.update_layout(margin=dict(l=10, r=10, t=50, b=10), height=340, yaxis_title="金额", xaxis_title="分类")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.bar_chart(df.set_index("分类"))


def render_monthly_spending_chart(rows: List[Dict[str, Any]], currency: str) -> None:
    by_cat: Dict[str, float] = {}
    for tx in rows:
        if tx.get("type") == "消费":
            by_cat[tx.get("category") or "其他"] = by_cat.get(tx.get("category") or "其他", 0.0) + fnum(tx.get("amount"))
    if not by_cat:
        st.info("该月暂无消费分类图。")
        return
    df = pd.DataFrame({"分类": list(by_cat.keys()), "金额": list(by_cat.values())}).sort_values("金额", ascending=False)
    if plotly_enabled():
        fig = px.bar(df, x="分类", y="金额", title=f"本月消费分类（{currency}）")
        fig.update_layout(margin=dict(l=10, r=10, t=50, b=10), height=340, yaxis_title="金额", xaxis_title="分类")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.bar_chart(df.set_index("分类"))


def render_monthly_cashflow_trend(data: Dict[str, Any]) -> None:
    txs = data.get("transactions", [])
    if not txs:
        st.info("暂无月度现金流趋势。")
        return
    rows = []
    for tx in txs:
        mth = tx_month(tx)
        if not mth:
            continue
        typ = tx.get("type")
        amount = fnum(tx.get("amount"))
        if typ == "收入":
            rows.append({"月份": mth, "项目": "收入", "金额": amount})
        elif typ == "消费":
            rows.append({"月份": mth, "项目": "消费", "金额": amount})
        elif typ == "还款":
            rows.append({"月份": mth, "项目": "利息收入", "金额": fnum(tx.get("interest_received"))})
    df = pd.DataFrame(rows)
    if df.empty:
        st.info("暂无可绘制的月度现金流趋势。")
        return
    grouped = df.groupby(["月份", "项目"], as_index=False)["金额"].sum().sort_values("月份")
    if plotly_enabled():
        fig = px.bar(grouped, x="月份", y="金额", color="项目", barmode="group", title="月度现金流趋势")
        fig.update_layout(margin=dict(l=10, r=10, t=50, b=10), height=340, yaxis_title="金额", xaxis_title="月份")
        st.plotly_chart(fig, use_container_width=True)
    else:
        pivot = grouped.pivot(index="月份", columns="项目", values="金额").fillna(0)
        st.bar_chart(pivot)


def build_excel_export(data: Dict[str, Any]) -> Optional[bytes]:
    if openpyxl is None:
        return None
    normalized = normalize_data(data)
    metrics = calc_financials(normalized)
    usage = metrics.get("budget_usage", {})
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame([
            {"指标": "总资产", "数值": metrics.get("total_assets")},
            {"指标": "净资产", "数值": metrics.get("net_assets")},
            {"指标": "现金", "数值": metrics.get("cash")},
            {"指标": "储蓄", "数值": metrics.get("savings")},
            {"指标": "应收贷款本金", "数值": metrics.get("receivables")},
            {"指标": "负债", "数值": metrics.get("liabilities")},
            {"指标": "信用分", "数值": metrics.get("credit_score")},
            {"指标": "本月收入", "数值": metrics.get("month_income")},
            {"指标": "本月消费", "数值": metrics.get("month_expense")},
        ]).to_excel(writer, sheet_name="summary", index=False)
        pd.DataFrame(normalized.get("transactions", [])).to_excel(writer, sheet_name="transactions", index=False)
        pd.DataFrame(normalized.get("pending_requests", [])).to_excel(writer, sheet_name="pending", index=False)
        pd.DataFrame(normalized.get("bounties", [])).to_excel(writer, sheet_name="bounties", index=False)
        pd.DataFrame(normalized.get("goals", [])).to_excel(writer, sheet_name="goals", index=False)
        pd.DataFrame(normalized.get("score_history", [])).to_excel(writer, sheet_name="score_history", index=False)
        pd.DataFrame(normalized.get("audit_log", [])).to_excel(writer, sheet_name="audit_log", index=False)
        pd.DataFrame(metrics.get("loans", [])).to_excel(writer, sheet_name="loans", index=False)
        pd.DataFrame([{"分类": k, **v} for k, v in usage.items()]).to_excel(writer, sheet_name="budgets", index=False)
        pd.DataFrame(normalized.get("weekly_reports", [])).to_excel(writer, sheet_name="weekly_reports", index=False)
    output.seek(0)
    return output.getvalue()

def render_decision(decision: Dict[str, Any]) -> None:
    css = decision_class(decision.get("result", "观察"))
    st.markdown(f"<div class='decision {css}'><div class='pill'>{decision.get('kind')}</div><h3>结论：{decision.get('result')}</h3><p class='small'>{decision.get('action','')}</p></div>", unsafe_allow_html=True)
    c1, c2 = st.columns([1, 1])
    with c1:
        st.markdown("#### 风控原因")
        for r in decision.get("reasons", []):
            st.write(f"- {r}")
    with c2:
        st.markdown("#### 关键指标")
        metric_rows = [{"指标": k, "数值": pretty_value(k, v)} for k, v in decision.get("metrics", {}).items()]
        st.dataframe(pd.DataFrame(metric_rows), use_container_width=True, hide_index=True)


def render_sidebar(data: Dict[str, Any]) -> None:
    st.sidebar.title("私人银行控制台")
    if configured_passwords():
        st.sidebar.success(f"已登录：{current_operator()}")
        if st.sidebar.button("退出登录"):
            logout()
    else:
        current = st.session_state.get("current_operator", "爸爸")
        idx = OPERATORS.index(current) if current in OPERATORS else 0
        st.sidebar.selectbox("当前操作人", OPERATORS, index=idx, key="current_operator", help="正式部署建议配置 DAD_PASSWORD / MOM_PASSWORD / ASU_PASSWORD。")
        st.sidebar.warning("未启用密码登录，当前身份可手动切换。")
    st.sidebar.caption(f"存储：{st.session_state.get('storage_backend', '未读取')}")
    if can_parent():
        st.sidebar.success("爸爸妈妈权限相同：审批、后台、设置、赏金任务。")
    else:
        st.sidebar.info("阿苏可以提交申请、领取任务、提交完成。")
    m = calc_financials(data)
    st.sidebar.metric("总资产", money(m["total_assets"], m["currency"]))
    st.sidebar.metric("信用分", score_text(m["credit_score"]))
    st.sidebar.metric("今日可安全花", money(m["safe_to_spend"], m["currency"]))


# ============================================================
# 10. 阿苏首页
# ============================================================

def page_asu_home(data: Dict[str, Any]) -> None:
    m = calc_financials(data)
    ccy = m["currency"]
    render_hero(data, m)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        metric_card("净资产", money(m["net_assets"], ccy), "总资产 - 负债")
    with c2:
        metric_card("现金", money(m["cash"], ccy), "可以立即使用，但要守安全线")
    with c3:
        metric_card("今日可安全花", money(m["safe_to_spend"], ccy), "现金安全线 + 预算剩余额")
    with c4:
        metric_card("信用分", score_text(m["credit_score"]), "越稳，权限越大")

    st.markdown("### 我要申请")
    st.caption("阿苏用一句话输入，系统会先做硬规则评估，再进入爸爸/妈妈审批。")
    with st.form("asu_quick_request"):
        text = st.text_input("申请内容", placeholder="例如：我想买一个 $12 的 Minecraft 皮肤 / 我想借给爸爸 $20 下周还 $22")
        submitted = st.form_submit_button("评估并提交申请", type="primary", use_container_width=True)
    if submitted:
        decision = parse_natural_request(data, text)
        submit_pending_request(data, decision, text, applicant=current_operator())
        commit(data, event="阿苏提交申请", reason=text)
        st.success("申请已进入爸爸/妈妈审批队列。")
        render_decision(decision)
        st.rerun()

    st.markdown("### 目标和任务")
    left, right = st.columns([1, 1])
    with left:
        st.markdown("#### 我的储蓄目标")
        if not m["goals"]:
            st.info("还没有储蓄目标。可以让爸爸妈妈在后台建立：猫咪基金、游戏基金、礼物基金。")
        for g in m["goals"]:
            st.write(f"**{g['name']}** · {money(g['current'], ccy)} / {money(g['target'], ccy)}")
            st.progress(min(1.0, fnum(g["progress"])))
            st.caption(f"还差 {money(g['remaining'], ccy)}；截止日：{g.get('deadline') or '未设置'}")
    with right:
        st.markdown("#### 可领取赏金任务")
        open_bounties = [b for b in data.get("bounties", []) if b.get("status") == "开放"]
        if not open_bounties:
            st.info("暂无开放任务。")
        for b in open_bounties[:6]:
            st.markdown(f"<div class='bounty'><h4>{b.get('title')}</h4><p class='small'>{b.get('description')}</p><div class='pill'>赏金 {money(b.get('reward_amount'), ccy)} · {b.get('reward_points')}分 · {b.get('difficulty')}</div></div>", unsafe_allow_html=True)
            if st.button("领取这个任务", key=f"claim_{b['id']}", disabled=not can_child()):
                b["status"] = "已领取"
                b["assigned_to"] = current_operator()
                b["claimed_at"] = now_str()
                commit(data, event="领取赏金任务", reason=b.get("title", ""))
                st.rerun()

    st.markdown("### 我的任务进度")
    my_bounties = [b for b in data.get("bounties", []) if b.get("assigned_to") == current_operator() or (current_operator() == CHILD and b.get("assigned_to") == CHILD)]
    if not my_bounties:
        st.info("你还没有领取任务。")
    for b in my_bounties:
        with st.expander(f"{b.get('title')} · {b.get('status')} · {money(b.get('reward_amount'), ccy)}", expanded=b.get("status") in {"已领取", "已退回"}):
            st.write(b.get("description") or "无说明")
            st.caption(f"截止日：{b.get('deadline') or '未设置'}；发布人：{b.get('created_by') or '未记录'}")
            if b.get("status") in {"已领取", "已退回"} and can_child():
                note = st.text_area("完成说明", value=b.get("submission_note", ""), key=f"submit_note_{b['id']}")
                if st.button("提交完成，等待家长审核", key=f"submit_bounty_{b['id']}"):
                    b["status"] = "已提交"
                    b["submission_note"] = note
                    b["submitted_at"] = now_str()
                    commit(data, event="提交赏金任务", reason=b.get("title", ""))
                    st.rerun()

    st.markdown("### 我的审批状态")
    mine = [r for r in data.get("pending_requests", []) if r.get("applicant") in {current_operator(), CHILD}]
    show_rows = []
    for r in sorted(mine, key=lambda x: str(x.get("created_at", "")), reverse=True)[:10]:
        show_rows.append({"时间": r.get("created_at"), "内容": r.get("request_text"), "金额": money(r.get("amount"), ccy), "系统结论": r.get("python_decision"), "家长状态": r.get("parent_status"), "最终状态": r.get("final_status"), "审批人": r.get("approved_by")})
    dataframe_or_empty(show_rows, empty_text="暂无申请记录。")

    st.markdown("### 信用分原因")
    for factor in m.get("score_factors", []):
        st.write(f"- {factor}")


# ============================================================
# 11. 家长工作台
# ============================================================

def approve_pending_request(data: Dict[str, Any], req: Dict[str, Any], note: str = "") -> None:
    tx = req.get("pending_tx")
    if tx:
        tx = normalize_tx(tx)
        tx["created_by"] = current_operator()
        tx["created_at"] = now_str()
        data.setdefault("transactions", []).append(tx)
        req["final_status"] = "已入账"
        req["booked_at"] = now_str()
    else:
        req["final_status"] = "已批准未入账"
    req["parent_status"] = "已批准"
    req["approved_by"] = current_operator()
    req["approved_at"] = now_str()
    req["parent_note"] = note


def reject_pending_request(req: Dict[str, Any], note: str = "") -> None:
    req["parent_status"] = "已拒绝"
    req["final_status"] = "不入账"
    req["approved_by"] = current_operator()
    req["approved_at"] = now_str()
    req["parent_note"] = note


def page_parent_workspace(data: Dict[str, Any]) -> None:
    st.title("家长工作台")
    if not can_parent():
        st.warning("当前身份不是爸爸/妈妈，只能查看部分信息。")
    tab1, tab2, tab3, tab4 = st.tabs(["审批中心", "发布赏金任务", "任务审核", "AI评估器"])

    with tab1:
        st.subheader("待审批申请")
        pending = [r for r in data.get("pending_requests", []) if r.get("parent_status") == "待审批"]
        if not pending:
            st.info("暂无待审批申请。")
        for r in sorted(pending, key=lambda x: str(x.get("created_at", "")), reverse=True):
            css = decision_class(r.get("python_decision", "观察"))
            st.markdown(f"<div class='decision {css}'><div class='pill'>{r.get('python_decision')}</div><h3>{r.get('request_type')}：{money(r.get('amount'))} · {r.get('category')}</h3><p class='small'>{r.get('request_text')}</p><p class='small'>申请人：{r.get('applicant')}；AI分类：{r.get('ai_category')}；必要性：{r.get('necessity')}；冲动等级：{r.get('impulse_level')}</p></div>", unsafe_allow_html=True)
            note = st.text_input("家长备注", key=f"parent_note_{r['id']}")
            c1, c2, c3 = st.columns([1, 1, 2])
            with c1:
                if st.button("批准并入账", key=f"approve_{r['id']}", disabled=not can_parent(), type="primary"):
                    approve_pending_request(data, r, note)
                    commit(data, event="批准申请", reason=r.get("request_text", ""))
                    st.rerun()
            with c2:
                if st.button("拒绝", key=f"reject_{r['id']}", disabled=not can_parent()):
                    reject_pending_request(r, note)
                    commit(data, event="拒绝申请", reason=r.get("request_text", ""))
                    st.rerun()
            with c3:
                if r.get("decision_json"):
                    with st.expander("查看系统评估"):
                        st.markdown(local_report(r.get("decision_json") or {}))

        st.divider()
        st.subheader("最近审批记录")
        rows = []
        for r in sorted(data.get("pending_requests", []), key=lambda x: str(x.get("created_at", "")), reverse=True)[:30]:
            rows.append({"时间": r.get("created_at"), "申请": r.get("request_text"), "金额": money(r.get("amount")), "系统结论": r.get("python_decision"), "家长状态": r.get("parent_status"), "最终状态": r.get("final_status"), "审批人": r.get("approved_by")})
        dataframe_or_empty(rows)

    with tab2:
        st.subheader("发布赏金任务")
        st.caption("任务收入会进入阿苏账户，形成劳动—收入—预算的闭环。")
        with st.form("bounty_form"):
            title = st.text_input("任务标题", "整理书桌并拍照提交")
            desc = st.text_area("任务说明", "把书桌清理干净，书本分类，垃圾扔掉，最后拍照或文字说明。")
            c1, c2, c3 = st.columns(3)
            with c1:
                reward_amount = st.number_input("赏金金额", min_value=0.0, value=5.0, step=1.0, format="%.2f")
            with c2:
                reward_points = st.number_input("信用积分", min_value=0, value=5, step=1)
            with c3:
                difficulty = st.selectbox("难度", ["简单", "普通", "困难"])
            c4, c5 = st.columns(2)
            with c4:
                category = st.text_input("任务分类", "家庭任务")
            with c5:
                deadline = st.date_input("截止日", value=date.today() + timedelta(days=7)).isoformat()
            submit = st.form_submit_button("发布任务", type="primary", disabled=not can_parent())
        if submit:
            data.setdefault("bounties", []).append({
                "id": uid(),
                "title": title.strip() or "未命名任务",
                "description": desc.strip(),
                "reward_amount": reward_amount,
                "reward_points": int(reward_points),
                "category": category.strip() or "家庭任务",
                "difficulty": difficulty,
                "deadline": deadline,
                "status": "开放",
                "created_by": current_operator(),
                "created_at": now_str(),
                "assigned_to": "",
                "claimed_at": "",
                "submitted_at": "",
                "submission_note": "",
                "reviewed_by": "",
                "reviewed_at": "",
                "parent_note": "",
                "paid_tx_id": "",
            })
            commit(data, event="发布赏金任务", reason=title)
            st.success("赏金任务已发布。")
            st.rerun()

        rows = []
        for b in sorted(data.get("bounties", []), key=lambda x: str(x.get("created_at", "")), reverse=True):
            rows.append({"标题": b.get("title"), "状态": b.get("status"), "赏金": money(b.get("reward_amount")), "积分": b.get("reward_points"), "截止日": b.get("deadline"), "发布人": b.get("created_by"), "领取人": b.get("assigned_to")})
        dataframe_or_empty(rows, empty_text="暂无赏金任务。")

    with tab3:
        st.subheader("任务审核")
        submitted = [b for b in data.get("bounties", []) if b.get("status") == "已提交"]
        if not submitted:
            st.info("暂无待审核任务。")
        for b in submitted:
            st.markdown(f"<div class='bounty'><h4>{b.get('title')}</h4><p class='small'>{b.get('description')}</p><p>提交说明：{b.get('submission_note') or '未填写'}</p><div class='pill'>赏金 {money(b.get('reward_amount'))} · 积分 {b.get('reward_points')}</div></div>", unsafe_allow_html=True)
            note = st.text_input("审核备注", key=f"review_note_{b['id']}")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("通过并发放赏金", key=f"approve_bounty_{b['id']}", disabled=not can_parent(), type="primary"):
                    b["status"] = "已支付"
                    b["reviewed_by"] = current_operator()
                    b["reviewed_at"] = now_str()
                    b["parent_note"] = note
                    tx = make_tx("收入", fnum(b.get("reward_amount")), "赏金任务", party=current_operator(), memo=f"赏金任务：{b.get('title')}", bounty_id=b.get("id"))
                    b["paid_tx_id"] = tx["id"]
                    data.setdefault("transactions", []).append(tx)
                    commit(data, event="发放赏金", reason=b.get("title", ""))
                    st.rerun()
            with c2:
                if st.button("退回修改", key=f"return_bounty_{b['id']}", disabled=not can_parent()):
                    b["status"] = "已退回"
                    b["reviewed_by"] = current_operator()
                    b["reviewed_at"] = now_str()
                    b["parent_note"] = note
                    commit(data, event="退回赏金任务", reason=b.get("title", ""))
                    st.rerun()

    with tab4:
        st.subheader("AI 决策评估器")
        st.caption("这里保留原来的输入—输出—评估能力：自然语言输入，Python 先算硬指标，AI 只解释。")
        text = st.text_input("输入一条申请", "我想买一个 $12 的 Minecraft 皮肤", key="parent_ai_eval_text")
        if st.button("生成评估", type="primary"):
            decision = parse_natural_request(data, text)
            st.session_state["parent_ai_eval_decision"] = decision
            st.session_state["parent_ai_eval_source_text"] = text

        decision = st.session_state.get("parent_ai_eval_decision")
        source_text = st.session_state.get("parent_ai_eval_source_text", text)
        if decision:
            render_decision(decision)
            st.markdown("#### 三种情景")
            amount = fnum(decision.get("metrics", {}).get("购买金额"))
            cat = str(decision.get("metrics", {}).get("消费分类") or "其他")
            if amount > 0 and decision.get("kind") == "消费审批":
                scenarios = scenario_planning(data, amount, cat, source_text, str(decision.get("metrics", {}).get("商户") or ""))
                st.dataframe(pd.DataFrame(scenarios), use_container_width=True, hide_index=True)
            st.markdown("#### 中文解释")
            st.markdown(deepseek_decision_report(decision))
            if st.button("把这条评估加入待审批", key="add_eval_pending"):
                submit_pending_request(data, decision, source_text, applicant=CHILD)
                commit(data, event="家长代提交评估申请", reason=source_text)
                st.session_state.pop("parent_ai_eval_decision", None)
                st.session_state.pop("parent_ai_eval_source_text", None)
                st.rerun()


# ============================================================
# 12. 私人银行后台
# ============================================================

def page_bank_backend(data: Dict[str, Any]) -> None:
    st.title("私人银行后台")
    if not can_parent():
        st.warning("后台操作需要爸爸或妈妈权限。当前身份可查看，不建议修改。")
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs(["总览", "新增交易", "贷款台账", "预算/目标", "月度账单", "风险雷达", "商户权益", "周报"])

    with tab1:
        m = calc_financials(data)
        render_hero(data, m)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("净资产", money(m["net_assets"], m["currency"]))
        c2.metric("现金", money(m["cash"], m["currency"]))
        c3.metric("储蓄", money(m["savings"], m["currency"]))
        c4.metric("应收贷款", money(m["receivables"], m["currency"]))
        c5, c6, c7, c8 = st.columns(4)
        c5.metric("本月收入", money(m["month_income"], m["currency"]))
        c6.metric("本月消费", money(m["month_expense"], m["currency"]))
        c7.metric("支出收入比", percent(m["spend_income_ratio"]))
        c8.metric("信用分", score_text(m["credit_score"]))

        st.markdown("#### 最近交易")
        rows = []
        for tx in sorted(data.get("transactions", []), key=lambda x: str(x.get("date", "")), reverse=True)[:20]:
            rows.append({"日期": tx.get("date"), "类型": tx.get("type"), "金额": money(tx.get("amount"), m["currency"]), "分类": tx.get("category"), "对象": tx.get("party"), "备注": tx.get("memo")})
        dataframe_or_empty(rows, empty_text="暂无交易。")
        st.markdown("#### 资产与现金流图表")
        chart_left, chart_right = st.columns(2)
        with chart_left:
            render_asset_mix_chart(m)
        with chart_right:
            render_monthly_cashflow_trend(data)

        if data.get("score_history"):
            st.markdown("#### 信用分走势")
            render_credit_score_chart(data.get("score_history", []))

    with tab2:
        st.subheader("新增交易")
        with st.form("tx_form"):
            c1, c2, c3 = st.columns(3)
            with c1:
                tx_date = st.date_input("日期", value=date.today()).isoformat()
                tx_type = st.selectbox("类型", TX_TYPES)
            with c2:
                amount = st.number_input("金额", min_value=0.0, value=5.0, step=1.0, format="%.2f")
                category = st.selectbox("分类", category_options(data))
            with c3:
                party = st.text_input("对象/商户/借款人", "")
                due_date = st.date_input("到期日/还款日", value=date.today() + timedelta(days=7)).isoformat()
            memo = st.text_input("备注", "")
            c4, c5, c6 = st.columns(3)
            with c4:
                expected = st.number_input("预计回款", min_value=0.0, value=0.0, step=1.0, format="%.2f")
            with c5:
                principal_repaid = st.number_input("还款本金", min_value=0.0, value=0.0, step=1.0, format="%.2f")
            with c6:
                interest_received = st.number_input("收到利息", min_value=0.0, value=0.0, step=1.0, format="%.2f")
            submitted = st.form_submit_button("保存交易", type="primary", disabled=not can_parent())
        if submitted:
            tx = make_tx(tx_type, amount, category, party=party, memo=memo, tx_date=tx_date, expected_repayment=expected, principal_repaid=principal_repaid, interest_received=interest_received, due_date=due_date if tx_type == "放贷" else "")
            data.setdefault("transactions", []).append(tx)
            commit(data, event="新增交易", reason=f"{tx_type} {money(amount)} {category}")
            st.success("交易已保存。")
            st.rerun()

        st.markdown("#### 交易流水")
        m = calc_financials(data)
        rows = [{"日期": tx.get("date"), "类型": tx.get("type"), "金额": money(tx.get("amount"), m["currency"]), "分类": tx.get("category"), "对象": tx.get("party"), "备注": tx.get("memo"), "贷款ID": tx.get("loan_id"), "赏金ID": tx.get("bounty_id")} for tx in sorted(data.get("transactions", []), key=lambda x: str(x.get("date", "")), reverse=True)]
        dataframe_or_empty(rows)

    with tab3:
        st.subheader("贷款台账")
        m = calc_financials(data)
        rows = []
        for l in m["loans"]:
            rows.append({"贷款ID": l.get("loan_id"), "日期": l.get("date"), "借款人": l.get("borrower"), "本金": money(l.get("principal"), m["currency"]), "已还本金": money(l.get("repaid"), m["currency"]), "剩余本金": money(l.get("remaining"), m["currency"]), "预计回款": money(l.get("expected_repayment"), m["currency"]), "预计利息": money(l.get("expected_interest"), m["currency"]), "已收利息": money(l.get("interest_received"), m["currency"]), "到期日": l.get("due_date"), "状态": l.get("status")})
        dataframe_or_empty(rows, empty_text="暂无贷款。")
        st.divider()
        st.markdown("#### 贷款审批模拟")
        with st.form("loan_eval_form"):
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                principal = st.number_input("拟借出本金", min_value=0.0, value=10.0, step=1.0, format="%.2f")
            with c2:
                borrower = st.text_input("借款人", "爸爸")
            with c3:
                expected = st.number_input("预计回款", min_value=0.0, value=11.0, step=1.0, format="%.2f")
            with c4:
                due = st.date_input("预计还款日", value=date.today() + timedelta(days=7)).isoformat()
            memo = st.text_input("说明", "家庭短期借款")
            submit = st.form_submit_button("评估贷款")
        if submit:
            decision = evaluate_loan_decision(data, principal, borrower, expected, due, memo)
            st.session_state["loan_eval_decision"] = decision
            st.session_state["loan_eval_text"] = f"放贷给{borrower} {money(principal)}，预计回款 {money(expected)}，到期 {due}"

        loan_decision = st.session_state.get("loan_eval_decision")
        if loan_decision:
            render_decision(loan_decision)
            if st.button("加入待审批队列", key="loan_add_pending"):
                submit_pending_request(data, loan_decision, st.session_state.get("loan_eval_text", "贷款审批"), applicant=CHILD)
                commit(data, event="提交贷款审批", reason=str(loan_decision.get("metrics", {}).get("借款人") or "贷款审批"))
                st.session_state.pop("loan_eval_decision", None)
                st.session_state.pop("loan_eval_text", None)
                st.rerun()

    with tab4:
        st.subheader("预算管理")
        m = calc_financials(data)
        usage_rows = [{"分类": cat, "预算": money(row["limit"], m["currency"]), "已花": money(row["spent"], m["currency"]), "剩余": money(row["remaining"], m["currency"]), "使用率": percent(row["ratio"])} for cat, row in m["budget_usage"].items()]
        dataframe_or_empty(usage_rows, empty_text="暂无预算。")
        render_budget_usage_chart(m)
        with st.form("budget_form"):
            c1, c2 = st.columns(2)
            with c1:
                cat = st.text_input("预算分类", "游戏")
            with c2:
                limit = st.number_input("月度预算", min_value=0.0, value=20.0, step=1.0, format="%.2f")
            submit = st.form_submit_button("保存预算", type="primary", disabled=not can_parent())
        if submit:
            data.setdefault("budgets", {})[cat.strip() or "其他"] = limit
            commit(data, event="保存预算", reason=f"{cat}={limit}")
            st.rerun()

        st.divider()
        st.subheader("储蓄目标")
        for g in m["goals"]:
            st.write(f"**{g['name']}** · {money(g['current'], m['currency'])} / {money(g['target'], m['currency'])}")
            st.progress(min(1.0, fnum(g["progress"])))
        with st.form("goal_form"):
            c1, c2, c3 = st.columns(3)
            with c1:
                name = st.text_input("目标名称", "猫咪基金")
                category = st.text_input("目标分类", "长期储蓄")
            with c2:
                target = st.number_input("目标金额", min_value=0.0, value=100.0, step=5.0, format="%.2f")
                current = st.number_input("当前金额", min_value=0.0, value=0.0, step=5.0, format="%.2f")
            with c3:
                deadline = st.date_input("截止日", value=date.today() + timedelta(days=90)).isoformat()
                note = st.text_input("说明", "")
            submit = st.form_submit_button("新增储蓄目标", disabled=not can_parent())
        if submit:
            data.setdefault("goals", []).append({"id": uid(), "name": name, "target": target, "current": current, "deadline": deadline, "category": category, "note": note})
            commit(data, event="新增储蓄目标", reason=name)
            st.rerun()

    with tab5:
        st.subheader("月度账单")
        selected_month = st.text_input("月份", month_str())
        rows = [tx for tx in data.get("transactions", []) if tx_month(tx) == selected_month]
        m2 = calc_financials(data, selected_month)
        c1, c2, c3 = st.columns(3)
        c1.metric("月收入", money(m2["month_income"], m2["currency"]))
        c2.metric("月消费", money(m2["month_expense"], m2["currency"]))
        c3.metric("支出收入比", percent(m2["spend_income_ratio"]))
        render_monthly_spending_chart(rows, m2["currency"])
        dataframe_or_empty([{"日期": tx.get("date"), "类型": tx.get("type"), "金额": money(tx.get("amount"), m2["currency"]), "分类": tx.get("category"), "对象": tx.get("party"), "备注": tx.get("memo")} for tx in rows], empty_text="该月暂无交易。")

    with tab6:
        st.subheader("风险雷达")
        risks = build_risk_radar(data)
        for level in ["高风险", "中风险", "低风险"]:
            st.markdown(f"#### {level}")
            if not risks[level]:
                st.caption("无。")
            for r in risks[level]:
                css = "bad" if level == "高风险" else "warn" if level == "中风险" else "ok"
                st.markdown(f"<div class='decision {css}'><strong>{r['title']}</strong><p class='small'>{r['detail']}</p></div>", unsafe_allow_html=True)
        st.markdown("#### 本周行动")
        for i, a in enumerate(build_weekly_plan(data), 1):
            st.write(f"{i}. {a}")

    with tab7:
        st.subheader("商户权益")
        if not data.get("merchants"):
            st.info("暂无商户。")
        for mer in data.get("merchants", []):
            res = evaluate_merchant_access(data, mer)
            css = decision_class(res["result"])
            st.markdown(f"<div class='decision {css}'><div class='pill'>{res['result']}</div><h3>{mer.get('name')} · {mer.get('category')}</h3><p class='small'>{res['message']}</p><p class='small'>失败条件：{'、'.join(res['failed']) if res['failed'] else '无'}</p></div>", unsafe_allow_html=True)
        with st.form("merchant_form"):
            c1, c2, c3 = st.columns(3)
            with c1:
                name = st.text_input("商户名称", "书店奖励")
                category = st.selectbox("分类", category_options(data))
            with c2:
                discount = st.number_input("折扣比例", min_value=0.0, max_value=1.0, value=0.10, step=0.05)
                required_score = st.number_input("最低信用分", min_value=0, max_value=850, value=720)
            with c3:
                cap = st.number_input("品类预算开放上限", min_value=0.0, max_value=2.0, value=0.90, step=0.05)
                note = st.text_input("说明", "信用好时开放")
            submit = st.form_submit_button("新增商户权益", disabled=not can_parent())
        if submit:
            data.setdefault("merchants", []).append({"id": uid(), "name": name, "category": category, "discount": discount, "required_score": required_score, "category_budget_cap": cap, "note": note})
            commit(data, event="新增商户权益", reason=name)
            st.rerun()

    with tab8:
        st.subheader("周报")
        if st.button("生成本周周报", type="primary", disabled=not can_parent()):
            report = generate_weekly_report(data)
            data.setdefault("weekly_reports", []).append(report)
            data["weekly_reports"] = data["weekly_reports"][-52:]
            commit(data, event="生成周报", reason=report["week_start"])
            st.rerun()
        if data.get("weekly_reports"):
            latest = data["weekly_reports"][-1]
            st.text_area("最新周报", latest.get("content", ""), height=260)
        rows = [{"周起始日": w.get("week_start"), "生成时间": w.get("generated_at"), "收入": money((w.get("metrics") or {}).get("income")), "消费": money((w.get("metrics") or {}).get("expense")), "信用分": (w.get("metrics") or {}).get("score"), "最大消费分类": (w.get("metrics") or {}).get("top_category")} for w in sorted(data.get("weekly_reports", []), key=lambda x: x.get("week_start", ""), reverse=True)]
        dataframe_or_empty(rows, empty_text="暂无周报。")


# ============================================================
# 13. 系统设置
# ============================================================

def page_settings(data: Dict[str, Any]) -> None:
    st.title("系统设置")
    if not can_parent():
        st.warning("系统设置需要爸爸或妈妈权限。")
    tab1, tab2, tab3, tab4 = st.tabs(["账户设置", "风控规则", "Google Sheet/备份", "数据导出"])

    with tab1:
        s = data.setdefault("settings", {})
        with st.form("settings_form"):
            c1, c2 = st.columns(2)
            with c1:
                owner = st.text_input("账户名称", s.get("owner", "阿苏"))
                currency = st.selectbox("币种", ["USD", "CNY"], index=0 if s.get("currency", "USD") == "USD" else 1)
                start_cash = st.number_input("初始现金", value=fnum(s.get("start_cash")), step=1.0, format="%.2f")
                start_savings = st.number_input("初始储蓄", value=fnum(s.get("start_savings")), step=1.0, format="%.2f")
            with c2:
                base_score = st.number_input("基础信用分", min_value=0, max_value=850, value=inum(s.get("base_score"), 720))
                cash_floor = st.number_input("现金安全线", value=fnum(s.get("cash_floor")), step=1.0, format="%.2f")
                approval_threshold = st.number_input("家长审批金额线", value=fnum(s.get("approval_threshold")), step=1.0, format="%.2f")
                sheet_sync = st.checkbox("保存时自动刷新 Google Sheet 展示页（不建议，较慢）", value=bool(s.get("sheet_report_sync_on_save", False)))
            submit = st.form_submit_button("保存账户设置", type="primary", disabled=not can_parent())
        if submit:
            s.update({"owner": owner, "currency": currency, "start_cash": start_cash, "start_savings": start_savings, "base_score": base_score, "cash_floor": cash_floor, "approval_threshold": approval_threshold, "sheet_report_sync_on_save": sheet_sync})
            commit(data, event="保存账户设置", reason="settings")
            st.rerun()

    with tab2:
        st.subheader("风控规则")
        rules = data.setdefault("rules", {})
        with st.form("rules_form"):
            c1, c2, c3 = st.columns(3)
            with c1:
                budget_warning = st.number_input("预算预警线", min_value=0.0, max_value=2.0, value=fnum(rules.get("budget_warning_line"), 0.90), step=0.05)
                budget_reject = st.number_input("预算拒绝线", min_value=0.0, max_value=3.0, value=fnum(rules.get("budget_reject_line"), 1.20), step=0.05)
            with c2:
                score_reject = st.number_input("信用分拒绝线", min_value=300, max_value=850, value=inum(rules.get("score_reject_line"), 650))
                spend_warning = st.number_input("支出收入比预警线", min_value=0.0, max_value=2.0, value=fnum(rules.get("spend_income_warning_line"), 0.85), step=0.05)
            with c3:
                loan_soft = st.number_input("贷款资产建议线", min_value=0.0, max_value=2.0, value=fnum(rules.get("loan_asset_soft_limit"), 0.45), step=0.05)
                loan_hard = st.number_input("贷款资产红线", min_value=0.0, max_value=2.0, value=fnum(rules.get("loan_asset_hard_limit"), 0.65), step=0.05)
            submit = st.form_submit_button("保存风控规则", type="primary", disabled=not can_parent())
        if submit:
            rules.update({"budget_warning_line": budget_warning, "budget_reject_line": budget_reject, "score_reject_line": score_reject, "spend_income_warning_line": spend_warning, "loan_asset_soft_limit": loan_soft, "loan_asset_hard_limit": loan_hard})
            commit(data, event="保存风控规则", reason="rules")
            st.rerun()

    with tab3:
        st.subheader("Google Sheet 和备份")
        st.write(f"当前存储：{st.session_state.get('storage_backend', '未知')}")
        c1, c2, c3 = st.columns(3)
        with c1:
            if st.button("重新从存储读取"):
                reload_data()
                st.rerun()
        with c2:
            if st.button("手动刷新 Google Sheet 展示页", disabled=not can_parent()):
                try:
                    sync_readable_sheets(data)
                    st.success("展示页已刷新。")
                except Exception as e:
                    st.error(f"刷新失败：{e}")
        with c3:
            confirm = st.checkbox("确认清空所有数据")
            if st.button("清空为空库", disabled=(not confirm or not can_parent())):
                reset_empty()
                st.success("已清空。")
                st.rerun()
        st.divider()
        st.markdown("#### 备份恢复")
        backups = data.get("backups", [])
        if not backups:
            st.info("暂无备份。每次保存会自动保留最近 30 个快照。")
        else:
            rows = [{"备份ID": b.get("id"), "时间": b.get("time"), "操作人": b.get("operator"), "事件": b.get("event"), "原因": b.get("reason")} for b in sorted(backups, key=lambda x: x.get("time", ""), reverse=True)]
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
            selected = st.selectbox("选择要恢复的备份", [b.get("id") for b in reversed(backups)])
            if st.button("恢复选中备份", disabled=not can_parent()):
                backup = next((b for b in backups if b.get("id") == selected), None)
                if backup:
                    restored = normalize_data(json.loads(backup.get("snapshot_json") or "{}"))
                    restored["backups"] = backups
                    commit(restored, event="恢复备份", reason=selected)
                    st.success("已恢复。")
                    st.rerun()

    with tab4:
        st.subheader("数据导出")
        json_text = json.dumps(normalize_data(data), ensure_ascii=False, indent=2)
        st.download_button("下载完整 JSON 备份", json_text.encode("utf-8"), "asu_money3_backup.json", "application/json")
        excel_bytes = build_excel_export(data)
        if excel_bytes is not None:
            st.download_button("下载 Excel 总账包", excel_bytes, "asu_money3_private_bank.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.info("如需 Excel 导出，请在 requirements.txt 中加入 openpyxl。")
        tx_df = pd.DataFrame(data.get("transactions", []))
        if not tx_df.empty:
            st.download_button("下载交易 CSV", tx_df.to_csv(index=False).encode("utf-8-sig"), "asu_money3_transactions.csv", "text/csv")
        bounty_df = pd.DataFrame(data.get("bounties", []))
        if not bounty_df.empty:
            st.download_button("下载赏金任务 CSV", bounty_df.to_csv(index=False).encode("utf-8-sig"), "asu_money3_bounties.csv", "text/csv")
        st.markdown("#### 审计日志")
        rows = [{"时间": a.get("time"), "操作人": a.get("operator"), "事件": a.get("event"), "详情": a.get("detail")} for a in sorted(data.get("audit_log", []), key=lambda x: x.get("time", ""), reverse=True)[:100]]
        dataframe_or_empty(rows, empty_text="暂无审计日志。")


# ============================================================
# 14. 主程序
# ============================================================

def main() -> None:
    st.set_page_config(page_title=APP_NAME, page_icon="🏦", layout="wide")
    inject_css()
    require_login()
    data = get_data()
    render_sidebar(data)

    mode = st.sidebar.radio("页面", ["阿苏首页", "家长工作台", "私人银行后台", "系统设置"], index=0)
    if mode == "阿苏首页":
        page_asu_home(data)
    elif mode == "家长工作台":
        page_parent_workspace(data)
    elif mode == "私人银行后台":
        page_bank_backend(data)
    else:
        page_settings(data)

    # 可选：不推荐开启；保留给用户显式选择。
    if data.get("settings", {}).get("sheet_report_sync_on_save") and gsheet_enabled():
        st.caption("提示：已开启保存时自动刷新展示页，这会显著降低速度。建议关闭。")


if __name__ == "__main__":
    main()
