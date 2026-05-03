# -*- coding: utf-8 -*-
"""
阿苏私人银行 3.1 — 全量重写版

核心修正：
1. 所有修改均 copy -> mutate -> commit，避免备份、审计、信用分历史失真。
2. 贷款、负债使用独立台账；还款必须绑定 loan_id / debt_id。
3. 赏金任务 reward_points 纳入信用分，且有领取、提交、审核、支付、关闭、过期。
4. 阿苏首页显示总可安全花与分类可安全花。
5. Google Sheet 日常只写 state；报表表页手动刷新，避免页面慢。

运行：streamlit run asu_money3.py
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
from typing import Any, Callable, Dict, List, Optional, Tuple

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
except Exception:
    px = None
try:
    import numpy as np
except Exception:
    np = None

APP_NAME = "阿苏私人银行"
SCHEMA_VERSION = "asu-bank-3-1"
SESSION_KEY = "asu_bank_3_1_state"
DATA_PATH = Path("asu_money3_data.json")
OLD_DATA_PATH = Path("asu_money2_data.json")
STATE_WS = "state"
STATE_KEY = "bank_data"
LAST_GOOD_KEY = "last_good_state"
UPDATED_AT_KEY = "updated_at"
PARENTS = {"爸爸", "妈妈"}
ROLES = ["阿苏", "爸爸", "妈妈"]
CATS = ["游戏", "甜品", "玩具", "学习", "宠物", "餐饮", "交通", "礼物", "奖励", "赏金任务", "家庭贷款", "负债", "其他"]
GSHEET_SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

# ------------------------- 基础工具 -------------------------
def uid(prefix: str = "") -> str:
    x = str(uuid.uuid4())
    return f"{prefix}_{x}" if prefix else x

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
    return f"${v:,.2f}" if currency == "USD" else f"{currency} {v:,.2f}"

def percent(x: Any) -> str:
    return f"{fnum(x)*100:.0f}%"

def tx_month(tx: Dict[str, Any]) -> str:
    return str(tx.get("date", ""))[:7]

def parse_date(x: Any) -> Optional[date]:
    try:
        if not x:
            return None
        return date.fromisoformat(str(x)[:10])
    except Exception:
        return None

def score_text(score: Optional[int]) -> str:
    return "未建立" if score is None else str(score)

def secret_value(key: str, default: str = "") -> str:
    try:
        if key in st.secrets:
            return str(st.secrets[key])
    except Exception:
        pass
    return os.getenv(key, default)

def current_operator() -> str:
    return st.session_state.get("current_operator", "阿苏")

def can_parent() -> bool:
    return current_operator() in PARENTS

def safe_load_json(text: str) -> Optional[Dict[str, Any]]:
    try:
        obj = json.loads(text)
        return obj if isinstance(obj, dict) else None
    except Exception:
        return None

def decision_css(x: str) -> str:
    if x in {"批准", "通过", "开放", "已批准", "已入账", "可执行"}:
        return "ok"
    if x in {"延迟", "限额通过", "暂缓开放", "待审批", "待家长审批", "观察", "未建立"}:
        return "warn"
    return "bad"

def col_letters(n: int) -> str:
    r = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        r = chr(65 + rem) + r
    return r

# ------------------------- 数据模型 -------------------------
def empty_data() -> Dict[str, Any]:
    return {
        "schema_version": SCHEMA_VERSION,
        "settings": {
            "owner": "阿苏", "currency": "USD", "start_cash": 0.0, "start_savings": 0.0,
            "base_score": 700, "cash_floor": 10.0, "approval_threshold": 15.0,
            "created_at": now_str(),
        },
        "rules": {
            "budget_warning_line": 0.90, "budget_reject_line": 1.20, "score_reject_line": 650,
            "loan_asset_soft_limit": 0.45, "loan_asset_hard_limit": 0.65,
            "debt_asset_soft_limit": 0.25, "debt_asset_hard_limit": 0.50,
            "safe_spend_savings_ratio": 0.25,
            "task_credit_window_days": 90, "task_credit_cap": 30,
        },
        "budgets": {}, "goals": [], "merchants": [], "transactions": [],
        "loans": [], "debts": [], "pending_requests": [], "bounties": [], "rewards": [],
        "badges": [], "score_history": [], "weekly_reports": [], "backups": [], "audit_log": [],
    }

def make_tx(tx_type: str, amount: float, category: str, party: str = "", memo: str = "", tx_date: str = "",
            loan_id: str = "", debt_id: str = "", goal_id: str = "", bounty_id: str = "", reward_id: str = "",
            principal_repaid: float = 0.0, interest_received: float = 0.0, expected_repayment: float = 0.0,
            due_date: str = "") -> Dict[str, Any]:
    return {
        "id": uid("tx"), "date": tx_date or today_str(), "type": tx_type, "amount": round(fnum(amount), 2),
        "category": category or "其他", "party": party or "", "memo": memo or "",
        "loan_id": loan_id or "", "debt_id": debt_id or "", "goal_id": goal_id or "", "bounty_id": bounty_id or "", "reward_id": reward_id or "",
        "principal_repaid": round(fnum(principal_repaid), 2), "interest_received": round(fnum(interest_received), 2),
        "expected_repayment": round(fnum(expected_repayment), 2), "due_date": due_date or "",
        "created_at": now_str(), "created_by": current_operator(),
    }

def normalize_tx(tx: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "id": tx.get("id") or uid("tx"), "date": tx.get("date") or today_str(), "type": tx.get("type") or "消费",
        "amount": round(fnum(tx.get("amount")), 2), "category": tx.get("category") or "其他",
        "party": tx.get("party") or tx.get("counterparty") or "", "memo": tx.get("memo") or "",
        "loan_id": tx.get("loan_id") or "", "debt_id": tx.get("debt_id") or "", "goal_id": tx.get("goal_id") or "",
        "bounty_id": tx.get("bounty_id") or "", "reward_id": tx.get("reward_id") or "",
        "principal_repaid": round(fnum(tx.get("principal_repaid")), 2),
        "interest_received": round(fnum(tx.get("interest_received")), 2),
        "expected_repayment": round(fnum(tx.get("expected_repayment")), 2), "due_date": tx.get("due_date") or "",
        "created_at": tx.get("created_at") or now_str(), "created_by": tx.get("created_by") or "",
    }

def make_loan(principal: float, borrower: str, expected_repayment: float, due_date: str, memo: str = "") -> Dict[str, Any]:
    p = round(fnum(principal), 2); e = round(fnum(expected_repayment) or p, 2)
    return {"id": uid("loan"), "created_at": now_str(), "created_by": current_operator(), "borrower": borrower or "未填写",
            "principal": p, "expected_repayment": e, "expected_interest": round(max(0, e-p), 2),
            "repaid_principal": 0.0, "interest_received": 0.0, "remaining": p, "due_date": due_date or "", "status": "未还", "memo": memo or ""}

def make_debt(amount: float, creditor: str, due_date: str = "", memo: str = "") -> Dict[str, Any]:
    a = round(fnum(amount), 2)
    return {"id": uid("debt"), "created_at": now_str(), "created_by": current_operator(), "creditor": creditor or "未填写",
            "principal": a, "repaid_principal": 0.0, "remaining": a, "due_date": due_date or "", "status": "未还", "memo": memo or ""}

def update_loan_status(l: Dict[str, Any]) -> Dict[str, Any]:
    p = fnum(l.get("principal")); repaid = fnum(l.get("repaid_principal")); rem = max(0.0, round(p - repaid, 2))
    l["remaining"] = rem; l["expected_interest"] = round(max(0.0, fnum(l.get("expected_repayment")) - p), 2)
    due = parse_date(l.get("due_date"))
    if rem <= 0.0001: l["status"] = "已结清"
    elif due and due < date.today(): l["status"] = "逾期"
    elif repaid > 0: l["status"] = "部分还款"
    else: l["status"] = "未还"
    return l

def update_debt_status(d: Dict[str, Any]) -> Dict[str, Any]:
    p = fnum(d.get("principal")); repaid = fnum(d.get("repaid_principal")); rem = max(0.0, round(p - repaid, 2))
    d["remaining"] = rem; due = parse_date(d.get("due_date"))
    if rem <= 0.0001: d["status"] = "已结清"
    elif due and due < date.today(): d["status"] = "逾期"
    elif repaid > 0: d["status"] = "部分偿还"
    else: d["status"] = "未还"
    return d

def migrate_loans(txs: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    loans: List[Dict[str, Any]] = []
    for tx in sorted(txs, key=lambda x: (str(x.get("date", "")), str(x.get("id", "")))):
        if tx.get("type") == "放贷" and fnum(tx.get("amount")) > 0:
            l = make_loan(fnum(tx.get("amount")), tx.get("party") or "未填写", fnum(tx.get("expected_repayment")) or fnum(tx.get("amount")), tx.get("due_date") or "", tx.get("memo") or "")
            l["id"] = tx.get("loan_id") or l["id"]; tx["loan_id"] = l["id"]; loans.append(l)
        elif tx.get("type") == "还款":
            amt = fnum(tx.get("principal_repaid")) or fnum(tx.get("amount")); interest = fnum(tx.get("interest_received"))
            target = next((x for x in loans if x["id"] == tx.get("loan_id")), None)
            if target is None:
                target = next((x for x in loans if x.get("remaining", 0) > 0 and (not tx.get("party") or x.get("borrower") == tx.get("party"))), None)
            if target:
                applied = min(fnum(target.get("remaining")), amt)
                target["repaid_principal"] = round(fnum(target.get("repaid_principal")) + applied, 2)
                target["interest_received"] = round(fnum(target.get("interest_received")) + interest, 2)
                tx["loan_id"] = target["id"]; update_loan_status(target)
    return [update_loan_status(x) for x in loans]

def migrate_debts(txs: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    debts: List[Dict[str, Any]] = []
    for tx in sorted(txs, key=lambda x: (str(x.get("date", "")), str(x.get("id", "")))):
        if tx.get("type") == "借入" and fnum(tx.get("amount")) > 0:
            d = make_debt(fnum(tx.get("amount")), tx.get("party") or "未填写", tx.get("due_date") or "", tx.get("memo") or "")
            d["id"] = tx.get("debt_id") or d["id"]; tx["debt_id"] = d["id"]; debts.append(d)
        elif tx.get("type") == "偿还负债":
            target = next((x for x in debts if x["id"] == tx.get("debt_id")), None)
            if target is None:
                target = next((x for x in debts if x.get("remaining", 0) > 0 and (not tx.get("party") or x.get("creditor") == tx.get("party"))), None)
            if target:
                applied = min(fnum(target.get("remaining")), fnum(tx.get("amount")))
                target["repaid_principal"] = round(fnum(target.get("repaid_principal")) + applied, 2)
                tx["debt_id"] = target["id"]; update_debt_status(target)
    return [update_debt_status(x) for x in debts]

def expire_bounties(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    out = copy.deepcopy(rows)
    for b in out:
        due = parse_date(b.get("deadline"))
        if b.get("status") == "开放" and due and due < date.today():
            b["status"] = "已过期"
    return out

def normalize_data(raw: Dict[str, Any]) -> Dict[str, Any]:
    if not isinstance(raw, dict): return empty_data()
    d = empty_data()
    if isinstance(raw.get("settings"), dict): d["settings"].update(raw["settings"])
    for k in ["start_cash", "start_savings", "cash_floor", "approval_threshold"]: d["settings"][k] = fnum(d["settings"].get(k))
    d["settings"]["base_score"] = inum(d["settings"].get("base_score"), 700)
    if isinstance(raw.get("rules"), dict):
        for k, v in raw["rules"].items(): d["rules"][str(k)] = fnum(v, d["rules"].get(str(k), 0.0))
    d["budgets"] = {str(k).strip(): round(fnum(v), 2) for k, v in (raw.get("budgets") or {}).items() if str(k).strip()}
    txs = [normalize_tx(x) for x in (raw.get("transactions") or []) if isinstance(x, dict)]
    d["transactions"] = txs
    if raw.get("loans"):
        d["loans"] = [update_loan_status({"id": x.get("id") or uid("loan"), "created_at": x.get("created_at") or now_str(), "created_by": x.get("created_by") or "", "borrower": x.get("borrower") or x.get("party") or "未填写", "principal": round(fnum(x.get("principal")),2), "expected_repayment": round(fnum(x.get("expected_repayment")) or fnum(x.get("principal")),2), "expected_interest": round(fnum(x.get("expected_interest")),2), "repaid_principal": round(fnum(x.get("repaid_principal")),2), "interest_received": round(fnum(x.get("interest_received")),2), "remaining": round(fnum(x.get("remaining")),2), "due_date": x.get("due_date") or "", "status": x.get("status") or "未还", "memo": x.get("memo") or ""}) for x in raw.get("loans", []) if isinstance(x, dict)]
    else:
        d["loans"] = migrate_loans(txs)
    if raw.get("debts"):
        d["debts"] = [update_debt_status({"id": x.get("id") or uid("debt"), "created_at": x.get("created_at") or now_str(), "created_by": x.get("created_by") or "", "creditor": x.get("creditor") or x.get("party") or "未填写", "principal": round(fnum(x.get("principal")),2), "repaid_principal": round(fnum(x.get("repaid_principal")),2), "remaining": round(fnum(x.get("remaining")),2), "due_date": x.get("due_date") or "", "status": x.get("status") or "未还", "memo": x.get("memo") or ""}) for x in raw.get("debts", []) if isinstance(x, dict)]
    else:
        d["debts"] = migrate_debts(txs)
    d["goals"] = [{"id": x.get("id") or uid("goal"), "name": x.get("name") or "未命名目标", "target": round(fnum(x.get("target")),2), "current": round(fnum(x.get("current")),2), "deadline": x.get("deadline") or "", "category": x.get("category") or "储蓄", "note": x.get("note") or ""} for x in raw.get("goals", []) if isinstance(x, dict)]
    d["merchants"] = [{"id": x.get("id") or uid("merchant"), "name": x.get("name") or "未命名商户", "category": x.get("category") or "其他", "discount": fnum(x.get("discount")), "required_score": inum(x.get("required_score"),0), "category_budget_cap": fnum(x.get("category_budget_cap"),1.0), "note": x.get("note") or ""} for x in raw.get("merchants", []) if isinstance(x, dict)]
    d["pending_requests"] = [{"id": x.get("id") or uid("req"), "created_at": x.get("created_at") or now_str(), "applicant": x.get("applicant") or "阿苏", "request_text": x.get("request_text") or "", "request_type": x.get("request_type") or "消费", "amount": round(fnum(x.get("amount")),2), "category": x.get("category") or "其他", "merchant": x.get("merchant") or "", "ai_category": x.get("ai_category") or x.get("category") or "其他", "necessity": x.get("necessity") or "未知", "impulse_level": x.get("impulse_level") or "未知", "python_decision": x.get("python_decision") or "观察", "parent_status": x.get("parent_status") or "待审批", "final_status": x.get("final_status") or "未入账", "decision_json": x.get("decision_json") or {}, "pending_tx": x.get("pending_tx"), "pending_loan": x.get("pending_loan"), "pending_debt": x.get("pending_debt"), "pending_reward_id": x.get("pending_reward_id") or "", "parent_note": x.get("parent_note") or "", "approved_by": x.get("approved_by") or "", "approved_at": x.get("approved_at") or "", "booked_at": x.get("booked_at") or ""} for x in raw.get("pending_requests", []) if isinstance(x, dict)]
    d["bounties"] = expire_bounties([{"id": x.get("id") or uid("bounty"), "title": x.get("title") or "未命名任务", "description": x.get("description") or "", "amount": round(fnum(x.get("amount")),2), "reward_points": inum(x.get("reward_points"),0), "category": x.get("category") or "家务/学习", "difficulty": x.get("difficulty") or "普通", "deadline": x.get("deadline") or "", "status": x.get("status") or "开放", "created_at": x.get("created_at") or now_str(), "created_by": x.get("created_by") or "", "claimed_at": x.get("claimed_at") or "", "submitted_at": x.get("submitted_at") or "", "paid_at": x.get("paid_at") or "", "submission_note": x.get("submission_note") or "", "review_note": x.get("review_note") or ""} for x in raw.get("bounties", []) if isinstance(x, dict)])
    d["rewards"] = [{"id": x.get("id") or uid("reward"), "title": x.get("title") or "未命名奖励券", "created_at": x.get("created_at") or now_str(), "created_by": x.get("created_by") or "", "cost": round(fnum(x.get("cost")),2), "required_score": inum(x.get("required_score"),0), "status": x.get("status") or "可用", "used_at": x.get("used_at") or "", "note": x.get("note") or ""} for x in raw.get("rewards", []) if isinstance(x, dict)]
    for k in ["badges", "score_history", "weekly_reports", "backups", "audit_log"]:
        if isinstance(raw.get(k), list): d[k] = raw[k]
    d["backups"] = d["backups"][-60:]; d["audit_log"] = d["audit_log"][-300:]; d["score_history"] = d["score_history"][-300:]
    d["schema_version"] = SCHEMA_VERSION
    return d

# ------------------------- 财务计算 -------------------------
def rule(data: Dict[str, Any], key: str, default: float) -> float:
    return fnum((data.get("rules") or {}).get(key), default)

def category_options(data: Dict[str, Any]) -> List[str]:
    s = set(CATS) | set(data.get("budgets", {}).keys()) | {tx.get("category") for tx in data.get("transactions", []) if tx.get("category")}
    return sorted(s)

def calc_budget_usage(data: Dict[str, Any], month: Optional[str] = None) -> Dict[str, Dict[str, float]]:
    month = month or month_str()
    u = {k: {"limit": fnum(v), "spent": 0.0, "remaining": fnum(v), "ratio": 0.0} for k, v in data.get("budgets", {}).items()}
    for tx in data.get("transactions", []):
        if tx.get("type") in {"消费", "奖励兑换"} and tx_month(tx) == month:
            c = tx.get("category") or "其他"; u.setdefault(c, {"limit":0.0,"spent":0.0,"remaining":0.0,"ratio":0.0}); u[c]["spent"] += fnum(tx.get("amount"))
    for row in u.values():
        row["remaining"] = max(0.0, row["limit"] - row["spent"]) if row["limit"] > 0 else 0.0
        row["ratio"] = row["spent"] / row["limit"] if row["limit"] > 0 else 0.0
    return u

def calc_goals(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    out=[]
    for g in data.get("goals", []):
        target=fnum(g.get("target")); cur=fnum(g.get("current")); due=parse_date(g.get("deadline"))
        out.append({**g, "progress": cur/target if target>0 else 0.0, "remaining": max(0.0,target-cur), "days_left": (due-date.today()).days if due else None})
    return out

def task_credit(data: Dict[str, Any]) -> Tuple[int, List[str]]:
    days=int(rule(data,"task_credit_window_days",90)); cap=int(rule(data,"task_credit_cap",30)); start=date.today()-timedelta(days=days); total=0
    for b in data.get("bounties", []):
        if b.get("status") == "已支付":
            paid=parse_date(b.get("paid_at")) or parse_date(b.get("submitted_at"))
            if paid and paid >= start: total += inum(b.get("reward_points"),0)
    total=min(cap,total)
    return total, ([f"最近 {days} 天任务信用积分：+{total}（封顶 +{cap}）"] if total else [])

def calc_credit_score(data: Dict[str, Any], cash: float, savings: float, receivables: float, liabilities: float, total_assets: float, month_income: float, month_expense: float, usage: Dict[str, Dict[str,float]], overdue_recv: float, overdue_debt: float) -> Tuple[Optional[int], List[str]]:
    base=inum(data.get("settings",{}).get("base_score"),0)
    if base<=0: return None,["账户尚未建立信用分。"]
    score=float(base); f=[]; floor=fnum(data.get("settings",{}).get("cash_floor"))
    if floor>0:
        if cash<0: score-=50; f.append("现金余额为负：-50")
        elif cash<floor: score-=18; f.append(f"现金低于安全线 {money(floor)}：-18")
        elif cash>=floor*2: score+=6; f.append("现金缓冲较充足：+6")
    if total_assets>0:
        sr=savings/total_assets
        if sr>=rule(data,"safe_spend_savings_ratio",0.25): score+=8; f.append("储蓄占比达标：+8")
        elif sr<0.10: score-=8; f.append("储蓄占比低于 10%：-8")
    if month_income<=0 and month_expense>0: score-=12; f.append("本月有消费但无收入：-12")
    elif month_income>0:
        r=month_expense/month_income
        if r>1: score-=35; f.append("本月支出超过收入：-35")
        elif r>0.85: score-=18; f.append("本月支出超过收入 85%：-18")
        elif r>0.65: score-=8; f.append("本月支出超过收入 65%：-8")
        elif r<=0.45: score+=5; f.append("本月支出控制较好：+5")
    loan_ratio=receivables/total_assets if total_assets>0 else 0
    if loan_ratio>rule(data,"loan_asset_hard_limit",0.65): score-=35; f.append("贷款资产占比超过硬红线：-35")
    elif loan_ratio>rule(data,"loan_asset_soft_limit",0.45): score-=18; f.append("贷款资产占比超过建议线：-18")
    debt_ratio=liabilities/total_assets if total_assets>0 else 0
    if debt_ratio>rule(data,"debt_asset_hard_limit",0.50): score-=45; f.append("负债占比超过硬红线：-45")
    elif debt_ratio>rule(data,"debt_asset_soft_limit",0.25): score-=25; f.append("负债占比超过建议线：-25")
    elif liabilities>0: score-=7; f.append("存在未偿还负债：-7")
    over=sum(1 for x in usage.values() if x.get("limit",0)>0 and x.get("ratio",0)>1)
    warn=sum(1 for x in usage.values() if x.get("limit",0)>0 and 0.85<x.get("ratio",0)<=1)
    if over: p=min(32,over*10); score-=p; f.append(f"{over} 个预算超支：-{p}")
    if warn: p=min(15,warn*4); score-=p; f.append(f"{warn} 个预算接近上限：-{p}")
    if overdue_recv>0: score-=24; f.append(f"存在逾期贷款 {money(overdue_recv)}：-24")
    if overdue_debt>0: score-=30; f.append(f"存在逾期负债 {money(overdue_debt)}：-30")
    pts, pf = task_credit(data); score += pts; f.extend(pf)
    return int(round(clamp(score,300,850))), (f or ["当前没有明显加分或扣分因素。"])

def calc_financials(data: Dict[str, Any], month: Optional[str] = None) -> Dict[str, Any]:
    month=month or month_str(); s=data.get("settings",{}); ccy=s.get("currency","USD"); cash=fnum(s.get("start_cash")); savings=fnum(s.get("start_savings")); mi=me=ti=te=interest=0.0
    for tx in sorted(data.get("transactions", []), key=lambda x:(str(x.get("date","")),str(x.get("created_at","")),str(x.get("id","")))):
        typ=tx.get("type"); amt=fnum(tx.get("amount")); is_m=tx_month(tx)==month
        if typ in {"收入","赏金收入"}: cash+=amt; ti+=amt; mi+=amt if is_m else 0
        elif typ in {"消费","奖励兑换"}: cash-=amt; te+=amt; me+=amt if is_m else 0
        elif typ=="转入储蓄": cash-=amt; savings+=amt
        elif typ=="储蓄取出": cash+=amt; savings-=amt
        elif typ=="放贷": cash-=amt
        elif typ=="还款":
            p=fnum(tx.get("principal_repaid")) or amt; i=fnum(tx.get("interest_received")); cash+=p+i; ti+=i; mi+=i if is_m else 0; interest+=i if is_m else 0
        elif typ=="借入": cash+=amt
        elif typ=="偿还负债": cash-=amt
    loans=[update_loan_status(copy.deepcopy(x)) for x in data.get("loans", [])]; debts=[update_debt_status(copy.deepcopy(x)) for x in data.get("debts", [])]
    recv=sum(fnum(x.get("remaining")) for x in loans); liab=sum(fnum(x.get("remaining")) for x in debts); overdue=sum(fnum(x.get("remaining")) for x in loans if x.get("status")=="逾期"); overdue_d=sum(fnum(x.get("remaining")) for x in debts if x.get("status")=="逾期")
    total=cash+savings+recv; net=total-liab; usage=calc_budget_usage(data,month); score,factors=calc_credit_score(data,cash,savings,recv,liab,total,mi,me,usage,overdue,overdue_d)
    return {"currency":ccy,"cash":cash,"savings":savings,"receivables":recv,"liabilities":liab,"total_assets":total,"net_assets":net,"month_income":mi,"month_expense":me,"month_interest":interest,"total_income":ti,"total_expense":te,"spend_income_ratio":me/mi if mi>0 else (1.0 if me>0 else 0.0),"loan_asset_ratio":recv/total if total>0 else 0.0,"debt_asset_ratio":liab/total if total>0 else 0.0,"savings_asset_ratio":savings/total if total>0 else 0.0,"budget_usage":usage,"goals":calc_goals(data),"loans":loans,"debts":debts,"overdue_principal":overdue,"overdue_debts":overdue_d,"credit_score":score,"score_factors":factors}

def calc_safe_spend(data: Dict[str, Any]) -> Dict[str, Any]:
    m=calc_financials(data); floor=fnum(data.get("settings",{}).get("cash_floor")); safe_cash=max(0.0,m["cash"]-floor); rows=[]; total_budget=0.0
    for cat,row in m["budget_usage"].items():
        rem=max(0.0,fnum(row.get("remaining"))); total_budget+=rem; rows.append({"分类":cat,"预算剩余":rem,"安全可花":min(safe_cash,rem) if row.get("limit",0)>0 else 0.0,"使用率":row.get("ratio",0.0)})
    if not rows: rows=[{"分类":"未设置预算","预算剩余":safe_cash,"安全可花":safe_cash,"使用率":0.0}]; total_safe=safe_cash
    else: total_safe=min(safe_cash,total_budget)
    return {"total_safe":round(total_safe,2),"safe_cash":round(safe_cash,2),"categories":rows}

# ------------------------- 存储 -------------------------
def gsheet_enabled() -> bool:
    if gspread is None or Credentials is None: return False
    try: return "gcp_service_account" in st.secrets and "SHEET_NAME" in st.secrets
    except Exception: return False

def get_spreadsheet():
    creds = Credentials.from_service_account_info(dict(st.secrets["gcp_service_account"]), scopes=GSHEET_SCOPES)
    return gspread.authorize(creds).open(str(st.secrets["SHEET_NAME"]))

def get_or_create_ws(sh, title: str, rows: int = 100, cols: int = 20):
    try: return sh.worksheet(title)
    except Exception: return sh.add_worksheet(title=title, rows=rows, cols=cols)

def ensure_state_header(ws) -> None:
    if ws.row_values(1)[:2] != ["key","value"]: ws.update("A1:B1", [["key","value"]], value_input_option="RAW")

def read_state(ws) -> Tuple[Dict[str,str], Dict[str,int]]:
    rows=ws.get_all_records(); vals={}; loc={}
    for i,r in enumerate(rows,start=2):
        if r.get("key"): vals[str(r.get("key"))]=str(r.get("value") or ""); loc[str(r.get("key"))]=i
    return vals,loc

def upsert(ws, loc: Dict[str,int], key: str, value: str):
    if key in loc: ws.update_cell(loc[key],2,value)
    else: ws.append_row([key,value], value_input_option="RAW")

def load_data_from_gsheet() -> Dict[str, Any]:
    sh=get_spreadsheet(); ws=get_or_create_ws(sh,STATE_WS,10,2); ensure_state_header(ws); vals,_=read_state(ws)
    for k in [STATE_KEY,LAST_GOOD_KEY]:
        raw=safe_load_json(vals.get(k,""))
        if raw is not None: return normalize_data(raw)
    data=empty_data(); save_data_to_gsheet(data); return data

def save_data_to_gsheet(data: Dict[str, Any]) -> None:
    sh=get_spreadsheet(); ws=get_or_create_ws(sh,STATE_WS,10,2); ensure_state_header(ws); vals,loc=read_state(ws)
    old=vals.get(STATE_KEY,"")
    if old.strip(): upsert(ws,loc,LAST_GOOD_KEY,old); vals,loc=read_state(ws)
    upsert(ws,loc,STATE_KEY,json.dumps(normalize_data(data),ensure_ascii=False,separators=(",",":"))); vals,loc=read_state(ws); upsert(ws,loc,UPDATED_AT_KEY,now_str())

def load_data_local() -> Dict[str, Any]:
    for p in [DATA_PATH, OLD_DATA_PATH]:
        if p.exists():
            raw=safe_load_json(p.read_text(encoding="utf-8"))
            if raw is not None:
                d=normalize_data(raw); save_data_local(d); return d
    d=empty_data(); save_data_local(d); return d

def save_data_local(data: Dict[str, Any]) -> None:
    DATA_PATH.write_text(json.dumps(normalize_data(data),ensure_ascii=False,indent=2),encoding="utf-8")

def load_data() -> Dict[str, Any]:
    if gsheet_enabled():
        try: st.session_state["storage_backend"]="Google Sheet"; return load_data_from_gsheet()
        except Exception as e: st.session_state["storage_backend"]=f"本地 JSON 兜底：{e}"; return load_data_local()
    st.session_state["storage_backend"]="本地 JSON"; return load_data_local()

def save_data(data: Dict[str, Any]) -> None:
    d=normalize_data(data)
    if gsheet_enabled():
        try: save_data_to_gsheet(d); st.session_state["storage_backend"]="Google Sheet"; return
        except Exception as e: st.session_state["storage_backend"]=f"本地 JSON 兜底：{e}"
    save_data_local(d)

def get_data() -> Dict[str, Any]:
    if SESSION_KEY not in st.session_state: st.session_state[SESSION_KEY]=load_data()
    data=normalize_data(st.session_state[SESSION_KEY])
    if any(b.get("status")=="开放" and parse_date(b.get("deadline")) and parse_date(b.get("deadline"))<date.today() for b in data.get("bounties",[])):
        mutate_and_commit(data, lambda d: d.__setitem__("bounties", expire_bounties(d.get("bounties",[]))), "任务自动过期", "截止日已过", rerun=False)
        data=normalize_data(st.session_state[SESSION_KEY])
    return data

# ------------------------- commit / 备份 / 审计 -------------------------
def strip_snapshot(data: Dict[str, Any]) -> Dict[str, Any]:
    x=normalize_data(data); x["backups"]=[]; return x

def state_summary(data: Dict[str, Any]) -> Dict[str, Any]:
    m=calc_financials(data)
    return {"cash":round(m["cash"],2),"savings":round(m["savings"],2),"receivables":round(m["receivables"],2),"liabilities":round(m["liabilities"],2),"total_assets":round(m["total_assets"],2),"net_assets":round(m["net_assets"],2),"credit_score":m["credit_score"],"tx_count":len(data.get("transactions",[])),"pending_count":len([r for r in data.get("pending_requests",[]) if r.get("parent_status")=="待审批"])}

def add_badge(data: Dict[str, Any], name: str, desc: str) -> None:
    if name not in {b.get("name") for b in data.get("badges",[])}:
        data.setdefault("badges",[]).append({"id":uid("badge"),"name":name,"earned_at":now_str(),"description":desc})

def evaluate_badges(data: Dict[str, Any]) -> None:
    m=calc_financials(data); txs=data.get("transactions",[])
    if any(t.get("type") in {"收入","赏金收入"} for t in txs): add_badge(data,"第一笔收入","已经建立现金流。")
    if len(data.get("budgets",{}))>=3: add_badge(data,"预算建筑师","建立至少三个预算分类。")
    if m.get("credit_score") is not None: add_badge(data,"信用分已建立","家庭风控系统开始运行。")
    if m.get("savings",0)>0: add_badge(data,"储蓄启动","开始把钱转向长期目标。")
    if any(b.get("status")=="已支付" for b in data.get("bounties",[])): add_badge(data,"赏金猎人","通过任务赚到收入。")
    if data.get("pending_requests"): add_badge(data,"合规申请人","学会先申请再消费。")
    if data.get("loans"): add_badge(data,"小小银行家","理解现金和应收资产转换。")
    data["badges"]=data.get("badges",[])[-80:]

def commit(new_data: Dict[str, Any], event: str, reason: str = "", rerun: bool = True) -> None:
    old=normalize_data(copy.deepcopy(st.session_state.get(SESSION_KEY, empty_data()))); new=normalize_data(copy.deepcopy(new_data))
    before=state_summary(old); evaluate_badges(new); after=state_summary(new)
    new["backups"]=(old.get("backups",[])+[{"id":uid("backup"),"time":now_str(),"operator":current_operator(),"event":event,"reason":reason,"snapshot_json":json.dumps(strip_snapshot(old),ensure_ascii=False)}])[-60:]
    new["audit_log"]=(old.get("audit_log",[])+[{"id":uid("audit"),"time":now_str(),"operator":current_operator(),"event":event,"reason":reason,"before":before,"after":after}])[-300:]
    if after.get("credit_score") is not None:
        change=None if before.get("credit_score") is None else int(after["credit_score"])-int(before["credit_score"])
        new.setdefault("score_history",[]).append({"time":now_str(),"score":after["credit_score"],"change":change,"event":event,"reason":reason or event})
        new["score_history"]=new["score_history"][-300:]
    st.session_state[SESSION_KEY]=new; save_data(new)
    if rerun: st.rerun()

def mutate_and_commit(data: Dict[str, Any], mutator: Callable[[Dict[str, Any]], None], event: str, reason: str = "", rerun: bool = True) -> None:
    new=normalize_data(copy.deepcopy(data)); mutator(new); commit(new,event,reason,rerun)

# ------------------------- 业务动作 / 审批 -------------------------
def apply_tx(d: Dict[str, Any], tx: Dict[str, Any], loan: Optional[Dict[str,Any]]=None, debt: Optional[Dict[str,Any]]=None) -> None:
    tx=normalize_tx(tx); d.setdefault("transactions",[]).append(tx)
    if loan: d.setdefault("loans",[]).append(update_loan_status(copy.deepcopy(loan)))
    if debt: d.setdefault("debts",[]).append(update_debt_status(copy.deepcopy(debt)))
    if tx.get("type")=="转入储蓄" and tx.get("goal_id"):
        for g in d.get("goals",[]):
            if g.get("id")==tx.get("goal_id"): g["current"]=round(fnum(g.get("current"))+fnum(tx.get("amount")),2)
    if tx.get("type")=="储蓄取出" and tx.get("goal_id"):
        for g in d.get("goals",[]):
            if g.get("id")==tx.get("goal_id"): g["current"]=max(0.0,round(fnum(g.get("current"))-fnum(tx.get("amount")),2))

def repay_loan_mutation(loan_id: str, principal: float, interest: float, memo: str) -> Callable[[Dict[str,Any]],None]:
    def m(d: Dict[str,Any]) -> None:
        l=next((x for x in d.get("loans",[]) if x.get("id")==loan_id),None)
        if not l: return
        applied=min(fnum(principal),fnum(l.get("remaining"))); l["repaid_principal"]=round(fnum(l.get("repaid_principal"))+applied,2); l["interest_received"]=round(fnum(l.get("interest_received"))+fnum(interest),2); update_loan_status(l)
        d.setdefault("transactions",[]).append(make_tx("还款",applied+fnum(interest),"家庭贷款",party=l.get("borrower",""),memo=memo,loan_id=loan_id,principal_repaid=applied,interest_received=interest))
    return m

def repay_debt_mutation(debt_id: str, amount: float, memo: str) -> Callable[[Dict[str,Any]],None]:
    def m(d: Dict[str,Any]) -> None:
        debt=next((x for x in d.get("debts",[]) if x.get("id")==debt_id),None)
        if not debt: return
        applied=min(fnum(amount),fnum(debt.get("remaining"))); debt["repaid_principal"]=round(fnum(debt.get("repaid_principal"))+applied,2); update_debt_status(debt)
        d.setdefault("transactions",[]).append(make_tx("偿还负债",applied,"负债",party=debt.get("creditor",""),memo=memo,debt_id=debt_id))
    return m

def simulate_tx(data: Dict[str,Any], tx: Dict[str,Any], loan: Optional[Dict[str,Any]]=None, debt: Optional[Dict[str,Any]]=None) -> Dict[str,Any]:
    before=calc_financials(data); after_data=normalize_data(copy.deepcopy(data)); apply_tx(after_data,tx,loan,debt); after=calc_financials(after_data)
    ds=None if before["credit_score"] is None or after["credit_score"] is None else after["credit_score"]-before["credit_score"]
    return {"before":before,"after":after,"delta":{"cash":after["cash"]-before["cash"],"receivables":after["receivables"]-before["receivables"],"liabilities":after["liabilities"]-before["liabilities"],"total_assets":after["total_assets"]-before["total_assets"],"net_assets":after["net_assets"]-before["net_assets"],"credit_score":ds},"tx":tx}

def evaluate_purchase(data: Dict[str,Any], amount: float, category: str, desc: str, merchant: str="") -> Dict[str,Any]:
    tx=make_tx("消费",amount,category,party=merchant,memo=desc); sim=simulate_tx(data,tx); before,after,delta=sim["before"],sim["after"],sim["delta"]; b=after["budget_usage"].get(category,{"limit":0,"spent":0,"ratio":0}); threshold=rule(data,"approval_threshold",fnum(data.get("settings",{}).get("approval_threshold"),15))
    result="批准"; reasons=[]
    if amount<=0: result="拒绝"; reasons.append("购买金额必须大于 0。")
    if after["cash"]<0: result="拒绝"; reasons.append("购买后现金为负。")
    if b.get("limit",0)>0 and b.get("ratio",0)>rule(data,"budget_reject_line",1.2): result="拒绝"; reasons.append(f"{category} 预算使用率将达 {percent(b.get('ratio'))}，超过红线。")
    if after["credit_score"] is not None and after["credit_score"]<rule(data,"score_reject_line",650): result="拒绝"; reasons.append("交易后信用分低于风控线。")
    if result!="拒绝":
        flags=[]
        if amount>=threshold: flags.append(f"金额达到审批线 {money(threshold)}。")
        if after["cash"]<fnum(data.get("settings",{}).get("cash_floor")): flags.append("购买后现金低于安全线。")
        if b.get("limit",0)>0 and b.get("ratio",0)>=rule(data,"budget_warning_line",0.9): flags.append(f"预算使用率将达 {percent(b.get('ratio'))}。")
        if delta["credit_score"] is not None and delta["credit_score"]<=-8: flags.append(f"信用分预计下降 {abs(delta['credit_score'])} 分。")
        if flags: result="待家长审批" if amount>=threshold else "延迟"; reasons.extend(flags)
    if result=="批准": reasons.append("现金、预算、信用分均未触发硬风控。")
    return {"kind":"消费审批","result":result,"reasons":reasons,"action":"进入审批或直接入账。" if result!="拒绝" else "不建议购买。", "metrics":{"购买金额":amount,"消费分类":category,"商户":merchant or "未填写","当前现金余额":before["cash"],"交易后现金余额":after["cash"],"当前信用分":before["credit_score"],"交易后信用分":after["credit_score"],"信用分变化":delta["credit_score"],"品类预算上限":b.get("limit",0),"交易后品类已花":b.get("spent",0),"交易后品类预算使用率":b.get("ratio",0),"交易后本月支出收入比":after["spend_income_ratio"]},"pending_tx":tx,"simulation":sim}

def evaluate_loan(data: Dict[str,Any], principal: float, borrower: str, expected: float, due: str, memo: str="") -> Dict[str,Any]:
    loan=make_loan(principal,borrower,expected,due,memo); tx=make_tx("放贷",principal,"家庭贷款",party=borrower,memo=memo,loan_id=loan["id"],expected_repayment=expected,due_date=due); sim=simulate_tx(data,tx,loan=loan); before,after,delta=sim["before"],sim["after"],sim["delta"]
    soft=rule(data,"loan_asset_soft_limit",0.45); hard=rule(data,"loan_asset_hard_limit",0.65); floor=fnum(data.get("settings",{}).get("cash_floor")); suggested=round(max(0.0,min(principal,max(0.0,before["cash"]-floor),max(0.0,soft*max(before["total_assets"],1)-before["receivables"]))),2)
    result="通过"; reasons=[]
    if principal<=0: result="拒绝"; reasons.append("放贷本金必须大于 0。")
    if after["cash"]<0: result="拒绝"; reasons.append("放贷后现金为负。")
    if after["loan_asset_ratio"]>hard: result="拒绝"; reasons.append(f"贷款资产占比超过硬红线 {percent(hard)}。")
    if result!="拒绝":
        flags=[]
        if after["cash"]<floor: flags.append("放贷后现金低于安全线。")
        if after["loan_asset_ratio"]>soft: flags.append(f"贷款资产占比超过建议线 {percent(soft)}。")
        if principal>suggested and suggested>0: flags.append(f"建议最高放贷 {money(suggested)}。")
        if not due: flags.append("未填写还款日。")
        if flags: result="限额通过"; reasons.extend(flags)
    if result=="通过": reasons.append("放贷后现金、贷款占比、信用分均可控。")
    return {"kind":"贷款审批","result":result,"reasons":reasons,"action":"可由家长确认或进入审批。" if result!="拒绝" else "不建议放贷。","metrics":{"拟借出本金":principal,"借款人":borrower,"预计回款":expected,"预计利息":max(0,expected-principal),"当前现金余额":before["cash"],"放贷后现金余额":after["cash"],"放贷后应收贷款本金":after["receivables"],"放贷后贷款资产占比":after["loan_asset_ratio"],"当前信用分":before["credit_score"],"放贷后信用分":after["credit_score"],"信用分变化":delta["credit_score"],"建议最高放贷金额":suggested},"pending_tx":tx,"pending_loan":loan,"simulation":sim}

def local_category(text: str) -> Dict[str,str]:
    s=(text or "").lower(); rules={"游戏":["minecraft","robux","游戏","道具","皮肤","steam"],"甜品":["甜品","奶茶","冰淇淋","蛋糕","糖","boba"],"玩具":["玩具","lego","乐高"],"学习":["书","学习","课程","文具","作业"],"宠物":["猫","猫粮","猫砂","宠物"],"餐饮":["饭","餐","pizza","披萨","汉堡"],"交通":["uber","车","公交"]}
    cat="其他"
    for k,arr in rules.items():
        if any(x in s for x in arr): cat=k; break
    return {"category":cat,"necessity":"半必要" if cat in {"学习","宠物","餐饮","交通"} else "非必要","impulse_level":"高" if cat in {"游戏","甜品","玩具"} else "中","merchant":"自然语言输入","note":"本地关键词分类"}

def ai_classify(text: str) -> Dict[str,str]:
    key=secret_value("DEEPSEEK_API_KEY"); model=secret_value("DEEPSEEK_MODEL","deepseek-chat")
    if not key or requests is None: return local_category(text)
    try:
        prompt=f"请把儿童消费/贷款请求分类，只输出JSON。字段category, necessity, impulse_level, merchant, note。可用分类：{','.join(CATS)}。输入：{text}"
        r=requests.post("https://api.deepseek.com/chat/completions",headers={"Authorization":f"Bearer {key}","Content-Type":"application/json"},json={"model":model,"messages":[{"role":"user","content":prompt}],"temperature":0,"max_tokens":200},timeout=12); r.raise_for_status(); content=r.json()["choices"][0]["message"]["content"]; m=re.search(r"\{.*\}",content,re.S)
        if not m: return local_category(text)
        o=json.loads(m.group(0)); return {"category":str(o.get("category") or "其他"),"necessity":str(o.get("necessity") or "未知"),"impulse_level":str(o.get("impulse_level") or "未知"),"merchant":str(o.get("merchant") or "自然语言输入"),"note":str(o.get("note") or "")}
    except Exception: return local_category(text)

def extract_amounts(text: str) -> List[float]:
    nums=[]
    for p in [r"\$\s*([0-9]+(?:\.[0-9]+)?)",r"([0-9]+(?:\.[0-9]+)?)\s*(?:美元|刀|块|元)"]:
        nums += [float(x.group(1)) for x in re.finditer(p,text or "",re.I)]
    if not nums: nums=[float(x.group(1)) for x in re.finditer(r"([0-9]+(?:\.[0-9]+)?)",text or "")]
    return nums

def infer_borrower(text: str) -> str:
    m=re.search(r"借给(.+?)(?:\$|[0-9]|，|,|。|$)",text or "")
    if m: return re.sub(r"\s+","",m.group(1))[:12] or "未填写"
    for name in ["爸爸","妈妈","法法","同学","朋友"]:
        if name in (text or ""): return name
    return "未填写"

def parse_request(data: Dict[str,Any], text: str) -> Dict[str,Any]:
    amts=extract_amounts(text)
    if any(x in text for x in ["借给","放贷","贷款给"]):
        p=amts[0] if amts else 0.0; e=amts[1] if len(amts)>=2 else p; return evaluate_loan(data,p,infer_borrower(text),e,"",text)
    cls=ai_classify(text); res=evaluate_purchase(data,amts[0] if amts else 0.0,cls["category"],text,cls["merchant"]); res["ai"]=cls; return res

def create_pending(data: Dict[str,Any], decision: Dict[str,Any], text: str, applicant: str) -> None:
    ai=decision.get("ai") or {}; tx=decision.get("pending_tx") or {}
    req={"id":uid("req"),"created_at":now_str(),"applicant":applicant,"request_text":text,"request_type":"贷款" if decision.get("kind")=="贷款审批" else "消费","amount":fnum(tx.get("amount")),"category":tx.get("category") or ai.get("category") or "其他","merchant":tx.get("party") or ai.get("merchant") or "","ai_category":ai.get("category") or tx.get("category") or "其他","necessity":ai.get("necessity") or "未知","impulse_level":ai.get("impulse_level") or "未知","python_decision":decision.get("result") or "观察","parent_status":"待审批","final_status":"未入账","decision_json":{"result":decision.get("result"),"reasons":decision.get("reasons"),"metrics":decision.get("metrics"),"action":decision.get("action")},"pending_tx":decision.get("pending_tx"),"pending_loan":decision.get("pending_loan"),"pending_debt":decision.get("pending_debt"),"pending_reward_id":"","parent_note":"","approved_by":"","approved_at":"","booked_at":""}
    mutate_and_commit(data, lambda d: d.setdefault("pending_requests",[]).append(req), "新增待审批申请", text)


# ------------------------- 完整业务功能面板辅助函数 -------------------------
def build_risk_radar(data: Dict[str, Any]) -> Dict[str, List[Dict[str, str]]]:
    m = calc_financials(data)
    ccy = m.get("currency", "USD")
    risks = {"高风险": [], "中风险": [], "低风险": []}
    floor = fnum(data.get("settings", {}).get("cash_floor"))
    if m["cash"] < 0:
        risks["高风险"].append({"title": "现金余额为负", "detail": f"当前现金 {money(m['cash'], ccy)}，应暂停非必要消费。"})
    elif floor > 0 and m["cash"] < floor:
        risks["中风险"].append({"title": "现金缓冲低于安全线", "detail": f"当前现金 {money(m['cash'], ccy)}，安全线 {money(floor, ccy)}。"})
    else:
        risks["低风险"].append({"title": "现金缓冲正常", "detail": f"当前现金 {money(m['cash'], ccy)}。"})
    if m["month_income"] > 0:
        if m["spend_income_ratio"] > 1:
            risks["高风险"].append({"title": "本月支出超过收入", "detail": f"支出/收入比 {percent(m['spend_income_ratio'])}。"})
        elif m["spend_income_ratio"] > 0.85:
            risks["中风险"].append({"title": "支出收入比偏高", "detail": f"支出/收入比 {percent(m['spend_income_ratio'])}。"})
        else:
            risks["低风险"].append({"title": "支出收入比可控", "detail": f"支出/收入比 {percent(m['spend_income_ratio'])}。"})
    elif m["month_expense"] > 0:
        risks["中风险"].append({"title": "有消费但无收入记录", "detail": "建议先建立收入来源，再判断消费纪律。"})
    if m["loan_asset_ratio"] > rule(data, "loan_asset_hard_limit", 0.65):
        risks["高风险"].append({"title": "贷款资产占比过高", "detail": f"应收贷款占总资产 {percent(m['loan_asset_ratio'])}。"})
    elif m["loan_asset_ratio"] > rule(data, "loan_asset_soft_limit", 0.45):
        risks["中风险"].append({"title": "贷款资产占比偏高", "detail": f"应收贷款占总资产 {percent(m['loan_asset_ratio'])}。"})
    else:
        risks["低风险"].append({"title": "贷款资产占比正常", "detail": f"应收贷款占总资产 {percent(m['loan_asset_ratio'])}。"})
    if m["debt_asset_ratio"] > rule(data, "debt_asset_hard_limit", 0.50):
        risks["高风险"].append({"title": "负债占比过高", "detail": f"负债占总资产 {percent(m['debt_asset_ratio'])}。"})
    elif m["liabilities"] > 0:
        risks["中风险"].append({"title": "存在未偿还负债", "detail": f"负债余额 {money(m['liabilities'], ccy)}。"})
    else:
        risks["低风险"].append({"title": "无未偿还负债", "detail": "净资产没有被负债侵蚀。"})
    for cat, row in m.get("budget_usage", {}).items():
        if row.get("limit", 0) > 0 and row.get("ratio", 0) > 1:
            risks["高风险"].append({"title": f"{cat} 预算超支", "detail": f"使用率 {percent(row.get('ratio'))}。"})
        elif row.get("limit", 0) > 0 and row.get("ratio", 0) >= rule(data, "budget_warning_line", 0.9):
            risks["中风险"].append({"title": f"{cat} 预算接近上限", "detail": f"使用率 {percent(row.get('ratio'))}。"})
    if m.get("overdue_principal", 0) > 0:
        risks["高风险"].append({"title": "存在逾期贷款", "detail": f"逾期应收本金 {money(m['overdue_principal'], ccy)}。"})
    if m.get("overdue_debts", 0) > 0:
        risks["高风险"].append({"title": "存在逾期负债", "detail": f"逾期负债 {money(m['overdue_debts'], ccy)}。"})
    return risks


def build_weekly_plan(data: Dict[str, Any]) -> List[str]:
    m = calc_financials(data)
    actions: List[str] = []
    overspent = [cat for cat, r in m.get("budget_usage", {}).items() if r.get("limit", 0) > 0 and r.get("ratio", 0) > 1]
    near = [cat for cat, r in m.get("budget_usage", {}).items() if r.get("limit", 0) > 0 and 0.85 < r.get("ratio", 0) <= 1]
    if overspent:
        actions.append(f"本周暂停 {', '.join(overspent)} 类非必要消费。")
    elif near:
        actions.append(f"本周 {', '.join(near)} 类消费必须先审批。")
    floor = fnum(data.get("settings", {}).get("cash_floor"))
    if floor > 0 and m["cash"] < floor:
        actions.append("下一笔收入先补现金安全线，不立刻消费。")
    if m["savings_asset_ratio"] < 0.25 and m["month_income"] > 0:
        actions.append("下一笔收入至少 20% 转入储蓄目标。")
    if m["loan_asset_ratio"] > rule(data, "loan_asset_soft_limit", 0.45):
        actions.append("优先催收或收回部分应收贷款本金。")
    if m["liabilities"] > 0:
        actions.append("有负债时，新增消费和放贷都要从严审批。")
    if not actions:
        actions = ["维持当前消费节奏，非必要消费继续走审批。", "下一笔收入优先补充储蓄目标或现金安全线。"]
    return actions[:6]


def evaluate_merchant_access(data: Dict[str, Any], merchant: Dict[str, Any]) -> Dict[str, Any]:
    m = calc_financials(data)
    cat = merchant.get("category") or "其他"
    usage = m.get("budget_usage", {}).get(cat, {"limit": 0.0, "spent": 0.0, "ratio": 0.0})
    required = inum(merchant.get("required_score"), 0)
    cap = fnum(merchant.get("category_budget_cap"), 1.0)
    score = m.get("credit_score")
    checks = {
        "信用分达标": score is not None and score >= required,
        "现金余额为正": m["cash"] > 0,
        "本月支出收入比不高于 90%": m["spend_income_ratio"] <= 0.90,
        "该品类预算未过高": usage.get("ratio", 0.0) <= cap if usage.get("limit", 0.0) > 0 else True,
    }
    failed = [k for k, ok in checks.items() if not ok]
    if all(checks.values()):
        result = "开放"
        message = f"折扣开放：{fnum(merchant.get('discount'))*100:.0f}%"
    elif checks.get("信用分达标"):
        result = "暂缓开放"
        message = "信用分达标，但现金、预算或支出收入比没有通过。"
    else:
        result = "关闭"
        message = "信用分未达标或尚未建立。"
    return {"result": result, "message": message, "failed": failed, "checks": checks, "metrics": {"当前信用分": score, "要求信用分": required, "现金余额": m["cash"], "本月支出收入比": m["spend_income_ratio"], "该品类预算使用率": usage.get("ratio", 0.0), "品类开放上限": cap}}


def build_scorecard(data: Dict[str, Any]) -> pd.DataFrame:
    m = calc_financials(data)
    rows = []
    rows.append({"模块": "基础分", "状态": data.get("settings", {}).get("base_score", 0), "说明": "家庭内部信用评分起点"})
    for factor in m.get("score_factors", []):
        rows.append({"模块": "信用因子", "状态": "已计入", "说明": factor})
    rows.append({"模块": "现金安全线", "状态": money(data.get("settings", {}).get("cash_floor"), m.get("currency", "USD")), "说明": "低于安全线会扣分并触发审批"})
    rows.append({"模块": "审批金额线", "状态": money(data.get("settings", {}).get("approval_threshold"), m.get("currency", "USD")), "说明": "超过该金额进入家长审批"})
    rows.append({"模块": "任务信用积分", "状态": task_credit(data)[0], "说明": "已支付赏金任务在窗口期内累计加分"})
    return pd.DataFrame(rows)


def scenario_planning(data: Dict[str, Any], amount: float, category: str, desc: str, merchant: str = "") -> pd.DataFrame:
    rows = []
    today_decision = evaluate_purchase(data, amount, category, desc, merchant)
    rows.append({"方案": "今天购买", "结论": today_decision["result"], "交易后现金": today_decision["metrics"].get("交易后现金余额"), "交易后信用分": today_decision["metrics"].get("交易后信用分"), "信用分变化": today_decision["metrics"].get("信用分变化"), "说明": today_decision.get("action", "")})
    next_data = normalize_data(copy.deepcopy(data))
    apply_tx(next_data, make_tx("收入", amount, "零花钱", memo="情景模拟：收入到账"))
    next_decision = evaluate_purchase(next_data, amount, category, desc + "（收入后购买）", merchant)
    rows.append({"方案": "收入后购买", "结论": next_decision["result"], "交易后现金": next_decision["metrics"].get("交易后现金余额"), "交易后信用分": next_decision["metrics"].get("交易后信用分"), "信用分变化": next_decision["metrics"].get("信用分变化"), "说明": "先等同等金额收入到账，再购买。"})
    half = round(fnum(amount) / 2, 2)
    split_decision = evaluate_purchase(data, half, category, desc + "（分两次第一笔）", merchant)
    rows.append({"方案": "分两次购买", "结论": split_decision["result"], "交易后现金": split_decision["metrics"].get("交易后现金余额"), "交易后信用分": split_decision["metrics"].get("交易后信用分"), "信用分变化": split_decision["metrics"].get("信用分变化"), "说明": f"先支出 {money(half)}，降低预算冲击。"})
    return pd.DataFrame(rows)


def local_decision_report(decision: Dict[str, Any]) -> str:
    result = decision.get("result", "观察")
    reasons = "；".join(decision.get("reasons", [])) or "没有触发明显风险。"
    if result in {"拒绝", "关闭"}:
        tone = "这不是惩罚，而是系统提醒：现在买会破坏现金安全或预算纪律。"
    elif result in {"延迟", "待家长审批", "限额通过", "暂缓开放"}:
        tone = "可以讨论，但需要先看预算、现金安全线和信用分变化。"
    else:
        tone = "当前风险可控，但仍建议记录原因，避免变成冲动消费。"
    return f"结论：{result}\n原因：{reasons}\n给阿苏的话：{tone}"


def pretty_df_money(df: pd.DataFrame, columns: List[str], currency: str) -> pd.DataFrame:
    out = df.copy()
    for col in columns:
        if col in out.columns:
            out[col] = out[col].map(lambda x: money(x, currency) if x != "" else "")
    return out


def add_download_buttons(data: Dict[str, Any], prefix: str = "asu_private_bank") -> None:
    try:
        st.download_button("下载 Excel 总账包", data=excel_bytes(data), file_name=f"{prefix}_{today_str()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.error(f"Excel 导出失败：{e}")
    st.download_button("下载 JSON 备份", data=json.dumps(data, ensure_ascii=False, indent=2), file_name=f"{prefix}_{today_str()}.json", mime="application/json")


# ------------------------- 报表 / 导出 -------------------------
def table_dfs(data: Dict[str,Any]) -> Dict[str,pd.DataFrame]:
    m=calc_financials(data)
    summary=pd.DataFrame([["总资产",m["total_assets"]],["净资产",m["net_assets"]],["现金",m["cash"]],["储蓄",m["savings"]],["应收贷款",m["receivables"]],["负债",m["liabilities"]],["信用分",m["credit_score"]]],columns=["指标","数值"])
    return {"summary":summary,"transactions":pd.DataFrame(data.get("transactions",[])),"loans":pd.DataFrame(m.get("loans",[])),"debts":pd.DataFrame(m.get("debts",[])),"pending":pd.DataFrame(data.get("pending_requests",[])),"bounties":pd.DataFrame(data.get("bounties",[])),"budgets":pd.DataFrame([{"category":k,"limit":v,**m["budget_usage"].get(k,{})} for k,v in data.get("budgets",{}).items()]),"goals":pd.DataFrame(m.get("goals",[])),"score_history":pd.DataFrame(data.get("score_history",[])),"badges":pd.DataFrame(data.get("badges",[])),"rewards":pd.DataFrame(data.get("rewards",[])),"audit_log":pd.DataFrame(data.get("audit_log",[])),"weekly_reports":pd.DataFrame(data.get("weekly_reports",[]))}

def excel_bytes(data: Dict[str,Any]) -> bytes:
    bio=BytesIO()
    with pd.ExcelWriter(bio,engine="openpyxl") as writer:
        for name,df in table_dfs(data).items(): df.to_excel(writer,sheet_name=name[:31],index=False)
    bio.seek(0); return bio.read()

def set_table(ws, header: List[str], rows: List[List[Any]]) -> None:
    ws.clear(); values=[header]+rows; ws.update(f"A1:{col_letters(len(header))}{len(values)}", values, value_input_option="RAW")

def sync_report_sheets(data: Dict[str,Any]) -> None:
    if not gsheet_enabled(): raise RuntimeError("Google Sheet 未启用。")
    sh=get_spreadsheet()
    for name,df in table_dfs(normalize_data(data)).items():
        header=list(df.columns.astype(str)) if len(df.columns) else ["empty"]; rows=df.fillna("").astype(str).values.tolist(); ws=get_or_create_ws(sh,name[:31],max(100,len(rows)+5),max(10,len(header)+2)); set_table(ws,header,rows)

def generate_weekly_report(data: Dict[str,Any]) -> Dict[str,Any]:
    start=week_start(); end=start+timedelta(days=7); rows=[tx for tx in data.get("transactions",[]) if parse_date(tx.get("date")) and start<=parse_date(tx.get("date"))<end]
    income=sum(fnum(tx.get("amount")) for tx in rows if tx.get("type") in {"收入","赏金收入"}); expense=sum(fnum(tx.get("amount")) for tx in rows if tx.get("type") in {"消费","奖励兑换"}); save=sum(fnum(tx.get("amount")) for tx in rows if tx.get("type")=="转入储蓄"); m=calc_financials(data)
    content="\n".join([f"阿苏私人银行周报：{start.isoformat()} 至 {(end-timedelta(days=1)).isoformat()}",f"本周收入：{money(income)}",f"本周消费：{money(expense)}",f"本周转入储蓄：{money(save)}",f"当前信用分：{score_text(m['credit_score'])}","下周行动：","1. 非必要消费继续走审批。","2. 下一笔收入优先补现金安全线或储蓄目标。"])
    return {"week_start":start.isoformat(),"generated_at":now_str(),"content":content,"metrics":{"income":income,"expense":expense,"savings":save,"score":m["credit_score"]}}

# ------------------------- UI -------------------------
def setup_page():
    st.set_page_config(page_title=APP_NAME, page_icon="🏦", layout="wide")
    st.markdown("""
<style>
:root{
  --chase-blue:#0b3d91;
  --chase-blue-2:#0f5bd8;
  --chase-navy:#071f49;
  --chase-red:#d71920;
  --ink:#111827;
  --muted:#6b7280;
  --line:#e5e7eb;
  --bg:#f6f8fb;
}
html,body,[data-testid="stAppViewContainer"]{background:var(--bg);}
.block-container{padding-top:1.2rem;padding-bottom:2.5rem;max-width:1180px;}
[data-testid="stSidebar"]{background:#ffffff;border-right:1px solid var(--line);}
.asu-hero{
  background:linear-gradient(135deg,var(--chase-navy),var(--chase-blue) 55%,var(--chase-blue-2));
  color:white;padding:28px 30px;border-radius:26px;
  box-shadow:0 20px 45px rgba(7,31,73,.22);margin-bottom:16px;
}
.asu-hero h1{margin:0;font-size:2.1rem;letter-spacing:.02em}.asu-hero p{margin:.5rem 0 0;color:#dbeafe;}
.card{
  border:1px solid rgba(17,24,39,.08);border-radius:20px;padding:19px;background:white;
  box-shadow:0 10px 25px rgba(7,31,73,.06);margin-bottom:13px;
}
.metric-title{font-size:.86rem;color:var(--muted);font-weight:650;letter-spacing:.01em}.metric-value{font-size:1.55rem;font-weight:850;color:var(--ink);line-height:1.25}.small-muted{color:var(--muted);font-size:.88rem;margin-top:3px}.badge{display:inline-block;padding:6px 11px;border-radius:999px;background:#eef4ff;color:var(--chase-blue);margin:4px;border:1px solid #dbeafe;font-weight:650}.ok{color:#067647;font-weight:800}.warn{color:#b54708;font-weight:800}.bad{color:#b42318;font-weight:800}
.login-wrap{max-width:520px;margin:5.5vh auto 0 auto;}
.login-card{
  background:#ffffff;border:1px solid rgba(7,31,73,.08);border-radius:28px;overflow:hidden;
  box-shadow:0 28px 70px rgba(7,31,73,.18);
}
.login-top{background:linear-gradient(135deg,var(--chase-navy),var(--chase-blue));padding:30px 32px;color:white;}
.login-bankmark{display:flex;align-items:center;gap:12px;font-weight:850;font-size:1.25rem;letter-spacing:.04em;}
.login-logo{width:34px;height:34px;border-radius:10px;background:#fff;display:inline-flex;align-items:center;justify-content:center;color:var(--chase-blue);font-weight:900;}
.login-title{font-size:2.1rem;font-weight:900;margin:24px 0 6px 0;letter-spacing:.03em;}
.login-sub{color:#dbeafe;font-size:.96rem;line-height:1.6;margin:0;}
.login-body{padding:26px 32px 30px 32px;background:white;}
.login-note{font-size:.88rem;color:var(--muted);line-height:1.55;margin-top:12px;}
div.stButton > button[kind="primary"]{background:var(--chase-blue);border-color:var(--chase-blue);border-radius:14px;font-weight:750;min-height:44px;}
div.stButton > button[kind="primary"]:hover{background:var(--chase-navy);border-color:var(--chase-navy);}
.stDataFrame{border-radius:16px;overflow:hidden;}
</style>""", unsafe_allow_html=True)


def card(title: str, value: str, note: str=""):
    st.markdown(f"<div class='card'><div class='metric-title'>{title}</div><div class='metric-value'>{value}</div><div class='small-muted'>{note}</div></div>",unsafe_allow_html=True)

def render_decision(dec: Dict[str,Any]):
    st.markdown(f"### 系统结论：<span class='{decision_css(dec.get('result','观察'))}'>{dec.get('result','观察')}</span>",unsafe_allow_html=True); st.write(dec.get("action",""))
    for r in dec.get("reasons",[]): st.write(f"- {r}")
    rows=[]
    for k,v in (dec.get("metrics") or {}).items():
        if isinstance(v,float) and ("占比" in k or "使用率" in k or "比" in k): vv=percent(v)
        elif isinstance(v,(int,float)) and "信用分" not in k and "变化" not in k: vv=money(v)
        else: vv=v
        rows.append({"指标":k,"数值":vv})
    if rows: st.dataframe(pd.DataFrame(rows),use_container_width=True,hide_index=True)

def require_login():
    keys = {"爸爸": "DAD_PASSWORD", "妈妈": "MOM_PASSWORD", "阿苏": "ASU_PASSWORD"}
    try:
        configured = any(k in st.secrets for k in keys.values())
    except Exception:
        configured = False

    if not configured:
        with st.sidebar:
            st.warning("未启用密码登录。公开部署前请配置三个密码。")
            current = st.session_state.get("current_operator", "阿苏")
            st.session_state["current_operator"] = st.selectbox(
                "当前身份", ROLES, index=ROLES.index(current) if current in ROLES else 0
            )
        return

    if st.session_state.get("authenticated"):
        st.session_state["current_operator"] = st.session_state.get("authenticated_role", "阿苏")
        return

    left, center, right = st.columns([1, 1.05, 1])
    with center:
        st.markdown(
            """
            <div class="login-card">
              <div class="login-top">
                <div class="login-bankmark"><span class="login-logo">A</span><span>ASU PRIVATE BANK</span></div>
                <div class="login-title">阿苏私人银行</div>
                <p class="login-sub">儿童版家庭银行系统。请验证身份后进入账户；爸爸和妈妈拥有相同家长权限。</p>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.write("")
        with st.container(border=True):
            role = st.selectbox("选择身份", ROLES, index=0)
            pwd = st.text_input("密码", type="password", placeholder="请输入密码")
            expected = secret_value(keys[role], "")
            login = st.button("安全进入", type="primary", use_container_width=True)
            st.caption("公开部署时请在 Streamlit Secrets 中配置 DAD_PASSWORD、MOM_PASSWORD、ASU_PASSWORD。")
            if login:
                if expected and pwd == expected:
                    st.session_state["authenticated"] = True
                    st.session_state["authenticated_role"] = role
                    st.session_state["current_operator"] = role
                    st.rerun()
                else:
                    st.error("密码错误，或该身份未配置密码。")
    st.stop()


def sidebar(data: Dict[str,Any]):
    with st.sidebar:
        st.markdown(f"### {APP_NAME}"); st.write(f"当前操作人：**{current_operator()}**"); st.caption(f"存储：{st.session_state.get('storage_backend','未知')}")
        m=calc_financials(data); st.metric("总资产",money(m["total_assets"],m["currency"])); st.metric("信用分",score_text(m["credit_score"]))
        if st.button("重新读取数据"): st.session_state.pop(SESSION_KEY,None); st.rerun()
        if st.session_state.get("authenticated") and st.button("退出登录"):
            st.session_state.pop("authenticated",None); st.session_state.pop("authenticated_role",None); st.rerun()

def require_parent() -> bool:
    if can_parent(): return True
    st.warning("这个页面需要爸爸或妈妈权限。")
    return False


# ------------------------- 可视化 / 删除交易 -------------------------
def style_plotly(fig, title: str = "", height: int = 360):
    fig.update_layout(
        title={"text": title, "x": 0.02, "xanchor": "left"} if title else None,
        height=height,
        margin=dict(l=18, r=18, t=58 if title else 24, b=28),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Arial, sans-serif", size=13, color="#111827"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    fig.update_xaxes(showgrid=True, gridcolor="#edf2f7", zeroline=False)
    fig.update_yaxes(showgrid=True, gridcolor="#edf2f7", zeroline=False)
    return fig

def asset_structure_chart(m: Dict[str, Any]):
    if px is None:
        return None
    df = pd.DataFrame({"项目": ["现金", "储蓄", "应收贷款"], "金额": [m["cash"], m["savings"], m["receivables"]]})
    df = df[df["金额"].abs() > 0.0001]
    if df.empty:
        df = pd.DataFrame({"项目": ["暂无资产"], "金额": [1]})
    fig = px.pie(df, names="项目", values="金额", hole=0.62)
    fig.update_traces(textposition="inside", textinfo="percent+label", marker=dict(line=dict(color="#ffffff", width=2)))
    return style_plotly(fig, "资产结构", 350)

def monthly_cashflow_df(data: Dict[str, Any], months: int = 6) -> pd.DataFrame:
    today = date.today().replace(day=1)
    keys = []
    y, mo = today.year, today.month
    for i in range(months - 1, -1, -1):
        mm = mo - i
        yy = y
        while mm <= 0:
            mm += 12
            yy -= 1
        keys.append(f"{yy:04d}-{mm:02d}")
    rows = {k: {"月份": k, "收入": 0.0, "消费": 0.0, "储蓄转入": 0.0, "利息收入": 0.0} for k in keys}
    for tx in data.get("transactions", []):
        k = tx_month(tx)
        if k not in rows:
            continue
        typ = tx.get("type")
        amt = fnum(tx.get("amount"))
        if typ in {"收入", "赏金收入"}:
            rows[k]["收入"] += amt
        elif typ in {"消费", "奖励兑换"}:
            rows[k]["消费"] += amt
        elif typ == "转入储蓄":
            rows[k]["储蓄转入"] += amt
        elif typ == "还款":
            rows[k]["利息收入"] += fnum(tx.get("interest_received"))
    return pd.DataFrame(list(rows.values()))

def monthly_cashflow_chart(data: Dict[str, Any]):
    if px is None:
        return None
    df = monthly_cashflow_df(data, 6)
    long = df.melt(id_vars="月份", value_vars=["收入", "消费", "储蓄转入", "利息收入"], var_name="项目", value_name="金额")
    fig = px.bar(long, x="月份", y="金额", color="项目", barmode="group", text="金额")
    fig.update_traces(texttemplate="$%{y:,.0f}", textposition="outside", cliponaxis=False)
    return style_plotly(fig, "近 6 个月现金流", 390)

def budget_usage_chart(data: Dict[str, Any], m: Dict[str, Any]):
    if px is None:
        return None
    rows = []
    for cat, row in m.get("budget_usage", {}).items():
        if fnum(row.get("limit")) > 0 or fnum(row.get("spent")) > 0:
            rows.append({"分类": cat, "已花": fnum(row.get("spent")), "剩余": max(0.0, fnum(row.get("remaining"))), "使用率": fnum(row.get("ratio"))})
    if not rows:
        return None
    df = pd.DataFrame(rows).sort_values("使用率", ascending=True)
    long = df.melt(id_vars=["分类", "使用率"], value_vars=["已花", "剩余"], var_name="项目", value_name="金额")
    fig = px.bar(long, y="分类", x="金额", color="项目", orientation="h", barmode="stack", text="金额", hover_data={"使用率": ":.0%"})
    fig.update_traces(texttemplate="$%{x:,.0f}", textposition="inside")
    return style_plotly(fig, "本月预算使用", max(330, 48 * len(df) + 120))

def expense_category_chart(data: Dict[str, Any], m: Dict[str, Any]):
    if px is None:
        return None
    rows = []
    for cat, row in m.get("budget_usage", {}).items():
        spent = fnum(row.get("spent"))
        if spent > 0:
            rows.append({"分类": cat, "金额": spent})
    if not rows:
        return None
    df = pd.DataFrame(rows).sort_values("金额", ascending=False)
    fig = px.bar(df, x="分类", y="金额", text="金额")
    fig.update_traces(texttemplate="$%{y:,.0f}", textposition="outside", cliponaxis=False)
    return style_plotly(fig, "本月消费分类", 360)

def score_trend_chart(data: Dict[str, Any]):
    if px is None or not data.get("score_history"):
        return None
    sdf = pd.DataFrame(data.get("score_history"))
    if "score" not in sdf.columns or "time" not in sdf.columns:
        return None
    sdf["time"] = pd.to_datetime(sdf["time"], errors="coerce")
    sdf["score"] = pd.to_numeric(sdf["score"], errors="coerce")
    sdf = sdf.dropna(subset=["time", "score"]).sort_values("time")
    if sdf.empty:
        return None
    sdf["三次移动平均"] = sdf["score"].rolling(3, min_periods=1).mean()
    fig = px.line(sdf, x="time", y=["score", "三次移动平均"], markers=True)
    fig.update_yaxes(range=[max(300, sdf[["score", "三次移动平均"]].min().min() - 25), min(850, sdf[["score", "三次移动平均"]].max().max() + 25)])
    return style_plotly(fig, "信用分走势", 360)

def delete_transaction_mutation(tx_id: str) -> Callable[[Dict[str, Any]], None]:
    def m(d: Dict[str, Any]) -> None:
        tx = next((x for x in d.get("transactions", []) if x.get("id") == tx_id), None)
        if not tx:
            return
        d["transactions"] = [x for x in d.get("transactions", []) if x.get("id") != tx_id]
        gid = tx.get("goal_id") or ""
        if gid:
            for g in d.get("goals", []):
                if g.get("id") == gid:
                    if tx.get("type") == "转入储蓄":
                        g["current"] = max(0.0, round(fnum(g.get("current")) - fnum(tx.get("amount")), 2))
                    elif tx.get("type") == "储蓄取出":
                        g["current"] = round(fnum(g.get("current")) + fnum(tx.get("amount")), 2)
        txs = [normalize_tx(copy.deepcopy(x)) for x in d.get("transactions", [])]
        d["loans"] = migrate_loans(txs)
        d["debts"] = migrate_debts(txs)
        d["transactions"] = txs
    return m

def transaction_display_df(data: Dict[str, Any], currency: str) -> pd.DataFrame:
    rows = []
    for tx in sorted(data.get("transactions", []), key=lambda x: (str(x.get("date", "")), str(x.get("created_at", ""))), reverse=True):
        rows.append({
            "日期": tx.get("date", ""),
            "类型": tx.get("type", ""),
            "金额": money(tx.get("amount"), currency),
            "分类": tx.get("category", ""),
            "对象/商户": tx.get("party", ""),
            "备注": tx.get("memo", ""),
            "ID": str(tx.get("id", ""))[-8:],
        })
    return pd.DataFrame(rows)

# ------------------------- pages -------------------------
def page_asu(data: Dict[str,Any]):
    m=calc_financials(data); safe=calc_safe_spend(data); ccy=m["currency"]
    st.markdown("<div class='asu-hero'><h1>阿苏私人银行</h1><p>看清资产，守住预算，赚到自己的钱。</p></div>",unsafe_allow_html=True); st.write("")
    c1,c2,c3,c4=st.columns(4)
    with c1: card("总资产",money(m["total_assets"],ccy),"现金+储蓄+应收贷款")
    with c2: card("净资产",money(m["net_assets"],ccy),"总资产-负债")
    with c3: card("今日可安全花",money(safe["total_safe"],ccy),"不碰安全线和预算")
    with c4: card("信用分",score_text(m["credit_score"]),"家庭风控评分")
    left,right=st.columns([1.2,1])
    with left:
        st.subheader("我要申请"); text=st.text_area("写清楚：想买什么、多少钱、为什么。也可以写：借给爸爸 20，预计还 22。",key="asu_req")
        if st.button("生成风控评估",type="primary") and text.strip(): st.session_state["asu_decision"]=parse_request(data,text)
        if st.session_state.get("asu_decision"):
            render_decision(st.session_state["asu_decision"])
            if st.button("提交给爸爸妈妈审批"): create_pending(data,st.session_state["asu_decision"],text,"阿苏")
        st.subheader("分类可安全花")
        df=pd.DataFrame(safe["categories"])
        if not df.empty:
            show=df.copy(); show["预算剩余"]=show["预算剩余"].map(lambda x:money(x,ccy)); show["安全可花"]=show["安全可花"].map(lambda x:money(x,ccy)); show["使用率"]=show["使用率"].map(percent); st.dataframe(show,use_container_width=True,hide_index=True)
        st.subheader("储蓄目标")
        if not m["goals"]: st.info("还没有储蓄目标。")
        for g in m["goals"]:
            st.write(f"**{g['name']}**：{money(g['current'],ccy)} / {money(g['target'],ccy)}"); st.progress(min(1.0,fnum(g.get("progress")))); st.caption(f"还差 {money(g['remaining'],ccy)}")
    with right:
        st.subheader("赏金任务")
        for b in [x for x in data.get("bounties",[]) if x.get("status")=="开放"][:8]:
            with st.container(border=True):
                st.write(f"**{b['title']}** · {money(b['amount'],ccy)} · +{b.get('reward_points',0)} 分"); st.caption(f"{b.get('description','')}｜截止：{b.get('deadline') or '无'}")
                if st.button("领取任务",key=f"claim_{b['id']}"):
                    mutate_and_commit(data,lambda d: [x.update({"status":"进行中","claimed_at":now_str()}) for x in d.get("bounties",[]) if x.get("id")==b["id"] and x.get("status")=="开放"],"领取赏金任务",b["title"])
        for b in [x for x in data.get("bounties",[]) if x.get("status") in {"进行中","待审核"}][:8]:
            with st.container(border=True):
                st.write(f"**{b['title']}** · {b['status']}")
                if b.get("status")=="进行中":
                    note=st.text_area("完成说明",key=f"sub_{b['id']}")
                    if st.button("提交审核",key=f"submit_{b['id']}"):
                        mutate_and_commit(data,lambda d: [x.update({"status":"待审核","submitted_at":now_str(),"submission_note":note}) for x in d.get("bounties",[]) if x.get("id")==b["id"] and x.get("status")=="进行中"],"提交赏金任务",b["title"])
        st.subheader("徽章墙")
        st.markdown(" ".join([f"<span class='badge'>{b.get('name')}</span>" for b in sorted(data.get("badges",[]),key=lambda x:str(x.get("earned_at","")),reverse=True)[:24]]) or "暂无徽章",unsafe_allow_html=True)
        st.subheader("奖励券")
        for r in [x for x in data.get("rewards",[]) if x.get("status")=="可用"][:8]:
            with st.container(border=True):
                st.write(f"**{r['title']}**"); st.caption(f"成本：{money(r.get('cost'),ccy)}｜要求信用分：{r.get('required_score')}")
                if st.button("申请兑换",key=f"redeem_{r['id']}"):
                    req={"id":uid("req"),"created_at":now_str(),"applicant":"阿苏","request_text":f"申请兑换奖励券：{r['title']}","request_type":"奖励兑换","amount":fnum(r.get("cost")),"category":"奖励","merchant":r["title"],"ai_category":"奖励","necessity":"非必要","impulse_level":"低","python_decision":"待审批","parent_status":"待审批","final_status":"未入账","decision_json":{},"pending_tx":None,"pending_loan":None,"pending_debt":None,"pending_reward_id":r["id"],"parent_note":"","approved_by":"","approved_at":"","booked_at":""}
                    mutate_and_commit(data,lambda d: d.setdefault("pending_requests",[]).append(req),"申请奖励兑换",r["title"])

def page_parent(data: Dict[str,Any]):
    if not require_parent(): return
    st.title("家长工作台"); tabs=st.tabs(["待审批","发布赏金","任务审核","AI评估","奖励券"]); ccy=data.get("settings",{}).get("currency","USD")
    with tabs[0]:
        pending=[r for r in data.get("pending_requests",[]) if r.get("parent_status")=="待审批"]
        if not pending: st.info("暂无待审批。")
        for r in sorted(pending,key=lambda x:str(x.get("created_at","")),reverse=True):
            with st.container(border=True):
                st.write(f"**{r.get('request_type')}｜{money(r.get('amount'),ccy)}｜{r.get('category')}**"); st.write(r.get("request_text")); st.caption(f"系统：{r.get('python_decision')}｜必要性：{r.get('necessity')}｜冲动：{r.get('impulse_level')}")
                note=st.text_input("备注",key=f"note_{r['id']}"); a,b=st.columns(2)
                if a.button("批准并入账",type="primary",key=f"ap_{r['id']}"):
                    def mm(d):
                        req=next((x for x in d.get("pending_requests",[]) if x.get("id")==r["id"]),None)
                        if not req or req.get("parent_status")!="待审批": return
                        rid=req.get("pending_reward_id")
                        if rid:
                            reward=next((x for x in d.get("rewards",[]) if x.get("id")==rid),None)
                            if reward:
                                reward["status"]="已兑换"; reward["used_at"]=now_str()
                                if fnum(reward.get("cost"))>0: d.setdefault("transactions",[]).append(make_tx("奖励兑换",reward.get("cost"),"奖励",party=reward.get("title"),memo="兑换奖励券",reward_id=rid))
                        elif req.get("pending_tx"): apply_tx(d,req.get("pending_tx"),req.get("pending_loan"),req.get("pending_debt"))
                        req.update({"parent_status":"已批准","final_status":"已入账","approved_by":current_operator(),"approved_at":now_str(),"booked_at":now_str(),"parent_note":note})
                    mutate_and_commit(data,mm,"批准申请并入账",r.get("request_text",""))
                if b.button("拒绝",key=f"rj_{r['id']}"):
                    mutate_and_commit(data,lambda d: [x.update({"parent_status":"已拒绝","final_status":"未入账","approved_by":current_operator(),"approved_at":now_str(),"parent_note":note}) for x in d.get("pending_requests",[]) if x.get("id")==r["id"]],"拒绝申请",r.get("request_text",""))
    with tabs[1]:
        with st.form("bounty_form",clear_on_submit=True):
            title=st.text_input("任务标题"); desc=st.text_area("说明"); c1,c2,c3=st.columns(3); amount=c1.number_input("赏金",min_value=0.0,step=1.0); pts=c2.number_input("信用积分",min_value=0,max_value=30,step=1); diff=c3.selectbox("难度",["简单","普通","困难","长期"]); deadline=st.date_input("截止日",value=date.today()+timedelta(days=7)); ok=st.form_submit_button("发布",type="primary")
        if ok:
            bounty={"id":uid("bounty"),"title":title or "未命名任务","description":desc,"amount":amount,"reward_points":pts,"category":"家务/学习","difficulty":diff,"deadline":deadline.isoformat(),"status":"开放","created_at":now_str(),"created_by":current_operator(),"claimed_at":"","submitted_at":"","paid_at":"","submission_note":"","review_note":""}
            mutate_and_commit(data,lambda d: d.setdefault("bounties",[]).append(bounty),"发布赏金任务",title)
        st.dataframe(pd.DataFrame(data.get("bounties",[])),use_container_width=True,hide_index=True)
        for b in data.get("bounties",[]):
            if b.get("status") in {"开放","进行中","已过期","已关闭"}:
                c1,c2=st.columns(2)
                if c1.button(f"关闭：{b.get('title')}",key=f"close_{b['id']}"): mutate_and_commit(data,lambda d:[x.update({"status":"已关闭"}) for x in d.get("bounties",[]) if x.get("id")==b["id"]],"关闭赏金任务",b.get("title",""))
                if c2.button(f"删除未领取：{b.get('title')}",key=f"del_{b['id']}"): mutate_and_commit(data,lambda d:d.__setitem__("bounties",[x for x in d.get("bounties",[]) if not (x.get("id")==b["id"] and x.get("status") in {"开放","已过期","已关闭"})]),"删除赏金任务",b.get("title",""))
    with tabs[2]:
        reviews=[b for b in data.get("bounties",[]) if b.get("status")=="待审核"]
        if not reviews: st.info("暂无待审核任务。")
        for b in reviews:
            with st.container(border=True):
                st.write(f"**{b['title']}**｜{money(b['amount'],ccy)}｜+{b.get('reward_points')}分"); st.write(b.get("submission_note") or "无说明"); note=st.text_input("审核备注",key=f"rv_{b['id']}"); c1,c2=st.columns(2)
                if c1.button("通过并支付",type="primary",key=f"pay_{b['id']}"):
                    def mm(d):
                        task=next((x for x in d.get("bounties",[]) if x.get("id")==b["id"]),None)
                        if task: task.update({"status":"已支付","paid_at":now_str(),"review_note":note}); d.setdefault("transactions",[]).append(make_tx("赏金收入",task.get("amount"),"赏金任务",party=current_operator(),memo=task.get("title"),bounty_id=task.get("id")))
                    mutate_and_commit(data,mm,"赏金任务通过并支付",b["title"])
                if c2.button("退回",key=f"ret_{b['id']}"): mutate_and_commit(data,lambda d:[x.update({"status":"进行中","review_note":note}) for x in d.get("bounties",[]) if x.get("id")==b["id"]],"赏金任务退回",b["title"])
    with tabs[3]:
        text=st.text_area("输入消费/贷款请求",key="parent_ai")
        if st.button("生成评估",type="primary") and text.strip(): st.session_state["parent_decision"]=parse_request(data,text)
        if st.session_state.get("parent_decision"):
            render_decision(st.session_state["parent_decision"])
            if st.button("加入待审批队列"): create_pending(data,st.session_state["parent_decision"],text,current_operator())
    with tabs[4]:
        with st.form("reward_form",clear_on_submit=True):
            title=st.text_input("奖励券名称"); cost=st.number_input("兑换成本",min_value=0.0,step=1.0); req=st.number_input("最低信用分",min_value=0,max_value=850,value=700,step=10); note=st.text_area("说明"); ok=st.form_submit_button("创建",type="primary")
        if ok:
            reward={"id":uid("reward"),"title":title or "未命名奖励券","created_at":now_str(),"created_by":current_operator(),"cost":cost,"required_score":req,"status":"可用","used_at":"","note":note}
            mutate_and_commit(data,lambda d:d.setdefault("rewards",[]).append(reward),"创建奖励券",title)
        st.dataframe(pd.DataFrame(data.get("rewards",[])),use_container_width=True,hide_index=True)


def page_bank(data: Dict[str,Any]):
    if not require_parent(): return
    st.title("私人银行后台")
    m = calc_financials(data); ccy = m.get("currency", "USD")
    tabs = st.tabs([
        "总览", "新增交易", "AI决策中心", "消费审批", "贷款审批", "负债管理", "待审批队列",
        "风险雷达", "月度账单", "预算管理", "储蓄目标", "信用分", "商户权益", "周报",
        "奖励徽章", "赏金任务", "备份审计", "数据表"
    ])

    with tabs[0]:
        c1, c2, c3, c4 = st.columns(4)
        with c1: card("总资产", money(m["total_assets"], ccy), "现金 + 储蓄 + 应收贷款")
        with c2: card("净资产", money(m["net_assets"], ccy), "总资产 - 负债")
        with c3: card("现金余额", money(m["cash"], ccy), f"安全线 {money(data.get('settings', {}).get('cash_floor'), ccy)}")
        with c4: card("信用分", score_text(m["credit_score"]), "家庭内部风控评分")
        g1, g2 = st.columns(2)
        with g1:
            fig = asset_structure_chart(m)
            if fig is not None: st.plotly_chart(fig, use_container_width=True)
            else: st.dataframe(pd.DataFrame([{"项目":"现金","金额":m["cash"]},{"项目":"储蓄","金额":m["savings"]},{"项目":"应收贷款","金额":m["receivables"]}]), use_container_width=True, hide_index=True)
        with g2:
            fig = monthly_cashflow_chart(data)
            if fig is not None: st.plotly_chart(fig, use_container_width=True)
            else: st.dataframe(monthly_cashflow_df(data), use_container_width=True, hide_index=True)
        g3, g4 = st.columns(2)
        with g3:
            fig = budget_usage_chart(data, m)
            if fig is not None: st.plotly_chart(fig, use_container_width=True)
            else: st.info("设置预算后显示预算柱状图。")
        with g4:
            fig = expense_category_chart(data, m)
            if fig is not None: st.plotly_chart(fig, use_container_width=True)
            else: st.info("本月有消费后显示分类柱状图。")
        st.subheader("本周行动建议")
        for i, a in enumerate(build_weekly_plan(data), 1):
            st.write(f"{i}. {a}")

    with tabs[1]:
        st.subheader("新增交易")
        with st.form("tx_form_full", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            typ = c1.selectbox("类型", ["收入", "消费", "转入储蓄", "储蓄取出", "放贷", "还款", "借入", "偿还负债", "奖励兑换", "赏金收入"])
            amount = c2.number_input("金额", min_value=0.0, step=1.0)
            dte = c3.date_input("日期", value=date.today())
            cat = st.selectbox("分类", category_options(data), index=category_options(data).index("其他") if "其他" in category_options(data) else 0)
            party = st.text_input("对象/商户/对方")
            memo = st.text_area("备注")
            target_goal = ""
            if typ in {"转入储蓄", "储蓄取出"} and data.get("goals"):
                opts = {f"{g.get('name')}｜{money(g.get('current'), ccy)} / {money(g.get('target'), ccy)}": g.get("id") for g in data.get("goals", [])}
                target_goal = opts[st.selectbox("绑定储蓄目标", list(opts.keys()))]
            ok = st.form_submit_button("保存交易", type="primary")
        if ok:
            if typ == "放贷":
                st.warning("放贷请优先使用“贷款审批”标签，以建立贷款台账。")
            elif typ == "借入":
                debt = make_debt(amount, party or "未填写", "", memo)
                tx = make_tx("借入", amount, "负债", party=party, memo=memo, tx_date=dte.isoformat(), debt_id=debt["id"])
                mutate_and_commit(data, lambda d: apply_tx(d, tx, None, debt), "新增借入交易", memo)
            else:
                tx = make_tx(typ, amount, cat, party=party, memo=memo, tx_date=dte.isoformat(), goal_id=target_goal)
                mutate_and_commit(data, lambda d: apply_tx(d, tx), "新增交易", f"{typ} {money(amount, ccy)}")
        st.subheader("交易流水")
        df = transaction_display_df(data, ccy)
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.subheader("清除单笔交易")
        txs = sorted(data.get("transactions", []), key=lambda x: (str(x.get("date", "")), str(x.get("created_at", ""))), reverse=True)
        if txs:
            opts = {f"{x.get('date')}｜{x.get('type')}｜{money(x.get('amount'), ccy)}｜{x.get('category')}｜{x.get('memo','')[:24]}｜{x.get('id','')[-6:]}": x.get("id") for x in txs}
            sel = st.selectbox("选择要清除的交易", list(opts.keys()))
            confirm = st.checkbox("我确认清除该交易，并由系统重建贷款/负债台账")
            if st.button("清除这笔交易", type="primary", disabled=not confirm):
                mutate_and_commit(data, delete_transaction_mutation(opts[sel]), "清除单笔交易", sel)
        else:
            st.info("暂无交易。")

    with tabs[2]:
        st.subheader("AI 决策中心")
        text = st.text_area("自然语言输入", placeholder="例如：我想买 Roblox 皮肤 12 美元；或者：借给爸爸 20，预计还 22。", key="ai_center_text")
        if st.button("生成 AI + Python 风控评估", type="primary") and text.strip():
            st.session_state["full_ai_decision"] = parse_request(data, text)
        dec = st.session_state.get("full_ai_decision")
        if dec:
            render_decision(dec)
            st.text_area("给阿苏看的解释", value=local_decision_report(dec), height=130)
            if dec.get("kind") == "消费审批":
                tx = dec.get("pending_tx") or {}
                sdf = scenario_planning(data, fnum(tx.get("amount")), tx.get("category") or "其他", text, tx.get("party") or "")
                st.dataframe(pretty_df_money(sdf, ["交易后现金"], ccy), use_container_width=True, hide_index=True)
            col_a, col_b = st.columns(2)
            if col_a.button("加入待审批队列", key="ai_to_pending"):
                create_pending(data, dec, text, current_operator())
            if col_b.button("家长直接入账", key="ai_direct_book") and can_parent():
                mutate_and_commit(data, lambda d, dec=dec: apply_tx(d, dec.get("pending_tx"), dec.get("pending_loan"), dec.get("pending_debt")), "AI评估后直接入账", text)

    with tabs[3]:
        st.subheader("消费审批")
        with st.form("purchase_form_full"):
            c1, c2, c3 = st.columns(3)
            amount = c1.number_input("消费金额", min_value=0.0, step=1.0, key="purchase_amt_full")
            cat = c2.selectbox("分类", category_options(data), key="purchase_cat_full")
            merchant = c3.text_input("商户", key="purchase_merchant_full")
            desc = st.text_area("消费说明", key="purchase_desc_full")
            ok = st.form_submit_button("评估消费", type="primary")
        if ok:
            st.session_state["purchase_decision_full"] = evaluate_purchase(data, amount, cat, desc, merchant)
        dec = st.session_state.get("purchase_decision_full")
        if dec:
            render_decision(dec)
            st.dataframe(pretty_df_money(scenario_planning(data, amount, cat, desc, merchant), ["交易后现金"], ccy), use_container_width=True, hide_index=True)
            c1, c2 = st.columns(2)
            if c1.button("加入待审批", key="purchase_queue_full"):
                create_pending(data, dec, desc or "消费审批", current_operator())
            if c2.button("家长确认入账", key="purchase_book_full"):
                mutate_and_commit(data, lambda d, dec=dec: apply_tx(d, dec.get("pending_tx")), "消费审批通过并入账", desc)

    with tabs[4]:
        st.subheader("贷款审批 / 贷款台账")
        with st.form("loan_form_full", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            principal = c1.number_input("拟借出本金", min_value=0.0, step=1.0)
            borrower = c2.text_input("借款人", value="爸爸")
            expected = c3.number_input("预计回款", min_value=0.0, step=1.0)
            due = st.date_input("预计还款日", value=date.today() + timedelta(days=14))
            memo = st.text_area("贷款说明")
            ok = st.form_submit_button("评估贷款", type="primary")
        if ok:
            st.session_state["loan_decision_full"] = evaluate_loan(data, principal, borrower, expected or principal, due.isoformat(), memo)
        if st.session_state.get("loan_decision_full"):
            dec = st.session_state["loan_decision_full"]
            render_decision(dec)
            c1, c2 = st.columns(2)
            if c1.button("家长确认放贷", key="loan_direct_full"):
                mutate_and_commit(data, lambda d, dec=dec: apply_tx(d, dec.get("pending_tx"), dec.get("pending_loan"), None), "确认放贷", memo)
            if c2.button("加入审批队列", key="loan_queue_full"):
                create_pending(data, dec, memo or "贷款申请", current_operator())
        st.dataframe(pd.DataFrame(m.get("loans", [])), use_container_width=True, hide_index=True)
        st.subheader("绑定 loan_id 登记还款")
        open_loans = [x for x in m.get("loans", []) if fnum(x.get("remaining")) > 0]
        if open_loans:
            opts = {f"{x['borrower']}｜剩余 {money(x['remaining'], ccy)}｜到期 {x.get('due_date') or '无'}｜{x['id'][-6:]}": x["id"] for x in open_loans}
            sel = st.selectbox("选择贷款", list(opts.keys()))
            c1, c2 = st.columns(2)
            rp = c1.number_input("归还本金", min_value=0.0, step=1.0, key="loan_repay_principal_full")
            ri = c2.number_input("收到利息", min_value=0.0, step=0.5, key="loan_interest_full")
            memo2 = st.text_input("还款备注", key="loan_repay_memo_full")
            if st.button("保存还款", type="primary", key="loan_repay_save_full"):
                mutate_and_commit(data, repay_loan_mutation(opts[sel], rp, ri, memo2), "登记贷款还款", sel)
        else:
            st.info("暂无未结清贷款。")

    with tabs[5]:
        st.subheader("负债管理")
        with st.form("debt_form_full", clear_on_submit=True):
            c1, c2 = st.columns(2)
            amount = c1.number_input("借入金额", min_value=0.0, step=1.0)
            creditor = c2.text_input("债权人", value="爸爸")
            due = st.date_input("偿还日", value=date.today() + timedelta(days=14))
            memo = st.text_area("借入说明")
            ok = st.form_submit_button("登记借入", type="primary")
        if ok:
            debt = make_debt(amount, creditor, due.isoformat(), memo)
            tx = make_tx("借入", amount, "负债", party=creditor, memo=memo, debt_id=debt["id"], due_date=due.isoformat())
            mutate_and_commit(data, lambda d: apply_tx(d, tx, None, debt), "登记借入", memo)
        st.dataframe(pd.DataFrame(m.get("debts", [])), use_container_width=True, hide_index=True)
        open_debts = [x for x in m.get("debts", []) if fnum(x.get("remaining")) > 0]
        if open_debts:
            opts = {f"{x['creditor']}｜剩余 {money(x['remaining'], ccy)}｜到期 {x.get('due_date') or '无'}｜{x['id'][-6:]}": x["id"] for x in open_debts}
            sel = st.selectbox("选择负债", list(opts.keys()))
            amount = st.number_input("偿还金额", min_value=0.0, step=1.0, key="debtpay_full")
            memo = st.text_input("偿还备注", key="debtpay_memo_full")
            if st.button("保存偿还", type="primary"):
                mutate_and_commit(data, repay_debt_mutation(opts[sel], amount, memo), "偿还负债", sel)

    with tabs[6]:
        st.subheader("待审批队列")
        rows = sorted(data.get("pending_requests", []), key=lambda x: str(x.get("created_at", "")), reverse=True)
        status_filter = st.selectbox("状态筛选", ["全部", "待审批", "已批准", "已拒绝"])
        if status_filter != "全部":
            rows = [r for r in rows if r.get("parent_status") == status_filter]
        if rows:
            for r in rows:
                with st.container(border=True):
                    st.write(f"**{r.get('request_type')}｜{money(r.get('amount'), ccy)}｜{r.get('category')}｜{r.get('parent_status')}**")
                    st.write(r.get("request_text"))
                    st.caption(f"系统结论：{r.get('python_decision')}｜必要性：{r.get('necessity')}｜冲动：{r.get('impulse_level')}｜申请人：{r.get('applicant')}")
                    if r.get("decision_json"):
                        with st.expander("查看系统评估"):
                            st.json(r.get("decision_json"))
        else:
            st.info("暂无符合条件的申请。")

    with tabs[7]:
        st.subheader("风险雷达")
        risks = build_risk_radar(data)
        c1, c2, c3 = st.columns(3)
        for col, level in zip([c1, c2, c3], ["高风险", "中风险", "低风险"]):
            with col:
                st.markdown(f"### {level}")
                if not risks[level]: st.caption("无")
                for r in risks[level]:
                    with st.container(border=True):
                        st.write(f"**{r['title']}**")
                        st.caption(r["detail"])
        st.subheader("下周行动计划")
        for i, a in enumerate(build_weekly_plan(data), 1): st.write(f"{i}. {a}")

    with tabs[8]:
        st.subheader("月度账单")
        months = sorted({tx_month(t) for t in data.get("transactions", []) if tx_month(t)}, reverse=True) or [month_str()]
        mo = st.selectbox("月份", months)
        mm = calc_financials(data, mo)
        c1, c2, c3, c4 = st.columns(4)
        with c1: card("本月收入", money(mm["month_income"], ccy), mo)
        with c2: card("本月消费", money(mm["month_expense"], ccy), mo)
        with c3: card("支出收入比", percent(mm["spend_income_ratio"]), mo)
        with c4: card("本月利息收入", money(mm.get("month_interest", 0), ccy), mo)
        tx_df = pd.DataFrame([t for t in data.get("transactions", []) if tx_month(t) == mo])
        if not tx_df.empty: st.dataframe(tx_df, use_container_width=True, hide_index=True)
        else: st.info("该月暂无交易。")

    with tabs[9]:
        st.subheader("预算管理")
        left, right = st.columns([1, 1])
        with left:
            with st.form("budget_add_full", clear_on_submit=True):
                cat = st.text_input("预算分类", value="游戏")
                limit = st.number_input("月度预算", min_value=0.0, step=1.0)
                ok = st.form_submit_button("新增/更新预算", type="primary")
            if ok:
                mutate_and_commit(data, lambda d: d.setdefault("budgets", {}).__setitem__(cat.strip() or "其他", limit), "更新预算", cat)
            for cat, val in sorted(data.get("budgets", {}).items()):
                c1, c2, c3 = st.columns([2, 1, 1])
                c1.write(f"**{cat}**：{money(val, ccy)}")
                new_val = c2.number_input("调整", min_value=0.0, value=float(fnum(val)), step=1.0, key=f"budget_edit_{cat}")
                if c3.button("保存", key=f"budget_save_{cat}"):
                    mutate_and_commit(data, lambda d, c=cat, v=new_val: d.setdefault("budgets", {}).__setitem__(c, v), "调整预算", cat)
                if st.button(f"删除预算：{cat}", key=f"budget_delete_{cat}"):
                    mutate_and_commit(data, lambda d, c=cat: d.setdefault("budgets", {}).pop(c, None), "删除预算", cat)
        with right:
            fig = budget_usage_chart(data, m)
            if fig is not None: st.plotly_chart(fig, use_container_width=True)
            st.dataframe(pd.DataFrame([{"分类": k, **v} for k, v in m.get("budget_usage", {}).items()]), use_container_width=True, hide_index=True)

    with tabs[10]:
        st.subheader("储蓄目标")
        with st.form("goal_form_full", clear_on_submit=True):
            name = st.text_input("目标名称")
            target = st.number_input("目标金额", min_value=0.0, step=1.0)
            cur = st.number_input("当前金额", min_value=0.0, step=1.0)
            due = st.date_input("截止日", value=date.today() + timedelta(days=90))
            note = st.text_area("说明")
            ok = st.form_submit_button("新增目标", type="primary")
        if ok:
            goal = {"id": uid("goal"), "name": name or "未命名目标", "target": target, "current": cur, "deadline": due.isoformat(), "category": "储蓄", "note": note}
            mutate_and_commit(data, lambda d: d.setdefault("goals", []).append(goal), "新增储蓄目标", name)
        for g in m.get("goals", []):
            with st.container(border=True):
                st.write(f"**{g.get('name')}**｜{money(g.get('current'), ccy)} / {money(g.get('target'), ccy)}")
                st.progress(min(1.0, fnum(g.get("progress"))))
                c1, c2, c3 = st.columns(3)
                add_amt = c1.number_input("转入金额", min_value=0.0, step=1.0, key=f"goal_add_{g['id']}")
                take_amt = c2.number_input("取出金额", min_value=0.0, step=1.0, key=f"goal_take_{g['id']}")
                if c3.button("删除目标", key=f"goal_del_{g['id']}"):
                    mutate_and_commit(data, lambda d, gid=g['id']: d.__setitem__("goals", [x for x in d.get("goals", []) if x.get("id") != gid]), "删除储蓄目标", g.get("name", ""))
                if st.button("转入该目标", key=f"goal_add_btn_{g['id']}"):
                    tx = make_tx("转入储蓄", add_amt, "储蓄", memo=f"转入目标：{g.get('name')}", goal_id=g["id"])
                    mutate_and_commit(data, lambda d, tx=tx: apply_tx(d, tx), "转入储蓄目标", g.get("name", ""))
                if st.button("从该目标取出", key=f"goal_take_btn_{g['id']}"):
                    tx = make_tx("储蓄取出", take_amt, "储蓄", memo=f"从目标取出：{g.get('name')}", goal_id=g["id"])
                    mutate_and_commit(data, lambda d, tx=tx: apply_tx(d, tx), "储蓄目标取出", g.get("name", ""))

    with tabs[11]:
        st.subheader("信用分")
        c1, c2 = st.columns([1, 1])
        with c1:
            card("当前信用分", score_text(m["credit_score"]), "300–850 区间")
            st.dataframe(build_scorecard(data), use_container_width=True, hide_index=True)
            st.subheader("信用分因子")
            for f in m.get("score_factors", []): st.write(f"- {f}")
        with c2:
            fig = score_trend_chart(data)
            if fig is not None: st.plotly_chart(fig, use_container_width=True)
            else: st.info("产生信用分历史后显示趋势图。")
        st.subheader("信用分历史")
        st.dataframe(pd.DataFrame(data.get("score_history", [])), use_container_width=True, hide_index=True)

    with tabs[12]:
        st.subheader("商户权益")
        with st.form("merchant_form_full", clear_on_submit=True):
            name = st.text_input("商户名称")
            cat = st.selectbox("分类", category_options(data), key="merchant_cat_full")
            discount = st.number_input("折扣", min_value=0.0, max_value=1.0, step=0.01)
            required = st.number_input("最低信用分", min_value=0, max_value=850, value=700, step=10)
            cap = st.number_input("预算使用率上限", min_value=0.0, max_value=2.0, value=1.0, step=0.05)
            note = st.text_area("说明")
            ok = st.form_submit_button("新增商户", type="primary")
        if ok:
            merchant = {"id": uid("merchant"), "name": name or "未命名商户", "category": cat, "discount": discount, "required_score": required, "category_budget_cap": cap, "note": note}
            mutate_and_commit(data, lambda d: d.setdefault("merchants", []).append(merchant), "新增商户", name)
        rows = []
        for mer in data.get("merchants", []):
            res = evaluate_merchant_access(data, mer)
            rows.append({"商户": mer.get("name"), "分类": mer.get("category"), "折扣": f"{fnum(mer.get('discount'))*100:.0f}%", "最低信用分": mer.get("required_score"), "状态": res["result"], "说明": res["message"], "未通过项": "；".join(res["failed"])})
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    with tabs[13]:
        st.subheader("周报")
        if st.button("生成本周周报", type="primary", key="weekly_generate_full"):
            rep = generate_weekly_report(data)
            mutate_and_commit(data, lambda d: d.setdefault("weekly_reports", []).append(rep), "生成周报", rep["week_start"])
        if data.get("weekly_reports"):
            latest = sorted(data.get("weekly_reports", []), key=lambda x: str(x.get("generated_at", "")), reverse=True)[0]
            st.text_area("最新周报", value=latest.get("content", ""), height=220)
        st.dataframe(pd.DataFrame(data.get("weekly_reports", [])), use_container_width=True, hide_index=True)

    with tabs[14]:
        st.subheader("奖励徽章")
        st.markdown("### 徽章墙")
        st.markdown(" ".join([f"<span class='badge'>{b.get('name')}</span>" for b in sorted(data.get("badges", []), key=lambda x: str(x.get("earned_at", "")), reverse=True)]) or "暂无徽章", unsafe_allow_html=True)
        st.dataframe(pd.DataFrame(data.get("badges", [])), use_container_width=True, hide_index=True)
        st.markdown("### 奖励券")
        with st.form("reward_form_bank", clear_on_submit=True):
            title = st.text_input("奖励券名称")
            cost = st.number_input("兑换成本", min_value=0.0, step=1.0)
            req_score = st.number_input("最低信用分", min_value=0, max_value=850, value=700, step=10)
            note = st.text_area("说明")
            ok = st.form_submit_button("创建奖励券", type="primary")
        if ok:
            reward = {"id": uid("reward"), "title": title or "未命名奖励券", "created_at": now_str(), "created_by": current_operator(), "cost": cost, "required_score": req_score, "status": "可用", "used_at": "", "note": note}
            mutate_and_commit(data, lambda d: d.setdefault("rewards", []).append(reward), "创建奖励券", title)
        st.dataframe(pd.DataFrame(data.get("rewards", [])), use_container_width=True, hide_index=True)

    with tabs[15]:
        st.subheader("赏金任务")
        st.write("任务完整闭环：发布 → 领取 → 提交 → 审核 → 支付 → 入账 → 信用积分。")
        with st.form("bounty_form_bank", clear_on_submit=True):
            title = st.text_input("任务标题")
            desc = st.text_area("任务说明")
            c1, c2, c3 = st.columns(3)
            amount = c1.number_input("赏金", min_value=0.0, step=1.0)
            pts = c2.number_input("信用积分", min_value=0, max_value=30, step=1)
            diff = c3.selectbox("难度", ["简单", "普通", "困难", "长期"])
            deadline = st.date_input("截止日", value=date.today() + timedelta(days=7), key="bank_bounty_deadline")
            ok = st.form_submit_button("发布任务", type="primary")
        if ok:
            bounty = {"id": uid("bounty"), "title": title or "未命名任务", "description": desc, "amount": amount, "reward_points": pts, "category": "家务/学习", "difficulty": diff, "deadline": deadline.isoformat(), "status": "开放", "created_at": now_str(), "created_by": current_operator(), "claimed_at": "", "submitted_at": "", "paid_at": "", "submission_note": "", "review_note": ""}
            mutate_and_commit(data, lambda d: d.setdefault("bounties", []).append(bounty), "发布赏金任务", title)
        st.dataframe(pd.DataFrame(data.get("bounties", [])), use_container_width=True, hide_index=True)

    with tabs[16]:
        st.subheader("备份与审计")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("### 备份")
            st.dataframe(pd.DataFrame(data.get("backups", [])), use_container_width=True, hide_index=True)
        with c2:
            st.markdown("### 审计日志")
            st.dataframe(pd.DataFrame(data.get("audit_log", [])), use_container_width=True, hide_index=True)
        add_download_buttons(data)

    with tabs[17]:
        st.subheader("全量数据表")
        for name, df in table_dfs(data).items():
            with st.expander(name, expanded=(name == "summary")):
                st.dataframe(df, use_container_width=True, hide_index=True)

def page_settings(data: Dict[str,Any]):
    if not require_parent(): return
    st.title("系统设置"); tabs=st.tabs(["账户规则","Google Sheet","导出","备份恢复","审计","危险操作"])
    with tabs[0]:
        s=data.get("settings",{})
        with st.form("settings"):
            owner=st.text_input("账户所有人",value=s.get("owner","阿苏")); currency=st.selectbox("币种",["USD","CNY"],index=0 if s.get("currency","USD")=="USD" else 1); c1,c2,c3=st.columns(3); start_cash=c1.number_input("初始现金",value=float(fnum(s.get("start_cash"))),step=1.0); start_savings=c2.number_input("初始储蓄",value=float(fnum(s.get("start_savings"))),step=1.0); base_score=c3.number_input("基础信用分",min_value=0,max_value=850,value=inum(s.get("base_score"),700),step=10); c4,c5=st.columns(2); cash_floor=c4.number_input("现金安全线",value=float(fnum(s.get("cash_floor"))),step=1.0); approval=c5.number_input("审批金额线",value=float(fnum(s.get("approval_threshold"))),step=1.0); ok=st.form_submit_button("保存",type="primary")
        if ok: mutate_and_commit(data,lambda d:d["settings"].update({"owner":owner,"currency":currency,"start_cash":start_cash,"start_savings":start_savings,"base_score":base_score,"cash_floor":cash_floor,"approval_threshold":approval}),"更新账户设置","")
        st.subheader("风控规则"); edits={}
        for k,v in data.get("rules",{}).items(): edits[k]=st.number_input(k,value=float(fnum(v)),step=0.01 if "line" in k or "limit" in k or "ratio" in k else 1.0,key=f"rule_{k}")
        if st.button("保存规则",type="primary"): mutate_and_commit(data,lambda d:d["rules"].update(edits),"更新风控规则","")
    with tabs[1]:
        st.info("日常保存只写 state；展示页手动刷新，避免每次操作都慢。")
        st.write(f"当前存储：{st.session_state.get('storage_backend','未知')}")
        if st.button("手动刷新 Google Sheet 展示页",type="primary"):
            try: sync_report_sheets(data); st.success("已刷新。")
            except Exception as e: st.error(f"失败：{e}")
    with tabs[2]:
        try: st.download_button("下载 Excel 总账包",data=excel_bytes(data),file_name=f"asu_private_bank_{today_str()}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e: st.error(f"Excel 导出失败：{e}")
        st.download_button("下载 JSON",data=json.dumps(data,ensure_ascii=False,indent=2),file_name=f"asu_private_bank_{today_str()}.json",mime="application/json")
    with tabs[3]:
        backups=data.get("backups",[])
        if not backups: st.info("暂无备份。")
        else:
            opts={f"{b.get('time')}｜{b.get('operator')}｜{b.get('event')}｜{b.get('id','')[-6:]}":b for b in reversed(backups)}; sel=st.selectbox("选择备份",list(opts.keys())); b=opts[sel]
            if st.button("恢复该备份",type="primary"):
                snap=safe_load_json(b.get("snapshot_json",""))
                if snap: restored=normalize_data(snap); restored["backups"]=data.get("backups",[]); restored["audit_log"]=data.get("audit_log",[]); commit(restored,"恢复备份",sel)
                else: st.error("备份不可解析。")
    with tabs[4]:
        st.dataframe(pd.DataFrame(data.get("audit_log",[])),use_container_width=True,hide_index=True); st.dataframe(pd.DataFrame(data.get("score_history",[])),use_container_width=True,hide_index=True)
    with tabs[5]:
        st.warning("输入 RESET 才能清空。"); confirm=st.text_input("确认词")
        if st.button("清空重建"):
            if confirm=="RESET": commit(empty_data(),"清空数据","RESET")
            else: st.error("未输入 RESET。")

def main():
    setup_page(); require_login(); data=get_data(); sidebar(data); page=st.sidebar.radio("主入口",["阿苏首页","家长工作台","私人银行后台","系统设置"])
    if page=="阿苏首页": page_asu(data)
    elif page=="家长工作台": page_parent(data)
    elif page=="私人银行后台": page_bank(data)
    else: page_settings(data)

if __name__ == "__main__":
    main()
