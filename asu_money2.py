# -*- coding: utf-8 -*-
"""
阿苏私人银行 2.0 — 单文件完全重写版

定位：
    不是普通记账 App，而是一个家庭版“私人银行风控系统”。

运行：
    streamlit run asu_money2.py

可选 DeepSeek：
    PowerShell:
        setx DEEPSEEK_API_KEY "你的key"

设计原则：
    1. Python 负责所有数字：资产、净资产、预算、信用分、风险等级、审批结论。
    2. DeepSeek 只负责把 Python 算好的指标解释成中文客户经理报告，不允许编数字。
    3. 没有 DeepSeek API Key 时，使用本地中文风控报告，程序照常运行。
    4. 全部写在一个文件里。数据自动保存到 asu_money2_data.json。
"""

from __future__ import annotations

import copy
import csv
import io
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


APP_NAME = "阿苏私人银行 2.0"
DATA_PATH = Path("asu_money2_data.json")


# ============================================================
# 1. 默认数据
# ============================================================

def today_str() -> str:
    return date.today().isoformat()


def month_str(d: Optional[date] = None) -> str:
    return (d or date.today()).strftime("%Y-%m")


DEFAULT_DATA: Dict[str, Any] = {
    "settings": {
        "owner": "阿苏",
        "currency": "USD",
        "start_cash": 42.00,
        "start_savings": 60.00,
        "base_score": 742,
        "cash_floor": 15.00,
        "loan_asset_limit": 0.45,
        "hard_loan_asset_limit": 0.65,
        "created_at": datetime.now().isoformat(timespec="seconds"),
    },
    "budgets": {
        "游戏": 25.00,
        "甜品": 18.00,
        "玩具": 30.00,
        "学习": 45.00,
        "宠物": 25.00,
        "餐饮": 35.00,
        "其他": 30.00,
    },
    "goals": [
        {
            "id": "goal-001",
            "name": "Minecraft 年度基金",
            "target": 80.00,
            "current": 18.00,
            "deadline": (date.today() + timedelta(days=90)).isoformat(),
            "category": "游戏",
        },
        {
            "id": "goal-002",
            "name": "宠物用品储备金",
            "target": 60.00,
            "current": 20.00,
            "deadline": (date.today() + timedelta(days=120)).isoformat(),
            "category": "宠物",
        },
    ],
    "merchants": [
        {
            "id": "m-001",
            "name": "Minecraft 商店",
            "category": "游戏",
            "discount": 0.08,
            "required_score": 720,
            "category_budget_cap": 0.90,
            "note": "游戏预算未接近上限时开放。",
        },
        {
            "id": "m-002",
            "name": "甜品小店",
            "category": "甜品",
            "discount": 0.10,
            "required_score": 735,
            "category_budget_cap": 0.80,
            "note": "甜品属于高冲动消费，需要更严格预算纪律。",
        },
        {
            "id": "m-003",
            "name": "学习用品店",
            "category": "学习",
            "discount": 0.12,
            "required_score": 700,
            "category_budget_cap": 1.00,
            "note": "学习类支出可以适度放宽。",
        },
        {
            "id": "m-004",
            "name": "宠物用品店",
            "category": "宠物",
            "discount": 0.06,
            "required_score": 710,
            "category_budget_cap": 0.95,
            "note": "宠物刚需，但仍要看现金余额。",
        },
    ],
    "transactions": [
        {
            "id": "seed-001",
            "date": date.today().replace(day=1).isoformat(),
            "type": "收入",
            "amount": 50.00,
            "category": "零花钱",
            "party": "",
            "memo": "本月零花钱",
            "expected_repayment": 0.00,
            "principal_repaid": 0.00,
            "interest_received": 0.00,
            "due_date": "",
        },
        {
            "id": "seed-002",
            "date": today_str(),
            "type": "消费",
            "amount": 7.50,
            "category": "甜品",
            "party": "甜品小店",
            "memo": "冰淇淋",
            "expected_repayment": 0.00,
            "principal_repaid": 0.00,
            "interest_received": 0.00,
            "due_date": "",
        },
        {
            "id": "seed-003",
            "date": today_str(),
            "type": "放贷",
            "amount": 20.00,
            "category": "家庭贷款",
            "party": "爸爸",
            "memo": "借给爸爸，预计一周后还 22",
            "expected_repayment": 22.00,
            "principal_repaid": 0.00,
            "interest_received": 0.00,
            "due_date": (date.today() + timedelta(days=7)).isoformat(),
        },
    ],
}


# ============================================================
# 2. 通用工具
# ============================================================

def uid() -> str:
    return str(uuid.uuid4())


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
    val = fnum(x)
    if currency == "USD":
        return f"${val:,.2f}"
    return f"{currency} {val:,.2f}"


def percent(x: Any) -> str:
    return f"{fnum(x) * 100:.0f}%"


def tx_month(tx: Dict[str, Any]) -> str:
    return str(tx.get("date", ""))[:7]


def css_class_by_decision(result: str) -> str:
    if result in ["批准", "通过", "开放"]:
        return "ok"
    if result in ["延迟", "限额通过", "暂缓开放", "观察"]:
        return "warn"
    return "bad"


def deep_merge(default: Dict[str, Any], incoming: Dict[str, Any]) -> Dict[str, Any]:
    out = copy.deepcopy(default)
    for k, v in (incoming or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = deep_merge(out[k], v)
        else:
            out[k] = v
    return out


def normalize_data(data: Dict[str, Any]) -> Dict[str, Any]:
    data = deep_merge(DEFAULT_DATA, data or {})

    # budget 兼容：只保留数字
    normalized_budgets = {}
    for k, v in data.get("budgets", {}).items():
        if isinstance(v, dict):
            normalized_budgets[k] = fnum(v.get("limit"), 0)
        else:
            normalized_budgets[k] = fnum(v, 0)
    data["budgets"] = normalized_budgets

    # transactions 兼容旧字段
    normalized_txs = []
    for tx in data.get("transactions", []):
        normalized_txs.append({
            "id": tx.get("id") or uid(),
            "date": tx.get("date") or today_str(),
            "type": tx.get("type") or "消费",
            "amount": fnum(tx.get("amount"), 0),
            "category": tx.get("category") or "其他",
            "party": tx.get("party") or tx.get("counterparty") or "",
            "memo": tx.get("memo") or "",
            "expected_repayment": fnum(tx.get("expected_repayment"), 0),
            "principal_repaid": fnum(tx.get("principal_repaid"), 0),
            "interest_received": fnum(tx.get("interest_received"), 0),
            "due_date": tx.get("due_date") or "",
        })
    data["transactions"] = normalized_txs

    for g in data.get("goals", []):
        g["id"] = g.get("id") or uid()
        g["target"] = fnum(g.get("target"), 0)
        g["current"] = fnum(g.get("current"), 0)
        g["deadline"] = g.get("deadline") or ""

    for m in data.get("merchants", []):
        m["id"] = m.get("id") or uid()
        m["discount"] = fnum(m.get("discount"), 0)
        m["required_score"] = inum(m.get("required_score"), 700)
        m["category_budget_cap"] = fnum(m.get("category_budget_cap"), 0.9)

    return data


def load_data() -> Dict[str, Any]:
    if DATA_PATH.exists():
        try:
            return normalize_data(json.loads(DATA_PATH.read_text(encoding="utf-8")))
        except Exception:
            broken = DATA_PATH.with_suffix(".broken.json")
            DATA_PATH.replace(broken)
            return normalize_data(copy.deepcopy(DEFAULT_DATA))
    return normalize_data(copy.deepcopy(DEFAULT_DATA))


def save_data(data: Dict[str, Any]) -> None:
    DATA_PATH.write_text(json.dumps(normalize_data(data), ensure_ascii=False, indent=2), encoding="utf-8")


def get_data() -> Dict[str, Any]:
    if "asu_bank_data" not in st.session_state:
        st.session_state.asu_bank_data = load_data()
    return st.session_state.asu_bank_data


def commit(data: Dict[str, Any]) -> None:
    st.session_state.asu_bank_data = normalize_data(data)
    save_data(st.session_state.asu_bank_data)


def new_tx(
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
# 3. 核心账务引擎
# ============================================================

def build_loan_book(transactions: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    应收贷款台账：
    - 放贷：现金减少，应收本金增加。
    - 还款：本金回收减少应收本金；利息计入收入。
    """
    loans: List[Dict[str, Any]] = []

    sorted_txs = sorted(transactions, key=lambda x: (str(x.get("date", "")), str(x.get("id", ""))))

    for tx in sorted_txs:
        if tx.get("type") == "放贷":
            principal = fnum(tx.get("amount"), 0)
            if principal <= 0:
                continue
            loans.append({
                "loan_id": tx.get("id"),
                "date": tx.get("date"),
                "borrower": tx.get("party") or "未填写",
                "principal": principal,
                "repaid": 0.0,
                "remaining": principal,
                "expected_repayment": fnum(tx.get("expected_repayment"), 0),
                "due_date": tx.get("due_date") or "",
                "memo": tx.get("memo") or "",
            })

        elif tx.get("type") == "还款":
            borrower = tx.get("party") or ""
            principal_repaid = fnum(tx.get("principal_repaid"), 0)
            if principal_repaid <= 0:
                principal_repaid = fnum(tx.get("amount"), 0)

            # 先匹配同一借款人，再按最早贷款冲抵
            candidates = [
                loan for loan in loans
                if loan["remaining"] > 0 and (not borrower or loan["borrower"] == borrower)
            ]
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
    usage = {
        category: {"limit": fnum(limit), "spent": 0.0, "ratio": 0.0}
        for category, limit in data.get("budgets", {}).items()
    }

    for tx in data.get("transactions", []):
        if tx.get("type") != "消费":
            continue
        if tx_month(tx) != month:
            continue

        cat = tx.get("category") or "其他"
        usage.setdefault(cat, {"limit": 0.0, "spent": 0.0, "ratio": 0.0})
        usage[cat]["spent"] += fnum(tx.get("amount"), 0)

    for cat, row in usage.items():
        limit = fnum(row.get("limit"), 0)
        row["ratio"] = row["spent"] / limit if limit > 0 else 0.0

    return usage


def calc_goals(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    out = []
    now = date.today()
    for g in data.get("goals", []):
        target = fnum(g.get("target"), 0)
        current = fnum(g.get("current"), 0)
        progress = current / target if target > 0 else 0.0
        days_left = None
        if g.get("deadline"):
            try:
                days_left = (date.fromisoformat(g["deadline"]) - now).days
            except Exception:
                days_left = None
        out.append({
            **g,
            "progress": progress,
            "remaining": max(0.0, target - current),
            "days_left": days_left,
        })
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
) -> Tuple[int, List[str]]:
    settings = data.get("settings", {})
    score = fnum(settings.get("base_score"), 742)
    factors: List[str] = []

    cash_floor = fnum(settings.get("cash_floor"), 15)

    # 现金缓冲
    if cash < 0:
        score -= 45
        factors.append("现金余额为负：-45")
    elif cash < cash_floor:
        score -= 16
        factors.append(f"现金低于安全线 {money(cash_floor)}：-16")
    elif cash >= cash_floor * 2:
        score += 6
        factors.append("现金缓冲较充足：+6")

    # 储蓄比例
    savings_ratio = savings / total_assets if total_assets > 0 else 0
    if savings_ratio >= 0.30:
        score += 8
        factors.append("储蓄占总资产 30% 以上：+8")
    elif total_assets > 0 and savings_ratio < 0.10:
        score -= 8
        factors.append("储蓄占总资产低于 10%：-8")

    # 支出收入比
    if month_income <= 0 and month_expense > 0:
        score -= 12
        factors.append("本月有消费但没有收入记录：-12")
    elif month_income > 0:
        ratio = month_expense / month_income
        if ratio > 1.00:
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

    # 贷款资产比例
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

    # 负债比例
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

    # 预算纪律
    overspent = sum(1 for row in budget_usage.values() if row["limit"] > 0 and row["ratio"] > 1)
    warning = sum(1 for row in budget_usage.values() if row["limit"] > 0 and 0.85 < row["ratio"] <= 1)

    if overspent:
        penalty = min(32, overspent * 10)
        score -= penalty
        factors.append(f"{overspent} 个预算品类已超支：-{penalty}")
    if warning:
        penalty = min(15, warning * 4)
        score -= penalty
        factors.append(f"{warning} 个预算品类接近上限：-{penalty}")

    # 逾期贷款
    if overdue_principal > 0:
        score -= 24
        factors.append(f"存在逾期未收回应收本金 {money(overdue_principal)}：-24")

    final_score = int(round(clamp(score, 300, 850)))
    if not factors:
        factors.append("暂无显著加分或扣分因素。")

    return final_score, factors


def calc_financials(data: Dict[str, Any], month: Optional[str] = None) -> Dict[str, Any]:
    month = month or month_str()
    settings = data.get("settings", {})
    currency = settings.get("currency", "USD")

    cash = fnum(settings.get("start_cash"), 0)
    savings = fnum(settings.get("start_savings"), 0)
    liabilities = 0.0
    month_income = 0.0
    month_expense = 0.0
    total_income = 0.0
    total_expense = 0.0

    for tx in sorted(data.get("transactions", []), key=lambda x: str(x.get("date", ""))):
        typ = tx.get("type")
        amt = fnum(tx.get("amount"), 0)
        is_month = tx_month(tx) == month

        if typ == "收入":
            cash += amt
            total_income += amt
            if is_month:
                month_income += amt

        elif typ == "消费":
            cash -= amt
            total_expense += amt
            if is_month:
                month_expense += amt

        elif typ == "转入储蓄":
            cash -= amt
            savings += amt

        elif typ == "储蓄取出":
            cash += amt
            savings -= amt

        elif typ == "放贷":
            cash -= amt

        elif typ == "还款":
            principal = fnum(tx.get("principal_repaid"), 0)
            if principal <= 0:
                principal = amt
            interest = fnum(tx.get("interest_received"), 0)
            cash += principal + interest
            total_income += interest
            if is_month:
                month_income += interest

        elif typ == "借入":
            cash += amt
            liabilities += amt

        elif typ == "偿还负债":
            cash -= amt
            liabilities -= amt

    liabilities = max(0.0, liabilities)
    loans = build_loan_book(data.get("transactions", []))
    receivables = sum(fnum(x.get("remaining"), 0) for x in loans)

    today = date.today()
    overdue_principal = 0.0
    for loan in loans:
        due = loan.get("due_date") or ""
        rem = fnum(loan.get("remaining"), 0)
        if rem <= 0 or not due:
            continue
        try:
            if date.fromisoformat(due) < today:
                overdue_principal += rem
        except Exception:
            pass

    total_assets = cash + savings + receivables
    net_assets = total_assets - liabilities
    budget_usage = calc_budget_usage(data, month)
    goals = calc_goals(data)

    score, score_factors = calc_credit_score(
        data=data,
        cash=cash,
        savings=savings,
        receivables=receivables,
        liabilities=liabilities,
        total_assets=total_assets,
        month_income=month_income,
        month_expense=month_expense,
        budget_usage=budget_usage,
        overdue_principal=overdue_principal,
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
        "credit_score": score,
        "score_factors": score_factors,
    }


# ============================================================
# 4. 模拟与审批引擎
# ============================================================

def simulate_transaction(data: Dict[str, Any], tx: Dict[str, Any]) -> Dict[str, Any]:
    before = calc_financials(data)
    after_data = copy.deepcopy(data)
    after_data.setdefault("transactions", []).append(tx)
    after = calc_financials(after_data)

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
            "credit_score": after["credit_score"] - before["credit_score"],
        },
        "tx": tx,
    }


def estimate_score_after_transaction(data: Dict[str, Any], tx: Dict[str, Any]) -> Dict[str, int]:
    sim = simulate_transaction(data, tx)
    return {
        "current_score": sim["before"]["credit_score"],
        "after_score": sim["after"]["credit_score"],
        "change": sim["delta"]["credit_score"],
    }


def category_usage(metrics: Dict[str, Any], category: str) -> Dict[str, float]:
    return metrics.get("budget_usage", {}).get(category, {"limit": 0.0, "spent": 0.0, "ratio": 0.0})


def evaluate_purchase_decision(
    data: Dict[str, Any],
    amount: float,
    category: str,
    description: str,
    merchant: str = "",
) -> Dict[str, Any]:
    tx = new_tx(
        tx_type="消费",
        amount=amount,
        category=category,
        party=merchant,
        memo=description,
    )
    sim = simulate_transaction(data, tx)
    before = sim["before"]
    after = sim["after"]
    delta = sim["delta"]
    budget = category_usage(after, category)
    cash_floor = fnum(data.get("settings", {}).get("cash_floor"), 15)

    result = "批准"
    reasons: List[str] = []

    # 硬拒绝
    if amount <= 0:
        result = "拒绝"
        reasons.append("购买金额必须大于 0。")
    if after["cash"] < 0:
        result = "拒绝"
        reasons.append("购买后现金余额为负，触发硬性拒绝。")
    if budget["limit"] > 0 and budget["ratio"] > 1.20:
        result = "拒绝"
        reasons.append(f"{category}预算使用率将达到 {percent(budget['ratio'])}，超过 120% 红线。")
    if after["credit_score"] < 650:
        result = "拒绝"
        reasons.append("交易后信用分低于 650，进入高风险区。")

    # 软拒绝 / 延迟
    if result != "拒绝":
        delay_flags = []
        if after["cash"] < cash_floor:
            delay_flags.append(f"购买后现金余额低于安全线 {money(cash_floor)}。")
        if budget["limit"] > 0 and budget["ratio"] >= 0.90:
            delay_flags.append(f"{category}预算使用率将达到 {percent(budget['ratio'])}，接近或超过上限。")
        if delta["credit_score"] <= -8:
            delay_flags.append(f"信用分预计下降 {abs(delta['credit_score'])} 分。")
        if after["spend_income_ratio"] >= 0.85:
            delay_flags.append(f"本月支出/收入比将达到 {percent(after['spend_income_ratio'])}。")

        if delay_flags:
            result = "延迟"
            reasons.extend(delay_flags)

    if result == "批准":
        reasons.append("现金余额、预算使用率、信用分变化均处在可接受区间。")

    if result == "拒绝":
        action = "不要购买。先恢复现金缓冲或等待下月预算重置。"
    elif result == "延迟":
        action = "建议延迟到下一笔收入到账后再买，或者把金额拆成两周预算。"
    else:
        action = "可以购买，但需要入账；非必要消费不要动用储蓄目标资金。"

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
        "simulation": sim,
        "pending_tx": tx,
    }


def evaluate_loan_decision(
    data: Dict[str, Any],
    principal: float,
    borrower: str,
    expected_repayment: float,
    due_date: str,
    memo: str = "",
) -> Dict[str, Any]:
    tx = new_tx(
        tx_type="放贷",
        amount=principal,
        category="家庭贷款",
        party=borrower,
        memo=memo,
        expected_repayment=expected_repayment,
        due_date=due_date,
    )
    sim = simulate_transaction(data, tx)
    before = sim["before"]
    after = sim["after"]
    delta = sim["delta"]

    settings = data.get("settings", {})
    cash_floor = fnum(settings.get("cash_floor"), 15)
    soft_limit = fnum(settings.get("loan_asset_limit"), 0.45)
    hard_limit = fnum(settings.get("hard_loan_asset_limit"), 0.65)

    result = "通过"
    reasons: List[str] = []

    max_by_cash = max(0.0, before["cash"] - cash_floor)
    safe_receivable_capacity = max(0.0, soft_limit * max(before["total_assets"], 1.0) - before["receivables"])
    suggested_limit = max(0.0, min(principal, max_by_cash, safe_receivable_capacity))

    if principal <= 0:
        result = "拒绝"
        reasons.append("放贷本金必须大于 0。")
    if after["cash"] < 0:
        result = "拒绝"
        reasons.append("放贷后现金余额为负，流动性不足。")
    if after["loan_asset_ratio"] > hard_limit:
        result = "拒绝"
        reasons.append(f"放贷后应收贷款本金占总资产 {percent(after['loan_asset_ratio'])}，超过 {percent(hard_limit)} 红线。")
    if after["credit_score"] < 650:
        result = "拒绝"
        reasons.append("放贷后信用分低于 650。")

    if result != "拒绝":
        flags = []
        if after["cash"] < cash_floor:
            flags.append(f"放贷后现金低于安全线 {money(cash_floor)}。")
        if after["loan_asset_ratio"] > soft_limit:
            flags.append(f"放贷后应收贷款本金占总资产 {percent(after['loan_asset_ratio'])}，超过建议线 {percent(soft_limit)}。")
        if principal > suggested_limit and suggested_limit > 0:
            flags.append(f"建议最高放贷金额为 {money(suggested_limit)}。")
        if not due_date:
            flags.append("没有填写预计还款日，回款纪律不足。")

        if flags:
            result = "限额通过"
            reasons.extend(flags)

    if result == "通过":
        reasons.append("放贷后现金缓冲、贷款资产占比、信用分均处于可接受区间。")

    if result == "拒绝":
        action = "本次不建议放贷。先收回应收本金，或降低放贷金额。"
    elif result == "限额通过":
        action = f"建议限额放贷，最高不超过 {money(suggested_limit)}；必须写清还款日。"
    else:
        action = "可以放贷，但回款本金到账后应优先补充现金余额。"

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
        "simulation": sim,
        "pending_tx": tx,
    }


def build_risk_radar(data: Dict[str, Any]) -> Dict[str, List[Dict[str, str]]]:
    m = calc_financials(data)
    ccy = m["currency"]
    risks = {"高风险": [], "中风险": [], "低风险": []}

    if m["cash"] < 0:
        risks["高风险"].append({"title": "现金余额为负", "detail": f"当前现金 {money(m['cash'], ccy)}，应暂停所有非必要消费。"})
    elif m["cash"] < fnum(data["settings"].get("cash_floor"), 15):
        risks["中风险"].append({"title": "现金缓冲偏低", "detail": f"当前现金 {money(m['cash'], ccy)}，低于安全线。"})
    else:
        risks["低风险"].append({"title": "现金余额正常", "detail": f"当前现金 {money(m['cash'], ccy)}。"})


    if m["month_income"] > 0:
        if m["spend_income_ratio"] > 1:
            risks["高风险"].append({"title": "本月支出超过收入", "detail": f"支出/收入比 {percent(m['spend_income_ratio'])}。"})
        elif m["spend_income_ratio"] > 0.85:
            risks["中风险"].append({"title": "本月支出接近收入", "detail": f"支出/收入比 {percent(m['spend_income_ratio'])}。"})
        else:
            risks["低风险"].append({"title": "支出收入比可控", "detail": f"支出/收入比 {percent(m['spend_income_ratio'])}。"})

    if m["loan_asset_ratio"] > 0.60:
        risks["高风险"].append({"title": "贷款资产占比过高", "detail": f"应收贷款本金占总资产 {percent(m['loan_asset_ratio'])}，流动性弱。"})
    elif m["loan_asset_ratio"] > 0.45:
        risks["中风险"].append({"title": "贷款资产占比偏高", "detail": f"应收贷款本金占总资产 {percent(m['loan_asset_ratio'])}，建议收回部分本金。"})
    else:
        risks["低风险"].append({"title": "贷款资产占比正常", "detail": f"应收贷款本金占总资产 {percent(m['loan_asset_ratio'])}。"})

    if m["liabilities"] > 0:
        if m["debt_asset_ratio"] > 0.35:
            risks["高风险"].append({"title": "负债偏高", "detail": f"负债占总资产 {percent(m['debt_asset_ratio'])}。"})
        else:
            risks["中风险"].append({"title": "存在未偿还负债", "detail": f"负债余额 {money(m['liabilities'], ccy)}。"})
    else:
        risks["低风险"].append({"title": "无未偿还负债", "detail": "净资产未被负债侵蚀。"})

    for cat, row in m["budget_usage"].items():
        if row["limit"] <= 0:
            continue
        if row["ratio"] > 1:
            risks["高风险"].append({"title": f"{cat}预算超支", "detail": f"使用率 {percent(row['ratio'])}。"})
        elif row["ratio"] > 0.85:
            risks["中风险"].append({"title": f"{cat}预算接近上限", "detail": f"使用率 {percent(row['ratio'])}。"})

    for goal in m["goals"]:
        if goal["days_left"] is not None and goal["days_left"] <= 45 and goal["progress"] < 0.60:
            risks["中风险"].append({"title": f"储蓄目标进度落后：{goal['name']}", "detail": f"剩余 {goal['days_left']} 天，完成度 {percent(goal['progress'])}。"})
        elif goal["progress"] >= 0.50:
            risks["低风险"].append({"title": f"储蓄目标进度正常：{goal['name']}", "detail": f"完成度 {percent(goal['progress'])}。"})

    if m["overdue_principal"] > 0:
        risks["高风险"].append({"title": "存在逾期贷款", "detail": f"逾期本金 {money(m['overdue_principal'], ccy)}，暂停新增放贷。"})

    return risks


def build_weekly_plan(data: Dict[str, Any]) -> List[str]:
    m = calc_financials(data)
    actions: List[str] = []

    overspent = [cat for cat, row in m["budget_usage"].items() if row["limit"] > 0 and row["ratio"] > 1]
    near = [cat for cat, row in m["budget_usage"].items() if row["limit"] > 0 and 0.85 < row["ratio"] <= 1]

    if overspent:
        actions.append(f"本周暂停 {', '.join(overspent)} 类非必要消费。")
    elif near:
        actions.append(f"本周 {', '.join(near)} 类消费必须先审批。")

    if m["loan_asset_ratio"] > 0.45:
        actions.append("优先收回应收贷款本金，目标至少收回 $10。")

    if m["cash"] < fnum(data["settings"].get("cash_floor"), 15):
        actions.append("下一笔收入先补现金缓冲，不立刻消费。")

    if m["savings_asset_ratio"] < 0.25:
        actions.append("下一笔收入的 30% 转入储蓄。")

    if m["liabilities"] > 0:
        actions.append("优先偿还负债，不新增借入资金。")

    if not actions:
        actions = [
            "维持当前消费节奏，但非必要消费继续走审批。",
            "下一笔收入至少 20% 转入储蓄。",
            "有人还款后，先补现金余额，再考虑消费。",
        ]

    return actions[:5]


# ============================================================
# 5. 商户权益风控
# ============================================================

def evaluate_merchant(data: Dict[str, Any], merchant: Dict[str, Any]) -> Dict[str, Any]:
    m = calc_financials(data)
    cat = merchant.get("category", "其他")
    usage = category_usage(m, cat)

    required_score = inum(merchant.get("required_score"), 700)
    category_cap = fnum(merchant.get("category_budget_cap"), 0.9)

    checks = {
        "信用分达标": m["credit_score"] >= required_score,
        "现金余额为正": m["cash"] > 0,
        "本月支出收入比不高于 90%": m["spend_income_ratio"] <= 0.90,
        "该品类预算未过高": usage["ratio"] <= category_cap if usage["limit"] > 0 else True,
    }

    open_status = all(checks.values())
    failed = [k for k, ok in checks.items() if not ok]

    if open_status:
        result = "开放"
        message = f"折扣开放：{fnum(merchant.get('discount')) * 100:.0f}%"
    elif m["credit_score"] >= required_score:
        result = "暂缓开放"
        message = "信用分达标，但预算或现金条件未通过。"
    else:
        result = "关闭"
        message = "信用分未达标。"

    return {
        "result": result,
        "message": message,
        "failed": failed,
        "checks": checks,
        "metrics": {
            "当前信用分": m["credit_score"],
            "要求信用分": required_score,
            "现金余额": m["cash"],
            "本月支出收入比": m["spend_income_ratio"],
            "该品类预算使用率": usage["ratio"],
            "品类开放上限": category_cap,
        },
    }


# ============================================================
# 6. DeepSeek 解释层
# ============================================================

def get_deepseek_key() -> str:
    key = os.getenv("DEEPSEEK_API_KEY", "")
    try:
        if "DEEPSEEK_API_KEY" in st.secrets:
            key = st.secrets["DEEPSEEK_API_KEY"]
    except Exception:
        pass
    return key or ""


def metric_lines(metrics: Dict[str, Any]) -> List[str]:
    lines = []
    for k, v in metrics.items():
        if isinstance(v, (int, float)):
            if "率" in k or "占比" in k or "使用率" in k or "上限" in k and fnum(v) <= 2:
                lines.append(f"- {k}：{percent(v)}")
            elif "信用分" in k:
                lines.append(f"- {k}：{v:.0f}")
            else:
                lines.append(f"- {k}：{money(v)}")
        else:
            lines.append(f"- {k}：{v}")
    return lines


def local_manager_report(decision: Dict[str, Any]) -> str:
    lines = [
        f"### {decision.get('kind', '风控报告')}",
        f"**审批结论：{decision.get('result')}**",
        "",
        "#### 关键硬指标",
        *metric_lines(decision.get("metrics", {})),
        "",
        "#### 风控原因",
    ]
    for r in decision.get("reasons", []):
        lines.append(f"- {r}")
    lines += [
        "",
        "#### 操作指令",
        decision.get("action", ""),
    ]
    return "\n".join(lines)


def deepseek_decision_report(decision: Dict[str, Any]) -> str:
    key = get_deepseek_key()
    if not key or requests is None:
        return local_manager_report(decision)

    payload = {
        "kind": decision.get("kind"),
        "result": decision.get("result"),
        "metrics": decision.get("metrics"),
        "reasons": decision.get("reasons"),
        "action": decision.get("action"),
    }

    system = """
你是“阿苏私人银行 2.0”的中文私人银行客户经理。
你只能解释 Python 已计算出的硬指标，不能编造任何数字，不能改变审批结论。
输出风格：克制、清楚、有银行风控感。
结构：
1. 审批结论
2. 关键数字
3. 风控原因
4. 操作指令
""".strip()

    try:
        resp = requests.post(
            "https://api.deepseek.com/chat/completions",
            headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"},
            json={
                "model": "deepseek-v4-pro",
                "messages": [
                    {"role": "system", "content": system},
                    {"role": "user", "content": json.dumps(payload, ensure_ascii=False, indent=2)},
                ],
                "temperature": 0.2,
                "max_tokens": 900,
            },
            timeout=20,
        )
        resp.raise_for_status()
        return resp.json()["choices"][0]["message"]["content"]
    except Exception as e:
        return local_manager_report(decision) + f"\n\n> DeepSeek 调用失败，已切换本地报告：{e}"


# ============================================================
# 7. 自然语言入口
# ============================================================

def extract_amounts(text: str) -> List[float]:
    if not text:
        return []
    patterns = [
        r"\$\s*([0-9]+(?:\.[0-9]+)?)",
        r"([0-9]+(?:\.[0-9]+)?)\s*(?:美元|刀|块|元)",
    ]
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
    for cat, keys in rules.items():
        if any(k in s for k in keys):
            return cat
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
        borrower = infer_borrower(text)
        return evaluate_loan_decision(
            data=data,
            principal=principal,
            borrower=borrower,
            expected_repayment=expected,
            due_date="",
            memo=text,
        )

    if is_purchase or amounts:
        amount = amounts[0] if amounts else 0.0
        category = infer_category(text)
        return evaluate_purchase_decision(
            data=data,
            amount=amount,
            category=category,
            description=text,
            merchant="自然语言输入",
        )

    return {
        "kind": "AI 决策中心",
        "result": "观察",
        "reasons": ["没有识别出明确金额或明确动作。"],
        "action": "请按格式输入：我想买 $18 的 Minecraft 道具；或：借给爸爸 $20，预计一周后还 $22。",
        "metrics": {},
        "pending_tx": None,
    }


# ============================================================
# 8. 页面样式
# ============================================================

def inject_css() -> None:
    st.markdown(
        """
        <style>
        .main .block-container {
            max-width: 1280px;
            padding-top: 1.2rem;
            padding-bottom: 2rem;
        }
        .hero {
            background: linear-gradient(135deg, #003C71 0%, #005EB8 48%, #0A74DA 100%);
            color: white;
            padding: 26px 30px;
            border-radius: 24px;
            margin-bottom: 18px;
            box-shadow: 0 16px 38px rgba(0, 62, 130, 0.25);
        }
        .hero h1 {
            margin: 0;
            font-size: 34px;
            letter-spacing: 0.3px;
        }
        .hero p {
            margin: 8px 0 0;
            opacity: 0.93;
            font-size: 16px;
        }
        .card {
            border: 1px solid #D6E4F5;
            border-radius: 18px;
            background: white;
            padding: 18px;
            box-shadow: 0 8px 22px rgba(15, 23, 42, 0.06);
            min-height: 112px;
        }
        .label {
            color: #6B7280;
            font-size: 13px;
            margin-bottom: 6px;
        }
        .value {
            font-size: 28px;
            font-weight: 780;
            color: #111827;
            line-height: 1.12;
        }
        .sub {
            color: #6B7280;
            font-size: 12px;
            margin-top: 6px;
        }
        .decision {
            border: 1px solid #D6E4F5;
            border-radius: 18px;
            background: white;
            padding: 18px;
            margin: 10px 0;
            box-shadow: 0 8px 22px rgba(15, 23, 42, 0.05);
        }
        .ok { border-left: 7px solid #166534; }
        .warn { border-left: 7px solid #B45309; }
        .bad { border-left: 7px solid #B91C1C; }
        .pill {
            display: inline-block;
            padding: 4px 10px;
            border-radius: 999px;
            background: #EAF3FF;
            color: #003C71;
            border: 1px solid #B9D7F6;
            font-size: 12px;
            font-weight: 700;
            margin-bottom: 8px;
        }
        .small {
            color: #6B7280;
            font-size: 13px;
        }
        div[data-testid="stMetric"] {
            border: 1px solid #D6E4F5;
            border-radius: 16px;
            padding: 12px 14px;
            background: white;
            box-shadow: 0 6px 18px rgba(15,23,42,0.05);
        }
        .stTabs [data-baseweb="tab-list"] { gap: 5px; }
        .stTabs [data-baseweb="tab"] {
            border-radius: 999px;
            background: #F2F6FB;
            padding: 8px 13px;
        }
        .stTabs [aria-selected="true"] {
            background: #005EB8 !important;
            color: white !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def hero(data: Dict[str, Any]) -> None:
    owner = data.get("settings", {}).get("owner", "阿苏")
    st.markdown(
        f"""
        <div class="hero">
            <h1>{APP_NAME}</h1>
            <p>{owner} 的家庭版私人银行：交易前审批 · 交易后模拟 · 信用分影响 · 风险雷达 · 本周行动指令</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def card(label: str, value: str, sub: str = "") -> None:
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


def render_decision(decision: Dict[str, Any], data: Dict[str, Any], key: str) -> None:
    cls = css_class_by_decision(decision.get("result", "观察"))
    st.markdown(
        f"""
        <div class="decision {cls}">
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
        rows = [{"指标": k, "值": pretty_value(k, v)} for k, v in decision["metrics"].items()]
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    with st.expander("AI 客户经理报告", expanded=True):
        st.markdown(deepseek_decision_report(decision))

    tx = decision.get("pending_tx")
    if tx:
        disabled = decision.get("result") == "拒绝"
        if st.button("确认入账" if not disabled else "拒绝结果不可入账", disabled=disabled, type="primary", key=key):
            data.setdefault("transactions", []).append(tx)
            commit(data)
            st.success("已入账。")
            st.rerun()


def pretty_value(k: str, v: Any) -> str:
    if isinstance(v, (int, float)):
        if "率" in k or "占比" in k or "使用率" in k or "上限" in k and fnum(v) <= 2:
            return percent(v)
        if "信用分" in k:
            return f"{v:.0f}"
        return money(v)
    return str(v)


# ============================================================
# 9. 页面
# ============================================================

def page_home(data: Dict[str, Any]) -> None:
    m = calc_financials(data)
    ccy = m["currency"]

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        card("总资产", money(m["total_assets"], ccy), "现金 + 储蓄 + 应收贷款本金")
    with c2:
        card("净资产", money(m["net_assets"], ccy), "总资产 - 负债")
    with c3:
        card("现金余额", money(m["cash"], ccy), "可立即使用资金")
    with c4:
        card("信用分", str(m["credit_score"]), "家庭内部风控评分，非 FICO")

    c5, c6, c7, c8 = st.columns(4)
    with c5:
        card("储蓄余额", money(m["savings"], ccy), f"储蓄占比 {percent(m['savings_asset_ratio'])}")
    with c6:
        card("应收贷款本金", money(m["receivables"], ccy), f"贷款资产占比 {percent(m['loan_asset_ratio'])}")
    with c7:
        card("负债余额", money(m["liabilities"], ccy), f"负债占比 {percent(m['debt_asset_ratio'])}")
    with c8:
        card("本月支出/收入", percent(m["spend_income_ratio"]), f"收入 {money(m['month_income'])} / 支出 {money(m['month_expense'])}")

    st.divider()

    left, right = st.columns([1.05, 1])
    with left:
        st.subheader("AI 本周行动计划")
        for i, action in enumerate(build_weekly_plan(data), 1):
            st.write(f"**{i}.** {action}")

        st.subheader("信用分因子")
        for x in m["score_factors"]:
            st.write(f"- {x}")

    with right:
        st.subheader("本月预算使用率")
        rows = []
        for cat, row in m["budget_usage"].items():
            rows.append({"分类": cat, "已花": row["spent"], "预算": row["limit"], "使用率": row["ratio"]})
        df = pd.DataFrame(rows)
        if not df.empty:
            st.bar_chart(df.set_index("分类")[["使用率"]])
            show = df.copy()
            show["已花"] = show["已花"].map(money)
            show["预算"] = show["预算"].map(money)
            show["使用率"] = show["使用率"].map(percent)
            st.dataframe(show, use_container_width=True, hide_index=True)


def page_add(data: Dict[str, Any]) -> None:
    st.subheader("新增交易")
    st.caption("正式入账区。购买和放贷建议先去审批页模拟。")

    tx_types = ["收入", "消费", "转入储蓄", "储蓄取出", "放贷", "还款", "借入", "偿还负债"]
    tx_type = st.selectbox("交易类型", tx_types, index=1)

    with st.form("add_tx"):
        d = st.date_input("日期", value=date.today())
        amount = st.number_input("金额", min_value=0.0, step=1.0, format="%.2f")
        cat_options = sorted(set(list(data.get("budgets", {}).keys()) + ["零花钱", "家庭贷款", "其他"]))
        category = st.selectbox("分类", cat_options)
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
        tx = new_tx(
            tx_type=tx_type,
            amount=amount,
            category=category,
            party=party,
            memo=memo,
            tx_date=d.isoformat(),
            expected_repayment=expected,
            principal_repaid=principal_repaid,
            interest_received=interest_received,
            due_date=due,
        )
        data.setdefault("transactions", []).append(tx)
        commit(data)
        st.success("交易已保存。")
        st.rerun()


def page_ai(data: Dict[str, Any]) -> None:
    st.subheader("AI 决策中心")
    st.caption("自然语言入口：先由 Python 算硬指标，再生成中文风控结论。")

    text = st.text_area(
        "输入请求",
        value="我想买 $18 的 Minecraft 道具",
        height=120,
        placeholder="例如：借给爸爸 $20，预计一周后还 $22",
    )

    if st.button("运行审批", type="primary"):
        st.session_state.latest_ai_decision = parse_natural_request(data, text)

    if "latest_ai_decision" in st.session_state:
        render_decision(st.session_state.latest_ai_decision, data, "ai_confirm")


def page_purchase(data: Dict[str, Any]) -> None:
    st.subheader("消费审批")
    st.caption("输出：批准 / 延迟 / 拒绝，并模拟信用分变化。")

    with st.form("purchase_form"):
        c1, c2, c3 = st.columns(3)
        with c1:
            amount = st.number_input("购买金额", value=18.0, min_value=0.0, step=1.0, format="%.2f")
        with c2:
            category = st.selectbox("消费分类", list(data.get("budgets", {}).keys()))
        with c3:
            merchant = st.text_input("商户", "Minecraft 商店")
        desc = st.text_area("购买说明", "Minecraft 道具")
        submitted = st.form_submit_button("模拟消费审批", type="primary")

    if submitted:
        st.session_state.purchase_decision = evaluate_purchase_decision(data, amount, category, desc, merchant)

    if "purchase_decision" in st.session_state:
        render_decision(st.session_state.purchase_decision, data, "purchase_confirm")


def page_loan(data: Dict[str, Any]) -> None:
    st.subheader("贷款审批")
    st.caption("放贷本金属于资产，但流动性弱，所以系统限制贷款资产占比。")

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
        st.session_state.loan_decision = evaluate_loan_decision(
            data=data,
            principal=principal,
            borrower=borrower,
            expected_repayment=expected,
            due_date=due.isoformat(),
            memo=memo,
        )

    if "loan_decision" in st.session_state:
        render_decision(st.session_state.loan_decision, data, "loan_confirm")


def page_radar(data: Dict[str, Any]) -> None:
    st.subheader("风险雷达")
    risks = build_risk_radar(data)

    cols = st.columns(3)
    for col, level in zip(cols, ["高风险", "中风险", "低风险"]):
        with col:
            st.markdown(f"### {level}")
            if not risks[level]:
                st.info("暂无")
            for r in risks[level]:
                css = "bad" if level == "高风险" else "warn" if level == "中风险" else "ok"
                st.markdown(
                    f"""
                    <div class="decision {css}">
                        <b>{r['title']}</b>
                        <p class="small">{r['detail']}</p>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

    st.divider()
    st.subheader("本周行动指令")
    for i, action in enumerate(build_weekly_plan(data), 1):
        st.write(f"**{i}.** {action}")


def page_statement(data: Dict[str, Any]) -> None:
    st.subheader("月度账单")

    df = tx_dataframe(data)
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
    c4.metric("信用分", m["credit_score"])

    month_df = df[df["日期"].astype(str).str[:7] == selected].copy()
    expense = month_df[month_df["类型"] == "消费"]
    if not expense.empty:
        by_cat = expense.groupby("分类")["金额"].sum().sort_values(ascending=False)
        st.subheader("消费分类")
        st.bar_chart(by_cat)

    st.subheader("明细")
    st.dataframe(month_df.drop(columns=["id"], errors="ignore"), use_container_width=True, hide_index=True)


def page_budget(data: Dict[str, Any]) -> None:
    st.subheader("预算管理")
    st.caption("预算是审批系统、风险雷达、商户折扣的基础输入。")

    with st.form("budget_form"):
        new_budgets = {}
        for cat, limit in data.get("budgets", {}).items():
            new_budgets[cat] = st.number_input(f"{cat} 月度预算", value=fnum(limit), min_value=0.0, step=1.0, format="%.2f", key=f"budget_{cat}")

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
        commit(data)
        st.success("预算已保存。")
        st.rerun()

    m = calc_financials(data)
    rows = []
    for cat, row in m["budget_usage"].items():
        rows.append({"分类": cat, "已花": money(row["spent"]), "预算": money(row["limit"]), "使用率": percent(row["ratio"])})
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


def page_goals(data: Dict[str, Any]) -> None:
    st.subheader("储蓄目标")

    for goal in calc_goals(data):
        st.markdown(f"#### {goal['name']}")
        st.progress(min(1.0, goal["progress"]))
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("目标", money(goal["target"]))
        c2.metric("当前", money(goal["current"]))
        c3.metric("完成度", percent(goal["progress"]))
        c4.metric("剩余", money(goal["remaining"]))
        if goal["days_left"] is not None:
            st.caption(f"截止日：{goal['deadline']}；剩余 {goal['days_left']} 天")
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
            category = st.selectbox("关联分类", list(data.get("budgets", {}).keys()))
        deadline = st.date_input("截止日", value=date.today() + timedelta(days=90))
        submitted = st.form_submit_button("新增目标", type="primary")

    if submitted and name.strip():
        data.setdefault("goals", []).append({
            "id": uid(),
            "name": name.strip(),
            "target": target,
            "current": current,
            "deadline": deadline.isoformat(),
            "category": category,
        })
        commit(data)
        st.success("目标已新增。")
        st.rerun()


def page_score(data: Dict[str, Any]) -> None:
    st.subheader("信用分")
    m = calc_financials(data)
    st.metric("当前信用分", m["credit_score"])
    st.caption("家庭内部风控评分，不是 FICO，不是银行真实征信。")

    st.markdown("#### 当前因子")
    for x in m["score_factors"]:
        st.write(f"- {x}")

    st.divider()
    st.markdown("#### 交易前模拟")
    with st.form("score_form"):
        typ = st.selectbox("交易类型", ["消费", "收入", "放贷", "转入储蓄", "借入", "偿还负债"])
        amount = st.number_input("金额", value=10.0, min_value=0.0, step=1.0, format="%.2f")
        category = st.selectbox("分类", sorted(set(list(data.get("budgets", {}).keys()) + ["零花钱", "家庭贷款", "其他"])))
        submitted = st.form_submit_button("模拟", type="primary")

    if submitted:
        score = estimate_score_after_transaction(data, new_tx(typ, amount, category, memo="信用分模拟"))
        c1, c2, c3 = st.columns(3)
        c1.metric("当前信用分", score["current_score"])
        c2.metric("交易后信用分", score["after_score"])
        c3.metric("变化", score["change"])


def page_merchants(data: Dict[str, Any]) -> None:
    st.subheader("商户权益")
    st.caption("不是信用分高就给折扣；必须同时满足信用分、现金、总预算纪律、品类预算。")

    for merchant in data.get("merchants", []):
        r = evaluate_merchant(data, merchant)
        css = css_class_by_decision(r["result"])
        failed = "无" if not r["failed"] else "；".join(r["failed"])
        st.markdown(
            f"""
            <div class="decision {css}">
                <div class="pill">{merchant.get('category')}</div>
                <h3>{merchant.get('name')}</h3>
                <p><b>{r['message']}</b></p>
                <p class="small">暂缓/失败原因：{failed}</p>
                <p class="small">{merchant.get('note', '')}</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        with st.expander(f"查看 {merchant.get('name')} 风控条件"):
            st.dataframe(pd.DataFrame([{"条件": k, "是否通过": "通过" if v else "未通过"} for k, v in r["checks"].items()]), use_container_width=True, hide_index=True)
            st.dataframe(pd.DataFrame([{"指标": k, "值": pretty_value(k, v)} for k, v in r["metrics"].items()]), use_container_width=True, hide_index=True)

    st.divider()
    with st.form("merchant_form"):
        st.markdown("#### 新增商户")
        name = st.text_input("商户名称", "")
        category = st.selectbox("商户分类", list(data.get("budgets", {}).keys()))
        c1, c2, c3 = st.columns(3)
        with c1:
            required = st.number_input("最低信用分", min_value=300, max_value=850, value=720)
        with c2:
            discount = st.number_input("折扣比例", min_value=0.0, max_value=0.9, value=0.08, step=0.01)
        with c3:
            cap = st.number_input("品类预算开放上限", min_value=0.0, max_value=2.0, value=0.90, step=0.05)
        note = st.text_input("说明", "")
        submitted = st.form_submit_button("新增商户", type="primary")

    if submitted and name.strip():
        data.setdefault("merchants", []).append({
            "id": uid(),
            "name": name.strip(),
            "category": category,
            "discount": discount,
            "required_score": int(required),
            "category_budget_cap": cap,
            "note": note,
        })
        commit(data)
        st.success("商户已新增。")
        st.rerun()


def tx_dataframe(data: Dict[str, Any]) -> pd.DataFrame:
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
    return pd.DataFrame(rows).sort_values(["日期"], ascending=False)


def page_transactions(data: Dict[str, Any]) -> None:
    st.subheader("交易流水")

    df = tx_dataframe(data)
    if df.empty:
        st.info("暂无交易。")
        return

    st.dataframe(df.drop(columns=["id"], errors="ignore"), use_container_width=True, hide_index=True)

    csv_text = df.drop(columns=["id"], errors="ignore").to_csv(index=False, encoding="utf-8-sig")
    st.download_button("下载交易 CSV", csv_text.encode("utf-8-sig"), "asu_money2_transactions.csv", "text/csv")

    json_text = json.dumps(data, ensure_ascii=False, indent=2)
    st.download_button("下载完整 JSON 数据", json_text.encode("utf-8"), "asu_money2_data.json", "application/json")

    st.divider()
    c1, c2 = st.columns(2)
    with c1:
        if st.button("删除最后一笔交易"):
            if data.get("transactions"):
                data["transactions"].pop()
                commit(data)
                st.success("已删除。")
                st.rerun()
    with c2:
        confirm = st.checkbox("确认重置为演示数据")
        if st.button("重置演示数据", disabled=not confirm):
            commit(copy.deepcopy(DEFAULT_DATA))
            st.success("已重置。")
            st.rerun()


# ============================================================
# 10. 侧边栏
# ============================================================

def sidebar(data: Dict[str, Any]) -> None:
    st.sidebar.title("系统设置")
    settings = data.setdefault("settings", {})

    with st.sidebar.form("settings"):
        owner = st.text_input("账户名称", settings.get("owner", "阿苏"))
        start_cash = st.number_input("初始现金", value=fnum(settings.get("start_cash")), step=1.0, format="%.2f")
        start_savings = st.number_input("初始储蓄", value=fnum(settings.get("start_savings")), step=1.0, format="%.2f")
        base_score = st.number_input("基础信用分", min_value=300, max_value=850, value=inum(settings.get("base_score"), 742))
        cash_floor = st.number_input("最低现金安全线", value=fnum(settings.get("cash_floor"), 15), step=1.0, format="%.2f")
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
        st.sidebar.success("设置已保存。")
        st.rerun()

    st.sidebar.divider()
    st.sidebar.caption("DeepSeek 状态")
    if get_deepseek_key():
        st.sidebar.success("已检测到 DEEPSEEK_API_KEY")
    else:
        st.sidebar.warning("未配置：使用本地中文风控报告")

    st.sidebar.caption(f"数据文件：{DATA_PATH.resolve()}")


# ============================================================
# 11. 主程序
# ============================================================

def main() -> None:
    st.set_page_config(
        page_title=APP_NAME,
        page_icon="🏦",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    inject_css()

    data = get_data()
    sidebar(data)
    hero(data)

    tabs = st.tabs([
        "1 首页总览",
        "2 新增交易",
        "3 AI 决策中心",
        "4 消费审批",
        "5 贷款审批",
        "6 风险雷达",
        "7 月度账单",
        "8 预算管理",
        "9 储蓄目标",
        "10 信用分",
        "11 商户权益",
        "12 交易流水",
    ])

    with tabs[0]:
        page_home(data)
    with tabs[1]:
        page_add(data)
    with tabs[2]:
        page_ai(data)
    with tabs[3]:
        page_purchase(data)
    with tabs[4]:
        page_loan(data)
    with tabs[5]:
        page_radar(data)
    with tabs[6]:
        page_statement(data)
    with tabs[7]:
        page_budget(data)
    with tabs[8]:
        page_goals(data)
    with tabs[9]:
        page_score(data)
    with tabs[10]:
        page_merchants(data)
    with tabs[11]:
        page_transactions(data)


if __name__ == "__main__":
    main()
