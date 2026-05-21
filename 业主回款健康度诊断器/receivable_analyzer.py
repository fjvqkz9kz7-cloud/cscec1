#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
receivable_analyzer.py
SAP FBL5N 客户行项目数据分析模块

功能：
    1. 读取并标准化 SAP FBL5N（客户行项目）导出的 Excel/CSV 数据
    2. 按客户编码 + 利润中心 + 合同维度进行应收账款分析
    3. 计算回款率、风险等级、催收优先级等核心指标
    4. 生成 plotly 可视化图表（仪表盘、对比图、瀑布图、饼图）
    5. 生成业主诊断报告文字

作者  ：AI Assistant
版本  ：2.0
日期  ：2025-06-12
"""

import warnings
import logging
from datetime import datetime, timedelta
from typing import Any

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ---------------------------------------------------------------------------
# 日志配置
# ---------------------------------------------------------------------------
logger = logging.getLogger("receivable_analyzer")
if not logger.handlers:
    _handler = logging.StreamHandler()
    _handler.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))
    logger.addHandler(_handler)
    logger.setLevel(logging.INFO)

# ---------------------------------------------------------------------------
# 列名映射（中文 ↔ 英文 SAP 列名）
# ---------------------------------------------------------------------------
COLUMN_MAP = {
    # 核心字段
    "已清项目/未清项目符号": "status_symbol",
    "Status symbol": "status_symbol",
    "凭证编号": "doc_no",
    "Document Number": "doc_no",
    "凭证类型": "doc_type",
    "Document Type": "doc_type",
    "本币金额": "amount_lc",
    "Amount in LC": "amount_lc",
    "Amount in local currency": "amount_lc",
    "清帐凭证": "clearing_doc",
    "Clearing Document": "clearing_doc",
    "清帐日期": "clearing_date",
    "Clearing Date": "clearing_date",
    "科目": "account",
    "Account": "account",
    "利润中心": "profit_center",
    "Profit Center": "profit_center",
    "总账科目": "gl_account",
    "G/L Account": "gl_account",
    "合同": "contract",
    "Contract": "contract",
    "WBS要素": "wbs_element",
    "WBS Element": "wbs_element",
    "客户": "customer_name",
    "Customer": "customer_name",
    "公司名称": "company_name",
    "Company Name": "company_name",
    "过账日期": "posting_date",
    "Posting Date": "posting_date",
    "文本": "text",
    "Text": "text",
    "分配": "assignment",
    "Assignment": "assignment",
    "到期日": "due_date",
    "Net Due Date": "due_date",
    "参照": "reference",
    "Reference": "reference",
}

# 正反向映射，方便复原展示
COLUMN_REVERSE_MAP = {v: k for k, v in COLUMN_MAP.items()}

# 金额列候选名（用于自动识别）
_AMOUNT_COL_CANDIDATES = [
    "amount_lc",
    "Amount in LC",
    "本币金额",
    "Amount in local currency",
]

# 凭证类型业务含义（可选扩展）
_DOC_TYPE_MEANING = {
    "YA": "预收",
    "Z2": "发票",
    "Z3": "发票冲销",
    "Z4": "贷项通知单",
    "Z5": "收款",
    "Z6": "付款",
    "DR": "借方凭证",
    "DZ": "客户付款",
    "RV": "发票凭证",
    "AB": "清账凭证",
}

# ---------------------------------------------------------------------------
# 辅助函数：列名标准化
# ---------------------------------------------------------------------------
def _detect_columns(df: pd.DataFrame) -> dict:
    """
    检测 DataFrame 的实际列名，返回 {标准英文名: 实际列名} 映射。

    兼容中文/英文/带空格等 SAP 导出格式。
    """
    actual_cols = list(df.columns)
    lower_actual = {c.strip().lower(): c for c in actual_cols}
    detected: dict[str, str] = {}

    for std_key, std_name in COLUMN_MAP.items():
        # 直接匹配
        if std_key in actual_cols:
            detected[std_name] = std_key
            continue
        # 大小写不敏感匹配
        lookup = std_key.strip().lower()
        if lookup in lower_actual:
            detected[std_name] = lower_actual[lookup]

    # 若 amount_lc 仍未检测到，尝试候选列名
    if "amount_lc" not in detected:
        for cand in _AMOUNT_COL_CANDIDATES:
            lookup = cand.strip().lower()
            if lookup in lower_actual:
                detected["amount_lc"] = lower_actual[lookup]
                break

    return detected


# ===========================================================================
# 1. 数据预处理
# ===========================================================================
def standardize_fbl5n(df: pd.DataFrame) -> pd.DataFrame:
    """
    标准化 SAP FBL5N 导出的客户行项目数据。

    处理步骤：
        1. 列名标准化为内部统一英文列名
        2. 过滤汇总行（status_symbol 含 '@' 的行）
        3. 过滤凭证编号为空的行
        4. 确保金额列为数值型
        5. 清洗日期列

    Parameters
    ----------
    df : pd.DataFrame
        原始 SAP FBL5N 导出数据（中文或英文列名均可）

    Returns
    -------
    pd.DataFrame
        清洗后的标准化 DataFrame，列名为内部英文名

    Examples
    --------
    >>> raw_df = pd.read_excel("FBL5N.xlsx")
    >>> clean_df = standardize_fbl5n(raw_df)
    """
    if df is None or df.empty:
        logger.warning("输入数据为空，返回空 DataFrame")
        return pd.DataFrame()

    df = df.copy()

    # 1) 检测列名映射
    col_map = _detect_columns(df)
    logger.info("检测到 %d 个可用列: %s", len(col_map), list(col_map.keys()))

    if "amount_lc" not in col_map:
        # 最后的兜底：尝试第一列包含 "金额" 或 "Amount" 的列
        for c in df.columns:
            if "金额" in str(c) or "amount" in str(c).lower():
                col_map["amount_lc"] = c
                logger.info("兜底识别金额列: %s", c)
                break

    # 2) 重命名可用列
    available_rename = {
        actual: std for std, actual in col_map.items() if actual in df.columns
    }
    df = df.rename(columns=available_rename)

    # 3) 过滤汇总行：status_symbol 含 '@' 符号的行
    if "status_symbol" in df.columns:
        before = len(df)
        mask = df["status_symbol"].astype(str).str.contains("@", na=False)
        df = df[~mask].copy()
        logger.info("过滤汇总行: %d -> %d 行", before, len(df))
    else:
        logger.warning("未检测到 status_symbol 列，跳过汇总行过滤")

    # 4) 过滤凭证编号为空的行
    if "doc_no" in df.columns:
        before = len(df)
        df = df[df["doc_no"].notna() & (df["doc_no"].astype(str).str.strip() != "")]
        logger.info("过滤空凭证编号: %d -> %d 行", before, len(df))

    # 5) 确保金额列为数值型
    if "amount_lc" in df.columns:
        df["amount_lc"] = pd.to_numeric(df["amount_lc"], errors="coerce")
        # 删除金额为空的关键行
        before = len(df)
        df = df[df["amount_lc"].notna()]
        logger.info("过滤空金额行: %d -> %d 行", before, len(df))
    else:
        logger.error("未检测到金额列，无法继续处理")
        raise ValueError(
            "未能识别金额列。支持的列名包括: "
            + ", ".join(_AMOUNT_COL_CANDIDATES)
        )

    # 6) 清洗日期列
    date_cols = ["clearing_date", "posting_date", "due_date"]
    for dcol in date_cols:
        if dcol in df.columns:
            df[dcol] = pd.to_datetime(df[dcol], errors="coerce")

    # 7) 确保关键维度列存在（缺失则填充空字符串）
    for dim_col in ["account", "profit_center", "contract"]:
        if dim_col not in df.columns:
            df[dim_col] = ""
        else:
            df[dim_col] = df[dim_col].fillna("").astype(str).str.strip()

    return df.reset_index(drop=True)


# ===========================================================================
# 2. 核心分析
# ===========================================================================
def analyze_fbl5n(
    df: pd.DataFrame,
    reference_date: datetime | None = None,
    default_payment_terms: int = 30
) -> dict:
    """
    对标准化的 FBL5N 数据进行应收账款核心分析。

    按客户编码 + 利润中心 + 合同维度分组，计算每个业主的回款率、
    风险等级、催收优先级等指标。

    Parameters
    ----------
    df : pd.DataFrame
        经过 standardize_fbl5n 处理后的标准化数据
    reference_date : datetime, optional
        分析基准日期，默认为当天

    Returns
    -------
    dict
        包含以下键值：
        - owner_summary: DataFrame —— 业主汇总（含催收优先级）
        - subject_breakdown: DataFrame —— 按总账科目汇总
        - detail: DataFrame —— 逐笔明细
        - stats: dict —— 整体统计指标

    业主汇总 DataFrame 列说明：
        - owner_key: 业主唯一标识（科目+利润中心+合同）
        - account: 客户编码
        - profit_center: 利润中心
        - contract: 合同号
        - ar_total: 应收总额（正金额之和，单位：元）
        - collected_total: 收回总额（|负金额|之和，单位：元）
        - net_balance: 净余额（元）
        - uncleared_balance: 未清余额 = max(净余额, 0)
        - cleared_count: 已清笔数
        - uncleared_count: 未清笔数
        - total_count: 总笔数
        - collection_rate: 回款率（0~100）
        - earliest_clearing: 最早清账日期
        - latest_clearing: 最晚清账日期
        - clearing_span_days: 清账跨度（天）
        - risk_level: 风险等级（高风险/预警/关注/正常）
        - priority_score: 催收优先级分数（越高越优先）
    """
    if df is None or df.empty:
        logger.warning("输入数据为空，返回空结果")
        return _empty_result(reference_date)

    ref_date = reference_date or datetime.now()
    detail = df.copy()

    # ---- 2a) 构建业主唯一键 ----
    detail["owner_key"] = (
        detail["account"].astype(str) + "|"
        + detail["profit_center"].astype(str) + "|"
        + detail["contract"].astype(str)
    )

    # ---- 2b) 识别正负金额 ----
    detail["amount_positive"] = detail["amount_lc"].clip(lower=0)   # 借方（应收增加）
    detail["amount_negative"] = detail["amount_lc"].clip(upper=0)  # 贷方（应收减少）

    # 清账状态
    if "clearing_date" in detail.columns:
        detail["is_cleared"] = detail["clearing_date"].notna()
    elif "clearing_doc" in detail.columns:
        detail["is_cleared"] = detail["clearing_doc"].notna() & (
            detail["clearing_doc"].astype(str).str.strip() != ""
        )
    else:
        # 无清账信息时，根据金额推断：负金额视为已清
        detail["is_cleared"] = detail["amount_lc"] < 0

    # ---- 2b+) 到期日与逾期天数计算 ----
    ref_date = pd.Timestamp(ref_date)
    # 如果数据中没有到期日列，用过账日期 + 默认付款条款推算
    if "due_date" not in detail.columns and "posting_date" in detail.columns:
        detail["due_date"] = detail["posting_date"] + pd.Timedelta(days=default_payment_terms)
    # 确保到期日是 datetime 类型
    if "due_date" in detail.columns:
        detail["due_date"] = pd.to_datetime(detail["due_date"], errors="coerce")
    else:
        # 没有到期日也没有过账日期，无法计算逾期
        detail["due_date"] = pd.NaT

    # 逾期天数 = 基准日期 - 到期日（仅在到期日有效时计算）
    detail["overdue_days"] = np.where(
        detail["due_date"].notna(),
        (ref_date - detail["due_date"]).dt.days,
        np.nan
    )
    # 未到期视为0（不逾期），NaT 保持 NaN
    detail["overdue_days"] = detail["overdue_days"].clip(lower=0).round(0)

    # 逾期状态：未清 + 已逾期（逾期天数 > 0）
    detail["is_overdue"] = (
        (~detail["is_cleared"])  # 未清
        & detail["overdue_days"].notna()
        & (detail["overdue_days"] > 0)  # 已逾期
    )

    # ---- 按业主分组聚合 ----
    grouped = detail.groupby("owner_key")

    agg = pd.DataFrame({
        "account": grouped["account"].first(),
        "profit_center": grouped["profit_center"].first(),
        "contract": grouped["contract"].first(),
        "ar_total": grouped["amount_positive"].sum().round(2),          # 应收总额
        "collected_total": grouped["amount_negative"].sum().abs().round(2),  # 收回总额
        "net_balance": grouped["amount_lc"].sum().round(2),              # 净余额
        "cleared_count": grouped["is_cleared"].sum().astype(int),
        "total_count": grouped.size(),
    })
    agg["uncleared_count"] = agg["total_count"] - agg["cleared_count"]

    # 最大逾期天数（仅统计未清且已逾期的交易）
    overdue_max = (
        detail[detail["is_overdue"]]
        .groupby("owner_key")["overdue_days"]
        .max()
        .reindex(agg.index)
        .fillna(0)
        .round(0)
        .astype(int)
    )
    agg["max_overdue_days"] = overdue_max

    # 预收金额（净余额为负表示多收了钱）
    agg["pre_receive_amount"] = (-agg["net_balance"]).clip(lower=0).round(2)
    # 未清余额（净余额为正表示还有钱没收回来）
    agg["uncleared_balance"] = agg["net_balance"].clip(lower=0).round(2)

    # ---- 回款率（不超过100%）----
    agg["collection_rate"] = np.where(
        agg["ar_total"] > 0,
        np.minimum((agg["collected_total"] / agg["ar_total"] * 100).round(2), 100),
        0.0
    )

    # ---- 2c) 清账日期统计 ----
    if "clearing_date" in detail.columns:
        clear_stats = grouped["clearing_date"].agg(
            earliest_clearing="min",
            latest_clearing="max"
        )
        clear_stats["clearing_span_days"] = (
            clear_stats["latest_clearing"] - clear_stats["earliest_clearing"]
        ).dt.days.fillna(0).astype(int)
        agg = agg.join(clear_stats, how="left")
    else:
        agg["earliest_clearing"] = pd.NaT
        agg["latest_clearing"] = pd.NaT
        agg["clearing_span_days"] = 0

    # ---- 2d) 按总账科目分组统计 ----
    if "gl_account" in detail.columns and detail["gl_account"].notna().any():
        subject_agg = detail.groupby("gl_account").agg(
            ar_total=("amount_positive", "sum"),
            collected_total=("amount_negative", "sum"),  # 负金额之和（负值）
            net_balance=("amount_lc", "sum"),
            transaction_count=("amount_lc", "size"),
        ).reset_index()
        # collected_total 取绝对值，uncleared_balance = max(net_balance, 0)
        subject_agg["collected_total"] = subject_agg["collected_total"].abs()
        subject_agg["uncleared_balance"] = subject_agg["net_balance"].clip(lower=0)
        subject_agg["collection_rate"] = np.where(
            subject_agg["ar_total"] > 0,
            np.minimum(
                subject_agg["collected_total"] / subject_agg["ar_total"] * 100, 100
            ),
            0.0
        )
    else:
        subject_agg = pd.DataFrame(columns=[
            "gl_account", "ar_total", "collected_total", "net_balance",
            "uncleared_balance", "transaction_count", "collection_rate"
        ])

    # ---- 2e) 风险等级评定 ----
    def _risk_level(row):
        rate = row["collection_rate"]
        bal = row["uncleared_balance"]
        pre = row["pre_receive_amount"]
        ar = row["ar_total"]
        # 预收预警：预收金额占应收总额 > 10%
        if ar > 0 and pre / ar > 0.10:
            return "预收预警"
        if rate < 50 or bal > 10_000_000:
            return "高风险"
        if rate < 80 or bal > 5_000_000:
            return "预警"
        if rate < 95:
            return "关注"
        return "正常"

    agg["risk_level"] = agg.apply(_risk_level, axis=1)

    # ---- 2f) 催收优先级排序 ----
    # 优先级分数 = 未清余额标准化得分 + 风险等级权重得分
    max_balance = agg["uncleared_balance"].max()
    if max_balance > 0:
        balance_score = agg["uncleared_balance"] / max_balance * 50  # 余额占50分
    else:
        balance_score = 0

    risk_weights = {"高风险": 40, "预警": 25, "关注": 15, "正常": 0, "预收预警": 10}
    risk_score = agg["risk_level"].map(risk_weights)

    # ---- 逾期时间维度（0-60分，仅对未清且已逾期赋分） ----
    def _overdue_score(days):
        if days <= 0:
            return 0          # 未清但未逾期 或 已清 → 不计分
        elif days <= 10:
            return 5           # 1-10天
        elif days <= 30:
            return 15          # 11-30天
        elif days <= 60:
            return 30          # 31-60天
        elif days <= 90:
            return 45          # 61-90天
        else:
            return 60          # 90天以上

    overdue_score = agg["max_overdue_days"].apply(_overdue_score)

    agg["priority_score"] = (balance_score + risk_score + overdue_score).round(2)
    agg = agg.sort_values("priority_score", ascending=False)
    agg = agg.reset_index()

    # ---- 2g) 整体汇总统计 ----
    total_ar = agg["ar_total"].sum()
    total_collected = agg["collected_total"].sum()
    net_balance = agg["net_balance"].sum()
    uncleared_balance = agg["uncleared_balance"].sum()
    overall_rate = (
        min(total_collected / total_ar * 100, 100) if total_ar > 0 else 0.0
    )

    # 逾期统计
    overdue_stats = agg["max_overdue_days"].agg(["max", "mean"]).round(0).astype(int)
    overdue_owners = (agg["max_overdue_days"] > 0).sum()

    stats = {
        "total_ar": round(float(total_ar), 2),
        "total_collected": round(float(total_collected), 2),
        "net_balance": round(float(net_balance), 2),
        "uncleared_balance": round(float(uncleared_balance), 2),
        "collection_rate": round(float(overall_rate), 2),
        "owner_count": int(len(agg)),
        "high_risk_count": int((agg["risk_level"] == "高风险").sum()),
        "warning_count": int((agg["risk_level"] == "预警").sum()),
        "attention_count": int((agg["risk_level"] == "关注").sum()),
        "normal_count": int((agg["risk_level"] == "正常").sum()),
        "pre_receive_count": int((agg["risk_level"] == "预收预警").sum()),
        "overdue_owner_count": int(overdue_owners),
        "max_overdue_days": int(overdue_stats["max"]),
        "avg_overdue_days": int(overdue_stats["mean"]),
        "analysis_date": ref_date.strftime("%Y-%m-%d %H:%M"),
    }

    return {
        "owner_summary": agg,
        "subject_breakdown": subject_agg,
        "detail": detail,
        "stats": stats,
    }


def _empty_result(reference_date: datetime | None = None) -> dict:
    """返回空的分析结果结构。"""
    ref = (reference_date or datetime.now()).strftime("%Y-%m-%d %H:%M")
    empty_owner = pd.DataFrame(columns=[
        "owner_key", "account", "profit_center", "contract",
        "ar_total", "collected_total", "net_balance", "uncleared_balance",
        "pre_receive_amount", "max_overdue_days",
        "cleared_count", "uncleared_count", "total_count",
        "collection_rate", "earliest_clearing", "latest_clearing",
        "clearing_span_days", "risk_level", "priority_score"
    ])
    empty_subject = pd.DataFrame(columns=[
        "gl_account", "ar_total", "collected_total", "net_balance",
        "uncleared_balance", "transaction_count", "collection_rate"
    ])
    return {
        "owner_summary": empty_owner,
        "subject_breakdown": empty_subject,
        "detail": pd.DataFrame(),
        "stats": {
            "total_ar": 0.0, "total_collected": 0.0, "net_balance": 0.0,
            "uncleared_balance": 0.0, "collection_rate": 0.0,
            "owner_count": 0, "high_risk_count": 0, "warning_count": 0,
            "attention_count": 0, "normal_count": 0, "pre_receive_count": 0,
            "analysis_date": ref,
        },
    }


# ===========================================================================
# 3. 图表生成函数
# ===========================================================================
# 颜色主题
_RISK_COLORS = {
    "高风险": "#DC143C",  # 深红
    "预警": "#FF8C00",    # 深橙
    "关注": "#FFD700",    # 金黄
    "正常": "#32CD32",    # 绿色
}

_RISK_ORDER = ["高风险", "预警", "关注", "正常"]


def create_risk_gauge(score: float) -> go.Figure:
    """
    创建健康度仪表盘图。

    Parameters
    ----------
    score : float
        健康度评分（0~100），越高越健康

    Returns
    -------
    plotly.graph_objects.Figure
        仪表盘图表对象
    """
    score = float(np.clip(score, 0, 100))

    # 根据分数确定颜色
    if score >= 80:
        color = "#32CD32"
    elif score >= 60:
        color = "#FFD700"
    elif score >= 40:
        color = "#FF8C00"
    else:
        color = "#DC143C"

    fig = go.Figure(
        go.Indicator(
            mode="gauge+number+delta",
            value=score,
            number={"suffix": "分", "font": {"size": 36, "color": color}},
            title={
                "text": "应收健康度评分",
                "font": {"size": 20},
            },
            delta={"reference": 60, "suffix": "分"},
            gauge={
                "axis": {
                    "range": [0, 100],
                    "tickwidth": 1,
                    "tickcolor": "#444",
                    "tickvals": [0, 20, 40, 60, 80, 100],
                },
                "bar": {"color": color, "thickness": 0.75},
                "bgcolor": "white",
                "borderwidth": 2,
                "bordercolor": "#ccc",
                "steps": [
                    {"range": [0, 40], "color": "#FFE4E1"},
                    {"range": [40, 60], "color": "#FFF8DC"},
                    {"range": [60, 80], "color": "#FFFACD"},
                    {"range": [80, 100], "color": "#E0FFE0"},
                ],
                "threshold": {
                    "line": {"color": "black", "width": 3},
                    "thickness": 0.8,
                    "value": 60,
                },
            },
        )
    )

    fig.update_layout(
        height=400,
        margin=dict(t=60, b=30, l=30, r=30),
        font=dict(family="Microsoft YaHei, SimHei, Arial, sans-serif"),
    )
    return fig


def create_owner_comparison(summary_df: pd.DataFrame, top_n: int = 20) -> go.Figure:
    """
    创建业主对比图：应收金额横向条形图 + 回款率对比。

    Parameters
    ----------
    summary_df : pd.DataFrame
        analyze_fbl5n 返回的 owner_summary DataFrame
    top_n : int, default 20
        显示前 N 个高优先级业主

    Returns
    -------
    plotly.graph_objects.Figure
        横向条形图，颜色按风险等级区分
    """
    if summary_df is None or summary_df.empty:
        fig = go.Figure()
        fig.add_annotation(text="暂无数据", showarrow=False, font_size=20)
        return fig

    df = summary_df.head(top_n).copy()
    df = df.iloc[::-1]  # 倒序，让最高分在最上面

    # 业主显示名称
    df["owner_display"] = (
        df["account"].astype(str).str[:8] + " | "
        + df["profit_center"].astype(str).str[-6:] + " | "
        + df["contract"].astype(str).str[:12]
    )

    # 转换为万元
    df["ar_total_wan"] = df["ar_total"] / 10_000
    df["uncleared_wan"] = df["uncleared_balance"] / 10_000

    # 悬停文本
    df["hover_text"] = (
        "客户: " + df["account"].astype(str) + "<br>"
        + "利润中心: " + df["profit_center"].astype(str) + "<br>"
        + "合同: " + df["contract"].astype(str) + "<br>"
        + "应收总额: " + df["ar_total_wan"].round(2).astype(str) + " 万元<br>"
        + "未清余额: " + df["uncleared_wan"].round(2).astype(str) + " 万元<br>"
        + "回款率: " + df["collection_rate"].round(2).astype(str) + "%<br>"
        + "风险等级: " + df["risk_level"]
    )

    colors = df["risk_level"].map(_RISK_COLORS).fillna("#999")

    fig = go.Figure()

    # 应收总额条形
    fig.add_trace(go.Bar(
        y=df["owner_display"],
        x=df["ar_total_wan"],
        orientation="h",
        name="应收总额(万元)",
        marker_color=colors,
        opacity=0.85,
        text=df["ar_total_wan"].round(1).astype(str),
        textposition="outside",
        hovertemplate="%{customdata}<extra></extra>",
        customdata=df["hover_text"],
    ))

    fig.update_layout(
        title=f"Top {min(top_n, len(summary_df))} 业主应收对比（按催收优先级排序）",
        xaxis_title="金额（万元）",
        yaxis_title="",
        height=max(400, len(df) * 32),
        margin=dict(t=50, b=40, l=200, r=40),
        font=dict(family="Microsoft YaHei, SimHei, Arial, sans-serif"),
        plot_bgcolor="#FAFAFA",
        showlegend=False,
    )

    # 添加风险等级图例（使用假数据点）
    for level in _RISK_ORDER:
        fig.add_trace(go.Scatter(
            x=[None], y=[None],
            mode="markers",
            marker=dict(size=12, color=_RISK_COLORS[level]),
            legendgroup=level,
            showlegend=True,
            name=level,
        ))

    fig.update_layout(legend=dict(
        orientation="h", yanchor="bottom", y=-0.15, xanchor="center", x=0.5,
        title="风险等级:"
    ))

    return fig


def create_ar_waterfall(summary_df: pd.DataFrame) -> go.Figure:
    """
    创建应收构成瀑布图：应收总额 -> 已收回 -> 未清余额。

    Parameters
    ----------
    summary_df : pd.DataFrame
        analyze_fbl5n 返回的 owner_summary DataFrame

    Returns
    -------
    plotly.graph_objects.Figure
        瀑布图展示应收构成
    """
    if summary_df is None or summary_df.empty:
        fig = go.Figure()
        fig.add_annotation(text="暂无数据", showarrow=False, font_size=20)
        return fig

    total_ar = summary_df["ar_total"].sum()
    total_collected = summary_df["collected_total"].sum()
    uncleared = summary_df["uncleared_balance"].sum()

    # 转换为万元
    ar_wan = total_ar / 10_000
    collected_wan = total_collected / 10_000
    uncleared_wan = uncleared / 10_000

    # 瀑布图数据
    categories = ["应收总额", "已收回金额", "未清余额"]
    values = [ar_wan, -collected_wan, uncleared_wan]
    measures = ["absolute", "relative", "total"]

    fig = go.Figure(go.Waterfall(
        orientation="v",
        measure=measures,
        x=categories,
        y=values,
        textposition="outside",
        text=[
            f"{ar_wan:,.2f}",
            f"-{collected_wan:,.2f}",
            f"{uncleared_wan:,.2f}",
        ],
        connector={"line": {"color": "rgb(150,150,150)", "dash": "dot"}},
        increasing={"marker": {"color": "#32CD32"}},
        decreasing={"marker": {"color": "#DC143C"}},
        totals={"marker": {"color": "#FF8C00"}},
    ))

    fig.update_layout(
        title="应收构成瀑布图（单位：万元）",
        yaxis_title="金额（万元）",
        xaxis_title="",
        height=450,
        margin=dict(t=60, b=40, l=60, r=40),
        font=dict(family="Microsoft YaHei, SimHei, Arial, sans-serif"),
        plot_bgcolor="#FAFAFA",
        showlegend=False,
    )

    # 添加回款率标注
    rate = min(total_collected / total_ar * 100, 100) if total_ar > 0 else 0
    fig.add_annotation(
        x="未清余额", y=uncleared_wan,
        text=f"综合回款率: {rate:.2f}%",
        showarrow=True,
        arrowhead=2,
        ax=60, ay=-40,
        font=dict(size=13, color="#333"),
        bgcolor="rgba(255,255,255,0.8)",
        bordercolor="#ccc",
    )

    return fig


def create_uncleared_detail(
    df: pd.DataFrame,
    owner_name: str
) -> go.Figure | None:
    """
    创建业主未清余额时间轴图。

    按每笔交易的时间顺序，逐笔累加计算累计未清余额，
    可视化展示未清金额随时间的变化过程（含分批付款场景）。
    红色柱状=应收增加（正金额），绿色柱状=回款减少（负金额），
    橙色折线=累计未清余额。

    Parameters
    ----------
    df : pd.DataFrame
        逐笔明细数据（analyze_fbl5n 返回的 detail DataFrame）
    owner_name : str
        业主唯一键

    Returns
    -------
    plotly.graph_objects.Figure or None
        时间轴双Y轴图表，未找到数据时返回 None
    """
    if df.empty or "owner_key" not in df.columns:
        return None

    owner_df = df[df["owner_key"] == owner_name].copy()
    if owner_df.empty:
        return None

    # ---- 按时间排序，逐笔计算累计未清余额 ----
    date_col = None
    for col in ["posting_date", "doc_date", "clearing_date"]:
        if col in owner_df.columns and owner_df[col].notna().any():
            date_col = col
            break

    if date_col is None:
        # 没有时间列，fallback 到凭证类型饼图
        return _fallback_doc_type_pie(owner_df, owner_name)

    # 按日期排序
    timeline = owner_df.sort_values(date_col).copy()
    timeline["running_balance"] = timeline["amount_lc"].cumsum().round(2)
    # 累计未清余额 = running_balance 为正的部分
    timeline["uncleared"] = timeline["running_balance"].clip(lower=0).round(2)

    # 交易类型标签
    timeline["txn_label"] = timeline["amount_lc"].apply(
        lambda a: f"应收 +{a:,.2f}" if a > 0 else f"回款 {a:,.2f}"
    )

    # ---- 双Y轴图表 ----
    fig = make_subplots(specs=[[{"secondary_y": True}]])

    # 柱状图：每笔交易金额
    colors = ["#E74C3C" if v > 0 else "#27AE60" for v in timeline["amount_lc"]]
    fig.add_trace(
        go.Bar(
            x=timeline[date_col],
            y=timeline["amount_lc"],
            name="单笔交易",
            marker_color=colors,
            opacity=0.6,
            hovertemplate=(
                "<b>%{customdata}</b><br>"
                "日期: %{x}<br>"
                "金额: ¥%{y:,.2f}<br>"
                "凭证: %{meta}<br>"
                "<extra></extra>"
            ),
            customdata=timeline["txn_label"],
            meta=timeline["doc_no"].astype(str),
        ),
        secondary_y=False,
    )

    # 折线图：累计未清余额
    fig.add_trace(
        go.Scatter(
            x=timeline[date_col],
            y=timeline["uncleared"],
            name="累计未清余额",
            mode="lines+markers",
            line=dict(color="#E67E22", width=2),
            marker=dict(size=6),
            hovertemplate=(
                "<b>累计未清余额</b><br>"
                "日期: %{x}<br>"
                "未清: ¥%{y:,.2f}<br>"
                "<extra></extra>"
            ),
        ),
        secondary_y=True,
    )

    # 零线
    fig.add_hline(y=0, line_dash="dash", line_color="gray", line_width=1)

    fig.update_layout(
        title=f"{owner_name} — 未清余额时间轴（分批付款追踪）",
        xaxis_title="日期",
        height=500,
        margin=dict(l=50, r=50, t=60, b=50),
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02,
                     xanchor="center", x=0.5),
    )
    fig.update_yaxes(title_text="单笔交易金额（元）", secondary_y=False)
    fig.update_yaxes(title_text="累计未清余额（元）", secondary_y=True)

    return fig


def _fallback_doc_type_pie(
    owner_df: pd.DataFrame,
    owner_name: str
) -> go.Figure:
    """
    无时间列时的 fallback：按凭证类型饼图。
    """
    doc_type_groups = owner_df.groupby("doc_type")["amount_lc"].sum().reset_index()
    doc_type_groups.columns = ["凭证类型", "金额"]

    if doc_type_groups.empty:
        fig = go.Figure()
        fig.update_layout(
            title=f"{owner_name} — 暂无数据",
            height=400,
        )
        return fig

    fig = go.Figure(data=[go.Pie(
        labels=doc_type_groups["凭证类型"],
        values=doc_type_groups["金额"],
        hole=0.4,
        textinfo="label+percent",
        textfont_size=12,
        marker=dict(colors=px.colors.qualitative.Set2),
        hovertemplate=(
            "<b>%{label}</b><br>"
            "金额: ¥%{value:,.2f}<br>"
            "占比: %{percent}<br>"
            "<extra></extra>"
        ),
    )])

    fig.update_layout(
        title=f"{owner_name} — 按凭证类型分布",
        height=450,
        margin=dict(l=20, r=20, t=50, b=20),
    )
    return fig


# ===========================================================================
# 4. 账龄分析
# ===========================================================================

def analyze_aging(
    detail: pd.DataFrame,
    reference_date: datetime | None = None,
    aging_bins: list[int] | None = None,
    aging_labels: list[str] | None = None
) -> dict:
    """
    应收账款账龄分析。

    基于 FBL5N 逐笔明细数据，以过账日期为起点计算每笔应收款的账龄天数，
    按账龄区间分类统计各业主的应收金额分布。

    只分析正金额（应收增加）的交易，已收回（负金额）和已清交易不计入账龄。

    Parameters
    ----------
    detail : pd.DataFrame
        analyze_fbl5n 返回的 detail DataFrame，需包含 posting_date、
        amount_lc、amount_positive、owner_key、account 等列
    reference_date : datetime, optional
        账龄计算基准日期，默认当天
    aging_bins : list[int], optional
        账龄区间边界（天数），默认 [0, 30, 60, 90, 180, 9999]
    aging_labels : list[str], optional
        账龄区间标签，默认 ["0-30天", "31-60天", "61-90天", "91-180天", "180天以上"]

    Returns
    -------
    dict
        {
            'owner_aging': DataFrame,      -- 按业主+账龄区间汇总
            'summary': DataFrame,          -- 全局账龄汇总
            'total_ar_by_owner': DataFrame, -- 各业主应收总额
            'reference_date': str,         -- 基准日期
            'aging_bins': list,            -- 区间边界
            'aging_labels': list,          -- 区间标签
        }

    owner_aging DataFrame 列：
        - owner_key: 业主标识
        - account: 客户编码
        - 账龄区间: 如 "0-30天"
        - amount: 该区间应收金额（元）
        - count: 该区间笔数
        - percent: 占该业主应收总额比例
    """
    if reference_date is None:
        reference_date = datetime.now()
    ref = pd.Timestamp(reference_date)

    # 默认账龄区间
    if aging_bins is None:
        aging_bins = [0, 30, 60, 90, 180, 9999]
    if aging_labels is None:
        aging_labels = ["0-30天", "31-60天", "61-90天", "91-180天", "180天以上"]

    # 确定日期列：优先用过账日期，其次用到期日，最后尝试清帐日期
    date_col = None
    for col in ["posting_date", "doc_date", "due_date"]:
        if col in detail.columns and detail[col].notna().any():
            date_col = col
            break

    if date_col is None:
        logger.warning("detail 数据中无有效日期列（posting_date/doc_date/due_date），无法计算账龄")
        empty_owner = pd.DataFrame(columns=[
            "owner_key", "account", "账龄区间", "amount", "count", "percent"
        ])
        empty_summary = pd.DataFrame({
            "账龄区间": aging_labels,
            "amount": [0.0] * len(aging_labels),
            "count": [0] * len(aging_labels),
            "percent": [0.0] * len(aging_labels),
        })
        return {
            "owner_aging": empty_owner,
            "summary": empty_summary,
            "total_ar_by_owner": pd.DataFrame(columns=["owner_key", "account", "total_ar"]),
            "reference_date": ref.strftime("%Y-%m-%d"),
            "aging_bins": aging_bins,
            "aging_labels": aging_labels,
        }

    # 过滤有效数据：有日期、且为正金额（应收增加）
    mask = (
        detail[date_col].notna()
        & (detail["amount_positive"] > 0)  # 只分析应收增加
    )
    ar_df = detail[mask].copy()

    if ar_df.empty:
        logger.warning("无有效过账日期的应收数据，返回空账龄分析")
        empty_owner = pd.DataFrame(columns=[
            "owner_key", "account", "账龄区间", "amount", "count", "percent"
        ])
        empty_summary = pd.DataFrame({
            "账龄区间": aging_labels,
            "amount": [0.0] * len(aging_labels),
            "count": [0] * len(aging_labels),
            "percent": [0.0] * len(aging_labels),
        })
        return {
            "owner_aging": empty_owner,
            "summary": empty_summary,
            "total_ar_by_owner": pd.DataFrame(columns=["owner_key", "account", "total_ar"]),
            "reference_date": ref.strftime("%Y-%m-%d"),
            "aging_bins": aging_bins,
            "aging_labels": aging_labels,
        }

    # 计算账龄天数 = 基准日期 - 日期列
    ar_df["aging_days"] = (ref - ar_df[date_col]).dt.days
    ar_df["aging_days"] = ar_df["aging_days"].clip(lower=0).round(0).astype(int)

    # 账龄区间分类
    ar_df["账龄区间"] = pd.cut(
        ar_df["aging_days"],
        bins=aging_bins,
        labels=aging_labels,
        right=True,       # 右闭区间
        include_lowest=True
    )

    # ---- 4a) 按业主+账龄区间汇总 ----
    owner_aging = (
        ar_df.groupby(["owner_key", "账龄区间"])
        .agg(
            amount=("amount_positive", "sum"),    # 金额
            count=("amount_positive", "size"),    # 笔数
        )
        .reset_index()
    )
    # 添加客户编码
    owner_acct_map = ar_df.groupby("owner_key")["account"].first().to_dict()
    owner_aging["account"] = owner_aging["owner_key"].map(owner_acct_map)

    # 计算各业主应收总额，用于计算占比
    total_ar_by_owner = (
        ar_df.groupby("owner_key")["amount_positive"]
        .sum()
        .reset_index()
        .rename(columns={"amount_positive": "total_ar"})
    )
    acct_map = ar_df.groupby("owner_key")["account"].first().to_dict()
    total_ar_by_owner["account"] = total_ar_by_owner["owner_key"].map(acct_map)

    # 计算占比
    owner_aging = owner_aging.merge(
        total_ar_by_owner[["owner_key", "total_ar"]],
        on="owner_key",
        how="left"
    )
    owner_aging["percent"] = np.where(
        owner_aging["total_ar"] > 0,
        (owner_aging["amount"] / owner_aging["total_ar"] * 100).round(2),
        0.0
    )
    owner_aging = owner_aging.reset_index(drop=True)
    owner_aging = owner_aging.drop(columns=["total_ar"])

    # ---- 4b) 全局账龄汇总 ----
    summary = (
        ar_df.groupby("账龄区间")
        .agg(
            amount=("amount_positive", "sum"),
            count=("amount_positive", "size"),
        )
        .reset_index()
    )
    total_ar = ar_df["amount_positive"].sum()
    summary["percent"] = np.where(
        total_ar > 0,
        (summary["amount"] / total_ar * 100).round(2),
        0.0
    )

    return {
        "owner_aging": owner_aging,
        "summary": summary,
        "total_ar_by_owner": total_ar_by_owner,
        "reference_date": ref.strftime("%Y-%m-%d"),
        "aging_bins": aging_bins,
        "aging_labels": aging_labels,
    }


# ===========================================================================
# 5. 账龄图表
# ===========================================================================

def create_aging_chart(
    owner_aging: pd.DataFrame,
    aging_labels: list[str],
    top_n: int = 15
) -> go.Figure:
    """
    创建账龄分析堆叠柱状图。

    按业主展示各账龄区间的应收金额分布，颜色区分账龄区间。

    Parameters
    ----------
    owner_aging : pd.DataFrame
        analyze_aging 返回的 owner_aging DataFrame
    aging_labels : list[str]
        账龄区间标签列表
    top_n : int
        展示前N个业主（按应收总额排序），默认15

    Returns
    -------
    plotly.graph_objects.Figure
        堆叠柱状图
    """
    if owner_aging.empty:
        fig = go.Figure()
        fig.update_layout(
            title="账龄分析 — 无数据",
            height=400,
        )
        return fig

    # 透视表：业主 x 账龄区间
    pivot = owner_aging.pivot_table(
        index="account",
        columns="账龄区间",
        values="amount",
        aggfunc="sum",
        fill_value=0,
        observed=False
    )

    # 确保所有账龄区间列都存在
    for label in aging_labels:
        if label not in pivot.columns:
            pivot[label] = 0
    pivot = pivot[aging_labels]  # 按顺序排列

    # 按总额排序，取前N
    pivot["_total"] = pivot.sum(axis=1)
    pivot = pivot.sort_values("_total", ascending=True).tail(top_n)
    pivot = pivot.drop(columns=["_total"])

    # 账龄区间颜色映射
    color_map = {
        "0-30天": "#27AE60",      # 绿
        "31-60天": "#F1C40F",     # 黄
        "61-90天": "#E67E22",     # 橙
        "91-180天": "#E74C3C",    # 红
        "180天以上": "#922B21",   # 深红
    }

    fig = go.Figure()
    for label in aging_labels:
        if label in pivot.columns:
            fig.add_trace(go.Bar(
                name=label,
                y=pivot.index.astype(str),
                x=pivot[label] / 10000,  # 转万元
                orientation="h",
                marker_color=color_map.get(label, "#95A5A6"),
                hovertemplate=(
                    f"<b>%{{y}}</b><br>"
                    f"{label}<br>"
                    "金额: ¥%{x:,.2f}万<br>"
                    "<extra></extra>"
                ),
            ))

    fig.update_layout(
        barmode="stack",
        title=f"业主账龄分析（前{min(top_n, len(pivot))}名，单位：万元）",
        xaxis_title="应收金额（万元）",
        yaxis_title="",
        height=max(400, len(pivot) * 35),
        margin=dict(l=120, r=30, t=60, b=50),
        legend=dict(
            orientation="h",
            yanchor="bottom", y=1.02,
            xanchor="center", x=0.5,
            title="账龄区间",
        ),
        hovermode="y unified",
    )
    return fig


def create_aging_summary_chart(summary: pd.DataFrame) -> go.Figure:
    """
    创建全局账龄汇总图（饼图+柱状图组合）。

    Parameters
    ----------
    summary : pd.DataFrame
        analyze_aging 返回的 summary DataFrame

    Returns
    -------
    plotly.graph_objects.Figure
        组合图
    """
    if summary.empty:
        fig = go.Figure()
        fig.update_layout(title="账龄汇总 — 无数据", height=400)
        return fig

    color_map = {
        "0-30天": "#27AE60",
        "31-60天": "#F1C40F",
        "61-90天": "#E67E22",
        "91-180天": "#E74C3C",
        "180天以上": "#922B21",
    }
    colors = [color_map.get(l, "#95A5A6") for l in summary["账龄区间"]]

    fig = make_subplots(
        rows=1, cols=2,
        subplot_titles=("金额分布", "占比分布"),
        specs=[[{"type": "bar"}, {"type": "pie"}]]
    )

    # 柱状图
    fig.add_trace(
        go.Bar(
            x=summary["账龄区间"],
            y=summary["amount"] / 10000,
            name="金额",
            marker_color=colors,
            text=[f"{v/10000:,.1f}万" for v in summary["amount"]],
            textposition="outside",
            hovertemplate="<b>%{x}</b><br>金额: ¥%{y:,.2f}万<br><extra></extra>",
        ),
        row=1, col=1
    )

    # 饼图
    fig.add_trace(
        go.Pie(
            labels=summary["账龄区间"],
            values=summary["amount"],
            hole=0.4,
            marker_colors=colors,
            textinfo="label+percent",
            textfont_size=11,
            hovertemplate=(
                "<b>%{label}</b><br>"
                "金额: ¥%{value:,.2f}<br>"
                "占比: %{percent}<br>"
                "<extra></extra>"
            ),
        ),
        row=1, col=2
    )

    fig.update_layout(
        title="应收账款账龄汇总",
        height=420,
        margin=dict(l=40, r=40, t=70, b=40),
        showlegend=False,
    )
    return fig


# ===========================================================================
# 6. 报告生成
# ===========================================================================
def generate_owner_report(row: pd.Series | dict) -> str:
    """
    生成单个业主的诊断报告文字。

    Parameters
    ----------
    row : pd.Series or dict
        业主汇总数据行，需包含以下字段：
        - account: 客户编码
        - profit_center: 利润中心
        - contract: 合同号
        - ar_total: 应收总额（元）
        - collected_total: 收回总额（元）
        - net_balance: 净余额（元）
        - uncleared_balance: 未清余额（元）
        - collection_rate: 回款率
        - cleared_count: 已清笔数
        - uncleared_count: 未清笔数
        - risk_level: 风险等级
        - priority_score: 优先级分数

    Returns
    -------
    str
        诊断报告文字（Markdown 格式）
    """
    if row is None:
        return "无数据，无法生成报告。"

    # 统一转为 dict
    if isinstance(row, pd.Series):
        row = row.to_dict()

    def _v(key, default=""):
        return row.get(key, default)

    # 提取数据
    account = str(_v("account", "未知"))
    profit_center = str(_v("profit_center", "未知"))
    contract = str(_v("contract", "未知"))
    ar_total = float(_v("ar_total", 0))
    collected = float(_v("collected_total", 0))
    net_balance = float(_v("net_balance", 0))
    uncleared = float(_v("uncleared_balance", 0))
    rate = float(_v("collection_rate", 0))
    cleared_cnt = int(_v("cleared_count", 0))
    uncleared_cnt = int(_v("uncleared_count", 0))
    risk = str(_v("risk_level", "未知"))
    priority = float(_v("priority_score", 0))
    overdue_days = int(_v("max_overdue_days", 0))

    # 转换为万元展示
    ar_wan = ar_total / 10_000
    collected_wan = collected / 10_000
    uncleared_wan = uncleared / 10_000
    net_wan = net_balance / 10_000

    # 风险等级对应建议
    risk_advice = {
        "高风险": (
            "【紧急催收】该业主回款率极低或欠款金额巨大，建议：\n"
            "  1. 立即启动法务催收程序，发送正式催款函\n"
            "  2. 暂停对该业主的新增业务及发货\n"
            "  3. 安排专人与业主高层对接，了解拖欠原因\n"
            "  4. 评估是否需要采取诉讼/仲裁等法律手段\n"
            "  5. 考虑将该业主列入黑名单"
        ),
        "预警": (
            "【加强催收】该业主存在较大回收风险，建议：\n"
            "  1. 加大催收力度，每周跟踪回款进度\n"
            "  2. 与业主协商制定分期还款计划\n"
            "  3. 评估业主的偿付能力和信用状况\n"
            "  4. 对新订单考虑要求预付款或保函"
        ),
        "关注": (
            "【持续跟踪】该业主回款情况一般，建议：\n"
            "  1. 保持正常催收节奏，定期对账\n"
            "  2. 关注业主经营动态，防范潜在风险\n"
            "  3. 对逾期款项及时发函提醒"
        ),
        "正常": (
            "【维护关系】该业主回款良好，建议：\n"
            "  1. 保持良好合作关系\n"
            "  2. 可作为优质客户给予适当信用支持\n"
            "  3. 关注是否有多付款项（预收）需要清退"
        ),
    }

    advice = risk_advice.get(risk, "请人工复核该业主情况。")

    # 判断是否存在预收
    pre_receive_note = ""
    if net_balance < 0:
        pre_receive_note = (
            f"\n> 注：该业主存在预收款 {-net_wan:.2f} 万元，"
            "可能有多收或提前回款的情况，建议核实。\n"
        )

    # ---- 预收预警行 ----
    pre_receive_row = ""
    pre_receive_note = ""
    if str(risk) == "预收预警":
        pre_val = _v("pre_receive_amount", 0)
        ar_val  = _v("ar_total", 0)
        if ar_val > 0:
            pre_ratio = pre_val / ar_val * 100
            pre_receive_row = f"| 预收预警 | 预收{pre_val/10000:,.2f}万元，占应收{pre_ratio:.2f}% | 超过10%阈值，关注资金占用 |\n    "
            pre_receive_note = f"**预收预警说明:** 该业主预收金额{pre_val:,.2f}元，占应收总额{pre_ratio:.2f}%，超过10%预警阈值。建议关注预收资金占用情况，避免影响现金流。\n\n"

    report = f"""## 业主诊断报告

### 基本信息
| 项目 | 内容 |
|------|------|
| 客户编码 | {account} |
| 利润中心 | {profit_center} |
| 合同号 | {contract} |
| 风险等级 | **{risk}** |
{pre_receive_row}    | 催收优先级分数 | {priority:.2f} |

### 核心指标
| 指标 | 金额（万元） | 说明 |
|------|-------------|------|
| 应收总额 | {ar_wan:,.2f} | 累计开票金额（借方） |
| 已收回金额 | {collected_wan:,.2f} | 累计收回金额（贷方） |
| 净余额 | {net_wan:,.2f} | 应收 - 已收回 |
| **未清余额** | **{uncleared_wan:,.2f}** | 实际尚未收回的应收 |
| 回款率 | **{rate:.2f}%** | 已收回 / 应收总额 |
    | 最大逾期天数 | **{overdue_days} 天** | 未清交易中逾期最久的一笔 |

### 笔数统计
- 已清笔数：{cleared_cnt} 笔
- 未清笔数：{uncleared_cnt} 笔
- 合计：{cleared_cnt + uncleared_cnt} 笔

{pre_receive_note}
### 建议措施
{advice}

---
*报告生成时间: {datetime.now().strftime("%Y-%m-%d %H:%M")}*
"""
    return report


# ===========================================================================
# 5. 综合报告生成
# ===========================================================================
def generate_full_report(result: dict) -> str:
    """
    生成完整的应收账款分析报告（Markdown 格式）。

    Parameters
    ----------
    result : dict
        analyze_fbl5n 函数返回的分析结果字典

    Returns
    -------
    str
        完整的 Markdown 格式分析报告
    """
    stats = result.get("stats", {})
    summary = result.get("owner_summary", pd.DataFrame())

    report = f"""# 应收账款分析报告

> 分析日期: {stats.get('analysis_date', 'N/A')}

---

## 一、整体概况

| 指标 | 数值 |
|------|------|
| 应收总额 | {stats.get('total_ar', 0) / 10_000:,.2f} 万元 |
| 已收回金额 | {stats.get('total_collected', 0) / 10_000:,.2f} 万元 |
| 净余额 | {stats.get('net_balance', 0) / 10_000:,.2f} 万元 |
| **未清余额** | **{stats.get('uncleared_balance', 0) / 10_000:,.2f} 万元** |
| 综合回款率 | **{stats.get('collection_rate', 0):.2f}%** |
| 业主数量 | {stats.get('owner_count', 0)} 家 |

## 二、风险分布

| 风险等级 | 业主数量 | 占比 |
|----------|---------|------|
| 高风险 | {stats.get('high_risk_count', 0)} | {stats.get('high_risk_count', 0) / max(stats.get('owner_count', 1), 1) * 100:.2f}% |
| 预警 | {stats.get('warning_count', 0)} | {stats.get('warning_count', 0) / max(stats.get('owner_count', 1), 1) * 100:.2f}% |
| 关注 | {stats.get('attention_count', 0)} | {stats.get('attention_count', 0) / max(stats.get('owner_count', 1), 1) * 100:.2f}% |
| 正常 | {stats.get('normal_count', 0)} | {stats.get('normal_count', 0) / max(stats.get('owner_count', 1), 1) * 100:.2f}% |

## 三、重点关注业主（Top 10）

"""
    if not summary.empty:
        top10 = summary.head(10)
        report += "| 排名 | 客户编码 | 利润中心 | 未清余额(万元) | 回款率 | 风险等级 |\n"
        report += "|------|----------|----------|---------------|--------|----------|\n"
        for i, (_, row) in enumerate(top10.iterrows(), 1):
            report += (
                f"| {i} | {row['account']} | {row['profit_center']} "
                f"| {row['uncleared_balance'] / 10_000:,.2f} "
                f"| {row['collection_rate']:.2f}% "
                f"| {row['risk_level']} |\n"
            )
    else:
        report += "暂无数据\n"

    report += f"""

---
*本报告由 receivable_analyzer 自动生成*
"""
    return report


# ===========================================================================
# 6. 主入口（示例用法）
# ===========================================================================
def main(file_path: str, reference_date: datetime | None = None) -> dict:
    """
    主入口函数：从文件读取数据并执行完整分析流程。

    Parameters
    ----------
    file_path : str
        SAP FBL5N 导出文件路径（.xlsx 或 .csv）
    reference_date : datetime, optional
        分析基准日期，默认当天

    Returns
    -------
    dict
        包含 analysis_result、figures、reports 的完整结果
    """
    logger.info("开始分析文件: %s", file_path)

    # 读取文件
    if file_path.lower().endswith(".csv"):
        raw_df = pd.read_csv(file_path, dtype=str)
    else:
        raw_df = pd.read_excel(file_path, dtype=str)

    logger.info("原始数据: %d 行 x %d 列", len(raw_df), len(raw_df.columns))

    # 标准化
    clean_df = standardize_fbl5n(raw_df)
    logger.info("清洗后数据: %d 行 x %d 列", len(clean_df), len(clean_df.columns))

    # 分析
    result = analyze_fbl5n(clean_df, reference_date)
    stats = result["stats"]
    logger.info(
        "分析完成: 业主=%d, 应收总额=%.2f, 回款率=%.2f%%",
        stats["owner_count"], stats["total_ar"], stats["collection_rate"]
    )

    # 生成图表
    figures = {
        "risk_gauge": create_risk_gauge(stats["collection_rate"]),
        "owner_comparison": create_owner_comparison(result["owner_summary"]),
        "ar_waterfall": create_ar_waterfall(result["owner_summary"]),
    }

    # 如果有业主，生成第一个业主的未清明细
    if not result["owner_summary"].empty:
        first_owner = result["owner_summary"].iloc[0]["owner_key"]
        figures["uncleared_detail"] = create_uncleared_detail(
            result["detail"], first_owner
        )

    # 生成报告
    full_report = generate_full_report(result)

    return {
        "analysis_result": result,
        "figures": figures,
        "reports": {
            "full_report": full_report,
        },
    }


# ---------------------------------------------------------------------------
# 直接运行时的示例
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    # 示例：创建模拟数据并测试
    print("=" * 60)
    print("SAP FBL5N 应收账款分析模块")
    print("=" * 60)

    np.random.seed(42)
    n_rows = 500

    sample_data = pd.DataFrame({
        "凭证编号": [f"DOC{i:06d}" for i in range(n_rows)],
        "凭证类型": np.random.choice(["Z2", "Z5", "YA", "Z4"], n_rows),
        "本币金额": np.concatenate([
            np.random.uniform(100_000, 5_000_000, n_rows // 2),
            np.random.uniform(-5_000_000, -50_000, n_rows // 2),
        ]),
        "科目": np.random.choice([f"101{i:04d}" for i in range(20)], n_rows),
        "利润中心": np.random.choice([f"L100400{i:03d}" for i in range(10)], n_rows),
        "合同": np.random.choice([f"HT2025-{i:04d}" for i in range(30)], n_rows),
        "总账科目": np.random.choice(
            ["1122010100", "1125030000"], n_rows, p=[0.7, 0.3]
        ),
        "清帐日期": np.random.choice(
            [None, "2025-01-15", "2025-02-20", "2025-03-10", "2025-04-05"],
            n_rows, p=[0.3, 0.2, 0.2, 0.2, 0.1]
        ),
        "清帐凭证": [f"CLR{i:06d}" if np.random.random() > 0.3 else None
                     for i in range(n_rows)],
    })

    # 插入一些汇总行（应被过滤）
    summary_rows = pd.DataFrame({
        "凭证编号": ["", "", ""],
        "凭证类型": ["", "", ""],
        "本币金额": [0, 0, 0],
        "科目": ["", "", ""],
        "利润中心": ["", "", ""],
        "合同": ["", "", ""],
        "总账科目": ["", "", ""],
        "已清项目/未清项目符号": ["@5C\\Q未核算@", "@5B\\Q结清的@", "@5B\\Q结清的@"],
    })
    sample_data = pd.concat([sample_data, summary_rows], ignore_index=True)

    print(f"\n模拟数据: {len(sample_data)} 行")

    # 测试完整流程
    clean = standardize_fbl5n(sample_data)
    print(f"清洗后: {len(clean)} 行")

    result = analyze_fbl5n(clean)
    stats = result["stats"]
    print("\n--- 分析结果 ---")
    print(f"业主数量: {stats['owner_count']}")
    print(f"应收总额: {stats['total_ar']:,.2f} 元")
    print(f"已收回:   {stats['total_collected']:,.2f} 元")
    print(f"未清余额: {stats['uncleared_balance']:,.2f} 元")
    print(f"综合回款率: {stats['collection_rate']:.2f}%")
    print(f"高风险业主: {stats['high_risk_count']} 家")

    print("\n--- 业主汇总（Top 5）---")
    top5 = result["owner_summary"].head(5)[[
        "account", "profit_center", "uncleared_balance",
        "collection_rate", "risk_level", "priority_score"
    ]]
    print(top5.to_string(index=False))

    # 测试报告生成
    if not result["owner_summary"].empty:
        report = generate_owner_report(result["owner_summary"].iloc[0])
        print("\n--- 首个业主诊断报告 ---")
        print(report[:800] + "...")

    print("\n" + "=" * 60)
    print("测试通过！模块运行正常。")
    print("=" * 60)
