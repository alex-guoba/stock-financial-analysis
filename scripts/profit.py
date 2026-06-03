"""导出利润表
1. 获取利润表
2. 清洗数据
3. 保存到Excel
"""

import akshare as ak
import pandas as pd
import numpy as np

from comm import ReportType, get_full_code, to_markdown


def get_xueqiu_style_profit(
    stock_code: str, unit: float, report_type: ReportType, limits=5
):
    """
    抓取个股利润表，字段对标雪球利润表，支持筛选报表类型
    :param stock_code: 纯数字股票代码，如 "601919"
    :param unit: 单位，如 "100000000"
    :param report_type: 报表筛选类型
        year   → 仅年报（12月31日）
        quarter → 单季报（3月31日、6月30日、9月30日、12月31日）
    :param limits: 显示行数，默认5
    :return: 清洗后的DataFrame，可直接复制Excel
    """
    # 市场前缀
    full_code = get_full_code(stock_code)

    # 获取利润表
    if report_type == ReportType.YEAR:
        raw_df = ak.stock_profit_sheet_by_yearly_em(symbol=full_code)
    elif report_type == ReportType.QUARTER:
        raw_df = ak.stock_profit_sheet_by_quarterly_em(symbol=full_code)
    else:
        raise ValueError(f"Invalid report_type: {report_type}")

    if raw_df.empty:
        raise ValueError("No data")

    stock_name = raw_df["SECURITY_NAME_ABBR"].iloc[0]

    # 雪球标准字段映射
    field_mapping = {
        "REPORT_DATE": "报告期",
        "REPORT_DATE_NAME": "报告期名称",
        "TOTAL_OPERATE_INCOME": "营业总收入",
        # （经营类）利息收入对于大部分公司无意义。
        # "INTEREST_INCOME": "其中：利息收入",
        "OPERATE_INCOME": "  营业收入",
        # 营业总成本 = 营业成本 + 税金及附加 + 销售费用 + 管理费用 + 研发费用 + 财务费用
        "TOTAL_OPERATE_COST": "营业总成本",
        "OPERATE_COST": "  营业成本",
        "OPERATE_TAX_ADD": "  营业税金及附加",
        "SALE_EXPENSE": "  销售费用",
        "MANAGE_EXPENSE": "  管理费用",
        "RESEARCH_EXPENSE": "  研发费用",
        "FINANCE_EXPENSE": "  财务费用",
        "FE_INTEREST_EXPENSE": "    财务费用-利息费用",
        "FE_INTEREST_INCOME": "    财务费用-利息收入",
        ##
        ## 2018 前：TotalImpair = ASSET_IMPAIRMENT_LOSS
        # "ASSET_IMPAIRMENT_LOSS": "资产减值损失",
        # "CREDIT_IMPAIRMENT_LOSS": "信用减值损失",
        ##
        # 其他经营类收益（信用减值损失、资产减值损失、其他收益、投资收益、公允价值变动收益、资产处置收益）
        "_EMPTY_OPERATE_INCOME": "其他经营类收益",
        "FAIRVALUE_CHANGE_INCOME": "  加：公允价值变动收益",
        "INVEST_INCOME": "    投资收益",
        # "INVEST_JOINT_INCOME": "    其中:对联营企业和合营企业的投资收益",
        "ASSET_DISPOSAL_INCOME": "  资产处置收益",
        "ASSET_IMPAIRMENT_INCOME": "  资产减值损失（新）",
        "CREDIT_IMPAIRMENT_INCOME": "  信用减值损失（新）",
        "OTHER_INCOME": "  其他收益",
        "OPERATE_PROFIT": "营业利润",
        "NONBUSINESS_INCOME": "  加：营业外收入",
        "NONBUSINESS_EXPENSE": "  减：营业外支出",
        "TOTAL_PROFIT": "利润总额",
        "INCOME_TAX": "  减：所得税费用",
        "NETPROFIT": "净利润",
        "NETPROFIT_YOY": "净利润yoy(%)",
        ## (一)按经营持续性分类：净利润 = 持续经营净利润 + 终止经营净利润
        # "CONTINUED_NETPROFIT": "持续经营净利润",
        # "DISCONTINUED_NETPROFIT": "终止经营净利润",
        ## （二）按所有权归属分类: 净利润 = 归母净利润 + 少数股东损益
        "PARENT_NETPROFIT": "  归属于母公司所有者的净利润",
        "MINORITY_INTEREST": "  少数股东损益",
        ## 剔除一次性收益
        "DEDUCT_PARENT_NETPROFIT": "扣非净利润",
        "_EMPTY_PROFIT_RATE_EMPTY": "利润率",
        "GROSS_MARGIN_RATE": "  毛利率(%)",
        "NETPROFIT_RATE": "  净利率(%)",
        "_EMPTY_EPS_EMPTY": "EPS(元)",
        "BASIC_EPS": "  基本每股收益(元)",
        "DILUTED_EPS": "  稀释每股收益(元)",
    }

    # 筛选
    df = raw_df.reindex(columns=list(field_mapping.keys()), fill_value=0).copy()

    # 日期处理
    df["REPORT_DATE"] = pd.to_datetime(df["REPORT_DATE"]).dt.strftime("%Y-%m-%d")

    # 从旧到新排序，取最近5期
    df = df.sort_values("REPORT_DATE", ascending=True).tail(limits)

    # 数值列转数字
    num_cols = [
        c
        for c in df.columns
        if c
        not in [
            "BASIC_EPS",
            "DILUTED_EPS",
            "NETPROFIT_YOY",
        ]
        and df[c].dtype == "float64"
    ]
    for col in num_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").div(unit).round(2)

    # 计算毛利润、净利率等
    df["GROSS_MARGIN_RATE"] = np.where(
        df["TOTAL_OPERATE_INCOME"] == 0,
        0,
        (
            (df["TOTAL_OPERATE_INCOME"] - df["OPERATE_COST"])
            / df["TOTAL_OPERATE_INCOME"]
            * 100
        ).round(2),
    )
    df["NETPROFIT_RATE"] = np.where(
        df["TOTAL_OPERATE_INCOME"] == 0,
        0,
        (df["NETPROFIT"] / df["TOTAL_OPERATE_INCOME"] * 100).round(2),
    )
    df["NETPROFIT_YOY"] = df["NETPROFIT_YOY"].round(2)

    # 空行占位
    for c in df.columns:
        if c.startswith("_EMPTY_"):
            df[c] = None

    # 转置时保留 财务项目名称 作为第一列
    df.rename(columns=field_mapping, inplace=True)
    df = df.set_index("报告期").T  # 把报告期当列名，转置后项目名保留

    return stock_name, df


def process_profit(
    stock_code: str, report_type: ReportType, limits: int, file_name: str | None
):
    """
    导出利润表
    """
    stock_name, profit_df = get_xueqiu_style_profit(
        stock_code, 1e8, report_type, limits=limits
    )

    title = f"利润表（单位：亿元） {stock_name}({stock_code})"

    print(to_markdown(profit_df, title))
    if file_name:
        profit_df.to_excel(file_name, sheet_name=f"{stock_name}-利润表", index=False)
        print(f"\n✅ 导出成功：{file_name}")


def export_profit(
    stock_code: str, report_type: str, limits: int = 5, to_excel: bool = False
):
    """
    导出利润表
    :param stock_code: 股票代码 如 "601919"
    :param report_type: year 或 quarter
    :param limits: 期数，默认5
    :param to_excel: 是否导出Excel，默认False
    """
    file_name = None
    if to_excel:
        file_name = f"{stock_code}_利润表_{report_type}.xlsx"

    process_profit(str(stock_code), ReportType(report_type), limits, file_name)
