"""导出资产负债表"""

import akshare as ak
import pandas as pd

from comm import ReportType, get_full_code, to_markdown


def get_xueqiu_style_balance(
    stock_code: str, unit: float, report_type: ReportType, limits=5
):
    """
    抓取个股资产负债表，字段对标雪球资产负债表，支持筛选报表类型
    :param stock_code: 纯数字股票代码，如 "601919"
    :param unit: 单位，如 1e8
    :param report_type: ReportType.YEAR / ReportType.QUARTER
    :param limits: 显示行数，默认5
    :return: 清洗后的DataFrame，可直接复制Excel
    """

    full_code = get_full_code(stock_code)

    # A 股接口（东方财富）
    if report_type == ReportType.YEAR:
        raw_df = ak.stock_balance_sheet_by_yearly_em(symbol=full_code)
    elif report_type == ReportType.QUARTER:
        raw_df = ak.stock_balance_sheet_by_report_em(symbol=full_code)
    else:
        raise ValueError(f"Invalid report_type: {report_type}")

    if raw_df.empty:
        raise ValueError("未获取到资产负债表数据")

    stock_name = raw_df["SECURITY_NAME_ABBR"].iloc[0]

    # 东方财富 资产负债表 字段中英文对照
    field_mapping = {
        # 基础信息
        # "SECUCODE": "证券代码",
        # "SECURITY_CODE": "股票代码",
        # "SECURITY_NAME_ABBR": "股票简称",
        # "ORG_CODE": "机构代码",
        # "ORG_TYPE": "机构类型",
        "REPORT_DATE": "报告期",
        # "REPORT_TYPE": "报告类型",
        "REPORT_DATE_NAME": "报告期名称",
        # "SECURITY_TYPE_CODE": "证券类型",
        # "NOTICE_DATE": "公告日期",
        # "UPDATE_DATE": "更新日期",
        # "CURRENCY": "货币类型",
        # 流动资产
        "TOTAL_CURRENT_ASSETS": "流动资产合计",
        "MONETARYFUNDS": "  货币资金",
        "NOTE_ACCOUNTS_RECE": "  应收票据及应收账款",
        "NOTE_RECE": "    应收票据",
        "ACCOUNTS_RECE": "    应收账款",
        # "FINANCE_RECE": "应收款项融资",
        # "PREMIUM_RECE": "应收保费",
        # "INTEREST_RECE": "应收利息",
        # "DIVIDEND_RECE": "应收股利",
        "PREPAYMENT": "  预付款项",
        "INVENTORY": "  存货",
        "CONTRACT_ASSET": "  合同资产",
        # "OTHER_RECE": "其他应收款",
        "TOTAL_OTHER_RECE": "  其他应收款合计",
        # "EXPORT_REFUND_RECE": "  应收出口退税",
        "NONCURRENT_ASSET_1YEAR": "  一年内到期的非流动资产",
        "OTHER_CURRENT_ASSET": "  其他流动资产",
        # 非流动资产
        "TOTAL_NONCURRENT_ASSETS": "非流动资产合计",
        "LONG_EQUITY_INVEST": "  长期股权投资",
        "OTHER_EQUITY_INVEST": "  其他权益工具投资",
        "OTHER_NONCURRENT_FINASSET": "  其他非流动金融资产",
        # "INVEST_REALESTATE": "投资性房地产",
        "FIXED_ASSET": "  固定资产",
        "CIP": "  在建工程",
        # "PROJECT_MATERIAL": "工程物资",
        "USERIGHT_ASSET": "  使用权资产",
        "INTANGIBLE_ASSET": "  无形资产",
        # "DEVELOP_EXPENSE": "开发支出",
        # "GOODWILL": "商誉",
        "LONG_PREPAID_EXPENSE": "  长期待摊费用",
        "DEFER_TAX_ASSET": "  递延所得税资产",
        "OTHER_NONCURRENT_ASSET": "  其他非流动资产",
        # 总资产
        "TOTAL_ASSETS": "总资产",
        # 流动负债
        "TOTAL_CURRENT_LIAB": "流动负债合计",
        "SHORT_LOAN": "  短期借款",
        "NOTE_ACCOUNTS_PAYABLE": "  应付票据及应付账款",
        # "NOTE_PAYABLE": "  应付票据",
        # "ACCOUNTS_PAYABLE": "  应付账款",
        "CONTRACT_LIAB": "  合同负债",
        "ADVANCE_RECEIVABLES": "  预收款项",
        "STAFF_SALARY_PAYABLE": "  应付职工薪酬",
        "TAX_PAYABLE": "  应交税费",
        # "INTEREST_PAYABLE": "应付利息",
        # "DIVIDEND_PAYABLE": "应付股利",
        # "OTHER_PAYABLE": "其他应付款",
        "TOTAL_OTHER_PAYABLE": "  其他应付款合计",
        "NONCURRENT_LIAB_1YEAR": "  一年内到期的非流动负债",
        "OTHER_CURRENT_LIAB": "  其他流动负债",
        # "DEFER_INCOME": "递延收益",
        # "DEFER_INCOME_1YEAR": "递延收益(一年以内)",
        # "SHORT_BOND_PAYABLE": "短期应付债券",
        # 非流动负债
        "TOTAL_NONCURRENT_LIAB": "非流动负债合计",
        "LONG_LOAN": "  长期借款",
        "BOND_PAYABLE": "  应付债券",
        "LEASE_LIAB": "  租赁负债",
        "LONG_PAYABLE": "  长期应付款",
        "LONG_STAFFSALARY_PAYABLE": "  长期应付职工薪酬",
        "PREDICT_LIAB": "  预计负债",
        "DEFER_TAX_LIAB": "  递延所得税负债",
        "NONCURRENT_LIAB_OTHER": "  其他非流动负债",
        # 总负债
        "TOTAL_LIABILITIES": "总负债",
        # 所有者权益
        "_EMPTY_EQUITY": "所有者权益(或股东权益)",
        # "SHARE_CAPITAL": "  实收资本(或股本)",
        "CAPITAL_RESERVE": "  资本公积",
        "OTHER_COMPRE_INCOME": "  其他综合收益",
        "SPECIAL_RESERVE": "  专项储备",
        "SURPLUS_RESERVE": "  盈余公积",
        "UNASSIGN_RPOFIT": "  未分配利润",
        "TOTAL_EQUITY": "所有者权益合计",
        "TOTAL_PARENT_EQUITY": "  归属于母公司股东权益合计",
        "MINORITY_EQUITY": "  少数股东权益",
        # 负债及权益总计
        "TOTAL_LIAB_EQUITY": "负债和所有者权益总计",
        "_EMPTY_RATIO": "重点指标",
        "DEBT_TO_ASSERTS_RATIO": "  资产负债率（%）",
        "LIQUIDITY_RATIO": "  流动比率",
        "QUICK_RATIO": "  速动比率",
    }

    df = raw_df.reindex(columns=list(field_mapping.keys()), fill_value=0).copy()
    df["REPORT_DATE"] = pd.to_datetime(df["REPORT_DATE"]).dt.strftime("%Y-%m-%d")
    df = df.sort_values("REPORT_DATE", ascending=True).tail(limits)

    num_cols = [c for c in df.columns if df[c].dtype == "float64"]
    for col in num_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").div(unit).round(2)

    # 计算资产负债率、流动比率
    df["DEBT_TO_ASSERTS_RATIO"] = (
        df["TOTAL_LIABILITIES"] / (df["TOTAL_ASSETS"]) * 100
    ).round(2)
    df["LIQUIDITY_RATIO"] = (
        df["TOTAL_CURRENT_ASSETS"] / df["TOTAL_CURRENT_LIAB"]
    ).round(2)
    df["QUICK_RATIO"] = (
        (df["TOTAL_CURRENT_ASSETS"] - df["INVENTORY"]) / df["TOTAL_CURRENT_LIAB"]
    ).round(2)

    # 空行占位
    for c in df.columns:
        if c.startswith("_EMPTY_"):
            df[c] = None

    df.rename(columns=field_mapping, inplace=True)
    df = df.set_index("报告期").T
    return stock_name, df


def export_balance(stock_code: str, report_type: str, limits=5, to_excel=False):
    """
    导出资产负债表
    :param stock_code: 股票代码 如 "601919"
    :param report_type: year 或 quarter
    :param limits: 期数，默认5
    :param to_excel: 是否导出Excel，默认False
    """
    file_name = f"{stock_code}_资产负债表_{report_type}.xlsx" if to_excel else None
    stock_name, df = get_xueqiu_style_balance(
        str(stock_code), 1e8, ReportType(report_type), limits
    )

    title = f"资产负债表（单位：亿元） {stock_name}({stock_code})"
    print(to_markdown(df, title))

    if file_name:
        df.to_excel(file_name, sheet_name="资产负债表", index=True)
        print(f"✅ 已导出：{file_name}")
