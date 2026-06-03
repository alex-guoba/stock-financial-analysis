"""导出现金流量表"""

import akshare as ak
import pandas as pd

from comm import ReportType, get_full_code, to_markdown


def get_xueqiu_style_cacheflow(
    stock_code: str, unit: float, report_type: ReportType, limits=5
):
    """
    抓取个股现金流量表，字段对标雪球现金流量表，支持筛选报表类型
    :param stock_code: 纯数字股票代码，如 "601919"
    :param unit: 单位，如 1e8
    :param report_type: ReportType.YEAR / ReportType.QUARTER
    :param limits: 显示行数，默认5
    :return: 清洗后的DataFrame，可直接复制Excel
    """

    full_code = get_full_code(stock_code)

    # A 股接口（东方财富）
    if report_type == ReportType.YEAR:
        raw_df = ak.stock_cash_flow_sheet_by_yearly_em(symbol=full_code)
    elif report_type == ReportType.QUARTER:
        raw_df = ak.stock_cash_flow_sheet_by_quarterly_em(symbol=full_code)
    else:
        raise ValueError(f"Invalid report_type: {report_type}")

    if raw_df.empty:
        raise ValueError("未获取到现金流量表数据")

    # debug_print_df(raw_df)

    stock_name = raw_df["SECURITY_NAME_ABBR"].iloc[0]

    # 东方财富 现金流量表 字段中英文对照
    field_mapping = {
        "REPORT_DATE": "报告期",
        "REPORT_DATE_NAME": "报告期名称",
        # 经营现金流
        "SALES_SERVICES": "  销售商品、提供劳务收到的现金",
        "RECEIVE_TAX_REFUND": "  收到的税费返还",
        "RECEIVE_OTHER_OPERATE": "  收到其他与经营活动有关现金",
        "TOTAL_OPERATE_INFLOW": "经营活动现金流入小计",
        "BUY_SERVICES": "  购买商品、接受劳务支付的现金",
        "PAY_STAFF_CASH": "  支付给职工以及为职工支付现金",
        "PAY_ALL_TAX": "  支付的各项税费",
        "PAY_OTHER_OPERATE": "  支付其他与经营活动有关现金",
        "TOTAL_OPERATE_OUTFLOW": "经营活动现金流出小计",
        "NETCASH_OPERATE": "经营活动现金流净额",
        "_EMPTY_OP_IN_OTHER": "",
        # 投资现金流
        "WITHDRAW_INVEST": "  收回投资收到的现金",
        "RECEIVE_INVEST_INCOME": "  取得投资收益收到现金",
        "DISPOSAL_LONG_ASSET": "  处置长期资产收回现金",
        "RECEIVE_OTHER_INVEST": "  收到其他投资现金",
        "TOTAL_INVEST_INFLOW": "投资活动现金流入小计",
        "CONSTRUCT_LONG_ASSET": "  购建长期资产支付现金(资本开支)",
        "INVEST_PAY_CASH": "  投资支付的现金",
        "TOTAL_INVEST_OUTFLOW": "投资活动现金流出小计",
        "NETCASH_INVEST": "投资活动现金流净额",
        "_EMPTY_INV_IN_OTHER": "",
        # 筹资现金流
        "ACCEPT_INVEST_CASH": "  吸收投资收到现金",
        "RECEIVE_LOAN_CASH": "  取得借款收到现金",
        "ISSUE_BOND": "  发行债券收到现金",
        "RECEIVE_OTHER_FINANCE": "  收到其他筹资现金",
        "TOTAL_FINANCE_INFLOW": "筹资活动现金流入小计",
        "PAY_DEBT_CASH": "  偿还债务支付现金",
        "ASSIGN_DIVIDEND_PORFIT": "  分配股利、利息支付现金",
        "PAY_OTHER_FINANCE": "  支付其他筹资现金",
        "TOTAL_FINANCE_OUTFLOW": "筹资活动现金流出小计",
        "NETCASH_FINANCE": "筹资活动现金流净额",
        "_EMPTY_FIN_IN_OTHER": "",
        "RATE_CHANGE_EFFECT": "汇率变动对现金影响",
        "CCE_ADD": "现金及等价物净增加额",
        "BEGIN_CCE": "期初现金及等价物余额",
        "END_CCE": "期末现金及等价物余额",
        # 关键附表项目（可选展示）
        "_EMPTY_NOTE_SPLIT": "净利润调节经营现金流(附表)",
        "NETPROFIT": "  净利润",
        "ASSET_IMPAIRMENT": "  资产减值损失",
        "FA_IR_DEPR": "  固定资产折旧",
        "IA_AMORTIZE": "  无形资产摊销",
        "USERIGHT_ASSET_AMORTIZE": "  使用权资产摊销",
        "FINANCE_EXPENSE": "  财务费用",
        "INVEST_LOSS": "  投资收益(-)",
        "INVENTORY_REDUCE": "  存货增减",
        "OPERATE_RECE_REDUCE": "  经营性应收增减",
        "OPERATE_PAYABLE_ADD": "  经营性应付增减",
        "NETCASH_OPERATENOTE": "附表-经营现金流净额",
    }

    df = raw_df.reindex(columns=list(field_mapping.keys()), fill_value=0).copy()
    df["REPORT_DATE"] = pd.to_datetime(df["REPORT_DATE"]).dt.strftime("%Y-%m-%d")
    df = df.sort_values("REPORT_DATE", ascending=True).tail(limits)

    num_cols = [c for c in df.columns if df[c].dtype == "float64"]
    for col in num_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").div(unit).round(2)

    # 空行占位
    for c in df.columns:
        if c.startswith("_EMPTY_"):
            df[c] = None

    df.rename(columns=field_mapping, inplace=True)
    df = df.set_index("报告期").T
    return stock_name, df


def export_cacheflow(stock_code: str, report_type: str, limits=5, to_excel=False):
    """
    导出现金流量表
    :param stock_code: 股票代码 如 "601919"
    :param report_type: year 或 quarter
    :param limits: 期数，默认5
    :param to_excel: 是否导出Excel，默认False
    """
    stock_name, df = get_xueqiu_style_cacheflow(
        str(stock_code), 1e8, ReportType(report_type), limits
    )
    title = f"现金流量表（单位：亿元） {stock_name}({stock_code})"

    print(to_markdown(df, title))
    if to_excel:
        file_name = f"{stock_code}_现金流量表_{report_type}.xlsx"
        df.to_excel(file_name, sheet_name="现金流量表", index=True)
        print(f"✅ 已导出：{file_name}")
