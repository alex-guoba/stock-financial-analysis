"""
导出三大财务报表
"""

from balance import export_balance
from profit import export_profit
from cashflow import export_cacheflow


def export_all(stock_code: str, report_type: str, limits=5, to_excel=False):
    """
    导出三大财务报表（利润表、资产负债表、现金流量表）
    :param stock_code: 股票代码 如 "601919"
    :param report_type: year 或 quarter
    :param limits: 期数，默认5
    :param to_excel: 是否导出Excel，默认False
    """

    export_profit(stock_code, report_type, limits, to_excel)
    export_balance(stock_code, report_type, limits, to_excel)
    export_cacheflow(stock_code, report_type, limits, to_excel)
