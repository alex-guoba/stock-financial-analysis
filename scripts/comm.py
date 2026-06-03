"""
公共模块
"""

from enum import Enum
import pandas as pd


class ReportType(Enum):
    """
    报告类型
    """

    QUARTER = "quarter"
    YEAR = "year"


def get_full_code(stock_code: str) -> str:
    """
    获取股票完整代码，格式为6位数字
    """
    if stock_code.startswith("6"):
        return f"SH{stock_code}"
    elif stock_code.startswith(("0", "3")):
        return f"SZ{stock_code}"
    else:
        return stock_code


def debug_print_df(df: pd.DataFrame):
    """
    打印DataFrame的所有列，用于调试
    """
    for col in df.columns:
        print(f"{col}: {df[col].dtype} {df[col].iloc[0]}")


def to_markdown(df: pd.DataFrame, title: str = "") -> str:
    """
    把 DataFrame 转换成 Markdown 表格
    """
    md = df.to_markdown(numalign="left", stralign="left", tablefmt="github")

    if title:
        return f"## {title}\n{md}\n"

    return f"{md}\n"
