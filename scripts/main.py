"""
A股股票分析器
"""

import fire

from profit import export_profit
from balance import export_balance
from cashflow import export_cacheflow
from all import export_all

if __name__ == "__main__":

    fire.Fire(
        {
            "profit": export_profit,
            "balance": export_balance,
            "cacheflow": export_cacheflow,
            "all": export_all,
        }
    )
