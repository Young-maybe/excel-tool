"""
发货时效表处理

只暴露对外的处理入口：输入 CSV 文件路径 + 输出目录 → 生成 Excel。
"""

from .logic import process_csv

__all__ = ["process_csv"]


