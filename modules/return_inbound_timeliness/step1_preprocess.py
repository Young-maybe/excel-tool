"""
退货入库时效 - 表的初步处理

内部实现来源：`modules/return_inbound_timeliness/_tool_step1.py`
仅做 I/O 适配：
- 输入由 UI 选取文件夹
- 输出写入统一输出目录下的独立子目录
不动态加载外部源脚本。
"""

from __future__ import annotations

import shutil
from pathlib import Path
from typing import Callable, List, Optional

from . import _tool_step1 as tool


def process_step1(
    input_dir: str,
    output_dir: str,
    progress_cb: Optional[Callable[[int, str], None]] = None,
) -> List[str]:
    def report(p: int, s: str) -> None:
        if progress_cb:
            progress_cb(p, s)

    in_dir = Path(input_dir)
    out_dir = Path(output_dir)
    if not in_dir.exists():
        raise FileNotFoundError(f"输入目录不存在：{input_dir}")

    work_dir = out_dir / "退货入库时效" / "表的初步处理"
    work_dir.mkdir(parents=True, exist_ok=True)

    # 必需文件（按原脚本的固定文件名/模板名准备到 work_dir，保证核心逻辑不变）
    report(5, "检查输入文件…")
    csv_candidates = sorted(in_dir.glob("退货商品明细汇总*.csv"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not csv_candidates:
        raise FileNotFoundError("缺少必需文件：退货商品明细汇总*.csv（至少需要一个）")
    csv_src = csv_candidates[0]
    csv_dst = work_dir / "退货商品明细汇总.csv"

    cfg_src = in_dir / "店铺匹配仓库配置.xlsx"
    if not cfg_src.exists():
        raise FileNotFoundError("缺少必需文件：店铺匹配仓库配置.xlsx")
    cfg_dst = work_dir / cfg_src.name

    tpl_src = in_dir / "12345.xlsx"
    if not tpl_src.exists():
        raise FileNotFoundError("缺少必需文件：12345.xlsx（模板文件）")
    tpl_dst = work_dir / tpl_src.name

    report(20, "复制输入文件到工作目录…")
    try:
        shutil.copy2(csv_src, csv_dst)
        shutil.copy2(cfg_src, cfg_dst)
        shutil.copy2(tpl_src, tpl_dst)
    except PermissionError as exc:
        raise PermissionError(f"文件正在被占用，请先关闭Excel/WPS后重试：{exc}") from exc

    # 覆盖内部脚本的 BASE_DIR，让其在 work_dir 内按原逻辑读取与输出
    tool.BASE_DIR = work_dir

    report(35, "读取 CSV 与配置…")
    csv_data = tool.read_csv_data()
    if csv_data is None or getattr(csv_data, "empty", False):
        raise RuntimeError("未能读取到有效的CSV数据，处理中止。")

    report(45, "读取仓库映射配置…")
    warehouse_mapping = tool.read_warehouse_config()

    report(55, "解析模板并计算…")
    template_wb, template_ws, template_headers, hidden_columns = tool.analyze_template_excel()
    if template_wb is None:
        raise FileNotFoundError("未能读取到有效的模板Excel文件：12345.xlsx")

    report(70, "填充模板并生成输出 Excel…")
    output_filename = tool.create_output_excel(
        csv_data, template_wb, template_ws, template_headers, hidden_columns, warehouse_mapping
    )
    if not output_filename:
        raise RuntimeError("处理失败：未生成输出文件。")

    out_path = work_dir / str(output_filename)
    if not out_path.exists():
        # 兜底：若内部脚本返回的是文件名但保存路径异常
        candidates = sorted(work_dir.glob("退货入库时效分析*月总表.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
        if candidates:
            out_path = candidates[0]

    if not out_path.exists():
        raise FileNotFoundError("未找到输出文件：退货入库时效分析*月总表.xlsx")

    report(100, "完成")
    return [str(out_path)]


