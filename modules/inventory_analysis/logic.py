"""
盘点的分析处理

说明：
- 你提供的源文件文件夹后续会删除，所以这里改为直接调用我们内部复制进来的实现：
  `modules/inventory_analysis/_tool_analysis.py`
- 仅做 I/O 适配：输入由 UI 选取文件夹；输出写到程序统一输出目录下的独立子目录。
- 核心处理逻辑保持与原脚本一致。
"""
from __future__ import annotations

import shutil
from pathlib import Path
from typing import Callable, List, Optional

from . import _tool_analysis as tool


def process_folder(
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
    work_dir = out_dir / "盘点的分析处理"
    work_dir.mkdir(parents=True, exist_ok=True)

    # 原脚本要求“程序与数据位于同一文件夹”，这里采用“复制 Excel 到输出目录再跑”的方式，
    # 保证核心逻辑不变且所有输出落在 output_dir。
    report(8, "扫描输入 Excel…")
    xlsx_files = list(in_dir.glob("*.xlsx"))
    xls_files = list(in_dir.glob("*.xls"))
    report(12, f"准备复制：.xlsx={len(xlsx_files)} .xls={len(xls_files)}")

    report(15, "复制 .xlsx 到工作目录…")
    total_xlsx = max(len(xlsx_files), 1)
    for i, p in enumerate(sorted(xlsx_files), start=1):
        if progress_cb and (i == 1 or i == total_xlsx or i % 5 == 0):
            report(15 + int(i / total_xlsx * 20), f"复制 .xlsx：{i}/{total_xlsx} {p.name}")
        try:
            shutil.copy2(p, work_dir / p.name)
        except PermissionError:
            pass

    report(35, "复制 .xls 到工作目录…")
    total_xls = max(len(xls_files), 1)
    for i, p in enumerate(sorted(xls_files), start=1):
        if progress_cb and (i == 1 or i == total_xls or i % 5 == 0):
            report(35 + int(i / total_xls * 10), f"复制 .xls：{i}/{total_xls} {p.name}")
        try:
            shutil.copy2(p, work_dir / p.name)
        except PermissionError:
            pass

    # 覆盖源脚本 BASE_DIR，指向 work_dir（其内已包含输入文件副本）
    setattr(tool, "BASE_DIR", work_dir)
    report(50, "开始执行盘点分析逻辑（可能需要几分钟）…")
    tool.process_workbook()

    report(92, "收集输出文件…")
    outputs = [str(p) for p in work_dir.glob("*-修改后.xlsx")]
    if not outputs:
        # 源脚本可能未找到主文件
        raise FileNotFoundError("未生成“*-修改后.xlsx”，请确认输入文件夹包含“附件一：自营仓盘点表?月.xlsx”等所需文件。")
    report(100, "完成")
    return outputs


