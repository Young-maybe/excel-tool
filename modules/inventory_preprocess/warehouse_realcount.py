"""
盘点的初步处理 - 仓库实盘表处理工具3

说明：
- 你提供的源文件文件夹后续会删除，所以这里改为直接调用我们内部复制进来的实现：
  `modules/inventory_preprocess/_tool_warehouse3.py`
- 仅做 I/O 适配：输入由 UI 选取文件夹；输出写到程序统一输出目录下的独立子目录。
- 核心处理逻辑保持与原脚本一致。
"""
from __future__ import annotations

import shutil
from pathlib import Path
from typing import Callable, List, Optional


def _load_source_module():
    # 保持函数名不变，便于最小改动；实际返回内部实现模块
    from . import _tool_warehouse3 as tool

    return tool


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
    work_dir = out_dir / "盘点的初步处理" / "仓库实盘表处理工具3"
    work_dir.mkdir(parents=True, exist_ok=True)

    # 将输入目录中相关 Excel 复制到 output_dir，再按原逻辑在 output_dir 内处理
    # - 信选/清元源文件
    report(10, "扫描信选/清元源文件…")
    candidates = []
    for p in in_dir.glob("*.xlsx"):
        stem = p.stem.lower()
        if stem.startswith("信选") or stem.startswith("清元") or stem in {"信选", "清元"}:
            candidates.append(p)

    report(14, f"待复制信选/清元：{len(candidates)} 个")
    report(16, "复制信选/清元源文件到工作目录…")
    total_c = max(len(candidates), 1)
    for i, p in enumerate(sorted(candidates), start=1):
        if progress_cb and (i == 1 or i == total_c or i % 3 == 0):
            report(16 + int(i / total_c * 10), f"复制 {i}/{total_c}：{p.name}")
        try:
            shutil.copy2(p, work_dir / p.name)
        except PermissionError:
            pass

    # 规格代码映射
    match_src = in_dir / "盘点-规格代码匹配.xlsx"
    if match_src.exists():
        try:
            shutil.copy2(match_src, work_dir / match_src.name)
        except PermissionError:
            pass

    # 商品库存导出（可能多个，全部复制更稳）
    report(20, "复制商品库存导出/映射文件…")
    for p in in_dir.glob("商品库存导出*.xlsx"):
        try:
            shutil.copy2(p, work_dir / p.name)
        except PermissionError:
            pass

    src = _load_source_module()

    load_mapping_pandian_spec = getattr(src, "load_mapping_pandian_spec")
    find_files_by_prefix = getattr(src, "find_files_by_prefix")
    make_modified_copy = getattr(src, "make_modified_copy")
    process_one_file = getattr(src, "process_one_file")

    code_mapping = load_mapping_pandian_spec(work_dir)
    targets = find_files_by_prefix(work_dir, prefixes=["信选", "清元"])
    if not targets:
        raise FileNotFoundError("未找到以“信选/清元”开头的源文件（.xlsx）。")

    outputs: List[str] = []
    total = len(targets)
    for i, src_file in enumerate(targets, start=1):
        base = 30 + int((i - 1) / max(total, 1) * 60)
        report(base, f"处理 {i}/{total}：{src_file.name}")
        modified = make_modified_copy(src_file)
        process_one_file(modified, work_dir, code_mapping)
        outputs.append(str(modified))

    # 同时返回被更新后的“商品库存导出*.xlsx”（原逻辑会保存）
    report(92, "收集输出文件…")
    for p in work_dir.glob("商品库存导出*.xlsx"):
        outputs.append(str(p))

    report(100, "完成")
    return outputs


