"""
盘点的初步处理 - 百合基础表处理工具2

说明：
- 你提供的源文件文件夹后续会删除，所以这里改为直接调用我们内部复制进来的实现：
  `modules/inventory_preprocess/_tool_baihe2.py`
- 仅做 I/O 适配：输入由 UI 选取文件夹；输出写到程序统一输出目录下的独立子目录。
- 核心处理逻辑保持与原脚本一致。
"""
from __future__ import annotations

import shutil
from pathlib import Path
from typing import Callable, List, Optional


from . import _tool_baihe2 as tool


def _latest_file(directory: Path, pattern: str) -> Optional[Path]:
    files = list(directory.glob(pattern))
    if not files:
        return None
    files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return files[0]


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
    work_dir = out_dir / "盘点的初步处理" / "百合基础表处理工具2"
    work_dir.mkdir(parents=True, exist_ok=True)

    # 选择最新库存快照明细
    report(5, "扫描输入文件…")
    inv_src = _latest_file(in_dir, "库存快照明细*.xlsx")
    if inv_src is None:
        raise FileNotFoundError("输入文件夹中未找到“库存快照明细*.xlsx”。")

    # 必需：盘点-规格代码匹配.xlsx
    match_src = in_dir / "盘点-规格代码匹配.xlsx"
    if not match_src.exists():
        raise FileNotFoundError("缺少必需文件：盘点-规格代码匹配.xlsx")

    # 必需：商品库存导出*.xlsx（用于透视表补商品名称等）
    sku_src = _latest_file(in_dir, "商品库存导出*.xlsx")
    if sku_src is None:
        raise FileNotFoundError("缺少必需文件：商品库存导出*.xlsx（至少需要一个）")

    # 将所需源文件复制到 output_dir 中再处理（确保所有输出落在 output_dir）
    inv_path = work_dir / inv_src.name
    report(20, "复制输入文件到工作目录…")
    try:
        shutil.copy2(inv_src, inv_path)
    except PermissionError as exc:
        raise PermissionError(f"文件正在被占用，请先关闭Excel/WPS后重试：{inv_src}") from exc

    try:
        shutil.copy2(match_src, work_dir / match_src.name)
    except PermissionError as exc:
        raise PermissionError(f"文件正在被占用，请先关闭Excel/WPS后重试：{match_src}") from exc

    try:
        shutil.copy2(sku_src, work_dir / sku_src.name)
    except PermissionError as exc:
        raise PermissionError(f"文件正在被占用，请先关闭Excel/WPS后重试：{sku_src}") from exc

    report(45, "开始处理库存快照（透视/匹配/回写，可能需要几分钟）…")
    report(55, "处理中…（此阶段以核心脚本为主，进度会相对稀疏）")
    ok, msg = tool.process_inventory_file(inv_path, work_dir)
    if not ok:
        raise RuntimeError(msg)

    report(95, "生成输出文件…")
    modified_path = work_dir / f"{inv_src.stem}-修改后.xlsx"
    report(100, "完成")
    return [str(modified_path)]


