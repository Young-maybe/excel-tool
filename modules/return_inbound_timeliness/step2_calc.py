"""
退货入库时效 - 表的时效计算

内部实现来源：`modules/return_inbound_timeliness/_tool_step2.py`
仅做 I/O 适配：
- 输入由 UI 选取文件夹（应包含“退货入库时效分析*.xlsx”）
- 输出写入统一输出目录下的独立子目录
不动态加载外部源脚本。
"""

from __future__ import annotations

from pathlib import Path
from typing import Callable, List, Optional

from . import _tool_step2 as tool


def process_step2(
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

    work_dir = out_dir / "退货入库时效" / "表的时效计算"
    work_dir.mkdir(parents=True, exist_ok=True)

    processor = tool.ReturnDataProcessor()
    report(10, "开始计算时效（读取总表）…")
    report(20, "处理中…（此阶段由核心脚本主导，进度会相对稀疏）")
    ok, msg = processor.process(str(in_dir), str(work_dir))
    if not ok:
        raise RuntimeError(msg)

    report(85, "收集输出文件…")
    outputs = sorted(work_dir.glob("退货商品明细汇总-*月推单.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not outputs:
        # 兜底：返回 work_dir 下最新的 xlsx
        outputs = sorted(work_dir.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not outputs:
        raise FileNotFoundError("未找到输出文件（.xlsx）。")

    report(100, "完成")
    return [str(outputs[0])]


