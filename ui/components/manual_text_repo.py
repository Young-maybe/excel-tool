"""
操作手册文本仓库：从项目根目录的 huancuntxt/ 读取各模块说明文本。

目的：
- 手册内容可直接编辑维护（不需要改代码）
- UI 运行时加载显示

注意：
- 默认按 UTF-8/UTF-8-SIG 读取；必要时回退 GBK
"""

from __future__ import annotations

import sys
from pathlib import Path


def _project_root() -> Path:
    # 兼容：源码运行 vs PyInstaller
    if hasattr(sys, "frozen") and getattr(sys, "frozen"):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parents[2]


def load_manual_text(pinyin_name: str) -> str:
    """
    从 huancuntxt/<pinyin_name>.txt 读取文本。不存在则返回空串。
    """
    root = _project_root()
    path = root / "huancuntxt" / f"{pinyin_name}.txt"
    if not path.exists():
        return ""

    for enc in ("utf-8-sig", "utf-8", "gbk", "gb18030"):
        try:
            return path.read_text(encoding=enc)
        except Exception:
            continue
    try:
        return path.read_bytes().decode("utf-8", errors="ignore")
    except Exception:
        return ""


