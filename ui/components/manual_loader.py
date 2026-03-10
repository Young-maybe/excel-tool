"""
读取 .docx 使用手册为纯文本，用于在 UI 内展示。
不引入第三方依赖（避免环境差异）。
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path


def load_docx_text(docx_path: str) -> str:
    p = Path(docx_path)
    if not p.exists():
        return f"未找到使用手册文件：{p}"

    try:
        with zipfile.ZipFile(p) as z:
            xml = z.read("word/document.xml").decode("utf-8", "ignore")
        # 极简提取：去掉标签，保留可读文本
        text = re.sub(r"<[^>]+>", "", xml)
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        return "\n".join(lines)
    except Exception as exc:
        return f"使用手册读取失败：{exc}"


