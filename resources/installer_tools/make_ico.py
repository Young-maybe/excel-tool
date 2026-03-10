"""
将 resources/app_icon.png 转为 resources/app_icon.ico（Windows 任务栏/安装包更友好）。
优先使用 PyQt6 的 QImage（不额外依赖 Pillow）。
"""

from __future__ import annotations

from pathlib import Path

from PyQt6.QtGui import QImage


def main() -> int:
    root = Path(__file__).resolve().parents[2]
    png_path = root / "resources" / "app_icon.png"
    ico_path = root / "resources" / "app_icon.ico"

    if not png_path.exists():
        print(f"[ERR] 未找到图标：{png_path}")
        return 1

    img = QImage(str(png_path))
    if img.isNull():
        print("[ERR] QImage 读取失败，请确认 png 文件有效")
        return 1

    ok = img.save(str(ico_path))
    if not ok:
        print("[ERR] 保存 ico 失败（可能是 Qt 插件不支持 ico 保存）。")
        return 1

    print(f"[OK] 已生成：{ico_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


