"""
统一的文件/文件夹选择对话框（马卡龙主题）。

说明：
- Windows 原生文件选择框无法完全被 QSS 控制；为了统一风格，这里使用 Qt 自带的非原生对话框。
- 保持简单：只提供“选择文件夹”的封装（当前需求用到）。
"""

from __future__ import annotations

from typing import Optional

from PyQt6.QtCore import Qt, QStandardPaths, QUrl
from PyQt6.QtWidgets import QFileDialog, QWidget


def _macaron_dialog_qss() -> str:
    # 尽量轻量：只改常见控件的背景/按钮/选中态
    return """
    QFileDialog {
      background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
        stop:0 rgba(255, 245, 250, 1),
        stop:1 rgba(238, 249, 255, 1));
    }
    QWidget {
      font-family: "Microsoft YaHei UI", "Segoe UI", sans-serif;
      font-size: 12px;
      color: #6d4c4c;
    }
    QLineEdit {
      background: rgba(255, 255, 255, 0.85);
      border: 2px solid rgba(231, 183, 198, 0.75);
      border-radius: 12px;
      padding: 6px 10px;
      color: #2f5f86;
    }
    QTreeView, QListView {
      background: rgba(255, 255, 255, 0.75);
      border: 2px solid rgba(231, 183, 198, 0.55);
      border-radius: 14px;
      padding: 6px;
      selection-background-color: rgba(133, 180, 255, 0.30);
      selection-color: #2f5f86;
    }
    QPushButton {
      padding: 8px 12px;
      border-radius: 14px;
      border: 2px solid rgba(133, 180, 255, 0.75);
      background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 rgba(227, 246, 255, 0.95),
        stop:1 rgba(246, 232, 255, 0.85));
      font-weight: 800;
      color: #2f5f86;
      min-height: 28px;
    }
    QPushButton:hover {
      background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 rgba(214, 242, 255, 1),
        stop:1 rgba(240, 225, 255, 1));
    }
    QDialogButtonBox QPushButton {
      min-width: 90px;
    }
    QScrollBar:vertical {
      background: rgba(255, 255, 255, 0.35);
      width: 12px;
      margin: 2px;
      border-radius: 6px;
    }
    QScrollBar::handle:vertical {
      background: rgba(133, 180, 255, 0.65);
      min-height: 28px;
      border-radius: 6px;
    }
    QScrollBar::handle:vertical:hover {
      background: rgba(133, 180, 255, 0.85);
    }
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
      height: 0px;
    }
    QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
      background: transparent;
    }
    """


def get_existing_directory(parent: QWidget, title: str, start_dir: str = "") -> Optional[str]:
    dialog = QFileDialog(parent, title, start_dir)
    dialog.setFileMode(QFileDialog.FileMode.Directory)
    dialog.setOption(QFileDialog.Option.ShowDirsOnly, True)
    dialog.setOption(QFileDialog.Option.DontResolveSymlinks, True)
    dialog.setOption(QFileDialog.Option.DontUseNativeDialog, True)
    dialog.setWindowModality(Qt.WindowModality.ApplicationModal)
    dialog.setStyleSheet(_macaron_dialog_qss())

    # 默认打开“桌面”，并把桌面加入左侧快捷入口（你截图红框区域）
    desktop = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DesktopLocation)
    home = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.HomeLocation)
    documents = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DocumentsLocation)
    downloads = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DownloadLocation)

    # 起始目录：优先外部传入；否则默认桌面；再不行回退到用户目录
    if start_dir:
        dialog.setDirectory(start_dir)
    elif desktop:
        dialog.setDirectory(desktop)
    elif home:
        dialog.setDirectory(home)

    # 侧边栏快捷入口：把桌面放第一个，方便最快点击
    sidebar_urls: list[QUrl] = []
    if desktop:
        sidebar_urls.append(QUrl.fromLocalFile(desktop))
    if documents:
        sidebar_urls.append(QUrl.fromLocalFile(documents))
    if downloads:
        sidebar_urls.append(QUrl.fromLocalFile(downloads))
    if home:
        sidebar_urls.append(QUrl.fromLocalFile(home))
    if sidebar_urls:
        dialog.setSidebarUrls(sidebar_urls)

    if dialog.exec():
        files = dialog.selectedFiles()
        if files:
            return str(files[0])
    return None


