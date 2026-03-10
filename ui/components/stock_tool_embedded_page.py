from __future__ import annotations

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import QLabel, QVBoxLayout, QWidget


class StockToolEmbeddedPage(QWidget):
    """
    将“库存处理工具”原界面嵌入到我们右侧内容区（允许被 Apple 样式影响）。
    """

    def __init__(self, tab_index: int, title: str):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # 顶部标题（可选）
        title_label = QLabel(title)
        title_label.setFont(QFont("Microsoft YaHei UI", 14, QFont.Weight.Bold))
        title_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
        title_label.setContentsMargins(24, 18, 24, 10)

        from ui.components.stock_tool_legacy_main import create_embedded_widget

        embedded = create_embedded_widget(initial_tab_index=tab_index)

        layout.addWidget(title_label)
        layout.addWidget(embedded)


