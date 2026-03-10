"""
使用统计（二级页面）

需求：
- 从主界面进入后展示报表，并且每次进入都实时刷新
- 左上角返回按钮回到主界面
- 所有时长用“小时”为单位展示
- 报表：按“大模块总计节省时长”降序；同时展示子模块节省
"""

from __future__ import annotations

from typing import Callable

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from services.usage_stats import build_report, hours_str


class UsageStatsPage(QWidget):
    def __init__(self, on_back: Callable[[], None]) -> None:
        super().__init__()
        self.setObjectName("usageStatsPage")
        self._on_back = on_back

        # 二级页面独立样式：更偏蓝的马卡龙风格
        self.setStyleSheet(
            """
            QWidget#usageStatsPage {
              background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 rgba(228, 246, 255, 1),
                stop:1 rgba(210, 235, 255, 1));
            }
            QPushButton#backButton {
              padding: 8px 12px;
              border-radius: 16px;
              border: 2px solid rgba(133, 180, 255, 0.85);
              background: rgba(255, 255, 255, 0.78);
              color: #2f5f86;
              font-weight: 900;
              min-width: 90px;
            }
            QPushButton#backButton:hover {
              background: rgba(214, 242, 255, 0.95);
            }
            QPushButton#backButton:pressed {
              background: rgba(133, 180, 255, 0.25);
            }
            QLabel#statsTitle {
              color: #2f5f86;
            }
            QLabel#statsTotalLabel {
              color: #2f5f86;
              font-weight: 900;
            }
            QTreeWidget#statsTable {
              background: rgba(255, 255, 255, 0.85);
              border: 2px solid rgba(133, 180, 255, 0.55);
              border-radius: 18px;
              padding: 6px;
            }
            QHeaderView::section {
              background: rgba(133, 180, 255, 0.18);
              color: #2f5f86;
              padding: 8px 10px;
              border: none;
              font-weight: 900;
            }
            QTreeWidget::item {
              padding: 8px 6px;
            }
            QTreeWidget::item:selected {
              background: rgba(133, 180, 255, 0.28);
              color: #2f5f86;
            }
            QTreeWidget::item:hover {
              background: rgba(231, 183, 198, 0.16);
            }
            QLabel#statsBottomTotal {
              color: #2f5f86;
              padding: 12px;
              border-radius: 18px;
              background: rgba(255, 255, 255, 0.65);
              border: 2px solid rgba(231, 183, 198, 0.65);
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
        )

        root = QVBoxLayout(self)
        root.setContentsMargins(20, 20, 20, 20)
        root.setSpacing(12)

        top = QHBoxLayout()
        top.setSpacing(10)

        self.back_btn = QPushButton("← 返回")
        self.back_btn.setObjectName("backButton")
        self.back_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.back_btn.clicked.connect(self._on_back)

        title = QLabel("使用统计")
        title.setFont(QFont("Microsoft YaHei UI", 18, QFont.Weight.Bold))
        title.setObjectName("statsTitle")

        top.addWidget(self.back_btn, 0, Qt.AlignmentFlag.AlignLeft)
        top.addWidget(title, 0, Qt.AlignmentFlag.AlignLeft)
        top.addStretch(1)

        self.total_label = QLabel("总计为您节省：0.00 小时")
        self.total_label.setFont(QFont("Microsoft YaHei UI", 12, QFont.Weight.Bold))
        self.total_label.setObjectName("statsTotalLabel")

        self.table = QTreeWidget()
        self.table.setObjectName("statsTable")
        self.table.setHeaderLabels(["模块", "节省(小时)", "运行次数"])
        self.table.setRootIsDecorated(True)
        self.table.setAlternatingRowColors(True)
        self.table.setIndentation(14)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        inner = QWidget()
        inner_layout = QVBoxLayout(inner)
        inner_layout.setContentsMargins(0, 0, 0, 0)
        inner_layout.addWidget(self.table)
        scroll.setWidget(inner)

        self.bottom_total = QLabel("您使用该程序以来共计节省了 0.00 小时")
        self.bottom_total.setFont(QFont("Microsoft YaHei UI", 16, QFont.Weight.Bold))
        self.bottom_total.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.bottom_total.setObjectName("statsBottomTotal")

        root.addLayout(top)
        root.addWidget(self.total_label)
        root.addWidget(scroll, 1)
        root.addWidget(self.bottom_total)

        self.refresh()

    def refresh(self) -> None:
        report = build_report()
        self.total_label.setText(f"总计为您节省：{hours_str(report.total_saved_sec)} 小时")
        self.bottom_total.setText(f"您使用该程序以来共计节省了 {hours_str(report.total_saved_sec)} 小时")

        self.table.clear()

        if not report.groups_sorted:
            QTreeWidgetItem(self.table, ["暂无统计数据（请先运行任意模块）", "0.00", "0"])
            self.table.expandAll()
            return

        for g in report.groups_sorted:
            total_h = hours_str(report.group_totals.get(g, 0.0))
            parent = QTreeWidgetItem([g, total_h, ""])
            parent.setFont(0, QFont("Microsoft YaHei UI", 11, QFont.Weight.Bold))
            self.table.addTopLevelItem(parent)

            children = report.modules_by_group.get(g, [])
            for m in children:
                child = QTreeWidgetItem([m.label, hours_str(m.total_saved_sec), str(m.runs)])
                parent.addChild(child)

        self.table.expandAll()


