"""
Apple 风格的进度与状态提示面板（可复用组件）。

用途：
- 为“耗时任务”提供：百分比进度条 + 当前状态文本 + 详细日志（可选）
- 可直接连接到 QThread / QObject 的信号：
  - progress_updated(int)
  - status_updated(str)
  - finished_signal(bool, str)  # 可选

注意：
- 该文件当前不被任何页面引用（按用户要求：先学习/沉淀，后续再接入各模块）。
"""

from __future__ import annotations

from typing import Any, Optional

from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QProgressBar,
    QSizePolicy,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)


class AppleProgressPanel(QWidget):
    """
    一个轻量的“进度 + 状态日志”组件，适配 Apple 风格 UI。
    """

    def __init__(
        self,
        title: str = "处理进度",
        show_log: bool = True,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)

        self._spinner_timer: Optional[QTimer] = None
        self._spinner_phase = 0

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)

        card = QFrame()
        card.setObjectName("appleProgressCard")
        card.setFrameShape(QFrame.Shape.NoFrame)
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(16, 14, 16, 14)
        card_layout.setSpacing(10)

        # 标题 + 运行指示
        header_row = QHBoxLayout()
        header_row.setContentsMargins(0, 0, 0, 0)
        header_row.setSpacing(10)

        self.title_label = QLabel(title)
        self.title_label.setFont(QFont("Microsoft YaHei UI", 12, QFont.Weight.Bold))

        self.spinner_label = QLabel("")
        self.spinner_label.setObjectName("appleProgressSpinner")
        self.spinner_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.spinner_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)

        header_row.addWidget(self.title_label)
        header_row.addWidget(self.spinner_label)
        card_layout.addLayout(header_row)

        # 当前状态 + 百分比
        status_row = QHBoxLayout()
        status_row.setContentsMargins(0, 0, 0, 0)
        status_row.setSpacing(10)

        self.status_label = QLabel("等待开始…")
        self.status_label.setWordWrap(True)
        self.status_label.setObjectName("appleProgressStatus")

        self.percent_label = QLabel("0%")
        self.percent_label.setObjectName("appleProgressPercent")
        self.percent_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.percent_label.setMinimumWidth(48)

        status_row.addWidget(self.status_label, 1)
        status_row.addWidget(self.percent_label, 0)
        card_layout.addLayout(status_row)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setObjectName("appleProgressBar")
        card_layout.addWidget(self.progress_bar)

        # 日志（可选）
        self.log_box: Optional[QTextEdit] = None
        if show_log:
            self.log_box = QTextEdit()
            self.log_box.setReadOnly(True)
            self.log_box.setMinimumHeight(140)
            self.log_box.setObjectName("appleProgressLog")
            card_layout.addWidget(self.log_box)

        root.addWidget(card)

        # Apple 风格（局部，不依赖 main_window 全局样式）
        self.setStyleSheet(
            """
            QFrame#appleProgressCard {
              background: rgba(255, 255, 255, 0.88);
              border: 2px solid rgba(231, 183, 198, 0.75);
              border-radius: 18px;
            }
            QLabel#appleProgressStatus {
              color: #6d4c4c;
              font-size: 12px;
            }
            QLabel#appleProgressPercent {
              color: #2f5f86;
              font-size: 12px;
              font-weight: 800;
            }
            QLabel#appleProgressSpinner {
              color: #2f5f86;
              font-size: 12px;
            }
            QProgressBar#appleProgressBar {
              background: rgba(255, 255, 255, 0.65);
              border: 2px solid rgba(133, 180, 255, 0.65);
              border-radius: 14px;
              height: 12px;
            }
            QProgressBar#appleProgressBar::chunk {
              background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                stop:0 rgba(133, 180, 255, 0.95),
                stop:1 rgba(106, 210, 255, 0.95));
              border-radius: 14px;
            }
            QTextEdit#appleProgressLog {
              background: rgba(255, 248, 251, 0.92);
              border: 2px solid rgba(231, 183, 198, 0.75);
              border-radius: 16px;
              padding: 12px;
              color: #6d4c4c;
              font-size: 12px;
            }
            """
        )

    # -------- 公共 API --------

    def reset(self) -> None:
        self.set_progress(0)
        self.set_status("等待开始…", append_log=False)
        if self.log_box is not None:
            self.log_box.clear()
        self.set_running(False)

    def set_progress(self, value: int) -> None:
        v = max(0, min(100, int(value)))
        self.progress_bar.setValue(v)
        self.percent_label.setText(f"{v}%")

    def set_status(self, text: str, append_log: bool = True) -> None:
        msg = "" if text is None else str(text)
        self.status_label.setText(msg)
        if append_log and self.log_box is not None and msg:
            self.append_log(msg)

    def append_log(self, text: str) -> None:
        if self.log_box is None:
            return
        msg = "" if text is None else str(text)
        if not msg:
            return
        self.log_box.append(msg)
        self.log_box.moveCursor(self.log_box.textCursor().MoveOperation.End)

    def set_running(self, running: bool) -> None:
        if running:
            if self._spinner_timer is None:
                self._spinner_timer = QTimer(self)
                self._spinner_timer.timeout.connect(self._tick_spinner)
            self._spinner_phase = 0
            self._spinner_timer.start(350)
            self.spinner_label.setText("处理中")
        else:
            if self._spinner_timer is not None:
                self._spinner_timer.stop()
            self.spinner_label.setText("")

    def bind_worker(self, worker: Any) -> None:
        """
        将常见信号绑定到面板：
        - worker.progress_updated(int)
        - worker.status_updated(str)
        - worker.finished_signal(bool, str)  # 可选（老格式）
        """
        if hasattr(worker, "progress_updated"):
            try:
                worker.progress_updated.connect(self.set_progress)  # type: ignore[attr-defined]
            except Exception:
                pass
        if hasattr(worker, "status_updated"):
            try:
                worker.status_updated.connect(lambda s: self.set_status(str(s), append_log=True))  # type: ignore[attr-defined]
            except Exception:
                pass
        if hasattr(worker, "finished_signal"):
            try:
                worker.finished_signal.connect(self._on_finished)  # type: ignore[attr-defined]
            except Exception:
                pass

    # -------- 内部实现 --------

    def _tick_spinner(self) -> None:
        self._spinner_phase = (self._spinner_phase + 1) % 4
        dots = "." * self._spinner_phase
        self.spinner_label.setText(f"处理中{dots}")

    def _on_finished(self, ok: bool, msg: str) -> None:
        self.set_running(False)
        # 保留最终状态到状态栏 & 日志
        if ok:
            self.set_status("完成", append_log=True)
            if msg:
                self.append_log(str(msg))
        else:
            self.set_status("失败", append_log=True)
            if msg:
                self.append_log(str(msg))


