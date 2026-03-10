"""
通用：单个文件输入 → 统一输出目录 的任务页面。

用于“发货时效表处理”等场景：用户从 UI 选择一个文件（如 CSV），然后运行处理逻辑。
"""

import logging
from time import perf_counter
from typing import Callable, List, Optional

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import QFileDialog, QLabel, QMessageBox, QPushButton, QTextEdit, QVBoxLayout, QWidget

from ui.components.apple_progress_panel import AppleProgressPanel
from ui.components.task_runner import TaskWorker, run_in_thread
from services.usage_stats import BASELINE_FIXED, record_event


def build_file_task_page(
    title: str,
    handler: Callable[[str, str], List[str]],
    get_output_dir: Callable[[], str | None],
    hint: str,
    manual_text: Optional[str] = None,
    file_filter: str = "CSV Files (*.csv);;All Files (*)",
    stats_key: Optional[str] = None,
    baseline_provider: Optional[Callable[[str, str], float]] = None,
) -> QWidget:
    widget = QWidget()
    widget.setMinimumWidth(700)
    layout = QVBoxLayout()
    layout.setContentsMargins(40, 40, 40, 40)
    layout.setSpacing(15)
    widget.setLayout(layout)

    title_label = QLabel(title)
    title_label.setFont(QFont("Microsoft YaHei UI", 24, QFont.Weight.Bold))
    title_label.setStyleSheet("color: #6d4c4c; margin-bottom: 10px;")
    title_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
    title_label.setWordWrap(True)

    hint_label = QLabel(hint)
    hint_label.setFont(QFont("Microsoft YaHei UI", 13))
    hint_label.setStyleSheet("color: #8b6f76; margin-bottom: 10px;")
    hint_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
    hint_label.setWordWrap(True)

    if manual_text:
        manual_title = QLabel("使用手册（输入文件需要符合什么）")
        manual_title.setFont(QFont("Microsoft YaHei UI", 12, QFont.Weight.Bold))
        manual_title.setStyleSheet("color: #6d4c4c; margin-top: 10px;")
        manual_title.setAlignment(Qt.AlignmentFlag.AlignLeft)

        manual_box = QTextEdit()
        manual_box.setReadOnly(True)
        manual_box.setPlainText(manual_text)
        manual_box.setMinimumHeight(180)
        manual_box.setStyleSheet(
            "QTextEdit { background: rgba(255, 248, 251, 0.9); border: 2px solid rgba(231, 183, 198, 0.85); border-radius: 16px; padding: 12px; color: #6d4c4c; }"
        )

    input_path_label = QLabel("未选择输入文件")
    input_path_label.setStyleSheet("color: #2f5f86; font-weight: 700;")
    input_path_label.setWordWrap(True)

    progress_panel = AppleProgressPanel(title="处理进度", show_log=True)
    progress_panel.reset()

    # 保存线程/worker引用，防止被 GC（否则会出现“点了开始处理但任务不执行/进度不动”）
    thread_holder = {"thread": None, "worker": None}
    run_ctx = {"t0": None, "input": None, "output": None}

    def choose_input_file() -> None:
        filename, _ = QFileDialog.getOpenFileName(
            widget,
            "选择输入文件",
            "",
            file_filter,
        )
        if filename:
            input_path_label.setText(filename)

    def run_task() -> None:
        input_file = input_path_label.text()
        if not input_file or input_file == "未选择输入文件":
            QMessageBox.warning(widget, "输入未选择", "请先选择输入文件。")
            return

        output_dir = get_output_dir()
        if not output_dir:
            QMessageBox.warning(widget, "输出未设置", "请先点击左下角“设置输出文件夹”。")
            return

        select_btn.setEnabled(False)
        run_btn.setEnabled(False)
        progress_panel.reset()
        progress_panel.set_running(True)
        progress_panel.set_progress(0)
        progress_panel.set_status("开始处理…", append_log=True)
        run_ctx["t0"] = perf_counter()
        run_ctx["input"] = input_file
        run_ctx["output"] = output_dir

        worker = TaskWorker(handler=handler, args=[input_file, output_dir], kwargs={})
        progress_panel.bind_worker(worker.signals)

        def on_finished(outputs: list) -> None:
            progress_panel.set_running(False)
            select_btn.setEnabled(True)
            run_btn.setEnabled(True)

            # 记录使用统计（成功才记；统计失败不影响主流程）
            try:
                if stats_key:
                    t0 = run_ctx.get("t0") or perf_counter()
                    runtime = perf_counter() - float(t0)
                    in_path = str(run_ctx.get("input") or input_file)
                    out_dir = str(run_ctx.get("output") or output_dir)
                    if baseline_provider:
                        baseline = float(baseline_provider(in_path, out_dir))
                    else:
                        baseline = float(BASELINE_FIXED.get(stats_key, 0.0))
                    record_event(stats_key, runtime_sec=runtime, baseline_sec=baseline)
            except Exception:
                pass

            if outputs:
                msg = "处理完成，输出文件：\n" + "\n".join([str(x) for x in outputs])
            else:
                msg = "处理完成，但未生成新文件，请检查输入数据。"
            QMessageBox.information(widget, "完成", msg)

        def on_failed(message: str) -> None:
            progress_panel.set_running(False)
            select_btn.setEnabled(True)
            run_btn.setEnabled(True)
            logging.exception("运行任务失败: %s", message)
            if "文件正在被占用" in message or "WinError 32" in message:
                QMessageBox.critical(widget, "文件被占用", message)
                return
            QMessageBox.warning(widget, "输入不符合要求", message)

        worker.signals.finished.connect(on_finished)
        worker.signals.failed.connect(on_failed)
        thread_holder["worker"] = worker
        thread_holder["thread"] = run_in_thread(worker)

    select_btn = QPushButton("选择输入文件")
    select_btn.setObjectName("taskSecondaryButton")
    select_btn.setCursor(Qt.CursorShape.PointingHandCursor)
    select_btn.clicked.connect(choose_input_file)

    run_btn = QPushButton("开始处理")
    run_btn.setObjectName("taskPrimaryButton")
    run_btn.setCursor(Qt.CursorShape.PointingHandCursor)
    run_btn.clicked.connect(run_task)

    layout.addWidget(title_label)
    layout.addWidget(hint_label)
    if manual_text:
        layout.addWidget(manual_title)
        layout.addWidget(manual_box)
    layout.addWidget(progress_panel)
    layout.addWidget(input_path_label)
    layout.addWidget(select_btn)
    layout.addWidget(run_btn)
    layout.addStretch()
    return widget


