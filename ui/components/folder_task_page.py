"""
通用：文件夹输入 → 统一输出目录 的任务页面。
"""
import logging
from time import perf_counter
from typing import Callable, List, Optional

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import QLabel, QMessageBox, QPushButton, QTextEdit, QVBoxLayout, QWidget

from ui.components.apple_progress_panel import AppleProgressPanel
from ui.components.task_runner import TaskWorker, run_in_thread
from services.usage_stats import BASELINE_FIXED, record_event
from ui.components.dialogs import get_existing_directory


def build_folder_task_page(
    title: str,
    handler: Callable[[str, str], List[str]],
    get_output_dir: Callable[[], str | None],
    hint: str,
    manual_text: Optional[str] = None,
    stats_key: Optional[str] = None,
    baseline_provider: Optional[Callable[[str, str], float]] = None,
    baseline_from_outputs: Optional[Callable[[str, str, list], float]] = None,
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
        manual_title = QLabel("使用手册（输入文件夹需要包含什么）")
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

    input_path_label = QLabel("未选择输入文件夹")
    input_path_label.setStyleSheet("color: #2f5f86; font-weight: 700;")
    input_path_label.setWordWrap(True)

    progress_panel = AppleProgressPanel(title="处理进度", show_log=True)
    progress_panel.reset()

    # 保存线程/worker引用，防止被 GC（否则会出现“点了开始处理但任务不执行/进度不动”）
    thread_holder = {"thread": None, "worker": None}
    run_ctx = {"t0": None, "input": None, "output": None}

    def choose_input_dir() -> None:
        directory = get_existing_directory(widget, "选择输入文件夹", start_dir="")
        if directory:
            input_path_label.setText(str(directory))

    def run_task() -> None:
        input_dir = input_path_label.text()
        if not input_dir or input_dir == "未选择输入文件夹":
            QMessageBox.warning(widget, "输入未选择", "请先选择输入文件夹。")
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
        run_ctx["input"] = input_dir
        run_ctx["output"] = output_dir

        worker = TaskWorker(handler=handler, args=[input_dir, output_dir], kwargs={})
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
                    in_dir = str(run_ctx.get("input") or input_dir)
                    out_dir = str(run_ctx.get("output") or output_dir)
                    if baseline_from_outputs:
                        baseline = float(baseline_from_outputs(in_dir, out_dir, outputs))
                    elif baseline_provider:
                        baseline = float(baseline_provider(in_dir, out_dir))
                    else:
                        baseline = float(BASELINE_FIXED.get(stats_key, 0.0))
                    record_event(stats_key, runtime_sec=runtime, baseline_sec=baseline)
            except Exception:
                pass

            if outputs:
                msg = "处理完成，输出文件：\n" + "\n".join([str(x) for x in outputs])
            else:
                msg = "处理完成，但未生成新文件，请检查输入数据或流程顺序。"
            QMessageBox.information(widget, "完成", msg)

        def on_failed(message: str) -> None:
            progress_panel.set_running(False)
            select_btn.setEnabled(True)
            run_btn.setEnabled(True)
            logging.exception("运行任务失败: %s", message)
            # 保持原来的弹窗分类体验
            if "文件正在被占用" in message or "WinError 32" in message:
                QMessageBox.critical(widget, "文件被占用", message)
                return
            if "缺少必需文件" in message or "未找到" in message:
                QMessageBox.warning(widget, "文件缺失", message)
                return
            QMessageBox.critical(widget, "处理失败", f"处理时出现错误：\n{message}\n\n请查看日志或检查输入文件。")

        worker.signals.finished.connect(on_finished)
        worker.signals.failed.connect(on_failed)
        thread_holder["worker"] = worker
        thread_holder["thread"] = run_in_thread(worker)

    select_btn = QPushButton("选择输入文件夹")
    select_btn.setObjectName("taskSecondaryButton")
    select_btn.setCursor(Qt.CursorShape.PointingHandCursor)
    select_btn.clicked.connect(choose_input_dir)

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


