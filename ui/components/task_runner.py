"""
后台任务执行器：把耗时处理放到 QThread 里跑，避免卡 UI。

设计目标：
- UI 只负责：选择输入、启动任务、展示进度/状态、显示结果
- 业务模块只需要（可选）调用 progress_cb(percent, status) 即可上报
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable, List, Optional

from PyQt6.QtCore import QObject, QThread, pyqtSignal


ProgressCallback = Callable[[int, str], None]


class TaskSignals(QObject):
    progress_updated = pyqtSignal(int)
    status_updated = pyqtSignal(str)
    finished = pyqtSignal(list)  # outputs: List[str]
    failed = pyqtSignal(str)  # message


@dataclass
class TaskContext:
    """
    传给业务函数的上下文（目前只放 progress_cb；后续需要可扩展）。
    """

    progress_cb: Optional[ProgressCallback] = None


class TaskWorker(QObject):
    """
    在后台线程执行 handler(input, output, progress_cb=?).
    """

    def __init__(self, handler: Callable[..., List[str]], args: list[Any], kwargs: dict[str, Any]) -> None:
        super().__init__()
        self._handler = handler
        self._args = args
        self._kwargs = kwargs
        self.signals = TaskSignals()

    def _make_progress_cb(self) -> ProgressCallback:
        def _cb(percent: int, status: str) -> None:
            try:
                self.signals.progress_updated.emit(int(percent))
            except Exception:
                pass
            if status is not None:
                self.signals.status_updated.emit(str(status))

        return _cb

    def run(self) -> None:
        try:
            progress_cb = self._make_progress_cb()
            # 先发一个“已启动”信号，避免界面长期停留 0% 造成误解
            self.signals.progress_updated.emit(1)
            self.signals.status_updated.emit("任务已启动，正在执行…")

            # 兼容：如果业务函数不接受 progress_cb（历史版本），就重试不带该参数
            try:
                outputs = self._handler(*self._args, **{**self._kwargs, "progress_cb": progress_cb})
            except TypeError as exc:
                # 只在“确实是不支持 progress_cb 参数”时回退；避免吞掉业务内部的 TypeError
                msg = str(exc)
                if "progress_cb" in msg and ("unexpected keyword argument" in msg or "got an unexpected keyword" in msg):
                    outputs = self._handler(*self._args, **self._kwargs)
                else:
                    raise

            if outputs is None:
                outputs = []
            self.signals.progress_updated.emit(100)
            self.signals.status_updated.emit("完成")
            self.signals.finished.emit(list(outputs))
        except Exception as exc:
            self.signals.failed.emit(str(exc))


def run_in_thread(worker: TaskWorker) -> QThread:
    """
    启动线程并返回 QThread（由调用方保存引用，防止被 GC）。
    """
    thread = QThread()
    worker.moveToThread(thread)
    thread.started.connect(worker.run)
    worker.signals.finished.connect(thread.quit)
    worker.signals.failed.connect(thread.quit)
    worker.signals.finished.connect(worker.deleteLater)
    worker.signals.failed.connect(worker.deleteLater)
    thread.finished.connect(thread.deleteLater)
    thread.start()
    return thread


