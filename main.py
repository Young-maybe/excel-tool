"""
程序入口文件
"""
import logging
import os
import site
import sys
import ctypes
from pathlib import Path

LOG_PATH = Path(__file__).resolve().parent / "resources" / "app.log"
ICON_PATH = Path(__file__).resolve().parent / "resources" / "app_icon.png"
ICON_ICO_PATH = Path(__file__).resolve().parent / "resources" / "app_icon.ico"


def ensure_qt_on_path() -> None:
    """在导入 PyQt6 前确保 Qt 动态库目录在 PATH 中。"""
    candidates = []
    try:
        candidates.extend(site.getsitepackages())
    except Exception:
        pass
    try:
        candidates.append(site.getusersitepackages())
    except Exception:
        pass

    for base in candidates:
        qt_bin = Path(base) / "PyQt6" / "Qt6" / "bin"
        if qt_bin.is_dir():
            qt_bin_str = str(qt_bin)
            current_path = os.environ.get("PATH", "")
            if qt_bin_str not in current_path:
                os.environ["PATH"] = qt_bin_str + os.pathsep + current_path
            return


ensure_qt_on_path()

from PyQt6.QtGui import QIcon  # noqa: E402
from PyQt6.QtGui import QImage  # noqa: E402
from PyQt6.QtWidgets import QApplication, QMessageBox  # noqa: E402

from ui.main_window import MainWindow  # noqa: E402


def setup_logging() -> None:
    """配置日志到本地文件，避免重复添加处理器。"""
    LOG_PATH.parent.mkdir(exist_ok=True)
    if logging.getLogger().handlers:
        return
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[logging.FileHandler(LOG_PATH, encoding="utf-8")],
    )


def show_error(message: str) -> None:
    """使用中文对话框提示错误，不暴露技术细节。"""
    msg_box = QMessageBox()
    msg_box.setIcon(QMessageBox.Icon.Critical)
    msg_box.setWindowTitle("程序错误")
    msg_box.setText(message)
    msg_box.exec()


def handle_exception(exc_type, exc_value, exc_traceback) -> None:
    """捕获未处理异常，记录日志并提示用户。"""
    logging.exception("发生未处理异常", exc_info=(exc_type, exc_value, exc_traceback))
    show_error("程序出现错误，详情已记录到日志。")


def ensure_app_icon() -> None:
    """
    确保 resources/app_icon.png 存在。
    说明：用户可能把图标临时放在项目根目录（例如“app图标.png”），后续会删除；
    这里启动时自动复制一份到 resources/ 作为永久资源。
    """
    try:
        ICON_PATH.parent.mkdir(exist_ok=True)
        if ICON_PATH.exists():
            return
        # 在项目根目录寻找“*图标*.png”
        root = Path(__file__).resolve().parent
        for p in root.glob("*图标*.png"):
            try:
                p_bytes = p.read_bytes()
                ICON_PATH.write_bytes(p_bytes)
                return
            except Exception:
                continue
    except Exception:
        pass


def find_icon_fallback() -> Path | None:
    """
    找一个可用的图标路径（resources 优先，其次项目根目录的 *图标*.png）。
    """
    if ICON_PATH.exists():
        return ICON_PATH
    root = Path(__file__).resolve().parent
    for p in root.glob("*图标*.png"):
        if p.exists():
            return p
    return None


def ensure_app_icon_ico() -> None:
    """
    Windows 任务栏对 .ico 更友好；尽量生成 resources/app_icon.ico。
    注意：这一步需要 Qt 图像模块（QImage），因此放在确保 Qt 可导入之后。
    """
    try:
        ICON_ICO_PATH.parent.mkdir(exist_ok=True)
        if ICON_ICO_PATH.exists():
            return
        src = None
        if ICON_PATH.exists():
            src = ICON_PATH
        else:
            src = find_icon_fallback()
        if src is None or not src.exists():
            return
        img = QImage(str(src))
        if img.isNull():
            return
        # Qt 支持保存为 ico（若平台插件不支持则会失败，但不会影响启动）
        img.save(str(ICON_ICO_PATH))
    except Exception:
        pass


def set_windows_app_user_model_id(app_id: str) -> None:
    """
    Windows：设置 AppUserModelID，帮助任务栏显示自定义图标/分组（而不是 python.exe 图标）。
    必须在创建窗口前调用更稳。
    """
    try:
        if sys.platform != "win32":
            return
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(str(app_id))
    except Exception:
        pass


def main() -> int:
    """程序主入口，保证异常被捕获。"""
    setup_logging()
    sys.excepthook = handle_exception

    # Windows 任务栏：尽早设置 AppUserModelID（否则容易沿用 python.exe 的默认图标）
    set_windows_app_user_model_id("com.yangxinpeng.exceltools")

    app = QApplication(sys.argv)
    # 应用图标（任务栏/窗口左上角）
    try:
        ensure_app_icon()
        ensure_app_icon_ico()
        icon = ICON_ICO_PATH if ICON_ICO_PATH.exists() else find_icon_fallback()
        if icon is not None and icon.exists():
            app.setWindowIcon(QIcon(str(icon)))
    except Exception:
        pass
    try:
        window = MainWindow()
        window.show()
        return app.exec()
    except Exception:
        logging.exception("程序启动失败")
        show_error("程序启动失败，请查看日志文件。")
        return 1


if __name__ == "__main__":
    sys.exit(main())

