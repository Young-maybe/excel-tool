"""
猴面包树 B2B 发货模块的 UI 页面（仅界面逻辑，不做 Excel 处理）。
"""
from typing import Callable, List

from ui.components.folder_task_page import build_folder_task_page


def build_b2b_page(
    title: str,
    handler: Callable[[str, str], List[str]],
    get_output_dir: Callable[[], str | None],
):
    return build_folder_task_page(
        title=f"猴面包树B2B发货 - {title}",
        handler=handler,
        get_output_dir=get_output_dir,
        hint="请选择输入文件夹；输出将写入左下角已设置的输出目录。",
    )


