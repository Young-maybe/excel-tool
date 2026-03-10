"""
主窗口文件
"""
from pathlib import Path

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QIcon
from PyQt6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QMainWindow,
    QPushButton,
    QScrollArea,
    QSplitter,
    QStackedWidget,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from modules.b2b_shipping import (
    carton_label_process,
    delivery_and_stock_process,
    picking_slip_process,
    template_match_process,
)
from modules.dangdang_sales import process_folder as dangdang_sales_process
from modules.inventory_analysis import process_folder as inventory_analysis_process
from modules.inventory_preprocess import baihe_process, guanyi_process, warehouse_process
from modules.pushed_order_not_inbound import process_folder as pushed_order_process
from modules.return_inbound_timeliness import process_step1 as return_step1_process, process_step2 as return_step2_process
from modules.shipping_timeliness import process_csv as shipping_timeliness_process
from ui.components.b2b_shipping_pages import build_b2b_page
from ui.components.file_task_page import build_file_task_page
from ui.components.folder_task_page import build_folder_task_page
from ui.components.stock_tool_embedded_page import StockToolEmbeddedPage
from ui.components.manual_text_repo import load_manual_text
from ui.components.usage_stats_page import UsageStatsPage
from services.usage_stats import BASELINE_FIXED
from ui.components.dialogs import get_existing_directory


class MainWindow(QMainWindow):
    """主窗口类，仅包含导航与内容占位。"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel 处理工具")
        self.resize(1200, 800)
        # 你的显示器为 1920*1080，希望最小可缩放到 1/2：960*540
        self.setMinimumSize(960, 540)

        # 标题栏左上角图标
        try:
            root = Path(__file__).resolve().parents[1]
            icon_path = root / "resources" / "app_icon.png"
            if not icon_path.exists():
                # 兜底：用户可能还没运行过一次复制逻辑，先用根目录临时图标
                candidates = sorted(root.glob("*图标*.png"))
                if candidates:
                    icon_path = candidates[0]
            if icon_path.exists():
                self.setWindowIcon(QIcon(str(icon_path)))
        except Exception:
            pass
        
        # 应用 Apple 风格样式
        self._apply_apple_style()

        # 根容器：主界面 + 使用统计（二级页面）
        self.root_stack = QStackedWidget()
        self.setCentralWidget(self.root_stack)

        # 主界面使用 QSplitter 实现可调节宽度的分栏
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setObjectName("mainSplitter")
        splitter.setHandleWidth(3)
        splitter.setChildrenCollapsible(False)

        # 左侧导航面板
        left_panel = QWidget()
        left_panel.setObjectName("leftPanel")
        left_layout = QVBoxLayout()
        left_layout.setContentsMargins(20, 20, 20, 20)
        left_layout.setSpacing(15)
        left_panel.setLayout(left_layout)

        # 标题标签
        title_label = QLabel("功能导航")
        title_label.setObjectName("titleLabel")
        title_label.setFont(QFont("Microsoft YaHei UI", 16, QFont.Weight.Bold))
        left_layout.addWidget(title_label)

        # 导航树
        self.nav_tree = QTreeWidget()
        self.nav_tree.setObjectName("navTree")
        self.nav_tree.setHeaderHidden(True)
        self.nav_tree.setIndentation(10)
        self.nav_tree.setAnimated(True)
        self.nav_tree.setRootIsDecorated(True)

        # 设置按钮
        self.settings_button = QPushButton("⚙ 设置输出文件夹")
        self.settings_button.setObjectName("settingsButton")
        self.settings_button.clicked.connect(self.choose_output_dir)
        self.settings_button.setCursor(Qt.CursorShape.PointingHandCursor)

        # 顶部一行：标题 + 使用统计按钮（在导航右侧）
        top_row = QHBoxLayout()
        top_row.setContentsMargins(0, 0, 0, 0)
        top_row.setSpacing(10)
        top_row.addWidget(title_label, 1)

        self.stats_button = QPushButton("使用统计")
        self.stats_button.setObjectName("statsButton")
        self.stats_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.stats_button.clicked.connect(self.open_usage_stats)
        top_row.addWidget(self.stats_button, 0)

        left_layout.addLayout(top_row)
        left_layout.addWidget(self.nav_tree)
        left_layout.addWidget(self.settings_button)

        # 右侧内容区域使用滚动区域
        self.content_stack = QStackedWidget()
        self.content_stack.setObjectName("contentStack")
        
        # 创建滚动区域
        scroll_area = QScrollArea()
        scroll_area.setObjectName("scrollArea")
        scroll_area.setWidget(self.content_stack)
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        # 右侧容器
        right_container = QWidget()
        right_container.setObjectName("rightContainer")
        right_layout = QHBoxLayout()
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_container.setLayout(right_layout)
        right_layout.addWidget(scroll_area)

        # 添加到分割器
        splitter.addWidget(left_panel)
        splitter.addWidget(right_container)
        
        # 设置初始比例（左侧350px，右侧占剩余空间）
        splitter.setSizes([350, 850])
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)

        self.module_map = {}
        self.output_dir = None
        self._build_navigation()
        self._build_content()
        self.nav_tree.itemClicked.connect(self.on_nav_item_clicked)
        self._select_default()

        # 二级页面：使用统计
        self.stats_page = UsageStatsPage(on_back=self.back_to_main)

        # 加入根 stack
        self.root_stack.addWidget(splitter)      # index 0
        self.root_stack.addWidget(self.stats_page)  # index 1
        self.root_stack.setCurrentIndex(0)

    def open_usage_stats(self) -> None:
        try:
            self.stats_page.refresh()
        except Exception:
            pass
        self.root_stack.setCurrentIndex(1)

    def back_to_main(self) -> None:
        self.root_stack.setCurrentIndex(0)

    def _apply_apple_style(self) -> None:
        """应用（粉蓝马卡龙）风格的样式表。"""
        style = """
            QMainWindow {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 rgba(255, 245, 250, 1),
                    stop:1 rgba(238, 249, 255, 1));
            }
            
            QSplitter::handle {
                background-color: rgba(231, 183, 198, 0.85);
            }
            
            QSplitter::handle:hover {
                background-color: rgba(133, 180, 255, 0.95);
            }
            
            QSplitter::handle:pressed {
                background-color: rgba(90, 150, 235, 1);
            }
            
            #rightContainer {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 rgba(225, 244, 255, 0.9),
                    stop:1 rgba(210, 235, 255, 0.9));
            }
            
            #leftPanel {
                background: rgba(255, 255, 255, 0.75);
                border-right: 2px solid rgba(231, 183, 198, 0.85);
                min-width: 280px;
                max-width: 560px;
                border-top-left-radius: 18px;
                border-bottom-left-radius: 18px;
            }
            
            #titleLabel {
                color: #6d4c4c;
                padding: 10px 0px;
                font-weight: 800;
            }
            
            #navTree {
                background-color: transparent;
                border: none;
                outline: none;
                font-size: 14px;
                color: #6d4c4c;
                font-family: "Microsoft YaHei UI", "Segoe UI", sans-serif;
            }
            
            #navTree::item {
                padding: 8px 12px;
                border-radius: 8px;
                margin: 2px 0px;
            }
            
            #navTree::item:hover {
                background-color: rgba(231, 183, 198, 0.18);
            }
            
            #navTree::item:selected {
                background: rgba(133, 180, 255, 0.28);
                border: 1px solid rgba(133, 180, 255, 0.55);
                color: #2f5f86;
            }
            
            #navTree::item:selected:hover {
                background: rgba(133, 180, 255, 0.36);
            }
            
            #navTree::branch:has-children:!has-siblings:closed,
            #navTree::branch:closed:has-children:has-siblings {
                border: none;
                background: transparent;
            }
            
            #navTree::branch:open:has-children:!has-siblings,
            #navTree::branch:open:has-children:has-siblings {
                border: none;
                background: transparent;
            }
            
            #settingsButton {
                background: rgba(255, 248, 251, 0.95);
                color: #6d4c4c;
                border: 2px solid rgba(231, 183, 198, 0.85);
                border-radius: 16px;
                padding: 12px 18px;
                font-size: 14px;
                font-weight: 800;
                font-family: "Microsoft YaHei UI", "Segoe UI", sans-serif;
            }
            
            #settingsButton:hover {
                background: rgba(255, 235, 244, 1);
            }
            
            #settingsButton:pressed {
                background: rgba(231, 183, 198, 0.35);
            }

            #statsButton {
                padding: 8px 12px;
                border-radius: 14px;
                border: 2px solid rgba(133, 180, 255, 0.75);
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 rgba(227, 246, 255, 0.95),
                    stop:1 rgba(246, 232, 255, 0.85));
                font-weight: 800;
                color: #2f5f86;
            }
            #statsButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 rgba(214, 242, 255, 1),
                    stop:1 rgba(240, 225, 255, 1));
            }

            /* 通用任务页按钮（folder_task_page / file_task_page 已设置 objectName） */
            #taskPrimaryButton {
                padding: 12px 16px;
                border-radius: 18px;
                border: 2px solid rgba(133, 180, 255, 0.9);
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 rgba(214, 242, 255, 1),
                    stop:1 rgba(232, 220, 255, 0.95));
                font-weight: 900;
                color: #2f5f86;
            }
            #taskPrimaryButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 rgba(198, 236, 255, 1),
                    stop:1 rgba(223, 206, 255, 1));
            }
            #taskSecondaryButton {
                padding: 10px 14px;
                border-radius: 18px;
                border: 2px solid rgba(231, 183, 198, 0.9);
                background: rgba(255, 248, 251, 0.95);
                font-weight: 800;
                color: #6d4c4c;
            }
            #taskSecondaryButton:hover {
                background: rgba(255, 235, 244, 1);
            }
            
            #scrollArea {
                background-color: transparent;
                border: none;
            }
            
            #contentStack {
                background: rgba(255, 255, 255, 0.82);
                border-radius: 18px;
                margin: 18px;
                border: 2px solid rgba(231, 183, 198, 0.65);
            }
            
            QScrollBar:horizontal {
                border: none;
                background-color: #f5f5f7;
                height: 12px;
                margin: 0px;
                border-radius: 6px;
            }
            
            QScrollBar::handle:horizontal {
                background-color: #d2d2d7;
                border-radius: 6px;
                min-width: 30px;
            }
            
            QScrollBar::handle:horizontal:hover {
                background-color: #86868b;
            }
            
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                border: none;
                background: none;
                width: 0px;
            }
            
            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
                background: none;
            }
            
            QScrollBar:vertical {
                border: none;
                background-color: rgba(255, 255, 255, 0.35);
                width: 12px;
                margin: 0px;
                border-radius: 6px;
            }
            
            QScrollBar::handle:vertical {
                background-color: rgba(133, 180, 255, 0.65);
                border-radius: 6px;
                min-height: 30px;
            }
            
            QScrollBar::handle:vertical:hover {
                background-color: rgba(133, 180, 255, 0.85);
            }
            
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                border: none;
                background: none;
                height: 0px;
            }
            
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: none;
            }
            
            QLabel {
                font-family: "Microsoft YaHei UI", "Segoe UI", sans-serif;
            }
            
            QMessageBox {
                background-color: #ffffff;
                font-family: "Microsoft YaHei UI", "Segoe UI", sans-serif;
            }
            
            QMessageBox QPushButton {
                background: rgba(133, 180, 255, 0.85);
                color: #2f5f86;
                border: 1px solid rgba(133, 180, 255, 0.9);
                border-radius: 12px;
                padding: 8px 20px;
                font-size: 13px;
                min-width: 80px;
            }
        """
        self.setStyleSheet(style)

    def _build_navigation(self) -> None:
        """构建导航树。"""
        QTreeWidgetItem(self.nav_tree, ["当当切出拉销量"])

        QTreeWidgetItem(self.nav_tree, ["已推单未入库表处理"])

        QTreeWidgetItem(self.nav_tree, ["发货时效表处理"])

        stock_item = QTreeWidgetItem(self.nav_tree, ["📁 库存处理工具"])
        stock_item.setFont(0, QFont("Microsoft YaHei UI", 13, QFont.Weight.Bold))
        QTreeWidgetItem(stock_item, ["未配货"])
        QTreeWidgetItem(stock_item, ["未发货"])
        QTreeWidgetItem(stock_item, ["库存报表"])

        return_item = QTreeWidgetItem(self.nav_tree, ["📁 退货入库时效"])
        return_item.setFont(0, QFont("Microsoft YaHei UI", 13, QFont.Weight.Bold))
        QTreeWidgetItem(return_item, ["表的初步处理"])
        QTreeWidgetItem(return_item, ["表的时效计算"])

        QTreeWidgetItem(self.nav_tree, ["盘点的分析处理"])

        inventory_item = QTreeWidgetItem(self.nav_tree, ["📁 盘点的初步处理"])
        inventory_item.setFont(0, QFont("Microsoft YaHei UI", 13, QFont.Weight.Bold))
        QTreeWidgetItem(inventory_item, ["管易基础表处理工具1"])
        QTreeWidgetItem(inventory_item, ["百合基础表处理工具2"])
        QTreeWidgetItem(inventory_item, ["仓库实盘表处理工具3"])

        b2b_item = QTreeWidgetItem(self.nav_tree, ["📁 猴面包树B2B发货"])
        b2b_item.setFont(0, QFont("Microsoft YaHei UI", 13, QFont.Weight.Bold))
        QTreeWidgetItem(b2b_item, ["送货单、备货单生成"])
        QTreeWidgetItem(b2b_item, ["送货单与模板匹配（需要备货单的规格箱数）"])
        QTreeWidgetItem(b2b_item, ["提货单生成（记得填sdo和箱数）"])
        QTreeWidgetItem(b2b_item, ["箱唛转换"])

        QTreeWidgetItem(self.nav_tree, ["对账（预留）"])

        self.nav_tree.expandAll()

    def _build_content(self) -> None:
        """构建右侧内容占位。"""
        modules = [
            "当当切出拉销量",
            "发货时效表处理",
            "已推单未入库表处理",
            "库存处理工具 - 未配货",
            "库存处理工具 - 未发货",
            "库存处理工具 - 库存报表",
            "退货入库时效 - 表的初步处理",
            "退货入库时效 - 表的时效计算",
            "盘点的分析处理",
            "盘点的初步处理 - 管易基础表处理工具1",
            "盘点的初步处理 - 百合基础表处理工具2",
            "盘点的初步处理 - 仓库实盘表处理工具3",
            "猴面包树B2B发货 - 送货单、备货单生成",
            "猴面包树B2B发货 - 送货单与模板匹配（需要备货单的规格箱数）",
            "猴面包树B2B发货 - 提货单生成（记得填sdo和箱数）",
            "猴面包树B2B发货 - 箱唛转换",
            "对账（预留）",
        ]

        handlers = {
            "当当切出拉销量": lambda: build_folder_task_page(
                title="当当拉销量",
                handler=dangdang_sales_process,
                get_output_dir=lambda: self.output_dir,
                hint="请选择包含待处理Excel的文件夹；程序会在输出目录生成处理后的文件与“存档”备份。",
                manual_text=load_manual_text("dangdang"),
                stats_key="dangdang",
            ),
            "发货时效表处理": lambda: build_file_task_page(
                title="发货时效表处理",
                handler=shipping_timeliness_process,
                get_output_dir=lambda: self.output_dir,
                hint="请选择“发货商品详情X月.csv”（必须符合命名规范）；输出写入统一输出目录。",
                manual_text=load_manual_text("fahuoshixiao"),
                file_filter="CSV Files (*.csv)",
                stats_key="fahuoshixiao",
            ),
            "已推单未入库表处理": lambda: build_folder_task_page(
                title="已推单未入库表处理",
                handler=pushed_order_process,
                get_output_dir=lambda: self.output_dir,
                hint="请选择包含“退货商品明细汇总*.csv”和“店铺匹配仓库配置.xlsx”的文件夹；输出写入输出目录。",
                manual_text=load_manual_text("yituidanweiruku"),
                stats_key="yituidanweiruku",
            ),
            "库存处理工具 - 未配货": lambda: StockToolEmbeddedPage(
                tab_index=2,
                title="库存处理工具 - 未配货",
            ),
            "库存处理工具 - 未发货": lambda: StockToolEmbeddedPage(
                tab_index=1,
                title="库存处理工具 - 未发货",
            ),
            "库存处理工具 - 库存报表": lambda: StockToolEmbeddedPage(
                tab_index=0,
                title="库存处理工具 - 库存报表",
            ),
            "退货入库时效 - 表的初步处理": lambda: build_folder_task_page(
                title="退货入库时效 - 表的初步处理",
                handler=return_step1_process,
                get_output_dir=lambda: self.output_dir,
                hint="请选择包含“退货商品明细汇总*.csv”“店铺匹配仓库配置.xlsx”“12345.xlsx”的文件夹；输出写入输出目录。",
                manual_text=load_manual_text("tuihuorukushixiao_chubu"),
                stats_key="tuihuo_chubu",
            ),
            "退货入库时效 - 表的时效计算": lambda: build_folder_task_page(
                title="退货入库时效 - 表的时效计算",
                handler=return_step2_process,
                get_output_dir=lambda: self.output_dir,
                hint="请选择包含“退货入库时效分析*月总表.xlsx”的文件夹；输出写入输出目录。",
                manual_text=load_manual_text("tuihuorukushixiao_jisuan"),
                stats_key="tuihuo_jisuan",
            ),
            "盘点的分析处理": lambda: build_folder_task_page(
                title="盘点的分析处理",
                handler=inventory_analysis_process,
                get_output_dir=lambda: self.output_dir,
                hint="请选择包含盘点分析所需全部Excel的文件夹；程序会在输出目录生成“-修改后.xlsx”。",
                manual_text=load_manual_text("pandianfenxi"),
                stats_key="pandianfenxi",
            ),
            "盘点的初步处理 - 管易基础表处理工具1": lambda: build_folder_task_page(
                title="盘点的初步处理 - 管易基础表处理工具1",
                handler=guanyi_process,
                get_output_dir=lambda: self.output_dir,
                hint="请选择包含“商品库存导出*.csv”（以及可选映射表）的文件夹；输出写入输出目录。",
                manual_text=load_manual_text("pandianchubu_guanyi"),
                stats_key="pandian_guanyi",
            ),
            "盘点的初步处理 - 百合基础表处理工具2": lambda: build_folder_task_page(
                title="盘点的初步处理 - 百合基础表处理工具2",
                handler=baihe_process,
                get_output_dir=lambda: self.output_dir,
                hint="请选择包含“库存快照明细*.xlsx”“盘点-规格代码匹配.xlsx”“商品库存导出*.xlsx”的文件夹；输出写入输出目录。",
                manual_text=load_manual_text("pandianchubu_baihe"),
                stats_key="pandian_baihe",
            ),
            "盘点的初步处理 - 仓库实盘表处理工具3": lambda: build_folder_task_page(
                title="盘点的初步处理 - 仓库实盘表处理工具3",
                handler=warehouse_process,
                get_output_dir=lambda: self.output_dir,
                hint="请选择包含“信选/清元*.xlsx”、以及“盘点-规格代码匹配.xlsx”“商品库存导出*.xlsx”的文件夹；输出写入输出目录。",
                manual_text=load_manual_text("pandianchubu_cangku"),
                stats_key="pandian_cangku",
            ),
            "猴面包树B2B发货 - 送货单、备货单生成": lambda: build_b2b_page(
                "送货单、备货单生成", delivery_and_stock_process, lambda: self.output_dir
            ),
            "猴面包树B2B发货 - 送货单与模板匹配（需要备货单的规格箱数）": lambda: build_b2b_page(
                "送货单与模板匹配（需要备货单的规格箱数）", template_match_process, lambda: self.output_dir
            ),
            "猴面包树B2B发货 - 提货单生成（记得填sdo和箱数）": lambda: build_b2b_page(
                "提货单生成（记得填sdo和箱数）", picking_slip_process, lambda: self.output_dir
            ),
            "猴面包树B2B发货 - 箱唛转换": lambda: build_b2b_page(
                "箱唛转换", carton_label_process, lambda: self.output_dir
            ),
        }

        for index, module_name in enumerate(modules):
            if module_name in handlers:
                content_widget = handlers[module_name]()
            else:
                content_widget = self._create_placeholder_widget(module_name)
            self.content_stack.addWidget(content_widget)
            self.module_map[module_name] = index

    def _select_default(self) -> None:
        """默认选中首个功能，确保有内容显示。"""
        first_item = self.nav_tree.topLevelItem(0)
        if first_item:
            self.nav_tree.setCurrentItem(first_item)
            self.on_nav_item_clicked(first_item, 0)

    def on_nav_item_clicked(self, item, column) -> None:
        """导航项点击事件，仅切换占位内容。"""
        path_parts = []
        current = item
        while current:
            # 清理文本：移除文件夹图标
            text = current.text(0).replace("📁 ", "")
            path_parts.insert(0, text)
            current = current.parent()

        path = " - ".join(path_parts)
        if path in self.module_map:
            self.content_stack.setCurrentIndex(self.module_map[path])

    def choose_output_dir(self) -> None:
        """弹出文件夹选择，配置输出目录。"""
        try:
            directory = get_existing_directory(self, "选择输出文件夹", start_dir="")
            if directory:
                self.output_dir = str(directory)
                msg_box = QMessageBox(self)
                msg_box.setWindowTitle("设置成功")
                msg_box.setText("输出文件夹已设置")
                msg_box.setInformativeText(str(directory))
                msg_box.setIcon(QMessageBox.Icon.Information)
                msg_box.exec()
        except Exception:
            msg_box = QMessageBox(self)
            msg_box.setWindowTitle("设置失败")
            msg_box.setText("选择输出文件夹时出现问题")
            msg_box.setInformativeText("请重试或检查文件夹权限")
            msg_box.setIcon(QMessageBox.Icon.Critical)
            msg_box.exec()

    def _create_placeholder_widget(self, title: str) -> QWidget:
        """未接入业务的占位界面。"""
        content_widget = QWidget()
        content_widget.setMinimumWidth(600)
        content_layout = QVBoxLayout()
        content_layout.setContentsMargins(40, 40, 40, 40)
        content_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        content_widget.setLayout(content_layout)

        title_label = QLabel(title)
        title_label.setFont(QFont("Microsoft YaHei UI", 24, QFont.Weight.Bold))
        title_label.setStyleSheet("color: #1d1d1f; margin-bottom: 20px;")
        title_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
        title_label.setWordWrap(True)

        desc_label = QLabel("功能模块占位区域\n具体功能将在后续开发中实现")
        desc_label.setFont(QFont("Microsoft YaHei UI", 14))
        desc_label.setStyleSheet("color: #86868b; line-height: 1.6;")
        desc_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
        desc_label.setWordWrap(True)

        content_layout.addWidget(title_label)
        content_layout.addWidget(desc_label)
        content_layout.addStretch()
        return content_widget

