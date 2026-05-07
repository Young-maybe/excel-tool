<p align="center">
  <img src="https://trae-api-cn.mchost.guru/api/ide/v1/text_to_image?prompt=A%20modern%20clean%20spreadsheet%20application%20icon%20with%20Excel%20grid%20pattern%2C%20soft%20pastel%20blue%20and%20pink%20gradient%2C%20rounded%20corners%2C%20minimalist%20design%2C%20white%20background%2C%20app%20icon%20style&image_size=square" width="120" alt="Excel Tool Logo"/>
</p>

<h1 align="center">Excel 处理工具</h1>

<p align="center">
  <strong>一款面向电商仓储场景的桌面级 Excel / CSV 批处理工具</strong>
</p>

<p align="center">
  <a href="#功能特性">功能特性</a> •
  <a href="#模块一览">模块一览</a> •
  <a href="#快速开始">快速开始</a> •
  <a href="#技术架构">技术架构</a> •
  <a href="#打包发布">打包发布</a>
</p>

---

## 功能特性

| 特性 | 说明 |
|:--|:--|
| 🎨 **马卡龙风格 UI** | 精心设计的渐变配色与圆角界面，操作体验舒适流畅 |
| ⚡ **后台线程执行** | 所有耗时任务在 QThread 后台运行，UI 永不卡顿 |
| 📊 **实时进度面板** | 百分比进度条 + 阶段状态文本 + 详细日志滚动 |
| 📈 **使用统计** | 自动记录每次任务节省时长，生成树形报表 |
| 📖 **外置使用手册** | 手册文本存于 `huancuntxt/`，无需改代码即可维护 |
| 🔗 **飞书集成** | 库存处理工具支持连接飞书多维表格 |
| 📦 **一键打包** | PyInstaller + Inno Setup，自动生成安装包 / 便携版 |
| 🔐 **代码加固** | 可选 PyArmor 混淆，提高反编译门槛 |

---

## 模块一览

### 当当切出拉销量
从当当供应链订单明细表中提取销量数据，自动生成处理后文件与 `.xlsx` 存档备份。

### 已推单未入库表处理
整合退货商品明细 CSV 与店铺仓库配置，输出包含未审核/仓库回复附加 sheet 的分析表。

### 发货时效表处理
读取 `发货商品详情X月.csv`，计算 24h / 48h 时效分布，输出三 sheet 分析文件。

### 库存处理工具（内嵌版）
保留原有 UI 与交互逻辑的三大子功能：
- **未配货** — 未配货商品分析与导出
- **未发货** — 未发货订单追踪
- **库存报表** — 库存数据汇总报表（支持飞书多维表格）

### 退货入库时效
两步式处理流程：
1. **表的初步处理** — 整合多源数据生成月度总表
2. **表的时效计算** — 基于总表推算各单据入库时效

### 盘点的初步处理
三个工具依次处理不同来源的基础数据：
1. **管易基础表处理工具1** — 商品库存导出 CSV 标准化
2. **百合基础表处理工具2** — 库存快照明细匹配规格代码
3. **仓库实盘表处理工具3** — 信选/清元实盘表清洗转换

### 盘点的分析处理
基于盘点基础数据完成差异分析与报表生成。

### 猴面包树 B2B 发货
四步式 B2B 发货单据生成流程：
1. **送货单、备货单生成** — 从易快报导单提取并生成送货/备货单
2. **送货单与模板匹配** — 回写模板规格箱数信息
3. **提货单生成** — 填入 SDO 和箱数后批量出提货单
4. **箱唛转换** — B2B 箱单格式标准化转换

---

## 快速开始

### 环境要求

- Windows 10 / 11
- Python 3.10+ （推荐 3.11）

### 安装依赖

```bash
pip install -r requirements.txt
```

> 主要依赖：`PyQt6`、`openpyxl`、`pandas`、`xlsxwriter`、`pywin32`、`requests`、`lark-oapi`

### 启动应用

```bash
python main.py
```

### 下载安装包

直接下载 [dist_installer](./dist_installer) 目录下的安装程序到桌面运行即可，无需 Python 环境。

---

## 技术架构

```
excel-tool/
├── main.py                          # 程序入口 / 日志 / 全局异常捕获
├── requirements.txt                 # 运行时依赖
├── ui/
│   ├── main_window.py               # 主窗口（左侧导航树 + 右侧内容区）
│   └── components/
│       ├── folder_task_page.py      # 文件夹输入通用任务页
│       ├── file_task_page.py        # 单文件输入通用任务页
│       ├── apple_progress_panel.py  # Apple 风格进度面板
│       ├── task_runner.py           # QThread 后台执行框架
│       ├── dialogs.py              # 马卡龙主题文件夹选择弹窗
│       ├── manual_text_repo.py     # 外置手册文本加载器
│       ├── usage_stats_page.py     # 使用统计二级页面
│       └── stock_tool_legacy_main.py # 库存处理工具（内嵌原 UI）
├── modules/                         # 业务逻辑层（UI 解耦）
│   ├── dangdang_sales/             # 当当拉销量
│   ├── pushed_order_not_inbound/   # 已推单未入库
│   ├── shipping_timeliness/         # 发货时效
│   ├── return_inbound_timeliness/   # 退货入库时效
│   ├── inventory_preprocess/        # 盘点初步处理
│   ├── inventory_analysis/          # 盘点分析处理
│   └── b2b_shipping/               # B2B 发货
├── services/
│   ├── excel_service.py            # Excel 服务占位
│   ├── file_service.py             # 文件服务
│   └── usage_stats.py              # 使用统计持久化
├── huancuntxt/                      # 外置使用手册（每模块一个 .txt）
├── resources/                       # 图标 / 日志 / 统计数据
│   └── installer_tools/            # 打包脚本集合
└── dist_installer/                  # 最终交付产物（安装包 / 便携版）
```

### 设计原则

- **UI 与业务严格分离** — `ui/` 只管界面，`modules/` 只管逻辑
- **统一 I/O 策略** — 输入从用户选择获取，输出写入统一的"输出目录"
- **安全复制机制** — 先复制到工作目录再处理，避免 WinError 32 文件占用冲突
- **同名自动改名** — 输出目录存在同名文件时自动追加 `_处理副本N`

---

## 打包发布

### 一键生成安装包

```bash
# 方式一：中文入口（内部转调英文脚本）
resources/installer_tools/一键生成安装包.bat

# 方式二：英文入口
resources/installer_tools/build_installer.bat
```

### 产物说明

| 产物 | 路径 | 说明 |
|:--|:--|:--|
| 安装向导 | `dist_installer/Install_ExcelTools.exe` | 支持自定义路径 + 桌面快捷方式 |
| 便携版 | `dist_installer/Portable_ExcelTools.zip` | 解压即用（兜底方案） |

### 清理中间产物

```bash
resources/installer_tools/清理打包中间产物.bat
```

---

## 项目信息

- **开发语言**: Python 3.11
- **GUI 框架**: PyQt6
- **目标平台**: Windows 10 / 11
- **许可协议**: Private

---

<p align="center">
  Made with ❤️ by <a href="https://github.com/Young-maybe">Young-maybe</a>
</p>
