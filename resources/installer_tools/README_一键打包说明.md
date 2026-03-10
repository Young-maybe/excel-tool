### 一键生成安装包（给非技术同事看的版本）

你只需要做 2 件事：

1. **安装 Inno Setup 6**（只需一次）
   - 安装完成后，系统里会有：`ISCC.exe`

2. 运行（推荐英文入口，避免中文路径乱码）：
   - **推荐双击**：`resources/installer_tools/build_installer.bat`
   - 或者 PowerShell：`resources/installer_tools/build_installer.ps1`
   - 中文入口也保留：`resources/installer_tools/一键生成安装包.bat`（内部已转调英文脚本）

成功后，你会得到：
- `dist_installer/Install_ExcelTools.exe`（这就是安装包，安装界面与快捷方式仍显示“Excel处理工具”）
- 如果机器无法安装 Inno Setup（没权限/被拦截），脚本会自动降级生成：
  - `dist_installer/Portable_ExcelTools.zip`（便携版：解压即用）

注意：
- 如果你的 Inno Setup 缺少 `ChineseSimplified.isl`（会导致编译失败），当前脚本会使用 **默认英文语言包** 来确保一定能生成安装包。

### 清理打包中间产物（建议每次打包后做一次）
打包会产生比较大的中间产物（`build/`、`dist/`、`*.spec`、部分日志/缓存）。

- 推荐双击运行：`resources/installer_tools/清理打包中间产物.bat`
- 或 PowerShell：`resources/installer_tools/清理打包中间产物.ps1`

清理脚本会保留：
- `dist_installer/Install_ExcelTools.exe`
- `dist_installer/Portable_ExcelTools.zip`

如果双击后“闪退/没反应”：
- 直接用 **英文入口**：`build_installer.bat`
- 查看日志：`resources/installer_tools/logs/build_log_*.txt`

安装包特性：
- 选择安装路径（向导）
- 可选创建桌面快捷方式
- 安装后即可运行（不依赖 Python 环境）

### 代码“加密/加固”说明（密钥：072554）
本项目使用 **PyArmor** 做“代码加固/混淆”，用于提高反编译门槛（不是绝对不可逆加密）。

- 默认会启用加固（密钥：`072554`）
- 如你想关闭加固，可在脚本中把 `USE_PYARMOR` 改为 `0`

注意：
- **PyArmor 的 pack（pyarmor-7）目前不支持 Python 3.11+**。如果你的 `.venv` 是 Python 3.11/3.12，脚本会自动关闭加固并用 PyInstaller 继续打包。
- 如需启用加固，请用 **Python 3.10.x** 创建用于打包的虚拟环境再运行脚本。


