Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# 说明：为了“看得懂 + 一键跑”，所有打包相关内容都集中在 resources/installer_tools 下。
# 生成结果：
# - dist\Excel处理工具\...    （PyInstaller 输出）
# - dist_installer\Install_Excel处理工具.exe （安装包）

# 尽量使用 UTF-8 输出，减少控制台/日志乱码
try {
  [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
  $OutputEncoding = [System.Text.Encoding]::UTF8
} catch {}

function Pause-Here {
  param([string]$Message = "按回车键退出…")
  try { Read-Host $Message | Out-Null } catch {}
}

function Run-Step {
  param(
    [Parameter(Mandatory=$true)][string]$Title,
    [Parameter(Mandatory=$true)][scriptblock]$Cmd
  )
  Write-Host $Title
  & $Cmd
  if ($LASTEXITCODE -ne 0) {
    throw "步骤失败：$Title（exit code=$LASTEXITCODE）"
  }
}

function Find-ISCC {
  param([string]$Root)

  $candidates = @(
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
    "${env:ProgramFiles}\Inno Setup 6\ISCC.exe",
    (Join-Path $Root "resources\installer_tools\third_party\innosetup\ISCC.exe")
  )

  foreach ($p in $candidates) {
    if ($p -and (Test-Path $p)) { return $p }
  }

  try {
    $cmd = Get-Command ISCC.exe -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Path) { return $cmd.Path }
  } catch {}

  return $null
}

function Ensure-InnoSetup {
  param([string]$Root)

  $iscc = Find-ISCC -Root $Root
  if ($iscc) { return $iscc }

  Write-Host "[警告] 未检测到 Inno Setup 6，尝试自动安装（静默）..." -ForegroundColor Yellow

  # 下载到 TEMP，避免把大体积安装器缓存留在工程里
  $dlDir = Join-Path $env:TEMP "exceltools_installer_cache"
  if (!(Test-Path $dlDir)) { New-Item -ItemType Directory -Path $dlDir | Out-Null }

  $url = "https://jrsoftware.org/download.php/is.exe"
  $out = Join-Path $dlDir "innosetup-installer.exe"

  Invoke-WebRequest -Uri $url -OutFile $out -UseBasicParsing

  Write-Host "[信息] 正在运行 Inno Setup 安装器（静默）..." -ForegroundColor Cyan
  $InnoArgs = @("/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART")
  $p = Start-Process -FilePath $out -ArgumentList $InnoArgs -PassThru -Wait
  if ($p.ExitCode -ne 0) {
    throw "Inno Setup 安装失败（exit code=$($p.ExitCode)）。请手动安装 Inno Setup 6。"
  }

  $iscc = Find-ISCC -Root $Root
  if (!$iscc) {
    throw "已尝试安装 Inno Setup，但仍找不到 ISCC.exe。请手动安装 Inno Setup 6。"
  }
  return $iscc
}

function Make-PortableZip {
  param([string]$Root, [string]$AppName)

  $distDir = Join-Path $Root "dist\$AppName"
  if (!(Test-Path $distDir)) {
    throw "便携版压缩失败：未找到 dist 目录：$distDir"
  }

  $outDir = Join-Path $Root "dist_installer"
  if (!(Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }

  $zipPath = Join-Path $outDir ("Portable_{0}.zip" -f $AppName)
  if (Test-Path $zipPath) { Remove-Item -Force $zipPath }

  Add-Type -AssemblyName System.IO.Compression.FileSystem
  [System.IO.Compression.ZipFile]::CreateFromDirectory($distDir, $zipPath)
  Write-Host ("[OK] 已生成便携版压缩包：" + $zipPath)
  return $zipPath
}

try {
  $Root = Resolve-Path (Join-Path $PSScriptRoot "..\..") | Select-Object -ExpandProperty Path
  Set-Location $Root

  # 日志固定写到 tools 目录，避免 dist_installer 尚未创建导致“无日志可查”
  $LogDir = Join-Path $Root "resources\installer_tools\logs"
  if (!(Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir | Out-Null }
  $LogPath = Join-Path $LogDir ("build_log_{0}.txt" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
  try { Start-Transcript -Path $LogPath -Append | Out-Null } catch {}

  Write-Host "==== 一键生成安装包（ExcelTools）===="
  Write-Host ("项目根目录：" + $Root)
  Write-Host ("日志文件：" + $LogPath)

  # 0) Python venv（打包用 Python）
  $Py = Join-Path $Root ".venv\Scripts\python.exe"
  if (!(Test-Path $Py)) {
    throw "未找到虚拟环境：$Py。请先创建/修复 .venv。"
  }

  # 如果当前 venv 是 Python 3.11+，尝试自动切换到 Python 3.10 的打包环境（用于 PyArmor）
  # 注意：不要直接调用 py -3.10（没装会弹很吓人的错误）
  $CurMinor = [int](& $Py -c "import sys; print(sys.version_info.minor)")
  if ($CurMinor -ge 11) {
    $hasPyLauncher = $false
    try { $hasPyLauncher = [bool](Get-Command py -ErrorAction SilentlyContinue) } catch { $hasPyLauncher = $false }
    if ($hasPyLauncher) {
      $pyList = @()
      try {
        $old = $ErrorActionPreference
        $ErrorActionPreference = "Continue"
        $pyList = (& py -0p 2>$null)
      } finally {
        $ErrorActionPreference = $old
      }
      $has310 = $false
      foreach ($ln in $pyList) {
        if ($ln -match "3\.10") { $has310 = $true; break }
      }
      if ($has310) {
        $BuildVenv = Join-Path $Root ".venv_build310"
        $BuildPy = Join-Path $BuildVenv "Scripts\python.exe"
        if (!(Test-Path $BuildPy)) {
          Write-Host "[信息] 检测到 Python 3.10，正在创建打包 venv：.venv_build310" -ForegroundColor Cyan
          & py -3.10 -m venv $BuildVenv
        }
        if (Test-Path $BuildPy) { $Py = $BuildPy }
      }
    }
  }

  $VenvScripts = Split-Path $Py -Parent

  # 强制按 requirements-build 锁定版本重新安装，避免 PyArmor 9.x（已移除 pack）
  Run-Step "[1/5] 安装打包依赖（pyinstaller/pyarmor）..." {
    & $Py -m pip uninstall -y pyarmor pyarmor.cli.core 2>$null
    & $Py -m pip install --upgrade --force-reinstall -r "resources\installer_tools\requirements-build.txt"
  }

  # 安装后再定位 pyarmor.exe（venv 未激活时 PATH 里可能找不到）
  # 定位 PyInstaller.exe（避免 python -m PyInstaller 传参异常）
  $PyInstallerExe = Join-Path $VenvScripts "pyinstaller.exe"
  if (!(Test-Path $PyInstallerExe)) { $PyInstallerExe = "pyinstaller" }

  Run-Step "[2/5] 生成 app_icon.ico（任务栏/安装包更稳定）..." { & $Py "resources\installer_tools\make_ico.py" }

  # 3) 打包 exe
  Write-Host "[3/5] 打包 exe（PyInstaller）..."
  # 用英文作为构建名，避免 PowerShell 5.1 / cmd 的编码导致路径乱码
  $AppBuildName = "ExcelTools"
  $AppDisplayName = "Excel处理工具"
  $DistDir = Join-Path $Root "dist\$AppBuildName"
  if (Test-Path $DistDir) { Remove-Item -Recurse -Force $DistDir }

  $IconIco = Join-Path $Root "resources\app_icon.ico"
  $AddData = @(
    "resources;resources",
    "huancuntxt;huancuntxt",
    "ui\components\stock_tool_lark_grammars;ui\components\stock_tool_lark_grammars"
  )

  # 你要求的“加密/加固”（密钥 072554）：用 PyArmor 做代码加固（提高反编译成本）
  # 说明：这是“加固/混淆”，不是绝对不可逆的加密。
  $PyMajor = [int](& $Py -c "import sys; print(sys.version_info.major)")
  $PyMinor = [int](& $Py -c "import sys; print(sys.version_info.minor)")
  $PyVer = ("{0}.{1}" -f $PyMajor, $PyMinor)

  $USE_PYARMOR = 1
  $PyArmorAttempted = $false
  if ($PyMinor -ge 11) {
    $USE_PYARMOR = 0
    Write-Host ("[警告] 当前 Python 版本为 {0}，PyArmor 加固已自动关闭（pack 不支持 Python 3.11+）" -f $PyVer) -ForegroundColor Yellow
    Write-Host "[警告] 如需启用加固（072554），请使用 Python 3.10.x 的环境进行打包。" -ForegroundColor Yellow
  }
  if ($USE_PYARMOR -eq 1) {
    $PyArmorAttempted = $true
    # PyArmor 8.x 的新 CLI 只有 gen/reg/cfg；旧命令（含 pack）需要 pyarmor-7
    $PyArmor7Exe = Join-Path $VenvScripts "pyarmor-7.exe"
    if (!(Test-Path $PyArmor7Exe)) {
      throw "未找到 pyarmor-7.exe：$PyArmor7Exe（PyArmor 的旧命令 pack 需要它）"
    }
    Write-Host "  - [加固] 使用 PyArmor（密钥=072554）..."
    $E = @(
      "--noconfirm",
      "--clean",
      "--name `"$AppBuildName`"",
      "--noconsole",
      "--icon `"$IconIco`""
    )
    foreach ($d in $AddData) { $E += "--add-data `"$d`"" }
    $E += "main.py"
    # PyArmor 8.x 才支持 pack（PyArmor 9.x 已移除 pack 子命令）
    # 尝试 PyArmor pack；失败则自动降级为纯 PyInstaller（保证先能出安装包）
    $pyarmor_ok = $true
    try {
      Run-Step "  - PyArmor pack (pyarmor-7)..." { & $PyArmor7Exe pack --clean --output "dist" -e ($E -join " ") }
      $Exe1 = Join-Path $Root ("dist\{0}\{0}.exe" -f $AppBuildName)
      $Exe2 = Join-Path $Root ("dist\{0}.exe" -f $AppBuildName)
      if (!(Test-Path $Exe1) -and !(Test-Path $Exe2)) {
        throw "PyArmor pack 执行结束但未找到输出 exe（检查：$Exe1 / $Exe2）"
      }
    } catch {
      $pyarmor_ok = $false
      Write-Host "[警告] PyArmor pack 失败，将降级为纯 PyInstaller（不加固）" -ForegroundColor Yellow
      Write-Host ("[警告] " + $_.Exception.Message) -ForegroundColor Yellow
    }
    if (-not $pyarmor_ok) {
      Write-Host "[3b] 使用 PyInstaller 打包（降级方案）..."
      $PyInstallerArgs = @(
        "--noconfirm",
        "--clean",
        "--name", $AppBuildName,
        "--noconsole",
        "--icon", $IconIco
      )
      foreach ($d in $AddData) { $PyInstallerArgs += @("--add-data", $d) }
      $PyInstallerArgs += "main.py"
      Run-Step "  - PyInstaller..." { & $PyInstallerExe @PyInstallerArgs }
    }
  } else {
    Write-Host "  - [未加固] 直接 PyInstaller..."
    $PyInstallerArgs = @(
      "--noconfirm",
      "--clean",
      "--name", $AppBuildName,
      "--noconsole",
      "--icon", $IconIco
    )
    foreach ($d in $AddData) { $PyInstallerArgs += @("--add-data", $d) }
    $PyInstallerArgs += "main.py"
    Run-Step "  - PyInstaller..." { & $PyInstallerExe @PyInstallerArgs }
  }

  # 如果 PyArmor 尝试过且失败了，这里再用 PyInstaller 兜底（避免在“自动关闭 PyArmor”场景重复打包）
  if ($PyArmorAttempted -and ($USE_PYARMOR -eq 0)) {
    Write-Host "[3b] 使用 PyInstaller 打包（兜底）..."
    $PyInstallerArgs = @(
      "--noconfirm",
      "--clean",
      "--name", $AppBuildName,
      "--noconsole",
      "--icon", $IconIco
    )
    foreach ($d in $AddData) { $PyInstallerArgs += @("--add-data", $d) }
    $PyInstallerArgs += "main.py"
    Run-Step "  - PyInstaller..." { & $PyInstallerExe @PyInstallerArgs }
  }

  # 4) 生成安装包
  Write-Host "[4/5] 生成安装包（Inno Setup）..."
  $InstallerBuilt = $false
  $PortableZipPath = $null
  try {
    $Iscc = Ensure-InnoSetup -Root $Root
    Run-Step "  - ISCC 编译 installer.iss..." { & $Iscc "resources\installer_tools\installer.iss" }
    $InstallerBuilt = $true
  } catch {
    Write-Host "[警告] 无法生成安装向导，将自动降级生成便携版压缩包（Portable_*.zip）" -ForegroundColor Yellow
    Write-Host ("[警告] " + $_.Exception.Message) -ForegroundColor Yellow
    $PortableZipPath = Make-PortableZip -Root $Root -AppName $AppBuildName
  }

  # 5) 完成提示
  Write-Host "[5/5] 完成！"
  if ($InstallerBuilt) {
    Write-Host "安装包在：dist_installer\\Install_ExcelTools.exe"
    Write-Host "你可以把该 exe 压缩成 zip 发给别人安装。"
  } else {
    Write-Host ("便携版在：" + $PortableZipPath)
    Write-Host "你可以把该 zip 直接发给别人，解压即用。"
  }

  try { Stop-Transcript | Out-Null } catch {}
  Pause-Here "成功完成。按回车键退出…"
} catch {
  try { Stop-Transcript | Out-Null } catch {}
  Write-Host ""
  Write-Host "===== 失败了 =====" -ForegroundColor Red
  Write-Host $_.Exception.Message -ForegroundColor Red
  Write-Host "请把 resources\\installer_tools\\logs 下最新的 build_log_*.txt 发我，我帮你定位。" -ForegroundColor Yellow
  Pause-Here
  exit 1
}


