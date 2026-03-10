Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# English entrypoint to avoid cmd/powershell path encoding issues with Chinese filenames.
# Outputs:
# - dist\Excel处理工具\...    (PyInstaller output)
# - dist_installer\Install_Excel处理工具.exe (installer)

# Prefer UTF-8 output to reduce mojibake in console/logs
try {
  [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
  $OutputEncoding = [System.Text.Encoding]::UTF8
} catch {}

function Pause-Here {
  param([string]$Message = "Press Enter to exit...")
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
    throw "Step failed: $Title (exit code=$LASTEXITCODE)"
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

  # fallback: search in PATH
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

  Write-Host "[WARN] Inno Setup not found. Trying to install it automatically..." -ForegroundColor Yellow

  # Download to TEMP so we don't leave large binaries in the repo
  $dlDir = Join-Path $env:TEMP "exceltools_installer_cache"
  if (!(Test-Path $dlDir)) { New-Item -ItemType Directory -Path $dlDir | Out-Null }

  # Official downloads (may change over time). If blocked, we'll fall back to portable zip later.
  $url = "https://jrsoftware.org/download.php/is.exe"
  $out = Join-Path $dlDir "innosetup-installer.exe"

  try {
    Invoke-WebRequest -Uri $url -OutFile $out -UseBasicParsing
  } catch {
    throw "Failed to download Inno Setup installer. Please install Inno Setup 6 manually, or use the portable zip fallback."
  }

  Write-Host "[INFO] Running Inno Setup installer (silent)..." -ForegroundColor Cyan
  # This may require admin privileges because it installs to Program Files.
  $InnoArgs = @("/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART")
  $p = Start-Process -FilePath $out -ArgumentList $InnoArgs -PassThru -Wait
  if ($p.ExitCode -ne 0) {
    throw "Inno Setup installer failed (exit code=$($p.ExitCode)). Please install it manually."
  }

  $iscc = Find-ISCC -Root $Root
  if (!$iscc) {
    throw "Inno Setup installed but ISCC.exe still not found. Please install Inno Setup 6 manually."
  }
  return $iscc
}

function Make-PortableZip {
  param([string]$Root, [string]$AppName)

  $distDir = Join-Path $Root "dist\$AppName"
  if (!(Test-Path $distDir)) {
    throw "Portable zip failed: dist folder not found: $distDir"
  }

  $outDir = Join-Path $Root "dist_installer"
  if (!(Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }

  $zipPath = Join-Path $outDir ("Portable_{0}.zip" -f $AppName)
  if (Test-Path $zipPath) { Remove-Item -Force $zipPath }

  Add-Type -AssemblyName System.IO.Compression.FileSystem
  [System.IO.Compression.ZipFile]::CreateFromDirectory($distDir, $zipPath)
  Write-Host ("[OK] Portable zip created: " + $zipPath)
  return $zipPath
}

try {
  $Root = Resolve-Path (Join-Path $PSScriptRoot "..\..") | Select-Object -ExpandProperty Path
  Set-Location $Root

  # Always write logs here (do not depend on dist_installer existing yet)
  $LogDir = Join-Path $Root "resources\installer_tools\logs"
  if (!(Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir | Out-Null }
  $LogPath = Join-Path $LogDir ("build_log_{0}.txt" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
  try { Start-Transcript -Path $LogPath -Append | Out-Null } catch {}

  Write-Host "==== Build Installer (ExcelTools) ===="
  Write-Host ("Project root: " + $Root)
  Write-Host ("Log file: " + $LogPath)

  # 0) Python venv (build Python)
  $Py = Join-Path $Root ".venv\Scripts\python.exe"
  if (!(Test-Path $Py)) {
    throw "venv python not found: $Py . Please create/fix .venv first."
  }

  # If current venv is Python 3.11+, try to auto-switch to a Python 3.10 build venv (for PyArmor support).
  # IMPORTANT: do not call "py -3.10" unless it exists (otherwise it prints a scary error).
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
          Write-Host "[INFO] Detected Python 3.10. Creating build venv: .venv_build310" -ForegroundColor Cyan
          & py -3.10 -m venv $BuildVenv
        }
        if (Test-Path $BuildPy) { $Py = $BuildPy }
      }
    }
  }

  $VenvScripts = Split-Path $Py -Parent

  # Force reinstall pinned build deps to avoid PyArmor 9.x (which removed "pack")
  Run-Step "[1/5] Install build deps (pyinstaller/pyarmor)..." {
    & $Py -m pip uninstall -y pyarmor pyarmor.cli.core 2>$null
    & $Py -m pip install --upgrade --force-reinstall -r "resources\installer_tools\requirements-build.txt"
  }

  # PyArmor legacy "pack" does NOT support Python 3.11+ (it errors on 3.11/3.12).
  $PyMajor = [int](& $Py -c "import sys; print(sys.version_info.major)")
  $PyMinor = [int](& $Py -c "import sys; print(sys.version_info.minor)")
  $PyVer = ("{0}.{1}" -f $PyMajor, $PyMinor)

  # Resolve PyInstaller entrypoint
  $PyInstallerExe = Join-Path $VenvScripts "pyinstaller.exe"
  if (!(Test-Path $PyInstallerExe)) { $PyInstallerExe = "pyinstaller" }

  Run-Step "[2/5] Generate app_icon.ico..." { & $Py "resources\installer_tools\make_ico.py" }

  # 3) Build exe
  Write-Host "[3/5] Build exe..."
  # Use ASCII build name to avoid mojibake in filesystem paths under PS 5.1/cmd
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

  # Code hardening (key=072554) via PyArmor
  # NOTE: legacy PyArmor pack doesn't support Python 3.11+; auto-disable to keep one-shot build working.
  $USE_PYARMOR = 1
  $PyArmorAttempted = $false
  if ($PyMinor -ge 11) {
    $USE_PYARMOR = 0
    Write-Host ("[WARN] Current Python is {0}, PyArmor hardening is disabled (PyArmor pack doesn't support Python 3.11+)." -f $PyVer) -ForegroundColor Yellow
    Write-Host "[WARN] To enable hardening (key=072554), please build with Python 3.10.x." -ForegroundColor Yellow
  }
  if ($USE_PYARMOR -eq 1) {
    $PyArmorAttempted = $true
    # Resolve pyarmor entry AFTER installation (venv may not be activated so PATH might not include it)
    # PyArmor 8.x uses a new CLI (gen/reg/cfg). The legacy CLI (with "pack") is provided as "pyarmor-7".
    $PyArmor7Exe = Join-Path $VenvScripts "pyarmor-7.exe"
    if (!(Test-Path $PyArmor7Exe)) {
      throw "pyarmor-7.exe not found in venv after install: $PyArmor7Exe (PyArmor legacy CLI is required for 'pack')"
    }

    Write-Host "  - [Harden] PyArmor (key=072554)..."
    $E = @(
      "--noconfirm",
      "--clean",
      "--name `"$AppBuildName`"",
      "--noconsole",
      "--icon `"$IconIco`""
    )
    foreach ($d in $AddData) { $E += "--add-data `"$d`"" }
    $E += "main.py"
    # PyArmor 8.x supports "pack" (PyArmor 9.x removed it).
    # Try PyArmor pack, otherwise fall back to plain PyInstaller.
    try {
      Run-Step "  - PyArmor pack (pyarmor-7)..." { & $PyArmor7Exe pack --clean --output "dist" -e ($E -join " ") }
      # Sanity check: ensure output exists; otherwise treat as failure and fall back.
      $Exe1 = Join-Path $Root ("dist\{0}\{0}.exe" -f $AppBuildName)
      $Exe2 = Join-Path $Root ("dist\{0}.exe" -f $AppBuildName)
      if (!(Test-Path $Exe1) -and !(Test-Path $Exe2)) {
        throw "PyArmor pack finished but output exe not found (checked: $Exe1 , $Exe2)"
      }
    } catch {
      Write-Host "[WARN] PyArmor pack failed. Falling back to PyInstaller without hardening." -ForegroundColor Yellow
      Write-Host ("[WARN] " + $_.Exception.Message) -ForegroundColor Yellow
      $USE_PYARMOR = 0
    }
  } else {
    Write-Host "  - PyInstaller (no hardening)..."
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

  # If hardening failed and we fell back, build with PyInstaller now.
  if ($PyArmorAttempted -and ($USE_PYARMOR -eq 0)) {
    Write-Host "[3b] Build exe with PyInstaller (fallback)..."
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

  # 4) Build installer
  Write-Host "[4/5] Build installer (Inno Setup)..."
  $InstallerBuilt = $false
  $PortableZipPath = $null
  try {
    $Iscc = Ensure-InnoSetup -Root $Root
    Run-Step "  - ISCC compile installer.iss..." { & $Iscc "resources\installer_tools\installer.iss" }
    $InstallerBuilt = $true
  } catch {
    Write-Host "[WARN] Unable to build Inno Setup installer. Falling back to portable zip." -ForegroundColor Yellow
    Write-Host ("[WARN] " + $_.Exception.Message) -ForegroundColor Yellow
    $PortableZipPath = Make-PortableZip -Root $Root -AppName $AppBuildName
  }

  Write-Host "[5/5] Done!"
  if ($InstallerBuilt) {
    Write-Host "Installer: dist_installer\Install_ExcelTools.exe"
  } else {
    Write-Host ("Portable: " + $PortableZipPath)
  }

  try { Stop-Transcript | Out-Null } catch {}
  Pause-Here "Done. Press Enter to exit..."
} catch {
  try { Stop-Transcript | Out-Null } catch {}
  Write-Host ""
  Write-Host "===== FAILED =====" -ForegroundColor Red
  Write-Host $_.Exception.Message -ForegroundColor Red
  Write-Host "Send me the latest log under resources\\installer_tools\\logs\\build_log_*.txt" -ForegroundColor Yellow
  Pause-Here
  exit 1
}


