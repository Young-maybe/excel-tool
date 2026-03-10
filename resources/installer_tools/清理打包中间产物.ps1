Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

try {
  [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
  $OutputEncoding = [System.Text.Encoding]::UTF8
} catch {}

$Root = Resolve-Path (Join-Path $PSScriptRoot "..\..") | Select-Object -ExpandProperty Path
Set-Location $Root

Write-Host "==== 清理打包中间产物 ===="
Write-Host ("项目根目录: " + $Root)

function Safe-Remove([string]$Path) {
  if (Test-Path $Path) {
    try {
      Remove-Item -Recurse -Force $Path
      Write-Host ("[DEL] " + $Path)
    } catch {
      Write-Host ("[WARN] 删除失败: " + $Path + " -> " + $_.Exception.Message) -ForegroundColor Yellow
    }
  } else {
    Write-Host ("[SKIP] " + $Path)
  }
}

# 1) PyInstaller 中间产物
Safe-Remove ".\build"
Safe-Remove ".\dist"

Get-ChildItem -LiteralPath . -Filter "*.spec" -ErrorAction SilentlyContinue | ForEach-Object {
  try { Remove-Item -Force $_.FullName; Write-Host ("[DEL] " + $_.Name) } catch {}
}

# 2) installer_tools 日志/缓存
Get-ChildItem -LiteralPath ".\resources\installer_tools\logs" -Filter "build_log_*.txt" -ErrorAction SilentlyContinue | ForEach-Object {
  try { Remove-Item -Force $_.FullName; Write-Host ("[DEL] " + $_.FullName) } catch {}
}
Safe-Remove ".\resources\installer_tools\third_party\innosetup-installer.exe"

# 3) 便携包旧乱码产物（保留 Portable_ExcelTools.zip）
Get-ChildItem -LiteralPath ".\dist_installer" -Filter "Portable_Excel*.zip" -ErrorAction SilentlyContinue | ForEach-Object {
  if ($_.Name -ne "Portable_ExcelTools.zip") {
    try { Remove-Item -Force $_.FullName; Write-Host ("[DEL] " + $_.FullName) } catch {}
  } else {
    Write-Host ("[KEEP] " + $_.FullName)
  }
}

Write-Host "[DONE] 清理完成。"
Read-Host "按回车键退出…" | Out-Null


