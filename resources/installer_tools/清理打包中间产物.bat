@echo off
setlocal
chcp 65001 >nul

REM 清理 PyInstaller/Inno Setup 过程中产生的中间产物/缓存
REM 说明：
REM - 保留 dist_installer\Install_ExcelTools.exe 与 dist_installer\Portable_ExcelTools.zip
REM - 会删除 build/ dist/ *.spec 以及乱码旧 zip、installer_tools 日志与缓存

set ROOT=%~dp0..\..
pushd "%ROOT%" >nul

echo [CLEAN] removing build/ dist/ *.spec ...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
del /q "*.spec" 2>nul

echo [CLEAN] removing old portable zip (garbled name) if any...
del /q "dist_installer\Portable_Excel*.zip" 2>nul
if exist "dist_installer\Portable_ExcelTools.zip" (
  echo [KEEP] dist_installer\Portable_ExcelTools.zip
) else (
  echo [WARN] Portable_ExcelTools.zip not found (did you build first?)
)

echo [CLEAN] removing installer_tools logs/cache...
if exist "resources\installer_tools\logs" del /q "resources\installer_tools\logs\build_log_*.txt" 2>nul
if exist "resources\installer_tools\third_party\innosetup-installer.exe" del /q "resources\installer_tools\third_party\innosetup-installer.exe" 2>nul

echo [DONE] cleanup finished.
pause

popd >nul
endlocal


