@echo off
setlocal

REM English entrypoint to avoid cmd Unicode/encoding issues
chcp 65001 >nul
set ROOT=%~dp0..\..
pushd "%ROOT%" >nul

powershell -NoProfile -ExecutionPolicy Bypass -NoExit -File "%~dp0build_installer.ps1"

echo.
echo ===== Finished (check resources\installer_tools\logs\build_log_*.txt if failed) =====
pause

popd >nul
endlocal


