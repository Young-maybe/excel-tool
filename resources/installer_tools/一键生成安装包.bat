@echo off
setlocal

REM 目的：双击运行不会闪退，并绕过 PowerShell 执行策略限制
REM 日志：脚本会写入 dist_installer\build_log_*.txt

set ROOT=%~dp0..\..
pushd "%ROOT%" >nul

chcp 65001 >nul
REM NOTE: Chinese filename can be garbled under cmd codepage; call English entrypoint.
powershell -NoProfile -ExecutionPolicy Bypass -NoExit -File "%~dp0build_installer.ps1"

echo.
echo ===== 运行结束（如有报错请看 dist_installer\build_log_*.txt）=====
pause

popd >nul
endlocal


