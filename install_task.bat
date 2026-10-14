@echo off
REM ============================================================
REM  install_task.bat -- Setup daily sync task for MoiCheki
REM  Requests administrator rights via UAC automatically.
REM ============================================================

setlocal enableextensions

net session >nul 2>&1
if not errorlevel 1 goto :run_as_admin

powershell -NoProfile -Command "Start-Process cmd.exe -ArgumentList "/c \"%~f0\"" -Verb RunAs -Wait"
exit /b

:run_as_admin
echo.
echo Registering task in Windows Task Scheduler...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_task.ps1"

if errorlevel 1 (
    echo.
    echo ERROR: Task was not registered.
    pause
    exit /b 1
)

echo.
pause
endlocal
exit /b 0
