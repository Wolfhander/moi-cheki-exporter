@echo off
REM ============================================================
REM  sync_runner.bat — Ежедневная синхронизация чеков ФНС
REM  Запускается через Windows Task Scheduler.
REM  Лог: %USERPROFILE%\.moi_cheki_exporter\sync.log
REM ============================================================

setlocal enableextensions enabledelayedexpansion
cd /d "%~dp0"

REM ── Найти Python (та же логика, что в build.bat) ─────────────
set "PY="

where py >nul 2>nul
if not errorlevel 1 (
    py -3 -c "import sys" >nul 2>nul
    if not errorlevel 1 set "PY=py -3"
)

if "%PY%"=="" (
    for /f "usebackq delims=" %%i in (`where python 2^>nul`) do (
        echo %%i | findstr /I "WindowsApps" >nul
        if errorlevel 1 (
            if "%PY%"=="" set "PY=%%i"
        )
    )
)

if "%PY%"=="" (
    echo ОШИБКА: Python не найден. Установите Python 3.10+ с python.org.
    exit /b 99
)

REM ── Запуск sync.py ───────────────────────────────────────────
%PY% "%~dp0sync.py" %*
set "EXIT_CODE=%errorlevel%"

REM ── Обработка кодов выхода ──────────────────────────────────
if %EXIT_CODE%==2 (
    echo.
    echo [sync_runner] ОШИБКА: Токены не найдены.
    echo Откройте MoiCheki.exe и войдите через Госуслуги.
)
if %EXIT_CODE%==3 (
    echo.
    echo [sync_runner] ОШИБКА: Ошибка API ФНС. Возможно, требуется повторная авторизация.
)

exit /b %EXIT_CODE%
