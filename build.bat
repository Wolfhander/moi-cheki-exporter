@echo off
setlocal
cd /d "%~dp0"
python -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto :fail
python -m PyInstaller --noconfirm moi_cheki_exporter.spec
if errorlevel 1 goto :fail
echo Done: dist\MoiCheki.exe
exit /b 0
:fail
echo Build failed.
exit /b 1
