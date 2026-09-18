@echo off
chcp 65001 >nul
cd /d "%~dp0"
py -m pip install -r requirements.txt
py -m pip install pyinstaller
py -m PyInstaller --noconfirm --clean --onefile --windowed --name BankPayaReconciliation app.py
echo.
echo EXE created in:
echo %~dp0dist\BankPayaReconciliation.exe
pause
