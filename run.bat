@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>&1
if %errorlevel%==0 (
    py -m pip install -r requirements.txt
    py app.py
    goto :eof
)
where python >nul 2>&1
if %errorlevel%==0 (
    python -m pip install -r requirements.txt
    python app.py
    goto :eof
)
echo Python was not found.
echo Install Python 3.11+ from python.org and enable "Add Python to PATH".
pause
