@echo off
cd /d "%~dp0"
pip install flask -q
start http://localhost:5055
python web.py
pause
