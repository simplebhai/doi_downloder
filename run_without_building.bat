@echo off
cd /d "%~dp0"
where py >nul 2>nul && (set PY=py -3) || (set PY=python)
if not exist .venv %PY% -m venv .venv
call .venv\Scripts\activate.bat
python -m pip install -q -r requirements.txt
start "" pythonw paper_downloader.py
