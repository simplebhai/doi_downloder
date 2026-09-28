@echo off
setlocal
title Build DOI Paper Downloader
cd /d "%~dp0"
echo ============================================
echo   Building DOI_Paper_Downloader.exe
echo ============================================
where py >nul 2>nul && (set PY=py -3) || (set PY=python)
%PY% --version >nul 2>nul || (echo Python 3.9+ is required. Install it from https://www.python.org/downloads/ and tick "Add python.exe to PATH". & pause & exit /b 1)
if not exist .venv (%PY% -m venv .venv || (echo Could not create virtual environment. & pause & exit /b 1))
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller || (echo Package installation failed - check internet connection. & pause & exit /b 1)
pyinstaller --noconfirm --clean --onefile --windowed --name DOI_Paper_Downloader --hidden-import openpyxl --hidden-import xlrd --collect-all playwright paper_downloader.py || (echo Build failed. & pause & exit /b 1)
copy /y dist\DOI_Paper_Downloader.exe . >nul
echo.
echo DONE: DOI_Paper_Downloader.exe is ready in this folder.
pause
