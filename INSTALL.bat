@echo off
cd /d "%~dp0"
echo.
echo Installing required Python packages...
py -m pip install --upgrade pymupdf openpyxl
if errorlevel 1 (
  echo.
  echo Installation failed. Make sure Python 3.10 or newer is installed and available on PATH.
  pause
  exit /b 1
)
echo.
echo Installation complete.
echo Starting PDF to Excel Universal Converter...
py PDF_to_Excel_Universal_v4.py
pause
