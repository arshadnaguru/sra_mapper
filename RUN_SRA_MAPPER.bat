@echo off
title SRA Mapper - RIT OVPR
color 0A
cd /d "%~dp0"

python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo.
    echo  [!] Python not found. Please install Python first.
    echo      Download from python.org
    pause
    exit /b 1
)

pip show pandas openpyxl >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo  Installing dependencies...
    pip install pandas openpyxl --quiet
)

echo.
python sra_mapper.py
