@echo off
title Energy Made Easy Enhanced Scraper
echo.
echo ============================================================
echo   Energy Made Easy Enhanced Scraper
echo   (with distributor selection for boundary postcodes)
echo ============================================================
echo.

set "VENV=%~dp0venv\Scripts\python.exe"

if not exist "%VENV%" (
    echo Virtual environment not found. Setting up...
    echo.
    python -m venv "%~dp0venv"
    "%VENV%" -m pip install -r "%~dp0requirements.txt"
    echo.
)

"%VENV%" "%~dp0scraper_enhanced.py"

if %ERRORLEVEL% EQU 0 (
    REM Find the most recently created spreadsheet and open it
    for /f "delims=" %%F in ('dir /b /od "%~dp0energy_plans_*.xlsx" 2^>nul') do set "LATEST=%%F"
    if defined LATEST (
        echo.
        echo   Opening %LATEST%...
        start "" "%~dp0%LATEST%"
    )
)

echo.
pause
