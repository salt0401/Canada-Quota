@echo off
echo ============================================
echo   Canada TRQ Tracker - Get Latest Data
echo ============================================
echo.

cd /d "%~dp0"

echo Pulling latest data from GitHub...
echo.
git pull origin master

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ============================================
    echo   SUCCESS - Data is up to date!
    echo ============================================
    echo.
    echo Open the Excel file at:
    echo   %~dp0data\canada_trq_tracker.xlsx
    echo.
    echo Opening Excel file now...
    start "" "%~dp0data\canada_trq_tracker.xlsx"
) else (
    echo.
    echo ERROR: Pull failed. Check your internet connection.
)
echo.
pause
