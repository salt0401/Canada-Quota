@echo off
echo ============================================
echo   Canada TRQ Tracker - Download Latest Data
echo ============================================
echo.
echo Downloading latest Excel file from GitHub...
echo.
curl.exe -L -o "%USERPROFILE%\Desktop\canada_trq_tracker.xlsx" ^
  "https://github.com/salt0401/Canada-Quota/releases/latest/download/canada_trq_tracker.xlsx"
echo.
if %ERRORLEVEL% EQU 0 (
    echo SUCCESS: File saved to your Desktop as "canada_trq_tracker.xlsx"
    echo.
    echo Opening Excel file now...
    start "" "%USERPROFILE%\Desktop\canada_trq_tracker.xlsx"
) else (
    echo ERROR: Download failed. Check your internet connection.
)
echo.
pause
