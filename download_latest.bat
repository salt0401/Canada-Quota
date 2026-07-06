@echo off
setlocal
echo ============================================
echo   Canada TRQ Tracker - Download Latest Data
echo ============================================
echo.

where curl.exe >nul 2>&1
if errorlevel 1 (
    echo ERROR: This PC is missing curl.exe ^(included in Windows 10 1803+^).
    echo        Contact Lin for an alternative download method.
    goto :end
)

rem Resolve the REAL Desktop folder. With OneDrive "Known Folder Move" the
rem Desktop lives under %USERPROFILE%\OneDrive\Desktop, not %USERPROFILE%\Desktop.
for /f "delims=" %%D in ('powershell -NoProfile -Command "[Environment]::GetFolderPath('Desktop')"') do set "DESKTOP=%%D"
if not defined DESKTOP set "DESKTOP=%USERPROFILE%\Desktop"
if not exist "%DESKTOP%\" (
    echo ERROR: Could not find your Desktop folder at:
    echo        "%DESKTOP%"
    echo        Contact Lin.
    goto :end
)
set "TARGET=%DESKTOP%\canada_trq_tracker.xlsx"
set "TMPFILE=%TEMP%\canada_trq_tracker.xlsx.download"

echo Downloading latest Excel file from GitHub...
echo.
rem Download to a temp file first so a failed download or a locked Excel file
rem never leaves a broken/half-written workbook on the Desktop.
curl.exe -L --fail -o "%TMPFILE%" ^
  "https://github.com/salt0401/Canada-Quota/releases/latest/download/canada_trq_tracker.xlsx"
if errorlevel 1 (
    echo ERROR: Download failed. Check your internet connection.
    echo        If the internet is fine, the file may be temporarily
    echo        unavailable on GitHub - try again in a few minutes,
    echo        or contact Lin.
    goto :end
)

move /y "%TMPFILE%" "%TARGET%" >nul 2>&1
if errorlevel 1 (
    echo ERROR: Could not save to:
    echo        "%TARGET%"
    echo        The file is probably still OPEN in Excel.
    echo        Close it, then run this again.
    goto :end
)

echo SUCCESS: File saved to your Desktop as "canada_trq_tracker.xlsx"
echo.
echo Opening Excel file now...
start "" "%TARGET%"

:end
echo.
pause
