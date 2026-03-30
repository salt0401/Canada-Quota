@echo off
echo ============================================
echo   Canada TRQ Tracker - First-Time Setup
echo ============================================
echo.
echo This will clone the GitHub repository to your Desktop.
echo You only need to run this ONCE.
echo.
echo IMPORTANT: You need Git installed on your computer.
echo   Download Git from: https://git-scm.com/download/win
echo.
pause

set REPO_URL=https://github.com/YOUR-USERNAME/Canada-Quota.git
set DEST=%USERPROFILE%\Desktop\Canada-Quota

if exist "%DEST%" (
    echo.
    echo Folder already exists: %DEST%
    echo If you want to start fresh, delete that folder first.
    echo.
    pause
    exit /b
)

echo.
echo Cloning repository to %DEST% ...
git clone %REPO_URL% "%DEST%"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ============================================
    echo   SUCCESS!
    echo ============================================
    echo.
    echo Repository cloned to: %DEST%
    echo.
    echo From now on, double-click "update.bat" inside
    echo that folder to get the latest data.
) else (
    echo.
    echo ERROR: Clone failed.
    echo Make sure Git is installed and the URL is correct.
    echo.
    echo Did you update REPO_URL in this file to your own GitHub repo URL?
)
echo.
pause
