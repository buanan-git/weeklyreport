@echo off
chcp 936 >nul
title WeeklyReportTool - Build Wizard

echo ========================================
echo   WeeklyReportTool - Build Wizard
echo ========================================
echo.

echo Please select build option:
echo.
echo  [1] PyInstaller + PyArmor (Recommended)
echo      - Build time: 2-5 minutes
echo      - Code protection: Medium
echo      - File size: ~50MB
echo      - Compatibility: Windows/Mac/Linux
echo.
echo  [2] Nuitka Compilation (Maximum Protection)
echo      - Build time: 10-30 minutes
echo      - Code protection: Maximum
echo      - File size: ~100MB+
echo      - Compatibility: Windows/Mac/Linux
echo.
echo  [3] Check Environment
echo.
echo  [Q] Quit
echo.

set /p choice="Enter option (1/2/3/Q): "

if /i "%choice%"=="1" goto pyinstaller
if /i "%choice%"=="2" goto nuitka
if /i "%choice%"=="3" goto check
if /i "%choice%"=="Q" goto end
goto invalid

:pyinstaller
echo.
echo ========================================
echo   Option 1: PyInstaller + PyArmor Build
echo ========================================
echo.
echo Starting build, please wait...
python "%~dp0build.py" --method pyinstaller
goto result

:nuitka
echo.
echo ========================================
echo   Option 2: Nuitka Compilation
echo ========================================
echo.
echo Warning: Nuitka compilation takes a long time (10-30 minutes)
echo.
set /p confirm="Confirm to start compilation? (Y/N): "
if /i not "%confirm%"=="Y" goto end

echo.
echo Starting compilation, please wait...
python "%~dp0build.py" --method nuitka
goto result

:check
echo.
echo ========================================
echo   Environment Check
echo ========================================
echo.
python "%~dp0build.py" --check
echo.
pause
goto end

:result
echo.
echo ========================================
if errorlevel 1 (
    echo   Build/Compilation FAILED, please check error messages
) else (
    echo   Build Completed Successfully!
)
echo ========================================
echo.
echo Output directory: %~dp0..\release\
echo.

:end
pause
