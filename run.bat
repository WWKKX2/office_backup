@echo off
rem Office Auto Backup -- console launcher
rem Double-click to run in a visible window (close the window to stop).
rem This script finds a working Python automatically; it does not rely on
rem "python" being on PATH (on some machines that is the Store alias).

setlocal EnableExtensions
cd /d "%~dp0"

set "PY="

rem 1) Official py launcher
py -3 -c "import sys" >nul 2>nul
if not errorlevel 1 set "PY=py -3"

rem 2) Per-user installs: %LOCALAPPDATA%\Programs\Python\Python3*
if not defined PY (
    for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do (
        if not defined PY if exist "%%~fD\python.exe" set "PY="%%~fD\python.exe""
    )
)

rem 3) Machine-wide installs: C:\Python3*
if not defined PY (
    for /d %%D in ("C:\Python3*") do (
        if not defined PY if exist "%%~fD\python.exe" set "PY="%%~fD\python.exe""
    )
)

rem 4) Last resort: python on PATH
if not defined PY (
    python -c "import sys" >nul 2>nul
    if not errorlevel 1 set "PY=python"
)

if not defined PY (
    echo [ERROR] Python 3.8+ was not found.
    echo         Install Python, or put python.exe / pythonw.exe next to this file.
    echo.
    pause
    exit /b 1
)

echo Using: %PY%
echo.
%PY% main.py %*
set "CODE=%ERRORLEVEL%"

if not "%CODE%"=="0" (
    echo.
    echo Exit code: %CODE%
    pause
)

exit /b %CODE%
