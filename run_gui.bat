@echo off
rem Office Auto Backup -- GUI launcher (no console window)
rem Double-click this file to open the graphical interface.
rem It finds a working Python 3 + tkinter automatically; the "python"
rem on PATH may be the Microsoft Store alias on some machines.

setlocal EnableExtensions
cd /d "%~dp0"

set "SCRIPT=%~dp0gui.py"
if not exist "%SCRIPT%" (
    echo [ERROR] gui.py was not found next to this file.
    pause
    exit /b 1
)

rem 1) Official py launcher: verify Python + tkinter first, then run windowless.
py -3 -c "import tkinter" >nul 2>nul
if not errorlevel 1 (
    where pyw.exe >nul 2>nul
    if not errorlevel 1 (
        start "" pyw.exe -3 "%SCRIPT%"
        exit /b 0
    )
    rem pyw.exe is missing: use pythonw.exe beside the interpreter.
    for /f "delims=" %%P in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do (
        if exist "%%~dpPpythonw.exe" (
            start "" "%%~dpPpythonw.exe" "%SCRIPT%"
            exit /b 0
        )
    )
)

rem 2) pythonw.exe on PATH
where pythonw.exe >nul 2>nul
if not errorlevel 1 (
    start "" pythonw.exe "%SCRIPT%"
    exit /b 0
)

rem 3) Per-user installs: %LOCALAPPDATA%\Programs\Python\Python3*
for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do (
    if exist "%%~fD\pythonw.exe" (
        start "" "%%~fD\pythonw.exe" "%SCRIPT%"
        exit /b 0
    )
)

rem 4) Machine-wide installs: C:\Python3*
for /d %%D in ("C:\Python3*") do (
    if exist "%%~fD\pythonw.exe" (
        start "" "%%~fD\pythonw.exe" "%SCRIPT%"
        exit /b 0
    )
)

echo [ERROR] Python 3 + tkinter was not found.
echo         Install Python 3, then run: py -3 gui.py
echo.
pause
exit /b 1
