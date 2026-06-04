@echo off
REM PPT Touch Controller Launcher - v2

set SCRIPT=%~dp0src\main.py
set PYTHON=

REM Try python from PATH first
where python >nul 2>&1
if %ERRORLEVEL% EQU 0 set PYTHON=python && goto :found

where python3 >nul 2>&1
if %ERRORLEVEL% EQU 0 set PYTHON=python3 && goto :found

REM Scan common install dirs
for %%d in (
    "%LOCALAPPDATA%\Programs\Python\Python314"
    "%LOCALAPPDATA%\Programs\Python\Python313"
    "%LOCALAPPDATA%\Programs\Python\Python312"
    "%LOCALAPPDATA%\Programs\Python\Python311"
    "%LOCALAPPDATA%\Programs\Python\Python310"
    "%LOCALAPPDATA%\Programs\Python\Python39"
    "C:\Program Files\Python314"
    "C:\Program Files\Python313"
    "C:\Python314"
    "C:\Python313"
) do (
    if exist "%%~d\python.exe" set PYTHON="%%~d\python.exe" && goto :found
)

REM Python not found
echo ERROR: Python not found.
echo Please install Python 3.9+ from https://www.python.org/downloads/
echo Make sure to check "Add Python to PATH" during installation.
pause
exit /b 1

:found
echo Starting PPT Touch Controller...
echo Python: %PYTHON%
echo.

%PYTHON% %SCRIPT% %*

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Program exited with error code: %ERRORLEVEL%
    pause
)
