@echo off
REM PPT Touch Controller Launcher
REM Use this to associate .pptx files: right-click .pptx → Open with → choose this .bat
REM Or run: python src/file_associator.py register

set PYTHON="C:\Users\Administrator\AppData\Local\Programs\Python\Python314\python.exe"
set SCRIPT=%~dp0src\main.py

if exist %PYTHON% (
    %PYTHON% %SCRIPT% %1
) else (
    python %SCRIPT% %1
)
