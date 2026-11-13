@echo off
setlocal

REM Juanzai launcher. Keep this file pure ASCII: cmd.exe parses .bat line by
REM line using the OEM code page, so UTF-8 Chinese text or "chcp 65001" can
REM break parsing.

set "TOOL_DIR=%~dp0"
set "LIB=%TOOL_DIR%lib"

REM Portable mode: release builds ship a bundled lib\ folder. If it exists,
REM use it and ignore the user's own site-packages.
if exist "%LIB%\PIL\__init__.py" (
    set "PYTHONNOUSERSITE=1"
    set "PYTHONPATH=%LIB%"
    set "PATH=%LIB%\pywin32_system32;%PATH%"
    echo [INFO] Using bundled dependencies from lib\
) else (
    echo [INFO] No lib\ folder found. Using the system Python instead.
    echo        Install dependencies first: pip install -r requirements.txt
)

REM Do not hardcode an interpreter path: use whatever python is on PATH.
set "PY=pythonw.exe"
where pythonw.exe >nul 2>nul || set "PY=python.exe"
where %PY% >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python not found on PATH.
    echo         Install Python 3.9 or newer, then run: pip install -r requirements.txt
    pause
    exit /b 1
)

start "" "%PY%" "%TOOL_DIR%app.py"
endlocal