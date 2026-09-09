@echo off
setlocal

set "PROJECT_DIR=%~dp0"
set "PYTHON_EXE=%PROJECT_DIR%.venv\Scripts\python.exe"

if not exist "%PYTHON_EXE%" (
    echo [outlook-mcp] Missing virtual environment: "%PYTHON_EXE%" 1>&2
    echo [outlook-mcp] Run install.bat first. 1>&2
    exit /b 1
)

REM Full mode registers write tools. ChatGPT should be configured with
REM "Allow read actions" so reads run silently and writes require approval.
set "OUTLOOK_MCP_ACCESS=full"

"%PYTHON_EXE%" -m outlook_mcp
