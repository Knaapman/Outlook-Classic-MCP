@echo off
REM ---------------------------------------------------------------------
REM outlook-classic-mcp Windows installer (uv-based, hardened fork)
REM
REM 1. Installs uv through winget if missing.
REM 2. Creates .venv with Python 3.11.
REM 3. Installs the package in editable mode (-e .).
REM 4. Pre-warms the pywin32 typelib for Outlook.
REM 5. Launches scripts\install_to_clients.py for local MCP clients.
REM
REM Security: this installer does NOT pipe remote PowerShell into iex.
REM The server defaults to read-only access unless OUTLOOK_MCP_ACCESS=full
REM is explicitly set before startup.
REM ---------------------------------------------------------------------

setlocal enabledelayedexpansion

set "INSTALL_DIR=%~dp0"
if "%INSTALL_DIR:~-1%"=="\" set "INSTALL_DIR=%INSTALL_DIR:~0,-1%"

set "VENV_DIR=%INSTALL_DIR%\.venv"
set "PYTHON_EXE=%VENV_DIR%\Scripts\python.exe"

echo.
echo ========================================================================
echo   outlook-classic-mcp hardened installer
echo   Install location: %INSTALL_DIR%
echo   Default access: READ ONLY
echo ========================================================================
echo.

REM ---- Check for uv, install with winget if missing ----
where uv >nul 2>nul
if %errorlevel% neq 0 (
    echo [1/5] uv not found. Installing Astral uv with winget ...
    where winget >nul 2>nul
    if !errorlevel! neq 0 (
        echo.
        echo [error] winget is not available.
        echo         Install uv manually from the official Astral documentation,
        echo         then re-run this installer.
        echo.
        pause
        exit /b 1
    )
    winget install --id=astral-sh.uv -e --accept-source-agreements --accept-package-agreements
    if !errorlevel! neq 0 (
        echo.
        echo [error] winget could not install uv.
        echo         Install uv manually from the official Astral documentation,
        echo         then re-run this installer.
        echo.
        pause
        exit /b 1
    )
    set "PATH=%USERPROFILE%\.local\bin;%LOCALAPPDATA%\Microsoft\WinGet\Links;%PATH%"
    where uv >nul 2>nul
    if !errorlevel! neq 0 (
        echo.
        echo [error] uv was installed but is not visible on PATH yet.
        echo         Open a NEW terminal and re-run install.bat.
        echo.
        pause
        exit /b 1
    )
) else (
    echo [1/5] uv is already installed.
)

REM ---- Create venv with Python 3.11 ----
echo [2/5] Creating virtual environment (Python 3.11) ...
cd /d "%INSTALL_DIR%"
uv venv --python 3.11 "%VENV_DIR%"
if %errorlevel% neq 0 (
    echo.
    echo [error] uv venv creation failed.
    pause
    exit /b 1
)

REM ---- Install the package in editable mode ----
echo [3/5] Installing outlook-classic-mcp and dependencies ...
uv pip install --python "%PYTHON_EXE%" -e "%INSTALL_DIR%"
if %errorlevel% neq 0 (
    echo.
    echo [error] Package install failed.
    pause
    exit /b 1
)

REM ---- Pre-warm pywin32 typelib ----
echo [4/5] Pre-warming pywin32 typelib cache for Outlook ...
> "%TEMP%\_outlook_mcp_warmup.py" echo import win32com.client; win32com.client.gencache.EnsureDispatch('Outlook.Application')
"%PYTHON_EXE%" "%TEMP%\_outlook_mcp_warmup.py" 2>nul
del /q "%TEMP%\_outlook_mcp_warmup.py" 2>nul
if %errorlevel% neq 0 (
    echo [warn] Could not pre-warm typelib. Fine if Outlook is not running yet.
)

REM ---- Smart client installer ----
echo [5/5] Launching local client installer ...
echo.
"%PYTHON_EXE%" "%INSTALL_DIR%\scripts\install_to_clients.py"

echo.
echo ========================================================================
echo   INSTALL COMPLETE - READ ONLY MODE

echo   Do not set OUTLOOK_MCP_ACCESS=full unless write access is intentionally
 echo   required. Full mode enables send, delete, move and other mutations.
echo ========================================================================
echo.
echo   Standalone smoke test:
echo     "%PYTHON_EXE%" -m outlook_mcp
echo.
echo   MCP Inspector:
set "FWD_PY=%PYTHON_EXE:\=/%"
echo     npx @modelcontextprotocol/inspector "%FWD_PY%" -m outlook_mcp
echo.
echo ========================================================================
echo.
pause

endlocal
