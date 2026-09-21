@echo off
title Fleet Dashboard Launcher

set "LOCAL_DIR=%TEMP%\FleetDashboard_Env"

:: Create local virtual environment if it doesn't exist
if not exist "%LOCAL_DIR%\Scripts\python.exe" (
    echo Setting up local environment in TEMP folder...
    echo This only happens on the first run...
    python -m venv "%LOCAL_DIR%"
)

:: Ensure all required dependencies are installed
"%LOCAL_DIR%\Scripts\python.exe" -m pip install streamlit pandas openpyxl xlwings pywin32 supabase

cd /d "%~dp0"

:: --------------------------------------------------
:: NEW: Bypass Streamlit Email Prompt
:: --------------------------------------------------
if not exist "%USERPROFILE%\.streamlit" mkdir "%USERPROFILE%\.streamlit"
echo [general]> "%USERPROFILE%\.streamlit\credentials.toml"
echo email = "" >> "%USERPROFILE%\.streamlit\credentials.toml"
:: --------------------------------------------------

echo Launching Fleet Dashboard...
"%LOCAL_DIR%\Scripts\python.exe" -m streamlit run app.py --browser.gatherUsageStats=false

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Streamlit failed to start.
    pause
)