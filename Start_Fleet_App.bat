@echo off
setlocal
title Fleet Dashboard Launcher

set "LOCAL_DIR=%TEMP%\FleetDashboard_Env"

echo.
echo ==========================================
echo       Fleet Maintenance Dashboard
echo ==========================================
echo.

:: --------------------------------------------------
:: Check whether the virtual environment is usable
:: --------------------------------------------------

if not exist "%LOCAL_DIR%\Scripts\python.exe" (
    echo Virtual environment not found.
    goto CREATE_ENV
)

echo Checking existing Python environment...

"%LOCAL_DIR%\Scripts\python.exe" -m pip --version >nul 2>&1

if errorlevel 1 (
    echo.
    echo Existing virtual environment is damaged.
    echo Recreating it...
    goto RECREATE_ENV
)

echo Existing Python environment is OK.
goto INSTALL_DEPENDENCIES


:: --------------------------------------------------
:: Create new virtual environment
:: --------------------------------------------------

:CREATE_ENV

echo.
echo Creating local Python environment...
echo This normally happens only once.

python -m venv "%LOCAL_DIR%"

if errorlevel 1 (
    echo.
    echo ERROR: Could not create Python virtual environment.
    echo Make sure Python is installed correctly.
    pause
    exit /b 1
)

goto REPAIR_PIP


:: --------------------------------------------------
:: Recreate damaged environment
:: --------------------------------------------------

:RECREATE_ENV

echo.
echo Removing damaged environment...

rmdir /s /q "%LOCAL_DIR%"

if exist "%LOCAL_DIR%" (
    echo.
    echo ERROR: Could not remove damaged environment.
    echo Close any Python/Streamlit processes and try again.
    pause
    exit /b 1
)

echo.
echo Creating fresh Python environment...

python -m venv "%LOCAL_DIR%"

if errorlevel 1 (
    echo.
    echo ERROR: Could not create Python virtual environment.
    pause
    exit /b 1
)

goto REPAIR_PIP


:: --------------------------------------------------
:: Repair / bootstrap pip
:: --------------------------------------------------

:REPAIR_PIP

echo.
echo Repairing pip...

"%LOCAL_DIR%\Scripts\python.exe" -m ensurepip --upgrade

if errorlevel 1 (
    echo.
    echo ERROR: Could not install pip into the virtual environment.
    pause
    exit /b 1
)

"%LOCAL_DIR%\Scripts\python.exe" -m pip install --upgrade pip setuptools wheel

if errorlevel 1 (
    echo.
    echo ERROR: Could not upgrade pip.
    pause
    exit /b 1
)


:: --------------------------------------------------
:: Install application dependencies
:: --------------------------------------------------

:INSTALL_DEPENDENCIES

echo.
echo Installing / checking Fleet Dashboard dependencies...

"%LOCAL_DIR%\Scripts\python.exe" -m pip install ^
    streamlit ^
    pandas ^
    openpyxl ^
    xlwings ^
    pywin32 ^
    supabase

if errorlevel 1 (
    echo.
    echo ==========================================
    echo ERROR: Dependency installation failed.
    echo ==========================================
    echo.
    echo The Fleet Dashboard cannot start.
    echo.
    pause
    exit /b 1
)


:: --------------------------------------------------
:: Configure pywin32
:: --------------------------------------------------

echo.
echo Configuring Windows COM support...

"%LOCAL_DIR%\Scripts\python.exe" -m pywin32_postinstall -install

if errorlevel 1 (
    echo.
    echo WARNING: pywin32 post-install returned an error.
    echo The application may still work, but Outlook/Excel
    echo integration could be affected.
    echo.
)


:: --------------------------------------------------
:: Move to application directory
:: --------------------------------------------------

cd /d "%~dp0"

if errorlevel 1 (
    echo.
    echo ERROR: Could not access application directory.
    pause
    exit /b 1
)


:: --------------------------------------------------
:: Configure Streamlit
:: --------------------------------------------------

if not exist "%USERPROFILE%\.streamlit" (
    mkdir "%USERPROFILE%\.streamlit"
)

(
    echo [general]
    echo email = ""
) > "%USERPROFILE%\.streamlit\credentials.toml"


:: --------------------------------------------------
:: Check application file
:: --------------------------------------------------

if not exist "%~dp0app.py" (
    echo.
    echo ==========================================
    echo ERROR: app.py was not found.
    echo ==========================================
    echo.
    echo Expected:
    echo %~dp0app.py
    echo.
    pause
    exit /b 1
)


:: --------------------------------------------------
:: Launch Fleet Dashboard
:: --------------------------------------------------

echo.
echo ==========================================
echo Launching Fleet Dashboard...
echo ==========================================
echo.

"%LOCAL_DIR%\Scripts\python.exe" -m streamlit run "%~dp0app.py" --browser.gatherUsageStats=false

if errorlevel 1 (
    echo.
    echo ==========================================
    echo Streamlit failed to start.
    echo ==========================================
    echo.
    pause
)

endlocal