@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
title PyRPA - PyInstaller Build
cd /d "%~dp0"

set "SRC_DIR=%~dp0"
set "VENV_PY=%SRC_DIR%.venv\Scripts\python.exe"
set "DIST_DIR=%SRC_DIR%dist\pyrpa"

echo ============================================
echo  PyRPA one-click build (PyInstaller OneDir)
echo ============================================

REM ---------- Step 1: Locate Python ----------
set "PY=%VENV_PY%"
if not exist "%PY%" (
    echo [WARN] Project venv not found: %VENV_PY%
    set PY=python
)
echo Using Python: %PY%
"%PY%" --version || goto :err

REM ---------- Step 2: Install/Upgrade PyInstaller ----------
echo.
echo [Step 2] Install / Upgrade PyInstaller ...
where uv >nul 2>nul && (
    echo   Installing PyInstaller with uv ...
    uv pip install --python "%PY%" "pyinstaller>=6.0" || goto :err
) || (
    "%PY%" -m pip install --upgrade "pyinstaller>=6.0" || goto :err
)

REM ---------- Step 3: Clean old artifacts ----------
echo.
echo [Step 3] Cleaning old build artifacts ...
if exist "%SRC_DIR%build" rmdir /s /q "%SRC_DIR%build"
if exist "%SRC_DIR%dist" rmdir /s /q "%SRC_DIR%dist"

REM ---------- Step 4: Run PyInstaller ----------
echo.
echo [Step 4] Running PyInstaller build (pyrpa.spec) ...
"%PY%" -m PyInstaller --noconfirm --clean "%SRC_DIR%pyrpa.spec" || goto :err

REM ---------- Step 5: Copy runtime data (config/example tasks) ----------
REM After packaging, rpa_data is a writable data directory beside the exe; it is
REM not bundled inside the program. Copy the source default config and example
REM tasks there so first run works out of the box.
echo.
echo [Step 5] Copying default runtime data to output directory ...
if not exist "%DIST_DIR%\rpa_data" (
    xcopy /e /i /y "%SRC_DIR%rpa_data" "%DIST_DIR%\rpa_data" >nul
) else (
    echo     Output dir already has rpa_data, skipped to avoid overwriting user data.
)

echo.
echo ============================================
echo  Build finished!
echo  Output directory: %DIST_DIR%
echo ============================================
goto :eof

:err
echo.
echo [ERROR] Build failed, please check the output above.
exit /b 1