@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
title PyRPA - PyInstaller 打包
cd /d "%~dp0"

set "SRC_DIR=%~dp0"
set "VENV_PY=%SRC_DIR%.venv\Scripts\python.exe"
set "DIST_DIR=%SRC_DIR%dist\pyrpa"

echo ============================================
echo  PyRPA 一键打包 (PyInstaller OneDir)
echo ============================================

REM ---------- 步骤1: 定位 Python ----------
set "PY=%VENV_PY%"
if not exist "%PY%" (
    echo [警告] 未找到项目虚拟环境: %VENV_PY%
    set PY=python
)
echo 使用 Python: %PY%
"%PY%" --version || goto :err

REM ---------- 步骤2: 安装/升级 PyInstaller ----------
echo.
echo [步骤2] 安装 / 升级 PyInstaller ...
where uv >nul 2>nul && (
    echo   使用 uv 安装 PyInstaller ...
    uv pip install --python "%PY%" "pyinstaller>=6.0" || goto :err
) || (
    "%PY%" -m pip install --upgrade "pyinstaller>=6.0" || goto :err
)

REM ---------- 步骤3: 清理旧产物 ----------
echo.
echo [步骤3] 清理旧构建产物 ...
if exist "%SRC_DIR%build" rmdir /s /q "%SRC_DIR%build"
if exist "%SRC_DIR%dist" rmdir /s /q "%SRC_DIR%dist"

REM ---------- 步骤4: 执行打包 ----------
echo.
echo [步骤4] 执行 PyInstaller 打包 (pyrpa.spec) ...
"%PY%" -m PyInstaller --noconfirm --clean "%SRC_DIR%pyrpa.spec" || goto :err

REM ---------- 步骤5: 附带运行时数据(配置/示例任务) ----------
REM 打包后 rpa_data 为 exe 旁可写数据目录, 不随程序内嵌; 把源码中的
REM 默认配置与示例任务复制过去, 保证首次运行开箱即用。
echo.
echo [步骤5] 复制默认运行时数据到输出目录 ...
if not exist "%DIST_DIR%\rpa_data" (
    xcopy /e /i /y "%SRC_DIR%rpa_data" "%DIST_DIR%\rpa_data" >nul
) else (
    echo     输出目录已存在 rpa_data, 跳过(避免覆盖用户数据)。
)

echo.
echo ============================================
echo  打包完成!
echo  输出目录: %DIST_DIR%
echo ============================================
goto :eof

:err
echo.
echo [错误] 打包失败, 请检查上方输出。
exit /b 1
