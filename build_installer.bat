@echo off
chcp 65001 >nul
REM ============================================================
REM   进度跟踪报表工具 - Windows 安装包构建
REM   前置：已安装 Inno Setup（iscc 在 PATH 中）
REM         并且已运行过 build.bat 生成 dist\*.exe
REM   产物：installer\进度跟踪报表工具_setup.exe
REM ============================================================
SETLOCAL
cd /d "%~dp0"

echo ========================================
echo   构建 Windows 安装包
echo ========================================
echo.

if not exist "dist\进度跟踪报表工具.exe" (
    echo [错误] 未找到 dist\进度跟踪报表工具.exe
    echo        请先运行 build.bat
    pause
    exit /b 1
)

where iscc >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 iscc（Inno Setup 编译器）
    echo        请从 https://jrsoftware.org/isdl.php 安装 Inno Setup
    pause
    exit /b 1
)

echo [1/2] 编译安装包...
iscc installer.iss
if errorlevel 1 (
    echo [错误] 安装包构建失败
    pause
    exit /b 1
)

echo.
echo [2/2] 完成!
echo 安装包: installer\进度跟踪报表工具_setup.exe
echo.
pause
ENDLOCAL
