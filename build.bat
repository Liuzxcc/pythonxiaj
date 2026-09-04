@echo off
chcp 65001 >nul
REM ============================================================
REM   进度跟踪报表工具 - Windows 打包脚本
REM   用法：在源码根目录（含 main.py）双击本文件，或在 cmd 中运行 build.bat
REM   前置：Windows + Python 3.9+（已加入 PATH）
REM   产物：dist\进度跟踪报表工具.exe
REM   注意：PyInstaller 不能跨平台编译，本脚本必须在 Windows 上运行。
REM ============================================================
SETLOCAL
cd /d "%~dp0"

echo ========================================
echo   进度跟踪报表工具 - Windows 打包
echo ========================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python，请先安装 Python 3.9+ 并勾选 "Add to PATH"
    pause
    exit /b 1
)

echo [1/3] 安装依赖包（xlrd==1.2.0 / xlwt / openpyxl / pyinstaller）...
python -m pip install -r requirements.txt pyinstaller -q
if errorlevel 1 (
    echo [错误] 依赖安装失败，请检查网络或 Python 环境
    pause
    exit /b 1
)

echo [2/3] 校验 xlrd 版本（必须 ^< 2.0，否则无法读取 .xls）...
python -c "import xlrd; assert int(xlrd.__version__.split('.')[0]) < 2, 'xlrd must be < 2.0'; print('xlrd OK:', xlrd.__version__)"
if errorlevel 1 (
    echo [错误] xlrd 版本不正确，请执行: pip install xlrd==1.2.0
    pause
    exit /b 1
)

echo [3/3] 开始打包（onefile / windowed）...
rem 用 spec 打包，hiddenimports 与 macOS 版完全一致（含 xlutils / xlutils.copy 等），
rem 避免漏依赖导致 exe 一启动就报「缺少模块」。
pyinstaller --noconfirm --windowed "进度跟踪报表工具_win.spec"
if errorlevel 1 (
    echo [错误] 打包失败，请查看上方日志
    pause
    exit /b 1
)

echo.
echo 打包完成!
echo 可执行文件: dist\进度跟踪报表工具.exe
echo.
pause
ENDLOCAL
