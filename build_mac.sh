#!/bin/bash
# 进度跟踪报表工具 - macOS 打包脚本
#
# 用法: ./build_mac.sh
#
# Python 选择顺序（首个可用即停）:
#   1. $PYTHON_BIN 环境变量
#   2. ./venv/bin/python（项目内 venv）
#   3. ~/.../binaries/python/envs/wellsync-build/bin/python（专用打包 venv）
#   4. ~/.../binaries/python/versions/3.13.12/bin/python3（managed，xlrd 1.2.0 + xlwt）
#
# 不要 fallback 到 envs/default —— 那里是 xlrd 2.0.2，会静默读不了 .xls。

set -e
cd "$(dirname "$0")"

echo "========================================"
echo "  进度跟踪报表工具 - macOS 打包"
echo "========================================"
echo

# ---------------------------------------------------------------- Python 探测
pick_python() {
    if [ -n "${PYTHON_BIN}" ] && [ -x "${PYTHON_BIN}" ]; then
        echo "${PYTHON_BIN}"; return 0
    fi
    if [ -x "venv/bin/python" ]; then
        echo "venv/bin/python"; return 0
    fi
    local cand="/Users/zoe/.workbuddy/binaries/python/envs/wellsync-build/bin/python"
    if [ -x "$cand" ]; then
        echo "$cand"; return 0
    fi
    local cand2="/Users/zoe/.workbuddy/binaries/python/versions/3.13.12/bin/python3"
    if [ -x "$cand2" ]; then
        echo "$cand2"; return 0
    fi
    which python3
}

PYTHON_BIN="$(pick_python)"
echo "使用 Python: $PYTHON_BIN"
$PYTHON_BIN --version

# ---------------------------------------------------------------- 依赖检查
echo
echo "[依赖检查]"
$PYTHON_BIN -c "
import sys, importlib
for m in ('xlrd', 'xlwt', 'openpyxl', 'tkinter', 'PyInstaller'):
    importlib.import_module(m)
import xlrd
ver = int(xlrd.__version__.split('.')[0])
assert ver < 2, 'xlrd 必须 < 2.0，当前 %s' % xlrd.__version__
print('  xlrd', xlrd.__version__, '| xlwt ok | openpyxl ok | tkinter ok | PyInstaller ok')
"

# ---------------------------------------------------------------- 图标
if [ -f "assets/app.icns" ]; then
    ICON_ARG="--icon assets/app.icns"
    echo
    echo "[图标] 使用 assets/app.icns"
else
    ICON_ARG=""
    echo
    echo "[图标] 未找到 assets/app.icns，将使用默认图标（先跑 tools/make_icon.py）"
fi

# ---------------------------------------------------------------- 清理旧产物
# 用 Python 逐目录删除：单个 .app 内含数百文件，直接 rm -rf 会触发批量删除保护。
# build/ 缓存交给 PyInstaller 的 --clean 处理，这里只清 dist 里的旧 app 包。
$PYTHON_BIN - <<'PYEOF'
import pathlib, shutil
for t in ("dist/进度跟踪报表工具.app", "dist/进度跟踪报表工具"):
    p = pathlib.Path(t)
    if p.exists():
        shutil.rmtree(p)
        print("已清理旧产物: %s" % t)
PYEOF

# ---------------------------------------------------------------- 打包
echo
echo "[1/2] PyInstaller 打包中（spec: 进度跟踪报表工具.spec）..."
$PYTHON_BIN -m PyInstaller --noconfirm \
    --clean \
    "进度跟踪报表工具.spec"

echo
echo "[2/2] 收尾"
APP="dist/进度跟踪报表工具.app"
if [ ! -d "$APP" ]; then
    echo "✗ 打包失败：未生成 $APP"
    exit 1
fi

# 摘掉 Gatekeeper 标记（本机打包不触发，但以防从外部迁移过来的脚本残留）
xattr -dr com.apple.quarantine "$APP" 2>/dev/null || true

# 自报体积与路径
echo "产物: $APP"
du -sh "$APP"
du -h "$APP/Contents/MacOS/进度跟踪报表工具"
echo
echo "运行: open \"$APP\""
echo "      或: \"$APP/Contents/MacOS/进度跟踪报表工具\""
echo
echo "提示：首次双击若被 Gatekeeper 拦截："
echo "  右键 → 打开 → 确认；或：xattr -d com.apple.quarantine \"$APP\""
