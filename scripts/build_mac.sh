#!/usr/bin/env bash
# ============================================================
#  WeeklyReportTool - macOS 一键打包脚本
#  用法: ./scripts/build_mac.sh
# ============================================================
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
cd "$PROJECT_DIR"

echo "========================================"
echo "  WeeklyReportTool - macOS 打包"
echo "========================================"
echo

# ---------- 1. 环境检查 ----------
echo "[1/5] 检查环境..."

# Python 3
if ! command -v python3 &>/dev/null; then
    echo "❌ 未找到 python3，请先安装："
    echo "   brew install python@3.12"
    exit 1
fi

PY_VER=$(python3 -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')")
echo "  Python: $PY_VER"

PY_MAJOR=$(python3 -c "import sys; print(sys.version_info.major)")
PY_MINOR=$(python3 -c "import sys; print(sys.version_info.minor)")
if [ "$PY_MAJOR" -lt 3 ] || { [ "$PY_MAJOR" -eq 3 ] && [ "$PY_MINOR" -lt 8 ]; }; then
    echo "❌ 需要 Python 3.8+，当前 $PY_VER"
    exit 1
fi

# Xcode Command Line Tools
if ! xcode-select -p &>/dev/null; then
    echo "⚠️  未安装 Xcode Command Line Tools，正在安装..."
    xcode-select --install
    echo "   安装完成后请重新运行此脚本"
    exit 1
fi
echo "  Xcode CLI Tools: 已安装"

# Chrome / Edge
CHROME_APP="/Applications/Google Chrome.app"
EDGE_APP="/Applications/Microsoft Edge.app"
if [ -d "$CHROME_APP" ]; then
    echo "  Chrome: 已安装"
elif [ -d "$EDGE_APP" ]; then
    echo "  Edge: 已安装（可替代Chrome）"
else
    echo "  ⚠️  未找到 Chrome 或 Edge！本工具运行时必须有 Chromium 内核浏览器"
    echo "     安装 Chrome: https://www.google.com/chrome/"
    echo "     或: brew install --cask google-chrome"
    echo "     Edge 也可以: brew install --cask microsoft-edge"
    echo "     (打包可以继续，但运行前必须安装)"
fi

echo

# ---------- 2. 安装依赖 ----------
echo "[2/5] 安装依赖..."

# 国内镜像配置
PIP_MIRROR="https://pypi.tuna.tsinghua.edu.cn/simple"
PIP_TRUSTED="pypi.tuna.tsinghua.edu.cn"
# 跳过 Playwright 内置浏览器下载（本项目使用系统 Chrome，不需要 Playwright 的 Chromium）
export PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD=1

echo "  使用清华 PyPI 镜像加速下载..."

echo "  安装 Python 依赖..."
pip3 install -i "$PIP_MIRROR" --trusted-host "$PIP_TRUSTED" -r requirements.txt
pip3 install -i "$PIP_MIRROR" --trusted-host "$PIP_TRUSTED" pyinstaller pyarmor 2>/dev/null || true
echo "  依赖安装完成"

# Playwright Python 包（只需要包，不需要下载浏览器）
if ! python3 -c "from playwright.sync_api import sync_playwright" &>/dev/null; then
    echo "  安装 Playwright Python 包..."
    pip3 install -i "$PIP_MIRROR" --trusted-host "$PIP_TRUSTED" playwright
fi
# 注意：不执行 playwright install chromium
# 本项目通过 CDP 连接系统已安装的 Chrome，无需 Playwright 内置浏览器
echo "  Playwright: 就绪（使用系统 Chrome）"
echo

# ---------- 3. 打包 ----------
echo "[3/5] 开始打包..."
echo

python3 scripts/build.py --method pyinstaller

echo

# ---------- 4. macOS 签名 ----------
echo "[4/5] macOS 签名处理..."

# 找到发布目录
RELEASE_DIR=$(ls -d release/WeeklyReportTool_v*_pyinstaller 2>/dev/null | head -1)
if [ -z "$RELEASE_DIR" ]; then
    echo "❌ 未找到发布目录"
    exit 1
fi

EXE_FILE="$RELEASE_DIR/WeeklyReportTool"
if [ -f "$EXE_FILE" ]; then
    # 清除隔离属性
    xattr -cr "$RELEASE_DIR/" 2>/dev/null || true
    echo "  ✅ 已清除隔离属性"

    # ad-hoc 签名
    codesign --force --deep --sign - "$EXE_FILE" 2>/dev/null || true
    echo "  ✅ 已完成 ad-hoc 签名"

    # 确保可执行权限
    chmod +x "$EXE_FILE"
    echo "  ✅ 已设置可执行权限"
else
    echo "  ⚠️  未找到可执行文件: $EXE_FILE"
fi

echo

# ---------- 5. 完成 ----------
echo "[5/5] 打包完成！"
echo
echo "========================================"
echo "  输出文件"
echo "========================================"

ZIP_FILE=$(ls release/WeeklyReportTool_v*_macos.zip 2>/dev/null | head -1)
echo "  发布目录: $PROJECT_DIR/$RELEASE_DIR"
if [ -n "$ZIP_FILE" ]; then
    echo "  ZIP 压缩包: $PROJECT_DIR/$ZIP_FILE"
fi
echo
echo "  运行方式: ./$RELEASE_DIR/WeeklyReportTool"
echo "========================================"
