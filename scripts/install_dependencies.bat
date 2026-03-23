@echo off
chcp 65001 >nul
title 周报自动化工具 - 依赖安装

echo ========================================
echo     周报自动化工具 - 依赖安装
echo ========================================
echo.

echo 正在检查Python环境...
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到Python，请先安装Python 3.8+
    pause
    exit /b 1
)

REM 国内镜像配置
set PIP_MIRROR=https://pypi.tuna.tsinghua.edu.cn/simple
set PIP_TRUSTED=pypi.tuna.tsinghua.edu.cn
REM 跳过 Playwright 内置浏览器下载（本项目使用系统 Chrome，不需要 Playwright 的 Chromium）
set PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD=1

echo 使用清华 PyPI 镜像加速下载...
echo.

echo 正在安装核心依赖...
python -m pip install --upgrade pip -i %PIP_MIRROR% --trusted-host %PIP_TRUSTED%
python -m pip install playwright requests -i %PIP_MIRROR% --trusted-host %PIP_TRUSTED%

REM 注意：不执行 playwright install chromium
REM 本项目通过 CDP 连接系统已安装的 Chrome，无需 Playwright 内置浏览器

echo.
echo ========================================
echo 依赖安装完成！
echo 本工具使用系统已安装的 Chrome 浏览器
echo 请确保已安装 Google Chrome
echo ========================================
pause
