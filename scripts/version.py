#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
周报自动化工具 - 版本信息管理
所有项目版本信息统一在此文件中管理
"""

import os
import sys
from pathlib import Path
from datetime import datetime

# ==================== 项目基本信息 ====================
PROJECT_NAME = "周报自动化工具"
PROJECT_NAME_EN = "WeeklyReportTool"
VERSION = "1.2.0"
BUILD_DATE = datetime.now().strftime("%Y-%m-%d")

# ==================== Python版本要求 ====================
PYTHON_VERSION_MIN = "3.8"
PYTHON_VERSION_RECOMMEND = "3.12"

# ==================== 依赖版本 ====================
DEPENDENCIES = {
    "playwright": ">=1.40.0",
    "requests": ">=2.28.0",
    "pyperclip": ">=1.8.0",
}

# ==================== 开发依赖 ====================
DEV_DEPENDENCIES = {
    "pyinstaller": ">=6.3.0",
    "pyarmor": ">=8.3.0",
    "nuitka": ">=1.8.0",
}

# ==================== 路径配置 ====================
def get_project_root():
    """获取项目根目录"""
    # 如果是打包后的可执行文件
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    # 如果是源码运行
    return Path(__file__).parent.parent.absolute()

PROJECT_ROOT = get_project_root()

# 相对路径（相对于项目根目录）
PATHS = {
    "scripts": "scripts",
    "config": "config",
    "output": "output",
    "chrome_data": "chrome_data",
    "libs": "libs",
    "build": "build",
    "dist": "dist",
    "release": "release",
}

# ==================== 核心脚本列表 ====================
CORE_SCRIPTS = [
    "weeklyreport_auto.py",
    "fetch_okr_final.py",
    "llmapi_v10.py",
    "llmchat_final.py",
    "submit_okr_ds_final.py",
    "config_loader.py",
]

# 需要加密的核心脚本
PROTECTED_SCRIPTS = [
    "weeklyreport_auto.py",
    "fetch_okr_final.py",
    "llmapi_v10.py",
    "llmchat_final.py",
    "submit_okr_ds_final.py",
]

# 共享库模块
SHARED_MODULES = [
    "browser_automation.py",
]

# ==================== 打包配置 ====================
BUILD_CONFIG = {
    # PyInstaller 配置
    "pyinstaller": {
        "console": False,  # 隐藏控制台窗口
        "upx": True,  # 是否使用UPX压缩
        "onefile": False,  # 改为 False，避免启动时解压卡顿
        "optimize_level": 2,  # Python优化级别
    },
    # PyArmor 配置
    "pyarmor": {
        "enabled": True,  # 是否启用代码混淆
        "runtime": True,  # 是否包含运行时
    },
    # Nuitka 配置
    "nuitka": {
        "standalone": True,  # 独立模式
        "onefile": True,  # 单文件模式
        "disable_console": False,  # 禁用控制台（Windows）
    },
}

# ==================== 排除的模块（减少打包体积）====================
EXCLUDED_MODULES = [
    "tkinter",
    "matplotlib",
    "numpy",
    "pandas",
    "scipy",
    "PIL",
    "Pillow",
    "cv2",
    "opencv",
    "pytest",
    "sphinx",
    "IPython",
    "jupyter",
    "notebook",
]

# ==================== 强制包含的模块 ====================
INCLUDED_MODULES = [
    "playwright",
    "playwright.sync_api",
    "playwright.async_api",
    "requests",
    "pyperclip",
    "asyncio",
    "json",
    "threading",
    "http.server",
    "socketserver",
    "webbrowser",
    "urllib",
    "urllib.request",
    "urllib.parse",
    "base64",
    "hashlib",
    "hmac",
    "secretlib",  # 如果使用了 cryptography
    "aiohttp",
]

# ==================== 平台检测 ====================
def get_platform():
    """获取当前平台"""
    return sys.platform

def is_windows():
    return sys.platform == "win32"

def is_macos():
    return sys.platform == "darwin"

def is_linux():
    return sys.platform.startswith("linux")

PLATFORM = "windows" if is_windows() else ("macos" if is_macos() else "linux")

# ==================== 浏览器配置 ====================
BROWSER_CONFIG = {
    "windows": [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%USERPROFILE%\AppData\Local\Google\Chrome\Application\chrome.exe"),
    ],
    "macos": [
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        "/Applications/Chromium.app/Contents/MacOS/Chromium",
    ],
    "linux": [
        "google-chrome",
        "chromium",
        "chromium-browser",
    ],
}

# ==================== 帮助函数 ====================
def get_path(path_name):
    """获取项目路径"""
    if path_name in PATHS:
        return PROJECT_ROOT / PATHS[path_name]
    return PROJECT_ROOT / path_name

def get_scripts_dir():
    return get_path("scripts")

def get_config_dir():
    return get_path("config")

def get_build_dir():
    return get_path("build")

def get_dist_dir():
    return get_path("dist")

def get_release_dir():
    return get_path("release")

def print_version_info():
    """打印版本信息"""
    print(f"""
{'=' * 60}
  {PROJECT_NAME} v{VERSION}
  Build Date: {BUILD_DATE}
  Platform: {PLATFORM}
  Python: {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}
{'=' * 60}
""")

# ==================== 入口点 ====================
if __name__ == "__main__":
    print_version_info()
