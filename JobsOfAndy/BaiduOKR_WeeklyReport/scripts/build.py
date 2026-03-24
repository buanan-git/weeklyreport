#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
周报自动化工具 - 统一打包脚本
功能：
1. 支持多平台打包（Windows、Mac、Linux）
2. 支持代码混淆（PyArmor）
3. 支持两种打包方案：PyInstaller 和 Nuitka
4. 统一版本管理
5. 增强路径容错性
"""

import os
import sys
import shutil
import subprocess
import json
import zipfile
import platform
import argparse
from pathlib import Path
from datetime import datetime
from typing import Optional, Tuple

# 设置UTF-8编码
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# 导入版本管理
SCRIPT_DIR = Path(__file__).parent
sys.path.insert(0, str(SCRIPT_DIR))
try:
    from version import (
        PROJECT_NAME, PROJECT_NAME_EN, VERSION, BUILD_DATE,
        CORE_SCRIPTS, PROTECTED_SCRIPTS, SHARED_MODULES,
        BUILD_CONFIG, EXCLUDED_MODULES, INCLUDED_MODULES,
        PLATFORM, BROWSER_CONFIG, get_platform,
        get_project_root, get_scripts_dir, get_config_dir,
        get_build_dir, get_dist_dir, get_release_dir,
        DEPENDENCIES, DEV_DEPENDENCIES
    )
except ImportError:
    # 如果 version.py 不存在，使用默认值
    PROJECT_NAME = "周报自动化工具"
    PROJECT_NAME_EN = "WeeklyReportTool"
    VERSION = "1.1.0"
    BUILD_DATE = datetime.now().strftime("%Y-%m-%d")
    PLATFORM = "windows" if sys.platform == "win32" else ("macos" if sys.platform == "darwin" else "linux")

    def get_project_root():
        return Path(__file__).parent.parent.absolute()

    def get_scripts_dir():
        return get_project_root() / "scripts"

    def get_config_dir():
        return get_project_root() / "config"

    def get_build_dir():
        return get_project_root() / "build"

    def get_dist_dir():
        return get_project_root() / "dist"

    def get_release_dir():
        return get_project_root() / "release"

    CORE_SCRIPTS = [
        "weeklyreport_auto.py",
        "fetch_okr_final.py",
        "llmapi_v10.py",
        "llmchat_final.py",
        "submit_okr_ds_final.py",
        "config_loader.py",
    ]

    SHARED_MODULES = [
        "browser_automation.py",
    ]

    EXCLUDED_MODULES = [
        "tkinter", "matplotlib", "numpy", "pandas", "scipy",
        "PIL", "Pillow", "cv2", "opencv", "pytest", "sphinx",
        "IPython", "jupyter", "notebook",
    ]

    INCLUDED_MODULES = [
        "playwright", "playwright.sync_api", "playwright.async_api",
        "requests", "pyperclip", "asyncio", "json", "threading",
        "http.server", "socketserver", "webbrowser",
        "aiohttp",
    ]

    BUILD_CONFIG = {
        "pyinstaller": {"console": False, "upx": True, "onefile": True},
        "pyarmor": {"enabled": True, "runtime": True},
    }

# ==================== 路径配置 ====================
BASE_DIR = get_project_root()
SCRIPTS_DIR = get_scripts_dir()
CONFIG_DIR = get_config_dir()
BUILD_DIR = get_build_dir()
DIST_DIR = get_dist_dir()
RELEASE_DIR = get_release_dir()

# 共享库目录（兼容旧路径）
LIBS_DIR = BASE_DIR.parent / "libs"
if not LIBS_DIR.exists():
    LIBS_DIR = BASE_DIR / "libs"

# TokenLess 共享脚本目录
TOKENLESS_SCRIPTS_DIR = BASE_DIR.parent.parent / "JobsOfAndy" / "TokenLess" / "scripts"
if not TOKENLESS_SCRIPTS_DIR.exists():
    TOKENLESS_SCRIPTS_DIR = BASE_DIR.parent / "TokenLess" / "scripts"
if not TOKENLESS_SCRIPTS_DIR.exists():
    TOKENLESS_SCRIPTS_DIR = Path(__file__).parent.parent.parent / "TokenLess" / "scripts"

# ==================== 国内镜像配置 ====================
# pip 使用清华镜像（国内最快最稳定）
PIP_MIRROR = "https://pypi.tuna.tsinghua.edu.cn/simple"
PIP_TRUSTED_HOST = "pypi.tuna.tsinghua.edu.cn"

# Playwright 浏览器下载使用 npmmirror（备用，当前项目不下载浏览器）
PLAYWRIGHT_MIRROR = "https://npmmirror.com/mirrors/playwright"

# 设置环境变量（对所有子进程生效）
os.environ["PLAYWRIGHT_DOWNLOAD_HOST"] = PLAYWRIGHT_MIRROR

# ==================== 工具函数 ====================
def print_header(text: str) -> None:
    """打印标题"""
    print("\n" + "=" * 70)
    print(f"  {text}")
    print("=" * 70)

def print_step(step: int, text: str) -> None:
    """打印步骤"""
    print(f"\n[步骤 {step}] {text}")
    print("-" * 50)

def print_success(text: str) -> None:
    """打印成功信息"""
    print(f"  \u2705 {text}")

def print_warning(text: str) -> None:
    """打印警告信息"""
    print(f"  \u26a0\ufe0f {text}")

def print_error(text: str) -> None:
    """打印错误信息"""
    print(f"  \u274c {text}")

def run_command(cmd: list, cwd: Optional[Path] = None, check: bool = True,
                show_output: bool = False) -> Tuple[bool, str]:
    """运行命令"""
    cmd_str = ' '.join(str(c) for c in cmd)
    print(f"  执行: {cmd_str}")

    try:
        if show_output:
            # 实时显示输出
            process = subprocess.Popen(
                cmd, cwd=cwd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding='utf-8',
                errors='replace'
            )
            output_lines = []
            for line in process.stdout:
                print(f"    {line.rstrip()}")
                output_lines.append(line)

            process.wait()
            output = ''.join(output_lines)

            if process.returncode != 0:
                if check:
                    print_error(f"\u547d\u4ee4\u6267\u884c\u5931\u8d25\uff0c\u8fd4\u56de\u7801: {process.returncode}")
                return False, output
            return True, output
        else:
            # 静默执行
            result = subprocess.run(
                cmd, cwd=cwd,
                capture_output=True,
                text=True,
                encoding='utf-8',
                errors='replace'
            )

            if result.returncode != 0:
                if check:
                    print_error(f"\u547d\u4ee4\u6267\u884c\u5931\u8d25: {result.stderr or result.stdout}")
                return False, result.stderr or result.stdout

            # \u663e\u793a\u90e8\u5206\u8f93\u51fa
            if result.stdout:
                lines = result.stdout.strip().split('\n')
                for line in lines[:10]:
                    print(f"    {line}")

            return True, result.stdout

    except Exception as e:
        if check:
            print_error(f"\u6267\u884c\u5f02\u5e38: {e}")
        return False, str(e)

def check_dir(path: Path, create: bool = True) -> bool:
    """\u68c0\u67e5\u76ee\u5f55\u662f\u5426\u5b58\u5728\uff0c\u4e0d\u5b58\u5728\u5219\u521b\u5efa"""
    if not path.exists():
        if create:
            path.mkdir(parents=True, exist_ok=True)
            print(f"  \u5df2\u521b\u5efa\u76ee\u5f55: {path}")
        return False
    return True

def clean_dir(path: Path) -> None:
    """\u6e05\u7406\u76ee\u5f55"""
    if path.exists():
        shutil.rmtree(path)
        print(f"  \u5df2\u6e05\u7406: {path}")

# ==================== \u73af\u5883\u68c0\u67e5 ====================
def check_python_version() -> bool:
    """\u68c0\u67e5Python\u7248\u672c"""
    print_step(1, "\u68c0\u67e5Python\u73af\u5883")
    version = sys.version_info
    print(f"  Python\u7248\u672c: {version.major}.{version.minor}.{version.micro}")

    if version.major < 3 or (version.major == 3 and version.minor < 8):
        print_error(f"Python\u7248\u672c\u8fc7\u4f4e\uff0c\u9700\u89813.8+\uff0c\u5f53\u524d: {version.major}.{version.minor}")
        return False

    print_success(f"Python\u7248\u672c\u7b26\u5408\u8981\u6c42")
    return True

def check_required_dirs() -> bool:
    """\u68c0\u67e5\u5fc5\u8981\u7684\u76ee\u5f55"""
    print_step(2, "\u68c0\u67e5\u9879\u76ee\u76ee\u5f55")

    dirs_to_check = {
        "scripts": SCRIPTS_DIR,
        "config": CONFIG_DIR,
    }

    all_ok = True
    for name, path in dirs_to_check.items():
        if path.exists():
            print_success(f"\u76ee\u5f55\u5b58\u5728: {name}/")
        else:
            print_error(f"\u76ee\u5f55\u7f3a\u5931: {name}/")
            all_ok = False

    # \u5c1d\u8bd5\u521b\u5faabuild/dist/release\u76ee\u5f55
    for name in ["build", "dist", "release"]:
        path = BASE_DIR / name
        check_dir(path, create=True)

    return all_ok

def check_build_tools(method: str) -> bool:
    """\u68c0\u67e5\u6253\u5305\u5de5\u5177"""
    print_step(3, f"\u68c0\u67e5{method}\u5de5\u5177")

    if method == "pyinstaller":
        try:
            result = subprocess.run(
                [sys.executable, "-m", "PyInstaller", "--version"],
                capture_output=True, text=True
            )
            version = result.stdout.strip()
            print_success(f"PyInstaller: {version}")
            return True
        except Exception:
            print_error("PyInstaller\u672a\u5b89\u88c5")
            print("  \u5b89\u88c5\u547d\u4ee4: pip install pyinstaller>=6.3.0")
            return False

    elif method == "nuitka":
        try:
            result = subprocess.run(
                [sys.executable, "-m", "nuitka", "--version"],
                capture_output=True, text=True
            )
            version = result.stdout.strip()
            print_success(f"Nuitka: {version}")

            # \u68c0\u67e5C\u7f16\u8bd1\u5668
            if sys.platform == "win32":
                # \u68c0\u67e5Visual Studio\u6216MinGW
                has_vs = shutil.which("vswhere") or shutil.which("cl.exe")
                has_gcc = shutil.which("gcc")
                if has_vs or has_gcc:
                    print_success("C\u7f16\u8bd1\u5668\u5df2\u5b89\u88c5")
                else:
                    print_warning("\u672a\u68c0\u6d4b\u5230C\u7f16\u8bd1\u5668\uff0cNuitka\u5c06\u81ea\u52a8\u4e0b\u8f7dMinGW")
            elif sys.platform == "darwin":
                result = subprocess.run(["xcode-select", "-p"], capture_output=True)
                if result.returncode == 0:
                    print_success("Xcode Command Line Tools\u5df2\u5b89\u88c5")
                else:
                    print_warning("\u8bf7\u5b89\u88c5Xcode Command Line Tools: xcode-select --install")
            else:
                if shutil.which("gcc"):
                    print_success("GCC\u5df2\u5b89\u88c5")
                else:
                    print_warning("\u8bf7\u5b89\u88c5GCC: sudo apt install build-essential")

            return True
        except Exception:
            print_error("Nuitka\u672a\u5b89\u88c5")
            print("  \u5b89\u88c5\u547d\u4ee4: pip install nuitka")
            return False

    return False

def install_dependencies(method: str) -> bool:
    """\u5b89\u88c5\u4f9d\u8d56"""
    print_step(4, "\u5b89\u88c5\u4f9d\u8d56")

    deps = []
    if method == "pyinstaller":
        deps = ["pyinstaller>=6.3.0", "pyarmor>=8.3.0"]
    elif method == "nuitka":
        deps = ["nuitka>=1.8.0", "ordered-set", "zstandard"]

    print(f"  使用清华 PyPI 镜像: {PIP_MIRROR}")
    for dep in deps:
        print(f"  安装 {dep}...")
        success, _ = run_command(
            [sys.executable, "-m", "pip", "install",
             "-i", PIP_MIRROR,
             "--trusted-host", PIP_TRUSTED_HOST,
             dep],
            check=False
        )
        if not success:
            print_warning(f"安装 {dep} 失败，请手动安装")

    return True

# ==================== \u6e90\u7801\u51c6\u5907 ====================
def prepare_source() -> Optional[Path]:
    """\u51c6\u5907\u6e90\u7801\u4ee3\u7801"""
    print_step(5, "\u51c6\u5907\u6e90\u7801\u4ee3\u7801")

    # \u6e05\u7406\u6784\u5efa\u76ee\u5f55
    build_work = BUILD_DIR / f"work_{PLATFORM}"
    if build_work.exists():
        shutil.rmtree(build_work)
    build_work.mkdir(parents=True, exist_ok=True)

    # \u590d\u5236main.py
    main_src = BASE_DIR / "main.py"
    if not main_src.exists():
        main_src = SCRIPTS_DIR / "main.py"

    if main_src.exists():
        shutil.copy2(main_src, build_work / "main.py")
        print_success(f"\u590d\u5236: main.py")
    else:
        print_error(f"\u4e3b\u811a\u672c\u4e0d\u5b58\u5728: {main_src}")
        return None

    # \u521b\u5efascripts\u76ee\u5f55
    scripts_work = build_work / "scripts"
    scripts_work.mkdir(parents=True, exist_ok=True)
    (scripts_work / "__init__.py").touch()

    # \u590d\u5236\u6838\u5fc3\u811a\u672c
    for script in CORE_SCRIPTS:
        src = SCRIPTS_DIR / script
        if src.exists():
            shutil.copy2(src, scripts_work / script)
            print_success(f"\u590d\u5236: scripts/{script}")
        else:
            # \u68c0\u67e5\u662f\u5426\u5728libs\u76ee\u5f55
            src_lib = LIBS_DIR / script
            if src_lib.exists():
                shutil.copy2(src_lib, scripts_work / script)
                print_success(f"\u590d\u5236: libs/{script} -> scripts/")
            elif TOKENLESS_SCRIPTS_DIR.exists():
                # \u68c0\u67e5\u662f\u5426\u5728 TokenLess/scripts \u76ee\u5f55
                src_tokenless = TOKENLESS_SCRIPTS_DIR / script
                if src_tokenless.exists():
                    shutil.copy2(src_tokenless, scripts_work / script)
                    print_success(f"\u590d\u5236: TokenLess/scripts/{script} -> scripts/")
                else:
                    print_warning(f"\u811a\u672c\u4e0d\u5b58\u5728: {script}")
            else:
                print_warning(f"\u811a\u672c\u4e0d\u5b58\u5728: {script}")

    # 复制共享模块（从libs目录，如 browser_automation.py）
    for module in SHARED_MODULES:
        dest = scripts_work / module
        if not dest.exists():
            src = LIBS_DIR / module
            if src.exists():
                shutil.copy2(src, dest)
                print_success(f"复制共享模块: libs/{module} -> scripts/")
            else:
                print_warning(f"共享模块不存在: libs/{module}")

    # \u590d\u5236config\u76ee\u5f55
    if CONFIG_DIR.exists():
        config_work = build_work / "config"
        shutil.copytree(CONFIG_DIR, config_work, dirs_exist_ok=True)
        print_success("\u590d\u5236: config/")

    return build_work

# ==================== \u4ee3\u7801\u6df7\u6dc6 ====================
def obfuscate_scripts(work_dir: Path) -> bool:
    """代码保护：PyArmor混淆（小文件）+ .pyc字节码编译（大文件fallback）"""
    print_step(6, "代码保护（PyArmor混淆 + .pyc编译）")

    # 检查PyArmor是否可用
    has_pyarmor = False
    if BUILD_CONFIG.get("pyarmor", {}).get("enabled", True):
        try:
            import pyarmor
            has_pyarmor = True
            print_success("PyArmor已安装")
        except ImportError:
            print_warning("PyArmor未安装，将仅使用.pyc编译保护")

    # 创建混淆输出目录
    obfuscated_dir = work_dir / "obfuscated"
    if obfuscated_dir.exists():
        shutil.rmtree(obfuscated_dir)
    obfuscated_dir.mkdir(parents=True, exist_ok=True)

    import py_compile

    def try_pyarmor(src_file: Path, out_dir: Path) -> bool:
        """尝试用PyArmor混淆单个文件，成功返回True"""
        if not has_pyarmor:
            return False
        success, _ = run_command([
            sys.executable, "-m", "pyarmor.cli", "gen",
            "-O", str(out_dir),
            str(src_file)
        ], check=False)
        return success and (out_dir / src_file.name).exists()

    def compile_pyc(src_file: Path, out_dir: Path) -> bool:
        """编译为.pyc字节码"""
        try:
            pyc_name = src_file.stem + ".pyc"
            pyc_path = out_dir / pyc_name
            py_compile.compile(str(src_file), cfile=str(pyc_path), doraise=True)
            return pyc_path.exists()
        except Exception as e:
            print_warning(f"  .pyc编译失败 {src_file.name}: {e}")
            return False

    # 1. 处理 main.py（PyInstaller入口必须是.py）
    main_script = work_dir / "main.py"
    if main_script.exists():
        if try_pyarmor(main_script, obfuscated_dir):
            print_success(f"main.py -> PyArmor混淆")
            for rt_dir in obfuscated_dir.glob("pyarmor_runtime_*"):
                print_success(f"  运行时: {rt_dir.name}")
        else:
            shutil.copy2(main_script, obfuscated_dir / "main.py")
            print_warning(f"main.py -> 未混淆（PyInstaller入口需要.py）")

    # 2. 处理 scripts/ 下的所有脚本
    scripts_src = work_dir / "scripts"
    scripts_dst = obfuscated_dir / "scripts"
    scripts_dst.mkdir(parents=True, exist_ok=True)

    if scripts_src.exists():
        obfuscated_count = 0
        pyc_count = 0
        copy_count = 0

        for py_file in sorted(scripts_src.glob("*.py")):
            if py_file.name == "__init__.py":
                shutil.copy2(py_file, scripts_dst / py_file.name)
                copy_count += 1
                continue

            if try_pyarmor(py_file, scripts_dst):
                print_success(f"  {py_file.name} -> PyArmor混淆")
                obfuscated_count += 1
            elif compile_pyc(py_file, scripts_dst):
                print_success(f"  {py_file.name} -> .pyc编译")
                pyc_count += 1
            else:
                shutil.copy2(py_file, scripts_dst / py_file.name)
                print_warning(f"  {py_file.name} -> 未保护（明文）")
                copy_count += 1

        # 复制子目录
        for item in scripts_src.iterdir():
            if item.is_dir():
                dest = scripts_dst / item.name
                if not dest.exists():
                    shutil.copytree(item, dest, dirs_exist_ok=True)

        print(f"\n  保护统计: PyArmor混淆 {obfuscated_count}, .pyc编译 {pyc_count}, 直接复制 {copy_count}")

    return True

# ==================== PyInstaller\u6253\u5305 ====================
def build_with_pyinstaller(work_dir: Path) -> Optional[Path]:
    """使用PyInstaller打包"""
    print_step(7, "PyInstaller打包")

    # 确定入口脚本
    main_script = work_dir / "obfuscated" / "main.py"
    if not main_script.exists():
        main_script = work_dir / "main.py"

    # 创建spec文件
    spec_file = BUILD_DIR / "app.spec"

    # 确定可执行文件名
    exe_name = PROJECT_NAME_EN
    if PLATFORM == "windows":
        exe_name += ".exe"

    # 构建hiddenimports
    hiddenimports = list(INCLUDED_MODULES)
    for script in CORE_SCRIPTS:
        name = script.replace(".py", "")
        if name not in hiddenimports:
            hiddenimports.append(name)

    # 构建excludes
    excludes = list(EXCLUDED_MODULES)

    # 获取配置
    onefile = BUILD_CONFIG.get("pyinstaller", {}).get("onefile", True)
    use_upx = BUILD_CONFIG.get("pyinstaller", {}).get("upx", True)
    show_console = BUILD_CONFIG.get("pyinstaller", {}).get("console", False)

    # spec文件内容 - 根据 onefile 配置生成不同结构
    if onefile:
        # 单文件模式
        spec_content = f"""# -*- mode: python ; coding: utf-8 -*-
# PyInstaller配置文件
# 由 build.py 自动生成
# 模式: onefile (单文件)

import os
import sys
from pathlib import Path

block_cipher = None

# 分析入口脚本
a = Analysis(
    [r'{main_script}'],
    pathex=[str(r'{work_dir}')],
    binaries=[],
    datas=[
        (str(r'{work_dir / "config"}'), 'config'),
        (str(r'{work_dir / "scripts"}'), 'scripts'),
    ],
    hiddenimports={hiddenimports},
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes={excludes},
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# PYZ归档
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# 可执行文件 (单文件模式)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='{exe_name}',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx={use_upx},
    upx_exclude=[],
    runtime_tmpdir=None,
    console={show_console},
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
"""
    else:
        # 目录模式 (onedir)
        spec_content = f"""# -*- mode: python ; coding: utf-8 -*-
# PyInstaller配置文件
# 由 build.py 自动生成
# 模式: onedir (目录)

import os
import sys
from pathlib import Path

block_cipher = None

# 分析入口脚本
a = Analysis(
    [r'{main_script}'],
    pathex=[str(r'{work_dir}')],
    binaries=[],
    datas=[
        (str(r'{work_dir / "config"}'), 'config'),
        (str(r'{work_dir / "scripts"}'), 'scripts'),
    ],
    hiddenimports={hiddenimports},
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes={excludes},
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# PYZ归档
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# 可执行文件 (目录模式)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='{exe_name}',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx={use_upx},
    console={show_console},
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

# 收集所有文件到目录
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx={use_upx},
    upx_exclude=[],
    name='{PROJECT_NAME_EN}',
)
"""

    # 写入spec文件
    with open(spec_file, 'w', encoding='utf-8') as f:
        f.write(spec_content)

    mode_str = "单文件" if onefile else "目录"
    print_success(f"创建spec文件: {spec_file} ({mode_str}模式)")

    # 运行PyInstaller（使用平台隔离的dist和build目录）
    dist_dir = BUILD_DIR / f"dist_{PLATFORM}"
    build_cache_dir = BUILD_DIR / f"build_{PLATFORM}"
    print("  开始打包...")
    success, output = run_command([
        sys.executable, "-m", "PyInstaller",
        "--clean",
        "--noconfirm",
        f"--distpath={dist_dir}",
        f"--workpath={build_cache_dir}",
        str(spec_file)
    ], cwd=BUILD_DIR, check=False, show_output=True)

    if not success:
        print_error("PyInstaller打包失败")
        return None

    # 查找生成的可执行文件

    # 可能的exe位置 (onedir模式会在 dist/PROJECT_NAME_EN/ 下生成)
    possible_paths = [
        dist_dir / PROJECT_NAME_EN / exe_name,  # onedir模式（优先检查）
        dist_dir / exe_name,          # onefile模式
    ]

    for exe_path in possible_paths:
        if exe_path.exists() and exe_path.is_file():
            size_mb = exe_path.stat().st_size / 1024 / 1024
            print_success(f"可执行文件: {exe_path}")
            print(f"  文件大小: {size_mb:.1f} MB")
            
            # 无论什么模式，都复制 scripts 目录到 exe 所在目录
            dist_scripts = exe_path.parent / "scripts"
            dist_scripts.mkdir(exist_ok=True)
            work_scripts = work_dir / "scripts"
            if work_scripts.exists():
                import shutil
                copied = 0
                for item in work_scripts.iterdir():
                    if item.is_file() and item.suffix == '.py':
                        shutil.copy2(item, dist_scripts / item.name)
                        copied += 1
                print_success(f"复制 {copied} 个脚本到 {dist_scripts}")
            else:
                print_warning(f"工作目录 scripts 不存在: {work_scripts}")
            
            return exe_path

    print_error("未找到生成的可执行文件")
    return None

def build_with_nuitka(work_dir: Path) -> Optional[Path]:
    """\u4f7f\u7528Nuitka\u7f16\u8bd1"""
    print_step(7, "Nuitka\u7f16\u8bd1")

    main_script = work_dir / "main.py"

    # \u6784\u5efa\u547d\u4ee4
    cmd = [
        sys.executable, "-m", "nuitka",
        "--standalone",
        "--assume-yes-for-downloads",
        f"--output-filename={PROJECT_NAME_EN}",
    ]

    # Windows\u7279\u5b9a\u914d\u7f6e
    if PLATFORM == "windows":
        # \u5c1d\u8bd1\u6dfb\u52a0\u56fe\u6807
        icon_path = SCRIPTS_DIR / "icon.ico"
        if icon_path.exists():
            cmd.append(f"--windows-icon-from-ico={icon_path}")
        # \u4e0d\u7981\u7528\u63a7\u5236\u53f0\uff0c\u65b9\u4fbf\u8c03\u8bd5

    # macOS\u7279\u5b9a\u914d\u7f6e
    elif PLATFORM == "macos":
        cmd.extend([
            "--macos-create-app-bundle",
            f"--macos-app-name={PROJECT_NAME}",
        ])

    # Linux\u7279\u5b9a\u914d\u7f6e
    else:
        pass

    # \u6dfb\u52a0\u6570\u636e\u76ee\u5f55
    config_dir = work_dir / "config"
    scripts_dir = work_dir / "scripts"

    if config_dir.exists():
        cmd.append(f"--include-data-dir={config_dir}=config")
    if scripts_dir.exists():
        cmd.append(f"--include-data-dir={scripts_dir}=scripts")

    # \u5305\u542b\u5fc5\u8981\u7684\u5305
    cmd.extend([
        "--include-package=playwright",
        "--include-package=requests",
    ])

    # \u6392\u9664\u4e0d\u5fc5\u8981\u7684\u6a21\u5757
    for module in EXCLUDED_MODULES:
        cmd.append(f"--nofollow-import-to={module}")

    # \u4f18\u5316\u9009\u9879
    cmd.extend([
        "--python-flag=no_site",
        "--python-flag=no_warnings",
    ])

    # \u8f93\u51fa\u76ee\u5f55
    cmd.append(f"--output-dir={DIST_DIR}")

    # \u6dfb\u52a0\u4e3b\u811a\u672c
    cmd.append(str(main_script))

    print("  \u7f16\u8bd1\u4e2d\uff0c\u8fd9\u53ef\u80fd\u9700\u898110-30\u5206\u949f...")
    print("  \u8bf7\u8010\u5fc3\u7b49\u5f85...\n")

    success, output = run_command(cmd, cwd=work_dir, check=False, show_output=True)

    if not success:
        print_error("Nuitka\u7f16\u8bd1\u5931\u8d25")
        return None

    # \u67e5\u627e\u751f\u6210\u7684\u53ef\u6267\u884c\u6587\u4ef6
    exe_name = PROJECT_NAME_EN + (".exe" if PLATFORM == "windows" else "")

    # Nuitka\u8f93\u51fa\u76ee\u5f55\u7ed3\u6784
    possible_paths = [
        DIST_DIR / "main_nuitka.dist" / exe_name,
        DIST_DIR / PROJECT_NAME_EN / exe_name,
        DIST_DIR / exe_name,
    ]

    for exe_path in possible_paths:
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / 1024 / 1024
            print_success(f"\u53ef\u6267\u884c\u6587\u4ef6: {exe_path}")
            print(f"  \u6587\u4ef6\u5927\u5c0f: {size_mb:.1f} MB")
            return exe_path

    print_error("\u672a\u627e\u5230\u751f\u6210\u7684\u53ef\u6267\u884c\u6587\u4ef6")
    return None

# ==================== \u53d1\u5e03\u5305\u521b\u5efa ====================
def create_release_package(exe_path: Path, method: str, work_dir: Path = None) -> Optional[Path]:
    """\u521b\u5efa\u53d1\u5e03\u5305"""
    print_step(8, "\u521b\u5efa\u53d1\u5e93\u5305")

    # \u521b\u5efa\u53d1\u5e03\u76ee\u5f55
    release_dir = RELEASE_DIR / f"{PROJECT_NAME_EN}_v{VERSION}_{PLATFORM}_{method}"
    if release_dir.exists():
        shutil.rmtree(release_dir)
    release_dir.mkdir(parents=True, exist_ok=True)

    # \u590d\u5236\u53ef\u6267\u884c\u6587\u4ef6
    exe_name = exe_path.name
    shutil.copy2(exe_path, release_dir / exe_name)
    print_success(f"\u590d\u5236: {exe_name}")

    # onedir模式：复制_internal目录（包含python312.dll等运行时依赖）
    internal_dir = exe_path.parent / "_internal"
    if internal_dir.exists() and internal_dir.is_dir():
        dest_internal = release_dir / "_internal"
        shutil.copytree(internal_dir, dest_internal, dirs_exist_ok=True)
        internal_size = sum(f.stat().st_size for f in dest_internal.rglob('*') if f.is_file()) / 1024 / 1024
        print_success(f"\u590d\u5236: _internal/ ({internal_size:.1f} MB)")

    # \u590d\u5236\u6216\u521b\u5efaconfig\u76ee\u5f55
    config_release = release_dir / "config"
    config_release.mkdir(parents=True, exist_ok=True)

    # \u590d\u5236\u914d\u7f6e\u6587\u4ef6
    if CONFIG_DIR.exists():
        for item in CONFIG_DIR.iterdir():
            if item.is_file():
                shutil.copy2(item, config_release / item.name)
            elif item.is_dir():
                dest_dir = config_release / item.name
                shutil.copytree(item, dest_dir, dirs_exist_ok=True)
        print_success("\u590d\u5236: config/")

    # \u521b\u5efa\u9ed8\u8ba4\u914d\u7f6e\u6587\u4ef6\uff08\u5982\u679c\u4e0d\u5b58\u5728\uff09
    default_config = config_release / "config.json"
    if not default_config.exists():
        config_data = {
            "version": VERSION,
            "user": {"my_id": "", "staff_ids": []},
            "weekly_report": {"target_week": 1, "default_platform": "BS"},
            "browser": {"mode": "headful", "headless": False},
            "logging": {"debug": False, "save_intermediate_files": True}
        }
        with open(default_config, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=2, ensure_ascii=False)
        print_success("\u521b\u5efa: config.json (\u9ed8\u8ba4\u914d\u7f6e)")

    # 强制确保发布版 debug=false（避免源码 debug 设置泄漏到发布包）
    for cfg_path in [default_config, release_dir / "_internal" / "config" / "config.json"]:
        if cfg_path.exists():
            try:
                with open(cfg_path, 'r', encoding='utf-8') as f:
                    cfg_data = json.load(f)
                if cfg_data.get("logging", {}).get("debug") is True:
                    cfg_data["logging"]["debug"] = False
                    with open(cfg_path, 'w', encoding='utf-8') as f:
                        json.dump(cfg_data, f, indent=2, ensure_ascii=False)
                    print_success(f"config debug=false: {cfg_path.relative_to(release_dir)}")
            except Exception:
                pass

    # 复制docs目录（用户手册）
    docs_dir = BASE_DIR / "docs"
    if docs_dir.exists():
        shutil.copytree(docs_dir, release_dir / "docs", dirs_exist_ok=True)
        print_success("复制: docs/")

    # \u521b\u5efa\u5fc5\u8981\u76ee\u5f55
    (release_dir / "chrome_data").mkdir(exist_ok=True)
    (release_dir / "output").mkdir(exist_ok=True)

    # 复制scripts目录（供 main.py 调用子脚本）- 使用保护后的版本
    scripts_release = release_dir / "scripts"
    scripts_release.mkdir(exist_ok=True)

    # 优先从 obfuscated/scripts 复制（已混淆或编译的版本）
    obfuscated_scripts = None
    if work_dir:
        obfuscated_scripts = work_dir / "obfuscated" / "scripts"

    if obfuscated_scripts and obfuscated_scripts.exists():
        scripts_src = obfuscated_scripts
        print(f"  使用保护后的脚本: {scripts_src}")
    elif work_dir and (work_dir / "scripts").exists():
        scripts_src = work_dir / "scripts"
        print(f"  使用构建工作目录的脚本: {scripts_src}")
    else:
        scripts_src = SCRIPTS_DIR

    if scripts_src.exists():
        for item in scripts_src.iterdir():
            if item.is_file() and item.suffix in ('.py', '.pyc', '.pyd'):
                shutil.copy2(item, scripts_release / item.name)
            elif item.is_dir():
                dest = scripts_release / item.name
                if not dest.exists():
                    shutil.copytree(item, dest, dirs_exist_ok=True)
        # 复制 pyarmor_runtime 目录（如果存在）
        for rt_dir in scripts_src.glob("pyarmor_runtime_*"):
            if rt_dir.is_dir():
                dest_rt = scripts_release / rt_dir.name
                if not dest_rt.exists():
                    shutil.copytree(rt_dir, dest_rt, dirs_exist_ok=True)
        py_count = sum(1 for f in scripts_release.glob('*.py'))
        pyc_count = sum(1 for f in scripts_release.glob('*.pyc'))
        print_success(f"复制: scripts/ ({py_count} .py + {pyc_count} .pyc)")

    # 同步保护后的脚本到 _internal/scripts/
    internal_scripts = release_dir / "_internal" / "scripts"
    if internal_scripts.exists():
        for item in scripts_release.iterdir():
            if item.is_file() and item.suffix in ('.py', '.pyc', '.pyd'):
                dest = internal_scripts / item.name
                # 如果有.pyc版本，删除_internal中对应的明文.py
                if item.suffix == '.pyc':
                    plain_py = internal_scripts / (item.stem + '.py')
                    if plain_py.exists():
                        plain_py.unlink()
                shutil.copy2(item, dest)
        for rt_dir in scripts_release.glob("pyarmor_runtime_*"):
            if rt_dir.is_dir():
                dest_rt = internal_scripts / rt_dir.name
                if not dest_rt.exists():
                    shutil.copytree(rt_dir, dest_rt, dirs_exist_ok=True)
        print_success("同步: _internal/scripts/ (保护版本)")

    # 创建 browser_debug 目录（浏览器运行时需要）
    browser_debug_dir = release_dir / "scripts" / "browser_debug"
    browser_debug_dir.mkdir(parents=True, exist_ok=True)
    print_success("\u521b\u5efa: scripts/browser_debug/")

    # macOS 后处理：清除隔离属性 + ad-hoc 签名 + 可执行权限
    if PLATFORM == "macos":
        print("\n  macOS 后处理...")
        # 清除隔离属性（防止 Gatekeeper 拦截）
        run_command(["xattr", "-cr", str(release_dir)], check=False)
        print_success("已清除隔离属性 (xattr -cr)")

        # ad-hoc 代码签名
        success, _ = run_command([
            "codesign", "--force", "--deep", "--sign", "-", str(exe_path)
        ], check=False)
        if success:
            print_success("已完成 ad-hoc 签名 (codesign)")
        else:
            print_warning("ad-hoc 签名失败（不影响本地使用）")

        # 确保可执行权限
        run_command(["chmod", "+x", str(exe_path)], check=False)
        print_success("已设置可执行权限")

    # \u8ba1\u7b97\u5927\u5c0f
    total_size = sum(f.stat().st_size for f in release_dir.rglob('*') if f.is_file()) / 1024 / 1024
    print_success(f"\u53d1\u5e93\u76ee\u5f55: {release_dir}")
    print(f"  \u603b\u5927\u5c0f: {total_size:.1f} MB")

    return release_dir

def create_zip_package(release_dir: Path) -> Optional[Path]:
    """\u521b\u5efaZIP\u538b\u7f29\u5305"""
    print_step(9, "\u521b\u5efaZIP\u538b\u7f29\u5305")

    zip_name = RELEASE_DIR / f"{PROJECT_NAME_EN}_v{VERSION}_{PLATFORM}.zip"

    try:
        with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file in release_dir.rglob('*'):
                if file.is_file():
                    arcname = file.relative_to(release_dir.parent)
                    zipf.write(file, arcname)

        size_mb = zip_name.stat().st_size / 1024 / 1024
        print_success(f"ZIP\u5305: {zip_name}")
        print(f"  \u538b\u7f29\u540e\u5927\u5c0f: {size_mb:.1f} MB")
        return zip_name
    except Exception as e:
        print_error(f"\u521b\u5efaZIP\u5931\u8d25: {e}")
        return None

# ==================== \u6253\u5305\u4e3b\u51fd\u6570 ====================
def build(method: str = "pyinstaller") -> int:
    """\u6253\u5305\u4e3b\u51fd\u6570"""
    print_header(f"{PROJECT_NAME} - \u7edf\u4e00\u6253\u5305\u811a\u672c")
    print(f"  \u7248\u672c: {VERSION}")
    print(f"  \u65f6\u95f4: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  \u5e73\u53f0: {PLATFORM}")
    print(f"  \u6253\u5305\u65b9\u5f0f: {method}")

    # 1. \u68c0\u67e5Python\u7248\u672c
    if not check_python_version():
        return 1

    # 2. \u68c0\u67e5\u76ee\u5f55
    if not check_required_dirs():
        print_warning("\u90e8\u5206\u76ee\u5f55\u7f3a\u5931\uff0c\u5c06\u4f7f\u7528\u9ed8\u8ba4\u7ed3\u6784")

    # 3. \u68c0\u67e5\u6253\u5305\u5de5\u5177
    if method == "pyinstaller":
        if not check_build_tools("pyinstaller"):
            install_dependencies("pyinstaller")
            if not check_build_tools("pyinstaller"):
                return 1
    elif method == "nuitka":
        if not check_build_tools("nuitka"):
            install_dependencies("nuitka")
            if not check_build_tools("nuitka"):
                return 1

    # 4. \u51c6\u5907\u6e90\u7801\u4ee3\u7801
    work_dir = prepare_source()
    if not work_dir:
        return 1

    # 5. \u4ee3\u7801\u6df7\u6dc6\uff08\u4ec5PyInstaller\uff09
    if method == "pyinstaller":
        obfuscate_scripts(work_dir)

    # 6. \u6253\u5305
    if method == "pyinstaller":
        exe_path = build_with_pyinstaller(work_dir)
    else:
        exe_path = build_with_nuitka(work_dir)

    if not exe_path:
        return 1

    # 7. \u521b\u5efa\u53d1\u5e93\u5305
    release_dir = create_release_package(exe_path, method, work_dir)
    if not release_dir:
        return 1

    # 8. \u521b\u5efaZIP
    zip_path = create_zip_package(release_dir)

    # \u603b\u7ed3
    print_header("\u6253\u5305\u5b8c\u6210")
    print(f"""
\u2705 \u6253\u5305\u6210\u529f\uff01

【\u8f93\u51fa\u6587\u4ef6】
  \u53d1\u5e93\u76ee\u5f55: {release_dir}
  ZIP\u538b\u7f29\u5305: {zip_path or 'N/A'}

【\u914d\u7f6e\u4fdd\u62a4】
  - {'PyArmor\u4ee3\u7801\u6df7\u6dc6' if method == 'pyinstaller' else 'Nuitka\u7f16\u8bd1\u4fdd\u62a4'}

【\u4e0b\u4e00\u6b65】
  1. \u6d4b\u8bd5\u8fd0\u884c: \u53cc\u51fb {exe_path.name}
  2. \u9a8c\u8bc1\u529f\u80fd\u6b63\u5e38
  3. \u5206\u53d1ZIP\u6587\u4ef6
""")

    return 0

# ==================== \u547d\u4ee4\u884c\u5165\u53e3 ====================
def main():
    parser = argparse.ArgumentParser(
        description=f"{PROJECT_NAME} - \u7edf\u4e00\u6253\u5305\u811a\u672c",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
\u793a\u4f8b:
  python build.py                     # \u4f7f\u7528PyInstaller\u6253\u5305
  python build.py --method pyinstaller  # \u4f7f\u7528PyInstaller\u6253\u5305
  python build.py --method nuitka    # \u4f7f\u7528Nuitka\u7f16\u8bd1
  python build.py --check             # \u4ec5\u68c0\u67e5\u73af\u5883
  python build.py --export-mac        # \u5bfc\u51fa macOS \u6253\u5305\u6e90\u7801
        """
    )

    parser.add_argument(
        "--method", "-m",
        choices=["pyinstaller", "nuitka"],
        default="pyinstaller",
        help="\u6253\u5305\u65b9\u5f0f: pyinstaller(\u5feb\u901f) \u6216 nuitka(\u9ad8\u4fdd\u62a4)"
    )

    parser.add_argument(
        "--check", "-c",
        action="store_true",
        help="\u4ec5\u68c0\u67e5\u73af\u5883\uff0c\u4e0d\u8fdb\u884c\u6253\u5305"
    )

    parser.add_argument(
        "--export-mac",
        action="store_true",
        help="\u5bfc\u51fa macOS \u6253\u5305\u6240\u9700\u7684\u5e72\u51c0\u6e90\u7801\u76ee\u5f55"
    )

    args = parser.parse_args()

    if args.export_mac:
        return export_mac_package()

    if args.check:
        print_header("\u73af\u5883\u68c0\u67e5")
        check_python_version()
        check_required_dirs()
        check_build_tools(args.method)
        return 0

    return build(args.method)


# ==================== macOS \u5bfc\u51fa ====================
def export_mac_package():
    """导出 macOS 打包所需的干净源码目录"""
    export_name = f"{PROJECT_NAME_EN}_mac_source"
    export_dir = BASE_DIR / export_name

    print_header("导出 macOS 打包源码")
    print(f"  目标目录: {export_dir}")
    print()

    # 清理旧的导出目录
    if export_dir.exists():
        shutil.rmtree(export_dir)
    export_dir.mkdir(parents=True)

    # ---- 1. 根目录文件 ----
    root_files = ["main.py", "config_page.py", "cleanup.py", "requirements.txt"]
    for f in root_files:
        src = BASE_DIR / f
        if src.exists():
            shutil.copy2(src, export_dir / f)
            print(f"  \u2705 {f}")
        else:
            print(f"  \u26a0\ufe0f  {f} 不存在，跳过")

    # ---- 2. scripts/ 目录（排除不需要的） ----
    scripts_exclude_dirs = {"历史版本", "temp", "__pycache__", "browser_debug",
                            "debug_screenshots", "browser_profile", "chrome_data",
                            "WeeklyReport", "llmRsp", "submitted_reports"}
    scripts_exclude_files = {"get-pip.py"}
    skip_exts = {".bat"}

    scripts_dst = export_dir / "scripts"
    scripts_dst.mkdir()

    for item in SCRIPTS_DIR.iterdir():
        if item.name in scripts_exclude_dirs and item.is_dir():
            continue
        if item.name in scripts_exclude_files:
            continue
        if item.suffix in skip_exts:
            continue
        if item.name == "__pycache__":
            continue
        if item.is_file():
            shutil.copy2(item, scripts_dst / item.name)
            print(f"  \u2705 scripts/{item.name}")
        elif item.is_dir():
            shutil.copytree(item, scripts_dst / item.name,
                            ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
            print(f"  \u2705 scripts/{item.name}/")

    # ---- 3. config/ 目录（排除运行时文件） ----
    config_exclude_files = {"browser_state.json"}
    config_exclude_dirs = {"history"}

    def config_ignore(directory, contents):
        ignored = set()
        for c in contents:
            full = Path(directory) / c
            if c in config_exclude_files:
                ignored.add(c)
            elif c in config_exclude_dirs and full.is_dir():
                ignored.add(c)
            elif c == "__pycache__":
                ignored.add(c)
        return ignored

    if CONFIG_DIR.exists():
        shutil.copytree(CONFIG_DIR, export_dir / "config", ignore=config_ignore)
        print(f"  \u2705 config/")

    # ---- 4. docs/ 目录 ----
    docs_dir = BASE_DIR / "docs"
    if docs_dir.exists():
        shutil.copytree(docs_dir, export_dir / "docs",
                        ignore=shutil.ignore_patterns("__pycache__"))
        print(f"  \u2705 docs/")

    # ---- 5. 共享模块（自动从外部目录查找） ----
    # 共享模块及其可能的来源目录（相对于项目根目录的上级）
    shared_modules = {
        "llmchat_final.py": ["TokenLess/scripts", "libs"],
        "llmapi_v10.py": ["TokenLess/scripts", "libs"],
        "browser_automation.py": ["libs", "TokenLess/scripts"],
    }
    # 搜索基准：项目根目录的上级（即 JobsOfAndy/ 或 openclaw/ 等）
    search_bases = [BASE_DIR.parent, BASE_DIR.parent.parent]

    shared_missing = []
    for mod, search_dirs in shared_modules.items():
        if (scripts_dst / mod).exists():
            continue  # 已在 scripts/ 中拷贝过
        found = False
        # 先在项目自己的 scripts/ 目录找
        if (SCRIPTS_DIR / mod).exists():
            shutil.copy2(SCRIPTS_DIR / mod, scripts_dst / mod)
            print(f"  \u2705 共享模块: {mod}")
            found = True
            continue
        # 再到外部目录找
        for base in search_bases:
            if found:
                break
            for rel_dir in search_dirs:
                src = base / rel_dir / mod
                if src.exists():
                    shutil.copy2(src, scripts_dst / mod)
                    print(f"  \u2705 共享模块: {mod} (来自 {src.relative_to(BASE_DIR.parent)})")
                    found = True
                    break
        if not found:
            shared_missing.append(mod)

    if shared_missing:
        print(f"  \u26a0\ufe0f  共享模块未找到: {', '.join(shared_missing)}")
        print(f"     搜索路径: scripts/, {', '.join(str(b) for b in search_bases)}")

    # ---- 6. 创建 ZIP ----
    zip_path = BASE_DIR / f"{export_name}.zip"
    print(f"\n  创建 ZIP: {zip_path.name}")
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(export_dir):
            dirs[:] = [d for d in dirs if d != "__pycache__"]
            for file in files:
                file_path = Path(root) / file
                arcname = Path(export_name) / file_path.relative_to(export_dir)
                zf.write(file_path, arcname)

    # ---- 7. 统计 ----
    file_count = sum(1 for _ in export_dir.rglob("*") if _.is_file())
    dir_size = sum(f.stat().st_size for f in export_dir.rglob("*") if f.is_file())
    zip_size = zip_path.stat().st_size

    print()
    print("=" * 50)
    print(f"  导出完成！")
    print(f"  文件数: {file_count}")
    print(f"  目录大小: {dir_size / 1024 / 1024:.1f} MB")
    print(f"  ZIP 大小: {zip_size / 1024 / 1024:.1f} MB")
    print("=" * 50)
    print()
    print("  使用方法：")
    print(f"  1. 将 {zip_path.name} 拷贝到 Mac")
    print(f"  2. unzip {export_name}.zip")
    print(f"  3. cd {export_name}")
    print(f"  4. chmod +x scripts/build_mac.sh && ./scripts/build_mac.sh")
    print("=" * 50)

    return 0


if __name__ == "__main__":
    sys.exit(main())
