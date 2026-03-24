#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
周报自动化工具 - 主入口
支持：直接运行、PyInstaller打包、Nuitka编译
"""

import os
import sys
import webbrowser
import json
import time
import subprocess
import threading
import platform
import shutil
from pathlib import Path
from datetime import datetime

# ==================== 日志重定向（GUI模式支持）====================
def setup_logging():
    """设置日志输出 - GUI模式下重定向到文件"""
    global LOG_FILE

    # 获取程序目录
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    # 创建日志目录
    log_dir = os.path.join(base_dir, 'logs')
    os.makedirs(log_dir, exist_ok=True)

    # 日志文件路径
    LOG_FILE = os.path.join(log_dir, f'run_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log')

    # 创建日志文件
    log_file = open(LOG_FILE, 'w', encoding='utf-8')

    # 重定向 stdout 和 stderr
    class TeeOutput:
        """同时输出到控制台和文件"""
        def __init__(self, file, stream=None):
            self.file = file
            self.stream = stream
            # 模拟标准流的属性，避免其他模块检查属性时报错
            self.encoding = 'utf-8'
            self.buffer = file  # 提供 buffer 属性

        def write(self, data):
            self.file.write(data)
            self.file.flush()
            if self.stream:
                try:
                    self.stream.write(data)
                except:
                    pass

        def flush(self):
            self.file.flush()
            if self.stream:
                try:
                    self.stream.flush()
                except:
                    pass

    # 在 GUI 模式下，stdout/stderr 可能为 None
    original_stdout = sys.stdout if sys.stdout else None
    original_stderr = sys.stderr if sys.stderr else None

    sys.stdout = TeeOutput(log_file, original_stdout)
    sys.stderr = TeeOutput(log_file, original_stderr)

    if globals().get('DEBUG', False):
        print(f"[日志] 日志文件: {LOG_FILE}")
    return LOG_FILE

# 初始化日志
LOG_FILE = None
if getattr(sys, 'frozen', False):
    # 仅在打包模式下重定向日志
    setup_logging()

# ==================== 浏览器工具函数 ====================
def get_python_exe():
    """获取Python解释器路径"""
    if getattr(sys, 'frozen', False):
        # PyInstaller打包版
        # 直接使用 python 命令，假设在 PATH 中
        # 或者尝试找到系统安装的 Python
        import shutil as shutil2
        python_cmd = shutil2.which('python') or shutil2.which('python3') or 'python'
        return python_cmd
    return sys.executable

def get_chrome_path():
    """获取Chromium内核浏览器路径（Chrome > Edge，跨平台）"""
    system = platform.system()
    if system == "Windows":
        # Chrome
        chrome_paths = [
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "Application", "chrome.exe"),
            os.path.join(os.environ.get("PROGRAMFILES", ""), "Google", "Chrome", "Application", "chrome.exe"),
            os.path.join(os.environ.get("PROGRAMFILES(X86)", ""), "Google", "Chrome", "Application", "chrome.exe"),
        ]
        for p in chrome_paths:
            if os.path.exists(p):
                return p
        if shutil.which("chrome"):
            return shutil.which("chrome")
        # Edge
        edge_paths = [
            os.path.join(os.environ.get("PROGRAMFILES(X86)", ""), "Microsoft", "Edge", "Application", "msedge.exe"),
            os.path.join(os.environ.get("PROGRAMFILES", ""), "Microsoft", "Edge", "Application", "msedge.exe"),
        ]
        for p in edge_paths:
            if os.path.exists(p):
                return p
        return shutil.which("msedge")
    elif system == "Darwin":
        # Chrome
        chrome_paths = [
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            os.path.expanduser("~/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"),
        ]
        for p in chrome_paths:
            if os.path.exists(p):
                return p
        if shutil.which("google-chrome"):
            return shutil.which("google-chrome")
        # Edge
        edge_paths = [
            "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
            os.path.expanduser("~/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge"),
        ]
        for p in edge_paths:
            if os.path.exists(p):
                return p
        return None
    else:
        # Linux: Chrome > Edge > Chromium
        return (shutil.which("google-chrome") or shutil.which("google-chrome-stable")
                or shutil.which("microsoft-edge-stable")
                or shutil.which("chromium") or shutil.which("chromium-browser"))

def open_browser(url):
    """使用Chromium内核浏览器打开配置页面"""
    # GUI模式下 subprocess 需要 CREATE_NO_WINDOW 标志
    creation_flags = 0
    if platform.system() == "Windows":
        creation_flags = getattr(subprocess, 'CREATE_NO_WINDOW', 0x08000000)

    chrome_path = get_chrome_path()

    if chrome_path:
        try:
            # 在已有Chrome窗口中打开新标签页
            subprocess.Popen([chrome_path, url],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                           stdin=subprocess.DEVNULL, creationflags=creation_flags)
            browser_name = "Edge" if "edge" in chrome_path.lower() else "Chrome"
            print(f"  [浏览器] 已启动{browser_name}")
            return True
        except Exception as e:
            print(f"  [浏览器] 启动失败: {e}")
    # 备用：使用系统默认浏览器
    try:
        webbrowser.open(url)
        print(f"  [浏览器] 使用默认浏览器打开")
    except Exception as e:
        print(f"  [浏览器] 默认浏览器打开失败: {e}")

# ==================== 路径配置 ====================
def get_base_dir():
    """获取程序运行目录"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_base_dir()
SCRIPTS_DIR = os.path.join(BASE_DIR, 'scripts')
CONFIG_DIR = os.path.join(BASE_DIR, 'config')
CHROME_DATA_DIR = os.path.join(BASE_DIR, 'chrome_data')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')

# 进度状态文件
PROGRESS_FILE = os.path.join(BASE_DIR, '.progress.json')

# 确保目录存在
for d in [SCRIPTS_DIR, CONFIG_DIR, CHROME_DATA_DIR, OUTPUT_DIR]:
    os.makedirs(d, exist_ok=True)

# 读取调试模式
try:
    import json as _json
    with open(os.path.join(BASE_DIR, "config", "config.json"), "r", encoding="utf-8") as _f:
        _cfg = _json.load(_f)
    DEBUG = _cfg.get("logging", {}).get("debug", False)
except:
    DEBUG = False

# 初始化进度文件
def init_progress():
    """初始化进度文件"""
    data = {
        "status": "idle",
        "current_step": 0,
        "total_steps": 4,
        "step_name": "",
        "message": "等待启动",
        "details": [],
        "elapsed": 0
    }
    try:
        with open(PROGRESS_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except:
        pass

def update_progress(step, step_name, message="", detail=None):
    """更新进度"""
    try:
        with open(PROGRESS_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except:
        data = {"details": [], "elapsed": 0}

    data["current_step"] = step
    data["step_name"] = step_name
    data["message"] = message or step_name
    data["status"] = "running"

    if detail:
        data["details"].append(f"[{step_name}] {detail}")
        # 只保留最近50条
        if len(data["details"]) > 50:
            data["details"] = data["details"][-50:]

    try:
        with open(PROGRESS_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except:
        pass

def complete_progress(message="执行完成"):
    """完成进度"""
    try:
        with open(PROGRESS_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except:
        data = {"details": [], "elapsed": 0}

    data["status"] = "completed"
    data["message"] = message
    data["current_step"] = 4

    try:
        with open(PROGRESS_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except:
        pass

def error_progress(message):
    """错误进度"""
    try:
        with open(PROGRESS_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except:
        data = {"details": [], "elapsed": 0}

    data["status"] = "error"
    data["message"] = message

    try:
        with open(PROGRESS_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except:
        pass

# ==================== 配置Web服务器 ====================
import http.server
import socketserver

class ConfigHandler(http.server.BaseHTTPRequestHandler):
    """配置页面请求处理器"""
    start_triggered = False
    start_time = None
    last_heartbeat = None  # 浏览器最后一次心跳时间

    def log_message(self, *args):
        pass

    def do_GET(self):
        # 每次收到请求都更新心跳时间
        ConfigHandler.last_heartbeat = time.time()

        if self.path == '/':
            self.send_response(200)
            self.send_header('Content-type', 'text/html; charset=utf-8')
            self.end_headers()
            self.wfile.write(CONFIG_HTML.encode('utf-8'))

        elif self.path == '/api/config':
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            cfg_file = os.path.join(CONFIG_DIR, 'config.json')
            if os.path.exists(cfg_file):
                with open(cfg_file, 'r', encoding='utf-8') as f:
                    self.wfile.write(f.read().encode('utf-8'))
            else:
                self.wfile.write(b'{}')

        elif self.path == '/api/progress':
            # 读取进度
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.end_headers()

            try:
                if os.path.exists(PROGRESS_FILE):
                    with open(PROGRESS_FILE, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                else:
                    data = {"status": "idle", "current_step": 0, "total_steps": 4, "step_name": "", "message": "等待启动", "details": [], "elapsed": 0}

                # 计算运行时间
                if ConfigHandler.start_time:
                    data["elapsed"] = int(time.time() - ConfigHandler.start_time)
            except:
                data = {"status": "idle", "message": "读取进度失败"}

            self.wfile.write(json.dumps(data, ensure_ascii=False).encode('utf-8'))

        elif self.path == '/api/start':
            ConfigHandler.start_triggered = True
            ConfigHandler.start_time = time.time()
            # 立即将进度设为running，防止前端轮询到idle就停止
            init_progress()
            update_progress(0, "初始化", "正在启动...")
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            self.wfile.write(b'{"success":true}')

        else:
            # 未知路径（如 /favicon.ico），返回404避免显示Method Not Allowed
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        ConfigHandler.last_heartbeat = time.time()
        if self.path == '/api/save':
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            try:
                new_config = json.loads(body)
                cfg_file = os.path.join(CONFIG_DIR, 'config.json')

                # 先读取现有配置（如果存在）
                existing_config = {}
                if os.path.exists(cfg_file):
                    try:
                        with open(cfg_file, 'r', encoding='utf-8') as f:
                            existing_config = json.load(f)
                    except:
                        pass

                # 递归合并配置
                def merge_config(base, update):
                    """递归合并配置，保留基础配置中存在但更新中没有的字段"""
                    result = base.copy()
                    for key, value in update.items():
                        if key in result and isinstance(result[key], dict) and isinstance(value, dict):
                            result[key] = merge_config(result[key], value)
                        else:
                            result[key] = value
                    return result

                merged_config = merge_config(existing_config, new_config)

                with open(cfg_file, 'w', encoding='utf-8') as f:
                    json.dump(merged_config, f, indent=2, ensure_ascii=False)
                self.send_response(200)
                self.send_header('Content-type', 'application/json')
                self.end_headers()
                self.wfile.write(b'{"success":true}')
            except Exception as e:
                self.send_response(500)
                self.end_headers()

# 配置页面HTML（带实时进度显示）
CONFIG_HTML = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>周报自动化工具 - 配置</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:linear-gradient(135deg,#667eea,#764ba2);min-height:100vh;padding:20px}
.container{max-width:900px;margin:0 auto}
.header{text-align:center;color:white;padding:30px 0}
.header h1{font-size:2em;margin-bottom:10px}
.card{background:white;border-radius:12px;box-shadow:0 10px 40px rgba(0,0,0,.15);margin-bottom:20px;overflow:hidden}
.card-header{background:#f8f9fa;padding:15px 20px;border-bottom:1px solid #dee2e6;font-weight:600}
.card-body{padding:20px}
.agreement-box{max-height:180px;overflow-y:auto;border:1px solid #ddd;padding:15px;margin-bottom:15px;background:#fafafa;font-size:14px;line-height:1.6}
.agreement-box h3{margin:10px 0 5px;color:#333}
.agreement-box ul{margin-left:20px}
.form-group{margin-bottom:15px}
.form-group label{display:block;margin-bottom:5px;font-weight:500;color:#333}
.form-group input,.form-group select,.form-group textarea{width:100%;padding:10px;border:1px solid #ddd;border-radius:6px;font-size:14px}
.form-group textarea{min-height:80px;font-family:monospace}
.form-group .hint{font-size:12px;color:#666;margin-top:4px}
.btn{padding:12px 24px;border:none;border-radius:6px;font-size:14px;font-weight:500;cursor:pointer;transition:all .2s}
.btn-primary{background:linear-gradient(135deg,#667eea,#764ba2);color:white}
.btn-primary:hover{transform:translateY(-2px);box-shadow:0 5px 15px rgba(102,126,234,.4)}
.btn-success{background:#28a745;color:white}
.btn-success:disabled{background:#ccc;cursor:not-allowed}
.actions{display:flex;gap:10px;justify-content:flex-end;padding:15px 20px;border-top:1px solid #eee;background:#fafafa}
.footer{text-align:center;margin-top:20px;color:rgba(255,255,255,.8);font-size:14px}
.section-title{font-size:16px;font-weight:600;color:#333;margin-bottom:10px;padding-bottom:8px;border-bottom:2px solid #667eea}
.progress-card{display:none}
.progress-card.active{display:block}
.progress-header{padding:20px;background:linear-gradient(135deg,#667eea,#764ba2);color:white}
.progress-title{font-size:1.3em;font-weight:600;margin-bottom:10px}
.progress-steps{display:flex;justify-content:space-between;margin:20px 0;position:relative}
.progress-steps::before{content:'';position:absolute;top:15px;left:0;right:0;height:2px;background:#e0e0e0;z-index:0}
.progress-step{position:relative;z-index:1;text-align:center}
.progress-step-icon{width:32px;height:32px;border-radius:50%;background:#e0e0e0;color:#999;display:flex;align-items:center;justify-content:center;margin:0 auto 8px;font-size:14px}
.progress-step.active .progress-step-icon{background:#667eea;color:white}
.progress-step.completed .progress-step-icon{background:#28a745;color:white}
.progress-step-label{font-size:12px;color:#666}
.progress-step.active .progress-step-label{color:#667eea;font-weight:600}
.progress-message{padding:15px;background:#f8f9fa;border-radius:8px;margin:15px 0;font-size:14px;color:#333}
.progress-details{max-height:200px;overflow-y:auto;background:#1e1e1e;color:#0f0;font-family:monospace;font-size:12px;padding:15px;border-radius:8px;margin:15px 0}
.progress-details div{margin:3px 0}
.progress-time{text-align:center;color:#666;font-size:12px;margin-top:10px}
.progress-bar{height:6px;background:#e0e0e0;border-radius:3px;overflow:hidden;margin:15px 0}
.progress-bar-fill{height:100%;background:linear-gradient(90deg,#667eea,#764ba2);transition:width .3s}
</style>
</head>
<body>
<div class="container" id="configPage">
<div class="header"><h1>📋 周报自动化工具</h1><p>配置向导</p></div>

<div class="card">
<div class="card-header">📜 用户协议与隐私政策</div>
<div class="card-body">
<div class="agreement-box">
<h3>用户协议</h3>
<p>欢迎使用周报自动化工具。使用本软件即表示您同意：</p>
<ul><li>本软件仅供授权用户内部测试Demo使用，请勿用于任何实际可能影响生产工作的场景</li>
<li>本程序运行过程会启动浏览器进入OKR和大模型页面且需要您完成登录，其他情况避免操作页面以免影响自动化执行</li>
<li>请注意绿色围墙，注意信息保密！请勿使用任何敏感信息作测试！</li>
<li>因未经全面测试，可能会有兼容性等各种异常问题出现，有任何问题欢迎联系buanan@baidu.com</li></ul>
<h3>隐私政策</h3>
<p>我们不收集任何个人数据，所有配置和数据仅保存在本地。</p>
</div>
<label><input type="checkbox" id="agree" onchange="document.getElementById('startBtn').disabled=!this.checked"> 我已阅读并同意以上条款</label>
</div>
</div>


<div class="card">
<div class="card-header">📖 使用说明 <span style="float:right;cursor:pointer;opacity:0.6" onclick="this.closest('.card').querySelector('.card-body').style.display=this.closest('.card').querySelector('.card-body').style.display==='none'?'block':'none'">[展开/收起]</span></div>
<div class="card-body">
<div style="margin-bottom:15px">
<h4 style="color:#e74c3c;margin-bottom:8px">⚠️ 环境要求（必读）</h4>
<table style="width:100%;border-collapse:collapse;font-size:14px;line-height:1.6">
<tr style="background:#f8f9fa"><th style="padding:8px;border:1px solid #ddd;text-align:left;width:120px">依赖项</th><th style="padding:8px;border:1px solid #ddd;text-align:left">说明</th></tr>
<tr><td style="padding:8px;border:1px solid #ddd"><b>Chrome / Edge</b></td><td style="padding:8px;border:1px solid #ddd">必须安装 <b>Chrome</b> 或 <b>Edge</b> 浏览器（Chromium 内核），Safari/Firefox 不支持。<br>macOS 用户如未安装，程序会自动打开下载页面。安装命令：<code>brew install --cask google-chrome</code></td></tr>
<tr><td style="padding:8px;border:1px solid #ddd"><b>内网环境</b></td><td style="padding:8px;border:1px solid #ddd">需连接百度内网（或 VPN），否则无法访问 OKR 系统。如页面显示"无法访问此网站"，请检查网络/VPN。</td></tr>
<tr><td style="padding:8px;border:1px solid #ddd"><b>OKR 登录</b></td><td style="padding:8px;border:1px solid #ddd">首次运行时，程序会打开 OKR 登录页面（UUAP），请在 <b>5 分钟内</b>完成扫码或账号登录。<br>登录状态保存在 <code>chrome_data/</code>，后续无需重复登录。</td></tr>
<tr><td style="padding:8px;border:1px solid #ddd"><b>大模型登录</b></td><td style="padding:8px;border:1px solid #ddd">使用 web 模式（非 API）时，首次需登录大模型平台（如百度 OneAPI）。<br>使用 API 模式则需提前在下方配置 API Key，无需浏览器登录。</td></tr>
<tr><td style="padding:8px;border:1px solid #ddd"><b>Python</b></td><td style="padding:8px;border:1px solid #ddd">仅 <b>macOS 源码版</b>需要：Python 3.8+（<code>brew install python@3.12</code>）。<br>Windows/Linux 已打包为独立程序，无需安装 Python。</td></tr>
</table>
</div>
<div style="margin-bottom:15px">
<h4 style="color:#667eea;margin-bottom:8px">🚀 快速开始</h4>
<ol style="margin-left:20px;line-height:1.8">
<li><b>检查环境</b> - 确认已安装 Chrome/Edge，已连接内网或 VPN</li>
<li><b>启动程序</b> - 双击 WeeklyReportTool（Windows 为 .exe），浏览器自动打开配置页面</li>
<li><b>配置账号</b> - 填写员工ID和团队成员ID，点击「保存配置」</li>
<li><b>开始执行</b> - 勾选同意条款，点击「启动程序」</li>
<li><b>完成登录</b> - 首次使用时按提示完成 OKR 和大模型平台的登录</li>
</ol>
</div>
<div style="margin-bottom:15px">
<h4 style="color:#667eea;margin-bottom:8px">⚙️ 配置说明</h4>
<ul style="margin-left:20px;line-height:1.8">
<li><b>员工ID</b> - 你的工号（如 s673090）</li>
<li><b>团队成员ID</b> - 需要汇总的成员工号，每行一个</li>
<li><b>目标周</b> - 本周/上周/上上周</li>
<li><b>LLM平台</b> - 大模型优化平台，推荐 BO</li>
</ul>
</div>
<div>
<h4 style="color:#667eea;margin-bottom:8px">❓ 常见问题</h4>
<ul style="margin-left:20px;line-height:1.8">
<li><b>浏览器没打开？</b> 确认已安装 Chrome 或 Edge，或手动访问 http://localhost:8081</li>
<li><b>"无法访问此网站"？</b> 请检查内网连接或 VPN 是否已开启</li>
<li><b>登录失败/超时？</b> 首次使用需在浏览器中手动完成登录，5分钟内完成</li>
<li><b>大模型优化失败？</b> web 模式检查平台登录状态；API 模式检查 API Key 配置</li>
<li><b>执行日志？</b> 查看 logs/ 目录下的日志文件</li>
</ul>
</div>
</div>
</div>

<div class="card">
<div class="card-header">⚙️ 基本配置</div>
<div class="card-body">
<div class="section-title">👤 用户信息</div>
<div class="form-group">
<label>我的员工ID：</label>
<input type="text" id="my_id" placeholder="例如: s673090">
<div class="hint">请输入你的员工ID（查看周报的URL地址中“id=”后的值）</div>
</div>

<div class="form-group">
<label>团队成员ID列表（查看周报的URL地址中“id=”后的值）：</label>
<textarea id="staff_ids" placeholder="s673090
s801573
s123456"></textarea>
<div class="hint">每行一个员工ID，或用英文逗号分隔</div>
</div>

<div class="section-title" style="margin-top:20px">📅 周报设置</div>
<div class="form-group">
<label>目标周：</label>
<select id="target_week">
<option value="0" selected>本周</option>
<option value="1" >上周</option>
<option value="2">上上周</option>
</select>
</div>

<div class="form-group">
<label>调用方式：</label>
<select id="call_type" onchange="onCallTypeChange()">
<option value="api">API 调用（推荐）</option>
<option value="web">Web 网页自动化</option>
</select>
<div class="hint">API模式需配置API Key，体验更好；Web模式通过浏览器自动化操作DeepSeek，无需API Key</div>
</div>

<div class="form-group" id="platform_group">
<label>默认LLM平台：</label>
<select id="default_platform">
<option value="BO">BO - 百度ONEAPI（推荐）</option>
<option value="DS">DS - DeepSeek</option>
</select>
</div>

<div class="form-group" id="apikey_group">
<label>API Key：</label>
<input type="password" id="api_key" placeholder="填写API Key">
</div>
</div>
<div class="actions">
<button class="btn btn-primary" onclick="save()">💾 保存配置</button>
<button class="btn btn-success" id="startBtn" onclick="start()" disabled>🚀 启动程序</button>
</div>
</div>

<div class="footer"><p>周报自动化工具 v1.2.0 © 2026</p></div>
</div>

<!-- 进度页面 -->
<div class="container" id="progressPage" style="display:none">
<div class="header"><h1>📋 周报自动化工具</h1><p>执行进度</p></div>

<div class="card progress-card active">
<div class="progress-header">
<div class="progress-title">🔄 正在执行...</div>
</div>
<div class="card-body">
<div class="progress-steps">
<div class="progress-step" id="step1">
<div class="progress-step-icon">1</div>
<div class="progress-step-label">读取周报</div>
</div>
<div class="progress-step" id="step2">
<div class="progress-step-icon">2</div>
<div class="progress-step-label">整合周报</div>
</div>
<div class="progress-step" id="step3">
<div class="progress-step-icon">3</div>
<div class="progress-step-label">优化周报</div>
</div>
<div class="progress-step" id="step4">
<div class="progress-step-icon">4</div>
<div class="progress-step-label">提交周报</div>
</div>
</div>

<div class="progress-bar">
<div class="progress-bar-fill" id="progressBar" style="width:0%"></div>
</div>

<div class="progress-message" id="progressMessage">准备启动...</div>

<div class="progress-details" id="progressDetails">
<div>等待开始执行...</div>
</div>

<div class="progress-time" id="progressTime">已运行: 0 秒</div>
</div>
</div>

<div class="footer"><p>周报自动化工具 v1.2.0 © 2026</p></div>
</div>

<script>
let pollInterval = null;

function onCallTypeChange(){
  const callType=document.getElementById('call_type').value;
  const pg=document.getElementById('platform_group');
  const ag=document.getElementById('apikey_group');
  if(callType==='web'){
    pg.innerHTML='<label>默认LLM平台：</label><input type="text" value="DS - DeepSeek" disabled style="background:#f0f0f0;padding:10px;border:1px solid #ddd;border-radius:6px;width:100%;font-size:14px"><div class="hint">Web模式仅支持DeepSeek</div>';
    ag.style.display='none';
  } else {
    pg.innerHTML='<label>默认LLM平台：</label><select id="default_platform" style="width:100%;padding:10px;border:1px solid #ddd;border-radius:6px;font-size:14px"><option value="BO">BO - 百度ONEAPI（推荐）</option><option value="DS">DS - DeepSeek</option></select>';
    ag.style.display='';
  }
}

function load(){
  fetch('/api/config').then(r=>r.json()).then(c=>{
    document.getElementById('my_id').value=c.user?.my_id||'';
    if(c.user?.staff_ids && c.user.staff_ids.length>0){
      document.getElementById('staff_ids').value=c.user.staff_ids.join('\\n');
    }
    document.getElementById('target_week').value=c.weekly_report?.target_week??0;
    // 读取平台和调用方式
    const platform = c.weekly_report?.default_platform||'BO';
    const platformConfig = c.llm_platforms?.[platform]||{};
    const callType = platformConfig.type||'api';
    document.getElementById('call_type').value=callType;
    onCallTypeChange();
    if(callType==='api'){
      document.getElementById('default_platform').value=platform;
    }
    document.getElementById('api_key').value=platformConfig.api_config?.api_key||c.llm_chat?.api_key||'';
  });
}

function save(){
  let staffIdsText=document.getElementById('staff_ids').value.trim();
  let staffIds=[];
  if(staffIdsText){
    staffIds=staffIdsText.split(/[\\n,]+/).map(s=>s.trim()).filter(s=>s.length>0);
  }

  let myId=document.getElementById('my_id').value.trim();

  if(!myId){
    alert('请输入你的员工ID');
    return false;
  }

  if(staffIds.length===0){
    alert('请输入至少一个团队成员ID');
    return false;
  }

  if(!staffIds.includes(myId)){
    staffIds.unshift(myId);
  }

  fetch('/api/save',{
    method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify((() => {
      const callType=document.getElementById('call_type').value;
      const platform=callType==='web'?'DS':(document.getElementById('default_platform')?.value||'BO');
      const apiKey=callType==='api'?document.getElementById('api_key').value.trim():'';
      const platformUpdate={};
      platformUpdate[platform]={type:callType,api_config:{api_key:apiKey}};
      return {
        version:'1.2.0',
        user:{my_id:myId,staff_ids:staffIds},
        weekly_report:{
          target_week:parseInt(document.getElementById('target_week').value),
          default_platform:platform
        },
        llm_platforms:platformUpdate,
        llm_chat:{api_key:apiKey},
        logging:{debug:false}
      };
    })())
  }).then(r=>r.json()).then(d=>{
    if(d.success){
      alert('配置已保存！团队成员: '+staffIds.length+' 人');
    }
  });
  return true;
}

function start(){
  if(!document.getElementById('agree').checked){
    alert('请先阅读并同意用户协议');
    return;
  }
  if(!save()) return;

  // 切换到进度页面
  document.getElementById('configPage').style.display='none';
  document.getElementById('progressPage').style.display='block';

  // 启动程序，等服务端确认后再开始轮询
  fetch('/api/start').then(()=>{
    setTimeout(pollProgress, 1000);
  }).catch(()=>{
    setTimeout(pollProgress, 2000);
  });
}

function pollProgress(){
  fetch('/api/progress')
    .then(r=>r.json())
    .then(data=>{
      updateProgress(data);

      if(data.status === 'completed'){
        document.querySelector('.progress-title').innerHTML = '✅ 执行完成！';
        document.querySelector('.progress-header').style.background = 'linear-gradient(135deg,#28a745,#20c997)';
      } else if(data.status === 'error'){
        document.querySelector('.progress-title').innerHTML = '❌ 执行出错';
        document.querySelector('.progress-header').style.background = 'linear-gradient(135deg,#dc3545,#c82333)';
      } else {
        // running 或 idle 都继续轮询
        pollInterval = setTimeout(pollProgress, 1500);
      }
    })
    .catch(err=>{
      console.error(err);
      pollInterval = setTimeout(pollProgress, 2000);
    });
}

function updateProgress(data){
  const steps = ['step1','step2','step3','step4'];
  steps.forEach((id,i)=>{
    const el = document.getElementById(id);
    if(i < data.current_step){
      el.className = 'progress-step completed';
      el.querySelector('.progress-step-icon').innerHTML = '✓';
    } else if(i === data.current_step){
      el.className = 'progress-step active';
      el.querySelector('.progress-step-icon').innerHTML = (i+1).toString();
    } else {
      el.className = 'progress-step';
      el.querySelector('.progress-step-icon').innerHTML = (i+1).toString();
    }
  });

  const percent = Math.round((data.current_step / data.total_steps) * 100);
  document.getElementById('progressBar').style.width = percent + '%';
  document.getElementById('progressMessage').innerHTML = data.message || data.step_name || '处理中...';

  const detailsEl = document.getElementById('progressDetails');
  if(data.details && data.details.length > 0){
    detailsEl.innerHTML = data.details.map(d=>'<div>'+escapeHtml(d)+'</div>').join('');
    detailsEl.scrollTop = detailsEl.scrollHeight;
  }

  if(data.elapsed){
    const mins = Math.floor(data.elapsed / 60);
    const secs = data.elapsed % 60;
    document.getElementById('progressTime').innerHTML = '已运行: ' + mins + '分' + secs + '秒';
  }
}

function escapeHtml(text){
  return text.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

window.onload=function(){
  load();
  // 心跳：每15秒ping一次服务器，让服务器知道浏览器还活着
  setInterval(()=>{fetch('/api/progress').catch(()=>{})},15000);
  // 检查是否正在执行中（支持页面刷新后恢复进度显示）
  fetch('/api/progress').then(r=>r.json()).then(data=>{
    if(data.status === 'running' || data.status === 'completed' || data.status === 'error'){
      document.getElementById('configPage').style.display='none';
      document.getElementById('progressPage').style.display='block';
      if(data.status === 'running'){
        pollProgress();
      } else {
        updateProgress(data);
        if(data.status === 'completed'){
          document.querySelector('.progress-title').innerHTML = '✅ 执行完成！';
          document.querySelector('.progress-header').style.background = 'linear-gradient(135deg,#28a745,#20c997)';
        } else if(data.status === 'error'){
          document.querySelector('.progress-title').innerHTML = '❌ 执行出错';
          document.querySelector('.progress-header').style.background = 'linear-gradient(135deg,#dc3545,#c82333)';
        }
      }
    }
  }).catch(()=>{});
};
</script>
</body>
</html>
"""

def run_config_server():
    """运行配置Web服务器，用户点击启动后转为后台线程继续服务进度轮询"""
    import socket
    import atexit
    import signal

    PID_FILE = os.path.join(BASE_DIR, '.server.pid')

    def _cleanup_stale_process():
        """启动时清理上一次残留的进程"""
        if not os.path.exists(PID_FILE):
            return
        try:
            with open(PID_FILE, 'r') as f:
                old_pid = int(f.read().strip())
            if platform.system() == "Windows":
                result = subprocess.run(
                    f'tasklist /FI "PID eq {old_pid}" /NH',
                    shell=True, capture_output=True, text=True, timeout=5
                )
                if str(old_pid) in result.stdout:
                    subprocess.run(f'taskkill /PID {old_pid} /F', shell=True, timeout=5)
            else:
                try:
                    os.kill(old_pid, signal.SIGTERM)
                except ProcessLookupError:
                    pass
        except (ValueError, FileNotFoundError, subprocess.TimeoutExpired):
            pass
        try:
            os.remove(PID_FILE)
        except:
            pass

    def _register_pid():
        """写入当前PID，注册退出清理"""
        try:
            with open(PID_FILE, 'w') as f:
                f.write(str(os.getpid()))
            atexit.register(lambda: os.remove(PID_FILE) if os.path.exists(PID_FILE) else None)
        except:
            pass

    def find_free_port(start_port=8080, max_attempts=50):
        for port in range(start_port, start_port + max_attempts):
            try:
                with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                    s.bind(('', port))
                    return port
            except OSError:
                continue
        # 所有端口都失败，让OS自动分配
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.bind(('', 0))
            return s.getsockname()[1]

    _cleanup_stale_process()
    _register_pid()

    port = find_free_port()
    socketserver.TCPServer.allow_reuse_address = False
    server = socketserver.TCPServer(("", port), ConfigHandler)
    print(f"  配置文件: http://localhost:{port}")

    # 先在后台线程处理请求，确保服务器就绪后再打开浏览器
    HEARTBEAT_TIMEOUT = 60  # 浏览器无心跳超时（秒）

    def serve_until_start():
        server.timeout = 5  # handle_request 每5秒超时一次，以便检查心跳
        while not ConfigHandler.start_triggered:
            server.handle_request()
            # 检查浏览器心跳：如果已收到过请求但超过HEARTBEAT_TIMEOUT秒没有新请求，说明浏览器已关闭
            if ConfigHandler.last_heartbeat and (time.time() - ConfigHandler.last_heartbeat > HEARTBEAT_TIMEOUT):
                print(f"  浏览器已关闭（{HEARTBEAT_TIMEOUT}秒无响应），程序退出")
                os._exit(0)

    srv_thread = threading.Thread(target=serve_until_start, daemon=True)
    srv_thread.start()

    print(f"  正在打开浏览器...")
    open_browser(f"http://localhost:{port}")
    ConfigHandler.last_heartbeat = time.time()  # 打开浏览器时初始化心跳
    print(f"  等待用户操作...")

    # 等待用户点击启动
    srv_thread.join()

    # 用户已点击启动，将服务器转为后台线程继续运行（支持进度轮询）
    def serve_background():
        try:
            server.serve_forever()
        except:
            pass

    bg_thread = threading.Thread(target=serve_background, daemon=True)
    bg_thread.start()
    print("  配置服务器已转入后台（支持进度轮询）")


class _ProgressOutputStream:
    """包装输出流，将print输出同步到进度文件的details数组，供前端实时展示"""
    def __init__(self, original, progress_file):
        self._original = original
        self._progress_file = progress_file
        self._line_buf = ""
        self.encoding = getattr(original, 'encoding', 'utf-8') if original else 'utf-8'
        self.buffer = getattr(original, 'buffer', None) if original else None
        self.closed = False

    def write(self, data):
        if self._original:
            try:
                self._original.write(data)
            except:
                pass
        self._line_buf += str(data)
        while '\n' in self._line_buf:
            line, self._line_buf = self._line_buf.split('\n', 1)
            line = line.strip()
            if line:
                self._append_to_progress(line)

    def flush(self):
        if self._original:
            try:
                self._original.flush()
            except:
                pass

    def _append_to_progress(self, line):
        """追加一行到进度文件"""
        try:
            if not os.path.exists(self._progress_file):
                return
            with open(self._progress_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            details = data.get("details", [])
            details.append(line)
            if len(details) > 200:
                details = details[-200:]
            data["details"] = details
            data["message"] = line  # 实时更新消息，让前端显示最新进度
            with open(self._progress_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except:
            pass


def run_main_script():
    """运行主脚本 - 直接导入调用，避免 subprocess 依赖"""
    import asyncio

    print("  启动主程序...")
    if DEBUG:
        print(f"  工作目录: {BASE_DIR}")
        print(f"  脚本目录: {SCRIPTS_DIR}")

    # 安装进度输出捕获，将所有print输出同步到进度页面
    sys.stdout = _ProgressOutputStream(sys.stdout, PROGRESS_FILE)
    sys.stderr = _ProgressOutputStream(sys.stderr, PROGRESS_FILE)

    try:
        # 打包环境下：将外部config同步到_internal/config，确保脚本读到用户最新配置
        if getattr(sys, 'frozen', False):
            internal_config = os.path.join(BASE_DIR, '_internal', 'config', 'config.json')
            external_config = os.path.join(CONFIG_DIR, 'config.json')
            if os.path.exists(external_config) and os.path.exists(os.path.dirname(internal_config)):
                shutil.copy2(external_config, internal_config)
                if DEBUG:
                    print(f"  已同步配置: config/config.json -> _internal/config/config.json")

        # 直接导入并运行 weeklyreport_auto 模块
        # 添加 scripts 目录到路径
        if SCRIPTS_DIR not in sys.path:
            sys.path.insert(0, SCRIPTS_DIR)

        print("  导入 weeklyreport_auto 模块...")
        import weeklyreport_auto

        print("  开始执行主函数...")
        update_progress(1, "读取周报", "正在执行...")

        # 运行异步主函数
        success = asyncio.run(weeklyreport_auto.main())

        if success:
            complete_progress("执行完成")
        else:
            error_progress("执行失败")

    except ImportError as e:
        print(f"  导入错误: {e}")
        error_progress(f"导入模块失败: {e}")
    except Exception as e:
        print(f"  运行错误: {e}")
        import traceback
        traceback.print_exc()
        error_progress(str(e))


def main():
    """主入口"""
    # 首先清理历史进度文件
    init_progress()

    print("=" * 60)
    print("  周报自动化工具 v1.1.0")
    print("=" * 60)

    # 运行配置服务器
    run_config_server()

    # 运行主程序
    run_main_script()

    print("\n" + "=" * 60)
    print("  执行完成")
    print("=" * 60)

    # 执行完成后等待一段时间让用户查看进度页面，然后自动退出
    # 心跳检测：浏览器关闭后60秒自动退出
    print("  关闭浏览器页面后程序将自动退出...")
    while True:
        if ConfigHandler.last_heartbeat and (time.time() - ConfigHandler.last_heartbeat > 60):
            print("  浏览器已关闭，程序退出")
            break
        time.sleep(5)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        # 确保任何崩溃都写入日志
        try:
            import traceback
            err_msg = f"[致命错误] {e}\n{traceback.format_exc()}"
            print(err_msg)
            # 如果日志文件存在，确保写入
            if LOG_FILE:
                with open(LOG_FILE, 'a', encoding='utf-8') as f:
                    f.write(err_msg + "\n")
        except:
            pass
        # GUI模式下保持窗口以便查看错误
        try:
            if sys.stdin and sys.stdin.isatty():
                input("发生错误，按回车键退出...")
        except:
            pass
