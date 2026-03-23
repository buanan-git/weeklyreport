#!/usr/bin/env python3
"""
通用表单自动填写工具 - 最终交互版
支持：DeepSeek、豆包、Kimi
功能：
1. 浏览器复用
2. 定位输入框时显示序号和高亮
3. 自动完成输入和提交（无需用户确认）
4. 智能检测页面加载完成（优化版）
5. 检测并关闭新增侧边栏
6. 逐个查找复制按钮，悬浮验证提示，用户交互确认
7. 保存剪贴板内容到文件
8. 支持文件上传（DeepSeek平台）- 增强版：精确定位输入框内回形针图标
"""
import traceback
import sys
import os
import io
import builtins
# ==================== 1. 先设置编码 ====================
os.environ["PYTHONUTF8"] = "1"
if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# ==================== 2. 打印启动信息 ====================
print("=== 脚本启动调试 ===", file=sys.stderr)
print(f"Python版本: {sys.version}", file=sys.stderr)
print(f"当前工作目录: {os.getcwd()}", file=sys.stderr)
print(f"命令行参数: {sys.argv}", file=sys.stderr)
print(f"环境变量 PYTHONUTF8: {os.environ.get('PYTHONUTF8', '未设置')}", file=sys.stderr)

# ==================== 3. 导入所有必要模块 ====================
try:
    import asyncio
    import json
    import argparse
    import time
    import platform
    import subprocess
    import socket
    import random
    import re
    import pyperclip
    from datetime import datetime  # ⭐ 提前导入datetime
    from urllib.parse import urlparse
    from playwright.async_api import async_playwright
    import warnings
    from pathlib import Path  # ⭐ 添加这行！
    from typing import Dict, Any, Optional, List  # ⭐ 添加这行（可选，但推荐）
    warnings.filterwarnings("ignore", category=ResourceWarning)
except Exception as e:
    print(f"导入失败: {e}", file=sys.stderr)
    traceback.print_exc(file=sys.stderr)
    sys.exit(1)

# ==================== 4. 现在可以定义日志函数 ====================
def debug_log(message):
    """调试日志"""
    with open(r"D:\OC_FATAL_ERROR.log", "a", encoding="utf-8") as f:
        f.write(f"{datetime.now()} - {message}\n")

def emergency_log(msg):
    """紧急日志"""
    with open(r"D:\OC_DEBUG.log", "a", encoding="utf-8") as f:
        f.write(f"{datetime.now()}: {msg}\n")

def fatal_debug(msg):
    """致命错误日志"""
    with open(r"D:\FATAL_DEBUG.log", "a", encoding="utf-8") as f:
        f.write(f"{datetime.now()} - {msg}\n")

def global_excepthook(exc_type, exc_value, exc_traceback):
    """全局异常钩子"""
    err_info = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    fatal_debug("检测到未捕获异常:\n" + err_info)

sys.excepthook = global_excepthook

# ==================== 5. 导入配置加载器 ====================
from config_loader import ConfigLoader

# 加载配置
config = ConfigLoader("llmchat")

# ==================== 6. 配置区域 ====================
DEBUG = config.debug
BASE_DIR = config.script_dir
BROWSER_STATE_FILE = config.get_browser_state_file()
BROWSER_DEBUG_DIR = config.get_path("browser_debug_dir")
LLMRSP_DIR = config.get_path("llm_response_dir")
SCREENSHOT_DIR = config.script_dir / "debug_screenshots"

os.makedirs(BROWSER_DEBUG_DIR, exist_ok=True)
os.makedirs(LLMRSP_DIR, exist_ok=True)
os.makedirs(SCREENSHOT_DIR, exist_ok=True)

# ==================== 7. Python 3.14+ 兼容性修复 ====================
PY_VERSION = sys.version_info
IS_PY314_OR_HIGHER = PY_VERSION.major == 3 and PY_VERSION.minor >= 14

if IS_PY314_OR_HIGHER:
    try:
        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
    except:
        pass
    
    print(f"[警告] 您正在使用 Python {PY_VERSION.major}.{PY_VERSION.minor}，这是测试版本")
    print("[警告] Playwright 可能不完全兼容，建议使用 Python 3.11 或 3.12")

# ==================== 8. 平台配置 ====================
platform_ds = config.get_platform_config("DS")
platform_db = config.get_platform_config("DB")
platform_km = config.get_platform_config("KM")

DS = platform_ds.get("url", "https://chat.deepseek.com") if platform_ds else "https://chat.deepseek.com"
DB = platform_db.get("url", "https://www.doubao.com/chat") if platform_db else "https://www.doubao.com/chat"
KM = platform_km.get("url", "https://www.kimi.com") if platform_km else "https://www.kimi.com"
# 打印调试信息
print(f"[配置] DeepSeek URL: {DS}")
print(f"[配置] 豆包 URL: {DB}")
print(f"[配置] Kimi URL: {KM}")

PLATFORM_NAMES = {
    'deepseek.com': 'DeepSeek',
    'doubao.com': '豆包',
    'kimi.com': 'Kimi',
    'moonshot.cn': 'Kimi'
}

# 文本大小阈值
TEXT_SIZE_THRESHOLD = 500
FILE_UPLOAD_THRESHOLD = 2048

# OpenClaw提示词常量
CLAW2DS_PROMPT_FILE = '''附件文件是openclaw的请求报文，请不要回显内容，而是根据你和openclaw之间的通信协议解析处理这个请求报文，并且按照协议规范回复完整报文并回显在页面。页面显示的内容应该按照如下完整json格式显示
{
  "content": "这是自然语言回复...",
  "tool_calls": [
    {
      "id": "call_12345",
      "type": "function",
      "function": {
        "name": "exec",
        "arguments": "{\\"command\\": \\"python3 --version\\"}"
      }
    }
  ]
}'''

CLAW2DS_PROMPT_DATA = '''以下是openclaw的请求报文，请不要回显内容，而是根据你和openclaw之间的通信协议解析处理这个请求报文，并且按照协议规范回复完整报文并回显在页面。页面显示的内容应该按照如下完整json格式显示
{
  "content": "这是自然语言回复...",
  "tool_calls": [
    {
      "id": "call_12345",
      "type": "function",
      "function": {
        "name": "exec",
        "arguments": "{\\"command\\": \\"python3 --version\\"}"
      }
    }
  ]
}'''

# ==================== 9. 打印函数 ====================
def dprint(*args, force=False, **kwargs):
    """调试打印函数"""
    if force or DEBUG:
        timestamp = datetime.now().strftime("%H:%M:%S")
        print(f"[{timestamp}]", *args, **kwargs)

def iprint(*args, **kwargs):
    """智能打印函数 - 兼容各种输出场景"""
    if sys.stdout is None or sys.stdout.closed:
        return
    
    msg = ' '.join(str(arg) for arg in args)
    
    # 替换Unicode符号
    replacements = {
        '✓': '[OK]', '✗': '[ERR]', '●': '-', '▶': '>',
        '「': '"', '」': '"', '：': ':', '；': ';', '！': '!', '？': '?',
        '\u2705': '[OK]', '\u26a0': '[WARN]', '\ufe0f': '',
    }
    for old, new in replacements.items():
        msg = msg.replace(old, new)
    
    kwargs['flush'] = True
    
    # 使用原始print
    original_print = getattr(builtins, '_original_print', print)
    
    try:
        original_print(msg, **kwargs)
        return
    except (UnicodeEncodeError, ValueError, BrokenPipeError):
        pass
    
    # 备用方案：尝试二进制写入
    try:
        if hasattr(sys.stdout, 'buffer') and sys.stdout.buffer and not sys.stdout.buffer.closed:
            sys.stdout.buffer.write(msg.encode('utf-8', errors='replace'))
            sys.stdout.buffer.write(b'\n')
            sys.stdout.buffer.flush()
            return
    except:
        pass
    
    # 最后尝试：强制ASCII
    try:
        safe_msg = msg.encode('ascii', errors='replace').decode('ascii')
        original_print(safe_msg, flush=True)
    except:
        pass

# 保存原始print
builtins._original_print = builtins.print
builtins.print = iprint

# ==================== 10. 安全输出函数 ====================
def safe_output(message):
    """安全输出函数"""
    try:
        if hasattr(sys.stdout, 'buffer') and sys.stdout.buffer is not None:
            sys.stdout.buffer.write(message.encode('utf-8'))
            sys.stdout.buffer.flush()
            return
    except:
        pass
    
    try:
        print(message, end='', flush=True)
    except:
        try:
            sys.stdout.write(message)
            sys.stdout.flush()
        except:
            pass

# ==================== 11. 事件循环清理 ====================
def cleanup_asyncio():
    """清理 asyncio 资源"""
    try:
        loop = asyncio.get_event_loop()
        if loop.is_running():
            return
        if hasattr(loop, '_proactor'):
            try:
                loop._proactor.close()
            except:
                pass
        if not loop.is_closed():
            loop.close()
    except:
        pass

# ==================== 工具函数 ====================
 
def get_url_domain(url):
    """从URL中提取域名"""
    if not url:
        return ""
    
    parsed = urlparse(url)
    domain = parsed.netloc.lower()
    
    # 确保 domain 是字符串
    if isinstance(domain, bytes):
        domain = domain.decode('utf-8')
    
    if domain.startswith('www.'):
        domain = domain[4:]
    return domain

def get_platform_name(url):
    """根据URL获取平台中文名称"""
    domain = get_url_domain(url)
    
    # 优先匹配常量
    if url == DS or 'deepseek.com' in domain:
        return 'DeepSeek'
    elif url == DB or 'doubao.com' in domain:
        return '豆包'
    elif url == KM or 'kimi.com' in domain or 'moonshot.cn' in domain:
        return 'Kimi'
    
    # 原有的映射检查
    for key, name in PLATFORM_NAMES.items():
        if key in domain:
            return name
    return "未知平台"

def get_platform_handler(page, url):
    """根据URL返回对应的平台处理器 - 增强调试版"""
    domain = get_url_domain(url)
    
    # 打印调试信息
    print(f"\n[平台识别调试] URL: {url}")
    print(f"[平台识别调试] 域名: {domain}")
    print(f"[平台识别调试] DS常量: {DS}")
    print(f"[平台识别调试] DB常量: {DB}")
    print(f"[平台识别调试] KM常量: {KM}")
    
    if url == DS or 'deepseek.com' in domain:
        iprint(f"\n[平台识别] DeepSeek - {url}")
        return DeepSeekHandler(page)
    elif url == DB or 'doubao.com' in domain:
        iprint(f"\n[平台识别] 豆包 - {url}")
        return DoubaoHandler(page)
    elif url == KM or 'kimi.com' in domain or 'moonshot.cn' in domain:
        iprint(f"\n[平台识别] Kimi - {url}")
        return KimiHandler(page)
    else:
        iprint(f"\n[平台识别] 未知平台，使用通用模式 - {url}")
        return DeepSeekHandler(page)

def sanitize_filename(text, max_bytes=50):
    """
    清理文件名，移除非法字符，并确保字节长度不超过限制
    
    Args:
        text: 原始文本
        max_bytes: 最大字节长度（默认50字节，留有余地）
    
    Returns:
        str: 安全的文件名
    """
    # 1. 移除非法字符
    text = re.sub(r'[<>:"/\\|?*]', '_', text)
    
    # 2. 如果为空，返回默认值
    if not text:
        return "unnamed"
    
    # 3. 按字节长度截断
    encoded = text.encode('utf-8')
    if len(encoded) > max_bytes:
        # 从开头逐字符截取，直到不超过字节限制
        result = ""
        for char in text:
            test = (result + char).encode('utf-8')
            if len(test) <= max_bytes:
                result += char
            else:
                break
        text = result
    
    return text

async def analyze_page(page, context=""):
    if not DEBUG:
        return
    dprint(f"\n[页面分析] {context}")
    try:
        title = await page.title()
        dprint(f"  标题: {title}")
        url = page.url
        dprint(f"  URL: {url}")
    except:
        pass

# ==================== 高亮显示工具 ====================

async def highlight_element_with_index(page, element, index, color='blue', duration=2):
    """高亮显示元素并显示序号"""
    if not element:
        return
    
    try:
        box = await element.bounding_box()
        if box:
            dprint(f"    元素 [{index}] 位置: ({box['x']:.0f}, {box['y']:.0f}) 大小: {box['width']:.0f}x{box['height']:.0f}")
        
        # 添加高亮和序号
        await element.evaluate(f"""
            (el, idx) => {{
                // 保存原始样式
                if (!el._originalStyles) {{
                    el._originalStyles = {{
                        outline: el.style.outline,
                        boxShadow: el.style.boxShadow,
                        position: el.style.position,
                        zIndex: el.style.zIndex
                    }};
                }}
                
                // 添加高亮
                el.style.outline = '3px solid {color}';
                el.style.boxShadow = '0 0 15px {color}';
                el.style.position = 'relative';
                el.style.zIndex = '9999';
                
                // 移除旧的序号标签
                const oldLabel = document.getElementById('highlight-label-' + idx);
                if (oldLabel) oldLabel.remove();
                
                // 添加新的序号标签
                const label = document.createElement('div');
                label.id = 'highlight-label-' + idx;
                label.textContent = '[' + idx + ']';
                label.style.position = 'absolute';
                label.style.top = '-25px';
                label.style.left = '0';
                label.style.backgroundColor = '{color}';
                label.style.color = 'white';
                label.style.padding = '2px 8px';
                label.style.borderRadius = '4px';
                label.style.fontSize = '14px';
                label.style.fontWeight = 'bold';
                label.style.zIndex = '10000';
                label.style.boxShadow = '0 2px 5px rgba(0,0,0,0.3)';
                el.appendChild(label);
                
                el.scrollIntoView({{ behavior: 'smooth', block: 'center' }});
            }}
        """, index)
        
        await asyncio.sleep(duration)
        
        # 移除高亮但保留序号（让用户能看到序号）
        await element.evaluate("""
            el => {
                if (el._originalStyles) {
                    el.style.outline = el._originalStyles.outline;
                    el.style.boxShadow = el._originalStyles.boxShadow;
                } else {
                    el.style.outline = '';
                    el.style.boxShadow = '';
                }
            }
        """)
        
    except Exception as e:
        dprint(f"    高亮失败: {e}")

async def clear_all_highlights(page):
    """清除所有高亮标签"""
    try:
        await page.evaluate("""
            () => {
                const labels = document.querySelectorAll('[id^="highlight-label-"]');
                labels.forEach(label => label.remove());
            }
        """)
    except:
        pass

async def print_element_info(element, name="元素"):
    if not DEBUG:
        return
    
    try:
        info = await element.evaluate("""el => {
            const rect = el.getBoundingClientRect();
            return {
                tag: el.tagName,
                id: el.id,
                class: el.className,
                type: el.getAttribute('type'),
                name: el.getAttribute('name'),
                value: el.value,
                text: el.textContent?.trim().substring(0, 100),
                title: el.getAttribute('title'),
                aria_label: el.getAttribute('aria-label'),
                placeholder: el.getAttribute('placeholder'),
                isContentEditable: el.isContentEditable,
                rect: {
                    top: rect.top,
                    left: rect.left,
                    width: rect.width,
                    height: rect.height
                }
            };
        }""")
        
        dprint(f"\n  === {name} 详细信息 ===")
        for key, value in info.items():
            if key != 'rect' and value:
                dprint(f"    {key}: {value}")
        dprint(f"    位置: ({info['rect']['left']:.0f}, {info['rect']['top']:.0f})")
    except Exception as e:
        dprint(f"    获取元素信息失败: {e}")

# ==================== 侧边栏检测和关闭 ====================

async def check_and_close_sidebar(page):
    """检测并关闭新打开的侧边栏 - 优化版，针对豆包右上角"""
    if DEBUG:
        print("\n[侧边栏检测] 检查是否有新增侧边栏...")
    else:
        # 非调试模式只打印简单信息
        pass
    
    # 等待一下让侧边栏可能出现的元素加载
    await asyncio.sleep(1)
    
    # 获取视口大小
    viewport = page.viewport_size
    if not viewport:
        viewport = {'width': 1280, 'height': 720}
    
    # 豆包侧边栏关闭按钮的精确选择器
    close_selectors = [
        'button[aria-label="关闭"]',
        'button[aria-label="Close"]',
        'button[class*="close"]',
        'button[title*="关闭"]',
        '.semi-button.semi-button-tertiary.semi-button-with-icon',
        'button.semi-button:has(svg)',
        '.semi-icon-close',
        'svg[class*="close"]',
        'div[class*="sidebar"] button:last-child',
        'div[class*="drawer"] button:last-child',
        'button:has(svg[viewBox="0 0 14 14"])',
        'button > svg[width="14"][height="14"]'
    ]
    
    # 查找并点击关闭按钮
    for selector in close_selectors:
        try:
            buttons = await page.locator(selector).all()
            for btn in buttons:
                if not await btn.is_visible(timeout=1000):
                    continue
                    
                box = await btn.bounding_box()
                if not box:
                    continue
                
                # 检查是否在右上角区域
                if (box['x'] + box['width'] > viewport['width'] - 150 and 
                    box['y'] < 150):
                    
                    if DEBUG:
                        print(f"  [侧边栏] 发现右上角关闭按钮 ({box['x']:.0f}, {box['y']:.0f})")
                    
                    await btn.click()
                    if DEBUG:
                        print(f"  [OK] 侧边栏已关闭")
                    
                    await asyncio.sleep(0.5)
                    return True
        except Exception as e:
            continue
    
    # 如果找不到按钮，尝试直接点击右上角
    try:
        click_x = viewport['width'] - 50
        click_y = 50
        
        if DEBUG:
            print(f"  [尝试] 直接点击右上角坐标 ({click_x}, {click_y})")
        
        await page.mouse.click(click_x, click_y)
        await asyncio.sleep(0.5)
        
        if DEBUG:
            print(f"  [OK] 侧边栏可能已关闭")
        return True
    except Exception as e:
        if DEBUG:
            print(f"  [调试] 坐标点击异常: {e}")
    
    if DEBUG:
        print(f"  [OK] 未检测到需要关闭的侧边栏")
    return False

async def get_page_content(page):
    """获取页面主要内容"""
    try:
        # 尝试多个选择器获取响应内容
        content_selectors = [
            '.markdown-body',
            '.ds-markdown',
            '[class*="message"]:last-child',
            '[class*="response"]:last-child',
            '.semi-typography:last-child',
            'main',
            'body'
        ]
        
        for selector in content_selectors:
            try:
                elements = await page.locator(selector).all()
                if elements:
                    # 取最后一个元素
                    element = elements[-1]
                    if await element.is_visible(timeout=500):
                        text = await element.text_content() or ""
                        if text.strip():
                            return text
            except:
                continue
        
        # 如果找不到特定元素，获取整个页面的文本
        return await page.evaluate("document.body.innerText") or ""
        
    except Exception as e:
        print(f"  [错误] 获取页面内容失败: {e}")
        return ""

# ==================== 流式响应检测（优化版） ====================
async def wait_for_page_stable_streaming(page, timeout=300, check_interval=1.0):
    """
    针对流式输出的智能加载检测 - 增强版
    监测点：
    1. 等待第一次内容变化
    2. 然后监测直到连续5次稳定
    3. 检测是否有"继续生成"按钮
    """
    print("\n[加载检测] 开始监测流式响应...")
    start_time = time.time()
    
    # 获取初始内容
    prev_content = await get_page_content(page)
    prev_length = len(prev_content)
    print(f"  [初始] 内容长度: {prev_length} 字符")
    
    # 等待第一次变化
    stable_count = 0
    first_change = False
    max_wait_first = 10  # 增加等待第一次变化的时间
    
    # 等待第一次内容变化
    wait_start = time.time()
    while not first_change and (time.time() - wait_start) < max_wait_first:
        await asyncio.sleep(check_interval)
        current_content = await get_page_content(page)
        current_length = len(current_content)
        
        if current_length != prev_length:
            first_change = True
            print(f"  [首次变化] {prev_length} -> {current_length}")
            prev_length = current_length
            break
        
        # 每5秒输出一次状态
        if int(time.time() - wait_start) % 5 == 0:
            print(f"  [等待首次变化] 已等待 {time.time()-wait_start:.1f}秒...")
    
    if not first_change:
        print("  [警告] 未检测到内容变化，但继续监测...")
    
    # 持续监测直到连续5次稳定（增加稳定性要求）
    no_change_count = 0
    last_log_time = time.time()
    stable_required = 5  # 
    
    while True:
        await asyncio.sleep(check_interval)
        
        # 获取当前内容
        current_content = await get_page_content(page)
        current_length = len(current_content)
        
        # 检查是否超时
        elapsed = time.time() - start_time
        if elapsed > timeout:
            print(f"  ⏰ 监测超时 ({timeout}秒)")
            break
        
        # 每10秒输出一次状态
        if time.time() - last_log_time >= 10:
            print(f"  [等待中] 已加载 {elapsed:.1f}秒, 当前内容长度: {current_length} 字符, 稳定次数: {no_change_count}/{stable_required}")
            last_log_time = time.time()
        
        # 检查是否有"继续生成"按钮
        try:
            continue_button = await page.query_selector('button:has-text("继续生成")')
            if continue_button and await continue_button.is_visible():
                print(f"  [检测] 发现'继续生成'按钮，点击继续...")
                await continue_button.click()
                await asyncio.sleep(2)
                no_change_count = 0  # 重置稳定计数
                continue
        except:
            pass
        
        # 检查内容是否变化
        if current_length != prev_length:
            #print(f"  [状态] 长度变化: {prev_length} -> {current_length} (变化量: {current_length - prev_length})")
            prev_length = current_length
            no_change_count = 0  # 重置计数
        else:
            no_change_count += 1
            if no_change_count % 2 == 0:  # 每2次稳定打印一次
                print(f"  [稳定检测] 第{no_change_count}/{stable_required}次稳定")
            
            if no_change_count >= stable_required:
                print(f"  [OK] 流式响应完成 (耗时: {elapsed:.1f}秒, 连续{stable_required}次稳定)")
                return True, current_content, current_length
    
    return False, current_content, current_length


# ==================== 剪贴板操作 ====================

def get_clipboard_content():
    try:
        return pyperclip.paste()
    except Exception as e:
        iprint(f"[剪贴板] 读取失败: {e}")
        return None

def format_clipboard_content(content):
    """格式化剪贴板内容：删除多余空行，清理格式符"""
    
    if not content:
        iprint("[剪贴板] 无内容可保存")
        sys.stderr.write("[DEBUG] content为空，返回None\n")
        sys.stderr.flush()
        sys.stdout.write('\n')
        sys.stdout.flush()
        return None
    
    iprint(f"[剪贴板] 准备保存内容，长度: {len(content)} 字符")
    sys.stderr.write(f"[DEBUG] 准备保存内容，长度: {len(content)}\n")
    sys.stderr.flush()
    
    # 按行分割
    lines = content.splitlines()
    
    # 处理每一行：去除首尾空白，但保留行内空格
    processed_lines = []
    for line in lines:
        # 去除行首行尾的空白字符
        stripped = line.strip()
        # 如果去除后还有内容，保留原格式（只去除首尾空白）
        if stripped:
            # 移除行内的多余空格（可选，根据需求）
            # cleaned_line = re.sub(r'\s+', ' ', stripped)
            processed_lines.append(line.rstrip('\r'))  # 只去除行尾的回车
        # 空行跳过
    
    # 用单个换行符连接非空行
    formatted = '\n'.join(processed_lines)
    
    return formatted

# 修改 save_clipboard_to_file 函数，添加格式化处理
def save_clipboard_to_file(content, platform_name, input_preview):
    """
    保存剪贴板内容到文件（自动格式化）
    兼容 WSL 跨系统调用和本地函数调用
    """
    if not content:
        iprint("[剪贴板] 无内容可保存")
        sys.stdout.write('\n')
        sys.stdout.flush()
        return None
    
    iprint(f"[剪贴板] 准备保存内容，长度: {len(content)} 字符")
    
    # ===== 清理 Markdown 代码块标记 =====
    original_content = content
    # 移除开头的 ```json 或 ```
    content = re.sub(r'^```json\s*\n', '', content, flags=re.MULTILINE)
    content = re.sub(r'^```\s*\n', '', content, flags=re.MULTILINE)
    # 移除结尾的 ```
    content = re.sub(r'\n```\s*$', '', content, flags=re.MULTILINE)
    # 移除可能存在的多余空白行
    content = content.strip()
    
    if original_content != content:
        iprint(f"[剪贴板] 已移除 Markdown 标记")
    
    # 格式化内容（删除多余空行）
    formatted_content = format_clipboard_content(content)
    
    if formatted_content != content:
        iprint(f"[剪贴板] 内容已格式化：原长度 {len(content)} -> 新长度 {len(formatted_content)}")
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    
    # 调试信息输出到 stderr
    import sys
    print(f"[DEBUG] input_preview原始: {repr(input_preview)}", file=sys.stderr)
    print(f"[DEBUG] input_preview长度: {len(input_preview)}", file=sys.stderr)
    
    # 使用字节长度截断，而不是字符数
    safe_preview = sanitize_filename(input_preview, max_bytes=50)
    
    print(f"[DEBUG] safe_preview处理后: {repr(safe_preview)}", file=sys.stderr)
    print(f"[DEBUG] safe_preview长度: {len(safe_preview)}", file=sys.stderr)
    
    # 生成文件名
    filename = f"{platform_name}_rsp__{timestamp}.txt"
    
    print(f"[DEBUG] 生成的文件名: {repr(filename)}", file=sys.stderr)
    filepath = os.path.join(LLMRSP_DIR, filename)
    
    try:
        # 确保目录存在
        os.makedirs(LLMRSP_DIR, exist_ok=True)
        iprint(f"[剪贴板] 目录已确保: {LLMRSP_DIR}")
        
        # ===== 使用二进制模式写入，确保精确 =====
        iprint(f"[剪贴板] 正在写入文件: {filepath}")
        
        # 将内容编码为字节
        content_bytes = formatted_content.encode('utf-8')
        
        # 使用二进制模式写入
        with open(filepath, 'wb') as f:
            f.write(content_bytes)
            f.flush()  # 强制刷新缓冲区
            os.fsync(f.fileno())  # 确保写入磁盘
        
        # 验证文件大小
        file_size = os.path.getsize(filepath)
        iprint(f"[剪贴板] 文件写入完成: 预期 {len(content_bytes)} 字节, 实际 {file_size} 字节")
        
        if file_size != len(content_bytes):
            iprint(f"[警告] 文件大小不匹配！可能写入不完整")
        
        # ===== 输出标记（兼容两种场景）=====
        result_token = f"\n###RESULT_PATH###:{filepath}\n"
        path_token = f"\n@@@PATH_START@@@{filepath}@@@PATH_END@@@\n"
        
        safe_output(result_token)
        safe_output(path_token)
        
        sys.stderr.write(f"[DEBUG] 文件路径已输出: {filepath}\n")
        sys.stderr.flush()
        
        return filepath
        
    except Exception as e:
        iprint(f"[错误] 保存文件失败: {e}")
        return None
# ==================== 浏览器管理器 ====================

class BrowserManager:
    """浏览器管理器，支持复用已打开的浏览器"""
    
    def __init__(self):
        self.browser = None
        self.context = None
        self.page = None
        self.playwright = None
        self.cdp_port = None
        self.browser_pid = None
        self.browser_process = None
        # 从配置获取浏览器模式
        self.headless = config.should_use_headless
        self.screenshot_on_error = config.screenshot_on_error
        
    def save_state(self):
        if self.cdp_port and self.browser_pid:
            state = {
                "cdp_port": self.cdp_port,
                "browser_pid": self.browser_pid,
                "timestamp": time.time()
            }
            try:
                with open(BROWSER_STATE_FILE, "w", encoding="utf-8") as f:
                    json.dump(state, f)
                iprint(f"[浏览器] 状态已保存到 {BROWSER_STATE_FILE}")
            except Exception as e:
                iprint(f"[浏览器] 保存状态失败: {e}")
    
    def load_state(self):
        if os.path.exists(BROWSER_STATE_FILE):
            try:
                with open(BROWSER_STATE_FILE, "r", encoding="utf-8") as f:
                    state = json.load(f)
                
                if self.is_process_running(state.get("browser_pid")):
                    self.cdp_port = state.get("cdp_port")
                    self.browser_pid = state.get("browser_pid")
                    iprint(f"[浏览器] 从文件加载状态成功: 端口={self.cdp_port}, PID={self.browser_pid}")
                    return True
                else:
                    iprint("[浏览器] 保存的浏览器进程已不存在")
            except Exception as e:
                iprint(f"[浏览器] 加载状态失败: {e}")
        return False
    
    def is_process_running(self, pid):
        if not pid:
            return False
        try:
            if platform.system() == "Windows":
                result = subprocess.run(f'tasklist /FI "PID eq {pid}"', 
                                       shell=True, capture_output=True, text=True)
                return str(pid) in result.stdout
            else:
                os.kill(pid, 0)
                return True
        except:
            return False
    
    async def connect_to_existing_browser(self):
        if not self.cdp_port:
            return False
        
        try:
            self.playwright = await async_playwright().start()
            self.browser = await self.playwright.chromium.connect_over_cdp(
                f"http://127.0.0.1:{self.cdp_port}"
            )
            
            contexts = self.browser.contexts
            if contexts:
                self.context = contexts[0]
                pages = self.context.pages
                if pages:
                    self.page = pages[0]
                    iprint("[浏览器] 成功连接到已存在的浏览器")
                    return True
                else:
                    self.page = await self.context.new_page()
                    return True
            
            iprint("[浏览器] 未找到现有的context")
            return False
        except Exception as e:
            iprint(f"[浏览器] 连接失败: {e}")
            return False
    
    def get_chrome_path(self):
        """获取Chrome浏览器路径 - 从配置读取"""
        return config.get_chrome_path()

    async def launch_new_browser(self, url=None):
        """启动新浏览器"""
        chrome_path = self.get_chrome_path()
        if not chrome_path:
            iprint("[X] 未找到Chrome浏览器，请手动安装")
            return False
        
        self.cdp_port = random.randint(9222, 9299)
        
        iprint(f"\n[浏览器] 启动新浏览器，调试端口: {self.cdp_port}")
        iprint(f"[浏览器] 模式: {'无头' if self.headless else '有头'}")
        iprint(f"[浏览器] Chrome路径: {chrome_path}")
        
        # 构建启动命令
        cmd = [
            chrome_path,
            f"--remote-debugging-port={self.cdp_port}",
            f"--user-data-dir={os.path.abspath(BROWSER_DEBUG_DIR)}",
            "--no-first-run",
            "--no-default-browser-check"
        ]
        
        # 添加无头模式参数
        if self.headless:
            cmd.append("--headless=new")
            cmd.append("--disable-gpu")
            cmd.append("--no-sandbox")
            cmd.append("--disable-dev-shm-usage")
            cmd.append("--disable-blink-features=AutomationControlled")
        
        if url:
            cmd.append(url)
        
        try:
            # 启动浏览器进程
            if platform.system() == "Windows":
                self.browser_process = subprocess.Popen(
                    cmd,
                    shell=False,
                    creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0
                )
            else:
                self.browser_process = subprocess.Popen(cmd)
            
            self.browser_pid = self.browser_process.pid
            iprint(f"[浏览器] 进程PID: {self.browser_pid}")
            
            # 等待浏览器启动
            iprint("[浏览器] 等待浏览器启动...")
            await asyncio.sleep(5)
            
            # 通过playwright连接
            self.playwright = await async_playwright().start()
            
            # 尝试多次连接
            max_retries = 5
            for i in range(max_retries):
                try:
                    self.browser = await self.playwright.chromium.connect_over_cdp(
                        f"http://127.0.0.1:{self.cdp_port}"
                    )
                    iprint(f"[浏览器] 第{i+1}次连接成功")
                    break
                except Exception as e:
                    if i < max_retries - 1:
                        iprint(f"[浏览器] 连接尝试 {i+1}/{max_retries} 失败，等待重试...")
                        await asyncio.sleep(2)
                    else:
                        raise e
            
            self.context = self.browser.contexts[0] if self.browser.contexts else await self.browser.new_context()
            self.page = self.context.pages[0] if self.context.pages else await self.context.new_page()
            
            # 保存状态
            self.save_state()
            iprint("[浏览器] 新浏览器启动并连接成功")
            return True
            
        except Exception as e:
            iprint(f"[浏览器] 连接新浏览器失败: {e}")
            return False  
  
    async def close(self, keep_browser=True):
        if self.playwright:
            if keep_browser:
                await self.playwright.stop()
                iprint("[浏览器] 已断开连接，浏览器进程继续运行")
            else:
                if self.browser:
                    await self.browser.close()
                await self.playwright.stop()
                if self.browser_process:
                    self.browser_process.terminate()
                if os.path.exists(BROWSER_STATE_FILE):
                    os.remove(BROWSER_STATE_FILE)
                iprint("[浏览器] 浏览器已关闭")

class BaseElementFinder:
    """通用元素查找器，提供统一的查找策略"""
    
    def __init__(self, page, platform_name, element_name, history_manager):
        self.page = page
        self.platform_name = platform_name
        self.element_name = element_name
        self.history = history_manager
     
    async def safe_find(self, coro, error_msg="操作失败"):
        """安全的查找操作"""
        try:
            return await coro
        except Exception as e:
            # 检查是否是页面关闭相关的错误
            error_str = str(e).lower()
            if any(keyword in error_str for keyword in ['target closed', 'page closed', 'browser closed', 'execution context']):
                dprint(f"  [安全退出] 页面已关闭: {e}")
                raise asyncio.CancelledError()
            dprint(f"  [安全退出] {error_msg}: {e}")
            return None
     
     # 在 BaseElementFinder 类中增强 find_by_history 方法
    async def find_by_history(self, selectors, validator_func, search_radius=50):
        print(f"[DEBUG][find_by_history] 开始 - {datetime.now().strftime('%H:%M:%S.%f')}")
        
        try:
            await self.page.evaluate('1')
        except:
            dprint("  页面已关闭，find_by_history 退出")
            return None
        
        best_coord = self.history.get_best_coord()
        if not best_coord:
            iprint(f"  [历史] 无历史记录")
            return None
                    
        target_x, target_y = best_coord
        iprint(f"  [历史] {self.element_name} 使用坐标 ({target_x:.0f}, {target_y:.0f}) 半径:{search_radius}px")
        
        total_elements = 0
        for selector in selectors:
            try:
                elements = await self.page.locator(selector).all()
                if elements:
                    dprint(f"    选择器 '{selector}' 找到 {len(elements)} 个元素")
                
                for elem in elements:
                    total_elements += 1
                    if not await elem.is_visible():
                        continue
                    
                    box = await elem.bounding_box()
                    if not box:
                        continue
                    
                    center_x = box['x'] + box['width'] / 2
                    center_y = box['y'] + box['height'] / 2
                    
                    distance = abs(center_x - target_x) + abs(center_y - target_y)
                    
                    if distance < search_radius * 2:  # 记录附近元素
                        dprint(f"      附近元素: ({box['x']:.0f},{box['y']:.0f}) 距离:{distance:.0f}px")
                    
                    if abs(center_x - target_x) < search_radius and abs(center_y - target_y) < search_radius:
                        dprint(f"      尝试验证: ({box['x']:.0f},{box['y']:.0f})")
                        is_valid, metadata = await validator_func(elem)
                        if is_valid:
                            iprint(f"  [OK] 历史附近找到 {self.element_name}")
                            self.history.update_cache(elem, box)
                            self.history.save(box['x'], box['y'], metadata)
                            return elem
            except Exception as e:
                dprint(f"    选择器 '{selector}' 异常: {e}")
                continue
        
        iprint(f"  [历史] 共检查 {total_elements} 个元素，未找到匹配")
        return None


    # 增强 find_by_area 方法
    async def find_by_area(self, center_x, center_y, selectors, validator_func, radius=100):
        print(f"[DEBUG][find_by_area] 开始 - {datetime.now().strftime('%H:%M:%S.%f')}")
        search_area = {
            'x_min': center_x - radius,
            'x_max': center_x + radius,
            'y_min': center_y - radius,
            'y_max': center_y + radius,
        }
        
        iprint(f"  [区域] 搜索范围: X[{search_area['x_min']:.0f}-{search_area['x_max']:.0f}] Y[{search_area['y_min']:.0f}-{search_area['y_max']:.0f}]")
        
        candidates = []
        total_elements = 0
        
        for selector in selectors:
            try:
                elements = await self.page.locator(selector).all()
                for elem in elements:
                    total_elements += 1
                    if not await elem.is_visible():
                        continue
                    
                    box = await elem.bounding_box()
                    if not box:
                        continue
                    
                    center_x_elem = box['x'] + box['width'] / 2
                    center_y_elem = box['y'] + box['height'] / 2
                    
                    if (search_area['x_min'] <= center_x_elem <= search_area['x_max'] and
                        search_area['y_min'] <= center_y_elem <= search_area['y_max']):
                        
                        distance = abs(center_x_elem - center_x) + abs(center_y_elem - center_y)
                        dprint(f"      区域候选: ({box['x']:.0f},{box['y']:.0f}) 距离:{distance:.0f}px")
                        
                        is_valid, metadata = await validator_func(elem)
                        if is_valid:
                            candidates.append({
                                'element': elem,
                                'box': box,
                                'metadata': metadata,
                                'distance': distance
                            })
            except Exception as e:
                continue
        
        iprint(f"  [区域] 共检查 {total_elements} 个元素，找到 {len(candidates)} 个候选")
        
        if candidates:
            candidates.sort(key=lambda x: x['distance'])
            best = candidates[0]
            
            self.history.update_cache(best['element'], best['box'])
            self.history.save(best['box']['x'], best['box']['y'], best.get('metadata'))
            
            iprint(f"  [OK] 选择最佳候选: ({best['box']['x']:.0f}, {best['box']['y']:.0f}) 距离:{best['distance']:.0f}px")
            return best['element']
        
        return None
     
# ==================== 通用历史记录管理器 ====================
class HistoryManager:
    """通用历史记录管理器 - 增强健壮性"""

    def __init__(self, name, platform=None):
        """
        初始化历史记录管理器
        Args:
            name: 记录名称（如 'copy_button', 'upload_icon'）
            platform: 平台名称（用于区分不同平台的历史记录，如 'DeepSeek'）
        """
        self.name = name
        self.platform = platform
        self.cached_element = None
        self.cached_box = None
        
        # ===== 构建历史文件名 =====
        if platform:
            self.history_name = f"{name}_{platform}"
        else:
            self.history_name = name
        
        iprint(f"[历史] 初始化 {self.history_name}")
        
        # ===== 默认值，确保不会因为配置问题而崩溃 =====
        self.max_records = 5
        self.history = []
        self.history_file = None
        
        # ===== 尝试加载配置，但失败时不报错 =====
        try:
            if hasattr(config, 'config') and 'history' in config.config:
                self.max_records = config.config['history'].get('max_records', 5)
                iprint(f"[历史] 从配置加载 max_records = {self.max_records}")
        except Exception as e:
            iprint(f"[历史] 加载配置失败: {e}，使用默认值")
            pass
        
        # ===== 尝试获取历史文件路径，失败时使用内存模式 =====
        try:
            if hasattr(config, 'get_history_file'):
                self.history_file = config.get_history_file(self.name, self.platform)
                iprint(f"[历史] 文件路径: {self.history_file}")
                
                if self.history_file:
                    # 确保目录存在
                    self.history_file.parent.mkdir(parents=True, exist_ok=True)
                    iprint(f"[历史] 目录已创建/存在: {self.history_file.parent}")
                    
                    # 检查目录是否可写
                    if os.access(self.history_file.parent, os.W_OK):
                        iprint(f"[历史] 目录可写")
                    else:
                        iprint(f"[历史] 警告: 目录不可写")
        except Exception as e:
            iprint(f"[历史] 获取历史文件路径失败: {e}，将使用内存记录")
            self.history_file = None
        
        # ===== 加载历史记录（如果文件存在） =====
        if self.history_file and self.history_file.exists():
            self.history = self._load()
            iprint(f"[历史] {self.history_name} 加载 {len(self.history)} 条记录")
        else:
            self.history = []
            if self.history_file:
                iprint(f"[历史] {self.history_name} 历史文件不存在，将使用内存记录")
   
    def _load(self):
        """加载历史记录 - 失败时返回空列表"""
        if not self.history_file:
            return []
        
        try:
            with open(self.history_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                # 确保是列表
                if not isinstance(data, list):
                    return []
                
                # 按时间戳排序，最新的在前
                data.sort(key=lambda r: -r.get('timestamp', 0))
                if len(data) > self.max_records:
                    data = data[:self.max_records]
                
                return data
        except Exception as e:
            # 任何错误都不报错，静默返回空列表
            iprint(f"[历史] {self.history_name} 加载失败: {e}，使用空记录")
            return []

    def save(self, x, y, metadata=None):
        """保存成功坐标 - 增强版，保存完整的元素特征"""
        try:
            # 构建完整的记录
            record = {
                'x': x,
                'y': y,
                'timestamp': time.time(),
                'date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            
            # 如果提供了metadata，确保包含完整的元素信息
            if metadata:
                # 确保metadata中有完整的box信息
                if 'box' not in metadata and 'width' in metadata and 'height' in metadata:
                    metadata['box'] = {
                        'x': x - metadata.get('width', 0) / 2,  # 估算左上角x
                        'y': y - metadata.get('height', 0) / 2,  # 估算左上角y
                        'width': metadata.get('width', 0),
                        'height': metadata.get('height', 0)
                    }
                
                # 确保有classes字段
                if 'classes' not in metadata:
                    metadata['classes'] = metadata.get('class', '')
                
                # 确保有tooltip字段
                if 'tooltip' not in metadata:
                    metadata['tooltip'] = ''
                
                record['metadata'] = metadata
            else:
                # 创建基本的metadata
                record['metadata'] = {
                    'tooltip': '',
                    'classes': '',
                    'aria': '',
                    'title': '',
                    'box': {
                        'x': x - 17,  # 估算（34/2）
                        'y': y - 17,
                        'width': 34,
                        'height': 34
                    }
                }
            
            self.history.append(record)
            
            # 按时间戳排序并保留最新的
            self.history.sort(key=lambda r: -r['timestamp'])
            self.history = self.history[:self.max_records]
            
            # 尝试保存到文件
            if self.history_file:
                try:
                    self.history_file.parent.mkdir(parents=True, exist_ok=True)
                    
                    iprint(f"[历史] 尝试保存到文件: {self.history_file}")
                    
                    with open(self.history_file, 'w', encoding='utf-8') as f:
                        json.dump(self.history, f, indent=2, ensure_ascii=False)
                    
                    if self.history_file.exists():
                        file_size = self.history_file.stat().st_size
                        iprint(f"[历史] 文件已保存，大小: {file_size} 字节")
                        
                except Exception as e:
                    iprint(f"[历史] 文件保存失败: {e}")
                    pass
            
            return True
        except Exception as e:
            iprint(f"[历史] {self.history_name} 保存失败: {e}")
            return False
        
    def get_best_coord(self):
        """获取最佳历史坐标 - 没有就返回None"""
        if self.history and len(self.history) > 0:
            best = self.history[0]  # 最新的
            return (best['x'], best['y'])
        return None
    
    def get_best_coord_with_metadata(self):
        """获取最佳历史坐标及元数据"""
        if self.history and len(self.history) > 0:
            best = self.history[0]  # 最新的
            return {
                'x': best['x'],
                'y': best['y'],
                'metadata': best.get('metadata', {}),
                'timestamp': best.get('timestamp', 0)
            }
        return None
    
    def extract_stable_features(self):
        """从历史记录提取稳定的特征"""
        if not self.history:
            return {}
        
        # 分析多个历史记录，找出稳定的特征
        classes_count = {}
        tooltips = set()
        sizes = []
        
        for record in self.history[:3]:  # 分析最近3条
            metadata = record.get('metadata', {})
            
            # 统计类名
            classes = metadata.get('classes', '')
            for cls in classes.split():
                classes_count[cls] = classes_count.get(cls, 0) + 1
            
            # 收集tooltip
            if metadata.get('tooltip'):
                tooltips.add(metadata['tooltip'])
            
            # 收集尺寸
            if metadata.get('box'):
                sizes.append((metadata['box']['width'], metadata['box']['height']))
        
        # 找出出现次数最多的类名（稳定的特征）
        stable_classes = [cls for cls, count in classes_count.items() 
                         if count >= len(self.history) // 2]
        
        # 计算平均尺寸
        avg_width = sum(w for w, _ in sizes) / len(sizes) if sizes else None
        avg_height = sum(h for _, h in sizes) / len(sizes) if sizes else None
        
        return {
            'stable_classes': stable_classes,
            'tooltips': list(tooltips),
            'avg_width': avg_width,
            'avg_height': avg_height,
            'has_ds_icon': any('ds-icon' in cls for cls in stable_classes)
        }
       
    def update_cache(self, element, box):
        """更新缓存"""
        self.cached_element = element
        self.cached_box = box
    
    async def get_cached(self, page, tolerance=10):
        """获取缓存的元素（检查是否仍有效）"""
        if self.cached_element and self.cached_box:
            try:
                is_connected = await page.evaluate("el => el.isConnected", self.cached_element)
                if is_connected:
                    new_box = await self.cached_element.bounding_box()
                    if new_box and abs(new_box['x'] - self.cached_box['x']) < tolerance:
                        return self.cached_element
            except:
                pass
        return None
    
    def clear_cache(self):
        """清除缓存"""
        self.cached_element = None
        self.cached_box = None
        
# ==================== 基础Handler类 ====================

class BaseHandler:
    """基础处理器，包含通用方法"""
    
    def __init__(self, page, platform_name):
        self.page = page
        self.platform_name = platform_name

    async def check_login_status(self):
        """通用登录状态检查 - 仅通过登录按钮是否存在判断"""
        iprint(f"\n[{self.platform_name}] 检查登录状态...")
            
        # 获取视口大小，用于定位侧边栏区域
        viewport = self.page.viewport_size
        if not viewport:
            viewport = {'width': 1280, 'height': 720}
        
        # 通用登录按钮选择器
        login_selectors = [
            # 通用按钮选择器
            'button:has-text("登录")',
            'button:has-text("注册")',
            'button:has-text("Log in")',
            'button:has-text("Sign in")',
            'button:has-text("Register")',
            '.login-btn',
            '[class*="login"] button',
            '[class*="Login"] button',
            'a:has-text("登录")',
            'a:has-text("Log in")',
            'a:has-text("Sign in")',
            '[class*="login"] a',
            
            # 针对Kimi的侧边栏登录按钮
            '[class*="sidebar"] button:has-text("登录")',
            '[class*="Sidebar"] button:has-text("登录")',
            '[class*="menu"] button:has-text("登录")',
            '[class*="Menu"] button:has-text("登录")',
            'aside button:has-text("登录")',
            'nav button:has-text("登录")',
            
            # 针对Kimi的特定类名（根据实际情况可能需要调整）
            '.sidebar-footer button',
            '.menu-footer button',
            '[class*="user"] button:has-text("登录")',
            '[class*="User"] button:has-text("登录")',
            
            # 通用的侧边栏底部查找
            'aside:last-child button:last-child',
            'nav:last-child button:last-child',
            '[class*="sidebar"]:last-child button:last-child',
            '[class*="Sidebar"]:last-child button:last-child'
        ]
        
        # 如果是Kimi平台，增加特定的搜索区域
        if self.platform_name == "Kimi":
            # 专门针对Kimi左侧边栏的查找策略
            kimi_specific_selectors = [
                # 左侧边栏区域
                '[class*="sidebar"]',
                '[class*="Sidebar"]',
                'aside',
                'nav',
                '[class*="menu"]',
                '[class*="Menu"]'
            ]
            
            # 在左侧边栏区域查找登录按钮
            for area_selector in kimi_specific_selectors:
                try:
                    # 查找侧边栏元素
                    sidebars = await self.page.locator(area_selector).all()
                    for sidebar in sidebars:
                        if not await sidebar.is_visible():
                            continue
                        
                        # 在侧边栏内查找登录按钮
                        for btn_selector in ['button:has-text("登录")', 'a:has-text("登录")', '[class*="login"]']:
                            buttons = await sidebar.locator(btn_selector).all()
                            for btn in buttons:
                                if await btn.is_visible(timeout=500):
                                    iprint(f"  [发现] 在左侧边栏找到登录按钮: {area_selector} -> {btn_selector}")
                                    return False
                except:
                    continue
        
        # 检查是否存在登录按钮
        for selector in login_selectors:
            try:
                elements = await self.page.locator(selector).all()
                for elem in elements:
                    if await elem.is_visible(timeout=500):
                        # 获取元素位置信息，用于调试
                        if DEBUG:
                            box = await elem.bounding_box()
                            if box:
                                dprint(f"    [调试] 检测到登录按钮: {selector} 位置: ({box['x']:.0f}, {box['y']:.0f})")
                        return False  # 有登录按钮，说明未登录
            except Exception as e:
                if DEBUG:
                    dprint(f"    [调试] 选择器 {selector} 检查失败: {e}")
                continue
        
        # 最后的尝试：查找任何包含"登录"文字的可见元素
        try:
            # 使用更通用的方式查找包含"登录"文字的可见元素
            login_text_elements = await self.page.locator('text="登录"').all()
            for elem in login_text_elements:
                if await elem.is_visible():
                    # 检查是否是按钮或可点击元素
                    tag_name = await elem.evaluate('el => el.tagName.toLowerCase()')
                    role = await elem.get_attribute('role') or ''
                    
                    if tag_name in ['button', 'a'] or 'button' in role:
                        if DEBUG:
                            box = await elem.bounding_box()
                            if box:
                                dprint(f"    [调试] 检测到包含登录文字的按钮: {tag_name} 位置: ({box['x']:.0f}, {box['y']:.0f})")
                        return False
        except:
            pass
        
        # 没有找到登录按钮，认为已登录
        iprint("  [OK] 未检测到登录按钮，认为已登录")
        return True

    async def wait_for_login(self, timeout=300):
        """通用等待用户手动登录"""
        
        # 根据平台定制提示信息
        login_prompts = {
            "DeepSeek": "   1. 微信扫码登录\n   2. 或使用账号密码登录",
            "豆包": "   1. 手机号登录\n   2. 或使用账号密码登录",
            "Kimi": "   1. 微信扫码登录\n   2. 或使用手机号登录"
        }
        
        prompt = login_prompts.get(self.platform_name, "   请完成登录")
        
        iprint("\n" + "="*60)
        iprint(f"[WARN]  请在浏览器中完成{self.platform_name}登录")
        iprint(prompt)
        iprint("="*60)
        
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            await asyncio.sleep(2)
            
            # 检查登录按钮是否消失（说明登录成功）
            if await self.check_login_status():
                elapsed = time.time() - start_time
                iprint(f"  [OK] 登录成功！ (耗时: {elapsed:.1f}秒)")
                return True
            
            # 每10秒输出一次等待信息
            elapsed = time.time() - start_time
            if int(elapsed) % 10 == 0:
                iprint(f"  [等待登录] 已等待 {elapsed:.1f}秒...")
        
        iprint(f"  [WARN] 登录超时 ({timeout}秒)")
        return False
    
    async def get_input_value(self, input_element):
        """获取输入框的值"""
        try:
            tag_name = await input_element.evaluate('el => el.tagName.toLowerCase()')
            is_content_editable = await input_element.evaluate('el => el.isContentEditable')
            
            if tag_name in ['input', 'textarea']:
                return await input_element.input_value()
            elif is_content_editable or tag_name == 'div':
                return await input_element.evaluate('el => el.textContent')
            return None
        except Exception as e:
            dprint(f"    获取输入值失败: {e}")
            return None
    
    async def fill_input(self, input_element, text):
        """填写输入框"""
        try:
            tag_name = await input_element.evaluate('el => el.tagName.toLowerCase()')
            is_content_editable = await input_element.evaluate('el => el.isContentEditable')
            
            if tag_name in ['input', 'textarea']:
                await input_element.fill(text)
            elif is_content_editable or tag_name == 'div':
                await input_element.click()
                await input_element.evaluate('el => el.textContent = ""')
                for char in text:
                    await input_element.type(char, delay=0.05)
                    await asyncio.sleep(0.01)
            else:
                await input_element.type(text, delay=0.05)
            
            return True
        except Exception as e:
            dprint(f"    填写失败: {e}")
            return False
    
    async def focus_input_box(self, input_element):
        """让输入框聚焦"""
        if not input_element:
            return
        
        iprint(f"\n[{self.platform_name}] 聚焦输入框...")
        try:
            await input_element.scroll_into_view_if_needed()
            await input_element.click()
            await input_element.focus()
            iprint("  [OK] 输入框已聚焦")
        except Exception as e:
            iprint(f"  [ERR] 聚焦失败: {e}")

# ==================== 复制按钮查找器（通用）====================

class CopyButtonFinder:
    """复制按钮查找器，支持逐个验证和用户交互"""
    
    def __init__(self, page, platform_name):
        self.page = page
        self.platform_name = platform_name
        self.history = HistoryManager("copy_button", platform_name)
        self.finder = BaseElementFinder(page, platform_name, "复制按钮", self.history)
        self._tasks = []  # 添加任务跟踪列表
        self._closed = False  # 关闭标志
        
    async def cleanup(self):
        """清理所有任务"""
        self._closed = True
        for task in self._tasks:
            if not task.done():
                task.cancel()
        if self._tasks:
            await asyncio.gather(*self._tasks, return_exceptions=True)
        self._tasks.clear()
    
    async def find_all_buttons(self):
        """查找页面上所有可能的按钮"""
        buttons = []
        
        # 查找所有可见的按钮元素
        button_selectors = [
            'button',
            'input[type="button"]',
            'input[type="submit"]',
            '[role="button"]',
            '[class*="btn"]',
            '[class*="button"]'
        ]
        
        for selector in button_selectors:
            elements = await self.page.locator(selector).all()
            for elem in elements:
                if await elem.is_visible():
                    buttons.append(elem)
        
        return buttons

    async def get_button_info(self, button):
        """获取按钮的详细信息（确保返回的字典不包含None）"""
        try:
            if not button:
                return {'tag': '', 'id': '', 'class': '', 'type': '', 'title': '', 'aria_label': '', 'text': '', 'value': '', 
                       'rect': {'top': 0, 'left': 0, 'width': 0, 'height': 0}}
            
            try:
                info = await button.evaluate("""el => {
                    const rect = el.getBoundingClientRect();
                    return {
                        tag: el.tagName || '',
                        id: el.id || '',
                        class: el.className || '',
                        type: el.getAttribute('type') || '',
                        title: el.getAttribute('title') || '',
                        aria_label: el.getAttribute('aria-label') || '',
                        text: (el.textContent || '').trim() || '',
                        value: el.value || '',
                        rect: {
                            top: rect.top || 0,
                            left: rect.left || 0,
                            width: rect.width || 0,
                            height: rect.height || 0
                        }
                    };
                }""")
                return info
            except:
                return {'tag': '', 'id': '', 'class': '', 'type': '', 'title': '', 'aria_label': '', 'text': '', 'value': '', 
                       'rect': {'top': 0, 'left': 0, 'width': 0, 'height': 0}}
        except:
            return {'tag': '', 'id': '', 'class': '', 'type': '', 'title': '', 'aria_label': '', 'text': '', 'value': '', 
                   'rect': {'top': 0, 'left': 0, 'width': 0, 'height': 0}}

    async def hover_and_check_tooltip(self, button, duration=2):
        """鼠标悬浮并检查提示（确保返回的字典不包含None）"""
        try:
            # 新增：检查按钮是否有效
            if not button:
                return {'before': {'title': '', 'aria': ''}, 'after': {'title': '', 'aria': ''}, 'tooltip': ''}
            
            # 获取悬浮前的提示属性
            try:
                before_title = await button.get_attribute('title') or ''
                before_aria = await button.get_attribute('aria-label') or ''
            except:
                before_title, before_aria = '', ''
            
            # 执行鼠标悬浮
            try:
                await button.hover()
                await asyncio.sleep(duration)
            except Exception as e:
                dprint(f"    悬浮操作失败: {e}")
                return {'before': {'title': before_title, 'aria': before_aria}, 
                       'after': {'title': '', 'aria': ''}, 
                       'tooltip': ''}
            
            # 获取悬浮后的提示属性
            try:
                after_title = await button.get_attribute('title') or ''
                after_aria = await button.get_attribute('aria-label') or ''
            except:
                after_title, after_aria = '', ''
            
            # 尝试获取可能出现的悬浮层元素
            try:
                tooltip = await self.page.evaluate("""() => {
                    const tooltips = document.querySelectorAll('[role="tooltip"], .tooltip, [class*="tooltip"]');
                    for (const t of tooltips) {
                        if (t.offsetParent !== null) {
                            return t.textContent || '';
                        }
                    }
                    return '';
                }""") or ''
            except:
                tooltip = ''
            
            return {
                'before': {'title': before_title, 'aria': before_aria},
                'after': {'title': after_title, 'aria': after_aria},
                'tooltip': tooltip
            }
        except Exception as e:
            dprint(f"    悬浮检查失败: {e}")
            return {'before': {'title': '', 'aria': ''}, 'after': {'title': '', 'aria': ''}, 'tooltip': ''}
 
    async def is_copy_button(self, button_info, hover_info):
        """判断是否为复制按钮（彻底修复NoneType错误）"""
        try:
            # 检查所有可能的提示文本
            check_texts = []
            
            if button_info and isinstance(button_info, dict):
                title = button_info.get('title')
                check_texts.append(str(title) if title is not None else '')
                
                aria = button_info.get('aria_label')
                check_texts.append(str(aria) if aria is not None else '')
                
                text = button_info.get('text')
                check_texts.append(str(text) if text is not None else '')
            
            if hover_info and isinstance(hover_info, dict):
                after = hover_info.get('after', {})
                if isinstance(after, dict):
                    after_title = after.get('title')
                    check_texts.append(str(after_title) if after_title is not None else '')
                    after_aria = after.get('aria')
                    check_texts.append(str(after_aria) if after_aria is not None else '')
                
                tooltip = hover_info.get('tooltip')
                check_texts.append(str(tooltip) if tooltip is not None else '')
            
            valid_texts = [text for text in check_texts if text and isinstance(text, str) and text.strip()]
            
            if not valid_texts:
                return False
            
            all_text = ' '.join(valid_texts).lower()
            copy_keywords = ['复制', 'copy', '拷贝', '📋', '📄', '⎘']
            
            for keyword in copy_keywords:
                if keyword in all_text:
                    return True
            
            return False
        except:
            return False
 
    async def scroll_to_bottom(self):
        """滚动页面到底部，确保最新内容可见"""
        iprint("  执行滚动操作，将页面滚动到底部...")
        try:
            # 方法1: 滚动主窗口
            await self.page.evaluate("""
                () => {
                    window.scrollTo({
                        top: document.documentElement.scrollHeight,
                        behavior: 'smooth'
                    });
                }
            """)
            await asyncio.sleep(1)
            
            # 方法2: 查找并滚动可能的聊天容器
            await self.page.evaluate("""
                () => {
                    const chatContainers = document.querySelectorAll(
                        '[class*="chat"], [class*="message"], [class*="conversation"], ' +
                        '[class*="scroll"], .rc-scroll, .ant-scroll, .semi-scroll'
                    );
                    chatContainers.forEach(container => {
                        if (container.scrollHeight > container.clientHeight) {
                            container.scrollTop = container.scrollHeight;
                        }
                    });
                }
            """)
            await asyncio.sleep(1)
            
            # 方法3: 模拟键盘操作（End键）
            await self.page.keyboard.press('End')
            await asyncio.sleep(1)
            
            iprint("  [OK] 滚动完成")
            return True
        except Exception as e:
            dprint(f"  滚动失败: {e}")
            return False

           
    #验证函数 - 注意不要捕获外部变量
    async def validate_copy_button(self, elem):
        """验证函数，添加关闭检查"""
        if self._closed:
            raise asyncio.CancelledError()
        
        try:
            # 检查页面是否还在
            try:
                await self.page.evaluate('1')
            except:
                raise asyncio.CancelledError()
            
            button_info = await self.get_button_info(elem)
            hover_info = await self.hover_and_check_tooltip(elem, duration=1)
            is_copy = await self.is_copy_button(button_info, hover_info)
            return (is_copy, {})
        except Exception as e:
            if self._closed or isinstance(e, asyncio.CancelledError):
                raise asyncio.CancelledError()
            return (False, {})

    # 修改 CopyButtonFinder 类的 find_copy_buttons_interactive 方法，限制最多遍历5个按钮
    async def find_copy_buttons_interactive(self, input_element=None):
        """交互式查找复制按钮，最多遍历5个最近的按钮"""
        iprint(f"\n[{self.platform_name}] 开始查找复制按钮...")
        
        # 先滚动页面到底部，确保最新的回复和复制按钮可见
        await self.scroll_to_bottom()
        
        # 获取所有按钮
        all_buttons = await self.find_all_buttons()
        
        if not all_buttons:
            iprint("  [信息] 页面上没有找到任何按钮")
            return None
        
        iprint(f"  找到 {len(all_buttons)} 个按钮")
        
        # 优先考虑输入框上方的按钮，并按距离排序
        target_buttons = []
        if input_element:
            input_box = await input_element.bounding_box()
            if input_box:
                input_center_y = input_box['y'] + input_box['height'] / 2
                
                # 计算每个按钮到输入框的距离
                buttons_with_distance = []
                for btn in all_buttons:
                    btn_box = await btn.bounding_box()
                    if btn_box:
                        btn_center_y = btn_box['y'] + btn_box['height'] / 2
                        # 只考虑输入框上方的按钮（比输入框位置高的）
                        if btn_center_y < input_center_y:
                            distance = input_center_y - btn_center_y
                            buttons_with_distance.append({
                                'button': btn,
                                'distance': distance,
                                'y': btn_center_y
                            })
                
                # 按距离排序（距离越近越优先）
                buttons_with_distance.sort(key=lambda x: x['distance'])
                
                if buttons_with_distance:
                    target_buttons = [item['button'] for item in buttons_with_distance]
                    iprint(f"  找到 {len(target_buttons)} 个在输入框上方的按钮")
                    
                    # 显示最近几个按钮的距离信息
                    for i, item in enumerate(buttons_with_distance[:3]):
                        iprint(f"    最近按钮 {i+1}: 距离 {item['distance']:.0f}px")
        
        # 如果没有找到上方的按钮，使用所有按钮
        if not target_buttons:
            target_buttons = all_buttons
            iprint("  未找到输入框上方的按钮，将检查所有按钮")
        
        # 最多只遍历5个最近的按钮
        max_buttons = min(5, len(target_buttons))
        iprint(f"  将检查最近的 {max_buttons} 个按钮...")
        
        # 逐个验证按钮
        copy_buttons = []
        for i in range(max_buttons):
            button = target_buttons[i]
            dprint(f"\n  [按钮 {i+1}/{max_buttons}]")
            
            if DEBUG:
                # 高亮显示当前按钮
                await highlight_element_with_index(self.page, button, i+1, 'orange', duration=1)
            
            # 获取按钮信息
            button_info = await self.get_button_info(button)
            if button_info:
                dprint(f"    位置: ({button_info['rect']['left']:.0f}, {button_info['rect']['top']:.0f})")
                if button_info['text']:
                    dprint(f"    文本: {button_info['text'][:50]}")
                if button_info['title']:
                    dprint(f"    title: {button_info['title']}")
                if button_info['aria_label']:
                    dprint(f"    aria-label: {button_info['aria_label']}")
            
            # 确保按钮可见（滚动到视图）
            try:
                await button.scroll_into_view_if_needed()
                await asyncio.sleep(0.5)
            except:
                pass
            
            # 鼠标悬浮查看提示
            dprint("    执行鼠标悬浮查看提示...")
            hover_info = await self.hover_and_check_tooltip(button, duration=3)
            
            if hover_info and hover_info.get('tooltip'):
                dprint(f"    悬浮提示: {hover_info['tooltip']}")
            
            # 判断是否为复制按钮
            is_copy = await self.is_copy_button(button_info, hover_info)
            
            if is_copy:
                iprint(f"    [OK] 检测到复制按钮!")
                copy_buttons.append(button)

                # 调试模式下才需要用户确认
                if DEBUG:
                    resp = input(f"\n    是否使用此复制按钮？(y/n/q-退出): ").strip().lower()
                    if resp == 'y':
                        return button
                    elif resp == 'q':
                        return None
                return button
            else:
                iprint(f"    [ERR] 不是复制按钮")
        
        if copy_buttons:
            iprint(f"\n  共找到 {len(copy_buttons)} 个复制按钮，但未选择")
        else:
            iprint(f"\n  [信息] 在最近的 {max_buttons} 个按钮中未找到复制按钮")
        
        return None

# ==================== DeepSeek Handler（最终优化版：快速定位 + 简化上传检测）====================

class DeepSeekHandler(BaseHandler):
    """DeepSeek聊天平台处理器 - 最终优化版：快速定位 + 简化上传检测"""
    
    def __init__(self, page):
        super().__init__(page, "DeepSeek")
        self.copy_finder = CopyButtonFinder(page, "DeepSeek")
        self.uploaded_file_path = None
        # 添加上传图标历史管理器
        self.upload_history = HistoryManager("upload_icon", "DeepSeek")
        self.upload_finder = BaseElementFinder(page, "DeepSeek", "上传图标", self.upload_history)
        self.history_file = self.upload_history.history_file if self.upload_history else None
        
        # ===== 新增：缓存 =====
        self.upload_features = self.upload_history.extract_stable_features()
        self.last_success_coord = None
        self._cached_upload_element = None  # 缓存元素句柄
        self._pending_file_chooser = None   # 待处理的文件选择器

    # 在 DeepSeekHandler 类中添加历史坐标加载方法
    def _load_history_coords(self):
        """加载历史成功坐标记录"""
        try:
            if os.path.exists(self.history_file):
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    # 只保留最近5次记录
                    if len(data) > 5:
                        data = data[-5:]
                    iprint(f"[历史] 加载 {len(data)} 条上传图标坐标记录")
                    return data
        except Exception as e:
            iprint(f"[历史] 加载失败: {e}")
        return []

    # 在 DeepSeekHandler 类中添加历史坐标保存方法
    def _save_history_coords(self, x, y):
        """保存成功坐标到历史记录"""
        try:
            # 加载现有记录
            history = self._load_history_coords()
            
            # 添加新记录（包含时间戳）
            history.append({
                'x': x,
                'y': y,
                'timestamp': time.time(),
                'date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            })
            
            # 只保留最近5条
            if len(history) > 5:
                history = history[-5:]
            
            # 保存到文件
            with open(self.history_file, 'w', encoding='utf-8') as f:
                json.dump(history, f, indent=2, ensure_ascii=False)
            
            iprint(f"[历史] 已记录成功坐标 ({x:.0f}, {y:.0f})")
        except Exception as e:
            iprint(f"[历史] 保存失败: {e}")

    # ===== 将验证函数定义为类的方法 =====    

    async def validate_upload(self, elem):
        """增强版上传图标验证函数 - 返回完整metadata"""
        try:
            box = await elem.bounding_box()
            if not box or box['width'] < 10 or box['height'] < 10:
                return (False, {})
            
            # 获取元素完整信息
            classes = await elem.get_attribute('class') or ''
            aria = await elem.get_attribute('aria-label') or ''
            title = await elem.get_attribute('title') or ''
            
            # 获取内部HTML
            inner_html = await elem.evaluate('el => el.innerHTML') or ''
            outer_html = await elem.evaluate('el => el.outerHTML') or ''
            
            # 检查tag和role
            tag_name = await elem.evaluate('el => el.tagName.toLowerCase()') or ''
            role = await elem.get_attribute('role') or ''
            
            # 检查是否有SVG
            has_svg = 'svg' in inner_html or 'svg' in outer_html
            
            # 鼠标悬浮获取tooltip
            tooltip = ''
            try:
                await elem.hover()
                await asyncio.sleep(0.3)
                tooltip = await self.page.evaluate("""() => {
                    const tips = document.querySelectorAll('[role="tooltip"], .tooltip');
                    for (const t of tips) {
                        if (t.offsetParent !== null) return t.textContent || '';
                    }
                    return '';
                }""")
            except:
                pass
            
            # 检查是否是上传图标
            all_text = (classes + ' ' + aria + ' ' + title + ' ' + inner_html + ' ' + outer_html).lower()
            upload_keywords = ['upload', '上传', 'attachment', '附件', '回形针', 'paperclip']
            
            is_upload = any(keyword in all_text for keyword in upload_keywords)
            
            # 构建完整的metadata
            metadata = {
                'tooltip': tooltip,
                'classes': classes,
                'aria': aria,
                'title': title,
                'box': {
                    'x': box['x'],
                    'y': box['y'],
                    'width': box['width'],
                    'height': box['height']
                },
                'tag': tag_name,
                'role': role,
                'has_svg': has_svg,
                'match_type': 'keyword' if is_upload else 'unknown'
            }
            
            # 调试输出
            if DEBUG:
                dprint(f"    验证详情:")
                dprint(f"      tag: {tag_name}")
                dprint(f"      class: {classes[:50]}")
                dprint(f"      has_svg: {has_svg}")
                dprint(f"      tooltip: {tooltip[:50] if tooltip else '无'}")
                dprint(f"      is_upload: {is_upload}")
            
            return (is_upload, metadata)
            
        except Exception as e:
            dprint(f"    验证异常: {e}")
            return (False, {})

    async def validate_upload_fast(self, elem):
        """超快速验证函数 - 只检查最核心的特征"""
        try:
            # 1. 快速检查：是否可见
            is_visible = await elem.is_visible()
            if not is_visible:
                return (False, {})
            
            # 2. 快速检查：类名是否包含 ds-icon-button
            classes = await elem.get_attribute('class') or ''
            if 'ds-icon-button' not in classes:
                return (False, {})
            
            # 3. 快速检查：尺寸大约34x34
            box = await elem.bounding_box()
            if not box:
                return (False, {})
            
            if abs(box['width'] - 34) > 10 or abs(box['height'] - 34) > 10:
                return (False, {})
            
            return (True, {'method': 'ultra_fast'})
            
        except Exception as e:
            return (False, {})

    async def find_upload_icon_optimized(self):
        """极速版上传图标查找 - 最终优化版"""
        iprint(f"\n[{self.platform_name}] 快速定位上传图标...")
        
        start_time = time.time()
        
        # ===== 阶段0：使用缓存的元素（最快）=====
        if hasattr(self, '_cached_upload_element') and self._cached_upload_element:
            try:
                # 检查元素是否还在
                is_visible = await self._cached_upload_element.is_visible()
                if is_visible:
                    iprint(f"  [⚡阶段0] 使用缓存的元素")
                    await self._cached_upload_element.hover()
                    await asyncio.sleep(0.1)
                    
                    async with self.page.expect_file_chooser(timeout=2000) as fc_info:
                        await self._cached_upload_element.click()
                    
                    self._pending_file_chooser = await fc_info.value
                    elapsed = (time.time() - start_time) * 1000
                    iprint(f"  [✅成功] 缓存元素点击生效！耗时: {elapsed:.0f}ms")
                    return self._pending_file_chooser
            except:
                self._cached_upload_element = None
        
        # ===== 阶段1：历史坐标查找 =====
        best = self.upload_history.get_best_coord_with_metadata()
        metadata = {}  # 初始化metadata，避免变量未定义错误
        
        if best:
            x, y = best['x'], best['y']
            metadata = best.get('metadata', {})
            iprint(f"  [⚡阶段1] 尝试历史坐标 ({x:.0f}, {y:.0f})")
            
            # ===== 方法1：直接获取该位置的元素并验证 =====
            try:
                # 获取该位置的元素
                element = await self.page.evaluate_handle(f"document.elementFromPoint({x}, {y})")
                
                if element:
                    # 快速验证
                    is_valid, _ = await self.validate_upload_fast(element)
                    if is_valid:
                        # 元素有效，直接点击
                        await element.hover()
                        await asyncio.sleep(0.1)
                        
                        async with self.page.expect_file_chooser(timeout=2000) as fc_info:
                            await element.click()
                        
                        self._pending_file_chooser = await fc_info.value
                        elapsed = (time.time() - start_time) * 1000
                        iprint(f"  [✅成功] 历史坐标元素直接生效！耗时: {elapsed:.0f}ms")
                        
                        # 缓存元素
                        self._cached_upload_element = element
                        self.last_success_coord = (x, y)
                        
                        # 保存精确的历史记录（更新坐标）
                        box = await element.bounding_box()
                        if box:
                            classes = await element.get_attribute('class') or ''
                            save_metadata = {
                                'classes': classes,
                                'box': {
                                    'x': box['x'],
                                    'y': box['y'],
                                    'width': box['width'],
                                    'height': box['height']
                                },
                                'success_method': 'direct_hit'
                            }
                            self.upload_history.save(x, y, save_metadata)
                        
                        return self._pending_file_chooser
            except Exception as e:
                iprint(f"  [方法1] 直接元素获取失败: {e}")
            
            # ===== 方法2：在坐标附近查找元素 =====
            iprint(f"  [方法2] 在坐标附近查找元素")
            try:
                # 缩小搜索半径到30px
                search_radius = 30
                nearby_elements = await self.page.evaluate(f"""
                    () => {{
                        const results = [];
                        const allElements = document.querySelectorAll(
                            'button, [role="button"], .ds-icon-button, div[class*="icon"]'
                        );
                        const targetX = {x};
                        const targetY = {y};
                        
                        for (const el of allElements) {{
                            if (el.offsetParent === null) continue; // 只检查可见元素
                            
                            const rect = el.getBoundingClientRect();
                            const centerX = rect.left + rect.width/2;
                            const centerY = rect.top + rect.height/2;
                            
                            // 计算曼哈顿距离
                            const distance = Math.abs(centerX - targetX) + Math.abs(centerY - targetY);
                            
                            if (distance < {search_radius}) {{
                                results.push({{
                                    x: centerX,
                                    y: centerY,
                                    distance: distance
                                }});
                            }}
                        }}
                        
                        // 按距离排序
                        results.sort((a, b) => a.distance - b.distance);
                        return results.slice(0, 3);
                    }}
                """)
                
                if nearby_elements and len(nearby_elements) > 0:
                    iprint(f"    找到 {len(nearby_elements)} 个附近元素")
                    for elem_info in nearby_elements:
                        iprint(f"      候选: ({elem_info['x']:.0f}, {elem_info['y']:.0f}) 距离:{elem_info['distance']:.0f}px")
                        
                        # 获取该位置的元素
                        element = await self.page.evaluate_handle(
                            f"document.elementFromPoint({elem_info['x']}, {elem_info['y']})"
                        )
                        
                        if element:
                            is_valid, _ = await self.validate_upload_fast(element)
                            if is_valid:
                                await element.hover()
                                await asyncio.sleep(0.1)
                                
                                async with self.page.expect_file_chooser(timeout=2000) as fc_info:
                                    await element.click()
                                
                                self._pending_file_chooser = await fc_info.value
                                elapsed = (time.time() - start_time) * 1000
                                iprint(f"  [✅成功] 附近元素查找生效！耗时: {elapsed:.0f}ms")
                                
                                # 缓存元素
                                self._cached_upload_element = element
                                
                                # 保存精确坐标（更新历史记录）
                                box = await element.bounding_box()
                                if box:
                                    exact_x = box['x'] + box['width']/2
                                    exact_y = box['y'] + box['height']/2
                                    self.last_success_coord = (exact_x, exact_y)
                                    
                                    # 保存精确的历史记录
                                    classes = await element.get_attribute('class') or ''
                                    save_metadata = {
                                        'classes': classes,
                                        'box': {
                                            'x': box['x'],
                                            'y': box['y'],
                                            'width': box['width'],
                                            'height': box['height']
                                        },
                                        'success_method': 'nearby_hit',
                                        'original_coord': {'x': x, 'y': y}
                                    }
                                    self.upload_history.save(exact_x, exact_y, save_metadata)
                                    iprint(f"  [历史] 更新精确坐标: ({exact_x:.0f}, {exact_y:.0f})")
                                
                                return self._pending_file_chooser
            except Exception as e:
                iprint(f"  [方法2] 附近元素查找失败: {e}")
            
            # ===== 方法3：使用历史类名查找 =====
            if metadata.get('classes'):
                stable_class = None
                for cls in metadata['classes'].split():
                    if 'ds-icon-button' in cls:
                        stable_class = cls
                        break
                
                if stable_class:
                    iprint(f"  [方法3] 使用历史类名查找: {stable_class}")
                    try:
                        elements = await self.page.locator(f'.{stable_class}').all()
                        for elem in elements[:2]:
                            if await elem.is_visible():
                                box = await elem.bounding_box()
                                if box:
                                    iprint(f"    找到候选元素: ({box['x']:.0f}, {box['y']:.0f})")
                                    
                                    is_valid, _ = await self.validate_upload_fast(elem)
                                    if is_valid:
                                        await elem.hover()
                                        await asyncio.sleep(0.1)
                                        
                                        async with self.page.expect_file_chooser(timeout=2000) as fc_info:
                                            await elem.click()
                                        
                                        self._pending_file_chooser = await fc_info.value
                                        elapsed = (time.time() - start_time) * 1000
                                        iprint(f"  [✅成功] 历史类名查找生效！耗时: {elapsed:.0f}ms")
                                        
                                        self._cached_upload_element = elem
                                        self.last_success_coord = (box['x'] + box['width']/2, 
                                                                  box['y'] + box['height']/2)
                                        
                                        # 保存详细的历史记录
                                        classes = await elem.get_attribute('class') or ''
                                        save_metadata = {
                                            'classes': classes,
                                            'box': {
                                                'x': box['x'],
                                                'y': box['y'],
                                                'width': box['width'],
                                                'height': box['height']
                                            },
                                            'success_method': 'class_based'
                                        }
                                        self.upload_history.save(
                                            box['x'] + box['width']/2, 
                                            box['y'] + box['height']/2, 
                                            save_metadata
                                        )
                                        
                                        return self._pending_file_chooser
                    except Exception as e:
                        iprint(f"  [方法3] 类名查找失败: {e}")
            
            # ===== 方法4：直接坐标点击（最后手段）=====
            try:
                iprint(f"  [方法4] 尝试直接坐标点击")
                await self.page.mouse.move(x, y)
                await asyncio.sleep(0.1)
                
                async with self.page.expect_file_chooser(timeout=1500) as fc_info:
                    await self.page.mouse.click(x, y)
                
                self._pending_file_chooser = await fc_info.value
                elapsed = (time.time() - start_time) * 1000
                iprint(f"  [✅成功] 坐标点击生效！耗时: {elapsed:.0f}ms")
                self.last_success_coord = (x, y)
                return self._pending_file_chooser
            except Exception as e:
                iprint(f"  [方法4] 坐标点击未触发上传: {e}")
        
        # ===== 阶段1.5：基于输入框位置的快速点击 =====
        iprint(f"  [⚡阶段1.5] 尝试输入框右下角快速点击")
        input_element = await self.find_input_box()
        if input_element:
            input_box = await input_element.bounding_box()
            if input_box:
                # 根据经验，上传图标在输入框右下角
                candidate_x = input_box['x'] + input_box['width'] - 80
                candidate_y = input_box['y'] + input_box['height'] + 30
                
                try:
                    # 先移动鼠标
                    await self.page.mouse.move(candidate_x, candidate_y)
                    await asyncio.sleep(0.1)
                    
                    # 触发文件选择器
                    async with self.page.expect_file_chooser(timeout=2000) as fc_info:
                        await self.page.mouse.click(candidate_x, candidate_y)
                    
                    self._pending_file_chooser = await fc_info.value
                    elapsed = (time.time() - start_time) * 1000
                    iprint(f"  [✅成功] 右下角快速点击生效！耗时: {elapsed:.0f}ms")
                    
                    # 保存坐标到缓存
                    self.last_success_coord = (candidate_x, candidate_y)
                    
                    # 尝试获取该位置的元素信息用于历史记录
                    try:
                        element = await self.page.evaluate_handle(f"document.elementFromPoint({candidate_x}, {candidate_y})")
                        if element:
                            self._cached_upload_element = element
                            
                            # 保存历史记录
                            box = await element.bounding_box()
                            if box:
                                classes = await element.get_attribute('class') or ''
                                save_metadata = {
                                    'classes': classes,
                                    'box': {
                                        'x': box['x'],
                                        'y': box['y'],
                                        'width': box['width'],
                                        'height': box['height']
                                    },
                                    'success_method': 'stage1.5'
                                }
                                self.upload_history.save(candidate_x, candidate_y, save_metadata)
                    except:
                        # 保底：只保存坐标
                        self.upload_history.save(candidate_x, candidate_y, {'source': 'stage1.5'})
                    
                    return self._pending_file_chooser
                    
                except Exception as e:
                    iprint(f"  [❌失败] 右下角快速点击未触发上传: {e}")
        
        # ===== 阶段2：快速选择器 =====
        iprint(f"  [⚡阶段2] 尝试快速选择器")
        
        # 精简选择器
        quick_selectors = [
            'div.ds-icon-button',
            '.ds-icon-button',
            'button:has(svg)',
            'div[role="button"]:has(svg)'
        ]
        
        for selector in quick_selectors:
            try:
                elements = await self.page.locator(selector).all()
                for elem in elements[:3]:
                    if not await elem.is_visible():
                        continue
                    
                    is_valid, _ = await self.validate_upload_fast(elem)
                    if is_valid:
                        await elem.hover()
                        await asyncio.sleep(0.1)
                        
                        try:
                            async with self.page.expect_file_chooser(timeout=1500) as fc_info:
                                await elem.click()
                            
                            self._pending_file_chooser = await fc_info.value
                            elapsed = (time.time() - start_time) * 1000
                            iprint(f"  [✅成功] 快速选择器找到元素！耗时: {elapsed:.0f}ms")
                            
                            self._cached_upload_element = elem
                            box = await elem.bounding_box()
                            if box:
                                self.last_success_coord = (box['x'] + box['width']/2, 
                                                          box['y'] + box['height']/2)
                            return self._pending_file_chooser
                        except:
                            continue
            except Exception as e:
                continue
        
        # ===== 阶段3：精确区域扫描 =====
        iprint(f"  [⚡阶段3] 尝试精确区域扫描")
        
        if best:
            x, y = best['x'], best['y']
            
            element = await self.upload_finder.find_by_area(
                x, y,
                quick_selectors,
                self.validate_upload,
                radius=100
            )
            
            if element:
                await element.hover()
                await asyncio.sleep(0.1)
                
                try:
                    async with self.page.expect_file_chooser(timeout=2000) as fc_info:
                        await element.click()
                    
                    self._pending_file_chooser = await fc_info.value
                    elapsed = (time.time() - start_time) * 1000
                    iprint(f"  [✅成功] 精确区域扫描找到元素！耗时: {elapsed:.0f}ms")
                    
                    self._cached_upload_element = element
                    box = await element.bounding_box()
                    if box:
                        center_x = box['x'] + box['width']/2
                        center_y = box['y'] + box['height']/2
                        self.last_success_coord = (center_x, center_y)
                        
                        # 保存到历史记录
                        classes = await element.get_attribute('class') or ''
                        save_metadata = {
                            'classes': classes,
                            'box': {
                                'x': box['x'],
                                'y': box['y'],
                                'width': box['width'],
                                'height': box['height']
                            },
                            'source': 'area_scan'
                        }
                        self.upload_history.save(center_x, center_y, save_metadata)
                    
                    return self._pending_file_chooser
                except:
                    pass
        
        # ===== 阶段4：兜底 - 使用原始完整搜索 =====
        iprint(f"  [🐢阶段4] 使用完整搜索")
        element = await self.find_upload_icon_slow()
        
        if element:
            elapsed = (time.time() - start_time) * 1000
            iprint(f"  [✅成功] 完整搜索找到元素！耗时: {elapsed:.0f}ms")
            
            self._cached_upload_element = element
            box = await element.bounding_box()
            if box:
                self.last_success_coord = (box['x'] + box['width']/2, 
                                          box['y'] + box['height']/2)
            return element
        else:
            iprint(f"  [❌失败] 所有方法都失败")
        
        return None
    
    async def find_upload_icon_slow(self):
        """原始的完整搜索逻辑"""
        iprint(f"  [完整搜索] 开始...")        
        iprint(f"\n[{self.platform_name}] 快速定位上传图标...")
        # 调试：检查历史记录
        iprint(f"  [调试] 历史记录文件: {self.upload_history.history_file}")
        iprint(f"  [调试] 历史记录数量: {len(self.upload_history.history)}")
        
        # 调试：检查验证函数
        iprint(f"  [调试] validate_upload 类型: {type(self.validate_upload)}")
        
        # 1. 检查缓存
        try:
            cached = await self.upload_history.get_cached(self.page)
            if cached:
                iprint(f"  [缓存] 使用缓存的上传图标")
                return cached
        except Exception as e:
            iprint(f"  [缓存] 检查失败: {e}")
        
        # 2. 上传图标选择器（扩大范围）
        upload_selectors = [
            'div.ds-icon-button',
            'div[class*="ds-icon-button"]',
            'div[class*="icon-button"]',
            'button[class*="icon"]',
            'button:has(svg)',
            'div[role="button"]:has(svg)',
            '[class*="upload"]',
            '[aria-label*="上传"]',
        ]
        
        iprint(f"  [选择器] 将尝试 {len(upload_selectors)} 个选择器")
 
        # 3. 按历史查找
        iprint("  [尝试] 按历史坐标查找...")
        element = await self.upload_finder.find_by_history(
            upload_selectors, 
            self.validate_upload,  
            search_radius=80
        )
        if element:
            iprint(f"  [OK] 历史查找成功")
            return element
        
        # 4. 获取参考点
        iprint("  [尝试] 按区域查找...")
        input_element = await self.find_input_box()
        if input_element:
            input_box = await input_element.bounding_box()
            if input_box:
                # 根据之前成功经验，上传图标在输入框右下角
                center_x = input_box['x'] + input_box['width'] - 80  # 右侧向内80px
                center_y = input_box['y'] + input_box['height'] + 70  # 底部向下70px
                iprint(f"  [参考] 基于输入框: ({center_x:.0f}, {center_y:.0f})")
            else:
                viewport = self.page.viewport_size
                center_x = viewport['width'] - 150 if viewport else 1000
                center_y = viewport['height'] - 100 if viewport else 700
        else:
            viewport = self.page.viewport_size
            center_x = viewport['width'] - 150 if viewport else 1000
            center_y = viewport['height'] - 100 if viewport else 700
            iprint(f"  [参考] 基于视口: ({center_x:.0f}, {center_y:.0f})")
        
        # 5. 区域查找（扩大范围）
        element = await self.upload_finder.find_by_area(center_x, center_y, upload_selectors, self.validate_upload, radius=300)
        if element:
            iprint(f"  [OK] 区域查找成功")
            return element
        
        # 6. 最后尝试：直接查找已知坐标
        iprint("  [尝试] 直接查找已知坐标...")
        known_x, known_y = 1388, 870  # 之前成功的坐标
        element = await self.upload_finder.find_by_area(known_x, known_y, upload_selectors, self.validate_upload, radius=50)
        if element:
            iprint(f"  [OK] 已知坐标查找成功")
            # 保存到历史
            box = await element.bounding_box()
            if box:
                self.upload_history.save(box['x'], box['y'], {'source': 'known_coord'})
            return element
        
        iprint("  [错误] 所有查找方法都失败")
        
        return None  # 临时返回，需要迁移原来的完整逻辑
               
    async def wait_for_upload_complete(self, file_path):
        """
        基于文件大小估算上传等待时间
        千兆网络理论速度：125MB/s，实际按50MB/s估算
        """
        # 获取文件大小（字节）
        file_size = os.path.getsize(file_path)
        file_size_mb = file_size / (1024 * 1024)
        
        # 基础等待时间1秒
        base_wait = 1.0
        
        # 按50MB/s估算传输时间
        transfer_speed = 50  # MB/s
        transfer_time = file_size_mb / transfer_speed
        
        # 加上基础完成时间1秒
        total_wait = base_wait + transfer_time + 1.0
        
        # 设置最小等待2秒，最大等待30秒
        total_wait = max(2.0, min(30.0, total_wait))
        
        iprint(f"\n  [上传估算] 文件大小: {file_size_mb:.2f}MB")
        iprint(f"  [上传估算] 预估传输时间: {transfer_time:.2f}秒")
        iprint(f"  [上传等待] 总计等待: {total_wait:.1f}秒")
        
        # 分阶段等待，便于显示进度
        elapsed = 0
        interval = 0.5
        steps = int(total_wait / interval)
        
        for i in range(steps):
            await asyncio.sleep(interval)
            elapsed += interval
            if i % 10 == 0:  # 每5秒显示一次进度
                iprint(f"  [上传中] 已等待 {elapsed:.1f}/{total_wait:.1f}秒")
        
        # 简单检查是否有明显的错误提示
        try:
            error_selectors = [
                'div[class*="error"]:has-text("上传失败")',
                'div[class*="error"]:has-text("upload failed")',
                'div[role="alert"]:has-text("失败")',
            ]
            
            for selector in error_selectors:
                elements = await self.page.locator(selector).all()
                for elem in elements:
                    if await elem.is_visible(timeout=500):
                        text = await elem.text_content() or ''
                        iprint(f"  [警告] 检测到错误提示: {text[:50]}")
                        break
        except:
            pass
        
        iprint(f"  [OK] 上传等待完成")
        return True

    async def upload_file_with_filechooser(self, file_path):
        """使用Playwright的file_chooser上传文件 - 最终优化版"""
        iprint(f"\n[{self.platform_name}] 开始上传文件: {file_path}")
        
        # 清除之前可能遗留的pending file_chooser
        self._pending_file_chooser = None
        
        # 查找上传图标
        result = await self.find_upload_icon_optimized()
        
        # 处理返回值
        file_chooser = None
        
        if hasattr(result, 'set_files'):  # 检查是否是file_chooser对象
            # 直接使用返回的file_chooser
            file_chooser = result
            iprint(f"  [OK] 已从 find_upload_icon_optimized 获取文件选择器")
            
            # 如果有缓存的坐标，保存精确的历史记录
            if self.last_success_coord:
                x, y = self.last_success_coord
                
                # 尝试获取精确的元素信息
                try:
                    # 获取该位置的元素
                    element = await self.page.evaluate_handle(f"document.elementFromPoint({x}, {y})")
                    if element:
                        box = await element.bounding_box()
                        if box:
                            # 计算精确的中心点
                            exact_x = box['x'] + box['width']/2
                            exact_y = box['y'] + box['height']/2
                            
                            # 如果坐标偏差较大，更新为精确坐标
                            if abs(exact_x - x) > 5 or abs(exact_y - y) > 5:
                                iprint(f"  [历史] 修正坐标: ({x:.0f}, {y:.0f}) -> ({exact_x:.0f}, {exact_y:.0f})")
                                self.last_success_coord = (exact_x, exact_y)
                                x, y = exact_x, exact_y
                            
                            # 获取元素完整信息
                            classes = await element.get_attribute('class') or ''
                            aria = await element.get_attribute('aria-label') or ''
                            title = await element.get_attribute('title') or ''
                            
                            # 获取tooltip
                            tooltip = ''
                            try:
                                await element.hover()
                                await asyncio.sleep(0.3)
                                tooltip = await self.page.evaluate("""() => {
                                    const tips = document.querySelectorAll('[role="tooltip"], .tooltip');
                                    for (const t of tips) {
                                        if (t.offsetParent !== null) return t.textContent || '';
                                    }
                                    return '';
                                }""")
                            except:
                                pass
                            
                            # 构建完整的metadata
                            metadata = {
                                'tooltip': tooltip,
                                'classes': classes,
                                'aria': aria,
                                'title': title,
                                'box': {
                                    'x': box['x'],
                                    'y': box['y'],
                                    'width': box['width'],
                                    'height': box['height']
                                },
                                'action': 'upload_success',
                                'source': 'final_save'
                            }
                            
                            # 保存到历史记录
                            self.upload_history.save(x, y, metadata)
                            iprint(f"[历史] 已保存上传图标精确坐标 ({x:.0f}, {y:.0f})")
                            
                except Exception as e:
                    iprint(f"  [历史] 保存详细信息失败: {e}")
                    # 保底：只保存坐标
                    self.upload_history.save(x, y, {'source': 'coord_fallback'})
                    iprint(f"[历史] 已保存上传图标坐标 ({x:.0f}, {y:.0f})")
        
        elif result:
            # 返回了元素对象，需要点击触发
            iprint(f"  [OK] 找到上传图标元素，准备点击...")
            
            try:
                # 点击前先hover
                await result.hover()
                await asyncio.sleep(0.1)
                
                # 触发文件选择器
                async with self.page.expect_file_chooser(timeout=5000) as fc_info:
                    await result.click()
                file_chooser = await fc_info.value
                iprint(f"  [OK] 成功捕获文件选择器")
                
                # 高亮显示
                await highlight_element_with_index(self.page, result, "上传", 'green', duration=1)
                
            except Exception as e:
                iprint(f"  [错误] 点击元素后等待文件选择器超时: {e}")
                return False
            
            # 获取元素完整信息后保存到历史记录
            box = await result.bounding_box()
            if box:
                # 计算中心点
                center_x = box['x'] + box['width']/2
                center_y = box['y'] + box['height']/2
                
                # 获取元素完整信息
                classes = await result.get_attribute('class') or ''
                aria = await result.get_attribute('aria-label') or ''
                title = await result.get_attribute('title') or ''
                
                # 获取tooltip
                tooltip = ''
                try:
                    await result.hover()
                    await asyncio.sleep(0.3)
                    tooltip = await self.page.evaluate("""() => {
                        const tips = document.querySelectorAll('[role="tooltip"], .tooltip');
                        for (const t of tips) {
                            if (t.offsetParent !== null) return t.textContent || '';
                        }
                        return '';
                    }""")
                except:
                    pass
                
                # 构建完整metadata
                metadata = {
                    'tooltip': tooltip,
                    'classes': classes,
                    'aria': aria,
                    'title': title,
                    'box': {
                        'x': box['x'],
                        'y': box['y'],
                        'width': box['width'],
                        'height': box['height']
                    },
                    'action': 'upload_click',
                    'source': 'element_click'
                }
                
                # 保存到历史记录
                self.upload_history.save(center_x, center_y, metadata)
                iprint(f"[历史] 已保存上传图标完整信息 ({center_x:.0f}, {center_y:.0f})")
                
                # 更新缓存
                self.last_success_coord = (center_x, center_y)
                self._cached_upload_element = result
            
        else:
            iprint(f"  [错误] 未找到上传图标")
            return False
        
        # 确保有file_chooser
        if not file_chooser:
            iprint(f"  [错误] 未能获取文件选择器")
            return False
        
        try:
            # 设置文件
            iprint("\n[阶段3] 设置文件")
            iprint("-"*40)
            iprint(f"  文件: {os.path.basename(file_path)}")
            
            await file_chooser.set_files(file_path)
            iprint("  [OK] 文件已添加到上传队列")
            
            # 等待上传完成
            iprint("\n[阶段4] 等待上传完成")
            iprint("-"*40)
            
            upload_success = await self.wait_for_upload_complete(file_path) 
            
            # 保存文件路径
            self.uploaded_file_path = file_path
            iprint("  [OK] 文件上传流程完成")
            
            await clear_all_highlights(self.page)
            return True
            
        except Exception as e:
            iprint(f"  [错误] 上传过程异常: {e}")
            await clear_all_highlights(self.page)
            return False

    async def find_input_box(self):
        """查找输入框，自动选择第一个"""
        quick_selectors = [
            'textarea[placeholder*="输入"]',
            'textarea[placeholder*="提问"]',
            'textarea:visible'
        ]
        
        for selector in quick_selectors:
            try:
                elements = await self.page.locator(selector).all()
                for elem in elements:
                    if await elem.is_visible():
                        box = await elem.bounding_box()
                        if box and box['width'] > 200:
                            return elem
            except:
                continue
        
        return None
        
    async def find_copy_buttons(self, input_element=None):
            """
            查找复制按钮 - 智能版本
            1. 优先使用历史记录（高效）
            2. 无历史记录时调用交互式方法（可靠）
            3. 找到后自动保存到历史记录
            """
            print("\n 进入find_copy_buttons...")
            
            # ===== 第一步：尝试使用历史记录 =====
            try:
                # 检查是否有历史记录
                best_coord = self.copy_finder.history.get_best_coord()
                if best_coord:
                    print(f"  [历史] 找到历史坐标 ({best_coord[0]:.0f}, {best_coord[1]:.0f})，尝试使用...")
                    
                    # 获取所有按钮
                    all_buttons = await self.copy_finder.find_all_buttons()
                    
                    # 在历史坐标附近查找
                    target_x, target_y = best_coord
                    search_radius = 80
                    
                    for btn in all_buttons:
                        try:
                            box = await btn.bounding_box()
                            if not box:
                                continue
                            
                            center_x = box['x'] + box['width'] / 2
                            center_y = box['y'] + box['height'] / 2
                            
                            # 检查是否在历史坐标附近
                            if (abs(center_x - target_x) < search_radius and 
                                abs(center_y - target_y) < search_radius):
                                
                                # 验证是否为复制按钮
                                button_info = await self.copy_finder.get_button_info(btn)
                                hover_info = await self.copy_finder.hover_and_check_tooltip(btn, duration=1)
                                is_copy = await self.copy_finder.is_copy_button(button_info, hover_info)
                                
                                if is_copy:
                                    print(f"  [历史] ✅ 在历史坐标附近找到复制按钮")
                                    
                                    # 调试模式下询问用户
                                    if DEBUG:
                                        resp = input(f"\n    是否使用此历史记录的复制按钮？(y/n): ").strip().lower()
                                        if resp == 'y':
                                            # 更新缓存
                                            self.copy_finder.history.update_cache(btn, box)
                                            return btn
                                    else:
                                        # 非调试模式直接使用
                                        self.copy_finder.history.update_cache(btn, box)
                                        return btn
                        except Exception as e:
                            print(f"    [历史] 检查按钮时出错: {e}")
                            continue
                    print("  [历史] 历史坐标附近未找到有效复制按钮")
            except Exception as e:
                print(f"  [历史] 历史记录查找失败: {e}")
            
            # ===== 第二步：无历史记录或历史记录失效，调用交互式方法 =====
            print("  [查找] 调用交互式查找...")
            result = await self.copy_finder.find_copy_buttons_interactive(input_element)
            
            # ===== 第三步：如果找到，保存到历史记录 =====
            if result:
                try:
                    box = await result.bounding_box()
                    if box:
                        # 保存坐标到历史记录
                        self.copy_finder.history.save(box['x'], box['y'], {'source': 'interactive'})
                        print(f"  [历史] 已保存新坐标 ({box['x']:.0f}, {box['y']:.0f})")
                        
                        # 更新缓存
                        self.copy_finder.history.update_cache(result, box)
                except Exception as e:
                    print(f"  [历史] 保存失败: {e}")
            
            return result

# ==================== 豆包 Handler ====================

class DoubaoHandler(BaseHandler):
    """豆包聊天平台处理器"""
    
    def __init__(self, page):
        super().__init__(page, "豆包")
        self.copy_finder = CopyButtonFinder(page, "豆包")
 
    # 在 DoubaoHandler 类中添加专门的侧边栏处理方法
    async def handle_doubao_sidebar(self):
        """专门处理豆包侧边栏 - 增强版"""
        print("\n[豆包侧边栏] 检查并处理...")
        
        try:
            # 等待侧边栏完全加载
            await asyncio.sleep(1)
            
            # 获取视口大小
            viewport = self.page.viewport_size
            if not viewport:
                viewport = {'width': 1280, 'height': 720}
            
            # 豆包侧边栏特定选择器
            sidebar_selectors = [
                'div[class*="sidebar"]',
                'div[class*="drawer"]',
                'div[class*="SlidePanel"]',
                'div[role="dialog"]',
                '.semi-modal',
                '.semi-drawer'
            ]
            
            # 关闭按钮特定选择器
            close_selectors = [
                'button[class*="close"]',
                'button[aria-label*="关闭"]',
                '.semi-button.semi-button-tertiary.semi-button-with-icon',
                'button.semi-button:has(svg)',
                '.semi-icon-close',
                'svg[class*="close"]',
                'button:has(svg[viewBox="0 0 14 14"])'
            ]
            
            # 先检查是否有侧边栏存在
            sidebar_exists = False
            for selector in sidebar_selectors:
                try:
                    if await self.page.locator(selector).first.is_visible(timeout=1000):
                        sidebar_exists = True
                        break
                except:
                    continue
            
            if not sidebar_exists:
                print("  [OK] 未检测到侧边栏")
                return True
            
            # 多次尝试关闭
            for attempt in range(3):
                print(f"  [尝试 {attempt + 1}/3]")
                closed = False
                
                # 方法1: 找关闭按钮点击
                for selector in close_selectors:
                    try:
                        buttons = await self.page.locator(selector).all()
                        for btn in buttons:
                            if not await btn.is_visible(timeout=500):
                                continue
                            
                            box = await btn.bounding_box()
                            if not box:
                                continue
                            
                            # 检查是否在右上角区域
                            if box['x'] > viewport['width'] - 300 and box['y'] < 150:
                                print(f"    找到关闭按钮 ({box['x']:.0f}, {box['y']:.0f})")
                                await btn.click()
                                print(f"    已点击")
                                await asyncio.sleep(1)
                                
                                # 验证是否关闭
                                still_exists = False
                                for s in sidebar_selectors:
                                    if await self.page.locator(s).first.is_visible(timeout=500):
                                        still_exists = True
                                        break
                                
                                if not still_exists:
                                    print(f"    [OK] 侧边栏已关闭")
                                    closed = True
                                    break
                    except:
                        continue
                    if closed:
                        break
                
                if closed:
                    break
                
                # 方法2: 直接点击右上角
                if not closed:
                    try:
                        click_x = viewport['width'] - 50
                        click_y = 50
                        print(f"    直接点击坐标 ({click_x}, {click_y})")
                        await self.page.mouse.click(click_x, click_y)
                        await asyncio.sleep(1)
                        
                        # 检查是否关闭
                        still_exists = False
                        for s in sidebar_selectors:
                            if await self.page.locator(s).first.is_visible(timeout=500):
                                still_exists = True
                                break
                        
                        if not still_exists:
                            print(f"    [OK] 侧边栏已关闭")
                            closed = True
                            break
                    except:
                        pass
                
                # 方法3: 按ESC键
                if not closed:
                    try:
                        print(f"    按ESC键")
                        await self.page.keyboard.press('Escape')
                        await asyncio.sleep(1)
                        
                        # 检查是否关闭
                        still_exists = False
                        for s in sidebar_selectors:
                            if await self.page.locator(s).first.is_visible(timeout=500):
                                still_exists = True
                                break
                        
                        if not still_exists:
                            print(f"    [OK] 侧边栏已关闭")
                            closed = True
                            break
                    except:
                        pass
                
                if not closed:
                    print(f"    第{attempt+1}次尝试未成功，继续...")
            
            # 最终检查
            sidebar_still_exists = False
            for selector in sidebar_selectors:
                if await self.page.locator(selector).first.is_visible(timeout=500):
                    sidebar_still_exists = True
                    break
            
            if sidebar_still_exists:
                print(f"  [警告] 侧边栏可能仍未完全关闭，执行页面刷新动作！")
                try:
                    await self.page.reload(wait_until='domcontentloaded', timeout=30000)
                    iprint("  [OK] 页面刷新成功")
                    await asyncio.sleep(3)  # 等待刷新后页面加载         
                except Exception as e:
                    iprint(f"  [WARN] 页面刷新失败: {e}")
                return True
            else:
                print(f"  [OK] 侧边栏处理完成")
                return True
                
        except Exception as e:
            print(f"  [错误] 处理侧边栏失败: {e}")
            return False
    
    async def find_input_box(self):
        """查找输入框，显示序号和高亮"""
        iprint(f"\n[{self.platform_name}] 定位输入框...")
        
        selectors = [
            'textarea[placeholder*="输入"]',
            'textarea[placeholder*="提问"]',
            '.chat-input textarea',
            '.input-area textarea',
            '[class*="input"] textarea',
            'div[contenteditable="true"]',
            '.semi-input-textarea'
        ]
        
        candidates = []
        for selector in selectors:
            elements = await self.page.locator(selector).all()
            for elem in elements:
                if await elem.is_visible():
                    box = await elem.bounding_box()
                    if box and box['width'] > 200 and box['height'] > 20:
                        candidates.append(elem)
                        dprint(f"    候选: {selector}")
        
        if not candidates:
            return None
        
        # 清除旧的高亮
        await clear_all_highlights(self.page)
        
        # 显示所有候选输入框的序号和高亮
        iprint(f"  找到 {len(candidates)} 个可能的输入框:")
        for i, elem in enumerate(candidates):
            await highlight_element_with_index(self.page, elem, i+1, 'blue', duration=1.5)
            await asyncio.sleep(0.3)
        
        # 自动选择第一个
        iprint(f"  [OK] 自动选择输入框 [1]")
        return candidates[0]

    async def find_copy_buttons(self, input_element=None):
            """查找复制按钮 - 豆包专用版"""
            try:              
                # 使用豆包专用的查找方法
                return await self.find_doubao_copy_buttons(input_element)
                
            except Exception as e:
                print(f"  [错误] 查找复制按钮失败: {e}")
                return None

    async def find_doubao_copy_buttons(self, input_element=None):
        """豆包专用复制按钮查找方法 - 只查找输入框左上方"""
        print(f"\n[豆包专用] 开始查找复制按钮...")
        
        # 先滚动页面到底部，确保最新内容可见
        await self.copy_finder.scroll_to_bottom()
        
        # 获取输入框位置
        if not input_element:
            print("  [错误] 没有输入框位置信息")
            return None
        
        try:
            input_box = await input_element.bounding_box()
            if not input_box:
                print("  [错误] 无法获取输入框位置")
                return None
            
            print(f"  [输入框位置] ({input_box['x']:.0f}, {input_box['y']:.0f}) 大小: {input_box['width']:.0f}x{input_box['height']:.0f}")
            
            # 定义搜索区域：输入框左上方区域
            # x: 输入框左侧300px到输入框左侧
            # y: 输入框上方400px到输入框顶部
            search_area = {
                'x_min': max(0, input_box['x'] - 300),  # 左侧300px内
                'x_max': input_box['x'],                 # 不超过输入框左侧
                'y_min': max(0, input_box['y'] - 400),   # 上方400px内
                'y_max': input_box['y']                   # 不超过输入框顶部
            }
            
            print(f"  [搜索区域] 左上方区域:")
            print(f"    X: {search_area['x_min']:.0f} -> {search_area['x_max']:.0f} (宽度: {search_area['x_max'] - search_area['x_min']:.0f}px)")
            print(f"    Y: {search_area['y_min']:.0f} -> {search_area['y_max']:.0f} (高度: {search_area['y_max'] - search_area['y_min']:.0f}px)")
            
            # 豆包特有的复制按钮选择器（缩小范围，只找小图标按钮）
            doubao_selectors = [
                'button[class*="copy"]',
                'button[class*="Copy"]',
                'button[aria-label*="copy"]',
                'button[title*="复制"]',
                'button:has(svg)',
                '.copy-button',
                '[class*="copy-btn"]',
                'button:has-text("复制")',
                'button:has(.semi-icon-copy)',
                'button.semi-button:has(svg[class*="copy"])',
                'button.w-fit',  # 根据您的打印信息，按钮有 w-fit 类
                '[class*="w-fit"]',  # 宽度自适应的元素
                'button.rounded-4',  # 圆角4px的按钮
                '[class*="rounded-4"]',  # 圆角4px的元素
                'button.flex.cursor-pointer',  # flex布局的按钮
                'button',  # 最后考虑所有按钮
            ]
            
            # 收集左上方区域的所有可能按钮
            area_buttons = []
            
            for selector in doubao_selectors:
                try:
                    elements = await self.page.locator(selector).all()
                    for elem in elements:
                        if not await elem.is_visible():
                            continue
                        
                        box = await elem.bounding_box()
                        if not box:
                            continue
                        
                        # 检查是否在左上方区域内
                        if (box['x'] + box['width']/2 >= search_area['x_min'] and
                            box['x'] + box['width']/2 <= search_area['x_max'] and
                            box['y'] + box['height']/2 >= search_area['y_min'] and
                            box['y'] + box['height']/2 <= search_area['y_max']):
                            
                            # 计算到输入框左上角的距离
                            btn_center_x = box['x'] + box['width']/2
                            btn_center_y = box['y'] + box['height']/2
                            distance = ((btn_center_x - input_box['x'])**2 + 
                                       (btn_center_y - input_box['y'])**2)**0.5
                            
                            # 获取元素信息用于判断
                            try:
                                # 尝试获取所有可能的提示信息
                                text = await elem.text_content() or ""
                                title = await elem.get_attribute('title') or ""
                                aria = await elem.get_attribute('aria-label') or ""
                                class_name = await elem.get_attribute('class') or ""
                                
                                # 鼠标悬浮，获取真正的提示（不使用高亮序号干扰）
                                try:
                                    # 先移除可能的高亮
                                    await self.page.evaluate("""
                                        () => {
                                            const labels = document.querySelectorAll('[id^="highlight-label-"]');
                                            labels.forEach(label => label.remove());
                                        }
                                    """)
                                    
                                    # 鼠标悬浮
                                    await elem.hover(timeout=5000)
                                    await asyncio.sleep(1)  # 等待提示出现
                                    
                                    # 获取悬浮提示（多种方式）
                                    tooltip_text = ""
                                    
                                    # 方法1: 通过 evaluate 获取可能的提示元素
                                    tooltip_text = await self.page.evaluate("""() => {
                                        // 查找常见的提示元素
                                        const selectors = [
                                            '[role="tooltip"]',
                                            '.tooltip',
                                            '[class*="tooltip"]',
                                            '[class*="Tooltip"]',
                                            '.semi-tooltip',
                                            '.ant-tooltip',
                                            'div[class*="popover"]',
                                            'div[class*="Popover"]'
                                        ];
                                        
                                        for (const selector of selectors) {
                                            const tooltips = document.querySelectorAll(selector);
                                            for (const tip of tooltips) {
                                                if (tip.offsetParent !== null) {
                                                    const text = tip.textContent || tip.innerText || '';
                                                    if (text.trim()) {
                                                        return text.trim();
                                                    }
                                                }
                                            }
                                        }
                                        return '';
                                    }""")
                                    
                                    # 方法2: 如果没找到，尝试从元素属性获取
                                    if not tooltip_text:
                                        # 检查鼠标悬浮后属性是否变化
                                        new_title = await elem.get_attribute('title') or ""
                                        new_aria = await elem.get_attribute('aria-label') or ""
                                        
                                        if new_title and new_title != title:
                                            tooltip_text = new_title
                                        elif new_aria and new_aria != aria:
                                            tooltip_text = new_aria
                                    
                                    # 方法3: 尝试从父元素获取提示
                                    if not tooltip_text:
                                        parent = await elem.evaluateHandle('el => el.parentElement')
                                        if parent:
                                            parent_title = await parent.get_attribute('title') or ""
                                            parent_aria = await parent.get_attribute('aria-label') or ""
                                            if parent_title:
                                                tooltip_text = parent_title
                                            elif parent_aria:
                                                tooltip_text = parent_aria
                                    
                                except Exception as e:
                                    print(f"    悬浮提示获取失败: {e}")
                                    tooltip_text = ""
                                
                                # 调试输出
                                print(f"\n  [候选元素] 位置: ({box['x']:.0f}, {box['y']:.0f}) 距离: {distance:.0f}px")
                                if text.strip():
                                    print(f"    文本: {text[:50]}")
                                if title:
                                    print(f"    title: {title}")
                                if aria:
                                    print(f"    aria-label: {aria}")
                                if class_name:
                                    print(f"    class: {class_name[:80]}")
                                if tooltip_text:
                                    print(f"    [悬浮提示] {tooltip_text}")
                                
                                # 综合判断是否为复制按钮
                                all_text = (text + " " + title + " " + aria + " " + tooltip_text).lower()
                                copy_keywords = ['复制', 'copy', '拷贝', 'content_copy', 'copy-icon', '📋', '⎘']
                                
                                is_copy = any(keyword in all_text for keyword in copy_keywords)
                                
                                if is_copy:
                                    print(f"    [OK] 检测到复制关键词!")
                                
                                area_buttons.append({
                                    'button': elem,
                                    'distance': distance,
                                    'box': box,
                                    'is_copy': is_copy,
                                    'tooltip': tooltip_text
                                })
                                
                            except Exception as e:
                                print(f"    获取元素信息失败: {e}")
                                continue
                                
                except Exception as e:
                    continue
            
            if not area_buttons:
                print("  [信息] 左上方区域未找到任何按钮")
                return None
            
            # 按距离排序（只找最近的）
            area_buttons.sort(key=lambda x: x['distance'])
            
            print(f"\n  在左上方区域找到 {len(area_buttons)} 个按钮")
            print(f"  其中疑似复制按钮: {sum(1 for b in area_buttons if b['is_copy'])} 个")
            
            # 只检查最近的6个按钮（避免被高亮序号干扰）
            max_buttons = min(6, len(area_buttons))
            print(f"  将检查最近的 {max_buttons} 个按钮...")

            for i, btn_info in enumerate(area_buttons[:max_buttons]):
                button = btn_info['button']
                
                if DEBUG:
                    print(f"\n  [按钮 {i+1}/{max_buttons}] 距离: {btn_info['distance']:.0f}px")
                    if btn_info.get('tooltip'):
                        print(f"    悬浮提示: {btn_info['tooltip']}")
                
                # 临时清除所有高亮
                await clear_all_highlights(self.page)
                
                # 使用绿色边框临时高亮（仅在调试模式）
                if DEBUG:
                    try:
                        await button.evaluate("""
                            el => {
                                el.style.outline = '3px solid green';
                                el.style.outlineOffset = '2px';
                            }
                        """)
                        await asyncio.sleep(1)
                        await button.evaluate("el => el.style.outline = ''")
                    except:
                        pass
                
                # 判断是否为复制按钮
                is_copy = btn_info.get('is_copy', False)
                
                if is_copy:
                    # 调试模式需要用户确认，非调试模式自动选择
                    if DEBUG:
                        resp = input(f"\n    是否使用此按钮？(y/n/q-退出): ").strip().lower()
                        if resp == 'y':
                            return button
                        elif resp == 'q':
                            return None
                    else:
                        # 非调试模式自动选择第一个检测到的复制按钮
                        return button
            
            print(f"\n  [信息] 未找到合适的复制按钮")
            return None
            
        except Exception as e:
            print(f"  [错误] 查找复制按钮失败: {e}")
            return None


# ==================== Kimi Handler ====================

class KimiHandler(BaseHandler):
    """Kimi聊天平台处理器"""
    
    def __init__(self, page):
        super().__init__(page, "Kimi")
        self.copy_finder = CopyButtonFinder(page, "Kimi")
  
    async def find_input_box(self):
        """查找输入框，显示序号和高亮"""
        iprint(f"\n[{self.platform_name}] 定位输入框...")
        
        selectors = [
            'textarea[placeholder*="输入"]',
            'textarea[placeholder*="提问"]',
            '.chat-input textarea',
            '.input-area textarea',
            '[contenteditable="true"]',
            '[role="textbox"]'
        ]
        
        candidates = []
        for selector in selectors:
            elements = await self.page.locator(selector).all()
            for elem in elements:
                if await elem.is_visible():
                    box = await elem.bounding_box()
                    if box and box['width'] > 200 and box['height'] > 20:
                        candidates.append(elem)
                        dprint(f"    候选: {selector}")
        
        if not candidates:
            return None
        
        # 清除旧的高亮
        await clear_all_highlights(self.page)
        
        # 显示所有候选输入框的序号和高亮
        iprint(f"  找到 {len(candidates)} 个可能的输入框:")
        for i, elem in enumerate(candidates):
            await highlight_element_with_index(self.page, elem, i+1, 'blue', duration=1.5)
            await asyncio.sleep(0.3)
        
        # 自动选择第一个
        iprint(f"  [OK] 自动选择输入框 [1]")
        return candidates[0]

    async def find_copy_buttons(self, input_element=None):
        """查找复制按钮，使用通用查找器"""
        return await self.copy_finder.find_copy_buttons_interactive(input_element)


# ==================== 填写输入框并提交 ====================

async def fill_and_submit(page, handler, input_element, form_data):
    """填写输入框并使用回车提交（自动完成，无需确认）"""
    if not input_element:
        iprint("[错误] 没有输入框可填写")
        return False, ""
    
    iprint("\n" + "="*60)
    iprint("开始填写并提交")
    iprint("="*60)
    
    if isinstance(form_data, dict):
        value = list(form_data.values())[0] if form_data.values() else ""
    else:
        value = form_data
    
    value_str = str(value)
    input_preview = value_str[:15] + "..." if len(value_str) > 15 else value_str
    
    iprint(f"\n[输入框] 准备填写: {input_preview}")
    
    try:
        await input_element.scroll_into_view_if_needed()
        await asyncio.sleep(0.5)
        
        await input_element.click()
        await asyncio.sleep(0.5)
        
        iprint(f"  开始输入 ({len(value_str)} 字符)...")
        
        fill_success = await handler.fill_input(input_element, value_str)
        if not fill_success:
            raise Exception("填写失败")
        
        filled_value = await handler.get_input_value(input_element)
        if filled_value:
            iprint(f"  [OK] 填写完成，长度: {len(filled_value)} 字符")
        else:
            iprint(f"  [WARN] 填写完成，但无法验证")
        
        await asyncio.sleep(0.5)
        await page.keyboard.press('Enter')
        iprint("  [OK] 已按回车提交")
        
        return True, value_str
        
    except Exception as e:
        iprint(f"  [ERR] 操作失败: {e}")
        return False, ""


# ==================== 点击复制按钮获取内容 ====================

async def click_copy_button_and_get_content(page, copy_button):
    """点击复制按钮并获取剪贴板内容"""
    iprint("\n" + "="*60)
    iprint("执行复制操作")
    iprint("="*60)
    
    if not copy_button:
        iprint("[错误] 没有找到复制按钮")
        return None
    
    try:
        pyperclip.copy('')
        iprint("[剪贴板] 已清空")
    except:
        pass
    
    await highlight_element_with_index(page, copy_button, "点击", 'green', duration=1)
    
    try:
        await copy_button.scroll_into_view_if_needed()
        await asyncio.sleep(0.5)
        await copy_button.click()
        iprint("  [OK] 已点击复制按钮")
        
        await asyncio.sleep(1)
        
        clipboard_content = get_clipboard_content()
        
        if clipboard_content:
            iprint(f"  [OK] 成功获取剪贴板内容，长度: {len(clipboard_content)} 字符")
            
            preview = clipboard_content[:200].replace('\n', ' ').replace('\r', '')
            if len(clipboard_content) > 200:
                preview += "..."
            iprint(f"\n[内容预览]")
            iprint("-" * 40)
            iprint(preview)
            iprint("-" * 40)
            
            return clipboard_content
        else:
            iprint("  [警告] 剪贴板内容为空")
            return None
            
    except Exception as e:
        iprint(f"  [ERR] 点击复制按钮失败: {e}")
        return None


# ==================== 数据加载 ====================

def load_input_data(file_path=None, data_string=None, is_base64=False, check_size_only=False, raw_content=False):
    """
    加载输入数据，支持Base64解码
    
    Args:
        file_path: 文件路径
        data_string: 数据字符串
        is_base64: 是否base64编码
        check_size_only: 如果为True，只返回数据大小，不返回完整数据
        raw_content: 如果为True，不添加OpenClaw提示词，直接返回原始内容
    
    Returns:
        如果check_size_only为True，返回数据大小
        否则返回处理后的数据
    """
    # Base64解码函数
    def decode_if_base64(content, is_base64):
        """如果is_base64为True，进行Base64解码"""
        if not is_base64:
            return content
        
        try:
            import base64
            decoded_bytes = base64.b64decode(content)
            decoded_content = decoded_bytes.decode('utf-8')
            return decoded_content
        except Exception as e:
            iprint(f"[Base64] 解码失败: {e}")
            return content
    
    # 优先使用data_string
    if data_string:
        # 如果是Base64编码，先解码
        if is_base64:
            data_string = decode_if_base64(data_string, is_base64)
        
        if check_size_only:
            return len(data_string)
        
        iprint("\n[数据加载] 从命令行参数读取")
        iprint(f"[数据加载] 参数长度: {len(data_string)} 字符")
        
        # 根据raw_content决定是否添加提示词
        if not raw_content:
            final_content = f"{CLAW2DS_PROMPT_DATA}\n{data_string}"
            iprint(f"[提示] 已添加OpenClaw请求报文Prompt")
        else:
            final_content = data_string
            iprint(f"[提示] 使用原始内容，不添加OpenClaw前缀")
        
        try:
            data = json.loads(final_content)
            iprint("[数据格式] JSON格式")
        except json.JSONDecodeError as e:
            iprint(f"[数据格式] 非JSON格式，作为文本处理: {e}")
            data = {"text": final_content}
            iprint("[数据格式] 文本格式")
        
        return data
    
    # 如果没有data_string但有file_path
    elif file_path and os.path.exists(file_path):
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read().strip()
            
            if check_size_only:
                return len(content)
            
            iprint(f"\n[数据加载] 从文件读取: {file_path}")
            iprint(f"[数据加载] 文件内容长度: {len(content)} 字符")
            
            # 如果是Base64编码，先解码
            if is_base64:
                content = decode_if_base64(content, is_base64)
            
            # 根据raw_content决定是否添加提示词
            if not raw_content:
                final_content = f"{CLAW2DS_PROMPT_FILE}\n{content}"
                iprint(f"[提示] 已添加OpenClaw请求报文文件Prompt")
            else:
                final_content = content
                iprint(f"[提示] 使用原始内容，不添加OpenClaw前缀")
            
            try:
                data = json.loads(final_content)
                iprint("[数据格式] JSON格式")
            except json.JSONDecodeError as e:
                iprint(f"[数据格式] 非JSON格式，作为文本处理: {e}")
                data = {"text": final_content}
                iprint("[数据格式] 文本格式")
            
            return data
            
        except Exception as e:
            iprint(f"[错误] 读取文件失败: {e}")
            return None
    
    return None if check_size_only else 0

# ==================== 主函数 ====================
async def main():
    parser = argparse.ArgumentParser(description='通用表单自动填写工具 - 最终交互版')
    
    # 平台选择参数（互斥）
    platform_group = parser.add_mutually_exclusive_group()
    platform_group.add_argument('--DS', '-DS', action='store_true', help='使用DeepSeek平台 (默认)')
    platform_group.add_argument('--DB', '-DB', action='store_true', help='使用豆包平台')
    platform_group.add_argument('--KM', '-KM', action='store_true', help='使用Kimi平台')
    
    # 原有的参数
    parser.add_argument('--url', '-u', help='目标页面URL（优先级高于平台简写）')
    parser.add_argument('--file', '-f', help='附件文件路径（用于上传给大模型）')
    parser.add_argument('--data', '-d', help='数据字符串（作为输入框的prompt）')
    
    # ===== 新增：直接文件内容作为prompt参数（与-d互斥）=====
    parser.add_argument('--df', '-df', dest='direct_file', help='直接读取文件内容作为输入框的prompt（不添加OpenClaw前缀）')
    
    parser.add_argument('--no-save', action='store_true', help='不保存响应结果')
    parser.add_argument('--debug', action='store_true', help='开启调试模式')
    parser.add_argument('--no-debug', action='store_true', help='强制关闭调试模式')
    parser.add_argument('--timeout', type=int, default=120, help='页面加载超时时间(秒)')

    # Base64解码支持
    parser.add_argument('--base64', action='store_true', help='输入数据是base64编码的（与-d、-f或-df配合使用）')
    
    # 文件上传相关参数
    parser.add_argument('--upload-only', action='store_true', help='只上传文件，不发送文本（仅DeepSeek平台）')
    
    args = parser.parse_args()
    
    # ===== 参数互斥检查 =====
    # -d 和 -df 不能同时使用
    if args.data and args.direct_file:
        parser.error("参数 --data/-d 和 --df/-df 不能同时使用")
    
    # 如果没有提供任何prompt源（-d 或 -df），则必须有 -f 参数（使用默认prompt）
    if not args.data and not args.direct_file and not args.file:
        parser.error("请提供输入内容：使用 --data/-d 或 --df/-df 提供prompt，或使用 --file/-f 提供附件文件（此时使用默认prompt）")
    
    global DEBUG
    if args.no_debug:
        DEBUG = False
    elif args.debug:
        DEBUG = True
    
    # 输出参数信息用于调试
    iprint("\n" + "="*60)
    iprint("命令行参数:")
    iprint(f"  --file/-f: {args.file if args.file else 'None'}")
    iprint(f"  --data/-d: {'[已提供]' if args.data else 'None'}")
    iprint(f"  --df/-df: {args.direct_file if args.direct_file else 'None'}")
    iprint(f"  --base64: {args.base64}")
    iprint(f"  --upload-only: {args.upload_only}")
    iprint(f"  --debug: {args.debug}")
    iprint(f"  --no-debug: {args.no_debug}")
    iprint("="*60 + "\n")
    
    # 确定目标URL
    target_url = args.url  # 优先使用用户指定的URL
    
    if not target_url:
        # 根据平台简写选择
        if args.DB:
            target_url = DB
            iprint(f"[平台选择] 使用豆包平台")
        elif args.KM:
            target_url = KM
            iprint(f"[平台选择] 使用Kimi平台")
        else:
            # 默认使用DeepSeek
            target_url = DS
            iprint(f"使用DeepSeek平台 (默认)")
    
    # 检查pyperclip
    try:
        pyperclip.copy('test')
        pyperclip.paste()
    except Exception as e:
        iprint("[警告] pyperclip初始化失败，请安装: pip install pyperclip")
        iprint(f"错误信息: {e}")
        return
    
    platform_name = get_platform_name(target_url)
    
    iprint("="*60)
    iprint(f"通用表单自动填写工具 - {platform_name}版")
    iprint("="*60)
    iprint(f"目标URL: {target_url}")
    iprint(f"调试模式: {'开启' if DEBUG else '关闭'}")
    iprint(f"加载超时: {args.timeout}秒")
    iprint("="*60)
    
    # ===== 修复：文件路径检查 =====
    # 获取当前工作目录
    current_dir = os.getcwd()
    iprint(f"[调试] 当前工作目录: {current_dir}")
    
    # 检查附件文件是否存在
    if args.file:
        # 先按原样检查
        if os.path.exists(args.file):
            iprint(f"[OK] 附件文件存在: {args.file}")
        else:
            # 尝试拼接当前目录
            full_path = os.path.join(current_dir, args.file)
            if os.path.exists(full_path):
                iprint(f"[修复] 附件文件路径修正: {full_path}")
                args.file = full_path
            else:
                iprint(f"[错误] 附件文件不存在: {args.file}")
                iprint(f"[尝试] 尝试路径: {full_path}")
                return
    
    # 检查prompt文件是否存在
    if args.direct_file:
        iprint(f"[调试] 检查prompt文件: {args.direct_file}")
        # 先按原样检查
        if os.path.exists(args.direct_file):
            iprint(f"[OK] prompt文件存在: {args.direct_file}")
        else:
            # 尝试拼接当前目录
            full_path = os.path.join(current_dir, args.direct_file)
            if os.path.exists(full_path):
                iprint(f"[修复] prompt文件路径修正: {full_path}")
                args.direct_file = full_path
            else:
                # 尝试使用绝对路径
                abs_path = os.path.abspath(args.direct_file)
                if os.path.exists(abs_path):
                    iprint(f"[修复] prompt文件使用绝对路径: {abs_path}")
                    args.direct_file = abs_path
                else:
                    iprint(f"[错误] prompt文件不存在: {args.direct_file}")
                    iprint(f"[尝试] 尝试路径: {full_path}")
                    iprint(f"[尝试] 绝对路径: {abs_path}")
                    return
    
    # 初始化浏览器
    browser_manager = BrowserManager()
    
    browser_connected = False
    if browser_manager.load_state():
        iprint("\n[浏览器] 尝试连接已保存的浏览器...")
        browser_connected = await browser_manager.connect_to_existing_browser()
    
    if not browser_connected:
        iprint("\n[浏览器] 启动新浏览器...")
        if not await browser_manager.launch_new_browser(target_url):
            iprint("[X] 浏览器启动失败")
            return
        iprint("[浏览器] 等待浏览器加载...")
        await asyncio.sleep(5)
    
    page = browser_manager.page
    
    # 导航
    current_url = page.url
    
    if target_url not in current_url:
        iprint(f"\n[导航] 当前URL: {current_url}，导航到目标页面...")
        try:
            await page.goto(target_url, wait_until='domcontentloaded', timeout=30000)
            iprint("导航成功")
        except Exception as e:
            iprint(f"导航超时: {e}")
        await asyncio.sleep(3)
    else:
        iprint(f"\n[刷新] 已在目标页面，执行刷新操作...")
        try:
            await page.reload(wait_until='domcontentloaded', timeout=30000)
            iprint("页面刷新成功")
            await asyncio.sleep(3)
        except Exception as e:
            iprint(f"页面刷新失败: {e}")
    
    await analyze_page(page, "初始页面")
    
    # 获取平台处理器
    handler = get_platform_handler(page, target_url)
    
    # 检查并处理登录
    logged_in = await handler.check_login_status()
    if not logged_in:
        iprint("[错误] 请登录")
        login_success = await handler.wait_for_login(timeout=300)
        if not login_success:
            iprint("[错误] 登录失败，退出")
            await browser_manager.close(keep_browser=True)
            return
        iprint("✓ 登录验证通过")

    # ===== 核心逻辑：根据参数组合确定操作模式 =====
    """
    参数逻辑说明：
    ----------------------------------------------------------------------
    1. -f 参数：附件文件（用于上传给大模型）
    2. -d 参数：数据字符串（作为输入框的prompt）
    3. -df 参数：prompt文件（文件内容作为输入框的prompt）
    
    关键关系：
    - -d 与 -df 互斥（不能同时使用）
    - -f 可以与其他参数任意组合
    
    详细逻辑：
    
    情况1：带 -f 参数（有附件上传）
        a) 有 -d 参数 → 使用 -d 的datastring作为输入框prompt
        b) 没有 -d 参数，但有 -df 参数 → 使用 -df 的文件内容作为输入框prompt
        c) 没有 -d 和 -df 参数 → 使用默认的 CLAW2DS_PROMPT_FILE 作为输入框prompt
    
    情况2：带 -d 参数（无附件上传）
        a) 有 -f 参数 → 见情况1
        b) 没有 -f 和 -df 参数 → 直接使用 -d 的datastring作为输入框prompt
    
    情况3：带 -df 参数（无附件上传）
        a) 有 -f 参数 → 见情况1
        b) 没有 -f 参数 → 直接使用 -df 的文件内容作为输入框prompt
    ----------------------------------------------------------------------
    """
    
    use_file_upload = False  # 是否使用附件上传模式
    text_data = None         # 输入框的prompt内容
    
    # ===== 第一步：判断是否使用附件上传模式 =====
    # 如果有 -f 参数，则使用附件上传模式
    if args.file:
        use_file_upload = True
        iprint(f"\n[模式识别] 附件上传模式 (使用文件: {args.file})")
        
        # ===== 第二步：确定附件上传模式下的prompt来源 =====
        # 情况1a: 有 -d 参数
        if args.data:
            iprint("[prompt来源] 使用 --data 参数内容")
            # 检查数据大小
            data_size = load_input_data(None, args.data, args.base64, check_size_only=True, raw_content=True)
            if data_size > FILE_UPLOAD_THRESHOLD:
                iprint(f"[警告] --data 内容较大 ({data_size} 字符)，可能影响上传性能")
            
            # 加载数据（raw_content=True 表示不添加OpenClaw前缀）
            text_data = load_input_data(None, args.data, args.base64, raw_content=True)
        
        # 情况1b: 没有 -d 参数，但有 -df 参数
        elif args.direct_file:
            iprint(f"[prompt来源] 使用 --df 文件内容: {args.direct_file}")
            # 检查文件大小
            file_size = load_input_data(args.direct_file, None, args.base64, check_size_only=True, raw_content=True)
            if file_size > FILE_UPLOAD_THRESHOLD:
                iprint(f"[警告] --df 文件内容较大 ({file_size} 字符)，可能影响上传性能")
            
            # 加载数据（raw_content=True 表示不添加OpenClaw前缀）
            text_data = load_input_data(args.direct_file, None, args.base64, raw_content=True)
        
        # 情况1c: 没有 -d 和 -df 参数
        else:
            iprint("[prompt来源] 使用默认 OpenClaw 提示词 (CLAW2DS_PROMPT_FILE)")
            text_data = {"text": CLAW2DS_PROMPT_FILE}
    
    # ===== 第三步：无附件上传的文本模式 =====
    else:
        iprint("\n[模式识别] 纯文本模式")
        
        # 情况2: 有 -d 参数
        if args.data:
            iprint("[prompt来源] 使用 --data 参数内容")
            # 检查数据大小
            data_size = load_input_data(None, args.data, args.base64, check_size_only=True, raw_content=True)
            if data_size > FILE_UPLOAD_THRESHOLD:
                iprint(f"[警告] --data 内容较大 ({data_size} 字符)，建议考虑使用附件上传模式")
            
            # 加载数据（raw_content=True 表示不添加OpenClaw前缀）
            text_data = load_input_data(None, args.data, args.base64, raw_content=True)
        
        # 情况3: 有 -df 参数
        elif args.direct_file:
            iprint(f"[prompt来源] 使用 --df 文件内容: {args.direct_file}")
            # 检查文件大小
            file_size = load_input_data(args.direct_file, None, args.base64, check_size_only=True, raw_content=True)
            if file_size > FILE_UPLOAD_THRESHOLD:
                iprint(f"[警告] --df 文件内容较大 ({file_size} 字符)，建议考虑使用附件上传模式")
            
            # 加载数据（raw_content=True 表示不添加OpenClaw前缀）
            text_data = load_input_data(args.direct_file, None, args.base64, raw_content=True)
        
        # 这种情况不应该发生（已经在参数检查中拦截）
        else:
            iprint("[错误] 未提供prompt内容")
            await browser_manager.close(keep_browser=True)
            return
    
    # ===== 根据模式执行相应操作 =====
    if use_file_upload:
        # 附件上传模式
        iprint(f"\n[调试] === 附件上传模式调试 ===")
        iprint(f"[调试] platform_name = '{platform_name}'")
        iprint(f"[调试] handler 类型 = {type(handler).__name__}")
        iprint(f"[调试] 是否有 upload_file_with_filechooser = {hasattr(handler, 'upload_file_with_filechooser')}")
        iprint(f"[调试] platform_name == 'DeepSeek' = {platform_name == 'DeepSeek'}")
        iprint(f"[调试] 条件结果 = {platform_name == 'DeepSeek' and hasattr(handler, 'upload_file_with_filechooser')}")
        if platform_name == "DeepSeek" and hasattr(handler, 'upload_file_with_filechooser'):
            iprint("\n" + "="*60)
            iprint("执行附件上传模式")
            iprint("="*60)
            
            # 上传附件文件
            upload_success = await handler.upload_file_with_filechooser(args.file)
            if not upload_success:
                iprint("[错误] 附件上传失败")
                await browser_manager.close(keep_browser=True)
                return
            
            iprint("\n[OK] 附件上传成功，准备发送prompt")
            
            # 查找输入框
            input_element = await handler.find_input_box()
            if not input_element:
                iprint("[错误] 没有找到输入框")
                await browser_manager.close(keep_browser=True)
                return
            
            # 聚焦输入框
            await handler.focus_input_box(input_element)
            
            # 填写prompt并提交
            submit_success, input_text = await fill_and_submit(page, handler, input_element, text_data)
            
            if not submit_success:
                iprint("[错误] 提交失败")
                await browser_manager.close(keep_browser=True)
                return
            
            # 等待响应
            iprint("\n[等待响应] 等待AI响应...")
            stable, content, length = await wait_for_page_stable_streaming(page, timeout=args.timeout)
            
            # 查找复制按钮
            copy_button = await handler.find_copy_buttons(input_element)
            
            if copy_button:
                clipboard_content = await click_copy_button_and_get_content(page, copy_button)
                if clipboard_content:
                    sys.stderr.write("[DEBUG] 成功获取剪贴板内容，准备保存到文件\n")
                    sys.stderr.flush()
                    
                    save_path = save_clipboard_to_file(clipboard_content, platform_name, input_text[:15])
                    
                    sys.stderr.write(f"[DEBUG] save_path: {save_path}\n")
                    sys.stderr.flush()
                    
                    if save_path:
                        sys.stderr.write(f"[DEBUG] 文件已保存，准备输出路径到 stdout\n")
                        sys.stderr.flush()
                        
                        # 统一使用 safe_output
                        safe_output(f"\n@@@PATH_START@@@{save_path}@@@PATH_END@@@\n")
                        safe_output(f"\n###RESULT_PATH###:{save_path}\n")
                        
                        sys.stderr.write("[DEBUG] 路径已输出到 stdout\n")
                        sys.stderr.flush()
                    else:
                        sys.stderr.write("[DEBUG] save_path 为空，文件保存失败\n")
                        sys.stderr.flush()
                else:
                    sys.stderr.write("[DEBUG] clipboard_content 为空\n")
                    sys.stderr.flush()
            else:
                sys.stderr.write("[DEBUG] 未找到复制按钮\n")
                sys.stderr.flush()
                print("", flush=True)
        
        else:
            iprint(f"[错误] 当前平台 {platform_name} 不支持附件上传模式")
            # 如果是 DeepSeek 但没有方法，给出更具体的错误
            if platform_name == "DeepSeek" and not hasattr(handler, 'upload_file_with_filechooser'):
                iprint(f"[调试] 处理器是 DeepSeekHandler 但缺少 upload_file_with_filechooser 方法")
            elif platform_name != "DeepSeek":
                iprint(f"[调试] 平台名称应该是 'DeepSeek'，但却是 '{platform_name}'")
    
    else:
        # 纯文本模式
        iprint("\n" + "="*60)
        iprint("执行纯文本模式")
        iprint("="*60)
        
        # 定位输入框
        input_element = await handler.find_input_box()
        
        if not input_element:
            iprint("[错误] 没有找到输入框，退出")
            await browser_manager.close(keep_browser=True)
            return
        
        # 聚焦输入框
        await handler.focus_input_box(input_element)
        
        # 填写prompt并提交
        submit_success, input_text = await fill_and_submit(page, handler, input_element, text_data)
        
        if not submit_success:
            iprint("[错误] 提交失败")
            await browser_manager.close(keep_browser=True)
            return
        
        # 等待响应
        stable, content, length = await wait_for_page_stable_streaming(page, timeout=args.timeout)
        
        # 如果是豆包平台，处理侧边栏
        if platform_name == "豆包":
            if hasattr(handler, 'handle_doubao_sidebar'):
                await handler.handle_doubao_sidebar()
        
        # 查找复制按钮
        copy_button = await handler.find_copy_buttons(input_element)
        
        if copy_button:
            clipboard_content = await click_copy_button_and_get_content(page, copy_button)
            if clipboard_content:
                sys.stderr.write("[DEBUG] 成功获取剪贴板内容，准备保存到文件\n")
                sys.stderr.flush()
                
                save_path = save_clipboard_to_file(clipboard_content, platform_name, input_text[:15])
                
                sys.stderr.write(f"[DEBUG] save_path: {save_path}\n")
                sys.stderr.flush()
                
                if save_path:                                    
                    safe_output(f"\n@@@PATH_START@@@{save_path}@@@PATH_END@@@\n")
                    safe_output(f"\n###RESULT_PATH###:{save_path}\n")
                    
                    sys.stderr.write("[DEBUG] 路径已输出到 stdout\n")
                    sys.stderr.flush()
                else:
                    sys.stderr.write("[DEBUG] save_path 为空，文件保存失败\n")
                    sys.stderr.flush()
            else:
                sys.stderr.write("[DEBUG] clipboard_content 为空\n")
                sys.stderr.flush()
        else:
            sys.stderr.write("[DEBUG] 未找到复制按钮\n")
            sys.stderr.flush()
            print("", flush=True)
    
    # 输出完成信息
    iprint("\n" + "="*60)
    iprint("操作完成！")
    iprint("="*60)
    iprint(f"浏览器调试端口: {browser_manager.cdp_port}")
    iprint(f"浏览器进程PID: {browser_manager.browser_pid}")
    iprint("浏览器保持运行，下次运行可复用")

    await browser_manager.close(keep_browser=True)
    await asyncio.sleep(0.5)
    return None


async def llmchat_resp(attachment_file=None, platform="DS", prompt_file=None, prompt_text=None, timeout=300):
    """
    通用大模型调用函数
    :param attachment_file: 附件文件路径（对应 -f 参数）
    :param platform: 大模型平台 "DS"/"DB"/"KM"
    :param prompt_file: prompt文件路径（对应 -df 参数）
    :param prompt_text: prompt文本内容（对应 -d 参数）
    :param timeout: 超时时间（秒）
    :return: 大模型响应的原始数据
    """
    iprint(f"\n{'='*60}")
    iprint(f"【大模型调用】平台: {platform}, 超时: {timeout}秒")
    if attachment_file:
        iprint(f"【附件文件】: {attachment_file}")
    if prompt_file:
        iprint(f"【prompt文件】: {prompt_file}")
    elif prompt_text:
        iprint(f"【prompt文本】: {prompt_text[:50]}...")
    iprint(f"{'='*60}")
    
    # 参数校验
    if prompt_file and prompt_text:
        raise ValueError("prompt_file 和 prompt_text 不能同时使用")
    
    if not attachment_file and not prompt_file and not prompt_text:
        raise ValueError("必须提供 attachment_file、prompt_file 或 prompt_text 之一")
    
    # 备份原始命令行参数
    sys_argv = sys.argv.copy()
    global DEBUG
    original_debug = DEBUG
    
    # 临时文件列表（用于清理）
    temp_files = []
    
    try:
        # 构建命令行参数
        sys.argv = [sys.argv[0]]
        
        # 平台参数
        if platform == "DS":
            sys.argv.append("--DS")
        elif platform == "DB":
            sys.argv.append("--DB")
        elif platform == "KM":
            sys.argv.append("--KM")
        
        # -f 参数：附件文件
        if attachment_file:
            sys.argv.extend(["--file", attachment_file])
        
        # ===== 关键修正：根据参数类型选择不同的参数 =====
        if prompt_file:
            # 情况：有 prompt 文件 → 使用 --df
            sys.argv.append("--df")
            sys.argv.append(prompt_file)
            iprint(f"[参数] 使用 --df: {prompt_file}")
            
        elif prompt_text:
            # 情况：有 prompt 文本 → 使用 --d（数据字符串）
            sys.argv.append("--d")
            sys.argv.append(prompt_text)
            iprint(f"[参数] 使用 --d: {prompt_text[:50]}...")
        
        # 超时参数
        sys.argv.extend(["--timeout", str(timeout)])
        
        # 强制关闭调试模式
        DEBUG = False
        
        # 捕获输出
        import io
        stdout_capture = io.StringIO()
        stderr_capture = io.StringIO()
        
        iprint("[调用] 执行原脚本主函数...")
        iprint(f"[命令行] {' '.join(sys.argv)}")
        
        original_stdout = sys.stdout
        original_stderr = sys.stderr
        start_time = time.time()
        
        try:
            sys.stdout = stdout_capture
            sys.stderr = stderr_capture
            await main()
        finally:
            sys.stdout = original_stdout
            sys.stderr = original_stderr
        
        elapsed = time.time() - start_time
        iprint(f"[耗时] 大模型调用完成，耗时: {elapsed:.1f}秒")
        
        DEBUG = original_debug
        output = stdout_capture.getvalue()
        
        # 解析输出找到结果文件
        import re
        result_file = None
        
        # 尝试多种路径标记
        path_match = re.search(r'@@@PATH_START@@@(.+?)@@@PATH_END@@@', output)
        if path_match:
            result_file = path_match.group(1).strip()
            iprint(f"[结果] 从PATH标记找到文件: {result_file}")
        
        if not result_file:
            path_match = re.search(r'###RESULT_PATH###:([^\n]+)', output)
            if path_match:
                result_file = path_match.group(1).strip()
                iprint(f"[结果] 从RESULT_PATH找到文件: {result_file}")
        
        if not result_file:
            import glob
            txt_files = glob.glob(os.path.join(LLMRSP_DIR, "*.txt"))
            if txt_files:
                txt_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
                result_file = txt_files[0]
                iprint(f"[结果] 使用最新生成的txt文件: {result_file}")
        
        if result_file and os.path.exists(result_file):
            with open(result_file, 'r', encoding='utf-8') as f:
                result_content = f.read()
            
            iprint(f"[结果] 成功读取响应文件，大小: {len(result_content)} 字符")
            
            try:
                return json.loads(result_content)
            except json.JSONDecodeError:
                return result_content
        else:
            iprint(f"[错误] 未找到生成的结果文件")
            return None
            
    except Exception as e:
        iprint(f"[错误] 调用过程异常: {e}")
        traceback.print_exc()
        return None
    finally:
        sys.argv = sys_argv
        DEBUG = original_debug
        
        # 清理临时文件（如果有）
        for temp_file in temp_files:
            try:
                if os.path.exists(temp_file):
                    os.unlink(temp_file)
                    iprint(f"[清理] 已删除临时文件: {temp_file}")
            except:
                pass

def save_optimized_to_file(data, file_path=None):
    """
    【新增】保存优化后的数据到文件
    :param data: 优化后的数据
    :param file_path: 指定文件路径（可选）
    :return: 保存的文件路径
    """
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    
    if file_path is None:
        file_path = os.path.join(LLMRSP_DIR, f"optimized_data_{timestamp}.json")
    
    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    
    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    iprint(f"[保存] 优化数据已保存到: {file_path}")
    return file_path

if __name__ == "__main__":
    # ...
    import logging
    # 彻底关闭 asyncio 的资源警告，防止干扰 stderr
    logging.getLogger('asyncio').setLevel(logging.CRITICAL)
    # 强制让剩下的垃圾回收静默
    sys.stderr = open(os.devnull, 'w')   
    
    try:
        asyncio.run(main())
    except Exception as e:
        # 不要用 original_print，直接用内置的 print
        # 加上 try 防止打印本身又崩了
        try:
            print(f"[错误] {e}")
        except:
            pass