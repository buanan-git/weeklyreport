#!/usr/bin/env python3
"""百度OKR周报自动提交脚本 - V2.4 (修复嵌套content结构)"""

import re
import asyncio
import socket
import json
import os
import traceback
import sys
from datetime import datetime, timedelta
from pathlib import Path
from playwright.async_api import async_playwright, Error as PlaywrightError
import platform
import time
import subprocess

# ==================== 导入浏览器自动化模块 ====================
# 添加共享库路径
_libs_path = Path(__file__).parent.parent.parent / "libs"
if _libs_path.exists():
    sys.path.insert(0, str(_libs_path))
# 打包环境：browser_automation.pyc 与本脚本同目录
_script_dir = str(Path(__file__).parent)
if _script_dir not in sys.path:
    sys.path.insert(0, _script_dir)

_BROWSER_AUTOMATION_LOADED = False
try:
    from browser_automation import (
        BrowserManager as _ExternalBrowserManager,
        set_debug as set_automation_debug,
        # 优化4、5、6: 新增工具
        find_first_available, find_all_parallel,
        SelectorPriorityManager, PreloadManager
    )
    _BROWSER_AUTOMATION_LOADED = True
    # 注：dprint 尚未定义，此处不打印（模块加载阶段）
except ImportError as e:
    print(f"[警告] 无法导入 browser_automation 模块: {e}", flush=True)
    print("[警告] 将使用脚本内建的简化版本", flush=True)

# ==================== 配置区域 ====================

# ==================== 新增：导入配置加载器 ====================
from config_loader import ConfigLoader

# 加载配置
config = ConfigLoader("submit_okr")

# ==================== 替换原有的硬编码配置 ====================
MY_ID = config.my_id
TARGET_WEEK = config.target_week #0=当周，1=上一周，2=上上周
DEBUG = config.debug
DRY_RUN = not config.save_intermediate
AUTO_SUBMIT = True

# 路径配置
BASE_DIR = config.script_dir
WEEKLYREPORT_DIR = config.get_path("weekly_report_dir")
BROWSER_STATE_FILE = config.get_browser_state_file()
BROWSER_DEBUG_DIR = config.get_path("browser_debug_dir")

# Chrome路径
CHROME_PATH = config.get_chrome_path()

# 调试端口范围（从config读取）
DEBUG_PORT_RANGE = config.debug_port_range

# 其他常量
BASE_URL = "https://okr.baidu-int.com/pages/weekly.html"


WEEKLY_DATA = {
    "part1_work_summary": None,
    "part2_okr_structure": None,
    "part3_metrics": None,
    "base_employee_name": None
}

# 定义 is_management_code 函数
def is_management_code(code: str) -> bool:
    """
    判断是否为管理类代码
    """
    if not isinstance(code, str):
        return False
    
    # 这里根据你的实际需求修改判断逻辑
    management_keywords = ['M', 'MGMT', 'ADMIN', 'LEAD', 'MANAGER']
    code_upper = code.upper()
    return any(keyword in code_upper for keyword in management_keywords)
# ==================== 工具函数 ====================

def iprint(msg):
    """智能打印，自动添加时间戳"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    print(f"[{timestamp}] {msg}")

# 设置DEBUG模式（从配置读取，默认关闭）
try:
    from config_loader import ConfigLoader
    _config = ConfigLoader("submit_okr")
    DEBUG = _config.debug
except:
    DEBUG = False

def dprint(*args, **kwargs):
    """调试输出，仅DEBUG模式显示"""
    if DEBUG:
        print(*args, **kwargs, flush=True)


# ==================== 快速输入函数 ====================

async def fast_type_text(page, text, element=None):
    """
    快速输入文本 - 使用JavaScript一次性插入，比逐字输入快100倍以上
    同时触发必要的事件，确保表单能够识别输入内容
    :param page: Playwright page对象
    :param text: 要输入的文本
    :param element: 目标元素（可选，如果不提供则使用当前焦点元素）
    """
    try:
        if not text:
            return True

        # 方法1：使用 document.execCommand('insertText') 并触发事件
        js_code = '''
            (el, text) => {
                el.focus();
                // 先插入文本
                document.execCommand('insertText', false, text);

                // 触发完整的事件链，确保Angular/React等框架能识别
                const events = [
                    new InputEvent('input', { bubbles: true, cancelable: true, data: text }),
                    new InputEvent('change', { bubbles: true, cancelable: true }),
                    new Event('keydown', { bubbles: true, cancelable: true, key: 'Insert', keyCode: 45 }),
                    new Event('keyup', { bubbles: true, cancelable: true, key: 'Insert', keyCode: 45 })
                ];

                // 触发input事件（最重要）
                el.dispatchEvent(events[0]);

                // 触发其他事件
                events.slice(1).forEach(evt => {
                    try { el.dispatchEvent(evt); } catch(e) {}
                });

                // 尝试触发原生事件（针对Angular等框架）
                try {
                    const nativeInputValueSetter = Object.getOwnPropertyDescriptor(el.tagName === 'INPUT' ? HTMLInputElement.prototype : HTMLTextAreaElement.prototype, 'value').set;
                    nativeInputValueSetter.call(el, text);
                    el.dispatchEvent(new Event('input', { bubbles: true }));
                } catch(e) {}
            }
        '''

        if element:
            await element.evaluate(js_code, text)
        else:
            await page.evaluate(js_code, text)

        return True
    except Exception as e:
        # 如果execCommand失败，回退到无延迟type
        dprint(f"  [fast_type] execCommand失败，使用快速type: {e}")
        try:
            await page.keyboard.type(text, delay=0)
            return True
        except Exception as e2:
            dprint(f"  [fast_type] 快速type也失败: {e2}")
            return False


def get_week_range(weeks_ago=0):
    """
    根据传入的参数获取指定周的日期范围（周一至周日）
    :param weeks_ago: 0=当周，1=上一周，2=上上周
    :return: 指定周的周一、周日datetime对象
    """
    today = datetime.now()
    # 计算基准周一（当周周一）
    base_monday = today - timedelta(days=today.weekday())
    # 根据传入的参数偏移周数
    target_monday = base_monday - timedelta(weeks=weeks_ago)
    target_sunday = target_monday + timedelta(days=6)
    return target_monday, target_sunday

def get_week_date(weeks_ago=0):
    """
    根据传入的参数获取指定周的标识日期（以周一为基准）
    :param weeks_ago: 0=当周，1=上一周，2=上上周
    :return: 格式化的日期字符串，如s2026-03-03
    """
    # 获取指定周的周一
    target_monday, _ = get_week_range(weeks_ago)
    return f"{target_monday.strftime('%Y-%m-%d')}"

def get_week_filename_suffix(weeks_ago=0):
    """
    根据传入的参数获取指定周的周报文件名后缀，格式为：2026.3.2-3.8
    :param weeks_ago: 0=当周，1=上一周，2=上上周
    :return: 格式化的文件名后缀字符串
    """
    monday, sunday = get_week_range(weeks_ago)
    
    # 提取年份、月、日信息
    year = monday.year  # 获取年份（如2026）
    start_month = monday.month  # 周一的月份
    start_day = monday.day  # 周一的日期
    end_month = sunday.month  # 周日的月份
    end_day = sunday.day  # 周日的日期
    
    # 按指定格式拼接（年份.月.日-月.日）
    return f"{year}.{start_month}.{start_day}-{end_month}.{end_day}.llm"

def get_target_week_report_filename(weeks_ago=0):
    """获取要填报的目标周报文件名"""
    week_suffix = get_week_filename_suffix(weeks_ago)
    return os.path.join(WEEKLYREPORT_DIR,f'团队周报_{week_suffix}.json')

# ==================== 浏览器管理器（适配器）====================
# 根据外部模块是否加载成功，选择使用外部类或内置类

if _BROWSER_AUTOMATION_LOADED:
    # 使用外部模块的适配器类
    class BrowserManager(_ExternalBrowserManager):
        """浏览器管理器适配器"""
        def __init__(self):
            super().__init__(
                config=config,
                chrome_path=CHROME_PATH,
                debug_dir=str(BROWSER_DEBUG_DIR),
                headless=False,
                screenshot_on_error=True,
                state_file=str(BROWSER_STATE_FILE)
            )

        # 方法别名 - 兼容旧接口
        async def connect_to_existing_browser(self):
            """兼容旧接口"""
            return await self.connect_to_existing()

        async def launch_new_browser(self, headless=False, url=None):
            """兼容旧接口"""
            if url is None:
                url = BASE_URL
            return await self.launch_new(url=url)

else:
    # 使用内置精简版类
    class BrowserManager:
        """浏览器管理器（精简版）"""
        def __init__(self):
            self.browser = None
            self.context = None
            self.page = None
            self.playwright = None
            self.cdp_port = None
            self.browser_pid = None
            self.browser_process = None

        def _is_port_in_use(self, port):
            """检查端口是否被占用"""
            import socket
            try:
                with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                    s.settimeout(1)
                    s.bind(('127.0.0.1', port))
                    return False
            except:
                return True

        def _get_available_port(self):
            """获取可用端口"""
            import random
            port_start, port_end = DEBUG_PORT_RANGE[0], DEBUG_PORT_RANGE[1]
            ports = list(range(port_start + 1, port_end + 1))
            random.shuffle(ports)
            for port in ports:
                if not self._is_port_in_use(port):
                    return port
            return random.randint(port_start + 1, port_end)

        async def _cleanup_playwright(self):
            """清理 playwright 资源"""
            try:
                if self.browser:
                    await self.browser.close()
            except: pass
            try:
                if self.playwright:
                    await self.playwright.stop()
            except: pass
            self.browser = self.context = self.page = self.playwright = None

        def save_state(self):
            if self.cdp_port and self.browser_pid:
                try:
                    state = {"cdp_port": self.cdp_port, "browser_pid": self.browser_pid, "timestamp": time.time()}
                    with open(BROWSER_STATE_FILE, "w", encoding="utf-8") as f:
                        json.dump(state, f)
                except: pass

        def load_state(self):
            if os.path.exists(BROWSER_STATE_FILE):
                try:
                    with open(BROWSER_STATE_FILE, "r", encoding="utf-8") as f:
                        state = json.load(f)
                    port = state.get("cdp_port")
                    pid = state.get("browser_pid")
                    # 验证进程和端口
                    if pid and self._is_running(pid):
                        if port and not self._is_port_in_use(port):
                            self.cdp_port = port
                            self.browser_pid = pid
                            return True
                except: pass
                # 清理无效状态文件
                try:
                    os.remove(BROWSER_STATE_FILE)
                except: pass
            return False

        def _is_running(self, pid):
            if not pid: return False
            try:
                if platform.system() == "Windows":
                    return str(pid) in subprocess.run(f'tasklist /FI "PID eq {pid}"', shell=True, capture_output=True, text=True).stdout
                os.kill(pid, 0)
                return True
            except: return False

        async def connect_to_existing_browser(self):
            if not self.cdp_port: return False
            await self._cleanup_playwright()
            try:
                self.playwright = await async_playwright().start()
                self.browser = await self.playwright.chromium.connect_over_cdp(f"http://127.0.0.1:{self.cdp_port}")
                self.context = self.browser.contexts[0] if self.browser.contexts else await self.browser.new_context()
                self.page = self.context.pages[0] if self.context.pages else await self.context.new_page()
                return True
            except:
                await self._cleanup_playwright()
                return False

        async def launch_new_browser(self, headless=False, url=None):
            if not CHROME_PATH: return False
            await self._cleanup_playwright()
            if url is None: url = BASE_URL
            self.cdp_port = self._get_available_port()
            cmd = [CHROME_PATH, f"--remote-debugging-port={self.cdp_port}", f"--user-data-dir={os.path.abspath(BROWSER_DEBUG_DIR)}", "--no-first-run", "--no-default-browser-check"]
            if headless: cmd.extend(["--headless=new"])
            cmd.append(url)
            try:
                self.browser_process = subprocess.Popen(cmd, shell=False, creationflags=subprocess.CREATE_NO_WINDOW if platform.system()=="Windows" else 0)
                self.browser_pid = self.browser_process.pid
                await asyncio.sleep(5)
                self.playwright = await async_playwright().start()
                for i in range(5):
                    try:
                        self.browser = await self.playwright.chromium.connect_over_cdp(f"http://127.0.0.1:{self.cdp_port}")
                        break
                    except: await asyncio.sleep(2) if i < 4 else None
                if not self.browser:
                    raise Exception("无法连接到浏览器")
                self.context = self.browser.contexts[0] if self.browser.contexts else await self.browser.new_context()
                self.page = self.context.pages[0] if self.context.pages else await self.context.new_page()
                self.save_state()
                return True
            except:
                await self._cleanup_playwright()
                if self.browser_process:
                    try:
                        self.browser_process.terminate()
                    except: pass
                    self.browser_process = None
                return False

        async def ensure_page(self):
            if not self.browser: return False
            self.context = self.browser.contexts[0] if self.browser.contexts else await self.browser.new_context()
            self.page = self.context.pages[0] if self.context.pages else await self.context.new_page()
            return True

        async def close(self, keep_browser=True):
            """关闭浏览器连接

            Args:
                keep_browser: True=保留浏览器进程供后续脚本复用，False=完全关闭浏览器
            """
            if keep_browser:
                # 保存状态，供下一个脚本复用
                self.save_state()
                # 只断开引用，不断开连接，让浏览器继续运行
                self.browser = self.context = self.page = self.playwright = None
            else:
                # 完全关闭：断开连接并终止浏览器进程
                await self._cleanup_playwright()
                if self.browser_process:
                    try:
                        self.browser_process.terminate()
                    except: pass
                    self.browser_process = None
                # 删除状态文件
                if os.path.exists(BROWSER_STATE_FILE):
                    try:
                        os.remove(BROWSER_STATE_FILE)
                    except: pass
                self.browser = self.context = self.page = self.playwright = None

        async def reconnect_with_timeout(self, max_retries=3):
            """Chrome 断连后尝试重连，最多重试指定次数

            Args:
                max_retries: 最大重试次数，默认3次
            """
            import time as _time
            start = _time.time()
            iprint(f"⚠️  浏览器连接断开，正在尝试重连（最多 {max_retries} 次）...")

            for attempt in range(1, max_retries + 1):
                elapsed = int(_time.time() - start)
                iprint(f"  重连尝试 #{attempt}/{max_retries}（已用时 {elapsed}s）...")

                await self._cleanup_playwright()

                # 尝试1: 通过已有 CDP 端口重连
                if self.cdp_port:
                    try:
                        self.playwright = await async_playwright().start()
                        self.browser = await self.playwright.chromium.connect_over_cdp(
                            f"http://127.0.0.1:{self.cdp_port}", timeout=10000)
                        self.context = self.browser.contexts[0] if self.browser.contexts else await self.browser.new_context()
                        self.page = self.context.pages[0] if self.context.pages else await self.context.new_page()
                        iprint(f"  ✅ 重连成功！（第{attempt}次尝试，耗时 {int(_time.time() - start)}s）")
                        return True
                    except:
                        await self._cleanup_playwright()

                # 尝试2: 从状态文件恢复
                if self.load_state() and self.cdp_port:
                    try:
                        self.playwright = await async_playwright().start()
                        self.browser = await self.playwright.chromium.connect_over_cdp(
                            f"http://127.0.0.1:{self.cdp_port}", timeout=10000)
                        self.context = self.browser.contexts[0] if self.browser.contexts else await self.browser.new_context()
                        self.page = self.context.pages[0] if self.context.pages else await self.context.new_page()
                        iprint(f"  ✅ 通过状态文件重连成功！（第{attempt}次尝试，耗时 {int(_time.time() - start)}s）")
                        return True
                    except:
                        await self._cleanup_playwright()

                if attempt < max_retries:
                    await asyncio.sleep(10)

            iprint(f"❌ 重连失败（已尝试{max_retries}次），程序将自动退出。")
            return False


# ==================== 格式化配置 ====================

FORMAT_CONFIG = {
    "part1_work_summary": {
        "template": "{project_name}: {work_desc} (进度: {progress}, 状态: {status})",
        "fallback": "{item}",
        "separator": "\n"
    },
    "part2_okr_structure": {
        "weekly_work_template": "已完成: {completed}\n进行中: {in_progress}\n阻塞: {blocked}",
        "fallback": "{weekly_work}",
        "include_progress": True,
        "include_confidence": True
    },
    "part3_metrics": {
        "sections": {
            "业务核心指标": "{metric_name}: 当前{current_value} (目标{target_value}, 变化{change})",
            "主要项目": "{project_name}: 进度{progress} ({status})",
            "下周重点工作": "[{priority}] {work_item} (负责人: {owner}, 截止: {deadline})"
        },
        "default": "{item}",
        "separator": "\n"
    }
}

# ==================== 解析函数 ====================

def format_item(item, template, fallback="{item}"):
    if isinstance(item, dict):
        try:
            return template.format(**item)
        except KeyError as e:
            dprint(f"      [WARN] 缺少字段: {e}")
            return ", ".join([f"{k}={v}" for k, v in item.items()])
    elif isinstance(item, str):
        return fallback.format(item=item)
    return str(item)

def flatten_nested_content(data, depth=0, max_depth=10):
    """递归扁平化嵌套结构，将嵌套的字典和列表转换为扁平列表"""
    if depth > max_depth:
        return [str(data)]

    if isinstance(data, list):
        result = []
        for item in data:
            result.extend(flatten_nested_content(item, depth + 1, max_depth))
        return result
    elif isinstance(data, dict):
        # 如果是字典，检查是否有特定字段
        if "content" in data:
            return flatten_nested_content(data["content"], depth + 1, max_depth)
        elif "items" in data:
            return flatten_nested_content(data["items"], depth + 1, max_depth)
        elif "data" in data:
            return flatten_nested_content(data["data"], depth + 1, max_depth)
        else:
            # 普通字典，保留为单个条目
            return [data]
    else:
        # 基本类型
        return [data] if data else []

def parse_part1(data):
    dprint(f"\n  [解析] 第一部分: 本周工作总结")

    # 处理直接是列表的情况
    if isinstance(data, list):
        dprint(f"    [INFO] 数据为列表类型，包含 {len(data)} 项")
        content_list = data
        title = "本周工作总结"
    elif isinstance(data, dict):
        title = data.get("title", "本周工作总结")
        content_list = data.get("content", [])
        # 如果content不是列表，尝试其他常见键
        if not isinstance(content_list, list):
            for key in ["items", "data", "work_items", "summary"]:
                if key in data and isinstance(data[key], list):
                    content_list = data[key]
                    dprint(f"    [INFO] 从 '{key}' 键提取内容")
                    break
        dprint(f"    标题: {title}")
        dprint(f"    content类型: {type(content_list)}")
    else:
        dprint(f"    [WARN] 不支持的数据类型: {type(data)}")
        return None

    # 扁平化嵌套结构
    content_list = flatten_nested_content(content_list)

    if not content_list:
        dprint(f"    [WARN] 内容为空")
        return None

    config = FORMAT_CONFIG["part1_work_summary"]
    formatted_items = []
    raw_items = []

    for idx, item in enumerate(content_list):
        dprint(f"    处理 item {idx+1}: {type(item)}")
        if isinstance(item, dict):
            raw_items.append(item)
            text = format_item(item, config["template"], config["fallback"])
            formatted_items.append(f"{idx+1}. {text}")
            dprint(f"      字段: {list(item.keys())}")
        elif isinstance(item, str):
            raw_items.append({"text": item})
            formatted_items.append(f"{idx+1}. {item}")
        else:
            # 其他类型转为字符串
            text = str(item)
            raw_items.append({"text": text})
            formatted_items.append(f"{idx+1}. {text}")

    result = {
        "title": title,
        "content": config["separator"].join(formatted_items),
        "raw_content": raw_items,
        "item_count": len(formatted_items)
    }

    dprint(f"    [OK] 完成: {result['item_count']} 条")
    dprint(f"    预览: {result['content'][:1000]}...")
    return result

def parse_part2(data):
    dprint(f"\n  [解析] 第二部分: OKR进度")

    # 处理直接是列表的情况（objectives列表）
    if isinstance(data, list):
        dprint(f"    [INFO] 数据为列表类型，包含 {len(data)} 个Objective")
        objectives = data
        title = "OKR进度"
    elif isinstance(data, dict):
        title = data.get("title", "OKR进度")
        objectives = data.get("objectives", [])
        # 如果objectives不是列表，尝试其他常见键
        if not isinstance(objectives, list):
            for key in ["items", "data", "okrs", "goals"]:
                if key in data and isinstance(data[key], list):
                    objectives = data[key]
                    dprint(f"    [INFO] 从 '{key}' 键提取objectives")
                    break
        dprint(f"    标题: {title}")
        dprint(f"    Objectives: {len(objectives)}")
    else:
        dprint(f"    [WARN] 不支持的数据类型: {type(data)}")
        return None

    if not isinstance(objectives, list):
        dprint(f"    [WARN] objectives 不是列表: {type(objectives)}")
        return None

    config = FORMAT_CONFIG["part2_okr_structure"]
    okr_data = {"title": title, "objectives": []}

    for obj_idx, obj in enumerate(objectives):
        if not isinstance(obj, dict):
            dprint(f"    [WARN] Objective {obj_idx} 不是字典: {type(obj)}")
            continue

        o_id = obj.get("o_id", f"O{obj_idx+1}")
        o_title = obj.get("o_title", obj.get("title", ""))
        dprint(f"    Objective: {o_id} - {o_title}")

        o_data = {"o_id": o_id, "o_title": o_title, "key_results": []}
        key_results = obj.get("key_results", [])

        # 如果key_results不是列表，尝试其他常见键
        if not isinstance(key_results, list):
            for key in ["krs", "items", "results", "data"]:
                if key in obj and isinstance(obj[key], list):
                    key_results = obj[key]
                    dprint(f"      [INFO] 从 '{key}' 键提取key_results")
                    break

        if not isinstance(key_results, list):
            dprint(f"      [WARN] key_results 不是列表: {type(key_results)}")
            key_results = [key_results] if key_results else []

        for kr_idx, kr in enumerate(key_results):
            if not isinstance(kr, dict):
                dprint(f"      [WARN] KR {kr_idx} 不是字典: {type(kr)}")
                continue

            kr_id = kr.get("kr_id", f"KR{kr_idx+1}")
            kr_title = kr.get("kr_title", kr.get("title", ""))
            weekly_work = kr.get("weekly_work", "")
            progress = kr.get("progress", "")
            confidence = kr.get("confidence", "")

            dprint(f"      KR: {kr_id}")

            # 处理weekly_work可能是嵌套结构的情况
            if isinstance(weekly_work, dict):
                dprint(f"        weekly_work字段: {list(weekly_work.keys())}")
                formatted_work = format_item(weekly_work, config["weekly_work_template"], config["fallback"])
            elif isinstance(weekly_work, list):
                # 如果是列表，递归处理
                flattened = flatten_nested_content(weekly_work)
                formatted_work = "\n".join([str(x) for x in flattened])
            elif isinstance(weekly_work, str):
                formatted_work = weekly_work
            else:
                formatted_work = str(weekly_work)

            if config["include_progress"] and progress:
                formatted_work += f"\n进度: {progress}"
            if config["include_confidence"] and confidence:
                formatted_work += f" | 信心: {confidence}"

            o_data["key_results"].append({
                "kr_id": kr_id,
                "kr_title": kr_title,
                "weekly_work": formatted_work,
                "raw_weekly_work": weekly_work,
                "progress": progress,
                "confidence": confidence
            })

        okr_data["objectives"].append(o_data)

    total_krs = sum(len(o["key_results"]) for o in okr_data["objectives"])
    dprint(f"    [OK] 完成: {len(okr_data['objectives'])} 个O, {total_krs} 个KR")
    return okr_data

# ==================== 修复第三部分数据解析 ====================

def parse_part3(data):
    """解析第三部分：关键指标 - 修复数据结构"""
    dprint(f"\n  [解析] 第三部分: 关键指标")

    # 处理直接是列表的情况（兼容旧格式）
    if isinstance(data, list):
        dprint(f"    [INFO] 数据为列表类型，包含 {len(data)} 项")
        data = {"title": "关键指标", "sections": data}
    
    # 如果不是字典，返回None
    if not isinstance(data, dict):
        dprint(f"    [WARN] 不支持的数据类型: {type(data)}")
        return None

    title = data.get("title", "关键指标")
    
    # 新格式：直接包含二级标题作为键
    # 例如：{"业务核心指标": [...], "主要项目": [...], "下周重点工作": [...]}
    sections = {}
    
    # 定义标准的二级标题
    standard_sections = ["业务核心指标", "主要项目", "下周重点工作"]
    
    for section_name in standard_sections:
        if section_name in data:
            content_list = data[section_name]
            dprint(f"    [OK] 找到 {section_name}: {len(content_list)} 项")
            
            # 确保是列表
            if not isinstance(content_list, list):
                content_list = [content_list] if content_list else []
            
            # 扁平化嵌套结构
            content_list = flatten_nested_content(content_list)
            
            # 格式化内容
            formatted_items = []
            raw_items = []
            
            for idx, item in enumerate(content_list):
                if isinstance(item, dict):
                    raw_items.append(item)
                    # 尝试格式化为字符串
                    text = ", ".join([f"{k}={v}" for k, v in item.items()])
                    formatted_items.append(f"{idx+1}. {text}")
                elif isinstance(item, str):
                    raw_items.append({"text": item})
                    formatted_items.append(f"{idx+1}. {item}")
                else:
                    text = str(item)
                    raw_items.append({"text": text})
                    formatted_items.append(f"{idx+1}. {text}")
            
            sections[section_name] = {
                "raw_content": raw_items,
                "formatted_content": "\n".join(formatted_items),
                "item_count": len(formatted_items)
            }
    
    # 兼容旧格式：如果上面没找到，尝试从sections字段获取
    if not sections and "sections" in data:
        old_sections = data.get("sections", [])
        if isinstance(old_sections, list):
            dprint(f"    [INFO] 使用旧格式sections列表")
            for idx, section in enumerate(old_sections):
                if isinstance(section, dict):
                    section_title = section.get("section_title", f"Section_{idx}")
                    content_list = section.get("content", [])
                    content_list = flatten_nested_content(content_list)
                    
                    formatted_items = []
                    raw_items = []
                    
                    for item_idx, item in enumerate(content_list):
                        if isinstance(item, dict):
                            raw_items.append(item)
                            text = ", ".join([f"{k}={v}" for k, v in item.items()])
                            formatted_items.append(f"{item_idx+1}. {text}")
                        elif isinstance(item, str):
                            raw_items.append({"text": item})
                            formatted_items.append(f"{item_idx+1}. {item}")
                        else:
                            text = str(item)
                            raw_items.append({"text": text})
                            formatted_items.append(f"{item_idx+1}. {text}")
                    
                    sections[section_title] = {
                        "raw_content": raw_items,
                        "formatted_content": "\n".join(formatted_items),
                        "item_count": len(formatted_items)
                    }

    dprint(f"\n    [统计] 共解析出 {len(sections)} 个部分:")
    for section_name, section_data in sections.items():
        dprint(f"      {section_name}: {section_data['item_count']} 项")

    metrics_data = {
        "title": title,
        "sections": sections
    }

    return metrics_data

def parse_json_file(json_path):
    dprint(f"\n{'='*60}")
    dprint("[DEBUG] 开始解析JSON文件")
    dprint(f"{'='*60}")

    dprint(f"\n[DEBUG] 检查文件: {json_path}")
    if not os.path.exists(json_path):
        print("  [ERROR] 文件不存在")
        try:
            files = os.listdir(os.path.dirname(json_path) or ".")
            dprint(f"  目录文件: {files}")
        except:
            pass
        return None

    try:
        with open(json_path, "r", encoding="utf-8") as f:
            raw_content = f.read()
        dprint(f"  [OK] 读取: {len(raw_content)} 字符")
        dprint(f"  前150字符: {raw_content[:150]}")
        dprint(f"  后150字符: {raw_content[-150:]}")
    except Exception as e:
        print(f"  [ERROR] 读取失败: {e}")
        return None

    try:
        data = json.loads(raw_content)
        dprint(f"\n[DEBUG] JSON解析成功")
        dprint(f"  顶层键: {list(data.keys())}")
    except json.JSONDecodeError as e:
        print(f"  [ERROR] JSON解析失败: {e}")
        dprint(f"  错误位置: 第 {e.lineno} 行, 第 {e.colno} 列 (字符 {e.pos})")
        
        # ===== 尝试修复截断的 JSON =====
        dprint(f"\n[DEBUG] 尝试修复截断的 JSON...")
        
        # 如果是 content 字符串被截断
        if "content" in raw_content:
            try:
                # 尝试提取完整的 content 字符串
                import re
                match = re.search(r'"content":\s*"(.+)"', raw_content, re.DOTALL)
                if match:
                    content_str = match.group(1)
                    # 检查是否以完整的 JSON 结尾
                    if content_str.endswith('"}'):
                        dprint(f"  [OK] 找到完整的 content 字符串")
                        # 尝试解析 content 字符串
                        try:
                            inner_data = json.loads(content_str)
                            dprint(f"  [OK] 成功解析 content 字符串")
                            
                            # 构造完整的数据
                            data = {"content": inner_data}
                            dprint(f"  [OK] 已修复 JSON 数据")
                        except json.JSONDecodeError as e2:
                            dprint(f"  [ERROR] content 字符串解析失败: {e2}")
                            return None
                    else:
                        dprint(f"  [WARN] content 字符串可能被截断")
                        # 尝试补全
                        fixed_content = content_str + '"}' 
                        try:
                            inner_data = json.loads(fixed_content)
                            dprint(f"  [OK] 补全后解析成功")
                            data = {"content": inner_data}
                        except:
                            dprint(f"  [ERROR] 补全后仍然解析失败")
                            return None
            except Exception as e2:
                dprint(f"  [ERROR] 修复失败: {e2}")
                return None
        else:
            return None

    # ===== 智能解析，兼容两种格式 =====
    final_content = None
    
    # 情况1: 直接是团队周报格式（有 report_type 字段，并且 content 是字典）
    if "report_type" in data and "content" in data and isinstance(data["content"], dict):
        dprint(f"\n[DEBUG] 情况1: 直接是团队周报格式")
        final_content = data["content"]  # 直接取 content 字典
        dprint(f"  从 content 字典中提取数据")
    
    # 情况2: 有 content 字段，且是字符串（大模型返回的格式）
    elif "content" in data and isinstance(data["content"], str):
        dprint(f"\n[DEBUG] 情况2: 发现 content 字符串字段")
        content_str = data["content"]
        dprint(f"  content字符串长度: {len(content_str)}")
        dprint(f"  content前100字符: {content_str[:100]}")
        dprint(f"  content后100字符: {content_str[-100:]}")
        
        try:
            # 解析 content 字符串为 JSON
            inner_data = json.loads(content_str)
            dprint(f"  [OK] 成功解析content字符串为JSON")
            dprint(f"  解析后的内层键: {list(inner_data.keys())}")
            
            # 检查内层是否有 content 字典
            if "content" in inner_data and isinstance(inner_data["content"], dict):
                dprint(f"  [INFO] 从内层JSON中提取 content 字典")
                final_content = inner_data["content"]
            else:
                dprint(f"  [WARN] 内层没有 content 字典，直接使用内层数据")
                final_content = inner_data
                
        except json.JSONDecodeError as e:
            print(f"  [ERROR] content字符串解析失败: {e}")
            dprint(f"  错误位置: 第 {e.lineno} 行, 第 {e.colno} 列")
            return None
    
    # 情况3: 其他未知格式，尝试找 content 字段
    elif "content" in data:
        dprint(f"\n[DEBUG] 情况3: 使用 content 字段")
        final_content = data["content"]
    
    else:
        dprint(f"\n[DEBUG] 情况4: 未知格式，使用原始数据")
        final_content = data

    if final_content is None:
        print(f"  [ERROR] 无法识别有效的数据格式")
        return None

    # 确保 final_content 是字典
    if not isinstance(final_content, dict):
        dprint(f"  [ERROR] final_content 不是字典类型: {type(final_content)}")
        return None

    dprint(f"\n{'='*60}")
    dprint("开始解析各部分内容")
    dprint(f"{'='*60}")

    result = {
        "part1_work_summary": None,
        "part2_okr_structure": None,
        "part3_metrics": None
    }

    # 从final_content中解析各部分
    if "part1_work_summary" in final_content:
        dprint(f"\n  [OK] 找到 part1_work_summary")
        result["part1_work_summary"] = parse_part1(final_content["part1_work_summary"])
    else:
        print(f"\n  [WARN] 未找到 part1_work_summary")
        dprint(f"  可用键: {list(final_content.keys())}")

    if "part2_okr_structure" in final_content:
        dprint(f"\n  [OK] 找到 part2_okr_structure")
        result["part2_okr_structure"] = parse_part2(final_content["part2_okr_structure"])
    else:
        print(f"\n  [WARN] 未找到 part2_okr_structure")

    if "part3_metrics" in final_content:
        dprint(f"\n  [OK] 找到 part3_metrics")
        result["part3_metrics"] = parse_part3(final_content["part3_metrics"])
    else:
        print(f"\n  [WARN] 未找到 part3_metrics")

    dprint(f"\n{'='*60}")
    dprint("[DEBUG] 解析结果总结")
    dprint(f"{'='*60}")
    dprint(f"  第一部分: {'OK' if result['part1_work_summary'] else 'FAIL'}")
    dprint(f"  第二部分: {'OK' if result['part2_okr_structure'] else 'FAIL'}")
    dprint(f"  第三部分: {'OK' if result['part3_metrics'] else 'FAIL'}")

    return result

# 修改 load_content_from_json 函数
def load_content_from_json():
    global WEEKLY_DATA

    dprint(f"\n{'='*60}")
    dprint("[系统] 加载JSON文件")
    dprint(f"{'='*60}")

    try:
        script_dir = Path(__file__).parent.absolute()
    except:
        script_dir = Path(os.getcwd())

    # 使用统一的文件名格式
    target_file = get_target_week_report_filename(TARGET_WEEK)
    json_path = script_dir / target_file
    dprint(f"文件路径: {json_path}")

    parsed = parse_json_file(str(json_path))

    if parsed:
        WEEKLY_DATA["part1_work_summary"] = parsed.get("part1_work_summary")
        WEEKLY_DATA["part2_okr_structure"] = parsed.get("part2_okr_structure")
        WEEKLY_DATA["part3_metrics"] = parsed.get("part3_metrics")

        # 从原始JSON文件中提取 base_employee.name（parse_json_file 只返回 content 部分）
        try:
            with open(str(json_path), "r", encoding="utf-8") as f:
                raw_data = json.loads(f.read())
            WEEKLY_DATA["base_employee_name"] = raw_data.get("base_employee", {}).get("name")
            if WEEKLY_DATA["base_employee_name"]:
                dprint(f"  员工姓名: {WEEKLY_DATA['base_employee_name']}")
        except Exception:
            pass

        dprint(f"\n[OK] 数据加载成功")
        dprint(f"\n[验证] 数据预览:")
        if WEEKLY_DATA["part1_work_summary"]:
            dprint(f"  第一部分: {len(WEEKLY_DATA['part1_work_summary']['content'])} 字符")
        if WEEKLY_DATA["part2_okr_structure"]:
            dprint(f"  第二部分: {len(WEEKLY_DATA['part2_okr_structure']['objectives'])} 个O")
        if WEEKLY_DATA["part3_metrics"]:
            dprint(f"  第三部分: {len(WEEKLY_DATA['part3_metrics']['sections'])} 个section")

        return True
    else:
        print(f"\n[ERROR] 加载失败")
        return False

# 添加提取项目数据的函数（用于后续填报）
def extract_projects_from_weekly_data():
    """
    从WEEKLY_DATA中提取需要填报的项目数据
    """
    projects = []
    
    part2 = WEEKLY_DATA.get("part2_okr_structure", {})
    objectives = part2.get("objectives", [])
    
    for o in objectives:
        for kr in o.get("key_results", []):
            title = kr.get("title", "")
            
            # 尝试提取项目代码
            code_match = re.search(r'(AC\d{5}|SQ\d{5}|ac\w+)', title, re.IGNORECASE)
            if code_match:
                project_code = code_match.group(0)
                project_name = title.replace(project_code, '').strip()
                
                projects.append({
                    'work_code': project_code,
                    'project_name': project_name,
                    'work_type': 'management' if is_management_code(project_code) else 'project',
                    'weekly_work': kr.get('weekly_work', [])
                })
                dprint(f"  [提取] 项目: {project_code} - {project_name}")
    
    return projects

def check_network_connection():
    print("\n[网络检测] 检查网络连接...")
    try:
        ip = socket.gethostbyname("okr.baidu-int.com")
        dprint(f"  [OK] 内网解析正常: {ip}")
        return True
    except:
        print(f"  [X] 内网解析失败，请检查VPN")
        return False

# ==================== 页面操作 ====================

async def analyze_page_structure(page):
    dprint("\n[页面分析] 分析页面结构...")
    structure = await page.evaluate("""() => {
        const result = {has_part1: false, has_part2: false, has_part3: false};
        document.querySelectorAll('h1, h2, h3, h4, [class*="title"]').forEach(h => {
            const t = h.textContent || '';
            if (t.includes('本周工作总结') || t.includes('工作总结')) result.has_part1 = true;
            if (t.includes('OKR') || t.includes('okr')) result.has_part2 = true;
            if (t.includes('指标') || t.includes('项目')) result.has_part3 = true;
        });
        return result;
    }""")
    dprint(f"  第一部分: {'存在' if structure['has_part1'] else '不存在'}")
    dprint(f"  第二部分: {'存在' if structure['has_part2'] else '不存在'}")
    dprint(f"  第三部分: {'存在' if structure['has_part3'] else '不存在'}")
    return structure

async def find_scroll_container(page):
    containers = await page.evaluate("""() => {
        const result = [];
        document.querySelectorAll("*").forEach(el => {
            const s = window.getComputedStyle(el);
            if (el.scrollHeight > el.clientHeight + 100 && el.scrollHeight > 500 &&
                (s.overflowY === 'auto' || s.overflowY === 'scroll')) {
                const r = el.getBoundingClientRect();
                result.push({class: el.className, id: el.id, tag: el.tagName, x: r.left + r.width/2, y: r.top + r.height/2});
            }
        });
        result.sort((a, b) => b.scrollHeight - a.scrollHeight);
        return result.slice(0, 1);
    }""")
    if containers:
        first = containers[0]
        selector = f"[class*='{first['class'].split()[0]}']" if first.get("class") else f"#{first['id']}" if first.get("id") else first["tag"].lower()
        return selector, (first.get("x", 960), first.get("y", 600))
    return "body", (960, 600)

async def scroll_container(page, selector, center):
    dprint("    滚动加载...", end=" ")
    await page.mouse.move(center[0], center[1])
    for i in range(15):  # 优化：30次 → 15次（每次滚屏更多）
        await page.mouse.wheel(0, 800)  # 优化：400 → 800
        await asyncio.sleep(0.15)  # 优化：0.2秒 → 0.15秒
        if i % 5 == 0:
            dprint(f"{i*800}px", end=" ")
    dprint("OK")
    dprint("    等待渲染...", end=" ")
    await asyncio.sleep(0.5)  # 优化：1秒 → 0.5秒
    dprint("OK")


# ==================== 模拟输入函数（备选方案） ====================

async def simulate_user_input(page, content):
    """模拟完整用户输入（备选方案）- 优化版"""

    dprint("\n[模拟输入] 开始模拟用户输入...")

    try:
        # 1. 找到本周工作总结标题并点击
        await page.evaluate("""() => {
            const headings = document.querySelectorAll('h3.mp-heading, h1.mp-heading');
            for (const h of headings) {
                if (h.textContent && h.textContent.includes('本周工作总结')) {
                    h.scrollIntoView({behavior: 'smooth', block: 'center'});
                    h.click();
                    break;
                }
            }
        }""")

        await asyncio.sleep(0.2)  # 优化：0.3秒 → 0.2秒  # 优化：0.5秒 → 0.3秒  # 优化：1秒 → 0.5秒

        # 2. 按回车创建新段落
        await page.keyboard.press('Enter')
        await asyncio.sleep(0.2)  # 优化：0.3秒 → 0.2秒  # 优化：0.5秒 → 0.3秒

        # 3. 逐字输入内容（带延迟，模拟真实输入）
        dprint("    开始逐字输入...")
        for i, char in enumerate(content):
            await page.keyboard.type(char)
            # 每10个字符稍微延迟一下
            if i % 10 == 0:
                await asyncio.sleep(0.05)  # 优化：0.1秒 → 0.05秒

        # 4. 按回车添加换行
        await page.keyboard.press('Enter')
        await asyncio.sleep(0.2)  # 优化：0.3秒 → 0.2秒

        # 5. 按Tab键移动焦点，触发blur
        await page.keyboard.press('Tab')
        await asyncio.sleep(0.2)  # 优化：0.3秒 → 0.2秒

        # 6. 点击页面其他地方，确保编辑器失去焦点
        await page.mouse.click(10, 10)

        dprint("  [OK] 模拟输入完成")

        # 7. 等待用户验证
        try:
            response = input("\n  模拟输入完成，请手动保存后输入 y 确认内容被保留 (y/n): ").strip().lower()
        except (EOFError, OSError, RuntimeError):
            response = "n"  # 无控制台时默认不确认
        return response == 'y'

    except Exception as e:
        print(f"  [X] 模拟输入失败: {e}")
        return False


# ==================== 第一部分填写函数 ====================

async def fill_part1_work_summary(page, part1_data):
    """填写第一部分：本周工作总结 - 动态定位版（自动模式）"""
    
    print("\n" + "="*80)
    print("[填写] 第一部分：本周工作总结（动态定位版 - 自动模式）")
    print("="*80)
    
    content = part1_data.get("content", "")
    if not content:
        print("  [X] 内容为空")
        return False
    
    # 清理不可见字符
    import re
    clean_content = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f\u200b]', '', content)
    
    dprint(f"  清理后内容长度: {len(clean_content)} 字符")
    dprint(f"  预览: {clean_content[:100]}...")
    
    try:
        # ==================== 第一步：滚动到标题位置 ====================
        dprint("\n" + "-"*60)
        dprint("  [步骤1] 滚动到标题位置")
        dprint("-"*60)

        # 清除可能存在的旧标记
        await page.evaluate("""() => {
            document.querySelectorAll('[style*="border"]').forEach(el => {
                el.style.border = '';
                el.style.backgroundColor = '';
            });
            document.querySelectorAll('.debug-marker').forEach(el => el.remove());
        }""")

        # 找到标题并滚动到视图 - 恢复到历史版本方式
        scroll_result = await page.evaluate("""() => {
            const headings = document.querySelectorAll('h1, h2, h3, .mp-heading');
            for (const h of headings) {
                if (h.textContent && h.textContent.includes('本周工作总结')) {
                    h.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    return { success: true, title: h.textContent };
                }
            }
            return { success: false };
        }""")

        if not scroll_result.get('success'):
            print("  [X] 未找到本周工作总结标题")
            return False

        dprint(f"  [OK] 已滚动到标题")
        await asyncio.sleep(1.0)  # 等待滚动完成（历史版本）

        # ==================== 第二步：获取精确位置并高亮 ====================
        dprint("\n" + "-"*60)
        dprint("  [步骤2] 获取精确位置")
        dprint("-"*60)
        
        title_info = await page.evaluate("""() => {
            const result = {
                success: false,
                clickX: 0,
                clickY: 0,
                viewX: 0,
                viewY: 0,
                hasTextElement: false,
                debug: ''
            };

            // 重新获取标题
            const headings = document.querySelectorAll('h1, h2, h3, .mp-heading');
            let titleElement = null;

            for (const h of headings) {
                if (h.textContent && h.textContent.includes('本周工作总结')) {
                    titleElement = h;
                    const rect = h.getBoundingClientRect();
                    result.debug = '标题位置: top=' + rect.top + ', bottom=' + rect.bottom + ', height=' + rect.height;
                    break;
                }
            }

            if (!titleElement) {
                result.debug = '未找到标题';
                return result;
            }

            // 高亮标题
            titleElement.style.border = '2px solid red';
            titleElement.style.backgroundColor = 'rgba(255, 200, 200, 0.3)';

            // 查找标题内的文字元素
            const textElements = titleElement.querySelectorAll('span[data-slate-string="true"]');

            if (textElements.length > 0) {
                // 使用最后一个文字元素 - 使用视口坐标（和历史版本一致）
                const lastText = textElements[textElements.length - 1];
                const textRect = lastText.getBoundingClientRect();

                // 点击位置：文字右侧15像素（避免太近选中文字）
                // 直接使用视口坐标，不需要滚动偏移
                result.clickX = textRect.right + 15;
                result.clickY = textRect.top + textRect.height / 2;
                result.hasTextElement = true;
                result.debug += '; 文字元素: top=' + textRect.top + ', bottom=' + textRect.bottom;

                // 高亮文字元素
                lastText.style.backgroundColor = 'rgba(255, 255, 0, 0.5)';
                lastText.style.border = '1px solid orange';
            } else {
                // 降级方案：使用标题元素本身
                const rect = titleElement.getBoundingClientRect();
                result.clickX = rect.right - 20;
                result.clickY = rect.top + rect.height / 2;
                result.debug += '; 使用降级方案';
            }

            // 添加绿色标记（更明显的标记）- 使用视口坐标
            const marker = document.createElement('div');
            marker.className = 'debug-marker';
            marker.style.position = 'fixed';
            marker.style.left = result.clickX + 'px';
            marker.style.top = result.clickY + 'px';
            marker.style.width = '16px';
            marker.style.height = '16px';
            marker.style.backgroundColor = '#00FF00';
            marker.style.border = '3px solid #008800';
            marker.style.borderRadius = '50%';
            marker.style.zIndex = '10000';
            marker.style.pointerEvents = 'none';
            marker.style.transform = 'translate(-50%, -50%)';
            marker.style.boxShadow = '0 0 15px #00FF00, 0 0 5px #FFFFFF';
            marker.style.animation = 'pulse 1s infinite';
            document.body.appendChild(marker);

            // 添加箭头指示 - 使用视口坐标
            const arrow = document.createElement('div');
            arrow.className = 'debug-marker';
            arrow.style.position = 'fixed';
            arrow.style.left = (result.clickX - 20) + 'px';
            arrow.style.top = (result.clickY - 20) + 'px';
            arrow.style.width = '0';
            arrow.style.height = '0';
            arrow.style.borderLeft = '10px solid transparent';
            arrow.style.borderRight = '10px solid transparent';
            arrow.style.borderBottom = '15px solid #00FF00';
            arrow.style.zIndex = '10000';
            arrow.style.pointerEvents = 'none';
            arrow.style.filter = 'drop-shadow(0 2px 2px black)';
            document.body.appendChild(arrow);
            
            result.success = true;
            return result;
        }""")
        
        if not title_info.get('success'):
            print("  [X] 无法获取标题位置")
            return False

        dprint(f"  [OK] 找到目标位置")
        dprint(f"  点击坐标: ({title_info['clickX']:.1f}, {title_info['clickY']:.1f})")
        dprint(f"  找到内部文字元素: {title_info.get('hasTextElement', False)}")
        dprint(f"  调试信息: {title_info.get('debug', 'N/A')}")

        # 注释掉用户交互确认
        # input("\n  请确认绿色闪烁圆点和箭头指向的位置，按Enter继续...")
        dprint("  [自动] 继续执行...")
        await asyncio.sleep(1.0)  # 大幅增加等待时间确保页面稳定

        # ==================== 第三步：动态定位并点击激活 ====================
        dprint("\n" + "-"*60)
        dprint("  [步骤3] 动态定位并点击激活")
        dprint("-"*60)

        await asyncio.sleep(0.5)  # 等待页面稳定
        dprint(f"  [INFO] 开始实时定位...")

        # 在点击前重新获取实时位置
        realtime_info = await page.evaluate("""() => {
            const result = {
                success: false,
                clickX: 0,       // 页面绝对坐标（用于点击）
                clickY: 0,
                viewX: 0,        // 视口坐标（用于标记显示）
                viewY: 0,
                debug: ''
            };

            // 1. 先找到"本周工作总结"标题，获取其位置范围
            let titleRect = null;
            const headings = document.querySelectorAll('h1, h2, h3, .mp-heading');
            for (const h of headings) {
                if (h.textContent && h.textContent.includes('本周工作总结')) {
                    titleRect = h.getBoundingClientRect();
                    result.debug += '找到标题 at ' + titleRect.top + '; ';
                    break;
                }
            }

            if (!titleRect) {
                result.debug += '未找到标题; ';
                return result;
            }

            // 标题底部位置（作为基准）
            const titleBottom = titleRect.bottom;
            const scrollX = window.scrollX || window.pageXOffset;
            const scrollY = window.scrollY || window.pageYOffset;

            // 2. 查找编辑器元素，但必须在标题下方范围内查找
            const editors = document.querySelectorAll('[data-slate-editor="true"], [contenteditable="true"]');
            for (const editor of editors) {
                const rect = editor.getBoundingClientRect();

                // 校验规则：
                // - 编辑器顶部必须在标题底部下方 >= 20px
                // - 编辑器顶部必须在标题底部下方 <= 60px
                const offsetFromTitle = rect.top - titleBottom;

                if (offsetFromTitle >= 20 && offsetFromTitle <= 60) {
                    // 使用编辑器顶部下方30px的位置作为输入点
                    result.viewX = rect.left + 50;  // 左侧50px
                    result.viewY = rect.top + 30;   // 顶部下方30px
                    result.clickX = result.viewX + scrollX;
                    result.clickY = result.viewY + scrollY;

                    result.success = true;
                    result.debug += '找到编辑器，偏移:' + offsetFromTitle + '; ';

                    // 临时高亮显示
                    const tempMarker = document.createElement('div');
                    tempMarker.className = 'temp-marker';
                    tempMarker.style.position = 'fixed';
                    tempMarker.style.left = result.viewX + 'px';
                    tempMarker.style.top = result.viewY + 'px';
                    tempMarker.style.width = '20px';
                    tempMarker.style.height = '20px';
                    tempMarker.style.backgroundColor = 'red';
                    tempMarker.style.border = '3px solid white';
                    tempMarker.style.borderRadius = '50%';
                    tempMarker.style.zIndex = '10001';
                    tempMarker.style.pointerEvents = 'none';
                    tempMarker.style.transform = 'translate(-50%, -50%)';
                    document.body.appendChild(tempMarker);

                    setTimeout(() => {
                        document.querySelector('.temp-marker')?.remove();
                    }, 3000);

                    break;
                } else {
                    result.debug += '编辑器偏移:' + offsetFromTitle + '不符合; ';
                }
            }

            // 3. 如果严格模式失败，放宽到标题下方20px-100px
            if (!result.success) {
                for (const editor of editors) {
                    const rect = editor.getBoundingClientRect();
                    const offsetFromTitle = rect.top - titleBottom;

                    if (offsetFromTitle >= 20 && offsetFromTitle <= 100) {
                        result.viewX = rect.left + 50;
                        result.viewY = rect.top + 30;
                        result.clickX = result.viewX + scrollX;
                        result.clickY = result.viewY + scrollY;

                        result.success = true;
                        result.debug += '放宽范围找到编辑器，偏移:' + offsetFromTitle + '; ';
                        break;
                    }
                }
            }

            return result;
        }""")

        if realtime_info.get('success'):
            dprint(f"  实时定位坐标: ({realtime_info['viewX']:.1f}, {realtime_info['viewY']:.1f}) [视口]")
            dprint(f"  页面绝对坐标: ({realtime_info['clickX']:.1f}, {realtime_info['clickY']:.1f})")
            dprint(f"  调试信息: {realtime_info.get('debug', 'N/A')}")
            click_x = realtime_info['clickX']
            click_y = realtime_info['clickY']
        else:
            dprint("  [⚠️] 实时定位失败，使用之前保存的坐标")
            dprint(f"  调试信息: {realtime_info.get('debug', 'N/A')}")
            click_x = title_info['clickX']
            click_y = title_info['clickY']

        # 确保元素在视图中 - 恢复到历史版本方式
        dprint(f"  [INFO] 滚动到目标位置...")
        await page.evaluate(f"window.scrollTo(0, {click_y - 200})")
        await asyncio.sleep(1.0)  # 等待滚动完成

        # 重新获取实时位置 - 使用视口坐标（和历史版本一致）
        dprint(f"  [INFO] 重新获取实时位置...")
        realtime_info = await page.evaluate("""() => {
            const result = { success: false, clickX: 0, clickY: 0 };

            // 重新获取标题
            const headings = document.querySelectorAll('h1, h2, h3, .mp-heading');
            for (const h of headings) {
                if (h.textContent && h.textContent.includes('本周工作总结')) {
                    // 查找标题内的文字元素
                    const textElements = h.querySelectorAll('span[data-slate-string="true"]');
                    if (textElements.length > 0) {
                        const lastText = textElements[textElements.length - 1];
                        const rect = lastText.getBoundingClientRect();
                        // 直接使用视口坐标
                        result.clickX = rect.right + 15;
                        result.clickY = rect.top + rect.height / 2;
                        result.success = true;

                        // 临时高亮显示实时位置
                        const tempMarker = document.createElement('div');
                        tempMarker.className = 'temp-marker';
                        tempMarker.style.position = 'fixed';
                        tempMarker.style.left = result.clickX + 'px';
                        tempMarker.style.top = result.clickY + 'px';
                        tempMarker.style.width = '20px';
                        tempMarker.style.height = '20px';
                        tempMarker.style.backgroundColor = 'red';
                        tempMarker.style.border = '3px solid white';
                        tempMarker.style.borderRadius = '50%';
                        tempMarker.style.zIndex = '10001';
                        tempMarker.style.pointerEvents = 'none';
                        tempMarker.style.transform = 'translate(-50%, -50%)';
                        document.body.appendChild(tempMarker);

                        setTimeout(() => {
                            document.querySelector('.temp-marker')?.remove();
                        }, 3000);
                    }
                    break;
                }
            }
            return result;
        }""")

        if realtime_info.get('success'):
            click_x = realtime_info['clickX']
            click_y = realtime_info['clickY']
            dprint(f"  实时定位坐标: ({click_x:.1f}, {click_y:.1f})")
        else:
            dprint(f"  [⚠️] 实时定位失败，使用之前保存的坐标")

        dprint(f"  [INFO] 最终点击坐标: ({click_x:.1f}, {click_y:.1f})")
        await asyncio.sleep(0.5)  # 确保稳定

        # 多次点击策略
        click_strategies = [
            ("单击", 1),
            ("双击", 2),
            ("三连击", 3)
        ]

        activated = False
        for strategy_name, click_count in click_strategies:
            dprint(f"    尝试{strategy_name} at ({click_x:.1f}, {click_y:.1f})...")
            for i in range(click_count):
                await page.mouse.click(click_x, click_y)
                await asyncio.sleep(0.2)  # 增加点击间隔
            await asyncio.sleep(0.5)  # 增加点击后等待时间
            
            # 验证是否激活 - 输入测试字符
            await page.keyboard.type('>', delay=30)
            await asyncio.sleep(0.1)
            
            has_test = await page.evaluate("""() => {
                const editor = document.querySelector('[data-slate-editor="true"]');
                return editor ? editor.textContent.includes('>') : false;
            }""")
            
            if has_test:
                dprint(f"    [OK] {strategy_name}成功激活")
                await page.keyboard.press('Backspace')
                await asyncio.sleep(0.1)
                activated = True
                break
        
        if not activated:
            dprint("  [⚠️] 自动点击未能激活")
            dprint("  尝试备用方案...")
            await asyncio.sleep(0.5)  # 增加等待时间

            # 备用方案：JavaScript直接聚焦
            await page.evaluate("""() => {
                const editor = document.querySelector('[data-slate-editor="true"]');
                if (editor) {
                    editor.focus();
                    editor.click();

                    // 尝试将光标移动到末尾
                    const range = document.createRange();
                    const selection = window.getSelection();
                    range.selectNodeContents(editor);
                    range.collapse(false);
                    selection.removeAllRanges();
                    selection.addRange(range);
                }
            }""")
            await asyncio.sleep(0.5)  # 增加等待时间
            
            # 再次测试
            await page.keyboard.type('>', delay=30)
            await asyncio.sleep(0.1)
            
            has_test = await page.evaluate("""() => {
                const editor = document.querySelector('[data-slate-editor="true"]');
                return editor ? editor.textContent.includes('>') : false;
            }""")
            
            if has_test:
                dprint("    [OK] JavaScript聚焦成功")
                await page.keyboard.press('Backspace')
                await asyncio.sleep(0.1)
                activated = True
            else:
                dprint("  [❌] 所有激活尝试失败，尝试继续...")
        
        # ==================== 第四步：验证光标 ====================
        dprint("\n" + "-"*60)
        dprint("  [步骤4] 验证光标")
        dprint("-"*60)
        
        # 输入测试字符
        await page.keyboard.type('测试', delay=50)
        await asyncio.sleep(0.2)  # 优化：0.3秒 → 0.2秒
        
        has_test = await page.evaluate("""() => {
            const editor = document.querySelector('[data-slate-editor="true"]');
            return editor ? editor.textContent.includes('测试') : false;
        }""")
        
        if has_test:
            dprint("  [OK] 光标已激活")
            # 删除测试字符
            for _ in range(2):
                await page.keyboard.press('Backspace')
                await asyncio.sleep(0.1)
        else:
            dprint("  [❌] 光标仍未激活，尝试继续...")
        
        # ==================== 第五步：插入空行 ====================
        dprint("\n" + "-"*60)
        dprint("  [步骤5] 插入空行")
        dprint("-"*60)

        # 输入2个空格
        dprint("    输入空格...")
        for i in range(2):
            await page.keyboard.press('Space')
            await asyncio.sleep(0.1)
        
        # 按两次回车
        dprint("    按两次回车...")
        await page.keyboard.press('Enter')
        await asyncio.sleep(0.2)  # 优化：0.3秒 → 0.2秒
        await page.keyboard.press('Enter')
        await asyncio.sleep(0.2)  # 优化：0.3秒 → 0.2秒
        
        # 按上箭头回到第一行
        await page.keyboard.press('ArrowUp')
        await asyncio.sleep(0.2)

        # 重新聚焦到编辑器元素，确保焦点正确
        await page.evaluate("""() => {
            const editor = document.querySelector('[data-slate-editor="true"]');
            if (editor) {
                editor.focus();
            }
        }""")
        await asyncio.sleep(0.1)

        # ==================== 第六步：输入内容 ====================
        dprint("\n" + "-"*60)
        dprint("  [步骤6] 输入内容")
        dprint("-"*60)

        dprint("    开始输入...")
        # 一次性粘贴全部内容（保持原有步骤和时序）
        dprint(f"      内容长度: {len(clean_content)} 字符")
        await page.keyboard.insert_text(clean_content)
        await asyncio.sleep(0.5)  # 等待粘贴完成

        await page.keyboard.press('Enter')
        await asyncio.sleep(0.3)  # 增加等待时间确保内容保存

        dprint("    [OK] 输入完成")

        # 清除标记
        await page.evaluate("""() => {
            document.querySelectorAll('.debug-marker, .temp-marker').forEach(el => el.remove());
            document.querySelectorAll('[style*="border"]').forEach(el => {
                el.style.border = '';
                el.style.backgroundColor = '';
            });
        }""")

        print("\n  [✅] 所有步骤完成")
        return True
        
    except Exception as e:
        print(f"\n  [X] 填写过程出错: {e}")
        import traceback
        traceback.print_exc()
        return False

# ==================== 第二部分填写函数（精确定位激活版） ====================

async def fill_okr_section(page, part2_data, base_employee_name=None):
    """填写第二部分：OKR进度 - 精确定位激活版（自动模式）"""
    
    print("\n" + "="*80)
    print("[填写] 第二部分：OKR进度（精确定位激活版 - 自动模式）")
    print("="*80)
    
    objectives = part2_data.get("objectives", [])
    if not objectives:
        print("  [X] 没有OKR数据")
        return False
    
    total_krs = sum(len(obj.get("key_results", [])) for obj in objectives)
    dprint(f"  共有 {len(objectives)} 个Objective, {total_krs} 个KR需要填写")
    
    # 定义正确的填写顺序
    fill_order = []
    for obj_idx, obj in enumerate(objectives):
        for kr_idx, kr in enumerate(obj.get("key_results", [])):
            weekly_work = kr.get("weekly_work", "")
            if isinstance(weekly_work, list):
                weekly_work = '\n'.join(weekly_work)
            fill_order.append({
                "o": obj_idx + 1,
                "kr": kr_idx + 1,
                "content": weekly_work
            })
    
    dprint("\n  [填写顺序]")
    for item in fill_order:
        if item['content']:
            dprint(f"    O{item['o']}KR{item['kr']}: {item['content'][:30]}...")
        else:
            dprint(f"    O{item['o']}KR{item['kr']}: [空内容]")
    
    try:
        # ==================== 第一步：定位OKR表格并获取输入框 ====================
        dprint("\n" + "-"*60)
        dprint("  [步骤1] 定位OKR表格")
        dprint("-"*60)

        # 预滚动：确保页面OKR区域完全加载（编辑模式可能有懒加载）
        await page.evaluate("""() => {
            const scrollable = document.querySelector('.mp-editor-content, .weekly-content, [class*="scroll"]') || document.documentElement;
            scrollable.scrollTop = scrollable.scrollHeight;
        }""")
        await asyncio.sleep(1.0)
        await page.evaluate("""() => {
            const scrollable = document.querySelector('.mp-editor-content, .weekly-content, [class*="scroll"]') || document.documentElement;
            scrollable.scrollTop = 0;
        }""")
        await asyncio.sleep(0.5)
        
        # 清除可能存在的旧标记
        await page.evaluate("""() => {
            document.querySelectorAll('[style*="border"]').forEach(el => {
                el.style.border = '';
                el.style.backgroundColor = '';
            });
            document.querySelectorAll('.debug-marker').forEach(el => el.remove());
        }""")
        
        # 定位OKR表格并高亮
        employee_name = base_employee_name or ''
        table_info = await page.evaluate("""(employeeName) => {
            const result = {
                success: false,
                message: '',
                tableFound: false,
                inputCount: 0,
                inputElements: [],
                debugInputs: [],
                anchorMethod: ''
            };

            let okrTitle = null;

            // ====== 策略1（主策略）：精确匹配 "XXX的OKR" ======
            if (employeeName) {
                const targetText = employeeName + '的OKR';
                result.message = '主策略: 搜索 "' + targetText + '"';
                const allEls = document.querySelectorAll('div, span, p, td, th, label, h1, h2, h3, h4, h5, section');
                for (const el of allEls) {
                    const text = (el.textContent || '').trim();
                    if (text === targetText) {
                        // 确认是叶子节点或直接包含目标文本的节点（排除大容器）
                        if (el.children.length === 0 || el.innerHTML.length < 100) {
                            okrTitle = el;
                            result.anchorMethod = 'exact_name_match';
                            el.scrollIntoView({ behavior: 'smooth', block: 'center' });
                            el.style.border = '3px solid blue';
                            el.style.backgroundColor = 'rgba(0, 0, 255, 0.1)';
                            break;
                        }
                    }
                }
            }

            // ====== 策略2：搜索 "OKR进度" 标题 ======
            if (!okrTitle) {
                const headings = document.querySelectorAll('h1, h2, h3, h4, h5, .mp-heading, [class*="title"], [class*="heading"], [class*="header"]');
                for (const h of headings) {
                    const text = (h.textContent || '').trim();
                    if (text.includes('OKR进度') && text.length < 200) {
                        okrTitle = h;
                        result.anchorMethod = 'okr_progress_heading';
                        h.scrollIntoView({ behavior: 'smooth', block: 'center' });
                        h.style.border = '3px solid blue';
                        h.style.backgroundColor = 'rgba(0, 0, 255, 0.1)';
                        break;
                    }
                }
            }

            if (!okrTitle) {
                // 策略2b：搜索任意包含"OKR进度"的元素
                const anyElements = document.querySelectorAll('div, span, p, td, th, label, section');
                for (const el of anyElements) {
                    const text = (el.textContent || '').trim();
                    if (text.includes('OKR进度') && text.length < 50) {
                        okrTitle = el;
                        result.anchorMethod = 'okr_progress_any';
                        el.style.border = '3px solid blue';
                        el.style.backgroundColor = 'rgba(0, 0, 255, 0.1)';
                        break;
                    }
                }
            }

            // ====== 策略3：搜索任意 /的OKR/ 模式 ======
            if (!okrTitle) {
                const anyElements2 = document.querySelectorAll('div, span, p, td, th, label');
                for (const el of anyElements2) {
                    const text = (el.textContent || '').trim();
                    if (/^.{1,10}的OKR$/.test(text)) {
                        okrTitle = el;
                        result.anchorMethod = 'generic_okr_pattern';
                        el.style.border = '3px solid blue';
                        el.style.backgroundColor = 'rgba(0, 0, 255, 0.1)';
                        break;
                    }
                }
            }

            // ====== 策略4：通过"本周工作"标签数量定位 ======
            if (!okrTitle) {
                const spans = document.querySelectorAll('span');
                let weeklyWorkSpans = [];
                for (const s of spans) {
                    if (s.textContent && s.textContent.trim() === '本周工作') {
                        weeklyWorkSpans.push(s);
                    }
                }
                if (weeklyWorkSpans.length >= 2) {
                    okrTitle = weeklyWorkSpans[0];
                    result.anchorMethod = 'weekly_work_spans';
                } else {
                    result.message = '未找到OKR进度标题（已尝试多种备用方案）';
                    return result;
                }
            }
            
            // 2. 找到包含多个"本周工作"标签的OKR表格容器
            //    核心判据：容器内必须有>=2个"本周工作"span（证明是OKR表格区域）
            let tableContainer = okrTitle;
            for (let i = 0; i < 15; i++) {
                tableContainer = tableContainer.parentElement;
                if (!tableContainer) break;

                const spans = tableContainer.querySelectorAll('span');
                let weeklyWorkCount = 0;
                spans.forEach(s => {
                    if (s.textContent && s.textContent.trim() === '本周工作') weeklyWorkCount++;
                });
                if (weeklyWorkCount >= 2) {
                    result.tableFound = true;
                    break;
                }
            }

            if (!result.tableFound) {
                // 备用方案：全局搜索包含"XXX的OKR"的容器
                const okrPattern = /的OKR/;
                const allElements = document.querySelectorAll('div, section');
                for (const el of allElements) {
                    if (el.textContent && okrPattern.test(el.textContent)) {
                        const spans = el.querySelectorAll('span');
                        let wc = 0;
                        spans.forEach(s => {
                            if (s.textContent && s.textContent.trim() === '本周工作') wc++;
                        });
                        if (wc >= 2) {
                            tableContainer = el;
                            result.tableFound = true;
                            break;
                        }
                    }
                }
            }

            if (!result.tableFound) {
                // 最终备用方案：使用整个页面body作为搜索范围
                tableContainer = document.body;
                const bodySpans = tableContainer.querySelectorAll('span');
                let weeklyWorkCount = 0;
                bodySpans.forEach(s => {
                    if (s.textContent && s.textContent.trim() === '本周工作') weeklyWorkCount++;
                });
                if (weeklyWorkCount >= 2) {
                    result.tableFound = true;
                    result.message += '（使用全页面搜索）';
                } else {
                    result.message = '未找到OKR表格';
                    return result;
                }
            }
            
            tableContainer.style.border = '3px solid red';
            tableContainer.style.backgroundColor = 'rgba(255, 0, 0, 0.05)';
            
            // 3. 在表格容器内查找"本周工作"输入区域
            //    回归历史稳定版本的简洁做法
            const allSpans = tableContainer.querySelectorAll('span');
            const processedInputs = new Set();
            let inputCounter = 0;

            // 诊断信息
            const editableSelectors = '[contenteditable="true"], .mp-paragraph-wrapper';
            let weeklyWorkSpanCount = 0;
            allSpans.forEach(s => {
                if (s.textContent && s.textContent.trim() === '本周工作') weeklyWorkSpanCount++;
            });
            result.diagnostics = {
                weeklyWorkLabelsFound: weeklyWorkSpanCount,
                editableElementsInContainer: tableContainer.querySelectorAll(editableSelectors).length,
                editableElementsOnPage: document.querySelectorAll(editableSelectors).length,
                containerTag: tableContainer.tagName,
                containerClass: tableContainer.className,
                containerInnerHTMLLength: tableContainer.innerHTML.length
            };

            allSpans.forEach((span) => {
                if (span.textContent && span.textContent.trim() === '本周工作') {
                    // 从"本周工作"标签向上遍历5层父元素，找第一个可编辑元素
                    let container = span;
                    for (let i = 0; i < 5; i++) {
                        container = container.parentElement;
                        if (!container) break;

                        const editable = container.querySelector('[contenteditable="true"], .mp-paragraph-wrapper');
                        if (editable && !processedInputs.has(editable)) {
                            processedInputs.add(editable);
                            inputCounter++;

                            const rect = editable.getBoundingClientRect();
                            const spanRect = span.getBoundingClientRect();

                            result.debugInputs.push({
                                index: inputCounter,
                                tagName: editable.tagName,
                                className: editable.className,
                                id: editable.id || '(无ID)',
                                parentTag: editable.parentElement ? editable.parentElement.tagName : '(无)',
                                parentClass: editable.parentElement ? editable.parentElement.className : '(无)',
                                containerClass: tableContainer.className,
                                spanText: span.textContent,
                                coordinates: { x: rect.left + 10, y: rect.top + 5, rawX: rect.left, rawY: rect.top },
                                spanCoordinates: { x: spanRect.left + 10, y: spanRect.top + 5, rawX: spanRect.left, rawY: spanRect.top }
                            });

                            result.inputElements.push({
                                span: span,
                                index: inputCounter,
                                x: rect.left + 10,
                                y: rect.top + 5,
                                rawY: rect.top,
                                spanX: spanRect.left + 10,
                                spanY: spanRect.top + 5
                            });

                            // 高亮
                            editable.style.border = '3px solid green';
                            editable.style.backgroundColor = 'rgba(0, 255, 0, 0.1)';
                            span.style.backgroundColor = 'rgba(255, 255, 0, 0.3)';
                            span.style.border = '1px solid orange';

                            const inputMarker = document.createElement('div');
                            inputMarker.className = 'debug-marker';
                            inputMarker.style.cssText = `position:fixed;left:${rect.left-30}px;top:${rect.top}px;background:green;color:white;padding:2px 6px;border-radius:3px;z-index:10000;font-size:12px;`;
                            inputMarker.textContent = '输入框 ' + inputCounter;
                            document.body.appendChild(inputMarker);

                            result.inputCount++;
                            break;  // 找到就停，不继续向上
                        }
                    }
                }
            });
            
            // 按从上到下的顺序排序
            result.inputElements.sort((a, b) => a.y - b.y);
            
            result.success = true;
            result.message = `找到OKR表格（${result.anchorMethod}），包含 ${result.inputCount} 个输入区域`;
            return result;
        }""", employee_name)
        
        if not table_info.get('success'):
            print(f"  [X] {table_info.get('message')}")
            return False

        dprint(f"  [OK] {table_info.get('message')}")
        print(f"  [定位] 锚点方法: {table_info.get('anchorMethod', '未知')}")
        dprint(f"  在OKR表格内找到输入区域: {table_info.get('inputCount', 0)} 个")

        # 输出诊断信息（尤其在检测失败时）
        diag = table_info.get('diagnostics', {})
        if diag:
            print(f"  [诊断] '本周工作'标签数量: {diag.get('weeklyWorkLabelsFound', '?')}")
            print(f"  [诊断] 容器内可编辑元素: {diag.get('editableElementsInContainer', '?')}")
            print(f"  [诊断] 全页面可编辑元素: {diag.get('editableElementsOnPage', '?')}")
            print(f"  [诊断] 容器: <{diag.get('containerTag', '?')}> class='{(diag.get('containerClass', '') or '')[:80]}'")
            print(f"  [诊断] 容器innerHTML长度: {diag.get('containerInnerHTMLLength', '?')}")
            for i, debug_info in enumerate(table_info.get('debugInputs', [])):
                coords = debug_info.get('coordinates', {})
                span_coords = debug_info.get('spanCoordinates', {})
                print(f"  [诊断] 输入框{i+1}: <{debug_info.get('tagName')}> class='{debug_info.get('className', '')[:40]}' pos=({coords.get('rawX', 0):.0f},{coords.get('rawY', 0):.0f}), 标签pos=({span_coords.get('rawX', 0):.0f},{span_coords.get('rawY', 0):.0f})")

        if table_info.get('inputCount') != total_krs:
            print(f"  [⚠️] 输入区域数量({table_info.get('inputCount')})与KR数量({total_krs})不匹配")
            try:
                response = input("  是否继续？(y/n): ").strip().lower()
            except (EOFError, OSError, RuntimeError):
                response = "y"  # 无控制台时默认继续
            if response != 'y':
                return False
        
        # 显示详细的输入框信息
        dprint("\n  [输入框详细信息]")
        for i, inp in enumerate(table_info['inputElements']):
            dprint(f"    输入框 {i+1}:")
            dprint(f"      输入框坐标: ({inp['x']:.0f}, {inp['y']:.0f})")
            dprint(f"      本周工作标签坐标: ({inp['spanX']:.0f}, {inp['spanY']:.0f})")
        
        dprint("  [自动] 继续执行...")
        await asyncio.sleep(0.2)  # 优化：0.3秒 → 0.2秒  # 优化：1秒 → 0.3秒

        # ==================== 第二步：逐个确认并填写（精确定位激活） ====================
        dprint("\n" + "-"*60)
        dprint("  [步骤2] 逐个确认并填写")
        dprint("-"*60)
        
        success_count = 0
        
        for order_idx, kr_info in enumerate(fill_order):
            if order_idx >= len(table_info['inputElements']):
                print(f"\n  [跳过] O{kr_info['o']}KR{kr_info['kr']}: 无对应输入框")
                continue
            
            # 如果没有内容，跳过
            if not kr_info['content']:
                print(f"\n  [跳过] O{kr_info['o']}KR{kr_info['kr']}: 内容为空")
                continue
            
            print(f"\n  [准备填写 O{kr_info['o']}KR{kr_info['kr']}]")
            dprint(f"  {'='*40}")
            
            # 获取当前要填写的输入框信息
            input_info = table_info['inputElements'][order_idx]
            weekly_work = kr_info['content']
            
            dprint(f"    目标输入框 {order_idx + 1}:")
            dprint(f"      输入框坐标: ({input_info['x']:.0f}, {input_info['y']:.0f})")
            dprint(f"      本周工作标签坐标: ({input_info['spanX']:.0f}, {input_info['spanY']:.0f})")
            dprint(f"    待填内容预览: {weekly_work[:60]}...")
            
            # 高亮当前要填写的输入框
            await page.evaluate(f"""() => {{
                // 清除之前的高亮
                document.querySelectorAll('.current-input-highlight').forEach(el => {{
                    el.classList.remove('current-input-highlight');
                    el.style.boxShadow = '';
                }});
                
                // 高亮当前输入框
                const inputs = document.querySelectorAll('[style*="border: 3px solid green"]');
                if (inputs[{order_idx}]) {{
                    const current = inputs[{order_idx}];
                    current.classList.add('current-input-highlight');
                    current.style.boxShadow = '0 0 0 4px yellow, 0 0 15px orange';
                    
                    // 添加箭头指示
                    const rect = current.getBoundingClientRect();
                    const arrow = document.createElement('div');
                    arrow.className = 'current-input-arrow';
                    arrow.style.position = 'fixed';
                    arrow.style.left = (rect.left - 40) + 'px';
                    arrow.style.top = (rect.top + rect.height/2 - 10) + 'px';
                    arrow.style.width = '0';
                    arrow.style.height = '0';
                    arrow.style.borderTop = '10px solid transparent';
                    arrow.style.borderBottom = '10px solid transparent';
                    arrow.style.borderRight = '20px solid yellow';
                    arrow.style.zIndex = '10001';
                    arrow.style.filter = 'drop-shadow(0 0 5px orange)';
                    document.body.appendChild(arrow);
                }}
            }}""")
            
            # ==================== 移除用户交互确认 ====================
            dprint("  [自动] 继续执行...")
            await asyncio.sleep(0.2)  # 优化：0.3秒 → 0.2秒  # 优化：0.5秒 → 0.3秒

            # ==================== 精确定位激活 ====================
            dprint("\n    [激活输入框] 尝试精确定位激活...")
            
            # 滚动到输入区域
            scroll_y = max(0, input_info['y'] - 200)
            await page.evaluate(f"window.scrollTo(0, {scroll_y})")
            await asyncio.sleep(0.2)  # 优化：0.3秒 → 0.2秒  # 优化：0.5秒 → 0.3秒
            
            activated = False
            
            # 首先尝试点击"本周工作"标签，这样可以确保焦点在正确的KR区域
            dprint("      点击本周工作标签...")
            await page.mouse.click(input_info['spanX'], input_info['spanY'])
            await asyncio.sleep(0.2)  # 优化：0.3秒 → 0.2秒
            
            # 然后点击输入框内部
            dprint("      点击输入框内部...")
            await page.mouse.click(input_info['x'], input_info['y'])
            await asyncio.sleep(0.2)  # 优化：0.3秒 → 0.2秒
            
            # 双击确保激活
            await page.mouse.dblclick(input_info['x'], input_info['y'])
            await asyncio.sleep(0.2)  # 优化：0.3秒 → 0.2秒
            
            # 验证是否激活 - 输入一个唯一的测试字符
            test_char = f'[{order_idx}]'
            await page.keyboard.type(test_char, delay=30)
            await asyncio.sleep(0.2)
            
            # 检查测试字符是否出现在编辑器中
            has_test = await page.evaluate(f"""() => {{
                const editor = document.querySelector('[data-slate-editor="true"]');
                return editor ? editor.textContent.includes('[{order_idx}]') : false;
            }}""")
            
            if has_test:
                dprint(f"      [OK] 输入框已激活，测试字符 '{test_char}' 已输入")
                # 删除测试字符
                for _ in range(len(test_char)):
                    await page.keyboard.press('Backspace')
                    await asyncio.sleep(0.1)
                activated = True
            else:
                dprint(f"      [⚠️] 测试字符未出现，可能焦点不在正确位置")

                # 备用方案：直接使用JavaScript聚焦
                dprint("      尝试JavaScript直接聚焦...")
                await page.evaluate(f"""() => {{
                    const inputs = document.querySelectorAll('[style*="border: 3px solid green"]');
                    if (inputs[{order_idx}]) {{
                        const input = inputs[{order_idx}];
                        input.focus();
                        input.click();
                        
                        // 创建并设置光标
                        const range = document.createRange();
                        const selection = window.getSelection();
                        range.selectNodeContents(input);
                        range.collapse(false);
                        selection.removeAllRanges();
                        selection.addRange(range);
                    }}
                }}""")
                await asyncio.sleep(0.2)  # 优化：0.3秒 → 0.2秒
                
                # 再次输入测试字符
                await page.keyboard.type(test_char, delay=30)
                await asyncio.sleep(0.2)
                
                has_test = await page.evaluate(f"""() => {{
                    const editor = document.querySelector('[data-slate-editor="true"]');
                    return editor ? editor.textContent.includes('[{order_idx}]') : false;
                }}""")
                
                if has_test:
                    dprint(f"      [OK] JavaScript聚焦成功")
                    for _ in range(len(test_char)):
                        await page.keyboard.press('Backspace')
                        await asyncio.sleep(0.1)
                    activated = True

            if not activated:
                dprint("      [❌] 无法自动激活，尝试继续...")
            
            # ==================== 插入空行 ====================
            dprint("\n    [插入空行] 准备输入位置...")

            # 输入2个空格
            for i in range(2):
                await page.keyboard.press('Space')
                await asyncio.sleep(0.1)

            # 按两次回车创建空行，挤出原有内容
            await page.keyboard.press('Enter')
            await asyncio.sleep(0.2)
            await page.keyboard.press('Enter')
            await asyncio.sleep(0.2)

            # 按上箭头回到第一行
            await page.keyboard.press('ArrowUp')
            await asyncio.sleep(0.2)

            # ==================== 输入内容 ====================
            dprint("\n    [输入内容] 开始输入...")
            # 一次性粘贴全部内容
            dprint(f"      内容长度: {len(weekly_work)} 字符")
            await page.keyboard.insert_text(weekly_work)
            await asyncio.sleep(0.5)  # 等待粘贴完成

            await page.keyboard.press('Enter')
            await asyncio.sleep(0.1)

            print(f"    [✅] O{kr_info['o']}KR{kr_info['kr']} 填写完成")
            success_count += 1
            
            # 移除箭头标记
            await page.evaluate("""() => {
                document.querySelector('.current-input-arrow')?.remove();
            }""")
            
            # 自动继续，无需用户确认
            dprint("  [自动] 继续下一个输入框...")
            await asyncio.sleep(0.2)  # 优化：0.3秒 → 0.2秒  # 优化：0.5秒 → 0.3秒
        
        print(f"\n  [填写统计] {success_count}/{total_krs} 个KR填写成功")

        # ==================== 清除标记 ====================
        dprint("\n" + "-"*60)
        dprint("  [步骤4] 清除标记")
        dprint("-"*60)
        
        await page.evaluate("""() => {
            document.querySelectorAll('[style*="border"]').forEach(el => {
                el.style.border = '';
                el.style.backgroundColor = '';
            });
            document.querySelectorAll('.debug-marker, .current-input-arrow').forEach(el => el.remove());
        }""")
        
        dprint("  [OK] 标记已清除")
        
        return success_count > 0
        
    except Exception as e:
        print(f"\n  [X] 填写过程出错: {e}")
        import traceback
        traceback.print_exc()
        return False

# ==================== 辅助函数：在OKR表格下方创建新section ====================

async def _create_section_below_okr(page, section_name):
    """滚动到编辑器底部（跳出OKR表格），动态创建section标题并定位光标到内容输入位置"""
    dprint(f"\n  [创建] 在OKR表格下方创建 '{section_name}' ...")

    # 滚动到编辑器最底部，确保跳出OKR表格区域
    await page.evaluate("""() => {
        const editor = document.querySelector('[data-slate-editor="true"]');
        if (editor) {
            editor.scrollTop = editor.scrollHeight;
        }
        window.scrollTo(0, document.body.scrollHeight - 500);
    }""")
    await asyncio.sleep(1)

    # 激活编辑器末尾光标
    await page.evaluate("""() => {
        const editor = document.querySelector('[data-slate-editor="true"]');
        if (editor) {
            editor.focus();
            editor.click();

            const range = document.createRange();
            const selection = window.getSelection();
            range.selectNodeContents(editor);
            range.collapse(false);
            selection.removeAllRanges();
            selection.addRange(range);
        }
    }""")
    await asyncio.sleep(0.2)

    # 创建新段落和标题
    dprint("    创建新段落...")
    for i in range(3):
        await page.keyboard.press('Enter')
        await asyncio.sleep(0.2)

    dprint(f"    输入标题: {section_name}")
    await page.keyboard.type(section_name, delay=30)
    await page.keyboard.press('Enter')
    await asyncio.sleep(0.2)
    await page.keyboard.press('Enter')
    await asyncio.sleep(0.2)

    dprint("    光标已在标题下方")
    await asyncio.sleep(0.2)


# ==================== 第三部分填写函数 ====================

async def fill_part3_metrics(page, part3_data):
    """
    填写第三部分：关键指标 - 优化版（基于成功经验）
    核心改进：
    1. 移除用户交互确认（改为自动执行）
    2. 优先使用 JavaScript 直接聚焦（策略3）
    3. 添加更精确的标题下方定位
    4. 保留调试信息但改为非交互式
    """
    print("\n" + "="*80)
    print("[填写] 第三部分：关键指标（优化版 - 基于成功经验）")
    print("="*80)
    
    # ==================== 调试函数（非交互式）====================
    async def debug_editor_state(context=""):
        """调试打印当前编辑器状态（仅输出，不交互）"""
        try:
            editor_info = await page.evaluate("""() => {
                const editor = document.querySelector('[data-slate-editor="true"]');
                if (!editor) return { exists: false };
                
                const selection = window.getSelection();
                let hasFocus = false;
                let cursorPosition = null;
                
                if (selection && selection.rangeCount > 0) {
                    const range = selection.getRangeAt(0);
                    const activeElement = document.activeElement;
                    hasFocus = activeElement === editor || editor.contains(activeElement);
                    
                    const text = range.startContainer.textContent || '';
                    cursorPosition = {
                        container: range.startContainer.nodeName,
                        offset: range.startOffset,
                        textAround: text.substring(Math.max(0, range.startOffset - 20), 
                                                   Math.min(text.length, range.startOffset + 20))
                    };
                }
                
                const paragraphs = editor.querySelectorAll('.mp-paragraph-wrapper, [class*="paragraph"]');
                
                return {
                    exists: true,
                    hasFocus: hasFocus,
                    activeElement: document.activeElement ? document.activeElement.tagName : 'none',
                    editorClass: editor.className,
                    paragraphCount: paragraphs.length
                };
            }""")
            
            dprint(f"\n  [编辑器{context}]")
            if editor_info.get('exists'):
                dprint(f"    聚焦: {editor_info.get('hasFocus', False)}")
                dprint(f"    激活: {editor_info.get('activeElement', 'unknown')}")
                dprint(f"    段落: {editor_info.get('paragraphCount', 0)}")
        except Exception as e:
            dprint(f"    调试失败: {e}")
    
    # ==================== 开始主要逻辑 ====================
    
    sections = part3_data.get("sections", {})
    if not sections:
        print("  [X] 没有指标数据")
        return False
    
    section_names = ["业务核心指标", "主要项目", "下周重点工作"]
    total_sections = len([s for s in section_names if s in sections and sections[s].get('formatted_content')])
    dprint(f"\n  [信息] 需要填写 {total_sections} 个部分")
    
    # 数据预览（缩短长度）
    dprint("\n  [数据预览]")
    for section_name in section_names:
        if section_name in sections:
            section_data = sections[section_name]
            content = section_data.get("formatted_content", "")
            if not content:
                content = section_data.get("raw_content", "")
                if isinstance(content, list):
                    content = "\n".join([str(item) for item in content])
            preview = content[:100] + "..." if len(content) > 100 else content
            dprint(f"    {section_name}: {len(content)} 字符")
            dprint(f"      预览: {preview}")
    
    try:
        # ==================== 第一步：定位各个区域 ====================
        dprint("\n" + "-"*60)
        dprint("  [步骤1] 定位各个区域")
        dprint("-"*60)
        
        # 清除标记
        await page.evaluate("""() => {
            document.querySelectorAll('[style*="border"]').forEach(el => {
                el.style.border = '';
                el.style.backgroundColor = '';
            });
            document.querySelectorAll('.debug-marker').forEach(el => el.remove());
        }""")
        
        # 查找各个标题的位置
        section_info = await page.evaluate("""() => {
            const result = {
                success: false,
                sections: {},
                foundTitles: []
            };
            
            const targetTitles = ['业务核心指标', '主要项目', '下周重点工作'];
            const headings = document.querySelectorAll('h1, h2, h3, .mp-heading');
            
            targetTitles.forEach(title => {
                for (const h of headings) {
                    if (h.textContent && h.textContent.includes(title)) {
                        result.foundTitles.push(title);
                        
                        const rect = h.getBoundingClientRect();
                        
                        // 查找该标题下方的输入区域
                        let inputArea = null;
                        let nextElement = h.nextElementSibling;
                        
                        for (let i = 0; i < 5; i++) {
                            if (!nextElement) break;
                            
                            const editable = nextElement.querySelector('[contenteditable="true"], .mp-paragraph-wrapper');
                            if (editable) {
                                inputArea = editable;
                                break;
                            }
                            
                            if (nextElement.matches && nextElement.matches('[contenteditable="true"], .mp-paragraph-wrapper')) {
                                inputArea = nextElement;
                                break;
                            }
                            
                            nextElement = nextElement.nextElementSibling;
                        }
                        
                        // 记录关键坐标 - 添加滚动偏移
                        const scrollX = window.scrollX || window.pageXOffset;
                        const scrollY = window.scrollY || window.pageYOffset;
                        result.sections[title] = {
                            title: {
                                rect: {
                                    left: rect.left,
                                    top: rect.top,
                                    bottom: rect.bottom,
                                    right: rect.right
                                },
                                // 标题右侧点击点（适合点击新增区域）
                                clickX: (rect.right - 20) + scrollX,
                                clickY: (rect.top + rect.height / 2) + scrollY,
                                // 标题下方点击点
                                belowX: (rect.left + 50) + scrollX,
                                belowY: (rect.bottom + 20) + scrollY
                            },
                            inputArea: inputArea ? {
                                rect: {
                                    left: inputArea.getBoundingClientRect().left,
                                    top: inputArea.getBoundingClientRect().top
                                },
                                // 输入区域点击点 - 添加滚动偏移
                                clickX: (inputArea.getBoundingClientRect().left + 50) + (window.scrollX || window.pageXOffset),
                                clickY: (inputArea.getBoundingClientRect().top + 10) + (window.scrollY || window.pageYOffset)
                            } : null
                        };
                        break;
                    }
                }
            });
            
            result.success = true;
            return result;
        }""")
        
        # 显示找到的区域
        dprint("\n  [找到的区域]")
        for section_name in section_names:
            if section_name in section_info['sections']:
                dprint(f"    ✅ {section_name}")
            else:
                dprint(f"    ❌ {section_name}: 未找到")
        
        await asyncio.sleep(1)
        
        # ==================== 第二步：逐个填写 ====================
        dprint("\n" + "-"*60)
        dprint("  [步骤2] 逐个填写")
        dprint("-"*60)
        
        success_count = 0
        
        for section_idx, section_name in enumerate(section_names):
            if section_name not in sections:
                print(f"\n  [跳过] {section_name}: 无数据")
                continue
            
            section_data = sections[section_name]
            content = section_data.get("formatted_content", "")
            if not content:
                content = section_data.get("raw_content", "")
                if isinstance(content, list):
                    content = "\n".join([str(item) for item in content])
            
            if not content:
                print(f"\n  [跳过] {section_name}: 内容为空")
                continue
            
            print(f"\n  [准备填写 {section_name} ({section_idx+1}/{total_sections})]")
            
            # 检查页面是否有该标题
            if section_name in section_info['sections']:
                dprint(f"  [信息] 页面已有 '{section_name}' 标题")
                section_info_item = section_info['sections'][section_name]
                
                # ==================== 激活输入区域（基于成功经验）====================
                dprint(f"\n    [激活] 准备填写 {section_name}...")
                
                # 滚动到标题位置
                title_y = section_info_item['title']['rect']['top']
                await page.evaluate(f"window.scrollTo(0, {max(0, title_y - 200)})")
                await asyncio.sleep(1)
                
                # 策略3（成功策略）：JavaScript直接聚焦
                dprint("      使用JavaScript直接聚焦（基于成功经验）...")
                
                # 记录聚焦前状态
                await debug_editor_state("(JS聚焦前)")
                
                # JavaScript强制聚焦到正确位置
                focus_success = await page.evaluate(f"""() => {{
                    // 查找目标区域的编辑器
                    const findTargetEditor = () => {{
                        // 方法1: 通过标题查找
                        const headings = document.querySelectorAll('h1, h2, h3, .mp-heading');
                        for (const h of headings) {{
                            if (h.textContent && h.textContent.includes('{section_name}')) {{
                                // 查找标题下方的可编辑区域
                                let next = h.nextElementSibling;
                                for (let i = 0; i < 5; i++) {{
                                    if (!next) break;
                                    const editable = next.querySelector('[contenteditable="true"], .mp-paragraph-wrapper');
                                    if (editable) {{
                                        console.log('找到标题下方的可编辑区域');
                                        return editable;
                                    }}
                                    if (next.matches && next.matches('[contenteditable="true"], .mp-paragraph-wrapper')) {{
                                        console.log('标题本身可编辑');
                                        return next;
                                    }}
                                    next = next.nextElementSibling;
                                }}
                            }}
                        }}
                        
                        // 方法2: 标题存在但下方无可编辑区域 → 需要在OKR表格外动态创建
                        // 不再盲目选择最后一个可编辑元素（那通常是OKR KR输入框）
                        console.log('标题存在但下方无可编辑区域，需要动态创建');
                        return null;
                    }};

                    const target = findTargetEditor();
                    if (!target) return 'create';
                    
                    // 强制聚焦
                    target.focus();
                    target.click();
                    
                    // 尝试将光标移动到末尾
                    const range = document.createRange();
                    const selection = window.getSelection();
                    
                    // 尝试找到目标区域内的最后一个段落
                    const paragraphs = target.querySelectorAll('.mp-paragraph-wrapper, [class*="paragraph"]');
                    if (paragraphs.length > 0) {{
                        // 定位到最后一个段落的末尾
                        const lastPara = paragraphs[paragraphs.length - 1];
                        const textNodes = [];
                        const walk = document.createTreeWalker(lastPara, NodeFilter.SHOW_TEXT);
                        let node;
                        while (node = walk.nextNode()) {{
                            textNodes.push(node);
                        }}
                        if (textNodes.length > 0) {{
                            const lastTextNode = textNodes[textNodes.length - 1];
                            range.setStart(lastTextNode, lastTextNode.length);
                            range.setEnd(lastTextNode, lastTextNode.length);
                        }} else {{
                            range.selectNodeContents(lastPara);
                            range.collapse(false);
                        }}
                    }} else {{
                        range.selectNodeContents(target);
                        range.collapse(false);
                    }}
                    
                    selection.removeAllRanges();
                    selection.addRange(range);
                    
                    // 高亮目标（短暂显示）
                    target.style.boxShadow = '0 0 0 3px #00FF00';
                    target.style.backgroundColor = 'rgba(0, 255, 0, 0.1)';
                    setTimeout(() => {{
                        target.style.boxShadow = '';
                        target.style.backgroundColor = '';
                    }}, 1000);
                    
                    return true;
                }}""")
                
                await asyncio.sleep(1)
                
                # 记录聚焦后状态
                await debug_editor_state("(JS聚焦后)")
                
                if focus_success == 'create':
                    # 标题存在但下方无可编辑区域，滚动到OKR表格下方动态创建
                    print(f"      [信息] '{section_name}' 标题存在但下方无输入区域，使用动态创建模式")
                    await _create_section_below_okr(page, section_name)
                elif not focus_success:
                    # JS聚焦完全失败，也尝试动态创建
                    print(f"      [信息] JavaScript聚焦失败，尝试动态创建模式")
                    await _create_section_below_okr(page, section_name)
                else:
                    dprint("      [✓] 成功激活")
                
                # 可选：输入测试字符验证（仅在DEBUG模式下）
                if DEBUG:
                    test_char = f'[{section_name[:2]}]'
                    await page.keyboard.type(test_char, delay=30)
                    await asyncio.sleep(0.2)
                    
                    has_test = await page.evaluate(f"""() => {{
                        const editor = document.querySelector('[data-slate-editor="true"]');
                        return editor ? editor.textContent.includes('{test_char}') : false;
                    }}""")
                    
                    if has_test:
                        dprint(f"      [✓] 验证成功，测试字符已输入")
                        # 删除测试字符
                        for _ in range(len(test_char)):
                            await page.keyboard.press('Backspace')
                            await asyncio.sleep(0.1)
                    else:
                        dprint(f"      [⚠️] 验证失败，但继续尝试")
                
                # 等待确保聚焦稳定
                await asyncio.sleep(0.2)  # 优化：0.3秒 → 0.2秒  # 优化：0.5秒 → 0.3秒
                
            else:
                # 页面没有该标题，需要动态创建
                dprint(f"\n  [创建] 页面缺少 '{section_name}' 标题，准备动态创建...")
                await _create_section_below_okr(page, section_name)
            
            # ==================== 插入空行 ====================
            dprint("\n    [插入空行] 准备输入位置...")
            
            # 输入2个空格
            for i in range(2):
                await page.keyboard.press('Space')
                await asyncio.sleep(0.1)
            
            # 按两次回车
            await page.keyboard.press('Enter')
            await asyncio.sleep(0.2)
            await page.keyboard.press('Enter')
            await asyncio.sleep(0.2)
            
            # 按上箭头回到第一行
            await page.keyboard.press('ArrowUp')
            await asyncio.sleep(0.2)

            # ==================== 输入内容 ====================
            dprint(f"\n    [输入内容] 开始输入 {section_name}...")
            # 一次性粘贴全部内容
            dprint(f"      内容长度: {len(content)} 字符")
            await page.keyboard.insert_text(content)
            await asyncio.sleep(0.5)  # 等待粘贴完成

            await page.keyboard.press('Enter')
            await asyncio.sleep(0.1)

            print(f"    [✅] {section_name} 填写完成")
            
            # 填写后检查状态（仅DEBUG）
            if DEBUG:
                await debug_editor_state(f"(填写{section_name}后)")
            
            success_count += 1
            
            # 短暂等待再继续
            await asyncio.sleep(1)
        
        print(f"\n  [填写统计] {success_count}/{total_sections} 个部分填写成功")

        return success_count > 0
        
    except Exception as e:
        print(f"\n  [X] 填写过程出错: {e}")
        import traceback
        traceback.print_exc()
        return False

# ==================== 遮挡层处理 ====================

async def remove_overlay_elements(page):
    """移除页面上的遮挡层元素"""
    try:
        await page.evaluate("""() => {
            // 移除常见的遮挡层类名
            const overlay_classes = [
                'mp_ai_plugin_shade',
                'MuiBackdrop-root',
                'MuiPopover-root',
                'ant-modal-mask',
                'ant-modal-wrap',
                'el-overlay',
                'ui-widget-overlay',
                'modal-backdrop',
                'overlay',
                'shade'
            ];

            overlay_classes.forEach(cls => {
                document.querySelectorAll('.' + cls).forEach(el => {
                    el.style.display = 'none';
                    el.style.pointerEvents = 'none';
                    if (el.parentNode) el.parentNode.removeChild(el);
                });
            });

            // 移除固定定位的全屏元素（可能是遮挡层）
            document.querySelectorAll('*').forEach(el => {
                const style = window.getComputedStyle(el);
                if (style.position === 'fixed' && 
                    style.zIndex > 1000 && 
                    el.offsetWidth > window.innerWidth * 0.5 &&
                    el.offsetHeight > window.innerHeight * 0.5) {
                    el.style.pointerEvents = 'none';
                }
            });

            return 'ok';
        }""")
        dprint("    [OK] 已清理遮挡层")
        return True
    except Exception as e:
        dprint(f"    [!] 清理遮挡层失败: {e}")
        return False

async def smart_click_with_fallback(page, locator, timeout=5000):
    """智能点击：尝试标准点击 -> 强制点击 -> JavaScript点击"""
    # 先尝试移除遮挡层
    await remove_overlay_elements(page)

    # 1. 尝试标准点击
    try:
        await locator.click(timeout=timeout)
        dprint("    [OK] 标准点击成功")
        return True
    except Exception as e:
        if "intercepts pointer events" in str(e):
            dprint("    [!] 被遮挡层拦截，尝试强制点击...")
        else:
            dprint(f"    [!] 标准点击失败: {e}")

    # 2. 尝试强制点击
    try:
        await locator.click(force=True, timeout=timeout)
        dprint("    [OK] 强制点击成功")
        return True
    except Exception as e:
        dprint(f"    [!] 强制点击失败: {e}")

    # 3. 使用JavaScript点击
    try:
        handle = await locator.element_handle()
        if handle:
            await handle.evaluate("el => { el.click(); el.dispatchEvent(new Event('click', {bubbles: true})); }")
            dprint("    [OK] JavaScript点击成功")
            return True
    except Exception as e:
        dprint(f"    [!] JavaScript点击失败: {e}")

    return False

async def smart_fill_input(page, label_text, content, create_if_missing=False, section_type=None):
    if not content or not content.strip():
        dprint(f"  [跳过] 内容为空: {label_text}")
        return False

    content = content.strip()
    dprint(f"\n  处理: {label_text[:40]}...")
    dprint(f"    内容长度: {len(content)}")
    dprint(f"    预览: {content[:80]}...")

    variants = [label_text]
    if "本周工作总结" in label_text:
        variants.extend(["工作总结", "本周工作"])
    elif "OKR" in label_text:
        variants.extend(["OKR进度", "OKR"])

    for variant in variants:
        try:
            locator = page.get_by_text(variant, exact=False).first
            if await locator.is_visible(timeout=2000):
                dprint(f"    [OK] 找到标签: {variant}")
                # 使用新的点击方式
                if await smart_click_with_fallback(page, locator):
                    await asyncio.sleep(0.2)  # 优化：0.3秒 → 0.2秒
                    if await find_and_fill_nearby_input(page, locator, content):
                        return True
        except Exception as e:
            dprint(f"    [!] 处理 {variant} 时出错: {e}")
            continue

    dprint("    尝试JavaScript直接填写...")
    if await fill_by_javascript(page, label_text, content):
        return True

    if create_if_missing and section_type:
        dprint("    尝试创建新区域...")
        if await create_new_section(page, section_type, label_text, content):
            return True

    dprint(f"    [X] 失败: {label_text}")
    return False

async def find_and_fill_nearby_input(page, label_locator, content):
    try:
        # 使用智能点击代替直接点击
        if not await smart_click_with_fallback(page, label_locator):
            dprint("    [!] 标签点击失败，尝试直接查找输入框...")

        await asyncio.sleep(0.2)  # 优化：0.3秒 → 0.2秒  # 优化：0.5秒 → 0.3秒
        handle = await label_locator.element_handle()
        if not handle:
            return False
        text = await handle.inner_text()
        selectors = ['textarea', 'input[type="text"]', '[contenteditable="true"]', '.ant-input']
        for sel in selectors:
            try:
                inp = page.locator(f'xpath=//*[contains(text(), "{text}")]/following::{sel}[1] | xpath=//*[contains(text(), "{text}")]/ancestor::div[1]//{sel}').first
                if await inp.is_visible(timeout=2000):
                    dprint(f"    [OK] 找到输入框: {sel}")
                    await inp.fill(content)
                    await asyncio.sleep(0.2)  # 优化：0.3秒 → 0.2秒  # 优化：0.5秒 → 0.3秒
                    return True
            except:
                continue
    except Exception as e:
        dprint(f"    [!] 失败: {e}")
    return False

async def fill_by_javascript(page, label_text, content):
    dprint(f"    [DEBUG-JS] 开始JavaScript填写，标签: {label_text[:30]}...")
    dprint(f"    [DEBUG-JS] 内容长度: {len(content)} 字符")

    try:
        # 在Python中预处理内容
        normalized = content.replace("\r\n", "\n").replace("\r", "\n")
        lines = normalized.split("\n")

        # 使用原始字符串避免转义问题
        js_code = """
        (args) => {
            const labelText = args.labelText;
            const lines = args.lines;

            console.log('JS: 开始查找标签:', labelText);

            // 改进的标签查找：支持部分匹配
            const findLabel = (text) => {
                // 首先尝试精确匹配
                let elements = document.querySelectorAll('*');
                for (let el of elements) {
                    if (el.textContent && el.textContent.trim() === text) {
                        return el;
                    }
                }
                // 然后尝试包含匹配
                for (let el of elements) {
                    if (el.textContent && el.textContent.trim().includes(text)) {
                        return el;
                    }
                }
                return null;
            };

            // 改进的输入框查找
            const findInput = (labelEl) => {
                if (!labelEl) return null;

                // 1. 检查标签本身是否是输入框
                if (labelEl.matches('textarea, input:not([type="hidden"]), [contenteditable="true"]')) {
                    return labelEl;
                }

                // 2. 查找子元素
                let inputs = labelEl.querySelectorAll('textarea, input:not([type="hidden"]), [contenteditable="true"]');
                if (inputs.length) return inputs[0];

                // 3. 查找兄弟元素（在同一父元素下）
                let parent = labelEl.parentElement;
                if (parent) {
                    inputs = parent.querySelectorAll('textarea, input:not([type="hidden"]), [contenteditable="true"]');
                    if (inputs.length) return inputs[0];
                }

                // 4. 向上查找5层祖先元素
                let ancestor = labelEl;
                for (let i=0; i<5; i++) {
                    ancestor = ancestor.parentElement;
                    if (!ancestor) break;
                    inputs = ancestor.querySelectorAll('textarea, input:not([type="hidden"]), [contenteditable="true"]');
                    if (inputs.length) return inputs[0];
                }

                // 5. 通过for属性查找（如果标签是label元素）
                if (labelEl.tagName === 'LABEL' && labelEl.htmlFor) {
                    const target = document.getElementById(labelEl.htmlFor);
                    if (target && (target.matches('textarea, input') || target.isContentEditable)) {
                        return target;
                    }
                }

                // 6. 在标签附近查找（下一个兄弟或子元素的下一个兄弟）
                let nextEl = labelEl.nextElementSibling;
                if (nextEl) {
                    if (nextEl.matches('textarea, input:not([type="hidden"]), [contenteditable="true"]')) {
                        return nextEl;
                    }
                    inputs = nextEl.querySelectorAll('textarea, input:not([type="hidden"]), [contenteditable="true"]');
                    if (inputs.length) return inputs[0];
                }

                // 7. 在整个文档中查找可见的输入框（作为兜底）
                const allInputs = document.querySelectorAll('textarea, input:not([type="hidden"])');
                for (let inp of allInputs) {
                    const style = window.getComputedStyle(inp);
                    if (style.display !== 'none' && style.visibility !== 'hidden') {
                        console.log('JS: 使用兜底方案，找到可见输入框');
                        return inp;
                    }
                }

                return null;
            };

            const label = findLabel(labelText);
            if (!label) {
                console.log('JS: 未找到标签');
                return 'no_label';
            }
            console.log('JS: 找到标签:', label.tagName, label.textContent.substring(0, 50));

            const input = findInput(label);
            if (!input) {
                console.log('JS: 未找到输入框');
                return 'no_input';
            }
            console.log('JS: 找到输入框:', input.tagName, input.className);

            // 确保输入框可见且可交互
            input.scrollIntoView({behavior:'smooth', block:'center'});
            input.focus();
            input.click();

            // 将行数组合并
            const fullContent = lines.join('\\n');

            if (input.tagName === 'TEXTAREA' || input.tagName === 'INPUT') {
                input.value = fullContent;
                // 触发input事件
                input.dispatchEvent(new Event('input', {bubbles: true}));
                input.dispatchEvent(new Event('change', {bubbles: true}));
                console.log('JS: 已设置value, 长度:', fullContent.length);
            } else if (input.isContentEditable || input.getAttribute('contenteditable') === 'true') {
                // 对于contenteditable，使用<br>标签
                const htmlContent = lines.join('<br>');
                input.innerHTML = htmlContent;
                // 触发input事件
                input.dispatchEvent(new Event('input', {bubbles: true}));
                console.log('JS: 已设置innerHTML');
            }

            // 触发其他事件
            input.dispatchEvent(new Event('blur', {bubbles: true}));

            console.log('JS: 已触发所有事件');
            return 'ok';
        }
        """

        result = await page.evaluate(js_code, {"labelText": label_text, "lines": lines})

        dprint(f"    [DEBUG-JS] 执行结果: {result}")
        if result == 'ok':
            dprint(f"    [OK] JavaScript填写成功")
            await asyncio.sleep(0.2)  # 优化：0.3秒 → 0.2秒  # 优化：0.5秒 → 0.3秒
            return True
        else:
            dprint(f"    [!] JavaScript返回: {result}")
    except Exception as e:
        dprint(f"    [!] JavaScript错误: {e}")
        import traceback
        traceback.print_exc()
    return False


async def create_new_section(page, section_type, title, content):
    dprint(f"    [DEBUG] 创建新区域: {title[:30]}...")
    try:
        result = await page.evaluate("""(args) => {
            const title = args.title;
            const content = args.content;

            const container = document.querySelector('.ant-form, .weekly-form, main') || document.body;
            const div = document.createElement('div');
            div.style.cssText = 'margin-top:20px;padding:16px;border:1px solid #d9d9d9;border-radius:4px;';

            const h3 = document.createElement('h3');
            h3.textContent = title;
            h3.style.marginBottom = '12px';

            const textarea = document.createElement('textarea');
            textarea.value = content;
            textarea.style.cssText = 'width:100%;min-height:120px;padding:8px;';

            div.appendChild(h3);
            div.appendChild(textarea);
            container.appendChild(div);

            textarea.focus();
            textarea.dispatchEvent(new Event('input', {bubbles:true}));

            return 'created';
        }""", {"title": title, "content": content})

        if result == 'created':
            dprint(f"    [OK] 创建成功")
            return True
    except Exception as e:
        dprint(f"    [X] 创建失败: {e}")
    return False


# ==================== 提交和确认 ====================

# ==================== 修复自动提交确认函数 ====================

async def auto_submit_and_confirm(page):
    """自动提交并确认 - 优化版（基于成功方法）"""
    
    print("\n[提交] 自动提交...")

    # 先滚动到页面底部，确保提交按钮可见
    dprint("  滚动到页面底部...")
    await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
    await asyncio.sleep(1.0)

    submit_clicked = False
    submit_button_text = ""
    
    # 方法1: 直接使用JavaScript点击提交/更新按钮（优先查找"提交"，再查找"更新"）
    dprint("  尝试JavaScript点击提交/更新按钮...")
    result = await page.evaluate("""() => {
        // 查找提交/更新按钮
        const buttons = Array.from(document.querySelectorAll('button, [role="button"]'));
        
        // 第一步：查找包含"提交"文本的按钮
        for (let btn of buttons) {
            const text = (btn.textContent || btn.value || btn.getAttribute('aria-label') || '').trim();
            if (text.includes('提交') && 
                btn.offsetParent !== null && 
                window.getComputedStyle(btn).display !== 'none') {
                btn.click();
                console.log('JavaScript点击提交:', text);
                return text;
            }
        }
        
        // 第二步：查找包含"更新"文本的按钮
        for (let btn of buttons) {
            const text = (btn.textContent || btn.value || btn.getAttribute('aria-label') || '').trim();
            if (text.includes('更新') && 
                btn.offsetParent !== null && 
                window.getComputedStyle(btn).display !== 'none') {
                btn.click();
                console.log('JavaScript点击更新:', text);
                return text;
            }
        }
        return false;
    }""")
    
    if result:
        submit_clicked = True
        submit_button_text = result
        dprint(f"  JavaScript点击: {result}")
    else:
        # 备用：尝试普通点击（调整为"提交"优先）
        for text in ["提交", "更新", "保存", "Submit", "Save"]:
            try:
                button = page.get_by_role("button", name=text).first
                if await button.count() > 0 and await button.is_visible(timeout=2000):
                    dprint(f"  尝试普通点击: {text}")
                    await button.click(force=True)  # 使用force=True忽略遮挡
                    submit_clicked = True
                    break
            except Exception as e:
                dprint(f"  尝试点击 '{text}' 失败: {e}")
                continue

    # 备用方案3：使用固定坐标点击右上角（页面右上角固定位置）
    if not submit_clicked:
        dprint("  尝试固定坐标点击右上角...")
        try:
            # 获取视口大小
            viewport = await page.evaluate("""() => {
                return { width: window.innerWidth, height: window.innerHeight };
            }""")
            # 右上角位置：X = 宽度 - 150, Y = 50
            x = viewport['width'] - 150
            y = 50
            dprint(f"  点击坐标: ({x}, {y})")
            await page.mouse.click(x, y)
            await asyncio.sleep(0.5)
            submit_clicked = True
            submit_button_text = "固定坐标点击"
            dprint("  固定坐标点击成功")
        except Exception as e:
            dprint(f"  固定坐标点击失败: {e}")

    if not submit_clicked:
        print("  [X] 未找到提交按钮")
        return False

    print("\n[确认] 处理确认弹框...")
    
    # 等待弹框出现
    dprint("  等待2秒让弹框出现...")
    await asyncio.sleep(2)
    
    # 先尝试滚动到底部，确保弹框可见
    await page.evaluate("""() => {
        window.scrollTo(0, document.body.scrollHeight);
    }""")
    await asyncio.sleep(0.5)
    
    confirm_clicked = False
    
    # 方法1: 使用成功的选择器 .erp-modal-footer button:has-text("确定")
    dprint("\n  方法1: 使用成功选择器 .erp-modal-footer button:has-text('确定')...")
    try:
        confirm_btn = page.locator('.erp-modal-footer button:has-text("确定")').first
        if await confirm_btn.count() > 0:
            dprint("  找到按钮")
            try:
                await confirm_btn.click(force=True, timeout=3000)
                dprint("  强制点击成功")
                confirm_clicked = True
            except:
                try:
                    await confirm_btn.evaluate("btn => btn.click()")
                    dprint("  JavaScript点击成功")
                    confirm_clicked = True
                except:
                    # 获取坐标点击
                    box = await confirm_btn.bounding_box()
                    if box:
                        x = box['x'] + box['width'] / 2
                        y = box['y'] + box['height'] / 2
                        await page.mouse.click(x, y)
                        dprint(f"  坐标点击成功: ({x:.0f}, {y:.0f})")
                        confirm_clicked = True
    except Exception as e:
        dprint(f"  方法1失败: {e}")

    # 方法2: 使用通用选择器 .mp-btn.mp-modal-enter-btn（备用）
    if not confirm_clicked:
        dprint("\n  方法2: 使用选择器 .mp-btn.mp-modal-enter-btn...")
        try:
            confirm_btn = page.locator('.mp-btn.mp-modal-enter-btn').first
            if await confirm_btn.count() > 0:
                dprint("  找到按钮")

                # 使用 evaluate 直接点击
                await confirm_btn.evaluate("btn => btn.click()")
                dprint("  JavaScript点击成功")
                confirm_clicked = True
        except Exception as e:
            dprint(f"  方法2失败: {e}")

    # 方法3: XPath查找
    if not confirm_clicked:
        dprint("\n  方法3: XPath查找...")
        try:
            confirm_btn = page.locator('xpath=//button[contains(text(), "确定")]').first
            if await confirm_btn.count() > 0:
                dprint("  找到包含'确定'的按钮")
                await confirm_btn.evaluate("btn => btn.click()")
                dprint("  JavaScript点击成功")
                confirm_clicked = True
        except Exception as e:
            dprint(f"  方法3失败: {e}")

    # 方法4: JavaScript兜底查找
    if not confirm_clicked:
        dprint("\n  方法4: JavaScript兜底查找...")
        result = await page.evaluate("""() => {
            // 查找所有可能包含确认按钮的容器
            const dialogs = document.querySelectorAll('.erp-dialog-container, .mp-modal, [class*="dialog"], [class*="modal"]');
            
            for (let dialog of dialogs) {
                if (dialog.style.display === 'none' || window.getComputedStyle(dialog).display === 'none') continue;
                
                const buttons = dialog.querySelectorAll('button');
                for (let btn of buttons) {
                    const text = (btn.textContent || '').trim();
                    if (text.includes('确定') || text.includes('确认')) {
                        btn.click();
                        console.log('JavaScript点击确认:', text);
                        return { method: 'dialog', text: text };
                    }
                }
            }
            
            // 如果没找到对话框，查找所有按钮
            const allButtons = document.querySelectorAll('button');
            for (let btn of allButtons) {
                const text = (btn.textContent || '').trim();
                if (text.includes('确定') || text.includes('确认')) {
                    btn.click();
                    console.log('JavaScript点击确认(全局):', text);
                    return { method: 'global', text: text };
                }
            }
            return false;
        }""")
        
        if result:
            dprint(f"  JavaScript点击成功: {result}")
            confirm_clicked = True

    # 等待确认完成
    if confirm_clicked:
        print("  [OK] 确认完成")
        
        # 等待2秒让弹框消失（不检查对话框状态，因为知道最终会消失）
        await asyncio.sleep(2)
        
        # 可选：简单检查一下，但不影响流程
        try:
            dialog_exists = await page.evaluate("""() => {
                const dialogs = document.querySelectorAll('.erp-dialog-container, .mp-modal');
                for (let d of dialogs) {
                    if (d.offsetParent !== null) return true;
                }
                return false;
            }""")
            
            if dialog_exists:
                dprint("  [信息] 对话框可能还在，但已知最终会消失")
            else:
                dprint("  [OK] 对话框已关闭")
        except:
            pass
            
        return True
    else:
        print("  [X] 未找到确认按钮")
        return False

async def diagnose_save_issue(page):
    """诊断保存问题的原因"""
    
    dprint("\n" + "="*60)
    dprint("[诊断] 开始分析保存问题")
    dprint("="*60)
    
    # 1. 获取当前编辑器内容
    current_content = await page.evaluate("""() => {
        const editor = document.querySelector('[data-slate-editor="true"]');
        if (!editor) return [];
        
        const paragraphs = editor.querySelectorAll('.mp-paragraph-wrapper');
        const content = [];
        paragraphs.forEach((p, index) => {
            const text = p.textContent || '';
            content.push({
                index: index,
                text: text.substring(0, 100),
                length: text.length,
                charCodes: Array.from(text).map(c => c.charCodeAt(0)).slice(0, 20)
            });
        });
        return content;
    }""")
    
    dprint("\n  [分析] 当前编辑器内容:")
    has_invisible = False
    for p in current_content:
        dprint(f"    段落 {p['index']}: {p['text']}")
        dprint(f"        长度: {p['length']}, 字符编码: {p['charCodes']}")
        
        # 检查不可见字符
        for code in p['charCodes']:
            if code < 32 or code == 8203:  # 控制字符或零宽空格
                has_invisible = True
                dprint(f"        ⚠️ 发现不可见字符: {code}")
    
    if has_invisible:
        dprint("\n  [⚠️ 发现] 存在不可见字符，可能需要清理")
    else:
        dprint("\n  [✅ 正常] 未发现不可见字符")
    
    # 2. 检查编辑器类型
    editor_info = await page.evaluate("""() => {
        const editor = document.querySelector('[data-slate-editor="true"]');
        if (!editor) return null;
        
        const reactProps = Object.keys(editor).filter(key => key.startsWith('__reactProps$'));
        const slateProps = Object.keys(editor).filter(key => key.includes('slate'));
        
        return {
            hasReactProps: reactProps.length > 0,
            reactPropsCount: reactProps.length,
            hasSlateProps: slateProps.length > 0,
            slateProps: slateProps.slice(0, 5),
            className: editor.className,
            id: editor.id,
            tagName: editor.tagName
        };
    }""")
    
    dprint(f"\n  [分析] 编辑器信息:")
    if editor_info:
        dprint(f"    类型: {editor_info['tagName']}.{editor_info['className']}")
        dprint(f"    有React属性: {editor_info['hasReactProps']} (找到 {editor_info['reactPropsCount']} 个)")
        dprint(f"    有Slate属性: {editor_info['hasSlateProps']}")
    
    # 3. 测试写入并触发事件
    dprint("\n  [测试] 尝试写入测试内容...")
    
    test_result = await page.evaluate("""() => {
        return new Promise((resolve) => {
            const editor = document.querySelector('[data-slate-editor="true"]');
            if (!editor) {
                resolve({success: false, error: '未找到编辑器'});
                return;
            }
            
            // 找到本周工作总结标题
            const headings = editor.querySelectorAll('h3.mp-heading, h1.mp-heading');
            let titleElement = null;
            for (const h of headings) {
                if (h.textContent && h.textContent.includes('本周工作总结')) {
                    titleElement = h;
                    break;
                }
            }
            
            if (!titleElement) {
                resolve({success: false, error: '未找到本周工作总结标题'});
                return;
            }
            
            // 创建测试段落
            const testPara = document.createElement('div');
            testPara.className = 'mp-paragraph-wrapper';
            testPara.setAttribute('data-morpho-type', 'paragraph');
            
            const timestamp = Date.now();
            const testContent = `[诊断测试] ${timestamp} - 这是一段测试内容，请手动保存看看能否保留`;
            
            const spanWrapper = document.createElement('span');
            spanWrapper.setAttribute('data-slate-node', 'text');
            
            const spanLeaf = document.createElement('span');
            spanLeaf.className = '';
            spanLeaf.setAttribute('data-slate-leaf', 'true');
            
            const contentSpan = document.createElement('span');
            contentSpan.setAttribute('data-slate-string', 'true');
            contentSpan.textContent = testContent;
            
            spanLeaf.appendChild(contentSpan);
            spanWrapper.appendChild(spanLeaf);
            testPara.appendChild(spanWrapper);
            
            // 高亮显示
            testPara.style.backgroundColor = 'yellow';
            testPara.style.border = '3px solid red';
            
            // 在标题后插入
            const nextElement = titleElement.nextElementSibling;
            if (nextElement) {
                titleElement.parentNode.insertBefore(testPara, nextElement);
            } else {
                titleElement.parentNode.appendChild(testPara);
            }
            
            // 触发事件
            const events = [
                new Event('input', {bubbles: true}),
                new Event('change', {bubbles: true}),
                new Event('blur', {bubbles: true})
            ];
            
            events.forEach(event => {
                testPara.dispatchEvent(event);
            });
            
            // 滚动到可见区域
            testPara.scrollIntoView({behavior: 'smooth', block: 'center'});
            
            resolve({
                success: true,
                testContent: testContent,
                timestamp: timestamp
            });
        });
    }""")
    
    if test_result.get('success'):
        dprint(f"\n  [测试结果]")
        dprint(f"    写入内容: {test_result['testContent']}")
        dprint(f"    时间戳: {test_result['timestamp']}")
        dprint("\n  ⚠️ 黄色高亮段落是测试内容，请现在手动点击保存按钮")
    else:
        dprint(f"\n  [测试失败] {test_result.get('error')}")
        return
    
    try:
        if sys.stdin and sys.stdin.isatty():
            input("\n  手动保存后按Enter继续验证...")
    except:
        pass  # 无控制台时跳过
    
    # 4. 验证测试内容是否被保存
    verify_result = await page.evaluate("""() => {
        const editor = document.querySelector('[data-slate-editor="true"]');
        if (!editor) return {found: false, error: '未找到编辑器'};
        
        const paragraphs = editor.querySelectorAll('.mp-paragraph-wrapper');
        let found = false;
        let content = '';
        
        for (const p of paragraphs) {
            if (p.textContent && p.textContent.includes('[诊断测试]')) {
                found = true;
                content = p.textContent;
                break;
            }
        }
        
        return {
            found: found,
            content: content || '未找到'
        };
    }""")
    
    dprint(f"\n  [验证结果]")
    dprint(f"    测试内容是否还在: {verify_result['found']}")
    if verify_result['found']:
        dprint(f"    保存后的内容: {verify_result['content']}")
        dprint("\n  ✅ 测试内容被成功保存！说明事件触发正确")
    else:
        dprint("\n  ❌ 测试内容丢失！说明事件触发可能有问题")
        dprint("     可能原因：")
        dprint("     - 需要触发更多事件")
        dprint("     - 编辑器有内部状态需要更新")
        dprint("     - 需要聚焦编辑器")
    
    return {
        "has_invisible_chars": has_invisible,
        "editor_info": editor_info,
        "test_saved": verify_result['found']
    }


# ==================== 主流程 ====================

async def submit_single_weekly_v24(page, target_date=None, dry_run=True, auto_submit=False, is_first_page=False):
    target_date = target_date or get_week_date(TARGET_WEEK)
    edit_url = f"{BASE_URL}#/home?date=s{target_date}&mandatorId={MY_ID}&isEdit=true"

    dprint("="*60)
    dprint(f"目标日期: {target_date}")
    dprint(f"编辑模式: {'预览' if dry_run else '实际填写'}")
    dprint(f"自动提交: {'是' if auto_submit else '否'}")
    dprint("="*60)

    try:
        try:
            await page.goto(edit_url, wait_until='domcontentloaded', timeout=60000)
        except Exception as nav_error:
            if _is_browser_disconnected(nav_error):
                raise
            if _is_network_error(nav_error):
                return {'status': 'error', 'error': f'网络不可达: {nav_error}'}
            if 'uuap' in page.url:
                return {'status': 'login_required', 'url': page.url}

        if is_first_page:
            dprint("\n(首次页面等待4秒)...", end=" ")
            await asyncio.sleep(4)  # 增加等待时间确保页面完全加载
            dprint("OK")
        else:
            await asyncio.sleep(2)

        if 'uuap' in page.url:
            return {'status': 'login_required', 'url': page.url}

        dprint("[OK] 页面加载完成")

    except Exception as e:
        print(f"\n[X] 页面加载失败: {e}")
        return {'status': 'error', 'error': str(e)}

    await analyze_page_structure(page)

    dprint("\n滚动加载页面内容...")
    selector, center = await find_scroll_container(page)
    await scroll_container(page, selector, center)

    dprint("\n" + "="*60)
    dprint("[验证] 数据状态检查")
    dprint("="*60)
    dprint(f"  part1_work_summary: {WEEKLY_DATA.get('part1_work_summary') is not None}")
    dprint(f"  part2_okr_structure: {WEEKLY_DATA.get('part2_okr_structure') is not None}")
    dprint(f"  part3_metrics: {WEEKLY_DATA.get('part3_metrics') is not None}")

    if dry_run:
        print("\n" + "="*60)
        print("【预览模式】")
        print("="*60)

        if WEEKLY_DATA.get("part1_work_summary"):
            print(f"\n1. 本周工作总结:")
            print(f"{WEEKLY_DATA['part1_work_summary']['content'][:300]}...")

        if WEEKLY_DATA.get("part2_okr_structure"):
            print(f"\n2. OKR进度:")
            for obj in WEEKLY_DATA["part2_okr_structure"]["objectives"]:
                print(f"  {obj['o_id']}: {obj['o_title']}")
                for kr in obj["key_results"]:
                    print(f"    {kr['kr_id']}: {kr['weekly_work'][:50]}...")

        if WEEKLY_DATA.get("part3_metrics"):
            print(f"\n3. 关键指标:")
            for section_name, section_data in WEEKLY_DATA["part3_metrics"]["sections"].items():
                print(f"  {section_name}:")
                print(f"    {section_data['formatted_content'][:100]}...")

        return {'status': 'preview', 'date': target_date}

    print("\n开始填写内容...")

    results = {}

    part1_data = WEEKLY_DATA.get("part1_work_summary")
    part2_data = WEEKLY_DATA.get("part2_okr_structure")
    part3_data = WEEKLY_DATA.get("part3_metrics")

    dprint(f"\n[填写前数据检查]")
    dprint(f"  part1_data: {part1_data is not None}")
    dprint(f"  part2_data: {part2_data is not None}")
    dprint(f"  part3_data: {part3_data is not None}")

    if part1_data:
        results['part1'] = await fill_part1_work_summary(page, part1_data)
    else:
        print("\n[!] 无第一部分数据，跳过")
        results['part1'] = False

    # 等待页面稳定后再进行下一部分
    await asyncio.sleep(1.0)
    dprint("\n[INFO] 第一部分完成，页面稳定等待...")

    if part2_data:
        results['part2'] = await fill_okr_section(page, part2_data, WEEKLY_DATA.get("base_employee_name"))
    else:
        print("\n[!] 无第二部分数据，跳过")
        results['part2'] = False

    # 等待页面稳定后再进行下一部分
    await asyncio.sleep(1.0)
    dprint("\n[INFO] 第二部分完成，页面稳定等待...")

    if part3_data:
        results['part3'] = await fill_part3_metrics(page, part3_data)
    else:
        print("\n[!] 无第三部分数据，跳过")
        results['part3'] = False

    print("\n" + "="*60)
    print("填写完成统计：")
    print(f"  本周工作总结: {'OK' if results.get('part1') else 'FAIL'}")
    print(f"  OKR本周工作: {'OK' if results.get('part2') else 'FAIL'}")
    print(f"  关键指标: {'OK' if results.get('part3') else 'FAIL'}")
    print("="*60)

    if auto_submit and any(results.values()):
        print("\n[自动提交] 准备自动提交...")
        submit_result = await auto_submit_and_confirm(page)
        results['submitted'] = submit_result

        if submit_result:
            print("[OK] 自动提交完成")
        else:
            print("[!] 自动提交可能失败")

    return {'status': 'success', 'date': target_date, 'results': results}

async def submit_weekly_with_retry_v24(page, **kwargs):
    result = await submit_single_weekly_v24(page, is_first_page=True, **kwargs)

    if result['status'] == 'error':
        print(f"\n{'='*60}")
        print("[重试机制] 首次失败，准备重试...")
        print(f"{'='*60}")
        await asyncio.sleep(2)

        result = await submit_single_weekly_v24(page, is_first_page=False, **kwargs)

        if result['status'] == 'success':
            print("\n[OK] 重试成功")
        else:
            print("\n[X] 重试失败")

    return result

async def create_section_with_content(page, title, content):
    """创建新区域并填写内容（兜底方案）"""
    dprint(f"    [创建区域] 创建 '{title}' 区域...")

    try:
        # 使用 JavaScript 创建新区域
        result = await page.evaluate("""({title, content}) => {
            // 查找或创建容器
            let container = document.querySelector('.weekly-form, .ant-form, main, body');
            if (!container) container = document.body;

            // 创建区域
            const section = document.createElement('div');
            section.style.cssText = 'margin: 20px 0; padding: 16px; border: 1px solid #d9d9d9; border-radius: 4px; background: #fff;';

            // 创建标题
            const h3 = document.createElement('h3');
            h3.textContent = title;
            h3.style.cssText = 'margin: 0 0 12px 0; font-weight: bold; color: #333;';

            // 创建文本框
            const textarea = document.createElement('textarea');
            textarea.value = content;
            textarea.style.cssText = 'width: 100%; min-height: 150px; padding: 8px; border: 1px solid #d9d9d9; border-radius: 4px; font-family: inherit; font-size: 14px; line-height: 1.5; resize: vertical;';
            textarea.placeholder = '请输入' + title;

            // 组装
            section.appendChild(h3);
            section.appendChild(textarea);

            // 添加到页面（尝试添加到合适的位置）
            // 优先添加到表单末尾
            const form = document.querySelector('form, .ant-form');
            if (form) {
                form.appendChild(section);
            } else {
                container.appendChild(section);
            }

            // 触发事件
            textarea.focus();
            textarea.dispatchEvent(new Event('input', {bubbles: true}));
            textarea.dispatchEvent(new Event('change', {bubbles: true}));

            return 'created';
        }""", {"title": title, "content": content})

        if result == 'created':
            dprint(f"    [OK] 已创建 '{title}' 区域并填写内容")
            return True
        else:
            dprint(f"    [!] 创建区域返回: {result}")
            return False

    except Exception as e:
        dprint(f"    [X] 创建区域失败: {e}")
        return False

async def find_editor_api(page):
    """查找编辑器可用的API"""
    
    dprint("\n[API诊断] 查找编辑器内部API...")
    
    api_info = await page.evaluate("""() => {
        const editor = document.querySelector('[data-slate-editor="true"]');
        if (!editor) return {error: '未找到编辑器', hasSlate: false};
        
        const result = {
            hasSlate: false,
            slateEditor: null,
            props: [],
            methods: []
        };
        
        // 1. 查找所有可能的React属性
        const allProps = Object.keys(editor).filter(key => 
            key.startsWith('__reactProps$') || 
            key.includes('slate') || 
            key.includes('Slate') ||
            key.includes('editor')
        );
        
        result.props = allProps;
        
        // 2. 尝试通过React属性获取编辑器实例
        for (const prop of allProps) {
            try {
                const value = editor[prop];
                if (value && typeof value === 'object') {
                    // 可能是React props
                    if (value.editor && typeof value.editor === 'object') {
                        result.slateEditor = value.editor;
                        result.hasSlate = true;
                        break;
                    }
                    // 或者直接是编辑器实例
                    if (value.insertText && value.delete) {
                        result.slateEditor = value;
                        result.hasSlate = true;
                        break;
                    }
                    // 或者有children属性
                    if (value.children && Array.isArray(value.children)) {
                        result.slateEditor = value;
                        result.hasSlate = true;
                        break;
                    }
                }
            } catch (e) {
                // 忽略访问错误
            }
        }
        
        // 3. 如果找到编辑器实例，检查它的方法
        if (result.slateEditor) {
            const methods = [];
            for (const key in result.slateEditor) {
                try {
                    if (typeof result.slateEditor[key] === 'function') {
                        methods.push(key);
                    }
                } catch (e) {
                    // 忽略
                }
            }
            result.methods = methods.slice(0, 30); // 取前30个
            
            // 4. 尝试获取编辑器的选择状态
            try {
                if (result.slateEditor.selection) {
                    result.hasSelection = true;
                }
            } catch (e) {}
        }
        
        return result;
    }""")
    
    dprint(f"\n  [诊断详情]")
    if api_info.get('error'):
        dprint(f"    错误: {api_info['error']}")
    else:
        dprint(f"    找到React属性: {len(api_info.get('props', []))} 个")
        for i, prop in enumerate(api_info.get('props', [])[:5]):
            dprint(f"      - {prop}")

        dprint(f"\n    找到Slate编辑器: {api_info.get('hasSlate', False)}")
        if api_info.get('hasSlate'):
            dprint(f"    编辑器方法数量: {len(api_info.get('methods', []))}")
            if api_info.get('methods'):
                dprint(f"    前10个方法: {api_info['methods'][:10]}")
    
    return api_info

# ==================== 浏览器断连检测 ====================
def _is_browser_disconnected(error):
    """检测异常是否为浏览器断连"""
    err_str = str(error).lower()
    disconnect_keywords = [
        "browser has been closed", "browser closed", "target closed",
        "connection refused", "disconnected", "not connected",
        "cdp session closed", "page closed", "context closed",
    ]
    return any(kw in err_str for kw in disconnect_keywords)

def _is_network_error(error):
    """检测异常是否为网络不可达（DNS解析失败等致命网络错误）"""
    err_str = str(error).lower()
    network_keywords = [
        "net::err_name_not_resolved",
        "net::err_connection_refused",
        "net::err_connection_timed_out",
        "net::err_internet_disconnected",
        "net::err_network_changed",
        "net::err_address_unreachable",
        "net::err_connection_reset",
        "net::err_network_io_suspended",
    ]
    return any(kw in err_str for kw in network_keywords)

async def _is_page_unreachable(page):
    """检测页面是否显示网络错误（无法访问此网站等）"""
    try:
        title = await page.title()
        title_lower = title.lower() if title else ""
        body_text = await page.evaluate("document.body ? document.body.innerText.substring(0, 500) : ''")
        body_lower = body_text.lower() if body_text else ""
        error_indicators = [
            "无法访问此网站", "无法访问此页面",
            "err_name_not_resolved", "err_connection_refused",
            "err_connection_timed_out", "err_internet_disconnected",
            "err_address_unreachable",
            "this site can't be reached", "this page isn't working",
            "dns_probe_finished", "took too long to respond",
        ]
        return any(indicator in body_lower or indicator in title_lower for indicator in error_indicators)
    except Exception:
        return False

# ==================== 新增：模块化导出函数 ====================

async def submit_okr_report(optimized_data):
    """
    【新增】核心导出函数 - 提交优化后的周报
    :param optimized_data: 优化后的周报数据（字典格式，与团队周报JSON结构一致）
    :return: 提交结果（True/False）
    """
    iprint(f"\n{'='*60}")
    iprint(f"【模块调用】开始提交优化后的周报")
    iprint(f"{'='*60}")
    
    # 保存数据到临时文件（因为原脚本依赖文件读取）
    import tempfile
    import json
    
    # 创建临时文件
    with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', suffix='.json', delete=False) as f:
        json.dump(optimized_data, f, ensure_ascii=False, indent=2)
        temp_file = f.name
    
    iprint(f"[临时文件] 已创建: {temp_file}")
    
    # 备份原始命令行参数
    original_argv = sys.argv.copy()
    
    try:
        # 修改全局变量，使脚本读取我们的临时文件
        global TARGET_WEEK
        global DRY_RUN
        global AUTO_SUBMIT
        
        # 保存原始值
        original_week = TARGET_WEEK
        original_dry_run = DRY_RUN
        original_auto_submit = AUTO_SUBMIT
        
        # 设置为实际提交模式
        DRY_RUN = False
        AUTO_SUBMIT = True
        
        # 临时将团队周报文件复制到目标位置（因为原脚本会按文件名查找）
        week_suffix = get_week_filename_suffix(TARGET_WEEK)
        target_file = os.path.join(WEEKLYREPORT_DIR, f'团队周报_{week_suffix}.json')
        
        # 确保目录存在
        os.makedirs(WEEKLYREPORT_DIR, exist_ok=True)
        
        # 复制临时文件到目标位置
        import shutil
        shutil.copy2(temp_file, target_file)
        iprint(f"[文件] 已复制到目标位置: {target_file}")
        
        # 执行提交
        iprint("[调用] 执行原脚本主函数...")
        
        # 调用原脚本的main函数
        await main()
        
        iprint("[完成] 提交脚本执行完成")
        
        return True
        
    except Exception as e:
        iprint(f"[错误] 提交过程异常: {e}")
        import traceback
        traceback.print_exc()
        return False
        
    finally:
        # 恢复原始值
        sys.argv = original_argv
        TARGET_WEEK = original_week
        DRY_RUN = original_dry_run
        AUTO_SUBMIT = original_auto_submit
        
        # 清理临时文件
        try:
            if os.path.exists(temp_file):
                os.unlink(temp_file)
                iprint(f"[清理] 已删除临时文件: {temp_file}")
        except:
            pass

def save_submit_result_to_file(result, file_path=None):
    """
    【新增】保存提交结果到文件
    :param result: 提交结果
    :param file_path: 指定文件路径（可选）
    :return: 保存的文件路径
    """
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    
    if file_path is None:
        file_path = os.path.join(WEEKLYREPORT_DIR, f"submit_result_{timestamp}.json")
    
    # 确保目录存在
    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    
    result_data = {
        "timestamp": timestamp,
        "result": result,
        "status": "success" if result else "failed"
    }
    
    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(result_data, f, ensure_ascii=False, indent=2)
    
    iprint(f"[保存] 提交结果已保存到: {file_path}")
    return file_path

async def main():
    print("\n" + "="*60)
    print("程序启动--OKR周报自动提交")
    print("="*60)

    if not check_network_connection():
        print("\n[!] 网络检查失败")
        try:
            response = input("\n是否仍要继续? (y/N): ")
        except (EOFError, OSError, RuntimeError):
            response = "n"  # 无控制台时默认不继续
        if response.lower() != 'y':
            return

    global TARGET_WEEK
    if len(sys.argv) > 1:
        try:
            TARGET_WEEK = int(sys.argv[1])
        except ValueError:
            print(f"❌ 错误：参数「{sys.argv[1]}」不是有效数字！")
            return
    print(f"📄 获取目标周：{TARGET_WEEK}")

    target_date = get_week_date(TARGET_WEEK)

    print("="*60)
    print(f"\n[配置]")
    print(f"  用户ID: {MY_ID}")
    print(f"  目标日期: {target_date}")
    print(f"  编辑模式: {'预览' if DRY_RUN else '实际填写'}")
    print(f"  自动提交: {'是' if AUTO_SUBMIT else '否'}")

    json_loaded = load_content_from_json()

    if json_loaded:
        print(f"  ✓ [数据加载成功]")
        dprint(f"\n  [数据状态确认]")
        dprint(f"    part1_work_summary: {WEEKLY_DATA.get('part1_work_summary') is not None}")
        dprint(f"    part2_okr_structure: {WEEKLY_DATA.get('part2_okr_structure') is not None}")
        dprint(f"    part3_metrics: {WEEKLY_DATA.get('part3_metrics') is not None}")

        if WEEKLY_DATA.get("part1_work_summary"):
            dprint(f"    - 工作总结: {len(WEEKLY_DATA['part1_work_summary']['content'])} 字符")
        if WEEKLY_DATA.get("part2_okr_structure"):
            kr_count = sum(len(o.get("key_results", [])) for o in WEEKLY_DATA["part2_okr_structure"].get("objectives", []))
            dprint(f"    - OKR工作: {kr_count} 条KR")
        if WEEKLY_DATA.get("part3_metrics"):
            dprint(f"    - 关键指标: {len(WEEKLY_DATA['part3_metrics'].get('sections', {}))} 项")
    else:
        print(f"  ✗ 未能加载数据")
        return

    print(f"\n{'='*60}")
    print("[系统] 初始化浏览器管理器...")
    dprint(f"[系统] 外部模块加载状态: {_BROWSER_AUTOMATION_LOADED}")

    # 使用浏览器管理器
    browser_manager = BrowserManager()

    # 尝试连接到已存在的浏览器
    browser_connected = False
    dprint(f"[浏览器] 状态文件: {BROWSER_STATE_FILE}")
    dprint(f"[浏览器] 状态文件存在: {os.path.exists(str(BROWSER_STATE_FILE))}")

    if browser_manager.load_state():
        dprint(f"[浏览器] 加载状态成功: 端口={browser_manager.cdp_port}, PID={browser_manager.browser_pid}")
        browser_connected = await browser_manager.connect_to_existing_browser()
        if not browser_connected:
            dprint("[浏览器] 连接已存在的浏览器失败")
    else:
        dprint("[浏览器] 加载状态失败或浏览器进程不存在")

    if not browser_connected:
        # 启动新浏览器
        print("\n[浏览器] 启动新浏览器...")
        if not await browser_manager.launch_new_browser(headless=False):
            print("[X] 浏览器启动失败")
            return

    iprint("[浏览器] 等待浏览器加载完成...")
    await asyncio.sleep(5)
# 2. 确保页面实例存在
    await browser_manager.ensure_page()
    page = browser_manager.page

    # 3. 获取当前URL并判断是否需要初始导航
    current_url = page.url
    if BASE_URL not in current_url:
        iprint(f"\n[导航] 1当前URL: {current_url}，导航到目标页面...")
        # 增加异常捕获，允许重试一次
        try:
            await page.goto(BASE_URL, wait_until='domcontentloaded', timeout=30000)
        except asyncio.TimeoutError:
            iprint("[!] 页面加载超时，重试一次...")
            await page.goto(BASE_URL, wait_until='domcontentloaded', timeout=30000)
        except Exception as e:
            if _is_network_error(e):
                iprint(f"\n❌ 网络错误，无法访问OKR系统: {e}")
                iprint("   请检查网络连接或VPN是否已开启")
                await browser_manager.close(keep_browser=True)
                return
        iprint("[浏览器] 等待浏览器加载完成...")
        await asyncio.sleep(5)
        current_url = page.url  # 更新URL

    # 4. 分析页面并检查登录状态
    #await analyze_page(page, "2当前页面")
    need_login = False
    # 更严谨的登录判断逻辑
    if (current_url and 'uuap' in current_url) or current_url in ('', 'about:blank'):
        need_login = True

    # 检查页面是否显示网络错误（"无法访问此网站"等）
    if not need_login and await _is_page_unreachable(page):
        iprint("\n❌ 页面无法访问（可能是网络不通或VPN未连接）")
        iprint("   请检查网络连接后重试")
        await browser_manager.close(keep_browser=True)
        return

    # 5. 引导用户登录（自动检测登录状态）
    if need_login:
        iprint("\n" + "="*60)
        iprint("⚠️  请在浏览器中完成OKR系统登录")
        iprint("   登录完成后将自动继续...")
        iprint("="*60)
        
        # 自动等待登录完成（轮询检测）
        login_timeout = 300  # 5分钟超时
        login_check_interval = 2  # 每2秒检测一次
        login_start = asyncio.get_event_loop().time()
        
        while True:
            try:
                current_url = page.url
                if current_url and 'uuap' not in current_url and current_url not in ('', 'about:blank'):
                    iprint("\n✓ 检测到登录成功，自动继续...")
                    break
                
                elapsed = asyncio.get_event_loop().time() - login_start
                if elapsed > login_timeout:
                    iprint("\n✗ 登录超时，退出")
                    await browser_manager.close(keep_browser=True)
                    return
                    
                remaining = int(login_timeout - elapsed)
                if int(elapsed) % 10 == 0:
                    iprint(f"   等待登录中... ({remaining}秒后超时)")
                    
                await asyncio.sleep(login_check_interval)
            except Exception as e:
                iprint(f"   登录检测异常: {e}")
                await asyncio.sleep(login_check_interval)
        
        # 仅当用户手动登录后，才重新导航（避免冗余跳转）
        iprint("\n[导航] 验证登录状态...")
        await page.goto(BASE_URL, wait_until='domcontentloaded', timeout=30000)
        current_url = page.url

    # 6. 登录状态最终验证
    #await analyze_page(page, "登录后页面")
    if 'uuap' in current_url:
        iprint("\n✗ 未登录成功，请检查账号密码或网络")
        await browser_manager.close(keep_browser=False)
        return
    # 开始正式填报主流程（含浏览器断连重试，最多3次）
    max_disconnect_retries = 3
    result = None
    for disconnect_attempt in range(max_disconnect_retries + 1):
        try:
            result = await submit_weekly_with_retry_v24(
                page=page,
                target_date=target_date,
                dry_run=DRY_RUN,
                auto_submit=AUTO_SUBMIT
            )
            break  # 成功则跳出重试循环
        except Exception as e:
            if _is_browser_disconnected(e) and disconnect_attempt < max_disconnect_retries:
                iprint(f"\n⚠️  提交过程中浏览器断连（第{disconnect_attempt+1}次），尝试重连...")
                if await browser_manager.reconnect_with_timeout(max_retries=3):
                    page = browser_manager.page
                    try:
                        await page.goto(BASE_URL, wait_until='domcontentloaded', timeout=30000)
                        await asyncio.sleep(3)
                        iprint(f"重连成功，重试提交（第{disconnect_attempt+2}次尝试）...")
                        continue  # 继续循环，重新执行submit
                    except Exception as nav_e:
                        iprint(f"❌ 重连后导航失败: {nav_e}")
                        if disconnect_attempt < max_disconnect_retries - 1:
                            continue
                        result = {'status': 'error', 'error': str(nav_e)}
                        break
                else:
                    iprint(f"❌ 浏览器重连失败，已尝试{disconnect_attempt+1}次，程序自动退出")
                    return
            elif _is_browser_disconnected(e):
                iprint(f"❌ 浏览器断连，已达最大重试次数({max_disconnect_retries})，程序自动退出")
                return
            else:
                raise

    if result is None or not isinstance(result, dict):
        iprint("❌ 提交未返回有效结果")
        return

    if result['status'] == 'login_required':
        print("\n[!] 需要重新登录")
    elif result['status'] == 'error':
        print(f"\n[X] 错误: {result.get('error')}")
    elif result['status'] == 'preview':
        print("\n[OK] 预览完成")
        print("\n要实际提交，请设置:")
        print("  DRY_RUN = False")
        print("  AUTO_SUBMIT = True")
    else:
        print("\n[OK] 填写流程完成")
        if AUTO_SUBMIT:
            if result.get('results', {}).get('submitted'):
                print("[OK] 已成功自动提交并确认")
            else:
                print("[!] 自动提交可能未完成，请检查页面")
        else:
            print("[!] 请手动点击保存按钮")

    print("\n" + "="*60)
    print("[系统] 执行完成")
    print("="*60)
    print("\n选项:")
    print("  1. 保持浏览器运行并退出（下次可快速复用）- 直接按Enter")
    print("  2. 关闭浏览器并退出 - 输入 'close' 后按Enter")

    try:
        try:
            user_input = input("\n请选择 (默认: 保持运行): ").strip().lower()
        except (EOFError, OSError, RuntimeError):
            user_input = ""  # 无控制台时使用默认值

        if user_input == 'close':
            print("[系统] 正在关闭浏览器...")
            await browser_manager.close(keep_browser=False)
        else:
            print("[系统] 保持浏览器运行，可以复用")
            await browser_manager.close(keep_browser=True)
            print(f"\n[提示] 下次运行脚本时将自动复用此浏览器")
            print(f"[提示] 状态文件: {BROWSER_STATE_FILE}")
    except KeyboardInterrupt:
        print("\n[系统] 用户中断，保持浏览器运行")
        await browser_manager.close(keep_browser=True)

    except Exception as e:
        print(f"\n[X] 严重错误: {e}")
        traceback.print_exc()
        if 'browser_manager' in locals():
            await browser_manager.close(keep_browser=True)
# ==================== 入口点 ====================
if __name__ == "__main__":
    try:
        asyncio.run(main())
    except Exception as e:
        print(f"错误: {e}")
        traceback.print_exc()
    finally:
        print("\n按回车键退出...")
        try:
            if sys.stdin and sys.stdin.isatty():
                input()  # 等待用户按回车
        except:
            pass