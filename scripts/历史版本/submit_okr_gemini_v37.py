#!/usr/bin/env python3
"""百度OKR周报自动提交脚本 - V2.4 (修复嵌套content结构)"""

import asyncio
import socket
import json
import os
from datetime import datetime, timedelta
from pathlib import Path
from playwright.async_api import async_playwright, Error as PlaywrightError
import platform
import traceback

# ==================== 工具函数 ====================

def iprint(msg):
    """智能打印，自动添加时间戳"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    print(f"[{timestamp}] {msg}")
	
# ==================== 页面上下文工具 ====================

async def get_working_context(page):
    """自动穿透进入百度周报的核心 iframe"""
    try:
        # 百度周报/OKR系统主要业务都在 id="main" 的 iframe 中
        iframe_locator = page.frame_locator('iframe#main')
        if await page.locator('iframe#main').count() > 0:
            return iframe_locatorx
    except:
        pass
    return page # 找不到 iframe 则回退到主页面上下文

async def smart_activate_and_fill(locator, content):
    """暴力填充：点击 + 注入 + 模拟按键"""
    if not locator or await locator.count() == 0:
        return False
    try:
        await locator.scroll_into_view_if_needed()
        # 激活元素
        await locator.click()
        await locator.focus()
        
        # 使用 JavaScript 强行赋值并触发 React 监听的所有事件
        await locator.evaluate(f"""(el, val) => {{
            const inputEvent = new Event('input', {{ bubbles: true, cancelable: true }});
            const changeEvent = new Event('change', {{ bubbles: true, cancelable: true }});
            
            if (el.isContentEditable) {{
                el.innerText = val;
            }} else {{
                // 针对原生 textarea 的特殊 Setter 调用
                const nativeValueSetter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, "value").set;
                nativeValueSetter.call(el, val);
            }}
            el.dispatchEvent(inputEvent);
            el.dispatchEvent(changeEvent);
        }}""", content)
        
        # 关键的一步：模拟输入一个空格再退格，强制触发表单验证
        await locator.press("End")
        await locator.press("Space")
        await asyncio.sleep(0.1)
        return True
    except Exception as e:
        iprint(f"    [!] 填充执行出错: {e}")
        return False

# ==================== 配置区域 ====================

# Chrome浏览器路径（Windows需要指定）
if platform.system() == "Windows":
    # 常见的Chrome安装路径
    CHROME_PATHS = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        r"C:\Users\%USERNAME%\AppData\Local\Google\Chrome\Application\chrome.exe"
    ]
    # 使用第一个存在的路径
    CHROME_PATH = None
    for path in CHROME_PATHS:
        expanded_path = os.path.expandvars(path)
        if os.path.exists(expanded_path):
            CHROME_PATH = expanded_path
            break
else:
    # Linux/Mac
    CHROME_PATH = "google-chrome"

MY_ID = "s673090"
EDIT_BASE_URL = "https://okr.baidu-int.com/pages/weekly.html#/home"


# ==================== 浏览器复用功能 ====================
import time
import subprocess

BROWSER_STATE_FILE = "browser_state.json"

# ==================== 浏览器复用功能 ====================

BROWSER_DEBUG_DIR = "browser_debug"


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
    
    def save_state(self):
        """保存浏览器状态到文件"""
        if self.cdp_port and self.browser_pid:
            state = {
                "cdp_port": self.cdp_port,
                "browser_pid": self.browser_pid,
                "timestamp": time.time()
            }
            with open(BROWSER_STATE_FILE, "w", encoding="utf-8") as f:
                json.dump(state, f)
            iprint(f"[浏览器] 状态已保存到 {BROWSER_STATE_FILE}")
    
    def load_state(self):
        """从文件加载浏览器状态"""
        if os.path.exists(BROWSER_STATE_FILE):
            try:
                with open(BROWSER_STATE_FILE, "r", encoding="utf-8") as f:
                    state = json.load(f)
                
                # 检查浏览器进程是否还在运行
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
        """检查进程是否在运行"""
        if not pid:
            return False
        try:
            if platform.system() == "Windows":
                # Windows下使用tasklist检查进程
                result = subprocess.run(f'tasklist /FI "PID eq {pid}"', shell=True, capture_output=True, text=True)
                return str(pid) in result.stdout
            else:
                os.kill(pid, 0)
                return True
        except:
            return False
    
    async def connect_to_existing_browser(self):
        """连接到已存在的浏览器"""
        if not self.cdp_port:
            return False
        
        try:
            # 尝试连接到已存在的浏览器
            self.playwright = await async_playwright().start()
            self.browser = await self.playwright.chromium.connect_over_cdp(
                f"http://127.0.0.1:{self.cdp_port}"
            )
            
            # 获取现有的context和page
            contexts = self.browser.contexts
            if contexts:
                self.context = contexts[0]
                pages = self.context.pages
                if pages:
                    self.page = pages[0]
                    iprint("[浏览器] 成功连接到已存在的浏览器")
                    return True
                else:
                    # 如果没有页面，创建一个新页面
                    self.page = await self.context.new_page()
                    return True
            
            iprint("[浏览器] 未找到现有的context")
            return False
        except Exception as e:
            iprint(f"[浏览器] 连接失败: {e}")
            return False
    
    async def launch_new_browser(self, headless=False):
        """启动新的浏览器"""
        if not CHROME_PATH:
            iprint("[X] 未找到Chrome浏览器，请手动安装或配置CHROME_PATH")
            return False
        
        # 创建调试目录
        os.makedirs(BROWSER_DEBUG_DIR, exist_ok=True)
        
        # 生成一个随机端口
        import random
        self.cdp_port = random.randint(9222, 9299)
        
        iprint(f"[浏览器] 启动新浏览器，调试端口: {self.cdp_port}")
        iprint(f"[浏览器] Chrome路径: {CHROME_PATH}")
        
        # 构建启动命令
        cmd = [
            CHROME_PATH,
            f"--remote-debugging-port={self.cdp_port}",
            f"--user-data-dir={os.path.abspath(BROWSER_DEBUG_DIR)}",
            "--no-first-run",
            "--no-default-browser-check",
        ]

        # 添加 headless 参数（如果需要）
        if headless:
            cmd.append("--headless=new")

        cmd.append(EDIT_BASE_URL)
        
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
            iprint(f"[浏览器] 浏览器进程PID: {self.browser_pid}")
            
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
    
    async def ensure_page(self):
        """确保有可用的页面"""
        if not self.browser:
            return False
        
        # 更新contexts和pages
        self.context = self.browser.contexts[0] if self.browser.contexts else None
        if not self.context:
            self.context = await self.browser.new_context()
        
        self.page = self.context.pages[0] if self.context.pages else None
        if not self.page:
            self.page = await self.context.new_page()
        
        return True
    
    async def close(self, keep_browser=True):
        """关闭连接，可选择是否保留浏览器"""
        if self.playwright:
            if keep_browser:
                # 只断开连接，不关闭浏览器
                await self.playwright.stop()
                iprint("[浏览器] 已断开连接，浏览器进程继续运行")
            else:
                # 关闭浏览器和连接
                if self.browser:
                    await self.browser.close()
                await self.playwright.stop()
                if self.browser_process:
                    self.browser_process.terminate()
                # 删除状态文件
                if os.path.exists(BROWSER_STATE_FILE):
                    os.remove(BROWSER_STATE_FILE)
                iprint("[浏览器] 浏览器已关闭")

# ==================== 配置区域 ====================



# ==================== 浏览器复用功能 ====================


# ==================== 浏览器复用功能 ====================


# ==================== 配置区域 ====================



def get_default_date():
    today = datetime.now()
    last_monday = today - timedelta(days=today.weekday() + 7)
    return f"s{last_monday.strftime('%Y-%m-%d')}"

DRY_RUN = False
AUTO_SUBMIT = True
JSON_FILE_PATH = "team_weekly_report_v10.json"

WEEKLY_DATA = {
    "part1_work_summary": None,
    "part2_okr_structure": None,
    "part3_metrics": None
}

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
            print(f"      [WARN] 缺少字段: {e}")
            return ", ".join([f"{k}={v}" for k, v in item.items()])
    elif isinstance(item, str):
        return fallback.format(item=item)
    return str(item)

def parse_part1(data):
    print(f"\n  [解析] 第一部分: 本周工作总结")

    # 处理直接是列表的情况
    if isinstance(data, list):
        print(f"    [INFO] 数据为列表类型，包含 {len(data)} 项")
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
                    print(f"    [INFO] 从 '{key}' 键提取内容")
                    break
        print(f"    标题: {title}")
        print(f"    content类型: {type(content_list)}")
    else:
        print(f"    [WARN] 不支持的数据类型: {type(data)}")
        return None

    # 扁平化嵌套结构
    content_list = flatten_nested_content(content_list)

    if not content_list:
        print(f"    [WARN] 内容为空")
        return None

    config = FORMAT_CONFIG["part1_work_summary"]
    formatted_items = []
    raw_items = []

    for idx, item in enumerate(content_list):
        print(f"    处理 item {idx+1}: {type(item)}")
        if isinstance(item, dict):
            raw_items.append(item)
            text = format_item(item, config["template"], config["fallback"])
            formatted_items.append(f"{idx+1}. {text}")
            print(f"      字段: {list(item.keys())}")
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

    print(f"    [OK] 完成: {result['item_count']} 条")
    print(f"    预览: {result['content'][:150]}...")
    return result

def parse_part2(data):
    print(f"\n  [解析] 第二部分: OKR进度")

    # 处理直接是列表的情况（objectives列表）
    if isinstance(data, list):
        print(f"    [INFO] 数据为列表类型，包含 {len(data)} 个Objective")
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
                    print(f"    [INFO] 从 '{key}' 键提取objectives")
                    break
        print(f"    标题: {title}")
        print(f"    Objectives: {len(objectives)}")
    else:
        print(f"    [WARN] 不支持的数据类型: {type(data)}")
        return None

    if not isinstance(objectives, list):
        print(f"    [WARN] objectives 不是列表: {type(objectives)}")
        return None

    config = FORMAT_CONFIG["part2_okr_structure"]
    okr_data = {"title": title, "objectives": []}

    for obj_idx, obj in enumerate(objectives):
        if not isinstance(obj, dict):
            print(f"    [WARN] Objective {obj_idx} 不是字典: {type(obj)}")
            continue

        o_id = obj.get("o_id", f"O{obj_idx+1}")
        o_title = obj.get("o_title", obj.get("title", ""))
        print(f"    Objective: {o_id} - {o_title}")

        o_data = {"o_id": o_id, "o_title": o_title, "key_results": []}
        key_results = obj.get("key_results", [])

        # 如果key_results不是列表，尝试其他常见键
        if not isinstance(key_results, list):
            for key in ["krs", "items", "results", "data"]:
                if key in obj and isinstance(obj[key], list):
                    key_results = obj[key]
                    print(f"      [INFO] 从 '{key}' 键提取key_results")
                    break

        if not isinstance(key_results, list):
            print(f"      [WARN] key_results 不是列表: {type(key_results)}")
            key_results = [key_results] if key_results else []

        for kr_idx, kr in enumerate(key_results):
            if not isinstance(kr, dict):
                print(f"      [WARN] KR {kr_idx} 不是字典: {type(kr)}")
                continue

            kr_id = kr.get("kr_id", f"KR{kr_idx+1}")
            kr_title = kr.get("kr_title", kr.get("title", ""))
            weekly_work = kr.get("weekly_work", "")
            progress = kr.get("progress", "")
            confidence = kr.get("confidence", "")

            print(f"      KR: {kr_id}")

            # 处理weekly_work可能是嵌套结构的情况
            if isinstance(weekly_work, dict):
                print(f"        weekly_work字段: {list(weekly_work.keys())}")
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
    print(f"    [OK] 完成: {len(okr_data['objectives'])} 个O, {total_krs} 个KR")
    return okr_data

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

def parse_part3(data):
    print(f"\n  [解析] 第三部分: 关键指标")

    # 处理直接是列表的情况
    if isinstance(data, list):
        print(f"    [INFO] 数据为列表类型，包含 {len(data)} 项")
        # 将列表转换为标准格式
        data = {"title": "关键指标", "sections": data}
    elif not isinstance(data, dict):
        print(f"    [WARN] 不支持的数据类型: {type(data)}")
        return None

    title = data.get("title", "关键指标")
    sections = data.get("sections", [])

    # 如果sections是字典，转换为列表
    if isinstance(sections, dict):
        print(f"    [INFO] sections是字典，转换为列表")
        sections_list = []
        for key, value in sections.items():
            if isinstance(value, dict):
                value["section_title"] = value.get("section_title", key)
                sections_list.append(value)
            elif isinstance(value, list):
                sections_list.append({"section_title": key, "content": value})
            else:
                sections_list.append({"section_title": key, "content": [value]})
        sections = sections_list

    print(f"    标题: {title}")
    print(f"    Sections: {len(sections)}")

    if not isinstance(sections, list):
        print(f"    [WARN] sections 不是列表: {type(sections)}")
        return None

    config = FORMAT_CONFIG["part3_metrics"]
    metrics_data = {"title": title, "sections": {}}

    for section_idx, section in enumerate(sections):
        if not isinstance(section, dict):
            print(f"    [WARN] Section {section_idx} 不是字典: {type(section)}")
            continue

        section_title = section.get("section_title", f"Section_{section_idx}")
        content_list = section.get("content", [])

        # 递归处理嵌套结构
        content_list = flatten_nested_content(content_list)

        print(f"    Section: {section_title} ({len(content_list)} 项)")

        template = config["sections"].get(section_title, config["default"])
        formatted_items = []
        raw_items = []

        for idx, item in enumerate(content_list):
            if isinstance(item, dict):
                raw_items.append(item)
                text = format_item(item, template, config["default"])
                formatted_items.append(f"{idx+1}. {text}")
                print(f"      字段: {list(item.keys())}")
            elif isinstance(item, str):
                raw_items.append({"text": item})
                formatted_items.append(f"{idx+1}. {item}")
            elif isinstance(item, (list, tuple)):
                # 处理嵌套列表
                nested_text = "; ".join([str(x) for x in item])
                raw_items.append({"nested_list": item})
                formatted_items.append(f"{idx+1}. {nested_text}")

        metrics_data["sections"][section_title] = {
            "raw_content": raw_items,
            "formatted_content": config["separator"].join(formatted_items),
            "item_count": len(formatted_items)
        }

        print(f"      [OK] {len(formatted_items)} 条")

    print(f"    [OK] 完成: {len(metrics_data['sections'])} 个section")
    return metrics_data

def parse_json_file(json_path):
    print(f"\n{'='*60}")
    print("[DEBUG] 开始解析JSON文件")
    print(f"{'='*60}")

    print(f"\n[DEBUG] 检查文件: {json_path}")
    if not os.path.exists(json_path):
        print("  [ERROR] 文件不存在")
        try:
            files = os.listdir(os.path.dirname(json_path) or ".")
            print(f"  目录文件: {files}")
        except:
            pass
        return None

    try:
        with open(json_path, "r", encoding="utf-8") as f:
            raw_content = f.read()
        print(f"  [OK] 读取: {len(raw_content)} 字符")
        print(f"  前500字符: {raw_content[:500]}")
    except Exception as e:
        print(f"  [ERROR] 读取失败: {e}")
        return None

    try:
        data = json.loads(raw_content)
        print(f"\n[DEBUG] JSON解析成功")
        print(f"  顶层键: {list(data.keys())}")
    except Exception as e:
        print(f"  [ERROR] JSON解析失败: {e}")
        return None

    # 检查是否在'content'键下
    if "content" in data:
        print(f"\n[DEBUG] 发现'content'键，使用嵌套结构")
        content_data = data.get("content", {})
        print(f"  content键下的键: {list(content_data.keys())}")
    else:
        print(f"\n[DEBUG] 使用顶层结构")
        content_data = data

    print(f"\n{'='*60}")
    print("开始解析各部分内容")
    print(f"{'='*60}")

    result = {
        "part1_work_summary": None,
        "part2_okr_structure": None,
        "part3_metrics": None
    }

    # 从content_data中解析各部分
    if "part1_work_summary" in content_data:
        print(f"\n  [OK] 找到 part1_work_summary")
        result["part1_work_summary"] = parse_part1(content_data["part1_work_summary"])
    else:
        print(f"\n  [WARN] 未找到 part1_work_summary")
        print(f"  可用键: {list(content_data.keys())}")

    if "part2_okr_structure" in content_data:
        print(f"\n  [OK] 找到 part2_okr_structure")
        result["part2_okr_structure"] = parse_part2(content_data["part2_okr_structure"])
    else:
        print(f"\n  [WARN] 未找到 part2_okr_structure")

    if "part3_metrics" in content_data:
        print(f"\n  [OK] 找到 part3_metrics")
        result["part3_metrics"] = parse_part3(content_data["part3_metrics"])
    else:
        print(f"\n  [WARN] 未找到 part3_metrics")

    print(f"\n{'='*60}")
    print("[DEBUG] 解析结果总结")
    print(f"{'='*60}")
    print(f"  第一部分: {'OK' if result['part1_work_summary'] else 'FAIL'}")
    print(f"  第二部分: {'OK' if result['part2_okr_structure'] else 'FAIL'}")
    print(f"  第三部分: {'OK' if result['part3_metrics'] else 'FAIL'}")

    return result

def load_content_from_json():
    global WEEKLY_DATA

    print(f"\n{'='*60}")
    print("[系统] 加载JSON文件")
    print(f"{'='*60}")

    try:
        script_dir = Path(__file__).parent.absolute()
    except:
        script_dir = Path(os.getcwd())

    json_path = script_dir / JSON_FILE_PATH
    print(f"文件路径: {json_path}")

    parsed = parse_json_file(str(json_path))

    if parsed:
        WEEKLY_DATA["part1_work_summary"] = parsed.get("part1_work_summary")
        WEEKLY_DATA["part2_okr_structure"] = parsed.get("part2_okr_structure")
        WEEKLY_DATA["part3_metrics"] = parsed.get("part3_metrics")

        print(f"\n[OK] 数据加载成功")
        print(f"\n[验证] 数据预览:")
        if WEEKLY_DATA["part1_work_summary"]:
            print(f"  第一部分: {len(WEEKLY_DATA['part1_work_summary']['content'])} 字符")
        if WEEKLY_DATA["part2_okr_structure"]:
            print(f"  第二部分: {len(WEEKLY_DATA['part2_okr_structure']['objectives'])} 个O")
        if WEEKLY_DATA["part3_metrics"]:
            print(f"  第三部分: {len(WEEKLY_DATA['part3_metrics']['sections'])} 个section")

        return True
    else:
        print(f"\n[ERROR] 加载失败")
        return False

def check_network_connection():
    print("\n[网络检测] 检查网络连接...")
    try:
        ip = socket.gethostbyname("okr.baidu-int.com")
        print(f"  [OK] 内网解析正常: {ip}")
        return True
    except:
        print(f"  [X] 内网解析失败，请检查VPN")
        return False

# ==================== 页面操作 ====================

async def analyze_page_structure(page):
    print("\n[页面分析] 分析页面结构...")
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
    print(f"  第一部分: {'存在' if structure['has_part1'] else '不存在'}")
    print(f"  第二部分: {'存在' if structure['has_part2'] else '不存在'}")
    print(f"  第三部分: {'存在' if structure['has_part3'] else '不存在'}")
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
    print("    滚动加载...", end=" ", flush=True)
    await page.mouse.move(center[0], center[1])
    for i in range(30):
        await page.mouse.wheel(0, 400)
        await asyncio.sleep(0.3)
        if i % 10 == 0:
            print(f"{i*400}px", end=" ", flush=True)
    print("OK")
    print("    等待渲染...", end=" ", flush=True)
    await asyncio.sleep(5)
    print("OK")

# ==================== 各部分填写函数 (优化版) ====================
async def ask_user(msg):
    """在终端暂停并等待用户确认"""
    print(f"\n>>>> [人工确认] {msg}")
    print(">>>> 按 Enter 键 [继续执行] | 输入 's' [跳过此步] | 输入 'q' [退出脚本]")
    choice = await asyncio.to_thread(input, "等待指令... ")
    if choice.lower() == 'q':
        exit()
    return choice.lower() != 's'

async def draw_debug_dot(page, x, y):
    """在页面指定坐标绘制一个红色圆点，用于人工校验点击位置"""
    await page.evaluate(f"""() => {{
        let dot = document.getElementById('ai-target-dot') || document.createElement('div');
        dot.id = 'ai-target-dot';
        dot.style.position = 'fixed'; 
        dot.style.width = '16px'; 
        dot.style.height = '16px';
        dot.style.backgroundColor = 'rgba(255, 0, 0, 0.8)'; 
        dot.style.borderRadius = '50%';
        dot.style.zIndex = '999999'; 
        dot.style.pointerEvents = 'none';
        dot.style.border = '2px solid white'; 
        document.body.appendChild(dot);
        dot.style.left = '{x - 8}px'; 
        dot.style.top = '{y - 8}px';
        dot.style.display = 'block';
    }}""")


async def fill_part1_work_summary(page, part1_data):
    """第一部分：锁定末尾 + 双向换行隔离"""
    iprint("\n[填写] 第一部分：本周工作总结")
    # 前后追加换行符，确保不干扰已有内容
    content = "\n\n" + part1_data.get("content", "") + "\n\n"
    ctx = await get_working_context(page)
    
    anchor = ctx.get_by_text("本周工作总结", exact=True).first
    if await anchor.count() > 0:
        await anchor.scroll_into_view_if_needed()
        box = await anchor.bounding_box()
        
        if box:
            target_x = box['x'] + 5 
            target_y = box['y'] + box['height'] + 15 

            await draw_debug_dot(page, target_x, target_y)
            iprint(f"红点已定位。坐标: ({int(target_x)}, {int(target_y)})")
            await ask_user("确认红点是否在总结输入框起始位置？确认后将强制锁定末尾输入。")

            # 物理点击激活
            await page.mouse.click(target_x, target_y)
            await asyncio.sleep(0.8) # 等待页面脚本反应，防止光标跳动

            # 强制光标到最后
            await page.keyboard.press("Control+End")
            await asyncio.sleep(0.2)

            iprint("正在注入内容（含前后换行隔离）...")
            await page.keyboard.insert_text(content)
            
            await ask_user("第一部分追加完成。请确认原有内容是否受影响？")
            return True
    return False
	


async def fill_part2_okr_weekly_work(page, part2_data):
    """第二部分：5px 偏移 + 十字探测激活输入框"""
    iprint("\n[填写] 第二部分：OKR 进度更新")
    objectives = part2_data.get("objectives", [])
    ctx = await get_working_context(page)

    for obj in objectives:
        for kr_idx, kr in enumerate(obj.get("key_results", [])):
            kr_id = f"KR{kr_idx + 1}"
            # 同样在内容前后追加换行
            weekly_content = "\n\n" + kr.get("weekly_work", "") + "\n\n"
            
            # 定位 KR 锚点
            kr_anchor = ctx.get_by_text(kr_id, exact=False).first
            if await kr_anchor.count() > 0:
                await kr_anchor.scroll_into_view_if_needed()
                await kr_anchor.evaluate("el => el.style.backgroundColor = 'yellow'")
                
                # 在 KR 后寻找“本周工作”标签
                work_label = ctx.locator(f"xpath=//*[contains(text(), '{kr_id}')]/following::*[text()='本周工作' or contains(., '本周工作')]").first
                
                if await work_label.count() > 0:
                    l_box = await work_label.bounding_box()
                    
                    # --- 精准 5px 偏移定位 ---
                    target_x = l_box['x'] + 5 
                    target_y = l_box['y'] + l_box['height'] + 5 

                    await draw_debug_dot(page, target_x, target_y)
                    await page.mouse.move(target_x, target_y)
                    await ask_user(f"红点已对齐 [{kr_id}] ‘本’字下方 5px。准备执行十字探测激活？")

                    # --- 十字探测激活逻辑 ---
                    # 在 5px 处及周边微调点击，确保激活光标
                    for dx, dy in [(0,0), (2,0), (-2,0), (0,2)]:
                        await page.mouse.click(target_x + dx, target_y + dy)
                        await asyncio.sleep(0.1)

                    # 尝试移动光标确保可输入
                    await page.keyboard.press("End")
                    await asyncio.sleep(0.3)

                    iprint(f"  [注入] 正在 5px 处输入 {kr_id} 内容...")
                    await page.keyboard.insert_text(weekly_content)
                    
                    await ask_user(f"{kr_id} 填写完成。确认后处理下一项。")
                    await page.evaluate("document.getElementById('ai-target-dot').style.display = 'none'")
    return True


async def fill_part3_metrics(page, part3_data):
    """填写第三部分：关键指标 - 增强变体匹配"""
    iprint("\n[填写] 第三部分：关键指标")
    sections = part3_data.get("sections", {})
    if not sections: return False

    ctx = await get_working_context(page)
    success_count = 0

    for name, data in sections.items():
        content = data.get("formatted_content", data.get("raw_content", ""))
        if not content: continue

        # 处理标签变体
        search_labels = [name]
        if "指标" in name: search_labels.extend(["核心指标", "业务数据"])
        if "下周" in name: search_labels.extend(["下周重点", "未来计划"])

        filled = False
        for label_text in search_labels:
            # 这里的 logic 参考 smart_fill_input，但使用 ctx 穿透
            label_locator = ctx.get_by_text(label_text, exact=False).first
            if await label_locator.count() > 0:
                parent = label_locator.locator('xpath=ancestor::div[contains(@class, "section")] | ancestor::div[2]').first
                target = parent.locator("textarea.ui-textarea-element, textarea, [contenteditable='true']").first
                if await smart_activate_and_fill(target, content):
                    iprint(f"  [OK] {name} 填写成功")
                    filled = True
                    success_count += 1
                    break
        
        if not filled:
            iprint(f"  [!] 未能通过标签定位到: {name}")

    return success_count > 0



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
        print("    [OK] 已清理遮挡层")
        return True
    except Exception as e:
        print(f"    [!] 清理遮挡层失败: {e}")
        return False
		
async def smart_activate_and_fill(locator, content):
    """暴力填充：JS 注入 + 事件强制触发表单同步"""
    if not locator or await locator.count() == 0:
        return False
    try:
        await locator.scroll_into_view_if_needed()
        await locator.click()
        
        # 使用 JS 强制修改值，并手动分发 input/change 事件，让 React 知道有输入
        await locator.evaluate(f"""(el, val) => {{
            if (el.isContentEditable) {{
                el.innerText = val;
            }} else {{
                // 解决 React 不同步问题的 Setter 调用
                const nativeValueSetter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, "value").set;
                nativeValueSetter.call(el, val);
            }}
            el.dispatchEvent(new Event('input', {{ bubbles: true }}));
            el.dispatchEvent(new Event('change', {{ bubbles: true }}));
        }}""", content)
        
        # 补一刀模拟按键，确保触发页面的‘自动保存’
        await locator.focus()
        await locator.press("End")
        await locator.press("Space") 
        return True
    except:
        return False

async def smart_fill_input(page, label_text, content, create_if_missing=False, section_type=None):
    if not content or not content.strip():
        print(f"  [跳过] 内容为空: {label_text}")
        return False

    content = content.strip()
    print(f"\n  处理: {label_text[:40]}...")
    print(f"    内容长度: {len(content)}")
    print(f"    预览: {content[:80]}...")

    variants = [label_text]
    if "本周工作总结" in label_text:
        variants.extend(["工作总结", "本周工作"])
    elif "OKR" in label_text:
        variants.extend(["OKR进度", "OKR"])

    for variant in variants:
        try:
            locator = page.get_by_text(variant, exact=False).first
            if await locator.is_visible(timeout=2000):
                print(f"    [OK] 找到标签: {variant}")
                # 使用新的点击方式
                if await smart_click_with_fallback(page, locator):
                    await asyncio.sleep(0.3)
                    if await find_and_fill_nearby_input(page, locator, content):
                        return True
        except Exception as e:
            print(f"    [!] 处理 {variant} 时出错: {e}")
            continue

    print("    尝试JavaScript直接填写...")
    if await fill_by_javascript(page, label_text, content):
        return True

    if create_if_missing and section_type:
        print("    尝试创建新区域...")
        if await create_new_section(page, section_type, label_text, content):
            return True

    print(f"    [X] 失败: {label_text}")
    return False
async def find_and_fill_nearby_input(page, label_locator, content):
    try:
        # 使用智能点击代替直接点击
        if not await smart_click_with_fallback(page, label_locator):
            print("    [!] 标签点击失败，尝试直接查找输入框...")

        await asyncio.sleep(0.5)
        handle = await label_locator.element_handle()
        if not handle:
            return False
        text = await handle.inner_text()
        selectors = ['textarea', 'input[type="text"]', '[contenteditable="true"]', '.ant-input']
        for sel in selectors:
            try:
                inp = page.locator(f'xpath=//*[contains(text(), "{text}")]/following::{sel}[1] | xpath=//*[contains(text(), "{text}")]/ancestor::div[1]//{sel}').first
                if await inp.is_visible(timeout=2000):
                    print(f"    [OK] 找到输入框: {sel}")
                    await inp.fill(content)
                    await asyncio.sleep(0.5)
                    return True
            except:
                continue
    except Exception as e:
        print(f"    [!] 失败: {e}")
    return False
async def fill_by_javascript(page, label_text, content):
    print(f"    [DEBUG-JS] 开始JavaScript填写，标签: {label_text[:30]}...")
    print(f"    [DEBUG-JS] 内容长度: {len(content)} 字符")

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

        print(f"    [DEBUG-JS] 执行结果: {result}")
        if result == 'ok':
            print(f"    [OK] JavaScript填写成功")
            await asyncio.sleep(0.5)
            return True
        else:
            print(f"    [!] JavaScript返回: {result}")
    except Exception as e:
        print(f"    [!] JavaScript错误: {e}")
        import traceback
        traceback.print_exc()
    return False





async def create_new_section(page, section_type, title, content):
    print(f"    [DEBUG] 创建新区域: {title[:30]}...")
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
            print(f"    [OK] 创建成功")
            return True
    except Exception as e:
        print(f"    [X] 创建失败: {e}")
    return False


# ==================== 提交和确认 ====================

async def auto_submit_and_confirm(page):
    print("\n[提交] 自动提交...")

    submit_clicked = False
    for text in ["保存", "提交", "更新", "Save", "Submit"]:
        try:
            button = page.get_by_role("button", name=text).first
            if await button.is_visible(timeout=2000):
                print(f"  点击: {text}")
                await button.click()
                submit_clicked = True
                break
        except:
            continue

    if not submit_clicked:
        result = await page.evaluate("""() => {
            const buttons = Array.from(document.querySelectorAll('button, input[type="submit"]'));
            for (let btn of buttons) {
                const text = (btn.textContent || btn.value || '').trim();
                if ((text.includes('保存') || text.includes('提交')) && btn.offsetParent !== null) {
                    btn.click();
                    return true;
                }
            }
            return false;
        }""")
        if result:
            submit_clicked = True
            print("  JavaScript点击提交")

    if not submit_clicked:
        print("  [X] 未找到提交按钮")
        return False

    print("\n[确认] 处理确认弹框...")
    await asyncio.sleep(1.5)

    confirm_clicked = False
    for text in ["确认", "确定", "Yes", "OK", "Confirm"]:
        try:
            confirm_btn = page.get_by_role("dialog").get_by_role("button", name=text).first
            if await confirm_btn.is_visible(timeout=3000):
                print(f"  点击确认: {text}")
                await confirm_btn.click()
                confirm_clicked = True
                break
        except:
            try:
                confirm_btn = page.get_by_role("button", name=text).first
                if await confirm_btn.is_visible(timeout=2000):
                    await confirm_btn.click()
                    confirm_clicked = True
                    break
            except:
                continue

    if not confirm_clicked:
        result = await page.evaluate("""() => {
            const dialogs = document.querySelectorAll('[role="dialog"], .modal, .ant-modal');
            for (let dialog of dialogs) {
                if (dialog.style.display === 'none') continue;
                const buttons = dialog.querySelectorAll('button');
                for (let btn of buttons) {
                    const text = (btn.textContent || '').trim();
                    if ((text.includes('确认') || text.includes('确定')) && btn.offsetParent !== null) {
                        btn.click();
                        return true;
                    }
                }
            }
            return false;
        }""")
        if result:
            confirm_clicked = True
            print("  JavaScript点击确认")

    if confirm_clicked:
        print("  [OK] 提交完成")
        await asyncio.sleep(2)
        return True
    return False

# ==================== 主流程 ====================

async def submit_single_weekly_v24(page, target_date=None, dry_run=True, auto_submit=False, is_first_page=False):
    target_date = target_date or get_default_date()
    edit_url = f"{EDIT_BASE_URL}?date={target_date}&mandatorId={MY_ID}&isEdit=true"

    print("="*60)
    print(f"目标日期: {target_date}")
    print(f"编辑模式: {'预览' if dry_run else '实际填写'}")
    print(f"自动提交: {'是' if auto_submit else '否'}")
    print("="*60)

    try:
        try:
            await page.goto(edit_url, wait_until='domcontentloaded', timeout=60000)
        except Exception as nav_error:
            if 'uuap' in page.url:
                return {'status': 'login_required', 'url': page.url}

        if is_first_page:
            print("\n(首次页面等待5秒)...", end=" ", flush=True)
            await asyncio.sleep(5)
            print("OK")
        else:
            await asyncio.sleep(2)

        if 'uuap' in page.url:
            return {'status': 'login_required', 'url': page.url}

        print("[OK] 页面加载完成")

    except Exception as e:
        print(f"\n[X] 页面加载失败: {e}")
        return {'status': 'error', 'error': str(e)}

    await analyze_page_structure(page)

    print("\n滚动加载页面内容...")
    selector, center = await find_scroll_container(page)
    await scroll_container(page, selector, center)

    print("\n" + "="*60)
    print("[验证] 数据状态检查")
    print("="*60)
    print(f"  part1_work_summary: {WEEKLY_DATA.get('part1_work_summary') is not None}")
    print(f"  part2_okr_structure: {WEEKLY_DATA.get('part2_okr_structure') is not None}")
    print(f"  part3_metrics: {WEEKLY_DATA.get('part3_metrics') is not None}")

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

    print(f"\n[填写前数据检查]")
    print(f"  part1_data: {part1_data is not None}")
    print(f"  part2_data: {part2_data is not None}")
    print(f"  part3_data: {part3_data is not None}")

    if part1_data:
        results['part1'] = await fill_part1_work_summary(page, part1_data)
    else:
        print("\n[!] 无第一部分数据，跳过")
        results['part1'] = False

    if part2_data:
        results['part2'] = await fill_part2_okr_weekly_work(page, part2_data)
    else:
        print("\n[!] 无第二部分数据，跳过")
        results['part2'] = False

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
    print(f"    [创建区域] 创建 '{title}' 区域...")

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
            print(f"    [OK] 已创建 '{title}' 区域并填写内容")
            return True
        else:
            print(f"    [!] 创建区域返回: {result}")
            return False

    except Exception as e:
        print(f"    [X] 创建区域失败: {e}")
        return False


async def main():
    print("\n" + "="*60)
    print("程序启动 - 开始加载数据")
    print("="*60)

    json_loaded = load_content_from_json()

    if not check_network_connection():
        print("\n[!] 网络检查失败")
        response = input("\n是否仍要继续? (y/N): ")
        if response.lower() != 'y':
            return

    target_date = get_default_date()

    print("="*60)
    print("百度OKR周报自动提交 - V2.4 (修复嵌套content结构)")
    print("="*60)
    print("\n【V2.4改进】")
    print("  ✓ 支持嵌套在'content'键下的JSON结构")
    print("  ✓ 支持content字段下的详细关键字段解析")
    print("  ✓ 可配置的格式化模板")
    print("  ✓ 详细的调试信息")
    print(f"\n[配置]")
    print(f"  用户ID: {MY_ID}")
    print(f"  目标日期: {target_date}")
    print(f"  编辑模式: {'预览' if DRY_RUN else '实际填写'}")
    print(f"  自动提交: {'是' if AUTO_SUBMIT else '否'}")
    print(f"\n[内容来源]")
    if json_loaded:
        print(f"  ✓ {JSON_FILE_PATH} 加载成功")
        print(f"\n  [数据状态确认]")
        print(f"    part1_work_summary: {WEEKLY_DATA.get('part1_work_summary') is not None}")
        print(f"    part2_okr_structure: {WEEKLY_DATA.get('part2_okr_structure') is not None}")
        print(f"    part3_metrics: {WEEKLY_DATA.get('part3_metrics') is not None}")

        if WEEKLY_DATA.get("part1_work_summary"):
            print(f"    - 工作总结: {len(WEEKLY_DATA['part1_work_summary']['content'])} 字符")
        if WEEKLY_DATA.get("part2_okr_structure"):
            kr_count = sum(len(o.get("key_results", [])) for o in WEEKLY_DATA["part2_okr_structure"].get("objectives", []))
            print(f"    - OKR工作: {kr_count} 条KR")
        if WEEKLY_DATA.get("part3_metrics"):
            print(f"    - 关键指标: {len(WEEKLY_DATA['part3_metrics'].get('sections', {}))} 项")
    else:
        print(f"  ✗ 未能加载数据")

    print(f"\n{'='*60}")
    print("[系统] 初始化浏览器管理器...")

    # 使用浏览器管理器
    browser_manager = BrowserManager()

    # 尝试连接到已存在的浏览器
    browser_connected = False
    if browser_manager.load_state():
        browser_connected = await browser_manager.connect_to_existing_browser()

    if not browser_connected:
        # 启动新浏览器
        print("\n[浏览器] 启动新浏览器...")
        if not await browser_manager.launch_new_browser(headless=False):
            print("[X] 浏览器启动失败")
            return

    # 确保页面可用
    await browser_manager.ensure_page()
    page = browser_manager.page
    print("[OK] 浏览器已就绪")

    try:
        # 检查当前URL，如果不是目标页面则导航
        current_url = page.url
        base_url = "https://okr.baidu-int.com"

        if base_url not in current_url:
            print(f"\n[导航] 当前URL: {current_url}")
            print(f"[导航] 导航到OKR首页...")
            try:
                await page.goto(base_url, wait_until='domcontentloaded', timeout=30000)
                print("[OK] 页面已加载")
            except PlaywrightError as e:
                if "ERR_NAME_NOT_RESOLVED" in str(e):
                    print(f"\n[X] 网络错误，无法访问OKR系统")
                    await browser_manager.close(keep_browser=False)
                    return
                else:
                    raise
        else:
            print(f"\n[导航] 已在OKR页面: {current_url}")

        # 检查登录状态（通过查找密码输入框）
        try:
            password_input = page.locator('input[type="password"]').first
            if await password_input.is_visible(timeout=5000):
                print("\n" + "="*60)
                print("[!] 检测到登录页面，请先手动登录")
                print("="*60)
                input("\n登录完成后按 Enter 继续...")
        except:
            # 没有找到密码输入框，可能已经登录
            pass

        if 'uuap' in page.url:
            print("\n[X] 未登录成功或会话已过期")
            await browser_manager.close(keep_browser=False)
            return
        print("[OK] 登录验证通过")

        result = await submit_weekly_with_retry_v24(
            page=page,
            target_date=target_date,
            dry_run=DRY_RUN,
            auto_submit=AUTO_SUBMIT
        )

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

# 在 main 函数末尾执行提交的部分
        if AUTO_SUBMIT and not DRY_RUN:
            iprint("[提交] 正在准备提交...")
            ctx = await get_working_context(page)
            # 定位 iframe 内部最下方的“确定”或“保存”按钮
            btn_submit = ctx.locator("button:has-text('确定'), button:has-text('保存')").last
            
            if await btn_submit.is_visible():
                # 清理掉可能挡住按钮的 AI 助手或蒙层
                await ctx.evaluate("() => document.querySelectorAll('.mp_ai_plugin_shade, .erp-guide-mask').forEach(e => e.remove())")
                await btn_submit.click()
                iprint("[OK] 提交按钮已点击")
                result['results']['submitted'] = True
            else:
                iprint("[X] 无法定位提交按钮")

        print("\n" + "="*60)
        print("[系统] 执行完成")
        print("="*60)
        print("\n选项:")
        print("  1. 保持浏览器运行并退出（下次可快速复用）- 直接按Enter")
        print("  2. 关闭浏览器并退出 - 输入 'close' 后按Enter")

        try:
            user_input = input("\n请选择 (默认: 保持运行): ").strip().lower()

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

if __name__ == "__main__":
    asyncio.run(main())