#!/usr/bin/env python3
"""百度OKR周报自动提交脚本 - V2.4 (修复嵌套content结构)"""

import asyncio
import socket
import json
import os
import traceback
from datetime import datetime, timedelta
from pathlib import Path
from playwright.async_api import async_playwright, Error as PlaywrightError
import platform

# ==================== 工具函数 ====================

def iprint(msg):
    """智能打印，自动添加时间戳"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    print(f"[{timestamp}] {msg}")

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
    print(f"    预览: {result['content'][:1000]}...")
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


# ==================== 各部分填写函数 ====================
    """模拟完整用户输入（备选方案）"""
    
    print("\n[模拟输入] 开始模拟用户输入...")
    
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
        }""")
        
        await asyncio.sleep(1)
        
        # 2. 按回车创建新段落
        await page.keyboard.press('Enter')
        await asyncio.sleep(0.5)
        
        # 3. 逐字输入内容（带延迟，模拟真实输入）
        print("    开始逐字输入...")
        for i, char in enumerate(content):
            await page.keyboard.type(char)
            # 每10个字符稍微延迟一下
            if i % 10 == 0:
                await asyncio.sleep(0.1)
        
        # 4. 按回车添加换行
        await page.keyboard.press('Enter')
        await asyncio.sleep(0.3)
        
        # 5. 按Tab键移动焦点，触发blur
        await page.keyboard.press('Tab')
        await asyncio.sleep(0.3)
        
        # 6. 点击页面其他地方，确保编辑器失去焦点
        await page.mouse.click(10, 10)
        
        print("  [OK] 模拟输入完成")
        
        # 7. 等待用户验证
        response = input("\n  模拟输入完成，请手动保存后输入 y 确认内容被保留 (y/n): ").strip().lower()
        return response == 'y'
        
    except Exception as e:
        print(f"  [X] 模拟输入失败: {e}")
        return False
    """模拟完整用户输入（备选方案）"""
    
    print("\n[模拟输入] 开始模拟用户输入...")
    
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
        }""")
        
        await asyncio.sleep(1)
        
        # 2. 按回车创建新段落
        await page.keyboard.press('Enter')
        await asyncio.sleep(0.5)
        
        # 3. 逐字输入内容（带延迟，模拟真实输入）
        print("    开始逐字输入...")
        for i, char in enumerate(content):
            await page.keyboard.type(char)
            # 每10个字符稍微延迟一下
            if i % 10 == 0:
                await asyncio.sleep(0.1)
        
        # 4. 按回车添加换行
        await page.keyboard.press('Enter')
        await asyncio.sleep(0.3)
        
        # 5. 按Tab键移动焦点，触发blur
        await page.keyboard.press('Tab')
        await asyncio.sleep(0.3)
        
        # 6. 点击页面其他地方，确保编辑器失去焦点
        await page.mouse.click(10, 10)
        
        print("  [OK] 模拟输入完成")
        
        # 7. 等待用户验证
        response = input("\n  模拟输入完成，请手动保存后输入 y 确认内容被保留 (y/n): ").strip().lower()
        return response == 'y'
        
    except Exception as e:
        print(f"  [X] 模拟输入失败: {e}")
        return False

async def fill_part1_work_summary(page, part1_data):
    """填写第一部分：本周工作总结 - 强化自动点击版"""
    
    print("\n" + "="*80)
    print("[填写] 第一部分：本周工作总结（强化点击版）")
    print("="*80)
    
    content = part1_data.get("content", "")
    if not content:
        print("  [X] 内容为空")
        return False
    
    # 清理不可见字符
    import re
    clean_content = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f\u200b]', '', content)
    
    print(f"  清理后内容长度: {len(clean_content)} 字符")
    print(f"  预览: {clean_content[:100]}...")
    
    try:
        # ==================== 第一步：滚动到标题位置 ====================
        print("\n" + "-"*60)
        print("  [步骤1] 滚动到标题位置")
        print("-"*60)
        
        # 清除可能存在的旧标记
        await page.evaluate("""() => {
            document.querySelectorAll('[style*="border"]').forEach(el => {
                el.style.border = '';
                el.style.backgroundColor = '';
            });
            document.querySelectorAll('.debug-marker').forEach(el => el.remove());
        }""")
        
        # 找到标题并滚动到视图
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
        
        print(f"  [OK] 已滚动到标题")
        await asyncio.sleep(1)  # 等待滚动完成
        
        # ==================== 第二步：获取精确位置并高亮 ====================
        print("\n" + "-"*60)
        print("  [步骤2] 获取精确位置")
        print("-"*60)
        
        title_info = await page.evaluate("""() => {
            const result = {
                success: false,
                clickX: 0,
                clickY: 0,
                hasTextElement: false
            };
            
            // 重新获取标题
            const headings = document.querySelectorAll('h1, h2, h3, .mp-heading');
            let titleElement = null;
            
            for (const h of headings) {
                if (h.textContent && h.textContent.includes('本周工作总结')) {
                    titleElement = h;
                    break;
                }
            }
            
            if (!titleElement) return result;
            
            // 高亮标题
            titleElement.style.border = '2px solid red';
            titleElement.style.backgroundColor = 'rgba(255, 200, 200, 0.3)';
            
            // 查找标题内的文字元素
            const textElements = titleElement.querySelectorAll('span[data-slate-string="true"]');
            
            if (textElements.length > 0) {
                // 使用最后一个文字元素
                const lastText = textElements[textElements.length - 1];
                const textRect = lastText.getBoundingClientRect();
                
                // 点击位置：文字右侧15像素（避免太近选中文字）
                result.clickX = textRect.right + 15;
                result.clickY = textRect.top + textRect.height / 2;
                result.hasTextElement = true;
                
                // 高亮文字元素
                lastText.style.backgroundColor = 'rgba(255, 255, 0, 0.5)';
                lastText.style.border = '1px solid orange';
            } else {
                // 降级方案
                const rect = titleElement.getBoundingClientRect();
                result.clickX = rect.right - 20;
                result.clickY = rect.top + rect.height / 2;
            }
            
            // 添加绿色标记（更明显的标记）
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
            
            // 添加箭头指示
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
        
        print(f"  [OK] 找到目标位置")
        print(f"  点击坐标: ({title_info['clickX']}, {title_info['clickY']})")
        print(f"  找到内部文字元素: {title_info.get('hasTextElement', False)}")
        
        input("\n  请确认绿色闪烁圆点和箭头指向的位置，按Enter继续...")
        
        # ==================== 第三步：强化点击激活 ====================
        print("\n" + "-"*60)
        print("  [步骤3] 强化点击激活")
        print("-"*60)
        
        click_x = title_info['clickX']
        click_y = title_info['clickY']
        
        # 多次点击策略
        click_strategies = [
            ("单击", 1),
            ("双击", 2),
            ("三连击", 3)
        ]
        
        for strategy_name, click_count in click_strategies:
            print(f"    尝试{strategy_name}...")
            for i in range(click_count):
                await page.mouse.click(click_x, click_y)
                await asyncio.sleep(0.1)
            await asyncio.sleep(0.3)
            
            # 验证是否激活
            is_active = await page.evaluate("""() => {
                const active = document.activeElement;
                const editor = document.querySelector('[data-slate-editor="true"]');
                return active === editor;
            }""")
            
            if is_active:
                print(f"    [OK] {strategy_name}成功激活")
                break
        else:
            print("  [⚠️] 自动点击未能激活")
            print("  请手动点击绿色箭头指向的位置，然后按Enter")
            input("    按Enter继续...")
        
        # ==================== 第四步：验证光标 ====================
        print("\n" + "-"*60)
        print("  [步骤4] 验证光标")
        print("-"*60)
        
        # 输入测试字符
        await page.keyboard.type('测试', delay=50)
        await asyncio.sleep(0.3)
        
        has_test = await page.evaluate("""() => {
            const editor = document.querySelector('[data-slate-editor="true"]');
            if (!editor) return false;
            return editor.textContent.includes('测试');
        }""")
        
        if has_test:
            print("  [OK] 光标已激活")
            # 删除测试字符
            for _ in range(2):
                await page.keyboard.press('Backspace')
                await asyncio.sleep(0.1)
        else:
            print("  [❌] 光标仍未激活，请手动点击后继续")
            input("    按Enter继续...")
        
        # ==================== 第五步：插入空行 ====================
        print("\n" + "-"*60)
        print("  [步骤5] 插入空行")
        print("-"*60)
        
        # 输入2个空格
        print("    输入空格...")
        for i in range(2):
            await page.keyboard.press('Space')
            await asyncio.sleep(0.1)
        
        # 按两次回车
        print("    按两次回车...")
        await page.keyboard.press('Enter')
        await asyncio.sleep(0.3)
        await page.keyboard.press('Enter')
        await asyncio.sleep(0.3)
        
        # 按上箭头回到第一行
        await page.keyboard.press('ArrowUp')
        await asyncio.sleep(0.2)
        
        # ==================== 第六步：输入内容 ====================
        print("\n" + "-"*60)
        print("  [步骤6] 输入内容")
        print("-"*60)
        
        print("    开始输入...")
        lines = clean_content.split('\n')
        for i, line in enumerate(lines):
            if line.strip():
                print(f"      第 {i+1} 行: {line[:30]}...")
                await page.keyboard.type(line, delay=30)
                if i < len(lines) - 1:
                    await page.keyboard.press('Enter')
                    await asyncio.sleep(0.2)
        
        await page.keyboard.press('Enter')
        await asyncio.sleep(0.3)
        
        print("    [OK] 输入完成")
        
        # ==================== 第七步：触发保存 ====================
        print("\n" + "-"*60)
        print("  [步骤7] 触发保存")
        print("-"*60)
        
        await page.mouse.click(10, 10)
        await asyncio.sleep(0.5)
        
        try:
            save_button = await page.get_by_role("button", name="保存").first
            if await save_button.count() > 0:
                await save_button.click()
                print("    [OK] 已点击保存")
        except:
            pass
        
        # 清除标记
        await page.evaluate("""() => {
            document.querySelectorAll('.debug-marker').forEach(el => el.remove());
            document.querySelectorAll('[style*="border"]').forEach(el => {
                el.style.border = '';
                el.style.backgroundColor = '';
            });
        }""")
        
        print("\n  [✅] 所有步骤完成")
        
        response = input("\n  请检查内容是否正确插入，输入 y 确认 (y/n): ").strip().lower()
        return response == 'y'
        
    except Exception as e:
        print(f"\n  [X] 填写过程出错: {e}")
        import traceback
        traceback.print_exc()
        return False
    """填写第一部分：本周工作总结 - 滚动到视图版"""
    
    print("\n" + "="*80)
    print("[填写] 第一部分：本周工作总结（滚动视图版）")
    print("="*80)
    
    content = part1_data.get("content", "")
    if not content:
        print("  [X] 内容为空")
        return False
    
    # 清理不可见字符
    import re
    clean_content = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f\u200b]', '', content)
    
    print(f"  清理后内容长度: {len(clean_content)} 字符")
    print(f"  预览: {clean_content[:100]}...")
    
    try:
        # ==================== 第一步：滚动到标题位置 ====================
        print("\n" + "-"*60)
        print("  [步骤1] 滚动到标题位置")
        print("-"*60)
        
        # 清除可能存在的旧标记
        await page.evaluate("""() => {
            document.querySelectorAll('[style*="border"]').forEach(el => {
                el.style.border = '';
                el.style.backgroundColor = '';
            });
            document.querySelectorAll('.debug-marker').forEach(el => el.remove());
        }""")
        
        # 找到标题并滚动到视图
        scroll_result = await page.evaluate("""() => {
            const result = {
                success: false,
                message: '',
                titleText: '',
                titleRect: null
            };
            
            // 找到标题
            const headings = document.querySelectorAll('h1, h2, h3, .mp-heading');
            let titleElement = null;
            
            for (const h of headings) {
                if (h.textContent && h.textContent.includes('本周工作总结')) {
                    titleElement = h;
                    result.titleText = h.textContent;
                    break;
                }
            }
            
            if (!titleElement) {
                result.message = '未找到本周工作总结标题';
                return result;
            }
            
            // 滚动到标题位置（平滑滚动）
            titleElement.scrollIntoView({ behavior: 'smooth', block: 'center' });
            
            // 获取滚动后的位置
            setTimeout(() => {
                const rect = titleElement.getBoundingClientRect();
                result.titleRect = {
                    top: rect.top,
                    bottom: rect.bottom,
                    left: rect.left,
                    right: rect.right,
                    width: rect.width,
                    height: rect.height
                };
            }, 500);
            
            result.success = true;
            result.message = '已滚动到标题位置';
            return result;
        }""")
        
        if not scroll_result.get('success'):
            print(f"  [X] {scroll_result.get('message')}")
            return False
        
        print(f"  [OK] 已滚动到标题: {scroll_result['titleText']}")
        await asyncio.sleep(1)  # 等待滚动完成
        
        # ==================== 第二步：获取精确位置 ====================
        print("\n" + "-"*60)
        print("  [步骤2] 获取精确位置")
        print("-"*60)
        
        title_info = await page.evaluate("""() => {
            const result = {
                success: false,
                message: '',
                clickX: 0,
                clickY: 0,
                titleText: '',
                hasTextElement: false
            };
            
            // 重新获取标题（确保是最新位置）
            const headings = document.querySelectorAll('h1, h2, h3, .mp-heading');
            let titleElement = null;
            
            for (const h of headings) {
                if (h.textContent && h.textContent.includes('本周工作总结')) {
                    titleElement = h;
                    result.titleText = h.textContent;
                    break;
                }
            }
            
            if (!titleElement) {
                result.message = '未找到本周工作总结标题';
                return result;
            }
            
            // 高亮标题
            titleElement.style.border = '2px solid red';
            titleElement.style.backgroundColor = 'rgba(255, 200, 200, 0.3)';
            
            // 查找标题内的文字元素
            const textElements = titleElement.querySelectorAll('span[data-slate-string="true"]');
            let clickElement = null;
            
            if (textElements.length > 0) {
                // 使用最后一个文字元素
                clickElement = textElements[textElements.length - 1];
                result.hasTextElement = true;
                
                // 高亮文字元素
                clickElement.style.backgroundColor = 'rgba(255, 255, 0, 0.5)';
                clickElement.style.border = '1px solid orange';
            }
            
            // 获取点击位置
            if (clickElement) {
                const textRect = clickElement.getBoundingClientRect();
                result.clickX = textRect.right + 8;  // 文字右侧8像素
                result.clickY = textRect.top + textRect.height / 2;
            } else {
                const rect = titleElement.getBoundingClientRect();
                result.clickX = rect.right - 30;
                result.clickY = rect.top + rect.height / 2;
            }
            
            // 添加绿色标记
            const marker = document.createElement('div');
            marker.className = 'debug-marker';
            marker.style.position = 'fixed';
            marker.style.left = result.clickX + 'px';
            marker.style.top = result.clickY + 'px';
            marker.style.width = '12px';
            marker.style.height = '12px';
            marker.style.backgroundColor = '#00FF00';
            marker.style.border = '2px solid #008800';
            marker.style.borderRadius = '50%';
            marker.style.zIndex = '10000';
            marker.style.pointerEvents = 'none';
            marker.style.transform = 'translate(-50%, -50%)';
            marker.style.boxShadow = '0 0 10px #00FF00';
            document.body.appendChild(marker);
            
            result.success = true;
            result.message = '已获取精确位置';
            return result;
        }""")
        
        if not title_info.get('success'):
            print(f"  [X] {title_info.get('message')}")
            return False
        
        print(f"  [OK] 找到标题: {title_info['titleText']}")
        print(f"  找到内部文字元素: {title_info.get('hasTextElement', False)}")
        print(f"  点击坐标: ({title_info['clickX']}, {title_info['clickY']})")
        
        input("\n  请确认绿色圆点是否在'本周工作总结'文字后方，按Enter继续...")
        
        # ==================== 第三步：点击激活 ====================
        print("\n" + "-"*60)
        print("  [步骤3] 点击激活")
        print("-"*60)
        
        click_x = title_info['clickX']
        click_y = title_info['clickY']
        
        # 移动鼠标
        await page.mouse.move(click_x, click_y)
        await asyncio.sleep(0.2)
        
        # 单击
        print(f"    单击坐标: ({click_x}, {click_y})")
        await page.mouse.click(click_x, click_y)
        await asyncio.sleep(0.5)
        
        # 再次点击确保
        await page.mouse.click(click_x, click_y)
        await asyncio.sleep(0.3)
        
        # ==================== 第四步：验证光标 ====================
        print("\n" + "-"*60)
        print("  [步骤4] 验证光标")
        print("-"*60)
        
        # 输入测试字符
        await page.keyboard.type('测试', delay=50)
        await asyncio.sleep(0.3)
        
        # 检查是否输入成功
        has_test = await page.evaluate("""() => {
            const editor = document.querySelector('[data-slate-editor="true"]');
            if (!editor) return false;
            return editor.textContent.includes('测试');
        }""")
        
        if has_test:
            print("  [OK] 光标已激活")
            # 删除测试字符
            for _ in range(2):
                await page.keyboard.press('Backspace')
                await asyncio.sleep(0.1)
        else:
            print("  [⚠️] 光标未激活")
            print("  请手动点击绿色圆点位置，然后按Enter")
            input("    按Enter继续...")
        
        # ==================== 第五步：插入空行 ====================
        print("\n" + "-"*60)
        print("  [步骤5] 插入空行")
        print("-"*60)
        
        # 输入2个空格
        print("    输入空格...")
        for i in range(2):
            await page.keyboard.press('Space')
            await asyncio.sleep(0.1)
        
        # 按两次回车
        print("    按两次回车...")
        await page.keyboard.press('Enter')
        await asyncio.sleep(0.3)
        await page.keyboard.press('Enter')
        await asyncio.sleep(0.3)
        
        # 按上箭头回到第一行
        await page.keyboard.press('ArrowUp')
        await asyncio.sleep(0.2)
        
        # ==================== 第六步：输入内容 ====================
        print("\n" + "-"*60)
        print("  [步骤6] 输入内容")
        print("-"*60)
        
        print("    开始输入...")
        lines = clean_content.split('\n')
        for i, line in enumerate(lines):
            if line.strip():
                print(f"      第 {i+1} 行: {line[:30]}...")
                await page.keyboard.type(line, delay=30)
                if i < len(lines) - 1:
                    await page.keyboard.press('Enter')
                    await asyncio.sleep(0.2)
        
        await page.keyboard.press('Enter')
        await asyncio.sleep(0.3)
        
        print("    [OK] 输入完成")
        
        # ==================== 第七步：触发保存 ====================
        print("\n" + "-"*60)
        print("  [步骤7] 触发保存")
        print("-"*60)
        
        await page.mouse.click(10, 10)
        await asyncio.sleep(0.5)
        
        try:
            save_button = await page.get_by_role("button", name="保存").first
            if await save_button.count() > 0:
                await save_button.click()
                print("    [OK] 已点击保存")
        except:
            pass
        
        # 清除标记
        await page.evaluate("""() => {
            document.querySelectorAll('.debug-marker').forEach(el => el.remove());
            document.querySelectorAll('[style*="border"]').forEach(el => {
                el.style.border = '';
                el.style.backgroundColor = '';
            });
        }""")
        
        print("\n  [✅] 所有步骤完成")
        
        response = input("\n  请检查内容是否正确插入，输入 y 确认 (y/n): ").strip().lower()
        return response == 'y'
        
    except Exception as e:
        print(f"\n  [X] 填写过程出错: {e}")
        import traceback
        traceback.print_exc()
        return False


    """填写第二部分：OKR进度 - 基于div结构版"""
    
    print("\n" + "="*80)
    print("[填写] 第二部分：OKR进度（div结构版）")
    print("="*80)
    
    objectives = part2_data.get("objectives", [])
    if not objectives:
        print("  [X] 没有OKR数据")
        return False
    
    print(f"  共有 {len(objectives)} 个Objective需要填写")
    
    # 打印JSON数据结构预览
    print("\n  [JSON数据结构预览]")
    for obj_idx, obj in enumerate(objectives):
        o_title = obj.get('o_title', f'O{obj_idx+1}')
        print(f"    O{obj_idx+1}: {o_title[:50]}...")
        krs = obj.get("key_results", [])
        for kr_idx, kr in enumerate(krs):
            kr_title = kr.get('kr_title', f'KR{kr_idx+1}')
            weekly_work = kr.get('weekly_work', '')
            print(f"      KR{kr_idx+1}: {kr_title[:30]}...")
            if weekly_work:
                print(f"        本周工作: {weekly_work[:50]}...")
    
    try:
        # ==================== 第一步：定位OKR区域 ====================
        print("\n" + "-"*60)
        print("  [步骤1] 定位OKR区域")
        print("-"*60)
        
        # 清除可能存在的旧标记
        await page.evaluate("""() => {
            document.querySelectorAll('[style*="border"]').forEach(el => {
                el.style.border = '';
                el.style.backgroundColor = '';
            });
            document.querySelectorAll('.debug-marker').forEach(el => el.remove());
        }""")
        
        # 定位OKR区域
        okr_region = await page.evaluate("""() => {
            const result = {
                success: false,
                message: '',
                oElements: [],
                krElements: [],
                weeklyWorkElements: []
            };
            
            // 1. 找到OKR进度标题并滚动
            const headings = document.querySelectorAll('h1, h2, h3, .mp-heading');
            let okrTitle = null;
            for (const h of headings) {
                if (h.textContent && h.textContent.includes('OKR进度')) {
                    okrTitle = h;
                    h.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    h.style.border = '3px solid blue';
                    h.style.backgroundColor = 'rgba(0, 0, 255, 0.1)';
                    break;
                }
            }
            
            if (!okrTitle) {
                result.message = '未找到OKR进度标题';
                return result;
            }
            
            // 2. 找到OKR卡片区域（包含"卜安安的OKR"的区域）
            const allElements = document.querySelectorAll('div, section');
            let okrCard = null;
            
            for (const el of allElements) {
                if (el.textContent && el.textContent.includes('卜安安的OKR')) {
                    // 找到包含OKR卡片的容器
                    okrCard = el.closest('div[class*="card"], div[class*="container"], section') || el;
                    
                    // 高亮整个OKR区域
                    okrCard.style.border = '3px solid red';
                    okrCard.style.backgroundColor = 'rgba(255, 0, 0, 0.05)';
                    break;
                }
            }
            
            if (!okrCard) {
                result.message = '未找到OKR卡片';
                return result;
            }
            
            // 3. 在OKR卡片内查找所有O元素
            const oCandidates = okrCard.querySelectorAll('[class*="objective"], [class*="Objective"]');
            oCandidates.forEach((el, idx) => {
                if (el.textContent && el.textContent.match(/O[0-9]/)) {
                    result.oElements.push({
                        index: idx,
                        element: el,
                        text: el.textContent
                    });
                    
                    // 高亮O元素
                    el.style.border = '2px solid orange';
                    el.style.backgroundColor = 'rgba(255, 165, 0, 0.1)';
                    
                    // 添加O序号标记
                    const marker = document.createElement('div');
                    marker.className = 'debug-marker';
                    marker.style.position = 'absolute';
                    marker.style.left = '10px';
                    marker.style.top = el.getBoundingClientRect().top + 'px';
                    marker.style.backgroundColor = 'orange';
                    marker.style.color = 'black';
                    marker.style.padding = '2px 5px';
                    marker.style.zIndex = '10000';
                    marker.style.fontSize = '12px';
                    marker.textContent = `O${idx + 1}`;
                    document.body.appendChild(marker);
                }
            });
            
            // 4. 查找所有KR元素
            const krCandidates = okrCard.querySelectorAll('[class*="key-result"], [class*="KeyResult"]');
            krCandidates.forEach((el, idx) => {
                if (el.textContent && el.textContent.includes('KR')) {
                    result.krElements.push({
                        index: idx,
                        element: el,
                        text: el.textContent
                    });
                    
                    // 高亮KR元素
                    el.style.border = '2px solid purple';
                    el.style.backgroundColor = 'rgba(128, 0, 128, 0.1)';
                    
                    // 添加KR序号标记
                    const marker = document.createElement('div');
                    marker.className = 'debug-marker';
                    marker.style.position = 'absolute';
                    marker.style.left = '30px';
                    marker.style.top = el.getBoundingClientRect().top + 'px';
                    marker.style.backgroundColor = 'purple';
                    marker.style.color = 'white';
                    marker.style.padding = '2px 5px';
                    marker.style.zIndex = '10000';
                    marker.style.fontSize = '12px';
                    marker.textContent = `KR${idx + 1}`;
                    document.body.appendChild(marker);
                }
            });
            
            // 5. 查找所有"本周工作"输入区域
            const allSpans = okrCard.querySelectorAll('span');
            allSpans.forEach((span) => {
                if (span.textContent && span.textContent.includes('本周工作')) {
                    // 找到包含"本周工作"的容器
                    const container = span.closest('div[class*="row"], div[class*="cell"], tr, div[class*="container"]');
                    if (container) {
                        // 在容器内查找可编辑区域
                        const editableArea = container.querySelector('[contenteditable="true"], .mp-paragraph-wrapper, [class*="input"], [class*="editor"]');
                        
                        if (editableArea) {
                            result.weeklyWorkElements.push({
                                span: span,
                                container: container,
                                element: editableArea
                            });
                            
                            // 高亮可编辑区域
                            editableArea.style.border = '2px solid green';
                            editableArea.style.backgroundColor = 'rgba(0, 255, 0, 0.1)';
                        }
                    }
                }
            });
            
            result.success = true;
            result.message = '成功定位OKR区域';
            return result;
        }""")
        
        if not okr_region.get('success'):
            print(f"  [X] {okr_region.get('message')}")
            return False
        
        print(f"  [OK] {okr_region.get('message')}")
        print(f"  找到O元素: {len(okr_region.get('oElements', []))} 个")
        print(f"  找到KR元素: {len(okr_region.get('krElements', []))} 个")
        print(f"  找到本周工作输入区域: {len(okr_region.get('weeklyWorkElements', []))} 个")
        
        await asyncio.sleep(1)
        
        input("\n  请确认页面元素已正确高亮，按Enter继续填写...")
        
        # ==================== 第二步：填写每个KR的本周工作 ====================
        print("\n" + "-"*60)
        print("  [步骤2] 填写本周工作")
        print("-"*60)
        
        success_count = 0
        total_krs = sum(len(obj.get("key_results", [])) for obj in objectives)
        
        # 遍历JSON中的每个Objective
        input_index = 0
        for obj_idx, obj in enumerate(objectives):
            key_results = obj.get("key_results", [])
            
            print(f"\n  [Objective {obj_idx+1}] 处理 {len(key_results)} 个KR")
            
            for kr_idx, kr in enumerate(key_results):
                weekly_work = kr.get("weekly_work", "")
                
                if not weekly_work:
                    print(f"    [跳过] KR{kr_idx+1}: 无内容")
                    continue
                
                print(f"\n    [KR{kr_idx+1}] 准备填写...")
                print(f"      内容预览: {weekly_work[:60]}...")
                
                # 找到对应的输入区域
                if input_index < len(okr_region['weeklyWorkElements']):
                    input_info = okr_region['weeklyWorkElements'][input_index]
                    
                    # 获取输入区域的位置
                    try:
                        input_rect = await page.evaluate("""(element) => {
                            const rect = element.getBoundingClientRect();
                            return {
                                x: rect.left + 10,
                                y: rect.top + 5,
                                width: rect.width,
                                height: rect.height,
                                exists: rect.width > 0 && rect.height > 0,
                                visible: window.getComputedStyle(element).display !== 'none'
                            };
                        }""", input_info['element'])
                        
                        if not input_rect.get('exists') or not input_rect.get('visible'):
                            print(f"      [⚠️] 输入区域不可见")
                            input_index += 1
                            continue
                        
                        print(f"      输入区域位置: ({input_rect['x']:.0f}, {input_rect['y']:.0f})")
                        
                        # 滚动到输入区域
                        await input_info['element'].scroll_into_view_if_needed()
                        await asyncio.sleep(0.3)
                        
                        # 点击激活
                        await page.mouse.click(input_rect['x'], input_rect['y'])
                        await asyncio.sleep(0.2)
                        await page.mouse.dblclick(input_rect['x'], input_rect['y'])
                        await asyncio.sleep(0.3)
                        
                        # 清空现有内容
                        await page.keyboard.down('Control')
                        await page.keyboard.press('A')
                        await page.keyboard.up('Control')
                        await asyncio.sleep(0.2)
                        await page.keyboard.press('Delete')
                        await asyncio.sleep(0.2)
                        
                        # 输入新内容
                        print(f"      输入内容中...")
                        lines = weekly_work.split('\n')
                        for i, line in enumerate(lines):
                            if line.strip():
                                await page.keyboard.type(line, delay=20)
                                if i < len(lines) - 1:
                                    await page.keyboard.press('Enter')
                                    await asyncio.sleep(0.2)
                        
                        await page.keyboard.press('Enter')
                        await asyncio.sleep(0.2)
                        
                        print(f"      [OK] 填写完成")
                        success_count += 1
                    except Exception as e:
                        print(f"      [X] 填写失败: {e}")
                else:
                    print(f"      [⚠️] 未找到对应的输入区域")
                
                input_index += 1
        
        print(f"\n  [填写统计] {success_count}/{total_krs} 个KR填写成功")
        
        # ==================== 第三步：清除标记 ====================
        print("\n" + "-"*60)
        print("  [步骤3] 清除标记")
        print("-"*60)
        
        await page.evaluate("""() => {
            document.querySelectorAll('[style*="border"]').forEach(el => {
                el.style.border = '';
                el.style.backgroundColor = '';
            });
            document.querySelectorAll('.debug-marker').forEach(el => el.remove());
        }""")
        
        print("  [OK] 标记已清除")
        
        return success_count > 0
        
    except Exception as e:
        print(f"\n  [X] 填写过程出错: {e}")
        import traceback
        traceback.print_exc()
        return False
    """填写第二部分：OKR进度 - 基于表格结构版"""
    
    print("\n" + "="*80)
    print("[填写] 第二部分：OKR进度（表格结构版）")
    print("="*80)
    
    objectives = part2_data.get("objectives", [])
    if not objectives:
        print("  [X] 没有OKR数据")
        return False
    
    print(f"  共有 {len(objectives)} 个Objective需要填写")
    
    # 打印JSON数据结构预览
    print("\n  [JSON数据结构预览]")
    for obj_idx, obj in enumerate(objectives):
        o_title = obj.get('o_title', f'O{obj_idx+1}')
        print(f"    O{obj_idx+1}: {o_title[:50]}...")
        krs = obj.get("key_results", [])
        for kr_idx, kr in enumerate(krs):
            kr_title = kr.get('kr_title', f'KR{kr_idx+1}')
            weekly_work = kr.get('weekly_work', '')
            print(f"      KR{kr_idx+1}: {kr_title[:30]}...")
            if weekly_work:
                print(f"        本周工作: {weekly_work[:50]}...")
    
    try:
        # ==================== 第一步：定位OKR表格 ====================
        print("\n" + "-"*60)
        print("  [步骤1] 定位OKR表格")
        print("-"*60)
        
        # 清除可能存在的旧标记
        await page.evaluate("""() => {
            document.querySelectorAll('[style*="border"]').forEach(el => {
                el.style.border = '';
                el.style.backgroundColor = '';
            });
            document.querySelectorAll('.debug-marker').forEach(el => el.remove());
        }""")
        
        # 定位OKR表格
        table_info = await page.evaluate("""() => {
            const result = {
                success: false,
                message: '',
                tableFound: false,
                oRows: [],
                krRows: [],
                weeklyWorkRows: []
            };
            
            // 1. 找到OKR进度标题并滚动到视图
            const headings = document.querySelectorAll('h1, h2, h3, .mp-heading');
            for (const h of headings) {
                if (h.textContent && h.textContent.includes('OKR进度')) {
                    h.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    // 高亮标题
                    h.style.border = '3px solid blue';
                    h.style.backgroundColor = 'rgba(0, 0, 255, 0.1)';
                    break;
                }
            }
            
            // 2. 找到包含"卜安安的OKR"的表格
            const tables = document.querySelectorAll('table');
            for (const table of tables) {
                const tableText = table.textContent || '';
                if (tableText.includes('卜安安的OKR')) {
                    result.tableFound = true;
                    
                    // 高亮表格
                    table.style.border = '3px solid red';
                    table.style.backgroundColor = 'rgba(255, 0, 0, 0.05)';
                    
                    // 3. 分析表格行
                    const rows = table.querySelectorAll('tr');
                    let currentO = null;
                    let oIndex = 0;
                    let krIndex = 0;
                    
                    rows.forEach((row, rowIdx) => {
                        const rowText = row.textContent || '';
                        
                        // 识别O行（包含O1、O2等字样）- 使用正则表达式匹配 O后面跟数字
                        const oMatch = rowText.match(/O[0-9]+/);
                        if (oMatch && !rowText.includes('KR')) {
                            oIndex++;
                            const oInfo = {
                                index: oIndex,
                                rowIndex: rowIdx,
                                text: rowText,
                                element: row
                            };
                            result.oRows.push(oInfo);
                            currentO = oInfo;
                            
                            // 高亮O行
                            row.style.border = '2px solid orange';
                            row.style.backgroundColor = 'rgba(255, 165, 0, 0.1)';
                            
                            // 添加O序号标记
                            const oLabel = document.createElement('div');
                            oLabel.className = 'debug-marker';
                            oLabel.style.position = 'absolute';
                            oLabel.style.left = '10px';
                            oLabel.style.top = row.offsetTop + 'px';
                            oLabel.style.backgroundColor = 'orange';
                            oLabel.style.color = 'black';
                            oLabel.style.padding = '2px 5px';
                            oLabel.style.zIndex = '10000';
                            oLabel.style.fontSize = '12px';
                            oLabel.textContent = `O${oIndex}`;
                            document.body.appendChild(oLabel);
                        }
                        
                        // 识别KR行（包含KR字样）
                        if (rowText.includes('KR')) {
                            krIndex++;
                            const krInfo = {
                                index: krIndex,
                                rowIndex: rowIdx,
                                text: rowText,
                                element: row,
                                parentO: currentO ? currentO.index : null
                            };
                            result.krRows.push(krInfo);
                            
                            // 高亮KR行
                            row.style.border = '2px solid purple';
                            row.style.backgroundColor = 'rgba(128, 0, 128, 0.1)';
                            
                            // 添加KR序号标记
                            const krLabel = document.createElement('div');
                            krLabel.className = 'debug-marker';
                            krLabel.style.position = 'absolute';
                            krLabel.style.left = '30px';
                            krLabel.style.top = row.offsetTop + 'px';
                            krLabel.style.backgroundColor = 'purple';
                            krLabel.style.color = 'white';
                            krLabel.style.padding = '2px 5px';
                            krLabel.style.zIndex = '10000';
                            krLabel.style.fontSize = '12px';
                            krLabel.textContent = `KR${krIndex}`;
                            document.body.appendChild(krLabel);
                            
                            // 4. 查找该KR行下的"本周工作"输入区域
                            const cells = row.querySelectorAll('td');
                            for (let i = 0; i < cells.length; i++) {
                                const cellText = cells[i].textContent || '';
                                if (cellText.includes('本周工作')) {
                                    // 下一个单元格通常是输入区域
                                    const inputCell = cells[i + 1];
                                    if (inputCell) {
                                        // 查找可编辑区域
                                        const inputArea = inputCell.querySelector('[contenteditable="true"], .mp-paragraph-wrapper, [class*="editor"]');
                                        
                                        if (inputArea) {
                                            result.weeklyWorkRows.push({
                                                krIndex: krIndex,
                                                oIndex: currentO ? currentO.index : null,
                                                cell: inputCell,
                                                input: inputArea
                                            });
                                            
                                            // 高亮输入区域
                                            inputArea.style.border = '2px solid green';
                                            inputArea.style.backgroundColor = 'rgba(0, 255, 0, 0.1)';
                                        }
                                    }
                                    break;
                                }
                            }
                        }
                    });
                    
                    break;
                }
            }
            
            result.success = result.tableFound;
            result.message = result.tableFound ? '找到OKR表格' : '未找到OKR表格';
            return result;
        }""")
        
        if not table_info.get('success'):
            print(f"  [X] {table_info.get('message')}")
            return False
        
        print(f"  [OK] {table_info.get('message')}")
        print(f"  找到O行: {len(table_info.get('oRows', []))} 个")
        print(f"  找到KR行: {len(table_info.get('krRows', []))} 个")
        print(f"  找到本周工作输入区域: {len(table_info.get('weeklyWorkRows', []))} 个")
        
        await asyncio.sleep(1)
        
        input("\n  请确认页面元素已正确高亮，按Enter继续填写...")
        
        # ==================== 第二步：填写每个KR的本周工作 ====================
        print("\n" + "-"*60)
        print("  [步骤2] 填写本周工作")
        print("-"*60)
        
        success_count = 0
        total_krs = sum(len(obj.get("key_results", [])) for obj in objectives)
        
        # 遍历JSON中的每个Objective
        json_kr_index = 0
        for obj_idx, obj in enumerate(objectives):
            key_results = obj.get("key_results", [])
            
            print(f"\n  [Objective {obj_idx+1}] 处理 {len(key_results)} 个KR")
            
            for kr_idx, kr in enumerate(key_results):
                weekly_work = kr.get("weekly_work", "")
                
                if not weekly_work:
                    print(f"    [跳过] KR{kr_idx+1}: 无内容")
                    continue
                
                print(f"\n    [KR{kr_idx+1}] 准备填写...")
                print(f"      内容预览: {weekly_work[:60]}...")
                
                # 找到对应的输入区域
                if json_kr_index < len(table_info['weeklyWorkRows']):
                    input_info = table_info['weeklyWorkRows'][json_kr_index]
                    
                    # 获取输入区域的位置
                    input_rect = await page.evaluate("""(element) => {
                        const rect = element.getBoundingClientRect();
                        return {
                            x: rect.left + 10,
                            y: rect.top + 5,
                            width: rect.width,
                            height: rect.height,
                            exists: rect.width > 0 && rect.height > 0
                        };
                    }""", input_info['input'])
                    
                    if not input_rect.get('exists'):
                        print(f"      [⚠️] 输入区域不可见")
                        json_kr_index += 1
                        continue
                    
                    print(f"      输入区域位置: ({input_rect['x']:.0f}, {input_rect['y']:.0f})")
                    
                    # 滚动到输入区域
                    await input_info['input'].scroll_into_view_if_needed()
                    await asyncio.sleep(0.3)
                    
                    # 多次点击激活
                    click_x = input_rect['x']
                    click_y = input_rect['y']
                    
                    # 单击
                    await page.mouse.click(click_x, click_y)
                    await asyncio.sleep(0.2)
                    
                    # 双击确保激活
                    await page.mouse.dblclick(click_x, click_y)
                    await asyncio.sleep(0.3)
                    
                    # 再次单击
                    await page.mouse.click(click_x, click_y)
                    await asyncio.sleep(0.2)
                    
                    # 清空现有内容
                    await page.keyboard.down('Control')
                    await page.keyboard.press('A')
                    await page.keyboard.up('Control')
                    await asyncio.sleep(0.2)
                    await page.keyboard.press('Delete')
                    await asyncio.sleep(0.2)
                    
                    # 输入新内容
                    print(f"      输入内容中...")
                    
                    # 分行输入
                    lines = weekly_work.split('\n')
                    for i, line in enumerate(lines):
                        if line.strip():
                            await page.keyboard.type(line, delay=20)
                            if i < len(lines) - 1:
                                await page.keyboard.press('Enter')
                                await asyncio.sleep(0.2)
                    
                    await page.keyboard.press('Enter')
                    await asyncio.sleep(0.2)
                    
                    print(f"      [OK] 填写完成")
                    success_count += 1
                else:
                    print(f"      [⚠️] 未找到对应的输入区域")
                
                json_kr_index += 1
        
        print(f"\n  [填写统计] {success_count}/{total_krs} 个KR填写成功")
        
        # ==================== 第三步：清除标记 ====================
        print("\n" + "-"*60)
        print("  [步骤3] 清除标记")
        print("-"*60)
        
        await page.evaluate("""() => {
            document.querySelectorAll('[style*="border"]').forEach(el => {
                el.style.border = '';
                el.style.backgroundColor = '';
            });
            document.querySelectorAll('.debug-marker').forEach(el => el.remove());
        }""")
        
        print("  [OK] 标记已清除")
        
        return success_count > 0
        
    except Exception as e:
        print(f"\n  [X] 填写过程出错: {e}")
        import traceback
        traceback.print_exc()
        return False
    """填写第二部分：OKR进度 - 基于表格结构版"""
    
    print("\n" + "="*80)
    print("[填写] 第二部分：OKR进度（表格结构版）")
    print("="*80)
    
    objectives = part2_data.get("objectives", [])
    if not objectives:
        print("  [X] 没有OKR数据")
        return False
    
    print(f"  共有 {len(objectives)} 个Objective需要填写")
    
    # 打印JSON数据结构预览
    print("\n  [JSON数据结构预览]")
    for obj_idx, obj in enumerate(objectives):
        o_title = obj.get('o_title', f'O{obj_idx+1}')
        print(f"    O{obj_idx+1}: {o_title[:50]}...")
        krs = obj.get("key_results", [])
        for kr_idx, kr in enumerate(krs):
            kr_title = kr.get('kr_title', f'KR{kr_idx+1}')
            weekly_work = kr.get('weekly_work', '')
            print(f"      KR{kr_idx+1}: {kr_title[:30]}...")
            print(f"        本周工作: {weekly_work[:50]}...")
    
    try:
        # ==================== 第一步：定位OKR表格 ====================
        print("\n" + "-"*60)
        print("  [步骤1] 定位OKR表格")
        print("-"*60)
        
        # 清除可能存在的旧标记
        await page.evaluate("""() => {
            document.querySelectorAll('[style*="border"]').forEach(el => {
                el.style.border = '';
                el.style.backgroundColor = '';
            });
            document.querySelectorAll('.debug-marker').forEach(el => el.remove());
        }""")
        
        # 定位OKR表格
        table_info = await page.evaluate("""() => {
            const result = {
                success: false,
                message: '',
                tableFound: false,
                oRows: [],
                krRows: [],
                weeklyWorkRows: []
            };
            
            // 1. 找到OKR进度标题并滚动到视图
            const headings = document.querySelectorAll('h1, h2, h3, .mp-heading');
            for (const h of headings) {
                if (h.textContent && h.textContent.includes('OKR进度')) {
                    h.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    // 高亮标题
                    h.style.border = '3px solid blue';
                    h.style.backgroundColor = 'rgba(0, 0, 255, 0.1)';
                    break;
                }
            }
            
            // 2. 找到包含"卜安安的OKR"的表格
            const tables = document.querySelectorAll('table');
            for (const table of tables) {
                const tableText = table.textContent || '';
                if (tableText.includes('卜安安的OKR')) {
                    result.tableFound = true;
                    
                    // 高亮表格
                    table.style.border = '3px solid red';
                    table.style.backgroundColor = 'rgba(255, 0, 0, 0.05)';
                    
                    // 3. 分析表格行
                    const rows = table.querySelectorAll('tr');
                    let currentO = null;
                    
                    rows.forEach((row, index) => {
                        const rowText = row.textContent || '';
                        
                        // 识别O行（包含O1、O2等字样）
                        if (rowText.match(/O\d+/) && !rowText.includes('KR')) {
                            const oInfo = {
                                index: index,
                                text: rowText,
                                element: row
                            };
                            result.oRows.push(oInfo);
                            currentO = oInfo;
                            
                            // 高亮O行
                            row.style.border = '2px solid orange';
                            row.style.backgroundColor = 'rgba(255, 165, 0, 0.1)';
                        }
                        
                        // 识别KR行（包含KR字样）
                        if (rowText.includes('KR')) {
                            const krInfo = {
                                index: index,
                                text: rowText,
                                element: row,
                                parentO: currentO
                            };
                            result.krRows.push(krInfo);
                            
                            // 高亮KR行
                            row.style.border = '2px solid purple';
                            row.style.backgroundColor = 'rgba(128, 0, 128, 0.1)';
                            
                            // 4. 查找该KR行下的"本周工作"输入区域
                            // 查找包含"本周工作"的单元格
                            const cells = row.querySelectorAll('td');
                            for (let i = 0; i < cells.length; i++) {
                                if (cells[i].textContent && cells[i].textContent.includes('本周工作')) {
                                    // 下一个单元格或当前单元格后的输入区域
                                    const inputCell = cells[i + 1] || cells[i];
                                    const inputArea = inputCell.querySelector('div[contenteditable="true"], [class*="input"], [class*="editor"]');
                                    
                                    if (inputArea) {
                                        result.weeklyWorkRows.push({
                                            krIndex: index,
                                            cell: inputCell,
                                            input: inputArea
                                        });
                                        
                                        // 高亮输入区域
                                        inputArea.style.border = '2px solid green';
                                        inputArea.style.backgroundColor = 'rgba(0, 255, 0, 0.1)';
                                    }
                                    break;
                                }
                            }
                        }
                    });
                    
                    break;
                }
            }
            
            result.success = result.tableFound;
            result.message = result.tableFound ? '找到OKR表格' : '未找到OKR表格';
            return result;
        }""")
        
        if not table_info.get('success'):
            print(f"  [X] {table_info.get('message')}")
            return False
        
        print(f"  [OK] {table_info.get('message')}")
        print(f"  找到O行: {len(table_info.get('oRows', []))} 个")
        print(f"  找到KR行: {len(table_info.get('krRows', []))} 个")
        print(f"  找到本周工作输入区域: {len(table_info.get('weeklyWorkRows', []))} 个")
        
        await asyncio.sleep(1)  # 等待滚动完成
        
        input("\n  请确认：1.OKR进度标题蓝色 2.表格红色边框 3.O行橙色 4.KR行紫色 5.输入区域绿色，按Enter继续...")
        
        # ==================== 第二步：匹配JSON和表格行 ====================
        print("\n" + "-"*60)
        print("  [步骤2] 匹配JSON和表格行")
        print("-"*60)
        
        # 这里需要根据实际表格结构进行匹配
        # 暂时先打印信息
        print(f"\n  JSON中Objective数量: {len(objectives)}")
        print(f"  表格中KR行数量: {len(table_info['krRows'])}")
        
        if len(objectives) != len(table_info['oRows']):
            print(f"  [⚠️] JSON中O数量({len(objectives)})与表格中O数量({len(table_info['oRows'])})不匹配")
        
        # ==================== 第三步：填写每个KR的本周工作 ====================
        print("\n" + "-"*60)
        print("  [步骤3] 填写本周工作")
        print("-"*60)
        
        success_count = 0
        total_krs = sum(len(obj.get("key_results", [])) for obj in objectives)
        
        # 遍历每个Objective
        kr_counter = 0
        for obj_idx, obj in enumerate(objectives):
            o_title = obj.get("o_title", f"O{obj_idx+1}")
            key_results = obj.get("key_results", [])
            
            print(f"\n  [Objective {obj_idx+1}] {o_title[:50]}...")
            
            for kr_idx, kr in enumerate(key_results):
                weekly_work = kr.get("weekly_work", "")
                
                if not weekly_work:
                    print(f"    [跳过] KR{kr_idx+1}: 无内容")
                    continue
                
                print(f"\n    [KR{kr_idx+1}] 待填写内容: {weekly_work[:50]}...")
                
                # 找到对应的输入区域
                if kr_counter < len(table_info['weeklyWorkRows']):
                    input_info = table_info['weeklyWorkRows'][kr_counter]
                    
                    # 获取输入区域的位置
                    input_rect = await page.evaluate("""(element) => {
                        const rect = element.getBoundingClientRect();
                        return {
                            x: rect.left + 10,
                            y: rect.top + 5,
                            width: rect.width,
                            height: rect.height
                        };
                    }""", input_info['input'])
                    
                    print(f"      输入区域位置: ({input_rect['x']}, {input_rect['y']})")
                    
                    # 点击激活输入区域
                    await page.mouse.click(input_rect['x'], input_rect['y'])
                    await asyncio.sleep(0.3)
                    
                    # 双击确保激活
                    await page.mouse.dblclick(input_rect['x'], input_rect['y'])
                    await asyncio.sleep(0.3)
                    
                    # 清空现有内容（如果有）
                    await page.keyboard.down('Control')
                    await page.keyboard.press('A')
                    await page.keyboard.up('Control')
                    await asyncio.sleep(0.2)
                    await page.keyboard.press('Delete')
                    await asyncio.sleep(0.2)
                    
                    # 输入内容
                    print(f"      输入内容...")
                    await page.keyboard.type(weekly_work, delay=30)
                    await asyncio.sleep(0.3)
                    
                    # 按回车完成
                    await page.keyboard.press('Enter')
                    await asyncio.sleep(0.2)
                    
                    print(f"    [OK] KR{kr_idx+1} 填写完成")
                    success_count += 1
                else:
                    print(f"    [⚠️] 未找到对应的输入区域")
                
                kr_counter += 1
        
        print(f"\n  [填写统计] {success_count}/{total_krs} 个KR填写成功")
        
        # ==================== 第四步：清除标记 ====================
        print("\n" + "-"*60)
        print("  [步骤4] 清除标记")
        print("-"*60)
        
        await page.evaluate("""() => {
            document.querySelectorAll('[style*="border"]').forEach(el => {
                el.style.border = '';
                el.style.backgroundColor = '';
            });
            document.querySelectorAll('.debug-marker').forEach(el => el.remove());
        }""")
        
        print("  [OK] 标记已清除")
        
        return success_count > 0
        
    except Exception as e:
        print(f"\n  [X] 填写过程出错: {e}")
        import traceback
        traceback.print_exc()
        return False
    """填写第二部分：OKR进度 - 带调试信息的强化版"""
    
    print("\n" + "="*80)
    print("[填写] 第二部分：OKR进度（强化调试版）")
    print("="*80)
    
    objectives = part2_data.get("objectives", [])
    if not objectives:
        print("  [X] 没有OKR数据")
        return False
    
    print(f"  共有 {len(objectives)} 个Objective需要填写")
    
    # 打印JSON数据结构预览
    print("\n  [JSON数据结构预览]")
    for obj_idx, obj in enumerate(objectives):
        print(f"    Objective {obj_idx + 1}: {obj.get('o_title', '无标题')[:50]}...")
        krs = obj.get("key_results", [])
        print(f"      包含 {len(krs)} 个KR:")
        for kr_idx, kr in enumerate(krs):
            kr_title = kr.get('kr_title', '无标题')
            weekly_work = kr.get('weekly_work', '')
            print(f"        KR {kr_idx + 1}: {kr_title[:30]}...")
            print(f"          本周工作长度: {len(weekly_work)} 字符")
            if weekly_work:
                print(f"          预览: {weekly_work[:50]}...")
    
    try:
        # ==================== 第一步：定位OKR区域 ====================
        print("\n" + "-"*60)
        print("  [步骤1] 定位OKR区域")
        print("-"*60)
        
        # 清除可能存在的旧标记
        await page.evaluate("""() => {
            document.querySelectorAll('[style*="border"]').forEach(el => {
                el.style.border = '';
                el.style.backgroundColor = '';
            });
            document.querySelectorAll('.debug-marker').forEach(el => el.remove());
        }""")
        
        # 找到OKR进度标题并高亮
        okr_info = await page.evaluate("""() => {
            const result = {
                success: false,
                message: '',
                titleRect: null,
                hasOkrBlock: false
            };
            
            // 查找OKR进度标题
            const headings = document.querySelectorAll('h1, h2, h3, .mp-heading');
            let titleElement = null;
            
            for (const h of headings) {
                if (h.textContent && h.textContent.includes('OKR进度')) {
                    titleElement = h;
                    result.message = '找到OKR进度标题';
                    break;
                }
            }
            
            if (!titleElement) {
                result.message = '未找到OKR进度标题';
                return result;
            }
            
            // 高亮标题
            titleElement.style.border = '3px solid blue';
            titleElement.style.backgroundColor = 'rgba(0, 0, 255, 0.1)';
            
            // 获取标题位置
            const rect = titleElement.getBoundingClientRect();
            result.titleRect = {
                top: rect.top,
                bottom: rect.bottom,
                left: rect.left,
                right: rect.right,
                width: rect.width,
                height: rect.height
            };
            
            // 滚动到标题位置
            titleElement.scrollIntoView({ behavior: 'smooth', block: 'center' });
            
            // 查找OKR卡片区域
            const okrCards = document.querySelectorAll('.mp-okr-card, [class*="okr"], [class*="OKR"]');
            result.hasOkrBlock = okrCards.length > 0;
            
            result.success = true;
            return result;
        }""")
        
        print(f"  [OK] {okr_info.get('message', '已定位')}")
        print(f"  找到OKR卡片: {okr_info.get('hasOkrBlock', False)}")
        
        await asyncio.sleep(1)  # 等待滚动完成
        
        input("\n  请确认OKR进度标题已高亮为蓝色，按Enter继续...")
        
        # ==================== 第二步：分析OKR结构 ====================
        print("\n" + "-"*60)
        print("  [步骤2] 分析OKR结构")
        print("-"*60)
        
        structure_info = await page.evaluate("""() => {
            const result = {
                objectiveCount: 0,
                krCount: 0,
                objectives: [],
                weeklyWorkInputs: []
            };
            
            // 查找所有Objective容器
            const objectiveContainers = document.querySelectorAll('[class*="objective"], [class*="Objective"]');
            result.objectiveCount = objectiveContainers.length;
            
            // 查找所有KR容器
            const krContainers = document.querySelectorAll('[class*="key-result"], [class*="KeyResult"]');
            result.krCount = krContainers.length;
            
            // 收集Objective信息
            objectiveContainers.forEach((obj, idx) => {
                const title = obj.querySelector('[class*="title"]')?.textContent || `Objective ${idx + 1}`;
                result.objectives.push({
                    index: idx,
                    title: title,
                    element: obj
                });
                
                // 高亮Objective
                obj.style.border = '2px solid orange';
                obj.style.backgroundColor = 'rgba(255, 165, 0, 0.1)';
            });
            
            // 查找所有"本周工作"输入区域
            const weeklyWorkLabels = document.querySelectorAll('span:contains("本周工作")');
            weeklyWorkLabels.forEach((label, idx) => {
                // 查找对应的输入区域
                let inputArea = label.closest('tr') || label.closest('div[class*="kr"]') || label.parentElement;
                if (inputArea) {
                    result.weeklyWorkInputs.push({
                        index: idx,
                        label: label.textContent,
                        area: inputArea
                    });
                    
                    // 高亮输入区域
                    inputArea.style.border = '2px dashed green';
                }
            });
            
            return result;
        }""")
        
        print(f"\n  [结构分析]")
        print(f"    找到Objective容器: {structure_info['objectiveCount']} 个")
        print(f"    找到KR容器: {structure_info['krCount']} 个")
        print(f"    找到'本周工作'输入区域: {len(structure_info['weeklyWorkInputs'])} 个")
        
        input("\n  请确认页面元素已高亮，按Enter继续...")
        
        # ==================== 第三步：遍历填写每个KR ====================
        print("\n" + "-"*60)
        print("  [步骤3] 开始填写KR本周工作")
        print("-"*60)
        
        success_count = 0
        total_krs = sum(len(obj.get("key_results", [])) for obj in objectives)
        
        for obj_idx, obj in enumerate(objectives):
            o_title = obj.get("o_title", f"O{obj_idx+1}")
            key_results = obj.get("key_results", [])
            
            print(f"\n  [Objective {obj_idx+1}] {o_title[:50]}...")
            
            for kr_idx, kr in enumerate(key_results):
                kr_title = kr.get("kr_title", f"KR{kr_idx+1}")
                weekly_work = kr.get("weekly_work", "")
                
                if not weekly_work:
                    print(f"    [跳过] KR{kr_idx+1}: 无内容")
                    continue
                
                print(f"\n    [KR{kr_idx+1}] {kr_title[:40]}...")
                print(f"      待填写内容长度: {len(weekly_work)} 字符")
                print(f"      预览: {weekly_work[:60]}...")
                
                # 这里需要根据实际页面结构实现点击和输入
                # 暂时先用占位
                print(f"      [待实现] 需要定位到对应输入区域")
                
                # 模拟成功计数（实际应该根据填写结果判断）
                success_count += 1
        
        print(f"\n  [填写统计] {success_count}/{total_krs} 个KR待填写")
        
        # ==================== 第四步：暂不实际填写，仅调试 ====================
        print("\n" + "-"*60)
        print("  [步骤4] 调试完成")
        print("-"*60)
        
        print("\n  [⚠️] 当前版本仅用于调试，未实际填写内容")
        print("  请根据页面高亮情况分析结构，后续再实现实际填写逻辑")
        
        # 清除标记
        await page.evaluate("""() => {
            document.querySelectorAll('[style*="border"]').forEach(el => {
                el.style.border = '';
                el.style.backgroundColor = '';
            });
            document.querySelectorAll('.debug-marker').forEach(el => el.remove());
        }""")
        
        response = input("\n  调试完成，是否继续执行后续部分？(y/n): ").strip().lower()
        return response == 'y'
        
    except Exception as e:
        print(f"\n  [X] 填写过程出错: {e}")
        import traceback
        traceback.print_exc()
        return False
    """填写第二部分：OKR进度 - 精准匹配"""
    print("\n[填写] 第二部分：OKR进度")

    objectives = part2_data.get("objectives", [])
    if not objectives:
        print("  [X] 没有OKR数据")
        return False

    print(f"  共有 {len(objectives)} 个Objective需要填写")

    # 策略1: 找到包含 "OKR进度" 或 "卜安安的OKR" 的表格区域
    okr_table = None

    try:
        # 先尝试找到 "OKR进度" 标题
        print("  [定位] 尝试找到 'OKR进度' 区域...")
        okr_heading = page.get_by_text("OKR进度", exact=False).first

        if await okr_heading.count() > 0 and await okr_heading.is_visible(timeout=3000):
            print("    [OK] 找到 'OKR进度' 标题")
            # 向上查找包含表格的容器
            okr_section = okr_heading.locator('xpath=ancestor::div[contains(@class, "section") or contains(@class, "okr") or contains(@class, "card")]').first
            if await okr_section.count() == 0:
                okr_section = okr_heading.locator('xpath=ancestor::div[3]').first

            if await okr_section.count() > 0:
                # 在区域内查找表格
                okr_table = okr_section.locator('table').first
                if await okr_table.count() > 0:
                    print("    [OK] 找到OKR表格")

        # 如果没找到，尝试直接找包含 "卜安安的OKR" 的表格
        if not okr_table or await okr_table.count() == 0:
            print("  [定位] 尝试找到 '卜安安的OKR' 表格...")

            # 查找包含 "卜安安的OKR" 的表格
            tables = await page.locator('table').all()
            for table in tables:
                table_text = await table.text_content()
                if "卜安安的OKR" in table_text or "OKR" in table_text:
                    okr_table = table
                    print("    [OK] 找到包含 '卜安安的OKR' 的表格")
                    break

        if not okr_table or await okr_table.count() == 0:
            print("  [X] 未找到OKR表格")
            return False

        # 获取表格的所有行
        rows = await okr_table.locator('tr').all()
        print(f"    表格共有 {len(rows)} 行")

        success_count = 0

        # 遍历每个 Objective
        for obj_idx, obj in enumerate(objectives):
            o_title = obj.get("o_title", "")
            key_results = obj.get("key_results", [])

            print(f"\n  处理 Objective {obj_idx+1}: {o_title[:50]}...")

            # 在表格行中查找匹配的 O
            o_row_idx = None
            for idx, row in enumerate(rows):
                row_text = await row.text_content()
                if o_title in row_text:
                    o_row_idx = idx
                    print(f"    [OK] 找到O所在行: 第{idx+1}行")
                    break

            if o_row_idx is None:
                print(f"    [!] 未找到O '{o_title}' 所在行，尝试模糊匹配...")
                # 尝试模糊匹配
                for idx, row in enumerate(rows):
                    row_text = await row.text_content()
                    # 提取O的简短标题进行匹配
                    short_title = o_title[:20] if len(o_title) > 20 else o_title
                    if short_title in row_text:
                        o_row_idx = idx
                        print(f"    [OK] 模糊匹配到O所在行: 第{idx+1}行")
                        break

            # 处理该O下的所有KR
            for kr_idx, kr in enumerate(key_results):
                kr_title = kr.get("kr_title", "")
                weekly_work = kr.get("weekly_work", "")

                if not weekly_work:
                    print(f"    [跳过] KR {kr_idx+1}: 无内容")
                    continue

                print(f"\n    处理 KR {kr_idx+1}: {kr_title[:40]}...")

                # 在表格中查找该KR对应的行
                kr_row_idx = None
                search_start = o_row_idx if o_row_idx is not None else 0

                for idx in range(search_start, len(rows)):
                    row_text = await rows[idx].text_content()
                    if kr_title in row_text:
                        kr_row_idx = idx
                        print(f"      [OK] 找到KR所在行: 第{idx+1}行")
                        break

                if kr_row_idx is None:
                    print(f"      [!] 未找到KR '{kr_title}' 所在行")
                    continue

                # 在该行查找 "本周工作" 输入框
                kr_row = rows[kr_row_idx]

                # 策略：查找包含 "本周工作" 文本的单元格，然后找到其输入框
                filled = False

                try:
                    # 方法1: 在该行内查找 "本周工作" 标签
                    weekly_work_label = kr_row.locator('text=本周工作').first

                    if await weekly_work_label.count() > 0:
                        print(f"      [OK] 找到 '本周工作' 标签")

                        # 查找该标签下的输入框
                        # 先向上找单元格，再找输入框
                        cell = weekly_work_label.locator('xpath=ancestor::td | ancestor::div[contains(@class, "cell")]').first

                        if await cell.count() > 0:
                            # 在单元格内查找输入框
                            input_box = cell.locator('textarea, input[type="text"], [contenteditable="true"]').first

                            if await input_box.count() > 0 and await input_box.is_visible():
                                await input_box.fill(weekly_work)
                                print(f"      [OK] 已填写 '本周工作' ({len(weekly_work)}字符)")
                                filled = True
                                success_count += 1
                            else:
                                print(f"      [!] 未找到输入框")
                        else:
                            print(f"      [!] 未找到单元格容器")
                    else:
                        # 方法2: 直接在该行查找所有输入框，选择第一个可见的
                        print(f"      [尝试] 直接查找输入框...")
                        inputs = await kr_row.locator('textarea:visible, input[type="text"]:visible').all()

                        if inputs:
                            await inputs[0].fill(weekly_work)
                            print(f"      [OK] 已填写到第一个可见输入框 ({len(weekly_work)}字符)")
                            filled = True
                            success_count += 1
                        else:
                            print(f"      [!] 未找到可见输入框")

                except Exception as e:
                    print(f"      [X] 填写失败: {e}")
                    continue

                await asyncio.sleep(0.5)  # 填写后稍等

        total_krs = sum(len(o.get("key_results", [])) for o in objectives)
        print(f"\n  填写统计: {success_count}/{total_krs} 个KR成功")
        return success_count > 0

    except Exception as e:
        print(f"  [X] OKR填写过程出错: {e}")
        import traceback
        traceback.print_exc()
        return False

async def fill_okr_section(page, part2_data):
    """填写第二部分：OKR进度 - 精确表格定位版（过滤负坐标）"""
    
    print("\n" + "="*80)
    print("[填写] 第二部分：OKR进度（精确表格定位版）")
    print("="*80)
    
    objectives = part2_data.get("objectives", [])
    if not objectives:
        print("  [X] 没有OKR数据")
        return False
    
    total_krs = sum(len(obj.get("key_results", [])) for obj in objectives)
    print(f"  共有 {len(objectives)} 个Objective, {total_krs} 个KR需要填写")
    
    try:
        # ==================== 第一步：定位OKR表格 ====================
        print("\n" + "-"*60)
        print("  [步骤1] 定位OKR表格")
        print("-"*60)
        
        # 清除可能存在的旧标记
        await page.evaluate("""() => {
            document.querySelectorAll('[style*="border"]').forEach(el => {
                el.style.border = '';
                el.style.backgroundColor = '';
            });
            document.querySelectorAll('.debug-marker').forEach(el => el.remove());
        }""")
        
        # 定位OKR表格并高亮
        table_info = await page.evaluate("""() => {
            const result = {
                success: false,
                message: '',
                tableFound: false,
                inputCount: 0,
                inputElements: []
            };
            
            // 1. 找到OKR进度标题
            const headings = document.querySelectorAll('h1, h2, h3, .mp-heading');
            let okrTitle = null;
            for (const h of headings) {
                if (h.textContent && h.textContent.includes('OKR进度')) {
                    okrTitle = h;
                    h.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    h.style.border = '3px solid blue';
                    h.style.backgroundColor = 'rgba(0, 0, 255, 0.1)';
                    
                    // 添加标题标记
                    const titleMarker = document.createElement('div');
                    titleMarker.className = 'debug-marker';
                    titleMarker.style.position = 'fixed';
                    titleMarker.style.left = '10px';
                    titleMarker.style.top = h.getBoundingClientRect().top + 'px';
                    titleMarker.style.backgroundColor = 'blue';
                    titleMarker.style.color = 'white';
                    titleMarker.style.padding = '4px 8px';
                    titleMarker.style.borderRadius = '4px';
                    titleMarker.style.zIndex = '10000';
                    titleMarker.style.fontSize = '14px';
                    titleMarker.textContent = 'OKR进度标题';
                    document.body.appendChild(titleMarker);
                    break;
                }
            }
            
            if (!okrTitle) {
                result.message = '未找到OKR进度标题';
                return result;
            }
            
            // 2. 找到包含"卜安安的OKR"的表格容器
            let tableContainer = okrTitle;
            for (let i = 0; i < 10; i++) {
                tableContainer = tableContainer.parentElement;
                if (!tableContainer) break;
                
                if (tableContainer.textContent && 
                    tableContainer.textContent.includes('卜安安的OKR') &&
                    (tableContainer.querySelector('table') || 
                     tableContainer.querySelector('[class*="okr-card"]') ||
                     tableContainer.querySelector('[class*="OKR"]'))) {
                    result.tableFound = true;
                    break;
                }
            }
            
            if (!result.tableFound) {
                const allElements = document.querySelectorAll('div, section');
                for (const el of allElements) {
                    if (el.textContent && el.textContent.includes('卜安安的OKR')) {
                        tableContainer = el.closest('div[class*="container"], div[class*="card"], section') || el;
                        result.tableFound = true;
                        break;
                    }
                }
            }
            
            if (!result.tableFound) {
                result.message = '未找到OKR表格';
                return result;
            }
            
            tableContainer.style.border = '3px solid red';
            tableContainer.style.backgroundColor = 'rgba(255, 0, 0, 0.05)';
            
            const tableRect = tableContainer.getBoundingClientRect();
            const tableMarker = document.createElement('div');
            tableMarker.className = 'debug-marker';
            tableMarker.style.position = 'fixed';
            tableMarker.style.left = tableRect.left + 'px';
            tableMarker.style.top = (tableRect.top - 25) + 'px';
            tableMarker.style.backgroundColor = 'red';
            tableMarker.style.color = 'white';
            tableMarker.style.padding = '4px 8px';
            tableMarker.style.borderRadius = '4px';
            tableMarker.style.zIndex = '10000';
            tableMarker.style.fontSize = '14px';
            tableMarker.textContent = 'OKR表格区域';
            document.body.appendChild(tableMarker);
            
            // 3. 在表格容器内查找"本周工作"输入区域
            const allSpans = tableContainer.querySelectorAll('span');
            const processedInputs = new Set();
            
            allSpans.forEach((span) => {
                if (span.textContent && span.textContent.includes('本周工作')) {
                    let container = span;
                    for (let i = 0; i < 5; i++) {
                        container = container.parentElement;
                        if (!container) break;
                        
                        const editable = container.querySelector('[contenteditable="true"], .mp-paragraph-wrapper');
                        if (editable && !processedInputs.has(editable)) {
                            processedInputs.add(editable);
                            
                            const rect = editable.getBoundingClientRect();
                            
                            // 【关键修改】只保存坐标为正的输入区域
                            if (rect.top > 0 && rect.left > 0) {
                                result.inputElements.push({
                                    element: editable,
                                    index: result.inputCount,
                                    x: rect.left + 10,
                                    y: rect.top + 5,
                                    rawY: rect.top
                                });
                                
                                editable.style.border = '3px solid green';
                                editable.style.backgroundColor = 'rgba(0, 255, 0, 0.1)';
                                
                                const inputMarker = document.createElement('div');
                                inputMarker.className = 'debug-marker';
                                inputMarker.style.position = 'fixed';
                                inputMarker.style.left = (rect.left - 30) + 'px';
                                inputMarker.style.top = rect.top + 'px';
                                inputMarker.style.backgroundColor = 'green';
                                inputMarker.style.color = 'white';
                                inputMarker.style.padding = '2px 6px';
                                inputMarker.style.borderRadius = '3px';
                                inputMarker.style.zIndex = '10000';
                                inputMarker.style.fontSize = '12px';
                                inputMarker.textContent = `KR输入${result.inputCount + 1}`;
                                document.body.appendChild(inputMarker);
                                
                                result.inputCount++;
                            } else {
                                console.log(`过滤掉负坐标输入框: y=${rect.top}`);
                            }
                            break;
                        }
                    }
                }
            });
            
            // 按从上到下的顺序排序
            result.inputElements.sort((a, b) => a.y - b.y);
            
            result.success = true;
            result.message = `找到OKR表格，包含 ${result.inputCount} 个输入区域（已过滤负坐标）`;
            return result;
        }""")
        
        if not table_info.get('success'):
            print(f"  [X] {table_info.get('message')}")
            return False
        
        print(f"  [OK] {table_info.get('message')}")
        print(f"  在OKR表格内找到可见输入区域: {table_info.get('inputCount', 0)} 个")
        
        if table_info.get('inputCount') != total_krs:
            print(f"  [⚠️] 表格内可见输入区域数量({table_info.get('inputCount')})与KR数量({total_krs})不匹配")
            print(f"  预期应有 {total_krs} 个输入框，实际找到 {table_info.get('inputCount')} 个可见输入框")
            
            response = input("  是否继续？(y/n): ").strip().lower()
            if response != 'y':
                return False
        
        await asyncio.sleep(1)
        input("\n  请确认：1.蓝色OKR进度标题 2.红色表格区域 3.绿色输入区域（应为8个，且坐标为正），按Enter继续...")
        
        # ==================== 第二步：获取排序后的输入区域位置 ====================
        print("\n" + "-"*60)
        print("  [步骤2] 获取输入区域位置")
        print("-"*60)
        
        # 重新获取位置（确保最新）
        input_positions = []
        for i, input_info in enumerate(table_info['inputElements']):
            pos = await page.evaluate("""(element) => {
                const rect = element.getBoundingClientRect();
                return {
                    x: rect.left + 10,
                    y: rect.top + 5,
                    rawY: rect.top,
                    exists: rect.width > 0 && rect.height > 0
                };
            }""", input_info['element'])
            
            # 再次确认坐标为正
            if pos.get('exists') and pos['rawY'] > 0:
                input_positions.append(pos)
                print(f"    输入区域 {i+1}: 坐标({pos['x']:.0f}, {pos['y']:.0f}), 原始Y={pos['rawY']:.0f}")
        
        print(f"  最终获取到 {len(input_positions)} 个有效输入区域")
        
        if len(input_positions) != total_krs:
            print(f"  [⚠️] 有效输入区域数量({len(input_positions)})与KR数量({total_krs})不匹配")
            response = input("  是否继续？(y/n): ").strip().lower()
            if response != 'y':
                return False
        
        input("\n  确认所有输入区域坐标均为正数后按Enter开始填写...")
        
        # ==================== 第三步：按顺序填写 ====================
        print("\n" + "-"*60)
        print("  [步骤3] 开始填写")
        print("-"*60)
        
        success_count = 0
        input_index = 0
        
        for obj_idx, obj in enumerate(objectives):
            json_krs = obj.get("key_results", [])
            print(f"\n  [O{obj_idx+1}] 处理 {len(json_krs)} 个KR")
            
            for kr_idx, json_kr in enumerate(json_krs):
                weekly_work = json_kr.get("weekly_work", "")
                
                if isinstance(weekly_work, list):
                    weekly_work = '\n'.join(weekly_work)
                
                if not weekly_work:
                    print(f"    [跳过] O{obj_idx+1}KR{kr_idx+1}: 无内容")
                    continue
                
                print(f"\n    [O{obj_idx+1}KR{kr_idx+1}] 准备填写...")
                print(f"      内容预览: {weekly_work[:60]}...")
                
                if input_index < len(input_positions):
                    pos = input_positions[input_index]
                    
                    print(f"      点击位置: ({pos['x']:.0f}, {pos['y']:.0f})")
                    
                    await page.evaluate(f"window.scrollTo(0, {pos['y'] - 200})")
                    await asyncio.sleep(0.5)
                    
                    await page.mouse.click(pos['x'], pos['y'])
                    await asyncio.sleep(0.2)
                    await page.mouse.dblclick(pos['x'], pos['y'])
                    await asyncio.sleep(0.3)
                    await page.mouse.click(pos['x'], pos['y'])
                    await asyncio.sleep(0.2)
                    
                    await page.keyboard.down('Control')
                    await page.keyboard.press('A')
                    await page.keyboard.up('Control')
                    await asyncio.sleep(0.2)
                    await page.keyboard.press('Delete')
                    await asyncio.sleep(0.2)
                    
                    print(f"      输入内容中...")
                    lines = weekly_work.split('\n')
                    for i, line in enumerate(lines):
                        if line.strip():
                            await page.keyboard.type(line, delay=20)
                            if i < len(lines) - 1:
                                await page.keyboard.press('Enter')
                                await asyncio.sleep(0.2)
                    
                    await page.keyboard.press('Enter')
                    await asyncio.sleep(0.2)
                    
                    print(f"      [OK] 填写完成")
                    success_count += 1
                else:
                    print(f"      [⚠️] 没有更多输入区域")
                
                input_index += 1
        
        print(f"\n  [填写统计] {success_count}/{total_krs} 个KR填写成功")
        
        # ==================== 第四步：清除标记 ====================
        print("\n" + "-"*60)
        print("  [步骤4] 清除标记")
        print("-"*60)
        
        await page.evaluate("""() => {
            document.querySelectorAll('[style*="border"]').forEach(el => {
                el.style.border = '';
                el.style.backgroundColor = '';
            });
            document.querySelectorAll('.debug-marker').forEach(el => el.remove());
        }""")
        
        print("  [OK] 标记已清除")
        
        return success_count > 0
        
    except Exception as e:
        print(f"\n  [X] 填写过程出错: {e}")
        import traceback
        traceback.print_exc()
        return False

async def fill_part3_metrics(page, part3_data):
    """填写第三部分：关键指标"""
    print("\n[填写] 第三部分：关键指标")

    sections = part3_data.get("sections", {})
    if not sections:
        print("  [X] 没有指标数据")
        return False

    success_count = 0
    total_sections = len(sections)

    for section_name, section_data in sections.items():
        content = section_data.get("formatted_content", "")
        if not content:
            content = section_data.get("raw_content", "")
            if isinstance(content, list):
                content = "\n".join([str(item) for item in content])

        if not content:
            print(f"  [跳过] {section_name}: 无内容")
            continue

        # 显示准备写入的内容
        print(f"\n  {'='*60}")
        print(f"  [准备写入 - {section_name}]")
        print(f"  {'='*60}")
        print(f"  内容长度: {len(content)} 字符")
        print(f"  完整内容:\n{content}")
        print(f"  {'='*60}")

        # 尝试多种标签匹配
        label_variants = [section_name]

        # 添加常见变体
        if "业务核心指标" in section_name:
            label_variants.extend(["核心指标", "业务指标", "指标"])
        elif "主要项目" in section_name:
            label_variants.extend(["项目", "重点项目"])
        elif "下周重点工作" in section_name:
            label_variants.extend(["下周工作", "重点工作", "下周计划"])

        filled = False
        for label in label_variants:
            try:
                result = await smart_fill_input(page, label, content)
                if result:
                    print(f"    [OK] 已填写: {label} ({len(content)}字符)")
                    filled = True
                    break
            except:
                continue

        if not filled:
            # 使用 JavaScript 兜底
            result = await fill_by_javascript(page, section_name, content)
            if result:
                print(f"    [OK] JavaScript填写成功")
                filled = True
            else:
                print(f"    [X] 填写失败: {section_name}")

        if filled:
            success_count += 1

    print(f"\n  填写统计: {success_count}/{total_sections} 个section成功")
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

async def smart_click_with_fallback(page, locator, timeout=5000):
    """智能点击：尝试标准点击 -> 强制点击 -> JavaScript点击"""
    # 先尝试移除遮挡层
    await remove_overlay_elements(page)

    # 1. 尝试标准点击
    try:
        await locator.click(timeout=timeout)
        print("    [OK] 标准点击成功")
        return True
    except Exception as e:
        if "intercepts pointer events" in str(e):
            print("    [!] 被遮挡层拦截，尝试强制点击...")
        else:
            print(f"    [!] 标准点击失败: {e}")

    # 2. 尝试强制点击
    try:
        await locator.click(force=True, timeout=timeout)
        print("    [OK] 强制点击成功")
        return True
    except Exception as e:
        print(f"    [!] 强制点击失败: {e}")

    # 3. 使用JavaScript点击
    try:
        handle = await locator.element_handle()
        if handle:
            await handle.evaluate("el => { el.click(); el.dispatchEvent(new Event('click', {bubbles: true})); }")
            print("    [OK] JavaScript点击成功")
            return True
    except Exception as e:
        print(f"    [!] JavaScript点击失败: {e}")

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

async def diagnose_save_issue(page):
    """诊断保存问题的原因"""
    
    print("\n" + "="*60)
    print("[诊断] 开始分析保存问题")
    print("="*60)
    
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
    
    print("\n  [分析] 当前编辑器内容:")
    has_invisible = False
    for p in current_content:
        print(f"    段落 {p['index']}: {p['text']}")
        print(f"        长度: {p['length']}, 字符编码: {p['charCodes']}")
        
        # 检查不可见字符
        for code in p['charCodes']:
            if code < 32 or code == 8203:  # 控制字符或零宽空格
                has_invisible = True
                print(f"        ⚠️ 发现不可见字符: {code}")
    
    if has_invisible:
        print("\n  [⚠️ 发现] 存在不可见字符，可能需要清理")
    else:
        print("\n  [✅ 正常] 未发现不可见字符")
    
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
    
    print(f"\n  [分析] 编辑器信息:")
    if editor_info:
        print(f"    类型: {editor_info['tagName']}.{editor_info['className']}")
        print(f"    有React属性: {editor_info['hasReactProps']} (找到 {editor_info['reactPropsCount']} 个)")
        print(f"    有Slate属性: {editor_info['hasSlateProps']}")
    
    # 3. 测试写入并触发事件
    print("\n  [测试] 尝试写入测试内容...")
    
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
        print(f"\n  [测试结果]")
        print(f"    写入内容: {test_result['testContent']}")
        print(f"    时间戳: {test_result['timestamp']}")
        print("\n  ⚠️ 黄色高亮段落是测试内容，请现在手动点击保存按钮")
    else:
        print(f"\n  [测试失败] {test_result.get('error')}")
        return
    
    input("\n  手动保存后按Enter继续验证...")
    
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
    
    print(f"\n  [验证结果]")
    print(f"    测试内容是否还在: {verify_result['found']}")
    if verify_result['found']:
        print(f"    保存后的内容: {verify_result['content']}")
        print("\n  ✅ 测试内容被成功保存！说明事件触发正确")
    else:
        print("\n  ❌ 测试内容丢失！说明事件触发可能有问题")
        print("     可能原因：")
        print("     - 需要触发更多事件")
        print("     - 编辑器有内部状态需要更新")
        print("     - 需要聚焦编辑器")
    
    return {
        "has_invisible_chars": has_invisible,
        "editor_info": editor_info,
        "test_saved": verify_result['found']
    }
    """诊断保存问题的原因"""
    
    print("\n[诊断] 开始分析保存问题")
    
    # 1. 获取当前编辑器内容
    current_content = await page.evaluate("""() => {
        const editor = document.querySelector('[data-slate-editor="true"]');
        if (!editor) return null;
        
        // 获取所有段落内容
        const paragraphs = editor.querySelectorAll('.mp-paragraph-wrapper');
        const content = [];
        paragraphs.forEach(p => {
            const text = p.textContent || '';
            content.push({
                text: text,
                length: text.length,
                charCodes: Array.from(text).map(c => c.charCodeAt(0))
            });
        });
        return content;
    }""")
    
    print("\n  [分析] 当前编辑器内容:")
    for i, p in enumerate(current_content):
        print(f"    段落 {i}: {p['text'][:50]}...")
        print(f"        长度: {p['length']}, 字符编码: {p['charCodes'][:10]}...")
    
    # 2. 检查是否有不可见字符
    has_invisible = False
    for p in current_content:
        for code in p['charCodes']:
            if code < 32 or code == 8203:  # 控制字符或零宽空格
                has_invisible = True
                print(f"    ⚠️ 发现不可见字符: {code}")
    
    if has_invisible:
        print("\n  [发现] 存在不可见字符，可能需要清理")
    
    # 3. 检查编辑器的事件监听
    event_listeners = await page.evaluate("""() => {
        const editor = document.querySelector('[data-slate-editor="true"]');
        if (!editor) return null;
        
        // 获取编辑器绑定的React属性
        const reactProps = Object.keys(editor).filter(key => key.startsWith('__reactProps$'));
        
        return {
            hasReactProps: reactProps.length > 0,
            reactProps: reactProps
        };
    }""")
    
    print(f"\n  [分析] 编辑器React属性: {event_listeners}")
    
    # 4. 模拟手动编辑并捕获事件
    print("\n  [测试] 模拟手动编辑...")
    
    # 找到第一个有内容的段落
    target_para_index = None
    for i, p in enumerate(current_content):
        if p['text'] and not p['text'].includes('本周工作总结') and not p['text'].includes('重要会议总结'):
            target_para_index = i
            break
    
    if target_para_index is not None:
        # 先清除已有内容
        await page.evaluate(f"""() => {{
            const paragraphs = document.querySelectorAll('.mp-paragraph-wrapper');
            if (paragraphs[{target_para_index}]) {{
                const p = paragraphs[{target_para_index}];
                while (p.firstChild) p.removeChild(p.firstChild);
                
                // 添加测试内容
                const spanWrapper = document.createElement('span');
                spanWrapper.setAttribute('data-slate-node', 'text');
                
                const spanLeaf = document.createElement('span');
                spanLeaf.className = '';
                spanLeaf.setAttribute('data-slate-leaf', 'true');
                
                const contentSpan = document.createElement('span');
                contentSpan.setAttribute('data-slate-string', 'true');
                contentSpan.textContent = '测试内容 ' + Date.now();
                
                spanLeaf.appendChild(contentSpan);
                spanWrapper.appendChild(spanLeaf);
                p.appendChild(spanWrapper);
                
                // 触发各种事件
                p.dispatchEvent(new Event('input', {{bubbles: true}}));
                p.dispatchEvent(new Event('change', {{bubbles: true}}));
                p.dispatchEvent(new Event('blur', {{bubbles: true}}));
                p.dispatchEvent(new Event('keydown', {{bubbles: true}}));
                p.dispatchEvent(new Event('keyup', {{bubbles: true}}));
            }}
        }}""")
        
        print("  [OK] 已写入测试内容，请手动点击保存按钮")
        input("  请手动点击保存，然后观察是否保存成功。完成后按Enter继续...")
        
        # 检查保存后内容是否还在
        after_save = await page.evaluate("""() => {
            const paragraphs = document.querySelectorAll('.mp-paragraph-wrapper');
            const lastPara = paragraphs[paragraphs.length - 1];
            return lastPara ? lastPara.textContent : null;
        }""")
        
        print(f"  保存后最后段落内容: {after_save}")
    
    # 5. 建议的修复方案
    print("\n  [建议] 可能的解决方案:")
    print("    1. 使用更完整的事件触发:")
    print("       - input (必须)")
    print("       - change (必须)")
    print("       - blur (让编辑器认为焦点离开)")
    print("       - keydown/keyup (模拟用户输入)")
    print("    2. 清理不可见字符")
    print("    3. 尝试先聚焦再输入")
    print("    4. 使用编辑器的API（如果能找到）")
    
    return has_invisible

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

    # if part1_data:
        # results['part1'] = await fill_part1_work_summary(page, part1_data)    
    # else:
        # print("\n[!] 无第一部分数据，跳过")
        # results['part1'] = False


    if part2_data:
        results['part2'] = await fill_okr_section(page, part2_data)
    else:
        print("\n[!] 无第二部分数据，跳过")
        results['part2'] = False

    # if part3_data:
        # results['part3'] = await fill_part3_metrics(page, part3_data)
    # else:
        # print("\n[!] 无第三部分数据，跳过")
        # results['part3'] = False
####
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

async def find_editor_api(page):
    """查找编辑器可用的API"""
    
    print("\n[API诊断] 查找编辑器内部API...")
    
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
    
    print(f"\n  [诊断详情]")
    if api_info.get('error'):
        print(f"    错误: {api_info['error']}")
    else:
        print(f"    找到React属性: {len(api_info.get('props', []))} 个")
        for i, prop in enumerate(api_info.get('props', [])[:5]):
            print(f"      - {prop}")
        
        print(f"\n    找到Slate编辑器: {api_info.get('hasSlate', False)}")
        if api_info.get('hasSlate'):
            print(f"    编辑器方法数量: {len(api_info.get('methods', []))}")
            if api_info.get('methods'):
                print(f"    前10个方法: {api_info['methods'][:10]}")
    
    return api_info

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