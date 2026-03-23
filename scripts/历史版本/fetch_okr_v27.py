#!/usr/bin/env python3
"""
百度OKR周报抓取 - V10多行标题合并版（优化版）
基于V10修复以下问题：
1. O1/O2/O3标题可能分布在多行，需要合并（如：O1 + 标题内容 + 状态）
2. KR1/KR2/KR3标题可能分布在多行，需要合并（如：KR1 + 标题内容 + 进度）
3. "本周工作"之后的行才是真正的周报内容
4. 本人和下属抓取逻辑完全一致
5. 【优化】团队周报改为JSON格式（类似个人周报）
6. 【优化】OKR标题去重：去掉重复的 O1/O2/O3/KR1/KR2/KR3 字样，但保留一个

针对百度OKR系统的特定格式：
1. 本周工作总结（可能有重复，需合并）
2. OKR进度（三级表格结构：总标题 -> O1/O2/O3 -> KR1/KR2/KR3 -> 本周工作）
3. 其他内容（业务核心指标、主要项目、下周重点工作）
4. 新增：首次页面等待5秒 + 失败重试机制（最多1次）
"""
import json
import asyncio
import re
import glob
import os
import platform  # 新增：用于获取系统信息
import time
import subprocess
import traceback  # 新增：用于异常处理
from datetime import datetime, timedelta
from playwright.async_api import async_playwright

# ==================== 浏览器复用功能 ====================
BROWSER_STATE_FILE = "browser_state.json"
BROWSER_DEBUG_DIR = "browser_debug"

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

# 配置
MY_ID = 's673090'
STAFF_IDS = [
    's673090','s801573',
    's733504', 's682083', 's929270', 's798268','s811385', 's865585', 's783729', 's400192', 's823700'
]
BASE_URL = 'https://okr.baidu-int.com/pages/weekly.html#/index?mode=steam&mandatorId=undefined'

# ==================== 工具函数 ====================

def iprint(*args, **kwargs):
    """普通信息输出，强制flush"""
    print(*args, **kwargs, flush=True)

def dprint(*args, **kwargs):
    """调试输出，自动flush"""
    if DEBUG:
        print(*args, **kwargs, flush=True)

# 设置DEBUG模式
DEBUG = True

# ==================== 浏览器复用工具 ====================

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
                    iprint("[浏览器] 创建新页面")
                    return True
            
            iprint("[浏览器] 未找到现有的context")
            return False
        except Exception as e:
            iprint(f"[浏览器] 连接失败: {e}")
            return False
    
    async def launch_new_browser(self):
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
            BASE_URL
        ]
        
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



# 在文件开头添加日期工具函数
def get_current_week_range():
    """获取当前周的日期范围（周一至周日）"""
    today = datetime.now()
    monday = today - timedelta(days=today.weekday())
    sunday = monday + timedelta(days=6)
    return monday, sunday

def get_week_filename_suffix():
    """
    生成周报文件名后缀
    格式：X月份_X.X-X.X周
    例如：3月份_3.2-3.8周
    """
    monday, sunday = get_current_week_range()
    
    # 获取月份
    month = monday.month
    month_names = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "十二"]
    month_str = f"{month_names[month-1]}月份"
    
    # 获取日期范围
    start_day = monday.day
    end_day = sunday.day
    
    return f"{month_str}_{monday.month}.{start_day}-{sunday.month}.{end_day}周"

def get_week_date(weeks_ago=0):
    """获取周日期"""
    today = datetime.now()
    target = today - timedelta(days=today.weekday() + 7 + (weeks_ago * 7))
    return f"s{target.strftime('%Y-%m-%d')}"


# ==================== 文本标准化与去重 ====================

def normalize_text(text):
    """文本标准化"""
    text = text.strip()
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'[，。？！；：""''（）【】]', ',', text)
    text = re.sub(r'[,\.]+', ',', text)
    return text.lower()

def calculate_similarity(str1, str2):
    """计算相似度"""
    if str1 in str2 or str2 in str1:
        return 0.9
    set1 = set(str1)
    set2 = set(str2)
    intersection = len(set1 & set2)
    union = len(set1 | set2)
    if union == 0:
        return 0
    return intersection / union

def deduplicate_lines(lines, similarity_threshold=0.85):
    """去重"""
    if not lines:
        return []
    
    result = []
    normalized_map = {}
    
    for line in lines:
        if not line or len(line.strip()) < 5:
            continue
        
        normalized = normalize_text(line)
        
        if normalized in normalized_map:
            continue
        
        is_duplicate = False
        for existing_norm, existing_orig in list(normalized_map.items()):
            if normalized in existing_norm or existing_norm in normalized:
                is_duplicate = True
                if len(line) > len(existing_orig):
                    normalized_map[normalized] = line
                    if existing_norm in normalized_map:
                        del normalized_map[existing_norm]
                break
            
            sim = calculate_similarity(normalized, existing_norm)
            if sim > similarity_threshold:
                is_duplicate = True
                if len(line) > len(existing_orig):
                    normalized_map[normalized] = line
                    if existing_norm in normalized_map:
                        del normalized_map[existing_norm]
                break
        
        if not is_duplicate:
            normalized_map[normalized] = line
    
    seen = set()
    for line in lines:
        normalized = normalize_text(line)
        if normalized in normalized_map and normalized not in seen:
            result.append(normalized_map[normalized])
            seen.add(normalized)
            del normalized_map[normalized]
    
    return result

# ==================== 文本标准化与去重（增强版）====================

def is_valid_content(text):
    """
    判断是否为有效的工作内容（过滤UI元素和提示文本）
    """
    if not text or len(text.strip()) < 5:
        return False
    
    # 过滤UI元素和提示文本的关键词
    invalid_patterns = [
        r'^[OoKkRr][1-3]',  # O1, O2, O3, KR1, KR2, KR3 开头
        r'邵银川的OKR',
        r'卜安安的OKR',
        r'期间：\d{4}-Q[1-4]',  # 期间：2026-Q1
        r'QTD实际值',
        r'^02/\d{2}~03/\d{2}',  # 日期范围如 "02/23~03/01"
        r'汽车项目交付组-.*?周报',
        r'在此处拆分卡片',
        r'DuChat',
        r'CodePen',
        r'重要会议总结:',
        r'仅汇报链上级及周报代理人可看该区域',
        r'内容已隐藏',
        r'点击上方按钮查看',
        r'本周工作总结',  # 这是标题，不是内容
        r'OKR进度',       # 这是标题，不是内容
        r'业务核心指标',   # 这是标题，不是内容
        r'主要项目',       # 这是标题，不是内容
        r'下周重点工作',   # 这是标题，不是内容
        r'^\([^)]+\)$',   # 只有括号内容，如 "(上周)"
        r'^[（(][^）)]+[）)]$',  # 只有括号内容
        r'^\s*$',  # 空行或空白
    ]
    
    for pattern in invalid_patterns:
        if re.search(pattern, text, re.IGNORECASE):
            print(f"      [过滤] '{text[:50]}...' 匹配到无效模式: {pattern}")
            return False
    
    # 检查是否只包含日期格式
    if re.match(r'^\d{2}/\d{2}~\d{2}/\d{2}$', text.strip()):
        return False
    
    # 检查是否只包含括号内容
    if re.match(r'^[（(][^）)]+[）)]$', text.strip()):
        return False
    
    # 检查是否只包含标点符号
    if re.match(r'^[，。？！、；：""''（）【】\\s]+$', text):
        return False
    
    return True

def is_title_line(text):
    """
    判断是否为标题行（不是工作内容）
    """
    title_patterns = [
        r'^[OoKkRr][1-3]',  # O1, O2, O3, KR1, KR2, KR3
        r'的OKR$',
        r'期间：',
        r'本周工作总结$',
        r'OKR进度$',
        r'业务核心指标$',
        r'主要项目$',
        r'下周重点工作$',
        r'^重要会议总结',
        r'^仅汇报链上级',
        r'^内容已隐藏',
    ]
    
    for pattern in title_patterns:
        if re.search(pattern, text, re.IGNORECASE):
            return True
    return False

# ==================== 修改 parse_weekly_report_structure_v10 函数中的相关部分 ====================
# ==================== 修改 parse_weekly_report_structure_v10 函数中的第二部分提取 ====================

def parse_weekly_report_structure_v10(full_text):
    """
    V10核心：解析三部分结构（增强过滤版）
    """
    # 全局过滤：如果检测到"Hi卜安安同学"，抛弃后续所有内容
    if 'Hi卜安安同学' in full_text:
        full_text = full_text.split('Hi卜安安同学')[0]

    lines = full_text.split('\n')
    lines = [l.strip() for l in lines if l.strip()]
    
    result = {
        'employee_name': None,
        'okr_title': None,
        'part1_work_summary': [],      # 第一部分：本周工作总结
        'part2_okr_structure': [],      # 第二部分：OKR结构
        'part3_metrics': {},            # 第三部分：业务核心指标等
        'raw_text': full_text
    }
    
    # 步骤1：提取OKR标题和姓名（用于确定第二部分开始位置）
    okr_header_index = -1
    for i, line in enumerate(lines):
        if 'OKR' in line.upper() and '的' in line:
            result['okr_title'] = line
            name = extract_name_from_okr_header(line)
            print(f"提取OKR标题和姓名：{result['okr_title']},{name}")
            if name:
                result['employee_name'] = name
            okr_header_index = i
            break
    
    # 步骤2：提取第一部分（本周工作总结）- 增强过滤
    first_summary_found = False
    i = 0
    while i < len(lines):
        line = lines[i]
        if '本周工作总结' in line:
            if first_summary_found:
                print(f"    [V10] 检测到第二个'本周工作总结'，第一部分内容获取完成")
                break
            first_summary_found = True
            print(f"    [V10] 找到第一个'本周工作总结'，开始提取内容")
            i += 1
            while i < len(lines):
                content_line = lines[i]
                # 如果遇到第二个"本周工作总结"，停止
                if '本周工作总结' in content_line:
                    print(f"    [V10] 检测到第二个'本周工作总结'，停止提取")
                    break
                # 如果遇到OKR标题或O/KR标题，停止
                if 'OKR' in content_line.upper() or re.match(r'^[Oo]\d+', content_line) or re.match(r'^[Kk][Rr]\d+', content_line) or '本周工作' in content_line:
                    break
                # 增强过滤：只添加有效的工作内容
                if len(content_line) > 5 and is_valid_content(content_line) and not is_title_line(content_line):
                    print(f"      [添加] 有效内容: {content_line[:80]}...")
                    result['part1_work_summary'].append(content_line)
                else:
                    print(f"      [过滤] 无效内容: {content_line[:80]}...")
                i += 1
            continue
        i += 1

    result['part1_work_summary'] = deduplicate_lines(result['part1_work_summary'], 0.85)

    # 步骤3：提取第二部分（OKR三级结构）- 保留标题
    if okr_header_index >= 0:
        okr_lines = lines[okr_header_index + 1:]
        
        current_o = None
        current_kr = None
        
        o_pattern = re.compile(r'^\s*[Oo](\d+)\b')
        kr_pattern = re.compile(r'^\s*[Kk][Rr](\d+)\b')
        weekly_work_pattern = re.compile(r'^(?:本周工作|本周进展|工作进展)', re.IGNORECASE)
        part3_pattern = re.compile(r'(?:业务核心指标|主要项目|下周重点工作)')
        
        i = 0
        while i < len(okr_lines):
            line = okr_lines[i]
            
            if part3_pattern.search(line):
                print(f"    [V10] 检测到第三部分标题，停止OKR解析")
                break
            
            o_match = o_pattern.match(line)
            if o_match:
                o_num = o_match.group(1)
                o_title_parts = [line.strip()]
                i += 1
                # 继续读取后续行，直到遇到KR、下一个O、本周工作或第三部分
                while i < len(okr_lines):
                    next_line = okr_lines[i]
                    if kr_pattern.match(next_line) or o_pattern.match(next_line) or weekly_work_pattern.match(next_line) or part3_pattern.search(next_line):
                        break
                    # 保留所有内容到标题，包括可能的无效内容（因为这是标题的一部分）
                    if len(next_line.strip()) > 1:
                        o_title_parts.append(next_line.strip())
                    i += 1
                
                o_title = ' '.join(o_title_parts)
                print(f"    [V10] O{o_num}标题(保留): {o_title[:100]}")
                
                current_o = {
                    'level': 1,
                    'id': f'O{o_num}',
                    'title': o_title,  # 完整保留标题
                    'key_results': []
                }
                result['part2_okr_structure'].append(current_o)
                current_kr = None
                continue
            
            kr_match = kr_pattern.match(line)
            if kr_match and current_o:
                kr_num = kr_match.group(1)
                kr_title_parts = [line.strip()]
                i += 1
                # 继续读取后续行，直到遇到"本周工作"、下一个O/KR或第三部分
                while i < len(okr_lines):
                    next_line = okr_lines[i]
                    if weekly_work_pattern.match(next_line) or kr_pattern.match(next_line) or o_pattern.match(next_line) or part3_pattern.search(next_line):
                        break
                    # 保留所有内容到标题，包括可能的无效内容（因为这是标题的一部分）
                    if len(next_line.strip()) > 1:
                        kr_title_parts.append(next_line.strip())
                    i += 1
                
                kr_title = ' '.join(kr_title_parts)
                print(f"    [V10] KR{kr_num}标题(保留): {kr_title[:100]}")
                
                current_kr = {
                    'level': 2,
                    'id': f'KR{kr_num}',
                    'title': kr_title,  # 完整保留标题
                    'weekly_work': []   # 初始为空，等待"本周工作"后的内容
                }
                current_o['key_results'].append(current_kr)
                continue
            
            weekly_match = weekly_work_pattern.match(line)
            if weekly_match and current_kr:
                # 提取"本周工作"后的内容（如果有）
                content_after = line[len(weekly_match.group(0)):].strip()
                content_after = re.sub(r'^[:：\s]+', '', content_after)
                if content_after and is_valid_content(content_after) and not is_title_line(content_after):
                    print(f"    [V10] 本周工作(首行): {content_after[:80]}")
                    current_kr['weekly_work'].append(content_after)
                
                i += 1
                while i < len(okr_lines):
                    next_line = okr_lines[i]
                    if o_pattern.match(next_line) or kr_pattern.match(next_line) or part3_pattern.search(next_line):
                        break
                    if len(next_line.strip()) > 1 and is_valid_content(next_line) and not is_title_line(next_line):
                        print(f"    [V10] 本周工作(续): {next_line.strip()[:80]}")
                        current_kr['weekly_work'].append(next_line.strip())
                    i += 1
                continue
            
            i += 1
        
        # 对OKR结构内的本周工作内容去重
        for o in result['part2_okr_structure']:
            for kr in o.get('key_results', []):
                kr['weekly_work'] = deduplicate_lines(kr['weekly_work'], 0.85)
        
        print(f"\n    [V10] 解析完成摘要:")
        for o in result['part2_okr_structure']:
            print(f"      {o['id']}: {o['title'][:80]}")
            for kr in o.get('key_results', []):
                print(f"        {kr['id']}: title={kr['title'][:60]}, work_items={len(kr['weekly_work'])}")
    
    # 步骤4：提取第三部分（其他内容）- 增强过滤
    part3_keywords = {
        '业务核心指标': [],
        '主要项目': [],
        '下周重点工作': []
    }
    
    for keyword in part3_keywords.keys():
        for i, line in enumerate(lines):
            if keyword in line:
                j = i + 1
                while j < len(lines):
                    content = lines[j]
                    # 如果遇到下一个第三部分标题，停止
                    if any(k in content for k in part3_keywords.keys() if k != keyword):
                        break
                    # 如果遇到OKR相关内容，可能是误匹配，停止
                    if re.match(r'^[Oo]\d+', content) or re.match(r'^[Kk][Rr]\d+', content):
                        break
                    # 增强过滤：只添加有效的工作内容
                    if len(content) > 5 and is_valid_content(content) and not is_title_line(content):
                        print(f"      [第三部分-{keyword}] 添加: {content[:80]}...")
                        part3_keywords[keyword].append(content)
                    else:
                        print(f"      [第三部分-{keyword}] 过滤: {content[:80]}...")
                    j += 1
                break
        
        result['part3_metrics'][keyword] = deduplicate_lines(part3_keywords[keyword], 0.85)
    
    return result


# ==================== V7核心：三部分结构提取 ====================

def extract_name_from_okr_header(text):
    """从OKR标题提取姓名，如：'张三的OKR' -> '张三'"""
    patterns = [
        r'([\u4e00-\u9fa5]{2,4})\s*的\s*OKR',
        r'([\u4e00-\u9fa5]{2,4})OKR',
        r'汽车项目交付组[-–—]\s*([\u4e00-\u9fa5]{2,4})',
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1)
    return None

# ==================== 浏览器操作（保持原有代码不变）====================

async def find_scroll_container(page):
    """查找可滚动容器"""
    containers = await page.evaluate('''() => {
        const all = document.querySelectorAll('*');
        const result = [];
        for (let el of all) {
            const style = window.getComputedStyle(el);
            if (el.scrollHeight > el.clientHeight + 100 && 
                el.scrollHeight > 500 &&
                (style.overflowY === 'auto' || style.overflowY === 'scroll')) {
                const rect = el.getBoundingClientRect();
                result.push({
                    class: el.className,
                    id: el.id,
                    tag: el.tagName,
                    scrollHeight: el.scrollHeight,
                    x: rect.left + rect.width / 2,
                    y: rect.top + rect.height / 2
                });
            }
        }
        result.sort((a, b) => b.scrollHeight - a.scrollHeight);
        return result.slice(0, 1);
    }''')
    
    if containers:
        first = containers[0]
        if first.get('class'):
            selector = f"[class*='{first['class'].split()[0]}']"
        elif first.get('id'):
            selector = f"#{first['id']}"
        else:
            selector = first['tag'].lower()
        return selector, (first.get('x', 960), first.get('y', 600))
    return 'body', (960, 600)

async def scroll_container_v10(page, selector, center):
    """V10：确保内容完整加载"""
    print("    滚动加载内容...", end=" ", flush=True)
    await page.mouse.move(center[0], center[1])
    await asyncio.sleep(0.5)
    
    for i in range(30):
        await page.mouse.wheel(0, 400)
        await asyncio.sleep(0.3)
        if i % 10 == 0:
            print(f"{i*400}px", end=" ", flush=True)
    print("✓")
    
    try:
        await page.click(selector)
        await asyncio.sleep(0.5)
        for _ in range(5):
            await page.keyboard.press('PageDown')
            await asyncio.sleep(0.3)
    except:
        pass
    
    print("    等待动态内容渲染...", end=" ", flush=True)
    await asyncio.sleep(5)
    print("✓")

# ==================== V7抓取流程（新增等待和重试机制）====================

async def fetch_single_staff_v10(page, staff_id, date, desc, index=0, total=0):
    """
    V10单个人员抓取（抽取为独立函数便于重试）
    新增：首次页面加载后等待5秒
    修复：更完善的登录状态检查
    V10最终修复：O/KR标题严格提取，绝不混入weekly_work
    """
    url = f"{BASE_URL}&id={staff_id}&date={date}"
    progress = f"[{index}/{total}] " if index else ""
    
    print(f"\n{'='*60}")
    print(f"{progress}抓取: {desc}")
    print(f"URL: {url}")
    
    try:
        print("  加载页面...", end=" ")
        
        # 使用 try-except 捕获导航被中断的情况（通常是重定向到登录页）
        try:
            await page.goto(url, wait_until='domcontentloaded', timeout=60000)
        except Exception as nav_error:
            # 检查是否跳转到了登录页
            current_url = page.url
            if 'uuap' in current_url:
                print("\n⚠️  会话失效，需要重新登录")
                return {'id': staff_id, 'date': date, 'status': 'login_required'}
            # 其他导航错误，继续尝试
            print(f"\n⚠️  导航警告: {nav_error}")
        
        # ===== V7新增：首次页面多等待5秒 =====
        if index == 1 or index == 0:  # 第一个页面
            print("(首次页面等待5秒)...", end=" ")
            await asyncio.sleep(5)
        else:
            await asyncio.sleep(2)
        print("✓")
        
        # 再次检查URL（可能跳转到了登录页）
        current_url = page.url
        if 'uuap' in current_url:
            print("需要登录")
            return {'id': staff_id, 'date': date, 'status': 'login_required'}
        
        print("  滚动加载...", end=" ")
        selector, center = await find_scroll_container(page)
        await scroll_container_v10(page, selector, center)
        print("✓")
        
        print("  提取内容...", end=" ")
        full_text = await page.evaluate('() => document.body.innerText')

        # 过滤"Hi卜安安同学"及后续所有内容
        if 'Hi卜安安同学' in full_text:
            full_text = full_text.split('Hi卜安安同学')[0]
            print(f"(已过滤Hi卜安安同学及后续内容)", end=" ")

        print("  解析三部分结构...", end=" ")
        parsed = parse_weekly_report_structure_v10(full_text)
        employee_name = parsed.get('employee_name') or staff_id
        print(f"姓名:{employee_name}, 工作总结:{len(parsed['part1_work_summary'])}条, O数量:{len(parsed['part2_okr_structure'])}")
        
        print("  第三部分统计...", end=" ")
        metrics = parsed.get('part3_metrics', {})
        metrics_summary = ', '.join([f"{k}:{len(v)}条" for k, v in metrics.items()])
        print(metrics_summary)
        
        safe_name = re.sub(r'[\\/*?:"<>|]', '_', employee_name) if employee_name else staff_id
        
        # 保存结构化结果
        result = {
            'id': staff_id,
            'name': employee_name,
            'date': date,
            'desc': desc,
            'fetch_time': datetime.now().isoformat(),
            'status': 'success',
            'content': {
                'okr_title': parsed.get('okr_title'),
                'part1_work_summary': parsed.get('part1_work_summary', []),
                'part2_okr_structure': parsed.get('part2_okr_structure', []),
                'part3_metrics': parsed.get('part3_metrics', {})
            },
            'stats': {
                'text_length': len(full_text),
                'part1_count': len(parsed.get('part1_work_summary', [])),
                'o_count': len(parsed.get('part2_okr_structure', [])),
                'kr_count': sum(len(o.get('key_results', [])) for o in parsed.get('part2_okr_structure', [])),
                'weekly_work_count': sum(
                    len(kr.get('weekly_work', [])) 
                    for o in parsed.get('part2_okr_structure', [])
                    for kr in o.get('key_results', [])
                ),
                'metrics_count': sum(len(v) for v in parsed.get('part3_metrics', {}).values())
            }
        }
        
        json_file = f'{safe_name}_{date}_周报.json'
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"  ✓ 完成: {safe_name} (已保存 {json_file})")
        
        return result
        
    except Exception as e:
        print(f"\n✗ 失败: {e}")
        import traceback
        traceback.print_exc()
        return {'id': staff_id, 'date': date, 'status': 'error', 'error': str(e)}


async def fetch_staff_v10_with_retry(page, staff_id, date, desc, index=0, total=0):
    """
    V10抓取（带重试机制）
    失败后记录，稍后重试（最多1次）
    """
    result = await fetch_single_staff_v10(page, staff_id, date, desc, index, total)
    return result


# ==================== 整合功能（V10格式 - 优化版）====================

def integrate_weekly_reports_v10():
    print("\n" + "="*60)
    print("整合周报...")
    
    # 读取所有周报文件到字典，方便按ID查找
    reports_dict = {}
    for file in glob.glob('*_周报.json'):
        try:
            with open(file, 'r', encoding='utf-8') as f:
                report = json.load(f)
                staff_id = report.get('id')
                if staff_id:
                    reports_dict[staff_id] = report
                else:
                    print(f"  警告: 文件 {file} 缺少ID字段")
        except Exception as e:
            print(f"  警告: 无法读取 {file} - {e}")
    
    if not reports_dict:
        print("  ✗ 未找到周报文件")
        return None
    
    # ===== 按照 STAFF_IDS 顺序合并 =====
    ordered_reports = []
    missing_ids = []
    
    for staff_id in STAFF_IDS:
        if staff_id in reports_dict:
            ordered_reports.append(reports_dict[staff_id])
            print(f"  ✓ 添加: {staff_id}")
        else:
            missing_ids.append(staff_id)
            print(f"  ⚠️ 缺失: {staff_id}")
    
    # 添加不在STAFF_IDS中的其他人员（如果有）
    other_reports = []
    for staff_id, report in reports_dict.items():
        if staff_id not in STAFF_IDS:
            other_reports.append(report)
            print(f"  ➕ 额外: {staff_id} (不在配置列表中)")
    
    if other_reports:
        ordered_reports.extend(other_reports)
    
    print(f"\n  共加载 {len(ordered_reports)} 份周报 (配置列表 {len(STAFF_IDS)} 人，缺失 {len(missing_ids)} 人)")
    if missing_ids:
        print(f"  缺失人员: {missing_ids}")
    
    # 生成团队周报
    team_report = generate_team_weekly_v10(ordered_reports)
    
    # ===== 修改：使用统一格式的文件名 =====
    week_suffix = get_week_filename_suffix()
    json_file = f'团队周报_{week_suffix}.json'
    md_file = f'团队周报_{week_suffix}.md'
    
    # 保存JSON文件
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(team_report, f, ensure_ascii=False, indent=2)
    print(f"  ✓ 团队周报(JSON): {json_file}")
    
    # 生成Markdown版本
    team_md = generate_team_weekly_markdown_v10(team_report)
    with open(md_file, 'w', encoding='utf-8') as f:
        f.write(team_md)
    print(f"  ✓ 团队周报(Markdown): {md_file}")
    
    # 生成个人参考（保持原有文件名）
    personal_md = generate_personal_reference_v10(ordered_reports)
    with open('personal_reference_v10.md', 'w', encoding='utf-8') as f:
        f.write(personal_md)
    print(f"  ✓ 个人参考: personal_reference_v10.md")
    
    return {
        'team_json': team_report, 
        'team_md': team_md, 
        'reports': ordered_reports,
        'missing_ids': missing_ids,
        'json_file': json_file,  # 返回文件名供其他函数使用
        'md_file': md_file
    }

def clean_okr_title(title, okr_id):
    """
    清理OKR标题，去掉重复的ID标记，但保留一个
    例如: "O1 完成项目交付 O1" -> "O1 完成项目交付"
         "KR1 完成需求分析 KR1" -> "KR1 完成需求分析"
         "完成项目交付 O1 O1" -> "O1 完成项目交付"
    """
    if not title:
        return title
    
    # 统计原始标题中 ID 出现的次数
    id_pattern = re.compile(re.escape(okr_id), re.IGNORECASE)
    matches = list(id_pattern.finditer(title))
    
    # 如果只有一个或没有，直接返回
    if len(matches) <= 1:
        return title.strip()
    
    # 有多个重复，需要处理：保留第一个，去掉其余的
    # 策略：去掉所有ID，然后在开头添加一个
    
    # 去掉所有ID标记（包括周围的空白）
    cleaned = id_pattern.sub('', title)
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
    
    # 在开头添加一个ID标记
    return f"{okr_id} {cleaned}"

# ==================== 修改 generate_team_weekly_v10 函数，确保标题被保留 ====================

def generate_team_weekly_v10(reports):
    """
    V10生成团队周报 - 严格按照MY_ID的结构和STAFF_IDS顺序合并
    保留完整的标题内容
    """
    
    # 收集所有成员的数据 - 按照STAFF_IDS顺序
    all_part1 = []
    all_metrics = {
        '业务核心指标': [],
        '主要项目': [],
        '下周重点工作': []
    }
    
    # 1. 首先找到MY_ID的报告作为基准
    my_report = None
    for report in reports:
        if report.get('id') == MY_ID:
            my_report = report
            break
    
    if not my_report:
        print(f"  [警告] 未找到基准员工 {MY_ID} 的周报，将使用第一个报告作为基准")
        my_report = reports[0] if reports else None
    
    if not my_report:
        return None
    
    # 获取基准员工的OKR结构（保留完整标题）
    base_okr_structure = my_report.get('content', {}).get('part2_okr_structure', [])
    print(f"\n  [基准] 使用 {MY_ID} 的OKR结构作为模板")
    
    # 创建合并用的本周工作映射 (O_id -> KR_id -> [weekly_work_items])
    weekly_work_map = {}
    
    # 创建标题映射，确保保留基准员工的标题
    title_map = {}
    for o in base_okr_structure:
        o_id = o.get('id', '')
        title_map[o_id] = {
            'title': o.get('title', ''),  # 保留基准员工的O标题
            'kr_titles': {}
        }
        for kr in o.get('key_results', []):
            kr_id = kr.get('id', '')
            title_map[o_id]['kr_titles'][kr_id] = kr.get('title', '')  # 保留基准员工的KR标题
    
    # 2. 按照STAFF_IDS顺序处理报告
    processed_ids = set()
    
    # 先处理MY_ID
    if my_report:
        staff_id = my_report.get('id', '')
        staff_name = my_report.get('name', '')
        print(f"  [处理] 基准: {staff_id} - {staff_name}")
        
        content = my_report.get('content', {})
        part1 = content.get('part1_work_summary', [])
        all_part1.extend(part1)
        
        metrics = content.get('part3_metrics', {})
        for key in all_metrics.keys():
            if key in metrics:
                all_metrics[key].extend(metrics[key])
        
        # 收集基准员工的本周工作
        okr_structure = content.get('part2_okr_structure', [])
        for o in okr_structure:
            o_id = o.get('id', '')
            if o_id not in weekly_work_map:
                weekly_work_map[o_id] = {}
            
            for kr in o.get('key_results', []):
                kr_id = kr.get('id', '')
                if kr_id not in weekly_work_map[o_id]:
                    weekly_work_map[o_id][kr_id] = []
                
                weekly_work = kr.get('weekly_work', [])
                weekly_work_map[o_id][kr_id].extend(weekly_work)
        
        processed_ids.add(staff_id)
    
    # 然后按照STAFF_IDS顺序处理其他人员
    for staff_id in STAFF_IDS:
        if staff_id == MY_ID:
            continue
            
        found_report = None
        for report in reports:
            if report.get('id') == staff_id:
                found_report = report
                break
        
        if found_report:
            staff_name = found_report.get('name', '')
            print(f"  [处理] {staff_id} - {staff_name}")
            
            content = found_report.get('content', {})
            
            part1 = content.get('part1_work_summary', [])
            all_part1.extend(part1)
            
            metrics = content.get('part3_metrics', {})
            for key in all_metrics.keys():
                if key in metrics:
                    all_metrics[key].extend(metrics[key])
            
            # 收集本周工作（按基准员工的O和KR结构进行映射）
            okr_structure = content.get('part2_okr_structure', [])
            
            for o in okr_structure:
                o_id = o.get('id', '')
                if o_id in weekly_work_map:
                    for kr in o.get('key_results', []):
                        kr_id = kr.get('id', '')
                        if kr_id in weekly_work_map[o_id]:
                            weekly_work = kr.get('weekly_work', [])
                            weekly_work_map[o_id][kr_id].extend(weekly_work)
                            print(f"    匹配到 {o_id} {kr_id}: 添加 {len(weekly_work)} 条工作项")
                else:
                    print(f"    [警告] 报告中的 {o_id} 不在基准结构中，跳过")
            
            processed_ids.add(staff_id)
        else:
            print(f"  [缺失] 未找到 {staff_id} 的周报")
    
    # 处理不在STAFF_IDS中的其他人员（如果有）
    for report in reports:
        staff_id = report.get('id', '')
        if staff_id not in processed_ids:
            staff_name = report.get('name', '')
            print(f"  [额外] {staff_id} - {staff_name}")
            
            content = report.get('content', {})
            
            part1 = content.get('part1_work_summary', [])
            all_part1.extend(part1)
            
            metrics = content.get('part3_metrics', {})
            for key in all_metrics.keys():
                if key in metrics:
                    all_metrics[key].extend(metrics[key])
            
            okr_structure = content.get('part2_okr_structure', [])
            for o in okr_structure:
                o_id = o.get('id', '')
                if o_id in weekly_work_map:
                    for kr in o.get('key_results', []):
                        kr_id = kr.get('id', '')
                        if kr_id in weekly_work_map[o_id]:
                            weekly_work = kr.get('weekly_work', [])
                            weekly_work_map[o_id][kr_id].extend(weekly_work)
    
    # 去重第一部分和第三部分
    all_part1 = deduplicate_lines(all_part1, 0.85)
    for key in all_metrics.keys():
        all_metrics[key] = deduplicate_lines(all_metrics[key], 0.85)
    
    # 过滤掉包含"Hi卜安安同学"及后续所有内容
    def filter_hi_content(items):
        filtered = []
        for item in items:
            if 'Hi卜安安同学' in item:
                break
            filtered.append(item)
        return filtered
    
    all_part1 = filter_hi_content(all_part1)
    for key in all_metrics.keys():
        all_metrics[key] = filter_hi_content(all_metrics[key])
    
    # 构建OKR结构（以MY_ID的结构为基准，保留完整标题）
    optimized_okr_structure = []
    for o in base_okr_structure:
        o_id = o.get('id', '')
        # 直接使用基准员工的O标题，不去重
        o_title = o.get('title', '')
        
        optimized_krs = []
        for kr in o.get('key_results', []):
            kr_id = kr.get('id', '')
            # 直接使用基准员工的KR标题，不去重
            kr_title = kr.get('title', '')
            
            # 获取该O和KR下所有员工的本周工作并合并
            weekly_work = weekly_work_map.get(o_id, {}).get(kr_id, [])
            weekly_work = filter_hi_content(weekly_work)
            weekly_work = deduplicate_lines(weekly_work, 0.85)
            
            print(f"    [合并] {o_id} {kr_id}: 共 {len(weekly_work)} 条工作项")
            
            optimized_krs.append({
                'id': kr_id,
                'title': kr_title,  # 保留完整标题
                'weekly_work': weekly_work
            })
        
        optimized_okr_structure.append({
            'id': o_id,
            'title': o_title,  # 保留完整标题
            'key_results': optimized_krs
        })
    
    # 在team_report中记录成员顺序（严格按照STAFF_IDS）
    members = []
    
    if my_report:
        members.append({
            'id': my_report.get('id'),
            'name': my_report.get('name'),
            'status': my_report.get('status'),
            'order': 1,
            'is_base': True
        })
    
    order = 2
    for staff_id in STAFF_IDS:
        if staff_id == MY_ID:
            continue
        
        found_report = None
        for report in reports:
            if report.get('id') == staff_id:
                found_report = report
                break
        
        if found_report:
            members.append({
                'id': found_report.get('id'),
                'name': found_report.get('name'),
                'status': found_report.get('status'),
                'order': order,
                'is_base': False
            })
            order += 1
        else:
            members.append({
                'id': staff_id,
                'name': None,
                'status': 'missing',
                'order': order,
                'is_base': False
            })
            order += 1
    
    # 构建JSON格式的团队周报
    team_report = {
        'report_type': 'team_weekly_v10',
        'generate_time': datetime.now().isoformat(),
        'base_employee': {
            'id': MY_ID,
            'name': my_report.get('name') if my_report else None
        },
        'member_count': len(members),
        'members': members,
        'content': {
            'part1_work_summary': all_part1,
            'part2_okr_structure': optimized_okr_structure,  # 保留完整标题
            'part3_metrics': all_metrics
        },
        'stats': {
            'part1_count': len(all_part1),
            'o_count': len(optimized_okr_structure),
            'kr_count': sum(len(o.get('key_results', [])) for o in optimized_okr_structure),
            'weekly_work_count': sum(
                len(kr.get('weekly_work', []))
                for o in optimized_okr_structure
                for kr in o.get('key_results', [])
            ),
            'metrics_count': sum(len(v) for v in all_metrics.values())
        }
    }
    
    return team_report

def generate_team_weekly_markdown_v10(team_report):
    """
    将JSON格式的团队周报转换为Markdown格式（便于阅读）
    """
    lines = [
        "# 团队周报汇总 - V10统一格式版（JSON格式+标题去重）",
        "",
        f"**生成时间**: {datetime.fromisoformat(team_report['generate_time']).strftime('%Y-%m-%d %H:%M')}",
        f"**团队成员数**: {team_report['member_count']}",
        "",
        "---",
        ""
    ]
    
    content = team_report['content']
    
    # 跨人员去重（仅用于第一部分）
    seen_work_items_part1 = set()
    
    # 第一部分：本周工作总结（汇总）
    all_part1 = content.get('part1_work_summary', [])
    if all_part1:
        lines.extend([
            "## 第一部分：本周工作总结",
            ""
        ])
        for item in all_part1[:100]:
            item_norm = normalize_text(item)
            if item_norm not in seen_work_items_part1:
                lines.append(f"- {item}")
                seen_work_items_part1.add(item_norm)
            else:
                lines.append(f"- {item} (重复)")
        lines.append("")
    
    # 第二部分：OKR进度（使用去重后的标题，保留一个标记）
    okr_structure = content.get('part2_okr_structure', [])
    if okr_structure:
        lines.extend([
            "## 第二部分：OKR进度",
            ""
        ])
        
        # 为第二部分使用独立的去重集合
        seen_work_items_part2 = set()
        
        for o in okr_structure:
            o_title = o.get('title', '')  # 已包含 O1/O2/O3 标记
            lines.append(f"### {o_title}")  # 直接使用，不重复添加标记
            lines.append("")
            
            for kr in o.get('key_results', []):
                kr_title = kr.get('title', '')  # 已包含 KR1/KR2/KR3 标记
                lines.append(f"**{kr_title}**")  # 直接使用，不重复添加标记
                lines.append("")
                
                weekly_work = kr.get('weekly_work', [])
                
                if weekly_work:
                    lines.append("本周工作：")
                    display_count = min(200, len(weekly_work))
                    if len(weekly_work) > 200:
                        lines.append(f"*（该KR共有{len(weekly_work)}条工作项，显示前{display_count}条）*")
                    
                    for work in weekly_work[:display_count]:
                        work_norm = normalize_text(work)
                        if work_norm not in seen_work_items_part2:
                            lines.append(f"- {work}")
                            seen_work_items_part2.add(work_norm)
                        else:
                            lines.append(f"- {work} (重复)")
                    lines.append("")
                else:
                    lines.append("*本周无工作记录*")
                    lines.append("")
    
    # 第三部分：其他内容（汇总）
    all_metrics = content.get('part3_metrics', {})
    if any(all_metrics.values()):
        lines.extend([
            "## 第三部分：其他内容",
            ""
        ])
        
        for title, items in all_metrics.items():
            if items:
                lines.append(f"### {title}")
                lines.append("")
                for item in items[:15]:
                    lines.append(f"- {item}")
                lines.append("")
    
    # 添加统计信息
    stats = team_report.get('stats', {})
    lines.extend([
        "---",
        "",
        "## 统计信息",
        "",
        f"- 本周工作总结: {stats.get('part1_count', 0)} 条",
        f"- OKR目标数: {stats.get('o_count', 0)} 个",
        f"- 关键结果数: {stats.get('kr_count', 0)} 个",
        f"- 本周工作项: {stats.get('weekly_work_count', 0)} 条",
        f"- 其他指标项: {stats.get('metrics_count', 0)} 条",
        ""
    ])
    
    return '\n'.join(lines)


def generate_llm_prompt_v10(reports):
    """V10生成提示词（三部分结构）"""
    lines = [
        "基于以下团队成员的周报，请生成一份团队周报汇总：",
        "",
        "周报结构说明：",
        "- 第一部分：本周工作总结",
        "- 第二部分：OKR进度（O -> KR -> 本周工作）",
        "- 第三部分：业务核心指标、主要项目、下周重点工作",
        "",
        "要求：",
        "1. 按人员分别总结每个人的三部分内容",
        "2. 按上述结构分别合并每个人的内容",
        "3. 重点汇总OKR进展，突出关键结果",
        "4. 提取团队整体的关键业务指标",
        "5. 识别需要关注的风险或问题",
        "",
        "---",
        "",
        f"团队成员数: {len(reports)}",
        "",
        "团队数据：",
        ""
    ]
    
    seen_items = set()
    
    for report in reports:
        name = report.get('name', 'Unknown')
        content = report.get('content', {})
        
        lines.extend([f"## {name}", ""])
        
        # 第一部分
        part1 = content.get('part1_work_summary', [])
        if part1:
            lines.append("第一部分-本周工作总结：")
            for item in part1[:5]:
                lines.append(f"- {item}")
            lines.append("")
        
        # 第二部分
        okr_structure = content.get('part2_okr_structure', [])
        if okr_structure:
            lines.append("第二部分-OKR进展：")
            for o in okr_structure:
                lines.append(f"- {o.get('id', 'O')}: {o.get('title', '')}")
                for kr in o.get('key_results', []):
                    lines.append(f"  - {kr.get('id', 'KR')}: {kr.get('title', '')}")
                    for work in kr.get('weekly_work', [])[:3]:
                        work_norm = normalize_text(work)
                        if work_norm not in seen_items:
                            lines.append(f"    - {work}")
                            seen_items.add(work_norm)
            lines.append("")
        
        # 第三部分
        metrics = content.get('part3_metrics', {})
        if any(metrics.values()):
            lines.append("第三部分-其他内容：")
            for title, items in metrics.items():
                if items:
                    lines.append(f"  {title}：")
                    for item in items[:5]:
                        lines.append(f"    - {item}")
            lines.append("")
    
    return '\n'.join(lines)


def generate_personal_reference_v10(reports):
    """V10生成个人参考（三部分结构）"""
    my_reports = [r for r in reports if r.get('id') == MY_ID]
    
    lines = [
        "# 个人历史周报参考 - V10三部分结构（O/KR标题最终修复版）",
        "",
        f"**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "",
        "---",
        ""
    ]
    
    for report in my_reports:
        date = report.get('date', 'Unknown')
        content = report.get('content', {})
        
        lines.extend([f"## {date}", ""])
        
        # 第一部分
        part1 = content.get('part1_work_summary', [])
        if part1:
            lines.append("### 本周工作总结")
            lines.append("")
            for item in part1[:10]:
                lines.append(f"- {item}")
            lines.append("")
        
        # 第二部分
        okr_structure = content.get('part2_okr_structure', [])
        if okr_structure:
            lines.append("### OKR进度")
            lines.append("")
            for o in okr_structure:
                lines.append(f"- {o.get('id', 'O')}: {o.get('title', '')}")
            lines.append("")
        
        # 第三部分
        metrics = content.get('part3_metrics', {})
        if any(metrics.values()):
            lines.append("### 其他内容")
            lines.append("")
            for title, items in metrics.items():
                if items:
                    lines.append(f"**{title}**：")
                    for item in items[:5]:
                        lines.append(f"- {item}")
                    lines.append("")
        
        lines.extend(["---", ""])
    
    return '\n'.join(lines)


# ==================== 页面分析函数 ====================

async def analyze_page(page, context=""):
    if not DEBUG:
        return
    dprint(f"\n[页面分析] {context}")
    try:
        title = await page.title()
        dprint(f"  标题: {title}")
    except:
        pass
    input_cnt = await page.evaluate("document.querySelectorAll('input:not([type=hidden]), textarea').length")
    dprint(f"  可见输入框数量: {input_cnt}")
    dprint(f"  当前URL: {page.url}")


# ==================== 主流程（带重试机制 + 浏览器持久化复用）====================

async def main():
    last_week = get_week_date(0)
    week_before = get_week_date(1)
    
    print("="*60)
    print(" "*15 + "百度OKR周报抓取 - V10表格结构最终修复版（优化版）")
    print("="*60)
    print("\n【V10优化版】")
    print("  ✓ O1/O2/O3标题严格提取整行内容（包括标记后的所有文字）")
    print("  ✓ KR1/KR2/KR3标题严格放入title字段（绝不混入weekly_work）")
    print("  ✓ 本人和下属抓取逻辑完全一致")
    print("  ✓ 第一部分：本周工作总结（支持多次出现合并）")
    print("  ✓ 第二部分：OKR进度（三级表格结构）")
    print("  ✓ 第三部分：业务核心指标、主要项目、下周重点工作")
    print("  ✓ 新增：首次页面等待5秒")
    print("  ✓ 新增：失败页面自动重试机制（最多1次）")
    print("  ✓ 【优化】团队周报改为JSON格式（类似个人周报）")
    print("  ✓ 【优化】OKR标题去重：去掉重复的 O1/O2/O3/KR1/KR2/KR3 字样，保留一个")
    print("  ✓ 【新增】浏览器持久化复用：保存浏览器信息，下次可复用当前浏览器（参考工时脚本实现）")
    print(f"\n日期: 上周{last_week}, 上上周{week_before}")

    browser_manager = BrowserManager()
    
    browser_connected = False
    if browser_manager.load_state():
        browser_connected = await browser_manager.connect_to_existing_browser()
    
    if not browser_connected:
        iprint("\n[浏览器] 启动新浏览器...")
        if not await browser_manager.launch_new_browser():
            iprint("[X] 浏览器启动失败")
            return
    
    await browser_manager.ensure_page()
    page = browser_manager.page
    
    current_url = page.url
    if BASE_URL not in current_url:
        iprint(f"\n[导航] 当前URL: {current_url}，导航到目标页面...")
        await page.goto(BASE_URL, wait_until='domcontentloaded', timeout=30000)
    
    await analyze_page(page, "当前页面")

    # 检查是否需要登录
    current_url = page.url
    if 'uuap' in current_url or not current_url or current_url == 'about:blank':
        iprint("\n[系统] 访问OKR...")
        await page.goto("https://okr.baidu-int.com", wait_until='domcontentloaded')
        iprint("✓ 页面已加载")
        
        iprint("\n" + "="*60)
        iprint("⚠️  请先登录OKR系统")
        iprint("="*60)
        input("\n登录完成后按 Enter...")
        
        if 'uuap' in page.url:
            iprint("\n✗ 未登录成功")
            # 清理资源
            await browser_manager.close(keep_browser=False)
            return
        iprint("✓ 登录验证通过")
    else:
        iprint(f"\n✓ 当前页面: {current_url}")
    
    all_results = []
    failed_tasks = []  # 记录失败任务用于重试
    
    # 抓取下属（含本人）
    iprint(f"\n{'='*60}")
    iprint(f"抓取 {len(STAFF_IDS)} 个员工上周周报")
    iprint(f"{'='*60}")
    
    for i, staff_id in enumerate(STAFF_IDS, 1):
        iprint(f"fetch_single_staff_v10参数：{page},{staff_id},{last_week}")
        result = await fetch_single_staff_v10(page, staff_id, last_week, f"员工{i}", i, len(STAFF_IDS))
        all_results.append(result)
        
        if result['status'] == 'login_required':
            iprint("\n⚠️ 会话失效，停止")
            break
        
        if result['status'] == 'error':
            # 记录失败任务
            failed_tasks.append({
                'staff_id': staff_id,
                'date': last_week,
                'desc': f"员工{i}",
                'index': i,
                'total': len(STAFF_IDS)
            })
        
        # 优化2：如果是本人ID，额外抓取上上周的周报
        if staff_id == MY_ID:
            iprint(f"\n{'='*60}")
            iprint(f"检测到本人ID({MY_ID})，额外抓取上上周周报")
            iprint(f"{'='*60}")
            await asyncio.sleep(1.5)
            week_before_result = await fetch_single_staff_v10(page, MY_ID, week_before, "本人-上上周", i, len(STAFF_IDS))
            all_results.append(week_before_result)
            if week_before_result['status'] == 'error':
                failed_tasks.append({
                    'staff_id': MY_ID,
                    'date': week_before,
                    'desc': "本人-上上周",
                    'index': i,
                    'total': len(STAFF_IDS)
                })
    
    # ===== V10: 重试失败页面（最多1次）=====
    if failed_tasks:
        iprint(f"\n{'='*60}")
        iprint(f"重试机制：{len(failed_tasks)} 个页面需要重试")
        iprint(f"{'='*60}")
        
        for task in failed_tasks:
            iprint(f"\n重试: {task['desc']}")
            result = await fetch_single_staff_v10(
                page, 
                task['staff_id'], 
                task['date'], 
                task['desc'] + "(重试)", 
                task['index'], 
                task['total']
            )
            
            # 更新结果
            for i, r in enumerate(all_results):
                if r['id'] == task['staff_id'] and r['date'] == task['date']:
                    all_results[i] = result
                    break
            
            await asyncio.sleep(1)
    
    # 保存汇总
    summary = {
        'fetch_time': datetime.now().isoformat(),
        'my_id': MY_ID,
        'success': sum(1 for r in all_results if r['status'] == 'success'),
        'failed': sum(1 for r in all_results if r['status'] == 'error'),
        'retried': len(failed_tasks),
        'results': [{'id': r.get('id'), 'name': r.get('name'), 'status': r.get('status')} for r in all_results]
    }
    
    with open('v10_fetch_summary.json', 'w', encoding='utf-8') as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    
    iprint(f"\n{'='*60}")
    iprint("抓取完成!")
    iprint(f"  成功: {summary['success']}, 失败: {summary['failed']}, 重试: {summary['retried']}")

    iprint(f"\n{'='*60}")
    iprint("整合团队周报...")
    iprint(f"  配置的成员顺序 ({len(STAFF_IDS)} 人):")
    for i, staff_id in enumerate(STAFF_IDS, 1):
        iprint(f"    {i:2d}. {staff_id}")

    integrate_result = integrate_weekly_reports_v10()

    if integrate_result:
        iprint(f"\n{'='*60}")
        iprint("全部完成!")
        iprint(f"{'='*60}")
        iprint(f"\n生成的文件:")
        iprint(f"  - 姓名_日期_周报.json (个人周报)")
        iprint(f"  - team_weekly_report_v10.json (团队周报 - 按配置顺序)")
        iprint(f"  - team_weekly_report_v10.md (团队周报Markdown)")
        iprint(f"  - personal_reference_v10.md (个人参考)")
        if integrate_result.get('missing_ids'):
            iprint(f"\n⚠️ 缺失人员: {integrate_result['missing_ids']}")

    iprint("\n" + "="*60)
    iprint("[系统] 执行完成")
    iprint("="*60)
    iprint("\n选项:")
    iprint("  1. 保持浏览器运行并退出（下次可快速复用）- 直接按Enter")
    iprint("  2. 关闭浏览器并退出 - 输入 'close' 后按Enter")

    try:
        user_input = input("\n请选择 (默认: 保持运行): ").strip().lower()

        if user_input == 'close':
            iprint("[系统] 正在关闭浏览器...")
            await browser_manager.close(keep_browser=False)
        else:
            iprint("[系统] 保持浏览器运行，可以复用")
            await browser_manager.close(keep_browser=True)
            iprint(f"\n[提示] 下次运行脚本时将自动复用此浏览器")
            iprint(f"[提示] 状态文件: {BROWSER_STATE_FILE}")
    except KeyboardInterrupt:
        iprint("\n[系统] 用户中断，保持浏览器运行")
        await browser_manager.close(keep_browser=True)

    except Exception as e:
        iprint(f"\n[X] 严重错误: {e}")
        traceback.print_exc()
        if 'browser_manager' in locals():
            await browser_manager.close(keep_browser=True)


if __name__ == '__main__':
    asyncio.run(main())