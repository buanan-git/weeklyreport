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

【2026-03-14 更新】增加模块化导出函数，支持被主脚本直接调用
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
import sys
import random  # 新增：用于随机端口
from pathlib import Path
from datetime import datetime, timedelta
from playwright.async_api import async_playwright
from config_loader import ConfigLoader
# ==================== 新增：导入浏览器自动化模块 ====================
# 添加共享库路径
_libs_path = Path(__file__).parent.parent.parent / "libs"
if _libs_path.exists():
    sys.path.insert(0, str(_libs_path))

try:
    from browser_automation import (
        BrowserManager, HistoryManager, BaseElementFinder,
        highlight_element_with_index, clear_all_highlights,
        get_element_info, get_selector_for_element,
        analyze_page, set_debug as set_automation_debug
    )
    USE_AUTOMATION_MODULE = True
    print("[导入] 成功加载 browser_automation 通用模块", flush=True)
except ImportError as e:
    print(f"[警告] 无法导入 browser_automation 模块: {e}", flush=True)
    print("[警告] 将使用脚本内建的简化版本", flush=True)
    USE_AUTOMATION_MODULE = False

    # 定义简化的占位类，防止代码出错
    class BrowserManager:
        """浏览器管理器（精简备用版）"""
        def __init__(self, *args, **kwargs):
            self.browser = self.context = self.page = self.playwright = None
            self.cdp_port = self.browser_pid = self.browser_process = None
            self.headless = False
            self.state_file = BROWSER_STATE_FILE

        def save_state(self):
            if self.cdp_port and self.browser_pid:
                try:
                    with open(self.state_file, "w", encoding="utf-8") as f:
                        json.dump({"cdp_port": self.cdp_port, "browser_pid": self.browser_pid, "timestamp": time.time()}, f)
                except: pass

        def load_state(self):
            if os.path.exists(self.state_file):
                try:
                    with open(self.state_file, "r", encoding="utf-8") as f:
                        state = json.load(f)
                    if self._is_running(state.get("browser_pid")):
                        self.cdp_port, self.browser_pid = state.get("cdp_port"), state.get("browser_pid")
                        return True
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

        async def connect_to_existing(self):
            if not self.cdp_port: return False
            try:
                self.playwright = await async_playwright().start()
                self.browser = await self.playwright.chromium.connect_over_cdp(f"http://127.0.0.1:{self.cdp_port}")
                self.context = self.browser.contexts[0] if self.browser.contexts else await self.browser.new_context()
                self.page = self.context.pages[0] if self.context.pages else await self.context.new_page()
                return True
            except: return False

        async def launch_new(self, url=None):
            if not CHROME_PATH: return False
            self.cdp_port = random.randint(9222, 9299)
            cmd = [CHROME_PATH, f"--remote-debugging-port={self.cdp_port}", f"--user-data-dir={os.path.abspath(BROWSER_DEBUG_DIR)}", "--no-first-run", "--no-default-browser-check"]
            if self.headless: cmd.extend(["--headless=new"])
            if url: cmd.append(url)
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
                self.context = self.browser.contexts[0] if self.browser.contexts else await self.browser.new_context()
                self.page = self.context.pages[0] if self.context.pages else await self.context.new_page()
                self.save_state()
                return True
            except: return False

        async def ensure_page(self):
            if not self.browser: return False
            self.context = self.browser.contexts[0] if self.browser.contexts else await self.browser.new_context()
            self.page = self.context.pages[0] if self.context.pages else await self.context.new_page()
            return True

        async def close(self, keep_browser=True):
            if self.playwright:
                await self.playwright.stop()
                if not keep_browser:
                    if self.browser: await self.browser.close()
                    if self.browser_process: self.browser_process.terminate()
                    if os.path.exists(self.state_file): os.remove(self.state_file)

    class HistoryManager:
        def __init__(self, *args, **kwargs): pass
        def get_best_coord(self): return None
        def save(self, *args): pass
        
    class BaseElementFinder:
        def __init__(self, *args, **kwargs): pass
        async def find_by_history(self, *args, **kwargs): return None
        async def find_by_area(self, *args, **kwargs): return None
    
    async def highlight_element_with_index(*args, **kwargs): pass
    async def clear_all_highlights(*args, **kwargs): pass
    async def get_selector_for_element(*args, **kwargs): return "body"
    def set_automation_debug(*args): pass
    
# 加载配置
config = ConfigLoader("fetch_okr")

# ==================== 替换原有的硬编码配置 ====================
MY_ID = config.my_id
STAFF_IDS = config.staff_ids
TARGET_WEEK = config.target_week
DEBUG = config.debug

# 设置自动化模块的调试模式
set_automation_debug(DEBUG)

# 路径配置
BASE_DIR = config.script_dir
WEEKLYREPORT_DIR = config.get_path("weekly_report_dir")
BROWSER_STATE_FILE = config.get_browser_state_file()
BROWSER_DEBUG_DIR = config.get_path("browser_debug_dir")

# 常量
BASE_URL = 'https://okr.baidu-int.com/pages/weekly.html'


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

# 日期工具函数

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
    return f"{year}.{start_month}.{start_day}-{end_month}.{end_day}"

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
        r'的OKR',
        r'的OKR',
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
        r'Hi',
        r'同学，欢迎使用辅助区',
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
            # print(f"      [过滤] '{text[:50]}...' 匹配到无效模式: {pattern}")
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
                # else:
                    # print(f"      [过滤] 无效内容: {content_line[:80]}...")
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
                    # else:
                        # print(f"      [第三部分-{keyword}] 过滤: {content[:80]}...")
                    j += 1
                break
        
        result['part3_metrics'][keyword] = deduplicate_lines(part3_keywords[keyword], 0.85)
    
    return result

def extract_name_from_okr_header_enhanced(text):
    """增强版姓名提取函数"""
    if not text:
        return None
    
    # 清理文本
    text = text.replace('\ufeff', '').replace('﻿', '')
    
    # 打印调试信息
    print(f"[调试] 尝试从文本提取姓名: {text[:100]}")
    
    patterns = [
        r'([\u4e00-\u9fa5]{2,4})\s*的\s*OKR',
        r'([\u4e00-\u9fa5]{2,4})\s*OKR',
        r'汽车项目交付组[-–—]\s*([\u4e00-\u9fa5]{2,4})',
        r'[-–—]\s*([\u4e00-\u9fa5]{2,4})\s*$',  # 结尾的姓名
        r'^([\u4e00-\u9fa5]{2,4})'  # 开头的姓名
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            name = match.group(1)
            print(f"[调试] 通过模式 '{pattern}' 提取到姓名: {name}")
            return name
    
    # 如果都失败，尝试提取任何中文字符
    chinese_chars = re.findall(r'[\u4e00-\u9fa5]{2,4}', text)
    if chinese_chars:
        print(f"[调试] 通过中文字符提取到姓名: {chinese_chars[0]}")
        return chinese_chars[0]
    
    print("[调试] 未提取到姓名")
    return None
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
    """V10：确保内容完整加载（优化版：根据元素类型决定滚动方式）"""
    print("    滚动加载内容...", end=" ", flush=True)
    
    # 检查选择器是否可能是主要内容区域
    is_main_content = any(keyword in selector.lower() for keyword in ['content', 'weekly', 'main', 'report'])
    
    if is_main_content and selector != 'body':
        # 如果是主要内容区域，使用内部滚动
        try:
            await page.evaluate('''
                (selector) => {
                    const element = document.querySelector(selector);
                    if (element) {
                        const totalHeight = element.scrollHeight;
                        const clientHeight = element.clientHeight;
                        const steps = Math.ceil(totalHeight / clientHeight);
                        
                        const scrollStep = () => {
                            return new Promise(resolve => {
                                let currentStep = 0;
                                const interval = setInterval(() => {
                                    if (currentStep >= steps) {
                                        clearInterval(interval);
                                        resolve();
                                        return;
                                    }
                                    element.scrollTop = (currentStep + 1) * clientHeight;
                                    currentStep++;
                                }, 300);
                            });
                        };
                        
                        return scrollStep();
                    }
                }
            ''', selector)
        except:
            # 失败时回退到原有方法
            await page.mouse.move(center[0], center[1])
            for i in range(30):
                await page.mouse.wheel(0, 400)
                await asyncio.sleep(0.3)
    else:
        # 原有滚动方法
        await page.mouse.move(center[0], center[1])
        for i in range(30):
            await page.mouse.wheel(0, 400)
            await asyncio.sleep(0.3)
    
    print("✓")

async def extract_clean_content(page, selector):
    """
    从指定选择器提取干净的内容，排除所有UI元素
    返回纯净的文本内容
    """
    return await page.evaluate('''
        (selector) => {
            const element = document.querySelector(selector);
            if (!element) return '';
            
            // 需要排除的UI元素选择器（完整列表）
            const uiSelectors = [
                // 工具栏和按钮
                'button', '[role="button"]', '.semi-button', '.ds-icon-button',
                '[class*="toolbar"]', '[class*="Toolbar"]', '.lake-toolbar', '.editor-toolbar',
                '[class*="btn"]', '[class*="Btn"]', '[class*="button"]', '[class*="Button"]',
                
                // 菜单和侧边栏
                '[class*="menu"]', '[class*="Menu"]', '[role="menu"]',
                '[class*="sidebar"]', '[class*="Sidebar"]', 'aside', 'nav',
                
                // 操作区域
                '[class*="operation"]', '[class*="Operation"]',
                '[class*="action"]', '[class*="Action"]',
                
                // 图标和装饰
                'svg', '[class*="icon"]', '[class*="Icon"]', 'i[class*="icon"]',
                '[class*="emoji"]', '[class*="Emoji"]',
                
                // 分割线和装饰
                'hr', '[class*="divider"]', '[class*="Divider"]',
                '[class*="separator"]', '[class*="Separator"]',
                
                // 提示和帮助
                '[class*="tip"]', '[class*="Tip"]', '[class*="hint"]', '[class*="Hint"]',
                '[class*="help"]', '[class*="Help"]', '[class*="guide"]', '[class*="Guide"]',
                '[class*="welcome"]', '[class*="assistant"]',
                
                // 特定的UI文本
                '[class*="copy"]', '[class*="Copy"]',  // 复制按钮
                '[class*="share"]', '[class*="Share"]', // 分享按钮
                '[class*="more"]', '[class*="More"]',   // 更多按钮
                
                // 编辑器特有元素
                '.lake-*, [class*="lake-"]',
                '.ProseMirror-menubar', '.ProseMirror-tooltip',
                
                // 日期选择器等
                '[class*="date"]', '[class*="Date"]',
                '[class*="picker"]', '[class*="Picker"]',
                
                // 分页和导航
                '[class*="pagination"]', '[class*="Pagination"]',
                '[class*="breadcrumb"]', '[class*="Breadcrumb"]',
                
                // 任何包含这些文本的元素（通过属性）
                '[title*="复制"]', '[title*="Copy"]',
                '[aria-label*="复制"]', '[aria-label*="Copy"]',
                '[title*="关闭"]', '[title*="Close"]',
                '[aria-label*="关闭"]', '[aria-label*="Close"]'
            ];
            
            // 克隆节点
            const clone = element.cloneNode(true);
            
            // 移除所有UI元素
            uiSelectors.forEach(selector => {
                try {
                    const elements = clone.querySelectorAll(selector);
                    elements.forEach(el => el.remove());
                } catch(e) {}
            });
            
            // 获取文本
            let text = clone.innerText || '';
            
            // 清理：移除空行、特殊字符、多余空格
            text = text.split('\\n')
                .map(line => line.trim())
                .filter(line => {
                    // 过滤掉太短的行（可能是UI残留）
                    if (line.length < 2) return false;
                    // 过滤掉只包含特殊字符的行
                    if (/^[\\s\\-—―~・●○■□◆◇★☆♪♫♬©®™]+$/.test(line)) return false;
                    return true;
                })
                .join('\\n');
            
            return text;
        }
    ''', selector)

# ==================== 精确定位主要内容区域 ====================

async def find_main_content_area_exact(page):
    """
    精确定位周报内容区域 - 排除编辑器UI
    返回: 一个可靠的CSS选择器字符串
    """
    print("  精确定位内容区域...")
    
    # 1. 首先找到包含"本周工作总结"的标题元素
    try:
        # 查找"本周工作总结"标题
        summary_title = await page.query_selector("xpath=//*[contains(text(), '本周工作总结')]")
        if not summary_title:
            # 如果找不到，试试"OKR进度"
            summary_title = await page.query_selector("xpath=//*[contains(text(), 'OKR进度')]")
        
        if summary_title:
            # 找到标题后，向上查找真正的内容容器
            content_container = await page.evaluate('''(titleEl) => {
                try {
                    let current = titleEl;
                    let bestContainer = null;
                    let bestScore = 0;
                    
                    // 向上遍历最多8层，寻找最合适的容器
                    for (let i = 0; i < 8; i++) {
                        if (!current) break;
                        
                        const rect = current.getBoundingClientRect();
                        const text = current.innerText || '';
                        
                        // 检查这个容器是否包含典型的周报结构
                        const hasWeeklySummary = text.includes('本周工作总结');
                        const hasOKR = text.includes('OKR');
                        const hasOKRProgress = text.includes('OKR进度');
                        const hasMetrics = text.includes('业务核心指标') || text.includes('主要项目') || text.includes('下周重点工作');
                        
                        // 检查是否包含编辑器UI的关键词
                        const uiKeywords = ['复制', 'AI', '图片', '表格', '分栏', 'DuChat', 'CodePen', 
                                           '欢迎使用辅助区', '在此处拆分卡片', '布局', '视图', '图表',
                                           '阅读统计', '高亮信息', '流程图', '思维导图', '保密区',
                                           '默认分栏', '模块分栏', '编号分栏', '时间线分栏', '箭头分栏',
                                           '文本格式', '一级标题', '二级标题', '第三方应用'];
                        
                        let uiScore = 0;
                        for (const keyword of uiKeywords) {
                            if (text.includes(keyword)) {
                                uiScore += 5;
                            }
                        }
                        
                        // 计算内容得分
                        let contentScore = 0;
                        if (hasWeeklySummary) contentScore += 30;
                        if (hasOKR) contentScore += 20;
                        if (hasOKRProgress) contentScore += 20;
                        if (hasMetrics) contentScore += 15;
                        
                        // 检测OKR结构
                        const oMatches = (text.match(/O[1-3]\\b/g) || []).length;
                        const krMatches = (text.match(/KR[1-3]\\b/g) || []).length;
                        contentScore += oMatches * 10;
                        contentScore += krMatches * 5;
                        
                        // 检测实际工作内容
                        const workKeywords = ['完成', '推进', '上线', '交付', '支持', '开发', '测试', '优化'];
                        for (const kw of workKeywords) {
                            if (text.includes(kw)) contentScore += 1;
                        }
                        
                        // 净得分 = 内容得分 - UI得分
                        const netScore = contentScore - uiScore;
                        
                        // 如果净得分高，且容器足够大，说明是好的内容区域
                        if (rect.width > 600 && rect.height > 300 && netScore > 30) {
                            if (netScore > bestScore) {
                                bestScore = netScore;
                                bestContainer = current;
                            }
                        }
                        
                        current = current.parentElement;
                    }
                    
                    // 如果找到了好的容器，返回它
                    if (bestContainer) {
                        return bestContainer;
                    }
                    
                    // 如果没有找到，尝试在当前容器内查找可能的内容区域
                    current = titleEl;
                    for (let i = 0; i < 3; i++) {
                        if (!current) break;
                        
                        const contentCandidates = current.querySelectorAll(
                            'div[class*="content"], div[class*="main"], div[class*="body"], article, section'
                        );
                        
                        for (const candidate of contentCandidates) {
                            const candidateText = candidate.innerText || '';
                            const candidateRect = candidate.getBoundingClientRect();
                            
                            // 候选区域应该包含关键标题，且不包含太多UI元素
                            if (candidateRect.width > 500 && 
                                candidateRect.height > 200 &&
                                candidateText.includes('本周工作总结') &&
                                candidateText.includes('OKR')) {
                                
                                let uiCount = 0;
                                for (const keyword of ['复制', 'AI', '图片', '表格']) {
                                    if (candidateText.includes(keyword)) uiCount++;
                                }
                                
                                if (uiCount <= 1) {
                                    return candidate;
                                }
                            }
                        }
                        
                        current = current.parentElement;
                    }
                    
                    return null;
                } catch (e) {
                    return null;
                }
            }''', summary_title)
            
            if content_container:
                # 为容器生成选择器
                try:
                    # 尝试获取ID
                    element_id = await content_container.get_attribute('id')
                    if element_id:
                        selector = f"#{element_id}"
                        print(f"  找到精确内容区域 (ID): {selector}")
                        return selector
                    
                    # 尝试获取class
                    class_name = await content_container.get_attribute('class')
                    if class_name:
                        # 取第一个稳定的class
                        classes = class_name.split()
                        for cls in classes:
                            if cls and len(cls) > 2 and not any(c.isdigit() for c in cls[-4:]):
                                selector = f".{cls}"
                                print(f"  找到精确内容区域 (class): {selector}")
                                return selector
                    
                    # 如果都没有，返回xpath
                    print(f"  找到精确内容区域 (元素)")
                    return "xpath=//div[contains(@class, 'content')]"
                except:
                    pass
    except Exception as e:
        print(f"  精确查找异常: {e}")
    
    # 2. 尝试查找编辑器的实际内容区域
    try:
        content_selectors = [
            '.lake-content-editor-core',
            '.editor-content',
            '.ProseMirror',
            '[contenteditable="true"]',
            'div[class*="editor"] div[class*="content"]',
            'div[class*="lake"] div[class*="content"]',
            'div[class*="weekly"] div[class*="content"]',
        ]
        
        for selector in content_selectors:
            try:
                elements = await page.locator(selector).all()
                for elem in elements:
                    if not await elem.is_visible():
                        continue
                    
                    box = await elem.bounding_box()
                    if not box or box['width'] < 400:
                        continue
                    
                    text = await elem.text_content() or ''
                    # 确认这是周报内容区域
                    if '本周工作总结' in text and 'OKR' in text:
                        # 快速检查是否包含太多UI元素
                        ui_count = sum(1 for ui in ['复制', 'AI', '图片'] if ui in text)
                        if ui_count <= 1:
                            print(f"  找到编辑器内容区域: {selector}")
                            return selector
            except:
                continue
    except Exception as e:
        print(f"  编辑器查找失败: {e}")
    
    # 3. 如果都失败，回退到原有的智能定位
    print("  使用原有的智能定位作为后备")
    return await find_main_content_area_original(page)

async def find_main_content_area_with_history_exact(page, history_manager):
    """
    使用历史记录精确定位主要内容区域 - 优化版
    优先使用CSS选择器，历史坐标作为备选
    """
    if not USE_AUTOMATION_MODULE or not history_manager:
        return await find_main_content_area_exact(page)
    
    try:
        # ===== 第1步：先尝试快速CSS选择器（最高效）=====
        quick_selectors = [
            "[class*='content']",  # 从日志看这个最有效
            ".weekly-content",
            ".weekly-detail",
            ".report-content",
            ".main-content",
            "[class*='weekly']",
            "main",
            "article"
        ]
        
        for selector in quick_selectors:
            try:
                element = await page.query_selector(selector)
                if not element:
                    continue
                
                if not await element.is_visible():
                    continue
                
                box = await element.bounding_box()
                if not box or box['width'] < 400:
                    continue
                
                text = await element.text_content() or ''
                if '本周工作总结' in text and 'OKR' in text:
                    # 额外验证：确保文本中包含O和KR标记
                    if 'O1' in text or 'O2' in text or 'O3' in text:
                        print(f"  快速定位成功: {selector}")
                        # 保存成功坐标到历史
                        center_x = box['x'] + box['width']/2
                        center_y = box['y'] + box['height']/2
                        history_manager.save(center_x, center_y, {'method': 'quick_selector'})
                        return selector
            except:
                continue
        
        # ===== 第2步：尝试历史坐标精确查找（缩小搜索范围）=====
        best_coord = history_manager.get_best_coord_with_metadata()
        if best_coord:
            x, y = best_coord['x'], best_coord['y']
            metadata = best_coord.get('metadata', {})
            
            print(f"  [历史] 尝试坐标 ({x:.0f}, {y:.0f})")
            
            # 直接获取该位置的元素
            element = await page.evaluate_handle(f"document.elementFromPoint({x}, {y})")
            if element:
                # 快速验证
                is_valid, _ = await validate_main_content_quick(element)
                if is_valid:
                    print(f"  [历史] 坐标点直接命中")
                    box = await element.bounding_box()
                    selector = await get_selector_for_element_safe(page, element)
                    return selector
            
            # 如果直接点没命中，缩小搜索半径到20px
            nearby_element = await find_element_nearby(page, x, y, quick_selectors, radius=20)
            if nearby_element:
                print(f"  [历史] 附近20px内找到元素")
                box = await nearby_element.bounding_box()
                selector = await get_selector_for_element_safe(page, nearby_element)
                # 保存修正后的坐标
                center_x = box['x'] + box['width']/2
                center_y = box['y'] + box['height']/2
                history_manager.save(center_x, center_y, {'method': 'nearby_correction'})
                return selector
        
        # ===== 第3步：如果都失败，使用精确定位 =====
        print("  [历史] 未找到，使用精确定位")
        return await find_main_content_area_exact(page)
        
    except Exception as e:
        dprint(f"  [错误] find_main_content_area_with_history_exact 异常: {e}")
        return await find_main_content_area_exact(page)


async def validate_main_content_quick(elem):
    """快速验证函数 - 只检查关键特征"""
    try:
        if not elem:
            return (False, {})
        
        box = await elem.bounding_box()
        if not box or box['width'] < 400:
            return (False, {})
        
        text = await elem.text_content() or ''
        
        # 只检查最关键的特征
        has_weekly = '本周工作总结' in text
        has_okr = 'OKR' in text
        
        return (has_weekly and has_okr, {})
    except:
        return (False, {})


async def find_element_nearby(page, center_x, center_y, selectors, radius=20):
    """在指定点附近查找元素 - 缩小搜索范围"""
    try:
        for selector in selectors:
            elements = await page.locator(selector).all()
            for elem in elements:
                if not await elem.is_visible():
                    continue
                
                box = await elem.bounding_box()
                if not box:
                    continue
                
                elem_center_x = box['x'] + box['width']/2
                elem_center_y = box['y'] + box['height']/2
                
                # 计算距离
                distance = ((elem_center_x - center_x)**2 + (elem_center_y - center_y)**2)**0.5
                
                if distance < radius:
                    # 快速验证
                    is_valid, _ = await validate_main_content_quick(elem)
                    if is_valid:
                        return elem
        return None
    except:
        return None

async def find_main_content_area_original(page):
    """
    原有的定位函数 - 保留作为后备
    """
    # 1. 尝试预设的CSS选择器
    css_selectors = [
        ".weekly-content",
        ".weekly-detail", 
        ".report-content",
        ".main-content",
        "[class*='content']",
        "[class*='weekly']",
    ]
    
    for selector in css_selectors:
        try:
            element = await page.query_selector(selector)
            if element and await element.is_visible():
                text = await element.text_content()
                if text and ("本周工作总结" in text or "OKR" in text):
                    print(f"  通过CSS选择器定位成功: {selector}")
                    return selector
        except Exception:
            continue

    # 2. 基于内容特征的智能定位
    print("  CSS选择器失效，尝试基于内容特征定位...")
    try:
        title_element = await page.query_selector(
            "xpath=//*[contains(text(), '本周工作总结') or contains(text(), 'OKR进度')]"
        )
        
        if title_element:
            main_content_selector = await page.evaluate('''(titleElement) => {
                try {
                    const element = titleElement;
                    if (!element) return null;
                    
                    let current = element;
                    let candidate = element;
                    const visited = new Set();
                    
                    while (current && current.tagName !== 'BODY' && !visited.has(current)) {
                        visited.add(current);
                        
                        const rect = current.getBoundingClientRect();
                        const text = current.innerText || '';
                        
                        const hasMultipleTitles = (text.includes('本周工作总结') ? 1 : 0) +
                                                  (text.includes('OKR进度') ? 1 : 0) +
                                                  (text.includes('业务核心指标') ? 1 : 0) +
                                                  (text.includes('主要项目') ? 1 : 0) +
                                                  (text.includes('下周重点工作') ? 1 : 0) >= 2;
                        
                        if (rect.width > 500 && rect.height > 200 && 
                            (hasMultipleTitles || 
                             (text.includes('本周工作总结') && text.includes('OKR进度')))) {
                            candidate = current;
                        } else {
                            break;
                        }
                        
                        current = current.parentElement;
                    }
                    
                    if (candidate && candidate !== element) {
                        if (candidate.id && candidate.id.trim()) {
                            return `#${candidate.id}`;
                        }
                        if (candidate.classList && candidate.classList.length > 0) {
                            const classList = Array.from(candidate.classList);
                            for (const cls of classList) {
                                if (cls && typeof cls === 'string' && cls.length > 2 && !/\\d{4,}/.test(cls)) {
                                    return `[class*="${cls}"]`;
                                }
                            }
                        }
                    }
                    return null;
                } catch (e) {
                    return null;
                }
            }''', title_element)
            
            if main_content_selector:
                print(f"  通过内容特征定位成功: {main_content_selector}")
                return main_content_selector
    except Exception as e:
        print(f"  内容特征定位失败: {e}")
    
    # 3. 后备方案
    print("  使用后备方案: body")
    return "body"

# ==================== 添加缺失的 find_main_content_area 函数 ====================
async def find_main_content_area(page):
    """
    查找主要内容区域的通用入口函数
    根据可用函数智能选择最佳定位方法
    """
    try:
        # 首先尝试使用精确查找
        selector = await find_main_content_area_exact(page)
        if selector and selector != "body":
            return selector
        
        # 如果精确查找失败，使用原始方法
        return await find_main_content_area_original(page)
    except Exception as e:
        dprint(f"  [错误] find_main_content_area 异常: {e}")
        return "body"

# ==================== 新增：使用历史记录定位主要内容区域 ====================
async def find_main_content_area_with_history(page, history_manager):
    """
    使用历史记录定位主要内容区域
    增强版：添加了完整的错误处理和空值检查
    """
    if not USE_AUTOMATION_MODULE or not history_manager:
        return await find_main_content_area(page)
    
    try:
        # 创建元素查找器
        finder = BaseElementFinder(page, "OKR", "main_content", history_manager)
        
        # 定义验证函数
        async def validate_main_content(elem):
            try:
                if not elem:
                    return (False, {})
                
                # 检查元素是否可见
                is_visible = await elem.is_visible()
                if not is_visible:
                    return (False, {})
                
                box = await elem.bounding_box()
                if not box or box['width'] < 400 or box['height'] < 200:
                    return (False, {})
                
                text = await elem.text_content() or ''
                
                # 检查关键标题
                has_weekly = '本周工作总结' in text
                has_okr = 'OKR' in text
                has_metrics = any(k in text for k in ['业务核心指标', '主要项目', '下周重点工作'])
                
                # 检查是否包含UI元素（这些是排除条件）
                ui_keywords = ['复制', 'AI', '图片', '表格', '分栏', 'DuChat', 'CodePen', 
                               '欢迎使用辅助区', '在此处拆分卡片', '布局', '视图', '图表']
                has_ui = any(keyword in text for keyword in ui_keywords)
                
                # 如果包含UI元素，说明这不是纯内容区域
                if has_ui:
                    return (False, {})
                
                # 得分计算
                score = 0
                if has_weekly: score += 10
                if has_okr: score += 10
                if has_metrics: score += 5
                
                # 检查OKR结构
                o_count = len(re.findall(r'O[1-3]\b', text))
                kr_count = len(re.findall(r'KR[1-3]\b', text))
                score += o_count * 3 + kr_count * 2
                
                # 还需要检查是否包含典型的周报内容
                has_work_items = any(word in text for word in ['完成', '推进', '上线', '交付', '支持'])
                if has_work_items:
                    score += 5
                
                classes = await elem.get_attribute('class') or ''
                tag = await elem.evaluate('el => el.tagName.toLowerCase()')
                
                metadata = {
                    'score': score,
                    'has_weekly': has_weekly,
                    'has_okr': has_okr,
                    'has_metrics': has_metrics,
                    'o_count': o_count,
                    'kr_count': kr_count,
                    'classes': classes,
                    'tag': tag,
                    'box': {
                        'x': box['x'],
                        'y': box['y'],
                        'width': box['width'],
                        'height': box['height']
                    }
                }
                
                # 要求分数足够高，且包含OKR结构
                return (score > 20 and o_count >= 2, metadata)
                
            except Exception as e:
                dprint(f"    验证函数异常: {e}")
                return (False, {})       
                
        # 查找选择器
        selectors = [
            ".weekly-content", ".weekly-detail", ".report-content",
            ".main-content", "[class*='content']", "[class*='weekly']",
            "main", "article"
        ]
        
        # 尝试历史查找
        element = await finder.find_by_history(selectors, validate_main_content, search_radius=50)
        if element:
            iprint("  [历史] 成功定位主要内容区域")
            # 保存成功的历史记录
            try:
                box = await element.bounding_box()
                if box:
                    center_x = box['x'] + box['width']/2
                    center_y = box['y'] + box['height']/2
                    history_manager.save(center_x, center_y, {'method': 'history_success'})
                
                # 安全地获取选择器
                selector = await get_selector_for_element_safe(page, element)
                return selector
            except Exception as e:
                dprint(f"    保存历史记录失败: {e}")
                # 如果获取选择器失败，使用备选方法
                return await find_main_content_area(page)
        
        # 如果历史查找失败，使用原有的智能定位
        iprint("  [历史] 未找到，使用智能定位")
        selector = await find_main_content_area(page)
        
        # 保存智能定位的结果到历史
        try:
            element = await page.query_selector(selector)
            if element:
                box = await element.bounding_box()
                if box:
                    center_x = box['x'] + box['width']/2
                    center_y = box['y'] + box['height']/2
                    # 获取元素特征
                    classes = await element.get_attribute('class') or ''
                    text = await element.text_content() or ''
                    metadata = {
                        'classes': classes,
                        'has_weekly': '本周工作总结' in text,
                        'has_okr': 'OKR' in text,
                        'method': 'smart_fallback'
                    }
                    history_manager.save(center_x, center_y, metadata)
        except Exception as e:
            dprint(f"    保存智能定位历史失败: {e}")
        
        return selector
        
    except Exception as e:
        dprint(f"  [错误] find_main_content_area_with_history 异常: {e}")
        # 发生任何错误时回退到原始方法
        return await find_main_content_area(page)

async def get_selector_for_element_safe(page, element):
    """
    安全地获取元素选择器，带有完整的错误处理
    """
    if not element:
        return "body"
    
    try:
        # 尝试使用通用模块的函数
        if USE_AUTOMATION_MODULE:
            return await get_selector_for_element(page, element)
    except Exception as e:
        dprint(f"  [警告] 通用模块获取选择器失败: {e}")
    
    # 备选方案：手动构建简单选择器
    try:
        # 尝试获取ID
        element_id = await element.get_attribute('id')
        if element_id:
            return f"#{element_id}"
        
        # 尝试获取class
        class_name = await element.get_attribute('class')
        if class_name:
            # 取第一个class
            first_class = class_name.split()[0] if class_name.split() else None
            if first_class:
                return f".{first_class}"
        
        # 最后返回body
        return "body"
    except:
        return "body"

async def fetch_single_staff_v10(page, staff_id, date, desc, index=0, total=0, history_manager=None):
    """
    V10单个人员抓取（优化版：只提取主要内容区域）
    新增：使用历史记录管理器精确定位
    """
    url = f"{BASE_URL}#/index?mode=steam&id={staff_id}&date=s{date}"
    progress = f"[{index}/{total}] " if index else ""
    
    print(f"\n{'='*60}")
    print(f"{progress}抓取: {desc}")
    print(f"URL: {url}")
    
    try:
        print("  加载页面...", end=" ")
        
        try:
            await page.goto(url, wait_until='domcontentloaded', timeout=60000)
        except Exception as nav_error:
            current_url = page.url
            if 'uuap' in current_url:
                print("\n⚠️  会话失效，需要重新登录")
                return {'id': staff_id, 'date': date, 'status': 'login_required'}
            print(f"\n⚠️  导航警告: {nav_error}")
        
        if index == 1 or index == 0:
            print("(首次页面等待5秒)...", end=" ")
            await asyncio.sleep(5)
        else:
            await asyncio.sleep(2)
        print("✓")
        
        current_url = page.url
        if 'uuap' in current_url:
            print("需要登录")
            return {'id': staff_id, 'date': date, 'status': 'login_required'}
        
        # ===== 使用历史记录定位主要内容区域 =====
        print("  定位主要内容区域...", end=" ")
        try:
            if history_manager:
                main_content_selector = await find_main_content_area_with_history_exact(page, history_manager)
            else:
                main_content_selector = await find_main_content_area_exact(page)
            
            # 确保选择器不为空
            if not main_content_selector or main_content_selector == 'undefined':
                main_content_selector = "body"
                
            print(f"使用选择器: {main_content_selector}")
        except Exception as e:
            print(f"定位失败: {e}, 使用body")
            main_content_selector = "body"
            import traceback
            traceback.print_exc()
        
        # 滚动主要内容区域而不是整个页面
        print("  滚动加载内容...", end=" ")
        try:
            # 获取主要内容区域的位置
            element = await page.query_selector(main_content_selector)
            if element:
                # 滚动该元素内部
                await page.evaluate('''
                    (selector) => {
                        const element = document.querySelector(selector);
                        if (element) {
                            const scroll = () => {
                                element.scrollTop = element.scrollHeight;
                                setTimeout(() => {
                                    element.scrollTop = 0;
                                    setTimeout(() => {
                                        element.scrollTop = element.scrollHeight / 2;
                                    }, 300);
                                }, 300);
                            };
                            scroll();
                        }
                    }
                ''', main_content_selector)
                await asyncio.sleep(2)
                
                # 连续滚动
                for i in range(10):
                    await page.evaluate('''
                        (selector) => {
                            const element = document.querySelector(selector);
                            if (element) {
                                element.scrollBy(0, 200);
                            }
                        }
                    ''', main_content_selector)
                    await asyncio.sleep(0.3)
            else:
                # 备选：使用之前的滚动方法
                selector, center = await find_scroll_container(page)
                await scroll_container_v10(page, selector, center)
        except Exception as e:
            print(f"滚动失败: {e}, 使用备选方法")
            selector, center = await find_scroll_container(page)
            await scroll_container_v10(page, selector, center)
        print("✓")
        

        print("  提取主要内容...", end=" ")
        # 暂时恢复简单的 innerText 提取
        if main_content_selector != "body":
            full_text = await page.evaluate('''
                (selector) => {
                    const element = document.querySelector(selector);
                    if (element) {
                        return element.innerText;
                    }
                    return document.body.innerText;
                }
            ''', main_content_selector)
        else:
            full_text = await page.evaluate('() => document.body.innerText')
        
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
        
        json_file = os.path.join(WEEKLYREPORT_DIR, f'{safe_name}_{date}_周报.json')
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"  ✓ 完成: {safe_name} (已保存 {json_file})")
        
        return result
        
    except Exception as e:
        print(f"\n✗ 失败: {e}")
        import traceback
        traceback.print_exc()
        
        # 尝试截图保存错误现场 - 修改这里，不再使用 browser_manager
        try:
            # 使用 page 对象直接截图
            if page:
                screenshot_dir = Path(__file__).parent / "debug_screenshots"
                screenshot_dir.mkdir(exist_ok=True)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                screenshot_file = screenshot_dir / f"error_{staff_id}_{timestamp}.png"
                await page.screenshot(path=str(screenshot_file))
                print(f"  错误截图已保存: {screenshot_file}")
            else:
                print("  无法截图：page对象不存在")
        except Exception as screenshot_error:
            print(f"  截图失败: {screenshot_error}")
            
        return {'id': staff_id, 'date': date, 'status': 'error', 'error': str(e)}

# ==================== 整合功能（V10格式 - 优化版）====================
def integrate_weekly_reports_v10():

    targetweek = get_week_date(TARGET_WEEK)
    week_before = get_week_date(TARGET_WEEK+1)

    print("\n" + "="*60)
    print(f'整合周报...{targetweek}_周报.json')
    print(f'整合周报...{week_before}_周报.json')

    reports_dict = {}
    # 读取所有周报文件到字典，方便按ID查找
    # 定义需要查找的所有文件模式
    patterns = [
        f'WeeklyReport/*{targetweek}*_周报.json',
        f'WeeklyReport/*{week_before}*_周报.json'
    ]
    # 遍历所有匹配模式的文件
    for pattern in patterns:
        for file in glob.glob(pattern):
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
    week_suffix = get_week_filename_suffix(TARGET_WEEK)
    json_file = os.path.join(WEEKLYREPORT_DIR,f'团队周报_{week_suffix}.json')
    md_file = os.path.join(WEEKLYREPORT_DIR,f'团队周报_{week_suffix}.md')
    
    # 保存JSON文件
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(team_report, f, ensure_ascii=False, indent=2)
    print(f"  ✓ 团队周报(JSON): {json_file}")
    
    # 生成Markdown版本
    team_md = generate_team_weekly_markdown_v10(team_report)
    with open(md_file, 'w', encoding='utf-8') as f:
        f.write(team_md)
    print(f"  ✓ 团队周报(Markdown): {md_file}")
    
    
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
    
    # 去重第一部分和第三部分（删除filter_hi_content调用）
    all_part1 = deduplicate_lines(all_part1, 0.85)
    for key in all_metrics.keys():
        all_metrics[key] = deduplicate_lines(all_metrics[key], 0.85)
    
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
            # 删除filter_hi_content调用
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


# ==================== 新增：模块化导出函数 ====================

async def fetch_okr_data(weeks_ago=0):
    """
    【新增】核心导出函数 - 获取团队周报数据并返回数据对象
    :param weeks_ago: 0=当周，1=上一周，2=上上周
    :return: 团队周报数据字典（与integrate_weekly_reports_v10返回的team_json相同）
    """
    global TARGET_WEEK
    TARGET_WEEK = weeks_ago
    
    # 创建浏览器管理器
    iprint(f"[浏览器] 外部模块加载状态: {USE_AUTOMATION_MODULE}")
    iprint(f"[浏览器] 状态文件: {BROWSER_STATE_FILE}")
    iprint(f"[浏览器] 状态文件存在: {os.path.exists(str(BROWSER_STATE_FILE))}")

    # 传递 config 和 state_file 参数，确保状态文件被正确保存
    if USE_AUTOMATION_MODULE:
        browser_manager = BrowserManager(
            config=config,
            chrome_path=config.get_chrome_path() if hasattr(config, 'get_chrome_path') else None,
            debug_dir=str(BROWSER_DEBUG_DIR),
            headless=config.should_use_headless if hasattr(config, 'should_use_headless') else False,
            screenshot_on_error=config.screenshot_on_error if hasattr(config, 'screenshot_on_error') else True,
            state_file=str(BROWSER_STATE_FILE)
        )
    else:
        # 使用内置精简版
        browser_manager = BrowserManager()

    # 尝试连接已有浏览器或启动新浏览器
    browser_connected = False
    if browser_manager.load_state():
        iprint(f"[浏览器] 加载状态成功: 端口={browser_manager.cdp_port}, PID={browser_manager.browser_pid}")
        browser_connected = await browser_manager.connect_to_existing()
        if not browser_connected:
            iprint("[浏览器] 连接已存在的浏览器失败")
    else:
        iprint("[浏览器] 加载状态失败或浏览器进程不存在")

    if not browser_connected:
        iprint("\n[浏览器] 启动新浏览器...")
        if not await browser_manager.launch_new():
            iprint("[X] 浏览器启动失败")
            return None

    iprint("[浏览器] 等待浏览器加载完成...")
    await asyncio.sleep(5)
    
    # 确保页面实例存在
    await browser_manager.ensure_page()
    page = browser_manager.page

    # 获取当前URL并判断是否需要初始导航
    current_url = page.url
    if BASE_URL not in current_url:
        iprint(f"\n[导航] 当前URL: {current_url}，导航到目标页面...")
        try:
            await page.goto(BASE_URL, wait_until='domcontentloaded', timeout=30000)
        except asyncio.TimeoutError:
            iprint("[!] 页面加载超时，重试一次...")
            await page.goto(BASE_URL, wait_until='domcontentloaded', timeout=30000)
        iprint("[浏览器] 等待浏览器加载完成...")
        await asyncio.sleep(5)
        current_url = page.url

    # 分析页面并检查登录状态
    await analyze_page(page, "当前页面")
    need_login = False
    if (current_url and 'uuap' in current_url) or current_url in ('', 'about:blank'):
        need_login = True

    # 引导用户登录
    if need_login:
        iprint("\n" + "="*60)
        iprint("⚠️  请在浏览器中完成OKR系统登录")
        iprint("   登录完成后请确保停留在目标页面再按Enter")
        iprint("="*60)
        await asyncio.get_event_loop().run_in_executor(None, input, "\n登录完成后按 Enter...")
        
        iprint("\n[导航] 验证登录状态...")
        await page.goto(BASE_URL, wait_until='domcontentloaded', timeout=30000)
        current_url = page.url

    # 登录状态最终验证
    await analyze_page(page, "登录后页面")
    if 'uuap' in current_url:
        iprint("\n✗ 未登录成功，请检查账号密码或网络")
        await browser_manager.close(keep_browser=False)
        return None
    
    # 抓取数据
    targetweek = get_week_date(TARGET_WEEK)
    week_before = get_week_date(TARGET_WEEK+1)
    
    all_results = []
    failed_tasks = []

    iprint(f"\n{'='*60}")
    iprint(f"抓取 {len(STAFF_IDS)} 个员工目标周的周报")
    iprint(f"{'='*60}")

    for i, staff_id in enumerate(STAFF_IDS, 1):
        if staff_id == MY_ID:
            iprint(f"抓取本人 {staff_id} 目标周的上周周报")
            result = await fetch_single_staff_v10(page, staff_id, week_before, f"员工{i}", i, len(STAFF_IDS))
        else:
            result = await fetch_single_staff_v10(page, staff_id, targetweek, f"员工{i}", i, len(STAFF_IDS))
        all_results.append(result)
        
        if result['status'] == 'login_required':
            iprint("\n⚠️ 会话失效，停止")
            break
        
        if result['status'] == 'error':
            failed_tasks.append({
                'staff_id': staff_id,
                'date': targetweek,
                'desc': f"员工{i}",
                'index': i,
                'total': len(STAFF_IDS)
            })
    
    # 重试失败页面
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
            
            for i, r in enumerate(all_results):
                if r['id'] == task['staff_id'] and r['date'] == task['date']:
                    all_results[i] = result
                    break
            
            await asyncio.sleep(1)
    
    # 整合团队周报
    iprint(f"\n{'='*60}")
    iprint("整合团队周报...")
    
    integrate_result = integrate_weekly_reports_v10()
    
    # 关闭浏览器（保持运行）
    await browser_manager.close(keep_browser=True)
    
    if integrate_result:
        return integrate_result.get('team_json')
    else:
        return None

def save_okr_to_file(data, weeks_ago=0):
    """
    【新增】保存团队周报数据到文件
    :param data: 团队周报数据字典
    :param weeks_ago: 周数
    :return: 保存的文件路径
    """
    week_suffix = get_week_filename_suffix(weeks_ago)
    json_file = os.path.join(WEEKLYREPORT_DIR, f'团队周报_{week_suffix}.json')
    
    # 确保目录存在
    os.makedirs(WEEKLYREPORT_DIR, exist_ok=True)
    
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    iprint(f"团队周报已保存到: {json_file}")
    return json_file


# ==================== 主流程（兼容原有命令行调用）====================

async def main_cli():
    """命令行版本的主函数
    Args:
        target_week: 目标周数（0=本周，1=上周，2=上上周）
        llm_platform: 大模型平台（BS/DS/DB/KM）
    """
    global TARGET_WEEK, LLM_PLATFORM
    
    # 1. 确定目标周：优先级 命令行参数 > 配置文件
    if len(sys.argv) > 1:
        # 命令行参数优先级最高
        try:
            TARGET_WEEK = int(sys.argv[1])
            print(f"📋 使用命令行参数: 周数={TARGET_WEEK}")
        except ValueError:
            print(f"❌ 错误：参数「{sys.argv[1]}」不是有效数字！")
            return
    else:
        # 使用配置文件默认值
        TARGET_WEEK = config.target_week
        print(f"📋 使用配置文件: 周数={TARGET_WEEK}")
    
    # 2. 确定平台：优先级 命令行参数 > 配置文件
    if len(sys.argv) > 2:
        # 命令行参数优先级最高
        platform_arg = sys.argv[2].upper()
        if platform_arg in ["BS", "DS", "DB", "KM"]:
            LLM_PLATFORM = platform_arg
            print(f"📋 使用命令行参数: 平台={LLM_PLATFORM}")
    else:
        # 使用配置文件默认值
        LLM_PLATFORM = config.default_platform
        print(f"📋 使用配置文件: 平台={LLM_PLATFORM}")
    
    # 显示最终使用的参数
    print(f"\n🚀 最终运行参数:")
    print(f"   - 目标周: {TARGET_WEEK} ({'本周' if TARGET_WEEK==0 else '上周' if TARGET_WEEK==1 else '上上周'})")
    print(f"   - 平台: {LLM_PLATFORM}")
    print("="*60 + "\n")
    
    targetweek = get_week_date(TARGET_WEEK)
    week_before = get_week_date(TARGET_WEEK+1)
    
    print("="*60)
    print(" "*15 + "百度OKR周报抓取")
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
    print(f"\n日期: 获取目标周{targetweek}, 获取目标周的上一周{week_before}")

    if USE_AUTOMATION_MODULE:
        # 使用通用模块的 BrowserManager
        browser_manager = BrowserManager(
            config=config,
            headless=False,  # 根据需要设置
            screenshot_on_error=True
        )
    else:
        # 回退到原有的 BrowserManager（需要保留原有定义）
        browser_manager = BrowserManager()  # 原有的类
    
    browser_connected = False
    if browser_manager.load_state():
        browser_connected = await browser_manager.connect_to_existing()
    
    if not browser_connected:
        iprint("\n[浏览器] 启动新浏览器...")
        if not await browser_manager.launch_new(BASE_URL):
            iprint("[X] 浏览器启动失败")
            return

    iprint("[浏览器] 等待浏览器加载完成...")
    await asyncio.sleep(5)
    
    # 2.确保页面实例存在
    await browser_manager.ensure_page()
    page = browser_manager.page

    # ===== 新增：初始化主要内容区域的历史管理器 =====
    content_history = None
    if USE_AUTOMATION_MODULE:
        content_history = HistoryManager(
            name="main_content", 
            platform="OKR",
            config=config,
            base_dir=BASE_DIR
        )
        iprint(f"[历史] 已初始化主要内容区域历史记录")

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
        iprint("[浏览器] 等待浏览器加载完成...")
        await asyncio.sleep(5)
        current_url = page.url  # 更新URL

    # 4. 分析页面并检查登录状态
    await analyze_page(page, "2当前页面")
    need_login = False
    # 更严谨的登录判断逻辑
    if (current_url and 'uuap' in current_url) or current_url in ('', 'about:blank'):
        need_login = True

    # 5. 引导用户登录
    if need_login:
        iprint("\n" + "="*60)
        iprint("⚠️  请在浏览器中完成OKR系统登录")
        iprint("   登录完成后请确保停留在目标页面再按Enter")
        iprint("="*60)
        # 使用input的异步兼容写法（避免阻塞事件循环）
        await asyncio.get_event_loop().run_in_executor(None, input, "\n登录完成后按 Enter...")
        
        # 仅当用户手动登录后，才重新导航（避免冗余跳转）
        iprint("\n[导航] 验证登录状态...")
        await page.goto(BASE_URL, wait_until='domcontentloaded', timeout=30000)
        current_url = page.url

    # 6. 登录状态最终验证
    await analyze_page(page, "登录后页面")
    if 'uuap' in current_url:
        iprint("\n✗ 未登录成功，请检查账号密码或网络")
        await browser_manager.close(keep_browser=False)
        return
    
    all_results = []
    failed_tasks = []  # 记录失败任务用于重试
    
    # 抓取下属（含本人）
    iprint(f"\n{'='*60}")
    iprint(f"抓取 {len(STAFF_IDS)} 个员工目标周的周报")
    iprint(f"{'='*60}")

    for i, staff_id in enumerate(STAFF_IDS, 1):
        iprint(f"fetch_single_staff_v10参数：{page},{staff_id},{targetweek}")
        if staff_id == MY_ID:        # 如果是本人ID，抓取目标周上周的周报
            iprint(f"抓取本人 {staff_id} 目标周的上周周报")
            result = await fetch_single_staff_v10(
                page, staff_id, week_before, f"员工{i}", 
                i, len(STAFF_IDS), history_manager=content_history
            )
        else:
            result = await fetch_single_staff_v10(
                page, staff_id, targetweek, f"员工{i}", 
                i, len(STAFF_IDS), history_manager=content_history
            )
        all_results.append(result)
        
        if result['status'] == 'login_required':
            iprint("\n⚠️ 会话失效，停止")
            break
        
        if result['status'] == 'error':
            # 记录失败任务
            failed_tasks.append({
                'staff_id': staff_id,
                'date': targetweek,
                'desc': f"员工{i}",
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
    
    #保存汇总
    summary = {
        'fetch_time': datetime.now().isoformat(),
        'my_id': MY_ID,
        'success': sum(1 for r in all_results if r['status'] == 'success'),
        'failed': sum(1 for r in all_results if r['status'] == 'error'),
        'retried': len(failed_tasks),
        'results': [{'id': r.get('id'), 'name': r.get('name'), 'status': r.get('status')} for r in all_results]
    }
    
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
        iprint(f"  - (团队周报 -json)")
        iprint(f"  - (团队周报 -Markdown)")
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
    try:
        asyncio.run(main_cli())
    except Exception as e:
        print(f"错误: {e}")
        traceback.print_exc()
    finally:
        print("\n按回车键退出...")
        input()  # 等待用户按回车