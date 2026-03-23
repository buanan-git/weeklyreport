#!/usr/bin/env python3
"""
百度工时填报自动脚本 - 精确提取版
目标URL: https://eop.baidu-int.com/product/weekly/working
功能：
1. 自动切换到上个月，遍历所有周
2. 对于每一周，遍历所有项目行的每个工作日单元格
3. 点击单元格，从右侧详情区域读取：
   - 工时code及中文名（如 acxiujia (休假)）
   - 日期（如 2026-01-26（星期一））
   - 根据code类型提取不同字段：
       管理类（acguanli, acpeixun, acxiujia）：
         - 合计工时(H)
         - 备注（解析为：【】内内容、工作内容、备注）
       项目类（AC12345, SQ12345）：
         - 工时阶段
         - 工时角色
         - 工时(H)
         - 产品(L3)工时
         - 备注
   - 单元格内的工时数值
4. 所有数据以JSON格式保存

使用前请修改下方配置区域：
    DRY_RUN: 预览模式，只读取不填写。
    AUTO_SUBMIT: 是否自动点击保存/提交按钮。
    DEBUG: 输出详细调试信息。
"""

import asyncio
import socket
import json
import re
from datetime import datetime, timedelta, date
from playwright.async_api import async_playwright

# ==================== 配置区域 ====================

DRY_RUN = True               # 预览模式：只读取不填写
AUTO_SUBMIT = True           # 是否自动保存/提交
DEBUG = True                  # 输出详细调试信息

BASE_URL = "https://eop.baidu-int.com/product/weekly/working"

# 管理类工时code
MANAGEMENT_CODES = ["acguanli", "acpeixun", "acxiujia"]

# 有效的工时code模式
VALID_CODE_PATTERNS = [
    r'acguanli',
    r'acpeixun',
    r'acxiujia',
    r'AC\d{5}',
    r'SQ\d{5}'
]

# ==================== 工具函数 ====================

def dprint(*args, **kwargs):
    """调试输出，自动flush"""
    if DEBUG:
        print(*args, **kwargs, flush=True)

def iprint(*args, **kwargs):
    """普通信息输出，强制flush"""
    print(*args, **kwargs, flush=True)

def is_valid_code(text: str) -> bool:
    """判断文本是否包含有效的工时code"""
    if not text:
        return False
    text_lower = text.lower()
    for pattern in VALID_CODE_PATTERNS:
        if re.search(pattern, text_lower):
            return True
    return False

def is_management_code(code_str: str) -> bool:
    """判断是否为管理类工时code"""
    if not code_str:
        return False
    code_lower = code_str.lower()
    return any(mgmt in code_lower for mgmt in MANAGEMENT_CODES)

def is_project_code(code_str: str) -> bool:
    """判断是否为项目类工时code (AC+5位数字 或 SQ+5位数字)"""
    if not code_str:
        return False
    pattern = r'^(AC|SQ)\d{5}$'
    return bool(re.match(pattern, code_str.upper()))

def extract_code_and_name(text: str) -> dict:
    """
    从文本中提取code和中文名称
    例如: "acxiujia (休假)" -> {"code": "acxiujia", "name": "休假"}
    """
    result = {"code": None, "name": None}
    
    if not text or not is_valid_code(text):
        return result
    
    # 清理文本
    text = re.sub(r'\s+', ' ', text).strip()
    
    # 匹配 "code (中文名)"
    pattern = r'([a-zA-Z]+\d*)\s*\(([^)]+)\)'
    match = re.search(pattern, text)
    if match:
        result["code"] = match.group(1).strip().lower()
        result["name"] = match.group(2).strip()
        return result
    
    # 如果没有括号，只提取code
    for pattern in VALID_CODE_PATTERNS:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            result["code"] = match.group(0).strip().lower()
            break
    
    return result

def extract_date(text: str) -> str:
    """
    从文本中提取日期
    例如: "日期：2026-01-26（星期一）" -> "2026-01-26（星期一）"
    """
    if not text:
        return None
    
    # 匹配 "日期：" 后面的完整日期
    match = re.search(r'日期[：:]\s*(\d{4}-\d{2}-\d{2}（[^）]+）)', text)
    if match:
        return match.group(1).strip()
    
    # 匹配 "日期：" 后面的简单日期
    match = re.search(r'日期[：:]\s*(\d{4}-\d{2}-\d{2})', text)
    if match:
        return match.group(1).strip()
    
    # 直接匹配日期
    match = re.search(r'(\d{4}-\d{2}-\d{2}（[^）]+）)', text)
    if match:
        return match.group(1).strip()
    
    match = re.search(r'(\d{4}-\d{2}-\d{2})', text)
    if match:
        return match.group(1).strip()
    
    return None

def parse_management_remark(remark_text: str) -> dict:
    """
    解析管理类工时的备注文本
    格式: 【第一部分】；第二部分；备注内容
    返回: {"part1": "...", "work_content": "...", "remark": "..."}
    """
    result = {
        "part1": None,
        "work_content": None,
        "remark": None
    }
    
    if not remark_text or not remark_text.strip():
        return result
    
    # 匹配【】内的内容
    bracket_match = re.search(r'【([^】]+)】', remark_text)
    if bracket_match:
        result["part1"] = bracket_match.group(1).strip()
        remaining = remark_text.replace(f'【{result["part1"]}】', '').strip()
    else:
        remaining = remark_text
    
    # 按分号分割
    parts = [p.strip() for p in remaining.split('；') if p.strip()]
    
    if len(parts) >= 1:
        result["work_content"] = parts[0]
    if len(parts) >= 2:
        result["remark"] = parts[1]
    
    return result

# ==================== 日期工具 ====================

def get_current_month_range():
    today = date.today()
    first_day = today.replace(day=1)
    if first_day.month == 12:
        next_month_first = first_day.replace(year=first_day.year+1, month=1, day=1)
    else:
        next_month_first = first_day.replace(month=first_day.month+1, day=1)
    last_day = next_month_first - timedelta(days=1)
    return first_day, last_day

def get_last_month_range():
    today = date.today()
    first_day_this = today.replace(day=1)
    last_day_last = first_day_this - timedelta(days=1)
    first_day_last = last_day_last.replace(day=1)
    return first_day_last, last_day_last

def format_month_param(dt: date) -> str:
    return dt.strftime("%Y-%m")

def get_target_url(month_param: str = None) -> str:
    if month_param:
        return f"{BASE_URL}?month={month_param}"
    return BASE_URL

# ==================== 网络检测 ====================

def check_network():
    iprint("\n[网络检测] 检查内网连接...")
    try:
        socket.gethostbyname("eop.baidu-int.com")
        iprint("  [OK] 内网域名解析正常")
        return True
    except:
        iprint("  [X] 内网域名解析失败，请检查VPN")
        return False

# ==================== 页面辅助函数 ====================

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

async def find_scroll_container(page):
    containers = await page.evaluate("""() => {
        const all = document.querySelectorAll("*");
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
    }""")
    if containers:
        first = containers[0]
        if first.get("class"):
            selector = f"[class*='{first['class'].split()[0]}']"
        elif first.get("id"):
            selector = f"#{first['id']}"
        else:
            selector = first["tag"].lower()
        return selector, (first.get("x", 960), first.get("y", 600))
    return "body", (960, 600)

async def scroll_page(page):
    iprint("  滚动加载...", end=" ")
    selector, center = await find_scroll_container(page)
    await page.mouse.move(center[0], center[1])
    await asyncio.sleep(0.5)
    for i in range(20):
        await page.mouse.wheel(0, 300)
        await asyncio.sleep(0.2)
    iprint("OK")
    await asyncio.sleep(2)

# ==================== 月份切换 ====================

async def switch_to_month(page, target_month: str):
    iprint(f"\n[月份切换] 切换到 {target_month}")
    try:
        month_picker = page.locator('.ant-calendar-picker, .month-picker, [class*="month"]:has-text("/")').first
        if await month_picker.count() > 0 and await month_picker.is_visible():
            await month_picker.click()
            await asyncio.sleep(1)
            target_opt = page.get_by_text(target_month.replace('-', '/'), exact=False).first
            if await target_opt.is_visible():
                await target_opt.click()
                iprint(f"  [OK] 点击月份选择器选择 {target_month}")
                await asyncio.sleep(2)
                return True
            else:
                iprint(f"  [!] 未找到目标月份选项: {target_month}")
        else:
            iprint("  [!] 未找到月份选择器")
    except Exception as e:
        iprint(f"  [!] 月份选择器点击失败: {e}")

    try:
        url = get_target_url(target_month)
        iprint(f"  [尝试] 直接导航到 {url}")
        await page.goto(url, wait_until='domcontentloaded', timeout=15000)
        iprint(f"  [OK] 通过URL导航到 {target_month}")
        await asyncio.sleep(2)
        return True
    except Exception as e:
        iprint(f"  [!] URL导航失败: {e}")

    iprint("  [X] 无法切换月份，请手动检查页面")
    return False

# ==================== 周切换 ====================

async def switch_to_week(page, week_label):
    dprint(f"  [周切换] 尝试切换到: {week_label}")
    try:
        week_locator = page.get_by_text(week_label, exact=False).first
        if await week_locator.is_visible(timeout=3000):
            await week_locator.click()
            dprint(f"    [OK] 点击周标签: {week_label}")
            await asyncio.sleep(2)
            return True
    except Exception as e:
        dprint(f"    [!] 点击失败: {e}")

    if "（本周）" in week_label:
        base = week_label.replace("（本周）", "")
        try:
            week_locator = page.get_by_text(base, exact=False).first
            if await week_locator.is_visible():
                await week_locator.click()
                dprint(f"    [OK] 点击周标签(去后缀): {base}")
                await asyncio.sleep(2)
                return True
        except Exception as e:
            dprint(f"    [!] 点击失败: {e}")

    dprint(f"    [X] 无法找到周标签: {week_label}")
    return False

# ==================== 获取表格结构 ====================

async def get_table_structure(page):
    structure = {
        "projects": [],
        "day_headers": []
    }
    
    try:
        await page.wait_for_selector("table", timeout=10000)
    except:
        dprint("    [X] 未找到表格")
        return structure
    
    # 获取表头（星期几）
    structure["day_headers"] = await page.evaluate("""() => {
        const headers = [];
        const headerRow = document.querySelector('thead tr, .ant-table-thead tr');
        if (headerRow) {
            const cells = headerRow.querySelectorAll('th, td');
            for (let cell of cells) {
                let text = cell.textContent.trim();
                if (text.includes('星期一')) headers.push('周一');
                else if (text.includes('星期二')) headers.push('周二');
                else if (text.includes('星期三')) headers.push('周三');
                else if (text.includes('星期四')) headers.push('周四');
                else if (text.includes('星期五')) headers.push('周五');
                else if (text.includes('星期六')) headers.push('周六');
                else if (text.includes('星期日')) headers.push('周日');
            }
        }
        return headers;
    }""")
    
    # 获取所有项目行
    structure["projects"] = await page.evaluate("""() => {
        const projects = [];
        const tbody = document.querySelector('tbody');
        if (tbody) {
            const rows = tbody.querySelectorAll('tr');
            for (let row of rows) {
                if (row.style.display === 'none') continue;
                const cells = row.querySelectorAll('td');
                if (cells.length < 2) continue;
                const projectCell = cells[0];
                let projectName = projectCell.textContent.trim();
                if (projectName && !projectName.includes('合计') && !projectName.includes('总计')) {
                    projects.push(projectName);
                }
            }
        }
        return projects;
    }""")
    
    dprint(f"      表格结构: {len(structure['projects'])}个项目, {len(structure['day_headers'])}天")
    return structure

# ==================== 核心：点击单元格并读取右侧信息 ====================

async def click_cell_and_read_info(page, row_index, col_index):
    """
    点击单元格，等待右侧更新，读取所有信息
    """
    dprint(f"      [点击] 行{row_index+1}, 列{col_index+2}")
    
    # 定位单元格
    try:
        cell = page.locator(f'tbody tr:nth-child({row_index+1}) td:nth-child({col_index+2})').first
        if await cell.count() == 0 or not await cell.is_visible():
            dprint(f"      [X] 找不到单元格")
            return None
        
        # 获取单元格内的工时数值
        hour_value = None
        try:
            input_in_cell = cell.locator('input[type="number"]').first
            if await input_in_cell.count() > 0:
                val = await input_in_cell.input_value()
                hour_value = float(val) if val else 0
        except:
            pass
        
        await cell.click()
        await asyncio.sleep(1.5)  # 等待右侧更新
    except Exception as e:
        dprint(f"      [X] 点击单元格失败: {e}")
        return None
    
    # 读取右侧信息
    info = await read_right_panel_info(page)
    if info:
        info["hour_value"] = hour_value
    
    return info

async def read_right_panel_info(page):
    """
    读取右侧详情区域的所有信息
    """
    result = {
        "work_code": None,
        "code_info": None,
        "date": None,
        "total_hours": None,
        "phase": None,
        "role": None,
        "product_hours": None,
        "remark": None
    }
    
    # 获取右侧面板的HTML用于调试
    if DEBUG:
        panel_html = await page.evaluate("""() => {
            const panel = document.querySelector('[class*="right"], [class*="detail"], .ant-card, .ant-modal');
            return panel ? panel.outerHTML : 'not found';
        }""")
        dprint(f"        右侧面板HTML预览: {panel_html[:200]}...")
    
    # 1. 查找工时code
    try:
        # 方式1：查找包含"工时code"的标签
        code_label = page.locator('text=/工时code/i').first
        if await code_label.count() > 0:
            # 获取父元素或相邻元素
            code_container = code_label.locator('xpath=..').first
            if await code_container.count() > 0:
                code_text = await code_container.text_content()
                if code_text and is_valid_code(code_text):
                    result["code_info"] = extract_code_and_name(code_text)
                    result["work_code"] = result["code_info"]["code"]
                    dprint(f"        找到工时code: {result['work_code']}")
        
        # 方式2：如果没找到，直接搜索包含code的文本
        if not result["work_code"]:
            code_elements = await page.locator('text=/ac|AC|SQ/').all()
            for el in code_elements[:3]:  # 只检查前几个
                if await el.is_visible():
                    text = await el.text_content()
                    if text and is_valid_code(text):
                        result["code_info"] = extract_code_and_name(text)
                        result["work_code"] = result["code_info"]["code"]
                        dprint(f"        通过文本找到工时code: {result['work_code']}")
                        break
    except Exception as e:
        dprint(f"        查找工时code失败: {e}")
    
    # 2. 查找日期
    try:
        date_label = page.locator('text=/日期/i').first
        if await date_label.count() > 0:
            date_container = date_label.locator('xpath=..').first
            if await date_container.count() > 0:
                date_text = await date_container.text_content()
                result["date"] = extract_date(date_text)
                if result["date"]:
                    dprint(f"        找到日期: {result['date']}")
    except Exception as e:
        dprint(f"        查找日期失败: {e}")
    
    # 如果没有找到有效code，提前返回
    if not result["work_code"]:
        return result
    
    # 3. 根据类型提取详细信息
    if is_management_code(result["work_code"]):
        # 管理类
        try:
            total_label = page.locator('text=/合计工时\\(H\\)/i').first
            if await total_label.count() > 0:
                total_input = total_label.locator('xpath=following::input[1]').first
                if await total_input.count() > 0:
                    val = await total_input.input_value()
                    result["total_hours"] = float(val) if val else 0
                    dprint(f"        合计工时: {result['total_hours']}")
        except:
            pass
        
        try:
            remark_input = page.locator('textarea').first
            if await remark_input.count() > 0:
                result["remark"] = await remark_input.input_value()
                dprint(f"        备注: {result['remark'][:50]}...")
        except:
            pass
    
    elif is_project_code(result["work_code"]):
        # 项目类
        try:
            phase_label = page.locator('text=/工时阶段/i').first
            if await phase_label.count() > 0:
                phase_select = phase_label.locator('xpath=following::select[1]').first
                if await phase_select.count() > 0:
                    result["phase"] = await phase_select.input_value()
                    dprint(f"        工时阶段: {result['phase']}")
        except:
            pass
        
        try:
            role_label = page.locator('text=/工时角色/i').first
            if await role_label.count() > 0:
                role_select = role_label.locator('xpath=following::select[1]').first
                if await role_select.count() > 0:
                    result["role"] = await role_select.input_value()
                    dprint(f"        工时角色: {result['role']}")
        except:
            pass
        
        try:
            hours_label = page.locator('text=/工时\\(H\\)/i').first
            if await hours_label.count() > 0:
                hours_input = hours_label.locator('xpath=following::input[1]').first
                if await hours_input.count() > 0:
                    val = await hours_input.input_value()
                    result["total_hours"] = float(val) if val else 0
                    dprint(f"        工时(H): {result['total_hours']}")
        except:
            pass
        
        try:
            product_label = page.locator('text=/产品\\(L3\\)/i').first
            if await product_label.count() > 0:
                product_input = product_label.locator('xpath=following::input[1]').first
                if await product_input.count() > 0:
                    val = await product_input.input_value()
                    result["product_hours"] = float(val) if val else 0
                    dprint(f"        产品工时: {result['product_hours']}")
        except:
            pass
        
        try:
            remark_input = page.locator('textarea').first
            if await remark_input.count() > 0:
                result["remark"] = await remark_input.input_value()
                dprint(f"        备注: {result['remark'][:50]}...")
        except:
            pass
    
    return result

# ==================== 提取单周数据 ====================

async def extract_week_data(page):
    result = {
        "week_label": None,
        "projects": [],
        "days": [],
        "cells": []
    }
    
    # 获取当前周标签
    try:
        active_week = await page.locator('.ant-radio-button-wrapper-checked, [class*="active"], [class*="selected"]').first
        if await active_week.count() > 0:
            result["week_label"] = await active_week.text_content()
    except:
        pass
    
    # 获取表格结构
    structure = await get_table_structure(page)
    result["projects"] = structure["projects"]
    result["days"] = structure["day_headers"]
    
    if not structure["projects"] or not structure["day_headers"]:
        dprint("    [X] 表格结构不完整")
        return result
    
    # 遍历每个单元格
    for p_idx, project in enumerate(structure["projects"]):
        for d_idx, day in enumerate(structure["day_headers"]):
            dprint(f"    [处理] 项目[{p_idx}]: {project}, 天[{d_idx}]: {day}")
            
            cell_info = await click_cell_and_read_info(page, p_idx, d_idx)
            
            if cell_info and cell_info.get("work_code"):
                cell_data = {
                    "project": project,
                    "project_index": p_idx,
                    "day": day,
                    "day_index": d_idx,
                    "work_code": cell_info["work_code"],
                    "code_name": cell_info["code_info"]["name"] if cell_info["code_info"] else None,
                    "date": cell_info["date"],
                    "work_type": "management" if is_management_code(cell_info["work_code"]) else "project",
                    "details": {},
                    "hour_value": cell_info.get("hour_value", 0)
                }
                
                if cell_data["work_type"] == "management":
                    cell_data["details"] = {
                        "total_hours": cell_info.get("total_hours"),
                        "remark_parts": parse_management_remark(cell_info.get("remark")) if cell_info.get("remark") else {}
                    }
                else:
                    cell_data["details"] = {
                        "phase": cell_info.get("phase"),
                        "role": cell_info.get("role"),
                        "total_hours": cell_info.get("total_hours"),
                        "product_hours": cell_info.get("product_hours"),
                        "remark": cell_info.get("remark")
                    }
                
                result["cells"].append(cell_data)
                dprint(f"      [OK] 已保存: {cell_data['work_code']} - {cell_data['date']}")
            else:
                dprint(f"      [跳过] 无有效工时code")
            
            await asyncio.sleep(0.5)
    
    return result

# ==================== 遍历上月所有周 ====================

async def extract_last_month_data(page):
    iprint("\n[数据提取] 开始遍历上个月的所有周...")

    weeks = await page.evaluate("""() => {
        const weekElements = [];
        const selectors = ['.week-item', '.ant-radio-button-wrapper', 'a', 'span[class*="week"]', 'div[class*="week"]', '.week-tab'];
        for (let sel of selectors) {
            const elements = document.querySelectorAll(sel);
            for (let el of elements) {
                const text = el.textContent.trim();
                if (text.match(/\\d{2}\\/\\d{2}-\\d{2}\\/\\d{2}/) || text.includes('本周')) {
                    if (!weekElements.includes(text)) weekElements.push(text);
                }
            }
            if (weekElements.length > 0) break;
        }
        return weekElements;
    }""")

    iprint(f"  找到的周选项: {weeks}")

    if not weeks:
        iprint("  [警告] 未找到周选项，尝试使用硬编码默认周")
        weeks = ["01/26-02/01", "02/02-02/08", "02/09-02/15", "02/16-02/22", "02/23-03/01"]
        iprint(f"  使用硬编码周: {weeks}")

    all_weeks_data = []
    for week_label in weeks:
        iprint(f"\n  正在处理周: {week_label}")
        if not await switch_to_week(page, week_label):
            iprint(f"    切换失败，跳过")
            continue
        
        await scroll_page(page)
        week_data = await extract_week_data(page)
        if week_data["cells"]:
            all_weeks_data.append(week_data)
            iprint(f"    该周提取到 {len(week_data['cells'])} 个有效单元格数据")
        else:
            iprint(f"    该周未提取到有效数据")

    if all_weeks_data:
        iprint(f"\n  共提取 {len(all_weeks_data)} 周的数据")
        
        json_output = json.dumps(all_weeks_data, ensure_ascii=False, indent=2)
        iprint("\n  提取的JSON数据预览：")
        preview = json_output[:2000] + "..." if len(json_output) > 2000 else json_output
        iprint(preview)
        
        with open("extracted_work_hours.json", "w", encoding="utf-8") as f:
            f.write(json_output)
        iprint("\n  数据已保存到 extracted_work_hours.json")
        
        return all_weeks_data
    else:
        iprint("  [X] 未能提取任何周数据")
        return []

# ==================== 填写当月数据（预留）====================

async def fill_current_week(page, last_month_data):
    iprint("\n[填写] 开始填写当月...")
    # 填写逻辑可根据需要实现
    return True

# ==================== 自动提交 ====================

async def auto_submit(page):
    iprint("\n[提交] 自动查找提交按钮...")
    submit_texts = ["保存", "提交", "更新", "确认", "Save", "Submit"]
    clicked = False

    for text in submit_texts:
        try:
            btn = page.get_by_role("button", name=text).first
            if await btn.is_visible(timeout=2000):
                await btn.click()
                iprint(f"  [OK] 点击 '{text}' 按钮")
                clicked = True
                break
        except:
            continue

    if not clicked:
        result = await page.evaluate("""(texts) => {
            const btns = document.querySelectorAll('button, input[type="submit"], [role="button"]');
            for (let btn of btns) {
                const txt = (btn.textContent || btn.value || '').trim();
                if (texts.some(t => txt.includes(t)) && btn.offsetParent !== null && !btn.disabled) {
                    btn.click();
                    return true;
                }
            }
            return false;
        }""", submit_texts)
        if result:
            iprint("  [OK] JavaScript点击提交按钮")
            clicked = True

    if not clicked:
        iprint("  [X] 未找到提交按钮")
        return False

    await asyncio.sleep(1.5)
    return True

# ==================== 主流程 ====================

async def main():
    iprint("="*60)
    iprint("百度工时填报自动脚本 - 精确提取版")
    iprint("="*60)

    if not check_network():
        resp = input("\n网络检查失败，是否继续? (y/N): ")
        if resp.lower() != 'y':
            return

    today = date.today()
    current_month = today.strftime("%Y-%m")
    last_month = (today.replace(day=1) - timedelta(days=1)).strftime("%Y-%m")
    iprint(f"当前月份: {current_month}")
    iprint(f"参考月份: {last_month}")
    iprint(f"预览模式: {'是' if DRY_RUN else '否'}")
    iprint(f"自动提交: {'是' if AUTO_SUBMIT else '否'}")
    iprint("="*60)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page(viewport={"width": 1920, "height": 1080})

        iprint("\n[导航] 打开工时填报首页...")
        await page.goto(BASE_URL, wait_until='domcontentloaded', timeout=30000)
        await analyze_page(page, "首页")

        iprint("\n" + "="*60)
        input("请确认已登录，然后按 Enter 继续...")

        iprint("等待页面加载完成...")
        await page.wait_for_load_state("networkidle", timeout=30000)

        try:
            await page.wait_for_selector("text=工时填报", timeout=10000)
            iprint("  [OK] 检测到‘工时填报’标题")
        except:
            iprint("  [!] 未找到‘工时填报’标题，继续...")

        await scroll_page(page)

        password_input = page.locator('input[type="password"]').first
        if await password_input.is_visible(timeout=5000):
            iprint("[X] 检测到密码输入框，请确保已登录成功")
            await browser.close()
            return

        iprint("[OK] 登录验证通过")

        # 步骤1: 提取上月数据
        if not await switch_to_month(page, last_month):
            iprint("[X] 无法切换到上月，请检查月份选择器")
            await browser.close()
            return
        
        await scroll_page(page)
        last_month_data = await extract_last_month_data(page)
        
        if not last_month_data:
            iprint("[X] 未能提取到上月数据，退出")
            await browser.close()
            return

        # 步骤2: 切换回当月
        if not await switch_to_month(page, current_month):
            iprint("[X] 无法切换回当月，退出")
            await browser.close()
            return
        
        await scroll_page(page)

        # 步骤3: 填写当月（预览模式不填写）
        if DRY_RUN:
            iprint("\n[预览模式] 数据提取完成，未实际填写")
        else:
            filled = await fill_current_week(page, last_month_data)
            if filled and AUTO_SUBMIT:
                await auto_submit(page)

        iprint("\n" + "="*60)
        input("按 Enter 关闭浏览器...")
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())