#!/usr/bin/env python3
"""
百度OKR周报抓取 - V6结构化提取版
针对百度OKR系统的特定格式：
1. 本周工作总结（可能有重复）
2. 员工OKR（O1/O2/O3一级标题 + KR1/KR2/KR3二级标题 + 本周工作内容）
"""

import json
import asyncio
import re
import glob
from datetime import datetime, timedelta
from playwright.async_api import async_playwright

# 配置
MY_ID = 's673090'
STAFF_IDS = [
    's801573', 's733504', 's682083', 's929270', 's798268',
    's811385', 's865585', 's783729', 's400192', 's823700'
]
BASE_URL = 'https://okr.baidu-int.com/pages/weekly.html#/index?mode=steam'

def get_week_date(weeks_ago=0):
    """获取周日期"""
    today = datetime(2026, 2, 11)
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

# ==================== V6核心：结构化提取 ====================

def extract_name_from_okr_header(text):
    """从OKR标题提取姓名，如：'张三的OKR' -> '张三'"""
    # 匹配 "XXX的OKR" 或 "XXX的 OKR"
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

def parse_okr_structure_v6(full_text):
    """
    V6核心：解析OKR结构
    格式：
    - 一级标题：O1/O2/O3 (通常有小图标前缀)
    - 二级标题：KR1/KR2/KR3
    - 内容：二级标题下的"本周工作"及后续内容
    """
    lines = full_text.split('\n')
    lines = [l.strip() for l in lines if l.strip()]
    
    result = {
        'employee_name': None,
        'okr_title': None,  # 如："张三的OKR"
        'okr_structure': [],  # 结构化的OKR
        'general_work_summary': [],  # 第一部分：通用工作总结
        'raw_text': full_text
    }
    
    # 步骤1：提取OKR标题和姓名
    for i, line in enumerate(lines[:30]):
        if 'OKR' in line.upper():
            result['okr_title'] = line
            name = extract_name_from_okr_header(line)
            if name:
                result['employee_name'] = name
            break
    
    # 步骤2：识别OKR结构
    current_o = None  # 当前Objective
    current_kr = None  # 当前Key Result
    
    # 模式定义
    o_pattern = re.compile(r'^(?:.*?)\s*[Oo](\d+)[:：\s]*(.*)', re.IGNORECASE)  # O1, O2, O3
    kr_pattern = re.compile(r'^[Kk][Rr](\d+)[:：\s]*(.*)')  # KR1, KR2, KR3
    weekly_work_pattern = re.compile(r'^(?:本周工作|本周进展|工作进展|进展)[:：\s]*(.*)', re.IGNORECASE)
    
    in_okr_section = False
    
    for i, line in enumerate(lines):
        # 检测进入OKR部分
        if result['okr_title'] and line == result['okr_title']:
            in_okr_section = True
            continue
        
        # 检测O（一级标题）
        o_match = o_pattern.match(line)
        if o_match:
            in_okr_section = True
            o_num = o_match.group(1)
            o_title = o_match.group(2).strip() or line
            
            current_o = {
                'level': 1,
                'id': f'O{o_num}',
                'title': o_title,
                'key_results': []
            }
            result['okr_structure'].append(current_o)
            current_kr = None
            continue
        
        # 检测KR（二级标题）
        kr_match = kr_pattern.match(line)
        if kr_match and current_o:
            kr_num = kr_match.group(1)
            kr_title = kr_match.group(2).strip() or line
            
            current_kr = {
                'level': 2,
                'id': f'KR{kr_num}',
                'title': kr_title,
                'weekly_work': []  # KR下的本周工作内容
            }
            current_o['key_results'].append(current_kr)
            continue
        
        # 检测"本周工作"内容
        weekly_match = weekly_work_pattern.match(line)
        if weekly_match and current_kr:
            content = weekly_match.group(1).strip()
            if content:
                current_kr['weekly_work'].append(content)
            continue
        
        # 收集KR下的内容（如果当前在KR中，且不是新的O或KR）
        if current_kr and not o_pattern.match(line) and not kr_pattern.match(line):
            # 可能是多行内容的后续
            if len(line) > 5:  # 忽略太短的行
                current_kr['weekly_work'].append(line)
    
    # 步骤3：提取第一部分（通用工作总结）
    # 通常在OKR之前的"本周工作总结"部分
    general_work_section = []
    found_okr = False
    
    for line in lines:
        if 'OKR' in line.upper() or o_pattern.match(line):
            found_okr = True
            break
        
        # 收集OKR之前的工作总结内容
        if any(keyword in line for keyword in ['本周工作', '工作总结', '工作进展']):
            continue  # 跳过标题行
        
        if not found_okr and len(line) > 10:
            general_work_section.append(line)
    
    result['general_work_summary'] = deduplicate_lines(general_work_section, 0.85)
    
    # 对OKR结构内的本周工作内容也去重
    for o in result['okr_structure']:
        for kr in o.get('key_results', []):
            kr['weekly_work'] = deduplicate_lines(kr['weekly_work'], 0.85)
    
    return result

def extract_next_week_plan(full_text):
    """提取下周计划"""
    lines = full_text.split('\n')
    lines = [l.strip() for l in lines if l.strip()]
    
    plan_section = []
    in_plan = False
    
    for line in lines:
        # 检测进入下周计划部分
        if any(keyword in line for keyword in ['下周计划', '下周工作', 'Next Week', '后续计划']):
            in_plan = True
            continue
        
        # 检测退出（遇到新的标题）
        if in_plan and re.match(r'^[Oo]\d+|^KR\d+|^本周|^#', line, re.IGNORECASE):
            break
        
        if in_plan and len(line) > 5:
            plan_section.append(line)
    
    return deduplicate_lines(plan_section, 0.85)

# ==================== 浏览器操作 ====================

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

async def scroll_container_v6(page, selector, center):
    """V6：确保内容完整加载"""
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

# ==================== 抓取流程 ====================

async def fetch_staff_v6(page, staff_id, date, desc, index=0, total=0):
    """V6抓取 - 结构化提取"""
    url = f"{BASE_URL}&id={staff_id}&date={date}"
    progress = f"[{index}/{total}] " if index else ""
    
    print(f"\n{'='*60}")
    print(f"{progress}抓取: {desc}")
    print(f"URL: {url}")
    
    try:
        print("  加载页面...", end=" ")
        await page.goto(url, wait_until='domcontentloaded', timeout=60000)
        await asyncio.sleep(5)  # 等待初始渲染
        
        if 'uuap' in page.url:
            print("需要登录")
            return {'id': staff_id, 'date': date, 'status': 'login_required'}
        print("✓")
        
        print("  滚动加载...", end=" ")
        selector, center = await find_scroll_container(page)
        await scroll_container_v6(page, selector, center)
        print("✓")
        
        print("  提取内容...", end=" ")
        full_text = await page.evaluate('() => document.body.innerText')
        print(f"{len(full_text)}字符")
        
        print("  解析OKR结构...", end=" ")
        parsed = parse_okr_structure_v6(full_text)
        employee_name = parsed.get('employee_name') or staff_id
        print(f"姓名:{employee_name}, O数量:{len(parsed['okr_structure'])}")
        
        print("  提取下周计划...", end=" ")
        next_week = extract_next_week_plan(full_text)
        print(f"{len(next_week)}条")
        
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
                'okr_structure': parsed.get('okr_structure', []),
                'general_work_summary': parsed.get('general_work_summary', []),
                'next_week_plan': next_week
            },
            'stats': {
                'text_length': len(full_text),
                'o_count': len(parsed.get('okr_structure', [])),
                'kr_count': sum(len(o.get('key_results', [])) for o in parsed.get('okr_structure', [])),
                'work_item_count': len(parsed.get('general_work_summary', [])) + sum(
                    len(kr.get('weekly_work', [])) 
                    for o in parsed.get('okr_structure', [])
                    for kr in o.get('key_results', [])
                )
            }
        }
        
        json_file = f'{safe_name}_{date}_周报.json'
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        screenshot = f'{safe_name}_{date}_周报.png'
        await page.screenshot(path=screenshot, full_page=True)
        
        print(f"  ✓ 完成: {safe_name}")
        
        return result
        
    except Exception as e:
        print(f"\n✗ 失败: {e}")
        import traceback
        traceback.print_exc()
        return {'id': staff_id, 'date': date, 'status': 'error', 'error': str(e)}

# ==================== 整合功能 ====================

def integrate_weekly_reports_v6():
    """V6整合 - 基于结构化数据"""
    print("\n" + "="*60)
    print("整合周报...")
    
    reports = []
    for file in glob.glob('*_周报.json'):
        try:
            with open(file, 'r', encoding='utf-8') as f:
                reports.append(json.load(f))
        except Exception as e:
            print(f"  警告: 无法读取 {file} - {e}")
    
    if not reports:
        print("  ✗ 未找到周报文件")
        return None
    
    print(f"  加载了 {len(reports)} 份周报")
    
    # 生成报告
    team_md = generate_team_weekly_v6(reports)
    with open('team_weekly_report.md', 'w', encoding='utf-8') as f:
        f.write(team_md)
    print(f"  ✓ 团队周报: team_weekly_report.md")
    
    llm_prompt = generate_llm_prompt_v6(reports)
    with open('llm_prompt.txt', 'w', encoding='utf-8') as f:
        f.write(llm_prompt)
    print(f"  ✓ 提示词: llm_prompt.txt")
    
    personal_md = generate_personal_reference_v6(reports)
    with open('personal_reference.md', 'w', encoding='utf-8') as f:
        f.write(personal_md)
    print(f"  ✓ 个人参考: personal_reference.md")
    
    return {'team_md': team_md, 'llm_prompt': llm_prompt, 'reports': reports}

def generate_team_weekly_v6(reports):
    """V6生成团队周报 - 按照OKR结构"""
    lines = [
        "# 团队周报汇总",
        "",
        f"**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        f"**团队成员数**: {len(reports)}",
        "",
        "---",
        ""
    ]
    
    # 跨人员去重
    seen_work_items = set()
    
    for report in reports:
        name = report.get('name', 'Unknown')
        date = report.get('date', 'Unknown')
        content = report.get('content', {})
        
        lines.extend([
            f"## {name}",
            "",
            f"**周期**: {date}",
            ""
        ])
        
        # 第一部分：通用工作总结
        general_work = content.get('general_work_summary', [])
        if general_work:
            lines.extend([
                "### 本周工作总结",
                ""
            ])
            for item in general_work[:15]:
                item_norm = normalize_text(item)
                if item_norm not in seen_work_items:
                    lines.append(f"- {item}")
                    seen_work_items.add(item_norm)
                else:
                    lines.append(f"- {item} (重复)")
            lines.append("")
        
        # 第二部分：OKR结构
        okr_structure = content.get('okr_structure', [])
        if okr_structure:
            lines.append("### OKR进展")
            lines.append("")
            
            for o in okr_structure:
                o_title = o.get('title', '')
                lines.append(f"#### {o.get('id', 'O')} {o_title}")
                lines.append("")
                
                for kr in o.get('key_results', []):
                    kr_title = kr.get('title', '')
                    lines.append(f"**{kr.get('id', 'KR')} {kr_title}**")
                    lines.append("")
                    
                    weekly_work = kr.get('weekly_work', [])
                    if weekly_work:
                        lines.append("本周工作：")
                        for work in weekly_work[:10]:
                            work_norm = normalize_text(work)
                            if work_norm not in seen_work_items:
                                lines.append(f"- {work}")
                                seen_work_items.add(work_norm)
                            else:
                                lines.append(f"- {work} (重复)")
                        lines.append("")
        
        # 下周计划
        next_week = content.get('next_week_plan', [])
        if next_week:
            lines.extend([
                "### 下周工作计划",
                ""
            ])
            for item in next_week[:10]:
                lines.append(f"- {item}")
            lines.append("")
        
        lines.extend(["---", ""])
    
    return '\n'.join(lines)

def generate_llm_prompt_v6(reports):
    """V6生成提示词"""
    lines = [
        "基于以下团队成员的周报，请生成一份团队周报汇总：",
        "",
        "要求：",
        "1. 按人员分别总结每个人的本周工作",
        "2. 重点汇总OKR进展",
        "3. 提取团队整体的关键进展",
        "4. 识别需要关注的风险或问题",
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
        
        # OKR结构
        okr_structure = content.get('okr_structure', [])
        if okr_structure:
            lines.append("OKR进展：")
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
        
        # 通用工作总结
        general_work = content.get('general_work_summary', [])
        if general_work:
            lines.append("其他工作总结：")
            for item in general_work[:5]:
                lines.append(f"- {item}")
            lines.append("")
    
    return '\n'.join(lines)

def generate_personal_reference_v6(reports):
    """V6生成个人参考"""
    my_reports = [r for r in reports if '本人' in r.get('desc', '')]
    
    lines = [
        "# 个人历史周报参考",
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
        
        # OKR
        okr_structure = content.get('okr_structure', [])
        if okr_structure:
            lines.append("### OKR")
            lines.append("")
            for o in okr_structure:
                lines.append(f"- {o.get('id', 'O')}: {o.get('title', '')}")
            lines.append("")
        
        # 工作
        general_work = content.get('general_work_summary', [])
        if general_work:
            lines.append("### 工作总结")
            lines.append("")
            for item in general_work[:10]:
                lines.append(f"- {item}")
            lines.append("")
        
        lines.extend(["---", ""])
    
    return '\n'.join(lines)

# ==================== 主流程 ====================

async def main():
    last_week = get_week_date(0)
    week_before = get_week_date(1)
    
    print("="*60)
    print(" "*15 + "百度OKR周报抓取 - V6结构化版")
    print("="*60)
    print("\n【V6特性】")
    print("  ✓ 针对百度OKR特定格式优化")
    print("  ✓ 提取O1/O2/O3一级标题")
    print("  ✓ 提取KR1/KR2/KR3二级标题")
    print("  ✓ 提取KR下的'本周工作'内容")
    print("  ✓ 第一部分通用工作总结去重")
    print(f"\n日期: 上周{last_week}, 上上周{week_before}")
    
    async with async_playwright() as p:
        print("\n" + "="*60)
        print("[系统] 启动Chrome...")
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page(viewport={'width': 1920, 'height': 1080})
        print("✓ Chrome已启动")
        
        print("\n[系统] 访问OKR...")
        await page.goto("https://okr.baidu-int.com", wait_until='domcontentloaded')
        print("✓ 页面已加载")
        
        print("\n" + "="*60)
        print("⚠️  请先登录OKR系统")
        print("="*60)
        input("\n登录完成后按 Enter...")
        
        if 'uuap' in page.url:
            print("\n✗ 未登录成功")
            return
        
        print("✓ 登录验证通过")
        
        all_results = []
        
        # 抓取下属
        print(f"\n{'='*60}")
        print(f"抓取 {len(STAFF_IDS)} 个下属上周周报")
        print(f"{'='*60}")
        
        for i, staff_id in enumerate(STAFF_IDS, 1):
            result = await fetch_staff_v6(page, staff_id, last_week, f"下属{i}", i, len(STAFF_IDS))
            all_results.append(result)
            
            if result['status'] == 'login_required':
                print("\n⚠️ 会话失效，停止")
                break
            
            if i < len(STAFF_IDS):
                await asyncio.sleep(1.5)
        
        # 抓取本人
        if not any(r['status'] == 'login_required' for r in all_results):
            print(f"\n{'='*60}")
            print("抓取本人历史2周")
            print(f"{'='*60}")
            
            r1 = await fetch_staff_v6(page, MY_ID, last_week, "本人-上周", 1, 2)
            all_results.append(r1)
            
            if r1['status'] != 'login_required':
                await asyncio.sleep(2)
                r2 = await fetch_staff_v6(page, MY_ID, week_before, "本人-上上周", 2, 2)
                all_results.append(r2)
        
        await browser.close()
        
        # 保存汇总
        summary = {
            'fetch_time': datetime.now().isoformat(),
            'my_id': MY_ID,
            'success': sum(1 for r in all_results if r['status'] == 'success'),
            'failed': sum(1 for r in all_results if r['status'] == 'error'),
            'results': [{'id': r.get('id'), 'name': r.get('name'), 'status': r.get('status')} for r in all_results]
        }
        
        with open('v6_fetch_summary.json', 'w', encoding='utf-8') as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)
        
        print(f"\n{'='*60}")
        print("抓取完成!")
        print(f"  成功: {summary['success']}, 失败: {summary['failed']}")
        
        # 自动整合
        integrate_result = integrate_weekly_reports_v6()
        
        if integrate_result:
            print(f"\n{'='*60}")
            print("全部完成!")
            print(f"{'='*60}")
            print(f"\n生成的文件:")
            print(f"  - 姓名_日期_周报.json (结构化数据)")
            print(f"  - 姓名_日期_周报.png (截图)")
            print(f"  - team_weekly_report.md (按OKR结构组织的团队周报)")
            print(f"  - llm_prompt.txt (大模型提示词)")
            print(f"  - personal_reference.md (个人历史参考)")
        
        input("\n按Enter退出...")

if __name__ == '__main__':
    asyncio.run(main())
