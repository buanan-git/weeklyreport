#!/usr/bin/env python3
"""
百度OKR周报抓取 - V10多行标题合并版
基于V10修复以下问题：
1. O1/O2/O3标题可能分布在多行，需要合并（如：O1 + 标题内容 + 状态）
2. KR1/KR2/KR3标题可能分布在多行，需要合并（如：KR1 + 标题内容 + 进度）
3. "本周工作"之后的行才是真正的周报内容
4. 本人和下属抓取逻辑完全一致

针对百度OKR系统的特定格式：
1. 本周工作总结（可能有重复，需合并）
2. OKR进度（三级表格结构：总标题 -> O1/O2/O3 -> KR1/KR2/KR3 -> 本周工作）
3. 其他内容（业务核心指标、主要项目、下周重点工作）
4. 新增：首次页面等待5秒 + 失败重试机制（最多1次）
针对百度OKR系统的特定格式优化：
1. 本周工作总结（可能有重复，需合并）
2. OKR进度（三级表格结构：总标题 -> O1/O2/O3 -> KR1/KR2/KR3 -> 本周工作）
3. 其他内容（业务核心指标、主要项目、下周重点工作）
4. 新增：首次页面等待5秒 + 失败重试机制（最多1次）
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
    's673090','s801573'
    #'s801573', 's733504', 's682083', 's929270', 's798268',
    #'s811385', 's865585', 's783729', 's400192', 's823700'
]
BASE_URL = 'https://okr.baidu-int.com/pages/weekly.html#/index?mode=steam&mandatorId=undefined'

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

def parse_weekly_report_structure_v10(full_text):
    """
    V10核心：解析三部分结构
    
    V10多行标题合并修复：
    1. O1/O2/O3标题可能分布在多行（如：O1 + 标题内容 + 状态），需要合并到title
    2. KR1/KR2/KR3标题可能分布在多行（如：KR1 + 标题内容 + 进度），需要合并到title
    3. "本周工作"之后的行才是真正的周报内容，放入weekly_work
    4. 严格的匹配优先级：O标题 > KR标题 > 本周工作 > 其他内容
    
    第一部分：本周工作总结
    - 以"本周工作总结"为标题
    - 多次出现则合并
    
    第二部分：OKR进度（三级表格结构）
    - 总标题（如"张三的OKR"）
    - O1标题（可能多行合并）
    - KR1标题（可能多行合并）
    - 本周工作内容
    - 后续KR和O...
    
    第三部分：其他内容
    - 业务核心指标
    - 主要项目
    - 下周重点工作
    """
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
    
    # 步骤2：提取第一部分（本周工作总结）
    # 查找所有"本周工作总结"标题，提取其后的内容直到遇到OKR相关标题
    i = 0
    while i < len(lines):
        line = lines[i]
        if '本周工作总结' in line:
            # 提取此标题后的内容
            i += 1
            while i < len(lines):
                content = lines[i]
                # 如果遇到OKR标题或O/KR标题，停止
                if 'OKR' in content.upper() or re.match(r'^[Oo]\d+', content) or \
                   re.match(r'^[Kk][Rr]\d+', content) or '本周工作' in content:
                    break
                if len(content) > 5:  # 忽略太短的行
                    result['part1_work_summary'].append(content)
                i += 1
            continue
        i += 1
    
    result['part1_work_summary'] = deduplicate_lines(result['part1_work_summary'], 0.85)
    
    # 步骤3：提取第二部分（OKR三级结构）- V10多行标题合并
    if okr_header_index >= 0:
        # 从OKR标题后开始解析
        okr_lines = lines[okr_header_index + 1:]
        
        # V10: 打印前30行用于调试
        print(f"\n    [V10调试] OKR部分前30行内容:")
        for idx, l in enumerate(okr_lines[:30]):
            print(f"      [{idx}] {l[:100]}")
        print()
        
        current_o = None
        current_kr = None
        
        # V10: 匹配O1/O2/O3或KR1/KR2/KR3开头（仅匹配标记部分）
        o_pattern = re.compile(r'^\s*[Oo](\d+)\b')
        kr_pattern = re.compile(r'^\s*[Kk][Rr](\d+)\b')
        weekly_work_pattern = re.compile(r'^(?:本周工作|本周进展|工作进展)', re.IGNORECASE)
        part3_pattern = re.compile(r'(?:业务核心指标|主要项目|下周重点工作)')
        
        i = 0
        while i < len(okr_lines):
            line = okr_lines[i]
            
            # V10: 检测第三部分标题，停止OKR解析
            if part3_pattern.search(line):
                print(f"    [V10] 检测到第三部分标题，停止OKR解析")
                break
            
            # V10: 检测O（二级标题）- 优先级最高
            o_match = o_pattern.match(line)
            if o_match:
                o_num = o_match.group(1)
                # V10: 收集O标题的多行内容
                o_title_parts = [line.strip()]
                i += 1
                # 继续读取后续行，直到遇到KR、下一个O、本周工作或第三部分
                while i < len(okr_lines):
                    next_line = okr_lines[i]
                    if kr_pattern.match(next_line) or o_pattern.match(next_line) or \
                       weekly_work_pattern.match(next_line) or part3_pattern.search(next_line):
                        break
                    # 将非标题行追加到O标题
                    o_title_parts.append(next_line.strip())
                    i += 1
                
                o_title = ' '.join(o_title_parts)
                print(f"    [V10] O{o_num}标题(合并): {o_title[:100]}")
                
                current_o = {
                    'level': 1,
                    'id': f'O{o_num}',
                    'title': o_title,  # V10: 合并后的多行内容
                    'key_results': []
                }
                result['part2_okr_structure'].append(current_o)
                current_kr = None  # 重置当前KR
                continue  # i已经指向下一个待处理行
            
            # V10: 检测KR（三级标题）- 优先级第二
            kr_match = kr_pattern.match(line)
            if kr_match and current_o:
                kr_num = kr_match.group(1)
                # V10: 收集KR标题的多行内容
                kr_title_parts = [line.strip()]
                i += 1
                # 继续读取后续行，直到遇到"本周工作"、下一个O/KR或第三部分
                while i < len(okr_lines):
                    next_line = okr_lines[i]
                    if weekly_work_pattern.match(next_line) or \
                       kr_pattern.match(next_line) or o_pattern.match(next_line) or \
                       part3_pattern.search(next_line):
                        break
                    # 将非标题行追加到KR标题
                    kr_title_parts.append(next_line.strip())
                    i += 1
                
                kr_title = ' '.join(kr_title_parts)
                print(f"    [V10] KR{kr_num}标题(合并): {kr_title[:100]}")
                
                current_kr = {
                    'level': 2,
                    'id': f'KR{kr_num}',
                    'title': kr_title,  # V10: 合并后的多行内容
                    'weekly_work': []   # V10: 初始为空，等待"本周工作"后的内容
                }
                current_o['key_results'].append(current_kr)
                continue  # i已经指向下一个待处理行
            
            # V10: 检测"本周工作"开头的输入框内容 - 优先级第三
            weekly_match = weekly_work_pattern.match(line)
            if weekly_match and current_kr:
                # 提取"本周工作"后的内容（如果有）
                content_after = line[len(weekly_match.group(0)):].strip()
                # 移除可能的前缀符号（如+10%后面的内容）
                content_after = re.sub(r'^[:：\s]+', '', content_after)
                if content_after:
                    print(f"    [V10] 本周工作(首行): {content_after[:80]}")
                    current_kr['weekly_work'].append(content_after)
                
                # V10: 继续读取后续行，直到遇到下一个O/KR/第三部分标题
                i += 1
                while i < len(okr_lines):
                    next_line = okr_lines[i]
                    # V10: 严格检查：如果遇到新的O/KR或第三部分标题，立即停止
                    if o_pattern.match(next_line) or kr_pattern.match(next_line) or \
                       part3_pattern.search(next_line):
                        break
                    # V10: 添加到当前KR的本周工作
                    if len(next_line.strip()) > 1:  # V10: 很低的阈值，几乎接收所有内容
                        print(f"    [V10] 本周工作(续): {next_line.strip()[:80]}")
                        current_kr['weekly_work'].append(next_line.strip())
                    i += 1
                continue
            
            # V10: 跳过不属于任何类别的行
            i += 1
        
        # V10: 对OKR结构内的本周工作内容去重
        for o in result['part2_okr_structure']:
            for kr in o.get('key_results', []):
                kr['weekly_work'] = deduplicate_lines(kr['weekly_work'], 0.85)
        
        # V10: 打印解析结果摘要
        print(f"\n    [V10] 解析完成摘要:")
        for o in result['part2_okr_structure']:
            print(f"      {o['id']}: {o['title'][:80]}")
            for kr in o.get('key_results', []):
                print(f"        {kr['id']}: title={kr['title'][:60]}, work_items={len(kr['weekly_work'])}")
    
    # 步骤4：提取第三部分（其他内容）
    # 查找"业务核心指标"、"主要项目"、"下周重点工作"
    part3_keywords = {
        '业务核心指标': [],
        '主要项目': [],
        '下周重点工作': []
    }
    
    for keyword in part3_keywords.keys():
        for i, line in enumerate(lines):
            if keyword in line:
                # 提取此标题后的内容
                j = i + 1
                while j < len(lines):
                    content = lines[j]
                    # 如果遇到下一个第三部分标题，停止
                    if any(k in content for k in part3_keywords.keys() if k != keyword):
                        break
                    # 如果遇到OKR相关内容，可能是误匹配，停止
                    if re.match(r'^[Oo]\d+', content) or re.match(r'^[Kk][Rr]\d+', content):
                        break
                    if len(content) > 5:
                        part3_keywords[keyword].append(content)
                    j += 1
                break
        
        result['part3_metrics'][keyword] = deduplicate_lines(part3_keywords[keyword], 0.85)
    
    return result

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
        print(f"{len(full_text)}字符")
        print(f"&&&&&&&&&&&&&&&&&&&&&&&&&&&&{full_text}")
        print(f"_+_+_++++++++++++++++++++++++++++++++++++")    
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
        
        # V7: 移除截图功能（因未实现滚动截图，内容不全）
        # screenshot = f'{safe_name}_{date}_周报.png'
        # await page.screenshot(path=screenshot, full_page=True)
        
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


# ==================== 整合功能（V10格式）====================

def integrate_weekly_reports_v10():
    """V10整合 - 基于三部分结构"""
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
    team_md = generate_team_weekly_v10(reports)
    with open('team_weekly_report_v10.md', 'w', encoding='utf-8') as f:
        f.write(team_md)
    print(f"  ✓ 团队周报: team_weekly_report_v10.md")
    
    llm_prompt = generate_llm_prompt_v10(reports)
    with open('llm_prompt_v10.txt', 'w', encoding='utf-8') as f:
        f.write(llm_prompt)
    print(f"  ✓ 提示词: llm_prompt_v10.txt")
    
    personal_md = generate_personal_reference_v10(reports)
    with open('personal_reference_v10.md', 'w', encoding='utf-8') as f:
        f.write(personal_md)
    print(f"  ✓ 个人参考: personal_reference_v10.md")
    
    return {'team_md': team_md, 'llm_prompt': llm_prompt, 'reports': reports}

def generate_team_weekly_v10(reports):
    """V10生成团队周报 - 统一格式汇总（以s673090为基准结构，合并本周工作）"""
    lines = [
        "# 团队周报汇总 - V10统一格式版（基准结构+合并本周工作）",
        "",
        f"**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        f"**团队成员数**: {len(reports)}",
        "",
        "---",
        ""
    ]

    # 跨人员去重
    seen_work_items = set()

    # 收集所有成员的数据
    all_part1 = []
    all_metrics = {
        '业务核心指标': [],
        '主要项目': [],
        '下周重点工作': []
    }

    # 查找基准员工(s673090)的OKR结构
    base_okr_structure = []
    for report in reports:
        if report.get('id') == 's673090':
            base_okr_structure = report.get('content', {}).get('part2_okr_structure', [])
            break

    # 如果没有找到基准员工，使用第一个员工的结构
    if not base_okr_structure and reports:
        base_okr_structure = reports[0].get('content', {}).get('part2_okr_structure', [])

    # 创建合并用的本周工作映射 (O_id -> KR_id -> [weekly_work_items])
    weekly_work_map = {}

    for report in reports:
        content = report.get('content', {})

        # 收集第一部分
        part1 = content.get('part1_work_summary', [])
        all_part1.extend(part1)

        # 收集所有员工的本周工作（按O和KR分组）
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

        # 收集第三部分
        metrics = content.get('part3_metrics', {})
        for key in all_metrics.keys():
            if key in metrics:
                all_metrics[key].extend(metrics[key])

    # 去重第一部分和第三部分
    all_part1 = deduplicate_lines(all_part1, 0.85)
    for key in all_metrics.keys():
        all_metrics[key] = deduplicate_lines(all_metrics[key], 0.85)

    # 过滤掉包含"欢迎使用辅助区"及后续7行的内容
    def filter_welcome_text(items):
        filtered = []
        skip_count = 0
        for item in items:
            if skip_count > 0:
                skip_count -= 1
                continue
            if '欢迎使用辅助区' in item:
                skip_count = 7
                continue
            filtered.append(item)
        return filtered

    all_part1 = filter_welcome_text(all_part1)
    for key in all_metrics.keys():
        all_metrics[key] = filter_welcome_text(all_metrics[key])

    # 第一部分：本周工作总结（汇总）
    if all_part1:
        lines.extend([
            "## 第一部分：本周工作总结",
            ""
        ])
        for item in all_part1[:30]:
            item_norm = normalize_text(item)
            if item_norm not in seen_work_items:
                lines.append(f"- {item}")
                seen_work_items.add(item_norm)
            else:
                lines.append(f"- {item} (重复)")
        lines.append("")

    # 第二部分：OKR进度（以s673090为基准，合并所有员工的本周工作）
    if base_okr_structure:
        lines.extend([
            "## 第二部分：OKR进度",
            ""
        ])

        for o in base_okr_structure:
            o_id = o.get('id', '')
            o_title = o.get('title', '')
            lines.append(f"### {o_id} {o_title}")
            lines.append("")

            for kr in o.get('key_results', []):
                kr_id = kr.get('id', '')
                kr_title = kr.get('title', '')
                lines.append(f"**{kr_id} {kr_title}**")
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
        
        # 第三部分：其他内容
        metrics = content.get('part3_metrics', {})
        if any(metrics.values()):
            lines.extend([
                "### 第三部分：其他内容",
                ""
            ])
            
            for title, items in metrics.items():
                if items:
                    lines.append(f"**{title}**：")
                    for item in items[:10]:
                        lines.append(f"- {item}")
                    lines.append("")
        
        lines.extend(["---", ""])
    
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


# ==================== 主流程（带重试机制）====================

async def main():
    last_week = get_week_date(0)
    week_before = get_week_date(1)
    
    print("="*60)
    print(" "*15 + "百度OKR周报抓取 - V10表格结构最终修复版")
    print("="*60)
    print("\n【V10最终修复】")
    print("  ✓ O1/O2/O3标题严格提取整行内容（包括标记后的所有文字）")
    print("  ✓ KR1/KR2/KR3标题严格放入title字段（绝不混入weekly_work）")
    print("  ✓ 本人和下属抓取逻辑完全一致")
    print("  ✓ 第一部分：本周工作总结（支持多次出现合并）")
    print("  ✓ 第二部分：OKR进度（三级表格结构）")
    print("  ✓ 第三部分：业务核心指标、主要项目、下周重点工作")
    print("  ✓ 新增：首次页面等待5秒")
    print("  ✓ 新增：失败页面自动重试机制（最多1次）")
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
        failed_tasks = []  # 记录失败任务用于重试
        
        # 抓取下属（含本人）
        print(f"\n{'='*60}")
        print(f"抓取 {len(STAFF_IDS)} 个员工上周周报")
        print(f"{'='*60}")
        
        for i, staff_id in enumerate(STAFF_IDS, 1):
            print(f"fetch_single_staff_v10参数：{page},{staff_id},{last_week}")
            result = await fetch_single_staff_v10(page, staff_id, last_week, f"员工{i}", i, len(STAFF_IDS))
            all_results.append(result)
            
            if result['status'] == 'login_required':
                print("\n⚠️ 会话失效，停止")
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
                print(f"\n{'='*60}")
                print(f"检测到本人ID({MY_ID})，额外抓取上上周周报")
                print(f"{'='*60}")
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
            print(f"\n{'='*60}")
            print(f"重试机制：{len(failed_tasks)} 个页面需要重试")
            print(f"{'='*60}")
            
            for task in failed_tasks:
                print(f"\n重试: {task['desc']}")
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
        
        # 关闭浏览器
        await browser.close()
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
        
        print(f"\n{'='*60}")
        print("抓取完成!")
        print(f"  成功: {summary['success']}, 失败: {summary['failed']}, 重试: {summary['retried']}")
        
        # 自动整合
        integrate_result = integrate_weekly_reports_v10()
        
        if integrate_result:
            print(f"\n{'='*60}")
            print("全部完成!")
            print(f"{'='*60}")
            print(f"\n生成的文件:")
            print(f"  - 姓名_日期_周报.json (结构化数据)")
            print(f"  - team_weekly_report_v10.md (V10三部分结构团队周报 - O/KR标题最终修复)")
            print(f"  - llm_prompt_v10.txt (大模型提示词)")
            print(f"  - personal_reference_v10.md (个人历史参考)")
            print(f"  - v10_fetch_summary.json (抓取汇总)")
        
        input("\n按Enter退出...")

if __name__ == '__main__':
    asyncio.run(main())