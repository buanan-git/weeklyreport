#!/usr/bin/env python3
"""
百度OKR周报抓取 - V5.2修正版
变更说明:
1. 从"汽车项目交付组-"后提取员工姓名
2. 增加等待时间确保内容完整加载（解决摘要遗漏问题）
3. 移除摘要功能，保留完整原文（仅过滤无效信息+去重）
4. 抓取和整合在一个脚本内完成
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

# ==================== 内容提取与过滤（优化版） ====================

def extract_employee_name_v5(full_text):
    """
    V5优化：从"汽车项目交付组-"后提取员工姓名
    """
    # 策略1：直接匹配 "汽车项目交付组-姓名" 模式
    pattern = r'汽车项目交付组[-–—]\s*([\u4e00-\u9fa5]{2,4})'
    match = re.search(pattern, full_text)
    if match:
        return match.group(1)
    
    # 策略2：寻找包含该前缀的行
    lines = full_text.split('\n')
    for line in lines[:50]:  # 只看前50行
        if '汽车项目交付组' in line:
            # 提取连字符后的内容
            parts = re.split(r'[-–—]', line)
            if len(parts) >= 2:
                name_part = parts[-1].strip()
                # 提取2-4个中文字符
                name_match = re.search(r'([\u4e00-\u9fa5]{2,4})', name_part)
                if name_match:
                    return name_match.group(1)
    
    # 策略3：寻找 "交付组-姓名" 变体
    pattern2 = r'交付组[-–—]\s*([\u4e00-\u9fa5]{2,4})'
    match2 = re.search(pattern2, full_text)
    if match2:
        return match2.group(1)
    
    # 策略4：回退到通用方法
    return extract_employee_name_fallback(full_text)

def extract_employee_name_fallback(full_text):
    """备用姓名提取方法"""
    lines = full_text.split('\n')
    
    # 寻找包含"周报"、"周总结"附近的文本
    for i, line in enumerate(lines[:30]):
        line = line.strip()
        if any(keyword in line for keyword in ['周报', '周总结', 'Weekly', '的OKR']):
            for j in range(max(0, i-3), min(len(lines), i+3)):
                candidate = lines[j].strip()
                if re.match(r'^[\u4e00-\u9fa5]{2,4}$', candidate):
                    if candidate not in ['周报', '总结', '本周', '下周', '上周']:
                        return candidate
    
    # 寻找特定模式
    patterns = [
        r'([\u4e00-\u9fa5]{2,4})的周报',
        r'([\u4e00-\u9fa5]{2,4})的周总结',
        r'([\u4e00-\u9fa5]{2,4})的OKR',
    ]
    for pattern in patterns:
        match = re.search(pattern, full_text)
        if match:
            return match.group(1)
    
    return None

def extract_key_content_v5(full_text):
    """
    V5优化：更高效的内容提取
    """
    result = {
        'employee_name': None,
        'week_period': None,
        'okr_content': [],
        'work_summary': [],
        'next_week_plan': [],
        'raw_filtered': ""
    }
    
    # 提取姓名（使用V5方法）
    result['employee_name'] = extract_employee_name_v5(full_text)
    
    # 提取时间段
    date_pattern = r'(20\d{2})[-./年](\d{1,2})[-./月](\d{1,2})'
    dates = re.findall(date_pattern, full_text)
    if len(dates) >= 2:
        result['week_period'] = f"{dates[0][0]}-{dates[0][1]:0>2}-{dates[0][2]:0>2} 至 {dates[1][0]}-{dates[1][1]:0>2}-{dates[1][2]:0>2}"
    
    lines = full_text.split('\n')
    lines = [l.strip() for l in lines if l.strip()]
    
    # 优化：更精简的噪声模式
    noise_patterns = [
        r'^百度$', r'^OKR系统$', r'^首页$', r'^目标$', r'^绩效$', r'^人才$',
        r'^考核$', r'^管理$', r'^设置$', r'^帮助$', r'^退出$', r'^搜索$',
        r'^消息$', r'^通知$', r'^个人中心$', r'^工作台$', r'^更多$', r'^菜单$',
        r'^Copyright', r'^©', r'^版权所有', r'^百度在线', r'^京ICP',
        r'^隐私政策', r'^用户协议', r'^联系我们$', r'^反馈$', r'^加载中',
        r'^Loading', r'^请稍候', r'^确定$', r'^取消$', r'^保存$', r'^提交$',
        r'^编辑$', r'^删除$', r'^查看$', r'^展开$', r'^收起$',
        r'^汽车项目交付组$',  # 只保留组名，去掉前缀
    ]
    
    filtered_lines = []
    for line in lines:
        is_noise = False
        for pattern in noise_patterns:
            if re.match(pattern, line, re.IGNORECASE):
                is_noise = True
                break
        if len(line) <= 2 and not re.match(r'^[Oo]\d+$', line) and not re.match(r'^[Kk][Rr]\d+$', line):
            is_noise = True
        if not is_noise:
            filtered_lines.append(line)
    
    # 识别OKR部分
    current_section = None
    okr_pattern = re.compile(r'^[Oo]\d+[:：\s]*(.*)', re.IGNORECASE)
    kr_pattern = re.compile(r'^[Kk][Rr]\d+[:：\s]*(.*)', re.IGNORECASE)
    
    for line in filtered_lines:
        okr_match = okr_pattern.match(line)
        if okr_match:
            current_section = {'type': 'objective', 'title': line, 'krs': []}
            result['okr_content'].append(current_section)
            continue
        
        kr_match = kr_pattern.match(line)
        if kr_match and current_section and current_section['type'] == 'objective':
            current_section['krs'].append(line)
            continue
        
        if any(keyword in line for keyword in ['本周工作总结', '本周工作', '工作总结', 'Work Summary']):
            current_section = {'type': 'work_summary', 'content': []}
            result['work_summary'].append(current_section)
            continue
        
        if any(keyword in line for keyword in ['下周计划', 'Next Week', '下周工作']):
            current_section = {'type': 'next_week', 'content': []}
            result['next_week_plan'].append(current_section)
            continue
        
        if current_section and isinstance(current_section.get('content'), list):
            current_section['content'].append(line)
    
    # 去重OKR内容
    result['okr_content'] = deduplicate_okr_content(result['okr_content'])
    
    # 去重工作总结和计划的原始内容
    for section in result['work_summary']:
        if isinstance(section.get('content'), list):
            section['content'] = deduplicate_lines(section['content'], similarity_threshold=0.85)
    
    for section in result['next_week_plan']:
        if isinstance(section.get('content'), list):
            section['content'] = deduplicate_lines(section['content'], similarity_threshold=0.85)
    
    result['raw_filtered'] = '\n'.join(filtered_lines)
    return result

def generate_local_summary_v5(extracted_content):
    """V5优化：更精简的摘要（带去重）"""
    summary = {
        'employee_name': extracted_content.get('employee_name', '未知'),
        'week_period': extracted_content.get('week_period', '未知'),
        'okr_summary': [],
        'work_highlights': [],
        'next_week_focus': [],
        'total_length': len(extracted_content.get('raw_filtered', ''))
    }
    
    # OKR摘要（先去重）
    deduped_okrs = deduplicate_okr_content(extracted_content.get('okr_content', []))
    for obj in deduped_okrs:
        obj_summary = {
            'objective': obj.get('title', '')[:80],
            'kr_count': len(obj.get('krs', [])),
            'key_results': [kr[:100] for kr in obj.get('krs', [])[:3]]  # 最多3个KR
        }
        summary['okr_summary'].append(obj_summary)
    
    # 工作总结（提取关键句子 + 去重）
    work_content = []
    for section in extracted_content.get('work_summary', []):
        work_content.extend(section.get('content', []))
    
    # 先筛选关键句
    important_keywords = ['完成', '上线', '发布', '开发', '测试', '修复', '优化', '%', '个', '项']
    filtered_work = []
    for line in work_content:
        if any(kw in line for kw in important_keywords) and len(line) > 10:
            filtered_work.append(line[:150])  # 先截断
    
    # 再去重
    summary['work_highlights'] = deduplicate_lines(filtered_work[:20], similarity_threshold=0.8)[:15]
    
    # 下周计划（去重）
    plan_content = []
    for section in extracted_content.get('next_week_plan', []):
        plan_content.extend(section.get('content', []))
    
    raw_plans = [line[:150] for line in plan_content if len(line) > 5]
    summary['next_week_focus'] = deduplicate_lines(raw_plans[:10], similarity_threshold=0.8)[:5]
    
    return summary

# ==================== 去重功能 ====================

def normalize_text(text):
    """
    文本标准化，用于去重比较
    - 去除多余空格
    - 统一标点
    - 转为小写（英文部分）
    """
    text = text.strip()
    text = re.sub(r'\s+', ' ', text)  # 多个空格变单个
    text = re.sub(r'[，。？！；：""''（）【】]', ',', text)  # 统一标点
    text = re.sub(r'[,\.]+', ',', text)  # 多个标点变单个
    return text.lower()

def calculate_similarity(str1, str2):
    """
    计算两个字符串的相似度（简单版）
    使用最长公共子序列思想
    """
    # 完全包含
    if str1 in str2 or str2 in str1:
        return 0.9
    
    # 计算共同字符数
    set1 = set(str1)
    set2 = set(str2)
    intersection = len(set1 & set2)
    union = len(set1 | set2)
    
    if union == 0:
        return 0
    
    return intersection / union

def deduplicate_lines(lines, similarity_threshold=0.85):
    """
    去除列表中的重复/相似行
    
    策略：
    1. 完全相同的行 -> 只保留一个
    2. 相似度 > threshold 的行 -> 保留较长的那个
    3. 包含关系的行 -> 保留较长的那个
    
    Args:
        lines: 文本行列表
        similarity_threshold: 相似度阈值（0-1）
    
    Returns:
        去重后的列表
    """
    if not lines:
        return []
    
    result = []
    normalized_map = {}  # 标准化文本 -> 原始文本
    
    for line in lines:
        if not line or len(line.strip()) < 5:  # 跳过太短的内容
            continue
            
        normalized = normalize_text(line)
        
        # 检查是否已存在完全相同的
        if normalized in normalized_map:
            continue
        
        # 检查是否有相似的内容
        is_duplicate = False
        for existing_norm, existing_orig in list(normalized_map.items()):
            # 1. 互相包含（一个包含另一个）
            if normalized in existing_norm or existing_norm in normalized:
                is_duplicate = True
                # 保留较长的
                if len(line) > len(existing_orig):
                    normalized_map[normalized] = line
                    # 移除旧的
                    if existing_norm in normalized_map:
                        del normalized_map[existing_norm]
                break
            
            # 2. 相似度超过阈值
            sim = calculate_similarity(normalized, existing_norm)
            if sim > similarity_threshold:
                is_duplicate = True
                # 保留较长的
                if len(line) > len(existing_orig):
                    normalized_map[normalized] = line
                    if existing_norm in normalized_map:
                        del normalized_map[existing_norm]
                break
        
        if not is_duplicate:
            normalized_map[normalized] = line
    
    # 保持原始顺序（按照第一次出现的顺序）
    seen = set()
    for line in lines:
        normalized = normalize_text(line)
        if normalized in normalized_map and normalized not in seen:
            result.append(normalized_map[normalized])
            seen.add(normalized)
    
    return result

def deduplicate_okr_content(okr_content):
    """
    去除OKR内容中的重复
    - O之间去重
    - 每个O内的KR去重
    """
    seen_objectives = set()
    unique_okrs = []
    
    for obj in okr_content:
        title = obj.get('title', '')
        normalized_title = normalize_text(title)
        
        # 检查O是否重复
        if normalized_title in seen_objectives:
            continue
        seen_objectives.add(normalized_title)
        
        # 对KR去重
        krs = obj.get('krs', [])
        unique_krs = deduplicate_lines(krs, similarity_threshold=0.9)
        
        unique_okrs.append({
            'type': obj.get('type', 'objective'),
            'title': title,
            'krs': unique_krs
        })
    
    return unique_okrs

def sanitize_filename(name):
    """将姓名转换为安全的文件名"""
    if not name:
        return "unknown"
    safe = re.sub(r'[\\/*?:"<>|]', '_', name)
    return safe.strip() if safe.strip() else "unknown"

# ==================== 浏览器操作（优化等待时间） ====================

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
        selector = None
        if first.get('class'):
            selector = f"[class*='{first['class'].split()[0]}']"
        elif first.get('id'):
            selector = f"#{first['id']}"
        else:
            selector = first['tag'].lower()
        return selector, (first.get('x', 960), first.get('y', 600))
    return 'body', (960, 600)

async def scroll_container_v5(page, selector, center):
    """V5修正：增加等待时间确保内容完整加载"""
    await page.mouse.move(center[0], center[1])
    await asyncio.sleep(0.5)
    
    # 增加滚动次数确保懒加载触发
    print("    滚动加载内容...", end=" ", flush=True)
    for i in range(30):  # 恢复到30次
        await page.mouse.wheel(0, 400)
        await asyncio.sleep(0.3)  # 增加等待
        if i % 10 == 0:
            print(f"{i*400}px", end=" ", flush=True)
    print("✓")
    
    # 键盘确认
    try:
        await page.click(selector)
        await asyncio.sleep(0.5)
        for _ in range(5):  # 恢复到5次
            await page.keyboard.press('PageDown')
            await asyncio.sleep(0.3)
    except:
        pass
    
    # 增加最终等待确保动态内容渲染完成
    print("    等待动态内容渲染...", end=" ", flush=True)
    await asyncio.sleep(5)  # 增加到5秒
    print("✓")

# ==================== 抓取流程（优化版） ====================

async def fetch_staff_v5(page, staff_id, date, desc, index=0, total=0):
    """V5优化版抓取 - 减少等待时间"""
    url = f"{BASE_URL}&id={staff_id}&date={date}"
    progress = f"[{index}/{total}] " if index else ""
    
    print(f"\n{'='*60}")
    print(f"{progress}抓取: {desc}")
    print(f"URL: {url}")
    
    try:
        # 步骤合并，减少等待
        print("  加载页面...", end=" ")
        await page.goto(url, wait_until='domcontentloaded', timeout=60000)
        await asyncio.sleep(3)  # 从5秒减少到3秒
        
        if 'uuap' in page.url:
            print("需要登录")
            return {'id': staff_id, 'date': date, 'status': 'login_required'}
        print("✓")
        
        # 查找并滚动
        print("  滚动加载...", end=" ")
        selector, center = await find_scroll_container(page)
        await scroll_container_v5(page, selector, center)
        print("✓")
        
        # 提取内容
        print("  提取内容...", end=" ")
        full_text = await page.evaluate('() => document.body.innerText')
        print(f"{len(full_text)}字符")
        
        # 本地处理（提取姓名 + 过滤 + 去重，不生成摘要）
        print("  提取姓名...", end=" ")
        extracted = extract_key_content_v5(full_text)
        employee_name = extracted.get('employee_name') or staff_id
        print(f"{employee_name}")
        
        safe_name = sanitize_filename(employee_name)
        
        # 统计去重效果
        original_lines = len(full_text.split('\n'))
        filtered_lines = len(extracted.get('raw_filtered', '').split('\n'))
        print(f"  原文: {original_lines}行 → 过滤后: {filtered_lines}行")
        
        # 保存完整数据（原文 + 结构化提取）
        result = {
            'id': staff_id,
            'name': employee_name,
            'date': date,
            'desc': desc,
            'fetch_time': datetime.now().isoformat(),
            'status': 'success',
            'content': {
                'okr': extracted.get('okr_content', []),
                'work_summary': extracted.get('work_summary', []),
                'next_week_plan': extracted.get('next_week_plan', []),
                'raw_filtered': extracted.get('raw_filtered', '')
            },
            'stats': {
                'original_length': len(full_text),
                'filtered_length': len(extracted.get('raw_filtered', '')),
                'original_lines': original_lines,
                'filtered_lines': filtered_lines,
                'has_okr': len(extracted.get('okr_content', [])) > 0
            }
        }
        
        json_file = f'{safe_name}_{date}_周报.json'
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        screenshot = f'{safe_name}_{date}_周报.png'
        await page.screenshot(path=screenshot, full_page=True)
        
        print(f"  ✓ 完成: {safe_name} ({len(extracted.get('raw_filtered', ''))}字符)")
        
        return result
        
    except Exception as e:
        print(f"\n✗ 失败: {e}")
        return {'id': staff_id, 'date': date, 'status': 'error', 'error': str(e)}

# ==================== 整合功能（基于完整原文） ====================

def integrate_weekly_reports():
    """整合所有周报，生成最终报告（基于完整原文）"""
    print("\n" + "="*60)
    print("整合周报...")
    
    reports = []
    for file in glob.glob('*_周报.json'):
        try:
            with open(file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                reports.append(data)
        except Exception as e:
            print(f"  警告: 无法读取 {file} - {e}")
    
    if not reports:
        print("  ✗ 未找到周报文件")
        return None
    
    print(f"  加载了 {len(reports)} 份周报")
    
    # 生成分隔的团队周报
    team_md = generate_team_weekly_md(reports)
    with open('team_weekly_report.md', 'w', encoding='utf-8') as f:
        f.write(team_md)
    print(f"  ✓ 团队周报: team_weekly_report.md")
    
    # 生成大模型提示词
    llm_prompt = generate_llm_prompt(reports)
    with open('llm_prompt.txt', 'w', encoding='utf-8') as f:
        f.write(llm_prompt)
    print(f"  ✓ 提示词: llm_prompt.txt")
    
    # 生成本人历史周报（用于参考）
    personal_md = generate_personal_reference(reports)
    with open('personal_reference.md', 'w', encoding='utf-8') as f:
        f.write(personal_md)
    print(f"  ✓ 个人参考: personal_reference.md")
    
    return {
        'team_md': team_md,
        'llm_prompt': llm_prompt,
        'reports': reports
    }

def generate_team_weekly_md(reports):
    """生成团队周报 Markdown（基于完整原文）"""
    lines = [
        "# 团队周报汇总",
        "",
        f"**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        f"**团队成员数**: {len(reports)}",
        "",
        "---",
        ""
    ]
    
    for report in reports:
        name = report.get('name', 'Unknown')
        date = report.get('date', 'Unknown')
        content = report.get('content', {})
        
        lines.extend([
            f"## {name}",
            "",
            f"**周期**: {date}",
            "",
            "### OKR进展",
            ""
        ])
        
        okr_content = content.get('okr', [])
        if okr_content:
            for obj in okr_content:
                obj_title = obj.get('title', '无')
                lines.append(f"- **{obj_title}**")
                for kr in obj.get('krs', []):
                    lines.append(f"  - {kr}")
        else:
            lines.append("- 暂无OKR数据")
        
        lines.extend([
            "",
            "### 本周工作总结",
            ""
        ])
        
        work_sections = content.get('work_summary', [])
        if work_sections:
            for section in work_sections:
                section_content = section.get('content', [])
                for item in section_content[:20]:  # 显示前20条
                    lines.append(f"- {item}")
        else:
            lines.append("- 暂无工作总结")
        
        lines.extend([
            "",
            "### 下周工作计划",
            ""
        ])
        
        plan_sections = content.get('next_week_plan', [])
        if plan_sections:
            for section in plan_sections:
                section_content = section.get('content', [])
                for item in section_content[:10]:  # 显示前10条
                    lines.append(f"- {item}")
        else:
            lines.append("- 暂无计划")
        
        lines.extend(["", "---", ""])
    
    return '\n'.join(lines)

def generate_llm_prompt(reports):
    """生成给大模型的提示词（基于完整原文）"""
    lines = [
        "基于以下团队成员的周报原文，请生成一份团队周报汇总：",
        "",
        "要求：",
        "1. 按人员分别总结每个人的本周工作",
        "2. 提取团队整体的关键进展",
        "3. 识别需要关注的风险或问题",
        "4. 格式清晰，便于阅读",
        "",
        "---",
        "",
        f"团队成员数: {len(reports)}",
        "",
        "团队数据（已过滤无效信息并去重）：",
        ""
    ]
    
    for report in reports:
        name = report.get('name', 'Unknown')
        content = report.get('content', {})
        
        lines.extend([f"## {name}", ""])
        
        lines.append("OKR进展：")
        for obj in content.get('okr', []):
            lines.append(f"- 目标: {obj.get('title', '')}")
            for kr in obj.get('krs', []):
                lines.append(f"  - {kr}")
        lines.append("")
        
        lines.append("本周工作：")
        work_sections = content.get('work_summary', [])
        for section in work_sections:
            for item in section.get('content', [])[:10]:  # 限制数量
                lines.append(f"- {item}")
        lines.append("")
    
    return '\n'.join(lines)

def generate_personal_reference(reports):
    """生成本人历史周报参考（基于完整原文）"""
    # 查找本人的周报
    my_reports = [r for r in reports if '本人' in r.get('desc', '')]
    
    lines = [
        "# 个人历史周报参考",
        "",
        f"**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "",
        "可用于撰写个人周报时参考历史内容",
        "",
        "---",
        ""
    ]
    
    for report in my_reports:
        date = report.get('date', 'Unknown')
        content = report.get('content', {})
        
        lines.extend([
            f"## {date}",
            "",
            "### OKR",
            ""
        ])
        
        for obj in content.get('okr', []):
            lines.append(f"- {obj.get('title', '')}")
        
        lines.extend([
            "",
            "### 工作总结",
            ""
        ])
        
        work_sections = content.get('work_summary', [])
        for section in work_sections:
            for item in section.get('content', []):
                lines.append(f"- {item}")
        
        lines.extend(["", "---", ""])
    
    return '\n'.join(lines)

# ==================== 主流程 ====================

async def main():
    last_week = get_week_date(0)
    week_before = get_week_date(1)
    
    print("="*60)
    print(" "*15 + "百度OKR周报抓取 - V5.2修正版")
    print("="*60)
    print("\n【特性说明】")
    print("  ✓ 从'汽车项目交付组-'后提取员工姓名")
    print("  ✓ 增加等待时间确保内容完整加载")
    print("  ✓ 保留完整原文（仅过滤无效信息+去重）")
    print("  ✓ 抓取+整合在一个脚本完成")
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
        
        # 抓取所有下属上周
        print(f"\n{'='*60}")
        print(f"抓取 {len(STAFF_IDS)} 个下属上周周报")
        print(f"{'='*60}")
        
        for i, staff_id in enumerate(STAFF_IDS, 1):
            result = await fetch_staff_v5(page, staff_id, last_week, f"下属{i}", i, len(STAFF_IDS))
            all_results.append(result)
            
            if result['status'] == 'login_required':
                print("\n⚠️ 会话失效，停止")
                break
            
            if i < len(STAFF_IDS):
                await asyncio.sleep(1.5)  # 减少间隔
        
        # 抓取本人历史2周
        if not any(r['status'] == 'login_required' for r in all_results):
            print(f"\n{'='*60}")
            print("抓取本人历史2周")
            print(f"{'='*60}")
            
            r1 = await fetch_staff_v5(page, MY_ID, last_week, "本人-上周", 1, 2)
            all_results.append(r1)
            
            if r1['status'] != 'login_required':
                await asyncio.sleep(2)  # 减少间隔
                r2 = await fetch_staff_v5(page, MY_ID, week_before, "本人-上上周", 2, 2)
                all_results.append(r2)
        
        await browser.close()
        
        # 保存抓取汇总
        summary = {
            'fetch_time': datetime.now().isoformat(),
            'my_id': MY_ID,
            'success': sum(1 for r in all_results if r['status'] == 'success'),
            'failed': sum(1 for r in all_results if r['status'] == 'error'),
            'results': [{'id': r.get('id'), 'name': r.get('name'), 'status': r.get('status')} for r in all_results]
        }
        
        with open('v5_fetch_summary.json', 'w', encoding='utf-8') as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)
        
        print(f"\n{'='*60}")
        print("抓取完成!")
        print(f"{'='*60}")
        print(f"  成功: {summary['success']}, 失败: {summary['failed']}")
        
        # 自动整合
        integrate_result = integrate_weekly_reports()
        
        if integrate_result:
            print(f"\n{'='*60}")
            print("全部完成!")
            print(f"{'='*60}")
            print(f"\n生成的文件:")
            print(f"  - 姓名_日期_周报.json (个人完整数据 - 已过滤+去重)")
            print(f"  - 姓名_日期_周报.png (截图)")
            print(f"  - team_weekly_report.md (团队周报 - 基于完整原文)")
            print(f"  - llm_prompt.txt (大模型提示词)")
            print(f"  - personal_reference.md (个人历史参考)")
            print(f"\n使用建议:")
            print(f"  1. 查看 team_weekly_report.md 了解团队整体情况")
            print(f"  2. 将 llm_prompt.txt 发送给大模型生成更完善的周报")
            print(f"  3. 参考 personal_reference.md 撰写个人周报")
        
        input("\n按Enter退出...")

if __name__ == '__main__':
    asyncio.run(main())
