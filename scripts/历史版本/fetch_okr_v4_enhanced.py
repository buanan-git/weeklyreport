#!/usr/bin/env python3
"""
百度OKR周报抓取 - V4增强版
改进项:
1. 文件名使用员工姓名，并在周报显著标明
2. 本地提取关键内容，过滤无关信息
3. 本地初步摘要，减少大模型输入量
"""

import json
import asyncio
import re
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

# ==================== 内容提取与过滤 ====================

def extract_employee_name(full_text):
    """从页面文本中提取员工姓名"""
    lines = full_text.split('\n')
    
    # 策略1: 寻找包含"周报"、"周总结"附近的文本
    for i, line in enumerate(lines[:30]):  # 只看前30行
        line = line.strip()
        if any(keyword in line for keyword in ['周报', '周总结', 'Weekly', '的OKR']):
            # 向前向后查找名字
            for j in range(max(0, i-3), min(len(lines), i+3)):
                candidate = lines[j].strip()
                # 名字特征: 2-4个中文字符，不含标点
                if re.match(r'^[\u4e00-\u9fa5]{2,4}$', candidate):
                    if candidate not in ['周报', '总结', '本周', '下周', '上周']:
                        return candidate
    
    # 策略2: 寻找特定模式
    patterns = [
        r'([\u4e00-\u9fa5]{2,4})的周报',
        r'([\u4e00-\u9fa5]{2,4})的周总结',
        r'([\u4e00-\u9fa5]{2,4})的OKR',
        r'员工[:\s]*([\u4e00-\u9fa5]{2,4})',
        r'姓名[:\s]*([\u4e00-\u9fa5]{2,4})',
    ]
    for pattern in patterns:
        match = re.search(pattern, full_text)
        if match:
            return match.group(1)
    
    # 策略3: 寻找独立的中文人名（2-4个字）
    for line in lines[:20]:
        line = line.strip()
        if re.match(r'^[\u4e00-\u9fa5]{2,4}$', line):
            # 排除常见非人名词
            exclude = ['百度', '周报', '总结', '本周', '下周', '上周', '提交', 
                      '保存', '取消', '编辑', '查看', '更多', '设置', '帮助',
                      '首页', '目标', '绩效', '人才', '考核', '系统', '管理']
            if line not in exclude:
                return line
    
    return None

def extract_key_content(full_text, html_content=""):
    """
    提取关键内容，过滤无关信息
    返回结构化的周报内容
    """
    result = {
        'employee_name': None,
        'week_period': None,
        'okr_content': [],
        'work_summary': [],
        'next_week_plan': [],
        'other_notes': [],
        'raw_filtered': ""
    }
    
    # 提取姓名
    result['employee_name'] = extract_employee_name(full_text)
    
    # 提取时间段
    date_pattern = r'(20\d{2})[-./年](\d{1,2})[-./月](\d{1,2})'
    dates = re.findall(date_pattern, full_text)
    if len(dates) >= 2:
        result['week_period'] = f"{dates[0][0]}-{dates[0][1]:0>2}-{dates[0][2]:0>2} 至 {dates[1][0]}-{dates[1][1]:0>2}-{dates[1][2]:0>2}"
    
    lines = full_text.split('\n')
    lines = [l.strip() for l in lines if l.strip()]
    
    # 过滤掉导航、菜单等无关内容
    noise_patterns = [
        r'^百度$',
        r'^OKR系统$',
        r'^首页$',
        r'^目标$',
        r'^绩效$',
        r'^人才$',
        r'^考核$',
        r'^管理$',
        r'^设置$',
        r'^帮助$',
        r'^退出$',
        r'^搜索$',
        r'^消息$',
        r'^通知$',
        r'^个人中心$',
        r'^工作台$',
        r'^更多$',
        r'^菜单$',
        r'^导航$',
        r'^Copyright',
        r'^©',
        r'^版权所有',
        r'^百度在线网络技术',
        r'^京ICP',
        r'^隐私政策',
        r'^用户协议',
        r'^联系我们$',
        r'^反馈$',
        r'^加载中',
        r'^Loading',
        r'^请稍候',
        r'^确定$',
        r'^取消$',
        r'^保存$',
        r'^提交$',
        r'^编辑$',
        r'^删除$',
        r'^查看$',
        r'^展开$',
        r'^收起$',
    ]
    
    filtered_lines = []
    for line in lines:
        is_noise = False
        for pattern in noise_patterns:
            if re.match(pattern, line, re.IGNORECASE):
                is_noise = True
                break
        # 过滤过短的行（可能是按钮或标签）
        if len(line) <= 2 and not re.match(r'^[O o]\d+$', line) and not re.match(r'^[K k][R r]\d+$', line):
            is_noise = True
        if not is_noise:
            filtered_lines.append(line)
    
    # 识别OKR部分
    current_section = None
    okr_pattern = re.compile(r'^[Oo]\d+[:：\s]*(.*)', re.IGNORECASE)
    kr_pattern = re.compile(r'^[Kk][Rr]\d+[:：\s]*(.*)', re.IGNORECASE)
    
    for line in filtered_lines:
        # 检测O（目标）
        okr_match = okr_pattern.match(line)
        if okr_match:
            current_section = {'type': 'objective', 'title': line, 'krs': []}
            result['okr_content'].append(current_section)
            continue
        
        # 检测KR（关键结果）
        kr_match = kr_pattern.match(line)
        if kr_match and current_section and current_section['type'] == 'objective':
            current_section['krs'].append(line)
            continue
        
        # 检测工作总结部分
        if any(keyword in line for keyword in ['本周工作总结', '本周工作', '工作总结', 'Work Summary', '本周进展']):
            current_section = {'type': 'work_summary', 'content': []}
            result['work_summary'].append(current_section)
            continue
        
        # 检测下周计划
        if any(keyword in line for keyword in ['下周计划', 'Next Week', '下周工作', '后续计划']):
            current_section = {'type': 'next_week', 'content': []}
            result['next_week_plan'].append(current_section)
            continue
        
        # 收集当前章节的内容
        if current_section:
            if isinstance(current_section.get('content'), list):
                current_section['content'].append(line)
    
    # 如果没有识别出OKR结构，尝试文本分析
    if not result['okr_content']:
        # 寻找可能的目标描述（包含"目标"、"Objective"等关键词的段落）
        for i, line in enumerate(filtered_lines):
            if any(keyword in line for keyword in ['目标', 'Objective', 'O1', 'O2', 'O3']):
                if len(line) > 10:  # 有一定长度的描述
                    result['okr_content'].append({
                        'type': 'objective',
                        'title': line,
                        'krs': []
                    })
    
    # 构建过滤后的原始文本
    result['raw_filtered'] = '\n'.join(filtered_lines)
    
    return result

def generate_local_summary(extracted_content):
    """
    在本地生成初步摘要
    提取关键信息，减少后续给大模型的数据量
    """
    summary = {
        'employee_name': extracted_content.get('employee_name', '未知'),
        'week_period': extracted_content.get('week_period', '未知'),
        'okr_summary': [],
        'work_highlights': [],
        'next_week_focus': [],
        'total_length': len(extracted_content.get('raw_filtered', ''))
    }
    
    # 摘要OKR
    for obj in extracted_content.get('okr_content', []):
        obj_summary = {
            'objective': obj.get('title', '')[:100],  # 限制长度
            'kr_count': len(obj.get('krs', [])),
            'key_results': [kr[:150] for kr in obj.get('krs', [])[:5]]  # 最多5个KR，每个限制长度
        }
        summary['okr_summary'].append(obj_summary)
    
    # 摘要工作总结（提取关键句子）
    work_content = []
    for section in extracted_content.get('work_summary', []):
        work_content.extend(section.get('content', []))
    
    # 简单的关键词提取：寻找包含数字、完成、上线、发布等词的句子
    important_keywords = ['完成', '上线', '发布', '开发', '测试', '修复', '优化', '设计', '评审', 
                         '会议', '对接', '协调', '推进', '解决', '实现', '%', '个', '次', '项']
    for line in work_content[:20]:  # 只看前20行
        if any(kw in line for kw in important_keywords) and len(line) > 10:
            summary['work_highlights'].append(line[:200])  # 限制长度
    
    # 摘要下周计划
    plan_content = []
    for section in extracted_content.get('next_week_plan', []):
        plan_content.extend(section.get('content', []))
    
    for line in plan_content[:10]:
        if len(line) > 5:
            summary['next_week_focus'].append(line[:200])
    
    return summary

def sanitize_filename(name):
    """将姓名转换为安全的文件名"""
    if not name:
        return "unknown"
    # 移除或替换不安全字符
    safe = re.sub(r'[\\/*?:"<>|]', '_', name)
    safe = safe.strip()
    return safe if safe else "unknown"

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
                result.push({
                    class: el.className,
                    id: el.id,
                    tag: el.tagName,
                    scrollHeight: el.scrollHeight,
                    clientHeight: el.clientHeight
                });
            }
        }
        result.sort((a, b) => b.scrollHeight - a.scrollHeight);
        return result.slice(0, 3);
    }''')
    
    if containers:
        first = containers[0]
        if first.get('class'):
            selector = f"[class*='{first['class'].split()[0]}']"
        elif first.get('id'):
            selector = f"#{first['id']}"
        else:
            selector = first['tag'].lower()
        return selector
    return 'body'

async def scroll_container(page, selector):
    """在容器内滚动"""
    center = await page.evaluate(f'''
        () => {{
            const el = document.querySelector("{selector}");
            if (el) {{
                const rect = el.getBoundingClientRect();
                return {{
                    x: rect.left + rect.width / 2,
                    y: rect.top + rect.height / 2
                }};
            }}
            return {{x: 960, y: 600}};
        }}
    ''')
    
    await page.mouse.move(center['x'], center['y'])
    await asyncio.sleep(0.5)
    
    for i in range(30):
        await page.mouse.wheel(0, 400)
        await asyncio.sleep(0.3)
    
    await page.click(selector)
    await asyncio.sleep(0.3)
    for _ in range(5):
        await page.keyboard.press('PageDown')
        await asyncio.sleep(0.3)
    
    await asyncio.sleep(3)

# ==================== 主流程 ====================

async def fetch_staff_enhanced(page, staff_id, date, desc, index=0, total=0):
    """增强版抓取 - 包含内容提取和本地摘要"""
    url = f"{BASE_URL}&id={staff_id}&date={date}"
    progress = f"[{index}/{total}] " if index else ""
    
    print(f"\n{'='*70}")
    print(f"{progress}抓取: {desc}")
    print(f"日期: {date}")
    print(f"{'='*70}")
    
    try:
        # 1. 打开页面
        print("\n[步骤1/6] 打开页面...")
        await page.goto(url, wait_until='domcontentloaded', timeout=60000)
        await asyncio.sleep(5)
        
        if 'uuap' in page.url:
            print("⚠️ 需要重新登录")
            return {'id': staff_id, 'date': date, 'status': 'login_required'}
        
        print("✓ 页面已打开")
        
        # 2. 等待渲染
        print("\n[步骤2/6] 等待初始渲染...")
        await asyncio.sleep(5)
        
        # 3. 滚动加载
        print("\n[步骤3/6] 滚动加载内容...")
        container = await find_scroll_container(page)
        await scroll_container(page, container)
        
        # 4. 等待动态内容
        print("\n[步骤4/6] 等待动态内容稳定...")
        await asyncio.sleep(5)
        
        # 5. 提取页面内容
        print("\n[步骤5/6] 提取页面内容...")
        full_text = await page.evaluate('() => document.body.innerText')
        html_content = await page.content()
        print(f"  原始内容: {len(full_text)} 字符")
        
        # 6. 本地提取关键内容
        print("\n[步骤6/6] 本地提取关键内容...")
        extracted = extract_key_content(full_text, html_content)
        print(f"  识别姓名: {extracted.get('employee_name') or '未识别'}")
        print(f"  OKR目标数: {len(extracted.get('okr_content', []))}")
        print(f"  工作总结段落: {len(extracted.get('work_summary', []))}")
        
        # 7. 生成本地摘要
        print("\n  生成本地摘要...")
        local_summary = generate_local_summary(extracted)
        print(f"  摘要后数据量: {len(json.dumps(local_summary))} 字符")
        print(f"  压缩比: {len(json.dumps(local_summary)) / len(full_text) * 100:.1f}%")
        
        # 确定文件名（使用姓名）
        employee_name = extracted.get('employee_name') or staff_id
        safe_name = sanitize_filename(employee_name)
        
        # 8. 保存结果
        result = {
            'id': staff_id,
            'name': employee_name,
            'date': date,
            'desc': desc,
            'fetch_time': datetime.now().isoformat(),
            'status': 'success',
            'extracted_content': extracted,
            'local_summary': local_summary,
            'original_length': len(full_text),
            'filtered_length': len(extracted.get('raw_filtered', '')),
            'has_okr': len(extracted.get('okr_content', [])) > 0
        }
        
        # 文件名使用姓名
        json_file = f'{safe_name}_{date}_周报.json'
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        # 同时保存一份纯摘要版本（供后续使用）
        summary_file = f'{safe_name}_{date}_摘要.json'
        with open(summary_file, 'w', encoding='utf-8') as f:
            json.dump({
                'name': employee_name,
                'date': date,
                'summary': local_summary,
                'okr_summary': local_summary.get('okr_summary', []),
                'work_highlights': local_summary.get('work_highlights', []),
                'next_week_focus': local_summary.get('next_week_focus', [])
            }, f, ensure_ascii=False, indent=2)
        
        # 截图
        screenshot = f'{safe_name}_{date}_周报.png'
        await page.screenshot(path=screenshot, full_page=True)
        
        print(f"\n{'='*70}")
        print(f"✓ 抓取成功!")
        print(f"  员工姓名: {employee_name}")
        print(f"  完整数据: {json_file}")
        print(f"  摘要数据: {summary_file}")
        print(f"  截图: {screenshot}")
        print(f"  数据压缩: {local_summary.get('total_length', 0)} -> {len(json.dumps(local_summary))} 字符")
        print(f"{'='*70}")
        
        return result
        
    except Exception as e:
        print(f"\n✗ 抓取失败: {e}")
        import traceback
        traceback.print_exc()
        return {'id': staff_id, 'date': date, 'status': 'error', 'error': str(e)}

async def main():
    last_week = get_week_date(0)
    week_before = get_week_date(1)
    
    print("="*70)
    print(" "*15 + "百度OKR周报抓取 - V4增强版")
    print("="*70)
    print("\n【改进特性】")
    print("  ✓ 文件名使用员工姓名")
    print("  ✓ 本地提取关键内容，过滤无关信息")
    print("  ✓ 本地生成摘要，减少大模型输入量")
    print("\n" + "="*70)
    print(f"员工ID: {MY_ID}")
    print(f"下属数量: {len(STAFF_IDS)}")
    print(f"上周: {last_week}")
    print(f"上上周: {week_before}")
    
    async with async_playwright() as p:
        print("\n" + "="*70)
        print("[系统] 启动Chrome...")
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page(viewport={'width': 1920, 'height': 1080})
        print("✓ Chrome已启动")
        
        print("\n[系统] 访问OKR...")
        await page.goto("https://okr.baidu-int.com", wait_until='domcontentloaded')
        print("✓ 页面已加载")
        
        print("\n" + "="*70)
        print("⚠️  请先登录OKR系统")
        print("="*70)
        input("\n登录完成后按 Enter...")
        
        if 'uuap' in page.url:
            print("\n✗ 未登录成功")
            return
        
        print("✓ 登录验证通过")
        
        all_results = []
        
        # 抓取所有下属上周
        print(f"\n{'='*70}")
        print(f"开始抓取 {len(STAFF_IDS)} 个下属上周周报")
        print(f"{'='*70}")
        
        for i, staff_id in enumerate(STAFF_IDS, 1):
            result = await fetch_staff_enhanced(page, staff_id, last_week, f"下属{i}", i, len(STAFF_IDS))
            all_results.append(result)
            
            if result['status'] == 'login_required':
                print("\n⚠️ 会话失效，停止")
                break
            
            if i < len(STAFF_IDS):
                await asyncio.sleep(2)
        
        # 抓取本人历史2周
        if not any(r['status'] == 'login_required' for r in all_results):
            print(f"\n{'='*70}")
            print("抓取本人历史2周")
            print(f"{'='*70}")
            
            r1 = await fetch_staff_enhanced(page, MY_ID, last_week, "本人-上周", 1, 2)
            all_results.append(r1)
            
            if r1['status'] != 'login_required':
                await asyncio.sleep(3)
                r2 = await fetch_staff_enhanced(page, MY_ID, week_before, "本人-上上周", 2, 2)
                all_results.append(r2)
        
        # 汇总
        summary = {
            'fetch_time': datetime.now().isoformat(),
            'my_id': MY_ID,
            'staff_count': len(STAFF_IDS),
            'success': sum(1 for r in all_results if r['status'] == 'success'),
            'failed': sum(1 for r in all_results if r['status'] == 'error'),
            'has_okr': sum(1 for r in all_results if r.get('has_okr')),
            'results': [{
                'id': r.get('id'),
                'name': r.get('name'),
                'status': r.get('status'),
                'has_okr': r.get('has_okr'),
                'original_length': r.get('original_length'),
                'filtered_length': r.get('filtered_length')
            } for r in all_results]
        }
        
        with open('v4_fetch_summary.json', 'w', encoding='utf-8') as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)
        
        print("\n" + "="*70)
        print("全部完成!")
        print("="*70)
        print(f"\n统计:")
        print(f"  总数: {len(all_results)}")
        print(f"  成功: {summary['success']}")
        print(f"  失败: {summary['failed']}")
        print(f"  包含OKR: {summary['has_okr']}")
        print(f"\n生成的文件:")
        print(f"  - {safe_name}_*_周报.json (完整提取数据)")
        print(f"  - {safe_name}_*_摘要.json (精简摘要，用于大模型)")
        print(f"  - {safe_name}_*_周报.png (截图)")
        print(f"  - v4_fetch_summary.json (汇总)")
        
        await browser.close()
        input("\n按Enter退出...")

if __name__ == '__main__':
    asyncio.run(main())
