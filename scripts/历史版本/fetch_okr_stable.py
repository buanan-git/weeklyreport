#!/usr/bin/env python3
"""
百度OKR周报抓取 - 最终稳定版（含重试机制）
特性：
1. 记录所有失败页面
2. 全部抓取完成后自动重试（最多1次）
3. 详细的日志和统计
"""

import json
import asyncio
import socket
from datetime import datetime, timedelta
from playwright.async_api import async_playwright

# 配置
MY_ID = 's673090'
STAFF_IDS = [
    's801573', 's733504', 's682083', 's929270', 's798268',
    's811385', 's865585', 's783729', 's400192', 's823700'
]
BASE_URL = 'https://okr.baidu-int.com/pages/weekly.html#/index?mode=steam'
OKR_HOME = 'https://okr.baidu-int.com'
MAX_RETRIES = 1  # 最大重试次数

def get_week_date(weeks_ago=0):
    """获取周日期"""
    today = datetime(2026, 2, 11)
    target = today - timedelta(days=today.weekday() + 7 + (weeks_ago * 7))
    return f"s{target.strftime('%Y-%m-%d')}"

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
                    scrollHeight: el.scrollHeight
                });
            }
        }
        result.sort((a, b) => b.scrollHeight - a.scrollHeight);
        return result.slice(0, 3);
    }''')
    
    if containers:
        first = containers[0]
        if first.get('class'):
            return f"[class*='{first['class'].split()[0]}']"
        elif first.get('id'):
            return f"#{first['id']}"
        else:
            return first['tag'].lower()
    
    return 'body'

async def scroll_container(page, selector):
    """滚动容器"""
    center = await page.evaluate(f'''
        () => {{
            const el = document.querySelector("{selector}");
            if (el) {{
                const rect = el.getBoundingClientRect();
                return {{x: rect.left + rect.width / 2, y: rect.top + rect.height / 2}};
            }}
            return {{x: 960, y: 600}};
        }}
    ''')
    
    await page.mouse.move(center['x'], center['y'])
    await asyncio.sleep(0.5)
    
    for i in range(30):
        await page.mouse.wheel(0, 400)
        await asyncio.sleep(0.3)
    
    # 键盘确认
    await page.click(selector)
    await asyncio.sleep(0.3)
    for _ in range(5):
        await page.keyboard.press('PageDown')
        await asyncio.sleep(0.3)
    
    await asyncio.sleep(3)

async def fetch_single_staff(page, staff_id, date, desc, is_retry=False):
    """
    抓取单个员工周报
    返回: (result_dict, success_bool)
    """
    url = f"{BASE_URL}&id={staff_id}&date={date}"
    retry_mark = "[重试] " if is_retry else ""
    
    print(f"\n{'='*70}")
    print(f"{retry_mark}抓取: {desc}")
    print(f"员工ID: {staff_id}, 日期: {date}")
    print(f"{'='*70}")
    
    try:
        # 1. 打开页面
        print("[1/5] 打开页面...")
        await page.goto(url, wait_until='domcontentloaded', timeout=60000)
        await asyncio.sleep(5)
        
        if 'uuap' in page.url:
            print("⚠️ 需要重新登录")
            return {
                'id': staff_id,
                'date': date,
                'desc': desc,
                'status': 'login_required',
                'error': '页面跳转到登录页'
            }, False
        
        print("✓ 页面已打开")
        
        # 2. 等待渲染
        print("[2/5] 等待渲染...")
        await asyncio.sleep(5)
        
        # 3. 查找滚动容器
        print("[3/5] 查找滚动容器...")
        container = await find_scroll_container(page)
        print(f"✓ 使用: {container[:50]}")
        
        # 4. 滚动加载
        print("[4/5] 滚动加载内容...")
        await scroll_container(page, container)
        
        # 5. 等待动态内容
        print("[5/5] 等待动态内容...")
        await asyncio.sleep(5)
        
        # 6. 提取内容
        print("提取页面内容...")
        full_text = await page.evaluate('() => document.body.innerText')
        print(f"✓ 获取 {len(full_text)} 字符")
        
        # 7. 截图
        screenshot = f'{staff_id}_{date}.png'
        await page.screenshot(path=screenshot, full_page=True)
        print(f"✓ 截图: {screenshot}")
        
        # 8. 保存
        result = {
            'id': staff_id,
            'date': date,
            'desc': desc,
            'container': container,
            'fetch_time': datetime.now().isoformat(),
            'status': 'success',
            'is_retry': is_retry,
            'content': full_text[:50000],
            'screenshot': screenshot,
            'has_okr': any(k in full_text.lower() for k in ['okr', 'objective', 'o1', 'kr1']),
            'content_length': len(full_text)
        }
        
        json_file = f'{staff_id}_{date}.json'
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"\n✓ 抓取成功!")
        print(f"  JSON: {json_file}")
        print(f"  包含OKR: {'是' if result['has_okr'] else '否'}")
        
        return result, True
        
    except Exception as e:
        error_msg = str(e)
        print(f"\n✗ 抓取失败: {error_msg}")
        
        return {
            'id': staff_id,
            'date': date,
            'desc': desc,
            'status': 'error',
            'is_retry': is_retry,
            'error': error_msg,
            'fetch_time': datetime.now().isoformat()
        }, False

async def fetch_with_retry(page, staff_id, date, desc):
    """
    抓取员工（含重试逻辑）
    先尝试一次，失败则记录，不立即重试
    """
    result, success = await fetch_single_staff(page, staff_id, date, desc, is_retry=False)
    
    if success:
        return result, None  # 成功，无失败记录
    else:
        # 失败，返回失败信息用于后续重试
        return result, {
            'staff_id': staff_id,
            'date': date,
            'desc': desc,
            'error': result.get('error', '未知错误')
        }

async def process_failed_tasks(page, failed_tasks):
    """处理失败的任务（重试）"""
    if not failed_tasks:
        print("\n没有失败的任务需要重试")
        return [], []
    
    print(f"\n{'='*70}")
    print(f"开始重试 {len(failed_tasks)} 个失败任务")
    print(f"{'='*70}")
    
    retry_results = []
    still_failed = []
    
    for i, task in enumerate(failed_tasks, 1):
        print(f"\n[重试 {i}/{len(failed_tasks)}]")
        print(f"员工: {task['staff_id']}, 日期: {task['date']}")
        print(f"之前错误: {task['error']}")
        
        result, success = await fetch_single_staff(
            page, 
            task['staff_id'], 
            task['date'], 
            task['desc'],
            is_retry=True
        )
        
        retry_results.append(result)
        
        if success:
            print("✓ 重试成功!")
        else:
            print("✗ 重试仍然失败")
            still_failed.append({
                'staff_id': task['staff_id'],
                'date': task['date'],
                'desc': task['desc'],
                'first_error': task['error'],
                'retry_error': result.get('error', '未知错误')
            })
        
        if i < len(failed_tasks):
            await asyncio.sleep(3)
    
    return retry_results, still_failed

async def main():
    last_week = get_week_date(0)
    week_before = get_week_date(1)
    
    print("="*70)
    print(" "*20 + "百度OKR周报抓取")
    print(" "*18 + "【稳定版-含重试】")
    print("="*70)
    print(f"\n员工ID: {MY_ID}")
    print(f"下属数量: {len(STAFF_IDS)}")
    print(f"上周: {last_week}")
    print(f"上上周: {week_before}")
    print(f"\n特性:")
    print(f"  - 失败自动记录")
    print(f"  - 全部完成后重试（最多{MAX_RETRIES}次）")
    print(f"  - 详细统计报告")
    
    async with async_playwright() as p:
        print("\n" + "="*70)
        print("[系统] 启动Chrome...")
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page(viewport={'width': 1920, 'height': 1080})
        print("✓ Chrome已启动")
        
        # 访问OKR（带错误处理）
        print("\n[系统] 访问OKR...")
        try:
            await page.goto(OKR_HOME, wait_until='domcontentloaded', timeout=30000)
            print("✓ OKR页面已加载")
        except Exception as e:
            print(f"⚠️  访问失败: {e}")
            print("\n请手动在Chrome中打开 https://okr.baidu-int.com")
            input("完成后按Enter...")
        
        print("\n" + "="*70)
        print("请登录OKR系统")
        print("="*70)
        input("\n登录完成后按Enter...")
        
        # 验证
        current_url = page.url
        if 'uuap' in current_url or 'login' in current_url:
            print("\n✗ 未登录成功")
            await browser.close()
            return
        
        print("✓ 登录验证通过")
        
        # ========== 第一轮：抓取所有下属 ==========
        print(f"\n{'='*70}")
        print(f"第一轮：抓取 {len(STAFF_IDS)} 个下属")
        print(f"{'='*70}")
        
        all_results = []
        failed_tasks = []
        
        for i, staff_id in enumerate(STAFF_IDS, 1):
            result, failed_info = await fetch_with_retry(
                page, staff_id, last_week, f"下属{i}"
            )
            
            all_results.append(result)
            
            if failed_info:
                failed_tasks.append(failed_info)
                print(f"  → 已记录失败，稍后重试")
            
            if i < len(STAFF_IDS):
                await asyncio.sleep(2)
        
        # ========== 第二轮：抓取本人 ==========
        print(f"\n{'='*70}")
        print("第二轮：抓取本人历史2周")
        print(f"{'='*70}")
        
        # 上周
        result1, failed1 = await fetch_with_retry(page, MY_ID, last_week, "本人-上周")
        all_results.append(result1)
        if failed1:
            failed_tasks.append(failed1)
        
        await asyncio.sleep(3)
        
        # 上上周
        result2, failed2 = await fetch_with_retry(page, MY_ID, week_before, "本人-上上周")
        all_results.append(result2)
        if failed2:
            failed_tasks.append(failed2)
        
        # ========== 第三轮：重试失败任务 ==========
        if failed_tasks:
            retry_results, still_failed = await process_failed_tasks(page, failed_tasks)
            all_results.extend(retry_results)
        else:
            still_failed = []
        
        # ========== 最终统计 ==========
        success_first = sum(1 for r in all_results if r.get('status') == 'success' and not r.get('is_retry'))
        success_retry = sum(1 for r in all_results if r.get('status') == 'success' and r.get('is_retry'))
        total_failed = len(still_failed)
        has_okr = sum(1 for r in all_results if r.get('has_okr'))
        
        summary = {
            'fetch_time': datetime.now().isoformat(),
            'my_id': MY_ID,
            'dates': {'last_week': last_week, 'week_before': week_before},
            'statistics': {
                'total_attempted': len(STAFF_IDS) + 2,  # 下属 + 本人2周
                'success_first_try': success_first,
                'success_after_retry': success_retry,
                'total_success': success_first + success_retry,
                'total_failed': total_failed,
                'has_okr_content': has_okr
            },
            'failed_tasks': still_failed,
            'all_results': all_results
        }
        
        with open('final_report.json', 'w', encoding='utf-8') as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)
        
        # ========== 输出报告 ==========
        print("\n" + "="*70)
        print("最终报告")
        print("="*70)
        print(f"\n抓取统计:")
        print(f"  总任务数: {len(STAFF_IDS) + 2}")
        print(f"  首次成功: {success_first}")
        print(f"  重试成功: {success_retry}")
        print(f"  总成功: {success_first + success_retry}")
        print(f"  最终失败: {total_failed}")
        print(f"  包含OKR: {has_okr}")
        
        if still_failed:
            print(f"\n最终失败的页面:")
            for i, task in enumerate(still_failed, 1):
                print(f"  {i}. {task['desc']} (ID: {task['staff_id']}, 日期: {task['date']})")
                print(f"     首次错误: {task['first_error'][:50]}")
                print(f"     重试错误: {task['retry_error'][:50]}")
        else:
            print(f"\n✓ 所有页面抓取成功!")
        
        print(f"\n生成的文件:")
        print(f"  - staff_*.json / staff_*.png (各员工数据)")
        print(f"  - final_report.json (完整报告)")
        
        await browser.close()
        input("\n按Enter退出...")

if __name__ == '__main__':
    asyncio.run(main())
