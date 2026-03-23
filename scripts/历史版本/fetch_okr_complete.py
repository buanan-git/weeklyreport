#!/usr/bin/env python3
"""
百度OKR周报抓取 - 完整版
修复内容:
1. 本人周报读取历史2周（上周+上上周）
2. 增加滚动等待，确保OKR部分完全加载
3. 增加页面结构检查
"""

import json
import asyncio
from datetime import datetime, timedelta
from playwright.async_api import async_playwright

MY_ID = 's673090'
STAFF_IDS = [
    's801573', 's733504', 's682083', 's929270', 's798268',
    's811385', 's865585', 's783729', 's400192', 's823700'
]
BASE_URL = 'https://okr.baidu-int.com/pages/weekly.html#/index?mode=steam'

def get_week_date(weeks_ago=0):
    """获取指定周数的周一日期
    weeks_ago: 0=上周, 1=上上周, 2=三周前
    """
    today = datetime(2026, 2, 11)
    days_since_monday = today.weekday()
    # 上周一 = today - (days_since_monday + 7)
    # 上上周 = today - (days_since_monday + 14)
    target_monday = today - timedelta(days=days_since_monday + 7 + (weeks_ago * 7))
    return f"s{target_monday.strftime('%Y-%m-%d')}"

async def scroll_page(page):
    """滚动页面确保所有内容加载"""
    try:
        # 滚动到底部
        await page.evaluate('''() => {
            window.scrollTo(0, document.body.scrollHeight);
        }''')
        await asyncio.sleep(2)
        
        # 再滚动一下，确保动态内容加载
        await page.evaluate('''() => {
            window.scrollTo(0, document.body.scrollHeight / 2);
        }''')
        await asyncio.sleep(1)
        
        # 回到顶部
        await page.evaluate('''() => {
            window.scrollTo(0, 0);
        }''')
        await asyncio.sleep(1)
    except:
        pass

async def fetch_staff_weekly(page, staff_id, date, desc=""):
    """抓取单个员工指定周的周报
    desc: 描述（如"上周"、"上上周"）
    """
    url = f"{BASE_URL}&id={staff_id}&date={date}"
    prefix = f"[{desc}] " if desc else ""
    print(f"{prefix}{staff_id} ({date})", end=" ")
    
    try:
        # 访问页面
        await page.goto(url, wait_until='domcontentloaded', timeout=60000)
        await asyncio.sleep(3)  # 初始等待
        
        # 检查是否被重定向到登录页
        current_url = page.url
        if 'uuap' in current_url or 'login' in current_url:
            print(f"⚠️ 需登录")
            return {'id': staff_id, 'date': date, 'status': 'login_required'}
        
        # 滚动页面确保内容加载
        await scroll_page(page)
        
        # 额外等待，确保OKR部分加载
        print("等待OKR加载...", end=" ")
        await asyncio.sleep(5)  # 增加等待时间到总共约10秒
        
        # 再次滚动，确保懒加载内容
        await scroll_page(page)
        await asyncio.sleep(2)
        
        # 获取页面完整内容
        text = await page.evaluate('() => document.body.innerText')
        html = await page.content()
        
        # 检查内容是否包含OKR关键词
        has_okr = any(keyword in text.lower() for keyword in ['okr', 'objective', 'key result', 'o1', 'o2', 'o3'])
        has_kr = any(keyword in text.lower() for keyword in ['kr1', 'kr2', 'kr3'])
        
        content_preview = text[:500] if text else ""
        
        result = {
            'id': staff_id,
            'date': date,
            'date_desc': desc,
            'url': url,
            'fetch_time': datetime.now().isoformat(),
            'status': 'success',
            'has_okr_section': has_okr,
            'has_kr_section': has_kr,
            'content_length': len(text) if text else 0,
            'html_length': len(html) if html else 0,
            'content': text[:50000] if text else '',  # 增加保存长度
            'content_preview': content_preview
        }
        
        # 保存到文件
        filename = f"{staff_id}_{date}.json"
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        okr_status = "✓OKR" if has_okr else "✗无OKR"
        print(f"✓ ({len(text)}字符, {okr_status})")
        return result
        
    except Exception as e:
        print(f"✗ {str(e)[:40]}")
        return {
            'id': staff_id,
            'date': date,
            'status': 'error',
            'error': str(e)
        }

async def main():
    # 获取三个日期：上周、上上周、三周前
    last_week = get_week_date(0)      # s2026-02-02
    week_before_last = get_week_date(1)  # s2026-01-26
    three_weeks_ago = get_week_date(2)   # s2026-01-19
    
    print("="*60)
    print("百度OKR周报抓取 - 完整版")
    print(f"上周: {last_week}")
    print(f"上上周: {week_before_last}")
    print(f"三周前: {three_weeks_ago}")
    print("="*60)
    
    async with async_playwright() as p:
        print("\n启动Chrome...")
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(viewport={'width': 1920, 'height': 1080})
        page = await context.new_page()
        
        # 先访问OKR并登录
        print("\n访问OKR系统...")
        await page.goto("https://okr.baidu-int.com", wait_until='domcontentloaded')
        
        print("\n⚠️  请手动登录OKR系统")
        input("登录完成后按Enter...")
        
        # 验证登录
        if 'uuap' in page.url or 'login' in page.url:
            print("✗ 未登录成功")
            await browser.close()
            return
        
        all_results = []
        
        # ========== 第一部分：抓取所有下属上周周报 ==========
        print(f"\n{'='*60}")
        print(f"第一部分：抓取 {len(STAFF_IDS)} 个下属上周周报")
        print(f"{'='*60}")
        
        for i, staff_id in enumerate(STAFF_IDS, 1):
            result = await fetch_staff_weekly(page, staff_id, last_week, f"{i}/{len(STAFF_IDS)}")
            all_results.append(result)
            
            if result['status'] == 'login_required':
                print("\n⚠️  会话失效，停止抓取")
                break
            
            if i < len(STAFF_IDS):
                await asyncio.sleep(2)
        
        # ========== 第二部分：抓取本人历史2周 ==========
        if not any(r['status'] == 'login_required' for r in all_results):
            print(f"\n{'='*60}")
            print("第二部分：抓取本人历史2周周报")
            print(f"{'='*60}")
            
            # 上周
            result1 = await fetch_staff_weekly(page, MY_ID, last_week, "上周")
            all_results.append(result1)
            await asyncio.sleep(2)
            
            # 上上周
            result2 = await fetch_staff_weekly(page, MY_ID, week_before_last, "上上周")
            all_results.append(result2)
            await asyncio.sleep(2)
            
            # 可选：三周前
            # result3 = await fetch_staff_weekly(page, MY_ID, three_weeks_ago, "三周前")
            # all_results.append(result3)
        
        # ========== 保存汇总 ==========
        summary = {
            'fetch_time': datetime.now().isoformat(),
            'dates': {
                'last_week': last_week,
                'week_before_last': week_before_last
            },
            'my_id': MY_ID,
            'total': len(all_results),
            'success': sum(1 for r in all_results if r['status'] == 'success'),
            'failed': sum(1 for r in all_results if r['status'] == 'error'),
            'login_required': sum(1 for r in all_results if r['status'] == 'login_required'),
            'has_okr': sum(1 for r in all_results if r.get('has_okr_section')),
            'results': all_results
        }
        
        with open('all_results_complete.json', 'w', encoding='utf-8') as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)
        
        # ========== 输出统计 ==========
        print(f"\n{'='*60}")
        print("抓取完成统计")
        print(f"{'='*60}")
        print(f"总数量: {summary['total']}")
        print(f"成功: {summary['success']}")
        print(f"失败: {summary['failed']}")
        print(f"需登录: {summary['login_required']}")
        print(f"包含OKR部分: {summary['has_okr']}")
        print(f"\n文件保存在当前目录")
        print(f"汇总文件: all_results_complete.json")
        print(f"{'='*60}")
        
        await browser.close()
        input("\n按Enter退出")

if __name__ == '__main__':
    asyncio.run(main())
