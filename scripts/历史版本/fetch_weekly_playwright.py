#!/usr/bin/env python3
"""
百度OKR周报抓取脚本（Playwright版）
作为Browser Relay不稳定时的备选方案
"""

import json
import asyncio
from datetime import datetime, timedelta
from playwright.async_api import async_playwright

# 员工ID
MY_ID = 's673090'
STAFF_IDS = [
    's801573', 's733504', 's682083', 's929270', 's798268',
    's811385', 's865585', 's783729', 's400192', 's823700'
]

BASE_URL = 'https://okr.baidu-int.com/pages/weekly.html#/index?mode=steam'

def get_last_monday():
    """获取上周一日期"""
    today = datetime(2026, 2, 11)
    days_since_monday = today.weekday()
    last_monday = today - timedelta(days=days_since_monday + 7)
    return f"s{last_monday.strftime('%Y-%m-%d')}"

async def fetch_weekly_report(page, staff_id, date):
    """抓取单个员工的周报"""
    url = f"{BASE_URL}&id={staff_id}&date={date}"
    print(f"正在抓取: {staff_id} ...")
    
    try:
        await page.goto(url, wait_until='networkidle', timeout=30000)
        await asyncio.sleep(2)  # 等待页面渲染
        
        # 获取员工姓名
        name_elem = await page.query_selector('[class*="name"], .user-name, .staff-name')
        name = await name_elem.inner_text() if name_elem else staff_id
        
        # 获取工作总结部分
        work_summary = await page.evaluate('''() => {
            const sections = document.querySelectorAll('.weekly-section, .summary-section, [class*="summary"]');
            const summaries = [];
            sections.forEach(section => {
                const text = section.innerText?.trim();
                if (text && text.length > 10) summaries.push(text);
            });
            return summaries.join('\\n\\n');
        }''')
        
        # 获取OKR部分
        okr_content = await page.evaluate('''() => {
            const okrSections = document.querySelectorAll('.okr-section, .objective-item, [class*="okr"], [class*="objective"]');
            const okrs = [];
            okrSections.forEach(section => {
                const text = section.innerText?.trim();
                if (text && text.length > 5) okrs.push(text);
            });
            return okrs.join('\\n\\n');
        }''')
        
        # 获取完整页面文本作为备份
        full_text = await page.evaluate('() => document.body.innerText')
        
        return {
            'id': staff_id,
            'name': name.strip() if name else staff_id,
            'url': url,
            'work_summary': work_summary or '未找到工作总结',
            'okr_content': okr_content or '未找到OKR内容',
            'full_text': full_text[:5000] if full_text else '',  # 限制长度
            'status': 'success'
        }
    except Exception as e:
        print(f"  ✗ 抓取失败 {staff_id}: {str(e)}")
        return {
            'id': staff_id,
            'name': staff_id,
            'url': url,
            'work_summary': f'抓取失败: {str(e)}',
            'okr_content': '',
            'full_text': '',
            'status': 'error',
            'error': str(e)
        }

async def main():
    date = get_last_monday()
    print(f"="*60)
    print(f"百度OKR周报抓取")
    print(f"日期: {date}")
    print(f"="*60)
    print()
    
    results = []
    
    async with async_playwright() as p:
        # 连接到已打开的Chrome（通过CDP）
        # 或者启动新的浏览器
        try:
            browser = await p.chromium.connect_over_cdp('http://127.0.0.1:18792')
            print("✓ 已连接到Chrome浏览器")
        except:
            print("正在启动新浏览器...")
            browser = await p.chromium.launch(headless=False)
        
        context = browser.contexts[0] if browser.contexts else await browser.new_context()
        page = context.pages[0] if context.pages else await context.new_page()
        
        # 先访问OKR主页确保登录
        print("检查登录状态...")
        await page.goto('https://okr.baidu-int.com', timeout=30000)
        await asyncio.sleep(2)
        
        # 抓取所有下属周报
        print(f"\\n开始抓取 {len(STAFF_IDS)} 个下属的周报...")
        for i, staff_id in enumerate(STAFF_IDS, 1):
            print(f"\\n[{i}/{len(STAFF_IDS)}] ", end='')
            result = await fetch_weekly_report(page, staff_id, date)
            results.append(result)
            
            # 保存中间结果
            with open(f'staff_{staff_id}.json', 'w', encoding='utf-8') as f:
                json.dump(result, f, ensure_ascii=False, indent=2)
        
        # 抓取本人周报（历史2周）
        print(f"\\n\\n抓取本人周报...")
        my_result = await fetch_weekly_report(page, MY_ID, date)
        results.append(my_result)
        
        # 保存所有结果
        output = {
            'date': date,
            'fetch_time': datetime.now().isoformat(),
            'my_id': MY_ID,
            'staff_count': len(STAFF_IDS),
            'results': results
        }
        
        with open('weekly_reports_raw.json', 'w', encoding='utf-8') as f:
            json.dump(output, f, ensure_ascii=False, indent=2)
        
        print(f"\\n✓ 所有周报已保存到 weekly_reports_raw.json")
        
        await browser.close()

if __name__ == '__main__':
    asyncio.run(main())
