#!/usr/bin/env python3
"""
百度OKR周报抓取脚本 - Windows版本
在PowerShell中运行: python fetch_okr.py
"""

import json
import asyncio
from datetime import datetime, timedelta
from playwright.async_api import async_playwright

# 员工ID配置
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

async def fetch_single_staff(page, staff_id, date, index, total):
    """抓取单个员工的周报"""
    url = f"{BASE_URL}&id={staff_id}&date={date}"
    print(f"\n[{index}/{total}] 抓取: {staff_id}")
    print(f"  URL: {url}")
    
    try:
        # 访问页面
        await page.goto(url, wait_until='networkidle', timeout=30000)
        await asyncio.sleep(3)  # 等待页面渲染
        
        # 获取员工姓名
        name = staff_id
        try:
            name_elem = await page.query_selector('[class*="name"], .user-name, .staff-name')
            if name_elem:
                name_text = await name_elem.inner_text()
                if name_text and len(name_text.strip()) > 0:
                    name = name_text.strip()
        except:
            pass
        
        print(f"  姓名: {name}")
        
        # 获取页面完整文本
        full_text = await page.evaluate('() => document.body.innerText')
        
        # 保存结果
        result = {
            'id': staff_id,
            'name': name,
            'url': url,
            'fetch_time': datetime.now().isoformat(),
            'status': 'success',
            'content': full_text[:20000] if full_text else ''  # 限制长度
        }
        
        # 保存到文件
        filename = f'staff_{staff_id}.json'
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"  ✓ 已保存: {filename} (内容长度: {len(full_text) if full_text else 0})")
        return result
        
    except Exception as e:
        print(f"  ✗ 抓取失败: {e}")
        return {
            'id': staff_id,
            'name': staff_id,
            'url': url,
            'fetch_time': datetime.now().isoformat(),
            'status': 'error',
            'error': str(e)
        }

async def main():
    date = get_last_monday()
    
    print("="*60)
    print("百度OKR周报抓取 - Windows Playwright")
    print(f"日期: {date}")
    print("="*60)
    print("\n⚠️  请确保:")
    print("  1. 已连接公司VPN")
    print("  2. Chrome浏览器已安装")
    print("  3. 已在Chrome中登录OKR系统")
    print("\n按 Enter 开始抓取...")
    input()
    
    results = []
    
    async with async_playwright() as p:
        # 启动Chrome
        print("\n启动Chrome浏览器...")
        browser = await p.chromium.launch(
            headless=False,  # 可见模式，便于调试
            args=['--disable-gpu']
        )
        
        # 创建新页面
        context = await browser.new_context(viewport={'width': 1920, 'height': 1080})
        page = await context.new_page()
        
        print(f"\n开始抓取 {len(STAFF_IDS)} 个下属...")
        
        # 逐个抓取下属
        for i, staff_id in enumerate(STAFF_IDS, 1):
            result = await fetch_single_staff(page, staff_id, date, i, len(STAFF_IDS))
            results.append(result)
            
            # 间隔2秒
            if i < len(STAFF_IDS):
                await asyncio.sleep(2)
        
        # 抓取本人
        print(f"\n抓取本人周报...")
        my_result = await fetch_single_staff(page, MY_ID, date, len(STAFF_IDS)+1, len(STAFF_IDS)+1)
        results.append(my_result)
        
        # 保存汇总
        summary = {
            'date': date,
            'fetch_time': datetime.now().isoformat(),
            'my_id': MY_ID,
            'staff_count': len(STAFF_IDS),
            'success_count': sum(1 for r in results if r['status'] == 'success'),
            'error_count': sum(1 for r in results if r['status'] == 'error'),
            'results': results
        }
        
        with open('all_results.json', 'w', encoding='utf-8') as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)
        
        print(f"\n{'='*60}")
        print("抓取完成!")
        print(f"成功: {summary['success_count']}, 失败: {summary['error_count']}")
        print(f"文件保存在: {os.getcwd()}")
        print(f"{'='*60}")
        
        await browser.close()
        
        print("\n按 Enter 退出...")
        input()

if __name__ == '__main__':
    import os
    asyncio.run(main())
