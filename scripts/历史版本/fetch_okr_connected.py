#!/usr/bin/env python3
"""
百度OKR周报抓取脚本 - Windows版本（连接已运行的Chrome）
修复问题:
1. 连接已运行的Chrome（而不是启动新窗口），保持登录状态
2. 增加页面加载等待时间（8秒）

在PowerShell中运行: python fetch_okr_connected.py
"""

import json
import asyncio
import subprocess
import os
import time
import socket
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
    
    try:
        print("  加载页面...")
        await page.goto(url, wait_until='networkidle', timeout=60000)
        print("  等待渲染(8秒)...")
        await asyncio.sleep(8)  # 充分等待页面加载
        
        # 获取姓名
        name = staff_id
        try:
            elem = await page.query_selector('[class*="name"], .user-name')
            if elem:
                text = await elem.inner_text()
                if text:
                    name = text.strip()
        except:
            pass
        
        print(f"  姓名: {name}")
        
        # 获取内容
        full_text = await page.evaluate('() => document.body.innerText')
        
        result = {
            'id': staff_id,
            'name': name,
            'url': url,
            'fetch_time': datetime.now().isoformat(),
            'status': 'success',
            'content': full_text[:30000] if full_text else ''
        }
        
        # 保存到文件
        with open(f'staff_{staff_id}.json', 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"  ✓ 已保存 ({len(full_text) if full_text else 0} 字符)")
        return result
        
    except Exception as e:
        print(f"  ✗ 失败: {e}")
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
    
    print("=" * 60)
    print("百度OKR周报抓取 - 连接现有Chrome")
    print(f"日期: {date}")
    print("=" * 60)
    input("\n按 Enter 开始...")
    
    # 启动Chrome调试模式
    print("\n检查Chrome调试端口...")
    chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    if not os.path.exists(chrome_path):
        chrome_path = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    
    # 检查端口9222
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    if sock.connect_ex(('localhost', 9222)) != 0:
        print("启动Chrome调试模式...")
        subprocess.Popen(
            [chrome_path, '--remote-debugging-port=9222'],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
        time.sleep(5)
    else:
        print("Chrome调试端口已存在")
    sock.close()
    
    results = []
    
    async with async_playwright() as p:
        print("\n连接Chrome...")
        browser = await p.chromium.connect_over_cdp('http://localhost:9222')
        print("  ✓ 已连接")
        
        # 获取或创建页面
        if browser.contexts and browser.contexts[0].pages:
            page = browser.contexts[0].pages[0]
        else:
            page = await browser.contexts[0].new_page()
        
        print(f"\n抓取 {len(STAFF_IDS)} 个下属...")
        for i, staff_id in enumerate(STAFF_IDS, 1):
            result = await fetch_single_staff(page, staff_id, date, i, len(STAFF_IDS))
            results.append(result)
            if i < len(STAFF_IDS):
                await asyncio.sleep(3)
        
        # 本人
        print("\n抓取本人...")
        results.append(await fetch_single_staff(page, MY_ID, date, len(STAFF_IDS) + 1, len(STAFF_IDS) + 1))
        
        # 汇总
        with open('all_results.json', 'w', encoding='utf-8') as f:
            json.dump({
                'date': date,
                'fetch_time': datetime.now().isoformat(),
                'success_count': sum(1 for r in results if r['status'] == 'success'),
                'results': results
            }, f, ensure_ascii=False, indent=2)
        
        print(f"\n完成! 成功: {sum(1 for r in results if r['status'] == 'success')}")
        print(f"文件保存在: {os.getcwd()}")
    
    input("\n按 Enter 退出")

if __name__ == '__main__':
    asyncio.run(main())
