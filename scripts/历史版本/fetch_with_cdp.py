#!/usr/bin/env python3
"""
百度OKR周报抓取脚本 - 使用系统Chrome
通过CDP连接到已运行的Chrome实例
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

async def fetch_weekly_report(page, staff_id, date, log_file):
    """抓取单个员工的周报"""
    url = f"{BASE_URL}&id={staff_id}&date={date}"
    log_msg = f"\n[{datetime.now().strftime('%H:%M:%S')}] 抓取: {staff_id}"
    print(log_msg)
    log_file.write(log_msg + '\n')
    
    try:
        await page.goto(url, wait_until='networkidle', timeout=30000)
        await asyncio.sleep(3)  # 等待页面渲染
        
        # 获取员工姓名
        name = staff_id
        try:
            # 尝试多种选择器
            selectors = [
                '[class*="name"]',
                '.user-name', 
                '.staff-name',
                '.avatar-name',
                '[class*="user"]'
            ]
            for sel in selectors:
                elem = await page.query_selector(sel)
                if elem:
                    text = await elem.inner_text()
                    if text and len(text.strip()) > 0 and len(text.strip()) < 20:
                        name = text.strip()
                        break
        except:
            pass
        
        log_msg = f"  员工姓名: {name}"
        print(log_msg)
        log_file.write(log_msg + '\n')
        
        # 获取页面完整文本
        full_text = await page.evaluate('() => document.body.innerText')
        
        # 获取页面HTML结构（用于后续解析）
        page_html = await page.content()
        
        # 保存原始内容
        result = {
            'id': staff_id,
            'name': name,
            'url': url,
            'fetch_time': datetime.now().isoformat(),
            'status': 'success',
            'content': {
                'text': full_text[:20000] if full_text else '',  # 限制长度
                'html': page_html[:50000] if page_html else ''   # 保存部分HTML
            }
        }
        
        # 保存到文件
        filename = f'../output/staff_{staff_id}.json'
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        log_msg = f"  ✓ 已保存: {filename} (内容长度: {len(full_text) if full_text else 0})"
        print(log_msg)
        log_file.write(log_msg + '\n')
        
        return result
        
    except Exception as e:
        error_msg = f"  ✗ 抓取失败: {str(e)}"
        print(error_msg)
        log_file.write(error_msg + '\n')
        
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
    print(f"{'='*60}")
    print(f"百度OKR周报抓取 - Playwright方案")
    print(f"日期: {date}")
    print(f"{'='*60}\n")
    
    # 打开日志文件
    with open('../output/playwright_log.txt', 'w', encoding='utf-8') as log_file:
        log_file.write(f"抓取开始: {datetime.now().isoformat()}\n")
        log_file.write(f"日期: {date}\n\n")
        
        results = []
        
        async with async_playwright() as p:
            # 尝试连接到已打开的Chrome（OpenClaw Browser的CDP端口）
            print("尝试连接到已运行的Chrome...")
            print("  CDP地址: http://127.0.0.1:18792")
            log_file.write("尝试连接到已运行的Chrome (CDP: 18792)\n")
            
            try:
                # 先获取WebSocket URL
                import urllib.request
                import json as json_lib
                
                req = urllib.request.Request('http://127.0.0.1:18792/json/version')
                with urllib.request.urlopen(req) as response:
                    version_info = json_lib.loads(response.read().decode())
                    ws_url = version_info.get('webSocketDebuggerUrl', 'ws://127.0.0.1:18792/devtools/browser')
                
                print(f"  WebSocket URL: {ws_url[:50]}...")
                log_file.write(f"WebSocket: {ws_url}\n")
                
                browser = await p.chromium.connect_over_cdp(ws_url)
                print("  ✓ 已连接到Chrome")
                log_file.write("✓ 连接成功\n")
            except Exception as e:
                print(f"  ✗ 连接失败: {e}")
                log_file.write(f"✗ 连接失败: {e}\n")
                return
            
            # 使用第一个context和page
            context = browser.contexts[0] if browser.contexts else await browser.new_context()
            page = context.pages[0] if context.pages else await context.new_page()
            
            log_file.write(f"\n开始抓取 {len(STAFF_IDS)} 个下属...\n")
            
            # 逐个抓取下属
            for i, staff_id in enumerate(STAFF_IDS, 1):
                log_file.write(f"\n[{i}/{len(STAFF_IDS)}] ")
                result = await fetch_weekly_report(page, staff_id, date, log_file)
                results.append(result)
                
                # 每个员工间隔2秒
                if i < len(STAFF_IDS):
                    await asyncio.sleep(2)
            
            # 抓取本人
            log_file.write(f"\n\n抓取本人周报...\n")
            my_result = await fetch_weekly_report(page, MY_ID, date, log_file)
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
            
            with open('../output/all_results.json', 'w', encoding='utf-8') as f:
                json.dump(summary, f, ensure_ascii=False, indent=2)
            
            log_msg = f"\n\n抓取完成!\n成功: {summary['success_count']}, 失败: {summary['error_count']}\n"
            print(log_msg)
            log_file.write(log_msg)
            
            await browser.close()

if __name__ == '__main__':
    asyncio.run(main())
