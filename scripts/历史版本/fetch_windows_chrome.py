#!/usr/bin/env python3
"""
百度OKR周报抓取脚本 - 使用Windows宿主机Chrome
通过WSL调用Windows Chrome
"""

import json
import asyncio
import subprocess
from datetime import datetime, timedelta
from playwright.async_api import async_playwright

# 员工ID
MY_ID = 's673090'
STAFF_IDS = [
    's801573', 's733504', 's682083', 's929270', 's798268',
    's811385', 's865585', 's783729', 's400192', 's823700'
]

BASE_URL = 'https://okr.baidu-int.com/pages/weekly.html#/index?mode=steam'

# Windows Chrome路径
WIN_CHROME = '/mnt/c/Program Files/Google/Chrome/Application/chrome.exe'

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
            selectors = ['[class*="name"]', '.user-name', '.staff-name', '.avatar-name']
            for sel in selectors:
                elem = await page.query_selector(sel)
                if elem:
                    text = await elem.inner_text()
                    if text and 0 < len(text.strip()) < 20:
                        name = text.strip()
                        break
        except:
            pass
        
        log_msg = f"  员工姓名: {name}"
        print(log_msg)
        log_file.write(log_msg + '\n')
        
        # 获取页面内容
        full_text = await page.evaluate('() => document.body.innerText')
        page_html = await page.content()
        
        # 截图保存（用于验证）
        screenshot_path = f'../output/screenshot_{staff_id}.png'
        await page.screenshot(path=screenshot_path, full_page=True)
        log_msg = f"  截图已保存: {screenshot_path}"
        print(log_msg)
        log_file.write(log_msg + '\n')
        
        # 保存结果
        result = {
            'id': staff_id,
            'name': name,
            'url': url,
            'fetch_time': datetime.now().isoformat(),
            'status': 'success',
            'content': {
                'text': full_text[:30000] if full_text else '',
                'html_length': len(page_html) if page_html else 0
            }
        }
        
        filename = f'../output/staff_{staff_id}.json'
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        log_msg = f"  ✓ 已保存: {filename}"
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
    print(f"百度OKR周报抓取 - Windows Chrome方案")
    print(f"日期: {date}")
    print(f"Chrome路径: {WIN_CHROME}")
    print(f"{'='*60}\n")
    
    # 检查Chrome是否存在
    import os
    if not os.path.exists(WIN_CHROME):
        print(f"错误: 未找到Windows Chrome: {WIN_CHROME}")
        return
    
    with open('../output/windows_chrome_log.txt', 'w', encoding='utf-8') as log_file:
        log_file.write(f"抓取开始: {datetime.now().isoformat()}\n")
        log_file.write(f"日期: {date}\n")
        log_file.write(f"Chrome: {WIN_CHROME}\n\n")
        
        results = []
        
        async with async_playwright() as p:
            print("启动Windows Chrome...")
            log_file.write("启动Windows Chrome...\n")
            
            try:
                # 使用Windows Chrome
                browser = await p.chromium.launch(
                    executable_path=WIN_CHROME,
                    headless=False,  # 可见模式，便于调试
                    args=['--disable-gpu', '--no-sandbox']
                )
                print("  ✓ Chrome已启动")
                log_file.write("✓ Chrome已启动\n")
            except Exception as e:
                print(f"  ✗ 启动失败: {e}")
                log_file.write(f"✗ 启动失败: {e}\n")
                return
            
            # 创建新页面
            context = await browser.new_context(
                viewport={'width': 1920, 'height': 1080}
            )
            page = await context.new_page()
            
            log_file.write(f"\n开始抓取 {len(STAFF_IDS)} 个下属...\n")
            
            # 逐个抓取下属（只抓取第一个测试）
            test_mode = True  # 测试模式，只抓第一个
            
            if test_mode:
                print("\n*** 测试模式: 只抓取第一个员工 ***\n")
                log_file.write("\n*** 测试模式: 只抓取第一个员工 ***\n")
                staff_list = [STAFF_IDS[0]]  # 只抓第一个
            else:
                staff_list = STAFF_IDS
            
            for i, staff_id in enumerate(staff_list, 1):
                log_file.write(f"\n[{i}/{len(staff_list)}] ")
                result = await fetch_weekly_report(page, staff_id, date, log_file)
                results.append(result)
                
                if i < len(staff_list):
                    await asyncio.sleep(2)
            
            # 如果测试成功，询问是否继续抓取全部
            if test_mode and results and results[0]['status'] == 'success':
                print("\n" + "="*60)
                print("✓ 测试抓取成功!")
                print("="*60)
                log_file.write("\n\n测试抓取成功!\n")
                
                # 保存测试成功标记
                with open('../output/test_success.flag', 'w') as f:
                    f.write('1')
            
            # 抓取本人（仅非测试模式）
            if not test_mode:
                log_file.write(f"\n\n抓取本人周报...\n")
                my_result = await fetch_weekly_report(page, MY_ID, date, log_file)
                results.append(my_result)
            
            # 保存汇总
            summary = {
                'date': date,
                'fetch_time': datetime.now().isoformat(),
                'my_id': MY_ID,
                'staff_count': len(staff_list),
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
