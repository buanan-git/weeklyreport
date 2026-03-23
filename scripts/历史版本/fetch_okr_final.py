#!/usr/bin/env python3
"""
百度OKR周报抓取 - 网络修复版
处理VPN连接和DNS问题
"""

import json
import asyncio
import socket
from datetime import datetime, timedelta
from playwright.async_api import async_playwright

MY_ID = 's673090'
STAFF_IDS = [
    's801573', 's733504', 's682083', 's929270', 's798268',
    's811385', 's865585', 's783729', 's400192', 's823700'
]
BASE_URL = 'https://okr.baidu-int.com/pages/weekly.html#/index?mode=steam'
OKR_HOME = 'https://okr.baidu-int.com'

def get_week_date(weeks_ago=0):
    today = datetime(2026, 2, 11)
    target = today - timedelta(days=today.weekday() + 7 + (weeks_ago * 7))
    return f"s{target.strftime('%Y-%m-%d')}"

def check_network():
    """检查网络连接"""
    print("检查网络...")
    try:
        # 尝试解析域名
        socket.gethostbyname('okr.baidu-int.com')
        print("✓ DNS解析正常")
        return True
    except:
        print("⚠️  无法解析 okr.baidu-int.com")
        print("   请确保:")
        print("   1. Windows已连接公司VPN")
        print("   2. 或手动在浏览器中打开OKR")
        return False

async def find_scroll_container(page):
    """查找可滚动容器"""
    print("  查找滚动容器...")
    
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
        print(f"  找到 {len(containers)} 个容器")
        for i, c in enumerate(containers, 1):
            cls = c.get('class', '')[:40]
            print(f"    {i}. {c['tag']}.{cls}: {c['scrollHeight']}px")
        
        first = containers[0]
        if first.get('class'):
            selector = f"[class*='{first['class'].split()[0]}']"
        elif first.get('id'):
            selector = f"#{first['id']}"
        else:
            selector = first['tag'].lower()
        
        print(f"  ✓ 使用: {selector[:50]}")
        return selector
    
    print("  ⚠️ 使用body")
    return 'body'

async def scroll_container(page, selector):
    """滚动容器"""
    print(f"  滚动容器...")
    
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
    
    print(f"  在({int(center['x'])}, {int(center['y'])})滚轮滚动...")
    await page.mouse.move(center['x'], center['y'])
    await asyncio.sleep(0.5)
    
    for i in range(30):
        await page.mouse.wheel(0, 400)
        await asyncio.sleep(0.3)
        if i % 10 == 0:
            print(f"    已滚动 {(i+1)*400}px...", end="\r")
    
    print("  滚动完成          ")
    
    # 键盘确认
    await page.click(selector)
    await asyncio.sleep(0.3)
    for _ in range(5):
        await page.keyboard.press('PageDown')
        await asyncio.sleep(0.3)
    
    print("  ✓ 完成")
    await asyncio.sleep(3)

async def fetch_staff(page, staff_id, date, desc, index=0, total=0):
    """抓取员工周报"""
    url = f"{BASE_URL}&id={staff_id}&date={date}"
    progress = f"[{index}/{total}] " if index else ""
    
    print(f"\n{'='*70}")
    print(f"{progress}抓取: {desc}")
    print(f"日期: {date}")
    print(f"{'='*70}")
    
    try:
        print("\n[1/5] 打开页面...")
        await page.goto(url, wait_until='domcontentloaded', timeout=60000)
        await asyncio.sleep(5)
        
        if 'uuap' in page.url:
            print("⚠️ 需要重新登录")
            return {'id': staff_id, 'date': date, 'status': 'login_required'}
        
        print("✓ 页面已打开")
        
        print("\n[2/5] 等待渲染(5秒)...")
        await asyncio.sleep(5)
        
        print("\n[3/5] 查找滚动容器...")
        container = await find_scroll_container(page)
        
        print("\n[4/5] 滚动加载内容...")
        await scroll_container(page, container)
        
        print("\n[5/5] 等待动态内容(5秒)...")
        await asyncio.sleep(5)
        
        print("\n  提取页面内容...")
        full_text = await page.evaluate('() => document.body.innerText')
        print(f"  ✓ 获取 {len(full_text)} 字符")
        
        print("  截图...")
        screenshot = f'{staff_id}_{date}.png'
        await page.screenshot(path=screenshot, full_page=True)
        print(f"  ✓ {screenshot}")
        
        result = {
            'id': staff_id,
            'date': date,
            'desc': desc,
            'container': container,
            'fetch_time': datetime.now().isoformat(),
            'status': 'success',
            'content': full_text[:50000],
            'screenshot': screenshot,
            'has_okr': any(k in full_text.lower() for k in ['okr', 'objective', 'o1', 'kr1'])
        }
        
        json_file = f'{staff_id}_{date}.json'
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"\n{'='*70}")
        print(f"✓ 抓取成功!")
        print(f"  JSON: {json_file}")
        print(f"  截图: {screenshot}")
        print(f"  包含OKR: {'是' if result['has_okr'] else '否'}")
        print(f"{'='*70}")
        
        return result
        
    except Exception as e:
        print(f"\n✗ 抓取失败: {e}")
        return {'id': staff_id, 'date': date, 'status': 'error', 'error': str(e)}

async def main():
    last_week = get_week_date(0)
    week_before = get_week_date(1)
    
    print("="*70)
    print(" "*20 + "百度OKR周报抓取")
    print(" "*18 + "【网络修复版】")
    print("="*70)
    print(f"\n员工ID: {MY_ID}")
    print(f"下属数量: {len(STAFF_IDS)}")
    print(f"上周: {last_week}")
    print(f"上上周: {week_before}")
    
    # 检查网络
    network_ok = check_network()
    
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
            print(f"⚠️  访问OKR失败: {e}")
            print("\n" + "="*70)
            print("切换到手动模式:")
            print("  1. 请在Chrome中手动打开 https://okr.baidu-int.com")
            print("  2. 确保VPN已连接")
            print("  3. 完成登录")
            print("="*70)
            input("\n完成后按 Enter...")
        
        print("\n" + "="*70)
        print("⚠️  请确保已登录OKR系统")
        print("="*70)
        input("\n登录完成后按 Enter...")
        
        # 验证当前页面
        current_url = page.url
        print(f"\n当前页面: {current_url[:60]}...")
        
        if 'uuap' in current_url or 'login' in current_url:
            print("\n✗ 似乎未登录，请重试")
            await browser.close()
            return
        
        if 'okr' not in current_url:
            print("\n⚠️  未在OKR页面，请手动导航到OKR后重试")
            await browser.close()
            return
        
        print("✓ 登录验证通过")
        
        all_results = []
        
        # 抓取所有下属上周
        print(f"\n{'='*70}")
        print(f"开始抓取 {len(STAFF_IDS)} 个下属上周周报")
        print(f"{'='*70}")
        
        for i, staff_id in enumerate(STAFF_IDS, 1):
            result = await fetch_staff(page, staff_id, last_week, f"下属{i}", i, len(STAFF_IDS))
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
            
            r1 = await fetch_staff(page, MY_ID, last_week, "本人-上周", 1, 2)
            all_results.append(r1)
            
            if r1['status'] != 'login_required':
                await asyncio.sleep(3)
                r2 = await fetch_staff(page, MY_ID, week_before, "本人-上上周", 2, 2)
                all_results.append(r2)
        
        # 汇总
        summary = {
            'fetch_time': datetime.now().isoformat(),
            'my_id': MY_ID,
            'staff_count': len(STAFF_IDS),
            'success': sum(1 for r in all_results if r['status'] == 'success'),
            'failed': sum(1 for r in all_results if r['status'] == 'error'),
            'has_okr': sum(1 for r in all_results if r.get('has_okr')),
            'results': all_results
        }
        
        with open('final_results.json', 'w', encoding='utf-8') as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)
        
        print("\n" + "="*70)
        print("全部完成!")
        print("="*70)
        print(f"\n统计:")
        print(f"  总数: {len(all_results)}")
        print(f"  成功: {summary['success']}")
        print(f"  失败: {summary['failed']}")
        print(f"  包含OKR: {summary['has_okr']}")
        print(f"\n文件: staff_*.json, staff_*.png, final_results.json")
        
        await browser.close()
        input("\n按Enter退出...")

if __name__ == '__main__':
    asyncio.run(main())
