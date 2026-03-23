#!/usr/bin/env python3
"""
单人测试版 - 使用Playwright原生滚动方法
针对SPA应用（React/Vue）的滚动处理
"""

import json
import asyncio
from datetime import datetime, timedelta
from playwright.async_api import async_playwright

MY_ID = 's673090'
BASE_URL = 'https://okr.baidu-int.com/pages/weekly.html#/index?mode=steam'

def get_week_date(weeks_ago=0):
    today = datetime(2026, 2, 11)
    target = today - timedelta(days=today.weekday() + 7 + (weeks_ago * 7))
    return f"s{target.strftime('%Y-%m-%d')}"

async def find_scrollable_element(page):
    """查找页面上可滚动的元素"""
    print("  查找滚动容器...")
    
    # 常见的滚动容器选择器
    scroll_selectors = [
        'html',
        'body',
        '[class*="scroll"]',
        '[class*="content"]',
        'main',
        '[role="main"]',
        '.app',
        '#app',
        '[class*="layout"]'
    ]
    
    for selector in scroll_selectors:
        try:
            elem = await page.query_selector(selector)
            if elem:
                # 检查元素是否可以滚动
                is_scrollable = await elem.evaluate('''el => {
                    return el.scrollHeight > el.clientHeight || 
                           window.getComputedStyle(el).overflow === 'auto' ||
                           window.getComputedStyle(el).overflow === 'scroll' ||
                           window.getComputedStyle(el).overflowY === 'auto' ||
                           window.getComputedStyle(el).overflowY === 'scroll';
                }''')
                if is_scrollable:
                    print(f"  ✓ 找到滚动容器: {selector}")
                    return selector
        except:
            continue
    
    print("  ⚠️ 未找到特定滚动容器，使用body")
    return 'body'

async def scroll_page_native(page, scroll_selector):
    """使用Playwright原生方法滚动"""
    print("  开始滚动页面...")
    
    try:
        # 获取元素
        element = await page.query_selector(scroll_selector)
        if not element:
            element = page
        
        # 方法1: 使用鼠标滚轮
        print("  方法1: 鼠标滚轮滚动...")
        for i in range(20):
            await page.mouse.wheel(0, 500)  # 每次滚动500像素
            await asyncio.sleep(0.5)
            if i % 5 == 0:
                print(f"    滚动 {i+1}/20...", end="\r")
        print("  鼠标滚轮滚动完成      ")
        
        await asyncio.sleep(2)
        
        # 方法2: 使用键盘
        print("  方法2: 键盘滚动...")
        await page.keyboard.press('End')  # 跳到页面底部
        await asyncio.sleep(2)
        await page.keyboard.press('Home')  # 回到顶部
        await asyncio.sleep(2)
        
        # 方法3: 再次使用鼠标滚动确保
        print("  方法3: 二次滚动确认...")
        for i in range(10):
            await page.mouse.wheel(0, 300)
            await asyncio.sleep(0.3)
        
        print("  ✓ 滚动操作完成")
        
    except Exception as e:
        print(f"  ⚠️ 滚动出错: {e}")
        # 备用方案：使用JavaScript
        print("  使用备用方案...")
        await page.evaluate('''() => {
            const scrollable = document.querySelector('.scroll-container, [class*="scroll"], main, body');
            if (scrollable) {
                for (let i = 0; i <= 10; i++) {
                    setTimeout(() => {
                        scrollable.scrollTop = scrollable.scrollHeight * (i / 10);
                    }, i * 200);
                }
            }
        }''')
        await asyncio.sleep(3)

async def fetch_staff_native(page, staff_id, date, desc):
    """抓取员工周报"""
    url = f"{BASE_URL}&id={staff_id}&date={date}"
    print(f"\n{'='*70}")
    print(f"抓取: {desc} ({date})")
    print(f"{'='*70}")
    
    try:
        # 访问页面
        print("\n[1/5] 访问页面...")
        await page.goto(url, wait_until='domcontentloaded', timeout=60000)
        await asyncio.sleep(5)
        
        if 'uuap' in page.url:
            print("⚠️ 需要登录")
            return {'status': 'login_required'}
        
        print(f"  ✓ 页面加载完成")
        
        # 查找滚动容器
        print("\n[2/5] 查找滚动容器...")
        scroll_selector = await find_scrollable_element(page)
        
        # 等待初始内容
        print("\n[3/5] 等待初始渲染(5秒)...")
        await asyncio.sleep(5)
        
        # 滚动页面
        print("\n[4/5] 滚动页面加载内容...")
        await scroll_page_native(page, scroll_selector)
        
        # 等待动态内容
        print("\n  等待动态内容(5秒)...")
        await asyncio.sleep(5)
        
        # 再次滚动确保
        print("  二次滚动...")
        await page.keyboard.press('End')
        await asyncio.sleep(2)
        await page.keyboard.press('Home')
        await asyncio.sleep(2)
        
        # 提取内容
        print("\n[5/5] 提取页面内容...")
        full_text = await page.evaluate('() => document.body.innerText')
        print(f"  ✓ 获取 {len(full_text)} 字符")
        
        # 截图
        print("  截取完整页面...")
        screenshot = f'{staff_id}_{date}_native.png'
        await page.screenshot(path=screenshot, full_page=True)
        print(f"  ✓ 截图: {screenshot}")
        
        # 保存
        result = {
            'id': staff_id, 'date': date, 'desc': desc,
            'content': full_text[:50000],
            'screenshot': screenshot,
            'has_okr': any(k in full_text.lower() for k in ['okr', 'objective', 'o1', 'kr1'])
        }
        
        with open(f'{staff_id}_{date}_native.json', 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False)
        
        print(f"\n{'='*70}")
        print(f"✓ 完成!")
        print(f"  JSON: {staff_id}_{date}_native.json")
        print(f"  截图: {screenshot}")
        print(f"{'='*70}")
        
        print(f"\n内容预览 (前600字符):")
        print("-" * 70)
        print(full_text[:600] if full_text else "(无内容)")
        print("-" * 70)
        
        return result
        
    except Exception as e:
        print(f"\n✗ 失败: {e}")
        import traceback
        traceback.print_exc()
        return {'status': 'error', 'error': str(e)}

async def main():
    last = get_week_date(0)
    before = get_week_date(1)
    
    print("="*70)
    print(" "*18 + "百度OKR周报抓取")
    print(" "*15 + "【原生滚动版】")
    print("="*70)
    print(f"\n员工ID: {MY_ID}")
    print(f"上周: {last}")
    print(f"上上周: {before}")
    
    async with async_playwright() as p:
        print("\n【1】启动Chrome...")
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page(viewport={'width': 1920, 'height': 1080})
        print("✓ Chrome已启动")
        
        print("\n【2】访问OKR...")
        await page.goto("https://okr.baidu-int.com", wait_until='domcontentloaded')
        print("✓ 页面已加载")
        
        print("\n" + "="*70)
        print("⚠️ 请在Chrome中登录OKR系统")
        print("="*70)
        input("\n登录完成后按 Enter...")
        
        if 'uuap' in page.url:
            print("✗ 未登录成功")
            return
        
        print("✓ 登录验证通过")
        
        # 抓取两周
        r1 = await fetch_staff_native(page, MY_ID, last, "上周周报")
        if r1.get('status') == 'success':
            await asyncio.sleep(3)
            r2 = await fetch_staff_native(page, MY_ID, before, "上上周周报")
        
        print("\n" + "="*70)
        print("全部完成!")
        print("="*70)
        print(f"\n文件:")
        print(f"  1. {MY_ID}_{last}_native.png")
        print(f"  2. {MY_ID}_{last}_native.json")
        print(f"  3. {MY_ID}_{before}_native.png")
        print(f"  4. {MY_ID}_{before}_native.json")
        print(f"\n请检查截图是否包含OKR部分")
        
        await browser.close()
        input("\n按Enter退出...")

if __name__ == '__main__':
    asyncio.run(main())
