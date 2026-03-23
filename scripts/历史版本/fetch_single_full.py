#!/usr/bin/env python3
"""
单人测试版 - 修复滚动和完整页面抓取
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

async def scroll_to_bottom(page):
    """滚动到页面底部，确保所有内容加载"""
    print("  开始滚动页面...")
    
    # 获取页面高度
    last_height = await page.evaluate('() => document.body.scrollHeight')
    print(f"  初始页面高度: {last_height}")
    
    # 分步滚动到底部
    steps = 10
    for i in range(steps):
        position = (i + 1) / steps
        await page.evaluate(f'() => {{ window.scrollTo(0, document.body.scrollHeight * {position}); }}')
        await asyncio.sleep(1)
        print(f"  滚动 {i+1}/{steps} ({int(position*100)}%)...", end="\r")
    
    print("  滚动完成          ")
    
    # 等待可能的懒加载
    await asyncio.sleep(3)
    
    # 检查页面高度是否变化（动态加载）
    new_height = await page.evaluate('() => document.body.scrollHeight')
    print(f"  滚动后页面高度: {new_height}")
    
    if new_height > last_height:
        print(f"  检测到动态加载，新内容增加了 {new_height - last_height} 像素")
        # 再次滚动
        await page.evaluate('() => { window.scrollTo(0, document.body.scrollHeight); }')
        await asyncio.sleep(3)
    
    # 回到顶部
    await page.evaluate('() => { window.scrollTo(0, 0); }')
    await asyncio.sleep(1)
    
    # 再次缓慢滚动一次，确保所有内容都触发加载
    print("  第二次完整滚动...")
    for i in range(20):
        await page.evaluate(f'() => {{ window.scrollTo(0, {i * 200}); }}')
        await asyncio.sleep(0.5)
    
    # 最后回到顶部
    await page.evaluate('() => { window.scrollTo(0, 0); }')
    await asyncio.sleep(2)
    print("  滚动操作完成")

async def fetch_staff_complete(page, staff_id, date, desc):
    """完整抓取单个员工的周报"""
    url = f"{BASE_URL}&id={staff_id}&date={date}"
    print(f"\n{'='*70}")
    print(f"抓取: {desc}")
    print(f"日期: {date}")
    print(f"{'='*70}")
    
    try:
        # 1. 访问页面
        print("\n[1/6] 访问页面...")
        await page.goto(url, wait_until='domcontentloaded', timeout=60000)
        await asyncio.sleep(5)
        
        current_url = page.url
        if 'uuap' in current_url:
            print("⚠️ 需要登录")
            return {'status': 'login_required'}
        
        print(f"  ✓ 页面加载完成")
        
        # 2. 等待初始渲染
        print("\n[2/6] 等待初始渲染(5秒)...")
        await asyncio.sleep(5)
        
        # 3. 完整滚动页面
        print("\n[3/6] 滚动页面加载所有内容...")
        await scroll_to_bottom(page)
        
        # 4. 再次等待动态内容
        print("\n[4/6] 等待动态内容稳定(5秒)...")
        await asyncio.sleep(5)
        
        # 5. 最终滚动确保OKR部分加载
        print("\n[5/6] 最终确认滚动...")
        # 慢慢滚动到OKR部分通常所在的位置（页面中部）
        await page.evaluate('() => { window.scrollTo(0, 800); }')
        await asyncio.sleep(2)
        await page.evaluate('() => { window.scrollTo(0, 1500); }')
        await asyncio.sleep(2)
        await page.evaluate('() => { window.scrollTo(0, 0); }')
        await asyncio.sleep(2)
        
        # 6. 提取完整内容
        print("\n[6/6] 提取完整页面内容...")
        
        # 获取完整文本
        full_text = await page.evaluate('() => document.body.innerText')
        print(f"  ✓ 获取文本: {len(full_text)} 字符")
        
        # 获取完整HTML
        full_html = await page.content()
        print(f"  ✓ 获取HTML: {len(full_html)} 字符")
        
        # 尝试通过选择器获取主要内容区域
        main_content = ""
        selectors = [
            'main',
            '[class*="content"]',
            '[class*="main"]',
            '#app',
            'body'
        ]
        
        for selector in selectors:
            try:
                elem = await page.query_selector(selector)
                if elem:
                    text = await elem.inner_text()
                    if len(text) > len(main_content):
                        main_content = text
                        print(f"  ✓ 从 '{selector}' 获取 {len(text)} 字符")
            except:
                pass
        
        # 7. 截取完整页面截图
        print("\n  截取完整页面截图...")
        screenshot_file = f'{staff_id}_{date}_fullpage.png'
        
        # 先滚动到顶部确保从头开始截图
        await page.evaluate('() => { window.scrollTo(0, 0); }')
        await asyncio.sleep(1)
        
        # 截取全页面
        await page.screenshot(path=screenshot_file, full_page=True)
        print(f"  ✓ 截图保存: {screenshot_file}")
        
        # 8. 保存数据
        result = {
            'id': staff_id,
            'date': date,
            'desc': desc,
            'fetch_time': datetime.now().isoformat(),
            'status': 'success',
            'content_lengths': {
                'innerText': len(full_text),
                'html': len(full_html),
                'main_content': len(main_content)
            },
            'content': {
                'innerText': full_text[:50000],
                'main_content': main_content[:30000]
            },
            'screenshot': screenshot_file
        }
        
        json_file = f'{staff_id}_{date}_complete.json'
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"\n{'='*70}")
        print(f"✓ 抓取完成!")
        print(f"  JSON: {json_file}")
        print(f"  截图: {screenshot_file}")
        print(f"{'='*70}")
        
        # 显示内容统计
        print(f"\n内容统计:")
        print(f"  文本总长度: {len(full_text)} 字符")
        print(f"  HTML总长度: {len(full_html)} 字符")
        print(f"  主要内容: {len(main_content)} 字符")
        
        # 检测关键词
        has_okr = any(k in full_text.lower() for k in ['okr', 'objective', 'key result', 'o1', 'kr1'])
        print(f"\n内容检测:")
        print(f"  包含OKR内容: {'✓' if has_okr else '✗'}")
        
        # 显示预览
        print(f"\n内容预览 (前800字符):")
        print("-" * 70)
        preview = full_text[:800] if full_text else "(无内容)"
        print(preview)
        print("-" * 70)
        
        return result
        
    except Exception as e:
        print(f"\n✗ 抓取失败: {e}")
        import traceback
        traceback.print_exc()
        return {'status': 'error', 'error': str(e)}

async def main():
    last = get_week_date(0)
    before = get_week_date(1)
    
    print("="*70)
    print(" "*20 + "百度OKR周报抓取")
    print(" "*18 + "【单人完整版】")
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
        r1 = await fetch_staff_complete(page, MY_ID, last, "上周周报")
        
        if r1.get('status') == 'success':
            await asyncio.sleep(3)
            r2 = await fetch_staff_complete(page, MY_ID, before, "上上周周报")
        
        print("\n" + "="*70)
        print("全部完成!")
        print("="*70)
        print(f"\n生成的文件:")
        print(f"  1. {MY_ID}_{last}_fullpage.png (上周截图)")
        print(f"  2. {MY_ID}_{last}_complete.json (上周数据)")
        print(f"  3. {MY_ID}_{before}_fullpage.png (上上周截图)")
        print(f"  4. {MY_ID}_{before}_complete.json (上上周数据)")
        print(f"\n请打开PNG文件检查是否包含完整的OKR部分")
        
        await browser.close()
        input("\n按Enter退出...")

if __name__ == '__main__':
    asyncio.run(main())
