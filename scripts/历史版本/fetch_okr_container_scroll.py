#!/usr/bin/env python3
"""
百度OKR周报抓取 - 局部滚动版
针对周报内容区域的独立滚动容器
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

async def find_scrollable_container(page):
    """查找周报内容的滚动容器"""
    print("  查找周报滚动容器...")
    
    # 常见的局部滚动容器选择器
    scrollable_selectors = [
        '[class*="weekly"] [class*="content"]',
        '[class*="weekly"] [class*="scroll"]',
        '[class*="report"] [class*="content"]',
        '[class*="detail"] [class*="content"]',
        '.weekly-report-content',
        '.report-content',
        '.content-area',
        '[class*="main"] [class*="scroll"]',
        '.scroll-container',

        'div[style*="overflow"]',
        'div[style*="scroll"]'
    ]
    
    for selector in scrollable_selectors:
        try:
            elem = await page.query_selector(selector)
            if elem:
                # 检查是否可滚动
                is_scrollable = await elem.evaluate('''el => {
                    const style = window.getComputedStyle(el);
                    const canScroll = (el.scrollHeight > el.clientHeight) ||
                                      style.overflow === 'auto' ||
                                      style.overflow === 'scroll' ||
                                      style.overflowY === 'auto' ||
                                      style.overflowY === 'scroll';
                    return canScroll;
                }''')
                
                if is_scrollable:
                    # 获取容器信息
                    info = await elem.evaluate('''el => ({
                        scrollHeight: el.scrollHeight,
                        clientHeight: el.clientHeight,
                        scrollTop: el.scrollTop
                    })''')
                    print(f"  ✓ 找到滚动容器: {selector}")
                    print(f"    滚动高度: {info['scrollHeight']}, 可视高度: {info['clientHeight']}")
                    return selector
        except:
            continue
    
    # 如果没找到特定容器，尝试获取所有可能滚动的div
    print("  尝试查找所有可滚动元素...")
    try:
        scrollables = await page.evaluate('''() => {
            const allDivs = document.querySelectorAll('div');
            const scrollableDivs = [];
            for (let div of allDivs) {
                if (div.scrollHeight > div.clientHeight && div.scrollHeight > 500) {
                    scrollableDivs.push({
                        selector: div.className || div.id || 'unnamed',
                        scrollHeight: div.scrollHeight,
                        clientHeight: div.clientHeight
                    });
                }
            }
            return scrollableDivs.slice(0, 5);  // 返回前5个
        }''')
        
        if scrollables:
            print(f"  找到 {len(scrollables)} 个可滚动元素:")
            for s in scrollables:
                print(f"    - {s['selector']}: {s['scrollHeight']}px")
            # 使用第一个
            return f"[class*='{scrollables[0]['selector'].split()[0]}']" if scrollables[0]['selector'] else 'div'
    except Exception as e:
        print(f"  查找失败: {e}")
    
    print("  ⚠️ 未找到特定滚动容器，使用默认body")
    return 'body'

async def scroll_container(page, container_selector):
    """在特定容器内滚动"""
    print(f"  在容器内滚动: {container_selector}")
    
    try:
        element = await page.query_selector(container_selector)
        if not element:
            print("  ⚠️ 未找到容器，使用页面滚动")
            # 使用鼠标滚轮在页面中心滚动
            for i in range(30):
                await page.mouse.wheel(0, 300)
                await asyncio.sleep(0.5)
            return
        
        # 获取容器信息
        container_info = await element.evaluate('''el => ({
            scrollHeight: el.scrollHeight,
            clientHeight: el.clientHeight
        })''')
        
        total_height = container_info['scrollHeight']
        viewport_height = container_info['clientHeight']
        
        print(f"  容器总高度: {total_height}px, 可视高度: {viewport_height}px")
        
        if total_height <= viewport_height:
            print("  容器无需滚动")
            return
        
        # 在容器内分步滚动
        steps = min(20, int(total_height / 200))  # 每步约200px
        print(f"  分 {steps} 步滚动...")
        
        for i in range(steps):
            scroll_position = int((total_height / steps) * (i + 1))
            
            # 使用JavaScript滚动特定容器
            await element.evaluate(f'el => {{ el.scrollTop = {scroll_position}; }}')
            await asyncio.sleep(0.8)
            
            if i % 5 == 0:
                print(f"    滚动 {i+1}/{steps} ({int((i+1)/steps*100)}%)...", end="\r")
        
        print(f"  滚动完成 (100%)      ")
        
        # 回到顶部再慢慢滚动一次，确保懒加载
        print("  二次滚动确认...")
        await element.evaluate('el => { el.scrollTop = 0; }')
        await asyncio.sleep(1)
        
        for i in range(10):
            await element.evaluate(f'el => {{ el.scrollTop = {i * 300}; }}')
            await asyncio.sleep(0.5)
        
        # 最后滚动到底部
        await element.evaluate('el => { el.scrollTop = el.scrollHeight; }')
        await asyncio.sleep(2)
        
        print("  ✓ 容器滚动完成")
        
    except Exception as e:
        print(f"  ⚠️ 容器滚动出错: {e}")
        print("  使用备用方案...")
        # 备用：页面中心鼠标滚轮
        await page.mouse.move(960, 540)  # 屏幕中心
        for i in range(20):
            await page.mouse.wheel(0, 400)
            await asyncio.sleep(0.5)

async def fetch_staff_with_container_scroll(page, staff_id, date, desc):
    """抓取员工周报（局部滚动版）"""
    url = f"{BASE_URL}&id={staff_id}&date={date}"
    
    print(f"\n{'='*70}")
    print(f"抓取: {desc} ({date})")
    print(f"{'='*70}")
    
    try:
        # 1. 打开页面
        print("\n[1/6] 打开页面...")
        await page.goto(url, wait_until='domcontentloaded', timeout=60000)
        await asyncio.sleep(5)
        
        if 'uuap' in page.url:
            print("⚠️ 需要重新登录")
            return {'status': 'login_required'}
        
        print(f"  ✓ 页面已打开")
        
        # 2. 等待初始渲染
        print("\n[2/6] 等待初始渲染(5秒)...")
        await asyncio.sleep(5)
        
        # 3. 查找滚动容器
        print("\n[3/6] 查找滚动容器...")
        container_selector = await find_scrollable_container(page)
        
        # 4. 在容器内滚动
        print("\n[4/6] 在容器内滚动加载内容...")
        await scroll_container(page, container_selector)
        
        # 5. 等待动态内容
        print("\n[5/6] 等待动态内容稳定(5秒)...")
        await asyncio.sleep(5)
        
        # 6. 提取内容
        print("\n[6/6] 提取页面内容...")
        full_text = await page.evaluate('() => document.body.innerText')
        print(f"  ✓ 获取 {len(full_text)} 字符")
        
        # 尝试获取容器内的具体内容
        container_text = ""
        try:
            elem = await page.query_selector(container_selector)
            if elem:
                container_text = await elem.inner_text()
                print(f"  ✓ 容器内容: {len(container_text)} 字符")
        except:
            pass
        
        # 截图
        print("  截取完整页面...")
        screenshot = f'{staff_id}_{date}_container.png'
        await page.screenshot(path=screenshot, full_page=True)
        print(f"  ✓ 截图: {screenshot}")
        
        # 保存
        result = {
            'id': staff_id,
            'date': date,
            'desc': desc,
            'container_selector': container_selector,
            'fetch_time': datetime.now().isoformat(),
            'status': 'success',
            'content': {
                'full_text': full_text[:50000],
                'container_text': container_text[:30000] if container_text else ''
            },
            'screenshot': screenshot,
            'has_okr': any(k in full_text.lower() for k in ['okr', 'objective', 'o1', 'kr1'])
        }
        
        json_file = f'{staff_id}_{date}_container.json'
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"\n{'='*70}")
        print(f"✓ 抓取完成!")
        print(f"  JSON: {json_file}")
        print(f"  截图: {screenshot}")
        print(f"{'='*70}")
        
        print(f"\n内容预览 (前800字符):")
        print("-" * 70)
        print(full_text[:800])
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
    print(" "*16 + "【局部滚动版】")
    print("="*70)
    print(f"\n员工ID: {MY_ID}")
    print(f"上周: {last}")
    print(f"上上周: {before}")
    print("\n此版本会自动查找周报区域的滚动容器并滚动")
    
    async with async_playwright() as p:
        print("\n" + "="*70)
        print("[系统] 启动Chrome...")
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page(viewport={'width': 1920, 'height': 1080})
        print("✓ Chrome已启动")
        
        print("\n[系统] 访问OKR...")
        await page.goto("https://okr.baidu-int.com", wait_until='domcontentloaded')
        print("✓ 页面已加载")
        
        print("\n" + "="*70)
        print("⚠️  请先登录OKR系统")
        print("="*70)
        input("\n登录完成后按 Enter...")
        
        if 'uuap' in page.url:
            print("✗ 未登录成功")
            return
        
        print("✓ 登录验证通过")
        
        # 抓取上周
        r1 = await fetch_staff_with_container_scroll(page, MY_ID, last, "上周周报")
        
        if r1.get('status') == 'success':
            await asyncio.sleep(3)
            r2 = await fetch_staff_with_container_scroll(page, MY_ID, before, "上上周周报")
        
        print("\n" + "="*70)
        print("全部完成!")
        print("="*70)
        print(f"\n文件:")
        print(f"  1. {MY_ID}_{last}_container.png")
        print(f"  2. {MY_ID}_{last}_container.json")
        print(f"  3. {MY_ID}_{before}_container.png") 
        print(f"  4. {MY_ID}_{before}_container.json")
        print(f"\n请检查截图是否包含OKR部分")
        
        await browser.close()
        input("\n按Enter退出...")

if __name__ == '__main__':
    asyncio.run(main())
