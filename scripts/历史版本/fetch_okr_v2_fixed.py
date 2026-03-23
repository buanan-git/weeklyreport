#!/usr/bin/env python3
"""
百度OKR周报抓取 - 6种滚动方法（修复版）
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

async def scroll_all_methods(page, selector):
    """尝试所有滚动方法"""
    print("  尝试6种滚动方法...")
    
    # 1. scrollTop
    print("  1. scrollTop...")
    try:
        await page.evaluate(f'''
            () => {{
                const el = document.querySelector("{selector}");
                if (el) {{
                    const end = el.scrollHeight - el.clientHeight;
                    for (let i = 0; i <= 10; i++) {{
                        setTimeout(() => {{ el.scrollTop = end * (i/10); }}, i * 200);
                    }}
                }}
            }}
        ''')
        await asyncio.sleep(3)
        print("    ✓")
    except Exception as e:
        print(f"    ✗ {e}")
    
    # 2. scrollIntoView
    print("  2. scrollIntoView...")
    try:
        await page.evaluate(f'''
            () => {{
                const el = document.querySelector("{selector}");
                if (el && el.children.length > 3) {{
                    el.children[Math.floor(el.children.length/2)].scrollIntoView({{behavior: 'smooth'}});
                }}
            }}
        ''')
        await asyncio.sleep(2)
        print("    ✓")
    except Exception as e:
        print(f"    ✗ {e}")
    
    # 3. 动画
    print("  3. 平滑滚动...")
    try:
        await page.evaluate(f'''
            () => {{
                const el = document.querySelector("{selector}");
                if (el) {{
                    const start = el.scrollTop;
                    const end = el.scrollHeight - el.clientHeight;
                    const duration = 2000;
                    const startTime = performance.now();
                    
                    function step(currentTime) {{
                        const elapsed = currentTime - startTime;
                        const progress = Math.min(elapsed / duration, 1);
                        el.scrollTop = start + (end - start) * progress;
                        if (progress < 1) {{
                            requestAnimationFrame(step);
                        }}
                    }}
                    requestAnimationFrame(step);
                }}
            }}
        ''')
        await asyncio.sleep(3)
        print("    ✓")
    except Exception as e:
        print(f"    ✗ {e}")
    
    # 4. 拖拽滚动条
    print("  4. 拖拽滚动条...")
    try:
        box = await page.evaluate(f'''
            () => {{
                const el = document.querySelector("{selector}");
                if (el) {{
                    const rect = el.getBoundingClientRect();
                    return {{x: rect.right - 10, y: rect.top + 100}};
                }}
                return {{x: 960, y: 500}};
            }}
        ''')
        
        await page.mouse.move(box['x'], box['y'])
        await page.mouse.down()
        for i in range(10):
            await page.mouse.move(box['x'], box['y'] + (i * 40))
            await asyncio.sleep(0.2)
        await page.mouse.up()
        print("    ✓")
    except Exception as e:
        print(f"    ✗ {e}")
    
    # 5. 键盘
    print("  5. 键盘PageDown...")
    try:
        await page.click(selector)
        await asyncio.sleep(0.5)
        for _ in range(12):
            await page.keyboard.press('PageDown')
            await asyncio.sleep(0.3)
        print("    ✓")
    except Exception as e:
        print(f"    ✗ {e}")
    
    # 6. 滚轮
    print("  6. 滚轮...")
    try:
        center = await page.evaluate(f'''
            () => {{
                const el = document.querySelector("{selector}");
                if (el) {{
                    const rect = el.getBoundingClientRect();
                    return {{x: rect.left + rect.width/2, y: rect.top + rect.height/2}};
                }}
                return {{x: 960, y: 600}};
            }}
        ''')
        
        await page.mouse.move(center['x'], center['y'])
        await asyncio.sleep(0.5)
        for _ in range(25):
            await page.mouse.wheel(0, 400)
            await asyncio.sleep(0.3)
        print("    ✓")
    except Exception as e:
        print(f"    ✗ {e}")
    
    print("  ✓ 所有滚动方法完成")
    await asyncio.sleep(3)

async def fetch_staff(page, staff_id, date, desc):
    """抓取员工周报"""
    url = f"{BASE_URL}&id={staff_id}&date={date}"
    
    print(f"\n{'='*60}")
    print(f"抓取: {desc} ({date})")
    print(f"{'='*60}")
    
    try:
        # 打开页面
        print("\n[1] 打开页面...")
        await page.goto(url, wait_until='domcontentloaded', timeout=60000)
        await asyncio.sleep(5)
        
        if 'uuap' in page.url:
            print("⚠️ 需要重新登录")
            return {'status': 'login_required'}
        
        print("✓ 页面已打开")
        
        # 等待渲染
        print("\n[2] 等待渲染(5秒)...")
        await asyncio.sleep(5)
        
        # 查找滚动容器
        print("\n[3] 查找滚动容器...")
        containers = await page.evaluate('''
            () => {
                const all = document.querySelectorAll('*');
                const result = [];
                for (let el of all) {
                    const style = window.getComputedStyle(el);
                    if (el.scrollHeight > el.clientHeight + 50 && 
                        el.scrollHeight > 400 &&
                        (style.overflowY === 'auto' || style.overflowY === 'scroll')) {
                        result.push({
                            class: el.className,
                            id: el.id,
                            tag: el.tagName,
                            scrollHeight: el.scrollHeight,
                            clientHeight: el.clientHeight
                        });
                    }
                }
                result.sort((a, b) => b.scrollHeight - a.scrollHeight);
                return result.slice(0, 3);
            }
        ''')
        
        if containers:
            print(f"  找到 {len(containers)} 个可滚动容器:")
            for i, c in enumerate(containers, 1):
                cls = c.get('class', '')
                print(f"    {i}. {c['tag']}.{cls[:40]}: {c['scrollHeight']}px")
            
            # 构建选择器
            first = containers[0]
            if first.get('class'):
                selector = f"[class*='{first['class'].split()[0]}']"
            elif first.get('id'):
                selector = f"#{first['id']}"
            else:
                selector = first['tag'].lower()
        else:
            print("  ⚠️ 未找到可滚动容器，使用body")
            selector = 'body'
        
        print(f"  ✓ 使用选择器: {selector[:50]}")
        
        # 滚动
        print("\n[4] 尝试所有滚动方法...")
        await scroll_all_methods(page, selector)
        
        # 等待
        print("\n[5] 等待动态内容(5秒)...")
        await asyncio.sleep(5)
        
        # 提取内容
        print("\n[6] 提取页面内容...")
        full_text = await page.evaluate('() => document.body.innerText')
        print(f"  ✓ 获取 {len(full_text)} 字符")
        
        # 截图
        print("\n[7] 截图...")
        screenshot = f'{staff_id}_{date}_v2.png'
        await page.screenshot(path=screenshot, full_page=True)
        print(f"  ✓ {screenshot}")
        
        # 保存
        result = {
            'id': staff_id,
            'date': date,
            'desc': desc,
            'container_selector': selector,
            'fetch_time': datetime.now().isoformat(),
            'status': 'success',
            'content': full_text[:50000],
            'screenshot': screenshot,
            'has_okr': any(k in full_text.lower() for k in ['okr', 'objective', 'o1', 'kr1'])
        }
        
        json_file = f'{staff_id}_{date}_v2.json'
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"\n{'='*60}")
        print(f"✓ 抓取完成!")
        print(f"  JSON: {json_file}")
        print(f"  截图: {screenshot}")
        print(f"  包含OKR: {'是' if result['has_okr'] else '否'}")
        print(f"{'='*60}")
        
        print(f"\n内容预览 (前500字符):")
        print("-" * 60)
        print(full_text[:500])
        print("-" * 60)
        
        return result
        
    except Exception as e:
        print(f"\n✗ 抓取失败: {e}")
        import traceback
        traceback.print_exc()
        return {'status': 'error', 'error': str(e)}

async def main():
    last = get_week_date(0)      # 上周
    before = get_week_date(1)    # 上上周
    
    print("="*60)
    print(" "*18 + "百度OKR周报抓取")
    print(" "*16 + "【6种滚动方法】")
    print("="*60)
    print(f"\n员工ID: {MY_ID}")
    print(f"上周: {last}")
    print(f"上上周: {before}")
    print("\n此版本会尝试6种不同的滚动方法")
    
    async with async_playwright() as p:
        print("\n" + "="*60)
        print("[系统] 启动Chrome浏览器...")
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page(viewport={'width': 1920, 'height': 1080})
        print("✓ Chrome已启动")
        
        print("\n[系统] 访问OKR登录页...")
        await page.goto("https://okr.baidu-int.com", wait_until='domcontentloaded')
        print("✓ 页面已加载")
        
        print("\n" + "="*60)
        print("⚠️  请先登录OKR系统")
        print("="*60)
        input("\n登录完成后按 Enter 键继续...")
        
        # 验证登录
        current_url = page.url
        if 'uuap' in current_url or 'login' in current_url:
            print("\n✗ 检测到仍在登录页面，请重新登录")
            await browser.close()
            return
        
        print(f"\n✓ 当前页面: {current_url[:50]}...")
        print("✓ 登录验证通过")
        
        # 抓取上周
        result1 = await fetch_staff(page, MY_ID, last, "上周周报")
        
        # 如果成功，抓取上上周
        if result1.get('status') == 'success':
            print("\n" + "="*60)
            input("准备抓取上上周，按 Enter 键继续...")
            
            result2 = await fetch_staff(page, MY_ID, before, "上上周周报")
        
        # 汇总
        print("\n" + "="*60)
        print("全部完成!")
        print("="*60)
        print(f"\n生成的文件:")
        print(f"  1. {MY_ID}_{last}_v2.png (上周截图)")
        print(f"  2. {MY_ID}_{last}_v2.json (上周数据)")
        print(f"  3. {MY_ID}_{before}_v2.png (上上周截图)")
        print(f"  4. {MY_ID}_{before}_v2.json (上上周数据)")
        print(f"\n请检查截图文件，确认是否包含完整的OKR部分内容")
        
        await browser.close()
        input("\n按 Enter 键退出...")

if __name__ == '__main__':
    asyncio.run(main())
