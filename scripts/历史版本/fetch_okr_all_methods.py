#!/usr/bin/env python3
"""
百度OKR周报抓取 - 多方法滚动尝试
尝试：拖拽滚动条、jQuery动画、scrollIntoView等
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

async def try_all_scroll_methods(page, container_selector):
    """尝试所有可能的滚动方法"""
    print(f"  尝试多种滚动方法...")
    
    # 方法1: 原生scrollTop
    print("  方法1: scrollTop...")
    try:
        await page.evaluate(f'''
            () => {{
                const el = document.querySelector("{container_selector}");
                if (el) {{
                    const scroll = () => {{
                        for (let i = 0; i <= 10; i++) {{
                            setTimeout(() => {{
                                el.scrollTop = (el.scrollHeight - el.clientHeight) * (i / 10);
                            }}, i * 200);
                        }}
                    }};
                    scroll();
                }}
            }}
        ''')
        await asyncio.sleep(3)
        print("    ✓ scrollTop完成")
    except Exception as e:
        print(f"    ✗ {e}")
    
    # 方法2: scrollIntoView
    print("  方法2: scrollIntoView...")
    try:
        await page.evaluate(f'''
            () => {{
                const container = document.querySelector("{container_selector}");
                if (container) {{
                    const children = container.children;
                    if (children.length > 2) {{
                        children[Math.floor(children.length/2)].scrollIntoView({{behavior: 'smooth', block: 'center'}});
                    }}
                }}
            }}
        ''')
        await asyncio.sleep(2)
        print("    ✓ scrollIntoView完成")
    except Exception as e:
        print(f"    ✗ {e}")
    
    # 方法3: jQuery animate (如果页面有jQuery)
    print("  方法3: jQuery animate...")
    try:
        await page.evaluate(f'''
            () => {{
                if (typeof $ !== 'undefined') {{
                    $("{container_selector}").animate({{scrollTop: $("{container_selector}")[0].scrollHeight}}, 2000);
                }} else {{
                    // 原生模拟smooth scroll
                    const el = document.querySelector("{container_selector}");
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
            }}
        ''')
        await asyncio.sleep(3)
        print("    ✓ jQuery/原生动画完成")
    except Exception as e:
        print(f"    ✗ {e}")
    
    # 方法4: 鼠标拖拽滚动条
    print("  方法4: 鼠标拖拽滚动条...")
    try:
        # 获取容器位置
        box = await page.evaluate(f'''
            () => {{
                const el = document.querySelector("{container_selector}");
                if (el) {{
                    const rect = el.getBoundingClientRect();
                    return {{
                        x: rect.right - 10,
                        y: rect.top + 50,
                        height: rect.height
                    }};
                }}
                return null;
            }}
        ''')
        
        if box:
            # 移动鼠标到滚动条位置
            await page.mouse.move(box['x'], box['y'])
            await asyncio.sleep(0.5)
            # 按下并向下拖拽
            await page.mouse.down()
            for i in range(10):
                await page.mouse.move(box['x'], box['y'] + (i * 50))
                await asyncio.sleep(0.2)
            await page.mouse.up()
            print("    ✓ 拖拽完成")
        else:
            print("    ⚠️ 无法获取容器位置")
    except Exception as e:
        print(f"    ✗ {e}")
    
    # 方法5: 键盘事件
    print("  方法5: 键盘事件...")
    try:
        # 先点击容器获得焦点
        await page.click(container_selector)
        await asyncio.sleep(0.5)
        # 发送多次Page Down
        for i in range(15):
            await page.keyboard.press('PageDown')
            await asyncio.sleep(0.3)
        print("    ✓ 键盘PageDown完成")
    except Exception as e:
        print(f"    ✗ {e}")
    
    # 方法6: wheel事件在特定位置
    print("  方法6: 特定位置滚轮...")
    try:
        # 获取容器中心位置
        center = await page.evaluate(f'''
            () => {{
                const el = document.querySelector("{container_selector}");
                if (el) {{
                    const rect = el.getBoundingClientRect();
                    return {{
                        x: rect.left + rect.width / 2,
                        y: rect.top + rect.height / 2
                    }};
                }}
                return {{x: 960, y: 540}};
            }}
        ''')
        
        await page.mouse.move(center['x'], center['y'])
        await asyncio.sleep(0.5)
        
        for i in range(30):
            await page.mouse.wheel(0, 300)
            await asyncio.sleep(0.4)
        print("    ✓ 滚轮完成")
    except Exception as e:
        print(f"    ✗ {e}")
    
    print("  所有滚动方法尝试完成")
    await asyncio.sleep(3)

async def find_scroll_container_with_debug(page):
    """查找滚动容器（带调试信息）"""
    print("  查找滚动容器...")
    
    # 获取所有div的滚动信息
    containers = await page.evaluate('''() => {
        const allElements = document.querySelectorAll('*');
        const scrollables = [];
        
        for (let el of allElements) {
            const style = window.getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            
            // 检查是否可滚动
            const canScrollY = el.scrollHeight > el.clientHeight + 10;
            const overflowY = style.overflowY;
            const hasScrollbar = canScrollY && (overflowY === 'auto' || overflowY === 'scroll' || overflowY === 'overlay');
            
            if (hasScrollbar && rect.height > 200) {
                scrollables.push({
                    tag: el.tagName,
                    class: el.className,
                    id: el.id,
                    scrollHeight: el.scrollHeight,
                    clientHeight: el.clientHeight,
                    scrollTop: el.scrollTop,
                    overflowY: overflowY,
                    rect: {
                        top: rect.top,
                        left: rect.left,
                        width: rect.width,
                        height: rect.height
                    }
                });
            }
        }
        
        // 按高度排序
        scrollables.sort((a, b) => b.scrollHeight - a.scrollHeight);
        return scrollables.slice(0, 5);
    }''')
    
    print(f"  找到 {len(containers)} 个可滚动元素:")
    for i, c in enumerate(containers, 1):
        print(f"    {i}. {c['tag']}.{c['class'][:40]}")
        print(f"       滚动: {c['scrollHeight']}/{c['clientHeight']}, 位置: ({int(c['rect']['left'])}, {int(c['rect']['top'])})")
    
    if containers:
        # 使用第一个（最大的）
        first = containers[0]
        selector = first['class'] ? `[class*="${first['class'].split()[0]}"]` : (first['id'] ? `#${first['id']}` : first['tag'].toLowerCase())
        print(f"  ✓ 选择: {selector}")
        return selector, first
    
    return 'body', None

async def fetch_staff_all_methods(page, staff_id, date, desc):
    """使用所有方法抓取"""
    url = f"{BASE_URL}&id={staff_id}&date={date}"
    
    print(f"\n{'='*70}")
    print(f"抓取: {desc} ({date})")
    print(f"{'='*70}")
    
    try:
        print("\n[1] 打开页面...")
        await page.goto(url, wait_until='domcontentloaded', timeout=60000)
        await asyncio.sleep(5)
        
        if 'uuap' in page.url:
            print("⚠️ 需要登录")
            return {'status': 'login_required'}
        
        print("✓ 页面打开")
        
        print("\n[2] 等待渲染...")
        await asyncio.sleep(5)
        
        print("\n[3] 查找滚动容器...")
        container, container_info = await find_scroll_container_with_debug(page)
        
        print("\n[4] 尝试所有滚动方法...")
        await try_all_scroll_methods(page, container)
        
        print("\n[5] 等待内容稳定...")
        await asyncio.sleep(5)
        
        print("\n[6] 提取内容...")
        full_text = await page.evaluate('() => document.body.innerText')
        print(f"✓ 获取 {len(full_text)} 字符")
        
        # 尝试获取容器内容
        container_text = ""
        try:
            elem = await page.query_selector(container)
            if elem:
                container_text = await elem.inner_text()
                print(f"✓ 容器内容: {len(container_text)} 字符")
        except:
            pass
        
        print("\n[7] 截图...")
        screenshot = f'{staff_id}_{date}_all.png'
        await page.screenshot(path=screenshot, full_page=True)
        print(f"✓ {screenshot}")
        
        # 保存
        result = {
            'id': staff_id,
            'date': date,
            'desc': desc,
            'container': container,
            'container_info': container_info,
            'fetch_time': datetime.now().isoformat(),
            'status': 'success',
            'content': {
                'full_text': full_text[:50000],
                'container_text': container_text[:30000] if container_text else ''
            },
            'screenshot': screenshot,
            'has_okr': any(k in full_text.lower() for k in ['okr', 'objective', 'o1', 'kr1'])
        }
        
        with open(f'{staff_id}_{date}_all.json', 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"\n{'='*70}")
        print(f"✓ 完成!")
        print(f"  JSON: {staff_id}_{date}_all.json")
        print(f"  截图: {screenshot}")
        print(f"{'='*70}")
        
        print(f"\n内容预览 (前600字符):")
        print(full_text[:600])
        
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
    print(" "*16 + "【多方法滚动版】")
    print("="*70)
    print(f"\n员工ID: {MY_ID}")
    print(f"上周: {last}")
    print(f"上上周: {before}")
    print("\n此版本会尝试6种不同的滚动方法")
    
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
        r1 = await fetch_staff_all_methods(page, MY_ID, last, "上周周报")
        
        if r1.get('status') == 'success':
            await asyncio.sleep(3)
            r2 = await fetch_staff_all_methods(page, MY_ID, before, "上上周周报")
        
        print("\n" + "="*70)
        print("全部完成!")
        print("="*70)
        print(f"\n文件:")
        print(f"  1. {MY_ID}_{last}_all.png")
        print(f"  2. {MY_ID}_{last}_all.json")
        print(f"  3. {MY_ID}_{before}_all.png")
        print(f"  4. {MY_ID}_{before}_all.json")
        print(f"\n请检查截图是否包含OKR部分")
        
        await browser.close()
        input("\n按Enter退出...")

if __name__ == '__main__':
    asyncio.run(main())
