#!/usr/bin/env python3
"""
单人测试版 - 只抓取本人周报
完整流程：启动Chrome -> 等待登录 -> 抓取内容
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

async def fetch_single_staff(page, staff_id, date, week_desc):
    """抓取单个员工的周报"""
    url = f"{BASE_URL}&id={staff_id}&date={date}"
    print(f"\n{'='*60}")
    print(f"抓取: {week_desc}")
    print(f"日期: {date}")
    print(f"URL: {url}")
    print(f"{'='*60}")
    
    try:
        # 1. 访问页面
        print("\n[步骤1] 访问页面...")
        await page.goto(url, wait_until='domcontentloaded', timeout=60000)
        await asyncio.sleep(5)
        
        # 检查是否被重定向到登录页
        current_url = page.url
        if 'uuap' in current_url or 'login' in current_url:
            print("⚠️  页面被重定向到登录页，会话可能已失效")
            return {'id': staff_id, 'date': date, 'status': 'login_required'}
        
        print(f"  ✓ 当前页面: {current_url[:80]}...")
        
        # 2. 等待OKR相关元素
        print("\n[步骤2] 等待OKR内容加载...")
        try:
            # 尝试等待常见的OKR元素
            await page.wait_for_selector('text=/OKR|目标|Objective|O1|KR1/', timeout=15000)
            print("  ✓ 检测到OKR相关元素")
        except:
            print("  ⚠️  未检测到OKR元素，继续等待...")
        
        await asyncio.sleep(5)
        
        # 3. 滚动页面触发懒加载
        print("\n[步骤3] 滚动页面触发OKR加载...")
        scroll_positions = [0.2, 0.4, 0.6, 0.8, 1.0, 0.5, 0]
        for i, pos in enumerate(scroll_positions, 1):
            await page.evaluate(f'() => {{ window.scrollTo(0, document.body.scrollHeight * {pos}); }}')
            await asyncio.sleep(1.5)
            print(f"  滚动 {i}/{len(scroll_positions)}...", end="\r")
        print("  ✓ 滚动完成")
        
        # 4. 再次等待
        print("\n[步骤4] 等待动态内容...")
        await asyncio.sleep(5)
        
        # 5. 提取内容（多种方式）
        print("\n[步骤5] 提取页面内容...")
        
        # 方式A: 页面完整文本
        print("  获取 innerText...")
        full_text = await page.evaluate('() => document.body.innerText')
        print(f"  ✓ innerText: {len(full_text)} 字符")
        
        # 方式B: 页面HTML
        print("  获取 HTML...")
        full_html = await page.content()
        print(f"  ✓ HTML: {len(full_html)} 字符")
        
        # 方式C: 通过CSS选择器查找OKR容器
        print("  查找OKR容器...")
        container_text = ""
        okr_selectors = [
            '[class*="okr"]',
            '[class*="objective"]',
            '[class*="weekly-content"]',
            '[class*="main-content"]',
            '.content-wrapper',
            '#app',
            'main',
            'article'
        ]
        
        for selector in okr_selectors:
            try:
                elements = await page.query_selector_all(selector)
                for elem in elements:
                    text = await elem.inner_text()
                    if text and len(text) > 200:  # 找内容较多的容器
                        container_text = text
                        print(f"  ✓ 找到容器: '{selector}' ({len(text)} 字符)")
                        break
                if container_text:
                    break
            except:
                continue
        
        if not container_text:
            print("  ⚠️  未找到特定容器")
        
        # 6. 截图保存（用于人工核对）
        print("\n[步骤6] 保存截图...")
        screenshot_file = f'{staff_id}_{date}_screenshot.png'
        await page.screenshot(path=screenshot_file, full_page=True)
        print(f"  ✓ 截图已保存: {screenshot_file}")
        
        # 7. 分析内容
        print("\n[步骤7] 分析内容...")
        content_lower = full_text.lower()
        analysis = {
            'has_okr': 'okr' in content_lower or 'objective' in content_lower,
            'has_o1': 'o1' in content_lower or '01' in content_lower,
            'has_kr': 'kr1' in content_lower or 'key result' in content_lower,
            'has_weekly': '本周' in full_text or 'weekly' in content_lower,
            'has_summary': '总结' in full_text or 'summary' in content_lower
        }
        
        print("  内容分析:")
        for key, value in analysis.items():
            print(f"    {key}: {'✓' if value else '✗'}")
        
        # 保存结果
        result = {
            'id': staff_id,
            'date': date,
            'week_desc': week_desc,
            'url': url,
            'fetch_time': datetime.now().isoformat(),
            'status': 'success',
            'page_url_after_load': current_url,
            'content_analysis': analysis,
            'content': {
                'innerText': {
                    'length': len(full_text),
                    'preview': full_text[:1000] if full_text else '',
                    'full': full_text[:50000]
                },
                'html': {
                    'length': len(full_html),
                    'preview': full_html[:1000] if full_html else ''
                },
                'container': {
                    'length': len(container_text),
                    'content': container_text[:30000] if container_text else ''
                }
            },
            'screenshot': screenshot_file
        }
        
        json_file = f'{staff_id}_{date}_full.json'
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"\n{'='*60}")
        print(f"✓ 抓取成功!")
        print(f"  JSON文件: {json_file}")
        print(f"  截图文件: {screenshot_file}")
        print(f"{'='*60}")
        
        # 显示预览
        print(f"\n内容预览 (innerText前500字符):")
        print("-" * 60)
        print(full_text[:500] if full_text else "(无内容)")
        print("-" * 60)
        
        return result
        
    except Exception as e:
        print(f"\n✗ 抓取失败: {e}")
        import traceback
        traceback.print_exc()
        return {
            'id': staff_id,
            'date': date,
            'status': 'error',
            'error': str(e)
        }

async def main():
    last_week = get_week_date(0)
    week_before_last = get_week_date(1)
    
    print("=" * 70)
    print(" " * 20 + "百度OKR周报抓取")
    print(" " * 15 + "【单人测试版 - 只抓本人】")
    print("=" * 70)
    print(f"\n员工ID: {MY_ID}")
    print(f"上周: {last_week}")
    print(f"上上周: {week_before_last}")
    print("\n" + "=" * 70)
    
    async with async_playwright() as p:
        # 启动Chrome
        print("\n【系统准备】")
        print("启动Chrome浏览器...")
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page(viewport={'width': 1920, 'height': 1080})
        print("✓ Chrome已启动")
        
        # 访问OKR并登录
        print("\n【登录步骤】")
        print("访问OKR系统...")
        await page.goto("https://okr.baidu-int.com", wait_until='domcontentloaded')
        print("✓ 页面已加载")
        
        print("\n" + "=" * 70)
        print("⚠️  请在Chrome浏览器中完成登录:")
        print("   1. 在打开的Chrome窗口中输入用户名密码")
        print("   2. 确保能看到OKR主页")
        print("   3. 可以手动打开你的周报页面确认正常")
        print("=" * 70)
        
        input("\n登录完成后，按 Enter 键继续...")
        
        # 验证登录状态
        current_url = page.url
        if 'uuap' in current_url or 'login' in current_url:
            print("\n✗ 检测到仍在登录页面，请重新登录")
            await browser.close()
            return
        
        print(f"\n✓ 当前页面: {current_url[:60]}...")
        print("✓ 登录验证通过")
        
        # 抓取上周
        result1 = await fetch_single_staff(page, MY_ID, last_week, "上周周报")
        
        # 间隔
        await asyncio.sleep(3)
        
        # 抓取上上周
        result2 = await fetch_single_staff(page, MY_ID, week_before_last, "上上周周报")
        
        # 保存汇总
        print("\n" + "=" * 70)
        print("【汇总结果】")
        print("=" * 70)
        
        summary = {
            'test_time': datetime.now().isoformat(),
            'my_id': MY_ID,
            'results': [result1, result2]
        }
        
        with open('test_summary.json', 'w', encoding='utf-8') as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)
        
        success_count = sum(1 for r in [result1, result2] if r.get('status') == 'success')
        
        print(f"\n抓取完成!")
        print(f"  成功: {success_count}/2")
        print(f"\n生成的文件:")
        print(f"  1. {MY_ID}_{last_week}_screenshot.png (上周截图)")
        print(f"  2. {MY_ID}_{last_week}_full.json (上周数据)")
        print(f"  3. {MY_ID}_{week_before_last}_screenshot.png (上上周截图)")
        print(f"  4. {MY_ID}_{week_before_last}_full.json (上上周数据)")
        print(f"  5. test_summary.json (汇总)")
        print(f"\n请重点检查截图文件，确认是否包含OKR部分内容")
        
        await browser.close()
        
        print("\n" + "=" * 70)
        input("按 Enter 键退出...")

if __name__ == '__main__':
    asyncio.run(main())
