#!/usr/bin/env python3
"""
百度OKR周报自动提交脚本 - V1.3 (参考V23版本)
基于fetch_okr_v23.py的改进机制：
1. 首次页面等待5秒
2. 导航中断错误处理（登录跳转）
3. 失败重试机制（最多1次）
4. 更完善的登录状态检查

编辑页面URL格式：
https://okr.baidu-int.com/pages/weekly.html#/home?date=s2026-02-16&mandatorId=s&isEdit=true
"""

import asyncio
import socket
import sys
from datetime import datetime, timedelta
from playwright.async_api import async_playwright, Error as PlaywrightError

# ==================== 配置区域 ====================

MY_ID = "s673090"
EDIT_BASE_URL = "https://okr.baidu-int.com/pages/weekly.html#/home"

def get_default_date():
    """获取默认提交日期（上周一）"""
    today = datetime.now()
    last_monday = today - timedelta(days=today.weekday() + 7)
    return f"s{last_monday.strftime('%Y-%m-%d')}"

# 第一部分：本周工作总结
WEEKLY_SUMMARY_CONTENT = """
1. 完成XX项目的需求评审和技术方案设计
2. 推进YY模块的开发工作，完成核心功能80%
3. 修复ZZ系统3个线上bug，保障系统稳定性
"""

# 第二部分：OKR本周工作内容
OKR_WEEKLY_WORK = {
    "O1": {
        "KR1": "完成需求文档编写，与产品团队确认3个核心功能点",
        "KR2": "完成技术方案评审，确定微服务拆分策略"
    },
    "O2": {
        "KR1": "修复登录模块的内存泄漏问题，已通过测试验证",
        "KR2": "优化数据库查询性能，响应时间降低40%"
    }
}

# 第三部分：其他内容
METRICS_CONTENT = {
    "业务核心指标": "本周系统可用性99.9%，接口平均响应时间120ms",
    "主要项目": "XX项目进入测试阶段，YY项目完成技术预研",
    "下周重点工作": "1. 完成XX项目上线 2. 启动YY项目开发"
}

DRY_RUN = False  # True=仅预览, False=实际填写

# ==================== 网络检测工具 ====================

def check_network_connection():
    """检测网络连接状态"""
    print("\n[网络检测] 检查网络连接...")

    try:
        ip = socket.gethostbyname("okr.baidu-int.com")
        print(f"  [OK] 内网域名解析正常: okr.baidu-int.com -> {ip}")
        return True
    except socket.gaierror as e:
        print(f"  [X] 内网域名解析失败: {e}")
        print("  [!] 请检查是否已连接公司VPN或内网")
        return False
    except Exception as e:
        print(f"  [X] 网络检测异常: {e}")
        return False


# ==================== 浏览器操作（与V23保持一致）====================

async def find_scroll_container(page):
    """查找可滚动容器（与V23一致）"""
    containers = await page.evaluate("""() => {
        const all = document.querySelectorAll("*");
        const result = [];
        for (let el of all) {
            const style = window.getComputedStyle(el);
            if (el.scrollHeight > el.clientHeight + 100 && 
                el.scrollHeight > 500 &&
                (style.overflowY === 'auto' || style.overflowY === 'scroll')) {
                const rect = el.getBoundingClientRect();
                result.push({
                    class: el.className,
                    id: el.id,
                    tag: el.tagName,
                    scrollHeight: el.scrollHeight,
                    x: rect.left + rect.width / 2,
                    y: rect.top + rect.height / 2
                });
            }
        }
        result.sort((a, b) => b.scrollHeight - a.scrollHeight);
        return result.slice(0, 1);
    }""")

    if containers:
        first = containers[0]
        if first.get("class"):
            selector = f"[class*='{first['class'].split()[0]}']"
        elif first.get("id"):
            selector = f"#{first['id']}"
        else:
            selector = first["tag"].lower()
        return selector, (first.get("x", 960), first.get("y", 600))
    return "body", (960, 600)


async def scroll_container_v23(page, selector, center):
    """滚动加载内容（与V23保持一致，30次滚动+5秒等待）"""
    print("    滚动加载内容...", end=" ", flush=True)
    await page.mouse.move(center[0], center[1])
    await asyncio.sleep(0.5)

    for i in range(30):
        await page.mouse.wheel(0, 400)
        await asyncio.sleep(0.3)
        if i % 10 == 0:
            print(f"{i*400}px", end=" ", flush=True)
    print("OK")

    try:
        await page.click(selector)
        await asyncio.sleep(0.5)
        for _ in range(5):
            await page.keyboard.press("PageDown")
            await asyncio.sleep(0.3)
    except:
        pass

    print("    等待动态内容渲染...", end=" ", flush=True)
    await asyncio.sleep(5)
    print("OK")


# ==================== 周报提交核心功能 ====================

async def fill_textarea_by_label(page, label_text, content):
    """根据标签文本找到对应的textarea并填写内容"""
    if not content or not content.strip():
        print(f"  [跳过] 内容为空: {label_text}")
        return False

    content = content.strip()
    print(f"  正在填写: {label_text[:30]}...")

    try:
        escaped_label = label_text.replace('"', '\\"')
        escaped_content = content.replace("\\", "\\\\").replace("`", "\\`").replace("$", "\\$")

        result = await page.evaluate(f"""() => {{
            const labelText = "{escaped_label}";
            const content = `{escaped_content}`;

            const labels = Array.from(document.querySelectorAll('label, span, div, h1, h2, h3, h4, p'));
            for (let label of labels) {{
                if (label.textContent.includes(labelText)) {{
                    let parent = label.parentElement;
                    for (let i = 0; i < 5 && parent; i++) {{
                        const textarea = parent.querySelector('textarea');
                        if (textarea) {{
                            textarea.value = content;
                            textarea.dispatchEvent(new Event('input', {{ bubbles: true }}));
                            textarea.dispatchEvent(new Event('change', {{ bubbles: true }}));
                            return 'found_by_label';
                        }}
                        parent = parent.parentElement;
                    }}
                }}
            }}

            if (labelText.match(/^O\\d+/) || labelText.match(/^KR\\d+/)) {{
                const allElements = Array.from(document.querySelectorAll("*"));
                for (let el of allElements) {{
                    if (el.textContent && el.textContent.includes(labelText)) {{
                        let parent = el.parentElement;
                        for (let i = 0; i < 4 && parent; i++) {{
                            const textarea = parent.querySelector('textarea');
                            if (textarea) {{
                                textarea.value = content;
                                textarea.dispatchEvent(new Event('input', {{ bubbles: true }}));
                                return 'found_by_okr';
                            }}
                            parent = parent.parentElement;
                        }}
                    }}
                }}
            }}

            return 'not_found';
        }}""")

        if result != "not_found":
            print(f"    [成功] JavaScript注入 ({result})")
            return True

        print(f"    [失败] 未找到输入框")
        return False

    except Exception as e:
        print(f"    [错误] {e}")
        return False


async def fill_weekly_summary(page, content):
    """填写本周工作总结（第一部分）"""
    print("\n[1/3] 填写本周工作总结...")
    labels = ["本周工作总结", "工作总结", "本周工作"]
    for label in labels:
        if await fill_textarea_by_label(page, label, content):
            return True
    return False


async def fill_okr_weekly_work(page, okr_work_dict):
    """填写OKR本周工作内容（第二部分）"""
    print("\n[2/3] 填写OKR本周工作...")
    success_count = 0
    fail_count = 0

    for o_id, kr_dict in okr_work_dict.items():
        print(f"  处理 {o_id}...")
        for kr_id, work_content in kr_dict.items():
            labels = [f"{o_id}{kr_id}", f"{o_id} {kr_id}", kr_id, f"{o_id}-{kr_id}"]
            filled = False
            for label in labels:
                if await fill_textarea_by_label(page, label, work_content):
                    success_count += 1
                    filled = True
                    break
            if not filled:
                print(f"    [失败] 无法定位 {o_id}/{kr_id}")
                fail_count += 1

    print(f"\n  统计: 成功{success_count}条, 失败{fail_count}条")
    return fail_count == 0


async def fill_other_metrics(page, metrics_dict):
    """填写其他内容（第三部分）"""
    print("\n[3/3] 填写其他内容...")
    for title, content in metrics_dict.items():
        await fill_textarea_by_label(page, title, content)
        await asyncio.sleep(0.3)
    return True


async def submit_single_weekly_v23(page, target_date=None, summary_content=None,
                                   okr_work=None, metrics=None, dry_run=True,
                                   is_first_page=False):
    """
    提交周报的核心函数（参考V23的fetch_single_staff_v10）
    包含：首次页面等待5秒、导航中断处理、登录检测
    """
    target_date = target_date or get_default_date()
    summary_content = summary_content or WEEKLY_SUMMARY_CONTENT
    okr_work = okr_work or OKR_WEEKLY_WORK
    metrics = metrics or METRICS_CONTENT

    edit_url = f"{EDIT_BASE_URL}?date={target_date}&mandatorId={MY_ID}&isEdit=true"

    print("="*60)
    print(f"目标日期: {target_date}")
    print(f"编辑模式: {'预览' if dry_run else '实际提交'}")
    print(f"URL: {edit_url}")
    print("="*60)

    print("\n正在加载编辑页面...")

    try:
        # 使用 try-except 捕获导航被中断的情况（通常是重定向到登录页）
        try:
            await page.goto(edit_url, wait_until='domcontentloaded', timeout=60000)
        except Exception as nav_error:
            # 检查是否跳转到了登录页
            current_url = page.url
            if 'uuap' in current_url:
                print("\n[!] 会话失效，需要重新登录")
                return {'status': 'login_required', 'url': current_url}
            # 其他导航错误，继续尝试
            print(f"\n[!] 导航警告: {nav_error}")

        # ===== V23改进：首次页面多等待5秒 =====
        if is_first_page:
            print("(首次页面等待5秒)...", end=" ", flush=True)
            await asyncio.sleep(5)
            print("OK")
        else:
            await asyncio.sleep(2)

        # 再次检查URL（可能跳转到了登录页）
        current_url = page.url
        if 'uuap' in current_url:
            print("\n[!] 检测到登录页面，需要登录")
            return {'status': 'login_required', 'url': current_url}

        print("[OK] 页面加载完成")

    except Exception as e:
        print(f"\n[X] 页面加载失败: {e}")
        return {'status': 'error', 'error': str(e)}

    print("\n滚动加载页面内容...")
    selector, center = await find_scroll_container(page)
    await scroll_container_v23(page, selector, center)

    if dry_run:
        print("\n" + "="*60)
        print("【预览模式】以下内容将被填写：")
        print("="*60)
        print(f"\n1. 本周工作总结:\n{summary_content[:200]}...")
        print(f"\n2. OKR工作项数: {sum(len(v) for v in okr_work.values())}")
        print(f"\n3. 其他内容项数: {len(metrics)}")
        print("\n" + "="*60)
        print("预览完成，未实际填写")
        return {'status': 'preview', 'date': target_date}

    print("\n开始填写内容...")
    results = {
        'part1_summary': await fill_weekly_summary(page, summary_content),
        'part2_okr': await fill_okr_weekly_work(page, okr_work),
        'part3_metrics': await fill_other_metrics(page, metrics)
    }

    print("\n" + "="*60)
    print("填写完成统计：")
    print(f"  本周工作总结: {'OK' if results['part1_summary'] else 'FAIL'}")
    print(f"  OKR本周工作: {'OK' if results['part2_okr'] else 'FAIL'}")
    print(f"  其他内容: {'OK' if results['part3_metrics'] else 'FAIL'}")
    print("="*60)

    return {'status': 'success', 'date': target_date, 'results': results}


async def submit_weekly_with_retry_v23(page, target_date=None, summary_content=None,
                                       okr_work=None, metrics=None, dry_run=True):
    """
    带重试机制的周报提交（参考V23的fetch_staff_v10_with_retry）
    失败后重试（最多1次）
    """
    # 首次尝试（标记为首次页面）
    result = await submit_single_weekly_v23(
        page, target_date, summary_content, okr_work, metrics, dry_run,
        is_first_page=True
    )

    # 如果失败且不是登录问题，重试一次
    if result['status'] == 'error':
        print(f"\n{'='*60}")
        print("[重试机制] 首次尝试失败，准备重试...")
        print(f"{'='*60}")
        await asyncio.sleep(2)

        result = await submit_single_weekly_v23(
            page, target_date, summary_content, okr_work, metrics, dry_run,
            is_first_page=False  # 重试时不再额外等待
        )

        if result['status'] == 'success':
            print("\n[OK] 重试成功")
        else:
            print("\n[X] 重试失败")

    return result


# ==================== 主流程（参考V23）====================

async def main():
    """主流程（参考V23脚本结构）"""

    # 先进行网络检测
    if not check_network_connection():
        print("\n[!] 网络检查失败，请连接公司VPN或内网")
        response = input("\n是否仍要继续尝试? (y/N): ")
        if response.lower() != 'y':
            print("退出程序")
            return
        print("\n[!] 用户选择继续，尝试访问...")

    target_date = get_default_date()

    print("="*60)
    print("百度OKR周报自动提交 - V1.3 (参考V23版本)")
    print("="*60)
    print("\n【V1.3改进】")
    print("  ✓ 首次页面等待5秒（参考V23）")
    print("  ✓ 导航中断错误处理（登录跳转）")
    print("  ✓ 失败重试机制（最多1次）")
    print("  ✓ 更完善的登录状态检查")
    print(f"\n[配置]")
    print(f"  用户ID: {MY_ID}")
    print(f"  目标日期: {target_date}")
    print(f"  运行模式: {'预览' if DRY_RUN else '实际填写'}")
    print(f"\n[内容]")
    print(f"  - 工作总结: {len(WEEKLY_SUMMARY_CONTENT.strip())} 字符")
    print(f"  - OKR工作: {sum(len(v) for v in OKR_WEEKLY_WORK.values())} 条")
    print(f"  - 其他内容: {len(METRICS_CONTENT)} 项")

    print(f"\n{'='*60}")
    print("[系统] 启动Chrome...")

    try:
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=False)
            page = await browser.new_page(viewport={"width": 1920, "height": 1080})
            print("[OK] Chrome已启动")

            print("\n[系统] 访问OKR...")
            try:
                await page.goto("https://okr.baidu-int.com", wait_until="domcontentloaded", timeout=30000)
                print("[OK] 页面已加载")
            except PlaywrightError as e:
                if "ERR_NAME_NOT_RESOLVED" in str(e):
                    print(f"\n[X] 网络错误: 无法解析域名")
                    print("[!] 请确保已连接到公司VPN或内网")
                    await browser.close()
                    return
                else:
                    raise

            print("\n" + "="*60)
            print("[!] 请先登录OKR系统")
            print("="*60)
            input("\n登录完成后按 Enter 继续...")

            # 检查登录状态
            if 'uuap' in page.url:
                print("\n[X] 未登录成功，退出")
                await browser.close()
                return
            print("[OK] 登录验证通过")

            # 执行周报填写（带重试机制）
            result = await submit_weekly_with_retry_v23(
                page=page,
                target_date=target_date,
                summary_content=WEEKLY_SUMMARY_CONTENT,
                okr_work=OKR_WEEKLY_WORK,
                metrics=METRICS_CONTENT,
                dry_run=DRY_RUN
            )

            if result['status'] == 'login_required':
                print("\n[!] 会话失效或需要登录")
                print(f"[!] 当前页面: {result.get('url', '未知')}")
                print("[!] 请在浏览器中完成登录后，重新运行脚本")
            elif result['status'] == 'error':
                print(f"\n[X] 错误: {result.get('error')}")
            elif result['status'] == 'preview':
                print("\n[OK] 预览完成")
                print("提示: 修改 DRY_RUN = False 后重新运行以实际填写")
            else:
                print("\n[OK] 填写完成")
                print("[!] 请手动检查内容并点击保存按钮")

            print("\n" + "="*60)
            input("\n按 Enter 关闭浏览器...")
            await browser.close()

    except Exception as e:
        print(f"\n[X] 严重错误: {e}")
        import traceback
        traceback.print_exc()
        return


if __name__ == "__main__":
    asyncio.run(main())