#!/usr/bin/env python3
"""
百度OKR周报自动提交脚本 - V1.5 (完整功能版)

修复内容：
1. 使用更精确的定位策略（page.get_by_text）
2. 明确配置文件路径提示
3. 添加填写验证机制（填写后检查内容）
4. 自动点击"保存"或"提交"按钮
5. 添加调试截图功能

配置文件：
- 默认配置在脚本顶部的"配置区域"
- 也可从外部JSON文件加载（见load_config_from_file函数）

编辑页面URL格式：
https://okr.baidu-int.com/pages/weekly.html#/home?date=s2026-02-16&mandatorId=s&isEdit=true
"""

import asyncio
import socket
import json
import os
import sys
from datetime import datetime, timedelta
from pathlib import Path
from playwright.async_api import async_playwright, Error as PlaywrightError

# ==================== 配置区域（默认配置）====================

MY_ID = "s673090"
EDIT_BASE_URL = "https://okr.baidu-int.com/pages/weekly.html#/home"

def get_default_date():
    """获取默认提交日期（上周一）"""
    today = datetime.now()
    last_monday = today - timedelta(days=today.weekday() + 7)
    return f"s{last_monday.strftime('%Y-%m-%d')}"

# 第一部分：本周工作总结（必填）
WEEKLY_SUMMARY_CONTENT = """
1. 完成XX项目的需求评审和技术方案设计
2. 推进YY模块的开发工作，完成核心功能80%
3. 修复ZZ系统3个线上bug，保障系统稳定性
"""

# 第二部分：OKR本周工作内容（必填）
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

# 第三部分：其他内容（可选）
METRICS_CONTENT = {
    "业务核心指标": "本周系统可用性99.9%，接口平均响应时间120ms",
    "主要项目": "XX项目进入测试阶段，YY项目完成技术预研",
    "下周重点工作": "1. 完成XX项目上线 2. 启动YY项目开发"
}

# 运行模式
DRY_RUN = False  # True=仅预览, False=实际填写并提交
AUTO_SUBMIT = True  # True=自动点击保存/提交按钮, False=手动提交

# ==================== 外部配置文件加载 ====================

def load_config_from_file():
    """
    从外部JSON文件加载配置
    配置文件路径优先级：
    1. 当前目录下的 weekly_config.json
    2. 用户主目录下的 .okr_submit_config.json
    3. 环境变量 OKR_CONFIG_PATH 指定的路径
    """
    config_paths = [
        Path("weekly_config.json"),  # 当前目录
        Path.home() / ".okr_submit_config.json",  # 用户主目录
    ]

    # 检查环境变量
    env_path = os.environ.get("OKR_CONFIG_PATH")
    if env_path:
        config_paths.insert(0, Path(env_path))

    for config_path in config_paths:
        if config_path.exists():
            print(f"[配置] 从文件加载配置: {config_path.absolute()}")
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                return config
            except Exception as e:
                print(f"[警告] 加载配置文件失败: {e}")
                continue

    print("[配置] 使用脚本内置默认配置")
    print("[提示] 可创建 weekly_config.json 文件覆盖默认配置")
    return None


def apply_config(config):
    """应用外部配置到全局变量"""
    global MY_ID, WEEKLY_SUMMARY_CONTENT, OKR_WEEKLY_WORK, METRICS_CONTENT, DRY_RUN, AUTO_SUBMIT

    if not config:
        return

    if 'MY_ID' in config:
        MY_ID = config['MY_ID']
    if 'WEEKLY_SUMMARY_CONTENT' in config:
        WEEKLY_SUMMARY_CONTENT = config['WEEKLY_SUMMARY_CONTENT']
    if 'OKR_WEEKLY_WORK' in config:
        OKR_WEEKLY_WORK = config['OKR_WEEKLY_WORK']
    if 'METRICS_CONTENT' in config:
        METRICS_CONTENT = config['METRICS_CONTENT']
    if 'DRY_RUN' in config:
        DRY_RUN = config['DRY_RUN']
    if 'AUTO_SUBMIT' in config:
        AUTO_SUBMIT = config['AUTO_SUBMIT']


# ==================== 网络检测工具 ====================

def check_network_connection():
    """检测网络连接状态"""
    print("\n[网络检测] 检查网络连接...")

    try:
        ip = socket.gethostbyname("okr.baidu-int.com")
        print(f"  [OK] 内网域名解析正常: okr.baidu-int.com -> {ip}")
        return True
    except socket.gaierror:
        print(f"  [X] 内网域名解析失败")
        print("  [!] 请检查是否已连接公司VPN或内网")
        return False
    except Exception as e:
        print(f"  [X] 网络检测异常: {e}")
        return False


# ==================== 浏览器操作 ====================

async def find_scroll_container(page):
    """查找可滚动容器"""
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


async def scroll_container(page, selector, center):
    """滚动加载内容（30次滚动+5秒等待）"""
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


# ==================== 核心填写功能（V15改进版）====================

async def find_and_fill_input_v15(page, label_text, content, screenshot_on_fail=True):
    """
    V15改进版：更精确的定位和填写，带验证
    """
    if not content or not content.strip():
        print(f"  [跳过] 内容为空: {label_text}")
        return False

    content = content.strip()
    print(f"  正在填写: {label_text[:40]}...")

    try:
        # 策略1: 使用get_by_text查找标签（Playwright推荐方式）
        label_locator = page.get_by_text(label_text, exact=False).first

        if await label_locator.is_visible(timeout=3000):
            print(f"    [1] 找到标签: {label_text}")

            # 获取标签元素，向上查找父容器
            input_filled = False

            # 1a. 尝试直接点击标签，看是否能激活输入框
            try:
                await label_locator.click()
                print(f"    [2] 点击标签激活")
                await asyncio.sleep(0.5)

                # 查找当前聚焦的元素
                active_element = await page.evaluate("""() => {
                    const el = document.activeElement;
                    if (el && (el.tagName === 'TEXTAREA' || el.tagName === 'INPUT' || el.isContentEditable)) {
                        return {
                            tag: el.tagName,
                            editable: el.isContentEditable,
                            hasValue: el.value || el.innerText
                        };
                    }
                    return null;
                }""")

                if active_element:
                    print(f"    [3] 发现聚焦元素: {active_element}")
                    # 直接输入到聚焦元素
                    await page.keyboard.type(content[:100])  # 先输入一部分测试
                    await asyncio.sleep(0.3)
                    input_filled = True
            except Exception as e:
                print(f"    [!] 点击标签策略失败: {e}")

            # 1b. 查找相邻的输入框
            if not input_filled:
                try:
                    # 使用xpath查找相邻的textarea/input
                    input_locator = page.locator(f"xpath=//*[contains(text(), '{label_text}')]/following::textarea[1] | //*[contains(text(), '{label_text}')]/following::input[1] | //*[contains(text(), '{label_text}')]/following::*[@contenteditable='true'][1]").first

                    if await input_locator.is_visible(timeout=2000):
                        print(f"    [4] 找到相邻输入框")
                        await input_locator.fill(content)
                        input_filled = True
                except Exception as e:
                    print(f"    [!] 相邻输入框策略失败: {e}")

            # 1c. 使用JavaScript在父容器内查找
            if not input_filled:
                escaped_label = label_text.replace('"', '\\"')
                escaped_content = content.replace("\\", "\\\\").replace("`", "\\`").replace("$", "\\$")

                result = await page.evaluate(f"""() => {{
                    const labelText = "{escaped_label}";
                    const content = `{escaped_content}`;

                    // 找到标签元素
                    const allElements = Array.from(document.querySelectorAll('*'));
                    let labelEl = null;

                    for (let el of allElements) {{
                        if (el.textContent && el.textContent.trim().includes(labelText)) {{
                            labelEl = el;
                            break;
                        }}
                    }}

                    if (!labelEl) return 'label_not_found';

                    // 向上查找5层父元素
                    let parent = labelEl.parentElement;
                    for (let i = 0; i < 5 && parent; i++) {{
                        // 查找所有可能的输入元素
                        const inputs = parent.querySelectorAll('textarea, input[type="text"], [contenteditable="true"]');

                        for (let input of inputs) {{
                            // 检查元素是否可见
                            const rect = input.getBoundingClientRect();
                            if (rect.width > 0 && rect.height > 0) {{
                                // 聚焦并填写
                                input.focus();
                                input.scrollIntoView({{ behavior: 'smooth', block: 'center' }});

                                if (input.tagName === 'TEXTAREA' || input.tagName === 'INPUT') {{
                                    input.value = content;
                                }} else {{
                                    input.innerHTML = content.replace(/\\n/g, '<br>');
                                }}

                                // 触发事件
                                input.dispatchEvent(new Event('focus', {{ bubbles: true }}));
                                input.dispatchEvent(new Event('input', {{ bubbles: true }}));
                                input.dispatchEvent(new Event('change', {{ bubbles: true }}));
                                input.dispatchEvent(new KeyboardEvent('keydown', {{ key: 'a' }}));
                                input.dispatchEvent(new KeyboardEvent('keyup', {{ key: 'a' }}));

                                return 'filled_' + input.tagName.toLowerCase();
                            }}
                        }}
                        parent = parent.parentElement;
                    }}

                    return 'input_not_found';
                }}""")

                if result and not result.includes('not_found'):
                    print(f"    [5] JavaScript填写成功: {result}")
                    input_filled = True

            # 验证填写结果
            if input_filled:
                await asyncio.sleep(0.5)
                verify_result = await page.evaluate(f"""() => {{
                    const labelText = "{escaped_label}";
                    const allElements = Array.from(document.querySelectorAll('*'));

                    for (let el of allElements) {{
                        if (el.textContent && el.textContent.includes(labelText)) {{
                            let parent = el.parentElement;
                            for (let i = 0; i < 5 && parent; i++) {{
                                const input = parent.querySelector('textarea, input, [contenteditable="true"]');
                                if (input) {{
                                    const value = input.value || input.innerText || '';
                                    return value.length > 10 ? 'has_content' : 'empty';
                                }}
                                parent = parent.parentElement;
                            }}
                        }}
                    }}
                    return 'not_found';
                }}""")

                if verify_result == 'has_content':
                    print(f"    [OK] 验证通过，内容已填写")
                    return True
                else:
                    print(f"    [X] 验证失败，内容可能未保存")
                    if screenshot_on_fail:
                        await page.screenshot(path=f"fill_fail_{label_text.replace(' ', '_')[:20]}.png")
                    return False
        else:
            print(f"    [X] 标签不可见: {label_text}")
            return False

    except Exception as e:
        print(f"    [错误] {type(e).__name__}: {e}")
        return False


async def fill_weekly_summary_v15(page, content):
    """填写本周工作总结（V15改进版）"""
    print("\n[1/3] 填写本周工作总结...")

    # 尝试多种可能的标签文本
    labels = ["本周工作总结", "工作总结", "本周工作", "周报总结"]

    for label in labels:
        result = await find_and_fill_input_v15(page, label, content)
        if result:
            return True

    print("  [X] 所有策略均失败，无法填写本周工作总结")
    return False


async def fill_okr_weekly_work_v15(page, okr_work_dict):
    """填写OKR本周工作内容（V15改进版）"""
    print("\n[2/3] 填写OKR本周工作...")
    success_count = 0
    fail_count = 0

    for o_id, kr_dict in okr_work_dict.items():
        print(f"  处理 {o_id}...")
        for kr_id, work_content in kr_dict.items():
            # 尝试多种标签格式
            labels = [
                f"{o_id}{kr_id}",
                f"{o_id} {kr_id}",
                kr_id,
                f"{o_id}-{kr_id}",
                f"{o_id}/{kr_id}"
            ]

            filled = False
            for label in labels:
                if await find_and_fill_input_v15(page, label, work_content):
                    success_count += 1
                    filled = True
                    break

            if not filled:
                print(f"    [X] 无法定位 {o_id}/{kr_id}")
                fail_count += 1

    print(f"\n  统计: 成功{success_count}条, 失败{fail_count}条")
    return fail_count == 0


async def fill_other_metrics_v15(page, metrics_dict):
    """填写其他内容（V15改进版）"""
    print("\n[3/3] 填写其他内容...")
    for title, content in metrics_dict.items():
        await find_and_fill_input_v15(page, title, content)
        await asyncio.sleep(0.3)
    return True


async def click_submit_button(page):
    """
    点击保存/提交按钮
    """
    print("\n[提交] 查找并点击保存/提交按钮...")

    # 可能的按钮文本
    button_texts = ["保存", "提交", "更新", "确认", "Save", "Submit", "Update"]

    for text in button_texts:
        try:
            # 使用get_by_role查找按钮
            button = page.get_by_role("button", name=text).first
            if await button.is_visible(timeout=2000):
                print(f"  找到按钮: {text}")
                await button.click()
                print(f"  [OK] 已点击 '{text}' 按钮")
                await asyncio.sleep(2)  # 等待提交完成
                return True
        except:
            continue

    # 尝试使用JavaScript查找并点击
    result = await page.evaluate("""() => {
        const buttonTexts = ['保存', '提交', '更新', '确认', 'Save', 'Submit', 'Update'];
        const buttons = Array.from(document.querySelectorAll('button, input[type="submit"], a.btn, div.btn'));

        for (let btn of buttons) {
            const btnText = btn.textContent || btn.value || '';
            for (let text of buttonTexts) {
                if (btnText.includes(text) && btn.offsetParent !== null) {
                    btn.click();
                    return 'clicked: ' + btnText;
                }
            }
        }
        return 'not_found';
    }""")

    if result != 'not_found':
        print(f"  [OK] JavaScript点击: {result}")
        await asyncio.sleep(2)
        return True

    print("  [X] 未找到保存/提交按钮，请手动点击")
    return False


async def submit_single_weekly_v15(page, target_date=None, summary_content=None,
                                   okr_work=None, metrics=None, dry_run=True,
                                   auto_submit=False, is_first_page=False):
    """
    V15完整版：填写 + 验证 + 提交
    """
    target_date = target_date or get_default_date()
    summary_content = summary_content or WEEKLY_SUMMARY_CONTENT
    okr_work = okr_work or OKR_WEEKLY_WORK
    metrics = metrics or METRICS_CONTENT

    edit_url = f"{EDIT_BASE_URL}?date={target_date}&mandatorId={MY_ID}&isEdit=true"

    print("="*60)
    print(f"目标日期: {target_date}")
    print(f"编辑模式: {'预览' if dry_run else '实际填写'}")
    print(f"自动提交: {'是' if auto_submit else '否'}")
    print(f"URL: {edit_url}")
    print("="*60)

    print("\n正在加载编辑页面...")

    try:
        try:
            await page.goto(edit_url, wait_until='domcontentloaded', timeout=60000)
        except Exception as nav_error:
            current_url = page.url
            if 'uuap' in current_url:
                print("\n[!] 会话失效，需要重新登录")
                return {'status': 'login_required', 'url': current_url}
            print(f"\n[!] 导航警告: {nav_error}")

        if is_first_page:
            print("(首次页面等待5秒)...", end=" ", flush=True)
            await asyncio.sleep(5)
            print("OK")
        else:
            await asyncio.sleep(2)

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
    await scroll_container(page, selector, center)

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
        'part1_summary': await fill_weekly_summary_v15(page, summary_content),
        'part2_okr': await fill_okr_weekly_work_v15(page, okr_work),
        'part3_metrics': await fill_other_metrics_v15(page, metrics)
    }

    print("\n" + "="*60)
    print("填写完成统计：")
    print(f"  本周工作总结: {'OK' if results['part1_summary'] else 'FAIL'}")
    print(f"  OKR本周工作: {'OK' if results['part2_okr'] else 'FAIL'}")
    print(f"  其他内容: {'OK' if results['part3_metrics'] else 'FAIL'}")
    print("="*60)

    # 自动提交
    if auto_submit and any(results.values()):
        submit_result = await click_submit_button(page)
        results['submitted'] = submit_result

    return {'status': 'success', 'date': target_date, 'results': results}


async def submit_weekly_with_retry_v15(page, **kwargs):
    """带重试机制的周报提交"""
    result = await submit_single_weekly_v15(page, is_first_page=True, **kwargs)

    if result['status'] == 'error':
        print(f"\n{'='*60}")
        print("[重试机制] 首次尝试失败，准备重试...")
        print(f"{'='*60}")
        await asyncio.sleep(2)

        result = await submit_single_weekly_v15(page, is_first_page=False, **kwargs)

        if result['status'] == 'success':
            print("\n[OK] 重试成功")
        else:
            print("\n[X] 重试失败")

    return result


# ==================== 主流程 ====================

async def main():
    """主流程"""

    # 加载外部配置
    external_config = load_config_from_file()
    apply_config(external_config)

    # 网络检测
    if not check_network_connection():
        print("\n[!] 网络检查失败，请连接公司VPN或内网")
        response = input("\n是否仍要继续尝试? (y/N): ")
        if response.lower() != 'y':
            print("退出程序")
            return
        print("\n[!] 用户选择继续，尝试访问...")

    target_date = get_default_date()

    print("="*60)
    print("百度OKR周报自动提交 - V1.5 (完整功能版)")
    print("="*60)
    print("\n【V1.5改进】")
    print("  ✓ 更精确的定位策略（get_by_text）")
    print("  ✓ 填写后验证机制")
    print("  ✓ 自动点击保存/提交按钮")
    print("  ✓ 支持外部配置文件")
    print("  ✓ 调试截图功能")
    print(f"\n[配置]")
    print(f"  用户ID: {MY_ID}")
    print(f"  目标日期: {target_date}")
    print(f"  运行模式: {'预览' if DRY_RUN else '实际填写'}")
    print(f"  自动提交: {'是' if AUTO_SUBMIT else '否'}")
    print(f"\n[内容]")
    print(f"  - 工作总结: {len(WEEKLY_SUMMARY_CONTENT.strip())} 字符")
    print(f"  - OKR工作: {sum(len(v) for v in OKR_WEEKLY_WORK.values())} 条")
    print(f"  - 其他内容: {len(METRICS_CONTENT)} 项")

    if external_config:
        print(f"\n[配置来源] 外部配置文件")
    else:
        print(f"\n[配置来源] 脚本内置（可创建 weekly_config.json 覆盖）")

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

            if 'uuap' in page.url:
                print("\n[X] 未登录成功，退出")
                await browser.close()
                return
            print("[OK] 登录验证通过")

            # 执行填写
            result = await submit_weekly_with_retry_v15(
                page=page,
                target_date=target_date,
                summary_content=WEEKLY_SUMMARY_CONTENT,
                okr_work=OKR_WEEKLY_WORK,
                metrics=METRICS_CONTENT,
                dry_run=DRY_RUN,
                auto_submit=AUTO_SUBMIT
            )

            if result['status'] == 'login_required':
                print("\n[!] 会话失效或需要登录")
                print(f"[!] 当前页面: {result.get('url', '未知')}")
            elif result['status'] == 'error':
                print(f"\n[X] 错误: {result.get('error')}")
            elif result['status'] == 'preview':
                print("\n[OK] 预览完成")
                print("\n要实际提交，请修改配置：")
                print("  DRY_RUN = False      # 实际填写")
                print("  AUTO_SUBMIT = True   # 自动点击保存")
            else:
                print("\n[OK] 填写完成")
                if AUTO_SUBMIT:
                    if result.get('results', {}).get('submitted'):
                        print("[OK] 已自动提交")
                    else:
                        print("[!] 自动提交失败，请手动点击保存按钮")
                else:
                    print("[!] 请手动检查内容并点击保存按钮")
                    print("\n或设置 AUTO_SUBMIT = True 启用自动提交")

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