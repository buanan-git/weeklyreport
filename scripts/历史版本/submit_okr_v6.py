#!/usr/bin/env python3
"""
百度OKR周报自动提交脚本 - V1.6 (弹框处理+MD文件读取版)
"""

import asyncio
import socket
import json
import os
import re
from datetime import datetime, timedelta
from pathlib import Path
from playwright.async_api import async_playwright, Error as PlaywrightError

# ==================== 配置区域 ====================

MY_ID = "s673090"
EDIT_BASE_URL = "https://okr.baidu-int.com/pages/weekly.html#/home"

def get_default_date():
    today = datetime.now()
    last_monday = today - timedelta(days=today.weekday() + 7)
    return f"s{last_monday.strftime('%Y-%m-%d')}"

DRY_RUN = False
AUTO_SUBMIT = True

WEEKLY_SUMMARY_CONTENT = ""
OKR_WEEKLY_WORK = {}
METRICS_CONTENT = {}
MD_FILE_PATH = "team_weekly_report_v10.md"

# ==================== MD文件解析器 ====================

def parse_md_file(md_path):
    if not os.path.exists(md_path):
        print(f"[MD解析] 文件不存在: {md_path}")
        return None

    print(f"[MD解析] 读取文件: {os.path.abspath(md_path)}")

    try:
        with open(md_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except Exception as e:
        print(f"[MD解析] 读取失败: {e}")
        return None

    result = {'summary': '', 'okr_work': {}, 'metrics': {}}

    # 解析第一部分
    summary_match = re.search(r'## 第一部分：本周工作总结\s*\n\n(.*?)(?=\n## |\Z)', content, re.DOTALL)
    if summary_match:
        items = re.findall(r'^[-*]\s*(.+)$', summary_match.group(1), re.MULTILINE)
        result['summary'] = '\n'.join([f"{i+1}. {item}" for i, item in enumerate(items)])
        print(f"[MD解析] 本周工作总结: {len(items)} 条")

    # 解析第二部分 - OKR
    o_pattern = r'### (O\d+)\s+(.*?)\n\n(.*?)'
    for o_match in re.finditer(r'### (O\d+)\s+(.*?)\n\n(.*?)(?=\n### |\n## |\Z)', content, re.DOTALL):
        o_id, o_title, o_content = o_match.groups()
        result['okr_work'][o_id] = {}

        for kr_match in re.finditer(r'\*\*(KR\d+)\s+(.*?)\*\*\s*\n\n(.*?)(?=\*\*KR|\*\*O|### |\n## |\Z)', o_content + '\n\n**O', re.DOTALL):
            kr_id, kr_title, kr_content = kr_match.groups()
            if kr_id == 'O':
                break
            work_items = re.findall(r'^[-*]\s*(.+)$', kr_content, re.MULTILINE)
            if work_items:
                result['okr_work'][o_id][kr_id] = '\n'.join(work_items)
                print(f"[MD解析] {o_id}/{kr_id}: {len(work_items)} 条")

    # 解析第三部分
    for title in ['业务核心指标', '主要项目', '下周重点工作']:
        pattern = rf'### {title}\s*\n\n(.*?)(?=\n### |\n## |\Z)'
        match = re.search(pattern, content, re.DOTALL)
        if match:
            items = re.findall(r'^[-*]\s*(.+)$', match.group(1), re.MULTILINE)
            if items:
                result['metrics'][title] = '\n'.join(items)
                print(f"[MD解析] {title}: {len(items)} 条")

    return result

def load_content_from_md():
    global WEEKLY_SUMMARY_CONTENT, OKR_WEEKLY_WORK, METRICS_CONTENT

    script_dir = Path(__file__).parent.absolute()
    md_path = script_dir / MD_FILE_PATH

    parsed = parse_md_file(str(md_path))

    if parsed:
        if parsed['summary']:
            WEEKLY_SUMMARY_CONTENT = parsed['summary']
        if parsed['okr_work']:
            OKR_WEEKLY_WORK = parsed['okr_work']
        if parsed['metrics']:
            METRICS_CONTENT = parsed['metrics']
        print(f"[配置] 已从MD文件加载内容")
        return True
    else:
        print(f"[警告] 无法从MD文件加载，使用内置配置")
        return False

# ==================== 网络检测 ====================

def check_network_connection():
    print("\n[网络检测] 检查网络连接...")
    try:
        ip = socket.gethostbyname("okr.baidu-int.com")
        print(f"  [OK] 内网域名解析正常: {ip}")
        return True
    except:
        print(f"  [X] 内网域名解析失败，请检查VPN")
        return False

# ==================== 浏览器操作 ====================

async def find_scroll_container(page):
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

# ==================== 核心填写功能 ====================

async def find_and_fill_input_v16(page, label_text, content):
    if not content or not content.strip():
        print(f"  [跳过] 内容为空: {label_text}")
        return False

    content = content.strip()
    print(f"  正在填写: {label_text[:40]}...")

    try:
        label_locator = page.get_by_text(label_text, exact=False).first

        if await label_locator.is_visible(timeout=3000):
            print(f"    [1] 找到标签")

            try:
                await label_locator.click()
                print(f"    [2] 点击标签激活")
                await asyncio.sleep(0.5)
            except:
                pass

            input_filled = False

            try:
                input_locator = page.locator(
                    f"xpath=//*[contains(text(), '{label_text}')]/following::textarea[1] | "
                    f"//*[contains(text(), '{label_text}')]/following::input[1] | "
                    f"//*[contains(text(), '{label_text}')]/following::*[@contenteditable='true'][1]"
                ).first

                if await input_locator.is_visible(timeout=2000):
                    print(f"    [3] 找到输入框，开始填写...")
                    await input_locator.fill(content)
                    input_filled = True
            except Exception as e:
                print(f"    [!] 相邻输入框策略: {e}")

            if not input_filled:
                escaped_label = label_text.replace('"', '\\"')
                escaped_content = content.replace("\\", "\\\\").replace("`", "\\`").replace("$", "\\$")

                result = await page.evaluate(f"""() => {{
                    const labelText = "{escaped_label}";
                    const content = `{escaped_content}`;

                    const allElements = Array.from(document.querySelectorAll('*'));
                    let labelEl = null;

                    for (let el of allElements) {{
                        if (el.textContent && el.textContent.trim().includes(labelText)) {{
                            labelEl = el;
                            break;
                        }}
                    }}

                    if (!labelEl) return 'label_not_found';

                    let parent = labelEl.parentElement;
                    for (let i = 0; i < 5 && parent; i++) {{
                        const inputs = parent.querySelectorAll('textarea, input[type="text"], [contenteditable="true"]');

                        for (let input of inputs) {{
                            const rect = input.getBoundingClientRect();
                            if (rect.width > 0 && rect.height > 0) {{
                                input.focus();
                                input.scrollIntoView({{ behavior: 'smooth', block: 'center' }});

                                if (input.tagName === 'TEXTAREA' || input.tagName === 'INPUT') {{
                                    input.value = content;
                                }} else {{
                                    input.innerHTML = content.replace(/\n/g, '<br>');
                                }}

                                input.dispatchEvent(new Event('focus', {{ bubbles: true }}));
                                input.dispatchEvent(new Event('input', {{ bubbles: true }}));
                                input.dispatchEvent(new Event('change', {{ bubbles: true }}));

                                return 'filled_' + input.tagName.toLowerCase();
                            }}
                        }}
                        parent = parent.parentElement;
                    }}

                    return 'input_not_found';
                }}""")

                if result and not result.includes('not_found'):
                    print(f"    [4] JavaScript填写: {result}")
                    input_filled = True

            if input_filled:
                await asyncio.sleep(0.5)
                return True
        else:
            print(f"    [X] 标签不可见: {label_text}")
            return False

    except Exception as e:
        print(f"    [错误] {e}")
        return False

async def fill_weekly_summary_v16(page, content):
    print("\n[1/3] 填写本周工作总结...")
    labels = ["本周工作总结", "工作总结", "本周工作"]

    for label in labels:
        if await find_and_fill_input_v16(page, label, content):
            return True

    print("  [X] 无法填写本周工作总结")
    return False

async def fill_okr_weekly_work_v16(page, okr_work_dict):
    print("\n[2/3] 填写OKR本周工作...")
    success_count = 0
    fail_count = 0

    for o_id, kr_dict in okr_work_dict.items():
        print(f"  处理 {o_id}...")
        for kr_id, work_content in kr_dict.items():
            labels = [f"{o_id}{kr_id}", f"{o_id} {kr_id}", kr_id]

            filled = False
            for label in labels:
                if await find_and_fill_input_v16(page, label, work_content):
                    success_count += 1
                    filled = True
                    break

            if not filled:
                print(f"    [X] 无法定位 {o_id}/{kr_id}")
                fail_count += 1

    print(f"\n  统计: 成功{success_count}条, 失败{fail_count}条")
    return fail_count == 0

async def fill_other_metrics_v16(page, metrics_dict):
    print("\n[3/3] 填写其他内容...")
    for title, content in metrics_dict.items():
        await find_and_fill_input_v16(page, title, content)
        await asyncio.sleep(0.3)
    return True

# ==================== 提交和确认 ====================

async def click_submit_and_confirm(page):
    print("\n[提交] 查找并点击保存/提交按钮...")

    button_texts = ["保存", "提交", "更新", "确认", "Save", "Submit", "Update"]
    submit_clicked = False

    for text in button_texts:
        try:
            button = page.get_by_role("button", name=text).first
            if await button.is_visible(timeout=2000):
                print(f"  找到按钮: {text}")
                await button.click()
                print(f"  [OK] 已点击 '{text}' 按钮")
                submit_clicked = True
                break
        except:
            continue

    if not submit_clicked:
        result = await page.evaluate("""() => {
            const buttonTexts = ['保存', '提交', '更新', '确认', 'Save', 'Submit', 'Update'];
            const buttons = Array.from(document.querySelectorAll('button, input[type="submit"], a.btn'));

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
            submit_clicked = True

    if not submit_clicked:
        print("  [X] 未找到提交按钮")
        return False

    # 等待弹框
    print("\n[确认] 等待确认弹框...")
    await asyncio.sleep(1.5)

    # 处理确认弹框
    confirm_texts = ["确认", "确定", "是的", "Yes", "OK", "Confirm"]

    for text in confirm_texts:
        try:
            confirm_btn = page.get_by_role("button", name=text).first
            if await confirm_btn.is_visible(timeout=3000):
                print(f"  找到确认按钮: {text}")
                await confirm_btn.click()
                print(f"  [OK] 已点击确认按钮 '{text}'")
                await asyncio.sleep(2)
                return True
        except:
            continue

    # JavaScript点击确认
    confirm_result = await page.evaluate("""() => {
        const confirmTexts = ['确认', '确定', '是的', 'Yes', 'OK', 'Confirm'];

        const dialogs = document.querySelectorAll('[role="dialog"], .modal, .ant-modal, .el-dialog');
        for (let dialog of dialogs) {
            const buttons = dialog.querySelectorAll('button');
            for (let btn of buttons) {
                const text = btn.textContent || btn.value || '';
                for (let ct of confirmTexts) {
                    if (text.includes(ct) && btn.offsetParent !== null) {
                        btn.click();
                        return 'dialog_clicked: ' + text;
                    }
                }
            }
        }

        const allButtons = Array.from(document.querySelectorAll('button'));
        for (let btn of allButtons) {
            const text = btn.textContent || '';
            if ((text.includes('确认') || text.includes('确定')) && btn.offsetParent !== null) {
                btn.click();
                return 'clicked: ' + text;
            }
        }

        return 'not_found';
    }""")

    if confirm_result != 'not_found':
        print(f"  [OK] JavaScript确认: {confirm_result}")
        await asyncio.sleep(2)
        return True

    print("  [!] 未找到确认按钮，请手动点击")
    return False

# ==================== 主流程 ====================

async def submit_single_weekly_v16(page, target_date=None, dry_run=True, auto_submit=False, is_first_page=False):
    target_date = target_date or get_default_date()
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
            print("\n[!] 检测到登录页面")
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
        print(f"\n1. 本周工作总结:\n{WEEKLY_SUMMARY_CONTENT[:200]}...")
        print(f"\n2. OKR工作项数: {sum(len(v) for v in OKR_WEEKLY_WORK.values())}")
        print(f"\n3. 其他内容项数: {len(METRICS_CONTENT)}")
        print("\n" + "="*60)
        print("预览完成，未实际填写")
        return {'status': 'preview', 'date': target_date}

    print("\n开始填写内容...")
    results = {
        'part1_summary': await fill_weekly_summary_v16(page, WEEKLY_SUMMARY_CONTENT),
        'part2_okr': await fill_okr_weekly_work_v16(page, OKR_WEEKLY_WORK),
        'part3_metrics': await fill_other_metrics_v16(page, METRICS_CONTENT)
    }

    print("\n" + "="*60)
    print("填写完成统计：")
    print(f"  本周工作总结: {'OK' if results['part1_summary'] else 'FAIL'}")
    print(f"  OKR本周工作: {'OK' if results['part2_okr'] else 'FAIL'}")
    print(f"  其他内容: {'OK' if results['part3_metrics'] else 'FAIL'}")
    print("="*60)

    if auto_submit and any(results.values()):
        submit_result = await click_submit_and_confirm(page)
        results['submitted'] = submit_result

    return {'status': 'success', 'date': target_date, 'results': results}

async def submit_weekly_with_retry_v16(page, **kwargs):
    result = await submit_single_weekly_v16(page, is_first_page=True, **kwargs)

    if result['status'] == 'error':
        print(f"\n{'='*60}")
        print("[重试机制] 首次失败，准备重试...")
        print(f"{'='*60}")
        await asyncio.sleep(2)

        result = await submit_single_weekly_v16(page, is_first_page=False, **kwargs)

        if result['status'] == 'success':
            print("\n[OK] 重试成功")
        else:
            print("\n[X] 重试失败")

    return result

async def main():
    md_loaded = load_content_from_md()

    if not check_network_connection():
        print("\n[!] 网络检查失败")
        response = input("\n是否仍要继续? (y/N): ")
        if response.lower() != 'y':
            return

    target_date = get_default_date()

    print("="*60)
    print("百度OKR周报自动提交 - V1.6 (弹框处理+MD文件读取版)")
    print("="*60)
    print("\n【V1.6改进】")
    print("  ✓ 处理提交后的确认弹框")
    print("  ✓ 读取team_weekly_report_v10.md文件")
    print("  ✓ 智能解析MD结构匹配页面")
    print(f"\n[配置]")
    print(f"  用户ID: {MY_ID}")
    print(f"  目标日期: {target_date}")
    print(f"  编辑模式: {'预览' if DRY_RUN else '实际填写'}")
    print(f"  自动提交: {'是' if AUTO_SUBMIT else '否'}")
    print(f"\n[内容来源]")
    if md_loaded:
        print(f"  ✓ team_weekly_report_v10.md")
    else:
        print(f"  ✗ 内置默认配置")
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
                    print(f"\n[X] 网络错误")
                    await browser.close()
                    return
                else:
                    raise

            print("\n" + "="*60)
            print("[!] 请先登录OKR系统")
            print("="*60)
            input("\n登录完成后按 Enter 继续...")

            if 'uuap' in page.url:
                print("\n[X] 未登录成功")
                await browser.close()
                return
            print("[OK] 登录验证通过")

            result = await submit_weekly_with_retry_v16(
                page=page,
                target_date=target_date,
                dry_run=DRY_RUN,
                auto_submit=AUTO_SUBMIT
            )

            if result['status'] == 'login_required':
                print("\n[!] 需要登录")
            elif result['status'] == 'error':
                print(f"\n[X] 错误: {result.get('error')}")
            elif result['status'] == 'preview':
                print("\n[OK] 预览完成")
                print("\n要实际提交：")
                print("  DRY_RUN = False")
                print("  AUTO_SUBMIT = True")
            else:
                print("\n[OK] 填写完成")
                if AUTO_SUBMIT:
                    if result.get('results', {}).get('submitted'):
                        print("[OK] 已提交并确认")
                    else:
                        print("[!] 提交或确认失败")
                else:
                    print("[!] 请手动点击保存按钮")

            print("\n" + "="*60)
            input("\n按 Enter 关闭浏览器...")
            await browser.close()

    except Exception as e:
        print(f"\n[X] 严重错误: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    asyncio.run(main())