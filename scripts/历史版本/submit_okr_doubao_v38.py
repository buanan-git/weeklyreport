import asyncio
import json
import time
from playwright.async_api import async_playwright

# ===================== 基础配置 =====================
USER_DATA_DIR = "./browser_profile"  # 持久化登录态目录
TARGET_URL = "https://okr.baidu-int.com/pages/weekly.html#/home?date=s2026-02-23&mandatorId=s&isEdit=true"
TIMEOUT_CONFIG = {
    "navigation": 60000,  # 页面导航超时
    "element": 30000,     # 元素定位超时
    "action": 5000        # 操作间隔超时
}

# ===================== 工具函数 =====================
def iprint(msg):
    """带时间戳的打印函数"""
    print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}")

async def ask_user(prompt: str) -> bool:
    """交互式确认（异步版）"""
    loop = asyncio.get_event_loop()
    response = await loop.run_in_executor(None, input, f"\n{prompt} (y/n): ")
    return response.strip().lower() == "y"

async def smart_activate_and_fill(locator, content):
    """智能激活并填充内容（备用方案）"""
    try:
        # 方案1：直接设置value/innerText
        await locator.evaluate(f"""
            (el, content) => {{
                if (el.tagName === 'TEXTAREA' || el.tagName === 'INPUT') {{
                    el.value = content;
                    el.dispatchEvent(new Event('input', {{ bubbles: true }}));
                    el.dispatchEvent(new Event('change', {{ bubbles: true }}));
                }} else if (el.isContentEditable) {{
                    el.innerText = content;
                    el.dispatchEvent(new Event('input', {{ bubbles: true }}));
                }}
            }}
        """, content)
        iprint("  [备用方案] 内容已通过JS注入")
        return True
    except Exception as e:
        iprint(f"  [备用方案] 注入失败: {e}")
        # 方案2：模拟键盘输入
        await locator.click()
        await locator.fill(content)
        return True

async def get_working_context(page):
    """递归穿透所有嵌套iframe，返回核心业务上下文"""
    try:
        # 第一层：main iframe（适配你的页面结构）
        main_iframe = page.frame_locator('iframe')
        if await main_iframe.locator('body').count() > 0:
            # 检查是否有第二层嵌套iframe
            nested_iframe = main_iframe.frame_locator('iframe')
            if await nested_iframe.locator('body').count() > 0:
                return nested_iframe
            return main_iframe
    except Exception as e:
        iprint(f"[iframe穿透] 警告: {e}")
    return page

async def ensure_editable_focus(page, locator, max_retries=3):
    """确保输入框获得真正的编辑焦点"""
    for retry in range(max_retries):
        try:
            await locator.scroll_into_view_if_needed()
            await locator.click(position={"x": 10, "y": 10})  # 点击元素内部10px位置
            await locator.focus()
            
            # 验证焦点是否成功
            is_focused = await locator.evaluate("el => el === document.activeElement")
            if is_focused:
                await page.keyboard.press("Escape")  # 清除默认选中
                await asyncio.sleep(0.2)
                return True
            
            iprint(f"[焦点验证] 重试 {retry+1}/{max_retries}...")
            await asyncio.sleep(0.5)
        except Exception as e:
            iprint(f"[焦点获取] 错误: {e}")
    return False

# ===================== 核心填写函数（优化版） =====================
async def fill_part1_work_summary(page, part1_data):
    """第一部分：本周工作总结 - 精准定位优化版"""
    iprint("\n[填写] 第一部分：本周工作总结")
    if not part1_data or not part1_data.get("content"):
        iprint("  [跳过] 无有效内容")
        return False

    # 处理内容（前后加换行，避免格式问题）
    content = "\n" + part1_data.get("content", "").strip() + "\n"
    ctx = await get_working_context(page)
    
    # 多维度定位器（优先级从高到低）
    locators = [
        ctx.locator("div:has-text('本周工作总结')").locator("textarea, [contenteditable='true']").first,
        ctx.get_by_placeholder("请输入本周工作总结").first,
        ctx.locator("[name*='summary'], [class*='summary']").locator("textarea").first,
        ctx.locator("//*[contains(text(), '本周工作总结')]/following-sibling::*[@contenteditable='true']").first
    ]
    
    target_locator = None
    for loc in locators:
        try:
            if await loc.count() > 0:
                target_locator = loc
                break
        except:
            continue
    
    if not target_locator:
        iprint("  [错误] 未找到工作总结输入区域")
        return False
    
    # 获取元素位置，确保可见
    box = await target_locator.bounding_box()
    if not box:
        iprint("  [错误] 无法获取元素位置")
        return False
    
    # 激活输入框（强化版）
    if not await ensure_editable_focus(page, target_locator):
        iprint("  [降级] 尝试坐标点击激活")
        click_x = box['x'] + box['width'] * 0.1
        click_y = box['y'] + box['height'] * 0.5
        await page.mouse.click(click_x, click_y, delay=100)
        await page.keyboard.press("Control+End")
    
    # 强制光标到末尾
    await target_locator.evaluate("el => {el.focus(); el.scrollTop = el.scrollHeight;}")
    await page.keyboard.press("Control+End")
    await asyncio.sleep(0.3)
    
    # 分批次输入内容（避免长文本丢失）
    try:
        iprint("  [注入] 开始输入工作总结内容...")
        chunk_size = 200
        content_chunks = [content[i:i+chunk_size] for i in range(0, len(content), chunk_size)]
        
        for idx, chunk in enumerate(content_chunks):
            await page.keyboard.insert_text(chunk)
            await asyncio.sleep(0.1)
            iprint(f"  [进度] 已输入 {min((idx+1)*chunk_size, len(content))}/{len(content)} 字符")
        
        # 验证输入是否成功
        actual_content = await target_locator.evaluate("el => el.innerText || el.value || ''")
        if content.strip() in actual_content:
            iprint("  [成功] 工作总结内容注入验证通过")
            await ask_user("第一部分填写完成，请确认内容是否正确显示？")
            return True
        else:
            iprint("  [警告] 内容验证失败，尝试备用方案...")
            await smart_activate_and_fill(target_locator, content)
            return True
    except Exception as e:
        iprint(f"  [错误] 输入失败: {e}")
        await smart_activate_and_fill(target_locator, content)
        return True

async def fill_part2_okr_weekly_work(page, part2_data):
    """第二部分：OKR进度更新 - 层级定位优化版"""
    iprint("\n[填写] 第二部分：OKR 进度更新")
    if not part2_data or not part2_data.get("objectives"):
        iprint("  [跳过] 无有效OKR数据")
        return False

    objectives = part2_data.get("objectives", [])
    ctx = await get_working_context(page)
    success_count = 0

    for obj_idx, obj in enumerate(objectives):
        obj_title = obj.get("o_title", f"O{obj_idx+1}")
        iprint(f"\n  [处理] Objective: {obj_title}")
        
        key_results = obj.get("key_results", [])
        for kr_idx, kr in enumerate(key_results):
            kr_id = kr.get("kr_id", f"KR{kr_idx + 1}")
            kr_title = kr.get("kr_title", "")
            weekly_content = "\n" + kr.get("weekly_work", "").strip() + "\n"
            
            iprint(f"    [处理] KR: {kr_id} - {kr_title[:30]}...")
            
            # 多维度定位KR锚点
            kr_anchor = None
            anchor_locators = [
                ctx.get_by_text(kr_id, exact=False),
                ctx.locator(f"//*[contains(text(), '{kr_id}')]"),
                ctx.locator(f"[data-id*='{kr_id}']"),
                ctx.locator(f"//*[contains(text(), '{kr_title[:20]}')]")
            ]
            
            for loc in anchor_locators:
                if await loc.count() > 0:
                    kr_anchor = loc
                    break
            
            if not kr_anchor:
                iprint(f"    [跳过] 未找到KR {kr_id} 定位锚点")
                continue
            
            # 基于KR锚点定位"本周工作"输入框（相对定位）
            work_locator = None
            work_locators = [
                # 方案1：KR锚点后跟随的"本周工作"输入框
                kr_anchor.locator("xpath=following::*[contains(text(), '本周工作')]/following-sibling::textarea[1]"),
                kr_anchor.locator("xpath=following::*[contains(text(), '本周工作')]/following-sibling::*[@contenteditable='true'][1]"),
                # 方案2：KR容器内的输入框
                ctx.locator(f"//*[contains(text(), '{kr_id}')]/ancestor::div[contains(@class, 'kr-item')]//textarea"),
                ctx.locator(f"//*[contains(text(), '{kr_id}')]/ancestor::div[contains(@class, 'kr-item')]//*[@contenteditable='true']")
            ]
            
            for loc in work_locators:
                if await loc.count() > 0:
                    work_locator = loc
                    break
            
            if not work_locator:
                # 备用方案：坐标定位
                iprint(f"    [降级] 使用坐标定位KR {kr_id} 输入区域")
                kr_box = await kr_anchor.bounding_box()
                if kr_box:
                    click_x = kr_box['x'] + kr_box['width'] + 20
                    click_y = kr_box['y'] + kr_box['height'] + 10
                    await page.mouse.click(click_x, click_y, delay=150)
                    work_locator = ctx.locator(":focus")  # 获取当前焦点元素
            
            if not work_locator or await work_locator.count() == 0:
                iprint(f"    [失败] 未找到KR {kr_id} 输入框")
                continue
            
            # 激活并填充
            if await ensure_editable_focus(page, work_locator):
                await page.keyboard.press("Control+A")  # 清空原有内容
                await page.keyboard.press("Backspace")
                await asyncio.sleep(0.2)
                
                # 注入内容
                await smart_activate_and_fill(work_locator, weekly_content)
                
                # 验证
                filled_content = await work_locator.evaluate("el => el.innerText || el.value || ''")
                if weekly_content.strip() in filled_content:
                    iprint(f"    [成功] KR {kr_id} 内容填写完成")
                    success_count += 1
                else:
                    iprint(f"    [警告] KR {kr_id} 内容验证失败")
            else:
                iprint(f"    [失败] KR {kr_id} 无法获取输入焦点")
    
    iprint(f"\n  [总结] OKR填写完成 - 成功 {success_count}/{sum(len(obj.get('key_results', [])) for obj in objectives)}")
    await ask_user("第二部分OKR填写完成，请确认所有KR内容是否正确显示？")
    return success_count > 0

async def fill_part3_metrics(page, part3_data):
    """第三部分：关键指标 - 板块定位优化版"""
    iprint("\n[填写] 第三部分：关键指标")
    if not part3_data or not part3_data.get("sections"):
        iprint("  [跳过] 无有效指标数据")
        return False

    ctx = await get_working_context(page)
    sections = part3_data.get("sections", {})
    success_count = 0

    # 定位指标区域根容器
    metric_root = None
    root_locators = [
        ctx.get_by_text("关键指标", exact=True),
        ctx.locator("[class*='metrics'], [class*='kpi']"),
        ctx.locator("div:has(h2:has-text('关键指标'))")
    ]
    
    for loc in root_locators:
        if await loc.count() > 0:
            metric_root = loc
            break
    
    if not metric_root:
        iprint("  [错误] 未找到关键指标根区域")
        return False

    # 遍历所有指标板块
    for section_title, section_data in sections.items():
        iprint(f"\n  [处理] 指标板块: {section_title}")
        content = section_data.get("formatted_content", "").strip()
        if not content:
            iprint(f"    [跳过] 无内容")
            continue
        
        # 定位当前板块的输入框
        section_input = None
        input_locators = [
            metric_root.locator(f"//*[contains(text(), '{section_title}')]/following::textarea[1]"),
            metric_root.locator(f"//*[contains(text(), '{section_title}')]/following::*[@contenteditable='true'][1]"),
            metric_root.locator(f"textarea[name*='{section_title[:4]}']"),
            metric_root.locator(f"[class*='section'][contains(text(), '{section_title[:6]}')] [contenteditable='true']")
        ]
        
        for loc in input_locators:
            if await loc.count() > 0:
                section_input = loc
                break
        
        if not section_input:
            # 备用方案：坐标偏移定位
            iprint(f"    [降级] 使用坐标定位{section_title}输入区域")
            root_box = await metric_root.bounding_box()
            if root_box:
                offset_y = 100 + (success_count * 150)  # 按板块顺序偏移
                click_x = root_box['x'] + 50
                click_y = root_box['y'] + offset_y
                await page.mouse.click(click_x, click_y, delay=100)
                section_input = ctx.locator(":focus")
        
        if not section_input or await section_input.count() == 0:
            iprint(f"    [失败] 未找到{section_title}输入框")
            continue
        
        # 激活并填充
        if await ensure_editable_focus(page, section_input):
            # 清空原有内容
            await page.keyboard.press("Control+A")
            await page.keyboard.press("Backspace")
            await asyncio.sleep(0.2)
            
            # 格式化内容并填充
            full_content = f"【{section_title}】\n{content}"
            await smart_activate_and_fill(section_input, full_content)
            
            # 验证
            actual = await section_input.evaluate("el => el.innerText || el.value || ''")
            if section_title in actual:
                iprint(f"    [成功] {section_title} 填写完成")
                success_count += 1
            else:
                iprint(f"    [警告] {section_title} 内容验证失败")
        else:
            iprint(f"    [失败] {section_title} 无法获取焦点")
    
    iprint(f"\n  [总结] 指标填写完成 - 成功 {success_count}/{len(sections)}")
    await ask_user("第三部分关键指标填写完成，请确认所有内容是否正确显示？")
    return success_count > 0

# ===================== 主流程函数 =====================
async def main():
    """脚本主入口"""
    iprint("===== 百度OKR周报自动填报脚本（优化版）=====")
    
    # 1. 初始化Playwright
    playwright = await async_playwright().start()
    browser = None
    try:
        # 启动带持久化登录的浏览器
        browser = await playwright.chromium.launch_persistent_context(
            user_data_dir=USER_DATA_DIR,
            headless=False,
            viewport={"width": 1280, "height": 800},
            ignore_https_errors=True,
            args=[
                "--no-sandbox",
                "--disable-blink-features=AutomationControlled",
                "--disable-web-security"
            ]
        )
        
        # 创建页面并配置超时
        page = await browser.new_page()
        page.set_default_navigation_timeout(TIMEOUT_CONFIG["navigation"])
        page.set_default_timeout(TIMEOUT_CONFIG["element"])
        
        # 2. 访问目标页面
        iprint(f"访问目标页面: {TARGET_URL}")
        await page.goto(TARGET_URL, wait_until="domcontentloaded")
        
        # 等待页面稳定
        await page.wait_for_load_state("domcontentloaded")
        await asyncio.sleep(3)
        
        # 3. 手动操作提示（登录/展开目录）
        iprint("\n===== 手动操作阶段 =====")
        iprint("请完成以下操作后按回车继续：")
        iprint("1. 登录系统（如果需要）")
        iprint("2. 展开所有需要填报的目录树结构")
        iprint("3. 确保所有输入框已显示在页面上")
        await ask_user("操作完成后请确认")
        
        # 4. 加载测试数据（你可替换为实际数据）
        iprint("\n加载填报数据...")
        test_data = {
            "part1": {
                "content": "本周完成XX项目需求评审2次，开发接口5个，联调通过；处理线上问题3个，完成技术文档编写1篇。"
            },
            "part2": {
                "objectives": [
                    {
                        "o_title": "O1: 完成XX核心功能开发",
                        "key_results": [
                            {
                                "kr_id": "KR1.1",
                                "kr_title": "完成XX接口开发",
                                "weekly_work": "本周完成接口编码，覆盖率95%，联调通过。"
                            },
                            {
                                "kr_id": "KR1.2",
                                "kr_title": "完成XX页面开发",
                                "weekly_work": "本周完成页面布局和交互，兼容主流浏览器。"
                            }
                        ]
                    }
                ]
            },
            "part3": {
                "sections": {
                    "核心指标1": {
                        "formatted_content": "接口响应时间：平均150ms（目标200ms）；Bug率：0.5%（目标1%）。"
                    },
                    "核心指标2": {
                        "formatted_content": "需求交付率：90%（目标85%）；代码评审通过率：100%。"
                    }
                }
            }
        }
        
        # 5. 执行填报流程
        fill_results = []
        fill_results.append(await fill_part1_work_summary(page, test_data["part1"]))
        fill_results.append(await fill_part2_okr_weekly_work(page, test_data["part2"]))
        fill_results.append(await fill_part3_metrics(page, test_data["part3"]))
        
        # 6. 提交确认（可选）
        if all(fill_results):
            iprint("\n===== 填报完成 =====")
            if await ask_user("是否确认提交周报？"):
                ctx = await get_working_context(page)
                submit_btn = ctx.locator("button:has-text('提交'), .erp-btn-primary").first
                if await submit_btn.count() > 0:
                    await submit_btn.click()
                    iprint("提交按钮已点击！")
                else:
                    iprint("未找到提交按钮，请手动提交。")
        else:
            iprint("\n===== 部分板块填写失败 =====")
            iprint("请检查失败板块，手动补充填写后提交。")
        
        # 7. 最终提示
        iprint("\n===== 脚本执行完成 =====")
        iprint("✅ 登录态已持久化保存（下次启动无需登录）")
        iprint("✅ 浏览器保持打开状态，可检查/修改内容")
        iprint("❌ 如需关闭浏览器，手动关闭窗口即可")
        
    except Exception as e:
        iprint(f"\n[致命错误] 脚本执行失败: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # 不关闭浏览器，仅保持运行
        if browser:
            iprint("\n脚本已退出，浏览器保持打开...")
            # 注释掉关闭逻辑，确保浏览器不关闭
            # await browser.close()
        # await playwright.stop()

# ===================== 执行入口 =====================
if __name__ == "__main__":
    # 确保登录态目录存在
    import os
    if not os.path.exists(USER_DATA_DIR):
        os.makedirs(USER_DATA_DIR)
        iprint(f"创建登录态目录: {USER_DATA_DIR}")
    
    # 运行主流程
    asyncio.run(main())