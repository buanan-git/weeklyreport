import asyncio
import json
import time
import os
from playwright.async_api import async_playwright

# ===================== 基础配置 =====================
USER_DATA_DIR = "./browser_profile"  # 持久化登录态目录
TARGET_URL = "https://okr.baidu-int.com/pages/weekly.html#/home?date=s2026-02-23&mandatorId=s&isEdit=true"
DATA_FILE_PATH = r"D:\个人\学习\OpenClaw\JobsOfAndy\BaiduOKR_WeeklyReport\scripts\team_weekly_report_v10.json"
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

def load_weekly_data(file_path):
    """加载周报填报数据（从指定JSON文件读取）"""
    iprint(f"读取填报数据文件: {file_path}")
    try:
        # 检查文件是否存在
        if not os.path.exists(file_path):
            iprint(f"[错误] 数据文件不存在: {file_path}")
            return None
        
        # 读取并解析JSON
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        # 验证数据结构
        required_keys = ["part1", "part2", "part3"]
        if not all(key in data for key in required_keys):
            iprint(f"[错误] 数据结构不完整，缺少关键字段: {required_keys}")
            return None
        
        iprint("✅ 数据文件读取成功")
        return data
    except json.JSONDecodeError as e:
        iprint(f"[错误] JSON解析失败: {e}")
        return None
    except Exception as e:
        iprint(f"[错误] 读取数据文件失败: {e}")
        return None

async def smart_activate_and_fill(locator, content):
    """智能激活并填充内容（适配富文本/普通输入框）"""
    try:
        # 先判断元素类型
        tag_name = await locator.evaluate("el => el.tagName.toLowerCase()")
        is_editable = await locator.evaluate("el => el.isContentEditable")
        
        if is_editable or tag_name in ["div", "p"]:
            # 适配富文本框（本周工作总结）
            await locator.evaluate(f"""
                (el, content) => {{
                    el.innerHTML = '';
                    const textNode = document.createTextNode(content);
                    el.appendChild(textNode);
                    el.dispatchEvent(new Event('input', {{ bubbles: true }}));
                    el.dispatchEvent(new Event('change', {{ bubbles: true }}));
                }}
            """, content)
            iprint("  [成功] 富文本内容已通过JS注入")
        else:
            # 适配普通输入框（OKR本周工作）
            await locator.fill(content)
            iprint("  [成功] 普通输入框内容已填充")
        return True
    except Exception as e:
        iprint(f"  [备用方案] 注入失败，尝试键盘输入: {e}")
        await locator.click(position={"x": 20, "y": 20})
        await locator.focus()
        await locator.press("Ctrl+A")
        await locator.press("Backspace")
        await locator.type(content, delay=50)
        return True

async def get_core_iframe(page):
    """精准定位核心业务iframe（基于截图结构）"""
    try:
        # 优先定位包含周报核心内容的iframe
        main_iframe = page.frame_locator("iframe#main, iframe[src*='weekly'], iframe[class*='weekly-container']")
        if await main_iframe.locator("body").count() == 0:
            main_iframe = page.frame_locator("iframe").first
        
        # 验证是否包含核心内容
        weekly_content = await main_iframe.locator("text=本周工作总结").count()
        if weekly_content > 0:
            iprint("✅ 定位到核心业务iframe")
            return main_iframe
        
        iprint("⚠️ 未找到核心iframe，使用页面根上下文")
        return page
    except Exception as e:
        iprint(f"[iframe定位] 错误: {e}")
        return page

# ===================== 核心填写函数（100%适配截图结构） =====================
async def fill_part1_work_summary(page, part1_data):
    """第一部分：本周工作总结 - 基于截图的精准定位（富文本编辑器）"""
    iprint("\n[填写] 第一部分：本周工作总结")
    if not part1_data or not part1_data.get("content"):
        iprint("  [跳过] 无有效内容")
        return False

    content = part1_data.get("content", "").strip()
    ctx = await get_core_iframe(page)
    
    # ===================== 核心定位（适配截图中的富文本编辑器） =====================
    # 方案1：通过编辑器工具栏定位（截图中最明显特征）
    target_locator = ctx.locator("div:has(.editor-toolbar)").locator("[contenteditable='true']").first
    
    # 方案2：通过提示文字兜底（截图中有“点击「智能生成」”提示）
    if await target_locator.count() == 0:
        target_locator = ctx.locator("div:has-text('点击「智能生成」')").locator("[contenteditable='true']").first

    try:
        await target_locator.wait_for(timeout=TIMEOUT_CONFIG["element"])
        iprint("✅ 精准定位到「本周工作总结」富文本编辑器")
    except Exception as e:
        iprint(f"❌ 定位失败: {e}")
        return False

    # 激活并填充内容
    await target_locator.scroll_into_view_if_needed()
    await target_locator.click(position={"x": 30, "y": 30})  # 点击编辑器中间位置
    await target_locator.focus()
    await smart_activate_and_fill(target_locator, content)
    
    # 验证填写结果
    actual_content = await target_locator.evaluate("el => el.innerText || el.textContent || ''")
    if content[:20] in actual_content:
        iprint("✅ 本周工作总结内容填写完成并验证通过")
        await ask_user("第一部分填写完成，请确认内容是否正确显示？")
        return True
    else:
        iprint(f"⚠️ 内容验证失败，实际值: {actual_content[:50]}...")
        return False

async def fill_part2_okr_weekly_work(page, part2_data):
    """第二部分：OKR进度更新 - 基于截图的精准定位（单行输入框）"""
    iprint("\n[填写] 第二部分：OKR 进度更新")
    if not part2_data or not part2_data.get("objectives"):
        iprint("  [跳过] 无有效OKR数据")
        return False

    objectives = part2_data.get("objectives", [])
    ctx = await get_core_iframe(page)
    success_count = 0

    for obj_idx, obj in enumerate(objectives):
        obj_title = obj.get("o_title", f"O{obj_idx+1}")
        iprint(f"\n  [处理] Objective: {obj_title}")
        
        key_results = obj.get("key_results", [])
        for kr_idx, kr in enumerate(key_results):
            kr_id = kr.get("kr_id", f"KR{kr_idx + 1}")
            kr_title = kr.get("kr_title", "")
            weekly_content = kr.get("weekly_work", "").strip()
            
            iprint(f"    [处理] KR: {kr_id} - {kr_title[:30]}...")
            
            # ===================== 核心定位（适配截图中的OKR输入框） =====================
            # 精准逻辑：先找KR编号 → 再找下方“本周工作”标签 → 最后定位兄弟输入框
            kr_locator = ctx.locator(f"""
                div:has-text("{kr_id}") 
                >> xpath=following::span[text()='本周工作']/following-sibling::input[1] | 
                   following::span[text()='本周工作']/following-sibling::textarea[1] |
                   following::span[text()='本周工作']/following-sibling::div[@contenteditable='true'][1]
            """).first

            try:
                await kr_locator.wait_for(timeout=10000)
                iprint(f"✅ 定位到KR {kr_id} 的本周工作输入框")
            except Exception as e:
                iprint(f"❌ KR {kr_id} 定位失败: {e}")
                continue

            # 激活并填充
            await kr_locator.scroll_into_view_if_needed()
            await kr_locator.click()
            await smart_activate_and_fill(kr_locator, weekly_content)
            
            # 验证
            if await kr_locator.evaluate("el => el.isContentEditable"):
                actual_content = await kr_locator.evaluate("el => el.innerText || ''")
            else:
                actual_content = await kr_locator.input_value()
            
            if weekly_content in actual_content:
                iprint(f"✅ KR {kr_id} 内容填写完成")
                success_count += 1
            else:
                iprint(f"⚠️ KR {kr_id} 内容验证失败")
    
    iprint(f"\n  [总结] OKR填写完成 - 成功 {success_count}/{sum(len(obj.get('key_results', [])) for obj in objectives)}")
    await ask_user("第二部分OKR填写完成，请确认所有KR内容是否正确显示？")
    return success_count > 0

async def fill_part3_metrics(page, part3_data):
    """第三部分：关键指标 - 适配截图结构"""
    iprint("\n[填写] 第三部分：关键指标")
    if not part3_data or not part3_data.get("sections"):
        iprint("  [跳过] 无有效指标数据")
        return False

    ctx = await get_core_iframe(page)
    sections = part3_data.get("sections", {})
    success_count = 0

    # 定位关键指标根容器（截图中“关键指标”标题下方的容器）
    metric_root = ctx.locator("div:has-text('关键指标')").locator("..").first
    try:
        await metric_root.wait_for(timeout=TIMEOUT_CONFIG["element"])
    except:
        iprint("⚠️ 未找到关键指标根容器，使用页面根上下文")
        metric_root = ctx

    for section_title, section_data in sections.items():
        iprint(f"\n  [处理] 指标板块: {section_title}")
        content = section_data.get("formatted_content", "").strip()
        if not content:
            iprint(f"    [跳过] 无内容")
            continue
        
        # 定位指标输入框（截图中指标标题后的输入框）
        metric_locator = metric_root.locator(f"""
            div:has-text("{section_title}") 
            >> xpath=following-sibling::input[1] | 
               following-sibling::textarea[1] |
               following-sibling::div[@contenteditable='true'][1]
        """).first

        try:
            await metric_locator.wait_for(timeout=10000)
        except:
            iprint(f"❌ {section_title}定位失败")
            continue

        # 填充内容
        await metric_locator.scroll_into_view_if_needed()
        if await smart_activate_and_fill(metric_locator, content):
            iprint(f"✅ {section_title} 填写完成")
            success_count += 1
        else:
            iprint(f"❌ {section_title} 填写失败")
    
    iprint(f"\n  [总结] 指标填写完成 - 成功 {success_count}/{len(sections)}")
    await ask_user("第三部分关键指标填写完成，请确认所有内容是否正确显示？")
    return success_count > 0

# ===================== 主流程函数 =====================
async def main():
    """脚本主入口（完整流程）"""
    iprint("===== 百度OKR周报自动填报脚本（截图适配最终版）=====")
    
    # 1. 加载填报数据
    weekly_data = load_weekly_data(DATA_FILE_PATH)
    if not weekly_data:
        iprint("[致命错误] 无法加载填报数据，脚本退出")
        return
    
    # 2. 初始化Playwright
    playwright = await async_playwright().start()
    browser = None
    try:
        # 启动带持久化登录的浏览器
        browser = await playwright.chromium.launch_persistent_context(
            user_data_dir=USER_DATA_DIR,
            headless=False,
            viewport={"width": 1920, "height": 1080},  # 适配大屏，避免元素被遮挡
            ignore_https_errors=True,
            args=[
                "--no-sandbox",
                "--disable-blink-features=AutomationControlled",
                "--disable-web-security",
                "--start-maximized"  # 浏览器最大化
            ]
        )
        
        # 创建页面并配置超时
        page = await browser.new_page()
        page.set_default_navigation_timeout(TIMEOUT_CONFIG["navigation"])
        page.set_default_timeout(TIMEOUT_CONFIG["element"])
        
        # 3. 访问目标页面
        iprint(f"访问目标页面: {TARGET_URL}")
        await page.goto(TARGET_URL, wait_until="networkidle")  # 等待网络空闲，确保页面加载完整
        
        # 等待动态内容渲染（适配内网慢加载）
        await asyncio.sleep(5)
        
        # 4. 手动操作提示
        iprint("\n===== 手动操作阶段 =====")
        iprint("请完成以下操作后按回车继续：")
        iprint("1. 登录系统（如果需要）")
        iprint("2. 展开所有需要填报的目录树结构")
        iprint("3. 确保所有输入框已显示在页面上")
        await ask_user("操作完成后请确认")
        
        # 5. 执行填报流程
        fill_results = []
        fill_results.append(await fill_part1_work_summary(page, weekly_data["part1"]))
        fill_results.append(await fill_part2_okr_weekly_work(page, weekly_data["part2"]))
        fill_results.append(await fill_part3_metrics(page, weekly_data["part3"]))
        
        # 6. 提交确认
        if all(fill_results):
            iprint("\n===== 填报完成 =====")
            if await ask_user("是否确认提交周报？"):
                ctx = await get_core_iframe(page)
                # 定位提交按钮（适配截图中的按钮样式）
                submit_btn = ctx.locator("""
                    button.erp-btn-primary:has-text('提交'), 
                    button[type='submit']:has-text('提交'),
                    div:has-text('提交') >> button
                """).first
                if await submit_btn.count() > 0:
                    await submit_btn.click()
                    iprint("✅ 提交按钮已点击！")
                else:
                    iprint("⚠️ 未找到提交按钮，请手动提交。")
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
        if browser:
            iprint("\n脚本已退出，浏览器保持打开...")

# ===================== 执行入口 =====================
if __name__ == "__main__":
    # 确保登录态目录存在
    if not os.path.exists(USER_DATA_DIR):
        os.makedirs(USER_DATA_DIR)
        iprint(f"创建登录态目录: {USER_DATA_DIR}")
    
    # 运行主流程
    asyncio.run(main())