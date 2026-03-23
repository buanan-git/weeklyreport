#!/usr/bin/env python3
"""
OKR周报自动填报脚本 - v17
修复：基于静态页面结构精确定位，不依赖动态占位符
"""

import asyncio
import json
import os
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

from playwright.async_api import async_playwright, Page, Locator, TimeoutError as PlaywrightTimeout


class OKRSubmitter:
    def __init__(self):
        self.browser = None
        self.context = None
        self.page = None
        self.base_url = "https://okr.baidu-int.com"
        self.max_retries = 1

    async def init_browser(self, headless: bool = False):
        """初始化浏览器"""
        playwright = await async_playwright().start()
        self.browser = await playwright.chromium.launch(
            headless=headless,
            args=['--disable-blink-features=AutomationControlled']
        )
        self.context = await self.browser.new_context(
            viewport={'width': 1920, 'height': 1080},
            user_agent='Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        )
        self.page = await self.context.new_page()

        # 设置超时
        self.page.set_default_timeout(30000)
        self.page.set_default_navigation_timeout(60000)

    async def close_browser(self):
        """关闭浏览器"""
        if self.context:
            await self.context.close()
        if self.browser:
            await self.browser.close()

    async def check_login_required(self) -> bool:
        """检查是否需要登录"""
        current_url = self.page.url
        if "uuap.baidu.com" in current_url or "login" in current_url.lower():
            print("⚠️ 检测到需要登录，请手动完成登录...")
            print(f"   当前URL: {current_url}")
            return True
        return False

    async def wait_for_login(self, timeout: int = 300):
        """等待用户完成登录"""
        print("⏳ 等待登录完成，请在浏览器中完成登录操作...")
        start_time = asyncio.get_event_loop().time()

        while asyncio.get_event_loop().time() - start_time < timeout:
            current_url = self.page.url
            if "okr.baidu-int.com" in current_url and "uuap" not in current_url:
                print("✅ 登录成功！")
                await asyncio.sleep(2)
                return True
            await asyncio.sleep(1)

        raise TimeoutError("登录等待超时")

    async def navigate_with_retry(self, url: str):
        """带重试的页面导航"""
        for attempt in range(self.max_retries + 1):
            try:
                print(f"🌐 正在访问: {url} (尝试 {attempt + 1}/{self.max_retries + 1})")

                response = await self.page.goto(url, wait_until="domcontentloaded", timeout=60000)

                # 等待页面加载
                await asyncio.sleep(5)

                # 检查是否需要登录
                if await self.check_login_required():
                    await self.wait_for_login()
                    # 登录后重新导航
                    response = await self.page.goto(url, wait_until="domcontentloaded", timeout=60000)
                    await asyncio.sleep(3)

                if response and response.ok:
                    print(f"✅ 页面加载成功: {self.page.url}")
                    return True

            except Exception as e:
                print(f"⚠️ 导航失败: {e}")
                if attempt < self.max_retries:
                    print(f"🔄 {3}秒后重试...")
                    await asyncio.sleep(3)
                else:
                    raise

        return False

    def parse_report_file(self, file_path: str) -> Dict[str, Any]:
        """解析周报文件"""
        print(f"📄 正在读取周报文件: {file_path}")

        if not os.path.exists(file_path):
            raise FileNotFoundError(f"找不到文件: {file_path}")

        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()

        print(f"   文件大小: {len(content)} 字符")

        # 解析JSON
        try:
            data = json.loads(content)
            print(f"✅ JSON解析成功")
            return data
        except json.JSONDecodeError as e:
            print(f"⚠️ JSON解析失败，尝试按文本解析: {e}")
            return self._parse_text_content(content)

    def _parse_text_content(self, content: str) -> Dict[str, Any]:
        """解析文本格式的周报"""
        result = {
            "part1_work_summary": "",
            "part2_okr": {},
            "part3_metrics": ""
        }

        # 尝试提取各个部分
        lines = content.split('\n')
        current_section = None
        current_content = []

        for line in lines:
            line = line.strip()
            if not line:
                continue

            if "本周工作总结" in line or "工作摘要" in line:
                if current_section and current_content:
                    result[current_section] = '\n'.join(current_content)
                current_section = "part1_work_summary"
                current_content = []
            elif "OKR" in line or "关键结果" in line:
                if current_section and current_content:
                    result[current_section] = '\n'.join(current_content)
                current_section = "part2_okr"
                current_content = []
            elif "指标" in line or "数据" in line:
                if current_section and current_content:
                    result[current_section] = '\n'.join(current_content)
                current_section = "part3_metrics"
                current_content = []
            elif current_section:
                current_content.append(line)

        if current_section and current_content:
            result[current_section] = '\n'.join(current_content)

        return result

    def flatten_nested_content(self, data: Any, prefix: str = "") -> str:
        """递归展平嵌套内容"""
        lines = []

        if isinstance(data, dict):
            for key, value in data.items():
                if key.lower() in ['content', 'value', 'text', 'description', 'detail', 'details', 'work', 'works']:
                    # 直接提取内容字段
                    if isinstance(value, str):
                        lines.append(value)
                    else:
                        lines.append(self.flatten_nested_content(value))
                elif isinstance(value, (dict, list)):
                    nested = self.flatten_nested_content(value, prefix + "  ")
                    if nested:
                        lines.append(f"{prefix}{key}: {nested}")
                elif value:
                    lines.append(f"{prefix}{key}: {value}")

        elif isinstance(data, list):
            for item in data:
                if isinstance(item, str):
                    lines.append(item)
                else:
                    nested = self.flatten_nested_content(item, prefix + "  ")
                    if nested:
                        lines.append(nested)
        elif isinstance(data, str):
            lines.append(data)

        return '\n'.join(filter(None, lines))

    async def remove_overlay_elements(self):
        """移除可能遮挡的覆盖层"""
        try:
            await self.page.evaluate("""
                () => {
                    // 移除常见的覆盖层
                    const selectors = [
                        '.ant-modal-mask',
                        '.ant-modal-wrap',
                        '.ant-drawer-mask',
                        '[class*="overlay"]',
                        '[class*="backdrop"]',
                        '[class*="modal-mask"]',
                        '.ant-notification',
                        '.ant-message',
                        '.ant-popover',
                        '.ant-tooltip'
                    ];
                    selectors.forEach(selector => {
                        document.querySelectorAll(selector).forEach(el => {
                            el.style.display = 'none';
                            el.style.visibility = 'hidden';
                        });
                    });
                }
            """)
        except:
            pass

    async def smart_click_with_fallback(self, locator: Locator, timeout: int = 5000):
        """智能点击，带多种回退策略"""
        try:
            # 策略1: 标准点击
            await locator.click(timeout=timeout)
            return True
        except Exception as e1:
            print(f"   标准点击失败: {e1}")
            try:
                # 策略2: 强制点击
                await locator.click(force=True, timeout=timeout)
                return True
            except Exception as e2:
                print(f"   强制点击失败: {e2}")
                try:
                    # 策略3: JavaScript点击
                    element = await locator.element_handle()
                    if element:
                        await element.evaluate("el => el.click()")
                        return True
                except Exception as e3:
                    print(f"   JS点击失败: {e3}")
                    return False

    async def smart_fill(self, locator: Locator, text: str):
        """智能填充，确保内容正确填入"""
        try:
            # 先点击激活
            await locator.click()
            await asyncio.sleep(0.5)

            # 清除内容
            await self.page.keyboard.press("Control+a")
            await self.page.keyboard.press("Delete")
            await asyncio.sleep(0.3)

            # 填入内容
            await locator.fill(text)
            await asyncio.sleep(0.5)

            # 验证内容
            actual_value = await locator.input_value()
            if actual_value == text:
                print(f"   ✅ 内容验证成功")
                return True
            else:
                print(f"   ⚠️ 内容验证失败，尝试备用方法")
                # 备用：使用键盘输入
                await locator.click()
                await self.page.keyboard.press("Control+a")
                await self.page.keyboard.press("Delete")
                await self.page.keyboard.type(text)
                return True
        except Exception as e:
            print(f"   ❌ 填充失败: {e}")
            return False

    async def find_main_content_area(self) -> Optional[Locator]:
        """找到主内容区（排除侧边栏）"""
        # 策略1: 通过页面结构 - 左侧主要内容区
        selectors = [
            # 常见的布局类名
            'main.content-area',
            'main.main-content',
            'div.main-content',
            'div.content-area',
            'div[class*="main-content"]',
            'div[class*="content-area"]',
            '#main-content',
            '#content-area',
            # 通过位置 - 左侧区域
            'div[class*="layout"] > div:first-child',
            'div[class*="container"] > div:first-child',
        ]

        for selector in selectors:
            try:
                area = self.page.locator(selector).first
                if await area.count() > 0:
                    # 验证不是侧边栏
                    class_attr = await area.get_attribute('class') or ''
                    if 'sidebar' not in class_attr.lower() and 'side' not in class_attr.lower():
                        print(f"   找到主内容区: {selector}")
                        return area
            except:
                continue

        # 策略2: 通过页面标题位置推断
        try:
            # 找到周报标题，其所在区域即为主内容区
            title = self.page.locator('h1:has-text("周报"), h2:has-text("周报"), div:has-text("周报")').first
            if await title.count() > 0:
                # 向上找到容器
                for parent_level in range(1, 5):
                    parent = title.locator(f'xpath=ancestor::div[{parent_level}]').first
                    if await parent.count() > 0:
                        box = await parent.bounding_box()
                        if box and box['x'] < 500:  # 在页面左侧
                            print(f"   通过标题推断主内容区")
                            return parent
        except:
            pass

        return None

    async def fill_part1_work_summary(self, content: str):
        """填写第一部分：本周工作总结 - 基于静态结构定位"""
        print("\n📝 正在填写【本周工作总结】...")

        if not content or not content.strip():
            print("   ⚠️ 内容为空，跳过")
            return True

        try:
            # 找到主内容区
            main_area = await self.find_main_content_area()

            if main_area:
                # 在主内容区内找"本周工作总结"标题
                work_summary_heading = main_area.locator('h1, h2, h3, h4, div[role="heading"], .section-title').filter(has_text=re.compile(r"本周工作总结|工作摘要|工作总结"))

                if await work_summary_heading.count() > 0:
                    print("   ✅ 在主内容区找到'本周工作总结'标题")

                    # 找到该标题后的第一个编辑器
                    # 策略A: 直接兄弟元素
                    editor = work_summary_heading.locator('xpath=following-sibling::div[contains(@class,"editor") or @contenteditable="true"][1] | xpath=following-sibling::*//div[contains(@class,"editor") or @contenteditable="true"][1]').first

                    if await editor.count() == 0:
                        # 策略B: 父容器内的编辑器
                        parent = work_summary_heading.locator('xpath=ancestor::section | ancestor::div[contains(@class,"section")] | ancestor::div[contains(@class,"card")]').first
                        if await parent.count() > 0:
                            editor = parent.locator('div[contenteditable="true"], .editor, .rich-text-editor, [class*="editor"]').first

                    if await editor.count() > 0:
                        print("   ✅ 找到输入框，正在填入内容...")
                        await self.smart_fill(editor, content)
                        return True

        except Exception as e:
            print(f"   策略1失败: {e}")

        try:
            # 策略2: 通过页面顺序 - 第一个富文本编辑器通常是"本周工作总结"
            # 但要在主内容区内
            if main_area:
                first_editor = main_area.locator('div[contenteditable="true"], .editor, .rich-text-editor').first
                if await first_editor.count() > 0:
                    # 验证位置 - 应该在页面上半部分
                    box = await first_editor.bounding_box()
                    if box and box['y'] < 600:
                        print("   ✅ 通过顺序找到第一个输入框")
                        await self.smart_fill(first_editor, content)
                        return True

        except Exception as e:
            print(f"   策略2失败: {e}")

        try:
            # 策略3: 通过文本标签的层级结构
            # 找到包含"重要会议总结"或相关文本的区域
            section = self.page.locator('div:has-text("重要会议总结"):not(:has(div:has-text("辅助区"))) | div:has-text("工作总结"):not(:has(div:has-text("辅助区")))').first

            if await section.count() > 0:
                editor = section.locator('div[contenteditable="true"]').first
                if await editor.count() > 0:
                    print("   ✅ 通过区域文本找到输入框")
                    await self.smart_fill(editor, content)
                    return True

        except Exception as e:
            print(f"   策略3失败: {e}")

        print("   ❌ 无法找到【本周工作总结】的输入框")
        return False

    async def fill_okr_section(self, okr_data: Dict[str, Any]):
        """填写第二部分：OKR进度 - 基于KR标签定位"""
        print("\n📝 正在填写【OKR进度】...")

        if not okr_data:
            print("   ⚠️ OKR数据为空，跳过")
            return True

        success_count = 0
        main_area = await self.find_main_content_area()

        # 处理每个KR的本周工作
        for kr_key, kr_content in okr_data.items():
            if not kr_content:
                continue

            print(f"\n   正在处理 {kr_key}...")

            # 展平嵌套内容
            if isinstance(kr_content, (dict, list)):
                flat_content = self.flatten_nested_content(kr_content)
            else:
                flat_content = str(kr_content)

            if not flat_content.strip():
                continue

            try:
                # 策略1: 通过KR标签精确定位
                # 找到KR1/KR2文本，然后找它所在的卡片/行
                kr_locator = None
                if main_area:
                    kr_locator = main_area.locator(f'div:has-text("{kr_key}"):not(:has(div:has-text("辅助区")))').filter(has_text=re.compile(rf"^{kr_key}\b|\b{kr_key}\b")).first
                else:
                    kr_locator = self.page.locator(f'div:has-text("{kr_key}"):not(:has(div:has-text("辅助区")))').filter(has_text=re.compile(rf"^{kr_key}\b|\b{kr_key}\b")).first

                if await kr_locator.count() > 0:
                    print(f"      ✅ 找到{kr_key}标签")

                    # 向上找到KR容器（行或卡片）
                    kr_container = kr_locator.locator('xpath=ancestor::div[contains(@class,"kr") or contains(@class,"row") or contains(@class,"item") or contains(@class,"card")][1] | xpath=ancestor::tr[1]').first

                    if await kr_container.count() > 0:
                        # 在容器内找"本周工作"标签
                        work_label = kr_container.locator('div:has-text("本周工作"), span:has-text("本周工作"), label:has-text("本周工作")').first

                        if await work_label.count() > 0:
                            # 找到对应的输入框
                            editor = work_label.locator('xpath=following::div[contains(@class,"editor") or @contenteditable="true"][1] | xpath=following-sibling::div[contains(@class,"editor") or @contenteditable="true"][1] | xpath=ancestor::div[1]//div[contains(@class,"editor") or @contenteditable="true"]').first

                            if await editor.count() > 0:
                                print(f"      ✅ 找到输入框，正在填入内容...")
                                await self.smart_fill(editor, flat_content)
                                success_count += 1
                                continue

            except Exception as e:
                print(f"      策略1失败: {e}")

            try:
                # 策略2: 通过OKR进度区域内的所有编辑器顺序
                if main_area:
                    okr_section = main_area.locator('div:has-text("OKR进度"), section:has-text("OKR进度")').first
                    if await okr_section.count() > 0:
                        editors = okr_section.locator('div[contenteditable="true"], .editor')
                        total = await editors.count()

                        # KR1对应第1个编辑器（在OKR区域内），KR2对应第2个
                        kr_index = int(re.search(r"KR(\d+)", kr_key).group(1)) - 1 if re.search(r"KR(\d+)", kr_key) else 0

                        if total > kr_index:
                            target_editor = editors.nth(kr_index)
                            print(f"      ✅ 通过OKR区域内顺序找到输入框 (索引 {kr_index})")
                            await self.smart_fill(target_editor, flat_content)
                            success_count += 1
                            continue

            except Exception as e:
                print(f"      策略2失败: {e}")

            try:
                # 策略3: 通过页面全局编辑器顺序（排除第一个，那是工作总结）
                if main_area:
                    all_editors = main_area.locator('div[contenteditable="true"], .editor')
                    total = await all_editors.count()

                    # 第1个是工作总结，第2个是KR1，第3个是KR2，以此类推
                    kr_index = int(re.search(r"KR(\d+)", kr_key).group(1)) if re.search(r"KR(\d+)", kr_key) else 1
                    target_index = kr_index  # +1因为第0个是工作总结

                    if total > target_index:
                        target_editor = all_editors.nth(target_index)
                        print(f"      ✅ 通过全局顺序找到输入框 (索引 {target_index})")
                        await self.smart_fill(target_editor, flat_content)
                        success_count += 1
                        continue

            except Exception as e:
                print(f"      策略3失败: {e}")

            print(f"      ❌ 无法找到 {kr_key} 的输入框")

        print(f"\n   ✅ 成功填写 {success_count} 个KR")
        return success_count > 0

    async def fill_part3_metrics(self, content: str):
        """填写第三部分：指标数据"""
        print("\n📝 正在填写【指标数据】...")

        if not content or not content.strip():
            print("   ⚠️ 内容为空，跳过")
            return True

        main_area = await self.find_main_content_area()

        try:
            if main_area:
                metrics_section = main_area.locator('div:has-text("指标"), section:has-text("指标")').first
                if await metrics_section.count() > 0:
                    editor = metrics_section.locator('div[contenteditable="true"]').first
                    if await editor.count() > 0:
                        print("   ✅ 找到指标数据输入框")
                        await self.smart_fill(editor, content)
                        return True

        except Exception as e:
            print(f"   填写失败: {e}")

        return False

    async def click_submit_button(self):
        """点击提交按钮"""
        print("\n🚀 正在点击提交按钮...")

        try:
            # 查找提交/更新按钮 - 通常在页面右上角或底部
            submit_btn = self.page.locator('button:has-text("提交"), button:has-text("更新"), button:has-text("保存"), div[role="button"]:has-text("提交"), div[role="button"]:has-text("更新")').filter(has_text=re.compile(r"^(提交|更新|保存)$")).first

            if await submit_btn.count() > 0:
                await self.smart_click_with_fallback(submit_btn)
                print("   ✅ 已点击提交按钮")
                await asyncio.sleep(2)

                # 处理确认对话框
                await self.handle_confirmation_dialog()
                return True
            else:
                print("   ⚠️ 未找到提交按钮")
                return False

        except Exception as e:
            print(f"   ❌ 点击失败: {e}")
            return False

    async def handle_confirmation_dialog(self):
        """处理确认对话框"""
        print("\n🔄 检查确认对话框...")

        try:
            # 查找确认按钮
            confirm_btn = self.page.locator('button:has-text("确认"), button:has-text("确定"), button:has-text("是"), div[role="button"]:has-text("确认")').filter(has_text=re.compile(r"^(确认|确定|是)$")).first

            if await confirm_btn.count() > 0 and await confirm_btn.is_visible():
                print("   ✅ 发现确认对话框，点击确认")
                await confirm_btn.click()
                await asyncio.sleep(2)
                return True

        except Exception as e:
            print(f"   无需确认或处理失败: {e}")

        return False

    async def run(self, report_file: str, url: str = None):
        """运行填报流程"""
        try:
            # 1. 解析周报文件
            report_data = self.parse_report_file(report_file)

            # 2. 初始化浏览器
            await self.init_browser(headless=False)

            # 3. 导航到填报页面
            if url:
                await self.navigate_with_retry(url)
            else:
                # 默认URL
                default_url = f"{self.base_url}/pages/weekly.html"
                await self.navigate_with_retry(default_url)

            # 4. 移除覆盖层
            await self.remove_overlay_elements()

            # 5. 填写各部分
            await self.fill_part1_work_summary(report_data.get("part1_work_summary", ""))
            await self.fill_okr_section(report_data.get("part2_okr", {}))
            await self.fill_part3_metrics(report_data.get("part3_metrics", ""))

            # 6. 点击提交
            await self.click_submit_button()

            print("\n✅ 填报流程完成！")
            print("\n⏳ 等待10秒后关闭浏览器...")
            await asyncio.sleep(10)

        except Exception as e:
            print(f"\n❌ 运行出错: {e}")
            import traceback
            traceback.print_exc()

        finally:
            await self.close_browser()


def main():
    """主函数"""
    # 获取脚本所在目录
    script_dir = Path(__file__).parent.absolute()

    # 默认周报文件路径
    default_report_file = script_dir / "team_weekly_report_v10.md"

    # 检查命令行参数
    if len(sys.argv) > 1:
        report_file = sys.argv[1]
    else:
        report_file = str(default_report_file)

    if len(sys.argv) > 2:
        url = sys.argv[2]
    else:
        url = None

    # 检查文件是否存在
    if not os.path.exists(report_file):
        print(f"❌ 错误：找不到文件 {report_file}")
        print(f"   请确保文件存在于: {report_file}")
        print(f"\n用法: python {sys.argv[0]} [周报文件路径] [填报页面URL]")
        sys.exit(1)

    print("=" * 60)
    print("OKR周报自动填报工具 v17")
    print("=" * 60)
    print(f"📄 周报文件: {report_file}")
    print(f"🌐 目标URL: {url or '默认'}")
    print("=" * 60)

    # 运行填报
    submitter = OKRSubmitter()
    asyncio.run(submitter.run(report_file, url))


if __name__ == "__main__":
    main()