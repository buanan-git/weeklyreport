#!/usr/bin/env python3
"""
百度OKR周报自动优化与提交脚本
功能：
1. 读取团队周报文件(team_weekly_report_v10.md)
2. 调用大模型API优化内容（保留核心信息，删除冗余）
3. 自动登录OKR系统并提交到本人(s673090)的周报页面

使用方法：
1. 确保已安装依赖: pip install openai playwright asyncio
2. 安装playwright浏览器: playwright install chromium  
3. 设置环境变量: export OPENAI_API_KEY="your-api-key"
4. 运行脚本: python auto_submit_weekly.py
"""

import os
import re
import json
import asyncio
from datetime import datetime
from openai import OpenAI
from playwright.async_api import async_playwright

# ==================== 配置 ====================
MY_ID = "s673090"
TEAM_WEEKLY_FILE = "team_weekly_report_v10.md"
BASE_URL = "https://okr.baidu-int.com/pages/weekly.html"

API_KEY = os.getenv("OPENAI_API_KEY", "")
API_BASE = os.getenv("OPENAI_API_BASE", "https://api.openai.com/v1")
MODEL_NAME = os.getenv("MODEL_NAME", "gpt-4")

# ==================== 第一步：读取并优化周报内容 ====================

def read_team_weekly(file_path):
    if not os.path.exists(file_path):
        print(f"✗ 未找到文件: {file_path}")
        import glob
        possible_files = glob.glob("*team_weekly*.md") + glob.glob("*周报*.md")
        if possible_files:
            print(f"  发现可能的文件: {possible_files}")
            return possible_files[0]
        return None

    with open(file_path, "r", encoding="utf-8") as f:
        content = f.read()
    print(f"✓ 已读取团队周报: {len(content)} 字符")
    return content

def optimize_weekly_content(team_content):
    if not API_KEY:
        print("✗ 未设置API密钥，请设置环境变量 OPENAI_API_KEY")
        return None

    try:
        client = OpenAI(api_key=API_KEY, base_url=API_BASE)
    except Exception as e:
        print(f"✗ 初始化API客户端失败: {e}")
        return None

    prompt = f"""请基于以下团队周报内容，为员工【s673090】生成一份个人周报。

要求：
1. 保留核心信息：关键项目进展、重要成果、数据指标
2. 删除冗余信息：重复描述、过于详细的日常事务、已过期内容  
3. 删除无效信息：空泛的陈述、与本周工作无关的内容
4. 优化格式：使用简洁清晰的语言，分点说明
5. 保持三部分结构：
   - 第一部分：本周工作总结（3-5条核心内容）
   - 第二部分：OKR进度（按O/KR结构，突出关键进展）
   - 第三部分：下周重点工作（2-3条）

团队周报内容：
{team_content}

请生成优化后的周报内容（使用Markdown格式）："""

    try:
        print(f"  正在调用大模型API ({MODEL_NAME}) 优化内容...")
        response = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[
                {"role": "system", "content": "你是一位专业的OKR周报优化助手，擅长提炼核心信息，删除冗余内容。"},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3,
            max_tokens=4000
        )

        optimized_content = response.choices[0].message.content
        print(f"✓ 内容优化完成: {len(optimized_content)} 字符")
        print(f"  优化前: {len(team_content)} 字符")
        print(f"  优化后: {len(optimized_content)} 字符")
        print(f"  压缩率: {(1 - len(optimized_content)/len(team_content))*100:.1f}%")
        return optimized_content

    except Exception as e:
        print(f"✗ API调用失败: {e}")
        return None

def parse_optimized_content(optimized_md):
    result = {"part1": "", "part2": "", "part3": ""}

    # 提取第一部分
    part1_match = re.search(r"#{1,2}\s*第一部分.*?(?=#{1,2}\s*第二部分|$)", optimized_md, re.DOTALL | re.IGNORECASE)
    if part1_match:
        result["part1"] = part1_match.group(0).strip()

    # 提取第二部分  
    part2_match = re.search(r"#{1,2}\s*第二部分.*?(?=#{1,2}\s*第三部分|$)", optimized_md, re.DOTALL | re.IGNORECASE)
    if part2_match:
        result["part2"] = part2_match.group(0).strip()

    # 提取第三部分
    part3_match = re.search(r"#{1,2}\s*第三部分.*?(?=$)", optimized_md, re.DOTALL | re.IGNORECASE)
    if part3_match:
        result["part3"] = part3_match.group(0).strip()

    # 清理标题行
    for key in result:
        if result[key]:
            lines = result[key].split("\n")
            if lines and ("第一部分" in lines[0] or "第二部分" in lines[0] or "第三部分" in lines[0]):
                result[key] = "\n".join(lines[1:]).strip()

    print(f"  解析结果:")
    print(f"    第一部分: {len(result['part1'])} 字符")
    print(f"    第二部分: {len(result['part2'])} 字符")  
    print(f"    第三部分: {len(result['part3'])} 字符")

    return result

# ==================== 第二步：自动提交周报 ====================

async def submit_weekly_report(content_parts):
    async with async_playwright() as p:
        print("\n=== 启动浏览器 ===")
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page(viewport={"width": 1920, "height": 1080})

        try:
            # 1. 访问OKR系统
            print("  访问OKR系统...")
            target_url = f"{BASE_URL}#/index?mode=steam&id={MY_ID}"
            await page.goto(target_url, wait_until="domcontentloaded")
            await asyncio.sleep(3)

            # 检查是否需要登录
            if "uuap" in page.url or "login" in page.url:
                print("  检测到登录页面，请手动登录...")
                print("  登录完成后，脚本将继续自动执行...")

                max_wait = 120
                waited = 0
                while "uuap" in page.url or "login" in page.url:
                    await asyncio.sleep(1)
                    waited += 1
                    if waited % 10 == 0:
                        print(f"  已等待 {waited} 秒...")
                    if waited > max_wait:
                        print("  ✗ 登录超时，退出")
                        return False

                print("  ✓ 登录完成")
                await asyncio.sleep(2)

            # 2. 等待页面加载
            print("  等待周报页面加载...")
            await page.wait_for_load_state("networkidle")
            await asyncio.sleep(2)

            # 3. 查找并点击"编辑"按钮
            print("  查找编辑按钮...")
            edit_selectors = [
                'button:has-text("编辑")',
                'button:has-text("填写")',
                'a:has-text("编辑")',
                'a:has-text("填写")',
                '[class*="edit"]',
                '[class*="write"]'
            ]

            edit_button = None
            for selector in edit_selectors:
                try:
                    edit_button = await page.wait_for_selector(selector, timeout=5000)
                    if edit_button:
                        print(f"    找到编辑按钮: {selector}")
                        break
                except:
                    continue

            if not edit_button:
                print("  ✗ 未找到编辑按钮")
                await page.screenshot(path="debug_page.png")
                return False

            await edit_button.click()
            print("  ✓ 已点击编辑按钮")
            await asyncio.sleep(3)

            # 4. 填写内容
            if content_parts["part1"]:
                print("  填写第一部分: 本周工作总结...")
                await fill_textarea(page, "本周工作总结", content_parts["part1"])

            if content_parts["part2"]:
                print("  填写第二部分: OKR进度...")
                await fill_okr_section(page, content_parts["part2"])

            if content_parts["part3"]:
                print("  填写第三部分: 下周重点工作...")
                await fill_textarea(page, "下周重点工作", content_parts["part3"])

            # 5. 提交
            print("  查找提交按钮...")
            submit_selectors = [
                'button:has-text("提交")',
                'button:has-text("更新")', 
                'button:has-text("保存")',
                'button[type="submit"]'
            ]

            submit_button = None
            for selector in submit_selectors:
                try:
                    submit_button = await page.wait_for_selector(selector, timeout=5000)
                    if submit_button:
                        print(f"    找到提交按钮: {selector}")
                        break
                except:
                    continue

            if submit_button:
                confirm = input("\n内容已填写完毕，是否确认提交? (y/n): ")
                if confirm.lower() == "y":
                    await submit_button.click()
                    print("  ✓ 已点击提交按钮")
                    await asyncio.sleep(3)
                    print("  ✓ 周报提交完成!")
                    return True
                else:
                    print("  已取消提交")
                    return False
            else:
                print("  ✗ 未找到提交按钮")
                return False

        except Exception as e:
            print(f"✗ 提交过程出错: {e}")
            await page.screenshot(path="error_screenshot.png")
            return False

        finally:
            input("\n按 Enter 键关闭浏览器...")
            await browser.close()
            print("\n=== 浏览器已关闭 ===")

async def fill_textarea(page, section_name, content):
    try:
        selectors = [
            f'textarea[placeholder*="{section_name}"]',
            'textarea',
            '[contenteditable="true"]'
        ]

        textarea = None
        for selector in selectors:
            try:
                elements = await page.query_selector_all(selector)
                for elem in elements:
                    is_visible = await elem.is_visible()
                    if is_visible:
                        textarea = elem
                        break
                if textarea:
                    break
            except:
                continue

        if textarea:
            await textarea.fill("")
            await textarea.fill(content)
            print(f"    ✓ 已填写: {section_name} ({len(content)} 字符)")
            await asyncio.sleep(1)
        else:
            print(f"    ✗ 未找到{section_name}的文本框")
    except Exception as e:
        print(f"    ✗ 填写{section_name}失败: {e}")

async def fill_okr_section(page, content):
    try:
        inputs = await page.query_selector_all('textarea, [contenteditable="true"]')
        if inputs:
            print(f"    找到 {len(inputs)} 个输入框")

            lines = [line.strip() for line in content.split("\n") if line.strip() and not line.startswith("#")]
            full_text = "\n".join(lines)

            for i, input_elem in enumerate(inputs[:3]):
                try:
                    is_visible = await input_elem.is_visible()
                    if is_visible:
                        await input_elem.fill("")
                        await input_elem.fill(full_text)
                        print(f"    ✓ 已填写OKR进度 ({len(full_text)} 字符)")
                        await asyncio.sleep(1)
                        break
                except:
                    continue
        else:
            print("    ✗ 未找到OKR输入框")
    except Exception as e:
        print(f"    ✗ 填写OKR进度失败: {e}")

# ==================== 主流程 ====================

async def main():
    print("="*60)
    print("百度OKR周报自动优化与提交工具")
    print("="*60)
    print(f"\n配置: 用户ID={MY_ID}, 模型={MODEL_NAME}")

    # 第一步：读取团队周报
    print("\n=== 第一步：读取团队周报 ===")
    team_content = read_team_weekly(TEAM_WEEKLY_FILE)
    if not team_content:
        print("✗ 无法读取团队周报，退出")
        return

    # 第二步：优化内容
    print("\n=== 第二步：优化周报内容 ===")
    optimized_content = optimize_weekly_content(team_content)
    if not optimized_content:
        print("⚠️ 内容优化失败，使用原始内容")
        optimized_content = team_content

    # 保存优化后的内容
    output_file = f'optimized_weekly_{datetime.now().strftime("%Y%m%d_%H%M%S")}.md'
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(optimized_content)
    print(f"✓ 已保存优化后的内容到: {output_file}")

    # 第三步：解析内容
    print("\n=== 第三步：解析内容结构 ===")
    content_parts = parse_optimized_content(optimized_content)

    # 显示预览
    print("\n=== 内容预览 ===")
    for part_name, part_content in content_parts.items():
        if part_content:
            preview = part_content[:150].replace("\n", " ")
            if len(part_content) > 150:
                preview += "..."
            print(f"\n{part_name}: {preview}")

    # 第四步：自动提交
    print("\n=== 第四步：自动提交周报 ===")
    confirm = input("  是否启动浏览器自动填写? (y/n): ")

    if confirm.lower() == "y":
        success = await submit_weekly_report(content_parts)
        if success:
            print("\n" + "="*60)
            print("✓ 周报提交流程完成!")
            print("="*60)
        else:
            print("\n" + "="*60)
            print("✗ 周报提交未完成")
            print(f"  优化后的内容已保存到: {output_file}")
            print("="*60)
    else:
        print("\n已取消自动提交")
        print(f"优化后的内容已保存到: {output_file}")

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n\n用户中断执行")
    except Exception as e:
        print(f"\n\n程序出错: {e}")
        import traceback
        traceback.print_exc()