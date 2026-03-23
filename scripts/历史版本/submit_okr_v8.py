#!/usr/bin/env python3
"""
百度OKR周报自动提交脚本 - V2.1 (增强调试信息版)
"""

import asyncio
import socket
import json
import os
import re
import sys
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
JSON_FILE_PATH = "team_weekly_report_v10.json"

# 全局数据存储
WEEKLY_DATA = {
    "part1_work_summary": None,
    "part2_okr_structure": None,
    "part3_metrics": None
}

# ==================== JSON文件解析器（增强调试版） ====================

def parse_json_file(json_path):
    """解析JSON格式的周报文件 - 带详细调试"""
    print(f"\n{'='*60}")
    print("[DEBUG] 开始解析JSON文件")
    print(f"{'='*60}")

    # 1. 检查文件路径
    print(f"\n[DEBUG] 步骤1: 检查文件路径")
    print(f"  输入路径: {json_path}")
    print(f"  绝对路径: {os.path.abspath(json_path)}")
    print(f"  路径是否存在: {os.path.exists(json_path)}")

    if not os.path.exists(json_path):
        print(f"  [ERROR] 文件不存在: {json_path}")
        # 尝试列出当前目录文件
        try:
            current_dir = os.path.dirname(json_path) or '.'
            files = os.listdir(current_dir)
            print(f"  [DEBUG] 当前目录 '{current_dir}' 下的文件:")
            for f in files:
                print(f"    - {f}")
        except Exception as e:
            print(f"  [ERROR] 无法列出目录内容: {e}")
        return None

    # 2. 检查文件权限和大小
    print(f"\n[DEBUG] 步骤2: 检查文件属性")
    try:
        stat = os.stat(json_path)
        print(f"  文件大小: {stat.st_size} bytes")
        print(f"  最后修改: {datetime.fromtimestamp(stat.st_mtime)}")
        print(f"  可读权限: {os.access(json_path, os.R_OK)}")
    except Exception as e:
        print(f"  [ERROR] 无法获取文件属性: {e}")

    # 3. 读取文件内容
    print(f"\n[DEBUG] 步骤3: 读取文件内容")
    raw_content = None
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            raw_content = f.read()
        print(f"  成功读取, 字符数: {len(raw_content)}")
        print(f"  前200字符预览: {raw_content[:200]}")
    except Exception as e:
        print(f"  [ERROR] 读取失败: {e}")
        import traceback
        traceback.print_exc()
        return None

    # 4. 解析JSON
    print(f"\n[DEBUG] 步骤4: 解析JSON")
    data = None
    try:
        data = json.loads(raw_content)
        print(f"  [OK] JSON解析成功")
        print(f"  顶层键: {list(data.keys())}")
    except json.JSONDecodeError as e:
        print(f"  [ERROR] JSON解析失败: {e}")
        print(f"  错误位置: 第{e.lineno}行, 第{e.colno}列")
        # 显示错误附近的上下文
        lines = raw_content.split('\n')
        if e.lineno <= len(lines):
            print(f"  错误行内容: {lines[e.lineno-1]}")
        return None
    except Exception as e:
        print(f"  [ERROR] 解析异常: {e}")
        import traceback
        traceback.print_exc()
        return None

    # 5. 提取各部分内容
    print(f"\n[DEBUG] 步骤5: 提取各部分内容")
    result = {
        "part1_work_summary": None,
        "part2_okr_structure": None,
        "part3_metrics": None
    }

    # 解析第一部分
    print(f"\n  [DEBUG] 5.1 解析第一部分 (part1_work_summary)")
    if "part1_work_summary" in data:
        part1 = data["part1_work_summary"]
        print(f"    找到 part1_work_summary")
        print(f"    类型: {type(part1)}")

        if isinstance(part1, dict):
            content_list = part1.get("content", [])
            print(f"    content字段类型: {type(content_list)}")
            print(f"    content内容: {content_list}")

            if isinstance(content_list, list):
                formatted_content = '\n'.join([f"{i+1}. {item}" for i, item in enumerate(content_list)])
                result["part1_work_summary"] = {
                    "title": part1.get("title", "本周工作总结"),
                    "content": formatted_content,
                    "raw_content": content_list
                }
                print(f"    [OK] 第一部分解析成功: {len(content_list)} 条")
            else:
                print(f"    [WARN] content不是列表类型")
                result["part1_work_summary"] = {
                    "title": part1.get("title", "本周工作总结"),
                    "content": str(content_list),
                    "raw_content": [str(content_list)]
                }
        else:
            print(f"    [WARN] part1不是字典类型，尝试直接转换")
            result["part1_work_summary"] = {
                "title": "本周工作总结",
                "content": str(part1),
                "raw_content": [str(part1)]
            }
    else:
        print(f"    [WARN] 未找到 part1_work_summary 字段")

    # 解析第二部分
    print(f"\n  [DEBUG] 5.2 解析第二部分 (part2_okr_structure)")
    if "part2_okr_structure" in data:
        part2 = data["part2_okr_structure"]
        print(f"    找到 part2_okr_structure")
        print(f"    类型: {type(part2)}")

        if isinstance(part2, dict):
            okr_data = {
                "title": part2.get("title", "OKR进度"),
                "objectives": []
            }

            objectives = part2.get("objectives", [])
            print(f"    objectives数量: {len(objectives)}")

            for idx, obj in enumerate(objectives):
                print(f"    处理 Objective {idx+1}:")
                if not isinstance(obj, dict):
                    print(f"      [WARN] Objective不是字典类型: {type(obj)}")
                    continue

                o_data = {
                    "o_id": obj.get("o_id", f"O{idx+1}"),
                    "o_title": obj.get("o_title", ""),
                    "key_results": []
                }
                print(f"      O_ID: {o_data['o_id']}, 标题: {o_data['o_title']}")

                key_results = obj.get("key_results", [])
                print(f"      KR数量: {len(key_results)}")

                for kr_idx, kr in enumerate(key_results):
                    if not isinstance(kr, dict):
                        print(f"        [WARN] KR不是字典类型: {type(kr)}")
                        continue

                    kr_data = {
                        "kr_id": kr.get("kr_id", f"KR{kr_idx+1}"),
                        "kr_title": kr.get("kr_title", ""),
                        "weekly_work": kr.get("weekly_work", "")
                    }
                    print(f"        {kr_data['kr_id']}: {len(kr_data['weekly_work'])} 字符")
                    o_data["key_results"].append(kr_data)

                okr_data["objectives"].append(o_data)

            result["part2_okr_structure"] = okr_data
            total_krs = sum(len(o["key_results"]) for o in okr_data["objectives"])
            print(f"    [OK] 第二部分解析成功: {len(okr_data['objectives'])} 个O, {total_krs} 个KR")
        else:
            print(f"    [WARN] part2不是字典类型")
    else:
        print(f"    [WARN] 未找到 part2_okr_structure 字段")

    # 解析第三部分
    print(f"\n  [DEBUG] 5.3 解析第三部分 (part3_metrics)")
    if "part3_metrics" in data:
        part3 = data["part3_metrics"]
        print(f"    找到 part3_metrics")
        print(f"    类型: {type(part3)}")

        if isinstance(part3, dict):
            metrics_data = {
                "title": part3.get("title", "关键指标"),
                "sections": {}
            }

            sections = part3.get("sections", [])
            print(f"    sections数量: {len(sections)}")

            if isinstance(sections, list):
                for idx, section in enumerate(sections):
                    print(f"    处理 Section {idx+1}:")
                    if not isinstance(section, dict):
                        print(f"      [WARN] Section不是字典类型: {type(section)}")
                        continue

                    section_title = section.get("section_title", f"Section{idx+1}")
                    content_list = section.get("content", [])

                    if isinstance(content_list, list):
                        formatted_content = '\n'.join([f"{i+1}. {item}" for i, item in enumerate(content_list)])
                    else:
                        formatted_content = str(content_list)
                        content_list = [str(content_list)]

                    metrics_data["sections"][section_title] = {
                        "raw_content": content_list,
                        "formatted_content": formatted_content
                    }
                    print(f"      {section_title}: {len(content_list)} 条")
            else:
                print(f"    [WARN] sections不是列表类型")

            result["part3_metrics"] = metrics_data
            print(f"    [OK] 第三部分解析成功: {len(metrics_data['sections'])} 个section")
        else:
            print(f"    [WARN] part3不是字典类型")
    else:
        print(f"    [WARN] 未找到 part3_metrics 字段")

    # 6. 最终总结
    print(f"\n{'='*60}")
    print("[DEBUG] 解析结果总结")
    print(f"{'='*60}")
    print(f"  第一部分(工作总结): {'✓ 成功' if result['part1_work_summary'] else '✗ 失败/无数据'}")
    print(f"  第二部分(OKR进度): {'✓ 成功' if result['part2_okr_structure'] else '✗ 失败/无数据'}")
    print(f"  第三部分(关键指标): {'✓ 成功' if result['part3_metrics'] else '✗ 失败/无数据'}")

    return result

def load_content_from_json():
    """从JSON文件加载内容 - 带详细调试"""
    global WEEKLY_DATA

    print(f"\n{'='*60}")
    print("[DEBUG] 开始加载JSON文件")
    print(f"{'='*60}")

    # 1. 确定脚本目录
    print(f"\n[DEBUG] 步骤1: 确定脚本目录")
    try:
        script_path = Path(__file__).resolve()
        script_dir = script_path.parent.absolute()
        print(f"  脚本路径: {script_path}")
        print(f"  脚本目录: {script_dir}")
    except Exception as e:
        print(f"  [ERROR] 无法获取脚本路径: {e}")
        script_dir = Path(os.getcwd())
        print(f"  [FALLBACK] 使用当前工作目录: {script_dir}")

    # 2. 构建JSON文件路径
    print(f"\n[DEBUG] 步骤2: 构建文件路径")
    json_path = script_dir / JSON_FILE_PATH
    print(f"  JSON文件名: {JSON_FILE_PATH}")
    print(f"  完整路径: {json_path}")
    print(f"  路径是否存在: {json_path.exists()}")

    # 3. 解析文件
    print(f"\n[DEBUG] 步骤3: 调用解析函数")
    parsed = parse_json_file(str(json_path))

    # 4. 应用结果
    print(f"\n[DEBUG] 步骤4: 应用解析结果")
    if parsed:
        WEEKLY_DATA = parsed
        print(f"  [OK] 数据已加载到全局变量 WEEKLY_DATA")

        # 详细显示加载的内容
        if WEEKLY_DATA["part1_work_summary"]:
            content = WEEKLY_DATA["part1_work_summary"]["content"]
            print(f"  第一部分内容长度: {len(content)} 字符")
            print(f"  第一部分预览: {content[:100]}...")

        if WEEKLY_DATA["part2_okr_structure"]:
            obj_count = len(WEEKLY_DATA["part2_okr_structure"]["objectives"])
            print(f"  第二部分Objectives数量: {obj_count}")

        if WEEKLY_DATA["part3_metrics"]:
            section_count = len(WEEKLY_DATA["part3_metrics"]["sections"])
            print(f"  第三部分Sections数量: {section_count}")

        return True
    else:
        print(f"  [ERROR] 解析失败，使用空数据")
        WEEKLY_DATA = {
            "part1_work_summary": None,
            "part2_okr_structure": None,
            "part3_metrics": None
        }
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

# ==================== 页面结构分析和操作 ====================

async def analyze_page_structure(page):
    """分析页面现有结构，识别已存在的部分"""
    print("\n[页面分析] 分析页面结构...")

    structure = await page.evaluate("""() => {
        const result = {
            has_part1: false,
            has_part2: false,
            has_part3: false,
            sections: []
        };

        const headers = document.querySelectorAll('h1, h2, h3, h4, .ant-typography, .section-title, [class*="title"], [class*="header"]');

        headers.forEach(header => {
            const text = header.textContent || '';

            if (text.includes('本周工作总结') || text.includes('工作总结') || text.includes('本周工作')) {
                result.has_part1 = true;
                result.part1_element = text;
            }
            if (text.includes('OKR') || text.includes('okr') || text.includes('目标')) {
                result.has_part2 = true;
                result.part2_element = text;
            }
            if (text.includes('指标') || text.includes('metrics') || text.includes('数据') || text.includes('项目')) {
                result.has_part3 = true;
                result.part3_element = text;
            }

            result.sections.push({
                text: text.substring(0, 50),
                tag: header.tagName
            });
        });

        return result;
    }""")

    print(f"  第一部分(工作总结): {'存在' if structure['has_part1'] else '不存在'}")
    print(f"  第二部分(OKR进度): {'存在' if structure['has_part2'] else '不存在'}")
    print(f"  第三部分(关键指标): {'存在' if structure['has_part3'] else '不存在'}")

    return structure

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
    """滚动加载内容"""
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

async def smart_fill_input(page, label_text, content, create_if_missing=False, section_type=None):
    """智能填写输入框"""
    if not content or not content.strip():
        print(f"  [跳过] 内容为空: {label_text}")
        return False

    content = content.strip()
    print(f"\n  正在处理: {label_text[:40]}...")

    # 策略1: 尝试通过标签文本查找输入框
    try:
        label_variants = [label_text]
        if "本周工作总结" in label_text:
            label_variants.extend(["工作总结", "本周工作", "summary", "work summary"])
        elif "OKR" in label_text or "O" in label_text:
            label_variants.extend(["OKR进度", "OKR", "目标与关键结果"])

        for variant in label_variants:
            try:
                label_locator = page.get_by_text(variant, exact=False).first
                if await label_locator.is_visible(timeout=2000):
                    print(f"    [OK] 找到标签: {variant}")

                    input_filled = await find_and_fill_nearby_input(page, label_locator, content)
                    if input_filled:
                        return True
            except:
                continue

    except Exception as e:
        print(f"    [!] 标签查找失败: {e}")

    # 策略2: 使用JavaScript深度查找
    print("    尝试JavaScript深度查找...")
    js_result = await fill_by_javascript(page, label_text, content)
    if js_result:
        return True

    # 策略3: 如果需要，创建新区域
    if create_if_missing and section_type:
        print(f"    [!] 未找到{label_text}，尝试创建...")
        created = await create_new_section(page, section_type, label_text, content)
        if created:
            return True

    print(f"    [X] 无法填写: {label_text}")
    return False

async def find_and_fill_nearby_input(page, label_locator, content):
    """查找标签附近的输入框并填写"""
    try:
        await label_locator.click()
        await asyncio.sleep(0.5)

        label_handle = await label_locator.element_handle()
        if not label_handle:
            return False

        input_selectors = [
            'textarea',
            'input[type="text"]',
            '[contenteditable="true"]',
            '.ant-input',
            '.rich-text-editor',
            '[role="textbox"]'
        ]

        for selector in input_selectors:
            try:
                input_locator = page.locator(
                    f'xpath=//*[contains(text(), "{await label_handle.inner_text()}")]/following::{selector}[1] | '
                    f'xpath=//*[contains(text(), "{await label_handle.inner_text()}")]/ancestor::div[1]//{selector}'
                ).first

                if await input_locator.is_visible(timeout=2000):
                    print(f"    [OK] 找到输入框: {selector}")
                    await input_locator.fill(content)
                    await asyncio.sleep(0.5)
                    await input_locator.press('End')
                    await asyncio.sleep(0.3)

                    return True
            except:
                continue

    except Exception as e:
        print(f"    [!] 附近输入框查找失败: {e}")

    return False

async def fill_by_javascript(page, label_text, content):
    """使用JavaScript深度查找并填写"""
    try:
        escaped_label = label_text.replace('"', '\\"').replace("'", "\\'")
        escaped_content = content.replace("\\", "\\\\").replace("`", "\\`").replace("$", "\\$").replace('"', '\\"')

        result = await page.evaluate(f"""(labelText, content) => {{
            const findElementByText = (text) => {{
                const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
                let node;
                while (node = walker.nextNode()) {{
                    if (node.textContent.trim().includes(text)) {{
                        return node.parentElement;
                    }}
                }}
                return null;
            }};

            const findInputNearElement = (element) => {{
                if (element.matches('textarea, input[type="text"], [contenteditable="true"]')) {{
                    return element;
                }}

                let inputs = element.querySelectorAll('textarea, input[type="text"], [contenteditable="true"]');
                if (inputs.length > 0) return inputs[0];

                let parent = element.parentElement;
                for (let i = 0; i < 5 && parent; i++) {{
                    inputs = parent.querySelectorAll('textarea, input[type="text"], [contenteditable="true"]');
                    if (inputs.length > 0) return inputs[0];
                    parent = parent.parentElement;
                }}

                return null;
            }};

            const fillInput = (input, text) => {{
                input.scrollIntoView({{ behavior: 'smooth', block: 'center' }});
                input.focus();

                if (input.tagName === 'TEXTAREA' || input.tagName === 'INPUT') {{
                    input.value = text;
                }} else if (input.contentEditable === 'true') {{
                    input.innerHTML = text.replace(/\\n/g, '<br>');
                }}

                ['focus', 'input', 'change', 'blur'].forEach(eventType => {{
                    const event = new Event(eventType, {{ bubbles: true }});
                    input.dispatchEvent(event);
                }});

                return true;
            }};

            const labelEl = findElementByText(labelText);
            if (!labelEl) return 'label_not_found';

            const input = findInputNearElement(labelEl);
            if (!input) return 'input_not_found';

            fillInput(input, content);
            return 'filled_success';
        }}""", label_text, content)

        if result == 'filled_success':
            print(f"    [OK] JavaScript填写成功")
            await asyncio.sleep(0.5)
            return True
        else:
            print(f"    [!] JavaScript结果: {result}")

    except Exception as e:
        print(f"    [!] JavaScript填写失败: {e}")

    return False

async def create_new_section(page, section_type, title, content):
    """在页面上创建新的部分"""
    print(f"    [创建] 尝试创建新区域: {title}")

    try:
        add_buttons = await page.evaluate("""(sectionType) => {
            const buttons = Array.from(document.querySelectorAll('button, a, [role="button"]'));
            const keywords = {
                'part1': ['添加', '新建', '增加', '工作总结', '添加总结'],
                'part2': ['添加OKR', '新建OKR', '添加目标', '添加O'],
                'part3': ['添加指标', '新建指标', '添加项目']
            };

            const targets = keywords[sectionType] || ['添加', '新建', '+'];

            return buttons
                .filter(btn => {
                    const text = btn.textContent || btn.title || '';
                    return targets.some(kw => text.includes(kw)) && btn.offsetParent !== null;
                })
                .map(btn => ({{
                    text: btn.textContent.substring(0, 30),
                    class: btn.className,
                    tag: btn.tagName
                }}));
        }""", section_type)

        if add_buttons.length > 0:
            print(f"    找到 {add_buttons.length} 个可能的添加按钮")
            for btn_info in add_buttons[:3]:
                try:
                    btn = page.locator(f'button:has-text("{btn_info.text}")').first
                    if await btn.is_visible(timeout=2000):
                        await btn.click()
                        print(f"    [OK] 点击添加按钮: {btn_info.text}")
                        await asyncio.sleep(2)

                        return await fill_by_javascript(page, title, content)
                except:
                    continue

        print("    [!] 未找到添加按钮，尝试直接插入...")

        await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        await asyncio.sleep(1)

        result = await page.evaluate("""(title, content) => {
            const container = document.querySelector('.ant-form, .weekly-form, [class*="form"], main, article') || document.body;

            const newSection = document.createElement('div');
            newSection.style.marginTop = '20px';
            newSection.style.padding = '16px';
            newSection.style.border = '1px solid #d9d9d9';
            newSection.style.borderRadius = '4px';

            const titleEl = document.createElement('h3');
            titleEl.textContent = title;
            titleEl.style.marginBottom = '12px';

            const textarea = document.createElement('textarea');
            textarea.value = content;
            textarea.style.width = '100%';
            textarea.style.minHeight = '120px';
            textarea.style.padding = '8px';
            textarea.style.border = '1px solid #d9d9d9';
            textarea.style.borderRadius = '4px';

            newSection.appendChild(titleEl);
            newSection.appendChild(textarea);
            container.appendChild(newSection);

            textarea.focus();
            textarea.dispatchEvent(new Event('input', { bubbles: true }));

            return 'created';
        }""", title, content)

        if result == 'created':
            print(f"    [OK] 已创建新区域并填写")
            return True

    except Exception as e:
        print(f"    [X] 创建区域失败: {e}")

    return False

# ==================== OKR特定处理 ====================

async def fill_okr_section(page, okr_data):
    """专门处理OKR部分的填写"""
    print(f"\n[OKR填写] 开始填写OKR进度...")

    objectives = okr_data.get("objectives", [])
    success_count = 0
    fail_count = 0

    for obj in objectives:
        o_id = obj.get("o_id", "")
        o_title = obj.get("o_title", "")
        print(f"\n  处理目标: {o_id} - {o_title}")

        for kr in obj.get("key_results", []):
            kr_id = kr.get("kr_id", "")
            kr_title = kr.get("kr_title", "")
            weekly_work = kr.get("weekly_work", "")

            print(f"    处理KR: {kr_id} - {kr_title}")

            label_variants = [
                f"{o_id}{kr_id}",
                f"{o_id} {kr_id}",
                kr_id,
                f"{kr_id} {kr_title}",
                kr_title,
                f"{o_id} {kr_title}",
            ]

            filled = False

            for variant in label_variants:
                try:
                    kr_locator = page.get_by_text(variant, exact=False).first
                    if await kr_locator.is_visible(timeout=2000):
                        print(f"      [OK] 找到KR标签: {variant}")

                        kr_handle = await kr_locator.element_handle()
                        if kr_handle:
                            filled = await fill_weekly_work_in_kr_context(page, kr_handle, weekly_work)
                            if filled:
                                success_count += 1
                                break
                except:
                    continue

            if not filled:
                print(f"      [!] 尝试全局搜索填写...")
                filled = await fill_by_javascript(page, f"{o_id}{kr_id} 本周工作", weekly_work)
                if filled:
                    success_count += 1
                else:
                    fail_count += 1
                    print(f"      [X] 无法填写 {o_id}/{kr_id}")

            await asyncio.sleep(0.5)

    print(f"\n  OKR填写统计: 成功{success_count}条, 失败{fail_count}条")
    return fail_count == 0

async def fill_weekly_work_in_kr_context(page, kr_element, content):
    """在KR元素的上下文中查找"本周工作"输入框并填写"""
    try:
        result = await page.evaluate("""(krElement, content) => {
            const findWeeklyWorkInput = (context) => {
                const allElements = Array.from(context.querySelectorAll('*'));

                let weeklyWorkLabel = allElements.find(el => 
                    el.textContent && el.textContent.trim() === '本周工作'
                );

                if (!weeklyWorkLabel) {
                    weeklyWorkLabel = allElements.find(el => 
                        el.textContent && el.textContent.includes('本周工作')
                    );
                }

                if (weeklyWorkLabel) {
                    let parent = weeklyWorkLabel.parentElement;
                    for (let i = 0; i < 3 && parent; i++) {
                        const inputs = parent.querySelectorAll('textarea, input[type="text"], [contenteditable="true"]');
                        if (inputs.length > 0) return inputs[0];
                        parent = parent.parentElement;
                    }
                }

                const inputs = context.querySelectorAll('textarea, input[type="text"], [contenteditable="true"]');
                for (let input of inputs) {
                    const rect = input.getBoundingClientRect();
                    if (rect.width > 0 && rect.height > 0) {
                        return input;
                    }
                }

                return null;
            };

            let input = findWeeklyWorkInput(krElement);

            if (!input) {
                const oParent = krElement.closest('[class*="objective"], [class*="O"], section, div[class*="okr"]');
                if (oParent) {
                    input = findWeeklyWorkInput(oParent);
                }
            }

            if (!input) return false;

            input.scrollIntoView({ behavior: 'smooth', block: 'center' });
            input.focus();

            if (input.tagName === 'TEXTAREA' || input.tagName === 'INPUT') {
                input.value = content;
            } else {
                input.innerHTML = content.replace(/\\n/g, '<br>');
            }

            ['focus', 'input', 'change', 'blur'].forEach(eventType => {
                input.dispatchEvent(new Event(eventType, { bubbles: true }));
            });

            return true;
        }""", kr_element, content)

        if result:
            print(f"      [OK] 已填写本周工作")
            await asyncio.sleep(0.5)
            return True

    except Exception as e:
        print(f"      [!] KR上下文填写失败: {e}")

    return False

# ==================== 其他部分填写 ====================

async def fill_part1_work_summary(page, part1_data):
    """填写第一部分：本周工作总结"""
    print("\n[1/3] 填写本周工作总结...")

    if not part1_data:
        print("  [!] 无第一部分数据")
        return False

    title = part1_data.get("title", "本周工作总结")
    content = part1_data.get("content", "")

    result = await smart_fill_input(page, title, content, create_if_missing=True, section_type="part1")

    if not result:
        alternative_labels = ["工作总结", "本周工作", "工作汇总", "周报总结"]
        for alt_label in alternative_labels:
            result = await smart_fill_input(page, alt_label, content)
            if result:
                break

    return result

async def fill_part3_metrics(page, part3_data):
    """填写第三部分：关键指标"""
    print("\n[3/3] 填写关键指标...")

    if not part3_data:
        print("  [!] 无第三部分数据")
        return False

    title = part3_data.get("title", "关键指标")
    sections = part3_data.get("sections", {})
    all_success = True

    for section_title, section_data in sections.items():
        content = section_data.get("formatted_content", "")

        print(f"  处理: {section_title}")

        result = await smart_fill_input(
            page, 
            section_title, 
            content, 
            create_if_missing=True, 
            section_type="part3"
        )

        if not result:
            variants = []
            if "业务核心指标" in section_title:
                variants = ["核心指标", "业务指标", "指标", "数据"]
            elif "主要项目" in section_title:
                variants = ["项目", "项目进度", "项目进展"]
            elif "下周重点工作" in section_title:
                variants = ["下周工作", "下周计划", "下周", "计划"]

            for variant in variants:
                result = await smart_fill_input(page, variant, content)
                if result:
                    break

        if not result:
            all_success = False
            print(f"    [X] 无法填写: {section_title}")

        await asyncio.sleep(0.3)

    return all_success

# ==================== 提交和确认（自动版） ====================

async def auto_submit_and_confirm(page):
    """自动点击提交并处理确认弹框"""
    print("\n[提交] 自动查找并点击保存/提交按钮...")

    submit_clicked = False

    button_texts = ["保存", "提交", "更新", "确认", "Save", "Submit", "Update"]

    for text in button_texts:
        try:
            button = page.get_by_role("button", name=text).first
            if await button.is_visible(timeout=2000):
                print(f"  [OK] 找到提交按钮: {text}")
                await button.click()
                print(f"  [OK] 已点击 '{text}' 按钮")
                submit_clicked = True
                break
        except:
            continue

    if not submit_clicked:
        print("  [!] 使用JavaScript查找提交按钮...")
        result = await page.evaluate("""() => {
            const buttonTexts = ['保存', '提交', '更新', '确认', 'Save', 'Submit', 'Update'];
            const buttons = Array.from(document.querySelectorAll('button, input[type="submit"], a.btn, [role="button"]'));

            for (let btn of buttons) {
                const btnText = (btn.textContent || btn.value || '').trim();
                const isVisible = btn.offsetParent !== null && !btn.disabled;

                for (let text of buttonTexts) {
                    if (btnText.includes(text) && isVisible) {
                        btn.click();
                        return { success: true, text: btnText, class: btn.className };
                    }
                }
            }

            const primaryBtns = document.querySelectorAll('.ant-btn-primary, .btn-primary, [type="submit"]');
            for (let btn of primaryBtns) {
                if (btn.offsetParent !== null && !btn.disabled) {
                    btn.click();
                    return { success: true, text: btn.textContent, type: 'primary' };
                }
            }

            return { success: false };
        }""")

        if result.get('success'):
            print(f"  [OK] JavaScript点击提交: {result.get('text')}")
            submit_clicked = True

    if not submit_clicked:
        print("  [X] 未找到提交按钮")
        return False

    print("\n[确认] 等待确认弹框...")
    await asyncio.sleep(1.5)

    confirm_clicked = False

    confirm_texts = ["确认", "确定", "是的", "Yes", "OK", "Confirm", "继续", "保存"]

    for text in confirm_texts:
        try:
            confirm_btn = page.get_by_role("dialog").get_by_role("button", name=text).first
            if await confirm_btn.is_visible(timeout=3000):
                print(f"  [OK] 找到确认按钮: {text}")
                await confirm_btn.click()
                print(f"  [OK] 已自动点击确认按钮 '{text}'")
                confirm_clicked = True
                break
        except:
            try:
                confirm_btn = page.get_by_role("button", name=text).first
                if await confirm_btn.is_visible(timeout=2000):
                    print(f"  [OK] 找到确认按钮: {text}")
                    await confirm_btn.click()
                    print(f"  [OK] 已自动点击确认按钮 '{text}'")
                    confirm_clicked = True
                    break
            except:
                continue

    if not confirm_clicked:
        print("  [!] 使用JavaScript自动确认...")
        confirm_result = await page.evaluate("""() => {
            const confirmTexts = ['确认', '确定', '是的', 'Yes', 'OK', 'Confirm', '继续', '保存'];

            const dialogs = document.querySelectorAll('[role="dialog"], .modal, .ant-modal, .el-dialog, .ant-modal-wrap');
            for (let dialog of dialogs) {
                if (dialog.style.display === 'none') continue;

                const buttons = dialog.querySelectorAll('button, [role="button"], .ant-btn');
                for (let btn of buttons) {
                    const text = (btn.textContent || btn.value || '').trim();
                    const isVisible = btn.offsetParent !== null;

                    for (let ct of confirmTexts) {
                        if (text.includes(ct) && isVisible) {
                            btn.click();
                            return { success: true, location: 'dialog', text: text };
                        }
                    }
                }
            }

            const allButtons = Array.from(document.querySelectorAll('button, [role="button"]'));
            const confirmKeywords = ['确认', '确定', '是的', 'OK', 'Confirm', '保存', '提交'];

            for (let btn of allButtons) {
                const text = (btn.textContent || '').trim();
                const isVisible = btn.offsetParent !== null && !btn.disabled;
                const isConfirm = confirmKeywords.some(kw => text.includes(kw));

                if (isConfirm && isVisible) {
                    const isPrimary = btn.classList.contains('ant-btn-primary') || 
                                     btn.classList.contains('btn-primary') ||
                                     btn.getAttribute('type') === 'submit';

                    if (isPrimary || text.length < 10) {
                        btn.click();
                        return { success: true, location: 'global', text: text, primary: isPrimary };
                    }
                }
            }

            return { success: false };
        }""")

        if confirm_result.get('success'):
            print(f"  [OK] JavaScript自动确认: {confirm_result.get('text')}")
            confirm_clicked = True

    if confirm_clicked:
        print("  [OK] 提交并确认完成")
        await asyncio.sleep(2)
        return True
    else:
        print("  [!] 未找到确认按钮，但已尝试提交")
        return True

# ==================== 主流程 ====================

async def submit_single_weekly_v21(page, target_date=None, dry_run=True, auto_submit=False, is_first_page=False):
    """提交单周周报 - V2.1版本"""
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

    page_structure = await analyze_page_structure(page)

    print("\n滚动加载页面内容...")
    selector, center = await find_scroll_container(page)
    await scroll_container(page, selector, center)

    if dry_run:
        print("\n" + "="*60)
        print("【预览模式】以下内容将被填写：")
        print("="*60)

        if WEEKLY_DATA["part1_work_summary"]:
            print(f"\n1. 本周工作总结:\n{WEEKLY_DATA['part1_work_summary']['content'][:200]}...")

        if WEEKLY_DATA["part2_okr_structure"]:
            okr_count = sum(len(o.get("key_results", [])) for o in WEEKLY_DATA["part2_okr_structure"].get("objectives", []))
            print(f"\n2. OKR工作项数: {okr_count} 条KR")

        if WEEKLY_DATA["part3_metrics"]:
            metrics_count = len(WEEKLY_DATA["part3_metrics"].get("sections", {}))
            print(f"\n3. 其他内容项数: {metrics_count} 项")

        print("\n" + "="*60)
        print("预览完成，未实际填写")
        return {'status': 'preview', 'date': target_date}

    print("\n开始填写内容...")

    results = {}

    if WEEKLY_DATA["part1_work_summary"]:
        results['part1'] = await fill_part1_work_summary(page, WEEKLY_DATA["part1_work_summary"])
    else:
        print("\n[!] 无第一部分数据，跳过")
        results['part1'] = False

    if WEEKLY_DATA["part2_okr_structure"]:
        results['part2'] = await fill_okr_section(page, WEEKLY_DATA["part2_okr_structure"])
    else:
        print("\n[!] 无第二部分数据，跳过")
        results['part2'] = False

    if WEEKLY_DATA["part3_metrics"]:
        results['part3'] = await fill_part3_metrics(page, WEEKLY_DATA["part3_metrics"])
    else:
        print("\n[!] 无第三部分数据，跳过")
        results['part3'] = False

    print("\n" + "="*60)
    print("填写完成统计：")
    print(f"  本周工作总结: {'OK' if results.get('part1') else 'FAIL'}")
    print(f"  OKR本周工作: {'OK' if results.get('part2') else 'FAIL'}")
    print(f"  关键指标: {'OK' if results.get('part3') else 'FAIL'}")
    print("="*60)

    if auto_submit and any(results.values()):
        print("\n[自动提交] 准备自动提交...")
        submit_result = await auto_submit_and_confirm(page)
        results['submitted'] = submit_result

        if submit_result:
            print("[OK] 自动提交完成")
        else:
            print("[!] 自动提交可能失败")

    return {'status': 'success', 'date': target_date, 'results': results}

async def submit_weekly_with_retry_v21(page, **kwargs):
    """带重试机制的提交"""
    result = await submit_single_weekly_v21(page, is_first_page=True, **kwargs)

    if result['status'] == 'error':
        print(f"\n{'='*60}")
        print("[重试机制] 首次失败，准备重试...")
        print(f"{'='*60}")
        await asyncio.sleep(2)

        result = await submit_single_weekly_v21(page, is_first_page=False, **kwargs)

        if result['status'] == 'success':
            print("\n[OK] 重试成功")
        else:
            print("\n[X] 重试失败")

    return result

async def main():
    # 加载JSON数据（带详细调试）
    json_loaded = load_content_from_json()

    if not check_network_connection():
        print("\n[!] 网络检查失败")
        response = input("\n是否仍要继续? (y/N): ")
        if response.lower() != 'y':
            return

    target_date = get_default_date()

    print("="*60)
    print("百度OKR周报自动提交 - V2.1 (增强调试信息版)")
    print("="*60)
    print("\n【V2.1改进】")
    print("  ✓ 详细的JSON加载和解析调试信息")
    print("  ✓ 智能识别页面结构，自动创建缺失标题")
    print("  ✓ 精确定位OKR下的'本周工作'输入框")
    print("  ✓ 全自动提交+确认，无需手动干预")
    print(f"\n[配置]")
    print(f"  用户ID: {MY_ID}")
    print(f"  目标日期: {target_date}")
    print(f"  编辑模式: {'预览' if DRY_RUN else '实际填写'}")
    print(f"  自动提交: {'是' if AUTO_SUBMIT else '否'}")
    print(f"\n[内容来源]")
    if json_loaded:
        print(f"  ✓ {JSON_FILE_PATH}")
        if WEEKLY_DATA["part1_work_summary"]:
            print(f"  - 工作总结: {len(WEEKLY_DATA['part1_work_summary']['content'])} 字符")
        if WEEKLY_DATA["part2_okr_structure"]:
            kr_count = sum(len(o.get("key_results", [])) for o in WEEKLY_DATA["part2_okr_structure"].get("objectives", []))
            print(f"  - OKR工作: {kr_count} 条KR")
        if WEEKLY_DATA["part3_metrics"]:
            print(f"  - 关键指标: {len(WEEKLY_DATA['part3_metrics'].get('sections', {}))} 项")
    else:
        print(f"  ✗ 未能加载数据")

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
                    print(f"\n[X] 网络错误，无法访问OKR系统")
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

            # 执行提交
            result = await submit_weekly_with_retry_v21(
                page=page,
                target_date=target_date,
                dry_run=DRY_RUN,
                auto_submit=AUTO_SUBMIT
            )

            # 处理结果
            if result['status'] == 'login_required':
                print("\n[!] 需要重新登录")
            elif result['status'] == 'error':
                print(f"\n[X] 错误: {result.get('error')}")
            elif result['status'] == 'preview':
                print("\n[OK] 预览完成")
                print("\n要实际提交，请设置:")
                print("  DRY_RUN = False")
                print("  AUTO_SUBMIT = True")
            else:
                print("\n[OK] 填写流程完成")
                if AUTO_SUBMIT:
                    if result.get('results', {}).get('submitted'):
                        print("[OK] 已成功自动提交并确认")
                    else:
                        print("[!] 自动提交可能未完成，请检查页面")
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