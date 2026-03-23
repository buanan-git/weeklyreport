#!/usr/bin/env python3
"""百度OKR周报自动提交脚本 - V2.4 (修复嵌套content结构)"""

import asyncio
import socket
import json
import os
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

WEEKLY_DATA = {
    "part1_work_summary": None,
    "part2_okr_structure": None,
    "part3_metrics": None
}

# ==================== 格式化配置 ====================

FORMAT_CONFIG = {
    "part1_work_summary": {
        "template": "{project_name}: {work_desc} (进度: {progress}, 状态: {status})",
        "fallback": "{item}",
        "separator": "\n"
    },
    "part2_okr_structure": {
        "weekly_work_template": "已完成: {completed}\n进行中: {in_progress}\n阻塞: {blocked}",
        "fallback": "{weekly_work}",
        "include_progress": True,
        "include_confidence": True
    },
    "part3_metrics": {
        "sections": {
            "业务核心指标": "{metric_name}: 当前{current_value} (目标{target_value}, 变化{change})",
            "主要项目": "{project_name}: 进度{progress} ({status})",
            "下周重点工作": "[{priority}] {work_item} (负责人: {owner}, 截止: {deadline})"
        },
        "default": "{item}",
        "separator": "\n"
    }
}

# ==================== 解析函数 ====================

def format_item(item, template, fallback="{item}"):
    if isinstance(item, dict):
        try:
            return template.format(**item)
        except KeyError as e:
            print(f"      [WARN] 缺少字段: {e}")
            return ", ".join([f"{k}={v}" for k, v in item.items()])
    elif isinstance(item, str):
        return fallback.format(item=item)
    return str(item)

def parse_part1(data):
    print(f"\n  [解析] 第一部分: 本周工作总结")

    # 处理直接是列表的情况
    if isinstance(data, list):
        print(f"    [INFO] 数据为列表类型，包含 {len(data)} 项")
        content_list = data
        title = "本周工作总结"
    elif isinstance(data, dict):
        title = data.get("title", "本周工作总结")
        content_list = data.get("content", [])
        # 如果content不是列表，尝试其他常见键
        if not isinstance(content_list, list):
            for key in ["items", "data", "work_items", "summary"]:
                if key in data and isinstance(data[key], list):
                    content_list = data[key]
                    print(f"    [INFO] 从 '{key}' 键提取内容")
                    break
        print(f"    标题: {title}")
        print(f"    content类型: {type(content_list)}")
    else:
        print(f"    [WARN] 不支持的数据类型: {type(data)}")
        return None

    # 扁平化嵌套结构
    content_list = flatten_nested_content(content_list)

    if not content_list:
        print(f"    [WARN] 内容为空")
        return None

    config = FORMAT_CONFIG["part1_work_summary"]
    formatted_items = []
    raw_items = []

    for idx, item in enumerate(content_list):
        print(f"    处理 item {idx+1}: {type(item)}")
        if isinstance(item, dict):
            raw_items.append(item)
            text = format_item(item, config["template"], config["fallback"])
            formatted_items.append(f"{idx+1}. {text}")
            print(f"      字段: {list(item.keys())}")
        elif isinstance(item, str):
            raw_items.append({"text": item})
            formatted_items.append(f"{idx+1}. {item}")
        else:
            # 其他类型转为字符串
            text = str(item)
            raw_items.append({"text": text})
            formatted_items.append(f"{idx+1}. {text}")

    result = {
        "title": title,
        "content": config["separator"].join(formatted_items),
        "raw_content": raw_items,
        "item_count": len(formatted_items)
    }

    print(f"    [OK] 完成: {result['item_count']} 条")
    print(f"    预览: {result['content'][:150]}...")
    return result

def parse_part2(data):
    print(f"\n  [解析] 第二部分: OKR进度")

    # 处理直接是列表的情况（objectives列表）
    if isinstance(data, list):
        print(f"    [INFO] 数据为列表类型，包含 {len(data)} 个Objective")
        objectives = data
        title = "OKR进度"
    elif isinstance(data, dict):
        title = data.get("title", "OKR进度")
        objectives = data.get("objectives", [])
        # 如果objectives不是列表，尝试其他常见键
        if not isinstance(objectives, list):
            for key in ["items", "data", "okrs", "goals"]:
                if key in data and isinstance(data[key], list):
                    objectives = data[key]
                    print(f"    [INFO] 从 '{key}' 键提取objectives")
                    break
        print(f"    标题: {title}")
        print(f"    Objectives: {len(objectives)}")
    else:
        print(f"    [WARN] 不支持的数据类型: {type(data)}")
        return None

    if not isinstance(objectives, list):
        print(f"    [WARN] objectives 不是列表: {type(objectives)}")
        return None

    config = FORMAT_CONFIG["part2_okr_structure"]
    okr_data = {"title": title, "objectives": []}

    for obj_idx, obj in enumerate(objectives):
        if not isinstance(obj, dict):
            print(f"    [WARN] Objective {obj_idx} 不是字典: {type(obj)}")
            continue

        o_id = obj.get("o_id", f"O{obj_idx+1}")
        o_title = obj.get("o_title", obj.get("title", ""))
        print(f"    Objective: {o_id} - {o_title}")

        o_data = {"o_id": o_id, "o_title": o_title, "key_results": []}
        key_results = obj.get("key_results", [])

        # 如果key_results不是列表，尝试其他常见键
        if not isinstance(key_results, list):
            for key in ["krs", "items", "results", "data"]:
                if key in obj and isinstance(obj[key], list):
                    key_results = obj[key]
                    print(f"      [INFO] 从 '{key}' 键提取key_results")
                    break

        if not isinstance(key_results, list):
            print(f"      [WARN] key_results 不是列表: {type(key_results)}")
            key_results = [key_results] if key_results else []

        for kr_idx, kr in enumerate(key_results):
            if not isinstance(kr, dict):
                print(f"      [WARN] KR {kr_idx} 不是字典: {type(kr)}")
                continue

            kr_id = kr.get("kr_id", f"KR{kr_idx+1}")
            kr_title = kr.get("kr_title", kr.get("title", ""))
            weekly_work = kr.get("weekly_work", "")
            progress = kr.get("progress", "")
            confidence = kr.get("confidence", "")

            print(f"      KR: {kr_id}")

            # 处理weekly_work可能是嵌套结构的情况
            if isinstance(weekly_work, dict):
                print(f"        weekly_work字段: {list(weekly_work.keys())}")
                formatted_work = format_item(weekly_work, config["weekly_work_template"], config["fallback"])
            elif isinstance(weekly_work, list):
                # 如果是列表，递归处理
                flattened = flatten_nested_content(weekly_work)
                formatted_work = "\n".join([str(x) for x in flattened])
            elif isinstance(weekly_work, str):
                formatted_work = weekly_work
            else:
                formatted_work = str(weekly_work)

            if config["include_progress"] and progress:
                formatted_work += f"\n进度: {progress}"
            if config["include_confidence"] and confidence:
                formatted_work += f" | 信心: {confidence}"

            o_data["key_results"].append({
                "kr_id": kr_id,
                "kr_title": kr_title,
                "weekly_work": formatted_work,
                "raw_weekly_work": weekly_work,
                "progress": progress,
                "confidence": confidence
            })

        okr_data["objectives"].append(o_data)

    total_krs = sum(len(o["key_results"]) for o in okr_data["objectives"])
    print(f"    [OK] 完成: {len(okr_data['objectives'])} 个O, {total_krs} 个KR")
    return okr_data

def flatten_nested_content(data, depth=0, max_depth=10):
    """递归扁平化嵌套结构，将嵌套的字典和列表转换为扁平列表"""
    if depth > max_depth:
        return [str(data)]

    if isinstance(data, list):
        result = []
        for item in data:
            result.extend(flatten_nested_content(item, depth + 1, max_depth))
        return result
    elif isinstance(data, dict):
        # 如果是字典，检查是否有特定字段
        if "content" in data:
            return flatten_nested_content(data["content"], depth + 1, max_depth)
        elif "items" in data:
            return flatten_nested_content(data["items"], depth + 1, max_depth)
        elif "data" in data:
            return flatten_nested_content(data["data"], depth + 1, max_depth)
        else:
            # 普通字典，保留为单个条目
            return [data]
    else:
        # 基本类型
        return [data] if data else []

def parse_part3(data):
    print(f"\n  [解析] 第三部分: 关键指标")

    # 处理直接是列表的情况
    if isinstance(data, list):
        print(f"    [INFO] 数据为列表类型，包含 {len(data)} 项")
        # 将列表转换为标准格式
        data = {"title": "关键指标", "sections": data}
    elif not isinstance(data, dict):
        print(f"    [WARN] 不支持的数据类型: {type(data)}")
        return None

    title = data.get("title", "关键指标")
    sections = data.get("sections", [])

    # 如果sections是字典，转换为列表
    if isinstance(sections, dict):
        print(f"    [INFO] sections是字典，转换为列表")
        sections_list = []
        for key, value in sections.items():
            if isinstance(value, dict):
                value["section_title"] = value.get("section_title", key)
                sections_list.append(value)
            elif isinstance(value, list):
                sections_list.append({"section_title": key, "content": value})
            else:
                sections_list.append({"section_title": key, "content": [value]})
        sections = sections_list

    print(f"    标题: {title}")
    print(f"    Sections: {len(sections)}")

    if not isinstance(sections, list):
        print(f"    [WARN] sections 不是列表: {type(sections)}")
        return None

    config = FORMAT_CONFIG["part3_metrics"]
    metrics_data = {"title": title, "sections": {}}

    for section_idx, section in enumerate(sections):
        if not isinstance(section, dict):
            print(f"    [WARN] Section {section_idx} 不是字典: {type(section)}")
            continue

        section_title = section.get("section_title", f"Section_{section_idx}")
        content_list = section.get("content", [])

        # 递归处理嵌套结构
        content_list = flatten_nested_content(content_list)

        print(f"    Section: {section_title} ({len(content_list)} 项)")

        template = config["sections"].get(section_title, config["default"])
        formatted_items = []
        raw_items = []

        for idx, item in enumerate(content_list):
            if isinstance(item, dict):
                raw_items.append(item)
                text = format_item(item, template, config["default"])
                formatted_items.append(f"{idx+1}. {text}")
                print(f"      字段: {list(item.keys())}")
            elif isinstance(item, str):
                raw_items.append({"text": item})
                formatted_items.append(f"{idx+1}. {item}")
            elif isinstance(item, (list, tuple)):
                # 处理嵌套列表
                nested_text = "; ".join([str(x) for x in item])
                raw_items.append({"nested_list": item})
                formatted_items.append(f"{idx+1}. {nested_text}")

        metrics_data["sections"][section_title] = {
            "raw_content": raw_items,
            "formatted_content": config["separator"].join(formatted_items),
            "item_count": len(formatted_items)
        }

        print(f"      [OK] {len(formatted_items)} 条")

    print(f"    [OK] 完成: {len(metrics_data['sections'])} 个section")
    return metrics_data

def parse_json_file(json_path):
    print(f"\n{'='*60}")
    print("[DEBUG] 开始解析JSON文件")
    print(f"{'='*60}")

    print(f"\n[DEBUG] 检查文件: {json_path}")
    if not os.path.exists(json_path):
        print("  [ERROR] 文件不存在")
        try:
            files = os.listdir(os.path.dirname(json_path) or ".")
            print(f"  目录文件: {files}")
        except:
            pass
        return None

    try:
        with open(json_path, "r", encoding="utf-8") as f:
            raw_content = f.read()
        print(f"  [OK] 读取: {len(raw_content)} 字符")
        print(f"  前500字符: {raw_content[:500]}")
    except Exception as e:
        print(f"  [ERROR] 读取失败: {e}")
        return None

    try:
        data = json.loads(raw_content)
        print(f"\n[DEBUG] JSON解析成功")
        print(f"  顶层键: {list(data.keys())}")
    except Exception as e:
        print(f"  [ERROR] JSON解析失败: {e}")
        return None

    # 检查是否在'content'键下
    if "content" in data:
        print(f"\n[DEBUG] 发现'content'键，使用嵌套结构")
        content_data = data.get("content", {})
        print(f"  content键下的键: {list(content_data.keys())}")
    else:
        print(f"\n[DEBUG] 使用顶层结构")
        content_data = data

    print(f"\n{'='*60}")
    print("开始解析各部分内容")
    print(f"{'='*60}")

    result = {
        "part1_work_summary": None,
        "part2_okr_structure": None,
        "part3_metrics": None
    }

    # 从content_data中解析各部分
    if "part1_work_summary" in content_data:
        print(f"\n  [OK] 找到 part1_work_summary")
        result["part1_work_summary"] = parse_part1(content_data["part1_work_summary"])
    else:
        print(f"\n  [WARN] 未找到 part1_work_summary")
        print(f"  可用键: {list(content_data.keys())}")

    if "part2_okr_structure" in content_data:
        print(f"\n  [OK] 找到 part2_okr_structure")
        result["part2_okr_structure"] = parse_part2(content_data["part2_okr_structure"])
    else:
        print(f"\n  [WARN] 未找到 part2_okr_structure")

    if "part3_metrics" in content_data:
        print(f"\n  [OK] 找到 part3_metrics")
        result["part3_metrics"] = parse_part3(content_data["part3_metrics"])
    else:
        print(f"\n  [WARN] 未找到 part3_metrics")

    print(f"\n{'='*60}")
    print("[DEBUG] 解析结果总结")
    print(f"{'='*60}")
    print(f"  第一部分: {'OK' if result['part1_work_summary'] else 'FAIL'}")
    print(f"  第二部分: {'OK' if result['part2_okr_structure'] else 'FAIL'}")
    print(f"  第三部分: {'OK' if result['part3_metrics'] else 'FAIL'}")

    return result

def load_content_from_json():
    global WEEKLY_DATA

    print(f"\n{'='*60}")
    print("[系统] 加载JSON文件")
    print(f"{'='*60}")

    try:
        script_dir = Path(__file__).parent.absolute()
    except:
        script_dir = Path(os.getcwd())

    json_path = script_dir / JSON_FILE_PATH
    print(f"文件路径: {json_path}")

    parsed = parse_json_file(str(json_path))

    if parsed:
        WEEKLY_DATA["part1_work_summary"] = parsed.get("part1_work_summary")
        WEEKLY_DATA["part2_okr_structure"] = parsed.get("part2_okr_structure")
        WEEKLY_DATA["part3_metrics"] = parsed.get("part3_metrics")

        print(f"\n[OK] 数据加载成功")
        print(f"\n[验证] 数据预览:")
        if WEEKLY_DATA["part1_work_summary"]:
            print(f"  第一部分: {len(WEEKLY_DATA['part1_work_summary']['content'])} 字符")
        if WEEKLY_DATA["part2_okr_structure"]:
            print(f"  第二部分: {len(WEEKLY_DATA['part2_okr_structure']['objectives'])} 个O")
        if WEEKLY_DATA["part3_metrics"]:
            print(f"  第三部分: {len(WEEKLY_DATA['part3_metrics']['sections'])} 个section")

        return True
    else:
        print(f"\n[ERROR] 加载失败")
        return False

def check_network_connection():
    print("\n[网络检测] 检查网络连接...")
    try:
        ip = socket.gethostbyname("okr.baidu-int.com")
        print(f"  [OK] 内网解析正常: {ip}")
        return True
    except:
        print(f"  [X] 内网解析失败，请检查VPN")
        return False

# ==================== 页面操作 ====================

async def analyze_page_structure(page):
    print("\n[页面分析] 分析页面结构...")
    structure = await page.evaluate("""() => {
        const result = {has_part1: false, has_part2: false, has_part3: false};
        document.querySelectorAll('h1, h2, h3, h4, [class*="title"]').forEach(h => {
            const t = h.textContent || '';
            if (t.includes('本周工作总结') || t.includes('工作总结')) result.has_part1 = true;
            if (t.includes('OKR') || t.includes('okr')) result.has_part2 = true;
            if (t.includes('指标') || t.includes('项目')) result.has_part3 = true;
        });
        return result;
    }""")
    print(f"  第一部分: {'存在' if structure['has_part1'] else '不存在'}")
    print(f"  第二部分: {'存在' if structure['has_part2'] else '不存在'}")
    print(f"  第三部分: {'存在' if structure['has_part3'] else '不存在'}")
    return structure

async def find_scroll_container(page):
    containers = await page.evaluate("""() => {
        const result = [];
        document.querySelectorAll("*").forEach(el => {
            const s = window.getComputedStyle(el);
            if (el.scrollHeight > el.clientHeight + 100 && el.scrollHeight > 500 &&
                (s.overflowY === 'auto' || s.overflowY === 'scroll')) {
                const r = el.getBoundingClientRect();
                result.push({class: el.className, id: el.id, tag: el.tagName, x: r.left + r.width/2, y: r.top + r.height/2});
            }
        });
        result.sort((a, b) => b.scrollHeight - a.scrollHeight);
        return result.slice(0, 1);
    }""")
    if containers:
        first = containers[0]
        selector = f"[class*='{first['class'].split()[0]}']" if first.get("class") else f"#{first['id']}" if first.get("id") else first["tag"].lower()
        return selector, (first.get("x", 960), first.get("y", 600))
    return "body", (960, 600)

async def scroll_container(page, selector, center):
    print("    滚动加载...", end=" ", flush=True)
    await page.mouse.move(center[0], center[1])
    for i in range(30):
        await page.mouse.wheel(0, 400)
        await asyncio.sleep(0.3)
        if i % 10 == 0:
            print(f"{i*400}px", end=" ", flush=True)
    print("OK")
    print("    等待渲染...", end=" ", flush=True)
    await asyncio.sleep(5)
    print("OK")


# ==================== 各部分填写函数 ====================

async def fill_part1_work_summary(page, part1_data):
    """填写第一部分：本周工作总结"""
    print("\n[填写] 第一部分：本周工作总结")

    content = part1_data.get("content", "")
    if not content:
        print("  [X] 内容为空")
        return False

    # 显示准备写入的内容
    print(f"\n{'='*60}")
    print("[准备写入 - 第一部分]")
    print(f"{'='*60}")
    print(f"内容长度: {len(content)} 字符")
    print(f"完整内容:\n{content}")
    print(f"{'='*60}\n")

    # 尝试多种标签文本匹配
    label_variants = ["本周工作总结", "工作总结", "本周工作", "工作总"]

    for label in label_variants:
        try:
            # 先尝试使用 smart_fill_input
            result = await smart_fill_input(page, label, content, create_if_missing=True, section_type="part1")
            if result:
                print(f"  [OK] 成功填写: {label}")
                return True
        except Exception as e:
            print(f"  [!] {label} 尝试失败: {e}")
            continue

    # 如果都失败了，使用 JavaScript 兜底
    print("  使用 JavaScript 兜底填写...")
    return await fill_by_javascript(page, "本周工作总结", content)

async def fill_okr_section(page, part2_data):
    """填写第二部分：OKR进度"""
    print("\n[填写] 第二部分：OKR进度")

    objectives = part2_data.get("objectives", [])
    if not objectives:
        print("  [X] 没有OKR数据")
        return False

    success_count = 0
    total_krs = 0

    for obj in objectives:
        o_title = obj.get("o_title", "")
        key_results = obj.get("key_results", [])

        print(f"\n  Objective: {o_title}")

        for kr in key_results:
            kr_title = kr.get("kr_title", "")
            weekly_work = kr.get("weekly_work", "")
            kr_id = kr.get("kr_id", "")

            total_krs += 1

            if not weekly_work:
                print(f"    [跳过] {kr_id}: 无内容")
                continue

            # 构建搜索标签：优先使用KR标题，其次使用KR ID
            search_labels = [kr_title] if kr_title else []
            if kr_id and kr_id not in search_labels:
                search_labels.append(kr_id)

            filled = False
            for label in search_labels:
                try:
                    result = await smart_fill_input(page, label, weekly_work)
                    if result:
                        print(f"    [OK] {kr_id}: 已填写 ({len(weekly_work)}字符)")
                        filled = True
                        break
                except:
                    continue

            if not filled:
                # 尝试通过O标题找到区域再填写
                try:
                    result = await fill_by_javascript(page, o_title, weekly_work)
                    if result:
                        print(f"    [OK] {kr_id}: JavaScript填写成功")
                        filled = True
                except Exception as e:
                    print(f"    [X] {kr_id}: 填写失败 - {e}")

            if filled:
                success_count += 1

    print(f"\n  填写统计: {success_count}/{total_krs} 个KR成功")
    return success_count > 0

async def fill_part3_metrics(page, part3_data):
    """填写第三部分：关键指标"""
    print("\n[填写] 第三部分：关键指标")

    sections = part3_data.get("sections", {})
    if not sections:
        print("  [X] 没有指标数据")
        return False

    success_count = 0
    total_sections = len(sections)

    for section_name, section_data in sections.items():
        content = section_data.get("formatted_content", "")
        if not content:
            content = section_data.get("raw_content", "")
            if isinstance(content, list):
                content = "\n".join([str(item) for item in content])

        if not content:
            print(f"  [跳过] {section_name}: 无内容")
            continue

        # 显示准备写入的内容
        print(f"\n  {'='*60}")
        print(f"  [准备写入 - {section_name}]")
        print(f"  {'='*60}")
        print(f"  内容长度: {len(content)} 字符")
        print(f"  完整内容:\n{content}")
        print(f"  {'='*60}")

        # 尝试多种标签匹配
        label_variants = [section_name]

        # 添加常见变体
        if "业务核心指标" in section_name:
            label_variants.extend(["核心指标", "业务指标", "指标"])
        elif "主要项目" in section_name:
            label_variants.extend(["项目", "重点项目"])
        elif "下周重点工作" in section_name:
            label_variants.extend(["下周工作", "重点工作", "下周计划"])

        filled = False
        for label in label_variants:
            try:
                result = await smart_fill_input(page, label, content)
                if result:
                    print(f"    [OK] 已填写: {label} ({len(content)}字符)")
                    filled = True
                    break
            except:
                continue

        if not filled:
            # 使用 JavaScript 兜底
            result = await fill_by_javascript(page, section_name, content)
            if result:
                print(f"    [OK] JavaScript填写成功")
                filled = True
            else:
                print(f"    [X] 填写失败: {section_name}")

        if filled:
            success_count += 1

    print(f"\n  填写统计: {success_count}/{total_sections} 个section成功")
    return success_count > 0



# ==================== 遮挡层处理 ====================

async def remove_overlay_elements(page):
    """移除页面上的遮挡层元素"""
    try:
        await page.evaluate("""() => {
            // 移除常见的遮挡层类名
            const overlay_classes = [
                'mp_ai_plugin_shade',
                'MuiBackdrop-root',
                'MuiPopover-root',
                'ant-modal-mask',
                'ant-modal-wrap',
                'el-overlay',
                'ui-widget-overlay',
                'modal-backdrop',
                'overlay',
                'shade'
            ];

            overlay_classes.forEach(cls => {
                document.querySelectorAll('.' + cls).forEach(el => {
                    el.style.display = 'none';
                    el.style.pointerEvents = 'none';
                    if (el.parentNode) el.parentNode.removeChild(el);
                });
            });

            // 移除固定定位的全屏元素（可能是遮挡层）
            document.querySelectorAll('*').forEach(el => {
                const style = window.getComputedStyle(el);
                if (style.position === 'fixed' && 
                    style.zIndex > 1000 && 
                    el.offsetWidth > window.innerWidth * 0.5 &&
                    el.offsetHeight > window.innerHeight * 0.5) {
                    el.style.pointerEvents = 'none';
                }
            });

            return 'ok';
        }""")
        print("    [OK] 已清理遮挡层")
        return True
    except Exception as e:
        print(f"    [!] 清理遮挡层失败: {e}")
        return False

async def smart_click_with_fallback(page, locator, timeout=5000):
    """智能点击：尝试标准点击 -> 强制点击 -> JavaScript点击"""
    # 先尝试移除遮挡层
    await remove_overlay_elements(page)

    # 1. 尝试标准点击
    try:
        await locator.click(timeout=timeout)
        print("    [OK] 标准点击成功")
        return True
    except Exception as e:
        if "intercepts pointer events" in str(e):
            print("    [!] 被遮挡层拦截，尝试强制点击...")
        else:
            print(f"    [!] 标准点击失败: {e}")

    # 2. 尝试强制点击
    try:
        await locator.click(force=True, timeout=timeout)
        print("    [OK] 强制点击成功")
        return True
    except Exception as e:
        print(f"    [!] 强制点击失败: {e}")

    # 3. 使用JavaScript点击
    try:
        handle = await locator.element_handle()
        if handle:
            await handle.evaluate("el => { el.click(); el.dispatchEvent(new Event('click', {bubbles: true})); }")
            print("    [OK] JavaScript点击成功")
            return True
    except Exception as e:
        print(f"    [!] JavaScript点击失败: {e}")

    return False

async def smart_fill_input(page, label_text, content, create_if_missing=False, section_type=None):
    if not content or not content.strip():
        print(f"  [跳过] 内容为空: {label_text}")
        return False

    content = content.strip()
    print(f"\n  处理: {label_text[:40]}...")
    print(f"    内容长度: {len(content)}")
    print(f"    预览: {content[:80]}...")

    variants = [label_text]
    if "本周工作总结" in label_text:
        variants.extend(["工作总结", "本周工作"])
    elif "OKR" in label_text:
        variants.extend(["OKR进度", "OKR"])

    for variant in variants:
        try:
            locator = page.get_by_text(variant, exact=False).first
            if await locator.is_visible(timeout=2000):
                print(f"    [OK] 找到标签: {variant}")
                # 使用新的点击方式
                if await smart_click_with_fallback(page, locator):
                    await asyncio.sleep(0.3)
                    if await find_and_fill_nearby_input(page, locator, content):
                        return True
        except Exception as e:
            print(f"    [!] 处理 {variant} 时出错: {e}")
            continue

    print("    尝试JavaScript直接填写...")
    if await fill_by_javascript(page, label_text, content):
        return True

    if create_if_missing and section_type:
        print("    尝试创建新区域...")
        if await create_new_section(page, section_type, label_text, content):
            return True

    print(f"    [X] 失败: {label_text}")
    return False
async def find_and_fill_nearby_input(page, label_locator, content):
    try:
        # 使用智能点击代替直接点击
        if not await smart_click_with_fallback(page, label_locator):
            print("    [!] 标签点击失败，尝试直接查找输入框...")

        await asyncio.sleep(0.5)
        handle = await label_locator.element_handle()
        if not handle:
            return False
        text = await handle.inner_text()
        selectors = ['textarea', 'input[type="text"]', '[contenteditable="true"]', '.ant-input']
        for sel in selectors:
            try:
                inp = page.locator(f'xpath=//*[contains(text(), "{text}")]/following::{sel}[1] | xpath=//*[contains(text(), "{text}")]/ancestor::div[1]//{sel}').first
                if await inp.is_visible(timeout=2000):
                    print(f"    [OK] 找到输入框: {sel}")
                    await inp.fill(content)
                    await asyncio.sleep(0.5)
                    return True
            except:
                continue
    except Exception as e:
        print(f"    [!] 失败: {e}")
    return False
async def fill_by_javascript(page, label_text, content):
    print(f"    [DEBUG-JS] 开始JavaScript填写，标签: {label_text[:30]}...")
    print(f"    [DEBUG-JS] 内容长度: {len(content)} 字符")

    try:
        # 在Python中预处理内容
        normalized = content.replace("\r\n", "\n").replace("\r", "\n")
        lines = normalized.split("\n")

        # 使用原始字符串避免转义问题
        js_code = """
        (args) => {
            const labelText = args.labelText;
            const lines = args.lines;

            console.log('JS: 开始查找标签:', labelText);
            const findEl = (text) => {
                const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
                let node; 
                while (node = walker.nextNode()) {
                    if (node.textContent.trim().includes(text)) return node.parentElement;
                }
                return null;
            };

            const findInput = (el) => {
                if (el.matches('textarea, input[type="text"], [contenteditable="true"]')) return el;
                let inputs = el.querySelectorAll('textarea, input[type="text"], [contenteditable="true"]');
                if (inputs.length) return inputs[0];
                let parent = el.parentElement;
                for (let i=0; i<5 && parent; i++) {
                    inputs = parent.querySelectorAll('textarea, input[type="text"], [contenteditable="true"]');
                    if (inputs.length) return inputs[0];
                    parent = parent.parentElement;
                }
                return null;
            };

            const label = findEl(labelText);
            if (!label) {
                console.log('JS: 未找到标签');
                return 'no_label';
            }

            const input = findInput(label);
            if (!input) {
                console.log('JS: 未找到输入框');
                return 'no_input';
            }
            console.log('JS: 找到输入框:', input.tagName);

            input.scrollIntoView({behavior:'smooth', block:'center'});
            input.focus();

            // 将行数组合并 - 使用\\n表示换行符
            const fullContent = lines.join('\\n');

            if (input.tagName === 'TEXTAREA' || input.tagName === 'INPUT') {
                input.value = fullContent;
                console.log('JS: 已设置value, 长度:', fullContent.length);
            } else {
                // 对于contenteditable，使用<br>标签
                const htmlContent = lines.join('<br>');
                input.innerHTML = htmlContent;
                console.log('JS: 已设置innerHTML');
            }

            // 触发事件
            const events = ['focus', 'input', 'change', 'blur'];
            events.forEach(eventType => {
                const event = new Event(eventType, {bubbles: true});
                input.dispatchEvent(event);
            });

            console.log('JS: 已触发事件');
            return 'ok';
        }
        """

        result = await page.evaluate(js_code, {"labelText": label_text, "lines": lines})

        print(f"    [DEBUG-JS] 执行结果: {result}")
        if result == 'ok':
            print(f"    [OK] JavaScript填写成功")
            await asyncio.sleep(0.5)
            return True
        else:
            print(f"    [!] JavaScript返回: {result}")
    except Exception as e:
        print(f"    [!] JavaScript错误: {e}")
        import traceback
        traceback.print_exc()
    return False




async def create_new_section(page, section_type, title, content):
    print(f"    [DEBUG] 创建新区域: {title[:30]}...")
    try:
        result = await page.evaluate("""(args) => {
            const title = args.title;
            const content = args.content;

            const container = document.querySelector('.ant-form, .weekly-form, main') || document.body;
            const div = document.createElement('div');
            div.style.cssText = 'margin-top:20px;padding:16px;border:1px solid #d9d9d9;border-radius:4px;';

            const h3 = document.createElement('h3');
            h3.textContent = title;
            h3.style.marginBottom = '12px';

            const textarea = document.createElement('textarea');
            textarea.value = content;
            textarea.style.cssText = 'width:100%;min-height:120px;padding:8px;';

            div.appendChild(h3);
            div.appendChild(textarea);
            container.appendChild(div);

            textarea.focus();
            textarea.dispatchEvent(new Event('input', {bubbles:true}));

            return 'created';
        }""", {"title": title, "content": content})

        if result == 'created':
            print(f"    [OK] 创建成功")
            return True
    except Exception as e:
        print(f"    [X] 创建失败: {e}")
    return False


# ==================== 提交和确认 ====================

async def auto_submit_and_confirm(page):
    print("\n[提交] 自动提交...")

    submit_clicked = False
    for text in ["保存", "提交", "更新", "Save", "Submit"]:
        try:
            button = page.get_by_role("button", name=text).first
            if await button.is_visible(timeout=2000):
                print(f"  点击: {text}")
                await button.click()
                submit_clicked = True
                break
        except:
            continue

    if not submit_clicked:
        result = await page.evaluate("""() => {
            const buttons = Array.from(document.querySelectorAll('button, input[type="submit"]'));
            for (let btn of buttons) {
                const text = (btn.textContent || btn.value || '').trim();
                if ((text.includes('保存') || text.includes('提交')) && btn.offsetParent !== null) {
                    btn.click();
                    return true;
                }
            }
            return false;
        }""")
        if result:
            submit_clicked = True
            print("  JavaScript点击提交")

    if not submit_clicked:
        print("  [X] 未找到提交按钮")
        return False

    print("\n[确认] 处理确认弹框...")
    await asyncio.sleep(1.5)

    confirm_clicked = False
    for text in ["确认", "确定", "Yes", "OK", "Confirm"]:
        try:
            confirm_btn = page.get_by_role("dialog").get_by_role("button", name=text).first
            if await confirm_btn.is_visible(timeout=3000):
                print(f"  点击确认: {text}")
                await confirm_btn.click()
                confirm_clicked = True
                break
        except:
            try:
                confirm_btn = page.get_by_role("button", name=text).first
                if await confirm_btn.is_visible(timeout=2000):
                    await confirm_btn.click()
                    confirm_clicked = True
                    break
            except:
                continue

    if not confirm_clicked:
        result = await page.evaluate("""() => {
            const dialogs = document.querySelectorAll('[role="dialog"], .modal, .ant-modal');
            for (let dialog of dialogs) {
                if (dialog.style.display === 'none') continue;
                const buttons = dialog.querySelectorAll('button');
                for (let btn of buttons) {
                    const text = (btn.textContent || '').trim();
                    if ((text.includes('确认') || text.includes('确定')) && btn.offsetParent !== null) {
                        btn.click();
                        return true;
                    }
                }
            }
            return false;
        }""")
        if result:
            confirm_clicked = True
            print("  JavaScript点击确认")

    if confirm_clicked:
        print("  [OK] 提交完成")
        await asyncio.sleep(2)
        return True
    return False

# ==================== 主流程 ====================

async def submit_single_weekly_v24(page, target_date=None, dry_run=True, auto_submit=False, is_first_page=False):
    target_date = target_date or get_default_date()
    edit_url = f"{EDIT_BASE_URL}?date={target_date}&mandatorId={MY_ID}&isEdit=true"

    print("="*60)
    print(f"目标日期: {target_date}")
    print(f"编辑模式: {'预览' if dry_run else '实际填写'}")
    print(f"自动提交: {'是' if auto_submit else '否'}")
    print("="*60)

    try:
        try:
            await page.goto(edit_url, wait_until='domcontentloaded', timeout=60000)
        except Exception as nav_error:
            if 'uuap' in page.url:
                return {'status': 'login_required', 'url': page.url}

        if is_first_page:
            print("\n(首次页面等待5秒)...", end=" ", flush=True)
            await asyncio.sleep(5)
            print("OK")
        else:
            await asyncio.sleep(2)

        if 'uuap' in page.url:
            return {'status': 'login_required', 'url': page.url}

        print("[OK] 页面加载完成")

    except Exception as e:
        print(f"\n[X] 页面加载失败: {e}")
        return {'status': 'error', 'error': str(e)}

    await analyze_page_structure(page)

    print("\n滚动加载页面内容...")
    selector, center = await find_scroll_container(page)
    await scroll_container(page, selector, center)

    print("\n" + "="*60)
    print("[验证] 数据状态检查")
    print("="*60)
    print(f"  part1_work_summary: {WEEKLY_DATA.get('part1_work_summary') is not None}")
    print(f"  part2_okr_structure: {WEEKLY_DATA.get('part2_okr_structure') is not None}")
    print(f"  part3_metrics: {WEEKLY_DATA.get('part3_metrics') is not None}")

    if dry_run:
        print("\n" + "="*60)
        print("【预览模式】")
        print("="*60)

        if WEEKLY_DATA.get("part1_work_summary"):
            print(f"\n1. 本周工作总结:")
            print(f"{WEEKLY_DATA['part1_work_summary']['content'][:300]}...")

        if WEEKLY_DATA.get("part2_okr_structure"):
            print(f"\n2. OKR进度:")
            for obj in WEEKLY_DATA["part2_okr_structure"]["objectives"]:
                print(f"  {obj['o_id']}: {obj['o_title']}")
                for kr in obj["key_results"]:
                    print(f"    {kr['kr_id']}: {kr['weekly_work'][:50]}...")

        if WEEKLY_DATA.get("part3_metrics"):
            print(f"\n3. 关键指标:")
            for section_name, section_data in WEEKLY_DATA["part3_metrics"]["sections"].items():
                print(f"  {section_name}:")
                print(f"    {section_data['formatted_content'][:100]}...")

        return {'status': 'preview', 'date': target_date}

    print("\n开始填写内容...")

    results = {}

    part1_data = WEEKLY_DATA.get("part1_work_summary")
    part2_data = WEEKLY_DATA.get("part2_okr_structure")
    part3_data = WEEKLY_DATA.get("part3_metrics")

    print(f"\n[填写前数据检查]")
    print(f"  part1_data: {part1_data is not None}")
    print(f"  part2_data: {part2_data is not None}")
    print(f"  part3_data: {part3_data is not None}")

    if part1_data:
        results['part1'] = await fill_part1_work_summary(page, part1_data)
    else:
        print("\n[!] 无第一部分数据，跳过")
        results['part1'] = False

    if part2_data:
        results['part2'] = await fill_okr_section(page, part2_data)
    else:
        print("\n[!] 无第二部分数据，跳过")
        results['part2'] = False

    if part3_data:
        results['part3'] = await fill_part3_metrics(page, part3_data)
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

async def submit_weekly_with_retry_v24(page, **kwargs):
    result = await submit_single_weekly_v24(page, is_first_page=True, **kwargs)

    if result['status'] == 'error':
        print(f"\n{'='*60}")
        print("[重试机制] 首次失败，准备重试...")
        print(f"{'='*60}")
        await asyncio.sleep(2)

        result = await submit_single_weekly_v24(page, is_first_page=False, **kwargs)

        if result['status'] == 'success':
            print("\n[OK] 重试成功")
        else:
            print("\n[X] 重试失败")

    return result

async def main():
    print("\n" + "="*60)
    print("程序启动 - 开始加载数据")
    print("="*60)

    json_loaded = load_content_from_json()

    if not check_network_connection():
        print("\n[!] 网络检查失败")
        response = input("\n是否仍要继续? (y/N): ")
        if response.lower() != 'y':
            return

    target_date = get_default_date()

    print("="*60)
    print("百度OKR周报自动提交 - V2.4 (修复嵌套content结构)")
    print("="*60)
    print("\n【V2.4改进】")
    print("  ✓ 支持嵌套在'content'键下的JSON结构")
    print("  ✓ 支持content字段下的详细关键字段解析")
    print("  ✓ 可配置的格式化模板")
    print("  ✓ 详细的调试信息")
    print(f"\n[配置]")
    print(f"  用户ID: {MY_ID}")
    print(f"  目标日期: {target_date}")
    print(f"  编辑模式: {'预览' if DRY_RUN else '实际填写'}")
    print(f"  自动提交: {'是' if AUTO_SUBMIT else '否'}")
    print(f"\n[内容来源]")
    if json_loaded:
        print(f"  ✓ {JSON_FILE_PATH} 加载成功")
        print(f"\n  [数据状态确认]")
        print(f"    part1_work_summary: {WEEKLY_DATA.get('part1_work_summary') is not None}")
        print(f"    part2_okr_structure: {WEEKLY_DATA.get('part2_okr_structure') is not None}")
        print(f"    part3_metrics: {WEEKLY_DATA.get('part3_metrics') is not None}")

        if WEEKLY_DATA.get("part1_work_summary"):
            print(f"    - 工作总结: {len(WEEKLY_DATA['part1_work_summary']['content'])} 字符")
        if WEEKLY_DATA.get("part2_okr_structure"):
            kr_count = sum(len(o.get("key_results", [])) for o in WEEKLY_DATA["part2_okr_structure"].get("objectives", []))
            print(f"    - OKR工作: {kr_count} 条KR")
        if WEEKLY_DATA.get("part3_metrics"):
            print(f"    - 关键指标: {len(WEEKLY_DATA['part3_metrics'].get('sections', {}))} 项")
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

            result = await submit_weekly_with_retry_v24(
                page=page,
                target_date=target_date,
                dry_run=DRY_RUN,
                auto_submit=AUTO_SUBMIT
            )

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