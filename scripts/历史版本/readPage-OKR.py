from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import json
import time

def extract_element_details(element, element_type, idx):
    """提取单个元素的完整属性（不止基础属性，增加更多定位维度）"""
    details = {
        "index": idx,
        "selector": f"{element_type} >> nth={idx}",
        # 核心定位属性
        "id": element.get_attribute("id") or "",
        "name": element.get_attribute("name") or "",
        "class": element.get_attribute("class") or "",
        "data_attrs": {},  # 提取所有data-*自定义属性（前端常用）
        "xpath": element.locator("self::*").xpath()  # 生成精准xpath
    }
    
    # 提取所有data-*属性（比如data-id、data-name）
    all_attrs = element.get_attribute_names()
    for attr in all_attrs:
        if attr.startswith("data-"):
            details["data_attrs"][attr] = element.get_attribute(attr)
    
    # 按元素类型补充专属属性
    if element_type == "select":
        details["type"] = "select"
        details["options"] = [
            {
                "value": opt.get_attribute("value") or "",
                "text": opt.inner_text().strip() or "",
                "selected": opt.get_attribute("selected") is not None
            } 
            for opt in element.query_selector_all("option")
        ]
    elif element_type == "input":
        details["type"] = element.get_attribute("type") or "text"
        details["placeholder"] = element.get_attribute("placeholder") or ""
        details["value"] = element.get_attribute("value") or ""
    elif element_type == "button":
        details["type"] = element.get_attribute("type") or "button"
        details["text"] = element.inner_text().strip() or ""
        details["aria_label"] = element.get_attribute("aria-label") or ""  # 无障碍标签（常用）
    elif element_type == "form":
        details["action"] = element.get_attribute("action") or ""
        details["method"] = element.get_attribute("method") or "get"
    
    return details

def extract_frame_full_elements(frame, frame_name):
    """提取单个frame的所有元素（适配动态渲染）"""
    frame_elements = {
        "forms": [],
        "selects": [],
        "inputs": [],
        "buttons": [],
        "dynamic_check": f"提取时间：{time.strftime('%Y-%m-%d %H:%M:%S')}"
    }
    
    # 定义要提取的元素类型及定位器（覆盖更多场景）
    element_configs = [
        ("form", "form"),
        ("select", "select"),
        ("input", "input, [role='textbox'], .ant-input"),  # 增加前端框架类名（如antd）
        ("button", "button, [role='button'], .ant-btn, input[type='submit']")  # 覆盖自定义按钮
    ]
    
    for elem_type, locator in element_configs:
        try:
            # 等待元素加载（最长30秒，适配动态渲染）
            frame.wait_for_selector(locator, timeout=30000, state="attached")
            # 获取所有匹配元素（包括动态生成的）
            elements = frame.query_selector_all(locator)
            
            for idx, elem in enumerate(elements):
                try:
                    elem_details = extract_element_details(elem, elem_type, idx)
                    frame_elements[f"{elem_type}s"].append(elem_details)
                except Exception as e:
                    print(f"⚠️ 提取{frame_name}的{elem_type}[{idx}]失败：{str(e)[:50]}")
        except PlaywrightTimeoutError:
            print(f"⚠️ {frame_name}未找到{elem_type}元素（超时30秒）")
        except Exception as e:
            print(f"⚠️ 提取{frame_name}的{elem_type}元素失败：{str(e)[:50]}")
    
    return frame_elements

def extract_all_dynamic_elements(url, save_path="full_elements.json"):
    """提取所有动态渲染的元素（包括跨域iframe）"""
    with sync_playwright() as p:
        # 启动浏览器，禁用跨域限制+启用调试模式
        browser = p.chromium.launch(
            headless=False,
            slow_mo=500,  # 放慢操作，适配动态加载
            args=[
                "--disable-web-security",
                "--disable-features=CrossSiteDocumentBlockingIfIsolating",
                "--disable-blink-features=AutomationControlled"  # 避免被页面检测
            ]
        )
        
        # 创建上下文，仅配置视口（移除错误的default_navigation_timeout）
        context = browser.new_context(
            viewport={"width": 1920, "height": 1080}
        )
        page = context.new_page()
        
        # 为page设置导航超时（正确的位置）
        page.set_default_navigation_timeout(60000)
        
        try:
            # 访问页面，等待所有资源加载（包括动态JS）
            page.goto(
                url,
                wait_until="networkidle",  # 等待网络空闲（动态加载完成）
                timeout=60000  # 这里也可以单独设置goto的超时
            )
            
            # 额外等待5秒，确保JS渲染完成（针对慢加载页面）
            time.sleep(5)
            
            all_elements = {
                "主页面": extract_frame_full_elements(page.main_frame, "主页面"),
                "iframes": {}
            }
            
            # 提取所有iframe（包括嵌套、跨域）
            all_frames = page.frames
            for frame_idx, frame in enumerate(all_frames):
                if frame == page.main_frame:
                    continue  # 跳过主框架
                
                frame_name = f"iframe[{frame_idx}]"
                print(f"🔍 开始提取 {frame_name} 的元素...")
                try:
                    # 等待iframe加载完成
                    frame.wait_for_load_state("networkidle", timeout=30000)
                    all_elements["iframes"][frame_name] = extract_frame_full_elements(frame, frame_name)
                except PlaywrightTimeoutError:
                    all_elements["iframes"][frame_name] = {"error": "iframe加载超时30秒"}
                except Exception as e:
                    all_elements["iframes"][frame_name] = {"error": f"提取失败：{str(e)[:100]}"}
            
            # 保存完整结果
            with open(save_path, "w", encoding="utf-8") as f:
                json.dump(all_elements, f, ensure_ascii=False, indent=4)
            
            print(f"\n✅ 提取完成！结果已保存到 {save_path}")
            return all_elements
        
        except Exception as e:
            print(f"\n❌ 整体提取失败：{e}")
            return None
        finally:
            browser.close()

# 调用示例：替换为你的百度OKR内网地址
if __name__ == "__main__":
    # 请替换为实际的百度OKR页面URL
    TARGET_URL = "https://okr.baidu-int.com/pages/weekly.html#/home?date=s2026-02-16&mandatorId=s&isEdit=true"
    extract_all_dynamic_elements(TARGET_URL)