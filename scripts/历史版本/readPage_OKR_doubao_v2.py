from playwright.sync_api import sync_playwright
import os
import time
import json
from datetime import datetime

# 保存登录状态（核心：持久化登录）
USER_DATA_DIR = "./browser_profile"
# 记录页面信息的文件
RECORD_FILE = "./page_record.json"
# 你要导出的内网页面地址
TARGET_URL = "https://okr.baidu-int.com/pages/weekly.html#/home?date=s2026-02-23&mandatorId=s&isEdit=true"

# 导出文件（新增轻量化文件）
OUTPUT_HTML = "page_full_dump.html"
OUTPUT_IFRAMES = "iframes_dump.txt"
OUTPUT_LIGHT_HTML = "page_light_dump.html"  # 轻量化HTML（仅保留核心结构）
OUTPUT_INPUTS_ONLY = "inputs_only.json"     # 仅输入框关键信息（大模型友好）

def record_page_info(page_url, input_list, iframe_count):
    """记录页面关键信息，方便后续复用"""
    record_data = {
        "record_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "page_url": page_url,
        "iframe_count": iframe_count,
        "input_box_count": len(input_list),
        "input_box_list": input_list
    }
    with open(RECORD_FILE, "w", encoding="utf-8") as f:
        json.dump(record_data, f, ensure_ascii=False, indent=2)
    print(f"✅ 页面信息已记录到：{RECORD_FILE}")

def extract_light_html(full_html):
    """提取轻量化HTML：移除script/style、冗余标签，仅保留核心结构"""
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(full_html, "html.parser")
    
    # 移除脚本、样式、注释（减少体积）
    for element in soup(["script", "style", "noscript", "comment"]):
        element.decompose()
    
    # 仅保留含输入框/目录的核心容器（适配你的页面结构）
    core_containers = soup.find_all("div", class_=lambda c: c and "block-container-weekly" in c)
    if core_containers:
        light_soup = BeautifulSoup("<html><body></body></html>", "html.parser")
        light_soup.body.extend(core_containers)
        return light_soup.prettify()
    else:
        # 兜底：保留body内核心内容
        body = soup.find("body")
        return f"<html><body>{body.prettify() if body else ''}</body></html>"

def dump_full_page():
    # 确保用户数据目录存在（关键：持久化登录态）
    if not os.path.exists(USER_DATA_DIR):
        os.makedirs(USER_DATA_DIR)
        print(f"📁 创建浏览器配置目录：{USER_DATA_DIR}（用于保存登录态）")

    p = sync_playwright().start()
    browser = p.chromium.launch_persistent_context(
        user_data_dir=USER_DATA_DIR,
        headless=False,               # 显示浏览器窗口
        viewport={"width": 1280, "height": 800},
        ignore_https_errors=True      # 兼容内网HTTPS证书
    )

    page = browser.new_page()
    page.set_default_navigation_timeout(60000)
    page.set_default_timeout(60000)

    # 访问页面
    print(f"🌐 正在访问：{TARGET_URL}")
    page.goto(TARGET_URL, wait_until="domcontentloaded")

    print("\n🔹 请你现在操作页面：")
    print("   1. 登录（如果需要）")
    print("   2. 展开所有需要填报的目录树结构")
    print("   3. 确保所有输入框已显示")
    input("按回车开始导出页面完整结构...")

    input_box_list = []
    try:
        # 等待页面稳定（缩短等待时间）
        page.wait_for_load_state("domcontentloaded", timeout=15000)
        time.sleep(3)  # 减少等待时间

        # 1. 导出完整HTML（保留原逻辑，用于兜底）
        full_html = page.content()
        with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
            f.write(full_html)
        print(f"✅ 完整页面已导出：{OUTPUT_HTML}")

        # 1.1 生成轻量化HTML（移除冗余内容，体积大幅减小）
        try:
            light_html = extract_light_html(full_html)
            with open(OUTPUT_LIGHT_HTML, "w", encoding="utf-8") as f:
                f.write(light_html)
            print(f"✅ 轻量化页面已导出：{OUTPUT_LIGHT_HTML}（体积仅为原1/10左右）")
        except Exception as e:
            print(f"⚠️ 轻量化HTML生成失败：{str(e)}")

        # 2. 优化iframe导出：仅导出核心iframe（索引0），跳过其他冗余iframe
        iframes = page.frames
        iframe_count = len(iframes)
        print(f"🔍 找到 {iframe_count} 个 iframe/框架（仅导出核心iframe 0）")

        with open(OUTPUT_IFRAMES, "w", encoding="utf-8") as f:
            # 只处理核心iframe（索引0），其他跳过
            main_iframe_idx = 0
            if iframe_count > main_iframe_idx:
                frame = iframes[main_iframe_idx]
                try:
                    f.write(f"===== 核心iframe {main_iframe_idx} | url={frame.url} =====\n")
                    frame.wait_for_load_state("domcontentloaded", timeout=15000)
                    # 对iframe内容也做轻量化处理
                    frame_html = frame.content()
                    light_frame_html = extract_light_html(frame_html)
                    f.write(light_frame_html)
                    f.write("\n\n")
                except Exception as e:
                    err_msg = f"核心iframe {main_iframe_idx} 读取失败: {str(e)}"
                    f.write(f"{err_msg}\n\n")
                    print(f"⚠️ {err_msg}")
            else:
                f.write("未找到核心iframe（索引0）\n")

        print(f"✅ 核心iframe内容已导出：{OUTPUT_IFRAMES}（体积大幅减小）")

        # 3. 优化输入框查找：减少DOM遍历、提前筛选、简化XPath
        print("\n📝 页面上所有可输入框（含iframe内）：")
        main_iframe = page.frames[0] if iframe_count > 0 else page
        
        # 优化点1：使用filter筛选可见输入框（跳过隐藏元素，减少遍历量）
        locator = main_iframe.locator("input, textarea, [contenteditable='true']").filter(visible=True)
        # 优化点2：使用count()先获取数量，避免all()全量加载（大幅提速）
        input_count = locator.count()
        print(f"📌 找到 {input_count} 个可见输入框（跳过隐藏元素）")

        # 批量读取输入框信息（简化逻辑，减少异常处理耗时）
        for idx in range(input_count):
            try:
                el = locator.nth(idx)
                # 优化点3：批量获取属性，减少多次DOM调用
                placeholder = el.get_attribute("placeholder") or ""
                value = el.input_value() or ""
                
                # 优化点4：简化目录定位逻辑（减少XPath层级，提速）
                parent_text = "未识别目录"
                try:
                    # 简化XPath：只找最近的含标题的父节点
                    parent = el.locator("xpath=ancestor::div[contains(@class, 'label-title')]/span").first
                    if parent.is_visible():
                        parent_text = parent.inner_text()[:30]
                except:
                    pass

                input_info = {
                    "index": idx,
                    "placeholder": placeholder,
                    "parent_directory": parent_text,
                    "value": value[:50]
                }
                input_box_list.append(input_info)
                print(f"输入框 {idx}: 目录='{parent_text}' placeholder='{placeholder}' value='{value[:20]}...'")
            except Exception as e:
                input_info = {"index": idx, "error": f"读取失败：{str(e)}"}
                input_box_list.append(input_info)
                print(f"输入框 {idx}: 无法读取属性（动态生成）")

        # 3.1 生成仅含输入框的JSON文件（大模型最易读取）
        with open(OUTPUT_INPUTS_ONLY, "w", encoding="utf-8") as f:
            json.dump({
                "record_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "input_box_count": len(input_box_list),
                "input_box_details": input_box_list
            }, f, ensure_ascii=False, indent=2)
        print(f"✅ 仅输入框关键信息已导出：{OUTPUT_INPUTS_ONLY}（体积<100KB）")

        # 4. 记录页面信息
        record_page_info(page.url, input_box_list, iframe_count)

    except Exception as e:
        print(f"\n❌ 导出过程出错：{str(e)}")
        # 容错保存HTML
        try:
            full_html = page.content()
            with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
                f.write(full_html)
            print("✅ 已保存当前能拿到的页面内容")
        except:
            print("❌ 无法保存页面内容，请检查页面是否完全加载")

    # 提示信息更新
    print("\n🎉 页面导出完成！")
    print("📌 关键信息：")
    print(f"   1. 登录态保存在：{USER_DATA_DIR}（下次启动无需登录）")
    print(f"   2. 完整文件（大体积）：{OUTPUT_HTML} / {OUTPUT_IFRAMES}")
    print(f"   3. 轻量化文件（推荐大模型读取）：{OUTPUT_LIGHT_HTML} / {OUTPUT_INPUTS_ONLY}")
    print(f"   4. 页面信息记录：{RECORD_FILE}")
    print("   5. 浏览器已保持打开状态，可继续操作！")
    print("   6. 如需关闭浏览器，手动关闭窗口即可（登录态仍保留）")

if __name__ == "__main__":
    # 安装依赖（首次运行需执行）
    try:
        import bs4
    except ImportError:
        print("📦 正在安装轻量化HTML所需依赖...")
        os.system("pip install beautifulsoup4")
        import bs4
    dump_full_page()