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

# 导出文件
OUTPUT_HTML = "page_full_dump.html"
OUTPUT_IFRAMES = "iframes_dump.txt"

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

def dump_full_page():
    # 确保用户数据目录存在（关键：持久化登录态）
    if not os.path.exists(USER_DATA_DIR):
        os.makedirs(USER_DATA_DIR)
        print(f"📁 创建浏览器配置目录：{USER_DATA_DIR}（用于保存登录态）")

    # 核心修改：不使用 with 语句，手动控制 browser 生命周期，避免自动关闭
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
        # 等待页面稳定
        page.wait_for_load_state("domcontentloaded", timeout=30000)
        time.sleep(5)  # 给内网页面渲染时间

        # 1. 导出完整HTML
        full_html = page.content()
        with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
            f.write(full_html)
        print(f"✅ 完整页面已导出：{OUTPUT_HTML}")

        # 2. 导出iframe内容（已确认有1个核心iframe）
        iframes = page.frames
        iframe_count = len(iframes)
        print(f"🔍 找到 {iframe_count} 个 iframe/框架（核心填报内容在iframe 0中）")

        with open(OUTPUT_IFRAMES, "w", encoding="utf-8") as f:
            for i, frame in enumerate(iframes):
                try:
                    f.write(f"===== iframe {i} | url={frame.url} =====\n")
                    frame.wait_for_load_state("domcontentloaded", timeout=30000)
                    f.write(frame.content())
                    f.write("\n\n")
                except Exception as e:
                    err_msg = f"iframe {i} 读取失败: {str(e)}"
                    f.write(f"{err_msg}\n\n")
                    print(f"⚠️ {err_msg}（不影响主页面）")

        print(f"✅ 所有 iframe 内容已导出：{OUTPUT_IFRAMES}")

        # 3. 读取所有可输入框（包括iframe内的输入框）
        print("\n📝 页面上所有可输入框（含iframe内）：")
        # 先定位核心iframe（已确认索引为0）
        main_iframe = page.frames[0] if iframe_count > 0 else page
        
        # 读取iframe内所有输入框（适配你的多层目录输入框）
        inputs = main_iframe.locator("input, textarea, [contenteditable='true']").all()
        for idx, el in enumerate(inputs):
            try:
                placeholder = el.get_attribute("placeholder") or ""
                value = el.input_value() or ""
                # 获取输入框所在的目录文本（辅助定位层级）
                parent_text = ""
                try:
                    # 查找输入框上方最近的目录标题
                    parent = el.locator("xpath=ancestor::div[contains(@class, 'block-container-weekly')]//span[contains(@class, 'label-title')]").first
                    parent_text = parent.inner_text()[:30]  # 截取前30字符
                except:
                    parent_text = "未识别目录"

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

    # 核心修改：不关闭浏览器，仅提示脚本执行完毕
    print("\n🎉 页面导出完成！")
    print("📌 关键信息：")
    print(f"   1. 登录态保存在：{USER_DATA_DIR}（下次启动无需登录）")
    print(f"   2. 页面文件：{OUTPUT_HTML} / {OUTPUT_IFRAMES}")
    print(f"   3. 页面信息记录：{RECORD_FILE}")
    print("   4. 浏览器已保持打开状态，可继续操作！")
    print("   5. 如需关闭浏览器，手动关闭窗口即可（登录态仍保留）")

    # 不调用 browser.close() 和 p.stop()，保持浏览器打开
    # 手动关闭浏览器窗口后，登录态仍保存在 USER_DATA_DIR 中

if __name__ == "__main__":
    dump_full_page()