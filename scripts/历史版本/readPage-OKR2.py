# enhanced_page_extractor.py
import asyncio
from playwright.async_api import async_playwright
import json
from typing import Dict, Any, List
from datetime import datetime

class EnhancedPageExtractor:
    def __init__(self, headless=False, debug=False):
        self.headless = headless
        self.debug = debug
        self.browser = None
        self.page = None
    
    async def start(self):
        """启动浏览器"""
        p = await async_playwright().start()
        self.browser = await p.chromium.launch(
            headless=self.headless,
            args=['--start-maximized']  # 启动时最大化窗口
        )
        context = await self.browser.new_context(viewport=None)  # 使用实际窗口大小
        self.page = await context.new_page()
        
        if self.debug:
            # 添加事件监听用于调试
            self.page.on("console", lambda msg: print(f"浏览器日志: {msg.text}"))
    
    async def navigate(self, url: str, wait_for_network=True):
        """访问目标页面，包含多种等待策略"""
        print(f"正在访问: {url}")
        
        try:
            # 设置超时
            await self.page.goto(url, timeout=30000, wait_until="domcontentloaded")
            
            if wait_for_network:
                # 等待网络空闲
                await self.page.wait_for_load_state("networkidle", timeout=10000)
            
            # 额外等待动态内容
            await asyncio.sleep(2)
            
            # 尝试等待常见的选择器出现
            try:
                await self.page.wait_for_selector(
                    "input, select, textarea, button, form", 
                    timeout=5000,
                    state="attached"
                )
            except:
                print("未找到表单元素，可能是页面结构特殊")
            
            print("页面加载完成")
            
        except Exception as e:
            print(f"页面加载异常: {e}")
            # 截图保存现场
            await self.page.screenshot(path="error_screenshot.png")
            print("已保存错误截图到 error_screenshot.png")
    
    async def extract_detailed_elements(self) -> Dict[str, Any]:
        """提取元素的详细定位信息"""
        
        extract_js = r"""
        () => {
            function getDetailedInfo(el) {
                if (!el) return null;
                
                // 基础信息
                const info = {
                    tag: el.tagName.toLowerCase(),
                    id: el.id || null,
                    name: el.name || null,
                    type: el.type || null,
                    value: el.value || null,
                    
                    // 文本信息
                    text: el.innerText ? el.innerText.trim().substring(0, 100) : null,
                    placeholder: el.placeholder || null,
                    title: el.title || null,
                    alt: el.alt || null,
                    aria_label: el.getAttribute('aria-label') || null,
                    aria_labelledby: el.getAttribute('aria-labelledby') || null,
                    
                    // 类名和样式
                    class_list: el.className ? el.className.split(/\s+/) : [],
                    
                    // 位置和尺寸
                    bounding_rect: el.getBoundingClientRect ? {
                        x: el.getBoundingClientRect().x,
                        y: el.getBoundingClientRect().y,
                        width: el.getBoundingClientRect().width,
                        height: el.getBoundingClientRect().height
                    } : null,
                    
                    // 状态
                    is_visible: !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length),
                    is_enabled: !el.disabled,
                    is_readonly: el.readOnly || false,
                    is_required: el.required || false,
                    
                    // 表单特有属性
                    form_id: el.form ? el.form.id : null,
                    form_name: el.form ? el.form.name : null,
                    
                    // 父子关系
                    parent_id: el.parentElement ? el.parentElement.id : null,
                    parent_class: el.parentElement ? el.parentElement.className : null,
                    
                    // 关联的label
                    label_text: null,
                    label_for: null,
                    
                    // 定位器
                    xpath: getXPath(el),
                    css_selectors: getAllCssSelectors(el)
                };
                
                // 查找关联的label
                if (el.id) {
                    const label = document.querySelector(`label[for="${el.id}"]`);
                    if (label) {
                        info.label_text = label.innerText?.trim() || null;
                        info.label_for = label.htmlFor;
                    }
                }
                
                // 如果被label包裹，获取label文本
                if (el.closest && el.closest('label')) {
                    const parentLabel = el.closest('label');
                    info.parent_label_text = parentLabel.innerText?.trim() || null;
                }
                
                return info;
            }
            
            function getXPath(element) {
                if (element.id) return `//*[@id="${element.id}"]`;
                
                const parts = [];
                while (element && element.nodeType === Node.ELEMENT_NODE) {
                    let index = 0;
                    const siblings = element.parentNode ? element.parentNode.children : [];
                    for (let i = 0; i < siblings.length; i++) {
                        const sibling = siblings[i];
                        if (sibling === element) {
                            parts.unshift(element.tagName.toLowerCase() + '[' + (index + 1) + ']');
                            break;
                        }
                        if (sibling.tagName === element.tagName) index++;
                    }
                    element = element.parentNode;
                }
                return '/' + parts.join('/');
            }
            
            function getAllCssSelectors(element) {
                const selectors = [];
                
                // 1. 最优先：ID选择器
                if (element.id) {
                    selectors.push(`#${element.id}`);
                }
                
                // 2. 名称选择器
                if (element.name) {
                    selectors.push(`[name="${element.name}"]`);
                }
                
                // 3. 类型+属性组合
                if (element.placeholder) {
                    selectors.push(`${element.tagName.toLowerCase()}[placeholder="${element.placeholder}"]`);
                }
                
                // 4. 类选择器组合
                if (element.className && typeof element.className === 'string') {
                    const classes = element.className.split(/\s+/).filter(c => c);
                    if (classes.length) {
                        selectors.push(`${element.tagName.toLowerCase()}.${classes.join('.')}`);
                    }
                }
                
                // 5. 通过文本内容（对按钮和链接特别有用）
                if (element.innerText && (element.tagName === 'BUTTON' || element.tagName === 'A')) {
                    const text = element.innerText.trim();
                    if (text) {
                        selectors.push(`${element.tagName.toLowerCase()}:has-text("${text}")`);
                    }
                }
                
                // 6. 通过关联label（如果存在）
                if (element.id) {
                    const label = document.querySelector(`label[for="${element.id}"]`);
                    if (label && label.innerText) {
                        selectors.push(`input[id="${element.id}"]`); // 已经包含在ID选择器中
                    }
                }
                
                return selectors;
            }
            
            // 收集所有可交互元素
            const elements = {
                all: [],
                by_tag: {},
                by_type: {},
                in_forms: []
            };
            
            // 常见可交互元素选择器
            const selectors = [
                'input', 'select', 'textarea', 'button',
                'a[href]', '[onclick]', '[role="button"]',
                '[role="textbox"]', '[role="checkbox"]', '[role="radio"]',
                '[contenteditable="true"]', 'label'
            ];
            
            selectors.forEach(selector => {
                document.querySelectorAll(selector).forEach(el => {
                    const info = getDetailedInfo(el);
                    elements.all.push(info);
                    
                    // 按标签分组
                    if (!elements.by_tag[info.tag]) {
                        elements.by_tag[info.tag] = [];
                    }
                    elements.by_tag[info.tag].push(info);
                    
                    // 按类型分组（对input特别有用）
                    if (info.type) {
                        if (!elements.by_type[info.type]) {
                            elements.by_type[info.type] = [];
                        }
                        elements.by_type[info.type].push(info);
                    }
                });
            });
            
            // 收集表单内的元素分组
            document.querySelectorAll('form').forEach(form => {
                const formInfo = {
                    form_id: form.id || null,
                    form_name: form.name || null,
                    form_action: form.action || null,
                    elements: []
                };
                
                form.querySelectorAll('input, select, textarea, button').forEach(el => {
                    formInfo.elements.push({
                        tag: el.tagName.toLowerCase(),
                        name: el.name || null,
                        id: el.id || null,
                        type: el.type || null,
                        placeholder: el.placeholder || null
                    });
                });
                
                elements.in_forms.push(formInfo);
            });
            
            return elements;
        }
        """
        
        return await self.page.evaluate(extract_js)
    
    async def get_page_fingerprint(self, url: str) -> Dict[str, Any]:
        """获取完整的页面指纹信息"""
        await self.start()
        await self.navigate(url)
        
        # 获取页面基本信息
        page_info = await self.page.evaluate("""
        () => ({
            title: document.title,
            url: window.location.href,
            charset: document.characterSet,
            doctype: document.doctype ? document.doctype.name : null,
            forms_count: document.forms.length,
            iframes_count: document.querySelectorAll('iframe').length,
            scripts_count: document.querySelectorAll('script').length,
            last_modified: document.lastModified
        })
        """)
        
        # 提取详细元素信息
        elements = await self.extract_detailed_elements()
        
        # 保存截图用于参考
        screenshot_path = f"page_screenshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        await self.page.screenshot(path=screenshot_path, full_page=True)
        print(f"已保存页面截图: {screenshot_path}")
        
        # 保存HTML源码
        html_path = f"page_source_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
        html_content = await self.page.content()
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html_content)
        print(f"已保存HTML源码: {html_path}")
        
        await self.browser.close()
        
        return {
            "page_info": page_info,
            "elements": elements,
            "total_elements": len(elements.get("all", [])),
            "screenshot": screenshot_path,
            "html_source": html_path,
            "extract_time": datetime.now().isoformat()
        }
    
    def generate_llm_prompt(self, data: Dict[str, Any]) -> str:
        """生成优化的大模型提示词"""
        
        lines = []
        lines.append("=" * 80)
        lines.append("📄 页面结构分析报告")
        lines.append("=" * 80)
        lines.append(f"页面标题: {data['page_info']['title']}")
        lines.append(f"页面URL: {data['page_info']['url']}")
        lines.append(f"提取时间: {data['extract_time']}")
        lines.append(f"总计可交互元素: {data['total_elements']}")
        lines.append("=" * 80)
        
        # 按类型统计
        lines.append("\n📊 元素统计:")
        for tag, items in data['elements']['by_tag'].items():
            lines.append(f"  - {tag}: {len(items)}个")
        
        # 表单分组信息
        if data['elements']['in_forms']:
            lines.append("\n📋 表单分组:")
            for i, form in enumerate(data['elements']['in_forms']):
                lines.append(f"\n  表单 {i+1}:")
                if form['form_id']:
                    lines.append(f"    ID: {form['form_id']}")
                if form['form_name']:
                    lines.append(f"    Name: {form['form_name']}")
                lines.append(f"    包含 {len(form['elements'])} 个元素:")
                for el in form['elements']:
                    lines.append(f"      - {el['tag']}: name='{el['name'] or ''}', type='{el['type'] or ''}'")
        
        # 详细的元素定位信息
        lines.append("\n" + "=" * 80)
        lines.append("🔍 详细元素定位信息")
        lines.append("=" * 80)
        
        # 按类型输出
        element_types = {
            'input_text': [e for e in data['elements']['all'] if e['tag'] == 'input' and e['type'] in ['text', 'search', 'email', 'tel', 'url', None]],
            'input_password': [e for e in data['elements']['all'] if e['tag'] == 'input' and e['type'] == 'password'],
            'input_checkbox': [e for e in data['elements']['all'] if e['tag'] == 'input' and e['type'] == 'checkbox'],
            'input_radio': [e for e in data['elements']['all'] if e['tag'] == 'input' and e['type'] == 'radio'],
            'select': [e for e in data['elements']['all'] if e['tag'] == 'select'],
            'textarea': [e for e in data['elements']['all'] if e['tag'] == 'textarea'],
            'button': [e for e in data['elements']['all'] if e['tag'] in ['button', 'a'] or (e['tag'] == 'input' and e['type'] in ['submit', 'button'])]
        }
        
        for type_name, items in element_types.items():
            if items:
                lines.append(f"\n{type_name.replace('_', ' ').title()}:")
                for idx, item in enumerate(items[:15]):  # 限制数量
                    lines.append(f"\n  [{idx+1}] {item['tag']}")
                    
                    # 所有可能的定位方式
                    if item['id']:
                        lines.append(f"      ID: {item['id']}")
                    if item['name']:
                        lines.append(f"      Name: {item['name']}")
                    if item['placeholder']:
                        lines.append(f"      Placeholder: {item['placeholder']}")
                    if item['label_text']:
                        lines.append(f"      Label: {item['label_text']}")
                    if item['parent_label_text']:
                        lines.append(f"      Parent Label: {item['parent_label_text']}")
                    if item['text'] and len(item['text']) < 50:
                        lines.append(f"      Text: {item['text']}")
                    
                    # CSS选择器建议
                    if item.get('css_selectors'):
                        lines.append(f"      CSS建议: {', '.join(item['css_selectors'][:3])}")
                    
                    if item['xpath']:
                        lines.append(f"      XPath: {item['xpath']}")
        
        # 添加使用建议
        lines.append("\n" + "=" * 80)
        lines.append("💡 使用建议")
        lines.append("=" * 80)
        lines.append("1. 优先使用ID定位，这是最稳定的方式")
        lines.append("2. 其次使用name属性，特别适合表单元素")
        lines.append("3. 使用placeholder或label文本作为备用方案")
        lines.append("4. 对于动态元素，使用CSS选择器的组合")
        lines.append("5. 可以结合多种定位方式，增加稳定性")
        
        return "\n".join(lines)

# 使用示例
async def main():
    # 创建提取器
    extractor = EnhancedPageExtractor(
        headless=False,  # 设为False可以看到浏览器操作
        debug=True       # 开启调试模式
    )
    
    # 提取页面信息
    url = input("请输入目标页面URL: ").strip()
    if not url:
        url = "https://okr.baidu-int.com/pages/weekly.html#/home?date=s2026-02-16&mandatorId=s&isEdit=true"  # 替换为你的内网地址
    
    print(f"\n开始提取页面结构: {url}")
    data = await extractor.get_page_fingerprint(url)
    
    # 生成大模型提示词
    prompt = extractor.generate_llm_prompt(data)
    
    # 保存到文件
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    with open(f"page_structure_{timestamp}.txt", "w", encoding="utf-8") as f:
        f.write(prompt)
    
    with open(f"page_data_{timestamp}.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str)
    
    print(f"\n✅ 提取完成！")
    print(f"📄 结构文件: page_structure_{timestamp}.txt")
    print(f"📊 数据文件: page_data_{timestamp}.json")
    print(f"📸 截图文件: {data['screenshot']}")
    print(f"📝 HTML源码: {data['html_source']}")
    
    # 显示摘要
    print(f"\n📈 摘要:")
    print(f"  共发现 {data['total_elements']} 个可交互元素")
    print(f"  包含 {data['page_info']['forms_count']} 个表单")
    print(f"  截图已保存，可以查看页面布局")

if __name__ == "__main__":
    asyncio.run(main())