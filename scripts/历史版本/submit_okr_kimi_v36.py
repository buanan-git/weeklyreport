#!/usr/bin/env python3
"""百度OKR周报自动提交脚本 - V2.5 (完整修复版)"""

import asyncio
import socket
import json
import os
import time
import subprocess
from datetime import datetime, timedelta
from pathlib import Path
from playwright.async_api import async_playwright, Error as PlaywrightError
import platform
import traceback

# ==================== 工具函数 ====================

def iprint(msg):
    """智能打印，自动添加时间戳"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    print(f"[{timestamp}] {msg}")

# ==================== 浏览器管理器 ====================

BROWSER_STATE_FILE = "browser_state.json"
BROWSER_DEBUG_DIR = "browser_debug"

class BrowserManager:
    """浏览器管理器，支持复用已打开的浏览器"""
    
    def __init__(self):
        self.browser = None
        self.context = None
        self.page = None
        self.playwright = None
        self.cdp_port = None
        self.browser_pid = None
        self.browser_process = None
    
    def save_state(self):
        """保存浏览器状态到文件"""
        if self.cdp_port and self.browser_pid:
            state = {
                "cdp_port": self.cdp_port,
                "browser_pid": self.browser_pid,
                "timestamp": time.time()
            }
            with open(BROWSER_STATE_FILE, "w", encoding="utf-8") as f:
                json.dump(state, f)
            iprint(f"[浏览器] 状态已保存到 {BROWSER_STATE_FILE}")
    
    def load_state(self):
        """从文件加载浏览器状态"""
        if os.path.exists(BROWSER_STATE_FILE):
            try:
                with open(BROWSER_STATE_FILE, "r", encoding="utf-8") as f:
                    state = json.load(f)
                
                if self.is_process_running(state.get("browser_pid")):
                    self.cdp_port = state.get("cdp_port")
                    self.browser_pid = state.get("browser_pid")
                    iprint(f"[浏览器] 从文件加载状态成功: 端口={self.cdp_port}, PID={self.browser_pid}")
                    return True
                else:
                    iprint("[浏览器] 保存的浏览器进程已不存在")
            except Exception as e:
                iprint(f"[浏览器] 加载状态失败: {e}")
        return False
    
    def is_process_running(self, pid):
        """检查进程是否在运行"""
        if not pid:
            return False
        try:
            if platform.system() == "Windows":
                result = subprocess.run(f'tasklist /FI "PID eq {pid}"', shell=True, capture_output=True, text=True)
                return str(pid) in result.stdout
            else:
                os.kill(pid, 0)
                return True
        except:
            return False
    
    async def connect_to_existing_browser(self):
        """连接到已存在的浏览器"""
        if not self.cdp_port:
            return False
        
        try:
            self.playwright = await async_playwright().start()
            self.browser = await self.playwright.chromium.connect_over_cdp(
                f"http://127.0.0.1:{self.cdp_port}"
            )
            
            contexts = self.browser.contexts
            if contexts:
                self.context = contexts[0]
                pages = self.context.pages
                if pages:
                    self.page = pages[0]
                    iprint("[浏览器] 成功连接到已存在的浏览器")
                    return True
                else:
                    self.page = await self.context.new_page()
                    return True
            
            iprint("[浏览器] 未找到现有的context")
            return False
        except Exception as e:
            iprint(f"[浏览器] 连接失败: {e}")
            return False
    
    async def launch_new_browser(self, headless=False):
        """启动新的浏览器"""
        if not CHROME_PATH:
            iprint("[X] 未找到Chrome浏览器，请手动安装或配置CHROME_PATH")
            return False
        
        os.makedirs(BROWSER_DEBUG_DIR, exist_ok=True)
        
        import random
        self.cdp_port = random.randint(9222, 9299)
        
        iprint(f"[浏览器] 启动新浏览器，调试端口: {self.cdp_port}")
        iprint(f"[浏览器] Chrome路径: {CHROME_PATH}")
        
        cmd = [
            CHROME_PATH,
            f"--remote-debugging-port={self.cdp_port}",
            f"--user-data-dir={os.path.abspath(BROWSER_DEBUG_DIR)}",
            "--no-first-run",
            "--no-default-browser-check",
        ]

        if headless:
            cmd.append("--headless=new")

        cmd.append(EDIT_BASE_URL)
        
        try:
            if platform.system() == "Windows":
                self.browser_process = subprocess.Popen(
                    cmd,
                    shell=False,
                    creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0
                )
            else:
                self.browser_process = subprocess.Popen(cmd)
            
            self.browser_pid = self.browser_process.pid
            iprint(f"[浏览器] 浏览器进程PID: {self.browser_pid}")
            
            iprint("[浏览器] 等待浏览器启动...")
            await asyncio.sleep(5)
            
            self.playwright = await async_playwright().start()
            
            max_retries = 5
            for i in range(max_retries):
                try:
                    self.browser = await self.playwright.chromium.connect_over_cdp(
                        f"http://127.0.0.1:{self.cdp_port}"
                    )
                    break
                except Exception as e:
                    if i < max_retries - 1:
                        iprint(f"[浏览器] 连接尝试 {i+1}/{max_retries} 失败，等待重试...")
                        await asyncio.sleep(2)
                    else:
                        raise e
            
            self.context = self.browser.contexts[0] if self.browser.contexts else await self.browser.new_context()
            self.page = self.context.pages[0] if self.context.pages else await self.context.new_page()
            
            self.save_state()
            iprint("[浏览器] 新浏览器启动并连接成功")
            return True
        except Exception as e:
            iprint(f"[浏览器] 连接新浏览器失败: {e}")
            return False
    
    async def ensure_page(self):
        """确保有可用的页面"""
        if not self.browser:
            return False
        
        self.context = self.browser.contexts[0] if self.browser.contexts else None
        if not self.context:
            self.context = await self.browser.new_context()
        
        self.page = self.context.pages[0] if self.context.pages else None
        if not self.page:
            self.page = await self.context.new_page()
        
        return True
    
    async def close(self, keep_browser=True):
        """关闭连接，可选择是否保留浏览器"""
        if self.playwright:
            if keep_browser:
                await self.playwright.stop()
                iprint("[浏览器] 已断开连接，浏览器进程继续运行")
            else:
                if self.browser:
                    await self.browser.close()
                await self.playwright.stop()
                if self.browser_process:
                    self.browser_process.terminate()
                if os.path.exists(BROWSER_STATE_FILE):
                    os.remove(BROWSER_STATE_FILE)
                iprint("[浏览器] 浏览器已关闭")

# ==================== 配置区域 ====================

# Chrome浏览器路径
if platform.system() == "Windows":
    CHROME_PATHS = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        r"C:\Users\%USERNAME%\AppData\Local\Google\Chrome\Application\chrome.exe"
    ]
    CHROME_PATH = None
    for path in CHROME_PATHS:
        expanded_path = os.path.expandvars(path)
        if os.path.exists(expanded_path):
            CHROME_PATH = expanded_path
            break
else:
    CHROME_PATH = "google-chrome"

MY_ID = "s673090"
EDIT_BASE_URL = "https://okr.baidu-int.com/pages/weekly.html#/home"

DRY_RUN = False
AUTO_SUBMIT = False  # 建议首次运行设为False，确认无误后再设为True
JSON_FILE_PATH = "team_weekly_report_v10.json"

WEEKLY_DATA = {
    "part1_work_summary": None,
    "part2_okr_structure": None,
    "part3_metrics": None
}

# ==================== 核心修复函数 ====================

async def get_working_context_fixed(page):
    """修复版：正确穿透进入百度周报的核心 iframe"""
    try:
        # 等待 iframe 加载
        await page.wait_for_selector('iframe#main', timeout=10000)
        
        # 尝试通过 name 获取
        frame = page.frame('main')
        
        # 如果失败，尝试通过 URL 匹配
        if not frame:
            for f in page.frames:
                if 'weekly.html' in f.url:
                    frame = f
                    break
        
        if frame:
            iprint("[OK] 已进入 iframe#main 上下文")
            return frame
        
        iprint("[!] 找到 iframe 但无法获取 frame 对象，回退到主页面")
    except Exception as e:
        iprint(f"[!] iframe 穿透失败: {e}")
    
    return page


async def fill_react_editor(locator, content):
    """通用：填写React富文本编辑器（增强版）"""
    try:
        if await locator.count() == 0:
            return False
        
        # 确保元素可见
        if not await locator.is_visible():
            await locator.scroll_into_view_if_needed()
            await asyncio.sleep(0.3)
        
        # 聚焦并点击激活
        await locator.focus()
        await asyncio.sleep(0.2)
        await locator.click()
        await asyncio.sleep(0.3)
        
        # 使用 JavaScript 强制设置内容（最可靠的方式）
        success = await locator.evaluate("""(el, content) => {
            try {
                // 滚动到可视区域
                el.scrollIntoView({behavior: 'smooth', block: 'center'});
                
                // 聚焦
                el.focus();
                
                // 清空现有内容
                if (el.isContentEditable) {
                    el.innerHTML = '';
                } else {
                    el.value = '';
                }
                
                // 设置新内容
                const lines = content.split('\\n');
                
                if (el.isContentEditable) {
                    // 富文本编辑器：使用 <br> 分隔
                    el.innerHTML = lines.map((line, i) => {
                        if (i === lines.length - 1) return line;
                        return line + '<br>';
                    }).join('');
                } else {
                    // 原生输入框
                    const nativeSetter = Object.getOwnPropertyDescriptor(
                        window.HTMLTextAreaElement.prototype || window.HTMLInputElement.prototype, 
                        "value"
                    )?.set;
                    
                    if (nativeSetter) {
                        nativeSetter.call(el, content);
                    } else {
                        el.value = content;
                    }
                }
                
                // 触发完整事件序列，确保 React 能捕获
                const eventSequence = [
                    new Event('focus', {bubbles: true}),
                    new InputEvent('beforeinput', {bubbles: true, data: content}),
                    new InputEvent('input', {bubbles: true, data: content}),
                    new Event('change', {bubbles: true}),
                    new KeyboardEvent('keydown', {key: 'Enter', bubbles: true}),
                    new KeyboardEvent('keyup', {key: 'Enter', bubbles: true}),
                    new Event('blur', {bubbles: true})
                ];
                
                eventSequence.forEach(event => {
                    el.dispatchEvent(event);
                });
                
                // 额外触发 composition 事件（某些编辑器需要）
                if (el.isContentEditable) {
                    el.dispatchEvent(new CompositionEvent('compositionstart', {bubbles: true}));
                    el.dispatchEvent(new CompositionEvent('compositionend', {bubbles: true}));
                }
                
                return true;
            } catch (err) {
                console.error('Fill error:', err);
                return false;
            }
        }""", content)
        
        await asyncio.sleep(0.5)
        return success
        
    except Exception as e:
        iprint(f"    [!] 填写编辑器失败: {e}")
        return False


async def fill_part1_work_summary_fixed(page, part1_data):
    """修复版第一部分：精准定位"本周工作总结" """
    iprint("\n[填写] 第一部分：本周工作总结")
    
    content = part1_data.get("content", "")
    if not content:
        iprint("  [!] 内容为空，跳过")
        return False
    
    # 前后加换行隔离，避免与原有内容粘连
    content = "\n\n" + content + "\n\n"
    
    # 获取 iframe 上下文
    ctx = await get_working_context_fixed(page)
    
    # ===== 策略1：精准匹配"本周工作总结"标签 =====
    try:
        # 精确匹配
        label_locator = ctx.locator("text=本周工作总结").first
        
        if await label_locator.count() == 0:
            # 尝试部分匹配
            label_locator = ctx.locator("text=/本周工作/").first
        
        if await label_locator.count() > 0:
            iprint("  [OK] 找到'本周工作总结'标签")
            
            await label_locator.scroll_into_view_if_needed()
            await asyncio.sleep(0.5)
            
            # 尝试多种选择器查找输入框
            input_selectors = [
                # 直接跟随的可编辑区域
                'xpath=following::div[@contenteditable="true"][1]',
                'xpath=following::textarea[1]',
                # 同一父级内的编辑器
                'xpath=ancestor::div[contains(@class, "section") or contains(@class, "block")][1]//div[@contenteditable="true"]',
                # 通过类名匹配百度OKR的编辑器
                'xpath=following::*[contains(@class, "slate") or contains(@class, "editor")][1]',
                # 更宽松的匹配
                'xpath=ancestor::div[contains(@class, "weekly")][1]//div[@contenteditable="true"]'
            ]
            
            for sel in input_selectors:
                try:
                    input_locator = label_locator.locator(sel).first
                    if await input_locator.count() > 0 and await input_locator.is_visible():
                        if await fill_react_editor(input_locator, content):
                            iprint(f"  [OK] 使用策略1填写成功")
                            return True
                except:
                    continue
            
            iprint("  [!] 策略1未找到输入框，尝试策略2")
    except Exception as e:
        iprint(f"  [!] 策略1出错: {e}")
    
    # ===== 策略2：全局搜索"本周工作总结"区域 =====
    try:
        iprint("  [策略2] 全局搜索工作区域...")
        
        # 使用 JavaScript 全局查找
        result = await ctx.evaluate("""(content) => {
            // 1. 查找所有包含"本周工作"文本的元素
            const allElements = Array.from(document.querySelectorAll('*'));
            const labelEl = allElements.find(el => {
                const text = el.textContent || '';
                return text.trim() === '本周工作总结' || 
                       text.includes('本周工作总结') ||
                       text.trim() === '工作总结';
            });
            
            if (!labelEl) {
                return {status: 'label_not_found'};
            }
            
            // 2. 从标签出发查找输入框
            let inputEl = null;
            
            // 方法A：查找后续兄弟元素
            let sibling = labelEl.nextElementSibling;
            while (sibling) {
                if (sibling.getAttribute('contenteditable') === 'true' ||
                    sibling.tagName === 'TEXTAREA') {
                    inputEl = sibling;
                    break;
                }
                const inner = sibling.querySelector('[contenteditable="true"], textarea');
                if (inner) {
                    inputEl = inner;
                    break;
                }
                sibling = sibling.nextElementSibling;
            }
            
            // 方法B：向上查找父级容器
            if (!inputEl) {
                let parent = labelEl.parentElement;
                for (let i = 0; i < 5 && parent; i++) {
                    const editables = parent.querySelectorAll('[contenteditable="true"], textarea');
                    if (editables.length > 0) {
                        // 找最近的一个
                        inputEl = Array.from(editables).sort((a, b) => {
                            const distA = Math.abs(a.getBoundingClientRect().top - labelEl.getBoundingClientRect().top);
                            const distB = Math.abs(b.getBoundingClientRect().top - labelEl.getBoundingClientRect().top);
                            return distA - distB;
                        })[0];
                        break;
                    }
                    parent = parent.parentElement;
                }
            }
            
            if (!inputEl) {
                return {status: 'input_not_found'};
            }
            
            // 3. 填写内容
            inputEl.scrollIntoView({behavior: 'smooth', block: 'center'});
            inputEl.focus();
            
            const lines = content.split('\\n');
            if (inputEl.isContentEditable) {
                inputEl.innerHTML = lines.join('<br>');
            } else {
                const nativeSetter = Object.getOwnPropertyDescriptor(
                    window.HTMLTextAreaElement.prototype, "value"
                )?.set;
                if (nativeSetter) {
                    nativeSetter.call(inputEl, content);
                } else {
                    inputEl.value = content;
                }
            }
            
            // 触发事件
            ['focus', 'input', 'change', 'blur'].forEach(evt => {
                inputEl.dispatchEvent(new Event(evt, {bubbles: true}));
            });
            
            return {status: 'filled', tagName: inputEl.tagName};
        }""", content)
        
        if result.get('status') == 'filled':
            iprint(f"  [OK] 策略2填写成功 (元素: {result.get('tagName')})")
            return True
        else:
            iprint(f"  [!] 策略2失败: {result.get('status')}")
            
    except Exception as e:
        iprint(f"  [!] 策略2出错: {e}")
    
    # ===== 策略3：创建新区域（兜底） =====
    try:
        iprint("  [策略3] 创建新的工作总结区域...")
        
        result = await ctx.evaluate("""({title, content}) => {
            // 检查是否已存在
            const existing = document.querySelector('.work-summary-section, [data-section="work-summary"]');
            if (existing) {
                // 已存在，直接填写
                const inputEl = existing.querySelector('[contenteditable="true"], textarea');
                if (inputEl) {
                    if (inputEl.isContentEditable) {
                        inputEl.innerHTML = content.replace(/\\n/g, '<br>');
                    } else {
                        inputEl.value = content;
                    }
                    ['focus', 'input', 'change', 'blur'].forEach(evt => {
                        inputEl.dispatchEvent(new Event(evt, {bubbles: true}));
                    });
                    return 'filled_existing';
                }
            }
            
            // 创建新区域
            const section = document.createElement('div');
            section.className = 'weekly-section work-summary-section';
            section.setAttribute('data-section', 'work-summary');
            section.style.cssText = 'margin: 20px 0; padding: 16px; background: #fafafa; border: 1px solid #e8ebf0; border-radius: 4px;';
            
            const h3 = document.createElement('h3');
            h3.textContent = title;
            h3.style.cssText = 'margin: 0 0 12px 0; font-size: 16px; font-weight: 500; color: #11141a;';
            
            const editorDiv = document.createElement('div');
            editorDiv.contentEditable = 'true';
            editorDiv.className = 'slate-editor';
            editorDiv.style.cssText = 'min-height: 150px; padding: 12px; background: white; border: 1px solid #dfe4ed; border-radius: 4px; outline: none; font-size: 14px; line-height: 1.6;';
            editorDiv.innerHTML = content.replace(/\\n/g, '<br>');
            
            section.appendChild(h3);
            section.appendChild(editorDiv);
            
            // 尝试插入到合适位置（OKR部分之前）
            const okrSection = allElements.find(el => 
                el.textContent && (el.textContent.includes('OKR') || el.textContent.includes('卜安安'))
            );
            
            if (okrSection && okrSection.parentElement) {
                okrSection.parentElement.insertBefore(section, okrSection);
            } else {
                // 添加到主容器
                const container = document.querySelector('.weekly-main, .doc-editor-wrapper, .block-container-weekly, form, main');
                if (container) {
                    container.appendChild(section);
                } else {
                    document.body.appendChild(section);
                }
            }
            
            // 触发事件
            editorDiv.focus();
            ['focus', 'input', 'change', 'blur'].forEach(evt => {
                editorDiv.dispatchEvent(new Event(evt, {bubbles: true}));
            });
            
            return 'created';
        }""", {"title": "本周工作总结", "content": content})
        
        if result in ['created', 'filled_existing']:
            iprint(f"  [OK] 策略3成功: {result}")
            return True
            
    except Exception as e:
        iprint(f"  [!] 策略3失败: {e}")
    
    return False


async def fill_part2_okr_fixed(page, part2_data):
    """修复版第二部分：精准匹配OKR表格结构（卜安安的OKR）"""
    iprint("\n[填写] 第二部分：OKR进度")
    
    objectives = part2_data.get("objectives", [])
    if not objectives:
        iprint("  [!] 无OKR数据")
        return False
    
    ctx = await get_working_context_fixed(page)
    success_count = 0
    
    # ===== 第一步：定位OKR表格区域 =====
    try:
        # 关键：找到"卜安安的OKR"或"OKR进度"
        okr_header = ctx.locator("text=/卜安安.*OKR/i").first
        
        if await okr_header.count() == 0:
            okr_header = ctx.locator("text=OKR进度").first
        
        if await okr_header.count() > 0:
            iprint("  [OK] 找到OKR表格区域")
            await okr_header.scroll_into_view_if_needed()
            await asyncio.sleep(1)
        else:
            iprint("  [!] 未找到OKR表格区域，将尝试直接查找KR")
    except Exception as e:
        iprint(f"  [!] 定位OKR区域失败: {e}")
    
    # ===== 第二步：处理每个KR的"本周工作" =====
    for obj_idx, obj in enumerate(objectives):
        obj_id = obj.get("o_id", f"O{obj_idx+1}")
        obj_title = obj.get("o_title", "")
        
        iprint(f"\n  [Objective] {obj_id}: {obj_title[:40]}...")
        
        key_results = obj.get("key_results", [])
        
        for kr_idx, kr in enumerate(key_results):
            kr_id = kr.get("kr_id", f"KR{kr_idx+1}")
            kr_title = kr.get("kr_title", "")
            weekly_work = kr.get("weekly_work", "")
            
            if not weekly_work:
                continue
            
            iprint(f"    [KR] {kr_id}: {kr_title[:30] if kr_title else '无标题'}...")
            
            # 尝试填写这个KR的本周工作
            filled = await fill_kr_weekly_work(ctx, kr_id, weekly_work)
            
            if filled:
                success_count += 1
                iprint(f"      [OK] 填写成功")
            else:
                iprint(f"      [!] 填写失败")
    
    iprint(f"\n  [统计] 成功填写 {success_count} 个KR")
    return success_count > 0


async def fill_kr_weekly_work(ctx, kr_id, content):
    """填写指定KR的"本周工作"输入框"""
    
    # 前后加换行隔离
    content = "\n" + content + "\n"
    
    try:
        # ===== 方法1：通过KR ID直接定位 =====
        kr_locator = ctx.locator(f"text={kr_id}").first
        
        if await kr_locator.count() == 0:
            # 尝试模糊匹配
            kr_locator = ctx.locator(f"text=/{kr_id}/i").first
        
        if await kr_locator.count() > 0:
            await kr_locator.scroll_into_view_if_needed()
            await asyncio.sleep(0.3)
            
            # 从KR元素出发，查找"本周工作"输入框
            result = await ctx.evaluate("""({krId, content}) => {
                console.log('JS: 查找KR:', krId);
                
                // 查找KR元素
                const allElements = Array.from(document.querySelectorAll('*'));
                const krEl = allElements.find(el => {
                    const text = el.textContent || '';
                    // 精确匹配 KR1, KR2 等，避免匹配到包含这些文本的其他元素
                    const match = text.trim().match(new RegExp('^' + krId + '\\\\s*$')) ||
                                  text.trim().match(new RegExp('^' + krId + '\\\\s')) ||
                                  text.includes(krId + ' ');
                    return match;
                });
                
                if (!krEl) {
                    console.log('JS: 未找到KR元素:', krId);
                    return {status: 'kr_not_found'};
                }
                
                console.log('JS: 找到KR元素:', krEl.tagName, krEl.textContent.substring(0, 50));
                
                // 查找"本周工作"输入框
                let inputEl = null;
                
                // 策略A：在KR的父级行/容器中查找
                let row = krEl.closest('tr, [class*="row"], [class*="item"], div');
                if (row) {
                    // 查找包含"本周工作"文本的元素
                    const weeklyWorkLabels = Array.from(row.querySelectorAll('*')).filter(el => {
                        const text = el.textContent || '';
                        return text.includes('本周工作') || text.includes('本周进展');
                    });
                    
                    for (let label of weeklyWorkLabels) {
                        // 查找这个标签后面的可编辑元素
                        let nextEl = label.nextElementSibling;
                        while (nextEl) {
                            if (nextEl.getAttribute('contenteditable') === 'true' ||
                                nextEl.tagName === 'TEXTAREA' ||
                                nextEl.tagName === 'INPUT') {
                                inputEl = nextEl;
                                break;
                            }
                            const inner = nextEl.querySelector('[contenteditable="true"], textarea, input:not([type="hidden"])');
                            if (inner) {
                                inputEl = inner;
                                break;
                            }
                            nextEl = nextEl.nextElementSibling;
                        }
                        
                        if (inputEl) break;
                    }
                    
                    // 如果没找到，查找行内所有可编辑元素，排除标题等
                    if (!inputEl) {
                        const editables = row.querySelectorAll('[contenteditable="true"], textarea');
                        // 找高度较大的（输入框通常比标签高）
                        for (let ed of editables) {
                            const rect = ed.getBoundingClientRect();
                            if (rect.height > 30) {  // 假设输入框高度大于30px
                                inputEl = ed;
                                break;
                            }
                        }
                    }
                }
                
                // 策略B：从KR元素向后查找
                if (!inputEl) {
                    let nextEl = krEl.nextElementSibling;
                    let depth = 0;
                    while (nextEl && depth < 10) {
                        if (nextEl.getAttribute('contenteditable') === 'true' ||
                            nextEl.tagName === 'TEXTAREA') {
                            inputEl = nextEl;
                            break;
                        }
                        
                        // 检查是否包含"本周工作"
                        if (nextEl.textContent && nextEl.textContent.includes('本周')) {
                            const inner = nextEl.querySelector('[contenteditable="true"], textarea');
                            if (inner) {
                                inputEl = inner;
                                break;
                            }
                        }
                        
                        nextEl = nextEl.nextElementSibling;
                        depth++;
                    }
                }
                
                // 策略C：向上查找祖先，然后在祖先内搜索
                if (!inputEl) {
                    let ancestor = krEl.parentElement;
                    for (let i = 0; i < 3 && ancestor; i++) {
                        const weeklyWorkEl = Array.from(ancestor.querySelectorAll('*')).find(el => {
                            const text = el.textContent || '';
                            return text.includes('本周工作') && el !== krEl;
                        });
                        
                        if (weeklyWorkEl) {
                            // 找到"本周工作"元素后，找其相邻的可编辑元素
                            let sibling = weeklyWorkEl.nextElementSibling;
                            if (sibling && sibling.getAttribute('contenteditable') === 'true') {
                                inputEl = sibling;
                                break;
                            }
                            
                            // 或者在同一父级内查找
                            const parent = weeklyWorkEl.parentElement;
                            if (parent) {
                                const editables = parent.querySelectorAll('[contenteditable="true"]');
                                if (editables.length > 0) {
                                    inputEl = editables[0];
                                    break;
                                }
                            }
                        }
                        
                        ancestor = ancestor.parentElement;
                    }
                }
                
                if (!inputEl) {
                    console.log('JS: 未找到输入框 for', krId);
                    return {status: 'input_not_found'};
                }
                
                console.log('JS: 找到输入框:', inputEl.tagName, inputEl.className);
                
                // 填写内容
                inputEl.scrollIntoView({behavior: 'smooth', block: 'center'});
                inputEl.focus();
                
                const lines = content.split('\\n');
                
                if (inputEl.isContentEditable) {
                    inputEl.innerHTML = lines.join('<br>');
                } else {
                    const nativeSetter = Object.getOwnPropertyDescriptor(
                        window.HTMLTextAreaElement.prototype, "value"
                    )?.set;
                    if (nativeSetter) {
                        nativeSetter.call(inputEl, content);
                    } else {
                        inputEl.value = content;
                    }
                }
                
                // 触发完整事件序列
                const events = [
                    new Event('focus', {bubbles: true}),
                    new InputEvent('beforeinput', {bubbles: true, data: content}),
                    new InputEvent('input', {bubbles: true, data: content}),
                    new Event('change', {bubbles: true}),
                    new KeyboardEvent('keydown', {key: 'Enter', bubbles: true}),
                    new KeyboardEvent('keyup', {key: 'Enter', bubbles: true}),
                    new Event('blur', {bubbles: true})
                ];
                
                events.forEach(e => inputEl.dispatchEvent(e));
                
                console.log('JS: 已填写', krId);
                return {status: 'filled'};
            }""", {"krId": kr_id, "content": content})
            
            if result.get('status') == 'filled':
                return True
            
            iprint(f"      [!] JS方法失败: {result.get('status')}")
        
        # ===== 方法2：如果方法1失败，尝试 Playwright 原生操作 =====
        # 重新定位KR，然后尝试相对选择器
        if await kr_locator.count() > 0:
            # 尝试在同一父级内查找可编辑元素
            parent = kr_locator.locator('xpath=ancestor::div[contains(@class, "row") or contains(@class, "item") or contains(@class, "kr")][1]').first
            
            if await parent.count() > 0:
                editor = parent.locator('[contenteditable="true"], textarea').first
                if await editor.count() > 0:
                    if await fill_react_editor(editor, content):
                        return True
        
    except Exception as e:
        iprint(f"      [!] 填写KR出错: {e}")
    
    return False


async def fill_part3_metrics_fixed(page, part3_data):
    """修复版第三部分：关键指标"""
    iprint("\n[填写] 第三部分：关键指标")
    
    sections = part3_data.get("sections", {})
    if not sections:
        iprint("  [!] 无指标数据")
        return False
    
    ctx = await get_working_context_fixed(page)
    success_count = 0
    
    for section_name, section_data in sections.items():
        content = section_data.get("formatted_content", "")
        if not content:
            continue
        
        iprint(f"\n  [Section] {section_name}")
        
        # 查找这个section的标签
        try:
            label = ctx.locator(f"text={section_name}").first
            
            if await label.count() > 0:
                await label.scroll_into_view_if_needed()
                
                # 查找后续输入框
                input_selectors = [
                    'xpath=following::div[@contenteditable="true"][1]',
                    'xpath=following::textarea[1]',
                    'xpath=ancestor::div[contains(@class, "section")][1]//div[@contenteditable="true"]'
                ]
                
                for sel in input_selectors:
                    try:
                        inp = label.locator(sel).first
                        if await inp.count() > 0:
                            if await fill_react_editor(inp, content):
                                iprint(f"    [OK] 填写成功")
                                success_count += 1
                                break
                    except:
                        continue
            else:
                iprint(f"    [!] 未找到标签: {section_name}")
                
        except Exception as e:
            iprint(f"    [!] 出错: {e}")
    
    iprint(f"\n  [统计] 成功填写 {success_count}/{len(sections)} 个section")
    return success_count > 0


# ==================== JSON解析函数（保持原有逻辑） ====================

def flatten_nested_content(data, depth=0, max_depth=10):
    """递归扁平化嵌套结构"""
    if depth > max_depth:
        return [str(data)]
    
    if isinstance(data, list):
        result = []
        for item in data:
            result.extend(flatten_nested_content(item, depth + 1, max_depth))
        return result
    elif isinstance(data, dict):
        if "content" in data:
            return flatten_nested_content(data["content"], depth + 1, max_depth)
        elif "items" in data:
            return flatten_nested_content(data["items"], depth + 1, max_depth)
        elif "data" in data:
            return flatten_nested_content(data["data"], depth + 1, max_depth)
        else:
            return [data]
    else:
        return [data] if data else []


def parse_part1(data):
    """解析第一部分：本周工作总结"""
    iprint(f"\n  [解析] 第一部分: 本周工作总结")
    
    if isinstance(data, list):
        content_list = data
        title = "本周工作总结"
    elif isinstance(data, dict):
        title = data.get("title", "本周工作总结")
        content_list = data.get("content", [])
        if not isinstance(content_list, list()):
            for key in ["items", "data", "work_items", "summary"]:
                if key in data and isinstance(data[key], list):
                    content_list = data[key]
                    break
    else:
        iprint(f"    [WARN] 不支持的数据类型: {type(data)}")
        return None
    
    content_list = flatten_nested_content(content_list)
    
    if not content_list:
        iprint(f"    [WARN] 内容为空")
        return None
    
    formatted_items = []
    for idx, item in enumerate(content_list):
        if isinstance(item, dict):
            text = ", ".join([f"{k}={v}" for k, v in item.items()])
            formatted_items.append(f"{idx+1}. {text}")
        elif isinstance(item, str):
            formatted_items.append(f"{idx+1}. {item}")
        else:
            formatted_items.append(f"{idx+1}. {str(item)}")
    
    return {
        "title": title,
        "content": "\n".join(formatted_items),
        "raw_content": content_list,
        "item_count": len(formatted_items)
    }


def parse_part2(data):
    """解析第二部分：OKR进度"""
    iprint(f"\n  [解析] 第二部分: OKR进度")
    
    if isinstance(data, list):
        objectives = data
        title = "OKR进度"
    elif isinstance(data, dict):
        title = data.get("title", "OKR进度")
        objectives = data.get("objectives", [])
        if not isinstance(objectives, list):
            for key in ["items", "data", "okrs", "goals"]:
                if key in data and isinstance(data[key], list):
                    objectives = data[key]
                    break
    else:
        iprint(f"    [WARN] 不支持的数据类型: {type(data)}")
        return None
    
    if not isinstance(objectives, list):
        iprint(f"    [WARN] objectives 不是列表")
        return None
    
    okr_data = {"title": title, "objectives": []}
    
    for obj_idx, obj in enumerate(objectives):
        if not isinstance(obj, dict):
            continue
        
        o_id = obj.get("o_id", f"O{obj_idx+1}")
        o_title = obj.get("o_title", obj.get("title", ""))
        
        o_data = {"o_id": o_id, "o_title": o_title, "key_results": []}
        key_results = obj.get("key_results", [])
        
        if not isinstance(key_results, list):
            for key in ["krs", "items", "results"]:
                if key in obj and isinstance(obj[key], list):
                    key_results = obj[key]
                    break
        
        for kr_idx, kr in enumerate(key_results):
            if not isinstance(kr, dict):
                continue
            
            kr_id = kr.get("kr_id", f"KR{kr_idx+1}")
            kr_title = kr.get("kr_title", kr.get("title", ""))
            weekly_work = kr.get("weekly_work", "")
            progress = kr.get("progress", "")
            confidence = kr.get("confidence", "")
            
            # 处理weekly_work的嵌套结构
            if isinstance(weekly_work, dict):
                parts = []
                if "completed" in weekly_work:
                    parts.append(f"已完成: {weekly_work['completed']}")
                if "in_progress" in weekly_work:
                    parts.append(f"进行中: {weekly_work['in_progress']}")
                if "blocked" in weekly_work:
                    parts.append(f"阻塞: {weekly_work['blocked']}")
                formatted_work = "\n".join(parts) if parts else str(weekly_work)
            elif isinstance(weekly_work, list):
                flattened = flatten_nested_content(weekly_work)
                formatted_work = "\n".join([str(x) for x in flattened])
            else:
                formatted_work = str(weekly_work)
            
            if progress:
                formatted_work += f"\n进度: {progress}"
            if confidence:
                formatted_work += f" | 信心: {confidence}"
            
            o_data["key_results"].append({
                "kr_id": kr_id,
                "kr_title": kr_title,
                "weekly_work": formatted_work,
                "progress": progress,
                "confidence": confidence
            })
        
        okr_data["objectives"].append(o_data)
    
    total_krs = sum(len(o["key_results"]) for o in okr_data["objectives"])
    iprint(f"    [OK] 完成: {len(okr_data['objectives'])} 个O, {total_krs} 个KR")
    return okr_data


def parse_part3(data):
    """解析第三部分：关键指标"""
    iprint(f"\n  [解析] 第三部分: 关键指标")
    
    if isinstance(data, list):
        data = {"title": "关键指标", "sections": data}
    elif not isinstance(data, dict):
        iprint(f"    [WARN] 不支持的数据类型: {type(data)}")
        return None
    
    title = data.get("title", "关键指标")
    sections = data.get("sections", [])
    
    if isinstance(sections, dict):
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
    
    if not isinstance(sections, list):
        iprint(f"    [WARN] sections 不是列表")
        return None
    
    metrics_data = {"title": title, "sections": {}}
    
    for section in sections:
        if not isinstance(section, dict):
            continue
        
        section_title = section.get("section_title", "未命名")
        content_list = section.get("content", [])
        content_list = flatten_nested_content(content_list)
        
        formatted_items = []
        for idx, item in enumerate(content_list):
            if isinstance(item, dict):
                text = ", ".join([f"{k}={v}" for k, v in item.items()])
                formatted_items.append(f"{idx+1}. {text}")
            elif isinstance(item, str):
                formatted_items.append(f"{idx+1}. {item}")
            else:
                formatted_items.append(f"{idx+1}. {str(item)}")
        
        metrics_data["sections"][section_title] = {
            "raw_content": content_list,
            "formatted_content": "\n".join(formatted_items),
            "item_count": len(formatted_items)
        }
    
    iprint(f"    [OK] 完成: {len(metrics_data['sections'])} 个section")
    return metrics_data


def parse_json_file(json_path):
    """解析JSON文件"""
    iprint(f"\n{'='*60}")
    iprint("[DEBUG] 开始解析JSON文件")
    iprint(f"{'='*60}")
    
    if not os.path.exists(json_path):
        iprint("  [ERROR] 文件不存在")
        return None
    
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            raw_content = f.read()
        iprint(f"  [OK] 读取: {len(raw_content)} 字符")
    except Exception as e:
        iprint(f"  [ERROR] 读取失败: {e}")
        return None
    
    try:
        data = json.loads(raw_content)
        iprint(f"\n[DEBUG] JSON解析成功")
        iprint(f"  顶层键: {list(data.keys())}")
    except Exception as e:
        iprint(f"  [ERROR] JSON解析失败: {e}")
        return None
    
    # 检查是否在'content'键下
    if "content" in data:
        iprint(f"\n[DEBUG] 发现'content'键，使用嵌套结构")
        content_data = data.get("content", {})
        iprint(f"  content键下的键: {list(content_data.keys())}")
    else:
        iprint(f"\n[DEBUG] 使用顶层结构")
        content_data = data
    
    iprint(f"\n{'='*60}")
    iprint("开始解析各部分内容")
    iprint(f"{'='*60}")
    
    result = {
        "part1_work_summary": None,
        "part2_okr_structure": None,
        "part3_metrics": None
    }
    
    if "part1_work_summary" in content_data:
        iprint(f"\n  [OK] 找到 part1_work_summary")
        result["part1_work_summary"] = parse_part1(content_data["part1_work_summary"])
    
    if "part2_okr_structure" in content_data:
        iprint(f"\n  [OK] 找到 part2_okr_structure")
        result["part2_okr_structure"] = parse_part2(content_data["part2_okr_structure"])
    
    if "part3_metrics" in content_data:
        iprint(f"\n  [OK] 找到 part3_metrics")
        result["part3_metrics"] = parse_part3(content_data["part3_metrics"])
    
    iprint(f"\n{'='*60}")
    iprint("[DEBUG] 解析结果总结")
    iprint(f"{'='*60}")
    iprint(f"  第一部分: {'OK' if result['part1_work_summary'] else 'FAIL'}")
    iprint(f"  第二部分: {'OK' if result['part2_okr_structure'] else 'FAIL'}")
    iprint(f"  第三部分: {'OK' if result['part3_metrics'] else 'FAIL'}")
    
    return result


def load_content_from_json():
    """从JSON文件加载内容"""
    global WEEKLY_DATA
    
    iprint(f"\n{'='*60}")
    iprint("[系统] 加载JSON文件")
    iprint(f"{'='*60}")
    
    try:
        script_dir = Path(__file__).parent.absolute()
    except:
        script_dir = Path(os.getcwd())
    
    json_path = script_dir / JSON_FILE_PATH
    iprint(f"文件路径: {json_path}")
    
    parsed = parse_json_file(str(json_path))
    
    if parsed:
        WEEKLY_DATA["part1_work_summary"] = parsed.get("part1_work_summary")
        WEEKLY_DATA["part2_okr_structure"] = parsed.get("part2_okr_structure")
        WEEKLY_DATA["part3_metrics"] = parsed.get("part3_metrics")
        
        iprint(f"\n[OK] 数据加载成功")
        return True
    else:
        iprint(f"\n[ERROR] 加载失败")
        return False


def check_network_connection():
    """检查网络连接"""
    iprint("\n[网络检测] 检查网络连接...")
    try:
        ip = socket.gethostbyname("okr.baidu-int.com")
        iprint(f"  [OK] 内网解析正常: {ip}")
        return True
    except:
        iprint(f"  [X] 内网解析失败，请检查VPN")
        return False


def get_default_date():
    """获取默认日期（上周一）"""
    today = datetime.now()
    last_monday = today - timedelta(days=today.weekday() + 7)
    return f"s{last_monday.strftime('%Y-%m-%d')}"


# ==================== 页面辅助函数 ====================

async def analyze_page_structure(page):
    """分析页面结构"""
    iprint("\n[页面分析] 分析页面结构...")
    structure = await page.evaluate("""() => {
        const result = {has_part1: false, has_part2: false, has_part3: false, iframes: []};
        
        // 检查iframe
        document.querySelectorAll('iframe').forEach(f => {
            result.iframes.push({id: f.id, name: f.name, src: f.src});
        });
        
        // 检查关键文本
        document.querySelectorAll('h1, h2, h3, h4, [class*="title"], [class*="header"]').forEach(h => {
            const t = h.textContent || '';
            if (t.includes('本周工作总结') || t.includes('工作总结')) result.has_part1 = true;
            if (t.includes('OKR') || t.includes('okr')) result.has_part2 = true;
            if (t.includes('指标') || t.includes('卜安安')) result.has_part2 = true;
            if (t.includes('核心指标') || t.includes('项目')) result.has_part3 = true;
        });
        
        return result;
    }""")
    
    iprint(f"  iframe数量: {len(structure.get('iframes', []))}")
    for f in structure.get('iframes', []):
        iprint(f"    - id={f.get('id')}, name={f.get('name')}")
    iprint(f"  第一部分: {'存在' if structure['has_part1'] else '不存在'}")
    iprint(f"  第二部分: {'存在' if structure['has_part2'] else '不存在'}")
    iprint(f"  第三部分: {'存在' if structure['has_part3'] else '不存在'}")
    return structure


async def remove_overlay_elements(page):
    """移除遮挡层"""
    try:
        await page.evaluate("""() => {
            const selectors = [
                '.mp_ai_plugin_shade', '.MuiBackdrop-root', '.ant-modal-mask',
                '.erp-guide-mask', '.overlay', '.shade', '[class*="modal-mask"]'
            ];
            selectors.forEach(sel => {
                document.querySelectorAll(sel).forEach(el => {
                    el.style.display = 'none';
                    el.style.pointerEvents = 'none';
                    if (el.parentNode) el.parentNode.removeChild(el);
                });
            });
            return 'ok';
        }""")
        iprint("    [OK] 已清理遮挡层")
        return True
    except Exception as e:
        iprint(f"    [!] 清理遮挡层失败: {e}")
        return False


async def auto_submit_and_confirm(page):
    """自动提交并确认"""
    iprint("\n[提交] 自动提交...")
    
    ctx = await get_working_context_fixed(page)
    
    # 查找提交按钮
    submit_btn = ctx.locator("button:has-text('确定'), button:has-text('保存'), button:has-text('提交'), button[type='submit']").last
    
    if await submit_btn.count() > 0 and await submit_btn.is_visible():
        iprint("  [OK] 找到提交按钮")
        await remove_overlay_elements(page)
        await submit_btn.click()
        iprint("  [OK] 已点击提交")
        await asyncio.sleep(2)
        return True
    else:
        iprint("  [!] 未找到提交按钮")
        return False


# ==================== 主流程 ====================

async def submit_weekly_report(page, target_date=None, dry_run=True, auto_submit=False):
    """提交周报主流程"""
    target_date = target_date or get_default_date()
    edit_url = f"{EDIT_BASE_URL}?date={target_date}&mandatorId={MY_ID}&isEdit=true"
    
    iprint("="*60)
    iprint(f"目标日期: {target_date}")
    iprint(f"编辑模式: {'预览' if dry_run else '实际填写'}")
    iprint(f"自动提交: {'是' if auto_submit else '否'}")
    iprint("="*60)
    
    # 导航到页面
    try:
        await page.goto(edit_url, wait_until='domcontentloaded', timeout=60000)
        await asyncio.sleep(3)  # 等待React渲染
        
        if 'uuap' in page.url:
            return {'status': 'login_required', 'url': page.url}
        
        iprint("[OK] 页面加载完成")
    except Exception as e:
        iprint(f"\n[X] 页面加载失败: {e}")
        return {'status': 'error', 'error': str(e)}
    
    # 分析页面结构
    await analyze_page_structure(page)
    
    # 检查数据状态
    iprint("\n" + "="*60)
    iprint("[验证] 数据状态检查")
    iprint("="*60)
    iprint(f"  part1_work_summary: {WEEKLY_DATA.get('part1_work_summary') is not None}")
    iprint(f"  part2_okr_structure: {WEEKLY_DATA.get('part2_okr_structure') is not None}")
    iprint(f"  part3_metrics: {WEEKLY_DATA.get('part3_metrics') is not None}")
    
    if dry_run:
        iprint("\n" + "="*60)
        iprint("【预览模式】不实际填写")
        iprint("="*60)
        
        if WEEKLY_DATA.get("part1_work_summary"):
            iprint(f"\n1. 本周工作总结 ({len(WEEKLY_DATA['part1_work_summary']['content'])} 字符):")
            iprint(WEEKLY_DATA['part1_work_summary']['content'][:200] + "...")
        
        if WEEKLY_DATA.get("part2_okr_structure"):
            iprint(f"\n2. OKR进度:")
            for obj in WEEKLY_DATA["part2_okr_structure"]["objectives"]:
                iprint(f"  {obj['o_id']}: {obj['o_title'][:40]}...")
                for kr in obj["key_results"]:
                    iprint(f"    {kr['kr_id']}: {kr['weekly_work'][:50]}...")
        
        return {'status': 'preview', 'date': target_date}
    
    # 实际填写
    iprint("\n" + "="*60)
    iprint("开始填写内容")
    iprint("="*60)
    
    results = {}
    
    # 第一部分
    if WEEKLY_DATA.get("part1_work_summary"):
        results['part1'] = await fill_part1_work_summary_fixed(page, WEEKLY_DATA["part1_work_summary"])
    else:
        iprint("\n[!] 无第一部分数据，跳过")
        results['part1'] = False
    
    # 第二部分
    if WEEKLY_DATA.get("part2_okr_structure"):
        results['part2'] = await fill_part2_okr_fixed(page, WEEKLY_DATA["part2_okr_structure"])
    else:
        iprint("\n[!] 无第二部分数据，跳过")
        results['part2'] = False
    
    # 第三部分（可选）
    if WEEKLY_DATA.get("part3_metrics"):
        results['part3'] = await fill_part3_metrics_fixed(page, WEEKLY_DATA["part3_metrics"])
    else:
        iprint("\n[!] 无第三部分数据，跳过")
        results['part3'] = False
    
    # 统计结果
    iprint("\n" + "="*60)
    iprint("填写完成统计：")
    iprint(f"  本周工作总结: {'OK' if results.get('part1') else 'FAIL'}")
    iprint(f"  OKR本周工作: {'OK' if results.get('part2') else 'FAIL'}")
    iprint(f"  关键指标: {'OK' if results.get('part3') else 'FAIL'}")
    iprint("="*60)
    
    # 自动提交
    if auto_submit and (results.get('part1') or results.get('part2')):
        iprint("\n[自动提交] 准备提交...")
        submit_result = await auto_submit_and_confirm(page)
        results['submitted'] = submit_result
    
    return {'status': 'success', 'date': target_date, 'results': results}


async def main():
    """主函数"""
    iprint("\n" + "="*60)
    iprint("百度OKR周报自动提交脚本 - V2.5 (完整修复版)")
    iprint("="*60)
    iprint("\n【V2.5改进】")
    iprint("  ✓ 修复iframe穿透逻辑")
    iprint("  ✓ 三层策略填写工作总结（精准匹配→全局搜索→创建新区域）")
    iprint("  ✓ 精准匹配OKR表格结构（卜安安的OKR）")
    iprint("  ✓ 增强React组件兼容性（完整事件序列）")
    iprint("  ✓ 智能KR定位（通过KR ID精准查找本周工作输入框）")
    
    # 加载数据
    json_loaded = load_content_from_json()
    
    # 检查网络
    if not check_network_connection():
        response = input("\n网络检查失败，是否仍要继续? (y/N): ")
        if response.lower() != 'y':
            return
    
    # 初始化浏览器
    iprint(f"\n{'='*60}")
    iprint("[系统] 初始化浏览器...")
    
    browser_manager = BrowserManager()
    browser_connected = False
    
    if browser_manager.load_state():
        browser_connected = await browser_manager.connect_to_existing_browser()
    
    if not browser_connected:
        iprint("\n[浏览器] 启动新浏览器...")
        if not await browser_manager.launch_new_browser(headless=False):
            iprint("[X] 浏览器启动失败")
            return
    
    await browser_manager.ensure_page()
    page = browser_manager.page
    iprint("[OK] 浏览器已就绪")
    
    try:
        # 检查当前页面
        current_url = page.url
        base_url = "https://okr.baidu-int.com"
        
        if base_url not in current_url:
            iprint(f"\n[导航] 当前URL: {current_url}")
            iprint(f"[导航] 导航到OKR首页...")
            await page.goto(base_url, wait_until='domcontentloaded', timeout=30000)
            await asyncio.sleep(2)
        
        # 检查登录状态
        try:
            password_input = page.locator('input[type="password"]').first
            if await password_input.is_visible(timeout=5000):
                iprint("\n" + "="*60)
                iprint("[!] 检测到登录页面，请先手动登录")
                iprint("="*60)
                input("\n登录完成后按 Enter 继续...")
        except:
            pass
        
        if 'uuap' in page.url:
            iprint("\n[X] 未登录成功或会话已过期")
            await browser_manager.close(keep_browser=False)
            return
        
        iprint("[OK] 登录验证通过")
        
        # 执行填报
        target_date = get_default_date()
        result = await submit_weekly_report(
            page=page,
            target_date=target_date,
            dry_run=DRY_RUN,
            auto_submit=AUTO_SUBMIT
        )
        
        # 处理结果
        if result['status'] == 'login_required':
            iprint("\n[!] 需要重新登录")
        elif result['status'] == 'error':
            iprint(f"\n[X] 错误: {result.get('error')}")
        elif result['status'] == 'preview':
            iprint("\n[OK] 预览完成")
            iprint("\n要实际提交，请修改脚本设置:")
            iprint("  DRY_RUN = False")
            iprint("  AUTO_SUBMIT = True")
        else:
            iprint("\n[OK] 填写流程完成")
        
        # 用户选择
        iprint("\n" + "="*60)
        iprint("[系统] 执行完成")
        iprint("="*60)
        iprint("\n选项:")
        iprint("  1. 保持浏览器运行并退出（下次可快速复用）- 直接按Enter")
        iprint("  2. 关闭浏览器并退出 - 输入 'close' 后按Enter")
        
        try:
            user_input = input("\n请选择 (默认: 保持运行): ").strip().lower()
            if user_input == 'close':
                iprint("[系统] 正在关闭浏览器...")
                await browser_manager.close(keep_browser=False)
            else:
                iprint("[系统] 保持浏览器运行，可以复用")
                await browser_manager.close(keep_browser=True)
                iprint(f"\n[提示] 下次运行脚本时将自动复用此浏览器")
                iprint(f"[提示] 状态文件: {BROWSER_STATE_FILE}")
        except KeyboardInterrupt:
            iprint("\n[系统] 用户中断，保持浏览器运行")
            await browser_manager.close(keep_browser=True)
    
    except Exception as e:
        iprint(f"\n[X] 严重错误: {e}")
        traceback.print_exc()
        await browser_manager.close(keep_browser=True)


if __name__ == "__main__":
    asyncio.run(main())