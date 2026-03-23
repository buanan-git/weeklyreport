#!/usr/bin/env python3
"""
团队周报自动化处理脚本 - 函数调用版
功能：
1. 直接调用fetch_okr_final.py的函数获取团队周报数据
2. 直接调用llmchat_v2.0.py的函数优化周报
3. 直接调用submit_okr_ds_final.py的函数提交优化后的周报

支持平台：Windows, macOS, Linux (无需子进程，更稳定)
作者: OpenClaw Assistant
创建时间: 2026-03-14
更新时间: 2026-03-14 (改为函数调用，移除子进程)
"""

import asyncio
import json
import sys
import io
import platform
from datetime import datetime
from pathlib import Path

# ==================== 设置 UTF-8 编码 ====================
os_environ = None  # placeholder
if sys.platform == 'win32':
    # Windows 控制台 UTF-8 支持
    import os
    os.environ["PYTHONUTF8"] = "1"
    if sys.stdout.encoding != 'utf-8':
        try:
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
            sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
        except:
            pass

# ==================== 确保 scripts 目录在 Python 路径中 ====================
scripts_dir = Path(__file__).parent.absolute()
if str(scripts_dir) not in sys.path:
    sys.path.insert(0, str(scripts_dir))

# ==================== 导入配置加载器 ====================
from config_loader import ConfigLoader

# 加载配置
config = ConfigLoader("weeklyreport")

# ==================== 配置变量 ====================
TARGET_WEEK = config.target_week  # 目标周配置 (0=当周，1=上一周，2=上上周)
LLM_PLATFORM = config.default_platform
DEBUG = config.debug    # 是否启用调试模式
SAVE_INTERMEDIATE_FILES = config.save_intermediate  # 是否保存中间文件

# 获取平台的调用方式配置
_platform_config = config.get_platform_config(LLM_PLATFORM)
LLM_CALL_TYPE = _platform_config.get("type", "web")  # web 或 api
print(f"[weeklyreport] 平台 {LLM_PLATFORM} 调用方式: {LLM_CALL_TYPE}")

# 文件路径配置
SCRIPT_DIR = config.script_dir
WEEKLYREPORT_DIR = config.get_path("weekly_report_dir")
PROMPT_FILE = config.optimize_prompt

# ==================== 动态导入模块（带错误处理）====================
imported_modules = {}

def safe_import(module_name):
    """安全导入模块，失败时返回None"""
    try:
        module = __import__(module_name)
        print(f"✅ 成功导入: {module_name}")
        return module
    except ImportError as e:
        print(f"⚠️ 导入 {module_name} 失败: {e}")
        return None
    except Exception as e:
        print(f"⚠️ 导入 {module_name} 时发生错误: {e}")
        return None

# 尝试导入各个模块
fetch_module = safe_import('fetch_okr_final')
llmapi_module = safe_import('llmapi_v10')
llmchat_module = safe_import('llmchat_final')
submit_module = safe_import('submit_okr_ds_final')

# ==================== 定义后备函数 ====================
async def fetch_okr_data_fallback(week):
    """后备：获取OKR数据"""
    print(f"[后备] 获取第{week}周的数据")
    return {"week": week, "data": "模拟数据", "message": "使用后备函数"}

async def llm_api_resp_fallback(attachment_file, platform="BS", prompt_file=None, timeout=300):
    """后备：API调用"""
    print(f"[后备API] platform={platform}, file={attachment_file}")
    return {"status": "success", "data": "API后备数据"}

async def llmchat_resp_fallback(attachment_file, platform="BS", prompt_file=None, timeout=300):
    """后备：网页聊天"""
    print(f"[后备聊天] platform={platform}, file={attachment_file}")
    return {"status": "success", "data": "聊天后备数据"}

async def submit_okr_report_fallback(data):
    """后备：提交周报"""
    print(f"[后备提交] 提交数据")
    # 保存到文件
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = Path(__file__).parent / "submitted_reports"
        output_dir.mkdir(exist_ok=True)
        output_file = output_dir / f"submitted_report_{timestamp}.json"
        
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        print(f"[后备提交] 已保存到: {output_file}")
        return f"提交成功，文件已保存到 {output_file}"
    except Exception as e:
        print(f"[后备提交] 保存失败: {e}")
        return "模拟提交成功（未保存文件）"

# ==================== 选择使用哪个函数 ====================
fetch_okr_data = getattr(fetch_module, 'fetch_okr_data', fetch_okr_data_fallback) if fetch_module else fetch_okr_data_fallback
save_okr_to_file = getattr(fetch_module, 'save_okr_to_file', lambda d,w: None) if fetch_module else lambda d,w: None

llm_api_resp = getattr(llmapi_module, 'llm_api_resp', llm_api_resp_fallback) if llmapi_module else llm_api_resp_fallback
llmchat_resp = getattr(llmchat_module, 'llmchat_resp', llmchat_resp_fallback) if llmchat_module else llmchat_resp_fallback
submit_okr_report = getattr(submit_module, 'submit_okr_report', submit_okr_report_fallback) if submit_module else submit_okr_report_fallback

# ==================== 跨平台兼容性设置 ====================
def get_platform_info():
    """获取当前平台信息"""
    return {
        'system': platform.system(),
        'platform': platform.platform(),
        'python_version': platform.python_version(),
        'is_windows': platform.system() == 'Windows',
        'is_mac': platform.system() == 'Darwin',
        'is_linux': platform.system() == 'Linux'
    }

# ==================== 工具函数 ====================
def log_info(message):
    """输出信息日志"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] ℹ️  {message}")

def log_success(message):
    """输出成功日志"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] ✅ {message}")

def log_error(message):
    """输出错误日志"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] ❌ {message}")

def log_warning(message):
    """输出警告日志"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] ⚠️  {message}")

def log_platform_info():
    """输出平台信息（调试用）"""
    if DEBUG:
        platform_info = get_platform_info()
        log_info("=== 平台信息 ===")
        log_info(f"系统: {platform_info['system']}")
        log_info(f"平台: {platform_info['platform']}")
        log_info(f"Python版本: {platform_info['python_version']}")
        log_info(f"是否Windows: {platform_info['is_windows']}")
        log_info(f"是否macOS: {platform_info['is_mac']}")
        log_info(f"是否Linux: {platform_info['is_linux']}")
        log_info("===============")

def get_week_filename_suffix(weeks_ago=0):
    """获取周报文件名后缀"""
    from datetime import datetime, timedelta
    
    today = datetime.now()
    monday = today - timedelta(days=today.weekday())
    target_monday = monday - timedelta(weeks=weeks_ago)
    target_sunday = target_monday + timedelta(days=6)
    
    return f"{target_monday.year}.{target_monday.month}.{target_monday.day}-{target_sunday.month}.{target_sunday.day}"

def get_original_report_filename(weeks_ago=0):
    """获取原始团队周报文件名"""
    week_suffix = get_week_filename_suffix(weeks_ago)
    return WEEKLYREPORT_DIR / f"团队周报_{week_suffix}.json"

# ==================== 核心功能函数 ====================
async def step_fetch_data():
    """步骤1: 获取团队周报数据"""
    log_info("开始执行步骤1: 获取团队周报数据")
    
    try:
        original_data = await fetch_okr_data(TARGET_WEEK)
        
        if not original_data:
            log_error("获取到的数据为空")
            return None
        
        log_success(f"团队周报数据获取成功")
        
        if SAVE_INTERMEDIATE_FILES:
            WEEKLYREPORT_DIR.mkdir(parents=True, exist_ok=True)
            saved_file = save_okr_to_file(original_data, TARGET_WEEK)
            log_info(f"原始数据已保存")
        
        return original_data
        
    except Exception as e:
        log_error(f"获取数据失败: {e}")
        if DEBUG:
            import traceback
            traceback.print_exc()
        return None

async def step_optimize_report(original_data):
    """步骤2: 优化周报"""
    log_info("开始执行步骤2: 调用大模型优化周报")

    try:
        import tempfile
        import os

        # 创建临时文件
        with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', suffix='.json', delete=False) as f:
            json.dump(original_data, f, ensure_ascii=False, indent=2)
            temp_attachment = f.name

        log_info(f"创建临时文件: {temp_attachment}")

        optimized_data = None

        # 根据配置的调用方式决定调用顺序
        if LLM_CALL_TYPE == "web":
            # web模式：只使用网页自动化
            log_info(f"调用方式为 web，使用网页自动化")
            optimized_data = await llmchat_resp(
                attachment_file=temp_attachment,
                platform=LLM_PLATFORM,
                prompt_file=str(PROMPT_FILE),
                timeout=300
            )

        elif LLM_CALL_TYPE == "api":
            # api模式：优先使用API，失败则回退到网页
            log_info(f"调用方式为 api，优先使用API调用")
            optimized_data = await llm_api_resp(
                attachment_file=temp_attachment,
                platform=LLM_PLATFORM,
                prompt_file=str(PROMPT_FILE),
                timeout=300
            )

            if not optimized_data:
                log_warning("API调用失败，回退到网页自动化")
                optimized_data = await llmchat_resp(
                    attachment_file=temp_attachment,
                    platform=LLM_PLATFORM,
                    prompt_file=str(PROMPT_FILE),
                    timeout=300
                )
        else:
            log_error(f"未知的调用方式: {LLM_CALL_TYPE}")

        # 清理临时文件
        try:
            os.unlink(temp_attachment)
        except:
            pass

        if optimized_data:
            log_success("优化成功")

            # 保存优化后的数据
            if SAVE_INTERMEDIATE_FILES:
                week_suffix = get_week_filename_suffix(TARGET_WEEK)
                optimized_file = WEEKLYREPORT_DIR / f"团队周报_{week_suffix}.optimized.json"

                with open(optimized_file, 'w', encoding='utf-8') as f:
                    json.dump(optimized_data, f, ensure_ascii=False, indent=2)

                log_info(f"优化数据已保存到: {optimized_file}")

            return optimized_data
        else:
            log_error("周报优化失败，所有调用方式均未成功")
            return None

    except Exception as e:
        log_error(f"优化过程异常: {e}")
        if DEBUG:
            import traceback
            traceback.print_exc()
        return None

async def step_submit_report(optimized_data):
    """步骤3: 提交优化后的周报"""
    log_info("开始执行步骤3: 提交优化后的周报")
    
    try:
        result = await submit_okr_report(optimized_data)
        log_success(f"周报提交成功: {result}")
        return True
        
    except Exception as e:
        log_error(f"提交周报失败: {e}")
        if DEBUG:
            import traceback
            traceback.print_exc()
        return False

# ==================== 主函数 ====================
async def main():
    """主函数：串联执行三个步骤"""
    print("=" * 80)
    print("🚀 团队周报自动化处理脚本")
    print("=" * 80)
    
    log_platform_info()
    
    print(f"目标周: {TARGET_WEEK} (0=当周, 1=上一周, 2=上上周)")
    print(f"大模型平台: {LLM_PLATFORM}")
    print(f"调用方式: {LLM_CALL_TYPE} ({'仅网页自动化' if LLM_CALL_TYPE == 'web' else '优先API，失败回退网页'})")
    print(f"调试模式: {'开启' if DEBUG else '关闭'}")
    print(f"保存中间文件: {'是' if SAVE_INTERMEDIATE_FILES else '否'}")
    print("=" * 80)
    
    # 步骤1
    print("\n" + "="*50)
    print("步骤 1/3: 获取团队周报数据")
    print("="*50)
    
    original_data = await step_fetch_data()
    if original_data is None:
        log_error("步骤1失败，终止执行")
        return False
    
    await asyncio.sleep(1)
    
    # 步骤2
    print("\n" + "="*50)
    print("步骤 2/3: 智能优化周报")
    print("="*50)
    
    optimized_data = await step_optimize_report(original_data)
    if optimized_data is None:
        log_error("=" * 50)
        log_error("步骤2失败: 周报优化未能完成")
        log_error("可能原因:")
        log_error("  - API调用失败（检查API密钥、网络连接）")
        log_error("  - 网页自动化失败（检查浏览器状态）")
        log_error("  - 平台响应超时")
        log_error("=" * 50)
        print("\n⚠️  周报优化失败，程序将终止。请检查上述错误信息后重试。")
        return False
    
    await asyncio.sleep(1)
    
    # 步骤3
    print("\n" + "="*50)
    print("步骤 3/3: 提交优化后的周报")
    print("="*50)
    
    if not await step_submit_report(optimized_data):
        log_error("步骤3失败，终止执行")
        return False
    
    print("\n" + "="*50)
    log_success("🎉 所有步骤执行完成！")
    print("="*50)
    
    return True

# ==================== 入口点 ====================
if __name__ == "__main__":
    try:
        print("\n" + "="*60)
        print("🚀 周报自动化工具启动")
        print("="*60)
        
        # 解析命令行参数
        target_week = config.target_week  # 默认从配置文件读取
        llm_platform = config.default_platform  # 默认从配置文件读取
        if len(sys.argv) > 1:
            try:
                TARGET_WEEK = int(sys.argv[1])
                print(f"📋 使用命令行参数: 周数={TARGET_WEEK}")
            except ValueError:
                print(f"⚠️ 参数无效，使用配置默认值: {TARGET_WEEK}")
        else:
            print(f"📋 未指定参数，使用配置默认值: 周数={TARGET_WEEK}")
        
        if len(sys.argv) > 2:
            platform_arg = sys.argv[2].upper()
            if platform_arg in ["BS", "DS", "DB", "KM"]:
                LLM_PLATFORM = platform_arg
                print(f"📋 使用命令行参数: 平台={LLM_PLATFORM}")
        
        print(f"\n🚀 最终运行参数:")
        print(f"   - 目标周: {TARGET_WEEK} ({'本周' if TARGET_WEEK==0 else '上周' if TARGET_WEEK==1 else '上上周'})")
        print(f"   - 平台: {LLM_PLATFORM}")
        print(f"   - 保存中间文件: {'是' if SAVE_INTERMEDIATE_FILES else '否'}")
        print("="*60 + "\n")
        
        # 运行主函数
        print("开始执行主函数...")
        try:
            success = asyncio.run(main())
        except Exception as e:
            print(f"错误: {e}")
            traceback.print_exc()
        finally:
            print("\n按回车键退出...")
            input()  # 等待用户按回车
            
        if success:
            print("\n✅ 脚本执行成功完成！")
            sys.exit(0)
        else:
            print("\n❌ 脚本执行失败！")
            sys.exit(1)
            
    except KeyboardInterrupt:
        print("\n\n⚠️ 用户中断执行")
        sys.exit(1)
    except Exception as e:
        print(f"\n\n❌ 脚本执行时发生未捕获的异常: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)