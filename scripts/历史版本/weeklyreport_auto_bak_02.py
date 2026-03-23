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
import sys
import platform
from datetime import datetime
from pathlib import Path

# ==================== 新增：导入配置加载器 ====================
from config_loader import ConfigLoader

# 加载配置
config = ConfigLoader("weeklyreport")

# ==================== 替换原有的硬编码配置 ====================
TARGET_WEEK = config.target_week  # 目标周配置 (0=当周，1=上一周，2=上上周)
LLM_PLATFORM = config.default_platform
DEBUG = config.debug    # 是否启用调试模式
SAVE_INTERMEDIATE_FILES = config.save_intermediate      # 是否保存中间文件（可选，默认保存以便追溯）

# 文件路径配置
SCRIPT_DIR = config.script_dir
WEEKLYREPORT_DIR = config.get_path("weekly_report_dir")
PROMPT_FILE = config.optimize_prompt

# 导入各模块的函数
try:
    from fetch_okr_final import fetch_okr_data, save_okr_to_file
    from llmchat_final import llmchat_resp
    from submit_okr_ds_final import submit_okr_report
except ImportError as e:
    print(f"❌ 导入模块失败: {e}")
    print("请确保以下脚本已改造为可导入模块：")
    print("  - fetch_okr_final.py (需要包含 fetch_okr_data 函数)")
    print("  - llmchat4weeklyReport_v2.0.py (需要包含 optimize_report 函数)")
    print("  - submit_okr_ds_final.py (需要包含 submit_okr_report 函数)")
    sys.exit(1)

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
    """
    根据传入的参数获取指定周的周报文件名后缀，格式为：2026.3.2-3.8
    :param weeks_ago: 0=当周，1=上一周，2=上上周
    :return: 格式化的文件名后缀字符串
    """
    def get_week_range(weeks_ago):
        """获取指定周的周一和周日日期"""
        from datetime import datetime, timedelta
        
        today = datetime.now()
        # 找到本周的周一
        monday = today - timedelta(days=today.weekday())
        # 调整到目标周
        target_monday = monday - timedelta(weeks=weeks_ago)
        target_sunday = target_monday + timedelta(days=6)
        
        return target_monday, target_sunday
    
    monday, sunday = get_week_range(weeks_ago)
    
    # 提取年份、月、日信息
    year = monday.year
    start_month = monday.month
    start_day = monday.day
    end_month = sunday.month
    end_day = sunday.day
    
    # 按指定格式拼接（年份.月.日-月.日）
    return f"{year}.{start_month}.{start_day}-{end_month}.{end_day}"

def get_original_report_filename(weeks_ago=0):
    """获取原始团队周报文件名"""
    week_suffix = get_week_filename_suffix(weeks_ago)
    return WEEKLYREPORT_DIR / f"团队周报_{week_suffix}.json"

def get_optimized_report_filename(weeks_ago=0):
    """获取优化后的团队周报文件名"""
    week_suffix = get_week_filename_suffix(weeks_ago)
    return WEEKLYREPORT_DIR / f"团队周报_{week_suffix}.llm.json"

# ==================== 核心功能函数 ====================

async def step_fetch_data():
    """步骤1: 获取团队周报数据"""
    log_info("开始执行步骤1: 获取团队周报数据")
    
    try:
        # 直接调用函数获取数据
        original_data = await fetch_okr_data(TARGET_WEEK)
        
        if not original_data:
            log_error("获取到的数据为空")
            return None
        
        log_success(f"团队周报数据获取成功，共 {len(original_data) if isinstance(original_data, (list, dict)) else '?'} 条记录")
        
        # 如果需要保存中间文件
        if SAVE_INTERMEDIATE_FILES:
            WEEKLYREPORT_DIR.mkdir(parents=True, exist_ok=True)
            saved_file = save_okr_to_file(original_data, TARGET_WEEK)
            log_info(f"原始数据已保存到: {saved_file}")
        
        return original_data
        
    except Exception as e:
        log_error(f"获取数据失败: {e}")
        if DEBUG:
            import traceback
            traceback.print_exc()
        return None



        
async def step_submit_report(optimized_data):
    """步骤3: 提交优化后的周报"""
    log_info("开始执行步骤3: 提交优化后的周报")
    
    try:
        # 直接调用提交函数
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
    print("🚀 团队周报自动化处理脚本 - 函数调用版")
    print("=" * 80)
    
    # 输出平台信息
    log_platform_info()
    
    print(f"目标周: {TARGET_WEEK} (0=当周, 1=上一周, 2=上上周)")
    print(f"大模型平台: {LLM_PLATFORM}")
    print(f"调试模式: {'开启' if DEBUG else '关闭'}")
    print(f"保存中间文件: {'是' if SAVE_INTERMEDIATE_FILES else '否'}")
    print("=" * 80)
    
    # 检查必要目录
    if SAVE_INTERMEDIATE_FILES and not WEEKLYREPORT_DIR.exists():
        log_info(f"创建WeeklyReport目录: {WEEKLYREPORT_DIR}")
        WEEKLYREPORT_DIR.mkdir(parents=True, exist_ok=True)
    
    # 步骤1: 获取团队周报数据
    log_info("=" * 50)
    log_info("步骤 1/3: 获取团队周报数据")
    log_info("=" * 50)
    
    original_data = await step_fetch_data()
    if original_data is None:
        log_error("步骤1失败，终止执行")
        return False
    
    # 步骤1完成后等待1秒
    log_info("步骤1完成，等待1秒...")
    await asyncio.sleep(1)
    
    # 步骤2: 调用大模型优化周报
    log_info("=" * 50)
    log_info("步骤 2/3: 调用大模型优化周报")
    log_info("=" * 50)
    
    optimized_data = await step_optimize_report(original_data)
    if optimized_data is None:
        log_error("步骤2失败，终止执行")
        return False
    
    # 步骤2完成后等待2秒，确保文件完全写入
    log_info("步骤2完成，等待2秒确保文件写入...")
    await asyncio.sleep(2)
    
    # 步骤3: 提交优化后的周报
    log_info("=" * 50)
    log_info("步骤 3/3: 提交优化后的周报")
    log_info("=" * 50)
    
    if not await step_submit_report(optimized_data):
        log_error("步骤3失败，终止执行")
        return False
    
    # 全部完成
    log_info("=" * 50)
    log_success("🎉 所有步骤执行完成！")
    log_info("=" * 50)
    
    return True

# ==================== 入口点 ====================

if __name__ == "__main__":
    try:
        # 解析命令行参数
        if len(sys.argv) > 1:
            try:
                TARGET_WEEK = int(sys.argv[1])
            except ValueError:
                print(f"❌ 错误：参数「{sys.argv[1]}」不是有效数字！")
                sys.exit(1)
        
        if len(sys.argv) > 2:
            platform_arg = sys.argv[2].upper()
            if platform_arg in ["DS", "DB", "KM"]:
                LLM_PLATFORM = platform_arg
            else:
                print(f"❌ 错误：平台参数「{sys.argv[2]}」无效！支持: DS, DB, KM")
                sys.exit(1)
        
        if len(sys.argv) > 3:
            # 第三个参数可选：是否保存文件
            SAVE_INTERMEDIATE_FILES = sys.argv[3].lower() in ['true', '1', 'yes', 'y']
        
        # 运行主函数
        success = asyncio.run(main())
        
        if success:
            print("\n🎊 脚本执行成功完成！")
            sys.exit(0)
        else:
            print("\n💥 脚本执行失败！")
            sys.exit(1)
            
    except KeyboardInterrupt:
        print("\n\n⚠️  用户中断执行")
        sys.exit(1)
    except Exception as e:
        print(f"\n\n💥 脚本执行时发生未捕获的异常: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)