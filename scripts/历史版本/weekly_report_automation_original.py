#!/usr/bin/env python3
"""
团队周报自动化处理脚本 - 跨平台兼容版
功能：
1. 调用fetch_okr_final.py获取团队周报数据
2. 使用llmchat_v2.0.py调用大模型优化周报
3. 调用submit_okr_ds_final.py提交优化后的周报

支持平台：Windows, macOS, Linux
作者: OpenClaw Assistant
创建时间: 2026-03-14
更新时间: 2026-03-14 (增强跨平台兼容性)
"""

import asyncio
import sys
import os
import json
import subprocess
import time
import shutil
import platform
from datetime import datetime
from pathlib import Path

# ==================== 配置区域 ====================

# 脚本所在目录
SCRIPT_DIR = Path(__file__).parent.absolute()

# 目标周配置 (0=当周，1=上一周，2=上上周)
TARGET_WEEK = 0

# 大模型平台配置
LLM_PLATFORM = "DS"  # DS=DeepSeek, DB=豆包, KM=Kimi

# 是否启用调试模式
DEBUG = True

# 文件路径配置
FETCH_SCRIPT = SCRIPT_DIR / "fetch_okr_final.py"
LLMCHAT_SCRIPT = SCRIPT_DIR / "llmchat4weeklyReport_v2.0.py"
SUBMIT_SCRIPT = SCRIPT_DIR / "submit_okr_ds_final.py"
WEEKLYREPORT_DIR = SCRIPT_DIR / "WeeklyReport"
PROMPT_FILE = SCRIPT_DIR / "WeeklyReportOptimizePrompt.md"

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

def ensure_executable(script_path):
    """确保脚本文件可执行（非Windows系统）"""
    platform_info = get_platform_info()
    if not platform_info['is_windows'] and script_path.exists():
        try:
            current_mode = script_path.stat().st_mode
            # 添加执行权限 (755)
            script_path.chmod(current_mode | 0o755)
            return True
        except Exception as e:
            log_warning(f"无法设置脚本执行权限 {script_path}: {e}")
            return False
    return True

def normalize_path(path):
    """标准化路径，确保跨平台兼容"""
    if isinstance(path, str):
        path = Path(path)
    return path.resolve()

def safe_subprocess_run(cmd, cwd=None, timeout=None):
    """
    跨平台安全的子进程运行函数
    :param cmd: 命令列表
    :param cwd: 工作目录
    :param timeout: 超时时间（秒）
    :return: subprocess.CompletedProcess
    """
    platform_info = get_platform_info()
    
    # 确保命令中的路径是字符串格式
    cmd = [str(arg) for arg in cmd]
    
    # Windows特定的环境设置
    env = os.environ.copy()
    if platform_info['is_windows']:
        # 在Windows上确保Python输出使用UTF-8
        env['PYTHONIOENCODING'] = 'utf-8'
        # 禁用Windows上的快速编辑模式，避免子进程被意外暂停
        env['PYTHONUNBUFFERED'] = '1'
    
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            cwd=cwd,
            timeout=timeout,
            env=env,
            # Windows上不需要shell，Unix-like系统上也不需要shell
            shell=False
        )
        return result
    except subprocess.TimeoutExpired:
        log_error(f"命令执行超时: {' '.join(cmd)}")
        raise
    except FileNotFoundError:
        log_error(f"命令或文件不存在: {cmd[0]}")
        raise
    except Exception as e:
        log_error(f"执行命令时发生异常: {' '.join(cmd)}, 错误: {e}")
        raise

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

async def run_fetch_script():
    """步骤1: 运行fetch_okr_final.py获取团队周报数据"""
    log_info("开始执行步骤1: 获取团队周报数据")
    
    # 检查脚本文件是否存在
    if not FETCH_SCRIPT.exists():
        log_error(f"fetch脚本不存在: {FETCH_SCRIPT}")
        return False
    
    # 构建命令
    cmd = [sys.executable, str(FETCH_SCRIPT), str(TARGET_WEEK)]
    
    log_info(f"执行命令: {' '.join(cmd)}")
    
    try:
        # 运行脚本
        result = subprocess.run(cmd, capture_output=True, text=True, cwd=SCRIPT_DIR)
        
        if result.returncode == 0:
            log_success("团队周报数据获取成功")
            if DEBUG:
                print("=== fetch脚本输出 ===")
                print(result.stdout)
                print("=====================")
            
            # 检查生成的文件是否存在
            original_file = get_original_report_filename(TARGET_WEEK)
            if original_file.exists():
                log_success(f"原始周报文件已生成: {original_file}")
                return True
            else:
                log_error(f"原始周报文件未生成: {original_file}")
                return False
        else:
            log_error(f"fetch脚本执行失败，返回码: {result.returncode}")
            log_error(f"错误输出: {result.stderr}")
            return False
            
    except Exception as e:
        log_error(f"运行fetch脚本时发生异常: {e}")
        return False

async def run_llm_optimization():
    """步骤2: 使用llmchat_v2.0.py调用大模型优化周报"""
    log_info("开始执行步骤2: 调用大模型优化周报")
    
    # 检查必要文件
    if not LLMCHAT_SCRIPT.exists():
        log_error(f"llmchat脚本不存在: {LLMCHAT_SCRIPT}")
        return False
    
    if not PROMPT_FILE.exists():
        log_error(f"prompt文件不存在: {PROMPT_FILE}")
        return False
    
    original_file = get_original_report_filename(TARGET_WEEK)
    if not original_file.exists():
        log_error(f"原始周报文件不存在: {original_file}")
        return False
    #####################注释如下通过改目标脚本文件的方式；而是优化目标脚本文件，支持传入文件作为prompt###################
    # # 读取WeeklyReportOptimizePrompt.md的内容
    # try:
        # with open(PROMPT_FILE, 'r', encoding='utf-8') as f:
            # prompt_content = f.read()
        # log_success(f"成功读取prompt文件: {PROMPT_FILE}")
        # if DEBUG:
            # print(f"=== Prompt内容预览 ===")
            # print(prompt_content[:200] + "..." if len(prompt_content) > 200 else prompt_content)
            # print("=====================")
    # except Exception as e:
        # log_error(f"读取prompt文件失败: {e}")
        # return False
    
    # # 临时修改llmchat_v2.0.py中的CLAW2DS_PROMPT_FILE常量
    # # 先备份原文件
    # import shutil
    # backup_file = LLMCHAT_SCRIPT.with_suffix('.py.backup')
    # if not backup_file.exists():
        # shutil.copy2(LLMCHAT_SCRIPT, backup_file)
        # log_info(f"已备份原脚本: {backup_file}")
    
    # try:
        # # 读取原脚本内容
        # with open(LLMCHAT_SCRIPT, 'r', encoding='utf-8') as f:
            # script_content = f.read()
        
        # # 替换CLAW2DS_PROMPT_FILE常量
        # old_prompt = 'CLAW2DS_PROMPT_FILE = "附件是openclaw的请求报文，请不要回显内容，而是根据报文信息及请求进行响应。"'
        # new_prompt = f'CLAW2DS_PROMPT_FILE = """{prompt_content}"""'
        
        # if old_prompt in script_content:
            # modified_content = script_content.replace(old_prompt, new_prompt)
            
            # # 写入修改后的脚本
            # with open(LLMCHAT_SCRIPT, 'w', encoding='utf-8') as f:
                # f.write(modified_content)
            
            # log_success("已成功修改llmchat脚本的提示词")
        # else:
            # log_error("未找到要替换的提示词常量")
            # return False
        
        # 构建llmchat命令
        # 使用文件上传模式，因为周报文件通常较大
        cmd = [
            sys.executable, str(LLMCHAT_SCRIPT),
            f"--{LLM_PLATFORM}",  # 选择平台
            "--file", str(original_file),  # 上传原始周报文件
            "-df",
            "--upload-only",  # 仅上传文件，不发送文本
        ]
        
        log_info(f"执行命令: {' '.join(cmd)}")
        
        # 运行llmchat脚本
        result = subprocess.run(cmd, capture_output=True, text=True, cwd=SCRIPT_DIR)
        
        if result.returncode == 0:
            log_success("大模型优化调用成功")
            if DEBUG:
                print("=== llmchat脚本输出 ===")
                print(result.stdout)
                print("=======================")
                if result.stderr:
                    print("=== llmchat脚本错误输出 ===")
                    print(result.stderr)
                    print("===========================")
            
            # 等待一段时间让文件生成完成
            await asyncio.sleep(5)
            
            # 检查优化后的文件是否存在
            optimized_file = get_optimized_report_filename(TARGET_WEEK)
            if optimized_file.exists():
                log_success(f"优化周报文件已生成: {optimized_file}")
                return True
            else:
                log_error(f"优化周报文件未生成: {optimized_file}")
                return False
        else:
            log_error(f"llmchat脚本执行失败，返回码: {result.returncode}")
            log_error(f"错误输出: {result.stderr}")
            return False
            
    except Exception as e:
        log_error(f"运行llmchat脚本时发生异常: {e}")
        return False
    finally:
        # 恢复原脚本
        try:
            shutil.copy2(backup_file, LLMCHAT_SCRIPT)
            log_info("已恢复原脚本文件")
        except Exception as e:
            log_warning(f"恢复原脚本文件时发生警告: {e}")

async def run_submit_script():
    """步骤3: 运行submit_okr_ds_final.py提交优化后的周报"""
    log_info("开始执行步骤3: 提交优化后的周报")
    
    # 检查脚本文件是否存在
    if not SUBMIT_SCRIPT.exists():
        log_error(f"submit脚本不存在: {SUBMIT_SCRIPT}")
        return False
    
    # 检查优化后的文件是否存在
    optimized_file = get_optimized_report_filename(TARGET_WEEK)
    if not optimized_file.exists():
        log_error(f"优化周报文件不存在: {optimized_file}")
        return False
    
    # 构建命令
    cmd = [sys.executable, str(SUBMIT_SCRIPT), str(TARGET_WEEK)]
    
    log_info(f"执行命令: {' '.join(cmd)}")
    
    try:
        # 运行脚本
        result = subprocess.run(cmd, capture_output=True, text=True, cwd=SCRIPT_DIR)
        
        if result.returncode == 0:
            log_success("周报提交成功")
            if DEBUG:
                print("=== submit脚本输出 ===")
                print(result.stdout)
                print("=======================")
            
            return True
        else:
            log_error(f"submit脚本执行失败，返回码: {result.returncode}")
            log_error(f"错误输出: {result.stderr}")
            return False
            
    except Exception as e:
        log_error(f"运行submit脚本时发生异常: {e}")
        return False

# ==================== 主函数 ====================

async def main():
    """主函数：串联执行三个步骤"""
    print("=" * 80)
    print("🚀 团队周报自动化处理脚本 - 串联版")
    print("=" * 80)
    print(f"目标周: {TARGET_WEEK} (0=当周, 1=上一周, 2=上上周)")
    print(f"大模型平台: {LLM_PLATFORM}")
    print(f"调试模式: {'开启' if DEBUG else '关闭'}")
    print("=" * 80)
    
    # 检查必要目录
    if not WEEKLYREPORT_DIR.exists():
        log_info(f"创建WeeklyReport目录: {WEEKLYREPORT_DIR}")
        WEEKLYREPORT_DIR.mkdir(parents=True, exist_ok=True)
    
    # 步骤1: 获取团队周报数据
    log_info("=" * 50)
    log_info("步骤 1/3: 获取团队周报数据")
    log_info("=" * 50)
    
    if not await run_fetch_script():
        log_error("步骤1失败，终止执行")
        return False
    
    # 步骤2: 调用大模型优化周报
    log_info("=" * 50)
    log_info("步骤 2/3: 调用大模型优化周报")
    log_info("=" * 50)
    
    if not await run_llm_optimization():
        log_error("步骤2失败，终止执行")
        return False
    
    # 步骤3: 提交优化后的周报
    log_info("=" * 50)
    log_info("步骤 3/3: 提交优化后的周报")
    log_info("=" * 50)
    
    if not await run_submit_script():
        log_error("步骤3失败，终止执行")
        return False
    
    # 全部完成
    log_info("=" * 50)
    log_success("🎉 所有步骤执行完成！")
    log_info("=" * 50)
    
    # 显示生成的文件
    original_file = get_original_report_filename(TARGET_WEEK)
    optimized_file = get_optimized_report_filename(TARGET_WEEK)
    
    log_info("生成的文件:")
    if original_file.exists():
        log_info(f"  - 原始周报: {original_file}")
    if optimized_file.exists():
        log_info(f"  - 优化周报: {optimized_file}")
    
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