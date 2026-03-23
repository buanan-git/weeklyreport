#!/usr/bin/env python3
"""
LLM API 调用脚本 - 同步版
支持 DeepSeek、豆包、Kimi 等所有OpenAI兼容API
所有差异通过 config.json 配置，无需为每个平台单独编写代码
"""

import os
import json
import requests
from pathlib import Path
from typing import Optional, Dict, Any, List, Union
import sys
from datetime import datetime
import re

# 添加项目根目录到路径
sys.path.insert(0, str(Path(__file__).parent.parent))
from config_loader import ConfigLoader

# 加载配置
config = ConfigLoader("llm_api")

class SyncAPIClient:
    """
    同步API客户端 - 支持所有OpenAI兼容的LLM API
    """
    
    def __init__(self, platform: str):
        self.platform = platform
        self.platform_config = config.get_platform_config(platform)
        self.api_config = self.platform_config.get("api_config", {})
        
        # 基础配置
        self.base_url = self.api_config.get("base_url", "").rstrip('/')
        self.api_key = self.api_config.get("api_key", "")
        self.model = self.api_config.get("model", "")
        
        # API调用配置
        self.timeout = self.api_config.get("timeout", 180)
        self.max_tokens = self.api_config.get("max_tokens", 4000)
        self.temperature = self.api_config.get("temperature", 0.7)
        
        # API端点路径
        self.api_endpoint = self.api_config.get("api_endpoint", "/chat/completions")
        
        # 完整URL
        self.api_url = self.base_url + self.api_endpoint
        
        # 响应路径
        self.response_path = self.api_config.get("response_path", "choices.0.message.content")
        
        print(f"[{self.platform} API] 客户端初始化: {self.api_url}, 模型: {self.model}")
        print(f"[{self.platform} API] 响应路径: {self.response_path}")
    
    def call_api(self, system_prompt: str, user_content: str) -> Optional[str]:
        """
        调用API
        """
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_key}"
        }
        
        # 构建消息
        messages = []
        if system_prompt:
            messages.append({"role": "system", "content": system_prompt})
        messages.append({"role": "user", "content": user_content})
        
        # 构建请求体
        payload = {
            "model": self.model,
            "messages": messages,
            "max_tokens": self.max_tokens,
            "temperature": self.temperature
        }
        
        print(f"[{self.platform} API] 发送请求到: {self.api_url}")
        print(f"[{self.platform} API] 消息数: {len(messages)}, 模型: {self.model}")
        print(f"[{self.platform} API] 请求头: {headers}")
        
        try:
            print("正在调用API...")
            response = requests.post(
                self.api_url, 
                headers=headers, 
                json=payload, 
                timeout=self.timeout
            )
            
            if response.status_code == 200:
                result = response.json()
                print("API调用成功!")
                
                # 提取响应内容
                content = self._extract_response(result)
                if content:
                    print(f"响应长度: {len(content)} 字符")
                    return content
                else:
                    print("无法提取响应内容")
                    return None
            else:
                print(f"API调用失败: {response.status_code}")
                print(f"错误信息: {response.text}")
                return None
                
        except requests.exceptions.Timeout:
            print(f"API调用超时 ({self.timeout}秒)")
            return None
        except Exception as e:
            print(f"API调用异常: {e}")
            return None
    
    def _extract_response(self, response_data: Dict[str, Any]) -> Optional[str]:
        """
        从响应中提取内容
        """
        try:
            # 按照路径提取内容
            path_parts = self.response_path.split('.')
            current = response_data
            
            for part in path_parts:
                if isinstance(current, dict) and part in current:
                    current = current[part]
                else:
                    print(f"无法找到路径: {part}")
                    return None
            
            return str(current) if current else None
            
        except Exception as e:
            print(f"提取响应时出错: {e}")
            return None

def log_msg(message: str, level: str = "INFO"):
    """日志输出"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    print(f"[{timestamp}] [{level}] {message}")

def main():
    import argparse
    
    parser = argparse.ArgumentParser(description="LLM API调用脚本 - 同步版")
    parser.add_argument("json_file", help="JSON数据文件路径")
    parser.add_argument("-P", "--prompt", help="提示词文件路径")
    parser.add_argument("-p", "--platform", help="使用的平台 (BS/DS/DB/KM)")
    parser.add_argument("-t", "--timeout", type=int, help="超时时间（秒）")
    parser.add_argument("-o", "--output", help="输出文件路径")
    parser.add_argument("-v", "--verbose", action="store_true", help="详细输出")
    
    args = parser.parse_args()
    
    # 确定平台
    platform = args.platform or config.default_platform
    timeout = args.timeout or 180
    
    log_msg(f"平台: {platform}, 超时: {timeout}秒")
    
    # 检查文件
    json_file = Path(args.json_file)
    if not json_file.exists():
        log_msg(f"JSON文件不存在: {json_file}", "ERROR")
        return
    
    prompt_file = Path(args.prompt) if args.prompt else config.optimize_prompt
    if not prompt_file.exists():
        log_msg(f"提示词文件不存在: {prompt_file}", "ERROR")
        return
    
    # 读取数据
    try:
        with open(json_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        log_msg(f"数据加载成功: {json_file.name}")
        log_msg(f"数据大小: {len(json.dumps(data, ensure_ascii=False))} 字符")
    except Exception as e:
        log_msg(f"读取JSON文件失败: {e}", "ERROR")
        return
    
    # 读取提示词
    try:
        with open(prompt_file, 'r', encoding='utf-8') as f:
            prompt = f.read()
        log_msg(f"提示词加载成功: {prompt_file.name}")
        log_msg(f"提示词长度: {len(prompt)} 字符")
    except Exception as e:
        log_msg(f"读取提示词文件失败: {e}", "ERROR")
        return
    
    # 准备请求数据
    data_str = json.dumps(data, ensure_ascii=False, indent=2)
    
    # 如果数据太大，进行截断
    if len(data_str) > 50000:
        log_msg(f"数据较大 ({len(data_str)} 字符)，进行截断处理", "WARNING")
        data_str = data_str[:25000] + "\n\n[...数据截断...]\n\n" + data_str[-25000:]
    
    # 构建用户消息
    user_message = f"请优化以下团队周报数据：\n\n{data_str}"
    
    # 创建API客户端
    client = SyncAPIClient(platform)
    
    # 调用API
    start_time = datetime.now()
    result = client.call_api(prompt, user_message)
    end_time = datetime.now()
    
    if result:
        # 计算耗时
        duration = (end_time - start_time).total_seconds()
        log_msg(f"API调用成功 (耗时: {duration:.1f}秒)")
        
        # 保存结果
        if args.output:
            output_file = Path(args.output)
        else:
            # 默认输出文件名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = json_file.parent / f"llm_response_{timestamp}.json"
        
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(result)
            log_msg(f"结果已保存到: {output_file}")
            
            # 显示预览
            print("\n" + "="*60)
            print("API调用成功，结果预览:")
            print("="*60)
            preview = result[:500] + "..." if len(result) > 500 else result
            print(preview)
            print("="*60)
            print(f"\n完整结果已保存到: {output_file}")
            
        except Exception as e:
            log_msg(f"保存结果失败: {e}", "ERROR")
    else:
        log_msg("API调用失败")

if __name__ == "__main__":
    main()