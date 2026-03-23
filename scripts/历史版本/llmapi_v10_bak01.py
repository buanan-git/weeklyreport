#!/usr/bin/env python3
"""
LLM API 调用脚本 - 统一版
支持 DeepSeek、豆包、Kimi 等所有OpenAI兼容API
所有差异通过 config.json 配置，无需为每个平台单独编写代码
"""

import os
import json
import asyncio
import aiohttp
from pathlib import Path
from typing import Optional, Dict, Any, List, Union
import sys
from datetime import datetime
import re

# 添加项目根目录到路径
sys.path.insert(0, str(Path(__file__).parent.parent))

@property
def prompt_file(self) -> Path:
    # 优先使用新的通用配置
    llm_chat = self.config.get("llm_chat", {})
    if llm_chat.get("prompt_file"):
        return self.get_prompt_file(llm_chat["prompt_file"])
    
    # 回退到原有配置
    return self.optimize_prompt

class UnifiedAPIClient:
    """
    统一API客户端 - 支持所有OpenAI兼容的LLM API
    所有差异通过配置管理，无需子类化
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
        self.timeout = self.api_config.get("timeout", 300)
        self.max_tokens = self.api_config.get("max_tokens", 4000)
        self.temperature = self.api_config.get("temperature", 0.7)
        
        # ===== 新增：平台差异配置 =====
        # API端点路径（如 /chat/completions, /v1/chat/completions, /chat）
        self.api_endpoint = self.api_config.get("api_endpoint", "/chat/completions")
        
        # 完整URL
        self.api_url = f"{self.base_url}{self.api_endpoint}"
        
        # 响应内容提取路径（点号分隔，如 "choices.0.message.content"）
        self.response_path = self.api_config.get("response_path", "choices.0.message.content")
        
        # 请求格式适配
        self.request_format = self.api_config.get("request_format", "openai")  # openai 或 custom
        
        # 自定义请求模板（如果request_format为custom）
        self.request_template = self.api_config.get("request_template", {})
        
        # 自定义响应解析函数（可选）
        self.response_parser = self.api_config.get("response_parser")
        
        # 额外请求头
        self.extra_headers = self.api_config.get("extra_headers", {})
        
        # 会话管理
        self.session = None
        
        self._log(f"客户端初始化: {self.api_url}, 模型: {self.model}")
        self._log(f"响应路径: {self.response_path}")
        
    async def __aenter__(self):
        self.session = aiohttp.ClientSession()
        return self
        
    async def __aexit__(self, exc_type, exc_val, exc_tb):
        if self.session:
            await self.session.close()
    
    def _log(self, message: str, level: str = "INFO"):
        """统一日志输出"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{timestamp}] [{self.platform} API] {message}", flush=True)
    
    async def check_api_available(self) -> bool:
        """
        检查API是否可用（API密钥是否配置正确）
        """
        if not self.api_key:
            self._log("API密钥未配置", "WARNING")
            return False
        
        if self.api_key.startswith("your-") or "your" in self.api_key.lower():
            self._log(f"API密钥为默认值，请修改: {self.api_key[:10]}...", "WARNING")
            return False
        
        if not self.base_url:
            self._log("API基础地址未配置", "WARNING")
            return False
        
        self._log(f"API配置检查通过: {self.api_url}")
        return True
    
    def _build_request_payload(self, messages: List[Dict[str, str]], temperature: float = None) -> Dict:
        """
        根据配置构建请求载荷
        """
        if self.request_format == "openai":
            # OpenAI 标准格式
            return {
                "model": self.model,
                "messages": messages,
                "temperature": temperature or self.temperature,
                "max_tokens": self.max_tokens,
                "stream": False
            }
        elif self.request_format == "custom" and self.request_template:
            # 自定义格式 - 使用模板并替换变量
            template_str = json.dumps(self.request_template)
            # 替换模板变量
            template_str = template_str.replace("${model}", self.model)
            template_str = template_str.replace("${temperature}", str(temperature or self.temperature))
            template_str = template_str.replace("${max_tokens}", str(self.max_tokens))
            
            # 处理消息（需要特殊处理，因为消息是列表）
            import copy
            payload = copy.deepcopy(self.request_template)
            
            # 递归替换消息中的占位符
            def replace_placeholders(obj):
                if isinstance(obj, dict):
                    for key, value in obj.items():
                        if isinstance(value, str) and "${messages}" in value:
                            obj[key] = messages
                        elif isinstance(value, (dict, list)):
                            replace_placeholders(value)
                elif isinstance(obj, list):
                    for i, item in enumerate(obj):
                        if isinstance(item, str) and "${messages}" in item:
                            obj[i] = messages
                        elif isinstance(item, (dict, list)):
                            replace_placeholders(item)
            
            replace_placeholders(payload)
            return payload
        else:
            # 默认使用OpenAI格式
            return {
                "model": self.model,
                "messages": messages,
                "temperature": temperature or self.temperature,
                "max_tokens": self.max_tokens
            }
    
    def _extract_content_from_response(self, response_data: Dict) -> Optional[str]:
        """
        根据配置的响应路径从响应中提取内容
        支持点号路径，如 "choices.0.message.content"
        """
        if self.response_parser:
            # 如果有自定义解析器，使用eval执行（注意安全）
            try:
                # 限制eval的环境，只允许基本的操作
                safe_dict = {
                    'data': response_data,
                    'json': json,
                    'str': str,
                    'dict': dict,
                    'list': list
                }
                result = eval(self.response_parser, {"__builtins__": {}}, safe_dict)
                if result:
                    return str(result)
            except Exception as e:
                self._log(f"自定义解析器执行失败: {e}", "WARNING")
        
        # 使用点号路径提取
        if not self.response_path:
            # 如果没有配置响应路径，尝试常见的格式
            if isinstance(response_data, dict):
                if "response" in response_data:
                    return response_data["response"]
                elif "content" in response_data:
                    return response_data["content"]
                elif "text" in response_data:
                    return response_data["text"]
                elif "choices" in response_data and len(response_data["choices"]) > 0:
                    choice = response_data["choices"][0]
                    if isinstance(choice, dict):
                        if "message" in choice and "content" in choice["message"]:
                            return choice["message"]["content"]
                        elif "text" in choice:
                            return choice["text"]
            return None
        
        # 按点号路径提取
        try:
            parts = self.response_path.split('.')
            value = response_data
            for part in parts:
                if part.isdigit():
                    # 数字索引
                    idx = int(part)
                    if isinstance(value, list) and idx < len(value):
                        value = value[idx]
                    else:
                        return None
                else:
                    # 字典键
                    if isinstance(value, dict) and part in value:
                        value = value[part]
                    else:
                        return None
            return str(value) if value is not None else None
        except Exception as e:
            self._log(f"响应提取失败: {e}", "ERROR")
            return None
    
    async def call_api(self, messages: List[Dict[str, str]], temperature: float = None) -> Optional[str]:
        """
        调用API - 统一方法
        """
        if not self.session:
            self.session = aiohttp.ClientSession()
        
        # 构建请求头
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_key}"
        }
        # 添加额外请求头
        headers.update(self.extra_headers)
        
        # 构建请求载荷
        payload = self._build_request_payload(messages, temperature)
        
        self._log(f"发送请求到 {self.api_url}")
        self._log(f"消息数: {len(messages)}, 模型: {self.model}")
        
        # 调试模式下打印请求摘要
        if config.debug:
            self._log(f"请求头: { {k:v[:20] + '...' if k=='Authorization' else v for k,v in headers.items()} }")
            self._log(f"请求载荷摘要: {json.dumps(payload, ensure_ascii=False)[:200]}...")
        
        try:
            async with self.session.post(
                self.api_url, 
                headers=headers, 
                json=payload, 
                timeout=self.timeout
            ) as response:
                
                # 读取响应
                response_text = await response.text()
                
                if response.status != 200:
                    self._log(f"API 返回错误: HTTP {response.status}", "ERROR")
                    self._log(f"错误详情: {response_text[:500]}", "ERROR")
                    return None
                
                # 解析JSON
                try:
                    result = json.loads(response_text)
                except json.JSONDecodeError:
                    self._log(f"响应不是有效的JSON: {response_text[:200]}", "ERROR")
                    return None
                
                # 提取内容
                content = self._extract_content_from_response(result)
                
                if content:
                    self._log(f"API 调用成功，返回内容长度: {len(content)}")
                    return content
                else:
                    self._log(f"无法从响应中提取内容", "WARNING")
                    self._log(f"响应结构: {json.dumps(result, ensure_ascii=False)[:500]}", "DEBUG")
                    return None
                    
        except asyncio.TimeoutError:
            self._log(f"API 调用超时 ({self.timeout}秒)", "ERROR")
            return None
        except aiohttp.ClientError as e:
            self._log(f"HTTP 客户端错误: {e}", "ERROR")
            return None
        except Exception as e:
            self._log(f"API 调用异常: {e}", "ERROR")
            import traceback
            traceback.print_exc()
            return None


def get_api_client(platform: str) -> UnifiedAPIClient:
    """获取对应平台的API客户端（统一版本）"""
    return UnifiedAPIClient(platform)


async def llm_api_resp(
    attachment_file: str = None,
    platform: str = "DS",
    prompt_file: str = None,
    timeout: int = 300,
    temperature: float = None
) -> Optional[Dict[str, Any]]:
    """
    LLM API 调用主函数（统一版）
    
    :param attachment_file: 附件文件路径（包含要处理的数据）
    :param platform: 平台名称 (DS/DB/KM)
    :param prompt_file: 提示词文件路径
    :param timeout: 超时时间（秒）
    :param temperature: 温度参数（可选）
    :return: 优化后的数据，失败返回 None
    """
    
    def log_msg(msg, level="INFO"):
        print(f"[{datetime.now().strftime('%H:%M:%S')}] [LLM-API] {msg}", flush=True)
    
    log_msg(f"平台: {platform}, 超时: {timeout}秒")
    
    # 1. 检查附件文件
    if not attachment_file or not os.path.exists(attachment_file):
        log_msg(f"附件文件不存在: {attachment_file}", "ERROR")
        return None
    
    # 2. 读取附件数据
    try:
        with open(attachment_file, 'r', encoding='utf-8') as f:
            attachment_data = json.load(f)
        log_msg(f"附件数据加载成功: {os.path.basename(attachment_file)}")
        log_msg(f"文件字节数: {len(json.dumps(attachment_data, ensure_ascii=False, indent=None).encode('utf-8'))} 字节 | 内容长度(~Token): {len(json.dumps(attachment_data, ensure_ascii=False, indent=None))}")
    except Exception as e:
        log_msg(f"读取附件失败: {e}", "ERROR")
        return None
    
    # 3. 读取提示词
    prompt_content = ""
    if prompt_file and os.path.exists(prompt_file):
        try:
            with open(prompt_file, 'r', encoding='utf-8') as f:
                prompt_content = f.read()
            log_msg(f"提示词加载成功: {os.path.basename(prompt_file)}")
            log_msg(f"提示词长度: {len(prompt_content)} 字符")
        except Exception as e:
            log_msg(f"读取提示词失败: {e}", "WARNING")
    
    # 4. 构建消息
    messages = []
    
    # 系统提示
    system_prompt = prompt_content if prompt_content else "你是一个专业的团队周报优化助手，擅长提炼和总结团队工作内容。"
    messages.append({"role": "system", "content": system_prompt})
    
    # 用户消息 - 将附件数据转换为文本
    # 注意：这里不是上传文件，而是将文件内容作为文本发送
    data_str = json.dumps(attachment_data, ensure_ascii=False, indent=2)
    
    # 如果数据太大，可能需要截断，但周报数据通常不大
    if len(data_str) > 100000:  # 10万字符
        log_msg(f"数据较大 ({len(data_str)} 字符)，可能超过token限制", "WARNING")
    
    user_content = f"请根据以下团队周报数据，生成一份优化后的团队周报汇总：\n\n{data_str}"
    messages.append({"role": "user", "content": user_content})
    
    # 5. 获取API客户端并调用
    try:
        async with get_api_client(platform) as client:
            # 检查API是否可用
            if not await client.check_api_available():
                log_msg("API不可用（密钥未配置或无效）", "WARNING")
                return None
            
            # 设置超时
            client.timeout = timeout
            
            # 调用API
            log_msg("正在调用API...")
            start_time = datetime.now()
            
            response_text = await client.call_api(messages, temperature)
            
            elapsed = (datetime.now() - start_time).total_seconds()
            
            if not response_text:
                log_msg(f"API返回为空 (耗时: {elapsed:.1f}秒)", "ERROR")
                return None
            
            log_msg(f"API调用成功，响应长度: {len(response_text)} 字符 (耗时: {elapsed:.1f}秒)")
            
            # 6. 尝试解析JSON响应
            try:
                # 查找JSON部分（如果响应包含markdown代码块）
                json_str = response_text
                
                # 提取 ```json ... ``` 或 ``` ... ``` 中的内容
                json_pattern = r'```(?:json)?\n(.*?)\n```'
                match = re.search(json_pattern, response_text, re.DOTALL)
                if match:
                    json_str = match.group(1)
                    log_msg("从markdown代码块中提取JSON")
                
                # 清理可能的额外文本
                json_str = json_str.strip()
                
                # 尝试解析JSON
                if json_str.startswith('{') and json_str.endswith('}'):
                    result_data = json.loads(json_str)
                    log_msg("响应解析为JSON成功")
                elif json_str.startswith('[') and json_str.endswith(']'):
                    # 如果是数组，包装成对象
                    array_data = json.loads(json_str)
                    result_data = {"optimized_content": array_data}
                    log_msg("响应解析为JSON数组")
                else:
                    # 不是JSON，返回原始文本
                    result_data = {
                        "optimized_content": response_text,
                        "raw_response": response_text,
                        "format": "text"
                    }
                    log_msg("响应不是JSON格式，返回原始文本")
                
                return result_data
                
            except json.JSONDecodeError as e:
                log_msg(f"JSON解析失败: {e}", "WARNING")
                # 返回原始文本
                return {
                    "optimized_content": response_text,
                    "raw_response": response_text,
                    "parse_error": str(e)
                }
            
    except Exception as e:
        log_msg(f"API调用过程异常: {e}", "ERROR")
        import traceback
        traceback.print_exc()
        return None


async def main():
    """测试用主函数"""
    import argparse
    
    parser = argparse.ArgumentParser(description="LLM API 测试工具")
    parser.add_argument("file", nargs="?", help="附件文件路径")
    parser.add_argument("-p", "--platform", default="BS", choices=["BS","DS", "DB", "KM"], help="平台 (默认: BS)")
    parser.add_argument("-P", "--prompt", help="提示词文件路径")
    parser.add_argument("-t", "--timeout", type=int, default=300, help="超时时间")
    parser.add_argument("-T", "--temperature", type=float, help="温度参数")
    parser.add_argument("--debug", action="store_true", help="调试模式")
    
    args = parser.parse_args()
    
    if args.file:
        result = await llm_api_resp(
            attachment_file=args.file,
            platform=args.platform,
            prompt_file=args.prompt,
            timeout=args.timeout,
            temperature=args.temperature
        )
        
        if result:
            print("\n" + "="*60)
            print("✅ API调用成功，结果预览:")
            print("="*60)
            
            # 智能预览
            if "optimized_content" in result:
                preview = result["optimized_content"]
                if len(preview) > 1000:
                    preview = preview[:1000] + "...\n[内容已截断]"
                print(preview)
            else:
                preview = json.dumps(result, ensure_ascii=False, indent=2)
                if len(preview) > 1000:
                    preview = preview[:1000] + "...\n[内容已截断]"
                print(preview)
            
            # 保存结果到文件（可选）
            output_file = f"api_response_{args.platform}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(result, f, ensure_ascii=False, indent=2)
            print(f"\n📁 完整结果已保存到: {output_file}")
        else:
            print("\n❌ API调用失败")
    else:
        parser.print_help()


if __name__ == "__main__":
    asyncio.run(main())