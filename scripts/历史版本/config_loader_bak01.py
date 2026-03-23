#!/usr/bin/env python3
"""
统一配置加载器 - 供所有脚本使用
"""

import json
import os
from pathlib import Path
from typing import Optional, Dict, Any, List
import sys
import platform
import subprocess
import traceback

sys.stderr.flush()

class ConfigLoader:
    """统一的配置加载器 - 保持各脚本独立性"""
    
    _instance = None  # 单例模式
    
    def __new__(cls, script_name: str = None):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self, script_name: str = None):
        """
        初始化配置加载器
        :param script_name: 调用脚本名称（用于日志）
        """
        if hasattr(self, 'initialized'):
            return
        
        try:
            # 判断是开发环境还是打包后的EXE
            if getattr(sys, 'frozen', False):
                # 打包成EXE后的运行路径
                base_path = os.path.dirname(sys.executable)
                print(f"运行模式: 打包EXE, 基础路径: {base_path}", file=sys.stderr)
            else:
                # 开发环境下的运行路径
                base_path = os.path.dirname(os.path.abspath(__file__))
                print(f"运行模式: 开发环境, 基础路径: {base_path}", file=sys.stderr)
            
            self.script_dir = Path(base_path).absolute()
            print(f"脚本目录: {self.script_dir}", file=sys.stderr)
            
            # ⭐ 重要：项目根目录就是脚本目录（对于EXE来说）
            self.project_root = self.script_dir
            print(f"项目根目录: {self.project_root}", file=sys.stderr)
            
            # 查找配置文件
            self.config_file = self._find_config_file()
            print(f"配置文件路径: {self.config_file}", file=sys.stderr)
            
            self.script_name = script_name or Path(sys.argv[0]).stem
            print(f"脚本名称: {self.script_name}", file=sys.stderr)
            
            self.config = self._load_config()
            print("配置文件加载成功", file=sys.stderr)
            print(f"配置键: {list(self.config.keys())}", file=sys.stderr)  # 调试输出
            
            self.initialized = True
            print("ConfigLoader 初始化完成", file=sys.stderr)
            
        except Exception as e:
            print(f"ConfigLoader.__init__ 出错: {e}", file=sys.stderr)
            traceback.print_exc(file=sys.stderr)
            raise
    
    def _load_config(self) -> Dict[str, Any]:
        """加载配置文件"""
        if not self.config_file.exists():
            # 如果配置文件不存在，尝试复制模板
            template_file = self.script_dir / "config" / "config.template.json"
            if template_file.exists():
                import shutil
                shutil.copy2(template_file, self.config_file)
                print(f"[{self.script_name}] 已从模板创建配置文件: {self.config_file}")
                print(f"[{self.script_name}] 请编辑配置文件后重新运行")
            else:
                raise FileNotFoundError(
                    f"[{self.script_name}] 配置文件不存在: {self.config_file}\n"
                    f"请创建配置文件或复制 config.template.json"
                )
            sys.exit(1)
        
        with open(self.config_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def get_path(self, key: str) -> Path:
        """获取路径（自动转换为绝对路径）"""
        path_str = self.config["paths"].get(key)
        if not path_str:
            raise KeyError(f"[{self.script_name}] 配置中缺少路径: {key}")
        
        # 相对路径基于项目根目录
        return self.project_root / path_str
    
    def get_history_file(self, name: str, platform: str = None) -> Path:
        """获取历史记录文件路径"""
        history_dir = self.get_path("history_dir")
        history_dir.mkdir(parents=True, exist_ok=True)  # ← 这里会创建目录
        
        # 从配置获取文件模板
        files_config = self.config.get("history", {}).get("files", {})
        template = files_config.get(name)
        
        if not template:
            # 如果配置中没有，使用默认模板
            if name == "copy_button":
                template = "copy_button_{platform}.json"
            elif name == "upload_icon":
                template = "upload_icon_{platform}.json"
            else:
                template = f"{name}_{{platform}}.json"
        
        # 替换平台占位符
        if platform and "{platform}" in template:
            filename = template.format(platform=platform)
        else:
            filename = template
        
        file_path = history_dir / filename
        return file_path
    
    def get_prompt_file(self, name: str) -> Path:
        """获取prompt文件路径"""
        # 如果name是绝对路径，直接返回
        if Path(name).is_absolute():
            return Path(name)
        
        # 如果name已经是相对路径（包含config/prompts），直接基于项目根目录拼接
        if name.startswith("config/prompts/") or name.startswith("config\\prompts\\"):
            return self.project_root / name
        
        # 否则，拼接到prompts_dir
        prompts_dir = self.get_path("prompts_dir")
        prompts_dir.mkdir(parents=True, exist_ok=True)
        return prompts_dir / name
    
    def get_browser_state_file(self) -> Path:
        """获取浏览器状态文件路径"""
        # browser_state.json 放在项目根目录的 config/ 下
        return self.project_root / self.config["browser"]["state_file"]
    
    def get_platform_config(self, platform: str) -> Dict[str, Any]:
        """获取平台完整配置"""
        platforms = self.config.get("llm_platforms", {})
        config = platforms.get(platform, {})
        
        # 确保返回的配置包含必要的字段
        if not config:
            # 返回默认配置
            return {
                "name": platform,
                "type": "web",
                "url": self._get_default_url(platform),
                "web_config": {
                    "input_selectors": [
                        'textarea[placeholder*="输入"]',
                        'textarea[placeholder*="提问"]',
                        'textarea:visible'
                    ],
                    "submit_method": "enter"
                }
            }
        
        # 确保有 web_config（如果是 web 类型）
        if config.get("type") == "web" and "web_config" not in config:
            config["web_config"] = {
                "input_selectors": [
                    'textarea[placeholder*="输入"]',
                    'textarea[placeholder*="提问"]',
                    'textarea:visible'
                ],
                "submit_method": "enter"
            }
        
        return config
    
    def _get_default_url(self, platform: str) -> str:
        """获取默认平台URL"""
        defaults = {
            "DS": "https://chat.deepseek.com",
            "DB": "https://www.doubao.com/chat",
            "KM": "https://kimi.moonshot.cn"
        }
        return defaults.get(platform, "")
    
    def get_platform_url(self, platform: str) -> str:
        """获取平台URL"""
        config = self.get_platform_config(platform)
        return config.get("url", self._get_default_url(platform))
    
    def get_platform_web_config(self, platform: str) -> Dict[str, Any]:
        """获取平台Web自动化配置"""
        config = self.get_platform_config(platform)
        
        # 如果配置中有 web_config，直接使用
        if "web_config" in config:
            return config["web_config"]
        
        # 否则返回默认配置
        return {
            "input_selectors": [
                'textarea[placeholder*="输入"]',
                'textarea[placeholder*="提问"]',
                'textarea:visible'
            ],
            "submit_method": "enter",
            "timeout": 300,
            "wait_for_stable": 8,
            "check_interval": 1.0
        }
    
    def _find_config_file(self) -> Path:
        """智能查找配置文件"""
        possible_locations = [
            # EXE同级目录的 config/config.json
            self.project_root / "config" / "config.json",
            # EXE同级目录的 config.json
            self.project_root / "config.json",
            # 脚本所在目录的 config/config.json
            self.script_dir / "config" / "config.json",
        ]
        
        for loc in possible_locations:
            if loc.exists():
                print(f"[调试] 找到配置文件: {loc}", file=sys.stderr)
                return loc
        
        # 如果都没找到，创建默认配置
        default_loc = self.project_root / "config" / "config.json"
        print(f"[调试] 未找到配置文件，将创建默认配置: {default_loc}", file=sys.stderr)
        
        # 创建默认配置
        self._create_default_config(default_loc)
        
        return default_loc

    def _create_default_config(self, config_path: Path):
        """创建默认配置文件"""
        default_config = {
            "version": "1.0.0",
            "user": {
                "my_id": "user001",
                "staff_ids": ["staff001", "staff002"]
            },
            "weekly_report": {  # 注意这里用的是下划线
                "target_week": 1,
                "optimize_prompt": "optimize_prompt.txt",
                "default_platform": "DS"
            },
            "paths": {
                "browser_debug_dir": "browser_debug",
                "history_dir": "config/history",
                "prompts_dir": "config/prompts",
                "llm_response_dir": "llmRsp"
            },
            "history": {
                "files": {
                    "copy_button": "copy_button_{platform}.json",
                    "upload_icon": "upload_icon_{platform}.json"
                }
            },
            "browser": {
                "state_file": "config/browser_state.json",
                "mode": "headful",
                "headless": False,
                "screenshot_on_error": True,
                "chrome_paths": {
                    "windows": [
                        "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
                        "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe"
                    ],
                    "linux": "/usr/bin/google-chrome",
                    "mac": "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
                }
            },
            "llm_platforms": {
                "DS": {
                    "name": "DeepSeek",
                    "type": "web",
                    "url": "https://chat.deepseek.com",
                    "web_config": {
                        "input_selectors": [
                            'textarea[placeholder*="输入"]',
                            'textarea[placeholder*="提问"]',
                            'textarea:visible'
                        ],
                        "submit_method": "enter",
                        "timeout": 300,
                        "wait_for_stable": 8,
                        "check_interval": 1.0
                    }
                }
            },
            "logging": {
                "debug": True,
                "save_intermediate_files": True
            }
        }
        
        config_path.parent.mkdir(parents=True, exist_ok=True)
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(default_config, f, ensure_ascii=False, indent=2)
        print(f"[调试] 已创建默认配置文件: {config_path}", file=sys.stderr)
  
    @property
    def browser_mode(self) -> str:
        """获取浏览器模式：headless/headful/auto"""
        return self.config["browser"].get("mode", "headful")
    
    @property
    def should_use_headless(self) -> bool:
        """
        判断是否应该使用无头模式
        基于配置决定
        """
        mode = self.browser_mode
        
        if mode == "headless":
            return True
        elif mode == "headful":
            return False
        elif mode == "auto":
            # auto模式：如果有已保存的浏览器状态且进程存在，用无头；否则用有头
            state_file = self.get_browser_state_file()
            if state_file.exists():
                try:
                    with open(state_file, 'r', encoding='utf-8') as f:
                        state = json.load(f)
                    pid = state.get("browser_pid")
                    if self._is_process_running(pid):
                        return True
                except:
                    pass
            return False
        
        return False  # 默认有头
    
    def _is_process_running(self, pid):
        """检查进程是否在运行"""
        if not pid:
            return False
        try:
            if platform.system() == "Windows":
                result = subprocess.run(f'tasklist /FI "PID eq {pid}"', 
                                       shell=True, capture_output=True, text=True)
                return str(pid) in result.stdout
            else:
                os.kill(pid, 0)
                return True
        except:
            return False
    
    @property
    def my_id(self) -> str:
        """获取我的ID"""
        user_config = self.config.get("user", {})
        return user_config.get("my_id", "")
    
    @property
    def staff_ids(self) -> List[str]:
        """获取员工ID列表"""
        user_config = self.config.get("user", {})
        return user_config.get("staff_ids", [])
    
    @property
    def target_week(self) -> int:
        """获取目标周数"""
        # 尝试多种可能的键名
        weekly_report = self.config.get("weekly_report") or self.config.get("weeklyReport") or {}
        return weekly_report.get("target_week", 1)
    
    @property
    def optimize_prompt(self) -> Path:
        """获取优化提示词文件路径"""
        # 尝试多种可能的键名
        weekly_report = self.config.get("weekly_report") or self.config.get("weeklyReport") or {}
        prompt_name = weekly_report.get("optimize_prompt", "optimize_prompt.txt")
        return self.get_prompt_file(prompt_name)
    
    @property
    def default_platform(self) -> str:
        """获取默认平台"""
        # 尝试多种可能的键名
        weekly_report = self.config.get("weekly_report") or self.config.get("weeklyReport") or {}
        return weekly_report.get("default_platform", "DS")
    
    @property
    def debug(self) -> bool:
        """是否调试模式"""
        logging_config = self.config.get("logging", {})
        return logging_config.get("debug", True)
    
    @property
    def save_intermediate(self) -> bool:
        """是否保存中间文件"""
        logging_config = self.config.get("logging", {})
        return logging_config.get("save_intermediate_files", True)
    
    @property
    def screenshot_on_error(self) -> bool:
        """错误时是否截图"""
        browser_config = self.config.get("browser", {})
        return browser_config.get("screenshot_on_error", True)