#!/usr/bin/env python3
"""
统一配置加载器 - 支持开发环境和打包环境
"""

import json
import os
import sys
import platform
import subprocess
import traceback
from pathlib import Path
from typing import Optional, Dict, Any, List

class ConfigLoader:
    """统一的配置加载器 - 支持打包环境"""
    
    _instance = None
    
    def __new__(cls, script_name: str = None):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self, script_name: str = None):
        if hasattr(self, 'initialized'):
            return

        # Debug flag: defaults to True so early diagnostic prints show during development.
        # Will be updated from config after config is loaded.
        self._debug = False

        try:
            # 判断运行环境
            self.is_frozen = getattr(sys, 'frozen', False)
            self.script_name = script_name or Path(sys.argv[0]).stem

            # 确定基础路径
            if self.is_frozen:
                # 打包后：
                # - base_dir: EXE所在目录（用于写入数据，如日志、输出）
                # - resource_dir: sys._MEIPASS（PyInstaller解压的内部资源目录，用于读取打包的资源）
                self.base_dir = Path(sys.executable).parent.absolute()
                self.resource_dir = Path(getattr(sys, '_MEIPASS', self.base_dir))
                if self._debug:
                    print(f"[{self.script_name}] 运行模式: 打包EXE", file=sys.stderr)
                    print(f"[{self.script_name}] EXE目录: {self.base_dir}", file=sys.stderr)
                    print(f"[{self.script_name}] 资源目录: {self.resource_dir}", file=sys.stderr)
            else:
                # 开发环境：脚本所在目录
                self.base_dir = Path(__file__).parent.absolute()
                self.resource_dir = self.base_dir
                if self._debug:
                    print(f"[{self.script_name}] 运行模式: 开发环境", file=sys.stderr)
                    print(f"[{self.script_name}] 脚本目录: {self.base_dir}", file=sys.stderr)
            
            # 为了兼容旧代码，设置 script_dir 属性
            self.script_dir = self.base_dir
            
            # 查找项目根目录（包含config目录的目录）
            self.project_root = self._find_project_root()
            if self._debug:
                print(f"[{self.script_name}] 项目根目录: {self.project_root}", file=sys.stderr)

            # 查找配置文件
            self.config_file = self._find_config_file()
            if self._debug:
                print(f"[{self.script_name}] 配置文件: {self.config_file}", file=sys.stderr)

            # 加载配置
            self.config = self._load_config()

            # Update debug flag from config now that config is loaded
            self._debug = self.config.get("logging", {}).get("debug", False)

            if self._debug:
                print(f"[{self.script_name}] 配置加载成功", file=sys.stderr)
                print(f"[{self.script_name}] 配置键: {list(self.config.keys())}", file=sys.stderr)
            
            # 确保必要的目录存在
            self._ensure_directories()

            # ⭐ 添加路径验证
            weekly_report_dir = self.get_path("weekly_report_dir", create=False)
            if self._debug:
                print(f"[{self.script_name}] 周报目录: {weekly_report_dir}", file=sys.stderr)
                print(f"[{self.script_name}] 周报目录是否存在: {weekly_report_dir.exists()}", file=sys.stderr)

            self.initialized = True
            if self._debug:
                print(f"[{self.script_name}] ConfigLoader 初始化完成", file=sys.stderr)
            
        except Exception as e:
            print(f"[{self.script_name}] 初始化出错: {e}", file=sys.stderr)
            traceback.print_exc(file=sys.stderr)
            raise
    
    def _find_project_root(self) -> Path:
        """查找项目根目录（包含config目录的目录）

        打包后优先使用 EXE 所在目录（base_dir），这样：
        - 用户编辑的 config/config.json 能被读取
        - browser_state.json 等运行时数据写在 EXE 目录下，更新时不丢失
        - 只有当 EXE 目录没有 config/ 时，才回退到 _internal/（resource_dir）
        """

        if self.is_frozen:
            # ⭐ 优先使用 EXE 目录（base_dir）：用户配置和运行时数据应在此
            if (self.base_dir / "config").exists():
                if self._debug:
                    print(f"[{self.script_name}] 打包环境：使用EXE目录 {self.base_dir}", file=sys.stderr)
                return self.base_dir

            # 从EXE目录向上查找
            current = self.base_dir.parent
            for _ in range(2):
                if (current / "config").exists():
                    if self._debug:
                        print(f"[{self.script_name}] 打包环境：找到外部配置目录 {current}", file=sys.stderr)
                    return current
                parent = current.parent
                if parent == current:
                    break
                current = parent

            # 回退到 _internal/（打包的资源目录）
            if (self.resource_dir / "config").exists():
                if self._debug:
                    print(f"[{self.script_name}] 打包环境：回退到资源目录 {self.resource_dir}", file=sys.stderr)
                return self.resource_dir

            if self._debug:
                print(f"[{self.script_name}] 打包环境：使用EXE目录（默认） {self.base_dir}", file=sys.stderr)
            return self.base_dir
        else:
            # 开发环境：脚本在 scripts/，项目根目录是上一级
            parent_dir = self.base_dir.parent
            if (parent_dir / "config").exists():
                if self._debug:
                    print(f"[{self.script_name}] 开发环境：找到项目根目录 {parent_dir}", file=sys.stderr)
                return parent_dir
            if self._debug:
                print(f"[{self.script_name}] 开发环境：使用默认项目根目录 {self.base_dir}", file=sys.stderr)
            return self.base_dir
    
    def _find_config_file(self) -> Path:
        """智能查找配置文件"""
        
        # 可能的配置文件位置（按优先级）
        possible_locations = [
            # 1. 项目根目录的 config/config.json
            self.project_root / "config" / "config.json",
            # 2. 当前目录的 config/config.json
            self.base_dir / "config" / "config.json",
            # 3. 项目根目录的 config.json
            self.project_root / "config.json",
            # 4. 上一级目录的 config/config.json
            self.base_dir.parent / "config" / "config.json",
            # 5. 脚本所在目录的 config/config.json（兼容旧代码）
            self.script_dir / "config" / "config.json",
        ]
        
        # 添加调试信息
        if self._debug:
            print(f"[{self.script_name}] 搜索配置文件:", file=sys.stderr)
        for loc in possible_locations:
            exists = "✓" if loc.exists() else "✗"
            if self._debug:
                print(f"  [{exists}] {loc}", file=sys.stderr)
            if loc.exists():
                return loc
        
        # 如果都不存在，在项目根目录创建默认配置
        default_loc = self.project_root / "config" / "config.json"
        print(f"[{self.script_name}] 未找到配置文件，将在以下位置创建默认配置: {default_loc}", file=sys.stderr)
        
        default_loc.parent.mkdir(parents=True, exist_ok=True)
        self._create_default_config(default_loc)
        
        return default_loc
    
    def _load_config(self) -> Dict[str, Any]:
        """加载配置文件"""
        if not self.config_file.exists():
            # 如果配置文件不存在，尝试复制模板
            template_file = self.project_root / "config" / "config.template.json"
            if template_file.exists():
                import shutil
                shutil.copy2(template_file, self.config_file)
                print(f"[{self.script_name}] 已从模板创建配置文件: {self.config_file}")
                print(f"[{self.script_name}] 请编辑配置文件后重新运行")
            else:
                # 如果没有模板，创建默认配置
                self._create_default_config(self.config_file)
                print(f"[{self.script_name}] 已创建默认配置文件: {self.config_file}")
                print(f"[{self.script_name}] 请编辑配置文件后重新运行")
            sys.exit(1)
        
        with open(self.config_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def _ensure_directories(self):
        """确保必要的目录存在"""
        try:
            # 从配置中获取需要创建的目录
            dirs_to_create = [
                self.get_path("history_dir", create=False),
                self.get_path("llm_response_dir", create=False),
                self.get_path("browser_debug_dir", create=False),
            ]
            
            for dir_path in dirs_to_create:
                try:
                    dir_path.mkdir(parents=True, exist_ok=True)
                    if self._debug:
                        print(f"[{self.script_name}] 确保目录存在: {dir_path}", file=sys.stderr)
                except Exception as e:
                    print(f"[{self.script_name}] 创建目录失败 {dir_path}: {e}", file=sys.stderr)
        except Exception as e:
            print(f"[{self.script_name}] 确保目录存在时出错: {e}", file=sys.stderr)
    
    def get_path(self, key: str, create: bool = True) -> Path:
        """获取路径
        
        :param key: 配置中的路径键名
        :param create: 是否自动创建目录
        """
        # 路径键名映射（处理可能的命名差异）
        key_mapping = {
            "weekly_report_dir": "weekly_report_dir",
            "llm_response_dir": "llm_response_dir",
            "browser_debug_dir": "browser_debug_dir",
            "prompts_dir": "prompts_dir",
            "history_dir": "history_dir",
        }
        
        # ⭐ 重要：修改默认路径，去掉 scripts/ 前缀
        default_paths = {
            "weekly_report_dir": "WeeklyReport",           # 原来是 "scripts/WeeklyReport"
            "llm_response_dir": "llmRsp",                  # 原来是 "scripts/llmRsp"
            "browser_debug_dir": "browser_debug",          # 原来是 "scripts/browser_debug"
            "prompts_dir": "config/prompts",
            "history_dir": "config/history",
        }
        
        actual_key = key_mapping.get(key, key)
        
        # 从配置获取路径
        paths_config = self.config.get("paths", {})
        path_str = paths_config.get(actual_key)
        
        # 如果配置中没有，使用默认值
        if not path_str:
            path_str = default_paths.get(key)
            if path_str:
                if self._debug:
                    print(f"[{self.script_name}] 使用默认路径: {key} -> {path_str}", file=sys.stderr)
            else:
                raise KeyError(f"[{self.script_name}] 配置中缺少路径: {key}")
        
        # 构建绝对路径
        full_path = self.project_root / path_str

        # 打包环境下，只读资源（如prompts）可能仅存在于 _internal/ 中
        # 如果 project_root 下不存在，回退到 resource_dir
        if self.is_frozen and not full_path.exists() and self.resource_dir != self.project_root:
            fallback_path = self.resource_dir / path_str
            if fallback_path.exists():
                if self._debug:
                    print(f"[{self.script_name}] 路径 {key}: 回退到资源目录 {fallback_path}", file=sys.stderr)
                return fallback_path

        if self._debug:
            print(f"[{self.script_name}] 路径 {key} = {full_path}", file=sys.stderr)
        
        # 自动创建目录
        if create and key.endswith("_dir") and not full_path.suffix:
            try:
                full_path.mkdir(parents=True, exist_ok=True)
                if self._debug:
                    print(f"[{self.script_name}] 确保目录存在: {full_path}", file=sys.stderr)
            except Exception as e:
                print(f"[{self.script_name}] 创建目录失败 {full_path}: {e}", file=sys.stderr)
        
        return full_path
    
    def get_history_file(self, name: str, platform: str = None) -> Path:
        """获取历史记录文件路径"""
        history_dir = self.get_path("history_dir")
        
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
            path = self.project_root / name
            # 打包环境：project_root下找不到时回退到 resource_dir
            if self.is_frozen and not path.exists() and self.resource_dir != self.project_root:
                fallback = self.resource_dir / name
                if fallback.exists():
                    return fallback
            return path

        # 否则，拼接到prompts_dir
        prompts_dir = self.get_path("prompts_dir")
        return prompts_dir / name
    
    def get_browser_state_file(self) -> Path:
        """获取浏览器状态文件路径"""
        # browser_state.json 放在项目根目录的 config/ 下
        return self.project_root / self.config.get("browser", {}).get("state_file", "config/browser_state.json")
    
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
            "BO": "https://chat.deepseek.com",
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

    def get_chrome_path(self) -> Optional[str]:
        """获取当前系统的Chrome路径"""
        system = platform.system()

        browser_config = self.config.get("browser", {})
        chrome_paths = browser_config.get("chrome_paths", {})

        if system == "Windows":
            paths = chrome_paths.get("windows", [])
            for path in paths:
                expanded = os.path.expandvars(path)
                if os.path.exists(expanded):
                    return expanded
        elif system == "Linux":
            return chrome_paths.get("linux")
        elif system == "Darwin":
            return chrome_paths.get("mac")

        return None

    def get_browser_state_file(self) -> Path:
        """获取浏览器状态文件路径"""
        browser_config = self.config.get("browser", {})
        state_file = browser_config.get("state_file", "config/browser_state.json")
        return self.project_root / state_file

    @property
    def browser_mode(self) -> str:
        """获取浏览器模式：headless/headful/auto"""
        return self.config.get("browser", {}).get("mode", "headful")

    @property
    def should_use_headless(self) -> bool:
        """判断是否应该使用无头模式"""
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
    
    def _create_default_config(self, config_path: Path):
        """创建默认配置文件"""
        default_config = {
            "version": "1.0.0",
            "user": {
                "my_id": "user001",
                "staff_ids": ["staff001", "staff002"]
            },
            "paths": {
                # ⭐ 修改这些路径，去掉 scripts/ 前缀
                "weekly_report_dir": "WeeklyReport",           # 原来是 "scripts/WeeklyReport"
                "llm_response_dir": "llmRsp",                  # 原来是 "scripts/llmRsp"
                "browser_debug_dir": "browser_debug",          # 原来是 "scripts/browser_debug"
                "prompts_dir": "config/prompts",
                "history_dir": "config/history"
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
            "weekly_report": {
                "target_week": 1,
                "optimize_prompt": "WeeklyReportOptimizePrompt.md",
                "default_platform": "DS"
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
            "history": {
                "max_records": 8,
                "files": {
                    "copy_button": "copy_button_{platform}.json",
                    "upload_icon": "upload_icon_{platform}.json"
                }
            },
            "logging": {
                "debug": False,
                "save_intermediate_files": True,
                "log_level": "DEBUG"
            }
        }
        
        config_path.parent.mkdir(parents=True, exist_ok=True)
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(default_config, f, ensure_ascii=False, indent=2)
        print(f"[{self.script_name}] 已创建默认配置文件: {config_path}", file=sys.stderr)

    @property
    def browser_mode(self) -> str:
        """获取浏览器模式：headless/headful/auto"""
        return self.config.get("browser", {}).get("mode", "headful")

    @property
    def debug_port_range(self) -> list:
        """获取浏览器调试端口范围 [start, end]"""
        return self.config.get("browser", {}).get("debug_port_range", [9222, 9299])

    @property
    def should_use_headless(self) -> bool:
        """判断是否应该使用无头模式"""
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
        
        return False
    
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
        weekly_report = self.config.get("weekly_report", {})
        return weekly_report.get("target_week", 1)
    
    @property
    def optimize_prompt(self) -> Path:
        """获取优化提示词文件路径"""
        weekly_report = self.config.get("weekly_report", {})
        prompt_name = weekly_report.get("optimize_prompt", "WeeklyReportOptimizePrompt.md")
        return self.get_prompt_file(prompt_name)
    
    @property
    def default_platform(self) -> str:
        """获取默认平台"""
        weekly_report = self.config.get("weekly_report", {})
        return weekly_report.get("default_platform", "DS")
    
    @property
    def debug(self) -> bool:
        """是否调试模式"""
        logging_config = self.config.get("logging", {})
        return logging_config.get("debug", False)
    
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

    # ==================== 日志配置 ====================
    @property
    def log_level(self) -> str:
        """获取日志级别"""
        return self.config.get("logging", {}).get("log_level", "DEBUG")

    @property
    def log_files(self) -> Dict[str, str]:
        """获取日志文件路径配置"""
        default_log_dir = os.path.join(self.project_root, "logs")
        return self.config.get("logging", {}).get("log_files", {
            "debug": os.path.join(default_log_dir, "OC_DEBUG.log"),
            "fatal": os.path.join(default_log_dir, "FATAL_DEBUG.log"),
            "error": os.path.join(default_log_dir, "OC_FATAL_ERROR.log")
        })


import sys

def safe_input(prompt: str = "", default: str = "") -> str:
    """
    安全的 input 函数，支持无控制台环境
    
    Args:
        prompt: 提示信息
        default: 默认返回值（无控制台时使用）
    
    Returns:
        用户输入或默认值
    """
    try:
        # 检查是否有可用的标准输入
        if sys.stdin is None or not sys.stdin.isatty():
            print(prompt, "(自动使用默认值)" if default else "")
            return default
        return input(prompt)
    except (EOFError, OSError, RuntimeError):
        # 无控制台环境
        print(prompt, "(自动使用默认值)" if default else "")
        return default

def safe_input_or_skip(prompt: str = "") -> str:
    """
    在无控制台环境下跳过输入，直接返回空字符串
    用于"按回车继续"类场景
    """
    return safe_input(prompt, default="")



# ==================== 统一日志管理器 ====================

class Logger:
    """
    统一日志管理器
    支持多级别日志、文件输出、控制台输出
    用法：
        logger = Logger(config)  # 或 Logger() 使用默认配置
        logger.debug("调试信息")
        logger.info("普通信息")
        logger.warning("警告信息")
        logger.error("错误信息")
        logger.fatal("致命错误")
    """

    # 日志级别映射
    LEVELS = {
        'DEBUG': 10,
        'INFO': 20,
        'WARNING': 30,
        'ERROR': 40,
        'FATAL': 50
    }

    _instance = None  # 单例模式

    def __new__(cls, config=None, **kwargs):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self, config=None, debug=None, log_files=None):
        """
        初始化日志器
        Args:
            config: ConfigLoader 实例（可选）
            debug: 是否启用调试模式（可选，默认从 config 读取）
            log_files: 日志文件路径字典（可选）
        """
        if hasattr(self, 'initialized') and self.initialized:
            return

        # 获取调试模式
        if debug is not None:
            self.debug_mode = debug
        elif config and hasattr(config, 'debug'):
            self.debug_mode = config.debug
        else:
            self.debug_mode = True  # 默认开启调试

        # 获取日志文件路径
        if log_files:
            self.log_files = log_files
        elif config and hasattr(config, 'log_files'):
            self.log_files = config.log_files
        else:
            # 默认日志路径：EXE所在目录或脚本所在目录的 logs/ 下
            if getattr(sys, 'frozen', False):
                _log_base = os.path.join(os.path.dirname(sys.executable), "logs")
            else:
                _log_base = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "logs")
            _log_base = os.path.abspath(_log_base)
            self.log_files = {
                "debug": os.path.join(_log_base, "OC_DEBUG.log"),
                "fatal": os.path.join(_log_base, "FATAL_DEBUG.log"),
                "error": os.path.join(_log_base, "OC_FATAL_ERROR.log")
            }

        self.initialized = True
        self._setup_excepthook()

    def _setup_excepthook(self):
        """设置全局异常钩子"""
        import traceback
        from datetime import datetime

        def custom_excepthook(exc_type, exc_value, exc_traceback):
            err_info = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
            self.fatal(f"检测到未捕获异常:\n{err_info}")

        sys.excepthook = custom_excepthook

    def _write_file(self, filepath: str, message: str):
        """写入日志文件"""
        try:
            log_dir = os.path.dirname(filepath)
            if log_dir:
                os.makedirs(log_dir, exist_ok=True)
            with open(filepath, "a", encoding="utf-8") as f:
                f.write(message + "\n")
        except Exception as e:
            print(f"[Logger] 写入日志文件失败: {e}", file=sys.stderr)

    def _format_message(self, level: str, message: str) -> str:
        """格式化日志消息"""
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        return f"{timestamp} [{level}] {message}"

    def debug(self, message: str):
        """调试日志"""
        if self.debug_mode:
            msg = self._format_message("DEBUG", message)
            self._write_file(self.log_files.get("debug", "OC_DEBUG.log"), msg)

    def info(self, message: str):
        """信息日志"""
        msg = self._format_message("INFO", message)
        print(f"[INFO] {message}", flush=True)

    def warning(self, message: str):
        """警告日志"""
        msg = self._format_message("WARNING", message)
        print(f"[WARNING] {message}", flush=True)

    def error(self, message: str):
        """错误日志"""
        msg = self._format_message("ERROR", message)
        self._write_file(self.log_files.get("error", "OC_FATAL_ERROR.log"), msg)
        print(f"[ERROR] {message}", flush=True)

    def fatal(self, message: str):
        """致命错误日志"""
        msg = self._format_message("FATAL", message)
        self._write_file(self.log_files.get("fatal", "FATAL_DEBUG.log"), msg)
        print(f"[FATAL] {message}", flush=True)

    # 兼容旧接口的别名
    def emergency(self, message: str):
        """紧急日志（兼容旧接口）"""
        self.error(message)

    def debug_log(self, message: str):
        """调试日志（兼容旧接口）"""
        self.debug(message)

    def fatal_debug(self, message: str):
        """致命调试日志（兼容旧接口）"""
        self.fatal(message)


# ==================== 打印函数工具 ====================

def create_print_functions(debug_mode=True):
    """
    创建打印函数（兼容旧代码）
    Returns:
        tuple: (dprint, iprint, safe_output)
    """
    import io
    import builtins
    from datetime import datetime

    def dprint(*args, force=False, **kwargs):
        """调试打印函数"""
        if force or debug_mode:
            timestamp = datetime.now().strftime("%H:%M:%S")
            print(f"[{timestamp}]", *args, **kwargs)

    def iprint(*args, **kwargs):
        """智能打印函数 - 兼容各种输出场景"""
        if sys.stdout is None or sys.stdout.closed:
            return

        msg = ' '.join(str(arg) for arg in args)

        # 替换Unicode符号
        replacements = {
            '✓': '[OK]', '✗': '[ERR]', '●': '-', '▶': '>',
            '「': '"', '」': '"', '：': ':', '；': ';', '！': '!', '？': '?',
            '\u2705': '[OK]', '\u26a0': '[WARN]', '\ufe0f': '',
        }
        for old, new in replacements.items():
            msg = msg.replace(old, new)

        kwargs['flush'] = True

        # 使用原始print
        original_print = getattr(builtins, '_original_print', print)

        try:
            original_print(msg, **kwargs)
            return
        except (UnicodeEncodeError, ValueError, BrokenPipeError):
            pass

        # 备用方案：尝试二进制写入
        try:
            if hasattr(sys.stdout, 'buffer') and sys.stdout.buffer and not sys.stdout.buffer.closed:
                sys.stdout.buffer.write(msg.encode('utf-8', errors='replace'))
                sys.stdout.buffer.write(b'\n')
                sys.stdout.buffer.flush()
                return
        except:
            pass

        # 最后尝试：强制ASCII
        try:
            safe_msg = msg.encode('ascii', errors='replace').decode('ascii')
            original_print(safe_msg, flush=True)
        except:
            pass

    def safe_output(message):
        """安全输出函数"""
        try:
            if hasattr(sys.stdout, 'buffer') and sys.stdout.buffer is not None:
                sys.stdout.buffer.write(message.encode('utf-8'))
                sys.stdout.buffer.flush()
                return
        except:
            pass

        try:
            print(message, end='', flush=True)
        except:
            try:
                sys.stdout.write(message)
                sys.stdout.flush()
            except:
                pass

    # 保存原始print
    if not hasattr(builtins, '_original_print'):
        builtins._original_print = builtins.print

    return dprint, iprint, safe_output


# ==================== 导出 ====================

__all__ = [
    'ConfigLoader',
    'Logger',
    'create_print_functions'
]