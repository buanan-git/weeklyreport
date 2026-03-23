# 团队周报自动化处理脚本使用说明

## 📋 脚本概述

`weekly_report_automation.py` 是一个串联脚本，用于自动化处理团队周报的完整流程：

1. **获取数据** - 调用 `fetch_okr_final.py` 从百度OKR系统获取团队成员周报
2. **智能优化** - 使用 `llmchat_v2.0.py` 调用大模型按照主管视角优化周报
3. **自动提交** - 调用 `submit_okr_ds_final.py` 将优化后的周报提交到系统

## 🚀 快速开始

### 基本用法

```bash
# 处理当周周报（默认）
python weekly_report_automation.py

# 处理上一周周报
python weekly_report_automation.py 1

# 处理上上周周报
python weekly_report_automation.py 2
```

### 指定大模型平台

```bash
# 使用DeepSeek平台（默认）
python weekly_report_automation.py 0 DS

# 使用豆包平台
python weekly_report_automation.py 0 DB

# 使用Kimi平台
python weekly_report_automation.py 0 KM
```

## ⚙️ 配置说明

### 目标周配置
- `0` = 当周（默认）
- `1` = 上一周  
- `2` = 上上周

### 大模型平台
- `DS` = DeepSeek（默认，推荐）
- `DB` = 豆包
- `KM` = Kimi

### 调试模式
脚本内置调试模式，会输出详细的执行信息。如需关闭，请修改脚本中的 `DEBUG = False`。

## 📁 文件结构

脚本执行后会生成以下文件：

```
WeeklyReport/
├── 团队周报_2026.3.9-3.15.json          # 原始团队周报数据
└── 团队周报_2026.3.9-3.15.llm.json       # 大模型优化后的周报数据
```

## 🔧 执行流程

### 步骤 1: 获取团队周报数据
- 调用 `fetch_okr_final.py`
- 从百度OKR系统抓取指定周的团队成员周报
- 生成原始JSON格式的团队周报文件

### 步骤 2: 调用大模型优化周报
- 调用 `llmchat_v2.0.py`
- 使用 `WeeklyReportOptimizePrompt.md` 中的优化要求
- 将原始周报从"团队成员汇总"视角转换为"主管视角"
- 生成优化后的 `.llm.json` 文件

### 步骤 3: 提交优化后的周报
- 调用 `submit_okr_ds_final.py`
- 读取优化后的周报数据
- 自动填写到百度OKR系统
- 完成周报提交

## ✅ 成功标志

脚本执行成功后会显示：
```
🎉 所有步骤执行完成！
生成的文件:
  - 原始周报: /path/to/团队周报_2026.3.9-3.15.json
  - 优化周报: /path/to/团队周报_2026.3.9-3.15.llm.json
```

## ⚠️ 注意事项

1. **网络连接**：确保能够访问百度OKR系统和选择的大模型平台
2. **浏览器环境**：脚本会自动启动和管理浏览器实例
3. **文件权限**：确保脚本有读写 `WeeklyReport` 目录的权限
4. **登录状态**：首次运行可能需要手动登录相关系统

## 🐛 故障排除

### 常见问题

1. **脚本文件不存在**
   - 确保 `fetch_okr_final.py`、`llmchat_v2.0.py`、`submit_okr_ds_final.py` 都在同一目录下

2. **大模型调用失败**
   - 检查网络连接
   - 尝试切换其他大模型平台
   - 确认 `WeeklyReportOptimizePrompt.md` 文件存在

3. **周报提交失败**
   - 检查是否已登录百度OKR系统
   - 确认优化后的周报文件格式正确

### 调试模式

开启调试模式（默认开启）可以查看详细的执行信息：
```python
# 在脚本中修改
DEBUG = True
```

## 📝 日志说明

脚本使用以下日志级别：
- `ℹ️` 信息日志 - 正常执行流程
- `✅` 成功日志 - 步骤完成
- `❌` 错误日志 - 执行失败
- `⚠️` 警告日志 - 需要注意的情况

## 🔄 中断恢复

如果脚本执行过程中被中断：
1. 检查当前执行到哪个步骤
2. 根据生成的文件判断进度
3. 重新运行脚本，会从第一步开始重新执行

## 📞 技术支持

如遇到问题，请检查：
1. 各个子脚本是否能独立正常运行
2. 网络连接状态
3. 文件权限设置
4. 浏览器环境配置