#!/usr/bin/env python3
"""
周报整合脚本 - 基于V4生成的摘要文件
整合团队周报并生成个人周报
"""

import json
import glob
from datetime import datetime

def load_summaries(pattern='*_摘要.json'):
    """加载所有摘要文件"""
    summaries = []
    for file in glob.glob(pattern):
        try:
            with open(file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                summaries.append(data)
                print(f"  ✓ 加载: {file} -> {data.get('name', 'Unknown')}")
        except Exception as e:
            print(f"  ✗ 失败: {file} - {e}")
    return summaries

def generate_team_weekly(summaries, output_file='team_weekly.md'):
    """生成团队周报 Markdown"""
    
    lines = [
        "# 团队周报汇总",
        f"",
        f"**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        f"**团队成员数**: {len(summaries)}",
        "",
        "---",
        ""
    ]
    
    for summary in summaries:
        name = summary.get('name', 'Unknown')
        date = summary.get('date', 'Unknown')
        data = summary.get('summary', {})
        
        lines.extend([
            f"## {name}",
            f"",
            f"**周期**: {data.get('week_period', date)}",
            f"",
            "### OKR进展",
            ""
        ])
        
        # OKR摘要
        for okr in data.get('okr_summary', []):
            obj = okr.get('objective', '无目标')
            lines.append(f"- **{obj}**")
            for kr in okr.get('key_results', []):
                lines.append(f"  - {kr}")
            lines.append("")
        
        if not data.get('okr_summary'):
            lines.append("- 暂无OKR数据")
            lines.append("")
        
        # 工作亮点
        lines.extend([
            "### 本周工作亮点",
            ""
        ])
        
        highlights = data.get('work_highlights', [])
        if highlights:
            for item in highlights[:5]:  # 最多5条
                lines.append(f"- {item}")
        else:
            lines.append("- 暂无数据")
        
        lines.extend([
            "",
            "### 下周重点",
            ""
        ])
        
        plans = data.get('next_week_focus', [])
        if plans:
            for item in plans[:3]:  # 最多3条
                lines.append(f"- {item}")
        else:
            lines.append("- 暂无数据")
        
        lines.extend([
            "",
            "---",
            ""
        ])
    
    content = '\n'.join(lines)
    
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f"\n✓ 团队周报已保存: {output_file}")
    print(f"  总字符数: {len(content)}")
    
    return content

def generate_prompt_for_llm(summaries):
    """生成给大模型的提示词"""
    
    prompt_lines = [
        "基于以下团队成员的周报摘要，请生成一份团队周报汇总：",
        "",
        "要求：",
        "1. 按人员分别总结每个人的本周工作",
        "2. 提取团队整体的关键进展",
        "3. 识别需要关注的风险或问题",
        "4. 格式清晰，便于阅读",
        "",
        "---",
        "",
        "团队数据：",
        ""
    ]
    
    for summary in summaries:
        name = summary.get('name', 'Unknown')
        data = summary.get('summary', {})
        
        prompt_lines.extend([
            f"## {name}",
            "",
            "OKR进展："
        ])
        
        for okr in data.get('okr_summary', []):
            prompt_lines.append(f"- 目标: {okr.get('objective', '')}")
            for kr in okr.get('key_results', []):
                prompt_lines.append(f"  - {kr}")
        
        prompt_lines.extend([
            "",
            "本周工作："
        ])
        
        for item in data.get('work_highlights', [])[:3]:
            prompt_lines.append(f"- {item}")
        
        prompt_lines.extend([
            "",
            "下周计划："
        ])
        
        for item in data.get('next_week_focus', [])[:2]:
            prompt_lines.append(f"- {item}")
        
        prompt_lines.append("")
    
    prompt = '\n'.join(prompt_lines)
    
    # 保存提示词
    with open('llm_prompt.txt', 'w', encoding='utf-8') as f:
        f.write(prompt)
    
    print(f"\n✓ 大模型提示词已保存: llm_prompt.txt")
    print(f"  字符数: {len(prompt)} (适合直接发送给大模型)")
    
    return prompt

def main():
    print("="*70)
    print(" "*20 + "周报整合工具")
    print("="*70)
    
    # 1. 加载所有摘要
    print("\n[步骤1] 加载摘要文件...")
    summaries = load_summaries()
    
    if not summaries:
        print("\n✗ 未找到摘要文件，请先运行 fetch_okr_v4_enhanced.py")
        return
    
    print(f"\n✓ 共加载 {len(summaries)} 份周报")
    
    # 2. 生成团队周报
    print("\n[步骤2] 生成团队周报...")
    team_content = generate_team_weekly(summaries)
    
    # 3. 生成大模型提示词
    print("\n[步骤3] 生成大模型提示词...")
    prompt = generate_prompt_for_llm(summaries)
    
    # 4. 统计信息
    print("\n" + "="*70)
    print("整合完成!")
    print("="*70)
    print(f"\n生成的文件：")
    print(f"  1. team_weekly.md - 团队周报（可直接阅读）")
    print(f"  2. llm_prompt.txt - 大模型提示词（用于AI生成）")
    print(f"\n数据量统计：")
    print(f"  原始数据（估算）: {len(summaries) * 50000} 字符")
    print(f"  处理后数据: {len(team_content) + len(prompt)} 字符")
    print(f"  压缩率: {(len(team_content) + len(prompt)) / (len(summaries) * 50000) * 100:.1f}%")
    print(f"\n使用建议：")
    print(f"  - 直接阅读: 查看 team_weekly.md")
    print(f"  - AI生成: 将 llm_prompt.txt 发送给大模型")

if __name__ == '__main__':
    main()
