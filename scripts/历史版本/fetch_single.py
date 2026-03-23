#!/usr/bin/env python3
"""
OpenClaw Browser Relay 单员工抓取脚本
逐个抓取，每步保存，避免上下文过大
"""

import subprocess
import json
import os
from datetime import datetime

# 配置
MY_ID = 's673090'
STAFF_IDS = [
    's801573', 's733504', 's682083', 's929270', 's798268',
    's811385', 's865585', 's783729', 's400192', 's823700'
]

BASE_URL = 'https://okr.baidu-int.com/pages/weekly.html#/index?mode=steam'
DATE = 's2026-02-02'  # 上周一

OUTPUT_DIR = '../output'

class WeeklyFetcher:
    def __init__(self):
        self.current_target = None
        self.results = []
        
    def log(self, message):
        """记录日志"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        log_msg = f"[{timestamp}] {message}"
        print(log_msg)
        
        # 追加到日志文件
        with open(f'{OUTPUT_DIR}/fetch_log.txt', 'a', encoding='utf-8') as f:
            f.write(log_msg + '\n')
    
    def run_openclaw(self, args, timeout=30):
        """执行openclaw命令"""
        cmd = ['openclaw', 'browser'] + args
        try:
            result = subprocess.run(
                cmd, 
                capture_output=True, 
                text=True, 
                timeout=timeout
            )
            return result.returncode, result.stdout, result.stderr
        except subprocess.TimeoutExpired:
            return -1, '', 'timeout'
        except Exception as e:
            return -1, '', str(e)
    
    def fetch_single_staff(self, staff_id, index, total):
        """抓取单个员工的周报"""
        self.log(f"\n{'='*60}")
        self.log(f"[{index}/{total}] 开始抓取员工: {staff_id}")
        self.log(f"{'='*60}")
        
        url = f"{BASE_URL}&id={staff_id}&date={DATE}"
        result_data = {
            'id': staff_id,
            'url': url,
            'fetch_time': datetime.now().isoformat(),
            'status': 'pending',
            'content': None,
            'error': None
        }
        
        try:
            # 步骤1: 检查当前标签页
            self.log("步骤1: 检查浏览器标签页...")
            code, stdout, stderr = self.run_openclaw(['tabs', '--profile=chrome'], timeout=10)
            self.log(f"  返回码: {code}")
            
            # 步骤2: 打开URL
            self.log(f"步骤2: 打开URL: {url[:80]}...")
            code, stdout, stderr = self.run_openclaw([
                'open', 
                '--target=host',
                '--profile=chrome',
                f'--targetUrl={url}'
            ], timeout=15)
            
            if code != 0:
                raise Exception(f"打开URL失败: {stderr}")
            
            # 提取targetId
            import re
            target_match = re.search(r'"targetId":\s*"([^"]+)"', stdout)
            if target_match:
                target_id = target_match.group(1)
                self.log(f"  标签页ID: {target_id[:20]}...")
            else:
                raise Exception("无法获取标签页ID")
            
            # 步骤3: 等待页面加载
            self.log("步骤3: 等待页面加载 (5秒)...")
            import time
            time.sleep(5)
            
            # 步骤4: 获取页面快照
            self.log("步骤4: 获取页面快照...")
            code, stdout, stderr = self.run_openclaw([
                'snapshot',
                '--target=host',
                '--profile=chrome',
                f'--targetId={target_id}',
                '--maxChars=20000'
            ], timeout=20)
            
            if code != 0:
                # 如果失败，尝试不指定targetId
                self.log("  尝试不指定targetId...")
                code, stdout, stderr = self.run_openclaw([
                    'snapshot',
                    '--target=host',
                    '--profile=chrome',
                    '--maxChars=20000'
                ], timeout=20)
            
            if code == 0 and stdout:
                result_data['status'] = 'success'
                result_data['content'] = stdout[:10000]  # 限制长度
                self.log(f"  ✓ 抓取成功，内容长度: {len(stdout)}")
            else:
                raise Exception(f"获取快照失败: {stderr}")
            
            # 步骤5: 关闭标签页
            self.log("步骤5: 关闭标签页...")
            self.run_openclaw([
                'close',
                '--target=host',
                '--profile=chrome',
                f'--targetId={target_id}'
            ], timeout=10)
            
        except Exception as e:
            result_data['status'] = 'error'
            result_data['error'] = str(e)
            self.log(f"  ✗ 抓取失败: {e}")
        
        # 保存单个结果
        filename = f'{OUTPUT_DIR}/staff_{staff_id}.json'
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(result_data, f, ensure_ascii=False, indent=2)
        self.log(f"  已保存: {filename}")
        
        return result_data
    
    def run(self):
        """运行抓取"""
        self.log(f"\n{'='*60}")
        self.log(f"开始抓取周报")
        self.log(f"日期: {DATE}")
        self.log(f"员工数: {len(STAFF_IDS)}")
        self.log(f"{'='*60}\n")
        
        # 先抓取下属
        for i, staff_id in enumerate(STAFF_IDS, 1):
            result = self.fetch_single_staff(staff_id, i, len(STAFF_IDS))
            self.results.append(result)
            
            # 每个员工处理完后暂停，让用户确认
            if i < len(STAFF_IDS):
                self.log(f"\n  已处理 {i}/{len(STAFF_IDS)} 个员工")
                self.log(f"  按Ctrl+C停止，或直接继续下一个...")
                import time
                time.sleep(2)
        
        # 最后抓取本人
        self.log(f"\n{'='*60}")
        self.log("抓取本人周报...")
        my_result = self.fetch_single_staff(MY_ID, len(STAFF_IDS)+1, len(STAFF_IDS)+1)
        self.results.append(my_result)
        
        # 保存汇总结果
        summary = {
            'date': DATE,
            'fetch_time': datetime.now().isoformat(),
            'total_staff': len(STAFF_IDS),
            'success_count': sum(1 for r in self.results if r['status'] == 'success'),
            'error_count': sum(1 for r in self.results if r['status'] == 'error'),
            'results': self.results
        }
        
        with open(f'{OUTPUT_DIR}/all_results.json', 'w', encoding='utf-8') as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)
        
        self.log(f"\n{'='*60}")
        self.log("抓取完成!")
        self.log(f"成功: {summary['success_count']}, 失败: {summary['error_count']}")
        self.log(f"汇总文件: all_results.json")
        self.log(f"{'='*60}")

if __name__ == '__main__':
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    fetcher = WeeklyFetcher()
    fetcher.run()
