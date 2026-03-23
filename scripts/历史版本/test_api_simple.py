#!/usr/bin/env python3
"""
简单API测试脚本
"""

import requests
import json
import time

def test_api():
    url = "http://10.213.206.148:8080/apis/ais-v2/chat/completions"
    headers = {
        "Content-Type": "application/json",
        "Authorization": "Bearer sk-fab01582-97ed-4132-7d1b-65e8e4b7bf82"
    }
    
    # 简单的测试消息
    payload = {
        "model": "glm-45-v1",
        "messages": [
            {"role": "user", "content": "你好，请回复'测试成功'"}
        ],
        "max_tokens": 50,
        "temperature": 0.7
    }
    
    print("发送API请求...")
    start_time = time.time()
    
    try:
        response = requests.post(url, headers=headers, json=payload, timeout=30)
        end_time = time.time()
        
        print(f"响应时间: {end_time - start_time:.2f}秒")
        print(f"状态码: {response.status_code}")
        
        if response.status_code == 200:
            result = response.json()
            print("API调用成功!")
            print(f"响应内容: {json.dumps(result, indent=2, ensure_ascii=False)}")
        else:
            print(f"API调用失败: {response.text}")
            
    except requests.exceptions.Timeout:
        print("请求超时")
    except Exception as e:
        print(f"请求异常: {e}")

if __name__ == "__main__":
    test_api()