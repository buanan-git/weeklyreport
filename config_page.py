#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""只启动配置页面 - 直接从 main.py 读取最新内容"""

import http.server
import socketserver
import socket
import webbrowser
import os
import re
import json

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_DIR = os.path.join(BASE_DIR, "config")
MAIN_FILE = os.path.join(BASE_DIR, "main.py")


def get_config_html_from_main():
    """从 main.py 动态读取 CONFIG_HTML"""
    with open(MAIN_FILE, 'r', encoding='utf-8') as f:
        content = f.read()

    # 匹配 CONFIG_HTML = """...""" 的内容
    match = re.search(r'CONFIG_HTML\s*=\s*"""(.+?)"""', content, re.DOTALL)
    if match:
        return match.group(1)
    raise Exception("无法从 main.py 找到 CONFIG_HTML")


class ConfigHandler(http.server.BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass

    def do_GET(self):
        if self.path == '/':
            self.send_response(200)
            self.send_header('Content-type', 'text/html; charset=utf-8')
            self.end_headers()
            self.wfile.write(CONFIG_HTML.encode('utf-8'))

        elif self.path == '/api/config':
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            cfg_file = os.path.join(CONFIG_DIR, 'config.json')
            if os.path.exists(cfg_file):
                with open(cfg_file, 'r', encoding='utf-8') as f:
                    self.wfile.write(f.read().encode('utf-8'))
            else:
                self.wfile.write(b'{}')

        elif self.path == '/api/start':
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            self.wfile.write(b'{"success":true}')

    def do_POST(self):
        if self.path == '/api/save':
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            try:
                config = json.loads(body)
                cfg_file = os.path.join(CONFIG_DIR, 'config.json')
                os.makedirs(os.path.dirname(cfg_file), exist_ok=True)
                with open(cfg_file, 'w', encoding='utf-8') as f:
                    json.dump(config, f, indent=2, ensure_ascii=False)
                self.send_response(200)
                self.send_header('Content-type', 'application/json')
                self.end_headers()
                self.wfile.write(b'{"success":true}')
            except Exception as e:
                self.send_response(500)
                self.end_headers()


def find_free_port(start_port=8080, max_attempts=10):
    for port in range(start_port, start_port + max_attempts):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind(('', port))
                return port
        except OSError:
            continue
    return start_port + max_attempts


if __name__ == "__main__":
    # 每次运行时都从 main.py 读取最新内容
    CONFIG_HTML = get_config_html_from_main()

    port = find_free_port()
    print(f"启动配置页面: http://localhost:{port}")
    print(f"从 main.py 读取最新配置页面内容")
    webbrowser.open(f"http://localhost:{port}")

    with socketserver.TCPServer(("", port), ConfigHandler) as httpd:
        print("按 Ctrl+C 停止服务器")
        httpd.serve_forever()
