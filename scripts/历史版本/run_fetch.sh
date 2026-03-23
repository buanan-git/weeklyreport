#!/bin/bash
# 周报抓取启动脚本
# 提供多种抓取方式

echo "=============================================="
echo "    百度OKR周报抓取工具"
echo "=============================================="
echo ""

# 检查Python环境
if ! command -v python3 &> /dev/null; then
    echo "错误: 未找到 python3"
    exit 1
fi

echo "选择抓取方式:"
echo "1. 使用 Playwright (推荐，需要安装: pip install playwright)"
echo "2. 使用 OpenClaw Browser Relay (当前连接不稳定)"
echo "3. 手动导出 - 生成操作指南"
echo ""
read -p "请输入选项 (1-3): " choice

case $choice in
    1)
        echo ""
        echo "使用 Playwright 抓取..."
        echo "确保已安装: pip install playwright"
        echo "并且已运行: playwright install chromium"
        echo ""
        python3 fetch_weekly_playwright.py
        ;;
    2)
        echo ""
        echo "使用 OpenClaw Browser Relay..."
        echo "确保Chrome扩展已连接 (显示 ON)"
        echo ""
        python3 fetch_weekly_openclaw.py
        ;;
    3)
        echo ""
        echo "生成手动操作指南..."
        cat << 'EOF'

手动导出周报步骤:
==================

1. 打开Chrome浏览器，确保已登录OKR系统

2. 依次访问以下URL（上周一日期：s2026-02-02）:

下属1:  https://okr.baidu-int.com/pages/weekly.html#/index?mode=steam&id=s801573&date=s2026-02-02
下属2:  https://okr.baidu-int.com/pages/weekly.html#/index?mode=steam&id=s733504&date=s2026-02-02
下属3:  https://okr.baidu-int.com/pages/weekly.html#/index?mode=steam&id=s682083&date=s2026-02-02
下属4:  https://okr.baidu-int.com/pages/weekly.html#/index?mode=steam&id=s929270&date=s2026-02-02
下属5:  https://okr.baidu-int.com/pages/weekly.html#/index?mode=steam&id=s798268&date=s2026-02-02
下属6:  https://okr.baidu-int.com/pages/weekly.html#/index?mode=steam&id=s811385&date=s2026-02-02
下属7:  https://okr.baidu-int.com/pages/weekly.html#/index?mode=steam&id=s865585&date=s2026-02-02
下属8:  https://okr.baidu-int.com/pages/weekly.html#/index?mode=steam&id=s783729&date=s2026-02-02
下属9:  https://okr.baidu-int.com/pages/weekly.html#/index?mode=steam&id=s400192&date=s2026-02-02
下属10: https://okr.baidu-int.com/pages/weekly.html#/index?mode=steam&id=s823700&date=s2026-02-02

本人:   https://okr.baidu-int.com/pages/weekly.html#/index?mode=steam&id=s673090&date=s2026-02-02

3. 复制每个页面的以下内容：
   - 员工姓名
   - 本周工作总结（自动+手动）
   - OKR部分（O1/O2/O3 + KR）

4. 将内容保存到文本文件，交给AI整理

EOF
        ;;
    *)
        echo "无效选项"
        exit 1
        ;;
esac
