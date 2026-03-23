@echo off
chcp 65001 >nul
echo ========================================
echo 周报自动化工具 - 调试模式
echo 错误日志将保存到 error_log.txt
echo ========================================
echo.

rem 查找可执行文件
setlocal enabledelayedexpansion
set "exe_file="

for %%i in (*.exe) do (
    set "exe_file=%%i"
    goto :found
)

:found
if "%exe_file%"=="" (
    echo 错误: 未找到可执行文件
    echo 请先运行打包脚本生成可执行文件
    pause
    exit /b 1
)

echo 找到可执行文件: %exe_file%
echo 正在启动调试模式...
echo.

rem 运行程序并捕获所有输出
"%exe_file%" > output.log 2>&1

echo.
echo ========================================
echo 程序已退出，退出码: %ERRORLEVEL%
echo 请查看 output.log 查看错误信息
echo ========================================
pause
