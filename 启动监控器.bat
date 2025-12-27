@echo off
chcp 65001 >nul
title 微信统计信息监控器

echo 🚀 微信统计信息监控器启动中...
echo ================================

:: 检查Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ 未找到Python，请先安装Python
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

:: 检查依赖库
echo 📋 检查依赖库...
python -c "import pyperclip, pandas, openpyxl" >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ 缺少依赖库，正在安装...
    pip install pyperclip pandas openpyxl
    if %errorlevel% neq 0 (
        echo ❌ 依赖库安装失败，请手动运行:
        echo pip install pyperclip pandas openpyxl
        pause
        exit /b 1
    )
)

echo ✅ 环境检查通过

:: 启动程序
echo 🎯 启动GUI程序...
python run_monitor.py

if %errorlevel% neq 0 (
    echo ❌ 程序启动失败
    pause
)

echo.
echo 程序已退出，按任意键关闭窗口...
pause >nul