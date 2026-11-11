@echo off
chcp 65001 >nul
title 华为基站MML任务结果智能解析与转换工具
echo 正在启动华为MML解析转换工具图形界面...
start "" pythonw "%~dp0mml_gui.py"
if %errorlevel% neq 0 (
    echo [提示] 尝试使用 python 启动...
    python "%~dp0mml_gui.py"
    pause
)
