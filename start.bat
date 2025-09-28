@echo off
echo 启动Excel/剪贴板群发邮件自动化工具...
echo.

REM 检查Python是否已安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    python3 --version >nul 2>&1
    if %errorlevel% neq 0 (
        echo 错误: 未检测到Python，请先安装Python 3.8或更高版本
        pause
        exit /b 1
    ) else (
        set PYTHON_CMD=python3
    )
) else (
    set PYTHON_CMD=python
)

REM 检查并安装依赖
echo 检查并安装依赖包...
pip install -r requirements.txt

REM 启动应用程序
echo 启动应用程序...
%PYTHON_CMD% email_automation.py

pause