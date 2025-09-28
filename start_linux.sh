#!/bin/bash

echo "启动Excel/剪贴板群发邮件自动化工具..."
echo

# 检查Python是否已安装
if ! command -v python3 &> /dev/null; then
    echo "错误: 未检测到Python3，请先安装Python 3.8或更高版本"
    exit 1
fi

# 检查虚拟环境是否存在
if [ ! -d "email_env" ]; then
    echo "创建Python虚拟环境..."
    python3 -m venv email_env
fi

# 激活虚拟环境
echo "激活虚拟环境..."
source email_env/bin/activate

# 检查并安装依赖
echo "检查并安装依赖包..."
pip install -r requirements_linux.txt

# 启动应用程序
echo "启动应用程序..."
python3 email_automation.py

# 暂停以便查看错误信息
echo "按任意键继续..."
read -n 1 -s