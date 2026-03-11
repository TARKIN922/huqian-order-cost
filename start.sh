#!/bin/bash

# 订单处理平台启动脚本 - Mac/Linux版

cd "$(dirname "$0")"

echo ""
echo "================================================================"
echo ""
echo "                  订单处理平台 - 自动启动"
echo ""
echo "================================================================"
echo ""

# 检查Python
if ! command -v python3 &> /dev/null; then
    echo "  错误: 未检测到Python3安装"
    echo "  请从以下链接下载: https://www.python.org/downloads/"
    echo ""
    exit 1
fi

echo "  检查Python版本..."
python3 --version

# 检查并安装依赖
echo ""
echo "  检查依赖包..."

pip3 show flask &> /dev/null
if [ $? -ne 0 ]; then
    echo ""
    echo "  正在安装依赖包..."
    echo "  (首次运行会下载dependencies，可能需要几分钟)"
    echo ""
    pip3 install -r requirements.txt
    if [ $? -ne 0 ]; then
        echo ""
        echo "  错误: 依赖安装失败"
        echo "  请检查网络连接或手动运行:"
        echo "    pip3 install -r requirements.txt"
        echo ""
        exit 1
    fi
fi

# 启动应用
echo ""
echo "  启动Web服务..."
echo ""
echo "  ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★"
echo "  ★                                           ★"
echo "  ★    服务器已启动！                         ★"
echo "  ★    访问地址: http://127.0.0.1:5000      ★"
echo "  ★                                           ★"
echo "  ★    如未自动打开浏览器，请手动访问上述地址 ★"
echo "  ★                                           ★"
echo "  ★    按 Ctrl+C 停止服务                    ★"
echo "  ★                                           ★"
echo "  ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★"
echo ""

python3 run.py

echo ""
echo "  程序已退出"
echo ""
