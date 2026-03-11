#!/bin/bash

# 快速重启脚本 - Mac/Linux版

echo ""
echo "================================================================"
echo ""
echo "              Web服务 - 快速重启 (关闭 + 启动)"
echo ""
echo "================================================================"
echo ""

# 第一步：关闭所有进程
echo "  [步骤 1/3] 关闭现有服务..."
pkill -9 -f "python.*run.py" 2>/dev/null
if [ $? -eq 0 ]; then
    echo "  ✓ 已关闭Python进程"
else
    echo "  ℹ 没有运行的Python进程"
fi
echo ""

# 等待清理
sleep 2

# 第二步：清空uploads文件夹（可选）
echo "  [步骤 2/3] 清理临时文件..."
if [ -d "uploads" ]; then
    rm -rf uploads
    echo "  ✓ 已清理上传文件夹"
else
    echo "  ℹ 上传文件夹不存在，跳过"
fi
echo ""

# 第三步：启动应用
echo "  [步骤 3/3] 启动服务..."
echo ""
echo "  ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★"
echo "  ★                                           ★"
echo "  ★    服务器启动中...                        ★"
echo "  ★    访问地址: http://127.0.0.1:5000      ★"
echo "  ★                                           ★"
echo "  ★    按 Ctrl+C 停止服务                    ★"
echo "  ★                                           ★"
echo "  ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★"
echo ""

python3 run.py

echo ""
echo "  程序已退出"
