#!/bin/bash

# 关闭5000端口程序脚本 - Mac/Linux版

echo ""
echo "================================================================"
echo ""
echo "                    关闭Web服务 - 5000端口清理"
echo ""
echo "================================================================"
echo ""

# 查找5000端口上的进程
echo "  正在查找5000端口上的进程..."
pids=$(lsof -ti :5000 2>/dev/null)

if [ -z "$pids" ]; then
    echo "  ℹ 5000端口当前无运行进程"
else
    echo "  找到以下进程ID: $pids"
    for pid in $pids; do
        echo "  - 关闭进程 $pid..."
        kill -9 $pid 2>/dev/null
        if [ $? -eq 0 ]; then
            echo "    ✓ 已关闭进程: $pid"
        else
            echo "    ✗ 关闭失败: $pid"
        fi
    done
fi

# 等待一秒
sleep 1

# 验证5000端口是否已释放
echo ""
echo "  验证5000端口状态..."
if lsof -ti :5000 >/dev/null 2>&1; then
    echo "  ⚠ 5000端口仍有进程运行"
    echo "  尝试强制关闭所有Python进程..."
    pkill -9 -f "python.*run.py" 2>/dev/null
    echo "  ✓ 已强制关闭Python进程"
else
    echo "  ✓ 5000端口已完全释放！"
fi

echo ""
echo "================================================================"
echo "✓ 服务已关闭，可以重新启动应用"
echo "================================================================"
echo ""
