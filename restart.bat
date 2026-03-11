@echo off
REM 快速重启脚本 - Windows版

setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo ================================================================
echo.
echo              Web服务 - 快速重启 (关闭 + 启动)
echo.
echo ================================================================
echo.

REM 第一步：关闭所有进程
echo  [步骤 1/3] 关闭现有服务...
taskkill /F /IM python.exe >nul 2>&1

REM 等待清理
timeout /t 2 /nobreak >nul

echo  ✓ 已关闭所有Python进程
echo.

REM 第二步：清空uploads文件夹（可选）
echo  [步骤 2/3] 清理临时文件...
if exist "uploads" (
    rmdir /s /q uploads >nul 2>&1
    echo  ✓ 已清理上传文件夹
) else (
    echo  ℹ 上传文件夹不存在，跳过
)
echo.

REM 第三步：启动应用
echo  [步骤 3/3] 启动服务...
echo.
echo  ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
echo  ★                                           ★
echo  ★    服务器启动中...                        ★
echo  ★    访问地址: http://127.0.0.1:5000      ★
echo  ★                                           ★
echo  ★    按 Ctrl+C 停止服务                    ★
echo  ★                                           ★
echo  ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
echo.

python run.py

echo.
echo  程序已退出
pause
