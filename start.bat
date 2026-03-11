@echo off
REM 订单处理平台启动脚本 - Windows版
REM 自动安装依赖并启动应用

setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo ================================================================
echo.
echo                  订单处理平台 - 自动启动
echo.
echo ================================================================
echo.

REM 检查Python是否已安装
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo  错误: 未检测到Python安装
    echo  请从以下链接下载Python: https://www.python.org/downloads/
    echo  安装时请勾选 "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

echo  检查Python版本...
python --version

REM 检查并安装依赖
echo.
echo  检查依赖包...
pip show flask >nul 2>&1
if errorlevel 1 (
    echo.
    echo  正在安装依赖包...
    echo  (首次运行会下载dependencies，可能需要几分钟)
    echo.
    pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo  错误: 依赖安装失败
        echo  请检查网络连接或手动运行:
        echo    pip install -r requirements.txt
        echo.
        pause
        exit /b 1
    )
)

REM 启动应用
echo.
echo  启动Web服务...
echo.
echo  ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
echo  ★                                           ★
echo  ★    服务器已启动！                         ★
echo  ★    访问地址: http://127.0.0.1:5000      ★
echo  ★                                           ★
echo  ★    如未自动打开浏览器，请手动访问上述地址 ★
echo  ★                                           ★
echo  ★    按 Ctrl+C 停止服务                    ★
echo  ★                                           ★
echo  ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
echo.

python run.py

REM 如果程序异常退出
echo.
echo  程序已退出
pause
