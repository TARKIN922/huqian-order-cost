@echo off
REM 关闭5000端口程序脚本 - Windows版

setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo ================================================================
echo.
echo                    关闭Web服务 - 5000端口清理
echo.
echo ================================================================
echo.

REM 查找5000端口上的进程
echo  正在查找5000端口上的进程...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":5000"') do (
    if not "%%a"=="0" (
        echo  找到进程ID: %%a
        taskkill /PID %%a /F >nul 2>&1
        if errorlevel 1 (
            echo  - 关闭失败: %%a
        ) else (
            echo  ✓ 已关闭进程: %%a
        )
    )
)

REM 等待一秒
timeout /t 1 /nobreak >nul

REM 验证5000端口是否已释放
echo.
echo  验证5000端口状态...
netstat -ano | findstr ":5000" >nul 2>&1
if errorlevel 1 (
    echo  ✓ 5000端口已完全释放！
    echo.
    echo  ================================================================
    echo  ✓ 服务已关闭，可以重新启动应用
    echo  ================================================================
    echo.
) else (
    echo  ⚠ 5000端口仍有进程运行，尝试强制关闭所有Python进程...
    taskkill /F /IM python.exe >nul 2>&1
    echo  ✓ 已强制关闭所有Python进程
    echo.
)

pause
