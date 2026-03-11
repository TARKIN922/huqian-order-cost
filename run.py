#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
订单处理平台启动脚本
"""

import os
import sys
import webbrowser
import time
from pathlib import Path

# 添加当前目录到Python路径
sys.path.insert(0, str(Path(__file__).parent))

def check_dependencies():
    """检查必要的依赖"""
    required_packages = ['flask', 'pandas', 'openpyxl', 'chardet']
    missing = []
    
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing.append(package)
    
    if missing:
        print("❌ 缺少必要的Python包:")
        print(f"   {', '.join(missing)}")
        print("\n请运行以下命令安装:")
        print(f"   pip install {' '.join(missing)}")
        return False
    
    return True


def main():
    """主函数"""
    print("=" * 50)
    print("📊 订单处理平台")
    print("=" * 50)
    
    # 检查依赖
    print("\n✓ 检查依赖...")
    if not check_dependencies():
        sys.exit(1)
    
    print("✓ 所有依赖已安装")
    
    # 启动Flask应用
    print("\n🚀 启动服务器...")
    print("   服务地址: http://127.0.0.1:5000")
    print("   按 Ctrl+C 停止服务器")
    print("\n")
    
    try:
        # 延迟打开浏览器，给服务器时间启动
        def open_browser():
            time.sleep(2)
            webbrowser.open('http://127.0.0.1:5000')
        
        import threading
        browser_thread = threading.Thread(target=open_browser, daemon=True)
        browser_thread.start()
        
        # 导入并运行Flask应用
        from app import app
        
        app.run(
            host='127.0.0.1',
            port=5000,
            debug=False,
            use_reloader=False
        )
    
    except Exception as e:
        print(f"❌ 启动失败: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
