#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
微信统计信息监控器启动脚本
"""

import sys
import os

# 添加当前目录到Python路径
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

def check_dependencies():
    """检查依赖库"""
    required_packages = ['pyperclip', 'pandas', 'openpyxl']
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)
    
    if missing_packages:
        print("❌ 缺少以下依赖库:")
        for package in missing_packages:
            print(f"   - {package}")
        print("\n请运行以下命令安装依赖:")
        print(f"pip install {' '.join(missing_packages)}")
        print(f"\n或者使用 requirements.txt:")
        print(f"pip install -r requirements.txt")
        return False
    
    print("✅ 所有依赖库检查通过")
    return True

def main():
    """主函数"""
    print("🚀 微信统计信息监控器启动中...")
    print("=" * 50)
    
    # 检查依赖
    if not check_dependencies():
        input("\n按 Enter 键退出...")
        return
    
    print("\n📋 功能说明:")
    print("• 自动监控剪贴板内容")
    print("• 识别以'统计'开头的微信消息")
    print("• 自动解析数据并导出到Excel")
    print("• 提供实时状态显示和数据管理")
    
    print("\n🎯 支持的消息格式:")
    print("   统计，名字，位置，租金，租期，押金方式")
    print("   例如: 统计，张三，深圳福田公寓，1500元，24个月，押二付一")
    
    print("\n⏰ 正在启动GUI界面...")
    
    try:
        # 导入主程序
        from wetchat_GUI import main as gui_main
        gui_main()
    except Exception as e:
        print(f"\n❌ 启动失败: {e}")
        import traceback
        traceback.print_exc()
        input("\n按 Enter 键退出...")

if __name__ == "__main__":
    main()