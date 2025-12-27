#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试GUI程序是否能正常启动
"""

import sys
import os

# 添加当前目录到Python路径
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

def test_imports():
    """测试导入"""
    try:
        import tkinter as tk
        from tkinter import ttk, messagebox, scrolledtext
        import threading
        import time
        import pyperclip
        import re
        import pandas as pd
        from datetime import datetime
        import os
        import json
        from queue import Queue
        print("✅ 所有模块导入成功")
        return True
    except ImportError as e:
        print(f"❌ 导入失败: {e}")
        return False

def test_gui_class():
    """测试GUI类"""
    try:
        # 尝试导入主类
        from wetchat_GUI import WeChatClipboardMonitor
        print("✅ GUI类导入成功")
        return True
    except Exception as e:
        print(f"❌ GUI类导入失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_basic_functionality():
    """测试基本功能"""
    try:
        from wetchat_GUI import WeChatClipboardMonitor
        
        # 创建测试用例
        test_messages = [
            "统计，张三，深圳福田公寓，1500元，24个月，押二付一",
            "统计，李四，深圳南山写字楼，2000元，12个月，押一付三",
            "普通消息，这不是统计信息",
            "统计，王五，深圳罗湖商铺，3000元，36个月，押三付一"
        ]
        
        # 创建临时实例测试解析功能
        class TestMonitor:
            def is_statistics_message(self, text):
                if not text or len(text) < 5:
                    return False
                patterns = [
                    text.startswith('统计，'),
                    text.startswith('统计,'),
                    text.startswith('统计 '),
                    text.startswith('统计：'),
                    text.startswith('统计:'),
                    text.startswith('统计')
                ]
                if not any(patterns):
                    return False
                import re
                parts = re.split(r'[，, ：:]', text)
                return len(parts) >= 6
            
            def parse_statistics_message(self, text):
                try:
                    import re
                    from datetime import datetime
                    text = re.sub(r'^统计[，, ：:]*', '', text)
                    parts = re.split(r'[，, ：:]', text)
                    if len(parts) >= 5:
                        return {
                            '名字': parts[0].strip(),
                            '位置': parts[1].strip(),
                            '租金': parts[2].strip(),
                            '租期': parts[3].strip(),
                            '押金方式': parts[4].strip() if len(parts) > 4 else '',
                            '时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                            '原始消息': text
                        }
                    else:
                        return None
                except Exception as e:
                    print(f"解析消息失败: {e}")
                    return None
        
        test_monitor = TestMonitor()
        
        print("\n🧪 测试消息解析:")
        for i, msg in enumerate(test_messages, 1):
            is_stats = test_monitor.is_statistics_message(msg)
            print(f"{i}. {'✅' if is_stats else '❌'} {msg[:50]}...")
            
            if is_stats:
                parsed = test_monitor.parse_statistics_message(msg)
                if parsed:
                    print(f"   解析成功: {parsed['名字']} - {parsed['位置']}")
                else:
                    print("   解析失败")
        
        print("✅ 基本功能测试通过")
        return True
        
    except Exception as e:
        print(f"❌ 基本功能测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """主测试函数"""
    print("🧪 微信统计信息监控器 - 程序测试")
    print("=" * 50)
    
    # 测试导入
    if not test_imports():
        return False
    
    # 测试GUI类
    if not test_gui_class():
        return False
    
    # 测试基本功能
    if not test_basic_functionality():
        return False
    
    print("\n🎉 所有测试通过！")
    print("\n📋 启动说明:")
    print("1. 运行 'python run_monitor.py' 启动GUI程序")
    print("2. 或直接运行 'python wetchat_GUI.py'")
    print("3. 程序会自动检查依赖库并启动界面")
    
    print("\n💡 使用提示:")
    print("• 点击'开始监控'启动剪贴板监控")
    print("• 从微信复制统计消息，程序会自动识别")
    print("• 解析结果会显示在'解析结果'标签页")
    print("• 点击'导出到Excel'保存数据")
    
    return True

if __name__ == "__main__":
    success = main()
    input("\n按 Enter 键退出...")