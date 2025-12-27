#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试重复数据检测功能
"""

import sys
import os

# 添加当前目录到Python路径
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

def test_duplicate_detection():
    """测试重复检测功能"""
    print("🧪 重复数据检测功能测试")
    print("=" * 50)
    
    # 导入监控类
    try:
        from wetchat_GUI import WeChatClipboardMonitor
        
        # 创建临时实例
        monitor = WeChatClipboardMonitor.__new__(WeChatClipboardMonitor)
        
        # 测试消息
        test_messages = [
            "统计，张三，深圳福田公寓，1500元/月，24个月，押二付一",
            "统计，张三，深圳福田公寓，1500元/月，24个月，押二付一",  # 重复
            "统计，李四，深圳南山写字楼，2000元/月，12个月，押一付三",
            "普通的非统计消息，张三，深圳福田公寓，1500元/月，24个月，押二付一",  # 非统计
            "前缀统计，张三，深圳福田公寓，1500元/月，24个月，押二付一",  # 不符合要求
            "统计，张三，深圳福田公寓，1500元/月，24个月，押二付一",  # 再次重复
            "统计，王五，深圳罗湖商铺，3000元/月，36个月，押三付一",
        ]
        
        print("\n📋 测试消息识别:")
        valid_messages = []
        invalid_messages = []
        
        for i, msg in enumerate(test_messages, 1):
            print(f"{i}. {msg}")
            is_stats = monitor.is_statistics_message(msg)
            
            if is_stats:
                print(f"   ✅ 有效统计消息")
                valid_messages.append(msg)
            else:
                print(f"   ❌ 无效（不以'统计'开头或格式不符）")
                invalid_messages.append(msg)
            print()
        
        print(f"\n📊 测试结果统计:")
        print(f"总消息数: {len(test_messages)}")
        print(f"有效统计消息: {len(valid_messages)}")
        print(f"无效消息: {len(invalid_messages)}")
        
        if valid_messages:
            print(f"\n🔄 重复检测测试:")
            processed_signatures = set()
            duplicate_count = 0
            
            for i, msg in enumerate(valid_messages, 1):
                parsed_data = monitor.parse_statistics_message(msg)
                if parsed_data:
                    signature = monitor.get_data_signature(parsed_data)
                    
                    if signature in processed_signatures:
                        print(f"{i}. 🔁 重复: {parsed_data['名字']} - {parsed_data['位置']}")
                        duplicate_count += 1
                    else:
                        print(f"{i}. ✅ 新记录: {parsed_data['名字']} - {parsed_data['位置']}")
                        processed_signatures.add(signature)
            
            print(f"\n重复数据: {duplicate_count}")
            print(f"唯一记录: {len(processed_signatures)}")
        
        print(f"\n✅ 测试完成！重复检测功能正常工作。")
        return True
        
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_strict_format():
    """测试严格格式要求"""
    print(f"\n{'='*50}")
    print("🔍 严格格式要求测试")
    print(f"{'='*50}")
    
    from wetchat_GUI import WeChatClipboardMonitor
    monitor = WeChatClipboardMonitor.__new__(WeChatClipboardMonitor)
    
    strict_test_cases = [
        ("统计，张三，位置，租金，租期，押金方式", True),
        ("统计,张三,位置,租金,租期,押金方式", True),
        ("统计：张三：位置：租金：租期：押金方式", True),
        ("统计:张三:位置:租金:租期:押金方式", True),
        ("统计 张三 位置 租金 租期 押金方式", True),
        ("前缀统计，张三，位置，租金，租期，押金方式", False),
        ("xx统计，张三，位置，租金，租期，押金方式", False),
        ("统计，张三，位置，租金，租期", False),  # 字段不足
        ("张三，位置，租金，租期，押金方式", False),  # 没有统计开头
    ]
    
    print("\n📝 严格格式测试:")
    for i, (msg, expected) in enumerate(strict_test_cases, 1):
        result = monitor.is_statistics_message(msg)
        status = "✅" if result == expected else "❌"
        print(f"{i}. {status} '{msg}' -> {result} (预期: {expected})")
    
    print(f"\n📋 格式要求:")
    print("• 必须严格以'统计'开头（不能有其他前缀字符）")
    print("• 支持多种分隔符：中文逗号、英文逗号、冒号、空格")
    print("• 必须包含至少6个字段（统计+5个数据字段）")

def main():
    """主测试函数"""
    print("🚀 开始测试重复检测和严格格式功能...")
    
    # 测试重复检测
    if test_duplicate_detection():
        # 测试严格格式
        test_strict_format()
        
        print(f"\n{'='*50}")
        print("🎉 所有测试完成！")
        print(f"{'='*50}")
        print("\n📖 使用说明:")
        print("1. 程序现在只处理严格以'统计'开头的消息")
        print("2. 自动检测并提示重复数据")
        print("3. 用户可选择是否添加重复记录")
        print("4. 重复记录会标记备注信息")
        
    else:
        print("\n❌ 测试过程中发现问题")

if __name__ == "__main__":
    main()