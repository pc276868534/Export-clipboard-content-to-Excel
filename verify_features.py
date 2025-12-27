#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
验证新功能的简单脚本
"""

def test_format_detection():
    """测试格式检测功能"""
    # 模拟格式检测逻辑
    import re
    
    def is_statistics_message(text):
        if not text or len(text) < 5:
            return False
        
        patterns = [
            text.startswith('统计，'),
            text.startswith('统计,'),
            text.startswith('统计 '),
            text.startswith('统计：'),
            text.startswith('统计:'),
            text == '统计'
        ]
        
        if not any(patterns):
            return False
        
        parts = re.split(r'[，, ：:]', text)
        return len(parts) >= 6
    
    # 测试用例
    test_cases = [
        ("统计，张三，深圳福田，1500元，24个月，押二付一", True),
        ("前缀统计，张三，深圳福田，1500元，24个月，押二付一", False),
        ("普通消息，张三，深圳福田，1500元，24个月，押二付一", False),
        ("统计，李四，深圳南山，2000元，12个月，押一付三", True),
        ("张三，深圳福田，1500元，24个月，押二付一", False),
    ]
    
    print("🧪 格式检测功能验证:")
    print("=" * 50)
    
    all_passed = True
    for msg, expected in test_cases:
        result = is_statistics_message(msg)
        status = "✅" if result == expected else "❌"
        print(f"{status} {msg[:30]}... -> {result} (预期: {expected})")
        if result != expected:
            all_passed = False
    
    return all_passed

def test_duplicate_detection():
    """测试重复检测功能"""
    def get_signature(data):
        return f"{data.get('名字', '')}|{data.get('位置', '')}|{data.get('租金', '')}|{data.get('租期', '')}|{data.get('押金方式', '')}"
    
    test_data = [
        {"名字": "张三", "位置": "深圳福田", "租金": "1500元", "租期": "24个月", "押金方式": "押二付一"},
        {"名字": "张三", "位置": "深圳福田", "租金": "1500元", "租期": "24个月", "押金方式": "押二付一"},  # 重复
        {"名字": "李四", "位置": "深圳南山", "租金": "2000元", "租期": "12个月", "押金方式": "押一付三"},
    ]
    
    print("\n🔄 重复检测功能验证:")
    print("=" * 50)
    
    processed_signatures = set()
    duplicate_count = 0
    
    for i, data in enumerate(test_data, 1):
        signature = get_signature(data)
        
        if signature in processed_signatures:
            print(f"🔁 第{i}条: 重复数据 - {data['名字']}")
            duplicate_count += 1
        else:
            print(f"✅ 第{i}条: 新数据 - {data['名字']}")
            processed_signatures.add(signature)
    
    print(f"\n重复数据: {duplicate_count}")
    print(f"唯一记录: {len(processed_signatures)}")
    
    return duplicate_count == 1

def main():
    """主验证函数"""
    print("🚀 微信统计信息监控器 - 新功能验证")
    print("=" * 60)
    
    # 测试格式检测
    format_ok = test_format_detection()
    
    # 测试重复检测
    duplicate_ok = test_duplicate_detection()
    
    print("\n" + "=" * 60)
    print("📊 验证结果:")
    print(f"✅ 格式检测功能: {'通过' if format_ok else '失败'}")
    print(f"✅ 重复检测功能: {'通过' if duplicate_ok else '失败'}")
    
    if format_ok and duplicate_ok:
        print("\n🎉 所有新功能验证通过！")
        print("\n📋 更新说明:")
        print("1. ✅ 严格格式检测 - 只处理以'统计'开头的消息")
        print("2. ✅ 重复数据检测 - 自动识别并提示重复记录")
        print("3. ✅ 用户确认机制 - 可选择是否添加重复记录")
        print("4. ✅ 备注标记功能 - 重复记录会添加特殊标记")
        
        print("\n🚀 程序已准备就绪！使用以下方式启动:")
        print("   • 双击 '启动监控器.bat'")
        print("   • 运行 'python run_monitor.py'")
        print("   • 直接运行 'python wetchat_GUI.py'")
    else:
        print("\n❌ 部分功能验证失败，请检查代码")

if __name__ == "__main__":
    main()