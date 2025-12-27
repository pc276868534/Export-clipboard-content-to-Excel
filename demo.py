#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
微信统计信息监控器演示脚本
展示如何使用工具处理示例数据
"""

import sys
import os

# 添加当前目录到Python路径
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

def demo_parsing():
    """演示消息解析功能"""
    print("📋 微信统计信息监控器 - 功能演示")
    print("=" * 50)
    
    # 示例数据
    demo_messages = [
        "统计，张三，深圳福田公寓，1500元/月，24个月，押二付一",
        "统计，李四，深圳南山写字楼，2000元/月，12个月，押一付三", 
        "统计，王五，深圳罗湖商铺，3000元/月，36个月，押三付一",
        "这是一条普通消息，不是统计信息",
        "统计，赵六，深圳宝安厂房，5000元/月，48个月，押四付一",
        "统计格式不完整的消息，只有名字和位置",
        "统计，钱七，深圳龙岗住宅，1800元/月，18个月，押二付二"
    ]
    
    # 导入解析功能
    try:
        from wetchat_GUI import WeChatClipboardMonitor
        
        # 创建临时实例用于演示
        monitor = WeChatClipboardMonitor.__new__(WeChatClipboardMonitor)
        
        print("\n🧪 测试消息识别和解析:")
        print("-" * 50)
        
        valid_count = 0
        for i, message in enumerate(demo_messages, 1):
            print(f"\n{i}. 测试消息: {message}")
            
            # 检查是否为统计消息
            is_stats = monitor.is_statistics_message(message) if hasattr(monitor, 'is_statistics_message') else False
            
            if is_stats:
                try:
                    # 解析消息
                    parsed = monitor.parse_statistics_message(message) if hasattr(monitor, 'parse_statistics_message') else None
                    
                    if parsed:
                        print(f"   ✅ 识别为有效统计消息")
                        print(f"   📊 解析结果:")
                        print(f"      名字: {parsed['名字']}")
                        print(f"      位置: {parsed['位置']}")
                        print(f"      租金: {parsed['租金']}")
                        print(f"      租期: {parsed['租期']}")
                        print(f"      押金方式: {parsed['押金方式']}")
                        valid_count += 1
                    else:
                        print("   ⚠️  格式不符合要求")
                        
                except Exception as e:
                    print(f"   ❌ 解析失败: {e}")
            else:
                print("   ❌ 不是统计消息格式")
        
        print(f"\n📊 演示结果:")
        print(f"   总消息数: {len(demo_messages)}")
        print(f"   有效统计消息: {valid_count}")
        print(f"   无效消息: {len(demo_messages) - valid_count}")
        
        # 演示Excel导出
        if valid_count > 0:
            print(f"\n💾 可以导出的数据示例:")
            
            # 模拟数据导出
            valid_data = []
            for message in demo_messages:
                if monitor.is_statistics_message(message):
                    try:
                        parsed = monitor.parse_statistics_message(message)
                        if parsed:
                            parsed['序号'] = len(valid_data) + 1
                            valid_data.append(parsed)
                    except:
                        continue
            
            # 使用pandas创建DataFrame
            import pandas as pd
            if valid_data:
                df = pd.DataFrame(valid_data)
                export_columns = ['序号', '名字', '位置', '租金', '租期', '押金方式']
                print(df[export_columns].to_string(index=False))
                
                # 保存演示Excel
                demo_file = "demo_data.xlsx"
                df[export_columns].to_excel(demo_file, index=False, engine='openpyxl')
                print(f"\n📄 演示数据已保存到: {demo_file}")
        
        return True
        
    except Exception as e:
        print(f"❌ 演示失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def usage_guide():
    """使用指南"""
    print(f"\n{'='*50}")
    print("📖 使用指南")
    print(f"{'='*50}")
    
    print("\n🚀 启动方法:")
    print("   方法1: python run_monitor.py")
    print("   方法2: python wetchat_GUI.py")
    
    print("\n📋 基本流程:")
    print("   1. 启动程序后点击'开始监控'")
    print("   2. 打开微信，复制聊天内容（Ctrl+C）")
    print("   3. 程序自动识别统计格式消息")
    print("   4. 在'解析结果'标签页查看数据")
    print("   5. 点击'导出到Excel'保存数据")
    
    print("\n🎯 支持的消息格式:")
    print("   必须以'统计'开头")
    print("   包含5个字段：名字、位置、租金、租期、押金方式")
    print("   分隔符可以是逗号、顿号、冒号等")
    
    print("\n💡 使用技巧:")
    print("   • 复制多条消息会逐条处理")
    print("   • 非统计消息会被忽略")
    print("   • 可以随时查看历史记录")
    print("   • 支持删除无效记录")
    print("   • Excel文件可以自定义路径")

def main():
    """主演示函数"""
    print("🎬 开始演示...")
    
    # 演示解析功能
    if demo_parsing():
        # 显示使用指南
        usage_guide()
        
        print(f"\n🎉 演示完成！")
        print("现在你可以运行 'python run_monitor.py' 启动真实的监控程序了。")
    else:
        print("\n❌ 演示过程中出现问题，请检查程序和依赖库。")

if __name__ == "__main__":
    main()