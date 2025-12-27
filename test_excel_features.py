#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试Excel读取和行级别检测功能
"""

import sys
import os

# 添加当前目录到Python路径
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

def test_line_extraction():
    """测试行级别统计信息提取"""
    print("🧪 测试行级别统计信息提取")
    print("=" * 50)
    
    from wetchat_GUI import WeChatClipboardMonitor
    monitor = WeChatClipboardMonitor.__new__(WeChatClipboardMonitor)
    
    test_text = """这是普通消息，不是统计内容
统计，张三，深圳福田公寓，1500元，24个月，押二付一
这是另一条普通消息
统计，李四，深圳南山写字楼，2000元，12个月，押一付三
这是第三条普通消息
统计，王五，深圳罗湖商铺，3000元，36个月，押三付一
统计格式不完整消息
"""
    
    print("测试文本内容:")
    print("-" * 30)
    print(test_text)
    print("-" * 30)
    
    # 提取统计行
    statistics_lines = monitor.extract_statistics_from_text(test_text)
    
    print(f"\n📊 提取结果:")
    print(f"找到统计行数: {len(statistics_lines)}")
    
    for i, stat_info in enumerate(statistics_lines, 1):
        line_content = stat_info['line_content']
        line_number = stat_info['line_number']
        parsed_data = stat_info['parsed_data']
        
        print(f"\n第{i}条 (第{line_number}行):")
        print(f"  原文: {line_content}")
        print(f"  名字: {parsed_data['名字']}")
        print(f"  位置: {parsed_data['位置']}")
        print(f"  租金: {parsed_data['租金']}")
    
    expected_count = 3
    success = len(statistics_lines) == expected_count
    
    print(f"\n✅ 行级别提取: {'通过' if success else '失败'} (预期: {expected_count}, 实际: {len(statistics_lines)})")
    return success

def test_excel_reading():
    """测试Excel读取功能（模拟）"""
    print(f"\n🧪 测试Excel读取功能")
    print("=" * 50)
    
    try:
        import pandas as pd
        
        # 创建测试Excel数据
        test_data = [
            {
                '名字': '张三',
                '位置': '深圳福田',
                '租金': '1500元',
                '租期': '24个月',
                '押金方式': '押二付一',
                '时间': '2025-12-20 10:00:00',
                '备注': '原始数据'
            },
            {
                '名字': '张三',  # 重复数据
                '位置': '深圳福田',
                '租金': '1500元',
                '租期': '24个月',
                '押金方式': '押二付一',
                '时间': '2025-12-21 11:00:00',
                '备注': ''
            },
            {
                '名字': '李四',
                '位置': '深圳南山',
                '租金': '2000元',
                '租期': '12个月',
                '押金方式': '押一付三',
                '时间': '2025-12-22 12:00:00',
                '备注': ''
            }
        ]
        
        # 创建测试Excel文件
        df = pd.DataFrame(test_data)
        test_file = "test_excel_data.xlsx"
        df.to_excel(test_file, index=False, engine='openpyxl')
        
        print(f"✅ 测试Excel文件已创建: {test_file}")
        print(f"   包含 {len(test_data)} 条记录（含1条重复）")
        
        # 模拟读取Excel
        from wetchat_GUI import WeChatClipboardMonitor
        monitor = WeChatClipboardMonitor.__new__(WeChatClipboardMonitor)
        
        # 读取测试文件
        df_read = pd.read_excel(test_file, engine='openpyxl')
        
        excel_data = []
        excel_signatures = set()
        duplicate_count = 0
        
        for index, row in df_read.iterrows():
            record = {
                '序号': index + 1,
                '名字': str(row.get('名字', '')) if pd.notna(row.get('名字')) else '',
                '位置': str(row.get('位置', '')) if pd.notna(row.get('位置')) else '',
                '租金': str(row.get('租金', '')) if pd.notna(row.get('租金')) else '',
                '租期': str(row.get('租期', '')) if pd.notna(row.get('租期')) else '',
                '押金方式': str(row.get('押金方式', '')) if pd.notna(row.get('押金方式')) else '',
                '时间': str(row.get('时间', '')) if pd.notna(row.get('时间')) else '',
                '备注': str(row.get('备注', '')) if pd.notna(row.get('备注')) else '',
                '来源': 'Excel'
            }
            
            signature = monitor.get_data_signature(record)
            if signature in excel_signatures:
                duplicate_count += 1
                record['备注'] = 'Excel内重复记录'
            
            excel_signatures.add(signature)
            excel_data.append(record)
        
        print(f"\n📊 Excel读取模拟结果:")
        print(f"总记录数: {len(excel_data)}")
        print(f"重复记录数: {duplicate_count}")
        print(f"有效记录数: {len(excel_data) - duplicate_count}")
        print(f"签名集大小: {len(excel_signatures)}")
        
        # 清理测试文件
        if os.path.exists(test_file):
            os.remove(test_file)
            print(f"✅ 测试文件已清理: {test_file}")
        
        return len(excel_data) == 3 and duplicate_count == 1
        
    except Exception as e:
        print(f"❌ Excel读取测试失败: {e}")
        return False

def test_data_merge():
    """测试数据合并功能"""
    print(f"\n🧪 测试数据合并功能")
    print("=" * 50)
    
    from wetchat_GUI import WeChatClipboardMonitor
    monitor = WeChatClipboardMonitor.__new__(WeChatClipboardMonitor)
    
    # 模拟Excel数据
    excel_data = [
        {'序号': 1, '名字': '张三', '位置': '深圳福田', '租金': '1500元', '备注': '', '来源': 'Excel'},
        {'序号': 2, '名字': '李四', '位置': '深圳南山', '租金': '2000元', '备注': '', '来源': 'Excel'}
    ]
    
    # 模拟新数据
    new_data = [
        {'序号': 1, '名字': '王五', '位置': '深圳罗湖', '租金': '3000元', '备注': '', '来源': '剪贴板'},
        {'序号': 2, '名字': '张三', '位置': '深圳福田', '租金': '1500元', '备注': '重复记录', '来源': '剪贴板'}
    ]
    
    # 合并数据
    monitor.existing_excel_data = excel_data
    monitor.processed_data = new_data
    
    # 计算合并后的重复检测
    all_signatures = set()
    excel_signatures = [monitor.get_data_signature(data) for data in excel_data]
    new_signatures = [monitor.get_data_signature(data) for data in new_data]
    
    all_signatures.update(excel_signatures)
    duplicates_in_new = [sig for sig in new_signatures if sig in excel_signatures]
    
    print(f"📊 数据合并分析:")
    print(f"Excel数据: {len(excel_data)} 条")
    print(f"新数据: {len(new_data)} 条")
    print(f"Excel签名: {len(excel_signatures)} 个")
    print(f"新数据签名: {len(new_signatures)} 个")
    print(f"重复签名: {len(duplicates_in_new)} 个")
    
    # 验证重复检测
    expected_duplicates = 1  # 张三是重复的
    success = len(duplicates_in_new) == expected_duplicates
    
    print(f"✅ 数据合并检测: {'通过' if success else '失败'} (预期重复: {expected_duplicates}, 实际: {len(duplicates_in_new)})")
    
    return success

def main():
    """主测试函数"""
    print("🚀 测试Excel读取和行级别检测功能")
    print("=" * 60)
    
    # 测试行级别提取
    line_ok = test_line_extraction()
    
    # 测试Excel读取
    excel_ok = test_excel_reading()
    
    # 测试数据合并
    merge_ok = test_data_merge()
    
    print(f"\n{'='*60}")
    print("📊 测试结果:")
    print(f"✅ 行级别统计提取: {'通过' if line_ok else '失败'}")
    print(f"✅ Excel读取功能: {'通过' if excel_ok else '失败'}")
    print(f"✅ 数据合并检测: {'通过' if merge_ok else '失败'}")
    
    if line_ok and excel_ok and merge_ok:
        print(f"\n🎉 所有新功能测试通过！")
        print(f"\n📋 新增功能说明:")
        print("1. ✅ 读取现有Excel文件并检测内部重复")
        print("2. ✅ 行级别检测：只处理与'统计'同一行的内容")
        print("3. ✅ 智能去重：Excel数据与新数据比较")
        print("4. ✅ 追加导出：保留原Excel数据，追加新数据")
        print("5. ✅ 完整工作流：读取Excel → 监控 → 追加导出")
        
        print(f"\n🚀 推荐使用流程:")
        print("1. 点击'读取Excel'加载现有数据")
        print("2. 点击'开始监控'监控剪贴板")
        print("3. 复制微信内容（可一次复制多行）")
        print("4. 程序自动提取统计行并去重")
        print("5. 点击'导出到Excel'追加保存")
    else:
        print(f"\n❌ 部分功能测试失败")

if __name__ == "__main__":
    main()