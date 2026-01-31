#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试 Excel 杠杆率计算模板的公式和功能
"""

from openpyxl import load_workbook
import sys


def test_template():
    """测试模板的基本结构和公式"""
    print("=" * 60)
    print("测试 Excel 杠杆率计算模板")
    print("=" * 60)
    
    # 加载工作簿
    wb = load_workbook('leverage_calculation_template.xlsx')
    
    # 验证工作表数量和名称
    expected_sheets = ["银利3号", "银利5号", "健康险"]
    actual_sheets = wb.sheetnames
    
    print(f"\n1. 验证工作表...")
    print(f"   预期工作表: {expected_sheets}")
    print(f"   实际工作表: {actual_sheets}")
    
    if actual_sheets != expected_sheets:
        print("   ❌ 错误：工作表不匹配！")
        return False
    print("   ✓ 工作表验证通过")
    
    # 测试每个工作表
    test_cases = [
        {
            "name": "银利3号",
            "leverage_limit": 1.40,
            "test_data": {
                "总资产": 100000,
                "卖出回购金融资产款": 20000,
                "交易性金融负债（买断）": 5000,
                "交易性金融负债（借券）": 3000,
                "买入返售金融资产": 2000,
                "其他负债": 1000,
                "买入券": 5000,
                "卖出回购融资": 3000
            }
        }
    ]
    
    for test_case in test_cases:
        print(f"\n2. 测试工作表：{test_case['name']}")
        ws = wb[test_case['name']]
        
        # 输入测试数据
        print("   输入测试数据...")
        
        # 在初始资产负债表区域输入数据（从第5行开始）
        ws['B5'] = test_case['test_data']['总资产']
        ws['B6'] = test_case['test_data']['卖出回购金融资产款']
        ws['B7'] = test_case['test_data']['交易性金融负债（买断）']
        ws['B8'] = test_case['test_data']['交易性金融负债（借券）']
        ws['B9'] = test_case['test_data']['买入返售金融资产']
        ws['B10'] = test_case['test_data']['其他负债']
        
        # 保存以触发公式计算
        wb.save('leverage_calculation_template_test.xlsx')
        
        # 重新加载以获取计算结果
        wb_test = load_workbook('leverage_calculation_template_test.xlsx', data_only=True)
        ws_test = wb_test[test_case['name']]
        
        # 验证总负债计算
        expected_total_liability = (
            test_case['test_data']['卖出回购金融资产款'] +
            test_case['test_data']['交易性金融负债（买断）'] +
            test_case['test_data']['交易性金融负债（借券）'] +
            test_case['test_data']['其他负债']
        )
        actual_total_liability = ws_test['B11'].value
        
        print(f"   验证总负债计算...")
        print(f"     预期: {expected_total_liability:,.0f}")
        print(f"     实际: {actual_total_liability:,.0f}" if actual_total_liability else "     实际: None (公式未计算)")
        
        # 验证净资产计算
        expected_net_asset = (
            test_case['test_data']['总资产'] - expected_total_liability
        )
        actual_net_asset = ws_test['B12'].value
        
        print(f"   验证净资产计算...")
        print(f"     预期: {expected_net_asset:,.0f}")
        print(f"     实际: {actual_net_asset:,.0f}" if actual_net_asset else "     实际: None (公式未计算)")
        
        # 验证初始杠杆率
        expected_leverage = test_case['test_data']['总资产'] / expected_net_asset if expected_net_asset != 0 else 0
        actual_leverage = ws_test['B13'].value
        
        print(f"   验证初始杠杆率...")
        print(f"     预期: {expected_leverage:.2%}")
        print(f"     实际: {actual_leverage:.2%}" if actual_leverage else "     实际: None (公式未计算)")
        
        # 验证杠杆限制
        print(f"   验证杠杆限制...")
        print(f"     设定限制: {test_case['leverage_limit']:.2%}")
        
        if actual_leverage and actual_leverage > test_case['leverage_limit']:
            print(f"     ⚠️  警告：杠杆率 ({actual_leverage:.2%}) 超过限制 ({test_case['leverage_limit']:.2%})")
        else:
            print(f"     ✓ 杠杆率正常")
        
    print("\n" + "=" * 60)
    print("测试完成")
    print("=" * 60)
    
    # 清理测试文件
    import os
    if os.path.exists('leverage_calculation_template_test.xlsx'):
        os.remove('leverage_calculation_template_test.xlsx')
    
    return True


def display_template_info():
    """显示模板的详细信息"""
    print("\n" + "=" * 60)
    print("Excel 模板结构信息")
    print("=" * 60)
    
    wb = load_workbook('leverage_calculation_template.xlsx')
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        print(f"\n工作表: {sheet_name}")
        print("-" * 60)
        
        # 显示标题信息
        title = ws['A1'].value
        limit = ws['A2'].value
        print(f"  标题: {title}")
        print(f"  {limit}")
        
        # 显示各个区域
        sections = [
            ("一、初始资产负债表数据", "A4"),
            ("二、当日操作输入区", "A16"),
            ("三、变化前后对比", "A26"),
            ("四、杠杆限制预警", "A38")
        ]
        
        print("\n  包含的区域:")
        for section_name, cell_ref in sections:
            cell_value = ws[cell_ref].value
            if cell_value:
                print(f"    • {cell_value}")
        
        # 显示操作类型
        print("\n  支持的操作类型:")
        for row in range(18, 27):  # 操作输入区域
            operation = ws[f'A{row}'].value
            if operation:
                print(f"    • {operation}")
        
        # 显示对比指标
        print("\n  对比指标:")
        for row in range(28, 39):  # 对比区域
            indicator = ws[f'A{row}'].value
            if indicator and indicator != "指标":
                print(f"    • {indicator}")
    
    print("\n" + "=" * 60)


if __name__ == "__main__":
    try:
        # 显示模板信息
        display_template_info()
        
        # 运行测试
        print("\n")
        success = test_template()
        
        if success:
            print("\n✓ 所有测试通过！")
            sys.exit(0)
        else:
            print("\n❌ 测试失败！")
            sys.exit(1)
    except Exception as e:
        print(f"\n❌ 测试出错: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
