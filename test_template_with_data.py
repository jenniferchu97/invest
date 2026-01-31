#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试杠杆率计算模板 - 填充示例数据
Test leverage calculation template with sample data
"""

from openpyxl import load_workbook
from openpyxl.styles.numbers import FORMAT_NUMBER_00
import datetime

def test_with_sample_data():
    """使用示例数据测试模板"""
    
    print("=" * 80)
    print("测试杠杆率计算模板 - 填充示例数据")
    print("=" * 80)
    
    # 加载模板文件
    wb = load_workbook('/home/runner/work/invest/invest/leverage_calculation_template.xlsx')
    
    # 测试数据
    test_scenarios = {
        '银利3号': {
            '数据日期': '2026-01-31',
            '总资产': 1000000,
            '总负债': 600000,
            '卖出回购金融资产款': 300000,
            '交易性金融负债-买断': 200000,
            '交易性金融负债-借券': 100000,
            '质押式逆回购': 50000,
            '买卖券净额': 50000,
            '借券卖空': 0,
            '买断逆回购': 0,
            '增减资': 0,
            '其他调整': 0
        },
        '银利5号': {
            '数据日期': '2026-01-31',
            '总资产': 2000000,
            '总负债': 800000,
            '卖出回购金融资产款': 400000,
            '交易性金融负债-买断': 250000,
            '交易性金融负债-借券': 150000,
            '质押式逆回购': 80000,
            '买卖券净额': -30000,
            '借券卖空': 20000,
            '买断逆回购': 10000,
            '增减资': 0,
            '其他调整': 0
        },
        '健康险': {
            '数据日期': '2026-01-31',
            '总资产': 5000000,
            '总负债': 3000000,
            '卖出回购金融资产款': 1500000,
            '交易性金融负债-买断': 1000000,
            '交易性金融负债-借券': 500000,
            '质押式逆回购': 200000,
            '买卖券净额': 100000,
            '借券卖空': 50000,
            '买断逆回购': 30000,
            '增减资': 0,
            '其他调整': 0
        }
    }
    
    for sheet_name, data in test_scenarios.items():
        print(f"\n处理 Sheet: {sheet_name}")
        ws = wb[sheet_name]
        
        # 填充数据日期
        ws['B3'] = data['数据日期']
        
        # 填充初始数据
        ws['B6'] = data['总资产']
        ws['B7'] = data['总负债']
        
        # 填充详细项目
        ws['B11'] = data['卖出回购金融资产款']
        ws['B12'] = data['交易性金融负债-买断']
        ws['B13'] = data['交易性金融负债-借券']
        ws['B14'] = data['质押式逆回购']
        
        # 填充当日操作
        ws['B19'] = data['买卖券净额']
        ws['B20'] = data['借券卖空']
        ws['B21'] = data['买断逆回购']
        ws['B22'] = data['增减资']
        ws['B23'] = data['其他调整']
        
        # 读取计算结果
        # 等待公式计算（在实际 Excel 中会自动计算）
        wb.save('/home/runner/work/invest/invest/leverage_calculation_template_with_data.xlsx')
        # 重新加载以获取计算结果
        wb_reload = load_workbook('/home/runner/work/invest/invest/leverage_calculation_template_with_data.xlsx', data_only=True)
        ws_reload = wb_reload[sheet_name]
        
        print(f"  ✓ 数据日期: {ws_reload['B3'].value}")
        print(f"  ✓ 总资产: {ws_reload['B6'].value:,.2f}")
        print(f"  ✓ 总负债: {ws_reload['B7'].value:,.2f}")
        print(f"  ✓ 净资产: {ws_reload['B8'].value:,.2f}" if ws_reload['B8'].value else "  ✓ 净资产: (公式)")
        print(f"  ✓ 基准杠杆率: {ws_reload['B9'].value:.2f}" if ws_reload['B9'].value else "  ✓ 基准杠杆率: (公式)")
        
        # 重新加载原始模板以继续填充其他账户
        wb = load_workbook('/home/runner/work/invest/invest/leverage_calculation_template_with_data.xlsx')
    
    print("\n" + "=" * 80)
    print("✓ 测试数据填充完成！")
    print("=" * 80)
    print(f"\n已创建测试文件: leverage_calculation_template_with_data.xlsx")
    print("\n可以在 Excel 中打开该文件查看：")
    print("  1. 所有公式是否正确计算")
    print("  2. 格式是否符合要求")
    print("  3. 预警功能是否正常工作")
    print("\n各账户预期结果：")
    print("  - 银利3号: 杠杆率应该超过 1.40 的限制，显示红色警告")
    print("  - 银利5号: 杠杆率情况需要在 Excel 中查看")
    print("  - 健康险: 杠杆率应该在 1.80 限制内")


if __name__ == "__main__":
    test_with_sample_data()
