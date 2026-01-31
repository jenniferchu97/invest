#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成杠杆率计算模板 Excel 文件
Generate leverage calculation template Excel file
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.styles.numbers import FORMAT_NUMBER_00

def create_account_sheet(wb, sheet_name, account_name, leverage_limit):
    """
    创建单个账户的 Sheet
    
    Args:
        wb: Workbook object
        sheet_name: Sheet 名称
        account_name: 账户名称
        leverage_limit: 杠杆限制百分比
    """
    ws = wb.create_sheet(title=sheet_name)
    
    # 定义样式
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    subheader_fill = PatternFill(start_color="B4C7E7", end_color="B4C7E7", fill_type="solid")
    subheader_font = Font(bold=True, size=11)
    input_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    formula_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    warning_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    warning_font = Font(bold=True, color="FFFFFF", size=12)
    
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    center_alignment = Alignment(horizontal='center', vertical='center')
    left_alignment = Alignment(horizontal='left', vertical='center')
    
    # 设置列宽
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 20
    
    row = 1
    
    # ========== 第一部分：初始资产负债表数据区 ==========
    ws.merge_cells(f'A{row}:E{row}')
    cell = ws[f'A{row}']
    cell.value = "第一部分：初始资产负债表数据区"
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = center_alignment
    row += 1
    
    # 账户名称
    ws[f'A{row}'] = "账户名称:"
    ws[f'A{row}'].font = Font(bold=True)
    ws[f'B{row}'] = account_name
    ws.merge_cells(f'B{row}:C{row}')
    ws[f'B{row}'].fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    row += 1
    
    # 数据日期
    ws[f'A{row}'] = "数据日期:"
    ws[f'A{row}'].font = Font(bold=True)
    ws[f'B{row}'].fill = input_fill
    ws.merge_cells(f'B{row}:C{row}')
    row += 1
    
    # 空行
    row += 1
    
    # 初始数据标题
    ws[f'A{row}'] = "初始数据"
    ws[f'A{row}'].font = subheader_font
    ws[f'A{row}'].fill = subheader_fill
    row += 1
    
    # 总资产
    ws[f'A{row}'] = "总资产"
    ws[f'B{row}'].fill = input_fill
    ws[f'B{row}'].number_format = FORMAT_NUMBER_00
    initial_total_asset_row = row
    row += 1
    
    # 总负债
    ws[f'A{row}'] = "总负债"
    ws[f'B{row}'].fill = input_fill
    ws[f'B{row}'].number_format = FORMAT_NUMBER_00
    initial_total_liability_row = row
    row += 1
    
    # 净资产（公式）
    ws[f'A{row}'] = "净资产"
    ws[f'B{row}'] = f"=B{initial_total_asset_row}-B{initial_total_liability_row}"
    ws[f'B{row}'].fill = formula_fill
    ws[f'B{row}'].number_format = FORMAT_NUMBER_00
    initial_net_asset_row = row
    row += 1
    
    # 基准杠杆率（公式）
    ws[f'A{row}'] = "基准杠杆率"
    ws[f'B{row}'] = f"=IFERROR(B{initial_total_asset_row}/B{initial_net_asset_row},0)"
    ws[f'B{row}'].fill = formula_fill
    ws[f'B{row}'].number_format = FORMAT_NUMBER_00
    row += 1
    
    # 空行
    row += 1
    
    # 详细项目标题
    ws[f'A{row}'] = "详细项目"
    ws[f'A{row}'].font = subheader_font
    ws[f'A{row}'].fill = subheader_fill
    row += 1
    
    # 卖出回购金融资产款
    ws[f'A{row}'] = "卖出回购金融资产款"
    ws[f'B{row}'].fill = input_fill
    ws[f'B{row}'].number_format = FORMAT_NUMBER_00
    repo_row = row
    row += 1
    
    # 交易性金融负债-买断
    ws[f'A{row}'] = "交易性金融负债-买断"
    ws[f'B{row}'].fill = input_fill
    ws[f'B{row}'].number_format = FORMAT_NUMBER_00
    buyout_row = row
    row += 1
    
    # 交易性金融负债-借券
    ws[f'A{row}'] = "交易性金融负债-借券"
    ws[f'B{row}'].fill = input_fill
    ws[f'B{row}'].number_format = FORMAT_NUMBER_00
    borrow_row = row
    row += 1
    
    # 质押式逆回购
    ws[f'A{row}'] = "质押式逆回购"
    ws[f'B{row}'].fill = input_fill
    ws[f'B{row}'].number_format = FORMAT_NUMBER_00
    reverse_repo_row = row
    row += 1
    
    # 空行
    row += 1
    
    # ========== 第二部分：当日操作输入区 ==========
    ws.merge_cells(f'A{row}:E{row}')
    cell = ws[f'A{row}']
    cell.value = "第二部分：当日操作输入区"
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = center_alignment
    row += 1
    
    # 表头
    ws[f'A{row}'] = "操作类型"
    ws[f'B{row}'] = "金额"
    ws[f'C{row}'] = "说明"
    for col in ['A', 'B', 'C']:
        ws[f'{col}{row}'].fill = subheader_fill
        ws[f'{col}{row}'].font = subheader_font
        ws[f'{col}{row}'].alignment = center_alignment
        ws[f'{col}{row}'].border = thin_border
    row += 1
    
    # 买卖券净额
    ws[f'A{row}'] = "买卖券净额"
    ws[f'B{row}'].fill = input_fill
    ws[f'B{row}'].number_format = FORMAT_NUMBER_00
    ws[f'C{row}'] = "正数为净买入，负数为净卖出"
    buy_sell_row = row
    row += 1
    
    # 借券卖空
    ws[f'A{row}'] = "借券卖空"
    ws[f'B{row}'].fill = input_fill
    ws[f'B{row}'].number_format = FORMAT_NUMBER_00
    ws[f'C{row}'] = "增加借券杠杆"
    borrow_short_row = row
    row += 1
    
    # 买断逆回购
    ws[f'A{row}'] = "买断逆回购"
    ws[f'B{row}'].fill = input_fill
    ws[f'B{row}'].number_format = FORMAT_NUMBER_00
    ws[f'C{row}'] = "增加回购和借券杠杆"
    buyout_reverse_row = row
    row += 1
    
    # 增减资
    ws[f'A{row}'] = "增减资"
    ws[f'B{row}'].fill = input_fill
    ws[f'B{row}'].number_format = FORMAT_NUMBER_00
    ws[f'C{row}'] = "正数增资，负数减资"
    capital_change_row = row
    row += 1
    
    # 其他调整
    ws[f'A{row}'] = "其他调整"
    ws[f'B{row}'].fill = input_fill
    ws[f'B{row}'].number_format = FORMAT_NUMBER_00
    ws[f'C{row}'] = ""
    other_adjust_row = row
    row += 1
    
    # 空行
    row += 1
    
    # ========== 第三部分：变化前后对比区 ==========
    ws.merge_cells(f'A{row}:E{row}')
    cell = ws[f'A{row}']
    cell.value = "第三部分：变化前后对比区"
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = center_alignment
    row += 1
    
    # 表头
    ws[f'A{row}'] = "指标"
    ws[f'B{row}'] = "变化前"
    ws[f'C{row}'] = "变化后"
    ws[f'D{row}'] = "变化量"
    for col in ['A', 'B', 'C', 'D']:
        ws[f'{col}{row}'].fill = subheader_fill
        ws[f'{col}{row}'].font = subheader_font
        ws[f'{col}{row}'].alignment = center_alignment
        ws[f'{col}{row}'].border = thin_border
    comparison_header_row = row
    row += 1
    
    # 总资产
    ws[f'A{row}'] = "总资产"
    ws[f'B{row}'] = f"=B{initial_total_asset_row}"
    ws[f'C{row}'] = f"=B{row}+B{buy_sell_row}+B{borrow_short_row}+B{buyout_reverse_row}"
    ws[f'D{row}'] = f"=C{row}-B{row}"
    for col in ['B', 'C', 'D']:
        ws[f'{col}{row}'].fill = formula_fill
        ws[f'{col}{row}'].number_format = FORMAT_NUMBER_00
    total_asset_after_row = row
    row += 1
    
    # 净资产
    ws[f'A{row}'] = "净资产"
    ws[f'B{row}'] = f"=B{initial_net_asset_row}"
    ws[f'C{row}'] = f"=B{row}+B{capital_change_row}+B{other_adjust_row}"
    ws[f'D{row}'] = f"=C{row}-B{row}"
    for col in ['B', 'C', 'D']:
        ws[f'{col}{row}'].fill = formula_fill
        ws[f'{col}{row}'].number_format = FORMAT_NUMBER_00
    net_asset_after_row = row
    row += 1
    
    # 杠杆率
    ws[f'A{row}'] = "杠杆率"
    ws[f'B{row}'] = f"=IFERROR(B{total_asset_after_row}/B{net_asset_after_row},0)"
    ws[f'C{row}'] = f"=IFERROR(C{total_asset_after_row}/C{net_asset_after_row},0)"
    ws[f'D{row}'] = f"=C{row}-B{row}"
    for col in ['B', 'C', 'D']:
        ws[f'{col}{row}'].fill = formula_fill
        ws[f'{col}{row}'].number_format = FORMAT_NUMBER_00
    leverage_after_row = row
    row += 1
    
    # 回购规模
    ws[f'A{row}'] = "回购规模"
    ws[f'B{row}'] = f"=B{repo_row}"
    ws[f'C{row}'] = f"=B{row}+B{buyout_reverse_row}"
    ws[f'D{row}'] = f"=C{row}-B{row}"
    for col in ['B', 'C', 'D']:
        ws[f'{col}{row}'].fill = formula_fill
        ws[f'{col}{row}'].number_format = FORMAT_NUMBER_00
    repo_after_row = row
    row += 1
    
    # 回购杠杆
    ws[f'A{row}'] = "回购杠杆"
    ws[f'B{row}'] = f"=IFERROR(B{repo_after_row}/B{net_asset_after_row},0)"
    ws[f'C{row}'] = f"=IFERROR(C{repo_after_row}/C{net_asset_after_row},0)"
    ws[f'D{row}'] = f"=C{row}-B{row}"
    for col in ['B', 'C', 'D']:
        ws[f'{col}{row}'].fill = formula_fill
        ws[f'{col}{row}'].number_format = FORMAT_NUMBER_00
    row += 1
    
    # 借券规模
    ws[f'A{row}'] = "借券规模"
    ws[f'B{row}'] = f"=B{buyout_row}+B{borrow_row}"
    ws[f'C{row}'] = f"=B{row}+B{borrow_short_row}+B{buyout_reverse_row}"
    ws[f'D{row}'] = f"=C{row}-B{row}"
    for col in ['B', 'C', 'D']:
        ws[f'{col}{row}'].fill = formula_fill
        ws[f'{col}{row}'].number_format = FORMAT_NUMBER_00
    borrow_after_row = row
    row += 1
    
    # 借券杠杆
    ws[f'A{row}'] = "借券杠杆"
    ws[f'B{row}'] = f"=IFERROR(B{borrow_after_row}/B{net_asset_after_row},0)"
    ws[f'C{row}'] = f"=IFERROR(C{borrow_after_row}/C{net_asset_after_row},0)"
    ws[f'D{row}'] = f"=C{row}-B{row}"
    for col in ['B', 'C', 'D']:
        ws[f'{col}{row}'].fill = formula_fill
        ws[f'{col}{row}'].number_format = FORMAT_NUMBER_00
    row += 1
    
    # 质押式逆回购规模
    ws[f'A{row}'] = "质押式逆回购规模"
    ws[f'B{row}'] = f"=B{reverse_repo_row}"
    ws[f'C{row}'] = f"=B{row}"  # 质押式逆回购不受当日操作影响
    ws[f'D{row}'] = f"=C{row}-B{row}"
    for col in ['B', 'C', 'D']:
        ws[f'{col}{row}'].fill = formula_fill
        ws[f'{col}{row}'].number_format = FORMAT_NUMBER_00
    reverse_repo_after_row = row
    row += 1
    
    # 质押式逆回购杠杆
    ws[f'A{row}'] = "质押式逆回购杠杆"
    ws[f'B{row}'] = f"=IFERROR(B{reverse_repo_after_row}/B{net_asset_after_row},0)"
    ws[f'C{row}'] = f"=IFERROR(C{reverse_repo_after_row}/C{net_asset_after_row},0)"
    ws[f'D{row}'] = f"=C{row}-B{row}"
    for col in ['B', 'C', 'D']:
        ws[f'{col}{row}'].fill = formula_fill
        ws[f'{col}{row}'].number_format = FORMAT_NUMBER_00
    row += 1
    
    # 空行
    row += 1
    
    # ========== 第四部分：杠杆限制预警区 ==========
    ws.merge_cells(f'A{row}:E{row}')
    cell = ws[f'A{row}']
    cell.value = "第四部分：杠杆限制预警区"
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = center_alignment
    row += 1
    
    # 当前杠杆率
    ws[f'A{row}'] = "当前杠杆率:"
    ws[f'A{row}'].font = Font(bold=True, size=12)
    ws[f'B{row}'] = f"=C{leverage_after_row}"
    ws[f'B{row}'].number_format = FORMAT_NUMBER_00
    ws[f'B{row}'].font = Font(bold=True, size=12)
    current_leverage_cell = f'B{row}'
    row += 1
    
    # 杠杆限制
    ws[f'A{row}'] = "杠杆限制:"
    ws[f'A{row}'].font = Font(bold=True, size=12)
    ws[f'B{row}'] = leverage_limit / 100.0  # 转换为小数
    ws[f'B{row}'].number_format = FORMAT_NUMBER_00
    ws[f'B{row}'].font = Font(bold=True, size=12)
    limit_cell = f'B{row}'
    row += 1
    
    # 预警状态
    ws[f'A{row}'] = "预警状态:"
    ws[f'A{row}'].font = Font(bold=True, size=12)
    ws.merge_cells(f'B{row}:D{row}')
    ws[f'B{row}'] = f'=IF({current_leverage_cell}>{limit_cell},"超过杠杆限制！","正常")'
    ws[f'B{row}'].font = Font(bold=True, size=12)
    
    # 添加条件格式：如果超过限制，显示红色
    from openpyxl.formatting.rule import CellIsRule
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    red_text = Font(color="FFFFFF", bold=True, size=12)
    ws.conditional_formatting.add(f'B{row}:D{row}',
                                  CellIsRule(operator='containsText',
                                            formula=['"超过"'],
                                            fill=red_fill,
                                            font=red_text))
    
    # 冻结窗格（冻结前3行和第一列）
    ws.freeze_panes = 'B4'
    
    return ws


def main():
    """主函数：创建完整的杠杆率计算模板"""
    
    # 创建工作簿
    wb = Workbook()
    
    # 删除默认的 Sheet
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])
    
    # 创建三个账户的 Sheet
    accounts = [
        ("银利3号", "银利3号", 140),
        ("银利5号", "银利5号", 140),
        ("健康险", "健康险", 180)
    ]
    
    for sheet_name, account_name, leverage_limit in accounts:
        create_account_sheet(wb, sheet_name, account_name, leverage_limit)
    
    # 保存文件
    output_file = "/home/runner/work/invest/invest/leverage_calculation_template.xlsx"
    wb.save(output_file)
    print(f"模板文件已创建: {output_file}")
    print(f"包含 {len(accounts)} 个账户的 Sheet:")
    for sheet_name, account_name, leverage_limit in accounts:
        print(f"  - {sheet_name}: {account_name}, 杠杆限制 {leverage_limit}%")


if __name__ == "__main__":
    main()
