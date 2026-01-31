#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
杠杆率计算模板生成器
生成用于三个债券账户的实时杠杆率计算和管理的 Excel 模板
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 账户配置
ACCOUNTS = [
    {"name": "银利3号", "leverage_limit": 1.40},
    {"name": "银利5号", "leverage_limit": 1.40},
    {"name": "健康险", "leverage_limit": 1.80}
]


def set_cell_style(cell, bold=False, bg_color=None, font_color="000000", 
                   border=True, align_center=False):
    """设置单元格样式"""
    cell.font = Font(name="微软雅黑", size=10, bold=bold, color=font_color)
    if bg_color:
        cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, 
                               fill_type="solid")
    if border:
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        cell.border = thin_border
    if align_center:
        cell.alignment = Alignment(horizontal='center', vertical='center')
    else:
        cell.alignment = Alignment(vertical='center')


def create_account_sheet(wb, account_name, leverage_limit):
    """为每个账户创建一个工作表"""
    ws = wb.create_sheet(title=account_name)
    
    # 设置列宽
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 18
    
    row = 1
    
    # 标题
    ws.merge_cells(f'A{row}:E{row}')
    cell = ws[f'A{row}']
    cell.value = f"{account_name} - 杠杆率计算表"
    set_cell_style(cell, bold=True, bg_color="4472C4", font_color="FFFFFF", 
                   align_center=True)
    ws.row_dimensions[row].height = 25
    row += 1
    
    # 杠杆限制说明
    ws.merge_cells(f'A{row}:E{row}')
    cell = ws[f'A{row}']
    cell.value = f"杠杆限制：{leverage_limit*100:.0f}%"
    set_cell_style(cell, bold=True, bg_color="D9E1F2", align_center=True)
    row += 1
    
    # 空行
    row += 1
    
    # ========== 第一部分：初始资产负债表数据 ==========
    ws.merge_cells(f'A{row}:E{row}')
    cell = ws[f'A{row}']
    cell.value = "一、初始资产负债表数据"
    set_cell_style(cell, bold=True, bg_color="FFC000", align_center=True)
    row += 1
    
    # 资产负债表数据标题
    headers = ["科目", "金额（万元）", "", "", ""]
    for col, header in enumerate(headers, start=1):
        if header:
            cell = ws.cell(row=row, column=col)
            cell.value = header
            set_cell_style(cell, bold=True, bg_color="E7E6E6", align_center=True)
    row += 1
    
    # 初始数据输入区域
    initial_data_start_row = row
    balance_sheet_items = [
        "总资产",
        "卖出回购金融资产款",
        "交易性金融负债（买断）",
        "交易性金融负债（借券）",
        "买入返售金融资产",
        "其他负债",
        "总负债",
        "净资产"
    ]
    
    for item in balance_sheet_items:
        cell = ws[f'A{row}']
        cell.value = item
        set_cell_style(cell, bold=True if item in ["总资产", "总负债", "净资产"] else False)
        
        cell = ws[f'B{row}']
        if item == "总负债":
            # 总负债 = 卖出回购 + 交易性金融负债（买断） + 交易性金融负债（借券） + 其他负债
            cell.value = f"=B{initial_data_start_row+1}+B{initial_data_start_row+2}+B{initial_data_start_row+3}+B{initial_data_start_row+5}"
        elif item == "净资产":
            # 净资产 = 总资产 - 总负债
            cell.value = f"=B{initial_data_start_row}-B{initial_data_start_row+6}"
        else:
            cell.value = 0
        set_cell_style(cell, bold=True if item in ["总资产", "总负债", "净资产"] else False)
        row += 1
    
    # 初始杠杆率
    cell = ws[f'A{row}']
    cell.value = "初始杠杆率"
    set_cell_style(cell, bold=True, bg_color="FFF2CC")
    cell = ws[f'B{row}']
    # 杠杆率 = 总资产 / 净资产
    cell.value = f"=IF(B{initial_data_start_row+7}=0,0,B{initial_data_start_row}/B{initial_data_start_row+7})"
    cell.number_format = '0.00%'
    set_cell_style(cell, bold=True, bg_color="FFF2CC")
    row += 1
    
    # 空行
    row += 1
    
    # ========== 第二部分：当日操作输入区 ==========
    ws.merge_cells(f'A{row}:E{row}')
    cell = ws[f'A{row}']
    cell.value = "二、当日操作输入区"
    set_cell_style(cell, bold=True, bg_color="FFC000", align_center=True)
    row += 1
    
    # 操作输入标题
    operation_start_row = row
    headers = ["操作类型", "金额（万元）", "说明", "", ""]
    for col, header in enumerate(headers, start=1):
        if header:
            cell = ws.cell(row=row, column=col)
            cell.value = header
            set_cell_style(cell, bold=True, bg_color="E7E6E6", align_center=True)
    row += 1
    
    # 操作输入项
    operations = [
        ("买入券", "净买入增加总资产，相应增加净资产"),
        ("卖出券", "净卖出减少总资产，相应减少净资产"),
        ("债券借贷（借券卖空）", "增加交易性金融负债（借券）"),
        ("买断逆回购借券", "当日增加买入返售 + 交易性金融负债（买断）"),
        ("买断逆回购借券释放", "第二天释放买入返售和交易性金融负债（买断）"),
        ("卖出回购融资", "增加卖出回购金融资产款"),
        ("卖出回购到期归还", "减少卖出回购金融资产款"),
        ("增资", "增加净资产和总资产"),
        ("减资", "减少净资产和总资产")
    ]
    
    for op_name, desc in operations:
        cell = ws[f'A{row}']
        cell.value = op_name
        set_cell_style(cell)
        
        cell = ws[f'B{row}']
        cell.value = 0
        set_cell_style(cell)
        
        ws.merge_cells(f'C{row}:E{row}')
        cell = ws[f'C{row}']
        cell.value = desc
        set_cell_style(cell)
        cell.font = Font(name="微软雅黑", size=9, italic=True, color="666666")
        row += 1
    
    # 空行
    row += 1
    
    # ========== 第三部分：变化后数据对比 ==========
    ws.merge_cells(f'A{row}:E{row}')
    cell = ws[f'A{row}']
    cell.value = "三、变化前后对比"
    set_cell_style(cell, bold=True, bg_color="FFC000", align_center=True)
    row += 1
    
    # 对比表标题
    comparison_start_row = row
    headers = ["指标", "变化前", "变化后", "变化额", "变化率"]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=col)
        cell.value = header
        set_cell_style(cell, bold=True, bg_color="E7E6E6", align_center=True)
    row += 1
    
    # 计算变化后的值
    # 需要引用初始数据和操作输入
    comparison_items = [
        "总资产",
        "净资产",
        "杠杆率",
        "卖出回购金融资产款",
        "回购杠杆",
        "交易性金融负债（买断）",
        "买断杠杆",
        "交易性金融负债（借券）",
        "借券杠杆",
        "买入返售金融资产",
        "质押式逆回购杠杆"
    ]
    
    # 操作行引用
    op_buy = operation_start_row + 1  # 买入券
    op_sell = operation_start_row + 2  # 卖出券
    op_borrow = operation_start_row + 3  # 债券借贷
    op_reverse_repo = operation_start_row + 4  # 买断逆回购借券
    op_reverse_release = operation_start_row + 5  # 买断逆回购释放
    op_repo_in = operation_start_row + 6  # 卖出回购融资
    op_repo_out = operation_start_row + 7  # 卖出回购到期归还
    op_capital_in = operation_start_row + 8  # 增资
    op_capital_out = operation_start_row + 9  # 减资
    
    for item in comparison_items:
        cell = ws[f'A{row}']
        cell.value = item
        is_bold = item in ["总资产", "净资产", "杠杆率"]
        set_cell_style(cell, bold=is_bold)
        
        # 变化前
        cell = ws[f'B{row}']
        if item == "总资产":
            cell.value = f"=B{initial_data_start_row}"
        elif item == "净资产":
            cell.value = f"=B{initial_data_start_row+7}"
        elif item == "杠杆率":
            cell.value = f"=IF(B{comparison_start_row+2}=0,0,B{comparison_start_row+1}/B{comparison_start_row+2})"
            cell.number_format = '0.00%'
        elif item == "卖出回购金融资产款":
            cell.value = f"=B{initial_data_start_row+1}"
        elif item == "回购杠杆":
            cell.value = f"=IF(B{comparison_start_row+2}=0,0,B{comparison_start_row+4}/B{comparison_start_row+2})"
            cell.number_format = '0.00%'
        elif item == "交易性金融负债（买断）":
            cell.value = f"=B{initial_data_start_row+2}"
        elif item == "买断杠杆":
            cell.value = f"=IF(B{comparison_start_row+2}=0,0,B{comparison_start_row+6}/B{comparison_start_row+2})"
            cell.number_format = '0.00%'
        elif item == "交易性金融负债（借券）":
            cell.value = f"=B{initial_data_start_row+3}"
        elif item == "借券杠杆":
            cell.value = f"=IF(B{comparison_start_row+2}=0,0,B{comparison_start_row+8}/B{comparison_start_row+2})"
            cell.number_format = '0.00%'
        elif item == "买入返售金融资产":
            cell.value = f"=B{initial_data_start_row+4}"
        elif item == "质押式逆回购杠杆":
            cell.value = f"=IF(B{comparison_start_row+2}=0,0,B{comparison_start_row+10}/B{comparison_start_row+2})"
            cell.number_format = '0.00%'
        set_cell_style(cell, bold=is_bold)
        
        # 变化后
        cell = ws[f'C{row}']
        if item == "总资产":
            # 总资产变化 = 买入 - 卖出 + 买断逆回购 - 买断释放 + 增资 - 减资
            cell.value = f"=B{comparison_start_row+1}+B{op_buy}-B{op_sell}+B{op_reverse_repo}-B{op_reverse_release}+B{op_capital_in}-B{op_capital_out}"
        elif item == "净资产":
            # 净资产变化 = 买入 - 卖出 + 增资 - 减资
            cell.value = f"=B{comparison_start_row+2}+B{op_buy}-B{op_sell}+B{op_capital_in}-B{op_capital_out}"
        elif item == "杠杆率":
            cell.value = f"=IF(C{comparison_start_row+2}=0,0,C{comparison_start_row+1}/C{comparison_start_row+2})"
            cell.number_format = '0.00%'
        elif item == "卖出回购金融资产款":
            # 卖出回购 = 初始 + 新增融资 - 到期归还
            cell.value = f"=B{comparison_start_row+4}+B{op_repo_in}-B{op_repo_out}"
        elif item == "回购杠杆":
            cell.value = f"=IF(C{comparison_start_row+2}=0,0,C{comparison_start_row+4}/C{comparison_start_row+2})"
            cell.number_format = '0.00%'
        elif item == "交易性金融负债（买断）":
            # 买断负债 = 初始 + 买断逆回购 - 买断释放
            cell.value = f"=B{comparison_start_row+6}+B{op_reverse_repo}-B{op_reverse_release}"
        elif item == "买断杠杆":
            cell.value = f"=IF(C{comparison_start_row+2}=0,0,C{comparison_start_row+6}/C{comparison_start_row+2})"
            cell.number_format = '0.00%'
        elif item == "交易性金融负债（借券）":
            # 借券负债 = 初始 + 新增借券
            cell.value = f"=B{comparison_start_row+8}+B{op_borrow}"
        elif item == "借券杠杆":
            cell.value = f"=IF(C{comparison_start_row+2}=0,0,C{comparison_start_row+8}/C{comparison_start_row+2})"
            cell.number_format = '0.00%'
        elif item == "买入返售金融资产":
            # 买入返售 = 初始 + 买断逆回购 - 买断释放
            cell.value = f"=B{comparison_start_row+10}+B{op_reverse_repo}-B{op_reverse_release}"
        elif item == "质押式逆回购杠杆":
            cell.value = f"=IF(C{comparison_start_row+2}=0,0,C{comparison_start_row+10}/C{comparison_start_row+2})"
            cell.number_format = '0.00%'
        set_cell_style(cell, bold=is_bold)
        
        # 变化额
        cell = ws[f'D{row}']
        cell.value = f"=C{row}-B{row}"
        if item in ["杠杆率", "回购杠杆", "买断杠杆", "借券杠杆", "质押式逆回购杠杆"]:
            cell.number_format = '0.00%'
        set_cell_style(cell)
        
        # 变化率
        cell = ws[f'E{row}']
        cell.value = f"=IF(B{row}=0,0,(C{row}-B{row})/B{row})"
        cell.number_format = '0.00%'
        set_cell_style(cell)
        
        row += 1
    
    # 空行
    row += 1
    
    # ========== 第四部分：杠杆限制预警 ==========
    ws.merge_cells(f'A{row}:E{row}')
    cell = ws[f'A{row}']
    cell.value = "四、杠杆限制预警"
    set_cell_style(cell, bold=True, bg_color="FFC000", align_center=True)
    row += 1
    
    # 预警区域
    warning_start_row = row
    
    # 杠杆限制
    cell = ws[f'A{row}']
    cell.value = "杠杆限制"
    set_cell_style(cell, bold=True)
    cell = ws[f'B{row}']
    cell.value = leverage_limit
    cell.number_format = '0.00%'
    set_cell_style(cell, bold=True)
    row += 1
    
    # 当前杠杆率
    cell = ws[f'A{row}']
    cell.value = "当前杠杆率"
    set_cell_style(cell, bold=True)
    cell = ws[f'B{row}']
    cell.value = f"=C{comparison_start_row+3}"
    cell.number_format = '0.00%'
    set_cell_style(cell, bold=True)
    row += 1
    
    # 距离限制
    cell = ws[f'A{row}']
    cell.value = "距离限制"
    set_cell_style(cell)
    cell = ws[f'B{row}']
    cell.value = f"=B{warning_start_row}-B{warning_start_row+1}"
    cell.number_format = '0.00%'
    set_cell_style(cell)
    row += 1
    
    # 预警状态
    cell = ws[f'A{row}']
    cell.value = "预警状态"
    set_cell_style(cell, bold=True)
    ws.merge_cells(f'B{row}:E{row}')
    cell = ws[f'B{row}']
    # 如果当前杠杆率 > 限制，显示"超限警告"，否则显示"正常"
    cell.value = f'=IF(B{warning_start_row+1}>B{warning_start_row},"⚠️ 超限警告！","✓ 正常")'
    set_cell_style(cell, bold=True)
    
    # 设置条件格式（通过背景色）
    # 注意：openpyxl 不支持直接设置条件格式公式，所以我们使用规则
    from openpyxl.formatting.rule import CellIsRule
    
    # 对预警状态单元格应用条件格式
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
    
    # 添加条件格式规则
    ws.conditional_formatting.add(
        f'B{warning_start_row+3}:E{warning_start_row+3}',
        CellIsRule(
            operator='containsText',
            formula=['"超限"'],
            fill=red_fill,
            font=Font(color="FFFFFF", bold=True)
        )
    )
    ws.conditional_formatting.add(
        f'B{warning_start_row+3}:E{warning_start_row+3}',
        CellIsRule(
            operator='containsText',
            formula=['"正常"'],
            fill=green_fill,
            font=Font(color="FFFFFF", bold=True)
        )
    )
    
    return ws


def main():
    """生成杠杆率计算模板"""
    wb = Workbook()
    
    # 移除默认的工作表
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])
    
    # 为每个账户创建工作表
    for account in ACCOUNTS:
        create_account_sheet(wb, account['name'], account['leverage_limit'])
    
    # 保存工作簿
    output_file = "leverage_calculation_template.xlsx"
    wb.save(output_file)
    print(f"✓ Excel 模板已生成: {output_file}")
    print(f"✓ 包含 {len(ACCOUNTS)} 个账户工作表：")
    for account in ACCOUNTS:
        print(f"  - {account['name']} (杠杆限制: {account['leverage_limit']*100:.0f}%)")


if __name__ == "__main__":
    main()
