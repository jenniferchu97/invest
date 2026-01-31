# invest
用于账户管理

## 杠杆率计算模板

本仓库包含一个完整的 Excel 杠杆率计算模板文件 `leverage_calculation_template.xlsx`，用于三个债券账户的实时杠杆率计算。

### 文件说明

- **leverage_calculation_template.xlsx**: 空白模板文件，可直接使用
- **leverage_calculation_template_with_data.xlsx**: 包含示例数据的模板文件，供参考
- **generate_leverage_template.py**: 生成模板文件的 Python 脚本
- **verify_template.py**: 验证模板文件的脚本
- **test_template_with_data.py**: 使用示例数据测试模板的脚本
- **TEMPLATE_USAGE.md**: 详细的使用说明文档

### 账户信息

| 账户名称 | 杠杆限制 |
|---------|---------|
| 银利3号 | 140% |
| 银利5号 | 140% |
| 健康险 | 180% |

### 快速开始

1. 打开 `leverage_calculation_template.xlsx` 文件
2. 选择要操作的账户 Sheet（银利3号、银利5号或健康险）
3. 填写初始资产负债表数据（黄色背景单元格）
4. 输入当日操作数据
5. 系统自动计算杠杆率和预警状态

### 详细使用说明

请参阅 [TEMPLATE_USAGE.md](TEMPLATE_USAGE.md) 文件获取详细的使用说明和示例。

### 生成新模板

```bash
python3 generate_leverage_template.py
```

### 验证模板

```bash
python3 verify_template.py
```

