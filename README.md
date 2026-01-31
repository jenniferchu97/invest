# invest - 账户管理系统

用于债券账户的杠杆率计算和管理。

## 功能

- **杠杆率实时计算**：支持三个债券账户（银利3号、银利5号、健康险）的杠杆率实时计算
- **操作影响分析**：自动计算各类操作（买卖券、借券卖空、增减资等）对杠杆率的影响
- **变化前后对比**：展示总资产、净资产、各类杠杆率等指标的变化情况
- **超限预警**：当杠杆率超过限制时自动发出警告

## 文件说明

- `leverage_calculation_template.xlsx` - Excel 杠杆率计算模板
- `generate_leverage_template.py` - 模板生成脚本
- `使用说明.md` - 详细使用说明文档

## 快速开始

1. 打开 `leverage_calculation_template.xlsx`
2. 选择对应的账户工作表（银利3号、银利5号或健康险）
3. 在"初始资产负债表数据"区域输入当日数据
4. 在"当日操作输入区"录入操作金额
5. 查看"变化前后对比"和"杠杆限制预警"

详细使用方法请参考 [使用说明.md](使用说明.md)

## 账户配置

| 账户名称 | 杠杆限制 |
|---------|---------|
| 银利3号  | 140%   |
| 银利5号  | 140%   |
| 健康险   | 180%   |

## 重新生成模板

```bash
python3 generate_leverage_template.py
```

## 依赖

- Python 3.x
- openpyxl

安装依赖：
```bash
pip install -r requirements.txt
```

## 文件结构

```
invest/
├── README.md                              # 项目说明
├── 使用说明.md                            # 详细使用说明
├── requirements.txt                       # Python 依赖
├── generate_leverage_template.py          # 模板生成脚本
├── test_template.py                       # 模板测试脚本
├── create_demo.py                         # 创建示例数据脚本
├── leverage_calculation_template.xlsx     # Excel 杠杆率计算模板（空模板）
└── .gitignore                             # Git 忽略文件配置
```

## 使用示例

### 生成空白模板
```bash
python3 generate_leverage_template.py
```

### 生成带示例数据的模板
```bash
python3 create_demo.py
```

### 测试模板结构
```bash
python3 test_template.py
```
