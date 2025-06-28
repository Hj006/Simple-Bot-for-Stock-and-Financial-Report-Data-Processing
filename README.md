## Project: Simple-Bot-for-Stock-and-Financial-Report-Data-Processing

简易股票与财报数据处理机器人

### Description | 项目简介

This is a lightweight and extensible bot designed for handling stock and financial report data.
It enables efficient data processing, flexible indicator queries, and customized filtering logic.

这是一个轻量级、可扩展的机器人，专为处理股票与财务报表数据而设计。
它支持高效的数据批处理、灵活的指标查询以及自定义筛选逻辑。

---

### Features | 项目功能

* Search and retrieve specific stock or financial indicators
  搜索和提取特定的股票或财务指标

* Output companies and dates that meet user-defined criteria
  输出满足用户设定条件的公司及其财报时间

* Efficient batch processing of large-scale financial data
  高效地批量处理大规模财务数据

* User-defined indicator customization
  支持用户自定义指标与计算逻辑

---

## Planned Features | 开发功能计划

1. **Read and process raw data, calculate and save key indicators**
   读取原始股票与财报数据，自动计算关键财务指标并保存

2. **Allow user to select indicators and threshold conditions to filter results**
   用户可选择多个指标及其阈值，筛选出满足条件的公司与对应财报时间

3. **Query single data point by company, date, or metric**
   用户可以按公司、日期或指标名称查询某个具体值

4. **Support user-defined custom indicators (e.g., combining multiple metrics)**
   用户可定义新的指标（如多个财务比率组合），并将其用于查询或筛选
---

## 暂时的结构

project_root/
│
├── main.py
├── config.py                     # 配置：路径、指标定义
├── utils/
│   ├── excel_link_breaker.py     # 外链处理
│   ├── metric_inserter.py        # add_new_metrics 相关函数
│   └── formulas.py               # 所有指标计算逻辑
├── processor/
│   ├── process_excel.py          # 处理 Excel 核心逻辑
│   └── process_csv.py            # 处理 CSV 的逻辑（可选）
└── data/
    └── 原始数据模板_财务分析_2025Q1.xlsx
