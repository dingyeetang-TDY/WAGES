# Attendance Hours Calculator

> 员工打卡考勤工时计算工具包 — 从原始打卡机导出文件到可视化报表的全流程自动化处理。

[![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)]()

---

## 目录

- [项目简介](#项目简介)
- [功能一览](#功能一览)
- [安装](#安装)
- [快速开始](#快速开始)
- [完整用法](#完整用法)
- [项目结构](#项目结构)
- [计算规则](#计算规则)
- [输入文件格式](#输入文件格式)
- [输出文件格式](#输出文件格式)
- [薪资配置](#薪资配置)
- [数据流程图](#数据流程图)
- [测试](#测试)
- [常见问题](#常见问题)
- [许可协议](#许可协议)

---

## 项目简介

本项目是一个 Python 工具包，用于处理企业员工打卡考勤数据。它将原始打卡机导出的 `.xls` 文件，经过解析、清洗、配对、计算，最终生成结构化的工时报表和可视化 HTML 报告。

**核心能力：**
- 将多 Sheet 的原始 `.xls` 转换为"员工 × 日期"矩阵格式
- 自动检测奇数打卡异常（无法配对的打卡记录）并高亮标注
- 按"上午班 / 中场休息 / 下午班 / 中场休息2 / 下午班2"分段计算净工时
- 生成带统计卡片、异常汇总、每日打卡明细、汇总表格的 HTML 报表
- 对 Excel 文件应用专业美化样式（斑马纹、标题栏、缺勤高亮等）
- 薪资计算引擎：支持全职/兼职、加班 OT、公共假期双倍、全勤奖

---

## 功能一览

| 模块 | 脚本 | 功能 |
|------|------|------|
| 格式转换 | `convert_xls.py` | 原始 .xls → 矩阵 .xlsx（工号/姓名/部门 + 日期列） |
| 异常检测 | `highlight_odd_punches.py` | 在矩阵中高亮奇数打卡格（红色背景 + 批注），幂等可反复运行 |
| 工时计算 | `attendance_calculator.py` | 从矩阵计算工作/休息时长，输出 4 个 Sheet |
| 一条龙 | `run_attendance_pipeline.py` | 转换 → 高亮 →（手动修正）→ 计算 |
| JUN 格式 | `matrix_to_jun_format.py` | 矩阵 → 每人一个 Sheet（DAY/START/END/WORKING/BREAK/NET） |
| 6次打卡提取 | `extract_six_punch_employees.py` | 提取有 6 次打卡日的员工，独立显示各休息段 |
| 分段明细 | `extract_six_punch_segment_rows.py` | 6 次打卡员工的逐段明细（每段一行） |
| 薪资引擎 | `payroll_engine.py` | 全职/兼职薪资、加班 OT、公共假期、全勤奖 |
| 薪资处理 | `process_jun_payroll.py` | 读取 JUN 格式 → 引擎计算 → 输出汇总 + 明细 |
| 验证 | `verify_salary_xlsx.py` | 验证手工薪资表中 totals 是否正确 |
| HTML 报表 | `generate_report.py` | xlsx（Logs 表）→ 带标注的 HTML 报表（推荐） |
| JSON 解析 | `parse_logs.py` | xlsx/xls → JSON（两步流程的 Step 1） |
| JSON→HTML | `generate_report_from_json.py` | JSON → HTML（两步流程的 Step 2） |
| Excel 美化 | `beautify_xlsx.py` | 对 .xlsx 应用专业样式 |

---

## 安装

### 前置条件

- Python 3.10 或更高版本
- pip 包管理器

### 步骤

```bash
# 1. 克隆仓库
git clone https://github.com/YOUR_USERNAME/attendance-hours-calculator.git

# 2. 进入项目目录
cd attendance-hours-calculator

# 3. 安装依赖
pip install -r requirements.txt
```

### 依赖说明

| 包名 | 版本要求 | 用途 |
|------|---------|------|
| `openpyxl` | >=3.1.0 | 读写 .xlsx 文件，支持样式/合并单元格 |
| `pandas` | >=2.0.0 | DataFrame 数据处理 |
| `xlrd` | >=2.0.1 | 读取旧版 .xls 文件（打卡机原始导出格式） |
| `PyYAML` | >=6.0 | 解析薪资配置 YAML 文件 |
| `holidays` | >=0.40 | 获取各国公共假期（薪资计算用） |

---

## 快速开始

### 场景一：从 Excel 生成 HTML 工时报表（最常用）

```bash
# 输入一个含 Logs 表的 .xlsx 文件，直接生成 HTML 报表
python scripts/generate_report.py attendance.xlsx

# 指定输出路径和月份
python scripts/generate_report.py attendance.xlsx -o report.html --month 7
```

### 场景二：从原始 .xls 走完整流水线

```bash
# Step 1-2: 转换格式 + 高亮奇数
python src/run_attendance_pipeline.py raw_export.xls

# Step 3: 在 Excel 中手动修正红色单元格（补录或删除一个打卡时间）

# Step 4-5: 刷新高亮 + 计算工时
python src/run_attendance_pipeline.py corrected.xlsx --calculate --header-row 2
```

---

## 完整用法

### 1. 分步执行（XLS → 矩阵 → 工时）

```bash
# Step 1: 转换 XLS 为矩阵
python src/convert_xls.py input.xls -o matrix.xlsx

# Step 2: 高亮奇数打卡
python src/highlight_odd_punches.py matrix.xlsx -o highlighted.xlsx

# Step 3: 手动修正（在 Excel 中操作）

# Step 4: 计算工时
python src/attendance_calculator.py highlighted.xlsx --header-row 2 -o results.xlsx

# 可选：自定义标准上班/下班时间
python src/attendance_calculator.py highlighted.xlsx --header-row 2 --am-in 09:00 --pm-out 18:00
```

### 2. 两步流程（xlsx → JSON → HTML）

```bash
# Step 1: 解析为 JSON
python scripts/parse_logs.py attendance.xlsx -o parsed.json --month 7

# Step 2: 从 JSON 生成 HTML
python scripts/generate_report_from_json.py parsed.json -o report.html
```

> 适用于需要中间检查或多次重生成报表（不重复解析 xlsx）的场景。

### 3. 矩阵 → JUN 格式 → 薪资计算

```bash
# 矩阵转 JUN 格式（每人一个 Sheet）
python src/matrix_to_jun_format.py matrix.xlsx -o jun_format.xlsx

# 薪资计算
python src/process_jun_payroll.py jun_format.xlsx -o payroll.xlsx --year 2026 --month 6
```

### 4. 提取 6 次打卡员工

```bash
# 独立休息段版（每员工一 Sheet，含 BREAK 1/BREAK 2 列）
python src/extract_six_punch_employees.py matrix.xlsx -o six_punch.xlsx

# 逐段明细版（每段一行：WORK/BREAK 分类）
python src/extract_six_punch_segment_rows.py matrix.xlsx -o segments.xlsx
```

### 5. Excel 美化

```bash
python scripts/beautify_xlsx.py plain.xlsx -o styled.xlsx
```

### 6. 运行测试

```bash
python -m pytest tests/test_attendance_calculator.py -v
```

---

## 项目结构

```
attendance-hours-calculator/
│
├── src/                              # 核心模块（可被 import 或独立运行）
│   ├── convert_xls.py                # .xls → 矩阵 .xlsx
│   ├── highlight_odd_punches.py     # 奇数打卡高亮标注
│   ├── attendance_calculator.py     # 工时计算器（核心）
│   ├── run_attendance_pipeline.py   # 一条龙流水线
│   ├── matrix_to_jun_format.py      # 矩阵 → JUN 格式
│   ├── extract_six_punch_employees.py
│   ├── extract_six_punch_segment_rows.py
│   ├── payroll_engine.py            # 薪资计算引擎
│   ├── process_jun_payroll.py       # JUN 薪资处理
│   └── verify_salary_xlsx.py        # 薪资验证
│
├── scripts/                          # 报表与样式脚本
│   ├── generate_report.py           # xlsx → HTML（推荐）
│   ├── parse_logs.py                # xlsx/xls → JSON
│   ├── generate_report_from_json.py # JSON → HTML
│   └── beautify_xlsx.py            # Excel 美化
│
├── config/                           # 配置文件
│   └── payroll_config.yaml           # 薪资规则模板
│
├── tests/                            # 测试
│   └── test_attendance_calculator.py
│
├── .gitignore                        # 忽略规则（屏蔽所有数据文件）
├── requirements.txt                  # 依赖清单
└── README.md                         # 本文件
```

---

## 计算规则

### 规则一：奇数打卡异常检测

如果某日的有效打卡时间数量 **不是偶数**（即奇数），则标注为异常打卡。

奇数打卡意味着存在无法配对的"孤儿"打卡记录，可能是漏打卡或多打了一次。

### 规则二：工时分段计算

打卡时间按时间顺序排列后，两两配对为工作段和休息段：

```
打卡次数    配对方式                              计入工时的段
─────────────────────────────────────────────────────────────────
  2 次      [1→2] 工作                             段 1
  4 次      [1→2]工作 [2→3]休息 [3→4]工作          段 1 + 段 3
  6 次      [1→2]工作 [2→3]休息 [3→4]工作          段 1 + 段 3 + 段 5
            [4→5]休息2 [5→6]工作2
  8+ 次     模式扩展：奇数对=工作，偶数对=休息       所有奇数对之和
```

**休息段（中场休息 / 中场休息2）不计入工时。**

### 规则三（attendance_calculator.py 额外功能）

- 迟到检测：首次打卡晚于标准上班时间（默认 08:30）
- 早退检测：末次打卡早于标准下班时间（默认 17:30）
- 奇数打卡日：工时和休息均计为 0（不计算任何分段）

### 分段命名对照

| 打卡次数 | 段 1 | 段 2 | 段 3 | 段 4 | 段 5 |
|---------|------|------|------|------|------|
| 2 次 | 上午班 | — | — | — | — |
| 4 次 | 上午班 | 中场休息 | 下午班 | — | — |
| 6 次 | 上午班 | 中场休息 | 下午班 | 中场休息2 | 下午班2 |

---

## 输入文件格式

### 原始 .xls 文件（打卡机导出）

包含以下 Sheet：

| Sheet | 用途 | 关键列位置 |
|-------|------|-----------|
| **Summary** | 员工汇总（工号/姓名/部门/应到工时/实到工时/迟到次数等） | 第 4 行起，A=工号, B=姓名, C=部门 |
| **Logs** | 每日打卡明细 | 见下方详细说明 |
| **日期分组表**（如 21.29.31） | 按日期范围分组的排班/打卡表 | 每组 14 列宽 |

### Logs 表结构（核心数据源）

```
行 3 (0-indexed):  日号（整数 1, 2, 3, ... 18）
行 4:              "No :" 标记行
                   A列="No :"  C列=工号  K列(10)=姓名  T列(19)=部门
行 5:              数据行（打卡时间）
                   B列(1)~S列(18) = 各日打卡时间
                   每格内换行分隔多个 HH:MM 时间
行 6:              下一个员工的 "No :" 行
...
```

### 矩阵 .xlsx 格式（转换后）

| 工号 | 姓名 | 部门 | 07/01 | 07/02 | ... | 07/18 |
|------|------|------|-------|-------|-----|-------|
| 1 | 张三 | 生产部 | 08:30\n17:30 | 08:25\n12:00\n13:00\n17:35 | ... | |
| 2 | 李四 | 人事部 | 09:00\n18:00 | | ... | 08:50\n17:20 |

---

## 输出文件格式

### HTML 报表（generate_report.py）

| 区域 | 内容 |
|------|------|
| 顶部统计卡片 | 员工总数 / 考勤天数 / 总出勤人次 / 奇数打卡异常数 / 总工时 |
| 规则说明 | 工时计算规则图例 |
| 异常汇总 | 奇数打卡员工列表（姓名 / 工号 / 部门 / 异常日期与次数） |
| 每日打卡明细 | 每位员工每天一张卡片：打卡时间标签 + 分段着色（绿=工作/黄=休息/紫=异常）+ 工时汇总 |
| 底部汇总表 | No. / 姓名 / 部门 / 有打卡天数 / 总工时 / 总时长 / 有效占比 / 奇数异常天数 |

### Excel 工时结果（attendance_calculator.py）

| Sheet 名 | 内容 |
|----------|------|
| 打卡记录 | 原始打卡时间（每格换行分隔） |
| 每日工时 | 每人每天的工作时段工时（小时） |
| 休息时长 | 每人每天的休息时段时长（小时） |
| 汇总统计 | 出勤天数 / 奇数打卡天数 / 总工时 / 休息时长 / 在司总时长 / 迟到次数 / 早退次数 |

### JUN 格式（matrix_to_jun_format.py）

每位员工一个 Sheet，列布局：

| DAY | START TIME | END TIME | WORKING HOURS | BREAK TIME | NET WORKING HOURS |
|-----|-----------|---------|---------------|-----------|-------------------|
| 1 | 08:30 | 17:30 | 9:00 | 1:00 | 8:00 |
| 2 | 08:25 | 17:35 | 9:10 | 1:10 | 8:00 |

底部含 TOTAL WORKING HOURS / TOTAL BREAK TIME / WORK DAYS 汇总行。

---

## 薪资配置

编辑 `config/payroll_config.yaml` 自定义以下规则：

```yaml
# 标准上下班时间（迟到/早退检测用）
schedule:
  am_in: "08:30"
  pm_out: "17:30"

# 休息扣除规则
rest:
  standard_minutes_per_day: 60      # 每日标准休息 60 分钟
  deduction_mode: "exceed_only"     # 仅扣除超出标准的部分
  multi_segment_rule: "per_day"     # 多段休息按日汇总

# 全职规则
full_time:
  standard_hours_per_day: 7.5       # 每日标准 7.5 小时
  standard_days_per_month: null     # 自动从 Sheet 解析，或手动指定

# 兼职规则
part_time:
  default_hourly_rate: 8.00         # 默认时薪
  public_holiday_double_pay: true   # 假期双倍

# 公共假期
public_holiday:
  full_time_ph_as_ot: true          # 全职假期工作算加班

# 全勤奖
full_attendance_bonus:
  enabled: true
  no_late: true                     # 无迟到
  no_early_leave: true              # 无早退
  no_absence: true                  # 无缺勤
  no_odd_punch: true                # 无奇数打卡
```

---

## 数据流程图

### 主流程：Excel → HTML 报表

```
┌─────────────────┐     ┌──────────────────┐     ┌─────────────────┐
│  attendance.xlsx│────▶│ generate_report  │────▶│  report.html    │
│  (Logs 表)      │     │     .py          │     │  (可视化报表)    │
└─────────────────┘     └──────────────────┘     └─────────────────┘
                              │
                              ├─ 解析员工信息（No/姓名/部门）
                              ├─ 读取每日打卡时间
                              ├─ 奇数标注检测
                              ├─ 分段配对计算工时
                              └─ 生成 HTML（卡片+表格）
```

### 完整流水线：XLS → 矩阵 → 工时

```
┌──────────┐    ┌────────────┐    ┌────────────┐    ┌────────────┐    ┌──────────────┐
│ input.xls│───▶│ convert_xls│───▶│ highlight  │───▶│  手动修正   │───▶│  attendance  │
│ (原始)   │    │    .py     │    │ _odd.py    │    │ (Excel中)  │    │ _calculator  │
└──────────┘    └────────────┘    └────────────┘    └────────────┘    └──────────────┘
                转换为矩阵xlsx      红色高亮奇数格     补录/删除打卡      计算4个Sheet
```

### 薪资流程：矩阵 → JUN → 薪资

```
┌────────────┐    ┌──────────────────┐    ┌──────────────────┐    ┌──────────────┐
│ matrix.xlsx│───▶│ matrix_to_jun    │───▶│ process_jun      │───▶│ payroll.xlsx │
│            │    │ _format.py       │    │ _payroll.py      │    │ (汇总+明细)  │
└────────────┘    └──────────────────┘    └──────────────────┘    └──────────────┘
                  每人一个Sheet            引擎计算OT/假期/全勤
```

---

## 测试

```bash
# 运行全部测试
python -m pytest tests/test_attendance_calculator.py -v

# 测试覆盖：
# - 时间解析（HH:MM 格式、无效输入、边界值）
# - 2次打卡配对（全部算工时）
# - 4次打卡配对（工作+休息+工作）
# - 6次打卡配对（3工作+2休息）
# - 奇数打卡异常检测
# - 无序输入自动排序
# - 空输入处理
# - 重复/无效条目过滤
```

---

## 常见问题

### Q1: 为什么我的 .xls 文件读取报错？

确保已安装 `xlrd>=2.0.1`。xlrd 2.0+ 仅支持 `.xls` 格式（不支持 `.xlsx`）。如果是 `.xlsx` 文件，不需要 xlrd，openpyxl 会自动处理。

### Q2: 奇数打卡怎么修正？

在 Excel 中打开高亮后的文件，找到 **红色背景** 的单元格。每个红色单元格内的打卡时间为奇数个，需要补录或删除一个时间使其变为偶数。修正后保存，重新运行计算步骤即可。

### Q3: 如何自定义标准上班/下班时间？

```bash
# 通过命令行参数
python src/attendance_calculator.py input.xlsx --am-in 09:00 --pm-out 18:00

# 或通过薪资配置文件
# 编辑 config/payroll_config.yaml 中的 schedule.am_in / schedule.pm_out
```

### Q4: HTML 报表支持哪些浏览器？

报表使用标准 HTML5 + CSS3（Grid/Flexbox 布局），兼容 Chrome、Edge、Firefox、Safari 等现代浏览器。

### Q5: 如何处理多个月的考勤数据？

每月运行一次流程即可。通过 `--month` 参数指定月份编号，影响日期列标签（如 `07/01` vs `08/01`）：

```bash
python scripts/generate_report.py august.xlsx --month 8
```

### Q6: 数据文件会被上传到 GitHub 吗？

不会。`.gitignore` 已配置屏蔽所有 `.xls`、`.xlsx`、`.csv`、`.json`、`.html` 文件。只有代码和配置文件会被提交。

---

## 许可协议

[MIT License](LICENSE) — 可自由使用、修改、分发。

---

## 贡献

欢迎提交 Issue 或 Pull Request。请确保：
1. 不提交任何含个人信息的文件（姓名、工号、打卡时间等）
2. 新增功能需附带测试
3. 代码风格保持一致（PEP 8）
