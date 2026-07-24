# BOM Normalization Tool
## Product Requirement Document (PRD)

Version: V1.0
Author: Yammy Zhao
Last Update: 2026-07-24

---

# 1. Product Goal

BOM Normalization Tool 用于协助 Hardware RD 对 BOM 进行标准化分析，并提供 Merge 建议，降低料号数量，提高 BOM 一致性。

本 Tool 的定位：

- 自动整理 BOM
- 自动辨识 R / MLCC / CAP
- 自动统一规格表示
- 自动寻找可 Merge 的料件
- 自动输出 RD Review Report
- 降低 BOM Complexity
- 最终 Merge 决策仍由 RD 判断

本 Tool 不会自动修改 BOM，仅提供建议。

---

# 2. Workflow

```
Input BOM
      │
      ▼
Read BOM
      │
      ▼
Column Mapping
      │
      ▼
Component Detection
      │
      ▼
Value Extraction
      │
      ▼
Normalization
      │
      ▼
Rule Engine
      │
      ▼
Similarity Analysis
      │
      ▼
Excel Report
```

---

# 3. Output Excel Structure

Output Excel 为 RD 最终 Review 文件。

## Sheet：Raw_BOM

用途：

保存原始 BOM。

来源：

Input Excel。

是否允许修改：

❌ 不允许。

---

## Sheet：Normalized_BOM

用途：

显示 Normalize 后的数据。

包含：

- Normalized Value
- Package
- Voltage
- Tolerance
- Dielectric
- Component Type

是否允许修改：

❌ 不允许。

---

## Sheet：Merge_Recommendation

用途：

提供建议 Merge 的料件。

包含：

- Target PN
- Member PN
- Merge Reason
- Priority
- Risk Level

是否允许修改：

❌ 不允许。

---

## Sheet：Review_Required

用途：

需要 RD 判断的项目。

例如：

- Voltage 不同
- Package 不同
- Tolerance 不同
- Dielectric 不同
- Function Unknown

是否允许修改：

✅ 允许 RD 填写 Review 结果。

---

## Sheet：Statistics

用途：

统计 Merge 成果。

包含：

- Total Component
- Merge Candidate
- Review Count
- Reduction Count

是否允许修改：

❌ 不允许。

---

# 4. Column Definition

## Part Number

料号。

---

## Description

原始 Description。

---

## Component Type

目前支持：

- R
- MLCC
- CAP

未来：

- L
- FB
- D
- Q
- IC

---

## Original Value

原始规格。

例如：

```
1000PF
```

---

## Display Value

提供 RD 阅读。

例如：

```
1nF
```

---

## Normalized Value

Tool 内部比较规格。

例如：

```
1000pF
```

统一格式后进行比较。

---

## Package

封装。

例如：

```
0201
0402
0603
0805
```

---

## Voltage

耐压。

例如：

```
6.3V
10V
16V
25V
50V
```

---

## Tolerance

例如：

```
±1%
±5%
±10%
```

---

## Dielectric

例如：

```
X5R
X7R
C0G
NP0
```

---

# 5. Rule Engine

## Rule 001

统一单位。

例如：

```
1000PF

↓

1nF
```

---

## Rule 002

统一大小写。

例如：

```
UF

uF

μF

↓

uF
```

---

## Rule 003

统一科学记号。

例如：

```
1E-7

↓

100nF
```

---

## Rule 004

Package 不同

进入 Review。

---

## Rule 005

Voltage 不同

进入 Review。

---

## Rule 006

Tolerance 不同

进入 Review。

---

## Rule 007

Dielectric 不同

进入 Review。

---

## Rule 008

未知规格

进入 Review。

---

# 6. Merge Priority

Priority 1

数量少于 10PCS 的料件优先 Merge。

---

Priority 2

Value 相同。

---

Priority 3

Package 相同。

---

Priority 4

Voltage 相同。

---

Priority 5

Tolerance 相同。

---

Priority 6

Dielectric 相同。

---

Priority 7

Unknown Function

禁止自动 Merge。

必须 RD Review。

---

# 7. RD Workflow

Step 1

打开 Review Sheet。

↓

Step 2

依照 Priority 排序。

↓

Step 3

Review Merge Reason。

↓

Step 4

确认是否 Merge。

↓

Step 5

修改正式 BOM。

---

# 8. Naming Convention

Python

Function：

snake_case

Class：

PascalCase

Constant：

UPPER_CASE

Excel Sheet：

固定名称。

Excel Column：

固定名称。

禁止任意修改。

---

# 9. Editable Policy

| Sheet | Editable |
|----------|----------|
| Raw_BOM | ❌ |
| Normalized_BOM | ❌ |
| Merge_Recommendation | ❌ |
| Review_Required | ✅ |
| Statistics | ❌ |

---

# 10. Design Principles

本 Tool 的设计原则：

1.

协助 RD，而非取代 RD。

2.

所有 Merge 建议必须可追溯。

3.

所有 Rule 必须透明。

4.

所有 Recommendation 必须有原因。

5.

宁可多 Review，不可错误 Merge。

6.

Output Excel 必须稳定，不随版本改变格式。

7.

Backward Compatibility 为最高优先级。

---

# 11. Future Roadmap

V4

完成 BOM Normalize。

---

V5

Resistor Function Analysis。

---

V6

AI Merge Recommendation。

---

V7

PCB Placement Analysis。

---

V8

BOM Cost Optimization。

---

V9

EFA Knowledge Integration。

---

V10

AI Hardware Design Assistant。