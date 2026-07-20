# 规则库指南

## 加载模型

规则文件必须是 UTF-8 JSON 对象，`schema_version` 当前为 `1`。文件内容会深度覆盖
内置默认规则，未填写的字段继续使用默认值。报告中的 `Rule Library` 保存实际生效
规则快照，`Report Metadata` 保存规则文件路径。

启动时会拒绝以下配置：

- 电阻或电容相近值比例不大于 `1`；
- 置信度不在 `0` 到 `1`；
- 同一厂商同时出现在批准和禁用列表；
- 自定义规则 ID 缺失或重复；
- 未支持的运算符或严重度；
- 应为数组、对象或数字的字段类型错误。

## 顶层配置

| 路径 | 含义 |
| --- | --- |
| `analysis.near_value_ratio.R` | 电阻高值/低值最大比例 |
| `analysis.near_value_ratio.C` | 电容高值/低值最大比例 |
| `risk.lifecycle_keywords` | 生命周期风险关键词 |
| `risk.single_source_vendor_limit` | 单一来源供应商上限 |
| `risk.high_fragmentation_pn_count` | 同规格高碎片料号阈值 |
| `avl.minimum_vendors` | 统一 AVL 目标供应商数 |
| `avl.approved_vendors` | 企业批准厂商；空数组表示不限制 |
| `avl.blocked_vendors` | 禁用厂商 |
| `cost_down.minimum_impact_quantity` | Cost Down 最小影响数量 |
| `confidence.*` | 内置规则置信度 |

## 自定义规则

`custom_rules` 可在不改 Python 的情况下检查任意标准字段或保留的源 BOM 字段。

支持运算符：

| 运算符 | 命中条件 |
| --- | --- |
| `contains_any` | 字段包含任一配置文本，不区分大小写 |
| `equals_any` | 字段等于任一配置文本，不区分大小写 |
| `not_in` | 非空字段不在允许列表 |
| `missing` | 字段为空，`values` 可留空 |

支持严重度为 `Critical`、`High`、`Medium`、`Low`。

```json
{
  "schema_version": 1,
  "avl": {
    "minimum_vendors": 2,
    "approved_vendors": ["Murata", "Samsung", "TDK"],
    "blocked_vendors": []
  },
  "custom_rules": [
    {
      "id": "ORG-001",
      "enabled": true,
      "field": "Lifecycle_Status",
      "operator": "equals_any",
      "values": ["Conditional", "Restricted"],
      "component_types": [],
      "severity": "High",
      "confidence": 0.95,
      "finding": "Component is restricted by lifecycle policy",
      "recommendation": "Obtain sourcing approval or qualify a replacement."
    }
  ]
}
```

`component_types` 为空表示适用于所有器件；设置为 `["R", "C"]` 可限制为电阻和
电容。命中结果会写入 `AI Rule Findings`，包含规则 ID、证据、置信度和建议。

## 变更管理建议

企业部署时应将规则文件与代码一同版本控制。每次规则变更至少记录变更原因、责任人、
验证 BOM 和生效日期；先在代表性项目上比较新旧报告，再发布到生产目录。规则命中是
辅助决策，不替代设计、元件、采购或质量部门的批准流程。