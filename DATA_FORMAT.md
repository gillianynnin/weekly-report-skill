# 示例数据文件说明

本工具需要 Excel 文件包含以下列（列名必须完全匹配）：

| 列名 | 说明 | 示例 |
|------|------|------|
| BU | 业务单元 | 心MTC、泛PAC、棱洋动态 |
| pmtu | PMTU分类 | MTC-Meta、PAC-海外管理型 |
| pa_name | 客户名称 | 茄子快传、点点互动 |
| week_number | 周数 | W01、W02、W14 |
| week_range | 周日期区间 | 2026-01-01 ~ 2026-01-07 |
| total_margin | 利润金额 | 12345.67 |

## 数据示例

```
BU,pmtu,pa_name,week_number,week_range,total_margin
心MTC,MTC-Meta,茄子快传,W01,2026-01-01 ~ 2026-01-07,50000
心MTC,MTC-Meta,点点互动,W01,2026-01-01 ~ 2026-01-07,30000
泛PAC,PAC-海外管理型,滴滴,W01,2026-01-01 ~ 2026-01-07,25000
```

## 注意事项

1. **列名大小写敏感**：必须完全匹配上述列名
2. **week_number 格式**：必须是 "W" + 两位数字（如 W01、W14）
3. **week_range 格式**：必须是 "YYYY-MM-DD ~ YYYY-MM-DD"
4. **total_margin**：可以是负数，表示亏损
5. **编码**：建议使用 UTF-8 编码保存 Excel 文件
