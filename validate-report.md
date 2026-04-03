---
name: validate-report
description: 校验报表数据准确性、完整性、预测模型合理性
---

<skill>
当用户输入 `/validate-report` 时，执行以下步骤：

## 步骤一：运行校验脚本

执行以下命令：

```bash
python "C:\Users\YOUR_USERNAME\YOUR_WORKSPACE\validate_report.py"
```

## 步骤二：展示校验结果

校验脚本会输出三部分内容：

1. **数据完整性校验**
   - 必要字段检查
   - 空值检查
   - 周次连续性检查
   - 各 pmtu 数据点数检查

2. **计算准确性校验**（独立重算）
   - 最新周 margin
   - 上周 margin
   - 周比百分比
   - YTD 累计
   - 进度百分比
   - 全局 YTD

3. **预测模型合理性**
   - 列出每个 pmtu 使用的算法
   - 验证算法选择是否符合 CV 和趋势规则

## 步骤三：根据校验结果引导用户

- **若存在错误（❌）**：
  > 发现 {N} 个计算错误，请检查原始数据或报表脚本后重新运行 `/weekly-report`。

- **若仅有警告（⚠️）**：
  > 数据校验通过，但有 {N} 个提示需要关注（不影响输出）。
  > 
  > 确认无误后，请运行 `/finalize-report` 正式输出报表。

- **若全部通过（✅）**：
  > 数据校验全部通过。请运行 `/finalize-report` 正式输出报表。

## 注意事项

- 校验脚本路径：`C:\Users\YOUR_USERNAME\YOUR_WORKSPACE\validate_report.py`
- 校验依赖 `staging_data.json` 和原始 Excel 文件
- 计算准确性采用独立重算方式，允许浮点误差 ±0.01
- 预测模型校验基于 CV 和趋势百分比规则
</skill>
