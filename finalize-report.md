---
name: finalize-report
description: 用户确认校验通过后，正式输出 HTML + Excel 报表
---

<skill>
当用户输入 `/finalize-report` 时，执行以下步骤：

## 步骤一：确认用户已完成校验

询问用户：

> 请确认你已运行 `/validate-report` 并查看了校验结果。是否继续正式输出报表？

- 若用户回答"是"、"确认"、"继续"等，执行步骤二
- 若用户回答"否"或"取消"，提示：请先运行 `/validate-report` 完成校验

## 步骤二：运行最终输出脚本

执行以下命令：

```bash
python "C:\Users\YOUR_USERNAME\YOUR_WORKSPACE\finalize_report.py"
```

## 步骤三：告知用户结果

脚本成功后向用户反馈：

> ✅ 报表已正式输出！
> 
> - HTML 报表：`C:\Users\YOUR_USERNAME\YOUR_WORKSPACE\weekly_report.html`
> - Excel 报表：`C:\Users\YOUR_USERNAME\YOUR_WORKSPACE\weekly_report.xlsx`（含总览 + 各 BU 分 sheet）
> 
> 你可以直接用浏览器打开 HTML 文件，或用 Excel 打开 xlsx 文件查看。

## 注意事项

- 脚本路径：`C:\Users\YOUR_USERNAME\YOUR_WORKSPACE\finalize_report.py`
- 此命令会将 `weekly_report_staging.html` 重命名为 `weekly_report.html`
- Excel 输出包含"总览" sheet + 每个 BU 独立 sheet，保留分组结构
- 若校验未通过就运行此命令，输出的报表可能包含错误数据
</skill>
