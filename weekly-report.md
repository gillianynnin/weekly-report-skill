---
name: weekly-report
description: 生成周度利润分析报表（暂存版，需校验后正式输出）
---

<skill>
当用户输入 `/weekly-report` 时，执行以下步骤：

## 步骤一：处理数据源路径参数

检查用户是否在 `/weekly-report` 后携带了参数（Excel 文件路径）。

- **无参数**：使用脚本中的默认路径，直接跳到步骤二。
- **有参数**（如 `/weekly-report C:\path\to\new_file.xlsx`）：
  1. 将用户提供的路径写入临时文件 `C:\Users\gillian.yin\data_workspace\03_aiagent\_data_source_path.txt`
  2. 读取 `generate_html_report.py`，找到 `file_path = ...` 那一行
  3. 替换为新路径：`file_path = r"用户提供的路径"`
  4. 保存文件后继续执行步骤二

## 步骤二：运行报表生成脚本

执行以下命令：

```bash
PYTHONIOENCODING=utf-8 python "C:\Users\gillian.yin\data_workspace\03_aiagent\generate_html_report.py"
```

脚本将输出：
- `weekly_report_staging.html` — 暂存报表（未正式输出）
- `staging_data.json` — 校验数据

## 步骤三：提示用户运行校验

向用户反馈：

> 暂存报表已生成，共 {行数} 行数据。
> 
> ⚠️  请运行 `/validate-report` 进行数据校验，确认无误后再正式输出。

## 注意事项

- 脚本路径固定为：`C:\Users\gillian.yin\data_workspace\03_aiagent\generate_html_report.py`
- 若用户传入相对路径，需转换为完整绝对路径
- 更新 `file_path` 时使用原始字符串格式 `r"..."`
- 预估完成率采用智能算法自动选择：根据各 pmtu 历史数据的波动性（CV）和趋势强度，自动选用 Holt 指数平滑、线性回归、或加权移动均值三种算法之一；PAC 系列统一使用整体均值法
- 此版本输出暂存文件，必须经过 `/validate-report` 校验并用户确认后，才能通过 `/finalize-report` 正式输出
</skill>
