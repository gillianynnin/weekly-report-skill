# 周度利润分析报表工具包

## 📦 包含文件

- `generate_html_report.py` - 核心报表生成脚本
- `pmtu_budget.json` - 预算配置文件
- `weekly-report.md` - Claude Code Skill 文件
- `README.md` - 本说明文档

## 🚀 快速开始

### 方式一：使用 Claude Code Skill（推荐）

1. **安装 Skill**
   ```bash
   # 将 weekly-report.md 复制到你的 Claude Code skills 目录
   # Windows: C:\Users\你的用户名\.claude\skills\
   # Mac/Linux: ~/.claude/skills/
   ```

2. **配置预算**
   - 编辑 `pmtu_budget.json`，设置各 pmtu 的季度预算
   - PAC 系列预算在脚本中统一设置（默认 10,000,000）

3. **使用命令**
   ```bash
   # 使用默认数据源
   /weekly-report
   
   # 指定数据源
   /weekly-report C:\路径\你的数据.xlsx
   ```

### 方式二：直接运行脚本

1. **修改数据源路径**
   打开 `generate_html_report.py`，修改第 5 行：
   ```python
   file_path = r"C:\你的路径\数据文件.xlsx"
   ```

2. **运行脚本**
   ```bash
   python generate_html_report.py
   ```

3. **查看报表**
   生成的 HTML 文件默认保存在脚本同目录下的 `weekly_report.html`

## 📊 数据格式要求

Excel 文件需包含以下列：
- `BU` - 业务单元
- `pmtu` - PMTU 分类
- `pa_name` - 客户名称
- `week_number` - 周数（如 "W01", "W02"）
- `week_range` - 周日期区间（如 "2026-01-01 ~ 2026-01-07"）
- `total_margin` - 利润金额

## ⚙️ 配置说明

### 预算配置（pmtu_budget.json）

```json
{
  "MTC-Meta": 10455326,
  "MTC-Tiktok": 4360000,
  "心VASC-CFC": 4360000,
  ...
}
```

- 所有金额为季度目标（13周）
- PAC 系列不在此配置，统一在脚本中设置年度目标

### PAC 预算修改

打开 `generate_html_report.py`，修改第 18 行：
```python
pac_total_budget = 10000000  # PAC 季度目标
```

## 📈 报表内容

### 主表格
- 按 BU 分组，每个 BU 内按当周 margin 降序排列
- 显示最新周 margin、周比、累计进度
- 进度颜色预警：<50% 红色，50-80% 黄色，≥80% 绿色

### AI 视角分析
- 每个 BU 的一句话总结
- 前三涨/跌 PA 明细（综合绝对值和周比排序）
- 预估季度完成率

## 🔧 依赖环境

```bash
pip install pandas openpyxl
```

## 📝 常见问题

**Q: 如何修改输出文件路径？**
A: 编辑脚本第 202 行的 `output_file` 变量

**Q: 报表样式可以调整吗？**
A: 可以，修改脚本中 210-220 行的 CSS 样式

**Q: 如何添加新的 pmtu 预算？**
A: 在 `pmtu_budget.json` 中添加新的键值对即可

**Q: 数据源切换后需要重启吗？**
A: 不需要，直接运行即可

## 📧 技术支持

如有问题，请联系报表开发团队。

---

**版本**: v1.0  
**更新日期**: 2026-04-02
