import shutil
import json
import pandas as pd

staging_html = r"C:\Users\YOUR_USERNAME\YOUR_WORKSPACE\weekly_report_staging.html"
final_html   = r"C:\Users\YOUR_USERNAME\YOUR_WORKSPACE\weekly_report.html"
staging_json = r"C:\Users\YOUR_USERNAME\YOUR_WORKSPACE\staging_data.json"

# ── 1. 将暂存 HTML 重命名为正式文件 ───────────────────────────
shutil.copy2(staging_html, final_html)
print(f"✅ HTML 报表已输出：{final_html}")

# ── 2. 从原始数据 + staging JSON 导出 Excel（保留 BU 分组）──
excel_path = r"C:\Users\YOUR_USERNAME\YOUR_WORKSPACE\_data_source_path.txt"
try:
    with open(excel_path, 'r', encoding='utf-8') as f:
        file_path = f.read().strip()
except FileNotFoundError:
    file_path = r"C:\Users\YOUR_USERNAME\Desktop\agent_test_V0402.xlsx"

with open(staging_json, 'r', encoding='utf-8') as f:
    staging = json.load(f)

df = pd.read_excel(file_path)
df['week_num'] = df['week_number'].str.extract(r'(\d+)').astype(int)

latest_week = staging['latest_week']
prev_week   = staging['prev_week']
budget_dict = staging['budget_dict']
pac_total_budget = staging['pac_total_budget']
pac_cumulative   = staging['pac_cumulative']

latest_range = df[df['week_num'] == latest_week]['week_range'].iloc[0]
margin_col       = f'W{latest_week:02d} margin（{latest_range}）'
week_compare_col = f'周比（W{latest_week:02d} vs W{prev_week:02d}）'

rows = []
for rec in staging['records']:
    bu   = rec['BU']
    pmtu = rec['pmtu']
    cur  = rec['current_margin']
    prv  = rec['prev_margin']
    ytd  = rec['cumulative']
    pct  = rec['progress_pct']
    proj = rec['projected_quarter']
    algo = rec['forecast_algo']

    if prv != 0:
        wow_str = f"{((cur - prv) / abs(prv) * 100):+.1f}%（W{latest_week:02d}: {cur:,.0f} vs W{prev_week:02d}: {prv:,.0f}）"
    else:
        wow_str = f"N/A（W{latest_week:02d}: {cur:,.0f} vs W{prev_week:02d}: 0）"

    if pmtu.startswith('PAC-'):
        progress_str = f"{pct:.1f}%（PAC整体: {pac_cumulative:,.0f}/{pac_total_budget:,.0f}）"
    elif budget_dict.get(pmtu, 0) > 0:
        budget = budget_dict[pmtu]
        progress_str = f"{pct:.1f}%（{ytd:,.0f}/{budget:,.0f}）"
    else:
        progress_str = f"未设置预算（累计: {ytd:,.0f}）"

    budget_val = pac_total_budget if pmtu.startswith('PAC-') else budget_dict.get(pmtu, 0)
    proj_rate  = f"{proj / budget_val * 100:.1f}%" if budget_val > 0 else "未设预算"

    rows.append({
        'BU': bu,
        'pmtu': pmtu,
        margin_col: f"{cur:,.0f}",
        week_compare_col: wow_str,
        '进度': progress_str,
        '预估季度完成（算法）': f"{proj:,.0f}（{algo}）",
        '预估完成率': proj_rate,
    })

excel_df = pd.DataFrame(rows)

output_excel = r"C:\Users\YOUR_USERNAME\YOUR_WORKSPACE\weekly_report.xlsx"
with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
    # 整体汇总 sheet
    excel_df.to_excel(writer, index=False, sheet_name='总览')

    # 每个 BU 独立 sheet
    for bu, group in excel_df.groupby('BU'):
        safe_name = str(bu)[:31]  # Excel sheet 名最长 31 字符
        group.reset_index(drop=True).to_excel(writer, index=False, sheet_name=safe_name)

print(f"✅ Excel 报表已输出：{output_excel}（含总览 + {excel_df['BU'].nunique()} 个 BU 分 sheet）")
