import pandas as pd
import json
import numpy as np

# 读取Excel文件
file_path = r"C:\Users\YOUR_USERNAME\Desktop\agent_test_V0402.xlsx"
df = pd.read_excel(file_path)

# 读取预算配置
budget_file = r"C:\Users\YOUR_USERNAME\YOUR_WORKSPACE\pmtu_budget.json"
try:
    with open(budget_file, 'r', encoding='utf-8') as f:
        budget_dict = json.load(f)
except FileNotFoundError:
    budget_dict = {}
    print("警告：未找到预算配置文件，将使用默认值")

# PAC 整体预算（季度目标）
pac_total_budget = 0  # TODO: 填入 PAC 季度目标预算金额

# 提取周数
df['week_num'] = df['week_number'].str.extract(r'(\d+)').astype(int)

# 按BU、pmtu、周数分组汇总
grouped = df.groupby(['BU', 'pmtu', 'week_num', 'week_range'])['total_margin'].sum().reset_index()

# 获取所有周数并排序，以及全局统一的week_range映射
all_weeks = sorted(grouped['week_num'].unique())
week_range_map = grouped.drop_duplicates('week_num').set_index('week_num')['week_range'].to_dict()

# 计算 PAC 整体累计
pac_cumulative = grouped[grouped['pmtu'].str.startswith('PAC-', na=False)]['total_margin'].sum()

latest_week = max(all_weeks)
prev_week = latest_week - 1
latest_range = week_range_map.get(latest_week, '')
margin_col = f'W{latest_week:02d} margin（{latest_range}）'
week_compare_col = f'周比（W{latest_week:02d} vs W{prev_week:02d}）'

# 构建报表
report_data = []

for (bu, pmtu), group in grouped.groupby(['BU', 'pmtu']):
    row = {'BU': bu, 'pmtu': pmtu}

    margins = {}
    for week in all_weeks:
        week_data = group[group['week_num'] == week]
        margins[week] = week_data['total_margin'].values[0] if not week_data.empty else 0

    # 最新周margin
    current_margin = margins[latest_week]
    row[margin_col] = f"{current_margin:,.0f}"

    # 周比：百分比 + 两周数据
    prev_margin = margins.get(prev_week, 0)
    if prev_margin != 0:
        week_change = ((current_margin - prev_margin) / prev_margin) * 100
        row[week_compare_col] = f"{week_change:+.1f}%（W{latest_week:02d}: {current_margin:,.0f} vs W{prev_week:02d}: {prev_margin:,.0f}）"
    else:
        row[week_compare_col] = f"N/A（W{latest_week:02d}: {current_margin:,.0f} vs W{prev_week:02d}: 0）"

    # 累计进度
    cumulative = sum(margins.values())

    if pmtu.startswith('PAC-'):
        progress_pct = (pac_cumulative / pac_total_budget) * 100
        row['进度'] = f"{progress_pct:.1f}%（PAC整体: {pac_cumulative:,.0f}/{pac_total_budget:,.0f}）"
    else:
        target_value = budget_dict.get(pmtu, 0)
        if target_value > 0:
            progress_pct = (cumulative / target_value) * 100
            row['进度'] = f"{progress_pct:.1f}%（{cumulative:,.0f}/{target_value:,.0f}）"
        else:
            progress_pct = None
            row['进度'] = f"未设置预算（累计: {cumulative:,.0f}）"

    row['_progress_pct'] = progress_pct
    row['_current_margin'] = current_margin
    report_data.append(row)

# 转换为DataFrame，按BU分组后按margin降序排列
report_df_raw = pd.DataFrame(report_data)
report_df_raw = report_df_raw.sort_values(['BU', '_current_margin'], ascending=[True, False])
report_df = report_df_raw[['BU', 'pmtu', margin_col, week_compare_col, '进度']].reset_index(drop=True)

# 进度颜色预警
def color_progress(pct):
    if pct is None:
        return ''
    elif pct < 50:
        return 'background-color: #fde8e8; color: #c0392b;'
    elif pct < 80:
        return 'background-color: #fef9e7; color: #b7770d;'
    else:
        return 'background-color: #e9f7ef; color: #1e8449;'

def build_colored_table(df, progress_pcts):
    headers = ''.join([f'<th>{col}</th>' for col in df.columns])
    rows_html = ''
    for i, (_, row) in enumerate(df.iterrows()):
        cells = ''
        for col in df.columns:
            if col == '进度':
                style = color_progress(progress_pcts[i])
                cells += f'<td style="{style}">{row[col]}</td>'
            else:
                cells += f'<td>{row[col]}</td>'
        rows_html += f'<tr>{cells}</tr>'
    return f'<table><thead><tr>{headers}</tr></thead><tbody>{rows_html}</tbody></table>'

progress_pcts = report_df_raw['_progress_pct'].tolist()
html_output = build_colored_table(report_df, progress_pcts)

# 计算YTD数据
ytd_margin = df['total_margin'].sum()
latest_date = df['week_range'].apply(lambda x: x.split('~')[1].strip()).max()

# ── 智能预估函数 ──────────────────────────────────────────────
def smart_forecast(weekly_margins, total_weeks=13):
    margins = np.array(weekly_margins)
    n = len(margins)
    if n < 3:
        return margins.mean() * total_weeks, "均值"

    weeks = np.arange(1, n + 1)
    mean = np.abs(margins.mean())
    cv = margins.std() / mean if mean > 0 else 0
    slope, intercept = np.polyfit(weeks, margins, 1)
    trend_pct = abs(slope) / mean * 100 if mean > 0 else 0
    weeks_remaining = max(total_weeks - n, 0)

    if cv > 0.5 and trend_pct > 5:
        alpha, beta = 0.4, 0.3
        level, trend_val = margins[0], margins[1] - margins[0]
        for m in margins[1:]:
            prev_level = level
            level = alpha * m + (1 - alpha) * (level + trend_val)
            trend_val = beta * (level - prev_level) + (1 - beta) * trend_val
        projected = margins.sum() + sum([level + trend_val * i for i in range(1, weeks_remaining + 1)])
        algo = "指数平滑"
    elif trend_pct > 5:
        future_weeks = np.arange(n + 1, total_weeks + 1)
        future_vals = np.maximum(slope * future_weeks + intercept, 0)
        projected = margins.sum() + future_vals.sum()
        algo = "线性回归"
    else:
        weights = np.array([0.1, 0.15, 0.25, 0.5]) if n >= 4 else np.ones(n) / n
        weights = weights[-n:]
        weights = weights / weights.sum()
        weighted_avg = np.dot(margins[-len(weights):], weights)
        projected = margins.sum() + weighted_avg * weeks_remaining
        algo = "加权移动均值"

    return projected, algo

# ── AI 视角分析 ──────────────────────────────────────────────
total_weeks_in_quarter = 13
pac_pmtu_weeks = df[df['pmtu'].str.startswith('PAC-', na=False)].groupby('week_num')['total_margin'].sum().reset_index()
pac_weeks_elapsed = len(pac_pmtu_weeks)
pac_avg_weekly = pac_pmtu_weeks['total_margin'].sum() / pac_weeks_elapsed if pac_weeks_elapsed > 0 else 0
pac_projected_quarter = pac_avg_weekly * total_weeks_in_quarter
pac_projected_rate = (pac_projected_quarter / pac_total_budget * 100) if pac_total_budget > 0 else 0

pa_grouped = df.groupby(['BU', 'pmtu', 'pa_name', 'week_num'])['total_margin'].sum().reset_index()
ai_rows = []

for (bu, pmtu), group in pa_grouped.groupby(['BU', 'pmtu']):
    latest = group[group['week_num'] == latest_week][['pa_name', 'total_margin']].rename(columns={'total_margin': 'cur'})
    prev   = group[group['week_num'] == prev_week][['pa_name', 'total_margin']].rename(columns={'total_margin': 'prv'})
    merged = latest.merge(prev, on='pa_name', how='left')
    merged['prv_missing'] = merged['prv'].isna()
    merged = merged.fillna(0)

    merged['delta'] = merged['cur'] - merged['prv']
    merged['pct'] = merged.apply(
        lambda r: r['delta'] / abs(r['prv']) * 100 if not r['prv_missing'] and r['prv'] != 0 else None,
        axis=1
    )

    abs_max = merged['delta'].abs().max()
    pct_max = merged['pct'].dropna().abs().max() if not merged['pct'].dropna().empty else 0

    def combined_score(r):
        abs_score = abs(r['delta']) / abs_max if abs_max > 0 else 0
        if r['pct'] is not None and not pd.isna(r['pct']):
            pct_score = abs(r['pct']) / pct_max if (pct_max and pct_max > 0) else 0
            return abs_score * 0.7 + pct_score * 0.3
        else:
            return abs_score * 0.7

    merged['score'] = merged.apply(combined_score, axis=1)

    top3_up   = merged[merged['delta'] > 100].nlargest(3, 'score')
    top3_down = merged[merged['delta'] < -100].nlargest(3, 'score')

    def fmt_pa(r, is_up):
        arrow = '<span style="color: #00c853; font-weight: bold; font-size: 16px;">↑</span>' if is_up else '<span style="color: #e74c3c; font-weight: bold; font-size: 16px;">↓</span>'
        pct_str = f"{r['pct']:+.1f}%" if r['pct'] is not None and not pd.isna(r['pct']) else "上周为0"
        return f"{r['pa_name']} {arrow}{abs(r['delta']):,.0f}（{pct_str}）"

    up_str   = "；".join([fmt_pa(r, True) for _, r in top3_up.iterrows()]) or "无"
    down_str = "；".join([fmt_pa(r, False) for _, r in top3_down.iterrows()]) or "无"
    pa_str = f"{up_str}<br>{down_str}"

    # 预估完成率（智能算法）
    if pmtu.startswith('PAC-'):
        forecast_str = f"按当前周均 {pac_avg_weekly:,.0f} 预估季度 {pac_projected_quarter:,.0f}，PAC整体预估完成率 {pac_projected_rate:.1f}%"
    else:
        pmtu_weekly = df[df['pmtu'] == pmtu].groupby('week_num')['total_margin'].sum().sort_index()
        budget = budget_dict.get(pmtu, 0)
        if len(pmtu_weekly) >= 3:
            projected, algo = smart_forecast(pmtu_weekly.tolist(), total_weeks_in_quarter)
            projected_rate = (projected / budget * 100) if budget > 0 else None
            forecast_str = (f"预估季度 {projected:,.0f}（{algo}），预估完成率 {projected_rate:.1f}%"
                            if projected_rate else f"预估季度 {projected:,.0f}（{algo}，未设预算）")
        else:
            avg = pmtu_weekly.mean()
            projected = avg * total_weeks_in_quarter
            projected_rate = (projected / budget * 100) if budget > 0 else None
            forecast_str = (f"预估季度 {projected:,.0f}（均值），预估完成率 {projected_rate:.1f}%"
                            if projected_rate else f"预估季度 {projected:,.0f}（均值，未设预算）")

    ai_rows.append({
        'BU': bu,
        'pmtu': pmtu,
        f'W{latest_week:02d} 前三波动 PA（vs W{prev_week:02d}）': pa_str,
        '预估完成率': forecast_str
    })

ai_df = pd.DataFrame(ai_rows)
sort_order = report_df_raw[['BU', 'pmtu', '_current_margin']].copy()
ai_df = ai_df.merge(sort_order, on=['BU', 'pmtu'], how='left')
ai_df = ai_df.sort_values(['BU', '_current_margin'], ascending=[True, False])
ai_df = ai_df.drop(columns=['_current_margin']).reset_index(drop=True)
ai_html = ai_df.to_html(index=False, border=1, justify='left', escape=False)

# ── 风险预警 ──────────────────────────────────────────────────
import re as _re

def extract_rate(forecast_str):
    """从预估完成率字符串中提取数值"""
    m = _re.search(r'预估完成率\s*([\d.]+)%', forecast_str)
    return float(m.group(1)) if m else None

weeks_elapsed = latest_week
weeks_remaining = total_weeks_in_quarter - weeks_elapsed

risk_rows = []
for row in ai_rows:
    pmtu = row['pmtu']
    rate = extract_rate(row['预估完成率'])
    if rate is None:
        continue
    # 获取该 pmtu 实际数据周数
    if pmtu.startswith('PAC-'):
        pmtu_elapsed = pac_weeks_elapsed
    else:
        pmtu_elapsed = len(df[df['pmtu'] == pmtu]['week_num'].unique())

    # 剩余周数紧迫时升级风险
    if rate < 50 or (rate < 80 and weeks_remaining <= 3):
        level = '🔴 高风险'
    elif rate < 80:
        level = '🟡 关注'
    else:
        continue

    urgency = f'<span style="color:#e74c3c;font-weight:600;">仅剩 {weeks_remaining} 周</span>' if weeks_remaining <= 3 else f'剩余 {weeks_remaining} 周'
    period_str = f'已过 W{pmtu_elapsed:02d} / 共{total_weeks_in_quarter}周，{urgency}'
    risk_rows.append({
        'BU': row['BU'], 'pmtu': pmtu,
        '预估完成率': rate,
        '进度周期': period_str,
        '风险等级': level
    })

if risk_rows:
    risk_rows_sorted = sorted(risk_rows, key=lambda x: x['预估完成率'])
    risk_items = ''.join([
        f'''<tr style="background:{'#fff5f5' if r['风险等级'].startswith('🔴') else '#fffbf0'};">
            <td style="padding:8px 14px;border-bottom:1px solid #eee;">{r['风险等级']}</td>
            <td style="padding:8px 14px;border-bottom:1px solid #eee;">{r['BU']}</td>
            <td style="padding:8px 14px;border-bottom:1px solid #eee;">{r['pmtu']}</td>
            <td style="padding:8px 14px;border-bottom:1px solid #eee;font-weight:600;color:{'#e74c3c' if r['风险等级'].startswith('🔴') else '#f39c12'};">{r['预估完成率']:.1f}%</td>
            <td style="padding:8px 14px;border-bottom:1px solid #eee;font-size:12px;color:#555;">{r['进度周期']}</td>
        </tr>'''
        for r in risk_rows_sorted
    ])
    risk_html = f'''
    <div style="background:#fff;border-radius:8px;padding:16px 20px;margin-bottom:20px;
                border-left:4px solid #e74c3c;box-shadow:0 2px 6px rgba(0,0,0,0.07);">
        <div style="font-size:14px;font-weight:700;color:#c0392b;margin-bottom:12px;">⚠️ 完成率风险预警（预估完成率 &lt; 80%）</div>
        <table style="width:100%;border-collapse:collapse;font-size:13px;">
            <thead>
                <tr style="background:#fdf2f2;">
                    <th style="padding:8px 14px;text-align:left;color:#555;font-weight:600;">风险等级</th>
                    <th style="padding:8px 14px;text-align:left;color:#555;font-weight:600;">BU</th>
                    <th style="padding:8px 14px;text-align:left;color:#555;font-weight:600;">pmtu</th>
                    <th style="padding:8px 14px;text-align:left;color:#555;font-weight:600;">预估完成率</th>
                    <th style="padding:8px 14px;text-align:left;color:#555;font-weight:600;">进度周期</th>
                </tr>
            </thead>
            <tbody>{risk_items}</tbody>
        </table>
        <div style="font-size:11px;color:#999;margin-top:10px;">🔴 高风险：完成率 &lt; 50%，或完成率 &lt; 80% 且剩余 ≤ 3 周　　🟡 关注：50% ≤ 完成率 &lt; 80%</div>
    </div>'''
else:
    risk_html = '<div style="font-size:13px;color:#27ae60;margin-bottom:16px;">✅ 所有 pmtu 预估完成率均 ≥ 80%，暂无风险项。</div>'

# ── 每个 BU 的进度条卡片 ─────────────────────────────────────
def progress_bar_html(pct, label=''):
    pct_clamped = min(max(pct, 0), 100)
    if pct < 50:
        color = '#e74c3c'
    elif pct < 80:
        color = '#f39c12'
    else:
        color = '#27ae60'
    return f'''
        <div style="margin-top:8px;">
            <div style="display:flex; justify-content:space-between; font-size:12px; color:#7f8c8d; margin-bottom:3px;">
                <span>{label}</span><span style="font-weight:600;color:{color};">{pct:.1f}%</span>
            </div>
            <div style="background:#e8ecef; border-radius:6px; height:10px; overflow:hidden;">
                <div style="width:{pct_clamped}%; background:{color}; height:100%; border-radius:6px; transition:width 0.4s;"></div>
            </div>
        </div>'''

bu_summary_html = '<div style="display:flex; flex-wrap:wrap; gap:16px; margin-bottom:32px;">'
for bu, bu_group in df.groupby('BU'):
    latest_margin = bu_group[bu_group['week_num'] == latest_week]['total_margin'].sum()
    prev_margin_bu = bu_group[bu_group['week_num'] == prev_week]['total_margin'].sum()
    ytd_bu = bu_group['total_margin'].sum()

    if prev_margin_bu != 0:
        wow = ((latest_margin - prev_margin_bu) / abs(prev_margin_bu)) * 100
        wow_color = '#27ae60' if wow >= 0 else '#e74c3c'
        wow_str = f'<span style="color:{wow_color};font-weight:600;">{"+" if wow>=0 else ""}{wow:.1f}%</span>'
    else:
        wow_str = '<span style="color:#95a5a6;">上周无数据</span>'

    top_pmtu = (bu_group[bu_group['week_num'] == latest_week]
                .groupby('pmtu')['total_margin'].sum()
                .idxmax() if not bu_group[bu_group['week_num'] == latest_week].empty else '—')

    # BU 级别预算 = 各 pmtu 预算加总（PAC 用 pac_total_budget）
    bu_pmtu_list = [p for p in bu_group['pmtu'].unique() if pd.notna(p)]
    bu_budget = 0
    for p in bu_pmtu_list:
        if str(p).startswith('PAC-'):
            bu_budget = pac_total_budget  # PAC 整体共用一个预算
            break
        else:
            bu_budget += budget_dict.get(p, 0)

    if bu_budget > 0:
        progress_pct = ytd_bu / bu_budget * 100
        bar = progress_bar_html(progress_pct, f'Q1累计 {ytd_bu:,.0f} / 目标 {bu_budget:,.0f}')
    else:
        progress_pct = None
        bar = f'<div style="font-size:12px;color:#95a5a6;margin-top:8px;">暂无BU目标数据</div>'

    bu_summary_html += f'''
    <div style="flex:1; min-width:220px; background:#fff; border-radius:10px; padding:16px 20px;
                box-shadow:0 2px 8px rgba(0,0,0,0.07); border-top:4px solid #2e86c1;">
        <div style="font-size:15px; font-weight:700; color:#1a1a1a; margin-bottom:10px;">{bu}</div>
        <div style="font-size:13px; color:#555; line-height:2;">
            <div>本周 margin：<strong>{latest_margin:,.0f}</strong> &nbsp; 环比 W{prev_week:02d}：{wow_str}</div>
            <div>主要贡献：<strong>{top_pmtu}</strong></div>
        </div>
        {bar}
    </div>'''

bu_summary_html += '</div>'

# ── 导出 staging_data.json 供校验师比对 ──────────────────────
import json as _json

class _NpEncoder(_json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, (np.integer,)):
            return int(obj)
        if isinstance(obj, (np.floating,)):
            return float(obj)
        if isinstance(obj, np.ndarray):
            return obj.tolist()
        return super().default(obj)

staging_data = {
    'latest_week': latest_week,
    'prev_week': prev_week,
    'ytd_margin': ytd_margin,
    'pac_total_budget': pac_total_budget,
    'pac_cumulative': pac_cumulative,
    'pac_projected_quarter': pac_projected_quarter,
    'pac_projected_rate': pac_projected_rate,
    'budget_dict': budget_dict,
    'records': []
}

for _, row in report_df_raw.iterrows():
    pmtu = row['pmtu']
    bu = row['BU']
    current_margin = row['_current_margin']
    progress_pct = row['_progress_pct']

    pmtu_weekly = df[df['pmtu'] == pmtu].groupby('week_num')['total_margin'].sum().sort_index()
    cumulative = pmtu_weekly.sum()
    prev_margin_val = pmtu_weekly.get(prev_week, 0)
    wow_pct = ((current_margin - prev_margin_val) / prev_margin_val * 100) if prev_margin_val != 0 else None

    if pmtu.startswith('PAC-'):
        algo = 'PAC整体均值'
        projected = pac_projected_quarter
    elif len(pmtu_weekly) >= 3:
        projected, algo = smart_forecast(pmtu_weekly.tolist(), total_weeks_in_quarter)
    else:
        projected = pmtu_weekly.mean() * total_weeks_in_quarter
        algo = '均值'

    budget = pac_total_budget if pmtu.startswith('PAC-') else budget_dict.get(pmtu, 0)
    staging_data['records'].append({
        'BU': bu,
        'pmtu': pmtu,
        'current_margin': current_margin,
        'prev_margin': prev_margin_val,
        'wow_pct': wow_pct,
        'cumulative': cumulative,
        'budget': budget,
        'progress_pct': progress_pct,
        'projected_quarter': projected,
        'forecast_algo': algo,
        'week_count': len(pmtu_weekly),
        'weekly_margins': pmtu_weekly.tolist(),
    })

staging_json = r"C:\Users\YOUR_USERNAME\YOUR_WORKSPACE\staging_data.json"
with open(staging_json, 'w', encoding='utf-8') as f:
    _json.dump(staging_data, f, ensure_ascii=False, indent=2, cls=_NpEncoder)

# 保存为暂存 HTML（校验通过后由 finalize_report.py 重命名为正式文件）
output_file = r"C:\Users\YOUR_USERNAME\YOUR_WORKSPACE\weekly_report_staging.html"
with open(output_file, 'w', encoding='utf-8') as f:
    f.write(f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>周度利润分析报表</title>
    <style>
        * {{ box-sizing: border-box; }}
        body {{ font-family: 'Microsoft YaHei', '微软雅黑', 'SimHei', '黑体', Arial, sans-serif; margin: 40px 50px; background: #f5f7fa; color: #2c3e50; line-height: 1.6; }}
        h1 {{ font-size: 24px; font-weight: 600; color: #1a1a1a; margin-bottom: 8px; letter-spacing: 0.5px; }}
        h2 {{ font-size: 18px; font-weight: 600; color: #34495e; margin-top: 48px; margin-bottom: 16px; border-left: 4px solid #2e86c1; padding-left: 12px; }}
        .report-meta {{ font-size: 12px; color: #999; margin-bottom: 20px; }}
        .model-note {{ font-size: 12px; color: #7f8c8d; background: #f4f6f7; border: 1px solid #e5e8ea; border-radius: 4px; padding: 8px 14px; margin-bottom: 16px; }}
        table {{ border-collapse: collapse; width: 100%; font-size: 13px; background: #fff; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 6px rgba(0,0,0,0.08); margin-bottom: 28px; }}
        thead th {{ background-color: #2e86c1; color: #fff; padding: 12px 14px; text-align: left; font-weight: 600; font-size: 13.5px; }}
        td {{ padding: 11px 14px; border-bottom: 1px solid #eaecee; }}
        tbody tr:hover {{ background-color: #f0f4f8; }}
        .ai-section thead th {{ background-color: #5d6d7e; }}
        .model-note {{ font-size: 12px; color: #7f8c8d; background: #f4f6f7; border: 1px solid #e5e8ea; border-radius: 4px; padding: 8px 14px; margin-bottom: 16px; }}
    </style>
</head>
<body>
    <h1>周度利润分析报表</h1>
    <p style="font-size:12px;color:#999;margin-bottom:20px;">报表更新至 {latest_date}</p>

    <h2>各 BU Q1 进度总览</h2>
    {bu_summary_html}

    <h2>利润分析明细</h2>
    {html_output}

    <h2>AI 视角分析</h2>
    <p class="model-note">📌 备注：预估完成率中所使用的预测模型（指数平滑 / 线性回归 / 加权移动均值）均由算法根据各 pmtu 历史数据的波动性与趋势特征自动选择，无需人工干预。PAC 系列统一采用整体均值法计算。</p>
    {risk_html}
    <div class="ai-section">
    {ai_html}
    </div>
</body>
</html>
""")

print(f"暂存报表已生成: {output_file}")
print(f"校验数据已导出: {staging_json}")
print(f"\n共 {len(report_df)} 行数据")
print("\n[!] 请运行 /validate-report 进行数据校验，确认无误后再正式输出报表。")
