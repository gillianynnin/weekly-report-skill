import pandas as pd
import numpy as np
import json

# ── 读取原始数据与暂存校验数据 ───────────────────────────────
excel_path = r"C:\Users\gillian.yin\data_workspace\03_aiagent\_data_source_path.txt"
try:
    with open(excel_path, 'r', encoding='utf-8') as f:
        file_path = f.read().strip()
except FileNotFoundError:
    file_path = r"C:\Users\gillian.yin\Desktop\agent_test_V0402.xlsx"

staging_json = r"C:\Users\gillian.yin\data_workspace\03_aiagent\staging_data.json"
with open(staging_json, 'r', encoding='utf-8') as f:
    staging = json.load(f)

df = pd.read_excel(file_path)
df['week_num'] = df['week_number'].str.extract(r'(\d+)').astype(int)

latest_week = staging['latest_week']
prev_week = staging['prev_week']
budget_dict = staging['budget_dict']
pac_total_budget = staging['pac_total_budget']

errors = []
warnings = []
info = []

TOLERANCE = 0.01  # 允许误差：1分钱以内视为一致（浮点精度）

# ══════════════════════════════════════════════════════════════
# 一、数据完整性校验
# ══════════════════════════════════════════════════════════════
print("=" * 60)
print("【一】数据完整性校验")
print("=" * 60)

# 1-1 必要字段检查
required_cols = ['BU', 'pmtu', 'pa_name', 'week_number', 'week_range', 'total_margin']
missing_cols = [c for c in required_cols if c not in df.columns]
if missing_cols:
    errors.append(f"缺少必要字段：{missing_cols}")
    print(f"  ❌ 缺少字段：{missing_cols}")
else:
    print(f"  ✅ 必要字段完整")

# 1-2 空值检查
null_counts = df[required_cols].isnull().sum()
null_cols = null_counts[null_counts > 0]
if not null_cols.empty:
    for col, cnt in null_cols.items():
        warnings.append(f"字段 [{col}] 存在 {cnt} 个空值")
        print(f"  ⚠️  字段 [{col}] 有 {cnt} 个空值")
else:
    print(f"  ✅ 无空值")

# 1-3 周次连续性检查
all_weeks = sorted(df['week_num'].unique())
gaps = [w for i, w in enumerate(all_weeks[1:]) if w - all_weeks[i] > 1]
if gaps:
    for g in gaps:
        warnings.append(f"周次存在断层：W{g-1:02d} → W{g:02d} 之间缺失数据")
        print(f"  ⚠️  周次断层：W{g-1:02d} → W{g:02d}")
else:
    print(f"  ✅ 周次连续（W{all_weeks[0]:02d} ~ W{all_weeks[-1]:02d}，共 {len(all_weeks)} 周）")

# 1-4 各 pmtu 数据点数检查
for rec in staging['records']:
    pmtu = rec['pmtu']
    wc = rec['week_count']
    if wc < 2:
        warnings.append(f"pmtu [{pmtu}] 历史数据仅 {wc} 周，计算结果参考价值低")
        print(f"  ⚠️  [{pmtu}] 仅 {wc} 周数据")

if not any(rec['week_count'] < 2 for rec in staging['records']):
    print(f"  ✅ 各 pmtu 数据点充足")

# ══════════════════════════════════════════════════════════════
# 二、计算准确性校验（独立重算，逐项比对）
# ══════════════════════════════════════════════════════════════
print()
print("=" * 60)
print("【二】计算准确性校验（独立重算）")
print("=" * 60)

# 独立汇总：按 BU+pmtu+week_num 求和
raw_grouped = df.groupby(['BU', 'pmtu', 'week_num'])['total_margin'].sum()

calc_errors = 0
for rec in staging['records']:
    pmtu = rec['pmtu']
    bu = rec['BU']
    label = f"[{bu} / {pmtu}]"

    # 2-1 最新周 margin
    try:
        expected_cur = raw_grouped.loc[(bu, pmtu, latest_week)]
    except KeyError:
        expected_cur = 0
    diff = abs(rec['current_margin'] - expected_cur)
    if diff > TOLERANCE:
        errors.append(f"{label} 最新周 margin 不符：报表={rec['current_margin']:,.2f}，重算={expected_cur:,.2f}，差异={diff:,.2f}")
        print(f"  ❌ {label} 最新周 margin 差异 {diff:,.2f}")
        calc_errors += 1

    # 2-2 上周 margin
    try:
        expected_prev = raw_grouped.loc[(bu, pmtu, prev_week)]
    except KeyError:
        expected_prev = 0
    diff = abs(rec['prev_margin'] - expected_prev)
    if diff > TOLERANCE:
        errors.append(f"{label} 上周 margin 不符：报表={rec['prev_margin']:,.2f}，重算={expected_prev:,.2f}，差异={diff:,.2f}")
        print(f"  ❌ {label} 上周 margin 差异 {diff:,.2f}")
        calc_errors += 1

    # 2-3 周比
    if expected_prev != 0:
        expected_wow = (expected_cur - expected_prev) / abs(expected_prev) * 100
        if rec['wow_pct'] is not None:
            diff_wow = abs(rec['wow_pct'] - expected_wow)
            if diff_wow > 0.05:  # 允许 0.05 个百分点误差
                errors.append(f"{label} 周比不符：报表={rec['wow_pct']:+.2f}%，重算={expected_wow:+.2f}%")
                print(f"  ❌ {label} 周比差异 {diff_wow:.3f}pp")
                calc_errors += 1

    # 2-4 YTD 累计
    pmtu_weeks = raw_grouped.xs((bu, pmtu), level=['BU', 'pmtu']) if (bu, pmtu) in raw_grouped.index.droplevel(2) else pd.Series()
    expected_ytd = pmtu_weeks.sum() if len(pmtu_weeks) > 0 else 0
    diff_ytd = abs(rec['cumulative'] - expected_ytd)
    if diff_ytd > TOLERANCE:
        errors.append(f"{label} YTD 累计不符：报表={rec['cumulative']:,.2f}，重算={expected_ytd:,.2f}，差异={diff_ytd:,.2f}")
        print(f"  ❌ {label} YTD 累计差异 {diff_ytd:,.2f}")
        calc_errors += 1

    # 2-5 进度百分比
    if not pmtu.startswith('PAC-'):
        budget = budget_dict.get(pmtu, 0)
        if budget > 0:
            expected_pct = expected_ytd / budget * 100
            if rec['progress_pct'] is not None:
                diff_pct = abs(rec['progress_pct'] - expected_pct)
                if diff_pct > 0.05:
                    errors.append(f"{label} 进度% 不符：报表={rec['progress_pct']:.2f}%，重算={expected_pct:.2f}%")
                    print(f"  ❌ {label} 进度% 差异 {diff_pct:.3f}pp")
                    calc_errors += 1
    else:
        # PAC 整体进度
        expected_pac_ytd = df[df['pmtu'].str.startswith('PAC-', na=False)]['total_margin'].sum()
        expected_pac_pct = expected_pac_ytd / pac_total_budget * 100
        diff_pac = abs(rec['progress_pct'] - expected_pac_pct)
        if diff_pac > 0.05:
            errors.append(f"{label} PAC 整体进度% 不符：报表={rec['progress_pct']:.2f}%，重算={expected_pac_pct:.2f}%")
            print(f"  ❌ {label} PAC 进度% 差异 {diff_pac:.3f}pp")
            calc_errors += 1

# 2-6 全局 YTD
expected_total_ytd = df['total_margin'].sum()
diff_total = abs(staging['ytd_margin'] - expected_total_ytd)
if diff_total > TOLERANCE:
    errors.append(f"全局 YTD margin 不符：报表={staging['ytd_margin']:,.2f}，重算={expected_total_ytd:,.2f}")
    print(f"  ❌ 全局 YTD 差异 {diff_total:,.2f}")
    calc_errors += 1

if calc_errors == 0:
    print(f"  ✅ 所有计算项校验通过（共 {len(staging['records'])} 个 pmtu，{len(staging['records'])*5+1} 项检查）")

# ══════════════════════════════════════════════════════════════
# 三、预测模型合理性校验
# ══════════════════════════════════════════════════════════════
print()
print("=" * 60)
print("【三】预测模型选用合理性")
print("=" * 60)
print(f"  {'pmtu':<30} {'周数':>4} {'CV':>6} {'趋势%/周':>8} {'报表算法':<12} 校验结论")
print(f"  {'-'*80}")

model_issues = 0
for rec in staging['records']:
    pmtu = rec['pmtu']
    margins = np.array(rec['weekly_margins'])
    n = len(margins)
    algo_reported = rec['forecast_algo']

    if pmtu.startswith('PAC-'):
        verdict = "✅ PAC整体均值（正常）"
    elif n < 3:
        verdict = "ℹ️  数据不足3周，用均值（正常）"
    else:
        weeks = np.arange(1, n + 1)
        mean = np.abs(margins.mean())
        cv = margins.std() / mean if mean > 0 else 0
        slope, _ = np.polyfit(weeks, margins, 1)
        trend_pct = abs(slope) / mean * 100 if mean > 0 else 0

        if cv > 0.5 and trend_pct > 5:
            expected_algo = '指数平滑'
        elif trend_pct > 5:
            expected_algo = '线性回归'
        else:
            expected_algo = '加权移动均值'

        if algo_reported == expected_algo:
            verdict = f"✅ {algo_reported}（CV={cv:.2f}, 趋势={trend_pct:.1f}%）"
        else:
            verdict = f"❌ 应为{expected_algo}，实为{algo_reported}（CV={cv:.2f}, 趋势={trend_pct:.1f}%）"
            errors.append(f"[{pmtu}] 算法选用有误：应为 {expected_algo}，报表用了 {algo_reported}")
            model_issues += 1

        print(f"  {pmtu:<30} {n:>4} {cv:>6.2f} {trend_pct:>7.1f}% {algo_reported:<12} {verdict}")
        continue

    print(f"  {pmtu:<30} {n:>4} {'—':>6} {'—':>8} {algo_reported:<12} {verdict}")

if model_issues == 0:
    print(f"\n  ✅ 所有 pmtu 算法选用正确")

# ══════════════════════════════════════════════════════════════
# 汇总结论
# ══════════════════════════════════════════════════════════════
print()
print("=" * 60)
print("【校验汇总】")
print("=" * 60)

if errors:
    print(f"\n  ❌ 发现 {len(errors)} 个错误，请修正后再确认输出：")
    for i, e in enumerate(errors, 1):
        print(f"     {i}. {e}")
else:
    print(f"\n  ✅ 计算准确性：通过")

if warnings:
    print(f"\n  ⚠️  {len(warnings)} 个提示（不影响输出，建议关注）：")
    for w in warnings:
        print(f"     - {w}")

print()
if not errors:
    print("  数据校验全部通过。请运行 /finalize-report 正式输出报表。")
else:
    print("  存在计算错误，请检查原始数据或报表脚本后重新运行 /weekly-report。")
