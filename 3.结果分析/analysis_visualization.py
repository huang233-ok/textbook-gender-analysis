"""
分析脚本：场景性别比 + 职业丰富度 + Top10职业集中度
参照 2026.1.29 log 分析框架，用人头计数和独立职业两套数据分别分析
- 命名标准化：合并"老师"→"教师"等常见变体
- 独立职业表已修正：按 (page, profession, gender) 去重，保留性别信息
输出：结果/4.分析图表/
"""
import sys, warnings
from pathlib import Path
import numpy as np
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
matplotlib.rcParams['font.family'] = 'sans-serif'
matplotlib.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'Arial Unicode MS']
matplotlib.rcParams['axes.unicode_minus'] = False
warnings.filterwarnings('ignore')

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

PROJECT_ROOT = Path(__file__).parent.parent.resolve()
CLASS_1_DIR  = PROJECT_ROOT / "结果" / "2.分类结果" / "一位码"
OUTPUT_DIR   = PROJECT_ROOT / "结果" / "4.分析图表"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

PUB_ORDER = ["人教版", "北师大版", "苏教版", "部编版"]
BLUE, RED = '#4472C4', '#C0504D'
NON_EMPLOYED = [0, 9, 10]     # 学生/儿童、家庭角色、其他非从业

# ── 命名标准化映射 ──
PROFESSION_MERGE = {
    '老师': '教师',
    '私塾老师': '教师',
    '体育老师': '教师',
    '幼儿园老师': '教师',
}


# ── 加载数据 ──
def load_classified(suffix: str) -> pd.DataFrame:
    frames = []
    for f in sorted(CLASS_1_DIR.glob(f"已分类_一位码_*_{suffix}.xlsx")):
        df = pd.read_excel(f)
        parts = f.stem.split("_")
        pub = next((p for p in parts if p in PUB_ORDER), "")
        df["版本"] = pub
        frames.append(df)
    out = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    # 统一 职业分类代码 为 int（有时读成 float 或 str）
    out['职业分类代码'] = pd.to_numeric(out['职业分类代码'], errors='coerce')
    return out

print("加载数据...")
df_head = load_classified("人头计数")
df_uniq = load_classified("独立职业")
# 命名标准化
df_head['profession'] = df_head['profession'].replace(PROFESSION_MERGE)
df_uniq['profession'] = df_uniq['profession'].replace(PROFESSION_MERGE)
print(f"  人头计数: {len(df_head)} 行, 独立职业: {len(df_uniq)} 行")


# ── 筛选条件 ──
def apply_cond(df: pd.DataFrame, cond: int) -> pd.DataFrame:
    """
    1 = 全部数据
    2 = 成人（排除 code=0 学生/儿童）
    3 = 工作场景（成人 & scenario='工作场所'）
    4 = 家庭场景（成人 & scenario='家庭'）
    """
    adult = df['职业分类代码'] != 0
    if cond == 1: return df
    if cond == 2: return df[adult]
    if cond == 3: return df[adult & (df['scenario'] == '工作场所')]
    if cond == 4: return df[adult & (df['scenario'] == '家庭')]


def gender_ratio(df_sub, pub):
    s = df_sub[df_sub['版本'] == pub]
    m, f = (s['gender']=='男').sum(), (s['gender']=='女').sum()
    return m / f if f > 0 else np.nan


# ════════════════════════════════════════════════════════
#  图1：成人分场景性别比散点图
# ════════════════════════════════════════════════════════
print("绘制图1：场景性别比散点图…")
COND_LABELS  = ['1. 全部数据', '2. 成人', '3. 工作场景', '4. 家庭场景']
COND_MARKERS = ['o', 'x', 's', '+']
COND_COLORS  = ['#1F497D', '#4BACC6', '#92D050', '#00B050']

fig, ax = plt.subplots(figsize=(9, 5.5))
for ci, cond in enumerate([1, 2, 3, 4]):
    df_c = apply_cond(df_head, cond)
    ratios = [gender_ratio(df_c, p) for p in PUB_ORDER]
    ms = 160 if COND_MARKERS[ci] == '+' else 90
    ax.scatter(PUB_ORDER, ratios, marker=COND_MARKERS[ci], s=ms,
               color=COND_COLORS[ci], label=COND_LABELS[ci], zorder=5)
    ax.plot(PUB_ORDER, ratios, color=COND_COLORS[ci], alpha=0.35, linewidth=0.9)

ax.axhline(1.0, color='red', linestyle='--', linewidth=1.5, alpha=0.85, label='平衡线 (1:1)')
ax.set_ylabel('性别比（男/女）', fontsize=12)
ax.set_title('不同版本的性别比（4条件）', fontsize=13, fontweight='bold')
ax.legend(loc='upper right', fontsize=9)
ax.set_ylim(bottom=0)
ax.grid(axis='y', alpha=0.25)
plt.tight_layout()
fig.savefig(OUTPUT_DIR / "fig1_场景性别比散点图.png", dpi=150, bbox_inches='tight')
plt.close()
print("  → fig1 已保存")


# ════════════════════════════════════════════════════════
#  图2：4条件下各版本男女人数柱状图
# ════════════════════════════════════════════════════════
print("绘制图2：4条件男女人数柱状图…")
COND_TITLES = ['条件=1  全部数据', '条件=2  成人', '条件=3  工作场景', '条件=4  家庭场景']
fig, axes = plt.subplots(2, 2, figsize=(13, 9))
x = np.arange(len(PUB_ORDER)); w = 0.36

for ci, (cond, title) in enumerate(zip([1,2,3,4], COND_TITLES)):
    ax = axes[ci//2][ci%2]
    df_c = apply_cond(df_head, cond)
    mc = [((df_c[df_c['版本']==p])['gender']=='男').sum() for p in PUB_ORDER]
    fc = [((df_c[df_c['版本']==p])['gender']=='女').sum() for p in PUB_ORDER]
    bm = ax.bar(x - w/2, mc, w, label='男', color=BLUE, alpha=0.85)
    bf = ax.bar(x + w/2, fc, w, label='女', color=RED,  alpha=0.85)
    ax.set_xticks(x); ax.set_xticklabels(PUB_ORDER, fontsize=10)
    ax.set_title(title, fontsize=10); ax.set_ylabel('人数'); ax.legend(fontsize=9)
    for bar, v in list(zip(bm, mc)) + list(zip(bf, fc)):
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+3, str(v),
                ha='center', va='bottom', fontsize=8)

fig.suptitle('不同版本4条件下的男女人数统计（人头计数）', fontsize=12, fontweight='bold')
plt.tight_layout()
fig.savefig(OUTPUT_DIR / "fig2_4条件男女人数.png", dpi=150, bbox_inches='tight')
plt.close()
print("  → fig2 已保存")


# ════════════════════════════════════════════════════════
#  图3：职业丰富度对比（人头计数 vs 独立职业）
# ════════════════════════════════════════════════════════
print("绘制图3：职业丰富度…")

def richness(df, pub, gender):
    s = df[(df['版本']==pub) & (df['gender']==gender)
           & (~df['职业分类代码'].isin(NON_EMPLOYED))
           & (df['profession'] != '未知')]
    return s['profession'].nunique()

fig, axes = plt.subplots(1, 2, figsize=(14, 5.5))
for ax_i, (df, sfx) in enumerate([(df_head, '人头计数'), (df_uniq, '独立职业')]):
    ax = axes[ax_i]
    fr = [richness(df, p, '女') for p in PUB_ORDER]
    mr = [richness(df, p, '男') for p in PUB_ORDER]
    x = np.arange(len(PUB_ORDER)); w = 0.36
    bf = ax.bar(x - w/2, fr, w, label='女', color=RED,  alpha=0.85)
    bm = ax.bar(x + w/2, mr, w, label='男', color=BLUE, alpha=0.85)
    ax.set_xticks(x); ax.set_xticklabels(PUB_ORDER)
    ax.set_ylabel('职业种类数（个）')
    ax.set_title(f'各版本教材男女职业种类丰富度对比（{sfx}）', fontsize=11)
    ax.legend()
    for bar, v in list(zip(bf, fr)) + list(zip(bm, mr)):
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.3, str(v),
                ha='center', va='bottom', fontsize=10, fontweight='bold')

plt.tight_layout()
fig.savefig(OUTPUT_DIR / "fig3_职业丰富度.png", dpi=150, bbox_inches='tight')
plt.close()
print("  → fig3 已保存")


# ════════════════════════════════════════════════════════
#  图4 & 5：Top10职业集中度（人头计数 / 独立职业）
# ════════════════════════════════════════════════════════
def plot_top10(df: pd.DataFrame, data_label: str, save_path: Path):
    """4版本×2性别的Top10水平柱状图，布局4行2列，所有子图共享X轴尺度"""
    fig, axes = plt.subplots(4, 2, figsize=(16, 26))

    # 第一遍：收集数据并计算全局最大百分比
    all_data = {}
    global_max_pct = 0
    for pub in PUB_ORDER:
        df_pub = df[(df['版本']==pub)
                    & (~df['职业分类代码'].isin(NON_EMPLOYED))
                    & (df['profession'] != '未知')]
        all_data[pub] = {}
        for gender in ['男', '女']:
            df_g = df_pub[df_pub['gender'] == gender]
            total = len(df_g)
            if total == 0:
                all_data[pub][gender] = (None, 0)
                continue
            top10 = df_g['profession'].value_counts().head(10)
            pct   = (top10 / total * 100).round(1)
            all_data[pub][gender] = (pct, total)
            if len(pct) > 0:
                global_max_pct = max(global_max_pct, pct.max())

    shared_xlim = global_max_pct * 1.2
    shared_xlim = max(shared_xlim, 5)

    # 第二遍：绘图
    for pi, pub in enumerate(PUB_ORDER):
        for gi, (gender, color) in enumerate([('男', BLUE), ('女', RED)]):
            ax = axes[pi][gi]
            pct, total = all_data[pub][gender]
            if total == 0 or pct is None:
                ax.set_visible(False); continue

            bars = ax.barh(range(len(pct)), pct.values, color=color, alpha=0.8)
            ax.set_yticks(range(len(pct)))
            ax.set_yticklabels(pct.index, fontsize=10)
            ax.invert_yaxis()
            ax.set_xlim(0, shared_xlim)
            ax.set_xlabel('占比 (%)')
            ax.set_title(f'【{pub}】{gender}性 Top10 占比 (N={total})', fontsize=10)
            for bar, v in zip(bars, pct.values):
                ax.text(bar.get_width()+0.15, bar.get_y()+bar.get_height()/2,
                        f'{v}%', va='center', fontsize=8.5)

    fig.suptitle(f'各版本 – 职业集中度（Top10职业占总从业人数比例）（{data_label}）',
                 fontsize=12, fontweight='bold')
    plt.tight_layout(rect=[0, 0, 1, 0.975])
    fig.savefig(save_path, dpi=130, bbox_inches='tight')
    plt.close()

print("绘制图4：Top10职业集中度（人头计数）…")
plot_top10(df_head, '人头计数', OUTPUT_DIR / "fig4_Top10职业集中度_人头计数.png")
print("  → fig4 已保存")

print("绘制图5：Top10职业集中度（独立职业）…")
plot_top10(df_uniq, '独立职业', OUTPUT_DIR / "fig5_Top10职业集中度_独立职业.png")
print("  → fig5 已保存")


# ════════════════════════════════════════════════════════
#  图6：各版本场景分布（男女分别）
# ════════════════════════════════════════════════════════
print("绘制图6：场景分布…")
SCENARIOS = ['家庭', '学校', '工作场所', '公共场所', '其他']
scene_colors = ['#ED7D31','#A9D18E','#4472C4','#FFC000','#9E9E9E']

fig, axes = plt.subplots(1, 2, figsize=(14, 5))
for gi, (gender, color_list) in enumerate([('男', None), ('女', None)]):
    ax = axes[gi]
    for si, scene in enumerate(SCENARIOS):
        vals = []
        for pub in PUB_ORDER:
            sub = df_head[(df_head['版本']==pub) & (df_head['gender']==gender)]
            total = len(sub)
            cnt = (sub['scenario']==scene).sum()
            vals.append(cnt/total*100 if total > 0 else 0)
        x = np.arange(len(PUB_ORDER))
        bottoms = [sum(
            (df_head[(df_head['版本']==p)&(df_head['gender']==gender)]['scenario']==SCENARIOS[s]).sum()
            / max(len(df_head[(df_head['版本']==p)&(df_head['gender']==gender)]),1)*100
            for s in range(si)
        ) for p in PUB_ORDER]
        ax.bar(x, vals, bottom=bottoms, label=scene, color=scene_colors[si], alpha=0.85)

    ax.set_xticks(x); ax.set_xticklabels(PUB_ORDER)
    ax.set_ylabel('占比 (%)'); ax.set_ylim(0, 105)
    ax.set_title(f'{gender}性人物场景分布（人头计数）', fontsize=11)
    ax.legend(loc='upper right', fontsize=9)

plt.tight_layout()
fig.savefig(OUTPUT_DIR / "fig6_场景分布_男女.png", dpi=150, bbox_inches='tight')
plt.close()
print("  → fig6 已保存")


# ════════════════════════════════════════════════════════
#  输出关键数值供文字分析
# ════════════════════════════════════════════════════════
print("\n" + "="*60)
print("关键数值汇总（供文字分析）")
print("="*60)

print("\n【1. 各条件性别比】")
for cond, label in zip([1,2,3,4], COND_LABELS):
    df_c = apply_cond(df_head, cond)
    print(f"\n  {label}:")
    for pub in PUB_ORDER:
        r = gender_ratio(df_c, pub)
        s = df_c[df_c['版本']==pub]
        m = (s['gender']=='男').sum(); f = (s['gender']=='女').sum()
        print(f"    {pub}: 男{m} / 女{f} = 比值{r:.2f}" if not np.isnan(r) else f"    {pub}: 无法计算（无女性）")

print("\n【2. 职业丰富度】")
for sfx, df in [('人头计数', df_head), ('独立职业', df_uniq)]:
    print(f"\n  {sfx}:")
    for pub in PUB_ORDER:
        fr = richness(df, pub, '女'); mr = richness(df, pub, '男')
        print(f"    {pub}: 女{fr}种 / 男{mr}种 (女/男={fr/mr:.2f})" if mr>0 else f"    {pub}: 女{fr}种 / 男{mr}种")

print("\n【3. 各场景性别构成（人头计数，全部数据）】")
for scene in SCENARIOS:
    print(f"\n  {scene}:")
    for pub in PUB_ORDER:
        sub = df_head[(df_head['版本']==pub) & (df_head['scenario']==scene)]
        m = (sub['gender']=='男').sum(); f = (sub['gender']=='女').sum(); tot = len(sub)
        if tot > 0:
            print(f"    {pub}: 总{tot}人, 男{m}({m/tot*100:.0f}%) 女{f}({f/tot*100:.0f}%)")

print("\n【4. 各版本Top5职业（人头计数，从业人员）】")
for pub in PUB_ORDER:
    print(f"\n  {pub}:")
    df_pub = df_head[(df_head['版本']==pub)
                     & (~df_head['职业分类代码'].isin(NON_EMPLOYED))
                     & (df_head['profession']!='未知')]
    for gender in ['男','女']:
        top5 = df_pub[df_pub['gender']==gender]['profession'].value_counts().head(5)
        print(f"    {gender}: {list(top5.items())}")

print(f"\n全部图表已保存至: {OUTPUT_DIR}")
