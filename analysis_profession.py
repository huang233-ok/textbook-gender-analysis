"""
职业分析脚本（完整版 v2）
- 职业丰富度：三种粒度（一位码/三位码/原始职业），仅用人头计数数据
- 职业集中度Top10：三种粒度 × 两套数据（人头/独立），男女X轴范围一致
- 命名标准化：合并"老师"→"教师"等常见变体
- 独立职业表已修正：按 (page, profession, gender) 去重，保留性别信息
输出：结果/4.分析图表/profession_*.png
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

PROJECT_ROOT = Path("F:/Desktop/教材/python project/2026.3_整体优化")
DIR_1 = PROJECT_ROOT / "结果" / "2.分类结果" / "一位码"
DIR_3 = PROJECT_ROOT / "结果" / "2.分类结果" / "三位码"
OUT   = PROJECT_ROOT / "结果" / "4.分析图表"
OUT.mkdir(parents=True, exist_ok=True)

PUB_ORDER = ["人教版", "北师大版", "苏教版", "部编版"]
BLUE, RED = '#4472C4', '#C0504D'
NON_EMP = {0, 9, 10}   # 排除：学生/儿童、家庭角色、其他非从业

# ── 命名标准化映射 ──
PROFESSION_MERGE = {
    '老师': '教师',
    '私塾老师': '教师',
    '体育老师': '教师',
    '幼儿园老师': '教师',
}

def normalize_profession(df: pd.DataFrame) -> pd.DataFrame:
    """标准化profession列中的常见变体"""
    df = df.copy()
    df['profession'] = df['profession'].replace(PROFESSION_MERGE)
    return df


# ── 加载数据 ──
def load_dir(directory: Path, suffix: str) -> pd.DataFrame:
    frames = []
    for f in sorted(directory.glob(f"已分类_*_{suffix}.xlsx")):
        df = pd.read_excel(f)
        parts = f.stem.split("_")
        pub = next((p for p in parts if p in PUB_ORDER), "")
        df["版本"] = pub
        df["职业分类代码"] = pd.to_numeric(df["职业分类代码"], errors='coerce')
        frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

print("加载数据…")
df_1h = normalize_profession(load_dir(DIR_1, "人头计数"))
df_1u = normalize_profession(load_dir(DIR_1, "独立职业"))
df_3h = normalize_profession(load_dir(DIR_3, "人头计数"))
df_3u = normalize_profession(load_dir(DIR_3, "独立职业"))
print(f"  一位码 人头{len(df_1h)} / 独立{len(df_1u)}")
print(f"  三位码 人头{len(df_3h)} / 独立{len(df_3u)}")

# 构建 一位码代码→名称 映射
_code_name_1 = (df_1h[['职业分类代码','职业分类名称']]
                .dropna()
                .drop_duplicates()
                .set_index('职业分类代码')['职业分类名称']
                .to_dict())
# 构建 三位码代码→名称 映射
_code_name_3 = (df_3h[['职业分类代码','职业分类名称']]
                .dropna()
                .drop_duplicates()
                .set_index('职业分类代码')['职业分类名称']
                .to_dict())


# ════════════════════════════════════════════════════════════
#  通用过滤函数
# ════════════════════════════════════════════════════════════
def filter_employed(df: pd.DataFrame) -> pd.DataFrame:
    """排除非从业人员（code 0/9/10）及 profession='未知'，code为NaN者（分类失败）亦排除"""
    mask = (
        df['职业分类代码'].notna()
        & (~df['职业分类代码'].isin(NON_EMP))
        & (df['profession'] != '未知')
    )
    return df[mask]


# ════════════════════════════════════════════════════════════
#  图A：职业丰富度（三种粒度）
#  注：丰富度只用人头计数，原因：独立职业表的性别用众数赋值，
#       会导致少数性别的职业被错误归入另一性别，低估其丰富度
# ════════════════════════════════════════════════════════════
print("\n绘制图A：职业丰富度…")

GRANULARITIES = [
    ("一位码", "职业分类代码", df_1h, "（分类代码，约8个有效类）"),
    ("三位码", "职业分类代码", df_3h, "（分类代码，最多78类）"),
    ("原始职业", "profession",  df_1h, "（模型识别原始文本）"),
]

fig, axes = plt.subplots(1, 3, figsize=(18, 5))

for ax, (gran_name, col, df_src, note) in zip(axes, GRANULARITIES):
    df_emp = filter_employed(df_src)
    x = np.arange(len(PUB_ORDER)); w = 0.36

    f_rich = [df_emp[(df_emp['版本']==p) & (df_emp['gender']=='女')][col].nunique() for p in PUB_ORDER]
    m_rich = [df_emp[(df_emp['版本']==p) & (df_emp['gender']=='男')][col].nunique() for p in PUB_ORDER]

    bf = ax.bar(x - w/2, f_rich, w, label='女', color=RED,  alpha=0.85)
    bm = ax.bar(x + w/2, m_rich, w, label='男', color=BLUE, alpha=0.85)
    ax.set_xticks(x); ax.set_xticklabels(PUB_ORDER)
    ax.set_ylabel('职业种类数')
    ax.set_title(f'{gran_name}丰富度\n{note}', fontsize=10)
    ax.legend(fontsize=9)
    for bar, v in list(zip(bf, f_rich)) + list(zip(bm, m_rich)):
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.3, str(v),
                ha='center', va='bottom', fontsize=10, fontweight='bold')

fig.suptitle('各版本男女职业丰富度对比（人头计数）', fontsize=12, fontweight='bold')
plt.tight_layout()
fig.savefig(OUT / "profA_职业丰富度.png", dpi=150, bbox_inches='tight')
plt.close()
print("  → profA 已保存")


# ════════════════════════════════════════════════════════════
#  图B：一位码职业分布（完整分布，非Top10）
#  展示所有有效职业大类（代码1-8），男女同轴
# ════════════════════════════════════════════════════════════
print("绘制图B：一位码完整分布…")

EMP_CODES_1 = sorted([c for c in df_1h['职业分类代码'].dropna().unique()
                       if c not in NON_EMP and not np.isnan(c)])

def short_name(name_str, n=12):
    """缩短名称用于Y轴标签"""
    s = str(name_str)
    return s[:n] + '…' if len(s) > n else s

for data_label, df_src in [("人头计数", df_1h), ("独立职业", df_1u)]:
    fig, axes = plt.subplots(1, 4, figsize=(20, 7), sharey=False)

    # 第一遍：计算所有子图中的最大百分比值
    all_pcts = []
    for pi, pub in enumerate(PUB_ORDER):
        df_emp = filter_employed(df_src)
        df_pub = df_emp[df_emp['版本'] == pub]
        m_total = (df_pub['gender']=='男').sum()
        f_total = (df_pub['gender']=='女').sum()
        for code in EMP_CODES_1:
            mc = ((df_pub['职业分类代码']==code) & (df_pub['gender']=='男')).sum()
            fc = ((df_pub['职业分类代码']==code) & (df_pub['gender']=='女')).sum()
            all_pcts.append(mc / m_total * 100 if m_total > 0 else 0)
            all_pcts.append(fc / f_total * 100 if f_total > 0 else 0)
    global_xlim = max(all_pcts) * 1.15 if all_pcts else 10
    global_xlim = max(global_xlim, 5)

    # 第二遍：绘图
    for pi, pub in enumerate(PUB_ORDER):
        ax = axes[pi]
        df_emp = filter_employed(df_src)
        df_pub = df_emp[df_emp['版本'] == pub]

        m_total = (df_pub['gender']=='男').sum()
        f_total = (df_pub['gender']=='女').sum()

        m_pct = []
        f_pct = []
        cat_names = []

        for code in EMP_CODES_1:
            mc = ((df_pub['职业分类代码']==code) & (df_pub['gender']=='男')).sum()
            fc = ((df_pub['职业分类代码']==code) & (df_pub['gender']=='女')).sum()
            m_pct.append(mc / m_total * 100 if m_total > 0 else 0)
            f_pct.append(fc / f_total * 100 if f_total > 0 else 0)
            cat_names.append(short_name(_code_name_1.get(code, str(code))))

        y = np.arange(len(EMP_CODES_1)); w = 0.36
        bm = ax.barh(y + w/2, m_pct, w, label='男', color=BLUE, alpha=0.8)
        bf = ax.barh(y - w/2, f_pct, w, label='女', color=RED,  alpha=0.8)
        ax.set_yticks(y)
        ax.set_yticklabels(cat_names, fontsize=8)
        ax.set_xlabel('占比 (%)'); ax.invert_yaxis()
        ax.set_xlim(0, global_xlim)
        ax.set_title(f'{pub}\n男N={m_total} 女N={f_total}', fontsize=9)
        ax.legend(fontsize=8); ax.grid(axis='x', alpha=0.25)

    fig.suptitle(f'各版本一位码职业大类分布（从业人员，{data_label}）', fontsize=12, fontweight='bold')
    plt.tight_layout()
    fig.savefig(OUT / f"profB_一位码分布_{data_label}.png", dpi=140, bbox_inches='tight')
    plt.close()
    print(f"  → profB_{data_label} 已保存")


# ════════════════════════════════════════════════════════════
#  通用Top10集中度函数（男女X轴范围相同）
# ════════════════════════════════════════════════════════════
def plot_top10_matched_axis(df_src, col, name_col, data_label, gran_label, save_path):
    """
    4×2 subplots (4版本 × 2性别), 所有子图共享同一X轴范围
    col        : 用于计数的列（'profession' 或 '职业分类名称'）
    name_col   : 用于显示的列（同col，或转换后的名称）
    """
    df_emp = filter_employed(df_src)

    fig, axes = plt.subplots(4, 2, figsize=(16, 28))

    # 第一遍：收集所有子图数据，计算全局最大百分比
    all_top_data = {}
    global_max_pct = 0
    for pi, pub in enumerate(PUB_ORDER):
        df_pub = df_emp[df_emp['版本'] == pub]
        all_top_data[pub] = {}
        for gender in ['男', '女']:
            df_g = df_pub[df_pub['gender'] == gender]
            total = len(df_g)
            if total == 0:
                all_top_data[pub][gender] = (pd.Series(dtype=float), 0)
                continue
            counts = df_g[col].value_counts().head(10)
            pcts   = (counts / total * 100).round(1)
            all_top_data[pub][gender] = (pcts, total)
            if len(pcts) > 0:
                global_max_pct = max(global_max_pct, pcts.max())

    shared_xlim = global_max_pct * 1.25
    shared_xlim = max(shared_xlim, 5)  # 至少5%

    # 第二遍：绘图
    for pi, pub in enumerate(PUB_ORDER):
        for gi, (gender, color) in enumerate([('男', BLUE), ('女', RED)]):
            ax = axes[pi][gi]
            pcts, total = all_top_data[pub][gender]

            if total == 0 or len(pcts) == 0:
                ax.text(0.5, 0.5, '无数据', ha='center', va='center',
                        transform=ax.transAxes, fontsize=12)
                ax.set_xlim(0, shared_xlim)
                continue

            bars = ax.barh(range(len(pcts)), pcts.values, color=color, alpha=0.8)
            ax.set_yticks(range(len(pcts)))
            ax.set_yticklabels(
                [short_name(str(idx), 14) for idx in pcts.index],
                fontsize=9
            )
            ax.invert_yaxis()
            ax.set_xlim(0, shared_xlim)
            ax.set_xlabel('占总从业人数 (%)', fontsize=8)
            ax.set_title(f'【{pub}】{gender}性 Top10  (N={total})', fontsize=9)

            for bar, v in zip(bars, pcts.values):
                ax.text(bar.get_width() + shared_xlim * 0.01,
                        bar.get_y() + bar.get_height() / 2,
                        f'{v}%', va='center', fontsize=8)
            ax.grid(axis='x', alpha=0.2)

    fig.suptitle(
        f'各版本 – {gran_label}职业集中度（Top10占总从业人数比例，{data_label}）\n'
        f'注：所有子图X轴范围统一',
        fontsize=11, fontweight='bold'
    )
    plt.tight_layout(rect=[0, 0, 1, 0.97])
    fig.savefig(save_path, dpi=130, bbox_inches='tight')
    plt.close()


# ════════════════════════════════════════════════════════════
#  图C：三位码 Top10 集中度
# ════════════════════════════════════════════════════════════
print("绘制图C：三位码 Top10 集中度…")
for data_label, df_src in [("人头计数", df_3h), ("独立职业", df_3u)]:
    save_path = OUT / f"profC_三位码Top10_{data_label}.png"
    plot_top10_matched_axis(df_src, '职业分类名称', '职业分类名称',
                            data_label, "三位码", save_path)
    print(f"  → profC_{data_label} 已保存")


# ════════════════════════════════════════════════════════════
#  图D：原始职业识别 Top10 集中度
# ════════════════════════════════════════════════════════════
print("绘制图D：原始职业 Top10 集中度…")
for data_label, df_src in [("人头计数", df_1h), ("独立职业", df_1u)]:
    save_path = OUT / f"profD_原始职业Top10_{data_label}.png"
    plot_top10_matched_axis(df_src, 'profession', 'profession',
                            data_label, "原始职业文本", save_path)
    print(f"  → profD_{data_label} 已保存")


# ════════════════════════════════════════════════════════════
#  打印关键数值
# ════════════════════════════════════════════════════════════
print("\n" + "="*60)
print("关键数值")
print("="*60)

print("\n【职业丰富度（人头计数，从业人员）】")
for gran_name, col, df_src, _ in GRANULARITIES:
    df_emp = filter_employed(df_src)
    print(f"\n  {gran_name}:")
    for pub in PUB_ORDER:
        fr = df_emp[(df_emp['版本']==pub)&(df_emp['gender']=='女')][col].nunique()
        mr = df_emp[(df_emp['版本']==pub)&(df_emp['gender']=='男')][col].nunique()
        print(f"    {pub}: 女{fr} / 男{mr} = {fr/mr:.2f}" if mr else f"    {pub}: 女{fr} / 男0")

print("\n【一位码职业分布（人头计数，各版本占比）】")
df_emp1 = filter_employed(df_1h)
for pub in PUB_ORDER:
    df_pub = df_emp1[df_emp1['版本']==pub]
    m_tot = (df_pub['gender']=='男').sum()
    f_tot = (df_pub['gender']=='女').sum()
    print(f"\n  {pub}（男N={m_tot} 女N={f_tot}）:")
    for code in EMP_CODES_1:
        name = _code_name_1.get(code,'?')[:15]
        mc = ((df_pub['职业分类代码']==code)&(df_pub['gender']=='男')).sum()
        fc = ((df_pub['职业分类代码']==code)&(df_pub['gender']=='女')).sum()
        if mc+fc > 0:
            print(f"    [{code}]{name}: 男{mc}({mc/m_tot*100:.0f}%) 女{fc}({fc/f_tot*100:.0f}%)")

print(f"\n所有图表已保存至: {OUT}")
