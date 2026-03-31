"""
统计分析脚本：汇总所有分类结果，生成多维度统计报表
输出：结果/3.统计结果/统计汇总.xlsx（多个sheet）
维度：全局 / 按版本 / 按年级 / 按版本×年级
"""
import sys, re
from pathlib import Path

import pandas as pd
import numpy as np

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

PROJECT_ROOT = Path(__file__).parent.resolve()
CLASS_1_DIR = PROJECT_ROOT / "结果" / "2.分类结果" / "一位码"
CLASS_3_DIR = PROJECT_ROOT / "结果" / "2.分类结果" / "三位码"
OUTPUT_DIR = PROJECT_ROOT / "结果" / "3.统计结果"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# ── 加载数据 ──
def load_all(directory: Path) -> pd.DataFrame:
    frames = []
    for f in sorted(directory.glob("已分类_*.xlsx")):
        df = pd.read_excel(f)
        df["来源文件"] = f.stem
        # 从文件名解析版本和年级
        # 格式: 已分类_一位码_人教版_1.1_一年级上_人头计数
        name = f.stem
        parts = name.split("_")
        # 找到版本名（人教版/北师大版/苏教版/部编版）
        publisher = ""
        grade_code = ""
        grade_name = ""
        count_type = ""
        for i, p in enumerate(parts):
            if p in ("人教版", "北师大版", "苏教版", "部编版"):
                publisher = p
                if i + 1 < len(parts):
                    grade_code = parts[i + 1]
                if i + 2 < len(parts):
                    grade_name = parts[i + 2]
                break
        if "人头计数" in name:
            count_type = "人头计数"
        elif "独立职业" in name:
            count_type = "独立职业"

        df["版本"] = publisher
        df["年级编号"] = grade_code
        df["年级"] = grade_name
        df["统计方式"] = count_type
        frames.append(df)
        print(f"  {f.name}: {len(df)} 行 → {publisher} {grade_name} ({count_type})")
    if frames:
        return pd.concat(frames, ignore_index=True)
    return pd.DataFrame()


print("=== 加载一位码分类结果 ===")
df_1 = load_all(CLASS_1_DIR)
print(f"\n=== 加载三位码分类结果 ===")
df_3 = load_all(CLASS_3_DIR)
print(f"\n一位码总计: {len(df_1)} 行, 三位码总计: {len(df_3)} 行")


# ── 统计函数 ──
def gender_crosstab(df, code_col="职业分类代码", name_col="职业分类名称"):
    """性别×职业交叉表"""
    if df.empty or code_col not in df.columns:
        return pd.DataFrame()
    return pd.crosstab(
        [df[code_col], df[name_col]],
        df["gender"],
        margins=True, margins_name="合计"
    )


def gender_ratio(df, code_col="职业分类代码", name_col="职业分类名称"):
    """各职业男女比例"""
    if df.empty or code_col not in df.columns:
        return pd.DataFrame()
    ct = pd.crosstab([df[code_col], df[name_col]], df["gender"])
    for col in ["男", "女", "未知"]:
        if col not in ct.columns:
            ct[col] = 0
    ct["合计"] = ct.sum(axis=1)
    ct["男占比%"] = (ct["男"] / ct["合计"] * 100).round(1)
    ct["女占比%"] = (ct["女"] / ct["合计"] * 100).round(1)
    return ct.sort_values("合计", ascending=False)


def gender_summary_row(df, label=""):
    """返回一行性别汇总：男/女/未知/合计/男占比/女占比"""
    if df.empty or "gender" not in df.columns:
        return {}
    vc = df["gender"].value_counts()
    total = vc.sum()
    m = vc.get("男", 0)
    f = vc.get("女", 0)
    u = vc.get("未知", 0)
    return {
        "分组": label,
        "男": m, "女": f, "未知": u, "合计": total,
        "男占比%": round(m / total * 100, 1) if total else 0,
        "女占比%": round(f / total * 100, 1) if total else 0,
    }


def build_gender_overview(df, group_col=None):
    """按分组生成性别概览表"""
    rows = []
    if group_col:
        for name, grp in df.groupby(group_col, sort=True):
            rows.append(gender_summary_row(grp, label=name))
    rows.append(gender_summary_row(df, label="总计"))
    return pd.DataFrame(rows)


# ── 分离人头计数 / 独立职业 ──
df_1_head = df_1[df_1["统计方式"] == "人头计数"].copy()
df_1_uniq = df_1[df_1["统计方式"] == "独立职业"].copy()
df_3_head = df_3[df_3["统计方式"] == "人头计数"].copy()
df_3_uniq = df_3[df_3["统计方式"] == "独立职业"].copy()

print(f"\n一位码 人头: {len(df_1_head)}, 独立: {len(df_1_uniq)}")
print(f"三位码 人头: {len(df_3_head)}, 独立: {len(df_3_uniq)}")


# ── 生成统计报表 ──
print(f"\n{'='*60}")
print("生成统计报表...")

output_file = OUTPUT_DIR / "统计汇总.xlsx"

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:

    # ===== Sheet 1: 性别概览（按版本）=====
    print("  性别概览（按版本）...")
    overview_pub_head = build_gender_overview(df_1_head, "版本")
    overview_pub_head.to_excel(writer, sheet_name="性别概览_按版本_人头", index=False)

    overview_pub_uniq = build_gender_overview(df_1_uniq, "版本")
    overview_pub_uniq.to_excel(writer, sheet_name="性别概览_按版本_独立", index=False)

    # ===== Sheet 2: 性别概览（按年级）=====
    print("  性别概览（按年级）...")
    overview_grade_head = build_gender_overview(df_1_head, "年级")
    overview_grade_head.to_excel(writer, sheet_name="性别概览_按年级_人头", index=False)

    overview_grade_uniq = build_gender_overview(df_1_uniq, "年级")
    overview_grade_uniq.to_excel(writer, sheet_name="性别概览_按年级_独立", index=False)

    # ===== Sheet 3: 性别概览（版本×年级）=====
    print("  性别概览（版本×年级）...")
    df_1_head["版本_年级"] = df_1_head["版本"] + "_" + df_1_head["年级"]
    overview_detail_head = build_gender_overview(df_1_head, "版本_年级")
    overview_detail_head.to_excel(writer, sheet_name="性别概览_版本×年级_人头", index=False)

    df_1_uniq["版本_年级"] = df_1_uniq["版本"] + "_" + df_1_uniq["年级"]
    overview_detail_uniq = build_gender_overview(df_1_uniq, "版本_年级")
    overview_detail_uniq.to_excel(writer, sheet_name="性别概览_版本×年级_独立", index=False)

    # ===== Sheet 4: 一位码 性别×职业 交叉表（全局）=====
    print("  一位码 性别×职业（全局）...")
    gender_crosstab(df_1_head).to_excel(writer, sheet_name="一位码_性别×职业_人头")
    gender_crosstab(df_1_uniq).to_excel(writer, sheet_name="一位码_性别×职业_独立")

    # ===== Sheet 5: 一位码 各职业男女比例（全局）=====
    print("  一位码 男女比例...")
    gender_ratio(df_1_head).to_excel(writer, sheet_name="一位码_男女比例_人头")
    gender_ratio(df_1_uniq).to_excel(writer, sheet_name="一位码_男女比例_独立")

    # ===== Sheet 6: 一位码 按版本分组的性别×职业 =====
    print("  一位码 按版本分组...")
    for pub in sorted(df_1_head["版本"].unique()):
        sub = df_1_head[df_1_head["版本"] == pub]
        ct = gender_crosstab(sub)
        sheet_name = f"一位码_{pub}_人头"
        if len(sheet_name) > 31:
            sheet_name = sheet_name[:31]
        ct.to_excel(writer, sheet_name=sheet_name)

    # ===== Sheet 7: 三位码 性别×职业 交叉表（全局）=====
    print("  三位码 性别×职业（全局）...")
    gender_crosstab(df_3_head).to_excel(writer, sheet_name="三位码_性别×职业_人头")
    gender_crosstab(df_3_uniq).to_excel(writer, sheet_name="三位码_性别×职业_独立")

    # ===== Sheet 8: 三位码 各职业男女比例（全局）=====
    print("  三位码 男女比例...")
    gender_ratio(df_3_head).to_excel(writer, sheet_name="三位码_男女比例_人头")
    gender_ratio(df_3_uniq).to_excel(writer, sheet_name="三位码_男女比例_独立")

    # ===== Sheet 9: 场景分布 =====
    print("  场景分布...")
    if "scenario" in df_1_head.columns:
        scene_gender = pd.crosstab(df_1_head["scenario"], df_1_head["gender"],
                                   margins=True, margins_name="合计")
        scene_gender.to_excel(writer, sheet_name="场景×性别_人头")

        scene_pub = pd.crosstab(df_1_head["scenario"], df_1_head["版本"],
                                margins=True, margins_name="合计")
        scene_pub.to_excel(writer, sheet_name="场景×版本_人头")

    # ===== Sheet 10: 版本间对比 - 一位码职业分布 =====
    print("  版本间对比 - 一位码职业分布...")
    # 每个版本各职业占比
    pub_prof_head = pd.crosstab(
        [df_1_head["职业分类代码"], df_1_head["职业分类名称"]],
        df_1_head["版本"],
        normalize="columns"
    ).round(3) * 100
    pub_prof_head.columns = [f"{c}_占比%" for c in pub_prof_head.columns]
    # 附加绝对数量
    pub_prof_count = pd.crosstab(
        [df_1_head["职业分类代码"], df_1_head["职业分类名称"]],
        df_1_head["版本"]
    )
    pub_compare = pd.concat([pub_prof_count, pub_prof_head], axis=1)
    pub_compare.to_excel(writer, sheet_name="版本对比_一位码职业分布")

    # ===== Sheet 11: 版本间对比 - 各版本男女比 =====
    print("  版本间对比 - 各版本男女比...")
    rows_compare = []
    for pub in sorted(df_1_head["版本"].unique()):
        sub = df_1_head[df_1_head["版本"] == pub]
        for prof_code, prof_name in sub.groupby(["职业分类代码", "职业分类名称"]).groups.keys():
            sub2 = sub[(sub["职业分类代码"] == prof_code) & (sub["职业分类名称"] == prof_name)]
            total = len(sub2)
            m = (sub2["gender"] == "男").sum()
            f = (sub2["gender"] == "女").sum()
            rows_compare.append({
                "版本": pub,
                "职业分类代码": prof_code,
                "职业分类名称": prof_name,
                "男": m, "女": f, "合计": total,
                "男占比%": round(m / total * 100, 1) if total else 0,
                "女占比%": round(f / total * 100, 1) if total else 0,
            })
    df_compare = pd.DataFrame(rows_compare)
    df_compare.to_excel(writer, sheet_name="版本对比_各职业男女比", index=False)

    # ===== Sheet 12: 版本间对比 - 三位码职业分布 =====
    print("  版本间对比 - 三位码职业分布...")
    pub_prof3 = pd.crosstab(
        [df_3_head["职业分类代码"], df_3_head["职业分类名称"]],
        df_3_head["版本"]
    )
    pub_prof3.to_excel(writer, sheet_name="版本对比_三位码职业分布")

    # ===== Sheet 13: 版本×年级 性别对比透视表 =====
    print("  版本×年级 性别对比...")
    pivot_data = []
    for (pub, grade), grp in df_1_head.groupby(["版本", "年级"]):
        total = len(grp)
        m = (grp["gender"] == "男").sum()
        f = (grp["gender"] == "女").sum()
        pivot_data.append({
            "版本": pub, "年级": grade,
            "男": m, "女": f, "合计": total,
            "男占比%": round(m / total * 100, 1) if total else 0,
            "女占比%": round(f / total * 100, 1) if total else 0,
        })
    df_pivot = pd.DataFrame(pivot_data)
    df_pivot.to_excel(writer, sheet_name="版本×年级_性别对比", index=False)

    # 宽格式透视：行=年级，列=版本，值=男占比%
    if not df_pivot.empty:
        pivot_wide = df_pivot.pivot_table(index="年级", columns="版本", values="男占比%")
        pivot_wide.to_excel(writer, sheet_name="透视_年级×版本_男占比")

    # ===== Sheet 14: 版本间场景分布对比 =====
    print("  版本间场景分布对比...")
    if "scenario" in df_1_head.columns:
        scene_pub_pct = pd.crosstab(
            df_1_head["scenario"], df_1_head["版本"],
            normalize="columns"
        ).round(3) * 100
        scene_pub_pct.to_excel(writer, sheet_name="版本对比_场景分布占比")

    # ===== Sheet 15: 原始汇总数据 =====
    print("  原始汇总数据...")
    df_1_head.to_excel(writer, sheet_name="原始_一位码_人头", index=False)
    df_1_uniq.to_excel(writer, sheet_name="原始_一位码_独立", index=False)

print(f"\n统计汇总已导出: {output_file}")

# ── 打印关键统计 ──
print(f"\n{'='*60}")
print(">>> 关键统计摘要 <<<\n")

print("【性别分布 - 全局人头计数】")
total = len(df_1_head)
m = (df_1_head["gender"] == "男").sum()
f = (df_1_head["gender"] == "女").sum()
u = (df_1_head["gender"] == "未知").sum()
print(f"  男: {m} ({m/total*100:.1f}%), 女: {f} ({f/total*100:.1f}%), 未知: {u} ({u/total*100:.1f}%), 合计: {total}")

print("\n【按版本 - 人头计数】")
print(overview_pub_head.to_string(index=False))

print("\n【按年级 - 人头计数】")
print(overview_grade_head.to_string(index=False))

print("\n【一位码职业分布 TOP10 - 人头计数】")
top10 = df_1_head.groupby(["职业分类代码", "职业分类名称"]).size().reset_index(name="数量")
top10 = top10.sort_values("数量", ascending=False).head(10)
print(top10.to_string(index=False))

print(f"\n统计分析完成!")
