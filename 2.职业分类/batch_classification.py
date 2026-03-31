"""
批量职业分类脚本：对 结果/1.识别结果/ 下所有Excel做一位码+三位码分类
复用 2.职业分类/ 下的分类表和API配置
"""
import os, json, time, sys
from pathlib import Path

import pandas as pd
from openai import OpenAI
from dotenv import load_dotenv

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
if sys.stderr.encoding != 'utf-8':
    sys.stderr.reconfigure(encoding='utf-8')

# ── 加载API配置（职业分类阶段）──
CLASS_DIR = Path(__file__).parent
load_dotenv(CLASS_DIR / ".env")

API_KEY = os.getenv("API_KEY")
BASE_URL = os.getenv("BASE_URL", "https://dashscope.aliyuncs.com/compatible-mode/v1")
MODEL_NAME = os.getenv("MODEL_NAME", "qwen-plus-2025-12-01")

if not API_KEY:
    raise ValueError("请在 2.职业分类/.env 文件中设置 API_KEY（当前脚本位于 2.职业分类/ 目录下）")

client = OpenAI(api_key=API_KEY, base_url=BASE_URL)

# ── 路径 ──
PROJECT_ROOT = Path(__file__).parent.parent.resolve()
INPUT_DIR = PROJECT_ROOT / "结果" / "1.识别结果"
OUTPUT_DIR_1 = PROJECT_ROOT / "结果" / "2.分类结果" / "一位码"
OUTPUT_DIR_3 = PROJECT_ROOT / "结果" / "2.分类结果" / "三位码"
OUTPUT_DIR_1.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR_3.mkdir(parents=True, exist_ok=True)

# ── 加载分类表 ──
df_class1 = pd.read_excel(CLASS_DIR / "职业分类.xlsx")
if 10 not in df_class1['代码'].values:
    df_class1 = pd.concat([df_class1, pd.DataFrame([{
        '代码': 10, '分类名称': '非从业人员-其他',
        '分类内容': '无法归入学生、儿童或家庭角色的其他非从业人员，或身份不明者。'
    }])], ignore_index=True)

df_class3 = pd.read_excel(CLASS_DIR / "职业分类（三位码）.xlsx")

# 一位码全量分类参考
class_info_1_full = ""
for _, row in df_class1.iterrows():
    content = str(row.get('分类内容', ''))[:50]
    class_info_1_full += f"代码[{row['代码']}] {row['分类名称']}: {content}\n"

# 一位码限制分类（0/9/10）
restricted_codes = [0, 9, 10]
class_info_1_restricted = ""
for _, row in df_class1[df_class1['代码'].isin(restricted_codes)].iterrows():
    content = str(row.get('分类内容', ''))[:50]
    class_info_1_restricted += f"代码[{row['代码']}] {row['分类名称']}: {content}\n"

# 三位码全量分类参考
class_info_3_full = ""
for _, row in df_class3.iterrows():
    content = str(row.get('分类内容', ''))[:60]
    class_info_3_full += f"代码[{row['代码']}] {row['分类名称']}: {content}\n"

# 代码→名称映射
code_to_name_1 = df_class1.set_index('代码')['分类名称'].to_dict()
code_to_name_3 = df_class3.set_index('代码')['分类名称'].to_dict()
for _, row in df_class1[df_class1['代码'].isin(restricted_codes)].iterrows():
    code_to_name_3[row['代码']] = row['分类名称']

print(f"一位码分类: {len(df_class1)} 类")
print(f"三位码分类: {len(df_class3)} 类")


# ── 分类函数 ──
def classify_batch(descriptions: list, class_info_str: str, restricted: bool = False, digit3: bool = False) -> dict:
    if restricted:
        constraint = """【重要限制】这组数据通过视觉无法确认职业，属于非从业人员场景。
请仅从以下三个类别中选择，严禁归入其他职业类：
- 0 (学生/儿童)
- 9 (家庭角色，如母亲、父亲、祖辈)
- 10 (其他非从业人员，或无法判断的非职业身份)"""
        example = '{"背书包的小学生": 0, "做饭的母亲": 9, "路人": 10}'
    elif digit3:
        constraint = """请根据描述归入最合适的三位码类别。
注意：三位码代码为5位数字（如10100、20200等），请务必使用完整的5位代码。
如果描述的是学生、儿童，请归入代码0；家庭角色归入代码9。"""
        example = '{"正在讲课的老师": 20800, "警察": 30100, "医生": 20600}'
    else:
        constraint = "请根据描述归入最合适的类别（0-10类）。如果描述的是学生、儿童或家庭角色，请优先归入0或9类。"
        example = '{"正在讲课的老师": 2, "背书包的小学生": 0, "做饭的母亲": 9}'

    prompt = f"""任务：请将以下人物描述归类到《职业分类大典》的对应类别中。

参考分类标准：
{class_info_str}

{constraint}

待分类的描述列表：
{json.dumps(descriptions, ensure_ascii=False)}

要求：
1. 严格输出JSON格式，Key为原描述，Value为对应的"代码"（数字）。
2. 不要输出任何解释性文字。

输出示例：
{example}"""

    try:
        response = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.1,
            response_format={"type": "json_object"}
        )
        return json.loads(response.choices[0].message.content)
    except Exception as e:
        print(f"    API调用出错: {e}")
        return {}


def process_file(file_path: Path, output_path: Path, class_info_full: str, class_info_restricted: str,
                 code_to_name: dict, batch_size: int = 50, digit3: bool = False):
    """处理单个识别结果文件"""
    df = pd.read_excel(file_path)
    if df.empty:
        print(f"    空文件，跳过")
        return

    mapping_dict = {}
    has_identifier = 'identifier' in df.columns

    # 轨道A：profession != "未知" → 用profession匹配全量分类表
    known_mask = df['profession'] != '未知'
    descs_known = df.loc[known_mask, 'profession'].dropna().unique().tolist()
    if descs_known:
        print(f"    轨道A: {len(descs_known)} 个描述")
        for i in range(0, len(descs_known), batch_size):
            batch = descs_known[i:i+batch_size]
            res = classify_batch(batch, class_info_full, restricted=False, digit3=digit3)
            mapping_dict.update(res)
            time.sleep(0.3)

    # 轨道B：profession == "未知" 且有identifier列 → 用identifier匹配限制分类表(0/9/10)
    if has_identifier:
        unknown_mask = df['profession'] == '未知'
        descs_unknown = df.loc[unknown_mask, 'identifier'].dropna().unique().tolist()
        if descs_unknown:
            print(f"    轨道B: {len(descs_unknown)} 个描述")
            for i in range(0, len(descs_unknown), batch_size):
                batch = descs_unknown[i:i+batch_size]
                res = classify_batch(batch, class_info_restricted, restricted=True)
                mapping_dict.update(res)
                time.sleep(0.3)

    # 回填
    def apply_mapping(row):
        if has_identifier and row['profession'] == '未知':
            key = row['identifier']
        else:
            key = row['profession']
        return mapping_dict.get(key, '匹配失败')

    df['职业分类代码'] = df.apply(apply_mapping, axis=1)
    df['职业分类名称'] = df['职业分类代码'].map(code_to_name)

    df.to_excel(output_path, index=False)
    matched = (df['职业分类代码'] != '匹配失败').sum()
    print(f"    匹配: {matched}/{len(df)}")


# ── 从已分类人头计数派生独立职业 ──
def derive_unique_from_headcount(headcount_path: Path, unique_path: Path):
    """从已分类的人头计数文件，按 (page, profession, gender) 去重得到独立职业表"""
    df = pd.read_excel(headcount_path)
    if df.empty:
        df.to_excel(unique_path, index=False)
        return 0
    df_unique = (
        df.groupby(["page", "profession", "gender"], as_index=False)
        .agg(
            source_type=("source_type", "first"),
            scenario=("scenario", "first"),
            count=("identifier", "count"),
            职业分类代码=("职业分类代码", "first"),
            职业分类名称=("职业分类名称", "first"),
        )
    )
    df_unique.to_excel(unique_path, index=False)
    return len(df_unique)


# ── 主流程 ──
if __name__ == "__main__":
    # 只找人头计数文件进行分类（独立职业从分类后的人头计数派生）
    headcount_files = [f for f in sorted(INPUT_DIR.glob("*_人头计数.xlsx"))]
    print(f"\n发现 {len(headcount_files)} 个人头计数识别结果文件")
    print("=" * 60)

    # === 一位码分类 ===
    print("\n>>> 阶段一：一位码分类（人头计数）<<<")
    for idx, file_path in enumerate(headcount_files, 1):
        output_path = OUTPUT_DIR_1 / f"已分类_一位码_{file_path.name}"
        if output_path.exists():
            print(f"[{idx}/{len(headcount_files)}] 跳过: {file_path.name}")
            continue
        print(f"[{idx}/{len(headcount_files)}] {file_path.name}")
        process_file(file_path, output_path, class_info_1_full, class_info_1_restricted,
                     code_to_name_1, batch_size=50, digit3=False)
        time.sleep(0.5)

    # 一位码：从已分类人头计数派生独立职业
    print("\n>>> 一位码：派生独立职业表 <<<")
    for hf in sorted(OUTPUT_DIR_1.glob("已分类_一位码_*_人头计数.xlsx")):
        unique_name = hf.name.replace("人头计数", "独立职业")
        unique_path = OUTPUT_DIR_1 / unique_name
        n = derive_unique_from_headcount(hf, unique_path)
        print(f"  {hf.name} → {n} 条独立职业")

    print(f"\n{'=' * 60}")
    print("一位码分类完成!")

    # === 三位码分类 ===
    print("\n>>> 阶段二：三位码分类（人头计数）<<<")
    for idx, file_path in enumerate(headcount_files, 1):
        output_path = OUTPUT_DIR_3 / f"已分类_三位码_{file_path.name}"
        if output_path.exists():
            print(f"[{idx}/{len(headcount_files)}] 跳过: {file_path.name}")
            continue
        print(f"[{idx}/{len(headcount_files)}] {file_path.name}")
        process_file(file_path, output_path, class_info_3_full, class_info_1_restricted,
                     code_to_name_3, batch_size=30, digit3=True)
        time.sleep(0.5)

    # 三位码：从已分类人头计数派生独立职业
    print("\n>>> 三位码：派生独立职业表 <<<")
    for hf in sorted(OUTPUT_DIR_3.glob("已分类_三位码_*_人头计数.xlsx")):
        unique_name = hf.name.replace("人头计数", "独立职业")
        unique_path = OUTPUT_DIR_3 / unique_name
        n = derive_unique_from_headcount(hf, unique_path)
        print(f"  {hf.name} → {n} 条独立职业")

    print(f"\n{'=' * 60}")
    print("三位码分类完成!")
    print("\n全部职业分类完毕!")
