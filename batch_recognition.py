"""
批量识别脚本：扫描 数据/道法教材/ 下所有PDF，运行人物识别，输出到 结果/1.识别结果/
输出命名规则：{版本}_{编号}_{科目}_人头计数.xlsx / _独立职业.xlsx
"""
import os, json, base64, time, re, sys

# 强制无缓冲输出
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
if sys.stderr.encoding != 'utf-8':
    sys.stderr.reconfigure(encoding='utf-8')
from pathlib import Path
from enum import Enum

import fitz  # PyMuPDF
import pandas as pd
from openai import OpenAI
from dotenv import load_dotenv
from pydantic import BaseModel, Field
from typing import List

# ── 加载API配置 ──
load_dotenv(Path(__file__).parent / "1.初始识别" / ".env")

client = OpenAI(
    base_url=os.getenv("OPENROUTER_BASE_URL", "https://openrouter.ai/api/v1"),
    api_key=os.getenv("OPENROUTER_API_KEY"),
)
MODEL_NAME = os.getenv("MODEL_NAME", "openai/gpt-5-chat")

PROJECT_ROOT = Path(__file__).parent.resolve()
DATA_DIR = PROJECT_ROOT / "数据" / "道法教材"
RESULT_DIR = PROJECT_ROOT / "结果" / "1.识别结果"
RESULT_DIR.mkdir(parents=True, exist_ok=True)

# ── 数据模型 ──
class GenderEnum(str, Enum):
    MALE = "男"; FEMALE = "女"; UNKNOWN = "未知"

class SourceTypeEnum(str, Enum):
    ILLUSTRATION = "插图"; ILLUSTRATION_AND_TEXT = "插图和文本"

class ScenarioEnum(str, Enum):
    FAMILY = "家庭"; SCHOOL = "学校"; WORKPLACE = "工作场所"; PUBLIC = "公共场所"; OTHER = "其他"

class Character(BaseModel):
    identifier: str = Field(description="人物标识")
    profession: str = Field(description="职业")
    gender: GenderEnum = Field(description="性别")
    source_type: SourceTypeEnum = Field(description="判断依据")
    scenario: ScenarioEnum = Field(description="场景类型")

class PageResult(BaseModel):
    page: int = Field(description="页码")
    characters: List[Character] = Field(default_factory=list)

class BookResult(BaseModel):
    results: List[PageResult] = Field(default_factory=list)

# ── Prompt ──
SYSTEM_PROMPT = """你是一个教材插图分析专家。你的任务是识别教材页面图片中出现的所有人物，并提取以下信息：

对于每个人物：
1. identifier: 人物标识——如果能从文本/插图中得知姓名则用姓名，否则用简短外貌描述（如"戴眼镜的中年男性"）
2. profession: 职业——根据插图和文本判断人物的职业或身份（如"医生"、"学生"、"警察"），如果无法判断则填"未知"
3. gender: 性别——"男"、"女"或"未知"
4. source_type: 判断依据——"插图"（仅根据插图判断）或"插图和文本"（结合了文本信息）
5. scenario: 场景类型——"家庭"、"学校"、"工作场所"、"公共场所"或"其他"

注意事项：
- 只识别插图中可见的人物，不要凭空想象
- 每个可辨认的人物都要单独记录（包括集体照中的每个人）
- 如果一页中没有人物插图，返回空的characters列表
- 仔细区分不同人物，不要遗漏也不要重复
- 职业判断要结合插图内容和页面文本信息

请以JSON格式返回结果。"""

# ── PDF处理 ──
def pdf_page_to_base64(pdf_path: str, page_num: int, dpi: int = 200) -> str:
    doc = fitz.open(pdf_path)
    page = doc[page_num]
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    pix = page.get_pixmap(matrix=mat)
    img_bytes = pix.tobytes("png")
    doc.close()
    return base64.b64encode(img_bytes).decode("utf-8")

def get_pdf_page_count(pdf_path: str) -> int:
    doc = fitz.open(pdf_path)
    count = len(doc)
    doc.close()
    return count

def clean_enum_value(val) -> str:
    if hasattr(val, 'value'):
        return val.value
    return str(val)

# ── API调用 ──
def analyze_page(pdf_path: str, page_num: int, max_retries: int = 3) -> PageResult:
    img_b64 = pdf_page_to_base64(pdf_path, page_num)
    page_display = page_num + 1

    user_content = [
        {"type": "text", "text": f"这是教材的第{page_display}页。请识别此页中所有人物并提取信息。返回JSON格式，包含page（页码={page_display}）和characters列表。"},
        {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{img_b64}"}}
    ]

    for attempt in range(max_retries):
        try:
            response = client.chat.completions.create(
                model=MODEL_NAME,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": user_content}
                ],
                response_format={"type": "json_object"},
                temperature=0.1,
                max_tokens=4096,
            )
            content = response.choices[0].message.content
            data = json.loads(content)

            if "characters" in data:
                chars_raw = data["characters"]
            elif "results" in data and len(data["results"]) > 0:
                chars_raw = data["results"][0].get("characters", [])
            else:
                chars_raw = []

            # 枚举合法值集合，用于容错映射
            VALID_GENDERS = {e.value for e in GenderEnum}
            VALID_SOURCES = {e.value for e in SourceTypeEnum}
            VALID_SCENARIOS = {e.value for e in ScenarioEnum}

            characters = []
            for c in chars_raw:
                try:
                    raw_gender = c.get("gender", "未知")
                    raw_source = c.get("source_type", "插图")
                    raw_scenario = c.get("scenario", "其他")
                    # 容错：非标准值映射到默认值
                    gender = raw_gender if raw_gender in VALID_GENDERS else "未知"
                    source_type = raw_source if raw_source in VALID_SOURCES else "插图"
                    scenario = raw_scenario if raw_scenario in VALID_SCENARIOS else "其他"

                    char = Character(
                        identifier=c.get("identifier", "未知"),
                        profession=c.get("profession", "未知"),
                        gender=gender,
                        source_type=source_type,
                        scenario=scenario,
                    )
                    characters.append(char)
                except Exception as e:
                    print(f"    警告: 第{page_display}页某人物解析失败: {e}")

            result = PageResult(page=page_display, characters=characters)
            print(f"    第{page_display}页: {len(characters)} 个人物")
            return result

        except Exception as e:
            print(f"    第{page_display}页 第{attempt+1}次尝试失败: {e}")
            if attempt < max_retries - 1:
                time.sleep(2 ** (attempt + 1))

    print(f"    第{page_display}页: 全部重试失败，返回空结果")
    return PageResult(page=page_display, characters=[])

# ── 结果转换 ──
def results_to_dataframes(book_result: BookResult):
    rows = []
    for page_result in book_result.results:
        for char in page_result.characters:
            rows.append({
                "page": page_result.page,
                "identifier": char.identifier,
                "profession": char.profession,
                "gender": clean_enum_value(char.gender),
                "source_type": clean_enum_value(char.source_type),
                "scenario": clean_enum_value(char.scenario),
            })

    df_headcount = pd.DataFrame(rows)
    if df_headcount.empty:
        df_unique = pd.DataFrame(columns=["page", "profession", "gender", "source_type", "scenario", "count"])
        return df_headcount, df_unique

    df_unique = (
        df_headcount
        .groupby(["page", "profession"], as_index=False)
        .agg(
            gender=("gender", lambda x: x.mode().iloc[0] if not x.mode().empty else "未知"),
            source_type=("source_type", "first"),
            scenario=("scenario", "first"),
            count=("identifier", "count"),
        )
    )
    return df_headcount, df_unique


# ── PDF文件名 → 规范化输出名 ──
GRADE_MAP = {
    "1.1": "一年级上", "1.2": "一年级下",
    "3.2": "三年级下",
    "6.1": "六年级上", "6.2": "六年级下",
}

def parse_pdf_info(pdf_path: Path):
    """从路径解析 版本、编号、年级名"""
    publisher = pdf_path.parent.name  # 人教版 / 北师大版 / 苏教版 / 部编版（新版）
    filename = pdf_path.name

    # 从文件名开头提取编号（如 "1.1", "3.2", "6.1"）
    m = re.match(r"(\d+\.\d+)", filename)
    if not m:
        return None
    code = m.group(1)

    # 部编版最后一个文件：名称写的 6.1 但内容是"六年级下册"，修正为 6.2
    if "部编版" in publisher and code == "6.1" and "下册" in filename:
        code = "6.2"

    grade_name = GRADE_MAP.get(code, code)

    # 简化版本名
    pub_short = publisher.replace("（新版）", "")

    return {
        "publisher": pub_short,
        "code": code,
        "grade": grade_name,
        "output_prefix": f"{pub_short}_{code}_{grade_name}",
    }


def discover_pdfs():
    """扫描所有子目录下的PDF文件"""
    pdfs = []
    for pub_dir in sorted(DATA_DIR.iterdir()):
        if not pub_dir.is_dir():
            continue
        for pdf_file in sorted(pub_dir.glob("*.pdf")):
            info = parse_pdf_info(pdf_file)
            if info:
                pdfs.append((pdf_file, info))
            else:
                print(f"  警告: 无法解析文件名 {pdf_file.name}，跳过")
    return pdfs


def process_one_pdf(pdf_path: Path, output_prefix: str):
    """识别单本PDF，保存人头计数表和独立职业表"""
    path_head = RESULT_DIR / f"{output_prefix}_人头计数.xlsx"
    path_unique = RESULT_DIR / f"{output_prefix}_独立职业.xlsx"

    # 跳过已处理的文件
    if path_head.exists() and path_unique.exists():
        print(f"  已存在，跳过: {output_prefix}")
        return True

    total_pages = get_pdf_page_count(str(pdf_path))
    print(f"  总页数: {total_pages}")

    all_page_results = []
    for i in range(total_pages):
        result = analyze_page(str(pdf_path), i)
        all_page_results.append(result)
        # 避免API限流
        time.sleep(0.5)

    book_result = BookResult(results=all_page_results)
    df_headcount, df_unique = results_to_dataframes(book_result)

    df_headcount.to_excel(path_head, index=False)
    df_unique.to_excel(path_unique, index=False)

    print(f"  人头计数: {len(df_headcount)} 条 → {path_head.name}")
    print(f"  独立职业: {len(df_unique)} 条 → {path_unique.name}")
    return True


# ── 主流程 ──
if __name__ == "__main__":
    print(f"模型: {MODEL_NAME}")
    print(f"数据目录: {DATA_DIR}")
    print(f"结果目录: {RESULT_DIR}")
    print("=" * 60)

    pdfs = discover_pdfs()
    print(f"\n发现 {len(pdfs)} 个PDF文件:")
    for pdf_path, info in pdfs:
        print(f"  {info['output_prefix']}  ←  {pdf_path.name[:50]}...")

    print(f"\n{'=' * 60}")
    print("开始批量识别...\n")

    success = 0
    failed = []
    for idx, (pdf_path, info) in enumerate(pdfs, 1):
        print(f"\n[{idx}/{len(pdfs)}] {info['output_prefix']}")
        try:
            process_one_pdf(pdf_path, info["output_prefix"])
            success += 1
        except Exception as e:
            print(f"  处理失败: {e}")
            failed.append(info["output_prefix"])

    print(f"\n{'=' * 60}")
    print(f"批量识别完成! 成功: {success}, 失败: {len(failed)}")
    if failed:
        print(f"失败列表: {failed}")
