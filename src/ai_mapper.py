import json
import os
from pathlib import Path

import pandas as pd
from openai import OpenAI


# =====================================
# 1. 项目路径
# =====================================

BASE_DIR = Path(__file__).resolve().parent.parent

SOURCE_FILE = BASE_DIR / "data" / "employees.xlsx"
HRIS_SCHEMA_FILE = BASE_DIR / "config" / "hris_schema.xlsx"
OUTPUT_FILE = BASE_DIR / "output" / "ai_mapping_suggestion.xlsx"


# =====================================
# 2. 加载 API Key
# =====================================

api_key = os.getenv("DEEPSEEK_API_KEY")

if not api_key:
    raise ValueError(
        "没有找到 DEEPSEEK_API_KEY，请检查 Streamlit Secrets。"
    )


# =====================================
# 3. 创建 DeepSeek 客户端
# =====================================

client = OpenAI(
    api_key=api_key,
    base_url="https://api.deepseek.com"
)


# =====================================
# 4. 读取 Excel
# =====================================

if not SOURCE_FILE.exists():
    raise FileNotFoundError(
        f"找不到原始员工文件：{SOURCE_FILE}"
    )

if not HRIS_SCHEMA_FILE.exists():
    raise FileNotFoundError(
        f"找不到 HRIS 字段文件：{HRIS_SCHEMA_FILE}"
    )


source_df = pd.read_excel(SOURCE_FILE)

hris_df = pd.read_excel(HRIS_SCHEMA_FILE)


source_fields = source_df.columns.tolist()

sample_values = {}

for column in source_df.columns:
    values = (
        source_df[column]
        .dropna()
        .astype(str)
        .drop_duplicates()
        .head(20)
        .tolist()
    )

    sample_values[column] = values

hris_schema = hris_df.to_dict(orient="records")

allowed_target_fields = list(dict.fromkeys(
    str(f).strip() for f in hris_df["field_name"].tolist()
))

# 从 schema 提取每个目标字段已定义的枚举代码（data_type 不是类型关键字的即为代码）
TYPE_KEYWORDS = {"string", "enum", "date"}

target_allowed_codes = {}

for _, row in hris_df.iterrows():

    field = str(row["field_name"]).strip()
    code = str(row["data_type"]).strip()

    if code.lower() in TYPE_KEYWORDS:
        continue

    target_allowed_codes.setdefault(field, set()).add(code)


# =====================================
# 5. 构造 AI Prompt
# =====================================

system_prompt = """
你是一名 HRIS 数据集成专家。

你的任务是分析“原始员工 Excel 字段”和“目标 HRIS 字段定义”，
生成字段映射建议。

请遵守以下规则：

1. 只能根据提供的信息进行判断。
2. 不确定时不要猜测。
3. 如果无法确定对应关系，将 review_required 设置为 true。
4. 返回合法 JSON。
5. 每个原始字段都应该给出处理建议。
6. 输出字段包括：
   source_field
   target_field
   confidence
   transformation
   review_required
   reason
7. confidence 使用 high / medium / low。
8. target_field 如果无法确定，填写 null。
9. transformation 用简短文字描述需要进行的转换。
10. review_required 为 true 时，必须说明人工需要检查什么。
11. target_field 必须严格从“目标 HRIS 字段定义”的 field_name 中取值，
    禁止自行创造、缩写或替换字段名（例如不要用 employment_status 代替 worker_status）。
"""


user_prompt = f"""
请分析以下 HRIS 字段映射，并进一步分析字段值转换。

原始 Excel 字段：
{json.dumps(source_fields, ensure_ascii=False, indent=2)}

原始字段样例值：
{json.dumps(sample_values, ensure_ascii=False, indent=2)}

目标 HRIS 字段定义：
{json.dumps(hris_schema, ensure_ascii=False, indent=2)}

target_field 白名单（只能从以下值中选择，禁止使用其他任何名称）：
{json.dumps(allowed_target_fields, ensure_ascii=False, indent=2)}

请完成两项任务：

第一项：字段映射
判断原始字段对应哪个 HRIS 字段。

第二项：值映射
如果原始字段的值需要转换成 HRIS 代码，
请给出 source_value → target_value 的建议。

特别注意：

1. 不确定时不要猜。
2. 不确定的字段或值设置 review_required=true。
3. 不要虚构不存在的 HRIS 代码。
4. 返回合法 JSON。
5. 所有 AI 推测都必须给出 reason。
6. target_field 必须严格等于上面白名单中的某个值，禁止创造新名称。

严格按照下面的 JSON 结构返回：

{{
    "mappings": [
        {{
            "source_field": "性别",
            "target_field": "gender",
            "confidence": "high",
            "transformation": "值映射",
            "review_required": false,
            "reason": "性别字段与 HRIS gender 含义一致"
        }}
    ],
    "value_mappings": [
        {{
            "field": "性别",
            "source_value": "男",
            "target_value": "M",
            "confidence": "high",
            "review_required": false,
            "reason": "HRIS 明确要求 M/F"
        }}
    ]
}}
"""


# =====================================
# 6. 调用 DeepSeek
# =====================================

print("正在调用 DeepSeek API...")

response = client.chat.completions.create(
    model="deepseek-flash",
    messages=[
        {
            "role": "system",
            "content": system_prompt
        },
        {
            "role": "user",
            "content": user_prompt
        }
    ],
    response_format={
        "type": "json_object"
    },
    stream=False
)


# =====================================
# 7. 获取 AI 返回结果
# =====================================

content = response.choices[0].message.content

if not content:
    raise ValueError("DeepSeek 返回了空内容。")

result = json.loads(content)


# =====================================
# 8. 检查结果
# =====================================

if "mappings" not in result:
    raise ValueError(
        "AI 返回结果中没有 mappings 字段。"
    )


mapping_df = pd.DataFrame(result["mappings"])


# 校验 target_field 是否都在 schema 白名单内，越界则标记为需人工复核
invalid_targets = []

if not mapping_df.empty and "target_field" in mapping_df.columns:

    for idx, target in mapping_df["target_field"].items():

        if target is None or pd.isna(target):
            continue

        if str(target).strip() not in allowed_target_fields:

            mapping_df.at[idx, "review_required"] = True

            note = "（target_field 不在 schema 白名单中，需人工核对）"

            current_reason = mapping_df.at[idx, "reason"]

            if pd.isna(current_reason):
                mapping_df.at[idx, "reason"] = note
            else:
                mapping_df.at[idx, "reason"] = str(current_reason) + note

            invalid_targets.append(str(target).strip())

if invalid_targets:
    print(
        f"警告：AI 返回了不在 schema 中的 target_field："
        f"{invalid_targets}，已标记 review_required=true"
    )


# =====================================
# 9. 保存 AI 映射建议
# =====================================

OUTPUT_FILE.parent.mkdir(exist_ok=True)

value_mapping_df = pd.DataFrame(
    result.get("value_mappings", [])
)

# 构建 source_field -> target_field 映射（用于值代码校验）
source_to_target = {}

if (
    not mapping_df.empty
    and "source_field" in mapping_df.columns
    and "target_field" in mapping_df.columns
):

    for _, row in mapping_df.iterrows():

        src = row["source_field"]
        tgt = row["target_field"]

        if pd.isna(src) or pd.isna(tgt):
            continue

        source_to_target[str(src).strip()] = str(tgt).strip()

# 校验 value_mappings 的 target_value 是否在 schema 已定义代码内
invalid_codes = []

if not value_mapping_df.empty:

    for idx, row in value_mapping_df.iterrows():

        source_field = row.get("field")
        target_value = row.get("target_value")

        if source_field is None or pd.isna(source_field):
            continue

        target_field = source_to_target.get(str(source_field).strip())

        if target_field is None or target_field not in target_allowed_codes:
            continue

        if target_value is None or pd.isna(target_value):
            continue

        if str(target_value).strip() not in target_allowed_codes[target_field]:

            value_mapping_df.at[idx, "target_value"] = None
            value_mapping_df.at[idx, "review_required"] = True

            note = "（target_value 不在 schema 已定义代码中，已清空，需人工确认）"

            current_reason = value_mapping_df.at[idx, "reason"]

            if pd.isna(current_reason):
                value_mapping_df.at[idx, "reason"] = note
            else:
                value_mapping_df.at[idx, "reason"] = str(current_reason) + note

            invalid_codes.append(
                f"{source_field} -> {target_value}"
            )

if invalid_codes:
    print(
        f"警告：AI 返回了 schema 未定义的代码："
        f"{invalid_codes}，已清空并标记 review_required=true"
    )

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:

    mapping_df.to_excel(
        writer,
        sheet_name="Field_Mapping",
        index=False
    )

    value_mapping_df.to_excel(
        writer,
        sheet_name="Value_Mapping",
        index=False
    )


# =====================================
# 10. 输出结果
# =====================================

print("=" * 60)
print("AI 字段映射分析完成")
print("=" * 60)

print(f"原始字段数量：{len(source_fields)}")
print(f"AI 生成映射数量：{len(mapping_df)}")

print()
print("AI 映射结果：")
print(mapping_df.to_string(index=False))

print()
print(f"结果文件：{OUTPUT_FILE}")

print("=" * 60)
