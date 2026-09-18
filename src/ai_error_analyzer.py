import json
import os
import subprocess
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from dotenv import load_dotenv
from openai import OpenAI


# =========================================
# 1. 项目路径
# =========================================

BASE_DIR = Path(__file__).resolve().parent.parent

ERROR_FILE = (
    BASE_DIR
    / "output"
    / "sync_errors.xlsx"
)

SCHEMA_FILE = (
    BASE_DIR
    / "config"
    / "hris_schema.xlsx"
)

CONFIG_FILE = (
    BASE_DIR
    / "config"
    / "mapping_config.xlsx"
)

OUTPUT_FILE = (
    BASE_DIR
    / "output"
    / "ai_error_analysis.xlsx"
)

MAIN_SCRIPT = (
    BASE_DIR
    / "src"
    / "main.py"
)


# =========================================
# 2. 加载 API Key
# =========================================

load_dotenv(
    BASE_DIR / ".env"
)

api_key = os.getenv(
    "DEEPSEEK_API_KEY"
)

if not api_key:

    raise ValueError(
        "没有找到 DEEPSEEK_API_KEY"
    )


# =========================================
# 3. 创建 DeepSeek 客户端
# =========================================

client = OpenAI(
    api_key=api_key,
    base_url="https://api.deepseek.com"
)


# =========================================
# 4. 检查文件
# =========================================

for path in (ERROR_FILE, SCHEMA_FILE, CONFIG_FILE):

    if not path.exists():

        raise FileNotFoundError(
            f"找不到文件：{path}"
        )


# =========================================
# 5. 读取数据
# =========================================

error_df = pd.read_excel(
    ERROR_FILE
)

schema_df = pd.read_excel(
    SCHEMA_FILE
)

config_sheets = pd.read_excel(
    CONFIG_FILE,
    sheet_name=None
)

field_mapping = config_sheets["Field_Mapping"]
value_mapping = config_sheets["Value_Mapping"]
validation_rules = config_sheets["Validation_Rules"]


# =========================================
# 6. 提取目标字段合法代码
# =========================================

TYPE_KEYWORDS = {"string", "enum", "date"}

target_allowed_codes = {}

for _, row in schema_df.iterrows():

    field = str(row["field_name"]).strip()
    code = str(row["data_type"]).strip()

    if code.lower() in TYPE_KEYWORDS:
        continue

    target_allowed_codes.setdefault(field, []).append(code)

for field in target_allowed_codes:

    target_allowed_codes[field] = list(
        dict.fromkeys(target_allowed_codes[field])
    )


# =========================================
# 7. 如果没有错误
# =========================================

if error_df.empty:

    print(
        "当前没有发现错误，不需要进行 AI 分析。"
    )

    empty_analysis = pd.DataFrame(
        columns=[
            "error_code",
            "field",
            "problem",
            "possible_cause",
            "recommended_action",
            "review_required"
        ]
    )

    empty_fixes = pd.DataFrame(
        columns=[
            "error_code",
            "field",
            "value",
            "action",
            "suggested_value",
            "applied",
            "confidence",
            "reason"
        ]
    )

    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl"
    ) as writer:

        empty_analysis.to_excel(
            writer,
            sheet_name="Analysis",
            index=False
        )

        empty_fixes.to_excel(
            writer,
            sheet_name="Fix_Actions",
            index=False
        )

    raise SystemExit


# =========================================
# 8. 整理错误数据
# =========================================

error_summary = (
    error_df
    .groupby(
        ["error_code", "field", "error"],
        dropna=False
    )
    .agg(
        sample_values=(
            "value",
            lambda values:
            list(
                pd.Series(values)
                .dropna()
                .astype(str)
                .drop_duplicates()
                .head(5)
            )
        ),
        affected_rows=(
            "excel_row",
            "count"
        )
    )
    .reset_index()
)


# =========================================
# 9. JSON 安全的序列化辅助
# =========================================

def to_native(obj):

    if isinstance(obj, dict):
        return {
            k: to_native(v)
            for k, v in obj.items()
        }

    if isinstance(obj, (list, tuple)):
        return [
            to_native(x)
            for x in obj
        ]

    if isinstance(obj, np.generic):
        return obj.item()

    if isinstance(obj, (pd.Timestamp, pd.Timedelta)):
        return str(obj)

    if obj is None:
        return None

    try:
        if pd.isna(obj):
            return None
    except (TypeError, ValueError):
        pass

    return obj


error_records = to_native(
    error_summary.to_dict(orient="records")
)

schema_records = to_native(
    schema_df.to_dict(orient="records")
)

field_mapping_records = to_native(
    field_mapping[["source_field", "target_field"]]
    .to_dict(orient="records")
)

value_mapping_records = to_native(
    value_mapping[["field", "source_value", "target_value"]]
    .to_dict(orient="records")
)

validation_rules_records = to_native(
    validation_rules[["field", "rule_type", "rule"]]
    .to_dict(orient="records")
)


# =========================================
# 10. Prompt
# =========================================

system_prompt = """
你是一名 HRIS 数据质量分析专家。

你的任务是分析员工数据同步过程中产生的错误，并给出可自动执行的修复方案。

注意：

1. 不要编造 HRIS 不存在的规则。
2. 只能根据提供的信息进行判断。
3. 如果无法确定原因，明确写“需要人工确认”。
4. 不要处理或推测身份证号、手机号、住址等敏感个人信息。
5. 重点分析字段、错误类型、样例值和影响行数。
6. 返回合法 JSON。

修复方案 action 只能是以下三种之一：

- add_enum_code：错误值本身就是目标字段的合法代码，只是校验规则枚举遗漏了它。
  此时 suggested_value 填该合法代码。
- add_value_mapping：错误值是源字段的原始值，可以映射到某个已存在的合法代码。
  此时 source_value 填原始值，suggested_value 填合法代码。
- needs_manual_review：无法安全自动修复（需要人工决定取值、修改源数据、处理重复编号等）。

suggested_value 必须严格来自“目标字段合法代码”，禁止虚构代码。
不确定时使用 needs_manual_review。
"""


user_prompt = f"""
请分析以下 HRIS 数据同步错误，并给出修复方案。

错误汇总：

{json.dumps(
    error_records,
    ensure_ascii=False,
    indent=2
)}

HRIS 字段定义：

{json.dumps(
    schema_records,
    ensure_ascii=False,
    indent=2
)}

目标字段合法代码：

{json.dumps(
    target_allowed_codes,
    ensure_ascii=False,
    indent=2
)}

当前字段映射（source_field -> target_field）：

{json.dumps(
    field_mapping_records,
    ensure_ascii=False,
    indent=2
)}

当前值映射（field / source_value -> target_value）：

{json.dumps(
    value_mapping_records,
    ensure_ascii=False,
    indent=2
)}

当前校验规则（field / rule_type / rule）：

{json.dumps(
    validation_rules_records,
    ensure_ascii=False,
    indent=2
)}

请完成两项任务：

第一项：错误分析（analysis）
针对每一种错误给出：
- error_code
- field
- problem
- possible_cause
- recommended_action
- review_required

第二项：修复方案（fixes）
针对每一种错误给出可自动执行的修复动作，字段包括：
- error_code
- field
- value
- action
- suggested_value
- source_value
- confidence
- reason

请严格返回以下 JSON 格式：

{{
    "analysis": [
        {{
            "error_code": "INVALID_ENUM",
            "field": "gender",
            "problem": "原始值无法匹配 HRIS 允许的枚举值",
            "possible_cause": "原始数据中出现了未定义的性别值",
            "recommended_action": "人工确认该值对应的 HRIS 标准代码",
            "review_required": true
        }}
    ],
    "fixes": [
        {{
            "error_code": "INVALID_ENUM",
            "field": "department",
            "value": "RD",
            "action": "add_enum_code",
            "suggested_value": "RD",
            "source_value": null,
            "confidence": "high",
            "reason": "研发部在 HRIS schema 中对应 RD，但 department 枚举遗漏了 RD"
        }}
    ]
}}
"""


# =========================================
# 11. 调用 DeepSeek
# =========================================

print(
    "正在调用 DeepSeek 分析错误..."
)


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


# =========================================
# 12. 解析 JSON
# =========================================

content = (
    response
    .choices[0]
    .message
    .content
)

if not content:

    raise ValueError(
        "DeepSeek 返回了空内容"
    )


result = json.loads(
    content
)


if "analysis" not in result:

    raise ValueError(
        "AI 返回结果缺少 analysis"
    )


analysis_df = pd.DataFrame(
    result["analysis"]
)

fixes = result.get("fixes", [])


# =========================================
# 13. 应用安全修复
# =========================================

applied_count = 0

fix_rows = []

for fix in fixes:

    error_code = str(fix.get("error_code", "")).strip()
    field = str(fix.get("field", "")).strip()
    value = fix.get("value")
    action = str(fix.get("action", "")).strip()
    suggested = fix.get("suggested_value")
    source_value = fix.get("source_value")
    confidence = str(fix.get("confidence", "")).strip()
    reason = str(fix.get("reason", "")).strip()

    applied = False

    # ---------------------------------
    # 补枚举代码
    # ---------------------------------

    if action == "add_enum_code" and suggested is not None:

        code = str(suggested).strip()
        allowed = target_allowed_codes.get(field, [])

        if code in allowed:

            mask = (
                validation_rules["field"]
                .astype(str)
                .str.strip() == field
            ) & (
                validation_rules["rule_type"]
                .astype(str)
                .str.strip()
                .str.lower() == "enum"
            )

            if mask.any():

                idx = validation_rules.index[mask][0]

                current_rule = str(
                    validation_rules.at[idx, "rule"]
                ).strip()

                current_codes = [
                    x.strip()
                    for x in current_rule.replace("，", ",").split(",")
                    if x.strip()
                ]

                if code not in current_codes:

                    current_codes.append(code)

                    validation_rules.at[idx, "rule"] = ",".join(
                        current_codes
                    )

                applied = True

    # ---------------------------------
    # 补值映射
    # ---------------------------------

    elif action == "add_value_mapping" and suggested is not None:

        code = str(suggested).strip()

        raw_source = (
            source_value
            if source_value is not None
            else value
        )

        raw_source = (
            ""
            if raw_source is None
            else str(raw_source).strip()
        )

        allowed = target_allowed_codes.get(field, [])

        if code in allowed and raw_source != "":

            source_field = None

            for _, m in field_mapping.iterrows():

                if str(m["target_field"]).strip() == field:

                    source_field = str(
                        m["source_field"]
                    ).strip()

                    break

            if source_field is not None:

                new_row = {
                    "field": source_field,
                    "source_value": raw_source,
                    "target_value": code,
                    "confidence": confidence,
                    "review_required": False,
                    "reason": reason
                }

                value_mapping = pd.concat(
                    [
                        value_mapping,
                        pd.DataFrame([new_row])
                    ],
                    ignore_index=True
                )

                config_sheets["Value_Mapping"] = value_mapping

                applied = True

    if applied:
        applied_count += 1

    fix_rows.append({
        "error_code": error_code,
        "field": field,
        "value": value,
        "action": action,
        "suggested_value": suggested,
        "applied": applied,
        "confidence": confidence,
        "reason": reason
    })


fix_df = pd.DataFrame(fix_rows)


# =========================================
# 14. 写回配置文件
# =========================================

with pd.ExcelWriter(
    CONFIG_FILE,
    engine="openpyxl"
) as writer:

    for sheet_name, df in config_sheets.items():

        df.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False
        )


# =========================================
# 15. 重新运行 main.py
# =========================================

print()
print("正在重新运行 main.py ...")
print()

env = os.environ.copy()
env["PYTHONIOENCODING"] = "utf-8"

rerun = subprocess.run(
    [sys.executable, str(MAIN_SCRIPT)],
    cwd=str(BASE_DIR),
    capture_output=True,
    text=True,
    encoding="utf-8",
    errors="replace",
    env=env
)

print(rerun.stdout)

if rerun.returncode != 0:

    print(rerun.stderr)

    raise SystemExit(
        f"main.py 运行失败（退出码 {rerun.returncode}）"
    )


# =========================================
# 16. 保存结果
# =========================================

OUTPUT_FILE.parent.mkdir(exist_ok=True)

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    analysis_df.to_excel(
        writer,
        sheet_name="Analysis",
        index=False
    )

    fix_df.to_excel(
        writer,
        sheet_name="Fix_Actions",
        index=False
    )


# =========================================
# 17. 打印结果
# =========================================

print()
print("=" * 60)

print(
    "AI 错误分析完成"
)

print("=" * 60)

print(
    f"分析的错误类型：{len(analysis_df)}"
)

print(
    f"自动修复：{applied_count}"
)

print(
    f"需人工处理：{len(fix_df) - applied_count}"
)

print()
print(
    f"结果文件：{OUTPUT_FILE}"
)

print("=" * 60)
