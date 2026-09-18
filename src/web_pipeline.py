import io
import json
import os
import re

import pandas as pd
from openai import OpenAI


# =========================================================
# 1. 获取 Secret
# =========================================================

def get_secret(name: str):
    """
    优先从 Streamlit Secrets 获取。
    如果本地没有，再尝试环境变量。

    这样本地和 Streamlit Cloud 都可以使用。
    """

    try:

        import streamlit as st

        if name in st.secrets:
            return st.secrets[name]

    except Exception:
        pass

    return os.getenv(name)


# =========================================================
# 2. Excel DataFrame → Excel bytes
# =========================================================

def dataframe_to_excel_bytes(
    df: pd.DataFrame,
    sheet_name: str = "Sheet1"
) -> bytes:

    buffer = io.BytesIO()

    with pd.ExcelWriter(
        buffer,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False
        )

    buffer.seek(0)

    return buffer.getvalue()


# =========================================================
# 3. 多 Sheet Excel
# =========================================================

def dataframes_to_excel_bytes(
    sheets: dict[str, pd.DataFrame]
) -> bytes:

    buffer = io.BytesIO()

    with pd.ExcelWriter(
        buffer,
        engine="openpyxl"
    ) as writer:

        for sheet_name, df in sheets.items():

            df.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )

    buffer.seek(0)

    return buffer.getvalue()


# =========================================================
# 4. 读取项目配置
# =========================================================

def load_project_config(base_dir):

    config_dir = base_dir / "config"

    mapping_file = (
        config_dir
        / "mapping_config.xlsx"
    )

    schema_file = (
        config_dir
        / "hris_schema.xlsx"
    )

    if not mapping_file.exists():

        raise FileNotFoundError(
            f"找不到 mapping_config.xlsx：{mapping_file}"
        )

    if not schema_file.exists():

        raise FileNotFoundError(
            f"找不到 hris_schema.xlsx：{schema_file}"
        )

    field_mapping = pd.read_excel(
        mapping_file,
        sheet_name="Field_Mapping"
    )

    value_mapping = pd.read_excel(
        mapping_file,
        sheet_name="Value_Mapping"
    )

    validation_rules = pd.read_excel(
        mapping_file,
        sheet_name="Validation_Rules"
    )

    hris_schema = pd.read_excel(
        schema_file
    )

    return (
        field_mapping,
        value_mapping,
        validation_rules,
        hris_schema
    )


# =========================================================
# 5. 判断敏感字段
# =========================================================

def is_sensitive_field(field_name: str) -> bool:

    field = str(field_name).lower()

    keywords = [
        "姓名",
        "name",
        "email",
        "邮箱",
        "手机",
        "phone",
        "电话",
        "身份证",
        "证件",
        "地址",
        "address",
        "银行卡",
        "银行",
        "id"
    ]

    return any(
        keyword in field
        for keyword in keywords
    )


# =========================================================
# 6. 创建数据画像
# =========================================================

def build_source_profile(
    source_df: pd.DataFrame
) -> dict:

    profile = {}

    for column in source_df.columns:

        series = (
            source_df[column]
            .dropna()
        )

        unique_values = (
            series
            .astype(str)
            .drop_duplicates()
            .tolist()
        )

        profile[column] = {

            "data_type":
                str(source_df[column].dtype),

            "non_null_count":
                int(series.shape[0]),

            "unique_count":
                len(unique_values),

            # 敏感字段不把真实值发送给 AI
            "sample_values":
                []
                if is_sensitive_field(column)
                else unique_values[:10]
        }

    return profile


# =========================================================
# 7. 初始化 DeepSeek Client
# =========================================================

def create_deepseek_client():

    api_key = get_secret(
        "DEEPSEEK_API_KEY"
    )

    if not api_key:

        raise ValueError(
            "没有找到 DEEPSEEK_API_KEY。"
        )

    return OpenAI(
        api_key=api_key,
        base_url="https://api.deepseek.com"
    )


# =========================================================
# 8. DeepSeek 字段映射
# =========================================================

def ai_analyze_mapping(
    source_df: pd.DataFrame,
    field_mapping: pd.DataFrame,
    value_mapping: pd.DataFrame,
    hris_schema: pd.DataFrame
):

    client = create_deepseek_client()

    source_fields = (
        source_df
        .columns
        .tolist()
    )

    source_profile = build_source_profile(
        source_df
    )

    schema_records = (
        hris_schema
        .to_dict(
            orient="records"
        )
    )

    # 正式字段配置
    # 只使用实际存在的列，避免配置表缺少可选字段时直接报错
    preferred_fields = [
        "standard_field",
        "target_field",
        "data_type",
        "required"
    ]

    available_fields = [
        field
        for field in preferred_fields
        if field in field_mapping.columns
    ]

    formal_field_records = (
        field_mapping[
            available_fields
        ]
        .to_dict(
            orient="records"
        )
    )

    # 现有的标准值映射
    formal_value_records = (
        value_mapping
        .to_dict(
            orient="records"
        )
    )

    system_prompt = """
你是一名 HRIS 数据集成专家。

你的任务是分析新的员工 Excel 字段，
并将其映射到已有的标准 HRIS 字段。

严格遵守：

1. 只能根据提供的信息判断。
2. 不确定时不要猜。
3. 不得编造不存在的 HRIS 字段。
4. 不得编造不存在的 HRIS 代码。
5. confidence 使用 high / medium / low。
6. 不确定的映射必须 review_required=true。
7. 不需要人工确认时 review_required=false。
8. 不要输出员工姓名、邮箱、手机号等个人数据。
9. 返回合法 JSON。
"""

    user_prompt = f"""
请分析新的员工数据字段。

新的 Excel 字段：
{json.dumps(
    source_fields,
    ensure_ascii=False,
    indent=2
)}

数据画像：
{json.dumps(
    source_profile,
    ensure_ascii=False,
    indent=2
)}

HRIS Schema：
{json.dumps(
    schema_records,
    ensure_ascii=False,
    indent=2
)}

现有标准字段：
{json.dumps(
    formal_field_records,
    ensure_ascii=False,
    indent=2
)}

已有值映射：
{json.dumps(
    formal_value_records,
    ensure_ascii=False,
    indent=2
)}

请完成：

一、字段映射

将 source_field 映射到：
standard_field
target_field

二、值映射

如果某个字段的值需要转换，
给出：
standard_field
source_value
target_value

严格按照下面结构返回：

{{
    "mappings": [
        {{
            "source_field": "性别",
            "standard_field": "gender",
            "target_field": "gender_code",
            "confidence": "high",
            "transformation": "男→M，女→F",
            "review_required": false,
            "reason": "字段含义一致"
        }}
    ],
    "value_mappings": [
        {{
            "standard_field": "gender",
            "source_value": "男性",
            "target_value": "M",
            "confidence": "high",
            "review_required": false,
            "reason": "与已有标准值含义一致"
        }}
    ]
}}
"""

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

    content = (
        response
        .choices[0]
        .message
        .content
    )

    if not content:

        raise ValueError(
            "DeepSeek 返回为空。"
        )

    result = json.loads(
        content
    )

    field_df = pd.DataFrame(
        result.get(
            "mappings",
            []
        )
    )

    value_df = pd.DataFrame(
        result.get(
            "value_mappings",
            []
        )
    )

    return (
        field_df,
        value_df
    )


# =========================================================
# 9. 根据 AI 审核结果建立“本次运行”的映射
# =========================================================

def build_working_field_mapping(
    formal_mapping: pd.DataFrame,
    approved_ai_mapping: pd.DataFrame
) -> pd.DataFrame:

    working = formal_mapping.copy()

    if approved_ai_mapping.empty:

        return working

    for _, ai_row in (
        approved_ai_mapping
        .iterrows()
    ):

        source_field = str(
            ai_row.get(
                "source_field",
                ""
            )
        ).strip()

        standard_field = str(
            ai_row.get(
                "standard_field",
                ""
            )
        ).strip()

        target_field = str(
            ai_row.get(
                "target_field",
                ""
            )
        ).strip()

        if (
            source_field == ""
            or standard_field == ""
            or target_field == ""
        ):
            continue

        mask = (
            working[
                "standard_field"
            ]
            .astype(str)
            .str.strip()
            == standard_field
        )

        if mask.any():

            working.loc[
                mask,
                "source_field"
            ] = source_field

            working.loc[
                mask,
                "target_field"
            ] = target_field

    return working


# =========================================================
# 10. 根据 AI 审核结果建立“本次值映射”
# =========================================================

def build_working_value_mapping(
    formal_value_mapping: pd.DataFrame,
    approved_ai_values: pd.DataFrame
) -> pd.DataFrame:

    working = (
        formal_value_mapping
        .copy()
    )

    if approved_ai_values.empty:

        return working

    new_rows = []

    for _, row in (
        approved_ai_values
        .iterrows()
    ):

        standard_field = str(
            row.get(
                "standard_field",
                ""
            )
        ).strip()

        source_value = row.get(
            "source_value"
        )

        target_value = row.get(
            "target_value"
        )

        if (
            standard_field == ""
            or pd.isna(source_value)
            or pd.isna(target_value)
        ):

            continue

        new_rows.append({

            "field":
                standard_field,

            "source_value":
                source_value,

            "target_value":
                target_value
        })

    if new_rows:

        new_df = pd.DataFrame(
            new_rows
        )

        working = pd.concat(
            [
                working,
                new_df
            ],
            ignore_index=True
        )

        working = (
            working
            .drop_duplicates(
                subset=[
                    "field",
                    "source_value",
                    "target_value"
                ]
            )
        )

    return working


# =========================================================
# 11. 数据转换
# =========================================================

def transform_employee_data(
    source_df: pd.DataFrame,
    working_mapping: pd.DataFrame,
    working_value_mapping: pd.DataFrame
):

    standard_df = pd.DataFrame(
        index=source_df.index
    )

    # -----------------------------------------
    # 建立 standard data
    # -----------------------------------------

    for _, mapping in (
        working_mapping
        .iterrows()
    ):

        source_field = str(
            mapping[
                "source_field"
            ]
        ).strip()

        standard_field = str(
            mapping[
                "standard_field"
            ]
        ).strip()

        if source_field in source_df.columns:

            standard_df[
                standard_field
            ] = source_df[
                source_field
            ]

        else:

            standard_df[
                standard_field
            ] = pd.NA

    # -----------------------------------------
    # 基础清洗
    # -----------------------------------------

    for column in (
        standard_df.columns
    ):

        standard_df[column] = (
            standard_df[column]
            .astype("string")
            .str.strip()
        )

    # -----------------------------------------
    # Value Mapping
    # -----------------------------------------

    for _, mapping in (
        working_value_mapping
        .iterrows()
    ):

        field = str(
            mapping[
                "field"
            ]
        ).strip()

        source_value = str(
            mapping[
                "source_value"
            ]
        ).strip()

        target_value = mapping[
            "target_value"
        ]

        if field not in standard_df.columns:

            continue

        standard_df[field] = (
            standard_df[field].apply(

                lambda value:
                target_value

                if (
                    not pd.isna(value)
                    and str(value).strip()
                    == source_value
                )

                else value
            )
        )

    return standard_df


# =========================================================
# 12. 数据校验
# =========================================================

def validate_employee_data(
    standard_df: pd.DataFrame,
    validation_rules: pd.DataFrame
):

    errors = []

    bad_indexes = set()

    # -----------------------------------------
    # 基础规则
    # -----------------------------------------

    for index, row in (
        standard_df.iterrows()
    ):

        employee_id = row.get(
            "employee_id",
            ""
        )

        for _, rule in (
            validation_rules
            .iterrows()
        ):

            field = str(
                rule["field"]
            ).strip()

            rule_type = str(
                rule["rule_type"]
            ).strip().lower()

            rule_value = str(
                rule["rule"]
            ).strip()

            error_message = str(
                rule["error_message"]
            ).strip()

            value = row.get(
                field
            )

            # =================================
            # REQUIRED
            # =================================

            if rule_type == "required":

                if (
                    pd.isna(value)
                    or str(value).strip()
                    == ""
                ):

                    errors.append({
                        "excel_row":
                            index + 2,

                        "employee_id":
                            employee_id,

                        "field":
                            field,

                        "error_code":
                            "REQUIRED",

                        "error":
                            error_message,

                        "value":
                            value
                    })

                    bad_indexes.add(index)

            # =================================
            # ENUM
            # =================================

            elif rule_type == "enum":

                if not pd.isna(value):

                    allowed_values = [
                        x.strip()
                        for x in
                        rule_value.split(",")
                    ]

                    if (
                        str(value).strip()
                        not in allowed_values
                    ):

                        errors.append({
                            "excel_row":
                                index + 2,

                            "employee_id":
                                employee_id,

                            "field":
                                field,

                            "error_code":
                                "INVALID_ENUM",

                            "error":
                                error_message,

                            "value":
                                value
                        })

                        bad_indexes.add(index)

            # =================================
            # DATE
            # =================================

            elif rule_type == "date":

                if not pd.isna(value):

                    converted = (
                        pd.to_datetime(
                            value,
                            errors="coerce"
                        )
                    )

                    if pd.isna(
                        converted
                    ):

                        errors.append({
                            "excel_row":
                                index + 2,

                            "employee_id":
                                employee_id,

                            "field":
                                field,

                            "error_code":
                                "INVALID_DATE",

                            "error":
                                error_message,

                            "value":
                                value
                        })

                        bad_indexes.add(index)

                    else:

                        standard_df.at[
                            index,
                            field
                        ] = (
                            converted
                            .strftime(
                                "%Y-%m-%d"
                            )
                        )

            # =================================
            # EMAIL
            # =================================

            elif rule_type == "email":

                if not pd.isna(value):

                    email = str(
                        value
                    ).strip()

                    pattern = (
                        r"^[^@\s]+@[^@\s]+\.[^@\s]+$"
                    )

                    if not re.match(
                        pattern,
                        email
                    ):

                        errors.append({
                            "excel_row":
                                index + 2,

                            "employee_id":
                                employee_id,

                            "field":
                                field,

                            "error_code":
                                "INVALID_EMAIL",

                            "error":
                                error_message,

                            "value":
                                value
                        })

                        bad_indexes.add(index)

    # -----------------------------------------
    # Duplicate employee_id
    # -----------------------------------------

    if "employee_id" in standard_df.columns:

        duplicated = standard_df[
            standard_df[
                "employee_id"
            ]
            .duplicated(
                keep=False
            )
        ]

        for index, row in (
            duplicated.iterrows()
        ):

            errors.append({
                "excel_row":
                    index + 2,

                "employee_id":
                    row[
                        "employee_id"
                    ],

                "field":
                    "employee_id",

                "error_code":
                    "DUPLICATE_ID",

                "error":
                    "员工编号重复",

                "value":
                    row[
                        "employee_id"
                    ]
            })

            bad_indexes.add(index)

    # -----------------------------------------
    # 生成错误报告
    # -----------------------------------------

    error_df = pd.DataFrame(
        errors
    )

    # -----------------------------------------
    # 生成 HRIS 数据
    # -----------------------------------------

    return (
        standard_df,
        error_df,
        bad_indexes
    )


# =========================================================
# 13. Standard → HRIS
# =========================================================

def create_hris_dataframe(
    standard_df: pd.DataFrame,
    working_mapping: pd.DataFrame,
    bad_indexes: set
):

    hris_df = pd.DataFrame(
        index=standard_df.index
    )

    for _, mapping in (
        working_mapping
        .iterrows()
    ):

        standard_field = str(
            mapping[
                "standard_field"
            ]
        ).strip()

        target_field = str(
            mapping[
                "target_field"
            ]
        ).strip()

        if standard_field in standard_df.columns:

            hris_df[
                target_field
            ] = standard_df[
                standard_field
            ]

    # 删除有错误的员工
    valid_indexes = [
        index
        for index in hris_df.index
        if index not in bad_indexes
    ]

    valid_hris_df = (
        hris_df
        .loc[
            valid_indexes
        ]
        .reset_index(
            drop=True
        )
    )

    return valid_hris_df


# =========================================================
# 14. Demo 错误分析
# =========================================================

def demo_error_analysis(
    error_df: pd.DataFrame
) -> pd.DataFrame:

    if error_df.empty:

        return pd.DataFrame(
            columns=[
                "error_code",
                "field",
                "problem",
                "possible_cause",
                "recommended_action",
                "review_required"
            ]
        )

    explanations = {

        "INVALID_ENUM": {
            "problem":
                "字段值不符合 HRIS 允许的枚举值。",

            "cause":
                "原始数据出现未定义的标准值。",

            "action":
                "确认该值对应的 HRIS 标准代码。"
        },

        "DUPLICATE_ID": {
            "problem":
                "员工编号重复。",

            "cause":
                "多个员工记录使用了相同的员工编号。",

            "action":
                "检查原始数据并确认唯一员工编号。"
        },

        "INVALID_EMAIL": {
            "problem":
                "邮箱格式不符合基本格式。",

            "cause":
                "邮箱缺少必要的格式结构。",

            "action":
                "修改为有效工作邮箱。"
        },

        "INVALID_DATE": {
            "problem":
                "日期无法转换成合法日期。",

            "cause":
                "日期格式错误或日期本身不存在。",

            "action":
                "确认并修改入职日期。"
        },

        "REQUIRED": {
            "problem":
                "必填字段为空。",

            "cause":
                "原始数据缺少必要字段。",

            "action":
                "补充字段后重新处理。"
        }
    }

    results = []

    for error_code, group in (
        error_df
        .groupby(
            "error_code"
        )
    ):

        explanation = explanations.get(
            error_code,
            {
                "problem":
                    "检测到数据异常。",

                "cause":
                    "需要进一步检查原始数据。",

                "action":
                    "人工确认并修改。"
            }
        )

        results.append({

            "error_code":
                error_code,

            "field":
                ", ".join(
                    group[
                        "field"
                    ]
                    .astype(str)
                    .unique()
                ),

            "problem":
                explanation[
                    "problem"
                ],

            "possible_cause":
                explanation[
                    "cause"
                ],

            "recommended_action":
                explanation[
                    "action"
                ],

            "review_required":
                True
        })

    return pd.DataFrame(
        results
    )


# =========================================================
# 15. DeepSeek 错误分析
# =========================================================

def ai_analyze_errors(
    error_df: pd.DataFrame,
    hris_schema: pd.DataFrame
):

    client = create_deepseek_client()

    if error_df.empty:

        return demo_error_analysis(
            error_df
        )

    # 不把真实员工值发给 AI
    grouped = (
        error_df
        .groupby(
            [
                "error_code",
                "field",
                "error"
            ]
        )
        .size()
        .reset_index(
            name="affected_rows"
        )
    )

    schema_records = (
        hris_schema
        .to_dict(
            orient="records"
        )
    )

    error_records = (
        grouped
        .to_dict(
            orient="records"
        )
    )

    system_prompt = """
你是一名 HRIS 数据质量分析专家。

请分析员工数据同步错误。

要求：

1. 只能根据提供的信息判断。
2. 不要猜测不存在的系统规则。
3. 不处理个人敏感数据。
4. 无法确定时明确写“需要人工确认”。
5. 返回合法 JSON。
"""

    user_prompt = f"""
错误汇总：

{json.dumps(
    error_records,
    ensure_ascii=False,
    indent=2
)}

HRIS Schema：

{json.dumps(
    schema_records,
    ensure_ascii=False,
    indent=2
)}

请严格返回：

{{
    "analysis": [
        {{
            "error_code": "INVALID_ENUM",
            "field": "gender",
            "problem": "问题说明",
            "possible_cause": "可能原因",
            "recommended_action": "建议处理方式",
            "review_required": true
        }}
    ]
}}
"""

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

    content = (
        response
        .choices[0]
        .message
        .content
    )

    if not content:

        raise ValueError(
            "DeepSeek 返回为空。"
        )

    result = json.loads(
        content
    )

    return pd.DataFrame(
        result.get(
            "analysis",
            []
        )
    )