import re
from pathlib import Path

import pandas as pd


# =========================================
# 1. 项目路径
# =========================================

BASE_DIR = Path(__file__).resolve().parent.parent

INPUT_FILE = BASE_DIR / "data" / "employees.xlsx"
CONFIG_FILE = BASE_DIR / "config" / "mapping_config.xlsx"

OUTPUT_DIR = BASE_DIR / "output"

OUTPUT_FILE = OUTPUT_DIR / "hris_import.xlsx"
ERROR_FILE = OUTPUT_DIR / "sync_errors.xlsx"


# =========================================
# 2. 创建输出文件夹
# =========================================

OUTPUT_DIR.mkdir(exist_ok=True)


# =========================================
# 3. 检查文件是否存在
# =========================================

if not INPUT_FILE.exists():
    raise FileNotFoundError(
        f"找不到员工文件：{INPUT_FILE}"
    )

if not CONFIG_FILE.exists():
    raise FileNotFoundError(
        f"找不到配置文件：{CONFIG_FILE}"
    )


# =========================================
# 4. 读取 Excel
# =========================================

employees = pd.read_excel(INPUT_FILE)

field_mapping = pd.read_excel(
    CONFIG_FILE,
    sheet_name="Field_Mapping"
)

value_mapping = pd.read_excel(
    CONFIG_FILE,
    sheet_name="Value_Mapping"
)

validation_rules = pd.read_excel(
    CONFIG_FILE,
    sheet_name="Validation_Rules"
)


# =========================================
# 5. 检查配置文件字段
# =========================================

required_mapping_columns = [
    "source_field",
    "target_field"
]

for column in required_mapping_columns:

    if column not in field_mapping.columns:
        raise ValueError(
            f"Field_Mapping 缺少字段：{column}"
        )

required_value_columns = [
    "field",
    "source_value",
    "target_value"
]

for column in required_value_columns:

    if column not in value_mapping.columns:
        raise ValueError(
            f"Value_Mapping 缺少字段：{column}"
        )

required_validation_columns = [
    "field",
    "rule_type",
    "rule",
    "error_message"
]

for column in required_validation_columns:

    if column not in validation_rules.columns:
        raise ValueError(
            f"Validation_Rules 缺少字段：{column}"
        )


# =========================================
# 6. 建立目标数据（source_field -> target_field）
# =========================================

target_df = pd.DataFrame(
    index=employees.index
)

source_to_target = {}

for _, mapping in field_mapping.iterrows():

    source_field = str(mapping["source_field"]).strip()
    target_field = str(mapping["target_field"]).strip()

    if source_field not in employees.columns:

        raise ValueError(
            f"原始 Excel 中找不到字段：{source_field}"
        )

    target_df[target_field] = employees[source_field]
    source_to_target[source_field] = target_field


# =========================================
# 7. 基础清洗
# =========================================

for column in target_df.columns:

    target_df[column] = (
        target_df[column]
        .astype("string")
        .str.strip()
    )


# =========================================
# 8. 执行 Value Mapping（作用于源字段）
# =========================================

for _, mapping in value_mapping.iterrows():

    source_field = str(mapping["field"]).strip()

    target_field = source_to_target.get(source_field)

    if target_field is None or target_field not in target_df.columns:
        continue

    source_value = str(
        mapping["source_value"]
    ).strip()

    target_value = mapping["target_value"]

    # 目标值为空表示“无有效代码”，保持原值交给校验规则处理，
    # 避免把非法值静默清空导致校验漏报。
    if pd.isna(target_value):
        continue

    target_df[target_field] = target_df[target_field].apply(
        lambda value:
        target_value
        if not pd.isna(value)
        and str(value).strip() == source_value
        else value
    )


# =========================================
# 9. 确定员工唯一标识字段
# =========================================

id_field = None

for _, rule in validation_rules.iterrows():

    if str(rule.get("rule_type", "")).strip().lower() == "unique":

        id_field = str(rule["field"]).strip()
        break

if id_field is None and "worker_id" in target_df.columns:
    id_field = "worker_id"


# =========================================
# 10. 创建错误列表
# =========================================

errors = []


def add_error(
    index,
    employee_id,
    field,
    error_code,
    error_message,
    value
):

    errors.append({
        "excel_row": index + 2,
        "employee_id": employee_id,
        "field": field,
        "error_code": error_code,
        "error": error_message,
        "value": value
    })


def get_employee_id(index):

    if id_field is None:
        return ""

    return target_df.at[index, id_field]


# =========================================
# 11. 数据校验：required
# =========================================

required_fields = []

required_message = {}

for _, rule in validation_rules.iterrows():

    field = str(rule["field"]).strip()

    rule_type = str(
        rule.get("rule_type", "")
    ).strip().lower()

    if rule_type == "required":

        required_message[field] = str(
            rule.get("error_message", "")
        ).strip()

    is_required = (
        str(rule.get("required", "")).strip().upper() == "YES"
    )

    if is_required and field not in required_fields:
        required_fields.append(field)

for field in required_message:

    if field not in required_fields:
        required_fields.append(field)

for field in required_fields:

    if field not in target_df.columns:
        continue

    message = required_message.get(
        field,
        f"{field} 不能为空"
    )

    for index, value in target_df[field].items():

        if pd.isna(value) or str(value).strip() == "":

            add_error(
                index,
                get_employee_id(index),
                field,
                "REQUIRED",
                message,
                value
            )


# =========================================
# 12. 数据校验：rule_type
# =========================================

for _, rule in validation_rules.iterrows():

    field = str(rule["field"]).strip()

    rule_type = str(
        rule.get("rule_type", "")
    ).strip().lower()

    rule_value = str(
        rule.get("rule", "")
    ).strip()

    error_message = str(
        rule.get("error_message", "")
    ).strip()

    if field not in target_df.columns:
        continue

    # ---------------------------------
    # unique
    # ---------------------------------

    if rule_type == "unique":

        duplicated = target_df[
            target_df[field].duplicated(keep=False)
        ]

        for index, value in duplicated[field].items():

            add_error(
                index,
                get_employee_id(index),
                field,
                "DUPLICATE_ID",
                error_message,
                value
            )

    # ---------------------------------
    # enum
    # ---------------------------------

    elif rule_type == "enum":

        allowed_values = [
            x.strip()
            for x in rule_value.replace("，", ",").split(",")
            if x.strip()
        ]

        for index, value in target_df[field].items():

            if pd.isna(value):
                continue

            if str(value).strip() not in allowed_values:

                add_error(
                    index,
                    get_employee_id(index),
                    field,
                    "INVALID_ENUM",
                    error_message,
                    value
                )

    # ---------------------------------
    # date
    # ---------------------------------

    elif rule_type == "date":

        for index, value in target_df[field].items():

            if pd.isna(value) or str(value).strip() == "":
                continue

            converted_date = pd.to_datetime(
                value,
                errors="coerce"
            )

            if pd.isna(converted_date):

                add_error(
                    index,
                    get_employee_id(index),
                    field,
                    "INVALID_DATE",
                    error_message,
                    value
                )

            else:

                target_df.at[index, field] = (
                    converted_date.strftime("%Y-%m-%d")
                )

    # ---------------------------------
    # email
    # ---------------------------------

    elif rule_type == "email":

        pattern = r"^[^@\s]+@[^@\s]+\.[^@\s]+$"

        for index, value in target_df[field].items():

            if pd.isna(value) or str(value).strip() == "":
                continue

            if not re.match(pattern, str(value).strip()):

                add_error(
                    index,
                    get_employee_id(index),
                    field,
                    "INVALID_EMAIL",
                    error_message,
                    value
                )

    # ---------------------------------
    # required 已在第 11 步处理
    # ---------------------------------

    elif rule_type == "required":
        continue


# =========================================
# 13. 找出有错误的员工
# =========================================

error_employee_ids = {
    str(error["employee_id"])
    for error in errors
    if (
        str(error["employee_id"]).strip() != ""
        and str(error["employee_id"]).lower() != "nan"
        and str(error["employee_id"]).lower() != "<na>"
    )
}


# =========================================
# 14. 只保留没有错误的员工
# =========================================

if id_field is not None and id_field in target_df.columns:

    valid_df = target_df[
        ~target_df[id_field]
        .astype(str)
        .isin(error_employee_ids)
    ].copy()

else:

    valid_df = target_df.copy()


# =========================================
# 15. 保存 HRIS 文件
# =========================================

valid_df.to_excel(
    OUTPUT_FILE,
    index=False
)


# =========================================
# 16. 保存错误报告
# =========================================

if errors:

    error_df = pd.DataFrame(errors)

else:

    error_df = pd.DataFrame(
        columns=[
            "excel_row",
            "employee_id",
            "field",
            "error_code",
            "error",
            "value"
        ]
    )


error_df.to_excel(
    ERROR_FILE,
    index=False
)


# =========================================
# 17. 打印运行结果
# =========================================

print()
print("=" * 60)

print("HRIS 数据处理完成")

print("=" * 60)

print(
    f"原始员工数量：{len(employees)}"
)

print(
    f"成功员工数量：{len(valid_df)}"
)

print(
    f"错误记录数量：{len(errors)}"
)

print()
print(
    f"HRIS 导入文件：{OUTPUT_FILE}"
)

print(
    f"错误报告：{ERROR_FILE}"
)

print("=" * 60)
