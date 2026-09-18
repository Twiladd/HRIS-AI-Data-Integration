import json
import sys
from pathlib import Path

import pandas as pd
import requests

sys.stdout.reconfigure(encoding="utf-8")


# =========================================
# 1. 项目路径
# =========================================

BASE_DIR = Path(__file__).resolve().parent.parent

HRIS_FILE = (
    BASE_DIR
    / "output"
    / "hris_import.xlsx"
)

OUTPUT_FILE = (
    BASE_DIR
    / "output"
    / "hris_sync_result.xlsx"
)


# =========================================
# 2. HRIS API 地址
# =========================================

HRIS_API_URL = (
    "http://127.0.0.1:8000/employees"
)


# =========================================
# 3. 检查输入文件
# =========================================

if not HRIS_FILE.exists():

    raise FileNotFoundError(
        f"找不到 HRIS 文件：{HRIS_FILE}"
    )


# =========================================
# 4. 读取 HRIS 数据
# =========================================

df = pd.read_excel(
    HRIS_FILE
)


# =========================================
# 5. 创建同步结果
# =========================================

results = []


# =========================================
# 6. 一个员工一个员工发送
# =========================================

for _, row in df.iterrows():

    employee = {

        "employee_id":
            str(row["worker_id"]),

        "full_name":
            str(row["legal_name"]),

        "gender_code":
            str(row["gender"]),

        "department_code":
            str(row["department"]),

        "hire_date":
            str(row["start_date"]),

        "work_email":
            None
            if pd.isna(row["business_email"])
            else str(row["business_email"]),

        "employment_status":
            str(row["employment_status"])
    }


    print(
        f"正在同步：{employee['employee_id']}"
    )


    try:

        response = requests.post(
            HRIS_API_URL,
            json=employee,
            timeout=10
        )


        # =====================================
        # 同步成功
        # =====================================

        if response.status_code == 200:

            result = response.json()

            results.append({
                "employee_id":
                    employee["employee_id"],

                "status":
                    "SUCCESS",

                "http_status":
                    response.status_code,

                "message":
                    result.get(
                        "message",
                        ""
                    )
            })


            print(
                f"✓ 同步成功："
                f"{employee['employee_id']}"
            )


        # =====================================
        # 同步失败
        # =====================================

        else:

            try:

                error_detail = (
                    response.json()
                    .get(
                        "detail",
                        response.text
                    )
                )

            except Exception:

                error_detail = response.text


            results.append({

                "employee_id":
                    employee["employee_id"],

                "status":
                    "FAILED",

                "http_status":
                    response.status_code,

                "message":
                    str(error_detail)
            })


            print(
                f"✗ 同步失败："
                f"{employee['employee_id']}"
            )


    # =========================================
    # 网络错误
    # =========================================

    except requests.RequestException as e:

        results.append({

            "employee_id":
                employee["employee_id"],

            "status":
                "ERROR",

            "http_status":
                None,

            "message":
                str(e)
        })


        print(
            f"✗ 网络错误："
            f"{employee['employee_id']}"
        )


# =========================================
# 7. 保存同步结果
# =========================================

result_df = pd.DataFrame(
    results
)


result_df.to_excel(
    OUTPUT_FILE,
    index=False
)


# =========================================
# 8. 统计
# =========================================

success_count = len(
    result_df[
        result_df["status"] == "SUCCESS"
    ]
)

failed_count = len(
    result_df[
        result_df["status"] != "SUCCESS"
    ]
)


print()
print("=" * 60)

print("HRIS API 同步完成")

print("=" * 60)

print(
    f"总数据：{len(df)}"
)

print(
    f"同步成功：{success_count}"
)

print(
    f"同步失败：{failed_count}"
)

print()
print(
    f"同步结果：{OUTPUT_FILE}"
)

print("=" * 60)