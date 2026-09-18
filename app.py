import os
import subprocess
import sys
import shutil
from pathlib import Path
from io import BytesIO

import pandas as pd
import streamlit as st


# =========================================================
# 1. 项目路径
# =========================================================

BASE_DIR = Path(__file__).resolve().parent

DATA_DIR = BASE_DIR / "data"
CONFIG_DIR = BASE_DIR / "config"
OUTPUT_DIR = BASE_DIR / "output"

INPUT_FILE = DATA_DIR / "employees.xlsx"
DEMO_FILE = DATA_DIR / "demo_employees.xlsx"

MAPPING_FILE = CONFIG_DIR / "mapping_config.xlsx"

AI_MAPPING_FILE = OUTPUT_DIR / "ai_mapping_suggestion.xlsx"
HRIS_FILE = OUTPUT_DIR / "hris_import.xlsx"
ERROR_FILE = OUTPUT_DIR / "sync_errors.xlsx"
AI_ERROR_FILE = OUTPUT_DIR / "ai_error_analysis.xlsx"


# =========================================================
# 2. 页面设置
# =========================================================

st.set_page_config(
    page_title="HRIS AI Data Sync",
    page_icon="👥",
    layout="wide",
    initial_sidebar_state="expanded"
)


# =========================================================
# 3. 简单页面样式
# =========================================================

st.markdown(
    """
    <style>

    .main-title {
        font-size: 36px;
        font-weight: 700;
        margin-bottom: 5px;
    }

    .subtitle {
        color: #6b7280;
        font-size: 16px;
        margin-bottom: 25px;
    }

    .step-card {
        padding: 15px;
        border-radius: 10px;
        border: 1px solid #e5e7eb;
        background-color: #fafafa;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# =========================================================
# 4. 工具函数
# =========================================================

def normalize_df(df):
    """
    将 object 列统一转成 string，避免 Arrow 序列化失败。
    """

    for column in df.columns:
        if df[column].dtype == object:
            df[column] = df[column].astype("string")

    return df


def run_python_script(script_name):
    """
    执行项目中的 Python 脚本。
    """

    script_path = BASE_DIR / "src" / script_name

    if not script_path.exists():

        return False, "", f"找不到脚本：{script_path}"

    result = subprocess.run(
        [
            sys.executable,
            str(script_path)
        ],
        cwd=BASE_DIR,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env={
            **os.environ,
            "PYTHONUTF8": "1"
        }
    )

    return (
        result.returncode == 0,
        result.stdout,
        result.stderr
    )


def save_uploaded_file(uploaded_file):
    """
    将网页上传的文件保存成项目中的 employees.xlsx。
    """

    DATA_DIR.mkdir(exist_ok=True)

    with open(INPUT_FILE, "wb") as file:

        file.write(
            uploaded_file.getvalue()
        )


def to_bool(value):
    """
    将 Excel / AI 返回的 TRUE/FALSE 转换成 Python bool。
    """

    if isinstance(value, bool):
        return value

    return str(value).strip().lower() in {
        "true",
        "yes",
        "1",
        "是"
    }


def excel_bytes(file_path):
    """
    将 Excel 文件转换成可下载的 bytes。
    """

    if not file_path.exists():
        return None

    return file_path.read_bytes()


def load_ai_mapping():

    if not AI_MAPPING_FILE.exists():
        return None, None

    field_mapping = pd.read_excel(
        AI_MAPPING_FILE,
        sheet_name="Field_Mapping"
    )

    value_mapping = pd.read_excel(
        AI_MAPPING_FILE,
        sheet_name="Value_Mapping"
    )

    return field_mapping, value_mapping


def apply_approved_mapping(
    approved_field_mapping,
    approved_value_mapping
):
    """
    将人工确认后的 AI 映射写入正式 mapping_config.xlsx。

    Field_Mapping 的字段是 source_field → target_field，
    Value_Mapping 的 field 是源字段名（main.py 通过
    source_to_target 查找），因此直接采用 AI 返回的字段名即可。
    """

    if not MAPPING_FILE.exists():

        raise FileNotFoundError(
            f"找不到正式配置文件：{MAPPING_FILE}"
        )

    # 读取正式配置
    formal_field_mapping = pd.read_excel(
        MAPPING_FILE,
        sheet_name="Field_Mapping"
    )

    formal_value_mapping = pd.read_excel(
        MAPPING_FILE,
        sheet_name="Value_Mapping"
    )

    validation_rules = pd.read_excel(
        MAPPING_FILE,
        sheet_name="Validation_Rules"
    )

    # -----------------------------------------------------
    # 1. 更新正式 Field_Mapping
    # -----------------------------------------------------

    for _, ai_row in approved_field_mapping.iterrows():

        source_field = str(
            ai_row.get("source_field", "")
        ).strip()

        target_field = ai_row.get(
            "target_field"
        )

        if (
            source_field == ""
            or pd.isna(target_field)
            or str(target_field).strip() == ""
        ):
            continue

        target_field = str(
            target_field
        ).strip()

        mask = (
            formal_field_mapping["source_field"]
            .astype(str)
            .str.strip()
            == source_field
        )

        if mask.any():

            formal_field_mapping.loc[
                mask,
                "target_field"
            ] = target_field

    # -----------------------------------------------------
    # 2. 追加 Value_Mapping
    # -----------------------------------------------------

    if not approved_value_mapping.empty:

        new_value_rows = []

        for _, row in approved_value_mapping.iterrows():

            ai_field = str(
                row.get("field", "")
            ).strip()

            source_value = row.get(
                "source_value"
            )

            target_value = row.get(
                "target_value"
            )

            if (
                ai_field == ""
                or pd.isna(source_value)
                or pd.isna(target_value)
            ):
                continue

            new_value_rows.append({
                "field": ai_field,
                "source_value": source_value,
                "target_value": target_value,
                "confidence": row.get("confidence", ""),
                "review_required": False,
                "reason": row.get("reason", "")
            })

        if new_value_rows:

            new_values_df = pd.DataFrame(
                new_value_rows
            )

            formal_value_mapping = pd.concat(
                [
                    formal_value_mapping,
                    new_values_df
                ],
                ignore_index=True
            )

            formal_value_mapping = (
                formal_value_mapping
                .drop_duplicates(
                    subset=[
                        "field",
                        "source_value",
                        "target_value"
                    ]
                )
            )

    # -----------------------------------------------------
    # 3. 写回正式配置
    # -----------------------------------------------------

    with pd.ExcelWriter(
        MAPPING_FILE,
        engine="openpyxl",
        mode="w"
    ) as writer:

        formal_field_mapping.to_excel(
            writer,
            sheet_name="Field_Mapping",
            index=False
        )

        formal_value_mapping.to_excel(
            writer,
            sheet_name="Value_Mapping",
            index=False
        )

        validation_rules.to_excel(
            writer,
            sheet_name="Validation_Rules",
            index=False
        )


# =========================================================
# 5. 左侧 Sidebar
# =========================================================

with st.sidebar:

    st.markdown("## 👥 HRIS AI")

    demo_mode = st.toggle(
        "🧪 Demo Mode（不调用 API）",
        value=True
    )

    if demo_mode:

        st.success(
            "当前为 Demo 模式\n\n"
            "使用虚拟数据，不调用 DeepSeek API。"
        )

    else:

        st.warning(
            "当前为真实分析模式\n\n"
            "上传新数据后会调用 DeepSeek API。"
        )

    real_mode_authorized = False

    if not demo_mode:

        try:

            password = st.text_input(
                "Real Mode 密码",
                type="password"
            )

            if "REAL_MODE_PASSWORD" not in st.secrets:

                st.error(
                    "检测不到 REAL_MODE_PASSWORD。"
                )

                st.caption(
                    f"当前可读取的 Secrets："
                    f"{list(st.secrets.keys())}"
                )

                real_mode_authorized = False

            else:

                real_mode_authorized = (
                    password
                    == st.secrets["REAL_MODE_PASSWORD"]
                )

                if real_mode_authorized:

                    st.success(
                        "真实分析模式已解锁"
                    )

                else:

                    st.info(
                        "请输入正确的 Real Mode 密码。"
                    )

        except Exception as e:

            st.error(
                f"读取 Streamlit Secrets 时发生错误：{e}"
            )

            real_mode_authorized = False

    st.caption(
        "AI-assisted HRIS Data Integration"
    )

    st.divider()

    st.markdown("### 工作流")

    st.markdown(
        """
        **① 数据上传**
        上传员工 Excel

        **② AI 映射审核**
        DeepSeek 生成字段映射建议

        **③ 数据校验**
        检查格式、重复值、枚举值等

        **④ 结果报告**
        查看 AI 分析并下载结果
        """
    )

    st.divider()

    st.markdown("### 当前版本")

    st.info(
        "Prototype v1.0\n\n"
        "测试环境使用虚拟员工数据。"
    )

    st.divider()

    st.caption(
        "Python · pandas · DeepSeek · Streamlit · FastAPI"
    )

    st.divider()

    st.subheader("系统状态")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.success("● DeepSeek API")

    with col2:
        st.success("● Data Validation")

    with col3:
        st.success("● HRIS Mapping")


# =========================================================
# 6. 页面标题
# =========================================================

st.markdown(
    '<div class="main-title">👥 AI-assisted HRIS Data Integration</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle">'
    'Employee Data Mapping · Validation · AI Error Analysis'
    '</div>',
    unsafe_allow_html=True
)


# =========================================================
# 7. 工作流 Tab
# =========================================================

tab1, tab2, tab3, tab4 = st.tabs(
    [
        "① 数据上传",
        "② AI 映射审核",
        "③ 数据校验",
        "④ 结果报告"
    ]
)


# =========================================================
# TAB 1：数据上传
# =========================================================

with tab1:

    st.header("员工数据")

    st.caption(
        "支持 .xlsx 员工数据文件。"
        "演示环境请使用虚拟员工数据。"
    )

    st.write(
        "上传 Excel 后，系统会用它作为本次 HRIS 数据处理的数据源。"
    )

    st.download_button(
        label="📥 下载 120 条测试数据",
        data=(
            BASE_DIR
            / "data"
            / "employees.xlsx"
        ).read_bytes()
        if (
            BASE_DIR
            / "data"
            / "employees.xlsx"
        ).exists()
        else b"",
        file_name="sample_employees.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )

    uploaded_file = st.file_uploader(
        "上传员工 Excel",
        type=["xlsx"],
        key="employee_upload"
    )

    if uploaded_file is not None:

        save_uploaded_file(
            uploaded_file
        )

        st.success(
            f"已加载：{uploaded_file.name}"
        )

        try:

            employees = pd.read_excel(
                BytesIO(uploaded_file.getvalue())
            )

            normalize_df(employees)

            st.divider()

            col1, col2, col3, col4 = st.columns(4)

            with col1:
                st.metric(
                    "员工总数",
                    len(employees)
                )

            with col2:
                st.metric(
                    "字段数",
                    len(employees.columns)
                )

            with col3:
                st.metric(
                    "重复工号",
                    employees["员工编号"].duplicated().sum()
                    if "员工编号" in employees.columns
                    else 0
                )

            with col4:
                st.metric(
                    "空值数量",
                    int(
                        employees.isna()
                        .sum()
                        .sum()
                    )
                )

            st.subheader(
                "数据预览"
            )

            st.dataframe(
                employees.head(30),
                width="stretch",
                height=500
            )

        except Exception as e:

            st.error(
                f"读取 Excel 失败：{e}"
            )


# =========================================================
# TAB 2：AI 映射审核
# =========================================================

with tab2:

    st.header("AI 字段映射")

    st.write(
        "DeepSeek 会根据原始 Excel 字段和 HRIS Schema "
        "生成映射建议。审核后才能正式采用。"
    )

    if not INPUT_FILE.exists():

        st.info(
            "请先在“① 数据上传”中上传员工 Excel。"
        )

    else:

        if demo_mode:

            if st.button(
                "🧪 查看 Demo 映射（不调用 API）",
                type="primary",
                width="stretch"
            ):

                if not MAPPING_FILE.exists():

                    st.error(
                        "找不到 mapping_config.xlsx"
                    )

                else:

                    formal_field_df = pd.read_excel(
                        MAPPING_FILE,
                        sheet_name="Field_Mapping"
                    )

                    formal_value_df = pd.read_excel(
                        MAPPING_FILE,
                        sheet_name="Value_Mapping"
                    )

                    # 创建“演示版 AI 映射”
                    demo_field_df = formal_field_df[
                        [
                            "source_field",
                            "target_field"
                        ]
                    ].copy()

                    demo_field_df[
                        "confidence"
                    ] = "high"

                    demo_field_df[
                        "transformation"
                    ] = "按预设规则转换"

                    demo_field_df[
                        "review_required"
                    ] = False

                    demo_field_df[
                        "approved"
                    ] = True

                    demo_field_df[
                        "reason"
                    ] = "Demo 模式使用预先确认的映射规则"

                    demo_value_df = formal_value_df.copy()

                    demo_value_df[
                        "confidence"
                    ] = "high"

                    demo_value_df[
                        "review_required"
                    ] = False

                    demo_value_df[
                        "approved"
                    ] = True

                    demo_value_df[
                        "reason"
                    ] = "Demo 模式使用预先确认的值映射"

                    st.session_state[
                        "ai_field_mapping"
                    ] = demo_field_df

                    st.session_state[
                        "ai_value_mapping"
                    ] = demo_value_df

                    st.success(
                        "Demo 映射已加载，不消耗 DeepSeek API。"
                    )


        else:

            if not real_mode_authorized:

                st.warning(
                    "请先解锁 Real Mode。"
                )

            else:

                if st.button(
                    "🤖 开始 AI 字段分析（调用 DeepSeek）",
                    type="primary",
                    width="stretch"
                ):

                    with st.spinner(
                        "正在调用 DeepSeek 分析字段..."
                    ):

                        success, stdout, stderr = (
                            run_python_script(
                                "ai_mapper.py"
                            )
                        )

                    if success:

                        field_df, value_df = (
                            load_ai_mapping()
                        )

                        if field_df is not None:

                            st.session_state[
                                "ai_field_mapping"
                            ] = field_df

                            st.session_state[
                                "ai_value_mapping"
                            ] = value_df

                            st.success(
                                "DeepSeek 字段分析完成。"
                            )

                        else:

                            st.error(
                                "AI 分析完成，但没有找到结果。"
                            )

                    else:

                        st.error(
                            "DeepSeek 字段分析失败。"
                        )

                        st.code(
                            stderr,
                            language="text"
                        )

        # -------------------------------------------------
        # 显示 AI 映射
        # -------------------------------------------------

        if (
            "ai_field_mapping"
            in st.session_state
        ):

            field_df = (
                st.session_state[
                    "ai_field_mapping"
                ].copy()
            )

            # 转换审核状态
            if "review_required" in field_df.columns:

                field_df[
                    "review_required"
                ] = field_df[
                    "review_required"
                ].apply(to_bool)

            else:

                field_df[
                    "review_required"
                ] = True

            # 默认批准：
            # 不需要审核 → 自动勾选
            field_df[
                "approved"
            ] = ~field_df[
                "review_required"
            ]

            st.subheader(
                "AI 字段映射建议"
            )

            st.caption(
                "你可以直接在表格里修改 target_field，"
                "并勾选 approved。"
            )

            edited_field_df = st.data_editor(
                field_df,
                width="stretch",
                height=450,
                num_rows="fixed",
                key="field_mapping_editor"
            )

            st.session_state[
                "edited_field_mapping"
            ] = edited_field_df

            # -------------------------------------------------
            # Value Mapping
            # -------------------------------------------------

            st.subheader(
                "AI 值映射建议"
            )

            value_df = (
                st.session_state[
                    "ai_value_mapping"
                ].copy()
            )

            if not value_df.empty:

                if "review_required" in value_df.columns:

                    value_df[
                        "review_required"
                    ] = value_df[
                        "review_required"
                    ].apply(to_bool)

                else:

                    value_df[
                        "review_required"
                    ] = False

                value_df[
                    "approved"
                ] = ~value_df[
                    "review_required"
                ]

                edited_value_df = st.data_editor(
                    value_df,
                    width="stretch",
                    height=350,
                    num_rows="fixed",
                    key="value_mapping_editor"
                )

                st.session_state[
                    "edited_value_mapping"
                ] = edited_value_df

            else:

                edited_value_df = pd.DataFrame()

                st.info(
                    "本次 AI 没有生成值映射建议。"
                )

            st.divider()

            st.warning(
                "请确认无误后再点击下面的按钮。"
                "点击后，批准的映射会写入正式 mapping_config.xlsx。"
            )

            if st.button(
                "✅ 确认审核结果并采用映射",
                type="primary",
                width="stretch"
            ):

                approved_fields = (
                    edited_field_df[
                        edited_field_df[
                            "approved"
                        ] == True
                    ].copy()
                )

                if not edited_value_df.empty:

                    approved_values = (
                        edited_value_df[
                            edited_value_df[
                                "approved"
                            ] == True
                        ].copy()
                    )

                else:

                    approved_values = (
                        pd.DataFrame()
                    )

                try:

                    apply_approved_mapping(
                        approved_fields,
                        approved_values
                    )

                    st.success(
                        f"已正式采用 "
                        f"{len(approved_fields)} 条字段映射。"
                    )

                    st.session_state[
                        "mapping_confirmed"
                    ] = True

                except Exception as e:

                    st.error(
                        f"保存正式映射失败：{e}"
                    )

        else:

            st.info(
                "点击“开始 AI 字段分析”后，这里会出现 DeepSeek 的映射建议。"
            )


# =========================================================
# TAB 3：数据校验
# =========================================================

with tab3:

    st.header("数据校验与 HRIS 转换")

    if not INPUT_FILE.exists():

        st.info(
            "请先上传员工数据。"
        )

    else:

        if not st.session_state.get(
            "mapping_confirmed",
            False
        ):

            st.info(
                "建议先到“② AI 映射审核”确认映射。"
            )

        if demo_mode:

            button_text = (
                "🧪 运行 Demo 数据校验"
            )

        else:

            button_text = (
                "🚀 开始真实数据校验"
            )

        if st.button(
            button_text,
            type="primary",
            width="stretch"
        ):

            with st.spinner(
                "正在进行字段转换、值映射和数据校验..."
            ):

                success, stdout, stderr = (
                    run_python_script(
                        "main.py"
                    )
                )

            if success:

                st.success(
                    "数据校验与 HRIS 转换完成。"
                )

                st.session_state[
                    "validation_stdout"
                ] = stdout

            else:

                st.error(
                    "数据处理失败。"
                )

                st.code(
                    stderr,
                    language="text"
                )

        # -------------------------------------------------
        # 显示结果统计
        # -------------------------------------------------

        if (
            HRIS_FILE.exists()
            and ERROR_FILE.exists()
        ):

            try:

                hris_df = pd.read_excel(
                    HRIS_FILE
                )

                error_df = pd.read_excel(
                    ERROR_FILE
                )

                if not error_df.empty:

                    st.subheader("错误类型分布")

                    error_summary = (
                        error_df[
                            "error_code"
                        ]
                        .value_counts()
                        .rename_axis("error_code")
                        .reset_index(
                            name="count"
                        )
                    )

                    st.bar_chart(
                        error_summary.set_index(
                            "error_code"
                        )
                    )

                employees = pd.read_excel(
                    INPUT_FILE
                )

                total = len(
                    employees
                )

                success_count = len(
                    hris_df
                )

                error_rows = len(
                    error_df
                )

                st.divider()

                col1, col2, col3 = st.columns(3)

                with col1:
                    st.metric(
                        "原始员工",
                        total
                    )

                with col2:

                    success_rate = (
                        success_count / total
                        if total > 0
                        else 0
                    )

                    st.metric(
                        "通过校验",
                        success_count,
                        delta=f"{success_rate:.1%}"
                    )

                with col3:

                    error_rate = (
                        error_rows / total
                        if total > 0
                        else 0
                    )

                    st.metric(
                        "异常记录",
                        error_rows,
                        delta=f"-{error_rate:.1%}"
                    )

                st.subheader(
                    "错误数据"
                )

                if error_df.empty:

                    st.success(
                        "没有发现数据错误。"
                    )

                else:

                    st.dataframe(
                        normalize_df(error_df),
                        width="stretch",
                        height=400
                    )

            except Exception as e:

                st.error(
                    f"读取处理结果失败：{e}"
                )


# =========================================================
# TAB 4：结果报告
# =========================================================

with tab4:

    st.header("处理结果")

    # -----------------------------------------------------
    # AI 错误分析
    # -----------------------------------------------------

    if ERROR_FILE.exists():

        error_df = pd.read_excel(
            ERROR_FILE
        )

        # =====================================================
        # Demo Mode：不调用 API
        # =====================================================

        if demo_mode:

            st.info(
                "🧪 Demo 模式："
                "以下为规则化演示分析，不调用 DeepSeek API。"
            )

            if error_df.empty:

                st.success(
                    "没有发现数据错误。"
                )

            else:

                demo_analysis = []

                explanations = {

                    "INVALID_ENUM": {
                        "problem":
                            "字段值不符合当前允许的枚举值。",
                        "cause":
                            "原始数据出现未定义的标准值。",
                        "action":
                            "确认该值对应的 HRIS 标准代码。"
                    },

                    "DUPLICATE_ID": {
                        "problem":
                            "员工编号已经出现重复。",
                        "cause":
                            "多个员工记录使用了相同的员工编号。",
                        "action":
                            "检查原始数据并确认唯一员工编号。"
                    },

                    "INVALID_EMAIL": {
                        "problem":
                            "邮箱格式不符合基本邮箱格式。",
                        "cause":
                            "邮箱可能缺少域名或其他必要部分。",
                        "action":
                            "修改为有效的工作邮箱。"
                    },

                    "INVALID_DATE": {
                        "problem":
                            "日期无法转换为合法日期。",
                        "cause":
                            "日期格式错误或日期本身不存在。",
                        "action":
                            "确认并修改入职日期。"
                    },

                    "REQUIRED": {
                        "problem":
                            "必填字段为空。",
                        "cause":
                            "原始员工数据缺少必要字段。",
                        "action":
                            "补充该字段后重新处理。"
                    }
                }

                for error_code, group in (
                    error_df
                    .groupby(
                        "error_code",
                        dropna=False
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

                    demo_analysis.append({

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

                demo_analysis_df = pd.DataFrame(
                    demo_analysis
                )

                st.subheader(
                    "Demo 错误分析"
                )

                for _, row in (
                    demo_analysis_df.iterrows()
                ):

                    with st.expander(
                        f"{row['error_code']} · "
                        f"{row['field']}"
                    ):

                        st.write(
                            "**问题：**",
                            row["problem"]
                        )

                        st.write(
                            "**可能原因：**",
                            row["possible_cause"]
                        )

                        st.write(
                            "**建议处理：**",
                            row["recommended_action"]
                        )

                        st.warning(
                            "需要人工确认"
                        )


        # =====================================================
        # Real Mode：调用 DeepSeek
        # =====================================================

        else:

            if not real_mode_authorized:

                st.warning(
                    "请先解锁 Real Mode。"
                )

            else:

                if not error_df.empty:

                    if st.button(
                        "🧠 运行 DeepSeek 错误分析（消耗 API）",
                        type="primary",
                        width="stretch"
                    ):

                        with st.spinner(
                            "正在调用 DeepSeek 分析错误..."
                        ):

                            success, stdout, stderr = (
                                run_python_script(
                                    "ai_error_analyzer.py"
                                )
                            )

                        if success:

                            st.success(
                                "DeepSeek 错误分析完成。"
                            )

                        else:

                            st.error(
                                "DeepSeek 错误分析失败。"
                            )

                            st.code(
                                stderr,
                                language="text"
                            )

    # -----------------------------------------------------
    # AI 分析结果
    # -----------------------------------------------------

    if AI_ERROR_FILE.exists():

        try:

            ai_error_df = pd.read_excel(
                AI_ERROR_FILE
            )

            st.subheader(
                "DeepSeek 错误分析"
            )

            if ai_error_df.empty:

                st.success(
                    "当前没有需要 AI 分析的问题。"
                )

            else:

                for _, row in ai_error_df.iterrows():

                    with st.expander(
                        f"{row.get('error_code', '')} · "
                        f"{row.get('field', '')}"
                    ):

                        st.write(
                            "**问题：**",
                            row.get(
                                "problem",
                                ""
                            )
                        )

                        st.write(
                            "**可能原因：**",
                            row.get(
                                "possible_cause",
                                ""
                            )
                        )

                        st.write(
                            "**建议处理：**",
                            row.get(
                                "recommended_action",
                                ""
                            )
                        )

                        review = row.get(
                            "review_required",
                            False
                        )

                        if str(review).lower() == "true":

                            st.warning(
                                "需要人工确认"
                            )

                        else:

                            st.success(
                                "无需人工确认"
                            )

        except Exception as e:

            st.error(
                f"读取 AI 分析失败：{e}"
            )

    st.divider()

    # -----------------------------------------------------
    # 下载文件
    # -----------------------------------------------------

    st.subheader(
        "下载结果文件"
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        data = excel_bytes(
            HRIS_FILE
        )

        if data:

            st.download_button(
                "⬇️ HRIS 导入文件",
                data=data,
                file_name="hris_import.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                width="stretch"
            )

    with col2:

        data = excel_bytes(
            ERROR_FILE
        )

        if data:

            st.download_button(
                "⬇️ 错误报告",
                data=data,
                file_name="sync_errors.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                width="stretch"
            )

    with col3:

        data = excel_bytes(
            AI_ERROR_FILE
        )

        if data:

            st.download_button(
                "⬇️ AI 错误分析",
                data=data,
                file_name="ai_error_analysis.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                width="stretch"
            )

    # -----------------------------------------------------
    # 最终文件状态
    # -----------------------------------------------------

    st.divider()

    st.subheader(
        "系统文件状态"
    )

    files_status = {
        "AI 映射建议": AI_MAPPING_FILE,
        "HRIS 导入文件": HRIS_FILE,
        "错误报告": ERROR_FILE,
        "AI 错误分析": AI_ERROR_FILE
    }

    for label, path in files_status.items():

        if path.exists():

            st.success(
                f"✓ {label}"
            )

        else:

            st.caption(
                f"○ {label}：尚未生成"
            )


# =========================================================
# 8. 关于本项目
# =========================================================

with st.expander("ℹ️ 关于这个项目"):

    st.write(
        """
        本项目是一个 AI-assisted HRIS Data Integration Prototype。

        主要用于模拟员工数据从 Excel 到 HRIS 的处理流程，
        包括字段映射、值映射、数据校验、AI 异常分析和
        Mock HRIS API 同步。

        当前版本使用虚拟员工数据，不连接真实企业 HRIS。
        """
    )
