import hashlib
from pathlib import Path

import pandas as pd
import streamlit as st


def normalize_bool(value, default=False):
    """
    将 Excel / AI 返回的各种 True/False 表示
    统一转换成 Python bool。
    """

    if pd.isna(value):
        return default

    if isinstance(value, bool):
        return value

    text = str(value).strip().lower()

    if text in {
        "true",
        "1",
        "yes",
        "y",
        "是"
    }:
        return True

    if text in {
        "false",
        "0",
        "no",
        "n",
        "否"
    }:
        return False

    return default


from src.web_pipeline import (
    load_project_config,
    ai_analyze_mapping,
    build_working_field_mapping,
    build_working_value_mapping,
    transform_employee_data,
    validate_employee_data,
    create_hris_dataframe,
    demo_error_analysis,
    ai_analyze_errors,
    dataframe_to_excel_bytes,
    dataframes_to_excel_bytes,
)


# =========================================================
# 1. 项目路径
# =========================================================

BASE_DIR = Path(__file__).resolve().parent

DATA_DIR = BASE_DIR / "data"

DEMO_FILE = (
    DATA_DIR
    / "demo_employees.xlsx"
)


# =========================================================
# 2. 页面设置
# =========================================================

st.set_page_config(
    page_title="AI-assisted HRIS Data Integration",
    page_icon="👥",
    layout="wide",
    initial_sidebar_state="expanded"
)


# =========================================================
# 3. Session State
# =========================================================

defaults = {

    "source_df": None,

    "ai_mapping_done": False,

    "ai_field_df": None,

    "ai_value_df": None,

    "file_signature": None,

    "approved_field_mapping": None,

    "approved_value_mapping": None,

    "mapping_confirmed": False,

    "validation_done": False,

    "validation_result": None,

    "report_ready": False,

    "goto_report": False,

    "hris_df": None,

    "error_df": None,

    "ai_error_df": None,

}

for key, value in defaults.items():

    if key not in st.session_state:

        st.session_state[key] = value


# =========================================================
# 4. 读取配置
# =========================================================

try:

    (
        formal_field_mapping,
        formal_value_mapping,
        validation_rules,
        hris_schema

    ) = load_project_config(
        BASE_DIR
    )

except Exception as e:

    st.error(
        f"读取项目配置失败：{e}"
    )

    st.stop()


# =========================================================
# 5. Sidebar
# =========================================================

with st.sidebar:

    st.markdown(
        "## 👥 HRIS AI"
    )

    st.caption(
        "AI-assisted HRIS Data Integration"
    )

    st.divider()

    demo_mode = st.toggle(
        "🧪 Demo Mode（不调用 API）",
        value=True
    )

    if demo_mode:

        st.success(
            "Demo 模式\n\n"
            "使用固定虚拟数据。\n"
            "不会调用 DeepSeek API。"
        )

    else:

        st.warning(
            "Real Mode\n\n"
            "上传的新数据可能调用 DeepSeek API。"
        )

    # -----------------------------------------
    # Real Mode Password
    # -----------------------------------------

    real_mode_authorized = False

    if not demo_mode:

        try:

            real_password = st.text_input(
                "Real Mode 密码",
                type="password"
            )

            if "REAL_MODE_PASSWORD" not in st.secrets:

                st.error(
                    "未检测到 REAL_MODE_PASSWORD。"
                )

            else:

                real_mode_authorized = (
                    real_password
                    == st.secrets[
                        "REAL_MODE_PASSWORD"
                    ]
                )

                if real_mode_authorized:

                    st.success(
                        "✓ Real Mode 已解锁"
                    )

                else:

                    st.info(
                        "请输入正确的 Real Mode 密码。"
                    )

        except Exception as e:

            st.error(
                f"读取 Real Mode Secret 失败：{e}"
            )

    st.divider()

    st.sidebar.subheader(
        "工作流"
    )


    upload_done = (
        st.session_state.get(
            "source_df"
        )
        is not None
    )


    mapping_done = (
        st.session_state.get(
            "mapping_confirmed",
            False
        )
    )


    validation_done = (
        st.session_state.get(
            "validation_done",
            False
        )
    )


    report_done = (
        st.session_state.get(
            "report_ready",
            False
        )
    )


    def workflow_status(done):
        if done:
            return "✅"
        else:
            return "⚪"



    st.sidebar.write(
        f"{workflow_status(upload_done)} 数据上传"
    )


    st.sidebar.write(
        f"{workflow_status(mapping_done)} AI映射审核"
    )


    st.sidebar.write(
        f"{workflow_status(validation_done)} 数据校验"
    )


    st.sidebar.write(
        f"{workflow_status(report_done)} 结果报告"
    )

    st.divider()

    st.caption(
        "Prototype v2.0"
    )

    st.caption(
        "Python · pandas · DeepSeek · Streamlit"
    )


# =========================================================
# 6. Title
# =========================================================

st.title(
    "👥 AI-assisted HRIS Data Integration"
)

st.caption(
    "Employee Data Mapping · Validation · AI Error Analysis"
)

st.divider()


# =========================================================
# 7. Tab
# =========================================================

if st.session_state.get("goto_report"):

    st.session_state["main_tabs"] = "④ 结果报告"

    st.session_state["goto_report"] = False


tab1, tab2, tab3, tab4 = st.tabs(
    [
        "① 数据上传",
        "② AI 映射审核",
        "③ 数据校验",
        "④ 结果报告"
    ],
    default="① 数据上传",
    key="main_tabs",
    on_change="rerun"
)


# =========================================================
# TAB 1：数据
# =========================================================

with tab1:

    st.header(
        "员工数据"
    )

    # =====================================================
    # Demo Mode
    # =====================================================

    if demo_mode:

        if not DEMO_FILE.exists():

            st.error(
                "找不到 Demo 数据："
                "data/demo_employees.xlsx"
            )

            st.stop()

        try:

            demo_df = pd.read_excel(
                DEMO_FILE
            )

            st.session_state[
                "source_df"
            ] = demo_df

            st.success(
                "🧪 当前使用 120 条虚拟员工数据"
            )

            st.caption(
                "Demo 数据不会上传到 DeepSeek，"
                "也不会消耗 API。"
            )

            if st.button(
                "🔄 重新加载 Demo 数据",
                width="stretch"
            ):

                st.session_state[
                    "source_df"
                ] = pd.read_excel(
                    DEMO_FILE
                )

                st.rerun()

        except Exception as e:

            st.error(
                f"读取 Demo 数据失败：{e}"
            )


    # =====================================================
    # Real Mode
    # =====================================================

    else:

        if not real_mode_authorized:

            st.warning(
                "请先输入 Real Mode 密码。"
            )

        else:

            uploaded_file = st.file_uploader(
                "上传新的员工 Excel",
                type=["xlsx"],
                key="real_employee_upload"
            )

            if uploaded_file is not None:

                try:

                    new_df = pd.read_excel(
                        uploaded_file
                    )

                    file_bytes = uploaded_file.getvalue()

                    file_signature = hashlib.md5(
                        file_bytes
                    ).hexdigest()

                    if (
                        "file_signature" not in st.session_state
                        or st.session_state["file_signature"] != file_signature
                    ):

                        st.session_state["file_signature"] = file_signature

                        st.session_state["ai_mapping_done"] = False
                        st.session_state["ai_field_df"] = None
                        st.session_state["ai_value_df"] = None
                        st.session_state["mapping_confirmed"] = False

                    st.session_state[
                        "source_df"
                    ] = new_df

                    st.success(
                        f"已读取 {len(new_df)} 条员工数据。"
                    )

                    st.caption(
                        "当前数据只在本次网页会话中处理，"
                        "不会保存成你的电脑上的 employees.xlsx。"
                    )

                except Exception as e:

                    st.error(
                        f"读取 Excel 失败：{e}"
                    )


    # =====================================================
    # Data Preview
    # =====================================================

    source_df = st.session_state[
        "source_df"
    ]

    if source_df is not None:

        st.divider()

        col1, col2, col3, col4 = (
            st.columns(4)
        )

        with col1:

            st.metric(
                "员工总数",
                len(source_df)
            )

        with col2:

            st.metric(
                "字段数量",
                len(source_df.columns)
            )

        with col3:

            duplicate_count = 0

            if "员工编号" in source_df.columns:

                duplicate_count = int(
                    source_df[
                        "员工编号"
                    ]
                    .duplicated()
                    .sum()
                )

            st.metric(
                "重复工号",
                duplicate_count
            )

        with col4:

            st.metric(
                "空值数量",
                int(
                    source_df
                    .isna()
                    .sum()
                    .sum()
                )
            )

        st.subheader(
            "数据预览"
        )

        st.dataframe(
            source_df.head(30),
            width="stretch",
            height=500
        )


# =========================================================
# TAB 2：AI 映射审核
# =========================================================


def go_to_validation():
    st.session_state["main_tabs"] = "③ 数据校验"


with tab2:

    st.subheader("AI 映射审核")

    # =========================================================
    # 1. Session State 初始化
    #    只能在 key 不存在时初始化
    # =========================================================

    if "ai_mapping_done" not in st.session_state:
        st.session_state["ai_mapping_done"] = False

    if "ai_field_df" not in st.session_state:
        st.session_state["ai_field_df"] = None

    if "ai_value_df" not in st.session_state:
        st.session_state["ai_value_df"] = None

    if "mapping_confirmed" not in st.session_state:
        st.session_state["mapping_confirmed"] = False


    # =========================================================
    # 2. Demo Mode
    # =========================================================

    if demo_mode:

        st.info(
            "Demo Mode：使用项目预设映射，不调用 DeepSeek。"
        )

        if not st.session_state["ai_mapping_done"]:

            field_demo = formal_field_mapping.copy()
            value_demo = formal_value_mapping.copy()

            field_demo["approved"] = True

            if not value_demo.empty:
                value_demo["approved"] = True

            st.session_state["ai_field_df"] = field_demo
            st.session_state["ai_value_df"] = value_demo
            st.session_state["ai_mapping_done"] = True

        field_df = st.session_state["ai_field_df"].copy()
        value_df = st.session_state["ai_value_df"].copy()

    # =========================================================
    # 3. Real Mode
    # =========================================================

    else:

        st.info(
            "Real Mode：上传的新员工数据将由 DeepSeek 进行字段映射分析。"
        )

        # -----------------------------------------------------
        # 只有用户主动点击按钮，才调用 DeepSeek
        # -----------------------------------------------------

        if not real_mode_authorized:

            st.warning("请先在侧边栏解锁 Real Mode。")

        elif source_df is None:

            st.info("请先在“数据上传”页面上传 Excel 文件。")

        elif not st.session_state["ai_mapping_done"]:

            start_analysis = st.button(
                "开始字段分析",
                type="primary",
                key="start_ai_mapping"
            )

            if start_analysis:

                with st.spinner("DeepSeek 正在分析字段，请稍候……"):

                    try:

                        field_df, value_df = ai_analyze_mapping(
                            source_df,
                            formal_field_mapping,
                            formal_value_mapping,
                            hris_schema
                        )

                        field_df = field_df.copy()
                        field_df["approved"] = True

                        value_df = value_df.copy()

                        if not value_df.empty:
                            value_df["approved"] = True

                        st.session_state["ai_field_df"] = field_df
                        st.session_state["ai_value_df"] = value_df
                        st.session_state["ai_mapping_done"] = True
                        st.session_state["mapping_confirmed"] = False

                        st.rerun()

                    except Exception as e:

                        st.error(
                            f"DeepSeek 分析失败：{e}"
                        )

        # -----------------------------------------------------
        # AI 分析完成后
        # -----------------------------------------------------

        else:

            field_df = st.session_state["ai_field_df"].copy()

            value_df = st.session_state["ai_value_df"].copy()


    # =========================================================
    # 4. 分析结果展示
    # =========================================================

    if st.session_state["ai_mapping_done"]:

        st.success("字段分析已完成，请审核 AI 映射结果。")

        field_df = st.session_state["ai_field_df"].copy()
        value_df = st.session_state["ai_value_df"].copy()


        # =====================================================
        # 表一：字段映射
        # =====================================================

        st.markdown("### 表一：字段映射")

        field_df = field_df.copy()

        if "reason" not in field_df.columns:
            field_df["reason"] = (
                "AI根据字段名称和数据样例进行匹配"
            )

        if "confidence" in field_df.columns:

            def confidence_status(x):

                text = str(x).strip().lower()

                if text in {"high", "高"}:
                    return "🟢 High Confidence"

                elif text in {"medium", "中"}:
                    return "🟡 Review Suggested"

                elif text in {"low", "低"}:
                    return "🔴 Manual Review"

                return "Unknown"


            field_df["confidence_status"] = (
                field_df["confidence"]
                .apply(confidence_status)
            )

        field_df["approved"] = True

        if "review_required" in field_df.columns:
            field_df["review_required"] = (
                field_df["review_required"]
                .apply(lambda x: normalize_bool(x, False))
            )
        else:
            field_df["review_required"] = False

        edited_field_df = st.data_editor(
            field_df,
            key="field_mapping_editor",
            hide_index=True,
            width="stretch",
            num_rows="fixed",
            disabled=[
                column
                for column in field_df.columns
                if column not in ["approved", "review_required"]
            ],
        )

        # -----------------------------------------------------
        # 保存用户刚刚编辑后的结果
        # -----------------------------------------------------

        st.session_state["ai_field_df"] = edited_field_df.copy()

        st.subheader("AI Mapping Explanation")

        for _, row in field_df.iterrows():

            with st.expander(
                f"{row['source_field']} → {row['standard_field']}"
            ):

                st.write(
                    "Confidence:",
                    row["confidence"]
                )

                st.write(
                    "Reason:",
                    row["reason"]
                )


        # =====================================================
        # 表二：值映射
        # =====================================================

        st.markdown("### 表二：值映射")

        if not value_df.empty:

            value_df = value_df.copy()

            value_df["approved"] = True

            if "review_required" in value_df.columns:
                value_df["review_required"] = (
                    value_df["review_required"]
                    .apply(lambda x: normalize_bool(x, False))
                )
            else:
                value_df["review_required"] = False

            edited_value_df = st.data_editor(
                value_df,
                key="value_mapping_editor",
                hide_index=True,
                width="stretch",
                num_rows="fixed",
                disabled=[
                    column
                    for column in value_df.columns
                    if column not in ["approved", "review_required"]
                ],
            )

            st.session_state["ai_value_df"] = edited_value_df.copy()

        else:

            st.info("AI 未发现需要新增的值映射。")


        # =====================================================
        # 5. 当前审核状态
        # =====================================================

        current_field_df = st.session_state["ai_field_df"]

        current_value_df = st.session_state["ai_value_df"]

        field_total = len(current_field_df)

        field_approved = (
            int(current_field_df["approved"].sum())
            if "approved" in current_field_df.columns
            else 0
        )

        if (
            current_value_df is not None
            and not current_value_df.empty
            and "approved" in current_value_df.columns
        ):

            value_total = len(current_value_df)

            value_approved = int(
                current_value_df["approved"].sum()
            )

        else:

            value_total = 0
            value_approved = 0


        st.caption(
            f"字段映射：{field_approved}/{field_total} 已审核"
        )

        if value_total > 0:

            st.caption(
                f"值映射：{value_approved}/{value_total} 已审核"
            )


        # =====================================================
        # 6. 下一步
        # =====================================================

        st.divider()

        confirm_mapping = st.button(
            "确认映射并进入数据校验 →",
            type="primary",
            key="confirm_mapping",
            on_click=go_to_validation
        )

        if confirm_mapping:

            approved_fields = (
                st.session_state["ai_field_df"]
                .copy()
            )

            approved_values = (
                st.session_state["ai_value_df"]
                .copy()
            )

            if "approved" in approved_fields.columns:
                approved_fields = approved_fields[
                    approved_fields["approved"] == True
                ].copy()

            if (
                approved_values is not None
                and not approved_values.empty
                and "approved" in approved_values.columns
            ):
                approved_values = approved_values[
                    approved_values["approved"] == True
                ].copy()

            st.session_state["approved_field_mapping"] = (
                approved_fields
            )

            st.session_state["approved_value_mapping"] = (
                approved_values
            )

            st.session_state["mapping_confirmed"] = True

            st.session_state["validation_done"] = False
            st.session_state["validation_result"] = None

            st.success(
                "映射审核已确认，可以进入“数据校验”页面。"
            )


# =========================================================
# TAB 3：数据校验
# =========================================================

with tab3:

    st.header(
        "数据校验与 HRIS 转换"
    )

    source_df = st.session_state[
        "source_df"
    ]

    if source_df is None:

        st.info(
            "请先准备员工数据。"
        )

    elif not st.session_state[
        "mapping_confirmed"
    ]:

        st.warning(
            "请先完成 AI 映射审核。"
        )

    else:

        if (
            st.session_state.get("mapping_confirmed")
            and not st.session_state.get("validation_done")
        ):

            with st.spinner(
                "正在执行 HRIS 数据校验，请稍候..."
            ):

                try:

                    # -------------------------------------
                    # Demo 使用正式配置
                    # -------------------------------------

                    if demo_mode:

                        working_field_mapping = (
                            formal_field_mapping.copy()
                        )

                        working_value_mapping = (
                            formal_value_mapping.copy()
                        )

                    # -------------------------------------
                    # Real 使用人工批准的 AI Mapping
                    # -------------------------------------

                    else:

                        working_field_mapping = (
                            build_working_field_mapping(

                                formal_field_mapping,

                                st.session_state[
                                    "approved_field_mapping"
                                ]

                            )
                        )

                        working_value_mapping = (
                            build_working_value_mapping(

                                formal_value_mapping,

                                st.session_state[
                                    "approved_value_mapping"
                                ]

                            )
                        )

                    # -------------------------------------
                    # Transform
                    # -------------------------------------

                    standard_df = (
                        transform_employee_data(

                            source_df,

                            working_field_mapping,

                            working_value_mapping

                        )
                    )

                    # -------------------------------------
                    # Validate
                    # -------------------------------------

                    (
                        standard_df,
                        error_df,
                        bad_indexes
                    ) = validate_employee_data(

                        standard_df,

                        validation_rules

                    )

                    # -------------------------------------
                    # HRIS
                    # -------------------------------------

                    hris_df = (
                        create_hris_dataframe(

                            standard_df,

                            working_field_mapping,

                            bad_indexes

                        )
                    )

                    st.session_state[
                        "hris_df"
                    ] = hris_df

                    st.session_state[
                        "error_df"
                    ] = error_df

                    st.session_state[
                        "ai_error_df"
                    ] = None

                    st.session_state[
                        "validation_result"
                    ] = {
                        "total": len(source_df),
                        "success": len(hris_df),
                        "errors": len(error_df),
                    }

                    st.session_state[
                        "validation_done"
                    ] = True

                    st.session_state["report_ready"] = True
                    st.session_state["goto_report"] = True

                    st.rerun()

                except Exception as e:

                    st.error(
                        f"数据校验失败：{e}"
                    )

        # ---------------------------------------------
        # 校验完成后的指标
        # ---------------------------------------------

        if st.session_state.get("validation_done"):

            result = st.session_state[
                "validation_result"
            ]

            st.success(
                "数据校验完成"
            )

            col1, col2, col3 = st.columns(3)

            with col1:

                st.metric(
                    "总记录",
                    result["total"]
                )

            with col2:

                st.metric(
                    "通过",
                    result["success"]
                )

            with col3:

                st.metric(
                    "错误",
                    result["errors"]
                )

        # ---------------------------------------------
        # 重新执行
        # ---------------------------------------------

        if st.button(
            "🔄 重新执行数据校验"
        ):

            st.session_state[
                "validation_done"
            ] = False

            st.session_state[
                "validation_result"
            ] = None

            st.rerun()

        # ---------------------------------------------
        # Results
        # ---------------------------------------------

        hris_df = st.session_state[
            "hris_df"
        ]

        error_df = st.session_state[
            "error_df"
        ]

        if (
            hris_df is not None
            and error_df is not None
        ):

            total = len(
                source_df
            )

            success_count = len(
                hris_df
            )

            error_count = len(
                error_df
            )

            success_rate = (
                success_count / total
                if total > 0
                else 0
            )

            error_rate = (
                error_count / total
                if total > 0
                else 0
            )

            st.divider()

            col1, col2, col3 = (
                st.columns(3)
            )

            with col1:

                st.metric(
                    "原始员工",
                    total
                )

            with col2:

                st.metric(
                    "通过校验",
                    success_count,
                    delta=f"{success_rate:.1%}"
                )

            with col3:

                st.metric(
                    "错误记录",
                    error_count,
                    delta=f"-{error_rate:.1%}"
                )

            if not error_df.empty:

                st.subheader(
                    "错误记录"
                )

                st.dataframe(
                    error_df,
                    width="stretch",
                    height=450
                )

                # 错误类型统计
                st.subheader(
                    "错误类型分布"
                )

                error_summary = (
                    error_df[
                        "error_code"
                    ]
                    .value_counts()
                    .rename_axis(
                        "error_code"
                    )
                    .reset_index(
                        name="count"
                    )
                )

                st.bar_chart(
                    error_summary.set_index(
                        "error_code"
                    )
                )

            else:

                st.success(
                    "没有发现数据错误。"
                )


# =========================================================
# TAB 4：结果报告
# =========================================================

with tab4:

    st.header(
        "结果报告"
    )

    error_df = st.session_state[
        "error_df"
    ]

    hris_df = st.session_state[
        "hris_df"
    ]

    validation_result = st.session_state.get(
        "validation_result"
    )

    if validation_result:

        st.subheader(
            "📊 数据处理摘要"
        )

        total = validation_result.get(
            "total",
            0
        )

        success = validation_result.get(
            "success",
            0
        )

        errors = validation_result.get(
            "errors",
            0
        )

        rate = (
            success / total * 100
            if total > 0
            else 0
        )


        col1, col2, col3, col4 = st.columns(4)


        with col1:
            st.metric(
                "总员工记录",
                total
            )


        with col2:
            st.metric(
                "成功导入",
                success
            )


        with col3:
            st.metric(
                "异常记录",
                errors
            )


        with col4:
            st.metric(
                "数据质量",
                f"{rate:.1f}%"
            )

    # =====================================================
    # AI Error Analysis
    # =====================================================

    if error_df is None:

        st.info(
            "请先在“③ 数据校验”运行数据校验。"
        )

    elif error_df.empty:

        st.success(
            "✅ 数据校验通过，未发现异常记录。"
        )

        st.info(
            "AI 错误分析未触发（当前无异常数据）。"
        )

    else:

        if demo_mode:

            st.info(
                "🧪 Demo 模式："
                "使用规则化错误解释，不调用 API。"
            )

            ai_error_df = demo_error_analysis(
                error_df
            )

            st.session_state[
                "ai_error_df"
            ] = ai_error_df

        else:

            if not real_mode_authorized:

                st.warning(
                    "请先解锁 Real Mode。"
                )

            else:

                if st.button(
                    "🧠 使用 DeepSeek 分析错误",
                    type="primary",
                    width="stretch"
                ):

                    with st.spinner(
                        "正在调用 DeepSeek..."
                    ):

                        try:

                            ai_error_df = (
                                ai_analyze_errors(

                                    error_df,

                                    hris_schema

                                )
                            )

                            st.session_state[
                                "ai_error_df"
                            ] = ai_error_df

                            st.success(
                                "AI 错误分析完成。"
                            )

                        except Exception as e:

                            st.error(
                                f"AI 错误分析失败：{e}"
                            )

    # =====================================================
    # Display AI Error Analysis
    # =====================================================

    ai_error_df = st.session_state[
        "ai_error_df"
    ]

    if ai_error_df is not None:

        st.subheader(
            "错误分析"
        )

        for _, row in (
            ai_error_df
            .iterrows()
        ):

            title = (
                f"{row.get('error_code', '')}"
                f" · "
                f"{row.get('field', '')}"
            )

            with st.expander(
                title
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

                if str(
                    row.get(
                        "review_required",
                        True
                    )
                ).lower() == "true":

                    st.warning(
                        "需要人工确认"
                    )

    # =====================================================
    # Downloads
    # =====================================================

    if (
        hris_df is not None
        and error_df is not None
     ):


        st.subheader(
            "📦 输出文件"
        )

        st.caption(
            "以下文件可直接用于 HRIS 数据导入与质量审核。"
        )

        col1, col2, col3 = (
            st.columns(3)
        )

        # -----------------------------------------------
        # HRIS Excel
        # -----------------------------------------------

        with col1:

            st.download_button(

                "⬇️ HRIS 导入文件",

                data=dataframe_to_excel_bytes(
                    hris_df,
                    "HRIS_Import"
                ),

                file_name="hris_import.xlsx",

                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),

                width="stretch"
            )

        # -----------------------------------------------
        # Error Excel
        # -----------------------------------------------

        with col2:

            st.download_button(

                "⬇️ 错误报告",

                data=dataframe_to_excel_bytes(
                    error_df,
                    "Errors"
                ),

                file_name="sync_errors.xlsx",

                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),

                width="stretch"
            )

        # -----------------------------------------------
        # AI Error Excel
        # -----------------------------------------------

        if ai_error_df is not None:

            with col3:

                st.download_button(

                    "⬇️ AI 错误分析",

                    data=dataframe_to_excel_bytes(
                        ai_error_df,
                        "AI_Error_Analysis"
                    ),

                    file_name=(
                        "ai_error_analysis.xlsx"
                    ),

                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),

                    width="stretch"
                )

        else:

            with col3:

                st.info(
                    "无异常，因此未生成 AI 分析报告"
                )