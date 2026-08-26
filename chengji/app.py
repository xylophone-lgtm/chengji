import streamlit as st
import pandas as pd
import numpy as np
import re

st.set_page_config(page_title="高二成绩分析系统", layout="wide")
st.title("📊 高二年级成绩分析系统")

DEFAULT_HIGH_SCORE = 505.0
DEFAULT_LOW_SCORE = 495.0
DEFAULT_BENCHMARK = 400.0

DEFAULT_SUBJECT_LINES = {
    "语文": 92.5,
    "数学": 114.0,
    "英语": 78.8,
    "物理": 73.0,
    "化学": 73.0,
    "生物": 73.0,
    "历史": 73.0,
    "地理": 73.0,
    "政治": 73.0,
}
ALL_SUBJECTS = list(DEFAULT_SUBJECT_LINES.keys())

@st.cache_data
def load_all_data(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names
    admin_sheet = None
    zb_sheets = {}
    for name in sheet_names:
        if "总表" in name:
            admin_sheet = name
            break
    if admin_sheet is None:
        admin_sheet = sheet_names[0]
    zb_pattern = re.compile(r"(物理|化学|生物|历史|地理|政治)走(\d+)班")
    for name in sheet_names:
        match = zb_pattern.match(name)
        if match:
            zb_sheets[name] = match.groups()
    try:
        df_admin = pd.read_excel(uploaded_file, sheet_name=admin_sheet)
        df_admin = df_admin.dropna(how="all")
        if "班级" not in df_admin.columns:
            st.error("Excel中找不到'班级'列，请检查格式")
            return None, {}, {}
        df_admin["班级_num"] = df_admin["班级"].astype(str).str.replace("班", "").str.strip()
        df_admin["班级_num"] = pd.to_numeric(df_admin["班级_num"], errors="coerce")
        df_admin = df_admin[df_admin["班级_num"].between(1, 18)]
        df_admin = df_admin.sort_values(["班级_num", "姓名"])
    except Exception as e:
        st.error(f"读取行政班失败：{e}")
        df_admin = None
    df_zb_dict = {}
    for name in zb_sheets:
        try:
            df = pd.read_excel(uploaded_file, sheet_name=name)
            df = df.dropna(how="all")
            df_zb_dict[name] = df
        except:
            df_zb_dict[name] = None
    return df_admin, df_zb_dict, zb_sheets

def calc_metrics(df_class, high_score, low_score, benchmark, subject_lines):
    possible_total_cols = ["总分（折）", "总分(折)", "赋分总分", "总分", "总分（赋分）"]
    total_col = None
    for col in possible_total_cols:
        if col in df_class.columns:
            total_col = col
            break
    if total_col is None:
        return None
    high_total = (df_class[total_col] >= high_score).sum()
    low_total = (df_class[total_col] >= low_score).sum()
    avg_total = df_class[total_col].mean()
    def calc_contrib(subject, line_score, total_line):
        if subject not in df_class.columns:
            return 0, 0
        mask = (df_class[subject] >= line_score) & (df_class[total_col] >= total_line)
        contrib = mask.sum()
        total_online = (df_class[total_col] >= total_line).sum()
        rate = contrib / total_online if total_online > 0 else 0
        return contrib, rate
    high_contrib, high_rate = {}, {}
    low_contrib, low_rate = {}, {}
    bench_contrib = {}
    for subj in ALL_SUBJECTS:
        c, r = calc_contrib(subj, subject_lines[subj], high_score)
        high_contrib[subj] = c
        high_rate[subj] = r
        c, r = calc_contrib(subj, subject_lines[subj], low_score)
        low_contrib[subj] = c
        low_rate[subj] = r
        if subj in df_class.columns:
            mask = (df_class[subj] >= subject_lines[subj]) & (df_class[total_col] >= benchmark)
            bench_contrib[subj] = mask.sum()
        else:
            bench_contrib[subj] = 0
    return {
        "avg": avg_total,
        "high_total": high_total,
        "low_total": low_total,
        "high_contrib": high_contrib,
        "high_rate": high_rate,
        "low_contrib": low_contrib,
        "low_rate": low_rate,
        "bench_contrib": bench_contrib,
        "total_col": total_col,
    }

with st.sidebar:
    st.header("📁 数据上传")
    uploaded_file = st.file_uploader("上传 Excel 成绩表", type=["xls", "xlsx"])
    if uploaded_file is not None:
        df_admin, df_zb_dict, zb_sheets = load_all_data(uploaded_file)
        if df_admin is not None or df_zb_dict:
            st.success("✅ 数据读取成功！")
        else:
            st.error("❌ 读取失败，请检查文件格式")
            st.stop()
    else:
        st.info("请上传成绩文件")
        st.stop()
    st.markdown("---")
    st.header("📂 选择分析对象")
    analysis_type = st.radio("分析类型", ["行政班", "走班班"])
    if analysis_type == "行政班":
        if df_admin is not None and len(df_admin) > 0:
            class_list = sorted(df_admin["班级_num"].unique().astype(int))
            class_options = [f"{int(c)}班" for c in class_list]
            selected_name = st.selectbox("选择班级", class_options)
            selected_class_num = int(selected_name.replace("班", ""))
            df_selected = df_admin[df_admin["班级_num"] == selected_class_num].copy()
            display_name = selected_name
        else:
            st.warning("无行政班数据")
            st.stop()
    else:
        if df_zb_dict:
            zb_options = list(df_zb_dict.keys())
            selected_zb = st.selectbox("选择走班班", zb_options)
            df_selected = df_zb_dict[selected_zb].copy()
            display_name = selected_zb
        else:
            st.warning("无走班班数据")
            st.stop()
    if df_selected is None or len(df_selected) == 0:
        st.warning("该班级无数据")
        st.stop()
    st.markdown("---")
    st.header("⚙️ 分数线设置")
    high_score = st.number_input("自招高线", value=DEFAULT_HIGH_SCORE, step=1.0)
    low_score = st.number_input("自招低线", value=DEFAULT_LOW_SCORE, step=1.0)
    benchmark = st.number_input("本科线", value=DEFAULT_BENCHMARK, step=1.0)
    st.subheader("各科单科线")
    subject_lines = {}
    cols = st.columns(3)
    for i, subj in enumerate(ALL_SUBJECTS):
        col = cols[i % 3]
        subject_lines[subj] = col.number_input(f"{subj}", value=DEFAULT_SUBJECT_LINES[subj], step=0.5)

if uploaded_file is not None:
    metrics = calc_metrics(df_selected, high_score, low_score, benchmark, subject_lines)
    if metrics is None:
        st.error("找不到总分列，请检查Excel中是否包含：总分（折）、总分(折)、赋分总分等列名")
        st.stop()
    st.subheader(f"📈 {display_name} 成绩分析报告")
    col1, col2, col3 = st.columns(3)
    col1.metric("班级平均分", f"{metrics['avg']:.2f}")
    col2.metric(f"高线上线人数（≥{high_score:.0f}分）", metrics['high_total'])
    col3.metric(f"低线上线人数（≥{low_score:.0f}分）", metrics['low_total'])
    st.markdown("---")
    st.subheader(f"📊 自招高线（≥{high_score:.0f}分）有效贡献")
    high_df = pd.DataFrame({
        "科目": ALL_SUBJECTS,
        "贡献人数": [metrics["high_contrib"][s] for s in ALL_SUBJECTS],
        "有效率": [f"{metrics['high_rate'][s]*100:.1f}%" for s in ALL_SUBJECTS],
    })
    if metrics['high_total'] == 0:
        high_df["有效率"] = "-"
    st.table(high_df.style.hide(axis="index"))
    st.subheader(f"📊 自招低线（≥{low_score:.0f}分）有效贡献")
    low_df = pd.DataFrame({
        "科目": ALL_SUBJECTS,
        "贡献人数": [metrics["low_contrib"][s] for s in ALL_SUBJECTS],
        "有效率": [f"{metrics['low_rate'][s]*100:.1f}%" for s in ALL_SUBJECTS],
    })
    if metrics['low_total'] == 0:
        low_df["有效率"] = "-"
    st.table(low_df.style.hide(axis="index"))
    st.subheader(f"📊 本科（≥{benchmark:.0f}分）贡献人数")
    bench_df = pd.DataFrame({
        "科目": ALL_SUBJECTS,
        "贡献人数": [metrics["bench_contrib"][s] for s in ALL_SUBJECTS],
    })
    st.table(bench_df.style.hide(axis="index"))
    with st.expander("📋 查看本班学生明细"):
        display_cols = ["姓名", "语文", "数学", "英语", "物理", "化学", "生物", "历史", "地理", "政治", metrics['total_col']]
        existing_cols = [c for c in display_cols if c in df_selected.columns]
        st.dataframe(df_selected[existing_cols], use_container_width=True)
    st.caption("💡 修改左侧分数线后，所有数据将自动更新。")
