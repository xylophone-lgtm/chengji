import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO

# ============================================================
# 🚀 页面配置 —— 中二标题
# ============================================================
st.set_page_config(page_title="超·成绩解析领域", layout="wide")
st.markdown("""
    <h1 style='text-align: center; font-size: 3.5rem; font-weight: 900; 
               background: linear-gradient(135deg, #ff6b6b, #ffd93d, #6bcb77, #4d96ff);
               -webkit-background-clip: text; -webkit-text-fill-color: transparent;'>
        ⚡ 超·成绩解析领域 ⚡
    </h1>
    <p style='text-align: center; font-size: 1.2rem; color: #888; margin-top: -10px;'>
        全维度成绩统御系统 · 高中
    </p>
""", unsafe_allow_html=True)

# ============================================================
# 默认分数线
# ============================================================
DEFAULT_HIGH_SCORE = 515.0          # 自招高线
DEFAULT_LOW_SCORE = 505.0           # 自招低线
DEFAULT_BENCHMARK_HIGH = 425.0      # 本科高线
DEFAULT_BENCHMARK_LOW = 415.0       # 本科低线

# 自招单科线
DEFAULT_SUBJECT_LINES = {
    "语文": 93.5,
    "数学": 108.0,
    "英语": 83.0,
    "物理": 73.0,
    "化学": 73.0,
    "生物": 73.0,
    "历史": 73.0,
    "地理": 73.0,
    "政治": 73.0,
}

# 本科单科线
DEFAULT_BENCHMARK_SUBJECT_LINES = {
    "语文": 80.0,
    "数学": 80.0,
    "英语": 57.0,
    "物理": 61.0,
    "化学": 61.0,
    "生物": 61.0,
    "历史": 61.0,
    "地理": 61.0,
    "政治": 61.0,
}

ALL_SUBJECTS = list(DEFAULT_SUBJECT_LINES.keys())

# 定科组合（用于计算班级总分平均分，仅导出用）
FIXED_SUBJECTS = {
    "1班": ["历史", "政治"],
    "2班": ["历史", "政治"],
    "3班": ["历史", "政治"],
    "4班": ["地理", "历史"],
    "5班": ["地理", "政治"],
    "6班": ["生物", "物理"],
    "7班": ["生物", "化学"],
    "8班": ["生物", "化学"],
    "9班": ["生物", "物理"],
    "10班": ["物理", "化学"],
    "11班": ["物理", "化学"],
    "12班": ["物理", "化学"],
    "13班": ["物理", "化学"],
    "14班": ["物理", "化学"],
    "15班": ["物理", "化学"],
    "16班": ["物理", "化学"],
    "17班": ["物理", "化学"],
    "18班": ["物理", "化学"],
}


# ============================================================
# 数据读取
# ============================================================
@st.cache_data
def load_data(uploaded_file):
    """读取总表，自动识别班级列、总分列、走班列"""
    try:
        df = pd.read_excel(uploaded_file, sheet_name=0)
        df = df.dropna(how="all")
        
        # ---------- 自动识别班级列 ----------
        class_col = None
        for col in df.columns:
            if "班级" in str(col) or "班" in str(col) or "class" in str(col).lower():
                class_col = col
                break
        if class_col is None:
            st.error(f"找不到班级列，当前列名：{list(df.columns)}")
            return None, None, None
        df.rename(columns={class_col: "班级"}, inplace=True)
        
        # 清理班级列（提取数字）
        df["班级"] = df["班级"].astype(str).str.replace("班", "").str.strip()
        df["班级"] = pd.to_numeric(df["班级"], errors="coerce")
        df = df[df["班级"].between(1, 18)]
        
        # ---------- 自动识别走班列 ----------
        zb_col = None
        for col in df.columns:
            if "走班" in str(col):
                zb_col = col
                break
        if zb_col:
            df.rename(columns={zb_col: "走班"}, inplace=True)
        
        # ---------- 自动识别总分列（优先赋分总分） ----------
        total_col = None
        for col in df.columns:
            if "总分（折）" in str(col) or "赋分总分" in str(col):
                total_col = col
                break
        if total_col is None:
            for col in df.columns:
                if "总分(折)" in str(col):
                    total_col = col
                    break
        if total_col is None:
            for col in df.columns:
                if "总分" in str(col):
                    total_col = col
                    break
        if total_col is None:
            st.error(f"找不到总分列，当前列名：{list(df.columns)}")
            return None, None, None
        df.rename(columns={total_col: "总分（折）"}, inplace=True)
        total_col = "总分（折）"
        
        # ---------- 清洗数据 ----------
        for subj in ALL_SUBJECTS:
            if subj in df.columns:
                df[subj] = pd.to_numeric(df[subj], errors='coerce')
        df[total_col] = pd.to_numeric(df[total_col], errors='coerce')
        df = df.dropna(subset=[total_col])  # 缺考者不计入任何统计
        
        df = df.sort_values(["班级", total_col], ascending=[True, False])
        return df, total_col, (zb_col is not None)
    except Exception as e:
        st.error(f"读取文件失败：{e}")
        return None, None, None


# ============================================================
# 预览指标计算（行政班或走班班通用）
# ============================================================
def calc_preview_metrics(df_class, high_score, low_score, benchmark_high, benchmark_low,
                          subject_lines, bench_subject_lines, total_col):
    """计算单个班级（行政或走班）的预览指标"""
    # 防御性转换
    for subj in ALL_SUBJECTS:
        if subj in df_class.columns:
            df_class[subj] = pd.to_numeric(df_class[subj], errors='coerce')
    df_class[total_col] = pd.to_numeric(df_class[total_col], errors='coerce')
    df_class = df_class.dropna(subset=[total_col])
    
    total_students = len(df_class)
    high_total = (df_class[total_col] >= high_score).sum()
    low_total = (df_class[total_col] >= low_score).sum()
    bench_high_total = (df_class[total_col] >= benchmark_high).sum()
    
    metrics = {
        "avg": df_class[total_col].mean(),
        "high_total": high_total,
        "low_total": low_total,
        "bench_high_total": bench_high_total,
        "high_rate_total": high_total / total_students if total_students > 0 else 0,
        "bench_rate_total": bench_high_total / total_students if total_students > 0 else 0,
        "high_contrib": {},
        "high_rate": {},
        "bench_contrib": {},
        "bench_rate": {},
    }
    for subj in ALL_SUBJECTS:
        if subj not in df_class.columns:
            metrics["high_contrib"][subj] = 0
            metrics["high_rate"][subj] = 0
            metrics["bench_contrib"][subj] = 0
            metrics["bench_rate"][subj] = 0
            continue
        # 自招有效贡献率
        total = (df_class[subj] >= subject_lines[subj]).sum()
        contrib = ((df_class[subj] >= subject_lines[subj]) & (df_class[total_col] >= high_score)).sum()
        metrics["high_contrib"][subj] = contrib
        metrics["high_rate"][subj] = contrib / total if total > 0 else 0
        # 本科有效贡献率
        total_b = (df_class[subj] >= bench_subject_lines[subj]).sum()
        contrib_b = ((df_class[subj] >= bench_subject_lines[subj]) & (df_class[total_col] >= benchmark_high)).sum()
        metrics["bench_contrib"][subj] = contrib_b
        metrics["bench_rate"][subj] = contrib_b / total_b if total_b > 0 else 0
    return metrics


# ============================================================
# 导出Excel生成函数（41个Sheet）
# ============================================================
def generate_excel_report(df, total_col, high_score, low_score, benchmark_high, benchmark_low,
                          subject_lines, bench_subject_lines):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')

    # 1-18 Sheet：行政班
    for cls in range(1, 19):
        df_cls = df[df["班级"] == cls].copy()
        if len(df_cls) == 0:
            continue
        df_cls.to_excel(writer, sheet_name=f"{cls}班", index=False)

    # 19 Sheet：过线统计（含5张表）
    df_stats = create_stats_sheet(df, total_col, high_score, low_score, benchmark_high, benchmark_low,
                                   subject_lines, bench_subject_lines)
    df_stats.to_excel(writer, sheet_name="过线统计", index=False)

    # 20+ Sheet：走班班
    zb_sheets = create_zb_sheets(df, total_col)
    for name, df_zb in zb_sheets.items():
        df_zb.to_excel(writer, sheet_name=name, index=False)

    # 班级各科平均分
    df_class_avg = create_class_avg_sheet(df, total_col)
    df_class_avg.to_excel(writer, sheet_name="班级各科平均分", index=False)

    # 走班学科均分
    df_zb_avg = create_zb_avg_sheet(df, total_col, high_score, low_score, benchmark_high, benchmark_low,
                                     subject_lines, bench_subject_lines)
    df_zb_avg.to_excel(writer, sheet_name="走班学科均分", index=False)

    writer.close()
    output.seek(0)
    return output


def create_stats_sheet(df, total_col, high_score, low_score, benchmark_high, benchmark_low,
                       subject_lines, bench_subject_lines):
    rows = []
    for cls in range(1, 19):
        df_cls = df[df["班级"] == cls]
        row = {"班级": f"{cls}班", "有效参考人数": len(df_cls)}
        # 总分过线人数
        row[f"自招高线(≥{high_score})"] = (df_cls[total_col] >= high_score).sum()
        row[f"自招低线(≥{low_score})"] = (df_cls[total_col] >= low_score).sum()
        row[f"本科高线(≥{benchmark_high})"] = (df_cls[total_col] >= benchmark_high).sum()
        row[f"本科低线(≥{benchmark_low})"] = (df_cls[total_col] >= benchmark_low).sum()
        # 各科过自招线人数
        for subj in ALL_SUBJECTS:
            if subj in df_cls.columns:
                row[f"{subj}过自招"] = (df_cls[subj] >= subject_lines[subj]).sum()
        # 各科过本科线人数
        for subj in ALL_SUBJECTS:
            if subj in df_cls.columns:
                row[f"{subj}过本科"] = (df_cls[subj] >= bench_subject_lines[subj]).sum()
        # 自招有效贡献率
        for subj in ALL_SUBJECTS:
            if subj in df_cls.columns:
                total_online = (df_cls[subj] >= subject_lines[subj]).sum()
                contrib = ((df_cls[subj] >= subject_lines[subj]) & (df_cls[total_col] >= high_score)).sum()
                row[f"{subj}自招贡献率"] = contrib / total_online if total_online > 0 else 0
        # 本科有效贡献率
        for subj in ALL_SUBJECTS:
            if subj in df_cls.columns:
                total_online = (df_cls[subj] >= bench_subject_lines[subj]).sum()
                contrib = ((df_cls[subj] >= bench_subject_lines[subj]) & (df_cls[total_col] >= benchmark_high)).sum()
                row[f"{subj}本科贡献率"] = contrib / total_online if total_online > 0 else 0
        rows.append(row)
    return pd.DataFrame(rows)


def create_zb_sheets(df, total_col):
    if "走班" not in df.columns:
        return {}
    zb_dict = {}
    for zb_name in df["走班"].unique():
        if pd.isna(zb_name) or zb_name == "":
            continue
        df_zb = df[df["走班"] == zb_name].copy()
        df_zb = df_zb.sort_values(total_col, ascending=False)
        zb_dict[zb_name] = df_zb
    return zb_dict


def create_class_avg_sheet(df, total_col):
    rows = []
    for cls in range(1, 19):
        df_cls = df[df["班级"] == cls]
        row = {"班级": f"{cls}班", "有效人数": len(df_cls)}
        for subj in ALL_SUBJECTS:
            if subj in df_cls.columns:
                row[f"{subj}均分"] = df_cls[subj].mean()
        row["总分均分"] = df_cls[total_col].mean()
        rows.append(row)
    return pd.DataFrame(rows)


def create_zb_avg_sheet(df, total_col, high_score, low_score, benchmark_high, benchmark_low,
                        subject_lines, bench_subject_lines):
    if "走班" not in df.columns:
        return pd.DataFrame()
    rows = []
    for zb_name in df["走班"].unique():
        if pd.isna(zb_name) or zb_name == "":
            continue
        df_zb = df[df["走班"] == zb_name]
        row = {"走班班": zb_name, "人数": len(df_zb)}
        # 识别走班对应的学科
        subject = None
        for subj in ["物理", "化学", "生物", "历史", "地理", "政治"]:
            if subj in zb_name:
                subject = subj
                break
        if subject and subject in df_zb.columns:
            row[f"{subject}赋分均分"] = df_zb[subject].mean()
            raw_col = f"{subject}原始分"
            if raw_col in df_zb.columns:
                row[f"{subject}原始分均分"] = df_zb[raw_col].mean()
            # 达线人数与有效贡献率
            row["自招达线人数"] = (df_zb[subject] >= subject_lines[subject]).sum()
            total_zb = (df_zb[subject] >= subject_lines[subject]).sum()
            contrib_zb = ((df_zb[subject] >= subject_lines[subject]) & (df_zb[total_col] >= high_score)).sum()
            row["自招有效贡献率"] = contrib_zb / total_zb if total_zb > 0 else 0
            row["本科达线人数"] = (df_zb[subject] >= bench_subject_lines[subject]).sum()
            total_zb_b = (df_zb[subject] >= bench_subject_lines[subject]).sum()
            contrib_zb_b = ((df_zb[subject] >= bench_subject_lines[subject]) & (df_zb[total_col] >= benchmark_high)).sum()
            row["本科有效贡献率"] = contrib_zb_b / total_zb_b if total_zb_b > 0 else 0
        rows.append(row)
    return pd.DataFrame(rows)


# ============================================================
# 获取走班班列表
# ============================================================
def get_zb_list(df):
    if "走班" not in df.columns:
        return []
    zb_list = []
    for name in df["走班"].unique():
        if pd.isna(name) or name == "":
            continue
        zb_list.append(name)
    return sorted(zb_list)


# ============================================================
# 侧边栏
# ============================================================
with st.sidebar:
    st.header("📁 数据上传")
    uploaded_file = st.file_uploader("上传 Excel 成绩表", type=["xls", "xlsx"])

    if uploaded_file is not None:
        df, total_col, has_zb = load_data(uploaded_file)
        if df is not None:
            st.success(f"✅ 读取成功！共 {len(df)} 名学生")
            
            # 分析模式切换
            mode = st.radio("分析模式", ["行政班", "走班班"], horizontal=True)
            
            if mode == "行政班":
                class_list = sorted(df["班级"].unique().astype(int))
                class_options = [f"{int(c)}班" for c in class_list]
                selected_name = st.selectbox("选择班级", class_options)
                selected_class_num = int(selected_name.replace("班", ""))
                df_selected = df[df["班级"] == selected_class_num].copy()
                display_name = selected_name
            else:
                if has_zb:
                    zb_list = get_zb_list(df)
                    if zb_list:
                        selected_zb = st.selectbox("选择走班班", zb_list)
                        df_selected = df[df["走班"] == selected_zb].copy()
                        display_name = selected_zb
                    else:
                        st.warning("数据中无走班信息")
                        st.stop()
                else:
                    st.warning("未找到'走班'列，请检查数据")
                    st.stop()
            
            if len(df_selected) == 0:
                st.warning("该班级无数据")
                st.stop()
        else:
            st.stop()
    else:
        st.info("请上传成绩文件开始分析")
        st.stop()

    st.markdown("---")
    st.header("⚙️ 分数线设置")

    high_score = st.number_input("自招高线", value=DEFAULT_HIGH_SCORE, step=1.0)
    low_score = st.number_input("自招低线", value=DEFAULT_LOW_SCORE, step=1.0)
    benchmark_high = st.number_input("本科高线", value=DEFAULT_BENCHMARK_HIGH, step=1.0)
    benchmark_low = st.number_input("本科低线", value=DEFAULT_BENCHMARK_LOW, step=1.0)

    st.subheader("自招单科线")
    subject_lines = {}
    cols = st.columns(3)
    for i, subj in enumerate(ALL_SUBJECTS):
        col = cols[i % 3]
        subject_lines[subj] = col.number_input(f"{subj}", value=DEFAULT_SUBJECT_LINES[subj], step=0.5)

    st.subheader("本科单科线")
    bench_subject_lines = {}
    cols = st.columns(3)
    for i, subj in enumerate(ALL_SUBJECTS):
        col = cols[i % 3]
        bench_subject_lines[subj] = col.number_input(f"{subj}", value=DEFAULT_BENCHMARK_SUBJECT_LINES[subj], step=0.5)

    st.markdown("---")
    if st.button("📥 导出完整报表（41个Sheet）"):
        with st.spinner("正在生成报表..."):
            excel_data = generate_excel_report(df, total_col, high_score, low_score,
                                                benchmark_high, benchmark_low,
                                                subject_lines, bench_subject_lines)
            st.download_button(
                label="📥 下载报表",
                data=excel_data,
                file_name="成绩分析报表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


# ============================================================
# 主区域预览
# ============================================================
if uploaded_file is not None and df is not None:
    metrics = calc_preview_metrics(df_selected, high_score, low_score, benchmark_high, benchmark_low,
                                    subject_lines, bench_subject_lines, total_col)

    st.subheader(f"⚡ {display_name} 成绩分析报告")

    col1, col2, col3, col4, col5, col6 = st.columns(6)
    col1.metric("平均分", f"{metrics['avg']:.2f}")
    col2.metric(f"自招高线达线率", f"{metrics['high_rate_total']*100:.1f}%")
    col3.metric(f"自招高线人数", metrics['high_total'])
    col4.metric(f"自招低线人数", metrics['low_total'])
    col5.metric(f"本科高线达线率", f"{metrics['bench_rate_total']*100:.1f}%")
    col6.metric(f"本科高线人数", metrics['bench_high_total'])

    st.markdown("---")

    # 自招有效贡献
    st.subheader(f"📊 自招线（≥{high_score:.0f}分）有效贡献")
    high_df = pd.DataFrame({
        "科目": ALL_SUBJECTS,
        "贡献人数": [metrics["high_contrib"][s] for s in ALL_SUBJECTS],
        "有效贡献率": [f"{metrics['high_rate'][s]*100:.1f}%" for s in ALL_SUBJECTS],
    })
    st.table(high_df.style.hide(axis="index"))

    # 本科有效贡献
    st.subheader(f"📊 本科线（≥{benchmark_high:.0f}分）有效贡献")
    bench_df = pd.DataFrame({
        "科目": ALL_SUBJECTS,
        "贡献人数": [metrics["bench_contrib"][s] for s in ALL_SUBJECTS],
        "有效贡献率": [f"{metrics['bench_rate'][s]*100:.1f}%" for s in ALL_SUBJECTS],
    })
    st.table(bench_df.style.hide(axis="index"))

    # 学生明细
    with st.expander("📋 查看本班学生明细"):
        display_cols = ["姓名"] + ALL_SUBJECTS + [total_col]
        existing_cols = [c for c in display_cols if c in df_selected.columns]
        st.dataframe(df_selected[existing_cols], use_container_width=True)

    st.caption("💡 修改左侧分数线后，所有数据自动更新。点击'导出完整报表'可下载41个Sheet的Excel文件。")
