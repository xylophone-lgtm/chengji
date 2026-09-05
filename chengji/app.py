import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO

# ============================================================
# 页面配置
# ============================================================
st.set_page_config(page_title="高二成绩分析系统", layout="wide")
st.title("📊 高二年级成绩分析系统")

# ============================================================
# 默认分数线
# ============================================================
DEFAULT_HIGH_SCORE = 515.0
DEFAULT_LOW_SCORE = 505.0
DEFAULT_BENCHMARK_HIGH = 425.0
DEFAULT_BENCHMARK_LOW = 415.0

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

# 定科组合
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
    try:
        df = pd.read_excel(uploaded_file, sheet_name=0)
        df = df.dropna(how="all")
        
        # ----- 自动识别班级列 -----
        class_col = None
        for col in df.columns:
            if "班级" in str(col) or "班" in str(col) or "class" in str(col).lower():
                class_col = col
                break
        if class_col is None:
            st.error(f"找不到班级列，当前列名：{list(df.columns)}")
            return None, None
        df.rename(columns={class_col: "班级"}, inplace=True)
        
        # 清理班级列（提取数字）
        df["班级"] = df["班级"].astype(str).str.replace("班", "").str.strip()
        df["班级"] = pd.to_numeric(df["班级"], errors="coerce")
        df = df[df["班级"].between(1, 18)]
        
        # ----- 自动识别总分列 -----
        total_col = None
        for col in df.columns:
            if "总分" in str(col) or "赋分总分" in str(col):
                total_col = col
                break
        if total_col is None:
            st.error(f"找不到总分列，当前列名：{list(df.columns)}")
            return None, None
        df.rename(columns={total_col: "总分（折）"}, inplace=True)
        total_col = "总分（折）"
        
        # ========== 关键：清洗所有科目和总分，转换为数字 ==========
        # 对所有科目列进行数值转换
        for subj in ALL_SUBJECTS:
            if subj in df.columns:
                df[subj] = pd.to_numeric(df[subj], errors='coerce')
        # 总分列也转换
        df[total_col] = pd.to_numeric(df[total_col], errors='coerce')
        
        # 删除总分缺失的行（缺考或无效数据）
        df = df.dropna(subset=[total_col])
        
        # 按班级和总分排序
        df = df.sort_values(["班级", total_col], ascending=[True, False])
        
        return df, total_col
    except Exception as e:
        st.error(f"读取文件失败：{e}")
        return None, None

# ============================================================
# 生成导出Excel（41个Sheet）
# ============================================================
def generate_excel_report(df, total_col, high_score, low_score, benchmark_high, benchmark_low,
                          subject_lines, bench_subject_lines):
    """生成完整的41个Sheet的Excel报表"""
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')

    # 1-18 Sheet：行政班
    for cls in range(1, 19):
        df_cls = df[df["班级"] == cls].copy()
        if len(df_cls) == 0:
            continue
        # 添加第六科成绩列
        df_cls = add_sixth_subject(df_cls)
        sheet_name = f"{cls}班"
        df_cls.to_excel(writer, sheet_name=sheet_name, index=False)

    # 19 Sheet：过线统计
    df_stats = create_stats_sheet(df, high_score, low_score, benchmark_high, benchmark_low,
                                   subject_lines, bench_subject_lines, total_col)
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


def add_sixth_subject(df):
    """为每个学生添加第六科成绩"""
    # 根据定科组合，找到走班科
    # 简化版：从走班列提取
    if "走班" in df.columns:
        df["第六科"] = df["走班"]
    return df


def create_stats_sheet(df, high_score, low_score, benchmark_high, benchmark_low,
                       subject_lines, bench_subject_lines, total_col):
    """生成过线统计表（5张表合并）"""
    rows = []
    for cls in range(1, 19):
        df_cls = df[df["班级"] == cls]
        row = {"班级": f"{cls}班", "有效参考人数": len(df_cls)}
        # 总分过线人数
        row["特招高线(≥{})".format(high_score)] = (df_cls[total_col] >= high_score).sum()
        row["特招低线(≥{})".format(low_score)] = (df_cls[total_col] >= low_score).sum()
        row["本科高线(≥{})".format(benchmark_high)] = (df_cls[total_col] >= benchmark_high).sum()
        row["本科低线(≥{})".format(benchmark_low)] = (df_cls[total_col] >= benchmark_low).sum()
        # 各科过特招线人数
        for subj in ALL_SUBJECTS:
            if subj in df_cls.columns:
                row[f"{subj}过特招"] = (df_cls[subj] >= subject_lines[subj]).sum()
        # 各科过本科线人数
        for subj in ALL_SUBJECTS:
            if subj in df_cls.columns:
                row[f"{subj}过本科"] = (df_cls[subj] >= bench_subject_lines[subj]).sum()
        # 特招有效贡献率
        for subj in ALL_SUBJECTS:
            if subj in df_cls.columns:
                total_online = (df_cls[subj] >= subject_lines[subj]).sum()
                contrib = ((df_cls[subj] >= subject_lines[subj]) & (df_cls[total_col] >= high_score)).sum()
                row[f"{subj}特招贡献率"] = contrib / total_online if total_online > 0 else 0
        # 本科有效贡献率
        for subj in ALL_SUBJECTS:
            if subj in df_cls.columns:
                total_online = (df_cls[subj] >= bench_subject_lines[subj]).sum()
                contrib = ((df_cls[subj] >= bench_subject_lines[subj]) & (df_cls[total_col] >= benchmark_high)).sum()
                row[f"{subj}本科贡献率"] = contrib / total_online if total_online > 0 else 0
        rows.append(row)
    return pd.DataFrame(rows)


def create_zb_sheets(df, total_col):
    """生成走班班Sheet"""
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
    """生成班级各科平均分"""
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
    """生成走班学科均分"""
    if "走班" not in df.columns:
        return pd.DataFrame()
    rows = []
    for zb_name in df["走班"].unique():
        if pd.isna(zb_name) or zb_name == "":
            continue
        df_zb = df[df["走班"] == zb_name]
        row = {"走班班": zb_name, "人数": len(df_zb)}
        # 识别该走班对应的学科
        subject = None
        for subj in ["物理", "化学", "生物", "历史", "地理", "政治"]:
            if subj in zb_name:
                subject = subj
                break
        if subject and subject in df_zb.columns:
            row[f"{subject}赋分均分"] = df_zb[subject].mean()
            # 原始分均分（如果有）
            raw_col = f"{subject}原始分"
            if raw_col in df_zb.columns:
                row[f"{subject}原始分均分"] = df_zb[raw_col].mean()
            # 达线人数
            row["特招达线人数"] = (df_zb[subject] >= subject_lines[subject]).sum()
            row["特招贡献率"] = ((df_zb[subject] >= subject_lines[subject]) & (df_zb[total_col] >= high_score)).sum() / max((df_zb[subject] >= subject_lines[subject]).sum(), 1)
            row["本科达线人数"] = (df_zb[subject] >= bench_subject_lines[subject]).sum()
            row["本科贡献率"] = ((df_zb[subject] >= bench_subject_lines[subject]) & (df_zb[total_col] >= benchmark_high)).sum() / max((df_zb[subject] >= bench_subject_lines[subject]).sum(), 1)
        rows.append(row)
    return pd.DataFrame(rows)


# ============================================================
# 计算预览指标
# ============================================================
def calc_preview_metrics(df_class, high_score, low_score, benchmark_high, benchmark_low,
                          subject_lines, bench_subject_lines, total_col):
    """网页预览用的指标计算"""
    # 防御性转换（以防万一）
    for subj in ALL_SUBJECTS:
        if subj in df_class.columns:
            df_class[subj] = pd.to_numeric(df_class[subj], errors='coerce')
    df_class[total_col] = pd.to_numeric(df_class[total_col], errors='coerce')
    df_class = df_class.dropna(subset=[total_col])
    
    metrics = {
        "avg": df_class[total_col].mean(),
        "high_total": (df_class[total_col] >= high_score).sum(),
        "low_total": (df_class[total_col] >= low_score).sum(),
        "bench_high_total": (df_class[total_col] >= benchmark_high).sum(),
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
        # 特招
        total = (df_class[subj] >= subject_lines[subj]).sum()
        contrib = ((df_class[subj] >= subject_lines[subj]) & (df_class[total_col] >= high_score)).sum()
        metrics["high_contrib"][subj] = contrib
        metrics["high_rate"][subj] = contrib / total if total > 0 else 0
        # 本科
        total_b = (df_class[subj] >= bench_subject_lines[subj]).sum()
        contrib_b = ((df_class[subj] >= bench_subject_lines[subj]) & (df_class[total_col] >= benchmark_high)).sum()
        metrics["bench_contrib"][subj] = contrib_b
        metrics["bench_rate"][subj] = contrib_b / total_b if total_b > 0 else 0
    return metrics


# ============================================================
# 侧边栏
# ============================================================
with st.sidebar:
    st.header("📁 数据上传")
    uploaded_file = st.file_uploader("上传 Excel 成绩表", type=["xls", "xlsx"])

    if uploaded_file is not None:
        df, total_col = load_data(uploaded_file)
        if df is not None:
            st.success(f"✅ 读取成功！共 {len(df)} 名学生")
            class_list = sorted(df["班级"].unique().astype(int))
            class_options = [f"{int(c)}班" for c in class_list]
            selected_class_str = st.selectbox("选择班级", class_options)
            selected_class_num = int(selected_class_str.replace("班", ""))
        else:
            st.stop()
    else:
        st.info("请上传成绩文件开始分析")
        st.stop()

    st.markdown("---")
    st.header("⚙️ 分数线设置")

    high_score = st.number_input("特招高线", value=515.0, step=1.0)
    low_score = st.number_input("特招低线", value=505.0, step=1.0)
    benchmark_high = st.number_input("本科高线", value=425.0, step=1.0)
    benchmark_low = st.number_input("本科低线", value=415.0, step=1.0)

    st.subheader("特招单科线")
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
# 主区域：预览
# ============================================================
if uploaded_file is not None and df is not None:
    df_class = df[df["班级"] == selected_class_num].copy()
    if len(df_class) == 0:
        st.warning("该班级无数据")
        st.stop()

    metrics = calc_preview_metrics(df_class, high_score, low_score, benchmark_high, benchmark_low,
                                    subject_lines, bench_subject_lines, total_col)

    st.subheader(f"📈 {selected_class_str} 成绩分析报告")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("班级平均分", f"{metrics['avg']:.2f}")
    col2.metric(f"特招高线（≥{high_score:.0f}）", metrics['high_total'])
    col3.metric(f"特招低线（≥{low_score:.0f}）", metrics['low_total'])
    col4.metric(f"本科高线（≥{benchmark_high:.0f}）", metrics['bench_high_total'])

    st.markdown("---")

    # 特招贡献
    st.subheader(f"📊 特招线（≥{high_score:.0f}分）有效贡献")
    high_df = pd.DataFrame({
        "科目": ALL_SUBJECTS,
        "贡献人数": [metrics["high_contrib"][s] for s in ALL_SUBJECTS],
        "贡献率": [f"{metrics['high_rate'][s]*100:.1f}%" for s in ALL_SUBJECTS],
    })
    st.table(high_df.style.hide(axis="index"))

    # 本科贡献
    st.subheader(f"📊 本科线（≥{benchmark_high:.0f}分）有效贡献")
    bench_df = pd.DataFrame({
        "科目": ALL_SUBJECTS,
        "贡献人数": [metrics["bench_contrib"][s] for s in ALL_SUBJECTS],
        "贡献率": [f"{metrics['bench_rate'][s]*100:.1f}%" for s in ALL_SUBJECTS],
    })
    st.table(bench_df.style.hide(axis="index"))

    # 学生明细
    with st.expander("📋 查看本班学生明细"):
        display_cols = ["姓名"] + ALL_SUBJECTS + [total_col]
        existing_cols = [c for c in display_cols if c in df_class.columns]
        st.dataframe(df_class[existing_cols], use_container_width=True)

    st.caption("💡 修改左侧分数线后，所有数据自动更新。点击'导出完整报表'可下载41个Sheet的Excel文件。")
