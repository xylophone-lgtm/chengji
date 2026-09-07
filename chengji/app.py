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
# 初始化 session_state
# ============================================================
if 'high_score' not in st.session_state:
    st.session_state.high_score = 515.499
if 'low_score' not in st.session_state:
    st.session_state.low_score = 505.499
if 'bench_high' not in st.session_state:
    st.session_state.bench_high = 425.499
if 'bench_low' not in st.session_state:
    st.session_state.bench_low = 415.499

# 自招高线单科线
if 'subject_high_lines' not in st.session_state:
    st.session_state.subject_high_lines = {
        "语文": 93.5, "数学": 108.0, "英语": 83.0,
        "物理": 73.0, "化学": 73.0, "生物": 73.0,
        "历史": 73.0, "地理": 73.0, "政治": 73.0,
    }
# 自招低线单科线
if 'subject_low_lines' not in st.session_state:
    st.session_state.subject_low_lines = {
        "语文": 92.5, "数学": 107.0, "英语": 82.0,
        "物理": 72.0, "化学": 72.0, "生物": 72.0,
        "历史": 72.0, "地理": 72.0, "政治": 72.0,
    }
# 本科高线单科线
if 'bench_subject_high_lines' not in st.session_state:
    st.session_state.bench_subject_high_lines = {
        "语文": 80.0, "数学": 80.0, "英语": 57.0,
        "物理": 61.0, "化学": 61.0, "生物": 61.0,
        "历史": 61.0, "地理": 61.0, "政治": 61.0,
    }
# 本科低线单科线
if 'bench_subject_low_lines' not in st.session_state:
    st.session_state.bench_subject_low_lines = {
        "语文": 79.0, "数学": 79.0, "英语": 56.0,
        "物理": 60.0, "化学": 60.0, "生物": 60.0,
        "历史": 60.0, "地理": 60.0, "政治": 60.0,
    }

ALL_SUBJECTS = list(st.session_state.subject_high_lines.keys())

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

        class_col = None
        for col in df.columns:
            if "班级" in str(col) or "班" in str(col) or "class" in str(col).lower():
                class_col = col
                break
        if class_col is None:
            st.error(f"找不到班级列，当前列名：{list(df.columns)}")
            return None, None, None
        df.rename(columns={class_col: "班级"}, inplace=True)

        df["班级"] = df["班级"].astype(str).str.replace("班", "").str.strip()
        df["班级"] = pd.to_numeric(df["班级"], errors="coerce")
        df = df[df["班级"].between(1, 18)]

        zb_col = None
        for col in df.columns:
            if "走班" in str(col):
                zb_col = col
                break
        if zb_col:
            df.rename(columns={zb_col: "走班"}, inplace=True)

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

        for subj in ALL_SUBJECTS:
            if subj in df.columns:
                df[subj] = pd.to_numeric(df[subj], errors='coerce')
        df[total_col] = pd.to_numeric(df[total_col], errors='coerce')
        df = df.dropna(subset=[total_col])

        df = df.sort_values(["班级", total_col], ascending=[True, False])
        return df, total_col, (zb_col is not None)
    except Exception as e:
        st.error(f"读取文件失败：{e}")
        return None, None, None


# ============================================================
# 导入分数线（支持高线/低线/本科线）
# ============================================================
def parse_score_file(uploaded_score_file):
    try:
        df = pd.read_excel(uploaded_score_file, header=0)
        if len(df) == 0:
            st.error("分数线文件为空")
            return None
        row = df.iloc[0].to_dict()
        result = {}

        # 总分线
        for key in ["自招高线", "自招低线", "本科高线", "本科低线"]:
            if key in row:
                result[key] = float(row[key])

        # 单科线：高线/低线/本科线
        # 对每个科目，尝试 "语文高线", "语文低线", "语文本科线" 等
        subject_lines = {}
        for subj in ALL_SUBJECTS:
            # 高线
            if f"{subj}高线" in row:
                subject_lines[f"{subj}_high"] = float(row[f"{subj}高线"])
            # 低线
            if f"{subj}低线" in row:
                subject_lines[f"{subj}_low"] = float(row[f"{subj}低线"])
            # 本科线（可能叫“语文本科线”或“语文本科线”）
            if f"{subj}本科线" in row:
                subject_lines[f"{subj}_bench"] = float(row[f"{subj}本科线"])
            elif f"{subj}本科" in row:
                subject_lines[f"{subj}_bench"] = float(row[f"{subj}本科"])
        result['subject_lines'] = subject_lines
        return result
    except Exception as e:
        st.error(f"解析分数线文件失败：{e}")
        return None


# ============================================================
# 预览指标计算
# ============================================================
def calc_preview_metrics(df_class, high_score, low_score, bench_high, bench_low,
                         subject_high_lines, subject_low_lines,
                         bench_subject_high, bench_subject_low, total_col):
    for subj in ALL_SUBJECTS:
        if subj in df_class.columns:
            df_class[subj] = pd.to_numeric(df_class[subj], errors='coerce')
    df_class[total_col] = pd.to_numeric(df_class[total_col], errors='coerce')
    df_class = df_class.dropna(subset=[total_col])

    total_students = len(df_class)
    high_total = (df_class[total_col] >= high_score).sum()
    low_total = (df_class[total_col] >= low_score).sum()
    bench_high_total = (df_class[total_col] >= bench_high).sum()
    bench_low_total = (df_class[total_col] >= bench_low).sum()

    metrics = {
        "avg": df_class[total_col].mean(),
        "high_total": high_total,
        "low_total": low_total,
        "bench_high_total": bench_high_total,
        "bench_low_total": bench_low_total,
        "high_rate_total": high_total / total_students if total_students > 0 else 0,
        "low_rate_total": low_total / total_students if total_students > 0 else 0,
        "bench_high_rate_total": bench_high_total / total_students if total_students > 0 else 0,
        "bench_low_rate_total": bench_low_total / total_students if total_students > 0 else 0,
        # 各科贡献数据（高/低/本科高/本科低）
        "high_contrib": {}, "high_rate": {},
        "low_contrib": {}, "low_rate": {},
        "bench_high_contrib": {}, "bench_high_rate": {},
        "bench_low_contrib": {}, "bench_low_rate": {},
    }
    for subj in ALL_SUBJECTS:
        if subj not in df_class.columns:
            # 置空
            for key in ["high_contrib", "high_rate", "low_contrib", "low_rate",
                        "bench_high_contrib", "bench_high_rate", "bench_low_contrib", "bench_low_rate"]:
                if "contrib" in key:
                    metrics[key][subj] = 0
                else:
                    metrics[key][subj] = 0
            continue
        # 自招高线
        total = (df_class[subj] >= subject_high_lines[subj]).sum()
        contrib = ((df_class[subj] >= subject_high_lines[subj]) & (df_class[total_col] >= high_score)).sum()
        metrics["high_contrib"][subj] = contrib
        metrics["high_rate"][subj] = contrib / total if total > 0 else 0
        # 自招低线
        total = (df_class[subj] >= subject_low_lines[subj]).sum()
        contrib = ((df_class[subj] >= subject_low_lines[subj]) & (df_class[total_col] >= low_score)).sum()
        metrics["low_contrib"][subj] = contrib
        metrics["low_rate"][subj] = contrib / total if total > 0 else 0
        # 本科高线
        total = (df_class[subj] >= bench_subject_high[subj]).sum()
        contrib = ((df_class[subj] >= bench_subject_high[subj]) & (df_class[total_col] >= bench_high)).sum()
        metrics["bench_high_contrib"][subj] = contrib
        metrics["bench_high_rate"][subj] = contrib / total if total > 0 else 0
        # 本科低线
        total = (df_class[subj] >= bench_subject_low[subj]).sum()
        contrib = ((df_class[subj] >= bench_subject_low[subj]) & (df_class[total_col] >= bench_low)).sum()
        metrics["bench_low_contrib"][subj] = contrib
        metrics["bench_low_rate"][subj] = contrib / total if total > 0 else 0

    return metrics


# ============================================================
# 导出Excel（包含所有线的统计）
# ============================================================
def generate_excel_report(df, total_col, high_score, low_score, bench_high, bench_low,
                          subject_high_lines, subject_low_lines,
                          bench_subject_high, bench_subject_low):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')

    # 行政班
    for cls in range(1, 19):
        df_cls = df[df["班级"] == cls].copy()
        if len(df_cls) == 0:
            continue
        df_cls.to_excel(writer, sheet_name=f"{cls}班", index=False)

    # 过线统计（含四套线）
    df_stats = create_stats_sheet(df, total_col, high_score, low_score, bench_high, bench_low,
                                   subject_high_lines, subject_low_lines,
                                   bench_subject_high, bench_subject_low)
    df_stats.to_excel(writer, sheet_name="过线统计", index=False)

    # 走班班
    zb_sheets = create_zb_sheets(df, total_col)
    for name, df_zb in zb_sheets.items():
        df_zb.to_excel(writer, sheet_name=name, index=False)

    # 班级各科平均分
    df_class_avg = create_class_avg_sheet(df, total_col)
    df_class_avg.to_excel(writer, sheet_name="班级各科平均分", index=False)

    # 走班学科均分（含四套线）
    df_zb_avg = create_zb_avg_sheet(df, total_col, high_score, low_score, bench_high, bench_low,
                                     subject_high_lines, subject_low_lines,
                                     bench_subject_high, bench_subject_low)
    df_zb_avg.to_excel(writer, sheet_name="走班学科均分", index=False)

    writer.close()
    output.seek(0)
    return output


def create_stats_sheet(df, total_col, high_score, low_score, bench_high, bench_low,
                       subject_high_lines, subject_low_lines,
                       bench_subject_high, bench_subject_low):
    rows = []
    for cls in range(1, 19):
        df_cls = df[df["班级"] == cls]
        row = {"班级": f"{cls}班", "有效参考人数": len(df_cls)}
        # 总分过线
        row[f"自招高线(≥{high_score})"] = (df_cls[total_col] >= high_score).sum()
        row[f"自招低线(≥{low_score})"] = (df_cls[total_col] >= low_score).sum()
        row[f"本科高线(≥{bench_high})"] = (df_cls[total_col] >= bench_high).sum()
        row[f"本科低线(≥{bench_low})"] = (df_cls[total_col] >= bench_low).sum()
        # 各科过线（高线/低线）
        for subj in ALL_SUBJECTS:
            if subj in df_cls.columns:
                row[f"{subj}过自招高线"] = (df_cls[subj] >= subject_high_lines[subj]).sum()
                row[f"{subj}过自招低线"] = (df_cls[subj] >= subject_low_lines[subj]).sum()
                row[f"{subj}过本科高线"] = (df_cls[subj] >= bench_subject_high[subj]).sum()
                row[f"{subj}过本科低线"] = (df_cls[subj] >= bench_subject_low[subj]).sum()
        # 贡献率（四套）
        for subj in ALL_SUBJECTS:
            if subj in df_cls.columns:
                for label, (sline, tline) in [
                    ("自招高线贡献率", (subject_high_lines[subj], high_score)),
                    ("自招低线贡献率", (subject_low_lines[subj], low_score)),
                    ("本科高线贡献率", (bench_subject_high[subj], bench_high)),
                    ("本科低线贡献率", (bench_subject_low[subj], bench_low)),
                ]:
                    total_online = (df_cls[subj] >= sline).sum()
                    contrib = ((df_cls[subj] >= sline) & (df_cls[total_col] >= tline)).sum()
                    row[f"{subj}{label}"] = contrib / total_online if total_online > 0 else 0
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


def create_zb_avg_sheet(df, total_col, high_score, low_score, bench_high, bench_low,
                        subject_high_lines, subject_low_lines,
                        bench_subject_high, bench_subject_low):
    if "走班" not in df.columns:
        return pd.DataFrame()
    rows = []
    for zb_name in df["走班"].unique():
        if pd.isna(zb_name) or zb_name == "":
            continue
        df_zb = df[df["走班"] == zb_name]
        row = {"走班班": zb_name, "人数": len(df_zb)}
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

            # 四套达线及贡献率
            for label, (sline, tline) in [
                ("自招高线", (subject_high_lines[subject], high_score)),
                ("自招低线", (subject_low_lines[subject], low_score)),
                ("本科高线", (bench_subject_high[subject], bench_high)),
                ("本科低线", (bench_subject_low[subject], bench_low)),
            ]:
                total_zb = (df_zb[subject] >= sline).sum()
                contrib_zb = ((df_zb[subject] >= sline) & (df_zb[total_col] >= tline)).sum()
                row[f"{label}达线人数"] = total_zb
                row[f"{label}有效贡献率"] = contrib_zb / total_zb if total_zb > 0 else 0
        rows.append(row)
    return pd.DataFrame(rows)


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
    uploaded_file = st.file_uploader("上传成绩表 (Excel)", type=["xls", "xlsx"], key="score_file")
    if uploaded_file is not None:
        df, total_col, has_zb = load_data(uploaded_file)
        if df is not None:
            st.success(f"✅ 读取成功！共 {len(df)} 名学生")
        else:
            st.stop()
    else:
        st.info("请先上传成绩表")
        st.stop()

    # ---------- 导入分数线 ----------
    st.markdown("---")
    st.subheader("📥 导入分数线")
    score_file = st.file_uploader("上传分数线文件 (可选)", type=["xls", "xlsx"], key="score_line_file")
    if score_file is not None:
        parsed = parse_score_file(score_file)
        if parsed is not None:
            # 更新总分
            for key in ["自招高线", "自招低线", "本科高线", "本科低线"]:
                if key in parsed:
                    val = parsed[key]
                    if key == "自招高线":
                        st.session_state.high_score = val
                    elif key == "自招低线":
                        st.session_state.low_score = val
                    elif key == "本科高线":
                        st.session_state.bench_high = val
                    elif key == "本科低线":
                        st.session_state.bench_low = val
            # 更新单科
            if 'subject_lines' in parsed:
                for key, val in parsed['subject_lines'].items():
                    # key 形如 "语文_high", "语文_low", "语文_bench"
                    parts = key.split('_')
                    if len(parts) == 2:
                        subj, typ = parts
                        if subj in ALL_SUBJECTS:
                            if typ == "high":
                                st.session_state.subject_high_lines[subj] = val
                            elif typ == "low":
                                st.session_state.subject_low_lines[subj] = val
                            elif typ == "bench":
                                st.session_state.bench_subject_high_lines[subj] = val  # 默认用高线
                                # 如果需要本科低线，需要额外字段
            st.success("✅ 分数线已导入！")
            st.rerun()

    # ---------- 分析模式 ----------
    st.markdown("---")
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
                st.warning("无走班数据")
                st.stop()
        else:
            st.warning("未找到走班列")
            st.stop()

    if len(df_selected) == 0:
        st.warning("该班级无数据")
        st.stop()

    # ---------- 分数线输入 ----------
    st.markdown("---")
    st.header("⚙️ 分数线设置")

    # 总分线
    high_score = st.number_input("自招高线", value=st.session_state.high_score, step=0.001, key="input_high")
    low_score = st.number_input("自招低线", value=st.session_state.low_score, step=0.001, key="input_low")
    bench_high = st.number_input("本科高线", value=st.session_state.bench_high, step=0.001, key="input_bench_high")
    bench_low = st.number_input("本科低线", value=st.session_state.bench_low, step=0.001, key="input_bench_low")
    # 同步
    st.session_state.high_score = high_score
    st.session_state.low_score = low_score
    st.session_state.bench_high = bench_high
    st.session_state.bench_low = bench_low

    # 自招高线单科
    st.subheader("自招高线单科线")
    subject_high_lines = {}
    cols = st.columns(3)
    for i, subj in enumerate(ALL_SUBJECTS):
        col = cols[i % 3]
        val = col.number_input(f"{subj}", value=st.session_state.subject_high_lines[subj], step=0.5, key=f"sub_high_{subj}")
        subject_high_lines[subj] = val
        st.session_state.subject_high_lines[subj] = val

    # 自招低线单科
    st.subheader("自招低线单科线")
    subject_low_lines = {}
    cols = st.columns(3)
    for i, subj in enumerate(ALL_SUBJECTS):
        col = cols[i % 3]
        val = col.number_input(f"{subj}", value=st.session_state.subject_low_lines[subj], step=0.5, key=f"sub_low_{subj}")
        subject_low_lines[subj] = val
        st.session_state.subject_low_lines[subj] = val

    # 本科高线单科
    st.subheader("本科高线单科线")
    bench_subject_high = {}
    cols = st.columns(3)
    for i, subj in enumerate(ALL_SUBJECTS):
        col = cols[i % 3]
        val = col.number_input(f"{subj}", value=st.session_state.bench_subject_high_lines[subj], step=0.5, key=f"bench_high_{subj}")
        bench_subject_high[subj] = val
        st.session_state.bench_subject_high_lines[subj] = val

    # 本科低线单科
    st.subheader("本科低线单科线")
    bench_subject_low = {}
    cols = st.columns(3)
    for i, subj in enumerate(ALL_SUBJECTS):
        col = cols[i % 3]
        val = col.number_input(f"{subj}", value=st.session_state.bench_subject_low_lines[subj], step=0.5, key=f"bench_low_{subj}")
        bench_subject_low[subj] = val
        st.session_state.bench_subject_low_lines[subj] = val

    # ---------- 导出 ----------
    st.markdown("---")
    if st.button("📥 导出完整报表（41个Sheet）"):
        with st.spinner("正在生成报表..."):
            excel_data = generate_excel_report(
                df, total_col, high_score, low_score, bench_high, bench_low,
                subject_high_lines, subject_low_lines,
                bench_subject_high, bench_subject_low
            )
            st.download_button(
                label="📥 下载报表",
                data=excel_data,
                file_name="成绩分析报表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


# ============================================================
# 主区域预览（显示四套线）
# ============================================================
if uploaded_file is not None and df is not None:
    metrics = calc_preview_metrics(
        df_selected, high_score, low_score, bench_high, bench_low,
        subject_high_lines, subject_low_lines,
        bench_subject_high, bench_subject_low, total_col
    )

    st.subheader(f"⚡ {display_name} 成绩分析报告")

    col1, col2, col3, col4, col5, col6, col7, col8 = st.columns(8)
    col1.metric("平均分", f"{metrics['avg']:.2f}")
    col2.metric("自招高线达线率", f"{metrics['high_rate_total']*100:.1f}%")
    col3.metric("自招高线人数", metrics['high_total'])
    col4.metric("自招低线达线率", f"{metrics['low_rate_total']*100:.1f}%")
    col5.metric("自招低线人数", metrics['low_total'])
    col6.metric("本科高线达线率", f"{metrics['bench_high_rate_total']*100:.1f}%")
    col7.metric("本科高线人数", metrics['bench_high_total'])
    col8.metric("本科低线人数", metrics['bench_low_total'])

    st.markdown("---")

    # 自招高线有效贡献
    st.subheader(f"📊 自招高线（≥{high_score:.0f}分）有效贡献")
    high_df = pd.DataFrame({
        "科目": ALL_SUBJECTS,
        "贡献人数": [metrics["high_contrib"][s] for s in ALL_SUBJECTS],
        "有效贡献率": [f"{metrics['high_rate'][s]*100:.1f}%" for s in ALL_SUBJECTS],
    })
    st.table(high_df.style.hide(axis="index"))

    # 自招低线有效贡献
    st.subheader(f"📊 自招低线（≥{low_score:.0f}分）有效贡献")
    low_df = pd.DataFrame({
        "科目": ALL_SUBJECTS,
        "贡献人数": [metrics["low_contrib"][s] for s in ALL_SUBJECTS],
        "有效贡献率": [f"{metrics['low_rate'][s]*100:.1f}%" for s in ALL_SUBJECTS],
    })
    st.table(low_df.style.hide(axis="index"))

    # 本科高线有效贡献
    st.subheader(f"📊 本科高线（≥{bench_high:.0f}分）有效贡献")
    bench_high_df = pd.DataFrame({
        "科目": ALL_SUBJECTS,
        "贡献人数": [metrics["bench_high_contrib"][s] for s in ALL_SUBJECTS],
        "有效贡献率": [f"{metrics['bench_high_rate'][s]*100:.1f}%" for s in ALL_SUBJECTS],
    })
    st.table(bench_high_df.style.hide(axis="index"))

    # 本科低线有效贡献
    st.subheader(f"📊 本科低线（≥{bench_low:.0f}分）有效贡献")
    bench_low_df = pd.DataFrame({
        "科目": ALL_SUBJECTS,
        "贡献人数": [metrics["bench_low_contrib"][s] for s in ALL_SUBJECTS],
        "有效贡献率": [f"{metrics['bench_low_rate'][s]*100:.1f}%" for s in ALL_SUBJECTS],
    })
    st.table(bench_low_df.style.hide(axis="index"))

    with st.expander("📋 查看本班学生明细"):
        display_cols = ["姓名"] + ALL_SUBJECTS + [total_col]
        existing_cols = [c for c in display_cols if c in df_selected.columns]
        st.dataframe(df_selected[existing_cols], use_container_width=True)

    st.caption("💡 修改左侧分数线后，所有数据自动更新。点击'导出完整报表'可下载41个Sheet的Excel文件。")
