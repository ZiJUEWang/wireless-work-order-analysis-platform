import pandas as pd
import streamlit as st
import os

# 页面基础配置（宽屏+标题）
st.set_page_config(page_title="无线工单分析平台", layout="wide")
st.title("📊 无线工单分析平台")

# ---------------------- 区县提取函数 ----------------------
def extract_maintenance_station(dept_str):
    """提取现场维护综合化后的XXX维护站"""
    if pd.isna(dept_str):
        return ""
    dept_str = str(dept_str)
    # 分割路径并查找"现场维护综合化"
    parts = dept_str.split("/")
    try:
        idx = parts.index("现场维护综合化")
        if idx + 1 < len(parts):
            station = parts[idx + 1]
            # 只保留"XXX维护站"（过滤后续层级，如"/威县综合维护划小区域二"）
            if "维护站" in station:
                return station
            # 兼容后续层级包含维护站的情况（如威县综合维护站划小区域二）
            for part in parts[idx+1:]:
                if "维护站" in part:
                    return part
    except ValueError:
        pass
    return ""

def get_district(row, simplify=False):
    """
    获取最终区县名称
    simplify: True=仅保留区县名（如威县），False=保留完整维护站名（如威县综合维护站）
    """
    accept_dept = str(row.get("受理部门", ""))
    audit_dept = str(row.get("审核部门", ""))
    
    # 规则1：受理部门是无线失败单处理组 → 从审核部门提取
    if "无线失败单处理组" in accept_dept:
        station = extract_maintenance_station(audit_dept)
    # 规则2：其他情况 → 从受理部门提取
    else:
        station = extract_maintenance_station(accept_dept)
    
    # 精简为纯区县名（如"威县综合维护站"→"威县"）
    if simplify and station:
        return station.replace("综合维护站", "").replace("维护站", "")
    return station

# ---------------------- 1. 文件上传与数据读取（核心前置） ----------------------
uploaded_file = st.file_uploader("📤 上传Excel工单文件", type=["xlsx", "xls"])
if uploaded_file is None:
    st.info("请先上传Excel格式的工单文件，支持.xlsx/.xls格式")
    st.stop()

# 读取数据（异常处理+友好提示）
try:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    # 自动添加提取后的区县列（两种格式，按需使用）
    df["维护站"] = df.apply(get_district, axis=1, simplify=False)  # 维护站名
    df["区县"] = df.apply(get_district, axis=1, simplify=True)     # 区县名
    st.success(f"✅ 文件读取成功！共加载 {len(df)} 条工单数据，已自动提取区县信息")
    
    # 原始数据预览（新增区县列）
    with st.expander("📄 查看原始数据+区县提取结果", expanded=False):
        st.dataframe(df[["受理部门", "审核部门", "维护站", "区县"] + list(df.columns[:5])].head(15), 
                     use_container_width=True)
except Exception as e:
    st.error(f"❌ 文件读取失败：{str(e)}")
    st.error("请确认已安装依赖：pip install pandas openpyxl")
    st.stop()

# ---------------------- 通用工具函数（避免重复代码） ----------------------
def check_required_cols(required_cols):
    """检查必要列是否存在，返回缺失列列表"""
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        st.warning(f"⚠️ 缺少必要列：{', '.join(missing_cols)}，相关指标暂无法统计")
    return len(missing_cols) == 0

def format_value_counts(series, col1, col2):
    """将value_counts的Series转为指定列名的DataFrame，避免columns参数报错"""
    df_formatted = series.reset_index()
    df_formatted.columns = [col1, col2]
    return df_formatted

# ---------------------- 2. 核心基础功能：所有统计指标（统一模块） ----------------------
st.divider()
st.subheader("📋 核心工单指标统计（按提取的区县维度）")

# 选择显示的区县格式（完整/精简）
district_col = st.radio("📌 选择区县显示格式", 
                        options=["维护站", "区县"], 
                        index=0, horizontal=True)

# 2.1 第一行：未恢复回单 + 超时工单
col1, col2 = st.columns(2)

# 2.1.1 未恢复回单统计（按提取的区县）
with col1:
    st.markdown("### 🔴 未恢复回单")
    not_recovered = pd.DataFrame()
    if check_required_cols(["是否恢复", "恢复时间", district_col]):
        not_recovered = df[df["是否恢复"] == "否"]
        st.metric("未恢复工单总数", len(not_recovered))
        
        # 按提取后的区县统计（替代原受理部门）
        if not not_recovered.empty:
            st.markdown(f"#### 按{district_col}分布")
            county_stats = not_recovered[district_col].value_counts()
            county_df = format_value_counts(county_stats, district_col, "工单数量")
            st.dataframe(county_df, use_container_width=True)
        
        # 明细查看（显示区县列）
        with st.expander("📝 查看未恢复回单明细", expanded=False):
            st.dataframe(not_recovered[[district_col,"故障编码", "是否恢复", "恢复时间", "故障标题"]], 
                         use_container_width=True)
    else:
        st.empty()

# 2.1.2 超时工单统计（按提取的区县）
with col2:
    st.markdown("### 🟡 超时工单")
    timeout = pd.DataFrame()
    # 兼容两种列名格式（新增时限标识）
    if "时限标识" in df.columns:
        timeout = df[df["时限标识"] == "超时"]
        st.metric("超时工单总数", len(timeout))
    elif "处理状态" in df.columns:
        timeout = df[df["处理状态"].str.contains("超时", na=False)]
        st.metric("超时工单总数", len(timeout))
    else:
        st.warning("⚠️ 缺少「时限标识」或「处理状态」列，无法统计超时工单")
    
    # 按提取后的区县统计
    if district_col in df.columns and not timeout.empty:
        st.markdown(f"#### 按{district_col}分布")
        timeout_county = timeout[district_col].value_counts()
        timeout_df = format_value_counts(timeout_county, district_col, "工单数量")
        st.dataframe(timeout_df, use_container_width=True)
    
    # 明细查看
    with st.expander("📝 查看超时工单明细", expanded=False):
        st.dataframe(timeout[[district_col,"故障编码", "故障标题", "时限标识"]], use_container_width=True)

# 2.2 第二行：超频指标统计（省公司+区县考核）
st.divider()
col3, col4 = st.columns(2)

# 2.2.1 省公司超频（故障标题≥5次）
with col3:
    st.markdown("### 🔵 省公司超频指标（故障标题≥5次）")
    if check_required_cols(["故障标题"]):
        title_count = df["故障标题"].value_counts()
        province_overfreq = title_count[title_count >= 5]
        st.metric("超频故障标题数", len(province_overfreq))
        province_df = format_value_counts(province_overfreq, "故障标题", "出现次数")
        st.dataframe(province_df, use_container_width=True)

# 2.2.2 区县考核超频（故障标题≥8次）
with col4:
    st.markdown("### 🟠 区县考核超频指标（故障标题≥8次）")
    if check_required_cols(["故障标题"]):
        title_count = df["故障标题"].value_counts()
        county_overfreq = title_count[title_count >= 8]
        st.metric("超频故障标题数", len(county_overfreq))
        county_df = format_value_counts(county_overfreq, "故障标题", "出现次数")
        st.dataframe(county_df, use_container_width=True)

# 2.3 第三行：回单定级错误 + 故障历时达标统计
st.divider()
col5, col6 = st.columns(2)

# 2.3.1 AB/CD类故障历时达标统计（剔除驻波、室分无备电）
with col5:
    st.markdown("### 🟣 AB/CD类故障历时达标统计（剔除驻波、室分无备电）")
    ab_not_meet = pd.DataFrame()
    cd_not_meet = pd.DataFrame()
    df_filtered = df
    if check_required_cols(["故障等级", "故障历时", "故障标题", "故障原因", district_col]):
        # 过滤驻波、室分无备电
        df_filtered = df[~df["故障标题"].str.contains("驻波", na=False)]
        df_filtered = df_filtered[~df_filtered["故障原因"].str.contains("室分站无备电", na=False)]
        st.info(f"剔除驻波、室分无备电类工单后，剩余工单：{len(df_filtered)} 条")
        
        # 筛选不达标工单
        ab_not_meet = df_filtered[
            (df_filtered["故障等级"].isin(["A", "B"])) & 
            (df_filtered["故障历时"] >= 240)
        ]
        cd_not_meet = df_filtered[
            (df_filtered["故障等级"].isin(["C", "D"])) & 
            (df_filtered["故障历时"] >= 360)
        ]
        
        # 展示核心指标
        col6_1, col6_2 = st.columns(2)
        with col6_1:
            st.metric("AB类不达标数（≥4小时）", len(ab_not_meet))
        with col6_2:
            st.metric("CD类不达标数（≥6小时）", len(cd_not_meet))
        
        # 按提取后的区县统计
        if not ab_not_meet.empty:
            st.markdown(f"#### AB类不达标 - 按{district_col}分布")
            ab_county = ab_not_meet[district_col].value_counts()
            ab_county_df = format_value_counts(ab_county, district_col, "不达标数量")
            st.dataframe(ab_county_df, use_container_width=True)
        
        if not cd_not_meet.empty:
            st.markdown(f"#### CD类不达标 - 按{district_col}分布")
            cd_county = cd_not_meet[district_col].value_counts()
            cd_county_df = format_value_counts(cd_county, district_col, "不达标数量")
            st.dataframe(cd_county_df, use_container_width=True)
        
        # 明细查看
        with st.expander("📝 查看不达标工单明细", expanded=False):
            st.markdown("##### AB类不达标明细")
            st.dataframe(ab_not_meet[[district_col, "故障等级", "故障历时", "故障编码","故障标题"]], use_container_width=True)
            st.markdown("##### CD类不达标明细")
            st.dataframe(cd_not_meet[[district_col, "故障等级", "故障历时", "故障编码","故障标题"]], use_container_width=True)
    else:
        st.warning("⚠️ 缺少「故障等级/故障历时/故障标题/故障原因/故障编码」列，无法统计历时达标情况")

# 2.3.2 回单定级错误统计（按提取的区县）
with col6:
    st.markdown("### 🟢 回单定级错误统计")
    wrong_level = pd.DataFrame()
    if check_required_cols(["回单定级", district_col]):
        wrong_level = df[df["回单定级"] != "较轻故障三级"]
        st.metric("非「较轻故障三级」工单总数", len(wrong_level))
        
        # 按提取后的区县分布
        st.markdown(f"#### 错误定级类型分布（按{district_col}）")
        level_stats = wrong_level[district_col].value_counts()
        level_df = format_value_counts(level_stats, district_col, "数量")
        st.dataframe(level_df, use_container_width=True)
        
        # 明细查看
        with st.expander("📝 查看定级错误明细", expanded=False):
            st.dataframe(wrong_level[[district_col, "回单定级", "故障编码","故障标题"]], use_container_width=True)


# ---------------------- 3. 数据导出功能（统一出口） ----------------------
st.divider()
st.subheader("💾 数据导出")
if st.button("📥 导出所有统计结果到Excel", type="primary"):
    export_path = "工单分析统计结果.xlsx"
    # 整合所有统计数据
    with pd.ExcelWriter(export_path, engine="openpyxl") as writer:
        # 原始数据+提取的区县列
        df.to_excel(writer, sheet_name="原始数据+区县", index=False)
        # 未恢复回单（含区县）
        if not not_recovered.empty:
            not_recovered[[district_col] + list(not_recovered.columns[:10])].to_excel(writer, sheet_name="未恢复回单", index=False)
        # 超时工单（含区县）
        if not timeout.empty:
            timeout[[district_col] + list(timeout.columns[:10])].to_excel(writer, sheet_name="超时工单", index=False)
        # 超频指标
        if check_required_cols(["故障标题"]):
            title_count = df["故障标题"].value_counts()
            title_df = format_value_counts(title_count, "故障标题", "出现次数")
            title_df.to_excel(writer, sheet_name="故障标题频次", index=False)
        # 回单定级错误（含区县）
        if not wrong_level.empty:
            wrong_level[[district_col, "回单定级", "故障标题"]].to_excel(writer, sheet_name="回单定级错误", index=False)
        # 故障历时不达标（含区县）
        if not ab_not_meet.empty:
            ab_not_meet[[district_col, "故障等级", "故障历时", "故障标题"]].to_excel(writer, sheet_name="AB类历时不达标", index=False)
        if not cd_not_meet.empty:
            cd_not_meet[[district_col, "故障等级", "故障历时", "故障标题"]].to_excel(writer, sheet_name="CD类历时不达标", index=False)
    
    # 下载按钮
    with open(export_path, "rb") as f:
        st.download_button(
            label="点击下载Excel文件",
            data=f,
            file_name="无线工单分析统计结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    st.success("✅ 统计结果已生成，可点击按钮下载！")
    
    # 清理临时文件
    if os.path.exists(export_path):
        os.remove(export_path)