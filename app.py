import streamlit as st
import pandas as pd
import numpy as np
from docxtpl import DocxTemplate
import io
import os
import sys
from datetime import datetime

# ---------- 基础函数 ----------
def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def find_column(df, keywords, allow_missing=False):
    for col in df.columns:
        col_lower = str(col).lower().replace(' ', '')
        for kw in keywords:
            if kw in col_lower:
                return col
    if allow_missing:
        return None
    raise ValueError(f"未找到包含 {keywords} 的列，可用列为：{list(df.columns)}")

def process_progress_sheet(df):
    progress_cols = {
        'priority': find_column(df, ['优先级', 'priority']),
        'test_func': find_column(df, ['测试功能', '功能', '模块', 'function']),
        'start_date': find_column(df, ['开测时间', '开始时间', 'start']),
        'end_date': find_column(df, ['完结时间', '结束时间', 'end']),
        'later_fix': find_column(df, ['后期修复', '延期修复']),
        'later_fix_level': find_column(df, ['后期修复bug等级', '后期修复等级']),
        'no_fix': find_column(df, ['不做修改', '不修改']),
        'no_fix_level': find_column(df, ['不做修改bug等级', '不修改等级']),
    }
    result_df = df[[v for v in progress_cols.values()]].copy()
    result_df.columns = list(progress_cols.keys())
    return result_df.to_dict(orient='records')

def process_bug_sheet(df, sheet_name):
    bug_level_col = find_column(df, ['bug等级', 'bug级别', '等级', 'level', '严重程度'])
    bug_desc_col = find_column(df, ['bug描述', '问题描述', '描述', 'description'])
    
    level_series = df[bug_level_col].astype(str).str.strip().str.upper()
    valid_mask = level_series != 'NA'
    df_clean = df[valid_mask].copy()
    
    total = len(df_clean)
    s_count = int((level_series[valid_mask] == 'S').sum())
    a_count = int((level_series[valid_mask] == 'A').sum())
    b_count = int((level_series[valid_mask] == 'B').sum())
    c_count = int((level_series[valid_mask] == 'C').sum())
    
    bugs = df_clean[[bug_level_col, bug_desc_col]].copy()
    bugs.columns = ['Bug等级', 'Bug描述']
    bugs_list = bugs.to_dict(orient='records')
    
    return {
        'sheet_name': sheet_name,
        'total_bugs': total,
        's_count': s_count,
        'a_count': a_count,
        'b_count': b_count,
        'c_count': c_count,
        'bugs': bugs_list
    }

def generate_report(context, template_path):
    doc = DocxTemplate(template_path)
    doc.render(context)
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# ---------- 页面配置 ----------
st.set_page_config(
    page_title="测试报告生成器",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 隐藏 Streamlit 默认样式，让界面更干净
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .stApp {margin-top: -50px;}
    div[data-testid="stExpander"] {border: 1px solid #ddd; border-radius: 10px; margin-bottom: 10px;}
</style>
""", unsafe_allow_html=True)

st.title("📄 测试报告自动生成工具")

base = get_base_dir()
default_template_path = os.path.join(base, "template.docx")

# 初始化 session
if 'progress_data' not in st.session_state:
    st.session_state.progress_data = None
if 'bug_stats' not in st.session_state:
    st.session_state.bug_stats = []
if 'sheet_names' not in st.session_state:
    st.session_state.sheet_names = []

# ==================== 步骤1：上传 Excel ====================
st.markdown("---")
st.subheader("📂 步骤1：上传 Excel 数据文件")

uploaded_excel = st.file_uploader(
    "选择从谷歌表格下载的 .xlsx 文件",
    type=['xlsx'],
    help="请先将谷歌表格下载为 Excel 格式（文件 → 下载 → Microsoft Excel）"
)

if uploaded_excel:
    try:
        xls = pd.ExcelFile(uploaded_excel)
        st.session_state.sheet_names = xls.sheet_names
        
        col_info, col_btn = st.columns([3, 1])
        with col_info:
            st.info(f"✅ 已读取 **{len(st.session_state.sheet_names)}** 个页签")
        
        progress_sheet = st.selectbox(
            "📋 请选择“每月更新进度情况”对应的页签",
            options=st.session_state.sheet_names,
            index=0 if st.session_state.sheet_names else 0
        )
        
        if st.button("🔍 提取数据", type="primary", use_container_width=True):
            with st.spinner("正在分析数据..."):
                all_sheets = {name: pd.read_excel(xls, sheet_name=name) for name in st.session_state.sheet_names}
                
                # 处理进度页签
                df_progress = all_sheets[progress_sheet]
                st.session_state.progress_data = process_progress_sheet(df_progress)
                
                # 处理其他页签
                bug_stats = []
                skipped = []
                for name in st.session_state.sheet_names:
                    if name == progress_sheet:
                        continue
                    try:
                        stats = process_bug_sheet(all_sheets[name], name)
                        bug_stats.append(stats)
                    except Exception as e:
                        skipped.append(name)
                
                st.session_state.bug_stats = bug_stats
                
                if bug_stats:
                    st.success(f"✅ 数据提取完成！成功处理 {len(bug_stats)} 个Bug页签")
                if skipped:
                    st.warning(f"⚠️ 跳过了 {len(skipped)} 个页签：{', '.join(skipped)}（缺少必要的Bug等级/描述列）")
                    
    except Exception as e:
        st.error(f"❌ 读取文件出错：{e}")

# 数据预览
if st.session_state.progress_data:
    with st.expander("📊 点击查看提取的数据预览"):
        tab1, tab2 = st.tabs(["进度数据", "Bug统计"])
        with tab1:
            df_preview = pd.DataFrame(st.session_state.progress_data)
            st.dataframe(df_preview, use_container_width=True)
        with tab2:
            for stat in st.session_state.bug_stats:
                cols = st.columns(5)
                cols[0].metric(f"📌 {stat['sheet_name']}", f"{stat['total_bugs']} 个Bug")
                cols[1].metric("S级", stat['s_count'])
                cols[2].metric("A级", stat['a_count'])
                cols[3].metric("B级", stat['b_count'])
                cols[4].metric("C级", stat['c_count'])
                st.divider()

# ==================== 步骤2：填写报告信息 ====================
st.markdown("---")
st.subheader("✍️ 步骤2：填写报告基本信息")

# 第一行：基本字段
col1, col2, col3, col4 = st.columns(4)
with col1:
    project = st.text_input("📁 项目名称", placeholder="例：XX项目V2.0")
with col2:
    version = st.text_input("🏷️ 版本号", placeholder="例：V2.0.1")
with col3:
    tester = st.text_input("👤 测试人员", placeholder="例：张三")
with col4:
    date = st.date_input("📅 报告日期", value=datetime.now().date())

# 第二行：测试环境和类型
col_env, col_type = st.columns(2)
with col_env:
    st.markdown("**🖥️ 测试环境（可多选）**")
    env_cols = st.columns(4)
    with env_cols[0]:
        env_ios = st.checkbox("iOS")
    with env_cols[1]:
        env_android = st.checkbox("Android")
    with env_cols[2]:
        env_web = st.checkbox("Web")
    with env_cols[3]:
        env_harmony = st.checkbox("鸿蒙")
    
    # 整理选中的环境
    selected_envs = []
    if env_ios: selected_envs.append("iOS")
    if env_android: selected_envs.append("Android")
    if env_web: selected_envs.append("Web")
    if env_harmony: selected_envs.append("鸿蒙")
    env_str = "、".join(selected_envs) if selected_envs else "未指定"

with col_type:
    test_type = st.radio(
        "**🧪 测试类型**",
        options=["功能测试", "集成测试", "回归测试", "验收测试", "性能测试", "其他"],
        horizontal=True
    )

# 第三行：测试结论
st.markdown("**📝 测试结论**")
conclusion = st.text_area(
    "测试结论",
    placeholder="请填写测试结论，例如：\n本次测试共发现Bug XX个，其中S级X个、A级X个...\n主要问题集中在登录模块和支付模块...\n建议修复S级和A级Bug后再上线...",
    height=120,
    label_visibility="collapsed"
)

# 第四行：备注
remark = st.text_area(
    "**⚠️ 备注/遗留风险**",
    placeholder="如有遗留风险或需要特别说明的事项，请在此填写...",
    height=80
)

# ==================== 步骤3：生成报告 ====================
st.markdown("---")
st.subheader("📃 步骤3：选择模板并生成报告")

col_template, col_generate = st.columns([2, 1])

with col_template:
    use_default = st.checkbox(
        "📄 使用程序同目录下的 template.docx",
        value=True,
        help="勾选后将自动使用 EXE 同目录下的模板文件"
    )
    if not use_default:
        template_file = st.file_uploader("手动上传报告模板", type=['docx'])
    else:
        template_file = None

with col_generate:
    st.markdown("<br>", unsafe_allow_html=True)
    generate_btn = st.button("🚀 生成并下载报告", type="primary", use_container_width=True)

if generate_btn:
    # 验证数据
    if st.session_state.progress_data is None:
        st.error("❌ 请先在步骤1中上传 Excel 文件并点击“提取数据”")
    elif not project:
        st.error("❌ 请填写项目名称")
    else:
        with st.spinner("正在生成报告..."):
            # 准备上下文
            context = {
                'project': project,
                'version': version,
                'tester': tester,
                'date': str(date),
                'test_type': test_type,
                'test_env': env_str,
                'conclusion': conclusion,
                'remark': remark,
                'progress_table': st.session_state.progress_data or [],
                'bug_stats': st.session_state.bug_stats or [],
            }
            
            # 确定模板路径
            if use_default:
                if not os.path.exists(default_template_path):
                    st.error("❌ 默认模板文件 template.docx 不存在，请与 EXE 放在同一目录，或手动上传模板。")
                    st.stop()
                template_path = default_template_path
            elif template_file is not None:
                temp_template = os.path.join(base, "temp_template.docx")
                with open(temp_template, "wb") as f:
                    f.write(template_file.getbuffer())
                template_path = temp_template
            else:
                st.error("❌ 请上传模板或勾选使用默认模板。")
                st.stop()
            
            try:
                output = generate_report(context, template_path)
                st.success("✅ 报告生成成功！点击下方按钮下载")
                st.download_button(
                    "⬇️ 下载测试报告",
                    data=output,
                    file_name=f"测试报告_{project}_{date}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"❌ 生成失败：{e}")

# 底部信息
st.markdown("---")
st.caption("💡 提示：所有数据均在本地处理，不会上传到任何服务器。")
