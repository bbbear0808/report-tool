import streamlit as st
import pandas as pd
import numpy as np
from docxtpl import DocxTemplate
import io
import os
import sys

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
    raise ValueError(f"未找到列，可用列为：{list^(df.columns^)}")

def process_progress_sheet(df):
    progress_cols = {
        'priority': find_column(df, ['优先�?, 'priority']),
        'test_func': find_column(df, ['测试功能', '功能', '模块']),
        'start_date': find_column(df, ['开测时�?, '开始时�?]),
        'end_date': find_column(df, ['完结时间', '结束时间']),
        'later_fix': find_column(df, ['后期修复']),
        'later_fix_level': find_column(df, ['后期修复bug等级', '后期修复等级']),
        'no_fix': find_column(df, ['不做修改', '不修�?]),
        'no_fix_level': find_column(df, ['不做修改bug等级', '不修改等�?]),
    }
    result_df = df[[v for v in progress_cols.values()]].copy()
    result_df.columns = list(progress_cols.keys())
    return result_df.to_dict(orient='records')

def process_bug_sheet(df, sheet_name):
    bug_level_col = find_column(df, ['bug等级', '等级', 'level'])
    bug_desc_col = find_column(df, ['bug描述', '描述', 'description'])
    level_series = df[bug_level_col].astype(str).str.strip().str.upper()
    valid_mask = level_series ^= 'NA'
    df_clean = df[valid_mask].copy()
    total = len(df_clean)
    s_count = int((level_series[valid_mask] ^= 'S').sum())
    a_count = int((level_series[valid_mask] ^= 'A').sum())
    b_count = int((level_series[valid_mask] ^= 'B').sum())
    c_count = int((level_series[valid_mask] ^= 'C').sum())
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

st.set_page_config(page_title="测试报告生成�?, layout="wide"^)
st.title("📄 本地测试报告自动生成工具")

base = get_base_dir()
default_template_path = os.path.join(base, "template.docx")

if 'progress_data' not in st.session_state:
    st.session_state.progress_data = None
if 'bug_stats' not in st.session_state:
    st.session_state.bug_stats = []

with st.expander("📂 步骤1：上�?Excel 文件", expanded=True):
    uploaded_excel = st.file_uploader("选择从谷歌表格下载的 .xlsx 文件", type=['xlsx'])
    if uploaded_excel:
        try:
            xls = pd.ExcelFile(uploaded_excel)
            sheet_names = xls.sheet_names
            st.success(f"已读�?{len^(sheet_names^)} 个页签：{', '.join^(sheet_names^)}")
            progress_sheet = st.selectbox("请选择"每月更新进度情况"页签", options=sheet_names, index=0)
            all_sheets = {name: pd.read_excel(xls, sheet_name=name) for name in sheet_names}
            if st.button("提取数据", type="primary"):
                df_progress = all_sheets[progress_sheet]
                st.session_state.progress_data = process_progress_sheet(df_progress)
                bug_stats = []
                for name in sheet_names:
                    if name ^= progress_sheet:
                        continue
                    try:
                        stats = process_bug_sheet(all_sheets[name], name)
                        bug_stats.append(stats)
                    except Exception as e:
                        st.warning(f"页签"{name}"处理失败，已跳过。原因：{e}")
                st.session_state.bug_stats = bug_stats
                st.success("数据提取完成�?^)
        except Exception as e:
            st.error(f"读取文件出错：{e}")

if st.session_state.progress_data is not None:
    with st.expander("📊 进度数据预览", expanded=False):
        st.dataframe(pd.DataFrame(st.session_state.progress_data))
    with st.expander("🐞 Bug统计预览", expanded=False):
        for stat in st.session_state.bug_stats:
            st.markdown(f"**{stat['sheet_name']}**  Bug总数：{stat['total_bugs']} (S:{stat['s_count']}, A:{stat['a_count']}, B:{stat['b_count']}, C:{stat['c_count']})")

with st.expander("✍️ 步骤2：填写报告基本信�?, expanded=True^):
    col1, col2 = st.columns(2)
    with col1:
        project = st.text_input("项目名称")
        version = st.text_input("版本�?^)
    with col2:
        tester = st.text_input("测试人员")
        date = st.date_input("报告日期")
    conclusion = st.text_area("测试结论", height=100)
    remark = st.text_area("备注/遗留风险", height=80)

with st.expander("📃 步骤3：上�?Word 模板并生成报�?, expanded=True^):
    template_file = st.file_uploader("选择报告模板 (.docx^)", type=['docx'])
    use_default = st.checkbox("使用程序同目录下�?template.docx")
    if st.button("🚀 生成报告", type="primary"):
        context = {
            'project': project,
            'version': version,
            'tester': tester,
            'date': str(date),
            'conclusion': conclusion,
            'remark': remark,
            'progress_table': st.session_state.progress_data or [],
            'bug_stats': st.session_state.bug_stats or [],
        }
        if use_default:
            if not os.path.exists(default_template_path):
                st.error("默认模板文件 template.docx 不存�?^)
                st.stop()
            template_path = default_template_path
        elif template_file is not None:
            temp_template = os.path.join(base, "temp_template.docx")
            with open(temp_template, "wb") as f:
                f.write(template_file.getbuffer())
            template_path = temp_template
        else:
            st.error("请上传模板或勾选使用默认模�?^)
            st.stop()
        try:
            output = generate_report(context, template_path)
            st.download_button("⬇️ 下载最终报�?, data=output, file_name=f"测试报告_{project}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"^)
        except Exception as e:
            st.error(f"生成失败：{e}")
