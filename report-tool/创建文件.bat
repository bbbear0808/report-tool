@echo off
mkdir .github\workflows 2>nul

echo 文件1/5: app.py
(
echo import streamlit as st
echo import pandas as pd
echo import numpy as np
echo from docxtpl import DocxTemplate
echo import io
echo import os
echo import sys
echo.
echo def get_base_dir^(^):
echo     if getattr^(sys, 'frozen', False^):
echo         return os.path.dirname^(sys.executable^)
echo     return os.path.dirname^(os.path.abspath^(__file__^)^)
echo.
echo def find_column^(df, keywords, allow_missing=False^):
echo     for col in df.columns:
echo         col_lower = str^(col^).lower^(^).replace^(' ', ''^)
echo         for kw in keywords:
echo             if kw in col_lower:
echo                 return col
echo     if allow_missing:
echo         return None
echo     raise ValueError^(f"未找到列，可用列为：{list^(df.columns^)}"^)
echo.
echo def process_progress_sheet^(df^):
echo     progress_cols = {
echo         'priority': find_column^(df, ['优先级', 'priority']^),
echo         'test_func': find_column^(df, ['测试功能', '功能', '模块']^),
echo         'start_date': find_column^(df, ['开测时间', '开始时间']^),
echo         'end_date': find_column^(df, ['完结时间', '结束时间']^),
echo         'later_fix': find_column^(df, ['后期修复']^),
echo         'later_fix_level': find_column^(df, ['后期修复bug等级', '后期修复等级']^),
echo         'no_fix': find_column^(df, ['不做修改', '不修改']^),
echo         'no_fix_level': find_column^(df, ['不做修改bug等级', '不修改等级']^),
echo     }
echo     result_df = df[[v for v in progress_cols.values(^)]].copy(^)
echo     result_df.columns = list^(progress_cols.keys(^)^)
echo     return result_df.to_dict^(orient='records'^)
echo.
echo def process_bug_sheet^(df, sheet_name^):
echo     bug_level_col = find_column^(df, ['bug等级', '等级', 'level']^)
echo     bug_desc_col = find_column^(df, ['bug描述', '描述', 'description']^)
echo     level_series = df[bug_level_col].astype^(str^).str.strip(^).str.upper(^)
echo     valid_mask = level_series ^^= 'NA'
echo     df_clean = df[valid_mask].copy(^)
echo     total = len^(df_clean^)
echo     s_count = int^((level_series[valid_mask] ^^= 'S'^).sum(^)^)
echo     a_count = int^((level_series[valid_mask] ^^= 'A'^).sum(^)^)
echo     b_count = int^((level_series[valid_mask] ^^= 'B'^).sum(^)^)
echo     c_count = int^((level_series[valid_mask] ^^= 'C'^).sum(^)^)
echo     bugs = df_clean[[bug_level_col, bug_desc_col]].copy(^)
echo     bugs.columns = ['Bug等级', 'Bug描述']
echo     bugs_list = bugs.to_dict^(orient='records'^)
echo     return {
echo         'sheet_name': sheet_name,
echo         'total_bugs': total,
echo         's_count': s_count,
echo         'a_count': a_count,
echo         'b_count': b_count,
echo         'c_count': c_count,
echo         'bugs': bugs_list
echo     }
echo.
echo def generate_report^(context, template_path^):
echo     doc = DocxTemplate^(template_path^)
echo     doc.render^(context^)
echo     output = io.BytesIO(^)
echo     doc.save^(output^)
echo     output.seek^(0^)
echo     return output
echo.
echo st.set_page_config^(page_title="测试报告生成器", layout="wide"^)
echo st.title^("📄 本地测试报告自动生成工具"^)
echo.
echo base = get_base_dir(^)
echo default_template_path = os.path.join^(base, "template.docx"^)
echo.
echo if 'progress_data' not in st.session_state:
echo     st.session_state.progress_data = None
echo if 'bug_stats' not in st.session_state:
echo     st.session_state.bug_stats = []
echo.
echo with st.expander^("📂 步骤1：上传 Excel 文件", expanded=True^):
echo     uploaded_excel = st.file_uploader^("选择从谷歌表格下载的 .xlsx 文件", type=['xlsx']^)
echo     if uploaded_excel:
echo         try:
echo             xls = pd.ExcelFile^(uploaded_excel^)
echo             sheet_names = xls.sheet_names
echo             st.success^(f"已读取 {len^(sheet_names^)} 个页签：{', '.join^(sheet_names^)}"^)
echo             progress_sheet = st.selectbox^("请选择"每月更新进度情况"页签", options=sheet_names, index=0^)
echo             all_sheets = {name: pd.read_excel^(xls, sheet_name=name^) for name in sheet_names}
echo             if st.button^("提取数据", type="primary"^):
echo                 df_progress = all_sheets[progress_sheet]
echo                 st.session_state.progress_data = process_progress_sheet^(df_progress^)
echo                 bug_stats = []
echo                 for name in sheet_names:
echo                     if name ^^= progress_sheet:
echo                         continue
echo                     try:
echo                         stats = process_bug_sheet^(all_sheets[name], name^)
echo                         bug_stats.append^(stats^)
echo                     except Exception as e:
echo                         st.warning^(f"页签"{name}"处理失败，已跳过。原因：{e}"^)
echo                 st.session_state.bug_stats = bug_stats
echo                 st.success^("数据提取完成！"^)
echo         except Exception as e:
echo             st.error^(f"读取文件出错：{e}"^)
echo.
echo if st.session_state.progress_data is not None:
echo     with st.expander^("📊 进度数据预览", expanded=False^):
echo         st.dataframe^(pd.DataFrame^(st.session_state.progress_data^)^)
echo     with st.expander^("🐞 Bug统计预览", expanded=False^):
echo         for stat in st.session_state.bug_stats:
echo             st.markdown^(f"**{stat['sheet_name']}**  Bug总数：{stat['total_bugs']} (S:{stat['s_count']}, A:{stat['a_count']}, B:{stat['b_count']}, C:{stat['c_count']})"^)
echo.
echo with st.expander^("✍️ 步骤2：填写报告基本信息", expanded=True^):
echo     col1, col2 = st.columns^(2^)
echo     with col1:
echo         project = st.text_input^("项目名称"^)
echo         version = st.text_input^("版本号"^)
echo     with col2:
echo         tester = st.text_input^("测试人员"^)
echo         date = st.date_input^("报告日期"^)
echo     conclusion = st.text_area^("测试结论", height=100^)
echo     remark = st.text_area^("备注/遗留风险", height=80^)
echo.
echo with st.expander^("📃 步骤3：上传 Word 模板并生成报告", expanded=True^):
echo     template_file = st.file_uploader^("选择报告模板 (.docx^)", type=['docx']^)
echo     use_default = st.checkbox^("使用程序同目录下的 template.docx"^)
echo     if st.button^("🚀 生成报告", type="primary"^):
echo         context = {
echo             'project': project,
echo             'version': version,
echo             'tester': tester,
echo             'date': str^(date^),
echo             'conclusion': conclusion,
echo             'remark': remark,
echo             'progress_table': st.session_state.progress_data or [],
echo             'bug_stats': st.session_state.bug_stats or [],
echo         }
echo         if use_default:
echo             if not os.path.exists^(default_template_path^):
echo                 st.error^("默认模板文件 template.docx 不存在"^)
echo                 st.stop(^)
echo             template_path = default_template_path
echo         elif template_file is not None:
echo             temp_template = os.path.join^(base, "temp_template.docx"^)
echo             with open^(temp_template, "wb"^) as f:
echo                 f.write^(template_file.getbuffer(^)^)
echo             template_path = temp_template
echo         else:
echo             st.error^("请上传模板或勾选使用默认模板"^)
echo             st.stop(^)
echo         try:
echo             output = generate_report^(context, template_path^)
echo             st.download_button^("⬇️ 下载最终报告", data=output, file_name=f"测试报告_{project}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"^)
echo         except Exception as e:
echo             st.error^(f"生成失败：{e}"^)
) > app.py

echo 文件2/5: run_app.py
(
echo import os
echo import sys
echo import streamlit.web.bootstrap as bootstrap
echo.
echo def main(^):
echo     if getattr^(sys, 'frozen', False^):
echo         base_dir = sys._MEIPASS
echo     else:
echo         base_dir = os.path.dirname^(os.path.abspath^(__file__^)^)
echo     app_path = os.path.join^(base_dir, "app.py"^)
echo     sys.argv = ["streamlit", "run", app_path, "--global.developmentMode=false", "--server.headless=true", "--browser.serverAddress=localhost"]
echo     bootstrap.run(^)
echo.
echo if __name__ == "__main__":
echo     main(^)
) > run_app.py

echo 文件3/5: requirements.txt
(
echo streamlit==1.28.1
echo pandas==2.0.3
echo openpyxl==3.1.2
echo python-docx==0.8.11
echo docxtpl==0.16.4
) > requirements.txt

echo 文件4/5: build.spec
(
echo # -*- mode: python ; coding: utf-8 -*-
echo from PyInstaller.utils.hooks import collect_submodules
echo.
echo a = Analysis(
echo     ['run_app.py'],
echo     pathex=[],
echo     binaries=[],
echo     datas=[('app.py', '.')],
echo     hiddenimports=['streamlit'] + collect_submodules('streamlit'),
echo     hookspath=[],
echo     hooksconfig={},
echo     runtime_hooks=[],
echo     excludes=[],
echo     noarchive=False,
echo )
echo.
echo pyz = PYZ(a.pure)
echo.
echo exe = EXE(
echo     pyz,
echo     a.scripts,
echo     a.binaries,
echo     a.datas,
echo     [],
echo     name='TestReportTool',
echo     debug=False,
echo     bootloader_ignore_signals=False,
echo     strip=False,
echo     upx=True,
echo     runtime_tmpdir=None,
echo     console=True,
echo )
) > build.spec

echo 文件5/5: .github/workflows/build.yml
(
echo name: Build EXE
echo.
echo on:
echo   workflow_dispatch:
echo.
echo jobs:
echo   build:
echo     runs-on: windows-latest
echo.
echo     steps:
echo     - uses: actions/checkout@v4
echo.
echo     - name: Set up Python
echo       uses: actions/setup-python@v5
echo       with:
echo         python-version: '3.10'
echo.
echo     - name: Install dependencies
echo       run: ^|
echo         python -m pip install --upgrade pip
echo         pip install -r requirements.txt
echo         pip install pyinstaller
echo.
echo     - name: Build with PyInstaller
echo       run: ^|
echo         pyinstaller build.spec
echo.
echo     - name: Upload EXE artifact
echo       uses: actions/upload-artifact@v4
echo       with:
echo         name: TestReportTool
echo         path: dist/TestReportTool.exe
) > .github\workflows\build.yml

echo.
echo ================================
echo  全部5个文件创建完成！
echo  文件夹内容：
dir /s /b
echo ================================
pause