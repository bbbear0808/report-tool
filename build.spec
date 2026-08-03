# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

a = Analysis(
    ['run_app.py'],
    pathex=[],
    binaries=[],
    datas=[('app.py', '.')],
    hiddenimports=[
        'streamlit',
        'streamlit.web.bootstrap',
        'streamlit.runtime',
        'pandas',
        'openpyxl',
        'docx',
        'docxtpl',
        'lxml',
    ] + collect_submodules('streamlit'),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['streamlit.runtime.caching.cache_data_api'],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='TestReportTool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    runtime_tmpdir=None,
    console=True,  # 保留控制台窗口，方便查看状态
)
