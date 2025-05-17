# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['mummy.py'],
    pathex=[],
    binaries=[],
    datas=[('iso_excel.xlsx', '.'), ('executive_summary_template.docx', '.'), ('_internal', '_internal')],
    hiddenimports=['win32timezone', 'win32com', 'customtkinter'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='mummy',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='mummy',
)
