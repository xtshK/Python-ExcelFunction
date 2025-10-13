# -*- mode: python ; coding: utf-8 -*-
import os
from ttkbootstrap import __path__ as ttk_path

# 把 ttkbootstrap 的內建樣式、圖示資料夾打包進去
datas = [
    (os.path.join(ttk_path[0], "_tk_data"), "ttkbootstrap/_tk_data"),
    ('app/assets', 'app/assets'),  # 你原本的 assets
]

a = Analysis(
    ['run.py'],
    pathex=['.'],
    binaries=[],
    datas=datas,
    hiddenimports=[
        'openpyxl',
        'et_xmlfile',
        'ttkbootstrap',
        # 'xlrd',
    ],
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
    a.binaries,
    a.zipfiles,
    a.datas,
    name='CSC 4M Excel Formatter',
    icon='app/assets/4m_icon.ico',
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
