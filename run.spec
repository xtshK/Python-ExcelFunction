# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['run.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('app/assets','app/assets'),   # 只留你的資源
    ],
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
    pyz, a.scripts, a.binaries, a.zipfiles, a.datas,
    name='CSC 4M Excel Formatter',
    icon='app/assets/4m_icon.ico',
    debug=False, bootloader_ignore_signals=False, strip=False,
    upx=True, console=False, disable_windowed_traceback=False,
    argv_emulation=False, target_arch=None, codesign_identity=None,
    entitlements_file=None,
)
