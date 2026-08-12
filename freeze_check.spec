# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['freeze_check.py'],
    pathex=['src'],
    binaries=[],
    datas=[],
    hiddenimports=['reportlab', 'reportlab.pdfbase.cidfonts', 'reportlab.pdfbase.ttfonts'],
    hookspath=['pyinstaller_hooks'],
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
    a.datas,
    [],
    name='freeze_check',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
