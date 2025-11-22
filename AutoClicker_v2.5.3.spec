# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['AutoClicker_2.5.py'],
    pathex=[],
    binaries=[],
    datas=[('ICON', 'ICON')],
    hiddenimports=['pystray._win32', 'PIL._imaging', 'PIL._imagingtk', 'PIL._webp', 'win32timezone', 'win32api'],
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
    a.datas,
    [],
    name='AutoClicker_v2.5.3',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['ICON\\64.png'],
)
