# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 目录版打包配置（onedir，启动快、资源路径稳定）。"""

a = Analysis(
    ["app.py"],
    pathex=["src"],
    binaries=[],
    datas=[
        ("assets", "assets"),
        ("templates", "templates"),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["matplotlib", "numpy", "pandas", "scipy"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="TXT重建Excel",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon="assets/ui/icon.ico",
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="TXT重建Excel",
)
