# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['launcher_clientes_onedrive.pyw'],
    pathex=[],
    binaries=[],
    datas=[('logo_buscador.png', '.'), ('logo_buscador.ico', '.')],
    hiddenimports=[],
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
    name='Buscador Cliente HeadCargo',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version='buscador_cliente_headcargo_version_info.txt',
    icon=['logo_buscador.ico'],
)
