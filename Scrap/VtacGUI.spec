# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['vtac_gui_standalone.py'],
    pathex=[],
    binaries=[],
    datas=[('Scrap\\spiders\\handlers', 'Scrap\\spiders\\handlers')],
    hiddenimports=['twisted.internet.asyncioreactor', 'twisted.internet.selectreactor', 'service_identity', 'certifi', 'Scrap.spiders.vtac_spider', 'Scrap.spiders.handlers.vtac_es'],
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
    name='VtacGUI',
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
)
