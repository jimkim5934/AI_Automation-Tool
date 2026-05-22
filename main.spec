# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'ont_automation',
        'ont_automation.config',
        'ont_automation.database',
        'ont_automation.gui_helpers',
        'ont_automation.gui_app',
        'ont_automation.excel_report',
        'ont_automation.olt_connection',
        'ont_automation.ont_lookup',
        'ont_automation.ont_registration',
        'ont_automation.ont_provisioning',
        'ont_automation.ont_cleanup',
        'ont_automation.throughput_test',
        'ont_automation.nokia_olt',
        'ont_automation.test_engine',
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
    a.datas,
    [],
    name='main',
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
