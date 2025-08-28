# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['serverRevision.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('.env.secure', '.'), 
        ('email_template.html', '.')
    ],
    hiddenimports=[
        'wmi',
        'requests',
        'requests.adapters',
        'requests.packages',
        'urllib3',
        'json',
        'dotenv'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'office365',  # Excluir la biblioteca antigua
        'matplotlib',
        'numpy',
        'pandas'
    ],
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
    name='serverRevision',
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