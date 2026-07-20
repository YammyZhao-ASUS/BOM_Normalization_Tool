# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path


project_root = Path(SPECPATH)

analysis = Analysis(
    [str(project_root / "src" / "desktop_app.py")],
    pathex=[str(project_root / "src")],
    binaries=[],
    datas=[
        (str(project_root / "config" / "default_rules.json"), "config"),
    ],
    hiddenimports=[
        "openpyxl",
        "openpyxl.chart",
        "pandas",
        "xlrd",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "IPython",
        "PIL",
        "matplotlib",
        "pandas.plotting",
        "pandas.tests",
        "pytest",
        "scipy",
    ],
    noarchive=False,
    optimize=1,
)

python_archive = PYZ(analysis.pure)

executable = EXE(
    python_archive,
    analysis.scripts,
    analysis.binaries,
    analysis.datas,
    [],
    name="BOM_Intelligence_Platform",
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
    version=str(project_root / "packaging" / "windows_version_info.txt"),
)