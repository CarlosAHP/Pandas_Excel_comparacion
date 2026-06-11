# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec — genera ComparadorCotizaciones.exe (Windows)."""

from PyInstaller.utils.hooks import collect_all, collect_data_files, collect_submodules

block_cipher = None

datas = [("config/field_mapping.yaml", "config")]
binaries = []
hiddenimports = [
    "tkinter",
    "tkinter.ttk",
    "tkinter.filedialog",
    "tkinter.messagebox",
    "yaml",
    "dotenv",
    "certifi",
    "anyio",
    "httpx",
    "httpcore",
    "h11",
    "sniffio",
    "idna",
    "pydantic",
    "pydantic_core",
    "typing_extensions",
    "dateutil",
    "numpy",
    "pandas._libs",
    "pandas._libs.tslibs.timedeltas",
    "openpyxl.cell._writer",
]

for pkg in ("pandas", "openpyxl", "supabase", "postgrest", "storage3", "supabase_auth", "realtime"):
    try:
        tmp = collect_all(pkg)
        datas += tmp[0]
        binaries += tmp[1]
        hiddenimports += tmp[2]
    except Exception:
        hiddenimports += collect_submodules(pkg)

hiddenimports += collect_submodules("httpx")
datas += collect_data_files("certifi")

a = Analysis(
    ["src/main.py"],
    pathex=["src"],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="ComparadorCotizaciones",
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
