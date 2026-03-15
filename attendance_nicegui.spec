# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

project_root = Path.cwd()
data_files = []
hiddenimports = []

usage_file = project_root / "data" / "月度文件" / "使用说明.txt"
if usage_file.exists():
    data_files.append((str(usage_file), "data/月度文件"))

usage_text = project_root / "使用说明-给同事看.txt"
if usage_text.exists():
    data_files.append((str(usage_text), "."))

try:
    data_files += collect_data_files("nicegui")
    hiddenimports += collect_submodules("nicegui")
except Exception:
    pass

try:
    hiddenimports += collect_submodules("starlette")
except Exception:
    pass

try:
    hiddenimports += collect_submodules("uvicorn")
except Exception:
    pass

a = Analysis(
    ["attendance_nicegui.py"],
    pathex=[str(project_root)],
    binaries=[],
    datas=data_files,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="财务公司考勤统计助手_NiceGUI",
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
