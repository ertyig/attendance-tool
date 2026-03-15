# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
from PyInstaller.utils.hooks import collect_all

project_root = Path.cwd()
data_files = []
ctk_datas = []
ctk_binaries = []
ctk_hiddenimports = []

usage_file = project_root / "data" / "月度文件" / "使用说明.txt"
if usage_file.exists():
    data_files.append((str(usage_file), "data/月度文件"))

detailed_usage_file = project_root / "使用说明-给同事看.txt"
if detailed_usage_file.exists():
    data_files.append((str(detailed_usage_file), "."))

try:
    ctk_datas, ctk_binaries, ctk_hiddenimports = collect_all("customtkinter")
except Exception:
    ctk_datas, ctk_binaries, ctk_hiddenimports = [], [], []

a = Analysis(
    ["attendance_customtkinter.py"],
    pathex=[str(project_root)],
    binaries=ctk_binaries,
    datas=data_files + ctk_datas,
    hiddenimports=ctk_hiddenimports,
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
    name="考勤统计助手_CustomTkinter",
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
    version="attendance_gui_version_info.txt",
    codesign_identity=None,
    entitlements_file=None,
)
