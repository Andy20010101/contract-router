# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

def safe_collect_submodules(package_name):
    try:
        return collect_submodules(package_name)
    except Exception:
        return []


def safe_collect_data_files(package_name):
    try:
        return collect_data_files(package_name)
    except Exception:
        return []


hiddenimports = []
datas = []
for package_name in ("rapidocr_onnxruntime", "rapidocr", "onnxruntime", "PIL"):
    hiddenimports += safe_collect_submodules(package_name)

for package_name in ("rapidocr_onnxruntime", "rapidocr"):
    datas += safe_collect_data_files(package_name)

a = Analysis(
    ["file_classifier.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
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
    name="contract-router",
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
    icon=None,
)
