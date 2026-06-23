# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for 极简票流 (FlowInvoice) desktop build.

Build:
    pyinstaller app.spec --clean

Output:
    dist/FlowInvoice.exe  (single-file, no console)
"""

import os
import tkinter as tk
from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files, collect_submodules


def _configure_tcl_tk_from_env():
    """Allow builds from portable Python paths whose Tcl/Tk folder is relocated."""
    tcl_root = os.environ.get('FLOWINVOICE_TCL_ROOT')
    if not tcl_root:
        return

    root = Path(tcl_root)
    os.environ.setdefault('TCL_LIBRARY', str(root / 'tcl8.6'))
    os.environ.setdefault('TK_LIBRARY', str(root / 'tk8.6'))


def _assert_tkinter_runtime():
    """Fail the build if CustomTkinter cannot start at runtime."""
    try:
        root = tk.Tk()
        root.withdraw()
        root.destroy()
    except Exception as exc:
        raise SystemExit(
            'Tkinter/Tcl runtime is not usable; refusing to build a broken '
            'FlowInvoice.exe. If Python is installed under a non-ASCII path, '
            'copy its tcl folder to an ASCII path and set FLOWINVOICE_TCL_ROOT.'
        ) from exc


_configure_tcl_tk_from_env()
_assert_tkinter_runtime()
# RapidOCR 的 ONNX 模型 + yaml 配置必须随包，否则首次启动会尝试联网下载
rapidocr_datas = collect_data_files('rapidocr_onnxruntime', include_py_files=False)
rapidocr_hidden = collect_submodules('rapidocr_onnxruntime')

# CustomTkinter 的主题 json/字体资源
ctk_datas = collect_data_files('customtkinter')

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=rapidocr_datas + ctk_datas,
    hiddenimports=rapidocr_hidden + [
        'onnxruntime',
        'pywinstyles',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'notebook', 'IPython', 'pytest'],
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
    name='FlowInvoice',
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
    icon=None,  # 有 .ico 图标后在这里填路径
)
