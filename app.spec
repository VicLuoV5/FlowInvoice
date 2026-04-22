# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for 极简票流 (FlowInvoice) desktop build.

Build:
    pyinstaller app.spec --clean

Output:
    dist/FlowInvoice.exe  (single-file, no console)
"""

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

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
