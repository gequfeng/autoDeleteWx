# -*- mode: python ; coding: utf-8 -*-
# Win7 兼容版打包配置（Python 3.8 + PyInstaller 5.x）
from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = ['rapidocr', 'win32api', 'win32gui', 'win32con', 'win32process', 'pywintypes']
# Win7 兼容：强制收集 onnxruntime / rapidocr 的全部依赖 DLL，否则 onnxruntime_pybind11_state 会找不到依赖
tmp_ret = collect_all('onnxruntime')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('rapidocr')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
# Win7 还需要 Microsoft VC++ 2015-2022 运行库；环境缺的话打包自带一份
try:
    tmp_ret = collect_all('vcruntime')
    datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
except Exception:
    pass


a = Analysis(
    ['wxwork_cleanup.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
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
    name='wxwork_cleanup_win7',
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
