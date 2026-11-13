# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = ['openpyxl', 'ttkbootstrap', 'PIL.ImageTk',
                 'pcf_renderer',   # 轴测图渲染适配层（pcf_app 内为惰性导入）
                 'pcf', 'fpdf']    # pcf2iso 渲染引擎 + fpdf2
tmp_ret = collect_all('ttkbootstrap')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
datas += [('symbols', 'symbols')]   # Isogen .ASC 符号库（管架等符号定义）


a = Analysis(
    ['pcf_app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # PIL.Image/ImageShow 的惰性 import IPython（notebook 显示探测）会被
    # 静态分析跟进去，连带 matplotlib→QtAgg→PyQt5 全家桶（93MB 的来源）；
    # openpyxl 的 numpy、cryptography 链同样是非必需的可选依赖，全部排除。
    excludes=['pymupdf', 'fitz', 'IPython', 'matplotlib', 'mpl_toolkits',
              'PyQt5', 'PySide6', 'shiboken6', 'pygame',
              'jedi', 'parso', 'pygments', 'wcwidth', 'prompt_toolkit',
              'numpy', 'cryptography', 'cffi', 'pycparser',
              'requests', 'certifi', 'chardet', 'charset_normalizer',
              'setuptools', 'wheel'],
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
    name='PCF收集器',
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
    version='version_info.py',
)
