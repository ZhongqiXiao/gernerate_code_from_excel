# -*- mode: python ; coding: utf-8 -*-

import os

a = Analysis(
    ['src\\gui\\qrcode_gui.py'],
    pathex=[os.path.abspath('.')],  # 添加当前路径到搜索路径
    binaries=[],
    datas=[
        (os.path.join('src', 'core'), 'core')  # 将core目录添加到datas，确保打包后能找到core模块
    ],
    hiddenimports=[
        'pandas', 
        'openpyxl', 
        'qrcode', 
        'PIL', 
        'tkinter',
        'concurrent.futures',  # 添加concurrent.futures
        'pandas._libs.tslibs.parquet',  # pandas相关依赖
        'pandas._libs.tslibs.nattype',  # pandas相关依赖
        'core',  # 显式添加core模块作为hiddenimport
        'core.qrcode_processor',  # 显式添加core.qrcode_processor模块
        'core.config'  # 显式添加core.config模块
    ],
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
    name='QRCodeGenerator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,  # 禁用UPX压缩，可能会解决资源访问问题
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # 启用控制台以便调试
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # 添加以下参数解决Windows上的文件访问问题
    icon='',  # 设置空图标路径以避免图标文件访问问题
    # 增加这个参数可以避免某些Windows系统上的资源访问冲突
    manifest=None,
)
