# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 打包配置 (OneDir 目录模式)
# 生成命令: pyinstaller pyrpa.spec  (Windows 下可用 build.bat 一键处理)

import os

project_root = os.path.abspath(SPECPATH)
icon_path = os.path.join(project_root, 'rpa_data', 'pyrpa.ico')

# 只读资源: 主题 tcl + 主题图片目录, 与 sun-valley.tcl 保持同目录结构
datas = [(os.path.join(project_root, 'sun-valley.tcl'), '.'),
         (os.path.join(project_root, 'theme'), 'theme')]

a = Analysis(
    ['main.py'],
    pathex=[project_root],
    binaries=[],
    datas=datas,
    hiddenimports=[],
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
    [],
    exclude_binaries=True,
    name='pyrpa',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_path if os.path.exists(icon_path) else None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='pyrpa',
)
