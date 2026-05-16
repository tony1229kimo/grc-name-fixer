# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for GRC 團名修正工具
打包方式：pyinstaller grc-tool.spec --clean

產出：
  dist/GRC-團名修正工具/
    ├── GRC-團名修正工具.exe
    ├── _internal/...                (Python runtime + libraries)
    ├── templates/, static/          (網頁資源)
    └── dictionary.seed.db           (預灌字典種子)

第一次執行時，launcher 會檢查 exe 目錄是否有 dictionary.db，
若無則從 dictionary.seed.db 複製一份過去（之後持續累積使用）。

這樣設計的好處：升級時新 zip 解壓覆蓋會更新 dictionary.seed.db，
但同事的 dictionary.db（含累積學習）不會被沖掉。
"""

block_cipher = None

a = Analysis(
    ['launcher.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('templates', 'templates'),
        ('static', 'static'),
        ('dictionary.seed.db', '.'),   # 預灌字典(種子);launcher 第一次跑會複製成 dictionary.db
    ],
    hiddenimports=[
        'flask',
        'openpyxl',
        'database',
        'parser',
        'matcher',
        'app',
        '_version',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter', 'unittest', 'pydoc', 'doctest',
        # 不能排除: email, html, http, xmlrpc（Flask/werkzeug 會用）
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='GRC-團名修正工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,                   # 保留終端視窗（顯示狀態 + 關閉用）
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='GRC-團名修正工具',
)
