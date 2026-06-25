# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file - Đóng gói thành .exe cho Windows
Chạy: pyinstaller build.spec
"""

block_cipher = None

a = Analysis(
    ['gui_app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('excel_to_word.py', '.'),
        ('merge_feedbacks.py', '.'),
        ('data_utils.py', '.'),
        ('app_logging.py', '.'),
        ('benchmark.py', '.'),
        ('feedback_classifier.py', '.'),
    ],
    hiddenimports=[
        'xlrd',
        'openpyxl',
        'openpyxl.cell._writer',
        'docx',
        'docx.opc',
        'docx.opc.constants',
        'docx.opc.part',
        'docx.opc.package',
        'docx.opc.pkgreader',
        'docx.opc.pkgwriter',
        'docx.opc.rel',
        'docx.oxml',
        'docx.oxml.ns',
        'lxml',
        'lxml.etree',
        'lxml._elementpath',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'cv2', 'torch', 'tensorflow',
        'notebook', 'IPython', 'pytest',
        'matplotlib',
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='TongHopKhaoSat',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,           # Không hiện cửa sổ console
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,               # Thêm icon.ico nếu có
)
