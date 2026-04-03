# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec - macOS .app bundle"""

block_cipher = None

a = Analysis(
    ['gui_app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('excel_to_word.py', '.'),
    ],
    hiddenimports=[
        'xlrd', 'docx', 'docx.opc', 'docx.opc.constants',
        'docx.opc.part', 'docx.opc.package', 'docx.opc.pkgreader',
        'docx.opc.pkgwriter', 'docx.opc.rel', 'docx.oxml', 'docx.oxml.ns',
        'lxml', 'lxml.etree', 'lxml._elementpath',
    ],
    hookspath=[],
    excludes=['matplotlib', 'numpy', 'scipy', 'pandas', 'PIL', 'cv2',
              'torch', 'tensorflow', 'notebook', 'IPython', 'pytest'],
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='TongHopKhaoSat',
    debug=False,
    strip=False,
    upx=True,
    console=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    name='TongHopKhaoSat',
)

app = BUNDLE(
    coll,
    name='TongHopKhaoSat.app',
    icon=None,
    bundle_identifier='com.dbcl.tongkhopkhaosat',
    info_plist={
        'CFBundleName': 'Tổng hợp Khảo sát',
        'CFBundleDisplayName': 'Tổng hợp Khảo sát',
        'CFBundleShortVersionString': '1.0.0',
        'NSHighResolutionCapable': True,
    },
)
