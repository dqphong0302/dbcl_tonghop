# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec dùng chung cho Windows và macOS.

  pyinstaller TongHopKhaoSat.spec --clean --noconfirm

- Windows: dist/TongHopKhaoSat.exe (1 file, chạy không cần Python)
- macOS:   dist/TongHopKhaoSat.app
"""
import os
import sys

from PyInstaller.utils.hooks import collect_data_files

APP_NAME = "TongHopKhaoSat"
APP_VERSION = os.environ.get("APP_VERSION", "3.0.0").lstrip("v")
IS_MAC = sys.platform == "darwin"

datas = [
    ("assets/icon_256.png", "assets"),
    ("assets/icon.ico", "assets"),
]
datas += collect_data_files("docx")           # templates/default.docx của python-docx
datas += collect_data_files("customtkinter")  # theme JSON + font của CustomTkinter

a = Analysis(
    ["gui_app.py"],
    datas=datas,
    hiddenimports=["openpyxl.cell._writer", "PIL._tkinter_finder"],
    excludes=["cv2", "torch", "tensorflow", "numpy", "pandas", "matplotlib",
              "notebook", "IPython", "pytest"],
    noarchive=False,
)
pyz = PYZ(a.pure)

if IS_MAC:
    exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name=APP_NAME,
              console=False, upx=False, argv_emulation=False)
    coll = COLLECT(exe, a.binaries, a.datas, name=APP_NAME, upx=False)
    app = BUNDLE(
        coll,
        name=f"{APP_NAME}.app",
        icon="assets/icon.icns",
        bundle_identifier="vn.edu.ump.dbcl.tonghopkhaosat",
        info_plist={
            "CFBundleName": "Tổng hợp Khảo sát",
            "CFBundleDisplayName": "Tổng hợp Khảo sát",
            "CFBundleShortVersionString": APP_VERSION,
            "CFBundleVersion": APP_VERSION,
            "NSHighResolutionCapable": True,
            "NSRequiresAquaSystemAppearance": True,  # luôn dùng giao diện sáng
        },
    )
else:
    exe = EXE(pyz, a.scripts, a.binaries, a.datas, [], name=APP_NAME,
              console=False, upx=False, icon="assets/icon.ico")
