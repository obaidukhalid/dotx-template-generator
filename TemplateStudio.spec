# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for the Windows build of Template Studio.

Build it with build_windows_exe.ps1, or by hand on a Windows machine:

    py -m pip install pyinstaller flask python-docx
    py -m PyInstaller --noconfirm TemplateStudio.spec

The result is dist\\TemplateStudio.exe, a single file with Python, Flask,
python-docx, the HTML pages and the shipped defaults inside it.
"""

block_cipher = None

# Read-only payload. These land in the temporary folder the one-file exe
# unpacks itself into, which app.py reads as BUNDLE_DIR. The writable copies of
# config.json and presets.json are seeded next to the exe on first run.
datas = [
    ("templates/index.html", "templates"),
    ("templates/feedback.html", "templates"),
    ("config.json", "."),
    ("presets.json", "."),
    ("HOW_TO_USE.txt", "."),
]

# python-docx ships the default .docx package it builds every document from as
# package data, and PyInstaller does not pick up package data by itself.
# Without this, build_template fails at Document() inside the exe.
hiddenimports = [
    "docx",
    "tkinter",
    "tkinter.filedialog",
]

a = Analysis(
    ["app.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Nothing here needs a scientific stack or a GUI toolkit beyond tkinter.
    # Excluding them keeps the exe to a size that is sane to email.
    excludes=[
        "numpy", "pandas", "matplotlib", "scipy", "PIL", "PyQt5", "PyQt6",
        "PySide2", "PySide6", "IPython", "pytest", "setuptools._distutils",
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
    name="TemplateStudio",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    # A console window is deliberate. It shows the address the app is serving
    # on, tells the user where settings are saved, and gives them somewhere to
    # press Ctrl+C. Closing the window stops the server, which is what someone
    # testing a local tool expects.
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
