# PyInstaller spec for FPC_Preview.exe
# Run via build.bat — do not run directly with pyinstaller from this folder.

from pathlib import Path

block_cipher = None

repo_root = str(Path(SPECPATH).parent)
pandoc_src = r"C:\Users\rbhall\AppData\Local\Pandoc\pandoc.exe"

a = Analysis(
    [str(Path(repo_root) / "preview.py")],
    pathex=[repo_root],
    binaries=[
        (pandoc_src, "."),
    ],
    datas=[
        (str(Path(repo_root) / "style" / "publication.css"), "style"),
        (str(Path(repo_root) / "style" / "rs_header"), "style/rs_header"),
    ],
    hiddenimports=[
        "utils",
        "utils.conversion",
        "pypandoc",
        "xml.etree.ElementTree",
        "zipfile",
        "tkinter",
        "tkinter.filedialog",
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=["gspread", "google"],
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
    name="FPC_Preview",
    debug=False,
    strip=False,
    upx=False,
    console=True,
    icon=None,
)
