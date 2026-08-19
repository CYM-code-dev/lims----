# app.spec — onedir，入口 entry.py → app.exe
# datas：icons/ 整目录 → _internal/icons（各处 iconbitmap 经 paths.resource /
# login.find_icon_file 的 icons/ 兜底命中）；MethodRule Editor.py 落 _internal 根，
# 供 SequenceMaster 的 __file__ 同级查找命中；
# ttkbootstrap 的主题字体/图标等非 .py 资源需显式收集（缺则启动即 FileNotFoundError）。
from PyInstaller.utils.hooks import collect_data_files

a = Analysis(
    ["entry.py"],
    pathex=[],
    binaries=[],
    datas=[("icons", "icons"), ("MethodRule Editor.py", ".")]
          + collect_data_files("ttkbootstrap"),
    hiddenimports=[],
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
    [],
    exclude_binaries=True,
    name="app",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    icon="icons/app.ico",
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="app",
)
