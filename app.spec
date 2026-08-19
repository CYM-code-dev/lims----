# app.spec — onedir，入口 entry.py → app.exe
# datas：login.ico 与 MethodRule Editor.py 落在 _internal 根，
# 分别供 login._get_resource_path 与 SequenceMaster 的 __file__ 同级查找命中；
# ttkbootstrap 的主题字体/图标等非 .py 资源需显式收集（缺则启动即 FileNotFoundError）。
from PyInstaller.utils.hooks import collect_data_files

a = Analysis(
    ["entry.py"],
    pathex=[],
    binaries=[],
    datas=[("login.ico", "."), ("MethodRule Editor.py", ".")]
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
    icon="login.ico",
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="app",
)
