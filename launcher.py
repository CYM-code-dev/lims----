"""launcher：安装器 + 自动更新引导（onefile，仅标准库 + update_check）。
不带参数默认装 C:\\LIMS数据登记\\；`launcher.exe D:\\LIMS` 装指定目录。
已安装则启动前检查更新：落后→下载校验→备份旧 app\\→换新；服务器连不上→放行当前版本并提示。"""
import ctypes
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile

import update_check

APP_NAME = "LIMS数据登记"
DEFAULT_ROOT = r"C:\LIMS数据登记"
MB = ctypes.windll.user32.MessageBoxW


def _msg(text, style=0x40):
    MB(0, text, APP_NAME, style)


def _nc(p):
    return os.path.normcase(os.path.abspath(p))


def _desktop_shortcut(root):
    """桌面建快捷方式指向 launcher.exe（PowerShell WScript.Shell，覆盖同名，幂等）。
    仅首次安装调用；失败静默——快捷方式不是关键路径。"""
    launcher = os.path.join(root, "launcher.exe")
    ps = (
        "$d=[Environment]::GetFolderPath('Desktop');"
        "$s=(New-Object -ComObject WScript.Shell).CreateShortcut(\"$d\\%s.lnk\");"
        "$s.TargetPath='%s';$s.WorkingDirectory='%s';$s.Save()"
    ) % (APP_NAME, launcher, root)
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
            capture_output=True, timeout=30,
        )
    except Exception:
        pass


def _apply_release(zip_path, root):
    """解压发行 zip（顶层为 app/）到安装根。"""
    with zipfile.ZipFile(zip_path) as z:
        z.extractall(root)


def _download_app(manifest):
    """下载并校验最新 app zip，返回 zip 路径；失败返回 None。"""
    tmp = os.path.join(tempfile.gettempdir(), "lims_app_%s.zip" % manifest["version"])
    if not update_check.download_and_verify(
            manifest.get("zip"), manifest.get("sha256"), tmp):
        _msg("下载失败或校验不通过（网络中断/文件损坏），未做任何改动，请重试。", 0x10)
        return None
    return tmp


def _launch(app_exe, root):
    subprocess.Popen([app_exe], cwd=root)


def _do_update(manifest, root, app_dir, app_exe):
    """下载→备份旧 app\\→换新。失败时从备份还原，绝不让用户无程序可用。"""
    zip_path = _download_app(manifest)
    if not zip_path:
        _launch(app_exe, root)  # 旧版先放行；下次启动重试更新
        return 1

    backup = os.path.join(root, "app_backup")
    try:
        shutil.rmtree(backup, ignore_errors=True)
        os.rename(app_dir, backup)
        _apply_release(zip_path, root)
    except Exception as e:
        # 换新失败：清掉残缺 app\，从备份还原
        shutil.rmtree(app_dir, ignore_errors=True)
        try:
            os.rename(backup, app_dir)
        except OSError:
            pass
        _msg("更新失败：%s\n已还原旧版本。" % e, 0x10)
        _launch(app_exe, root)
        return 1
    _launch(app_exe, root)
    return 0


def main():
    root = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_ROOT
    app_dir = os.path.join(root, "app")
    app_exe = os.path.join(app_dir, "app.exe")
    version_file = os.path.join(app_dir, "VERSION")
    target_exe = os.path.join(root, "launcher.exe")
    exe = _nc(sys.executable if getattr(sys, "frozen", False) else sys.argv[0])

    first_install = not os.path.exists(version_file)

    # 自安装：不在安装根的 launcher.exe（如 %TEMP% 里 curl 下来的）先拷过去。
    # onefile 运行的是临时解压副本，拷贝原件无文件锁问题。
    if exe != _nc(target_exe):
        os.makedirs(root, exist_ok=True)
        if not os.path.exists(target_exe):
            shutil.copy2(exe, target_exe)

    manifest = update_check.fetch_manifest(timeout=10 if first_install else 5)

    if first_install:
        if not manifest:
            _msg("无法连接更新服务器，请检查网络后重新运行安装命令。", 0x10)
            return 1
        _msg("正在安装 %s v%s，请稍候 ..." % (APP_NAME, manifest["version"]))
        zip_path = _download_app(manifest)
        if not zip_path:
            return 1
        try:
            _apply_release(zip_path, root)
        except Exception as e:
            _msg("安装失败：%s" % e, 0x10)
            return 1
        _desktop_shortcut(root)
        _launch(app_exe, root)
        return 0

    # 日常启动：读本地版本，检查更新
    try:
        local = open(version_file, encoding="utf-8").read().strip()
    except OSError:
        local = "?"
    if manifest is None:
        _msg("无法连接更新服务器，以当前版本 v%s 启动。" % local, 0x30)
        _launch(app_exe, root)
        return 0
    if update_check.is_behind(manifest, local):
        _msg("发现新版本 v%s（当前 v%s），点击确定自动更新。" % (manifest["version"], local))
        return _do_update(manifest, root, app_dir, app_exe)
    _launch(app_exe, root)
    return 0


if __name__ == "__main__":
    sys.exit(main())
