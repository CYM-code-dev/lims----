"""发布脚本（dev 机跑）：两次 PyInstaller → 组装 dist/release（客户端安装后的布局）
→ 产 dist/publish（上传服务器的 3 个文件：launcher.exe / app-{ver}.zip / manifest.json）。
用法：.venv/Scripts/python publish.py
"""
import hashlib
import json
import os
import shutil
import subprocess
import zipfile

import version

ROOT = os.path.dirname(os.path.abspath(__file__))
VENV_PYINSTALLER = os.path.join(ROOT, ".venv", "Scripts", "pyinstaller.exe")


def sh(*args):
    print(">", *args)
    subprocess.check_call(args, cwd=ROOT)


def sha256_of(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def build_release_layout(dist, ver):
    """dist/release/LIMS数据登记/ = 客户端安装后的目录样子（本机测试用）。"""
    rel = os.path.join(dist, "release", "LIMS数据登记")
    os.makedirs(rel)
    shutil.copy2(os.path.join(dist, "launcher.exe"), rel)
    shutil.copytree(os.path.join(dist, "app"), os.path.join(rel, "app"))
    shutil.copytree(os.path.join(ROOT, "methods"), os.path.join(rel, "app", "methods"),
                    ignore=shutil.ignore_patterns("运行日志_*"))  # 运行日志不进发行包
    shutil.copy2(os.path.join(ROOT, "switch_rules.mtd"),
                 os.path.join(rel, "app", "switch_rules.mtd"))
    with open(os.path.join(rel, "app", "VERSION"), "w", encoding="utf-8") as f:
        f.write(ver)
    return rel


def build_publish(dist, rel, ver):
    """dist/publish/ = 上传服务器的全部内容。zip 顶层为 app/，供 launcher 解压。"""
    pub = os.path.join(dist, "publish")
    os.makedirs(pub)
    shutil.copy2(os.path.join(dist, "launcher.exe"), pub)
    shutil.copy2(os.path.join(ROOT, "server_setup.bat"), pub)
    shutil.copy2(os.path.join(ROOT, "server_serve.py"), pub)
    zpath = os.path.join(pub, "app-%s.zip" % ver)
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        for base, _dirs, files in os.walk(os.path.join(rel, "app")):
            for fn in files:
                full = os.path.join(base, fn)
                z.write(full, os.path.relpath(full, rel).replace(os.sep, "/"))
    manifest = {
        "version": ver,
        "zip": os.path.basename(zpath),
        "sha256": sha256_of(zpath),
        "size": os.path.getsize(zpath),
    }
    with open(os.path.join(pub, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    return pub, manifest


def main():
    ver = version.APP_VERSION
    dist = os.path.join(ROOT, "dist")

    # 1. 清理旧产物
    for name in ("app", "launcher.exe", "release", "publish"):
        p = os.path.join(dist, name)
        if os.path.isdir(p):
            shutil.rmtree(p)
        elif os.path.exists(p):
            os.remove(p)
    shutil.rmtree(os.path.join(ROOT, "build"), ignore_errors=True)

    # 2. 打包（onedir app + onefile launcher）
    sh(VENV_PYINSTALLER, "--noconfirm", "app.spec")
    sh(VENV_PYINSTALLER, "--noconfirm", "launcher.spec")

    # 3./4. 组装
    rel = build_release_layout(dist, ver)
    pub, manifest = build_publish(dist, rel, ver)

    print("release 布局:", rel)
    print("服务器上传物:", pub, "->", sorted(os.listdir(pub)))
    print("manifest:", json.dumps(manifest, ensure_ascii=False))


if __name__ == "__main__":
    main()
