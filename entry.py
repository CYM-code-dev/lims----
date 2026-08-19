"""打包唯一入口：默认 SequenceMaster；`app.exe detection` → 检测数据录入。
启动前门禁（防绕过 launcher 直接跑 app.exe）：服务器可达且版本落后 → 拦截退出。"""
import sys

import update_check
from version import APP_VERSION
import paths


def _gate():
    manifest = update_check.fetch_manifest(timeout=3)
    if manifest and update_check.is_behind(manifest, APP_VERSION):
        import ctypes
        ctypes.windll.user32.MessageBoxW(
            0,
            f"已有新版本 {manifest['version']}（当前 {APP_VERSION}）。\n"
            "请从桌面快捷方式或 launcher.exe 启动，将自动更新。",
            "版本过旧",
            0x30,
        )
        sys.exit(1)


def main():
    paths.set_dpi_awareness()
    if "--selfcheck" in sys.argv:
        import SequenceMaster
        SequenceMaster._selfcheck()
        sys.exit(0)
    _gate()
    if len(sys.argv) > 1 and sys.argv[1] == "detection":
        sys.argv = [sys.argv[0]]  # 剥掉调度参数，detection_entry_main 的 argparse 不认识它
        import detection_entry_main
        detection_entry_main.main()
    else:
        import SequenceMaster
        SequenceMaster.main()


if __name__ == "__main__":
    main()
