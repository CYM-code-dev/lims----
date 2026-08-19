import os
import sys

# 每台机器各自不同、随版本更新时绝不覆盖的用户数据文件
_DATA_FILES = ["users_config.json", "session_info.json", "last_config.txt", "原始记录登记.xlsx"]


def app_dir():
    """程序内容目录（随版本更新整体替换）：methods/、switch_rules.mtd 所在处。
    frozen（onedir）→ app.exe 所在的 app/；dev → 脚本根目录。
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def data_dir():
    """用户数据目录（更新时绝不覆盖）。

    优先级：环境变量 LIMS_DATA_DIR（Phase 2 launcher 传入）>
    frozen 且为 launcher+app 布局时取 app.exe 上一级（安装根）>
    dev 时取脚本同级 data/。
    """
    env = os.environ.get("LIMS_DATA_DIR")
    if env:
        d = env
    elif getattr(sys, "frozen", False):
        # launcher+app 布局：app.exe 在 <安装根>\app\ 内，数据放安装根的 data\，
        # 否则更新覆盖 app\ 时会连带丢数据。
        d = os.path.join(os.path.dirname(os.path.dirname(sys.executable)), "data")
    else:
        d = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
    os.makedirs(d, exist_ok=True)
    return d


def migrate_data_files():
    """一次性把脚本目录下历史遗留的数据文件搬进 data/（不覆盖已有）。"""
    dest = data_dir()
    src_dir = os.path.dirname(os.path.abspath(__file__))
    for name in _DATA_FILES:
        src = os.path.join(src_dir, name)
        dst = os.path.join(dest, name)
        if os.path.exists(src) and not os.path.exists(dst):
            try:
                os.replace(src, dst)
            except OSError:
                pass


# 首次 import 时自动迁移一次；迁移完成后为空操作。
migrate_data_files()
