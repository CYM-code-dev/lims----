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


def set_dpi_awareness():
    """进程级 DPI 声明（任何 Tk 创建前调用一次，之后各机渲染行为一致）。
    重复调用/已声明/旧系统 → 静默。"""
    try:
        import ctypes
        try:
            ctypes.windll.user32.SetProcessDpiAwarenessContext(-4)  # Per-Monitor V2
        except Exception:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # System Aware
    except Exception:
        pass


def dpi_factor():
    """当前 DPI / 96（进程已声明感知后即真实值；失败回退 1.0）。"""
    try:
        import ctypes
        return ctypes.windll.user32.GetDpiForSystem() / 96.0
    except Exception:
        return 1.0


# 高 DPI/大屏机器上窗口放大倍数的折扣（1.0=完全等比；调小=高 DPI 机器窗口更紧凑）
_COMFORT = 0.75


def scaled_size(widget, w, h):
    """窗口尺寸缩放：普通屏（≤1920×1080、100% 缩放）→ 原样不变；
    高 DPI/高分屏 → 按 min(DPI, 屏幕相对 1920×1080 比例)×折扣 放大，但不低于 1.0
    （基准尺寸按内容设计，再小会裁切），最终仍夹在屏内。
    """
    f = min(dpi_factor(),
            widget.winfo_screenwidth() / 1920.0,
            widget.winfo_screenheight() / 1080.0)
    f = max(1.0, f * _COMFORT)
    w, h = min(w, 1880), min(h, 1010)
    return min(int(w * f), widget.winfo_screenwidth() - 40), \
           min(int(h * f), widget.winfo_screenheight() - 100)


def resource(name):
    """随包资源文件绝对路径（frozen → _internal；dev → 脚本根目录；
    图标已整合到 icons/ 子目录，自动兜底查找）。"""
    base = getattr(sys, "_MEIPASS", None) or app_dir()
    for p in (os.path.join(base, name), os.path.join(base, "icons", name)):
        if os.path.exists(p):
            return p
    return os.path.join(base, name)


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
