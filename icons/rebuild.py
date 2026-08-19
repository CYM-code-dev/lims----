# -*- coding: utf-8 -*-
"""SVG → 加粗笔画 + 4×超采样 → 各尺寸独立渲染 → 组装 ICO。
小尺寸不再糊/弱：stroke 按 SW 表加粗，先渲 4× 再 LANCZOS 缩小。
运行: .venv/Scripts/python icons/rebuild.py（SVG 源在项目根目录）"""
import os
import subprocess

from PIL import Image

EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
OUT = os.path.dirname(os.path.abspath(__file__))          # icons/
ROOT = os.path.dirname(OUT)                               # 项目根
TMP = os.path.join(OUT, "_tmp")
SIZES = [16, 24, 32, 48, 64, 128, 256]
# 小尺寸笔画加粗（stroke 4/64 在 24px 下仅 1.5px 线宽，直渲太弱）
SW = {16: 5, 24: 5, 32: 5, 48: 4.5, 64: 4, 128: 4, 256: 4}
SS = 4  # 超采样倍数

HTML = """<!doctype html><style>html,body{{margin:0;padding:0;width:{s}px;height:{s}px;
overflow:hidden;background:transparent}}</style>
<img src="file:///{svg}" width="{s}" height="{s}">"""


def frame(svg, s):
    """s 尺寸帧：按 SW 加粗笔画 → 4× 渲染 → LANCZOS 缩小。"""
    sw = SW[s]
    svg_text = (open(svg, encoding="utf-8").read()
                .replace('stroke-width="4"', f'stroke-width="{sw}"'))
    tmp_svg = os.path.join(TMP, f"s{s}.svg")
    open(tmp_svg, "w", encoding="utf-8").write(svg_text)
    html = os.path.join(TMP, f"w{s}.html")
    open(html, "w", encoding="utf-8").write(
        HTML.format(s=s * SS, svg=tmp_svg.replace("\\", "/")))
    png = os.path.join(TMP, f"w{s}.png")
    subprocess.run(
        [EDGE, "--headless", "--disable-gpu", f"--screenshot={png}",
         f"--window-size={s * SS},{s * SS}", "--default-background-color=00000000",
         "--hide-scrollbars", "file:///" + html.replace("\\", "/")],
        capture_output=True, timeout=60, check=True)
    img = Image.open(png).convert("RGBA")
    assert img.size == (s * SS, s * SS), f"{s} 尺寸不符: {img.size}"
    assert img.getchannel("A").getbbox(), f"{s} 全透明"
    return img.resize((s, s), Image.Resampling.LANCZOS)


def check(ico, frames):
    """逐帧校验：ICO 里存的必须是上面的直渲像素，而不是缩放产物。"""
    for f in frames:
        s = f.size[0]
        got = Image.open(ico)
        got.size = (s, s)
        got.load()
        assert got.tobytes() == f.tobytes(), f"{s}px 帧与直渲不一致"


def build(name):
    svg = os.path.join(ROOT, f"{name}.svg")
    if not os.path.exists(svg):
        print(name, "无 svg，跳过")
        return
    frames = sorted((frame(svg, s) for s in SIZES),
                    key=lambda im: im.size[0], reverse=True)
    ico = os.path.join(OUT, f"{name}.ico")
    # 基准必须取最大帧：Pillow 门禁按基准尺寸跳过更大的 sizes
    frames[0].save(ico, format="ICO", append_images=frames[1:],
                   sizes=[(s, s) for s in SIZES])
    check(ico, frames)
    print(name, "ico ok")


os.makedirs(TMP, exist_ok=True)
for name in ["app", "method", "launcher", "detection", "device"]:
    build(name)
