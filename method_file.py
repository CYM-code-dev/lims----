"""方法文件读写：gzip 压缩的 JSON 二进制（记事本打开为乱码，防普通用户查看/篡改）。

两层回退保证迁移期健壮：
- 后缀回退 _resolve：旧 .seq 序列里 method_file 仍指向 .yaml，迁移后实际文件是 .mtd，自动续链。
- 内容回退：gzip 解压失败则按明文 yaml 读，防止迁移遗漏导致方法文件无法加载。
"""
import gzip
import json
import os

import yaml  # 仅过渡期兜底：读尚未迁移的明文 yaml


def _resolve(path):
    """后缀回退：.yaml 不存在则尝试同名 .mtd（迁移期旧序列续链）。"""
    if path and path.endswith(".yaml") and not os.path.isfile(path):
        alt = path[:-5] + ".mtd"
        if os.path.isfile(alt):
            return alt
    return path


def load_method(path):
    """加载方法文件（gzip+JSON 二进制）。失败/空/不存在返回 {}。"""
    path = _resolve(path)
    if not path or not os.path.isfile(path):
        return {}
    try:
        with open(path, "rb") as f:
            raw = f.read()
    except Exception:
        return {}
    try:
        return json.loads(gzip.decompress(raw).decode("utf-8")) or {}
    except Exception:
        try:
            return yaml.safe_load(raw.decode("utf-8")) or {}  # 过渡期明文兜底
        except Exception:
            return {}


def save_method(path, data):
    """保存为 gzip+JSON 二进制（记事本打开为乱码）。"""
    with open(path, "wb") as f:
        f.write(gzip.compress(json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")))


if __name__ == "__main__":
    # 自检：round-trip + 后缀回退
    import shutil
    import tempfile

    d = {"query_rules": {"x": 1}, "items": [1, 2, 3], "s": "<0.004", "中文": "胶水"}
    p = tempfile.mktemp(suffix=".mtd")
    save_method(p, d)
    assert load_method(p) == d, "round-trip failed"

    # 后缀回退：只有 a.mtd 存在，load a.yaml 应自动找到 a.mtd
    pyaml = p[:-4] + ".yaml"  # 同名 .yaml（不存在文件）
    fake_yaml = pyaml.replace(".mtd", "X.mtd")  # 占位避免混淆
    assert load_method(pyaml) == d, "suffix fallback failed"

    os.remove(p)
    if os.path.exists(fake_yaml):
        os.remove(fake_yaml)
    print("method_file self-check OK")
