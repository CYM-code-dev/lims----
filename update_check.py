"""更新检查/下载共用逻辑：launcher（onefile）与 entry（app 内）都 import 本模块。
仅标准库——launcher.spec 不引入任何第三方依赖。"""
import hashlib
import json
import os
import urllib.request

# 部署时改为实际服务器地址；本机闭环测试用环境变量 LIMS_UPDATE_URL 覆盖
UPDATE_BASE = "http://10.1.93.25:8010"


def update_base():
    return os.environ.get("LIMS_UPDATE_URL") or UPDATE_BASE


def fetch_manifest(timeout=5):
    """拉 manifest.json；连不上/坏 JSON 返回 None（调用方据此放行当前版本）。"""
    try:
        url = update_base().rstrip("/") + "/manifest.json"
        with urllib.request.urlopen(url, timeout=timeout) as r:
            return json.loads(r.read().decode("utf-8"))
    except Exception:
        return None


def ver_tuple(v):
    return tuple(int(x) for x in str(v).split("."))


def is_behind(manifest, local_ver):
    """服务器版本 > 本地版本 → True；数据异常一律 False（不拦人）。"""
    try:
        return ver_tuple(manifest["version"]) > ver_tuple(local_ver)
    except Exception:
        return False


def download_and_verify(zip_name, sha256, dest, timeout=120):
    """流式下载 zip 到 dest 并校验 sha256；成功 True，失败 False（不留坏文件）。"""
    try:
        url = update_base().rstrip("/") + "/" + zip_name
        h = hashlib.sha256()
        with urllib.request.urlopen(url, timeout=timeout) as r, open(dest, "wb") as f:
            while True:
                chunk = r.read(1 << 16)
                if not chunk:
                    break
                f.write(chunk)
                h.update(chunk)
        if h.hexdigest() != str(sha256 or "").lower():
            os.remove(dest)
            return False
        return True
    except Exception:
        try:
            os.remove(dest)
        except OSError:
            pass
        return False


if __name__ == "__main__":
    # 自检：版本比较 + manifest 拉取（无服务器时打印 None）
    assert ver_tuple("1.0.1") > ver_tuple("1.0.0")
    assert ver_tuple("2.0") > ver_tuple("1.9.9")
    assert not is_behind({"version": "1.0.0"}, "1.0.0")
    assert not is_behind({}, "1.0.0")
    print("update_check self-check OK; manifest:", fetch_manifest(timeout=3))
