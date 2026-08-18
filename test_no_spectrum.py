"""无需谱图方法自检：precheck 放行 + _resolve_samples 无谱图路径时按称量记录展开。无网络/GUI。
运行：.venv/Scripts/python test_no_spectrum.py"""
import os
import tempfile
import types

import yaml

from SequenceMaster import SequenceMaster


def _make_method(mode):
    fd, path = tempfile.mkstemp(suffix=".yaml")
    os.close(fd)
    with open(path, "w", encoding="utf-8") as f:
        yaml.dump({"spectrum_upload_settings": {"upload_mode": mode}}, f)
    return path


def test_no_spectrum_flag():
    m = _make_method("no_spectrum")
    m2 = _make_method("local_upload")
    try:
        ns = types.SimpleNamespace(_method_no_spectrum=SequenceMaster._method_no_spectrum)
        assert SequenceMaster._method_no_spectrum(ns, m) is True
        assert SequenceMaster._method_no_spectrum(ns, m2) is False
        assert SequenceMaster._method_no_spectrum(ns, "") is False
    finally:
        os.unlink(m)
        os.unlink(m2)


def test_resolve_from_weighing_record_only():
    ns = types.SimpleNamespace(_method_no_spectrum=lambda mf: True)
    wmap = {"TS26081242001": {"masses": [], "remarks": ["r"], "extra": {}}}
    row = {"method_file": "m.yaml", "spectrum_path": "", "sample_code": ""}
    samples, err = SequenceMaster._resolve_samples(ns, row, wmap, lambda *a: None)
    assert err is None and samples == [("TS26081242001", [])]
    # 有称量记录但方法非无需谱图 → 仍要求谱图
    ns2 = types.SimpleNamespace(_method_no_spectrum=lambda mf: False)
    _, err = SequenceMaster._resolve_samples(ns2, row, wmap, lambda *a: None)
    assert err == "谱图路径无效或目录内无PDF"
    # 无需谱图但无称量记录 → 明确报错
    _, err = SequenceMaster._resolve_samples(ns, row, None, lambda *a: None)
    assert "称量记录无编号" in err


if __name__ == "__main__":
    test_no_spectrum_flag()
    test_resolve_from_weighing_record_only()
    print("test_no_spectrum: all OK")
