# -*- coding: utf-8 -*-
"""冲击吸收 CSV 称量记录自检：python test_impact_csv.py"""
import atexit
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from SequenceMaster import (_read_weighing_records, _variant_project_match,
                            _extra_field_by_variant)

SRC = None  # 测试自带固定数据，不依赖用户随时替换的 谱图/冲击吸收数据导出.csv
_CSV = ("称样时间,样品编号,检测项目,试样描述,第一次冲击结果(N),第二次冲击结果(N),第三次冲击结果(N),Fc参照力值(N)\n"
        "2026/8/20,tn26080619,冲击吸收|0℃,,3951.16,4162.61,4138.01,6589.17\n"
        "2026/8/20,tn26080619,冲击吸收|23℃,,3522.05,3834.13,3797.36,6589.17\n"
        "2026/8/20,tn26080619,冲击吸收|50℃,,3816.74,4047.56,4085.83,6589.17\n"
        "2026/8/20,tn26080911,冲击吸收|0℃,,4898.65,4921.15,4958.41,6589.17\n"
        "2026/8/20,tn26080911,冲击吸收|23℃,,4684.64,4781.37,4678.92,6589.17\n"
        "2026/8/20,tn26080911,冲击吸收|50℃,,4553.94,4781.37,4848.14,6589.17\n")


def _src():
    global SRC
    if SRC is None:
        fd, SRC = tempfile.mkstemp(suffix=".csv")
        os.close(fd)
        with open(SRC, "w", encoding="gbk", newline="") as f:  # 设备导出为 GBK
            f.write(_CSV)
        atexit.register(os.remove, SRC)
    return SRC


def test_read():
    m, err = _read_weighing_records(_src())
    assert err is None, err
    assert sorted(m) == ["TN26080619", "TN26080911"], m.keys()  # 小写键统一大写
    e = m["TN26080619"]
    assert e["masses"] == [] and e["extra"] == {}, (e["masses"], e["extra"])  # 无称样量、不走旧 extra
    assert e["time"] is not None and e["time"].strftime("%Y-%m-%d") == "2026-08-20", e["time"]
    assert [r["proj"] for r in e["rows"]] == ["冲击吸收|0℃", "冲击吸收|23℃", "冲击吸收|50℃"]
    r0 = e["rows"][0]
    assert r0["vals"]["第一次冲击结果(N)"] == "3951.16"
    assert r0["vals"]["Fc参照力值(N)"] == "6589.17"
    assert "试样描述" not in r0["vals"]  # 保留列不进 vals


def test_read_old_format():  # 无「检测项目」列 → 旧 extra 通道不变(回归)
    fd, p = tempfile.mkstemp(suffix=".csv")
    os.close(fd)
    try:
        with open(p, "w", encoding="utf-8-sig", newline="") as f:
            f.write("称样时间,样品编号,浸泡面积s (cm2)\n2026/8/20,tn26080619,1.5\n")
        m, err = _read_weighing_records(p)
        assert err is None, err
        e = m["TN26080619"]
        assert e["extra"] == {"浸泡面积s (cm2)": ["1.5"]} and e["rows"] == []
    finally:
        os.remove(p)


def test_match():
    assert _variant_project_match("冲击吸收|0℃", "冲击吸收0℃")
    assert _variant_project_match("冲击吸收|23℃", "冲击吸收（23℃）")  # 全角括号
    assert _variant_project_match("冲击吸收｜50℃", "冲击吸收(50℃)")  # 全角竖线
    assert not _variant_project_match("冲击吸收|0℃", "冲击吸收23℃")  # 条件不同
    assert not _variant_project_match("冲击吸收|0℃", "冲击吸收50℃")  # 0℃ 不是 50℃ 的子串(数字边界)
    assert not _variant_project_match("冲击吸收|0℃", "拉伸强度")  # 基名不同
    assert _variant_project_match("", "任意")  # 空值=全中
    assert _variant_project_match("冲击吸收", "冲击吸收0℃")  # 无| 按包含


def test_route():
    m, _ = _read_weighing_records(_src())
    rows = {"TN26080619001": m["TN26080619"]["rows"]}
    p2s = {"101": "TN26080619001", "102": "TN26080619001", "103": "TN26080619001"}
    p2n = {"101": "冲击吸收0℃", "102": "冲击吸收23℃", "103": "冲击吸收50℃"}
    recs = [{"projectId": "101", "serialNumber": 1},
            {"projectId": "102", "serialNumber": 1},
            {"projectId": "103", "serialNumber": 1}]
    for hdr, want in [("第一次冲击结果(N)", ["3951.16", "3522.05", "3816.74"]),
                      ("第二次冲击结果(N)", ["4162.61", "3834.13", "4047.56"]),
                      ("Fc参照力值(N)", ["6589.17"] * 3)]:
        got = [b.get() for b in _extra_field_by_variant(recs, p2s, p2n, rows, hdr, hdr)]
        assert got == want, (hdr, got)
    # 项目名匹配不上该行变体 → 保留记录原值
    recs_x = [{"projectId": "999", "serialNumber": 1, "第三次冲击结果(N)": "旧值"}]
    got = [b.get() for b in _extra_field_by_variant(
        recs_x, {"999": "TN26080619001"}, {"999": "拉伸强度"}, rows,
        "第三次冲击结果(N)", "第三次冲击结果(N)")]
    assert got == ["旧值"], got


if __name__ == "__main__":
    test_read()
    test_read_old_format()
    test_match()
    test_route()
    print("all ok")
