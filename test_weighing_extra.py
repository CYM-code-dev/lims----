"""称量记录同名列/备注自检：模板 xlsx 读取(extra/remarks/称样量兜底守卫)、平行合并透传、
_vol_field_by_project 合并列语义(一行两实验共用/两行各行各值)。无网络/GUI。
运行：.venv/Scripts/python test_weighing_extra.py"""
import os
import tempfile

import openpyxl

from SequenceMaster import (_merge_parallel_groups, _read_weighing_records,
                            _vol_field_by_project, _Box)

_HDRS = ["称样时间", "样品编号", "试样描述", "浸泡面积s (cm2)", " 浸泡体积V (mL)", "解析", "备注"]


def _make_xlsx(rows):
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(_HDRS)
    for r in rows:
        ws.append(r)
    wb.save(path)
    return path


def test_read_extra_and_remarks():
    path = _make_xlsx([
        ["2026-08-17", "TS26081242001", "红色颗粒", 110.0, 120, None, "第1次条件"],   # 一行=合并列
        [None, "TS26081243001", "胶水", 110, 120, None, "第1次条件"],                 # 两行=实验1/2
        [None, "TS26081243001", None, None, 125, None, "第2次条件"],                 # 空格=沿用上行
    ])
    try:
        m, err = _read_weighing_records(path)
        assert err is None, err
        a = m["TS26081242001"]
        assert a["extra"]["浸泡面积s (cm2)"] == ["110"] and a["extra"]["浸泡体积V (mL)"] == ["120"]  # 表头strip；110.0→"110"
        assert a["remarks"] == ["第1次条件"]
        b = m["TS26081243001"]
        assert b["extra"]["浸泡面积s (cm2)"] == ["110"]          # 第2行空格不收 → 沿用
        assert b["extra"]["浸泡体积V (mL)"] == ["120", "125"]    # 两行各异
        assert b["remarks"] == ["第1次条件", "第2次条件"]
        assert a["masses"] == [] and b["masses"] == []           # 兜底守卫：D列(浸泡面积)不当称样量
    finally:
        os.unlink(path)


def test_mass_col_still_read_when_headered():
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["称样时间", "样品编号", "试样描述", "称样量", "备注"])
    ws.append(["2026-08-17", "TN01A", "颗粒", 0.5448, "备注1"])
    wb.save(path)
    try:
        m, _ = _read_weighing_records(path)
        assert m["TN01A"]["masses"] == [0.5448] and m["TN01A"]["remarks"] == ["备注1"]
        assert not m["TN01A"]["extra"]  # 保留列不进 extra
    finally:
        os.unlink(path)


def test_merge_parallel_keeps_extra():
    wmap = {
        "TN26080729A": {"masses": [0.5], "cells": [], "time": None, "desc": "颗粒",
                        "remarks": ["r1"], "extra": {"浸泡面积s (cm2)": ["110"]}},
        "TN26080729B": {"masses": [0.6], "cells": [], "time": None, "desc": "",
                        "remarks": ["r2"], "extra": {"浸泡面积s (cm2)": ["111"], "浸泡体积V (mL)": ["125"]}},
    }
    out = _merge_parallel_groups(wmap)
    e = out["TN26080729"]
    assert e["masses"] == [0.5, 0.6]
    assert e["remarks"] == ["r1", "r2"]
    assert e["extra"]["浸泡面积s (cm2)"] == ["110", "111"]
    assert e["extra"]["浸泡体积V (mL)"] == ["125"]
    assert out.get("TN26080729001") is e  # 短报验编号键：按全码(补001)容错查到同一条目


def test_vol_field_merged_vs_split():
    records = [
        {"projectId": "p1", "serialNumber": 1},
        {"projectId": "p1", "serialNumber": 2},
    ]
    pid2sc = {"p1": "TS001"}
    # 一行：实验1/2 共用同值(前端合并列语义)
    out = _vol_field_by_project(records, pid2sc, {"TS001": ["110"]}, "c1")
    assert [b.get() for b in out] == ["110", "110"]
    # 两行：实验1/2 各取各行；无值样品保留记录原值
    out = _vol_field_by_project(records, pid2sc, {"TS001": ["110", "111"]}, "c1")
    assert [b.get() for b in out] == ["110", "111"]
    records2 = [{"projectId": "p2", "serialNumber": 1, "c1": "旧值"}]
    out = _vol_field_by_project(records2, {"p2": "TS002"}, {"TS001": ["110"]}, "c1")
    assert out[0].get() == "旧值"


if __name__ == "__main__":
    test_read_extra_and_remarks()
    test_mass_col_still_read_when_headered()
    test_merge_parallel_keeps_extra()
    test_vol_field_merged_vs_split()
    print("test_weighing_extra: all OK")
