# -*- coding: utf-8 -*-
"""_resolve_equipment_config 混标批自检：
同行样品分属不同标准(前段 GB 36246、TN26080640 为 GB 18583)时，设备查询须逐样品尝试至命中，
并按子方法ID(decideProjectMethodId)逐个取设备合并，而非只查首个样品+混 sp_ids 调 getOcExperiment。
运行: .venv/Scripts/python test_equipment_query.py
"""
import io
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from SequenceMaster import SequenceMaster


def _proj(pid, mid, name, std, code):
    return {"projectId": pid, "sampleId": 9, "sampleCode": code, "projectName": name,
            "standardNo": std, "decideProjectMethodId": mid, "oldSampleProjectId": None}


P36246 = [_proj(1, 4678, "苯", "GB 36246-2018 6.15.2", "TN26080626001"),
          _proj(2, 4679, "甲苯、二甲苯及乙苯总和", "GB 36246-2018 6.15.2", "TN26080626001")]
P18583 = [_proj(3, 1560, "苯", "GB 18583-2008 附录B", "TN26080640001"),
          _proj(4, 1561, "甲苯和二甲苯总和", "GB 18583-2008 附录C", "TN26080640001")]


class _API:
    def __init__(self, by_code):
        self.by_code = by_code
        self.calls = []

    def query_samples_by_conditions(self, sample_code=None, exact_match=False, **k):
        self.calls.append(("q", sample_code))
        return list(self.by_code.get(sample_code) or [])

    def get_detection_equipment(self, mid, log_func=None):
        self.calls.append(("e", mid))
        return {"raw_data": [{"usedCategory": "检测设备", "no": f"CK-{mid}", "name": "气相色谱仪"}]}

    _std_no_name_to_id = {"GB 18583-2008 附录B": 1560, "GB 18583-2008 附录C": 1561}

    def _prefetch_std_no_name_ids(self, names, log_func=None):
        pass

    def get_method_id_by_standard_no_name(self, name, log_func=None):
        return self._std_no_name_to_id.get(name)


def _make(by_code, rules):
    app = object.__new__(SequenceMaster)
    app.api = _API(by_code)
    app._env_eq_cfg_cache = {}
    app._row_sample_codes = lambda row, log: (list(by_code), None)
    app._read_query_rules = lambda mf: rules
    app._read_switch_rules = lambda mf: []
    app._read_exclusion_rules = lambda mf: []
    return app


RULES_18583 = [{"project": "苯", "method": "GB 18583-2008 附录B", "cancel_test": False},
               {"project": "甲苯和二甲苯总和", "method": "GB 18583-2008 附录C", "cancel_test": False}]
logs = []

# 用例1(混批)：首个样品是 GB 36246(18583 规则滤空)，须跳过并命中 TN26080640，
# 设备来自两个子方法 1560+1561 的并集
app = _make({"TN26080626001": P36246, "TN26080640001": P18583}, RULES_18583)
codes = app._resolve_equipment_config({"method_file": "x"}, logs.append)["raw_data"]
got = [e["no"] for e in codes]
assert sorted(set(got)) == ["CK-1560", "CK-1561"], got
assert ("q", "TN26080640001") in app.api.calls, "未尝试后段样品"
assert not any("getOcExperiment" in str(c) or c[0] == "e" and c[1] not in (1560, 1561)
               for c in app.api.calls), "设备来源不是子方法ID"
assert any("不在此样品项目中" in m and "TN26080626001" in m for m in logs), logs

# 缓存：二次调用零 API 调用
n = len(app.api.calls)
assert app._resolve_equipment_config({"method_file": "x"}, logs.append) is not None
assert len(app.api.calls) == n, "缓存未生效"

# 用例2(回归)：单样品单子方法，设备来源等价(该 mid 的设备)
app2 = _make({"TN26080640001": P18583[:1]}, RULES_18583[:1])
codes2 = app2._resolve_equipment_config({"method_file": "x"}, lambda m: None)["raw_data"]
assert [e["no"] for e in codes2] == ["CK-1560"], codes2

# 用例3(全不命中)：返回 None 且不抛异常，逐样品都探测过
app3 = _make({"TN26080626001": P36246}, RULES_18583)
assert app3._resolve_equipment_config({"method_file": "x"}, lambda m: None) is None

print("PASS 3/3")
