# -*- coding: utf-8 -*-
"""混标批运行优化自检：
1) _collect_sample_projects 方法过滤失败只登记 mismatched、不再逐样品串行查已登记(混标批26个≈25s)；
2) _cross_row_summary 跨行去重——他行已录入的未录入样品不算真跳过(成功运行不再误报"跳过27个")；
3) _fmt_skipped 同因归组(26 行重复长原因 → 1 行)。
运行: .venv/Scripts/python test_skip_summary.py
"""
import io
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from SequenceMaster import SequenceMaster


class _API:
    def __init__(self, pool):
        self.pool = pool
        self.already_calls = []

    def query_samples_by_conditions(self, sample_code=None, exact_match=False, check_in_status=None, **k):
        if check_in_status == "CHECK_IN_STATUS_ALREADY":
            self.already_calls.append(sample_code)
        return list(self.pool.get(sample_code) or [])

    _std_no_name_to_id = {"GB 18583-2008 附录B": 1560}

    def _prefetch_std_no_name_ids(self, names, log_func=None):
        pass

    def get_method_id_by_standard_no_name(self, name, log_func=None):
        return self._std_no_name_to_id.get(name)


RULES_18583 = [{"project": "苯", "method": "GB 18583-2008 附录B"}]


def _make(pool):
    app = object.__new__(SequenceMaster)
    app.api = _API(pool)
    app._read_query_rules = lambda mf: RULES_18583
    app._read_switch_rules = lambda mf: []
    app._read_exclusion_rules = lambda mf: []
    return app


def _proj(pid, mid, name, std, code):
    return {"projectId": pid, "sampleId": 9, "sampleCode": code, "projectName": name,
            "standardNo": std, "decideProjectMethodId": mid}


P36246 = [_proj(1, 4678, "苯", "GB 36246-2018 6.15.2", "TN26080626001")]
P18583 = [_proj(3, 1560, "苯", "GB 18583-2008 附录B", "TN26080640001")]

# 用例1：方法过滤失败 → 登记 mismatched，函数内零已登记查询(复核移交调用方批量并行)
app = _make({"TN26080626001": P36246})
mism = []
items, serr = app._collect_sample_projects(0, {"method_file": "x"}, "TN26080626001",
                                           [], {}, lambda m: None, None, mism)
assert items == [] and "不在此样品项目中" in serr, (items, serr)
assert mism == ["TN26080626001"], mism
assert app.api.already_calls == [], "函数内仍有串行已登记查询"

# 调用方复核语义：ALREADY 列表(同规则方法)过滤命中 → 该样品应细化为"已登记"
_filt, _ferr = app._filter_projects_by_method(P18583, {"method_file": "x"}, lambda *a, **k: None)
assert _filt and not _ferr, (_filt, _ferr)

# 用例2(混标批跨行)：行1录入26个GB36246样品+跳过TN640；行2录入TN640+跳过26个 → 27/27、零真跳过
gb = [f"TN2608{i:04d}001" for i in range(26)]
row1 = {"merged_list": gb, "skipped_list": [("TN26080640001", "方法文件指定的方法不在此样品项目中。")]}
row2 = {"merged_list": ["TN26080640001"], "skipped_list": [(c, "方法文件指定的方法不在此样品项目中。") for c in gb]}
m, tot, skipped, cross = app._cross_row_summary([row1, row2])
assert (m, tot) == (27, 27), (m, tot)
assert skipped == [] and len(cross) == 27, (skipped, cross)

# 用例3(真跳过保留)：无任何行录入 C/D；行内 4 样品录入 2(与旧口径 len-n_skip 一致)
row3 = {"merged_list": ["A", "B"], "skipped_list": [("C", "已登记"), ("D", "方法文件指定的方法不在此样品项目中。")]}
m, tot, skipped, cross = app._cross_row_summary([row3])
assert (m, tot) == (2, 4) and cross == [], (m, tot, skipped, cross)
assert [sc for sc, _ in skipped] == ["C", "D"], skipped

# 用例4：同因归组 26 个 → 1 行(截断到6个+等26个)；不同原因分行
lines = app._fmt_skipped([(c, "同因") for c in gb])
assert len(lines) == 1 and "等26个" in lines[0] and gb[0] in lines[0], lines
lines = app._fmt_skipped([("A", "同因"), ("B", "同因"), ("C", "别因")])
assert len(lines) == 2 and lines[0].count("、") == 1, lines

print("PASS 4/4")
