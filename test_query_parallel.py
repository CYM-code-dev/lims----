# -*- coding: utf-8 -*-
"""_query_samples_parallel 分发归属自检：
报验编号级编号(称样记录只写报验编号，如 TS26071303)代表 001，必须命中该报验编号的 001 记录
且排除兄弟小号(003)；完整样品号行为不变。复现 苯系物-025 20260817 0/27 命中事故。
运行: .venv/Scripts/python test_query_parallel.py
"""
import io
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from SequenceMaster import SequenceMaster

RECORDS = [  # 同一报验编号下 2 个小号、共 3 条项目(模拟 LIMS 全码返回)
    {"sampleCode": "TS26071303001", "projectId": 11, "sampleId": 1},
    {"sampleCode": "TS26071303001", "projectId": 12, "sampleId": 1},
    {"sampleCode": "TS26071303003", "projectId": 13, "sampleId": 3},
]


class _Api:
    @staticmethod
    def query_samples_by_conditions(**k):
        return list(RECORDS)


sm = object.__new__(SequenceMaster)
sm.api = _Api()
_log = lambda *a, **k: None

# 用例1：报验编号级(无小号=代表001) → 命中 001 的 2 条，排除 003（修复前 0 命中）
o1 = sm._query_samples_parallel(["TS26071303"], _log)
assert [p["projectId"] for p in o1.get("TS26071303", [])] == [11, 12], o1

# 用例2：完整样品号 → 仍只精确匹配自己的记录
o2 = sm._query_samples_parallel(["TS26071303003"], _log)
assert [p["projectId"] for p in o2.get("TS26071303003", [])] == [13], o2

# 用例3：同报验编号混入(报验编号级 + 完整码) → 去重 1 次请求，各自正确分发
o3 = sm._query_samples_parallel(["TS26071303", "TS26071303003"], _log)
assert [p["projectId"] for p in o3["TS26071303"]] == [11, 12], o3
assert [p["projectId"] for p in o3["TS26071303003"]] == [13], o3

# 用例4：方法池(全码键)与并行查询(报验编号键)各存一份同一项目 → _collect_sample_projects
# 扫描按 projectId 去重，每项目只留一份(苯系物-025 105956 重跑事故：4 项目应为 2)
sm._filter_projects_by_method = lambda projects, row, log: (projects, None)
pool = {"TS26071303001": list(RECORDS[:2]), "TS26071303": list(RECORDS[:2])}
items, err = sm._collect_sample_projects(0, {"method_file": ""}, "TS26071303", [], {}, _log, pool)
assert err is None and [it["project"]["projectId"] for it in items] == [11, 12], (err, items)

print("PASS 4/4")
