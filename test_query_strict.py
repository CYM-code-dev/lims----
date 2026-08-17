# -*- coding: utf-8 -*-
"""query_samples_by_conditions 精确模式严格过滤自检：
完整样品号未命中时必须返回空(而非整个报验编号全集)，报验编号级查询仍返回全集。
复现 XRF 0813 事故：兜底吞下兄弟样品 2797 条导致谱图错绑。运行: .venv/Scripts/python test_query_strict.py
"""
import io
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from detection_entry_api import DetectionAPI

DNO = "TS26071303"
RECORDS = [  # 同一报验编号下 2 个小号、共 3 条项目
    {"detectionNo": DNO, "smallNo": "001", "id": 11, "sampleId": 1},
    {"detectionNo": DNO, "smallNo": "001", "id": 12, "sampleId": 1},
    {"detectionNo": DNO, "smallNo": "003", "id": 13, "sampleId": 3},
]


class _Resp:
    status_code = 200

    @staticmethod
    def json():
        return {"success": True, "resultData": {"voList": RECORDS}}


class _Sess:
    calls = 0

    @classmethod
    def get(cls, *a, **k):
        cls.calls += 1
        return _Resp()


class _Login:
    current_user = {"name": "tester"}
    session = _Sess
    base_url = "http://x"


def _make():
    api = object.__new__(DetectionAPI)
    api.login_system = _Login()
    api.get_user_pid = lambda: "1"
    api.get_user_pname = lambda: "t"
    api.get_user_login_id = lambda: "1"
    return api


api = _make()

# 用例1：完整样品号命中 → 早退只返回该样品 2 条
r1 = api.query_samples_by_conditions(sample_code=f"{DNO}001", exact_match=True)
assert [p["projectId"] for p in r1] == [11, 12], r1

# 用例2：完整样品号未命中(如已登记) → 空；旧代码会返回整个报验编号全集 3 条
r2 = api.query_samples_by_conditions(sample_code=f"{DNO}170", exact_match=True)
assert r2 == [], r2

# 用例3：报验编号级查询(并行分发路径) → 仍返回全集 3 条
r3 = api.query_samples_by_conditions(sample_code=DNO, exact_match=True)
assert len(r3) == 3, r3

print(f"PASS 3/3 (http calls={_Sess.calls})")
