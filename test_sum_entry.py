"""总和录入自检：桩数据验证 _fill_result_from_sum 的全号派生、组分列定位、取值字段、@ 别名匹配、平行分槽。无网络/GUI。
运行：.venv/Scripts/python test_sum_entry.py"""
import types
from SequenceMaster import SequenceMaster, _is_sum_project


class _Host:
    def __init__(self):
        self.data_fields = {}


class _StubAPI:
    """记录调用参数，按样品返回预设组分记录。"""
    def __init__(self, by_sample):
        self.by_sample = by_sample
        self.calls = []

    def get_oc_compare_show_data(self, sample_code, items, log_func=None):
        self.calls.append((sample_code, list(items)))
        return self.by_sample.get(sample_code, [])


def _run(dyn_cols, records, batch_items, comp_by_sample):
    api = _StubAPI(comp_by_sample)
    self_ = types.SimpleNamespace(api=api)
    host = _Host()
    SequenceMaster._fill_result_from_sum(
        self_, host, dyn_cols, {"ocAnalysisRecordList": records}, batch_items, {}, "", {}, lambda *a: None)
    return api, host


# batch_items 样板：project 带 detectionNo/sampleSmallNo（全号源）；sample_code 为报验编号(旧字段，不应被采用)
def _item(det_no, small_no):
    return {"sample_code": det_no, "project": {"detectionNo": det_no, "sampleSmallNo": small_no}}


def test_is_sum_project():
    assert _is_sum_project("甲苯、二甲苯及乙苯总和")
    assert _is_sum_project("AfPS GS 2019:01 PAK 4项之和")
    assert _is_sum_project("18种多环芳烃总和*")
    assert not _is_sum_project("苯并[a]芘")
    assert not _is_sum_project("苯")
    print("ok _is_sum_project")


def test_full_sample_code_and_report_value():
    # Fix-A：sampleCodes 用全号 detectionNo+sampleSmallNo，不是报验编号
    dyn_cols = [
        {"columeCode": "dynamic6321", "compareShowItemName": "菲（PHE）", "compareShowTitleName": "报告值"},
        {"columeCode": "dynamic6322", "compareShowItemName": "蒽（Ant）", "compareShowTitleName": "报告值"},
        {"columeCode": "dynamic9999", "columeName": "称样量m(g)"},  # 非组分列，应被忽略
    ]
    records = [{"serialNumber": 1, "sampleCode": "TS26072901001", "projectId": 3985248}]
    comp = [
        {"projectName": "菲（PHE）", "serialNumber": 1, "reportValue": "1.0", "calculatedValue": "0.976"},
        {"projectName": "蒽（Ant）", "serialNumber": 1, "reportValue": "1.6", "calculatedValue": "1.601"},
    ]
    api, host = _run(dyn_cols, records, [_item("TS26072901", "001")], {"TS26072901001": comp})
    assert api.calls == [("TS26072901001", ["菲（PHE）", "蒽（Ant）"])], api.calls  # 全号，非 TS26072901
    assert [b.get() for b in host.data_fields["dynamic6321"]] == ["1.0"]
    assert [b.get() for b in host.data_fields["dynamic6322"]] == ["1.6"]
    assert "dynamic9999" not in host.data_fields  # 非组分列不经本方法
    print("ok 全号派生/报告值/组分列筛选")


def test_at_alias_matching():
    # Fix-B：compareShowItemName 含 @ 别名；记录 projectName 命中任一归一化别名即取值
    dyn_cols = [{"columeCode": "d1",
                 "compareShowItemName": "二苯并(a，h)蒽（DBA）@二苯并(a,h)蒽（DBA）",
                 "compareShowTitleName": "报告值"}]
    records = [{"serialNumber": 1, "sampleCode": "S1"}]
    comp = [{"projectName": "二苯并(a,h)蒽（DBA）", "serialNumber": 1, "reportValue": "0.5"}]  # 半角逗号变体
    api, host = _run(dyn_cols, records, [_item("S", "1")], {"S1": comp})
    assert api.calls == [("S1", ["二苯并(a，h)蒽（DBA）", "二苯并(a,h)蒽（DBA）"])], api.calls  # @ 拆成两条别名
    assert [b.get() for b in host.data_fields["d1"]] == ["0.5"]
    print("ok @ 别名匹配")


def test_calc_value_field():
    dyn_cols = [{"columeCode": "d1", "compareShowItemName": "菲（PHE）", "compareShowTitleName": "计算值"}]
    records = [{"serialNumber": 1, "sampleCode": "S1"}]
    comp = [{"projectName": "菲（PHE）", "serialNumber": 1, "reportValue": "1.0", "calculatedValue": "0.976"}]
    _, host = _run(dyn_cols, records, [_item("S", "1")], {"S1": comp})
    assert [b.get() for b in host.data_fields["d1"]] == ["0.976"]  # compareShowTitleName=计算值→calculatedValue
    print("ok 计算值字段")


def test_two_parallels():
    # 2 平行：菲 A=1.0/B=2.0；总和实验 2 条记录(serialNumber 1,2)
    dyn_cols = [{"columeCode": "d1", "compareShowItemName": "菲（PHE）", "compareShowTitleName": "报告值"}]
    records = [{"serialNumber": 1, "sampleCode": "S1"}, {"serialNumber": 2, "sampleCode": "S1"}]
    comp = [
        {"projectName": "菲（PHE）", "serialNumber": 1, "reportValue": "1.0"},
        {"projectName": "菲（PHE）", "serialNumber": 2, "reportValue": "2.0"},
    ]
    _, host = _run(dyn_cols, records, [_item("S", "1")], {"S1": comp})
    assert [b.get() for b in host.data_fields["d1"]] == ["1.0", "2.0"], host.data_fields["d1"]
    print("ok 多平行按 serialNumber 分槽")


if __name__ == "__main__":
    test_is_sum_project()
    test_full_sample_code_and_report_value()
    test_at_alias_matching()
    test_calc_value_field()
    test_two_parallels()
    print("sum-entry self-check OK")
