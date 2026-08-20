# 只读探针：查一个样品，dump 受理时间相关字段的实际值
import json, sys
from login import MultiUserLoginSystem
from detection_entry_api import DetectionAPI

ls = MultiUserLoginSystem()
if not ls.load_session() or not ls.current_user:
    print("会话不可用"); sys.exit(1)
api = DetectionAPI(ls)

rows = api.query_samples_by_conditions(sample_code="TS26082291", exact_match=False,
                                       log_func=lambda m: None)
print(f"记录数: {len(rows)}")
seen = set()
for p in rows[:5]:
    raw = p.get("_raw") or {}
    key = (p.get("sampleCode"), raw.get("acceptTime"))
    if key in seen:
        continue
    seen.add(key)
    print(json.dumps({
        "sampleCode": p.get("sampleCode"),
        "checkInStatus": p.get("checkInStatus"),
        "acceptTime": raw.get("acceptTime"),
        "analyzeStartDateBy": raw.get("analyzeStartDateBy"),
        "analyzeEndDateBy": raw.get("analyzeEndDateBy"),
        "createDatetime": raw.get("createDatetime"),
        "requireTime": raw.get("requireTime"),
    }, ensure_ascii=False, indent=1))
