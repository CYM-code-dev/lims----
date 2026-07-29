import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import ttkbootstrap as ttkb  # 档1: 现代主题(sandstone-light)，ttk 控件自动套用
import os
import re
import sys
import json
import fnmatch
import glob
import threading
import queue
import types
import random
import yaml
import openpyxl
from datetime import datetime
from PIL import Image, ImageTk

sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from login import MultiUserLoginSystem
from detection_entry_api import DetectionAPI, build_grouped_experiment_data, _Box


class DraggableHeader:
    """可拖动的表头分隔线"""

    def __init__(self, parent, index, on_drag):
        self.parent = parent
        self.index = index
        self.on_drag = on_drag
        self.dragging = False

        # 创建分隔线 - 使用绝对定位
        self.separator = tk.Frame(parent, width=1, cursor="sb_h_double_arrow", bg="#e0e0e0")
        self.separator.place(in_=parent, x=0, y=0, relheight=1)

        # 绑定事件
        self.separator.bind("<ButtonPress-1>", self.on_press)
        self.separator.bind("<B1-Motion>", self.on_drag_motion)
        self.separator.bind("<ButtonRelease-1>", self.on_release)

        # 鼠标悬停效果
        self.separator.bind("<Enter>", self.on_enter)
        self.separator.bind("<Leave>", self.on_leave)

    def update_position(self, x_position):
        """更新分隔线位置 - 使用绝对坐标"""
        self.separator.place_configure(x=x_position)

    def on_enter(self, event):
        """鼠标进入分隔线区域"""
        self.separator.configure(bg="#3498db", width=2)

    def on_leave(self, event):
        """鼠标离开分隔线区域"""
        if not self.dragging:
            self.separator.configure(bg="#e0e0e0", width=1)

    def on_press(self, event):
        """鼠标按下分隔线"""
        self.dragging = True
        self.start_x = event.x_root
        self.separator.configure(bg="#2980b9", width=2)

    def on_drag_motion(self, event):
        """拖动分隔线"""
        if self.dragging:
            delta = event.x_root - self.start_x
            self.start_x = event.x_root
            self.on_drag(self.index, delta)

    def on_release(self, event):
        """释放鼠标"""
        self.dragging = False
        self.separator.configure(bg="#e0e0e0", width=4)


class CellSelectionManager:
    """单元格选择管理器"""

    def __init__(self):
        self.selected_cells = set()  # 存储选中的单元格 (row, col)
        self.selected_rows = set()  # 存储选中的整行
        self.drag_start = None  # 拖动起始位置
        self.dragging = False  # 是否正在拖动
        self.drag_type = None  # 拖动类型: 'row' 或 'cell'
        self.drag_column = None  # 拖动起始列（用于限制列范围）

    def clear_all_selection(self):
        """清除所有选择"""
        self.selected_cells.clear()
        self.selected_rows.clear()

    def select_cell(self, row, col):
        """选择单个单元格"""
        self.selected_cells.add((row, col))

    def deselect_cell(self, row, col):
        """取消选择单元格"""
        self.selected_cells.discard((row, col))

    def select_row(self, row):
        """选择整行"""
        self.selected_rows.add(row)

    def deselect_row(self, row):
        """取消选择行"""
        self.selected_rows.discard(row)

    def is_cell_selected(self, row, col):
        """检查单元格是否被选中"""
        return (row, col) in self.selected_cells or row in self.selected_rows

    def is_row_selected(self, row):
        """检查整行是否被选中"""
        return row in self.selected_rows

    def start_drag(self, row, col, drag_type='cell'):
        """开始拖动选择"""
        self.drag_start = (row, col)
        self.dragging = True
        self.drag_type = drag_type
        self.drag_column = col  # 记录起始列

    def update_drag(self, row, col, ctrl_pressed=False, shift_pressed=False):
        """更新拖动选择"""
        if not self.dragging or self.drag_start is None:
            return

        start_row, start_col = self.drag_start

        if self.drag_type == 'row':
            # 行拖动选择
            min_row = min(start_row, row)
            max_row = max(start_row, row)

            # 如果不是Ctrl或Shift操作，清除之前的选择
            if not ctrl_pressed and not shift_pressed:
                self.selected_rows.clear()

            for r in range(min_row, max_row + 1):
                self.selected_rows.add(r)
        else:
            # 单元格拖动选择 - 限制在当前列内
            min_row = min(start_row, row)
            max_row = max(start_row, row)

            # 使用起始列，而不是当前列，确保只选中当前列
            min_col = self.drag_column
            max_col = self.drag_column

            # 如果不是Ctrl或Shift操作，清除之前的选择
            if not ctrl_pressed and not shift_pressed:
                # 只清除当前列的选择
                cells_to_remove = [(r, c) for r, c in self.selected_cells.copy() if c == self.drag_column]
                for cell in cells_to_remove:
                    self.selected_cells.discard(cell)

            # 添加新选择 - 只选择当前列
            for r in range(min_row, max_row + 1):
                self.selected_cells.add((r, self.drag_column))

    def end_drag(self):
        """结束拖动选择"""
        self.dragging = False
        self.drag_start = None
        self.drag_type = None
        self.drag_column = None


def _format_path(value):
    """路径显示：文件名(末级目录)在前 - 目录在后，避免长路径盖住文件名"""
    v = (value or "").strip()
    if not v:
        return ""
    base = os.path.basename(v)
    d = os.path.dirname(v)
    return f"{base} - {d}" if d else base


# 称量记录列候选键名(方法表头属性「称量记录=是」标记的即称样量列；LIMS 实测键名为 isWeighing)
_WEIGHING_COL_KEYS = ("isWeighing", "称量记录", "isWeighingRecord", "weighingRecord", "isRecordColumn", "recordColumn")
_WEIGHING_TRUE = ("是", "1", "true", "True", "yes", "Y")


def _find_weighing_column(dynamic_columns):
    """返回标记为称量记录(称样量)的动态列对象；按方法表头属性「称量记录=是」识别。
    兼容中英文键名；未找到返回 None。"""
    for col in dynamic_columns or []:
        if not isinstance(col, dict):
            continue
        # 1. 直接命中候选键
        if any(str(col.get(k, "")).strip() in _WEIGHING_TRUE for k in _WEIGHING_COL_KEYS):
            return col
        # 2. 兜底：任一键值为真且键名含 称量/record/weigh
        for k, v in col.items():
            if str(v).strip() in _WEIGHING_TRUE and ("称量" in str(k) or "record" in str(k).lower() or "weigh" in str(k).lower()):
                return col
    return None


def _gen_random_mass(wp):
    """按 weighing_params 生成随机称样量：[min_value, max_value]，保留 decimal_places 位小数。"""
    try:
        lo = float(wp.get("min_value") or 0)
    except (TypeError, ValueError):
        lo = 0.0
    try:
        hi = float(wp.get("max_value") or 1)
    except (TypeError, ValueError):
        hi = lo + 1.0
    if hi <= lo:
        hi = lo + 1.0
    try:
        dp = int(wp.get("decimal_places") or 2)
    except (TypeError, ValueError):
        dp = 2
    return f"{random.uniform(lo, hi):.{dp}f}"


def _parallel_indices(records):
    """返回 (parallel_of, n_par)：parallel_of[全局记录序号]=平行索引，n_par=平行数。
    优先用记录的 serialNumber(1-based 平行号)；缺失才按 projectId 分组枚举。
    用 serialNumber 是为兼容「同 projectId 多组分」方法(如 总和：5组分×2平行=10条同 projectId)，
    按 projectId 枚举会把 10 条误当成 10 个平行。"""
    serials = [r.get("serialNumber") for r in (records or [])]
    if any(s is not None for s in serials):
        parallel_of, max_par = {}, 0
        for g, r in enumerate(records or []):
            sn = r.get("serialNumber")
            try:
                par = int(sn) - 1 if sn is not None else 0
            except (TypeError, ValueError):
                par = 0
            parallel_of[g] = par
            if par > max_par:
                max_par = par
        return parallel_of, max_par + 1
    by_proj = {}
    for g, r in enumerate(records or []):
        by_proj.setdefault(r.get("projectId"), []).append(g)
    parallel_of = {}
    for glist in by_proj.values():
        for par, g in enumerate(glist):
            parallel_of[g] = par
    n_par = (max(parallel_of.values()) + 1) if parallel_of else 1
    return parallel_of, n_par


def _mass_field_by_parallel(records, pmasses):
    """按平行值列表 pmasses(每个平行一个称样量字符串) 构造对齐 ocAnalysisRecordList 的字段列表。
    同一平行(不同项目)共用 pmasses[平行索引]；build 按记录全局位置取 list[i]。"""
    parallel_of, _ = _parallel_indices(records)
    return [_Box(pmasses[parallel_of[g]]) for g in range(len(records or []))]


def _mass_field_by_project(records, pid_to_sample, pmasses_by_sample):
    """跨样品合并提交时按 projectId -> 样品 -> 该样品 pmasses[平行索引] 构造称样量字段列表。
    不同样品用各自的称样量；同样品同平行共用一值。pid_to_sample: {projectId_str: sample_code}；
    pmasses_by_sample: {sample_code: [mass_str,...]}(按平行序)。"""
    parallel_of, _ = _parallel_indices(records)
    out = []
    for g, r in enumerate(records or []):
        sc = pid_to_sample.get(str(r.get("projectId")), "")
        pmasses = pmasses_by_sample.get(sc) or []
        par = parallel_of.get(g, 0)
        val = pmasses[par] if par < len(pmasses) else (pmasses[-1] if pmasses else "")
        out.append(_Box(val))
    return out


def _project_match(project_name, project_val):
    """单条 query rule 的 project 匹配：rule.project 空=全中；含 * 用 fnmatch；否则精确相等。"""
    pname = (project_name or "").strip()
    pv = (project_val or "").strip()
    if not pv:
        return True
    if "*" in pv:
        return fnmatch.fnmatchcase(pname, pv)
    return pname == pv


def _plan_submission_batches(query_rules, items, max_sel):
    """按 query_rules 顺序规划提交批次（纯函数，可单测）。
    items: 每项为 dict，需含 switch_mid / sample_code / projectName / wdate(称样日期,可空)。
    返回 [{"switch_mid","wdate","items","force_new"}, ...]，顺序 = 规则顺序，同规则内按 switch_mid、
    再按称样日期、再按样品。同 switch_mid 内称样日期不同(跨天)拆独立批(各自实验编号)；同日/无日期仍合并。
    input_method=方法：同(switch_mid,日期)的样品合并，max_sel 超限切片(多片 force_new=True)；
    input_method=样品：每样品各一片。无 query_rules 退化为单条空规则(全中,方法)。"""
    plan = []
    rules = query_rules or [{"project": "", "input_method": "方法"}]
    for rule in rules:
        rp = str(rule.get("project") or "").strip()
        mode = str(rule.get("input_method") or "方法").strip()
        rule_items = [it for it in items if _project_match(it.get("projectName", ""), rp)]
        if not rule_items:
            continue
        groups = {}  # switch_mid -> [items]（保序）
        for it in rule_items:
            groups.setdefault(it.get("switch_mid", ""), []).append(it)
        for mid, g_items in groups.items():
            # 称样日期二次分组：同 switch_mid 但称样日期不同(跨天)的样品拆独立批，各自生成独立实验编号；
            # 同日(含无称样时间的 "" 组)仍合并。日期取自 item["wdate"](_run_row 由称样时间归一)
            by_date = {}  # wdate -> [items]（保序）
            for it in g_items:
                by_date.setdefault(it.get("wdate", ""), []).append(it)
            sub_batches = []  # [(wdate, items, [sample 切片])]
            for d, d_items in by_date.items():
                samples = list(dict.fromkeys(it.get("sample_code", "") for it in d_items))
                if mode == "样品":
                    slices = [[sc] for sc in samples]
                elif max_sel > 0 and len(samples) > max_sel:
                    slices = [samples[i:i + max_sel] for i in range(0, len(samples), max_sel)]
                else:
                    slices = [samples]
                sub_batches.append((d, d_items, slices))
            force_new = sum(len(sl) for _, _, sl in sub_batches) > 1
            for d, d_items, slices in sub_batches:
                for sb in slices:
                    sb_set = set(sb)
                    plan.append({
                        "switch_mid": mid,
                        "wdate": d,
                        "items": [it for it in d_items if it.get("sample_code", "") in sb_set],
                        "force_new": force_new,
                    })
    return plan


def _match_filename_rule(filename_rules, project_name, pdf_paths, desc=""):
    """名称切换规则匹配（纯函数，可单测）。
    规则各非空条件均需满足(AND): project_name 关键字包含于样品项目名(空=任意项目)、
    filename 关键字出现在某谱图PDF文件名(空=任意文件)、desc 关键字(逗号分隔多个、任一)
    出现在样品"试样描述"(空=任意描述)。即"文件名"与"试样描述"都填时需同时命中。
    多规则命中取首条，均不命中返回 ''。三项匹配均大小写不敏感、子串包含。"""
    bases = [os.path.basename(p).lower() for p in (pdf_paths or []) if p]
    pn = (project_name or "").strip()
    d = (desc or "").strip().lower()
    for r in filename_rules or []:
        rp = str(r.get("project_name") or "").strip()
        if rp and rp.lower() not in pn.lower():
            continue
        fk = str(r.get("filename") or "").strip().lower()
        if fk and not any(fk in b for b in bases):
            continue
        dk = str(r.get("desc") or "")
        if dk.strip():
            tmp = dk
            for sep in (",", "，", ";", "；"):
                tmp = tmp.replace(sep, " ")
            dks = [k for k in tmp.lower().split() if k]
            if not any(k in d for k in dks):
                continue
        return str(r.get("to_id") or "").strip()
    return ""


def _rule_desc_match(filename_rules, desc):
    """混目录分流用：样品"试样描述"是否命中本方法任一 filename_rules 的 desc 关键字。
    desc 逗号分隔、任一命中、大小写不敏感(与 _match_filename_rule 的 desc 口径一致)；
    任一规则未设 desc 视为"不限描述"(命中)。全部不命中返回 False。"""
    d = (desc or "").strip().lower()
    for r in filename_rules or []:
        dk = str(r.get("desc") or "")
        if not dk.strip():
            return True
        tmp = dk
        for sep in (",", "，", ";", "；"):
            tmp = tmp.replace(sep, " ")
        if any(k and k in d for k in tmp.lower().split()):
            return True
    return False


_PARALLEL_SUFFIX_RE = re.compile(r'^([A-Za-z]+\d{8})\d{3}([A-Za-z]*)$')


def _strip_parallel_suffix(code):
    """去掉称样编号里的平行小号(001)，用于匹配无小号的谱图文件名。
    约定编号 = 字母前缀 + 8位流水 + 3位平行小号(001/002…) + 可选字母(A/B)，
    如 TN26070466001→TN26070466、TN26070474001A→TN26070474A；
    不符该结构(无小号)原样返回。注意：提交LIMS仍用带小号的原始编号(谱图文件名才不带)。"""
    m = _PARALLEL_SUFFIX_RE.match((code or "").strip())
    return (m.group(1) + m.group(2)) if m else (code or "")


_PARALLEL_LETTER_RE = re.compile(r'^(\D*\d+)([A-Za-z]+)$')


def _strip_parallel_letter(code):
    """去掉称样编号末尾的平行字母(如 TN26070474001A→TN26070474001)，用于把 A/B 平行样归为同一样品。
    末尾非字母(普通样品)原样返回。"""
    m = _PARALLEL_LETTER_RE.match((code or "").strip())
    return m.group(1) if m else (code or "")


def _merge_parallel_groups(wmap):
    """合并称样记录里的 A/B(及更多) 平行样：编号去尾字母后相同的归为一个样品，称样量按字母序
    (A→平行1、B→平行2…)拼接。合并后样品编号 = 去尾字母编号(如 TN26070474001A/B→TN26070474001)，
    作为 LIMS 查询键——平行数/实验次数由 LIMS 该样品 ocAnalysisRecordList 的 serialNumber 数决定。
    普通样品(无配对)保持不变；无称样记录返回原值。"""
    if not wmap:
        return wmap
    groups, order = {}, []
    for code in wmap:
        mc = _strip_parallel_letter(code)
        if mc not in groups:
            groups[mc] = []
            order.append(mc)
        groups[mc].append(code)
    out = {}
    for mc in order:
        members = sorted(groups[mc])  # A 在 B 前 → 平行序
        if len(members) == 1:
            out[members[0]] = wmap[members[0]]
            continue
        base = wmap[members[0]]
        masses = []
        for m in members:
            masses.extend((wmap[m].get("masses") or []))
        out[mc] = {"masses": masses, "time": base.get("time"), "desc": base.get("desc") or ""}
    return out


def _expand_parallel_records(experiment_config, pid_to_sample, par_by_sample):
    """按各样品称样量平行数扩展 ocAnalysisRecordList：样品需 N 个平行而 LIMS 仅返回更少时，
    以该样品各组分的最小 serialNumber 记录为模板、serialNumber 递增复制到 N，使每个平行都有记录槽
    (实验次数=平行数；与 LIMS 前端"加平行"生成的 row_X.0001 子行同构)。
    就地改 experiment_config['ocAnalysisRecordList']，返回 (旧条数, 新条数)。"""
    cfg = experiment_config or {}
    records = list(cfg.get("ocAnalysisRecordList") or [])
    if not records:
        return 0, 0

    def _sn(r):
        try:
            return int(r.get("serialNumber"))
        except (TypeError, ValueError):
            return 1

    pids, groups = [], {}
    for r in records:
        pid = str(r.get("projectId"))
        if pid not in groups:
            groups[pid] = []; pids.append(pid)
        groups[pid].append(r)

    expanded = []
    for pid in pids:
        recs = groups[pid]
        sc = pid_to_sample.get(pid)
        need = par_by_sample.get(sc) if sc else 0
        min_sn, cur_max = min(_sn(r) for r in recs), max(_sn(r) for r in recs)
        do_expand = bool(sc and need and need > 1 and cur_max < min_sn + need - 1)
        for r in recs:
            expanded.append(r)
            # 各组分(最小 serial)记录后紧跟其新增 serial，保持 组分×serial 交错顺序
            # (如 DBP-1,DBP-2,BBP-1,BBP-2,...)，与 LIMS 前端加平行的子行顺序一致
            if do_expand and _sn(r) == min_sn:
                for k in range(cur_max + 1, min_sn + need):
                    nr = dict(r)
                    nr["serialNumber"] = k
                    expanded.append(nr)
    if len(expanded) != len(records):
        cfg["ocAnalysisRecordList"] = expanded
    return len(records), len(expanded)


def _raw_mass_str(raw):
    """称样量原值转字符串：%g 去浮点尾噪并保留有效位(0.5655→"0.5655"，0.5→"0.5")。"""
    return f"{float(raw):.10g}"


def _read_weighing_records(path):
    """读取称量记录 xlsx：返回 ({样品编号: {masses:[float...], time, desc}}, err)。
    行序即平行序；time/desc 取该样品首行(称样时间/试样描述)。表头按列名定位，缺失按 A/B/C/D 兜底。"""
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception as e:
        return None, f"读取称量记录失败: {e}"
    ws = wb[wb.sheetnames[0]] if wb.sheetnames else None
    if ws is None:
        return None, "称量记录无工作表"
    rows = list(ws.iter_rows(values_only=True))
    header = [str(c or "").strip() for c in rows[0]] if rows else []

    def find(key, default):
        for i, h in enumerate(header):
            if key in h:
                return i
        return default

    code_col, mass_col = find("样品编号", 1), find("称样量", 2)
    time_col, desc_col = find("称样时间", 0), find("试样描述", 3)
    m = {}
    for r in rows[1:]:
        if not r:
            continue
        code = r[code_col] if code_col < len(r) else None
        if code is None or str(code).strip() == "":
            continue
        entry = m.setdefault(str(code).strip(), {"masses": [], "time": None, "desc": ""})
        # 称样量可能为空(random 模式称样量随机生成，excel 仅记录编号/试样描述)；有值才追加
        mass = r[mass_col] if mass_col < len(r) else None
        if mass is not None:
            try:
                entry["masses"].append(float(mass))
            except (TypeError, ValueError):
                pass
        if entry["time"] is None:
            entry["time"] = r[time_col] if time_col < len(r) else None
        if not entry["desc"]:
            d = r[desc_col] if desc_col < len(r) else None
            entry["desc"] = str(d).strip() if d is not None else ""
    return m, None


def _find_column_by_name(dynamic_columns, *substrs):
    """按 columeName 子串匹配返回首个动态列；未找到返回 None。"""
    for col in dynamic_columns or []:
        if not isinstance(col, dict):
            continue
        name = str(col.get("columeName") or "")
        if any(s and s in name for s in substrs):
            return col
    return None


# 温湿度记录文件路径：本地当前路径；后续迁到共享盘时改这一处即可
_ENV_RECORD_PATH = r"D:\Agent\lims数据登记\温湿度记录.xlsx"


def _date_str(v):
    """日期值归一为 'YYYY-MM-DD'：兼容 datetime / 字符串。"""
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    return str(v).strip().replace("/", "-")[:10]


def _read_env_records(path):
    """读取温湿度记录 xlsx：返回 (equip_to_room, room_env, err)。
    equip_to_room: {设备编号: 房间名称}；room_env: {(房间名称,'YYYY-MM-DD'): (温度,湿度)}。
    按 sheet 表头识别列(设备编号/房间名称/日期/温度/湿度)；当前文件仅含房间↔设备映射，温湿度值待补。"""
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception as e:
        return None, None, f"读取温湿度记录失败: {e}"
    equip_to_room, room_env = {}, {}
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        header = [str(c or "").strip() for c in rows[0]]

        def hcol(*keys):
            for k in keys:
                for i, h in enumerate(header):
                    if k in h:
                        return i
            return None

        c_room, c_eq = hcol("房间名称", "房间"), hcol("设备编号", "设备")
        c_date, c_temp, c_hum = hcol("日期", "时间"), hcol("温度"), hcol("湿度", "相对湿度")
        if c_eq is None and c_temp is None:
            continue  # 本 sheet 既无设备映射也无温湿度
        for r in rows[1:]:
            if not r:
                continue

            def cell(i):
                return r[i] if (i is not None and i < len(r)) else None

            room = str(cell(c_room)).strip() if c_room is not None and cell(c_room) is not None else ""
            eq = str(cell(c_eq)).strip() if c_eq is not None and cell(c_eq) is not None else ""
            if eq and room:
                equip_to_room[eq] = room
            if c_temp is not None and c_hum is not None and room:
                ds = _date_str(cell(c_date))
                if ds:
                    room_env[(room, ds)] = (str(cell(c_temp)).strip(), str(cell(c_hum)).strip())
    return equip_to_room, room_env, None


def _resolve_env(equip_to_room, room_env, candidates, date_str):
    """候选设备编号 → 房间 → (房间,date) 温湿度。返回 (温度,湿度) 或 (None,None)。
    匹配：设备精确→包含；日期精确命中，否则取该房间任一日期(兼容无日期列文件)。"""
    room = None
    for c in candidates or []:
        c = (c or "").strip()
        if not c:
            continue
        room = (equip_to_room or {}).get(c)
        if not room:
            for k, v in (equip_to_room or {}).items():
                if c in k or k in c:
                    room = v
                    break
        if room:
            break
    if not room or not room_env:
        return None, None
    if (room, date_str) in room_env:
        return room_env[(room, date_str)]
    for (rm, _ds), v in room_env.items():
        if rm == room:
            return v
    return None, None


def _write_env_records(path, entries, date_str):
    """把 entries={房间名称:(温度,湿度)} 按 date_str 写入 path 的「温湿度」sheet：
    (房间,日期)已存在则更新温度/湿度，否则追加。保留其它日期数据。失败抛异常。"""
    wb = openpyxl.load_workbook(path)
    if "温湿度" not in wb.sheetnames:
        ws = wb.create_sheet("温湿度")
        ws.append(["房间名称", "日期", "温度", "湿度"])
    else:
        ws = wb["温湿度"]
    idx = {}  # (房间,日期) -> 行号
    for r in range(2, ws.max_row + 1):
        room = ws.cell(r, 1).value
        if room is not None:
            idx[(str(room).strip(), _date_str(ws.cell(r, 2).value))] = r
    for room, (t, h) in (entries or {}).items():
        key = (room, date_str)
        if key in idx:
            r = idx[key]
            ws.cell(r, 3).value = t
            ws.cell(r, 4).value = h
        else:
            ws.append([room, date_str, t, h])
    wb.save(path)


def _equipment_env_candidates(row, equipment_config, matched_eq):
    """收集用于查房间的设备编号候选字符串(去重保序)：表格设备 + 匹配设备 + 方法默认主检设备。
    mainEquipmentNames 形如 '编号,名称'，取首段即设备编号。"""
    cands = []

    def add(*vals):
        for v in vals:
            s = (str(v or "")).strip().split(",")[0].strip()
            if s and s not in cands:
                cands.append(s)

    add(row.get("equipment"))
    if matched_eq:
        add(matched_eq.get("mainEquipmentNames"), matched_eq.get("name"),
            matched_eq.get("no") or matched_eq.get("code") or matched_eq.get("equipmentCode"))
    main_ids = set(str((equipment_config or {}).get("mainEquipmentIds", "")).split(","))
    for eq in (equipment_config or {}).get("raw_data") or []:
        if eq.get("usedCategory") != "检测设备":
            continue
        eid = str(eq.get("equipmentBillId") or eq.get("id") or "")
        if eid and eid in main_ids:
            add(eq.get("mainEquipmentNames"), eq.get("name"),
                eq.get("no") or eq.get("code") or eq.get("equipmentCode"))
    return cands


def _match_processing_rule(processing_rules, std_no):
    """按标准号(standardNo)匹配处理规则；支持精确→包含双向。无匹配返回 None(→直接读取兜底)。"""
    std = (std_no or "").strip()
    if std:
        for r in processing_rules or []:
            rm = str((r or {}).get("method") or "").strip()
            if rm and (rm == std or rm in std or std in rm):
                return r
    return None


def _apply_processing(raw, rule, wp):
    """按处理规则把原始称样量(float)转为提交字符串。
    type ∈ 直接读取/小数位补充/换算处理/换算加补充；
    换算两类型相同(=raw×factor 四舍五入到 rule.decimal_places)，小数位补充用 wp.decimal_places。"""
    rtype = (rule or {}).get("type") or "直接读取"
    if rtype in ("换算处理", "换算加补充"):
        try:
            factor = float((rule or {}).get("factor") or 1.0)
        except (TypeError, ValueError):
            factor = 1.0
        try:
            n = int((rule or {}).get("decimal_places") or 2)
        except (TypeError, ValueError):
            n = 2
        return f"{float(raw) * factor:.{n}f}"
    if rtype == "小数位补充":
        try:
            n = int((wp or {}).get("decimal_places") or 2)
        except (TypeError, ValueError):
            n = 2
        return f"{float(raw):.{n}f}"
    return _raw_mass_str(raw)  # 直接读取 / 未知


class UniversalCell:
    """通用单元格组件，支持单击选中和双击编辑"""

    def __init__(self, parent, width, row_index, col_index, initial_value,
                 selection_manager, on_data_update, on_drag_start,
                 on_drag_update, on_drag_end, on_cell_select, has_button=False, on_button_click=None,
                 placeholder=None, readonly=False, display_transform=None):
        self.parent = parent
        self.width = width
        self.row_index = row_index
        self.col_index = col_index
        self.selection_manager = selection_manager
        self.on_data_update = on_data_update
        self.on_drag_start = on_drag_start
        self.on_drag_update = on_drag_update
        self.on_drag_end = on_drag_end
        self.on_cell_select = on_cell_select
        self.has_button = has_button
        self.on_button_click = on_button_click
        self.placeholder = placeholder  # 占位提示文字（仅标准物质列用）
        self.readonly = readonly        # 只读态：禁止双击编辑（none 用）
        self.display_transform = display_transform  # 显示值转换（路径列：文件名-路径）

        # 编辑状态
        self.editing = False
        self.backup_value = ""
        self.ignore_focus_out = False
        self.click_timer = None  # 用于处理双击事件
        self.destroyed = False  # 标记部件是否已被销毁
        self.is_double_click = False  # 标记是否为双击事件

        self.create_widgets(initial_value)

    def create_widgets(self, initial_value):
        # 创建单元格框架
        self.cell_frame = tk.Frame(self.parent, width=self.width, height=40, bg="white")
        self.cell_frame.pack_propagate(False)

        # 添加右边框
        border_right = tk.Frame(self.cell_frame, width=1, bg="#c0c0c0")
        border_right.pack(side='right', fill='y')

        # 添加下边框
        border_bottom = tk.Frame(self.cell_frame, height=1, bg="#c0c0c0")
        border_bottom.pack(side='bottom', fill='x')

        # 创建按钮（如果有）
        if self.has_button and self.on_button_click:
            self.button = tk.Button(self.cell_frame, text="...", width=3, command=self.on_button_click)
            self.button.pack(side='right', padx=2, pady=1)

        # 创建标签（显示模式）
        self.value_var = tk.StringVar(value=initial_value)
        if self.placeholder or self.display_transform:
            # 手动渲染：占位文字 / 路径缩写显示，label 不绑 textvariable
            self.label = tk.Label(self.cell_frame, anchor="w", bg="white", fg="black")
            self._render_label()
            self.value_var.trace("w", lambda *a: self._render_label())
        else:
            self.label = tk.Label(self.cell_frame, textvariable=self.value_var, anchor="w", bg="white", fg="black")

        # 创建输入框（编辑模式）- 初始隐藏，移除所有边框
        self.entry = tk.Entry(self.cell_frame, textvariable=self.value_var, bg="white",
                              relief="flat", bd=0, highlightthickness=0)  # 移除所有边框

        # 绑定输入框事件
        self.entry.bind("<Return>", self.on_entry_return)
        self.entry.bind("<Escape>", self.on_entry_escape)
        self.entry.bind("<FocusOut>", self.on_entry_focus_out)
        self.entry.bind("<Button-1>", self.on_entry_click)

        # 默认显示标签
        self.label.pack(fill='both', expand=True, padx=2, pady=1)

        # 绑定鼠标事件到整个单元格框架
        self.cell_frame.bind("<Button-1>", self.on_click)
        self.cell_frame.bind("<B1-Motion>", self.on_drag)
        self.cell_frame.bind("<ButtonRelease-1>", self.on_release)
        self.cell_frame.bind("<Double-Button-1>", self.on_double_click)

        # 绑定鼠标事件到标签
        self.label.bind("<Button-1>", self.on_click)
        self.label.bind("<B1-Motion>", self.on_drag)
        self.label.bind("<ButtonRelease-1>", self.on_release)
        self.label.bind("<Double-Button-1>", self.on_double_click)

        # 绑定值变化事件
        self.value_var.trace("w", self.on_value_change)

    def on_entry_click(self, event):
        """输入框点击事件 - 阻止事件传播到单元格"""
        return "break"

    def on_value_change(self, *args):
        """值变化时更新数据"""
        if not self.destroyed:
            self.on_data_update()

    def get_value(self):
        """获取单元格值"""
        return self.value_var.get()

    def set_value(self, value):
        """设置单元格值"""
        self.value_var.set(value)

    def on_click(self, event):
        """单击事件 - 选择单元格"""
        if self.destroyed:
            return "break"

        # 按下时重置上一次(单击经延时定时器遗留的)拖拽状态，避免污染本次交互
        self.selection_manager.end_drag()

        # 取消之前的计时器（如果有）
        if self.click_timer:
            self.cell_frame.after_cancel(self.click_timer)
            self.click_timer = None

        # 设置计时器，延迟执行选择操作
        self.click_timer = self.cell_frame.after(200, self.perform_single_click)
        return "break"

    def perform_single_click(self):
        """执行单击操作"""
        if not self.destroyed:
            self.is_double_click = False
            self.on_cell_select(self.row_index, self.col_index, None)
        self.click_timer = None

    def on_double_click(self, event):
        """双击事件 - 进入编辑模式"""
        if self.destroyed:
            return "break"

        # 取消单击计时器
        if self.click_timer:
            self.cell_frame.after_cancel(self.click_timer)
            self.click_timer = None

        # 标记为双击事件
        self.is_double_click = True

        # 选择单元格
        self.on_cell_select(self.row_index, self.col_index, event)

        # 立即进入编辑模式
        self.start_editing()
        return "break"

    def on_drag(self, event):
        """拖动事件"""
        if self.destroyed:
            return "break"

        # 取消单击计时器(避免误触发单击选择)
        if self.click_timer:
            self.cell_frame.after_cancel(self.click_timer)
            self.click_timer = None

        # 首次移动时启动拖拽选择：序号列→选行范围，数据列→选单元格(限本列)
        if not self.selection_manager.dragging:
            drag_type = 'row' if self.col_index == 0 else 'cell'
            self.on_drag_start(self.row_index, self.col_index, drag_type)

        # 按鼠标当前位置更新选中范围(跨行)，而非按下那格的固定行号
        self.on_drag_update(event)
        return "break"

    def on_release(self, event):
        """释放事件"""
        if not self.destroyed:
            self.on_drag_end()
        return "break"

    def _render_label(self):
        """占位模式：按值/空渲染 label 文字与颜色"""
        if not getattr(self, "label", None) or not self.label.winfo_exists():
            return
        v = self.value_var.get()
        if v:
            text = self.display_transform(v) if self.display_transform else v
            self.label.configure(text=text, fg="#2c3e50")
        else:
            self.label.configure(text=self.placeholder or "", fg="#999")

    # 标液类型 → (placeholder, readonly)
    _STANDARD_MODES = {
        "fresh": ("现配现用", False),
        "fixed": ("配制序号", False),
        "none":  ("无需标液", True),
        "":      ("配制序号", False),
    }
    # 仪器设置 → (placeholder, readonly)
    _EQUIPMENT_MODES = {
        "default":   ("默认设备", True),
        "specified": ("设备编号", False),
        "":          ("设备编号", False),
    }

    def _set_placeholder(self, placeholder, readonly):
        """切换占位文字与只读态，并重渲染"""
        self.placeholder = placeholder
        self.readonly = readonly
        self._render_label()

    def set_standard_mode(self, mode):
        """按标液类型切换标准物质列（cells[3]）"""
        self._set_placeholder(*self._STANDARD_MODES.get(mode, ("配制序号", False)))

    def set_equipment_mode(self, mode):
        """按仪器设置切换设备列（cells[4]）"""
        self._set_placeholder(*self._EQUIPMENT_MODES.get(mode, ("设备编号", False)))

    # 称样量模式 → (placeholder, readonly)
    _WEIGHING_MODES = {
        "random":  ("随机称样", True),
        "none":    ("无需称样", True),
        "record":  ("称量记录", False),
        "process": ("过程称量", False),
        "":        ("称样记录", False),
    }

    def set_weighing_mode(self, mode):
        """按方法称样量模式切换称样记录列（cells[1]）：模式作为占位提示显示"""
        self._set_placeholder(*self._WEIGHING_MODES.get((mode or "").strip(), ("称样记录", False)))

    def start_editing(self):
        """开始编辑"""
        if self.editing or self.destroyed or self.readonly:
            return

        self.editing = True
        self.backup_value = self.value_var.get()
        self.ignore_focus_out = True

        # 隐藏标签，显示输入框
        self.label.pack_forget()

        # 检查部件是否存在
        if hasattr(self, 'entry') and self.entry.winfo_exists():
            self.entry.pack(fill='both', expand=True, padx=2, pady=1)  # 增加内边距
        else:
            self.editing = False
            return

        # 设置焦点并选中文本
        if hasattr(self, 'entry') and self.entry.winfo_exists():
            self.entry.focus_set()
            self.entry.icursor(tk.END)
            self.entry.select_range(0, tk.END)

            # 强制获取焦点并确保光标可见
            self.entry.focus_force()

        # 强制更新显示
        if hasattr(self, 'entry') and self.entry.winfo_exists():
            self.entry.update_idletasks()

        # 短暂延迟后重置标志
        if hasattr(self, 'entry') and self.entry.winfo_exists():
            self.entry.after(150, self.reset_ignore_focus_out)

    def reset_ignore_focus_out(self):
        """重置忽略焦点失去标志"""
        self.ignore_focus_out = False

    def stop_editing(self, confirm=True):
        """停止编辑"""
        if not self.editing or self.destroyed:
            return

        self.editing = False

        if not confirm:
            # 取消编辑，恢复原值
            self.value_var.set(self.backup_value)

        # 隐藏输入框，显示标签
        if hasattr(self, 'entry') and self.entry.winfo_exists():
            self.entry.pack_forget()

        if hasattr(self, 'label') and self.label.winfo_exists():
            self.label.pack(fill='both', expand=True, padx=2, pady=1)

        # 通知数据更新
        self.on_data_update()

    def on_entry_return(self, event):
        """输入框回车事件"""
        if not self.destroyed:
            self.stop_editing(confirm=True)
        return "break"

    def on_entry_escape(self, event):
        """输入框ESC事件"""
        if not self.destroyed:
            self.stop_editing(confirm=False)
        return "break"

    def on_entry_focus_out(self, event):
        """输入框失去焦点事件"""
        if self.ignore_focus_out or self.destroyed:
            return

        if self.editing:
            self.stop_editing(confirm=True)

    def update_display(self):
        """更新显示状态"""
        if self.destroyed:
            return

        is_selected = self.selection_manager.is_cell_selected(self.row_index, self.col_index)
        bg_color = "#e6f3ff" if is_selected else "white"

        if hasattr(self, 'cell_frame') and self.cell_frame.winfo_exists():
            self.cell_frame.configure(bg=bg_color)

        if hasattr(self, 'label') and self.label.winfo_exists():
            self.label.configure(bg=bg_color)

        if self.editing and hasattr(self, 'entry') and self.entry.winfo_exists():
            self.entry.configure(bg=bg_color)

        if hasattr(self, 'button') and self.button.winfo_exists():
            self.button.configure(bg=bg_color)

    def destroy(self):
        """销毁部件"""
        self.destroyed = True
        if self.click_timer:
            try:
                self.cell_frame.after_cancel(self.click_timer)
            except:
                pass
            self.click_timer = None
        if hasattr(self, 'cell_frame') and self.cell_frame.winfo_exists():
            try:
                self.cell_frame.destroy()
            except:
                pass


class SelectableRow:
    """可选择的行"""

    def __init__(self, parent, index, data, column_widths, selection_manager,
                 on_path_select, on_method_select, on_spectrum_select, on_cell_select,
                 on_drag_start, on_drag_update, on_drag_end, on_data_update):
        self.parent = parent
        self.index = index
        self.data = data
        self.column_widths = column_widths
        self.selection_manager = selection_manager
        self.on_path_select = on_path_select
        self.on_method_select = on_method_select
        self.on_spectrum_select = on_spectrum_select
        self.on_cell_select = on_cell_select
        self.on_drag_start = on_drag_start
        self.on_drag_update = on_drag_update
        self.on_drag_end = on_drag_end
        self.on_data_update = on_data_update

        # 所有单元格
        self.cells = []
        self.destroyed = False  # 标记部件是否已被销毁

        self.create_widgets()

    def create_widgets(self):
        # 创建行框架
        self.row_frame = tk.Frame(self.parent, bg="#c0c0c0")
        self.row_frame.pack(fill='x', pady=0)

        # 创建单元格
        self.create_cells()

        # 更新显示
        self.update_display()

    def create_cells(self):
        # 序号列 - 单击选择整行，双击编辑序号
        id_cell = UniversalCell(
            self.row_frame, self.column_widths[0], self.index, 0,
            str(self.data["id"]), self.selection_manager,
            self.on_id_data_update, self.on_drag_start,
            self.on_drag_update, self.on_drag_end,
            lambda row, col, event: self.on_row_select(row, event)
        )
        id_cell.cell_frame.pack(side='left', fill='y')
        self.cells.append(id_cell)

        # 称样记录路径 - 单击选择单元格，双击编辑内容
        weighing_cell = UniversalCell(
            self.row_frame, self.column_widths[1], self.index, 1,
            self.data["weighing_path"], self.selection_manager,
            self.on_weighing_data_update, self.on_drag_start,
            self.on_drag_update, self.on_drag_end,
            self.on_cell_select,
            has_button=True, on_button_click=lambda: self.on_weighing_button_click(), display_transform=_format_path
        )
        weighing_cell.cell_frame.pack(side='left', fill='y')
        weighing_cell.set_weighing_mode(self.data.get("weighing_mode", ""))
        self.cells.append(weighing_cell)

        # 录入方法文件 - 单击选择单元格，双击编辑内容
        method_cell = UniversalCell(
            self.row_frame, self.column_widths[2], self.index, 2,
            self.data["method_file"], self.selection_manager,
            self.on_method_data_update, self.on_drag_start,
            self.on_drag_update, self.on_drag_end,
            self.on_cell_select,
            has_button=True, on_button_click=lambda: self.on_method_button_click(), display_transform=_format_path
        )
        method_cell.cell_frame.pack(side='left', fill='y')
        self.cells.append(method_cell)

        # 标准物质 - 单击选择单元格，双击编辑内容（值绑定 configure_order，按标液类型切换占位/只读）
        reference_cell = UniversalCell(
            self.row_frame, self.column_widths[3], self.index, 3,
            self.data.get("configure_order", ""), self.selection_manager,
            self.on_reference_data_update, self.on_drag_start,
            self.on_drag_update, self.on_drag_end,
            self.on_cell_select,
            placeholder="配制序号"
        )
        reference_cell.cell_frame.pack(side='left', fill='y')
        reference_cell.set_standard_mode(self.data.get("standard_type", ""))
        self.cells.append(reference_cell)

        # 设备 - 单击选择单元格，双击编辑内容（按仪器设置切换占位/只读，值绑 equipment）
        equipment_cell = UniversalCell(
            self.row_frame, self.column_widths[4], self.index, 4,
            self.data.get("equipment", ""), self.selection_manager,
            self.on_equipment_data_update, self.on_drag_start,
            self.on_drag_update, self.on_drag_end,
            self.on_cell_select,
            placeholder="设备编号"
        )
        equipment_cell.cell_frame.pack(side='left', fill='y')
        equipment_cell.set_equipment_mode(self.data.get("instrument_setting", ""))
        self.cells.append(equipment_cell)

        # 谱图文件路径 - 单击选择单元格，双击编辑内容
        spectrum_cell = UniversalCell(
            self.row_frame, self.column_widths[5], self.index, 5,
            self.data["spectrum_path"], self.selection_manager,
            self.on_spectrum_data_update, self.on_drag_start,
            self.on_drag_update, self.on_drag_end,
            self.on_cell_select,
            has_button=True, on_button_click=lambda: self.on_spectrum_button_click(), display_transform=_format_path
        )
        spectrum_cell.cell_frame.pack(side='left', fill='y')
        self.cells.append(spectrum_cell)

        # 温度 - 单击选择，双击编辑（按设备房间自动填充，可手填覆盖）
        temp_cell = UniversalCell(
            self.row_frame, self.column_widths[6], self.index, 6,
            self.data.get("temperature", ""), self.selection_manager,
            self.on_temperature_data_update, self.on_drag_start,
            self.on_drag_update, self.on_drag_end,
            self.on_cell_select, placeholder="温度"
        )
        temp_cell.cell_frame.pack(side='left', fill='y')
        self.cells.append(temp_cell)

        # 湿度 - 单击选择，双击编辑
        humidity_cell = UniversalCell(
            self.row_frame, self.column_widths[7], self.index, 7,
            self.data.get("humidity", ""), self.selection_manager,
            self.on_humidity_data_update, self.on_drag_start,
            self.on_drag_update, self.on_drag_end,
            self.on_cell_select, placeholder="湿度"
        )
        humidity_cell.cell_frame.pack(side='left', fill='y')
        self.cells.append(humidity_cell)

        # 运行状态列 - 固定宽度只读，与 cells 同级，但不动 cells 列索引
        self.status_cell = tk.Frame(self.row_frame, width=self.column_widths[8], bg="white")
        self.status_cell.pack(side='left', fill='y')
        self.status_cell.pack_propagate(False)
        # 网格线（与 UniversalCell 一致）
        tk.Frame(self.status_cell, width=1, bg="#c0c0c0").pack(side='right', fill='y')
        tk.Frame(self.status_cell, height=1, bg="#c0c0c0").pack(side='bottom', fill='x')
        self.status_label = tk.Label(self.status_cell, text=self.data.get("status", "待运行"),
                                     anchor='center', relief='flat', padx=4,
                                     bg="white", fg="#888888",
                                     font=("Segoe UI", 9))
        self.status_label.pack(fill='both', expand=True)

    # 状态显示样式（阶段1）
    STATUS_STYLE = {
        "待运行": ("white", "#888888"),
        "运行中": ("#dbeafe", "#2563eb"),
        "成功":   ("#dcfce7", "#16a34a"),
        "失败":   ("#fee2e2", "#dc2626"),
        "跳过":   ("#f5f5f5", "#9ca3af"),
    }

    def set_status(self, status, error_msg=""):
        """更新本行状态显示（由主线程队列轮询调用）"""
        if self.destroyed:
            return
        self.data["status"] = status
        self.data["error_msg"] = error_msg
        bg, fg = self.STATUS_STYLE.get(status, ("#f0f0f0", "#888888"))
        text = status
        if status == "失败" and error_msg:
            short = error_msg if len(error_msg) <= 24 else error_msg[:24] + "…"
            text = f"失败:{short}"
        if self.status_label.winfo_exists():
            self.status_label.configure(text=text, bg=bg, fg=fg)

    def on_weighing_button_click(self):
        """称样记录路径按钮点击事件"""
        if self.selection_manager.is_cell_selected(self.index, 1):
            # 如果单元格被选中，进入编辑模式
            self.cells[1].start_editing()
        else:
            # 否则执行原来的路径选择功能
            self.on_path_select(self.index)

    def on_method_button_click(self):
        """方法文件按钮点击事件"""
        if self.selection_manager.is_cell_selected(self.index, 2):
            # 如果单元格被选中，进入编辑模式
            self.cells[2].start_editing()
        else:
            # 否则执行原来的路径选择功能
            self.on_method_select(self.index)

    def on_spectrum_button_click(self):
        """谱图文件路径按钮点击事件"""
        if self.selection_manager.is_cell_selected(self.index, 5):
            # 如果单元格被选中，进入编辑模式
            self.cells[5].start_editing()
        else:
            # 否则执行原来的路径选择功能
            self.on_spectrum_select(self.index)

    def on_row_select(self, row, event):
        """处理行选择（序号列）"""
        if not self.destroyed:
            self.on_drag_start(row, 0, 'row')

    def on_id_data_update(self):
        """序号数据更新"""
        if not self.destroyed:
            try:
                self.data["id"] = int(self.cells[0].get_value())
            except ValueError:
                self.cells[0].set_value(str(self.data["id"]))
            self.on_data_update()

    def on_weighing_data_update(self):
        """称样记录路径数据更新"""
        if not self.destroyed:
            self.data["weighing_path"] = self.cells[1].get_value()
            self.on_data_update()

    def on_method_data_update(self):
        """方法文件路径数据更新"""
        if not self.destroyed:
            self.data["method_file"] = self.cells[2].get_value()
            self.on_data_update()

    def on_reference_data_update(self):
        """标准物质数据更新"""
        if not self.destroyed:
            self.data["configure_order"] = self.cells[3].get_value()
            self.on_data_update()

    def on_equipment_data_update(self):
        """设备数据更新"""
        if not self.destroyed:
            self.data["equipment"] = self.cells[4].get_value()
            self.on_data_update()

    def on_spectrum_data_update(self):
        """谱图文件路径数据更新"""
        if not self.destroyed:
            self.data["spectrum_path"] = self.cells[5].get_value()
            self.on_data_update()

    def on_temperature_data_update(self):
        """温度数据更新"""
        if not self.destroyed:
            self.data["temperature"] = self.cells[6].get_value()
            self.on_data_update()

    def on_humidity_data_update(self):
        """湿度数据更新"""
        if not self.destroyed:
            self.data["humidity"] = self.cells[7].get_value()
            self.on_data_update()

    def update_display(self):
        """更新显示状态"""
        if not self.destroyed:
            for cell in self.cells:
                cell.update_display()

    def apply_column_widths(self, widths):
        """列宽被手动拖动后，重设本行各单元格宽度。
        cell_frame/status_cell 均设了 pack_propagate(False)，configure(width=) 即时生效。"""
        if self.destroyed:
            return
        for i, cell in enumerate(self.cells):
            cell.width = widths[i]
            cell.cell_frame.configure(width=widths[i])
        if self.status_cell.winfo_exists():
            self.status_cell.configure(width=widths[8])

    def destroy(self):
        """销毁行"""
        self.destroyed = True
        for cell in self.cells:
            if hasattr(cell, 'destroy'):
                cell.destroy()
        if hasattr(self, 'row_frame') and self.row_frame.winfo_exists():
            try:
                self.row_frame.destroy()
            except:
                pass

    def update_index(self, new_index, new_id):
        """删除其它行后就地更新本行的位置/序号(不重建控件，避免闪烁)。
        new_index: 0-based 位置(选择/拖拽用)；new_id: 1-based 显示序号"""
        if self.destroyed:
            return
        self.index = new_index
        self.data["id"] = new_id
        for cell in self.cells:
            cell.row_index = new_index
        # No 列(cells[0])刷新显示的新序号
        if self.cells and hasattr(self.cells[0], 'set_value'):
            self.cells[0].set_value(str(new_id))

    def sync_from_data(self):
        """就地按当前 data 重渲染本行所有单元格(不销毁/重建控件，避免闪烁)。
        覆盖：序号、各路径、标准物质/设备的值与占位模式、运行状态、选中态"""
        if self.destroyed or len(self.cells) < 8:
            return
        self.cells[0].set_value(str(self.data.get("id", "")))
        self.cells[1].set_value(self.data.get("weighing_path", ""))
        self.cells[1].set_weighing_mode(self.data.get("weighing_mode", ""))
        self.cells[2].set_value(self.data.get("method_file", ""))
        self.cells[3].set_value(self.data.get("configure_order", ""))
        self.cells[3].set_standard_mode(self.data.get("standard_type", ""))
        self.cells[4].set_value(self.data.get("equipment", ""))
        self.cells[4].set_equipment_mode(self.data.get("instrument_setting", ""))
        self.cells[5].set_value(self.data.get("spectrum_path", ""))
        self.cells[6].set_value(self.data.get("temperature", ""))
        self.cells[7].set_value(self.data.get("humidity", ""))
        self.set_status(self.data.get("status", "待运行"), self.data.get("error_msg", ""))
        self.update_display()


class SequenceMaster:
    def __init__(self, root):
        self.root = root
        self.root.title("SequenceMaster - 序列编辑器")
        self.root.geometry("1400x850")

        # 存储序列数据
        self.sequence_data = []

        # 选择管理器
        self.selection_manager = CellSelectionManager()

        # 键盘状态
        self.ctrl_pressed = False
        self.shift_pressed = False

        # 列宽配置
        self.column_widths = [50, 260, 220, 140, 140, 260, 70, 70, 140]

        # 存储分隔线引用
        self.draggable_headers = []

        # 存储行组件引用
        self.row_widgets = []

        # LIMS 登录与 API（阶段1）
        self.login_system = MultiUserLoginSystem()
        self.api = DetectionAPI(self.login_system)
        self.logged_in = False
        self._env_eq_cache = {}  # sample_code -> 默认设备编号(温湿度按房间匹配用，跨行去重)

        # 序列运行器状态（阶段1）：1 worker 线程 + queue + 4 Event，UI 更新全部 marshal 回主线程
        self._ui_q = queue.Queue()
        self._worker = None
        self._running = False
        self._stop = threading.Event()
        self._pause = threading.Event(); self._pause.set()  # set=运行中
        self._confirm_done = threading.Event()
        self._confirm_result = None
        self._confirm_win = None

        self.create_widgets()

        # 启动即复用已存 LIMS 会话(避免每个动作都要求重新登录)
        try:
            if self.login_system.load_session() and self.login_system.verify_session():
                self.logged_in = True
                self._refresh_user_menu()
        except Exception:
            pass

        # 启动 UI 队列轮询（主线程）
        self.root.after(100, self._drain_ui_queue)

        # 绑定键盘事件
        self.root.bind('<Control_L>', self.on_ctrl_press)
        self.root.bind('<Control_R>', self.on_ctrl_press)
        self.root.bind('<KeyRelease-Control_L>', self.on_ctrl_release)
        self.root.bind('<KeyRelease-Control_R>', self.on_ctrl_release)
        self.root.bind('<Shift_L>', self.on_shift_press)
        self.root.bind('<Shift_R>', self.on_shift_press)
        self.root.bind('<KeyRelease-Shift_L>', self.on_shift_release)
        self.root.bind('<KeyRelease-Shift_R>', self.on_shift_release)
        self.root.bind('<Escape>', self.clear_selection)

        # 绑定全局点击事件，用于退出编辑模式
        self.root.bind('<Button-1>', self.on_global_click)

        # 初始添加一行
        self.add_row()

    def create_widgets(self):
        # 档2: 收紧 ttk 按钮默认内边距(否则 ttkbootstrap 按钮偏高偏大)
        ttkb.Style().configure("TButton", padding=(8, 2))
        # 主框架（档2: 转 ttkb，主题提供底色）
        main_frame = ttkb.Frame(self.root)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # 按钮框架
        top_button_frame = ttkb.Frame(main_frame)
        top_button_frame.pack(fill='x', pady=(0, 0))

        # 操作按钮（bootstyle 配色，弃 button_style 的 bg/fg/relief）
        add_btn = ttkb.Button(top_button_frame, text="Add", command=self.add_row, width=4, bootstyle="secondary")
        add_btn.pack(side='left', padx=(0, 0))

        delete_btn = ttkb.Button(top_button_frame, text="✕", command=self.delete_selected_rows,
                                 width=3, bootstyle="secondary")
        delete_btn.pack(side='left', padx=(0, 0))

        fill_btn = ttkb.Button(top_button_frame, text="↓", command=self.fill_down,
                               width=3, bootstyle="secondary")
        fill_btn.pack(side='left', padx=(0, 0))

        clear_btn = ttkb.Button(top_button_frame, text="Clear", command=self.clear_all, width=5, bootstyle="secondary")
        clear_btn.pack(side='left')

        edit_method_btn = ttkb.Button(top_button_frame, text="📝 方法", command=self.edit_method, width=7, bootstyle="secondary")
        edit_method_btn.pack(side='left', padx=(8, 0))

        env_btn = ttkb.Button(top_button_frame, text="🌡 保存温湿度", command=self.save_env_to_excel, width=12, bootstyle="secondary")
        env_btn.pack(side='left', padx=(8, 0))

        # 用户区（右上，整合为一个菜单按钮）：按钮文案=登录状态，下拉=登录/切换用户 + 用户管理
        _c = ttkb.Style().colors  # 主题配色：按钮底色/下拉菜单配色都由此取，与界面保持一致
        self.user_menu_btn = tk.Menubutton(top_button_frame, text="👤 未登录 ▾",
                                           bg=_c.bg, fg="#dc2626", relief="flat",
                                           font=("Segoe UI", 9, "bold"), padx=10, pady=2,
                                           activebackground=_c.bg, cursor="hand2")
        self.user_menu = tk.Menu(self.user_menu_btn, tearoff=0,
                                 bg=_c.bg, fg=_c.fg,
                                 activebackground=_c.selectbg, activeforeground=_c.selectfg,
                                 relief="flat", borderwidth=0,
                                 font=("Segoe UI", 9))
        self.user_menu.add_command(label="登录", command=self._show_login_dialog)
        self.user_menu.add_command(label="用户管理", command=self._open_user_management)
        self.user_menu_btn.configure(menu=self.user_menu)
        self.user_menu_btn.pack(side='right', padx=(0, 8))

        # 创建表格容器
        self.create_table_container(main_frame)

        # 运行控制按钮（阶段1）
        run_ctrl = ttkb.Frame(main_frame)
        run_ctrl.pack(fill='x', pady=(5, 0))
        ttkb.Label(run_ctrl, text="运行控制:").pack(side='left', padx=(0, 5))
        self.pause_btn = ttkb.Button(run_ctrl, text="⏸ 暂停/继续", command=self.toggle_pause, width=10, state='disabled', bootstyle="secondary")
        self.pause_btn.pack(side='left', padx=2)
        self.skip_btn = ttkb.Button(run_ctrl, text="⏭ 跳过当前", command=self.skip_current, width=10, state='disabled', bootstyle="secondary")
        self.skip_btn.pack(side='left', padx=2)
        self.abort_btn = ttkb.Button(run_ctrl, text="⏹ 中止", command=self.abort_run, width=10, state='disabled', bootstyle="danger")
        self.abort_btn.pack(side='left', padx=2)

        # 运行日志区（阶段1）：标题可点击折叠/展开
        log_frame = ttkb.Labelframe(main_frame)
        log_frame.pack(fill='both', expand=False, pady=(5, 0))
        toggle_lbl = ttkb.Label(log_frame, text="▼ 运行日志", bootstyle="secondary", cursor="hand2")
        log_frame.configure(labelwidget=toggle_lbl)

        self.log_text = tk.Text(log_frame, height=7, wrap='word', state='disabled',
                                bg="#f7f7f7", fg="#3e3f3a", relief="flat", bd=0,
                                highlightthickness=0, padx=6, pady=4,
                                font=("Consolas", 9))
        log_scroll = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side='left', fill='both', expand=True)
        log_scroll.pack(side='right', fill='y')

        # 折叠/展开：点击标题切换 log_text 与滚动条的显隐
        self._log_expanded = True

        def _toggle_log(_e=None):
            if self._log_expanded:
                self.log_text.pack_forget()
                log_scroll.pack_forget()
                toggle_lbl.configure(text="▶ 运行日志")
                self._log_expanded = False
            else:
                self.log_text.pack(side='left', fill='both', expand=True)
                log_scroll.pack(side='right', fill='y')
                toggle_lbl.configure(text="▼ 运行日志")
                self._log_expanded = True

        toggle_lbl.bind('<Button-1>', _toggle_log)

        # 底部按钮框架
        bottom_frame = ttkb.Frame(main_frame)
        bottom_frame.pack(fill='x', pady=(8, 4))

        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttkb.Label(bottom_frame, textvariable=self.status_var, anchor='w')
        status_bar.pack(side='left', fill='x', expand=True, padx=5, pady=8)

        # 右侧按钮（加大宽度+垂直内边距，使底部操作栏更醒目）
        run_btn = ttkb.Button(bottom_frame, text="运行", command=self.run_sequence, width=6, padding=(6, 4), bootstyle="primary")
        run_btn.pack(side='right', padx=(6, 5), pady=4)
        # 从指定行开始运行(默认1=从头)
        start_box = ttkb.Frame(bottom_frame)
        start_box.pack(side='right', padx=(10, 0), pady=4)
        ttkb.Label(start_box, text="从第").pack(side='left')
        self.start_row_var = tk.StringVar(value="1")
        ttkb.Entry(start_box, textvariable=self.start_row_var, width=4).pack(side='left', padx=3)
        ttkb.Label(start_box, text="行起").pack(side='left')

        load_btn = ttkb.Button(bottom_frame, text="加载", command=self.load_sequence, width=6, padding=(6, 4), bootstyle="secondary")
        load_btn.pack(side='right', padx=(6, 0), pady=4)

        save_btn = ttkb.Button(bottom_frame, text="保存", command=self.save_sequence, width=6, padding=(6, 4), bootstyle="secondary")
        save_btn.pack(side='right', padx=(6, 0), pady=4)

    def create_table_container(self, parent):
        """创建表格容器，包含表头和表格"""
        table_frame = tk.Frame(parent, bg="white", relief="solid", bd=1,
                               highlightbackground="#c0c0c0", highlightcolor="#c0c0c0",
                               highlightthickness=1)
        table_frame.pack(fill='both', expand=True)

        # 创建表头
        self.create_table_header(table_frame)

        # 创建表格主体
        self.create_table_body(table_frame)

    def create_table_header(self, parent):
        """创建表格表头"""
        self.header_frame = tk.Frame(parent, height=40, bg="white")
        self.header_frame.pack(fill='x')
        self.header_frame.pack_propagate(False)

        headers = [
            {"text": "No", "anchor": "center"},
            {"text": "称样记录路径", "anchor": "w"},
            {"text": "录入方法", "anchor": "w"},
            {"text": "标准物质", "anchor": "w"},
            {"text": "设备", "anchor": "w"},
            {"text": "谱图文件路径", "anchor": "w"},
            {"text": "温度(℃)", "anchor": "center"},
            {"text": "湿度(%RH)", "anchor": "center"},
            {"text": "运行状态", "anchor": "center"}
        ]

        self.header_cells = []
        self.draggable_headers = []

        # 创建表头标签
        for i, header in enumerate(headers):
            cell_frame = tk.Frame(self.header_frame, height=40, bg="white")

            if i == 0:
                cell_frame.place(x=0, y=0, width=self.column_widths[i], height=40)
            else:
                prev_width = sum(self.column_widths[:i])
                cell_frame.place(x=prev_width, y=0, width=self.column_widths[i], height=40)

            # 添加边框
            border_right = tk.Frame(cell_frame, width=1, bg="#c0c0c0")
            border_right.pack(side='right', fill='y')

            border_bottom = tk.Frame(cell_frame, height=1, bg="#c0c0c0")
            border_bottom.pack(side='bottom', fill='x')

            label = tk.Label(cell_frame, text=header["text"], anchor=header["anchor"],
                             bg="white", fg="#2c3e50",
                             font=("Segoe UI", 9))
            label.pack(fill='both', expand=True, padx=4, pady=3)

            self.header_cells.append(cell_frame)

        # 创建可拖动的分隔线
        for i in range(len(headers)):
            separator = DraggableHeader(self.header_frame, i, self.on_column_drag)
            self.draggable_headers.append(separator)

            separator_x = sum(self.column_widths[:i + 1])
            separator.update_position(separator_x)
            separator.separator.lift()

    def on_column_drag(self, column_index, delta):
        """处理列宽调整"""
        if column_index < len(self.column_widths) - 1:
            self.column_widths[column_index] += delta
            self.column_widths[column_index + 1] -= delta
        else:
            self.column_widths[column_index] += delta

        # 确保最小宽度
        min_width = 50
        for i in range(len(self.column_widths)):
            if self.column_widths[i] < min_width:
                self.column_widths[i] = min_width

        # 更新表头和分隔线位置
        self.update_header_layout()

        # 更新内容行
        self.refresh_table()

    def update_header_layout(self):
        """更新表头布局"""
        for i, cell in enumerate(self.header_cells):
            if i == 0:
                cell.place_configure(x=0, width=self.column_widths[i])
            else:
                prev_width = sum(self.column_widths[:i])
                cell.place_configure(x=prev_width, width=self.column_widths[i])

        for i, separator in enumerate(self.draggable_headers):
            separator_x = sum(self.column_widths[:i + 1])
            separator.update_position(separator_x)

    def create_table_body(self, parent):
        """创建表格主体"""
        table_container = tk.Frame(parent, bg="white")
        table_container.pack(fill='both', expand=True)

        # 创建Canvas和滚动条
        self.canvas = tk.Canvas(table_container, bg='white', highlightthickness=0, height=250)
        self.scrollbar = tk.Scrollbar(table_container, orient="vertical", command=self.canvas.yview)

        # 滚动区域框架
        self.scrollable_frame = tk.Frame(self.canvas, bg="white")

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # 绑定鼠标滚轮事件
        self.canvas.bind("<MouseWheel>", self.on_mousewheel)
        self.scrollable_frame.bind("<MouseWheel>", self.on_mousewheel)

    def on_mousewheel(self, event):
        """处理鼠标滚轮事件：仅当内容超出可视区时才滚动，
        否则行数少时上滚会把第一行往下拽（Tk 对内容小于画布的 yview 不夹紧）"""
        bbox = self.canvas.bbox("all")
        if not bbox:
            return
        content_h = bbox[3] - bbox[1]
        if content_h > self.canvas.winfo_height():
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def add_row(self):
        """添加新行(增量追加，不重建整表，避免闪烁)"""
        row_id = len(self.sequence_data) + 1
        new_row = {
            "id": row_id,
            "weighing_path": "",
            "method_file": "",
            "reference_material": "",
            "equipment": "",
            "spectrum_path": "",
            "sample_code": "",
            "configure_order": "",
            "standard_type": "",
            "instrument_setting": "",
            "weighing_mode": "",
            "temperature": "",
            "humidity": "",
            "status": "待运行",
            "error_msg": "",
            "experiment_code": ""
        }
        self.sequence_data.append(new_row)
        idx = len(self.sequence_data) - 1
        row_widget = SelectableRow(
            self.scrollable_frame,
            idx,
            new_row,
            self.column_widths,
            self.selection_manager,
            self.select_weighing_path,
            self.select_method_file,
            self.select_spectrum_path,
            self.handle_cell_select,
            self.handle_drag_start,
            self.handle_drag_update,
            self.handle_drag_end,
            self.on_data_update
        )
        self.row_widgets.append(row_widget)
        self.status_var.set(f"已添加第 {row_id} 行")

    def refresh_table(self):
        """刷新表格显示(就地更新，避免销毁重建导致的闪烁)。
        行数不变时只重渲染各单元格；行数变化时按差额增删再同步。"""
        n_data = len(self.sequence_data)

        # 多余的行：从末尾销毁
        while len(self.row_widgets) > n_data:
            w = self.row_widgets.pop()
            if hasattr(w, 'destroy'):
                w.destroy()

        # 不足的行：新建(旧行不动，只追加)
        while len(self.row_widgets) < n_data:
            idx = len(self.row_widgets)
            row_widget = SelectableRow(
                self.scrollable_frame,
                idx,
                self.sequence_data[idx],
                self.column_widths,
                self.selection_manager,
                self.select_weighing_path,
                self.select_method_file,
                self.select_spectrum_path,
                self.handle_cell_select,
                self.handle_drag_start,
                self.handle_drag_update,
                self.handle_drag_end,
                self.on_data_update
            )
            self.row_widgets.append(row_widget)

        # 全部行：同步数据/序号/显示(不销毁控件)
        for i, w in enumerate(self.row_widgets):
            w.index = i
            w.data = self.sequence_data[i]
            w.apply_column_widths(self.column_widths)
            w.sync_from_data()

    def on_data_update(self):
        """数据更新时的回调函数"""
        pass

    def on_global_click(self, event):
        """全局点击事件处理，用于退出编辑模式"""
        widget = event.widget

        # 检查是否点击了任何可编辑单元格
        for row_widget in self.row_widgets:
            for cell in row_widget.cells:
                if cell and hasattr(cell, 'entry') and widget == cell.entry:
                    return
                if cell and hasattr(cell, 'cell_frame') and widget == cell.cell_frame:
                    return
                if cell and hasattr(cell, 'label') and widget == cell.label:
                    return

        # 如果不是任何可编辑单元格，结束所有编辑状态
        for row_widget in self.row_widgets:
            for cell in row_widget.cells:
                if cell and hasattr(cell, 'editing') and cell.editing:
                    cell.stop_editing(confirm=True)

    # 键盘事件处理
    def on_ctrl_press(self, event):
        self.ctrl_pressed = True
        return "break"

    def on_ctrl_release(self, event):
        self.ctrl_pressed = False
        return "break"

    def on_shift_press(self, event):
        self.shift_pressed = True
        return "break"

    def on_shift_release(self, event):
        self.shift_pressed = False
        return "break"

    def clear_selection(self, event=None):
        """清除所有选择"""
        self.selection_manager.clear_all_selection()
        self.refresh_table()
        self.status_var.set("已清除选择")
        return "break"

    def handle_cell_select(self, row, col, event):
        """处理单元格选择"""
        # 检查是否是双击事件
        is_double_click = False
        if event and hasattr(event, 'widget'):
            # 通过检查事件来源的单元格来确定是否是双击
            for row_widget in self.row_widgets:
                if row_widget.index == row:
                    for cell in row_widget.cells:
                        if cell.col_index == col and hasattr(cell, 'is_double_click') and cell.is_double_click:
                            is_double_click = True
                            cell.is_double_click = False  # 重置标志
                            break
                    break

        if is_double_click:
            # 对于双击事件，我们直接选择单元格但不刷新表格
            self.selection_manager.clear_all_selection()
            self.selection_manager.select_cell(row, col)
            self.last_selected_cell = (row, col)

            # 只更新当前行的显示，不刷新整个表格
            if 0 <= row < len(self.row_widgets):
                self.row_widgets[row].update_display()

            self.update_status()
        else:
            # 对于单击和其他事件，使用原来的逻辑
            if self.ctrl_pressed:
                # Ctrl+点击：切换单元格选择状态
                if self.selection_manager.is_cell_selected(row, col):
                    self.selection_manager.deselect_cell(row, col)
                else:
                    self.selection_manager.select_cell(row, col)
            elif self.shift_pressed and hasattr(self, 'last_selected_cell') and self.last_selected_cell is not None:
                # Shift+点击：选择连续的单元格区域
                last_row, last_col = self.last_selected_cell
                start_row, end_row = sorted([last_row, row])
                start_col, end_col = sorted([last_col, col])

                for r in range(start_row, end_row + 1):
                    for c in range(start_col, end_col + 1):
                        self.selection_manager.select_cell(r, c)
            else:
                # 普通点击：只选择当前单元格
                self.selection_manager.clear_all_selection()
                self.selection_manager.select_cell(row, col)
                self.last_selected_cell = (row, col)

            self.refresh_table()
            self.update_status()

    def handle_drag_start(self, row, col, drag_type):
        """处理选择开始(点击序号列选行 / 点击数据格选单元格；亦为拖拽起点)"""
        if drag_type == 'row':
            anchor = getattr(self, 'last_selected_row', None)
            if self.shift_pressed and anchor is not None:
                # Shift：从锚点行到当前行连续选择(不清除、不改锚点)
                for r in range(min(anchor, row), max(anchor, row) + 1):
                    self.selection_manager.select_row(r)
            elif self.ctrl_pressed:
                # Ctrl：切换单行选中
                if self.selection_manager.is_row_selected(row):
                    self.selection_manager.deselect_row(row)
                else:
                    self.selection_manager.select_row(row)
            else:
                # 普通点击：只选当前行，并记为锚点
                self.selection_manager.clear_all_selection()
                self.selection_manager.select_row(row)
                self.last_selected_row = row
            self.selection_manager.start_drag(row, col, drag_type)
        else:
            if not self.ctrl_pressed and not self.shift_pressed:
                self.selection_manager.clear_all_selection()
            self.selection_manager.start_drag(row, col, drag_type)
            self.last_selected_cell = (row, col)
            self.selection_manager.select_cell(row, col)

        self.refresh_table()

    def handle_drag_update(self, event):
        """拖拽中：按鼠标当前所在行更新选中范围(跨行/跨单元格)"""
        row = self._row_at(event)
        if row is None:
            return
        self.selection_manager.update_drag(row, 0, self.ctrl_pressed, self.shift_pressed)
        self.refresh_table()

    def _row_at(self, event):
        """根据事件屏幕坐标定位鼠标当前所在的行索引(用于跨行拖拽)；不在任何行上返回 None"""
        try:
            w = self.root.winfo_containing(event.x_root, event.y_root)
        except Exception:
            return None
        seen = set()
        while w is not None and id(w) not in seen:
            seen.add(id(w))
            for i, rw in enumerate(self.row_widgets):
                if rw.row_frame is w:
                    return i
            w = w.master
        return None

    def handle_drag_end(self):
        """处理拖动结束"""
        self.selection_manager.end_drag()
        self.update_status()

    def update_status(self):
        """更新状态栏"""
        selected_rows = len(self.selection_manager.selected_rows)
        selected_cells = len(self.selection_manager.selected_cells)

        if selected_rows > 0:
            self.status_var.set(f"已选择 {selected_rows} 行")
        elif selected_cells > 0:
            self.status_var.set(f"已选择 {selected_cells} 个单元格")
        else:
            self.status_var.set("就绪")

    def select_weighing_path(self, row_index):
        """选择称样记录文件"""
        file_path = filedialog.askopenfilename(
            title="选择称样记录文件",
            filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")]
        )
        if file_path and 0 <= row_index < len(self.sequence_data):
            self.sequence_data[row_index]["weighing_path"] = file_path
            self.refresh_table()
            self.status_var.set(f"第 {row_index + 1} 行称样记录路径已设置")

    def select_method_file(self, row_index):
        """选择录入方法文件"""
        file_path = filedialog.askopenfilename(
            title="选择录入方法文件",
            filetypes=[("YAML files", "*.yaml"), ("配置文件", "*.yaml;*.yml;*.json"), ("All files", "*.*")]
        )
        if file_path and 0 <= row_index < len(self.sequence_data):
            row = self.sequence_data[row_index]
            row["method_file"] = file_path
            # 按方法回填标液+设备配置
            self._apply_method_params(row, file_path)
            self.refresh_table()
            self.status_var.set(f"第 {row_index + 1} 行录入方法文件已设置")

    def _read_other_params(self, method_file):
        """读取方法 yaml 的 other_params_settings，返回 dict"""
        if not method_file or not os.path.isfile(method_file):
            return {}
        try:
            with open(method_file, "r", encoding="utf-8") as f:
                y = yaml.safe_load(f) or {}
            return y.get("other_params_settings") or {}
        except Exception:
            return {}

    def _read_weighing_params(self, method_file):
        """读取方法 yaml 的 weighing_params（称样量模式与随机参数），返回 dict"""
        if not method_file or not os.path.isfile(method_file):
            return {}
        try:
            with open(method_file, "r", encoding="utf-8") as f:
                y = yaml.safe_load(f) or {}
            return y.get("weighing_params") or {}
        except Exception:
            return {}

    def _read_processing_rules(self, method_file):
        """读取方法 yaml 顶层 processing_rules（称量记录处理规则），返回 list"""
        if not method_file or not os.path.isfile(method_file):
            return []
        try:
            with open(method_file, "r", encoding="utf-8") as f:
                y = yaml.safe_load(f) or {}
            return y.get("processing_rules") or []
        except Exception:
            return []

    def _apply_standard_config(self, row, st, prep):
        """标液配置应用到行"""
        row["standard_type"] = st
        if st == "fixed" and prep:
            row["configure_order"] = prep      # 带出固定编号，可覆盖
        elif st == "none":
            row["configure_order"] = ""        # 无需标液，清空
        # fresh / 未知：保留用户已填或空

    def _apply_equipment_config(self, row, setting, device):
        """设备配置应用到行（仪器设置→显示模式；指定设备带出编号可覆盖）"""
        row["instrument_setting"] = setting
        if setting == "specified" and device:
            row["equipment"] = device          # 带出设备编号，可覆盖
        elif setting == "default":
            row["equipment"] = ""              # 默认设备，清空（显示占位）
        # 未配置：保留用户已填或空

    def _apply_method_params(self, row, method_file):
        """选方法 / 编辑器保存后：一次读取并回填标液+设备+称样量模式配置（共用）"""
        ops = self._read_other_params(method_file)
        self._apply_standard_config(row, ops.get("standard_type", ""), ops.get("preparation_number", ""))
        self._apply_equipment_config(row, ops.get("instrument_setting", ""), ops.get("device_number", ""))
        wp = self._read_weighing_params(method_file)
        row["weighing_mode"] = (wp.get("weighing_mode") or "").strip() if wp else ""
        self._autofill_env(row)

    def _autofill_env(self, row):
        """按 设备→房间→(房间,今天) 自动填温湿度(空列才填，已填保留)。
        指定设备直接匹配；默认设备(无编号)在已登录时查 LIMS 默认主检设备编号再匹配。
        每个未填充的原因都写日志，便于排查。"""
        if (row.get("temperature") or "").strip() and (row.get("humidity") or "").strip():
            return
        rid = row.get("id")
        equip_to_room, room_env, err = _read_env_records(_ENV_RECORD_PATH)
        if equip_to_room is None:
            self._log(f"行{rid}: 温湿度未自动填充(读不到温湿度记录文件: {err})")
            return
        eq = (row.get("equipment") or "").strip()
        cands = [eq] if eq else []
        src = "指定设备"
        if not cands:
            if not self._ensure_session_silent():
                self._log(f"行{rid}: 默认设备温湿度未填充(未登录且无可复用会话；登录后重选方法/谱图即可)")
                return
            code = self._resolve_default_equipment_code(row, self._log, self._env_eq_cache)
            if not code:
                self._log(f"行{rid}: 默认设备温湿度未填充(查不到默认设备；需该行先有谱图路径/样品编号且样品可查)")
                return
            cands = [code]
            src = "默认设备"
        today = datetime.now().strftime("%Y-%m-%d")
        temp, hum = _resolve_env(equip_to_room, room_env, cands, today)
        room = next((equip_to_room.get(c) or
                     next((v for k, v in equip_to_room.items() if c in k or k in c), None)
                     for c in cands), None)
        if temp and hum:
            row["temperature"], row["humidity"] = temp, hum
            self._log(f"行{rid}: 温湿度已自动填充({src}→{room} 当天) T={temp}℃ H={hum}%RH")
        else:
            self._log(f"行{rid}: 温湿度未填充({src}→房间「{room or '未匹配'}」在温湿度记录中无当天({today})数据)")

    def _ensure_session_silent(self):
        """确保有可用 LIMS 会话(已登录或可复用存档)；无会话返回 False，不弹登录框。"""
        if self.logged_in:
            return True
        try:
            if self.login_system.load_session() and self.login_system.verify_session():
                self.logged_in = True
                self._refresh_user_menu()
                return True
        except Exception:
            pass
        return False

    def save_env_to_excel(self):
        """把各行录入的温湿度按 设备编号→房间 聚合，回写当天值到 温湿度记录.xlsx 的「温湿度」sheet。
        指定设备直接匹配；默认设备(无表格编号)先查 LIMS 取默认主检设备编号再匹配。
        (房间,当天)已存在则更新，否则追加；保留其它日期。未录温湿度或设备无法匹配房间的行跳过。"""
        if not self.sequence_data:
            self.status_var.set("提示: 没有可保存的序列数据")
            return
        equip_to_room, _, err = _read_env_records(_ENV_RECORD_PATH)
        if equip_to_room is None:
            messagebox.showerror("错误", err or "无法读取温湿度记录文件")
            return
        today = datetime.now().strftime("%Y-%m-%d")
        # 需查默认设备的行(已录温湿度但无表格设备编号)
        default_rows = [r for r in self.sequence_data
                        if (r.get("temperature") or "").strip() and (r.get("humidity") or "").strip()
                        and not (r.get("equipment") or "").strip()]
        can_query = bool(default_rows) and self._ensure_session_silent()
        if default_rows:
            if can_query:
                self._log(f"查询 {len(default_rows)} 行的默认设备(用于温湿度按房间匹配)...")
            else:
                self._log("未登录且无可复用会话：默认设备的行跳过设备查询(仅保存指定设备行)")

        entries, skipped = {}, 0
        for row in self.sequence_data:
            t = (row.get("temperature") or "").strip()
            h = (row.get("humidity") or "").strip()
            if not t or not h:
                continue
            eq = (row.get("equipment") or "").strip()
            if not eq and can_query:
                eq = self._resolve_default_equipment_code(row, self._log, self._env_eq_cache) or ""
                if eq:
                    self._log(f"行{row.get('id')}: 默认设备编号={eq}")
                self.root.update_idletasks()  # 同步查 LIMS 期间刷新日志，避免界面假死
            room = None
            if eq:
                room = equip_to_room.get(eq)
                if not room:
                    for k, v in equip_to_room.items():
                        if eq in k or k in eq:
                            room = v
                            break
            if not room:
                skipped += 1
                continue
            entries[room] = (t, h)  # 同房间多行：后者覆盖(以最后一次录入为准)
        if not entries:
            self.status_var.set("提示: 没有可保存的温湿度(需录入温度+湿度，且设备能匹配到房间)")
            return
        try:
            _write_env_records(_ENV_RECORD_PATH, entries, today)
        except PermissionError:
            messagebox.showerror("错误", "温湿度记录.xlsx 被占用(可能正用 Excel 打开)，请关闭后重试")
            return
        except Exception as e:
            messagebox.showerror("错误", f"写入温湿度记录失败: {e}")
            return
        msg = f"已保存 {len(entries)} 个房间的温湿度({today})到 温湿度记录.xlsx"
        if skipped:
            msg += f"\n跳过 {skipped} 行(未录温湿度或设备无法匹配房间)"
        self._log(msg)
        self.status_var.set(msg)

    def _resolve_default_equipment_code(self, row, log, cache):
        """查询 LIMS 取默认主检设备编号(供温湿度按房间匹配)：样品→方法→全量配置→默认主检设备编号。
        复用运行期同一套查询(get_all_configs→equipment_config)；cache 按 sample_code 去重。
        失败/无法确定返回 None。"""
        _, sample_code = self._resolve_spectrum_pdf(row, log)
        if not sample_code:
            return None
        if sample_code in cache:
            return cache[sample_code]
        code = None
        try:
            projects = self.api.query_samples_by_conditions(
                sample_code=sample_code, exact_match=True, log_func=log)
            if projects:
                projects, ferr = self._filter_projects_by_method(projects, row, log)
                if projects and not ferr:
                    sample_id = projects[0].get("sampleId")
                    sp_ids = ",".join(str(p["projectId"]) for p in projects if p.get("projectId"))
                    method_name = projects[0].get("standardNo") or ""
                    initial = self.api.get_experiment_config(sp_ids, method_name, "", sample_id, log)
                    method_id = (initial.get("ocMethodSettings", {}) or {}).get("methodId")
                    pnames = [p.get("projectName", "") for p in projects]
                    all_cfg = self.api.get_all_configs(sp_ids, method_name, "", sample_id, log, method_id, pnames)
                    eq_cfg = (all_cfg or {}).get("equipment") or {}
                    cands = _equipment_env_candidates(row, eq_cfg, None)
                    code = cands[0] if cands else None
        except Exception as e:
            log(f"查询默认设备异常({sample_code}): {e}")
        cache[sample_code] = code
        return code

    def _override_equipment(self, equipment_config, device_number):
        """表格指定的设备编号覆盖默认主检设备（设备以序列表格为准）。
        返回 (equipment_config, matched_eq, error_msg)。表格空则用方法默认。
        matched_eq 为匹配到的原始设备条目，供 save_main_equipment 构造 items。"""
        device_number = (device_number or "").strip()
        if not device_number or not equipment_config:
            return equipment_config, None, None
        raw = equipment_config.get("raw_data") or []
        for eq in raw:
            if eq.get('usedCategory') != '检测设备':
                continue
            name = (eq.get("name") or "").strip()
            men = (eq.get("mainEquipmentNames") or "").strip()
            men_first = men.split(',')[0].strip() if men else ""
            if device_number == name or device_number == men or device_number == men_first:
                eid = eq.get("equipmentBillId") or eq.get("id")
                display = men or name
                cfg = dict(equipment_config)
                cfg["mainEquipment"] = display
                cfg["mainEquipmentNames"] = men or name
                cfg["mainEquipmentIds"] = str(eid) if eid else ""
                return cfg, eq, None
        return equipment_config, None, f"设备编号 {device_number} 未在方法检测设备中找到"

    def select_spectrum_path(self, row_index):
        """选择谱图文件路径"""
        path = filedialog.askdirectory(title="选择谱图文件夹路径")
        if path and 0 <= row_index < len(self.sequence_data):
            row = self.sequence_data[row_index]
            row["spectrum_path"] = path
            self._autofill_env(row)  # 谱图定了样品可查，默认设备温湿度此时再尝试自动填充(与选方法互不依赖先后)
            self.refresh_table()
            self.status_var.set(f"第 {row_index + 1} 行谱图文件路径已设置")

    def delete_selected_rows(self):
        """删除选中的行"""
        selected_rows = set()
        for row in self.selection_manager.selected_rows:
            selected_rows.add(row)
        for row, col in self.selection_manager.selected_cells:
            selected_rows.add(row)

        if not selected_rows:
            self.status_var.set("提示: 没有选中的行")
            return

        if not messagebox.askyesno("确认删除", f"确定要删除选中的 {len(selected_rows)} 行吗？"):
            return

        # 从后往前删：数据 + 对应行控件一起删(逆序保证低位索引仍有效)
        for index in sorted(selected_rows, reverse=True):
            if 0 <= index < len(self.sequence_data):
                del self.sequence_data[index]
            if 0 <= index < len(self.row_widgets):
                w = self.row_widgets[index]
                if hasattr(w, 'destroy'):
                    w.destroy()
                del self.row_widgets[index]

        self.selection_manager.clear_all_selection()
        self.renumber_rows()
        # 剩余行就地更新序号(不重建控件，避免闪烁)
        for i, w in enumerate(self.row_widgets):
            if hasattr(w, 'update_index'):
                w.update_index(i, i + 1)
        self.status_var.set(f"已删除 {len(selected_rows)} 行")

    def renumber_rows(self):
        """重新编号所有行"""
        for i, row_data in enumerate(self.sequence_data, 1):
            row_data["id"] = i

    def fill_down(self):
        """向下填充选中的内容"""
        if self.selection_manager.selected_rows:
            if len(self.selection_manager.selected_rows) > 1:
                self.status_var.set("提示: 只能选择一行作为填充源")
                return

            source_row = next(iter(self.selection_manager.selected_rows))

            if source_row >= len(self.sequence_data) - 1:
                self.status_var.set("提示: 已是最后一行，无需向下填充")
                return

            source_data = self.sequence_data[source_row]

            for i in range(source_row + 1, len(self.sequence_data)):
                self.sequence_data[i]["weighing_path"] = source_data["weighing_path"]
                self.sequence_data[i]["method_file"] = source_data["method_file"]
                self.sequence_data[i]["configure_order"] = source_data.get("configure_order", "")
                self.sequence_data[i]["standard_type"] = source_data.get("standard_type", "")
                self.sequence_data[i]["equipment"] = source_data["equipment"]
                self.sequence_data[i]["instrument_setting"] = source_data.get("instrument_setting", "")
                self.sequence_data[i]["spectrum_path"] = source_data["spectrum_path"]
                self.sequence_data[i]["temperature"] = source_data.get("temperature", "")
                self.sequence_data[i]["humidity"] = source_data.get("humidity", "")

            self.refresh_table()
            self.status_var.set(f"已从第 {source_row + 1} 行向下填充到第 {len(self.sequence_data)} 行")

        elif self.selection_manager.selected_cells:
            selected_cells = list(self.selection_manager.selected_cells)

            if not selected_cells:
                self.status_var.set("提示: 请先选择要填充的源单元格")
                return

            cells_by_column = {}
            for row, col in selected_cells:
                if col not in cells_by_column:
                    cells_by_column[col] = []
                cells_by_column[col].append(row)

            for col, rows in cells_by_column.items():
                if len(rows) != 1:
                    self.status_var.set(
                        f"提示: 第 {['序号', '称样记录路径', '录入方法', '标准物质', '设备', '谱图文件路径', '温度', '湿度'][col]} 列只能选择一个源单元格")
                    continue

                source_row = rows[0]

                if source_row >= len(self.sequence_data) - 1:
                    self.status_var.set("提示: 已是最后一行，无需向下填充")
                    continue

                if col == 0:
                    source_data = self.sequence_data[source_row]["id"]
                    for i in range(source_row + 1, len(self.sequence_data)):
                        self.sequence_data[i]["id"] = source_data
                elif col == 1:
                    source_data = self.sequence_data[source_row]["weighing_path"]
                    for i in range(source_row + 1, len(self.sequence_data)):
                        self.sequence_data[i]["weighing_path"] = source_data
                elif col == 2:
                    source_data = self.sequence_data[source_row]["method_file"]
                    for i in range(source_row + 1, len(self.sequence_data)):
                        self.sequence_data[i]["method_file"] = source_data
                elif col == 3:
                    source_data = self.sequence_data[source_row].get("configure_order", "")
                    for i in range(source_row + 1, len(self.sequence_data)):
                        self.sequence_data[i]["configure_order"] = source_data
                elif col == 4:
                    source_data = self.sequence_data[source_row]["equipment"]
                    for i in range(source_row + 1, len(self.sequence_data)):
                        self.sequence_data[i]["equipment"] = source_data
                elif col == 5:
                    source_data = self.sequence_data[source_row]["spectrum_path"]
                    for i in range(source_row + 1, len(self.sequence_data)):
                        self.sequence_data[i]["spectrum_path"] = source_data
                elif col == 6:
                    source_data = self.sequence_data[source_row].get("temperature", "")
                    for i in range(source_row + 1, len(self.sequence_data)):
                        self.sequence_data[i]["temperature"] = source_data
                elif col == 7:
                    source_data = self.sequence_data[source_row].get("humidity", "")
                    for i in range(source_row + 1, len(self.sequence_data)):
                        self.sequence_data[i]["humidity"] = source_data

            self.refresh_table()
            self.status_var.set("向下填充完成")
        else:
            self.status_var.set("提示: 请先选择要填充的行或单元格")

    def clear_all(self):
        """清空所有行内容"""
        if not self.sequence_data:
            self.status_var.set("提示: 没有数据可清空")
            return

        if not messagebox.askyesno("确认清空", "确定要清空所有行的内容吗？"):
            return

        for row_data in self.sequence_data:
            row_data["weighing_path"] = ""
            row_data["method_file"] = ""
            row_data["reference_material"] = ""
            row_data["configure_order"] = ""
            row_data["standard_type"] = ""
            row_data["equipment"] = ""
            row_data["instrument_setting"] = ""
            row_data["spectrum_path"] = ""
            row_data["temperature"] = ""
            row_data["humidity"] = ""

        self.selection_manager.clear_all_selection()
        self.refresh_table()
        self.status_var.set("已清空所有行内容")

    def save_sequence(self):
        """保存序列到文件"""
        if not self.sequence_data:
            self.status_var.set("提示: 没有数据可保存")
            return

        file_path = filedialog.asksaveasfilename(
            title="保存序列文件",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("Text files", "*.txt"), ("All files", "*.*")]
        )

        if not file_path:
            return

        try:
            save_data = {
                "sequence_data": [],
                "column_widths": self.column_widths
            }

            for row in self.sequence_data:
                save_data["sequence_data"].append({
                    "id": row["id"],
                    "weighing_path": row["weighing_path"],
                    "method_file": row["method_file"],
                    "reference_material": row["reference_material"],
                    "equipment": row["equipment"],
                    "spectrum_path": row["spectrum_path"],
                    "sample_code": row.get("sample_code", ""),
                    "configure_order": row.get("configure_order", ""),
                    "standard_type": row.get("standard_type", ""),
                    "instrument_setting": row.get("instrument_setting", ""),
                    "weighing_mode": row.get("weighing_mode", ""),
                    "temperature": row.get("temperature", ""),
                    "humidity": row.get("humidity", ""),
                    "status": row.get("status", "待运行"),
                    "error_msg": row.get("error_msg", ""),
                    "experiment_code": row.get("experiment_code", "")
                })

            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(save_data, f, ensure_ascii=False, indent=2)

            self.status_var.set(f"序列已保存到: {file_path}")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {str(e)}")

    def load_sequence(self):
        """从文件加载序列"""
        file_path = filedialog.askopenfilename(
            title="加载序列文件",
            filetypes=[("JSON files", "*.json"), ("Text files", "*.txt"), ("All files", "*.*")]
        )

        if not file_path:
            return

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                loaded_data = json.load(f)

            self.sequence_data = []
            self.selection_manager.clear_all_selection()

            if "column_widths" in loaded_data:
                self.column_widths = loaded_data["column_widths"]
                self.update_header_layout()

            sequence_data = loaded_data.get("sequence_data", loaded_data)
            for item in sequence_data:
                new_row = {
                    "id": item["id"],
                    "weighing_path": item["weighing_path"],
                    "method_file": item["method_file"],
                    "reference_material": item.get("reference_material", ""),
                    "equipment": item.get("equipment", ""),
                    "spectrum_path": item["spectrum_path"],
                    "sample_code": item.get("sample_code", ""),
                    "configure_order": item.get("configure_order", ""),
                    "standard_type": item.get("standard_type", ""),
                    "instrument_setting": item.get("instrument_setting", ""),
                    "weighing_mode": item.get("weighing_mode", ""),
                    "temperature": item.get("temperature", ""),
                    "humidity": item.get("humidity", ""),
                    "status": item.get("status", "待运行"),
                    "error_msg": item.get("error_msg", ""),
                    "experiment_code": item.get("experiment_code", "")
                }
                self.sequence_data.append(new_row)

            self.refresh_table()
            self.status_var.set(f"已加载序列文件: {file_path}")

        except Exception as e:
            messagebox.showerror("错误", f"加载失败: {str(e)}")

    def edit_method(self):
        """方法编辑入口：取选中行的方法文件并打开方法编辑器"""
        target = None
        target_idx = None
        sel = sorted(self.selection_manager.selected_rows) or sorted({r for r, _ in self.selection_manager.selected_cells})
        if sel:
            idx = sel[0]
            if 0 <= idx < len(self.sequence_data):
                mf = self.sequence_data[idx].get("method_file")
                if mf:
                    target = mf
                    target_idx = idx
        if not target:
            self._log("未选中含方法文件的行，将打开编辑器(可在编辑器内点'加载')")
        self._open_method_editor(target, target_idx)

    def _open_method_editor(self, file_path=None, target_idx=None):
        """用 Toplevel 打开 MethodRule Editor(带空格文件名，用 importlib 导入)，可选加载指定方法文件"""
        if hasattr(self, "_method_editor_top") and self._method_editor_top is not None and self._method_editor_top.winfo_exists():
            self._method_editor_top.lift()
            self._method_editor_top.focus_force()
            return
        import importlib.util
        editor_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "MethodRule Editor.py")
        if not os.path.exists(editor_path):
            messagebox.showerror("错误", "找不到 MethodRule Editor.py")
            return
        spec = importlib.util.spec_from_file_location("method_rule_editor", editor_path)
        mre = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mre)
        top = tk.Toplevel(self.root)
        top.title("录入方法编辑器")
        app = mre.QueryAppFixed(top)
        if file_path:
            app.load_config_file(file_path)
        self._method_editor_top = top
        top.protocol("WM_DELETE_WINDOW", lambda: self._on_method_editor_closed(top, file_path, target_idx))
        self._log(f"打开方法编辑器{': ' + os.path.basename(file_path) if file_path else ''}")

    def _on_method_editor_closed(self, top, file_path, target_idx):
        """编辑器关闭：若针对某行打开且该行未换文件，重读标液配置回填该行"""
        setattr(self, "_method_editor_top", None)
        top.destroy()
        if not file_path or target_idx is None:
            return
        if not (0 <= target_idx < len(self.sequence_data)):
            return
        row = self.sequence_data[target_idx]
        if row.get("method_file") != file_path:
            return  # 用户在编辑器内加载/另存了别的文件，不回填
        self._apply_method_params(row, file_path)
        self.refresh_table()
        self._log(f"已从方法同步标液+设备配置到第 {target_idx + 1} 行")

    def run_sequence(self):
        """运行序列（阶段1：真实逐行提交，半自动）。支持从指定行开始。"""
        if self._running:
            self.status_var.set("提示: 序列正在运行中")
            return
        if not self.sequence_data:
            self.status_var.set("提示: 没有可运行的序列数据")
            return
        # 起始行(1-based)；空或1=从头
        n = len(self.sequence_data)
        raw = (self.start_row_var.get() or "").strip() if hasattr(self, "start_row_var") else ""
        start_idx = 0
        if raw and raw != "1":
            try:
                sr = int(raw)
            except ValueError:
                self.status_var.set("提示: 起始行需为整数")
                return
            if sr < 1 or sr > n:
                self.status_var.set(f"提示: 起始行需在 1~{n} 之间")
                return
            start_idx = sr - 1
        invalid = [i + 1 for i, r in enumerate(self.sequence_data) if i >= start_idx
                   and (not r["method_file"] or not r["spectrum_path"])]
        if invalid:
            self.status_var.set(f"提示: 第 {', '.join(map(str, invalid))} 行缺少方法文件或谱图路径")
            return
        scope = f"第 {start_idx + 1}~{n} 行" if start_idx else f"全部 {n} 行"
        if not messagebox.askyesno("确认运行", f"确定要运行{scope}序列吗？"):
            return
        if not self.logged_in:
            # 先尝试复用已保存的会话(避免每次运行都重新登录)
            if self.login_system.load_session() and self.login_system.verify_session():
                self.logged_in = True
                self._refresh_user_menu()
                disp = self.login_system.users.get(self.login_system.current_user, {}).get(
                    "display_name", self.login_system.current_user)
                self._log(f"已复用会话自动登录: {disp}")
            if not self.logged_in:
                if not self._show_login_dialog():
                    return
        # 重置状态(仅运行区间；起始行之前的行保持原状态)
        self._stop.clear()
        self._pause.set()
        for i, r in enumerate(self.sequence_data):
            if i < start_idx:
                continue
            r["status"] = "待运行"
            r["error_msg"] = ""
            r["experiment_code"] = ""
            if i < len(self.row_widgets):
                self.row_widgets[i].set_status("待运行")
        self._set_running(True)
        self._log(f"==== 开始运行序列，{scope}（共 {n - start_idx} 行）====")
        self._worker = threading.Thread(target=self._run_worker, args=(start_idx,), daemon=True)
        self._worker.start()

    def _set_running(self, running):
        self._running = running
        self.pause_btn.configure(state='normal' if running else 'disabled', text="⏸ 暂停")
        self.skip_btn.configure(state='normal' if running else 'disabled')
        self.abort_btn.configure(state='normal' if running else 'disabled')
        self.status_var.set("运行中..." if running else "就绪")

    def _run_worker(self, start_idx=0):
        """worker 线程：从 start_idx 起逐行执行，所有 UI 更新经 _ui_q"""
        n = len(self.sequence_data)
        for idx in range(start_idx, n):
            if self._stop.is_set():
                break
            self._pause.wait()
            if self._stop.is_set():
                break
            self._ui_q.put(("status", (idx, "运行中", "")))
            self._log(f"--- 第 {idx + 1}/{n} 行 ---")
            try:
                outcome = self._run_one_row(idx)
            except Exception as e:
                self._ui_q.put(("status", (idx, "失败", str(e))))
                self._log(f"[行{idx + 1}] 异常: {e}")
                outcome = "fail"
            if outcome == "abort":
                self._log("用户中止序列")
                break
        self._ui_q.put(("done", None))

    def _run_one_row(self, idx):
        """单行流水线（worker 线程内）。返回 'ok'/'skip'/'abort'/'fail'，失败自行 put status。
        谱图为目录+多PDF+未填样品编号时，按「称样记录excel ∩ 谱图目录」展开为多个样品；
        跨样品按方法(default_rules.to_id)合并：所有样品的同方法项目合并到一个实验编号提交，
        再按 max_select 分批。某样品查询失败仅跳过(尽量多录入)，任一提交失败整行标失败。"""
        row = self.sequence_data[idx]
        log = lambda m: self._log(f"[行{idx + 1}] {m}")

        # 读称样参数 + 称样记录：record/process 模式必须读(注入称样量)；
        # 其他模式若填了路径也读(供多样品展开确定样品编号，但不注入称样量)
        wp = self._read_weighing_params(row.get("method_file"))
        wmode = (wp.get("weighing_mode") or "").strip() if wp else ""
        rec_path = (row.get("weighing_path") or "").strip()
        wmap = None
        if wmode in ("record", "process"):
            if not rec_path or not os.path.isfile(rec_path):
                self._ui_q.put(("status", (idx, "失败", "未设置称量记录文件")))
                log(f"失败: 称样量模式={wmode} 但未设置称量记录文件路径")
                return "fail"
            wmap, rerr = _read_weighing_records(rec_path)
            if wmap is None:
                self._ui_q.put(("status", (idx, "失败", "称量记录读取失败")))
                log(f"失败: {rerr}")
                return "fail"
        elif rec_path and os.path.isfile(rec_path):
            wmap, _ = _read_weighing_records(rec_path)

        # A/B 平行样合并：称样记录里同编号去尾字母后相同的(如 …001A/B)归为一个样品，
        # 称样量按 A→平行1、B→平行2 拼接，按去尾字母编号(…001)查 LIMS(实验次数由 LIMS 记录数决定)
        if wmap:
            wmap = _merge_parallel_groups(wmap)

        # 解析本行要录入的样品：单样品，或目录多PDF按 excel∩谱图 展开
        samples, serr = self._resolve_samples(row, wmap, log)
        if serr or not samples:
            self._ui_q.put(("status", (idx, "失败", "无法确定样品")))
            log(f"失败: {serr or '谱图目录无匹配PDF且未填样品编号'}")
            return "fail"
        if len(samples) == 1:
            self._ui_q.put(("rowdata", (idx, {"sample_code": samples[0][0]})))
        if len(samples) > 1:
            log(f"本行展开为 {len(samples)} 个样品(称样记录 ∩ 谱图目录)")

        # 谱图检查：按方法配置核对谱图目录PDF是否齐全(空白/标液/线性/样品 各类)，缺则阻止录入
        scp = self._read_spectrum_check_params(row.get("method_file"))
        if scp and any(p["enabled"] for p in scp.values()):
            ok, missing = self._check_spectrum(row, samples, scp, log)
            if not ok:
                self._ui_q.put(("status", (idx, "失败", "谱图检查未通过")))
                log("失败: 谱图检查未通过 - " + "；".join(missing))
                return "fail"
            log("谱图检查通过")

        configure_order = (row.get("configure_order") or "").strip()
        # 标准溶液预检（行级一次，各样品共用）
        if configure_order:
            log(f"校验标准溶液 {configure_order} ...")
            cid, emsg = self.api.get_solution_configure_id(configure_order, log)
            if not cid:
                self._ui_q.put(("status", (idx, "失败", f"标准溶液:{emsg}")))
                log(f"失败: 标准溶液校验未过 - {emsg}")
                return "fail"

        fixed_params = (self._read_other_params(row.get("method_file")) or {}).get("fixed_params") or []
        ctx = {
            "configure_order": configure_order, "wp": wp, "wmode": wmode, "wmap": wmap,
            "fixed_params": fixed_params, "method_file": row.get("method_file") or "",
            "primary_cache": {},  # {sample_code: {masses, desc}} 首项目(苯)生成后供依赖项目(总和)复用
        }

        # 阶段A：逐样品查询+过滤(此阶段不上传谱图)，收集所有样品的项目
        # 某样品查询失败仅跳过(尽量多录入)，不影响其余样品的合并提交
        all_items = []  # [{project, sample_code, sample_id, pdf_paths}, ...]
        n_skip = 0
        for si, (sc, pdf_paths) in enumerate(samples):
            if len(samples) > 1:
                log(f"=== 样品 {si + 1}/{len(samples)}：{sc} ===")
            items, serr = self._collect_sample_projects(idx, row, sc, pdf_paths, ctx, log)
            if serr:
                log(f"警告[{sc}]: {serr}，跳过该样品")
                n_skip += 1
                continue
            all_items.extend(items)
        if not all_items:
            self._ui_q.put(("status", (idx, "失败", "无可用样品项目")))
            log("失败: 所有样品均无可录入项目")
            return "fail"

        # 清空旧实验暂存与谱图(等价前端清空: cancleOcExperiment 连带删除旧谱图)；
        # 必须在上传新谱图之前，否则旧谱图未清、新谱图又传导致重复
        all_pids = ",".join(str(it["project"].get("projectId")) for it in all_items
                            if it["project"].get("projectId"))
        if all_pids:
            log("清空旧实验暂存与谱图 ...")
            self.api.clear_experiment_cache(all_pids, log)

        # 方法勾选「录入前清空谱图」时，额外调 deleteSpectrumByProjectIds 删除已绑定谱图
        if all_pids and self._read_clear_spectrum(row.get("method_file")):
            log("按方法配置删除已绑定谱图 ...")
            self.api.delete_spectrum_by_project_ids(all_pids, log)

        # 批次通用谱图：谱图检查启用类(空白/标液/线性)的 keyword 在谱图目录匹配到的 PDF。
        # 空白/标液/线性为批次共享，每个样品上传时都要带上；样品类不纳入(样品谱图按编号各自上传)
        sp = (row.get("spectrum_path") or "").strip()
        dir_pdfs = sorted(glob.glob(os.path.join(sp, "*.pdf"))) if sp and os.path.isdir(sp) else []
        common_pdfs = []
        for k in ("blank", "standard", "linearity"):
            p = (scp or {}).get(k) or {}
            if p.get("enabled") and p.get("keyword"):
                kw = p["keyword"]
                for n in dir_pdfs:
                    if kw.lower() in os.path.basename(n).lower() and n not in common_pdfs:
                        common_pdfs.append(n)
        if common_pdfs:
            log("批次通用谱图: " + ", ".join(os.path.basename(n) for n in common_pdfs))

        # 上传谱图(清空后旧谱图已删，此时上传不重复)；按样品记录 spectrum_uploaded
        for it in all_items:
            uploaded = []
            for pdf in list(dict.fromkeys((it.get("pdf_paths") or []) + common_pdfs)):
                up_ok, up_res = self.api.upload_spectrum_file(pdf, "", log)
                if up_ok and isinstance(up_res, dict):
                    uploaded.append({
                        "fileId": up_res.get("id"),
                        "fileName": up_res.get("orgName") or up_res.get("name"),
                        "url": up_res.get("url"),
                    })
            it["spectrum_uploaded"] = uploaded
            if it.get("pdf_paths"):
                if uploaded:
                    log(f"[{it['sample_code']}] 谱图已上传 {len(uploaded)} 个: "
                        f"{', '.join(f['fileName'] for f in uploaded)}")
                else:
                    log(f"警告[{it['sample_code']}]: 谱图上传失败({len(it['pdf_paths'])} 个)")

        # 阶段B：按 query_rules 顺序 + input_method(方法=跨样品合并 / 样品=按样品) 规划并提交
        # 方法模式下，同一目标方法(default_rules/filename_rules.to_id)的不同样品项目并入一个实验编号
        method_file = ctx["method_file"]
        default_rules = self._read_default_rules(method_file)
        filename_rules = self._read_filename_rules(method_file)
        if default_rules:
            to_id = {str(r.get("project_name") or "").strip(): str(r.get("to_id") or "").strip()
                     for r in default_rules}
            fallback_mid = ""
        else:
            to_id = {}
            fallback_mid = self._read_switch_method_id(method_file)
        for it in all_items:
            pname = (it["project"].get("projectName") or "").strip()
            it["projectName"] = pname            # 供 _plan_submission_batches 按 project 选
            # 名称切换：样品谱图PDF文件名 + 试样描述 命中 filename_rules 关键字时用其 to_id；否则回落 default_rules
            # 试样描述取自称量记录(无记录则为空，desc 条件不生效)
            wrec = (ctx.get("wmap") or {}).get(it["sample_code"]) or {}
            desc = wrec.get("desc") or ""
            it["wdate"] = _date_str(wrec.get("time"))   # 称样日期分组键；无时间→""(并入合并组)
            mid = _match_filename_rule(filename_rules, pname, it.get("pdf_paths") or [], desc)
            if filename_rules:
                # 测试期日志：显示项目名口径与命中与否，便于核对 filename_rules.project_name 是否对得上 LIMS
                if mid:
                    log(f"[{it['sample_code']}] 命中名称切换 → 方法ID {mid}（项目:{pname}）")
                else:
                    log(f"[{it['sample_code']}] 未命中名称切换（项目:{pname}），用默认方法")
            it["switch_mid"] = mid or to_id.get(pname, fallback_mid)

        rules = self._read_query_rules(method_file)
        max_sel = self._read_max_select(method_file)
        plan = _plan_submission_batches(rules, all_items, max_sel)

        codes = []
        for b in plan:
            sb = [it["sample_code"] for it in b["items"]]
            dw = f"〔称样{b['wdate']}〕" if b.get("wdate") else ""
            if b["switch_mid"]:
                log(f"切换到方法ID {b['switch_mid']}（{len(sb)} 个样品: {', '.join(sb)}）{dw}")
            else:
                log(f"未匹配切换规则（{len(sb)} 个样品）{dw}，用默认方法")
            ok, code = self._submit_batch(idx, row, b["items"], ctx, log, b["force_new"], b["switch_mid"])
            if not ok:
                return "fail"
            codes.append(code)

        real_code = " / ".join(codes)
        self._ui_q.put(("status", (idx, "成功", "")))
        self._ui_q.put(("rowdata", (idx, {"experiment_code": real_code})))
        extra = f"，跳过 {n_skip} 个样品" if n_skip else ""
        log(f"成功，{len(samples) - n_skip}/{len(samples)} 个样品参与合并，实验编号 {real_code}{extra}")
        return "ok"

    def _collect_sample_projects(self, idx, row, sample_code, pdf_paths, ctx, log):
        """阶段A：查询样品 + 按方法过滤（谱图上传推迟到清空旧数据之后，避免重跑重复上传）。
        返回 (items, err)。items = [{project, sample_code, sample_id, pdf_paths}, ...]。
        err 非空表示该样品不可用(查不到/无匹配项目)，调用方跳过该样品。"""
        log(f"查询样品 {sample_code} ...")
        projects = self.api.query_samples_by_conditions(
            sample_code=sample_code, exact_match=True, log_func=log)
        if not projects:
            return [], "LIMS 未查到该样品(可能未登记或超30天)"
        projects, ferr = self._filter_projects_by_method(projects, row, log)
        if ferr:
            return [], ferr
        if not projects:
            return [], "该样品下没有匹配方法文件的项目"
        sample_id = projects[0].get("sampleId")
        items = [{"project": p, "sample_code": sample_code, "sample_id": sample_id,
                  "pdf_paths": pdf_paths} for p in projects]
        return items, None

    def _submit_batch(self, idx, row, batch_items, ctx, log, force_new, switch_mid=""):
        """提交一批 projects（worker 线程内，跨样品合并）。返回 (ok, real_code)，失败自行 put status。
        batch_items: [{project, sample_code, sample_id, spectrum_uploaded}, ...]，含多个样品的
        同方法项目；同一目标方法(to_id)的不同样品项目合并到一个实验编号。
        switch_mid: 本批目标方法ID(按项目名匹配 default_rules)；空则不切换。
        force_new: 多批时传 True，使每批生成独立实验编号(避免同号冲突)。"""
        batch_projects = [it["project"] for it in batch_items]
        project_ids = [str(p["projectId"]) for p in batch_projects if p.get("projectId")]
        project_names = [p.get("projectName", "") for p in batch_projects]
        sp_ids_str = ",".join(project_ids)
        method_name = batch_projects[0].get("standardNo") or ""
        sample_id = batch_items[0]["sample_id"] if batch_items else ""  # 辅助参数(构建Referer)，跨样品取首个

        # 配置中段（复刻 submit_single_method_group 纯 API 部分）；clear 已在 _run_one_row
        # 上传谱图前统一完成——cancleOcExperiment 连带删谱图，此处再 clear 会删掉刚上传的新谱图
        if switch_mid:
            # 方法配置了目标切换方法ID：先切换服务端项目方法，再用目标方法取配置(默认方法可能缺计算公式)
            log(f"切换检测方法 -> 方法ID {switch_mid}")
            self.api.update_method(sp_ids_str, switch_mid, project_names, log)
            switched_std = self.api.get_method_standard_no_by_id(switch_mid, log)
            if switched_std:
                method_name = switched_std
            method_id = switch_mid
            actual_method_name = method_name
            actual_method_id = method_id
        else:
            log("取实验配置 ...")
            log(f"取实验配置: sp_ids={sp_ids_str} method={method_name!r} sample_id={sample_id!r}")
            initial = self.api.get_experiment_config(sp_ids_str, method_name, "", sample_id, log)
            if not initial:
                self._ui_q.put(("status", (idx, "失败", "取实验配置失败")))
                log("失败: get_experiment_config 返回空")
                return False, ""
            oc = initial.get("ocMethodSettings", {}) or {}
            method_id = oc.get("methodId")
            actual_method_name = method_name
            actual_method_id = method_id
            if method_id and str(method_id) in self.api.sub_method_map:
                sub = self.api.sub_method_map[str(method_id)]
                actual_method_id = sub.get("sub_method_id")
                self.api.update_method(sp_ids_str, actual_method_id, project_names, log)
                std = self.api.get_method_standard_no_by_id(actual_method_id, log)
                if std:
                    actual_method_name = std
        log("取全量配置 ...")
        all_cfg = self.api.get_all_configs(sp_ids_str, method_name, "", sample_id, log, method_id, project_names)
        if not all_cfg:
            self._ui_q.put(("status", (idx, "失败", "取全量配置失败")))
            return False, ""
        experiment_config = all_cfg.get("experiment")
        # 平行样扩展：称样记录平行数多于 LIMS 记录数时，复制各组分记录补平行(实验次数=平行数)
        pid_to_sample = {str(it["project"].get("projectId")): it["sample_code"] for it in batch_items}
        _par_by = {}
        for _sc in dict.fromkeys(it["sample_code"] for it in batch_items):
            _mm = ((ctx["wmap"] or {}).get(_sc) or {}).get("masses") or []
            if len(_mm) > 1:
                _par_by[_sc] = len(_mm)
        _old_n, _new_n = _expand_parallel_records(experiment_config, pid_to_sample, _par_by)
        if _new_n > _old_n:
            log(f"平行样扩展：ocAnalysisRecordList {_old_n}→{_new_n} 条(按称样记录平行数补平行)")
        equipment_config = all_cfg.get("equipment")
        # 设备以序列表格为准：表格指定设备编号则覆盖默认主检设备
        equipment_config, matched_eq, eq_err = self._override_equipment(equipment_config, row.get("equipment", ""))
        if eq_err:
            self._ui_q.put(("status", (idx, "失败", eq_err)))
            log(f"失败: {eq_err}")
            return False, ""
        dynamic_columns = all_cfg.get("dynamic_columns") or []
        actual_method_name = all_cfg.get("actual_method_name", actual_method_name)
        actual_method_id = all_cfg.get("actual_method_id", actual_method_id)

        # 实验编号（force_new=True 时每批独立，避免同号冲突）
        log("生成实验编号 ...")
        experiment_code = self.api.generate_experiment_code(
            method_name=actual_method_name, log_func=log, force_new=force_new)
        if not experiment_code:
            self._ui_q.put(("status", (idx, "失败", "生成实验编号失败")))
            return False, ""

        # 动态列取值：用列默认值（None→""，避免 build 把字符串"None"当用户输入、覆盖记录原值，如组分名称）
        values = {col.get("columeCode", ""): str(col.get("defaultVal") or "") for col in dynamic_columns}
        host = self._build_headless_host(
            values, experiment_config, equipment_config, dynamic_columns,
            actual_method_name, actual_method_id, ctx["configure_order"], ctx["fixed_params"])

        # 温湿度：表格手填优先 → 设备房间·当天自动(温湿度记录.xlsx) → 默认 22/55
        temp_val = (row.get("temperature") or "").strip()
        hum_val = (row.get("humidity") or "").strip()
        if not temp_val or not hum_val:
            e2r, renv, _ = _read_env_records(_ENV_RECORD_PATH)
            if e2r is not None:
                cands = _equipment_env_candidates(row, equipment_config, matched_eq)
                t, h = _resolve_env(e2r, renv, cands, datetime.now().strftime("%Y-%m-%d"))
                if t and not temp_val:
                    temp_val = t
                if h and not hum_val:
                    hum_val = h
        host.temperature_var = _Box(temp_val)
        host.humidity_var = _Box(hum_val)
        log(f"温湿度: 温度={temp_val or '空'}℃ 湿度={hum_val or '空'}%RH")

        # 称样量注入：跨样品按 projectId -> 样品 -> 各自 pmasses[平行索引]
        wmode = ctx["wmode"]
        wp = ctx["wp"]
        analysis_start = None  # 称量记录「称样时间」→ 实验分析开始时间(跨样品取首个样品)
        batch_samples = list(dict.fromkeys(it["sample_code"] for it in batch_items))

        def _fill_desc_column(records, desc_by_sample):
            """试样信息/试样描述列：按记录所属样品的试样描述逐条填充(跨样品各不同)。"""
            info_col = _find_column_by_name(dynamic_columns, "试样信息", "试样描述")
            if info_col is None:
                log("未找到试样信息/试样描述列，跳过试样描述填充")
                return
            info_code = info_col.get("columeCode", "")
            host.data_fields[info_code] = [
                _Box(desc_by_sample.get(pid_to_sample.get(str(r.get("projectId")), ""), ""))
                for r in records or []]
            log(f"试样描述({info_col.get('columeName', '')}) 已按样品填充")

        # 跨项目复用(苯→总和)：首项目生成的试样描述/称样量缓存后，后续项目直接复用，保证同样品一致
        # 免去 getOcCompareShowData 网络往返；query_rules 顺序保证苯先于总和生成
        primary_cache = ctx["primary_cache"]
        cached_before = set(primary_cache)  # 本批开始前已缓存的样品 → 这些样品将复用缓存

        if wmode == "random":
            mass_col = _find_weighing_column(dynamic_columns)
            if mass_col is not None:
                mass_code = mass_col.get("columeCode", "")
                records = (experiment_config or {}).get("ocAnalysisRecordList") or []
                _, n_par = _parallel_indices(records)
                pmasses_by_sample, desc_by_sample = {}, {}
                for it in batch_items:
                    sc = it["sample_code"]
                    if sc in pmasses_by_sample:
                        continue
                    rv = primary_cache.get(sc)
                    if rv:
                        pmasses_by_sample[sc] = list(rv["masses"])
                        desc_by_sample[sc] = rv["desc"]
                    else:
                        pmasses_by_sample[sc] = [_gen_random_mass(wp) for _ in range(n_par)]
                        desc_by_sample[sc] = ((ctx["wmap"] or {}).get(sc) or {}).get("desc") or ""
                        primary_cache[sc] = {"masses": list(pmasses_by_sample[sc]), "desc": desc_by_sample[sc]}
                host.data_fields[mass_code] = _mass_field_by_project(records, pid_to_sample, pmasses_by_sample)
                log(f"称样量(random) {mass_col.get('columeName', '')} 跨{len(pmasses_by_sample)}样品 按平行({n_par})")
                _fill_desc_column(records, desc_by_sample)
            else:
                log("称样量(random) 未找到称量记录列(isWeighing=1)，跳过")
            first_sc = batch_items[0]["sample_code"] if batch_items else ""
            analysis_start = ((ctx["wmap"] or {}).get(first_sc) or {}).get("time")
        elif wmode == "none":
            log("称样量模式=无需称样量，跳过称样列")
        elif wmode in ("record", "process"):
            mass_col = _find_weighing_column(dynamic_columns)
            if not mass_col:
                log(f"称样量({wmode}) 未找到称量记录列(isWeighing=1)，跳过")
            else:
                records = (experiment_config or {}).get("ocAnalysisRecordList") or []
                # 各样品平行数 = 该样品项目记录的最大 serialNumber(扩展后)；混批时各样品可不同
                n_par_by_sample = {}
                for _r in records:
                    _rsc = pid_to_sample.get(str(_r.get("projectId")))
                    if not _rsc:
                        continue
                    try:
                        _rsv = int(_r.get("serialNumber"))
                    except (TypeError, ValueError):
                        _rsv = 1
                    n_par_by_sample[_rsc] = max(n_par_by_sample.get(_rsc, 0), _rsv)
                prules = self._read_processing_rules(ctx["method_file"]) if wmode == "process" else []
                rule = _match_processing_rule(prules, actual_method_name) if wmode == "process" else None
                pmasses_by_sample, desc_by_sample = {}, {}
                for it in batch_items:
                    sc = it["sample_code"]
                    if sc in pmasses_by_sample:
                        continue
                    rv = primary_cache.get(sc)
                    if rv:
                        pmasses_by_sample[sc] = list(rv["masses"])
                        desc_by_sample[sc] = rv["desc"]
                        continue
                    samp = ctx["wmap"].get(sc) or {}
                    masses = samp.get("masses")
                    n_par = n_par_by_sample.get(sc, 1)
                    if not masses or len(masses) < n_par:
                        self._ui_q.put(("status", (idx, "失败", f"称量记录平行不足({len(masses or [])}/{n_par})")))
                        log(f"失败: 样品 {sc} 称量记录仅 {len(masses or [])} 个平行，实验需 {n_par}")
                        return False, ""
                    if wmode == "record":
                        pmasses_by_sample[sc] = [_raw_mass_str(masses[i]) for i in range(n_par)]
                    else:
                        pmasses_by_sample[sc] = [_apply_processing(masses[i], rule, wp) for i in range(n_par)]
                    desc_by_sample[sc] = samp.get("desc") or ""
                    primary_cache[sc] = {"masses": list(pmasses_by_sample[sc]), "desc": desc_by_sample[sc]}
                mass_code = mass_col.get("columeCode", "")
                host.data_fields[mass_code] = _mass_field_by_project(records, pid_to_sample, pmasses_by_sample)
                rname = ((rule or {}).get("type") or "直接读取(无匹配规则)") if wmode == "process" else "record"
                _npars = sorted(set(n_par_by_sample.values())) if n_par_by_sample else [1]
                log(f"称样量({wmode}) {mass_col.get('columeName', '')} 规则[{rname}] 跨{len(pmasses_by_sample)}样品 按平行({'/'.join(map(str, _npars))})")
                _fill_desc_column(records, desc_by_sample)
                first_sc = batch_items[0]["sample_code"] if batch_items else ""
                analysis_start = (ctx["wmap"].get(first_sc) or {}).get("time")
        elif wmode:
            log(f"称样量模式={wmode} 暂未接入(本轮支持 random/none/record/process)")

        _reused = [sc for sc in batch_samples if sc in cached_before]
        if _reused:
            log(f"复用首项目(苯)试样描述/称样量: {len(_reused)}/{len(batch_samples)} 样品")

        # 构造载荷 + 注入谱图 + 提交
        log("构建并提交 ...")
        experiment_data = build_grouped_experiment_data(host, batch_projects, experiment_code, actual_method_name)
        # 称量记录「称样时间」→ 覆盖实验分析开始时间 startTime
        if analysis_start is not None:
            try:
                experiment_data["startTime"] = analysis_start.strftime("%Y-%m-%d %H:%M:%S")
                log(f"分析开始时间(startTime) <- 称样时间: {experiment_data['startTime']}")
            except (AttributeError, ValueError):
                log(f"警告: 称样时间格式无法解析({analysis_start!r})，startTime 保持默认")
        # 注入谱图 fileIds/spectrumJsonList：跨样品按样品分组，每样品谱图绑定该样品在本批的 projectId 串
        # （对照 detection_entry_main:1718-1725）
        sample_pids = {}
        for it in batch_items:
            sample_pids.setdefault(it["sample_code"], []).append(str(it["project"].get("projectId")))
        spec_list, all_file_ids, seen = [], [], set()
        for it in batch_items:
            pids = ",".join(sample_pids.get(it["sample_code"], []))
            for f in it["spectrum_uploaded"] or []:
                fid = f.get("fileId")
                if not fid or fid in seen:
                    continue
                seen.add(fid)
                spec_list.append({"fileId": fid, "fileName": f.get("fileName"), "projectId": pids})
                all_file_ids.append(str(fid))
        if spec_list:
            experiment_data["fileIds"] = ",".join(all_file_ids)
            experiment_data["spectrumJsonList"] = json.dumps(spec_list)
        ok, _ = self.api.submit_experiment_data(experiment_data, actual_method_name, log)
        if not ok:
            self._ui_q.put(("status", (idx, "失败", "实验数据提交失败")))
            log("失败: submit_experiment_data 返回 False")
            return False, ""

        # 回读服务端真实实验编号 + experiment_id（对照 detection_entry_main:1730-1752）
        real_code = experiment_code
        experiment_id = 0
        first_pid = project_ids[0] if project_ids else ""
        if first_pid:
            try:
                cfg = self.api.get_experiment_config(first_pid, "", "", "", log)
                if isinstance(cfg, dict):
                    real_code = self.api.extract_experiment_code(cfg, experiment_code) or experiment_code
                    try:
                        experiment_id = int(cfg.get('id') or 0)
                    except (TypeError, ValueError):
                        experiment_id = 0
            except Exception as e:
                log(f"读取真实实验编号失败: {e}")

        # 主检设备保存（对照 detection_entry_main:1757-1766）：表格指定了设备则提交该设备；
        # 未指定(默认模式)则提交方法全部默认检测设备(isDefault=1)，与手动界面「查询设备」默认勾选一致
        if matched_eq and experiment_id:
            eid = matched_eq.get("equipmentBillId") or matched_eq.get("id")
            items = [{"id": eid, "label": "", "raw": matched_eq}]
            eq_ok, _ = self.api.save_main_equipment(experiment_id, items, log)
            log(f"主检设备{'提交成功' if eq_ok else '提交失败'}: {eid}")
        elif not matched_eq:
            if not experiment_id:
                log("主检设备未单独提交：experimentId=0（实验编号未生成）")
            else:
                # 默认主检设备 = 方法检测设备中 isDefault=1 的（即 mainEquipmentIds 列出的）
                main_ids = set((equipment_config or {}).get("mainEquipmentIds", "").split(","))
                default_items = [
                    {"id": (eq.get("equipmentBillId") or eq.get("id")), "label": "", "raw": eq}
                    for eq in (equipment_config or {}).get("raw_data") or []
                    if eq.get("usedCategory") == "检测设备"
                    and (eq.get("equipmentBillId") or eq.get("id"))
                    and str(eq.get("equipmentBillId") or eq.get("id")) in main_ids
                ]
                if default_items:
                    eq_ok, _ = self.api.save_main_equipment(experiment_id, default_items, log)
                    log(f"主检设备{'提交成功' if eq_ok else '提交失败'}(默认{len(default_items)}台): "
                        f"{','.join(str(i['id']) for i in default_items)}")
                else:
                    log("主检设备未单独提交：方法未配置默认检测设备(已随实验载荷提交)")

        # 标液关联——用真实编号（对照 detection_entry_main:1768-1776）
        if ctx["configure_order"]:
            experiment_codes = {pid: real_code for pid in project_ids}
            sok, serr = self.api.submit_solution_with_experiment(
                project_ids, ctx["configure_order"], experiment_codes, log)
            if not sok:
                self._ui_q.put(("status", (idx, "失败", f"标准溶液关联:{serr}")))
                log(f"失败: 标准溶液关联失败 - {serr}")
                return False, ""

        log(f"本批完成，实验编号 {real_code}")
        return True, real_code

    def _resolve_samples(self, row, wmap, log):
        """解析本行要录入的样品。返回 (samples, err)。
        samples = [(sample_code, [pdf_path,...]), ...]；err 非空=无法解析。
        - 谱图是文件 / 目录单PDF / 目录+已填样品编号：单样品
        - 目录+多PDF+未填样品编号：用 wmap(称样记录)样品编号 ∩ 目录PDF(startswith)展开；
          wmap 缺失则 err。一个样品匹配多个PDF(如A/B平行)全部收集、都上传。"""
        sp = (row.get("spectrum_path") or "").strip()
        sc = (row.get("sample_code") or "").strip()
        if os.path.isfile(sp):
            code = sc or os.path.splitext(os.path.basename(sp))[0].split("-", 1)[0]
            return [(code, [sp])], None
        pdfs = sorted(glob.glob(os.path.join(sp, "*.pdf"))) if os.path.isdir(sp) else []
        if not pdfs:
            return [], "谱图路径无效或目录内无PDF"
        if sc:
            key = _strip_parallel_suffix(sc)  # 称样编号带平行小号(001)、谱图文件名不带：用去小号后的前缀匹配
            matched = [p for p in pdfs if os.path.basename(p).startswith(key)]
            if not matched:
                log(f"警告: 目录内无文件名以 {key} 开头的PDF")
            return [(sc, matched)], None
        if len(pdfs) == 1:
            code = os.path.splitext(os.path.basename(pdfs[0]))[0].split("-", 1)[0]
            return [(code, [pdfs[0]])], None
        # 多PDF + 未填样品编号 → 用称样记录编号 ∩ 目录PDF 展开(只有两边都有的编号才参与录入)
        if not wmap:
            return [], "目录有多个PDF且未填样品编号，请在「称样记录路径」填称样记录excel(按 excel∩谱图 展开)"
        # 先收集有谱图的候选 (编号, 试样描述, 匹配PDF)
        cands = []
        for code, entry in wmap.items():
            key = _strip_parallel_suffix(code)  # 称样编号带平行小号(001)、谱图文件名不带：去小号后匹配
            matched = [p for p in pdfs if os.path.basename(p).startswith(key)]
            if matched:
                cands.append((code, (entry or {}).get("desc") or "", matched))
        if not cands:
            return [], "称样记录中的样品编号在谱图目录内均无匹配PDF"
        # 混目录分流：方法 filename_rules 含 desc 时，只保留试样描述命中本方法的样品
        # (固体/液体各走各自方法，避免液体样品混进固体方法卡在谱图检查)
        filename_rules = self._read_filename_rules(row.get("method_file"))
        if any(str(r.get("desc") or "").strip() for r in filename_rules):
            kept, skipped = [], []
            for code, desc, matched in cands:
                if _rule_desc_match(filename_rules, desc):
                    kept.append((code, matched))
                else:
                    skipped.append(code)
            if not kept:
                return [], "称样记录中无试样描述命中本方法(filename_rules.desc)的样品，请检查方法与称样记录"
            if skipped:
                log(f"样品分流(按 filename_rules.desc)：跳过 {len(skipped)} 个：{'、'.join(skipped)}")
            samples = kept
        else:
            samples = [(code, matched) for code, _desc, matched in cands]
        return samples, None

    def _resolve_spectrum_pdf(self, row, log):
        """定位本行谱图PDF(目录内按文件名前缀=sampleCode匹配)。返回 (pdf_path, sample_code)"""
        sp = row.get("spectrum_path") or ""
        sc = (row.get("sample_code") or "").strip()
        if os.path.isfile(sp):
            return sp, (sc or os.path.splitext(os.path.basename(sp))[0].split("-", 1)[0])
        pdfs = sorted(glob.glob(os.path.join(sp, "*.pdf"))) if os.path.isdir(sp) else []
        if not pdfs:
            return None, sc
        if sc:
            for p in pdfs:
                if os.path.basename(p).startswith(sc):
                    return p, sc
            log(f"警告: 目录内无文件名以 {sc} 开头的PDF")
            return None, sc
        if len(pdfs) == 1:
            return pdfs[0], os.path.splitext(os.path.basename(pdfs[0]))[0].split("-", 1)[0]
        log(f"目录有 {len(pdfs)} 个PDF且未填样品编号，无法确定(请一行一PDF或先填sample_code)")
        return None, ""

    def _read_query_rules(self, method_file):
        """读方法文件 query_rules，返回有序规则列表；兼容 dict/list 两种 yaml 结构。空/异常返回 []。"""
        if not method_file or not os.path.isfile(method_file):
            return []
        try:
            with open(method_file, "r", encoding="utf-8") as f:
                y = yaml.safe_load(f) or {}
        except Exception:
            return []
        qr_raw = y.get("query_rules")
        if isinstance(qr_raw, dict):
            return qr_raw.get("query_rules") or []
        if isinstance(qr_raw, list):
            return qr_raw
        return []

    def _filter_projects_by_method(self, projects, row, log):
        """按方法文件 query_rules 指定的方法过滤 projects。
        返回 (projects, error_msg)。error_msg 非空表示无法确定单一方法——
        不回退到全部：多方法会让 getOcExperiment 报"样品项目对应的方法不同"。"""
        # 样品实际含哪些方法(标准号)
        stdnos = []
        for p in projects:
            s = (p.get("standardNo") or "").strip()
            if s and s not in stdnos:
                stdnos.append(s)

        method_file = row.get("method_file") or ""
        qr = self._read_query_rules(method_file)
        method_val = str(qr[0]["method"]).strip() if (qr and qr[0].get("method")) else ""
        project_vals = []
        for q in qr:
            pv = str(q.get("project") or "").strip()
            if pv and pv not in project_vals:
                project_vals.append(pv)

        # project 过滤：项目名命中任一 query rule 的 project(含通配*)即保留；规则均无 project 则全中
        def _project_hit(p):
            if not project_vals:
                return True
            pname = (p.get("projectName") or "").strip()
            return any(_project_match(pname, pv) for pv in project_vals)

        if method_val:
            # method 可能是数字方法ID(如 '36246')，也可能是标准号文本(如 'GB/T 23991-2009' / 'GB 36246-2018 附录G')。
            # 文本直接当标准号匹配，避免被当ID查服务端(getObj?id=<标准号>)返回错误方法。
            if method_val.isdigit():
                try:
                    target_std = (self.api.get_method_standard_no_by_id(method_val, log) or "").strip()
                except Exception:
                    target_std = ""
            else:
                target_std = method_val
            if target_std:
                def _hit(std):
                    s = (std or "").strip()
                    if not s:
                        return False
                    if s == target_std:
                        return True
                    # 容忍"附录X"后缀差异：'GB 36246-2018' 匹配 'GB 36246-2018 附录G'；要求足够长防短串误匹配
                    if len(target_std) >= 6 and len(s) >= 6:
                        return s.startswith(target_std) or target_std.startswith(s)
                    return False
                filtered = [p for p in projects if _hit(p.get("standardNo"))]
                filtered = [p for p in filtered if _project_hit(p)]
                if filtered:
                    proj_hint = f"，项目名过滤={project_vals!r}" if project_vals else ""
                    log(f"按方法 {target_std} 过滤出 {len(filtered)} 个项目(查询方法值={method_val!r}{proj_hint})")
                    return filtered, None
                return None, (f"方法文件指定的方法「{target_std}」(查询方法值={method_val!r})不在此样品项目中。"
                              f"样品实际方法: {', '.join(stdnos)}。请检查录入方法文件。")
            return None, (f"方法文件指定了方法「{method_val}」但无法确定标准号。"
                          f"样品实际方法: {', '.join(stdnos)}")

        # 方法文件未指定方法：仅当样品只含单一方法时才可用全部(再按 project 过滤)
        if len(stdnos) <= 1:
            return [p for p in projects if _project_hit(p)], None
        return None, (f"样品含 {len(stdnos)} 个不同方法，但方法文件未在 query_rules 指定要提交的方法。"
                      f"方法列表: {', '.join(stdnos)}")

    def _read_max_select(self, method_file):
        """读方法文件 query_rules[0].max_select（最大选择量=每批样品数上限）。返回 int；
        空/非数字/≤0 返回 0(不限)。兼容 query_rules 的 dict/list 两种 yaml 结构。"""
        if not method_file or not os.path.isfile(method_file):
            return 0
        try:
            with open(method_file, "r", encoding="utf-8") as f:
                y = yaml.safe_load(f) or {}
            qr_raw = y.get("query_rules")
            if isinstance(qr_raw, dict):
                qr = qr_raw.get("query_rules") or []
            elif isinstance(qr_raw, list):
                qr = qr_raw
            else:
                qr = []
            if not qr:
                return 0
            ms = str(qr[0].get("max_select") or "").strip()
            if not ms:
                return 0
            n = int(ms)
            return n if n > 0 else 0
        except Exception:
            return 0

    def _read_switch_method_id(self, method_file):
        """读方法文件 query_rules[0].method_id（录入前要切换到的目标方法ID）。返回 str；空则 ''。
        用于 PD-苯 等需从默认方法切换到子方法才含计算公式的场景(对照主窗口 switch_method_id)。"""
        if not method_file or not os.path.isfile(method_file):
            return ""
        try:
            with open(method_file, "r", encoding="utf-8") as f:
                y = yaml.safe_load(f) or {}
            qr_raw = y.get("query_rules")
            if isinstance(qr_raw, dict):
                qr = qr_raw.get("query_rules") or []
            elif isinstance(qr_raw, list):
                qr = qr_raw
            else:
                qr = []
            if not qr:
                return ""
            return str(qr[0].get("method_id") or "").strip()
        except Exception:
            return ""

    def _read_default_rules(self, method_file):
        """读方法文件 default_rules（方法切换规则: project_name -> to_id）。返回 list；空/异常返回 []。"""
        if not method_file or not os.path.isfile(method_file):
            return []
        try:
            with open(method_file, "r", encoding="utf-8") as f:
                y = yaml.safe_load(f) or {}
            return y.get("default_rules") or []
        except Exception:
            return []

    def _read_filename_rules(self, method_file):
        """读方法文件 filename_rules（名称切换规则: filename 关键字 + project_name -> to_id）。返回 list；空/异常返回 []。"""
        if not method_file or not os.path.isfile(method_file):
            return []
        try:
            with open(method_file, "r", encoding="utf-8") as f:
                y = yaml.safe_load(f) or {}
            return y.get("filename_rules") or []
        except Exception:
            return []

    def _read_spectrum_check_params(self, method_file):
        """读方法文件 spectrum_upload_settings.spectrum_check_params。
        返回 {blank/standard/linearity/sample: {enabled,count,keyword}}；空/异常返回 {}。
        count 解析为 int(非数字/空=0)；keyword 去空白。"""
        if not method_file or not os.path.isfile(method_file):
            return {}
        try:
            with open(method_file, "r", encoding="utf-8") as f:
                y = yaml.safe_load(f) or {}
            params = (y.get("spectrum_upload_settings") or {}).get("spectrum_check_params") or {}
            out = {}
            for k in ("blank", "standard", "linearity", "sample"):
                item = params.get(k) or {}
                cnt = str(item.get("count") or "").strip()
                out[k] = {
                    "enabled": bool(item.get("enabled")),
                    "count": int(cnt) if cnt.lstrip("-").isdigit() else 0,
                    "keyword": str(item.get("keyword") or "").strip(),
                }
            return out
        except Exception:
            return {}

    def _read_clear_spectrum(self, method_file):
        """读方法文件 spectrum_upload_settings.clear_spectrum；异常返回 False。"""
        if not method_file or not os.path.isfile(method_file):
            return False
        try:
            with open(method_file, "r", encoding="utf-8") as f:
                y = yaml.safe_load(f) or {}
            return bool((y.get("spectrum_upload_settings") or {}).get("clear_spectrum"))
        except Exception:
            return False

    def _check_spectrum(self, row, samples, params, log):
        """谱图检查：核对本行谱图目录的 PDF 是否齐全。
        空白/标液/线性(批次共享)：目录中文件名含「关键字」的 PDF 数须 ≥「数量」(数量0按1)。
        样品(每样品维度)：关键字为后缀列表(逗号分隔，如"A,B")，每个样品对每个后缀都需有
        「样品编号+后缀」的 PDF 各 ≥「数量」份(如 TN...001A、TN...001B 各1份)。
        返回 (ok, missing_list)；missing_list 为各类缺失描述，通过时为 []。"""
        sp = (row.get("spectrum_path") or "").strip()
        if not sp or not os.path.isdir(sp):
            return False, ["谱图路径无效，无法检查"]
        names = [os.path.basename(p) for p in glob.glob(os.path.join(sp, "*.pdf"))]
        labels = {"blank": "空白", "standard": "标液", "linearity": "线性"}
        missing = []
        # 批次级：空白/标液/线性(全目录含关键字 ≥ 数量)
        for k in ("blank", "standard", "linearity"):
            p = params.get(k) or {}
            if not p.get("enabled"):
                continue
            kw = p.get("keyword")
            if not kw:
                continue  # 启用但无关键字：跳过(UI 校验已拦，运行时再防御)
            need = p.get("count") or 1
            actual = sum(1 for n in names if kw.lower() in n.lower())
            if actual < need:
                missing.append(f"{labels[k]}谱图需≥{need}个含'{kw}'，实际{actual}个")
        # 样品级：每样品 × 每后缀(后缀取自关键字逗号分隔) 各需 need 份；
        # 关键字为空时按样品编号前缀匹配(如 TN26070485 → TN26070485001.pdf)，与上传匹配逻辑一致
        ps = params.get("sample") or {}
        if ps.get("enabled"):
            suf_str = ps.get("keyword")
            suffixes = [s.strip() for s in suf_str.split(",") if s.strip()] if suf_str else [""]
            need = ps.get("count") or 1
            for sc, _pdfs in samples:
                base = _strip_parallel_suffix(sc).lower()  # 去平行小号(001)后的样品前缀
                for suf in suffixes:
                    suf_l = suf.lower()
                    # 文件名 = 前缀 [+ 平行字母 A/B] + 后缀(如 TN26070477AK)：以前缀开头且含后缀即算，
                    # 兼容平行样(字母夹在前缀与后缀之间)与单样(前缀直接接后缀)
                    actual = sum(1 for n in names if n.lower().startswith(base) and suf_l in n.lower())
                    if actual < need:
                        tag = f"{suf}谱图" if suf else "谱图"
                        missing.append(f"样品{sc}的{tag}需≥{need}个，实际{actual}个")
        return (not missing), missing

    def _build_headless_host(self, values, experiment_config, equipment_config,
                             dynamic_columns, actual_method_name, actual_method_id, configure_order,
                             fixed_params=None):
        """构造无头 host 供 build_grouped_experiment_data 使用"""
        data_fields = {code: _Box(val) for code, val in values.items()}
        today = datetime.now().strftime("%Y-%m-%d")
        return types.SimpleNamespace(
            login_system=self.login_system,
            api=self.api,
            log=lambda m: self._log(f"[行] {m}"),
            experiment_config=experiment_config or {},
            equipment_config=equipment_config or {},
            dynamic_columns=dynamic_columns or [],
            data_fields=data_fields,
            remark_text=_Box(""),
            solution_type_var=_Box(configure_order or ""),
            actual_method_name=actual_method_name,
            actual_method_id=actual_method_id,
            fixed_params=fixed_params or [],
            temperature_var=_Box(""),
            humidity_var=_Box(""),
            start_date_var=_Box(today),
            end_date_var=_Box(today),
        )

    def _prompt_confirm(self, prompt):
        """worker 线程：请求主线程弹确认面板并阻塞。返回 dict/'skip'/'abort'"""
        self._confirm_done.clear()
        self._confirm_result = None
        self._ui_q.put(("prompt", prompt))
        self._confirm_done.wait()
        return self._confirm_result

    def _drain_ui_queue(self):
        """主线程：轮询 worker 消息并更新 UI"""
        try:
            while True:
                kind, payload = self._ui_q.get_nowait()
                if kind == "log":
                    self._append_log(payload)
                elif kind == "status":
                    idx, st, err = payload
                    if idx < len(self.row_widgets):
                        self.row_widgets[idx].set_status(st, err)
                    if idx < len(self.sequence_data):
                        self.sequence_data[idx]["status"] = st
                        self.sequence_data[idx]["error_msg"] = err
                elif kind == "rowdata":
                    idx, extras = payload
                    if idx < len(self.sequence_data):
                        self.sequence_data[idx].update(extras)
                elif kind == "prompt":
                    self._build_confirm_dialog(payload)
                elif kind == "done":
                    self._on_run_done()
        except queue.Empty:
            pass
        self.root.after(100, self._drain_ui_queue)

    def _on_run_done(self):
        self._set_running(False)
        self._confirm_done.set()
        self._append_log("==== 序列运行结束 ====")
        ok = sum(1 for r in self.sequence_data if r.get("status") == "成功")
        fail = sum(1 for r in self.sequence_data if r.get("status") == "失败")
        self.status_var.set(f"完成: 成功 {ok} / 失败 {fail}")

    def _append_log(self, msg):
        self.log_text.configure(state='normal')
        self.log_text.insert('end', msg + "\n")
        self.log_text.see('end')
        self.log_text.configure(state='disabled')

    def _log(self, msg):
        self._ui_q.put(("log", msg))

    def _build_confirm_dialog(self, prompt):
        """主线程：构建确认面板（非 grab，不挡表格/日志）"""
        if self._confirm_win is not None and self._confirm_win.winfo_exists():
            self._confirm_win.destroy()
        win = tk.Toplevel(self.root)
        win.title(f"确认提交 - 第 {prompt['idx'] + 1} 行")
        win.geometry("640x560")
        win.transient(self.root)

        info = tk.Frame(win, padx=10, pady=8)
        info.pack(fill='x')
        tk.Label(info, text=f"样品: {prompt['sample_code']}", anchor='w').pack(fill='x')
        tk.Label(info, text=f"方法: {prompt['method']}", anchor='w').pack(fill='x')
        tk.Label(info, text=f"实验编号: {prompt['experiment_code']}", anchor='w', fg="#2563eb").pack(fill='x')
        co = prompt['configure_order']
        tk.Label(info, text=f"标准溶液: {co or '(无)'}", anchor='w',
                 fg=("#16a34a" if co else "#999")).pack(fill='x')

        fields_frame = tk.LabelFrame(win, text="结果数值（请逐项核对填写）", padx=8, pady=6)
        fields_frame.pack(fill='both', expand=True, padx=8, pady=6)
        entries = {}
        cols = prompt['dynamic_columns'] or []
        if not cols:
            tk.Label(fields_frame, text="(无动态字段配置)").pack()
        for col in cols:
            code = col.get("columeCode", "")
            name = col.get("columeName", code)
            r = tk.Frame(fields_frame)
            r.pack(fill='x', pady=2)
            tk.Label(r, text=name, width=22, anchor='w').pack(side='left')
            e = tk.Entry(r, width=24)
            e.insert(0, str(col.get("defaultVal", "")))
            e.pack(side='left', padx=4)
            entries[code] = e

        btns = tk.Frame(win, pady=8)
        btns.pack(fill='x')

        def on_submit():
            self._confirm_result = {code: e.get().strip() for code, e in entries.items()}
            self._confirm_done.set()
            win.destroy()

        def on_skip():
            self._confirm_result = "skip"
            self._confirm_done.set()
            win.destroy()

        def on_abort():
            self._confirm_result = "abort"
            self._confirm_done.set()
            win.destroy()

        tk.Button(btns, text="确认提交", command=on_submit, width=12, bg="#dbeafe").pack(side='left', padx=6)
        tk.Button(btns, text="跳过本行", command=on_skip, width=10).pack(side='left', padx=6)
        tk.Button(btns, text="中止序列", command=on_abort, width=10, bg="#fee2e2").pack(side='left', padx=6)
        win.protocol("WM_DELETE_WINDOW", on_skip)
        self._confirm_win = win

    def toggle_pause(self):
        if not self._running:
            return
        if self._pause.is_set():
            self._pause.clear()
            self.pause_btn.configure(text="▶ 继续")
            self._log("已暂停(行边界生效)")
            self.status_var.set("已暂停")
        else:
            self._pause.set()
            self.pause_btn.configure(text="⏸ 暂停")
            self._log("已继续")
            self.status_var.set("运行中...")

    def skip_current(self):
        if not self._running:
            return
        if not self._confirm_done.is_set():
            self._confirm_result = "skip"
            self._confirm_done.set()
        self._log("请求跳过当前行(确认面板或下一行生效)")

    def abort_run(self):
        if not self._running:
            return
        if messagebox.askyesno("确认", "确定中止序列运行吗？"):
            self._stop.set()
            self._pause.set()
            if not self._confirm_done.is_set():
                self._confirm_result = "abort"
                self._confirm_done.set()
            self._log("请求中止...")

    def _show_login_dialog(self):
        """主线程：弹出登录对话框(选用户名+验证码)。返回是否登录成功"""
        users = list(self.login_system.users.keys())
        if not users:
            messagebox.showerror("错误", "没有可用用户(请先在主程序登录保存用户)")
            return False

        MUTED = "#888888"  # 仅 cap_label(tk) 用

        win = ttkb.Toplevel(self.root)
        win.title("登录 LIMS")
        win.geometry("400x360")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()

        body = ttkb.Frame(win)
        body.pack(fill='both', expand=True, padx=30, pady=(28, 10))

        user_var = tk.StringVar(value=users[0])
        ttkb.Label(body, text="用户").pack(anchor='w')
        ttkb.Combobox(body, textvariable=user_var, values=users, state='readonly').pack(fill='x', pady=(2, 10))

        cap_entry_var = tk.StringVar()
        ttkb.Label(body, text="验证码").pack(anchor='w')
        cap_row = ttkb.Frame(body)
        cap_row.pack(fill='x', pady=(2, 2))

        cap_photo = {'img': None}
        # 先 pack 验证码图框固定右侧槽位，输入框再填左侧，确保图片有位置显示
        # 注意：不设 width/height，否则 tk.Label 会把图片裁进固定小框导致看不清
        cap_label = tk.Label(cap_row, text="加载中...", bg="#ffffff", fg=MUTED,
                             font=("Segoe UI", 8), relief="solid", bd=1, cursor="hand2")
        cap_label.pack(side='right')
        cap_entry = ttkb.Entry(cap_row, textvariable=cap_entry_var)
        cap_entry.pack(side='left', fill='x', expand=True)
        ttkb.Label(body, text="点击图片可刷新", bootstyle="secondary",
                   font=("Segoe UI", 8)).pack(anchor='w', pady=(2, 0))

        def fetch_captcha():
            img = self.login_system.get_captcha_image()
            if img is None:
                cap_label.configure(image="", text="获取失败 点击重试", fg="#dc2626")
                return
            # 等比缩放到与输入框等高(标签无 width/height 限制不裁剪)
            w, h = img.size
            target_h = max(20, cap_entry.winfo_reqheight())
            img = img.resize((max(1, round(w * target_h / h)), target_h), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            cap_photo['img'] = photo
            cap_label.image = photo  # 防止图片被回收
            cap_label.configure(image=photo, text="")

        def on_refresh():
            cap_label.configure(image="", text="刷新中...", fg=MUTED)
            win.after(10, fetch_captcha)

        def on_login():
            ok, msg = self.login_system.login_with_user(user_var.get(), cap_entry_var.get().strip())
            if ok:
                self.logged_in = True
                self._refresh_user_menu()
                self.status_var.set(f"已登录: {msg}")
                win.destroy()
            else:
                messagebox.showwarning("登录失败", msg, parent=win)
                cap_entry_var.set("")
                on_refresh()

        cap_label.bind("<Button-1>", lambda e: on_refresh())
        ttkb.Button(body, text="登 录", command=on_login, width=14, bootstyle="primary").pack(pady=(16, 14))
        win.bind('<Return>', lambda e: on_login())

        win.after(50, fetch_captcha)
        win.update_idletasks()
        px = self.root.winfo_x() + (self.root.winfo_width() - 400) // 2
        py = self.root.winfo_y() + (self.root.winfo_height() - 360) // 2
        win.geometry(f"+{max(px, 0)}+{max(py, 0)}")
        win.wait_window(win)
        return self.logged_in

    def _refresh_user_menu(self):
        """根据登录状态刷新用户菜单按钮(文案 + 菜单首项 登录/切换用户)"""
        if self.logged_in and self.login_system.current_user:
            disp = self.login_system.users.get(self.login_system.current_user, {}).get(
                'display_name', self.login_system.current_user)
            self.user_menu_btn.configure(text=f"👤 {disp} ▾", fg="#16a34a")
            self.user_menu.entryconfigure(0, label="切换用户")
        else:
            self.user_menu_btn.configure(text="👤 未登录 ▾", fg="#dc2626")
            self.user_menu.entryconfigure(0, label="登录")

    def _open_user_management(self):
        """打开用户管理对话框(复用 login.CompactLoginApp 的增/删/查用户 UI)"""
        from login import CompactLoginApp

        def _status_configure(text=None, foreground=None):
            # chamber 在删除/改密当前用户时调 status_label.configure("未登录", red)
            if text and "未登录" in text:
                self.logged_in = False
                self._refresh_user_menu()

        # ponytail: self=host 复用 chamber 的 show_new_user_dialog/delete_user/center_window，零重写
        host = types.SimpleNamespace()
        host.root = self.root
        host.login_system = self.login_system
        host.status_label = types.SimpleNamespace(configure=_status_configure)
        host.update_user_list = lambda: None  # 登录对话框每次重开即刷新用户列表
        host.find_icon_file = lambda name: name if os.path.exists(name) else None
        host.center_window = types.MethodType(CompactLoginApp.center_window, host)
        host.delete_user = types.MethodType(CompactLoginApp.delete_user, host)
        CompactLoginApp.show_new_user_dialog(host)


def _selfcheck():
    """_plan_submission_batches 自检：方法/样品/max_select 三场景。无依赖，纯逻辑。"""
    def item(sc, pn, mid):
        return {"sample_code": sc, "projectName": pn, "switch_mid": mid,
                "project": {"projectId": f"{sc}-{pn}", "projectName": pn}}
    items = [item("S1", "苯", "4678"), item("S2", "苯", "4678"),
             item("S1", "甲苯、二甲苯及乙苯总和", "4679"),
             item("S2", "甲苯、二甲苯及乙苯总和", "4679")]
    rules_m = [{"project": "苯", "input_method": "方法"},
               {"project": "甲苯、二甲苯及乙苯总和", "input_method": "方法"}]

    # 方法模式：苯(2样品合并)在前、总和(2样品合并)在后，各1批，force_new=False
    p = _plan_submission_batches(rules_m, items, 30)
    assert [(b["switch_mid"], len(b["items"]), b["force_new"]) for b in p] \
        == [("4678", 2, False), ("4679", 2, False)], p

    # 样品模式：每样品各1批，苯先于总和
    rules_s = [{"project": "苯", "input_method": "样品"},
               {"project": "甲苯、二甲苯及乙苯总和", "input_method": "样品"}]
    p2 = _plan_submission_batches(rules_s, items, 30)
    assert [(b["switch_mid"], b["items"][0]["sample_code"]) for b in p2] \
        == [("4678", "S1"), ("4678", "S2"), ("4679", "S1"), ("4679", "S2")], p2

    # 方法模式 max_select=1：苯拆2批，两批 force_new=True
    p3 = _plan_submission_batches(rules_m, items, 1)
    assert [b["force_new"] for b in p3 if b["switch_mid"] == "4678"] == [True, True], p3

    # 称样日期拆批(方法模式)：同 switch_mid 跨日 → 拆 2 批 force_new=True；同日仍合 1 批
    di = [item("S1", "苯", "4678"), item("S2", "苯", "4678")]
    di[0]["wdate"] = "2026-07-28"
    di[1]["wdate"] = "2026-07-29"
    pd = _plan_submission_batches([{"project": "苯", "input_method": "方法"}], di, 30)
    assert len(pd) == 2 and all(b["force_new"] for b in pd), pd
    assert [b["items"][0]["sample_code"] for b in pd] == ["S1", "S2"], pd
    di[1]["wdate"] = "2026-07-28"  # 改同日 → 合 1 批
    pd2 = _plan_submission_batches([{"project": "苯", "input_method": "方法"}], di, 30)
    assert len(pd2) == 1 and not pd2[0]["force_new"], pd2

    # _parallel_indices：同 projectId 多组分(总和 5组分×2平行) 按 serialNumber 正确归平行
    recs = [{"projectId": "P1", "serialNumber": s} for s in (1, 2, 1, 2, 1, 2, 1, 2, 1, 2)]
    po, n = _parallel_indices(recs)
    assert n == 2 and po == {0: 0, 1: 1, 2: 0, 3: 1, 4: 0, 5: 1, 6: 0, 7: 1, 8: 0, 9: 1}, (n, po)
    # 无 serialNumber 回退：按 projectId 枚举
    po2, n2 = _parallel_indices([{"projectId": "P1"}, {"projectId": "P1"}])
    assert n2 == 2 and po2 == {0: 0, 1: 1}, (n2, po2)

    # _match_filename_rule：名称切换按 project_name + PDF文件名关键字命中(desc 默认空=不限)
    fr = [{"to_id": "4481", "filename": "K", "project_name": "DEHP"},
          {"to_id": "4482", "filename": "K", "project_name": "DNOP"}]
    # 命中：DNOP + 文件名含 K -> 4482（大小写不敏感）
    assert _match_filename_rule(fr, "DNOP", ["D:/sp/K-001.pdf"]) == "4482"
    # 不命中：DBP 无名称规则 -> ''（回落 default_rules）
    assert _match_filename_rule(fr, "DBP", ["D:/sp/K-001.pdf"]) == ""
    # 不命中：DNOP 但文件名不含 K -> ''
    assert _match_filename_rule(fr, "DNOP", ["D:/sp/001.pdf"]) == ""
    # project_name 短关键字包含于 LIMS 长项目名(大小写不敏感) -> 路由到各自 to_id
    assert _match_filename_rule(fr, "3种邻苯二甲酸酯类化合物（DBP、BBP、DEHP）总和", ["K-001.pdf"]) == "4481"
    assert _match_filename_rule(fr, "3种邻苯二甲酸酯类化合物（DNOP、DINP、DIDP）总和", ["K-001.pdf"]) == "4482"

    # 名称切换 + 试样描述(desc)：文件名与描述均非空时需同时命中(AND)；任一为空=不限
    mr = [{"to_id": "5501", "filename": "K", "desc": "苯,甲苯"}]
    # 文件名命中但描述不含任一关键字 -> 不命中(两条件都填需同时满足)
    assert _match_filename_rule(mr, "", ["K-001.pdf"], "水溶液") == ""
    # 文件名 + 描述都命中 -> 5501
    assert _match_filename_rule(mr, "", ["K-001.pdf"], "苯溶液") == "5501"
    # 描述第二关键字"甲苯"命中 -> 5501
    assert _match_filename_rule(mr, "", ["K-001.pdf"], "甲苯") == "5501"
    # 只填描述(无文件名条件) -> 命中
    assert _match_filename_rule([{"to_id": "5502", "desc": "固体"}], "", [], "固体颗粒") == "5502"
    # 规则要求描述关键字但样品描述为空 -> 不命中
    assert _match_filename_rule([{"to_id": "5502", "desc": "固体"}], "", [], "") == ""
    # 大小写不敏感
    assert _match_filename_rule([{"to_id": "9", "desc": "Solid"}], "", [], "SOLID sample") == "9"
    print("selfcheck OK")


if __name__ == "__main__":
    if "--selfcheck" in sys.argv:
        _selfcheck()
        sys.exit(0)
    root = ttkb.Window(themename="sandstone-light")
    app = SequenceMaster(root)
    root.mainloop()