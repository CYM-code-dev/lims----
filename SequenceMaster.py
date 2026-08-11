import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import ttkbootstrap as ttkb  # 档1: 现代主题(sandstone-light)，ttk 控件自动套用
import os
import re
import sys
import time
import csv
import json
import fnmatch
import glob
import threading
import queue
import types
import random
import requests
import yaml
import openpyxl
from datetime import datetime, date, timedelta
from PIL import Image, ImageTk

sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from login import MultiUserLoginSystem
from method_file import load_method
from detection_entry_api import DetectionAPI, build_grouped_experiment_data, _Box, _norm_cn


# ==================== 节假日/工作日（法定假日+调休，数据源：标准品服务器 10.1.93.25:5000）====================
_HOLIDAY_HOST = "http://10.1.93.25:5000"
_holiday_year_cache = {}  # year -> {YYYY-MM-DD: bool}(true=放假/false=补班) 或 None


def _holiday_mapping(year):
    """取年度节假日映射（进程内按年缓存；服务器不可达返回 None，调用方回落周一~周五）。"""
    if year in _holiday_year_cache:
        return _holiday_year_cache[year]
    mapping = None
    try:
        r = requests.get(f"{_HOLIDAY_HOST}/holidays/{year}.json", timeout=5)
        if r.status_code == 200:
            data = r.json()
            if isinstance(data, dict):
                mapping = data
    except Exception:
        mapping = None
    _holiday_year_cache[year] = mapping
    return mapping


def _is_workday_holiday(d):
    """d: datetime.date。命中映射(true=放假→非工作日/false=补班→工作日)，否则周一~周五。"""
    mapping = _holiday_mapping(d.year)
    if mapping:
        v = mapping.get(d.strftime("%Y-%m-%d"))
        if v is True:
            return False
        if v is False:
            return True
    return d.weekday() < 5


def _prev_workday():
    """上一个工作日（考虑法定假日/调休）：从昨天起回退到首个工作日；上限40天防异常。"""
    cur = date.today() - timedelta(days=1)
    for _ in range(40):
        if _is_workday_holiday(cur):
            return cur
        cur -= timedelta(days=1)
    return cur


_ACCEPT_DATE_KEYS = ("acceptTime", "acceptDate", "checkInTime", "checkInDate",
                     "receiveTime", "receiveDate", "sampleAcceptDate", "registerDate")


def _accept_date_of(raw):
    """从 LIMS 原始 sample_data(_raw) 提取受理日期(date)。返回 (date, 命中字段名) 或 (None, "")。
    服务端005校验按日期比较(前端 startTime='受理日 00:00:00' 即通过)。解析 'YYYY-MM-DD...' 取前10位。"""
    if not isinstance(raw, dict):
        return None, ""
    for k in _ACCEPT_DATE_KEYS:
        v = raw.get(k)
        if not v:
            continue
        try:
            return date.fromisoformat(str(v).strip()[:10]), k
        except ValueError:
            continue
    return None, ""


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
            # 单元格拖动选择 - 矩形区域(支持横向多列)
            min_row, max_row = min(start_row, row), max(start_row, row)
            min_col, max_col = min(start_col, col), max(start_col, col)

            # 如果不是Ctrl或Shift操作，清除之前的选择
            if not ctrl_pressed and not shift_pressed:
                self.selected_cells.clear()

            # 添加矩形选择
            for r in range(min_row, max_row + 1):
                for c in range(min_col, max_col + 1):
                    self.selected_cells.add((r, c))

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


# 报告解析结果列定位 & 样品报告 PDF 挑选（批量浓度回填用）
_REPORT_RESULT_KW = ('计算值', '报告值', '测定值', '浓度', '结果值', '含量')
# 多报告按方法/项目类别关键词挑 PDF（同 detection_entry_main._CATEGORY_HINTS）
_REPORT_CATEGORY_HINTS = (
    ('PAE', ('PAE', '邻苯')), ('PAHS', ('PAHS', '多环', 'PAH')),
    ('PCN', ('PCN', '氯萘')), ('BXW', ('BXW', '苯系物')),
    ('氯苯', ('氯苯',)), ('SCCP', ('SCCP',)), ('SVHC', ('SVHC',)),
    ('AZO', ('AZO', '偶氮')), ('OT', ('OT',)),
    ('ICP', ('ICP', '元素', '重金属')),
)


def _find_result_column(dynamic_columns, pdf_headers=()):
    """挑数据采集结果列(浓度/计算值/报告值...)：
    1) equipRelativeTitle 精确命中 PDF 报告表头(原系统机制，浓度/校准浓度等都能对上)；
    2) 兜底列名关键词；排除合并列(isColumnMerge)/组分列(isMutiPolyColume)；
    多命中取 columeOrder 最小。返回 columeCode 或 None。"""
    # ponytail: 与 detection_entry_main._find_result_column 重复(~20行)；不抽公共件以避免改动已验证的单样品代码。
    def _is_merge(c):
        try:
            return int(c.get('isColumnMerge') or 0) == 1
        except (TypeError, ValueError):
            return False

    def _is_comp(c):
        try:
            return int(c.get('isMutiPolyColume') or 0) == 1
        except (TypeError, ValueError):
            return False

    cols = [c for c in (dynamic_columns or [])
            if isinstance(c, dict) and not _is_merge(c) and not _is_comp(c)]
    if pdf_headers:
        cands = [c for c in cols if (c.get('equipRelativeTitle') or '').strip() in pdf_headers]
        if cands:
            return min(cands, key=lambda c: c.get('columeOrder', 9999)).get('columeCode')
    cands = [c for c in cols if any(k in (c.get('columeName') or '') for k in _REPORT_RESULT_KW)]
    if not cands:
        return None
    return min(cands, key=lambda c: c.get('columeOrder', 9999)).get('columeCode')


def _match_sample_report_pdfs(spectrum_path, sample_code, method_hint=""):
    """在谱图目录中匹配该样品的报告 PDF，返回 (normal_pdfs, dil_pdfs)：按是否含 -NNX 稀释后缀
    分为正常/稀释；各自按方法/项目类别关键词过滤后按文件名序保留全部——同类多份即 A/B 平行
    (GCMS 每平行一份 PDF)，由调用方逐份解析按平行槽回填。ICP-MS 按报告内 Sample Name 认领时
    normal 为单元素。对应无则空列表。"""
    from report_parser import (_dilution_factor, extract_content_sample_ids,
                               extract_icp_ms_sample_name, _split_content_dilution)
    sc = (sample_code or '').lower()
    if not sc or not spectrum_path or not os.path.isdir(spectrum_path):
        return [], []
    # ICP-MS：按报告内 Sample Name 认领(优先于文件名——文件名可能与报告样品号不一致)
    # ponytail: ICP-OES 稀释报告成对置于 稀释/ 子目录(正常+稀释各一份)；一并扫描，无该目录则仅顶层。
    def _list_pdfs(d):
        try:
            return [os.path.join(d, fn) for fn in os.listdir(d) if fn.lower().endswith('.pdf')]
        except Exception:
            return []
    _pdf_dirs = [spectrum_path, os.path.join(spectrum_path, '稀释')]
    _all_pdfs = [p for d in _pdf_dirs if os.path.isdir(d) for p in _list_pdfs(d)]
    # ICP-MS：按报告内 Sample Name 认领(优先于文件名——文件名可能与报告样品号不一致)。
    # 稀释报告 Sample Name 带 -NNX(如 TS…001-10X)，去后缀后匹配样品号；
    # 正常+稀释(-NNX)成对，按 _dilution_factor 分流(供下游超线性换源 + 稀释列填倍数)。
    _icp_hits = []
    for p in _all_pdfs:
        _sn = extract_icp_ms_sample_name(p)
        if not _sn:
            continue
        _base, _ = _split_content_dilution(_sn)
        if _base == sample_code or sample_code.startswith(_base) or _base.startswith(sample_code):
            _icp_hits.append(p)
    if _icp_hits:
        _icp_hits = sorted(set(_icp_hits))
        _normal = [p for p in _icp_hits if _dilution_factor(p) == 1.0]
        _dil = [p for p in _icp_hits if _dilution_factor(p) != 1.0]
        if not _normal:      # 全 -NNX(无独立正常报告)：整体作正常源
            _normal = _icp_hits
            _dil = []
        return _normal, _dil
    pdfs = [p for p in _all_pdfs if os.path.basename(p).lower().startswith(sc)]
    # ponytail: 报验编号命名(无小号，如 TN26080187.pdf)的正常报告并入——
    # 不再仅兜底(if not pdfs)，否则只匹中稀释报告会令 normal_pdfs 为空触发"全-NNX→正常"误并。
    # 词干须精确等同报验编号(非前缀)：避免误并同报验编号其它小号样品(如 TS...893012)的报告。
    stripped = _strip_parallel_suffix(sample_code).lower()
    if stripped and stripped != sc:
        pdfs += [p for p in _all_pdfs
                 if p not in pdfs
                 and os.path.splitext(os.path.basename(p))[0].lower() == stripped]
    if not pdfs:  # 文件名前缀全 miss -> 按 PDF 内容 `样品 :` 字段关联(支持 文件名≠内容id 的 PAHS 报告)
        sc_stripped = _strip_parallel_suffix(sample_code).strip()
        for p in _all_pdfs:
            for cid in extract_content_sample_ids(p):
                if cid == sample_code or cid.startswith(sc_stripped) or sc_stripped.startswith(cid):
                    pdfs.append(p)
                    break
    if not pdfs:
        return [], []
    pdfs = sorted(set(pdfs))   # 文件名序：A 平行在 B 前(同基编号仅平行字母不同)
    normal_pdfs = [p for p in pdfs if _dilution_factor(p) == 1.0]
    dil_pdfs = [p for p in pdfs if _dilution_factor(p) != 1.0]
    if not normal_pdfs:   # 全部带 -NNX(GCMS 平行稀释: TN...KA-10X/TN...KB-10X)：报告即稀释源，无独立稀释 PDF，
        normal_pdfs = pdfs   # 各段按报告「名称」字段的 -NNX 自带稀释倍数；清空 dil 避免误把首份当稀释源
        dil_pdfs = []
    text = (method_hint or '').upper()
    hint = ''
    for cat, keys in _REPORT_CATEGORY_HINTS:
        if any(k.upper() in text for k in keys):
            hint = cat
            break

    def _by_hint(cands):
        if hint and len(cands) > 1:
            matched = [p for p in cands if hint.lower() in os.path.basename(p).lower()]
            if matched:
                return matched
        return cands

    return _by_hint(normal_pdfs), _by_hint(dil_pdfs)


def _pick_sample_report_pdf(spectrum_path, sample_code, method_hint=""):
    """在谱图目录中挑该样品的报告 PDF，返回 (normal_pdf, diluted_pdf)：取 _match_sample_report_pdfs
    的首个正常/稀释 PDF(单份用法，如称样量 PDF 路径)。对应无则该位为 None。"""
    normal_pdfs, dil_pdfs = _match_sample_report_pdfs(spectrum_path, sample_code, method_hint)
    return (normal_pdfs[0] if normal_pdfs else None), (dil_pdfs[0] if dil_pdfs else None)


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


def _random_masses(recorded, n_par, prule, wp, mdp):
    """random 称样量：记录值不少于本样品平行数(n_par)则用记录(按 prule 换算或按 mdp 格式化)，
    否则随机生成。平行数取本样品的，混批下不因他样多平行而丢弃本样记录值。"""
    if recorded and len(recorded) >= n_par:
        if prule:
            return [_apply_processing(recorded[i], prule, wp) for i in range(n_par)]
        return [f"{float(recorded[i]):.{mdp}f}" for i in range(n_par)]
    return [_gen_random_mass(wp) for _ in range(n_par)]


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


def _mass_field_by_project(records, pid_to_sample, pmasses_by_sample,
                           pid_filter=None, marker_masses_by_sample=None):
    """跨样品合并提交时按 projectId -> 样品 -> 该样品称样量[平行索引] 构造字段列表。
    不同样品用各自的称样量；同样品同平行共用一值。pid_to_sample: {projectId_str: sample_code}；
    pmasses_by_sample: {sample_code: [mass_str,...]}(按平行序)。
    pid_filter: {projectId_str: 谱图filter}；filter 为标记字母(如 'M')时，该记录改取
    marker_masses_by_sample[样品][标记]，供「5mm以内」等标记项目用标记(基体加标)称样量。"""
    parallel_of, _ = _parallel_indices(records)
    out = []
    for g, r in enumerate(records or []):
        pid = str(r.get("projectId"))
        sc = pid_to_sample.get(pid, "")
        flt = (pid_filter or {}).get(pid, "")
        mm = (marker_masses_by_sample or {}).get(sc) or {}
        masses = mm.get(flt) if (flt and mm.get(flt)) else (pmasses_by_sample.get(sc) or [])
        par = parallel_of.get(g, 0)
        val = masses[par] if par < len(masses) else (masses[-1] if masses else "")
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


def _is_sum_project(project_name):
    """是否"总和"项目：项目名含「总和」或「之和」(覆盖 甲苯..总和/四项之和/AfPS PAK 4项之和/18种多环芳烃总和 等)。"""
    pn = (project_name or "")
    return "总和" in pn or "之和" in pn




def _std_loose_match(std, target):
    """标准号宽松匹配：忽略大小写与所有空白(LIMS 标准号空格不一致，如 'IEC62321-4' vs 'IEC 62321-4')，
    == 或互为前缀(容忍'附录X'后缀)。"""
    s = "".join((std or "").split()).lower()
    t = "".join((target or "").split()).lower()
    if not s or not t:
        return False
    if s == t:
        return True
    if len(t) >= 6 and len(s) >= 6:
        return s.startswith(t) or t.startswith(s)
    return False


def _exclusion_match(project, rule):
    """单条排除规则是否命中某 LIMS 项目(双维度 AND)。
    project 空=任意项目，否则按 _project_match(精确/通配*)；method 空=任意方法，否则按 _std_loose_match 匹配 standardNo。
    各非空条件均需满足；与 query_rules 选项目的口径一致(精确/通配，非关键字包含)。"""
    rp = str(rule.get("project") or "").strip()
    if rp and not _project_match(project.get("projectName", ""), rp):
        return False
    rm = str(rule.get("method") or "").strip()
    if rm and not _std_loose_match(project.get("standardNo", ""), rm):
        return False
    return True


def _filter_pdfs_by_rule(pdf_paths, project_name, filter_rules):
    """按命中的 spectrum_filter_rules 过滤本项目的样品谱图PDF(大小写不敏感,匹配文件名子串)。
    filter_rules: [{project, filter}, ...]，取首条 project 命中规则的 filter；common_pdfs 不走这里。
    filter: 留空=不过滤；"kw"=只保留含 kw 的；"!kw"=排除含 kw 的。"""
    flt = ""
    for r in filter_rules or []:
        if _project_match(project_name, str(r.get("project") or "")):
            flt = str(r.get("filter") or "").strip()
            break
    paths = list(pdf_paths or [])
    if not flt:
        return paths
    exclude = flt.startswith("!")
    kw = flt[1:].strip().lower() if exclude else flt.lower()
    if not kw:
        return paths
    return [p for p in paths if (kw in os.path.basename(p).lower()) != exclude]


_MAX_BATCH_ITEMS = 500  # 单批录入条数硬上限：超限按 max_items 切片，每批独立实验编号(force_new)


def _plan_submission_batches(query_rules, items, max_items=_MAX_BATCH_ITEMS):
    """按 query_rules 顺序规划提交批次（纯函数，可单测）。
    items: 每项为 dict，需含 switch_mid / sample_code / projectName / wdate(称样日期,可空)。
    返回 [{"switch_mid","wdate","items","force_new","_equipment"}, ...]，顺序 = 规则顺序，同规则内按 switch_mid、
    再按条件设备(可选)、再按称样日期、再按样品。不同设备需拆独立批。
    同 switch_mid 内称样日期不同(跨天)拆独立批(各自实验编号)；同日/无日期仍合并。
    input_method=方法：同(switch_mid,设备,日期)的样品合并，每条规则各自的 max_select 超限切片(多片 force_new=True)；
    input_method=样品：每样品各一片。无 query_rules 退化为单条空规则(全中,方法)。
    多个不同 method(如 XRF 多元素方法)：每条规则按项目 _qr_idx 认领各自方法，一方法一实验、不重复。
    单批录入条数(len items)超 max_items 再切片(每片 force_new=True)，避免单实验编号录入过大。"""
    plan = []
    rules = query_rules or [{"project": "", "input_method": "方法"}]
    # 多个不同方法(如 XRF 多元素方法)→ 每条规则按 _qr_idx 认领各自方法的项目；否则按 project 拆分
    _methods = {str(r.get("method") or "").strip() for r in rules if str(r.get("method") or "").strip()}
    multi_method = len(_methods) > 1
    remaining = list(items)  # 已归批的不再参与后续规则，避免多条规则重复提交同一些项目
    for i, rule in enumerate(rules):
        if not remaining:
            break
        mode = str(rule.get("input_method") or "方法").strip()
        ms = str(rule.get("max_select") or "").strip()
        rule_max = int(ms) if ms.isdigit() else 0
        if multi_method:
            rule_items = [it for it in remaining if it.get("_qr_idx") == i]
        else:
            rp = str(rule.get("project") or "").strip()
            rule_items = [it for it in remaining if _project_match(it.get("projectName", ""), rp)]
        if not rule_items:
            continue
        remaining = [it for it in remaining if id(it) not in {id(x) for x in rule_items}]
        groups = {}  # switch_mid -> [items]（保序）
        for it in rule_items:
            groups.setdefault(it.get("switch_mid", ""), []).append(it)
        for mid, g_items in groups.items():
            # 条件设备二次分组：不同设备需拆独立批（分批录入）
            by_eq = {}  # equipment -> [items]（保序）
            for it in g_items:
                by_eq.setdefault(it.get("_equipment") or "", []).append(it)
            for eq_key, eq_items in by_eq.items():
                # 条件实验过程三次分组：不同 lab_proc 需拆独立批(各自 experimentProcess)
                by_lp = {}  # lab_proc -> [items]（保序）
                for it in eq_items:
                    by_lp.setdefault(it.get("_lab_proc") or "", []).append(it)
                for lp_key, lp_items in by_lp.items():
                    # 称样日期分组
                    by_date = {}  # wdate -> [items]（保序）
                    for it in lp_items:
                        by_date.setdefault(it.get("wdate", ""), []).append(it)
                    sub_batches = []  # [(wdate, items, [sample 切片])]
                    for d, d_items in by_date.items():
                        samples = list(dict.fromkeys(it.get("sample_code", "") for it in d_items))
                        if mode == "样品":
                            slices = [[sc] for sc in samples]
                        elif rule_max > 0 and len(samples) > rule_max:
                            slices = [samples[i:i + rule_max] for i in range(0, len(samples), rule_max)]
                        else:
                            slices = [samples]
                        sub_batches.append((d, d_items, slices))
                    # 多方法(XRF 各元素)每批是独立实验，必须各出新编号(否则同 actual_method_name 时
                    # force_new=False 会复用上一批缓存编号→撞号合并)；同方法则仅多片/跨日时才出新编号
                    force_new = multi_method or sum(len(sl) for _, _, sl in sub_batches) > 1
                    for d, d_items, slices in sub_batches:
                        for sb in slices:
                            sb_set = set(sb)
                            batch = [it for it in d_items if it.get("sample_code", "") in sb_set]
                            # 单批录入条数硬上限：超 max_items 再切片，每片独立实验编号
                            chunks = ([batch[i:i + max_items] for i in range(0, len(batch), max_items)]
                                      if len(batch) > max_items else [batch])
                            for ch in chunks:
                                plan.append({
                                    "switch_mid": mid,
                                    "_equipment": eq_key,
                                    "_lab_proc": lp_key,
                                    "wdate": d,
                                    "items": ch,
                                    "force_new": force_new or len(chunks) > 1,
                                })
    return plan


def _kw_match(keywords, target):
    """关键字匹配（desc/filename 等）：逗号=OR(任一)，分号=AND(均需命中)。
    如 '纸,布;涂层' = (纸 OR 布) AND 涂层。keywords 空=不限(命中)。大小写不敏感，子串包含。"""
    kk = (keywords or "").strip()
    if not kk:
        return True
    t = (target or "").lower()
    for g in [x for x in kk.replace("；", ";").split(";") if x.strip()]:
        kws = [k for k in g.replace("，", ",").split(",") if k.strip()]
        if kws and not any(k.lower() in t for k in kws):
            return False
    return True


_EQUIP_SEP_RE = re.compile(r"[;,，；]")


def _split_device_codes(field):
    """设备编号按 ; ; ， ， 任一分隔拆成多台(去空白/空串，保序)。条件/指定设备多台编号通用。"""
    return [c.strip() for c in _EQUIP_SEP_RE.split(field or "") if c.strip()]


def _match_switch_rule(switch_rules, project_name, pdf_paths, desc=""):
    """统一规则匹配（纯函数，可单测）。
    规则各非空条件均需满足(AND): project_name 精确/wildcard匹配(空=任意项目)、
    filename 关键字出现在某谱图PDF文件名(空=任意文件)、desc 关键字(逗号OR/分号AND)
    出现在样品"试样描述"(空=任意描述)。多规则命中取首条，均不命中返回 ''。
    三项匹配均大小写不敏感。"""
    bases = [os.path.basename(p).lower() for p in (pdf_paths or []) if p]
    pn = (project_name or "").strip()
    for r in switch_rules or []:
        rp = str(r.get("project_name") or "").strip()
        if rp and not _project_match(pn, rp):
            continue
        fk = str(r.get("filename") or "").strip().lower()
        if fk and not any(fk in b for b in bases):
            continue
        if not _kw_match(r.get("desc"), desc):
            continue
        return str(r.get("to_id") or "").strip()
    return ""


def _match_conditional_rule(rules, project_name, pdf_paths, desc=""):
    """条件规则匹配（纯函数）：project_name/filename/desc 三条件 AND，空=不限；多规则命中取首条。
    desc 关键字：逗号=OR(任一)，分号=AND(均需命中)。返回命中的规则 dict，均不命中返回 None。
    供条件设备/条件实验过程等共用。"""
    bases = [os.path.basename(p).lower() for p in (pdf_paths or []) if p]
    pn = (project_name or "").strip()
    for r in rules or []:
        rp = str(r.get("project_name") or "").strip()
        if rp and rp.lower() not in pn.lower():
            continue
        fk = str(r.get("filename") or "").strip().lower()
        if fk and not any(fk in b for b in bases):
            continue
        if not _kw_match(r.get("desc"), desc):
            continue
        return r
    return None


def _match_equipment_rule(equipment_rules, project_name, pdf_paths, desc=""):
    """条件设备：返回命中规则的 device_number，均不命中返回 ''。"""
    m = _match_conditional_rule(equipment_rules, project_name, pdf_paths, desc)
    return str((m or {}).get("device_number") or "").strip()


def _match_lab_proc_rule(lab_proc_rules, project_name, pdf_paths, desc=""):
    """条件实验过程：返回命中规则的 lab_proc(覆盖 experimentProcess)，均不命中返回 ''。"""
    m = _match_conditional_rule(lab_proc_rules, project_name, pdf_paths, desc)
    return str((m or {}).get("lab_proc") or "").strip()


def _rule_desc_match(switch_rules, desc):
    """混目录分流用：样品"试样描述"是否命中本方法任一 switch_rules 的 desc 关键字。
    desc 逗号分隔、任一命中、大小写不敏感(与 _match_switch_rule 的 desc 口径一致)；
    任一规则未设 desc 视为"不限描述"(命中)。全部不命中返回 False。"""
    d = (desc or "").strip().lower()
    for r in switch_rules or []:
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
    """去掉称样编号里的平行小号(001)，用作谱图前缀匹配。
    约定编号 = 字母前缀 + 8位流水 + 3位平行小号(001/002…) + 可选字母(A/B)，
    如 TN26070466001→TN26070466、TN26070474001A→TN26070474A；
    不符该结构(无小号)原样返回。提交 LIMS 仍用带小号的原始编号。"""
    m = _PARALLEL_SUFFIX_RE.match((code or "").strip())
    return (m.group(1) + m.group(2)) if m else (code or "")


class _ParallelWMap(dict):
    """称量记录 dict：.get 按样品编号容错查找。
    称样记录的 A/B 平行常省略 3 位小号(写 报验编号+字母，如 TN26070729A/B)，合并后 key=报验编号
    (TN26070729)；而实验里样品编号是 报验编号+001(TN26070729001)，直接 get 会查不到。
    另有种样记录用单字母后缀(如 TN26080226K，材质标记未配对、未合并为报验编号)，key 保留
    TN26080226K，按 LIMS 全码 TN26080226001 查会 miss → desc/称样量/称样时间丢失。
    故 get 依次：精确→去小号按报验编号→按报验编号 base 兜底匹配字母后缀键。仅覆盖 get——
    items/keys/迭代仍只含真实 key，不会冒出重复样品。"""

    def get(self, code, default=None):
        v = super().get(code)
        if v is not None:
            return v
        s = _strip_parallel_suffix(code)
        if s != code:
            v = super().get(s)
            if v is not None:
                return v
        # 单字母后缀称样记录(如 TN…K)按报验编号 base 兜底匹配
        for k, val in super().items():
            b, letter = _base_and_letter(k)
            if letter and b == s:
                return val
        return default


_PARALLEL_LETTER_RE = re.compile(r'^(\D*\d+)([A-Za-z]+)$')


def _base_and_letter(code):
    """拆出 (去尾字母的基编号, 尾字母)；无尾字母则 letter=''。
    TN26070591001A→('TN26070591001','A')、TN26070591001M→('TN26070591001','M')、TN26070591001→('TN26070591001','')。"""
    m = _PARALLEL_LETTER_RE.match((code or "").strip())
    if not m:
        return (code or ""), ""
    return m.group(1), m.group(2)


def _code_belongs_sample(any_code, code, key):
    """any_code(谱图文件名stem 或 PDF内容样品号) 是否归属样品 code(wmap合并后基编号)。
    按小号匹配，避免 startswith(key) 把同报验号的其它小号全吃进：
    - 精确: any_code == code
    - 平行字母文件: any_code 去尾字母 == code (TN…001A.pdf ↔ 样品 TN…001)
    - 简写报验编号(code 无小号=代表001)：any_code 去尾字母 == code+001
      (称样记录写报验编号 TN…20、谱图文件名带小号 TN…20001T/TS 也归属本样品，不依赖 PDF 内容)
    - 简写文件名(仅当 code 小号==001): any_code == key (TS26072277.pdf 代表 …001)"""
    any_code = (any_code or "").strip()
    if any_code == code:
        return True
    b, _ = _base_and_letter(any_code)
    if b == code and b != any_code:
        return True
    if code == key and b == code + "001":   # 称样记录用报验编号、文件名带001(+T/TS字母)
        return True
    return any_code == key and code == key + "001"


def _merge_parallel_groups(wmap, non_parallel_suffixes=()):
    """合并称样记录里的平行样：同基编号(去尾字母)归为一个样品。
    - 平行字母(A/B…)：按字母序(A→平行1、B→平行2)拼入 masses。
    - 标记后缀(non_parallel_suffixes，如 M)：不计入平行，存入 marker_masses[标记]，
      供「谱图 filter=该标记」的项目(如 5mm以内)取标记称样量——标记样在 LIMS 不是独立样品，挂在基样上。
    合并后样品编号 = 去尾字母基编号(如 TN…001A/B→TN…001)，作为 LIMS 查询键。
    普通样品(无配对、无标记)原样保留；无称样记录返回原值。"""
    if not wmap:
        return wmap
    non_par = {str(s).strip().upper() for s in (non_parallel_suffixes or []) if str(s).strip()}
    groups, order = {}, []
    for code in wmap:
        base, letter = _base_and_letter(code)
        g = groups.get(base)
        if g is None:
            g = {"par": [], "marker": {}, "time": None, "desc": ""}
            groups[base] = g
            order.append(base)
        src = wmap[code] or {}
        if g["time"] is None:
            g["time"] = src.get("time")
        if not g["desc"]:
            g["desc"] = src.get("desc") or ""
        if letter and letter.upper() in non_par:
            g["marker"].setdefault(letter.upper(), []).append(code)
        else:
            g["par"].append(code)
    out = {}
    for base in order:
        g = groups[base]
        par_codes = sorted(g["par"])  # A 在 B 前 → 平行序
        if not g["marker"] and len(par_codes) == 1:
            out[par_codes[0]] = wmap[par_codes[0]]  # 普通样品原样
            continue
        masses = []
        for c in par_codes:
            masses.extend((wmap[c] or {}).get("masses") or [])
        entry = {"masses": masses, "time": g["time"], "desc": g["desc"]}
        if any((wmap[c] or {}).get("force_parse") for c in par_codes):
            entry["force_parse"] = True  # E列「解析」标记随平行合并保留(任一平行行标即生效)
        if g["marker"]:
            entry["marker_masses"] = {}
            for mk, codes in g["marker"].items():
                mm = []
                for c in codes:
                    mm.extend((wmap[c] or {}).get("masses") or [])
                entry["marker_masses"][mk] = mm
        out[base] = entry
    return _ParallelWMap(out)


def _expand_parallel_records(experiment_config, pid_to_sample, par_by_sample, par_by_pid=None):
    """按各样品称样量平行数扩展 ocAnalysisRecordList：样品需 N 个平行而 LIMS 仅返回更少时，
    以该样品各组分的最小 serialNumber 记录为模板、serialNumber 递增复制到 N，使每个平行都有记录槽
    (实验次数=平行数；与 LIMS 前端"加平行"生成的 row_X.0001 子行同构)。
    par_by_pid: {projectId_str: 平行数} 按项目覆盖——标记项目(谱图filter=标记)用其标记称样量数，
    不跟随基样 A/B 平行数(如 5mm以内 只有 1 个 M 称样量，不应被 A/B 扩成 2 平行)。
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
        need = (par_by_pid or {}).get(pid)
        if need is None:
            sc = pid_to_sample.get(pid)
            need = par_by_sample.get(sc) if sc else 0
        min_sn, cur_max = min(_sn(r) for r in recs), max(_sn(r) for r in recs)
        do_expand = bool(need and need > 1 and cur_max < min_sn + need - 1)
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


def _pdf_parallel_count(sample_code, pdf_paths, known_suffixes):
    """从谱图文件名推断平行数：扣已知后缀(如 T/TS)后剩余单个字母=平行字母(A/B)，返回不同平行字母数(≥1)。
    random/无称样记录模式下称量记录不编码平行数，改由文件名 A/B 推断(对齐 _check_spectrum 的拆分口径)。
    known_suffixes 为小写后缀集合(如 {'','t','ts'})。"""
    base = _strip_parallel_suffix(sample_code).lower()
    head = re.compile(re.escape(base) + r"(\d{3})?(.*)$")
    suf_sorted = sorted((s for s in known_suffixes if s), key=len, reverse=True)
    letters = set()
    for p in pdf_paths or []:
        m = head.match(os.path.splitext(os.path.basename(p))[0].lower())
        if not m:
            continue
        tail = m.group(2)
        for s in suf_sorted:
            if tail.endswith(s):
                tail = tail[:-len(s)]
                break
        if len(tail) == 1 and tail.isalpha():  # 剩余单字母 = 平行字母
            letters.add(tail)
    return max(1, len(letters))


def _parse_weigh_time(v):
    """称样时间单元格 → datetime：兼容 datetime / 'YYYY/M/D' / 'YYYY-M-D'(可带时分)。解析不了返回 None。"""
    if v is None or isinstance(v, datetime):
        return v
    s = str(v).strip()
    if not s:
        return None
    for fmt in ("%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M", "%Y/%m/%d",
                "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def _read_weighing_records(path):
    """读取称量记录(xlsx 或 csv)：返回 ({样品编号: {masses:[float...], time, desc}}, err)。
    行序即平行序；time/desc 取该样品首行(称样时间/试样描述)。表头按列名定位，缺失按 A/B/C/D 兜底。"""
    try:
        if str(path).lower().endswith(".csv"):
            rows = None
            for enc in ("utf-8-sig", "gbk"):  # 中文 Excel 导出 CSV 常为 GBK
                try:
                    with open(path, encoding=enc, newline="") as f:
                        rows = [tuple(r) for r in csv.reader(f)]
                    break
                except UnicodeDecodeError:
                    continue
            if rows is None:
                return None, "读取称量记录失败: CSV 编码无法识别(尝试 utf-8/gbk)"
        else:
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
            ws = wb[wb.sheetnames[0]] if wb.sheetnames else None
            if ws is None:
                return None, "称量记录无工作表"
            rows = list(ws.iter_rows(values_only=True))
    except Exception as e:
        return None, f"读取称量记录失败: {e}"
    header = [str(c or "").strip() for c in rows[0]] if rows else []

    def find(key, default):
        for i, h in enumerate(header):
            if key in h:
                return i
        return default

    code_col, mass_col = find("样品编号", 1), find("称样量", 2)
    time_col, desc_col = find("称样时间", 0), find("试样描述", 3)
    parse_col = find("解析", 4)  # E列：强制解析标记
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
            entry["time"] = _parse_weigh_time(r[time_col] if time_col < len(r) else None)
        if not entry["desc"]:
            d = r[desc_col] if desc_col < len(r) else None
            entry["desc"] = str(d).strip() if d is not None else ""
        if parse_col < len(r):
            pv = r[parse_col]
            if pv is not None and str(pv).strip():
                entry["force_parse"] = True
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


def _equipment_env_candidates(row, equipment_config, matched_list):
    """收集用于查房间的设备编号候选字符串(去重保序)：表格设备 + 匹配设备 + 方法默认主检设备。
    matched_list 为 _override_equipment 匹配到的设备条目列表(可 None)；mainEquipmentNames 取首段即编号。"""
    cands = []

    def add(*vals):
        for v in vals:
            s = (str(v or "")).strip().split(",")[0].strip()
            if s and s not in cands:
                cands.append(s)

    add(row.get("equipment"))
    for eq in (matched_list or []):
        add(eq.get("mainEquipmentNames"), eq.get("name"),
            eq.get("no") or eq.get("code") or eq.get("equipmentCode"))
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
    """按标准号(standardNo)匹配处理规则；支持精确→包含双向。
    空 method=通配(对该方法文件所有样品生效，与目标标准号无关)；建议放末条作兜底。
    无匹配返回 None(→直接读取兜底)。"""
    std = (std_no or "").strip()
    for r in processing_rules or []:
        rm = str((r or {}).get("method") or "").strip()
        if not rm:                  # 空 method = 通配
            return r
        if std and (rm == std or rm in std or std in rm):
            return r
    return None


def _apply_processing(raw, rule, wp):
    """按处理规则把原始称样量(float)转为提交字符串。
    type ∈ 直接读取/小数位补充/换算处理/换算加补充；
    换算处理=raw×factor 修约到 rule.decimal_places；
    换算加补充=base=round(raw×factor,decimal_places) 后末尾补2位随机[01,49]
      (pad<50 不进位，修约回 decimal_places 位仍=base)，总小数位=decimal_places+2；
    小数位补充用 wp.decimal_places。"""
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
        if rtype == "换算加补充":
            base = round(float(raw) * factor, n)
            pad = random.randint(1, 49)  # ponytail: 01-49(<50不进位) 保证修约回 n 位=base
            return f"{base:.{n}f}{pad:02d}"
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
        "conditional": ("条件匹配", True),
        "":      ("配制序号", False),
    }
    # 仪器设置 → (placeholder, readonly)
    _EQUIPMENT_MODES = {
        "default":     ("默认设备", True),
        "specified":   ("设备编号", False),
        "conditional": ("条件匹配", True),
        "":            ("设备编号", False),
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
        "random":      ("随机称样", True),
        "none":        ("无需称样", True),
        "record":      ("称量记录", False),
        "process":     ("过程称量", False),
        "pdf":         ("PDF报告称样", True),
        "conditional": ("条件称样", True),
        "":            ("称样记录", False),
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
                 on_drag_start, on_drag_update, on_drag_end, on_data_update,
                 on_equipment_select=None):
        self.parent = parent
        self.index = index
        self.data = data
        self.column_widths = column_widths
        self.selection_manager = selection_manager
        self.on_path_select = on_path_select
        self.on_method_select = on_method_select
        self.on_spectrum_select = on_spectrum_select
        self.on_equipment_select = on_equipment_select
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
            has_button=True, on_button_click=lambda: self.on_equipment_button_click(),
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

        # 悬停显示完整运行状态/错误信息
        _tip = {'win': None}
        def _show_status_tip(_e):
            if _tip['win']:
                return
            msg = self.data.get("error_msg", "")
            full = msg if msg else self.data.get("status", "")
            if not full:
                return
            tw = tk.Toplevel(self.status_label)
            tw.wm_overrideredirect(True)
            tw.wm_geometry(f"+{self.status_label.winfo_rootx()+18}+{self.status_label.winfo_rooty()+self.status_label.winfo_height()+4}")
            ttk.Label(tw, text=full, background="#ffffe0", relief='solid', borderwidth=1,
                      font=("微软雅黑", 9), wraplength=400).pack(ipadx=4, ipady=2)
            _tip['win'] = tw
        def _hide_status_tip(_e):
            if _tip['win']:
                _tip['win'].destroy()
                _tip['win'] = None
        self.status_label.bind('<Enter>', _show_status_tip)
        self.status_label.bind('<Leave>', _hide_status_tip)

    # 状态显示样式（阶段1）
    STATUS_STYLE = {
        "待运行": ("white", "#888888"),
        "运行中": ("#dbeafe", "#2563eb"),
        "成功":   ("#dcfce7", "#16a34a"),
        "失败":   ("#fee2e2", "#dc2626"),
        "跳过":   ("#f5f5f5", "#9ca3af"),
        "中止":   ("#fef3c7", "#d97706"),
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

    def on_equipment_button_click(self):
        """设备列按钮：打开设备多选选择器（单元格值绑 equipment）"""
        if self.on_equipment_select and not self.destroyed:
            self.on_equipment_select(self.index)

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
        self.root.title("序列编辑器")
        self.root.geometry("1600x850")

        # 存储序列数据
        self.sequence_data = []
        self.current_file = None  # 当前已打开/已保存的序列文件路径（用于原地保存与标题显示）

        # 选择管理器
        self.selection_manager = CellSelectionManager()

        # 键盘状态
        self.ctrl_pressed = False
        self.shift_pressed = False

        # 列宽配置
        self.column_widths = [50, 260, 220, 140, 180, 260, 90, 90, 140]

        # 存储分隔线引用
        self.draggable_headers = []

        # 存储行组件引用
        self.row_widgets = []

        # LIMS 登录与 API（阶段1）
        self.login_system = MultiUserLoginSystem()
        self.api = DetectionAPI(self.login_system)
        self.logged_in = False
        self._env_eq_cfg_cache = {}  # sample_code -> 完整 equipment_config(设备选择器用，跨行去重)
        self._method_projects_cache = {}  # 方法查询键->该方法待登记项目列表(运行期跨行复用)

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
        # 实心彩色按钮统一加可见边框(深色 bevel 描边)；禁用态仍由各自 disabled 映射保持灰色不可用
        _sty = ttkb.Style()
        for _base in ("primary", "secondary", "danger"):
            _c = getattr(_sty.colors, _base)
            _d = "#%02x%02x%02x" % tuple(int(int(_c[i:i + 2], 16) * 0.55) for i in (1, 3, 5))
            _sty.configure(_base + ".TButton", relief="raised", borderwidth=2,
                           lightcolor=_d, darkcolor=_d, bordercolor=_d)
        # 主框架（档2: 转 ttkb，主题提供底色）
        main_frame = ttkb.Frame(self.root)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # 快捷工具栏：添加行/删除行/向下填充 + 登录状态显示
        quick_bar = ttkb.Frame(main_frame)
        quick_bar.pack(fill='x', pady=(0, 4))
        tb_add = ttkb.Button(quick_bar, text="+", command=self.add_row, width=3, padding=(2, 0), bootstyle="secondary")
        tb_add.pack(side='left', padx=(0, 2))
        tb_del = ttkb.Button(quick_bar, text="✕", command=self.delete_selected_rows, width=3, padding=(2, 0), bootstyle="secondary")
        tb_del.pack(side='left', padx=2)
        tb_fill = ttkb.Button(quick_bar, text="↓", command=self.fill_down, width=3, padding=(2, 0), bootstyle="secondary")
        tb_fill.pack(side='left', padx=2)
        tb_method = ttkb.Button(quick_bar, text="⚙", command=self.edit_method, width=3, padding=(2, 0), bootstyle="secondary")
        tb_method.pack(side='left', padx=2)
        _c = ttkb.Style().colors  # 主题配色，使 tk.Label 背景与界面一致
        self.login_status_label = tk.Label(quick_bar, text="👤 未登录", fg="#dc2626", bg=_c.bg,
                                          font=("Segoe UI", 9, "bold"))
        self.login_status_label.pack(side='right', padx=4)

        # 运行控制图标（与底部按钮同命令、同状态）：紧随编辑图标(➕✕↓)之后
        self.tb_run = ttkb.Button(quick_bar, text="▶", command=self.run_sequence, width=3, padding=(2, 0), bootstyle="primary")
        self.tb_run.pack(side='left', padx=(8, 2))
        self.tb_load = ttkb.Button(quick_bar, text="▼", command=self.load_sequence, width=3, padding=(2, 0), bootstyle="secondary")
        self.tb_load.pack(side='left', padx=2)
        self.tb_pause = ttkb.Button(quick_bar, text="⏸", command=self.toggle_pause, width=3, padding=(2, 0), state='disabled', bootstyle="secondary")
        self.tb_pause.pack(side='left', padx=2)
        self.tb_abort = ttkb.Button(quick_bar, text="⏹", command=self.abort_run, width=3, padding=(2, 0), state='disabled', bootstyle="danger")
        self.tb_abort.pack(side='left', padx=2)
        self.tb_clear_log = ttkb.Button(quick_bar, text="🧹", command=self.clear_log, width=3, padding=(2, 0), bootstyle="secondary")
        self.tb_clear_log.pack(side='left', padx=2)

        # 工具栏图标 tooltip（鼠标悬停显示功能）
        self._tips = [
            ttkb.ToolTip(tb_add, text="添加行"),
            ttkb.ToolTip(tb_del, text="删除选中行"),
            ttkb.ToolTip(tb_fill, text="向下填充"),
            ttkb.ToolTip(tb_method, text="编辑方法"),
            ttkb.ToolTip(self.tb_run, text="运行序列"),
            ttkb.ToolTip(self.tb_load, text="加载序列"),
            ttkb.ToolTip(self.tb_pause, text="暂停 / 继续"),
            ttkb.ToolTip(self.tb_abort, text="中止运行"),
            ttkb.ToolTip(self.tb_clear_log, text="清空日志"),
        ]

        # 创建表格容器
        self.create_table_container(main_frame)

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

        # 运行控制（运行/暂停/中止 合并到同一行；暂停/中止初始禁用，运行中启用）
        run_btn = ttkb.Button(bottom_frame, text="▶ 运行", command=self.run_sequence, width=8, padding=(6, 4), bootstyle="primary")
        run_btn.pack(side='right', padx=(6, 5), pady=4)
        self.abort_btn = ttkb.Button(bottom_frame, text="⏹ 中止", command=self.abort_run, width=8, padding=(6, 4), state='disabled', bootstyle="danger")
        self.abort_btn.pack(side='right', padx=2, pady=4)
        self.pause_btn = ttkb.Button(bottom_frame, text="⏸ 暂停", command=self.toggle_pause, width=8, padding=(6, 4), state='disabled', bootstyle="secondary")
        self.pause_btn.pack(side='right', padx=2, pady=4)

        # 顶部菜单栏（归类原工具栏按钮：序列/编辑/工具/用户）
        self.create_menubar()

    def create_menubar(self):
        """创建顶部菜单栏：序列/编辑/工具/用户（归类原界面按钮功能）"""
        c = ttkb.Style().colors  # 主题配色，与界面保持一致
        menu_opts = dict(
            bg=c.bg, fg=c.fg,
            activebackground=c.light, activeforeground=c.fg,
            borderwidth=0, relief="flat",
        )
        menubar = tk.Menu(self.root, **menu_opts)

        # 序列菜单：加载 / 保存 / 从选中行运行
        seq_menu = tk.Menu(menubar, tearoff=False, **menu_opts)
        seq_menu.add_command(label="加载序列", command=self.load_sequence)
        seq_menu.add_command(label="保存序列", command=self.save_sequence)
        seq_menu.add_command(label="序列另存为", command=self.save_sequence_as)
        seq_menu.add_separator()
        seq_menu.add_command(label="从选中行运行", command=self.run_from_selected)
        menubar.add_cascade(label="序列", menu=seq_menu)

        # 编辑菜单：添加行 / 删除行 / 向下填充 / 清空
        edit_menu = tk.Menu(menubar, tearoff=False, **menu_opts)
        edit_menu.add_command(label="添加行", command=self.add_row)
        edit_menu.add_command(label="删除行", command=self.delete_selected_rows)
        edit_menu.add_command(label="向下填充", command=self.fill_down)
        edit_menu.add_separator()
        edit_menu.add_command(label="清空", command=self.clear_all)
        menubar.add_cascade(label="编辑", menu=edit_menu)

        # 工具菜单：编辑方法 / 导出日志
        tool_menu = tk.Menu(menubar, tearoff=False, **menu_opts)
        tool_menu.add_command(label="编辑方法", command=self.edit_method)
        tool_menu.add_separator()
        tool_menu.add_command(label="导出日志", command=self._export_log)
        tool_menu.add_command(label="清空日志", command=self.clear_log)
        menubar.add_cascade(label="工具", menu=tool_menu)

        # 用户菜单：登录(切换用户) / 用户管理；self.user_menu 供 _refresh_user_menu 复用
        self.user_menu = tk.Menu(menubar, tearoff=False, **menu_opts)
        self.user_menu.add_command(label="登录", command=self._show_login_dialog)
        self.user_menu.add_command(label="用户管理", command=self._open_user_management)
        menubar.add_cascade(label="用户", menu=self.user_menu)

        self.root.config(menu=menubar)

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
            {"text": "T", "anchor": "w"},
            {"text": "RH", "anchor": "w"},
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
            self.on_data_update,
            self.select_equipment
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
                self.on_data_update,
                self.select_equipment
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
        """拖拽中：按鼠标当前所在行列更新选中范围(跨行/跨单元格，支持横向)"""
        row = self._row_at(event)
        if row is None:
            return
        col = self._col_at(event)
        if col is None:
            col = self.selection_manager.drag_column  # 离开单元格时保持起始列
        self.selection_manager.update_drag(row, col, self.ctrl_pressed, self.shift_pressed)
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

    def _col_at(self, event):
        """根据事件屏幕坐标定位鼠标当前所在的列索引(用于横向拖拽)；不在任何单元格上返回 None"""
        try:
            w = self.root.winfo_containing(event.x_root, event.y_root)
        except Exception:
            return None
        seen = set()
        while w is not None and id(w) not in seen:
            seen.add(id(w))
            for rw in self.row_widgets:
                for c, cell in enumerate(getattr(rw, "cells", None) or []):
                    if cell.cell_frame is w:
                        return c
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
        """选择称样记录文件（已有路径时，对话框定位到原路径所在目录）"""
        cur = ""
        if 0 <= row_index < len(self.sequence_data):
            cur = (self.sequence_data[row_index].get("weighing_path") or "").strip()
        file_path = filedialog.askopenfilename(
            title="选择称样记录文件",
            initialdir=os.path.dirname(cur) if cur and os.path.dirname(cur) else None,
            filetypes=[("Excel files", "*.xlsx;*.xls"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if file_path and 0 <= row_index < len(self.sequence_data):
            self.sequence_data[row_index]["weighing_path"] = file_path
            self.refresh_table()
            self.status_var.set(f"第 {row_index + 1} 行称样记录路径已设置")

    def select_method_file(self, row_index):
        """选择录入方法文件（已有路径时，对话框定位到原路径所在目录）"""
        cur = ""
        if 0 <= row_index < len(self.sequence_data):
            cur = (self.sequence_data[row_index].get("method_file") or "").strip()
        file_path = filedialog.askopenfilename(
            title="选择录入方法文件",
            initialdir=os.path.dirname(cur) if cur and os.path.dirname(cur) else None,
            filetypes=[("方法文件", "*.mtd"), ("All files", "*.*")]
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
        if not method_file:
            return {}
        try:
            y = load_method(method_file)
            return y.get("other_params_settings") or {}
        except Exception:
            return {}

    def _read_weighing_params(self, method_file):
        """读取方法 yaml 的 weighing_params（称样量模式与随机参数），返回 dict"""
        if not method_file:
            return {}
        try:
            y = load_method(method_file)
            return y.get("weighing_params") or {}
        except Exception:
            return {}

    def _method_default_desc(self, method_file):
        """判定方法是否支持「无称样记录默认录入」并取默认试样信息/描述值。
        返回 (supports, label, xx)：supports = 称样模式∉{record,process} 且 默认触发固定参数含
        试样信息/试样描述；label = 命中参数名(试样信息/试样描述，多个用/连)；
        xx = 命中值(多个用'；'连)。供运行前确认弹窗与默认 startTime 用。"""
        wp = self._read_weighing_params(method_file)
        wmode = (wp.get("weighing_mode") or "").strip()
        if wmode in ("record", "process"):
            return False, "", ""
        if wmode == "conditional":
            # 条件称样：任一规则需 record/process 则必须读称量记录 → 不支持无记录默认录入
            if any(str((r or {}).get("weighing_mode") or "") in ("record", "process")
                   for r in (wp.get("weighing_rules") or [])):
                return False, "", ""
        fps = (self._read_other_params(method_file) or {}).get("fixed_params") or []
        labels, vals = [], []
        for rule in fps:
            if (rule.get("trigger") or "").strip() not in ("默认", "默认触发"):
                continue
            for p in (rule.get("params") or []):
                if _norm_cn(p.get("name")) in ("试样信息", "试样描述"):
                    _nm = (p.get("name") or "").strip()
                    if _nm and _nm not in labels:
                        labels.append(_nm)
                    v = (p.get("value") or "").strip()
                    if v and v not in vals:
                        vals.append(v)
        return bool(vals), "/".join(labels), "；".join(vals)

    def _read_processing_rules(self, method_file):
        """读取方法 yaml 顶层 processing_rules（称量记录处理规则），返回 list"""
        if not method_file:
            return []
        try:
            y = load_method(method_file)
            return y.get("processing_rules") or []
        except Exception:
            return []

    def _apply_standard_config(self, row, st, prep, standard_rules=None):
        """标液配置应用到行"""
        row["standard_type"] = st
        if st == "fixed" and prep:
            row["configure_order"] = prep      # 带出固定编号，可覆盖
        elif st == "conditional":
            row["configure_order"] = ""        # 条件匹配时清空（运行时按样品匹配）
            row["_standard_rules"] = standard_rules  # 保存规则供 _submit_batch 使用
        elif st == "none":
            row["configure_order"] = ""        # 无需标液，清空
        # fresh / 未知：保留用户已填或空

    def _apply_equipment_config(self, row, setting, device, equipment_rules=None):
        """设备配置应用到行（仪器设置→显示模式；指定设备带出编号可覆盖）"""
        row["instrument_setting"] = setting
        if setting == "specified" and device:
            row["equipment"] = device          # 带出设备编号，可覆盖
        elif setting == "conditional":
            row["equipment"] = ""              # 条件匹配时清空（运行时按样品匹配）
            row["_equipment_rules"] = equipment_rules  # 保存规则供 _run_one_row 使用
        elif setting == "default":
            row["equipment"] = ""              # 默认设备，清空（显示占位）
        # 未配置：保留用户已填或空

    def _apply_method_params(self, row, method_file):
        """选方法 / 编辑器保存后：一次读取并回填标液+设备+称样量模式配置（共用）"""
        ops = self._read_other_params(method_file)
        self._apply_standard_config(row, ops.get("standard_type", ""), ops.get("preparation_number", ""),
                                    ops.get("standard_rules"))
        self._apply_equipment_config(row, ops.get("instrument_setting", ""),
                                     ops.get("device_number", ""), ops.get("equipment_rules"))
        wp = self._read_weighing_params(method_file)
        row["weighing_mode"] = (wp.get("weighing_mode") or "").strip() if wp else ""

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

    def _row_sample_codes(self, row, log):
        """本行全部样品编号(去重保序)。单样品走 _resolve_spectrum_pdf；
        多PDF目录按称样记录∩谱图展开(_resolve_samples)。同行样品方法可能不同，
        设备查询需逐个匹配行方法后再取设备，故返回全部而非仅首个。"""
        _, sc = self._resolve_spectrum_pdf(row, log)
        if sc:
            return [sc]
        rec_path = (row.get("weighing_path") or "").strip()
        if not (rec_path and os.path.isfile(rec_path)):
            return []
        wmap, _ = _read_weighing_records(rec_path)
        if wmap:
            _wp = self._read_weighing_params(row.get("method_file"))
            wmap = _merge_parallel_groups(wmap, (_wp or {}).get("non_parallel_suffixes"))
        samples, _ = self._resolve_samples(row, wmap, log)
        codes = []
        for s in samples or []:
            c = s[0] if isinstance(s, (list, tuple)) else s
            if c and c not in codes:
                codes.append(c)
        return codes

    def _row_first_sample_code(self, row, log):
        """取本行任一样品编号(设备列表/温湿度房间都是方法级，任一样品即可解析方法)。"""
        codes = self._row_sample_codes(row, log)
        return codes[0] if codes else None

    def _resolve_equipment_choices(self, row, log):
        """本行报验单的主检设备可选编号(设备选择器用)。
        样品→项目→ocMultipleChoicePage(报验单级，与网页端/手动录入界面一致)。
        同行样品方法可能不同，逐个匹配行方法文件，用首个匹配的样品查设备；
        返回去重保序的 [编号,...]；失败返回 None。"""
        # 报验单级候选(ocMultipleChoicePage)
        sample_codes = self._row_sample_codes(row, log)
        pid = None
        matched_sc = None
        last_reason = None
        for sc in sample_codes or []:
            try:
                projects = self.api.query_samples_by_conditions(
                    sample_code=sc, exact_match=True, log_func=log)
                projects, ferr = self._filter_projects_by_method(projects, row, log)
                if ferr or not projects:
                    last_reason = f"{sc}: {ferr or '无匹配项目'}"
                    continue
                pid = projects[0].get("projectId")
                if not pid:
                    last_reason = f"{sc}: 无 projectId"
                    continue
                matched_sc = sc
                break
            except Exception as e:
                last_reason = f"{sc}: 查询异常 {e}"
                log(f"查询设备可选列表异常({sc}): {e}")
                continue
        if not pid:
            log(f"设备查询：遍历 {len(sample_codes or [])} 个样品均未匹配方法项目(最后 {last_reason})，回退手动输入")
            return None
        try:
            choices = self.api.get_main_equipment_choices(pid, log)
        except Exception as e:
            log(f"查询设备可选列表异常({matched_sc}): {e}")
            return None
        if not choices:
            log(f"设备查询：样品 {matched_sc} 方法未配置主检设备(ocMultipleChoicePage 返回空)，回退手动输入")
            return None
        # 提编号
        codes = []
        for it in choices:
            raw = it.get("raw") or {}
            men = (raw.get("mainEquipmentNames") or "").strip()
            code = men.split(",")[0].strip() if men else ""  # "编号,名称" → 编号
            if not code:  # 无 mainEquipmentNames 时取编号字段，避免把"编号 名称"整体当编号(否则提交校验匹配不上)
                code = (raw.get("no") or raw.get("code") or raw.get("equipmentCode")
                        or raw.get("equipmentNo") or raw.get("number") or raw.get("billCode") or "").strip()
            if not code:
                code = (it.get("label") or "").split(",")[0].strip()
            if code and code not in codes:
                codes.append(code)
        return codes or None

    def _resolve_equipment_config(self, row, log):
        """查询 LIMS 取该行方法的完整 equipment_config(含 raw_data 全部检测设备)。
        样品→方法→get_all_configs→equipment；按 sample_code 缓存(_env_eq_cfg_cache)。
        需登录且行有样品(谱图)。失败/无样品返回 None。"""
        sample_code = self._row_first_sample_code(row, log)
        if not sample_code:
            return None
        cache = self._env_eq_cfg_cache
        if sample_code in cache:
            return cache[sample_code]
        eq_cfg = None
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
                    eq_cfg = (all_cfg or {}).get("equipment") or None
        except Exception as e:
            log(f"查询方法设备列表异常({sample_code}): {e}")
        cache[sample_code] = eq_cfg
        return eq_cfg

    def _override_equipment(self, equipment_config, device_field):
        """表格/设备规则指定的设备编号覆盖默认主检设备与称样设备（设备以序列表格为准；支持 ';' 分隔多个）。
        检测设备→主检设备(可多台)；称样设备→称样设备(单台，取首个)。
        返回 (equipment_config, matched_list, error_msg)。表格空则用方法默认(matched_list=None)。
        matched_list 为匹配到的【检测设备】条目，供 save_main_equipment 构造 items。"""
        codes = _split_device_codes(device_field)
        if not codes or not equipment_config:
            return equipment_config, None, None
        raw = equipment_config.get("raw_data") or []

        def _find(code):
            for eq in raw:
                name = (eq.get("name") or "").strip()
                men = (eq.get("mainEquipmentNames") or "").strip()
                men_first = men.split(',')[0].strip() if men else ""
                no = (eq.get("no") or eq.get("code") or eq.get("equipmentCode")
                      or eq.get("equipmentNo") or eq.get("number") or eq.get("billCode") or "").strip()
                if code == name or code == men or code == men_first or code == no:
                    return eq
            return None

        def _code_of(e):
            return ((e.get("mainEquipmentNames") or "").split(",")[0].strip()
                    or (e.get("no") or e.get("code") or e.get("equipmentCode")
                        or e.get("equipmentNo") or e.get("number") or e.get("billCode") or "").strip())

        matched = []      # 检测设备(主检)
        weigh_eq = None   # 称样设备(单台，取首个)
        for code in codes:
            eq = _find(code)
            if not eq:
                avail = [c for c in (_code_of(e) for e in raw) if c]
                return equipment_config, None, (
                    f"设备编号 {code} 未在方法设备中找到(可用编号: {', '.join(avail) or '无'})")
            if (eq.get('usedCategory') or '').strip() == '称样设备':
                if weigh_eq is None:
                    weigh_eq = eq
            elif id(eq) not in {id(m) for m in matched}:  # 检测设备去重
                matched.append(eq)
        # 主检设备(检测设备)：连接方式对齐 process_equipment_list(名称 "; "、id ",")
        # 每条按"编号,名称,有效期"构造：mainEquipmentNames 非空则用它，否则由 no/name/checkOutDate 补全——
        # 条件/指定设备常无 mainEquipmentNames(非默认设备)，不补全则 build 按 ',' 解析会丢有效期/编号(信息不全)
        displays, names, ids = [], [], []
        for eq in matched:
            men = (eq.get("mainEquipmentNames") or "").strip()
            name = (eq.get("name") or "").strip()
            if men:
                display = men
            else:
                _eno = (eq.get("no") or eq.get("code") or eq.get("equipmentCode")
                        or eq.get("equipmentNo") or eq.get("number") or eq.get("billCode") or "").strip()
                _cod = eq.get("checkOutDate")
                _edate = _cod[:10] if isinstance(_cod, str) else (str(_cod)[:10] if _cod else "")
                display = f"{_eno},{name},{_edate}" if _eno else name
            eid = eq.get("equipmentBillId") or eq.get("id")
            if not eid:
                continue   # 无 id 不能提交，跳过以保持 ids 与名称序列对齐(否则 build 按 index 错配 name↔id)
            displays.append(display)
            names.append(display)
            ids.append(str(eid))
        cfg = dict(equipment_config)
        cfg["mainEquipment"] = "; ".join(displays)
        cfg["mainEquipmentNames"] = "; ".join(names)
        cfg["mainEquipmentIds"] = ",".join(ids)
        # 称样设备(单台)：随 saveOcExperiment 提交，格式对齐 GUI(detection_entry_main:1982-1989)
        if weigh_eq is not None:
            wno = _code_of(weigh_eq)
            wname = (weigh_eq.get("name") or weigh_eq.get("equipmentName")
                     or weigh_eq.get("equipmentBillName") or "").strip()
            cod = weigh_eq.get("checkOutDate")
            wdate = cod[:10] if isinstance(cod, str) else (str(cod)[:10] if cod else '')
            wid = weigh_eq.get("equipmentBillId") or weigh_eq.get("id")
            cfg["weighingEquipmentId"] = str(wid) if wid else ""
            cfg["weighingEquipmentBaseName"] = wname
            cfg["weighingEquipment"] = f"{wno},{wname},{wdate}" if wno else wname
            cfg["weighingEquipmentRaw"] = weigh_eq
        return cfg, matched, None

    def _enrich_weighing_bill(self, equipment_config, sp_ids_str, log):
        """称样设备 weighingEquipmentRaw 回填真实台账对象。
        selectByDetectionMethodId 返回的 raw 是「方法-设备配置行」(元数据=配置人/时间，如 李金玲/2026-03)；
        总和实验无称样列时 LIMS 仅在 weighingEquipment 为真实台账对象(元数据=台账创建人/时间)时绑定称样设备。
        故用 ocChoicePage(网页端称样设备选择源) 取台账对象覆盖之；id/name 仍取自 weighingEquipmentId/BaseName。
        遍历批内全部样品项目查 ocChoicePage(设备可能只挂在批内某个样品项目下)，命中即止。"""
        wid = (equipment_config or {}).get("weighingEquipmentId")
        if not wid or not sp_ids_str:
            return equipment_config
        sp_ids = [s.strip() for s in sp_ids_str.split(",") if s.strip()]
        if not sp_ids:
            return equipment_config
        for sp in sp_ids:
            try:
                choices = self.api.get_weighing_equipment_choices(sp, log)
            except Exception as e:
                log(f"[设备] 取称样设备台账异常(sp={sp})，跳过: {e}")
                continue
            for it in choices:
                raw = it.get("raw") or {}
                _rids = {str(raw.get("id") or ""), str(raw.get("equipmentBillId") or "")}
                if str(wid) in _rids:
                    cfg = dict(equipment_config)
                    cfg["weighingEquipmentRaw"] = raw
                    log(f"[设备] 称样设备台账回填: id={wid} sp={sp} creator={raw.get('creatorName')!r} "
                        f"createDatetime={raw.get('createDatetime')!r}")
                    return cfg
        log(f"⚠ [设备] 称样设备台账未在批内任何样品项目找到(id={wid}, sp_ids={sp_ids_str})，本批将不绑定称样设备")
        return equipment_config

    def select_spectrum_path(self, row_index):
        """选择谱图文件路径（已有路径时，对话框定位到原路径）"""
        cur = ""
        if 0 <= row_index < len(self.sequence_data):
            cur = (self.sequence_data[row_index].get("spectrum_path") or "").strip()
        path = filedialog.askdirectory(
            title="选择谱图文件夹路径",
            initialdir=cur if cur and os.path.isdir(cur) else None
        )
        if path and 0 <= row_index < len(self.sequence_data):
            row = self.sequence_data[row_index]
            row["spectrum_path"] = path
            self.refresh_table()
            self.status_var.set(f"第 {row_index + 1} 行谱图文件路径已设置")

    def select_equipment(self, row_index):
        """打开设备选择：能取到方法设备列表则勾选(只显示编号)；取不到(如样品方法未切换/不可查)
        则回退手动输入(; 分隔)。结果写回行 equipment。"""
        if not (0 <= row_index < len(self.sequence_data)):
            return
        row = self.sequence_data[row_index]
        # 取方法主检设备可选编号(走 ocMultipleChoicePage，不经 getOcExperiment，避免'方法不同')；
        # 未登录/样品不可查/取空 → 走手动输入
        items = None
        if self._ensure_session_silent():
            try:
                items = self._resolve_equipment_choices(row, self._log)
            except Exception:
                items = None
        else:
            self._log("设备查询：未登录或会话失效，回退手动输入")
        current = row.get("equipment") or ""
        if items:
            checked = set(_split_device_codes(current))
            selected = self._open_equipment_picker(row_index, items, checked)
        else:
            # 未取到方法设备列表：不再弹出手动输入框，仅记日志/状态栏提示
            self._log(f"第 {row_index + 1} 行：未取到方法设备列表(样品不可查或方法未配置设备)，跳过设备选择")
            self.status_var.set(f"第 {row_index + 1} 行未取到设备列表，已跳过")
            return
        if selected is None:
            return  # 用户取消
        row["equipment"] = ";".join(selected)
        self.refresh_table()
        self.status_var.set(f"第 {row_index + 1} 行设备已设置: {row['equipment'] or '(默认)'}")

    def _open_equipment_picker(self, row_index, items, checked):
        """模态勾选对话框(只显示编号)。确定返回选中编号列表(按列表顺序)；取消返回 None。"""
        win = ttkb.Toplevel(self.root)  # 主题 Toplevel，与主界面风格一致
        win.title(f"选择设备 - 第 {row_index + 1} 行")
        win.transient(self.root)
        win.grab_set()
        _font = ("Microsoft YaHei", 10)
        # 放大勾选框：复用 detection_entry_main 的 ttk indicator 重建；仅创建一次并缓存——
        # 重复调用会重新注册同名 indicator element 并替换 PhotoImage，第二次起损坏 style 致复选框消失
        if not getattr(self, "_eq_picker_cb_style", None):
            try:
                from detection_entry_main import _make_large_checkbutton_style
                self._eq_picker_cb_style = _make_large_checkbutton_style(1.0) or "EqPicker.TCheckbutton"
            except Exception:
                self._eq_picker_cb_style = "EqPicker.TCheckbutton"
        cb_style = self._eq_picker_cb_style
        ttkb.Style().configure(cb_style, font=_font)
        result = {"value": None}
        cvars = []

        def on_ok():
            result["value"] = [code for code, v in cvars if v.get()]
            win.destroy()

        # 关键：按钮区先 pack 到底部，再 pack 内容区(expand)；按钮永不被内容挤出可视区
        btns = ttkb.Frame(win)
        btns.pack(side="bottom", fill="x", padx=12, pady=(0, 12))
        ttkb.Button(btns, text="取消", command=win.destroy, width=8, bootstyle="secondary").pack(side="right")
        ttkb.Button(btns, text="确定", command=on_ok, width=8, bootstyle="primary").pack(side="right", padx=(0, 8))

        # 内容区：canvas 可滚动(设备多时竖向滚动)
        body = ttkb.Frame(win)
        body.pack(fill="both", expand=True, padx=12, pady=(12, 4))
        _bg = ttkb.Style().colors.bg
        canvas = tk.Canvas(body, highlightthickness=0, width=320, bg=_bg)
        sb = ttkb.Scrollbar(body, orient="vertical", command=canvas.yview)
        inner = ttkb.Frame(canvas)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        _inner_id = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(_inner_id, width=e.width))  # inner 宽度跟随 canvas，防塌缩
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        checked = set(checked or [])
        for code in items:
            v = tk.BooleanVar(value=(code in checked))
            ttk.Checkbutton(inner, text=code, variable=v, style=cb_style).pack(fill="x", padx=8, pady=5)
            cvars.append((code, v))
        win.update_idletasks()
        win.geometry(f"360x{min(980, 160 + len(items) * 54)}")
        win.minsize(320, 360)
        win.update_idletasks()  # 确保 winfo 反映实际尺寸后再算居中
        rx, ry = self.root.winfo_rootx(), self.root.winfo_rooty()
        x = rx + (self.root.winfo_width() - win.winfo_width()) // 2
        y = ry + (self.root.winfo_height() - win.winfo_height()) // 2
        win.geometry(f"+{max(0, x)}+{max(0, y)}")  # 居中到主窗口
        win.wait_window()
        return result["value"]

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

    def _update_title(self):
        """根据当前打开的序列文件更新窗口标题"""
        if self.current_file:
            name = os.path.splitext(os.path.basename(self.current_file))[0]
            self.root.title(f"序列编辑器 · {name}")
        else:
            self.root.title("序列编辑器")

    def _write_sequence(self, file_path):
        """实际写序列到文件，并更新 current_file/标题。失败弹错。"""
        try:
            save_data = {
                "sequence_data": [],
                "column_widths": self.column_widths
            }

            for row in self.sequence_data:
                # 只持久化用户手填/路径/环境；equipment/configure_order/标液模式/仪器设置/称样模式
                # 运行时从方法派生(_apply_method_params)，sample_code 等隐藏字段不存(避免残留坑)
                save_data["sequence_data"].append({
                    "id": row["id"],
                    "weighing_path": row["weighing_path"],
                    "method_file": row["method_file"],
                    "spectrum_path": row["spectrum_path"],
                    "temperature": row.get("temperature", ""),
                    "humidity": row.get("humidity", ""),
                })

            with open(file_path, 'w', encoding='utf-8') as f:
                yaml.dump(save_data, f, allow_unicode=True, indent=2, sort_keys=False)

            self.current_file = file_path
            self._update_title()
            self.status_var.set(f"序列已保存到: {file_path}")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {str(e)}")

    def save_sequence(self):
        """保存序列到文件（已打开文件则原地覆盖保存，否则弹窗选择）"""
        if not self.sequence_data:
            self.status_var.set("提示: 没有数据可保存")
            return

        file_path = self.current_file
        if not file_path:
            file_path = filedialog.asksaveasfilename(
                title="保存序列文件",
                defaultextension=".seq",
                filetypes=[("序列文件", "*.seq"), ("YAML files", "*.yaml;*.yml"), ("All files", "*.*")]
            )
            if not file_path:
                return
        self._write_sequence(file_path)

    def save_sequence_as(self):
        """序列另存为：始终弹窗选新路径(默认定位到当前文件)，保存后切换到新文件"""
        if not self.sequence_data:
            self.status_var.set("提示: 没有数据可保存")
            return
        initdir = os.path.dirname(self.current_file) if self.current_file and os.path.dirname(self.current_file) else None
        initialfile = os.path.basename(self.current_file) if self.current_file else None
        file_path = filedialog.asksaveasfilename(
            title="序列另存为",
            initialdir=initdir,
            initialfile=initialfile,
            defaultextension=".seq",
            filetypes=[("序列文件", "*.seq"), ("YAML files", "*.yaml;*.yml"), ("All files", "*.*")]
        )
        if not file_path:
            return
        self._write_sequence(file_path)

    def load_sequence(self):
        """从文件加载序列"""
        file_path = filedialog.askopenfilename(
            title="加载序列文件",
            initialdir=os.path.dirname(self.current_file) if self.current_file and os.path.dirname(self.current_file) else None,
            filetypes=[("序列文件", "*.seq"), ("YAML files", "*.yaml;*.yml"), ("JSON files", "*.json"), ("All files", "*.*")]
        )

        if not file_path:
            return

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                loaded_data = yaml.safe_load(f) or {}

            self.sequence_data = []
            self.selection_manager.clear_all_selection()

            if "column_widths" in loaded_data:
                self.column_widths = loaded_data["column_widths"]
                self.update_header_layout()

            sequence_data = loaded_data.get("sequence_data", loaded_data)
            for item in sequence_data:
                mf = item["method_file"]
                new_row = {
                    "id": item["id"],
                    "weighing_path": item["weighing_path"],
                    "method_file": mf,
                    "spectrum_path": item["spectrum_path"],
                    "temperature": item.get("temperature", ""),
                    "humidity": item.get("humidity", ""),
                    # 以下不持久化：equipment/configure_order/standard_type/instrument_setting/
                    # weighing_mode 运行时从方法派生；sample_code 等隐藏字段不恢复(避免残留)
                    "reference_material": "",
                    "equipment": "",
                    "sample_code": "",
                    "configure_order": "",
                    "standard_type": "",
                    "instrument_setting": "",
                    "weighing_mode": "",
                    "experiment_code": "",
                    "status": item.get("status", "待运行"),
                    "error_msg": item.get("error_msg", ""),
                }
                # equipment/configure_order/标液模式/仪器设置/称样模式 从方法重派生(同选方法时一致)
                if mf:
                    try:
                        self._apply_method_params(new_row, mf)
                    except Exception:
                        pass
                self.sequence_data.append(new_row)

            self.refresh_table()
            self.current_file = file_path
            self._update_title()
            self.status_var.set(f"已加载序列文件: {file_path}")

        except Exception as e:
            messagebox.showerror("错误", f"加载失败: {str(e)}")

    def edit_method(self):
        """方法编辑入口：取序列里最上面的有方法的行，打开方法编辑器"""
        target = None
        for row in self.sequence_data:
            mf = row.get("method_file")
            if mf:
                target = mf
                break
        self._open_method_editor(target)

    def _open_method_editor(self, file_path=None):
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
        top = ttkb.Toplevel(self.root)
        top.title("录入方法编辑器")
        top.transient(self.root)  # 置于序列编辑器之上(owned)：加载方法时序列窗口不再抢占到最前端
        app = mre.QueryAppFixed(top)
        if file_path:
            app.load_config_file(file_path)
        self._method_editor_top = top
        top.protocol("WM_DELETE_WINDOW", lambda: self._on_method_editor_closed(top, file_path))

    def _on_method_editor_closed(self, top, file_path):
        """编辑器关闭：把方法里的标液+设备配置回填到所有使用该方法文件的行"""
        setattr(self, "_method_editor_top", None)
        top.destroy()
        if not file_path:
            return
        affected = [row for row in self.sequence_data if row.get("method_file") == file_path]
        for row in affected:
            self._apply_method_params(row, file_path)
        if affected:
            self.refresh_table()
            self._log(f"已从方法同步标液+设备配置到 {len(affected)} 行")

    def run_from_selected(self):
        """从当前选中行开始运行序列（未选中则提示）。"""
        if self._running:
            self.status_var.set("提示: 序列正在运行中")
            return
        if not self.sequence_data:
            self.status_var.set("提示: 没有可运行的序列数据")
            return
        sel = sorted(self.selection_manager.selected_rows) or sorted({r for r, _ in self.selection_manager.selected_cells})
        if not sel:
            self.status_var.set("提示: 请先选中要开始运行的行")
            return
        self.run_sequence(sel[0] + 1)  # 转为 1-based

    def run_sequence(self, start_row=None):
        """运行序列（阶段1：真实逐行提交，半自动）。start_row: 1-based 起始行，None=从头运行。"""
        if self._running:
            self.status_var.set("提示: 序列正在运行中")
            return
        if not self.sequence_data:
            self.status_var.set("提示: 没有可运行的序列数据")
            return
        # 起始行(1-based)；None 或 1=从头
        n = len(self.sequence_data)
        sr = start_row if start_row is not None else 1
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
        # 默认录入确认：称样模式非 record/process、且默认触发含试样信息/试样描述、且称样记录路径为空的行，
        # 会以谱图目录全部 PDF 作样品、试样信息取方法默认、开始时间取上一工作日——运行前统一确认。
        _def_rows = []
        for _i in range(start_idx, n):
            _r = self.sequence_data[_i]
            if (_r.get("weighing_path") or "").strip():
                continue
            _sup, _lbl, _xx = self._method_default_desc(_r.get("method_file"))
            if _sup:
                _mn = os.path.splitext(os.path.basename(_r.get("method_file") or ""))[0]
                if (_mn, _lbl, _xx) not in _def_rows:
                    _def_rows.append((_mn, _lbl, _xx))
        if _def_rows:
            if len(_def_rows) == 1:
                _info = f"录入方法中{_def_rows[0][1]}默认是「{_def_rows[0][2]}」"
            else:
                _info = "；".join(f"[{_m}]{_lbl}={_x}" for _m, _lbl, _x in _def_rows)
            _pwd = _prev_workday().strftime("%Y-%m-%d")
            _msg = (f"称量记录为空，将录入谱图文件夹中所有编号\n"
                    f"{_info}\n"
                    f"分析开始时间默认是{_pwd}。\n"
                    f"是否继续录入？")
            if not messagebox.askyesno("默认录入确认", _msg):
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
        self.abort_btn.configure(state='normal' if running else 'disabled')
        self.tb_pause.configure(state='normal' if running else 'disabled', text="⏸")
        self.tb_abort.configure(state='normal' if running else 'disabled')
        self.status_var.set("运行中..." if running else "就绪")

    def _run_worker(self, start_idx=0):
        """worker 线程：从 start_idx 起逐行执行，所有 UI 更新经 _ui_q。
        循环按当前行数动态推进：运行期间新增的行(append 到末尾)会在当行结束后自动纳入运行。"""
        self._method_projects_cache.clear()  # 每次运行重建方法查询缓存(运行中样品已登记会使旧池过期)
        weighing_caches = {}  # {组名: {sample_code: {masses,desc}}} 跨行共享称样量(称样量共享组)
        idx = start_idx
        while idx < len(self.sequence_data):
            if self._stop.is_set():
                break
            self._pause.wait()
            if self._stop.is_set():
                break
            n = len(self.sequence_data)
            self._ui_q.put(("status", (idx, "运行中", "")))
            self._log(f"--- 第 {idx + 1}/{n} 行 ---")
            try:
                outcome = self._run_one_row(idx, weighing_caches)
            except Exception as e:
                self._ui_q.put(("status", (idx, "失败", str(e))))
                self._log(f"[行{idx + 1}] 异常: {e}")
                outcome = "fail"
            if outcome == "abort":
                self._log("用户中止序列")
                break
            idx += 1
        self._ui_q.put(("done", None))

    def _aborted(self, idx):
        """worker 线程：检查中止标志，已中止则标记本行并返回 True（_run_one_row 各阶段边界调用）"""
        if self._stop.is_set():
            self._ui_q.put(("status", (idx, "中止", "用户中止")))
            return True
        return False

    def _run_one_row(self, idx, weighing_caches=None):
        """单行流水线（worker 线程内）。返回 'ok'/'skip'/'abort'/'fail'，失败自行 put status。
        谱图为目录+多PDF+未填样品编号时，按「称样记录excel ∩ 谱图目录」展开为多个样品；
        跨样品按方法(default_rules.to_id)合并：所有样品的同方法项目合并到一个实验编号提交，
        再按 max_select 分批。某样品查询失败仅跳过(尽量多录入)，任一提交失败整行标失败。
        weighing_caches: 运行级 {组名: 缓存dict}；方法设了称样量共享组时同组跨行共享缓存。"""
        row = self.sequence_data[idx]
        log = lambda m: self._log(f"[行{idx + 1}] {m}")

        # 读称样参数 + 称样记录：record/process 模式必须读(注入称样量)；
        # 其他模式若填了路径也读(供多样品展开确定样品编号，但不注入称样量)
        wp = self._read_weighing_params(row.get("method_file"))
        wmode = (wp.get("weighing_mode") or "").strip() if wp else ""
        rec_path = (row.get("weighing_path") or "").strip()
        # 无称样记录默认录入：方法支持(模式非record/process + 默认触发试样信息) 且 称样记录路径为空
        _sup_def, _, _ = self._method_default_desc(row.get("method_file"))
        defaults_no_record = bool(_sup_def and not rec_path)
        wmap = None
        # 是否强制需要称量记录文件：record/process 模式必读；conditional 时只要任一规则命中 record/process 即必读
        _needs_record = wmode in ("record", "process")
        if wmode == "conditional":
            _needs_record = any(str((r or {}).get("weighing_mode") or "") in ("record", "process")
                                for r in (wp.get("weighing_rules") or []))
        if _needs_record:
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
            wmap = _merge_parallel_groups(wmap, (wp or {}).get("non_parallel_suffixes"))

        # 解析本行要录入的样品：单样品，或目录多PDF按 excel∩谱图 展开
        samples, serr = self._resolve_samples(row, wmap, log, defaults_no_record=defaults_no_record)
        if serr or not samples:
            self._ui_q.put(("status", (idx, "失败", "无法确定样品")))
            log(f"失败: {serr or '谱图目录无匹配PDF且未填样品编号'}")
            return "fail"
        if len(samples) == 1:
            self._ui_q.put(("rowdata", (idx, {"sample_code": samples[0][0]})))
        if len(samples) > 1:
            log(f"本行展开为 {len(samples)} 个样品(称样记录 ∩ 谱图目录)")
        # 诊断：各样品称样量(基样平行 + 标记，核对 M 等是否归到基样标记而非平行)
        if wmap:
            for _sc, _pdfs in samples:
                _e = wmap.get(_sc) or {}
                _parts = [f"称样量{[f'{x:.4f}' for x in (_e.get('masses') or [])]}"]
                for _mk, _mv in (_e.get('marker_masses') or {}).items():
                    _parts.append(f"{_mk}标记{[f'{x:.4f}' for x in _mv]}")
                log(f"  样品 {_sc} → {' '.join(_parts)} 谱图{len(_pdfs)}个")

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
        # 标准溶液预检（行级一次，各样品共用；支持逗号分隔多个标液，逐个校验）
        if configure_order:
            # 称样时间(分析开始时间)：历史标液按「称样时间 ≤ 有效期」校验用；
            # 有称样记录取首个样品称样时间，无则取上一工作日
            sample_time = next(((_e or {}).get("time") for _e in (wmap or {}).values()
                                if (_e or {}).get("time")), None) or _prev_workday()
            for _co in [x.strip() for x in configure_order.replace("，", ",").split(",") if x.strip()]:
                log(f"校验标准溶液 {_co} ...")
                cid, emsg = self.api.get_solution_configure_id(_co, log, sample_time)
                if not cid:
                    self._ui_q.put(("status", (idx, "失败", f"标准溶液:{emsg}")))
                    log(f"失败: 标准溶液校验未过 - {emsg}")
                    return "fail"

        _ops = self._read_other_params(row.get("method_file")) or {}
        _grp = (wp.get("weighing_share_group") or "").strip() if wp else ""
        # 称样量共享组：同名组跨行复用同一缓存(同一样品首方法生成、后方法复用)；无组名则每行独立
        _pcache = weighing_caches.setdefault(_grp, {}) if (_grp and weighing_caches is not None) else {}
        ctx = {
            "configure_order": configure_order, "wp": wp, "wmode": wmode, "wmap": wmap,
            "weighing_rules": (wp.get("weighing_rules") or []) if wmode == "conditional" else [],
            "fixed_params": _ops.get("fixed_params") or [], "method_file": row.get("method_file") or "",
            "standard_type": (row.get("standard_type") or "").strip(),
            "standard_rules": (row.get("_standard_rules") or _ops.get("standard_rules") or []),
            "defaults_no_record": defaults_no_record,
            "primary_cache": _pcache,  # {sample_code: {masses, desc}} 首项目(组分)生成后供依赖项目(总和)复用
            # 标「解析」的样品若稀释，自动启用稀释备注(无需方法勾选)；下游 dil_extra>1 仅稀释记录实际生成备注
            "dilution_remark_enabled": bool(_ops.get("dilution_remark_enabled")) or any(
                ((wmap or {}).get(_sc) or {}).get("force_parse") for _sc, _ in samples),
            "dilution_volume_column": (_ops.get("dilution_volume_column") or "").strip(),
        }

        # 阶段A：逐样品查询+过滤(此阶段不上传谱图)，收集所有样品的项目
        # 某样品查询失败仅跳过(尽量多录入)，不影响其余样品的合并提交
        # 自适应：样品数≤阈值→按样品并行精确查(避免1样品却拉方法池全量234条)；>阈值→方法池(1请求)摊销固定开销。
        # 日期窗口统一喂方法池与逐样品精确查：按方法文件配置收窄提速/放宽找旧样品。
        mf = row.get("method_file") or ""
        threshold = self._read_sample_parallel_threshold(mf)
        window_days = self._read_date_window_days(mf)
        ctx["date_window_days"] = window_days  # 透传给 _collect_sample_projects 兜底查询
        method_names = self._method_query_names(mf)
        target_codes = [sc for sc, _ in samples]
        projects_by_sample = {}
        use_pool = len(target_codes) > threshold
        if use_pool and method_names and target_codes:
            pooled = self._query_projects_by_methods(method_names, log, days=window_days)
            if pooled:
                for _p in pooled:
                    projects_by_sample.setdefault(_p.get("sampleCode"), []).append(_p)
                log(f"按方法查询完成：{len(pooled)} 个项目，覆盖 {len(projects_by_sample)} 个样品（窗口{window_days}天）")
            else:
                log("按方法查询无结果，全部走逐样品精确查询")
        else:
            log(f"样品数 {len(target_codes)} ≤ 阈值 {threshold}，按样品并行精确查询（跳过方法池，窗口{window_days}天）")
        # 样品少 或 方法池未覆盖的样品：并行精确查询(同一窗口)
        missing = [sc for sc in target_codes if sc not in projects_by_sample]
        if missing:
            t0 = time.time()
            log(f"逐样品精确查询 {len(missing)} 个(并行, 窗口{window_days}天) ...")
            fetched = self._query_samples_parallel(missing, log, days=window_days)
            for sc, plist in fetched.items():
                projects_by_sample.setdefault(sc, []).extend(plist)
            log(f"逐样品精确查询完成：{len(fetched)}/{len(missing)} 命中，用时 {time.time()-t0:.1f}s")
        all_items = []  # [{project, sample_code, sample_id, pdf_paths}, ...]
        n_skip = 0
        skipped_samples = []  # [(code, reason), ...]
        for si, (sc, pdf_paths) in enumerate(samples):
            if self._aborted(idx):
                return "abort"
            if len(samples) > 1:
                log(f"=== 样品 {si + 1}/{len(samples)}：{sc} ===")
            items, serr = self._collect_sample_projects(idx, row, sc, pdf_paths, ctx, log, projects_by_sample)
            if serr:
                log(f"警告[{sc}]: {serr}，跳过该样品")
                n_skip += 1
                skipped_samples.append((sc, serr))
                continue
            all_items.extend(items)
        if not all_items:
            self._ui_q.put(("status", (idx, "失败", "无可用样品项目")))
            self._ui_q.put(("rowdata", (idx, {"skipped_list": list(skipped_samples)})))
            log("失败: 所有样品均无可录入项目")
            if skipped_samples:
                log(f"各样品未录入原因（{len(skipped_samples)} 个）：")
                for sc, reason in skipped_samples:
                    log(f"  - {sc}：{reason}")
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

        # 按 spectrum_filter_rules 过滤各项目的样品谱图(如 总和排除m、5mm以内仅m)；common_pdfs 不受影响
        filter_rules = self._read_spectrum_filter_rules(row.get("method_file") or "")
        for it in all_items:
            before = len(it.get("pdf_paths") or [])
            it["pdf_paths"] = _filter_pdfs_by_rule(
                it.get("pdf_paths") or [], it["project"].get("projectName", ""), filter_rules)
            after = len(it["pdf_paths"])
            if before != after:
                log(f"[{it['sample_code']}] {it['project'].get('projectName','')}: "
                    f"谱图筛选 {before}→{after} 个")

        # 上传谱图：同一文件只 POST 一次，fileId 跨样品/项目复用(提交阶段 fid_info 按 fileId 去重)
        _upload_cache = {}  # pdf_path -> uploaded entry or None(失败)
        _posted, _failed = [], []
        for it in all_items:
            if self._aborted(idx):
                return "abort"
            uploaded = []
            for pdf in list(dict.fromkeys((it.get("pdf_paths") or []) + common_pdfs)):
                if pdf not in _upload_cache:
                    up_ok, up_res = self.api.upload_spectrum_file(pdf, "", log)
                    rec = ({
                        "fileId": up_res.get("id"),
                        "fileName": up_res.get("orgName") or up_res.get("name"),
                        "url": up_res.get("url"),
                    } if up_ok and isinstance(up_res, dict) else None)
                    _upload_cache[pdf] = rec
                    if rec:
                        _posted.append(rec["fileName"])
                    else:
                        _failed.append(os.path.basename(pdf))
                rec = _upload_cache.get(pdf)
                if rec:
                    uploaded.append(rec)
            it["spectrum_uploaded"] = uploaded
        if _posted:
            log(f"谱图上传 {len(_posted)} 个(去重): {', '.join(_posted)}")
        if _failed:
            log(f"警告: 谱图上传失败 {len(_failed)} 个: {', '.join(_failed)}")

        # 阶段B：按 query_rules 顺序 + input_method(方法=跨样品合并 / 样品=按样品) 规划并提交
        # 方法模式下，同一目标方法(default_rules/filename_rules.to_id)的不同样品项目并入一个实验编号
        method_file = ctx["method_file"]
        switch_rules = self._read_switch_rules(method_file)
        equipment_rules = row.get("_equipment_rules") or []
        ops = self._read_other_params(row.get("method_file"))
        # 从保存的序列加载时行可能没有 _equipment_rules，从方法文件回退读取
        if not equipment_rules and row.get("instrument_setting") == "conditional":
            equipment_rules = ops.get("equipment_rules") or []
        # 条件实验过程(labProc)：方法勾选启用时按规则匹配，命中 lab_proc 覆盖 experimentProcess（并按之拆批）
        lab_proc_rules = (ops.get("lab_proc_rules") or []) if ops.get("lab_proc_enabled") else []
        fallback_mid = "" if switch_rules else self._read_switch_method_id(method_file)
        for it in all_items:
            pname = (it["project"].get("projectName") or "").strip()
            it["projectName"] = pname            # 供 _plan_submission_batches 按 project 选
            it["_qr_idx"] = it["project"].get("_qr_idx")  # 多方法分批：标记该项目命中的 query_rule 序号
            # 统一规则匹配：project_name / filename / desc 三条件 AND，空=不限
            # 试样描述取自称量记录(无记录则为空，desc 条件不生效)
            wrec = (ctx.get("wmap") or {}).get(it["sample_code"]) or {}
            desc = wrec.get("desc") or ""
            it["wdate"] = _date_str(wrec.get("time"))   # 称样日期分组键；无时间→""(并入合并组)
            mid = _match_switch_rule(switch_rules, pname, it.get("pdf_paths") or [], desc)
            if switch_rules:
                if mid:
                    log(f"[{it['sample_code']}] 命中切换规则 → 方法ID {mid}（项目:{pname}）")
                else:
                    log(f"[{it['sample_code']}] 未命中切换规则（项目:{pname}），用默认方法")
            it["switch_mid"] = mid or fallback_mid
            # 条件设备匹配：按 equipment_rules 为每个样品匹配设备（用于拆批）
            if equipment_rules:
                eq_dev = _match_equipment_rule(equipment_rules, pname, it.get("pdf_paths") or [], desc)
                it["_equipment"] = eq_dev
                if eq_dev:
                    log(f"[{it['sample_code']}] 命中设备规则 → {eq_dev}（项目:{pname}）")
                else:
                    log(f"[{it['sample_code']}] 未命中设备规则，使用默认设备")
            # 条件实验过程匹配：命中 lab_proc 覆盖该样品 experimentProcess（用于拆批）
            if lab_proc_rules:
                lp = _match_lab_proc_rule(lab_proc_rules, pname, it.get("pdf_paths") or [], desc)
                it["_lab_proc"] = lp
                if lp:
                    log(f"[{it['sample_code']}] 命中实验过程规则 → {lp}（项目:{pname}）")
            # 条件标液匹配：按 standard_rules 为每个样品匹配配制序号（用于按批提交）
            if ctx.get("standard_type") == "conditional" and ctx.get("standard_rules"):
                sm = _match_conditional_rule(ctx["standard_rules"], pname, it.get("pdf_paths") or [], desc)
                it["_standard"] = str((sm or {}).get("preparation_number") or "").strip()
                if it["_standard"]:
                    log(f"[{it['sample_code']}] 命中标液规则 → {it['_standard']}（项目:{pname}）")
                else:
                    log(f"[{it['sample_code']}] 未命中标液规则（项目:{pname}）")
            # 条件称样匹配：按 weighing_rules 为每个样品匹配称样方式（运行时按样品分发用）
            if ctx.get("wmode") == "conditional" and ctx.get("weighing_rules"):
                wm = _match_conditional_rule(ctx["weighing_rules"], pname, it.get("pdf_paths") or [], desc)
                it["_wmode"] = str((wm or {}).get("weighing_mode") or "").strip()
                if it["_wmode"]:
                    log(f"[{it['sample_code']}] 命中称样规则 → {it['_wmode']}（项目:{pname}）")
                else:
                    log(f"[{it['sample_code']}] 未命中称样规则（项目:{pname}），将跳过该样品称样")

        rules = self._read_query_rules(method_file)
        plan = _plan_submission_batches(rules, all_items)

        codes = []
        for b in plan:
            if self._aborted(idx):
                return "abort"
            sb = [it["sample_code"] for it in b["items"]]
            dw = f"〔称样{b['wdate']}〕" if b.get("wdate") else ""
            eq_info = f"〔设备:{b.get('_equipment')}〕" if b.get("_equipment") else "〔默认设备〕"
            lp_info = f"〔过程:{b.get('_lab_proc')}〕" if b.get("_lab_proc") else ""
            if b["switch_mid"]:
                log(f"切换到方法ID {b['switch_mid']} {eq_info}{lp_info}（{len(sb)} 个样品: {', '.join(sb)}）{dw}")
            else:
                log(f"未匹配切换规则（{len(sb)} 个样品）{dw} {eq_info}{lp_info}，用默认方法")
            ok, code = self._submit_batch(idx, row, b["items"], ctx, log, b["force_new"], b["switch_mid"], b.get("_lab_proc") or "")
            if not ok:
                return "fail"
            codes.append(code)

        real_code = " / ".join(codes)
        self._ui_q.put(("status", (idx, "成功", "")))
        self._ui_q.put(("rowdata", (idx, {
            "experiment_code": real_code,
            "samples_total": len(samples),
            "samples_merged": len(samples) - n_skip,
            "skipped_list": list(skipped_samples),  # [(code, reason), ...] 供 _on_run_done 跨行汇总
            "early_analysis": ctx.get("_early_analysis", []),  # [(code,分析日,受理日)] 分析时间<受理时间，供运行报告汇总
        })))
        extra = f"，跳过 {n_skip} 个样品" if n_skip else ""
        log(f"成功，{len(samples) - n_skip}/{len(samples)} 个样品参与合并，实验编号 {real_code}{extra}")
        if skipped_samples:
            log(f"未录入样品（{len(skipped_samples)} 个）：")
            for sc, reason in skipped_samples:
                log(f"  - {sc}：{reason}")
        return "ok"

    def _collect_sample_projects(self, idx, row, sample_code, pdf_paths, ctx, log, projects_by_sample=None):
        """阶段A：查询样品 + 按方法过滤（谱图上传推迟到清空旧数据之后，避免重跑重复上传）。
        返回 (items, err)。items = [{project, sample_code, sample_id, pdf_paths}, ...]。
        err 非空表示该样品不可用(查不到/无匹配项目)，调用方跳过该样品。
        projects_by_sample: 批量查询缓存 {sampleCode: [project...]}，提供则不再逐样品查 LIMS(省往返)。"""
        _days = (ctx or {}).get("date_window_days", 30)  # 与方法池/逐样品共用同一窗口
        if projects_by_sample is not None:
            projects = list(projects_by_sample.get(sample_code) or [])
            if not projects:  # 方法池未含该样品(方法关键字与服务端索引不一致)，精确查询兜底，防漏查
                log(f"方法池未含 {sample_code}，精确查询兜底 ...")
                projects = self.api.query_samples_by_conditions(
                    sample_code=sample_code, exact_match=True, log_func=log, days=_days)
        else:
            log(f"查询样品 {sample_code} ...")
            projects = self.api.query_samples_by_conditions(
                sample_code=sample_code, exact_match=True, log_func=log, days=_days)
        if not projects:
            return [], self.api.diagnose_missing_sample(sample_code, log, days=_days)
        projects, ferr = self._filter_projects_by_method(projects, row, log)
        if ferr:
            # 方法过滤失败：可能该样品的这些项目已登记(在 ALREADY 列表，未登记查询不返回)。
            # 复用同一过滤逻辑查 ALREADY 列表，命中则报"已登记"，避免误报"方法不在此样品中"。
            try:
                _done = self.api.query_samples_by_conditions(
                    sample_code=sample_code, exact_match=True, log_func=None, days=_days,
                    check_in_status="CHECK_IN_STATUS_ALREADY")
                if _done:
                    _filt, _ferr2 = self._filter_projects_by_method(_done, row, lambda *a, **k: None)
                    if _filt:
                        return [], "已登记"
            except Exception:
                pass
            return [], ferr
        if not projects:
            return [], "该样品下没有匹配方法文件的项目"
        sample_id = projects[0].get("sampleId")
        # 样品号统一用 LIMS 全码(project.sampleCode = detectionNo+smallNo)，不用传入的报验编号
        # (可能截断，如手填/旧序列残留的 TS26072901)，否则后续 wmap/n_par 按全码键查会 miss
        # (误报"称量记录平行不足")。与 4275 _full_code 同源。
        items = [{"project": p, "sample_code": p.get("sampleCode") or sample_code,
                  "sample_id": sample_id, "pdf_paths": pdf_paths} for p in projects]
        return items, None

    def _submit_batch(self, idx, row, batch_items, ctx, log, force_new, switch_mid="", lab_proc=""):
        """提交一批 projects（worker 线程内，跨样品合并）。返回 (ok, real_code)，失败自行 put status。
        batch_items: [{project, sample_code, sample_id, spectrum_uploaded}, ...]，含多个样品的
        同方法项目；同一目标方法(to_id)的不同样品项目合并到一个实验编号。
        switch_mid: 本批目标方法ID(按项目名匹配 default_rules)；空则不切换。
        force_new: 多批时传 True，使每批生成独立实验编号(避免同号冲突)。
        lab_proc: 本批条件实验过程(命中 lab_proc 规则)；非空则覆盖 experimentProcess，空用方法配置默认。"""
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
            method_id = switch_mid
            actual_method_name = method_name
            actual_method_id = method_id
        else:
            # 解析方法文件规则的子方法（如 'AfPS GS 2019:01 PAK 单组份' → methodId 4489），
            # 首次取配置即带 detectionMethod.id：一个标准号下多子方法（单组份/N项之和）若不带 id，
            # 服务端无法定位 → 报 005"样品项目对应的方法不同"
            sub_method_id = ""
            _qi = batch_projects[0].get("_qr_idx")
            if _qi is not None:
                _qr = self._read_query_rules(ctx.get("method_file") or "")
                if _qi < len(_qr):
                    _rm = str(_qr[_qi].get("method") or "").strip()
                    if _rm and not _rm.isdigit():
                        sub_method_id = self.api.get_method_id_by_standard_no_name(_rm, log) or ""
            log("取实验配置 ...")
            log(f"取实验配置: sp_ids={sp_ids_str} method={method_name!r} sample_id={sample_id!r}"
                + (f" subMethodId={sub_method_id}" if sub_method_id else ""))
            initial = self.api.get_experiment_config(sp_ids_str, method_name, "", sample_id, log, sub_method_id or None)
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
        log("取全量配置 ...")
        all_cfg = self.api.get_all_configs(sp_ids_str, method_name, "", sample_id, log, method_id, project_names)
        if not all_cfg:
            self._ui_q.put(("status", (idx, "失败", "取全量配置失败")))
            return False, ""
        experiment_config = all_cfg.get("experiment")
        # 平行样扩展：称样记录平行数多于 LIMS 记录数时，复制各组分记录补平行(实验次数=平行数)
        pid_to_sample = {str(it["project"].get("projectId")): it["sample_code"] for it in batch_items}
        _pname_filter = {r["project"]: r["filter"]
                         for r in (self._read_spectrum_filter_rules(ctx["method_file"]) or [])}
        _par_by = {}
        # 已知后缀(如 T/TS)：random/无称样记录模式下称量记录不编码平行数，改由谱图文件名 A/B 推断
        _scp = self._read_spectrum_check_params(ctx["method_file"]) or {}
        _psw = _scp.get("sample") or {}
        _known_suf = ({s.replace("*", "").lower() for s in (_psw.get("keyword") or "").replace("，", ",").split(",") if s.strip()}
                      if _psw.get("enabled") else set())
        _pdfs_by_sc = {}
        for _it in batch_items:
            _pdfs_by_sc.setdefault(_it["sample_code"], []).extend(_it.get("pdf_paths") or [])
        for _sc in dict.fromkeys(it["sample_code"] for it in batch_items):
            _mm = ((ctx["wmap"] or {}).get(_sc) or {}).get("masses") or []
            _n = len(_mm)
            if _n <= 1 and _known_suf:  # 称量记录未编码平行数 → 由文件名 A/B 推断
                _n = _pdf_parallel_count(_sc, _pdfs_by_sc.get(_sc), _known_suf)
            if _n > 1:
                _par_by[_sc] = _n
        # 标记项目(谱图filter=标记)按其标记称样量数扩展，不跟随基样 A/B 平行数(如 5mm以内 仅1个M)
        _par_by_pid = {}
        for it in batch_items:
            _pid = str(it["project"].get("projectId"))
            _flt = _pname_filter.get(it["project"].get("projectName", ""), "")
            _mmk = ((ctx["wmap"] or {}).get(it["sample_code"]) or {}).get("marker_masses") or {}
            if _flt and _mmk.get(_flt):
                _par_by_pid[_pid] = len(_mmk[_flt])
        _old_n, _new_n = _expand_parallel_records(experiment_config, pid_to_sample, _par_by, _par_by_pid)
        if _new_n > _old_n:
            log(f"平行样扩展：ocAnalysisRecordList {_old_n}→{_new_n} 条(按称样记录平行数补平行)")
        equipment_config = all_cfg.get("equipment")
        # 设备以序列表格为准：表格指定设备编号(可 ';' 多个)则覆盖默认主检设备；
        # 条件设备模式取批内 item 的 _equipment（所有 item 同批次故设备一致）
        batch_eq = (batch_items[0].get("_equipment") or "") if batch_items else ""
        override_eq = batch_eq or row.get("equipment", "")
        equipment_config, matched_list, eq_err = self._override_equipment(equipment_config, override_eq)
        if eq_err:
            self._ui_q.put(("status", (idx, "失败", eq_err)))
            log(f"失败: {eq_err}")
            return False, ""
        equipment_config = self._enrich_weighing_bill(equipment_config, sp_ids_str, log)
        # 标液：条件标液模式按批内 item 的 _standard 解析（同批同标液）；其余模式用行级 configure_order
        _co = ctx.get("configure_order") or ""
        if ctx.get("standard_type") == "conditional":
            _stds = [str(it.get("_standard") or "").strip() for it in batch_items
                     if str(it.get("_standard") or "").strip()]
            _co = _stds[0] if _stds else ""
            if _co:
                log(f"本批条件标液: {_co}")
                # 条件标液行级预检未覆盖(行级 configure_order 为空)，此处逐个校验
                _stime = next(((_e or {}).get("time") for _e in (ctx.get("wmap") or {}).values()
                               if (_e or {}).get("time")), None) or _prev_workday()
                for _c in [x.strip() for x in _co.replace("，", ",").split(",") if x.strip()]:
                    cid, emsg = self.api.get_solution_configure_id(_c, log, _stime)
                    if not cid:
                        self._ui_q.put(("status", (idx, "失败", f"标准溶液:{emsg}")))
                        log(f"失败: 标准溶液校验未过 - {emsg}")
                        return False, ""
            else:
                log("警告: 条件标液模式但本批未匹配到标液编号，跳过标液关联")
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
            actual_method_name, actual_method_id, _co, ctx["fixed_params"],
            dilution_remark_enabled=ctx.get("dilution_remark_enabled", False),
            dilution_volume_column=ctx.get("dilution_volume_column", ""))

        # 温湿度：表格手填优先 → 设备房间·当天自动(温湿度记录.xlsx) → 默认 22/55
        temp_val = (row.get("temperature") or "").strip()
        hum_val = (row.get("humidity") or "").strip()
        if not temp_val or not hum_val:
            e2r, renv, _ = _read_env_records(_ENV_RECORD_PATH)
            if e2r is not None:
                cands = _equipment_env_candidates(row, equipment_config, matched_list)
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
                _names = [str(c.get("columeName") or "") for c in dynamic_columns if isinstance(c, dict)]
                log(f"未找到试样信息/试样描述列，跳过填充；本方法动态列: {_names}")
                return
            info_code = info_col.get("columeCode", "")
            # projectId 未命中 pid_to_sample 时回落到批次首个样品描述(单样品批次如 XRF 三方法共享一样品)
            _fallback_sc = batch_samples[0] if batch_samples else ""
            host.data_fields[info_code] = [
                _Box(desc_by_sample.get(pid_to_sample.get(str(r.get("projectId")), "") or _fallback_sc, ""))
                for r in records or []]
            _filled = sum(1 for b in host.data_fields[info_code] if b.get())
            log(f"试样描述({info_col.get('columeName', '')}) 已按样品填充 {_filled}/{len(records or [])} 条")

        # 跨项目复用(组分→总和)：首项目生成的试样描述/称样量缓存后，后续项目直接复用，保证同样品一致
        # 免去 getOcCompareShowData 网络往返；query_rules 顺序保证组分先于总和生成
        primary_cache = ctx["primary_cache"]
        cached_before = set(primary_cache)  # 本批开始前已缓存的样品 → 这些样品将复用缓存

        # 称样量注入：统一按样品分发（支持 conditional：各样品按命中规则走不同称样方式）
        base_wmode = wmode
        _cond = (base_wmode == "conditional")
        if base_wmode == "none" and not _cond:
            log("称样量模式=无需称样量，跳过称样列")
        mass_col = _find_weighing_column(dynamic_columns)
        mass_code = mass_col.get("columeCode", "") if mass_col else ""
        records = (experiment_config or {}).get("ocAnalysisRecordList") or []
        # 各样品平行数：projectId->样品 的最大 serialNumber；混批各样品可不同(record/process/random 共用)
        _n_par_by_sample = {}
        for _r in records:
            _rsc = pid_to_sample.get(str(_r.get("projectId")))
            if not _rsc:
                continue
            try:
                _rsv = int(_r.get("serialNumber"))
            except (TypeError, ValueError):
                _rsv = 1
            _n_par_by_sample[_rsc] = max(_n_par_by_sample.get(_rsc, 0), _rsv)
        _prules = self._read_processing_rules(ctx["method_file"])
        _prule = _match_processing_rule(_prules, actual_method_name)  # random/process 共用同一匹配规则
        try:
            _mdp = int((wp or {}).get("decimal_places") or 4)  # 天平标准4位(0.0001g)
        except (TypeError, ValueError):
            _mdp = 4
        # marker 路由(与方式无关)：projectId -> 谱图filter；filter 为标记字母(如 M)时取标记称样量
        pid_filter = {str(it["project"].get("projectId")):
                      _pname_filter.get(it["project"].get("projectName", ""), "")
                      for it in batch_items}
        pmasses_by_sample, marker_masses_by_sample, desc_by_sample = {}, {}, {}
        _sp = (row.get("spectrum_path") or "").strip()

        for it in batch_items:
            sc = it["sample_code"]
            if sc in pmasses_by_sample:
                continue
            eff = (it.get("_wmode") or "") if _cond else base_wmode  # 条件称样按样品命中；否则整批统一
            samp = (ctx["wmap"] or {}).get(sc) or {}
            rv = primary_cache.get(sc)
            if rv:
                pmasses_by_sample[sc] = list(rv["masses"])
                marker_masses_by_sample[sc] = dict(rv.get("marker_masses") or {})
                desc_by_sample[sc] = rv["desc"]
                continue

            if eff == "random":
                # 平行数取本样品的(与 record/process 一致)，混批下不因他样多平行而丢弃本样记录值
                _n_par = _n_par_by_sample.get(sc, 1)
                pmasses_by_sample[sc] = _random_masses(samp.get("masses") or [], _n_par, _prule, wp, _mdp)
                desc_by_sample[sc] = samp.get("desc") or ""
            elif eff in ("record", "process"):
                masses = samp.get("masses")
                n_par = _n_par_by_sample.get(sc, 1)
                if not masses or len(masses) < n_par:
                    self._ui_q.put(("status", (idx, "失败", f"称量记录平行不足({len(masses or [])}/{n_par})")))
                    log(f"失败: 样品 {sc} 称量记录仅 {len(masses or [])} 个平行，实验需 {n_par}")
                    return False, ""
                raw_marker = samp.get("marker_masses") or {}
                if eff == "record":
                    # 按方法 decimal_places 格式化，保留末尾0(0.552→0.5520)；不再用 %g 吞末尾0
                    pmasses_by_sample[sc] = [f"{float(masses[i]):.{_mdp}f}" for i in range(n_par)]
                    marker_masses_by_sample[sc] = {mk: [f"{float(v):.{_mdp}f}" for v in mv]
                                                   for mk, mv in raw_marker.items()}
                else:
                    pmasses_by_sample[sc] = [_apply_processing(masses[i], _prule, wp) for i in range(n_par)]
                    marker_masses_by_sample[sc] = {mk: [_apply_processing(v, _prule, wp) for v in mv]
                                                   for mk, mv in raw_marker.items()}
                desc_by_sample[sc] = samp.get("desc") or ""
            elif eff == "pdf":
                # PDF报告：称样量取自报告(样品初始质量)，按 decimal_places 保留末尾0(0.31→0.3100)
                from report_parser import parse_pdf_report_meta, filter_samples_by_code
                field = ((mass_col or {}).get('equipRelativeTitle') or '').strip() or '样品初始质量'
                pdf, _dil = _pick_sample_report_pdf(_sp, sc, actual_method_name)
                raw = []
                if pdf:
                    try:
                        meta = filter_samples_by_code(parse_pdf_report_meta(pdf, field), sc)
                        raw = [v for _sid, v in meta if v]
                    except Exception as e:
                        log(f"称样量(pdf) 样品 {sc} 报告解析失败({e})")
                else:
                    log(f"称样量(pdf) 样品 {sc} 谱图目录未找到报告 PDF，称样量留空")
                pmasses_by_sample[sc] = [f"{float(v):.{_mdp}f}" for v in raw]  # 报告段顺序=平行槽(A→0,B→1)
                desc_by_sample[sc] = samp.get("desc") or ""
            elif eff in ("none", ""):
                # 无需称样 / 条件未命中：不产称样量，仅留描述
                pmasses_by_sample[sc] = []
                desc_by_sample[sc] = samp.get("desc") or ""
                if eff == "":
                    log(f"样品 {sc} 条件称样未命中任何规则，称样量留空")
            else:
                log(f"称样量模式={eff} 暂未接入(本轮支持 random/none/record/process/pdf)")
                pmasses_by_sample[sc] = []
                desc_by_sample[sc] = samp.get("desc") or ""

            primary_cache[sc] = {"masses": list(pmasses_by_sample[sc]),
                                 "marker_masses": dict(marker_masses_by_sample.get(sc) or {}),
                                 "desc": desc_by_sample[sc]}

        # 写入称量记录列(所有方式共用)；纯 none 模式不写称样列(保持原行为)
        _skip_mass = (base_wmode == "none" and not _cond)
        if not _skip_mass and mass_col is not None:
            host.data_fields[mass_code] = _mass_field_by_project(
                records, pid_to_sample, pmasses_by_sample, pid_filter, marker_masses_by_sample)
            if _cond:
                from collections import Counter
                _cnt = Counter((it.get("_wmode") or "未命中") for it in batch_items)
                log(f"称样量(条件) {mass_col.get('columeName', '')} 跨{len(pmasses_by_sample)}样品 方式分布{dict(_cnt)}")
            else:
                log(f"称样量({base_wmode}) {mass_col.get('columeName', '')} 跨{len(pmasses_by_sample)}样品")
        elif not _skip_mass:
            log(f"称样量({base_wmode}) 未找到称量记录列(isWeighing=1)，跳过称样量写入")
        _fill_desc_column(records, desc_by_sample)
        first_sc = batch_items[0]["sample_code"] if batch_items else ""
        analysis_start = ((ctx["wmap"] or {}).get(first_sc) or {}).get("time")

        _reused = [sc for sc in batch_samples if sc in cached_before]
        if _reused:
            log(f"复用首项目(组分)试样描述/称样量: {len(_reused)}/{len(batch_samples)} 样品")

        # 结果值回填：是否总和由匹配到的 query rule 的 sum_entry 标记决定(方法编辑器勾选)。
        # 勾选=是 → 读各组分已录入值(getOcCompareShowData)填组分列、服务端求和；否 → 按报告解析 PDF 浓度写入结果列
        _qr = self._read_query_rules(ctx.get("method_file") or "")
        _sum_on = False
        for _p in batch_projects:
            _qi = _p.get("_qr_idx")
            if isinstance(_qi, int) and 0 <= _qi < len(_qr) and bool(_qr[_qi].get("sum_entry")):
                _sum_on = True
                break
        if _sum_on:
            self._fill_result_from_sum(host, dynamic_columns, experiment_config,
                                       batch_items, row, actual_method_name, ctx, log)
        else:
            self._fill_result_from_report(host, dynamic_columns, experiment_config,
                                          batch_items, row, actual_method_name, ctx, log)
        # 启用稀释备注但未启用报告解析：扫稀释 PDF 文件名(-NNX)填稀释列(报告解析启用时不扫，已更精确填过)
        if ctx.get("dilution_remark_enabled"):
            self._fill_dilution_from_filename(host, dynamic_columns, experiment_config,
                                              batch_items, row, actual_method_name, ctx, log)

        # 构造载荷 + 注入谱图 + 提交
        log("构建并提交 ...")
        experiment_data = build_grouped_experiment_data(host, batch_projects, experiment_code, actual_method_name, experiment_process_override=lab_proc or None)
        log(f"[设备] 提交称样设备对象: {experiment_data.get('weighingEquipment')}")
        if lab_proc:
            log(f"实验过程(experimentProcess) <- 条件规则: {lab_proc}")
        # 称量记录「称样时间」→ 覆盖实验分析开始时间 startTime；无有效称样时间则开始时间=结束时间
        _start_set = False
        if analysis_start is not None:
            try:
                experiment_data["startTime"] = analysis_start.strftime("%Y-%m-%d %H:%M:%S")
                log(f"分析开始时间(startTime) <- 称样时间: {experiment_data['startTime']}")
                _start_set = True
            except (AttributeError, ValueError):
                log(f"警告: 称样时间格式无法解析({analysis_start!r})")
        if not _start_set:
            if ctx.get("defaults_no_record"):
                # 无称样记录默认模式：startTime = max(上一工作日, 批内最晚受理日期) + 00:00:00
                # 服务端005校验"分析时间不能早于受理时间"(按日期比较，与前端一致：受理日 00:00:00)
                _prev = _prev_workday()
                _latest, _hit_k = _prev, ""
                for _it in batch_items:
                    _d, _k = _accept_date_of((_it.get("project") or {}).get("_raw"))
                    if _d and _d > _latest:
                        _latest, _hit_k = _d, _k
                experiment_data["startTime"] = _latest.strftime("%Y-%m-%d") + " 00:00:00"
                _src = f"受理日(字段{_hit_k})" if _hit_k else "上一工作日(未取到受理时间)"
                log(f"无称样时间(默认模式)，开始时间(startTime) <- max(上一工作日,批内最晚受理日): "
                    f"{experiment_data['startTime']} 〔{_src}〕")
                if not _hit_k and batch_items:  # 未命中受理时间字段→dump raw keys 以便精准修正
                    _ks = list((batch_items[0].get("project") or {}).get("_raw") or {})
                    log(f"警告: 未取到受理时间字段，raw keys={_ks}")
            else:
                # 无有效称样时间：开始时间取结束时间
                experiment_data["startTime"] = experiment_data["endTime"]
                log(f"无称样时间，开始时间(startTime) <- 结束时间: {experiment_data['endTime']}")
        # 服务端005校验：分析时间(startTime)不得早于受理时间。startTime 须 ≥ 批内最晚受理日，
        # 否则把日期提到该受理日(保留称样时分)，记入运行报告。defaults_no_record 模式 startTime 已 ≥ 受理日，不触发。
        try:
            _start_d = date.fromisoformat(str(experiment_data.get("startTime"))[:10])
        except ValueError:
            _start_d = None
        if _start_d is not None:
            _accs = []  # [(sample_code, accept_date)]
            for _it in batch_items:
                _ad, _ = _accept_date_of((_it.get("project") or {}).get("_raw"))
                if _ad:
                    _accs.append((_it.get("sample_code"), _ad))
            _latest = max([_d for _, _d in _accs], default=_start_d)  # 批内最晚受理日
            if _latest > _start_d:
                _orig = experiment_data["startTime"]
                _time_part = str(_orig)[10:]                          # " HH:MM:SS" 保留称样时分
                experiment_data["startTime"] = _latest.strftime("%Y-%m-%d") + _time_part
                _violators = [(_sc, _orig, experiment_data["startTime"], _ad.isoformat())
                              for _sc, _ad in _accs if _ad > _start_d]
                log(f"分析时间早于受理时间，startTime 已调整: {_orig} → {experiment_data['startTime']}"
                    f"（提到最晚受理日 {_latest}，涉及 {len(_violators)} 个样品）")
                ctx.setdefault("_early_analysis", []).extend(_violators)
        # 注入谱图 fileIds/spectrumJsonList：每个谱图只绑定到「实际用到它的项目」(按 fileId 归并 projectId)。
        # 标记分流时(总和/苯并[a]芘用基样谱、5mm以内用M谱)，不可把整样所有谱图绑到全部 projectId，
        # 否则 5mm以内 会同时挂上基样谱与M谱。同谱图被多项目复用则 projectId 取并集。
        fid_info = {}  # fileId -> {"fileName":..., "pids":[projectId...有序去重]}
        for it in batch_items:
            pid = str(it["project"].get("projectId"))
            for f in it["spectrum_uploaded"] or []:
                fid = f.get("fileId")
                if not fid:
                    continue
                if fid not in fid_info:
                    fid_info[fid] = {"fileName": f.get("fileName"), "pids": []}
                if pid not in fid_info[fid]["pids"]:
                    fid_info[fid]["pids"].append(pid)
        spec_list = [{"fileId": fid, "fileName": v["fileName"], "projectId": ",".join(v["pids"])}
                     for fid, v in fid_info.items()]
        if spec_list:
            experiment_data["fileIds"] = ",".join(str(fid) for fid in fid_info)
            experiment_data["spectrumJsonList"] = json.dumps(spec_list)
        # ponytail: 服务端实验编号=lqy+秒级时间戳，同秒多批撞号(本地 experimentCode 被忽略)。
        # 保证每批 saveOcExperiment 距上一批≥1s，使服务端生成不同编号。若服务端改为支持客户端唯一编号，可移除此节流。
        _last = getattr(self, "_last_exp_submit_ts", None)
        _now = time.time()
        if _last and _now - _last < 1.0:
            time.sleep(1.0 - (_now - _last) + 0.15)
        self._last_exp_submit_ts = time.time()
        _require_sign = bool((self._read_other_params(ctx.get("method_file") or "") or {}).get("require_signature"))
        log(f"提交实验数据 ({'submitOcExperiment/提交签名' if _require_sign else 'saveOcExperiment/仅保存'}) ...")
        ok, _ = self.api.submit_experiment_data(experiment_data, actual_method_name, log, require_signature=_require_sign)
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

        # 主检设备 + 称样设备均随 saveOcExperiment 载荷提交（mainEquipment / weighingEquipment 字段），
        # 对齐网页端——网页端不单独调 saveMainEqubment。称样设备经 _enrich_weighing_bill 回填真实台账。
        _we = equipment_config or {}
        if _we.get("weighingEquipmentId"):
            log(f"[设备] 称样设备随实验载荷提交: {_we.get('weighingEquipment')} (id={_we.get('weighingEquipmentId')})")
        if (_we.get("mainEquipmentIds") or "").strip():
            log(f"[设备] 主检设备随实验载荷提交: ids={_we.get('mainEquipmentIds')}")

        # 标液关联——用真实编号（对照 detection_entry_main:1768-1776）
        if _co:
            experiment_codes = {pid: real_code for pid in project_ids}
            sok, serr = self.api.submit_solution_with_experiment(
                project_ids, _co, experiment_codes, log)
            if not sok:
                self._ui_q.put(("status", (idx, "失败", f"标准溶液关联:{serr}")))
                log(f"失败: 标准溶液关联失败 - {serr}")
                return False, ""

        log(f"本批完成，实验编号 {real_code}")
        return True, real_code

    def _resolve_samples(self, row, wmap, log, defaults_no_record=False):
        """解析本行要录入的样品。返回 (samples, err)。
        samples = [(sample_code, [pdf_path,...]), ...]；err 非空=无法解析。
        - 谱图是文件 / 目录单PDF / 目录+已填样品编号：单样品
        - 目录+多PDF+未填样品编号：用 wmap(称样记录)样品编号 ∩ 目录PDF(startswith)展开；
          wmap 缺失则 err。一个样品匹配多个PDF(如A/B平行)全部收集、都上传。
        - defaults_no_record=True(无称样记录默认模式)：wmap 缺失时不报错，目录内每个编号各成一样品。"""
        sp = (row.get("spectrum_path") or "").strip()
        sc = (row.get("sample_code") or "").strip()
        if os.path.isfile(sp):
            if sc:
                code = sc
            else:
                # 行无样品编号时优先用 PDF 内容 `样品 :` 字段(文件名号可能与内容不一致)
                from report_parser import extract_content_sample_ids
                cids = extract_content_sample_ids(sp)
                code = cids[0] if cids else os.path.splitext(os.path.basename(sp))[0].split("-", 1)[0]
            return [(code, [sp])], None
        pdfs = sorted(glob.glob(os.path.join(sp, "*.pdf"))) if os.path.isdir(sp) else []
        if not pdfs:
            return [], "谱图路径无效或目录内无PDF"
        # 目录=多样品意图。sample_code 是隐藏的自动回填字段(前端无单元格、不可编辑)，
        # 常为旧序列残留，不应在目录模式强制单样品。清空它，统一走下方 单PDF / 称样记录∩谱图 展开。
        if sc:
            log(f"忽略样品编号残留 '{sc}'，按谱图目录({len(pdfs)}个PDF)展开全部样品")
            row["sample_code"] = ""
            sc = ""
        if len(pdfs) == 1:
            code = os.path.splitext(os.path.basename(pdfs[0]))[0].split("-", 1)[0]
            return [(code, [pdfs[0]])], None
        # 多PDF + 未填样品编号 → 用称样记录编号 ∩ 目录PDF 展开(只有两边都有的编号才参与录入)
        if not wmap:
            if defaults_no_record:
                # 无称样记录默认模式：目录内每个编号(按文件名前缀去平行小号)各成一样品，全部参与录入
                _by_code = {}
                for _p in pdfs:
                    _c = os.path.splitext(os.path.basename(_p))[0].split("-", 1)[0]
                    _by_code.setdefault(_c, []).append(_p)
                return [(_c, _ps) for _c, _ps in _by_code.items()], None
            return [], "目录有多个PDF且未填样品编号，请在「称样记录路径」填称样记录excel(按 excel∩谱图 展开)"
        # 先收集有谱图的候选 (编号, 试样描述, 匹配PDF)
        from report_parser import extract_content_sample_ids
        content_index = {p: extract_content_sample_ids(p) for p in pdfs}
        cands = []
        for code, entry in wmap.items():
            key = _strip_parallel_suffix(code)  # 报验编号(去小号)；简写文件名(无小号)代表001时用它
            # 按小号匹配(精确/平行字母/简写)，避免 startswith(key) 吃进同报验号其它小号
            matched = [p for p in pdfs
                       if _code_belongs_sample(
                           os.path.splitext(os.path.basename(p))[0].split("-", 1)[0], code, key)
                       or any(_code_belongs_sample(cid, code, key) for cid in content_index.get(p, ()))]
            if matched:
                cands.append((code, (entry or {}).get("desc") or "", matched))
        if not cands:
            return [], "称样记录中的样品编号在谱图目录内均无匹配PDF"
        # 混目录分流：方法 switch_rules 含 desc 时，只保留试样描述命中本方法的样品
        # (固体/液体各走各自方法，避免液体样品混进固体方法卡在谱图检查)
        switch_rules = self._read_switch_rules(row.get("method_file"))
        if any(str(r.get("desc") or "").strip() for r in switch_rules):
            kept, skipped = [], []
            for code, desc, matched in cands:
                if _rule_desc_match(switch_rules, desc):
                    kept.append((code, matched))
                else:
                    skipped.append(code)
            if not kept:
                return [], "称样记录中无试样描述命中本方法(switch_rules.desc)的样品，请检查方法与称样记录"
            if skipped:
                log(f"样品分流(按 switch_rules.desc)：跳过 {len(skipped)} 个：{'、'.join(skipped)}")
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
        if not method_file:
            return []
        try:
            y = load_method(method_file)
        except Exception:
            return []
        qr_raw = y.get("query_rules")
        if isinstance(qr_raw, dict):
            return qr_raw.get("query_rules") or []
        if isinstance(qr_raw, list):
            return qr_raw
        return []

    def _read_sample_parallel_threshold(self, method_file):
        """query_rules.sample_parallel_threshold(按样品并行 vs 方法池 阈值)。默认 20。
        样品数≤阈值→按样品并行精确查；>阈值→方法池(1请求)摊销。"""
        try:
            y = load_method(method_file)
        except Exception:
            return 20
        qr = y.get("query_rules")
        if isinstance(qr, dict):
            try:
                n = int(qr.get("sample_parallel_threshold"))
                return n if n > 0 else 20
            except (TypeError, ValueError):
                pass
        return 20

    def _read_date_window_days(self, method_file):
        """query_rules.date_window_days(查询受理日期窗口)。默认 30。方法池与逐样品精确查共用。"""
        try:
            y = load_method(method_file)
        except Exception:
            return 30
        qr = y.get("query_rules")
        if isinstance(qr, dict):
            try:
                n = int(qr.get("date_window_days"))
                return n if n > 0 else 30
            except (TypeError, ValueError):
                pass
        return 30

    def _method_query_names(self, method_file):
        """方法文件 query_rules[].method 的非空文本方法名，去重保序。纯数字跳过(不支持方法ID查询)。"""
        names = []
        for q in self._read_query_rules(method_file):
            mv = str(q.get("method") or "").strip()
            if mv and not mv.isdigit() and mv not in names:
                names.append(mv)
        return names

    def _method_query_key(self, name):
        """方法查询关键字：剥掉标准号尾部的子方法后缀(如 ' XRF'/' 附录A'/' 单组份')，只用标准号主体查询。
        服务端 decideProjectMethodName 对 standardNo+' '+subMethod 做子串匹配，带子方法后缀的查询串常子串不命中、整方法漏查
        (实测 '...7-2:2017 XRF' 查不到六价铬，'...7-2:2017  XRF' 才能查到；GB/T 27947-2020 附录A 同理查不到，
        因方法定义把'附录A'写进 standardNo、而样品 standardNo 仅主干)。用 standardNo 主体(其前缀)查询必命中且是超集，
        _filter_projects_by_method 再按 standardNo+子方法ID 精准过滤。
        以真实'年份'(19xx/20xx)为分界，剥其后非年份 token：兼容 ':' 年份(IEC ...:2017) 与 '-' 年份(GB/T ...-2020)；
        'IEC62321XRF'/'62321' 等非年份数字不误判为分界(否则汞方法的 IEC62321XRF 尾巴剥不掉)。"""
        yr = r"(?:19|20)\d{2}"
        if not re.search(yr, name):
            return name
        parts = name.split()
        while len(parts) > 1 and not re.search(yr, parts[-1]):
            parts.pop()
        return " ".join(parts).strip() or name

    def _query_projects_by_methods(self, names, log, days=30):
        """逐规则查询待登记项目入池(命中缓存直接用)，按 projectId 合并去重。返回 dict(p) 副本，避免跨行 _qr_idx 互染。
        先用完整方法名(含子方法后缀)精确查——结果少、快；若 0(服务端 standardNo 末尾空格拼接出双空格、如六价铬 7-2:2017，
        或子方法名与样品侧不完全一致)，回退到标准号主体(_method_query_key 剥后缀)宽查——必命中但是超集，
        _filter_projects_by_method 再按 standardNo+子方法ID 精准过滤。
        days: 方法池受理日期窗口(按方法文件可配)，默认30；仅本方法池查询用。"""
        cache = self._method_projects_cache
        pooled, seen = [], set()
        for name in names:
            core = self._method_query_key(name)
            keys = [name, core] if name != core else [core]  # 先精确(full)，0 再宽查(core)
            chosen = None
            for key in keys:
                projects = cache.get(key)
                if projects is None:
                    projects = self.api.query_samples_by_conditions(
                        method_name=key, log_func=log, days=days) or []
                    cache[key] = projects
                    log(f"按方法查询 {key}：{len(projects)} 个项目")
                if projects:
                    chosen = projects
                    break
            for p in (chosen or []):
                pid = p.get("projectId")
                if pid not in seen:
                    seen.add(pid)
                    pooled.append(dict(p))
        return pooled

    def _query_samples_parallel(self, sample_codes, log, max_workers=6, days=30):
        """并行按样品编号精确查(I/O 密集、各样品独立)，返回 {sampleCode: [project...]}。
        days: 受理日期窗口(天)，默认30；与方法池共用 date_window_days 配置，按方法文件收窄/放宽。
        单样品异常不影响其余。worker 内静默，日志由调用方汇总。"""
        from concurrent.futures import ThreadPoolExecutor
        out = {}
        codes = [c for c in sample_codes if c]
        if not codes:
            return out

        def _one(code):
            try:
                return code, self.api.query_samples_by_conditions(
                    sample_code=code, exact_match=True, log_func=None, days=days) or []
            except Exception:
                return code, []  # ponytail: 单样品失败静默，调用方按空结果跳过

        with ThreadPoolExecutor(max_workers=max_workers) as ex:  # ponytail: 固定6线程，>50样品再调
            for code, plist in ex.map(_one, codes):
                if plist:
                    out[code] = plist
        return out

    def _read_exclusion_rules(self, method_file):
        """读方法文件 query_rules.exclusion_rules（排除项目规则: {project, method}）。
        返回 list；空/异常/旧版 list 结构返回 []。"""
        if not method_file:
            return []
        try:
            y = load_method(method_file)
        except Exception:
            return []
        qr_raw = y.get("query_rules")
        if isinstance(qr_raw, dict):
            return qr_raw.get("exclusion_rules") or []
        return []

    def _read_spectrum_filter_rules(self, method_file):
        """读方法文件 spectrum_upload_settings.spectrum_filter_rules（按项目筛选谱图文件名）。
        返回 [{project, filter}, ...]；空/异常返回 []。"""
        if not method_file:
            return []
        try:
            y = load_method(method_file)
        except Exception:
            return []
        rules = ((y.get("spectrum_upload_settings") or {}).get("spectrum_filter_rules")) or []
        return [{"project": str(r.get("project") or ""), "filter": str(r.get("filter") or "")}
                for r in rules if isinstance(r, dict)]

    def _read_report_parse_settings(self, method_file):
        """读方法文件 spectrum_upload_settings 的报告解析设置。
        返回 {enabled, instrument_type, marker}；空/异常返回 {}。"""
        if not method_file:
            return {}
        try:
            y = load_method(method_file)
        except Exception:
            return {}
        su = y.get("spectrum_upload_settings") or {}
        return {
            "enabled": bool(su.get("report_parse_enabled")),
            "instrument_type": su.get("instrument_type"),
            "marker": su.get("marker"),
        }

    def _fill_result_from_sum(self, host, dynamic_columns, experiment_config,
                              batch_items, row, actual_method_name, ctx, log):
        """总和录入（批量）：读各组分已录入结果值(getOcCompareShowData)填入总和实验的组分动态列，
        供 build_grouped_experiment_data→calcTheValue 服务端按方法公式求和+修约。
        组分列 = 带 compareShowItemName 的动态列；items 由其派生；取值字段由 compareShowTitleName(报告值/计算值)决定。
        按样品×平行(serialNumber)匹配组分记录；缺组分留空不阻断。"""
        comp_cols = [c for c in (dynamic_columns or [])
                     if isinstance(c, dict) and str(c.get("compareShowItemName") or "").strip()]
        if not comp_cols:
            log("总和录入: 未找到带 compareShowItemName 的组分列，跳过")
            return
        # items：各列 compareShowItemName 按 @ 拆别名后扁平去重（服务端做精确匹配也能命中其一）
        items = []
        for c in comp_cols:
            for _a in str(c.get("compareShowItemName") or "").split("@"):
                _a = _a.strip()
                if _a and _a not in items:
                    items.append(_a)

        records = (experiment_config or {}).get("ocAnalysisRecordList") or []
        # 样品号用完整号 detectionNo+sampleSmallNo（与 build_grouped_experiment_data 同源），
        # 不能用报验编号 it["sample_code"]——getOcCompareShowData 按记录 sampleCode 匹配，报验编号对不上会返 0 条
        def _full_code(it):
            p = it.get("project") or {}
            return f"{p.get('detectionNo','')}{p.get('sampleSmallNo','')}".strip()
        batch_samples = list(dict.fromkeys(_c for _c in (_full_code(it) for it in batch_items) if _c))
        if not batch_samples:  # 回落到报验编号(旧字段)
            batch_samples = list(dict.fromkeys(it.get("sample_code") for it in batch_items if it.get("sample_code")))
        _fallback_sc = batch_samples[0] if batch_samples else ""

        def _rec_sample(rec):
            sc = str(rec.get("sampleCode") or "").strip()
            if sc:
                return sc
            d = str(rec.get("detectionNo") or "").strip()
            s = str(rec.get("sampleSmallNo") or "").strip()
            return (d + s) if d and s else _fallback_sc

        # 逐样品取组分对比数据，索引 {sample: {norm_projectName: {serialNumber: rec}}}
        # ponytail: 多样品并行查(单请求~6s，串行=N×6s)；session 并发同 _query_samples_parallel
        comp_by_sample = {}
        from concurrent.futures import ThreadPoolExecutor
        with ThreadPoolExecutor(max_workers=min(6, len(batch_samples) or 1)) as _ex:
            _fetched = list(_ex.map(lambda sc: (sc, self.api.get_oc_compare_show_data(sc, items, log)),
                                    batch_samples))
        for sc, cr in _fetched:
            idx = {}
            for r in cr or []:
                pn = _norm_cn(r.get("projectName"))  # NFKC 归一，兼容 @ 别名的全角/半角逗号等变体
                try:
                    sn = int(r.get("serialNumber")) if r.get("serialNumber") is not None else 1
                except (TypeError, ValueError):
                    sn = 1
                idx.setdefault(pn, {})[sn] = r
            comp_by_sample[sc] = idx
            log(f"总和录入: 样品 {sc} 取到 {len(cr or [])} 条组分对比记录（{len(items)} 别名/{len(comp_cols)} 列）")

        def _val(rec, col):
            title = str(col.get("compareShowTitleName") or "报告值")
            field = "calculatedValue" if "计算值" in title else "reportValue"
            try:
                sn = int(rec.get("serialNumber")) if rec.get("serialNumber") is not None else 1
            except (TypeError, ValueError):
                sn = 1
            sample_map = comp_by_sample.get(_rec_sample(rec), {})
            # 列的 @ 别名逐个归一后命中任一即取值
            for _a in str(col.get("compareShowItemName") or "").split("@"):
                _a = _a.strip()
                if not _a:
                    continue
                item_map = sample_map.get(_norm_cn(_a))
                if item_map:
                    cr = item_map.get(sn) or (next(iter(item_map.values())) if item_map else None)
                    if cr is not None:
                        return str(cr.get(field) or "")
            return ""

        for col in comp_cols:
            col_code = col.get("columeCode", "")
            vals = [_Box(_val(rec, col)) for rec in records]
            host.data_fields[col_code] = vals
            _filled = sum(1 for b in vals if b.get())
            log(f"总和录入: 组分列 {col_code}({col.get('compareShowItemName')}/{col.get('compareShowTitleName')}) 填 {_filled}/{len(vals)}")

    def _fill_result_from_report(self, host, dynamic_columns, experiment_config,
                                 batch_items, row, actual_method_name, ctx, log):
        """报告解析→浓度回填（批量）：方法勾选「启用报告解析」时，按各样品谱图目录的报告 PDF
        解析浓度并写入结果列 host.data_fields[res_col]，供 build_grouped_experiment_data 提交。
        PDF 缺失/解析失败→告警留空，不阻断批量。复用 parse_pdf_report + evaluate_alias + get_project_alias。
        主支持无组分列场景(PAHs：每记录一化合物/项目，evaluate_alias 单段取值)；
        多组分列场景按记录组分名匹配段，best-effort(报告格式未全面验证)。"""
        settings = self._read_report_parse_settings(ctx.get("method_file"))
        method_enabled = settings and settings.get("enabled")
        wmap = ctx.get("wmap") or {}
        if not method_enabled and not any(
            (wmap.get(it.get("sample_code") or "") or {}).get("force_parse")
            for it in batch_items
        ):
            return  # 方法未启用报告解析且无样品强制解析：保持现状(结果列留默认)
        spectrum_path = (row.get("spectrum_path") or "").strip()
        if not spectrum_path:
            log("报告解析: 本行未设置谱图文件路径，跳过浓度回填")
            return

        # 1) 按样品解析报告 PDF（正常 + 可选稀释）；同一样品多项目共用
        # ICP 一 PDF 多样品(A/B 平行样)：samples = [(标识码, compounds), ...]，按平行槽分取
        from report_parser import (parse_pdf_report, parse_pdf_report_multi, _dilution_factor,
                                   filter_samples_by_code, _split_content_dilution, parse_pdf_report_meta)
        parsed_by_sample = {}   # {sample_code: (samples, diluted_compounds, headers)}
        factor_by_sample = {}   # {sample_code: 稀释倍数}
        missing_samples = []
        for it in batch_items:
            sc = it.get("sample_code")
            if sc in parsed_by_sample or sc in missing_samples:
                continue
            # ponytail: 方法未启用时，仅解析称量记录中标记了"解析"的样品
            if not method_enabled and not (wmap.get(sc) or {}).get("force_parse"):
                continue
            normal_pdfs, dil_pdfs = _match_sample_report_pdfs(spectrum_path, sc, actual_method_name)
            if not normal_pdfs:
                log(f"报告解析: 样品 {sc} 在谱图目录未找到报告 PDF，浓度留空")
                missing_samples.append(sc)
                continue
            dil_pdf = dil_pdfs[0] if dil_pdfs else None
            # 平行样分离 PDF(GCMS: A/B 各一份)：按文件名序(A→B)逐份解析，合并样品段按平行槽回填
            samples, headers = [], None
            for _pp in normal_pdfs:
                try:
                    _s, _h = parse_pdf_report_multi(_pp)
                except Exception as e:
                    log(f"报告解析: 样品 {sc} 报告 {os.path.basename(_pp)} 解析失败({e})，跳过该份")
                    continue
                if _s:
                    samples.extend(_s)
                    headers = headers or _h
            if not samples:
                log(f"报告解析: 样品 {sc} 未解析到化合物，浓度留空")
                missing_samples.append(sc)
                continue
            samples = filter_samples_by_code(samples, sc)  # ICP 多报验批共一份 PDF：只取当前样品段
            # 内容稀释: 段 id 形如 TS...001-10X -> 视作该样品的稀释源(镜像文件名 -NNX 稀释 PDF 流)
            normal_samples = []
            diluted = None
            content_factor = None
            for sid, cmp in samples:
                _base, f = _split_content_dilution(sid)
                if f != 1.0:
                    diluted = cmp              # 段即稀释源(原值不乘;LIMS 据稀释列自算)
                    content_factor = f
                else:
                    normal_samples.append((sid, cmp))
            samples = normal_samples or samples   # 全为稀释段时退回原样,保留旧行为
            # 文件名稀释 PDF 仍优先(两者互斥:单样品PDF文件名 vs 多样品PDF内容)
            if dil_pdf:
                try:
                    _dil_s, _ = parse_pdf_report_multi(dil_pdf)
                    _dil_s = filter_samples_by_code(_dil_s, sc)
                    # ponytail: 稀释源按平行槽对齐(A-10X→槽0, B-10X→槽1)，不再只取首样
                    diluted = [cmp for _sid, cmp in _dil_s] or None
                except Exception as e:
                    log(f"报告解析: 样品 {sc} 稀释报告解析失败({e})，按正常报告处理")
                # ponytail: 倍数优先取报告内「稀释：」字段(权威)，回退文件名 -NNX
                _fv = next((v for _s, v in parse_pdf_report_meta(dil_pdf, '稀释') if v), None)
                factor_by_sample[sc] = float(_fv) if _fv else _dilution_factor(dil_pdf)
            elif content_factor is not None:
                factor_by_sample[sc] = content_factor
            parsed_by_sample[sc] = (samples, diluted, headers)
            n_cmp = len(samples[0][1]) if samples else 0
            _names = "、".join(os.path.basename(p) for p in normal_pdfs)
            log(f"报告解析: 样品 {sc} 解析到 {len(samples)} 个样品×{n_cmp} 化合物 ({_names})"
                + (f"，含稀释报告 {os.path.basename(dil_pdf)}×{factor_by_sample[sc]:g}" if dil_pdf else ""))
        if not parsed_by_sample:
            return  # 无一样品可解析：结果列保持默认

        # 2) 定位结果列(用任一已解析样品的 PDF 表头做 equipRelativeTitle 精确匹配)
        any_headers = next(iter(parsed_by_sample.values()))[2]
        res_col = _find_result_column(dynamic_columns, any_headers)
        if not res_col:
            log("报告解析: 未找到结果列(计算值/报告值/浓度)，浓度留默认")
            return

        # 3) 逐记录求值：记录→projectId→样品→报告→别名→evaluate_alias→选段值
        from alias_evaluator import evaluate_alias
        records = (experiment_config or {}).get("ocAnalysisRecordList") or []
        pid_to_item = {str(it["project"].get("projectId")): it for it in batch_items
                       if it.get("project", {}).get("projectId") is not None}

        def _is_comp(c):
            try:
                return int(c.get('isMutiPolyColume') or 0) == 1
            except (TypeError, ValueError):
                return False

        comp_col_code = next((c.get("columeCode") for c in (dynamic_columns or [])
                              if isinstance(c, dict) and _is_comp(c)), None)
        alias_cache = {}   # {detectionProjectId: alias_str}

        parallel_of, _n_par = _parallel_indices(records)  # 每条记录的平行槽(用 serialNumber)

        # 告警：报告样品段数 > 该样品平行槽(=称样量数) → 多余段(如 B)被静默丢弃
        # 平行数取自称样量个数(见 _submit_batch _par_by)，称样量不足时扩不出对应平行槽，
        # _value_for_record 槽超界回落 samples[0] → 报告里 A/B 的 B 被丢。这里让它可见。
        _slots_by_sc = {}
        for _g, r in enumerate(records):
            _item = pid_to_item.get(str(r.get("projectId")))
            if _item:
                _sc = _item.get("sample_code")
                _slots_by_sc[_sc] = max(_slots_by_sc.get(_sc, 0), parallel_of.get(_g, 0) + 1)
        for _sc, (_smp, _dil, _h) in parsed_by_sample.items():
            _slots = _slots_by_sc.get(_sc, 1)
            if len(_smp) > _slots:
                log(f"报告解析警告: 样品 {_sc} 报告含 {len(_smp)} 个样品段，但称样量仅扩出 {_slots} 个平行槽——"
                    f"多余 {len(_smp) - _slots} 个(如 B)将被丢弃，请在称样记录补全平行称样量")

        def _value_for_record(rec, slot):
            item = pid_to_item.get(str(rec.get("projectId")))
            if not item:
                return "", False
            parsed = parsed_by_sample.get(item.get("sample_code"))
            if not parsed:
                return "", False  # 该样品 PDF 缺失/解析失败 → 留空
            samples, diluted_compounds, _headers = parsed
            # 取该平行槽样品；槽超界(平行数<样品数，如 N=1 而报告有 A/B)则取首个
            compounds = samples[slot][1] if slot < len(samples) else samples[0][1]
            # 稀释源按平行槽取(dil_pdf 多样品 A-10X/B-10X 各对其槽)；单 dict(内容稀释)或 None 原样
            if isinstance(diluted_compounds, list):
                diluted_compounds = (diluted_compounds[slot] if slot < len(diluted_compounds)
                                     else diluted_compounds[0]) if diluted_compounds else None
            det_pid = item["project"].get("detectionProjectId")
            if not det_pid:
                return "", False
            if det_pid not in alias_cache:
                alias, _detail = self.api.get_project_alias(det_pid, log)
                alias_cache[det_pid] = alias or ""
            alias = alias_cache.get(det_pid)
            if not alias:
                return "", False
            try:
                results = evaluate_alias(alias, compounds, diluted_compounds=diluted_compounds)
            except Exception:
                return "", False
            if not results:
                return "", False
            if comp_col_code:  # 多组分：按记录组分名匹配段(组分名取自记录该列值)
                comp_name = str(rec.get(comp_col_code) or "").strip()
                if comp_name:
                    seg = next((r for r in results if r.get("lims_component") == comp_name), None)
                    if seg is not None:
                        return seg.get("value", ""), bool(seg.get('raw', {}).get('diluted'))
            seg0 = results[0]
            return seg0.get("value", ""), bool(seg0.get('raw', {}).get('diluted'))  # 无组分列(PAHs)：单段即该化合物浓度

        rec_values = [_value_for_record(r, parallel_of.get(g, 0)) for g, r in enumerate(records)]
        values = [v for v, _d in rec_values]
        diluted_flags = [_d for _v, _d in rec_values]
        host.data_fields[res_col] = [_Box(v) for v in values]
        # 稀释列：超线性(稀释)记录填倍数，其余填 1（LIMS 据稀释列×结果自算）
        dil_col = next((c.get("columeCode") for c in (dynamic_columns or [])
                        if isinstance(c, dict) and (c.get('equipRelativeTitle') or '').strip() == '稀释'), None)
        if dil_col:
            def _dil_val(i, rec):
                if not diluted_flags[i]:
                    return "1"
                _item = pid_to_item.get(str(rec.get("projectId")))
                _f = factor_by_sample.get(_item.get("sample_code")) if _item else None
                return f"{_f:g}" if _f and _f != 1.0 else "1"
            host.data_fields[dil_col] = [_Box(_dil_val(i, rec)) for i, rec in enumerate(records)]
            _dil_filled = sum(1 for i in range(len(records)) if i < len(diluted_flags) and diluted_flags[i])
            if _dil_filled:
                log(f"报告解析: 稀释列 {dil_col} 填 {_dil_filled} 条超线性组分")
        filled = sum(1 for v in values if v)
        log(f"报告解析: 结果列 {res_col} 浓度回填 {filled}/{len(values)} 条"
            + (f"，{len(missing_samples)} 个样品缺报告PDF" if missing_samples else ""))

    def _fill_dilution_from_filename(self, host, dynamic_columns, experiment_config,
                                     batch_items, row, actual_method_name, ctx, log):
        """启用稀释备注 + 未启用报告解析时：扫各样品谱图目录的稀释 PDF(文件名-NNX)取额外倍数，
        填稀释列(整样品稀释→该样品全部记录)。报告解析启用或含强制解析样品时不运行(报告解析已填、更精确)。"""
        settings = self._read_report_parse_settings(ctx.get("method_file"))
        if settings and settings.get("enabled"):
            return  # 报告解析已填稀释列(文件名+内容两路)，不重复
        wmap = ctx.get("wmap") or {}
        if any((wmap.get(it.get("sample_code") or "") or {}).get("force_parse") for it in batch_items):
            return  # 含强制解析样品：报告解析已处理，避免整样品覆盖其逐组分结果
        spectrum_path = (row.get("spectrum_path") or "").strip()
        if not spectrum_path:
            return
        dil_col = next((c.get("columeCode") for c in (dynamic_columns or [])
                        if isinstance(c, dict) and (c.get('equipRelativeTitle') or '').strip() == '稀释'), None)
        if not dil_col:
            return
        from report_parser import _dilution_factor
        factor_by_sample = {}
        for it in batch_items:
            sc = it.get("sample_code")
            if not sc or sc in factor_by_sample:
                continue
            _pdf, dil_pdf = _pick_sample_report_pdf(spectrum_path, sc, actual_method_name)
            if dil_pdf:
                f = _dilution_factor(dil_pdf)
                if f and f != 1.0:
                    factor_by_sample[sc] = f
        if not factor_by_sample:
            return
        records = (experiment_config or {}).get("ocAnalysisRecordList") or []
        pid_to_item = {str(it["project"].get("projectId")): it for it in batch_items
                       if it.get("project", {}).get("projectId") is not None}

        def _dil_val(rec):
            _item = pid_to_item.get(str(rec.get("projectId")))
            _f = factor_by_sample.get(_item.get("sample_code")) if _item else None
            return f"{_f:g}" if _f and _f != 1.0 else "1"

        host.data_fields[dil_col] = [_Box(_dil_val(rec)) for rec in records]
        log(f"稀释备注(文件名): 稀释列 {dil_col} 按 -NNX 填 {len(factor_by_sample)} 个稀释样品 "
            f"({', '.join(f'{sc}×{f:g}' for sc, f in factor_by_sample.items())})")

    def _filter_projects_by_method(self, projects, row, log):
        """按方法文件 query_rules 指定的方法过滤 projects。
        返回 (projects, error_msg)。error_msg 非空表示无法确定单一方法——
        不回退到全部：多方法会让 getOcExperiment 报"样品项目对应的方法不同"。"""
        # 排除项目：先剔除命中 exclusion_rules 的项目(project + method 双维度)，再走方法/项目过滤
        excl = self._read_exclusion_rules(row.get("method_file") or "")
        if excl:
            kept = [p for p in projects if not any(_exclusion_match(p, r) for r in excl)]
            if len(kept) != len(projects):
                log(f"排除项目：{len(projects) - len(kept)} 个项目被排除规则剔除，剩余 {len(kept)} 个")
            projects = kept
            if not projects:
                return [], None
        # 样品实际含哪些方法(标准号)
        stdnos = []
        for p in projects:
            s = (p.get("standardNo") or "").strip()
            if s and s not in stdnos:
                stdnos.append(s)

        method_file = row.get("method_file") or ""
        # switch_rules 的 from_id(基方法)：规则方法解析出此ID时，项目实际带的是 switch 目标子方法(to_id)，
        # 不应按基方法ID剔除——子方法路由交由 switch_rules(project_name→to_id)负责
        _switch_from_ids = {str(r.get("from_id") or "").strip()
                            for r in (self._read_switch_rules(method_file) or [])
                            if str(r.get("from_id") or "").strip()}
        qr = self._read_query_rules(method_file)
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

        # 解析每条规则的方法 → (标准号, 子方法ID)。文本标准号名(如 'AfPS GS 2019:01 PAK 单组份')解析子方法ID。
        # 一个标准号下常有多个子方法(单组份/N项之和)，各为独立 methodId，须按子方法ID精确过滤，
        # 否则同标准号的项目混在一起会让 getOcExperiment 报 005"样品项目对应的方法不同"。
        # ponytail: 多规则首次解析并行(串行=N×RTT，afps 6规则≈6倍延迟)，缓存命中后样品2+零开销
        _need = [m for m in dict.fromkeys(str(q.get("method") or "").strip() for q in qr
                                          if str(q.get("method") or "").strip())
                 if m not in self.api._std_no_name_to_id]
        if _need:
            self.api._prefetch_std_no_name_ids(_need, log)  # 批量1次API(原逐规则串行N次≈N×3.5s)
        rule_stds = []  # [(规则序号, 标准号, 子方法ID或None)]，仅含有 method 的规则
        for _i, q in enumerate(qr):
            mv = str(q.get("method") or "").strip()
            if not mv:
                continue
            ts = mv
            _sub_id = self.api.get_method_id_by_standard_no_name(mv, log)
            if ts:
                rule_stds.append((_i, ts, _sub_id))

        if rule_stds:
            filtered = []
            _retest_skipped = []  # [(项目名, 记录是否复测, 方法是否复测)] 被 cancel_test 维度排除的项目
            for p in projects:
                _pname = (p.get("projectName") or "").strip()
                _pstd = p.get("standardNo", "")
                # 同时按 method+project 命中首条规则：方法匹配且(规则无 project 或项目名命中)。
                # 这样多方法各自独立分批(XRF 汞/六价铬/镉铅 各一实验)，单方法多项目(组分/总和)也各归其规则
                _matched = None
                for _i, ts, _sub_id in rule_stds:
                    if not _std_loose_match(_pstd, ts):
                        continue
                    # 规则解析出子方法ID时，项目 decideProjectMethodId 必须一致(剔除同标准号下别的子方法)；
                    # 但该ID若恰是 switch_rules.from_id(基方法)，项目带的是 switch 目标子方法(to_id)，跳过此过滤
                    if _sub_id and str(_sub_id) not in _switch_from_ids \
                            and str(p.get("decideProjectMethodId") or "") != str(_sub_id):
                        continue
                    _rp = str(qr[_i].get("project") or "").strip()
                    if not _rp or _project_match(_pname, _rp):
                        _matched = _i
                        break
                if _matched is None:
                    continue
                # 复测维度(cancel_test)：True 仅留复测(oldSampleProjectId 有值)，False 仅留非复测
                _is_retest = p.get("oldSampleProjectId") is not None
                if bool(qr[_matched].get("cancel_test")) != _is_retest:
                    _retest_skipped.append((_pname, _is_retest, bool(qr[_matched].get("cancel_test"))))
                    continue
                p["_qr_idx"] = _matched
                filtered.append(p)
            if filtered:
                distinct = sorted({ts for _i, ts, _ in rule_stds})
                _sub_ids = sorted({str(_s) for _i, _ts, _s in rule_stds if _s})
                proj_hint = f"，项目名过滤={project_vals!r}" if project_vals else ""
                sub_hint = f"，子方法ID={_sub_ids}" if _sub_ids else ""
                _pstds = sorted({str(p.get("standardNo") or "") for p in filtered})
                _pnames = [str(p.get("projectName") or "") for p in filtered]
                _skipped_hint = f"，复测维度(cancel_test)排除 {len(_retest_skipped)} 个" if _retest_skipped else ""
                log(f"按方法过滤出 {len(filtered)} 个项目(共 {len(rule_stds)} 条方法规则：{', '.join(distinct)}{proj_hint}{sub_hint})；"
                    f"标准号集合: {_pstds}；项目名: {_pnames}{_skipped_hint}")
                return filtered, None
            if _retest_skipped:
                _detail = "；".join(f"{n}(记录为{'复测' if a else '非复测'},方法设为{'复测' if e else '非复测'})"
                                    for n, a, e in _retest_skipped)
                return [], (f"{len(_retest_skipped)} 个项目因复测维度(cancel_test)不符被跳过：{_detail}。"
                            "请在方法编辑器调整该规则的 Retest 设置，或检查样品复测记录。")
            # 标准号命中却仍被剔除：细分是 子方法ID 还是 项目名 不匹配，避免误报"方法不在样品中"
            _std_hit = [p for p in projects
                        if any(_std_loose_match((p.get("standardNo") or ""), ts) for _i, ts, _ in rule_stds)]
            _sub_ids = sorted({str(_s) for _i, _ts, _s in rule_stds if _s})
            _detail = ""
            if _std_hit:
                _dpids = sorted({str(p.get("decideProjectMethodId") or "") for p in _std_hit})
                _pnames = [str(p.get("projectName") or "") for p in _std_hit]
                _sub_part = f"子方法ID不符(规则={_sub_ids}，项目decideProjectMethodId={_dpids})" if _sub_ids else ""
                _pn_part = f"项目名不符(规则={project_vals}，项目名={_pnames})"
                _detail = f" 其中 {len(_std_hit)} 个项目标准号已命中，但 {_sub_part + '；' if _sub_part else ''}{_pn_part}。"
            return None, (f"方法文件指定的方法不在此样品项目中。"
                          f"样品实际方法: {', '.join(stdnos)}。{_detail}请检查录入方法文件。")

        # 方法文件未指定方法：仅当样品只含单一方法时才可用全部(再按 project 过滤)
        if len(stdnos) <= 1:
            return [p for p in projects if _project_hit(p)], None
        return None, (f"样品含 {len(stdnos)} 个不同方法，但方法文件未在 query_rules 指定要提交的方法。"
                      f"方法列表: {', '.join(stdnos)}")

    def _read_switch_method_id(self, method_file):
        """读方法文件 query_rules[0].method_id（录入前要切换到的目标方法ID）。返回 str；空则 ''。
        用于 PD-苯 等需从默认方法切换到子方法才含计算公式的场景(对照主窗口 switch_method_id)。"""
        if not method_file:
            return ""
        try:
            y = load_method(method_file)
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

    def _read_switch_rules(self, method_file):
        """读方法文件 switch_rules（统一规则列表）。
        优先读新格式 switch_rules；回退合并旧格式 default_rules + filename_rules。"""
        if not method_file:
            return []
        try:
            y = load_method(method_file)
            switch_rules = y.get("switch_rules")
            if switch_rules is not None:
                return switch_rules
            # 向后兼容旧格式
            default_rules = y.get("default_rules") or []
            filename_rules = y.get("filename_rules") or []
            merged = []
            for r in default_rules:
                merged.append({
                    "from_id": r.get("from_id", ""),
                    "to_id": r.get("to_id", ""),
                    "project_name": r.get("project_name", ""),
                    "filename": "",
                    "desc": ""
                })
            for r in filename_rules:
                merged.append({
                    "from_id": r.get("from_id", ""),
                    "to_id": r.get("to_id", ""),
                    "project_name": r.get("project_name", ""),
                    "filename": r.get("filename", r.get("name", "")),
                    "desc": r.get("desc", "")
                })
            return merged
        except Exception:
            return []

    def _read_spectrum_check_params(self, method_file):
        """读方法文件 spectrum_upload_settings.spectrum_check_params。
        返回 {blank/standard/linearity/sample: {enabled,count,keyword}}；空/异常返回 {}。
        count 解析为 int(非数字/空=0)；keyword 去空白。"""
        if not method_file:
            return {}
        try:
            y = load_method(method_file)
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
        if not method_file:
            return False
        try:
            y = load_method(method_file)
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
        # 关键字为空时按样品编号前缀匹配(如 TN26070485 → TN26070485001.pdf)，与上传匹配逻辑一致。
        # 关键字含 * 时视为通配符：* 替换为样品编号前缀做精确 startswith 匹配（如 *,*T,*TS）。
        ps = params.get("sample") or {}
        if ps.get("enabled"):
            suf_str = ps.get("keyword")
            # 逗号/分号(中英文)都当分隔符(如 "A,B" / "*,*T；*TS")
            suffixes = _split_device_codes(suf_str) or [""]
            need = ps.get("count") or 1
            # 后缀表：关键字去 * 后(如 *,*T,*TS → ['','t','ts'])；非空后缀长者优先扣(避免 t 抢 ts)
            reqs = [s.replace("*", "").lower() for s in suffixes]
            known = sorted({s for s in reqs if s}, key=len, reverse=True)
            from report_parser import _DILUTION_RE  # 剥 -NNX 稀释后缀，避免稀释报告自成平行组被强求各后缀齐全
            for sc, _pdfs in samples:
                base = _strip_parallel_suffix(sc).lower()  # 报验号(去平行小号 001)
                head = re.compile(re.escape(base) + r"(\d{3})?(.*)$")
                # 拆每个文件为 (平行字母, 后缀)：先扣已知后缀，剩余单字母 = 平行字母
                # 文件名约定 base + 可选小号 + 可选平行字母(A/B) + 后缀(T/TS)；简写无小号
                groups = {}
                for n in names:
                    # 稀释报告 -NNX 先剥：并入基样(算作其谱图一份)，不另成平行组(否则被要求 T/TS 齐全)
                    stem = _DILUTION_RE.sub('', os.path.splitext(n)[0].lower())
                    m = head.match(stem)
                    if not m:
                        continue
                    tail = m.group(2)
                    suf = ""
                    for s in known:
                        if tail.endswith(s):
                            suf, tail = s, tail[:-len(s)]
                            break
                    groups[(tail, suf)] = groups.get((tail, suf), 0) + 1
                # 平行样：出现的每个平行字母，对方法各后缀各需 ≥need 份
                for raw, req in zip(suffixes, reqs):
                    tag = f"{raw}谱图" if raw else "谱图"
                    for p in sorted({par for (par, _s) in groups}):
                        actual = groups.get((p, req), 0)
                        if actual < need:
                            who = f"{sc}{p.upper()}" if p else sc
                            missing.append(f"样品{who}的{tag}需≥{need}个，实际{actual}个")
        return (not missing), missing

    def _build_headless_host(self, values, experiment_config, equipment_config,
                             dynamic_columns, actual_method_name, actual_method_id, configure_order,
                             fixed_params=None, dilution_remark_enabled=False, dilution_volume_column=""):
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
            is_headless=True,  # 序列无头模式：固定参数覆盖列默认值(无真实手填输入)
            temperature_var=_Box(""),
            humidity_var=_Box(""),
            start_date_var=_Box(today),
            end_date_var=_Box(today),
            dilution_remark_enabled=dilution_remark_enabled,
            dilution_volume_column=dilution_volume_column,
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
        ok = sum(1 for r in self.sequence_data if r.get("status") == "成功")
        fail = sum(1 for r in self.sequence_data if r.get("status") == "失败")
        abort = sum(1 for r in self.sequence_data if r.get("status") == "中止")
        total = ok + fail + abort
        self._append_log("==== 序列运行结束 ====")
        self._append_log(f"运行报告：共 {total} 行 — 成功 {ok} / 失败 {fail} / 中止 {abort}")
        # 汇总：各成功行样品合并情况(跨行合计)，格式对齐单行「成功，… 个样品参与合并」
        rows_ok = [r for r in self.sequence_data if r.get("status") == "成功"]
        tot = sum(r.get("samples_total", 0) for r in rows_ok)
        merged = sum(r.get("samples_merged", 0) for r in rows_ok)
        if tot:
            codes_all = [r.get("experiment_code") for r in rows_ok if r.get("experiment_code")]
            skipped = [(sc, reason) for r in rows_ok for (sc, reason) in r.get("skipped_list", [])]
            extra = f"，跳过 {len(skipped)} 个样品" if skipped else ""
            self._append_log(f"汇总：{merged}/{tot} 个样品参与合并，实验编号 {' / '.join(codes_all)}{extra}")
            if skipped:
                self._append_log(f"未录入样品（{len(skipped)} 个）：")
                for sc, reason in skipped:
                    self._append_log(f"  - {sc}：{reason}")
        early = [t for r in rows_ok for t in r.get("early_analysis", [])]
        if early:
            self._append_log(f"时间调整：{len(early)} 个样品分析时间早于受理时间，startTime 已提到受理日（服务端005校验）：")
            for sc, orig, new, acc in early:
                self._append_log(f"  - {sc}：{orig} → {new}（受理日 {acc}）")
        if fail:
            lines = []
            for i, r in enumerate(self.sequence_data):
                if r.get("status") == "失败":
                    _code = r.get('sample_code')
                    _sp = os.path.basename((r.get('spectrum_path') or '').rstrip('/\\')) or ''
                    # 标识：优先样品编号，无则用谱图文件夹名(多样品行从称样记录展开，前端无编号输入)
                    _tag = f"[{_code}]" if _code else (f"[{_sp}]" if _sp else "")
                    _head = f"{_tag} " if _tag else ""
                    lines.append(f"  · 第{i + 1}行 {_head}{r.get('error_msg') or ''}")
                    for sc, reason in r.get("skipped_list", []):
                        lines.append(f"      - {sc}：{reason}")
            self._append_log("失败明细：\n" + "\n".join(lines))
        self.status_var.set(f"完成: 成功 {ok} / 失败 {fail}")

    def _append_log(self, msg):
        self.log_text.configure(state='normal')
        self.log_text.insert('end', msg + "\n")
        self.log_text.see('end')
        self.log_text.configure(state='disabled')

    def clear_log(self):
        """清空运行日志"""
        self.log_text.configure(state='normal')
        self.log_text.delete('1.0', 'end')
        self.log_text.configure(state='disabled')

    def _export_log(self):
        """导出运行日志为 txt 文件"""
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
            initialfile=f"运行日志_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        )
        if path:
            with open(path, 'w', encoding='utf-8') as f:
                f.write(self.log_text.get("1.0", "end-1c"))

    def _log(self, msg):
        self._ui_q.put(("log", f"[{time.strftime('%H:%M:%S')}] {msg}"))

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
            self.tb_pause.configure(text="▶")
            self._log("已暂停(行边界生效)")
            self.status_var.set("已暂停")
        else:
            self._pause.set()
            self.pause_btn.configure(text="⏸ 暂停")
            self.tb_pause.configure(text="⏸")
            self._log("已继续")
            self.status_var.set("运行中...")

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
        """根据登录状态刷新用户菜单首项(登录/切换用户)与登录状态显示"""
        if self.logged_in and self.login_system.current_user:
            disp = self.login_system.users.get(self.login_system.current_user, {}).get(
                'display_name', self.login_system.current_user)
            self.user_menu.entryconfigure(0, label="切换用户")
            self.login_status_label.configure(text=f"👤 {disp}", fg="#16a34a")
        else:
            self.user_menu.entryconfigure(0, label="登录")
            self.login_status_label.configure(text="👤 未登录", fg="#dc2626")

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
    p = _plan_submission_batches(rules_m, items)
    assert [(b["switch_mid"], len(b["items"]), b["force_new"]) for b in p] \
        == [("4678", 2, False), ("4679", 2, False)], p

    # 样品模式：每样品各1批，苯先于总和
    rules_s = [{"project": "苯", "input_method": "样品"},
               {"project": "甲苯、二甲苯及乙苯总和", "input_method": "样品"}]
    p2 = _plan_submission_batches(rules_s, items)
    assert [(b["switch_mid"], b["items"][0]["sample_code"]) for b in p2] \
        == [("4678", "S1"), ("4678", "S2"), ("4679", "S1"), ("4679", "S2")], p2

    # 方法模式 max_select=1(按行)：苯规则限 1 → 拆2批，两批 force_new=True；总和规则不限 → 合1批
    rules_max = [{"project": "苯", "input_method": "方法", "max_select": "1"},
                 {"project": "甲苯、二甲苯及乙苯总和", "input_method": "方法"}]
    p3 = _plan_submission_batches(rules_max, items)
    assert [b["force_new"] for b in p3 if b["switch_mid"] == "4678"] == [True, True], p3
    assert [b["force_new"] for b in p3 if b["switch_mid"] == "4679"] == [False], p3

    # 单批录入条数硬上限(max_items)：1样品5项目→原 force_new=False，切片[2,2,1]各出新编号
    _one = [{"sample_code": "S1", "projectName": f"P{i}", "switch_mid": "4678",
             "project": {"projectId": f"S1-P{i}", "projectName": f"P{i}"}} for i in range(5)]
    p4 = _plan_submission_batches([{"project": "", "input_method": "方法"}], _one, max_items=2)
    assert [len(b["items"]) for b in p4] == [2, 2, 1], [len(b["items"]) for b in p4]
    assert all(b["force_new"] for b in p4), p4
    assert sum(len(b["items"]) for b in p4) == 5, p4
    _p4b = _plan_submission_batches([{"project": "", "input_method": "方法"}], _one, max_items=10)
    assert len(_p4b) == 1 and not _p4b[0]["force_new"] and len(_p4b[0]["items"]) == 5, _p4b

    # 称样日期拆批(方法模式)：同 switch_mid 跨日 → 拆 2 批 force_new=True；同日仍合 1 批
    di = [item("S1", "苯", "4678"), item("S2", "苯", "4678")]
    di[0]["wdate"] = "2026-07-28"
    di[1]["wdate"] = "2026-07-29"
    pd = _plan_submission_batches([{"project": "苯", "input_method": "方法"}], di)
    assert len(pd) == 2 and all(b["force_new"] for b in pd), pd
    assert [b["items"][0]["sample_code"] for b in pd] == ["S1", "S2"], pd
    di[1]["wdate"] = "2026-07-28"  # 改同日 → 合 1 批
    pd2 = _plan_submission_batches([{"project": "苯", "input_method": "方法"}], di)
    assert len(pd2) == 1 and not pd2[0]["force_new"], pd2

    # _parallel_indices：同 projectId 多组分(总和 5组分×2平行) 按 serialNumber 正确归平行
    recs = [{"projectId": "P1", "serialNumber": s} for s in (1, 2, 1, 2, 1, 2, 1, 2, 1, 2)]
    po, n = _parallel_indices(recs)
    assert n == 2 and po == {0: 0, 1: 1, 2: 0, 3: 1, 4: 0, 5: 1, 6: 0, 7: 1, 8: 0, 9: 1}, (n, po)
    # 无 serialNumber 回退：按 projectId 枚举
    po2, n2 = _parallel_indices([{"projectId": "P1"}, {"projectId": "P1"}])
    assert n2 == 2 and po2 == {0: 0, 1: 1}, (n2, po2)

    # _pdf_parallel_count：random/无称样记录模式，平行数由文件名 A/B 推断(扣已知后缀 T/TS)
    _suf = {"", "t", "ts"}
    assert _pdf_parallel_count("TN26070724001",
        ["TN26070724001A.pdf", "TN26070724001AT.pdf", "TN26070724001ATS.pdf",
         "TN26070724001B.pdf", "TN26070724001BT.pdf", "TN26070724001BTS.pdf"], _suf) == 2
    # 单样(仅 T/TS 后缀、无 A/B) → 1
    assert _pdf_parallel_count("TN26070729001",
        ["TN26070729001.pdf", "TN26070729001T.pdf", "TN26070729001TS.pdf"], _suf) == 1

    # _match_switch_rule：统一规则 project_name + filename + desc 匹配(desc 默认空=不限)
    # project_name 使用 wildcard 匹配(与 _project_match 同口径)
    fr = [{"to_id": "4481", "filename": "K", "project_name": "*DEHP*"},
          {"to_id": "4482", "filename": "K", "project_name": "*DNOP*"}]
    # 命中：DNOP + 文件名含 K -> 4482（大小写不敏感）
    assert _match_switch_rule(fr, "DNOP", ["D:/sp/K-001.pdf"]) == "4482"
    # 不命中：DBP 无名称规则 -> ''
    assert _match_switch_rule(fr, "DBP", ["D:/sp/K-001.pdf"]) == ""
    # 不命中：DNOP 但文件名不含 K -> ''
    assert _match_switch_rule(fr, "DNOP", ["D:/sp/001.pdf"]) == ""
    # project_name wildcard 包含于 LIMS 长项目名 -> 路由到各自 to_id
    assert _match_switch_rule(fr, "3种邻苯二甲酸酯类化合物（DBP、BBP、DEHP）总和", ["K-001.pdf"]) == "4481"
    assert _match_switch_rule(fr, "3种邻苯二甲酸酯类化合物（DNOP、DINP、DIDP）总和", ["K-001.pdf"]) == "4482"
    # 精确匹配：无通配符时需完全相等
    fr_exact = [{"to_id": "9901", "filename": "X", "project_name": "DNOP"}]
    assert _match_switch_rule(fr_exact, "DNOP", ["X-001.pdf"]) == "9901"
    assert _match_switch_rule(fr_exact, "3种邻苯二甲酸酯类化合物（DNOP、DINP、DIDP）总和", ["X-001.pdf"]) == ""

    # 统一规则 + 试样描述(desc)：文件名与描述均非空时需同时命中(AND)；任一为空=不限
    mr = [{"to_id": "5501", "filename": "K", "desc": "苯,甲苯"}]
    # 文件名命中但描述不含任一关键字 -> 不命中(两条件都填需同时满足)
    assert _match_switch_rule(mr, "", ["K-001.pdf"], "水溶液") == ""
    # 文件名 + 描述都命中 -> 5501
    assert _match_switch_rule(mr, "", ["K-001.pdf"], "苯溶液") == "5501"
    # 描述第二关键字"甲苯"命中 -> 5501
    assert _match_switch_rule(mr, "", ["K-001.pdf"], "甲苯") == "5501"
    # 只填描述(无文件名条件) -> 命中
    assert _match_switch_rule([{"to_id": "5502", "desc": "固体"}], "", [], "固体颗粒") == "5502"
    # 规则要求描述关键字但样品描述为空 -> 不命中
    assert _match_switch_rule([{"to_id": "5502", "desc": "固体"}], "", [], "") == ""
    # 大小写不敏感
    assert _match_switch_rule([{"to_id": "9", "desc": "Solid"}], "", [], "SOLID sample") == "9"
    # desc 逗号=OR(任一)，分号=AND(均需命中)；中英文标点等价
    assert _match_switch_rule([{"to_id": "1", "desc": "纸,木头"}], "", [], "木头样品") == "1"   # OR 命中其一
    assert _match_switch_rule([{"to_id": "1", "desc": "纸,木头"}], "", [], "塑料") == ""        # OR 都不中
    assert _match_switch_rule([{"to_id": "2", "desc": "纸;涂层"}], "", [], "纸基涂层") == "2"   # AND 都中
    assert _match_switch_rule([{"to_id": "2", "desc": "纸;涂层"}], "", [], "纸") == ""          # AND 缺一
    assert _match_switch_rule([{"to_id": "3", "desc": "纸，布；涂层"}], "", [], "布涂层") == "3" # (纸OR布)AND涂层，全角标点

    # _override_equipment 多设备：固定设备 + 替换可变设备；不依赖 self（用裸实例避免建 Tk）
    _eq_inst = object.__new__(SequenceMaster)
    _raw = [
        {"usedCategory": "检测设备", "name": "超声波清洗机", "mainEquipmentNames": "CK-SB063-CG,超声波清洗机,", "equipmentBillId": 1063, "isDefault": 1},
        {"usedCategory": "检测设备", "name": "气相色谱质谱联用仪", "mainEquipmentNames": "CK-SB017-CG,气相色谱质谱联用仪,2026-12-08", "equipmentBillId": 1017, "isDefault": 1},
        {"usedCategory": "检测设备", "name": "气相色谱质谱联用仪", "mainEquipmentNames": "CK-SB036-CG,气相色谱质谱联用仪,2026-12-08", "equipmentBillId": 1036, "isDefault": 0},
        {"usedCategory": "称样设备", "name": "分析天平", "mainEquipmentNames": "CK-WB001,分析天平,", "equipmentBillId": 9001, "isDefault": 1},
    ]
    _cfg = {"raw_data": _raw, "mainEquipmentIds": "1063,1017", "mainEquipmentNames": "", "mainEquipment": ""}
    # GC-MS 换 CK-SB036 + 保留清洗机 CK-SB063；旧 GC-MS(1017) 不在结果
    _o, _m, _e = _eq_inst._override_equipment(_cfg, "CK-SB036-CG;CK-SB063-CG")
    assert _e is None and len(_m) == 2, (_e, _m)
    assert set(_o["mainEquipmentIds"].split(",")) == {"1036", "1063"}, _o["mainEquipmentIds"]
    assert "CK-SB036-CG" in _o["mainEquipmentNames"] and "CK-SB063-CG" in _o["mainEquipmentNames"]
    # 单设备(向后兼容)
    _o1, _m1, _e1 = _eq_inst._override_equipment(_cfg, "CK-SB017-CG")
    assert _e1 is None and len(_m1) == 1 and _o1["mainEquipmentIds"] == "1017", (_e1, _m1)
    # 空值 -> 默认(matched_list=None)
    _o0, _m0, _e0 = _eq_inst._override_equipment(_cfg, "")
    assert _m0 is None and _e0 is None
    # 找不到 -> 报错指明编号
    _ox, _mx, _ex = _eq_inst._override_equipment(_cfg, "CK-NOPE;CK-SB063-CG")
    assert _ex and "CK-NOPE" in _ex, _ex
    # 设备无 mainEquipmentNames、仅编号字段(no)——编号须能匹配(修复"编号 名称"被整体当编号致提交校验失败)
    _raw2 = [{"usedCategory": "检测设备", "name": "气相色谱质谱联用仪", "no": "CK-SB005-EN", "equipmentBillId": 1005}]
    _cfg2 = {"raw_data": _raw2, "mainEquipmentIds": "1005", "mainEquipmentNames": "", "mainEquipment": ""}
    _o2, _m2, _e2 = _eq_inst._override_equipment(_cfg2, "CK-SB005-EN")
    assert _e2 is None and len(_m2) == 1 and _o2["mainEquipmentIds"] == "1005", (_e2, _m2)
    # 设备规则含称样设备(天平)→路由到称样设备(单台)，检测设备仍进主检
    _o3, _m3, _e3 = _eq_inst._override_equipment(_cfg, "CK-SB036-CG;CK-WB001")
    assert _e3 is None and len(_m3) == 1 and _o3["mainEquipmentIds"] == "1036", (_e3, _o3, _m3)
    assert _o3["weighingEquipmentId"] == "9001" and "CK-WB001" in _o3["weighingEquipment"], _o3
    # 称样量按 decimal_places 格式化，保留末尾0(0.552→0.5520)
    assert _apply_processing(0.552, {"type": "小数位补充"}, {"decimal_places": 4}) == "0.5520"
    assert f"{0.552:.4f}" == "0.5520"
    # 换算加补充：base=round(raw×factor,dp) 后末尾补2位随机[01,49]，修约回 dp 位仍=base，末两位非00
    _r = _apply_processing(0.52, {"type": "换算加补充", "factor": 4, "decimal_places": 2}, {})
    assert _r.startswith("2.08") and _r[-2:] != "00" and round(float(_r), 2) == 2.08, _r
    assert len(_r.split(".")[1]) == 4  # 总小数位 = dp(2) + 补2位
    for _ in range(200):  # pad 始终落在 01-49，修约回2位恒=2.08
        _rr = _apply_processing(0.52, {"type": "换算加补充", "factor": 4, "decimal_places": 2}, {})
        assert 1 <= int(_rr[-2:]) <= 49 and round(float(_rr), 2) == 2.08, _rr
    # 换算处理 不受影响(仍 raw×factor 修约)
    assert _apply_processing(0.52, {"type": "换算处理", "factor": 4, "decimal_places": 2}, {}) == "2.08"

    # _exclusion_match：排除规则双维度(project 精确/通配 + method 标准号宽松匹配)，非空条件 AND
    _p = lambda pn, std: {"projectName": pn, "standardNo": std}
    assert _exclusion_match(_p("苯", "GB 36246-2018 6.15.2"), {"project": "苯"}) is True
    assert _exclusion_match(_p("甲苯", "GB 36246-2018 6.15.2"), {"project": "苯"}) is False
    # project 通配
    assert _exclusion_match(_p("苯", ""), {"project": "*苯*"}) is True
    # method 维度：== 与 附录后缀容忍；大小写不敏感
    assert _exclusion_match(_p("苯", "GB 36246-2018 附录G"), {"project": "苯", "method": "gb 36246-2018"}) is True
    assert _exclusion_match(_p("苯", "GB/T 23991"), {"method": "GB 36246"}) is False
    # 两条件都填需同时命中
    assert _exclusion_match(_p("苯", "GB 36246-2018"), {"project": "甲苯", "method": "GB 36246-2018"}) is False

    # 多方法分批(XRF)：3 条不同 method 的规则，项目带 _qr_idx → 各方法各 1 批、不重复提交
    mm = [{"sample_code": "S1", "projectName": "镉", "switch_mid": "", "_qr_idx": 0, "project": {}},
          {"sample_code": "S1", "projectName": "铅", "switch_mid": "", "_qr_idx": 0, "project": {}},
          {"sample_code": "S1", "projectName": "汞", "switch_mid": "", "_qr_idx": 1, "project": {}},
          {"sample_code": "S1", "projectName": "六价铬", "switch_mid": "", "_qr_idx": 2, "project": {}}]
    mm_rules = [{"project": "", "method": "IEC 62321-5", "input_method": "方法"},
                {"project": "", "method": "IEC 62321-4", "input_method": "方法"},
                {"project": "", "method": "IEC 62321-7-2", "input_method": "方法"}]
    pm = _plan_submission_batches(mm_rules, mm)
    assert len(pm) == 3 and sum(len(b["items"]) for b in pm) == 4, pm  # 3 批、4 项目无重复
    assert sorted(len(b["items"]) for b in pm) == [1, 1, 2], pm        # 方法0=镉铅2个，其余各1
    assert all(b["force_new"] for b in pm), pm                         # 多方法各批独立编号

    # _std_loose_match：标准号空格不一致仍匹配(LIMS 汞 'IEC62321-4' vs 规则 'IEC 62321-4')，
    # 且各 XRF 元素方法不互配(汞≠镉铅≠六价铬)
    _hg = "IEC 62321-3-1:2013+IEC62321-4:2013+AMD1:2017"          # LIMS 实际值(第二段无空格)
    assert _std_loose_match(_hg, "IEC 62321-3-1:2013+IEC 62321-4:2013+AMD1:2017 IEC62321XRF") is True
    assert _std_loose_match(_hg, "IEC 62321-3-1:2013+IEC 62321-5:2013 IEC62321XRF") is False  # 汞≠镉铅
    assert _std_loose_match(_hg, "IEC 62321-3-1:2013+IEC 62321-7-2:2017 XRF") is False        # 汞≠六价铬
    assert _std_loose_match("IEC 62321-3-1:2013+IEC 62321-7-2:2017 XRF",
                           "IEC 62321-3-1:2013+IEC 62321-7-2:2017 XRF") is True

    # _enrich_weighing_bill：批内首个样品项目 ocChoicePage 没有该台账时，应遍历后续样品项目命中
    # (修复偶发「缺称样设备」：原实现只查 split(",")[0])
    _wid = "777"
    _raw_hit = {"id": 777, "creatorName": "系统管理员", "createDatetime": "2023-08-25 10:00:00"}
    class _FakeAPI:
        def __init__(self, table): self.t = table
        def get_weighing_equipment_choices(self, sp, log): return self.t.get(sp, [])
    _cfg = {"weighingEquipmentId": _wid, "weighingEquipmentRaw": None}
    _sm = SequenceMaster.__new__(SequenceMaster)
    _sm.api = _FakeAPI({"sp1": [{"raw": {"id": 111}}], "sp2": [{"raw": _raw_hit}]})
    _out = SequenceMaster._enrich_weighing_bill(_sm, _cfg, "sp1,sp2", lambda *a: None)
    assert _out["weighingEquipmentRaw"] is _raw_hit, _out          # sp1 未命中→遍历到 sp2 命中
    _sm2 = SequenceMaster.__new__(SequenceMaster)
    _sm2.api = _FakeAPI({"sp1": [{"raw": _raw_hit}]})
    _out2 = SequenceMaster._enrich_weighing_bill(_sm2, _cfg, "sp1,sp2", lambda *a: None)
    assert _out2["weighingEquipmentRaw"] is _raw_hit, _out2         # 回归：首个即命中

    # _code_belongs_sample：按小号匹配，不把同报验号其它小号吃进
    assert _code_belongs_sample("TS26072277087", "TS26072277087", "TS26072277")       # 精确
    assert not _code_belongs_sample("TS26072277055", "TS26072277087", "TS26072277")   # 同报验号不同小号不归属
    assert _code_belongs_sample("TN26070729001A", "TN26070729001", "TN26070729")     # 平行字母文件
    assert _code_belongs_sample("TS26072277", "TS26072277001", "TS26072277")         # 简写报验编号代表001
    assert not _code_belongs_sample("TS26072277", "TS26072277087", "TS26072277")     # 简写不代表087
    # 称样记录用报验编号(无小号=代表001)：谱图文件名带001(+T/TS字母)也归属，不依赖 PDF 内容
    assert _code_belongs_sample("TN26080320001", "TN26080320", "TN26080320")        # base 全码
    assert _code_belongs_sample("TN26080320001T", "TN26080320", "TN26080320")       # T 变体
    assert _code_belongs_sample("TN26080320001TS", "TN26080320", "TN26080320")      # TS 变体
    assert not _code_belongs_sample("TN26080320002T", "TN26080320", "TN26080320")   # 不同小号(002)不吃进

    # _ParallelWMap：单字母后缀称样记录(如 TN…K)按 LIMS 全码(…001)能查回 desc/称样量。
    # 回归：K 样品子方法切换(desc 匹配)与称样量回填不再因 key≠全码而 miss。
    _wm = _merge_parallel_groups(_ParallelWMap({
        "TN26080226K": {"desc": "草颗粒", "masses": [0.5140], "time": "2026-08-06 10:00:00"},
        "TN26080248001A": {"desc": "草", "masses": [0.5268], "time": "2026-08-06 10:00:00", "force_parse": True},
        "TN26080248001B": {"desc": "草", "masses": [0.5016], "time": "2026-08-06 10:00:00"},
    }))
    assert "TN26080226K" in _wm and "TN26080248001" in _wm, dict(_wm)   # K 保留原键；A/B 合并为全码
    assert _wm.get("TN26080226001").get("desc") == "草颗粒", "K 后缀按 LIMS 全码查应命中"
    assert _wm.get("TN26080248001").get("masses") == [0.5268, 0.5016]   # A/B 平行合并(原有行为)
    assert _wm.get("TN26080248001").get("force_parse") is True          # 「解析」标记随平行合并保留
    assert _wm.get("TN26080226K").get("desc") == "草颗粒"               # 原键仍可直接查
    # _accept_date_of：受理日期解析(含时分串/仅日期/空/多候选)
    assert _accept_date_of({"acceptTime": "2026-08-08 12:00:00"}) == (date(2026, 8, 8), "acceptTime")
    assert _accept_date_of({"acceptDate": "2026-08-09"}) == (date(2026, 8, 9), "acceptDate")
    assert _accept_date_of({})[0] is None

    # _match_processing_rule：空 method=通配(对该方法文件所有样品生效，与标准号无关)；
    # 非空 method 仍按标准号精确/包含双向匹配。回归 TDI/液体TDI 的 method="" 换算加补充规则。
    _pr_blank = {"method": "", "type": "换算加补充", "factor": 4}
    assert _match_processing_rule([_pr_blank], "GB/T 18446-2009") is _pr_blank   # 空通配命中
    assert _match_processing_rule([_pr_blank], "") is _pr_blank                  # 标准号为空也命中
    assert _match_processing_rule([{"method": "18446", "type": "换算处理"}], "GB/T 18446-2009") \
        == {"method": "18446", "type": "换算处理"}                              # 子串匹配仍有效
    assert _match_processing_rule([{"method": "GB/T 9999"}], "GB/T 18446-2009") is None  # 不中→None
    assert _match_processing_rule([], "GB/T 18446-2009") is None                          # 无规则→None

    # _random_masses：random 称样量按本样品平行数(n_par)决定用记录还是随机(不再用全局 max)。
    # 回归 PAE TN26080385：混批(他样2平行→旧代码全局 max=2)下本样1平行且录了0.2751 → 必须用记录。
    _wp_r = {"min_value": 1.9, "max_value": 2.2, "decimal_places": 4}
    assert _random_masses(["0.2751"], 1, None, _wp_r, 4) == ["0.2751"]            # 单值不被随机覆盖
    assert _random_masses(["0.5000", "0.6000"], 2, None, _wp_r, 4) == ["0.5000", "0.6000"]  # 2平行2值
    _r2 = _random_masses(["0.2751"], 2, None, _wp_r, 4)                          # 需2平行仅录1值→随机补
    assert len(_r2) == 2 and all(1.9 <= float(v) <= 2.2 for v in _r2)
    _r3 = _random_masses([], 1, None, _wp_r, 4)                                  # 无记录→随机
    assert len(_r3) == 1 and 1.9 <= float(_r3[0]) <= 2.2
    _pr_x4 = {"type": "换算处理", "factor": 4, "decimal_places": 4}              # prule 换算路径
    assert _random_masses(["0.2751"], 1, _pr_x4, _wp_r, 4) == ["1.1004"]
    print("selfcheck OK")


if __name__ == "__main__":
    if "--selfcheck" in sys.argv:
        _selfcheck()
        sys.exit(0)
    root = ttkb.Window(themename="sandstone-light")
    app = SequenceMaster(root)
    root.mainloop()