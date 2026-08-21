"""项目别名（解析规则）解析与求值。

「项目别名」是配置在 LIMS（检测标准管理→项目→属性→项目别名）的字符串，即谱图数据采集的
解析规则本身。运行时由 detection_entry_api.get_project_alias 取回 otherName 字段传入本模块。
语法见 希科仪器读取设置信息说明文档.doc：

  单组分: compound  |  compound;<blank>  |  compound;<blank>#内标  |  compound;<blank>/[lo-hi]
  多组分: 别名1=compound1;<blank>#内标/[lo-hi]|别名2=compound2;...

  - ;   未检出时报出的字面值（如 <0.005）；缺省则用默认 ND 字面值
  - #   内标物名（v1 仅记录，不改值——PDF 已给出目标物最终浓度）
  - /[lo-hi]  线性范围（v1 仅解析，越界告警不钳制，避免篡改数据）
  - =   多组分：左侧为 LIMS 组分/项目名，右侧为 PDF 中印的化合物名
  - |   多组分分隔
"""
import re
import unicodedata

# ponytail: 线性范围 /[lo-hi]，可选前置斜杠；ICP 用分号分隔(';[lo-hi]')，故前导符 [/;] 均吃掉避免尾分号
_RANGE_RE = re.compile(r'[;/]?\[([0-9]+(?:\.[0-9]+)?)-([0-9]+(?:\.[0-9]+)?)\]')

# 全角标点(U+FF01..FF5E)→半角；LIMS 中文前端常录成 ；＝｜／＃，规范化后才按分隔符切。
# ponytail: 一次性映射整个全角 ASCII 区，比逐个替换分隔符更省且覆盖所有分隔符。
_FW_MAP = {c: chr(c - 0xFEE0) for c in range(0xFF01, 0xFF5F)}

# 括号类统一为圆括号：别名常录 「苯并（a,h）蒽」，LIMS 组分名却用 「苯并[a,h]蒽」
_BR_MAP = {'[': '(', ']': ')', '{': '(', '}': ')', '（': '(', '）': ')', '【': '(', '】': ')'}


def norm_component(s):
    """LIMS 组分名/别名段名匹配归一：NFKC(全角→半角) + 去空白与逗号 + 括号统一为圆括号 + 小写。
    覆盖 「苯并（a,h）蒽」(别名) vs 「苯并[a,h]蒽」(LIMS) 这类录入风格差异；
    谐音错字(茚苯/茚并)不归一——归不出来，靠未匹配告警暴露。"""
    s = unicodedata.normalize('NFKC', str(s or ''))
    return ''.join(_BR_MAP.get(ch, ch) for ch in s
                   if not ch.isspace() and ch not in (',', '，', '、')).lower()


def parse_alias(alias_str):
    """解析项目别名字符串 -> 段列表。

    每段: {'lims_component': str|None, 'compound': str, 'blank_value': str|None,
          'internal_std': str|None, 'range': (lo,hi)|None}
    """
    alias_str = (alias_str or '').strip()
    if not alias_str:
        return []
    alias_str = alias_str.translate(_FW_MAP)  # 全角；＝｜／＃ → 半角
    segments = []
    for raw_seg in alias_str.split('|'):
        seg = raw_seg.strip()
        if not seg:
            continue
        # 多组分：左侧 lims_component = 右侧 compound 规则
        lims_component = None
        if '=' in seg:
            left, seg = seg.split('=', 1)
            lims_component = left.strip()
        # 线性范围：先抽出
        rng = None
        m = _RANGE_RE.search(seg)
        if m:
            rng = (float(m.group(1)), float(m.group(2)))
            seg = _RANGE_RE.sub('', seg, count=1)
        # 内标 #
        internal_std = None
        if '#' in seg:
            seg, internal_std = seg.split('#', 1)
            internal_std = internal_std.strip() or None
        # 空白值 ;
        blank_value = None
        if ';' in seg:
            compound, blank_value = seg.split(';', 1)
            compound = compound.strip()
            blank_value = blank_value.strip() or None
        else:
            compound = seg.strip()
        segments.append({
            'lims_component': lims_component,
            'compound': compound,
            'blank_value': blank_value,
            'internal_std': internal_std,
            'range': rng,
        })
    return segments


def _norm(name):
    """化合物名归一：去定量离子后缀(DBP-149→DBP)、去空白/连字符、小写。
    覆盖 MassHunter「名+定量离子」报告变体与大小写/连字差异。"""
    # ponytail: MassHunter 化合物名常带 -NN 定量离子且随批次/配置有无；剥末尾 -数字 后按基名匹配，
    # 使别名(带离子)与报告(带/不带离子、离子号不同)互通。仅作用于末尾 -纯数字，不伤 2,6-TDI/Pb 220.353 等。
    s = re.sub(r'-\d+$', '', str(name or ''))
    return re.sub(r'[\s\-]+', '', s).lower()


def _lookup(compounds, compound):
    """在解析结果里按精确名→归一名查化合物；返回 {'value','status'} 或 None。"""
    if compound in compounds:
        return compounds[compound]
    target = _norm(compound)
    for name, info in compounds.items():
        if _norm(name) == target:
            return info
    return None


def _fmt(num):
    """数值去尾零：0.360000 -> 0.36。"""
    s = f"{num:.6f}".rstrip('0').rstrip('.')
    return s if s else '0'


def _parse_limit(blank_value, default):
    """从 blank 值 '<0.005'/'＜0.005' 解析检出限；解析不出返回 default。"""
    if not blank_value:
        return default
    m = re.search(r'[<＜]\s*([0-9]+(?:\.[0-9]+)?)', str(blank_value))
    return float(m.group(1)) if m else default


def evaluate_alias(alias_str, compounds, nd_threshold=0.05, diluted_compounds=None):
    """按项目别名规则对解析出的化合物求值，返回每段一个结果。

    返回 [{'lims_component': str|None, 'value': str, 'raw': {...}}]。
      - value: 最终填报字符串（数值，或 blank 字面值如 '<0.005'）
      - compounds: {化合物名: {'value': float, 'status': str}}（来自 report_parser.parse_pdf_report）
      - nd_threshold: 固定兜底值(0.05)，仅当别名未给 `<X` 检出限时用作 ND 字面值/阈值；
        正常情况 ND 判断一律以别名 `<X` 为准（方法编辑器已移除 N.D判断值字段）。
      - diluted_compounds: 稀释报告(-NNX)的化合物表；某段正常值 > 线性范围上限时改用其中读数(原值不乘)，
        并在该段 raw 标记 'diluted'=True 供消费者填稀释列。None 则不切换。
    """
    out = []
    for seg in parse_alias(alias_str):
        looked = _lookup(compounds, seg['compound'])
        raw = {'compound': seg['compound'], 'found': looked is not None}
        nd_literal = seg['blank_value'] or '<%s' % _fmt(nd_threshold)
        limit = _parse_limit(seg['blank_value'], nd_threshold)  # 检出限取自别名 <X
        # 过线性范围上界 → 改用稀释报告读数（不乘；LIMS 据稀释列自算）
        if (seg['range'] and looked is not None
                and looked.get('status') == '检出'
                and float(looked.get('value') or 0) > seg['range'][1]
                and diluted_compounds is not None):
            dlooked = _lookup(diluted_compounds, seg['compound'])
            if dlooked is not None:
                looked = dlooked          # 换源；后续阈值/格式/越界逻辑照常跑
                raw['diluted'] = True     # 供消费者填稀释列
        if looked is None or looked.get('status') in ('未检出', '未校正'):
            # 未检出/未校正/谱图无此化合物 -> 报出 ND 字面值
            raw['status'] = (looked or {}).get('status', '未检出')
            value = nd_literal
        else:
            num = float(looked.get('value') or 0.0)
            raw['status'] = '检出'
            if num < limit:
                value = nd_literal  # 检出但低于检出限 -> 报 <限
            else:
                value = looked.get('raw') or _fmt(num)  # 保留报告原始有效位/末尾0(如0.100)
                if seg['range'] and (num < seg['range'][0] or num > seg['range'][1]):
                    raw['out_of_range'] = (num, seg['range'])  # 越界仅标记，不改值
        out.append({'lims_component': seg['lims_component'], 'value': value, 'raw': raw})
    return out


if __name__ == '__main__':
    # ponytail: 自检——覆盖多组分/;/#/[] 四种形式 + 求值（ND/检出/小于阈值）。
    segs = parse_alias('邻苯-1=Acenaphthylene;<0.01#Naphthalene-d8|邻苯-2=Acenaphthene;<0.01#Naphthalene-d8')
    assert len(segs) == 2, segs
    assert segs[0]['lims_component'] == '邻苯-1' and segs[0]['compound'] == 'Acenaphthylene'
    assert segs[0]['blank_value'] == '<0.01' and segs[0]['internal_std'] == 'Naphthalene-d8'

    segs2 = parse_alias('Naphthalene;<0.005/[0.005-2.00]')
    assert len(segs2) == 1 and segs2[0]['range'] == (0.005, 2.00), segs2
    assert segs2[0]['blank_value'] == '<0.005', segs2

    # 未检出 -> blank 字面值
    r = evaluate_alias('Naphthalene;<0.005', {'Naphthalene': {'value': 0.0, 'status': '未检出'}})
    assert r[0]['value'] == '<0.005', r
    # 检出且 >= 阈值 -> 数值
    r = evaluate_alias('DEHP;<0.05', {'DEHP': {'value': 0.36, 'status': '检出'}})
    assert r[0]['value'] == '0.36', r
    # 检出但 < 阈值 -> ND 字面值
    r = evaluate_alias('X;<0.05', {'X': {'value': 0.02, 'status': '检出'}})
    assert r[0]['value'] == '<0.05', r
    # 多组分：一个检出一个未检出
    r = evaluate_alias('a=DBP;<0.05|b=BBP;<0.05',
                       {'DBP': {'value': 0.1, 'status': '检出'}, 'BBP': {'value': 0.0, 'status': '未检出'}})
    assert r[0]['value'] == '0.1' and r[1]['value'] == '<0.05', r
    # 化合物名大小写/连字差异应能匹配
    r = evaluate_alias('Naphthalene', {'Naphthalene ': {'value': 0.5, 'status': '检出'}})
    assert r[0]['value'] == '0.5', r
    # 检出限取自别名 <X：检出值高于限->数值；低于限->报 <限（修复 0.05 硬编码误判）
    r = evaluate_alias('Phenanthrene;<0.005', {'Phenanthrene': {'value': 0.026, 'status': '检出'}})
    assert r[0]['value'] == '0.026', r

    # 组分名匹配归一：全角/半角括号、方括号、空白/逗号差异应相等；不同化合物不等
    assert norm_component('苯并（a,h）蒽') == norm_component('苯并[a,h]蒽')
    assert norm_component('茚苯（123-cd）芘') == norm_component('茚苯[1,2,3-cd]芘')
    assert norm_component('苯并（a）蒽') != norm_component('苯并（a）芘')
    print('alias_evaluator self-check OK')
    r = evaluate_alias('Anthracene;<0.005', {'Anthracene': {'value': 0.003, 'status': '检出'}})
    assert r[0]['value'] == '<0.005', r
    # 超线性(>范围上限) → 改用稀释报告读数(原值不乘)，并标记 diluted
    r = evaluate_alias('Naphthalene;<0.005/[0.005-2.00]',
                       {'Naphthalene': {'value': 3.0, 'status': '检出', 'raw': '3.0'}},
                       diluted_compounds={'Naphthalene': {'value': 0.064, 'status': '检出', 'raw': '0.064'}})
    assert r[0]['value'] == '0.064' and r[0]['raw'].get('diluted') is True, r
    # 在范围内 → 用正常值，不标记 diluted
    r = evaluate_alias('Naphthalene;<0.005/[0.005-2.00]',
                       {'Naphthalene': {'value': 0.5, 'status': '检出', 'raw': '0.5'}},
                       diluted_compounds={'Naphthalene': {'value': 0.064, 'status': '检出', 'raw': '0.064'}})
    assert r[0]['value'] == '0.5' and not r[0]['raw'].get('diluted'), r
    # ICP 别名格式：元素+波长;<空白>;[区间]，第二分号分隔区间（不应留尾分号）
    segs = parse_alias('Pb 220.353;<0.020;[0.01-0.50]')
    assert len(segs) == 1, segs
    assert segs[0]['compound'] == 'Pb 220.353', segs
    assert segs[0]['blank_value'] == '<0.020', segs  # 无尾分号
    assert segs[0]['range'] == (0.01, 0.50), segs
    # 无区间（只化合物+空白）
    assert parse_alias('Pb 220.353;<0.020')[0]['blank_value'] == '<0.020'
    # ICP：未检出 -> <0.020（不带尾分号）；检出值>=限 -> 原值
    r = evaluate_alias('Pb 220.353;<0.020;[0.01-0.50]', {})
    assert r[0]['value'] == '<0.020', r
    r = evaluate_alias('Pb 220.353;<0.020;[0.01-0.50]',
                       {'Pb 220.353': {'value': 0.5, 'status': '检出', 'raw': '0.500'}})
    assert r[0]['value'] == '0.500', r
    # 全角标点（中文前端常录成 ；＝）应按半角分隔符解析：blank_value 不丢、化合物名不粘连
    segs = parse_alias('苯=苯；<2.00/[1.00-50]')
    assert segs[0]['compound'] == '苯' and segs[0]['blank_value'] == '<2.00', segs
    r = evaluate_alias('苯=苯；<2.00/[1.00-50]', {'苯': {'value': 0.0, 'status': '未检出'}})
    assert r[0]['value'] == '<2.00', r  # 用别名的 <2.00，而非默认 <0.05
    # 定量离子后缀归一：别名化合物带 -NN 离子，报告带/不带离子、离子号不同都应按基名命中
    r = evaluate_alias('BBP=BBP-206;<0.50', {'BBP-149': {'value': 1.29, 'status': '检出', 'raw': '1.29'}})
    assert r[0]['value'] == '1.29', r  # BBP-206 别名命中 BBP-149 报告
    r = evaluate_alias('DBP=DBP-149;<0.50', {'DBP': {'value': 1.05, 'status': '检出', 'raw': '1.05'}})
    assert r[0]['value'] == '1.05', r  # 带离子别名命中不带离子报告
    print('alias_evaluator 自检通过')
