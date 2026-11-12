# -*- coding: utf-8 -*-
"""同步引擎：井号匹配 + 冲突决策。

决策矩阵（方案 §4.3 / §13.4）：
    ① 原值为空或 "/"            → 写入
    ② 原值"审核中"，新值日期     → 写入（审核中转已完成）
    ③ 原值日期，新值"审核中"     → 默认不覆盖（--allow-recheck-overwrite 可开）
    ④ 新旧相同                  → 跳过（幂等）
    ⑤ 新旧都是日期              → 取较新者
    ⑥ 明细表同井号多条          → 取流程次数最大者
    ⑦ 同 sheet 同阶段重复井号   → 仅第一条，其余告警
    ⑧ 明细值为空                → 跳过（避免清空跟踪大表现有日期）
"""

from __future__ import annotations

import dataclasses as dc
import datetime
import re

from . import config
from .pathutil import canon, nfc, well_core

# 严格日期：YYYY-M-D 或 YYYY/M/D
_CLEAN_DATE = re.compile(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$")
# 日期前缀：允许后面带时间
_DATE_PREFIX = re.compile(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})")


def date_only(s) -> str:
    """'2026-08-14 10:18:20' / '2026/8/14' → '2026-08-14'。非日期原样返回。"""
    s = nfc(s)
    if not s:
        return ""
    m = _DATE_PREFIX.match(s)
    if m:
        return "%04d-%02d-%02d" % (int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return s


def is_dirty_date(s) -> bool:
    """脏日期判定：空、"/"、不符合 YYYY-M-D、或日期不存在（如 '2025-03-38'）。

    只做格式校验是不够的 —— "2025/3/38" 能过正则但根本不是合法日期，
    若当成干净日期参与字典序比较，会把真实完成日期挤掉。故追加日历校验。
    """
    s = nfc(s)
    if s in ("", "/"):
        return True
    m = _CLEAN_DATE.match(s)
    if not m:
        return True
    try:
        datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return True
    return False


def format_value(payload: str, status: str) -> str:
    """把明细值格式化为写入值。

    已完成 → 'YYYY-MM-DD'
    审核中 → '{当前审核人}（审核中）'
    """
    payload = nfc(payload)
    if status == config.STATUS_REVIEW:
        if not payload:
            return ""
        return payload + config.REVIEW_SUFFIX
    return date_only(payload)


def decide(payload: str, status: str, old: str,
           allow_recheck_overwrite: bool | None = None) -> str | None:
    """冲突决策。返回写入值，或 None 表示跳过。"""
    if allow_recheck_overwrite is None:
        allow_recheck_overwrite = config.ALLOW_RECHECK_OVERWRITE

    # ⑧ 明细无值 → 不清空
    if not nfc(payload):
        return None

    new = format_value(payload, status)
    if not new:
        return None

    old_n = nfc(old)

    # ① 原值为空或 "/"
    if not old_n or old_n == "/":
        return new

    # ② 审核中 → 已完成
    if config.REVIEW_SUFFIX in old_n and status == config.STATUS_DONE:
        return new

    # ⑥ 兜底：原值是脏日期/脏数据 → 覆盖（但同值仍要幂等跳过）
    # 必须排在「已完成→审核中」保护之前，否则"当前审核人员"这类占位符会被
    # 误判为"已完成"而跳过，无法替换成真实的审核人姓名。
    # 真实的完成日期不是脏数据，因此仍受下一条规则保护。
    if is_dirty_date(old_n) and canon(old_n) != canon(new):
        return new

    # ③ 已完成 → 审核中（回退）
    if config.REVIEW_SUFFIX in new and config.REVIEW_SUFFIX not in old_n:
        return new if allow_recheck_overwrite else None

    old_d, new_d = date_only(old_n), date_only(new)

    # ④ 幂等
    if old_d == new_d:
        return None

    # ⑤ 都有日期 → 取较新
    if not is_dirty_date(old_d) and not is_dirty_date(new_d):
        return new if new_d > old_d else None

    return None


@dc.dataclass
class DetailRecord:
    """明细表一行。"""
    well: str            # 比较键（canon 后）
    well_display: str    # 原文
    stage: str
    status: str          # 已完成 / 审核中
    value: str           # 完成日期 或 当前审核人
    flow_count: int = 0
    source_name: str = ""
    measure: str = ""    # 措施类别 / 措施类型（决定新增行落到哪个 sheet）


@dc.dataclass
class Action:
    """一次写入/跳过动作。"""
    well: str
    well_display: str
    sheet: str
    row: int             # 1-based 行号（面向用户）
    stage: str
    target_col: int      # 0-based 列号
    old_value: str
    new_value: str
    action: str          # write / skip / warn-dup-well
    reason: str = ""


def pick_latest(records: list[DetailRecord]) -> DetailRecord:
    """⑥ 同井号多条明细 → 取流程次数最大者。"""
    return max(records, key=lambda r: (r.flow_count, r.source_name))


def unmatched_detail_wells(track_sheets: list, stage_index: dict) -> dict:
    """明细有、但大表所有 sheet 都找不到对应井号的井。

    返回 {stage: [well_display, ...]}。这些井在大表里没有行，
    工具只填已有行、不新建行，所以无法写入——这是"数据没写进去"最常见的真因，
    必须显式提示，避免用户误以为工具出错了。

    track_sheets: List[TrackSheet]（提供 .sheet / .name / .locate / .cell）。
    """
    # 收集每个阶段在大表里出现过的井号（canon 键）
    book_wells: dict[str, set] = {}
    for ts in track_sheets:
        for stage in stage_index:
            cols = ts.locate(stage)
            if cols is None:
                continue
            cw = cols["well"]
            s = book_wells.setdefault(stage, set())
            sh = ts.sheet
            for r in range(config.ROW_DATA_START, sh.nrows):
                w = well_core(ts.cell(r, cw))
                if w:
                    s.add(w)

    out: dict[str, list] = {}
    for stage, well2rec in stage_index.items():
        bset = book_wells.get(stage, set())
        miss = [r.well_display for r in well2rec.values() if r.well not in bset]
        if miss:
            out[stage] = miss
    return out


def load_measure_map(track_sheets) -> tuple[dict, str]:
    """从跟踪大表的「对照表」sheet 读取「措施类别 → 对应表单」映射。

    对照表示例（第 1 列为类别、第 2 列为表名关键字）：
        措施类别      | 对应表单
        永久性封井    | 封堵井
        老井回采      | 上试井
        一类井口整改  | 井口整改
        其他          | 修井作业   ← 作为兜底表关键字

    没有对照表 / 对照表为空时，回退到 config 里的默认规则，
    这样规则随大表走，用户改对照表即可生效，无需改代码。
    """
    mapping: dict[str, str] = {}
    fallback = config.MEASURE_FALLBACK_KEYWORD
    for ts in track_sheets:
        if "对照表" not in ts.name:
            continue
        sh = ts.sheet
        for r in range(sh.nrows):
            k = nfc(ts.cell(r, 0)).strip()
            v = nfc(ts.cell(r, 1)).strip()
            if not k or not v or k in ("措施类别", "措施类型", "对应表单"):
                continue
            if k in ("其他", "其它", "其它类别", "默认"):
                fallback = v
            else:
                mapping[k] = v
        break
    return (mapping or dict(config.MEASURE_SHEET_KEYWORDS)), fallback


def measure_sheet_keyword(measure: str, mapping: dict | None = None,
                          fallback: str | None = None) -> str:
    """措施类别 → 目标 sheet 名应包含的关键字。"""
    m = nfc(measure).strip()
    mp = mapping if mapping is not None else config.MEASURE_SHEET_KEYWORDS
    fb = fallback if fallback is not None else config.MEASURE_FALLBACK_KEYWORD
    for key, kw in mp.items():
        if m == key or (m and key in m):
            return kw
    return fb


def _sheet_by_keyword(track_sheets, keyword: str):
    """找表名包含关键字的第一个 sheet。"""
    for ts in track_sheets:
        if keyword in ts.name:
            return ts
    return None


class _VirtualSheet:
    """待新建的 sheet：列结构沿用模板表，本身没有数据行。"""

    def __init__(self, name: str, template):
        self.name = name
        self._tpl = template

    def well_col(self):
        return self._tpl.well_col()

    def seq_col(self):
        return self._tpl.seq_col()

    def locate(self, stage):
        return self._tpl.locate(stage)

    def last_data_row(self):
        return config.ROW_DATA_START - 1   # 新表从数据起始行开始写

    def max_seq(self):
        return 0                            # 新表序号从 1 开始


def _template_sheet(track_sheets):
    """挑一张业务表作模板：优先「修井作业」表（列最全），排除「对照表」。"""
    for ts in track_sheets:
        if config.MEASURE_FALLBACK_KEYWORD in ts.name:
            return ts
    for ts in track_sheets:
        if "对照表" not in ts.name and ts.well_col() >= 0:
            return ts
    return None


def _new_sheet_name(track_sheets, keyword: str) -> str:
    """按现有编号规律生成新表名，如 '9.井口整改-陈亚妮'。"""
    nums, suffix = [], ""
    for ts in track_sheets:
        m = re.match(r"^\s*(\d+)\.", ts.name)
        if not m:
            continue
        nums.append(int(m.group(1)))
        if "-" in ts.name:
            suffix = "-" + ts.name.split("-", 1)[1]
    n = max(nums) + 1 if nums else 1
    return "%d.%s%s" % (n, keyword, suffix)


def build_append_actions(track_sheets: list, stage_index: dict,
                         fallback_on_missing_sheet: bool | None = None,
                         create_missing_sheet: bool | None = None):
    """为「明细有、大表无」的井生成新增行的写入动作。

    规则（按需求）：
        永久性封井   → 表名含「封堵井」的 sheet
        老井回采     → 表名含「上试井」的 sheet
        一类井口整改 → 表名含「井口整改」的 sheet
        其余措施类别 → 表名含「修井作业」的 sheet

    只写三类字段：序号、井号、各参与节点的「实际完成」（完成时间）列，
    其余列一律留空。每个单元格生成一个 action="write" 的 Action，
    因此可直接复用现有 write_inplace / write_back 链路，无需改协议。

    幂等：新增行写入后，下次扫描该井号已在 book_wells 中，不会重复新增。

    返回 (actions, info)；info 含 by_sheet / no_sheet / total。
    """
    if fallback_on_missing_sheet is None:
        fallback_on_missing_sheet = config.APPEND_FALLBACK_ON_MISSING_SHEET
    if create_missing_sheet is None:
        create_missing_sheet = config.APPEND_CREATE_MISSING_SHEET

    # 1) 大表所有 sheet 里已存在的井号（跨 sheet、跨阶段统一判定）
    book_wells: set[str] = set()
    for ts in track_sheets:
        cw = ts.well_col()
        if cw < 0:
            continue
        for r in range(config.ROW_DATA_START, ts.sheet.nrows):
            w = well_core(ts.cell(r, cw))
            if w:
                book_wells.add(w)

    # 2) 按井号聚合：一口井只新增一行，多个节点的完成时间各填各的列
    #    key 统一再 canon 一次，避免调用方传入未归一化的井号导致重复新增
    by_well: dict[str, dict] = {}
    for stage, well2rec in stage_index.items():
        for well, rec in well2rec.items():
            wkey = well_core(well)
            if wkey in book_wells:
                continue
            item = by_well.setdefault(wkey, {"display": rec.well_display, "recs": []})
            item["recs"].append((stage, rec))

    actions: list[Action] = []
    info: dict = {"by_sheet": {}, "no_sheet": [], "total": 0,
                  "new_sheets": {}}   # new_sheets: {新表名: 模板表名}
    state: dict[str, dict] = {}   # 每个目标 sheet 的追加行号 / 序号游标
    virtuals: dict = {}           # keyword -> _VirtualSheet（同一次运行复用）

    # 规则优先取大表「对照表」，读不到才用 config 默认
    measure_map, fallback_kw = load_measure_map(track_sheets)
    info["measure_map"] = measure_map
    info["fallback_keyword"] = fallback_kw

    for wkey, item in sorted(by_well.items()):
        recs = item["recs"]
        # 措施类别：优先取有值、流程次数最大的那条记录
        best = max(recs, key=lambda p: (bool(p[1].measure), p[1].flow_count,
                                        p[1].source_name))
        measure = nfc(best[1].measure).strip()
        kw = measure_sheet_keyword(measure, measure_map, fallback_kw)
        ts = _sheet_by_keyword(track_sheets, kw)
        if ts is None and create_missing_sheet:
            # 目标表不存在：按模板表头自动新建（同一 keyword 只建一张）
            ts = virtuals.get(kw)
            if ts is None:
                tpl = _template_sheet(track_sheets)
                if tpl is not None:
                    ts = _VirtualSheet(_new_sheet_name(track_sheets, kw), tpl)
                    virtuals[kw] = ts
                    info["new_sheets"][ts.name] = tpl.name
        if ts is None and fallback_on_missing_sheet:
            kw = fallback_kw
            ts = _sheet_by_keyword(track_sheets, kw)
        if ts is None:
            # 目标表不存在：默认跳过并告警，绝不擅自把数据放到别的表
            info["no_sheet"].append({"well": item["display"],
                                     "measure": measure, "keyword": kw})
            continue
        cwell = ts.well_col()
        if cwell < 0:
            info["no_sheet"].append({"well": item["display"], "measure": measure,
                                     "keyword": kw + "（该表无井号列）"})
            continue

        st = state.get(ts.name)
        if st is None:
            st = {"row": ts.last_data_row() + 1, "seq": ts.max_seq()}
            state[ts.name] = st
        r = st["row"]
        st["row"] += 1
        st["seq"] += 1

        cells: dict[int, str] = {}
        cseq = ts.seq_col()
        if cseq >= 0:
            cells[cseq] = str(st["seq"])
        cells[cwell] = item["display"]
        for stage, rec in recs:
            cols = ts.locate(stage)
            if cols is None or cols["actual"] < 0:
                continue
            val = format_value(rec.value, rec.status)
            if val:
                cells[cols["actual"]] = val

        label = measure or "未填写"
        for c, v in sorted(cells.items()):
            actions.append(Action(
                well=wkey, well_display=item["display"], sheet=ts.name,
                row=r + 1, stage="新增井号", target_col=c,
                old_value="", new_value=v, action="write",
                reason="按措施类别「%s」新增到含「%s」的表（序号/井号/完成时间）"
                       % (label, kw),
            ))
        slot = info["by_sheet"].setdefault(ts.name, {"count": 0, "wells": []})
        slot["count"] += 1
        slot["wells"].append(item["display"])
        info["total"] += 1

    return actions, info


def build_stage_index(records: list[DetailRecord]) -> dict:
    """构建 {stage: {well_key: DetailRecord}}。"""
    by_stage: dict[str, dict[str, list[DetailRecord]]] = {}
    for r in records:
        by_stage.setdefault(r.stage, {}).setdefault(r.well, []).append(r)
    return {
        stage: {w: pick_latest(v) for w, v in wells.items()}
        for stage, wells in by_stage.items()
    }


def sync_sheet(track_sheet, stage_index: dict,
               allow_recheck_overwrite: bool | None = None) -> list[Action]:
    """在一个 sheet 上执行同步。track_sheet 需提供 .sheet / .name / .columns_for()。"""
    actions: list[Action] = []
    sh = track_sheet.sheet
    n = sh.nrows

    for stage, well2rec in stage_index.items():
        cols = track_sheet.locate(stage)
        if cols is None:
            continue  # 该 sheet 无此阶段，正常（如 sheet 6/8 无工艺设计）

        col_well, col_actual = cols["well"], cols["actual"]
        seen: set[str] = set()   # ⑦ 按阶段去重，不是按 sheet

        for r in range(config.ROW_DATA_START, n):
            try:
                well_disp = track_sheet.cell(r, col_well)
            except IndexError:
                continue
            well = well_core(well_disp)
            if not well:
                continue

            if well in seen:
                actions.append(Action(
                    well, well_disp, track_sheet.name, r + 1, stage, col_actual,
                    "", "", "warn-dup-well", "同 sheet 同阶段出现重复井号，仅第一条写入"))
                continue
            seen.add(well)

            rec = well2rec.get(well)
            if rec is None:
                continue

            old = track_sheet.cell(r, col_actual)
            new = decide(rec.value, rec.status, old, allow_recheck_overwrite)
            if new is None:
                actions.append(Action(
                    well, well_disp, track_sheet.name, r + 1, stage, col_actual,
                    old, "", "skip", _skip_reason(rec, old)))
            else:
                actions.append(Action(
                    well, well_disp, track_sheet.name, r + 1, stage, col_actual,
                    old, new, "write", _write_reason(rec, old, new)))

    return actions


def _skip_reason(rec: DetailRecord, old: str) -> str:
    if not nfc(rec.value):
        return "明细无值，跳过以免清空"
    if nfc(old) in ("", "/"):
        return "无变化"
    if config.REVIEW_SUFFIX in nfc(new_v := format_value(rec.value, rec.status)):
        if config.REVIEW_SUFFIX not in nfc(old):
            return "已有完成日期，审核中不覆盖（可开 --allow-recheck-overwrite）"
    if date_only(old) == date_only(new_v):
        return "日期相同，幂等跳过"
    if date_only(old) > date_only(new_v):
        return "原值日期更新，保留"
    return "无变化"


def _write_reason(rec: DetailRecord, old: str, new: str) -> str:
    if not nfc(old) or nfc(old) == "/":
        return "原值为空，写入"
    if config.REVIEW_SUFFIX in nfc(old):
        return "审核中 → 已完成"
    if is_dirty_date(nfc(old)):
        return "原值为脏日期，覆盖"
    if date_only(new) > date_only(old):
        return "新日期更晚，更新"
    return "写入"
