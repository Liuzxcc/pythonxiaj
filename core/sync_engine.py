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
from .pathutil import canon, nfc

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

    # 兜底：原值是脏日期 → 覆盖
    if is_dirty_date(old_n):
        return new

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
            well = canon(well_disp)
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
