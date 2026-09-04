# -*- coding: utf-8 -*-
"""sheet 7 列结构调整：把"工艺设计"下方的单列补成「计划完成 / 实际完成」两列。

背景：v1.0 模板中 sheet 7 的"工艺设计"只有一列，二级表头叫"当前进度"，
      与地质设计/工程设计等阶段的"计划完成 / 实际完成"两列不一致。

实现约束：xlwt 不能"插列"，所以必须整张 sheet 重建 ——
    xlrd 读全部单元格 → 内存中插入空列 → xlwt 重写到新 .xls。
原表永不被本模块直接修改（调用方决定何时用 --apply 覆盖）。

变更前（sheet 7，18 列）:
    L=工艺设计/当前进度   M=修前工程/计划完成   N=修前工程/实际完成 ...
变更后（19 列）:
    L=工艺设计/实际完成   M=工艺设计/计划完成(新增)  N=修前工程/计划完成 ...
"""

from __future__ import annotations

import csv
import dataclasses as dc
import datetime
import pathlib

import xlwt

from . import config
from .detail_reader import _cell
from .pathutil import nfc


@dc.dataclass
class SheetSnapshot:
    name: str
    rows: list[list]          # rows[r][c]，元素为原值
    ncols_old: int = 0


def snapshot(sheet) -> SheetSnapshot:
    """把 xlrd sheet 抓成内存快照。"""
    rows = []
    for r in range(sheet.nrows):
        row = []
        for c in range(sheet.ncols):
            v = sheet.cell_value(r, c)
            if sheet.cell_type(r, c) == 3:  # XL_CELL_DATE
                try:
                    import xlrd
                    v = xlrd.xldate_as_datetime(v, 0)
                except Exception:
                    pass
            row.append(v)
        rows.append(row)
    return SheetSnapshot(name=sheet.name, rows=rows, ncols_old=sheet.ncols)


@dc.dataclass
class ColumnPlan:
    """列重排计划。"""
    insert_at: int | None     # 在哪个新列号处插入空列；None 表示无需变更
    old_to_new: dict[int, int]
    actual_col: int
    plan_col: int


def plan_columns(snap: SheetSnapshot, stage: str | None = None) -> ColumnPlan:
    """根据一级/二级表头文本，规划列变更。"""
    stage = stage or config.STAGE_NEEDS_TWO_COLS
    if len(snap.rows) <= config.ROW_SUB_HEADER:
        raise ValueError("sheet 行数不足，无法规划列结构")

    r2 = [nfc(_row_at(snap, config.ROW_STAGE_HEADER, c)) for c in range(snap.ncols_old)]
    r3 = [nfc(_row_at(snap, config.ROW_SUB_HEADER, c)) for c in range(snap.ncols_old)]

    # 按一级表头切块
    blocks: list[tuple[str, list[int]]] = []
    cur_text, cur_cols = "", []
    for c, t in enumerate(r2):
        if t:
            if cur_cols:
                blocks.append((cur_text, cur_cols))
            cur_text, cur_cols = t, [c]
        else:
            cur_cols.append(c)
    if cur_cols:
        blocks.append((cur_text, cur_cols))

    target = next((cols for text, cols in blocks if text == nfc(stage)), None)
    if target is None:
        raise ValueError("未找到一级表头 %r" % stage)

    actual_col = target[0]
    # 已经是标准结构（该块内已含"实际完成"）→ 无需变更
    if any(r3[c] == config.COL_NAME_ACTUAL for c in target if c < len(r3)):
        return ColumnPlan(None, {c: c for c in range(snap.ncols_old)},
                          actual_col, actual_col + 1)

    plan_col = actual_col + 1
    old_to_new = {c: (c if c < plan_col else c + 1) for c in range(snap.ncols_old)}
    return ColumnPlan(plan_col, old_to_new, actual_col, plan_col)


def _row_at(snap: SheetSnapshot, r: int, c: int):
    if r < len(snap.rows) and c < len(snap.rows[r]):
        return snap.rows[r][c]
    return ""


def apply_plan(snap: SheetSnapshot, plan: ColumnPlan,
               stage: str | None = None) -> SheetSnapshot:
    """按计划重建行数据，并修正二级表头文本。"""
    stage = stage or config.STAGE_NEEDS_TWO_COLS
    if plan.insert_at is None:
        return snap

    new_ncols = max(plan.old_to_new.values()) + 1
    new_rows = []
    for row in snap.rows:
        new_row = [None] * new_ncols
        for old_c, val in enumerate(row):
            new_c = plan.old_to_new.get(old_c)
            if new_c is not None and new_c < new_ncols:
                new_row[new_c] = val
        new_rows.append(new_row)

    # 修正表头：actual_col → "实际完成"，plan_col → "计划完成"，一级表头都归该 stage
    def _set(r: int, c: int, v):
        while len(new_rows) <= r:
            new_rows.append([None] * new_ncols)
        new_rows[r][c] = v

    _set(config.ROW_STAGE_HEADER, plan.actual_col, stage)
    _set(config.ROW_STAGE_HEADER, plan.plan_col, stage)
    _set(config.ROW_SUB_HEADER, plan.actual_col, config.COL_NAME_ACTUAL)
    _set(config.ROW_SUB_HEADER, plan.plan_col, config.COL_NAME_PLAN)

    return SheetSnapshot(name=snap.name, rows=new_rows, ncols_old=snap.ncols_old)


def write_snapshots(snaps: list[SheetSnapshot], out_path: pathlib.Path) -> pathlib.Path:
    """把若干快照写成一份 .xls。"""
    out_path = pathlib.Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wbk = xlwt.Workbook(encoding="utf-8")
    header_xf = xlwt.easyxf("font: bold on; align: wrap on, vert center, horiz center;")
    data_xf = xlwt.easyxf()

    for snap in snaps:
        ws = wbk.add_sheet(snap.name[:31])
        ncols = max((len(r) for r in snap.rows), default=0)
        for r, row in enumerate(snap.rows):
            for c in range(ncols):
                v = row[c] if c < len(row) else None
                if v is None or v == "":
                    continue
                if isinstance(v, datetime.datetime):
                    v = v.strftime("%Y-%m-%d")
                ws.write(r, c, v,
                         header_xf if r <= config.ROW_SUB_HEADER else data_xf)
        for c in range(min(ncols, 30)):
            ws.col(c).width = 14 * 256

    wbk.save(str(out_path))
    return out_path


def col_letter(idx: int) -> str:
    s, idx = "", idx + 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def write_report(before: SheetSnapshot, after: SheetSnapshot,
                 plan: ColumnPlan, dst: pathlib.Path) -> pathlib.Path:
    """输出列结构变更前后对照 CSV。"""
    dst = pathlib.Path(dst)
    dst.parent.mkdir(parents=True, exist_ok=True)
    ncols_after = max((len(r) for r in after.rows), default=0)
    with dst.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["新列号", "新列字母", "新一级表头", "新二级表头",
                    "来源旧列号", "来源旧一级表头", "来源旧二级表头", "备注"])
        inv = {v: k for k, v in plan.old_to_new.items()}
        for c in range(ncols_after):
            h1 = nfc(_row_at(after, config.ROW_STAGE_HEADER, c))
            h2 = nfc(_row_at(after, config.ROW_SUB_HEADER, c))
            old_c = inv.get(c)
            if old_c is None:
                note = "新增列（%s）" % config.COL_NAME_PLAN
                o1 = o2 = ""
                src = ""
            else:
                o1 = nfc(_row_at(before, config.ROW_STAGE_HEADER, old_c))
                o2 = nfc(_row_at(before, config.ROW_SUB_HEADER, old_c))
                src = "%s(%d)" % (col_letter(old_c), old_c)
                note = ("表头改写 当前进度→实际完成" if o2 == config.COL_NAME_ACTUAL_LEGACY
                        else ("右移" if old_c != c else ""))
            w.writerow([c, col_letter(c), h1, h2, src, o1, o2, note])
    return dst


def find_sheet7(names: list[str]) -> str | None:
    """按 config.SHEET7_MARKERS 定位 sheet 7。"""
    for n in names:
        if any(m in nfc(n) for m in config.SHEET7_MARKERS):
            return n
    return None
