# -*- coding: utf-8 -*-
"""写回 .xls。

两种模式：

1. **原地写入**（inplace=True，默认）
   用 xlrd(formatting_info=True) + xlutils.copy 复制原表，只改目标单元格，
   **保留合并单元格、字体、边框、对齐、列宽**。
   先备份原表到备份目录，再写临时文件并原子替换原表。
   源目录不会多出 .synced-*.xls。

2. **另存副本**（inplace=False）
   用 xlwt 从头重建，产出 .synced-{ts}.xls，原表不动。
   ⚠️ 此模式会丢失合并单元格与条件格式，仅用于只读核对场景。

xlwt 本身不能原地改单元格，两种模式本质都是「读入内存 → 应用变更 → 重写整份」，
区别只在于是否沿用原表的格式信息。
"""

from __future__ import annotations

import datetime
import pathlib
import shutil
import time

import xlrd
import xlwt

from . import config
from .sync_engine import Action
from .trackbook import TrackBook

# 改动的单元格用红色字体标记，方便人工扫一眼
_CHANGED_XF = "font: colour_index red;"


def _style_map():
    """另存副本模式用的基础样式。"""
    return {
        "header": xlwt.easyxf("font: bold on; align: wrap on, vert center, horiz center;"),
        "data": xlwt.easyxf(),
        "changed": xlwt.easyxf(_CHANGED_XF),
    }


def _collect_plan(actions: list[Action]) -> dict[str, dict[tuple[int, int], str]]:
    """把 write 动作整理成 {sheet: {(row0, col): new_value}}。"""
    plan: dict[str, dict[tuple[int, int], str]] = {}
    for a in actions:
        if a.action != "write":
            continue
        plan.setdefault(a.sheet, {})[(a.row - 1, a.target_col)] = a.new_value
    return plan


def _write_value(ws, r: int, c: int, val, style) -> None:
    if val is None or val == "":
        return
    if isinstance(val, datetime.datetime):
        ws.write(r, c, val.strftime("%Y-%m-%d %H:%M:%S"), style)
    else:
        ws.write(r, c, val, style)


def xlrd_date_to_str(v) -> str:
    """xlrd 日期序列号 → 'YYYY-MM-DD'。"""
    dt = xlrd.xldate_as_datetime(v, 0)
    if dt.hour or dt.minute or dt.second:
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    return dt.strftime("%Y-%m-%d")


def _ts() -> str:
    return time.strftime("%Y%m%d-%H%M%S")


# ------------------------------------------------------------------ 原地写入

def write_inplace(track: TrackBook, actions: list[Action],
                  backup_dir: pathlib.Path | None = None,
                  make_backup: bool = True):
    """原地写入原表（保留格式）。返回 (原表路径, 备份路径, 是否保格式)。

    make_backup=False 时不生成任何备份文件（「真正零新文件」模式）。
    """
    plan = _collect_plan(actions)
    src = pathlib.Path(track.path)

    # 备份优先：任何写操作之前先把原表存一份（可选）
    backup = _backup(src, backup_dir) if make_backup else None

    tmp = src.with_name("%s.tmp-%s.xls" % (src.stem, _ts()))
    try:
        kept_format = _render_preserving(src, plan, tmp)
        shutil.copystat(src, tmp)          # 保留文件时间/权限
        import os
        os.replace(str(tmp), str(src))     # 原子替换：要么整个换掉，要么原表分毫不动
        return src, backup, kept_format
    finally:
        if tmp.exists():
            tmp.unlink()


def _backup(src: pathlib.Path, backup_dir) -> pathlib.Path:
    """把原表复制到备份目录。未指定备份目录时与源表同目录。"""
    dst_dir = pathlib.Path(backup_dir) if backup_dir else src.parent
    dst_dir.mkdir(parents=True, exist_ok=True)
    dst = dst_dir / ("%s.bak-%s.xls" % (src.stem, _ts()))
    shutil.copy2(src, dst)
    return dst


def _render_preserving(src: pathlib.Path, plan: dict, dst: pathlib.Path) -> bool:
    """用 xlutils 复制原表并应用变更。成功返回 True（保格式）。

    xlutils 不可用、或该 .xls 读不出格式信息时，回退到普通 xlwt 重建（丢格式），
    并返回 False —— 调用方应据此向用户告警。
    """
    rb = None
    try:
        from xlutils.copy import copy as xl_copy
        rb = xlrd.open_workbook(str(src), formatting_info=True)
        wb = xl_copy(rb)
    except Exception:
        _render_plain(src, plan, dst)
        return False

    red = xlwt.easyxf(_CHANGED_XF)
    for idx, name in enumerate(rb.sheet_names()):
        changes = plan.get(name)
        if not changes:
            continue
        ws = wb.get_sheet(idx)
        for (r, c), v in changes.items():
            _write_value(ws, r, c, v, red)

    wb.save(str(dst))
    return True


def _render_plain(src: pathlib.Path, plan: dict, dst: pathlib.Path) -> None:
    """回退方案：xlwt 从头重建（会丢失合并单元格与条件格式）。"""
    rb = xlrd.open_workbook(str(src), formatting_info=False)
    styles = _style_map()
    wb = xlwt.Workbook(encoding="utf-8")
    for name in rb.sheet_names():
        sh = rb.sheet_by_name(name)
        ws = wb.add_sheet(name[:31])
        changes = plan.get(name, {})
        for r in range(sh.nrows):
            for c in range(sh.ncols):
                if (r, c) in changes:
                    _write_value(ws, r, c, changes[(r, c)], styles["changed"])
                else:
                    v = sh.cell_value(r, c)
                    if sh.cell_type(r, c) == 3:      # XL_CELL_DATE
                        try:
                            v = xlrd_date_to_str(v)
                        except Exception:
                            v = str(v)
                    _write_value(ws, r, c, v,
                                 styles["header"] if r < config.ROW_DATA_START
                                 else styles["data"])
    wb.save(str(dst))


# ------------------------------------------------------------------ 另存副本

def write_back(track: TrackBook, actions: list[Action],
               out_path: pathlib.Path | None = None) -> pathlib.Path:
    """另存为 .synced-{ts}.xls（原表不动）。返回新文件路径。

    ⚠️ 用 xlwt 重建，合并单元格/条件格式会丢失，仅供核对。
    """
    plan = _collect_plan(actions)
    if out_path is None:
        out_path = track.path.with_suffix(".synced-%s.xls" % _ts())
    out_path = pathlib.Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    styles = _style_map()
    wbk = xlwt.Workbook(encoding="utf-8")
    for ts_ in track.sheets:
        ws = wbk.add_sheet(ts_.name[:31])
        sh = ts_.sheet
        changes = plan.get(ts_.name, {})
        for r in range(sh.nrows):
            for c in range(sh.ncols):
                if (r, c) in changes:
                    _write_value(ws, r, c, changes[(r, c)],
                                 styles["changed"] if r >= config.ROW_DATA_START
                                 else styles["header"])
                else:
                    v = sh.cell_value(r, c)
                    if sh.cell_type(r, c) == 3:
                        try:
                            v = xlrd_date_to_str(v)
                        except Exception:
                            v = str(v)
                    _write_value(ws, r, c, v,
                                 styles["header"] if r < config.ROW_DATA_START
                                 else styles["data"])
        for c in range(min(sh.ncols, 30)):
            ws.col(c).width = 14 * 256
    wbk.save(str(out_path))
    return out_path
