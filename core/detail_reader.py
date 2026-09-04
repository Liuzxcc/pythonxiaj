# -*- coding: utf-8 -*-
"""设计明细表读取：把"已完成 / 审核中"表加载为 DetailRecord 列表。

两类表的列结构不同，统一按表头文本定位：
    已完成表: 井号 / 单位 / 措施类别 / 资金来源 / 设计单位 / 设计人 / 流程次数 / 设计日期 / 完成日期 / 备注
    审核中表: 井号 / 单位 / 措施类型 / 资金来源 / 当前审核人 / 设计单位 / 设计人 / 设计日期 / 上报日期 / 备注
"""

from __future__ import annotations

import pathlib

import xlrd

from . import config
from .filename_parser import ParsedFile
from .pathutil import canon, nfc
from .sync_engine import DetailRecord


def _cell(sh, r: int, c: int) -> str:
    """读取单元格并转为字符串；日期类型格式化。"""
    if c < 0 or c >= sh.ncols or r >= sh.nrows:
        return ""
    v = sh.cell_value(r, c)
    if sh.cell_type(r, c) == xlrd.XL_CELL_DATE:
        try:
            import datetime
            dt = xlrd.xldate_as_datetime(v, 0)
            if dt.hour or dt.minute or dt.second:
                return dt.strftime("%Y-%m-%d %H:%M:%S")
            return dt.strftime("%Y-%m-%d")
        except Exception:
            return str(v)
    if v is None:
        return ""
    return str(v).strip()


def find_col(sh, header_row: int, candidates: list[str]) -> int:
    """按表头文本找列号。先精确后包含。找不到抛 KeyError。"""
    headers = [
        nfc(_cell(sh, header_row, c)).replace("\u3000", "").replace("\n", "").strip()
        for c in range(sh.ncols)
    ]
    for i, h in enumerate(headers):
        if h in candidates:
            return i
    for i, h in enumerate(headers):
        for cand in candidates:
            if cand and (cand in h or h in cand):
                return i
    raise KeyError("未找到列 %s（第 %d 行），实际表头=%s" % (candidates, header_row, headers))


def load_detail(pf: ParsedFile) -> list[DetailRecord]:
    """加载一份明细表为 DetailRecord 列表。"""
    path = pathlib.Path(pf.path)
    if path.suffix.lower() != ".xls":
        # .xlsx 暂不支持（xlrd>=2 只支持 xls）
        return []

    book = xlrd.open_workbook(str(path), formatting_info=False)
    sh = book.sheet_by_index(0)

    try:
        col_well = find_col(sh, config.ROW_DETAIL_HEADER, config.DETAIL_COL_WELL)
    except KeyError:
        return []

    # 取值列：按 status 决定取"完成日期"还是"当前审核人"
    if pf.status == config.STATUS_REVIEW:
        try:
            col_value = find_col(sh, config.ROW_DETAIL_HEADER, config.DETAIL_COL_REVIEWER)
        except KeyError:
            return []
    else:
        try:
            col_value = find_col(sh, config.ROW_DETAIL_HEADER, config.DETAIL_COL_DONE_DATE)
        except KeyError:
            # 退而求其次用"上报日期"
            try:
                col_value = find_col(sh, config.ROW_DETAIL_HEADER, config.DETAIL_COL_SUBMIT_DATE)
            except KeyError:
                return []

    try:
        col_flow = find_col(sh, config.ROW_DETAIL_HEADER, config.DETAIL_COL_FLOW)
    except KeyError:
        col_flow = -1

    out: list[DetailRecord] = []
    for r in range(config.ROW_DETAIL_HEADER + 1, sh.nrows):
        well_disp = _cell(sh, r, col_well)
        well = canon(well_disp)
        if not well:
            continue
        flow = 0
        if col_flow >= 0:
            try:
                flow = int(float(_cell(sh, r, col_flow)))
            except (TypeError, ValueError):
                flow = 0
        out.append(DetailRecord(
            well=well,
            well_display=well_disp,
            stage=pf.stage,
            status=pf.status,
            value=_cell(sh, r, col_value),
            flow_count=flow,
            source_name=path.name,
        ))
    return out


def load_all(parsed_files: list[ParsedFile]) -> list[DetailRecord]:
    """批量加载多份明细表。"""
    records: list[DetailRecord] = []
    for pf in parsed_files:
        records.extend(load_detail(pf))
    return records
