# -*- coding: utf-8 -*-
"""写回 .xls。

两种模式：

1. **原地写入**（inplace=True，默认）
   用 xlrd(formatting_info=True) + xlutils.copy 复制原表，再统一改写全表格式与目标单元格，
   **保留合并单元格、列宽、行高、窗格等结构**。
   先备份原表到备份目录（可选），再写临时文件并原子替换原表。
   源目录不会多出 .synced-*.xls。

2. **另存副本**（inplace=False）
   用 xlwt 从头重建，产出 .synced-{ts}.xls，原表不动。

v1.6 起默认执行「全表格式统一」：
   所有非空单元格 → 宋体 10 号 / 黑色细边框 / 水平垂直居中；
   日期（Excel 日期序列号、yyyy/m/d、带时间或备注的日期文本）→ yyyy-MM-dd。
   命令行可用 --no-unify-format 关闭，回到「只写变更单元格」的旧行为。

xlwt 本身不能原地改单元格，两种模式本质都是「读入内存 → 应用变更 → 重写整份」，
区别只在于是否沿用原表的格式信息。
"""

from __future__ import annotations

import datetime
import os
import pathlib
import re
import shutil
import time

import xlrd
import xlwt

from . import config
from .sync_engine import Action
from .trackbook import TrackBook

_EMPTY_TYPES = (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK)

# ------------------------------------------------------------------ 统一格式
# 宋体 10 号（xlwt height 单位为 1/20 pt，200 = 10 号）、水平垂直居中、黑色细边框。
# wrap on 必须保留：一级表头含换行（如「修前工程\n及概算完成情况」）。
_ALIGN_BORDER = (
    "align: horiz center, vert center, wrap on;"
    " borders: left thin, right thin, top thin, bottom thin,"
    " left_colour black, right_colour black,"
    " top_colour black, bottom_colour black;"
)
# 普通单元格：黑字（人工录入的数据）
_NORMAL_XF = "font: name 宋体, height 200; " + _ALIGN_BORDER
# 报表工具生成的数据：深蓝色（dark_blue = BIFF8 索引 18 / #000080），其余格式一致
_GENERATED_XF = "font: name 宋体, height 200, colour_index dark_blue; " + _ALIGN_BORDER
# BIFF8 颜色索引：10=red（v1.6.0 及以前工具写入的标记）、18=dark_blue（v1.6.1 起）、
# 32767=automatic（黑）。v1.6.2 起已有单元格的颜色一律保留原值。
_COLOR_LEGACY_RED = 10
_COLOR_GENERATED = 18
_COLOR_AUTO = 32767
# 文本单元格一律强制文本格式，避免日期/自定义数字格式显示异常（(fXt165) 乱码）
_CHANGED_NUM_FMT = "@"
# 按需构造统一样式：只替换字号/字体色，其余属性（宋体/黑框/居中/换行）永远一致
_FONT_XF = "font: name 宋体, height %d, colour_index %d; " + _ALIGN_BORDER

# 字符串里任意位置的日期片段：2026/5/30、2026-5-3、2026.5.3、2026年5月30日。
# 分隔符只吸收空格/制表符，结尾不吸收空白 —— 否则
# 「2026/5/30\n（董5同井场）」里的换行会被当成日期的一部分吃掉，备注与日期挤到一行。
_DATE_TOKEN_RE = re.compile(
    r"(?<!\d)(\d{4})[ \t]*[-/.年][ \t]*(\d{1,2})[ \t]*[-/.月][ \t]*(\d{1,2})日?(?!\d)"
)
# 纯时间尾巴，如明细表常见的 "\n\t\t09:05:39"
_TIME_TAIL_RE = re.compile(r"[\s\u3000]*\d{1,2}:\d{2}(?::\d{2})?\s*$")


def _fmt_date_match(m) -> str:
    """日历校验后再格式化；2025/3/38 这类脏日期保持原样，不当成日期改写。"""
    try:
        d = datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return m.group(0)
    return d.strftime(config.DATE_FMT)


def normalize_date_text(s) -> str:
    """把字符串里所有日期片段统一成 yyyy-MM-dd。

    - 只改日期写法，不丢备注：
      「2026/5/30\\n（董5同井场）」→「2026-05-30\\n（董5同井场）」
    - 多行多个日期逐个归一：「2025/6/19\\n2025/7/23」→ 两个 yyyy-MM-dd
    - 末尾纯时间去掉，满足「日期格式统一为 yyyy-MM-dd」
    - 非日期文本（井号、审核人、备注）原样返回
    """
    if not isinstance(s, str) or not s.strip():
        return s
    if not _DATE_TOKEN_RE.search(s):
        return s
    return _TIME_TAIL_RE.sub("", _DATE_TOKEN_RE.sub(_fmt_date_match, s))




def _uniform_styles() -> dict:
    """4 个默认统一样式：黑字 / 深蓝(工具生成) × 文本 / 数值，另有按需缓存 `_cache`。

    文本用 num_format_str='@'（避免历史自定义格式在 WPS 显示成 (fXt165)）；
    数值保留数字类型（序号等），用 General 显示，不产生「数字以文本形式存储」绿三角。
    """
    return {
        "text": xlwt.easyxf(_NORMAL_XF, num_format_str=_CHANGED_NUM_FMT),
        "num": xlwt.easyxf(_NORMAL_XF, num_format_str="General"),
        "text_gen": xlwt.easyxf(_GENERATED_XF, num_format_str=_CHANGED_NUM_FMT),
        "num_gen": xlwt.easyxf(_GENERATED_XF, num_format_str="General"),
        "_cache": {},
    }


def _xf(styles: dict, text: bool = True, height: int = 200,
        colour: int = _COLOR_AUTO):
    """取统一样式；字号/字体色按传入值覆盖，其它属性永远一致。

    - 表头及标题行传原 height → 不改动它们的字号（v1.6.2 用户要求）
    - 已有单元格传原 colour → 不覆盖人工设置的字体颜色
    - 只有工具本次新增/修改的单元格传 18（深蓝）
    """
    if height == 200 and colour == _COLOR_AUTO:
        return styles["text"] if text else styles["num"]
    if height == 200 and colour == _COLOR_GENERATED:
        return styles["text_gen"] if text else styles["num_gen"]
    cache = styles.setdefault("_cache", {})
    key = (text, height, colour)
    if key not in cache:
        try:
            cache[key] = xlwt.easyxf(
                _FONT_XF % (int(height), int(colour)),
                num_format_str=_CHANGED_NUM_FMT if text else "General")
        except Exception:
            # 极端的旧颜色索引不被 xlwt 接受时退回默认样式，不影响主流程
            cache[key] = (styles["text_gen"] if colour == _COLOR_GENERATED
                          else styles["text"])
    return cache[key]


def _orig_font(rb, sh, r: int, c: int):
    """读原表单元格的 (字号, 字体色)；超出原表范围时返回默认值。"""
    if r >= sh.nrows or c >= sh.ncols:
        return 200, _COLOR_AUTO
    try:
        xf = rb.xf_list[sh.cell_xf_index(r, c)]
        fo = rb.font_list[xf.font_index]
        return (int(fo.height) or 200), int(fo.colour_index)
    except Exception:
        return 200, _COLOR_AUTO


def _cell_style(styles: dict, rb, sh, r: int, c: int, text: bool,
                generated: bool = False):
    """按行号与「是否工具写入」决定字号 / 颜色。

    - r < ROW_DATA_START（标题 + 表头）：字号保留原值，不改成 10 号
    - r >= ROW_DATA_START（数据区）：统一 10 号（height=200）
    - generated=True：字体色强制深蓝；否则一律保留原色
    """
    orig_h, orig_c = _orig_font(rb, sh, r, c)
    height = (orig_h if (r < config.ROW_DATA_START
                         and config.PRESERVE_HEADER_FONT_SIZE) else 200)
    colour = _COLOR_GENERATED
    if not generated:
        colour = (_COLOR_GENERATED
                  if (config.MIGRATE_LEGACY_RED and orig_c == _COLOR_LEGACY_RED)
                  else orig_c)
    return _xf(styles, text=text, height=height, colour=colour)


def _cell_output_value(sh, r: int, c: int, datemode: int):
    """读一个非空单元格 → (写回值, 是否文本样式)。日期统一为 yyyy-MM-dd 文本。"""
    ctype = sh.cell_type(r, c)
    v = sh.cell_value(r, c)
    if ctype == xlrd.XL_CELL_DATE:
        try:
            return xlrd.xldate_as_datetime(v, datemode).strftime(config.DATE_FMT), True
        except Exception:
            return str(v), True
    if ctype == xlrd.XL_CELL_NUMBER:
        # 整数浮点转成 int，避免「1.0」这类显示；仍保持数字类型
        if isinstance(v, float) and v.is_integer():
            return int(v), False
        return v, False
    if ctype in (xlrd.XL_CELL_TEXT, xlrd.XL_CELL_BOOLEAN, xlrd.XL_CELL_ERROR):
        return normalize_date_text(v if isinstance(v, str) else str(v)), True
    return v, True


def _style_map():
    """另存副本 / 纯 xlwt 重建 / 新建 sheet 用的样式表（保留 `_cache` 供 _xf 复用）。"""
    s = _uniform_styles()
    s.update({
        "header": s["text"],
        "data": s["text"],
        "changed": s["text_gen"],
    })
    return s


def _collect_plan(actions: list[Action]) -> dict[str, dict[tuple[int, int], str]]:
    """把 write 动作整理成 {sheet: {(row0, col): new_value}}。"""
    plan: dict[str, dict[tuple[int, int], str]] = {}
    for a in actions:
        if a.action != "write":
            continue
        plan.setdefault(a.sheet, {})[(a.row - 1, a.target_col)] = a.new_value
    return plan


def _plan_max_row(sh, changes: dict) -> int:
    """写回时需要遍历的行数：原表行数与「新增行」行号二者取大。"""
    if not changes:
        return sh.nrows
    return max(sh.nrows, max(r for (r, _c) in changes) + 1)


def _create_sheets(rb, wb, new_sheets: dict, styles=None) -> dict:
    """按模板表的表头结构新建 sheet，返回 {新表名: xlwt worksheet}。

    用于「对照表声明了某类目标表、但大表里没有」的场景。
    只复制表头区（R0 ~ ROW_DATA_START-1）的文本与合并格，数据行留空，
    列结构与模板一致，因此 locate()/well_col() 等按表头定位的逻辑照样可用。

    注意：xlrd 的 merged_cells 是半开区间 (rlo, rhi, clo, chi)，
    而 xlwt.merge 用闭区间，故需 -1。
    """
    styles = styles or _style_map()
    existing = set(rb.sheet_names())
    created: dict = {}
    for new_name, tpl_name in (new_sheets or {}).items():
        if new_name in existing or tpl_name not in existing:
            continue
        tpl = rb.sheet_by_name(tpl_name)
        # 允许覆盖：表头区先铺满带黑框的空格，再由 merge() 重写合并区内空格
        ws = wb.add_sheet(new_name[:31], cell_overwrite_ok=True)
        for r in range(min(config.ROW_DATA_START, tpl.nrows)):
            for c in range(tpl.ncols):
                if (tpl.cell_type(r, c) in _EMPTY_TYPES
                        or tpl.cell_value(r, c) == ""):
                    ws.write(r, c, "", _cell_style(styles, rb, tpl, r, c,
                                                   text=True))
                    continue
                v, is_text = _cell_output_value(tpl, r, c, rb.datemode)
                ws.write(r, c, v, _cell_style(styles, rb, tpl, r, c,
                                              text=is_text))
        merged = getattr(tpl, "merged_cells", None) or ()
        for (r1, r2, c1, c2) in merged:
            if r1 >= config.ROW_DATA_START:
                continue
            try:
                # 带统一样式 merge：合并区域内的空白格也有黑色边框
                ws.merge(r1, r2 - 1, c1, c2 - 1,
                         _cell_style(styles, rb, tpl, r1, c1, text=True))
            except Exception:
                pass
        created[new_name] = ws
    return created


def xlrd_date_to_str(v) -> str:
    """xlrd 日期序列号 → 'YYYY-MM-DD'（兼容旧调用）。"""
    dt = xlrd.xldate_as_datetime(v, 0)
    if dt.hour or dt.minute or dt.second:
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    return dt.strftime("%Y-%m-%d")


def _ts() -> str:
    return time.strftime("%Y%m%d-%H%M%S")


# ------------------------------------------------------------------ 全表格式统一

def _apply_uniform_style(rb, wb, plan: dict, styles: dict,
                         full_grid: bool = True) -> None:
    """整表统一格式；同时把 plan 的变更写进去（深蓝）。

    在 xlutils.copy 的产物上逐格覆盖（xlutils 已把 cell_overwrite_ok 打开）：
      - 合并单元格数量 / 列宽 / 行高 / 窗格由 xlutils 原样保留；
      - full_grid=True 时把「使用区域矩形」内**所有空格**也写成带黑框的空白格，
        整张表是连续网格，中间不会断框（v1.6.1 用户反馈修正）；
      - 工具生成的数据（本次 plan 写入 / 历史标红）统一为深蓝色。
    """
    for si, name in enumerate(rb.sheet_names()):
        sh = rb.sheet_by_index(si)
        ws = wb.get_sheet(si)
        changes = plan.get(name, {})

        # 合并区域非锚点空格：full_grid 关闭时也要补黑框（字号/颜色同样保留原值）
        for (rlo, rhi, clo, chi) in (getattr(sh, "merged_cells", None) or ()):
            for r in range(rlo, rhi):
                for c in range(clo, chi):
                    if (r, c) == (rlo, clo) or (r, c) in changes:
                        continue
                    try:
                        ws.write(r, c, "",
                                 _cell_style(styles, rb, sh, r, c, text=True))
                    except Exception:
                        pass

        # 「整个活动表」= xlrd 读到的完整 UsedRange（含尾部预置空格行）+ 新增行。
        # 不能只画有数据的格子，否则中间整片空白会断框（v1.6.1 用户反馈）。
        rows = _plan_max_row(sh, changes)
        for r in range(rows):
            for c in range(sh.ncols):
                if (r, c) in changes:
                    # 只有工具「本次新增/修改」的单元格才用深蓝色
                    val = normalize_date_text(str(changes[(r, c)]))
                    style = _cell_style(styles, rb, sh, r, c,
                                        text=True, generated=True)
                elif (r < sh.nrows
                      and sh.cell_type(r, c) not in _EMPTY_TYPES
                      and sh.cell_value(r, c) != ""):
                    val, is_text = _cell_output_value(sh, r, c, rb.datemode)
                    # 已有数据：字号按行号决定，字体色一律保留原值
                    style = _cell_style(styles, rb, sh, r, c, text=is_text)
                elif full_grid:
                    # 空白格：写空值 + 黑框，保证整片网格连续；字号/颜色也保留原值
                    val = ""
                    style = _cell_style(styles, rb, sh, r, c, text=True)
                else:
                    continue
                try:
                    ws.write(r, c, val, style)
                except Exception:
                    pass


def _write_plain_sheet(ws, sh, changes: dict, styles: dict, rb,
                       full_grid: bool = True) -> None:
    """纯 xlwt 重建单张 sheet：统一格式 + 整片网格 + 深蓝 + 日期归一 + 合并格。"""
    datemode = rb.datemode
    merged = getattr(sh, "merged_cells", None) or ()
    covered = set()
    for (rlo, rhi, clo, chi) in merged:
        for r in range(rlo, rhi):
            for c in range(clo, chi):
                covered.add((r, c))
        if rlo >= sh.nrows or clo >= sh.ncols or (rhi - 1) >= 65536:
            continue
        key = (rlo, clo)
        if key in changes:
            val = normalize_date_text(str(changes[key]))
            style = _cell_style(styles, rb, sh, rlo, clo,
                                text=True, generated=True)
        elif (sh.cell_type(rlo, clo) not in _EMPTY_TYPES
              and sh.cell_value(rlo, clo) != ""):
            val, is_text = _cell_output_value(sh, rlo, clo, datemode)
            style = _cell_style(styles, rb, sh, rlo, clo, text=is_text)
        else:
            val = ""
            style = _cell_style(styles, rb, sh, rlo, clo, text=True)
        try:
            ws.write_merge(rlo, rhi - 1, clo, chi - 1, val, style)
        except Exception:
            pass

    # 副本模式同样铺满整个 UsedRange（含尾部预置空格行）+ 新增行
    rows = _plan_max_row(sh, changes)
    for r in range(rows):
        for c in range(sh.ncols):
            if (r, c) in covered:
                continue
            if (r, c) in changes:
                val = normalize_date_text(str(changes[(r, c)]))
                style = _cell_style(styles, rb, sh, r, c,
                                    text=True, generated=True)
            elif (r < sh.nrows
                  and sh.cell_type(r, c) not in _EMPTY_TYPES
                  and sh.cell_value(r, c) != ""):
                val, is_text = _cell_output_value(sh, r, c, datemode)
                style = _cell_style(styles, rb, sh, r, c, text=is_text)
            elif full_grid:
                val = ""
                style = _cell_style(styles, rb, sh, r, c, text=True)
            else:
                continue
            try:
                ws.write(r, c, val, style)
            except Exception:
                pass


# ------------------------------------------------------------------ 原地写入

def write_inplace(track: TrackBook, actions: list[Action],
                  backup_dir: pathlib.Path | None = None,
                  make_backup: bool = True,
                  new_sheets: dict | None = None,
                  unify_format: bool = True,
                  full_grid: bool = True):
    """原地写入原表（保留格式）。返回 (原表路径, 备份路径, 是否保格式)。

    make_backup=False 时不生成任何备份文件（「真正零新文件」模式）。
    unify_format=True 时顺带统一全表字体/边框/对齐/日期（v1.6 默认）。
    full_grid=True 时把使用区域矩形内所有空格也加上黑框（v1.6.1 默认）。
    """
    plan = _collect_plan(actions)
    src = pathlib.Path(track.path)

    # 备份优先：任何写操作之前先把原表存一份（可选）
    backup = _backup(src, backup_dir) if make_backup else None

    tmp = src.with_name("%s.tmp-%s.xls" % (src.stem, _ts()))
    try:
        kept_format = _render_preserving(src, plan, tmp,
                                         new_sheets=new_sheets,
                                         unify_format=unify_format,
                                         full_grid=full_grid)
        shutil.copystat(src, tmp)          # 保留文件时间/权限
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


def _sanitize_num_formats(wb):
    """xlutils.copy 复制某些 .xls 时会生成 key 为 None 的自定义数字格式，
    导致 xlwt 在 wb.save() 时抛 `TypeError: descriptor 'decode' ... NoneType`。
    同时清理历史版本误写入的 `'General (fmt%d)'` 等异常格式字符串，
    统一替换为文本格式 '@'，避免 WPS/Excel 把单元格显示成 `(fXt165)` 之类乱码。
    """
    styles = getattr(wb, '_Workbook__styles', None)
    if not styles:
        return
    nf = getattr(styles, '_num_formats', None)
    if not nf:
        return
    for fmtstr, idx in list(nf.items()):
        dirty = (
            fmtstr is None
            or fmtstr == ""
            or "(fmt" in fmtstr          # 历史误写入的占位格式
        )
        if dirty:
            del nf[fmtstr]
            # 统一用文本格式；若 '@' 已被占用，用带索引的变体保留该索引
            new_str = "@"
            if new_str in nf:
                new_str = "@ (fmt%d)" % idx
            nf[new_str] = idx


def _render_preserving(src: pathlib.Path, plan: dict, dst: pathlib.Path,
                       new_sheets: dict | None = None,
                       unify_format: bool = True,
                       full_grid: bool = True) -> bool:
    """用 xlutils 复制原表并应用变更。成功返回 True（保格式）。

    xlutils 不可用、或该 .xls 读不出格式信息、或 xlutils 复制后内部数字格式损坏时，
    回退到普通 xlwt 重建，并返回 False —— 调用方应据此向用户告警。
    """
    rb = None
    try:
        from xlutils.copy import copy as xl_copy
        rb = xlrd.open_workbook(str(src), formatting_info=True)
        wb = xl_copy(rb)
    except Exception:
        _render_plain(src, plan, dst, new_sheets=new_sheets,
                      full_grid=full_grid)
        return False

    # 修复 xlutils.copy 偶尔带入的 None 自定义数字格式
    _sanitize_num_formats(wb)

    styles = _uniform_styles()
    created = _create_sheets(rb, wb, new_sheets, styles=styles)
    names = rb.sheet_names()

    if unify_format:
        # v1.6：整表统一格式 + 日期归一 + 写入变更（一次遍历完成）
        _apply_uniform_style(rb, wb, plan, styles, full_grid=full_grid)
    else:
        # 旧行为（--no-unify-format）：只写变更单元格，原有格式一律不动
        blue = styles["text_gen"]
        for name, changes in plan.items():
            if not changes:
                continue
            if name in created:
                ws = created[name]
            elif name in names:
                ws = wb.get_sheet(names.index(name))
            else:
                continue
            for (r, c), v in changes.items():
                ws.write(r, c, normalize_date_text(str(v)), blue)

    # 新建 sheet 不在 rb.sheet_names() 里，其变更单元格需单独写；
    # 数据区空格也补黑框，保持与既有表一致的连续网格
    for name, ws in created.items():
        changes = plan.get(name, {})
        tpl = rb.sheet_by_name(new_sheets[name]) if name in new_sheets else None
        for (r, c), v in changes.items():
            try:
                gen = _cell_style(styles, rb, tpl, r, c,
                                  text=True, generated=True) if tpl else styles["text_gen"]
                ws.write(r, c, normalize_date_text(str(v)), gen)
            except Exception:
                pass
        if unify_format and full_grid and tpl is not None:
            last = max([config.ROW_DATA_START - 1]
                       + [rr for (rr, _c) in changes])
            for r in range(config.ROW_DATA_START, last + 1):
                for c in range(tpl.ncols):
                    if (r, c) in changes:
                        continue
                    try:
                        ws.write(r, c, "",
                                 _cell_style(styles, rb, tpl, r, c, text=True))
                    except Exception:
                        pass

    if not unify_format:
        # 统一格式模式下所有非空单元格已用干净格式重写，无需再扫描修复
        _repair_cells_with_bad_format(rb, wb, styles["text"])

    try:
        wb.save(str(dst))
    except Exception:
        _render_plain(src, plan, dst, new_sheets=new_sheets,
                      full_grid=full_grid)
        return False
    return True


def _repair_cells_with_bad_format(rb, wb, style):
    """把数字格式为乱码（含 `(fmt`）的单元格统一重写成文本格式。

    历史版本在修复 xlutils None 格式时写入了 'General (fmt%d)' 这类占位格式，
    WPS/Excel 会把对应单元格显示成 `(fXt165)`。本函数扫描所有单元格，
    若其数字格式含 `(fmt`，则按原值以文本格式重写。
    """
    try:
        for sidx, name in enumerate(rb.sheet_names()):
            sh = rb.sheet_by_index(sidx)
            ws = wb.get_sheet(sidx)
            for r in range(sh.nrows):
                for c in range(sh.ncols):
                    xf_idx = sh.cell_xf_index(r, c)
                    xf = rb.xf_list[xf_idx]
                    fmt_idx = xf.format_key
                    fmt_str = ""
                    fmt_obj = rb.format_map.get(fmt_idx)
                    if fmt_obj is not None:
                        fmt_str = fmt_obj.format_str or ""
                    if "(fmt" not in fmt_str:
                        continue
                    val = sh.cell_value(r, c)
                    if sh.cell_type(r, c) == xlrd.XL_CELL_DATE and val not in (None, ""):
                        try:
                            val = xlrd_date_to_str(val)
                        except Exception:
                            val = str(val)
                    _write_value(ws, r, c, val, style)
    except Exception:
        # 修复是增量优化，失败不应影响主流程
        pass


def _write_value(ws, r: int, c: int, val, style) -> None:
    """兼容旧调用：写一个值（空值跳过，datetime 转字符串）。"""
    if val is None or val == "":
        return
    if isinstance(val, datetime.datetime):
        ws.write(r, c, val.strftime("%Y-%m-%d %H:%M:%S"), style)
    else:
        ws.write(r, c, val, style)


def _render_plain(src: pathlib.Path, plan: dict, dst: pathlib.Path,
                  new_sheets: dict | None = None,
                  full_grid: bool = True) -> None:
    """回退方案：xlwt 从头重建（尽量保留合并单元格；条件格式等仍会丢失）。"""
    try:
        rb = xlrd.open_workbook(str(src), formatting_info=True)
    except Exception:
        rb = xlrd.open_workbook(str(src), formatting_info=False)
    styles = _style_map()
    wb = xlwt.Workbook(encoding="utf-8")
    for name in rb.sheet_names():
        sh = rb.sheet_by_name(name)
        ws = wb.add_sheet(name[:31])
        _write_plain_sheet(ws, sh, plan.get(name, {}), styles, rb,
                           full_grid=full_grid)
    created = _create_sheets(rb, wb, new_sheets, styles=styles)
    for name, ws in created.items():
        changes = plan.get(name, {})
        tpl = rb.sheet_by_name(new_sheets[name]) if name in new_sheets else None
        for (r, c), v in changes.items():
            gen = (_cell_style(styles, rb, tpl, r, c, text=True, generated=True)
                   if tpl else styles["text_gen"])
            ws.write(r, c, normalize_date_text(str(v)), gen)
        if full_grid and tpl is not None:
            last = max([config.ROW_DATA_START - 1]
                       + [rr for (rr, _c) in changes])
            for r in range(config.ROW_DATA_START, last + 1):
                for c in range(tpl.ncols):
                    if (r, c) not in changes:
                        try:
                            ws.write(r, c, "",
                                     _cell_style(styles, rb, tpl, r, c, text=True))
                        except Exception:
                            pass
    wb.save(str(dst))


# ------------------------------------------------------------------ 另存副本

def write_back(track: TrackBook, actions: list[Action],
               out_path: pathlib.Path | None = None,
               new_sheets: dict | None = None,
               full_grid: bool = True) -> pathlib.Path:
    """另存为 .synced-{ts}.xls（原表不动）。返回新文件路径。

    用 xlwt 重建；合并单元格按原表复制，条件格式会丢失，仅供核对。
    """
    plan = _collect_plan(actions)
    if out_path is None:
        out_path = track.path.with_suffix(".synced-%s.xls" % _ts())
    out_path = pathlib.Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        rb = xlrd.open_workbook(str(track.path), formatting_info=True)
    except Exception:
        rb = xlrd.open_workbook(str(track.path), formatting_info=False)
    styles = _style_map()
    wbk = xlwt.Workbook(encoding="utf-8")
    for ts_ in track.sheets:
        name = ts_.name
        sh = rb.sheet_by_name(name) if name in rb.sheet_names() else ts_.sheet
        ws = wbk.add_sheet(name[:31])
        _write_plain_sheet(ws, sh, plan.get(name, {}), styles, rb,
                           full_grid=full_grid)
        for c in range(min(sh.ncols, 30)):
            ws.col(c).width = 14 * 256
    created = _create_sheets(rb, wbk, new_sheets, styles=styles)
    for name, ws in created.items():
        changes = plan.get(name, {})
        tpl = rb.sheet_by_name(new_sheets[name]) if name in new_sheets else None
        for (r, c), v in changes.items():
            gen = (_cell_style(styles, rb, tpl, r, c, text=True, generated=True)
                   if tpl else styles["text_gen"])
            ws.write(r, c, normalize_date_text(str(v)), gen)
        if full_grid and tpl is not None:
            last = max([config.ROW_DATA_START - 1]
                       + [rr for (rr, _c) in changes])
            for r in range(config.ROW_DATA_START, last + 1):
                for c in range(tpl.ncols):
                    if (r, c) not in changes:
                        try:
                            ws.write(r, c, "",
                                     _cell_style(styles, rb, tpl, r, c, text=True))
                        except Exception:
                            pass
        for c in range(30):
            ws.col(c).width = 14 * 256
    wbk.save(str(out_path))
    return out_path
