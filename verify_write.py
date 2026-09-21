# -*- coding: utf-8 -*-
"""端到端验证：用真实生产代码证明「已完成→审核完成时间 / 审核中→审核人」写入正确。

分两段：
  0) 基线（只读真实样本大表）：当前真实状态 + 各阶段未匹配井数（"数据没写进去"真因）。
  1) 写入证明（样本大表副本）：把已匹配井的"实际完成"列清空，让工具重新写入，
     再读回核对每个井是否按 井号 精确写入了 日期 / 审核人（审核中）。
全程不修改原始样本大表（副本在 /tmp）。
"""
from __future__ import annotations

import pathlib
import shutil
import sys

sys.path.insert(0, "/Users/zoe/WorkBuddy/井下报表工具")

import xlrd
from xlutils.copy import copy as xlcopy

from core import config
from core.trackbook import TrackBook
from core.detail_reader import load_all
from core.filename_parser import scan_dir, is_trackbook
from core.sync_engine import build_stage_index, sync_sheet, unmatched_detail_wells, canon, format_value
from core.writer import write_inplace

SRC = pathlib.Path("/Users/zoe/Documents/重庆气矿项目/报表开发的支撑文件/井下作业报表生成")
BOOK = SRC / config.DEFAULT_TRACKBOOK_NAME
TMP = pathlib.Path("/tmp/demo_book")
TMP.mkdir(parents=True, exist_ok=True)
COPY = TMP / "井下作业设计节点跟踪大表.xls"

# ---- 加载明细 ----
parsed, rejected = scan_dir(SRC)
parsed = [p for p in parsed if not is_trackbook(p.path)]
records = load_all(parsed)
stage_index = build_stage_index(records)
print("[明细] %d 条记录，阶段=%s" % (len(records), ", ".join(stage_index.keys())))
for s, w2r in stage_index.items():
    done = sum(1 for r in w2r.values() if r.status == config.STATUS_DONE)
    rev = sum(1 for r in w2r.values() if r.status == config.STATUS_REVIEW)
    print("        %-10s 已完成 %3d / 审核中 %3d" % (s, done, rev))

# ============================================================= 0) 基线（真实样本，只读）
print("\n" + "=" * 64)
print("0) 基线：真实样本大表（只读）")
print("=" * 64)
tb_real = TrackBook(str(BOOK))
real_actions = []
for ts in tb_real.sheets:
    real_actions.extend(sync_sheet(ts, stage_index))
rw = sum(1 for a in real_actions if a.action == "write")
rs = sum(1 for a in real_actions if a.action == "skip")
rwarn = sum(1 for a in real_actions if a.action == "warn-dup-well")
print("匹配井动作：写入 %d / 跳过 %d / 告警 %d" % (rw, rs, rwarn))
unm = unmatched_detail_wells(tb_real.sheets, stage_index)
tot_unm = sum(len(v) for v in unm.values())
print("未匹配（明细有、大表无对应行）总计 %d 口井：" % tot_unm)
for s, wells in unm.items():
    print("        %-10s %d 口（如：%s）" % (s, len(wells), "、".join(wells[:8])))

# ============================================================= 1) 写入证明（副本，清空→重填）
print("\n" + "=" * 64)
print("1) 写入证明：副本大表先清空已匹配井的'实际完成'列，再让工具重新写入并读回核对")
print("=" * 64)
shutil.copy(BOOK, COPY)

# 找出副本中每个 (sheet, stage) 已匹配的井行，清空其 actual 列
rb = xlrd.open_workbook(str(COPY), formatting_info=True)
wb = xlcopy(rb)
sheet_idx = {name: i for i, name in enumerate(rb.sheet_names())}
tb0 = TrackBook(str(COPY))
blanked = []
for ts in tb0.sheets:
    for stage, well2rec in stage_index.items():
        cols = ts.locate(stage)
        if cols is None:
            continue
        cw, ca = cols["well"], cols["actual"]
        sh = ts.sheet
        for r in range(config.ROW_DATA_START, sh.nrows):
            w = canon(ts.cell(r, cw))
            if not w:
                continue
            if well2rec.get(w) is None:
                continue
            wb.get_sheet(sheet_idx[ts.name]).write(r, ca, "")
            blanked.append((ts.name, ts.cell(r, cw), stage))
wb.save(str(COPY))
print("已清空 %d 个已匹配井的'实际完成'单元格，准备验证写入。" % len(blanked))

# 真实写回
tb = TrackBook(str(COPY))
actions = []
for ts in tb.sheets:
    actions.extend(sync_sheet(ts, stage_index))
wr = sum(1 for a in actions if a.action == "write")
src_path, _backup, kept = write_inplace(tb, actions, backup_dir=TMP, make_backup=False)
print("工具写入 %d 个值（保留格式=%s）。" % (wr, kept))

# 读回核对
tb2 = TrackBook(str(COPY))
print("\n%-16s %-10s %-10s %-22s %s" % ("井号", "阶段", "状态", "写入值", "核对"))
ok = mismatch = 0
for ts in tb2.sheets:
    for stage, well2rec in stage_index.items():
        cols = ts.locate(stage)
        if cols is None:
            continue
        cw, ca = cols["well"], cols["actual"]
        sh = ts.sheet
        for r in range(config.ROW_DATA_START, sh.nrows):
            w = canon(ts.cell(r, cw))
            if not w:
                continue
            rec = well2rec.get(w)
            if rec is None:
                continue
            expect = format_value(rec.value, rec.status)
            got = ts.cell(r, ca)
            good = (canon(got) == canon(expect))
            ok += good
            mismatch += (not good)
            print("%-16s %-10s %-10s %-22s %s" % (
                ts.cell(r, cw), stage, rec.status, got, "✓" if good else "✗ 应为 " + expect))
print("\n结论：匹配井 %d 口，写入核对一致 %d / 不一致 %d" % (ok + mismatch, ok, mismatch))
