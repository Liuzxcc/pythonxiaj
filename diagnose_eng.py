# -*- coding: utf-8 -*-
"""工程数据写入诊断：扫描真实跟踪大表全部 sheet + 工程设计明细，
定位"为什么工程数据没有写入"。

用法（在 Bash 用 build venv 运行）：
    /Users/zoe/.workbuddy/binaries/python/envs/wellsync-build/bin/python diagnose_eng.py
"""
from __future__ import annotations

import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from core import config
from core import detail_reader
from core import trackbook
from core import filename_parser as fnp
from core.sync_engine import build_stage_index, decide, format_value, date_only
from core.pathutil import canon, nfc

STAGE = "工程设计"
SOURCE_DIR = config.DEFAULT_SOURCE_DIR
TRACKBOOK = SOURCE_DIR / config.DEFAULT_TRACKBOOK_NAME

print("=" * 70)
print("源目录 :", SOURCE_DIR)
print("跟踪大表:", TRACKBOOK, "存在?" , TRACKBOOK.exists())
print("=" * 70)

# ---- 1) 加载工程设计明细（已完成 + 审核中）----
detail_files = []
for f in sorted(SOURCE_DIR.iterdir()):
    if f.suffix.lower() != ".xls":
        continue
    pf = fnp.parse_filename(str(f))
    if pf is None:
        continue
    if pf.stage == STAGE:
        detail_files.append((f, pf))
        print("明细文件 :", f.name, "→ stage=%s status=%s" % (pf.stage, pf.status))

records = []
for f, pf in detail_files:
    recs = detail_reader.load_detail(pf)
    print("  加载 %s: %d 行" % (f.name, len(recs)))
    records.extend(recs)

idx = build_stage_index(records)
eng = idx.get(STAGE, {})
print("工程设计 去重后井数: %d" % len(eng))

# 展示几条明细原始 value（看是否带换行/时间）
print("\n--- 明细样本 value（前 5 条已完成）---")
shown = 0
for w, r in eng.items():
    if r.status == config.STATUS_DONE:
        print("  well=%-20s value=%r flow=%s" % (r.well_display, r.value, r.flow_count))
        shown += 1
        if shown >= 5:
            break

# ---- 2) 打开真实跟踪大表，扫描全部 sheet ----
print("\n" + "=" * 70)
print("扫描跟踪大表全部 sheet 的 工程设计 列")
print("=" * 70)

tb = trackbook.TrackBook(str(TRACKBOOK))
# 收集：book_wells[canon] = [(sheet, row1, value, col_actual)]
book_wells = {}
sheet_cols = {}
for tsh in tb.sheets:
    loc = tsh.locate(STAGE)
    if loc is None:
        continue
    sheet_cols[tsh.name] = (loc["well"], loc["actual"], loc["actual_text"])
    print("sheet %-30s 井号列=%d 工程设计实际列=%d (%s)"
          % (tsh.name, loc["well"], loc["actual"], loc["actual_text"]))
    sh = tsh.sheet
    for r in range(config.ROW_DATA_START, sh.nrows):
        wd = tsh.cell(r, loc["well"])
        w = canon(wd)
        if not w:
            continue
        val = tsh.cell(r, loc["actual"])
        book_wells.setdefault(w, []).append((tsh.name, r + 1, val, loc["actual"]))

print("工程设计 列中出现的井号总数(去重): %d" % len(book_wells))

# ---- 3) 交叉比对 ----
print("\n" + "=" * 70)
print("交叉比对：明细井 vs 大表井")
print("=" * 70)

matched = 0          # 明细井在大表里有行
not_in_book = []     # 明细井大表里完全没有
empty_but_detail = []  # 大表该行工程设计为空，但明细有值
would_write = []     # decide 返回写入
skipped = []         # decide 返回跳过
dirty_writes = []    # 写入值不同于明细原始（被归一化）

for w, rec in eng.items():
    if w in book_wells:
        matched += 1
        for (sname, row1, bval, _c) in book_wells[w]:
            new = decide(rec.value, rec.status, bval)
            if new is None:
                skipped.append((rec.well_display, sname, row1, bval, rec.value))
            else:
                would_write.append((rec.well_display, sname, row1, bval, new, rec.value))
                if new != rec.value:
                    dirty_writes.append((rec.well_display, sname, row1, bval, new, rec.value))
    else:
        not_in_book.append((rec.well_display, rec.status, rec.value))

# 反查：大表里工程设计为空的井，有多少能在明细中找到（应写但未写）
for w, rows in book_wells.items():
    rec = eng.get(w)
    if rec is None:
        continue
    for (sname, row1, bval, _c) in rows:
        if not nfc(bval):
            empty_but_detail.append((rec.well_display, sname, row1, rec.value))

print("明细井总数(去重)        : %d" % len(eng))
print("  在大表有对应行        : %d" % matched)
print("  大表里完全没有(新井)  : %d" % len(not_in_book))
print("  大表为空但明细有值    : %d" % len(empty_but_detail))
print("  → 经 decide 判定写入  : %d" % len(would_write))
print("  → 经 decide 判定跳过  : %d" % len(skipped))
print("  写入值被归一化(≠原始) : %d" % len(dirty_writes))

# ---- 4) 输出关键明细 ----
if would_write:
    print("\n--- 会写入的井（前 20）---")
    for (wd, s, r, old, new, raw) in would_write[:20]:
        print("  %-20s %s 行%d  旧=%r 新=%r (原=%r)" % (wd, s, r, old, new, raw))
if not_in_book:
    print("\n--- 大表中完全没有的明细井（前 30，这些无法写入，因为大表无对应行）---")
    for (wd, st, val) in not_in_book[:30]:
        print("  %-20s %s %r" % (wd, st, val))
    print("  ... 共 %d 口" % len(not_in_book))
if empty_but_detail:
    print("\n--- 大表为空但明细有值（前 20）---")
    for (wd, s, r, val) in empty_but_detail[:20]:
        print("  %-20s %s 行%d 明细值=%r" % (wd, s, r, val))

# ---- 5) 大表工程设计列空单元格总数 ----
print("\n" + "=" * 70)
print("大表 工程设计 列空单元格统计")
print("=" * 70)
total_eng_cells = 0
empty_eng_cells = 0
for w, rows in book_wells.items():
    for (_s, _r, bval, _c) in rows:
        total_eng_cells += 1
        if not nfc(bval):
            empty_eng_cells += 1
print("工程设计列有井号的数据行总数: %d" % total_eng_cells)
print("其中为空(未填)的单元格      : %d" % empty_eng_cells)
print("这些空单元格里能由明细补的  : %d" % len(empty_but_detail))
print("这些空单元格里明细也没有的  : %d" % (empty_eng_cells - len(empty_but_detail)))
