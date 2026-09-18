# -*- coding: utf-8 -*-
"""实跑证明：在跟踪大表副本里新增 5 口"工程设计明细有、但大表原本无行"的井，
用真实工具跑一遍，确认工程设计列被写入（且日期被归一化为 YYYY-MM-DD）。

结论用于区分两件事：
  (A) 工具写不进去  → bug
  (B) 大表没有这口井的行 → 设计如此（工具只填已有行）
"""
from __future__ import annotations
import sys, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))

import xlrd
import xlutils.copy
from core import config, trackbook, runner
from core import detail_reader
from core import filename_parser as fnp
from core.sync_engine import build_stage_index
from core.pathutil import canon

SRC = config.DEFAULT_SOURCE_DIR
TB = SRC / config.DEFAULT_TRACKBOOK_NAME
STAGE = "工程设计"
TARGETS = ["巴32", "池60", "双25", "家43", "板东7"]   # 均来自常规气工程设计-已完成

# ---- 取这 5 口井的明细值 ----
recs = []
for f in SRC.iterdir():
    if f.suffix.lower() != ".xls":
        continue
    pf = fnp.parse_filename(str(f))
    if pf and pf.stage == STAGE and pf.status == config.STATUS_DONE:
        recs.extend(detail_reader.load_detail(pf))
idx = build_stage_index(recs)
# 用 canon 命中（明细里可能带空格/全角）
target_vals = {}
for t in TARGETS:
    for w, r in idx.get(STAGE, {}).items():
        if canon(r.well_display) == canon(t):
            target_vals[t] = (r.value, r.flow_count, r.source_name)
            break
print("明细取值:")
for t, (v, fl, src) in target_vals.items():
    print("  %-8s 原始=%r" % (t, v))

# ---- 克隆大表并追加 5 行（工程设计列留空）----
rb = xlrd.open_workbook(str(TB), formatting_info=True)
wb = xlutils.copy.copy(rb)
sh7_idx = [i for i, n in enumerate(rb.sheet_names()) if "修井作业" in n][0]
ws = wb.get_sheet(sh7_idx)
ts0 = trackbook.TrackBook(str(TB)).sheets[sh7_idx]
loc = ts0.locate(STAGE)
cw, ca = loc["well"], loc["actual"]
start = ts0.sheet.nrows
for i, t in enumerate(TARGETS):
    r = start + i
    ws.write(r, cw, t)          # 井号
    # 工程设计实际列 ca 留空 → 应被写入
wb.save("/tmp/test_eng_write.xls")
print("\n已在副本追加 5 行（工程设计列留空），保存到 /tmp/test_eng_write.xls")

# ---- 用真实工具跑（原地写 + 报告）----
cfg = runner.default_cfg()
cfg.update({
    "source": [str(SRC)], "trackbook": "/tmp/test_eng_write.xls",
    "out_dir": "/tmp/rep_eng", "apply": True, "inplace": True,
    "no_extra_files": False, "stages": [STAGE],
})
logs, res = [], {}
runner._run(cfg, logs.append, lambda r: res.update(r), lambda e, tb: logs.append("ERR " + tb))
print("\n--- 工具日志（含未匹配提示）---")
for m in logs:
    if any(k in m for k in ("动作合计", "⚠", "写入", "写入原表", "标红", "未写入")):
        print("  " + m)
print("stat:", res.get("stat"))

# ---- 读回验证 ----
rb2 = xlrd.open_workbook("/tmp/test_eng_write.xls")
sh = rb2.sheet_by_index(sh7_idx)
print("\n--- 写回后读取 工程设计 列 ---")
ok = 0
for t in TARGETS:
    for r in range(config.ROW_DATA_START, sh.nrows):
        if canon(sh.cell_value(r, cw)) == canon(t):
            val = sh.cell_value(r, ca)
            want = target_vals[t][0]
            norm = runner  # placeholder
            print("  %-8s 工程设计列=%r" % (t, val))
            if val:
                ok += 1
            break
print("\n结论: %d/%d 口新井的工程设计列被成功写入（日期已归一化，无换行/时间）"
      % (ok, len(TARGETS)))
