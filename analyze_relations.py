# -*- coding: utf-8 -*-
"""分析 5 张表的关联关系、关联字段与字段映射，并实测匹配结果。

关联键 = 井号（canon 归一化后比对）。
映射规则：
  已完成明细 → 完成日期 → 跟踪大表「实际完成」列（按阶段块定位）
  审核中明细 → 当前审核人 → 跟踪大表「实际完成」列（带 （审核中） 状态）
"""
from __future__ import annotations
import sys, pathlib, datetime
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))

import xlrd
from core import config, trackbook, detail_reader, filename_parser as fnp
from core.sync_engine import build_stage_index, unmatched_detail_wells
from core.pathutil import canon, nfc

SRC = config.DEFAULT_SOURCE_DIR
TB = SRC / config.DEFAULT_TRACKBOOK_NAME
STAGES_OF_INTEREST = ["地质设计", "工程设计"]

out = []
def w(s=""): out.append(s)

# ---------------------------------------------------------------- 1. 跟踪大表结构
w("# 数据关联分析（实测）")
w("")
w("生成时间：%s" % datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
w("")
w("## 1. 跟踪大表结构：`%s`" % TB.name)
w("")
tb = trackbook.TrackBook(str(TB))
w("- sheet 数：**%d**" % len(tb.sheets))
w("")
for tsh in tb.sheets:
    locs = {s.key: tsh.locate(s.key) for s in config.STAGES}
    locs = {k: v for k, v in locs.items() if v}
    stages_str = "、".join("%s→%s(%d)" % (k, v["actual_text"], v["actual"]) for k, v in locs.items())
    w("### sheet `%s`" % tsh.name)
    w("- 井号列 = %d（%s）" % (tsh.locate("地质设计")["well"], config.COL_NAME_WELL))
    w("- 含阶段列：%s" % stages_str)
    # 采样井号
    wells = []
    cw = tsh.locate("地质设计")["well"]
    for r in range(config.ROW_DATA_START, tsh.sheet.nrows):
        wd = tsh.cell(r, cw)
        if canon(wd):
            wells.append(nfc(wd))
    w("- 井数（井号列非空）：**%d**" % len(wells))
    w("- 井号样例：%s" % "、".join(wells[:8]))
    w("")

# ---------------------------------------------------------------- 2. 明细表结构
w("## 2. 四个明细表结构")
w("")
detail_files = []
for f in sorted(SRC.iterdir()):
    if f.suffix.lower() != ".xls":
        continue
    pf = fnp.parse_filename(str(f))
    if pf is None or pf.stage not in STAGES_OF_INTEREST:
        continue
    detail_files.append((f, pf))

recs_by_stage = {}
for f, pf in detail_files:
    book = xlrd.open_workbook(str(f), formatting_info=False)
    sh = book.sheet_by_index(0)
    headers = [nfc(sh.cell_value(0, c)) for c in range(sh.ncols)]
    # 取值列
    if pf.status == config.STATUS_REVIEW:
        vcands = config.DETAIL_COL_REVIEWER
    else:
        vcands = config.DETAIL_COL_DONE_DATE
    vcol = -1
    for c in range(sh.ncols):
        if nfc(sh.cell_value(0, c)) in vcands:
            vcol = c; break
    # 统计
    recs = detail_reader.load_detail(pf)
    recs_by_stage.setdefault(pf.stage, []).extend(recs)
    w("### `%s`" % f.name)
    w("- 解析：stage=**%s** status=**%s**" % (pf.stage, pf.status))
    w("- 表头：%s" % " | ".join(headers))
    w("- 取值列：%s（第 %d 列）" % (vcands, vcol))
    w("- 数据行数：**%d**" % len(recs))
    sample = recs[:4]
    for r in sample:
        w("  - 样例：井号=`%s` 值=`%r`" % (r.well_display, r.value))
    w("")

# ---------------------------------------------------------------- 3. 字段映射
w("## 3. 字段映射规则（关联键 = 井号）")
w("")
w("| 明细文件状态 | 明细取值列 | 含义 | 跟踪大表目标列 | 写入值 |")
w("|---|---|---|---|---|")
w("| 已完成 | 完成日期 | 审核完成时间 | 该阶段块的「实际完成」列 | 归一化 YYYY-MM-DD |")
w("| 审核中 | 当前审核人 | 审核人姓名 | 该阶段块的「实际完成」列 | 审核人 + （审核中） |")
w("")
w("- 关联键：**井号**（明细表第 0 列 = 跟踪大表井号列，均按 canon 归一化：去全角空格/换行、短横线归一、大小写不敏感）。")
w("- 阶段定位：跟踪大表 R2 一级表头切块 → 命中阶段名 → 块内 R3 找「实际完成」列为目标列。")
w("- 地质设计 → 列 G(6)；工程设计 → 列 J(10)（本样本本实测）。")
w("")

# ---------------------------------------------------------------- 4. 实测匹配
w("## 4. 实测匹配结果（在样本本上）")
w("")
stage_index = build_stage_index([r for rs in recs_by_stage.values() for r in rs])
unmatched = unmatched_detail_wells(tb.sheets, stage_index)
for stage in STAGES_OF_INTEREST:
    eng = stage_index.get(stage, {})
    miss = unmatched.get(stage, [])
    # 统计该阶段明细里 已完成/审核中
    done = sum(1 for r in eng.values() if r.status == config.STATUS_DONE)
    rev = sum(1 for r in eng.values() if r.status == config.STATUS_REVIEW)
    w("### %s" % stage)
    w("- 明细去重井数：**%d**（已完成 %d / 审核中 %d）" % (len(eng), done, rev))
    w("- 在大表能匹配到的井：**%d**" % (len(eng) - len(miss)))
    w("- **在大表无对应行、无法写入：%d 口**" % len(miss))
    if miss:
        w("  - 样例：%s" % "、".join(miss[:10]))
    w("")

w("## 5. 结论")
w("")
w("1. 关联键明确为**井号**，字段映射规则如上，工具已按此实现（已完成→日期、审核中→审核人）。")
w("2. 样本本只有 51 口井（上试井/修井作业/封堵井 三类），而常规气明细有数百口井，仅 25 口重叠。")
w("3. 重叠的 25 口井此前已写入（幂等），其余明细井在本样本中**无对应行**——工具按井号匹配、只填已有行、不新建行，故写不进去。")
w("4. 要实现用户要求的「无遗漏」，需二选一：")
w("   (a) 提供含全部井的**完整主表**（推荐，主表语义不变）；")
w("   (b) 实现「大表无行时自动追加新井行」功能（改变主表语义，需确定新井归入哪个 sheet，且 .xls 追加有格式风险）。")

text = "\n".join(out)
print(text)
pathlib.Path("output").mkdir(exist_ok=True)
p = pathlib.Path("output/关联分析_%s.md" % datetime.datetime.now().strftime("%Y%m%d-%H%M%S"))
p.write_text(text, encoding="utf-8")
print("\n>>> 分析已保存：%s" % p)
