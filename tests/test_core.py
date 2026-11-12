# -*- coding: utf-8 -*-
"""测试脚本：生成测试数据 + 验证核心逻辑。

覆盖了方案 §4.3 / §13.4 的决策矩阵、文件名解析、列自适应定位、
sheet 7 列改造，以及一次完整端到端（造 .xls → 扫描 → 同步 → 出 diff）。

运行: python tests/test_core.py
"""

from __future__ import annotations

import datetime
import os
import pathlib
import shutil
import sys
import tempfile

# 添加项目根目录到路径（对齐参考项目 tests/test_core.py）
_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, _PROJECT_ROOT)

import xlwt  # noqa: E402

from core import config  # noqa: E402
from core import restructure, reports, sync_engine  # noqa: E402
from core.detail_reader import load_all  # noqa: E402
from core.filename_parser import _compile, classify, is_trackbook, parse_filename, scan_dir  # noqa: E402
from core.pathutil import canon, nfc, safe_filename, safe_path  # noqa: E402
from core.sync_engine import (  # noqa: E402
    Action, DetailRecord, build_stage_index, date_only, decide,
    format_value, is_dirty_date, pick_latest, sync_sheet,
)
from core.trackbook import TrackBook  # noqa: E402
from core.writer import write_inplace, write_back  # noqa: E402

# ------------------------------------------------------------------ 断言辅助

_RESULTS: list[tuple[bool, str, str]] = []
_SECTION = ""


def section(title: str) -> None:
    global _SECTION
    _SECTION = title
    print("\n" + "=" * 64)
    print("  " + title)
    print("=" * 64)


def check(desc: str, cond, detail="") -> bool:
    ok = bool(cond)
    if not _SECTION:
        section("未分组")
    _RESULTS.append((ok, "[%s] %s" % (_SECTION, desc), str(detail)))
    tail = ("  -> %s" % detail) if (not ok and detail != "") else ""
    print("  [%s] %s%s" % ("PASS" if ok else "FAIL", desc, tail))
    return ok


def eq(desc: str, got, want) -> bool:
    return check(desc, got == want, "期望 %r，实际 %r" % (want, got))


# ================================================================== A. 归一化

def test_pathutil() -> None:
    section("A. 跨平台归一化 pathutil")

    eq("NFC 去空白", canon("  卧122  "), "卧122")
    eq("全角破折号 U+FF0D 归一", canon("黄202H1－2"), canon("黄202H1-2"))
    eq("en dash U+2013 归一", canon("黄202H1–2"), canon("黄202H1-2"))
    eq("全角空格 U+3000 剔除", canon("卧　122"), "卧122")
    eq("NBSP 剔除", canon("卧\xa0122"), "卧122")
    eq("大小写无关", canon("Huang202H1-2"), canon("huang202h1-2"))
    eq("nfc(None) 安全", nfc(None), "")
    eq("井号含破折号可与普通写法互配",
       canon("黄202H1－2") == canon("黄202H1-2"), True)

    check("safe_filename 拦 Windows 非法字符 '?'",
          _raises(ValueError, safe_filename, "卧122?.xls"))
    check("safe_path 对不存在路径抛 FileNotFoundError",
          _raises(FileNotFoundError, safe_path, "/绝/不/存/在/的/路径.xls"))


def _raises(exc, fn, *a, **kw) -> bool:
    try:
        fn(*a, **kw)
    except exc:
        return True
    except Exception:
        return False
    return False


# ============================================================== B. 文件名解析

def test_filename() -> None:
    section("B. 文件名解析 filename_parser")

    pf = parse_filename("常规气地质设计-已完成.xls")
    eq("地质设计-已完成 → stage", pf.stage, "地质设计")
    eq("地质设计-已完成 → status", pf.status, config.STATUS_DONE)

    pf = parse_filename("常规气地质设计-审核中.xls")
    eq("地质设计-审核中 → status", pf.status, config.STATUS_REVIEW)

    pf = parse_filename("常规气工程方案审查-已完成.xls")
    eq("工程方案审查不被截断", pf.stage, "工程方案审查")

    pf = parse_filename("常规气工艺设计-已完成.XLS")
    eq("扩展名大小写不敏感", (pf.stage, pf.ext), ("工艺设计", ".xls"))

    check("长关键字优先（防御性：短词在前也命中长词）",
          _compile(["工程", "工程方案审查"]).search("工程方案审查").group(0) == "工程方案审查")

    eq("无扩展名 → None", parse_filename("费用清单"), None)

    _, reason = classify("常规气地质设计-已完成?.xls")
    check("非法文件名被拦下", reason.startswith("非法文件名"), reason)

    _, reason = classify("费用清单.txt")
    eq("扩展名不支持", reason, "扩展名不支持")

    _, reason = classify("费用清单.xlsx")
    check("关键字不匹配（未识别阶段）", reason.startswith("关键字不匹配"), reason)

    _, reason = classify("常规气地质设计-已驳回.xls")
    check("状态存疑落入人工确认桶", reason.startswith("状态存疑"), reason)

    eq("is_trackbook 命中", is_trackbook("井下作业设计节点跟踪大表.xls"), True)
    eq("is_trackbook 未命中", is_trackbook("常规气地质设计-已完成.xls"), False)


def test_scan_dir(tmp: pathlib.Path) -> None:
    """scan_dir 应跳过 macOS 隐藏文件与 Excel 锁文件。"""
    d = tmp / "scan"
    d.mkdir(parents=True, exist_ok=True)
    (d / "常规气地质设计-已完成.xls").write_bytes(b"")
    (d / "常规气地质设计-审核中.xls").write_bytes(b"")
    (d / "常规气工艺设计-已完成.xls").write_bytes(b"")
    (d / "井下作业设计节点跟踪大表.xls").write_bytes(b"")
    (d / "~$常规气地质设计-已完成.xls").write_bytes(b"")   # Excel 锁文件
    (d / ".DS_Store").write_bytes(b"")                      # macOS 隐藏文件
    (d / "子目录").mkdir(exist_ok=True)

    parsed, rejected = scan_dir(d)
    names = {p.name for p in parsed}
    eq("命中 3 个明细文件（集合比较，不受排序影响）", names,
       {"常规气地质设计-已完成.xls", "常规气地质设计-审核中.xls", "常规气工艺设计-已完成.xls"})
    check("锁文件 / 隐藏文件 / 子目录均未计入命中",
          all("~$" not in n and ".DS" not in n for n in names))
    check("跟踪大表走 rejected 桶（无阶段关键字）",
          any(reason.startswith("关键字不匹配") for reason, _ in rejected))
    check("锁文件被跳过、未进入 rejected",
          not any("~$" in name for _, name in rejected))


# ============================================================ C. 日期与格式化

def test_dates() -> None:
    section("C. 日期归一化与写入值格式化")

    eq("带时间的日期截断", date_only("2026-08-14 10:18:20"), "2026-08-14")
    eq("斜杠日期补零", date_only("2026/8/5"), "2026-08-05")
    eq("标准日期原样", date_only("2026-08-14"), "2026-08-14")
    eq("非日期原样返回", date_only("张三"), "张三")
    eq("空串原样", date_only(""), "")

    eq("空判脏", is_dirty_date(""), True)
    eq("'/' 判脏", is_dirty_date("/"), True)
    eq("合法日期不脏", is_dirty_date("2026-08-14"), False)
    eq("人名判脏（非日期）", is_dirty_date("张三（审核中）"), True)
    eq("越界日期判脏（2025-03-38 不存在）", is_dirty_date("2025/3/38"), True)
    eq("非闰年 2 月 29 日判脏", is_dirty_date("2026-02-29"), True)
    eq("闰年 2 月 29 日合法", is_dirty_date("2024-02-29"), False)
    eq("越界月份判脏", is_dirty_date("2026-13-01"), True)

    eq("已完成 → 纯日期", format_value("2026-08-14 10:18:20", config.STATUS_DONE), "2026-08-14")
    eq("审核中 → 人名 + 后缀", format_value("张三", config.STATUS_REVIEW), "张三（审核中）")
    eq("审核中且无人名 → 空", format_value("", config.STATUS_REVIEW), "")
    eq("已完成且无日期 → 空", format_value("", config.STATUS_DONE), "")


# ============================================================== D. 决策矩阵

def test_decide() -> None:
    """方案 §4.3 / §13.4 的 8 条规则 + 兜底。"""
    section("D. 冲突决策矩阵 decide()")

    D, R = config.STATUS_DONE, config.STATUS_REVIEW

    # ① 原值为空 / "/"
    eq("① 原值为空 → 写入", decide("2026-03-01", D, ""), "2026-03-01")
    eq("① 原值为 '/' → 写入", decide("2026-03-01", D, "/"), "2026-03-01")

    # ② 审核中 → 已完成
    eq("② 审核中转已完成 → 写入",
       decide("2026-03-01", D, "张三（审核中）"), "2026-03-01")

    # ③ 已完成 → 审核中（回退）
    eq("③ 已完成被审核中覆盖 → 默认拒绝",
       decide("李四", R, "2026-03-01"), None)
    eq("③ 开 allow_recheck_overwrite → 允许",
       decide("李四", R, "2026-03-01", True), "李四（审核中）")

    # ④ 幂等
    eq("④ 同日期 → 跳过",
       decide("2026-03-01", D, "2026-03-01"), None)
    eq("④ 同日期带时间 → 跳过",
       decide("2026-03-01 09:00:00", D, "2026-03-01"), None)
    eq("④ 同审核人 → 跳过",
       decide("张三", R, "张三（审核中）"), None)

    # ⑤ 两个日期 → 取较新
    eq("⑤ 新日期更晚 → 写入",
       decide("2026-05-01", D, "2026-03-01"), "2026-05-01")
    eq("⑤ 新日期更早 → 保留原值",
       decide("2025-12-20", D, "2026-01-05"), None)

    # ⑧ 明细值为空 → 绝不清空
    eq("⑧ 已完成空日期 → 跳过", decide("", D, "2026-03-01"), None)
    eq("⑧ 审核中空人名 → 跳过", decide("", R, "2026-03-01"), None)

    # 兜底：原值脏数据 → 覆盖
    eq("兜底 越界日期被覆盖",
       decide("2026-03-01", D, "2025/3/38"), "2026-03-01")
    eq("兜底 脏文本被覆盖",
       decide("2026-03-01", D, "待定"), "2026-03-01")

    # ⑨ 占位符"当前审核人员"不是真实完成值，应被覆盖；同审核人仍幂等
    eq("⑨ 当前审核人员占位符 → 被审核人覆盖",
       decide("李四", R, "当前审核人员"), "李四（审核中）")
    eq("⑨ 当前审核人员占位符 → 被完成日期覆盖",
       decide("2026-03-01", D, "当前审核人员"), "2026-03-01")
    eq("⑨ 同审核人占位符也幂等",
       decide("李四", R, "李四（审核中）"), None)


# ============================================================ E. 明细聚合

def test_aggregate() -> None:
    section("E. 明细聚合 pick_latest / build_stage_index")

    D = config.STATUS_DONE
    rs = [
        DetailRecord("卧122", "卧122", "地质设计", D, "2026-03-01", 3, "a.xls"),
        DetailRecord("卧122", "卧122", "地质设计", D, "2026-03-10", 7, "a.xls"),
        DetailRecord("峰2", "峰2", "地质设计", D, "2025-12-20", 1, "a.xls"),
        DetailRecord("黄202H1-2", "黄202H1-2", "工艺设计", D, "2026-04-05", 1, "b.xls"),
    ]
    eq("⑥ 取流程次数最大者", pick_latest(rs[:2]).value, "2026-03-10")

    idx = build_stage_index(rs)
    eq("按阶段分组数量", len(idx), 2)
    eq("同井号去重后剩 2 条", len(idx["地质设计"]), 2)
    eq("工艺设计独立成组", list(idx["工艺设计"]), ["黄202H1-2"])


# ============================================================ F. sync_sheet

class _FakeSheet:
    """最小 TrackSheet 替身，用于纯逻辑回归测试。

    必须提供 .sheet.nrows / .name / .cell(r,c) / .locate(stage)。
    """

    def __init__(self, name, rows, locate_map):
        self.name = name
        self.rows = rows
        self._map = locate_map
        self.sheet = self

    @property
    def nrows(self):
        return len(self.rows)

    def cell(self, r, c):
        if r < 0 or r >= len(self.rows):
            return ""
        row = self.rows[r]
        return "" if c < 0 or c >= len(row) else ("" if row[c] is None else str(row[c]))

    def locate(self, stage):
        return self._map.get(stage)


def test_sync_sheet() -> None:
    section("F. sync_sheet 回归（v1.1 关键修复）")

    D = config.STATUS_DONE
    # 行 4/5/6：卧122 / 峰2 / 卧122（重复井号）
    rows = [
        ["", "", "", ""],                                   # r0
        ["", "", "", ""],                                   # r1
        ["井号", "地质设计", "", "工艺设计"],                 # r2 一级表头
        ["", "计划完成", "实际完成", "实际完成"],             # r3 二级表头
        ["卧122", "", "", ""],                              # r4
        ["峰2", "", "2026-01-05", ""],                      # r5
        ["卧122", "", "", ""],                              # r6 重复井号
    ]
    # 地质设计 → 实际完成列 2；工艺设计 → 实际完成列 3
    locate_map = {
        "地质设计": {"well": 0, "actual": 2, "plan": 1, "legacy_used": False},
        "工艺设计": {"well": 0, "actual": 3, "plan": -1, "legacy_used": False},
    }
    sh = _FakeSheet("7.修井作业", rows, locate_map)

    idx = build_stage_index([
        DetailRecord("卧122", "卧122", "地质设计", D, "2026-03-10", 7, "a.xls"),
        DetailRecord("峰2", "峰2", "地质设计", D, "2025-12-20", 1, "a.xls"),
        DetailRecord("卧122", "卧122", "工艺设计", D, "2026-04-02", 1, "b.xls"),
    ])
    acts = sync_sheet(sh, idx)

    def _pick(stage, row):
        return next((a for a in acts if a.stage == stage and a.row == row), None)

    # 地质设计走满 3 行；工艺设计只有卧122 有记录，峰2 无记录不产生动作，
    # 卧122 的重复行在工艺设计阶段同样告警（去重按阶段而非按 sheet）
    eq("动作总数（地质 3 + 工艺 2）", len(acts), 5)

    a = _pick("地质设计", 5)   # 1-based 行号 → rows[4]
    eq("卧122 首行写入", (a.action, a.new_value), ("write", "2026-03-10"))

    a = _pick("地质设计", 6)   # rows[5] 峰2，原值更晚
    eq("峰2 原值更晚 → 跳过", (a.action, a.reason), ("skip", "原值日期更新，保留"))

    a = _pick("地质设计", 7)   # rows[6] 卧122 重复
    eq("同阶段重复井号 → 告警", a.action, "warn-dup-well")

    a = _pick("工艺设计", 5)
    eq("★ 同井号跨阶段仍写入（v1.1 修复点）",
       (a.action, a.new_value), ("write", "2026-04-02"))

    a = _pick("工艺设计", 7)
    eq("工艺设计阶段的重复井号同样告警", a.action, "warn-dup-well")

    a = _pick("工艺设计", 6)
    eq("峰2 无工艺设计明细 → 无动作", a, None)

    # sheet 无该阶段 → locate 返回 None → 不产生任何动作
    sh2 = _FakeSheet("6.上试井", rows, {"地质设计": locate_map["地质设计"]})
    acts2 = sync_sheet(sh2, idx)
    stages = {a.stage for a in acts2}
    check("sheet 无工艺设计列时静默跳过", "工艺设计" not in stages, stages)

    # 幂等：原值已等于新值 → skip
    rows3 = [r[:] for r in rows]
    rows3[4][2] = "2026-03-10"
    sh3 = _FakeSheet("7.修井作业", rows3, locate_map)
    a3 = next(a for a in sync_sheet(sh3, idx) if a.stage == "地质设计" and a.row == 5)
    eq("重复运行幂等", (a3.action, a3.reason), ("skip", "日期相同，幂等跳过"))


# ========================================================== G. 列结构调整

def _sheet7_snapshot(ncols=18):
    """构造 sheet 7 旧结构快照：工艺设计单列、二级表头'当前进度'。"""
    rows = [[None] * ncols for _ in range(8)]
    rows[0][0] = "井下作业设计节点跟踪表"
    rows[config.ROW_STAGE_HEADER] = [
        "井号", "地质设计", "", "工程设计", "", "工程方案审查", "",
        "工艺设计", "修前工程\n及概算完成情况", "", "", "",
        "修井作业开工", "", "修井作业完工", "", "", "",
    ]
    rows[config.ROW_SUB_HEADER] = [
        "", "计划完成", "实际完成", "计划完成", "实际完成", "计划完成", "实际完成",
        "当前进度", "计划完成", "实际完成", "计划完成", "实际完成",
        "计划完成", "实际完成", "计划完成", "实际完成", "", "",
    ]
    rows[4][0] = "卧122"
    rows[4][8] = "2026-01-01"     # 旧 M 列（修前工程-计划完成）
    rows[4][9] = "2026-01-02"     # 旧 N 列（修前工程-实际完成）
    rows[5][0] = "峰2"
    rows[5][8] = "2026-02-01"
    return restructure.SheetSnapshot(name="7.修井作业", rows=rows, ncols_old=ncols)


def test_restructure() -> None:
    section("G. sheet 7 列结构调整 restructure")

    before = _sheet7_snapshot()
    plan = restructure.plan_columns(before)
    eq("插入位置 = 工艺设计列 + 1", plan.insert_at, 8)
    eq("实际完成列保持不变", plan.actual_col, 7)
    eq("计划完成列 = 新增列", plan.plan_col, 8)
    eq("旧列 8 → 新列 9", plan.old_to_new[8], 9)
    eq("旧列 7 → 新列 7", plan.old_to_new[7], 7)

    after = restructure.apply_plan(before, plan)
    ncols_after = max(len(r) for r in after.rows)
    eq("列数 18 → 19", ncols_after, 19)

    r2 = after.rows[config.ROW_STAGE_HEADER]
    r3 = after.rows[config.ROW_SUB_HEADER]
    eq("新 L 一级表头", r2[7], "工艺设计")
    eq("新 M 一级表头", r2[8], "工艺设计")
    eq("新 L 二级表头", r3[7], "实际完成")
    eq("新 M 二级表头", r3[8], "计划完成")
    eq("新 N 一级表头右移到位", r2[9], "修前工程\n及概算完成情况")
    eq("新 N 二级表头右移到位", r3[9], "计划完成")
    eq("新 O 二级表头右移到位", r3[10], "实际完成")

    eq("数据行右移：旧 M(8) → 新 N(9)", after.rows[4][9], "2026-01-01")
    eq("数据行右移：旧 N(9) → 新 O(10)", after.rows[4][10], "2026-01-02")
    eq("工艺设计新列留空", after.rows[4][7], None)
    eq("工艺设计新列留空（计划）", after.rows[4][8], None)

    # 幂等：对已改造的表再跑一次应判定无需变更
    plan2 = restructure.plan_columns(after)
    eq("★ 重复改造幂等（insert_at=None）", plan2.insert_at, None)

    eq("find_sheet7 定位",
       restructure.find_sheet7(["6.上试井", "7.修井作业-陈亚妮", "8.封堵井"]),
       "7.修井作业-陈亚妮")

    eq("col_letter(0)", restructure.col_letter(0), "A")
    eq("col_letter(25)", restructure.col_letter(25), "Z")
    eq("col_letter(26)", restructure.col_letter(26), "AA")


def test_restructure_missing_stage() -> None:
    """sheet 6/8 没有工艺设计块 → plan_columns 应明确报错而不是乱改。"""
    snap = _sheet7_snapshot()
    snap.rows[config.ROW_STAGE_HEADER][7] = "修前工程"
    check("未找到目标阶段时抛 ValueError",
          _raises(ValueError, restructure.plan_columns, snap))


# ================================================================ H. 报告

def test_reports(tmp: pathlib.Path) -> None:
    section("H. 报告 reports")

    eq("_col_letter(0)", reports._col_letter(0), "A")
    eq("_col_letter(26)", reports._col_letter(26), "AA")
    eq("_col_letter(27)", reports._col_letter(27), "AB")

    acts = [
        Action("卧122", "卧122", "7.修井作业", 5, "地质设计", 2, "", "2026-03-10", "write", "原值为空，写入"),
        Action("峰2", "峰2", "6.上试井", 6, "地质设计", 2, "2026-01-05", "", "skip", "原值日期更新，保留"),
        Action("卧122", "卧122", "6.上试井", 7, "地质设计", 2, "", "", "warn-dup-well", "重复井号"),
    ]
    s = reports.summarize(acts)
    eq("汇总统计", (s["total"], s["write"], s["skip"], s["warn"]), (3, 1, 1, 1))

    dst = tmp / "reports" / "diff-测试.csv"
    reports.write_diff(acts, dst)
    check("diff CSV 已生成", dst.exists())
    raw = dst.read_bytes()
    check("CSV 含 UTF-8 BOM（Excel 中文不乱码）", raw.startswith(b"\xef\xbb\xbf"))
    text = raw.decode("utf-8-sig")
    check("CSV 表头完整", reports.DIFF_HEADER[0] in text and "C(2)" in text)


# ============================================================== I. 端到端

def _write_detail(path: pathlib.Path, headers: list[str], rows: list[list]) -> None:
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("Sheet1")
    for c, h in enumerate(headers):
        ws.write(0, c, h)
    date_xf = xlwt.easyxf(num_format_str="YYYY-MM-DD")
    dt_xf = xlwt.easyxf(num_format_str="YYYY-MM-DD HH:MM:SS")
    for r, row in enumerate(rows, start=1):
        for c, v in enumerate(row):
            if isinstance(v, datetime.datetime):
                ws.write(r, c, v, dt_xf if (v.hour or v.minute) else date_xf)
            else:
                ws.write(r, c, v)
    wb.save(str(path))


def _write_trackbook(path: pathlib.Path) -> None:
    """造一份两 sheet 的跟踪大表：sheet 6 无工艺设计，sheet 7 有。"""
    wb = xlwt.Workbook(encoding="utf-8")
    head_xf = xlwt.easyxf("font: bold on; align: wrap on, vert center, horiz center;")

    # ---------- 6.上试井（9 列，无工艺设计）----------
    ws = wb.add_sheet("6.上试井")
    h1 = ["井号", "地质设计", "", "工程设计", "", "工程方案审查", "", "修前工程及概算", ""]
    h2 = ["", "计划完成", "实际完成", "计划完成", "实际完成",
          "计划完成", "实际完成", "计划完成", "实际完成"]
    for c, v in enumerate(h1):
        ws.write(config.ROW_STAGE_HEADER, c, v, head_xf)
    for c, v in enumerate(h2):
        ws.write(config.ROW_SUB_HEADER, c, v, head_xf)
    data6 = [
        ["卧122", "", "", "", "", "", "", "", ""],                 # r4
        ["峰2", "", "2026-01-05", "", "", "", "", "", ""],         # r5 原值更晚
        ["卧122", "", "", "", "", "", "", "", ""],                 # r6 重复井号
    ]
    for r, row in enumerate(data6, start=config.ROW_DATA_START):
        for c, v in enumerate(row):
            if v:
                ws.write(r, c, v)

    # ---------- 7.修井作业（11 列，含工艺设计）----------
    ws = wb.add_sheet("7.修井作业")
    h1 = ["井号", "地质设计", "", "工程设计", "", "工程方案审查", "",
          "工艺设计", "", "修前工程及概算", ""]
    h2 = ["", "计划完成", "实际完成", "计划完成", "实际完成", "计划完成", "实际完成",
          "计划完成", "实际完成", "计划完成", "实际完成"]
    for c, v in enumerate(h1):
        ws.write(config.ROW_STAGE_HEADER, c, v, head_xf)
    for c, v in enumerate(h2):
        ws.write(config.ROW_SUB_HEADER, c, v, head_xf)
    data7 = [
        ["卧122", "", "", "", "", "", "", "", "", "", ""],         # r4
        ["黄202H1-2", "", "", "", "", "", "", "", "", "", ""],     # r5
    ]
    for r, row in enumerate(data7, start=config.ROW_DATA_START):
        for c, v in enumerate(row):
            if v:
                ws.write(r, c, v)

    wb.save(str(path))


def test_end_to_end(tmp: pathlib.Path) -> None:
    section("I. 端到端（造 .xls → 扫描 → 同步 → diff）")

    src = tmp / "e2e"
    src.mkdir(parents=True, exist_ok=True)

    # ---- 明细表 ----
    done_headers = ["井号", "单位", "措施类别", "资金来源", "设计单位",
                    "设计人", "流程次数", "设计日期", "完成日期", "备注"]
    _write_detail(src / "常规气地质设计-已完成.xls", done_headers, [
        ["卧122", "重庆气矿", "压裂", "上市", "院", "张三", 3, "", "2026-03-01", ""],
        ["峰2", "重庆气矿", "酸化", "未上市", "院", "李四", 1, "",
         datetime.datetime(2025, 12, 20), ""],     # 真实日期单元格
        ["卧122", "重庆气矿", "压裂", "上市", "院", "张三", 7, "",
         datetime.datetime(2026, 3, 10, 10, 18, 20), ""],   # 带时间 + 流程次数更大
    ])

    review_headers = ["井号", "单位", "措施类型", "资金来源", "当前审核人",
                      "设计单位", "设计人", "流程次数", "设计日期", "上报日期", "备注"]
    _write_detail(src / "常规气地质设计-审核中.xls", review_headers, [
        ["黄202H1-2", "重庆气矿", "压裂", "上市", "张三", "院", "王五", 2, "", "2026-03-05", ""],
    ])

    _write_detail(src / "常规气工艺设计-已完成.xls", done_headers, [
        ["卧122", "重庆气矿", "压裂", "上市", "院", "赵六", 1, "", "2026-04-02", ""],
        ["黄202H1-2", "重庆气矿", "压裂", "上市", "院", "赵六", 1, "", "2026-04-05", ""],
    ])

    # ---- 跟踪大表 ----
    _write_trackbook(src / "井下作业设计节点跟踪大表.xls")

    # ---- 1. 扫描（排除跟踪大表）----
    parsed, rejected = scan_dir(src)
    parsed = [p for p in parsed if not is_trackbook(p.path)]
    eq("解析出 3 份明细表", len(parsed), 3)

    # ---- 2. 加载明细 ----
    records = load_all(parsed)
    eq("明细记录 6 条（卧122 在地质设计有 2 条待去重）", len(records), 6)

    idx = build_stage_index(records)
    eq("阶段分组", sorted(idx), ["地质设计", "工艺设计"])
    eq("卧122 取流程次数 7 的日期", idx["地质设计"]["卧122"].value, "2026-03-10 10:18:20")

    # ---- 3. 同步 ----
    track = TrackBook(src / "井下作业设计节点跟踪大表.xls")
    actions: list[Action] = []
    for ts in track.sheets:
        actions.extend(sync_sheet(ts, idx))

    s = reports.summarize(actions)
    eq("动作总数", s["total"], 7)
    eq("写入数", s["write"], 5)
    eq("跳过数", s["skip"], 1)
    eq("告警数（重复井号）", s["warn"], 1)

    def _find(sheet, row, stage):
        return next((a for a in actions
                     if a.sheet == sheet and a.row == row and a.stage == stage), None)

    a = _find("6.上试井", 5, "地质设计")
    eq("6.上试井 卧122 写入（带时间已截断）", (a.action, a.new_value), ("write", "2026-03-10"))
    eq("写入列 = C 列（地质设计-实际完成）", a.target_col, 2)

    a = _find("6.上试井", 6, "地质设计")
    eq("6.上试井 峰2 原值更晚被保留", a.action, "skip")

    a = _find("6.上试井", 7, "地质设计")
    eq("6.上试井 重复井号告警", a.action, "warn-dup-well")

    a = _find("6.上试井", 5, "工艺设计")
    eq("6.上试井 无工艺设计列 → 不产生动作", a, None)

    a = _find("7.修井作业", 5, "地质设计")
    eq("7.修井作业 卧122 地质设计写入", (a.action, a.new_value), ("write", "2026-03-10"))

    a = _find("7.修井作业", 6, "地质设计")
    eq("7.修井作业 黄202H1-2 写入审核人", (a.action, a.new_value), ("write", "张三（审核中）"))

    a = _find("7.修井作业", 5, "工艺设计")
    eq("★ 7.修井作业 卧122 工艺设计写入（同井号跨阶段）",
       (a.action, a.new_value), ("write", "2026-04-02"))
    eq("工艺设计落在 I 列", a.target_col, 8)

    a = _find("7.修井作业", 6, "工艺设计")
    eq("7.修井作业 黄202H1-2 工艺设计写入", (a.action, a.new_value), ("write", "2026-04-05"))

    # ---- 4. 列自适应定位 ----
    sh6 = track.sheets[0]
    sh7 = track.sheets[1]
    eq("sheet6 地质设计实际完成列 = C(2)", sh6.locate("地质设计")["actual"], 2)
    eq("sheet6 无工艺设计", sh6.locate("工艺设计"), None)
    eq("sheet7 工艺设计实际完成列 = I(8)", sh7.locate("工艺设计")["actual"], 8)
    eq("同阶段在不同 sheet 列号一致（本例）",
       sh6.locate("地质设计")["actual"], sh7.locate("地质设计")["actual"])

    # ---- 5. diff 报告 ----
    dst = tmp / "reports" / "diff-e2e.csv"
    reports.write_diff(actions, dst)
    lines = dst.read_text(encoding="utf-8-sig").splitlines()
    eq("diff 行数 = 表头 + 动作数", len(lines), len(actions) + 1)

    # ---- 6. 幂等：重跑一次不应再产生写入 ----
    actions2: list[Action] = []
    plan = {}
    for a in actions:
        if a.action == "write":
            plan.setdefault(a.sheet, {})[(a.row - 1, a.target_col)] = a.new_value
    # 直接把写入值落到内存中的假表上，再跑一次
    fake_rows = {}
    for ts in track.sheets:
        rows = [[ts.cell(r, c) for c in range(ts.sheet.ncols)] for r in range(ts.sheet.nrows)]
        for (r, c), v in plan.get(ts.name, {}).items():
            rows[r][c] = v
        fake_rows[ts.name] = rows
    # 用真表结构重建替身：locate 结果一致
    for ts in track.sheets:
        fs = _FakeSheet(ts.name, fake_rows[ts.name],
                        {st: ts.locate(st) for st in idx})
        actions2.extend(sync_sheet(fs, idx))
    s2 = reports.summarize(actions2)
    eq("★ 重跑幂等：写入数归零", s2["write"], 0)


# ================================================================ J. Phase 4 多选源/多选节点/原地写入

def _write_trackbook_all_nodes(path: pathlib.Path) -> None:
    """造一份含全部 7 个节点的跟踪大表（单 sheet）。

    关键还原真实大表的坑：
      - 一级表头含换行：修前工程\n及概算完成情况、修井作业\n开工时间/完工时间
      - 二级表头不统一：修井作业用「实际开工 / 实际完工」而非「实际完成」
    """
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("6.上试井")
    h1 = [
        "井号",
        "地质设计", "",
        "工程方案审查", "",
        "工程设计", "",
        "工艺设计", "",
        "修前工程\n及概算完成情况", "",
        "修井作业\n开工时间", "",
        "修井作业\n完工时间", "",
    ]
    h2 = [
        "",
        "计划完成", "实际完成",
        "计划完成", "实际完成",
        "计划完成", "实际完成",
        "计划完成", "实际完成",
        "计划完成", "实际完成",
        "计划开工", "实际开工",
        "计划完工", "实际完工",
    ]
    for c, v in enumerate(h1):
        ws.write(config.ROW_STAGE_HEADER, c, v)
    for c, v in enumerate(h2):
        ws.write(config.ROW_SUB_HEADER, c, v)
    ws.write(config.ROW_DATA_START, 0, "卧122")
    wb.save(str(path))


def _write_trackbook_merge(path: pathlib.Path) -> None:
    """造一份含合并单元格的跟踪大表，用于验证原地写入保格式。"""
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("6.上试井")
    ws.write_merge(0, 0, 0, 8, "井下作业设计节点跟踪表")        # 1 个合并
    ws.write(config.ROW_STAGE_HEADER, 0, "井号")
    ws.write_merge(config.ROW_STAGE_HEADER, config.ROW_STAGE_HEADER, 1, 2, "地质设计")  # 1 个合并
    ws.write(config.ROW_SUB_HEADER, 1, "计划完成")
    ws.write(config.ROW_SUB_HEADER, 2, "实际完成")
    ws.write(config.ROW_DATA_START, 0, "卧122")                # 数据行 4 (0-based)
    wb.save(str(path))


def test_phase4_features(tmp: pathlib.Path) -> None:
    section("J. Phase 4 多选源 / 多选节点 / 原地写入")

    # ---- J1. 7 节点定位（含换行归一化 + 二级表头变体）----
    path = tmp / "all_nodes.xls"
    _write_trackbook_all_nodes(path)
    tb = TrackBook(path)
    ts = tb.sheets[0]
    for stage in config.STAGE_KEYS:
        loc = ts.locate(stage)
        check("定位节点「%s」" % stage, loc is not None,
              ("actual=%s" % loc["actual"]) if loc else "未命中")
    eq("地质设计 actual 列 = C(2)", ts.locate("地质设计")["actual"], 2)
    eq("工程方案审查 actual 列 = E(4)", ts.locate("工程方案审查")["actual"], 4)
    eq("工程设计 actual 列 = G(6)", ts.locate("工程设计")["actual"], 6)
    eq("工艺设计 actual 列 = I(8)", ts.locate("工艺设计")["actual"], 8)
    eq("修前工程及概算 actual 列 = K(10)",
       ts.locate("修前工程及概算完成情况")["actual"], 10)
    eq("修井作业开工时间 actual 列 = M(12)",
       ts.locate("修井作业开工时间")["actual"], 12)
    eq("修井作业开工时间 actual_text=实际开工",
       ts.locate("修井作业开工时间")["actual_text"], "实际开工")
    eq("修井作业完工时间 actual 列 = O(14)",
       ts.locate("修井作业完工时间")["actual"], 14)
    eq("修井作业完工时间 actual_text=实际完工",
       ts.locate("修井作业完工时间")["actual_text"], "实际完工")
    eq("available_stages 含全部 7 节点",
       sorted(ts.available_stages()), sorted(config.STAGE_KEYS))

    # ---- J2. 多源目录合并扫描 ----
    done_h = ["井号", "单位", "措施类别", "资金来源", "设计单位", "设计人",
              "流程次数", "设计日期", "完成日期", "备注"]
    review_h = ["井号", "单位", "措施类型", "资金来源", "当前审核人", "设计单位",
                "设计人", "流程次数", "设计日期", "上报日期", "备注"]
    d1 = tmp / "src1"; d2 = tmp / "src2"
    d1.mkdir(parents=True, exist_ok=True); d2.mkdir(parents=True, exist_ok=True)
    _write_detail(d1 / "常规气地质设计-已完成.xls", done_h,
                  [["卧122", "重庆", "压裂", "上市", "院", "张", 1, "", "2026-03-01", ""]])
    _write_detail(d2 / "常规气工程方案审查-审核中.xls", review_h,
                  [["峰2", "重庆", "酸化", "上市", "李", "院", "王", 1, "", "2026-03-05", ""]])
    p1, _ = scan_dir(d1)
    p2, _ = scan_dir(d2)
    merged = p1 + p2
    eq("合并两个源目录得到 2 个明细", len(merged), 2)
    from core.filename_parser import parse_filename
    stages = {parse_filename(p.name).stage for p in merged}
    check("两个节点都被识别", stages >= {"地质设计", "工程方案审查"}, stages)

    # ---- J3. 节点多选过滤 ----
    idx = build_stage_index([
        DetailRecord("卧122", "卧122", "地质设计", config.STATUS_DONE, "2026-03-01", 1, "a"),
        DetailRecord("峰2", "峰2", "工程方案审查", config.STATUS_REVIEW, "李", 1, "b"),
    ])
    filt = {"地质设计"}
    filtered = {k: v for k, v in idx.items() if k in filt}
    eq("过滤后仅剩勾选节点", sorted(filtered), ["地质设计"])

    # ---- J4. 原地写入（保留格式）----
    mp = tmp / "merge.xls"
    _write_trackbook_merge(mp)
    import xlrd
    before = xlrd.open_workbook(str(mp), formatting_info=True)
    before_merged = sum(len(s.merged_cells) for s in before.sheets())

    track = TrackBook(mp)
    act = Action(
        well="卧122", well_display="卧122", sheet="6.上试井",
        row=config.ROW_DATA_START + 1, stage="地质设计",
        target_col=2, old_value="", new_value="2026-03-10",
        action="write", reason="test")
    src_path, backup, kept = write_inplace(track, [act], backup_dir=tmp / "bk")
    eq("原地写入返回格式保留标志", kept, True)
    after = xlrd.open_workbook(str(src_path), formatting_info=True)
    after_merged = sum(len(s.merged_cells) for s in after.sheets())
    eq("合并单元格数量不变（保格式）", after_merged, before_merged)
    sh = after.sheet_by_name("6.上试井")
    eq("目标单元格已写入新值", sh.cell_value(config.ROW_DATA_START, 2), "2026-03-10")
    xf = sh.cell_xf_index(config.ROW_DATA_START, 2)
    font = after.font_list[after.xf_list[xf].font_index]
    eq("改动单元格标红", font.colour_index, 10)
    check("已自动生成备份", backup is not None and pathlib.Path(backup).exists(), backup)

    # ---- J5. 副本模式（另存 .synced）----
    copy_path = tmp / "copy_mode.xls"
    _write_trackbook_merge(copy_path)
    track2 = TrackBook(copy_path)
    out_path = tmp / "副本.synced-test.xls"
    res = write_back(track2, [act], out_path=out_path)
    check("副本文件已生成", pathlib.Path(res).exists(), res)
    rb = xlrd.open_workbook(str(res), formatting_info=False)
    eq("副本含写入值",
       rb.sheet_by_name("6.上试井").cell_value(config.ROW_DATA_START, 2), "2026-03-10")


# ================================================== K. 按措施类别自动新增井号行

_APPEND_SHEETS = ["6.上试井", "7.修井作业", "8.封堵井", "9.井口整改"]


def _write_append_trackbook(path: pathlib.Path, sheets=None,
                            with_map=None) -> None:
    """造一份带「序号」列的跟踪大表（每个 sheet 一行已有数据 + 1 个合并格）。

    with_map 不为空时额外造一张「对照表」sheet（规则随大表走）。
    """
    sheets = sheets if sheets is not None else _APPEND_SHEETS
    wb = xlwt.Workbook(encoding="utf-8")
    head_xf = xlwt.easyxf("font: bold on; align: wrap on, vert center, horiz center;")
    h1 = ["序号", "井号", "作业类型", "地质设计", "", "工程设计", "", "备注"]
    h2 = ["", "", "", "计划完成", "实际完成", "计划完成", "实际完成", ""]
    for name in sheets:
        ws = wb.add_sheet(name)
        ws.write_merge(0, 0, 0, 7, "测试表标题")   # 合并格，验证写回后仍保留
        for c, v in enumerate(h1):
            ws.write(config.ROW_STAGE_HEADER, c, v, head_xf)
        for c, v in enumerate(h2):
            ws.write(config.ROW_SUB_HEADER, c, v, head_xf)
        for c, v in enumerate([1, "卧122", "已有井", "", "", "", "", ""]):
            if v != "":
                ws.write(config.ROW_DATA_START, c, v)
    if with_map:
        ws = wb.add_sheet("对照表")
        for r, row in enumerate(with_map):
            for c, v in enumerate(row):
                ws.write(r, c, v)
    wb.save(str(path))


def _append_index() -> dict:
    """造 stage_index：4 口待新增井（4 类措施）+ 1 口大表已有的井。"""
    return {
        "地质设计": {
            "井A": DetailRecord("井A", "井A", "地质设计", config.STATUS_DONE,
                                "2026-01-01", 3, "d.xls", "永久性封井"),
            "井B": DetailRecord("井B", "井B", "地质设计", config.STATUS_DONE,
                                "2026-02-02", 1, "d.xls", "老井回采"),
            "井D": DetailRecord("井D", "井D", "地质设计", config.STATUS_REVIEW,
                                "张三", 0, "r.xls", "更换管柱"),
            "卧122": DetailRecord("卧122", "卧122", "地质设计", config.STATUS_DONE,
                                  "2025-01-01", 1, "d.xls", "永久性封井"),
        },
        "工程设计": {
            "井C": DetailRecord("井C", "井C", "工程设计", config.STATUS_DONE,
                                "2026-03-03", 2, "e.xls", "一类井口整改"),
        },
    }


def test_append_wells(tmp: pathlib.Path) -> None:
    import xlrd  # noqa: F401  (局部引入，避免影响其它分组)

    section("K. 按措施类别自动新增井号行")

    # ---- K1. 措施类别 → 目标表关键字 ----
    eq("永久性封井 → 封堵井",
       sync_engine.measure_sheet_keyword("永久性封井"), "封堵井")
    eq("老井回采 → 上试井",
       sync_engine.measure_sheet_keyword("老井回采"), "上试井")
    eq("一类井口整改 → 井口整改",
       sync_engine.measure_sheet_keyword("一类井口整改"), "井口整改")
    eq("更换管柱 → 修井作业（兜底）",
       sync_engine.measure_sheet_keyword("更换管柱"), "修井作业")
    eq("空措施类别 → 修井作业（兜底）",
       sync_engine.measure_sheet_keyword(""), "修井作业")

    # ---- K2. 分表新增 ----
    src = tmp / "append"
    src.mkdir(parents=True, exist_ok=True)
    tb = src / "井下作业设计节点跟踪大表.xls"
    _write_append_trackbook(tb)
    track = TrackBook(tb)
    idx = _append_index()

    acts, info = sync_engine.build_append_actions(track.sheets, idx)
    eq("新增 4 口井（卧122 已存在，不新增）", info["total"], 4)
    eq("分别落到 4 张表", sorted(info["by_sheet"]), sorted(_APPEND_SHEETS))
    eq("卧122 没有新增动作",
       any(a.well_display == "卧122" for a in acts), False)

    by_sheet: dict[str, list] = {}
    for a in acts:
        by_sheet.setdefault(a.sheet, []).append(a)
    eq("每张表只写 3 个单元格（序号/井号/完成时间）",
       {k: len(v) for k, v in by_sheet.items()},
       {k: 3 for k in _APPEND_SHEETS})

    amap = {(a.sheet, a.target_col): a for a in acts}
    eq("封堵井 新行序号续为 2", amap[("8.封堵井", 0)].new_value, "2")
    eq("封堵井 新行井号 = 井A", amap[("8.封堵井", 1)].new_value, "井A")
    eq("封堵井 地质设计完成时间 = 2026-01-01",
       amap[("8.封堵井", 4)].new_value, "2026-01-01")
    eq("上试井 新行井号 = 井B", amap[("6.上试井", 1)].new_value, "井B")
    eq("井口整改 完成时间写在工程设计实际列",
       amap[("9.井口整改", 6)].new_value, "2026-03-03")
    eq("修井作业 审核中写审核人（沿用既有规则）",
       amap[("7.修井作业", 4)].new_value, "张三（审核中）")
    eq("新行行号 = 第 6 行（1-based）", amap[("8.封堵井", 0)].row, 6)
    eq("新增动作全部是 write",
       all(a.action == "write" for a in acts), True)

    # ---- K3. 目标 sheet 不存在：跳过并告警 ----
    tb3 = src / "no_sheet.xls"
    _write_append_trackbook(tb3, _APPEND_SHEETS[:3])
    track3 = TrackBook(tb3)
    acts3, info3 = sync_engine.build_append_actions(
        track3.sheets, idx, create_missing_sheet=False)
    eq("缺井口整改表且不自动建表时只新增 3 口", info3["total"], 3)
    eq("井C 被跳过", [x["well"] for x in info3["no_sheet"]], ["井C"])
    eq("跳过原因关键字 = 井口整改",
       info3["no_sheet"][0]["keyword"], "井口整改")

    # ---- K4. --append-fallback 回退到修井作业表 ----
    acts4, info4 = sync_engine.build_append_actions(
        track3.sheets, idx, fallback_on_missing_sheet=True,
        create_missing_sheet=False)
    eq("回退后新增 4 口", info4["total"], 4)
    eq("井C 落到修井作业表",
       info4["by_sheet"]["7.修井作业"]["wells"], ["井C", "井D"])

    # ---- K5. 真实写回（原地保格式）+ 幂等 ----
    _src_path, _bak, kept = write_inplace(track, acts,
                                          backup_dir=src, make_backup=False)
    eq("原地写入保留格式", kept, True)
    track2 = TrackBook(tb)
    _acts5, info5 = sync_engine.build_append_actions(track2.sheets, idx)
    eq("重跑不再新增（幂等）", info5["total"], 0)

    ts8 = next(t for t in track2.sheets if "封堵井" in t.name)
    eq("封堵井行数 5 → 6", ts8.sheet.nrows, 6)
    eq("新行序号 = 2", ts8.cell(5, 0), "2")
    eq("新行井号 = 井A", ts8.cell(5, 1), "井A")
    eq("新行完成时间 = 2026-01-01", ts8.cell(5, 4), "2026-01-01")
    eq("新行其它列留空", ts8.cell(5, 2), "")
    rb = xlrd.open_workbook(str(tb), formatting_info=True)
    eq("合并单元格数量不变",
       len(rb.sheet_by_name("8.封堵井").merged_cells), 1)

    # ---- K6. 副本模式也能写出新增行 ----
    out = src / "copy.xls"
    write_back(track, acts, out_path=out)
    rb2 = xlrd.open_workbook(str(out))
    eq("副本模式行数 5 → 6", rb2.sheet_by_name("8.封堵井").nrows, 6)
    eq("副本新行井号 = 井A",
       rb2.sheet_by_name("8.封堵井").cell_value(5, 1), "井A")

    # ---- K7. detail_reader 认两种措施表头 ----
    d1 = src / "常规气地质设计-已完成.xls"
    _write_detail(d1, ["井号", "措施类别", "流程次数", "完成日期"],
                  [["井X", "永久性封井", 2, "2026-05-05"]])
    recs1 = load_all([parse_filename(str(d1))])
    eq("读到「措施类别」", recs1[0].measure, "永久性封井")

    d2 = src / "常规气地质设计-审核中.xls"
    _write_detail(d2, ["井号", "措施类型", "当前审核人", "上报日期"],
                  [["井Y", "老井回采", "李四", "2026-06-06"]])
    recs2 = load_all([parse_filename(str(d2))])
    eq("读到「措施类型」", recs2[0].measure, "老井回采")

    # ---- K8. 分表规则从大表「对照表」读取，改表即改行为 ----
    tb8 = src / "with_map.xls"
    custom_map = [
        ["措施类别", "对应表单"],
        ["永久性封井", "上试井"],      # 故意与默认规则不同
        ["老井回采", "上试井"],
        ["一类井口整改", "井口整改"],
        ["其他", "修井作业"],
    ]
    _write_append_trackbook(tb8, _APPEND_SHEETS, with_map=custom_map)
    track8 = TrackBook(tb8)
    mp, fb = sync_engine.load_measure_map(track8.sheets)
    eq("对照表规则覆盖默认（永久性封井→上试井）", mp["永久性封井"], "上试井")
    eq("对照表「其他」作为兜底表", fb, "修井作业")
    eq("自定义映射解析生效",
       sync_engine.measure_sheet_keyword("永久性封井", mp, fb), "上试井")
    _a8, info8 = sync_engine.build_append_actions(track8.sheets, idx)
    eq("井A 按对照表落到上试井表",
       info8["by_sheet"]["6.上试井"]["wells"], ["井A", "井B"])
    eq("「对照表」自身不参与井号匹配",
       "对照表" in info8["by_sheet"], False)

    # 无对照表时回退到 config 默认规则
    mp0, fb0 = sync_engine.load_measure_map(track.sheets)
    eq("无对照表时用默认规则", mp0["永久性封井"], "封堵井")
    eq("无对照表时用默认兜底", fb0, "修井作业")

    # ---- K9. 目标表缺失时按模板表头自动新建 sheet ----
    tb9 = src / "create_sheet.xls"
    _write_append_trackbook(tb9, _APPEND_SHEETS[:3])
    track9 = TrackBook(tb9)
    acts9, info9 = sync_engine.build_append_actions(track9.sheets, idx)
    eq("自动新建后 4 口井全部落表", info9["total"], 4)
    eq("新建表名按编号规律生成", list(info9["new_sheets"]), ["9.井口整改"])
    eq("表头模板取「修井作业」表",
       list(info9["new_sheets"].values()), ["7.修井作业"])
    eq("井C 落到新建表", info9["by_sheet"]["9.井口整改"]["wells"], ["井C"])
    eq("不再有被跳过的井", info9["no_sheet"], [])

    _p9, _b9, kept9 = write_inplace(track9, acts9, backup_dir=src,
                                    make_backup=False,
                                    new_sheets=info9["new_sheets"])
    eq("含新建表的写回仍保留格式", kept9, True)
    track9b = TrackBook(tb9)
    eq("sheet 数 3 → 4", len(track9b.sheets), 4)
    eq("新 sheet 名", track9b.sheet_names[-1], "9.井口整改")
    ns = track9b.sheets[-1]
    eq("新表井号列可按表头定位", ns.well_col(), 1)
    eq("新表序号列可按表头定位", ns.seq_col(), 0)
    eq("新表序号 = 1", ns.cell(4, 0), "1")
    eq("新表井号 = 井C", ns.cell(4, 1), "井C")
    eq("新表完成时间写在工程设计实际列", ns.cell(4, 6), "2026-03-03")
    eq("新表其它列留空", ns.cell(4, 2), "")
    _a9b, info9b = sync_engine.build_append_actions(track9b.sheets, idx)
    eq("新建表后重跑不再新增（幂等）", info9b["total"], 0)


# ================================================================ 主流程

def run_tests() -> bool:
    tmp = pathlib.Path(tempfile.mkdtemp(prefix="井下报表工具_test_"))
    print("测试临时目录: %s" % tmp)
    try:
        test_pathutil()
        test_filename()
        test_scan_dir(tmp)
        test_dates()
        test_decide()
        test_aggregate()
        test_sync_sheet()
        test_restructure()
        test_restructure_missing_stage()
        test_reports(tmp)
        test_end_to_end(tmp)
        test_phase4_features(tmp)
        test_append_wells(tmp)
    finally:
        pass  # 保留临时目录便于排查，见文末提示

    total = len(_RESULTS)
    failed = [r for r in _RESULTS if not r[0]]
    print("\n" + "=" * 64)
    print("  用例总数 %d，通过 %d，失败 %d" % (total, total - len(failed), len(failed)))
    if failed:
        print("\n  失败清单：")
        for _, desc, detail in failed:
            print("    - %s  %s" % (desc, detail))
    else:
        print("  所有测试通过！")
    print("=" * 64)
    print("  临时数据保留在：%s（可自行删除）" % tmp)
    return not failed


if __name__ == "__main__":
    sys.exit(0 if run_tests() else 1)
