# -*- coding: utf-8 -*-
"""统一执行入口：后台线程跑同步，通过回调把进度交回 GUI。

对齐参考项目 core/pivot_runner.py 的范式：
    run_sync(cfg, on_log, on_done, on_error)
回调均在后台线程触发，GUI 层需自行用 root.after 切回主线程。

cfg 字段:
    source       源目录（可多选，list[str] 或单个 str），含按规则命名的明细表
    trackbook    跟踪大表 .xls
    out_dir      报告输出目录（diff / 运行日志 / 自动备份落此）
    apply        True=写入；False=dry-run 只出 diff
    inplace      True=直接写入原表（默认，自动备份到 out_dir）；False=另存 .synced 副本
    stages       参与同步的节点 key 列表（多选过滤），默认全部
    allow_recheck_overwrite  允许"已完成"被"审核中"覆盖
    restructure  是否先执行 sheet 7 列结构调整
"""

from __future__ import annotations

import pathlib
import threading
import traceback

from . import config, restructure as rst
from .detail_reader import load_all
from .filename_parser import is_trackbook, scan_dir
from .reports import build_logger, summarize, write_diff
from .sync_engine import build_stage_index, sync_sheet
from .trackbook import TrackBook
from .writer import write_back, write_inplace


def _run(cfg: dict, on_log, on_done, on_error):
    try:
        raw_src = cfg.get("source")
        if isinstance(raw_src, (str, pathlib.Path)):
            sources = [pathlib.Path(raw_src)]
        else:
            sources = [pathlib.Path(s) for s in raw_src]
        if not sources:
            on_log("没有指定源目录，结束。")
            if on_done:
                on_done({"actions": [], "stat": {"total": 0, "write": 0, "skip": 0, "warn": 0},
                         "diff": None, "synced": None, "backup": None,
                         "kept_format": None, "mode": "dry", "log": logger.log_path})
            return

        reports_dir = pathlib.Path(cfg.get("out_dir") or (sources[0] / "reports"))
        no_extra = cfg.get("no_extra_files", False)
        logger = build_logger(reports_dir, to_file=not no_extra)
        on_log("=" * 56)
        on_log("井下作业设计节点同步")
        for i, s in enumerate(sources, 1):
            on_log("源目录 %d  : %s" % (i, s))
        on_log("跟踪大表  : %s" % cfg["trackbook"])
        inplace = cfg.get("inplace", True)
        if cfg.get("apply"):
            mode_txt = "写入(原地)" if inplace else "写入(副本)"
            if inplace and no_extra:
                mode_txt += " · 零新文件（不生成备份/报告）"
        else:
            mode_txt = "试运行(dry-run)"
        on_log("模式      : %s" % mode_txt)
        on_log("=" * 56)

        # ---- 1. 扫描 + 解析（逐目录合并）----
        parsed, rejected = [], []
        for src in sources:
            p, r = scan_dir(src)
            parsed.extend(p)
            rejected.extend(r)
        # 排除跟踪大表自身
        parsed = [p for p in parsed if not is_trackbook(p.path)]
        on_log("扫描完成：命中 %d 个明细文件（来自 %d 个源目录）" % (len(parsed), len(sources)))
        for p in parsed:
            on_log("  ✓ [%s / %s] %s" % (p.stage, p.status, p.name))
        for reason, name in rejected:
            on_log("  ✗ %s — %s" % (reason, name))

        if not parsed:
            on_log("没有可处理的明细文件，结束。")
            if on_done:
                on_done({"actions": [], "stat": {"total": 0, "write": 0, "skip": 0, "warn": 0},
                         "diff": None, "synced": None, "backup": None,
                         "kept_format": None, "mode": "dry", "log": logger.log_path})
            return

        # ---- 2. 加载明细 ----
        records = load_all(parsed)
        stage_index = build_stage_index(records)
        on_log("明细加载完成：%d 条记录，阶段=%s"
               % (len(records), ", ".join(stage_index.keys())))

        # ---- 2.5 节点多选过滤 ----
        stages_filter = set(cfg.get("stages") or config.STAGE_KEYS)
        present = set(stage_index.keys())
        stage_index = {k: v for k, v in stage_index.items() if k in stages_filter}
        skipped = present - stages_filter
        on_log("参与同步的节点：%s" % (", ".join(sorted(stages_filter & present)) or "（无）"))
        if skipped:
            on_log("  已排除节点：%s" % ", ".join(sorted(skipped)))

        # ---- 3. 可选：列结构调整 ----
        synced_note = None
        if cfg.get("restructure"):
            synced_note = _do_restructure(cfg, logger, on_log, first_source=sources[0])

        # ---- 4. 打开跟踪大表 + 同步 ----
        track = TrackBook(cfg["trackbook"])
        on_log("跟踪大表已打开：%s" % ", ".join(track.sheet_names))

        actions = []
        for ts in track.sheets:
            sheet_actions = sync_sheet(ts, stage_index,
                                       cfg.get("allow_recheck_overwrite"))
            if sheet_actions:
                on_log("  [%s] 产生 %d 个动作" % (ts.name, len(sheet_actions)))
            actions.extend(sheet_actions)
            for c in {a.stage for a in sheet_actions}:
                cols = ts.locate(c)
                if cols and cols.get("legacy_used"):
                    on_log("    ⚠ %s 缺少'实际完成'列，已回退到'当前进度'；"
                           "建议先执行列结构调整" % c)

        # ---- 5. 报告 ----
        stat = summarize(actions)
        diff_path = None
        if not no_extra:
            diff_path = write_diff(actions, reports_dir / ("diff-%s.csv" % _ts()))
        on_log("-" * 56)
        on_log("动作合计 %d：写入 %d / 跳过 %d / 告警 %d"
               % (stat["total"], stat["write"], stat["skip"], stat["warn"]))
        if diff_path:
            on_log("Diff 报告：%s" % diff_path)

        # ---- 6. 写回 ----
        backup = synced = kept_format = None
        result_mode = "dry"
        if cfg.get("apply"):
            if stat["write"] == 0:
                on_log("没有需要写入的变更，跳过写回。")
            else:
                result_mode = "inplace" if cfg.get("inplace", True) else "copy"
                if result_mode == "inplace":
                    src_path, backup, kept_format = write_inplace(
                        track, actions, backup_dir=reports_dir,
                        make_backup=not no_extra)
                    synced = src_path
                    on_log("已直接写入原表：%s" % src_path)
                    if backup:
                        on_log("原表已自动备份：%s" % backup)
                    else:
                        on_log("（零新文件模式：未生成备份，原表已直接修改）")
                    if kept_format:
                        on_log("（已保留合并单元格/字体/边框等原表格式，改动单元格标红）")
                    else:
                        on_log("⚠ 未能保留原表格式（xlutils 不可用），已用纯 xlwt 重建；请用 WPS 核对。")
                else:
                    backup = track.backup()
                    synced = write_back(
                        track, actions,
                        out_path=reports_dir / ("井下作业设计节点跟踪大表.synced-%s.xls" % _ts()))
                    on_log("已生成同步后副本（原表未改动）：%s" % synced)
                    on_log("原表已备份：%s" % backup)
                    on_log("⚠ 副本为 xlwt 重建，合并单元格/条件格式丢失，需核对后手动替换。")
        else:
            on_log("试运行模式：未修改任何文件。勾选『实际写入』才会写回。")

        on_log("运行日志：%s" % logger.log_path)
        on_log("=" * 56)

        if on_done:
            on_done({"actions": actions, "stat": stat, "diff": diff_path,
                     "synced": synced, "backup": backup, "kept_format": kept_format,
                     "mode": result_mode, "log": logger.log_path,
                     "restructure": synced_note})

    except Exception as e:
        tb = traceback.format_exc()
        if on_error:
            on_error(e, tb)
        else:
            on_log("[错误] %s" % e)
            on_log(tb)


def _do_restructure(cfg, logger, on_log, first_source=None) -> dict | None:
    import xlrd
    from .pathutil import safe_path

    tb_path = safe_path(cfg["trackbook"])
    book = xlrd.open_workbook(str(tb_path), formatting_info=False)
    names = book.sheet_names()
    s7 = rst.find_sheet7(names)
    if not s7:
        on_log("未找到 sheet 7，跳过列结构调整。")
        return None

    snaps = [rst.snapshot(book.sheet_by_name(n)) for n in names]
    idx = names.index(s7)
    before = snaps[idx]
    try:
        plan = rst.plan_columns(before)
    except ValueError as e:
        on_log("列结构调整跳过：%s" % e)
        return None

    if plan.insert_at is None:
        on_log("sheet 7 列结构已是标准（含'实际完成'），无需调整。")
        return None

    after = rst.apply_plan(before, plan)
    snaps[idx] = after
    out_dir = pathlib.Path(cfg.get("out_dir") or (pathlib.Path(first_source or sources[0]) / "reports"))
    new_xls = out_dir / ("井下作业设计节点跟踪大表_重列-%s.xls" % _ts())
    rst.write_snapshots(snaps, new_xls)
    rep = rst.write_report(before, after, plan,
                           out_dir / ("restructure-report-%s.csv" % _ts()))
    on_log("列结构调整：%s 列 %d → %d" % (s7, before.ncols_old,
                                        max(len(r) for r in after.rows)))
    on_log("  新表：%s" % new_xls)
    on_log("  对照：%s" % rep)
    return {"sheet": s7, "new_xls": new_xls, "report": rep,
            "ncols": (before.ncols_old, max(len(r) for r in after.rows))}


def _ts() -> str:
    import time
    return time.strftime("%Y%m%d-%H%M%S")


def run_sync(cfg: dict, on_log=None, on_done=None, on_error=None):
    """在后台线程执行同步。"""
    on_log = on_log or (lambda m: None)
    t = threading.Thread(target=_run, args=(cfg, on_log, on_done, on_error),
                         daemon=True)
    t.start()
    return t


def default_cfg() -> dict:
    return {
        "source": [str(config.DEFAULT_SOURCE_DIR)],
        "trackbook": str(config.DEFAULT_SOURCE_DIR / config.DEFAULT_TRACKBOOK_NAME),
        "out_dir": None,
        "apply": False,
        "inplace": True,
        "no_extra_files": True,
        "stages": list(config.STAGE_KEYS),
        "allow_recheck_overwrite": config.ALLOW_RECHECK_OVERWRITE,
        "restructure": False,
    }
