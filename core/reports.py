# -*- coding: utf-8 -*-
"""结果报告：diff CSV + 运行日志。

输出编码用 UTF-8 BOM（utf-8-sig），保证 Excel 双击直接打开中文不乱码。
"""

from __future__ import annotations

import csv
import logging
import pathlib
import time

from .sync_engine import Action

DIFF_HEADER = ["井号", "sheet", "行号", "阶段", "目标列", "改前", "改后", "动作", "说明"]


def _col_letter(idx: int) -> str:
    """0-based 列号 → Excel 列字母（0→A，26→AA）。"""
    s = ""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def write_diff(actions: list[Action], dst: pathlib.Path) -> pathlib.Path:
    """写 diff CSV。"""
    dst = pathlib.Path(dst)
    dst.parent.mkdir(parents=True, exist_ok=True)
    with dst.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(DIFF_HEADER)
        for a in actions:
            w.writerow([
                a.well_display, a.sheet, a.row, a.stage,
                "%s(%d)" % (_col_letter(a.target_col), a.target_col),
                a.old_value, a.new_value, a.action, a.reason,
            ])
    return dst


def summarize(actions: list[Action]) -> dict:
    """汇总统计。"""
    return {
        "total": len(actions),
        "write": sum(1 for a in actions if a.action == "write"),
        "skip": sum(1 for a in actions if a.action == "skip"),
        "warn": sum(1 for a in actions if a.action.startswith("warn")),
    }


def build_logger(reports_dir: pathlib.Path, name: str = "well-sync",
                 to_file: bool = True) -> logging.Logger:
    """建日志器。

    to_file=True  → 文件 + 控制台；to_file=False → 仅控制台（「真正零新文件」模式，不建目录、不写日志文件）。
    """
    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    fmt = logging.Formatter("%(asctime)s %(levelname)s %(message)s")

    if to_file:
        reports_dir = pathlib.Path(reports_dir)
        reports_dir.mkdir(parents=True, exist_ok=True)
        ts = time.strftime("%Y%m%d-%H%M%S")
        log_path = reports_dir / ("run-%s.log" % ts)
        fh = logging.FileHandler(log_path, encoding="utf-8")
        fh.setFormatter(fmt)
        logger.addHandler(fh)
        logger.log_path = log_path  # type: ignore[attr-defined]
    else:
        logger.log_path = None  # type: ignore[attr-defined]

    sh = logging.StreamHandler()
    sh.setFormatter(fmt)
    logger.addHandler(sh)
    return logger
