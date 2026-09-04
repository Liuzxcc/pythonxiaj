# -*- coding: utf-8 -*-
"""跟踪大表读写 + 列自适应定位。

关键约束：同一阶段在不同 sheet 的列号不同（工艺设计仅 sheet 7 有，
且 sheet 6 的 L 列是修前工程），必须按表头文本定位，严禁硬编码列号。
"""

from __future__ import annotations

import pathlib
import shutil
import time

import xlrd

from . import config
from .detail_reader import _cell, find_col
from .pathutil import nfc, norm_header, safe_path


def stage_config(stage_key: str):
    """按节点 key 取 Stage 配置；找不到返回一个宽松的默认配置。"""
    for s in config.STAGES:
        if s.key == stage_key:
            return s
    return None


class TrackSheet:
    """跟踪大表中的单个 sheet，提供列自适应定位与单元格读取。"""

    def __init__(self, sheet, name: str):
        self.sheet = sheet
        self.name = name
        self._cache: dict[str, dict | None] = {}

    def cell(self, r: int, c: int) -> str:
        return _cell(self.sheet, r, c)

    def locate(self, stage: str) -> dict | None:
        """定位某阶段的"实际"列。找不到返回 None（该 sheet 无此阶段）。

        返回 {'well': int, 'actual': int, 'plan': int,
              'legacy_used': bool, 'stage': str, 'actual_text': str}
        """
        if stage in self._cache:
            return self._cache[stage]
        result = self._locate_impl(stage)
        self._cache[stage] = result
        return result

    def _locate_impl(self, stage: str) -> dict | None:
        sh = self.sheet
        stage_norm = norm_header(stage)

        # 井号列（一级表头行）
        try:
            col_well = find_col(sh, config.ROW_STAGE_HEADER, [config.COL_NAME_WELL])
        except KeyError:
            return None

        # 按一级表头切块：[(块名, [列号...])]，块名已归一化（去掉换行等）
        blocks: list[tuple[str, list[int]]] = []
        cur_text, cur_cols = "", []
        for c in range(sh.ncols):
            v = norm_header(_cell(sh, config.ROW_STAGE_HEADER, c))
            if v:
                if cur_cols:
                    blocks.append((cur_text, cur_cols))
                cur_text, cur_cols = v, [c]
            else:
                cur_cols.append(c)
        if cur_cols:
            blocks.append((cur_text, cur_cols))

        # 命中该节点的块（支持配置里的多个表头变体）
        stage_cfg = stage_config(stage)
        wanted = {stage_norm}
        if stage_cfg:
            wanted |= set(stage_cfg.headers)
        target = None
        for text, cols in blocks:
            if text in wanted:
                target = cols
                break
        if target is None:
            return None

        # 二级表头候选：该节点配置的 actuals / plans 变体
        actual_names = list(stage_cfg.actuals) if stage_cfg else \
            [config.COL_NAME_ACTUAL, config.COL_NAME_ACTUAL_LEGACY]
        plan_names = list(stage_cfg.plans) if stage_cfg else [config.COL_NAME_PLAN]

        col_actual = col_plan = -1
        legacy_used = False
        actual_text = ""
        for c in target:
            v3 = norm_header(_cell(sh, config.ROW_SUB_HEADER, c))
            if col_actual < 0 and v3 in actual_names:
                col_actual, actual_text = c, v3
            elif col_plan < 0 and v3 in plan_names:
                col_plan = c

        # 回退：v1.0 旧模板用"当前进度"
        if col_actual < 0:
            for c in target:
                v3 = norm_header(_cell(sh, config.ROW_SUB_HEADER, c))
                if v3 == config.COL_NAME_ACTUAL_LEGACY:
                    col_actual, actual_text = c, v3
                    legacy_used = True
                    break
        if col_actual < 0:
            return None

        return {"well": col_well, "actual": col_actual, "plan": col_plan,
                "legacy_used": legacy_used, "stage": stage,
                "actual_text": actual_text}

    def available_stages(self) -> list[str]:
        """列出该 sheet 上实际存在的可同步节点（按配置顺序）。"""
        out = []
        for s in config.STAGES:
            if self.locate(s.key) is not None:
                out.append(s.key)
        return out


class TrackBook:
    """跟踪大表工作簿。"""

    def __init__(self, path):
        self.path = safe_path(path)
        if self.path.suffix.lower() != ".xls":
            raise ValueError("跟踪大表须为 .xls：%s" % self.path)
        self.book = xlrd.open_workbook(str(self.path), formatting_info=False)
        self.sheets = [TrackSheet(self.book.sheet_by_name(n), n)
                       for n in self.book.sheet_names()]

    @property
    def sheet_names(self) -> list[str]:
        return self.book.sheet_names()

    def backup(self) -> pathlib.Path:
        """备份原表，返回备份路径。"""
        ts = time.strftime("%Y%m%d-%H%M%S")
        dst = self.path.with_suffix(".bak-%s.xls" % ts)
        shutil.copy2(self.path, dst)
        return dst
