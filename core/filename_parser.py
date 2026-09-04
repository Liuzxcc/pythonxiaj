# -*- coding: utf-8 -*-
"""文件名解析：把 ``<类别前缀><阶段>-<状态>.xls`` 拆成 (stage, status)。

井号不在文件名里，而在明细表的 A 列 —— 文件名只描述"阶段 × 状态"。
"""

from __future__ import annotations

import dataclasses as dc
import os
import pathlib
import re

from . import config
from .pathutil import canon, nfc, safe_filename


@dc.dataclass
class ParsedFile:
    path: pathlib.Path
    name: str      # NFC 归一后的文件名
    stage: str     # 阶段，如 "地质设计"
    status: str    # 状态，"已完成" / "审核中"
    ext: str       # 小写扩展名


def _compile(words: list[str]) -> re.Pattern:
    """长关键字优先：按长度降序，避免 "工程" 截胡 "工程方案审查"。"""
    ordered = sorted(words, key=len, reverse=True)
    return re.compile("|".join(re.escape(w) for w in ordered))


_STAGE_RE = _compile(config.STAGE_KEYWORDS)
_STATUS_RE = _compile(config.STATUS_KEYWORDS)
_AMBIG_RE = _compile(config.STATUS_AMBIGUOUS)


def parse_filename(path) -> ParsedFile | None:
    """解析文件名。命中返回 ParsedFile，否则 None（调用方归入"未识别"桶）。"""
    path = pathlib.Path(path)
    name = safe_filename(path.name)
    stem, ext = os.path.splitext(name)
    ext = ext.lower()
    if ext not in config.SUPPORTED_EXTS:
        return None

    m_stage = _STAGE_RE.search(stem)
    m_status = _STATUS_RE.search(stem)
    if not (m_stage and m_status):
        return None

    return ParsedFile(
        path=path,
        name=name,
        stage=m_stage.group(0),
        status=m_status.group(0),
        ext=ext,
    )


def classify(path) -> tuple[ParsedFile | None, str]:
    """分类单个文件，返回 (ParsedFile|None, 归类原因)。

    归类原因: "ok" / "非法文件名" / "扩展名不支持" / "关键字不匹配" / "状态存疑"
    """
    try:
        name = safe_filename(pathlib.Path(path).name)
    except ValueError as e:
        return None, "非法文件名（%s）" % e

    stem, ext = os.path.splitext(name)
    if ext.lower() not in config.SUPPORTED_EXTS:
        return None, "扩展名不支持"

    m_stage = _STAGE_RE.search(stem)
    if not m_stage:
        return None, "关键字不匹配（未识别阶段）"

    m_status = _STATUS_RE.search(stem)
    if not m_status:
        if _AMBIG_RE.search(stem):
            return None, "状态存疑（需人工确认）"
        return None, "关键字不匹配（未识别状态）"

    return parse_filename(path), "ok"


def scan_dir(src) -> tuple[list[ParsedFile], list[tuple[str, str]]]:
    """扫描目录。返回 (命中列表, [(原因, 文件名), ...])。"""
    src = pathlib.Path(src)
    parsed: list[ParsedFile] = []
    rejected: list[tuple[str, str]] = []

    for p in sorted(src.iterdir()):
        if not p.is_file():
            continue
        if p.name.startswith("."):        # macOS 隐藏文件
            continue
        if p.name == "~$" or p.name.startswith("~$"):   # Excel 锁文件
            continue
        pf, reason = classify(p)
        if pf:
            parsed.append(pf)
        else:
            rejected.append((reason, p.name))
    return parsed, rejected


def is_trackbook(path) -> bool:
    """判断是否为跟踪大表（需从明细中排除）。"""
    return nfc(pathlib.Path(path).name) == nfc(config.DEFAULT_TRACKBOOK_NAME)
