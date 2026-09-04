# -*- coding: utf-8 -*-
"""井下作业设计节点同步系统 — 核心层。

模块划分：
    config          阶段/状态关键字、决策策略、默认路径
    pathutil        跨平台路径与文件名归一（NFC / casefold）
    filename_parser 文件名 → (stage, status)
    detail_reader   明细表 → DetailRecord 列表
    trackbook       跟踪大表读写 + 列自适应定位
    sync_engine     井号匹配 + 冲突决策
    restructure     sheet 7 列结构调整
    writer          写回 .xls
    reports         diff CSV + 日志
    runner          后台线程统一执行入口
"""

from . import (config, detail_reader, filename_parser, pathutil, reports,
               restructure, runner, sync_engine, trackbook, writer)
from .sync_engine import Action, DetailRecord, decide

__all__ = [
    "config", "pathutil", "filename_parser", "detail_reader",
    "trackbook", "sync_engine", "restructure", "writer", "reports", "runner",
    "Action", "DetailRecord", "decide",
]

__version__ = "1.1.0"
