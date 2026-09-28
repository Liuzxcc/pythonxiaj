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

# 应用版本号：主版本.次版本.修订号
#   1.1 初始写回能力     1.2 零新文件     1.3 GUI 极简（多步骤向导）
#   1.4 GUI 再简化为「单目录 + 状态记忆 + 自动识别跟踪大表」，
#       并修复 xlutils None 格式崩溃、写入单元格统一文本格式（(fXt165) 乱码）
#   1.4.1 修复打包后不记住上次目录：状态文件从「程序所在目录」移到用户级配置目录
#         （Windows onefile 解压目录退出即失效）
__version__ = "1.4.1"
