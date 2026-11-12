# -*- coding: utf-8 -*-
"""全局配置：节点定义、状态关键字、决策策略、默认路径。

所有业务关键字集中在此，新增节点只需在 STAGES 追加一项，无需改代码。
"""

from __future__ import annotations

import dataclasses as dc
import pathlib

# ---------------------------------------------------------------- 节点定义


@dc.dataclass(frozen=True)
class Stage:
    """一个同步节点（跟踪大表的一级表头块）。

    key       内部标识，也是 GUI 复选框的显示名
    headers   一级表头可能的文本（归一化后比对，换行/全角空格已剔除）
    actuals   二级表头中「实际」列可能的文本（按优先级）
    plans     二级表头中「计划」列可能的文本（仅用于块定位，不写入）
    keywords  文件名匹配关键字（长关键字在前，正则按长度降序拼接）
    """

    key: str
    headers: tuple
    actuals: tuple
    plans: tuple
    keywords: tuple


# 7 个可同步节点。
# 注意：跟踪大表的一级表头含换行（如 "修前工程\n及概算完成情况"），
#       二级表头也不统一（修井作业是「实际开工 / 实际完工」而非「实际完成」），
#       所以 header / actual 都必须是变体列表，不能硬编码单一文本。
STAGES: tuple = (
    Stage("地质设计",
          ("地质设计",),
          ("实际完成", "当前进度"),
          ("计划完成",),
          ("地质设计",)),
    Stage("工程方案审查",
          ("工程方案审查", "工程方案"),
          ("实际完成", "当前进度"),
          ("计划完成",),
          ("工程方案审查", "工程方案")),
    Stage("工程设计",
          ("工程设计",),
          ("实际完成", "当前进度"),
          ("计划设计", "计划完成"),
          ("工程设计",)),
    Stage("工艺设计",
          ("工艺设计",),
          ("实际完成", "当前进度"),
          ("计划完成",),
          ("工艺设计",)),
    Stage("修前工程及概算完成情况",
          ("修前工程及概算完成情况", "修前工程及概算", "修前工程"),
          ("实际完成", "当前进度"),
          ("计划完成",),
          ("修前工程及概算", "修前工程")),
    Stage("修井作业开工时间",
          ("修井作业开工时间", "修井作业开工"),
          ("实际开工", "实际完成"),
          ("计划开工",),
          ("修井作业开工",)),
    Stage("修井作业完工时间",
          ("修井作业完工时间", "修井作业完工"),
          ("实际完工", "实际完成"),
          ("计划完工",),
          ("修井作业完工",)),
)

# 全部节点的 key，按声明顺序
STAGE_KEYS: tuple = tuple(s.key for s in STAGES)

# 文件名阶段关键字（扁平列表，长关键字优先由 _compile 保证）
STAGE_KEYWORDS: list = [kw for s in STAGES for kw in s.keywords]

# 一级表头文本 → 节点 key 的反查表
HEADER_TO_STAGE: dict = {h: s.key for s in STAGES for h in s.headers}

# 状态关键字（文件名后缀）
STATUS_DONE = "已完成"
STATUS_REVIEW = "审核中"
STATUS_KEYWORDS = [STATUS_DONE, STATUS_REVIEW]

# 非预期状态：落入"人工确认"桶，不自动处理
STATUS_AMBIGUOUS = ["已驳回", "已退回", "草稿", "作废"]

# ---------------------------------------------------------------- 表头定位

# 一级表头所在行（0-based），三个 sheet 一致
ROW_STAGE_HEADER = 2
# 二级表头（计划完成 / 实际完成）所在行
ROW_SUB_HEADER = 3
# 数据起始行
ROW_DATA_START = 4

# 一级表头里的井号列名
COL_NAME_WELL = "井号"

# 二级表头默认名（新增节点/兜底时使用；具体节点以 Stage.actuals / Stage.plans 为准）
COL_NAME_ACTUAL = "实际完成"
COL_NAME_PLAN = "计划完成"
# v1.0 旧模板回退列名（sheet 7 工艺设计曾用"当前进度"）
COL_NAME_ACTUAL_LEGACY = "当前进度"

# ---------------------------------------------------------------- 明细表字段

# 明细表表头行（已完成 / 审核中 表的表头都在第 0 行）
ROW_DETAIL_HEADER = 0

# 明细表列名候选（按优先级）
DETAIL_COL_WELL = ["井号"]
DETAIL_COL_DONE_DATE = ["完成日期", "Completion Date"]
DETAIL_COL_SUBMIT_DATE = ["上报日期"]
DETAIL_COL_REVIEWER = ["当前审核人", "审核人"]
DETAIL_COL_FLOW = ["流程次数"]
# 措施类别：已完成表叫「措施类别」，审核中表叫「措施类型」，两种表头都要认
DETAIL_COL_MEASURE = ["措施类别", "措施类型"]

# ---------------------------------------------------------------- 自动新增井号行

# 措施类别 → 目标 sheet 名应包含的关键字（按需求：按措施类别分表新增）
# 说明：示例中「一类井口整改」对应的表在当前样本大表里尚不存在，
#       若大表补齐了表名含「井口整改」的 sheet，工具会自动识别并落入该表。
MEASURE_SHEET_KEYWORDS: dict = {
    "永久性封井": "封堵井",
    "老井回采": "上试井",
    "一类井口整改": "井口整改",
}
# 其余措施类别（更换管柱 / 回注井修井 / 连续油管作业 …）统一落此表
MEASURE_FALLBACK_KEYWORD = "修井作业"

# 一级表头里的序号列名
COL_NAME_SEQ = "序号"

# 功能开关：自动为「明细有、大表无」的井新增行（默认开启）
APPEND_MISSING_WELLS = True
# 目标 sheet 不存在时是否回退到兜底表：默认 False —— 宁可跳过并告警，也不把数据放错表
APPEND_FALLBACK_ON_MISSING_SHEET = False
# 目标 sheet 不存在时，按现有业务表的表头结构自动新建该 sheet（规则来自大表「对照表」）
# 例：对照表写了「一类井口整改 → 井口整改」但大表没有该表 → 自动建「9.井口整改-陈亚妮」
APPEND_CREATE_MISSING_SHEET = True

# ---------------------------------------------------------------- 全表格式统一

# v1.6：写回时把整表所有非空单元格统一成
#   宋体 10 号 / 黑色细边框 / 水平垂直居中，日期统一 yyyy-MM-dd。
# 默认开启；命令行 --no-unify-format 可关闭，回到「只写变更单元格」的旧行为。
UNIFY_FORMAT = True
# v1.6.1：边框铺满整个「使用区域矩形」，空白单元格也补黑框，中间不断框。
# 默认开启；命令行 --border-data-only 可退回 v1.6.0 的「只给有数据的格子加框」。
UNIFY_FULL_GRID = True
# v1.6.2：字号只统一「表头下面的数据行」，标题行 + 表头（r < ROW_DATA_START）
# 保留原本的字号，避免把大标题压成 10 号。设 False 可恢复 v1.6.1 全表 10 号。
PRESERVE_HEADER_FONT_SIZE = True
# v1.6.2：已有单元格的字体颜色一律保留；只有工具本次新增/修改的才上深蓝色。
# 设 True 可恢复 v1.6.1 的「历史红字一并迁移为深蓝」。
MIGRATE_LEGACY_RED = False

# ---------------------------------------------------------------- 决策策略

# 已完成日期 → 已完成的日期被"审核中"覆盖（规则③）
ALLOW_RECHECK_OVERWRITE = False
# 同井号在明细表多条记录时，取流程次数最大者
PREFER_MAX_FLOW = True
# 跟踪大表同一 sheet 内重复井号：仅第一条写入
DUPLICATE_WELL_STRATEGY = "first_row_only"

# ---------------------------------------------------------------- 输出

# 审核中写入值的后缀
REVIEW_SUFFIX = "（审核中）"

# 日期输出格式
DATE_FMT = "%Y-%m-%d"

# 支持的文件扩展名
SUPPORTED_EXTS = {".xls", ".xlsx"}

# ---------------------------------------------------------------- 默认路径

DEFAULT_SOURCE_DIR = pathlib.Path(
    "/Users/zoe/Documents/重庆气矿项目/报表开发的支撑文件/井下作业报表生成"
)
DEFAULT_TRACKBOOK_NAME = "井下作业设计节点跟踪大表.xls"

# sheet 7 标识（用于列结构调整）
SHEET7_MARKERS = ("7.", "修井作业")

# 需要补两列的阶段（v1.1）
STAGE_NEEDS_TWO_COLS = "工艺设计"
