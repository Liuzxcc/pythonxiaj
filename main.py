# -*- coding: utf-8 -*-
"""进度跟踪报表工具 - 主入口

用法:
    python main.py                 # 启动图形界面（默认）
    python main.py --auto          # 启动后立即执行一次（使用界面中的默认路径）
    python main.py --cli           # 命令行模式（无界面）
    python main.py --cli --apply   # 命令行模式并实际写回

命令行示例:
    python main.py --cli \
        --source  "/path/to/井下作业报表生成" \
        --trackbook "/path/to/井下作业设计节点跟踪大表.xls" \
        --out-dir  "/path/to/reports"
"""

import argparse
import os
import sys

# 确保项目根目录在 Python 路径中
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core import config  # noqa: E402  (供 argparse choices 使用)


def check_dependencies():
    """检查依赖是否安装"""
    missing = []
    for mod, pkg in [("xlrd", "xlrd==1.2.0"), ("xlwt", "xlwt"),
                     ("openpyxl", "openpyxl")]:
        try:
            __import__(mod)
        except ImportError:
            missing.append(pkg)

    if missing:
        print("缺少依赖包:", ", ".join(missing))
        print("请运行: pip install " + " ".join(missing))
        print("或运行: pip install -r requirements.txt")
        sys.exit(1)

    # xlrd >= 2.0 无法读取 .xls，需显式拦截
    import xlrd
    if int(xlrd.__version__.split(".")[0]) >= 2:
        print("错误：xlrd %s 不支持 .xls（BIFF8）格式。" % xlrd.__version__)
        print("请降级: pip install xlrd==1.2.0")
        sys.exit(1)


def run_cli(args):
    """命令行模式：同步执行（不走后台线程）。"""
    from core import config
    from core.runner import _run, default_cfg

    cfg = default_cfg()
    if args.source:
        cfg["source"] = args.source
    if args.trackbook:
        cfg["trackbook"] = args.trackbook
    if args.out_dir:
        cfg["out_dir"] = args.out_dir
    cfg["apply"] = args.apply
    cfg["allow_recheck_overwrite"] = args.allow_recheck_overwrite
    cfg["restructure"] = args.restructure
    cfg["stages"] = args.stage if args.stage else list(config.STAGE_KEYS)
    cfg["inplace"] = not bool(args.copy_mode)
    cfg["no_extra_files"] = not bool(args.report)

    result = {}

    def on_log(msg):
        print(msg)

    def on_done(r):
        result.update(r)

    def on_error(e, tb):
        print("[错误] %s" % e)
        print(tb)
        result["failed"] = True

    _run(cfg, on_log, on_done, on_error)
    return 0 if not result.get("failed") else 1


def main():
    check_dependencies()

    parser = argparse.ArgumentParser(description="进度跟踪报表工具")
    parser.add_argument("--cli", action="store_true",
                        help="命令行模式（无图形界面）")
    parser.add_argument("--auto", action="store_true",
                        help="（图形界面）启动后立即执行一次")
    parser.add_argument("--source", action="append",
                        help="源目录（可多次指定，合并扫描；默认取配置中的目录）")
    parser.add_argument("--trackbook", help="跟踪大表 .xls 路径")
    parser.add_argument("--out-dir", help="报告输出目录（diff/日志/自动备份落此）")
    parser.add_argument("--apply", action="store_true",
                        help="实际写回（默认试运行 dry-run，不改文件）")
    parser.add_argument("--copy-mode", action="store_true",
                        help="谨慎模式：不改原表，另存 .synced 副本（GUI 已无此选项，仅 CLI 可用；默认原地写原表、零新文件）")
    parser.add_argument("--report", action="store_true",
                        help="生成备份(.bak)与核对报告(diff CSV/日志)；默认不生成任何额外文件，仅改原表")
    parser.add_argument("--stage", action="append", choices=list(config.STAGE_KEYS),
                        help="参与的同步节点（可多次指定）；默认全部节点")
    parser.add_argument("--allow-recheck-overwrite", action="store_true",
                        help="允许『已完成』被『审核中』覆盖（回退场景）")
    parser.add_argument("--restructure", action="store_true",
                        help="先执行 sheet 7 列结构调整（工艺设计补两列）")
    args = parser.parse_args()

    if args.cli:
        sys.exit(run_cli(args))

    from gui.main_window import SyncWindow
    SyncWindow(auto=args.auto).run()


if __name__ == "__main__":
    main()
