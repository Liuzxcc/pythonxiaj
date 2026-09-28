# -*- coding: utf-8 -*-
"""主窗口：选源目录 → 自动读取同目录跟踪大表 → 一键按规则写入原表。

简化后的 GUI 范式：
    - 只保留一个「选择源目录」输入框 + 浏览按钮
    - 第二次打开默认使用上次选择的目录
    - 跟踪大表自动从源目录中读取（《井下作业设计节点跟踪大表.xls》）
    - 去掉节点多选区，启动时自动识别并同步全部可用节点
    - 顶部蓝色横幅 + 自绘主按钮 + 状态卡片 + 执行日志
"""

from __future__ import annotations

import json
import os
import pathlib
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from core import config
from core import __version__ as APP_VERSION
from core.runner import run_sync

# 主题色（与参考项目一致）
ACCENT = "#0078D4"
ACCENT_HOVER = "#106EBE"
BG = "#FFFFFF"
PANEL = "#F4F7FB"
OK_BG, OK_FG = "#E7F6EC", "#0E7C36"
ERR_BG, ERR_FG = "#FDECEA", "#C0392B"
RUN_BG, RUN_FG = "#FFF7E6", "#9A6B00"

FONT = "Microsoft YaHei"
FONT_MONO = "Consolas"

# 状态持久化文件：仅保存上次源目录
STATE_FILE = pathlib.Path(__file__).resolve().parent.parent / ".workbuddy" / "gui_state.json"


class SyncWindow:
    """同步工具主窗口。"""

    def __init__(self, auto: bool = False):
        self.auto = auto
        self.root = tk.Tk()
        self.root.title(f"进度跟踪报表工具 v{APP_VERSION}")
        self.root.geometry("900x620")
        self.root.minsize(820, 560)
        self.root.configure(bg=BG)

        # 源目录（单个）
        self.src_var = tk.StringVar(value=str(config.DEFAULT_SOURCE_DIR))
        # 跟踪大表路径（内部维护，不显示在 UI）
        self.track_path: str | None = None

        self._running = False

        self._build_ui()
        self._load_state()
        self._auto_detect_trackbook()

    # ================= 界面 =================

    def _accent_button(self, master, text, command):
        """自绘可悬停主按钮（macOS 兼容）。"""
        frame = tk.Frame(master, bg=ACCENT, cursor="hand2", relief=tk.FLAT, borderwidth=0)
        lbl = tk.Label(frame, text=text, bg=ACCENT, fg="white",
                       font=(FONT, 14, "bold"), cursor="hand2",
                       anchor="center", justify=tk.CENTER)
        lbl.pack(fill=tk.BOTH, expand=True, padx=10, pady=12)
        frame.label = lbl

        def on_click(e=None):
            if not self._running:
                command()

        def on_enter(e=None):
            if not self._running:
                frame.config(bg=ACCENT_HOVER)
                lbl.config(bg=ACCENT_HOVER)

        def on_leave(e=None):
            frame.config(bg=ACCENT)
            lbl.config(bg=ACCENT)

        for w in (frame, lbl):
            w.bind("<Button-1>", on_click)
            w.bind("<Enter>", on_enter)
            w.bind("<Leave>", on_leave)
        return frame

    def _build_ui(self):
        # ---- 顶部横幅 ----
        header = tk.Frame(self.root, bg=ACCENT, padx=22, pady=18)
        header.pack(fill=tk.X)
        title_row = tk.Frame(header, bg=ACCENT)
        title_row.pack(fill=tk.X)
        tk.Label(title_row, text="进度跟踪报表工具", bg=ACCENT, fg="white",
                 font=(FONT, 19, "bold")).pack(side=tk.LEFT, anchor=tk.W)
        # 右上角显示版本号，便于确认本机安装的版本（排查"跑的是不是旧包"）
        tk.Label(title_row, text=f"v{APP_VERSION}", bg=ACCENT, fg="#D6E9F8",
                 font=(FONT, 10)).pack(side=tk.RIGHT, anchor=tk.E)

        body = ttk.Frame(self.root, padding=(20, 16, 20, 8))
        body.pack(fill=tk.BOTH, expand=True)

        # ---- 源目录 ----
        src_box = ttk.LabelFrame(body, text="源目录（存放按规则命名的明细表）", padding=12)
        src_box.pack(fill=tk.X, pady=(0, 10))
        row = ttk.Frame(src_box)
        row.pack(fill=tk.X)
        ttk.Label(row, text="文件夹", width=8, font=(FONT, 11, "bold")).pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=self.src_var, font=(FONT, 10)).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=8)
        ttk.Button(row, text="浏览...", command=self._browse_src).pack(side=tk.LEFT)
        ttk.Label(src_box,
                  text="  提示：目录中需包含《井下作业设计节点跟踪大表.xls》，工具会自动识别并同步其中全部可用节点。",
                  foreground="#888888", font=(FONT, 9)).pack(anchor=tk.W, pady=(6, 0))

        # ---- 主按钮 ----
        self.run_btn = self._accent_button(body, "🔄   一 键 写 入 原 表", self._start)
        self.run_btn.pack(fill=tk.X, pady=(4, 10))

        # ---- 状态卡片 ----
        self.status_label = tk.Label(
            body, text="选择源目录后，点击「一键写入原表」。",
            bg=PANEL, fg="#555555", anchor=tk.W, justify=tk.LEFT,
            font=(FONT, 10), padx=12, pady=8, relief=tk.GROOVE,
            borderwidth=1, wraplength=820)
        self.status_label.pack(fill=tk.X, pady=(0, 8))

        # ---- 日志 ----
        log_box = ttk.LabelFrame(body, text="执行日志", padding=8)
        log_box.pack(fill=tk.BOTH, expand=True)
        self.log_text = tk.Text(log_box, height=14, wrap=tk.WORD, font=(FONT_MONO, 10))
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb = ttk.Scrollbar(log_box, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.config(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        # ---- 底部提示 ----
        footer = ttk.Frame(self.root, padding=(20, 10, 20, 14))
        footer.pack(fill=tk.X)
        ttk.Label(footer,
                  text="提示：文件名需含「阶段-状态」，如 地质设计-已完成 / 工艺设计-审核中；井号取自明细表 A 列。",
                  foreground="#888888").pack(side=tk.LEFT)

    # ================= 状态持久化 =================

    def _load_state(self):
        """加载上次使用的源目录。"""
        try:
            if STATE_FILE.exists():
                data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
                last_dir = data.get("last_source_dir")
                if last_dir and os.path.isdir(last_dir):
                    self.src_var.set(last_dir)
        except Exception:
            pass

    def _save_state(self):
        """保存当前源目录。"""
        try:
            STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
            data = {"last_source_dir": self.src_var.get().strip()}
            STATE_FILE.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        except Exception:
            pass

    # ================= 文件选择 =================

    def _browse_src(self):
        init = self.src_var.get().strip() or os.path.expanduser("~")
        d = filedialog.askdirectory(title="选择源目录", initialdir=init)
        if d:
            self.src_var.set(d)
            self._save_state()
            self._auto_detect_trackbook()

    def _auto_detect_trackbook(self):
        """从源目录自动探测跟踪大表。"""
        src = self.src_var.get().strip()
        if not src or not os.path.isdir(src):
            self.track_path = None
            return
        # 优先精确文件名
        default_name = config.DEFAULT_TRACKBOOK_NAME
        candidates = [
            os.path.join(src, default_name),
            os.path.join(src, default_name.replace(".xls", ")")),
        ]
        # 也允许目录中以 "跟踪大表" 结尾的 .xls
        for name in os.listdir(src):
            low = name.lower()
            if low.endswith(".xls") and "跟踪大表" in name:
                candidates.append(os.path.join(src, name))

        for p in candidates:
            if os.path.isfile(p):
                self.track_path = p
                self._log("已自动识别跟踪大表：%s" % p)
                return
        self.track_path = None

    # ================= 执行 =================

    def _start(self):
        if self._running:
            return

        src = self.src_var.get().strip()
        missing = []
        if not src:
            missing.append("源目录")
        elif not os.path.isdir(src):
            missing.append("源目录（路径无效）")

        self._auto_detect_trackbook()
        if not self.track_path or not os.path.exists(self.track_path):
            missing.append("跟踪大表（请在源目录中放置《井下作业设计节点跟踪大表.xls》）")

        if missing:
            messagebox.showerror("缺少路径",
                                 "以下路径无效，请重新选择：\n" +
                                 "\n".join("• " + m for m in missing))
            return

        # 极简模式：固定为「原地写原表 + 零新文件 + 保格式 + 不调整列 + 不允许回退覆盖 + 全部节点」
        cfg = {
            "source": src,
            "trackbook": self.track_path,
            "out_dir": None,
            "apply": True,
            "inplace": True,
            "no_extra_files": True,
            "stages": list(config.STAGE_KEYS),
            "allow_recheck_overwrite": False,
            "restructure": False,
        }

        self._running = True
        self.run_btn.config(bg="#9AA7B4")
        self.run_btn.label.config(text="🔄   写入中...", bg="#9AA7B4")
        self.log_text.delete("1.0", tk.END)
        self.status_label.config(
            text="⏳ 正在按规则写入原表...",
            fg=RUN_FG, bg=RUN_BG)
        self.root.config(cursor="watch")

        self._log("配置：%s" % cfg)

        run_sync(cfg,
                 on_log=self._log,
                 on_done=lambda r: self.root.after(0, self._on_done, r),
                 on_error=lambda e, tb: self.root.after(0, self._on_error, e, tb))

    def _on_done(self, result):
        self._running = False
        self.run_btn.config(bg=ACCENT)
        self.run_btn.label.config(text="🔄   一 键 写 入 原 表", bg=ACCENT)
        self.root.config(cursor="")

        stat = result["stat"]
        lines = [
            "%s ✅  写入 %d / 跳过 %d / 告警 %d（合计 %d）"
            % ("写入完成" if result.get("synced") else "本次无需写入",
               stat["write"], stat["skip"], stat["warn"], stat["total"]),
        ]
        if result.get("synced"):
            lines.append("已直接写入原表：%s" % result["synced"])
            lines.append("（零新文件：未生成任何 .bak / diff / 日志）")
            if result.get("kept_format"):
                lines.append("（已保留合并单元格/字体/边框等原表格式，改动单元格标红）")
            else:
                lines.append("⚠ 未能保留原表格式，已用纯 xlwt 重建，请用 WPS 核对")
        if stat["write"] == 0 and stat["total"] == 0:
            lines.append("未发现可处理的明细文件，请检查源目录与文件命名。")

        # 明细有、但大表无对应行的井（"数据没写进去"的真因）
        unmatched = result.get("unmatched") or {}
        if unmatched:
            total_miss = sum(len(v) for v in unmatched.values())
            stages = "、".join("%s %d 口" % (s, len(v)) for s, v in unmatched.items())
            lines.append("⚠ 明细有 %d 口井在大表找不到对应行、未写入（%s）"
                         % (total_miss, stages))
            lines.append("　大表是按井号匹配的主表，工具只填已有行、不新建行；"
                         "请把这些井补进大表，或确认是否在用正确的大表。")

        self.status_label.config(text="\n".join(lines), fg=OK_FG, bg=OK_BG)

    def _on_error(self, e, tb):
        self._running = False
        self.run_btn.config(bg=ACCENT)
        self.run_btn.label.config(text="🔄   一 键 写 入 原 表", bg=ACCENT)
        self.root.config(cursor="")
        self._log("[错误] %s" % e)
        self._log(tb)
        self.status_label.config(text="❌ 写入失败：%s" % e, fg=ERR_FG, bg=ERR_BG)
        messagebox.showerror("写入失败", "%s\n\n详细日志见上方「执行日志」。" % e)

    # ================= 工具 =================

    def _log(self, msg: str):
        self.root.after(0, self._log_impl, msg)

    def _log_impl(self, msg: str):
        self.log_text.insert(tk.END, str(msg) + "\n")
        self.log_text.see(tk.END)

    def run(self):
        if self.auto:
            self.root.after(600, self._start)
        self.root.mainloop()
