# -*- coding: utf-8 -*-
"""主窗口：选源目录（可多选）+ 选节点（自动识别）+ 选跟踪大表 → 一键按规则写入原表。

范式对齐参考项目 gui/simple_window.py：
    - 顶部蓝色横幅
    - 文件选择卡片（浏览按钮 + 输入框）
    - 自绘主按钮（macOS 兼容）
    - 状态卡片（成功绿 / 失败红 / 运行中黄）
    - 执行日志 Text 区

本次改造要点：
    - 源目录支持多选（Listbox + 添加/删除）
    - 节点支持多选（7 个复选框，载入跟踪大表后自动识别可用节点并置灰不可用者）
    - 极简 UI：去掉全部运行选项，点按钮即按规则直接写入原表
      （原地写、保格式、真正零新文件：不生成任何 .bak / diff / 日志），
      合并单元格/字体/边框等原表格式全部保留、改动单元格标红
"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from core import config
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


class SyncWindow:
    """同步工具主窗口。"""

    def __init__(self, auto: bool = False):
        self.auto = auto
        self.root = tk.Tk()
        self.root.title("进度跟踪报表工具")
        self.root.geometry("980x820")
        self.root.minsize(880, 760)
        self.root.configure(bg=BG)

        default_src = str(config.DEFAULT_SOURCE_DIR)
        self.track_var = tk.StringVar(
            value=str(config.DEFAULT_SOURCE_DIR / config.DEFAULT_TRACKBOOK_NAME))

        # 源目录列表（多选）
        self.src_paths = []
        self.src_listbox = None

        # 节点多选
        self.stage_vars = {k: tk.BooleanVar(value=True) for k in config.STAGE_KEYS}
        self.stage_cbs = {}

        # 运行选项已全部移除：点按钮即按规则直接写入原表、零新文件

        self._running = False

        self._build_ui()
        # 启动时若默认跟踪大表存在，自动扫描可用节点
        if os.path.exists(self.track_var.get()):
            self._scan_stages()

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
        tk.Label(header, text="进度跟踪报表工具", bg=ACCENT, fg="white",
                 font=(FONT, 19, "bold")).pack(anchor=tk.W)

        body = ttk.Frame(self.root, padding=(20, 16, 20, 8))
        body.pack(fill=tk.BOTH, expand=True)

        # ---- 源目录（多选）----
        src_box = ttk.LabelFrame(body, text="源目录（可多选，存放按规则命名的明细表）", padding=12)
        src_box.pack(fill=tk.X, pady=(0, 10))
        list_pane = ttk.Frame(src_box)
        list_pane.pack(fill=tk.X)
        self.src_listbox = tk.Listbox(list_pane, height=4, font=(FONT, 10))
        self.src_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        sb_src = ttk.Scrollbar(list_pane, orient=tk.VERTICAL, command=self.src_listbox.yview)
        self.src_listbox.config(yscrollcommand=sb_src.set)
        sb_src.pack(side=tk.LEFT, fill=tk.Y)
        btn_pane = ttk.Frame(list_pane)
        btn_pane.pack(side=tk.LEFT, fill=tk.Y, padx=(8, 0))
        ttk.Button(btn_pane, text="＋ 添加目录", command=self._add_src).pack(fill=tk.X, pady=2)
        ttk.Button(btn_pane, text="－ 删除选中", command=self._remove_src).pack(fill=tk.X, pady=2)
        # 默认源目录预填
        default_src = str(config.DEFAULT_SOURCE_DIR)
        if os.path.isdir(default_src) and default_src not in self.src_paths:
            self.src_paths.append(default_src)
            self.src_listbox.insert(tk.END, default_src)

        # ---- 跟踪大表 ----
        self._row_dir(src_box, "跟踪大表", self.track_var, self._browse_track,
                      "《井下作业设计节点跟踪大表.xls》")

        # ---- 节点多选 ----
        stage_box = ttk.LabelFrame(body, text="同步节点（载入跟踪大表后自动识别，可多选）", padding=14)
        stage_box.pack(fill=tk.X, pady=(0, 10))
        cols = 3
        rowf = None
        for i, key in enumerate(config.STAGE_KEYS):
            if i % cols == 0:
                rowf = ttk.Frame(stage_box)
                rowf.pack(fill=tk.X, pady=2)
            cb = ttk.Checkbutton(rowf, text=key, variable=self.stage_vars[key])
            cb.pack(side=tk.LEFT, padx=12, pady=2)
            self.stage_cbs[key] = cb
        ttk.Button(stage_box, text="重新扫描可用节点",
                   command=self._scan_stages).pack(anchor=tk.W, pady=(6, 0))

        # ---- 主按钮 ----
        self.run_btn = self._accent_button(body, "🔄   一 键 写 入 原 表", self._start)
        self.run_btn.pack(fill=tk.X, pady=(4, 10))

        # ---- 状态卡片 ----
        self.status_label = tk.Label(
            body, text="选择源目录与跟踪大表后，点击「一键写入原表」。",
            bg=PANEL, fg="#555555", anchor=tk.W, justify=tk.LEFT,
            font=(FONT, 10), padx=12, pady=8, relief=tk.GROOVE,
            borderwidth=1, wraplength=900)
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

    def _row_dir(self, master, label, var, cmd, hint):
        row = ttk.Frame(master)
        row.pack(fill=tk.X, pady=5)
        ttk.Label(row, text=label, width=10, font=(FONT, 11, "bold")).pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=var, font=(FONT, 10)).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=8)
        ttk.Button(row, text="浏览...", command=cmd).pack(side=tk.LEFT)
        ttk.Label(master, text="  " + hint, foreground="#888888",
                  font=(FONT, 9)).pack(anchor=tk.W, padx=(12, 0), pady=(0, 3))

    # ================= 文件选择 =================

    def _add_src(self):
        init = self.src_paths[-1] if self.src_paths else os.path.expanduser("~")
        d = filedialog.askdirectory(title="选择源目录", initialdir=init)
        if d and d not in self.src_paths:
            self.src_paths.append(d)
            self.src_listbox.insert(tk.END, d)

    def _remove_src(self):
        sel = self.src_listbox.curselection()
        if not sel:
            return
        i = sel[0]
        self.src_listbox.delete(i)
        self.src_paths.pop(i)

    def _browse_track(self):
        p = filedialog.askopenfilename(
            title="选择跟踪大表",
            initialdir=os.path.dirname(self.track_var.get() or "."),
            filetypes=[("Excel 97-2003", "*.xls"), ("所有文件", "*.*")])
        if p:
            self.track_var.set(p)
            self._scan_stages()

    def _scan_stages(self):
        """载入跟踪大表，识别当前存在的节点，**仅切换可用/禁用状态**——不取消勾选。

        7 个节点默认全部勾选；大表里不存在的节点置灰禁用，但勾选状态保留，
        保证「默认全部同步」的承诺。
        """
        track = self.track_var.get().strip()
        if not track or not os.path.exists(track):
            return
        try:
            from core.trackbook import TrackBook
            tb = TrackBook(track)
            avail = set()
            for ts in tb.sheets:
                avail |= set(ts.available_stages())
        except Exception as e:
            self._log("节点扫描失败（不影响同步，将按全部节点尝试）：%s" % e)
            avail = set(config.STAGE_KEYS)

        for key, var in self.stage_vars.items():
            cb = self.stage_cbs[key]
            if key in avail:
                var.set(True)
                cb.config(state=tk.NORMAL)
            else:
                # 保持勾选（默认全部同步），仅置灰禁用
                cb.config(state=tk.DISABLED)
        self._log("已扫描可用节点：%s" % ("、".join(sorted(avail)) or "（无）"))

    # ================= 执行 =================

    def _start(self):
        if self._running:
            return

        track = self.track_var.get().strip()
        missing = []
        if not self.src_paths:
            missing.append("源目录（请至少添加一个）")
        else:
            for s in self.src_paths:
                if not os.path.isdir(s):
                    missing.append("源目录：%s" % s)
        if not track or not os.path.exists(track):
            missing.append("跟踪大表")
        if missing:
            messagebox.showerror("缺少路径",
                                 "以下路径无效，请重新选择：\n" + "\n".join("• " + m for m in missing))
            return

        stages = [k for k, v in self.stage_vars.items() if v.get()]
        if not stages:
            messagebox.showerror("未选择节点", "请至少勾选一个要同步的节点。")
            return

        # 极简模式：固定为「原地写原表 + 零新文件 + 保格式 + 不调整列 + 不允许回退覆盖」
        cfg = {
            "source": list(self.src_paths),
            "trackbook": track,
            "out_dir": None,
            "apply": True,
            "inplace": True,
            "no_extra_files": True,
            "stages": stages,
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
