# -*- coding: utf-8 -*-
"""井下作业设计节点同步系统 — GUI 层（tkinter）。

与参考项目「报表转换工具」保持同一套技术栈与视觉范式：
    - tkinter / ttk，无额外 GUI 依赖
    - 后台线程执行 + on_log/on_done/on_error 回调，主线程用 root.after 刷新
    - macOS 上 tk.Button 忽略 bg，主按钮用 Frame+Label 自绘（同 simple_window.py）
"""

from .main_window import SyncWindow

__all__ = ["SyncWindow"]
