# -*- coding: utf-8 -*-
"""生成应用图标：macOS 的 assets/app.icns 与 Windows 的 assets/app.ico。

意象：蓝色圆角卡片（与 GUI 主题色 #0078D4 同族）+ 白色表格网格
      + 右下角绿色已写入单元格带对勾 —— 一眼看懂「把数据填进表格」。

用法: python tools/make_icon.py
产物:
  assets/app.icns        macOS（由 iconutil 从 iconset 打包）
  assets/app.ico         Windows（Pillow 直接输出多尺寸）
  assets/app-1024.png    基准图，供文档 / 其它用途复用
  build/app.iconset/     中间产物（各尺寸 PNG）
"""

from __future__ import annotations

import pathlib
import shutil
import subprocess
import sys

from PIL import Image, ImageDraw

ROOT = pathlib.Path(__file__).resolve().parent.parent
ICONSET = ROOT / "build" / "app.iconset"
OUT_ICNS = ROOT / "assets" / "app.icns"
OUT_ICO = ROOT / "assets" / "app.ico"
OUT_PNG = ROOT / "assets" / "app-1024.png"

# Windows .ico 内嵌的尺寸（Pillow 上限 256）
ICO_SIZES = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]

# 主色（与 gui/main_window.py 的 ACCENT 同族）
TOP = (26, 138, 224)     # #1A8AE0
BOTTOM = (0, 103, 184)   # #0067B8
LINE = (255, 255, 255, 242)
FILLED = (22, 163, 74, 255)   # 已写入单元格
TICK = (255, 255, 255, 255)

SIZES = {
    "icon_16x16.png": 16,
    "icon_16x16@2x.png": 32,
    "icon_32x32.png": 32,
    "icon_32x32@2x.png": 64,
    "icon_128x128.png": 128,
    "icon_128x128@2x.png": 256,
    "icon_256x256.png": 256,
    "icon_256x256@2x.png": 512,
    "icon_512x512.png": 512,
    "icon_512x512@2x.png": 1024,
}


def draw(size: int) -> Image.Image:
    """在 size×size 画布上绘制图标。"""
    S = size
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))

    # ---- 渐变背景（逐行填充）----
    bg = Image.new("RGB", (S, S))
    bd = ImageDraw.Draw(bg)
    for y in range(S):
        t = y / max(1, S - 1)
        bd.line([(0, y), (S, y)],
                fill=tuple(int(TOP[i] + (BOTTOM[i] - TOP[i]) * t) for i in range(3)))

    # ---- 圆角遮罩 ----
    mask = Image.new("L", (S, S), 0)
    ImageDraw.Draw(mask).rounded_rectangle(
        [0, 0, S - 1, S - 1], radius=int(S * 0.22), fill=255)
    img.paste(bg, (0, 0), mask)

    d = ImageDraw.Draw(img)

    # ---- 表格（2 列 × 3 行）----
    pad = S * 0.255
    x0, y0, x1, y1 = pad, pad, S - pad, S - pad
    lw = max(1, int(round(S * 0.035)))
    cw, ch = (x1 - x0) / 2, (y1 - y0) / 3

    d.rounded_rectangle([x0, y0, x1, y1],
                        radius=int(S * 0.035), outline=LINE, width=lw)
    d.line([(x0 + cw, y0), (x0 + cw, y1)], fill=LINE, width=lw)
    for i in (1, 2):
        y = y0 + ch * i
        d.line([(x0, y), (x1, y)], fill=LINE, width=lw)

    # ---- 右下角：已写入的单元格 ----
    gx0, gy0 = x0 + cw, y0 + ch * 2
    inset = lw * 0.78
    d.rounded_rectangle([gx0 + inset, gy0 + inset, x1 - inset, y1 - inset],
                        radius=int(S * 0.022), fill=FILLED)

    # ---- 对勾 ----
    cx, cy = (gx0 + x1) / 2, (gy0 + y1) / 2
    r = S * 0.052
    kw = max(1, int(round(lw * 0.95)))
    d.line([(cx - r * 0.92, cy + r * 0.02), (cx - r * 0.20, cy + r * 0.60),
            (cx + r * 0.95, cy - r * 0.70)],
           fill=TICK, width=kw, joint="curve")

    return img


def main() -> int:
    if ICONSET.exists():
        shutil.rmtree(ICONSET)
    ICONSET.mkdir(parents=True, exist_ok=True)
    OUT_ICNS.parent.mkdir(parents=True, exist_ok=True)

    base = draw(1024)
    base.save(OUT_PNG)
    print("基准图: %s" % OUT_PNG)

    # ---- Windows .ico ----
    base.save(OUT_ICO, sizes=ICO_SIZES)
    print("ico:    %s (%.1f KB, %d 个尺寸)"
          % (OUT_ICO, OUT_ICO.stat().st_size / 1024, len(ICO_SIZES)))

    for name, px in SIZES.items():
        (base.resize((px, px), Image.LANCZOS) if px != 1024 else base).save(ICONSET / name)
    print("iconset: %s（%d 个尺寸）" % (ICONSET, len(SIZES)))

    if sys.platform != "darwin":
        print("非 macOS，跳过 iconutil（icns 需在本机生成）")
        return 0

    subprocess.run(["iconutil", "-c", "icns", str(ICONSET),
                    "-o", str(OUT_ICNS)], check=True)
    print("icns: %s (%.1f KB)" % (OUT_ICNS, OUT_ICNS.stat().st_size / 1024))
    return 0


if __name__ == "__main__":
    sys.exit(main())
