# -*- coding: utf-8 -*-
"""跨平台路径与文件名工具。

Windows 与 macOS 的三处关键差异，统一在本模块收敛：

1. 路径分隔符：一律用 pathlib，禁止硬编码 ``\\`` 或 ``/``。
2. Unicode 归一化：macOS (HFS+/APFS) 倾向 NFD，Windows (NTFS) 用 NFC。
   同一份中文文件名在两端可能字节不同，比较前必须 NFC 归一。
3. 大小写：两端文件系统默认都不敏感但保留大小写，比较用 casefold。

参考项目：报表转换工具/core/excel_reader.py 的文件读取抽象。
"""

from __future__ import annotations

import os
import pathlib
import platform
import re
import unicodedata

# Windows 文件名非法字符
_ILLEGAL_CHARS = set('\\/:*?"<>|\0')

# 全角/各种短横线破折号 → 半角 "-"
_DASH_CLASS = re.compile(r"[\u2010-\u2015\u2212\uFF0D\u30FC]")


def nfc(s) -> str:
    """Unicode NFC 归一化 + 去首尾空白。None 安全。"""
    if s is None:
        return ""
    return unicodedata.normalize("NFC", str(s)).strip()


def canon(s) -> str:
    """井号/关键字的比较键：NFC + 去全角空格 + 短横线归一 + casefold。"""
    s = nfc(s).replace("\u3000", "").replace("\xa0", "")
    s = _DASH_CLASS.sub("-", s)
    return s.casefold()


# 表头里的干扰字符：换行、全角空格、NBSP
_HEADER_NOISE = re.compile(r"[\r\n\t\u3000\xa0]+")


def norm_header(s) -> str:
    """表头文本归一化：NFC + 剔除换行/全角空格/NBSP + 去首尾空白。

    跟踪大表的一级表头含换行（"修前工程\\n及概算完成情况"、"修井作业\\n开工时间"），
    必须与配置里的连续文本能比对上。
    """
    return _HEADER_NOISE.sub("", nfc(s)).strip()


def safe_path(p) -> pathlib.Path:
    """把任意输入解析为绝对路径并校验存在性。"""
    path = pathlib.Path(os.path.expanduser(str(p))).resolve()
    if not path.exists():
        raise FileNotFoundError("路径不存在: %s" % path)
    return path


def safe_filename(name: str) -> str:
    """校验文件名合法性（Windows 非法字符）。返回 NFC 归一后的名字。"""
    name = nfc(name)
    bad = [c for c in name if c in _ILLEGAL_CHARS]
    if bad:
        raise ValueError("文件名含非法字符 %r: %s" % ("".join(sorted(set(bad))), name))
    return name


def is_win() -> bool:
    return platform.system() == "Windows"


def is_mac() -> bool:
    return platform.system() == "Darwin"


def open_in_explorer(path) -> None:
    """跨平台在文件管理器中打开目录/文件（对齐参考项目 gui/simple_window.py）。"""
    path = pathlib.Path(path)
    target = path if path.is_dir() else path.parent
    target = str(target)
    if is_mac():
        os.system('open "%s"' % target)
    elif is_win():
        os.startfile(target)  # noqa: S606 - Windows 专用
    else:
        os.system('xdg-open "%s"' % target)


def long_path_warning(path, limit: int = 240) -> str | None:
    """Windows 长路径（>260）预警。返回提示文本或 None。"""
    if not is_win():
        return None
    if len(str(pathlib.Path(path))) > limit:
        return ("路径长度 %d 接近 Windows 260 字符上限，建议缩短目录层级：%s"
                % (len(str(pathlib.Path(path))), path))
    return None
