# -*- coding: utf-8 -*-
"""文本编码识别。"""
import os
from typing import List, Tuple


DEFAULT_ENCODINGS = ["utf-8-sig", "utf-8", "gb18030", "gbk"]


def read_text(path: str, encodings: List[str] = None) -> Tuple[str, str, List[str]]:
    """读取文本文件，返回 (内容, 实际编码, 尝试过的编码列表)。

    按 BOM / UTF-8 / GB18030 顺序尝试；GB18030 能兼容绝大多数 ANSI(GBK) 文件。
    """
    candidates = list(encodings or DEFAULT_ENCODINGS)
    with open(path, "rb") as f:
        raw = f.read()
    tried = []
    for enc in candidates:
        tried.append(enc)
        try:
            text = raw.decode(enc)
            return text, enc, tried
        except (UnicodeDecodeError, LookupError):
            continue
    raise UnicodeDecodeError("auto", bytes(raw[:64]), 0, 0, "无法识别文本编码: %s" % os.path.basename(path))


def preview(path: str, lines: int = 10, encodings: List[str] = None) -> str:
    text, _, _ = read_text(path, encodings)
    return "\n".join(text.splitlines()[:lines])
