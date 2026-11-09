# -*- coding: utf-8 -*-
"""单元格原值转换规则。"""
from __future__ import annotations

from typing import Optional, Tuple


NULL_TOKENS = {"", "/", "-", "—", "–", "N/A", "NA", "None", "null"}
ERROR_TOKENS = {"#VALUE!", "#DIV/0!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#NULL!", "#GETTING_DATA"}


def clean_raw(raw: str) -> str:
    if raw is None:
        return ""
    return str(raw).replace("\u3000", " ").strip()


def convert_value(raw: str, value_type: str) -> Tuple[Optional[object], str]:
    """返回 (写入 Excel 的值, 类型标记)。

    类型标记: empty / text / number / error / invalid
    - 异常标记（如 #VALUE!、/）不丢弃，按字符串原样保留；
    - 无法解析为数值的文本也按原样保留，并标记 invalid，由调用方记录日志。
    """
    value_type = (value_type or "text").lower()
    s = clean_raw(raw)

    if s in NULL_TOKENS:
        return None, "empty"

    if value_type in ("text", "date"):
        return s, "text"

    if s in ERROR_TOKENS:
        return s, "error"

    numeric_text = s.replace(",", "").replace("，", "").replace(" ", "")
    negative_parens = False
    if numeric_text.startswith("(") and numeric_text.endswith(")"):
        negative_parens = True
        numeric_text = numeric_text[1:-1]
    percent = False
    if numeric_text.endswith("%") or numeric_text.endswith("％"):
        percent = True
        numeric_text = numeric_text[:-1]

    try:
        number = float(numeric_text)
    except ValueError:
        return s, "invalid"

    if negative_parens:
        number = -number

    if value_type == "percent":
        number = number / 100.0 if percent else number
        return number, "number"

    if value_type in ("number_or_percent", "auto"):
        if percent:
            number = number / 100.0
        return number, "number"

    if value_type == "number":
        if percent:
            number = number / 100.0
        return number, "number"

    return s, "text"


def is_error_like(value) -> bool:
    return isinstance(value, str) and (value in ERROR_TOKENS or value.strip() == "/")
