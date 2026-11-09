# -*- coding: utf-8 -*-
"""模板公式的受控求值与“公式/TXT 原值”取舍。"""
from __future__ import annotations

import ast
import re
from dataclasses import dataclass
from typing import Dict, Optional, Tuple

_CELL_RE = re.compile(r"(?<![A-Za-z0-9_])(\$?[A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_])")


def _col_letter(ref: str) -> str:
    return ref.replace("$", "").upper()


def _allowed_node(node: ast.AST) -> bool:
    allowed = (
        ast.Expression, ast.BinOp, ast.UnaryOp, ast.Constant,
        ast.Add, ast.Sub, ast.Mult, ast.Div, ast.Pow, ast.Mod,
        ast.USub, ast.UAdd, ast.Load,
    )
    for child in ast.walk(node):
        if not isinstance(child, allowed):
            return False
    return True


def evaluate_expression(expr: str) -> float:
    tree = ast.parse(expr, mode="eval")
    if not _allowed_node(tree):
        raise ValueError("公式包含不支持的语法")
    return float(eval(compile(tree, "<formula>", "eval"), {"__builtins__": {}}, {}))


def evaluate_formula(formula: str, row: int, resolve) -> float:
    """求值只引用同一行的公式。resolve(column_letter) -> 数值。"""
    if not formula.startswith("="):
        raise ValueError("不是公式")
    body = formula[1:]

    def replace(match):
        col = _col_letter(match.group(1))
        ref_row = int(match.group(3))
        if ref_row != row:
            raise ValueError("公式引用了其他行: %s%d" % (col, ref_row))
        value = resolve(col)
        if value is None:
            # 与 Excel 一致：空单元格参与四则运算时按 0 处理。
            return "(0.0)"
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            raise ValueError("引用单元格 %s%d 不是数值" % (col, ref_row))
        return "(%r)" % float(value)

    expr = _CELL_RE.sub(replace, body)
    if re.search(r"[A-Za-z_]", expr):
        raise ValueError("公式仍含无法解析的标识符")
    return evaluate_expression(expr)


def is_close(a: float, b: float, rel: float = 0.005, floor: float = 0.02) -> bool:
    return abs(a - b) <= max(floor, abs(b) * rel)


@dataclass
class CellDecision:
    mode: str                 # "value" / "blank"（不再产生公式）
    value: object = None
    formula: Optional[str] = None
    numeric: Optional[float] = None
    reason: str = ""

    def display(self):
        return self.value


class RowResolver(object):
    """在同一行内决定每个单元格写入的值：TXT 优先，缺值时才用模板公式算出的数值补位。"""

    def __init__(self, row_number: int, values: Dict[str, object], formulas: Dict[str, str], tolerance_rel=0.005):
        self.row_number = row_number
        self.values = values
        self.formulas = formulas
        self.cache: Dict[str, CellDecision] = {}
        self.visiting: set = set()
        self.tolerance_rel = tolerance_rel
        self.deviations = []

    def _resolve_letter(self, col: str) -> object:
        col = col.upper()
        if col in self.cache:
            return self.cache[col].value
        if col in self.visiting:
            raise ValueError("公式存在循环引用: %s" % col)
        self.visiting.add(col)
        try:
            decision = self._decide(col)
            self.cache[col] = decision
            return decision.value
        finally:
            self.visiting.discard(col)

    def _eval_formula(self, formula: str) -> Tuple[Optional[float], str]:
        try:
            return evaluate_formula(formula, self.row_number, self._resolve_letter), ""
        except Exception as exc:
            return None, str(exc)

    def _decide(self, col: str) -> CellDecision:
        txt_value = self.values.get(col)
        formula = self.formulas.get(col)
        if formula is None:
            return CellDecision("value", txt_value)

        # TXT 中是非数值异常标记（如 "/"、"#VALUE!"）时原样保留。
        if isinstance(txt_value, str):
            return CellDecision("value", txt_value, reason="TXT 异常标记原样保留")

        if isinstance(txt_value, (int, float)):
            result, _ = self._eval_formula(formula)
            if result is not None and not is_close(result, float(txt_value), self.tolerance_rel):
                self.deviations.append(
                    (col, formula, txt_value, "公式=%.6g，TXT=%.6g" % (result, float(txt_value))))
            # 无论公式是否一致，都写入 TXT 的确定数值，绝不写公式。
            return CellDecision("value", txt_value, numeric=float(txt_value), reason="写 TXT 数值")

        # TXT 无值：公式能算就用算出的数值补位，算不了留空；不写公式本身。
        result, _ = self._eval_formula(formula)
        if result is not None:
            return CellDecision("value", result, numeric=result, reason="TXT 缺值，公式计算值补位")
        return CellDecision("blank", None, reason="TXT 缺值且公式无法求值")

    def decide(self, col: str) -> CellDecision:
        self._resolve_letter(col)
        return self.cache[col]
