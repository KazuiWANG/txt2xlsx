# -*- coding: utf-8 -*-
"""以范例工作表为样式源，清理数据区并按模板重建数据。"""
from __future__ import annotations

import bisect
import copy
import os
import re
import unicodedata
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

import openpyxl
from openpyxl.comments import Comment
from openpyxl.formula.translate import Translator
from openpyxl.utils import column_index_from_string, get_column_letter

from .formula_eval import RowResolver
from .models import ColumnSpec, FormulaSpec, Template
from .text_parser import ParsedDocument
from .values import ERROR_TOKENS, convert_value


@dataclass
class CommentRecord:
    rel_row: int
    column: int
    text: str
    author: str
    width: Optional[float] = None
    height: Optional[float] = None


@dataclass
class SheetBlueprint:
    header_rows: int
    max_column: int
    data_style: Dict[int, object]
    data_height: Optional[float]
    formulas: Dict[Tuple[str, int], str]
    comments: Dict[Tuple[str, int], List[CommentRecord]]
    freeze_cell: Optional[str]
    auto_filter_ref: Optional[str]
    warnings: List[str] = field(default_factory=list)


@dataclass
class WriteStats:
    rows: int = 0
    groups: int = 0
    merges: int = 0
    comments: int = 0
    formula_cells: int = 0      # 模板中带公式的单元格数量
    formulas_replaced: int = 0  # 这些格子中直接写 TXT 数值的数量
    formulas_filled: int = 0    # TXT 缺值时用公式计算出的数值补位的数量
    anomalies: int = 0
    invalid_values: int = 0
    outliers: int = 0


def normalize_key(value: str) -> str:
    s = unicodedata.normalize("NFKC", str(value or ""))
    s = re.sub(r"\s+", "", s)
    return s.upper()


def capture_blueprint(ws, template: Template) -> SheetBlueprint:
    header = template.sheet_header_rows
    data_row = header + 1
    max_col = ws.max_column
    styles = {}
    for c in range(1, max_col + 1):
        styles[c] = copy.copy(ws.cell(row=data_row, column=c)._style)

    data_height = ws.row_dimensions[data_row].height
    if data_height is None:
        data_height = ws.sheet_format.defaultRowHeight

    formulas = {}
    warnings = []
    for spec in template.formulas:
        key = (spec.excel, spec.source_row)
        value = ws["%s%d" % (spec.excel, spec.source_row)].value
        if isinstance(value, str) and value.startswith("="):
            formulas[key] = value
        else:
            warnings.append("模板公式缺失: %s%d 不是公式（%r）" % (spec.excel, spec.source_row, value))

    starts = []
    start_values = []
    for r in range(data_row, ws.max_row + 1):
        value = ws.cell(row=r, column=1).value
        if value not in (None, ""):
            starts.append(r)
            start_values.append(str(value))

    comments: Dict[Tuple[str, int], List[CommentRecord]] = {}
    occurrences = []
    occurrence_counter: Dict[str, int] = {}
    for value in start_values:
        key = normalize_key(value)
        occurrences.append(occurrence_counter.get(key, 0))
        occurrence_counter[key] = occurrences[-1] + 1
    for row in ws.iter_rows(min_row=data_row, max_row=ws.max_row):
        for cell in row:
            if not cell.comment:
                continue
            if not starts:
                continue
            pos = bisect.bisect_right(starts, cell.row) - 1
            if pos < 0:
                continue
            batch = normalize_key(start_values[pos])
            occurrence = occurrences[pos]
            rec = CommentRecord(
                rel_row=cell.row - starts[pos],
                column=cell.column,
                text=cell.comment.text,
                author=cell.comment.author,
                width=cell.comment.width,
                height=cell.comment.height,
            )
            comments.setdefault((batch, occurrence), []).append(rec)

    freeze_cell = None
    pane = ws.sheet_view.pane
    if pane is not None and pane.state in ("frozen", "frozenSplit"):
        xs = int(pane.xSplit or 0)
        ys = int(pane.ySplit or 0)
        if xs or ys:
            freeze_cell = "%s%d" % (get_column_letter(xs + 1), ys + 1)

    return SheetBlueprint(
        header_rows=header,
        max_column=max_col,
        data_style=styles,
        data_height=data_height,
        formulas=formulas,
        comments=comments,
        freeze_cell=freeze_cell,
        auto_filter_ref=ws.auto_filter.ref,
        warnings=warnings,
    )


def clean_data_area(ws, header_rows: int) -> None:
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row > header_rows:
            ws.unmerge_cells(str(rng))
    for row in ws.iter_rows(min_row=header_rows + 1):
        for cell in row:
            if cell.comment:
                cell.comment = None
    remaining = ws.max_row - header_rows
    if remaining > 0:
        ws.delete_rows(header_rows + 1, remaining)
    for key in list(ws.row_dimensions.keys()):
        if key > header_rows:
            try:
                del ws.row_dimensions[key]
            except KeyError:
                pass


class _Limiter(object):
    def __init__(self, limit: int = 40):
        self.limit = limit
        self.count = 0
        self.total = 0

    def hit(self) -> bool:
        self.total += 1
        self.count += 1
        return self.count <= self.limit


def _set_cell(cell, style, value) -> None:
    cell._style = copy.copy(style)
    if value is None:
        return
    cell.value = value


def _translate(formula: str, excel: str, source_row: int, target_row: int) -> str:
    origin = "%s%d" % (excel, source_row)
    return Translator(formula, origin=origin).translate_formula("%s%d" % (excel, target_row))


def write_data(ws, doc: ParsedDocument, template: Template, bp: SheetBlueprint, logger) -> WriteStats:
    scope = os.path.basename(doc.path)
    stats = WriteStats()
    formula_columns = {spec.excel.upper() for spec in template.formulas}
    column_specs = {col.excel.upper(): col for col in template.columns}
    for excel, spec in column_specs.items():
        if spec.txt is not None and spec.txt >= 0:
            pass

    invalid_log = _Limiter()
    anomaly_log = _Limiter()
    outlier_log = _Limiter()
    deviation_log = _Limiter()

    next_row = bp.header_rows + 1
    group_geometry = []
    for group in doc.groups:
        start = next_row
        end = start + len(group.rows) - 1
        group_geometry.append((start, end))
        batch = group.key
        for offset, row in enumerate(group.rows):
            target_row = start + offset
            values = {}
            for excel, spec in column_specs.items():
                if spec.txt is None:
                    continue
                raw = row.cells[spec.txt] if spec.txt < len(row.cells) else ""
                value, kind = convert_value(raw, spec.type)
                if kind == "invalid":
                    stats.invalid_values += 1
                    if invalid_log.hit():
                        logger.warn(
                            "第%d行 %s 列值 %r 无法按数值解析，已按文本原样写入" % (row.line_no, excel, raw),
                            scope,
                        )
                elif kind == "error":
                    stats.anomalies += 1
                    if anomaly_log.hit():
                        logger.info(
                            "第%d行 %s 列异常标记 %r 按原值保留" % (row.line_no, excel, raw),
                            scope,
                        )
                if isinstance(value, (int, float)) and excel in template.outliers:
                    lo, hi = template.outliers[excel]
                    if value < lo or value > hi:
                        stats.outliers += 1
                        if outlier_log.hit():
                            logger.warn(
                                "第%d行 %s 列数值 %s 超出参考范围 [%s, %s]，已按原值写入"
                                % (row.line_no, excel, value, lo, hi),
                                scope,
                            )
                values[excel] = value

            formulas = {}
            for excel in formula_columns:
                spec = template.choose_formula(excel, batch)
                if not spec:
                    continue
                formula = bp.formulas.get((excel, spec.source_row))
                if formula:
                    formulas[excel] = _translate(formula, excel, spec.source_row, target_row)

            resolver = RowResolver(target_row, values, formulas)
            decisions = {}
            for excel in formula_columns:
                decisions[excel] = resolver.decide(excel)
                if excel not in formulas:
                    continue
                stats.formula_cells += 1
                if values.get(excel) is None and decisions[excel].value is not None:
                    stats.formulas_filled += 1
                else:
                    stats.formulas_replaced += 1
            for deviation in resolver.deviations:
                if deviation_log.hit():
                    logger.info(
                        "第%d行 %s 列模板公式与 TXT 不一致（%s），采用 TXT 原值"
                        % (row.line_no, deviation[0], deviation[3]),
                        scope,
                    )

            for c in range(1, bp.max_column + 1):
                style = bp.data_style[c]
                excel = get_column_letter(c)
                if excel in decisions:
                    decision = decisions[excel]
                    cell_value = decision.display()
                else:
                    cell_value = values.get(excel)
                cell = ws.cell(row=target_row, column=c)
                _set_cell(cell, style, cell_value)
            ws.row_dimensions[target_row].height = bp.data_height

        next_row = end + 1
        stats.groups += 1

    last_row = next_row - 1
    stats.rows = last_row - bp.header_rows

    # 合并：整组字段列 + 组内相同值的连续段列
    for start, end in group_geometry:
        if end <= start:
            continue
        for letter in template.merge_group_columns:
            ci = column_index_from_string(letter)
            ws.merge_cells(start_row=start, start_column=ci, end_row=end, end_column=ci)
            stats.merges += 1
        for letter in template.merge_run_columns:
            ci = column_index_from_string(letter)
            run_start = start
            run_value = ws.cell(row=start, column=ci).value
            for r in range(start + 1, end + 2):
                value = ws.cell(row=r, column=ci).value if r <= end else object()
                if r <= end and value == run_value:
                    continue
                if run_value not in (None, "") and r - run_start > 1:
                    ws.merge_cells(start_row=run_start, start_column=ci, end_row=r - 1, end_column=ci)
                    stats.merges += 1
                run_start = r
                run_value = value

    # 冻结窗格与筛选范围随新数据更新
    ws.freeze_panes = bp.freeze_cell
    if bp.auto_filter_ref:
        last_col = get_column_letter(bp.max_column)
        ws.auto_filter.ref = "A1:%s%d" % (last_col, max(last_row, bp.header_rows + 1))

    # 批注：仅对范例中存在的同名组按相对行号回填；同名批号按出现次序区分。
    if template.carry_comments and bp.comments:
        group_occurrence: Dict[str, int] = {}
        for group, (start, end) in zip(doc.groups, group_geometry):
            key = normalize_key(group.key)
            occurrence = group_occurrence.get(key, 0)
            group_occurrence[key] = occurrence + 1
            records = bp.comments.get((key, occurrence)) or bp.comments.get((key, 0))
            if not records:
                continue
            for rec in records:
                target_row = start + rec.rel_row
                if target_row > end:
                    continue
                cell = ws.cell(row=target_row, column=rec.column)
                if cell.comment:
                    continue
                comment = Comment(rec.text, rec.author or "")
                if rec.width:
                    comment.width = rec.width
                if rec.height:
                    comment.height = rec.height
                cell.comment = comment
                stats.comments += 1

    for msg in bp.warnings:
        logger.warn(msg, scope)
    return stats


def build_workbook(output_path: str, documents: List[ParsedDocument], logger) -> Dict[str, WriteStats]:
    """documents 必须共享同一个范例工作簿；按模板分别写 sheet 后一次性保存。"""
    templates = [doc.template for doc in documents]
    sample_path = templates[0].sample_workbook_path
    keep = {t.sample_sheet for t in templates}
    wb = openpyxl.load_workbook(sample_path)
    for name in list(wb.sheetnames):
        if name not in keep:
            del wb[name]

    results = {}
    for doc in documents:
        template = doc.template
        ws = wb[template.sample_sheet]
        bp = capture_blueprint(ws, template)
        clean_data_area(ws, bp.header_rows)
        logger.info(
            "样式源 %s / %s：%d 行表头，%d 列，数据行高 %s，模板公式 %d 个，可回填批注组 %d 个"
            % (os.path.basename(sample_path), template.sample_sheet, bp.header_rows, bp.max_column,
               bp.data_height, len(bp.formulas), len(bp.comments)),
            os.path.basename(doc.path),
        )
        results[doc.path] = write_data(ws, doc, template, bp, logger)

    wb.calculation.fullCalcOnLoad = True
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    tmp_path = output_path + ".tmp"
    wb.save(tmp_path)
    os.replace(tmp_path, output_path)
    return results
