# -*- coding: utf-8 -*-
"""制表符分隔 TXT 的通用解析：组识别、组内字段向下填充、结构告警。"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional

from .encoding_utils import read_text
from .models import Template


@dataclass
class ParsedRow:
    line_no: int
    cells: List[str]


@dataclass
class ParsedGroup:
    index: int
    key: str
    start_line: int
    rows: List[ParsedRow] = field(default_factory=list)


@dataclass
class ParsedDocument:
    path: str
    encoding: str
    groups: List[ParsedGroup]
    warnings: List[str]
    blank_lines: int
    total_data_lines: int
    total_rows: int
    template: Optional[Template] = None


class _LimitLog(object):
    def __init__(self, limit: int = 30):
        self.limit = limit
        self.messages: List[str] = []
        self.counts: Dict[str, int] = {}

    def add(self, category: str, message: str) -> None:
        c = self.counts.get(category, 0) + 1
        self.counts[category] = c
        if c <= self.limit:
            self.messages.append(message)

    def finish(self) -> List[str]:
        extra = []
        for cat, n in self.counts.items():
            if n > self.limit:
                extra.append("%s 同类问题共 %d 处，日志仅展示前 %d 处" % (cat, n, self.limit))
        return self.messages + extra


def _is_blank_line(cells: List[str]) -> bool:
    return all((c or "").strip() == "" for c in cells)


def parse_text(path: str, template: Template) -> ParsedDocument:
    text, encoding, _ = read_text(path, template.encoding_candidates)
    lines = text.splitlines()
    log = _LimitLog()
    header = template.header_lines
    key_col = template.group_key_column
    if header < 0:
        header = 0
    if len(lines) < header:
        raise ValueError("文件行数少于模板声明的表头行数(%d)" % header)

    required_width = 0
    for col in template.columns:
        if col.txt is not None:
            required_width = max(required_width, col.txt + 1)
    required_width = max(required_width, key_col + 1)

    groups: List[ParsedGroup] = []
    current: Optional[ParsedGroup] = None
    blank_lines = 0
    total_data = 0
    total_rows = 0
    fill_count = 0
    last_values: Dict[int, str] = {}

    for idx in range(header, len(lines)):
        raw_line = lines[idx]
        line_no = idx + 1
        cells = raw_line.split("\t")
        if _is_blank_line(cells):
            blank_lines += 1
            continue
        total_data += 1

        if len(cells) < required_width:
            log.add(
                "列数不足",
                "第%d行只有 %d 列（模板需要至少 %d 列），缺失列按空值处理" % (line_no, len(cells), required_width),
            )
            cells = cells + [""] * (required_width - len(cells))
        elif len(cells) > required_width:
            log.add(
                "列数超出",
                "第%d行有 %d 列（模板映射到第 %d 列），多余列已忽略" % (line_no, len(cells), required_width - 1),
            )

        key = (cells[key_col] or "").strip() if key_col < len(cells) else ""
        if key:
            current = ParsedGroup(index=len(groups), key=key, start_line=line_no)
            groups.append(current)
            last_values = {}
        elif current is None:
            log.add("缺少组标识", "第%d行没有组标识且无上文组，按空组标识保留" % line_no)
            current = ParsedGroup(index=len(groups), key="", start_line=line_no)
            groups.append(current)
            last_values = {}

        # 组内向下填充：仅填充模板声明的列，避免把逐行变化的列误填。
        for col_idx in template.forward_fill_columns:
            if col_idx >= len(cells):
                continue
            value = (cells[col_idx] or "").strip()
            if value:
                last_values[col_idx] = cells[col_idx]
            elif col_idx in last_values:
                cells[col_idx] = last_values[col_idx]
                fill_count += 1

        current.rows.append(ParsedRow(line_no=line_no, cells=cells))
        total_rows += 1

    if not groups:
        raise ValueError("未解析到任何数据组（请检查表头行数与模板是否匹配）")

    # 组级字段（批号/日期/制作人等）可能在组内任意一行出现：用组内首个非空值补齐整组。
    group_fill_count = 0
    for group in groups:
        for col_idx in template.group_fill_columns:
            present = []
            for row in group.rows:
                if col_idx < len(row.cells):
                    v = (row.cells[col_idx] or "").strip()
                    if v:
                        present.append((row.line_no, row.cells[col_idx]))
            if not present:
                continue
            first = present[0][1]
            distinct = []
            for _, v in present:
                if v not in distinct:
                    distinct.append(v)
            if len(distinct) > 1:
                log.add(
                    "组内字段不一致",
                    "组 %s 第 %d 列出现多个值 %r，取首个值 %r 补齐"
                    % (group.key, col_idx, distinct[:4], first),
                )
            for row in group.rows:
                if col_idx < len(row.cells) and (row.cells[col_idx] or "").strip() == "":
                    row.cells[col_idx] = first
                    group_fill_count += 1

    warnings = log.finish()
    warnings.insert(0, "编码识别: %s；表头 %d 行；数据组 %d 个；数据行 %d 行" % (
        encoding, header, len(groups), total_rows))
    if fill_count:
        warnings.append("组内向下填充字段共 %d 处" % fill_count)
    if group_fill_count:
        warnings.append("组级字段整组补齐共 %d 处" % group_fill_count)
    if blank_lines:
        warnings.append("跳过空行 %d 行" % blank_lines)

    return ParsedDocument(
        path=path,
        encoding=encoding,
        groups=groups,
        warnings=warnings,
        blank_lines=blank_lines,
        total_data_lines=total_data,
        total_rows=total_rows,
        template=template,
    )
