# -*- coding: utf-8 -*-
"""数据模型与模板配置加载。"""
from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from openpyxl.utils import column_index_from_string, get_column_letter


from .paths import PROJECT_ROOT, resolve_data_path, templates_dir  # noqa: E402

TEMPLATES_DIR = templates_dir()


@dataclass
class ColumnSpec:
    excel: str
    txt: Optional[int]
    type: str = "text"


@dataclass
class FormulaSpec:
    excel: str
    source_row: int
    when_batch_regex: Optional[str] = None

    def matches(self, batch: str) -> bool:
        if not self.when_batch_regex:
            return True
        return bool(re.search(self.when_batch_regex, batch or ""))


@dataclass
class Template:
    id: str
    display_name: str
    filename_regex: str
    content_signature: List[str]
    header_lines: int
    group_key_column: int
    forward_fill_columns: List[int]
    group_fill_columns: List[int]
    sample_workbook: str
    sample_sheet: str
    output_workbook: str
    output_sheet: str
    columns: List[ColumnSpec]
    formulas: List[FormulaSpec]
    merge_group_columns: List[str]
    merge_run_columns: List[str]
    outliers: Dict[str, List[float]]
    carry_comments: bool
    encoding_candidates: List[str] = field(default_factory=list)
    signature_scan_lines: int = 8
    source_path: str = ""
    # 范例工作表中表头占用的行数，可能少于 TXT 的表头行数（如 SY6/SYLRM 的 TXT 有 3 行、Excel 只有 2 行）。
    sheet_header_rows: int = 0

    @property
    def sample_workbook_path(self) -> str:
        return resolve_data_path(self.sample_workbook)

    def formula_specs_for(self, excel: str, batch: str) -> List[FormulaSpec]:
        return [f for f in self.formulas if f.excel == excel and f.matches(batch)]

    def choose_formula(self, excel: str, batch: str) -> Optional[FormulaSpec]:
        for spec in self.formula_specs_for(excel, batch):
            return spec
        return None


def _expand_txt_int(value) -> int:
    if isinstance(value, int):
        return value
    return int(str(value).strip())


def load_template(path: str) -> Template:
    with open(path, "r", encoding="utf-8") as f:
        raw = json.load(f)

    source = raw["source"]
    output = raw["output"]
    txt = raw.get("txt", {})
    merge = raw.get("merge", {})
    txt_header_lines = int(txt.get("header_lines", raw.get("header_lines", 0)))
    sheet_header_rows = int(source.get("header_rows", txt_header_lines))

    columns: List[ColumnSpec] = []
    for item in raw.get("columns", []):
        columns.append(
            ColumnSpec(
                excel=str(item["excel"]).strip().upper(),
                txt=_expand_txt_int(item["txt"]) if item.get("txt") is not None else None,
                type=str(item.get("type", "text")).strip().lower(),
            )
        )

    formulas: List[FormulaSpec] = []
    for item in raw.get("formulas", []):
        formulas.append(
            FormulaSpec(
                excel=str(item["excel"]).strip().upper(),
                source_row=int(item["source_row"]),
                when_batch_regex=item.get("when_batch_regex"),
            )
        )

    for rg in raw.get("column_ranges", []):
        c_from = column_index_from_string(str(rg["excel_from"]).strip().upper())
        c_to = column_index_from_string(str(rg["excel_to"]).strip().upper())
        t_from = int(rg["txt_from"])
        col_type = str(rg.get("type", "number")).strip().lower()
        for offset, ci in enumerate(range(c_from, c_to + 1)):
            columns.append(ColumnSpec(get_column_letter(ci), t_from + offset, col_type))

        every = rg.get("formula_columns_every")
        if every:
            start = column_index_from_string(str(every["start"]).strip().upper())
            step = int(every.get("step", 2))
            end = column_index_from_string(str(every["end"]).strip().upper())
            source_row = int(rg.get("formula_source_row") or (sheet_header_rows + 1))
            for ci in range(start, end + 1, step):
                formulas.append(FormulaSpec(get_column_letter(ci), source_row))

    return Template(
        id=str(raw["id"]),
        display_name=str(raw.get("display_name", raw["id"])),
        filename_regex=str(raw.get("match", {}).get("filename_regex", "")),
        content_signature=list(raw.get("match", {}).get("content_signature", [])),
        header_lines=txt_header_lines,
        group_key_column=int(txt.get("group_key_column", 0)),
        forward_fill_columns=[int(x) for x in txt.get("forward_fill_columns", [])],
        group_fill_columns=[int(x) for x in txt.get("group_fill_columns", [])],
        sample_workbook=str(source["sample_workbook"]),
        sample_sheet=str(source["sample_sheet"]),
        output_workbook=str(output["workbook_name"]),
        output_sheet=str(output["sheet_name"]),
        columns=columns,
        formulas=formulas,
        merge_group_columns=[str(x).upper() for x in merge.get("group_columns", [])],
        merge_run_columns=[str(x).upper() for x in merge.get("run_columns", [])],
        outliers={str(k).upper(): [float(v[0]), float(v[1])] for k, v in raw.get("outliers", {}).items()},
        carry_comments=bool(raw.get("carry_comments", True)),
        encoding_candidates=list(txt.get("encodings", ["utf-8-sig", "utf-8", "gb18030"])),
        signature_scan_lines=int(raw.get("match", {}).get("scan_lines", 8)),
        source_path=path,
        sheet_header_rows=sheet_header_rows,
    )


class TemplateRegistry:
    def __init__(self, templates_dir: str = TEMPLATES_DIR):
        self.templates_dir = templates_dir
        self.templates: List[Template] = []
        self.reload()

    def reload(self) -> None:
        self.templates = []
        if not os.path.isdir(self.templates_dir):
            return
        for name in sorted(os.listdir(self.templates_dir)):
            if name.lower().endswith(".json"):
                self.templates.append(load_template(os.path.join(self.templates_dir, name)))

    def by_id(self, template_id: str) -> Optional[Template]:
        for t in self.templates:
            if t.id == template_id:
                return t
        return None

    def match_filename(self, path: str) -> Optional[Template]:
        name = os.path.basename(path)
        for t in self.templates:
            if t.filename_regex and re.search(t.filename_regex, name, re.IGNORECASE):
                return t
        return None

    def match_content(self, preview_text: str) -> Optional[Template]:
        best = None
        best_score = -1
        for t in self.templates:
            if not t.content_signature:
                continue
            score = sum(1 for token in t.content_signature if token in preview_text)
            if score == len(t.content_signature) and score > best_score:
                best = t
                best_score = score
        return best

    def match(self, path: str, preview_text: str = "") -> Optional[Template]:
        t = self.match_filename(path)
        if t:
            return t
        return self.match_content(preview_text)
