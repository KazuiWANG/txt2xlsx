# -*- coding: utf-8 -*-
"""从任意范例工作簿推断模板：结构、列映射、合并规则与置信度。

推断结果只用于生成候选模板；保存前用户可在界面逐项确认列映射。
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import tempfile
import unicodedata
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from typing import Dict, List, Optional, Tuple

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

from .encoding_utils import read_text
from .models import TEMPLATES_DIR, ColumnSpec, FormulaSpec, Template
from .paths import data_root, user_path
from .text_parser import parse_text
from .workbook_builder import build_workbook

CONFIDENCE_ORDER = {"high": 3, "medium": 2, "low": 1, "none": 0}

_UNIT_TAIL = re.compile(
    r"(mah/g|mah|ma|mv|kg|cm3|cm2|μm|um|mm|ml|mg|g|l|v|a|%|微安|毫安|安时|伏|克|毫克|毫升)$",
    re.IGNORECASE,
)


@dataclass
class Mapping:
    excel: str
    excel_header: str
    txt: Optional[int]
    txt_header: str = ""
    confidence: str = "none"
    note: str = ""


@dataclass
class LearnResult:
    workbook_path: str
    sheet_name: str
    header_rows: int
    data_start: int
    data_end: int
    txt_path: Optional[str]
    txt_header_lines: int
    excel_headers: Dict[str, str]
    txt_headers: List[str]
    column_types: Dict[str, str]
    mappings: List[Mapping] = field(default_factory=list)
    group_columns: List[str] = field(default_factory=list)
    run_columns: List[str] = field(default_factory=list)
    group_key_column: int = 0
    group_fill_columns: List[int] = field(default_factory=list)
    formula_source_row: int = 0
    formulas: Dict[str, str] = field(default_factory=dict)
    warnings: List[str] = field(default_factory=list)
    overall_confidence: str = "low"
    template: Optional[Template] = None
    preview_rows: List[List[str]] = field(default_factory=list)
    preview_note: str = ""

    def mapping_rate(self) -> float:
        mapped = [m for m in self.mappings if m.txt is not None]
        if not mapped:
            return 0.0
        good = sum(1 for m in mapped if m.confidence in ("high", "medium"))
        return float(good) / len(mapped)


class _TrialLogger(object):
    def info(self, *_args, **_kwargs):
        pass

    def warn(self, *_args, **_kwargs):
        pass

    def ok(self, *_args, **_kwargs):
        pass


def normalize_header(text) -> str:
    s = unicodedata.normalize("NFKC", str(text or ""))
    s = s.replace("\r", "").replace("\n", "")
    s = re.sub(r"\s+", "", s)
    s = s.replace("＊", "*")
    return s.strip(":：*")


def split_header(text) -> Tuple[str, Tuple[str, ...]]:
    """拆分“名称 + 单位”，避免把 % 和 mg 当成同一列。"""
    t = normalize_header(text)
    units: List[str] = []
    for inner in re.findall(r"[（(]([^）)]*)[）)]", t):
        units.append(normalize_header(inner))
    base = re.sub(r"[（(][^）)]*[）)]", "", t)
    if "%" in base:
        units.append("%")
        base = base.replace("%", "")
    base = base.replace("％", "")
    changed = True
    while changed:
        changed = False
        match = _UNIT_TAIL.search(base)
        if match and match.start() > 0:
            units.append(match.group(1).lower())
            base = base[:match.start()]
            changed = True
    base = base.strip(" /-_")
    return base, tuple(sorted(set(u for u in units if u)))


def header_score(excel_header: str, txt_header: str) -> Tuple[float, str, str]:
    eb, eu = split_header(excel_header)
    tb, tu = split_header(txt_header)
    if not eb or not tb:
        return 0.0, "none", "缺少表头文字"
    if eb == tb:
        if eu == tu:
            return 1.0, "high", "表头完全一致"
        return 0.8, "medium", "名称一致但单位不同（Excel %s / TXT %s）" % (
            "/".join(eu) or "无", "/".join(tu) or "无")
    ratio = SequenceMatcher(None, eb, tb).ratio()
    if eu and tu and eu != tu and ratio >= 0.55:
        return min(0.66, ratio), "low", "名称相近但单位不同（Excel %s / TXT %s），请确认" % (
            "/".join(eu), "/".join(tu))
    if ratio >= 0.86:
        return ratio, "medium", "名称近似（相似度 %.0f%%）" % (ratio * 100)
    if ratio >= 0.6:
        return ratio, "low", "名称相似度仅 %.0f%%" % (ratio * 100)
    return 0.0, "none", "未匹配"


def _is_number(value) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _to_float(text: str) -> bool:
    try:
        float(str(text).replace(",", "").replace("%", "").strip())
        return True
    except ValueError:
        return False


def _first_content_row(ws) -> int:
    for r in range(1, min(ws.max_row, 20) + 1):
        for c in range(1, ws.max_column + 1):
            if ws.cell(row=r, column=c).value not in (None, ""):
                return r
    return 1


def _detect_structure(ws) -> Tuple[int, int, int, List[str]]:
    warnings: List[str] = []
    first = _first_content_row(ws)
    data_start = None
    for r in range(first + 1, min(ws.max_row, first + 10) + 1):
        values = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        nonempty = [v for v in values if v not in (None, "")]
        numbers = [v for v in nonempty if _is_number(v)]
        if len(nonempty) >= 2 and len(numbers) >= 2:
            data_start = r
            break
    if data_start is None:
        data_start = first + 1
        warnings.append("未找到明显的数值数据行，暂按第 2 行开始为数据区")
    span = max(1, data_start - first)
    if span > 3:
        warnings.append("表头疑似超过 3 行（识别到 %d 行），仅保留前 3 行" % span)
        data_start = first + 3
    header_rows = min(3, span)

    last_col = 0
    for c in range(1, ws.max_column + 1):
        if any(ws.cell(row=r, column=c).value not in (None, "") for r in range(first, min(data_start - 1, first + 3) + 1)):
            last_col = c
        elif any(ws.cell(row=r, column=c).value not in (None, "") for r in range(data_start, min(data_start + 20, ws.max_row + 1))):
            last_col = c

    data_end = data_start - 1
    for r in range(ws.max_row, data_start - 1, -1):
        values = [ws.cell(row=r, column=c).value for c in range(1, max(last_col, 1) + 1)]
        nonempty = [v for v in values if v not in (None, "")]
        if len(nonempty) >= 2 and any(_is_number(v) for v in nonempty):
            data_end = r
            break
    if data_end < data_start:
        data_end = ws.max_row
        warnings.append("数据区没有检测到数值，按到最后一行")
    return header_rows, data_start, data_end, warnings


def _header_texts(ws, header_rows: int, last_col: int) -> Dict[str, str]:
    grid: Dict[Tuple[int, int], str] = {}
    for r in range(1, header_rows + 1):
        for c in range(1, last_col + 1):
            value = ws.cell(row=r, column=c).value
            grid[(r, c)] = "" if value is None else str(value).strip()
    for rng in ws.merged_cells.ranges:
        if rng.min_row > header_rows:
            continue
        value = ws.cell(row=rng.min_row, column=rng.min_col).value
        if value in (None, ""):
            continue
        for r in range(rng.min_row, min(rng.max_row, header_rows) + 1):
            for c in range(rng.min_col, min(rng.max_col, last_col) + 1):
                grid[(r, c)] = str(value).strip()
    headers: Dict[str, str] = {}
    for c in range(1, last_col + 1):
        parts: List[str] = []
        for r in range(1, header_rows + 1):
            value = grid.get((r, c), "")
            if value and value not in parts:
                parts.append(value)
        headers[get_column_letter(c)] = " ".join(parts)
    return headers


def _detect_txt_header_lines(lines: List[str]) -> int:
    count = 1
    for idx in range(1, min(len(lines), 6)):
        cells = [c.strip() for c in lines[idx].split("\t") if c.strip()]
        if len(cells) < 2:
            break
        numbers = sum(1 for c in cells if _to_float(c))
        if numbers >= 2:
            break
        if numbers <= max(1, int(len(cells) * 0.2)):
            count += 1
        else:
            break
    return max(1, count)


def _txt_headers(lines: List[str], header_lines: int) -> List[str]:
    width = 0
    for idx in range(header_lines):
        width = max(width, len(lines[idx].split("\t")))
    headers: List[str] = []
    for c in range(width):
        parts: List[str] = []
        for r in range(header_lines):
            cells = lines[r].split("\t")
            if c >= len(cells):
                continue
            value = cells[c].strip().strip('"').strip()
            if value and value not in parts:
                parts.append(value)
        headers.append(" ".join(parts))
    return headers


def _join_quoted_lines(lines: List[str]) -> List[str]:
    """表头里出现带换行的引号字段时，把物理行还原成一条逻辑行。"""
    merged: List[str] = []
    buffer = ""
    for line in lines:
        buffer = (buffer + "\n" + line) if buffer else line
        if buffer.count('"') % 2 == 0:
            merged.append(buffer)
            buffer = ""
    if buffer:
        merged.append(buffer)
    return merged


def _read_txt_headers(path: str) -> Tuple[List[str], int, List[str]]:
    text, _encoding, _ = read_text(path, ["utf-8-sig", "utf-8", "gb18030", "gbk"])
    lines = [line for line in text.splitlines() if line.strip()]
    if not lines:
        raise ValueError("样例 TXT 为空")
    header_lines = _detect_txt_header_lines(lines)
    logical = _join_quoted_lines(lines[:header_lines])
    return _txt_headers(logical, len(logical)), header_lines, lines


def best_header_score(excel_header: str, txt_header: str) -> Tuple[float, str, str]:
    """多层表头的拼接文本按片段再比一次，避免父级标题拉低子列匹配。"""
    parts = [excel_header] + [p for p in re.split(r"[\s/]+", excel_header) if p]
    best = (0.0, "none", "未匹配")
    for part in parts:
        score, confidence, note = header_score(part, txt_header)
        if CONFIDENCE_ORDER.get(confidence, 0) > CONFIDENCE_ORDER.get(best[1], 0) or (
                confidence == best[1] and score > best[0]):
            best = (score, confidence, note)
    return best


def _assign_mappings(excel_headers: Dict[str, str], txt_headers: List[str]) -> List[Mapping]:
    candidates = []
    for excel, eh in excel_headers.items():
        if not normalize_header(eh):
            candidates.append(Mapping(excel, eh, None, "", "low", "Excel 列为空表头，未自动映射"))
            continue
        for idx, th in enumerate(txt_headers):
            score, confidence, note = best_header_score(eh, th)
            if score > 0:
                candidates.append(Mapping(excel, eh, idx, th, confidence, note))
    candidates.sort(
        key=lambda m: (-CONFIDENCE_ORDER.get(m.confidence, 0),
                       -best_header_score(m.excel_header, m.txt_header)[0],
                       column_index_from_string(m.excel), m.txt or 0))

    used_txt = set()
    assigned: Dict[str, Mapping] = {}
    for item in candidates:
        if item.excel in assigned or item.txt in used_txt:
            continue
        assigned[item.excel] = item
        used_txt.add(item.txt)

    result: List[Mapping] = []
    for excel, eh in excel_headers.items():
        if excel in assigned:
            result.append(assigned[excel])
        elif normalize_header(eh):
            result.append(Mapping(excel, eh, None, "", "low", "未找到可信的 TXT 列，请手工指定"))
        else:
            result.append(Mapping(excel, eh, None, "", "low", "空表头，已跳过"))
    return result


def _positional_mappings(excel_headers: Dict[str, str]) -> List[Mapping]:
    result = []
    index = 0
    for excel, header in excel_headers.items():
        if not normalize_header(header):
            result.append(Mapping(excel, header, None, "", "low", "空表头，已跳过"))
            continue
        result.append(Mapping(excel, header, index, "(按位置)", "low", "未提供样例 TXT，按从左到右顺序占位"))
        index += 1
    return result


def _detect_group_columns(ws, data_start: int, last_col: int, header_rows: int) -> List[str]:
    columns = set()
    for rng in ws.merged_cells.ranges:
        if rng.min_row < data_start or rng.min_col != rng.max_col:
            continue
        if rng.max_row - rng.min_row >= 1:
            columns.add(get_column_letter(rng.min_col))
    return sorted(columns, key=column_index_from_string)


def _detect_run_columns(ws, data_start: int, data_end: int, last_col: int, skip: List[str]) -> List[str]:
    columns = []
    for c in range(1, last_col + 1):
        letter = get_column_letter(c)
        if letter in skip:
            continue
        values = [ws.cell(row=r, column=c).value for r in range(data_start, data_end + 1)]
        values = [v for v in values if v not in (None, "")]
        if len(values) < 6 or any(_is_number(v) for v in values):
            continue
        runs = 0
        current = None
        length = 0
        for value in values:
            if value == current:
                length += 1
            else:
                if length >= 3:
                    runs += 1
                current, length = value, 1
        if length >= 3:
            runs += 1
        if runs >= 2:
            columns.append(letter)
    return columns


def _infer_types(ws, data_start: int, data_end: int, last_col: int) -> Dict[str, str]:
    types = {}
    for c in range(1, last_col + 1):
        letter = get_column_letter(c)
        fmt = str(ws.cell(row=data_start, column=c).number_format or "")
        numeric = any(
            _is_number(ws.cell(row=r, column=c).value)
            for r in range(data_start, min(data_end, data_start + 20) + 1)
        )
        if "%" in fmt:
            types[letter] = "percent"
        elif numeric:
            types[letter] = "number"
        else:
            types[letter] = "text"
    return types


def _scan_formulas(ws, data_start: int, last_col: int) -> Dict[str, str]:
    formulas = {}
    for c in range(1, last_col + 1):
        value = ws.cell(row=data_start, column=c).value
        if isinstance(value, str) and value.startswith("="):
            formulas[get_column_letter(c)] = value
    return formulas


def _signature(txt_headers: List[str]) -> List[str]:
    tokens = []
    for header in txt_headers:
        token = normalize_header(header)
        if len(token) >= 2 and token not in tokens:
            tokens.append(token)
        if len(tokens) >= 3:
            break
    return tokens


def build_candidate(result: LearnResult) -> Template:
    columns = [
        ColumnSpec(excel=m.excel, txt=m.txt, type=result.column_types.get(m.excel, "text"))
        for m in result.mappings if m.txt is not None
    ]
    formulas = [
        FormulaSpec(excel=letter, source_row=result.formula_source_row)
        for letter in sorted(result.formulas, key=column_index_from_string)
    ]
    return Template(
        id="preview",
        display_name="%s（自动学习预览）" % result.sheet_name,
        filename_regex="",
        content_signature=_signature(result.txt_headers) if result.txt_path else [],
        header_lines=result.txt_header_lines,
        group_key_column=result.group_key_column,
        forward_fill_columns=[],
        group_fill_columns=list(result.group_fill_columns),
        sample_workbook=result.workbook_path,
        sample_sheet=result.sheet_name,
        output_workbook=os.path.basename(result.workbook_path),
        output_sheet=result.sheet_name,
        columns=columns,
        formulas=formulas,
        merge_group_columns=list(result.group_columns),
        merge_run_columns=list(result.run_columns),
        outliers={},
        carry_comments=True,
        encoding_candidates=["utf-8-sig", "utf-8", "gb18030", "gbk"],
        signature_scan_lines=max(8, result.txt_header_lines + 2),
        sheet_header_rows=result.header_rows,
    )


def _trial_preview(result: LearnResult) -> None:
    if not result.txt_path or result.template is None:
        result.preview_note = "未提供样例 TXT，无法试转；保存后请手工选一个 TXT 验证。"
        return
    tmp_dir = tempfile.mkdtemp(prefix="txt2xlsx_learn_")
    try:
        doc = parse_text(result.txt_path, result.template)
        out_path = os.path.join(tmp_dir, "preview.xlsx")
        build_workbook(out_path, [doc], _TrialLogger())
        wb = openpyxl.load_workbook(out_path)
        ws = wb[result.sheet_name]
        mapped = [m for m in result.mappings if m.txt is not None]
        for r in range(result.header_rows + 1, min(ws.max_row, result.header_rows + 5) + 1):
            cells = []
            for mapping in mapped:
                cell = ws["%s%d" % (mapping.excel, r)]
                value = cell.value
                if value is None:
                    text = ""
                elif isinstance(value, float):
                    text = ("%.6g" % value)
                else:
                    text = str(value)
                cells.append("%s=%s" % (mapping.excel, text))
            result.preview_rows.append(cells)
        result.preview_note = "试转样例 TXT 的前 %d 行（左为 Excel 列，右为写入值）" % len(result.preview_rows)
    except Exception as exc:
        result.preview_note = "试转失败：%s" % exc
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def learn_template(workbook_path: str, sheet_name: Optional[str] = None,
                   txt_path: Optional[str] = None) -> LearnResult:
    """读取范例工作簿与可选样例 TXT，返回可编辑的候选模板。"""
    workbook_path = os.path.abspath(workbook_path)
    wb = openpyxl.load_workbook(workbook_path, data_only=False)
    sheet_name = sheet_name or wb.sheetnames[0]
    if sheet_name not in wb.sheetnames:
        raise ValueError("工作簿中找不到 sheet：%s" % sheet_name)
    ws = wb[sheet_name]

    header_rows, data_start, data_end, warnings = _detect_structure(ws)
    last_col = 0
    for c in range(1, ws.max_column + 1):
        if any(ws.cell(row=r, column=c).value not in (None, "")
               for r in range(1, min(ws.max_row, data_start + 20) + 1)):
            last_col = c
    last_col = max(last_col, 1)
    excel_headers = _header_texts(ws, header_rows, last_col)
    column_types = _infer_types(ws, data_start, data_end, last_col)
    formulas = _scan_formulas(ws, data_start, last_col)

    txt_headers: List[str] = []
    txt_header_lines = 1
    if txt_path:
        txt_headers, txt_header_lines, _lines = _read_txt_headers(txt_path)
        mappings = _assign_mappings(excel_headers, txt_headers)
    else:
        mappings = _positional_mappings(excel_headers)
        warnings.append("未提供样例 TXT：列映射按从左到右位置占位，置信度全部为 low")

    group_columns = _detect_group_columns(ws, data_start, last_col, header_rows)
    if not group_columns:
        warnings.append("未发现数据区的纵向合并单元格，组字段不会自动合并/填充")
    run_columns = _detect_run_columns(ws, data_start, data_end, last_col, group_columns)

    mapping_by_excel = {m.excel: m for m in mappings}
    group_key = 0
    for letter in group_columns:
        mapping = mapping_by_excel.get(letter)
        if mapping and mapping.txt is not None:
            group_key = mapping.txt
            break
    group_fill = [
        mapping_by_excel[letter].txt for letter in group_columns
        if letter in mapping_by_excel and mapping_by_excel[letter].txt is not None
    ]

    result = LearnResult(
        workbook_path=workbook_path,
        sheet_name=sheet_name,
        header_rows=header_rows,
        data_start=data_start,
        data_end=data_end,
        txt_path=os.path.abspath(txt_path) if txt_path else None,
        txt_header_lines=txt_header_lines,
        excel_headers=excel_headers,
        txt_headers=txt_headers,
        column_types=column_types,
        mappings=mappings,
        group_columns=group_columns,
        run_columns=run_columns,
        group_key_column=group_key,
        group_fill_columns=group_fill,
        formula_source_row=data_start,
        formulas=formulas,
        warnings=warnings,
    )

    mapped = [m for m in mappings if m.txt is not None]
    rate = result.mapping_rate()
    low = sum(1 for m in mapped if m.confidence in ("low", "none"))
    if not mapped:
        result.overall_confidence = "low"
        warnings.append("没有任何列完成映射，必须手工配置")
    elif rate < 0.6:
        result.overall_confidence = "low"
        warnings.append("表头匹配率仅 %.0f%%（低于 60%%），建议逐项手工确认列映射" % (rate * 100))
    elif low or rate < 0.9:
        result.overall_confidence = "medium"
    else:
        result.overall_confidence = "high"

    result.template = build_candidate(result)
    _trial_preview(result)
    return result


def save_learned_template(result: LearnResult, templates_dir: str = TEMPLATES_DIR,
                         assets_dir: Optional[str] = None) -> str:
    """把确认后的候选模板写入 templates/*.json，并把范例工作簿复制进 assets/templates。"""
    key = "%s|%s" % (result.workbook_path, result.sheet_name)
    digest = hashlib.md5(key.encode("utf-8")).hexdigest()[:8]
    template_id = "learned_%s" % digest
    json_path = os.path.join(templates_dir, "%s.json" % template_id)
    assets_dir = assets_dir or user_path("assets", "templates")
    os.makedirs(assets_dir, exist_ok=True)
    asset_name = "learned_%s.xlsx" % digest
    asset_path = os.path.join(assets_dir, asset_name)
    shutil.copy2(result.workbook_path, asset_path)
    try:
        sample_ref = os.path.relpath(asset_path, data_root())
    except ValueError:
        sample_ref = asset_path

    raw = {
        "id": template_id,
        "display_name": "%s（自动学习：%s）" % (
            result.sheet_name, os.path.basename(result.workbook_path)),
        "match": {
            "filename_regex": "",
            "content_signature": _signature(result.txt_headers),
            "scan_lines": max(8, result.txt_header_lines + 2),
        },
        "txt": {
            "header_lines": result.txt_header_lines,
            "group_key_column": result.group_key_column,
            "group_fill_columns": list(result.group_fill_columns),
            "forward_fill_columns": [],
            "encodings": ["utf-8-sig", "utf-8", "gb18030", "gbk"],
        },
        "source": {
            "sample_workbook": sample_ref.replace("\\", "/"),
            "sample_sheet": result.sheet_name,
            "header_rows": result.header_rows,
        },
        "output": {
            "workbook_name": os.path.basename(result.workbook_path),
            "sheet_name": result.sheet_name,
        },
        "columns": [
            {"excel": m.excel, "txt": m.txt, "type": result.column_types.get(m.excel, "text")}
            for m in result.mappings if m.txt is not None
        ],
        "formulas": [
            {"excel": letter, "source_row": result.formula_source_row}
            for letter in sorted(result.formulas, key=column_index_from_string)
        ],
        "merge": {
            "group_columns": list(result.group_columns),
            "run_columns": list(result.run_columns),
        },
        "outliers": {},
        "carry_comments": True,
    }
    os.makedirs(templates_dir, exist_ok=True)
    tmp_path = json_path + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as fh:
        json.dump(raw, fh, ensure_ascii=False, indent=2)
    os.replace(tmp_path, json_path)
    return json_path
