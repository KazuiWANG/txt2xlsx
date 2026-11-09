# -*- coding: utf-8 -*-
"""格式还原验证：逐 sheet 对比范例工作簿与生成结果。

运行示例：
    python tools/verify.py --output output --report docs/验证结果.json
"""
from __future__ import annotations

import argparse
import bisect
import json
import os
import re
import sys
from typing import Dict, List, Optional, Tuple

import openpyxl
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "src"))

from txt2xlsx.models import Template, TemplateRegistry  # noqa: E402
from txt2xlsx.text_parser import parse_text  # noqa: E402
from txt2xlsx.values import convert_value  # noqa: E402
from txt2xlsx.workbook_builder import normalize_key  # noqa: E402


def color_key(color) -> Optional[Tuple]:
    if color is None:
        return None
    return (
        getattr(color, "type", None), getattr(color, "rgb", None),
        getattr(color, "indexed", None), getattr(color, "theme", None),
        round(getattr(color, "tint", 0.0) or 0.0, 4), getattr(color, "auto", None),
    )


def style_key(cell) -> Tuple:
    font = cell.font
    fill = cell.fill
    border = cell.border
    alignment = cell.alignment
    sides = tuple(
        (getattr(side, "style", None), color_key(getattr(side, "color", None)))
        for side in (border.left, border.right, border.top, border.bottom, border.diagonal)
    )
    return (
        (font.name, font.sz, font.b, font.i, font.u, font.strike, font.vertAlign,
         color_key(font.color), font.charset, font.family, font.scheme),
        (fill.patternType, color_key(fill.fgColor), color_key(fill.bgColor)),
        sides + (border.diagonalUp, border.diagonalDown, border.outline),
        (alignment.horizontal, alignment.vertical, alignment.wrap_text,
         alignment.shrink_to_fit, alignment.indent, alignment.relativeIndent,
         alignment.justifyLastLine, alignment.readingOrder, alignment.textRotation),
        cell.number_format,
        (cell.protection.locked, cell.protection.hidden),
    )


def add_check(checks: List[dict], name: str, ok: bool, detail: str,
              differences: Optional[List[str]] = None) -> None:
    checks.append({
        "name": name,
        "ok": bool(ok),
        "detail": detail,
        "differences": (differences or [])[:20],
    })


def page_signature(ws) -> Tuple:
    setup = ws.page_setup
    margins = ws.page_margins
    return (
        setup.orientation, setup.paperSize, setup.scale, setup.fitToWidth, setup.fitToHeight,
        bool(ws.sheet_properties.pageSetUpPr and ws.sheet_properties.pageSetUpPr.fitToPage),
        round(margins.left, 6), round(margins.right, 6), round(margins.top, 6),
        round(margins.bottom, 6), round(margins.header, 6), round(margins.footer, 6),
        ws.print_area, ws.print_title_rows, ws.print_title_cols,
        ws.oddHeader.left.text if ws.oddHeader else None,
        ws.oddHeader.center.text if ws.oddHeader else None,
        ws.oddFooter.left.text if ws.oddFooter else None,
        ws.oddFooter.center.text if ws.oddFooter else None,
        ws.sheet_view.showGridLines, ws.sheet_view.zoomScale,
    )


def freeze_expected(ws) -> Optional[str]:
    pane = ws.sheet_view.pane
    if pane is not None and pane.state in ("frozen", "frozenSplit"):
        xs = int(pane.xSplit or 0)
        ys = int(pane.ySplit or 0)
        if xs or ys:
            return "%s%d" % (get_column_letter(xs + 1), ys + 1)
    return None


def find_sample_txt(template: Template) -> Optional[str]:
    sample_dir = os.path.join(ROOT, "samples")
    for name in sorted(os.listdir(sample_dir)):
        if name.lower().endswith(".txt") and re.search(template.filename_regex, name, re.IGNORECASE):
            return os.path.join(sample_dir, name)
    return None


def group_starts(ws, header_rows: int) -> List[int]:
    starts = []
    for row in range(header_rows + 1, ws.max_row + 1):
        value = ws.cell(row=row, column=1).value
        if value not in (None, ""):
            starts.append(row)
    return starts


def comment_map(ws, header_rows: int) -> Dict[Tuple[str, int, int, int], Tuple[str, str]]:
    starts = group_starts(ws, header_rows)
    data: Dict[Tuple[str, int, int, int], Tuple[str, str]] = {}
    if not starts:
        return data
    occurrence_counter: Dict[str, int] = {}
    occurrences = []
    for row_number in starts:
        key = normalize_key(ws.cell(row=row_number, column=1).value)
        occurrences.append(occurrence_counter.get(key, 0))
        occurrence_counter[key] = occurrences[-1] + 1
    for row in ws.iter_rows(min_row=header_rows + 1, max_row=ws.max_row):
        for cell in row:
            if not cell.comment:
                continue
            pos = bisect.bisect_right(starts, cell.row) - 1
            if pos < 0:
                continue
            key = (normalize_key(ws.cell(row=starts[pos], column=1).value), occurrences[pos],
                   cell.row - starts[pos], cell.column)
            data[key] = (cell.comment.text, cell.comment.author or "")
    return data


def verify_comments(sample_ws, out_ws, header_rows: int, checks: List[dict]) -> None:
    sample = comment_map(sample_ws, header_rows)
    generated = comment_map(out_ws, header_rows)
    generated_occ = {(key[0], key[1]) for key in generated}
    sample_occ = {(key[0], key[1]) for key in sample}
    missing = [key for key in sample if (key[0], key[1]) in generated_occ and key not in generated]
    mismatched = [key for key in sample if key in generated and sample[key] != generated[key]]
    skipped = [key for key in sample if (key[0], key[1]) not in generated_occ]
    extra = [key for key in generated if key not in sample and (key[0], key[1]) in sample_occ]
    detail = ("样例批注 %d 条 / 生成 %d 条；同批号同次序组 %d 个，匹配回填 %d 条；"
              "样例中 %d 条批注所属组不在本次 TXT 中（未回填）"
              % (len(sample), len(generated), len(sample_occ), len(sample) - len(missing) - len(skipped),
                 len(skipped)))
    ok = not missing and not mismatched and not extra
    diffs = []
    for key in missing[:10]:
        diffs.append("缺批注 批号=%s 第%d次出现 组内第%d行 列%s"
                     % (key[0], key[1] + 1, key[2] + 1, get_column_letter(key[3])))
    for key in mismatched[:10]:
        diffs.append("批注内容不一致 批号=%s 第%d次出现 组内第%d行 列%s"
                     % (key[0], key[1] + 1, key[2] + 1, get_column_letter(key[3])))
    add_check(checks, "单元格批注", ok, detail, diffs)


def merged_covered(ws) -> Tuple[set, set]:
    covered = set()
    anchors = set()
    for rng in ws.merged_cells.ranges:
        anchors.add((rng.min_row, rng.min_col))
        for row in range(rng.min_row, rng.max_row + 1):
            for col in range(rng.min_col, rng.max_col + 1):
                covered.add((row, col))
    return covered, anchors


def verify_values(template: Template, out_ws, header_rows: int, doc, checks: List[dict]) -> None:
    covered, anchors = merged_covered(out_ws)
    rows_by_col = {col.excel.upper(): col for col in template.columns if col.txt is not None}
    formula_cols = {spec.excel.upper() for spec in template.formulas}

    checked = 0
    mismatches = []
    filled = 0
    error_cells = 0

    row_number = header_rows + 1
    for group in doc.groups:
        for row in group.rows:
            for letter, spec in rows_by_col.items():
                col = openpyxl.utils.column_index_from_string(letter)
                expected, kind = convert_value(
                    row.cells[spec.txt] if spec.txt < len(row.cells) else "", spec.type)
                cell = out_ws.cell(row=row_number, column=col)
                checked += 1
                if (row_number, col) in covered and (row_number, col) not in anchors:
                    if cell.value is not None:
                        mismatches.append("%s%d 位于合并区但仍有值 %r" % (letter, row_number, cell.value))
                    continue
                if kind == "error":
                    error_cells += 1
                value = cell.value
                if expected is None and letter in formula_cols and isinstance(value, (int, float)):
                    filled += 1
                    continue
                if not _same(value, expected):
                    mismatches.append("%s%d 期望 %r，实际 %r（类型 %s）"
                                      % (letter, row_number, expected, value, cell.data_type))
            row_number += 1

    total_rows = sum(len(group.rows) for group in doc.groups)
    last_row = header_rows + total_rows
    add_check(
        checks, "数值与 TXT 一致性", not mismatches,
        "核对 %d 个映射单元格；异常标记 %d 个；TXT 缺值用公式计算值补位 %d 个；"
        "全部为数值/文本值，无公式；生成末行 %d，TXT 数据 %d 行"
        % (checked, error_cells, filled, out_ws.max_row, total_rows),
        mismatches,
    )
    add_check(checks, "数据行数", out_ws.max_row == last_row,
              "期望末行 %d，实际 %d" % (last_row, out_ws.max_row))


def _same(actual, expected) -> bool:
    if expected is None:
        return actual is None or actual == ""
    if isinstance(expected, (int, float)) and not isinstance(expected, bool):
        return isinstance(actual, (int, float)) and abs(float(actual) - float(expected)) <= 1e-9
    return actual == expected


def verify_merges(template: Template, sample_ws, out_ws, header_rows: int, doc, checks: List[dict]) -> None:
    header_sample = {str(r) for r in sample_ws.merged_cells.ranges if r.min_row <= header_rows}
    header_out = {str(r) for r in out_ws.merged_cells.ranges if r.min_row <= header_rows}
    add_check(checks, "表头合并单元格", header_sample == header_out,
              "样例 %d 处 / 生成 %d 处" % (len(header_sample), len(header_out)),
              sorted(header_sample ^ header_out))

    ranges = {}
    for rng in out_ws.merged_cells.ranges:
        if rng.min_row > header_rows:
            ranges[(rng.min_row, rng.min_col, rng.max_row, rng.max_col)] = str(rng)

    expected_missing = []
    row_number = header_rows + 1
    group_spans = []
    for group in doc.groups:
        start = row_number
        end = start + len(group.rows) - 1
        group_spans.append((start, end))
        for letter in template.merge_group_columns:
            if end <= start:
                continue
            col = openpyxl.utils.column_index_from_string(letter)
            found = None
            for (r1, c1, r2, c2) in ranges:
                if r1 == start and r2 == end and c1 == c2 == col:
                    found = (r1, c1, r2, c2)
                    break
            if not found:
                expected_missing.append("缺少组字段合并 %s%d:%s%d" % (letter, start, letter, end))
        row_number = end + 1

    bad_cross = []
    for (r1, c1, r2, c2) in ranges:
        letter = get_column_letter(c1)
        if c1 != c2:
            bad_cross.append("跨列合并 %s%d:%s%d" % (letter, r1, get_column_letter(c2), r2))
            continue
        inside = any(r1 >= start and r2 <= end for start, end in group_spans)
        if not inside:
            bad_cross.append("跨组或越界合并 %s%d:%s%d" % (letter, r1, letter, r2))
        allowed = letter in template.merge_group_columns or letter in template.merge_run_columns
        if not allowed:
            bad_cross.append("非模板列合并 %s%d:%s%d" % (letter, r1, letter, r2))

    add_check(
        checks, "数据区合并单元格", not expected_missing and not bad_cross,
        "生成数据区合并 %d 处；组字段列 %s 全组覆盖；分段列 %s 按相同值合并"
        % (len(ranges), ",".join(template.merge_group_columns) or "无",
           ",".join(template.merge_run_columns) or "无"),
        expected_missing + bad_cross,
    )


def verify_sheet(template: Template, wb_out, output_dir: str) -> dict:
    checks: List[dict] = []
    sample_wb = openpyxl.load_workbook(template.sample_workbook_path)
    sample_ws = sample_wb[template.sample_sheet]
    out_ws = wb_out[template.output_sheet]
    header_rows = template.sheet_header_rows
    max_col = sample_ws.max_column

    header_diffs = []
    for row in range(1, header_rows + 1):
        for col in range(1, max_col + 1):
            a = sample_ws.cell(row=row, column=col).value
            b = out_ws.cell(row=row, column=col).value
            if a != b:
                header_diffs.append("%s%d 样例 %r / 生成 %r" % (get_column_letter(col), row, a, b))
    add_check(checks, "表头内容", not header_diffs,
              "%d 行 × %d 列" % (header_rows, max_col), header_diffs)

    formula_diffs = []
    scanned = 0
    for row in out_ws.iter_rows():
        for cell in row:
            scanned += 1
            value = cell.value
            if cell.data_type == "f":
                formula_diffs.append("%s data_type=f 公式 %r" % (cell.coordinate, value))
            elif isinstance(value, str) and (value.startswith("=") or value.startswith("'=")):
                formula_diffs.append("%s 文本公式 %r（类型 %s）" % (cell.coordinate, value, cell.data_type))
            elif getattr(cell, "quotePrefix", False):
                formula_diffs.append("%s 被标记为文本（quotePrefix）%r" % (cell.coordinate, value))
    add_check(
        checks, "无公式单元格", not formula_diffs,
        "遍历 %d 个单元格：无 data_type=f、无以 = 或 '= 开头的文本、无文本格式标记" % scanned,
        formula_diffs,
    )

    width_diffs = []
    for col in range(1, max_col + 1):
        letter = get_column_letter(col)
        a_dim = sample_ws.column_dimensions.get(letter)
        b_dim = out_ws.column_dimensions.get(letter)
        a_width = round(a_dim.width, 4) if a_dim and a_dim.width else None
        b_width = round(b_dim.width, 4) if b_dim and b_dim.width else None
        a_hidden = bool(a_dim.hidden) if a_dim else False
        b_hidden = bool(b_dim.hidden) if b_dim else False
        if a_width != b_width or a_hidden != b_hidden:
            width_diffs.append("%s 列 样例宽 %s 隐藏 %s / 生成宽 %s 隐藏 %s"
                               % (letter, a_width, a_hidden, b_width, b_hidden))
    add_check(checks, "列宽与隐藏", not width_diffs, "对比 %d 列" % max_col, width_diffs)

    style_diffs = []
    header_style_diffs = []
    for row in range(1, header_rows + 1):
        for col in range(1, max_col + 1):
            if style_key(sample_ws.cell(row=row, column=col)) != style_key(out_ws.cell(row=row, column=col)):
                header_style_diffs.append("%s%d 表头样式不同" % (get_column_letter(col), row))
    for col in range(1, max_col + 1):
        if style_key(sample_ws.cell(row=header_rows + 1, column=col)) != \
                style_key(out_ws.cell(row=header_rows + 1, column=col)):
            style_diffs.append("%s 列数据样式不同" % get_column_letter(col))
    add_check(checks, "表头样式", not header_style_diffs,
              "字体/填充/边框/对齐/数字格式/保护", header_style_diffs)
    add_check(checks, "数据样式", not style_diffs,
              "按第 %d 行样式逐列对比" % (header_rows + 1), style_diffs)

    base_height = sample_ws.row_dimensions[header_rows + 1].height
    height_diffs = []
    for row in range(header_rows + 1, out_ws.max_row + 1):
        if out_ws.row_dimensions[row].height != base_height:
            height_diffs.append("%d 行 生成 %r / 基准 %r"
                                % (row, out_ws.row_dimensions[row].height, base_height))
    sample_heights = {}
    for row in range(header_rows + 1, sample_ws.max_row + 1):
        h = sample_ws.row_dimensions[row].height
        sample_heights[h] = sample_heights.get(h, 0) + 1
    add_check(
        checks, "数据行高", not height_diffs,
        "生成统一使用基准行高 %s；样例数据区高度分布 %s（样例的 1:1 对应行并不存在，逐行微差见说明）"
        % (base_height, {str(k): v for k, v in sorted(sample_heights.items(), key=lambda x: -x[1])[:4]}),
        height_diffs,
    )

    sample_freeze = sample_ws.freeze_panes
    expected_freeze = freeze_expected(sample_ws)
    add_check(checks, "冻结窗格", out_ws.freeze_panes == expected_freeze,
              "按 xSplit/ySplit 归一化为 %s；样例记录 topLeftCell=%s；生成 %s"
              % (expected_freeze, sample_freeze, out_ws.freeze_panes))
    if sample_freeze != expected_freeze:
        checks[-1]["detail"] += "（样例 topLeftCell 为历史残留，按分屏尺寸还原）"

    sample_page = page_signature(sample_ws)
    out_page = page_signature(out_ws)
    page_diffs = []
    names = ["方向", "纸张", "缩放", "适应宽", "适应高", "fitToPage", "左", "右", "上", "下",
             "页眉", "页脚", "打印区域", "重复行", "重复列", "页眉左", "页眉中", "页脚左", "页脚中",
             "网格线", "缩放比例"]
    for index, name in enumerate(names):
        if sample_page[index] != out_page[index]:
            page_diffs.append("%s: 样例 %r / 生成 %r" % (name, sample_page[index], out_page[index]))
    add_check(checks, "页面设置", not page_diffs, "方向/纸张/缩放/边距/打印区域/页眉页脚/网格线", page_diffs)

    if sample_ws.auto_filter.ref:
        expected_ref = "A1:%s%d" % (get_column_letter(max_col), out_ws.max_row)
        add_check(checks, "筛选范围", out_ws.auto_filter.ref == expected_ref,
                  "样例 %s / 生成 %s（末端按本次数据行数更新）"
                  % (sample_ws.auto_filter.ref, out_ws.auto_filter.ref))
    else:
        add_check(checks, "筛选范围", out_ws.auto_filter.ref is None, "样例无筛选，生成无筛选")

    return {
        "template": template.id,
        "workbook": _relpath(os.path.join(output_dir, template.output_workbook)),
        "sheet": template.output_sheet,
        "header_rows": header_rows,
        "max_column": max_col,
        "checks": checks,
        "sample_comments": sum(1 for row in sample_ws.iter_rows() for c in row if c.comment),
        "output_comments": sum(1 for row in out_ws.iter_rows() for c in row if c.comment),
    }


def _relpath(path: str) -> str:
    """报告里只写相对路径，避免把本机绝对路径带进发布物。"""
    return os.path.relpath(path, ROOT).replace("\\", "/")


def main() -> int:
    parser = argparse.ArgumentParser(description="对比生成工作簿与范例工作簿的格式还原度")
    parser.add_argument("-o", "--output", default=os.path.join(ROOT, "output"))
    parser.add_argument("--report", default=os.path.join(ROOT, "docs", "验证结果.json"))
    args = parser.parse_args()
    output_dir = os.path.abspath(args.output)

    registry = TemplateRegistry()
    workbooks = {}
    results = []
    for template in registry.templates:
        path = os.path.join(output_dir, template.output_workbook)
        if not os.path.exists(path):
            results.append({"template": template.id, "error": "缺少输出文件 %s" % _relpath(path)})
            continue
        if path not in workbooks:
            workbooks[path] = openpyxl.load_workbook(path)
        result = verify_sheet(template, workbooks[path], output_dir)
        txt = find_sample_txt(template)
        if txt:
            doc = parse_text(txt, template)
            verify_comments(openpyxl.load_workbook(template.sample_workbook_path)[template.sample_sheet],
                            workbooks[path][template.output_sheet],
                            template.sheet_header_rows, result["checks"])
            verify_merges(template, openpyxl.load_workbook(template.sample_workbook_path)[template.sample_sheet],
                          workbooks[path][template.output_sheet], template.sheet_header_rows, doc,
                          result["checks"])
            verify_values(template, workbooks[path][template.output_sheet],
                          template.sheet_header_rows, doc, result["checks"])
        results.append(result)

    report = {"output_dir": _relpath(output_dir), "sheets": results}
    os.makedirs(os.path.dirname(os.path.abspath(args.report)), exist_ok=True)
    with open(args.report, "w", encoding="utf-8") as fh:
        json.dump(report, fh, ensure_ascii=False, indent=2)

    print("=" * 78)
    for result in results:
        if "error" in result:
            print("[失败] %s：%s" % (result["template"], result["error"]))
            continue
        print("\n%s / %s" % (result["workbook"], result["sheet"]))
        for check in result["checks"]:
            print("  %s %-12s %s" % ("PASS" if check["ok"] else "FAIL", check["name"], check["detail"]))
            for diff in check["differences"][:4]:
                print("        - %s" % diff)
    failed = sum(1 for result in results if "error" in result)
    failed += sum(1 for result in results if result.get("checks")
                  for check in result["checks"] if not check["ok"])
    print("\n报告：%s" % os.path.abspath(args.report))
    print("结论：%s" % ("全部关键项通过" if failed == 0 else "存在 %d 项差异，请查看报告" % failed))
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
