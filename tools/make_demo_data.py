# -*- coding: utf-8 -*-
"""生成完全虚构的演示数据（TXT、样式范例工作簿、模板 JSON）。

用法：
    python tools/make_demo_data.py

所有数据均为虚构，批号统一使用 DEMO- 前缀，不对应任何真实客户、项目或批次。
脚本可重复执行：会覆盖 samples/、assets/templates/、templates/ 下的演示文件。
"""
from __future__ import annotations

import json
import os

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SAMPLES = os.path.join(ROOT, "samples")
ASSETS = os.path.join(ROOT, "assets", "templates")
TEMPLATES = os.path.join(ROOT, "templates")

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill("solid", fgColor="FCE4D6")
DATA_FILL = PatternFill("solid", fgColor="F7F7F7")
HEADER_FONT = Font(name="Microsoft YaHei", size=10, bold=True, color="7F4F24")
DATA_FONT = Font(name="Microsoft YaHei", size=10, color="333333")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _write_txt(path: str, header, rows) -> None:
    lines = list(header) + ["\t".join(row) for row in rows]
    with open(path, "w", encoding="utf-8-sig", newline="\n") as fh:
        fh.write("\n".join(lines) + "\n")


def _style_header(ws, header_rows: int, ncols: int) -> None:
    for row in range(1, header_rows + 1):
        ws.row_dimensions[row].height = 24
        for col in range(1, ncols + 1):
            cell = ws.cell(row=row, column=col)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.border = BORDER
            cell.alignment = CENTER


def _style_data(ws, first_row: int, last_row: int, ncols: int, number_formats=None) -> None:
    number_formats = number_formats or {}
    for row in range(first_row, last_row + 1):
        ws.row_dimensions[row].height = 18
        for col in range(1, ncols + 1):
            cell = ws.cell(row=row, column=col)
            cell.font = DATA_FONT
            cell.fill = DATA_FILL
            cell.border = BORDER
            cell.alignment = CENTER
            fmt = number_formats.get(col)
            if fmt:
                cell.number_format = fmt


def _set_widths(ws, widths) -> None:
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def _page_setup(ws) -> None:
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.sheet_view.showGridLines = False


def _save_workbook(wb, path: str) -> None:
    """先删除旧文件再保存，避免重复运行时被占用或残留句柄影响。"""
    if os.path.exists(path):
        try:
            os.remove(path)
        except OSError:
            pass
    wb.save(path)


def build_batch_workbook() -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = "演示分容"

    headers = [
        ("A", "批号"), ("B", "制备日期"), ("C", "工序"), ("D", "班组"),
        ("E", "设备"), ("F", "备注"), ("G", "序号"), ("H", "极片重量g"),
    ]
    merged_pairs = ["A", "B", "C", "D", "E", "F", "G", "H"]
    for col, title in headers:
        ws["%s1" % col] = title
        ws["%s2" % col] = None
        ws.merge_cells("%s1:%s2" % (col, col))
    ws["I1"] = "电性能"
    ws.merge_cells("I1:N1")
    for col, title in zip(["I", "J", "K", "L", "M", "N"],
                          ["容量mAh", "能量mWh", "内阻mΩ", "差值1", "差值2", "校验值"]):
        ws["%s2" % col] = title
    for col, title in [("O", "电压V"), ("P", "电流mA"), ("Q", "效率%"),
                       ("R", "温度℃"), ("S", "时间min"), ("T", "校验值")]:
        ws["%s1" % col] = title
        ws["%s2" % col] = None
        ws.merge_cells("%s1:%s2" % (col, col))

    # 公式只用于“TXT 缺值补位”，输出一律写数值；必须放在首个数据行（第 3 行）。
    ws["I3"] = "=H3*10"
    ws["K3"] = "=H3*0.5"
    ws["L3"] = "=I3-K3"
    ws["M3"] = "=I3+K3"
    ws["N3"] = "=L3*2"
    for col, value in zip("ABCDEFGH", ["DEMO-A-0001", "2026-01-05", "化成", "甲班",
                                       "设备-01", "演示", 1, 4.62]):
        ws["%s3" % col] = value
    for col, value in zip("OPQRST", [3.92, 12.5, 0.975, 25.0, 90, 4.65]):
        ws["%s3" % col] = value
    ws["H3"].comment = Comment("演示批注：完全虚构的数据，仅用于验证批注回填。", "DemoAuthor")

    sample_rows = [
        ["DEMO-A-0002", "2026-01-06", "分容", "乙班", "设备-02", "演示", 1, 4.55,
         45.5, 40.9, 1.15, 44.35, 46.65, 88.7, 3.88, 11.9, 0.968, 24.8, 95, 4.58],
        ["DEMO-A-0003", "2026-01-07", "老化", "甲班", "设备-01", "无", 1, 4.49,
         44.9, 40.4, 1.18, 43.72, 46.08, 87.44, 3.86, 11.6, 0.961, 25.4, 120, 4.52],
    ]
    for r, values in enumerate(sample_rows, start=4):
        for c, value in enumerate(values, start=1):
            ws.cell(row=r, column=c, value=value)

    _style_header(ws, 2, 20)
    _style_data(ws, 3, 3 + len(sample_rows), 20, {
        8: "0.000", 9: "0.00", 10: "0.00", 11: "0.00", 12: "0.00",
        13: "0.00", 14: "0.00", 15: "0.0", 16: "0.0", 17: "0.00%",
        19: "0", 20: "0.00",
    })
    _set_widths(ws, {"A": 17, "B": 13, "C": 10, "D": 10, "E": 12, "F": 10,
                     "G": 8, "H": 13, "O": 9, "P": 10, "Q": 10, "R": 9,
                     "S": 10, "T": 10})
    for col in "IJKLMN":
        ws.column_dimensions[col].width = 11
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = "A1:T%d" % (3 + len(sample_rows))
    ws["H3"].comment = Comment("演示批注：完全虚构的数据，仅用于验证批注回填。", "DemoAuthor")
    _page_setup(ws)
    wb.properties.creator = "Demo Generator"
    wb.properties.lastModifiedBy = "Demo Generator"
    wb.properties.title = "演示分容范例（虚构数据）"

    path = os.path.join(ASSETS, "demo_batch.xlsx")
    _save_workbook(wb, path)
    return path


def build_cycle_workbook() -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = "演示循环"

    titles = ["批号", "循环开始日期", "存放位置", "工位", "倍率制度", "备注",
              "首次容量mAh", "首次效率%", "序号", "备用1", "备用2"]
    for i, title in enumerate(titles, start=1):
        ws.cell(row=1, column=i, value=title)
    for i in range(47):
        ws.cell(row=1, column=12 + i, value="指标%02d" % (i + 1))

    sample_rows = [
        ["DEMO-CYCLE-0001", "2026-02-01", "A 区", "工位-01", "1C/1C", "演示", 520.0, 94.8, 1],
        ["DEMO-CYCLE-0002", "2026-02-02", "B 区", "工位-02", "2C/1C", "无", 510.5, 94.2, 1],
    ]
    # 第 2 行是首个数据行：公式列写入公式（只用于补位），其余列放示例值。
    for r, values in enumerate(sample_rows, start=2):
        for c, value in enumerate(values, start=1):
            ws.cell(row=r, column=c, value=value)
        for i in range(47):
            col = 12 + i
            if r == 2 and col % 2 == 0:
                ws.cell(row=r, column=col, value="=L2*2")
            else:
                ws.cell(row=r, column=col, value=100 + (r - 2) * 10 + i)
    ws["G2"].comment = Comment("演示批注：完全虚构的数据，仅用于验证批注回填。", "DemoAuthor")

    _style_header(ws, 1, 58)
    _style_data(ws, 2, 1 + len(sample_rows), 58, {12 + i: "0.00" for i in range(47)})
    _set_widths(ws, {"A": 18, "B": 14, "C": 11, "D": 10, "E": 12, "F": 10,
                     "G": 14, "H": 12, "I": 8, "J": 9, "K": 9})
    for i in range(47):
        ws.column_dimensions[get_column_letter(12 + i)].width = 10
    ws.freeze_panes = "G2"
    ws.auto_filter.ref = "A1:BF%d" % (1 + len(sample_rows))
    ws["G2"].comment = Comment("演示批注：完全虚构的数据，仅用于验证批注回填。", "DemoAuthor")
    _page_setup(ws)

    wb.properties.creator = "Demo Generator"
    wb.properties.lastModifiedBy = "Demo Generator"
    wb.properties.title = "演示循环范例（虚构数据）"

    path = os.path.join(ASSETS, "demo_cycle.xlsx")
    _save_workbook(wb, path)
    return path


def write_batch_txt() -> str:
    header = [
        "批号\t制备日期\t工序\t班组\t设备\t备注\t序号\t极片重量g\t容量mAh\t能量mWh",
        "演示数据（全部为虚构，仅用于功能验证）",
        "字段顺序：组信息 + 过程数据 + 校验值",
    ]
    rows = [
        ["DEMO-A-0001", "2026-01-05", "化成", "甲班", "设备-01", "演示", "1", "4.620",
         "", "41.60", "", "", "", "", "3.92", "12.5", "97.50%", "25.0", "90", "4.65"],
        ["", "", "", "", "", "", "2", "4.580", "46.200", "41.50", "1.100", "", "", "",
         "3.91", "12.4", "#VALUE!", "25.2", "91", "4.61"],
        ["", "", "", "", "", "", "3", "4.600", "46.000", "41.40", "1.120", "", "", "",
         "3.90", "12.3", "97.20%", "25.1", "92", "4.63"],
        ["DEMO-A-0002", "2026-01-06", "分容", "乙班", "设备-02", "演示", "1", "4.550",
         "", "40.90", "", "", "", "", "3.88", "11.9", "96.80%", "24.8", "95", "4.58"],
        ["", "", "", "", "", "", "2", "4.510", "45.100", "40.60", "1.150", "", "", "",
         "3.87", "11.8", "96.50%", "24.9", "96", "4.54"],
        ["DEMO-A-0003", "2026-01-07", "老化", "甲班", "设备-01", "无", "1", "4.490",
         "", "40.40", "", "", "", "", "3.86", "11.6", "96.10%", "25.4", "120", "4.52"],
        ["", "", "", "", "", "", "2", "4.470", "44.700", "40.20", "1.180", "", "", "",
         "3.85", "11.5", "95.90%", "25.5", "121", "4.50"],
    ]
    path = os.path.join(SAMPLES, "DEMO-BATCH-0001.txt")
    _write_txt(path, header, rows)
    return path


def write_cycle_txt() -> str:
    header = ["批号\t循环开始日期\t存放位置\t工位\t倍率制度\t备注\t首次容量mAh\t首次效率%\t序号\t"
              "演示数据（全部为虚构，仅用于功能验证）"]

    def metrics(base, row_index, blank_even):
        values = []
        for i in range(47):
            if blank_even and i % 2 == 0 and i != 46:
                values.append("")
            else:
                values.append("%.2f" % (base + i * 3 + row_index))
        return values

    rows = []
    for row_index in range(3):
        first = row_index == 0
        row = [
            "DEMO-CYCLE-0001" if first else "",
            "2026-02-01" if first else "",
            "A 区" if first else "",
            "工位-01" if first else "",
            "1C/1C" if first else "",
            "演示" if first else "",
            "520.00" if first else "",
            "94.80" if first else "",
            str(row_index + 1),
        ] + metrics(100.0, row_index, blank_even=not first)
        rows.append(row)
    rows[1][9 + 8] = "#DIV/0!"
    for row_index in range(2):
        first = row_index == 0
        row = [
            "DEMO-CYCLE-0002" if first else "",
            "2026-02-02" if first else "",
            "B 区" if first else "",
            "工位-02" if first else "",
            "2C/1C" if first else "",
            "无" if first else "",
            "510.50" if first else "",
            "94.20" if first else "",
            str(row_index + 1),
        ] + metrics(90.0, row_index, blank_even=not first)
        rows.append(row)
    rows[-1][9 + 8] = "#VALUE!"
    path = os.path.join(SAMPLES, "DEMO-CYCLE-0002.txt")
    _write_txt(path, header, rows)
    return path


def write_templates() -> None:
    batch = {
        "id": "demo_batch",
        "display_name": "演示分容（虚构数据）",
        "match": {
            "filename_regex": "^DEMO-BATCH.*\\.txt$",
            "content_signature": ["批号", "制备日期", "演示数据"],
            "scan_lines": 8,
        },
        "txt": {
            "header_lines": 3,
            "group_key_column": 0,
            "forward_fill_columns": [],
            "group_fill_columns": [0, 1, 2, 3, 4, 5],
            "encodings": ["utf-8-sig", "utf-8", "gb18030", "gbk"],
        },
        "source": {
            "sample_workbook": "assets/templates/demo_batch.xlsx",
            "sample_sheet": "演示分容",
            "header_rows": 2,
        },
        "output": {"workbook_name": "演示分容数据.xlsx", "sheet_name": "演示分容"},
        "columns": [
            {"excel": "A", "txt": 0, "type": "text"},
            {"excel": "B", "txt": 1, "type": "text"},
            {"excel": "C", "txt": 2, "type": "text"},
            {"excel": "D", "txt": 3, "type": "text"},
            {"excel": "E", "txt": 4, "type": "text"},
            {"excel": "F", "txt": 5, "type": "text"},
            {"excel": "G", "txt": 6, "type": "number"},
            {"excel": "H", "txt": 7, "type": "number"},
            {"excel": "I", "txt": 8, "type": "number"},
            {"excel": "J", "txt": 9, "type": "number"},
            {"excel": "K", "txt": 10, "type": "number"},
            {"excel": "L", "txt": 11, "type": "number"},
            {"excel": "M", "txt": 12, "type": "number"},
            {"excel": "N", "txt": 13, "type": "number"},
            {"excel": "O", "txt": 14, "type": "number"},
            {"excel": "P", "txt": 15, "type": "number"},
            {"excel": "Q", "txt": 16, "type": "percent"},
            {"excel": "R", "txt": 17, "type": "number"},
            {"excel": "S", "txt": 18, "type": "number"},
            {"excel": "T", "txt": 19, "type": "number"},
        ],
        "formulas": [
            {"excel": "I", "source_row": 3},
            {"excel": "K", "source_row": 3},
            {"excel": "L", "source_row": 3},
            {"excel": "M", "source_row": 3},
            {"excel": "N", "source_row": 3},
        ],
        "merge": {"group_columns": ["A", "B", "C", "D", "E", "F"], "run_columns": []},
        "outliers": {},
        "carry_comments": True,
    }

    cycle = {
        "id": "demo_cycle",
        "display_name": "演示循环（虚构数据）",
        "match": {
            "filename_regex": "^DEMO-CYCLE.*\\.txt$",
            "content_signature": ["批号", "循环开始日期", "演示数据"],
            "scan_lines": 8,
        },
        "txt": {
            "header_lines": 1,
            "group_key_column": 0,
            "forward_fill_columns": [],
            "group_fill_columns": [0, 1, 4],
            "encodings": ["utf-8-sig", "utf-8", "gb18030", "gbk"],
        },
        "source": {
            "sample_workbook": "assets/templates/demo_cycle.xlsx",
            "sample_sheet": "演示循环",
            "header_rows": 1,
        },
        "output": {"workbook_name": "演示循环数据.xlsx", "sheet_name": "演示循环"},
        "columns": [
            {"excel": "A", "txt": 0, "type": "text"},
            {"excel": "B", "txt": 1, "type": "text"},
            {"excel": "C", "txt": 2, "type": "text"},
            {"excel": "D", "txt": 3, "type": "text"},
            {"excel": "E", "txt": 4, "type": "text"},
            {"excel": "F", "txt": 5, "type": "text"},
            {"excel": "G", "txt": 6, "type": "number"},
            {"excel": "H", "txt": 7, "type": "number"},
            {"excel": "I", "txt": 8, "type": "number"},
        ],
        "column_ranges": [
            {
                "excel_from": "L", "excel_to": "BF", "txt_from": 9, "type": "number",
                "formula_columns_every": {"start": "N", "end": "BF", "step": 2},
                "formula_source_row": 2,
            },
        ],
        "formulas": [],
        "merge": {"group_columns": ["A", "B"], "run_columns": ["E"]},
        "outliers": {},
        "carry_comments": True,
    }

    for name, data in (("demo_batch.json", batch), ("demo_cycle.json", cycle)):
        with open(os.path.join(TEMPLATES, name), "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False, indent=2)
            fh.write("\n")


def main() -> int:
    for folder in (SAMPLES, ASSETS, TEMPLATES):
        os.makedirs(folder, exist_ok=True)
    build_batch_workbook()
    build_cycle_workbook()
    write_batch_txt()
    write_cycle_txt()
    write_templates()
    print("已生成演示数据：")
    print("  samples/DEMO-BATCH-0001.txt")
    print("  samples/DEMO-CYCLE-0002.txt")
    print("  assets/templates/demo_batch.xlsx")
    print("  assets/templates/demo_cycle.xlsx")
    print("  templates/demo_batch.json")
    print("  templates/demo_cycle.json")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
