# -*- coding: utf-8 -*-
"""自动学习模板的端到端测试：结构识别、列映射、试转数值与无公式约束。"""
import json
import os
import shutil
import sys
import tempfile
import unittest

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "src"))

from txt2xlsx.models import load_template  # noqa: E402
from txt2xlsx.template_learner import (  # noqa: E402
    header_score, learn_template, save_learned_template,
)
from txt2xlsx.text_parser import parse_text  # noqa: E402
from txt2xlsx.workbook_builder import build_workbook  # noqa: E402

FIXTURES = os.path.join(ROOT, "tests", "fixtures")

GROUP_ROWS = [
    ("B001", "2026.1.1", "化成", [
        (3.852, 1250.5, 95.50, "首检"),
        (3.848, 1249.0, 95.20, "正常"),
        (3.845, 1248.2, 94.80, "正常"),
        (3.842, 1247.5, 94.50, "正常"),
    ]),
    ("B002", "2026.1.2", "分容", [
        (3.860, 1260.0, 96.00, "首检"),
        (3.858, 1258.8, 95.70, "正常"),
        (3.855, 1257.4, 95.40, "正常"),
    ]),
    ("B003", "2026.1.3", "老化", [
        (3.870, 1270.2, 96.50, "首检"),
        (3.868, 1269.0, 96.20, "正常"),
    ]),
]


class _SilentLogger(object):
    def info(self, *_args, **_kwargs):
        pass

    def warn(self, *_args, **_kwargs):
        pass


def build_fixture(folder: str):
    """结构与现有四种都不同的两行多层表头 + 纵向合并组块。"""
    os.makedirs(folder, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "新结构"
    ws["A1"], ws["B1"], ws["C1"] = "批号", "测试日期", "工序"
    ws["D1"], ws["G1"] = "电性能", "备注"
    ws["D2"], ws["E2"], ws["F2"] = "电压V", "容量mAh", "保持率%"
    ws["H2"] = "校验值"
    ws.merge_cells("A1:A2")
    ws.merge_cells("B1:B2")
    ws.merge_cells("C1:C2")
    ws.merge_cells("D1:F1")
    ws.merge_cells("G1:G2")
    ws.merge_cells("H1:H2")
    ws["D3"].number_format = "0.000"
    ws["E3"].number_format = "0.0"
    ws["F3"].number_format = "0.00%"

    row = 3
    txt_lines = ["批号\t测试日期\t工序\t电压V\t容量mAh\t保持率%\t备注"]
    for key, date, process, records in GROUP_ROWS:
        start = row
        for idx, (voltage, capacity, retention, note) in enumerate(records):
            ws.cell(row=row, column=1, value=key if idx == 0 else None)
            ws.cell(row=row, column=2, value=date if idx == 0 else None)
            ws.cell(row=row, column=3, value=process if idx == 0 else None)
            ws.cell(row=row, column=4, value=voltage)
            ws.cell(row=row, column=5, value=capacity)
            ws.cell(row=row, column=6, value=retention)
            ws.cell(row=row, column=7, value=note)
            if idx == 0:
                ws.cell(row=row, column=8, value="=D%d*E%d" % (row, row))
            txt_lines.append(
                "%s\t%s\t%s\t%s\t%s\t%s%%\t%s" % (
                    key if idx == 0 else "", date if idx == 0 else "", process if idx == 0 else "",
                    voltage, capacity, retention, note))
            row += 1
        ws.merge_cells(start_row=start, start_column=1, end_row=row - 1, end_column=1)
        ws.merge_cells(start_row=start, start_column=2, end_row=row - 1, end_column=2)
        ws.merge_cells(start_row=start, start_column=3, end_row=row - 1, end_column=3)

    wb.save(os.path.join(folder, "自动学习测试.xlsx"))
    with open(os.path.join(folder, "自动学习测试.txt"), "w", encoding="utf-8-sig") as fh:
        fh.write("\n".join(txt_lines) + "\n")
    return (os.path.join(folder, "自动学习测试.xlsx"),
            os.path.join(folder, "自动学习测试.txt"))


class LearnerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        os.makedirs(FIXTURES, exist_ok=True)
        cls.tmp = tempfile.mkdtemp(prefix="learn_", dir=FIXTURES)
        cls.xlsx, cls.txt = build_fixture(cls.tmp)
        cls.result = learn_template(cls.xlsx, txt_path=cls.txt)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def test_structure_and_mapping(self):
        result = self.result
        self.assertEqual(result.header_rows, 2)
        self.assertEqual(result.data_start, 3)
        self.assertEqual(result.data_end, 11)
        self.assertEqual(result.txt_header_lines, 1)
        self.assertEqual(result.group_columns, ["A", "B", "C"])
        self.assertEqual(result.group_fill_columns, [0, 1, 2])
        self.assertEqual(result.group_key_column, 0)
        expected = {"A": 0, "B": 1, "C": 2, "D": 3, "E": 4, "F": 5, "G": 6}
        actual = {m.excel: m.txt for m in result.mappings}
        for excel, txt in expected.items():
            self.assertEqual(actual.get(excel), txt, excel)
            self.assertEqual(next(m.confidence for m in result.mappings if m.excel == excel), "high")
        self.assertEqual(result.overall_confidence, "high")
        self.assertGreaterEqual(result.mapping_rate(), 0.85)

    def test_trial_conversion_values_and_no_formula(self):
        result = self.result
        doc = parse_text(self.txt, result.template)
        out_dir = tempfile.mkdtemp(prefix="out_", dir=self.tmp)
        out = os.path.join(out_dir, "结果.xlsx")
        build_workbook(out, [doc], _SilentLogger())
        ws = openpyxl.load_workbook(out)["新结构"]
        self.assertAlmostEqual(ws["D3"].value, 3.852)
        self.assertAlmostEqual(ws["E3"].value, 1250.5)
        self.assertAlmostEqual(ws["F3"].value, 0.955)
        self.assertEqual(ws["F3"].number_format, "0.00%")
        self.assertAlmostEqual(ws["D11"].value, 3.868)
        self.assertAlmostEqual(ws["H3"].value, 3.852 * 1250.5)
        self.assertEqual(str(ws["A3"].value), "B001")
        formula_cells = [
            cell.coordinate for row in ws.iter_rows() for cell in row if cell.data_type == "f"
        ]
        self.assertEqual(formula_cells, [])
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    self.assertFalse(cell.value.startswith("="), cell.coordinate)

    def test_saved_json_can_be_reloaded(self):
        result = self.result
        templates = os.path.join(self.tmp, "templates")
        assets = os.path.join(self.tmp, "assets")
        path = save_learned_template(result, templates_dir=templates, assets_dir=assets)
        with open(path, "r", encoding="utf-8") as fh:
            raw = json.load(fh)
        self.assertEqual(raw["txt"]["group_key_column"], 0)
        self.assertEqual(raw["merge"]["group_columns"], ["A", "B", "C"])
        self.assertEqual(len(raw["columns"]), 7)
        loaded = load_template(path)
        self.assertEqual(loaded.sheet_header_rows, 2)
        self.assertEqual(loaded.group_fill_columns, [0, 1, 2])
        self.assertEqual([c.txt for c in loaded.columns],
                         [0, 1, 2, 3, 4, 5, 6])
        self.assertTrue(os.path.exists(os.path.join(assets, os.path.basename(
            raw["source"]["sample_workbook"]))))

    def test_without_txt_is_low_confidence(self):
        result = learn_template(self.xlsx)
        self.assertEqual(result.overall_confidence, "low")
        self.assertTrue(all(m.confidence == "low" for m in result.mappings))
        self.assertTrue(any("未提供样例 TXT" in w for w in result.warnings))
        self.assertEqual([m.txt for m in result.mappings if m.txt is not None],
                         [0, 1, 2, 3, 4, 5, 6])


class HeaderScoreTests(unittest.TestCase):
    def test_unit_conflict_is_never_high(self):
        score, confidence, note = header_score("活性物质含量%", "活性物质量mg")
        self.assertNotEqual(confidence, "high")
        self.assertIn("单位", note)
        self.assertLess(score, 1.0)

    def test_exact_match_is_high(self):
        _score, confidence, _note = header_score("电压V", "电压V")
        self.assertEqual(confidence, "high")
