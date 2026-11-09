# -*- coding: utf-8 -*-
"""合成演示样例的端到端转换测试（不依赖任何真实业务数据）。"""
import os
import shutil
import sys
import tempfile
import unittest

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "src"))

from txt2xlsx.models import TemplateRegistry  # noqa: E402
from txt2xlsx.pipeline import run_conversion  # noqa: E402


class ConversionSmokeTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp(prefix="txt2xlsx_convert_", dir=os.path.join(ROOT, "tests"))
        cls.registry = TemplateRegistry()
        samples = os.path.join(ROOT, "samples")
        files = [os.path.join(samples, name) for name in sorted(os.listdir(samples))
                 if name.lower().endswith(".txt")]
        cls.summary = run_conversion(files, {}, cls.tmp, cls.registry)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def test_all_files_succeed(self):
        self.assertEqual(self.summary.total, 2)
        self.assertEqual(self.summary.succeeded, 2)
        self.assertEqual(self.summary.failed, 0)
        self.assertEqual(len(self.summary.outputs), 2)

    def test_sheet_rows_and_start(self):
        cases = [
            ("演示分容数据.xlsx", "演示分容", 3, 9),
            ("演示循环数据.xlsx", "演示循环", 2, 6),
        ]
        for name, sheet, first_row, last_row in cases:
            path = os.path.join(self.tmp, name)
            self.assertTrue(os.path.exists(path), name)
            ws = openpyxl.load_workbook(path)[sheet]
            self.assertEqual(ws.max_row, last_row, sheet)
            self.assertIsNotNone(ws.cell(row=first_row, column=1).value, sheet)
            self.assertIsNotNone(ws.cell(row=1, column=1).value, sheet)

    def test_error_tokens_are_kept_as_text(self):
        path = os.path.join(self.tmp, "演示循环数据.xlsx")
        ws = openpyxl.load_workbook(path)["演示循环"]
        total_errors = 0
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("#"):
                    total_errors += 1
        self.assertGreater(total_errors, 0)

    def test_formula_gap_filled_with_value(self):
        """TXT 缺值的公式列应写入计算出的数值，而不是公式文本。"""
        path = os.path.join(self.tmp, "演示分容数据.xlsx")
        ws = openpyxl.load_workbook(path)["演示分容"]
        self.assertAlmostEqual(ws["I3"].value, 46.2, places=2)
        self.assertAlmostEqual(ws["K3"].value, 2.31, places=2)
        self.assertAlmostEqual(ws["N3"].value, ws["L3"].value * 2, places=2)

    def test_no_formula_cells_in_any_sheet(self):
        """数据区只允许真实值：任何 sheet 都不能出现公式或文本公式。"""
        cases = ["演示分容数据.xlsx", "演示循环数据.xlsx"]
        for name in cases:
            path = os.path.join(self.tmp, name)
            wb = openpyxl.load_workbook(path)
            for ws in wb.worksheets:
                for row in ws.iter_rows():
                    for cell in row:
                        self.assertNotEqual(cell.data_type, "f", "%s!%s" % (ws.title, cell.coordinate))
                        self.assertFalse(getattr(cell, "quotePrefix", False),
                                         "%s!%s 被标记为文本" % (ws.title, cell.coordinate))
                        if isinstance(cell.value, str):
                            self.assertFalse(cell.value.startswith("="),
                                             "%s!%s 文本公式 %r" % (ws.title, cell.coordinate, cell.value))
                            self.assertFalse(cell.value.startswith("'="),
                                             "%s!%s 文本公式 %r" % (ws.title, cell.coordinate, cell.value))

    def test_log_file_written(self):
        self.assertTrue(os.path.exists(self.summary.log_path))
        self.assertGreater(os.path.getsize(self.summary.log_path), 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
