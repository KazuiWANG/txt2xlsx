# -*- coding: utf-8 -*-
"""核心解析/转换/模板匹配的单元测试（标准库 unittest，无额外依赖）。"""
import json
import os
import shutil
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "src"))

from txt2xlsx.encoding_utils import read_text  # noqa: E402
from txt2xlsx.formula_eval import RowResolver, evaluate_formula, is_close  # noqa: E402
from txt2xlsx.models import TemplateRegistry, load_template  # noqa: E402
from txt2xlsx.text_parser import parse_text  # noqa: E402
from txt2xlsx.values import convert_value  # noqa: E402


def _test_template(tmp_dir, **txt_options):
    raw = {
        "id": "unit",
        "display_name": "单元测试模板",
        "match": {"filename_regex": r"unit_.*\.txt"},
        "txt": {
            "header_lines": 1,
            "group_key_column": 0,
            "forward_fill_columns": [1, 3],
            "group_fill_columns": [0, 2],
            "encodings": ["utf-8-sig", "utf-8", "gb18030"],
        },
        "source": {
            "sample_workbook": "assets/templates/demo_cycle.xlsx",
            "sample_sheet": "演示循环",
            "header_rows": 1,
        },
        "output": {"workbook_name": "unit.xlsx", "sheet_name": "演示循环"},
        "columns": [
            {"excel": "A", "txt": 0, "type": "text"},
            {"excel": "B", "txt": 1, "type": "text"},
            {"excel": "C", "txt": 2, "type": "text"},
            {"excel": "D", "txt": 3, "type": "text"},
        ],
    }
    raw["txt"].update(txt_options)
    path = os.path.join(tmp_dir, "unit_template.json")
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(raw, fh, ensure_ascii=False)
    return load_template(path)


class ParseTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="txt2xlsx_test_", dir=os.path.join(ROOT, "tests"))
        self.template = _test_template(self.tmp)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _write(self, name, text, encoding):
        path = os.path.join(self.tmp, name)
        with open(path, "wb") as fh:
            fh.write(text.encode(encoding))
        return path

    def test_utf8_bom_and_gbk_auto_detect(self):
        body = "批次\t日期\t制作人\t备注\n甲\t2026/1/1\t王\t好\n"
        for encoding in ("utf-8-sig", "gb18030"):
            path = self._write("unit_%s.txt" % encoding.replace("-", ""), body, encoding)
            text, detected, tried = read_text(path, self.template.encoding_candidates)
            self.assertIn("甲", text)
            self.assertTrue(any(code in detected.lower() for code in ("utf", "gb")), (detected, tried))

    def test_group_fill_and_forward_fill(self):
        body = "\n".join([
            "批次\t日期\t制作人\t备注",
            "G1\t2026/1/1\t\t",
            "\t\t王\t",
            "G2\t2026/1/2\t李\t",
            "\t\t\t",
            "",
        ])
        path = self._write("unit_group.txt", body, "utf-8-sig")
        doc = parse_text(path, self.template)
        self.assertEqual(len(doc.groups), 2)
        self.assertEqual(doc.total_rows, 3)
        self.assertEqual(doc.blank_lines, 1)
        g1, g2 = doc.groups
        self.assertEqual([row.cells[0] for row in g1.rows], ["G1", "G1"])
        self.assertEqual([row.cells[1] for row in g1.rows], ["2026/1/1", "2026/1/1"])
        self.assertEqual([row.cells[3] for row in g1.rows], ["", ""])
        self.assertEqual([row.cells[2] for row in g1.rows], ["王", "王"])
        self.assertEqual([row.cells[2] for row in g2.rows], ["李"])
        self.assertTrue(any("组级字段整组补齐" in item for item in doc.warnings))

    def test_short_rows_are_padded_and_reported(self):
        body = "批次\t日期\t制作人\t备注\nG1\t2026/1/1\n"
        path = self._write("unit_short.txt", body, "utf-8-sig")
        doc = parse_text(path, self.template)
        self.assertEqual(len(doc.groups[0].rows[0].cells), 4)
        self.assertTrue(any("缺失列按空值处理" in item for item in doc.warnings))


class ValueTests(unittest.TestCase):
    def test_percent_error_and_invalid(self):
        self.assertEqual(convert_value("97.21%", "percent")[0], 0.9721)
        self.assertEqual(convert_value("97.21%", "number")[0], 0.9721)
        self.assertEqual(convert_value("#VALUE!", "number"), ("#VALUE!", "error"))
        self.assertEqual(convert_value("/", "number"), (None, "empty"))
        self.assertEqual(convert_value("(549.83)", "number")[0], -549.83)
        self.assertEqual(convert_value("1,234.50", "number")[0], 1234.5)
        self.assertEqual(convert_value("abc", "number"), ("abc", "invalid"))
        self.assertEqual(convert_value("文本", "text"), ("文本", "text"))


class FormulaTests(unittest.TestCase):
    def test_same_row_arithmetic_and_blank_as_zero(self):
        values = {"A": 4.6, "B": None}
        result = evaluate_formula("=(A1-B1)*2", 1, lambda col: values[col])
        self.assertAlmostEqual(result, 9.2, places=6)

    def test_other_row_is_rejected(self):
        with self.assertRaises(ValueError):
            evaluate_formula("=A2+1", 1, lambda col: 1)

    def test_tolerance(self):
        self.assertTrue(is_close(99.3, 99.0))
        self.assertFalse(is_close(1.0, 2.0))

    def test_resolver_only_returns_values(self):
        """模板公式只用于算数补位，绝不作为公式文本返回。"""
        values = {"A": 4.6, "B": None, "C": 10.0}
        formulas = {"C": "=A1*2", "D": "=C1+A1"}
        resolver = RowResolver(1, values, formulas)
        c = resolver.decide("C")
        self.assertEqual(c.mode, "value")
        self.assertAlmostEqual(c.display(), 10.0)
        d = resolver.decide("D")
        self.assertEqual(d.mode, "value")
        self.assertAlmostEqual(d.display(), 14.6)
        self.assertFalse(str(c.display()).startswith("="))
        self.assertFalse(str(d.display()).startswith("="))


class TemplateMatchTests(unittest.TestCase):
    def setUp(self):
        self.registry = TemplateRegistry()

    def test_all_templates_load(self):
        self.assertGreaterEqual(len(self.registry.templates), 2)
        ids = {t.id for t in self.registry.templates}
        self.assertTrue({"demo_batch", "demo_cycle"} <= ids)

    def test_filename_matching(self):
        cases = {
            "DEMO-BATCH-0001.txt": "demo_batch",
            "DEMO-CYCLE-0002.txt": "demo_cycle",
        }
        for name, expected in cases.items():
            if expected not in {t.id for t in self.registry.templates}:
                continue
            template = self.registry.match(os.path.join(ROOT, "samples", name), "")
            self.assertIsNotNone(template, name)
            self.assertEqual(template.id, expected, name)

    def test_manual_override_by_id(self):
        template = self.registry.by_id("demo_cycle")
        self.assertEqual(template.output_sheet, "演示循环")


if __name__ == "__main__":
    unittest.main(verbosity=2)
