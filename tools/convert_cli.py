# -*- coding: utf-8 -*-
"""命令行批量转换（用于回归测试，GUI 默认入口见 app.py）。"""
import argparse
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "src"))

from txt2xlsx.models import TemplateRegistry  # noqa: E402
from txt2xlsx.pipeline import run_conversion  # noqa: E402


def main():
    parser = argparse.ArgumentParser(description="TXT -> Excel 样式重建")
    parser.add_argument("files", nargs="*", help="TXT 文件；留空则转换 samples/ 下全部")
    parser.add_argument("-o", "--output", default=os.path.join(ROOT, "output"), help="输出目录")
    parser.add_argument("--template", default="", help="强制使用模板 id（可选）")
    args = parser.parse_args()

    files = args.files
    if not files:
        sample_dir = os.path.join(ROOT, "samples")
        files = [os.path.join(sample_dir, n) for n in sorted(os.listdir(sample_dir)) if n.lower().endswith(".txt")]
    files = [os.path.abspath(f) for f in files]

    overrides = {f: args.template for f in files} if args.template else {}
    registry = TemplateRegistry()
    summary = run_conversion(files, overrides, os.path.abspath(args.output), registry)
    print("=" * 70)
    print("成功 %d / 失败 %d" % (summary.succeeded, summary.failed))
    for path in summary.outputs:
        print("输出:", path)
    print("日志:", summary.log_path)
    return 0 if summary.failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
