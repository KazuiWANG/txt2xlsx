# -*- coding: utf-8 -*-
"""程序入口：启动图形界面。"""
import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(ROOT, "src"))

from txt2xlsx.gui import launch  # noqa: E402


if __name__ == "__main__":
    launch()
