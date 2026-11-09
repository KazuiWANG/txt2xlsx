# -*- coding: utf-8 -*-
"""运行日志：同时写入文件与回调（供 GUI 展示）。"""
from __future__ import annotations

import datetime
import os
from typing import Callable, Optional


LEVELS = ("DEBUG", "INFO", "WARN", "ERROR", "OK")


class RunLogger(object):
    def __init__(self, log_path: str, callback: Optional[Callable] = None):
        self.log_path = log_path
        self.callback = callback
        os.makedirs(os.path.dirname(os.path.abspath(log_path)), exist_ok=True)
        self._fh = open(log_path, "a", encoding="utf-8", newline="\n")

    def log(self, level: str, message: str, scope: str = "") -> None:
        level = level.upper()
        stamp = datetime.datetime.now().strftime("%H:%M:%S")
        prefix = "[%s] %s" % (stamp, level)
        if scope:
            prefix += " [%s]" % scope
        line = "%s %s" % (prefix, message)
        self._fh.write(line + "\n")
        self._fh.flush()
        if self.callback:
            self.callback({"type": "log", "level": level, "message": message, "scope": scope, "line": line})
        print(line)

    def info(self, message, scope=""):
        self.log("INFO", message, scope)

    def warn(self, message, scope=""):
        self.log("WARN", message, scope)

    def error(self, message, scope=""):
        self.log("ERROR", message, scope)

    def ok(self, message, scope=""):
        self.log("OK", message, scope)

    def close(self):
        try:
            self._fh.close()
        except Exception:
            pass
