# -*- coding: utf-8 -*-
"""源码与 PyInstaller 冻结环境的统一资源/可写目录寻址。"""
from __future__ import annotations

import os
import shutil
import sys

_SOURCE_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 只读资源根：源码运行时是项目根，冻结运行时是 PyInstaller 的解包目录。
PROJECT_ROOT = getattr(sys, "_MEIPASS", _SOURCE_ROOT)
# 别名，语义更明确。
BUNDLE_ROOT = PROJECT_ROOT
SOURCE_ROOT = _SOURCE_ROOT


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def data_root() -> str:
    """可写数据根：源码模式是项目根；exe 模式是 exe 所在目录。"""
    if is_frozen():
        return os.path.dirname(os.path.abspath(sys.executable))
    return _SOURCE_ROOT


def resource_path(*parts: str) -> str:
    return os.path.join(BUNDLE_ROOT, *parts)


def user_path(*parts: str) -> str:
    return os.path.join(data_root(), *parts)


def templates_dir() -> str:
    return user_path("templates")


def resolve_data_path(path: str) -> str:
    """按“用户目录优先、内置资源兜底”的顺序解析模板引用的相对路径。"""
    if not path:
        return path
    if os.path.isabs(path):
        return path
    for base in (data_root(), BUNDLE_ROOT):
        candidate = os.path.join(base, path)
        if os.path.exists(candidate):
            return candidate
    return os.path.join(BUNDLE_ROOT, path)


def ensure_user_dirs() -> None:
    """首次启动初始化 exe 同级的 templates/，并补齐新增的内置模板。"""
    target = templates_dir()
    os.makedirs(target, exist_ok=True)
    bundled = resource_path("templates")
    if os.path.abspath(bundled) == os.path.abspath(target):
        return
    if not os.path.isdir(bundled):
        return
    for name in sorted(os.listdir(bundled)):
        if not name.endswith(".json"):
            continue
        source = os.path.join(bundled, name)
        destination = os.path.join(target, name)
        if not os.path.exists(destination):
            shutil.copy2(source, destination)
