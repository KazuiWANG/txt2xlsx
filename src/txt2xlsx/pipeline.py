# -*- coding: utf-8 -*-
"""转换流水线：匹配模板、解析、按输出工作簿分组写盘。"""
from __future__ import annotations

import datetime
import os
from collections import OrderedDict
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional

from .encoding_utils import preview
from .logging_utils import RunLogger
from .models import PROJECT_ROOT, Template, TemplateRegistry
from .text_parser import ParsedDocument, parse_text
from .workbook_builder import WriteStats, build_workbook


@dataclass
class FileJob:
    path: str
    template: Optional[Template] = None
    document: Optional[ParsedDocument] = None
    status: str = "待处理"
    detail: str = ""
    output: str = ""
    stats: Optional[WriteStats] = None


@dataclass
class RunSummary:
    total: int = 0
    succeeded: int = 0
    failed: int = 0
    outputs: List[str] = field(default_factory=list)
    log_path: str = ""


def _status_cb(cb, job: FileJob, status: str, detail: str = "", progress: float = 0.0):
    job.status = status
    job.detail = detail
    if cb:
        cb({
            "type": "file",
            "path": job.path,
            "status": status,
            "detail": detail,
            "template": job.template.display_name if job.template else "未匹配",
            "output": job.output,
            "progress": progress,
        })


def run_conversion(
    files: List[str],
    overrides: Dict[str, str],
    output_dir: str,
    registry: TemplateRegistry,
    logger_callback: Optional[Callable] = None,
    status_callback: Optional[Callable] = None,
    progress_callback: Optional[Callable] = None,
) -> RunSummary:
    os.makedirs(output_dir, exist_ok=True)
    log_dir = os.path.join(output_dir, "logs")
    os.makedirs(log_dir, exist_ok=True)
    log_path = os.path.join(log_dir, "convert_%s.log" % datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
    logger = RunLogger(log_path, logger_callback)
    summary = RunSummary(total=len(files), log_path=log_path)
    logger.info("开始转换：%d 个文件，输出目录 %s" % (len(files), output_dir))

    jobs: List[FileJob] = []
    seen = set()
    for path in files:
        if os.path.abspath(path) in seen:
            logger.warn("忽略重复文件: %s" % path)
            continue
        seen.add(os.path.abspath(path))
        job = FileJob(path=path)
        jobs.append(job)

        override_id = overrides.get(path) if overrides else None
        if override_id:
            job.template = registry.by_id(override_id)
        if job.template is None:
            try:
                preview_text = preview(path, registry.templates[0].signature_scan_lines if registry.templates else 8)
            except Exception:
                preview_text = ""
            job.template = registry.match(path, preview_text)
        if job.template is None:
            job.status = "失败"
            summary.failed += 1
            _status_cb(status_callback, job, "失败", "未匹配到模板，可在界面手动指定", 0.0)
            logger.error("未匹配到模板: %s" % path, os.path.basename(path))
            continue

        _status_cb(status_callback, job, "解析中", "模板: %s" % job.template.display_name, 0.1)
        logger.info("使用模板 %s" % job.template.display_name, os.path.basename(path))
        try:
            job.document = parse_text(path, job.template)
        except Exception as exc:
            job.status = "失败"
            summary.failed += 1
            _status_cb(status_callback, job, "失败", "解析失败: %s" % exc, 0.0)
            logger.error("解析失败: %s" % exc, os.path.basename(path))
            continue

        for warning in job.document.warnings:
            logger.info(warning, os.path.basename(path))
        _status_cb(
            status_callback, job, "解析完成",
            "编码 %s，%d 组 / %d 行，等待写入" % (
                job.document.encoding, len(job.document.groups), job.document.total_rows),
            0.45,
        )
        if progress_callback:
            progress_callback({"type": "parsed", "path": path})

    # 按输出工作簿分组
    groups: "OrderedDict[str, List[FileJob]]" = OrderedDict()
    for job in jobs:
        if not job.document:
            continue
        key = os.path.abspath(os.path.join(output_dir, job.template.output_workbook))
        groups.setdefault(key, []).append(job)

    for output_path, group_jobs in groups.items():
        sample_paths = {os.path.abspath(j.template.sample_workbook_path) for j in group_jobs}
        if len(sample_paths) > 1:
            for job in group_jobs:
                summary.failed += 1
                _status_cb(status_callback, job, "失败", "同输出工作簿使用了不同范例工作簿，无法合并", 0.0)
                logger.error("范例工作簿冲突: %s" % output_path, os.path.basename(job.path))
            continue
        for job in group_jobs:
            _status_cb(status_callback, job, "写入中", "生成 %s" % os.path.basename(output_path), 0.6)
        try:
            stats = build_workbook(output_path, [j.document for j in group_jobs], logger)
        except PermissionError:
            for job in group_jobs:
                summary.failed += 1
                _status_cb(status_callback, job, "失败", "输出文件被占用（请关闭 Excel 后重试）", 0.0)
                logger.error("输出文件被占用: %s" % output_path, os.path.basename(job.path))
            continue
        except Exception as exc:
            for job in group_jobs:
                summary.failed += 1
                _status_cb(status_callback, job, "失败", "写入失败: %s" % exc, 0.0)
                logger.error("写入失败: %s" % exc, os.path.basename(job.path))
            continue

        summary.outputs.append(output_path)
        logger.ok("已生成 %s（%d 个 sheet）" % (output_path, len(group_jobs)))
        for job in group_jobs:
            job.stats = stats.get(job.path)
            job.output = output_path
            summary.succeeded += 1
            detail = "输出 %s" % output_path
            if job.stats:
                detail += "；%d 行 / %d 组 / %d 处合并 / %d 条批注" % (
                    job.stats.rows, job.stats.groups, job.stats.merges, job.stats.comments)
            _status_cb(status_callback, job, "成功", detail, 1.0)
            if progress_callback:
                progress_callback({"type": "done", "path": job.path})
            if job.stats:
                logger.info(
                    "统计: 行=%d 组=%d 合并=%d 批注=%d 模板公式格=%d 写TXT值=%d 公式补值=%d "
                    "异常标记=%d 非法数值=%d 离群=%d"
                    % (job.stats.rows, job.stats.groups, job.stats.merges, job.stats.comments,
                       job.stats.formula_cells, job.stats.formulas_replaced, job.stats.formulas_filled,
                       job.stats.anomalies, job.stats.invalid_values, job.stats.outliers),
                    os.path.basename(job.path),
                )

    logger.info("结束：成功 %d，失败 %d，生成文件 %d 个" % (summary.succeeded, summary.failed, len(summary.outputs)))
    logger.close()
    return summary
