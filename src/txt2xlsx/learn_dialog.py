# -*- coding: utf-8 -*-
"""自动学习模板的预览/确认面板。"""
from __future__ import annotations

import os
import tkinter as tk
from tkinter import messagebox, ttk
from typing import Dict, Optional

from .template_learner import LearnResult, build_candidate, save_learned_template

CONFIDENCE_LABEL = {
    "high": "高",
    "medium": "中",
    "low": "低（请确认）",
    "none": "未匹配",
}
CONFIDENCE_COLOR_ATTR = {
    "high": "success",
    "medium": "accent_soft",
    "low": "warning",
    "none": "error",
}


class LearnDialog(object):
    def __init__(self, app, result: LearnResult):
        self.app = app
        self.result = result
        self.palette = app.palette
        self.vars: Dict[str, tk.StringVar] = {}
        self.option_labels = ["（不映射）"]
        for idx, header in enumerate(result.txt_headers):
            text = header or "(空表头)"
            self.option_labels.append("%02d  %s" % (idx + 1, text))

        self.top = tk.Toplevel(app.root)
        self.top.title("学习新模板 — %s" % result.sheet_name)
        self.top.configure(bg=self.palette.panel)
        self.top.transient(app.root)
        self.top.geometry("1060x660")
        self._center()
        self._build()
        self.top.grab_set()

    def _center(self) -> None:
        self.top.update_idletasks()
        sw = self.top.winfo_screenwidth()
        sh = self.top.winfo_screenheight()
        self.top.geometry("1060x660+%d+%d" % (max(0, (sw - 1060) // 2), max(0, (sh - 660) // 2 - 30)))

    def _build(self) -> None:
        p = self.palette
        head = tk.Frame(self.top, bg=p.panel)
        head.pack(fill="x", padx=18, pady=(14, 6))
        tk.Label(head, text="自动识别结果", bg=p.panel, fg=p.text,
                 font=self.app.font_section).pack(side="left")
        tk.Label(
            head,
            text="整体置信度：%s" % CONFIDENCE_LABEL.get(self.result.overall_confidence, "低"),
            bg=p.panel,
            fg=getattr(p, CONFIDENCE_COLOR_ATTR.get(self.result.overall_confidence, "warning")),
            font=self.app.font_body_bold,
        ).pack(side="right")

        info = tk.Frame(self.top, bg=p.panel)
        info.pack(fill="x", padx=18)
        info_text = (
            "范例：%s    工作表：%s\n"
            "Excel 表头 %d 行，数据区第 %d–%d 行；TXT 表头 %d 行；组字段列：%s；连续值合并列：%s"
            % (
                os.path.basename(self.result.workbook_path), self.result.sheet_name,
                self.result.header_rows, self.result.data_start, self.result.data_end,
                self.result.txt_header_lines,
                "、".join(self.result.group_columns) or "无",
                "、".join(self.result.run_columns) or "无",
            )
        )
        tk.Label(info, text=info_text, bg=p.panel, fg=p.text_muted,
                 font=self.app.font_tiny, justify="left", anchor="w").pack(fill="x", pady=(2, 6))

        if self.result.warnings:
            warn_box = tk.Text(
                self.top, height=3, bg=p.log_bg, fg=p.warning, relief="flat", bd=0,
                font=self.app.font_tiny, wrap="word", padx=8, pady=5,
            )
            warn_box.pack(fill="x", padx=18)
            for line in self.result.warnings:
                warn_box.insert("end", "· " + line + "\n")
            warn_box.config(state="disabled")

        body = tk.Frame(self.top, bg=p.panel)
        body.pack(fill="both", expand=True, padx=18, pady=(10, 6))
        body.columnconfigure(0, weight=3)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(0, weight=1)

        left = tk.Frame(body, bg=p.panel)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(1, weight=1)
        tk.Label(left, text="列映射（Excel ← TXT，可下拉修改）", bg=p.panel, fg=p.text,
                 font=self.app.font_body_bold).grid(row=0, column=0, sticky="w", pady=(0, 4))

        wrap = tk.Frame(left, bg=p.panel)
        wrap.grid(row=1, column=0, sticky="nsew")
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)
        canvas = tk.Canvas(wrap, bg=p.panel_alt, highlightthickness=0, bd=0)
        canvas.grid(row=0, column=0, sticky="nsew")
        bar = ttk.Scrollbar(wrap, orient="vertical", command=canvas.yview,
                            style="Vertical.TScrollbar")
        bar.grid(row=0, column=1, sticky="ns", padx=(4, 0))
        canvas.configure(yscrollcommand=bar.set)
        inner = tk.Frame(canvas, bg=p.panel_alt)
        window = canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(window, width=e.width))
        canvas.bind("<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 120.0), "units"))

        tf = ttk.Style(self.top)
        tf.configure("Learn.TCombobox", fieldbackground=p.panel_raised, background=p.panel_raised,
                     foreground=p.text, arrowcolor=p.accent_soft, padding=2)

        for row_idx, mapping in enumerate(self.result.mappings):
            row_bg = p.panel_alt if row_idx % 2 == 0 else p.panel
            row = tk.Frame(inner, bg=row_bg, padx=6, pady=4)
            row.pack(fill="x", expand=True)
            row.columnconfigure(2, weight=1)
            tk.Label(row, text=mapping.excel, bg=row_bg, fg=p.accent_soft, width=3,
                     font=self.app.font_body_bold, anchor="w").grid(row=0, column=0, sticky="w")
            tk.Label(row, text=mapping.excel_header or "(空)", bg=row_bg, fg=p.text,
                     font=self.app.font_tiny, width=22, anchor="w", wraplength=190,
                     justify="left").grid(row=0, column=1, sticky="w", padx=(0, 6))
            var = tk.StringVar(value=self._mapping_label(mapping.txt))
            self.vars[mapping.excel] = var
            combo = ttk.Combobox(row, textvariable=var, values=self.option_labels,
                                 state="readonly", width=24, font=self.app.font_tiny,
                                 style="Learn.TCombobox")
            combo.grid(row=0, column=2, sticky="ew")
            confidence = mapping.confidence
            color = getattr(p, CONFIDENCE_COLOR_ATTR.get(confidence, "warning"))
            tk.Label(row, text=CONFIDENCE_LABEL.get(confidence, confidence), bg=row_bg,
                     fg=color, font=self.app.font_tiny, width=9, anchor="w").grid(
                row=0, column=3, sticky="w", padx=(6, 0))
            tk.Label(row, text=mapping.note, bg=row_bg, fg=p.text_muted,
                     font=self.app.font_tiny, anchor="w").grid(
                row=1, column=1, columnspan=3, sticky="ew")

        right = tk.Frame(body, bg=p.panel)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)
        tk.Label(right, text="试转预览（前 5 行写入值）", bg=p.panel, fg=p.text,
                 font=self.app.font_body_bold).grid(row=0, column=0, sticky="w", pady=(0, 4))
        preview = tk.Text(right, bg=p.log_bg, fg=p.text, relief="flat", bd=0,
                          font=self.app.font_tiny, wrap="word", padx=8, pady=6)
        preview.grid(row=1, column=0, sticky="nsew")
        if self.result.preview_rows:
            for idx, cells in enumerate(self.result.preview_rows, 1):
                preview.insert("end", "第 %d 行：%s\n\n" % (idx, " | ".join(cells)))
        else:
            preview.insert("end", self.result.preview_note or "（无预览）")
        if self.result.preview_note:
            preview.insert("end", "\n" + self.result.preview_note)
        preview.config(state="disabled")

        footer = tk.Frame(self.top, bg=p.panel)
        footer.pack(fill="x", padx=18, pady=(4, 14))
        tk.Label(
            footer, text="修改映射后点「保存为模板」立即生效；取消不会写入任何文件。",
            bg=p.panel, fg=p.text_muted, font=self.app.font_tiny,
        ).pack(side="left")
        tk.Button(footer, text="取消", command=self._cancel, bg=p.panel_raised, fg=p.text,
                  activebackground=p.accent_dark, relief="flat", bd=0, font=self.app.font_small,
                  padx=14, pady=5, cursor="hand2").pack(side="right")
        tk.Button(footer, text="保存为模板", command=self._save, bg=p.accent, fg="#10241A",
                  activebackground=p.accent_soft, relief="flat", bd=0, font=self.app.font_button,
                  padx=16, pady=5, cursor="hand2").pack(side="right", padx=(0, 8))

    def _mapping_label(self, txt: Optional[int]) -> str:
        if txt is None or txt < 0 or txt >= len(self.result.txt_headers):
            return "（不映射）"
        return self.option_labels[txt + 1]

    def _current_txt(self, excel: str) -> Optional[int]:
        label = self.vars[excel].get()
        if label not in self.option_labels[1:]:
            return None
        return self.option_labels.index(label) - 1

    def _save(self) -> None:
        for mapping in self.result.mappings:
            mapping.txt = self._current_txt(mapping.excel)
            mapping.txt_header = (
                self.result.txt_headers[mapping.txt] if mapping.txt is not None
                and mapping.txt < len(self.result.txt_headers) else "")
        self.result.group_fill_columns = [
            m.txt for m in self.result.mappings
            if m.excel in self.result.group_columns and m.txt is not None
        ]
        try:
            self.result.template = build_candidate(self.result)
            path = save_learned_template(self.result)
        except Exception as exc:
            messagebox.showerror("保存失败", str(exc), parent=self.top)
            return
        self.app.on_template_saved(path, self.result)
        self.top.grab_release()
        self.top.destroy()

    def _cancel(self) -> None:
        self.top.grab_release()
        self.top.destroy()
