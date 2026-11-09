# -*- coding: utf-8 -*-
"""tkinter 图形界面：背景图融入、批量转换、逐文件进度与日志。"""
from __future__ import annotations

import os
import queue
import threading
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional

from PIL import Image, ImageFilter, ImageTk

from .encoding_utils import preview
from .learn_dialog import LearnDialog
from .models import PROJECT_ROOT, Template, TemplateRegistry
from .paths import ensure_user_dirs, resource_path, user_path
from .pipeline import RunSummary, run_conversion
from .template_learner import learn_template
from .theme import Palette, ambient_backdrop, build_palette, character_image

ASSETS_DIR = resource_path("assets", "ui")
BACKGROUND_PATH = os.path.join(ASSETS_DIR, "background.png")
ICON_PATH = os.path.join(ASSETS_DIR, "icon.png")
ICON_ICO_PATH = os.path.join(ASSETS_DIR, "icon.ico")
DEFAULT_OUTPUT = user_path("output")
AUTO_TEMPLATE_LABEL = "自动匹配"


class FileRow(object):
    """文件列表中的一行：文件名、模板选择、状态与独立进度条。"""

    def __init__(self, app: "App", path: str, index: int):
        self.app = app
        self.path = path
        self.palette = app.palette
        self.template_id: Optional[str] = None
        self.matched: Optional[Template] = None
        self._silent = False
        self.progress_value = 0.0

        self.frame = tk.Frame(app.list_inner, bg=self.palette.panel_alt, padx=10, pady=8)
        self.frame.grid(row=index, column=0, sticky="ew", pady=(0, 6))
        self.frame.columnconfigure(1, weight=1)

        self.index_label = tk.Label(
            self.frame, text="%02d" % (index + 1), bg=self.palette.panel_alt,
            fg=self.palette.accent_soft, font=app.font_small, width=3, anchor="w",
        )
        self.index_label.grid(row=0, column=0, rowspan=2, sticky="w")

        display_name = os.path.basename(path)
        if len(display_name) > 30:
            display_name = display_name[:27] + "…"
        self.name_label = tk.Label(
            self.frame, text=display_name, bg=self.palette.panel_alt, fg=self.palette.text,
            font=app.font_body_bold, anchor="w",
        )
        self.name_label.grid(row=0, column=1, sticky="ew")

        self.detail_label = tk.Label(
            self.frame, text=os.path.dirname(path), bg=self.palette.panel_alt,
            fg=self.palette.text_muted, font=app.font_tiny, anchor="w",
        )
        self.detail_label.grid(row=1, column=1, sticky="ew")

        self.combo_var = tk.StringVar(value=AUTO_TEMPLATE_LABEL)
        self.combo = ttk.Combobox(
            self.frame, textvariable=self.combo_var, state="readonly",
            values=app.template_labels(), width=20, font=app.font_small,
        )
        self.combo.grid(row=0, column=2, rowspan=2, padx=(10, 10), sticky="ew")
        self.combo.bind("<<ComboboxSelected>>", self._on_combo)

        self.status_label = tk.Label(
            self.frame, text="待处理", bg=self.palette.panel_alt, fg=self.palette.text_muted,
            font=app.font_small, width=10, anchor="w",
        )
        self.status_label.grid(row=0, column=3, sticky="w")

        self.progress = ttk.Progressbar(
            self.frame, style="Row.Horizontal.TProgressbar", mode="determinate",
            maximum=100, value=0, length=84,
        )
        self.progress.grid(row=1, column=3, sticky="w", pady=(3, 0))

        self.remove_btn = tk.Button(
            self.frame, text="移除", command=lambda: app.remove_row(self),
            bg=self.palette.panel_alt, fg=self.palette.text_muted,
            activebackground=self.palette.panel_raised, activeforeground=self.palette.text,
            relief="flat", bd=0, font=app.font_tiny, cursor="hand2",
        )
        self.remove_btn.grid(row=0, column=4, rowspan=2, padx=(6, 0))

    def set_index(self, index: int) -> None:
        self.index_label.config(text="%02d" % (index + 1))

    def _on_combo(self, _event=None) -> None:
        if self._silent:
            return
        selected = self.combo_var.get()
        if selected == AUTO_TEMPLATE_LABEL:
            self.template_id = None
            self.app.rematch_row(self)
        else:
            template = self.app.template_by_label(selected)
            if template:
                self.template_id = template.id
                self.matched = template
                self.set_status("待处理", "手动指定：%s" % template.display_name, self.progress_value)

    def set_template(self, template: Optional[Template], manual: bool) -> None:
        self.matched = template
        if manual and template:
            self.template_id = template.id
            label = self.app.template_short_label(template)
        else:
            self.template_id = None
            label = AUTO_TEMPLATE_LABEL
        self._silent = True
        self.combo_var.set(label)
        self._silent = False

    def set_status(self, status: str, detail: str = "", progress: Optional[float] = None) -> None:
        color = self.app.status_color(status)
        self.status_label.config(text=status, fg=color)
        if detail:
            self.detail_label.config(text=detail)
        if progress is not None:
            self.progress_value = max(0.0, min(100.0, progress * 100.0))
            self.progress["value"] = self.progress_value

    def set_running(self, running: bool) -> None:
        state = "disabled" if running else "readonly"
        self.combo.config(state=state)
        self.remove_btn.config(state="disabled" if running else "normal",
                               fg=self.palette.text_muted if not running else self.palette.panel_raised)

    def destroy(self) -> None:
        self.frame.destroy()


class App(object):
    def __init__(self, root: tk.Tk):
        self.root = root
        self.palette = build_palette(BACKGROUND_PATH, ICON_PATH)
        self.registry = TemplateRegistry()
        self.rows: List[FileRow] = []
        self.rows_by_path: Dict[str, FileRow] = {}
        self.running = False
        self.learning = False
        self.queue: "queue.Queue[dict]" = queue.Queue()
        self.bg_photo: Optional[ImageTk.PhotoImage] = None
        self.bg_item: Optional[int] = None
        self._resize_job = None
        self._last_background_size = (0, 0)
        self.ambient_source = None
        self.list_bg_item: Optional[int] = None
        self.list_bg_photo: Optional[ImageTk.PhotoImage] = None
        self.list_bg_key = None
        self.list_char_item: Optional[int] = None
        self.list_char_photo: Optional[ImageTk.PhotoImage] = None
        self.list_char_key = None
        self.list_char_job = None
        self.list_hint_items: List[int] = []
        self.summary: Optional[RunSummary] = None
        self.last_output_dir = DEFAULT_OUTPUT

        self._setup_window()
        self._setup_fonts()
        self._setup_style()
        self._build_ui()
        self._render_background()
        self.root.after(80, self._layout_cards)
        self._poll_queue()

    # ------------------------------------------------------------------ 窗口
    def _setup_window(self) -> None:
        self.root.title("TXT → Excel 样式重建")
        self.root.geometry("1180x740")
        self.root.minsize(1000, 690)
        self.root.configure(bg=self.palette.panel)
        self._center_window(1180, 740)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.bind("<Configure>", self._on_configure)
        if os.path.exists(ICON_ICO_PATH):
            try:
                self.root.iconbitmap(ICON_ICO_PATH)
            except Exception:
                pass
        try:
            icon = Image.open(ICON_PATH).convert("RGBA").resize((64, 64), Image.LANCZOS)
            self.icon_photo = ImageTk.PhotoImage(icon)
            self.root.iconphoto(True, self.icon_photo)
        except Exception:
            self.icon_photo = None

    def _center_window(self, width: int, height: int) -> None:
        try:
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            x = max(0, (sw - width) // 2)
            y = max(0, (sh - height) // 2 - 20)
            self.root.geometry("%dx%d+%d+%d" % (width, height, x, y))
        except Exception:
            pass

    def _setup_fonts(self) -> None:
        families = set(tkfont.families(self.root))
        family = "Microsoft YaHei UI"
        if family not in families:
            family = "Microsoft YaHei" if "Microsoft YaHei" in families else "SimHei"
        self.font_family = family
        self.font_title = (family, 19, "bold")
        self.font_subtitle = (family, 10)
        self.font_section = (family, 11, "bold")
        self.font_body = (family, 9)
        self.font_body_bold = (family, 9, "bold")
        self.font_small = (family, 8)
        self.font_tiny = (family, 7)
        self.font_button = (family, 11, "bold")
        self.font_log = (family, 8)

    def _setup_style(self) -> None:
        p = self.palette
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(
            "TCombobox", fieldbackground=p.panel_raised, background=p.panel_raised,
            foreground=p.text, arrowcolor=p.accent_soft, bordercolor=p.border,
            lightcolor=p.panel_raised, darkcolor=p.panel_raised, padding=3,
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", p.panel_raised), ("disabled", p.panel_alt)],
            foreground=[("readonly", p.text), ("disabled", p.text_muted)],
            selectbackground=[("readonly", p.accent_dark)],
            selectforeground=[("readonly", p.text)],
        )
        style.configure(
            "Row.Horizontal.TProgressbar", troughcolor=p.log_bg, background=p.accent,
            bordercolor=p.panel_alt, lightcolor=p.accent, darkcolor=p.accent, thickness=7,
        )
        style.configure(
            "Overall.Horizontal.TProgressbar", troughcolor=p.log_bg, background=p.accent,
            bordercolor=p.panel, lightcolor=p.accent, darkcolor=p.accent, thickness=12,
        )
        style.configure(
            "Vertical.TScrollbar", background=p.panel_raised, troughcolor=p.panel,
            arrowcolor=p.accent_soft, bordercolor=p.border, darkcolor=p.panel_raised,
            lightcolor=p.panel_raised, gripcount=0,
        )
        style.map(
            "Vertical.TScrollbar",
            background=[("active", p.accent), ("pressed", p.accent_dark)],
        )

    # ------------------------------------------------------------------ 背景
    def _on_configure(self, event) -> None:
        if event.widget is not self.root:
            return
        if self._resize_job:
            self.root.after_cancel(self._resize_job)
        self._resize_job = self.root.after(180, self._after_resize)

    def _after_resize(self) -> None:
        self._resize_job = None
        self._render_background()
        self._layout_cards()

    def _render_background(self) -> None:
        width = max(1, self.root.winfo_width())
        height = max(1, self.root.winfo_height())
        if width < 200 or height < 200:
            return
        if (width, height) == self._last_background_size:
            return
        try:
            image = ambient_backdrop((width, height), BACKGROUND_PATH, blur=12, darkness=0.34)
            self.bg_photo = ImageTk.PhotoImage(image)
            self.ambient_source = image
            self.list_bg_key = None
        except Exception:
            self.bg_photo = None
            self.ambient_source = None
        if self.bg_photo is None:
            return
        if self.bg_item is None:
            self.bg_item = self.canvas.create_image(0, 0, anchor="nw", image=self.bg_photo, tags="bg")
            self.canvas.tag_lower("bg")
        else:
            self.canvas.itemconfigure(self.bg_item, image=self.bg_photo)
        self._last_background_size = (width, height)

    # ------------------------------------------------------------------ 布局
    def _layout_cards(self) -> None:
        width = max(1, self.root.winfo_width())
        height = max(1, self.root.winfo_height())
        margin = 20
        header_bottom = 96
        log_height = 178 if height >= 720 else 152
        middle_height = max(300, height - header_bottom - log_height - margin - 16)
        left_width = int((width - margin * 3) * 0.62)
        right_width = width - margin * 3 - left_width
        log_y = header_bottom + middle_height + 12

        self.left_card.place(x=margin, y=header_bottom, width=left_width, height=middle_height)
        self.right_card.place(x=margin + left_width + margin, y=header_bottom,
                              width=right_width, height=middle_height)
        self.log_card.place(x=margin, y=log_y, width=width - margin * 2, height=log_height)

        if self.status_item is not None:
            self.canvas.coords(self.status_item, width - margin, 34)
        if self.header_note_item is not None:
            self.canvas.coords(self.header_note_item, width - margin, 62)
        self._refresh_card_backdrops()

    def _blended_ambient(self, x: int, y: int, width: int, height: int):
        """取背景层对应区域，按 85% 面板色 + 15% 环境色混合，形成轻微透底。"""
        source = self.ambient_source
        if source is None:
            return None
        panel = self.palette.panel
        panel_rgb = (int(panel[1:3], 16), int(panel[3:5], 16), int(panel[5:7], 16))
        box = (max(0, x), max(0, y),
               min(source.width, x + width), min(source.height, y + height))
        if box[2] > box[0] and box[3] > box[1]:
            crop = source.crop(box)
        else:
            crop = Image.new("RGB", (width, height), panel_rgb)
        if crop.size != (width, height):
            crop = crop.resize((width, height), Image.BILINEAR)
        # 低不透明度下细节会被面板底色吃掉，先轻度锐化让轮廓在面板里更可读。
        crop = crop.filter(ImageFilter.UnsharpMask(radius=2, percent=140, threshold=2))
        return Image.blend(crop, Image.new("RGB", crop.size, panel_rgb), 0.85)

    def _refresh_card_backdrops(self) -> None:
        if self.ambient_source is None:
            return
        self.root.update_idletasks()
        root_x = self.root.winfo_rootx()
        root_y = self.root.winfo_rooty()
        for card in (self.left_card, self.right_card, self.log_card):
            label = getattr(card, "backdrop_label", None)
            if label is None:
                continue
            try:
                x = card.winfo_rootx() - root_x
                y = card.winfo_rooty() - root_y
                card_w = card.winfo_width()
                card_h = card.winfo_height()
            except Exception:
                continue
            if card_w < 20 or card_h < 20:
                continue
            key = (x, y, card_w, card_h)
            if getattr(card, "_backdrop_key", None) == key:
                continue
            blended = self._blended_ambient(x, y, card_w, card_h)
            if blended is None:
                continue
            photo = ImageTk.PhotoImage(blended)
            card._backdrop_photo = photo
            card._backdrop_key = key
            label.configure(image=photo)
        self._refresh_list_backdrop()

    def _refresh_list_backdrop(self) -> None:
        """把同样的透底铺进列表画布，让装饰在空白/行间区域隐约可辨。"""
        canvas = self.list_canvas
        width = max(1, canvas.winfo_width())
        height = max(1, canvas.winfo_height())
        if self.ambient_source is None or width < 40 or height < 40:
            return
        x = canvas.winfo_rootx() - self.root.winfo_rootx()
        y = canvas.winfo_rooty() - self.root.winfo_rooty()
        key = (x, y, width, height)
        if self.list_bg_key == key:
            return
        blended = self._blended_ambient(x, y, width, height)
        if blended is None:
            return
        photo = ImageTk.PhotoImage(blended)
        self.list_bg_photo = photo
        if self.list_bg_item is None:
            self.list_bg_item = canvas.create_image(
                0, 0, anchor="nw", image=photo, tags=("list_bg",))
        else:
            canvas.itemconfigure(self.list_bg_item, image=photo)
        canvas.tag_lower("list_bg")
        self.list_bg_key = key

    def _build_ui(self) -> None:
        p = self.palette
        self.canvas = tk.Canvas(self.root, highlightthickness=0, bd=0, bg=p.panel)
        self.canvas.place(x=0, y=0, relwidth=1, relheight=1)

        self.header_icon = None
        try:
            header_image = Image.open(ICON_PATH).convert("RGBA").resize((54, 54), Image.LANCZOS)
            self.header_icon = ImageTk.PhotoImage(header_image)
        except Exception:
            pass
        if self.header_icon is not None:
            self.canvas.create_image(20, 20, anchor="nw", image=self.header_icon, tags="header")
        self.canvas.create_text(
            20 + 66, 22, text="TXT → Excel 样式重建", anchor="nw",
            fill=p.text, font=self.font_title, tags="header",
        )
        self.canvas.create_text(
            20 + 66, 56, text="以范例工作簿为样式源，批量还原字体、边框、合并、批注与页面设置",
            anchor="nw", fill=p.text_muted, font=self.font_subtitle, tags="header",
        )
        self.status_item = self.canvas.create_text(
            1160, 34, text="就绪", anchor="e", fill=p.accent_soft, font=(self.font_family, 10, "bold"),
            tags="header",
        )
        self.header_note_item = self.canvas.create_text(
            1160, 62, text="Python 3.9 · openpyxl", anchor="e", fill=p.text_muted,
            font=self.font_tiny, tags="header",
        )

        self._build_file_card()
        self._build_action_card()
        self._build_log_card()

    def _card(self, parent=None) -> tk.Frame:
        p = self.palette
        frame = tk.Frame(parent or self.root, bg=p.panel, highlightbackground=p.border,
                         highlightthickness=1, bd=0)
        frame.grid_propagate(True)
        backdrop = tk.Label(frame, bg=p.panel, bd=0, highlightthickness=0)
        backdrop.place(x=0, y=0, relwidth=1, relheight=1)
        backdrop.lower()
        frame.backdrop_label = backdrop
        return frame

    def _build_file_card(self) -> None:
        p = self.palette
        self.left_card = self._card()
        self.left_card.columnconfigure(0, weight=1)
        self.left_card.rowconfigure(2, weight=1)

        header = tk.Frame(self.left_card, bg=p.panel)
        header.grid(row=0, column=0, sticky="ew", padx=16, pady=(13, 8))
        header.columnconfigure(0, weight=1)
        tk.Label(header, text="待转换文件", bg=p.panel, fg=p.text, font=self.font_section).grid(
            row=0, column=0, sticky="w")
        self.file_count_label = tk.Label(header, text="0 个", bg=p.panel, fg=p.text_muted,
                                         font=self.font_small)
        self.file_count_label.grid(row=0, column=1, sticky="e", padx=(0, 10))

        actions = tk.Frame(self.left_card, bg=p.panel)
        actions.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 8))
        self.add_btn = self._button(actions, "添加 TXT 文件", self.add_files, primary=True)
        self.add_btn.pack(side="left")
        self.clear_btn = self._button(actions, "清空列表", self.clear_rows)
        self.clear_btn.pack(side="left", padx=(8, 0))

        list_wrap = tk.Frame(self.left_card, bg=p.panel)
        list_wrap.grid(row=2, column=0, sticky="nsew", padx=(16, 8), pady=(0, 12))
        list_wrap.columnconfigure(0, weight=1)
        list_wrap.rowconfigure(0, weight=1)
        self.list_canvas = tk.Canvas(list_wrap, bg=p.panel, highlightthickness=0, bd=0)
        self.list_canvas.grid(row=0, column=0, sticky="nsew")
        self.list_scrollbar = self._scrollbar(list_wrap, self.list_canvas.yview)
        self.list_scrollbar.grid(row=0, column=1, sticky="ns", padx=(4, 0))
        self.list_canvas.configure(yscrollcommand=self._on_list_yscroll)
        self.list_inner = tk.Frame(self.list_canvas, bg=p.panel)
        self.list_window = self.list_canvas.create_window((0, 0), window=self.list_inner, anchor="nw")
        self.list_inner.bind("<Configure>", self._on_list_inner_configure)
        self.list_canvas.bind("<Configure>", self._on_list_canvas_configure)
        self.root.bind_all("<MouseWheel>", self._on_mousewheel)

    def _build_action_card(self) -> None:
        p = self.palette
        self.right_card = self._card()
        self.right_card.columnconfigure(0, weight=1)
        self.right_card.rowconfigure(2, weight=1)

        title = tk.Label(self.right_card, text="输出与执行", bg=p.panel, fg=p.text,
                         font=self.font_section)
        title.grid(row=0, column=0, sticky="w", padx=16, pady=(13, 8))

        form = tk.Frame(self.right_card, bg=p.panel)
        form.grid(row=1, column=0, sticky="nsew", padx=16)
        form.columnconfigure(0, weight=1)

        tk.Label(form, text="输出目录", bg=p.panel, fg=p.text_muted, font=self.font_small).grid(
            row=0, column=0, sticky="w", pady=(0, 4))
        path_row = tk.Frame(form, bg=p.panel)
        path_row.grid(row=1, column=0, sticky="ew")
        path_row.columnconfigure(0, weight=1)
        self.output_var = tk.StringVar(value=DEFAULT_OUTPUT)
        self.output_entry = tk.Entry(
            path_row, textvariable=self.output_var, bg=p.panel_raised, fg=p.text,
            insertbackground=p.accent, relief="flat", bd=0, font=self.font_small,
            highlightthickness=1, highlightbackground=p.border, highlightcolor=p.accent,
        )
        self.output_entry.grid(row=0, column=0, sticky="ew", ipady=5)
        self.browse_btn = self._button(path_row, "浏览", self.browse_output)
        self.browse_btn.grid(row=0, column=1, padx=(6, 0))

        divider = tk.Frame(form, bg=p.border, height=1)
        divider.grid(row=2, column=0, sticky="ew", pady=14)

        tk.Label(form, text="模板匹配", bg=p.panel, fg=p.text_muted, font=self.font_small).grid(
            row=3, column=0, sticky="w", pady=(0, 4))
        template_text = self._template_summary()
        self.template_label = tk.Label(
            form, text=template_text, bg=p.panel, fg=p.text_muted, font=self.font_tiny,
            justify="left", anchor="w", wraplength=320,
        )
        self.template_label.grid(row=4, column=0, sticky="ew")
        self.template_hint = tk.Label(
            form, text="每个文件右侧可单独指定模板；自动匹配失败时状态会提示。",
            bg=p.panel, fg=p.text_muted, font=self.font_tiny, justify="left", anchor="w",
            wraplength=320,
        )
        self.template_hint.grid(row=5, column=0, sticky="ew", pady=(4, 0))

        self.learn_btn = self._button(form, "学习新模板…", self.start_learning)
        self.learn_btn.grid(row=6, column=0, sticky="ew", pady=(10, 0))

        self.start_btn = tk.Button(
            form, text="开始转换", command=self.start_conversion, bg=p.accent, fg="#10241A",
            activebackground=p.accent_soft, activeforeground="#10241A", relief="flat", bd=0,
            font=self.font_button, cursor="hand2", padx=10, pady=9,
        )
        self.start_btn.grid(row=7, column=0, sticky="ew", pady=(12, 10))

        overall = tk.Frame(form, bg=p.panel)
        overall.grid(row=8, column=0, sticky="ew")
        overall.columnconfigure(0, weight=1)
        self.overall_bar = ttk.Progressbar(
            overall, style="Overall.Horizontal.TProgressbar", mode="determinate",
            maximum=100, value=0,
        )
        self.overall_bar.grid(row=0, column=0, sticky="ew")
        self.overall_label = tk.Label(overall, text="0 / 0", bg=p.panel, fg=p.text_muted,
                                      font=self.font_tiny)
        self.overall_label.grid(row=1, column=0, sticky="e", pady=(3, 0))

        self.open_btn = self._button(form, "打开输出目录", self.open_output_dir)
        self.open_btn.grid(row=9, column=0, sticky="ew", pady=(10, 0))
        self.open_btn.config(state="disabled")

        # 用一个可伸缩的空白行把内容顶在上方。
        form.rowconfigure(10, weight=1)

    def _build_log_card(self) -> None:
        p = self.palette
        self.log_card = self._card()
        self.log_card.columnconfigure(0, weight=1)
        self.log_card.rowconfigure(1, weight=1)

        header = tk.Frame(self.log_card, bg=p.panel)
        header.grid(row=0, column=0, sticky="ew", padx=14, pady=(9, 5))
        header.columnconfigure(0, weight=1)
        tk.Label(header, text="运行日志", bg=p.panel, fg=p.text, font=self.font_section).grid(
            row=0, column=0, sticky="w")
        self.log_path_label = tk.Label(header, text="", bg=p.panel, fg=p.text_muted,
                                       font=self.font_tiny, anchor="e")
        self.log_path_label.grid(row=0, column=1, sticky="e", padx=(0, 10))
        self.log_clear_btn = self._button(header, "清空", self.clear_log)
        self.log_clear_btn.grid(row=0, column=2)

        body = tk.Frame(self.log_card, bg=p.panel)
        body.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 10))
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)
        self.log_text = tk.Text(
            body, bg=p.log_bg, fg=p.log_text, insertbackground=p.log_text, relief="flat",
            bd=0, font=self.font_log, wrap="word", height=6, padx=8, pady=6,
            highlightthickness=0,
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")
        log_scroll = self._scrollbar(body, self.log_text.yview)
        log_scroll.grid(row=0, column=1, sticky="ns", padx=(4, 0))
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.tag_configure("INFO", foreground=p.log_text)
        self.log_text.tag_configure("OK", foreground=p.success)
        self.log_text.tag_configure("WARN", foreground=p.warning)
        self.log_text.tag_configure("ERROR", foreground=p.error)
        self.log_text.tag_configure("DEBUG", foreground=p.text_muted)
        self.log_text.config(state="disabled")

        self._append_log("INFO", "就绪：%d 个模板已加载。" % len(self.registry.templates))

    # ------------------------------------------------------------------ 小部件
    def _button(self, parent, text: str, command, primary: bool = False) -> tk.Button:
        p = self.palette
        if primary:
            bg, fg, active_bg = p.accent, "#10241A", p.accent_soft
        else:
            bg, fg, active_bg = p.panel_raised, p.text, p.accent_dark
        return tk.Button(
            parent, text=text, command=command, bg=bg, fg=fg, activebackground=active_bg,
            activeforeground=fg, relief="flat", bd=0, font=self.font_small, cursor="hand2",
            padx=10, pady=4,
        )

    def _scrollbar(self, parent, command) -> ttk.Scrollbar:
        return ttk.Scrollbar(parent, orient="vertical", command=command,
                             style="Vertical.TScrollbar")

    def template_labels(self) -> List[str]:
        return [AUTO_TEMPLATE_LABEL] + [self.template_short_label(t) for t in self.registry.templates]

    def template_short_label(self, template: Template) -> str:
        """下拉框用的短标签，完整名称在文件行详情与日志中展示。"""
        name = template.display_name
        for sep in ("（", "("):
            if sep in name:
                name = name.split(sep)[0]
        return name.strip()

    def template_by_label(self, label: str) -> Optional[Template]:
        for template in self.registry.templates:
            if self.template_short_label(template) == label:
                return template
        return None

    def _template_summary(self) -> str:
        lines = []
        for template in self.registry.templates:
            lines.append("· %s" % template.display_name)
        return "\n".join(lines) if lines else "未找到模板配置（templates/*.json）"

    def status_color(self, status: str) -> str:
        p = self.palette
        if status in ("成功", "完成"):
            return p.success
        if status in ("失败", "未匹配"):
            return p.error
        if status in ("解析中", "解析完成", "写入中", "转换中"):
            return p.accent_soft
        return p.text_muted

    # ------------------------------------------------------------------ 文件
    def add_files(self) -> None:
        if self.running:
            return
        initial = os.path.join(PROJECT_ROOT, "samples")
        paths = filedialog.askopenfilenames(
            parent=self.root, title="选择仪器导出的 TXT 文件", initialdir=initial,
            filetypes=[("TXT 文本", "*.txt"), ("所有文件", "*.*")],
        )
        added = 0
        for path in paths:
            path = os.path.abspath(path)
            key = os.path.normcase(path)
            if key in self.rows_by_path:
                continue
            row = FileRow(self, path, len(self.rows))
            self.rows.append(row)
            self.rows_by_path[key] = row
            self.rematch_row(row)
            added += 1
        self._refresh_file_list()
        if added:
            self._append_log("INFO", "已添加 %d 个文件。" % added)

    def remove_row(self, row: FileRow) -> None:
        if self.running or row not in self.rows:
            return
        self.rows.remove(row)
        self.rows_by_path.pop(os.path.normcase(row.path), None)
        row.destroy()
        self._refresh_file_list()
        self._append_log("INFO", "已移除：%s" % os.path.basename(row.path))

    def clear_rows(self) -> None:
        if self.running:
            return
        for row in list(self.rows):
            row.destroy()
        self.rows = []
        self.rows_by_path.clear()
        self._refresh_file_list()

    def _refresh_file_list(self) -> None:
        for index, row in enumerate(self.rows):
            row.frame.grid_configure(row=index)
            row.set_index(index)
        self.file_count_label.config(text="%d 个" % len(self.rows))
        self.list_inner.update_idletasks()
        self._update_list_scrollregion()
        self._render_list_char()
        self.overall_label.config(text="0 / %d" % len(self.rows))
        self.overall_bar["value"] = 0

    # ------------------------------------------------------------ 列表人物底图
    def _on_list_canvas_configure(self, event) -> None:
        self.list_canvas.itemconfigure(self.list_window, width=event.width)
        self._schedule_list_char()

    def _on_list_inner_configure(self, _event=None) -> None:
        self._update_list_scrollregion()

    def _on_list_yscroll(self, first, last) -> None:
        self.list_scrollbar.set(first, last)
        self._sync_list_overlay()

    def _update_list_scrollregion(self) -> None:
        """只用列表内容的包围盒，避免把人物底图算进可滚动区域。"""
        bbox = self.list_canvas.bbox(self.list_window)
        if bbox:
            self.list_canvas.configure(scrollregion=bbox)

    def _schedule_list_char(self) -> None:
        if self.list_char_job is not None:
            try:
                self.root.after_cancel(self.list_char_job)
            except Exception:
                pass
        self.list_char_job = self.root.after(120, self._render_list_char)

    def _render_list_char(self) -> None:
        self.list_char_job = None
        canvas = self.list_canvas
        width = max(1, canvas.winfo_width())
        height = max(1, canvas.winfo_height())
        if width < 60 or height < 60:
            return
        self._refresh_list_backdrop()
        if self.rows:
            # 有文件时缩小并降到低透明度，垫在列表行下面当水印。
            max_height = int(height * 0.55)
            max_width = int(width * 0.46)
            opacity = 0.22
        else:
            max_height = int(height * 0.9)
            max_width = int(width * 0.72)
            opacity = 1.0
        try:
            image = character_image(BACKGROUND_PATH, max_height, max_width, opacity)
        except Exception:
            self._clear_list_char()
            return
        key = (image.width, image.height, opacity)
        if self.list_char_item is None or self.list_char_key != key:
            photo = ImageTk.PhotoImage(image)
            self.list_char_photo = photo
            if self.list_char_item is None:
                self.list_char_item = canvas.create_image(
                    0, 0, anchor="se", image=photo, tags=("list_char",))
                canvas.tag_lower("list_char")
                canvas.tag_lower("list_bg")
            else:
                canvas.itemconfigure(self.list_char_item, image=photo)
            self.list_char_key = key
        self._render_list_hint()
        self._sync_list_overlay()

    def _render_list_hint(self) -> None:
        for item in self.list_hint_items:
            try:
                self.list_canvas.delete(item)
            except Exception:
                pass
        self.list_hint_items = []
        if self.rows:
            return
        text = "还没有文件\n点击「添加 TXT 文件」，支持一次选择多个 txt。"
        shadow = self.list_canvas.create_text(
            0, 0, anchor="nw", text=text, font=self.font_body,
            fill="#0A140F", justify="left",
        )
        label = self.list_canvas.create_text(
            0, 0, anchor="nw", text=text, font=self.font_body,
            fill=self.palette.text, justify="left",
        )
        self.list_hint_items = [shadow, label]

    def _sync_list_overlay(self) -> None:
        """把底图与提示固定在可视区域，不随列表内容滚动。"""
        canvas = self.list_canvas
        width = max(1, canvas.winfo_width())
        height = max(1, canvas.winfo_height())
        top = canvas.canvasy(0)
        if self.list_bg_item is not None:
            canvas.coords(self.list_bg_item, 0, top)
        if self.list_char_item is not None:
            canvas.coords(self.list_char_item, width - 8, top + height - 6)
        if len(self.list_hint_items) == 2:
            shadow, label = self.list_hint_items
            x = 14
            y = top + 12
            canvas.coords(shadow, x + 1, y + 1)
            canvas.coords(label, x, y)

    def _clear_list_char(self) -> None:
        if self.list_char_item is not None:
            self.list_canvas.delete(self.list_char_item)
        self.list_char_item = None
        self.list_char_photo = None
        self.list_char_key = None

    def rematch_row(self, row: FileRow) -> None:
        try:
            scan = self.registry.templates[0].signature_scan_lines if self.registry.templates else 6
            text = preview(row.path, scan)
        except Exception:
            text = ""
        template = self.registry.match(row.path, text)
        row.set_template(template, manual=False)
        if template:
            row.set_status("待处理", "匹配：%s" % template.display_name, 0.0)
        else:
            row.set_status("未匹配", "请点击右侧下拉框手动指定模板", 0.0)

    # -------------------------------------------------------------- 学习模板
    def start_learning(self) -> None:
        if self.running or self.learning:
            return
        path = filedialog.askopenfilename(
            parent=self.root, title="选择范例 Excel（用于学习格式）",
            initialdir=os.path.join(PROJECT_ROOT, "assets", "templates"),
            filetypes=[("Excel 工作簿", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not path:
            return
        txt_path = None
        if messagebox.askyesno(
            "样例 TXT（可选）",
            "是否同时选择一个样例 TXT？\n\n有 TXT 才能自动对齐列名并试转预览；\n没有也可以学习格式，但列映射按位置占位。",
            parent=self.root,
        ):
            txt_path = filedialog.askopenfilename(
                parent=self.root, title="选择样例 TXT（可取消）",
                initialdir=os.path.join(PROJECT_ROOT, "samples"),
                filetypes=[("TXT 文本", "*.txt"), ("所有文件", "*.*")],
            ) or None
        self.learning = True
        self._set_learning_ui(True)
        self._set_header_status("学习模板中…")
        self._append_log("INFO", "开始学习模板：%s" % os.path.basename(path))
        threading.Thread(
            target=self._learn_worker, args=(path, txt_path), daemon=True
        ).start()

    def _learn_worker(self, path: str, txt_path: Optional[str]) -> None:
        try:
            result = learn_template(path, txt_path=txt_path)
            self.queue.put({"type": "learn_ready", "result": result})
        except Exception as exc:
            self.queue.put({"type": "learn_error", "message": str(exc)})

    def _set_learning_ui(self, learning: bool) -> None:
        state = "disabled" if learning else "normal"
        for widget in (self.add_btn, self.clear_btn, self.browse_btn, self.start_btn):
            widget.config(state=state)
        self.learn_btn.config(state=state)
        self.output_entry.config(state=state)
        if learning:
            self.open_btn.config(state="disabled")

    def _on_learn_ready(self, result) -> None:
        self.learning = False
        self._set_learning_ui(False)
        self._set_header_status("就绪")
        mapped = sum(1 for m in result.mappings if m.txt is not None)
        self._append_log(
            "INFO", "识别完成：表头 %d 行，数据 %d–%d 行，映射 %d 列，整体置信度 %s" % (
                result.header_rows, result.data_start, result.data_end, mapped,
                result.overall_confidence),
        )
        self.learn_dialog = LearnDialog(self, result)

    def _on_learn_error(self, message: str) -> None:
        self.learning = False
        self._set_learning_ui(False)
        self._set_header_status("学习失败")
        self._append_log("ERROR", "学习模板失败：%s" % message)
        messagebox.showerror("学习失败", "无法从该 Excel 推断模板：\n%s" % message,
                             parent=self.root)

    def on_template_saved(self, json_path: str, result) -> None:
        self.registry.reload()
        labels = self.template_labels()
        for row in self.rows:
            row.combo.configure(values=labels)
        self.template_label.config(text=self._template_summary())
        self._set_header_status("模板已保存")
        self._append_log(
            "OK", "新模板已保存：%s（%s）；下拉框已更新。" % (
                os.path.basename(json_path), result.sheet_name),
        )

    def browse_output(self) -> None:
        if self.running:
            return
        initial = self.output_var.get() or DEFAULT_OUTPUT
        path = filedialog.askdirectory(parent=self.root, title="选择输出目录", initialdir=initial)
        if path:
            self.output_var.set(os.path.abspath(path))

    # ------------------------------------------------------------------ 转换
    def start_conversion(self) -> None:
        if self.running:
            return
        if not self.rows:
            messagebox.showinfo("没有文件", "请先添加至少一个 TXT 文件。", parent=self.root)
            return
        output_dir = self.output_var.get().strip()
        if not output_dir:
            messagebox.showinfo("缺少输出目录", "请先选择输出目录。", parent=self.root)
            return
        output_dir = os.path.abspath(output_dir)
        self.output_var.set(output_dir)
        self.last_output_dir = output_dir
        self.running = True
        self.summary = None
        self._set_running_ui(True)
        self.overall_bar["value"] = 0
        self.overall_label.config(text="0 / %d" % len(self.rows))
        for row in self.rows:
            row.set_status("待处理", "等待开始", 0.0)
            row.set_running(True)
        self._append_log("INFO", "开始转换：%d 个文件，输出到 %s" % (len(self.rows), output_dir))

        files = [row.path for row in self.rows]
        overrides = {}
        for row in self.rows:
            if row.template_id:
                overrides[row.path] = row.template_id
        self._worker = threading.Thread(
            target=self._run_worker, args=(files, overrides, output_dir), daemon=True
        )
        self._worker.start()

    def _run_worker(self, files: List[str], overrides: Dict[str, str], output_dir: str) -> None:
        try:
            summary = run_conversion(
                files, overrides, output_dir, self.registry,
                logger_callback=lambda event: self.queue.put(event),
                status_callback=lambda event: self.queue.put(event),
                progress_callback=lambda event: self.queue.put(event),
            )
            self.queue.put({"type": "summary", "summary": summary})
        except Exception as exc:
            self.queue.put({"type": "fatal", "message": str(exc)})

    def _poll_queue(self) -> None:
        finished = None
        try:
            while True:
                event = self.queue.get_nowait()
                kind = event.get("type")
                if kind == "log":
                    self._append_log(event.get("level", "INFO"), event.get("line", ""))
                elif kind == "file":
                    self._on_file_event(event)
                elif kind == "parsed":
                    row = self._row_for(event.get("path"))
                    if row:
                        row.set_status("解析完成", row.detail_label.cget("text"), 0.45)
                elif kind == "done":
                    row = self._row_for(event.get("path"))
                    if row:
                        row.set_status("成功", row.detail_label.cget("text"), 1.0)
                elif kind == "summary":
                    finished = event.get("summary")
                elif kind == "learn_ready":
                    self._on_learn_ready(event.get("result"))
                elif kind == "learn_error":
                    self._on_learn_error(event.get("message", "未知错误"))
                elif kind == "fatal":
                    self._append_log("ERROR", "转换线程异常：%s" % event.get("message", ""))
                    self._finish_run(None, fatal=True)
        except queue.Empty:
            pass
        if finished is not None:
            self._finish_run(finished, fatal=False)
        self._refresh_overall()
        self.root.after(80, self._poll_queue)

    def _row_for(self, path: Optional[str]) -> Optional[FileRow]:
        if not path:
            return None
        return self.rows_by_path.get(os.path.normcase(os.path.abspath(path)))

    def _on_file_event(self, event: dict) -> None:
        row = self._row_for(event.get("path"))
        if not row:
            return
        status = event.get("status", "")
        detail = event.get("detail", "")
        progress = event.get("progress")
        if status == "成功":
            detail = "已生成 %s" % os.path.basename(event.get("output", ""))
        elif status == "失败":
            detail = event.get("detail") or "转换失败，详见日志"
        row.set_status(status, detail, progress)
        if status == "失败":
            self._append_log("ERROR", "%s：%s" % (os.path.basename(row.path), detail))

    def _refresh_overall(self) -> None:
        if not self.rows:
            return
        if self.summary is not None and not self.running:
            return
        total = sum(row.progress_value for row in self.rows) / (100.0 * len(self.rows))
        self.overall_bar["value"] = total * 100.0
        done = sum(1 for row in self.rows if row.status_label.cget("text") in ("成功", "失败"))
        self.overall_label.config(text="%d / %d" % (done, len(self.rows)))
        if self.running:
            self._set_header_status("转换中 %d%%" % int(total * 100))

    def _finish_run(self, summary: Optional[RunSummary], fatal: bool) -> None:
        self.running = False
        for row in self.rows:
            row.set_running(False)
        self._set_running_ui(False)
        if fatal:
            self._set_header_status("已中止")
            messagebox.showerror("转换失败", "转换线程异常，请查看日志。", parent=self.root)
            return
        if summary is None:
            return
        self.summary = summary
        self._append_log(
            "OK" if summary.failed == 0 else "WARN",
            "结束：成功 %d，失败 %d，输出 %d 个工作簿。" % (
                summary.succeeded, summary.failed, len(summary.outputs)),
        )
        self.log_path_label.config(text=os.path.basename(summary.log_path))
        self.open_btn.config(state="normal")
        self._set_header_status("完成" if summary.failed == 0 else "部分失败")
        self.overall_bar["value"] = 100.0 * summary.succeeded / max(1, summary.total)
        self.overall_label.config(text="%d / %d" % (summary.succeeded, summary.total))
        if summary.outputs:
            detail = "\n".join(os.path.basename(p) for p in summary.outputs)
            title = "转换完成" if summary.failed == 0 else "转换完成（存在失败）"
            message = "成功 %d 个文件，失败 %d 个。\n\n输出工作簿：\n%s\n\n目录：%s" % (
                summary.succeeded, summary.failed, detail, self.last_output_dir)
            if summary.failed == 0:
                messagebox.showinfo(title, message, parent=self.root)
            else:
                messagebox.showwarning(title, message + "\n\n失败原因见日志面板。", parent=self.root)
        else:
            messagebox.showerror("转换失败", "没有生成任何工作簿，请查看日志。", parent=self.root)

    def _set_running_ui(self, running: bool) -> None:
        self.add_btn.config(state="disabled" if running else "normal")
        self.clear_btn.config(state="disabled" if running else "normal")
        self.browse_btn.config(state="disabled" if running else "normal")
        self.output_entry.config(state="disabled" if running else "normal")
        self.start_btn.config(state="disabled" if running else "normal")
        self.learn_btn.config(state="disabled" if running else "normal")
        if running:
            self.open_btn.config(state="disabled")

    def open_output_dir(self) -> None:
        path = self.last_output_dir
        if os.path.isdir(path):
            try:
                os.startfile(path)
            except Exception as exc:
                messagebox.showinfo("无法打开目录", "请手动打开：%s\n%s" % (path, exc), parent=self.root)
        else:
            messagebox.showinfo("目录不存在", "输出目录还不存在：%s" % path, parent=self.root)

    # ------------------------------------------------------------------ 日志
    def _append_log(self, level: str, line: str) -> None:
        if not line:
            return
        tag = level if level in ("INFO", "OK", "WARN", "ERROR", "DEBUG") else "INFO"
        self.log_text.config(state="normal")
        self.log_text.insert("end", line + "\n", tag)
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def clear_log(self) -> None:
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")

    def _set_header_status(self, text: str) -> None:
        self.canvas.itemconfigure(self.status_item, text=text)

    def _on_mousewheel(self, event) -> None:
        widget = self.root.winfo_containing(event.x_root, event.y_root)
        if widget is None:
            return None
        target = widget
        while target is not None:
            if target is self.list_canvas or target is self.list_inner:
                self.list_canvas.yview_scroll(int(-event.delta / 120.0), "units")
                return "break"
            target = getattr(target, "master", None)
        return None

    def _on_close(self) -> None:
        if self.running:
            if not messagebox.askyesno("正在转换", "转换尚未结束，确定要退出吗？", parent=self.root):
                return
        self.root.destroy()


def launch() -> None:
    ensure_user_dirs()
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = tk.Tk()
    App(root)
    root.mainloop()
