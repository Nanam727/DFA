from __future__ import annotations

import csv
from datetime import datetime
import json
import re
import time
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

from src.align_icons import create_align_icons
from src.docx_handler import DocxHandler


ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class FormatApp(ctk.CTk):
    A4_WIDTH_MM = 210.0
    A4_HEIGHT_MM = 297.0
    A4_WIDTH_IN = A4_WIDTH_MM / 25.4
    A4_HEIGHT_IN = A4_HEIGHT_MM / 25.4
    FONT_OPTIONS = ["微软雅黑", "宋体", "Calibri", "Times New Roman", "黑体", "楷体", "楷体GB2312"]
    WORD_MARGIN_PRESETS_MM: dict[str, dict[str, float]] = {
        "普通": {"top": 25.4, "bottom": 25.4, "left": 31.8, "right": 31.8},
        "窄": {"top": 12.7, "bottom": 12.7, "left": 12.7, "right": 12.7},
        "宽": {"top": 25.4, "bottom": 25.4, "left": 50.8, "right": 50.8},
    }
    WORD_DEFAULT_FONT_PT = {
        "body": 10.5,
        "h1": 22.0,
        "h2": 15.0,
        "h3": 14.0,
        "h4": 12.0,
        "h5": 10.5,
    }
    PREVIEW_BASELINE_CHARS_PER_LINE = 42
    PREVIEW_BASELINE_LINES_PER_PAGE = 46
    PREVIEW_SAFE_EDGE_MM = {
        "left": 3.0,
        "right": 3.0,
        "top": 2.2,
        "bottom": 2.2,
    }
    PREVIEW_SAFE_CHAR_BUFFER = 1.8
    PREVIEW_TEXT_INNER_PAD_X = 2
    PREVIEW_TEXT_INNER_PAD_Y = 2
    UI_BG = "#f4f7fb"
    UI_CARD_BG = "#ffffff"
    UI_BORDER = "#d8e1ef"
    UI_ACCENT = "#2f5fa7"
    UI_ACCENT_HOVER = "#244b83"
    UI_TEXT = "#1f2a3a"
    PREVIEW_MODE_OPTIONS = ["快速模式", "高精度模式"]
    BASELINE_NORMAL_CONTENT_MM = {
        "width": A4_WIDTH_MM - WORD_MARGIN_PRESETS_MM["普通"]["left"] - WORD_MARGIN_PRESETS_MM["普通"]["right"],
        "height": A4_HEIGHT_MM - WORD_MARGIN_PRESETS_MM["普通"]["top"] - WORD_MARGIN_PRESETS_MM["普通"]["bottom"],
    }
    LINE_START_FORBIDDEN = set("，。、；：？！’”）】》")
    LINE_END_FORBIDDEN = set("‘“（【《")
    CJK_RE = re.compile(r"[\u4e00-\u9fff\u3400-\u4dbf\uff00-\uffef]")
    HALF_TO_FULL_PUNC = {
        ",": "，",
        ".": "。",
        ";": "；",
        ":": "：",
        "?": "？",
        "!": "！",
        "(": "（",
        ")": "）",
    }
    FULL_TO_HALF_PUNC = {v: k for k, v in HALF_TO_FULL_PUNC.items()}

    def __init__(self) -> None:
        super().__init__()
        self.title("Word 文档在线预览与格式调整工具（本地桌面版）")
        self.geometry("1300x860")
        self.minsize(1050, 700)
        self.configure(fg_color=self.UI_BG)

        # Reserved fields for next rounds (do not remove).
        self.docx_handler = DocxHandler()
        self.docx_handler.set_auto_backup_enabled(True)
        self.current_file: str | None = None
        self.selected_indices: set[int] = set()
        self.para_ranges: dict[int, tuple[tk.Text, str, str]] = {}
        self.left_para_ranges: dict[int, tuple[str, str]] = {}
        self.line_to_para_index_left: dict[int, int] = {}
        self.line_to_para_index_right: dict[tk.Text, dict[int, int]] = {}
        self.paragraph_data: list[dict[str, object]] = []
        self.last_clicked_index: int | None = None
        self.current_view_page: int = 1
        self.total_view_pages: int = 1
        self.status_text: str = ""
        self.preview_margin_override_mm: dict[str, float] | None = None
        self.current_doc_margins_mm: dict[str, float] = dict(self.WORD_MARGIN_PRESETS_MM["普通"])
        self.current_margin_preset_name: str = "普通"
        self.preview_zoom_percent: int = 100
        self.external_preview_content: str = ""
        self.toc_levels: int = 3
        self.default_school_template: str = "通用模板"
        self.operation_logs: list[dict[str, str]] = []
        self.preference_file = Path.cwd() / ".adfa_user_prefs.json"

        self.align_icons: dict[str, ctk.CTkImage] = {}
        self.btn_style_bold: ctk.CTkButton | None = None
        self.btn_style_italic: ctk.CTkButton | None = None
        self.btn_style_underline: ctk.CTkButton | None = None
        self.btn_align_left: ctk.CTkButton | None = None
        self.btn_align_center: ctk.CTkButton | None = None
        self.btn_align_right: ctk.CTkButton | None = None
        self.btn_align_justify: ctk.CTkButton | None = None
        self.preview_mode_combo: ctk.CTkComboBox | None = None
        self.auto_backup_switch: ctk.CTkSwitch | None = None
        self.tool_group_vars: dict[str, tk.BooleanVar] = {}
        self.tool_group_frames: dict[str, list[ctk.CTkFrame]] = {}
        self.active_bold = False
        self.active_italic = False
        self.active_underline = False
        self.active_alignment = ""
        self.page_window_ids: list[int | None] = [None, None]
        self.page_rect_ids: list[int | None] = [None, None]
        self.page_frames: list[tk.Frame] = []
        self.page_inners: list[tk.Frame] = []
        self.preview_text_widgets: list[tk.Text] = []
        self.find_match_indices: list[int] = []
        self.find_match_cursor: int = -1
        self.find_last_signature: tuple[str, bool, bool, bool] | None = None
        self.preview_accuracy_mode: str = "快速模式"
        self.auto_backup_enabled: bool = True
        self.field_update_tip_shown: bool = False
        self.btn_undo: ctk.CTkButton | None = None
        self.btn_redo: ctk.CTkButton | None = None

        self._build_ui()
        self._load_user_preferences()
        self.bind_all("<Control-a>", self.on_select_all_shortcut)
        self.bind_all("<Control-A>", self.on_select_all_shortcut)
        self.bind_all("<Control-s>", self.on_save_shortcut)
        self.bind_all("<Control-S>", self.on_save_shortcut)
        self.bind_all("<Control-d>", self.on_cancel_selection_shortcut)
        self.bind_all("<Control-D>", self.on_cancel_selection_shortcut)
        self.bind_all("<Control-z>", self.on_undo_shortcut)
        self.bind_all("<Control-Z>", self.on_undo_shortcut)
        self.bind_all("<Control-y>", self.on_redo_shortcut)
        self.bind_all("<Control-Y>", self.on_redo_shortcut)
        self.bind_all("<Control-x>", self.on_refresh_shortcut)
        self.bind_all("<Control-X>", self.on_refresh_shortcut)

    def _build_ui(self) -> None:
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._build_top_bar()
        self._build_main_area()

    def _build_top_bar(self) -> None:
        top = ctk.CTkFrame(
            self,
            corner_radius=10,
            fg_color=self.UI_CARD_BG,
            border_width=1,
            border_color=self.UI_BORDER,
        )
        top.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 6))
        top.grid_columnconfigure(5, weight=1)

        self.btn_open = ctk.CTkButton(
            top,
            text="打开文档",
            width=110,
            command=self.on_open_document,
            corner_radius=8,
            fg_color=self.UI_ACCENT,
            hover_color=self.UI_ACCENT_HOVER,
        )
        self.btn_open.grid(row=0, column=0, padx=(10, 8), pady=10)

        self.btn_save = ctk.CTkButton(
            top,
            text="保存",
            width=90,
            command=self.on_save_document,
            corner_radius=8,
            fg_color=self.UI_ACCENT,
            hover_color=self.UI_ACCENT_HOVER,
        )
        self.btn_save.grid(row=0, column=1, padx=8, pady=10)

        self.btn_save_as = ctk.CTkButton(
            top,
            text="另存为",
            width=90,
            command=self.on_save_as_document,
            corner_radius=8,
            fg_color=self.UI_ACCENT,
            hover_color=self.UI_ACCENT_HOVER,
        )
        self.btn_save_as.grid(row=0, column=2, padx=8, pady=10)

        label_title = ctk.CTkLabel(
            top, text="当前文件名：", font=ctk.CTkFont(size=14, weight="bold"), text_color=self.UI_TEXT
        )
        label_title.grid(row=0, column=3, padx=(20, 5), pady=10, sticky="w")

        self.file_label = ctk.CTkLabel(top, text="未打开文档", anchor="w", text_color=self.UI_TEXT)
        self.file_label.grid(row=0, column=4, padx=(0, 10), pady=10, sticky="w")
        self.shortcut_hint_label = ctk.CTkLabel(
            top,
            text="快捷键：Ctrl+A 全选 | Ctrl+S 保存 | Ctrl+Z 撤销 | Ctrl+Y 重做 | Ctrl+X 刷新 | Ctrl+D 取消选取 | 右键局部取消",
            anchor="e",
            text_color="#4a5a72",
        )
        self.shortcut_hint_label.grid(row=0, column=5, padx=(10, 12), pady=10, sticky="e")

    def _build_main_area(self) -> None:
        main = ctk.CTkFrame(self, corner_radius=8, fg_color="transparent")
        main.grid(row=1, column=0, sticky="nsew", padx=12, pady=(6, 12))
        main.grid_columnconfigure(0, weight=3)
        main.grid_columnconfigure(1, weight=1)
        main.grid_rowconfigure(0, weight=1)

        self._build_left_panel(main)
        self._build_right_panel(main)

    def _build_left_panel(self, parent: ctk.CTkFrame) -> None:
        left = ctk.CTkFrame(parent, corner_radius=10, fg_color=self.UI_CARD_BG, border_width=1, border_color=self.UI_BORDER)
        left.grid(row=0, column=0, sticky="nsew", padx=(8, 6), pady=8)
        left.grid_columnconfigure(0, weight=1)
        left.grid_rowconfigure(1, weight=1)

        action_bar = ctk.CTkFrame(left, fg_color="transparent")
        action_bar.grid(row=0, column=0, sticky="ew", padx=10, pady=(6, 8))
        # 合并“选取/全选/取消”和“分类按钮”为一行，取消固定最右。
        action_bar.grid_columnconfigure(12, weight=1)

        self.btn_pick = ctk.CTkButton(action_bar, text="选取", width=84, command=self.on_select_pick, corner_radius=8)
        self.btn_pick.grid(row=0, column=0, padx=(2, 4), pady=4, sticky="w")
        self.btn_all = ctk.CTkButton(action_bar, text="全选", width=84, command=self.on_select_all, corner_radius=8)
        self.btn_all.grid(row=0, column=1, padx=4, pady=4, sticky="w")

        categories = [
            ("大标题", self.on_mark_title),
            ("一级", self.on_mark_h1),
            ("二级", self.on_mark_h2),
            ("三级", self.on_mark_h3),
            ("四级", self.on_mark_h4),
            ("正文", self.on_mark_body),
        ]
        for idx, (text, cmd) in enumerate(categories, start=2):
            btn = ctk.CTkButton(action_bar, text=text, width=82, command=cmd, corner_radius=8)
            btn.grid(row=0, column=idx, padx=3, pady=4, sticky="w")

        ctk.CTkButton(action_bar, text="批量选择", width=90, command=self.on_select_batch, corner_radius=8).grid(
            row=0, column=8, padx=(6, 3), pady=4, sticky="w"
        )
        ctk.CTkButton(action_bar, text="查找", width=82, command=self.on_find_dialog, corner_radius=8).grid(
            row=0, column=9, padx=(4, 3), pady=4, sticky="w"
        )
        ctk.CTkButton(action_bar, text="替代", width=82, command=self.on_replace_dialog, corner_radius=8).grid(
            row=0, column=10, padx=(3, 3), pady=4, sticky="w"
        )
        self.btn_undo = ctk.CTkButton(action_bar, text="撤销(0)", width=84, command=self.on_undo_click, corner_radius=8)
        self.btn_undo.grid(row=0, column=11, padx=(3, 3), pady=4, sticky="w")
        self.btn_redo = ctk.CTkButton(action_bar, text="重做(0)", width=84, command=self.on_redo_click, corner_radius=8)
        self.btn_redo.grid(row=0, column=12, padx=(3, 3), pady=4, sticky="w")
        self.btn_cancel = ctk.CTkButton(action_bar, text="取消", width=84, command=self.on_select_cancel, corner_radius=8)
        self.btn_cancel.grid(row=0, column=13, padx=(8, 2), pady=4, sticky="e")

        preview_wrap = ctk.CTkFrame(left, fg_color="#eef3fb", corner_radius=10)
        preview_wrap.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        preview_wrap.grid_rowconfigure(0, weight=1)
        preview_wrap.grid_columnconfigure(0, weight=0)
        preview_wrap.grid_columnconfigure(1, weight=1)

        # 标注层放在 A4 页面左侧，不占用白纸区域。
        left_mark_frame = tk.Frame(preview_wrap, bg="#e6edf8")
        left_mark_frame.grid(row=0, column=0, sticky="ns", padx=(0, 10))

        self.preview_canvas = tk.Canvas(preview_wrap, bg="#e6edf8", highlightthickness=0)
        self.preview_canvas.grid(row=0, column=1, sticky="nsew")
        self.preview_canvas.bind("<Configure>", self._on_preview_resize)
        self.preview_canvas.bind("<MouseWheel>", self._on_preview_mousewheel)
        self.preview_canvas.bind("<Button-4>", self._on_preview_mousewheel)
        self.preview_canvas.bind("<Button-5>", self._on_preview_mousewheel)

        self.left_class_text = tk.Text(
            left_mark_frame,
            width=16,
            wrap="word",
            bg="#f1f5fc",
            relief="flat",
            bd=0,
            highlightthickness=0,
            font=("Microsoft YaHei UI", 10),
        )
        self.left_class_text.pack(side="left", fill="both", expand=True)
        self.left_class_text.bind("<Button-1>", self.on_preview_click)
        self.left_class_text.bind("<Control-Button-1>", self.on_preview_click)
        self.left_class_text.bind("<Button-3>", self.on_preview_right_click)
        self.left_class_text.bind("<MouseWheel>", self._on_preview_mousewheel)
        self.left_class_text.bind("<Button-4>", self._on_preview_mousewheel)
        self.left_class_text.bind("<Button-5>", self._on_preview_mousewheel)
        for _ in range(2):
            page_frame = tk.Frame(self.preview_canvas, bg="#ffffff", bd=0, highlightthickness=0)
            page_inner = tk.Frame(page_frame, bg="#ffffff")
            page_inner.pack(fill="both", expand=True, padx=0, pady=0)
            page_inner.grid_rowconfigure(0, weight=1)
            page_inner.grid_columnconfigure(0, weight=1)

            text = tk.Text(
                page_inner,
                wrap="none",
                bg="#ffffff",
                relief="flat",
                bd=0,
                highlightthickness=0,
                padx=self.PREVIEW_TEXT_INNER_PAD_X,
                pady=self.PREVIEW_TEXT_INNER_PAD_Y,
                font=("Microsoft YaHei UI", 11),
            )
            text.grid(row=0, column=0, sticky="nsew")
            text.bind("<Button-1>", self.on_preview_click)
            text.bind("<Control-Button-1>", self.on_preview_click)
            text.bind("<Button-3>", self.on_preview_right_click)
            text.bind("<MouseWheel>", self._on_preview_mousewheel)
            text.bind("<Button-4>", self._on_preview_mousewheel)
            text.bind("<Button-5>", self._on_preview_mousewheel)

            self.page_frames.append(page_frame)
            self.page_inners.append(page_inner)
            self.preview_text_widgets.append(text)

        pager_bar = ctk.CTkFrame(left, fg_color="transparent")
        pager_bar.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 8))
        pager_bar.grid_columnconfigure(1, weight=1)

        self.btn_prev_page = ctk.CTkButton(pager_bar, text="◀", width=44, command=self.on_prev_page, corner_radius=8)
        self.btn_prev_page.grid(row=0, column=0, padx=(0, 8), pady=2, sticky="w")

        self.page_info_label = ctk.CTkLabel(pager_bar, text="第 1 / 1 页", anchor="center")
        self.page_info_label.grid(row=0, column=1, padx=4, pady=2, sticky="ew")

        self.btn_next_page = ctk.CTkButton(pager_bar, text="▶", width=44, command=self.on_next_page, corner_radius=8)
        self.btn_next_page.grid(row=0, column=2, padx=(8, 0), pady=2, sticky="e")

        self._insert_preview_placeholder()

    def _build_right_panel(self, parent: ctk.CTkFrame) -> None:
        right = ctk.CTkFrame(parent, corner_radius=10, fg_color=self.UI_CARD_BG, border_width=1, border_color=self.UI_BORDER)
        right.grid(row=0, column=1, sticky="nsew", padx=(6, 8), pady=6)
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(1, weight=1)

        tools_title = ctk.CTkLabel(
            right, text="格式工具栏", font=ctk.CTkFont(size=15, weight="bold"), text_color=self.UI_TEXT
        )
        tools_title.grid(row=0, column=0, sticky="w", padx=12, pady=(12, 8))

        tools_scroll_wrap = ctk.CTkFrame(right, fg_color="transparent")
        tools_scroll_wrap.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))
        tools_scroll_wrap.grid_rowconfigure(0, weight=1)
        tools_scroll_wrap.grid_columnconfigure(0, weight=1)

        self.right_tools_canvas = tk.Canvas(tools_scroll_wrap, bg="#f2f6fd", highlightthickness=0)
        self.right_tools_canvas.grid(row=0, column=0, sticky="nsew")
        self.right_tools_scrollbar = ctk.CTkScrollbar(
            tools_scroll_wrap, orientation="vertical", command=self.right_tools_canvas.yview
        )
        self.right_tools_scrollbar.grid(row=0, column=1, sticky="ns", padx=(6, 0))
        self.right_tools_canvas.configure(yscrollcommand=self.right_tools_scrollbar.set)

        self.right_tools_inner = ctk.CTkFrame(self.right_tools_canvas, fg_color="transparent")
        self.right_tools_window = self.right_tools_canvas.create_window(
            0, 0, anchor="nw", window=self.right_tools_inner
        )
        self.right_tools_inner.grid_columnconfigure(0, weight=1)
        self.right_tools_canvas.bind("<Configure>", self._on_right_tools_canvas_configure)
        self.right_tools_inner.bind("<Configure>", self._on_right_tools_inner_configure)
        self._bind_right_tools_mousewheel(self.right_tools_inner)

        group_bar = ctk.CTkFrame(self.right_tools_inner)
        group_bar.grid(row=0, column=0, sticky="ew", padx=10, pady=(6, 2))
        group_bar.grid_columnconfigure((0, 1), weight=1)
        group_defs = [
            ("base", "基础"),
            ("page_toc", "页码/目录"),
            ("refs", "参考文献"),
            ("clean", "清理"),
        ]
        for i, (key, label) in enumerate(group_defs):
            var = tk.BooleanVar(value=True)
            self.tool_group_vars[key] = var
            ctk.CTkCheckBox(
                group_bar,
                text=label,
                variable=var,
                command=self._apply_tool_group_visibility,
                checkbox_width=16,
                checkbox_height=16,
            ).grid(row=i // 2, column=i % 2, sticky="w", padx=4, pady=2)

        font_size_row = ctk.CTkFrame(self.right_tools_inner)
        font_size_row.grid(row=1, column=0, sticky="ew", padx=10, pady=6)
        font_size_row.grid_columnconfigure((0, 1), weight=1)
        font_box = ctk.CTkFrame(font_size_row)
        font_box.grid(row=0, column=0, sticky="ew", padx=(0, 4), pady=0)
        font_box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(font_box, text="字体").grid(row=0, column=0, sticky="w", padx=8, pady=(8, 4))
        self.font_combo = ctk.CTkComboBox(font_box, values=self.FONT_OPTIONS, command=self.on_font_option_change)
        self.font_combo.set("宋体")
        self.font_combo.grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 8))
        self.font_combo.bind("<KeyRelease>", self.on_font_option_keyup)

        size_box = ctk.CTkFrame(font_size_row)
        size_box.grid(row=0, column=1, sticky="ew", padx=(4, 0), pady=0)
        size_box.grid_columnconfigure(0, weight=1)
        size_head = ctk.CTkFrame(size_box, fg_color="transparent")
        size_head.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 4))
        size_head.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(size_head, text="字号").grid(row=0, column=0, sticky="w")
        ctk.CTkButton(
            size_head,
            text="-",
            width=26,
            height=24,
            corner_radius=7,
            command=self.on_decrease_font_size,
            fg_color="#e9eef8",
            hover_color="#dde6f5",
            text_color=self.UI_TEXT,
        ).grid(row=0, column=1, padx=(0, 4))
        ctk.CTkButton(
            size_head,
            text="+",
            width=26,
            height=24,
            corner_radius=7,
            command=self.on_increase_font_size,
            fg_color="#e9eef8",
            hover_color="#dde6f5",
            text_color=self.UI_TEXT,
        ).grid(row=0, column=2)
        self.size_combo = ctk.CTkComboBox(
            size_box,
            values=[
                "32",
                "28",
                "24",
                "二号",
                "22",
                "小二号",
                "18",
                "三号",
                "16",
                "小三号",
                "15",
                "四号",
                "14",
                "小四号",
                "12",
                "五号",
                "10.5",
            ],
            command=self.on_size_option_change,
        )
        self.size_combo.set("五号")
        self.size_combo.grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 6))
        self.size_combo.bind("<KeyRelease>", self.on_size_option_keyup)

        margin_spacing_row = ctk.CTkFrame(self.right_tools_inner)
        margin_spacing_row.grid(row=2, column=0, sticky="ew", padx=10, pady=6)
        margin_spacing_row.grid_columnconfigure((0, 1), weight=1)
        margin_box = ctk.CTkFrame(margin_spacing_row)
        margin_box.grid(row=0, column=0, sticky="ew", padx=(0, 4), pady=0)
        margin_box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(margin_box, text="页边距").grid(row=0, column=0, sticky="w", padx=8, pady=(8, 4))
        self.margin_combo = ctk.CTkComboBox(
            margin_box,
            values=["普通", "窄", "宽"],
            command=self.on_margin_option_change,
        )
        self.margin_combo.set("普通")
        self.margin_combo.grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 8))
        self.margin_combo.bind("<KeyRelease>", self.on_margin_option_keyup)

        spacing_box = ctk.CTkFrame(margin_spacing_row)
        spacing_box.grid(row=0, column=1, sticky="ew", padx=(4, 0), pady=0)
        spacing_box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(spacing_box, text="行间距").grid(row=0, column=0, sticky="w", padx=8, pady=(8, 4))
        self.spacing_combo = ctk.CTkComboBox(
            spacing_box, values=["单倍(1.0)", "1.5倍", "2倍", "固定18pt"], command=self.on_line_spacing_option_change
        )
        self.spacing_combo.set("1.5倍")
        self.spacing_combo.grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 8))
        self.spacing_combo.bind("<KeyRelease>", self.on_line_spacing_option_keyup)

        style_box = ctk.CTkFrame(self.right_tools_inner)
        style_box.grid(row=3, column=0, sticky="ew", padx=10, pady=6)
        style_box.grid_columnconfigure((0, 1, 2), weight=1)
        self.btn_style_bold = ctk.CTkButton(
            style_box,
            text="B",
            width=34,
            fg_color="transparent",
            hover_color="#e9eef6",
            text_color="#202020",
            border_width=1,
            border_color="#c7c7c7",
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.on_toggle_bold,
        )
        self.btn_style_bold.grid(row=0, column=0, sticky="ew", padx=(8, 4), pady=8)
        self.btn_style_italic = ctk.CTkButton(
            style_box,
            text="I",
            width=34,
            fg_color="transparent",
            hover_color="#e9eef6",
            text_color="#202020",
            border_width=1,
            border_color="#c7c7c7",
            font=ctk.CTkFont(size=14, slant="italic"),
            command=self.on_toggle_italic,
        )
        self.btn_style_italic.grid(row=0, column=1, sticky="ew", padx=4, pady=8)
        self.btn_style_underline = ctk.CTkButton(
            style_box,
            text="U",
            width=34,
            fg_color="transparent",
            hover_color="#e9eef6",
            text_color="#202020",
            border_width=1,
            border_color="#c7c7c7",
            font=ctk.CTkFont(size=14, underline=True),
            command=self.on_toggle_underline,
        )
        self.btn_style_underline.grid(row=0, column=2, sticky="ew", padx=(4, 8), pady=8)

        align_box = ctk.CTkFrame(self.right_tools_inner)
        align_box.grid(row=4, column=0, sticky="ew", padx=10, pady=6)
        align_box.grid_columnconfigure((0, 1, 2, 3), weight=1)
        ctk.CTkLabel(align_box, text="对齐").grid(row=0, column=0, columnspan=4, sticky="w", padx=8, pady=(8, 2))
        self.align_icons = create_align_icons(size=16)
        self.btn_align_left = ctk.CTkButton(
            align_box,
            text="",
            image=self.align_icons["left"],
            command=self.on_align_left,
            width=30,
            fg_color="transparent",
            hover_color="#eef2f7",
            border_width=1,
            border_color="#d8d8d8",
            corner_radius=6,
        )
        self.btn_align_left.grid(row=1, column=0, sticky="ew", padx=(8, 4), pady=(0, 8))
        self.btn_align_center = ctk.CTkButton(
            align_box,
            text="",
            image=self.align_icons["center"],
            command=self.on_align_center,
            width=30,
            fg_color="transparent",
            hover_color="#eef2f7",
            border_width=1,
            border_color="#d8d8d8",
            corner_radius=6,
        )
        self.btn_align_center.grid(row=1, column=1, sticky="ew", padx=4, pady=(0, 8))
        self.btn_align_right = ctk.CTkButton(
            align_box,
            text="",
            image=self.align_icons["right"],
            command=self.on_align_right,
            width=30,
            fg_color="transparent",
            hover_color="#eef2f7",
            border_width=1,
            border_color="#d8d8d8",
            corner_radius=6,
        )
        self.btn_align_right.grid(row=1, column=2, sticky="ew", padx=4, pady=(0, 8))
        self.btn_align_justify = ctk.CTkButton(
            align_box,
            text="",
            image=self.align_icons["justify"],
            command=self.on_align_justify,
            width=30,
            fg_color="transparent",
            hover_color="#eef2f7",
            border_width=1,
            border_color="#d8d8d8",
            corner_radius=6,
        )
        self.btn_align_justify.grid(row=1, column=3, sticky="ew", padx=(4, 8), pady=(0, 8))

        indent_box = ctk.CTkFrame(self.right_tools_inner)
        indent_box.grid(row=5, column=0, sticky="ew", padx=10, pady=6)
        indent_box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(indent_box, text="段落版式").grid(row=0, column=0, sticky="w", padx=8, pady=(8, 6))
        ctk.CTkButton(indent_box, text="全文首行缩进2字符", command=self.on_apply_first_line_indent).grid(
            row=1, column=0, sticky="ew", padx=8, pady=(0, 8)
        )

        page_box = ctk.CTkFrame(self.right_tools_inner)
        page_box.grid(row=6, column=0, sticky="ew", padx=10, pady=6)
        page_box.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkLabel(page_box, text="页码 / 页眉页脚").grid(row=0, column=0, columnspan=2, sticky="w", padx=8, pady=(8, 4))
        self.header_entry = ctk.CTkEntry(page_box, placeholder_text="页眉文本（可选）")
        self.header_entry.grid(row=1, column=0, columnspan=2, sticky="ew", padx=8, pady=(0, 6))
        ctk.CTkButton(page_box, text="应用页眉", command=self.on_apply_header_text).grid(
            row=2, column=0, sticky="ew", padx=(8, 4), pady=4
        )
        ctk.CTkButton(page_box, text="页码从1开始", command=self.on_apply_footer_page_number).grid(
            row=2, column=1, sticky="ew", padx=(4, 8), pady=4
        )
        ctk.CTkButton(page_box, text="断开页眉页脚链接", command=self.on_unlink_header_footer).grid(
            row=3, column=0, columnspan=2, sticky="ew", padx=8, pady=4
        )
        ctk.CTkButton(page_box, text="页码格式切换：罗马", command=self.on_switch_page_number_roman).grid(
            row=4, column=0, columnspan=2, sticky="ew", padx=8, pady=4
        )
        ctk.CTkButton(page_box, text="页码格式切换：阿拉伯", command=self.on_switch_page_number_arabic).grid(
            row=5, column=0, columnspan=2, sticky="ew", padx=8, pady=4
        )
        ctk.CTkButton(page_box, text="一键分节并设置页码", command=self.on_one_click_section_and_page_number).grid(
            row=6, column=0, columnspan=2, sticky="ew", padx=8, pady=4
        )
        ctk.CTkButton(page_box, text="清空页眉页脚", command=self.on_clear_header_footer).grid(
            row=7, column=0, columnspan=2, sticky="ew", padx=8, pady=(4, 8)
        )

        toc_box = ctk.CTkFrame(self.right_tools_inner)
        toc_box.grid(row=7, column=0, sticky="ew", padx=10, pady=6)
        toc_box.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkLabel(toc_box, text="目录工具").grid(row=0, column=0, columnspan=2, sticky="w", padx=8, pady=(8, 4))
        ctk.CTkButton(toc_box, text="自动生成目录", command=self.on_generate_toc).grid(
            row=1, column=0, columnspan=2, sticky="ew", padx=8, pady=4
        )
        ctk.CTkButton(toc_box, text="更新目录（仅页码）", command=self.on_update_toc_page_only).grid(
            row=2, column=0, columnspan=2, sticky="ew", padx=8, pady=4
        )
        ctk.CTkButton(toc_box, text="更新目录（整个目录）", command=self.on_update_toc_all).grid(
            row=3, column=0, columnspan=2, sticky="ew", padx=8, pady=4
        )
        ctk.CTkButton(toc_box, text="目录-页码一致性校验", command=self.on_check_toc_consistency).grid(
            row=4, column=0, columnspan=2, sticky="ew", padx=8, pady=(4, 8)
        )

        ref_box = ctk.CTkFrame(self.right_tools_inner)
        ref_box.grid(row=8, column=0, sticky="ew", padx=10, pady=6)
        ref_box.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkLabel(ref_box, text="参考文献工具（GB/T 7714）").grid(row=0, column=0, columnspan=2, sticky="w", padx=8, pady=(8, 4))
        self.ref_mode_combo = ctk.CTkComboBox(ref_box, values=["顺序编码制", "著者-出版年制"])
        self.ref_mode_combo.set("顺序编码制")
        self.ref_mode_combo.grid(row=1, column=0, columnspan=2, sticky="ew", padx=8, pady=(0, 6))
        ctk.CTkButton(ref_box, text="格式化参考文献", command=self.on_format_references_gbt).grid(
            row=2, column=0, sticky="ew", padx=(8, 4), pady=4
        )
        ctk.CTkButton(ref_box, text="悬挂缩进一键设置", command=self.on_apply_reference_hanging_indent).grid(
            row=2, column=1, sticky="ew", padx=(4, 8), pady=4
        )
        ctk.CTkButton(ref_box, text="文献全面校验", command=self.on_validate_references_gbt).grid(
            row=3, column=0, columnspan=2, sticky="ew", padx=8, pady=(4, 8)
        )
        ctk.CTkButton(ref_box, text="一键修改（校验后使用）", command=self.on_auto_fix_references_by_validation).grid(
            row=4, column=0, columnspan=2, sticky="ew", padx=8, pady=(0, 8)
        )

        trace_box = ctk.CTkFrame(self.right_tools_inner)
        trace_box.grid(row=9, column=0, sticky="ew", padx=10, pady=6)
        trace_box.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkLabel(trace_box, text="可追溯与可配置").grid(row=0, column=0, columnspan=2, sticky="w", padx=8, pady=(8, 4))
        ctk.CTkLabel(trace_box, text="预览模式").grid(row=1, column=0, sticky="w", padx=8, pady=(2, 2))
        self.preview_mode_combo = ctk.CTkComboBox(trace_box, values=self.PREVIEW_MODE_OPTIONS, command=self.on_preview_mode_change)
        self.preview_mode_combo.set(self.preview_accuracy_mode)
        self.preview_mode_combo.grid(row=1, column=1, sticky="ew", padx=(4, 8), pady=(2, 2))
        self.auto_backup_switch = ctk.CTkSwitch(
            trace_box,
            text="自动备份 .bak",
            command=self.on_toggle_auto_backup,
        )
        self.auto_backup_switch.select()
        self.auto_backup_switch.grid(row=2, column=0, columnspan=2, sticky="w", padx=8, pady=(2, 4))
        ctk.CTkButton(trace_box, text="操作日志查看 / 导出", command=self.on_operation_log_view_export).grid(
            row=3, column=0, columnspan=2, sticky="ew", padx=8, pady=4
        )
        ctk.CTkButton(trace_box, text="学校规则配置文件导入 / 导出", command=self.on_school_rules_import_export).grid(
            row=4, column=0, columnspan=2, sticky="ew", padx=8, pady=4
        )
        ctk.CTkButton(trace_box, text="个人偏好设置保存", command=self.on_save_user_preferences).grid(
            row=5, column=0, columnspan=2, sticky="ew", padx=8, pady=4
        )
        ctk.CTkButton(trace_box, text="恢复默认偏好", command=self.on_restore_default_preferences).grid(
            row=6, column=0, columnspan=2, sticky="ew", padx=8, pady=(2, 8)
        )

        clean_box = ctk.CTkFrame(self.right_tools_inner)
        clean_box.grid(row=10, column=0, sticky="ew", padx=10, pady=(6, 12))
        clean_box.grid_columnconfigure(0, weight=1)
        clean_box.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(clean_box, text="文本清理").grid(row=0, column=0, columnspan=2, sticky="w", padx=8, pady=(8, 6))
        ctk.CTkButton(clean_box, text="删除#", command=self.on_clean_hash).grid(row=1, column=0, sticky="ew", padx=(8, 4), pady=4)
        ctk.CTkButton(clean_box, text="删除*", command=self.on_clean_star).grid(row=1, column=1, sticky="ew", padx=(4, 8), pady=4)
        ctk.CTkButton(clean_box, text="删除占位.", command=self.on_clean_placeholder_dot).grid(
            row=2, column=0, sticky="ew", padx=(8, 4), pady=4
        )
        ctk.CTkButton(clean_box, text="保留数字间-", command=self.on_clean_keep_num_dash).grid(
            row=2, column=1, sticky="ew", padx=(4, 8), pady=4
        )
        ctk.CTkButton(clean_box, text="删除空段落", command=self.on_clean_empty).grid(
            row=3, column=0, columnspan=2, sticky="ew", padx=8, pady=(4, 8)
        )
        self.tool_group_frames = {
            "base": [font_size_row, margin_spacing_row, style_box, align_box, indent_box],
            "page_toc": [page_box, toc_box],
            "refs": [ref_box],
            "clean": [trace_box, clean_box],
        }
        self._apply_tool_group_visibility()
        self._refresh_style_toggle_buttons()
        self._refresh_align_buttons()
        self._bind_right_tools_mousewheel(self.right_tools_inner)

    def _on_right_tools_canvas_configure(self, event: tk.Event) -> None:
        if hasattr(self, "right_tools_window"):
            self.right_tools_canvas.itemconfigure(self.right_tools_window, width=max(120, event.width))

    def _on_right_tools_inner_configure(self, _event: tk.Event | None = None) -> None:
        if hasattr(self, "right_tools_canvas"):
            self.right_tools_canvas.configure(scrollregion=self.right_tools_canvas.bbox("all"))

    def _apply_tool_group_visibility(self) -> None:
        if not self.tool_group_frames:
            return
        for key, frames in self.tool_group_frames.items():
            visible = bool(self.tool_group_vars.get(key).get()) if key in self.tool_group_vars else True
            for frame in frames:
                if visible:
                    frame.grid()
                else:
                    frame.grid_remove()
        self._on_right_tools_inner_configure(None)

    def _bind_right_tools_mousewheel(self, widget: tk.Misc) -> None:
        widget.bind("<MouseWheel>", self._on_right_tools_mousewheel)
        widget.bind("<Button-4>", self._on_right_tools_mousewheel)
        widget.bind("<Button-5>", self._on_right_tools_mousewheel)
        for child in widget.winfo_children():
            self._bind_right_tools_mousewheel(child)

    def _on_right_tools_mousewheel(self, event: tk.Event | None = None) -> str:
        if event is None or not hasattr(self, "right_tools_canvas"):
            return "break"
        delta = getattr(event, "delta", 0)
        num = getattr(event, "num", 0)
        if delta < 0 or num == 5:
            self.right_tools_canvas.yview_scroll(1, "units")
        elif delta > 0 or num == 4:
            self.right_tools_canvas.yview_scroll(-1, "units")
        return "break"

    def _set_preview_widgets_state(self, state: str) -> None:
        self.left_class_text.configure(state=state)
        for text in self.preview_text_widgets:
            text.configure(state=state)

    def _insert_preview_placeholder(self) -> None:
        self._set_preview_widgets_state("normal")
        self.left_class_text.delete("1.0", "end")
        for text in self.preview_text_widgets:
            text.delete("1.0", "end")
        self.left_class_text.insert("1.0", "大标题\n一级\n二级\n三级\n四级\n正文\n正文\n")
        self.preview_text_widgets[0].insert(
            "1.0",
            "【预览占位】\n"
            "这是用于第一轮 UI 骨架演示的正文区域。\n\n"
            "点击任意段落位置，不会触发业务处理，只做占位交互。\n\n"
            "后续轮次将接入 docx 读取、选择段落与格式修改逻辑。\n",
        )
        self.preview_text_widgets[1].insert("1.0", "（右页占位）\n")
        self.left_class_text.tag_configure("selected", background="#ffe7a3")
        for text in self.preview_text_widgets:
            text.tag_configure("selected", background="#ffe7a3")
        self._set_preview_widgets_state("disabled")
        self._update_grid_info(self._get_active_margins_mm())

    def _on_preview_resize(self, event: tk.Event) -> None:
        width = max(120, event.width)
        height = max(120, event.height)
        self._draw_a4_page(width, height)

    def _draw_a4_page(self, canvas_w: int, canvas_h: int) -> None:
        ratio = self.A4_WIDTH_IN / self.A4_HEIGHT_IN
        gap = max(12, int(canvas_w * 0.02))
        available_w = max(240, canvas_w - 48)
        available_h = max(180, canvas_h - 36)
        by_width = int((available_w - gap) // 2)
        by_height = int(available_h * ratio)
        zoom_ratio = max(0.5, min(2.0, self.preview_zoom_percent / 100.0))
        page_w = max(100, int(min(by_width, by_height) * zoom_ratio))
        if page_w * 2 + gap > available_w:
            page_w = max(100, int((available_w - gap) / 2))
        if int(page_w / ratio) > available_h:
            page_w = max(100, int(available_h * ratio))
        page_h = int(page_w / ratio)

        total_w = page_w * 2 + gap
        start_x = (canvas_w - total_w) // 2
        y1 = (canvas_h - page_h) // 2
        margins = self._get_active_margins_mm()
        safe = self._get_preview_safe_edge_mm()
        margin_left = max(8, int(page_w * ((margins["left"] + safe["left"]) / self.A4_WIDTH_MM)))
        margin_right = max(8, int(page_w * ((margins["right"] + safe["right"]) / self.A4_WIDTH_MM)))
        margin_top = max(8, int(page_h * ((margins["top"] + safe["top"]) / self.A4_HEIGHT_MM)))
        margin_bottom = max(8, int(page_h * ((margins["bottom"] + safe["bottom"]) / self.A4_HEIGHT_MM)))
        self._update_grid_info(margins)
        # Keep the embedded paper window alive; only redraw page shapes.
        self.preview_canvas.delete("page-shape")
        for i in range(2):
            x1 = start_x + i * (page_w + gap)
            x2 = x1 + page_w
            y2 = y1 + page_h
            self.preview_canvas.create_rectangle(
                x1 + 6,
                y1 + 6,
                x2 + 6,
                y2 + 6,
                fill="#b6b6b6",
                outline="",
                tags="page-shape",
            )
            self.page_rect_ids[i] = self.preview_canvas.create_rectangle(
                x1,
                y1,
                x2,
                y2,
                fill="#ffffff",
                outline="#c9c9c9",
                width=2,
                tags="page-shape",
            )

            if self.page_window_ids[i] is None or not self.preview_canvas.type(self.page_window_ids[i]):
                self.page_window_ids[i] = self.preview_canvas.create_window(
                    x1,
                    y1,
                    window=self.page_frames[i],
                    anchor="nw",
                    width=max(100, page_w),
                    height=max(140, page_h),
                    tags="page-window",
                )
            else:
                self.preview_canvas.coords(self.page_window_ids[i], x1, y1)
                self.preview_canvas.itemconfigure(
                    self.page_window_ids[i],
                    width=max(100, page_w),
                    height=max(140, page_h),
                )
            self.page_inners[i].pack_configure(padx=(margin_left, margin_right), pady=(margin_top, margin_bottom))

    def _show_todo(self) -> None:
        messagebox.showinfo("提示", "功能待接入")

    def _block_scroll_event(self, _event: tk.Event | None = None) -> str:
        return "break"

    def _on_preview_mousewheel(self, event: tk.Event | None = None) -> str:
        if event is None:
            return "break"
        delta = getattr(event, "delta", 0)
        num = getattr(event, "num", 0)
        if delta < 0 or num == 5:
            self.on_next_page()
        elif delta > 0 or num == 4:
            self.on_prev_page()
        return "break"

    def _refresh_document_view(self, preserve_selection: bool = False) -> None:
        if self.external_preview_content.strip():
            self._refresh_external_view()
            return
        previous_selection = set(self.selected_indices) if preserve_selection else set()
        previous_clicked = self.last_clicked_index

        self._set_preview_widgets_state("normal")
        self.left_class_text.delete("1.0", "end")
        for text in self.preview_text_widgets:
            text.delete("1.0", "end")
            self._clear_para_style_tags(text)
        self.left_para_ranges.clear()
        self.para_ranges.clear()
        self.line_to_para_index_left.clear()
        self.line_to_para_index_right = {text: {} for text in self.preview_text_widgets}

        self.paragraph_data = self.docx_handler.get_paragraphs()
        max_page = 1
        if self.paragraph_data:
            max_page = max(1, max(int(para["page_est"]) for para in self.paragraph_data))
        self.total_view_pages = max_page
        self.current_view_page = min(max(1, self.current_view_page), self._max_pair_start())
        self._update_page_indicator()

        show_pages = [self.current_view_page]
        if self.current_view_page + 1 <= self.total_view_pages:
            show_pages.append(self.current_view_page + 1)
        page_to_widget: dict[int, tk.Text] = {show_pages[0]: self.preview_text_widgets[0]}
        if len(show_pages) > 1:
            page_to_widget[show_pages[1]] = self.preview_text_widgets[1]

        if not self.paragraph_data:
            self.left_class_text.insert("1.0", "P1 [正文]\n")
            self.preview_text_widgets[0].insert("1.0", "（文档无可显示段落）\n")
            self.preview_text_widgets[1].insert("1.0", "（右页无内容）\n")

        line_no = 1
        right_line_no: dict[tk.Text, int] = {text: 1 for text in self.preview_text_widgets}
        for para in self.paragraph_data:
            idx = int(para["index"])
            page = int(para["page_est"])
            if page not in page_to_widget:
                continue

            level_label = self._level_to_label(int(para["detected_level"]))
            left_text = f"P{page} [{level_label}]"
            right_text_raw = str(para["text"]).replace("\r", " ").replace("\n", " ")
            right_text = right_text_raw if right_text_raw.strip() else " "
            target_widget = page_to_widget[page]
            target_line = right_line_no[target_widget]
            font_size = para.get("font_size_pt")
            size = 11 if font_size in (None, "") else max(8, min(40, int(round(float(font_size)))))
            first_indent_pt = float(para.get("first_line_indent_pt", 0.0) or 0.0)
            wrapped_text = self._prepare_preview_text(right_text, size, first_line_indent_pt=first_indent_pt)
            wrapped_line_count = max(1, wrapped_text.count("\n") + 1)

            self.left_class_text.insert("end", left_text + "\n")
            target_widget.insert("end", wrapped_text + "\n")

            self.left_para_ranges[idx] = (f"{line_no}.0", f"{line_no}.end")
            self.para_ranges[idx] = (target_widget, f"{target_line}.0", f"{target_line + wrapped_line_count - 1}.end")
            self.line_to_para_index_left[line_no] = idx
            for ln in range(target_line, target_line + wrapped_line_count):
                self.line_to_para_index_right[target_widget][ln] = idx
            self._apply_para_style(target_widget, idx, para, target_line, wrapped_line_count)
            line_no += 1
            right_line_no[target_widget] += wrapped_line_count
        if line_no == 1 and self.paragraph_data:
            # 当前页无内容时，显示占位文本，避免空白误解为渲染失败。
            self.left_class_text.insert("1.0", f"P{self.current_view_page} [正文]\n")
            self.preview_text_widgets[0].insert("1.0", "（当前页无可显示段落）\n")
        if right_line_no[self.preview_text_widgets[0]] == 1 and self.paragraph_data:
            self.preview_text_widgets[0].insert("1.0", "（左页无内容）\n")
        if right_line_no[self.preview_text_widgets[1]] == 1:
            self.preview_text_widgets[1].insert("1.0", "（右页无内容）\n")

        self.left_class_text.tag_configure("selected", background="#ffe7a3")
        for text in self.preview_text_widgets:
            text.tag_configure("selected", background="#ffe7a3")
        if preserve_selection:
            valid_indices = {int(para["index"]) for para in self.paragraph_data}
            self.selected_indices = {idx for idx in previous_selection if idx in valid_indices}
            self.last_clicked_index = previous_clicked if previous_clicked in valid_indices else None
        else:
            self.selected_indices.clear()
            self.last_clicked_index = None
        self._apply_selection_highlight()
        self._sync_style_toggle_state_from_selection()
        self._set_preview_widgets_state("disabled")
        self._update_hint()

    def _refresh_external_view(self) -> None:
        self._set_preview_widgets_state("normal")
        self.left_class_text.delete("1.0", "end")
        for text in self.preview_text_widgets:
            text.delete("1.0", "end")
            self._clear_para_style_tags(text)

        pages = self._split_external_content_to_pages(self.external_preview_content)
        self.total_view_pages = max(1, len(pages))
        self.current_view_page = min(max(1, self.current_view_page), self._max_pair_start())
        self._update_page_indicator()
        show_pages = [self.current_view_page]
        if self.current_view_page + 1 <= self.total_view_pages:
            show_pages.append(self.current_view_page + 1)

        for i, page_no in enumerate(show_pages):
            text = self.preview_text_widgets[i]
            self.left_class_text.insert("end", f"P{page_no} [正文]\n")
            text.insert("1.0", pages[page_no - 1] if pages[page_no - 1].strip() else " ")
            text.tag_configure("external-body", font=("宋体", int(round(self.WORD_DEFAULT_FONT_PT["body"]))))
            text.tag_add("external-body", "1.0", "end")

        if len(show_pages) < 2:
            self.preview_text_widgets[1].insert("1.0", "（右页无内容）")
        self.selected_indices.clear()
        self.last_clicked_index = None
        self._sync_style_toggle_state_from_selection()
        self._set_preview_widgets_state("disabled")
        self._update_hint("已刷新外置浏览内容")

    def _split_external_content_to_pages(self, content: str) -> list[str]:
        margins = self._get_active_margins_mm()
        lines_per_page, chars_per_line = self._compute_grid_metrics(margins, self.WORD_DEFAULT_FONT_PT["body"])
        normalized = content.replace("\r\n", "\n").replace("\r", "\n")
        raw_lines: list[str] = []
        for para in normalized.split("\n"):
            wrapped = self._word_like_wrap(para if para else " ", float(chars_per_line))
            raw_lines.extend(wrapped.split("\n"))
        if not raw_lines:
            raw_lines = [" "]
        pages: list[str] = []
        for start in range(0, len(raw_lines), max(1, lines_per_page)):
            chunk = raw_lines[start : start + max(1, lines_per_page)]
            pages.append("\n".join(chunk))
        return pages

    def _update_page_indicator(self) -> None:
        end_page = min(self.total_view_pages, self.current_view_page + 1)
        self.page_info_label.configure(text=f"第 {self.current_view_page}-{end_page} / {self.total_view_pages} 页")
        if self.current_view_page <= 1:
            self.btn_prev_page.configure(state="disabled")
        else:
            self.btn_prev_page.configure(state="normal")
        if self.current_view_page >= self._max_pair_start():
            self.btn_next_page.configure(state="disabled")
        else:
            self.btn_next_page.configure(state="normal")

    def _max_pair_start(self) -> int:
        if self.total_view_pages <= 1:
            return 1
        return self.total_view_pages if self.total_view_pages % 2 == 1 else self.total_view_pages - 1

    def on_prev_page(self) -> None:
        if self.current_view_page <= 1:
            return
        self.current_view_page = max(1, self.current_view_page - 2)
        self._play_page_turn_effect(direction=-1)
        self._refresh_document_view(preserve_selection=True)
        self._update_hint("已切换到上一页")

    def on_next_page(self) -> None:
        max_start = self._max_pair_start()
        if self.current_view_page >= max_start:
            return
        self.current_view_page = min(max_start, self.current_view_page + 2)
        self._play_page_turn_effect(direction=1)
        self._refresh_document_view(preserve_selection=True)
        self._update_hint("已切换到下一页")

    def _play_page_turn_effect(self, direction: int) -> None:
        shift = 8 if direction > 0 else -8
        for win_id in self.page_window_ids:
            if win_id is not None and self.preview_canvas.type(win_id):
                self.preview_canvas.move(win_id, shift, 0)
        self.after(45, lambda: self._draw_a4_page(max(120, self.preview_canvas.winfo_width()), max(120, self.preview_canvas.winfo_height())))

    def _level_to_label(self, level: int) -> str:
        mapping = {
            DocxHandler.LEVEL_BIG_TITLE: "大标题",
            DocxHandler.LEVEL_H1: "1级",
            DocxHandler.LEVEL_H2: "2级",
            DocxHandler.LEVEL_H3: "3级",
            DocxHandler.LEVEL_H4: "4级",
            DocxHandler.LEVEL_BODY: "正文",
        }
        return mapping.get(level, "正文")

    def _apply_selection_highlight(self) -> None:
        self._set_preview_widgets_state("normal")
        self.left_class_text.tag_remove("selected", "1.0", "end")
        for text in self.preview_text_widgets:
            text.tag_remove("selected", "1.0", "end")
        for idx in sorted(self.selected_indices):
            left_range = self.left_para_ranges.get(idx)
            right_range = self.para_ranges.get(idx)
            if left_range:
                self.left_class_text.tag_add("selected", left_range[0], left_range[1])
            if right_range:
                target_widget, start, end = right_range
                target_widget.tag_add("selected", start, end)
        self._set_preview_widgets_state("disabled")

    def _clear_para_style_tags(self, text_widget: tk.Text) -> None:
        for tag in text_widget.tag_names():
            if tag.startswith("para-style-"):
                text_widget.tag_delete(tag)

    def _apply_para_style(
        self, text_widget: tk.Text, para_idx: int, para: dict[str, object], line_no: int, line_count: int
    ) -> None:
        tag_name = f"para-style-{para_idx}"
        level_to_size = {
            DocxHandler.LEVEL_BIG_TITLE: self.WORD_DEFAULT_FONT_PT["h1"],
            DocxHandler.LEVEL_H1: self.WORD_DEFAULT_FONT_PT["h2"],
            DocxHandler.LEVEL_H2: self.WORD_DEFAULT_FONT_PT["h3"],
            DocxHandler.LEVEL_H3: self.WORD_DEFAULT_FONT_PT["h4"],
            DocxHandler.LEVEL_H4: self.WORD_DEFAULT_FONT_PT["h5"],
            DocxHandler.LEVEL_BODY: self.WORD_DEFAULT_FONT_PT["body"],
        }
        detected_level = int(para.get("detected_level", DocxHandler.LEVEL_BODY))
        font_size = para.get("font_size_pt")
        fallback_size = level_to_size.get(detected_level, self.WORD_DEFAULT_FONT_PT["body"])
        size = int(round(fallback_size if font_size in (None, "") else max(8, min(40, float(font_size)))))
        weight = "bold" if bool(para.get("bold")) else "normal"
        italic = bool(para.get("italic"))
        underline = bool(para.get("underline"))
        family = str(para.get("font_name") or "宋体")
        align_raw = str(para.get("alignment") or "left").lower()
        justify = align_raw if align_raw in {"left", "center", "right"} else "left"

        line_spacing = para.get("line_spacing")
        spacing_ratio = 1.0
        if isinstance(line_spacing, (int, float)):
            spacing_ratio = float(line_spacing) if float(line_spacing) <= 6 else max(1.0, float(line_spacing) / max(size, 1))
        spacing_extra = max(0, int(round(size * (spacing_ratio - 1.0))))
        left_indent_pt = float(para.get("left_indent_pt", 0.0) or 0.0)
        first_line_indent_pt = float(para.get("first_line_indent_pt", 0.0) or 0.0)
        # Approximate Word point -> Tk pixel conversion for preview indentation.
        left_px = max(0, int(round(left_indent_pt * 1.33)))
        first_px = max(0, int(round((left_indent_pt + first_line_indent_pt) * 1.33)))

        font_style: list[str] = []
        if weight == "bold":
            font_style.append("bold")
        if italic:
            font_style.append("italic")
        font_value: tuple[object, ...] = (family, size, *font_style) if font_style else (family, size)
        text_widget.tag_configure(
            tag_name,
            font=font_value,
            justify=justify,
            spacing2=spacing_extra,
            underline=1 if underline else 0,
            lmargin1=first_px,
            lmargin2=left_px,
        )
        end_line = line_no + max(1, line_count) - 1
        text_widget.tag_add(tag_name, f"{line_no}.0", f"{end_line}.end+1c")

    def _prepare_preview_text(self, text: str, font_size: int, first_line_indent_pt: float = 0.0) -> str:
        normalized = self._normalize_punctuation_for_preview(text)
        if first_line_indent_pt > 0:
            # Force first-line indent visibility in preview across widget modes.
            indent_chars = max(1, int(round(first_line_indent_pt / max(8.0, float(font_size)))))
            normalized = ("　" * indent_chars) + normalized
        max_units = self._estimate_line_units(font_size)
        return self._word_like_wrap(normalized, max_units)

    def _normalize_punctuation_for_preview(self, text: str) -> str:
        # Keep author spacing as much as possible; only normalize tabs and
        # repeated spaces to reduce accidental text distortion in preview.
        content = text.replace("\t", " ")
        content = re.sub(r" {2,}", " ", content)
        content = content.replace("...", "……")
        chars = list(content)
        out: list[str] = []
        for i, ch in enumerate(chars):
            if ch in self.HALF_TO_FULL_PUNC:
                prev_cjk = self._is_cjk_neighbor(chars, i, -1)
                next_cjk = self._is_cjk_neighbor(chars, i, 1)
                if prev_cjk or next_cjk:
                    out.append(self.HALF_TO_FULL_PUNC[ch])
                else:
                    out.append(ch)
            elif ch in self.FULL_TO_HALF_PUNC:
                prev_cjk = self._is_cjk_neighbor(chars, i, -1)
                next_cjk = self._is_cjk_neighbor(chars, i, 1)
                if prev_cjk or next_cjk:
                    out.append(ch)
                else:
                    out.append(self.FULL_TO_HALF_PUNC[ch])
            else:
                out.append(ch)
        return "".join(out)

    def _is_cjk_neighbor(self, chars: list[str], idx: int, step: int) -> bool:
        i = idx + step
        while 0 <= i < len(chars):
            ch = chars[i]
            if ch.isspace():
                i += step
                continue
            return bool(self.CJK_RE.match(ch))
        return False

    def _estimate_line_units(self, font_size: int) -> float:
        margins = self._get_active_margins_mm()
        _lines, chars = self._compute_grid_metrics(margins, float(font_size))
        return float(max(8, chars - self.PREVIEW_SAFE_CHAR_BUFFER))

    def _word_like_wrap(self, text: str, max_units: float) -> str:
        if not text.strip():
            return " "

        tokens: list[str] = []
        i = 0
        while i < len(text):
            if text.startswith("——", i):
                tokens.append("——")
                i += 2
            elif text.startswith("……", i):
                tokens.append("……")
                i += 2
            else:
                tokens.append(text[i])
                i += 1

        lines: list[list[str]] = []
        current: list[str] = []
        used = 0.0

        for tok in tokens:
            w = self._token_units(tok)
            if current and used + w > max_units:
                # Never place forbidden punctuation at line start.
                if tok in self.LINE_START_FORBIDDEN:
                    current.append(tok)
                    lines.append(current)
                    current = []
                    used = 0.0
                    continue
                # Never keep opening punctuation at line end.
                if current and current[-1] in self.LINE_END_FORBIDDEN:
                    carry = current.pop()
                    had_prefix = bool(current)
                    lines.append(current if had_prefix else [carry])
                    current = [carry, tok] if had_prefix else [tok]
                    used = sum(self._token_units(x) for x in current)
                    continue
                lines.append(current)
                current = []
                used = 0.0
            current.append(tok)
            used += w
        if current:
            lines.append(current)

        return "\n".join("".join(line) for line in lines if line)

    def _token_units(self, token: str) -> float:
        if token in {"——", "……"}:
            return 2.0
        if self.CJK_RE.match(token) or token in self.HALF_TO_FULL_PUNC.values() or token in {"“", "”", "‘", "’", "《", "》", "【", "】"}:
            return 1.0
        return 0.5

    def _set_selection(self, indices: set[int], reason: str) -> None:
        valid_indices = {int(para["index"]) for para in self.paragraph_data}
        self.selected_indices = {idx for idx in indices if idx in valid_indices}
        self._apply_selection_highlight()
        self._sync_style_toggle_state_from_selection()
        self._update_hint(reason)

    def _selected_or_warn(self) -> set[int]:
        if not self.docx_handler.is_loaded:
            messagebox.showinfo("提示", "请先打开文档")
            return set()
        if not self.selected_indices:
            messagebox.showinfo("提示", "请先选择至少一个段落")
            return set()
        return set(self.selected_indices)

    def _selected_or_empty(self) -> set[int]:
        if not self.docx_handler.is_loaded:
            return set()
        return set(self.selected_indices)

    def _refresh_after_doc_mutation(self, reason: str) -> None:
        # Keep Word paragraph non-splitting rule consistent after edits.
        self.docx_handler.enforce_keep_together_for_all()
        self.find_match_indices = []
        self.find_match_cursor = -1
        self.find_last_signature = None
        self._refresh_document_view(preserve_selection=True)
        self._update_hint(reason)
        self._log_operation("格式操作", reason)

    def _refresh_undo_redo_buttons(self) -> None:
        if not self.docx_handler.is_loaded:
            undo_count = 0
            redo_count = 0
        else:
            counts = self.docx_handler.get_undo_redo_counts()
            undo_count = int(counts.get("undo", 0))
            redo_count = int(counts.get("redo", 0))
        if self.btn_undo is not None:
            self.btn_undo.configure(text=f"撤销({undo_count})")
        if self.btn_redo is not None:
            self.btn_redo.configure(text=f"重做({redo_count})")

    def _parse_font_size_value(self) -> float | None:
        raw = self.size_combo.get().strip()
        return self._parse_size_token_to_pt(raw)

    def _parse_size_token_to_pt(self, raw: str) -> float | None:
        mapping = {
            "五号": 10.5,
            "小四号": 12.0,
            "四号": 14.0,
            "小三号": 15.0,
            "三号": 16.0,
            "小二号": 18.0,
            "二号": 22.0,
            "24": 24.0,
            "28": 28.0,
            "32": 32.0,
        }
        if raw in mapping:
            return mapping[raw]
        try:
            return float(raw)
        except ValueError:
            return None

    def _size_step_points(self) -> list[float]:
        points: set[float] = set()
        for value in self.size_combo.cget("values"):
            parsed = self._parse_size_token_to_pt(str(value).strip())
            if parsed is not None:
                points.add(float(parsed))
        points.update({10.5, 12.0, 14.0, 15.0, 16.0, 18.0, 22.0, 24.0, 28.0, 32.0})
        return sorted(points)

    def _size_display_text(self, pt: float) -> str:
        prefer = {
            10.5: "五号",
            12.0: "小四号",
            14.0: "四号",
            15.0: "小三号",
            16.0: "三号",
            18.0: "小二号",
            22.0: "二号",
            24.0: "24",
            28.0: "28",
            32.0: "32",
        }
        key = round(float(pt), 1)
        return prefer.get(key, f"{pt:g}")

    def _step_font_size(self, step: int) -> None:
        points = self._size_step_points()
        if not points:
            return
        current = self._parse_font_size_value()
        if current is None:
            current = 10.5
        target = points[0]
        if step > 0:
            for p in points:
                if p > current + 1e-6:
                    target = p
                    break
            else:
                target = points[-1]
        else:
            target = points[0]
            for p in points:
                if p < current - 1e-6:
                    target = p
                else:
                    break
        self.size_combo.set(self._size_display_text(target))
        self.on_apply_font_size(silent=False)

    def on_increase_font_size(self) -> None:
        self._step_font_size(1)

    def on_decrease_font_size(self) -> None:
        self._step_font_size(-1)

    def _parse_font_name_value(self) -> str:
        raw = self.font_combo.get().strip()
        return raw if raw else "宋体"

    def _parse_line_spacing_value(self) -> float | str | None:
        raw = self.spacing_combo.get().strip()
        mapping: dict[str, float | str] = {
            "单倍(1.0)": 1.0,
            "1.5倍": 1.5,
            "2倍": 2.0,
            "固定18pt": "fixed18pt",
        }
        if raw in mapping:
            return mapping[raw]
        try:
            return float(raw)
        except ValueError:
            return None

    def _parse_margin_value(self) -> tuple[bool, str | None]:
        raw = self.margin_combo.get().strip()
        if raw in self.WORD_MARGIN_PRESETS_MM:
            return True, raw
        return False, None

    def _get_active_margins_mm(self) -> dict[str, float]:
        if self.preview_margin_override_mm is not None:
            return dict(self.preview_margin_override_mm)
        return dict(self.current_doc_margins_mm)

    def _compute_grid_metrics(self, margins_mm: dict[str, float], font_size_pt: float = 10.5) -> tuple[int, int]:
        safe = self._get_preview_safe_edge_mm()
        content_w = max(
            30.0,
            self.A4_WIDTH_MM
            - margins_mm["left"]
            - margins_mm["right"]
            - safe["left"]
            - safe["right"],
        )
        content_h = max(
            40.0,
            self.A4_HEIGHT_MM
            - margins_mm["top"]
            - margins_mm["bottom"]
            - safe["top"]
            - safe["bottom"],
        )
        width_ratio = content_w / self.BASELINE_NORMAL_CONTENT_MM["width"]
        height_ratio = content_h / self.BASELINE_NORMAL_CONTENT_MM["height"]
        size_ratio = self.WORD_DEFAULT_FONT_PT["body"] / max(8.0, font_size_pt)
        precision_factor = 0.96 if self.preview_accuracy_mode == "高精度模式" else 1.0
        chars_per_line = max(10, int(round(self.PREVIEW_BASELINE_CHARS_PER_LINE * width_ratio * size_ratio * precision_factor)))
        lines_per_page = max(10, int(round(self.PREVIEW_BASELINE_LINES_PER_PAGE * height_ratio * size_ratio * precision_factor)))
        return lines_per_page, chars_per_line

    def _get_preview_safe_edge_mm(self) -> dict[str, float]:
        return dict(self.PREVIEW_SAFE_EDGE_MM)

    def _update_grid_info(self, margins_mm: dict[str, float]) -> None:
        if not hasattr(self, "grid_info_label"):
            return
        lines, chars = self._compute_grid_metrics(margins_mm, self.WORD_DEFAULT_FONT_PT["body"])
        preset = self.current_margin_preset_name
        self.grid_info_label.configure(text=f"{preset}：每页 {lines} 行 / 每行 {chars} 字")

    def _match_margin_preset_name(self, margins_mm: dict[str, float]) -> str:
        for name, preset in self.WORD_MARGIN_PRESETS_MM.items():
            if (
                abs(margins_mm["top"] - preset["top"]) <= 0.2
                and abs(margins_mm["bottom"] - preset["bottom"]) <= 0.2
                and abs(margins_mm["left"] - preset["left"]) <= 0.2
                and abs(margins_mm["right"] - preset["right"]) <= 0.2
            ):
                return name
        return "普通"

    def _update_hint(self, reason: str | None = None) -> None:
        count = len(self.selected_indices)
        stack = self.docx_handler.get_undo_redo_counts() if self.docx_handler.is_loaded else {"undo": 0, "redo": 0}
        if reason:
            self.status_text = f"{reason}，已选中 {count} 段 | 可撤销 {stack['undo']} | 可重做 {stack['redo']}"
        else:
            self.status_text = f"当前已选中 {count} 段 | 可撤销 {stack['undo']} | 可重做 {stack['redo']}"
        self._refresh_undo_redo_buttons()

    def _show_action_error(self, title: str, what: str, how: str, exc: Exception | None = None) -> None:
        tech = f"\n\n技术信息：{exc}" if exc is not None else ""
        self._log_operation("异常", what, result="失败", affected_count=0)
        messagebox.showerror(title, f"发生了什么：{what}\n怎么处理：{how}{tech}")

    def _run_long_task(self, label: str, func) -> tuple[object, float]:
        self._update_hint(f"{label}进行中，请稍候")
        self.configure(cursor="watch")
        self.update_idletasks()
        start = time.perf_counter()
        try:
            result = func()
            elapsed = (time.perf_counter() - start) * 1000.0
            return result, elapsed
        finally:
            self.configure(cursor="")
            self.update_idletasks()

    def _show_field_update_tip_if_needed(self) -> None:
        if self.field_update_tip_shown:
            return
        self.field_update_tip_shown = True
        messagebox.showinfo(
            "提示",
            "若在 WPS/Word 中看到目录或页码未即时刷新，请执行一次“更新域/更新目录（F9）”。",
        )

    def _log_operation(
        self,
        op_type: str,
        detail: str,
        result: str = "成功",
        duration_ms: float | None = None,
        affected_count: int | None = None,
    ) -> None:
        row = {
            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "type": op_type,
            "detail": detail,
            "result": result,
            "duration_ms": "" if duration_ms is None else f"{duration_ms:.0f}",
            "affected_count": "" if affected_count is None else str(affected_count),
        }
        self.operation_logs.append(row)

    def _collect_current_rules(self) -> dict[str, object]:
        return {
            "school_template": self.default_school_template,
            "font": self.font_combo.get().strip(),
            "size": self.size_combo.get().strip(),
            "line_spacing": self.spacing_combo.get().strip(),
            "margin": self.margin_combo.get().strip(),
            "toc_levels": int(self.toc_levels),
            "reference_mode": self.ref_mode_combo.get().strip() if hasattr(self, "ref_mode_combo") else "顺序编码制",
            "preview_mode": self.preview_accuracy_mode,
            "auto_backup_enabled": bool(self.auto_backup_enabled),
        }

    def _apply_rules_to_ui(self, rules: dict[str, object]) -> None:
        if not isinstance(rules, dict):
            return
        if "school_template" in rules:
            self.default_school_template = str(rules["school_template"])
        if "font" in rules and str(rules["font"]).strip():
            self.font_combo.set(str(rules["font"]).strip())
            self.on_font_option_change(self.font_combo.get())
        if "size" in rules and str(rules["size"]).strip():
            self.size_combo.set(str(rules["size"]).strip())
            self.on_size_option_change(self.size_combo.get())
        if "line_spacing" in rules and str(rules["line_spacing"]).strip():
            self.spacing_combo.set(str(rules["line_spacing"]).strip())
            self.on_line_spacing_option_change(self.spacing_combo.get())
        if "margin" in rules and str(rules["margin"]).strip():
            self.margin_combo.set(str(rules["margin"]).strip())
            self.on_margin_option_change(self.margin_combo.get())
        if "toc_levels" in rules:
            try:
                self.toc_levels = max(1, min(9, int(rules["toc_levels"])))
            except Exception:
                self.toc_levels = 3
        if "reference_mode" in rules and hasattr(self, "ref_mode_combo"):
            mode = str(rules["reference_mode"]).strip()
            if mode in {"顺序编码制", "著者-出版年制"}:
                self.ref_mode_combo.set(mode)
        if "preview_mode" in rules:
            mode = str(rules["preview_mode"]).strip()
            if mode in self.PREVIEW_MODE_OPTIONS:
                self.preview_accuracy_mode = mode
                if self.preview_mode_combo is not None:
                    self.preview_mode_combo.set(mode)
        if "auto_backup_enabled" in rules:
            self.auto_backup_enabled = bool(rules["auto_backup_enabled"])
            self.docx_handler.set_auto_backup_enabled(self.auto_backup_enabled)
            if self.auto_backup_switch is not None:
                if self.auto_backup_enabled:
                    self.auto_backup_switch.select()
                else:
                    self.auto_backup_switch.deselect()

    def _load_user_preferences(self) -> None:
        if not self.preference_file.exists():
            return
        try:
            data = json.loads(self.preference_file.read_text(encoding="utf-8"))
            rules = data.get("rules", {}) if isinstance(data, dict) else {}
            self._apply_rules_to_ui(rules if isinstance(rules, dict) else {})
        except Exception:
            return

    def on_preview_click(self, event: tk.Event | None = None) -> str | None:
        if event is None:
            return None
        if not self.paragraph_data:
            self._update_hint("当前没有可选段落")
            return "break"

        para_idx = self._resolve_preview_para_index(event)

        if para_idx is None:
            return "break"

        ctrl_pressed = bool(event.state & 0x0004)
        if ctrl_pressed:
            if para_idx in self.selected_indices:
                self.selected_indices.remove(para_idx)
            else:
                self.selected_indices.add(para_idx)
        else:
            self.selected_indices = {para_idx}

        self.last_clicked_index = para_idx
        self._apply_selection_highlight()
        self._sync_style_toggle_state_from_selection()
        self._update_hint("已按段落点击选中")
        return "break"

    def on_preview_right_click(self, event: tk.Event | None = None) -> str | None:
        if event is None:
            return None
        if not self.paragraph_data:
            self._update_hint("当前没有可选段落")
            return "break"
        para_idx = self._resolve_preview_para_index(event)
        if para_idx is None:
            return "break"
        if para_idx not in self.selected_indices:
            self._update_hint("该段落未被选中")
            return "break"
        self.selected_indices.remove(para_idx)
        self._apply_selection_highlight()
        self._sync_style_toggle_state_from_selection()
        self._update_hint("已局部取消选中")
        return "break"

    def _resolve_preview_para_index(self, event: tk.Event) -> int | None:
        widget = event.widget
        if widget != self.left_class_text and widget not in self.preview_text_widgets:
            return None
        pos_index = widget.index(f"@{event.x},{event.y}")
        line_no = int(pos_index.split(".")[0])
        if widget == self.left_class_text:
            return self.line_to_para_index_left.get(line_no)
        return self.line_to_para_index_right.get(widget, {}).get(line_no)

    def on_open_document(self) -> None:
        file_path = filedialog.askopenfilename(
            title="选择 Word 文档",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")],
        )
        if not file_path:
            return

        try:
            self.docx_handler.load(file_path)
            self.find_match_indices = []
            self.find_match_cursor = -1
            self.find_last_signature = None
            self.current_file = file_path
            self.file_label.configure(text=Path(file_path).name)
            layout = self.docx_handler.get_page_layout_mm()
            self.current_doc_margins_mm = {
                "left": layout["margin_left_mm"],
                "right": layout["margin_right_mm"],
                "top": layout["margin_top_mm"],
                "bottom": layout["margin_bottom_mm"],
            }
            self.current_margin_preset_name = self._match_margin_preset_name(self.current_doc_margins_mm)
            self.preview_margin_override_mm = None
            self.margin_combo.set(self.current_margin_preset_name)
            self.current_view_page = 1
            self._refresh_document_view()
            self._update_hint("文档已打开")
            self.preview_canvas.update_idletasks()
        except Exception as exc:  # noqa: BLE001
            self._show_action_error("打开失败", "文档加载失败。", "请确认文件未损坏、未被占用，并重试“打开文档”。", exc)

    def on_save_document(self) -> None:
        if not self.docx_handler.is_loaded:
            messagebox.showinfo("提示", "请先打开文档")
            return
        try:
            self.docx_handler.save()
            messagebox.showinfo("提示", "文档已保存")
        except Exception:
            self.on_save_as_document()

    def on_save_as_document(self) -> None:
        if not self.docx_handler.is_loaded:
            messagebox.showinfo("提示", "请先打开文档")
            return
        file_path = filedialog.asksaveasfilename(
            title="另存为",
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")],
        )
        if not file_path:
            return
        try:
            self.docx_handler.save(file_path)
            self.current_file = file_path
            self.file_label.configure(text=Path(file_path).name)
            messagebox.showinfo("提示", "文档已另存为")
        except Exception as exc:  # noqa: BLE001
            self._show_action_error("保存失败", "文档保存失败。", "请检查目标路径权限，或改用“另存为”。", exc)

    def on_save_shortcut(self, _event: tk.Event | None = None) -> str:
        self.on_save_document()
        return "break"

    def on_select_all_shortcut(self, _event: tk.Event | None = None) -> str:
        self.on_select_all()
        return "break"

    def on_cancel_selection_shortcut(self, _event: tk.Event | None = None) -> str:
        self.on_select_cancel()
        return "break"

    def on_refresh_shortcut(self, _event: tk.Event | None = None) -> str:
        if not self.docx_handler.is_loaded:
            self._update_hint("请先打开文档")
            return "break"
        self._refresh_document_view(preserve_selection=True)
        self._update_hint("已刷新页面")
        return "break"

    def on_select_pick(self) -> None:
        if not self.paragraph_data:
            messagebox.showinfo("提示", "请先打开文档")
            return
        anchor = self.last_clicked_index
        if anchor is None and self.selected_indices:
            anchor = sorted(self.selected_indices)[0]
        if anchor is None:
            messagebox.showinfo("提示", "请先点击一个段落，再使用选取")
            return

        indices = set(self.docx_handler.get_range_under_title(anchor))
        self._set_selection(indices, "已按标题管辖范围选取")

    def on_select_all(self) -> None:
        if not self.paragraph_data:
            return
        indices = {int(para["index"]) for para in self.paragraph_data}
        self._set_selection(indices, "已全选")

    def on_select_cancel(self) -> None:
        self._set_selection(set(), "已取消选择")

    def on_mark_title(self) -> None:
        self._select_by_level(DocxHandler.LEVEL_BIG_TITLE, "大标题")

    def on_mark_h1(self) -> None:
        self._select_by_level(DocxHandler.LEVEL_H1, "一级")

    def on_mark_h2(self) -> None:
        self._select_by_level(DocxHandler.LEVEL_H2, "二级")

    def on_mark_h3(self) -> None:
        self._select_by_level(DocxHandler.LEVEL_H3, "三级")

    def on_mark_h4(self) -> None:
        self._select_by_level(DocxHandler.LEVEL_H4, "四级")

    def on_mark_body(self) -> None:
        self._select_by_level(DocxHandler.LEVEL_BODY, "正文")

    def on_select_batch(self) -> None:
        if not self.docx_handler.is_loaded:
            messagebox.showinfo("提示", "请先打开文档")
            return

        popup = tk.Toplevel(self)
        popup.title("批量选择")
        popup.transient(self)
        popup.grab_set()
        popup.resizable(False, False)
        popup.configure(bg="#f7f7f7")

        frame = tk.Frame(popup, bg="#f7f7f7", padx=14, pady=12)
        frame.pack(fill="both", expand=True)

        level_options = [
            ("一级标题", DocxHandler.LEVEL_H1),
            ("二级标题", DocxHandler.LEVEL_H2),
            ("三级标题", DocxHandler.LEVEL_H3),
            ("四级标题", DocxHandler.LEVEL_H4),
            ("正文", DocxHandler.LEVEL_BODY),
        ]
        vars_by_level: list[tuple[tk.BooleanVar, int, str]] = []
        for row_idx, (label, level) in enumerate(level_options):
            var = tk.BooleanVar(value=False)
            cb = tk.Checkbutton(frame, text=label, variable=var, bg="#f7f7f7", anchor="w")
            cb.grid(row=row_idx, column=0, sticky="w", pady=2)
            vars_by_level.append((var, level, label))

        btn_row = tk.Frame(frame, bg="#f7f7f7")
        btn_row.grid(row=len(level_options), column=0, sticky="e", pady=(10, 0))

        def _confirm() -> None:
            picked = [(level, label) for var, level, label in vars_by_level if var.get()]
            if not picked:
                messagebox.showinfo("提示", "请至少勾选一个分类")
                return
            indices: set[int] = set()
            labels: list[str] = []
            for level, label in picked:
                indices.update(self.docx_handler.get_indices_by_level(level))
                labels.append(label)
            popup.destroy()
            self._set_selection(indices, f"已批量选中：{'/'.join(labels)}")

        tk.Button(btn_row, text="确定", width=10, command=_confirm).pack(side="left", padx=(0, 6))
        tk.Button(btn_row, text="取消", width=10, command=popup.destroy).pack(side="left")

    def on_find_dialog(self) -> None:
        if not self.docx_handler.is_loaded:
            messagebox.showinfo("提示", "请先打开文档")
            return
        popup = tk.Toplevel(self)
        popup.title("查找")
        popup.transient(self)
        popup.grab_set()
        popup.resizable(False, False)
        popup.configure(bg="#f7f7f7")

        frame = tk.Frame(popup, bg="#f7f7f7", padx=14, pady=12)
        frame.pack(fill="both", expand=True)
        tk.Label(frame, text="查找内容：", bg="#f7f7f7").grid(row=0, column=0, sticky="w")
        keyword_var = tk.StringVar(value="")
        entry = tk.Entry(frame, textvariable=keyword_var, width=34)
        entry.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(4, 8))
        case_var = tk.BooleanVar(value=False)
        whole_var = tk.BooleanVar(value=False)
        tk.Checkbutton(frame, text="区分大小写", variable=case_var, bg="#f7f7f7").grid(row=2, column=0, sticky="w")
        tk.Checkbutton(frame, text="整词匹配", variable=whole_var, bg="#f7f7f7").grid(row=2, column=1, sticky="w")

        count_var = tk.StringVar(value="匹配：0")
        tk.Label(frame, textvariable=count_var, bg="#f7f7f7", fg="#3f4d63").grid(row=3, column=0, columnspan=2, sticky="w")

        def _find_next() -> None:
            query = keyword_var.get().strip()
            if not query:
                messagebox.showinfo("提示", "请输入查找内容")
                return
            self._find_and_focus_next(
                query,
                case_sensitive=case_var.get(),
                whole_word=whole_var.get(),
                selected_only=False,
                direction=1,
            )
            count_var.set(self._find_match_text())

        def _find_prev() -> None:
            query = keyword_var.get().strip()
            if not query:
                messagebox.showinfo("提示", "请输入查找内容")
                return
            self._find_and_focus_next(
                query,
                case_sensitive=case_var.get(),
                whole_word=whole_var.get(),
                selected_only=False,
                direction=-1,
            )
            count_var.set(self._find_match_text())

        btn_row = tk.Frame(frame, bg="#f7f7f7")
        btn_row.grid(row=4, column=0, columnspan=2, sticky="e", pady=(10, 0))
        tk.Button(btn_row, text="上一处", width=10, command=_find_prev).pack(side="left", padx=(0, 6))
        tk.Button(btn_row, text="下一处", width=10, command=_find_next).pack(side="left", padx=(0, 6))
        tk.Button(btn_row, text="关闭", width=10, command=popup.destroy).pack(side="left")
        entry.focus_set()
        entry.bind("<Return>", lambda _e: _find_next())

    def on_replace_dialog(self) -> None:
        if not self.docx_handler.is_loaded:
            messagebox.showinfo("提示", "请先打开文档")
            return
        popup = tk.Toplevel(self)
        popup.title("替代")
        popup.transient(self)
        popup.grab_set()
        popup.resizable(False, False)
        popup.configure(bg="#f7f7f7")

        frame = tk.Frame(popup, bg="#f7f7f7", padx=14, pady=12)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="查找内容：", bg="#f7f7f7").grid(row=0, column=0, sticky="w")
        find_var = tk.StringVar(value="")
        find_entry = tk.Entry(frame, textvariable=find_var, width=34)
        find_entry.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(4, 8))
        tk.Label(frame, text="替换为：", bg="#f7f7f7").grid(row=2, column=0, sticky="w")
        repl_var = tk.StringVar(value="")
        repl_entry = tk.Entry(frame, textvariable=repl_var, width=34)
        repl_entry.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(4, 8))

        case_var = tk.BooleanVar(value=False)
        whole_var = tk.BooleanVar(value=False)
        selected_only_var = tk.BooleanVar(value=False)
        tk.Checkbutton(frame, text="区分大小写", variable=case_var, bg="#f7f7f7").grid(row=4, column=0, sticky="w")
        tk.Checkbutton(frame, text="整词匹配", variable=whole_var, bg="#f7f7f7").grid(row=4, column=1, sticky="w")
        tk.Checkbutton(frame, text="仅选中范围替换", variable=selected_only_var, bg="#f7f7f7").grid(row=5, column=0, sticky="w")
        count_var = tk.StringVar(value="匹配：0")
        tk.Label(frame, textvariable=count_var, bg="#f7f7f7", fg="#3f4d63").grid(row=5, column=1, sticky="e")

        def _replace_current() -> None:
            query = find_var.get().strip()
            if not query:
                messagebox.showinfo("提示", "请输入查找内容")
                return
            if not self.find_match_indices:
                self._find_and_focus_next(
                    query,
                    case_sensitive=case_var.get(),
                    whole_word=whole_var.get(),
                    selected_only=selected_only_var.get(),
                    direction=1,
                )
            if not self.find_match_indices:
                return
            cur_idx = self.find_match_indices[self.find_match_cursor]
            result = self.docx_handler.replace_text(
                query,
                repl_var.get(),
                case_sensitive=case_var.get(),
                whole_word=whole_var.get(),
                indices={cur_idx},
                max_replacements_per_para=1,
            )
            replaced = int(result.get("replaced_count", 0))
            if replaced <= 0:
                messagebox.showinfo("提示", "当前匹配位置不可替换")
                return
            self._refresh_after_doc_mutation(f"已替换 {replaced} 处")
            self.find_match_indices = []
            self.find_match_cursor = -1
            self.find_last_signature = None
            self._find_and_focus_next(
                query,
                case_sensitive=case_var.get(),
                whole_word=whole_var.get(),
                selected_only=selected_only_var.get(),
                direction=1,
            )
            count_var.set(self._find_match_text())

        def _replace_all() -> None:
            query = find_var.get().strip()
            if not query:
                messagebox.showinfo("提示", "请输入查找内容")
                return
            target_indices = set(self.selected_indices) if selected_only_var.get() else None
            scope_count = (
                len(self.docx_handler.find_text(query, case_sensitive=case_var.get(), whole_word=whole_var.get()))
                if target_indices is None
                else len(
                    [
                        idx
                        for idx in self.docx_handler.find_text(
                            query, case_sensitive=case_var.get(), whole_word=whole_var.get()
                        )
                        if idx in target_indices
                    ]
                )
            )
            if scope_count <= 0:
                messagebox.showinfo("提示", "未找到可替换内容")
                return
            if not messagebox.askyesno("确认全部替换", f"预计匹配 {scope_count} 处，是否继续全部替换？"):
                return
            result = self.docx_handler.replace_text(
                query,
                repl_var.get(),
                case_sensitive=case_var.get(),
                whole_word=whole_var.get(),
                indices=target_indices,
            )
            replaced = int(result.get("replaced_count", 0))
            changed_paragraphs = int(result.get("changed_paragraphs", 0))
            if replaced <= 0:
                messagebox.showinfo("提示", "未找到可替换内容")
                return
            self._refresh_after_doc_mutation(f"已替换 {replaced} 处（影响 {changed_paragraphs} 段）")
            messagebox.showinfo("替换完成", f"共替换 {replaced} 处，影响 {changed_paragraphs} 段。")
            self.find_match_indices = []
            self.find_match_cursor = -1
            self.find_last_signature = None
            count_var.set("匹配：0")

        btn_row = tk.Frame(frame, bg="#f7f7f7")
        btn_row.grid(row=6, column=0, columnspan=2, sticky="e", pady=(10, 0))
        tk.Button(btn_row, text="替换当前", width=10, command=_replace_current).pack(side="left", padx=(0, 6))
        tk.Button(btn_row, text="全部替换", width=10, command=_replace_all).pack(side="left", padx=(0, 6))
        tk.Button(btn_row, text="关闭", width=10, command=popup.destroy).pack(side="left")
        find_entry.focus_set()
        repl_entry.bind("<Return>", lambda _e: _replace_current())

    def _find_match_text(self) -> str:
        if not self.find_match_indices:
            return "匹配：0"
        return f"匹配：{self.find_match_cursor + 1}/{len(self.find_match_indices)}"

    def _find_and_focus_next(
        self,
        keyword: str,
        case_sensitive: bool = False,
        whole_word: bool = False,
        selected_only: bool = False,
        direction: int = 1,
    ) -> None:
        query = keyword.strip()
        if not query:
            return
        signature = (query, bool(case_sensitive), bool(whole_word), bool(selected_only))
        refreshed = False
        if self.find_last_signature != signature:
            self.find_match_indices = []
            self.find_match_cursor = -1
            self.find_last_signature = signature
        if not self.find_match_indices:
            result = self.docx_handler.find_text(query, case_sensitive=case_sensitive, whole_word=whole_word)
            if selected_only:
                selected = set(self.selected_indices)
                result = [idx for idx in result if idx in selected]
            self.find_match_indices = result
            self.find_match_cursor = -1
            refreshed = True
        if not self.find_match_indices:
            messagebox.showinfo("查找结果", "未找到匹配内容")
            return
        step = 1 if direction >= 0 else -1
        self.find_match_cursor = (self.find_match_cursor + step) % len(self.find_match_indices)
        idx = self.find_match_indices[self.find_match_cursor]
        self.last_clicked_index = idx
        self._set_selection({idx}, f"查找命中 {self.find_match_cursor + 1}/{len(self.find_match_indices)}")
        if refreshed:
            self._update_hint(f"查找到 {len(self.find_match_indices)} 处，已定位第 1 处")

    def _select_by_level(self, level: int, label: str) -> None:
        if not self.docx_handler.is_loaded:
            messagebox.showinfo("提示", "请先打开文档")
            return
        indices = set(self.docx_handler.get_indices_by_level(level))
        self._set_selection(indices, f"已按{label}分类选取")

    def on_apply_font_size(self, silent: bool = False) -> None:
        target_indices = self._selected_or_empty() if silent else self._selected_or_warn()
        if not target_indices:
            return
        font_size = self._parse_font_size_value()
        if font_size is None:
            if not silent:
                messagebox.showerror("格式错误", "字号格式无效，请使用如 12 或 小四")
            return
        try:
            changed = self.docx_handler.apply_font_size(target_indices, font_size)
            if not changed:
                if not silent:
                    messagebox.showinfo("提示", "没有可修改的段落")
                return
            self._refresh_after_doc_mutation(f"已应用字号 {font_size:g}pt")
        except Exception as exc:  # noqa: BLE001
            if not silent:
                messagebox.showerror("操作失败", f"应用字号失败：{exc}")

    def on_font_option_change(self, _choice: str) -> None:
        self.on_apply_font_name(silent=True)

    def on_size_option_change(self, _choice: str) -> None:
        self.on_apply_font_size(silent=True)

    def on_font_option_keyup(self, _event: tk.Event | None = None) -> str | None:
        self.on_apply_font_name(silent=True)
        return "break"

    def on_size_option_keyup(self, _event: tk.Event | None = None) -> str | None:
        self.on_apply_font_size(silent=True)
        return "break"

    def on_line_spacing_option_change(self, _choice: str) -> None:
        self.on_apply_line_spacing(silent=True)

    def on_line_spacing_option_keyup(self, _event: tk.Event | None = None) -> str | None:
        self.on_apply_line_spacing(silent=True)
        return "break"

    def on_margin_option_change(self, _choice: str) -> None:
        self.on_apply_page_margin(silent=True)

    def on_margin_option_keyup(self, _event: tk.Event | None = None) -> str | None:
        self.on_apply_page_margin(silent=True)
        return "break"

    def write_external_content(self, content: str, refresh: bool = True) -> None:
        self.external_preview_content = content or ""
        if refresh:
            self.current_view_page = 1
            self._refresh_document_view(preserve_selection=False)

    def on_apply_external_content(self) -> None:
        content = self.external_input.get("1.0", "end").rstrip("\n")
        self.write_external_content(content, refresh=True)

    def on_clear_external_content(self) -> None:
        self.external_input.delete("1.0", "end")
        self.write_external_content("", refresh=True)

    def on_apply_font_name(self, silent: bool = False) -> None:
        target_indices = self._selected_or_empty() if silent else self._selected_or_warn()
        if not target_indices:
            return
        font_name = self._parse_font_name_value()
        try:
            changed = self.docx_handler.apply_font_name(target_indices, font_name)
            if not changed:
                if not silent:
                    messagebox.showinfo("提示", "没有可修改的段落")
                return
            self._refresh_after_doc_mutation(f"已应用字体 {font_name}")
        except Exception as exc:  # noqa: BLE001
            if not silent:
                messagebox.showerror("操作失败", f"应用字体失败：{exc}")

    def on_apply_line_spacing(self, silent: bool = False) -> None:
        target_indices = self._selected_or_empty() if silent else self._selected_or_warn()
        if not target_indices:
            return
        spacing = self._parse_line_spacing_value()
        if spacing is None:
            if not silent:
                messagebox.showerror("格式错误", "行间距格式无效")
            return
        try:
            changed = self.docx_handler.apply_line_spacing(target_indices, spacing)
            if not changed:
                if not silent:
                    messagebox.showinfo("提示", "没有可修改的段落")
                return
            label = "固定18pt" if spacing == "fixed18pt" else f"{float(spacing):g}倍"
            self._refresh_after_doc_mutation(f"已应用行间距 {label}")
        except Exception as exc:  # noqa: BLE001
            if not silent:
                messagebox.showerror("操作失败", f"应用行间距失败：{exc}")

    def on_apply_page_margin(self, silent: bool = False) -> None:
        valid, preset_name = self._parse_margin_value()
        if not valid:
            # Ignore transient typing content in silent mode.
            if not silent:
                messagebox.showerror("格式错误", "页边距选项无效")
            return
        try:
            assert preset_name is not None
            margin = self.WORD_MARGIN_PRESETS_MM[preset_name]
            changed = False
            if self.docx_handler.is_loaded:
                changed = self.docx_handler.apply_page_margins_mm(
                    top_mm=margin["top"],
                    bottom_mm=margin["bottom"],
                    left_mm=margin["left"],
                    right_mm=margin["right"],
                )
            self.current_margin_preset_name = preset_name
            self.preview_margin_override_mm = dict(margin)
            self.current_doc_margins_mm = dict(margin)
            self._draw_a4_page(max(120, self.preview_canvas.winfo_width()), max(120, self.preview_canvas.winfo_height()))
            self._refresh_document_view(preserve_selection=True)
            if changed:
                self._update_hint(
                    f"已应用页边距 {preset_name} (上{margin['top']/10:.2f}/下{margin['bottom']/10:.2f}/左{margin['left']/10:.2f}/右{margin['right']/10:.2f}cm)"
                )
                self._log_operation(
                    "格式操作",
                    f"应用页边距 {preset_name} (上{margin['top']/10:.2f}/下{margin['bottom']/10:.2f}/左{margin['left']/10:.2f}/右{margin['right']/10:.2f}cm)",
                )
            elif self.docx_handler.is_loaded and not silent:
                messagebox.showinfo("提示", "页边距未变化")
        except Exception as exc:  # noqa: BLE001
            if not silent:
                messagebox.showerror("操作失败", f"应用页边距失败：{exc}")

    def on_apply_first_line_indent(self) -> None:
        if not self.docx_handler.is_loaded:
            messagebox.showinfo("提示", "请先打开文档")
            return
        try:
            changed = self.docx_handler.apply_first_line_indent_chars(chars=2.0, base_font_pt=12.0)
            if changed:
                self._refresh_after_doc_mutation("已应用全文首行缩进2字符")
            else:
                messagebox.showinfo("提示", "首行缩进未变化")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("操作失败", f"设置首行缩进失败：{exc}")

    def on_apply_header_text(self) -> None:
        if not self.docx_handler.is_loaded:
            messagebox.showinfo("提示", "请先打开文档")
            return
        text = self.header_entry.get().strip() if hasattr(self, "header_entry") else ""
        try:
            changed = self.docx_handler.apply_header_text(text)
            if changed:
                self._refresh_after_doc_mutation("已应用页眉")
            else:
                messagebox.showinfo("提示", "页眉未变化")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("操作失败", f"应用页眉失败：{exc}")

    def on_apply_footer_page_number(self) -> None:
        if not self.docx_handler.is_loaded:
            messagebox.showinfo("提示", "请先打开文档")
            return
        try:
            changed = self.docx_handler.apply_page_number_footer(start_page=1, font_name="Times New Roman", font_size_pt=10.5)
            if changed:
                self._refresh_after_doc_mutation("已设置页码（页脚居中，从1开始）")
            else:
                messagebox.showinfo("提示", "页码未变化")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("操作失败", f"设置页码失败：{exc}")

    def on_unlink_header_footer(self) -> None:
        if not self.docx_handler.is_loaded:
            messagebox.showinfo("提示", "请先打开文档")
            return
        try:
            changed = self.docx_handler.unlink_headers_footers_all_sections()
            if changed:
                self._refresh_after_doc_mutation("已断开所有节的页眉页脚链接")
            else:
                messagebox.showinfo("提示", "页眉页脚链接已是断开状态")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("操作失败", f"断开链接失败：{exc}")

    def on_switch_page_number_roman(self) -> None:
        if not self.docx_handler.is_loaded:
            messagebox.showinfo("提示", "请先打开文档")
            return
        try:
            changed = self.docx_handler.switch_page_number_format_bulk("roman")
            if changed:
                self._refresh_after_doc_mutation("已批量切换为罗马页码")
                self._show_field_update_tip_if_needed()
            else:
                messagebox.showinfo("提示", "未找到可切换的页码")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("操作失败", f"切换罗马页码失败：{exc}")

    def on_switch_page_number_arabic(self) -> None:
        if not self.docx_handler.is_loaded:
            messagebox.showinfo("提示", "请先打开文档")
            return
        try:
            changed = self.docx_handler.switch_page_number_format_bulk("arabic")
            if changed:
                self._refresh_after_doc_mutation("已批量切换为阿拉伯页码")
                self._show_field_update_tip_if_needed()
            else:
                messagebox.showinfo("提示", "未找到可切换的页码")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("操作失败", f"切换阿拉伯页码失败：{exc}")

    def on_one_click_section_and_page_number(self) -> None:
        if not self.docx_handler.is_loaded:
            messagebox.showinfo("提示", "请先打开文档")
            return
        try:
            result = self.docx_handler.one_click_section_and_page_number()
            self._refresh_after_doc_mutation(
                f"已完成一键分节与页码：识别{result['structures']}块，插入{result['section_breaks']}处分节符"
            )
            self._show_field_update_tip_if_needed()
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("操作失败", f"一键分节并设置页码失败：{exc}")

    def on_clear_header_footer(self) -> None:
        if not self.docx_handler.is_loaded:
            messagebox.showinfo("提示", "请先打开文档")
            return
        if not messagebox.askyesno("确认操作", "该操作将清空全部节的页眉页脚，是否继续？"):
            return
        try:
            changed = self.docx_handler.clear_header_footer()
            if changed:
                self._refresh_after_doc_mutation("已清空页眉页脚")
            else:
                messagebox.showinfo("提示", "页眉页脚未变化")
        except Exception as exc:  # noqa: BLE001
            self._show_action_error("操作失败", "清空页眉页脚未完成。", "请检查文档是否只读，或重试“一键分节并设置页码”。", exc)

    def _show_toc_precheck_dialog(self, check: dict[str, object]) -> bool:
        dialog = ctk.CTkToplevel(self)
        dialog.title("目录生成提示")
        dialog.geometry("520x260")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()

        frame = ctk.CTkFrame(dialog)
        frame.pack(fill="both", expand=True, padx=12, pady=12)
        message = (
            f"发现 {check.get('mismatch_count', 0)} 处标题分级/样式不规范，可能影响目录准确性。\n"
            "您可以继续生成目录，或先去修复标题。"
        )
        ctk.CTkLabel(frame, text=message, justify="left", anchor="w").pack(fill="x", padx=10, pady=(10, 8))
        samples = check.get("samples", [])
        if samples:
            sample_text = "\n".join(str(x) for x in samples[:4])
            ctk.CTkLabel(frame, text=f"示例：\n{sample_text}", justify="left", anchor="w").pack(fill="x", padx=10, pady=(0, 8))

        result = {"continue": False}

        def on_fix() -> None:
            result["continue"] = False
            dialog.destroy()

        def on_continue() -> None:
            result["continue"] = True
            dialog.destroy()

        btn_row = ctk.CTkFrame(frame, fg_color="transparent")
        btn_row.pack(fill="x", padx=10, pady=(8, 10))
        btn_row.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkButton(btn_row, text="去修复标题", command=on_fix).grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ctk.CTkButton(btn_row, text="跳过检查，强制生成目录", command=on_continue).grid(
            row=0, column=1, sticky="ew", padx=(4, 0)
        )
        self.wait_window(dialog)
        return bool(result["continue"])

    def _precheck_heading_styles_or_warn(self) -> bool:
        check = self.docx_handler.precheck_heading_styles(strict_style=False)
        if bool(check.get("ok")) and int(check.get("mismatch_count", 0)) == 0:
            return True
        return self._show_toc_precheck_dialog(check)

    def on_generate_toc(self) -> None:
        if not self.docx_handler.is_loaded:
            messagebox.showinfo("提示", "请先打开文档")
            return
        try:
            force = self._precheck_heading_styles_or_warn()
            if not force:
                return
            result, elapsed = self._run_long_task(
                "正在生成目录",
                lambda: self.docx_handler.auto_generate_toc(levels=self.toc_levels, mode="smart"),
            )
            if bool(result.get("ok")):
                self._refresh_after_doc_mutation(f"{str(result.get('message', '已生成目录'))}（{elapsed:.0f}ms）")
                self._show_field_update_tip_if_needed()
            else:
                messagebox.showwarning("提示", str(result.get("message", "生成目录失败")))
        except Exception as exc:  # noqa: BLE001
            self._show_action_error("操作失败", "自动生成目录失败。", "请先检查标题结构，必要时点击“去修复标题”。", exc)

    def on_update_toc_page_only(self) -> None:
        if not self.docx_handler.is_loaded:
            messagebox.showinfo("提示", "请先打开文档")
            return
        try:
            result, elapsed = self._run_long_task(
                "正在更新目录页码",
                lambda: self.docx_handler.update_toc(page_only=True, levels=self.toc_levels, mode="smart"),
            )
            if bool(result.get("ok")):
                self._refresh_after_doc_mutation(f"{str(result.get('message', '已更新目录页码'))}（{elapsed:.0f}ms）")
                self._show_field_update_tip_if_needed()
            else:
                messagebox.showwarning("提示", str(result.get("message", "更新目录失败")))
        except Exception as exc:  # noqa: BLE001
            self._show_action_error("操作失败", "更新目录页码失败。", "请确认已生成目录，或尝试“更新目录（整个目录）”。", exc)

    def on_update_toc_all(self) -> None:
        if not self.docx_handler.is_loaded:
            messagebox.showinfo("提示", "请先打开文档")
            return
        try:
            if not self._precheck_heading_styles_or_warn():
                return
            result, elapsed = self._run_long_task(
                "正在更新整个目录",
                lambda: self.docx_handler.update_toc(page_only=False, levels=self.toc_levels, mode="smart"),
            )
            if bool(result.get("ok")):
                self._refresh_after_doc_mutation(f"{str(result.get('message', '已更新整个目录'))}（{elapsed:.0f}ms）")
                self._show_field_update_tip_if_needed()
            else:
                messagebox.showwarning("提示", str(result.get("message", "更新目录失败")))
        except Exception as exc:  # noqa: BLE001
            self._show_action_error("操作失败", "更新整个目录失败。", "请检查标题与分节后重试。", exc)

    def on_check_toc_consistency(self) -> None:
        if not self.docx_handler.is_loaded:
            messagebox.showinfo("提示", "请先打开文档")
            return
        try:
            result, elapsed = self._run_long_task("正在校验目录页码", self.docx_handler.check_toc_page_consistency)
            if bool(result.get("ok")):
                self._update_hint(f"{str(result.get('message', '目录校验通过'))}（{elapsed:.0f}ms）")
                messagebox.showinfo("校验结果", str(result.get("message", "目录校验通过")))
            else:
                items = result.get("items", [])
                lines = [f"{it['title']}：目录{it['toc_page']}，正文{it['actual_page']}" for it in items[:6] if isinstance(it, dict)]
                detail = "\n".join(lines)
                messagebox.showwarning(
                    "校验结果",
                    f"{result.get('message', '目录页码存在不一致')}\n{detail}" if detail else str(result.get("message", "目录页码存在不一致")),
                )
        except Exception as exc:  # noqa: BLE001
            self._show_action_error("操作失败", "目录一致性校验失败。", "请先生成目录，再执行校验。", exc)

    def on_apply_reference_hanging_indent(self) -> None:
        if not self.docx_handler.is_loaded:
            messagebox.showinfo("提示", "请先打开文档")
            return
        try:
            total = self.docx_handler.get_reference_entry_count()
            if total == 0:
                messagebox.showinfo("提示", "未找到可设置的参考文献条目")
                return
            changed, elapsed = self._run_long_task(
                "正在设置参考文献悬挂缩进",
                lambda: self.docx_handler.apply_reference_hanging_indent(chars=2.0, base_font_pt=12.0),
            )
            if changed:
                self._refresh_after_doc_mutation(f"已设置参考文献悬挂缩进2字符（{elapsed:.0f}ms）")
            else:
                messagebox.showinfo("提示", "参考文献排版已符合悬挂缩进要求")
        except Exception as exc:  # noqa: BLE001
            self._show_action_error("操作失败", "设置悬挂缩进失败。", "请确认文档中存在“参考文献”章节。", exc)

    def on_format_references_gbt(self) -> None:
        if not self.docx_handler.is_loaded:
            messagebox.showinfo("提示", "请先打开文档")
            return
        mode_label = self.ref_mode_combo.get().strip() if hasattr(self, "ref_mode_combo") else "顺序编码制"
        mode = "order" if mode_label.startswith("顺序") else "author_year"
        try:
            result, elapsed = self._run_long_task(
                "正在格式化参考文献",
                lambda: self.docx_handler.format_references_gbt(mode=mode),
            )
            if bool(result.get("ok")):
                self._refresh_after_doc_mutation(f"{str(result.get('message', '已完成参考文献格式化'))}（{elapsed:.0f}ms）")
            else:
                messagebox.showwarning("提示", str(result.get("message", "参考文献格式化失败")))
        except Exception as exc:  # noqa: BLE001
            self._show_action_error("操作失败", "参考文献格式化失败。", "请先确认“参考文献”章节和条目格式。", exc)

    def on_validate_references_gbt(self) -> None:
        if not self.docx_handler.is_loaded:
            messagebox.showinfo("提示", "请先打开文档")
            return
        mode_label = self.ref_mode_combo.get().strip() if hasattr(self, "ref_mode_combo") else "顺序编码制"
        mode = "order" if mode_label.startswith("顺序") else "author_year"
        try:
            result, elapsed = self._run_long_task(
                "正在校验参考文献",
                lambda: self.docx_handler.validate_references_gbt(mode=mode),
            )
            if bool(result.get("ok")):
                messagebox.showinfo("校验结果", str(result.get("message", "参考文献校验通过")))
                self._update_hint(f"参考文献校验通过（{elapsed:.0f}ms）")
                return
            items = result.get("items", [])
            lines = [f"第{it['row']}条：{it['issue']}" for it in items[:8] if isinstance(it, dict)]
            detail = "\n".join(lines)
            messagebox.showwarning(
                "校验结果",
                f"{result.get('message', '存在格式问题')}\n{detail}" if detail else str(result.get("message", "存在格式问题")),
            )
        except Exception as exc:  # noqa: BLE001
            self._show_action_error("操作失败", "参考文献校验失败。", "请先执行“格式化参考文献”后重试。", exc)

    def on_auto_fix_references_by_validation(self) -> None:
        if not self.docx_handler.is_loaded:
            messagebox.showinfo("提示", "请先打开文档")
            return
        mode_label = self.ref_mode_combo.get().strip() if hasattr(self, "ref_mode_combo") else "顺序编码制"
        mode = "order" if mode_label.startswith("顺序") else "author_year"
        try:
            result, elapsed = self._run_long_task(
                "正在一键修复参考文献",
                lambda: self.docx_handler.auto_fix_references_by_validation(mode=mode),
            )
            if not bool(result.get("ok")):
                messagebox.showwarning("提示", str(result.get("message", "一键修改失败")))
                return
            self._refresh_after_doc_mutation(f"{str(result.get('message', '已完成参考文献一键修改'))}（{elapsed:.0f}ms）")
        except Exception as exc:  # noqa: BLE001
            self._show_action_error("操作失败", "一键修改参考文献失败。", "请先执行文献全面校验，再尝试一键修改。", exc)

    def on_operation_log_view_export(self) -> None:
        choice = messagebox.askyesnocancel(
            "操作日志",
            "选择“是”查看操作日志；选择“否”导出CSV；“取消”退出。",
        )
        if choice is None:
            return
        if choice:
            self._view_operation_logs_window()
        else:
            self._export_operation_logs_csv()

    def _view_operation_logs_window(self) -> None:
        win = ctk.CTkToplevel(self)
        win.title("操作日志")
        win.geometry("760x420")
        win.transient(self)
        text_box = ctk.CTkTextbox(win)
        text_box.pack(fill="both", expand=True, padx=10, pady=10)
        if not self.operation_logs:
            text_box.insert("1.0", "暂无操作日志。")
        else:
            lines = [
                f"{x['time']} | {x['type']} | {x.get('result', '')} | {x.get('duration_ms', '')}ms | 影响:{x.get('affected_count', '')} | {x['detail']}"
                for x in self.operation_logs
            ]
            text_box.insert("1.0", "\n".join(lines))
        text_box.configure(state="disabled")

    def _export_operation_logs_csv(self) -> None:
        if not self.operation_logs:
            messagebox.showinfo("提示", "暂无可导出的操作日志")
            return
        file_path = filedialog.asksaveasfilename(
            title="导出操作日志",
            defaultextension=".csv",
            filetypes=[("CSV 文件", "*.csv"), ("所有文件", "*.*")],
        )
        if not file_path:
            return
        try:
            with open(file_path, "w", newline="", encoding="utf-8-sig") as fp:
                writer = csv.DictWriter(fp, fieldnames=["time", "type", "result", "duration_ms", "affected_count", "detail"])
                writer.writeheader()
                writer.writerows(self.operation_logs)
            messagebox.showinfo("提示", "操作日志已导出")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("导出失败", f"导出操作日志失败：{exc}")

    def on_school_rules_import_export(self) -> None:
        choice = messagebox.askyesnocancel(
            "学校规则配置",
            "选择“是”导入规则JSON；选择“否”导出当前规则JSON；“取消”退出。",
        )
        if choice is None:
            return
        if choice:
            self._import_school_rules_json()
        else:
            self._export_school_rules_json()

    def _export_school_rules_json(self) -> None:
        file_path = filedialog.asksaveasfilename(
            title="导出学校规则配置",
            defaultextension=".json",
            filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")],
        )
        if not file_path:
            return
        data = {
            "version": 1,
            "exported_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "rules": self._collect_current_rules(),
        }
        try:
            Path(file_path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            messagebox.showinfo("提示", "学校规则配置已导出")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("导出失败", f"导出规则失败：{exc}")

    def _import_school_rules_json(self) -> None:
        file_path = filedialog.askopenfilename(
            title="导入学校规则配置",
            filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")],
        )
        if not file_path:
            return
        try:
            data = json.loads(Path(file_path).read_text(encoding="utf-8"))
            rules = data.get("rules", data) if isinstance(data, dict) else {}
            if not isinstance(rules, dict):
                raise ValueError("配置格式无效")
            self._apply_rules_to_ui(rules)
            self._log_operation("规则导入", f"导入学校规则：{Path(file_path).name}")
            messagebox.showinfo("提示", "学校规则配置已导入并应用")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("导入失败", f"导入规则失败：{exc}")

    def on_save_user_preferences(self) -> None:
        data = {
            "version": 1,
            "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "rules": self._collect_current_rules(),
        }
        try:
            self.preference_file.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            self._log_operation("偏好保存", "已保存个人偏好设置")
            messagebox.showinfo("提示", "个人偏好设置已保存")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("保存失败", f"保存个人偏好失败：{exc}")

    def on_restore_default_preferences(self) -> None:
        self.font_combo.set("宋体")
        self.size_combo.set("五号")
        self.spacing_combo.set("1.5倍")
        self.margin_combo.set("普通")
        if hasattr(self, "ref_mode_combo"):
            self.ref_mode_combo.set("顺序编码制")
        self.preview_accuracy_mode = "快速模式"
        if self.preview_mode_combo is not None:
            self.preview_mode_combo.set(self.preview_accuracy_mode)
        self.auto_backup_enabled = True
        self.docx_handler.set_auto_backup_enabled(True)
        if self.auto_backup_switch is not None:
            self.auto_backup_switch.select()
        self.on_save_user_preferences()
        self._refresh_document_view(preserve_selection=True)
        self._update_hint("已恢复默认偏好")

    def on_preview_mode_change(self, choice: str) -> None:
        mode = choice.strip()
        if mode not in self.PREVIEW_MODE_OPTIONS:
            return
        self.preview_accuracy_mode = mode
        self._refresh_document_view(preserve_selection=True)
        self._update_hint(f"已切换预览模式：{mode}")

    def on_toggle_auto_backup(self) -> None:
        enabled = bool(self.auto_backup_switch.get()) if self.auto_backup_switch is not None else True
        self.auto_backup_enabled = enabled
        self.docx_handler.set_auto_backup_enabled(enabled)
        self._update_hint("已开启自动备份(.bak)" if enabled else "已关闭自动备份(.bak)")

    def _sync_style_toggle_state_from_selection(self) -> None:
        if not self.selected_indices or not self.paragraph_data:
            self.active_bold = False
            self.active_italic = False
            self.active_underline = False
            self.active_alignment = ""
            self._refresh_style_toggle_buttons()
            self._refresh_align_buttons()
            return
        selected = [p for p in self.paragraph_data if int(p.get("index", -1)) in self.selected_indices]
        if not selected:
            self.active_bold = False
            self.active_italic = False
            self.active_underline = False
            self.active_alignment = ""
            self._refresh_style_toggle_buttons()
            self._refresh_align_buttons()
            return
        self.active_bold = all(bool(p.get("bold")) for p in selected)
        self.active_italic = all(bool(p.get("italic")) for p in selected)
        self.active_underline = all(bool(p.get("underline")) for p in selected)
        alignments = {str(p.get("alignment") or "left").lower() for p in selected}
        self.active_alignment = next(iter(alignments)) if len(alignments) == 1 else ""
        self._refresh_style_toggle_buttons()
        self._refresh_align_buttons()

    def _refresh_style_toggle_buttons(self) -> None:
        selected_color = "#dbeafe"
        default_color = "transparent"
        if self.btn_style_bold is not None:
            self.btn_style_bold.configure(fg_color=selected_color if self.active_bold else default_color)
        if self.btn_style_italic is not None:
            self.btn_style_italic.configure(fg_color=selected_color if self.active_italic else default_color)
        if self.btn_style_underline is not None:
            self.btn_style_underline.configure(fg_color=selected_color if self.active_underline else default_color)

    def _refresh_align_buttons(self) -> None:
        active = "#111111"
        inactive = "#d8d8d8"
        active_bg = "#f9f9f9"
        inactive_bg = "transparent"
        mapping = {
            "left": self.btn_align_left,
            "center": self.btn_align_center,
            "right": self.btn_align_right,
            "justify": self.btn_align_justify,
        }
        for key, btn in mapping.items():
            if btn is None:
                continue
            is_active = self.active_alignment == key
            btn.configure(
                border_color=active if is_active else inactive,
                fg_color=active_bg if is_active else inactive_bg,
                border_width=2 if is_active else 1,
            )

    def on_toggle_bold(self) -> None:
        target_indices = self._selected_or_warn()
        if not target_indices:
            return
        next_value = not self.active_bold
        try:
            changed = self.docx_handler.apply_bold(target_indices, next_value)
            if changed:
                self._refresh_after_doc_mutation("已切换粗体")
            else:
                messagebox.showinfo("提示", "粗体状态未变化")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("操作失败", f"切换粗体失败：{exc}")

    def on_toggle_italic(self) -> None:
        target_indices = self._selected_or_warn()
        if not target_indices:
            return
        next_value = not self.active_italic
        try:
            changed = self.docx_handler.apply_italic(target_indices, next_value)
            if changed:
                self._refresh_after_doc_mutation("已切换斜体")
            else:
                messagebox.showinfo("提示", "斜体状态未变化")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("操作失败", f"切换斜体失败：{exc}")

    def on_toggle_underline(self) -> None:
        target_indices = self._selected_or_warn()
        if not target_indices:
            return
        next_value = not self.active_underline
        try:
            changed = self.docx_handler.apply_underline(target_indices, next_value)
            if changed:
                self._refresh_after_doc_mutation("已切换下划线")
            else:
                messagebox.showinfo("提示", "下划线状态未变化")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("操作失败", f"切换下划线失败：{exc}")

    def on_bold(self) -> None:
        self.on_toggle_bold()

    def on_unbold(self) -> None:
        self.on_toggle_bold()

    def on_align_left(self) -> None:
        self._apply_alignment("left", "左对齐")

    def on_align_center(self) -> None:
        self._apply_alignment("center", "居中对齐")

    def on_align_right(self) -> None:
        self._apply_alignment("right", "右对齐")

    def on_align_justify(self) -> None:
        self._apply_alignment("justify", "两端对齐")

    def _apply_alignment(self, alignment: str, label: str) -> None:
        target_indices = self._selected_or_warn()
        if not target_indices:
            return
        try:
            changed = self.docx_handler.apply_alignment(target_indices, alignment)
            if changed:
                self._refresh_after_doc_mutation(f"已应用{label}")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("操作失败", f"对齐失败：{exc}")

    def on_clean_hash(self) -> None:
        self._apply_text_clean("hash")

    def on_clean_star(self) -> None:
        self._apply_text_clean("star")

    def on_clean_dash(self) -> None:
        self._apply_text_clean("dash")

    def on_clean_placeholder_dot(self) -> None:
        self._apply_text_clean("placeholder_dot")

    def on_clean_keep_num_dash(self) -> None:
        self._apply_text_clean("keep_num_dash")

    def on_clean_empty(self) -> None:
        self._apply_text_clean("remove_empty")

    def _apply_text_clean(self, action: str) -> None:
        try:
            changed = 0
            if action == "remove_empty":
                empty_count = self.docx_handler.count_empty_paragraphs()
                if empty_count <= 0:
                    messagebox.showinfo("提示", "当前文档没有空段落")
                    return
                if not messagebox.askyesno("确认删除", f"检测到约 {empty_count} 个空段落，是否继续删除？"):
                    return
                changed = self.docx_handler.clean_remove_empty_paragraphs()
                reason = "已删除空段落"
                self._refresh_document_view(preserve_selection=True)
                self._update_hint(reason)
            else:
                target_indices = self._selected_or_warn()
                if not target_indices:
                    return
                if action == "hash":
                    changed = self.docx_handler.clean_delete_char(target_indices, "#")
                    reason = "已删除#"
                elif action == "star":
                    changed = self.docx_handler.clean_delete_char(target_indices, "*")
                    reason = "已删除*"
                elif action == "dash":
                    changed = self.docx_handler.clean_delete_char(target_indices, "-")
                    reason = "已删除-"
                elif action == "placeholder_dot":
                    changed = self.docx_handler.clean_delete_placeholder_dot(target_indices)
                    reason = "已删除占位."
                elif action == "keep_num_dash":
                    changed = self.docx_handler.clean_keep_numeric_dash(target_indices)
                    reason = "已保留数字间-"
                else:
                    return
                self._refresh_after_doc_mutation(reason)

            if changed == 0:
                messagebox.showinfo("提示", "清理未产生变化")
        except Exception as exc:  # noqa: BLE001
            self._show_action_error("清理失败", "文本清理未完成。", "请检查是否已选中段落，并重试。", exc)

    def on_undo_shortcut(self, _event: tk.Event | None = None) -> str:
        self._do_undo()
        return "break"

    def on_redo_shortcut(self, _event: tk.Event | None = None) -> str:
        self._do_redo()
        return "break"

    def on_undo_click(self) -> None:
        self._do_undo()

    def on_redo_click(self) -> None:
        self._do_redo()

    def _do_undo(self) -> None:
        if not self.docx_handler.is_loaded:
            return
        try:
            ok = self.docx_handler.undo_last_action()
            if not ok:
                self._update_hint("没有可撤销的上一步操作")
                return
            self._refresh_document_view(preserve_selection=False)
            self._update_hint("已撤销上一步操作")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("撤销失败", f"撤销上一步失败：{exc}")

    def _do_redo(self) -> None:
        if not self.docx_handler.is_loaded:
            return
        try:
            ok = self.docx_handler.redo_last_action()
            if not ok:
                self._update_hint("没有可重做的操作")
                return
            self._refresh_document_view(preserve_selection=False)
            self._update_hint("已重做上一步操作")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("重做失败", f"重做上一步失败：{exc}")


def main() -> None:
    app = FormatApp()
    app.mainloop()

