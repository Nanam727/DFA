"""Application state management."""

from __future__ import annotations

from pathlib import Path
from typing import Any
import tkinter as tk


class AppState:
    """Manages application state and selection data."""

    def __init__(self) -> None:
        # File and document state
        self.current_file: str | None = None

        # Selection state
        self.selected_indices: set[int] = set()
        self.last_clicked_index: int | None = None

        # Preview state
        self.current_view_page: int = 1
        self.total_view_pages: int = 1
        self.preview_zoom_percent: int = 100

        # Document layout state
        self.preview_margin_override_mm: dict[str, float] | None = None
        self.current_doc_margins_mm: dict[str, float] = {
            "top": 25.4,
            "bottom": 25.4,
            "left": 31.8,
            "right": 31.8,
        }
        self.current_margin_preset_name: str = "普通"

        # UI state
        self.tool_group_vars: dict[str, tk.BooleanVar] = {}
        self.tool_group_frames: dict[str, list[tk.Frame]] = {}
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

        # Paragraph data
        self.para_ranges: dict[int, tuple[tk.Text, str, str]] = {}
        self.left_para_ranges: dict[int, tuple[str, str]] = {}
        self.line_to_para_index_left: dict[int, int] = {}
        self.line_to_para_index_right: dict[tk.Text, dict[int, int]] = {}
        self.paragraph_data: list[dict[str, object]] = []

        # Operation tracking
        self.operation_logs: list[dict[str, str]] = []
        self.preference_file = Path.cwd() / ".adfa_user_prefs.json"

        # TOC settings
        self.toc_levels: int = 3
        self.default_school_template: str = "通用模板"

        # 论文结构选取（论文类型 + 勾选的章节 id）
        self.structure_selection: dict[str, Any] = {
            "paper_type": "undergrad_liberal",
            "selected": [],
        }

    @classmethod
    def get_constants(cls) -> dict[str, Any]:
        """Get application constants."""
        return {
            "A4_WIDTH_MM": 210.0,
            "A4_HEIGHT_MM": 297.0,
            "A4_WIDTH_IN": 210.0 / 25.4,
            "A4_HEIGHT_IN": 297.0 / 25.4,
            "FONT_OPTIONS": [
                "微软雅黑",
                "宋体",
                "Calibri",
                "Times New Roman",
                "黑体",
                "楷体",
                "楷体GB2312",
            ],
            "WORD_MARGIN_PRESETS_MM": {
                "普通": {"top": 25.4, "bottom": 25.4, "left": 31.8, "right": 31.8},
                "窄": {"top": 12.7, "bottom": 12.7, "left": 12.7, "right": 12.7},
                "宽": {"top": 25.4, "bottom": 25.4, "left": 50.8, "right": 50.8},
            },
            "WORD_DEFAULT_FONT_PT": {
                "body": 10.5,
                "h1": 22.0,
                "h2": 15.0,
                "h3": 14.0,
                "h4": 12.0,
                "h5": 10.5,
            },
            "PREVIEW_BASELINE_CHARS_PER_LINE": 42,
            "PREVIEW_BASELINE_LINES_PER_PAGE": 46,
            "PREVIEW_SAFE_EDGE_MM": {
                "left": 3.0,
                "right": 3.0,
                "top": 2.2,
                "bottom": 2.2,
            },
            "PREVIEW_SAFE_CHAR_BUFFER": 1.8,
            "PREVIEW_TEXT_INNER_PAD_X": 2,
            "PREVIEW_TEXT_INNER_PAD_Y": 2,
            "UI_BG": "#f4f7fb",
            "UI_CARD_BG": "#ffffff",
            "UI_BORDER": "#d8e1ef",
            "UI_ACCENT": "#2f5fa7",
            "UI_ACCENT_HOVER": "#244b83",
            "UI_TEXT": "#1f2a3a",
            "PREVIEW_MODE_OPTIONS": ["快速模式", "高精度模式"],
        }
