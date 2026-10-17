"""DOCX data handling and paragraph classification."""

from __future__ import annotations

from copy import deepcopy
from io import BytesIO
import math
import re
from pathlib import Path
from typing import Any

from docx import Document
from docx.document import Document as DocxDocument
from docx.enum.text import WD_BREAK, WD_LINE_SPACING, WD_PARAGRAPH_ALIGNMENT, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt
from docx.text.paragraph import Paragraph


class DocxHandler:
    """Load DOCX, expose paragraph metadata, and query selection ranges."""

    LEVEL_BIG_TITLE = 5
    LEVEL_BODY = 0
    LEVEL_H1 = 1
    LEVEL_H2 = 2
    LEVEL_H3 = 3
    LEVEL_H4 = 4

    # New thesis heading rules (with backward compatibility):
    # H1: 第 1 章 / 第 一 章
    # H1 extra: 1绪论 / 2理论基础 / 3现状...
    # H2: 1.1
    # H3: 1.1.1
    # H4: （1）
    _PATTERN_H1 = re.compile(r"^\s*第\s*([0-9]+|[一二三四五六七八九十百千万两零〇]+)\s*章(?:\s+|$)")
    _PATTERN_H1_NUM_TEXT = re.compile(r"^\s*\d+(?!\s*\.)\s*[^\W\d_].+")
    _PATTERN_H2 = re.compile(r"^\s*\d+\.\d+(?!\.)")
    _PATTERN_H3 = re.compile(r"^\s*\d+\.\d+\.\d+(?!\.)")
    _PATTERN_H4 = re.compile(r"^\s*[（(]\d+[）)]")
    # Legacy-compatible patterns retained for older documents.
    _PATTERN_H1_LEGACY = re.compile(r"^\s*[一二三四五六七八九十百千万两零〇]+[、.]")
    _PATTERN_H2_LEGACY = re.compile(r"^\s*[（(][一二三四五六七八九十百千万两零〇]+[）)]")
    _PATTERN_H3_LEGACY = re.compile(r"^\s*\d+[、.]")
    _CJK_PATTERN = re.compile(r"[\u4e00-\u9fff\u3400-\u4dbf\uff00-\uffef]")
    _H1_NUM_TEXT_KEYWORDS = (
        "绪论",
        "理论",
        "理论基础",
        "研究背景",
        "研究意义",
        "研究方法",
        "现状",
        "成因",
        "分析",
        "对策",
        "启示",
        "结论",
        "建议",
        "总结",
    )
    PAGE_BOTTOM_SAFE_MIN_PT = 2.0
    PAGE_BOTTOM_SAFE_MAX_PT = 10.0

    def __init__(self) -> None:
        self.file_path: Path | None = None
        self._document: DocxDocument | None = None
        self._paragraphs: list[dict[str, Any]] = []
        self._undo_stack: list[bytes] = []
        self._redo_stack: list[bytes] = []
        self._last_section_labels: list[str] = []
        self._auto_backup_enabled: bool = True

    @property
    def is_loaded(self) -> bool:
        return self._document is not None

    def load(self, path: str | Path) -> None:
        """Load .docx file and cache paragraph metadata."""
        target = Path(path)
        document = Document(str(target))
        self.file_path = target
        self._document = document
        self.enforce_word_pagination_rules()
        self._paragraphs = self._collect_paragraphs(document.paragraphs)
        self._undo_stack.clear()
        self._redo_stack.clear()

    def save(self, path: str | Path | None = None) -> None:
        if self._document is None:
            raise RuntimeError("Document is not loaded.")

        if path is not None:
            self.file_path = Path(path)
        if self.file_path is None:
            raise RuntimeError("No target path for saving.")
        # Ensure non-splitting settings are persisted before save.
        self.enforce_word_pagination_rules()
        if self._auto_backup_enabled and self.file_path.exists():
            backup_path = self.file_path.with_suffix(self.file_path.suffix + ".bak")
            backup_path.write_bytes(self.file_path.read_bytes())
        self._document.save(str(self.file_path))

    def set_auto_backup_enabled(self, enabled: bool) -> None:
        self._auto_backup_enabled = bool(enabled)

    def get_undo_redo_counts(self) -> dict[str, int]:
        return {"undo": len(self._undo_stack), "redo": len(self._redo_stack)}

    def enforce_keep_together_for_all(self) -> int:
        """Backward-compatible wrapper."""
        return self.enforce_word_pagination_rules()

    def enforce_word_pagination_rules(self) -> int:
        doc = self._require_document()
        changed = 0
        for para in doc.paragraphs:
            fmt = para.paragraph_format
            # Avoid Word left-side black square marker in show-format mode:
            # do not force "keep lines together / keep with next" globally.
            if fmt.keep_together is not False:
                fmt.keep_together = False
                changed += 1
            if fmt.widow_control is not False:
                fmt.widow_control = False
                changed += 1
            if fmt.keep_with_next is not False:
                fmt.keep_with_next = False
                changed += 1
            # 行距属于用户格式数据，此处一律不得改写。
            # 历史版本会把全文行距强制改成 1.25（单倍/双倍/固定值全被摧毁），
            # 导致仅打开文档就损坏用户格式，该逻辑已移除。
            if self._set_snap_to_grid_false(para._p):
                changed += 1

        # Apply same paragraph rules to paragraphs inside textboxes.
        changed += self._apply_xml_paragraph_rules_for_textboxes(doc)
        # Disable document grid in each section.
        changed += self._disable_document_grid(doc)
        # Disallow table rows from breaking across pages.
        changed += self._disable_table_row_breaks(doc.tables)

        if changed:
            self._paragraphs = self._collect_paragraphs(doc.paragraphs)
        return changed

    def get_page_layout_mm(self) -> dict[str, float]:
        doc = self._require_document()
        if not doc.sections:
            return {
                "page_width_mm": 210.0,
                "page_height_mm": 297.0,
                "margin_left_mm": 31.8,
                "margin_right_mm": 31.8,
                "margin_top_mm": 25.4,
                "margin_bottom_mm": 25.4,
            }
        section = doc.sections[0]
        return {
            "page_width_mm": float(section.page_width.mm),
            "page_height_mm": float(section.page_height.mm),
            "margin_left_mm": float(section.left_margin.mm),
            "margin_right_mm": float(section.right_margin.mm),
            "margin_top_mm": float(section.top_margin.mm),
            "margin_bottom_mm": float(section.bottom_margin.mm),
        }

    def apply_page_margin_mm(self, margin_mm: float) -> bool:
        return self.apply_page_margins_mm(
            top_mm=margin_mm,
            bottom_mm=margin_mm,
            left_mm=margin_mm,
            right_mm=margin_mm,
        )

    def apply_page_margins_mm(self, top_mm: float, bottom_mm: float, left_mm: float, right_mm: float) -> bool:
        doc = self._require_document()
        if min(top_mm, bottom_mm, left_mm, right_mm) <= 0:
            return False
        self._push_undo_snapshot()
        changed = False
        for section in doc.sections:
            old = (
                float(section.top_margin.mm),
                float(section.bottom_margin.mm),
                float(section.left_margin.mm),
                float(section.right_margin.mm),
            )
            section.top_margin = Mm(top_mm)
            section.bottom_margin = Mm(bottom_mm)
            section.left_margin = Mm(left_mm)
            section.right_margin = Mm(right_mm)
            if old != (top_mm, bottom_mm, left_mm, right_mm):
                changed = True
        if changed:
            self._paragraphs = self._collect_paragraphs(doc.paragraphs)
        return changed

    def apply_font_size(self, indices: set[int], font_size_pt: float) -> int:
        if not indices:
            return 0
        doc = self._require_document()
        self._push_undo_snapshot()
        changed = 0
        for idx in sorted(indices):
            if idx < 0 or idx >= len(doc.paragraphs):
                continue
            para = doc.paragraphs[idx]
            changed += self._set_paragraph_font_size(para, font_size_pt)
        if changed:
            self._paragraphs = self._collect_paragraphs(doc.paragraphs)
        return changed

    def apply_font_size_step(
        self, indices: set[int], size_points: list[float], step: int
    ) -> int:
        """按每段当前字号在阶梯上独立进/退一档。

        选中多种字号时不强制统一：每种字号各自放大/缩小一号，
        保持段落间的相对大小差异。返回实际改动的 run 数。
        """
        if not indices or step == 0:
            return 0
        points = sorted({float(p) for p in size_points if float(p) > 0})
        if not points:
            return 0
        doc = self._require_document()
        self._push_undo_snapshot()
        changed = 0
        for idx in sorted(indices):
            if idx < 0 or idx >= len(doc.paragraphs):
                continue
            para = doc.paragraphs[idx]
            current = self._extract_font_size_pt(para)
            if current is None:
                # 段内未显式设置字号时按正文默认字号参与步进。
                current = 12.0
            target = self._step_size_on_ladder(current, points, step)
            if target is None or abs(target - current) < 1e-6:
                continue
            changed += self._set_paragraph_font_size(para, target)
        if changed:
            self._paragraphs = self._collect_paragraphs(doc.paragraphs)
        return changed

    @staticmethod
    def _step_size_on_ladder(
        current: float, size_points: list[float], step: int
    ) -> float | None:
        """取 current 在阶梯上的下一档（step>0）或上一档（step<0）。

        current 不在阶梯上（自定义字号）时，按大于/小于它的最近档进退；
       已在两端则返回 None（不动）。
        """
        points = sorted(float(p) for p in size_points)
        if step > 0:
            for point in points:
                if point > current + 1e-6:
                    return point
            return None
        if step < 0:
            target: float | None = None
            for point in points:
                if point < current - 1e-6:
                    target = point
                else:
                    break
            return target
        return None

    def apply_font_name(self, indices: set[int], font_name: str) -> int:
        if not indices:
            return 0
        doc = self._require_document()
        if not font_name.strip():
            return 0
        self._push_undo_snapshot()
        changed = 0
        for idx in sorted(indices):
            if idx < 0 or idx >= len(doc.paragraphs):
                continue
            para = doc.paragraphs[idx]
            changed += self._set_paragraph_font_name(para, font_name.strip())
        if changed:
            self._paragraphs = self._collect_paragraphs(doc.paragraphs)
        return changed

    def apply_line_spacing(self, indices: set[int], line_spacing: float | str) -> int:
        if not indices:
            return 0
        doc = self._require_document()
        self._push_undo_snapshot()
        changed = 0
        for idx in sorted(indices):
            if idx < 0 or idx >= len(doc.paragraphs):
                continue
            para = doc.paragraphs[idx]
            if isinstance(line_spacing, str) and line_spacing.lower() == "fixed18pt":
                # 固定 18pt 必须用 EXACTLY 精确值实现；
                # 历史版本误用 MULTIPLE 1.5，且会被分页规则改写，双重失效。
                para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                para.paragraph_format.line_spacing = Pt(18)
            else:
                para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
                para.paragraph_format.line_spacing = float(line_spacing)
            changed += 1
        if changed:
            self.enforce_word_pagination_rules()
            self._paragraphs = self._collect_paragraphs(doc.paragraphs)
        return changed

    def apply_bold(self, indices: set[int], bold: bool) -> int:
        if not indices:
            return 0
        doc = self._require_document()
        self._push_undo_snapshot()
        changed = 0
        for idx in sorted(indices):
            if idx < 0 or idx >= len(doc.paragraphs):
                continue
            para = doc.paragraphs[idx]
            changed += self._set_paragraph_bold(para, bold)
        if changed:
            self._paragraphs = self._collect_paragraphs(doc.paragraphs)
        return changed

    def apply_italic(self, indices: set[int], italic: bool) -> int:
        if not indices:
            return 0
        doc = self._require_document()
        self._push_undo_snapshot()
        changed = 0
        for idx in sorted(indices):
            if idx < 0 or idx >= len(doc.paragraphs):
                continue
            para = doc.paragraphs[idx]
            changed += self._set_paragraph_italic(para, italic)
        if changed:
            self._paragraphs = self._collect_paragraphs(doc.paragraphs)
        return changed

    def apply_underline(self, indices: set[int], underline: bool) -> int:
        if not indices:
            return 0
        doc = self._require_document()
        self._push_undo_snapshot()
        changed = 0
        for idx in sorted(indices):
            if idx < 0 or idx >= len(doc.paragraphs):
                continue
            para = doc.paragraphs[idx]
            changed += self._set_paragraph_underline(para, underline)
        if changed:
            self._paragraphs = self._collect_paragraphs(doc.paragraphs)
        return changed

    def find_text(self, keyword: str, case_sensitive: bool = False, whole_word: bool = False) -> list[int]:
        doc = self._require_document()
        query = keyword.strip()
        if not query:
            return []
        pattern = self._build_search_pattern(query, case_sensitive=case_sensitive, whole_word=whole_word)
        matched_indices: list[int] = []
        for idx, para in enumerate(doc.paragraphs):
            text = para.text or ""
            if not text:
                continue
            if pattern.search(text):
                matched_indices.append(idx)
        return matched_indices

    def replace_text(
        self,
        keyword: str,
        replacement: str,
        case_sensitive: bool = False,
        whole_word: bool = False,
        indices: set[int] | None = None,
        max_replacements_per_para: int = 0,
    ) -> dict[str, int]:
        doc = self._require_document()
        query = keyword.strip()
        if not query:
            return {"changed_paragraphs": 0, "replaced_count": 0}
        pattern = self._build_search_pattern(query, case_sensitive=case_sensitive, whole_word=whole_word)
        target_indices = sorted(indices) if indices else list(range(len(doc.paragraphs)))
        changed_paragraphs = 0
        replaced_count = 0
        did_snapshot = False
        count_arg = 0 if max_replacements_per_para <= 0 else int(max_replacements_per_para)
        for idx in target_indices:
            if idx < 0 or idx >= len(doc.paragraphs):
                continue
            para = doc.paragraphs[idx]
            original = para.text or ""
            if not original:
                continue
            updated, num = pattern.subn(replacement, original, count=count_arg)
            if num <= 0:
                continue
            if not did_snapshot:
                self._push_undo_snapshot()
                did_snapshot = True
            self._replace_paragraph_text(para, updated)
            changed_paragraphs += 1
            replaced_count += num
        if changed_paragraphs:
            self._paragraphs = self._collect_paragraphs(doc.paragraphs)
        return {"changed_paragraphs": changed_paragraphs, "replaced_count": replaced_count}

    def apply_alignment(self, indices: set[int], alignment: str) -> int:
        if not indices:
            return 0
        doc = self._require_document()
        align_map = {
            "left": WD_PARAGRAPH_ALIGNMENT.LEFT,
            "center": WD_PARAGRAPH_ALIGNMENT.CENTER,
            "right": WD_PARAGRAPH_ALIGNMENT.RIGHT,
            "justify": WD_PARAGRAPH_ALIGNMENT.JUSTIFY,
        }
        align_value = align_map.get(alignment.lower())
        if align_value is None:
            return 0

        targets = [
            doc.paragraphs[idx]
            for idx in sorted(indices)
            if 0 <= idx < len(doc.paragraphs)
        ]
        # 无实际变化时不压入撤销快照、不计成功——历史版本会对
        # “已是对齐状态”的段落也推送快照并提示成功。
        if not any(para.alignment != align_value for para in targets):
            return 0
        self._push_undo_snapshot()
        changed = 0
        for para in targets:
            if para.alignment != align_value:
                para.alignment = align_value
                changed += 1
        if changed:
            self._paragraphs = self._collect_paragraphs(doc.paragraphs)
        return changed

    def apply_first_line_indent_chars(self, chars: float = 2.0, base_font_pt: float = 12.0) -> int:
        doc = self._require_document()
        if chars <= 0 or base_font_pt <= 0:
            return 0
        self._push_undo_snapshot()
        indent_pt = chars * base_font_pt
        changed = 0
        for para in doc.paragraphs:
            if not (para.text or "").strip():
                continue
            old = para.paragraph_format.first_line_indent
            old_pt = float(old.pt) if old is not None else 0.0
            if abs(old_pt - indent_pt) <= 0.2:
                continue
            para.paragraph_format.first_line_indent = Pt(indent_pt)
            changed += 1
        if changed:
            self._paragraphs = self._collect_paragraphs(doc.paragraphs)
        return changed

    def apply_header_text(self, text: str) -> int:
        doc = self._require_document()
        self._push_undo_snapshot()
        changed = 0
        for section in doc.sections:
            section.header.is_linked_to_previous = False
            para = section.header.paragraphs[0] if section.header.paragraphs else section.header.add_paragraph()
            old_text = para.text or ""
            if old_text != text:
                para.text = text
                para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                changed += 1
        if changed:
            self._paragraphs = self._collect_paragraphs(doc.paragraphs)
        return changed

    def apply_page_number_footer(self, start_page: int = 1, font_name: str = "Times New Roman", font_size_pt: float = 10.5) -> int:
        doc = self._require_document()
        self._push_undo_snapshot()
        changed = 0
        for section in doc.sections:
            section.footer.is_linked_to_previous = False
            para = section.footer.paragraphs[0] if section.footer.paragraphs else section.footer.add_paragraph()
            self._clear_paragraph_content(para)
            para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run = para.add_run()
            run.font.name = font_name
            run.font.size = Pt(font_size_pt)
            if run._element.rPr is not None and run._element.rPr.rFonts is not None:
                rfonts = run._element.rPr.rFonts
            else:
                rpr = run._element.get_or_add_rPr()
                rfonts = rpr.get_or_add_rFonts()
            rfonts.set(qn("w:ascii"), font_name)
            rfonts.set(qn("w:hAnsi"), font_name)
            self._append_field(run, "PAGE")
            self._set_section_page_start(section, start_page)
            changed += 1
        if changed:
            self._paragraphs = self._collect_paragraphs(doc.paragraphs)
        return changed

    def clear_header_footer(self) -> int:
        doc = self._require_document()
        self._push_undo_snapshot()
        changed = 0
        for section in doc.sections:
            section.header.is_linked_to_previous = False
            section.footer.is_linked_to_previous = False
            h_para = section.header.paragraphs[0] if section.header.paragraphs else section.header.add_paragraph()
            f_para = section.footer.paragraphs[0] if section.footer.paragraphs else section.footer.add_paragraph()
            if (h_para.text or "").strip():
                self._clear_paragraph_content(h_para)
                changed += 1
            if (f_para.text or "").strip() or f_para._p.xpath(".//w:fldChar"):
                self._clear_paragraph_content(f_para)
                changed += 1
            self._clear_section_page_start(section)
        if changed:
            self._paragraphs = self._collect_paragraphs(doc.paragraphs)
        return changed

    def precheck_heading_styles(self, strict_style: bool = False) -> dict[str, object]:
        self._refresh_paragraph_cache()
        items = [p for p in self._paragraphs if int(p.get("detected_level", 0)) in {self.LEVEL_H1, self.LEVEL_H2, self.LEVEL_H3}]
        h1_count = sum(1 for p in items if int(p.get("detected_level", 0)) == self.LEVEL_H1)
        warnings: list[str] = []

        # Keep big-title permanently out of H1 check pool.
        # _paragraphs already uses detected_level=LEVEL_BIG_TITLE for the total title.
        if h1_count == 0:
            warnings.append("未识别到有效一级标题（H1）")

        # Check level jump with original detected levels (non-blocking warning).
        prev_level = 0
        for p in items:
            level = int(p.get("detected_level", 0))
            if prev_level and level - prev_level > 1:
                warnings.append(f"存在标题跳级：{str(p.get('text', ''))[:24]}")
                break
            prev_level = level

        if strict_style:
            doc = self._require_document()
            style_mismatches = 0
            samples: list[str] = []
            for para in doc.paragraphs:
                text = (para.text or "").strip()
                if not text:
                    continue
                expected = self._detect_level_by_numbering(text)
                if expected not in {self.LEVEL_H1, self.LEVEL_H2, self.LEVEL_H3}:
                    continue
                style_name = (para.style.name if para.style is not None else "").lower()
                if expected == self.LEVEL_H1 and not ("heading 1" in style_name or "标题 1" in style_name):
                    style_mismatches += 1
                    if len(samples) < 5:
                        samples.append(f"H1 样式不匹配：{text[:30]}")
                if expected == self.LEVEL_H2 and not ("heading 2" in style_name or "标题 2" in style_name):
                    style_mismatches += 1
                    if len(samples) < 5:
                        samples.append(f"H2 样式不匹配：{text[:30]}")
                if expected == self.LEVEL_H3 and not ("heading 3" in style_name or "标题 3" in style_name):
                    style_mismatches += 1
                    if len(samples) < 5:
                        samples.append(f"H3 样式不匹配：{text[:30]}")
            if style_mismatches > 0:
                warnings.append(f"存在 {style_mismatches} 处标题样式未绑定 Word 内置 Heading")
                warnings.extend(samples)

        return {
            "ok": h1_count > 0,
            "checked": len(items),
            "mismatch_count": len(warnings),
            "samples": warnings[:8],
            "h1_count": h1_count,
        }

    def auto_generate_toc(self, levels: int = 3, mode: str = "smart") -> dict[str, object]:
        doc = self._require_document()
        self._push_undo_snapshot()
        levels = max(1, min(9, int(levels)))
        mode = mode.strip().lower()

        if mode == "smart":
            self._sync_heading_styles_from_detected_levels()

        # 先清理旧目录块，再计算插入位置。
        # 历史版本先算索引后删旧块：索引失效，且上次生成的折页空段
        # 位于目录标题之前、不在清理范围内，每更新一次就累积一个空白页。
        self._remove_existing_generated_toc_block()

        insert_idx = self._find_toc_insert_index()
        heading_idx = self._find_toc_heading_index()
        if heading_idx is not None:
            # 清理上次生成时残留在目录标题之前的折页空段。
            self._remove_stale_toc_break_paragraphs(heading_idx - 1, max_scan=2)

        toc_titles = self._build_toc_titles(levels=levels)
        if not toc_titles:
            return {"ok": False, "message": "未识别到可用于生成目录的标题（H1/H2/H3）。"}

        # Ensure TOC section starts on a dedicated page.
        before_break = self._insert_paragraph_by_index(insert_idx)
        before_break.add_run().add_break(WD_BREAK.PAGE)
        insert_idx += 1

        if heading_idx is None:
            heading_para = self._insert_paragraph_by_index(insert_idx)
            heading_para.text = "目录"
            self._apply_toc_heading_style(heading_para)
            insert_idx += 1
        else:
            # If there is already a TOC heading in doc, keep it and insert TOC right after
            # the first heading occurrence.
            existing_idx = self._find_toc_heading_index()
            if existing_idx is not None:
                insert_idx = existing_idx + 1

        inserted_count = self._insert_toc_entries(insert_idx, toc_titles)

        # Ensure TOC section ends and正文 starts on next page.
        anchor_idx = min(len(doc.paragraphs) - 1, insert_idx + max(0, inserted_count - 1))
        anchor_after_toc = doc.paragraphs[anchor_idx]
        after_break = self._insert_paragraph_after(anchor_after_toc)
        after_break.add_run().add_break(WD_BREAK.PAGE)
        self._paragraphs = self._collect_paragraphs(doc.paragraphs)
        return {"ok": True, "message": f"已生成完整目录（{inserted_count}条，1-{levels}级）"}

    def update_toc(self, page_only: bool = True, levels: int = 3, mode: str = "smart") -> dict[str, object]:
        # Manual TOC mode: both update actions regenerate visible entries.
        result = self.auto_generate_toc(levels=levels, mode=mode)
        if not bool(result.get("ok")):
            return result
        if page_only:
            return {"ok": True, "message": "目录已更新（仅页码模式已应用）。"}
        return {"ok": True, "message": "目录已更新（整个目录模式已应用）。"}

    def check_toc_page_consistency(self) -> dict[str, object]:
        doc = self._require_document()
        heading_idx = self._find_toc_heading_index()
        if heading_idx is None:
            return {"ok": False, "message": "未找到目录章节。", "mismatch_count": 0, "items": []}

        toc_entries = self._extract_toc_entries(doc, heading_idx)
        if not toc_entries:
            return {"ok": False, "message": "目录条目为空，请先在 Word 中更新目录。", "mismatch_count": 0, "items": []}

        paragraphs = self.get_paragraphs()
        body_start_idx = self._get_structure_start_index("正文")
        body_start_page = self._get_body_start_page_estimate(body_start_idx)
        heading_pages: dict[str, int] = {}
        for para in paragraphs:
            idx = int(para.get("index", -1))
            if body_start_idx is not None and idx < body_start_idx:
                continue
            level = int(para.get("detected_level", 0))
            text = str(para.get("text", ""))
            is_post_part_title = text.startswith("参考文献") or text.startswith("致谢") or text.startswith("附录")
            if level not in {self.LEVEL_H1, self.LEVEL_H2, self.LEVEL_H3} and not is_post_part_title:
                continue
            key = self._normalize_compare_text(text)
            if key and key not in heading_pages:
                page_est = int(para.get("page_est", 1))
                heading_pages[key] = max(1, page_est - body_start_page + 1)

        # Page estimate may drift by one page around boundary paragraphs.
        # Treat +/-1 as acceptable to avoid false mismatch alarms.
        allowed_delta = 1
        mismatches: list[dict[str, object]] = []
        for item in toc_entries:
            key = self._normalize_compare_text(item["title"])
            actual = heading_pages.get(key)
            if actual is None:
                continue
            if abs(actual - item["page"]) > allowed_delta:
                mismatches.append({"title": item["title"], "toc_page": item["page"], "actual_page": actual})

        if not mismatches:
            return {"ok": True, "message": "目录页码与正文估算页码一致（容差±1页）。", "mismatch_count": 0, "items": []}
        return {
            "ok": False,
            "message": f"发现 {len(mismatches)} 项目录页码不一致。",
            "mismatch_count": len(mismatches),
            "items": mismatches[:20],
        }

    def unlink_headers_footers_all_sections(self) -> int:
        doc = self._require_document()
        self._push_undo_snapshot()
        changed = 0
        for section in doc.sections:
            if section.header.is_linked_to_previous:
                section.header.is_linked_to_previous = False
                changed += 1
            if section.footer.is_linked_to_previous:
                section.footer.is_linked_to_previous = False
                changed += 1
        return changed

    def switch_page_number_format_bulk(self, mode: str) -> int:
        doc = self._require_document()
        self._push_undo_snapshot()
        mode_normalized = mode.strip().lower()
        if mode_normalized not in {"roman", "arabic"}:
            return 0
        fmt = "upperRoman" if mode_normalized == "roman" else "decimal"
        field_code = "PAGE \\* ROMAN" if mode_normalized == "roman" else "PAGE"
        changed = 0
        for idx, section in enumerate(doc.sections):
            if not self._section_has_page_number(section):
                continue
            section.footer.is_linked_to_previous = False
            para = section.footer.paragraphs[0] if section.footer.paragraphs else section.footer.add_paragraph()
            self._clear_paragraph_content(para)
            para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run = para.add_run()
            run.font.name = "Times New Roman"
            run.font.size = Pt(10.5)
            if run._element.rPr is not None and run._element.rPr.rFonts is not None:
                rfonts = run._element.rPr.rFonts
            else:
                rpr = run._element.get_or_add_rPr()
                rfonts = rpr.get_or_add_rFonts()
            rfonts.set(qn("w:ascii"), "Times New Roman")
            rfonts.set(qn("w:hAnsi"), "Times New Roman")
            self._append_field(run, field_code)
            self._set_section_page_number_type(section, fmt=fmt, start=1 if idx == 0 else None)
            changed += 1
        if changed:
            self._paragraphs = self._collect_paragraphs(doc.paragraphs)
        return changed

    def one_click_section_and_page_number(self) -> dict[str, int]:
        doc = self._require_document()
        self._push_undo_snapshot()
        starts = self._detect_structure_starts()
        labels = [label for label, _ in starts]
        section_breaks = self._apply_structure_section_breaks(starts)
        unlinked = self.unlink_headers_footers_all_sections()
        numbered = self._apply_page_numbers_by_structure(labels)
        self._paragraphs = self._collect_paragraphs(doc.paragraphs)
        return {
            "structures": len(starts),
            "section_breaks": section_breaks,
            "unlinked": unlinked,
            "numbered_sections": numbered,
        }

    def apply_reference_hanging_indent(self, chars: float = 2.0, base_font_pt: float = 12.0) -> int:
        doc = self._require_document()
        bounds = self._get_reference_section_bounds()
        if bounds is None:
            return 0
        self._push_undo_snapshot()
        ref_indices = self._get_reference_body_indices()
        if not ref_indices:
            return 0
        indent_pt = max(1.0, chars * base_font_pt)
        changed = 0
        for idx in ref_indices:
            if idx < 0 or idx >= len(doc.paragraphs):
                continue
            para = doc.paragraphs[idx]
            fmt = para.paragraph_format
            old_left = self._to_points(fmt.left_indent)
            old_first = self._to_points(fmt.first_line_indent)
            old_line = self._safe_float(fmt.line_spacing) or 0.0
            old_before = self._to_points(fmt.space_before)
            if (
                abs(old_left - indent_pt) <= 0.2
                and abs(old_first + indent_pt) <= 0.2
                and abs(old_line - 1.5) <= 0.02
                and abs(old_before - 6.75) <= 0.3
            ):
                continue
            self._apply_reference_para_style(para, indent_pt=indent_pt)
            changed += 1
        if changed:
            self._paragraphs = self._collect_paragraphs(doc.paragraphs)
        return changed

    # GB/T 7714 文献类型标识推断：仅当原文缺少 [X] 标识时使用。
    # 无法可靠判断时返回 None（保持原样），绝不默认补 [J]——
    # 历史版本对缺标识文献一律补 [J]，把专著/学位论文/会议论文都标成期刊。
    _REFERENCE_TYPE_PATTERNS: tuple[tuple[str, re.Pattern[str]], ...] = (
        ("D", re.compile(r"博士|硕士|学位论文|dissertation|thesis", re.I)),
        ("C", re.compile(r"会议|论文集|研讨会|conference|proceedings|symposium|congress", re.I)),
        ("N", re.compile(r"报纸|日报|晚报|newspaper", re.I)),
        ("R", re.compile(r"技术报告|研究报告|报告|\breport\b", re.I)),
        ("S", re.compile(r"标准|\bGB/T|ISO\b|\bstandard\b", re.I)),
        ("P", re.compile(r"专利|patent", re.I)),
        ("DB/OL", re.compile(r"数据库|database", re.I)),
        ("EB/OL", re.compile(r"电子|网页|网站|EB/OL|https?://", re.I)),
        ("M", re.compile(r"出版社|出版|图书|专著|\bbooks?\b|\bpress\b", re.I)),
        ("J", re.compile(r"期刊|学报|杂志|journal", re.I)),
    )

    def _infer_reference_type(self, text: str) -> str | None:
        """根据文献内容推断 GB/T 7714 类型标识，无法判断时返回 None。"""
        for marker, pattern in self._REFERENCE_TYPE_PATTERNS:
            if pattern.search(text):
                return marker
        return None

    def format_references_gbt(self, mode: str = "order") -> dict[str, object]:
        doc = self._require_document()
        self._push_undo_snapshot()
        mode_norm = mode.strip().lower()
        if mode_norm not in {"order", "author_year"}:
            mode_norm = "order"
        bounds = self._get_reference_section_bounds()
        if bounds is None:
            return {"ok": False, "message": "未找到参考文献章节或文献条目。", "changed": 0}
        entries = self._extract_reference_entries(bounds)
        if not entries:
            return {"ok": False, "message": "未识别到可处理的参考文献条目（需以[数字]开头）。", "changed": 0}

        formatted: list[str] = []
        seq_no = 1
        for entry in entries:
            cleaned = self._normalize_reference_text(entry["raw"])
            cleaned = self._strip_leading_reference_numbering(cleaned)
            cleaned = self._normalize_reference_punctuation(cleaned, bool(entry["is_english"]))
            cleaned = self._normalize_reference_authors(cleaned, bool(entry["is_english"]))
            # 先规整句末标点，再把类型标识插到句末句号之前，
            # 避免出现“句子已结束又在句末后追加 [J].”的污染形态。
            if not cleaned.endswith("."):
                cleaned = cleaned.rstrip("。.;；：:，,") + "."
            if not re.search(r"\[[A-Z]+(?:/[A-Z]+)?\]", cleaned):
                # 仅在有把握时补类型标识；判断不出就保留原样，避免误注入 [J]。
                inferred = self._infer_reference_type(cleaned)
                if inferred:
                    cleaned = f"{cleaned[:-1]} [{inferred}]."
            if mode_norm == "order":
                cleaned = f"[{seq_no}] {cleaned}"
                seq_no += 1
            formatted.append(cleaned)

        changed = self._rewrite_reference_section(bounds, formatted)
        self._paragraphs = self._collect_paragraphs(doc.paragraphs)
        label = "顺序编码制" if mode_norm == "order" else "著者-出版年制"
        return {"ok": True, "message": f"已按 {label} 格式化参考文献。", "changed": changed}

    def validate_references_gbt(self, mode: str = "order") -> dict[str, object]:
        self._require_document()
        mode_norm = mode.strip().lower()
        if mode_norm not in {"order", "author_year"}:
            mode_norm = "order"
        bounds = self._get_reference_section_bounds()
        if bounds is None:
            return {"ok": False, "message": "未找到参考文献章节或文献条目。", "error_count": 0, "items": []}
        entries = self._extract_reference_entries(bounds)
        if not entries:
            return {"ok": False, "message": "未识别到可校验的参考文献条目。", "error_count": 0, "items": []}

        issues: list[dict[str, object]] = []
        if mode_norm == "order":
            numbers = [int(e["num"]) for e in entries]
            for i, n in enumerate(numbers, start=1):
                if n != i:
                    issues.append({"row": i, "issue": f"编号不连续：应为[{i}]，实际[{n}]", "text": entries[i - 1]["raw"][:60]})

        for row_no, entry in enumerate(entries, start=1):
            entry_issues = self._validate_single_reference_fields(entry)
            for it in entry_issues:
                issues.append({"row": row_no, "issue": it, "text": str(entry["raw"])[:60]})

        if not issues:
            return {"ok": True, "message": "参考文献校验通过。", "error_count": 0, "items": []}
        return {
            "ok": False,
            "message": f"发现 {len(issues)} 处参考文献格式问题。",
            "error_count": len(issues),
            "items": issues[:50],
        }

    def auto_fix_references_by_validation(self, mode: str = "order") -> dict[str, object]:
        fmt = self.format_references_gbt(mode=mode)
        if not bool(fmt.get("ok")):
            return fmt
        check = self.validate_references_gbt(mode=mode)
        unresolved = int(check.get("error_count", 0))
        if unresolved == 0:
            return {"ok": True, "message": "已按校验结果自动修复参考文献，未发现剩余问题。", "unresolved": 0}
        return {
            "ok": True,
            "message": f"已完成一键修改，仍有 {unresolved} 项需人工确认。",
            "unresolved": unresolved,
            "items": check.get("items", []),
        }

    def get_reference_entry_count(self) -> int:
        bounds = self._get_reference_section_bounds()
        if bounds is None:
            return 0
        return len(self._extract_reference_entries(bounds))

    def _get_reference_section_bounds(self) -> tuple[int, int] | None:
        doc = self._require_document()
        start_idx: int | None = None
        for i, para in enumerate(doc.paragraphs):
            t = (para.text or "").strip()
            if not t:
                continue
            n = re.sub(r"\s+", "", t).lower()
            n = re.sub(r"[:：]+$", "", n)
            if n in {"参考文献", "reference", "references"}:
                start_idx = i + 1
                break
        if start_idx is None:
            return None
        end_idx = len(doc.paragraphs)
        for j in range(start_idx, len(doc.paragraphs)):
            t = (doc.paragraphs[j].text or "").strip()
            if not t:
                continue
            n = re.sub(r"\s+", "", t).lower()
            n = re.sub(r"[:：]+$", "", n)
            if n in {"致谢", "acknowledgements", "acknowledgments"}:
                end_idx = j
                break
        return (start_idx, end_idx)

    def _get_reference_body_indices(self) -> list[int]:
        doc = self._require_document()
        bounds = self._get_reference_section_bounds()
        if bounds is None:
            return []
        start_idx, end_idx = bounds
        return [k for k in range(start_idx, end_idx) if (doc.paragraphs[k].text or "").strip()]

    def get_reference_section_indices(self) -> list[int]:
        """返回“参考文献”章节的段落索引（含标题，到“致谢”或文末为止）。"""
        doc = self._require_document()
        start_idx: int | None = None
        for i, para in enumerate(doc.paragraphs):
            if self._match_section_heading(
                para.text, {"参考文献", "reference", "references"}
            ):
                start_idx = i
                break
        if start_idx is None:
            return []
        end_idx = len(doc.paragraphs)
        for j in range(start_idx + 1, len(doc.paragraphs)):
            if self._match_section_heading(
                doc.paragraphs[j].text, {"致谢", "acknowledgements", "acknowledgments"}
            ):
                end_idx = j
                break
        return list(range(start_idx, end_idx))

    def get_acknowledgment_indices(self) -> list[int]:
        """返回“致谢”章节的段落索引（含标题，到文末）。"""
        doc = self._require_document()
        start_idx: int | None = None
        for i, para in enumerate(doc.paragraphs):
            if self._match_section_heading(
                para.text, {"致谢", "acknowledgements", "acknowledgments"}
            ):
                start_idx = i
                break
        if start_idx is None:
            return []
        return list(range(start_idx, len(doc.paragraphs)))

    def _match_section_heading(self, text: str | None, names: set[str]) -> bool:
        """判断段落文本是否匹配指定章节标题（忽略空白与末尾冒号）。"""
        t = (text or "").strip()
        if not t:
            return False
        n = re.sub(r"\s+", "", t).lower()
        n = re.sub(r"[:：]+$", "", n)
        return n in names

    def _extract_reference_entries(self, bounds: tuple[int, int]) -> list[dict[str, object]]:
        doc = self._require_document()
        start_idx, end_idx = bounds
        lines = [(doc.paragraphs[i].text or "").strip() for i in range(start_idx, end_idx) if (doc.paragraphs[i].text or "").strip()]
        if not lines:
            return []
        out: list[dict[str, object]] = []
        numeric_prefix = re.compile(r"^\[(\d+)\]\s*(.+)$")
        has_numbered = any(numeric_prefix.match(x) for x in lines)
        for i, raw in enumerate(lines, start=1):
            m = numeric_prefix.match(raw)
            if has_numbered and not m:
                if out:
                    out[-1]["raw"] = f"{out[-1]['raw']} {raw}".strip()
                    out[-1]["body"] = f"{out[-1]['body']} {raw}".strip()
                    parsed = self._parse_reference_fields(str(out[-1]["body"]))
                    out[-1]["author"] = parsed["author"]
                    out[-1]["title"] = parsed["title"]
                    out[-1]["year"] = parsed["year"]
                    out[-1]["type"] = parsed["type"]
                    out[-1]["is_english"] = parsed["is_english"]
                continue
            if m:
                num = int(m.group(1))
                body = m.group(2).strip()
            else:
                num = i
                body = raw.strip()
            parsed = self._parse_reference_fields(body)
            out.append(
                {
                    "num": num,
                    "raw": raw.strip(),
                    "body": body,
                    "author": parsed["author"],
                    "title": parsed["title"],
                    "year": parsed["year"],
                    "type": parsed["type"],
                    "is_english": parsed["is_english"],
                }
            )
        return out

    def _parse_reference_fields(self, body: str) -> dict[str, object]:
        type_match = re.search(r"\[([A-Z]+(?:/[A-Z]+)?)\]", body)
        year_match = re.search(r"(19|20)\d{2}", body)
        type_code = type_match.group(1) if type_match else ""
        year = year_match.group(0) if year_match else ""
        is_eng = self._is_english_reference(body)
        parts = re.split(r"[.。]", body, maxsplit=1)
        author = parts[0].strip() if parts else ""
        title = ""
        if len(parts) > 1:
            left = parts[1]
            if type_match:
                left = left[: type_match.start() - len(parts[0]) - 1] if type_match.start() > len(parts[0]) else left
            title = left.strip(" .。；;,:：")
        return {"author": author, "title": title, "year": year, "type": type_code, "is_english": is_eng}

    def _normalize_reference_text(self, text: str) -> str:
        s = text.replace("\t", " ").replace("\u3000", " ")
        s = re.sub(r"\s+", " ", s).strip()
        s = re.sub(r"\s*([,.;:，。；：])\s*", r"\1 ", s)
        s = re.sub(r"\s+\]", "]", s)
        s = re.sub(r"\s+\.", ".", s)
        return s.strip()

    def _is_english_reference(self, text: str) -> bool:
        cjk = len(self._CJK_PATTERN.findall(text))
        alpha = len(re.findall(r"[A-Za-z]", text))
        return alpha > cjk * 2 and alpha > 8

    def _normalize_reference_punctuation(self, text: str, is_english: bool) -> str:
        half_to_full = {",": "，", ".": "。", ";": "；", ":": "："}
        full_to_half = {"，": ",", "。": ".", "；": ";", "：": ":"}
        s = text
        if is_english:
            for k, v in full_to_half.items():
                s = s.replace(k, v)
        else:
            for k, v in half_to_full.items():
                s = s.replace(k, v)
        s = re.sub(r"\s+", " ", s).strip()
        return s

    def _normalize_reference_authors(self, text: str, is_english: bool) -> str:
        s = text
        if re.search(r"\[[A-Z]", s) and is_english:
            s = re.sub(r"\bet al\.\s*et al\.\b", "et al.", s, flags=re.I)
            head, sep, tail = s.partition(".")
            names = [x.strip() for x in re.split(r",| and ", head) if x.strip()]
            if len(names) > 3 and "et al." not in head.lower():
                head = ", ".join(names[:3]) + ", et al."
                s = head + (sep + tail if sep else "")
            return s
        if "，" in s and "等" not in s:
            head = s.split("。")[0]
            if head.count("，") >= 3:
                parts = head.split("，")
                keep = "，".join(parts[:3]) + "，等"
                s = s.replace(head, keep, 1)
        return s

    def _strip_leading_reference_numbering(self, text: str) -> str:
        # Remove repeated leading numbering markers like [1] / 【1】 / [ 1 ].
        return re.sub(r"^\s*(?:(?:[\[【]\s*\d+\s*[\]】])[\s、.．]*)+", "", text)

    def _remove_empty_paragraphs_in_indices(self, indices: list[int]) -> int:
        doc = self._require_document()
        removed = 0
        for idx in sorted(indices, reverse=True):
            if idx < 0 or idx >= len(doc.paragraphs):
                continue
            para = doc.paragraphs[idx]
            if (para.text or "").strip():
                continue
            p = para._element
            p.getparent().remove(p)
            removed += 1
        return removed

    def _validate_single_reference_fields(self, entry: dict[str, object]) -> list[str]:
        issues: list[str] = []
        text = str(entry.get("body", "")).strip()
        type_code = str(entry.get("type", "")).strip()
        author = str(entry.get("author", "")).strip()
        title = str(entry.get("title", "")).strip()
        year = str(entry.get("year", "")).strip()
        is_english = bool(entry.get("is_english"))

        if type_code not in {"J", "M", "D", "C", "EB/OL"}:
            issues.append("缺少文献类型标识（如 [J]/[M]/[D]）")
        if not author:
            issues.append("缺少作者字段")
        if not title:
            issues.append("缺少标题字段")
        if not year:
            issues.append("缺少年份")
        if re.search(r"[,，]\s*[,，]|[.。]{2,}|[:：]{2,}|[;；]{2,}", text):
            issues.append("标点可能不规范（存在重复或异常）")
        if not text.endswith("."):
            issues.append("结尾应统一使用英文句点 .")

        if type_code == "J":
            if not re.search(r"\d+\s*\(\d+\)", text):
                issues.append("期刊文献缺少卷(期)信息")
            if not re.search(r":\s*\d+(-\d+)?|：\s*\d+(-\d+)?", text):
                issues.append("期刊文献缺少页码信息")

        if is_english:
            if re.search(r"[A-Za-z]+\s+[A-Za-z]+", author) and not re.search(r"[A-Za-z]+,\s*[A-Z]", author):
                issues.append("英文作者格式可能不规范（建议 姓, 名缩写）")
            if author.count(",") >= 3 and "et al." not in author.lower():
                issues.append("英文作者超过3人建议使用 et al.")
        else:
            if author.count("，") >= 3 and "等" not in author:
                issues.append("中文作者超过3人建议加“等”")
        return issues

    def _rewrite_reference_section(self, bounds: tuple[int, int], entries_text: list[str]) -> int:
        doc = self._require_document()
        start_idx, end_idx = bounds
        changed = 0
        for idx in range(end_idx - 1, start_idx - 1, -1):
            if idx < 0 or idx >= len(doc.paragraphs):
                continue
            p = doc.paragraphs[idx]._element
            p.getparent().remove(p)
            changed += 1
        indent_pt = 24.0
        for i, text in enumerate(entries_text):
            para = self._insert_paragraph_by_index(start_idx + i)
            self._replace_paragraph_text(para, text)
            self._apply_reference_para_style(para, indent_pt=indent_pt)
            changed += 1
        return changed

    def _apply_reference_para_style(self, para: Paragraph, indent_pt: float = 24.0) -> None:
        fmt = para.paragraph_format
        fmt.left_indent = Pt(indent_pt)
        fmt.first_line_indent = Pt(-indent_pt)
        fmt.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        fmt.line_spacing = 1.5
        fmt.space_before = Pt(6.75)
        fmt.space_after = Pt(0)
        if not para.runs:
            para.add_run("")
        for run in para.runs:
            run.font.size = Pt(9)
            run.font.name = "Times New Roman"
            if run._element.rPr is not None and run._element.rPr.rFonts is not None:
                rfonts = run._element.rPr.rFonts
            else:
                rpr = run._element.get_or_add_rPr()
                rfonts = rpr.get_or_add_rFonts()
            rfonts.set(qn("w:eastAsia"), "宋体")
            rfonts.set(qn("w:ascii"), "Times New Roman")
            rfonts.set(qn("w:hAnsi"), "Times New Roman")

    def clean_delete_char(self, indices: set[int], target_char: str) -> int:
        if not target_char or not indices:
            return 0
        doc = self._require_document()
        self._push_undo_snapshot()
        changed = 0
        for idx in sorted(indices):
            if idx < 0 or idx >= len(doc.paragraphs):
                continue
            para = doc.paragraphs[idx]
            original = para.text or ""
            updated = original.replace(target_char, "")
            if updated != original:
                self._replace_paragraph_text(para, updated)
                changed += 1
        if changed:
            self._paragraphs = self._collect_paragraphs(doc.paragraphs)
        return changed

    def clean_keep_numeric_dash(self, indices: set[int]) -> int:
        if not indices:
            return 0
        doc = self._require_document()
        self._push_undo_snapshot()
        changed = 0
        for idx in sorted(indices):
            if idx < 0 or idx >= len(doc.paragraphs):
                continue
            para = doc.paragraphs[idx]
            original = para.text or ""
            chars = list(original)
            for pos, ch in enumerate(chars):
                if ch != "-":
                    continue
                prev_ch = chars[pos - 1] if pos - 1 >= 0 else ""
                next_ch = chars[pos + 1] if pos + 1 < len(chars) else ""
                if not (prev_ch.isdigit() and next_ch.isdigit()):
                    chars[pos] = ""
            updated = "".join(chars)
            if updated != original:
                self._replace_paragraph_text(para, updated)
                changed += 1
        if changed:
            self._paragraphs = self._collect_paragraphs(doc.paragraphs)
        return changed

    def clean_delete_placeholder_dot(self, indices: set[int]) -> int:
        """Delete placeholder dot that occupies a whole line only."""
        if not indices:
            return 0
        doc = self._require_document()
        self._push_undo_snapshot()
        changed = 0
        for idx in sorted(indices):
            if idx < 0 or idx >= len(doc.paragraphs):
                continue
            para = doc.paragraphs[idx]
            original = para.text or ""
            updated = "" if re.fullmatch(r"\s*\.\s*", original) else original
            if updated != original:
                self._replace_paragraph_text(para, updated)
                changed += 1
        if changed:
            self._paragraphs = self._collect_paragraphs(doc.paragraphs)
        return changed

    def clean_remove_empty_paragraphs(self) -> int:
        doc = self._require_document()
        self._push_undo_snapshot()
        removed = 0
        for para in list(doc.paragraphs):
            if (para.text or "").strip():
                continue
            p_element = para._element
            p_element.getparent().remove(p_element)
            removed += 1
        if removed:
            self._paragraphs = self._collect_paragraphs(doc.paragraphs)
        return removed

    def count_empty_paragraphs(self) -> int:
        doc = self._require_document()
        return sum(1 for para in doc.paragraphs if not (para.text or "").strip())

    def undo_last_action(self) -> bool:
        doc = self._require_document()
        if not self._undo_stack:
            return False
        current = BytesIO()
        doc.save(current)
        self._redo_stack.append(current.getvalue())
        if len(self._redo_stack) > 30:
            self._redo_stack = self._redo_stack[-30:]
        last_snapshot = self._undo_stack.pop()
        self._document = Document(BytesIO(last_snapshot))
        self._paragraphs = self._collect_paragraphs(self._document.paragraphs)
        return True

    def redo_last_action(self) -> bool:
        doc = self._require_document()
        if not self._redo_stack:
            return False
        current = BytesIO()
        doc.save(current)
        self._undo_stack.append(current.getvalue())
        if len(self._undo_stack) > 30:
            self._undo_stack = self._undo_stack[-30:]
        snap = self._redo_stack.pop()
        self._document = Document(BytesIO(snap))
        self._paragraphs = self._collect_paragraphs(self._document.paragraphs)
        return True

    def get_paragraphs(self) -> list[dict[str, Any]]:
        return [dict(item) for item in self._paragraphs]

    def get_indices_by_level(self, level: int | str) -> list[int]:
        parsed_level = self._normalize_level(level)
        return [item["index"] for item in self._paragraphs if item["detected_level"] == parsed_level]

    def get_range_under_title(self, cursor_idx: int) -> list[int]:
        if not self._paragraphs:
            return []

        pos_map = {p["index"]: i for i, p in enumerate(self._paragraphs)}
        cursor_pos = pos_map.get(cursor_idx)
        if cursor_pos is None:
            return []

        heading_pos = self._resolve_heading_anchor(cursor_pos)
        if heading_pos is None:
            return [cursor_idx]

        heading_level = self._paragraphs[heading_pos]["detected_level"]
        heading_rank = self._heading_rank(heading_level)
        if heading_rank is None:
            return [cursor_idx]

        end_pos = len(self._paragraphs) - 1
        for pos in range(heading_pos + 1, len(self._paragraphs)):
            level = self._paragraphs[pos]["detected_level"]
            rank = self._heading_rank(level)
            if rank is not None and rank <= heading_rank:
                end_pos = pos - 1
                break

        return [self._paragraphs[i]["index"] for i in range(heading_pos, end_pos + 1)]

    def _collect_paragraphs(self, paragraphs: list[Paragraph]) -> list[dict[str, Any]]:
        result: list[dict[str, Any]] = []
        content_width_pt, content_height_pt = self._get_page_content_size_pt()
        bottom_safe_pt = self._get_dynamic_bottom_safe_pt(paragraphs)
        current_page = 1
        used_height_pt = 0.0

        for idx, para in enumerate(paragraphs):
            text = para.text or ""
            font_size_pt = self._extract_font_size_pt(para)
            font_name = self._extract_font_name(para)
            bold, italic, underline = self._extract_font_style(para)
            alignment = self._extract_alignment(para)
            line_spacing = self._extract_line_spacing(para)
            first_line_indent_pt = self._to_points(para.paragraph_format.first_line_indent)
            left_indent_pt = self._to_points(para.paragraph_format.left_indent)

            detected_level = self._detect_level_by_numbering(text)
            style_level = self._detect_level_by_style(para, font_size_pt, bold, italic)
            if detected_level == self.LEVEL_BODY and style_level != self.LEVEL_BODY:
                detected_level = style_level

            if self._has_page_break_before(para):
                current_page += 1
                used_height_pt = 0.0

            para_height_pt = self._estimate_paragraph_height_pt(
                para=para,
                text=text,
                font_size_pt=font_size_pt,
                line_spacing=line_spacing,
                content_width_pt=content_width_pt,
            )
            available_height = max(60.0, content_height_pt - bottom_safe_pt)
            if used_height_pt > 0 and used_height_pt + para_height_pt > available_height:
                current_page += 1
                used_height_pt = 0.0

            page_est = current_page

            result.append(
                {
                    "index": idx,
                    "text": text,
                    "font_size_pt": font_size_pt,
                    "font_name": font_name,
                    "bold": bold,
                    "italic": italic,
                    "underline": underline,
                    "alignment": alignment,
                    "line_spacing": line_spacing,
                    "first_line_indent_pt": first_line_indent_pt,
                    "left_indent_pt": left_indent_pt,
                    "page_est": page_est,
                    "detected_level": detected_level,
                }
            )

            remaining_height = para_height_pt
            while remaining_height > 0:
                page_left = max(1.0, available_height - used_height_pt)
                if remaining_height <= page_left:
                    used_height_pt += remaining_height
                    remaining_height = 0
                else:
                    remaining_height -= page_left
                    current_page += 1
                    used_height_pt = 0.0

            page_break_count = self._count_manual_page_breaks(para)
            if page_break_count > 0:
                current_page += page_break_count
                used_height_pt = 0.0

        self._mark_big_title(result)
        return result

    def _get_dynamic_bottom_safe_pt(self, paragraphs: list[Paragraph]) -> float:
        font_sizes: list[float] = []
        for para in paragraphs:
            if not (para.text or "").strip():
                continue
            size = self._extract_font_size_pt(para)
            if size is not None:
                font_sizes.append(float(size))
        if not font_sizes:
            main_size = 12.0
        else:
            font_sizes.sort()
            main_size = font_sizes[len(font_sizes) // 2]
        # Use a smaller reserve at page bottom; previous value was too conservative
        # and caused preview to overflow into extra pages compared with Word.
        dynamic_safe = main_size * 0.35
        return max(self.PAGE_BOTTOM_SAFE_MIN_PT, min(self.PAGE_BOTTOM_SAFE_MAX_PT, dynamic_safe))

    def _get_page_content_size_pt(self) -> tuple[float, float]:
        doc = self._require_document()
        if not doc.sections:
            return 451.0, 648.0  # A4 1-inch margins fallback

        section = doc.sections[0]
        page_width_pt = float(section.page_width.pt)
        page_height_pt = float(section.page_height.pt)
        margin_left_pt = float(section.left_margin.pt)
        margin_right_pt = float(section.right_margin.pt)
        margin_top_pt = float(section.top_margin.pt)
        margin_bottom_pt = float(section.bottom_margin.pt)

        content_width_pt = max(220.0, page_width_pt - margin_left_pt - margin_right_pt)
        content_height_pt = max(260.0, page_height_pt - margin_top_pt - margin_bottom_pt)

        return content_width_pt, content_height_pt

    def _estimate_paragraph_height_pt(
        self,
        para: Paragraph,
        text: str,
        font_size_pt: float | None,
        line_spacing: float | None,
        content_width_pt: float,
    ) -> float:
        size_pt = font_size_pt if font_size_pt is not None else 12.0

        if line_spacing is None:
            line_height_pt = size_pt * 1.0
        else:
            line_height_pt = size_pt * float(line_spacing) * 1.0 if line_spacing <= 10 else float(line_spacing)
            line_height_pt = max(line_height_pt, size_pt * 1.0)

        normalized_text = text.replace("\r", "").replace("\n", "")
        if not normalized_text.strip():
            line_count = 1
        else:
            cjk_count = len(self._CJK_PATTERN.findall(normalized_text))
            other_count = max(0, len(normalized_text) - cjk_count)
            char_width_units = cjk_count * 1.0 + other_count * 0.55
            # Relax width coefficient to reduce over-wrapping against Word preview.
            units_per_line = max(8.0, content_width_pt / max(6.0, size_pt * 0.70))
            line_count = max(1, int(math.ceil(char_width_units / units_per_line)))

        before_pt = self._to_points(para.paragraph_format.space_before)
        after_pt = self._to_points(para.paragraph_format.space_after)
        return line_count * line_height_pt + before_pt + after_pt

    def _to_points(self, value: Any) -> float:
        if value is None:
            return 0.0
        if hasattr(value, "pt"):
            return float(value.pt)
        try:
            return float(value)
        except (TypeError, ValueError):
            return 0.0

    def _has_page_break_before(self, para: Paragraph) -> bool:
        return bool(para.paragraph_format.page_break_before)

    def _count_manual_page_breaks(self, para: Paragraph) -> int:
        xml_text = para._p.xml
        return xml_text.count('w:type="page"')

    def _safe_float(self, value: Any) -> float | None:
        if value is None:
            return None
        try:
            return float(value)
        except (TypeError, ValueError):
            if hasattr(value, "pt"):
                pt = float(value.pt)
                return pt / 12.0
            return None

    def _set_snap_to_grid_false(self, p_element: Any) -> bool:
        p_pr = p_element.get_or_add_pPr()
        snap = p_pr.find(qn("w:snapToGrid"))
        if snap is None:
            snap = OxmlElement("w:snapToGrid")
            p_pr.append(snap)
            snap.set(qn("w:val"), "0")
            return True
        old = snap.get(qn("w:val"))
        snap.set(qn("w:val"), "0")
        return old != "0"

    def _apply_xml_paragraph_rules_for_textboxes(self, doc: DocxDocument) -> int:
        changed = 0
        try:
            elements = doc._element.xpath(".//w:txbxContent//w:p")
        except Exception:
            return 0
        for p in elements:
            p_pr = p.find(qn("w:pPr"))
            if p_pr is None:
                p_pr = OxmlElement("w:pPr")
                p.insert(0, p_pr)
                changed += 1
            changed += self._ensure_xml_on_off(p_pr, "w:keepLines", "0")
            changed += self._ensure_xml_on_off(p_pr, "w:widowControl", "0")
            changed += self._ensure_xml_on_off(p_pr, "w:keepNext", "0")
            changed += self._ensure_xml_line_spacing_multiple(p_pr, "300")
            changed += self._ensure_xml_snap_to_grid_off(p_pr)
        return changed

    def _ensure_xml_on_off(self, p_pr: Any, tag: str, val: str = "1") -> int:
        node = p_pr.find(qn(tag))
        if node is None:
            node = OxmlElement(tag)
            node.set(qn("w:val"), val)
            p_pr.append(node)
            return 1
        old = node.get(qn("w:val"))
        node.set(qn("w:val"), val)
        return 1 if old != val else 0

    def _ensure_xml_line_spacing_multiple(self, p_pr: Any, line_val: str) -> int:
        spacing = p_pr.find(qn("w:spacing"))
        if spacing is None:
            spacing = OxmlElement("w:spacing")
            p_pr.append(spacing)
            spacing.set(qn("w:lineRule"), "auto")
            spacing.set(qn("w:line"), line_val)
            return 1
        old_rule = spacing.get(qn("w:lineRule"))
        old_line = spacing.get(qn("w:line"))
        spacing.set(qn("w:lineRule"), "auto")
        spacing.set(qn("w:line"), line_val)
        return 1 if (old_rule != "auto" or old_line != line_val) else 0

    def _ensure_xml_snap_to_grid_off(self, p_pr: Any) -> int:
        snap = p_pr.find(qn("w:snapToGrid"))
        if snap is None:
            snap = OxmlElement("w:snapToGrid")
            p_pr.append(snap)
            snap.set(qn("w:val"), "0")
            return 1
        old = snap.get(qn("w:val"))
        snap.set(qn("w:val"), "0")
        return 1 if old != "0" else 0

    def _disable_document_grid(self, doc: DocxDocument) -> int:
        changed = 0
        for section in doc.sections:
            sect_pr = section._sectPr
            grids = sect_pr.findall(qn("w:docGrid"))
            for grid in grids:
                sect_pr.remove(grid)
                changed += 1
        return changed

    def _disable_table_row_breaks(self, tables: list[Any]) -> int:
        changed = 0
        for table in tables:
            for row in table.rows:
                tr_pr = row._tr.get_or_add_trPr()
                cant_split = tr_pr.find(qn("w:cantSplit"))
                if cant_split is None:
                    cant_split = OxmlElement("w:cantSplit")
                    tr_pr.append(cant_split)
                    changed += 1
                old = cant_split.get(qn("w:val"))
                cant_split.set(qn("w:val"), "1")
                if old not in (None, "1"):
                    changed += 1
                for cell in row.cells:
                    changed += self._disable_table_row_breaks(cell.tables)
        return changed

    def _append_field(self, run: Any, field_name: str) -> None:
        begin = OxmlElement("w:fldChar")
        begin.set(qn("w:fldCharType"), "begin")
        instr = OxmlElement("w:instrText")
        instr.set(qn("xml:space"), "preserve")
        instr.text = f" {field_name} "
        end = OxmlElement("w:fldChar")
        end.set(qn("w:fldCharType"), "end")
        run._r.append(begin)
        run._r.append(instr)
        run._r.append(end)

    def _append_field_with_result(self, run: Any, field_name: str, result_text: str) -> None:
        begin = OxmlElement("w:fldChar")
        begin.set(qn("w:fldCharType"), "begin")
        instr = OxmlElement("w:instrText")
        instr.set(qn("xml:space"), "preserve")
        instr.text = f" {field_name} "
        separate = OxmlElement("w:fldChar")
        separate.set(qn("w:fldCharType"), "separate")
        text = OxmlElement("w:t")
        text.text = result_text
        end = OxmlElement("w:fldChar")
        end.set(qn("w:fldCharType"), "end")
        run._r.append(begin)
        run._r.append(instr)
        run._r.append(separate)
        run._r.append(text)
        run._r.append(end)

    def _next_bookmark_id(self) -> int:
        doc = self._require_document()
        max_id = 0
        for para in doc.paragraphs:
            for node in para._p.xpath(".//w:bookmarkStart"):
                try:
                    max_id = max(max_id, int(node.get(qn("w:id")) or 0))
                except Exception:
                    continue
        return max_id + 1

    def _ensure_toc_heading_bookmark(self, para_index: int) -> str:
        doc = self._require_document()
        if para_index < 0 or para_index >= len(doc.paragraphs):
            return ""
        para = doc.paragraphs[para_index]
        for node in para._p.xpath(".//w:bookmarkStart"):
            name = str(node.get(qn("w:name")) or "")
            if name.startswith("ADFA_TOC_"):
                return name
        bookmark_id = self._next_bookmark_id()
        bookmark_name = f"ADFA_TOC_{para_index}_{bookmark_id}"
        start = OxmlElement("w:bookmarkStart")
        start.set(qn("w:id"), str(bookmark_id))
        start.set(qn("w:name"), bookmark_name)
        end = OxmlElement("w:bookmarkEnd")
        end.set(qn("w:id"), str(bookmark_id))
        para._p.insert(0, start)
        para._p.append(end)
        return bookmark_name

    def _set_section_page_start(self, section: Any, start_page: int) -> None:
        sect_pr = section._sectPr
        node = sect_pr.find(qn("w:pgNumType"))
        if node is None:
            node = OxmlElement("w:pgNumType")
            sect_pr.append(node)
        node.set(qn("w:start"), str(max(1, int(start_page))))

    def _clear_section_page_start(self, section: Any) -> None:
        sect_pr = section._sectPr
        node = sect_pr.find(qn("w:pgNumType"))
        if node is not None:
            sect_pr.remove(node)

    def _clear_paragraph_content(self, para: Paragraph) -> None:
        p = para._p
        for child in list(p):
            if child.tag != qn("w:pPr"):
                p.remove(child)

    def _refresh_paragraph_cache(self) -> None:
        doc = self._require_document()
        self._paragraphs = self._collect_paragraphs(doc.paragraphs)

    def _sync_heading_styles_from_detected_levels(self) -> int:
        doc = self._require_document()
        self._refresh_paragraph_cache()
        changed = 0
        for item in self._paragraphs:
            idx = int(item.get("index", -1))
            level = int(item.get("detected_level", self.LEVEL_BODY))
            if idx < 0 or idx >= len(doc.paragraphs):
                continue
            if level not in {self.LEVEL_H1, self.LEVEL_H2, self.LEVEL_H3}:
                continue
            para = doc.paragraphs[idx]
            target_name = f"Heading {level}"
            try:
                current_name = para.style.name if para.style is not None else ""
                if current_name != target_name:
                    para.style = doc.styles[target_name]
                    changed += 1
            except Exception:
                # Fallback for Chinese UI naming.
                zh_name = f"标题 {level}"
                try:
                    current_name = para.style.name if para.style is not None else ""
                    if current_name != zh_name:
                        para.style = doc.styles[zh_name]
                        changed += 1
                except Exception:
                    continue
        return changed

    def _find_toc_heading_index(self) -> int | None:
        doc = self._require_document()
        for idx, para in enumerate(doc.paragraphs):
            text = (para.text or "").strip()
            if not text:
                continue
            normalized = re.sub(r"\s+", "", text)
            if normalized == "目录" or normalized.startswith("目录"):
                return idx
        return None

    def _find_toc_insert_index(self) -> int:
        doc = self._require_document()
        first_body_idx: int | None = None
        last_front_idx: int | None = None
        for idx, para in enumerate(doc.paragraphs):
            text = (para.text or "").strip()
            if not text:
                continue
            normalized = re.sub(r"\s+", "", text).lower()
            if (
                normalized.startswith("摘要")
                or normalized.startswith("abstract")
                or normalized.startswith("关键词")
                or normalized.startswith("keywords")
            ):
                last_front_idx = idx
            level = self._detect_level_by_numbering(text)
            if first_body_idx is None and (level == self.LEVEL_H1 or text.startswith("绪论")):
                first_body_idx = idx
        if first_body_idx is None:
            return (last_front_idx + 1) if last_front_idx is not None else len(doc.paragraphs)
        if last_front_idx is None:
            return first_body_idx
        return min(first_body_idx, last_front_idx + 1)

    def _remove_all_existing_toc_fields(self) -> None:
        doc = self._require_document()
        to_remove: list[Paragraph] = []
        for para in doc.paragraphs:
            if self._is_toc_field_paragraph(para):
                to_remove.append(para)
        for para in to_remove:
            parent = para._element.getparent()
            if parent is not None:
                parent.remove(para._element)

    def _remove_existing_generated_toc_block(self) -> None:
        doc = self._require_document()
        heading_idx = self._find_toc_heading_index()
        if heading_idx is None:
            self._remove_all_existing_toc_fields()
            return
        start = heading_idx + 1
        end = min(len(doc.paragraphs), start + 300)
        remove_targets: list[Paragraph] = []
        for idx in range(start, end):
            para = doc.paragraphs[idx]
            text = (para.text or "").strip()
            if not text:
                remove_targets.append(para)
                continue
            if self._is_toc_field_paragraph(para):
                remove_targets.append(para)
                continue
            if re.match(r"^.+\s*[\.\t ]+\d+\s*$", text):
                remove_targets.append(para)
                continue
            if self._detect_level_by_numbering(text) == self.LEVEL_H1 or text.startswith("绪论"):
                break
            if text == "目录":
                continue
            if idx > start + 3:
                break
        for para in remove_targets:
            parent = para._element.getparent()
            if parent is not None:
                parent.remove(para._element)

    def _remove_stale_toc_break_paragraphs(self, start_idx: int, max_scan: int = 2) -> int:
        """移除上次生成目录时遗留的折页空段（无文字、仅含分页符的段落）。"""
        doc = self._require_document()
        removed = 0
        idx = start_idx
        while idx >= 0 and removed < max_scan and idx < len(doc.paragraphs):
            para = doc.paragraphs[idx]
            if (para.text or "").strip() or not self._paragraph_has_page_break(para):
                break
            parent = para._element.getparent()
            if parent is not None:
                parent.remove(para._element)
            removed += 1
            idx -= 1
        return removed

    def _paragraph_has_page_break(self, para: Paragraph) -> bool:
        for node in para._p.xpath(".//w:br"):
            if node.get(qn("w:type")) == "page":
                return True
        return False

    def _insert_paragraph_by_index(self, index: int) -> Paragraph:
        doc = self._require_document()
        if not doc.paragraphs:
            return doc.add_paragraph()
        if index <= 0:
            first = doc.paragraphs[0]
            new_p = OxmlElement("w:p")
            first._p.addprevious(new_p)
            return Paragraph(new_p, first._parent)
        if index >= len(doc.paragraphs):
            return doc.add_paragraph()
        current = doc.paragraphs[index]
        new_p = OxmlElement("w:p")
        current._p.addprevious(new_p)
        return Paragraph(new_p, current._parent)

    def _apply_toc_heading_style(self, para: Paragraph) -> None:
        doc = self._require_document()
        para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        try:
            para.style = doc.styles["Heading 2"]
        except Exception:
            try:
                para.style = doc.styles["标题 2"]
            except Exception:
                pass

    def _build_toc_titles(self, levels: int = 3) -> list[dict[str, object]]:
        self._refresh_paragraph_cache()
        rows: list[dict[str, object]] = []
        body_start_idx = self._get_structure_start_index("正文")
        body_start_page = self._get_body_start_page_estimate(body_start_idx)
        h1_pattern = re.compile(r"^\s*第\s*([0-9]+|[一二三四五六七八九十百千万两零〇]+)\s*章")
        hx_pattern = re.compile(r"^\s*\d+\.\d+(?:\.\d+)?")
        post_part_titles = {"参考文献", "致谢", "附录", "后记"}
        for item in self._paragraphs:
            level = int(item.get("detected_level", 0))
            text = str(item.get("text", "")).strip()
            if not text:
                continue
            idx = int(item.get("index", -1))
            if body_start_idx is not None and idx < body_start_idx:
                continue
            is_post_part_title = any(text.startswith(x) for x in post_part_titles)
            if level not in {self.LEVEL_H1, self.LEVEL_H2, self.LEVEL_H3} and not is_post_part_title:
                continue
            if level > levels and not is_post_part_title:
                continue
            # Additional TOC filter rules:
            # 1) must start with "第X章" or "X.X(.X)"
            # 2) length <= 50
            # 3) must not end with "。，；："
            if len(text) > 50:
                continue
            if text[-1] in {"。", "，", "；", "："}:
                continue
            if not (h1_pattern.match(text) or hx_pattern.match(text) or is_post_part_title):
                continue
            page_est = int(item.get("page_est", 1))
            page_arabic = max(1, page_est - body_start_page + 1)
            bookmark = self._ensure_toc_heading_bookmark(idx)
            rows.append(
                {
                    "level": self.LEVEL_H1 if is_post_part_title else level,
                    "text": text,
                    "page": page_arabic,
                    "bookmark": bookmark,
                }
            )
        return [r for r in rows if r["text"]]

    def _insert_toc_entries(self, insert_idx: int, titles: list[dict[str, object]]) -> int:
        doc = self._require_document()
        right_tab = self._get_toc_right_tab_position_pt()
        inserted = 0
        for row in titles:
            para = self._insert_paragraph_by_index(insert_idx + inserted)
            level = int(row["level"])
            text = str(row["text"])
            page = int(row["page"])
            bookmark = str(row.get("bookmark", "") or "")
            para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
            para.paragraph_format.line_spacing = 1.5
            para.paragraph_format.left_indent = Pt(max(0, (level - 1) * 12))
            para.paragraph_format.tab_stops.add_tab_stop(Pt(right_tab), alignment=WD_TAB_ALIGNMENT.RIGHT, leader=WD_TAB_LEADER.DOTS)
            run = para.add_run(f"{text}\t")
            run.font.size = Pt(12)
            run.font.name = "宋体"
            if run._element.rPr is not None and run._element.rPr.rFonts is not None:
                rfonts = run._element.rPr.rFonts
            else:
                rpr = run._element.get_or_add_rPr()
                rfonts = rpr.get_or_add_rFonts()
            rfonts.set(qn("w:eastAsia"), "宋体")
            rfonts.set(qn("w:ascii"), "Times New Roman")
            rfonts.set(qn("w:hAnsi"), "Times New Roman")
            page_run = para.add_run("")
            page_run.font.size = Pt(12)
            page_run.font.name = "Times New Roman"
            if page_run._element.rPr is not None and page_run._element.rPr.rFonts is not None:
                page_rfonts = page_run._element.rPr.rFonts
            else:
                page_rpr = page_run._element.get_or_add_rPr()
                page_rfonts = page_rpr.get_or_add_rFonts()
            page_rfonts.set(qn("w:ascii"), "Times New Roman")
            page_rfonts.set(qn("w:hAnsi"), "Times New Roman")
            if bookmark:
                self._append_field_with_result(page_run, f"PAGEREF {bookmark} \\h", str(page))
            else:
                page_run.text = str(page)
            if level == self.LEVEL_H1:
                run.bold = True
            inserted += 1
        return inserted

    def _get_toc_right_tab_position_pt(self) -> float:
        doc = self._require_document()
        if not doc.sections:
            return 430.0
        section = doc.sections[0]
        width = float(section.page_width.pt) - float(section.left_margin.pt) - float(section.right_margin.pt)
        return max(300.0, width - 6.0)

    def _remove_existing_toc_block_after(self, heading_idx: int) -> None:
        doc = self._require_document()
        if heading_idx + 1 >= len(doc.paragraphs):
            return
        para = doc.paragraphs[heading_idx + 1]
        if self._is_toc_field_paragraph(para):
            parent = para._element.getparent()
            parent.remove(para._element)

    def _insert_paragraph_after(self, para: Paragraph) -> Paragraph:
        new_p = OxmlElement("w:p")
        para._p.addnext(new_p)
        return Paragraph(new_p, para._parent)

    def _set_toc_field(self, para: Paragraph, levels: int = 3) -> None:
        self._clear_paragraph_content(para)
        run = para.add_run("")
        begin = OxmlElement("w:fldChar")
        begin.set(qn("w:fldCharType"), "begin")
        instr = OxmlElement("w:instrText")
        instr.set(qn("xml:space"), "preserve")
        instr.text = f' TOC \\o "1-{levels}" \\h \\z \\u '
        separate = OxmlElement("w:fldChar")
        separate.set(qn("w:fldCharType"), "separate")
        text = OxmlElement("w:t")
        text.text = "（目录域，打开 Word 后右键更新目录）"
        end = OxmlElement("w:fldChar")
        end.set(qn("w:fldCharType"), "end")
        run._r.append(begin)
        run._r.append(instr)
        run._r.append(separate)
        run._r.append(text)
        run._r.append(end)

    def _mark_field_dirty(self, para: Paragraph) -> None:
        for node in para._p.xpath(".//w:fldChar"):
            if node.get(qn("w:fldCharType")) == "begin":
                node.set(qn("w:dirty"), "true")

    def _is_toc_field_paragraph(self, para: Paragraph) -> bool:
        for node in para._p.xpath(".//w:instrText"):
            text = (node.text or "").upper()
            if " TOC " in f" {text} ":
                return True
        return False

    def _apply_toc_style_defaults(self, para: Paragraph) -> None:
        para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        para.paragraph_format.line_spacing = 1.5
        if not para.runs:
            para.add_run("")
        for run in para.runs:
            run.font.size = Pt(12)
            run.font.name = "宋体"
            if run._element.rPr is not None and run._element.rPr.rFonts is not None:
                rfonts = run._element.rPr.rFonts
            else:
                rpr = run._element.get_or_add_rPr()
                rfonts = rpr.get_or_add_rFonts()
            rfonts.set(qn("w:eastAsia"), "宋体")
            rfonts.set(qn("w:ascii"), "Times New Roman")
            rfonts.set(qn("w:hAnsi"), "Times New Roman")

    def _extract_toc_entries(self, doc: DocxDocument, heading_idx: int) -> list[dict[str, object]]:
        items: list[dict[str, object]] = []
        start = heading_idx + 1
        end = min(len(doc.paragraphs), start + 220)
        entry_re = re.compile(r"^\s*(.+?)\s*(?:\.{2,}|\t|\s)\s*(\d+)\s*$")
        for idx in range(start, end):
            text = (doc.paragraphs[idx].text or "").strip()
            if not text:
                continue
            if idx > start and self._detect_level_by_numbering(text) == self.LEVEL_H1:
                break
            m = entry_re.match(text)
            if not m:
                continue
            title = m.group(1).strip()
            page = int(m.group(2))
            items.append({"title": title, "page": page})
        return items

    def _normalize_compare_text(self, text: str) -> str:
        cleaned = re.sub(r"\s+", "", text or "")
        cleaned = re.sub(r"[\.。．·•]+$", "", cleaned)
        return cleaned

    def _detect_structure_starts(self) -> list[tuple[str, int]]:
        doc = self._require_document()
        labels: dict[str, int] = {"封面": 0}
        for idx, para in enumerate(doc.paragraphs):
            text = (para.text or "").strip()
            if not text:
                continue
            normalized = re.sub(r"\s+", "", text)
            if "声明" not in labels and ("声明" in normalized or "原创性声明" in normalized or "诚信声明" in normalized):
                labels["声明"] = idx
            if "摘要" not in labels and (normalized == "摘要" or normalized.startswith("摘要")):
                labels["摘要"] = idx
            if "目录" not in labels and (normalized == "目录" or normalized.startswith("目录")):
                labels["目录"] = idx
            if "参考文献" not in labels and normalized.startswith("参考文献"):
                labels["参考文献"] = idx
            if "致谢" not in labels and (normalized.startswith("致谢") or normalized == "致谢"):
                labels["致谢"] = idx

            if "正文" not in labels:
                level_by_num = self._detect_level_by_numbering(text)
                if level_by_num == self.LEVEL_H1 or re.match(r"^\s*绪论", text):
                    labels["正文"] = idx

        ordered = sorted(((label, pos) for label, pos in labels.items()), key=lambda x: x[1])
        return ordered

    def _apply_structure_section_breaks(self, starts: list[tuple[str, int]]) -> int:
        doc = self._require_document()
        if not doc.paragraphs:
            self._last_section_labels = []
            return 0
        ordered = sorted(starts, key=lambda x: x[1])
        self._clear_all_paragraph_section_breaks()
        breaks = 0
        section_labels: list[str] = []
        for i, (label, start_idx) in enumerate(ordered):
            section_labels.append(label)
            if i == len(ordered) - 1:
                break
            next_start = ordered[i + 1][1]
            end_idx = max(start_idx, next_start - 1)
            if self._set_section_break_after_paragraph(end_idx):
                breaks += 1
        self._last_section_labels = section_labels
        return breaks

    def _clear_all_paragraph_section_breaks(self) -> None:
        doc = self._require_document()
        for para in doc.paragraphs:
            p_pr = para._p.pPr
            if p_pr is None:
                continue
            sect_pr = p_pr.find(qn("w:sectPr"))
            if sect_pr is not None:
                p_pr.remove(sect_pr)

    def _set_section_break_after_paragraph(self, para_index: int) -> bool:
        doc = self._require_document()
        if para_index < 0 or para_index >= len(doc.paragraphs):
            return False
        para = doc.paragraphs[para_index]
        p_pr = para._p.get_or_add_pPr()
        old = p_pr.find(qn("w:sectPr"))
        if old is not None:
            p_pr.remove(old)

        body_sect = doc._element.body.sectPr
        if body_sect is not None:
            new_sect = deepcopy(body_sect)
        else:
            new_sect = OxmlElement("w:sectPr")
        sec_type = new_sect.find(qn("w:type"))
        if sec_type is None:
            sec_type = OxmlElement("w:type")
            new_sect.insert(0, sec_type)
        sec_type.set(qn("w:val"), "nextPage")
        p_pr.append(new_sect)
        return True

    def _apply_page_numbers_by_structure(self, structure_labels: list[str]) -> int:
        doc = self._require_document()
        sections = list(doc.sections)
        labels = structure_labels if len(structure_labels) == len(sections) else self._last_section_labels
        if len(labels) != len(sections):
            labels = [f"节{i + 1}" for i in range(len(sections))]

        changed = 0
        arabic_started = False
        roman_started = False
        body_pos: int | None = None
        for i, name in enumerate(labels):
            if name == "正文":
                body_pos = i
                break
        for idx, section in enumerate(sections):
            label = labels[idx]
            section.header.is_linked_to_previous = False
            section.footer.is_linked_to_previous = False

            if label == "封面":
                if self._clear_footer_page_number(section):
                    changed += 1
                self._clear_section_page_start(section)
                continue

            is_front_matter = body_pos is None or idx < body_pos
            if is_front_matter:
                start = 1 if not roman_started else None
                roman_started = True
                if self._set_footer_page_number(section, fmt="upperRoman", start=start):
                    changed += 1
                continue

            # 正文 / 参考文献 / 致谢 (and fallback sections) use Arabic page numbers.
            start = 1 if not arabic_started else None
            arabic_started = True
            if self._set_footer_page_number(section, fmt="decimal", start=start):
                changed += 1
        return changed

    def _get_structure_start_index(self, target_label: str) -> int | None:
        starts = self._detect_structure_starts()
        for label, pos in starts:
            if label == target_label:
                return pos
        return None

    def _get_body_start_page_estimate(self, body_start_idx: int | None) -> int:
        if not self._paragraphs:
            return 1
        if body_start_idx is None:
            return 1
        for item in self._paragraphs:
            idx = int(item.get("index", -1))
            if idx < body_start_idx:
                continue
            text = str(item.get("text", "")).strip()
            if not text:
                continue
            return max(1, int(item.get("page_est", 1)))
        return 1

    def _set_footer_page_number(self, section: Any, fmt: str, start: int | None = None) -> bool:
        para = section.footer.paragraphs[0] if section.footer.paragraphs else section.footer.add_paragraph()
        self._clear_paragraph_content(para)
        para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = para.add_run()
        run.font.name = "Times New Roman"
        run.font.size = Pt(10.5)
        if run._element.rPr is not None and run._element.rPr.rFonts is not None:
            rfonts = run._element.rPr.rFonts
        else:
            rpr = run._element.get_or_add_rPr()
            rfonts = rpr.get_or_add_rFonts()
        rfonts.set(qn("w:ascii"), "Times New Roman")
        rfonts.set(qn("w:hAnsi"), "Times New Roman")
        field_code = "PAGE \\* ROMAN" if fmt == "upperRoman" else "PAGE"
        self._append_field(run, field_code)
        self._set_section_page_number_type(section, fmt=fmt, start=start)
        return True

    def _clear_footer_page_number(self, section: Any) -> bool:
        para = section.footer.paragraphs[0] if section.footer.paragraphs else section.footer.add_paragraph()
        had_content = bool((para.text or "").strip() or para._p.xpath(".//w:fldChar"))
        if had_content:
            self._clear_paragraph_content(para)
        return had_content

    def _set_section_page_number_type(self, section: Any, fmt: str | None = None, start: int | None = None) -> None:
        sect_pr = section._sectPr
        node = sect_pr.find(qn("w:pgNumType"))
        if fmt is None and start is None:
            if node is not None:
                sect_pr.remove(node)
            return
        if node is None:
            node = OxmlElement("w:pgNumType")
            sect_pr.append(node)
        if fmt is not None:
            node.set(qn("w:fmt"), fmt)
        else:
            if node.get(qn("w:fmt")) is not None:
                del node.attrib[qn("w:fmt")]
        if start is not None:
            node.set(qn("w:start"), str(max(1, int(start))))
        elif node.get(qn("w:start")) is not None:
            del node.attrib[qn("w:start")]

    def _section_has_page_number(self, section: Any) -> bool:
        for para in section.footer.paragraphs:
            if para._p.xpath(".//w:fldChar"):
                return True
        return False

    def _require_document(self) -> DocxDocument:
        if self._document is None:
            raise RuntimeError("Document is not loaded.")
        return self._document

    def _set_paragraph_font_size(self, para: Paragraph, font_size_pt: float) -> int:
        changed = 0
        if para.runs:
            for run in para.runs:
                run.font.size = Pt(font_size_pt)
                changed += 1
        else:
            para.add_run("")
            para.runs[0].font.size = Pt(font_size_pt)
            changed += 1
        return changed

    def _set_paragraph_bold(self, para: Paragraph, bold: bool) -> int:
        changed = 0
        if para.runs:
            for run in para.runs:
                run.bold = bold
                changed += 1
        else:
            para.add_run("")
            para.runs[0].bold = bold
            changed += 1
        return changed

    def _set_paragraph_italic(self, para: Paragraph, italic: bool) -> int:
        changed = 0
        if para.runs:
            for run in para.runs:
                run.italic = italic
                changed += 1
        else:
            para.add_run("")
            para.runs[0].italic = italic
            changed += 1
        return changed

    def _set_paragraph_underline(self, para: Paragraph, underline: bool) -> int:
        changed = 0
        if para.runs:
            for run in para.runs:
                run.underline = underline
                changed += 1
        else:
            para.add_run("")
            para.runs[0].underline = underline
            changed += 1
        return changed

    def _set_paragraph_font_name(self, para: Paragraph, font_name: str) -> int:
        changed = 0
        if not para.runs:
            para.add_run("")
        for run in para.runs:
            run.font.name = font_name
            if run._element.rPr is not None and run._element.rPr.rFonts is not None:
                rfonts = run._element.rPr.rFonts
            else:
                rpr = run._element.get_or_add_rPr()
                rfonts = rpr.get_or_add_rFonts()
            rfonts.set(qn("w:eastAsia"), font_name)
            rfonts.set(qn("w:ascii"), font_name)
            rfonts.set(qn("w:hAnsi"), font_name)
            changed += 1
        return changed

    def _replace_paragraph_text(self, para: Paragraph, new_text: str) -> None:
        if para.runs:
            first_run = para.runs[0]
            first_run.text = new_text
            first_run.bold = first_run.bold
            first_run.italic = first_run.italic
            first_run.font.size = first_run.font.size
            for run in para.runs[1:]:
                run.text = ""
        else:
            para.add_run(new_text)

    def _build_search_pattern(self, query: str, case_sensitive: bool = False, whole_word: bool = False) -> re.Pattern[str]:
        escaped = re.escape(query)
        if whole_word and re.search(r"[A-Za-z0-9_]", query):
            escaped = rf"\b{escaped}\b"
        flags = 0 if case_sensitive else re.IGNORECASE
        return re.compile(escaped, flags)

    def _push_undo_snapshot(self) -> None:
        doc = self._require_document()
        bio = BytesIO()
        doc.save(bio)
        snap = bio.getvalue()
        if self._undo_stack and self._undo_stack[-1] == snap:
            return
        self._redo_stack.clear()
        self._undo_stack.append(snap)
        if len(self._undo_stack) > 30:
            self._undo_stack = self._undo_stack[-30:]

    def _extract_font_size_pt(self, para: Paragraph) -> float | None:
        for run in para.runs:
            if run.font.size is not None:
                return round(float(run.font.size.pt), 2)
        style = para.style
        if style is not None and style.font is not None and style.font.size is not None:
            return round(float(style.font.size.pt), 2)
        return None

    def _extract_font_name(self, para: Paragraph) -> str | None:
        for run in para.runs:
            if run.font.name:
                return str(run.font.name)
            if run._element.rPr is not None and run._element.rPr.rFonts is not None:
                east = run._element.rPr.rFonts.get(qn("w:eastAsia"))
                if east:
                    return str(east)
        style = para.style
        if style is not None and style.font is not None and style.font.name:
            return str(style.font.name)
        return None

    def _extract_font_style(self, para: Paragraph) -> tuple[bool, bool, bool]:
        style_font = para.style.font if para.style is not None and para.style.font is not None else None
        style_bold = style_font.bold if style_font is not None else None
        style_italic = style_font.italic if style_font is not None else None
        style_underline = style_font.underline if style_font is not None else None

        def _weighted_majority(attr: str, style_default: Any) -> bool:
            true_weight = 0
            false_weight = 0
            for run in para.runs:
                # Empty runs are still style carriers in Word. If ignored,
                # toggle state may look incorrect after clicking style buttons.
                text_weight = max(1, len(run.text or ""))
                raw = getattr(run, attr)
                effective = raw if raw is not None else style_default
                if bool(effective):
                    true_weight += text_weight
                else:
                    false_weight += text_weight
            if true_weight == 0 and false_weight == 0:
                return bool(style_default)
            return true_weight > false_weight

        return (
            _weighted_majority("bold", style_bold),
            _weighted_majority("italic", style_italic),
            _weighted_majority("underline", style_underline),
        )

    def _extract_alignment(self, para: Paragraph) -> str:
        align = para.alignment
        mapping = {
            WD_PARAGRAPH_ALIGNMENT.LEFT: "left",
            WD_PARAGRAPH_ALIGNMENT.CENTER: "center",
            WD_PARAGRAPH_ALIGNMENT.RIGHT: "right",
            WD_PARAGRAPH_ALIGNMENT.JUSTIFY: "justify",
            WD_PARAGRAPH_ALIGNMENT.DISTRIBUTE: "justify",
        }
        return mapping.get(align, "left")

    def _extract_line_spacing(self, para: Paragraph) -> float | None:
        spacing = para.paragraph_format.line_spacing
        if spacing is None:
            return None
        try:
            return round(float(spacing), 3)
        except (TypeError, ValueError):
            # Length (EMU) for exact spacing. 240 pt ~= single-line in Word style scale.
            if hasattr(spacing, "pt"):
                pt_value = float(spacing.pt)
                return round(pt_value / 12.0, 3)
            return None

    def _detect_level_by_numbering(self, text: str) -> int:
        content = text.strip()
        if not content:
            return self.LEVEL_BODY
        # Check more specific numeric levels first to avoid ambiguity.
        if self._PATTERN_H3.match(content):
            return self.LEVEL_H3
        if self._PATTERN_H2.match(content):
            return self.LEVEL_H2
        if self._PATTERN_H1.match(content):
            return self.LEVEL_H1
        if self._PATTERN_H1_NUM_TEXT.match(content) and self._is_valid_h1_num_text(content):
            return self.LEVEL_H1
        if self._PATTERN_H4.match(content):
            return self.LEVEL_H4
        # Legacy fallback, keep previous behavior compatibility.
        if self._PATTERN_H1_LEGACY.match(content):
            return self.LEVEL_H1
        if self._PATTERN_H2_LEGACY.match(content):
            return self.LEVEL_H2
        if self._PATTERN_H3_LEGACY.match(content):
            return self.LEVEL_H3
        return self.LEVEL_BODY

    def _detect_level_by_style(self, para: Paragraph, font_size_pt: float | None, bold: bool, italic: bool) -> int:
        style_name = (para.style.name if para.style is not None else "").lower()
        content = (para.text or "").strip()

        if "heading 4" in style_name or "标题 4" in style_name:
            return self.LEVEL_H4
        if "heading 3" in style_name or "标题 3" in style_name:
            return self.LEVEL_H3
        if "heading 2" in style_name or "标题 2" in style_name:
            return self.LEVEL_H2
        if "heading 1" in style_name or "标题 1" in style_name:
            return self.LEVEL_H1

        if font_size_pt is None:
            return self.LEVEL_BODY

        # Style-assisted fallback when numbering pattern is absent.
        # Constrain fallback to heading-like short lines to avoid sentence misclassification.
        if not self._is_heading_like_short_text(content, max_len=30):
            return self.LEVEL_BODY
        if bold and font_size_pt >= 16:
            return self.LEVEL_H1
        if bold and font_size_pt >= 14:
            return self.LEVEL_H2
        if bold and font_size_pt >= 12.5:
            return self.LEVEL_H3
        if italic and font_size_pt >= 12:
            return self.LEVEL_H4
        return self.LEVEL_BODY

    def _is_valid_h1_num_text(self, content: str) -> bool:
        text = content.strip()
        if not text:
            return False
        if not self._is_heading_like_short_text(text, max_len=30):
            return False
        if self._is_year_sentence_like(text):
            return False
        match = re.match(r"^\s*(\d+)\s*(.+)$", text)
        if not match:
            return False
        tail = match.group(2).strip()
        if not tail:
            return False
        # Block common time-expression sentences: 08年/2008年/3月/12日...
        if re.match(r"^(年|月|日|季度|学年|年代)", tail):
            return False
        # Keyword must appear near the beginning of numeric-title tail.
        return self._contains_h1_keyword_near_start(tail)

    def _is_year_sentence_like(self, text: str) -> bool:
        return bool(re.match(r"^\s*(?:\d{2}|\d{4})\s*年(?:\s*\d{1,2}\s*月)?", text))

    def _contains_h1_keyword_near_start(self, tail: str) -> bool:
        compact = re.sub(r"\s+", "", tail)
        if not compact:
            return False
        prefix = compact[:10].lower()
        return any(keyword.lower() in prefix for keyword in self._H1_NUM_TEXT_KEYWORDS)

    def _is_heading_like_short_text(self, text: str, max_len: int = 30) -> bool:
        stripped = text.strip()
        if not stripped:
            return False
        if len(stripped) >= max_len:
            return False
        if re.search(r"[。！？!?；;，,：:]", stripped):
            return False
        if stripped.endswith(("、",)):
            return False
        return True

    def _mark_big_title(self, items: list[dict[str, Any]]) -> None:
        if not items:
            return

        trigger_keywords = ("摘要", "abstract", "关键词", "keywords", "目录", "声明", "绪论", "第")
        module_keywords = ("绪论", "理论基础", "现状", "成因", "分析", "对策", "启示", "结论")
        numbering_like = re.compile(r"(第\s*[0-9一二三四五六七八九十]+\s*章|\d+\.\d+|\d+[、.]|[（(][一二三四五六七八九十0-9]+[）)])")

        first_heading_idx: int | None = None
        for item in items:
            if item["detected_level"] in {self.LEVEL_H1, self.LEVEL_H2, self.LEVEL_H3, self.LEVEL_H4}:
                first_heading_idx = item["index"]
                break

        trigger_pos: int | None = None
        for item in items:
            text_lower = str(item["text"]).strip().lower()
            if not text_lower:
                continue
            if any(k in text_lower for k in trigger_keywords):
                trigger_pos = int(item["index"])
                break

        candidate_pos: int | None = None
        for i, item in enumerate(items):
            text = item["text"].strip()
            if not text:
                continue
            if item["detected_level"] != self.LEVEL_BODY:
                continue
            if first_heading_idx is not None and item["index"] > first_heading_idx:
                break
            if trigger_pos is not None and int(item["index"]) >= trigger_pos:
                break
            if len(text) > 40:
                continue
            if len(text) <= 5:
                continue
            if any(k in text for k in module_keywords):
                continue
            if numbering_like.search(text):
                continue
            if str(item.get("alignment", "")).lower() != "center":
                continue
            is_title_like = (
                bool(item["bold"])
                or (item["font_size_pt"] is not None and item["font_size_pt"] >= 16)
            )
            if is_title_like:
                candidate_pos = i
                break

        if candidate_pos is not None:
            items[candidate_pos]["detected_level"] = self.LEVEL_BIG_TITLE

    def _normalize_level(self, level: int | str) -> int:
        if isinstance(level, int):
            return level
        text = level.strip().lower()
        mapping = {
            "title": self.LEVEL_BIG_TITLE,
            "big": self.LEVEL_BIG_TITLE,
            "big_title": self.LEVEL_BIG_TITLE,
            "h1": self.LEVEL_H1,
            "h2": self.LEVEL_H2,
            "h3": self.LEVEL_H3,
            "h4": self.LEVEL_H4,
            "body": self.LEVEL_BODY,
        }
        return mapping.get(text, self.LEVEL_BODY)

    def _resolve_heading_anchor(self, cursor_pos: int) -> int | None:
        current_level = self._paragraphs[cursor_pos]["detected_level"]
        if self._heading_rank(current_level) is not None:
            return cursor_pos

        for pos in range(cursor_pos - 1, -1, -1):
            level = self._paragraphs[pos]["detected_level"]
            if self._heading_rank(level) is not None:
                return pos
        return None

    def _heading_rank(self, level: int) -> int | None:
        # Smaller rank means higher-level heading.
        rank_map = {
            self.LEVEL_BIG_TITLE: 0,
            self.LEVEL_H1: 1,
            self.LEVEL_H2: 2,
            self.LEVEL_H3: 3,
            self.LEVEL_H4: 4,
        }
        return rank_map.get(level)
