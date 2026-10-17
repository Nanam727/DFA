"""论文结构检查器：把选取的章节蓝图与文档段落做容错匹配。

匹配策略（两阶段贪心，按文档顺序）：
1. 第一阶段只扫描“标题段”（detected_level >= 1），允许 编号命中/整体相等/模式包含；
2. 第二阶段对仍未命中的条目扫描全部段落，但只允许 整体相等/段落被模式包含
   （正文长句里出现关键词不算命中，避免误报）。
评分：编号相等 3 > 归一化后整体相等 2 > 模式被段落包含 2 > 段落被模式包含 1。
同一段落命中多个条目时，按蓝图顺序取先出现者（如“国内外研究现状”同时
像 1.3 与 2.2 时，先出现在文档里的段落优先归属 1.3）。
"""

from __future__ import annotations

import re
from collections.abc import Iterable
from typing import Any

from src.thesis_schema import (
    REQUIRED,
    applicable_items,
    match_patterns,
    req_for,
)

# 全角字母数字 -> 半角
_FW2HW = {ord("０") + i: ord("0") + i for i in range(10)}
_FW2HW.update({ord("Ａ") + i: ord("A") + i for i in range(26)})
_FW2HW.update({ord("ａ") + i: ord("a") + i for i in range(26)})
_FW2HW[ord("．")] = ord(".")
_FW2HW[ord("。")] = ord(".")

# 段首编号：第1章 / 第三章 / 1. / 1.1 / 1．1 、/ 一、
_CN_NUM = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
_CN_NUM.update({"十": 10})


def _cn_to_int(text: str) -> int | None:
    if text.isdigit():
        return int(text)
    if text == "十":
        return 10
    if "十" in text:
        left, _, right = text.partition("十")
        tens = _CN_NUM.get(left, 1) if left else 1
        ones = _CN_NUM.get(right, 0) if right else 0
        return tens * 10 + ones
    return _CN_NUM.get(text)


_NUM_PREFIX_RE = re.compile(
    r"^\s*(?:"
    r"第\s*([0-9０-９一二三四五六七八九十百]+)\s*[章节部篇]?"  # 第1章 / 第三章
    r"|([0-9０-９]+(?:\.[0-9０-９]+)*)\s*[\.、．\s]"            # 1. / 1.1 / 1．
    r"|[（(]\s*([0-9０-９一二三四五六七八九十]+)\s*[）)]\s*[、.．]?"  # （一）/ (2)
    r"|([一二三四五六七八九十]+)\s*[、.．]"                    # 一、
    r")"
)

# 归一化时剥离的空白与标点
_STRIP_CHARS_RE = re.compile(
    r"[\s，。、；：？！（）()【】\[\]{}《》<>«»“”\"'·—\-–_/\\|+*#@$%^&~`]+"
)

_SCORE_NUMBERING = 3
_SCORE_EXACT = 2
_SCORE_PATTERN_IN_PARA = 2
_SCORE_PARA_IN_PATTERN = 1


def to_half_width(text: str) -> str:
    return text.translate(_FW2HW)


def extract_numbering(text: str) -> str:
    """提取段首编号，归一为阿拉伯点分形式（如 "1.3"）；无编号返回 ""。"""
    if not text:
        return ""
    text = to_half_width(text)
    m = _NUM_PREFIX_RE.match(text)
    if not m:
        return ""
    cn_chapter, arabic_dotted, paren_num, cn_comma = m.groups()
    if arabic_dotted:
        return arabic_dotted.strip(".")
    if cn_chapter:
        value = _cn_to_int(cn_chapter)
        return str(value) if value is not None else ""
    if paren_num:
        value = _cn_to_int(paren_num)
        return str(value) if value is not None else ""
    if cn_comma:
        value = _cn_to_int(cn_comma)
        return str(value) if value is not None else ""
    return ""


def strip_numbering(text: str) -> str:
    """去掉段首编号前缀。"""
    if not text:
        return ""
    text = to_half_width(text)
    return _NUM_PREFIX_RE.sub("", text, count=1)


def normalize_title(text: str) -> str:
    """标题归一化：去编号、去空白标点、小写。用于包含/相等比较。"""
    if not text:
        return ""
    text = to_half_width(text)
    text = _NUM_PREFIX_RE.sub("", text, count=1)
    return _STRIP_CHARS_RE.sub("", text).casefold()


def _para_text(para: dict[str, Any]) -> str:
    text = para.get("text", "")
    return text if isinstance(text, str) else ""


def _para_level(para: dict[str, Any]) -> int:
    try:
        return int(para.get("detected_level", 0) or 0)
    except (TypeError, ValueError):
        return 0


def _is_heading(para: dict[str, Any]) -> bool:
    # 0=正文，1-4=一至四级标题，5=大标题；>=1 视为标题段
    return _para_level(para) >= 1


def _match_score(
    para_text: str,
    para_numbering: str,
    item: dict[str, Any],
    allow_pattern_in_para: bool,
) -> int:
    """计算段落与条目的匹配分；0 表示不匹配。"""
    norm_para = normalize_title(para_text)
    if len(norm_para) < 2:
        return 0

    item_num = item.get("num")
    if item_num and para_numbering and para_numbering == str(item_num):
        return _SCORE_NUMBERING

    norm_item = normalize_title(item.get("name", ""))
    if norm_item and norm_para == norm_item:
        return _SCORE_EXACT

    best = 0
    for pattern in match_patterns(item):
        norm_pattern = normalize_title(pattern)
        if len(norm_pattern) < 2:
            continue
        if norm_pattern in norm_para:
            if allow_pattern_in_para:
                best = max(best, _SCORE_PATTERN_IN_PARA)
            continue
        if norm_para in norm_pattern:
            best = max(best, _SCORE_PARA_IN_PATTERN)
    return best


def check_structure(
    paragraphs: Iterable[dict[str, Any]],
    selected_ids: Iterable[str],
    paper_type: str,
) -> dict[str, Any]:
    """检查文档结构。

    Args:
        paragraphs: docx_handler.get_paragraphs() 的段落字典列表，
            需含 "index" / "text" / "detected_level" 键。
        selected_ids: 用户在面板中勾选的条目 id 集合。
        paper_type: 论文类型（见 thesis_schema.PAPER_TYPES）。

    Returns:
        {
          "paper_type": str,
          "found": {item_id: para_index},        # 命中的条目 -> 段落索引
          "missing": [item_id],                  # 选取了但文档中没找到
          "not_selected_required": [item_id],    # 必写但用户没勾选
          "present_unselected": [item_id],       # 文档里有但用户没勾选
        }
        列表均按蓝图文档顺序排列。
    """
    para_list = [p for p in paragraphs if _para_text(p).strip()]
    selected = {str(i) for i in selected_ids}
    items = applicable_items(paper_type)
    item_order = {item["id"]: pos for pos, item in enumerate(items)}
    selected = {i for i in selected if i in item_order}

    claims: dict[str, int] = {}
    claimed_paras: set[int] = set()

    def _run_pass(candidates: list[dict[str, Any]], allow_pattern_in_para: bool) -> None:
        for para in candidates:
            para_index = para.get("index")
            if para_index in claimed_paras:
                continue
            para_text = _para_text(para)
            para_numbering = extract_numbering(para_text)
            best_item: str | None = None
            best_score = 0
            for item in items:
                if item["id"] in claims:
                    continue
                score = _match_score(
                    para_text, para_numbering, item, allow_pattern_in_para
                )
                if score > best_score:
                    best_score = score
                    best_item = item["id"]
            if best_item is not None and best_score >= 1:
                claims[best_item] = int(para_index) if para_index is not None else -1
                claimed_paras.add(para_index)

    # 第一阶段：仅标题段，允许关键词包含
    _run_pass([p for p in para_list if _is_heading(p)], allow_pattern_in_para=True)
    # 第二阶段：全部段落，只允许整体相等/被包含（防止正文误命中）
    _run_pass(para_list, allow_pattern_in_para=False)

    def _ordered(ids: Iterable[str]) -> list[str]:
        return sorted(set(ids), key=lambda i: item_order.get(i, 10**6))

    missing = _ordered(i for i in selected if i not in claims)
    not_selected_required = _ordered(
        item["id"]
        for item in items
        if item["id"] not in selected and req_for(item, paper_type) == REQUIRED
    )
    present_unselected = _ordered(i for i in claims if i not in selected)

    return {
        "paper_type": paper_type,
        "found": dict(claims),
        "missing": missing,
        "not_selected_required": not_selected_required,
        "present_unselected": present_unselected,
    }
