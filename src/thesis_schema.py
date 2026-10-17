"""本科毕业论文结构蓝图（声明式数据层）。

设计约束：
- 与 DocxHandler 的 LEVEL_H1/H2/H3 分级体系完全解耦，本模块只描述
  “章节蓝图 + 必写性 + 学科差异 + 匹配同义词”，不做任何文档操作。
- req 支持两种写法：
  - 字符串：对所有论文类型统一（"required" / "optional" / "school"）
  - 字典：按论文类型分别指定，如对策与建议对文科必写、对理工可选
- disc 指定适用学科：all / liberal_arts（文科经管）/ stem（理工）/ undergrad（本科两类）
- excl 列出不适用的论文类型（如课程论文不需要封面与原创性声明）
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import Any

# ---------------------------------------------------------------- 论文类型

PAPER_TYPES: dict[str, str] = {
    "undergrad_liberal": "本科·文科/经管",
    "undergrad_stem": "本科·理工科",
    "course_paper": "课程论文",
}

# ---------------------------------------------------------------- 必写性

REQUIRED = "required"
OPTIONAL = "optional"
SCHOOL = "school"  # 视学校要求

REQ_LABELS: dict[str, str] = {
    REQUIRED: "必写",
    OPTIONAL: "可选",
    SCHOOL: "视学校",
}

_DISC_SCOPES: dict[str, frozenset[str]] = {
    "all": frozenset(PAPER_TYPES),
    "undergrad": frozenset({"undergrad_liberal", "undergrad_stem"}),
    "liberal_arts": frozenset({"undergrad_liberal"}),
    "stem": frozenset({"undergrad_stem"}),
}

# ---------------------------------------------------------------- 章节蓝图
# 字段说明：
#   id       唯一标识（选取与检查结果都用它）
#   name     显示名（含编号，匹配时会剥离编号）
#   num      编号（如 "1.3"），匹配时优先按编号命中
#   req      必写性，见文件头
#   disc     适用学科
#   excl     不适用的论文类型
#   hint     格式/写作提示（展示在结构报告与悬浮提示中）
#   synonyms 匹配同义词（与 name 一起参与文档标题匹配）
#   level    建议标题级别：H1/H2/front（前置）/post（后置）/None
#   children 子级条目

SECTIONS: list[dict[str, Any]] = [
    # ======================= 第一部分：前置 =======================
    {
        "id": "cover",
        "name": "封面",
        "req": REQUIRED,
        "disc": "all",
        "excl": ["course_paper"],
        "level": None,
        "synonyms": ["封面", "题目", "标题页"],
        "hint": "使用学校统一模板，勿手打；含题目、姓名、学号、学院、专业、指导教师、完成时间。",
    },
    {
        "id": "declaration",
        "name": "原创性声明/版权使用授权书",
        "req": REQUIRED,
        "disc": "all",
        "excl": ["course_paper"],
        "level": "front",
        "synonyms": ["原创性声明", "声明", "版权使用授权书", "授权书", "学位论文使用授权书"],
        "hint": "学校模板签字页；电子稿保留页面，打印后手写签名，日期与封面提交时间一致。",
    },
    {
        "id": "abstract_cn",
        "name": "中文摘要+关键词",
        "req": REQUIRED,
        "disc": "all",
        "level": "front",
        "synonyms": ["摘要", "中文摘要", "内容摘要"],
        "hint": "200-300字：背景→目的→方法→主要结果→结论；关键词3-5个，分号隔开。",
    },
    {
        "id": "abstract_en",
        "name": "英文摘要(Abstract)+关键词",
        "req": SCHOOL,
        "disc": "all",
        "excl": ["course_paper"],
        "level": "front",
        "synonyms": ["abstract", "英文摘要"],
        "hint": "以学校通知为准；不要求可删。不是逐字对译，需扩展；方法用过去时、结论用现在时。",
    },
    {
        "id": "toc",
        "name": "目录",
        "req": REQUIRED,
        "disc": "all",
        "level": "front",
        "synonyms": ["目录", "目次"],
        "hint": "含一至三级标题及页码；摘要、参考文献、致谢、附录都编入目录。",
    },
    {
        "id": "fig_toc",
        "name": "图目录/表目录",
        "req": OPTIONAL,
        "disc": "all",
        "level": "front",
        "synonyms": ["图目录", "图表目录", "表目录", "插图目录"],
        "hint": "图表20个以上才做；位置紧跟目录之后。",
    },
    {
        "id": "symbols",
        "name": "符号、缩写说明",
        "req": OPTIONAL,
        "disc": "all",
        "level": "front",
        "synonyms": ["符号说明", "缩写说明", "符号与缩写说明", "术语表", "缩略语"],
        "hint": "理工科、公式缩写多时使用；文科可跳过。",
    },
    # ======================= 第二部分：正文 =======================
    {
        "id": "ch1",
        "name": "第1章 绪论",
        "num": "1",
        "req": REQUIRED,
        "disc": "all",
        "level": "H1",
        "synonyms": ["绪论", "绪言", "引言", "前言", "导论"],
        "hint": "讲“为什么做这个研究”；不强制写创新点，不要强行拔高。",
        "children": [
            {
                "id": "ch1_1",
                "name": "1.1 研究背景",
                "num": "1.1",
                "req": REQUIRED,
                "disc": "all",
                "level": "H2",
                "synonyms": ["研究背景", "背景"],
            },
            {
                "id": "ch1_2",
                "name": "1.2 研究意义",
                "num": "1.2",
                "req": REQUIRED,
                "disc": "all",
                "level": "H2",
                "synonyms": ["研究意义", "意义"],
                "hint": "理论意义+现实意义；应用型本科论文写不出理论意义时可合并为一段。",
            },
            {
                "id": "ch1_3",
                "name": "1.3 国内外研究现状",
                "num": "1.3",
                "req": REQUIRED,
                "disc": "all",
                "level": "H2",
                "synonyms": ["国内外研究现状", "国内外研究动态", "研究现状述评"],
                "hint": "与第2章二选一：若文献综述独立成章，此处压缩为2-3段概要，避免重复。",
            },
            {
                "id": "ch1_4",
                "name": "1.4 研究内容与研究问题",
                "num": "1.4",
                "req": REQUIRED,
                "disc": "all",
                "level": "H2",
                "synonyms": ["研究内容", "研究问题", "研究内容与方法"],
            },
            {
                "id": "ch1_5",
                "name": "1.5 研究方法",
                "num": "1.5",
                "req": REQUIRED,
                "disc": "all",
                "level": "H2",
                "synonyms": ["研究方法", "技术路线"],
                "hint": "文献研究法、案例分析法、问卷调查法等，简单说明怎么用。",
            },
            {
                "id": "ch1_6",
                "name": "1.6 论文框架",
                "num": "1.6",
                "req": OPTIONAL,
                "disc": "all",
                "level": "H2",
                "synonyms": ["论文框架", "论文结构", "组织结构"],
            },
        ],
    },
    {
        "id": "ch2",
        "name": "第2章 文献综述与理论基础",
        "num": "2",
        "req": REQUIRED,
        "disc": "all",
        "level": "H1",
        "synonyms": ["文献综述", "相关理论与技术基础", "理论基础", "理论与技术基础", "相关理论"],
        "hint": "文科经管：核心概念+国内外现状+述评；理工科：本文学用到的理论、模型、技术原理。",
        "children": [
            {
                "id": "ch2_1",
                "name": "2.1 核心概念与理论基础",
                "num": "2.1",
                "req": REQUIRED,
                "disc": "all",
                "level": "H2",
                "synonyms": ["核心概念界定", "概念界定", "理论基础", "核心概念"],
            },
            {
                "id": "ch2_2",
                "name": "2.2 国内外研究现状",
                "num": "2.2",
                "req": REQUIRED,
                "disc": "liberal_arts",
                "level": "H2",
                "synonyms": ["国内研究现状", "国外研究现状", "国内外研究现状"],
            },
            {
                "id": "ch2_3",
                "name": "2.3 研究述评",
                "num": "2.3",
                "req": REQUIRED,
                "disc": "liberal_arts",
                "level": "H2",
                "synonyms": ["研究述评", "文献述评", "述评"],
                "hint": "总结前人成果，点出现有研究不足，引出本文研究。",
            },
        ],
    },
    {
        "id": "ch3",
        "name": "第3章 现状分析（研究对象）",
        "num": "3",
        "req": REQUIRED,
        "disc": "all",
        "level": "H1",
        "synonyms": ["现状分析", "现状描述", "研究对象", "研究设计", "实验设计"],
        "hint": "文科经管：行业/案例现状、数据情况；理工科：实验对象、实验环境、数据来源。",
    },
    {
        "id": "ch4",
        "name": "第4章 问题与实证分析",
        "num": "4",
        "req": REQUIRED,
        "disc": "all",
        "level": "H1",
        "synonyms": ["问题与成因", "实证分析", "案例分析", "实验与结果", "结果分析", "问题分析"],
        "hint": "全文最重要一章：文科写存在问题+成因；经管写数据处理+模型分析；理工写实验过程+结果。",
    },
    {
        "id": "ch5",
        "name": "第5章 对策与建议",
        "num": "5",
        "req": {
            "undergrad_liberal": REQUIRED,
            "undergrad_stem": OPTIONAL,
            "course_paper": OPTIONAL,
        },
        "disc": "all",
        "level": "H1",
        "synonyms": ["对策与建议", "对策建议", "建议与对策"],
        "hint": "针对第4章发现的问题提出可行对策；理工科无对策可不设本章。",
    },
    {
        "id": "ch6",
        "name": "第6章 结论与展望",
        "num": "6",
        "req": REQUIRED,
        "disc": "all",
        "level": "H1",
        "synonyms": ["结论与展望", "总结与展望"],
        "hint": "提炼自己研究得出的结果，不要复制摘要。",
        "children": [
            {
                "id": "ch6_1",
                "name": "6.1 研究主要结论",
                "num": "6.1",
                "req": REQUIRED,
                "disc": "all",
                "level": "H2",
                "synonyms": ["研究结论", "主要结论"],
            },
            {
                "id": "ch6_2",
                "name": "6.2 研究不足",
                "num": "6.2",
                "req": REQUIRED,
                "disc": "all",
                "level": "H2",
                "synonyms": ["研究不足", "不足与展望", "局限性"],
                "hint": "客观写局限（样本小、资料有限等），本科一定要写，是加分项。",
            },
            {
                "id": "ch6_3",
                "name": "6.3 研究展望",
                "num": "6.3",
                "req": OPTIONAL,
                "disc": "all",
                "level": "H2",
                "synonyms": ["研究展望", "展望"],
            },
        ],
    },
    # ======================= 第三部分：后置 =======================
    {
        "id": "notes",
        "name": "注释",
        "req": OPTIONAL,
        "disc": "all",
        "level": "post",
        "synonyms": ["注释", "脚注"],
        "hint": "脚注优先（页面底部），用于补充解释、引用古籍；不要与参考文献混淆。",
    },
    {
        "id": "references",
        "name": "参考文献",
        "req": REQUIRED,
        "disc": "all",
        "level": "post",
        "synonyms": ["参考文献", "引用文献"],
        "hint": "GB/T 7714-2015；本科15-30篇，需含一定外文文献；只列正文引用过的。",
    },
    {
        "id": "appendix",
        "name": "附录",
        "req": OPTIONAL,
        "disc": "all",
        "level": "post",
        "synonyms": ["附录"],
        "hint": "放调查问卷全文、访谈提纲、原始数据、代码等；正文写“详见附录A”。",
    },
    {
        "id": "acknowledgments",
        "name": "致谢",
        "req": REQUIRED,
        "disc": "all",
        "level": "post",
        "synonyms": ["致谢", "鸣谢", "acknowledgements"],
        "hint": "感谢导师、同学、家人、调研协助人员；篇幅半页-1页。",
    },
]


# ---------------------------------------------------------------- 工具函数


def iter_all_items() -> Iterator[dict[str, Any]]:
    """按文档顺序产出自定义所有条目（含子级）。"""
    for item in SECTIONS:
        yield item
        yield from item.get("children", [])


def get_item(item_id: str) -> dict[str, Any] | None:
    for item in iter_all_items():
        if item.get("id") == item_id:
            return item
    return None


def item_depth(item: dict[str, Any]) -> int:
    """顶层为 0，子级为 1。"""
    for top in SECTIONS:
        if top.get("id") == item.get("id"):
            return 0
        for child in top.get("children", []):
            if child.get("id") == item.get("id"):
                return 1
    return 0


def is_applicable(item: dict[str, Any], paper_type: str) -> bool:
    """该条目是否适用于给定论文类型。"""
    if paper_type not in PAPER_TYPES:
        return False
    if paper_type in item.get("excl", []):
        return False
    scope = _DISC_SCOPES.get(item.get("disc", "all"), frozenset(PAPER_TYPES))
    return paper_type in scope


def req_for(item: dict[str, Any], paper_type: str) -> str:
    """取该条目在给定论文类型下的必写性。"""
    req = item.get("req", OPTIONAL)
    if isinstance(req, dict):
        return req.get(paper_type, OPTIONAL)
    return req


def applicable_items(paper_type: str) -> list[dict[str, Any]]:
    """按文档顺序返回适用条目（含 depth 信息，供 UI 复选框树使用）。"""
    result: list[dict[str, Any]] = []
    for top in SECTIONS:
        if not is_applicable(top, paper_type):
            continue
        result.append({**top, "depth": 0})
        for child in top.get("children", []):
            if is_applicable(child, paper_type):
                result.append({**child, "depth": 1})
    return result


def default_selection(paper_type: str) -> set[str]:
    """默认选取 = 该类型下所有“必写”条目（可选/视学校默认不勾）。"""
    return {
        item["id"]
        for item in applicable_items(paper_type)
        if req_for(item, paper_type) == REQUIRED
    }


def match_patterns(item: dict[str, Any]) -> list[str]:
    """参与文档标题匹配的模式：显示名 + 同义词。"""
    patterns = [item.get("name", "")]
    patterns.extend(item.get("synonyms", []))
    seen: set[str] = set()
    result: list[str] = []
    for p in patterns:
        if p and p not in seen:
            seen.add(p)
            result.append(p)
    return result


def req_label(item: dict[str, Any], paper_type: str) -> str:
    return REQ_LABELS.get(req_for(item, paper_type), req_for(item, paper_type))
