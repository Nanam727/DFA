"""批次4回归：论文结构蓝图 schema + 结构检查器 + UI 选取面板。

运行：python tests/test_batch4_structure.py
"""

from __future__ import annotations

import json
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from docx import Document

import src.ui as ui_module
from src.structure_checker import (
    check_structure,
    extract_numbering,
    normalize_title,
)
from src.thesis_schema import (
    OPTIONAL,
    REQUIRED,
    SCHOOL,
    applicable_items,
    default_selection,
    get_item,
    is_applicable,
    req_for,
)

PASS = "PASS"
FAIL = "FAIL"
_results: list[tuple[str, str, str]] = []


def _record(name: str, ok: bool, detail: str = "") -> None:
    _results.append((name, PASS if ok else FAIL, detail))
    print(f"{PASS if ok else FAIL} {name}" + (f" :: {detail}" if detail else ""))


def _check(name: str, condition: bool, detail: str = "") -> bool:
    _record(name, bool(condition), detail)
    return bool(condition)


# ---------------------------------------------------------------- schema 测试


def test_schema_defaults() -> None:
    liberal = default_selection("undergrad_liberal")
    stem = default_selection("undergrad_stem")
    course = default_selection("course_paper")

    for required_id in ("cover", "abstract_cn", "toc", "references", "acknowledgments"):
        _check(f"文科默认勾选 {required_id}", required_id in liberal)
    _check("文科默认勾选第5章对策", "ch5" in liberal)
    _check("文科默认勾选综述子节", {"ch2_2", "ch2_3"} <= liberal)
    _check("理工不含综述子节", "ch2_2" not in stem)
    _check("理工默认不勾第5章（可选）", "ch5" not in stem)
    _check("课程论文不含封面", "cover" not in course)
    _check("课程论文不含原创性声明", "declaration" not in course)
    for optional_id in ("fig_toc", "symbols", "notes", "appendix", "abstract_en"):
        _check(f"可选项默认不勾 {optional_id}", optional_id not in liberal)
    _check("默认选取仅含必写项", all(
        req_for(get_item(i), "undergrad_liberal") == REQUIRED for i in liberal
    ))


def test_schema_req_and_scope() -> None:
    _check("第5章对文科必写", req_for(get_item("ch5"), "undergrad_liberal") == REQUIRED)
    _check("第5章对理工可选", req_for(get_item("ch5"), "undergrad_stem") == OPTIONAL)
    _check("英文摘要视学校", req_for(get_item("abstract_en"), "undergrad_liberal") == SCHOOL)
    _check("封面不适用于课程论文", not is_applicable(get_item("cover"), "course_paper"))
    _check("研究述评仅文科", not is_applicable(get_item("ch2_3"), "undergrad_stem"))
    _check(
        "applicable_items 顺序含父子",
        [i["id"] for i in applicable_items("undergrad_liberal")][:4]
        == ["cover", "declaration", "abstract_cn", "abstract_en"],
    )


def test_normalize_helpers() -> None:
    cases = [
        ("第一章 绪论", "1", "绪论"),
        ("1.3 国内外研究现状", "1.3", "国内外研究现状"),
        ("（三）研究意义", "3", "研究意义"),
        ("2．1 核心概念", "2.1", "核心概念"),
        ("(2) 研究方法", "2", "研究方法"),
        ("摘要", "", "摘要"),
        ("Abstract", "", "abstract"),
    ]
    ok = True
    for text, num, norm in cases:
        got_num = extract_numbering(text)
        got_norm = normalize_title(text)
        if got_num != num or got_norm != norm:
            ok = False
            print(f"   mismatch {text!r}: num={got_num!r} (want {num!r}), "
                  f"norm={got_norm!r} (want {norm!r})")
    _check("编号提取与标题归一化（7 例）", ok)


# ---------------------------------------------------------------- 检查器测试


def _full_doc_paragraphs() -> list[dict]:
    """一份完整的本科文科论文标题序列（故意缺 1.4/1.5/6.1/6.2）。"""
    spec = [
        ("封面", 5),
        ("原创性声明", 1),
        ("摘要", 1),
        ("目录", 1),
        ("第1章 绪论", 1),
        ("1.1 研究背景", 2),
        ("1.2 研究意义", 2),
        ("1.3 国内外研究现状", 2),
        ("第2章 文献综述", 1),
        ("2.1 核心概念界定", 2),
        ("2.2 国内研究现状", 2),
        ("2.3 研究述评", 2),
        ("第3章 现状分析", 1),
        ("第4章 问题与成因", 1),
        ("第5章 对策与建议", 1),
        ("第6章 结论", 1),
        ("参考文献", 1),
        ("致谢", 1),
    ]
    return [
        {"index": i, "text": text, "detected_level": level}
        for i, (text, level) in enumerate(spec)
    ]


def test_checker_full_document() -> None:
    report = check_structure(
        _full_doc_paragraphs(), default_selection("undergrad_liberal"),
        "undergrad_liberal",
    )
    _check(
        "完整论文仅缺未写的 4 个子节",
        report["missing"] == ["ch1_4", "ch1_5", "ch6_1", "ch6_2"],
        f"missing={report['missing']}",
    )
    _check("必写项全部已勾选", report["not_selected_required"] == [])
    _check("无未选取的文档章节", report["present_unselected"] == [])
    _check("第6章定位到“第6章 结论”", report["found"].get("ch6") == 15)
    _check("1.3 定位正确", report["found"].get("ch1_3") == 7)


def test_checker_missing_required() -> None:
    paras = [
        {"index": 0, "text": "第1章 绪论", "detected_level": 1},
        {"index": 1, "text": "参考文献", "detected_level": 1},
    ]
    report = check_structure(
        paras, default_selection("undergrad_liberal"), "undergrad_liberal"
    )
    _check("缺失摘要被检出", "abstract_cn" in report["missing"])
    _check("缺失致谢被检出", "acknowledgments" in report["missing"])
    _check("缺失目录被检出", "toc" in report["missing"])
    _check("绪论已找到", "ch1" in report["found"])


def test_checker_synonym_and_body_text() -> None:
    # 同义词：“总结”应命中“结论与展望”
    paras = [{"index": 3, "text": "总结", "detected_level": 1}]
    report = check_structure(
        paras, {"ch6"}, "undergrad_liberal"
    )
    _check("同义词“总结”命中结论章", report["found"].get("ch6") == 3)
    _check("命中后无缺失", report["missing"] == [])

    # 正文长句含关键词不算命中（防止误报）
    body = [
        {"index": 0, "text": "本文的研究背景主要包括行业发展与政策变化两个方面。",
         "detected_level": 0},
    ]
    report = check_structure(body, {"ch1_1"}, "undergrad_liberal")
    _check("正文含关键词不算命中", "ch1_1" in report["missing"],
           f"found={report['found']}")


def test_checker_numbering_conflict() -> None:
    # 文档只有 2.2“国内研究现状”，1.3 缺失时不应被 2.2 冒充
    paras = [
        {"index": 0, "text": "第2章 文献综述", "detected_level": 1},
        {"index": 1, "text": "2.2 国内研究现状", "detected_level": 2},
    ]
    report = check_structure(
        paras, {"ch1_3", "ch2_2"}, "undergrad_liberal"
    )
    _check("1.3 缺失时正确报缺失", "ch1_3" in report["missing"])
    _check("2.2 正常命中", report["found"].get("ch2_2") == 1)


def test_checker_unselected_reports() -> None:
    paras = _full_doc_paragraphs() + [
        {"index": 18, "text": "附录", "detected_level": 1},
    ]
    selected = default_selection("undergrad_liberal") - {"references"}
    report = check_structure(paras, selected, "undergrad_liberal")
    _check("未勾选必写项被提醒", "references" in report["not_selected_required"])
    _check("未勾选项不进 missing", "references" not in report["missing"])
    _check("文档有但未选取被提示", "appendix" in report["present_unselected"])


# ---------------------------------------------------------------- UI 测试


def test_ui_structure_panel() -> None:
    app = ui_module.FormatApp()
    app.withdraw()
    original_showinfo = ui_module.messagebox.showinfo
    ui_module.messagebox.showinfo = lambda *a, **k: None
    try:
        with tempfile.TemporaryDirectory() as tmp:
            app.app_state.preference_file = Path(tmp) / "prefs.json"
            app._load_structure_selection()

            _check("复选框树已构建", len(app.structure_item_vars) > 10)
            liberal_default = default_selection("undergrad_liberal")
            checked = {
                i for i, v in app.structure_item_vars.items() if v.get()
            }
            _check("默认勾选与蓝图一致", checked == liberal_default,
                   f"checked={len(checked)}, default={len(liberal_default)}")
            _check("论文类型下拉存在", app.structure_type_combo is not None)

            # 切换为理工科：可选章不勾、综述子节消失
            app.on_structure_type_change("本科·理工科")
            stem_ids = set(app.structure_item_vars)
            _check("理工不含综述子节复选框", "ch2_2" not in stem_ids)
            _check("理工默认不勾第5章",
                   not app.structure_item_vars["ch5"].get())
            _check("理工仍勾必写项",
                   app.structure_item_vars["references"].get())
            _check("状态已记录论文类型",
                   app.app_state.structure_selection["paper_type"] == "undergrad_stem")

            # 切换勾选 -> 状态同步
            app.structure_item_vars["ch5"].set(True)
            app.on_structure_item_toggle()
            _check("勾选同步到 state", "ch5" in app._selected_structure_ids())

            # 全选
            app.on_select_all_structure_items()
            _check("全选后所有复选框为真",
                   all(v.get() for v in app.structure_item_vars.values()))

            # 重置
            app.on_reset_structure_selection()
            _check("重置后恢复默认",
                   {i for i, v in app.structure_item_vars.items() if v.get()}
                   == default_selection("undergrad_stem"))

            # 持久化往返
            app.on_save_structure_selection()
            saved = json.loads(
                app.app_state.preference_file.read_text(encoding="utf-8")
            )
            _check("偏好文件含 structure 键", "structure" in saved)
            _check("保存的论文类型正确",
                   saved["structure"]["paper_type"] == "undergrad_stem")
            app.app_state.structure_selection = {
                "paper_type": "undergrad_liberal", "selected": []
            }
            app._load_structure_selection()
            _check("重新加载恢复理工选取",
                   app.app_state.structure_selection["paper_type"] == "undergrad_stem")
            _check("重新加载恢复勾选项",
                   "references" in app._selected_structure_ids())

            # 检查结构（未打开文档时不弹窗、不异常）
            app.on_check_structure()
            _check("未打开文档时检查结构安全通过", True)

            # 打开文档后检查 + 定位
            doc_path = Path(tmp) / "thesis.docx"
            doc = Document()
            for para in _full_doc_paragraphs():
                doc.add_paragraph(para["text"])
            doc.save(str(doc_path))
            app.docx_handler.load(str(doc_path))
            app._refresh_document_view()

            captured: dict = {}
            original_report = app._show_structure_report

            def _capture(report: dict) -> None:
                captured.update(report)

            app._show_structure_report = _capture
            try:
                app.on_check_structure()
            finally:
                app._show_structure_report = original_report
            _check("检查产出报告", "missing" in captured)
            _check("报告缺失 1.4/1.5/6.1/6.2",
                   captured.get("missing") == ["ch1_4", "ch1_5", "ch6_1", "ch6_2"],
                   f"missing={captured.get('missing')}")

            # 一键定位：found 中的段落索引可用 _set_selection 选中
            idx = captured["found"]["references"]
            app._set_selection({idx}, "测试定位")
            _check("定位后段落被选中", app.selected_indices == {idx},
                   f"selected={app.selected_indices}")
    finally:
        ui_module.messagebox.showinfo = original_showinfo
        app.destroy()


def main() -> int:
    test_schema_defaults()
    test_schema_req_and_scope()
    test_normalize_helpers()
    test_checker_full_document()
    test_checker_missing_required()
    test_checker_synonym_and_body_text()
    test_checker_numbering_conflict()
    test_checker_unselected_reports()
    test_ui_structure_panel()
    failed = [r for r in _results if r[1] == FAIL]
    print(f"\n批次4回归：{len(_results) - len(failed)}/{len(_results)} 通过")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
