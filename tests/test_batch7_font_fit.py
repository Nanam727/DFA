"""批次7：预览整页高度自适应 + 字号+-多种字号各自步进。

覆盖：
A. apply_font_size_step：多种字号各自升/降一档（保持相对差异）、
   自定义字号走最近档、顶格/底格不动、空集合、撤销还原。
B. UI 分流：混合选中走"各自步进"且清空字号框；单一选中走统一步进。
C. 预览高度自适应：大字号整页超高时按比例缩小该页字号与段距，
   正常字号不触发；行宽估算使用最近绘制的页面尺寸。
"""

from __future__ import annotations

import re
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from docx import Document
from docx.shared import Pt

from src.docx_handler import DocxHandler
from src.ui import FormatApp

POINTS = [10.5, 12.0, 14.0, 15.0, 16.0, 18.0, 22.0, 24.0, 28.0, 32.0]

PASS = 0
FAIL = 0


def _check(name: str, cond: bool, detail: str = "") -> None:
    global PASS, FAIL
    if cond:
        PASS += 1
        print(f"  PASS  {name}")
    else:
        FAIL += 1
        print(f"  FAIL  {name}  {detail}")


def _make_doc(path: Path, sizes: list[float], body_chars: int = 14) -> None:
    doc = Document()
    for i, pt in enumerate(sizes):
        p = doc.add_paragraph(f"第{i}章 内容" + "正文内容文字，" * body_chars)
        for r in p.runs:
            r.font.size = Pt(pt)
    doc.save(str(path))


def _font_size_of(tag_font: str) -> int:
    m = re.search(r"(\d+)\s*$", str(tag_font))
    return int(m.group(1)) if m else -1


# ---------------- A. apply_font_size_step ----------------

def test_apply_font_size_step() -> None:
    print("[A] apply_font_size_step 多种字号各自步进")
    with tempfile.TemporaryDirectory() as tmp:
        f = Path(tmp) / "mix.docx"
        _make_doc(f, [12.0, 16.0, 18.0])
        h = DocxHandler()
        h.load(str(f))
        _check("原始字号", [p["font_size_pt"] for p in h.get_paragraphs()] == [12.0, 16.0, 18.0])

        changed = h.apply_font_size_step({0, 1, 2}, POINTS, 1)
        after_up = [p["font_size_pt"] for p in h.get_paragraphs()]
        _check("+1 档改动 3 段", changed == 3, f"changed={changed}")
        _check("+1 档各自升档", after_up == [14.0, 18.0, 22.0], f"{after_up}")

        changed = h.apply_font_size_step({0, 1, 2}, POINTS, -1)
        after_down = [p["font_size_pt"] for p in h.get_paragraphs()]
        _check("-1 档各自降档", after_down == [12.0, 16.0, 18.0], f"{after_down}")

        # 自定义字号 13pt：+1 -> 14（大于它的最近档）
        f2 = Path(tmp) / "custom.docx"
        _make_doc(f2, [13.0])
        h2 = DocxHandler()
        h2.load(str(f2))
        h2.apply_font_size_step({0}, POINTS, 1)
        _check("自定义 13pt +1 -> 14",
               h2.get_paragraphs()[0]["font_size_pt"] == 14.0)
        h2.apply_font_size_step({0}, POINTS, -1)
        _check("自定义 14pt -1 -> 12",
               h2.get_paragraphs()[0]["font_size_pt"] == 12.0)

        # 顶格不动
        f3 = Path(tmp) / "max.docx"
        _make_doc(f3, [32.0])
        h3 = DocxHandler()
        h3.load(str(f3))
        changed = h3.apply_font_size_step({0}, POINTS, 1)
        _check("顶格 32pt +1 不动", changed == 0
               and h3.get_paragraphs()[0]["font_size_pt"] == 32.0)

        # 空集合 / step=0
        _check("空集合返回 0", h.apply_font_size_step(set(), POINTS, 1) == 0)
        _check("step=0 返回 0", h.apply_font_size_step({0}, POINTS, 0) == 0)

        # 撤销还原
        h.apply_font_size_step({0, 1, 2}, POINTS, 1)
        h.undo_last_action()
        _check("撤销还原", [p["font_size_pt"] for p in h.get_paragraphs()] == [12.0, 16.0, 18.0])


# ---------------- B. UI 分流 ----------------

def test_ui_step_routing() -> None:
    print("[B] UI 字号+- 分流逻辑")
    with tempfile.TemporaryDirectory() as tmp:
        f = Path(tmp) / "mix.docx"
        _make_doc(f, [12.0, 16.0, 18.0], body_chars=6)
        app = FormatApp()
        app.withdraw()
        try:
            app.docx_handler.load(str(f))
            app._refresh_document_view()
            app._set_selection({0, 1, 2}, "全选")
            _check("混合字号集合大小 3", len(app._selected_font_sizes()) == 3)

            app.on_increase_font_size()
            sizes_up = [p["font_size_pt"] for p in app.docx_handler.get_paragraphs()]
            _check("混合 + 各自升一档", sizes_up == [14.0, 18.0, 22.0], f"{sizes_up}")
            _check("混合后字号框清空", app.size_combo.get() == "",
                   f"combo={app.size_combo.get()!r}")
            _check("提示语含各自字号", "各自字号" in app.status_text,
                   app.status_text)

            app.on_decrease_font_size()
            sizes_down = [p["font_size_pt"] for p in app.docx_handler.get_paragraphs()]
            _check("混合 - 各自降一档", sizes_down == [12.0, 16.0, 18.0], f"{sizes_down}")

            # 单一字号走统一步进
            app._set_selection({0}, "单段")
            app.size_combo.set("五号")  # 10.5 -> 下一档 12
            app.on_increase_font_size()
            sizes_single = [p["font_size_pt"] for p in app.docx_handler.get_paragraphs()]
            _check("单一字号统一升档", sizes_single == [12.0, 16.0, 18.0], f"{sizes_single}")
            _check("统一步进后字号框显示目标值", app.size_combo.get() == "小四号",
                   f"combo={app.size_combo.get()!r}")
        finally:
            app.destroy()


# ---------------- C. 预览高度自适应 ----------------

def test_preview_height_fit() -> None:
    print("[C] 预览整页高度自适应")
    with tempfile.TemporaryDirectory() as tmp:
        # 大字号整页超高（模拟截图场景：小二号、内容接近一页容量）
        big = Path(tmp) / "big.docx"
        _make_doc(big, [18.0] * 14)
        app = FormatApp()
        app.withdraw()
        try:
            app.docx_handler.load(str(big))
            app._draw_a4_page(900, 600)
            units = app._estimate_line_units(18)
            _check("行宽估算使用绘制页宽（>15 单位）", units > 15.0, f"units={units:.1f}")

            app._refresh_document_view()
            fit = app._last_preview_fit
            _check("大字号页触发高度适配", fit is not None)
            if fit:
                shrink = float(fit["shrink"])
                _check("适配比例小于 1", 0.25 <= shrink < 1.0, f"shrink={shrink:.3f}")
                _check("适配后估算高度不超可用高度",
                       float(fit["after_h"]) <= float(fit["avail_h"]) + 1,
                       f"after={fit['after_h']:.0f} avail={fit['avail_h']}")
                before, after = fit["sizes"][0]
                _check("字号确实变小", after < before, f"{before} -> {after}")
                # tag 字体同步变小
                for widget in app.preview_text_widgets:
                    for tag in widget.tag_names():
                        if not str(tag).startswith("para-style-"):
                            continue
                        size = _font_size_of(widget.tag_cget(tag, "font"))
                        _check(f"{tag} 字体已缩小", 0 < size <= before,
                               f"size={size} before={before}")
                        break
                    else:
                        continue
                    break
                _check("状态栏提示自动缩小", "自动缩小" in app.status_text, app.status_text)

            # 正常小字号不触发
            small = Path(tmp) / "small.docx"
            _make_doc(small, [10.5] * 5, body_chars=8)
            app.docx_handler.load(str(small))
            app._draw_a4_page(900, 600)
            app._refresh_document_view()
            _check("正常字号不触发适配", app._last_preview_fit is None,
                   f"fit={app._last_preview_fit}")
            _check("正常字号无提示", app._preview_fit_note == "")
        finally:
            app.destroy()


def test_preview_fit_extreme() -> None:
    print("[D] 极端内容（更大字号）适配稳定性")
    with tempfile.TemporaryDirectory() as tmp:
        big = Path(tmp) / "huge.docx"
        _make_doc(big, [22.0] * 20, body_chars=20)
        app = FormatApp()
        app.withdraw()
        try:
            app.docx_handler.load(str(big))
            app._draw_a4_page(900, 600)
            app._refresh_document_view()
            fit = app._last_preview_fit
            _check("极端内容仍触发适配", fit is not None)
            if fit:
                _check("极端内容后高度仍不超页",
                       float(fit["after_h"]) <= float(fit["avail_h"]) + 1,
                       f"after={fit['after_h']:.0f} avail={fit['avail_h']}")
                sizes = fit["sizes"]
                _check("所有段落字号被同步缩小",
                       all(a < b for b, a in sizes), f"{sizes[:3]}")
                _check("字号下限保护（>=4px）", all(a >= 4 for _, a in sizes),
                       f"{sizes[:3]}")
        finally:
            app.destroy()


if __name__ == "__main__":
    print("=== 批次7：预览高度自适应 + 字号各自步进 ===")
    test_apply_font_size_step()
    test_ui_step_routing()
    test_preview_height_fit()
    test_preview_fit_extreme()
    print(f"\n批次7结果: {PASS} 通过, {FAIL} 失败")
    sys.exit(1 if FAIL else 0)
