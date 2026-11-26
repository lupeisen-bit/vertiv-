import math
import os
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

sys.path.insert(0, os.path.dirname(__file__))
from data import GROUPS, STANDARDS  # noqa: E402

FONT = "Microsoft YaHei"
TEAL = RGBColor(0x00, 0x8C, 0x95)
DARK = RGBColor(0x1F, 0x2D, 0x3D)
LIGHT = RGBColor(0xEA, 0xF4, 0xF5)
GRAY = RGBColor(0x59, 0x65, 0x70)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ORANGE = RGBColor(0xE0, 0x7B, 0x1F)
CONF_COLOR = {"高": RGBColor(0x2E, 0x8B, 0x57), "中高": RGBColor(0x5B, 0x9B, 0x3A),
              "中": RGBColor(0xD9, 0x9A, 0x1E), "低": RGBColor(0xC0, 0x45, 0x3A)}

SW, SH = 13.333, 7.5
COLW = [2.35, 5.0, 5.2]
TBL_FONT = 12
HEAD_H = 0.38


def est_lines(text, width_in, pt):
    em = 0.0
    for ch in text:
        em += 1.0 if ord(ch) > 0x2E80 else 0.55
    usable = (width_in - 0.2) / (pt / 72.0) * 0.93
    return max(1, math.ceil(em / usable))


def row_height(row):
    lines = max(est_lines(t, w, TBL_FONT) for t, w in zip(row, COLW))
    return lines * TBL_FONT * 1.3 / 72 + 0.16


def set_text(tf, text, size, color=DARK, bold=False, align=None):
    tf.clear()
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.name = FONT
    r.font.color.rgb = color
    if align:
        p.alignment = align


def rect(slide, x, y, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
    s.shadow.inherit = False
    return s


def textbox(slide, x, y, w, h, text, size, color=DARK, bold=False, align=None, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.08)
    tf.margin_top = tf.margin_bottom = Inches(0.04)
    set_text(tf, text, size, color, bold, align)
    return tb


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def header(slide, title, tag=None, conf=None):
    rect(slide, 0, 0, SW, 0.85, TEAL)
    em = sum(1.0 if ord(ch) > 0x2E80 else 0.58 for ch in title)
    size = max(14, min(22, int(8.2 * 72 / em))) if tag else 22
    textbox(slide, 0.4, 0.1, 8.5 if tag else 12.5, 0.65, title, size, WHITE, True, anchor=MSO_ANCHOR.MIDDLE)
    x = SW - 0.4
    if conf:
        w = 1.45
        x -= w
        c = rect(slide, x, 0.22, w, 0.4, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        set_text(c.text_frame, "置信度：" + conf, 12, CONF_COLOR[conf], True, PP_ALIGN.CENTER)
        c.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        x -= 0.15
    if tag:
        w = 2.45
        x -= w
        c = rect(slide, x, 0.22, w, 0.4, LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        set_text(c.text_frame, tag, 11, TEAL, True, PP_ALIGN.CENTER)
        c.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE


def footer(slide, n):
    textbox(slide, SW - 1.6, SH - 0.35, 1.3, 0.3, str(n), 10, GRAY, align=PP_ALIGN.RIGHT)
    textbox(slide, 0.4, SH - 0.35, 8, 0.3, "UL 标准章节说明 · 内容依据公开资料整理，条款号与版本以正式文本为准", 9, GRAY)


def add_table(slide, y, rows, head=("章节 / 条目", "内容要点", "说明（为什么重要 / 怎么用）")):
    heights = [row_height(r) for r in rows]
    shape = slide.shapes.add_table(len(rows) + 1, 3, Inches(0.4), Inches(y), Inches(sum(COLW)),
                                   Inches(HEAD_H + sum(heights)))
    tbl = shape.table
    tblPr = shape._element.graphic.graphicData.tbl.tblPr
    tblPr.set("firstRow", "0")
    tblPr.set("bandRow", "0")
    for i, w in enumerate(COLW):
        tbl.columns[i].width = Inches(w)
    tbl.rows[0].height = Inches(HEAD_H)
    for j, h in enumerate(head):
        c = tbl.cell(0, j)
        c.fill.solid()
        c.fill.fore_color.rgb = TEAL
        set_text(c.text_frame, h, TBL_FONT, WHITE, True)
        c.vertical_anchor = MSO_ANCHOR.MIDDLE
    for i, row in enumerate(rows, start=1):
        tbl.rows[i].height = Inches(heights[i - 1])
        for j, t in enumerate(row):
            c = tbl.cell(i, j)
            c.fill.solid()
            c.fill.fore_color.rgb = LIGHT if i % 2 == 0 else WHITE
            set_text(c.text_frame, t, TBL_FONT, DARK, bold=(j == 0))
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            c.margin_left = c.margin_right = Inches(0.08)
            c.margin_top = c.margin_bottom = Inches(0.03)
    return sum(heights) + HEAD_H


def paginate(rows, first_top, cont_top, tip_h):
    bottom = SH - 0.5
    pages, cur, used = [], [], 0.0
    top = first_top
    for r in rows:
        h = row_height(r)
        if cur and top + HEAD_H + used + h > bottom:
            pages.append(cur)
            cur, used, top = [], 0.0, cont_top
        cur.append(r)
        used += h
    pages.append(cur)
    if top + HEAD_H + used + tip_h > bottom and len(pages[-1]) > 1:
        moved = pages[-1].pop()
        pages.append([moved])
    return pages


def build():
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)
    n = [0]

    def new():
        n[0] += 1
        return blank(prs)

    s = new()
    rect(s, 0, 0, SW, SH, TEAL)
    textbox(s, 0.9, 2.2, 11.5, 1.2, "800VDC 数据中心 UL 标准", 40, WHITE, True)
    textbox(s, 0.9, 3.35, 11.5, 0.9, "23 项标准的章节内容与逐节说明", 28, WHITE)
    textbox(s, 0.9, 4.6, 11.5, 0.9,
            "分组：路线图与总览 · 电源变换与 ICT 安全 · 保护、配电与预制系统 · 储能与备用电源 · 连接、浪涌、绝缘与材料",
            14, LIGHT)

    s = new()
    header(s, "阅读说明")
    items = [
        "1. 内容依据公开资料与行业通用结构整理，没有逐条对照正式文本；章节编号和条款号会随版本变化，请以购买的正式文本为准。",
        "2. 每项标准一节：先说明范围与定位，再逐章（或逐组条款）给出「内容要点」与「说明」，最后给出 800VDC 应用提示。",
        "3. 置信度表示章节结构的把握程度：高 = 章节结构明确；中高 = 主体明确，细节随版本变化；中 = 大框架正确，细分章节按同类标准惯例归纳；低 = OOI 调查大纲，公开信息少，结构为推断。",
        "4. UL 标准通用骨架：前言 → 引言（范围、术语、元件、单位、引用）→ 结构要求 Construction → 性能要求 Performance → 制造和生产试验 → 标志 Markings → 附录。",
        "5. UL 489 的 G/H/F/I 增补、UL 2877 / 2140 / 2755 的具体章节，请向 UL Solutions 获取目录页后再核对。",
        "6. 仓库中已有 ul-certification 目录下的深度手册（含条款号、依据标签），本套为章节级的轻量版说明，可配合使用。",
    ]
    tb = s.shapes.add_textbox(Inches(0.6), Inches(1.2), Inches(12.1), Inches(5.6))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, t in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r = p.add_run()
        r.text = t
        r.font.size = Pt(16)
        r.font.name = FONT
        r.font.color.rgb = DARK
        p.space_after = Pt(12)
    footer(s, n[0])

    s = new()
    header(s, "标准总览（共 23 项，5 个分组）")
    rows = []
    for g, lst in GROUPS:
        rows.append((g, "、".join(lst), ""))
    by = {st["id"]: st for st in STANDARDS}
    y = 0.98
    for g, lst in GROUPS:
        rect(s, 0.4, y, SW - 0.8, 0.32, TEAL)
        textbox(s, 0.5, y, 6, 0.32, g, 13, WHITE, True, anchor=MSO_ANCHOR.MIDDLE)
        y += 0.38
        for i in range(0, len(lst), 6):
            for k, sid in enumerate(lst[i:i + 6]):
                st = by[sid]
                x = 0.4 + k * 2.1
                short = st["name"].split("（")[0]
                if len(short) > 15:
                    short = short[:14] + "…"
                rect(s, x, y, 2.0, 0.6, LIGHT)
                textbox(s, x, y, 2.0, 0.3, sid, 12, TEAL, True)
                textbox(s, x, y + 0.28, 2.0, 0.32, short, 9, GRAY)
            y += 0.64
        y += 0.04
    footer(s, n[0])

    group_seen = set()
    for st in STANDARDS:
        if st["group"] not in group_seen:
            group_seen.add(st["group"])
            s = new()
            rect(s, 0, 0, SW, SH, DARK)
            textbox(s, 0.9, 2.6, 11.5, 1.0, st["group"], 36, WHITE, True)
            members = [x["id"] for x in STANDARDS if x["group"] == st["group"]]
            textbox(s, 0.9, 3.8, 11.5, 0.9, "、".join(members), 18, LIGHT)

        scope_h = 0.2 + est_lines(st["scope"], SW - 0.8, 12) * 12 * 1.35 / 72
        tip_h = 0.2 + est_lines(st["tip"], SW - 1.0, 12) * 12 * 1.35 / 72 + 0.15
        first_top = 0.95 + scope_h + 0.1
        pages = paginate(st["rows"], first_top, 1.0, tip_h)
        for pi, rows in enumerate(pages):
            s = new()
            title = f'{st["id"]}  {st["name"]}'
            if len(pages) > 1:
                title += f"（{pi + 1}/{len(pages)}）"
            header(s, title, st["group"], st["conf"])
            top = 1.0
            if pi == 0:
                rect(s, 0.4, 0.95, SW - 0.8, scope_h, LIGHT)
                textbox(s, 0.5, 0.95, SW - 1.0, scope_h, "范围与定位：" + st["scope"], 12, DARK,
                        anchor=MSO_ANCHOR.MIDDLE)
                top = first_top
            used = add_table(s, top, rows)
            if pi == len(pages) - 1:
                ty = top + used + 0.15
                rect(s, 0.4, ty, 0.08, tip_h - 0.15, ORANGE)
                textbox(s, 0.55, ty - 0.02, SW - 1.0, tip_h - 0.1, "800VDC 应用提示：" + st["tip"], 12, DARK,
                        anchor=MSO_ANCHOR.MIDDLE)
            footer(s, n[0])

    s = new()
    header(s, "使用建议与待核对事项")
    items = [
        "基础链路：UL 62368-1（ICT 设备）、UL 840（绝缘配合）、UL 94 / UL 746C（材料）是几乎所有设备的基础，其余标准针对特定设备。",
        "800VDC 的主要缺口：ES3 防护、直流分断与电弧、间距与爬电、预制系统、中压电源，现有标准常需要 CRD、OOI 或补充评估。",
        "系统层级：BESS 需要 UL 1973（电池）+ UL 1741（PCS）+ UL 9540（系统）+ UL 9540A（火灾传播数据）组合使用。",
        "待核对：UL 489 增补字母的对应关系；UL 2877 / 2140 / 2755 的正式章节与完整名称（原图被截断）；各标准当前有效版本。",
        "建议：向 UL Solutions 获取目录页与 CRD 清单，并在项目早期与认证机构确认 800VDC 的评估路径。",
    ]
    tb = s.shapes.add_textbox(Inches(0.6), Inches(1.2), Inches(12.1), Inches(5.6))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, t in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r = p.add_run()
        r.text = t
        r.font.size = Pt(16)
        r.font.name = FONT
        r.font.color.rgb = DARK
        p.space_after = Pt(14)
    footer(s, n[0])

    out = os.path.join(os.path.dirname(__file__), "UL标准章节内容说明.pptx")
    prs.save(out)
    print("saved", out, "slides:", n[0])


if __name__ == "__main__":
    build()
