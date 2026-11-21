# -*- coding: utf-8 -*-
import math
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

FONT = "Microsoft YaHei"
NAVY = RGBColor(0x0B, 0x1F, 0x3A)
TEAL = RGBColor(0x00, 0x8C, 0x95)
ORANGE = RGBColor(0xE0, 0x7A, 0x1F)
RED = RGBColor(0xC0, 0x39, 0x2B)
GREY = RGBColor(0x5F, 0x6B, 0x7A)
LIGHT = RGBColor(0xF2, 0xF5, 0xF8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x1F, 0x29, 0x37)
GREEN = RGBColor(0x2E, 0x8B, 0x57)
PALE = RGBColor(0xBF, 0xD4, 0xE6)
COLORS = {"teal": TEAL, "orange": ORANGE, "red": RED, "green": GREEN, "navy": NAVY, "grey": GREY}


def new_prs():
    p = Presentation()
    p.slide_width = Inches(13.333)
    p.slide_height = Inches(7.5)
    return p


def est_lines(txt, size_pt, width_in):
    usable = max(width_in * 72 - 10, 20)
    w = 0.0
    for ch in txt:
        w += size_pt * (1.0 if ord(ch) > 0x2E80 else 0.55)
    return max(1, math.ceil(w / usable))



import re as _re
_FILEMAP = {
    "UL489_2025_preview": "UL 489（2025）公开预览", "UL891_2019_preview": "UL 891（2019）公开预览",
    "CurrentOS_UL_800VDC_roadmap": "Current/OS × UL 800VDC 路线图资料", "UL840_old_edition_fulltext": "UL 840 第 2 版（1993）公开全文",
    "IEC62368-1_ed4_redline_preview": "IEC 62368-1 Ed.4 红线预览", "UL1973_summary_of_topics": "UL 1973 Summary of Topics",
    "UL1977_text": "UL 1977 公开文本", "UL1973_2022_preview": "UL 1973（2022）公开预览", "UL1449_Intertek_SUN": "Intertek UL 1449 SUN",
    "IEC62368-1_BIS_test_report_template": "IEC 62368-1 BIS 测试报告模板", "UL1741_2025_preview": "UL 1741（2025）公开预览",
    "UL9540A_2025_preview": "UL 9540A（2025）公开预览", "UL9540_2025_preview": "UL 9540（2025）公开预览",
    "UL746C_1995_fulltext": "UL 746C（1995）公开全文", "UL746C_scope": "UL 746C Scope 公开页", "UL1012_2014_fulltext": "UL 1012（2014）公开全文",
    "IEC62368-1_spilma_guide": "IEC 62368-1 SPILMA 指南", "UL_emerging_risks_AI_datacenters": "UL 关于 AI 数据中心新兴风险的文章",
}


def clean(t):
    if not isinstance(t, str):
        return t
    t = _re.sub(r"[（(]\s*agent-tools[^）)]*[）)]", "", t)
    t = _re.sub(r"与\s*agent-tools\s*缓存", "", t)
    t = _re.sub(r"agent-tools\s*[0-9a-f]*", "", t)
    t = _re.sub(r"(与\s*)?/tmp/src\s*(文件与)?", "", t)
    t = _re.sub(r"/tmp/src/", "", t)
    t = _re.sub(r"([A-Za-z0-9\-_]+)\.txt", lambda m: _FILEMAP.get(m.group(1), m.group(1).replace("_", " ")), t)
    t = t.replace("路线图 txt", "路线图资料")
    return t


def clean_obj(o):
    if isinstance(o, str):
        return clean(o)
    if isinstance(o, list):
        return [clean_obj(x) for x in o]
    if isinstance(o, dict):
        return {k: clean_obj(v) for k, v in o.items()}
    return o


WARN = []
import json as _json, glob as _glob, os as _os
DETAIL = {}
CARDX = {}
EXPORT = []
EXPORT_MODE = _os.environ.get("EXPORT") == "1"


def load_expansions(pattern="/tmp/deck/work/out/*.json"):
    for fp in sorted(_glob.glob(pattern)):
        try:
            arr = _json.load(open(fp, encoding="utf-8"))
        except Exception as e:
            print("BAD JSON", fp, e)
            continue
        for e in arr:
            if "slides" in e:
                CARDX[e["k"]] = e["slides"]
            else:
                DETAIL[e["k"]] = e



class Deck:
    def __init__(self, footer_label=""):
        self.prs = new_prs()
        self.n = 0
        self.footer_label = footer_label
        self.cur_key = None
        self.cur_meta = {}
        self.cur_rows = None
        self.cur_head = None
        self._busy = False
        self.expand = True

    # ---- primitives
    def blank(self):
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.n += 1
        return s

    def rect(self, s, x, y, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE):
        r = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
        r.fill.solid()
        r.fill.fore_color.rgb = fill
        if line is None:
            r.line.fill.background()
        else:
            r.line.color.rgb = line
            r.line.width = Pt(0.75)
        r.shadow.inherit = False
        return r

    def text(self, s, x, y, w, h, lines, size=12, color=DARK, bold=False, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, gap=0):
        tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.05)
        tf.margin_top = tf.margin_bottom = Inches(0.03)
        tf.vertical_anchor = anchor
        if isinstance(lines, str):
            lines = [lines]
        for i, ln in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            if gap:
                p.space_after = Pt(gap)
            r = p.add_run()
            r.text = ln
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.color.rgb = color
            r.font.name = FONT
        return tb

    def bullets(self, s, x, y, w, h, items, size=11, color=DARK, gap=4):
        tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.08)
        tf.margin_top = Inches(0.04)
        for i, it in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.space_after = Pt(gap)
            r = p.add_run()
            r.text = "\u2022 " + it
            r.font.size = Pt(size)
            r.font.color.rgb = color
            r.font.name = FONT
        return tb

    # ---- chrome
    def header(self, s, title, subtitle=None, tag=None, tag_color=TEAL):
        if not self._busy:
            self.cur_meta.update({'title': title, 'subtitle': subtitle, 'tag': tag, 'tag_color': tag_color})
            if self.cur_key is None:
                self.cur_key = 'main|' + title
        self.rect(s, 0, 0, 13.333, 1.0, NAVY)
        self.rect(s, 0, 1.0, 13.333, 0.05, TEAL)
        tsize = 26 if len(title) <= 26 else (22 if len(title) <= 34 else 19)
        self.text(s, 0.45, 0.08, 10.4, 0.55, title, size=tsize, color=WHITE, bold=True)
        if subtitle:
            self.text(s, 0.45, 0.6, 10.5, 0.35, subtitle, size=12.5, color=PALE)
        if tag:
            c = self.rect(s, 10.9, 0.28, 2.1, 0.42, tag_color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            tf = c.text_frame
            tf.margin_left = tf.margin_right = Inches(0.04)
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = tag
            r.font.size = Pt(12)
            r.font.bold = True
            r.font.color.rgb = WHITE
            r.font.name = FONT

    def footer(self, s, src=None, conf=None):
        self._footer(s, src, conf)
        self._after_slide(src, conf)

    def _footer(self, s, src=None, conf=None):
        self.rect(s, 0, 7.08, 13.333, 0.42, LIGHT)
        label = ("来源：" + src) if src else self.footer_label
        self.text(s, 0.4, 7.1, 10.5, 0.36, label[:150], size=9, color=GREY, anchor=MSO_ANCHOR.MIDDLE)
        if conf:
            col = GREEN if str(conf).startswith("A") else ORANGE
            c = self.rect(s, 11.1, 7.14, 1.75, 0.28, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            tf = c.text_frame
            tf.margin_top = tf.margin_bottom = 0
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = "引用核实 " + str(conf)
            r.font.size = Pt(9)
            r.font.bold = True
            r.font.color.rgb = WHITE
            r.font.name = FONT
        self.text(s, 12.8, 7.1, 0.5, 0.36, str(self.n), size=9, color=GREY,
                  anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.RIGHT)

    # ---- table
    def table(self, s, x, y, w, colw, head, rows, size=12, max_h=5.7, min_size=9.5, row_pad=0.2):
        if self.cur_rows is None:
            self.cur_rows = rows
            self.cur_head = head
        tot = float(sum(colw))
        widths = [w * c / tot for c in colw]

        def total_h(sz):
            h = 0.0
            hl = max(est_lines(t, sz, widths[i]) for i, t in enumerate(head))
            h += hl * sz * 1.25 / 72 + row_pad
            for r in rows:
                ml = max(est_lines(str(t), sz, widths[i]) for i, t in enumerate(r))
                h += ml * sz * 1.25 / 72 + row_pad
            return h

        sz = min(size + 2.5, 15.5)
        while total_h(sz) > max_h * 0.93 and sz > min_size:
            sz -= 0.5
        if total_h(sz) > max_h * 1.0:
            WARN.append(('table', self.n, round(total_h(sz), 2), max_h, sz))
        head_sz = min(sz, 12.5)
        nrows = len(rows) + 1
        gs = s.shapes.add_table(nrows, len(head), Inches(x), Inches(y), Inches(w), Inches(0.3 * nrows))
        tbl = gs.table
        tblPr = gs._element.graphic.graphicData.tbl.tblPr
        tblPr.set("firstRow", "0")
        tblPr.set("bandRow", "0")
        for i, cw in enumerate(widths):
            tbl.columns[i].width = Inches(cw)
        extra = 0.0
        for r in range(nrows):
            tbl.rows[r].height = Inches(0.3)

        def setcell(cell, txt, sz_, color, bold, fill):
            cell.fill.solid()
            cell.fill.fore_color.rgb = fill
            cell.margin_left = cell.margin_right = Inches(0.07)
            cell.margin_top = cell.margin_bottom = Inches(row_pad / 2 - 0.02)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            r = p.add_run()
            r.text = str(txt)
            r.font.size = Pt(sz_)
            r.font.bold = bold
            r.font.color.rgb = color
            r.font.name = FONT

        for c, h in enumerate(head):
            setcell(tbl.cell(0, c), h, head_sz, WHITE, True, TEAL)
        for r, row in enumerate(rows, start=1):
            fill = WHITE if r % 2 else LIGHT
            for c, v in enumerate(row):
                col = DARK
                bold = False
                if c == 0:
                    col, bold = NAVY, True
                if str(v) in ("原文", "目录", "解读") and c == len(row) - 1:
                    col = {"原文": GREEN, "目录": TEAL, "解读": ORANGE}[str(v)]
                    bold = True
                setcell(tbl.cell(r, c), v, sz, col, bold, fill)
        return sz

    # ---- cards
    def card(self, s, x, y, w, h, title, items, accent=TEAL, size=11.5, min_size=9):
        self.rect(s, x, y, w, h, WHITE, line=RGBColor(0xD5, 0xDC, 0xE4))
        self.rect(s, x, y, 0.07, h, accent)
        self.text(s, x + 0.15, y + 0.05, w - 0.25, 0.32, title, size=13, color=accent, bold=True)
        inner_w = w - 0.4

        def need(sz):
            t = 0.0
            for it in items:
                t += est_lines("\u2022 " + it, sz, inner_w) * sz * 1.3 / 72 + 4 / 72
            return t

        sz = min(size + 2.5, 16)
        while need(sz) > h - 0.55 and sz > min_size:
            sz -= 0.5
        if need(sz) > h - 0.45:
            WARN.append(('card', self.n, title, round(need(sz), 2), h))
        self.bullets(s, x + 0.12, y + 0.4, w - 0.25, h - 0.45, items, size=sz)


    # ---- expansion machinery
    def _after_slide(self, src, conf):
        if self._busy:
            return
        key = self.cur_key
        meta = self.cur_meta
        rows = self.cur_rows
        head = self.cur_head
        self.cur_key = None
        self.cur_rows = None
        self.cur_head = None
        self.cur_meta = {}
        if not key:
            return
        if EXPORT_MODE:
            EXPORT.append({"k": key, "title": meta.get("title"), "head": head, "rows": rows, "src": src, "conf": conf,
                           "cards": meta.get("cards"), "impact": meta.get("impact"), "advice": meta.get("advice")})
        if not self.expand:
            return
        self._busy = True
        try:
            if rows:
                ents = []
                for ri, r in enumerate(rows):
                    e = DETAIL.get(f"{key}|{ri}")
                    if e:
                        ents.append((r, e))
                if ents:
                    self.detail_slides(meta, ents, src, conf)
            for sd in CARDX.get(key, []):
                self.std_slide(dict(sd, src=sd.get("src") or src, conf=sd.get("conf") or conf), meta.get("tag"), meta.get("tag_color", TEAL), _nokey=True)
        finally:
            self._busy = False
            self.cur_rows = None
            self.cur_head = None

    def detail_slides(self, meta, ents, src, conf):
        per = 2
        chunks = [ents[i:i + per] for i in range(0, len(ents), per)]
        base = meta.get("title", "")
        for ci, ch in enumerate(chunks):
            s = self.blank()
            title = f"{base} · 逐条详解 第{ci + 1}页/共{len(chunks)}页"
            self.header(s, title, meta.get("subtitle") or "逐条说明：要求 / 800V 意义 / 验证方式 / 注意点与失败点 / 引用", meta.get("tag"), meta.get("tag_color", TEAL))
            bh = 2.78
            for ei, (row, e) in enumerate(ch):
                y = 1.2 + ei * (bh + 0.1)
                self.detail_entry(s, 0.35, y, 12.65, bh, row, e)
            self._footer(s, src, conf)

    def detail_entry(self, s, x, y, w, h, row, e):
        self.rect(s, x, y, w, h, WHITE, line=RGBColor(0xD5, 0xDC, 0xE4))
        basis = str(row[-1]) if str(row[-1]) in ("原文", "目录", "解读") else ""
        ref = str(row[0])
        body = str(row[1]) if len(row) > 2 else " ".join(str(c) for c in row[:-1])
        band = self.rect(s, x, y, w, 0.5, NAVY)
        head = f"{ref}  |  {body}"
        hs = 12.5
        while est_lines(head, hs, w - 1.2) > 2 and hs > 9.5:
            hs -= 0.5
        self.text(s, x + 0.08, y + 0.02, w - 1.1, 0.46, head, size=hs, color=WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
        if basis:
            col = {"原文": GREEN, "目录": TEAL, "解读": ORANGE}[basis]
            b = self.rect(s, x + w - 0.95, y + 0.1, 0.85, 0.3, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            tf = b.text_frame
            tf.margin_top = tf.margin_bottom = 0
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = basis
            r.font.size = Pt(11)
            r.font.bold = True
            r.font.color.rgb = WHITE
            r.font.name = FONT
        fields = [("具体要求", e.get("req", ""), TEAL), ("对 800V DC 的意义", e.get("why", ""), ORANGE),
                  ("怎么测 / 怎么证明", e.get("verify", ""), GREEN), ("注意点与常见失败", e.get("pitfall", ""), RED)]
        cw = (w - 0.45) / 2
        fh = (h - 0.5 - 0.36 - 0.1) / 2
        texts = [f[1] for f in fields]
        sz = 13.5
        while sz > 9 and max(est_lines(t, sz, cw - 0.1) for t in texts) * sz * 1.28 / 72 > fh - 0.26:
            sz -= 0.5
        for i, (lab, txt, col) in enumerate(fields):
            cx = x + 0.15 + (i % 2) * (cw + 0.15)
            cy = y + 0.55 + (i // 2) * fh
            self.rect(s, cx, cy + 0.03, 0.05, fh - 0.08, col)
            self.text(s, cx + 0.08, cy, cw - 0.1, 0.24, lab, size=10.5, color=col, bold=True)
            self.text(s, cx + 0.08, cy + 0.22, cw - 0.12, fh - 0.25, txt, size=sz, color=DARK)
        cite = "引用：" + str(e.get("cite", ""))
        cs = 10
        while est_lines(cite, cs, w - 0.3) > 1 and cs > 8:
            cs -= 0.5
        self.text(s, x + 0.12, y + h - 0.34, w - 0.24, 0.3, cite, size=cs, color=GREY, anchor=MSO_ANCHOR.MIDDLE)

    # ---- composite slide kinds
    def std_slide(self, d, tag, tag_color=TEAL, _nokey=False):
        d = clean_obj(d)
        if not _nokey:
            self.cur_key = d.get('_k')
            self.cur_rows = None
            self.cur_meta = {'cards': d.get('cards'), 'impact': d.get('impact'), 'advice': d.get('advice')}
        k = d.get("kind", "table")
        s = self.blank()
        self.header(s, d["title"], d.get("subtitle"), tag, tag_color)
        if k == "table":
            has_side = bool(d.get("impact") or d.get("advice"))
            if has_side:
                self.table(s, 0.35, 1.25, 8.1, d["colw"], d["head"], d["rows"], size=12.5, max_h=5.65)
                if d.get("impact") and d.get("advice"):
                    self.card(s, 8.7, 1.25, 4.3, 3.1, d.get("impact_title", "对 800V 设计的影响"), d["impact"], ORANGE, 11.5)
                    self.card(s, 8.7, 4.45, 4.3, 2.5, d.get("advice_title", "取证 / 设计建议"), d["advice"], TEAL, 11)
                elif d.get("impact"):
                    self.card(s, 8.7, 1.25, 4.3, 5.7, d.get("impact_title", "对 800V 设计的影响"), d["impact"], ORANGE, 12)
                else:
                    self.card(s, 8.7, 1.25, 4.3, 5.7, d.get("advice_title", "取证 / 设计建议"), d["advice"], TEAL, 12)
            else:
                self.table(s, 0.35, 1.25, 12.65, d["colw"], d["head"], d["rows"], size=12.5, max_h=5.65)
        elif k == "table_full":
            self.table(s, 0.35, 1.25, 12.65, d["colw"], d["head"], d["rows"], size=12.5, max_h=5.65)
        elif k == "cards":
            cards = d["cards"]
            n = len(cards)
            if n <= 2:
                w = (12.65 - 0.2 * (n - 1)) / n
                for i, c in enumerate(cards):
                    self.card(s, 0.35 + i * (w + 0.2), 1.25, w, 5.7, c["title"], c["items"],
                              COLORS.get(c.get("color", "teal"), TEAL), 13)
            elif n == 3:
                w = (12.65 - 0.4) / 3
                for i, c in enumerate(cards):
                    self.card(s, 0.35 + i * (w + 0.2), 1.25, w, 5.7, c["title"], c["items"],
                              COLORS.get(c.get("color", "teal"), TEAL), 12.5)
            else:
                w = (12.65 - 0.2) / 2
                h = (5.7 - 0.2) / 2
                for i, c in enumerate(cards[:4]):
                    self.card(s, 0.35 + (i % 2) * (w + 0.2), 1.25 + (i // 2) * (h + 0.2), w, h, c["title"],
                              c["items"], COLORS.get(c.get("color", "teal"), TEAL), 12)
        self.footer(s, d.get("src"), d.get("conf"))
        return s

    def divider(self, title, sub="", num=""):
        s = self.blank()
        self.rect(s, 0, 0, 13.333, 7.5, NAVY)
        self.rect(s, 0.8, 3.55, 1.6, 0.06, TEAL)
        if num:
            self.text(s, 0.8, 1.6, 4, 1.0, num, size=54, color=TEAL, bold=True)
        self.text(s, 0.8, 2.55, 11.5, 0.9, title, size=36, color=WHITE, bold=True)
        if sub:
            self.text(s, 0.8, 3.8, 11.5, 1.5, sub if isinstance(sub, list) else [sub], size=16, color=PALE, gap=6)
        self.text(s, 12.3, 7.0, 0.8, 0.4, str(self.n), size=10, color=PALE, align=PP_ALIGN.RIGHT)
        return s

    def flow(self, title, subtitle, steps, tag=None, src=None, conf=None, tag_color=TEAL):
        """steps: list of (label, desc) rendered as horizontal chevrons with descriptions below"""
        s = self.blank()
        self.header(s, title, subtitle, tag, tag_color)
        n = len(steps)
        per_row = 4 if n > 4 else n
        rows_n = math.ceil(n / per_row)
        w = (12.65 - 0.15 * (per_row - 1)) / per_row
        row_h = 5.7 / rows_n
        for i, (lab, desc) in enumerate(steps):
            r, c = divmod(i, per_row)
            x = 0.35 + c * (w + 0.15)
            y = 1.25 + r * row_h
            ch = self.rect(s, x, y, w, 0.62, TEAL if i % 2 == 0 else NAVY, shape=MSO_SHAPE.PENTAGON)
            tf = ch.text_frame
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.margin_left = Inches(0.12)
            p = tf.paragraphs[0]
            rr = p.add_run()
            rr.text = f"{i + 1}  {lab}"
            rr.font.size = Pt(14)
            rr.font.bold = True
            rr.font.color.rgb = WHITE
            rr.font.name = FONT
            self.rect(s, x, y + 0.7, w, row_h - 0.85, WHITE, line=RGBColor(0xD5, 0xDC, 0xE4))
            sz = 15
            while sum(est_lines("\u2022 " + t, sz, w - 0.3) * sz * 1.3 / 72 + 4 / 72 for t in desc) > row_h - 1.0 and sz > 9:
                sz -= 0.5
            self.bullets(s, x + 0.05, y + 0.75, w - 0.1, row_h - 0.95, desc, size=sz)
        self.footer(s, src, conf)
        return s
