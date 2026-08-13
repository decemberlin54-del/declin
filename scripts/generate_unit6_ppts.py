#!/usr/bin/env python3
"""Generate 5 distinct-style Unit 6 PPTs (4 lessons + showcase)."""

import shutil
from pathlib import Path

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

OUT = Path("/workspace/第六单元PPT")
IMG = OUT / "images"
W, H = Inches(13.333), Inches(7.5)
FL, FT, FW, FH = Inches(0.32), Inches(0.30), Inches(12.68), Inches(6.89)
CX = 6.667


def rgb(c):
    return RGBColor(*c)


def run(r, t, sz=28, bold=False, color=(50, 50, 50)):
    r.text = t
    r.font.size = Pt(sz)
    r.font.bold = bold
    r.font.name = "Microsoft YaHei"
    r.font.color.rgb = rgb(color)


def para(p, align=PP_ALIGN.CENTER, after=8):
    p.alignment = align
    p.space_after = Pt(after)


def new_prs():
    p = Presentation()
    p.slide_width = W
    p.slide_height = H
    return p


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def save(prs, name):
    OUT.mkdir(parents=True, exist_ok=True)
    prs.save(str(OUT / name))
    print(f"✓ {name} ({len(prs.slides)} slides)")


def grad_bg(path, c1, c2):
    IMG.mkdir(parents=True, exist_ok=True)
    if not path.exists():
        im = Image.new("RGB", (1920, 1080), c1)
        d = ImageDraw.Draw(im)
        for y in range(1080):
            t = y / 1080
            c = tuple(int(c1[i] * (1 - t) + c2[i] * t) for i in range(3))
            d.line([(0, y), (1920, y)], fill=c)
        im.save(path)
    return path


# ═══════════════════════════════════════════════════════════
# STYLE A — 论语风：左文右图、大标题、米色底（第21课）
# ═══════════════════════════════════════════════════════════
class LunyuStyle:
    BG1, BG2 = (248, 242, 228), (235, 220, 195)
    INK, ACCENT, GOLD = (60, 45, 30), (140, 60, 40), (180, 130, 60)

    def __init__(self):
        self.bg = grad_bg(IMG / "bg_lunyu.png", self.BG1, self.BG2)

    def base(self, slide):
        slide.shapes.add_picture(str(self.bg), 0, 0, W, H)
        # top ink bar
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, Inches(0.18))
        bar.fill.solid()
        bar.fill.fore_color.rgb = rgb(self.INK)
        bar.line.fill.background()

    def cover(self, prs, title, sub, extra=""):
        s = blank(prs)
        self.base(s)
        # vertical accent strip right
        strip = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(10.5), 0, Inches(2.8), H)
        strip.fill.solid()
        strip.fill.fore_color.rgb = rgb(self.ACCENT)
        strip.line.fill.background()
        # big title left
        box = s.shapes.add_textbox(Inches(1.2), Inches(2.0), Inches(8.5), Inches(2.5))
        p = box.text_frame.paragraphs[0]
        para(p, PP_ALIGN.LEFT)
        run(p.add_run(), title, 56, True, self.INK)
        p2 = box.text_frame.add_paragraph()
        para(p2, PP_ALIGN.LEFT, 14)
        run(p2.add_run(), sub, 28, False, self.ACCENT)
        if extra:
            p3 = box.text_frame.add_paragraph()
            para(p3, PP_ALIGN.LEFT, 14)
            run(p3.add_run(), extra, 22, False, self.GOLD)
        # decorative circle
        c = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(11.0), Inches(2.5), Inches(1.8), Inches(1.8))
        c.fill.solid()
        c.fill.fore_color.rgb = rgb(self.GOLD)
        c.fill.transparency = 0.3
        c.line.color.rgb = rgb(self.INK)

    def knowledge(self, prs, heading, name, items):
        """Like 论语 '孔丘' biography slide: name + content box."""
        s = blank(prs)
        self.base(s)
        # name block
        nb = s.shapes.add_textbox(Inches(1.5), Inches(0.7), Inches(5), Inches(0.8))
        np = nb.text_frame.paragraphs[0]
        para(np, PP_ALIGN.LEFT)
        run(np.add_run(), name, 44, True, self.INK)
        # heading underline
        ul = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.5), Inches(1.45), Inches(4), Inches(0.04))
        ul.fill.solid()
        ul.fill.fore_color.rgb = rgb(self.ACCENT)
        ul.line.fill.background()
        hb = s.shapes.add_textbox(Inches(1.5), Inches(1.55), Inches(6), Inches(0.5))
        hp = hb.text_frame.paragraphs[0]
        para(hp, PP_ALIGN.LEFT)
        run(hp.add_run(), heading, 24, False, self.ACCENT)
        # content panel left
        panel = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.3), Inches(2.2), Inches(7.8), Inches(4.5))
        panel.fill.solid()
        panel.fill.fore_color.rgb = rgb((255, 252, 245))
        panel.line.color.rgb = rgb(self.GOLD)
        panel.line.width = Pt(1.5)
        y = 2.4
        for key, val in items:
            if key:
                kb = s.shapes.add_textbox(Inches(1.6), Inches(y), Inches(2.2), Inches(0.45))
                kp = kb.text_frame.paragraphs[0]
                para(kp, PP_ALIGN.LEFT)
                run(kp.add_run(), f"【{key}】", 22, True, self.ACCENT)
                vb = s.shapes.add_textbox(Inches(3.8), Inches(y), Inches(5.0), Inches(0.7))
            else:
                vb = s.shapes.add_textbox(Inches(1.6), Inches(y), Inches(7.2), Inches(0.7))
            vp = vb.text_frame.paragraphs[0]
            vp.word_wrap = True
            para(vp, PP_ALIGN.LEFT)
            run(vp.add_run(), val, 22, False, self.INK)
            y += 0.85 if len(val) > 30 else 0.65
        # right decorative panel
        rp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(9.5), Inches(2.2), Inches(3.2), Inches(4.5))
        rp.fill.solid()
        rp.fill.fore_color.rgb = rgb(self.ACCENT)
        rp.fill.transparency = 0.15
        rp.line.color.rgb = rgb(self.ACCENT)
        tb = s.shapes.add_textbox(Inches(9.7), Inches(3.5), Inches(2.8), Inches(2))
        tp = tb.text_frame.paragraphs[0]
        para(tp, PP_ALIGN.CENTER)
        run(tp.add_run(), "笔\n记", 48, True, self.ACCENT)

    def twocol(self, prs, title, left_title, left_items, right_title, right_items):
        s = blank(prs)
        self.base(s)
        tb = s.shapes.add_textbox(Inches(1.5), Inches(0.55), Inches(10), Inches(0.7))
        tp = tb.text_frame.paragraphs[0]
        para(tp, PP_ALIGN.CENTER)
        run(tp.add_run(), title, 36, True, self.INK)
        for xt, tt, items in [(0.8, left_title, left_items), (6.8, right_title, right_items)]:
            pnl = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(xt), Inches(1.5), Inches(5.8), Inches(5.2))
            pnl.fill.solid()
            pnl.fill.fore_color.rgb = rgb((255, 252, 245))
            pnl.line.color.rgb = rgb(self.GOLD)
            hb = s.shapes.add_textbox(Inches(xt + 0.2), Inches(1.65), Inches(5.4), Inches(0.5))
            hp = hb.text_frame.paragraphs[0]
            para(hp, PP_ALIGN.CENTER)
            run(hp.add_run(), tt, 28, True, self.ACCENT)
            y = 2.3
            for k, v in items:
                lb = s.shapes.add_textbox(Inches(xt + 0.3), Inches(y), Inches(5.2), Inches(0.9))
                lp = lb.text_frame.paragraphs[0]
                lp.word_wrap = True
                para(lp, PP_ALIGN.LEFT)
                if k:
                    run(lp.add_run(), f"{k}：", 20, True, self.ACCENT)
                    run(lp.add_run(), v, 20, False, self.INK)
                else:
                    run(lp.add_run(), v, 20, False, self.INK)
                y += 0.95

    def activity(self, prs, title, items):
        s = blank(prs)
        self.base(s)
        tb = s.shapes.add_textbox(Inches(1.5), Inches(0.6), Inches(10), Inches(0.7))
        tp = tb.text_frame.paragraphs[0]
        para(tp, PP_ALIGN.LEFT)
        run(tp.add_run(), f"▶ {title}", 32, True, self.ACCENT)
        y = 1.6
        for i, item in enumerate(items, 1):
            circ = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.5), Inches(y), Inches(0.4), Inches(0.4))
            circ.fill.solid()
            circ.fill.fore_color.rgb = rgb(self.ACCENT)
            circ.line.fill.background()
            nb = s.shapes.add_textbox(Inches(1.5), Inches(y + 0.02), Inches(0.4), Inches(0.38))
            np = nb.text_frame.paragraphs[0]
            para(np, PP_ALIGN.CENTER)
            run(np.add_run(), str(i), 14, True, (255, 255, 255))
            bb = s.shapes.add_textbox(Inches(2.1), Inches(y - 0.02), Inches(10), Inches(0.55))
            bp = bb.text_frame.paragraphs[0]
            para(bp, PP_ALIGN.LEFT)
            run(bp.add_run(), item, 24, False, self.INK)
            y += 0.75


# ═══════════════════════════════════════════════════════════
# STYLE B — 济南冬天风：顶栏标题 + 左侧面板（第22课）
# ═══════════════════════════════════════════════════════════
class JinanStyle:
    BG = (240, 248, 255)
    BLUE, DARK, LIGHT = (50, 100, 180), (30, 50, 90), (180, 210, 240)

    def __init__(self):
        self.bg = grad_bg(IMG / "bg_jinan.png", (245, 250, 255), (210, 230, 250))

    def base(self, slide):
        slide.shapes.add_picture(str(self.bg), 0, 0, W, H)

    def cover(self, prs, title, author, sub=""):
        s = blank(prs)
        self.base(s)
        box = s.shapes.add_textbox(Inches(2.5), Inches(2.3), Inches(8.3), Inches(1.2))
        p = box.text_frame.paragraphs[0]
        para(p, PP_ALIGN.CENTER)
        run(p.add_run(), title, 52, True, self.DARK)
        line = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(3.5), Inches(3.5), Inches(6.3), Inches(0.04))
        line.fill.solid()
        line.fill.fore_color.rgb = rgb(self.BLUE)
        line.line.fill.background()
        ab = s.shapes.add_textbox(Inches(8.5), Inches(3.65), Inches(2), Inches(0.5))
        ap = ab.text_frame.paragraphs[0]
        para(ap, PP_ALIGN.RIGHT)
        run(ap.add_run(), author, 28, False, self.BLUE)
        if sub:
            sb = s.shapes.add_textbox(Inches(2.5), Inches(4.3), Inches(8.3), Inches(0.6))
            sp = sb.text_frame.paragraphs[0]
            para(sp, PP_ALIGN.CENTER)
            run(sp.add_run(), sub, 22, False, self.DARK)
        # decorative dots
        for dx, dy, ds in [(2.8, 2.5, 0.15), (2.6, 2.9, 0.1), (3.0, 3.1, 0.2)]:
            d = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(dx), Inches(dy), Inches(ds), Inches(ds))
            d.fill.solid()
            d.fill.fore_color.rgb = rgb(self.LIGHT)
            d.line.fill.background()

    def section(self, prs, header, title, items, logic=False):
        s = blank(prs)
        self.base(s)
        # top header strip
        hdr = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.35), Inches(12.3), Inches(0.55))
        hdr.fill.solid()
        hdr.fill.fore_color.rgb = rgb(self.BLUE)
        hdr.line.fill.background()
        hb = s.shapes.add_textbox(Inches(0.7), Inches(0.38), Inches(12), Inches(0.5))
        hp = hb.text_frame.paragraphs[0]
        para(hp, PP_ALIGN.LEFT)
        run(hp.add_run(), header, 22, True, (255, 255, 255))
        # title with dots
        tb = s.shapes.add_textbox(Inches(1.5), Inches(1.1), Inches(10), Inches(0.7))
        tp = tb.text_frame.paragraphs[0]
        para(tp, PP_ALIGN.CENTER)
        run(tp.add_run(), title, 36, True, self.DARK)
        # left panel
        pw = 7.5 if logic else 11.5
        px = 0.8 if logic else 0.9
        panel = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(px), Inches(2.0), Inches(pw), Inches(4.6))
        panel.fill.solid()
        panel.fill.fore_color.rgb = rgb((255, 255, 255))
        panel.line.color.rgb = rgb(self.LIGHT)
        panel.line.width = Pt(2)
        y = 2.2
        for item in items:
            if isinstance(item, tuple):
                key, val = item
                bb = s.shapes.add_textbox(Inches(px + 0.3), Inches(y), Inches(pw - 0.5), Inches(0.8))
                bp = bb.text_frame.paragraphs[0]
                bp.word_wrap = True
                para(bp, PP_ALIGN.LEFT)
                run(bp.add_run(), f"【{key}】", 24, True, self.BLUE)
                run(bp.add_run(), val, 24, False, self.DARK)
            else:
                bb = s.shapes.add_textbox(Inches(px + 0.3), Inches(y), Inches(pw - 0.5), Inches(0.7))
                bp = bb.text_frame.paragraphs[0]
                bp.word_wrap = True
                para(bp, PP_ALIGN.LEFT)
                run(bp.add_run(), item, 24 if not logic else 28, False, self.DARK)
            y += 0.85 if logic else 0.7
        if logic:
            # right keyword column
            rb = s.shapes.add_textbox(Inches(8.8), Inches(2.2), Inches(3.8), Inches(4.2))
            rp = rb.text_frame.paragraphs[0]
            para(rp, PP_ALIGN.CENTER)
            run(rp.add_run(), "全\n文\n主\n线", 32, True, self.BLUE)

    def activity(self, prs, title, items):
        self.section(prs, "课堂活动", title, items)


# ═══════════════════════════════════════════════════════════
# STYLE C — 古诗四首风：边框 + 居中 + 编号圆（第23课）
# ═══════════════════════════════════════════════════════════
class GushiStyle:
    GREEN, DARK, LIGHT = (80, 130, 90), (40, 70, 50), (180, 210, 180)

    def __init__(self):
        self.bg = grad_bg(IMG / "bg_gushi.png", (235, 245, 232), (200, 225, 200))

    def frame(self, slide):
        slide.shapes.add_picture(str(self.bg), 0, 0, W, H)
        rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, FL, FT, FW, FH)
        rect.fill.background()
        rect.line.color.rgb = rgb(self.GREEN)
        rect.line.width = Pt(2.5)

    def cover(self, prs, chars, sub):
        """Split-character cover like 古诗四首."""
        s = blank(prs)
        self.frame(s)
        cx = 5.5
        for i, ch in enumerate(chars):
            box = s.shapes.add_textbox(Inches(cx + i * 1.3), Inches(1.8 + i * 0.3), Inches(1.5), Inches(1.5))
            p = box.text_frame.paragraphs[0]
            para(p, PP_ALIGN.CENTER)
            run(p.add_run(), ch, 72, True, self.GREEN)
        sb = s.shapes.add_textbox(Inches(3.5), Inches(5.0), Inches(6.3), Inches(0.6))
        sp = sb.text_frame.paragraphs[0]
        para(sp, PP_ALIGN.CENTER)
        run(sp.add_run(), sub, 26, False, self.DARK)

    def titled(self, prs, title, items, numbered=True):
        s = blank(prs)
        self.frame(s)
        tb = s.shapes.add_textbox(Inches(3.5), Inches(0.7), Inches(6.3), Inches(0.8))
        tp = tb.text_frame.paragraphs[0]
        para(tp, PP_ALIGN.CENTER)
        run(tp.add_run(), title, 40, True, self.GREEN)
        y = 1.8
        for i, item in enumerate(items, 1):
            if numbered:
                circ = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.5), Inches(y), Inches(0.45), Inches(0.45))
                circ.fill.solid()
                circ.fill.fore_color.rgb = rgb(self.GREEN)
                circ.line.fill.background()
                nb = s.shapes.add_textbox(Inches(1.5), Inches(y + 0.03), Inches(0.45), Inches(0.42))
                np = nb.text_frame.paragraphs[0]
                para(np, PP_ALIGN.CENTER)
                run(np.add_run(), str(i), 16, True, (255, 255, 255))
                xb = 2.2
            else:
                xb = 1.5
            if isinstance(item, tuple):
                key, val = item[0], item[1]
                text = f"{key}：{val}"
            else:
                key, val = "", item
                text = item
            bb = s.shapes.add_textbox(Inches(xb), Inches(y - 0.02), Inches(10), Inches(0.65))
            bp = bb.text_frame.paragraphs[0]
            bp.word_wrap = True
            para(bp, PP_ALIGN.LEFT)
            if isinstance(item, tuple):
                run(bp.add_run(), f"{item[0]}：", 24, True, self.GREEN)
                run(bp.add_run(), item[1], 24, False, self.DARK)
            else:
                run(bp.add_run(), item, 24, False, self.DARK)
            y += 0.72

    def compare_table(self, prs, title, rows):
        s = blank(prs)
        self.frame(s)
        tb = s.shapes.add_textbox(Inches(2), Inches(0.6), Inches(9.3), Inches(0.7))
        tp = tb.text_frame.paragraphs[0]
        para(tp, PP_ALIGN.CENTER)
        run(tp.add_run(), title, 36, True, self.GREEN)
        cols = [1.2, 4.5, 8.5]
        headers = ["项目", "《风俗通》", "课文"]
        y0 = 1.6
        for j, h in enumerate(headers):
            hb = s.shapes.add_textbox(Inches(cols[j]), Inches(y0), Inches(3.2), Inches(0.45))
            hp = hb.text_frame.paragraphs[0]
            para(hp, PP_ALIGN.CENTER)
            run(hp.add_run(), h, 22, True, self.GREEN)
        y = 2.2
        for row in rows:
            for j, cell in enumerate(row):
                cb = s.shapes.add_textbox(Inches(cols[j]), Inches(y), Inches(3.2), Inches(0.7))
                cp = cb.text_frame.paragraphs[0]
                cp.word_wrap = True
                para(cp, PP_ALIGN.CENTER if j > 0 else PP_ALIGN.LEFT)
                run(cp.add_run(), cell, 20, j == 0, self.DARK)
            line = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), Inches(y + 0.75), Inches(11.3), Inches(0.015))
            line.fill.solid()
            line.fill.fore_color.rgb = rgb(self.LIGHT)
            line.line.fill.background()
            y += 0.85


# ═══════════════════════════════════════════════════════════
# STYLE D — 书卷风：上下分栏 + 文言注解（第24课）
# ═══════════════════════════════════════════════════════════
class ScrollStyle:
    AMBER, BROWN, CREAM = (200, 150, 60), (100, 70, 30), (255, 250, 235)

    def __init__(self):
        self.bg = grad_bg(IMG / "bg_scroll.png", (255, 248, 228), (240, 220, 180))

    def base(self, slide):
        slide.shapes.add_picture(str(self.bg), 0, 0, W, H)

    def cover(self, prs, title, sub):
        s = blank(prs)
        self.base(s)
        # scroll shape
        scroll = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(2), Inches(1.5), Inches(9.3), Inches(4.5))
        scroll.fill.solid()
        scroll.fill.fore_color.rgb = rgb(self.CREAM)
        scroll.line.color.rgb = rgb(self.AMBER)
        scroll.line.width = Pt(3)
        tb = s.shapes.add_textbox(Inches(2.5), Inches(2.5), Inches(8.3), Inches(1.2))
        tp = tb.text_frame.paragraphs[0]
        para(tp, PP_ALIGN.CENTER)
        run(tp.add_run(), title, 48, True, self.BROWN)
        sb = s.shapes.add_textbox(Inches(2.5), Inches(4.0), Inches(8.3), Inches(0.6))
        sp = sb.text_frame.paragraphs[0]
        para(sp, PP_ALIGN.CENTER)
        run(sp.add_run(), sub, 24, False, self.AMBER)

    def grid4(self, prs, title, items):
        """2x2 grid for four fables."""
        s = blank(prs)
        self.base(s)
        tb = s.shapes.add_textbox(Inches(2), Inches(0.5), Inches(9.3), Inches(0.7))
        tp = tb.text_frame.paragraphs[0]
        para(tp, PP_ALIGN.CENTER)
        run(tp.add_run(), title, 36, True, self.BROWN)
        positions = [(0.8, 1.4), (6.8, 1.4), (0.8, 4.0), (6.8, 4.0)]
        for (name, meaning), (x, y) in zip(items, positions):
            card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(5.5), Inches(2.3))
            card.fill.solid()
            card.fill.fore_color.rgb = rgb(self.CREAM)
            card.line.color.rgb = rgb(self.AMBER)
            nb = s.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.2), Inches(5.1), Inches(0.5))
            np = nb.text_frame.paragraphs[0]
            para(np, PP_ALIGN.CENTER)
            run(np.add_run(), name, 26, True, self.AMBER)
            mb = s.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.8), Inches(5.1), Inches(1.3))
            mp = mb.text_frame.paragraphs[0]
            mp.word_wrap = True
            para(mp, PP_ALIGN.CENTER)
            run(mp.add_run(), meaning, 20, False, self.BROWN)

    def classical(self, prs, title, original, notes, translation):
        """Like古诗 annotation slide: original + notes + translation."""
        s = blank(prs)
        self.base(s)
        tb = s.shapes.add_textbox(Inches(2), Inches(0.45), Inches(9.3), Inches(0.6))
        tp = tb.text_frame.paragraphs[0]
        para(tp, PP_ALIGN.CENTER)
        run(tp.add_run(), title, 32, True, self.BROWN)
        # original line
        ob = s.shapes.add_textbox(Inches(1), Inches(1.3), Inches(11.3), Inches(0.8))
        op = ob.text_frame.paragraphs[0]
        para(op, PP_ALIGN.CENTER)
        run(op.add_run(), original, 36, True, self.BROWN)
        # annotation lines
        for i, (word, note) in enumerate(notes):
            x = 1.5 + i * 2.8
            wb = s.shapes.add_textbox(Inches(x), Inches(2.5), Inches(2.5), Inches(0.5))
            wp = wb.text_frame.paragraphs[0]
            para(wp, PP_ALIGN.CENTER)
            run(wp.add_run(), word, 24, True, self.AMBER)
            nb = s.shapes.add_textbox(Inches(x), Inches(3.1), Inches(2.5), Inches(0.6))
            np = nb.text_frame.paragraphs[0]
            np.word_wrap = True
            para(np, PP_ALIGN.CENTER)
            run(np.add_run(), note, 20, False, self.BROWN)
            # connector line
            ln = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(2.3 + i * 0.05), Inches(11), Inches(0.015))
            ln.fill.solid()
            ln.fill.fore_color.rgb = rgb(self.AMBER)
            ln.line.fill.background()
        # translation
        lb = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(4.5), Inches(11.3), Inches(0.04))
        lb.fill.solid()
        lb.fill.fore_color.rgb = rgb(self.AMBER)
        lb.line.fill.background()
        trb = s.shapes.add_textbox(Inches(1.5), Inches(4.8), Inches(10.3), Inches(0.8))
        trp = trb.text_frame.paragraphs[0]
        para(trp, PP_ALIGN.CENTER)
        run(trp.add_run(), translation, 28, False, self.BROWN)

    def notes(self, prs, title, items):
        s = blank(prs)
        self.base(s)
        tb = s.shapes.add_textbox(Inches(1.5), Inches(0.5), Inches(10.3), Inches(0.7))
        tp = tb.text_frame.paragraphs[0]
        para(tp, PP_ALIGN.LEFT)
        run(tp.add_run(), title, 32, True, self.BROWN)
        y = 1.4
        for key, val in items:
            kb = s.shapes.add_textbox(Inches(1.2), Inches(y), Inches(2.5), Inches(0.5))
            kp = kb.text_frame.paragraphs[0]
            para(kp, PP_ALIGN.RIGHT)
            run(kp.add_run(), key, 22, True, self.AMBER)
            vb = s.shapes.add_textbox(Inches(4.0), Inches(y), Inches(8.5), Inches(0.7))
            vp = vb.text_frame.paragraphs[0]
            vp.word_wrap = True
            para(vp, PP_ALIGN.LEFT)
            run(vp.add_run(), val, 22, False, self.BROWN)
            y += 0.8

    def activity(self, prs, title, items):
        self.notes(prs, f"▶ {title}", [(f"{i}", v) for i, v in enumerate(items, 1)])


# ═══════════════════════════════════════════════════════════
# STYLE E — 目录展板式（单元成果展示）
# ═══════════════════════════════════════════════════════════
class ShowcaseStyle:
    PURPLE, PINK, DARK = (120, 70, 150), (230, 100, 150), (50, 30, 70)

    def __init__(self):
        self.bg = grad_bg(IMG / "bg_show.png", (245, 238, 252), (220, 200, 240))

    def frame(self, slide):
        slide.shapes.add_picture(str(self.bg), 0, 0, W, H)
        rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, FL, FT, FW, FH)
        rect.fill.background()
        rect.line.color.rgb = rgb(self.PURPLE)
        rect.line.width = Pt(2)

    def cover(self, prs, title, sub):
        s = blank(prs)
        self.frame(s)
        # big 目/录 style
        for ch, x, sz in [("想", 4.5, 80), ("象", 5.8, 80), ("力", 7.1, 80)]:
            box = s.shapes.add_textbox(Inches(x), Inches(1.5), Inches(1.5), Inches(1.5))
            p = box.text_frame.paragraphs[0]
            para(p, PP_ALIGN.CENTER)
            run(p.add_run(), ch, sz, True, self.PURPLE)
        sb = s.shapes.add_textbox(Inches(3.5), Inches(4.5), Inches(6.3), Inches(0.8))
        sp = sb.text_frame.paragraphs[0]
        para(sp, PP_ALIGN.CENTER)
        run(sp.add_run(), sub, 28, False, self.DARK)
        # directory items
        items = ["《小圣施威降大圣》", "《皇帝的新装》", "《女娲造人》", "《寓言四则》"]
        positions = [(1.8, 5.5), (7.0, 5.5), (1.8, 6.3), (7.0, 6.3)]
        for i, (text, (x, y)) in enumerate(zip(items, positions), 1):
            circ = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(0.45), Inches(0.45))
            circ.fill.solid()
            circ.fill.fore_color.rgb = rgb(self.PINK)
            circ.line.fill.background()
            nb = s.shapes.add_textbox(Inches(x), Inches(y + 0.03), Inches(0.45), Inches(0.42))
            np = nb.text_frame.paragraphs[0]
            para(np, PP_ALIGN.CENTER)
            run(np.add_run(), str(i), 16, True, (255, 255, 255))
            lb = s.shapes.add_textbox(Inches(x + 0.55), Inches(y - 0.02), Inches(5), Inches(0.5))
            lp = lb.text_frame.paragraphs[0]
            para(lp, PP_ALIGN.LEFT)
            run(lp.add_run(), text, 22, True, self.DARK)

    def content(self, prs, title, items):
        s = blank(prs)
        self.frame(s)
        tb = s.shapes.add_textbox(Inches(3), Inches(0.6), Inches(7.3), Inches(0.8))
        tp = tb.text_frame.paragraphs[0]
        para(tp, PP_ALIGN.CENTER)
        run(tp.add_run(), title, 38, True, self.PURPLE)
        y = 1.6
        for i, item in enumerate(items, 1):
            if isinstance(item, tuple):
                key, val = item
                text = f"{key}：{val}"
            else:
                text = item
            circ = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.5), Inches(y), Inches(0.42), Inches(0.42))
            circ.fill.solid()
            circ.fill.fore_color.rgb = rgb(self.PINK)
            circ.line.fill.background()
            nb = s.shapes.add_textbox(Inches(1.5), Inches(y + 0.02), Inches(0.42), Inches(0.38))
            np = nb.text_frame.paragraphs[0]
            para(np, PP_ALIGN.CENTER)
            run(np.add_run(), str(i), 14, True, (255, 255, 255))
            bb = s.shapes.add_textbox(Inches(2.1), Inches(y - 0.02), Inches(10), Inches(0.65))
            bp = bb.text_frame.paragraphs[0]
            bp.word_wrap = True
            para(bp, PP_ALIGN.LEFT)
            if isinstance(item, tuple):
                run(bp.add_run(), f"{item[0]}：", 24, True, self.PURPLE)
                run(bp.add_run(), item[1], 24, False, self.DARK)
            else:
                run(bp.add_run(), item, 24, False, self.DARK)
            y += 0.75


# ═══════════════════════════════════════════════════════════
# BUILD LESSONS
# ═══════════════════════════════════════════════════════════

def build_lesson21():
    s = LunyuStyle()
    prs = new_prs()
    s.cover(prs, "小圣施威降大圣", "第21课 · 吴承恩《西游记》", "母题：自由与规则")
    s.knowledge(prs, "选文背景", "文学常识", [
        ("作者", "吴承恩，明代小说家"),
        ("出处", "章回体长篇小说《西游记》"),
        ("文体", "神魔小说——以神魔斗法写人情世故"),
        ("情节", "悟空大闹天宫后，二郎神奉旨捉拿，二人变化斗法"),
    ])
    s.knowledge(prs, "变化相克", "情节梳理", [
        ("第1轮", "悟空变鸟→二郎变鹰；悟空变鱼→二郎变鱼鹰"),
        ("第2轮", "悟空变水蛇→二郎变灰鹤；悟空变花鸨→二郎现原身"),
        ("第3轮", "悟空变庙宇（尾巴变旗竿）→二郎识破"),
        ("规律", "孙悟空先变，二郎神以相克之物随即相对"),
    ])
    s.twocol(prs, "人物赏析", "孙悟空", [
        ("特点", "机敏善变、好胜不屈"),
        ("证据", "连续变化、不肯认输"),
        ("作用", "体现对自由的追求"),
    ], "二郎神", [
        ("特点", "沉着冷静、维护秩序"),
        ("证据", "紧追不舍、识破破绽"),
        ("作用", "体现规则的力量"),
    ])
    s.knowledge(prs, "名家点评", "主题探究", [
        ("鲁迅", "「使神魔皆有人情，精魅亦通世故」"),
        ("林庚", "不宜看得过于认真，应看到儿童的心理与行为"),
        ("方法", "证据—特点—作用"),
        ("策展主题", "自由可贵，但不可无视规则"),
    ])
    s.activity(prs, "学习任务", [
        "填写七十二变对照表，标出相克关系",
        "完成人物档案卡（证据—特点—作用）",
        "写80字「赛事快讯」，至少三次变化",
    ])
    save(prs, "第21课 小圣施威降大圣.pptx")


def build_lesson22():
    s = JinanStyle()
    prs = new_prs()
    s.cover(prs, "皇 帝 的 新 装", "安徒生", "母题：真实与虚假")
    s.section(prs, "文学常识", "知人论世", [
        ("作者", "安徒生，丹麦童话作家"),
        ("文体", "童话——用虚构故事反映现实生活"),
        ("特点", "以「新装」为线索，写一个荒唐的骗局"),
        ("核心", "为什么人们都不敢说自己看不见？"),
    ])
    s.section(prs, "构建逻辑：梳理情节", "给骗局画流程图", [
        "两个骗子：能织出只有聪明人才能看见的衣服",
        "老大臣、官员、皇帝：先后「看布」，不敢说真话",
        "全城游行：百姓假装看见新装",
        "小孩说真话→百姓跟着说→皇帝仍装模作样",
    ], logic=True)
    s.section(prs, "构建逻辑：品味手法", "拆解心理机关", [
        ("夸张", "把皇帝爱新衣写到极致"),
        ("反讽", "说「看见了」其实什么也没看见"),
        ("对比", "大人的怯懦 vs 孩子的天真诚实"),
        ("反复", "「我什么也没有看见」强化讽刺"),
    ])
    s.section(prs, "构建逻辑：探究主题", "真话为何难以说出口", [
        "害怕被认为「不聪明」或「不称职」",
        "讽刺盲从权威、虚伪逢迎的官场风气",
        "孩子的天真是不被虚假规则束缚的诚实",
        "策展主题：面对虚假，要敢于说真话",
    ])
    s.activity(prs, "学习任务", [
        "画骗局流程图，勾画人物心理",
        "完成人物分析卡（证据—特点—作用）",
        "用150字概括故事，写50字情境运用",
    ])
    save(prs, "第22课 皇帝的新装.pptx")


def build_lesson23():
    s = GushiStyle()
    prs = new_prs()
    s.cover(prs, "女娲造人", "袁珂 · 母题：创造与生命")
    s.titled(prs, "文学常识", [
        ("作者", "袁珂（现代作家，神话研究专家）"),
        ("文体", "神话——用想象解释自然与人类起源"),
        ("特点", "在古籍记载基础上增删改写，赋予现代意识"),
        ("材料", "对比教材「阅读提示」中的《风俗通》"),
    ])
    s.titled(prs, "造人过程", [
        ("起因", "女娲行走世间，感到孤独寂寞"),
        ("方法一", "黄泥和水揉成小泥人，落地即活"),
        ("方法二", "藤条蘸泥浆挥洒，泥点也变成人"),
        ("结果", "建立婚姻制度，让人类繁衍生息"),
    ])
    s.titled(prs, "想象特点", [
        ("具体", "池水照影见自己面容→想到造同类"),
        ("生动", "藤条一挥，满天泥浆洒落"),
        ("温暖", "造人后的疲倦、喜悦，神有人的情感"),
        ("规律", "想象基于生活经验，又超越现实"),
    ])
    s.compare_table(prs, "比较阅读：古籍与课文", [
        ("记载风格", "简洁，重在说明方法", "增加心理、细节、情感"),
        ("造人动机", "未详写", "突出孤独寂寞"),
        ("改写意图", "——", "表达对生命与创造的礼赞"),
    ])
    s.titled(prs, "学习任务", [
        "给女娲行动排序，比较两种造人方法",
        "完成「古籍与课文」比较表",
        "以「我看见女娲……」写100字画面描述",
    ], numbered=True)
    save(prs, "第23课 女娲造人.pptx")


def build_lesson24():
    s = ScrollStyle()
    prs = new_prs()
    s.cover(prs, "寓言四则", "母题：智慧与局限")
    s.notes(prs, "文体知识", [
        ("寓言", "短小故事，寄寓深刻道理，多运用拟人"),
        ("特点", "情节简短、人物典型、寓意明确"),
        ("出处", "《赫耳墨斯和雕像者》《蚊子和狮子》→《伊索寓言》"),
        ("", "《穿井得一人》《杞人忧天》→《吕氏春秋》"),
    ])
    s.grid4(prs, "四则寓意", [
        ("《赫耳墨斯和雕像者》", "讽刺自高自大、妄自尊重的人"),
        ("《蚊子和狮子》", "再小的个体也有长处，骄兵必败"),
        ("《穿井得一人》", "以讹传讹，调查求证才能辨明真相"),
        ("《杞人忧天》", "讽刺不必要的担忧（也可读出忧患意识）"),
    ])
    s.classical(prs, "文言积累",
                "得一人之使，非得一人于井中也。",
                [("闻", "听说"), ("道", "讲述"), ("亡", "同「无」"), ("只使", "纵使")],
                "节省了一个人的劳力，并不是从井里挖出一个人。")
    s.notes(prs, "写法探究", [
        ("情节设计", "寓意藏在欲望、选择与意外后果中"),
        ("改情节", "蚊子平静离开→寓意变为「懂得适可而止」"),
        ("阅读法", "读懂「为什么这样结尾」"),
        ("策展主题", "认识认知局限，善用智慧"),
    ])
    s.activity(prs, "学习任务", [
        "完成「寓言档案卡」，概括四则寓意",
        "解释加点词，翻译重点句",
        "任选一则寓言新编（100—200字）",
    ])
    save(prs, "第24课 寓言四则.pptx")


def build_showcase():
    s = ShowcaseStyle()
    prs = new_prs()
    s.cover(prs, "想象力博物馆", "第七单元学历案 · 成果展示")
    s.content(prs, "单元母题链", [
        ("单元母题", "想象与真实"),
        ("自由与规则", "《小圣施威降大圣》"),
        ("真实与虚假", "《皇帝的新装》"),
        ("创造与生命", "《女娲造人》"),
        ("智慧与局限", "《寓言四则》"),
    ])
    s.content(prs, "四类文本比较", [
        ("神魔小说", "想象依据：相克逻辑；表现：连续变化"),
        ("童话", "想象依据：生活经验；表现：夸张反讽"),
        ("神话", "想象依据：古籍传说；表现：创世画面"),
        ("寓言", "想象依据：生活现象；表现：拟人故事"),
    ])
    s.content(prs, "联想与想象", [
        ("联想", "由一事物想到另一事物（街灯→明星）"),
        ("想象", "在已有材料上创造新形象（天上的街市）"),
        ("写作要求", "有依据、合逻辑、有新意"),
        ("创作路径", "触发点→联想链→情节转折→主题"),
    ])
    s.content(prs, "想象说明卡", [
        ("文本片段", "从本单元选一个精彩片段"),
        ("想象依据", "从什么现实经验或材料出发？"),
        ("表现方式", "用了什么手法？神奇之处在哪？"),
        ("现实意味", "照见了怎样的生活道理？"),
    ])
    s.content(prs, "成果展示要求", [
        "布展：说明卡+创作作品",
        "讲解：30秒依据+60秒亮点+30秒思考",
        "评价：投「发现卡」——最有依据/最有新意",
        "作业：600字想象作文或5分钟课本剧",
    ])
    save(prs, "单元成果展示.pptx")


def cleanup():
    """Remove extra files."""
    keep = {
        "第21课 小圣施威降大圣.pptx",
        "第22课 皇帝的新装.pptx",
        "第23课 女娲造人.pptx",
        "第24课 寓言四则.pptx",
        "单元成果展示.pptx",
    }
    for f in OUT.glob("*.pptx"):
        if f.name not in keep:
            f.unlink()
            print(f"  删除 {f.name}")
    for f in Path("/workspace").glob("unit6-*.pptx"):
        f.unlink()
    zipf = Path("/workspace/unit6-all-ppts.zip")
    if zipf.exists():
        zipf.unlink()


def main():
    build_lesson21()
    build_lesson22()
    build_lesson23()
    build_lesson24()
    build_showcase()
    cleanup()
    # English copies
    copies = {
        "第21课 小圣施威降大圣.pptx": "unit6-lesson21-wukong.pptx",
        "第22课 皇帝的新装.pptx": "unit6-lesson22-emperor.pptx",
        "第23课 女娲造人.pptx": "unit6-lesson23-nuwa.pptx",
        "第24课 寓言四则.pptx": "unit6-lesson24-fables.pptx",
        "单元成果展示.pptx": "unit6-showcase.pptx",
    }
    for cn, en in copies.items():
        shutil.copy2(OUT / cn, Path("/workspace") / en)
    print("\n完成！仅保留5个PPT，每课风格不同。")


if __name__ == "__main__":
    main()
