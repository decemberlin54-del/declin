#!/usr/bin/env python3
"""第23课《女娲造人》新国风水墨青绿风格 PPT — 7 slides."""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

OUT = Path("/workspace/第六单元PPT/第23课 女娲造人.pptx")
OUT_EN = Path("/workspace/unit6-lesson23-nuwa.pptx")

SW = Inches(13.333)
SH = Inches(7.5)
PL, PT = Inches(0.6), Inches(0.38)
IW = SW - Inches(1.2)

# 水墨青绿 + 古风暖浅灰
BG = RGBColor(0xF4, 0xF1, 0xEA)       # 宣纸暖浅灰
GREEN = RGBColor(0x4A, 0x7C, 0x6F)   # 水墨青绿
GREEN_SOFT = RGBColor(0xE4, 0xED, 0xE8)
GREEN_WASH = RGBColor(0xD8, 0xE6, 0xDF)  # 淡青晕染
INK = RGBColor(0x1E, 0x1E, 0x1E)
RED = RGBColor(0xC8, 0x3C, 0x23)
MUTED = RGBColor(0x5A, 0x6A, 0x62)

FONT_TITLE = "KaiTi"
FONT_BODY = "Microsoft YaHei"


def font(run, size=16, bold=False, color=INK, title=False):
    run.font.name = FONT_TITLE if title else FONT_BODY
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color


def bg(slide):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    sh.fill.solid()
    sh.fill.fore_color.rgb = BG
    sh.line.fill.background()
    tree = slide.shapes._spTree
    tree.remove(sh._element)
    tree.insert(2, sh._element)


def wash(slide, l, t, w, h, color=GREEN_WASH):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = color
    sh.line.fill.background()
    return sh


def line_box(slide, l, t, w, h, fill=BG, border=GREEN, lw=Pt(1.5)):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    sh.line.color.rgb = border
    sh.line.width = lw
    return sh


def cloud_div(slide, y):
    ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, PL, y, IW, Inches(0.01))
    ln.fill.solid()
    ln.fill.fore_color.rgb = MUTED
    ln.line.fill.background()
    tb(slide, PL, y - Inches(0.07), IW, Inches(0.22), [("～ ～ ～ ～ ～ ～ ～ ～ ～ ～ ～ ～", 9, False, MUTED)], PP_ALIGN.CENTER)


def seal(slide, l, t, size, text, fs=11):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, size, size)
    sh.fill.background()
    sh.line.color.rgb = RED
    sh.line.width = Pt(2)
    tb(slide, l, t, size, size, [(text, fs, True, RED)], PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)


def green_bar(slide, l, t, h, w=Inches(0.05)):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = GREEN
    sh.line.fill.background()


def tb(slide, l, t, w, h, lines, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, title=False):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for i, item in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if isinstance(item, tuple):
            txt, sz, b, c = item
            r = p.add_run()
            font(r, sz, b, c, title=title or (b and i == 0))
            r.text = txt
        else:
            r = p.add_run()
            font(r, title=title)
            r.text = item
        p.space_after = Pt(3)
    return box


def hdr(slide, title, n, total=7):
    seal(slide, PL, PT, Inches(0.4), f"{n:02d}", 10)
    tb(slide, PL + Inches(0.52), PT, IW * 0.7, Inches(0.4), [(title, 20, True, INK)], title=True)
    tb(slide, PL + IW * 0.75, PT, IW * 0.25, Inches(0.35), [(f"卷 {n}/{total}", 11, False, MUTED)], PP_ALIGN.RIGHT)
    cloud_div(slide, PT + Inches(0.46))


def card(slide, l, t, w, h, lines, accent=True, fill=BG):
    if fill != BG:
        wash(slide, l, t, w, h, fill)
    else:
        line_box(slide, l, t, w, h, BG, GREEN)
    if accent:
        green_bar(slide, l, t, h)
    tb(slide, l + Inches(0.16), t + Inches(0.12), w - Inches(0.22), h - Inches(0.18), lines)


# ── Slide 1: 封面 · 左对齐大字 + 水墨晕染底 ──
def s01(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    # 水墨山水意象：层叠淡青块（非圆形）
    wash(s, Inches(7.5), Inches(0.5), Inches(5.5), Inches(3.2), GREEN_WASH)
    wash(s, Inches(8.8), Inches(2.8), Inches(4.0), Inches(2.5), GREEN_SOFT)
    wash(s, Inches(0.3), Inches(5.0), Inches(4.5), Inches(2.0), GREEN_WASH)  # 泥捏手艺意象区
    seal(s, PL, PT, Inches(0.48), "廿三", 11)
    tb(s, PL, Inches(1.5), Inches(8.5), Inches(1.5), [("女娲造人", 52, True, INK)], title=True)
    green_bar(s, PL, Inches(3.15), Inches(0.55), Inches(2.8))
    tb(s, PL + Inches(0.15), Inches(3.2), Inches(4.0), Inches(0.55), [("创造与生命", 24, True, GREEN)], title=True)
    tb(
        s,
        PL,
        Inches(4.0),
        Inches(7.0),
        Inches(1.2),
        [
            ("统编版 · 七年级上册第六单元第23课", 14, False, MUTED),
            ("想象力博物馆 · 策展主题：创造与生命", 14, False, MUTED),
        ],
    )
    cloud_div(s, Inches(6.6))
    tb(s, PL, Inches(6.8), IW, Inches(0.35), [("古籍为骨 · 想象为肉 · 情感为魂", 13, False, GREEN)], PP_ALIGN.CENTER)


# ── Slide 2: 文学常识 · 双栏卡片 ──
def s02(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    hdr(s, "文学常识", 2)
    cw = (IW - Inches(0.2)) / 2
    card(
        s,
        PL,
        Inches(0.72),
        cw,
        Inches(5.85),
        [
            ("作者与文体", 16, True, GREEN),
            ("", 6, False, INK),
            ("作者", 13, True, RED),
            ("袁珂（现代作家、神话研究专家）", 14, False, INK),
            ("", 6, False, INK),
            ("文体", 13, True, RED),
            ("神话——用想象解释自然与人类起源", 14, False, INK),
        ],
        fill=GREEN_SOFT,
    )
    card(
        s,
        PL + cw + Inches(0.2),
        Inches(0.72),
        cw,
        Inches(5.85),
        [
            ("创作特征", 16, True, GREEN),
            ("", 6, False, INK),
            ("改写特点", 13, True, RED),
            ("基于古籍记载增删改写，赋予现代意识", 14, False, INK),
            ("", 6, False, INK),
            ("对比材料", 13, True, RED),
            ("教材「阅读提示」中的《风俗通》记载", 14, False, INK),
        ],
    )


# ── Slide 3: 造人过程 · 时间轴流程图 ──
def s03(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    hdr(s, "造人起因与过程", 3)
    # 起点
    line_box(s, PL + Inches(3.5), Inches(0.75), Inches(5.3), Inches(0.95), GREEN_SOFT, GREEN, Pt(2))
    tb(
        s,
        PL + Inches(3.65),
        Inches(0.88),
        Inches(5.0),
        Inches(0.75),
        [("起因", 13, True, RED), ("女娲行走世间，感到孤独寂寞", 16, True, INK)],
        PP_ALIGN.CENTER,
        MSO_ANCHOR.MIDDLE,
    )
    # 竖线
    ln = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, PL + Inches(6.1), Inches(1.75), Inches(0.02), Inches(0.55))
    ln.fill.solid()
    ln.fill.fore_color.rgb = GREEN
    ln.line.fill.background()
    # 双分支
    card(
        s,
        PL,
        Inches(2.4),
        Inches(5.6),
        Inches(2.35),
        [
            ("方式一 · 精细制造", 14, True, GREEN),
            ("黄泥和水，揉成小泥人", 14, False, INK),
            ("泥人落地即活", 14, True, INK),
        ],
        fill=GREEN_SOFT,
    )
    card(
        s,
        PL + Inches(6.15),
        Inches(2.4),
        Inches(5.6),
        Inches(2.35),
        [
            ("方式二 · 批量制造", 14, True, GREEN),
            ("藤条蘸泥浆挥洒", 14, False, INK),
            ("溅落泥点化为人", 14, True, INK),
        ],
    )
    # 汇合线
    for x in (Inches(2.8), Inches(9.5)):
        ln2 = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, Inches(4.85), Inches(0.02), Inches(0.45))
        ln2.fill.solid()
        ln2.fill.fore_color.rgb = GREEN
        ln2.line.fill.background()
    hln = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, PL + Inches(2.8), Inches(5.25), Inches(6.72), Inches(0.02))
    hln.fill.solid()
    hln.fill.fore_color.rgb = GREEN
    hln.line.fill.background()
    vln = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, PL + Inches(6.1), Inches(5.25), Inches(0.02), Inches(0.45))
    vln.fill.solid()
    vln.fill.fore_color.rgb = GREEN
    vln.line.fill.background()
    # 终点
    line_box(s, PL + Inches(2.8), Inches(5.75), Inches(6.72), Inches(1.05), GREEN_SOFT, RED, Pt(2))
    tb(
        s,
        PL + Inches(3.0),
        Inches(5.88),
        Inches(6.3),
        Inches(0.85),
        [
            ("结果", 13, True, RED),
            ("建立婚姻制度，让人类繁衍生息", 16, True, INK),
        ],
        PP_ALIGN.CENTER,
        MSO_ANCHOR.MIDDLE,
    )


# ── Slide 4: 想象特点 · 2×2 四宫格 ──
def s04(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    hdr(s, "想象特点分析", 4)
    items = [
        ("壹", "具体", "池水照影见自己面容 → 想到造同类"),
        ("贰", "生动", "藤条一挥，满天泥浆洒落，画面感强"),
        ("叁", "温暖", "造人后的疲倦与喜悦，赋予神以人的情感"),
        ("肆", "规律", "想象基于生活经验（和泥、洒水），又超越现实"),
    ]
    cw, ch = (IW - Inches(0.15)) / 2, Inches(2.55)
    positions = [
        (PL, Inches(0.72)),
        (PL + cw + Inches(0.15), Inches(0.72)),
        (PL, Inches(0.72) + ch + Inches(0.15)),
        (PL + cw + Inches(0.15), Inches(0.72) + ch + Inches(0.15)),
    ]
    for (mark, title, desc), (x, y) in zip(items, positions):
        fill = GREEN_SOFT if mark in ("壹", "叁") else BG
        line_box(s, x, y, cw, ch, fill, GREEN)
        green_bar(s, x, y, ch)
        seal(s, x + cw - Inches(0.55), y + Inches(0.12), Inches(0.42), mark, 12)
        tb(
            s,
            x + Inches(0.16),
            y + Inches(0.15),
            cw - Inches(0.25),
            ch - Inches(0.2),
            [(title, 18, True, GREEN), ("", 4, False, INK), (desc, 14, False, INK)],
        )


# ── Slide 5: 比较阅读 · 左右对照 + 底部横幅 ──
def s05(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    hdr(s, "比较阅读", 5)
    cw = (IW - Inches(0.2)) / 2
    card(
        s,
        PL,
        Inches(0.72),
        cw,
        Inches(3.55),
        [
            ("《风俗通》古籍原文", 16, True, GREEN),
            ("", 6, False, INK),
            ("文字简洁，重在说明造人方法。", 15, False, INK),
            ("", 6, False, INK),
            ("「女娲抟黄土作人……", 14, False, MUTED),
            ("引绳于泥中，举以为人。」", 14, False, MUTED),
        ],
        fill=GREEN_WASH,
    )
    card(
        s,
        PL + cw + Inches(0.2),
        Inches(0.72),
        cw,
        Inches(3.55),
        [
            ("课文改写", 16, True, GREEN),
            ("", 6, False, INK),
            ("增加孤独心理、造人细节", 15, False, INK),
            ("与情感描写，更具温度。", 15, False, INK),
        ],
        fill=GREEN_SOFT,
    )
    line_box(s, PL, Inches(4.5), IW, Inches(2.05), BG, RED, Pt(2))
    green_bar(s, PL, Inches(4.5), Inches(2.05), Inches(0.06))
    tb(
        s,
        PL + Inches(0.2),
        Inches(4.65),
        IW - Inches(0.35),
        Inches(1.75),
        [
            ("改写意图与主题", 14, True, RED),
            ("让神话更生动，表达对生命与创造的礼赞", 16, True, INK),
            ("策展主题：生命来之不易，创造值得珍视（创造与生命）", 14, False, GREEN),
        ],
    )


# ── Slide 6: 课堂小结 · 双核心焦点 ──
def s06(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    hdr(s, "课堂小结", 6)
    cw = (IW - Inches(0.2)) / 2
    card(
        s,
        PL,
        Inches(0.85),
        cw,
        Inches(5.5),
        [
            ("核心一 · 神话想象", 18, True, GREEN),
            ("", 8, False, INK),
            ("基于古籍", 20, True, INK),
            ("超越古籍", 20, True, INK),
            ("富有温度", 20, True, RED),
        ],
        fill=GREEN_SOFT,
    )
    card(
        s,
        PL + cw + Inches(0.2),
        Inches(0.85),
        cw,
        Inches(5.5),
        [
            ("核心二 · 女娲形象", 18, True, GREEN),
            ("", 8, False, INK),
            ("神性与人性并存", 16, True, INK),
            ("", 6, False, INK),
            ("孤独 · 疲倦 · 喜悦", 22, True, RED),
            ("", 6, False, INK),
            ("涵盖情感变化的完整母亲形象", 14, False, MUTED),
        ],
    )


# ── Slide 7: 课后延伸 · 分层任务框 ──
def s07(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    hdr(s, "课后延伸", 7)
    line_box(s, PL, Inches(0.85), IW, Inches(2.45), GREEN_SOFT, GREEN, Pt(2))
    seal(s, PL + Inches(0.15), Inches(1.05), Inches(0.48), "必", 13)
    tb(
        s,
        PL + Inches(0.75),
        Inches(1.0),
        IW - Inches(0.9),
        Inches(2.1),
        [
            ("任务一 · 必做", 16, True, GREEN),
            ("以「我看见女娲……」为题，描写100字画面。", 16, False, INK),
            ("要求：有具体动作或细节，体现造人的神奇与温情。", 13, False, MUTED),
        ],
    )
    line_box(s, PL, Inches(3.55), IW, Inches(2.45), BG, RED, Pt(2))
    seal(s, PL + Inches(0.15), Inches(3.75), Inches(0.48), "选", 13)
    tb(
        s,
        PL + Inches(0.75),
        Inches(3.7),
        IW - Inches(0.9),
        Inches(2.1),
        [
            ("任务二 · 选做", 16, True, RED),
            ("查找一个中国创世神话，与本文进行比较阅读。", 16, False, INK),
            ("思考：想象依据、表现方式、现实意味有何异同。", 13, False, MUTED),
        ],
    )
    cloud_div(s, Inches(6.35))
    tb(
        s,
        PL,
        Inches(6.5),
        IW,
        Inches(0.4),
        [("创造与生命 · 在神话中照见人类对存在的追问", 13, False, GREEN)],
        PP_ALIGN.CENTER,
    )


def main():
    prs = Presentation()
    prs.slide_width = SW
    prs.slide_height = SH
    for fn in (s01, s02, s03, s04, s05, s06, s07):
        fn(prs)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(OUT))
    import shutil

    shutil.copy2(OUT, OUT_EN)
    # sync english alias used elsewhere
    shutil.copy2(OUT, Path("/workspace/nuwa-creation-myth-8slides.pptx"))
    print(f"✓ {OUT}")
    print(f"✓ {OUT_EN}")


if __name__ == "__main__":
    main()
