#!/usr/bin/env python3
"""12-slide light museum PPT: 想象与真实——想象力博物馆的追光之旅."""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

OUT = Path("/workspace/第六单元PPT/想象力博物馆追光之旅第1-2课时.pptx")
OUT_EN = Path("/workspace/imagination-museum-lesson1-2.pptx")

SW = Inches(13.333)
SH = Inches(7.5)
PL, PT = Inches(0.5), Inches(0.32)
IW = SW - Inches(1.0)

# Light museum palette (no dark backgrounds)
BG = RGBColor(0xFA, 0xF6, 0xEE)
CARD = RGBColor(0xFF, 0xFF, 0xFF)
GOLD = RGBColor(0xC4, 0x8A, 0x1A)
GOLD_SOFT = RGBColor(0xF6, 0xE7, 0xC1)
TEAL = RGBColor(0x1A, 0x8A, 0x7A)
TEAL_SOFT = RGBColor(0xE4, 0xF4, 0xF0)
CORAL = RGBColor(0xC4, 0x5C, 0x4A)
CORAL_SOFT = RGBColor(0xF8, 0xE8, 0xE4)
BLUE = RGBColor(0x3D, 0x6B, 0x8A)
BLUE_SOFT = RGBColor(0xE6, 0xEE, 0xF4)
TEXT = RGBColor(0x2F, 0x3E, 0x46)
MUTED = RGBColor(0x6B, 0x7B, 0x8C)
LINE = RGBColor(0xE8, 0xDE, 0xC8)


def font(run, size=16, bold=False, color=TEXT):
    run.font.name = "Microsoft YaHei"
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


def rect(slide, l, t, w, h, fill, line=None, width=Pt(1.5)):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = width
    return sh


def tb(slide, l, t, w, h, lines, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
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
            font(r, sz, b, c)
            r.text = txt
        else:
            r = p.add_run()
            font(r)
            r.text = item
        p.space_after = Pt(3)
    return box


def header(slide, title, num, total=12):
    tb(slide, PL, PT, IW * 0.72, Inches(0.32), [(title, 18, True, TEAL)])
    tb(
        slide,
        PL + IW * 0.72,
        PT,
        IW * 0.28,
        Inches(0.32),
        [(f"{num:02d} / {total:02d}", 13, False, GOLD)],
        PP_ALIGN.RIGHT,
    )
    ln = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, PL, PT + Inches(0.34), IW, Inches(0.025)
    )
    ln.fill.solid()
    ln.fill.fore_color.rgb = GOLD
    ln.line.fill.background()


def pill(slide, l, t, w, h, text, fill=GOLD_SOFT, color=GOLD):
    rect(slide, l, t, w, h, fill, GOLD)
    tb(slide, l, t, w, h, [(text, 11, True, color)], PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)


def s01(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    pill(s, PL + Inches(4.6), Inches(0.55), Inches(3.1), Inches(0.38), "展特编号：七上 · Unit 6")
    rect(s, PL + Inches(2.1), Inches(1.15), Inches(8.1), Inches(4.55), CARD, GOLD, Pt(2.25))
    tb(
        s,
        PL + Inches(2.3),
        Inches(1.55),
        Inches(7.7),
        Inches(1.4),
        [("想象与真实：看不见的线", 32, True, TEXT)],
        PP_ALIGN.CENTER,
    )
    tb(
        s,
        PL + Inches(2.3),
        Inches(2.95),
        Inches(7.7),
        Inches(1.0),
        [
            ("七年级上册第六单元", 18, False, MUTED),
            ("「想象力博物馆」策展人导学", 20, True, TEAL),
        ],
        PP_ALIGN.CENTER,
    )
    tb(
        s,
        PL + Inches(2.3),
        Inches(4.2),
        Inches(7.7),
        Inches(0.9),
        [
            ("首席导览员：语文教师", 14, False, GOLD),
            ("实习策展人：全体同学", 14, False, GOLD),
        ],
        PP_ALIGN.CENTER,
    )
    tb(
        s,
        PL,
        Inches(6.85),
        IW,
        Inches(0.35),
        [("神魔 · 童话 · 神话 · 寓言  |  一场追光之旅", 13, False, MUTED)],
        PP_ALIGN.CENTER,
    )


def s02(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "策展人任务书 · 单元挑战", 2)
    # left large
    rect(s, PL, Inches(0.85), Inches(6.35), Inches(5.7), CARD, CORAL, Pt(2))
    tb(s, PL + Inches(0.25), Inches(1.05), Inches(5.85), Inches(0.35), [("致命提问", 13, True, CORAL)])
    tb(
        s,
        PL + Inches(0.25),
        Inches(1.55),
        Inches(5.85),
        Inches(2.4),
        [
            ("为什么有的想象让人发笑，", 20, True, TEXT),
            ("有的却让我们突然看见", 20, True, TEXT),
            ("真实的自己？", 24, True, CORAL),
        ],
    )
    tb(
        s,
        PL + Inches(0.25),
        Inches(4.3),
        Inches(5.85),
        Inches(1.8),
        [
            ("一切想象都有来处。", 15, False, MUTED),
            ("本单元的挑战，是寻找想象与真实", 15, False, MUTED),
            ("之间那条看不见的线。", 15, False, MUTED),
        ],
    )
    # right top
    rect(s, PL + Inches(6.55), Inches(0.85), Inches(5.3), Inches(2.55), GOLD_SOFT, GOLD, Pt(2))
    tb(s, PL + Inches(6.75), Inches(1.05), Inches(4.9), Inches(0.35), [("身份解锁", 13, True, GOLD)])
    tb(
        s,
        PL + Inches(6.75),
        Inches(1.5),
        Inches(4.9),
        Inches(1.6),
        [
            ("你将担任", 16, False, TEXT),
            ("「想象力博物馆」", 22, True, TEXT),
            ("实习策展人", 22, True, TEAL),
        ],
    )
    # right bottom
    rect(s, PL + Inches(6.55), Inches(3.6), Inches(5.3), Inches(2.95), TEAL_SOFT, TEAL, Pt(2))
    tb(s, PL + Inches(6.75), Inches(3.8), Inches(4.9), Inches(0.35), [("终极任务", 13, True, TEAL)])
    tb(
        s,
        PL + Inches(6.75),
        Inches(4.3),
        Inches(4.9),
        Inches(1.9),
        [
            ("追寻来处", 20, True, TEXT),
            ("➔  读懂逻辑", 20, True, TEXT),
            ("➔  创编作品", 20, True, TEAL),
        ],
    )


def s03(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "展厅地图 · 文体与技能矩阵", 3)
    halls = [
        ("展厅 1 · 神魔", "《小圣施威降大圣》", "自由与规则", GOLD, GOLD_SOFT),
        ("展厅 2 · 童话", "《皇帝的新装》", "荒诞与清醒", CORAL, CORAL_SOFT),
        ("展厅 3 · 神话", "《女娲造人》", "诗意与起源", TEAL, TEAL_SOFT),
        ("展厅 4 · 寓言", "寓言四则", "小故事大智慧", BLUE, BLUE_SOFT),
    ]
    w = Inches(2.85)
    gap = Inches(0.18)
    x = PL
    for title, book, theme, border, fill in halls:
        rect(s, x, Inches(0.9), w, Inches(3.35), fill, border, Pt(2))
        tb(s, x + Inches(0.12), Inches(1.1), w - Inches(0.24), Inches(0.45), [(title, 14, True, border)])
        tb(s, x + Inches(0.12), Inches(1.7), w - Inches(0.24), Inches(1.2), [(book, 16, True, TEXT)])
        tb(s, x + Inches(0.12), Inches(3.15), w - Inches(0.24), Inches(0.7), [(theme, 15, True, border)])
        x += w + gap
    rect(s, PL, Inches(4.5), IW, Inches(2.15), CARD, GOLD)
    tb(
        s,
        PL + Inches(0.25),
        Inches(4.7),
        IW - Inches(0.5),
        Inches(1.8),
        [
            ("母题链：自由  →  真实  →  创造  →  智慧", 18, True, TEXT),
            ("寻找想象与真实之间那条看不见的线", 16, False, MUTED),
            ("今日入馆：展厅 1《小圣施威降大圣》", 15, True, TEAL),
        ],
    )


def s04(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "能力清单 · 你将学会什么", 4)
    items = [
        ("1", "快速默读", "用「节点词 + 简洁句」概括情节"),
        ("2", "想象阅读", "说出想象的依据、特点与表达效果"),
        ("3", "叙事鉴赏", "有证据地分析人物形象"),
        ("4", "文言积累", "疏通《穿井得一人》《杞人忧天》"),
        ("5", "联想想象写作", "续写、改写、课本剧创编"),
    ]
    y = Inches(0.85)
    for n, title, desc in items:
        rect(s, PL, y, IW, Inches(0.78), CARD, TEAL)
        rect(s, PL + Inches(0.12), y + Inches(0.16), Inches(0.48), Inches(0.46), TEAL_SOFT, TEAL)
        tb(
            s,
            PL + Inches(0.12),
            y + Inches(0.2),
            Inches(0.48),
            Inches(0.4),
            [(n, 16, True, TEAL)],
            PP_ALIGN.CENTER,
        )
        tb(s, PL + Inches(0.75), y + Inches(0.1), Inches(3.2), Inches(0.55), [(title, 18, True, TEXT)])
        tb(s, PL + Inches(4.1), y + Inches(0.18), Inches(7.4), Inches(0.5), [(desc, 15, False, MUTED)])
        y += Inches(0.88)
    tb(
        s,
        PL,
        Inches(6.55),
        IW,
        Inches(0.4),
        [("今日聚焦：第1课时单元感知+默读  |  第2课时变化链+自由与规则", 14, True, GOLD)],
        PP_ALIGN.CENTER,
    )


def s05(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "极速挑战 · 默读训练营（第1课时）", 5)
    rect(s, PL, Inches(0.85), IW, Inches(1.15), TEAL_SOFT, TEAL, Pt(2))
    tb(
        s,
        PL + Inches(0.25),
        Inches(1.05),
        IW - Inches(0.5),
        Inches(0.8),
        [
            ("目标：不出声 · 不指读 · 每分钟 ≥ 400 字", 22, True, TEAL),
            ("阅读文本：《小圣施威降大圣》  |  段末停顿，记录段意", 14, False, MUTED),
        ],
        PP_ALIGN.CENTER,
    )
    cols = [
        ("开端", "寻找节点词", GOLD, GOLD_SOFT),
        ("冲突", "锁定关键词", CORAL, CORAL_SOFT),
        ("结果", "一句话概括", TEAL, TEAL_SOFT),
    ]
    w = Inches(3.85)
    x = PL
    for title, desc, border, fill in cols:
        rect(s, x, Inches(2.25), w, Inches(3.0), fill, border, Pt(2))
        tb(s, x + Inches(0.2), Inches(2.5), w - Inches(0.4), Inches(0.8), [(title, 24, True, border)], PP_ALIGN.CENTER)
        tb(s, x + Inches(0.2), Inches(3.5), w - Inches(0.4), Inches(1.2), [(desc, 18, False, TEXT)], PP_ALIGN.CENTER)
        x += w + Inches(0.2)
    tb(
        s,
        PL,
        Inches(6.5),
        IW,
        Inches(0.4),
        [("课堂工具：教师计时 3 分钟，学生默读后填写情节节点卡", 14, False, MUTED)],
        PP_ALIGN.CENTER,
    )


def s06(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "情节节点卡 · 读得快也读得准", 6)
    headers = ["情节节点", "关键词", "一句话概括"]
    rows = [
        ["开端", "________", "________"],
        ["变化 / 冲突", "________", "________"],
        ["结果 / 转折", "________", "________"],
    ]
    col_w = [Inches(2.6), Inches(3.5), Inches(5.75)]
    y = Inches(0.9)
    x = PL
    for i, h in enumerate(headers):
        rect(s, x, y, col_w[i] - Inches(0.06), Inches(0.55), TEAL, None)
        tb(s, x, y, col_w[i] - Inches(0.06), Inches(0.55), [(h, 15, True, CARD)], PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
        x += col_w[i]
    y = Inches(1.55)
    for row in rows:
        x = PL
        for i, cell in enumerate(row):
            rect(s, x, y, col_w[i] - Inches(0.06), Inches(0.85), CARD, LINE)
            tb(s, x + Inches(0.1), y, col_w[i] - Inches(0.2), Inches(0.85), [(cell, 16, False, TEXT)], PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
            x += col_w[i]
        y += Inches(0.95)
    rect(s, PL, Inches(4.55), Inches(5.9), Inches(2.15), GOLD_SOFT, GOLD)
    tb(
        s,
        PL + Inches(0.2),
        Inches(4.75),
        Inches(5.5),
        Inches(1.8),
        [
            ("快速复述三要素", 16, True, GOLD),
            ("人物  ＋  主要事件  ＋  结果", 18, True, TEXT),
        ],
    )
    rect(s, PL + Inches(6.15), Inches(4.55), Inches(5.7), Inches(2.15), TEAL_SOFT, TEAL)
    tb(
        s,
        PL + Inches(6.35),
        Inches(4.75),
        Inches(5.3),
        Inches(1.8),
        [
            ("课前热身", 16, True, TEAL),
            ("伞能说话、星星掉进教室……", 15, False, TEXT),
            ("离奇念头，往往从真实经验出发。", 15, False, TEXT),
        ],
    )


def s07(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "斗法解说席 · 七十二变对决卡（第2课时）", 7)
    rect(s, PL, Inches(0.9), Inches(5.35), Inches(5.15), CORAL_SOFT, CORAL, Pt(2.25))
    tb(s, PL + Inches(0.2), Inches(1.1), Inches(4.95), Inches(0.5), [("孙悟空", 22, True, CORAL)], PP_ALIGN.CENTER)
    tb(
        s,
        PL + Inches(0.25),
        Inches(1.75),
        Inches(4.85),
        Inches(3.9),
        [
            ("麻雀", 20, True, TEXT),
            ("➔  大海鹤", 20, True, TEXT),
            ("➔  鱼儿", 20, True, TEXT),
            ("➔  水蛇", 20, True, TEXT),
            ("➔  土地庙", 20, True, CORAL),
        ],
        PP_ALIGN.CENTER,
    )
    rect(s, PL + Inches(5.55), Inches(2.55), Inches(1.25), Inches(1.6), GOLD_SOFT, GOLD, Pt(2))
    tb(
        s,
        PL + Inches(5.55),
        Inches(2.7),
        Inches(1.25),
        Inches(1.35),
        [("VS", 28, True, GOLD), ("一物降一物", 11, True, TEXT)],
        PP_ALIGN.CENTER,
        MSO_ANCHOR.MIDDLE,
    )
    rect(s, PL + Inches(7.0), Inches(0.9), Inches(5.35), Inches(5.15), BLUE_SOFT, BLUE, Pt(2.25))
    tb(s, PL + Inches(7.2), Inches(1.1), Inches(4.95), Inches(0.5), [("二郎神", 22, True, BLUE)], PP_ALIGN.CENTER)
    tb(
        s,
        PL + Inches(7.25),
        Inches(1.75),
        Inches(4.85),
        Inches(3.9),
        [
            ("鹞鹰", 20, True, TEXT),
            ("➔  大鹚老", 20, True, TEXT),
            ("➔  鱼鹰", 20, True, TEXT),
            ("➔  灰鹤", 20, True, TEXT),
            ("➔  破绽识破", 20, True, BLUE),
        ],
        PP_ALIGN.CENTER,
    )
    tb(
        s,
        PL,
        Inches(6.25),
        IW,
        Inches(0.55),
        [("互动：学生说一组变化，教师点一组对照，读懂「见招拆招」", 14, False, MUTED)],
        PP_ALIGN.CENTER,
    )


def s08(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "七十二变对照表 · 变化逻辑链", 8)
    headers = ["轮次", "孙悟空", "二郎神", "关系 / 效果"]
    rows = [
        ["1", "麻雀", "鹞鹰", "一物降一物"],
        ["2", "大海鹤", "大鹚老", "体型压制"],
        ["3", "鱼儿", "鱼鹰", "环境判断"],
        ["4", "水蛇", "灰鹤", "紧追不舍"],
        ["5", "土地庙", "识破旗竿", "细节破绽"],
    ]
    col_w = [Inches(1.5), Inches(3.1), Inches(3.1), Inches(4.15)]
    y = Inches(0.85)
    x = PL
    for i, h in enumerate(headers):
        rect(s, x, y, col_w[i] - Inches(0.06), Inches(0.5), TEAL)
        tb(s, x, y, col_w[i] - Inches(0.06), Inches(0.5), [(h, 14, True, CARD)], PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
        x += col_w[i]
    y = Inches(1.42)
    fills = [CARD, GOLD_SOFT, CARD, GOLD_SOFT, CARD]
    for ri, row in enumerate(rows):
        x = PL
        for i, cell in enumerate(row):
            rect(s, x, y, col_w[i] - Inches(0.06), Inches(0.72), fills[ri], LINE)
            tb(s, x, y, col_w[i] - Inches(0.06), Inches(0.72), [(cell, 16, True if i == 0 else False, TEXT)], PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
            x += col_w[i]
        y += Inches(0.78)
    tb(
        s,
        PL,
        Inches(6.45),
        IW,
        Inches(0.4),
        [("变化因果：设法脱身  ↔  据形识破  ↔  斗法推进", 16, True, TEAL)],
        PP_ALIGN.CENTER,
    )


def s09(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "细节放大镜 · 土地庙的尾巴", 9)
    rect(s, PL, Inches(0.9), Inches(5.7), Inches(4.35), CARD, GOLD, Pt(2))
    tb(s, PL + Inches(0.25), Inches(1.1), Inches(5.2), Inches(0.4), [("课文摘引", 14, True, GOLD)])
    tb(
        s,
        PL + Inches(0.25),
        Inches(1.7),
        Inches(5.2),
        Inches(3.1),
        [
            ("「只有尾巴不好收拾，", 18, False, TEXT),
            ("竖在后面，变做一根旗竿。」", 18, True, CORAL),
            ("", 12, False, MUTED),
            ("二郎神：「更不曾见一个", 16, False, TEXT),
            ("旗竿竖在后面的。」", 16, True, BLUE),
        ],
    )
    rect(s, PL + Inches(6.0), Inches(0.9), Inches(5.85), Inches(4.35), TEAL_SOFT, TEAL, Pt(2))
    tb(s, PL + Inches(6.2), Inches(1.1), Inches(5.45), Inches(0.4), [("放大镜 · 推理链", 14, True, TEAL)])
    tb(
        s,
        PL + Inches(6.2),
        Inches(1.7),
        Inches(5.45),
        Inches(3.2),
        [
            ("孙悟空的习惯破绽", 17, True, TEXT),
            ("➔  视觉异常（旗竿竖后）", 17, False, TEXT),
            ("➔  二郎神据形识破", 17, False, TEXT),
            ("➔  破局 · 悟空逃走", 17, True, TEAL),
        ],
    )
    rect(s, PL, Inches(5.5), IW, Inches(1.2), GOLD_SOFT, GOLD, Pt(2))
    tb(
        s,
        PL + Inches(0.2),
        Inches(5.7),
        IW - Inches(0.4),
        Inches(0.9),
        [("奇趣来源于合理的细节，无理之处见妙趣。", 22, True, TEXT)],
        PP_ALIGN.CENTER,
        MSO_ANCHOR.MIDDLE,
    )


def s10(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "环境推理图 · 想象的判断依据", 10)
    rect(s, PL, Inches(0.85), IW, Inches(1.35), CARD, GOLD)
    tb(
        s,
        PL + Inches(0.25),
        Inches(1.0),
        IW - Inches(0.5),
        Inches(1.1),
        [
            ("片段一：「这猢狲必然下水去也，定变作鱼虾之类。」", 18, True, TEXT),
            ("二郎神为什么想到「鱼虾之类」？——据环境作出判断。", 15, False, MUTED),
        ],
    )
    steps = [
        ("孙悟空", "入涧中"),
        ("环境", "水里可藏身"),
        ("判断", "必变鱼虾"),
        ("应对", "变作鱼鹰"),
    ]
    w = Inches(2.7)
    x = PL
    for i, (a, b) in enumerate(steps):
        rect(s, x, Inches(2.45), w, Inches(1.7), TEAL_SOFT if i % 2 == 0 else GOLD_SOFT, TEAL if i % 2 == 0 else GOLD)
        tb(s, x, Inches(2.6), w, Inches(0.55), [(a, 14, True, TEAL if i % 2 == 0 else GOLD)], PP_ALIGN.CENTER)
        tb(s, x, Inches(3.15), w, Inches(0.7), [(b, 18, True, TEXT)], PP_ALIGN.CENTER)
        if i < 3:
            tb(s, x + w - Inches(0.05), Inches(2.95), Inches(0.35), Inches(0.5), [("→", 20, True, GOLD)], PP_ALIGN.CENTER)
        x += w + Inches(0.22)
    rect(s, PL, Inches(4.4), IW, Inches(2.2), CARD, TEAL)
    tb(
        s,
        PL + Inches(0.25),
        Inches(4.6),
        IW - Inches(0.5),
        Inches(1.85),
        [
            ("土地庙推理卡", 15, True, TEAL),
            ("悟空想利用【环境】隐藏  →  留下【形态】破绽", 16, False, TEXT),
            ("→  二郎神发现【异常】  →  悟空决定【逃走】", 16, False, TEXT),
            ("小组讨论角度：环境  /  形态  /  对手", 15, True, GOLD),
        ],
    )


def s11(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "哲学天平 · 自由有边界吗？", 11)
    rect(s, PL, Inches(0.9), Inches(5.85), Inches(3.55), CARD, MUTED, Pt(1.5))
    tb(s, PL + Inches(0.25), Inches(1.1), Inches(5.35), Inches(0.45), [("虚线框 · 绝对自由？", 16, True, MUTED)])
    tb(
        s,
        PL + Inches(0.25),
        Inches(1.7),
        Inches(5.35),
        Inches(2.4),
        [
            ("想变什么就变什么", 18, False, TEXT),
            ("看似自由", 18, False, TEXT),
            ("实则无序、被动", 18, True, CORAL),
        ],
    )
    rect(s, PL + Inches(6.15), Inches(0.9), Inches(5.7), Inches(3.55), GOLD_SOFT, GOLD, Pt(2.25))
    tb(s, PL + Inches(6.4), Inches(1.1), Inches(5.2), Inches(0.45), [("金框 · 规则之下的自由", 16, True, GOLD)])
    tb(
        s,
        PL + Inches(6.4),
        Inches(1.7),
        Inches(5.2),
        Inches(2.4),
        [
            ("受制于环境", 18, True, TEXT),
            ("受制于形态", 18, True, TEXT),
            ("受制于对手判断与自身破绽", 18, True, TEAL),
        ],
    )
    rect(s, PL, Inches(4.7), IW, Inches(1.95), TEAL_SOFT, TEAL, Pt(2))
    tb(
        s,
        PL + Inches(0.3),
        Inches(4.9),
        IW - Inches(0.6),
        Inches(1.6),
        [
            ("最天马行空的想象，背后也是最扎实的现实逻辑；", 18, True, TEXT),
            ("神通越强大，越需在规则中寻得智慧。", 20, True, TEAL),
        ],
        PP_ALIGN.CENTER,
        MSO_ANCHOR.MIDDLE,
    )


def s12(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "闭馆小结 · 作业检测", 12)
    rect(s, PL, Inches(0.9), Inches(5.9), Inches(4.35), CARD, TEAL, Pt(2))
    tb(s, PL + Inches(0.25), Inches(1.1), Inches(5.4), Inches(0.4), [("今日策展笔记", 16, True, TEAL)])
    tb(
        s,
        PL + Inches(0.25),
        Inches(1.65),
        Inches(5.4),
        Inches(3.3),
        [
            ("✓  文体：神魔 / 童话 / 神话 / 寓言", 15, False, TEXT),
            ("✓  默读：节点词 + 简洁句，≥400字/分", 15, False, TEXT),
            ("✓  斗法：一物降一物，想象有逻辑", 15, False, TEXT),
            ("✓  母题：自由需在规则中寻得智慧", 15, False, TEXT),
        ],
    )
    rect(s, PL + Inches(6.15), Inches(0.9), Inches(5.7), Inches(4.35), GOLD_SOFT, GOLD, Pt(2))
    tb(s, PL + Inches(6.4), Inches(1.1), Inches(5.2), Inches(0.4), [("作业检测", 16, True, GOLD)])
    tb(
        s,
        PL + Inches(6.4),
        Inches(1.65),
        Inches(5.2),
        Inches(3.3),
        [
            ("1. 名著链接：作者______", 15, False, TEXT),
            ("    时代______  其他情节______", 15, False, TEXT),
            ("2. 斗法解说词（100—150字）", 15, True, TEXT),
            ("    写清至少三次变化，有动作细节", 14, False, MUTED),
        ],
    )
    rect(s, PL, Inches(5.5), IW, Inches(1.2), CORAL_SOFT, CORAL)
    tb(
        s,
        PL + Inches(0.2),
        Inches(5.7),
        IW - Inches(0.4),
        Inches(0.9),
        [("下站预告：展厅 2《皇帝的新装》——荒诞与清醒", 18, True, CORAL)],
        PP_ALIGN.CENTER,
        MSO_ANCHOR.MIDDLE,
    )


def main():
    prs = Presentation()
    prs.slide_width = SW
    prs.slide_height = SH
    for fn in (s01, s02, s03, s04, s05, s06, s07, s08, s09, s10, s11, s12):
        fn(prs)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(OUT))
    import shutil

    shutil.copy2(OUT, OUT_EN)
    print(f"✓ {OUT}")
    print(f"✓ {OUT_EN}")


if __name__ == "__main__":
    main()
