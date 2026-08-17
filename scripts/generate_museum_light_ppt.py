#!/usr/bin/env python3
"""12-slide 新国风水墨卷轴 PPT: 想象与真实——想象力博物馆的追光之旅."""

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
PL, PT = Inches(0.65), Inches(0.38)
IW = SW - Inches(1.3)

# 新国风水墨卷轴配色
BG = RGBColor(0xF4, 0xF1, 0xEA)       # 宣纸米白
INK = RGBColor(0x1E, 0x1E, 0x1E)       # 墨黑
RED = RGBColor(0xC8, 0x3C, 0x23)       # 朱砂红
MUTED = RGBColor(0x5A, 0x5A, 0x5A)    # 淡墨
WASH = RGBColor(0xE8, 0xE2, 0xD6)      # 水墨晕染底

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


def ink_wash(slide, l, t, w, h, alpha_wash=True):
    """水墨晕染底：无边框浅灰块."""
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = WASH if alpha_wash else BG
    sh.line.fill.background()
    return sh


def cloud_divider(slide, y, width=None):
    """云纹分割线."""
    w = width or IW
    ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, PL, y, w, Inches(0.012))
    ln.fill.solid()
    ln.fill.fore_color.rgb = MUTED
    ln.line.fill.background()
    tb(
        slide,
        PL,
        y - Inches(0.08),
        w,
        Inches(0.25),
        [("～ ～ ～ ～ ～ ～ ～ ～ ～ ～ ～ ～ ～ ～ ～", 9, False, MUTED)],
        PP_ALIGN.CENTER,
    )


def red_bar(slide, l, t, h, width=Inches(0.06)):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, width, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = RED
    sh.line.fill.background()
    return sh


def seal(slide, l, t, size, text, font_size=11):
    """朱砂印章小徽标."""
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, size, size)
    sh.fill.background()
    sh.line.color.rgb = RED
    sh.line.width = Pt(2)
    tb(slide, l, t, size, size, [(text, font_size, True, RED)], PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    return sh


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
            use_title = title or (i == 0 and b)
            r = p.add_run()
            font(r, sz, b, c, title=use_title)
            r.text = txt
        else:
            r = p.add_run()
            font(r, title=title)
            r.text = item
        p.space_after = Pt(4)
    return box


def header(slide, title, num, total=12):
    seal(slide, PL, PT, Inches(0.42), f"{num:02d}", 10)
    tb(slide, PL + Inches(0.55), PT, IW * 0.72, Inches(0.42), [(title, 20, True, INK)], title=True)
    tb(
        slide,
        PL + IW * 0.78,
        PT + Inches(0.05),
        IW * 0.22,
        Inches(0.35),
        [(f"卷 {num} / {total}", 12, False, MUTED)],
        PP_ALIGN.RIGHT,
    )
    cloud_divider(slide, PT + Inches(0.48))


def block(slide, l, t, w, h, lines, accent=True, wash=False):
    """无边框内容块，可选左侧朱砂粗线 + 水墨底."""
    if wash:
        ink_wash(slide, l, t, w, h)
    if accent:
        red_bar(slide, l, t, h)
    tb(slide, l + Inches(0.18), t + Inches(0.12), w - Inches(0.25), h - Inches(0.2), lines)


def s01(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    seal(s, PL + Inches(5.5), Inches(0.55), Inches(0.55), "六上")
    tb(
        s,
        PL,
        Inches(0.55),
        IW,
        Inches(0.35),
        [("展特编号 · 七年级上册第六单元", 13, False, MUTED)],
        PP_ALIGN.CENTER,
    )
    cloud_divider(s, Inches(1.0))
    tb(
        s,
        PL,
        Inches(1.35),
        IW,
        Inches(1.2),
        [("想象与真实", 40, True, INK), ("看不见的线", 36, True, RED)],
        PP_ALIGN.CENTER,
        title=True,
    )
    tb(
        s,
        PL,
        Inches(2.75),
        IW,
        Inches(0.55),
        [("「想象力博物馆」策展人导学", 20, True, INK)],
        PP_ALIGN.CENTER,
        title=True,
    )
    tb(
        s,
        PL,
        Inches(3.45),
        IW,
        Inches(0.45),
        [("课时六、七  ·  追光之旅", 16, False, MUTED)],
        PP_ALIGN.CENTER,
    )
    ink_wash(s, PL + Inches(2.5), Inches(4.2), Inches(7.3), Inches(1.5))
    tb(
        s,
        PL + Inches(2.7),
        Inches(4.45),
        Inches(6.9),
        Inches(1.0),
        [
            ("神魔 · 童话 · 神话 · 寓言", 16, False, INK),
            ("首席导览员：教师    实习策展人：全体同学", 14, False, MUTED),
        ],
        PP_ALIGN.CENTER,
    )
    cloud_divider(s, Inches(6.55))
    tb(
        s,
        PL,
        Inches(6.75),
        IW,
        Inches(0.35),
        [("新国风 · 水墨卷轴 · 教学演示", 12, False, MUTED)],
        PP_ALIGN.CENTER,
    )


def s02(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "策展人任务书", 2)
    block(
        s,
        PL,
        Inches(0.75),
        Inches(6.2),
        Inches(5.8),
        [
            ("致命提问", 15, True, RED),
            ("为什么有的想象让人发笑，", 19, True, INK),
            ("有的却让我们突然看见", 19, True, INK),
            ("真实的自己？", 22, True, RED),
            ("", 8, False, MUTED),
            ("一切想象都有来处。本单元的挑战，", 14, False, MUTED),
            ("是寻找想象与真实之间那条看不见的线。", 14, False, MUTED),
        ],
        wash=True,
    )
    block(
        s,
        PL + Inches(6.45),
        Inches(0.75),
        Inches(5.25),
        Inches(2.55),
        [
            ("身份解锁", 14, True, RED),
            ("你将担任", 15, False, INK),
            ("「想象力博物馆」", 20, True, INK),
            ("实习策展人", 20, True, RED),
        ],
    )
    block(
        s,
        PL + Inches(6.45),
        Inches(3.5),
        Inches(5.25),
        Inches(3.05),
        [
            ("终极任务", 14, True, RED),
            ("追寻来处", 18, True, INK),
            ("读懂逻辑", 18, True, INK),
            ("创编作品", 18, True, RED),
        ],
        wash=True,
    )


def s03(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "展厅地图", 3)
    halls = [
        ("壹 · 神魔", "《小圣施威降大圣》", "自由与规则"),
        ("贰 · 童话", "《皇帝的新装》", "荒诞与清醒"),
        ("叁 · 神话", "《女娲造人》", "诗意与起源"),
        ("肆 · 寓言", "寓言四则", "小故事大智慧"),
    ]
    w = Inches(2.85)
    gap = Inches(0.15)
    x = PL
    for title, book, theme in halls:
        ink_wash(s, x, Inches(0.75), w, Inches(3.2))
        red_bar(s, x, Inches(0.75), Inches(3.2))
        tb(s, x + Inches(0.2), Inches(0.95), w - Inches(0.3), Inches(0.5), [(title, 14, True, RED)], title=True)
        tb(s, x + Inches(0.2), Inches(1.55), w - Inches(0.3), Inches(1.2), [(book, 16, True, INK)], title=True)
        tb(s, x + Inches(0.2), Inches(2.95), w - Inches(0.3), Inches(0.7), [(theme, 14, False, MUTED)])
        x += w + gap
    block(
        s,
        PL,
        Inches(4.2),
        IW,
        Inches(2.35),
        [
            ("母题链", 14, True, RED),
            ("自由  →  真实  →  创造  →  智慧", 20, True, INK),
            ("今日入馆：展厅壹《小圣施威降大圣》", 15, False, MUTED),
        ],
        wash=True,
    )


def s04(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "能力清单", 4)
    items = [
        ("壹", "快速默读", "用「节点词 + 简洁句」概括情节"),
        ("贰", "想象阅读", "说出想象的依据、特点与表达效果"),
        ("叁", "叙事鉴赏", "有证据地分析人物形象"),
        ("肆", "文言积累", "疏通《穿井得一人》《杞人忧天》"),
        ("伍", "联想想象写作", "续写、改写、课本剧创编"),
    ]
    y = Inches(0.75)
    for mark, title, desc in items:
        red_bar(s, PL, y, Inches(0.72))
        seal(s, PL + Inches(0.12), y + Inches(0.12), Inches(0.48), mark, 12)
        tb(s, PL + Inches(0.75), y + Inches(0.08), Inches(2.8), Inches(0.55), [(title, 17, True, INK)], title=True)
        tb(s, PL + Inches(3.7), y + Inches(0.15), Inches(8.0), Inches(0.5), [(desc, 14, False, MUTED)])
        y += Inches(0.88)
    tb(
        s,
        PL,
        Inches(6.55),
        IW,
        Inches(0.4),
        [("今日聚焦：第1课时单元感知+默读  |  第2课时变化链+自由与规则", 13, True, RED)],
        PP_ALIGN.CENTER,
    )


def s05(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "极速挑战 · 默读训练营", 5)
    block(
        s,
        PL,
        Inches(0.75),
        IW,
        Inches(1.05),
        [
            ("目标：不出声 · 不指读 · 每分钟 ≥ 400 字", 21, True, INK),
            ("阅读文本：《小圣施威降大圣》  ·  段末停顿，记录段意", 13, False, MUTED),
        ],
        wash=True,
    )
    cols = [("开端", "寻找节点词"), ("冲突", "锁定关键词"), ("结果", "一句话概括")]
    w = Inches(3.85)
    x = PL
    for title, desc in cols:
        ink_wash(s, x, Inches(2.05), w, Inches(3.05))
        red_bar(s, x, Inches(2.05), Inches(3.05))
        tb(s, x + Inches(0.2), Inches(2.35), w - Inches(0.35), Inches(0.7), [(title, 24, True, RED)], PP_ALIGN.CENTER, title=True)
        tb(s, x + Inches(0.2), Inches(3.35), w - Inches(0.35), Inches(1.2), [(desc, 17, False, INK)], PP_ALIGN.CENTER)
        x += w + Inches(0.2)
    tb(
        s,
        PL,
        Inches(6.45),
        IW,
        Inches(0.35),
        [("课堂工具：教师计时 3 分钟，学生默读后填写情节节点卡", 13, False, MUTED)],
        PP_ALIGN.CENTER,
    )


def s06(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "情节节点卡", 6)
    headers = ["情节节点", "关键词", "一句话概括"]
    rows = [
        ["开端", "________", "________"],
        ["变化 / 冲突", "________", "________"],
        ["结果 / 转折", "________", "________"],
    ]
    col_w = [Inches(2.6), Inches(3.5), Inches(5.75)]
    y = Inches(0.75)
    x = PL
    for i, h in enumerate(headers):
        tb(s, x, y, col_w[i], Inches(0.45), [(h, 14, True, RED)], PP_ALIGN.CENTER, title=True)
        if i < 2:
            tb(s, x + col_w[i], y + Inches(0.05), Inches(0.3), Inches(0.35), [("·", 14, False, MUTED)])
        x += col_w[i]
    cloud_divider(s, y + Inches(0.5))
    y = Inches(1.45)
    for row in rows:
        x = PL
        red_bar(s, x, y, Inches(0.75))
        for i, cell in enumerate(row):
            tb(s, x + Inches(0.12), y + Inches(0.1), col_w[i] - Inches(0.15), Inches(0.55), [(cell, 15, False, INK)], PP_ALIGN.CENTER)
            x += col_w[i]
        y += Inches(0.95)
    block(
        s,
        PL,
        Inches(4.45),
        Inches(5.8),
        Inches(2.05),
        [
            ("快速复述三要素", 15, True, RED),
            ("人物  ＋  主要事件  ＋  结果", 18, True, INK),
        ],
        wash=True,
    )
    block(
        s,
        PL + Inches(6.05),
        Inches(4.45),
        Inches(5.65),
        Inches(2.05),
        [
            ("课前热身", 15, True, RED),
            ("伞能说话、星星掉进教室……", 14, False, INK),
            ("离奇念头，往往从真实经验出发。", 14, False, MUTED),
        ],
    )


def s07(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "斗法解说席 · 七十二变", 7)
    ink_wash(s, PL, Inches(0.75), Inches(5.35), Inches(5.0))
    red_bar(s, PL, Inches(0.75), Inches(5.0))
    tb(s, PL + Inches(0.25), Inches(0.95), Inches(4.85), Inches(0.5), [("孙悟空", 22, True, INK)], PP_ALIGN.CENTER, title=True)
    tb(
        s,
        PL + Inches(0.3),
        Inches(1.65),
        Inches(4.75),
        Inches(3.8),
        [
            ("麻雀  →  大海鹤", 18, False, INK),
            ("  →  鱼儿  →  水蛇", 18, False, INK),
            ("  →  土地庙", 18, True, RED),
        ],
        PP_ALIGN.CENTER,
    )
    seal(s, PL + Inches(5.85), Inches(2.55), Inches(0.65), "VS", 14)
    tb(s, PL + Inches(5.55), Inches(3.35), Inches(1.25), Inches(0.4), [("一物降一物", 11, True, RED)], PP_ALIGN.CENTER)
    ink_wash(s, PL + Inches(7.0), Inches(0.75), Inches(5.35), Inches(5.0))
    red_bar(s, PL + Inches(7.0), Inches(0.75), Inches(5.0))
    tb(s, PL + Inches(7.25), Inches(0.95), Inches(4.85), Inches(0.5), [("二郎神", 22, True, INK)], PP_ALIGN.CENTER, title=True)
    tb(
        s,
        PL + Inches(7.3),
        Inches(1.65),
        Inches(4.75),
        Inches(3.8),
        [
            ("鹞鹰  →  大鹚老", 18, False, INK),
            ("  →  鱼鹰  →  灰鹤", 18, False, INK),
            ("  →  破绽识破", 18, True, RED),
        ],
        PP_ALIGN.CENTER,
    )
    tb(
        s,
        PL,
        Inches(6.15),
        IW,
        Inches(0.45),
        [("互动：学生说一组变化，教师点一组对照，读懂「见招拆招」", 13, False, MUTED)],
        PP_ALIGN.CENTER,
    )


def s08(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "七十二变对照表", 8)
    headers = ["轮次", "孙悟空", "二郎神", "关系 / 效果"]
    rows = [
        ["壹", "麻雀", "鹞鹰", "一物降一物"],
        ["贰", "大海鹤", "大鹚老", "体型压制"],
        ["叁", "鱼儿", "鱼鹰", "环境判断"],
        ["肆", "水蛇", "灰鹤", "紧追不舍"],
        ["伍", "土地庙", "识破旗竿", "细节破绽"],
    ]
    col_w = [Inches(1.5), Inches(3.1), Inches(3.1), Inches(4.15)]
    y = Inches(0.75)
    x = PL
    for h in headers:
        tb(s, x, y, col_w[headers.index(h)], Inches(0.4), [(h, 13, True, RED)], PP_ALIGN.CENTER, title=True)
        x += col_w[headers.index(h)]
    cloud_divider(s, y + Inches(0.42))
    y = Inches(1.35)
    for row in rows:
        x = PL
        red_bar(s, x, y, Inches(0.65))
        for i, cell in enumerate(row):
            tb(s, x + Inches(0.1), y + Inches(0.08), col_w[i] - Inches(0.12), Inches(0.5), [(cell, 15, i == 0, INK)], PP_ALIGN.CENTER)
            x += col_w[i]
        y += Inches(0.72)
    tb(
        s,
        PL,
        Inches(6.4),
        IW,
        Inches(0.4),
        [("变化因果：设法脱身  ↔  据形识破  ↔  斗法推进", 16, True, INK)],
        PP_ALIGN.CENTER,
        title=True,
    )


def s09(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "细节放大镜 · 土地庙", 9)
    block(
        s,
        PL,
        Inches(0.75),
        Inches(5.7),
        Inches(4.2),
        [
            ("课文摘引", 14, True, RED),
            ("「只有尾巴不好收拾，", 17, False, INK),
            ("竖在后面，变做一根旗竿。」", 17, True, RED),
            ("", 8, False, MUTED),
            ("二郎神：「更不曾见一个", 15, False, INK),
            ("旗竿竖在后面的。」", 15, True, INK),
        ],
        wash=True,
    )
    block(
        s,
        PL + Inches(6.0),
        Inches(0.75),
        Inches(5.85),
        Inches(4.2),
        [
            ("推理链", 14, True, RED),
            ("孙悟空的习惯破绽", 16, True, INK),
            ("→  视觉异常（旗竿竖后）", 15, False, INK),
            ("→  二郎神据形识破", 15, False, INK),
            ("→  破局 · 悟空逃走", 15, True, RED),
        ],
    )
    ink_wash(s, PL, Inches(5.25), IW, Inches(1.15))
    tb(
        s,
        PL + Inches(0.2),
        Inches(5.45),
        IW - Inches(0.4),
        Inches(0.85),
        [("奇趣来源于合理的细节，无理之处见妙趣。", 20, True, INK)],
        PP_ALIGN.CENTER,
        MSO_ANCHOR.MIDDLE,
        title=True,
    )


def s10(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "环境推理图", 10)
    block(
        s,
        PL,
        Inches(0.75),
        IW,
        Inches(1.25),
        [
            ("片段一", 13, True, RED),
            ("「这猢狲必然下水去也，定变作鱼虾之类。」", 17, True, INK),
            ("二郎神据环境作出判断。", 14, False, MUTED),
        ],
        wash=True,
    )
    steps = [("孙悟空", "入涧中"), ("环境", "水里可藏身"), ("判断", "必变鱼虾"), ("应对", "变作鱼鹰")]
    w = Inches(2.7)
    x = PL
    for i, (a, b) in enumerate(steps):
        ink_wash(s, x, Inches(2.25), w, Inches(1.55))
        if i == 0:
            red_bar(s, x, Inches(2.25), Inches(1.55))
        tb(s, x + Inches(0.15), Inches(2.4), w - Inches(0.25), Inches(0.45), [(a, 13, True, RED)], PP_ALIGN.CENTER, title=True)
        tb(s, x + Inches(0.15), Inches(2.95), w - Inches(0.25), Inches(0.6), [(b, 16, True, INK)], PP_ALIGN.CENTER, title=True)
        if i < 3:
            tb(s, x + w, Inches(2.75), Inches(0.25), Inches(0.4), [("→", 18, False, RED)], PP_ALIGN.CENTER)
        x += w + Inches(0.22)
    block(
        s,
        PL,
        Inches(4.15),
        IW,
        Inches(2.15),
        [
            ("土地庙推理", 14, True, RED),
            ("悟空想利用【环境】隐藏  →  留下【形态】破绽", 15, False, INK),
            ("→  二郎神发现【异常】  →  悟空决定【逃走】", 15, False, INK),
            ("讨论角度：环境  /  形态  /  对手", 14, True, RED),
        ],
    )


def s11(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "自由有边界吗", 11)
    block(
        s,
        PL,
        Inches(0.75),
        Inches(5.85),
        Inches(3.35),
        [
            ("绝对自由？", 16, True, MUTED),
            ("想变什么就变什么", 17, False, INK),
            ("看似自由", 17, False, INK),
            ("实则无序、被动", 17, True, RED),
        ],
    )
    block(
        s,
        PL + Inches(6.15),
        Inches(0.75),
        Inches(5.7),
        Inches(3.35),
        [
            ("规则之下的自由", 16, True, RED),
            ("受制于环境", 17, True, INK),
            ("受制于形态", 17, True, INK),
            ("受制于对手判断与自身破绽", 17, True, INK),
        ],
        wash=True,
    )
    ink_wash(s, PL, Inches(4.35), IW, Inches(1.85))
    red_bar(s, PL, Inches(4.35), Inches(1.85), width=Inches(0.08))
    tb(
        s,
        PL + Inches(0.25),
        Inches(4.55),
        IW - Inches(0.4),
        Inches(1.5),
        [
            ("最天马行空的想象，背后也是最扎实的现实逻辑；", 17, True, INK),
            ("神通越强大，越需在规则中寻得智慧。", 19, True, RED),
        ],
        PP_ALIGN.CENTER,
        title=True,
    )


def s12(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "闭馆小结 · 作业", 12)
    block(
        s,
        PL,
        Inches(0.75),
        Inches(5.9),
        Inches(4.1),
        [
            ("今日策展笔记", 15, True, RED),
            ("文体：神魔 / 童话 / 神话 / 寓言", 14, False, INK),
            ("默读：节点词 + 简洁句，≥400字/分", 14, False, INK),
            ("斗法：一物降一物，想象有逻辑", 14, False, INK),
            ("母题：自由需在规则中寻得智慧", 14, False, INK),
        ],
        wash=True,
    )
    block(
        s,
        PL + Inches(6.15),
        Inches(0.75),
        Inches(5.7),
        Inches(4.1),
        [
            ("作业检测", 15, True, RED),
            ("1. 名著链接：作者______  时代______", 14, False, INK),
            ("2. 斗法解说词（100—150字）", 14, True, INK),
            ("    写清至少三次变化，有动作细节", 13, False, MUTED),
        ],
    )
    ink_wash(s, PL, Inches(5.15), IW, Inches(1.05))
    tb(
        s,
        PL + Inches(0.2),
        Inches(5.35),
        IW - Inches(0.4),
        Inches(0.75),
        [("下站预告：展厅贰《皇帝的新装》——荒诞与清醒", 17, True, RED)],
        PP_ALIGN.CENTER,
        MSO_ANCHOR.MIDDLE,
        title=True,
    )
    seal(s, PL + Inches(11.8), Inches(6.55), Inches(0.5), "续", 12)


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
