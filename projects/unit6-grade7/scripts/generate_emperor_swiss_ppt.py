#!/usr/bin/env python3
"""《皇帝的新装》瑞士杂志平面设计风格 PPT — 9 slides."""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

OUT = Path("/workspace/第六单元PPT/第22课 皇帝的新装.pptx")
OUT_EN = Path("/workspace/unit6-lesson22-emperor.pptx")

SW = Inches(13.333)
SH = Inches(7.5)
PL, PT = Inches(0.45), Inches(0.35)
IW = SW - Inches(0.9)

WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x00, 0x00, 0x00)
YELLOW = RGBColor(0xFF, 0xE6, 0x00)
FONT = "Arial"


def fnt(run, size=18, bold=False, color=BLACK):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color


def bg(slide):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    sh.fill.solid()
    sh.fill.fore_color.rgb = WHITE
    sh.line.fill.background()
    tree = slide.shapes._spTree
    tree.remove(sh._element)
    tree.insert(2, sh._element)


def box(slide, l, t, w, h, fill=WHITE, stroke=BLACK, lw=Pt(2)):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if stroke is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = stroke
        sh.line.width = lw
    return sh


def tb(slide, l, t, w, h, lines, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    box_shape = slide.shapes.add_textbox(l, t, w, h)
    tf = box_shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for i, item in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if isinstance(item, tuple):
            txt, sz, b, c = item
            r = p.add_run()
            fnt(r, sz, b, c)
            r.text = txt
        else:
            r = p.add_run()
            fnt(r)
            r.text = item
        p.space_after = Pt(2)
    return box_shape


def slide_num(slide, n, total=9):
    tb(slide, PL, PT, Inches(1.2), Inches(0.3), [(f"{n:02d}", 14, True, BLACK)])


# ── Slide 1: 全屏大字居中 · Hero = 巨型标题 ──
def s01(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    slide_num(s, 1)
    # Hero: title occupies ~1/3 vertical space
    tb(
        s,
        PL,
        Inches(1.0),
        IW,
        Inches(2.6),
        [("皇帝的新装", 72, True, BLACK)],
        PP_ALIGN.CENTER,
        MSO_ANCHOR.MIDDLE,
    )
    box(s, PL + Inches(1.5), Inches(3.85), IW - Inches(3.0), Inches(0.55), YELLOW, None)
    tb(
        s,
        PL + Inches(1.5),
        Inches(3.9),
        IW - Inches(3.0),
        Inches(0.45),
        [("现实的讽刺与人性的镜子", 22, True, BLACK)],
        PP_ALIGN.CENTER,
        MSO_ANCHOR.MIDDLE,
    )
    tb(
        s,
        PL,
        Inches(4.8),
        IW,
        Inches(0.5),
        [("［丹麦］安徒生  ·  世界儿童文学巨匠", 16, False, BLACK)],
        PP_ALIGN.CENTER,
    )
    tb(
        s,
        PL,
        Inches(5.4),
        IW,
        Inches(0.4),
        [("想象力博物馆 · 真实与虚假", 14, True, BLACK)],
        PP_ALIGN.CENTER,
    )
    box(s, PL, Inches(6.55), IW, Pt(2), BLACK, None)


# ── Slide 2: 三栏交错 · Hero = 柠檬黄「03」目标编号 ──
def s02(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    slide_num(s, 2)
    tb(s, PL, Inches(0.35), IW, Inches(0.55), [("单元情境与学习目标", 28, True, BLACK)])
    box(s, PL, Inches(0.95), IW, Pt(2), BLACK, None)
    # Col 1 — narrow, top high
    box(s, PL, Inches(1.15), Inches(3.5), Inches(2.4), WHITE, BLACK)
    tb(
        s,
        PL + Inches(0.15),
        Inches(1.3),
        Inches(3.2),
        Inches(2.1),
        [
            ("单元", 14, True, BLACK),
            ("情境", 14, True, BLACK),
            ("", 6, False, BLACK),
            ("探索想象与现实的交融。", 13, False, BLACK),
            ("在荒诞童话中照见真实，", 13, False, BLACK),
            ("在虚构故事里审视人心。", 13, False, BLACK),
        ],
    )
    # Col 2 — wide center, offset down (stagger)
    box(s, PL + Inches(3.75), Inches(1.65), Inches(5.0), Inches(3.5), YELLOW, BLACK)
    tb(
        s,
        PL + Inches(3.95),
        Inches(1.85),
        Inches(4.6),
        Inches(3.1),
        [
            ("本课", 16, True, BLACK),
            ("目标", 16, True, BLACK),
            ("", 8, False, BLACK),
            ("1. 受骗—展骗—穿骗—看骗—揭骗", 14, True, BLACK),
            ("2. 分析皇帝/大臣/百姓/孩子心理", 14, True, BLACK),
            ("3. 理解虚伪与真实的现实对应", 14, True, BLACK),
        ],
    )
    # Col 3 — narrow, bottom aligned
    box(s, PL + Inches(9.0), Inches(2.35), Inches(3.4), Inches(2.8), WHITE, BLACK)
    tb(
        s,
        PL + Inches(9.15),
        Inches(2.55),
        Inches(3.1),
        Inches(2.4),
        [
            ("策展", 14, True, BLACK),
            ("主题", 14, True, BLACK),
            ("", 6, False, BLACK),
            ("真实", 36, True, BLACK),
            ("与", 20, True, BLACK),
            ("虚假", 36, True, BLACK),
        ],
        PP_ALIGN.RIGHT,
    )
    # Hero number
    tb(s, PL, Inches(5.5), Inches(2.5), Inches(1.5), [("03", 96, True, BLACK)])


# ── Slide 3: 上下分割+左侧瀑布流 · Hero = 竖排「手法」大字 ──
def s03(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    slide_num(s, 3)
    # Top band
    box(s, PL, Inches(0.35), IW, Inches(1.05), YELLOW, BLACK)
    tb(
        s,
        PL + Inches(0.2),
        Inches(0.5),
        IW - Inches(0.4),
        Inches(0.75),
        [("艺术手法 · TYPOGRAPHY", 26, True, BLACK)],
        PP_ALIGN.LEFT,
        MSO_ANCHOR.MIDDLE,
    )
    # Bottom split line
    box(s, PL, Inches(1.45), IW, Pt(2), BLACK, None)
    # Left hero vertical label
    tb(
        s,
        PL,
        Inches(1.65),
        Inches(1.0),
        Inches(5.2),
        [("手", 48, True, BLACK), ("法", 48, True, BLACK), ("四", 48, True, BLACK), ("式", 48, True, BLACK)],
        PP_ALIGN.CENTER,
    )
    # Waterfall flow — 4 stacked boxes, each offset right
    items = [
        ("01  夸张", "「除非为了炫耀新衣服」——爱新衣写到极致"),
        ("02  反讽", "说「看见了」其实什么也没看见"),
        ("03  对比", "大人怯懦  vs  孩子天真诚实"),
        ("04  反复", "「我什么也没有看见」强化讽刺"),
    ]
    y = Inches(1.65)
    x_off = [Inches(1.15), Inches(1.55), Inches(1.95), Inches(2.35)]
    for i, (title, desc) in enumerate(items):
        w = Inches(9.8) - x_off[i]
        box(s, PL + x_off[i], y, w, Inches(1.15), WHITE if i % 2 else YELLOW, BLACK)
        tb(s, PL + x_off[i] + Inches(0.15), y + Inches(0.12), w - Inches(0.3), Inches(0.9), [(title, 16, True, BLACK), (desc, 13, False, BLACK)])
        y += Inches(1.28)


# ── Slide 4: Hero = 贯穿左右的变化时间轴 ──
def s04(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    slide_num(s, 4)
    tb(s, PL, Inches(0.35), IW, Inches(0.5), [("故事推演脉络", 28, True, BLACK)])
    steps = ["骗子做衣", "大臣探访", "皇帝试穿", "游行大典", "小孩揭穿"]
    n = len(steps)
    step_w = (IW - Inches(0.4)) / n
    y_line = Inches(2.8)
    # Timeline bar
    box(s, PL, y_line, IW, Pt(2), BLACK, None)
    for i, st in enumerate(steps):
        x = PL + i * step_w
        # node
        box(s, x + step_w * 0.35, y_line - Inches(0.12), Inches(0.35), Inches(0.35), YELLOW if i == 4 else WHITE, BLACK)
        tb(s, x, y_line - Inches(1.5), step_w, Inches(1.2), [(st, 14, True, BLACK)], PP_ALIGN.CENTER)
        tb(s, x, y_line + Inches(0.25), step_w, Inches(0.35), [(f"0{i+1}", 11, True, BLACK)], PP_ALIGN.CENTER)
    # Hero arrow label
    box(s, PL, Inches(4.2), IW, Inches(1.35), WHITE, BLACK)
    tb(
        s,
        PL + Inches(0.2),
        Inches(4.35),
        IW - Inches(0.4),
        Inches(1.05),
        [
            ("核心矛盾", 14, True, BLACK),
            ("虚荣心  VS  真实事实", 22, True, BLACK),
            ("谎言因「承认看不见」运转  ·  真话因「孩子无畏」刺破", 14, False, BLACK),
        ],
    )
    tb(s, PL + Inches(9.5), Inches(1.0), Inches(2.5), Inches(1.2), [("5", 120, True, YELLOW)], PP_ALIGN.RIGHT)


# ── Slide 5: Hero = 巨型「03」人物数量 ──
def s05(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    slide_num(s, 5)
    tb(s, PL, Inches(0.35), Inches(6.0), Inches(0.5), [("人物图鉴", 28, True, BLACK)])
    tb(s, PL + Inches(8.5), Inches(0.2), Inches(3.5), Inches(1.8), [("03", 110, True, BLACK)], PP_ALIGN.RIGHT)
    chars = [
        ("皇帝", "极致虚荣、愚蠢自私", "怕被认为不聪明、不称职"),
        ("大臣们", "自保阿谀、盲从从众", "怕失去官位与恩宠"),
        ("骗子", "狡黠敏锐、利用弱点", "精准猎杀人性虚荣"),
    ]
    w = Inches(3.85)
    x = PL
    for i, (name, trait, psyche) in enumerate(chars):
        fill = YELLOW if i == 1 else WHITE
        box(s, x, Inches(1.35), w, Inches(5.5), fill, BLACK)
        tb(s, x + Inches(0.15), Inches(1.55), w - Inches(0.3), Inches(0.7), [(name, 24, True, BLACK)])
        box(s, x + Inches(0.15), Inches(2.35), w - Inches(0.3), Pt(2), BLACK, None)
        tb(
            s,
            x + Inches(0.15),
            Inches(2.55),
            w - Inches(0.3),
            Inches(4.0),
            [
                ("性格", 12, True, BLACK),
                (trait, 14, False, BLACK),
                ("", 6, False, BLACK),
                ("心理", 12, True, BLACK),
                (psyche, 14, False, BLACK),
            ],
        )
        x += w + Inches(0.2)


# ── Slide 6: Hero = 巨型「？」试金石 ──
def s06(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    slide_num(s, 6)
    # Asymmetric: giant ? left
    tb(s, PL, Inches(0.8), Inches(3.5), Inches(3.5), [("?", 200, True, YELLOW)], PP_ALIGN.CENTER)
    tb(s, PL + Inches(3.8), Inches(0.35), Inches(8.5), Inches(0.55), [("关键道具 · 新装", 26, True, BLACK)])
    tb(
        s,
        PL + Inches(3.8),
        Inches(0.95),
        Inches(8.5),
        Inches(0.55),
        [("不存在的衣服，何以成为人性的「试金石」？", 16, False, BLACK)],
    )
    box(s, PL + Inches(3.8), Inches(1.65), Inches(4.1), Inches(4.8), WHITE, BLACK)
    tb(
        s,
        PL + Inches(4.0),
        Inches(1.85),
        Inches(3.7),
        Inches(4.4),
        [
            ("特征", 14, True, BLACK),
            ("· 愚蠢或不称职的人看不见", 13, False, BLACK),
            ("· 检验聪明与称职", 13, False, BLACK),
            ("· 以看不见制造恐惧", 13, False, BLACK),
        ],
    )
    box(s, PL + Inches(8.1), Inches(2.35), Inches(4.2), Inches(4.1), YELLOW, BLACK)
    tb(
        s,
        PL + Inches(8.3),
        Inches(2.55),
        Inches(3.8),
        Inches(3.7),
        [
            ("实质", 14, True, BLACK),
            ("· 权力压迫", 13, False, BLACK),
            ("· 群体谎言", 13, False, BLACK),
            ("· 虚荣与怯懦的遮羞布", 13, True, BLACK),
        ],
    )


# ── Slide 7: Hero = 贯穿左右的 ADULT / CHILD 分割条 ──
def s07(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    slide_num(s, 7)
    tb(s, PL, Inches(0.35), IW, Inches(0.55), [("为什么只有小孩敢说真话？", 24, True, BLACK)], PP_ALIGN.CENTER)
    # Hero split bar
    box(s, PL, Inches(1.15), IW * 0.48, Inches(4.8), WHITE, BLACK)
    box(s, PL + IW * 0.52, Inches(2.05), IW * 0.48, Inches(3.9), YELLOW, BLACK)
    tb(
        s,
        PL + Inches(0.2),
        Inches(1.4),
        IW * 0.44,
        Inches(1.0),
        [("成人的世界", 32, True, BLACK)],
    )
    tb(
        s,
        PL + Inches(0.2),
        Inches(2.6),
        IW * 0.44,
        Inches(3.0),
        [
            ("· 害怕被嘲笑", 16, False, BLACK),
            ("· 维护虚荣与体面", 16, False, BLACK),
            ("· 沉默成为谎言帮凶", 16, True, BLACK),
            ("", 8, False, BLACK),
            ("「我什么也没有看见」", 14, False, BLACK),
        ],
    )
    tb(
        s,
        PL + IW * 0.54,
        Inches(2.35),
        IW * 0.42,
        Inches(1.0),
        [("孩子的世界", 32, True, BLACK)],
    )
    tb(
        s,
        PL + IW * 0.54,
        Inches(3.55),
        IW * 0.42,
        Inches(2.2),
        [
            ("· 无畏说真话", 16, False, BLACK),
            ("· 天真实诚", 16, False, BLACK),
            ("· 「他什么衣服也没穿呀」", 16, True, BLACK),
        ],
    )
    box(s, PL + Inches(5.5), Inches(5.2), Inches(2.3), Inches(0.9), BLACK, None)
    tb(s, PL + Inches(5.5), Inches(5.25), Inches(2.3), Inches(0.8), [("VS", 28, True, WHITE)], PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)


# ── Slide 8: Hero = 巨型百分比/逻辑链「100%」 ──
def s08(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    slide_num(s, 8)
    tb(s, PL, Inches(0.35), Inches(5.0), Inches(0.5), [("想象力的构建", 28, True, BLACK)])
    tb(s, PL + Inches(8.0), Inches(0.15), Inches(4.0), Inches(1.5), [("100%", 72, True, YELLOW)], PP_ALIGN.RIGHT)
    rows = [
        ("现实基础", "社会上的虚伪现象与谄媚风气", False),
        ("夸张想象", "将「虚伪」实体化为「不存在的衣服」", True),
        ("逻辑", "人人害怕承认 → 谎言完美运转", False),
    ]
    y = Inches(1.1)
    for title, desc, hi in rows:
        w = Inches(10.5) if hi else Inches(9.0)
        box(s, PL, y, w, Inches(1.55), YELLOW if hi else WHITE, BLACK)
        tb(s, PL + Inches(0.2), y + Inches(0.15), Inches(2.5), Inches(1.2), [(title, 18, True, BLACK)])
        tb(s, PL + Inches(2.8), y + Inches(0.35), w - Inches(3.0), Inches(1.0), [(desc, 15, False, BLACK)])
        y += Inches(1.75)
    box(s, PL, Inches(6.35), IW, Inches(0.75), BLACK, None)
    tb(
        s,
        PL + Inches(0.2),
        Inches(6.45),
        IW - Inches(0.4),
        Inches(0.55),
        [("想象源于现实，夸张服务于讽刺。", 18, True, WHITE)],
        PP_ALIGN.CENTER,
        MSO_ANCHOR.MIDDLE,
    )


# ── Slide 9: Hero = 巨型「200」字数 ──
def s09(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    slide_num(s, 9)
    tb(s, PL + Inches(7.5), Inches(0.2), Inches(4.5), Inches(2.2), [("200", 140, True, YELLOW)], PP_ALIGN.RIGHT)
    tb(s, PL, Inches(0.35), Inches(6.5), Inches(0.55), [("写作任务", 32, True, BLACK)])
    box(s, PL, Inches(1.05), Inches(7.2), Inches(5.8), WHITE, BLACK)
    tb(
        s,
        PL + Inches(0.25),
        Inches(1.25),
        Inches(6.7),
        Inches(5.4),
        [
            ("任务", 14, True, BLACK),
            ("为《皇帝的新装》续写一个 200 字结尾。", 16, True, BLACK),
            ("", 10, False, BLACK),
            ("提示", 14, True, BLACK),
            ("游行结束后，皇帝回宫会发生什么？", 14, False, BLACK),
            ("大臣们会怎么做？百姓如何议论？", 14, False, BLACK),
            ("", 10, False, BLACK),
            ("要求", 14, True, BLACK),
            ("符合人物性格逻辑，具有讽刺效果。", 14, False, BLACK),
        ],
    )
    box(s, PL + Inches(7.5), Inches(2.8), Inches(4.9), Inches(4.05), YELLOW, BLACK)
    tb(
        s,
        PL + Inches(7.7),
        Inches(3.0),
        Inches(4.5),
        Inches(3.7),
        [
            ("WORDS", 14, True, BLACK),
            ("", 6, False, BLACK),
            ("续写", 36, True, BLACK),
            ("", 6, False, BLACK),
            ("讽刺", 36, True, BLACK),
            ("", 6, False, BLACK),
            ("逻辑", 36, True, BLACK),
        ],
        PP_ALIGN.CENTER,
    )


def main():
    prs = Presentation()
    prs.slide_width = SW
    prs.slide_height = SH
    for fn in (s01, s02, s03, s04, s05, s06, s07, s08, s09):
        fn(prs)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(OUT))
    import shutil

    shutil.copy2(OUT, OUT_EN)
    print(f"✓ {OUT}")
    print(f"✓ {OUT_EN}")


if __name__ == "__main__":
    main()
