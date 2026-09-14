#!/usr/bin/env python3
"""Generate 10-slide Morandi PPT: 《盲孩子和他的影子》与单元写作."""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

OUT = Path("/workspace/第六单元PPT/第4课 盲孩子和他的影子与单元写作.pptx")
OUT_EN = Path("/workspace/unit6-lesson4-blind-child-writing.pptx")

SW, SH = Inches(13.333), Inches(7.5)
PL, PT, PR = Inches(0.625), Inches(0.42), Inches(0.625)
IW = SW - PL - PR
TOTAL = 10

BG = RGBColor(0xF4, 0xF1, 0xEA)
CARD = RGBColor(0xFF, 0xFF, 0xFF)
TEXT = RGBColor(0x2B, 0x2D, 0x42)
MUTED = RGBColor(0x6C, 0x75, 0x7D)
TEAL = RGBColor(0x6B, 0x90, 0x80)
PINK = RGBColor(0xB5, 0x83, 0x8D)
OCHRE = RGBColor(0xE0, 0x9F, 0x68)
BORDER = RGBColor(0xE2, 0xE2, 0xD9)


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


def header(slide, title, n):
    ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, PL, PT + Inches(0.55), IW, Inches(0.02))
    ln.fill.solid()
    ln.fill.fore_color.rgb = BORDER
    ln.line.fill.background()
    tb(slide, PL, PT, IW * 0.78, Inches(0.48), [(title, 26, True, TEAL)])
    tb(slide, PL + IW * 0.78, PT, IW * 0.22, Inches(0.48), [(f"{n:02d} / {TOTAL}", 14, True, MUTED)], PP_ALIGN.RIGHT)


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


def card(slide, l, t, w, h, title, body, title_sz=18, body_sz=14, accent=TEAL):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = CARD
    sh.line.color.rgb = BORDER
    sh.line.width = Pt(1)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l + Inches(0.12), t + Inches(0.18), Inches(0.05), Inches(0.28))
    bar.fill.solid()
    bar.fill.fore_color.rgb = OCHRE
    bar.line.fill.background()
    tb(slide, l + Inches(0.22), t + Inches(0.12), w - Inches(0.3), Inches(0.38), [(title, title_sz, True, accent)])
    tb(slide, l + Inches(0.22), t + Inches(0.52), w - Inches(0.3), h - Inches(0.65), [(body, body_sz, False, TEXT)])


def add_table(slide, l, t, w, h, headers, rows, col_widths=None):
    tbl = slide.shapes.add_table(len(rows) + 1, len(headers), l, t, w, h).table
    if col_widths:
        for i, cw in enumerate(col_widths):
            tbl.columns[i].width = cw
    for i, htxt in enumerate(headers):
        cell = tbl.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = BG
        p = cell.text_frame.paragraphs[0]
        r = p.add_run()
        font(r, 12, True, TEAL)
        r.text = htxt
    for ri, row in enumerate(rows):
        for ci, txt in enumerate(row):
            cell = tbl.cell(ri + 1, ci)
            p = cell.text_frame.paragraphs[0]
            r = p.add_run()
            font(r, 11, ci == 0, TEXT)
            r.text = txt
    return tbl


def s01(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    pill = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, SW / 2 - Inches(2.8), Inches(1.6), Inches(5.6), Inches(0.45))
    pill.fill.solid()
    pill.fill.fore_color.rgb = PINK
    pill.line.fill.background()
    tb(s, SW / 2 - Inches(2.8), Inches(1.65), Inches(5.6), Inches(0.38), [("七年级语文下册第六单元 · 第4课", 13, True, CARD)], PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    tb(s, PL, Inches(2.5), IW, Inches(0.9), [("《盲孩子和他的影子》与单元写作", 36, True, TEXT)], PP_ALIGN.CENTER)
    tb(s, PL, Inches(3.5), IW, Inches(0.55), [("爱与光明的隐喻 & 从阅读到写作的迁移", 20, False, TEAL)], PP_ALIGN.CENTER)


def s02(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "故事脉络：从黑暗走向光明", 2)
    items = [
        ("1. 寂寞与黑暗", "盲孩子独自生活在无光的世界里，内心孤独。"),
        ("2. 影子的陪伴", "影子有了生命，带来欢声笑语与温情。"),
        ("3. 萤火虫的光芒", "微小力量的汇聚，驱散狂风暴雨中的恐慌。"),
        ("4. 重见光明", "盲孩子看见世界，影子也化作了真实的朋友。"),
    ]
    cw = (IW - Inches(0.45)) / 4
    x = PL
    for title, body in items:
        card(s, x, Inches(1.0), cw, Inches(5.6), title, body, 15, 13)
        x += cw + Inches(0.15)


def s03(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "核心意象的隐喻解读", 3)
    items = [
        ("意象 A：影子", "象征着纯洁的友谊、形影不离的真心陪伴与关爱。"),
        ("意象 B：萤火虫", "象征着社会中微小却坚韧的善意，是微光汇聚的力量。"),
        ("意象 C：光明", "象征着内心的希望、幸福的重塑以及爱的终极奇迹。"),
    ]
    cw = (IW - Inches(0.3)) / 3
    x = PL
    for title, body in items:
        card(s, x, Inches(1.0), cw, Inches(5.6), title, body)
        x += cw + Inches(0.15)


def s04(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "诗意与童话的融合", 4)
    cw = (IW - Inches(0.2)) / 2
    card(
        s,
        PL,
        Inches(1.0),
        cw,
        Inches(5.6),
        "语言形式的美感",
        "• 大量使用单句与短句，形成如流水的韵律。\n• 频繁重叠词汇（「轻轻的」「痒痒的」），充满童趣。\n• 犹如一首抒情散文诗。",
    )
    card(
        s,
        PL + cw + Inches(0.2),
        Inches(1.0),
        cw,
        Inches(5.6),
        "感官描写的细腻",
        "• 盲孩子失去视觉，但作者强化了听觉、触觉与嗅觉。\n• 风声、虫鸣、阳光的温度，丰富了想象的感知层次。",
    )


def s05(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "单元总结：想象力构建的模型", 5)
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PL, Inches(1.0), IW, Inches(5.6))
    sh.fill.solid()
    sh.fill.fore_color.rgb = CARD
    sh.line.color.rgb = BORDER
    sh.line.width = Pt(1)
    tb(
        s,
        PL,
        Inches(1.35),
        IW,
        Inches(0.6),
        [("想象力 = 现实依托 + 逻辑推演 + 新奇设想", 24, True, TEAL)],
        PP_ALIGN.CENTER,
    )
    cw = (IW - Inches(0.6)) / 2
    card(s, PL + Inches(0.15), Inches(2.3), cw, Inches(3.8), "《皇帝的新装》", "虚荣心/谄媚风气 + 衣服检验法则 + 皇帝裸奔游行")
    card(s, PL + cw + Inches(0.45), Inches(2.3), cw, Inches(3.8), "《盲孩子和他的影子》", "残障者的孤独 + 影子/萤火虫拟人化 + 光明的奇迹")


def s06(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "单元写作任务发布", 6)
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PL, Inches(1.0), IW, Inches(5.6))
    sh.fill.solid()
    sh.fill.fore_color.rgb = CARD
    sh.line.color.rgb = BORDER
    tb(s, PL + Inches(0.25), Inches(1.15), IW - Inches(0.5), Inches(0.4), [("写作题目选一", 18, True, TEAL)])
    tb(
        s,
        PL + Inches(0.25),
        Inches(1.65),
        IW - Inches(0.5),
        Inches(4.5),
        [
            ("1. 《假如我有一双翅膀》（侧重现实意义与理想追求）", 16, False, TEXT),
            ("2. 《十年后的我》（基于个人发展的合理推演）", 16, False, TEXT),
            ("3. 自选主题虚构故事（须包含明确的现实依托与新奇情节）", 16, False, TEXT),
            ("", 8, False, TEXT),
            ("写作要求：字数 500 字以上，逻辑严密，具有文学美感。", 16, True, TEXT),
        ],
    )


def s07(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "写作脚手架：想象说明卡", 7)
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PL, Inches(1.0), IW, Inches(5.6))
    sh.fill.solid()
    sh.fill.fore_color.rgb = CARD
    sh.line.color.rgb = BORDER
    tb(s, PL + Inches(0.25), Inches(1.15), IW - Inches(0.5), Inches(0.35), [("动笔前的思考模板", 18, True, TEAL)])
    add_table(
        s,
        PL + Inches(0.25),
        Inches(1.65),
        IW - Inches(0.5),
        Inches(4.5),
        ["维度", "说明", "填写示例（以《假如我有一双翅膀》为例）"],
        [
            ["1. 现实依托", "生活中的真实现象/困境", "看到鸟类迁徙被网捕，或者城市交通拥堵。"],
            ["2. 逻辑推演", "有了神奇设想后的发展因果", "有了翅膀 → 飞向高空 → 发现鸟类生存困境 → 展开守护。"],
            ["3. 新奇点", "突破常规的创新元素", "翅膀不是羽毛做的，而是由「透明的空气净化流体」构成。"],
        ],
        [Inches(1.4), Inches(2.8), Inches(7.5)],
    )


def s08(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "中考导向：想象作文 3 维评价量表", 8)
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PL, Inches(1.0), IW, Inches(5.6))
    sh.fill.solid()
    sh.fill.fore_color.rgb = CARD
    sh.line.color.rgb = BORDER
    add_table(
        s,
        PL + Inches(0.2),
        Inches(1.2),
        IW - Inches(0.4),
        Inches(5.0),
        ["评价维度", "权重", "低分段（1-2分）", "中分段（3-4分）", "高分段（5分）"],
        [
            ["有依托", "40%", "凭空捏造，毫无现实基础", "有现实对应，但结合牵强", "现实依托清晰，观察细致"],
            ["有逻辑", "30%", "前后矛盾，情节突兀", "基本合逻辑，细节有小漏洞", "推演严密，合情合理"],
            ["有新奇", "30%", "套路化严重，老生常谈", "意料之外，尚在情理之中", "视角独特，设想新颖深刻"],
        ],
        [Inches(1.2), Inches(0.7), Inches(2.8), Inches(2.8), Inches(2.8)],
    )


def s09(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "佳作片段升格剖析", 9)
    cw = (IW - Inches(0.2)) / 2
    card(
        s,
        PL,
        Inches(1.0),
        cw,
        Inches(5.6),
        "修改前（缺乏依托与细节）",
        "「一天早上醒来，我发现自己突然长出了一双大翅膀。我开心极了，立刻飞上了天空，飞到了太空里看月球，玩得非常开心。」\n\n问题：凭空出现，无心理准备，无细节描述，不符合物理逻辑。",
        accent=PINK,
    )
    card(
        s,
        PL + cw + Inches(0.2),
        Inches(1.0),
        cw,
        Inches(5.6),
        "修改后（注入现实感与逻辑）",
        "「肩胛骨处传来一阵微微的酸胀感，就像春笋突破冻土。我试着张开双臂，一双由光线折射构成的透明羽翼在晨光中舒展。这不是魔法，而是科技对残障人士的重塑……」\n\n优势：细节真实，逻辑自洽，富有现代感。",
    )


def s10(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "课堂小结与作业部署", 10)
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PL, Inches(1.0), IW, Inches(5.6))
    sh.fill.solid()
    sh.fill.fore_color.rgb = CARD
    sh.line.color.rgb = BORDER
    tb(s, PL + Inches(0.3), Inches(1.25), IW - Inches(0.6), Inches(0.4), [("结语", 20, True, TEAL)])
    tb(
        s,
        PL + Inches(0.3),
        Inches(1.75),
        IW - Inches(0.6),
        Inches(1.6),
        [
            (
                "想象力是人类最宝贵的财富。希望同学们在写作时，时刻牢记：飞得再高、再远的翅膀，其根基永远深扎于我们对生活最真实的爱与观察中。",
                16,
                False,
                TEXT,
            )
        ],
    )
    tb(s, PL + Inches(0.3), Inches(3.55), IW - Inches(0.6), Inches(0.4), [("课后作业", 20, True, TEAL)])
    tb(
        s,
        PL + Inches(0.3),
        Inches(4.05),
        IW - Inches(0.6),
        Inches(1.5),
        [
            ("1. 填好《想象说明卡》。", 15, False, TEXT),
            ("2. 根据本课学习的量表，完成 500 字单元想象作文初稿。", 15, False, TEXT),
        ],
    )


def main():
    prs = Presentation()
    prs.slide_width = SW
    prs.slide_height = SH
    for fn in (s01, s02, s03, s04, s05, s06, s07, s08, s09, s10):
        fn(prs)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(OUT))
    import shutil

    shutil.copy2(OUT, OUT_EN)
    print(f"✓ {OUT}")
    print(f"✓ {OUT_EN}")


if __name__ == "__main__":
    main()
