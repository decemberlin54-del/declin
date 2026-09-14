#!/usr/bin/env python3
"""Generate three-lesson Chinese poetry teaching deck with fade-in animations."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, List, Sequence, Tuple

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

# —— Design tokens ——
BG = RGBColor(0xF5, 0xF7, 0xFA)
ACCENT = RGBColor(0x1A, 0x6B, 0x7C)
ACCENT2 = RGBColor(0xC4, 0x5C, 0x26)
TEXT = RGBColor(0x2C, 0x3E, 0x50)
MUTED = RGBColor(0x5D, 0x6D, 0x7E)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
MARGIN = Inches(0.55)


@dataclass
class TextBlock:
    text: str
    left: float
    top: float
    width: float
    height: float
    font_size: int
    bold: bool = False
    color: RGBColor = TEXT
    align: PP_ALIGN = PP_ALIGN.LEFT
    line_spacing: float = 1.25


def get_spid(shape) -> int:
    el = shape._element
    cnv = el.find(f".//{{{P_NS}}}cNvPr")
    if cnv is None:
        cnv = el.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr")
    return int(cnv.get("id"))


def set_slide_bg(slide, color: RGBColor) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_accent_bar(slide, top_in: float, height_in: float = 0.08) -> None:
    bar = slide.shapes.add_shape(
        1, MARGIN, Inches(top_in), SLIDE_W - MARGIN * 2, Inches(height_in)
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = ACCENT
    bar.line.fill.background()


def add_textbox(slide, block: TextBlock):
    shape = slide.shapes.add_textbox(
        Inches(block.left),
        Inches(block.top),
        Inches(block.width),
        Inches(block.height),
    )
    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    p = tf.paragraphs[0]
    p.text = block.text
    p.alignment = block.align
    p.line_spacing = block.line_spacing
    run = p.runs[0]
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(block.font_size)
    run.font.bold = block.bold
    run.font.color.rgb = block.color
    return shape


def add_slide_header(slide, lesson_tag: str, title: str, subtitle: str = "") -> List:
    """Returns animated shapes (title area blocks)."""
    animated = []
    tag = add_textbox(
        slide,
        TextBlock(
            lesson_tag,
            0.55,
            0.35,
            12.2,
            0.45,
            22,
            bold=True,
            color=ACCENT2,
        ),
    )
    animated.append(tag)
    add_accent_bar(slide, 0.82)
    t = add_textbox(
        slide,
        TextBlock(
            title,
            0.55,
            0.95,
            12.2,
            0.85,
            36,
            bold=True,
            color=ACCENT,
        ),
    )
    animated.append(t)
    if subtitle:
        s = add_textbox(
            slide,
            TextBlock(
                subtitle,
                0.55,
                1.75,
                12.2,
                0.55,
                24,
                color=MUTED,
            ),
        )
        animated.append(s)
    return animated


def _set_vis(spid: int, val: str, delay: int, nid) -> str:
    cid = nid()
    return (
        f'<p:set xmlns:p="{P_NS}">'
        f'<p:cBhvr>'
        f'<p:cTn id="{cid}" dur="1" fill="hold">'
        f'<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst>'
        f"</p:cTn>"
        f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
        f"<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>"
        f"</p:cBhvr>"
        f'<p:to><p:strVal val="{val}"/></p:to>'
        f"</p:set>"
    )


def _anim_fade_in(spid: int, dur: int, nid) -> str:
    cid = nid()
    return (
        f'<p:animEffect transition="in" filter="fade" xmlns:p="{P_NS}">'
        f"<p:cBhvr>"
        f'<p:cTn id="{cid}" dur="{dur}" fill="hold"/>'
        f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
        f"</p:cBhvr>"
        f"</p:animEffect>"
    )


def _effect_par_fade_in(spid: int, node_type: str, dur: int, nid) -> str:
    par_id = nid()
    body = _set_vis(spid, "visible", 0, nid) + _anim_fade_in(spid, dur, nid)
    return (
        f'<p:par xmlns:p="{P_NS}">'
        f'<p:cTn id="{par_id}" presetID="10" presetClass="entr" presetSubtype="0" '
        f'fill="hold" grpId="0" nodeType="{node_type}">'
        f'<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
        f"<p:childTnLst>{body}</p:childTnLst>"
        f"</p:cTn>"
        f"</p:par>"
    )


def build_fade_timing(shape_ids: Sequence[int], dur_ms: int = 750) -> etree._Element:
    """One click per shape: fade-in entrance."""
    counter = [0]

    def nid() -> int:
        counter[0] += 1
        return counter[0]

    nid()  # tmRoot = 1
    nid()  # mainSeq = 2

    groups_xml = []
    for spid in shape_ids:
        effect = _effect_par_fade_in(spid, "clickEffect", dur_ms, nid)
        groups_xml.append(
            f'<p:par xmlns:p="{P_NS}">'
            f'<p:cTn id="{nid()}" fill="hold">'
            f'<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
            f"<p:childTnLst>{effect}</p:childTnLst>"
            f"</p:cTn>"
            f"</p:par>"
        )

    xml = (
        f'<p:timing xmlns:p="{P_NS}" xmlns:a="{A_NS}">'
        f"<p:tnLst>"
        f'<p:par>'
        f'<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
        f"<p:childTnLst>"
        f'<p:seq concurrent="1" nextAc="seek">'
        f'<p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
        f"<p:childTnLst>{''.join(groups_xml)}</p:childTnLst>"
        f"</p:cTn>"
        f"<p:prevCondLst>"
        f'<p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>'
        f"</p:prevCondLst>"
        f"<p:nextCondLst>"
        f'<p:cond evt="onClick" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>'
        f"</p:nextCondLst>"
        f"</p:seq>"
        f"</p:childTnLst>"
        f"</p:cTn>"
        f"</p:par>"
        f"</p:tnLst>"
        f"</p:timing>"
    )
    return etree.fromstring(xml.encode("utf-8"))


def attach_fade_animations(slide, shapes: Iterable) -> None:
    ids = [get_spid(s) for s in shapes]
    if not ids:
        return
    old = slide._element.find(f"{{{P_NS}}}timing")
    if old is not None:
        slide._element.remove(old)
    slide._element.append(build_fade_timing(ids))


def add_bullets_slide(
    prs,
    lesson_tag: str,
    title: str,
    bullets: Sequence[str],
    footer: str = "",
) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, BG)
    animated = add_slide_header(slide, lesson_tag, title)
    y = 2.45
    for line in bullets:
        box = add_textbox(
            slide,
            TextBlock(
                "▎ " + line,
                0.65,
                y,
                11.9,
                0.95,
                28,
                color=TEXT,
                line_spacing=1.2,
            ),
        )
        animated.append(box)
        y += 0.92
        if y > 6.5:
            break
    if footer and y <= 6.2:
        f = add_textbox(
            slide,
            TextBlock(
                footer,
                0.65,
                6.55,
                11.9,
                0.45,
                20,
                color=MUTED,
                align=PP_ALIGN.RIGHT,
            ),
        )
        animated.append(f)
    attach_fade_animations(slide, animated)


def add_cover_slide(
    prs,
    lesson_num: str,
    main_title: str,
    poem: str,
    positioning: str,
    focus_lines: Sequence[str],
) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, BG)
    animated = []
    ribbon = slide.shapes.add_shape(1, 0, Inches(2.55), SLIDE_W, Inches(2.35))
    ribbon.fill.solid()
    ribbon.fill.fore_color.rgb = ACCENT
    ribbon.line.fill.background()

    t1 = add_textbox(
        slide,
        TextBlock(
            f"第{lesson_num}课",
            0.55,
            0.45,
            3,
            0.5,
            24,
            bold=True,
            color=ACCENT2,
        ),
    )
    animated.append(t1)
    t2 = add_textbox(
        slide,
        TextBlock(
            main_title,
            0.55,
            2.75,
            12.2,
            1.0,
            44,
            bold=True,
            color=WHITE,
            align=PP_ALIGN.CENTER,
        ),
    )
    animated.append(t2)
    t3 = add_textbox(
        slide,
        TextBlock(
            poem,
            0.55,
            3.75,
            12.2,
            0.55,
            30,
            color=RGBColor(0xE8, 0xF4, 0xF6),
            align=PP_ALIGN.CENTER,
        ),
    )
    animated.append(t3)
    t4 = add_textbox(
        slide,
        TextBlock(
            positioning,
            0.75,
            5.05,
            11.8,
            0.9,
            26,
            color=TEXT,
            align=PP_ALIGN.CENTER,
            line_spacing=1.35,
        ),
    )
    animated.append(t4)
    y = 6.05
    for line in focus_lines:
        tb = add_textbox(
            slide,
            TextBlock(
                "◆ " + line,
                1.0,
                y,
                11.3,
                0.42,
                22,
                color=MUTED,
                align=PP_ALIGN.CENTER,
            ),
        )
        animated.append(tb)
        y += 0.38
    attach_fade_animations(slide, animated)


def lesson1(prs: Presentation) -> None:
    tag = "第一课 · 次北固山下"
    add_cover_slide(
        prs,
        "一",
        "风吹进新旧交替的江流",
        "《次北固山下》——对偶与理趣",
        "在江河行舟的风浪与新旧时光交替中，体悟自然规律与游子思乡。",
        [
            "律诗结构识别 · 颔联颈联对偶",
            "关键动词/形容词炼字 · 自然理趣与情感表达",
        ],
    )
    add_bullets_slide(
        prs,
        tag,
        "板块一 · 吹起江流的阵风（行舟视角与空间构图）",
        [
            "律诗结构：识别首联、颔联，掌握五言律诗押韵规则。",
            "对偶修辞：颔联词性相对（名词对名词，形容词对形容词）。",
            "炼字：“阔”——潮水涨平、江面开阔的壮阔视界。",
            "炼字：“悬”——风正顺风、帆高高悬挂的静态张力。",
            "细读：潮平/两岸/阔，风正/一帆/悬——梳理“客路”与“青山绿水”的空间关系。",
            "答题：形容词/动词表现力类——先释义，再联系画面，最后点情感。",
        ],
    )
    add_bullets_slide(
        prs,
        tag,
        "板块二 · 吹开新旧的风（自然理趣与时光交替）",
        [
            "名句：海日生残夜，江春入旧年。",
            "拟人与炼字：“生”“入”赋予自然以主观意图与生命力。",
            "自然理趣：旧事物中孕育新事物，新旧交替不可阻挡。",
            "对偶：颈联工整——“残夜↔海日”“旧年↔江春”对照。",
            "“生”：海日破空而出的动态；“入”：江南春意闯入旧年的势头。",
            "答题：诗中蕴含哲理——摘句 + 释义 + 概括规律 + 联系主旨。",
        ],
    )
    add_bullets_slide(
        prs,
        tag,
        "板块三 · 吹向远方的江流（意象传递与思乡主旨）",
        [
            "尾联：乡书何处达？归雁洛阳边。",
            "设问修辞 + 借物传书：一问一答，直抒胸臆。",
            "意象“归雁”：北方来雁、春归北飞，象征思乡与信息传递。",
            "收束全诗：由江行写景 → 理趣升华 → 点明思乡主旨。",
            "联想：风送江流、雁向北飞，触发“风送家书”的游子之思。",
            "答题：分析尾联情感与作用——手法 + 内容 + 结构 + 情感。",
        ],
    )
    add_bullets_slide(
        prs,
        tag,
        "本课考点 · 答题规范速记",
        [
            "结构题：分联标名 + 押韵字 + 对仗位置（颔联、颈联）。",
            "炼字题：字义 → 所写之景 → 所见之情/理。",
            "哲理题：抓“生/入”等动词，写“旧→新”的不可阻规律。",
            "尾联题：设问/意象 + 思乡主旨 + 与前文景物照应。",
            "朗读提示：重读“阔”“悬”“生”“入”，体会由壮景到理趣再到乡愁。",
        ],
        footer="次北固山下 · 王湾",
    )


def lesson2(prs: Presentation) -> None:
    tag = "第二课 · 闻王昌龄左迁龙标遥有此寄"
    add_cover_slide(
        prs,
        "二",
        "月亮替我去看你",
        "《闻王昌龄左迁龙标遥有此寄》——起兴与寄情",
        "跨越千山万水的阻隔，将无形情感托付给月光，表达对友人遭贬的深切牵挂。",
        [
            "七言绝句结构 · 起兴手法 · 典型意象",
            "拟人修辞 · 借景抒情与托物寄情",
        ],
    )
    add_bullets_slide(
        prs,
        tag,
        "板块一 · 遥寄牵挂的序曲（暮春意象与贬谪背景）",
        [
            "体裁：七言绝句（四句二十八字，押平声韵）。",
            "起兴：触景生情，先景后情，烘托气氛。",
            "杨花：漂泊无定、离散流离；子规：啼声“不如归去”，渲染悲凉。",
            "词语：左迁 = 古代贬官的委婉说法。",
            "诵读：杨花落尽子规啼——“落尽”写暮春衰败，“啼”写声响凄清。",
            "为何开篇不直写闻讯？以暮春荒凉起兴，映友人贬谪龙标之远与内心之悲。",
        ],
    )
    add_bullets_slide(
        prs,
        tag,
        "板块二 · 月亮替我去看你（拟人修辞与化虚为实）",
        [
            "名句：我寄愁心与明月。",
            "拟人：明月具人的知觉与情感，可承接“愁心”。",
            "借景抒情 / 托物寄情：抽象“愁心”化为可“寄”的实体。",
            "炼字“寄”：主动交付，打破空间阻隔。",
            "为何选明月？高悬天空、万里共照——与友人虽远犹同在一轮月下。",
            "答题：修辞 + 分析字词 + 情感效果（牵挂、安慰、无奈中的深情）。",
        ],
    )
    add_bullets_slide(
        prs,
        tag,
        "板块三 · 万里同辉的告白（空间跨越与主旨收束）",
        [
            "结句：随风直到夜郎西。",
            "空间跨越：从诗人所在地延伸至夜郎西（龙标一带）。",
            "“随”“直”：坚定、执着，月光一路陪伴友人。",
            "主旨：对友人遭贬的同情、关切与深厚友谊。",
            "艺术：不直接劝慰，而让月光“送”到贬所——情感更绵长含蓄。",
            "答题：尾联作用——拓展意境 + 强化情感 + 收束全诗。",
        ],
    )
    add_bullets_slide(
        prs,
        tag,
        "本课考点 · 答题规范速记",
        [
            "起兴题：写了什么景 → 营造什么氛围 → 为何情做铺垫。",
            "意象题：杨花/子规/明月各答文化内涵 + 与情感的关系。",
            "拟人题：指出赋予人的特征 + 表达效果。",
            "炼字“寄”：交付、传递 + 化虚为实 + 突破空间。",
            "主旨题：贬谪背景 + 牵挂友情 + 浪漫寄托方式。",
        ],
        footer="闻王昌龄左迁龙标遥有此寄 · 李白",
    )


def lesson3(prs: Presentation) -> None:
    tag = "第三课 · 天净沙·秋思"
    add_cover_slide(
        prs,
        "三",
        "夕阳照不亮归途",
        "《天净沙·秋思》——白描与反衬",
        "夕阳西下的羁旅途中，九个物象组合与对比，呈现游子无家可归的孤独与思乡之痛。",
        [
            "元曲小令体裁 · 纯名词白描",
            "乐景衬哀（反衬）· 羁旅思乡主旨",
        ],
    )
    add_bullets_slide(
        prs,
        tag,
        "板块一 · 九景并置的剪影（物象堆叠与白描画面）",
        [
            "体裁：元曲小令（曲牌《天净沙》，题目《秋思》）。",
            "手法：纯名词并置、白描——不用动词串联，直接组画。",
            "九名词：枯藤、老树、昏鸦、小桥、流水、人家、古道、西风、瘦马。",
            "前三句：枯藤老树昏鸦——衰败、苍凉的秋日气氛。",
            "白描质感：名词排列即画面，读者自行补全动态与情感。",
            "答题：白描 + 意象选择 + 氛围（萧瑟、苍凉、羁旅）。",
        ],
    )
    add_bullets_slide(
        prs,
        tag,
        "板块二 · 别人家与路上一人（乐景反衬与氛围对比）",
        [
            "对比组：小桥流水人家 ↔ 古道西风瘦马。",
            "反衬（以乐景衬哀景）：人家温馨安宁，反衬游子漂泊无依。",
            "炼字“瘦”：路途遥远、羁旅艰辛、游子憔悴形态。",
            "冷暖调性：暖（人家）与冷（古道西风）并置，强化孤独。",
            "为何嵌入“人家”？以他人安居反衬自己无家可归。",
            "答题：对比/反衬——两幅画面 + 突出什么情感。",
        ],
    )
    add_bullets_slide(
        prs,
        tag,
        "板块三 · 夕阳照不亮的归途（时间节点与主旨点睛）",
        [
            "结句：夕阳西下，断肠人在天涯。",
            "“夕阳”：日暮思归的特定心理时刻（该归而不能归）。",
            "直抒胸臆：“断肠人”点明极度思乡、悲伤至极的漂泊者。",
            "由景入情：夕阳余晖照古道，却照不亮游子的归家之路。",
            "主旨：羁旅思乡——天涯孤旅，有家难归。",
            "答题：结句作用——时间意象 + 点明人物 + 揭示主题。",
        ],
    )
    add_bullets_slide(
        prs,
        tag,
        "本课考点 · 答题规范速记",
        [
            "体裁题：小令、曲牌与题目、句数特点。",
            "白描题：名词罗列 + 组画方式 + 氛围效果。",
            "意象题：九个意象分类（衰败/旅途/温馨/羁旅）+ 整体基调。",
            "反衬题：乐景（人家）写哀情（游子）+ 对比强化孤独。",
            "主旨题：“断肠人在天涯”——羁旅 + 思乡 + 无家可归之痛。",
        ],
        footer="天净沙·秋思 · 马致远",
    )


def remove_all_timing(prs: Presentation) -> None:
    for slide in prs.slides:
        old = slide._element.find(f"{{{P_NS}}}timing")
        if old is not None:
            slide._element.remove(old)


def main() -> None:
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    lesson1(prs)
    lesson2(prs)
    lesson3(prs)

    animated_path = "/workspace/gushi-lessons-animated.pptx"
    prs.save(animated_path)

    remove_all_timing(prs)
    stable_paths = (
        "/workspace/gushi-lessons.pptx",
        "/workspace/古诗三课教学课件.pptx",
    )
    for path in stable_paths:
        prs.save(path)

    n = len(prs.slides)
    print(f"Saved stable (no animation): {stable_paths[0]} ({n} slides)")
    print(f"Saved animated: {animated_path} ({n} slides)")


if __name__ == "__main__":
    main()
