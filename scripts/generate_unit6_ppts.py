#!/usr/bin/env python3
"""Generate Unit 6 lesson PPTs — bordered, centered, distinct styles per lesson."""

import math
import os
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

BASE = Path("/workspace/第六单元PPT")
IMG_DIR = BASE / "images"
OUT_DIR = BASE

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
FRAME_L, FRAME_T = Inches(0.32), Inches(0.30)
FRAME_W, FRAME_H = Inches(12.68), Inches(6.89)
CX = 6.667  # slide center x in inches


def rgb(t):
    return RGBColor(*t)


def set_run(run, text, size=28, bold=False, color=(60, 60, 60)):
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = "Microsoft YaHei"
    run.font.color.rgb = rgb(color)


def set_para(p, align=PP_ALIGN.CENTER):
    p.alignment = align
    p.space_after = Pt(10)


# ── Theme definitions ──────────────────────────────────────────────
THEMES = {
    "unit": {
        "name": "博物馆展板风",
        "bg": (235, 240, 248),
        "bg2": (200, 215, 235),
        "border": (100, 130, 170),
        "accent": (70, 120, 180),
        "accent2": (160, 190, 220),
        "panel": (248, 250, 255),
        "text": (50, 60, 80),
        "corner": "museum",
    },
    "wukong": {
        "name": "仙侠斗法风",
        "bg": (255, 240, 228),
        "bg2": (255, 200, 160),
        "border": (200, 90, 50),
        "accent": (220, 80, 40),
        "accent2": (255, 180, 100),
        "panel": (255, 248, 240),
        "text": (80, 40, 20),
        "corner": "cloud",
    },
    "emperor": {
        "name": "童话宫廷风",
        "bg": (245, 238, 252),
        "bg2": (210, 190, 230),
        "border": (130, 90, 160),
        "accent": (120, 70, 150),
        "accent2": (200, 170, 220),
        "panel": (252, 248, 255),
        "text": (60, 40, 80),
        "corner": "crown",
    },
    "nuwa": {
        "name": "神话创世风",
        "bg": (235, 245, 232),
        "bg2": (170, 200, 165),
        "border": (90, 140, 90),
        "accent": (70, 130, 80),
        "accent2": (160, 200, 150),
        "panel": (248, 252, 245),
        "text": (40, 70, 45),
        "corner": "leaf",
    },
    "fable": {
        "name": "故事书卷风",
        "bg": (255, 248, 228),
        "bg2": (240, 210, 150),
        "border": (180, 140, 60),
        "accent": (200, 150, 40),
        "accent2": (240, 200, 100),
        "panel": (255, 252, 240),
        "text": (90, 70, 30),
        "corner": "scroll",
    },
    "final": {
        "name": "多巴胺展览风",
        "bg": (255, 240, 250),
        "bg2": (200, 230, 255),
        "border": (180, 100, 180),
        "accent": (230, 80, 150),
        "accent2": (100, 200, 220),
        "panel": (255, 250, 255),
        "text": (70, 50, 90),
        "corner": "festive",
    },
}


def make_corner(kind, colors, size=(400, 400)):
    """Generate unique corner decoration per theme."""
    img = Image.new("RGBA", size, (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    c1, c2, c3 = colors["accent"], colors["accent2"], colors["border"]
    w, h = size

    if kind == "museum":
        for i in range(5):
            d.rectangle([10 + i * 8, 10 + i * 8, w - 30, h - 30], outline=(*c1, 180 - i * 30), width=2)
        d.rectangle([40, 40, 120, 90], fill=(*c2, 120))
        d.text((50, 50), "展", fill=(*c1, 200))

    elif kind == "cloud":
        for cx, cy, r in [(120, 80, 50), (200, 60, 40), (280, 90, 45)]:
            d.ellipse([cx - r, cy - r // 2, cx + r, cy + r // 2], fill=(*c2, 150))
        d.arc([20, 20, 200, 200], 0, 90, fill=(*c1, 200), width=4)

    elif kind == "crown":
        pts = [(60, 200), (100, 80), (140, 160), (180, 60), (220, 160), (260, 80), (300, 200)]
        d.polygon(pts, fill=(*c1, 160), outline=(*c3, 200))
        for x in [100, 180, 260]:
            d.ellipse([x - 12, 68, x + 12, 92], fill=(*c2, 200))

    elif kind == "leaf":
        d.pieslice([50, 30, 250, 280], 200, 340, fill=(*c1, 140))
        d.pieslice([100, 80, 300, 330], 160, 300, fill=(*c2, 120))
        d.line([(30, 350), (200, 200)], fill=(*c3, 180), width=3)

    elif kind == "scroll":
        d.rectangle([60, 80, 340, 280], fill=(*c2, 100), outline=(*c1, 180), width=3)
        d.ellipse([40, 120, 100, 240], fill=(*c1, 150))
        d.ellipse([300, 120, 360, 240], fill=(*c1, 150))
        for y in range(110, 260, 30):
            d.line([(110, y), (290, y)], fill=(*c3, 80), width=1)

    elif kind == "festive":
        fest = [(255, 100, 130), (100, 200, 230), (255, 200, 80), (180, 130, 230), (100, 220, 160)]
        for i, col in enumerate(fest):
            x = 50 + i * 55
            d.polygon([(x, 40), (x + 20, 100), (x + 40, 40)], fill=(*col, 180))
            d.ellipse([x + 10, 100, x + 30, 120], fill=(*col, 200))

    return img


def make_bg(theme_key, filename):
    t = THEMES[theme_key]
    path = IMG_DIR / filename
    IMG_DIR.mkdir(parents=True, exist_ok=True)
    if path.exists():
        return path
    img = Image.new("RGB", (1920, 1080), t["bg"])
    d = ImageDraw.Draw(img)
    for i in range(0, 1080, 3):
        ratio = i / 1080
        c = tuple(int(t["bg"][j] * (1 - ratio * 0.3) + t["bg2"][j] * ratio * 0.3) for j in range(3))
        d.line([(0, i), (1920, i)], fill=c)
    # subtle pattern
    for x in range(0, 1920, 80):
        for y in range(0, 1080, 80):
            d.ellipse([x, y, x + 3, y + 3], fill=(*t["accent2"],))
    img.save(path)
    return path


def save_corner(theme_key, pos):
    t = THEMES[theme_key]
    path = IMG_DIR / f"corner_{theme_key}_{pos}.png"
    if not path.exists():
        img = make_corner(t["corner"], t)
        if pos == "tr":
            img = img.transpose(Image.FLIP_LEFT_RIGHT)
        if pos == "bl":
            img = img.transpose(Image.FLIP_TOP_BOTTOM)
        if pos == "br":
            img = img.transpose(Image.FLIP_LEFT_RIGHT).transpose(Image.FLIP_TOP_BOTTOM)
        img.save(path)
    return path


class ThemedDeck:
    def __init__(self, theme_key):
        self.key = theme_key
        self.t = THEMES[theme_key]
        self.prs = Presentation()
        self.prs.slide_width = SLIDE_W
        self.prs.slide_height = SLIDE_H
        self.bg = make_bg(theme_key, f"bg_{theme_key}.png")
        for pos in ("tl", "tr"):
            save_corner(theme_key, pos)

    def _blank(self):
        return self.prs.slides.add_slide(self.prs.slide_layouts[6])

    def _base(self, slide):
        slide.shapes.add_picture(str(self.bg), 0, 0, SLIDE_W, SLIDE_H)
        self._frame(slide)
        for pos, l, t in [("tl", -0.05, -0.05), ("tr", 10.8, -0.05)]:
            p = IMG_DIR / f"corner_{self.key}_{pos}.png"
            if p.exists():
                slide.shapes.add_picture(str(p), Inches(l), Inches(t), Inches(2.2), Inches(2.2))

    def _frame(self, slide, double=False):
        for i, (lw, col, inset) in enumerate([
            (Pt(2.5), self.t["border"], 0),
            (Pt(1), self.t["accent2"], 0.06 if double else None),
        ]):
            if inset is None and i == 1:
                continue
            l = FRAME_L + Inches(inset or 0)
            t = FRAME_T + Inches(inset or 0)
            w = FRAME_W - Inches((inset or 0) * 2)
            h = FRAME_H - Inches((inset or 0) * 2)
            rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
            rect.fill.background()
            rect.line.color.rgb = rgb(col)
            rect.line.width = lw

    def _panel(self, slide, top, height, width=10.5):
        left = Inches(CX - width / 2)
        panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(top), Inches(width), Inches(height))
        panel.fill.solid()
        panel.fill.fore_color.rgb = rgb(self.t["panel"])
        panel.fill.transparency = 0.15
        panel.line.color.rgb = rgb(self.t["accent2"])
        panel.line.width = Pt(1.5)
        return panel

    def _title(self, slide, text, top=0.55, size=40):
        box = slide.shapes.add_textbox(Inches(1.5), Inches(top), Inches(10.3), Inches(0.9))
        tf = box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        set_para(p, PP_ALIGN.CENTER)
        set_run(p.add_run(), text, size, True, self.t["accent"])

    def _subtitle(self, slide, text, top=1.35, size=24):
        box = slide.shapes.add_textbox(Inches(2), Inches(top), Inches(9.3), Inches(0.6))
        p = box.text_frame.paragraphs[0]
        set_para(p, PP_ALIGN.CENTER)
        set_run(p.add_run(), text, size, False, self.t["text"])

    def _tag(self, slide, text, top=1.15):
        w = max(2.5, len(text) * 0.28)
        left = Inches(CX - w / 2)
        tag = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(top), Inches(w), Inches(0.45))
        tag.fill.solid()
        tag.fill.fore_color.rgb = rgb(self.t["accent2"])
        tag.line.fill.background()
        box = slide.shapes.add_textbox(left, Inches(top + 0.02), Inches(w), Inches(0.42))
        p = box.text_frame.paragraphs[0]
        set_para(p, PP_ALIGN.CENTER)
        set_run(p.add_run(), text, 18, True, self.t["text"])

    def _divider(self, slide, top):
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(3.5), Inches(top), Inches(6.3), Inches(0.03))
        line.fill.solid()
        line.fill.fore_color.rgb = rgb(self.t["accent2"])
        line.line.fill.background()

    def _bullets(self, slide, items, top=2.0, centered=True):
        """items: list of str or (label, str)"""
        n = len(items)
        panel_h = min(4.5, 0.65 * n + 0.5)
        self._panel(slide, top - 0.2, panel_h)
        cx = CX - 4.2 if centered else 1.8
        for i, item in enumerate(items):
            y = top + i * 0.62
            if isinstance(item, tuple):
                label, text = item
                display = f"{label}：{text}"
                badge = str(i + 1)
            else:
                display = item
                badge = str(i + 1)
            circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx), Inches(y), Inches(0.42), Inches(0.42))
            circ.fill.solid()
            circ.fill.fore_color.rgb = rgb(self.t["accent"])
            circ.line.fill.background()
            nb = slide.shapes.add_textbox(Inches(cx), Inches(y + 0.02), Inches(0.42), Inches(0.4))
            np = nb.text_frame.paragraphs[0]
            set_para(np, PP_ALIGN.CENTER)
            set_run(np.add_run(), badge, 16, True, (255, 255, 255))
            box = slide.shapes.add_textbox(Inches(cx + 0.55), Inches(y - 0.02), Inches(8.5), Inches(0.55))
            p = box.text_frame.paragraphs[0]
            set_para(p, PP_ALIGN.LEFT)
            set_run(p.add_run(), display, 24, False, self.t["text"])

    def _question(self, slide, q, hint=None):
        self._tag(slide, "想一想", 1.1)
        self._title(slide, q, 1.75, 32)
        if hint:
            self._divider(slide, 4.8)
            box = slide.shapes.add_textbox(Inches(2), Inches(5.1), Inches(9.3), Inches(1.0))
            p = box.text_frame.paragraphs[0]
            set_para(p, PP_ALIGN.CENTER)
            set_run(p.add_run(), hint, 22, False, self.t["accent"])

    def cover(self, main, sub):
        slide = self._blank()
        self._base(slide)
        self._frame(slide, double=True)
        # center big title
        box = slide.shapes.add_textbox(Inches(1.5), Inches(2.2), Inches(10.3), Inches(1.8))
        p = box.text_frame.paragraphs[0]
        set_para(p, PP_ALIGN.CENTER)
        set_run(p.add_run(), main, 52, True, self.t["accent"])
        p2 = box.text_frame.add_paragraph()
        set_para(p2, PP_ALIGN.CENTER)
        set_run(p2.add_run(), sub, 26, False, self.t["text"])
        # bottom decorative line
        self._divider(slide, 5.5)

    def content(self, title, tag, items):
        slide = self._blank()
        self._base(slide)
        self._tag(slide, tag, 0.55)
        self._title(slide, title, 1.05, 36)
        self._divider(slide, 1.75)
        self._bullets(slide, items, 2.0)

    def think(self, q, hint=None):
        slide = self._blank()
        self._base(slide)
        self._question(slide, q, hint)

    def summary(self, lines):
        self.content("课堂小结", "回顾收获", lines)

    def save(self, name):
        OUT_DIR.mkdir(parents=True, exist_ok=True)
        path = OUT_DIR / name
        self.prs.save(str(path))
        print(f"✓ {path.name} ({len(self.prs.slides)} slides) [{self.t['name']}]")


# ── Lesson content (aligned with 学历案) ───────────────────────────

def build_unit_intro():
    d = ThemedDeck("unit")
    d.cover("想象力博物馆", "七年级上册 · 第六单元 · 第一课时")
    d.content("本课目标", "课时目标", [
        "感知单元学习目标、文体类型与挑战任务",
        "快速默读，用关键词和情节节点概括故事",
        "每分钟阅读不少于400字",
    ])
    d.content("大任务", "单元挑战", [
        "校园文学社举办「想象力博物馆」微展览",
        "小组担任策展人，制作「想象说明卡」",
        "完成原创想象文字或一幕微型课本剧",
    ])
    d.content("母题递进链", "想象与真实", [
        ("自由与规则", "《小圣施威降大圣》"),
        ("真实与虚假", "《皇帝的新装》"),
        ("创造与生命", "《女娲造人》"),
        ("智慧与局限", "《寓言四则》"),
    ])
    d.think("如果伞能说话、星星掉进教室……\n这些离奇念头从哪来？", "提示：从真实经验出发，联想与想象")
    d.content("学习任务一", "明确任务", [
        "阅读学历案，圈画挑战任务与学习目标",
        "同伴交流：说说你的策展想法",
        "填写：本单元有哪些文体？阅读方法是什么？",
    ])
    d.content("学习任务二", "快速阅读", [
        "快速默读（不出声、不指读）",
        "段末停顿，用5—8字记录段意",
        "归纳复述三要素：人物、事件、结果",
        "同桌互评：是否说清人物、事件、结果？",
    ])
    d.summary([
        "本单元学习重点：想象与真实",
        "四条策展主题：自由→真实→创造→智慧",
        "快速阅读法：人物—事件—结果",
    ])
    d.content("课后作业", "课后延伸", [
        ("必做", "预习一篇课文，写100字以内故事梗概"),
        ("选做", "记录生活中一个「如果……会怎样」的想象问题"),
    ])
    d.save("第1课时 单元导学.pptx")


def build_lesson21():
    d = ThemedDeck("wukong")
    d.cover("小圣施威降大圣", "第21课 · 第二、三课时 · 自由与规则")
    d.content("本课目标", "课时目标", [
        "梳理孙悟空与二郎神的变化链，概括故事经过",
        "分析神魔形象及小说想象的趣味",
        "探究「自由与规则」母题，提炼策展主题一",
    ])
    d.think("高手过招，吸引人的不只是「谁赢了」", "把自己当作赛事解说员，抓住每一次变化")
    d.content("学习任务一", "变形追逐图", [
        "默读课文，圈出每次「变作」对象，标序号",
        "填写七十二变对照表（孙悟空先变—二郎神相克）",
        "用「孙悟空先……，二郎神便……」复述最精彩一轮",
    ])
    d.content("学习任务二", "加解说词", [
        "变庙宇：尾巴为何变旗竿？写一句批注",
        "完成人物档案卡（证据—特点—作用）",
        "批注鲁迅、林庚的评论，找「人情」与「世故」",
    ])
    d.think("「使神魔皆有人情，精魅亦通世故」\n课文哪里体现了人情与世故？", "林庚：不宜看得过于认真，应看到儿童心理")
    d.content("学习任务三", "主题探究", [
        "讨论：孙悟空为何逃？天庭为何捉他？",
        "联系生活：追求自由与遵守规则矛盾吗？",
        "策展主题一：自由可贵，但需守规则",
    ])
    d.summary([
        "想象精彩，来自变化中的因果回应",
        "神魔有法力，也有机敏、急切、好胜等人情味",
        "策展主题：自由与规则",
    ])
    d.content("课后作业", "课后延伸", [
        ("必做", "写80字「赛事快讯」（至少三次变化）"),
        ("链接", "作者吴承恩，成书明代，再写一个经典情节"),
    ])
    d.save("第21课 小圣施威降大圣.pptx")


def build_lesson22():
    d = ThemedDeck("emperor")
    d.cover("皇帝的新装", "第22课 · 第四、五课时 · 真实与虚假")
    d.content("本课目标", "课时目标", [
        "梳理童话情节，用简洁语言复述故事",
        "赏析夸张、反讽等手法，理解讽刺艺术",
        "探究「真话为何难以说出口」，树立独立判断",
    ])
    d.think("爱新衣的皇帝、织空布的骗子\n不敢说真话的大人、说真话的孩子", "闹剧背后，藏着怎样的真相？")
    d.content("学习任务一", "画流程图", [
        "快速阅读，每读完一段记下「谁做了什么」",
        "用箭头串联骗局全过程",
        "骗子谎言 → 大臣附和 → 游行 → 孩子说真话",
    ])
    d.content("学习任务二", "拆心理机关", [
        "勾画老大臣、官员、皇帝的心理和语言",
        "品味夸张语句的讽刺效果",
        "完成人物分析卡（证据—特点—作用）",
    ])
    d.think("人们都不敢说自己看不见，这是为什么？", "思考：他们共同害怕失去什么？")
    d.content("学习任务三", "学会判断", [
        "完成课后「思考·探究·积累」三道思考题",
        "联系生活：你有没有「随大流」的经历？",
        "策展主题二：面对虚假，要敢于说真话",
    ])
    d.summary([
        "夸张、重复、反差制造笑声，指向盲从与虚伪",
        "孩子的天真，是不被虚假规则束缚的诚实",
        "策展主题：真实与虚假",
    ])
    d.content("课后作业", "课后延伸", [
        ("必做", "用150字概括《皇帝的新装》的故事"),
        ("情境", "班级群流传「震惊」消息，你会怎么做？（50字）"),
    ])
    d.save("第22课 皇帝的新装.pptx")


def build_lesson23():
    d = ThemedDeck("nuwa")
    d.cover("女娲造人", "第23课 · 第六、七课时 · 创造与生命")
    d.content("本课目标", "课时目标", [
        "概括女娲造人过程，梳理文章写作思路",
        "比较古籍与课文，体会作者的想象与创造",
        "感悟神话中蕴含的对生命与创造的礼赞",
    ])
    d.think("为什么古人要想象一位女神来创造人类？", "这是对「人从哪里来、为何需要同伴」的诗意追问")
    d.content("学习任务一", "追踪创造", [
        "给女娲行动排序：孤独→揉泥→挥洒→繁衍",
        "比较两种造人方法的不同",
        "复述：从孤独到热闹的过程",
    ])
    d.content("学习任务二", "比较阅读", [
        "阅读《风俗通》记载，完成「古籍与课文」比较表",
        "细读池水照影、黄泥捏人、藤条挥洒等画面",
        "说说想象如何具体、生动、温暖",
    ])
    d.content("学习任务三", "主题探究", [
        "讨论：神话用想象解释人的来处",
        "女娲既有神性，也有人的孤独、疲倦和喜悦",
        "策展主题三：生命来之不易，创造值得礼赞",
    ])
    d.summary([
        "神话改写基于古籍又超越古籍",
        "赋予创世故事现代意识与生命温度",
        "策展主题：创造与生命",
    ])
    d.content("课后作业", "课后延伸", [
        ("必做", "以「我看见女娲……」开头，写100字画面描述"),
        ("选做", "查找一个中国创世神话，与本文比较"),
    ])
    d.save("第23课 女娲造人.pptx")


def build_lesson24():
    d = ThemedDeck("fable")
    d.cover("寓言四则", "第24课 · 第八、九课时 · 智慧与局限")
    d.content("本课目标", "课时目标", [
        "归纳寓言文体特点，理解各则寓意",
        "借助注释疏通文言，积累重点词语",
        "分析寓意与情节设计，尝试寓言新编",
    ])
    d.think("寓言是「穿着外套的真理」", "伊索、吕不韦、列子……跨越时空，智慧相通")
    d.content("学习任务一", "初识寓言", [
        "默读四则寓言，完成「寓言档案卡」",
        "讨论：蚊子平静离开，寓意会发生什么变化？",
        "归纳寓言文体特点",
    ])
    d.content("学习任务二", "深读寓意", [
        "解释加点词：闻、传、道、亡、只使、中伤……",
        "翻译：「得一人之使，非得一人于井中也」",
        "讨论「杞人忧天」的两种理解，结合生活举例",
    ])
    d.content("学习任务三", "寓言新编", [
        "任选一则，重新设计情节（100—200字）",
        "策展主题四：认识认知局限，善用智慧",
        "同伴互评：寓意是否由情节自然推出？",
    ])
    d.summary([
        "寓言道理藏在欲望、选择与意外后果中",
        "读懂「为什么这样结尾」，才能读到真正的提醒",
        "策展主题：智慧与局限",
    ])
    d.content("课后作业", "课后延伸", [
        ("必做", "用「这则寓言告诉我们……」写《蚊子和狮子》启示"),
        ("背诵", "得一人之使，非得一人于井中也"),
    ])
    d.save("第24课 寓言四则.pptx")


def build_final():
    d = ThemedDeck("final")
    d.cover("想象力博物馆开幕", "第十—十二课时 · 综合·写作·展示")
    d.content("第十课时", "综合研读", [
        "比较四类文本想象的依据、表现方式和现实意味",
        "完成一张有文本依据的「想象说明卡」",
        "同伴审稿：事实准确？神奇具体？照见了什么？",
    ])
    d.content("想象的四种面貌", "比较阅读", [
        ("神魔小说", "变化斗法，人情世故"),
        ("童话", "夸张反讽，照见人性"),
        ("神话", "创世温情，生命礼赞"),
        ("寓言", "简短警醒，寓意深远"),
    ])
    d.content("第十一课时", "写作训练", [
        "朗读《天上的街市》，辨析联想与想象",
        "三项要求：有依据、合逻辑、有新意",
        "完成创作提纲，独立写作不少于300字",
    ])
    d.think("联想是搭桥，想象是飞翔\n你的作品从哪里起飞？", "续写童话 / 改写寓言 / 十年后的我 / 课本剧")
    d.content("第十二课时", "成果展示", [
        "布展：说明卡 + 创作作品，附一句「请看我如何从文本出发」",
        "讲解：30秒依据 + 60秒亮点 + 30秒思考",
        "参观者投「发现卡」并写理由",
    ])
    d.content("回望母题链", "单元总结", [
        "自由 → 真实 → 创造 → 智慧",
        "成长就是在四者间不断寻找平衡",
        "填写自评表，整理学习档案",
    ])
    d.summary([
        "想象既是阅读能力，也是表达能力",
        "尊重文本和生活的逻辑，看见习以为常之外的世界",
    ])
    d.content("单元作业", "课后延伸", [
        ("阅读", "完善策展卡，600字想象主题作文"),
        ("或", "5分钟以内课本剧脚本"),
        ("要求", "有依据、合逻辑、有新意，体现策展主题"),
    ])
    d.save("第10-12课时 单元成果展示.pptx")


def main():
    for fn in [build_unit_intro, build_lesson21, build_lesson22,
               build_lesson23, build_lesson24, build_final]:
        fn()
    print("\n全部生成完成！")


if __name__ == "__main__":
    main()
