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

    def _bullets(self, slide, items, top=2.0, centered=True, compact=False):
        """items: list of str or (label, str)"""
        n = len(items)
        row_h = 0.52 if compact else 0.62
        fs = 22 if compact else 24
        panel_h = min(5.0, row_h * n + 0.45)
        self._panel(slide, top - 0.2, panel_h)
        cx = CX - 4.2 if centered else 1.8
        for i, item in enumerate(items):
            y = top + i * row_h
            if isinstance(item, tuple):
                label, text = item
                display = f"{label}：{text}"
                badge = str(i + 1)
            else:
                display = item
                badge = str(i + 1)
            circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx), Inches(y), Inches(0.38), Inches(0.38))
            circ.fill.solid()
            circ.fill.fore_color.rgb = rgb(self.t["accent"])
            circ.line.fill.background()
            nb = slide.shapes.add_textbox(Inches(cx), Inches(y + 0.02), Inches(0.38), Inches(0.36))
            np = nb.text_frame.paragraphs[0]
            set_para(np, PP_ALIGN.CENTER)
            set_run(np.add_run(), badge, 14, True, (255, 255, 255))
            box = slide.shapes.add_textbox(Inches(cx + 0.5), Inches(y - 0.02), Inches(8.6), Inches(0.5))
            p = box.text_frame.paragraphs[0]
            set_para(p, PP_ALIGN.LEFT)
            set_run(p.add_run(), display, fs, False, self.t["text"])

    def _kv_notes(self, slide, items, top=1.95):
        """Knowledge notes: (keyword, explanation) pairs for note-taking."""
        n = len(items)
        panel_h = min(5.2, 0.72 * n + 0.35)
        self._panel(slide, top - 0.15, panel_h, width=10.8)
        for i, item in enumerate(items):
            y = top + i * 0.72
            if isinstance(item, tuple):
                key, val = item
            else:
                key, val = "", item
            if key:
                kb = slide.shapes.add_textbox(Inches(1.35), Inches(y), Inches(2.6), Inches(0.55))
                kp = kb.text_frame.paragraphs[0]
                set_para(kp, PP_ALIGN.RIGHT)
                set_run(kp.add_run(), key, 22, True, self.t["accent"])
                xb, xw = 4.1, 8.3
            else:
                xb, xw = 1.35, 11.0
            vb = slide.shapes.add_textbox(Inches(xb), Inches(y), Inches(xw), Inches(0.65))
            vp = vb.text_frame.paragraphs[0]
            vp.word_wrap = True
            set_para(vp, PP_ALIGN.LEFT)
            set_run(vp.add_run(), val, 22, False, self.t["text"])

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

    def content(self, title, tag, items, compact=False):
        slide = self._blank()
        self._base(slide)
        self._tag(slide, tag, 0.55)
        self._title(slide, title, 1.05, 36)
        self._divider(slide, 1.75)
        self._bullets(slide, items, 2.0, compact=compact)

    def knowledge(self, title, items, tag="记笔记"):
        """Substantive knowledge slide for student note-taking."""
        slide = self._blank()
        self._base(slide)
        self._tag(slide, tag, 0.55)
        self._title(slide, title, 1.05, 34)
        self._divider(slide, 1.75)
        self._kv_notes(slide, items, 1.95)

    def think(self, q, hint=None):
        slide = self._blank()
        self._base(slide)
        self._question(slide, q, hint)

    def summary(self, lines):
        self.content("课堂小结", "回顾收获", lines, compact=True)

    def save(self, name):
        OUT_DIR.mkdir(parents=True, exist_ok=True)
        path = OUT_DIR / name
        self.prs.save(str(path))
        print(f"✓ {path.name} ({len(self.prs.slides)} slides) [{self.t['name']}]")


# ── Lesson content (aligned with 学历案) ───────────────────────────

def build_unit_intro():
    d = ThemedDeck("unit")
    d.cover("想象力博物馆", "七年级上册 · 第六单元 · 第一课时")
    d.knowledge("单元知识框架", [
        ("单元母题", "想象与真实——想象源于生活经验，又超越现实"),
        ("文体类型", "神魔小说、童话、神话、寓言（四种富于想象力的文体）"),
        ("阅读方法", "快速默读：把握思路，发挥联想和想象"),
        ("阅读速度", "每分钟不少于400字，先抓「谁—做什么—结果」"),
    ])
    d.knowledge("母题递进链", [
        ("自由与规则", "《小圣施威降大圣》——神魔斗法中见人情"),
        ("真实与虚假", "《皇帝的新装》——荒诞童话照见人性"),
        ("创造与生命", "《女娲造人》——创世神话礼赞生命"),
        ("智慧与局限", "《寓言四则》——短小故事寄寓道理"),
    ], tag="核心知识")
    d.knowledge("阅读方法笔记", [
        ("情节概括", "用「节点词+简洁句」概括，避免逐字回读"),
        ("人物分析", "证据—特点—作用（先找细节，再概括，后说作用）"),
        ("想象追问", "它像现实中的什么？人物为什么这样做？"),
        ("创作方法", "触发点→联想链→情节转折→表达主题"),
    ])
    d.content("单元大任务", "挑战任务", [
        "举办「想象力博物馆」微展览，小组担任策展人",
        "制作「想象说明卡」：说清想象的依据与表达效果",
        "完成原创想象文字或一幕微型课本剧",
    ])
    d.content("快速默读训练", "学习方法", [
        "不出声、不指读，段末停顿记录段意",
        "用5—8个字概括每个情节段落",
        "复述三要素：人物、事件、结果",
    ])
    d.think("奇异情节从哪来？", "从真实经验出发——伞、星星、孙悟空都连接着我们的生活")
    d.summary([
        "四文体：神魔小说、童话、神话、寓言",
        "母题链：自由→真实→创造→智慧",
        "概括法：人物—事件—结果；分析法：证据—特点—作用",
    ])
    d.content("课后作业", "课后延伸", [
        ("必做", "预习一篇课文，写100字以内故事梗概"),
        ("选做", "记录一个「如果……会怎样」的想象问题"),
    ])
    d.save("第1课时 单元导学.pptx")


def build_lesson21():
    d = ThemedDeck("wukong")
    d.cover("小圣施威降大圣", "第21课 · 第二、三课时 · 自由与规则")
    d.knowledge("文学常识", [
        ("作者", "吴承恩，明代小说家"),
        ("出处", "章回体长篇小说《西游记》节选"),
        ("文体", "神魔小说——以神魔斗法写人情世故"),
        ("本文情节", "孙悟空大闹天宫后，二郎神奉旨捉拿，二人变化斗法"),
    ])
    d.knowledge("变化链梳理", [
        ("第1轮", "悟空变鸟→二郎变鹰；悟空变鱼→二郎变鱼鹰"),
        ("第2轮", "悟空变水蛇→二郎变灰鹤；悟空变花鸨→二郎现原身"),
        ("第3轮", "悟空变庙宇，尾巴变旗竿→二郎识破"),
        ("规律", "孙悟空先变，二郎神随即以相克之物相对"),
    ], tag="情节笔记")
    d.knowledge("人物赏析", [
        ("孙悟空", "机敏善变、好胜不屈，追求自由（证据：连续变化、不肯认输）"),
        ("二郎神", "沉着冷静、法力高强，维护天庭秩序（证据：紧追不舍、识破破绽）"),
        ("分析方法", "证据—特点—作用：找细节→概括性格→联系主题"),
        ("变旗竿", "紧张又好笑——神通广大也有露馅之时"),
    ], tag="人物笔记")
    d.knowledge("名家点评", [
        ("鲁迅", "「使神魔皆有人情，精魅亦通世故」——神魔也有人的情感"),
        ("林庚", "不宜看得过于认真，应看到儿童的心理与行为"),
        ("艺术特色", "想象基于「相克」逻辑，变化链环环相扣"),
        ("策展主题", "自由可贵，但不可无视规则（自由与规则）"),
    ], tag="主题笔记")
    d.content("学习任务", "课堂活动", [
        "填写七十二变对照表，标出相克关系",
        "完成人物档案卡（证据—特点—作用）",
        "用「孙悟空先……，二郎神便……」复述最精彩一轮",
    ])
    d.think("追求自由与遵守规则矛盾吗？", "结合课文和生活各举一例说明")
    d.summary([
        "神魔小说：想象在「相克」中见趣味，在斗法中见人情",
        "变化链是因果回应，不是随意编造",
        "策展主题一：自由与规则",
    ])
    d.content("课后作业", "课后延伸", [
        ("必做", "写80字「赛事快讯」（至少三次变化）"),
        ("积累", "《西游记》作者吴承恩，成书明代；再写一个孙悟空经典情节"),
    ])
    d.save("第21课 小圣施威降大圣.pptx")


def build_lesson22():
    d = ThemedDeck("emperor")
    d.cover("皇帝的新装", "第22课 · 第四、五课时 · 真实与虚假")
    d.knowledge("文学常识", [
        ("作者", "安徒生，丹麦童话作家"),
        ("文体", "童话——用虚构故事反映现实生活"),
        ("本文特点", "以「新装」为线索，写一个荒唐的骗局"),
        ("核心问题", "为什么人们都不敢说自己看不见？"),
    ])
    d.knowledge("情节结构", [
        ("开端", "两个骗子说能织出只有聪明人才能看见的衣服"),
        ("发展", "皇帝、老大臣、官员先后「看布」，都不敢说真话"),
        ("高潮", "皇帝穿上「新装」游行，百姓假装看见"),
        ("结局", "小孩子说「他没穿衣服」，百姓跟着说，皇帝仍装模作样"),
    ], tag="情节笔记")
    d.knowledge("艺术手法", [
        ("夸张", "把皇帝爱新衣写到极致：「除非为了炫耀新衣服」"),
        ("反讽", "说「看见了」其实什么也没看见，讽刺虚伪"),
        ("对比", "大人的怯懦 vs 孩子的天真诚实"),
        ("反复", "「我什么也没有看见」反复出现，强化讽刺效果"),
    ], tag="手法笔记")
    d.knowledge("主题理解", [
        ("心理机制", "害怕被认为「不聪明」或「不称职」而丢失地位"),
        ("社会批判", "讽刺盲从权威、虚伪逢迎的官场风气"),
        ("孩子作用", "天真不受虚假规则束缚，一语道破真相"),
        ("策展主题", "面对虚假，要敢于说真话，不做沉默的大多数"),
    ], tag="主题笔记")
    d.content("学习任务", "课堂活动", [
        "画骗局流程图，用箭头串联各环节",
        "勾画老大臣、官员、皇帝的心理和语言",
        "完成人物分析卡（证据—特点—作用）",
    ])
    d.summary([
        "夸张+反讽+对比，让笑声指向盲从与虚伪",
        "「新装」是谎言，更是照见人性的镜子",
        "策展主题二：真实与虚假",
    ])
    d.content("课后作业", "课后延伸", [
        ("必做", "用150字概括故事"),
        ("情境", "班级群流传「震惊」消息，你会怎么做？（50字）"),
    ])
    d.save("第22课 皇帝的新装.pptx")


def build_lesson23():
    d = ThemedDeck("nuwa")
    d.cover("女娲造人", "第23课 · 第六、七课时 · 创造与生命")
    d.knowledge("文学常识", [
        ("作者", "袁珂（现代作家，神话研究专家）"),
        ("文体", "神话——用想象解释自然与人类起源"),
        ("改写特点", "在古籍记载基础上增删改写，赋予现代意识"),
        ("对比材料", "教材「阅读提示」中的《风俗通》记载"),
    ])
    d.knowledge("造人过程", [
        ("起因", "女娲行走世间，感到孤独寂寞"),
        ("方法一", "黄泥和水，揉成小泥人，泥人落地即活"),
        ("方法二", "藤条蘸泥浆挥洒，溅落的泥点也变成人"),
        ("结果", "建立婚姻制度，让人类繁衍生息"),
    ], tag="情节笔记")
    d.knowledge("想象特点", [
        ("具体", "池水照影见自己面容→想到造同类"),
        ("生动", "「藤条一挥，满天泥浆洒落」画面感强"),
        ("温暖", "造人后的疲倦、喜悦，赋予神以人的情感"),
        ("规律", "想象基于生活经验（和泥、洒水），又超越现实"),
    ], tag="写法笔记")
    d.knowledge("比较阅读", [
        ("《风俗通》", "记载简洁，重在说明造人方法"),
        ("课文", "增加孤独心理、造人细节、情感描写"),
        ("改写意图", "让神话更生动，表达对生命与创造的礼赞"),
        ("策展主题", "生命来之不易，创造值得珍视（创造与生命）"),
    ], tag="主题笔记")
    d.content("学习任务", "课堂活动", [
        "给女娲行动排序，比较两种造人方法",
        "完成「古籍与课文」比较表",
        "以「我看见女娲……」写100字画面描述",
    ])
    d.summary([
        "神话想象：基于古籍，超越古籍，有温度",
        "女娲：神性与人性并存——孤独、疲倦、喜悦",
        "策展主题三：创造与生命",
    ])
    d.content("课后作业", "课后延伸", [
        ("必做", "以「我看见女娲……」写100字画面"),
        ("选做", "查找一个中国创世神话，与本文比较"),
    ])
    d.save("第23课 女娲造人.pptx")


def build_lesson24():
    d = ThemedDeck("fable")
    d.cover("寓言四则", "第24课 · 第八、九课时 · 智慧与局限")
    d.knowledge("文体知识", [
        ("寓言", "短小故事，寄寓深刻道理，多运用拟人手法"),
        ("特点", "情节简短、人物典型、寓意明确、多讽刺或警醒"),
        ("四则出处", "《赫耳墨斯和雕像者》《蚊子和狮子》出自《伊索寓言》"),
        ("", "《穿井得一人》《杞人忧天》出自《吕氏春秋》"),
    ])
    d.knowledge("四则寓意", [
        ("赫耳墨斯", "讽刺自高自大、妄自尊重的人"),
        ("蚊子和狮子", "再小的个体也有长处，骄兵必败"),
        ("穿井得一人", "以讹传讹，调查求证才能辨明真相"),
        ("杞人忧天", "讽刺不必要的担忧（也可读出忧患意识）"),
    ], tag="寓意笔记")
    d.knowledge("文言积累", [
        ("闻", "听说；传：传播、流传"),
        ("道", "讲述、说（国人道之）"),
        ("亡", "同「无」，没有（亡处亡气）"),
        ("只使", "纵使、即使；中伤：伤害"),
        ("重点句", "得一人之使，非得一人于井中也。（节省一人劳力，不是井里挖出一人）"),
    ], tag="文言笔记")
    d.knowledge("写法探究", [
        ("情节设计", "寓意藏在人物的欲望、选择与意外后果中"),
        ("改情节改寓意", "蚊子战胜后平静离开→寓意从「骄兵必败」变为「懂得适可而止」"),
        ("阅读方法", "读懂「为什么这样结尾」，才能读到真正的提醒"),
        ("策展主题", "认识自己的认知局限，善用智慧（智慧与局限）"),
    ], tag="主题笔记")
    d.content("学习任务", "课堂活动", [
        "完成「寓言档案卡」，概括四则寓意",
        "解释加点词，翻译重点句",
        "任选一则寓言新编（100—200字）",
    ])
    d.summary([
        "寓言=故事+道理，道理由情节自然推出",
        "文言：闻、传、道、亡；重点句要准确翻译",
        "策展主题四：智慧与局限",
    ])
    d.content("课后作业", "课后延伸", [
        ("必做", "用「这则寓言告诉我们……」写《蚊子和狮子》启示"),
        ("背诵", "得一人之使，非得一人于井中也"),
    ])
    d.save("第24课 寓言四则.pptx")


def build_final():
    d = ThemedDeck("final")
    d.cover("想象力博物馆开幕", "第十—十二课时 · 综合·写作·展示")
    d.knowledge("四类文本比较", [
        ("神魔小说", "想象依据：相克逻辑；表现：连续变化；意味：人情世故"),
        ("童话", "想象依据：生活经验；表现：夸张反讽；意味：照见人性"),
        ("神话", "想象依据：古籍传说；表现：创世画面；意味：礼赞生命"),
        ("寓言", "想象依据：生活现象；表现：拟人故事；意味：警醒智慧"),
    ], tag="比较笔记")
    d.knowledge("联想与想象", [
        ("联想", "由一事物想到另一事物（街灯→明星，因形状相似）"),
        ("想象", "在已有材料上创造新形象（天上的街市、牛郎织女骑牛）"),
        ("写作要求", "有依据、合逻辑、有新意"),
        ("创作路径", "触发点→联想链→情节转折→表达主题"),
    ], tag="写作笔记")
    d.knowledge("想象说明卡", [
        ("文本片段", "从本单元选一个精彩片段"),
        ("想象依据", "它从什么现实经验或材料出发？"),
        ("表现方式", "用了什么手法？（变化/夸张/改写/拟人）"),
        ("现实意味", "照见了怎样的生活道理？（联系策展主题）"),
    ], tag="策展笔记")
    d.content("成果展示要求", "展示活动", [
        "布展：说明卡+创作作品，附「请看我如何从文本出发」",
        "讲解：30秒依据+60秒亮点+30秒思考",
        "评价：投「发现卡」——最有依据 / 最有新意",
    ])
    d.knowledge("单元回顾", [
        ("母题链", "自由→真实→创造→智慧"),
        ("核心领悟", "想象必须扎根于真实，飞得再高也要落回地面"),
        ("单元作业", "600字想象作文，或5分钟课本剧脚本"),
        ("要求", "有依据、合逻辑、有新意，体现策展主题"),
    ], tag="总结笔记")
    d.save("第10-12课时 单元成果展示.pptx")


def main():
    for fn in [build_unit_intro, build_lesson21, build_lesson22,
               build_lesson23, build_lesson24, build_final]:
        fn()
    print("\n全部生成完成！")


if __name__ == "__main__":
    main()
