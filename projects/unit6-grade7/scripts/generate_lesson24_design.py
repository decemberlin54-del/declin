#!/usr/bin/env python3
"""Generate lesson 24 design docx in 学历案 format."""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

OUT = Path("/workspace/《寓言四则》短故事，大镜子课时设计.docx")
BLANK = " " * 60


def set_run_font(run, name="宋体", size=12, bold=False):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)


def add_para(doc, text="", bold=False, indent=0):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    p.paragraph_format.first_line_indent = Cm(indent)
    r = p.add_run(text)
    set_run_font(r, bold=bold)
    return p


def add_blank(doc, n=1):
    for _ in range(n):
        add_para(doc, BLANK)


def add_table(doc, headers, rows, col_widths=None):
    tbl = doc.add_table(rows=1 + len(rows), cols=len(headers))
    tbl.style = "Table Grid"
    for ci, h in enumerate(headers):
        tbl.rows[0].cells[ci].text = h
    for ri, row in enumerate(rows, 1):
        for ci, val in enumerate(row):
            tbl.rows[ri].cells[ci].text = val
    return tbl


def build():
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21)
    sec.page_height = Cm(29.7)
    sec.left_margin = Cm(2.5)
    sec.right_margin = Cm(2.5)

    add_para(doc, "第八、九课时   寓言四则   短故事，大镜子", bold=True)
    add_para(doc, "【课时目标】")
    add_para(doc, "1.通过阅读四则寓言，归纳寓言的文体特点，理解各则寓意。")
    add_para(doc, "2.借助注释疏通《穿井得一人》《杞人忧天》，积累重点文言词语。")
    add_para(
        doc,
        "3.分析寓意与情节设计的关系，认识寓言情节与寓意之间的关系，探究“杞人忧天”的多种理解。",
    )
    add_para(doc, "评价任务：")
    add_para(
        doc,
        "1.完成“寓言档案卡”，准确概括四则寓言的情节与寓意，归纳寓言文体特点。（检测目标1）",
    )
    add_para(
        doc,
        "2.借助注释准确翻译《穿井得一人》《杞人忧天》重点语句，解释“闻、传、道、亡”等文言词语。（检测目标2）",
    )
    add_para(
        doc,
        "3.结合具体情节分析寓意如何产生，说出“杞人忧天”的至少两种理解，并能联系生活谈启示。（检测目标3）",
    )
    add_para(doc, "学习过程：")
    add_para(doc, "课前热身：")
    add_para(
        doc,
        "在前面的学习中，我们读过神魔斗法的奇幻、童话的夸张反讽、神话的创世想象。"
        "今天，我们要读四则“短故事”——它们篇幅极短，却像一面面“大镜子”，照见人心与社会。"
        "请你想一想：为什么《赫耳墨斯和雕像者》里，雕像者一句“假如你买了那两个，这个算饶头，白送”，"
        "会让赫耳墨斯难堪？为什么《穿井得一人》里，丁氏一句“吾穿井得一人”会在国都传开？"
        "带着这些问题，走进寓言的世界。（指向目标1）",
    )

    add_para(doc, "学习任务一：初读四则寓言，概括情节，归纳文体特点")
    add_para(doc, "1.快速默读四则寓言，用“谁—做什么—结果怎样”概括每则故事。（指向目标1）")
    add_para(
        doc,
        "2.完成学习任务单一，然后以小组为单位讨论并完善，最后全班交流、展示。（检测目标1）",
    )
    add_para(doc, "学习任务单一：寓言档案卡")
    add_table(
        doc,
        ["篇目", "主要人物", "核心情节（一句话）", "寓意"],
        [
            ["《赫耳墨斯和雕像者》", "赫耳墨斯、雕像者", "", ""],
            ["《蚊子和狮子》", "蚊子、狮子", "", ""],
            ["《穿井得一人》", "丁氏、宋君", "", ""],
            ["《杞人忧天》", "杞人、晓之者", "", ""],
        ],
    )
    add_para(doc, "")
    add_para(
        doc,
        "3.链接文本：读《赫耳墨斯和雕像者》，赫耳墨斯先问宙斯“值多少钱”，再问赫拉“值多少”，"
        "最后看见自己的雕像，心想“他身为神使，又是商人的庇护神，人们会对他更尊重些”。"
        "请你圈画体现他心理变化的词句，想一想：如果没有前面两次“试探”，"
        "结尾雕像者说“这个算饶头，白送”还会有那么强的讽刺效果吗？（检测目标1）",
    )
    add_blank(doc, 2)
    add_para(
        doc,
        "4.链接文本：读《蚊子和狮子》，蚊子专挑“没有名气的狮子”挑战，战胜后“吹着喇叭，唱着凯歌”，"
        "却被蜘蛛网粘住。请你用两个词概括蚊子前后状态的变化，说说这一变化与寓意有什么关系。（检测目标1）",
    )
    add_blank(doc, 2)
    add_para(
        doc,
        "5.小组讨论：四则寓言在篇幅、人物、写法上有哪些共同点？尝试用一句话归纳“寓言”的文体特点。（指向目标1、检测目标1）",
    )
    add_para(doc, "我的发现：寓言的文体特点是——（提示：篇幅上/人物上/写法上/作用上）")
    add_blank(doc, 3)

    add_para(doc, "学习任务二：研读寓言，品味情节，读懂寓意")
    add_para(
        doc,
        "1.寓言的寓意往往藏在情节的“关键一步”里。请再读四则寓言，完成学习任务单二，"
        "分析“情节设计”与“寓意”的关系。（指向目标3）",
    )
    add_para(doc, "学习任务单二：情节与寓意")
    add_table(
        doc,
        ["篇目", "关键情节（请引用原文或概括）", "这一情节如何推出寓意"],
        [
            [
                "《赫耳墨斯和雕像者》",
                "雕像者答：“假如你买了那两个，这个算饶头，白送。”",
                "",
            ],
            [
                "《蚊子和狮子》",
                "蚊子“吹着喇叭，唱着凯歌”，却被蜘蛛网粘住。",
                "",
            ],
            [
                "《穿井得一人》",
                "丁氏对曰：“得一人之使，非得一人于井中也。”",
                "",
            ],
            [
                "《杞人忧天》",
                "晓之者曰：“天，积气耳，亡处亡气……”",
                "",
            ],
        ],
    )
    add_para(doc, "")
    add_para(
        doc,
        "2.深度探究：如果把《蚊子和狮子》的结尾改为“蚊子战胜狮子后悄悄离开”，"
        "寓意会从“骄兵必败”变成什么？这说明了什么？（检测目标3）",
    )
    add_para(doc, "我认为寓意会变成")
    add_blank(doc, 1)
    add_para(doc, "因为")
    add_blank(doc, 1)
    add_para(doc, "这说明：寓言的寓意与")
    add_blank(doc, 1)
    add_para(doc, "密切相关。")
    add_para(
        doc,
        "3.联系生活：下面哪种现象更像《穿井得一人》？哪种更像《赫耳墨斯和雕像者》？请各举一例并说明。（检测目标3）",
    )
    add_para(doc, "①未经核实就转发“某地挖出古墓”的消息  ②因为自己成绩好就看不起同学  ③听到一点风声就整日焦虑不安")
    add_para(doc, "《穿井得一人》像生活中的")
    add_blank(doc, 1)
    add_para(doc, "因为")
    add_blank(doc, 1)
    add_para(doc, "《赫耳墨斯和雕像者》像生活中的")
    add_blank(doc, 1)
    add_para(doc, "因为")
    add_blank(doc, 2)

    add_para(doc, "学习任务三：疏通文言，积累词语，探究“杞人忧天”的多种理解")
    add_para(doc, "1.借助注释，朗读《穿井得一人》《杞人忧天》，读准字音、停顿。（指向目标2）")
    add_para(doc, "2.完成学习任务单三，解释加点词，翻译重点句。（检测目标2）")
    add_para(doc, "学习任务单三：文言积累")
    add_table(
        doc,
        ["语句", "加点词", "释义"],
        [
            ["有闻而传之者", "闻", ""],
            ["国人道之", "道", ""],
            ["闻之于宋君", "闻", ""],
            ["身亡所寄", "亡", ""],
            ["因往晓之", "晓", ""],
            ["其人舍然大喜", "舍然", ""],
        ],
    )
    add_para(doc, "")
    add_para(doc, "重点句翻译：")
    add_para(doc, "（1）得一人之使，非得一人于井中也。")
    add_blank(doc, 2)
    add_para(doc, "（2）求闻之若此，不若无闻也。")
    add_blank(doc, 2)
    add_para(doc, "（3）若屈伸呼吸，终日在天中行止，奈何忧崩坠乎？")
    add_blank(doc, 2)
    add_para(
        doc,
        "3.多种理解探究：《杞人忧天》通常讽刺“不必要的担忧”，"
        "但也有人读出“对未知自然的忧患意识”。请结合原文完成学习任务单四。（指向目标3、检测目标3）",
    )
    add_para(doc, "学习任务单四：")
    add_table(
        doc,
        ["理解角度", "依据文本（请引用原文）", "启示"],
        [
            ["讽刺不必要的杞忧", "如“废寝食者”“奈何忧崩坠乎”", ""],
            ["体现忧患意识（可选）", "如“忧天地崩坠，身亡所寄”", ""],
        ],
    )
    add_para(doc, "")
    add_para(
        doc,
        "4.为“想象力博物馆”提炼第四条策展主题：智慧与局限。"
        "请结合四则寓言中“人物的认知局限”，写一句主题语。（检测目标3）",
    )
    add_para(doc, "主题四（智慧与局限）：")
    add_blank(doc, 2)

    add_para(doc, "四、课后小结")
    add_para(doc, "1.四则寓言的寓意分别是什么？你获得了怎样的人生启示？")
    add_para(doc, "《赫耳墨斯和雕像者》：")
    add_blank(doc, 1)
    add_para(doc, "《蚊子和狮子》：")
    add_blank(doc, 1)
    add_para(doc, "《穿井得一人》：")
    add_blank(doc, 1)
    add_para(doc, "《杞人忧天》：")
    add_blank(doc, 1)
    add_para(doc, "2.寓言的文体特点是什么？情节与寓意有什么关系？对你阅读其他寓言有什么帮助？")
    add_blank(doc, 2)

    add_para(doc, "五、作业与检测：")
    add_para(
        doc,
        "1.必做：用“这则寓言告诉我们……”的句式，写《蚊子和狮子》给你的启示（80字左右）。",
    )
    add_blank(doc, 2)
    add_para(doc, "2.必做：背诵并默写《穿井得一人》中的名句：")
    add_blank(doc, 1)
    add_para(
        doc,
        "3.选做：任选一则寓言进行新编（100—200字）。"
        "要求：保留原寓意，或改变一个关键情节使寓意发生变化，并写出你修改后的寓意。",
    )
    add_blank(doc, 3)
    add_para(
        doc,
        "4.选做：完善“寓言档案卡”，为“想象力博物馆”提交第四条展品说明——《寓言四则》。",
    )
    add_blank(doc, 2)

    doc.save(str(OUT))
    print(f"✓ {OUT}")


if __name__ == "__main__":
    build()
