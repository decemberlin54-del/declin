#!/usr/bin/env python3
"""Generate 七年级上册期末考试语文试卷多维细目表 (landscape Word)."""

from pathlib import Path

import docx
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor

OUT = Path("/workspace/七年级语文期末考试双向细目表.docx")
OUT_EN = Path("/workspace/grade7-final-exam-spec-table.docx")


def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)


def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def style_run(run, size=9, bold=False, color=None):
    run.font.name = "宋体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color


def create_document():
    doc = Document()

    section = doc.sections[0]
    section.orientation = docx.enum.section.WD_ORIENT.LANDSCAPE
    new_width, new_height = section.page_height, section.page_width
    section.page_width = new_width
    section.page_height = new_height
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("七年级上册期末考试语文试卷多维细目表")
    style_run(run, 16, True, RGBColor(0x11, 0x3F, 0x67))

    headers = [
        "板块",
        "考查内容",
        "题序",
        "考点",
        "分值",
        "题型",
        "设题特点或材料",
        "阅读量",
        "难度",
        "试题特点",
        "核心素养内涵",
    ]
    data = [
        [
            "基础积累与运用\n(16分)",
            "字词综合",
            "1",
            "汉字字音辨析",
            "2",
            "选择",
            '语段取材于阅读与成长历程，考查易错字读音（"涵"读音）',
            "约500字\n(1-5题共用)",
            "容易",
            "—",
            "语言运用",
        ],
        [
            "基础积累与运用",
            "字词综合",
            "2",
            "汉字字形辨析",
            "2",
            "选择",
            '考查常见成语易错字辨析（"浅尝辄止"）',
            "",
            "容易",
            "—",
            "语言运用",
        ],
        [
            "基础积累与运用",
            "字词综合",
            "3",
            "词语语境搭配",
            "2",
            "选择",
            "结合语段考查实词与成语在特定语境下的准确运用",
            "",
            "中档",
            "—",
            "语言运用",
        ],
        [
            "基础积累与运用",
            "语言运用",
            "4",
            "病句辨析与修改",
            "2",
            "选择",
            "考查滥用介词导致缺失主语及动宾搭配不当的修改",
            "",
            "中档",
            "—",
            "语言运用",
        ],
        [
            "基础积累与运用",
            "语言运用",
            "5",
            "句意衔接与逻辑排序",
            "2",
            "选择",
            "考查句间逻辑关联及与上下文的紧密衔接",
            "",
            "中档",
            "—",
            "思维能力",
        ],
        [
            "基础积累与运用",
            "诗文默写",
            "6",
            "名篇名句默写",
            "6",
            "填空",
            "考查《观沧海》《次北固山下》《论语·子罕》默写",
            "—",
            "容易",
            "—",
            "文化自信",
        ],
        [
            "综合性学习\n(10分)",
            "综合运用",
            "7",
            "思辨表达与口语交际",
            "2",
            "简答",
            '聚焦"AI摘要速读与传统阅读"，考查辩证思维与口语表达',
            "—",
            "中档",
            "开放性√\n情境性√\n综合性√",
            "思维能力",
        ],
        [
            "综合性学习",
            "名著阅读",
            "8",
            "名著推荐与情境迁移",
            "3",
            "简答",
            "结合《昆虫记》《朝花夕拾》《西游记》特质进行小分队选择",
            "—",
            "中档",
            "开放性√\n情境性√",
            "审美创造",
        ],
        [
            "综合性学习",
            "名著阅读",
            "9",
            "词典释义与书名内涵",
            "2",
            "简答",
            '借助字典"拾"释义，解析《朝花夕拾》书名内涵与情感',
            "—",
            "中档",
            "综合性√",
            "文化自信",
        ],
        [
            "综合性学习",
            "名著阅读",
            "10",
            "名著人物与演讲主题",
            "3",
            "简答",
            '结合孙悟空/法布尔遭遇，阐述"面对困境的姿态"演讲理由',
            "—",
            "较难",
            "开放性√\n情境性√\n综合性√",
            "思维能力",
        ],
        [
            "阅读能力\n(34分)",
            "古代诗歌阅读",
            "11",
            "诗歌内容与赏析",
            "2",
            "选择",
            "【甲】李白《闻王昌龄左迁龙标...》景物、意象与基调",
            "约70字",
            "容易",
            "—",
            "审美创造",
        ],
        [
            "阅读能力",
            "古代诗歌阅读",
            "12",
            "诗歌对比与表现手法",
            "2",
            "简答",
            '对比【甲】诗与【乙】《夜雨寄北》"超越时空想象"机制',
            "",
            "中档",
            "综合性√",
            "审美创造",
        ],
        [
            "阅读能力",
            "文言文阅读",
            "13",
            "文言实词解释",
            "2",
            "填空",
            "考查通假字、古今异义及常见实词（亡、晓、恶、愈）",
            "约300字\n(甲+乙)",
            "容易",
            "—",
            "语言运用",
        ],
        [
            "阅读能力",
            "文言文阅读",
            "14",
            "文言句子翻译",
            "3",
            "翻译",
            "考查重点文言句式翻译（舍然、晓、所以、沈疴）",
            "",
            "中档",
            "—",
            "语言运用",
        ],
        [
            "阅读能力",
            "文言文阅读",
            "15",
            "内容概括与对比",
            "2",
            "简答",
            '概括甲《杞人忧天》乙《乐广传》"不必要的忧虑"成因',
            "",
            "中档",
            "综合性√",
            "思维能力",
        ],
        [
            "阅读能力",
            "文言文阅读",
            "16",
            "情境劝慰与主旨迁移",
            "3",
            "简答",
            "跨文本角色扮演，结合两文主旨撰写劝慰开导语（60字）",
            "",
            "较难",
            "开放性√\n情境性√\n综合性√",
            "思维能力",
        ],
        [
            "阅读能力",
            "说明性文本阅读",
            "17",
            "信息提取与作用分析",
            "5",
            "填空/简答",
            "提取词汇/名称；概括保护方面；分析引用民歌作用",
            "约1200字",
            "中档",
            "—",
            "思维能力",
        ],
        [
            "阅读能力",
            "说明性文本阅读",
            "18",
            "补充段落作用分析",
            "2.5",
            "简答",
            "从内容、结构、主旨三维度分析第⑦段拓展作用",
            "",
            "较难",
            "综合性√",
            "思维能力",
        ],
        [
            "阅读能力",
            "说明性文本阅读",
            "19",
            "跨文本/文体写作目的",
            "2.5",
            "简答",
            "对比科普说明文与生态随笔《大雁归来》写作目的异同",
            "",
            "较难",
            "综合性√",
            "思维能力",
        ],
        [
            "阅读能力",
            "文学类文本阅读",
            "20",
            "人物心理状态概括",
            "2",
            "简答",
            '概括"我"面对母亲送伞时暴躁与后悔的心理变迁',
            "约1100字",
            "中档",
            "—",
            "思维能力",
        ],
        [
            "阅读能力",
            "文学类文本阅读",
            "21",
            "关键语句深层含义理解",
            "2",
            "简答",
            '剖析"母爱就像伞"的表层含义与深层哲理',
            "",
            "中档",
            "—",
            "审美创造",
        ],
        [
            "阅读能力",
            "文学类文本阅读",
            "22",
            "跨文本修辞效果对比",
            "3",
            "简答",
            "对比本文比喻（伞与龙眼）与《荷叶·母亲》效果异同",
            "",
            "较难",
            "综合性√",
            "审美创造",
        ],
        [
            "阅读能力",
            "文学类文本阅读",
            "23",
            "多文本联动与主题探究",
            "3",
            "简答",
            "结合《秋天的怀念》《寒风吹彻》系统阐述对母爱的理解",
            "",
            "较难",
            "情境性√\n综合性√",
            "思维能力",
        ],
        [
            "写作能力\n(40分)",
            "写作",
            "24/25",
            "命题/半命题作文",
            "40",
            "写作",
            "题目一《在那前，我停了下来》；题目二《唤醒心中的____》",
            "约90字",
            "较难",
            "开放性√\n情境性√\n综合性√",
            "语言运用\n思维能力\n审美创造\n文化自信",
        ],
    ]

    table = doc.add_table(rows=len(data) + 1, cols=11)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    hdr_cells = table.rows[0].cells
    for i, title_text in enumerate(headers):
        hdr_cells[i].text = title_text
        set_cell_background(hdr_cells[i], "113F67")
        hdr_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            style_run(r, 9, True, RGBColor(0xFF, 0xFF, 0xFF))

    center_cols = {2, 4, 5, 7, 8, 9}
    for row_idx, row_data in enumerate(data):
        row_cells = table.rows[row_idx + 1].cells
        bg_color = "F4F6F9" if row_idx % 2 == 1 else "FFFFFF"
        for col_idx, cell_value in enumerate(row_data):
            row_cells[col_idx].text = cell_value
            set_cell_background(row_cells[col_idx], bg_color)
            set_cell_margins(row_cells[col_idx], top=80, bottom=80, left=100, right=100)
            row_cells[col_idx].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = row_cells[col_idx].paragraphs[0]
            p.alignment = (
                WD_ALIGN_PARAGRAPH.CENTER if col_idx in center_cols else WD_ALIGN_PARAGRAPH.LEFT
            )
            for r in p.runs:
                style_run(r, 8.5)

    doc.save(str(OUT))
    import shutil

    shutil.copy2(OUT, OUT_EN)
    print(f"✓ {OUT}")
    print(f"✓ {OUT_EN}")


if __name__ == "__main__":
    create_document()
