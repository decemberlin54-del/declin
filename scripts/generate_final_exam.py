#!/usr/bin/env python3
"""Generate 7th grade Semester 1 final exam matching 2025 Yunnan format."""

from pathlib import Path
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, Cm

OUT = Path("/workspace")


def set_font(run, name="宋体", size=12, bold=False):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)


def add_p(doc, text, align=WD_ALIGN_PARAGRAPH.LEFT, font="宋体", size=12, bold=False,
          indent=None, space_after=0):
    p = doc.add_paragraph()
    p.alignment = align
    if indent:
        p.paragraph_format.first_line_indent = Cm(indent)
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text)
    set_font(r, font, size, bold)
    return p


def add_mixed(doc, parts, align=WD_ALIGN_PARAGRAPH.LEFT, indent=None):
    p = doc.add_paragraph()
    p.alignment = align
    if indent:
        p.paragraph_format.first_line_indent = Cm(indent)
    for text, font, size, bold in parts:
        r = p.add_run(text)
        set_font(r, font, size, bold)
    return p


def build_exam(with_answers=False):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21)
    sec.page_height = Cm(29.7)
    sec.left_margin = Cm(2.5)
    sec.right_margin = Cm(2.5)

    add_p(doc, "昆明市××学校2025—2026学年度第一学期期末考试", WD_ALIGN_PARAGRAPH.CENTER, size=16, bold=True)
    add_p(doc, "语  文", WD_ALIGN_PARAGRAPH.CENTER, size=16, bold=True)
    add_p(doc, "（全卷四个大题，共24个小题，共8页；满分100分，考试用时150分钟）", WD_ALIGN_PARAGRAPH.CENTER, size=12)
    add_p(doc, "注意事项：", bold=True)
    add_p(doc, "1．考生必须在答题卡上解题作答。答案应书写在答题卡的相应位置上，在试卷、草稿纸上作答无效。")
    add_p(doc, "2．考试结束后，请将试卷和答题卡一并交回。")

    # ── 一、积累与运用 ──
    add_p(doc, "一、积累与运用（1~5题，每题2分，第6题6分，共16分）", bold=True)
    add_p(doc, "阅读下面文字，按要求完成下面小题。")

    passage1 = (
        "阅读，是心灵的旅行。明代学者董其昌在《画禅室随笔》中写道：“读万卷书，行万里路。”"
        "读书与行路，都是拓展生命视野的方式。进入初中以来，同学们在语文课上读过朱自清笔下“吹面不寒杨柳风”的《春》，"
        "读过老舍笔下“温晴”的《济南的冬天》，也在《论语》中感受过“学而时习之，不亦说乎”的求学之乐。"
        "校园里，图书馆新设了“经典诵读角”，同学们或低声吟诵，或默读批注，书香与晨光一起流淌。"
        "有同学把阅读比作（yù）见老友——每一次翻开书页，都是一次温暖的相逢；也有同学把好书推荐卡贴在走廊里，"
        "让文字在校园里传（bō）。当然，阅读并不只是“看热闹”，更要学会思考：作者为什么这样写？"
        "想象与真实之间有着怎样的联系？这些思考，会让阅读由“浅尝”走向“深耕”。"
        "本学期，我们还将走进《朝花夕拾》，在鲁迅的回忆里看见童年；走进《西游记》，在神魔斗法中感受想象的魅力。"
        "愿同学们在阅读中涵养品格、开阔眼界，让好书成为成长路上最长情的陪伴。"
    )
    add_p(doc, passage1, font="楷体", indent=0.74)

    q1 = "1. 文中注音不正确的一项是（   ）"
    add_p(doc, q1)
    add_p(doc, "A. 遇（yù）\tB. 播（bō）\tC. 涵（hán）\tD. 长（zhǎng）")
    if with_answers:
        add_p(doc, "【答案】D")
        add_p(doc, "【解析】“长情”的“长”应读 cháng，此处误读为 zhǎng。故选D。")

    q2 = "2. 文中加点词语有错别字的一项是（   ）"
    add_p(doc, q2)
    add_p(doc, "A. 流淌\tB. 深耕\tC. 浅尝则止\tD. 涵养")
    if with_answers:
        add_p(doc, "【答案】C")
        add_p(doc, "【解析】“浅尝则止”应为“浅尝辄止”。故选C。")

    q3 = "3. 文中横线上应填入的词语，最恰当的一项是（   ）"
    add_p(doc, "有同学把阅读比作____老友，让文字在校园里____。")
    add_p(doc, "A. 重逢  传播\tB. 邂逅  传布\tC. 重逢  传布\tD. 邂逅  传播")
    if with_answers:
        add_p(doc, "【答案】A")
        add_p(doc, "【解析】“重逢”与“遇见老友”呼应；“传播”与“文字在校园”搭配恰当。故选A。")

    q4 = "4. 下列句子没有语病的一项是（   ）"
    add_p(doc, "A. 通过设立“经典诵读角”，使同学们的阅读兴趣明显提高了。")
    add_p(doc, "B. 阅读并不只是“看热闹”，更要学会思考作者为什么这样写。")
    add_p(doc, "C. 我们能否养成良好的阅读习惯，是提升语文素养的关键。")
    add_p(doc, "D. 同学们在走廊里贴了好书推荐卡，目的是为了让更多人爱上阅读。")
    if with_answers:
        add_p(doc, "【答案】B")
        add_p(doc, "【解析】A项缺主语，删“通过”或“使”；C项两面对一面；D项“目的是……为了”重复。故选B。")

    q5 = "5. 下列句子的排序，与上下文衔接最恰当的一项是（   ）"
    add_p(doc, "读书需要方法。______，______；______，______。______，才能在阅读中真正有所收获。")
    add_p(doc, "①先把握文章写了什么\t②再品味语言有什么特点\t③还要思考表达了怎样的情感或道理")
    add_p(doc, "④最后联系生活谈自己的感受\t⑤因此，我们要学会“整体—局部—探究”的阅读路径")
    add_p(doc, "A. ①②③④⑤\tB. ⑤①②③④\tC. ①③②④⑤\tD. ⑤②①③④")
    if with_answers:
        add_p(doc, "【答案】B")
        add_p(doc, "【解析】⑤提出路径，①②③④按阅读步骤排列。故选B。")

    add_p(doc, "6. 名篇名句默写。")
    add_p(doc, "（1）曹操在《观沧海》中，以“________，________”展现大海吞吐日月星辰的壮阔景象。")
    add_p(doc, "（2）王湾《次北固山下》中“________，________”蕴含新旧交替的自然理趣，历来为人称道。")
    add_p(doc, "（3）同窗分别之际，可用《论语》中的“________，________”表达惜别与勉励之情。")
    if with_answers:
        add_p(doc, "【答案】（1）日月之行  若出其中（2）海日生残夜  江春入旧年（3）有朋自远方来  不亦乐乎")
        add_p(doc, "（答对一句1分，共6分；错字、漏字、添字均不得分）")

    # ── 二、综合性学习 ──
    add_p(doc, "二、综合性学习（7~10题，共10分）", bold=True)
    add_p(doc, "（一）学习与探究（5分）", bold=True)
    add_p(doc, "学校开展“少年正是读书时”专题学习活动，请你参加。")
    add_p(doc, "材料一：", bold=True)
    add_p(doc, "某校七年级200名同学阅读时间调查（单位：人）", indent=0.74)
    add_p(doc, "每天阅读30分钟以下：68人；30—60分钟：95人；60分钟以上：37人。", indent=0.74)
    add_p(doc, "材料二：", bold=True)
    add_p(doc, (
        "于漪在《往事依依》中写道：“书，给我以广阔的天地，编织我美好的青春生活，"
        "日日用心灵品尝生活的滋味。”（选自人民教育出版社《语文》七年级上册第三单元）"
    ), font="楷体", indent=0.74)

    add_p(doc, "7. 看完材料后，小明和小丽展开了讨论。根据材料一、二将他们的对话补充完整。")
    add_p(doc, "小明：①________，因为②________________________________________________________。")
    add_p(doc, "小丽：我不这么认为。坚持阅读仍然很有意义，因为③________________________________。")
    if with_answers:
        add_p(doc, "【答案】示例：①近半数同学每天阅读不足30分钟  ②说明同学们阅读时间偏少，需要加强引导"
                       "  ③阅读能开阔视野、丰富精神生活（言之有理即可，共5分）")

    add_p(doc, "8. 假如你是班级“阅读推荐官”，要从下面三本书中推荐一本给同学，你会选择哪一本？请说明理由。")
    add_p(doc, "A.《朝花夕拾》（鲁迅）  B.《西游记》（吴承恩）  C.《昆虫记》（法布尔）")
    if with_answers:
        add_p(doc, "【答案】示例：选A。鲁迅以深情与批判回忆童年与青年，有助于我们了解作者成长经历，学习回忆性散文写法。（5分）")

    add_p(doc, "（二）名著阅读（5分）", bold=True)
    add_p(doc, "9. 《朝花夕拾》是鲁迅的散文集。请结合书中任一篇目，谈谈你对“童年”的理解。")
    if with_answers:
        add_p(doc, "【答案】示例：读《从百草园到三味书屋》，鲁迅笔下的童年既有自由快乐，也有束缚与无奈，"
                       "让我明白童年体验会影响人的一生。（5分，结合篇目2分，谈理解3分）")

    add_p(doc, "10. 学校文学社征集以“成长”为主题的演讲稿，你会选择下面哪组素材？结合作品内容说明理由。")
    add_p(doc, "史铁生——《秋天的怀念》  海伦·凯勒——《再塑生命的人》")
    if with_answers:
        add_p(doc, "【答案】示例：选海伦·凯勒。莎莉文老师用爱与耐心唤醒“我”的生命意识，"
                       "体现教育对成长的关键作用，契合“成长”主题。（5分）")

    # ── 三、阅读 ──
    add_p(doc, "三、阅读（11~23题，共34分）", bold=True)
    add_p(doc, "（一）（4分）", bold=True)
    add_p(doc, "阅读下面两首诗歌，完成下面小题。")
    add_p(doc, "【甲】", bold=True)
    add_p(doc, "闻王昌龄左迁龙标遥有此寄", WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    add_p(doc, "李白", WD_ALIGN_PARAGRAPH.CENTER)
    add_p(doc, "杨花落尽子规啼，\n闻道龙标过五溪。\n我寄愁心与明月，\n随君直到夜郎西。",
          WD_ALIGN_PARAGRAPH.CENTER, font="楷体")
    add_p(doc, "（选自《李太白全集》，人民教育出版社《语文》七年级上册第一单元）", size=10.5)

    add_p(doc, "【乙】", bold=True)
    add_p(doc, "夜雨寄北", WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    add_p(doc, "李商隐", WD_ALIGN_PARAGRAPH.CENTER)
    add_p(doc, "君问归期未有期，\n巴山夜雨涨秋池。\n何当共剪西烛，\n却话巴山夜雨时。",
          WD_ALIGN_PARAGRAPH.CENTER, font="楷体")
    add_p(doc, "（选自《李商隐诗选》，人民教育出版社《语文》七年级上册第六单元课外古诗词诵读）", size=10.5)

    add_p(doc, "11. 对甲诗的理解和赏析，不恰当的一项是（   ）")
    add_p(doc, "A. “杨花落尽子规啼”点明时令，渲染凄凉氛围，为全诗奠定感情基调。")
    add_p(doc, "B. “我寄愁心与明月”运用拟人，将抽象愁思化为可托付的具象形象。")
    add_p(doc, "C. “随君直到夜郎西”表明诗人愿随友人同赴贬所，当面表达安慰。")
    add_p(doc, "D. 全诗语言清新自然，表达了诗人对友人的深切关怀。")
    if with_answers:
        add_p(doc, "【答案】C")
        add_p(doc, "【解析】“随君”是寄月相随，非诗人亲身同往。故选C。")

    add_p(doc, "12. 这两首诗都写“雨夜”背景，但情感不同，请简要分析。")
    if with_answers:
        add_p(doc, "【答案】甲诗写对友人的牵挂与慰藉；乙诗写羁旅中对亲人的思念与重逢期盼。（各2分，共4分）")

    # （二）文言文
    add_p(doc, "（二）（10分）", bold=True)
    add_p(doc, "阅读下面文言文，完成下面小题。")
    add_p(doc, "【甲】", bold=True)
    add_p(doc, (
        "穿井得一人\n"
        "宋之丁氏，家无井而出溉汲，常一人居外。及其家穿井，告人曰：“吾穿井得一人。”"
        "有闻而传之者，曰：“丁氏穿井得一人。”国人道之，闻之于宋君。宋君令人问之于丁氏，"
        "丁氏对曰：“得一人之使，非得一人于井中也。”"
    ), font="楷体", indent=0.74)
    add_p(doc, "（选自《吕氏春秋·察传》，人民教育出版社《语文》七年级上册第六单元）", size=10.5)

    add_p(doc, "【乙】", bold=True)
    add_p(doc, (
        "陈太丘与友期行，期日中。过中不至，太丘舍去，去后乃至。元方时年七岁，门外戏。"
        "客问元方：“尊君在不？”答曰：“待君久不至，则舍去。”友人便怒：“非人哉！与人期行，相委而去。”"
        "元方曰：“君与家君期日中。日中不至，则是无信；对子骂父，则是无礼。”友人惭，下车引之。"
    ), font="楷体", indent=0.74)
    add_p(doc, "（选自《世说新语·方正》，人民教育出版社《语文》七年级上册第二单元）", size=10.5)

    add_p(doc, "13. 解释下列句子中加点词的意思。")
    add_p(doc, "（1）国人道之  道：________________  （2）闻之于宋君  闻：________________")
    add_p(doc, "（3）尊君在不  不：________________  （4）相委而去  委：________________")
    if with_answers:
        add_p(doc, "【答案】（1）讲述、说（2）听说，这里指上报（3）同“否”（4）舍弃、丢下（每空1分）")

    add_p(doc, "14. 用现代汉语翻译下列句子。")
    add_p(doc, "（1）得一人之使，非得一人于井中也。")
    add_p(doc, "（2）友人便怒：“非人哉！与人期行，相委而去。”")
    if with_answers:
        add_p(doc, "【答案】（1）得到一个人劳力，不是从井里挖出一个人。（2）友人便发怒说："
                       "“不是人啊！和别人约好同行，却丢下我先离开了。”（各2分）")

    add_p(doc, "15. 甲文讽刺了怎样的社会现象？请结合文意简要分析。")
    if with_answers:
        add_p(doc, "【答案】讽刺以讹传讹、不经核实就传播消息的现象。丁氏语有歧义，经人传播后严重失真。（3分）")

    add_p(doc, "16. 元方的做法是否符合“求真守礼”的要求？请结合甲、乙两文谈谈你的认识。")
    if with_answers:
        add_p(doc, "【答案】符合。乙文元方指出友人无信无礼，体现守礼；甲文启示我们传话要准确核实，体现求真。（3分）")

    # （三）说明文
    add_p(doc, "（三）（8分）", bold=True)
    add_p(doc, "阅读下面文章，完成下面小题。")
    add_p(doc, "大雁归来（节选）", WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    add_p(doc, "［美］利奥波德", WD_ALIGN_PARAGRAPH.CENTER)

    essay3 = [
        "①每年三月，大雁从南方飞到北方。它们排着整齐的队伍，在天空划出优美的曲线。",
        "②大雁是一种候鸟，它们的迁徙路线几乎固定不变。它们用叫声联络同伴，用翅膀丈量季节。",
        "③春天，它们带来泥土的湿润气息，也带来播种的消息。农民说，看见大雁北归，就知道该春耕了。",
        "④然而，随着湿地减少、气候变化，大雁的栖息地正在缩小。保护湿地，就是保护这些“春天的信使”。",
        "⑤我们应该学会倾听自然的声音。大雁归来，不仅是季节更替的标志，更提醒人类与自然和谐相处。",
    ]
    for para in essay3:
        add_p(doc, para, font="楷体", indent=0.74)
    add_p(doc, "（选自［美］利奥波德《沙乡年鉴·大雁归来》，朱曼译，人民教育出版社《语文》七年级上册第五单元）", size=10.5)

    add_p(doc, "17. 阅读选文，完成下面的学习任务单。")
    add_p(doc, "（1）大雁迁徙的时间是________，它们被比作________。（2分）")
    add_p(doc, "（2）选文从哪几个方面写了“大雁归来”的意义？请简要概括。（3分）")
    add_p(doc, "（3）第④段在文中有什么作用？（3分）")
    if with_answers:
        add_p(doc, "【答案】（1）三月（春天）  春天的信使（2）标示季节/带来春耕消息/提醒人与自然和谐相处"
                       "（3）补充现实问题，引出保护湿地、保护自然的主题，深化中心。")

    add_p(doc, "18. 选文语言生动形象，请以第①段画线句为例进行分析。")
    add_p(doc, "它们排着整齐的队伍，在天空划出优美的曲线。")
    if with_answers:
        add_p(doc, "【答案】运用拟人，把大雁飞行队形比作“队伍”“曲线”，生动展现大雁迁徙的整齐优美。（3分）")

    add_p(doc, "19. 链接单元课文《狼》（蒲松龄《聊斋志异》），有人认为“人与自然可以和谐相处”，"
               "也有人认为“人与自然存在冲突”。结合选文及你的阅读体验谈看法。")
    if with_answers:
        add_p(doc, "【答案】示例：我赞同和谐相处。选文写保护大雁与湿地，《狼》写人与动物的冲突，"
                       "启示我们要尊重自然规律，而非一味征服。（3分，观点1分，理由2分）")

    # （四）记叙文
    add_p(doc, "（四）（12分）", bold=True)
    add_p(doc, "阅读下面文章，完成下面小题。")
    add_p(doc, "秋天的怀念（节选）", WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    add_p(doc, "史铁生", WD_ALIGN_PARAGRAPH.CENTER)

    essay4 = [
        "①双腿瘫痪后，我的脾气变得暴怒无常。望着天上北归的雁阵，我会突然把面前的玻璃砸碎；",
        "听着李谷一甜美的歌声，我会猛地把手边的东西摔向四周的墙壁。母亲就悄悄地躲出去，",
        "在我看不见的地方偷偷地听着我的动静。当一切恢复沉寂，她又悄悄地进来，眼边红红的，看着我。",
        "②“听说北海的花儿都开了，我推着你去走走。”她总是这么说。母亲喜欢花，可自从我的腿瘫痪，",
        "她侍弄的那些花都死了。“不，我不去！”我狠命地捶打这两条可恨的腿，喊着：“我可活什么劲！”",
        "③母亲扑过来抓住我的手，忍住哭声说：“咱娘儿俩，好好儿活，好好儿活……”",
    ]
    for para in essay4:
        add_p(doc, para, font="楷体", indent=0.74)
    add_p(doc, "（选自史铁生《秋天的怀念》，人民教育出版社《语文》七年级上册第二单元）", size=10.5)

    add_p(doc, "20. 第①段中“悄悄地”“偷偷地”表现了母亲怎样的心理？请简要概括。")
    if with_answers:
        add_p(doc, "【答案】体贴儿子、担心刺激儿子、强忍内心痛苦却又无微不至地关怀儿子。（3分）")

    add_p(doc, "21. 结合选文内容，分析第③段加点词“忍住哭声”的含义。")
    if with_answers:
        add_p(doc, "【答案】母亲因儿子瘫痪而悲痛，却为了不加重儿子负担而压抑哭泣，表现母爱的隐忍与坚强。（3分）")

    add_p(doc, "22. 文中母亲说“好好儿活”，与《皇帝的新装》中孩子说“他没穿衣服”有何不同？"
               "请从人物与主题角度简要分析。")
    if with_answers:
        add_p(doc, "【答案】母亲的话是关爱与鼓励，主题在亲情与生命；孩子的话揭露虚伪，主题在真话与盲从。（3分）")

    add_p(doc, "23. 班级将举行“想象与真实”主题读书会，请结合选文和链接材料，谈谈文学阅读对我们认识生活的意义。")
    add_p(doc, "材料一：安徒生《皇帝的新装》写众人不敢说真话，唯有孩子道出真相。（人民教育出版社《语文》七年级上册第六单元）")
    add_p(doc, "材料二：袁珂《女娲造人》以神话想象解释人类起源，表达对生命与创造的礼赞。（同上）")
    if with_answers:
        add_p(doc, "【答案】文学以想象写真实：《秋天的怀念》写亲情，《皇帝的新装》写人性，《女娲造人》写生命。"
                       "阅读帮助我们理解生活、树立正确价值观。（3分）")

    # ── 四、写作 ──
    add_p(doc, "四、写作（40分）", bold=True)
    add_p(doc, "请从下面的题目中任选一题完成写作。（40分）")
    add_p(doc, "24. 题目：", bold=True)
    add_p(doc, (
        "进入初中，你读过许多动人的文章，见过许多难忘的画面。某一本书、某一次散步、"
        "某一个微笑，都可能成为记忆中闪光的坐标……"
    ), indent=0.74)
    add_p(doc, "请以“这也是一堂语文课”为题，写一篇作文。")
    add_p(doc, "要求：")
    add_p(doc, "（1）不少于600字，书写工整，字迹清楚；")
    add_p(doc, "（2）立意自定，文体自选（诗歌除外）；")
    add_p(doc, "（3）不得在文中泄露个人和学校信息。")

    add_p(doc, "25. 题目二：_________的力量", bold=True)
    add_p(doc, "要求：")
    add_p(doc, "（1）若选题目二，应将题目补充完整；")
    add_p(doc, "（2）立意自定，文体自选（诗歌除外）；")
    add_p(doc, "（3）不得在文中泄露个人和学校信息；")
    add_p(doc, "（4）不少于600字，书写工整，字迹清楚。")

    if with_answers:
        add_p(doc, "【写作指导】24题应写出“语文学习”在课堂之外的真实体验；25题可写“阅读”“想象”“母爱”等，"
                       "结合七年级上册单元主题，注意细节与真情实感。（40分）")

    return doc


def main():
    exam = build_exam(False)
    exam.save(str(OUT / "七年级上册期末考试卷（原卷版）.docx"))
    print("✓ 七年级上册期末考试卷（原卷版）.docx")

    ans = build_exam(True)
    ans.save(str(OUT / "七年级上册期末考试卷（参考答案及解析）.docx"))
    print("✓ 七年级上册期末考试卷（参考答案及解析）.docx")

    # English copies
    import shutil
    shutil.copy2(OUT / "七年级上册期末考试卷（原卷版）.docx", OUT / "grade7-final-exam-paper.docx")
    shutil.copy2(OUT / "七年级上册期末考试卷（参考答案及解析）.docx", OUT / "grade7-final-exam-answers.docx")


if __name__ == "__main__":
    main()
