# -*- coding: utf-8 -*-
"""生成八年级上册第六章《人体生命活动的调节》单元测试卷三份文档：
1) 试卷卷  2) 参考答案及评分标准  3) 命题多维度细目表
"""
import os
import shutil
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

OUT = "/sandbox/workspace/output/biovol6/"
os.makedirs(OUT, exist_ok=True)

SONG = "宋体"
HEI = "黑体"
KAI = "楷体"


def set_run(run, name=SONG, size=10.5, bold=False, color=None, italic=False):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = RGBColor(*color)
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), name)
    rFonts.set(qn('w:ascii'), name)
    rFonts.set(qn('w:hAnsi'), name)
    return run


def para(doc, text="", name=SONG, size=10.5, bold=False, align=None,
         color=None, space_before=0, space_after=2, indent=None,
         first_line=None, line_spacing=1.25, italic=False):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line_spacing
    if indent is not None:
        pf.left_indent = Cm(indent)
    if first_line is not None:
        pf.first_line_indent = Cm(first_line)
    if text:
        set_run(p.add_run(text), name, size, bold, color, italic)
    return p


def options_two(doc, opts):
    """两列排版选项：A B 一行，C D 一行。"""
    for pair in [(opts[0], opts[1]), (opts[2], opts[3])]:
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(2)
        pf.line_spacing = 1.25
        pf.left_indent = Cm(0.6)
        pf.tab_stops.add_tab_stop(Cm(8.6), WD_TAB_ALIGNMENT.LEFT)
        set_run(p.add_run(pair[0] + "\t" + pair[1]), SONG, 10.5)


def options_one(doc, opts):
    for o in opts:
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(1)
        pf.line_spacing = 1.2
        pf.left_indent = Cm(0.6)
        set_run(p.add_run(o), SONG, 10.5)


def add_pic(doc, fname, width=11.0, caption=None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(1)
    p.add_run().add_picture(OUT + fname, width=Cm(width))
    if caption:
        cp = doc.add_paragraph()
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cp.paragraph_format.space_before = Pt(0)
        cp.paragraph_format.space_after = Pt(4)
        set_run(cp.add_run(caption), KAI, 9, False, (0x40, 0x40, 0x40))
    return p


def blank_line(doc, n=1):
    for _ in range(n):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        set_run(p.add_run(""), SONG, 10.5)


# ============================================================
# 题目数据
# ============================================================
CHOICES = [
    dict(n=1, stem="关于眼球结构与功能的叙述，错误的是（　　）",
         opts=["A．角膜无色透明，能透过光线并对光线有折射作用",
               "B．虹膜中央的瞳孔是光进入眼球的通道",
               "C．视网膜能感受光的刺激，是形成视觉的部位",
               "D．晶状体透明而有弹性，能折射光线，其曲度可由睫状体调节"],
         layout="one", tag="[素养：生命观念｜难度：易]"),
    dict(n=2, stem="下列关于视觉形成过程的叙述，正确的是（　　）",
         opts=["A．物像形成的部位是大脑皮层",
               "B．视网膜上对光敏感的感觉细胞接受刺激后产生神经冲动，经视神经传到大脑皮层的特定区域，形成视觉",
               "C．光线经角膜、瞳孔、玻璃体、晶状体折射后到达视网膜",
               "D．只要眼球的结构完好，人就一定能形成视觉"],
         layout="one", tag="[素养：生命观念｜难度：中]"),
    dict(n=3, stem="小明看远处物体时感到模糊，经医生检查，其眼球前后径过长。下列分析正确的是（　　）",
         opts=["A．物像落在视网膜后方，需佩戴凸透镜矫正",
               "B．物像落在视网膜前方，需佩戴凸透镜矫正",
               "C．物像落在视网膜后方，需佩戴凹透镜矫正",
               "D．物像落在视网膜前方，需佩戴凹透镜矫正"],
         layout="one", tag="[素养：生命观念｜难度：中]"),
    dict(n=4, stem="下列用眼习惯中，不利于保护视力的是（　　）",
         opts=["A．读写时眼与书本保持约33厘米的距离",
               "B．使用电子产品30～40分钟后远眺休息",
               "C．在光线昏暗的公交车上看书",
               "D．每天保证2小时以上的户外活动"],
         layout="one", tag="[素养：态度责任｜难度：易]"),
    dict(n=5, stem="眼球中，能够调节进入眼内光线多少的结构是（　　）",
         opts=["A．角膜", "B．瞳孔", "C．巩膜", "D．视网膜"],
         layout="two", tag="[素养：生命观念｜难度：中]"),
    dict(n=6, stem="关于耳的结构与功能的叙述，错误的是（　　）",
         opts=["A．耳郭和外耳道组成外耳，能收集和传导声波",
               "B．鼓膜能把声波转换为振动",
               "C．耳蜗内有相应的感觉细胞，能接受刺激并产生神经冲动",
               "D．半规管和前庭与听觉的形成直接有关"],
         layout="one", tag="[素养：生命观念｜难度：中]"),
    dict(n=7, stem="下列关于听觉形成过程的叙述，正确的是（　　）",
         opts=["A．声波→鼓膜→耳蜗→听小骨→听神经→大脑皮层",
               "B．听觉在大脑皮层形成，耳蜗只起传导声波的作用",
               "C．声波→外耳道→鼓膜→听小骨→耳蜗→听神经→大脑皮层",
               "D．只要鼓膜完好，人就一定能形成听觉"],
         layout="one", tag="[素养：生命观念｜难度：中]"),
    dict(n=8, stem="下列做法中，有利于保护耳和听力的是（　　）",
         opts=["A．遇到巨大声响时迅速张口或闭嘴堵耳",
               "B．长时间大音量使用耳机听音乐",
               "C．用尖锐的物体掏挖外耳道",
               "D．擤鼻涕时捏紧两个鼻孔用力擤"],
         layout="one", tag="[素养：态度责任｜难度：易]"),
    dict(n=9, stem="神经系统由脑、脊髓和与它们相连的神经组成。下列关于神经系统的叙述，正确的是（　　）",
         opts=["A．脑和脊髓是神经系统的周围部分",
               "B．脑和脊髓是神经系统的中枢部分，脑神经和脊神经是周围部分",
               "C．脑神经和脊神经属于神经系统的中枢部分",
               "D．神经系统只包括脑和脊髓两部分"],
         layout="one", tag="[素养：生命观念｜难度：中]"),
    dict(n=10, stem="某患者因脑部受伤，出现身体平衡失调、走路不稳的症状。据此推测，其受损部位最可能是（　　）",
         opts=["A．小脑", "B．大脑皮层", "C．脑干", "D．脊髓"],
         layout="two", tag="[素养：生命观念｜难度：中]"),
    dict(n=11, stem="脑干中有些部位专门调节心跳、呼吸等基本生命活动。若这些部位受到严重损伤，最可能出现的后果是（　　）",
         opts=["A．身体平衡失调",
               "B．无法完成语言交流",
               "C．心跳和呼吸停止，危及生命",
               "D．运动协调能力下降"],
         layout="one", tag="[素养：生命观念｜难度：中]"),
    dict(n=12, stem="神经元是神经系统结构和功能的基本单位。下列关于神经元的叙述，正确的是（　　）",
         opts=["A．神经元一般具有一个较长的树突和多个较短的轴突",
               "B．神经元包括胞体和突起两部分",
               "C．神经纤维的末端分支叫作神经中枢",
               "D．神经元不能传导神经冲动"],
         layout="one", tag="[素养：生命观念｜难度：中]"),
    dict(n=13, stem="神经纤维集结成束，外面包有膜，构成了（　　）",
         opts=["A．神经元", "B．神经中枢", "C．神经末梢", "D．神经"],
         layout="two", tag="[素养：生命观念｜难度：易]"),
    dict(n=14, stem="机体通过神经系统，对外界或内部的各种刺激所作出的有规律的反应，叫作（　　）",
         opts=["A．反射", "B．反射弧", "C．神经中枢", "D．激素调节"],
         layout="two", tag="[素养：生命观念｜难度：易]"),
    dict(n=15, stem="反射弧是反射的结构基础。下列说法正确的是（　　）",
         opts=["A．反射弧中任何一段受损，其他部分仍能独立完成反射",
               "B．效应器只能由肌肉构成，不能由腺体构成",
               "C．反射弧由感受器、传入神经、神经中枢、传出神经、效应器五部分组成",
               "D．感受器受到刺激后能直接引起效应器反应，不需要神经中枢参与"],
         layout="one", tag="[素养：生命观念｜难度：中]"),
    dict(n=16, stem="短跑运动员听到发令枪响后迅速冲出起跑线。下列关于该反射的分析，正确的是（　　）",
         opts=["A．这是人生来就有的非条件反射",
               "B．这是后天形成的条件反射，需要大脑皮层参与",
               "C．该反射的结构基础是反射弧，不需要大脑参与",
               "D．该反射与膝跳反射属于同一类型"],
         layout="one", tag="[素养：生命观念｜难度：中]"),
    dict(n=17, stem="与语言文字有关的反射是人类所特有的。“谈虎色变”这一反射的神经中枢位于（　　）",
         opts=["A．脊髓", "B．大脑皮层", "C．小脑", "D．脑干"],
         layout="two", tag="[素养：生命观念｜难度：中]"),
    dict(n=18, stem="下列反射中，属于非条件反射的是（　　）",
         opts=["A．膝跳反射", "B．望梅止渴", "C．惊弓之鸟", "D．谈虎色变"],
         layout="two", tag="[素养：生命观念｜难度：易]"),
    dict(n=19, stem="人的运动系统主要由（　　）组成",
         opts=["A．骨骼、骨骼肌和神经",
               "B．骨、关节和神经",
               "C．骨和骨骼肌",
               "D．骨、关节和肌肉"],
         layout="one", tag="[素养：生命观念｜难度：易]"),
    dict(n=20, stem="关节既牢固又灵活。下列结构与关节的灵活性直接相关的是（　　）",
         opts=["A．关节囊及其周围的韧带",
               "B．关节腔内的滑液和关节软骨",
               "C．关节头",
               "D．骨松质"],
         layout="one", tag="[素养：生命观念｜难度：中]"),
    dict(n=21, stem="做屈肘动作时，上臂肌肉的状态是（　　）",
         opts=["A．肱二头肌舒张，肱三头肌收缩",
               "B．肱二头肌和肱三头肌同时收缩",
               "C．肱二头肌收缩，肱三头肌舒张",
               "D．肱二头肌和肱三头肌同时舒张"],
         layout="one", tag="[素养：生命观念｜难度：易]"),
    dict(n=22, stem="关于人体运动的产生，下列说法正确的是（　　）",
         opts=["A．运动仅靠运动系统就能完成",
               "B．人体的任何一个动作都只由一组肌肉完成",
               "C．骨骼肌的收缩与神经支配无关",
               "D．骨骼肌受神经传来的冲动而收缩，牵动骨绕关节活动，从而产生运动"],
         layout="one", tag="[素养：生命观念｜难度：中]"),
    dict(n=23, stem="下列关于内分泌腺的叙述，正确的是（　　）",
         opts=["A．内分泌腺有导管，分泌物由导管排出",
               "B．激素在人体内含量多，作用小",
               "C．甲状腺不属于内分泌腺",
               "D．内分泌腺分泌的激素直接进入腺体内的毛细血管，随血液循环运输"],
         layout="one", tag="[素养：生命观念｜难度：中]"),
    dict(n=24, stem="某儿童身材矮小但智力发育正常，经检查是幼年时某种激素分泌不足所致。该激素是（　　）",
         opts=["A．生长激素", "B．甲状腺激素", "C．胰岛素", "D．肾上腺素"],
         layout="two", tag="[素养：生命观念｜难度：中]"),
    dict(n=25, stem="胰岛素能促进组织细胞对葡萄糖的吸收、利用和转化。当人体内胰岛素分泌不足时，最可能出现的症状是（　　）",
         opts=["A．血糖浓度过高，出现多尿、多饮、多食和体重减轻",
               "B．血糖浓度过低，出现头晕、心慌",
               "C．生长迅速，身材过高",
               "D．情绪容易激动，身体消瘦"],
         layout="one", tag="[素养：生命观念｜难度：中]"),
]

ANSWERS = ["C", "B", "D", "C", "B", "D", "C", "A", "B", "A",
           "C", "B", "D", "A", "C", "B", "B", "A", "D", "B",
           "C", "D", "D", "A", "A"]

# 参考文档（细目表与答案共用）
META = [
    # n, 类型, 考点, 题源, 改编方式, 素养, 难度
    (1, "选择题/眼球结构", "眼球各结构的功能", "P2 探究题·眼球结构（找错误项）", "改表述", "生命观念", "0.80易"),
    (2, "选择题/视觉形成", "视觉的形成过程", "2025辽宁试题·视觉形成", "改表述", "生命观念", "0.65中"),
    (3, "选择题/近视矫正", "近视的成因与矫正", "课后提升·昆明模拟（近视）", "改表述", "生命观念", "0.70中"),
    (4, "选择题/用眼卫生", "科学用眼、预防近视", "第4题·用眼卫生", "调选项顺序", "态度责任", "0.85易"),
    (5, "选择题/眼球结构", "瞳孔调节进光量", "第7题·眼球结构", "调选项顺序", "生命观念", "0.75中"),
    (6, "选择题/耳的结构", "耳各结构的功能", "耳和听觉·探究题", "改表述", "生命观念", "0.75中"),
    (7, "选择题/听觉形成", "听觉的形成过程", "课后提升·贺州模拟（听觉）", "调选项顺序", "生命观念", "0.70中"),
    (8, "选择题/护耳", "保护耳和听力", "第4题·成语与护耳", "改表述", "态度责任", "0.85易"),
    (9, "选择题/神经系统组成", "神经系统的组成", "2025济南模拟·神经系统组成", "改表述", "生命观念", "0.70中"),
    (10, "选择题/脑的功能", "小脑协调运动、维持平衡", "第7题·脑的神经中枢", "改表述", "生命观念", "0.70中"),
    (11, "选择题/脑干功能", "脑干调节基本生命活动", "第5题·情境题（脑干）", "改表述", "生命观念", "0.75中"),
    (12, "选择题/神经元", "神经元的结构", "第6题·神经元结构", "改表述", "生命观念", "0.65中"),
    (13, "选择题/神经纤维", "神经纤维与神经的关系", "弹性作业（P11）第9题", "调选项顺序", "生命观念", "0.80易"),
    (14, "选择题/反射概念", "反射是神经调节的基本方式", "课堂探究一·神经调节的基本方式", "改表述", "生命观念", "0.85易"),
    (15, "选择题/反射弧", "反射弧的组成", "课后提升·徐州模拟第2题", "改表述", "生命观念", "0.65中"),
    (16, "选择题/反射类型", "条件反射与非条件反射", "课后提升第1题·昆明模拟（情境）", "改表述", "生命观念", "0.65中"),
    (17, "选择题/语言反射", "人类特有的语言反射中枢", "第5题·人工智能（模拟神经连接）", "改表述", "科学思维", "0.70中"),
    (18, "选择题/反射类型", "非条件反射的判断", "第7、8题·反射类型判断", "调选项顺序", "生命观念", "0.80易"),
    (19, "选择题/运动系统", "运动系统的组成", "课堂探究一·运动系统的组成", "改表述", "生命观念", "0.85易"),
    (20, "选择题/关节结构", "关节的牢固性与灵活性", "课后提升·郴州模拟（关节）", "改表述", "生命观念", "0.65中"),
    (21, "选择题/屈肘伸肘", "屈肘时肌肉的协调配合", "课堂探究二·神经系统支配下的运动", "调选项顺序", "生命观念", "0.85易"),
    (22, "选择题/运动的产生", "骨骼肌牵动骨绕关节活动", "弹性作业第10题·运动的产生", "改表述", "生命观念", "0.70中"),
    (23, "选择题/内分泌腺", "内分泌腺的特点", "课堂探究一·内分泌腺与激素", "改表述", "生命观念", "0.70中"),
    (24, "选择题/激素异常", "生长激素与侏儒症", "课后提升·绵阳模拟第7题", "改表述", "生命观念", "0.65中"),
    (25, "选择题/胰岛素", "胰岛素的功能与糖尿病", "课后提升·绵阳模拟第8题", "调选项顺序", "生命观念", "0.65中"),
    (26, "非选择题/识图分析", "眼球结构与视觉形成、近视矫正、用眼卫生", "P2/第4题/第9题（填空）·眼和视觉", "题型创新", "生命观念、态度责任", "0.65中"),
    (27, "非选择题/识图分析", "耳的结构与听觉形成、护耳", "耳和听觉·探究题第9题", "题型创新", "生命观念、态度责任", "0.60中"),
    (28, "非选择题/实验探究", "反射弧组成、反射类型、反射弧受损分析", "课堂探究一、二＋课后提升（神经调节）", "题型创新", "科学思维、探究实践", "0.50较难"),
    (29, "非选择题/识图分析", "关节结构、屈肘伸肘、运动的多系统配合", "课堂探究一、二＋课后提升（运动）", "题型创新", "生命观念、探究实践", "0.60中"),
    (30, "非选择题/资料分析", "内分泌腺、激素异常、神经与激素的关系、碘与甲状腺", "第9题（填空问答）·激素调节＋章末整合", "题型创新", "生命观念、态度责任", "0.55中"),
]

NONCHOICE = [
    dict(n=26, title="【识图分析】眼与视觉（10分）",
         pic="fig_eye.png",
         lead="人眼是人体重要的感觉器官。图1为人眼球结构示意图。请据图并结合所学知识回答：",
         items=[
             "（1）人的眼球近似球体。图中，无色透明、能对光线起折射作用且其曲度可以调节的结构是______。（2分）",
             "（2）外界物体反射来的光，经角膜、瞳孔等结构折射后，在______上形成一个物像；当该结构上对光敏感的感觉细胞接受刺激并产生神经冲动后，经视神经传到______的特定区域，人才能形成视觉。（2分）",
             "（3）如果眼球的前后径过长，或晶状体的曲度过大且不易恢复原状，远处物体反射来的光所形成的物像就会落在视网膜的______方，人看远处物体就会模糊，这种眼病称为近视。近视可以通过佩戴______（填“凸透镜”或“凹透镜”）加以矫正。（4分）",
             "（4）近年来青少年近视率不断上升。请写出一条有利于预防近视的用眼习惯：______。（2分）",
         ]),
    dict(n=27, title="【识图分析】耳与听觉（10分）",
         pic="fig_ear.png",
         lead="耳是人体重要的感觉器官。图2为耳的结构示意图。请据图并结合所学知识回答：",
         items=[
             "（1）耳分为外耳、中耳和内耳。图中属于中耳的结构有______（填图中结构名称，写出两个即可）。（2分）",
             "（2）听觉的形成：外界声波经外耳道传到______，其振动通过听小骨传到内耳，刺激______内相应的感觉细胞，这些细胞将声波信息通过听神经传给大脑的特定区域，人就产生了听觉。（4分）",
             "（3）遇到巨大声响时，应迅速张口或闭嘴堵耳，这样做可以保护______（填图中结构名称）。（2分）",
             "（4）长时间处于强噪声环境中会损伤听力。请写出一条保护听力的措施：______。（2分）",
         ]),
    dict(n=28, title="【实验探究】神经调节（10分）",
         pic="fig_reflex.png",
         lead="图3为反射弧结构模式图。某同学在老师的指导下进行了膝跳反射实验。请回答：",
         items=[
             "（1）图3中，反射弧由感受器、______、神经中枢、______和效应器五部分组成。（2分）",
             "（2）膝跳反射是人生来就有的反射，属于______（填“条件”或“非条件”）反射，其神经中枢位于______（填“大脑皮层”或“脊髓”）。（2分）",
             "（3）实验时，老师用橡胶锤叩击该同学膝盖下方的韧带（感受器），其小腿迅速抬起。若该同学的传出神经受到损伤，叩击韧带后小腿______（填“能”或“不能”）抬起，原因是______。（4分）",
             "（4）“望梅止渴”和“谈梅止渴”都属于______反射，其中与语言文字有关、人类所特有的是______。（2分）",
         ]),
    dict(n=29, title="【识图分析】神经系统支配下的运动（10分）",
         pic="fig_joint.png",
         lead="运动是在神经系统的支配下完成的。图4为关节结构示意图。请据图回答：",
         items=[
             "（1）人的运动系统主要由骨、关节和肌肉组成。图中，关节由关节面、______和______三部分组成。（2分）",
             "（2）关节面上覆盖着______，关节腔内有______，它们可以减少骨与骨之间的摩擦，使关节既牢固又灵活。（2分）",
             "（3）做屈肘动作时，以肱二头肌为主的一组肌肉处于______状态，以肱三头肌为主的另一组肌肉处于______状态。（2分）",
             "（4）运动不仅靠运动系统完成，还需要其他系统的配合。请列举一个参与运动的系统：______。（2分）",
             "（5）青少年在体育锻炼前应做好准备活动，其主要目的是防止______受到损伤。（2分）",
         ]),
    dict(n=30, title="【资料分析】激素调节（10分）",
         pic=None,
         lead="人体的生命活动既受神经系统的调节，也受激素调节的影响。请结合所学知识回答：",
         items=[
             "（1）激素是由内分泌腺或内分泌细胞产生的、对机体代谢和生理功能发挥调节作用的化学物质。与汗腺等外分泌腺不同，内分泌腺没有______，它们分泌的激素进入腺体内的______，并随血液循环输送到全身各处。（2分）",
             "（2）幼年时，如果______分泌不足，会患侏儒症；如果甲状腺激素分泌不足，则易患______，患者的智力发育会受到明显影响。（2分）",
             "（3）胰岛素的主要功能是促进组织细胞对葡萄糖的吸收、利用和转化。当人体内胰岛素分泌不足时，就可能患______，该病患者常通过______（填“口服”或“注射”）胰岛素进行治疗。（2分）",
             "（4）当人情绪激动时，大脑皮层会特别兴奋，促使肾上腺分泌较多的肾上腺素，使心跳加快、血压升高。这一实例说明，人体的生命活动受______调节和______调节的共同影响。（2分）",
             "（5）我国大力推广食用加碘食盐，其主要目的是预防因缺______而引起的______（填“地方性甲状腺肿”或“糖尿病”）。（2分）",
         ]),
]


# ============================================================
# 文档1：试卷
# ============================================================
def build_paper():
    doc = Document()
    # 页面
    sec = doc.sections[0]
    sec.top_margin = Cm(2.0)
    sec.bottom_margin = Cm(2.0)
    sec.left_margin = Cm(2.4)
    sec.right_margin = Cm(2.4)
    # 默认字体
    style = doc.styles['Normal']
    style.font.name = SONG
    style.font.size = Pt(10.5)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), SONG)

    para(doc, "八年级生物学（上册）第六章", HEI, 14, True, WD_ALIGN_PARAGRAPH.CENTER, space_after=0)
    para(doc, "《人体生命活动的调节》单元测试卷", HEI, 16, True, WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    para(doc, "（人教版　满分100分　考试时间60分钟）", KAI, 10.5, False, WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    para(doc, "班级：__________　姓名：__________　学号：__________　得分：__________",
         SONG, 10.5, False, WD_ALIGN_PARAGRAPH.LEFT, space_after=6)

    para(doc, "注意事项：", HEI, 10.5, True, space_after=2)
    notes = [
        "1.答题前，考生先将自己的姓名、班级、学号等信息填写清楚。",
        "2.选择题部分请按题号用2B铅笔填涂方框，修改时用橡皮擦干净，不留痕迹。",
        "3.非选择题部分请按题号用0.5毫米黑色墨水签字笔书写，否则作答无效。",
        "4.请勿折叠答题卡，保持字体工整、笔迹清晰、卡面清洁；本卷在草稿纸、试题卷上作答无效。",
    ]
    for t in notes:
        para(doc, t, SONG, 10, False, space_after=1)

    # 一、选择题
    para(doc, "一、选择题（本大题共25小题，每小题2分，共50分。在每小题给出的四个选项中，只有一项是符合题目要求的）",
         HEI, 11, True, space_before=8, space_after=4)
    for q in CHOICES:
        p = para(doc, f"{q['n']}. {q['stem']}", SONG, 10.5, False, space_before=3, space_after=2)
        if q['layout'] == "two":
            options_two(doc, q['opts'])
        else:
            options_one(doc, q['opts'])
        # 素养/难度标注（浅灰小字）
        para(doc, q['tag'], KAI, 8.5, False, color=(0x80, 0x80, 0x80), space_after=2)

    # 二、非选择题
    para(doc, "二、非选择题（本大题共5小题，共50分）", HEI, 11, True, space_before=10, space_after=4)
    cap = {26: "图1　眼球结构示意图", 27: "图2　耳的结构示意图",
           28: "图3　反射弧结构模式图", 29: "图4　关节结构示意图"}
    for q in NONCHOICE:
        para(doc, f"{q['n']}．{q['title']}", HEI, 10.5, True, space_before=8, space_after=2)
        para(doc, q['lead'], SONG, 10.5, False, space_after=2)
        if q['pic']:
            add_pic(doc, q['pic'], width=10.5, caption=cap.get(q['n']))
        for it in q['items']:
            p = para(doc, it, SONG, 10.5, False, space_after=3, line_spacing=1.5)

    tmp = OUT + "_paper_tmp.docx"
    doc.save(tmp)
    shutil.move(tmp, OUT + "八年级上册第六章《人体生命活动的调节》单元测试卷.docx")
    print("paper done")


# ============================================================
# 文档2：参考答案及评分标准
# ============================================================
ANS_SOL = {
    26: ["（1）晶状体（2分）",
         "（2）视网膜（1分）；大脑皮层的特定区域（或视觉中枢）（1分）",
         "（3）前（2分）；凹透镜（2分）",
         "（4）读写姿势要正确，眼与书本保持适当距离；连续用眼30～40分钟后远眺休息；不在光线昏暗或直射强光下看书；每天户外活动2小时以上等（2分，答出一条、合理即可）"],
    27: ["（1）鼓膜、听小骨（2分，每写对1个得1分；写“鼓室”也可得1分）",
         "（2）鼓膜（2分）；耳蜗（2分）",
         "（3）鼓膜（2分）",
         "（4）避免长时间使用耳机；不用尖锐物体掏耳；不让脏水进入外耳道；遇到巨大声响时迅速张口或闭嘴堵耳等（2分，答出一条、合理即可）"],
    28: ["（1）传入神经（1分）；传出神经（1分）",
         "（2）非条件（1分）；脊髓（1分）",
         "（3）不能（2分）；传出神经受损，神经冲动无法由神经中枢传到效应器（或反射弧不完整）（2分）",
         "（4）条件（1分）；谈梅止渴（1分）"],
    29: ["（1）关节囊（1分）；关节腔（1分）",
         "（2）关节软骨（1分）；滑液（1分）",
         "（3）收缩（1分）；舒张（1分）",
         "（4）呼吸系统（或循环系统、消化系统等，答出1个即可）（2分）",
         "（5）关节（或防止关节受损/脱臼）（2分）"],
    30: ["（1）导管（1分）；毛细血管（1分）",
         "（2）生长激素（1分）；呆小病（1分）",
         "（3）糖尿病（1分）；注射（1分）",
         "（4）神经（或神经系统）（1分）；激素（或激素调节）（1分）",
         "（5）碘（1分）；地方性甲状腺肿（1分）"],
}


def build_answer():
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(2.0); sec.bottom_margin = Cm(2.0)
    sec.left_margin = Cm(2.4); sec.right_margin = Cm(2.4)
    style = doc.styles['Normal']; style.font.name = SONG; style.font.size = Pt(10.5)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), SONG)

    para(doc, "八年级生物学（上册）第六章《人体生命活动的调节》", HEI, 13, True, WD_ALIGN_PARAGRAPH.CENTER, space_after=0)
    para(doc, "单元测试卷·参考答案及评分标准", HEI, 15, True, WD_ALIGN_PARAGRAPH.CENTER, space_after=6)

    para(doc, "一、选择题（每小题2分，共50分）", HEI, 11, True, space_after=3)
    for i in range(0, 25, 5):
        seg = ANSWERS[i:i + 5]
        line = "　".join([f"{i+j+1}—{seg[j]}" for j in range(len(seg))])
        # 更规范：分组展示
        line = f"{i+1}～{i+5}：" + "　".join(seg)
        para(doc, line, SONG, 10.5, False, space_after=2)

    para(doc, "二、非选择题（共50分）", HEI, 11, True, space_before=8, space_after=3)
    for q in NONCHOICE:
        para(doc, f"{q['n']}．（10分）", HEI, 10.5, True, space_before=4, space_after=2)
        for s in ANS_SOL[q['n']]:
            para(doc, s, SONG, 10.5, False, space_after=2, line_spacing=1.4)

    tmp = OUT + "_ans_tmp.docx"
    doc.save(tmp)
    shutil.move(tmp, OUT + "八年级上册第六章《人体生命活动的调节》参考答案及评分标准.docx")
    print("answer done")


# ============================================================
# 文档3：命题多维度细目表
# ============================================================
def build_spec():
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(2.0); sec.bottom_margin = Cm(2.0)
    sec.left_margin = Cm(1.8); sec.right_margin = Cm(1.8)
    style = doc.styles['Normal']; style.font.name = SONG; style.font.size = Pt(10.5)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), SONG)

    para(doc, "八年级生物学（上册）第六章《人体生命活动的调节》", HEI, 12, True, WD_ALIGN_PARAGRAPH.CENTER, space_after=0)
    para(doc, "单元测试卷·命题多维度细目表", HEI, 15, True, WD_ALIGN_PARAGRAPH.CENTER, space_after=8)

    para(doc, "一、试卷总体结构", HEI, 11.5, True, space_after=3)
    for t in ["满分：100分　　考试时长：60分钟",
              "题型结构：选择题（25题，每小题2分，共50分）；非选择题（5题，共50分）",
              "命题依据：依据《义务教育生物学课程标准（2022年版2025年修订）》概念5、5.5相关内容要求，并参照《湖南省初中学业水平考试试卷结构（2025年修订）》生物学试卷结构。",
              "素养导向：凸显生命观念、科学思维、探究实践、态度责任的核心素养立意。"]:
        para(doc, t, SONG, 10.5, False, space_after=2)

    para(doc, "二、命题多维度细目表", HEI, 11.5, True, space_before=8, space_after=3)
    headers = ["题号", "试卷结构及题目类型", "考查知识（核心考点）", "素材来源（对应《学练大视野》原题）", "改编方式", "核心素养", "难度系数"]
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = ""
        p = hdr[i].paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run(p.add_run(h), HEI, 9.5, True)
    for row in META:
        cells = table.add_row().cells
        vals = [str(row[0]), row[1], row[2], row[3], row[4], row[5], row[6]]
        for i, v in enumerate(vals):
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i in (0, 4, 6) else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_after = Pt(0)
            set_run(p.add_run(v), SONG, 9)
    # 列宽
    widths = [1.1, 2.6, 3.8, 4.2, 1.6, 2.0, 1.4]
    for r in table.rows:
        for i, w in enumerate(widths):
            r.cells[i].width = Cm(w)

    para(doc, "三、命题设计说明", HEI, 11.5, True, space_before=10, space_after=3)
    notes = [
        ("1.命题范围与依据　", "本卷聚焦人教版《生物学》八年级上册第六章“人体生命活动的调节”，涵盖第一节（眼和视觉、耳和听觉）、第二节（神经系统的组成、神经调节）、第三节（神经系统支配下的运动）、第四节（激素调节）四节六个课时。命题对照课标“概念5—5.5人体各系统在神经系统和内分泌系统的调节下……以适应机体内外环境的变化”及5.5.1～5.5.5的内容要求，题型结构参照《湖南省初中学业水平考试试卷结构（2025年修订）》：选择题25题（50分）＋非选择题5题（50分），时量60分钟，满分100分。"),
        ("2.试题来源与定位　", "本卷以复习资料《学练大视野》中被圈定的题目为命题素材，将每一道原题“定位”到相应的知识点与题型（详见细目表“素材来源”列），实现“题源可溯、考点可查”。"),
        ("3.改编策略　", "全卷不直接照搬原题，统一进行“同题创编”：约60%的题目在保持考查知识点不变的前提下调整选项或题干的表述（改表述）；约20%的题目主要调换选项顺序（调选项顺序）；约15%～17%的题目完全改变题型呈现方式而考点不变（题型创新，如将原选择题、填空题改造为识图分析、实验探究、资料分析等非选择题）。改编着眼于情境的更新与设问方式的转换，避免机械重复，突出能力的迁移。"),
        ("4.素养与难度分布　", "全卷以生命观念立意为主，兼顾科学思维、探究实践与态度责任；客观题以易、中档题为主夯实基础，主观题设置数据分析、结构分析与开放论证等高阶问题以保障区分度。全卷预估难度系数约为0.70。"),
        ("5.答案与评分　", "“参考答案及评分标准”单独成文，非选择题按空、按点赋分，开放性设问“答出其一、合理即可”，确保评分的一致性与可操作性。"),
    ]
    for head, body in notes:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.4
        set_run(p.add_run(head), HEI, 10.5, True)
        set_run(p.add_run(body), SONG, 10.5)

    tmp = OUT + "_spec_tmp.docx"
    doc.save(tmp)
    shutil.move(tmp, OUT + "八年级上册第六章《人体生命活动的调节》命题多维度细目表.docx")
    print("spec done")


if __name__ == "__main__":
    build_paper()
    build_answer()
    build_spec()
    for f in os.listdir(OUT):
        if f.endswith(".docx"):
            print(f, os.path.getsize(OUT + f), "bytes")
