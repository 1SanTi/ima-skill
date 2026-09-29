# -*- coding: utf-8 -*-
"""
在原《通知》docx 上直接填写其余附件，保持原表格式完全不变。
生成：汇总表（附件4）、征求意见表（附件5）、先进事迹材料（附件6）。

用法：
    python3 generate_attachments.py --template 通知.docx --outdir 输出目录 [--data 候选人.json]
"""
import argparse
import json
import os
import shutil
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

DEFAULT_DATA = {
    # 汇总表（附件4）
    "序号": "1",
    "姓名": "周立本",
    "工作单位": "云溪县青竹镇中学",
    "性别": "男",
    "民族": "汉",
    "出生年月": "197208",
    "政治面貌": "群众",
    "学历": "本科",
    "行政职务及年限": "",
    "专业技术职务": "中小学一级教师",
    "教龄": "31",
    "乡村工作时间": "31",
    "班主任时间": "14",
    "推荐奖励名称": "云溪县优秀班主任",
    "人选类型": "教师（初中）",
    "电话": "13800000001",
    "100字事迹": "该同志连续五年担任九年级班主任，长期扎根山村，坚守乡村教育一线。工作中恪尽职守、爱生如子，以高度的责任心和无私的奉献精神默默耕耘。所带班级中考成绩在全乡（县）名列前茅，教学成果显著，深受学生爱戴和家长信赖，是乡村教育战线上的优秀典型。",
    "民主测评满意率": "",
    # 征求意见表（附件5）
    "职务": "教师",
    # 事迹材料（附件6）
    "事迹标题": "申报“云溪县优秀班主任”事迹材料",
    "综合表现": "该同志忠诚党和人民的教育事业，全面贯彻党的教育方针，严格遵守教育法律法规和师德规范，为人师表、爱岗敬业、立德树人。从教三十一年，始终扎根乡村教育一线，以高度的责任感和使命感默默耕耘、无私奉献。政治立场坚定，工作态度端正，服从组织安排，纪律意识强，不论是学校安排的教学任务还是非教学工作都欣然接受、认真完成。团结同事、关爱学生、注重家校共育，深受学生爱戴、家长信赖和同事好评。历年年度考核均为合格及以上等次，多次获评优秀等次。",
    "主要业绩": "该同志长期坚守乡村教育一线，勇挑重担、实绩突出。\n\n（一）勇挑重担，教学成绩优异。近年来，因大批青年骨干教师考入县城，学校师资严重紧张。为顾全大局，他毫不犹豫地同时承担起初三数学、初三体育和初二生物三门中考科目的教学任务。繁重的教学任务没有使他退缩，反而激励他更加努力。他精心备课、认真上课、个别辅导、科学训练、严格要求，2021年中考三门科目均取得优异成绩，为学校圆满完成中考目标作出了应有贡献。他所带的班级中考成绩逐年攀升，2024年位列全县第三名、2025年位列全县第四名，创下学校近年最好成绩。\n\n（二）倾心育人，班级管理成效显著。他连续担任七届班主任，班级管理规范有序、班风正、学风浓。2018年上学期，临近中考仅一个多月时，他因长期劳累突发胃出血，连夜被送往医院救治。病情稍有好转，他便提前出院，带着大病初愈的身体立即投入工作，并主动向学校建议利用双休日为学生补课，不计任何报酬，独自排课表、补课程，连续补了四个双休日。功夫不负有心人，当年他所带的89班全县中考综合排名位列第22名，较前一年前进十一名，一举摘掉了学校长期落后的帽子。他爱生如子，善于站在学生和家长的角度开展教育，所带班级学生无一人失学，深受学生喜欢和爱戴。\n\n（三）甘于奉献，全面支持学校工作。作为一所乡镇中学，学校缺乏专职体育教师。2012年，学校面临全县中小学生运动会无人带队的困境，他自告奋勇接下训练任务。此后每年，他都利用清晨、傍晚和国庆假期、双休日，带领运动员到一公里外的小学运动场训练，虚心向小学体育教师请教训练方法，科学制定训练计划。他所带的队伍在所报12个项目中取得7个第一名、3个第二名的优异成绩，并勇夺团体总分第一名，此后学校体育成绩多次进入全县前八名，他本人也多次获评“县优秀教练员”。\n\n该同志以“吃得苦、霸得蛮、耐得烦”的精神，在平凡岗位上忘我工作、默默奉献，2024年获评云溪县优秀乡村教师。他用自己的实际行动，生动诠释了一名新时代乡村班主任的使命与担当，是乡村教育战线上的优秀典型。",
}


def strip_space(s):
    return s.replace(" ", "").replace("\u3000", "").replace("\t", "").replace("\n", "")


def save_doc(doc, out):
    """保存文档，中文路径先存英文临时文件再重命名，规避 zip 写入异常。"""
    if any(ord(c) > 127 for c in out):
        tmp = os.path.join(os.path.dirname(out) or ".", "_gen_tmp.docx")
        doc.save(tmp)
        if os.path.exists(out):
            os.remove(out)
        os.rename(tmp, out)
    else:
        save_doc(doc, out)


def logic_cells(row):
    seen = []
    for c in row.cells:
        if c._tc not in [s._tc for s in seen]:
            seen.append(c)
    return seen


def fill_cell(cell, text, align="center", font="FangSong", size=10.5):
    # 删除除第一个外的所有段落，避免残留旧内容
    for p in cell.paragraphs[1:]:
        p._element.getparent().remove(p._element)
    p = cell.paragraphs[0]
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    if align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "left":
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    else:
        p.alignment = None
    run = p.add_run(text)
    run.font.name = font
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font)
    run.font.size = Pt(size)


def extract_range(template, out, start_pred, end_pred):
    """复制模板，删除 start 之前、end 之后的内容，返回 doc。"""
    shutil.copy(template, out)
    doc = Document(out)
    body = doc.element.body
    children = list(body.iterchildren())
    start_idx = end_idx = None
    for i, ch in enumerate(children):
        if ch.tag != qn("w:p"):
            continue
        t = strip_space("".join(ch.itertext()))
        if start_idx is None and start_pred(t):
            start_idx = i
        elif start_idx is not None and end_pred(t):
            end_idx = i
            break
    if start_idx is None:
        raise SystemExit("未找到起始段落")
    if end_idx is None:
        end_idx = len(children)  # 直到文档末尾
    for i, ch in enumerate(children):
        if ch.tag in (qn("w:p"), qn("w:tbl")):
            if i < start_idx or i >= end_idx:
                body.remove(ch)
    return doc


def gen_summary(template, out, d):
    doc = extract_range(
        template, out,
        start_pred=lambda t: "附件4" in t and "汇总表" in t,
        end_pred=lambda t: "征求意见表" in t,
    )
    tbl = doc.tables[0]  # 删除后仅剩汇总表
    row = tbl.rows[1]  # 第一个数据行（替换示例）
    cols = [
        d["序号"], d["姓名"], d["工作单位"], d["性别"], d["民族"], d["出生年月"],
        d["政治面貌"], d["学历"], d["行政职务及年限"], d["专业技术职务"],
        d["教龄"], d["乡村工作时间"], d["班主任时间"], d["推荐奖励名称"],
        d["人选类型"], d["电话"], d["100字事迹"], d["民主测评满意率"],
    ]
    lc = logic_cells(row)
    for i, val in enumerate(cols):
        align = "left" if i in (16,) else "center"
        fill_cell(lc[i], val, align=align, size=10.5)
    save_doc(doc, out)


def gen_opinion(template, out, d):
    doc = extract_range(
        template, out,
        start_pred=lambda t: "征求意见表" in t,
        end_pred=lambda t: "正文内容中不能反映" in t,
    )
    # 填写姓名/单位/职务/推荐类别名称
    for p in doc.paragraphs:
        if "姓名" in p.text and "单位" in p.text and "职务" in p.text:
            runs = p.runs
            labels = {"姓名：": d["姓名"], "单位：": d["工作单位"],
                      "职务：": d["职务"], "推荐类别名称：": d["推荐奖励名称"]}
            for i, r in enumerate(runs):
                for label, val in list(labels.items()):
                    if r.text == label:
                        if i + 1 < len(runs) and runs[i + 1].text.strip() == "":
                            runs[i + 1].text = val
                        else:
                            r.text = label + val
                        del labels[label]
                        break
            break
    save_doc(doc, out)


def gen_deeds(template, out, d):
    doc = Document()
    # 标题：黑体小二（18pt）居中
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(d["事迹标题"])
    r.font.name = "SimHei"
    r._element.rPr.rFonts.set(qn("w:eastAsia"), "SimHei")
    r.font.size = Pt(18)
    r.font.bold = True
    # 正文：仿宋小三（15pt），行距28磅
    def add_body(text, bold=False):
        para = doc.add_paragraph()
        para.paragraph_format.line_spacing = Pt(28)
        para.paragraph_format.first_line_indent = Pt(30)  # 首行缩进2字符
        run = para.add_run(text)
        run.font.name = "FangSong"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "FangSong")
        run.font.size = Pt(15)
        run.font.bold = bold
        return para
    # 一、综合表现
    add_body("一、综合表现", bold=True)
    add_body(d["综合表现"])
    # 二、主要业绩
    add_body("二、主要业绩", bold=True)
    # 主要业绩可能含 \n 分段
    for seg in d["主要业绩"].split("\n"):
        if seg.strip():
            add_body(seg)
    save_doc(doc, out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True)
    ap.add_argument("--outdir", required=True)
    ap.add_argument("--data", default=None)
    args = ap.parse_args()

    d = DEFAULT_DATA
    if args.data:
        with open(args.data, "r", encoding="utf-8") as f:
            d = json.load(f)

    os.makedirs(args.outdir, exist_ok=True)
    name = d["姓名"]

    def safe_path(fn):
        return os.path.join(args.outdir, fn)

    # 汇总表
    summary_out = safe_path(f"3_拟推荐人选汇总表_{name}.docx")
    gen_summary(args.template, summary_out, d)
    # 征求意见表
    opinion_out = safe_path(f"4_拟推荐候选人征求意见表_{name}.docx")
    gen_opinion(args.template, opinion_out, d)
    # 事迹材料
    deeds_out = safe_path(f"5_先进事迹材料_{name}.docx")
    gen_deeds(args.template, deeds_out, d)

    print("已生成：")
    print(" ", summary_out)
    print(" ", opinion_out)
    print(" ", deeds_out)


if __name__ == "__main__":
    main()
