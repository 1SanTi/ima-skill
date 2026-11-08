#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
为各社团生成「活动计划」+「学生名单」
学生来自各班花名册，每人最多参加 MAX_CLUBS 个社团。

用法：
    python3 generate_plans_rosters.py
配置：修改下方 CONFIG 区（学校、日期、社团列表、花名册文件、每人最多社团数）
输出：默认 /sandbox/workspace/club_plans_output/，按社团类别分子目录
"""
import warnings
warnings.filterwarnings('ignore')
import os
import random
import xlrd
import openpyxl
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ==================== CONFIG ====================
SCHOOL = '云溪县青竹镇中学'
DATE = '2026年9月10日'
OUT = '/sandbox/workspace/club_plans_output'
MAX_CLUBS = 4                                   # 每个学生最多参加的社团数
ACTIVITY_TIME = '每周课后服务时间（每周2次，每次60分钟）'  # 社团活动时间

# 社团：(名称, 类型, 活动地点, 指导老师)
CLUBS = [
    ('篮球', '体育类社团', '学校篮球场', '陈静、郑伟、冯磊、周立本'),
    ('羽毛球', '体育类社团', '学校羽毛球场（操场）', '陈静、周立本'),
    ('乒乓球', '体育类社团', '学校乒乓球室（操场）', '陈静、周立本'),
    ('排球', '体育类社团', '学校排球场（操场）', '陈静、周立本'),
    ('跳绳', '体育类社团', '学校操场', '陈静、周立本'),
    ('阅读', '美育类社团', '学校图书室（阅览室）', '万灵谱尼'),
    ('音乐', '美育类社团', '学校音乐教室', '吴晓'),
    ('美术', '美育类社团', '学校美术教室', '吴晓'),
    ('综合实践', '美育类社团', '校园及周边', '万灵谱尼'),
    ('心理', '特色类社团', '学校心理健康辅导室', '万灵谱尼'),
    ('数学解题', '特色类社团', '教室', '数学教研组教师'),
    ('语文国学文化', '特色类社团', '教室', '万灵谱尼'),
    ('新闻', '特色类社团', '教室', '万灵谱尼'),
    ('科学兴趣', '特色类社团', '学校实验室（教室）', '科学（生物）教研组教师'),
]
CLUB_NAMES = [c[0] for c in CLUBS]
CLS_ORDER = {'初一109班': 0, '初二108班': 1, '初三107班': 2}   # 花名册名单排序

# 花名册文件：(班级名, 文件路径, 解析方式)
# 解析方式按文件结构选择：xls_107 / xls_108 / xlsx_109（新文件可按需增加解析函数）
ROSTER_FILES = [
    ('初三107班', '/sandbox/workspace/uploads/107班花名册全(1).xls', 'xls_107'),
    ('初二108班', '/sandbox/workspace/uploads/2026年下学期108班花名册全 （最新版）(2).xls', 'xls_108'),
    ('初一109班', '/sandbox/workspace/uploads/109班初一花名册.xlsx', 'xlsx_109'),
]
# 108班：以"全校所有学生名单"为准，个别缺性别者在此补充
GENDER_FALLBACK = {'张伟': '男', '李娜': '女'}


# ==================== 公文格式 helper ====================
def new_doc():
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(3.7)
    sec.bottom_margin = Cm(3.5)
    sec.left_margin = Cm(2.8)
    sec.right_margin = Cm(2.6)
    return doc


def set_font(run, name_east, name_ascii, size, bold=False):
    run.font.name = name_ascii
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0)
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), name_ascii)
    rFonts.set(qn('w:hAnsi'), name_ascii)
    rFonts.set(qn('w:eastAsia'), name_east)


def add_para(doc, text, font_east='仿宋_GB2312', font_ascii='Times New Roman',
             size=16, bold=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
             indent_first=True, line=28, space_before=0, space_after=0):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(line)
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    if indent_first:
        pf.first_line_indent = Pt(size * 2)
    run = p.add_run(text)
    set_font(run, font_east, font_ascii, size, bold)
    return p


def add_title(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(33)
    pf.space_after = Pt(16)
    run = p.add_run(text)
    set_font(run, '方正小标宋简体', 'Times New Roman', 22, bold=False)


def add_sign(doc):
    add_para(doc, '', indent_first=False)
    add_para(doc, SCHOOL, align=WD_ALIGN_PARAGRAPH.RIGHT, indent_first=False)
    add_para(doc, DATE, align=WD_ALIGN_PARAGRAPH.RIGHT, indent_first=False)


def h1(doc, t):
    add_para(doc, t, font_east='黑体')


def _clean(v):
    return str(v).strip().replace(' ', '').replace('\n', '') if v is not None else ''


def _valid_name(nm):
    return nm and 1 < len(nm) <= 4 and nm != '姓名' and not any(ch.isdigit() for ch in nm)


# ==================== 花名册解析（按文件结构选用，新文件可扩展） ====================
def parse_xls_107(path, cls):
    """107班花名册：Sheet1，姓名=col0，性别=col1，数据从第5行起"""
    out = []
    wb = xlrd.open_workbook(path)
    ws = wb.sheet_by_index(0)
    for i in range(4, ws.nrows):
        nm = _clean(ws.cell_value(i, 0))
        gd = _clean(ws.cell_value(i, 1))
        if _valid_name(nm):
            out.append({'name': nm, 'gender': gd, 'cls': cls})
    return out


def parse_xls_108(path, cls):
    """108班花名册：以"全校所有学生名单"为准，性别从"Sheet1 信息采集表"补充"""
    out = []
    wb = xlrd.open_workbook(path)
    gmap = {}
    ws1 = wb.sheet_by_name('Sheet1')
    for i in range(4, ws1.nrows):
        nm = _clean(ws1.cell_value(i, 3))
        gd = _clean(ws1.cell_value(i, 9))
        if _valid_name(nm):
            gmap[nm] = gd
    wsall = wb.sheet_by_name('全校所有学生名单')
    for i in range(1, wsall.nrows):
        if str(wsall.cell_value(i, 1)).strip() in ('108', '108.0'):
            nm = _clean(wsall.cell_value(i, 2))
            if nm:
                out.append({'name': nm, 'gender': gmap.get(nm, GENDER_FALLBACK.get(nm, '')), 'cls': cls})
    return out


def parse_xlsx_109(path, cls):
    """109班花名册：Sheet1，姓名=col0，性别=col1，数据从第3行起"""
    out = []
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['Sheet1']
    for row in ws.iter_rows(min_row=2, values_only=True):
        nm = _clean(row[0])
        gd = _clean(row[1]) if len(row) > 1 else ''
        if _valid_name(nm):
            out.append({'name': nm, 'gender': gd, 'cls': cls})
    return out


PARSERS = {'xls_107': parse_xls_107, 'xls_108': parse_xls_108, 'xlsx_109': parse_xlsx_109}


def extract_students():
    students = []
    for cls, path, how in ROSTER_FILES:
        students.extend(PARSERS[how](path, cls))
    return students


# ==================== 学生→社团 分配 ====================
def assign(students, seed=2026):
    """每人随机参加 2—MAX_CLUBS 个社团，优先分配给人数少的社团，使各社团人数均衡"""
    random.seed(seed)
    asg = {c: [] for c in CLUB_NAMES}
    studs = students[:]
    random.shuffle(studs)
    for s in studs:
        k = random.randint(2, MAX_CLUBS)
        order = sorted(CLUB_NAMES, key=lambda c: (len(asg[c]), random.random()))
        for c in order[:k]:
            asg[c].append(s)
    return asg


def sort_members(members):
    return sorted(members, key=lambda s: (CLS_ORDER.get(s['cls'], 9), s['name']))


# ==================== 生成文档 ====================
def gen_plan(club, typ, place, teacher, members):
    doc = new_doc()
    add_title(doc, f'{SCHOOL}{club}社团活动计划')
    add_para(doc, f'为规范{club}社团活动开展，落实课后服务要求，发展学生兴趣特长，促进学生全面发展，结合我校实际，特制定本社团活动计划。')
    h1(doc, '一、活动目标')
    add_para(doc, f'1.丰富学生课后服务生活，发展{club}方面的兴趣特长；')
    add_para(doc, '2.培养学生良好的活动习惯和团队协作精神；')
    add_para(doc, '3.提升学生的综合素质，促进学生全面发展。')
    h1(doc, '二、活动时间与地点')
    add_para(doc, f'活动时间：{ACTIVITY_TIME}', indent_first=False)
    add_para(doc, f'活动地点：{place}', indent_first=False)
    add_para(doc, f'指导老师：{teacher}', indent_first=False)
    h1(doc, '三、社团成员')
    add_para(doc, f'本社团现有成员{len(members)}名，具体名单见《{club}社团学生名单》。')
    h1(doc, '四、活动内容与安排')
    rows = [
        ('9月', f'{club}社团组建，明确活动纪律与要求，开展基础训练'),
        ('10月', '开展基础技能训练，夯实基本功，分组开展活动'),
        ('11月', '开展技能提升训练，组织小组练习与展示'),
        ('12月', '组织社团成果展示和交流活动，以展示检验成效'),
        ('1月', '进行学期总结与评价考核，做好活动资料归档'),
    ]
    table = doc.add_table(rows=len(rows) + 1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'
    for j, t in enumerate(['时间', '活动内容']):
        cell = table.rows[0].cells[j]
        cell.width = Cm(3.5) if j == 0 else Cm(11.5)
        cell.text = ''
        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_font(para.add_run(t), '黑体', 'Times New Roman', 14)
    for i, (m, c) in enumerate(rows, start=1):
        c0 = table.rows[i].cells[0]
        c0.width = Cm(3.5)
        c0.text = ''
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_font(p0.add_run(m), '仿宋_GB2312', 'Times New Roman', 14)
        c1 = table.rows[i].cells[1]
        c1.width = Cm(11.5)
        c1.text = ''
        set_font(c1.paragraphs[0].add_run(c), '仿宋_GB2312', 'Times New Roman', 14)
    h1(doc, '五、活动要求')
    for x in ['按时参加活动，不迟到、不早退；', '遵守活动纪律，注意活动安全；', '积极参与，认真训练，爱护活动器材；', '讲究卫生，保持活动场地整洁。']:
        add_para(doc, x)
    add_sign(doc)
    path = os.path.join(OUT, typ, f'{club}社团活动计划.docx')
    doc.save(path)
    return path


def gen_roster(club, typ, place, teacher, members):
    doc = new_doc()
    add_title(doc, f'{SCHOOL}{club}社团学生名单')
    add_para(doc, f'活动时间：{ACTIVITY_TIME}', indent_first=False)
    add_para(doc, f'指导老师：{teacher}', indent_first=False)
    add_para(doc, f'活动地点：{place}', indent_first=False)
    ms = sort_members(members)
    table = doc.add_table(rows=len(ms) + 1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'
    for j, t in enumerate(['序号', '姓名', '性别', '班级']):
        cell = table.rows[0].cells[j]
        cell.text = ''
        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_font(para.add_run(t), '黑体', 'Times New Roman', 14)
    for i, s in enumerate(ms, start=1):
        for j, v in enumerate([str(i), s['name'], s['gender'], s['cls']]):
            cell = table.rows[i].cells[j]
            cell.text = ''
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_font(para.add_run(v), '仿宋_GB2312', 'Times New Roman', 14)
    add_para(doc, f'本社团共有成员{len(ms)}名。', indent_first=False)
    add_sign(doc)
    path = os.path.join(OUT, typ, f'{club}社团学生名单.docx')
    doc.save(path)
    return path


if __name__ == '__main__':
    students = extract_students()
    from collections import Counter
    print('提取学生总数:', len(students), dict(Counter(s['cls'] for s in students)))
    for typ in ['体育类社团', '美育类社团', '特色类社团']:
        os.makedirs(os.path.join(OUT, typ), exist_ok=True)
    asg = assign(students)
    for club, typ, place, teacher in CLUBS:
        gen_plan(club, typ, place, teacher, asg[club])
        gen_roster(club, typ, place, teacher, asg[club])
        print(f'{club}: {len(asg[club])}人 -> 活动计划 + 学生名单')
    # 校验每人社团数 ≤ MAX_CLUBS
    cnt = {}
    for c, ms in asg.items():
        for s in ms:
            key = s['name'] + s['cls']
            cnt[key] = cnt.get(key, 0) + 1
    print('单个学生最多参加社团数:', max(cnt.values()) if cnt else 0, f'(应≤{MAX_CLUBS})')
    print('生成完成，共', len(CLUBS) * 2, '份文档')
