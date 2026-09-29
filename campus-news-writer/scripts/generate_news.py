# -*- coding: utf-8 -*-
"""
校园新闻稿 Word 生成脚本（格式对齐既有稿件规范）。

用法：
    python3 generate_news.py --config /path/to/news.json

JSON 配置结构（image、reporter 可省略）：
{
  "title": "金秋启新程，逐梦正当时",          # 主标题文字，脚本自动加中文双引号
  "subtitle": "蓬源中学2026年秋季开学典礼",    # 副标题文字，脚本自动加"——"前缀
  "paragraphs": ["第一段", "第二段", "第三段"], # 正文，通常三段
  "image": "/path/to/photo.jpg",               # 可选，现场照片
  "reporter": "通讯员：刘东安",                 # 可选，默认"通讯员：刘东安"
  "output": "/path/to/out.docx"                # 输出文件绝对路径
}

格式规范（自动应用）：
- 主标题：宋体 16pt，居中，中文双引号包裹
- 副标题：宋体 16pt，右对齐，"——"前缀
- 正文：宋体 14pt，左对齐
- 配图：居中，宽度 12cm，正文之后落款之前
- 落款：宋体 14pt，右对齐
"""
import json
import sys
import argparse
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn


def set_run_font(run, size_pt, name='宋体'):
    run.font.name = name
    run.font.size = Pt(size_pt)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn('w:eastAsia'), name)


def add_para(doc, text, size_pt, align):
    p = doc.add_paragraph()
    p.alignment = align
    run = p.add_run(text)
    set_run_font(run, size_pt)
    return p


def add_picture(doc, image_path, width_cm=12):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(image_path, width=Cm(width_cm))
    return p


def generate(cfg):
    doc = Document()

    title = cfg['title']
    if not (title.startswith('\u201c') and title.endswith('\u201d')):
        title = '\u201c' + title + '\u201d'

    subtitle = cfg['subtitle']
    if not subtitle.startswith('\u2014'):
        subtitle = '\u2014' + subtitle

    add_para(doc, title, 16, WD_ALIGN_PARAGRAPH.CENTER)
    add_para(doc, subtitle, 16, WD_ALIGN_PARAGRAPH.RIGHT)

    for ptext in cfg.get('paragraphs', []):
        add_para(doc, ptext, 14, WD_ALIGN_PARAGRAPH.LEFT)

    image_path = cfg.get('image')
    if image_path:
        add_picture(doc, image_path, 12)

    reporter = cfg.get('reporter', '通讯员：刘东安')
    add_para(doc, reporter, 14, WD_ALIGN_PARAGRAPH.RIGHT)

    doc.save(cfg['output'])
    print('已生成:', cfg['output'])


def main():
    parser = argparse.ArgumentParser(description='生成校园新闻稿 Word 文档')
    parser.add_argument('--config', required=True, help='JSON 配置文件路径')
    args = parser.parse_args()

    with open(args.config, 'r', encoding='utf-8') as f:
        cfg = json.load(f)

    generate(cfg)


if __name__ == '__main__':
    main()
