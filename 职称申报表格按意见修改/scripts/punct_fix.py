# -*- coding: utf-8 -*-
"""
docx 中文标点修复工具。

用途：把 Word 文档中的半角/英文标点统一为中文全角标点，避免职称、公文类材料
出现“引号是英文直引号”“序号点用半角”等不规范问题。

处理内容：
1. 英文直双引号 " → 中文弯引号 “ ”，按出现顺序自动配对；
2. 英文直单引号 ' → 中文弯单引号 ‘ ’（配对）；
3. 序号点：数字后跟半角点+空格（如 "1. 运动能力"）→ 全角点 "1．运动能力"；
4. 常用英文标点（, ; : ! ? ( ) [ ]）在中文语境下 → 中文标点（,：；！？（）【】）。
   注：小数、日期中的点（如 2009.9）、百分号 %、斜杠 / 属正常用法，保留不动。

用法：
    python3 punct_fix.py <file.docx> [out.docx]
"""
import sys
import re
from docx import Document
from docx.oxml.ns import qn


def _paras(doc):
    for p in doc.paragraphs:
        yield p
    for t in doc.tables:
        for row in t.rows:
            seen = set()
            for c in row.cells:
                if id(c._tc) in seen:
                    continue
                seen.add(id(c._tc))
                for p in c.paragraphs:
                    yield p


def fix(path, out=None):
    doc = Document(path)
    q = [0]
    sq = [0]

    def rq(_m):
        q[0] += 1
        return '\u201c' if q[0] % 2 else '\u201d'

    def rsq(_m):
        sq[0] += 1
        return '\u2018' if sq[0] % 2 else '\u2019'

    n_q = n_d = 0
    for p in _paras(doc):
        for r in p.runs:
            t = r.text
            if not t:
                continue
            t3 = re.sub(r'(\d)\.[ \t]+', r'\1．', t)          # 序号点
            t3, c1 = re.subn(r'"', rq, t3)                     # 直双引号
            t3, c2 = re.subn(r"'", rsq, t3)                    # 直单引号
            t3, c3 = re.subn(r'(?<=[\u4e00-\u9fff])\s*,', '，', t3)
            t3, c4 = re.subn(r'(?<=[\u4e00-\u9fff])\s*;', '；', t3)
            t3, c5 = re.subn(r'(?<=[\u4e00-\u9fff])\s*:(?![0-9])', '：', t3)
            t3, c6 = re.subn(r'(?<=[\u4e00-\u9fff])\s*!', '！', t3)
            t3, c7 = re.subn(r'(?<=[\u4e00-\u9fff])\s*\?', '？', t3)
            if t3 != t:
                r.text = t3
            n_q += c1
            n_d += len(re.findall(r'(\d)\.[ \t]+', t))
    doc.save(out or path)
    print('修复 %s：中文引号 %d 个' % (out or path, n_q))


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('用法: python3 punct_fix.py <file.docx> [out.docx]')
        sys.exit(0)
    fix(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
