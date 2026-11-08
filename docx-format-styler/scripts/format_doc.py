#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
docx 格式套用脚本
=================
分析"参考文档"里的标题格式与正文格式（字体、字号、加粗、首行缩进），
然后把"目标文档"改成与参考文档一致的字体字号，并保留目标文档自身的
文字内容、段落结构、对齐方式与图片。

用法：
    python3 format_doc.py --ref 参考.docx --target 目标.docx --out 输出.docx

可选：
    --title-pt 20      手动指定标题字号（磅），覆盖自动识别
    --body-pt 16       手动指定正文字号（磅），覆盖自动识别
    --font 宋体        手动指定字体，覆盖自动识别
"""
import argparse
import os
import sys
import zipfile
import binascii
from collections import Counter


def patch_zip_crc():
    """绕过 ZIP 的 CRC 校验。

    很多由微信 / WPS / 编辑器导出的 docx，内嵌图片的 CRC 记录损坏，
    会导致 python-docx 读取时报 BadZipFile 而整体打不开。
    这里把 CRC 校验方法替换为空操作，使文件仍可正常读取；
    保存时 python-docx 会重新计算 CRC，输出文件自然是好的。
    """
    def _patched(self, newdata):
        if self._expected_crc is None:
            return
        self._running_crc = binascii.crc32(newdata, self._running_crc)
    zipfile.ZipExtFile._update_crc = _patched


patch_zip_crc()

from docx import Document  # noqa: E402
from docx.shared import Pt  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402

CENTER = WD_ALIGN_PARAGRAPH.CENTER
RIGHT = WD_ALIGN_PARAGRAPH.RIGHT
TITLE_ALIGNS = (CENTER, RIGHT)


# ----------------------------------------------------------------------
# 读取
# ----------------------------------------------------------------------
def get_doc_defaults(doc):
    """读取文档默认字体（docDefaults），作为 run 无显式格式时的回退值。"""
    d = {"font": "宋体", "size": 10.5}
    try:
        styles_el = doc.styles.element
        dd = styles_el.find(qn("w:docDefaults"))
        if dd is not None:
            rpr = dd.find(qn("w:rPrDefault"))
            if rpr is not None:
                rpr = rpr.find(qn("w:rPr"))
            if rpr is not None:
                rf = rpr.find(qn("w:rFonts"))
                if rf is not None:
                    d["font"] = rf.get(qn("w:eastAsia")) or rf.get(qn("w:ascii")) or d["font"]
                sz = rpr.find(qn("w:sz"))
                if sz is not None:
                    d["size"] = int(sz.get(qn("w:val"))) / 2
    except Exception:
        pass
    return d


def run_fmt(run, ddef):
    """取某个 run 的实际格式（字体 / 字号 / 加粗），缺失时回退到文档默认值。"""
    rpr = run._element.find(qn("w:rPr"))
    name = None
    bold = None
    if rpr is not None:
        rf = rpr.find(qn("w:rFonts"))
        if rf is not None:
            name = rf.get(qn("w:eastAsia")) or rf.get(qn("w:ascii"))
        b = rpr.find(qn("w:b"))
        if b is not None:
            # 注意 w:val 可能是 "false" 而非 "0"
            bold = b.get(qn("w:val")) not in ("0", "false", "False")
    if not name:
        name = ddef["font"]
    size = run.font.size.pt if run.font.size else None
    if size is None:
        size = ddef["size"]
    if bold is None:
        bold = False
    return {"font": name, "size": size, "bold": bold}


def para_indent(p):
    """读取段落首行缩进，返回 (firstLineChars, firstLine) 或 None。"""
    ppr = p._p.find(qn("w:pPr"))
    if ppr is None:
        return None
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        return None
    chars = ind.get(qn("w:firstLineChars"))
    line = ind.get(qn("w:firstLine"))
    if chars is None and line is None:
        return None
    return (int(chars) if chars else None, int(line) if line else None)


def is_title_like(p):
    return p.alignment in TITLE_ALIGNS


# ----------------------------------------------------------------------
# 分析参考文档
# ----------------------------------------------------------------------
def analyze(doc):
    """识别参考文档的【标题格式】与【正文格式】。

    规则：
      - 正文格式 = 正文段落中出现次数最多的字号（众数），最稳。
      - 标题格式 = 字号最大者；若最大的不唯一，优先取居中/右对齐的段落。
      - 正文首行缩进 = 正文段落中缩进的众数。
    """
    ddef = get_doc_defaults(doc)
    para_formats = []  # 每段：{'fmt': 最大字号run格式, 'title_like': bool, 'indent': ...}
    for p in doc.paragraphs:
        if not p.text.strip():
            continue
        fmts = [run_fmt(r, ddef) for r in p.runs if (r.text or "").strip() != ""]
        if not fmts:
            continue
        maxf = max(fmts, key=lambda f: f["size"])
        fmt = maxf
        # 标题里若有多个 run（如标题主体 + 副标题），加粗沿用标题段中是否加粗
        para_formats.append({
            "fmt": fmt,
            "title_like": is_title_like(p),
            "indent": para_indent(p),
        })

    if not para_formats:
        return None, None, None

    max_size = max(x["fmt"]["size"] for x in para_formats)
    biggest = [x for x in para_formats if x["fmt"]["size"] == max_size]
    # 标题：优先居中/右对齐者
    title_fmt = next((x["fmt"] for x in biggest if x["title_like"]), biggest[0]["fmt"])

    body = [x for x in para_formats if x["fmt"]["size"] != max_size] or para_formats
    size_counter = Counter(x["fmt"]["size"] for x in body)
    common_size = size_counter.most_common(1)[0][0]
    body_fmt = next(x["fmt"] for x in body if x["fmt"]["size"] == common_size)

    indents = [x["indent"] for x in body if x["indent"]]
    indent = Counter(indents).most_common(1)[0][0] if indents else None

    return title_fmt, body_fmt, indent


# ----------------------------------------------------------------------
# 写入目标文档
# ----------------------------------------------------------------------
def set_run(run, font, size_pt, bold):
    run.font.name = font
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    for attr in ("w:eastAsia", "w:cs", "w:ascii", "w:hAnsi"):
        rfonts.set(qn(attr), font)
    run.font.size = Pt(size_pt)
    # szCs（复杂脚本字号），保证中文彻底生效
    for e in rpr.findall(qn("w:szCs")):
        rpr.remove(e)
    szcs = OxmlElement("w:szCs")
    szcs.set(qn("w:val"), str(int(size_pt * 2)))
    sz = rpr.find(qn("w:sz"))
    if sz is not None:
        sz.addnext(szcs)
    else:
        rpr.append(szcs)
    run.font.bold = bold


def set_indent(p, indent):
    if not indent:
        return
    chars, line = indent
    ppr = p._p.get_or_add_pPr()
    ind = ppr.get_or_add_ind()
    if chars:
        ind.set(qn("w:firstLineChars"), str(chars))
    if line:
        ind.set(qn("w:firstLine"), str(line))


def apply(doc, title_fmt, body_fmt, indent):
    n_title = n_body = 0
    for p in doc.paragraphs:
        if is_title_like(p):
            for r in p.runs:
                set_run(r, title_fmt["font"], title_fmt["size"], title_fmt["bold"])
            if p.runs:
                n_title += 1
        else:
            set_indent(p, indent)
            for r in p.runs:
                set_run(r, body_fmt["font"], body_fmt["size"], body_fmt["bold"])
            if p.runs:
                n_body += 1
    return n_title, n_body


# ----------------------------------------------------------------------
# 主流程
# ----------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="把一个文档的标题/正文字体字号套用到另一个文档")
    ap.add_argument("--ref", required=True, help="参考文档路径（提供格式模板）")
    ap.add_argument("--target", required=True, help="目标文档路径（被修改）")
    ap.add_argument("--out", required=True, help="输出文档路径")
    ap.add_argument("--title-pt", type=float, help="手动指定标题字号（磅）")
    ap.add_argument("--body-pt", type=float, help="手动指定正文字号（磅）")
    ap.add_argument("--font", help="手动指定字体名")
    args = ap.parse_args()

    ref = Document(args.ref)
    target = Document(args.target)

    title_fmt, body_fmt, indent = analyze(ref)
    if title_fmt is None or body_fmt is None:
        print("!! 参考文档里没有可识别的文字段落，无法提取格式", file=sys.stderr)
        sys.exit(1)

    # 手动覆盖
    if args.title_pt:
        title_fmt["size"] = args.title_pt
    if args.body_pt:
        body_fmt["size"] = args.body_pt
    if args.font:
        title_fmt["font"] = args.font
        body_fmt["font"] = args.font

    print("参考文档识别结果：")
    print(f"  标题格式：字体={title_fmt['font']} 字号={title_fmt['size']}磅 "
          f"加粗={title_fmt['bold']}")
    print(f"  正文格式：字体={body_fmt['font']} 字号={body_fmt['size']}磅 "
          f"加粗={body_fmt['bold']} 首行缩进={indent}")

    n_title, n_body = apply(target, title_fmt, body_fmt, indent)

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    target.save(args.out)
    print(f"\n已套用：标题段 {n_title} 个，正文段 {n_body} 个")
    print(f"已保存：{args.out}")


if __name__ == "__main__":
    main()
