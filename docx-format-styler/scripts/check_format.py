#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""抽查 docx 段落的实际格式（字体 / 字号 / 加粗 / 缩进），直接读底层 XML。

用法：python3 inspect.py 文档.docx
"""
import sys
import zipfile
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def q(tag):
    return f"{{{W}}}{tag}"


def run_fmt(r):
    rpr = r.find(q("rPr"))
    d = {}
    if rpr is None:
        return d
    rf = rpr.find(q("rFonts"))
    if rf is not None:
        d["eastAsia"] = rf.get(q("eastAsia"))
        d["ascii"] = rf.get(q("ascii"))
    sz = rpr.find(q("sz"))
    if sz is not None:
        d["sz_pt"] = int(sz.get(q("val"))) / 2
    b = rpr.find(q("b"))
    if b is not None:
        d["bold"] = b.get(q("val")) not in ("0", "false", "False")
    return d


def para_fmt(p):
    ppr = p.find(q("pPr"))
    d = {}
    if ppr is None:
        return d
    jc = ppr.find(q("jc"))
    if jc is not None:
        d["jc"] = jc.get(q("val"))
    ind = ppr.find(q("ind"))
    if ind is not None:
        d["ind"] = {k.split("}")[-1]: v for k, v in ind.attrib.items()}
    return d


def main():
    path = sys.argv[1]
    root = etree.fromstring(zipfile.ZipFile(path).read("word/document.xml"))
    body = root.find(q("body"))
    paras = body.findall(q("p"))
    print(f"段落数: {len(paras)}\n")
    for i, p in enumerate(paras):
        full = "".join(t.text or "" for t in p.findall(".//" + q("t")))
        print(f"--- P{i} {para_fmt(p)}")
        print(f"    TEXT: {full[:60]!r}")
        for j, r in enumerate(p.findall(q("r"))):
            t = "".join(x.text or "" for x in r.findall(q("t")))
            print(f"    R{j}: {t[:30]!r} -> {run_fmt(r)}")
        print()


if __name__ == "__main__":
    main()
