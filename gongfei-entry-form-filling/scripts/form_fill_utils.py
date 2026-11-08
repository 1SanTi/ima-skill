# -*- coding: utf-8 -*-
"""入编表格填写通用工具：保格式写值、按网格定位合并单元格、等比适配插图。

设计要点：
- docx 写值：不改单元格/行的既有属性，仅清空并写入 run，字体统一“方正仿宋简体”+Times New Roman。
- 定位：Word 表格有大量合并单元格，用“网格列号”定位最稳，不要用 row.cells[i]（会因合并错位）。
- 插图：按单元格内径等比缩放，避免撑大行高/改变表格尺寸。
"""
import copy
from docx.shared import Pt, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.table import _Cell

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


# ---------- docx ----------
def set_font(run, size=12, ea='方正仿宋简体'):
    run.font.size = Pt(size)
    run.font.name = 'Times New Roman'
    rpr = run._element.get_or_add_rPr()
    rf = rpr.get_or_add_rFonts()
    rf.set(qn('w:eastAsia'), ea)
    rf.set(qn('w:ascii'), 'Times New Roman')
    rf.set(qn('w:hAnsi'), 'Times New Roman')


def set_docx_cell(cell, text, align='center', size=12):
    """清空单元格并写入 text（支持 \n 换行），保留单元格既有格式。"""
    for p in cell.paragraphs[1:]:
        p._element.getparent().remove(p._element)
    p = cell.paragraphs[0]
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    p.alignment = {'center': WD_ALIGN_PARAGRAPH.CENTER,
                   'left': WD_ALIGN_PARAGRAPH.LEFT,
                   'right': WD_ALIGN_PARAGRAPH.RIGHT}[align]
    first = True
    for line in str(text).split('\n'):
        if not first:
            p.add_run().add_break()
        set_font(p.add_run(line), size)
        first = False
    return cell


def grid_cell(row, gridcol):
    """按网格列号返回 python-docx 单元格（正确处理 gridSpan 合并）。"""
    pos = 0
    for tc in row._tr.tc_lst:
        span = int(tc.tcPr.gridSpan.val) if (tc.tcPr is not None and tc.tcPr.gridSpan is not None) else 1
        if pos <= gridcol < pos + span:
            return _Cell(tc, row.table)
        pos += span
    raise IndexError(f'gridcol {gridcol} out of range')


def cell_size(row, gridcol):
    """返回 (宽cm, 高cm)：按网格列合计列宽、按该单元格实际跨行合计行高估算。"""
    tbl = row.table
    cols = [int(gc.get(W + 'w')) for gc in tbl._tbl.tblGrid.findall(W + 'gridCol')]
    # 列宽（twips -> cm）
    pos = 0
    cw = 0
    span = 1
    for tc in row._tr.tc_lst:
        s = int(tc.tcPr.gridSpan.val) if (tc.tcPr is not None and tc.tcPr.gridSpan is not None) else 1
        if pos <= gridcol < pos + s:
            span = s
            break
        pos += s
    cw = sum(cols[pos:pos + span]) * 2.54 / 1440
    # 行高（trHeight.val 由 python-docx 转成 EMU），按 vMerge 向下累计
    ridx = list(tbl.rows).index(row)
    tot = 0
    for r in tbl.rows[ridx:]:
        trpr = r._tr.trPr
        h = int(trpr.trHeight.val) if (trpr is not None and trpr.trHeight is not None) else None
        tot += (h or Emu(Cm(0.8)))
        # 判断该单元格是否继续向下合并（vMerge continue）
        c = grid_cell(r, gridcol)
        tcpr = c._tc.tcPr
        vm = tcpr.find(qn('w:vMerge')) if tcpr is not None else None
        cont = vm is not None and (vm.get(qn('w:val')) in (None, 'continue'))
        if not cont:
            break
    return cw, tot / 360000.0  # EMU -> cm


def insert_photo_fit(cell, photo_path, max_w_cm, max_h_cm, margin_cm=0.15):
    """等比缩放照片放入单元格，居中。max_w_cm/max_h_cm 为照片栏内径（可先用 cell_size 估算再减边距）。"""
    from PIL import Image
    iw, ih = Image.open(photo_path).size
    ratio = ih / iw
    W0, H0 = max_w_cm, max_h_cm
    w = W0 - margin_cm
    h = w * ratio
    if h > H0 - margin_cm:
        h = H0 - margin_cm
        w = h / ratio
    for p in cell.paragraphs[1:]:
        p._element.getparent().remove(p._element)
    p = cell.paragraphs[0]
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(photo_path, width=Cm(w), height=Cm(h))
    return w, h


# ---------- xlsx ----------
def xlsx_insert_photo(ws, cell, photo_path, max_w_cm=3.2, max_h_cm=4.0):
    """按 cm 上限等比缩放，锚定到 cell（左上角）。"""
    from openpyxl.drawing.image import Image as XLImage
    from PIL import Image
    iw, ih = Image.open(photo_path).size
    ratio = ih / iw
    w = max_w_cm
    h = w * ratio
    if h > max_h_cm:
        h = max_h_cm
        w = h / ratio
    px = lambda cm: int(cm / 2.54 * 96)
    img = XLImage(photo_path)
    img.width, img.height = px(w), px(h)
    ws.add_image(img, cell)


def load_docx_tolerant(path):
    """读取微信/WPS 导出的、CRC 损坏的 docx。"""
    import zipfile
    from docx import Document
    orig = zipfile.ZipExtFile._update_crc
    zipfile.ZipExtFile._update_crc = lambda self, data: None
    try:
        return Document(path)
    finally:
        zipfile.ZipExtFile._update_crc = orig
