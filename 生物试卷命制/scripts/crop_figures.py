# -*- coding: utf-8 -*-
"""从扫描版教辅/教材 PDF 中提取插图（供生物试卷配图使用）。

两步用法：
  1) 渲染：python crop_figures.py render <pdf> <outdir> <start_page> <end_page> [dpi]
     —— 把 PDF 指定页范围渲染为 PNG（1-based 页码，便于与书本页码对照）。
  2) 裁剪 + 拼校验图：
     python crop_figures.py crop <pagesdir> <outdir> <coords.json>
     —— coords.json 形如 {"图名": {"page": "p003", "box": [左,上,右,下]}, ...}
        box 为相对坐标(0~1，页面左上角=(0,0))；脚本会加少量 padding 并放大 2 倍输出，
        同时生成 overview.png（网格缩略图）供视觉复核。

配套建议：坐标不宜靠目测，先用连通域分析（cv2/scipy）区分“文字行（小连通域）”与
“图形（大连通域/长线条）”，得到图形主体的紧贴边界，再用视觉模型复核。
"""
import os, sys, json
from PIL import Image, ImageDraw, ImageFont


def render(pdf, outdir, start, end, dpi=170):
    import pymupdf
    d = pymupdf.open(pdf)
    os.makedirs(outdir, exist_ok=True)
    for i in range(start - 1, min(end, d.page_count)):
        pix = d[i].get_pixmap(dpi=dpi)
        pix.save(os.path.join(outdir, f"p{i+1:03d}.png"))
    print("rendered", start, end, "to", outdir)


def crop(pagesdir, outdir, coordfile, pad=0.006, scale_to=900):
    os.makedirs(outdir, exist_ok=True)
    coords = json.load(open(coordfile, encoding="utf-8"))
    cache = {}
    for name, spec in coords.items():
        pg, box = spec["page"], spec["box"]
        if pg not in cache:
            cache[pg] = Image.open(os.path.join(pagesdir, pg + ".png"))
        im = cache[pg]; w, h = im.size
        x0, y0, x1, y1 = box
        x0 = max(0, x0 - pad); y0 = max(0, y0 - pad); x1 = min(1, x1 + pad); y1 = min(1, y1 + pad)
        c = im.crop((int(x0 * w), int(y0 * h), int(x1 * w), int(y1 * h)))
        sc = max(1, int(scale_to / max(1, c.width)))
        if sc > 1:
            c = c.resize((c.width * sc, c.height * sc), Image.LANCZOS)
        c.save(os.path.join(outdir, name + ".png"))
    print("cropped", len(coords))
    montage(outdir)


def montage(outdir, cols=5, cell=300, lbl=26):
    files = sorted(f for f in os.listdir(outdir) if f.endswith(".png") and f != "overview.png")
    if not files:
        return
    rows = (len(files) + cols - 1) // cols
    W = cols * (cell + 8) + 8; H = rows * (cell + lbl + 8) + 8
    S = Image.new("RGB", (W, H), "white"); dr = ImageDraw.Draw(S)
    try:
        font = ImageFont.truetype("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", 18)
    except Exception:
        font = ImageFont.load_default()
    for i, f in enumerate(files):
        im = Image.open(os.path.join(outdir, f)).convert("RGB"); im.thumbnail((cell, cell))
        r, c = divmod(i, cols)
        x = 8 + c * (cell + 8); y = 8 + r * (cell + lbl + 8)
        dr.text((x, y), f, fill="red", font=font)
        S.paste(im, (x + (cell - im.width) // 2, y + lbl))
    S.save(os.path.join(outdir, "overview.png"))
    print("overview ->", os.path.join(outdir, "overview.png"))


if __name__ == "__main__":
    cmd = sys.argv[1]
    if cmd == "render":
        render(sys.argv[2], sys.argv[3], int(sys.argv[4]), int(sys.argv[5]),
               int(sys.argv[6]) if len(sys.argv) > 6 else 170)
    elif cmd == "crop":
        crop(sys.argv[2], sys.argv[3], sys.argv[4])
    else:
        print(__doc__)
