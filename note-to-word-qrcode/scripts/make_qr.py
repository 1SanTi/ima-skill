# -*- coding: utf-8 -*-
"""
链接/文本 → 二维码 PNG（默认 H 级纠错 + 中文标题 + 回读校验）
用法:
  python make_qr.py --url "https://..." --out qr.png --title "..." --subtitle "..."
  python make_qr.py --file urls.txt --outdir qrdir
"""
import argparse, os, sys

def ensure_qrcode():
    try:
        import qrcode  # noqa
    except ImportError:
        os.system(f'{sys.executable} -m pip install -q "qrcode[pil]"')

def find_font():
    cands = [
        '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
        '/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc',
        '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc',
        '/usr/share/fonts/truetype/arphic/uming.ttc',
        'C:/Windows/Fonts/msyh.ttc', 'C:/Windows/Fonts/simhei.ttf',
        '/System/Library/Fonts/PingFang.ttc',
    ]
    for c in cands:
        if os.path.exists(c):
            return c
    return None

def make_one(url, out, title=None, subtitle=None, box_size=12, ec='H',
             fg='black', bg='white', logo=None, scale=1.0):
    ensure_qrcode()
    import qrcode
    from qrcode.constants import (ERROR_CORRECT_L, ERROR_CORRECT_M,
                                  ERROR_CORRECT_Q, ERROR_CORRECT_H)
    from PIL import Image, ImageDraw, ImageFont

    ecl = {'L': ERROR_CORRECT_L, 'M': ERROR_CORRECT_M, 'Q': ERROR_CORRECT_Q, 'H': ERROR_CORRECT_H}[ec]
    qr = qrcode.QRCode(error_correction=ecl, box_size=box_size, border=4)
    qr.add_data(url); qr.make(fit=True)
    img = qr.make_image(fill_color=fg, back_color=bg).convert('RGB')

    if logo and os.path.exists(logo):
        lg = Image.open(logo).convert('RGBA')
        w, h = img.size
        size = int(w * 0.22)
        lg = lg.resize((size, size))
        img.paste(lg, ((w - size)//2, (h - size)//2), lg)

    font = find_font()
    pad = 40
    fpath = font
    def load(sz):
        try:
            return ImageFont.truetype(fpath, sz) if fpath else ImageFont.load_default()
        except Exception:
            return ImageFont.load_default()

    title_font = load(int(box_size * 2.2))
    sub_font = load(int(box_size * 1.3))
    top = pad + (title_font.size + 20 if title else 0)
    bot = pad + (sub_font.size + 12 if subtitle else 0)
    canvas = Image.new('RGB', (img.size[0] + pad*2, img.size[1] + top + bot), 'white')
    canvas.paste(img, (pad, top))
    d = ImageDraw.Draw(canvas)
    if title:
        tw = d.textlength(title, font=title_font)
        d.text(((canvas.size[0]-tw)/2, pad), title, fill='black', font=title_font)
    if subtitle:
        sw = d.textlength(subtitle, font=sub_font)
        d.text(((canvas.size[0]-sw)/2, img.size[1] + top + 10), subtitle, fill='#666666', font=sub_font)

    if scale != 1.0:
        canvas = canvas.resize((int(canvas.size[0]*scale), int(canvas.size[1]*scale)))
    canvas.save(out, dpi=(300, 300))

    # 回读校验（对纯二维码区域解码，避免标题文字干扰）
    ok = None
    try:
        import cv2, numpy as np
        det = cv2.QRCodeDetector()
        data, _, _ = det.detectAndDecode(cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR))
        ok = (data == url)
    except Exception:
        pass
    print(f'OK {out} verify={ok} url={url[:70]}...')
    return out

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--url'); ap.add_argument('--file')
    ap.add_argument('--out'); ap.add_argument('--outdir', default='.')
    ap.add_argument('--title'); ap.add_argument('--subtitle')
    ap.add_argument('--box-size', type=int, default=12)
    ap.add_argument('--ec', default='H', choices=list('LMQH'))
    ap.add_argument('--fg', default='black'); ap.add_argument('--bg', default='white')
    ap.add_argument('--logo'); ap.add_argument('--scale', type=float, default=1.0)
    a = ap.parse_args()
    os.makedirs(a.outdir, exist_ok=True)
    if a.file:
        items = [l.strip() for l in open(a.file, encoding='utf-8') if l.strip()]
        for i, u in enumerate(items, 1):
            make_one(u, os.path.join(a.outdir, f'qr_{i}.png'), a.title, a.subtitle,
                     a.box_size, a.ec, a.fg, a.bg, a.logo, a.scale)
    else:
        assert a.url and a.out
        make_one(a.url, a.out, a.title, a.subtitle, a.box_size, a.ec, a.fg, a.bg, a.logo, a.scale)

if __name__ == '__main__':
    main()
