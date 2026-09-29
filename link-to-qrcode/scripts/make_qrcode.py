#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通用链接 / 文本转二维码生成器。

生成高纠错等级（H 级）二维码 PNG，可选在上下方加中文标题与提示文字，
可选在二维码中心嵌入 logo 图片，并自动回读校验解码结果。

依赖: qrcode[pil]  (pip install "qrcode[pil]")
校验依赖: opencv-python（可选，缺失时自动跳过校验）

用法示例:
  python make_qrcode.py --url "https://example.com" --title "扫码访问官网"
  python make_qrcode.py --url "https://a.com" --title "标题" --subtitle "提示" --out out/a.png
  python make_qrcode.py --file links.txt --outdir out      # 批量，每行 一个URL[|标题]
  python make_qrcode.py --url "https://a.com" --no-text    # 只出纯二维码
  python make_qrcode.py --url "https://a.com" --logo logo.png
"""
import argparse
import os
import re
import sys
import time

try:
    import qrcode
    from qrcode.constants import ERROR_CORRECT_L, ERROR_CORRECT_M, ERROR_CORRECT_Q, ERROR_CORRECT_H
except ImportError:
    # 沙箱会话重置后第三方库可能丢失，这里自动补齐，避免任务中断。
    import subprocess
    print("[init] 首次运行，正在安装依赖 qrcode[pil] ...", file=sys.stderr)
    subprocess.run([sys.executable, "-m", "pip", "install", "-q", "qrcode[pil]"], check=False)
    try:
        import qrcode
        from qrcode.constants import ERROR_CORRECT_L, ERROR_CORRECT_M, ERROR_CORRECT_Q, ERROR_CORRECT_H
    except ImportError:
        sys.exit('缺少依赖，请手动执行: pip install "qrcode[pil]"')

from PIL import Image, ImageDraw, ImageFont

# ---------------------------------------------------------------- 字体解析
FONT_CANDIDATES = [
    "/usr/share/fonts/opentype/noto/NotoSerifCJK-Bold.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "C:/Windows/Fonts/msyh.ttc",
    "/System/Library/Fonts/PingFang.ttc",
]


def find_font(extra=None):
    """返回可用的中文字体路径，找不到时回退 PIL 默认字体（None）。"""
    for p in ([extra] if extra else []) + FONT_CANDIDATES:
        if p and os.path.exists(p):
            return p
    return None


def load_font(path, size):
    if not path:
        return ImageFont.load_default()
    try:
        return ImageFont.truetype(path, size, index=0)
    except Exception:
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            return ImageFont.load_default()


def text_width(draw, text, font):
    bbox = draw.textbbox((0, 0), text, font=font)
    return bbox[2] - bbox[0], bbox


def fit_font(draw, text, font_path, max_px, start, min_size=16):
    """自动收缩字号，使 text 宽度不超过 max_px。"""
    size = start
    while size > min_size:
        f = load_font(font_path, size)
        w, _ = text_width(draw, text, f)
        if w <= max_px:
            return f
        size -= 2
    return load_font(font_path, min_size)


EC_MAP = {"L": ERROR_CORRECT_L, "M": ERROR_CORRECT_M, "Q": ERROR_CORRECT_Q, "H": ERROR_CORRECT_H}

# ---------------------------------------------------------------- 核心生成
def build_qrcode(data, box_size, border, ec, fg, bg):
    qr = qrcode.QRCode(version=None, error_correction=EC_MAP[ec], box_size=box_size, border=border)
    qr.add_data(data)
    qr.make(fit=True)
    return qr.make_image(fill_color=fg, back_color=bg).convert("RGB")


def paste_logo(img, logo_path, ratio=0.20):
    """在二维码中心粘贴 logo（白底留白），H 级纠错下安全。"""
    logo = Image.open(logo_path).convert("RGBA")
    side = int(min(img.size) * ratio)
    logo = logo.copy()
    logo.thumbnail((side, side), Image.LANCZOS)
    pad = int(side * 0.12)
    bg = Image.new("RGBA", (logo.size[0] + pad * 2, logo.size[1] + pad * 2), (255, 255, 255, 255))
    bg.paste(logo, (pad, pad), logo)
    pos = ((img.size[0] - bg.size[0]) // 2, (img.size[1] - bg.size[1]) // 2)
    img.paste(bg, pos, bg)
    return img


def render(data, title="", subtitle="", box_size=16, border=2, ec="H",
           fg="#111111", bg="white", logo=None, font_path=None, margin=1.0):
    img = build_qrcode(data, box_size, border, ec, fg, bg)
    if logo:
        img = paste_logo(img, logo)

    if not title and not subtitle:
        return img

    W = img.size[0]
    font_path = find_font(font_path)
    probe = ImageDraw.Draw(Image.new("RGB", (10, 10)))

    # 预留装饰性留白
    pad_y = int(W * 0.055)
    has_t = bool(title)
    has_s = bool(subtitle)
    canvas_h = img.size[1] + pad_y + (int(W * 0.075) if has_t else 0) + (int(W * 0.07) if has_s else 0) + pad_y
    canvas = Image.new("RGB", (W, canvas_h), "white" if bg == "white" else bg)
    canvas.paste(img, (0, pad_y))
    draw = ImageDraw.Draw(canvas)

    def draw_center(text, font, y, fill):
        w, bbox = text_width(draw, text, font)
        draw.text(((W - w) / 2 - bbox[0], y - bbox[1]), text, font=font, fill=fill)

    y = int(pad_y * 0.35)
    if has_t:
        f = fit_font(probe, title, font_path, int(W * 0.88), int(W * 0.068))
        draw_center(title, f, y, fg)
    if has_s:
        f2 = fit_font(probe, subtitle, font_path, int(W * 0.90), int(W * 0.036))
        y2 = pad_y + img.size[1] + int(W * 0.022)
        draw_center(subtitle, f2, y2, "#555555")
    return canvas


def verify(path, expect=None):
    """回读校验：用 OpenCV 解码生成的二维码。返回 (ok, decoded)。"""
    try:
        import cv2
    except ImportError:
        return None, "(未安装 opencv，跳过校验)"
    img = cv2.imread(path)
    if img is None:
        return False, "(读取失败)"
    data, _, _ = cv2.QRCodeDetector().detectAndDecode(img)
    if not data:
        return False, "(未识别出二维码)"
    if expect is not None and data != expect:
        return False, data
    return True, data


def safe_name(s, fallback="qrcode"):
    s = re.sub(r"^https?://", "", s or "")
    s = re.sub(r"[^\w\u4e00-\u9fff.-]+", "_", s).strip("_")
    return (s[:60] or fallback)


def gen_one(data, out, title="", subtitle="", **kw):
    if kw.get("do_verify", True):
        ok0 = True
    img = render(data, title=title, subtitle=subtitle, **{k: v for k, v in kw.items() if k in (
        "box_size", "border", "ec", "fg", "bg", "logo", "font_path")})
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    img.save(out, dpi=(300, 300))
    res = {"path": out, "size": img.size, "data": data}
    if kw.get("do_verify", True):
        ok, decoded = verify(out, expect=data)
        res["verified"] = ok
        res["decoded"] = decoded
    return res


def main():
    ap = argparse.ArgumentParser(description="链接/文本转二维码生成器")
    src = ap.add_mutually_exclusive_group(required=True)
    src.add_argument("--url", help="要编码的链接或文本")
    src.add_argument("--file", help="批量文件，每行 `URL[|标题]`，空行与 # 开头行忽略")
    ap.add_argument("--title", default="", help="二维码上方标题（中文 OK）")
    ap.add_argument("--subtitle", default="", help="二维码下方提示文字")
    ap.add_argument("--out", default="", help="输出 PNG 路径（单条模式）")
    ap.add_argument("--outdir", default="", help="批量模式输出目录")
    ap.add_argument("--box-size", type=int, default=16, help="单格像素，越大图越大（默认 16）")
    ap.add_argument("--border", type=int, default=2, help="静默边格数（默认 2）")
    ap.add_argument("--ec", default="H", choices=list(EC_MAP), help="纠错等级，默认 H（最抗污损）")
    ap.add_argument("--fg", default="#111111", help="前景（码点）颜色")
    ap.add_argument("--bg", default="white", help="背景颜色")
    ap.add_argument("--logo", default="", help="中心 logo 图片路径")
    ap.add_argument("--font", default="", help="指定中文字体文件")
    ap.add_argument("--no-text", action="store_true", help="只输出纯二维码，不加任何文字")
    ap.add_argument("--scale", type=int, default=1, help="整体放大倍数（整数）")
    ap.add_argument("--no-verify", action="store_true", help="跳过解码回读校验")
    args = ap.parse_args()

    jobs = []
    if args.url:
        jobs.append((args.url, args.title, args.out))
    else:
        with open(args.file, encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                parts = line.split("|", 1)
                data = parts[0].strip()
                t = parts[1].strip() if len(parts) > 1 else ""
                outdir = args.outdir or "."
                jobs.append((data, t, os.path.join(outdir, safe_name(data) + ".png")))
        if not args.outdir:
            os.makedirs("qrcodes", exist_ok=True)
            jobs = [(d, t, os.path.join("qrcodes", os.path.basename(o))) for d, t, o in jobs]

    results = []
    for data, title, out in jobs:
        if not out:
            os.makedirs("qrcodes", exist_ok=True)
            out = os.path.join("qrcodes", safe_name(data) + ".png")
        r = gen_one(
            data, out,
            title=("" if args.no_text else title),
            subtitle=("" if args.no_text else args.subtitle),
            box_size=args.box_size, border=args.border, ec=args.ec,
            fg=args.fg, bg=args.bg, logo=args.logo, font_path=args.font,
            do_verify=not args.no_verify,
        )
        if args.scale > 1:
            im = Image.open(r["path"])
            im = im.resize((im.size[0] * args.scale, im.size[1] * args.scale), Image.LANCZOS)
            im.save(r["path"], dpi=(300, 300))
            r["size"] = im.size
        results.append(r)

    for r in results:
        flag = {True: "✔ 校验通过", False: "✘ 校验失败", None: "○ 未校验"}[r.get("verified")]
        print(f"[{flag}] {r['path']}  {r['size'][0]}x{r['size'][1]}  -> {r['data'][:70]}")
    bad = [r for r in results if r.get("verified") is False]
    if bad:
        print(f"\n⚠ {len(bad)} 个二维码回读失败，建议提高 --box-size 或改用 --ec H", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
