#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
一键生成「问卷 + 在线答题网站 + 数据后台」全套成果。

用法：
    python3 build_survey.py --spec questions.json --out ./输出目录 [--url http://192.168.1.10:8000/]

产物：
    输出目录/
    ├── <问卷标题>.docx        可直接打印/分发的 Word 问卷
    └── 在线问卷系统/           手机扫码答题 + 后台数据汇总（零依赖，双击即用）
        ├── server.py ... 等
        └── questions.json     本问卷的题目
"""
import argparse
import json
import os
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "..", "assets", "app")


def log(msg):
    print("[build] " + msg)


def copy_app(spec, out_dir):
    dest = os.path.join(out_dir, "在线问卷系统")
    if os.path.exists(dest):
        shutil.rmtree(dest)
    shutil.copytree(TEMPLATE, dest,
                    ignore=shutil.ignore_patterns("data", "__pycache__", "*.db"))
    with open(os.path.join(dest, "questions.json"), "w", encoding="utf-8") as f:
        json.dump(spec, f, ensure_ascii=False, indent=2)
    log("已生成在线问卷系统 -> " + dest)
    return dest


def make_word(spec, out_dir):
    title = spec.get("title", "调查问卷")
    safe = "".join(c for c in title if c not in '\\/:*?"<>|').strip() or "调查问卷"
    out_docx = os.path.join(out_dir, safe + ".docx")
    script = os.path.join(HERE, "spec_to_word.py")
    subprocess.run([sys.executable, script, os.path.join(out_dir, "在线问卷系统", "questions.json"), out_docx],
                   check=True)
    log("已生成 Word 问卷 -> " + out_docx)
    return out_docx


def make_qr(url, out_dir):
    if not url:
        return None
    try:
        try:
            import qrcode  # 优先用系统安装的完整版
            qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_H, box_size=10, border=3)
            qr.add_data(url)
            qr.make(fit=True)
            img = qr.make_image(fill_color="black", back_color="white")
            out_png = os.path.join(out_dir, "答题二维码.png")
            img.save(out_png)
            log("已生成二维码 -> " + out_png + "  内容: " + url)
            return out_png
        except Exception:
            # 退回到内置的纯矩阵实现，用 PIL 自行绘制
            import importlib
            sys.path.insert(0, os.path.join(HERE, "..", "assets", "app", "vendor"))
            mod = importlib.import_module("qrcode.main")
            consts = importlib.import_module("qrcode.constants")
            from PIL import Image, ImageDraw
            qr = mod.QRCode(error_correction=consts.ERROR_CORRECT_H, border=3)
            qr.add_data(url)
            qr.make(fit=True)
            m = qr.get_matrix()
            n = len(m)
            box = 10
            img = Image.new("RGB", (n * box, n * box), "white")
            d = ImageDraw.Draw(img)
            for y, row in enumerate(m):
                for x, val in enumerate(row):
                    if val:
                        d.rectangle([x * box, y * box, x * box + box - 1, y * box + box - 1], fill="black")
            out_png = os.path.join(out_dir, "答题二维码.png")
            img.save(out_png)
            log("已生成二维码 -> " + out_png + "  内容: " + url)
            return out_png
    except Exception as e:
        log("跳过二维码（%s）。可稍后在后台页面直接查看二维码。" % e)
        return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--spec", required=True, help="问卷题目 JSON 文件")
    ap.add_argument("--out", required=True, help="输出目录")
    ap.add_argument("--url", default="", help="答题网址（可选，用于生成二维码）")
    args = ap.parse_args()

    with open(args.spec, encoding="utf-8") as f:
        spec = json.load(f)
    os.makedirs(args.out, exist_ok=True)

    copy_app(spec, args.out)
    make_word(spec, args.out)
    make_qr(args.url, args.out)

    print()
    print("=" * 56)
    print("  全部完成！")
    print("  1) Word 问卷：  " + os.path.join(args.out, spec.get('title', '调查问卷') + '.docx'))
    print("  2) 在线系统：    进入「在线问卷系统」文件夹")
    print("     - Windows：双击 启动_Windows.bat")
    print("     - Mac/Linux：双击 启动_Mac_Linux.sh")
    print("     启动后浏览器打开 http://127.0.0.1:8000/admin 查看后台与二维码")
    print("=" * 56)


if __name__ == "__main__":
    main()
