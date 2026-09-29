# -*- coding: utf-8 -*-
"""生成答题链接的二维码 PNG。
用法：python3 make_qr.py "http://192.168.1.10:8000/"  [输出.png]
"""
import sys

url = sys.argv[1] if len(sys.argv) > 1 else "http://127.0.0.1:8000/"
out = sys.argv[2] if len(sys.argv) > 2 else "答题二维码.png"

try:
    import qrcode
except ImportError:
    print("缺少 qrcode 库，请先运行： pip install \"qrcode[pil]\"")
    sys.exit(1)

qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_H, box_size=10, border=3)
qr.add_data(url)
qr.make(fit=True)
img = qr.make_image(fill_color="black", back_color="white")
img.save(out)
print("已生成二维码：", out, "内容：", url)
