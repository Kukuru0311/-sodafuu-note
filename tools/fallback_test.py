"""フォントが置き換わったときの崩れを試す（購入者の環境の代わり）。

使い方: python3 tools/fallback_test.py 代わりのフォント名 出力フォルダ
pptxの中の Zen Maru Gothic を別のフォントに書き換えてから、PNGにする。
"""
import os
import subprocess
import sys
import zipfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
SRC = os.path.join(ROOT, "product", "01_taikenkai-kit", "src")
font, out = sys.argv[1], sys.argv[2]
os.makedirs(out, exist_ok=True)
jobs = []
for name, width in [("instagram_soda.pptx", "540"), ("a4_flyer.pptx", "600")]:
    dst = os.path.join(out, name)
    with zipfile.ZipFile(os.path.join(SRC, name)) as zin, \
            zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith(".xml"):
                data = data.replace(b"Zen Maru Gothic", font.encode("utf-8"))
            zout.writestr(item, data)
    jobs += [dst, out, width]
subprocess.run(["python3", os.path.join(ROOT, "tools", "render.py")] + jobs, check=True)
