"""販売ファイル・見本画像・見出し画像をまとめる。

使い方: python3 tools/package.py
前提: python3 tools/build_kit.py を実行済み
"""
import os
import shutil
import subprocess
import zipfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
P = os.path.join(ROOT, "product", "01_taikenkai-kit")
SRC = os.path.join(P, "src")
KIT = "SodaFuu_体験会おしらせキット_v0.1"
SALE = os.path.join(P, "販売ファイル", KIT)
SAMPLE = os.path.join(P, "見本画像")
NOTE_IMG = os.path.join(ROOT, "note", "01_taikenkai-kit")
TMP = os.path.join(P, "_render")


def render(*jobs):
    args = ["python3", os.path.join(ROOT, "tools", "render.py")]
    for j in jobs:
        args += list(j)
    subprocess.run(args, check=True, stdout=subprocess.DEVNULL)


def main():
    for d in (SALE, SAMPLE, TMP):
        shutil.rmtree(d, ignore_errors=True)
        os.makedirs(d)
    os.makedirs(NOTE_IMG, exist_ok=True)
    s = lambda n: os.path.join(SRC, n)
    render((s("instagram_soda.pptx"), TMP, "1080"), (s("instagram_mikan.pptx"), TMP, "1080"),
           (s("a4_flyer.pptx"), TMP, "1240"), (s("guide.pptx"), TMP, "-"),
           (s("note_header.pptx"), TMP, "1280"))
    t = lambda n: os.path.join(TMP, n)

    # 販売ファイル（購入者に渡すもの）
    shutil.copy(s("instagram_soda.pptx"), os.path.join(SALE, "01_Instagram用_ソーダ色_1080x1350.pptx"))
    shutil.copy(s("instagram_mikan.pptx"), os.path.join(SALE, "02_Instagram用_みかん色_1080x1350.pptx"))
    shutil.copy(s("a4_flyer.pptx"), os.path.join(SALE, "03_A4チラシ_2色.pptx"))
    shutil.copy(t("guide.pdf"), os.path.join(SALE, "00_はじめにお読みください_使い方ガイド.pdf"))
    mihon = os.path.join(SALE, "見本")
    os.makedirs(mihon)
    shutil.copy(t("instagram_soda-1.png"), os.path.join(mihon, "見本_Instagram_ソーダ色_1枚目.png"))
    shutil.copy(t("instagram_soda-2.png"), os.path.join(mihon, "見本_Instagram_ソーダ色_2枚目.png"))
    shutil.copy(t("instagram_mikan-1.png"), os.path.join(mihon, "見本_Instagram_みかん色_1枚目.png"))
    shutil.copy(t("instagram_mikan-2.png"), os.path.join(mihon, "見本_Instagram_みかん色_2枚目.png"))
    shutil.copy(t("a4_flyer.pdf"), os.path.join(mihon, "見本_A4チラシ_2色.pdf"))

    zpath = os.path.join(P, "販売ファイル", KIT + ".zip")
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        for base, _, files in os.walk(SALE):
            for f in sorted(files):
                full = os.path.join(base, f)
                z.write(full, os.path.relpath(full, os.path.dirname(SALE)))

    # 品質確認・記事用の見本画像
    for f in sorted(os.listdir(TMP)):
        if f.endswith(".png") or f.endswith(".pdf"):
            shutil.copy(t(f), os.path.join(SAMPLE, f))

    # note用：見出し画像と、記事に貼るサンプル画像（JPEG・軽量）
    from PIL import Image
    Image.open(t("note_header-1.png")).convert("RGB").save(
        os.path.join(NOTE_IMG, "見出し画像_1280x670.png"))
    for src, dst in [("instagram_soda-1.png", "記事用サンプル1_Instagram表紙.jpg"),
                     ("instagram_soda-2.png", "記事用サンプル2_Instagramくわしい内容.jpg"),
                     ("a4_flyer-1.png", "記事用サンプル3_A4チラシ_ソーダ色.jpg"),
                     ("a4_flyer-2.png", "記事用サンプル4_A4チラシ_みかん色.jpg"),
                     ("guide-3.png", "記事用サンプル5_告知チェックリスト.jpg")]:
        im = Image.open(t(src)).convert("RGB")
        if im.width > 1240:
            im = im.resize((1240, round(im.height * 1240 / im.width)))
        im.save(os.path.join(NOTE_IMG, dst), quality=88)
    # 無料サンプル：チェックリストのページだけのPDF
    from pypdf import PdfReader, PdfWriter
    w = PdfWriter()
    w.add_page(PdfReader(t("guide.pdf")).pages[2])
    w.write(os.path.join(NOTE_IMG, "無料サンプル_告知チェックリスト.pdf"))

    shutil.rmtree(TMP)
    print("ok", zpath, os.path.getsize(zpath))


if __name__ == "__main__":
    main()
