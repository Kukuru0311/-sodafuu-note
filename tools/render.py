"""pptx を LibreOffice で PDF にし、PNG も作る。

使い方: python3 tools/render.py 入力.pptx 出力フォルダ [PNGの幅px]
- 日本語と英数字の間に自動で入るすき間（LibreOfficeの設定）を切ってから書き出す。
"""
import os
import subprocess
import sys
import time

import uno
from com.sun.star.beans import PropertyValue


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def connect():
    proc = subprocess.Popen(["soffice", "--headless", "--norestore", "--nologo",
                             "--accept=socket,host=127.0.0.1,port=2002;urp;"],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(60):
        try:
            ctx = resolver.resolve(
                "uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext")
            return proc, ctx.ServiceManager.createInstanceWithContext(
                "com.sun.star.frame.Desktop", ctx)
        except Exception:
            time.sleep(0.5)
    raise RuntimeError("LibreOfficeに接続できませんでした")


def fix_spacing(doc):
    pages = doc.getDrawPages()
    for i in range(pages.getCount()):
        page = pages.getByIndex(i)
        for j in range(page.getCount()):
            sh = page.getByIndex(j)
            if not hasattr(sh, "createEnumeration"):
                continue
            try:
                en = sh.getText().createEnumeration()
            except Exception:
                continue
            while en.hasMoreElements():
                para = en.nextElement()
                try:
                    para.setPropertyValue("ParaIsCharacterDistance", False)
                except Exception:
                    pass


def render(desktop, src, outdir, width=None):
    os.makedirs(outdir, exist_ok=True)
    url = uno.systemPathToFileUrl(os.path.abspath(src))
    doc = desktop.loadComponentFromURL(url, "_blank", 0, (prop("Hidden", True),))
    fix_spacing(doc)
    base = os.path.splitext(os.path.basename(src))[0]
    pdf = os.path.join(outdir, base + ".pdf")
    doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(pdf)),
                   (prop("FilterName", "impress_pdf_Export"),))
    doc.close(True)
    args = ["pdftoppm", "-png"]
    args += ["-scale-to-x", str(width), "-scale-to-y", "-1"] if width else ["-r", "200"]
    subprocess.run(args + [pdf, os.path.join(outdir, base)], check=True)
    return pdf


if __name__ == "__main__":
    jobs = sys.argv[1:]
    proc, desktop = connect()
    try:
        # 引数: src outdir width src outdir width ...
        for k in range(0, len(jobs), 3):
            w = int(jobs[k + 2]) if jobs[k + 2] != "-" else None
            print(render(desktop, jobs[k], jobs[k + 1], w))
    finally:
        try:
            desktop.terminate()
        except Exception:
            pass
        proc.wait(timeout=30)
