"""試作品「小さな子ども教室の 体験会おしらせキット」を生成する。

使い方: python3 tools/build_kit.py
出力: product/01_taikenkai-kit/ 以下
  - 販売ファイル/  … 購入者に渡すファイル（編集用pptx・PDF・見本PNG）
  - preview/       … 品質確認用のPNG

文字はすべて見本（架空の教室）。購入者が自分の教室の内容に入れ替える。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
from pptx_helpers import Doc  # noqa: E402

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
OUT = os.path.join(ROOT, "product", "01_taikenkai-kit", "src")
os.makedirs(OUT, exist_ok=True)

THEMES = {
    "soda": dict(label="ソーダ", bg="#EAF6F7", main="#2A8C99", light="#CBEBEE",
                 sub="#F6C64F", accent="#EE8E7E", ink="#2F3B40", white="#FFFFFF"),
    "mikan": dict(label="みかん", bg="#FDF3E7", main="#D9692F", light="#F9DCC3",
                  sub="#7DB89E", accent="#F2B84B", ink="#3B332E", white="#FFFFFF"),
}

# 見本の文言（架空の教室）。ここを書き換えれば全サイズに反映される。
SAMPLE = dict(
    season="冬の体験会",
    title=["「やってみたい」を", "ためしに来ませんか"],
    school="おえかき教室 そらいろ",
    dates="1月16日（土）・23日（土）",
    time="10:00〜11:00",
    target="3〜6さい（未就学のお子さん）",
    target_short="3〜6さい",
    content="大きな紙に、絵の具で自由に描きます",
    fee="500円（材料費こみ）",
    fee_short="500円",
    capacity="各回4名（先着順）",
    capacity_short="各回4名",
    bring="よごれてもいい服・水筒",
    place="○○市○○町（くわしい場所はお申し込み後にお知らせします）",
    apply="InstagramのDMかLINEで、\n「お名前・年齢・希望日」を送ってください",
    deadline="しめきり：1月9日（土）",
    note="見学だけ、保護者の方のご相談だけでも大丈夫です",
)


def bubbles(d, s, t, specs):
    for cx, cy, r, key, ring in specs:
        if ring:
            d.circle(s, cx, cy, r, None, line=t[key], line_w=ring, name="飾り_泡")
        else:
            d.circle(s, cx, cy, r, t[key], name="飾り_泡")


# ---------- Instagram 1080×1350（px単位）----------
def insta(theme):
    t = THEMES[theme]
    d = Doc(1080, 1350, 9525)
    # 1枚目：表紙
    s = d.slide(bg=t["bg"], name="1枚目（表紙）")
    bubbles(d, s, t, [(960, 110, 70, "light", 0), (870, 230, 26, "sub", 0),
                      (1010, 300, 16, "main", 5), (95, 455, 18, "accent", 0)])
    d.rrect(s, 90, 100, 300, 72, t["main"], radius=0.5, name="ラベル背景")
    d.text(s, 90, 100, 300, 72, SAMPLE["season"], 30, t["white"], bold=True,
           align="center", anchor="middle", name="ラベル文字")
    d.text(s, 90, 215, 900, 230, "\n".join(SAMPLE["title"]), 58, t["ink"], bold=True,
           spacing=1.15, name="タイトル")
    d.scribble(s, 96, 452, 560, 9, 9, t["sub"], 5, name="飾り_下線")
    d.text(s, 90, 490, 900, 50, SAMPLE["school"], 26, t["main"], bold=True, name="教室名")
    ph = d.rrect(s, 90, 565, 900, 430, t["light"], radius=0.08, name="写真枠（写真に差し替え）")
    d.text(s, 90, 565, 900, 430, "ここに教室の写真\n（作品や道具の写真もおすすめ）", 24,
           t["main"], align="center", anchor="middle", name="写真枠の説明（差し替え後に削除）")
    d.wave(s, 0, 1050, 1080, 300, 16, 2.5, t["main"], name="下の帯")
    d.text(s, 90, 1095, 900, 70, SAMPLE["dates"], 40, t["white"], bold=True,
           align="center", name="日付")
    d.text(s, 90, 1172, 900, 50,
           f'{SAMPLE["time"]}　｜　{SAMPLE["target_short"]}　｜　{SAMPLE["fee_short"]}', 24,
           t["white"], align="center", name="時間・対象・料金")
    d.text(s, 90, 1252, 900, 44, "くわしくは次の画像へ　→", 22, t["white"], align="center",
           name="案内")

    # 2枚目：くわしい内容
    s = d.slide(bg=t["bg"], name="2枚目（くわしい内容）")
    d.rect(s, 0, 0, 1080, 190, t["main"], name="上の帯")
    d.wave(s, 0, 170, 1080, 40, 12, 2.5, t["bg"], name="波")
    d.text(s, 90, 52, 900, 80, "体験会のくわしい内容", 40, t["white"], bold=True,
           align="center", anchor="middle", name="見出し")
    rows = [("日時", f'{SAMPLE["dates"]}\n{SAMPLE["time"]}'),
            ("対象", SAMPLE["target"]),
            ("内容", SAMPLE["content"]),
            ("料金", SAMPLE["fee"]),
            ("定員", SAMPLE["capacity"]),
            ("持ち物", SAMPLE["bring"]),
            ("場所", SAMPLE["place"]),
            ("申込", SAMPLE["apply"])]
    heights = [118, 72, 72, 72, 72, 72, 118, 118]
    y = 248
    for (k, v), h in zip(rows, heights):
        d.rrect(s, 80, y + 4, 150, 52, t["light"], radius=0.5, name=f"項目名背景_{k}")
        d.text(s, 80, y + 4, 150, 52, k, 22, t["main"], bold=True, align="center",
               anchor="middle", name=f"項目名_{k}")
        d.text(s, 262, y + 8, 740, h, v, 23, t["ink"], spacing=1.3, name=f"内容_{k}")
        y += h + 24
    d.rrect(s, 80, 1225, 920, 80, t["white"], radius=0.5, line=t["sub"], line_w=4,
            name="ひとこと枠")
    d.text(s, 80, 1225, 920, 80, SAMPLE["note"], 22, t["ink"], align="center",
           anchor="middle", name="ひとこと")
    return d


# ---------- A4チラシ（mm単位）----------
def flyer(d, theme):
    t = THEMES[theme]
    s = d.slide(bg=t["white"], name=f"A4チラシ（{t['label']}）")
    d.rect(s, 0, 0, 210, 297, t["bg"], name="背景")
    bubbles(d, s, t, [(186, 22, 15, "light", 0), (166, 44, 5, "sub", 0),
                      (196, 56, 3.5, "main", 1.0), (190, 78, 3, "accent", 0)])
    d.rrect(s, 15, 15, 58, 13, t["main"], radius=0.5, name="ラベル背景")
    d.text(s, 15, 15, 58, 13, SAMPLE["season"], 15, t["white"], bold=True, align="center",
           anchor="middle", name="ラベル文字")
    d.text(s, 15, 35, 180, 40, "\n".join(SAMPLE["title"]), 32, t["ink"], bold=True,
           spacing=1.12, name="タイトル")
    d.scribble(s, 16, 69.5, 105, 1.7, 9, t["sub"], 1.0, name="飾り_下線")
    d.text(s, 15, 74.5, 180, 9, SAMPLE["school"], 15, t["main"], bold=True, name="教室名")
    d.rrect(s, 15, 87, 180, 78, t["light"], radius=0.08, name="写真枠（写真に差し替え）")
    d.text(s, 15, 87, 180, 78, "ここに教室の写真", 14, t["main"], align="center",
           anchor="middle", name="写真枠の説明（差し替え後に削除）")
    # 日付の帯
    d.rrect(s, 15, 172, 180, 24, t["main"], radius=0.3, name="日付背景")
    d.text(s, 15, 174, 180, 12, SAMPLE["dates"], 19, t["white"], bold=True, align="center",
           name="日付")
    d.text(s, 15, 186, 180, 8, SAMPLE["time"], 12, t["white"], align="center", name="時間")
    rows = [("対象", SAMPLE["target"]), ("内容", SAMPLE["content"]),
            ("料金", SAMPLE["fee"]), ("定員", SAMPLE["capacity"]),
            ("持ち物", SAMPLE["bring"]), ("場所", SAMPLE["place"])]
    y = 204
    for k, v in rows:
        d.rrect(s, 15, y, 22, 7, t["light"], radius=0.5, name=f"項目名背景_{k}")
        d.text(s, 15, y, 22, 7, k, 10, t["main"], bold=True, align="center", anchor="middle",
               name=f"項目名_{k}")
        d.text(s, 41, y + 0.6, 112, 12, v, 10.5, t["ink"], spacing=1.2, name=f"内容_{k}")
        y += 10
    # 申込
    d.rrect(s, 15, 266, 140, 19, t["white"], radius=0.25, line=t["sub"], line_w=0.8,
            name="申込枠")
    d.text(s, 20, 268.5, 132, 16,
           [("お申し込み　" + SAMPLE["deadline"], dict(bold=True, color=t["main"], size=10.5)),
            (SAMPLE["apply"].replace("\n", ""), dict(size=9.5))],
           10, t["ink"], spacing=1.2, name="申込方法")
    d.rrect(s, 163, 255, 32, 32, t["white"], radius=0.08, line=t["main"], line_w=0.5,
            name="QRコード枠（差し替え）")
    d.text(s, 163, 255, 32, 32, "QRコード\nを置く", 9, t["main"], align="center",
           anchor="middle", name="QR枠の説明（差し替え後に削除）")
    d.text(s, 15, 288.5, 140, 5, SAMPLE["note"], 8.5, t["ink"], name="ひとこと")


# ---------- note見出し画像 1280×670（px）----------
def header():
    t = THEMES["soda"]
    o = THEMES["mikan"]
    d = Doc(1280, 670, 9525)
    s = d.slide(bg=t["bg"], name="note見出し画像")
    bubbles(d, s, t, [(1210, 70, 60, "light", 0), (1120, 150, 20, "sub", 0),
                      (1225, 615, 40, "light", 0), (1140, 640, 12, "accent", 0)])
    d.text(s, 70, 85, 700, 60, "小さな子ども教室の先生へ", 30, t["main"], bold=True,
           name="対象")
    d.text(s, 70, 150, 700, 260, "体験会の\nおしらせキット", 70, t["ink"], bold=True,
           spacing=1.1, name="商品名")
    d.scribble(s, 74, 400, 420, 8, 8, t["sub"], 5, name="飾り")
    d.text(s, 70, 440, 700, 120,
           "Instagram投稿2枚＋A4チラシ\n文字を入れ替えるだけで完成",
           24, t["ink"], spacing=1.4, name="説明")
    d.text(s, 70, 590, 400, 40, "Sö-dä fuü", 24, t["main"], bold=True, name="ブランド名")
    # 右側：作品イメージ（縮小した2枚の見本を図形で表現）
    for i, (th, x, y, rot) in enumerate([(o, 905, 150, 6), (t, 800, 130, 0)]):
        d.rrect(s, x, y, 300, 380, th["bg"], radius=0.05, line=th["main"], line_w=3,
                rot=rot, name=f"見本{i+1}_台紙")
    # 手前の見本の中身（回転なしで近似配置）
    d.rrect(s, 830, 165, 110, 26, t["main"], radius=0.5, name="見本_ラベル")
    d.text(s, 830, 210, 250, 100, "「やってみたい」を\nためしに来ませんか", 20, t["ink"],
           bold=True, name="見本_タイトル")
    d.rrect(s, 830, 290, 240, 120, t["light"], radius=0.1, name="見本_写真枠")
    d.rrect(s, 830, 430, 240, 50, t["main"], radius=0.3, name="見本_日付帯")
    return d


def main():
    for theme in THEMES:
        insta(theme).save(os.path.join(OUT, f"instagram_{theme}.pptx"))
    d = Doc(210, 297, 36000)
    for theme in THEMES:
        flyer(d, theme)
    d.save(os.path.join(OUT, "a4_flyer.pptx"))
    header().save(os.path.join(OUT, "note_header.pptx"))
    print("built", OUT)


if __name__ == "__main__":
    main()


# ---------- 使い方ガイド（A4縦・mm単位）----------
def guide():
    t = THEMES["soda"]
    d = Doc(210, 297, 36000)

    def page(title, num):
        s = d.slide(bg=t["white"], name=title)
        d.rect(s, 0, 0, 210, 30, t["bg"], name="上の帯")
        d.text(s, 18, 9, 150, 14, title, 18, t["main"], bold=True, anchor="middle", name="見出し")
        d.text(s, 160, 12, 32, 8, f"{num} / 4", 9, t["ink"], align="right", name="ページ")
        d.text(s, 18, 284, 174, 6, "Sö-dä fuü　体験会のおしらせキット　v0.1", 8, t["main"],
               name="フッター")
        return s

    def body(s, y, lines, size=11, h=None, x=18, w=174):
        return d.text(s, x, y, w, h or 200, lines, size, t["ink"], spacing=1.45)

    # 1 表紙・収録内容
    s = d.slide(bg=t["bg"], name="表紙")
    bubbles(d, s, t, [(180, 30, 16, "light", 0), (160, 52, 5, "sub", 0),
                      (190, 62, 3, "main", 0.8)])
    d.text(s, 18, 30, 170, 10, "小さな子ども教室の先生へ", 14, t["main"], bold=True)
    d.text(s, 18, 44, 174, 40, "体験会の\nおしらせキット", 34, t["ink"], bold=True, spacing=1.1)
    d.scribble(s, 19, 88, 90, 1.6, 8, t["sub"], 1.0)
    d.text(s, 18, 95, 174, 10, "使い方ガイド", 16, t["ink"], bold=True)
    d.rrect(s, 18, 115, 174, 100, t["white"], radius=0.05, name="収録内容の枠")
    body(s, 123, [
        ("収録内容", dict(bold=True, color=t["main"], size=13)),
        ("・Instagram投稿用（1080×1350px）2枚組 × 2色", {}),
        ("　1枚目：表紙　2枚目：くわしい内容", dict(size=10)),
        ("・A4チラシ（印刷用）× 2色", {}),
        ("・このガイド（使い方／告知チェックリスト／文例／利用範囲）", {}),
        ("色：ソーダ（水色）、みかん（オレンジ）", dict(space_before=8)),
        ("必要なもの", dict(bold=True, color=t["main"], size=13, space_before=10)),
        ("・pptxファイルを編集できるアプリ（どれか1つ）", {}),
        ("　PowerPoint、Keynote（iPhone・Mac）、Googleスライド", dict(size=10)),
        ("・チラシを印刷する場合はA4のプリンター", {}),
    ], x=26, w=158)
    d.text(s, 18, 272, 174, 12, "※見本の教室名・日時・料金などは、すべて架空のものです。",
           9, t["ink"])

    # 2 使い方
    s = page("使い方（3ステップ）", 2)
    steps = [
        ("1　ファイルを開く",
         "使う色のpptxファイルを、PowerPoint・Keynote・Googleスライドのどれかで開きます。"
         "Googleスライドは、Googleドライブにアップロードしてから開きます。"),
        ("2　文字と写真を入れ替える",
         "文字をタップして、自分の教室の内容に書き換えます。\n"
         "水色（オレンジ）の四角は写真を入れる場所です。写真を入れたら、"
         "「ここに教室の写真」の文字は消してください。"
         "チラシのQRコードの枠も同じように差し替えます。"),
        ("3　画像やPDFにする",
         "アプリの「書き出し」や「ダウンロード」から保存します。\n"
         "・Instagram用：PNGかJPEG\n・チラシ：PDF（そのまま印刷できます）"),
    ]
    y = 40
    for head, txt in steps:
        d.rrect(s, 18, y, 174, 9, t["light"], radius=0.5)
        d.text(s, 22, y, 166, 9, head, 12, t["main"], bold=True, anchor="middle")
        body(s, y + 12, txt, size=10.5, h=40)
        y += 58
    d.text(s, 18, 214, 174, 70, [
        ("うまくいかないとき", dict(bold=True, color=t["main"], size=12)),
        ("・文字の形が見本と違う：見本は無料フォント「Zen Maru Gothic」を使っています。"
         "入っていない環境では別のフォントで表示されます。好きな丸ゴシック体に変えて使ってください。",
         {}),
        ("・文字がはみ出す：文字を少し減らすか、文字の大きさを1〜2段階小さくしてください。", {}),
        ("・アプリの細かい操作は、アプリごとに違います。各アプリのヘルプもあわせて確認してください。", {}),
    ], 10, t["ink"], spacing=1.45)

    # 3 チェックリスト
    s = page("告知に入れる項目チェックリスト", 3)
    d.text(s, 18, 36, 174, 10,
           "書く前と、公開する前に見直すためのリストです。全部入れる必要はありません。",
           10.5, t["ink"])
    items = [
        ("何の会か", "体験会・見学会・説明会など"),
        ("日にちと曜日", "カレンダーで曜日を確認"),
        ("時間", "始まりと終わり"),
        ("対象の年齢", "「3〜6さい」など"),
        ("何をするか", "1文で。持ち帰るものがあれば書く"),
        ("料金", "材料費が入っているか"),
        ("定員と決め方", "先着順・抽選など"),
        ("持ち物・服装", "よごれてもいい服など"),
        ("場所", "自宅教室は、どこまで公開するか決める"),
        ("申込方法", "何を送ってもらうか（名前・年齢・希望日）"),
        ("しめきり", "日付と曜日"),
        ("安心できる一言", "見学だけ・相談だけでもよい、など"),
    ]
    y = 50
    for k, v in items:
        d.rrect(s, 20, y + 0.8, 5, 5, t["white"], radius=0.2, line=t["main"], line_w=0.5)
        d.text(s, 29, y, 50, 7, k, 11, t["ink"], bold=True)
        d.text(s, 80, y + 0.3, 110, 7, v, 10, t["ink"])
        y += 11
    d.rrect(s, 18, 186, 174, 66, t["bg"], radius=0.05)
    body(s, 192, [
        ("公開・印刷の前に", dict(bold=True, color=t["main"], size=12)),
        ("・日付と曜日が合っているか", {}),
        ("・料金や時間の数字が、全部の画像とチラシでそろっているか", {}),
        ("・電話番号、URL、QRコードが正しく開けるか", {}),
        ("・写真に写っている人の許可があるか（お子さんの写真は保護者の許可）", {}),
        ("・住所や家の外観など、出したくない情報が写っていないか", {}),
        ("・チラシは、試しに1枚印刷して文字が切れていないか", {}),
    ], size=10.5, x=24, w=162)

    # 4 文例・利用範囲
    s = page("文例と利用範囲", 4)
    body(s, 38, [
        ("タイトルの文例（そのまま使えます）", dict(bold=True, color=t["main"], size=12)),
        ("・「やってみたい」を、ためしに来ませんか", {}),
        ("・はじめての○○、1回だけ体験してみませんか", {}),
        ("・冬の体験会のおしらせ", {}),
        ("・春から通える教室の、体験会をひらきます", {}),
        ("ひとことの文例", dict(bold=True, color=t["main"], size=12, space_before=8)),
        ("・見学だけ、保護者の方のご相談だけでも大丈夫です", {}),
        ("・とちゅうで疲れたら、休みながら参加できます", {}),
        ("・きょうだいで参加したい場合は、ご相談ください", {}),
        ("DM・LINEで送るお知らせ文", dict(bold=True, color=t["main"], size=12, space_before=8)),
        ("「冬の体験会のおしらせです。1月16日（土）と23日（土）の10時から、"
         "3〜6さいのお子さん向けに体験会をひらきます。料金は500円（材料費こみ）です。"
         "参加をご希望の方は、お名前・年齢・希望日を送ってください。」", dict(size=10)),
    ], size=10.5, h=120)
    d.rrect(s, 18, 186, 174, 72, t["bg"], radius=0.05)
    body(s, 192, [
        ("利用範囲", dict(bold=True, color=t["main"], size=12)),
        ("できること", dict(bold=True, size=10.5)),
        ("・ご自身の教室の告知に使う（SNS投稿・印刷・配布）。何回でも編集できます。", {}),
        ("・色や文字、配置を自由に変える。", {}),
        ("できないこと", dict(bold=True, size=10.5, space_before=4)),
        ("・テンプレートのファイルを、ほかの人に渡したり、販売・配布したりすること。", {}),
        ("・このテンプレートを使って、ほかの教室の告知を作って納品・販売すること。", {}),
        ("※このキットは告知づくりの手間を減らすためのものです。集客の結果を約束するものではありません。",
         dict(size=9, space_before=4)),
    ], size=10, x=24, w=162)
    return d


def build_guide():
    guide().save(os.path.join(OUT, "guide.pptx"))


if __name__ == "__main__":
    build_guide()
