# Sö-dä fuü 制作プロジェクト

「デザイン×子育て」の商品を、noteで販売するための制作場所です。
AI社員チーム（役割は `team/roles.md`）が、調査から販売準備までを進めます。

## まず読むファイル
1. `reports/` … 朝の報告書（いちばん新しいもの）
2. `ops/questions.md` … オーナーへの質問
3. `ops/progress.md` … 進捗と、次にやること

## フォルダ
- `team/` … 役割の定義（roles.md）、事業方針とルール（policy.md）
- `ops/` … 進捗、質問、再開方法、オーナーの実機確認チェック表
- `sales/` … 公開後の販売記録
- `research/` … 市場調査、note公式情報の確認
- `planning/` … 商品の選定と理由
- `product/` … 試作品（販売ファイル、見本画像、元データ）
- `note/` … noteに貼る記事の下書き、見出し画像、投稿手順
- `qa/` … 品質確認の結果
- `tools/` … 試作品を作るスクリプト

## 作り直す方法（作業環境で）
```
python3 tools/build_kit.py   # pptx（編集用の元データ）を作る
python3 tools/package.py     # PDF・PNG・zip・note用画像を作る
```
フォント Zen Maru Gothic を ~/.fonts に入れておく（ダウンロード元は product/01_taikenkai-kit/LICENSES.md）。
