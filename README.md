# Noto Day Match

海外の旅行者に、能登の1日ツアー案8枚を4方向スワイプで選んでもらう聞き取りページ。

- 公開URL: https://sandboxconnections-spec.github.io/noto-day-match/
- 回答の保存先: Google スプレッドシート（Google Apps Script のウェブアプリ経由）。受け口のコードは `gas/`
- 1人1行。カードごとの選択・迷った秒数・表示順、各画面の滞在秒数、元のデータ（raw_json）まで残る
- 自分で試すとき: URLの末尾に `#test`（保存されるが「集計」シートからは外れる）
- 日本語の画面で開くとき: `?lang=ja` を付ける（例 `…/noto-day-match/?lang=ja`）。最初の画面の English／日本語 でも切り替えられる。集計は「集計（英語）」「集計（日本語）」に分かれる
- 出どころを分けるとき: `#colive` などを付ける（`src` 列に残る）
- 検索よけ: `robots.txt` と `noindex`

## 更新のしかた
- ページ: `index.html` を直して push（GitHub Pages に1〜2分で反映）
- 受け口: `gas/` で `clasp push` → `clasp deploy -i <deploymentId>`（URLを変えずに更新）
