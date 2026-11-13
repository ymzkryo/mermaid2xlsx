# mermaid2xlsx

Mermaid のシーケンス図を、**Excel のネイティブ図形**（四角・線・矢印・テキスト）で描いた `.xlsx` に変換する。
納品物で「Excel の図形で編集できること」が求められる場面向け。

- 配置は Mermaid に任せる。`mmdc` で SVG を書き出し、その座標を DrawingML の図形に写す
- マクロ不要。Mac / Windows どちらの Excel でも開くだけで編集できる
- 図全体が 1 つのグループになっている。個別に動かすときは、クリックして中の図形を選ぶか、グループ解除する

## 使い方

```bash
python3 mermaid2xlsx.py examples/sample.mmd -o out/sample.xlsx   # .mmd → mmdc → .xlsx
python3 mermaid2xlsx.py a.mmd b.mmd -o docs.xlsx                  # 1 図 1 シート
python3 mermaid2xlsx.py examples/sample.svg -o out/sample.xlsx   # 書き出し済み SVG から
```

変換結果の見本は `sample-out/`（`examples/sample.mmd` を変換した `.xlsx` と、それを Mac の Excel で PDF に書き出したもの）。
`out/` は手元で変換するときの置き場で、中身は git に入らない。

| オプション | 内容 |
| --- | --- |
| `--font` | 図形の文字のフォント（既定: 游ゴシック。Mac / Win の両方にあるもの） |
| `--scale` | 拡大率（既定 1.0） |
| `--no-mirror-actors` | 図の下側の参加者を描かない |
| `--keep-svg DIR` | 中間の SVG を残す |
| `--mmdc CMD` | mmdc のコマンド（既定: PATH の `mmdc`、無ければ `npx @mermaid-js/mermaid-cli@11.4.2`） |

色・線の太さ・破線の種類は `mermaid2xlsx.py` 冒頭の `THEME` で変える。

## 依存

- Python 3.11+（標準ライブラリのみ）
- `.mmd` を渡すときは Node.js（`npx`）。初回は Chromium のダウンロードが走る

## 対応している要素

参加者（`participant` / `actor`）、ライフライン、メッセージ（`->>` `-->>` `-)` `--)` `-x`）、
自分宛てメッセージ、`autonumber`、活性化、`alt` / `else` / `loop` / `opt` などの枠、ノート、`rect` の背景色。

シーケンス図以外（フローチャートなど）は未対応。

## ライセンス

[MIT](LICENSE)
