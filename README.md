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

| 依存 | 必要な場面 | 内容 |
| --- | --- | --- |
| **Python 3.9 以上** | 常に | 標準ライブラリだけで動く。`pip install` は不要。3.9 / 3.10 / 3.11 / 3.12 / 3.13 / 3.14 で同じ出力になることを確認済み |
| **Node.js（`npx`）** | `.mmd` を渡すとき | SVG の書き出しに使う。`.svg` を渡すときは不要。Node.js 23 で確認済み |
| **[mermaid-cli](https://github.com/mermaid-js/mermaid-cli)** | `.mmd` を渡すとき | インストール不要。`npx` で `@mermaid-js/mermaid-cli@11.4.2` を固定して呼ぶ。PATH に `mmdc` があればそちらを使い、`--mmdc` で差し替えもできる |
| **Chromium** | `.mmd` を渡すとき | mermaid-cli が内部で使う（puppeteer 経由）。**初回の実行時に自動でダウンロードされる**ので、その回だけ時間がかかる |

### 出力した `.xlsx` を開く側

- **Excel**: Microsoft 365 の Mac 版で確認済み。**Windows 版は未確認**
- **フォント**: 既定の游ゴシックは Mac / Windows の両方に入っている。`--font` で変えるときは、開く側の環境にあるものを選ぶ

## 対応している要素

参加者（`participant` / `actor`）、ライフライン、メッセージ（`->>` `-->>` `-)` `--)` `-x`）、
自分宛てメッセージ、`autonumber`、活性化、`alt` / `else` / `loop` / `opt` などの枠、ノート、`rect` の背景色。

シーケンス図以外（フローチャートなど）は未対応。

## ライセンス

[MIT](LICENSE)
