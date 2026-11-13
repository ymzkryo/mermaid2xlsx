# mermaid2xlsx

Mermaid の図を、**Excel のネイティブ図形**（四角・線・矢印・テキスト）で描いた `.xlsx` に変換する。
シーケンス図・フローチャート・状態遷移図・ER 図・クラス図に対応している。
納品物で「Excel の図形で編集できること」が求められる場面向け。

- 配置は Mermaid に任せる。`mmdc` で SVG を書き出し、その座標を DrawingML の図形に写す
- マクロ不要。Mac / Windows どちらの Excel でも開くだけで編集できる
- 図全体が 1 つのグループになっている。個別に動かすときは、クリックして中の図形を選ぶか、グループ解除する

## 使い方

```bash
python3 mermaid2xlsx.py examples/sample.mmd -o out/sample.xlsx   # .mmd → mmdc → .xlsx
python3 mermaid2xlsx.py examples/*.mmd -o out/all.xlsx            # 1 図 1 シート
python3 mermaid2xlsx.py examples/sample.svg -o out/sample.xlsx   # 書き出し済み SVG から
```

変換結果の見本は `sample-out/`。どちらも `.xlsx` と、それを Mac の Excel で PDF に書き出したもの。

- `sample.xlsx` / `sample.pdf`: シーケンス図（`examples/sample.mmd`）
- `diagrams.xlsx` / `diagrams.pdf`: フローチャート・状態遷移図・ER 図・クラス図（`examples/` のほかの `.mmd`）。1 図 1 シート

`out/` は手元で変換するときの置き場で、中身は git に入らない。

| オプション | 内容 |
| --- | --- |
| `--theme FILE` | 見た目の設定（TOML）。書いた項目だけ既定の `theme.toml` を上書きする |
| `--font` | 図形の文字のフォント。`theme.toml` の `[text] font`（既定: 游ゴシック）を上書きする |
| `--scale` | 拡大率（既定 1.0） |
| `--no-mirror-actors` | 図の下側の参加者を描かない |
| `--keep-svg DIR` | 中間の SVG を残す |
| `--mmdc CMD` | mmdc のコマンド（既定: PATH の `mmdc`、無ければ `npx @mermaid-js/mermaid-cli@11.4.2`） |

## 見た目を変える

色・線の太さ・破線の種類・矢じり・フォントは **`theme.toml`** にまとめてある。各項目の意味と書ける値はファイル内のコメントに書いた。

- **シーケンス図**: `[actor]` `[message]` `[note]` など、要素ごとの色・線を `theme.toml` で決める
- **それ以外の図**: 色・線は Mermaid のテーマで決まり、その見た目をそのまま写す。`[mermaid] theme` で選ぶ（既定は白黒・グレー基調の `neutral`）
- **フォント**（`[text] font`）はすべての図に効く

- 全体の既定を変えるなら `theme.toml` を直接直す
- 案件ごとに変えるなら、**変えたいところだけ**を書いた TOML を用意して `--theme` で渡す

```toml
# my-theme.toml — 参加者の箱を青系にして、ノートの枠線を消す
[actor]
fill = "DCEBFF"
line = "1F4E79"

[note]
line = "none"
```

```bash
python3 mermaid2xlsx.py examples/sample.mmd -o out/sample.xlsx --theme my-theme.toml
```

書き間違い（色が 6 桁でない、知らない項目名、使えない破線の種類など）は、変換を始める前にまとめて表示して止まる。

## 依存

| 依存 | 必要な場面 | 内容 |
| --- | --- | --- |
| **Python 3.11 以上** | 常に | 標準ライブラリだけで動く。`pip install` は不要。`theme.toml` の読み込みに 3.11 からの `tomllib` を使う。3.11 / 3.12 / 3.13 / 3.14 で同じ出力になることを確認済み |
| **Node.js（`npx`）** | `.mmd` を渡すとき | SVG の書き出しに使う。`.svg` を渡すときは不要。Node.js 23 で確認済み |
| **[mermaid-cli](https://github.com/mermaid-js/mermaid-cli)** | `.mmd` を渡すとき | インストール不要。`npx` で `@mermaid-js/mermaid-cli@11.4.2` を固定して呼ぶ。PATH に `mmdc` があればそちらを使い、`--mmdc` で差し替えもできる |
| **Chromium** | `.mmd` を渡すとき | mermaid-cli が内部で使う（puppeteer 経由）。**初回の実行時に自動でダウンロードされる**ので、その回だけ時間がかかる |

### 出力した `.xlsx` を開く側

- **Excel**: Microsoft 365 の Mac 版で確認済み。**Windows 版は未確認**
- **フォント**: 既定の游ゴシックは Mac / Windows の両方に入っている。`--font` で変えるときは、開く側の環境にあるものを選ぶ

## 対応している図

| 図 | 変換のしかた | 確認した要素 |
| --- | --- | --- |
| **シーケンス図** | 図の意味を読み取り、`theme.toml` の色で描き直す | 参加者（`participant` / `actor`）、ライフライン、メッセージ（`->>` `-->>` `-)` `--)` `-x`）、自分宛て、`autonumber`、活性化、`alt` / `else` / `loop` / `opt`、ノート、`rect` の背景色 |
| **フローチャート** | SVG の見た目をそのまま写す | 四角・角丸・ひし形・円・円柱・平行四辺形・旗形、実線・点線・太線の矢印、矢印のラベル、`subgraph` |
| **状態遷移図** | 同上 | 状態、開始・終了、遷移とラベル、入れ子の状態、ノート |
| **ER 図** | 同上 | エンティティ（列の表）、リレーションと多重度の記号（鳥の足）、ラベル |
| **クラス図** | 同上 | クラス（3 段の箱）、`<<interface>>`、継承（白抜きの三角）・コンポジション（ひし形）・依存・関連、多重度、ラベル |

シーケンス図以外の変換について:

- **ノードの文字は図形の中に入れる**ので、箱をダブルクリックすればそのまま書き換えられる。ER 図とクラス図の表は、各セルの文字が独立したテキストボックスになる
- **矢じり**: 三角形の矢じりは Excel の線の矢じりにする（線を動かしても付いてくる）。白抜きの三角・ひし形・鳥の足など Excel に無い形は、記号を図形として線の端に置く（線を動かしたら記号も動かす必要がある）
- 上の表以外の図（ガントチャート・円グラフなど）も同じ方法で写すが、確認していないので崩れることがある。そのときは警告を出す

## ライセンス

[MIT](LICENSE)
