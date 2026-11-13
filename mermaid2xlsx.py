#!/usr/bin/env python3
"""Mermaid の図を、Excel のネイティブ図形で描いた .xlsx に変換する。
（シーケンス図・フローチャート・状態遷移図・ER 図・クラス図）

配置は Mermaid に任せる。mermaid-cli (mmdc) で SVG を書き出し、
その座標をそのまま DrawingML の図形（四角・線・矢印・テキスト）に写す。
マクロは使わないので、Mac / Windows どちらの Excel でも開くだけで編集できる。

使い方:
    python3 mermaid2xlsx.py examples/sample.mmd -o sample.xlsx
    python3 mermaid2xlsx.py a.mmd b.mmd -o docs.xlsx      # 1 図 1 シート
    python3 mermaid2xlsx.py examples/sample.svg            # 書き出し済みの SVG も可

依存: Python 3.11+（標準ライブラリのみ）。.mmd を渡すときは Node.js（npx）。詳しくは README。
見た目（色・線・フォント・矢じり）は同じディレクトリの theme.toml で変える。
"""

from __future__ import annotations

import sys

if sys.version_info < (3, 11):  # tomllib が 3.11 から
    sys.exit("mermaid2xlsx は Python 3.11 以上が必要です"
             f"（今は {sys.version_info.major}.{sys.version_info.minor}）")

import argparse
import json
import math
import re
import shutil
import subprocess
import tempfile
import tomllib
import unicodedata
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

MMDC_PACKAGE = "@mermaid-js/mermaid-cli@11.4.2"
SVG_NS = "{http://www.w3.org/2000/svg}"
EMU_PER_PX = 9525  # 96 dpi
EMU_PER_PT = 12700
DEFAULT_THEME = Path(__file__).resolve().parent / "theme.toml"

# SVG のマーカー id の末尾 → theme.toml の [arrows] のどれを使うか
MARKER_TO_ROLE = {
    "arrowhead": "sync",  # ->>  -->>
    "filled-head": "async",  # -)   --)
    "solidTopArrowHead": "sync",
    "solidBottomArrowHead": "sync",
    "stickTopArrowHead": "async",
    "stickBottomArrowHead": "async",
}


# ---------------------------------------------------------------------------
# 見た目の設定（theme.toml）
# ---------------------------------------------------------------------------
PRST_DASH = {"solid", "dot", "dash", "lgDash", "dashDot", "lgDashDot", "lgDashDotDot",
             "sysDash", "sysDot", "sysDashDot", "sysDashDotDot"}
ARROW_TYPES = {"triangle", "arrow", "stealth", "diamond", "oval", "none"}
ARROW_SIZES = {"sm", "med", "lg"}
MERMAID_THEMES = ("neutral", "default", "base", "forest", "dark")
STYLE_SECTIONS = {"actor", "actor_man", "lifeline", "message", "reply", "frame", "frame_sep",
                  "label_box", "note", "activation", "default"}


class ThemeError(ValueError):
    pass


def deep_merge(base: dict, over: dict) -> dict:
    out = dict(base)
    for k, v in over.items():
        out[k] = deep_merge(out[k], v) if isinstance(v, dict) and isinstance(out.get(k), dict) else v
    return out


def load_theme(override: Path | None = None) -> dict:
    """既定の theme.toml を読み、override があれば書いてある項目だけ上書きする。"""
    def read(p: Path) -> dict:
        try:
            with open(p, "rb") as f:
                return tomllib.load(f)
        except FileNotFoundError:
            raise ThemeError(f"テーマが見つかりません: {p}") from None
        except tomllib.TOMLDecodeError as e:
            raise ThemeError(f"{p} の書き方が TOML として正しくありません: {e}") from None

    theme = read(DEFAULT_THEME)
    if override is not None:
        theme = deep_merge(theme, read(override))
    validate_theme(theme)
    return theme


def validate_theme(t: dict) -> None:
    """打ち間違いを Excel で開くまで気づかない、を避けるために読み込み時に弾く。"""
    errors: list[str] = []

    def color(where: str, v) -> None:
        if not (isinstance(v, str) and (v == "none" or re.fullmatch(r"[0-9a-fA-F]{6}", v))):
            errors.append(f"{where} = {v!r}: 色は 16 進 6 桁（例 \"333333\"）か \"none\"")

    def number(where: str, v, lo: float = 0) -> None:
        if isinstance(v, bool) or not isinstance(v, (int, float)) or v < lo:
            errors.append(f"{where} = {v!r}: {lo} 以上の数")

    known = STYLE_SECTIONS | {"layout", "text", "arrows", "seq_number", "mermaid"}
    for sec in t:
        if sec not in known:
            errors.append(f"[{sec}]: 知らないセクション（使えるのは {', '.join(sorted(known))}）")
        elif not isinstance(t[sec], dict):
            # 形が崩れたセクションは以降の検査が読めないので、ここで止める
            raise ThemeError(f"テーマの設定に誤りがあります:\n  {sec} = {t[sec]!r}: "
                             f"[{sec}] はセクションとして書く")

    for sec in STYLE_SECTIONS:
        s = t.get(sec, {})
        for k, v in s.items():
            where = f"[{sec}] {k}"
            if k in ("fill", "line"):
                color(where, v)
            elif k == "w":
                number(where, v)
            elif k == "dash":
                if v not in PRST_DASH:
                    errors.append(f"{where} = {v!r}: 使えるのは {', '.join(sorted(PRST_DASH))}")
            else:
                errors.append(f"{where}: 知らない項目（使えるのは fill / line / w / dash）")

    allowed = {"layout": {"margin"}, "text": {"font", "color"}, "mermaid": {"theme"},
               "arrows": {"sync", "async", "size"}, "seq_number": {"fill", "text", "r"}}
    for sec, keys in allowed.items():
        for k in t.get(sec, {}):
            if k not in keys:
                errors.append(f"[{sec}] {k}: 知らない項目（使えるのは {', '.join(sorted(keys))}）")

    if t["mermaid"]["theme"] not in MERMAID_THEMES:
        errors.append(f"[mermaid] theme = {t['mermaid']['theme']!r}: 使えるのは {', '.join(MERMAID_THEMES)}")
    number("[layout] margin", t["layout"]["margin"])
    if not (isinstance(t["text"]["font"], str) and t["text"]["font"].strip()):
        errors.append("[text] font: フォント名を文字列で書く")
    color("[text] color", t["text"]["color"])
    for k in ("sync", "async"):
        if t["arrows"][k] not in ARROW_TYPES:
            errors.append(f"[arrows] {k} = {t['arrows'][k]!r}: 使えるのは {', '.join(sorted(ARROW_TYPES))}")
    if t["arrows"]["size"] not in ARROW_SIZES:
        errors.append(f"[arrows] size = {t['arrows']['size']!r}: 使えるのは sm / med / lg")
    color("[seq_number] fill", t["seq_number"]["fill"])
    color("[seq_number] text", t["seq_number"]["text"])
    number("[seq_number] r", t["seq_number"]["r"], lo=1)

    if errors:
        raise ThemeError("テーマの設定に誤りがあります:\n  " + "\n  ".join(errors))


# ---------------------------------------------------------------------------
# 中間表現
# ---------------------------------------------------------------------------
@dataclass
class Style:
    fill: str | None = None  # None = 塗りなし
    line: str | None = None  # None = 線なし
    w: float = 1.0  # pt
    dash: str | None = None  # DrawingML の prstDash
    fill_alpha: float = 1.0  # 0〜1
    line_alpha: float = 1.0


@dataclass
class TextSpec:
    lines: list[str]
    size_px: float = 16
    bold: bool = False
    color: str = "000000"
    align: str = "ctr"  # l / ctr / r


@dataclass
class Shape:
    kind: str  # rect / roundrect / ellipse / textbox / line / path
    name: str
    x: float = 0
    y: float = 0
    w: float = 0
    h: float = 0
    style: Style = field(default_factory=Style)
    text: TextSpec | None = None
    # line
    x1: float = 0
    y1: float = 0
    x2: float = 0
    y2: float = 0
    head: str | None = None  # 始点の矢じり
    tail: str | None = None  # 終点の矢じり
    # path: [("M", [(x,y)]), ("L", [(x,y)]), ("C", [(x,y)*3]), ("Z", [])]
    segments: list[tuple[str, list[tuple[float, float]]]] = field(default_factory=list)
    closed: bool = False
    radius: float = 0  # roundrect の角の半径（px）
    wrap: bool | None = None  # 図形内の文字を折り返すか。None は種類ごとの既定
    arrow_size: str | None = None  # 矢じりの大きさ。None はテーマの [arrows] size
    origin: int = 0  # 生成元の親要素（同じ g の兄弟かどうかの判定用）

    def bbox(self) -> tuple[float, float, float, float]:
        if self.kind == "line":
            return (
                min(self.x1, self.x2),
                min(self.y1, self.y2),
                max(self.x1, self.x2),
                max(self.y1, self.y2),
            )
        if self.kind == "path":
            pts = [p for _, ps in self.segments for p in ps]
            xs, ys = [p[0] for p in pts], [p[1] for p in pts]
            return min(xs), min(ys), max(xs), max(ys)
        return self.x, self.y, self.x + self.w, self.y + self.h


# ---------------------------------------------------------------------------
# SVG → 中間表現
# ---------------------------------------------------------------------------
def classes(el: ET.Element) -> set[str]:
    return set((el.get("class") or "").split())


def tag(el: ET.Element) -> str:
    return el.tag.replace(SVG_NS, "")


def num(v: str | None, default: float = 0.0) -> float:
    if v is None:
        return default
    m = re.match(r"-?\d*\.?\d+(?:e-?\d+)?", v.strip())
    return float(m.group()) if m else default


def style_map(el: ET.Element) -> dict[str, str]:
    out = {}
    for part in (el.get("style") or "").split(";"):
        if ":" in part:
            k, v = part.split(":", 1)
            out[k.strip()] = v.strip()
    return out


def parse_color(v: str | None) -> str | None:
    if not v or v == "none":
        return None
    v = v.strip()
    if m := re.fullmatch(r"#([0-9a-fA-F]{6})", v):
        return m.group(1).upper()
    if m := re.fullmatch(r"#([0-9a-fA-F]{3})", v):
        return "".join(c * 2 for c in m.group(1)).upper()
    if m := re.fullmatch(r"rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+).*\)", v):
        return "".join(f"{int(c):02X}" for c in m.groups())
    return None


def has_dash(el: ET.Element) -> bool:
    d = style_map(el).get("stroke-dasharray") or el.get("stroke-dasharray") or ""
    return any(num(x) > 0 for x in re.split(r"[,\s]+", d) if x)


def marker_id(v: str | None) -> str | None:
    if not v:
        return None
    m = re.search(r"url\(#(?:[\w-]*?-)?(arrowhead|filled-head|crosshead|sequencenumber|"
                  r"solidTopArrowHead|solidBottomArrowHead|stickTopArrowHead|stickBottomArrowHead)\)", v)
    return m.group(1) if m else None


def theme_color(v: str | None) -> str | None:
    return None if v in (None, "none") else v.upper()


def theme_style(theme: dict, key: str) -> Style:
    t = theme.get(key, {})
    return Style(fill=theme_color(t.get("fill")), line=theme_color(t.get("line")),
                 w=float(t.get("w", 1.0)), dash=t.get("dash"))


def text_lines(el: ET.Element) -> list[str]:
    tspans = [t for t in el if tag(t) == "tspan"]
    if tspans:
        return ["".join(t.itertext()).strip() for t in tspans]
    return ["".join(el.itertext()).strip()]


def font_px(el: ET.Element) -> float:
    s = style_map(el)
    return num(s.get("font-size") or el.get("font-size"), 16)


def is_bold(el: ET.Element) -> bool:
    w = style_map(el).get("font-weight") or el.get("font-weight") or "400"
    return w == "bold" or num(w, 400) >= 600


def text_width_px(s: str, size: float) -> float:
    w = 0.0
    for ch in s:
        w += size if unicodedata.east_asian_width(ch) in "WF" else size * 0.6
    return w


class SvgConverter:
    def __init__(self, theme: dict, mirror_actors: bool = True):
        self.theme = theme
        self.mirror_actors = mirror_actors
        self.shapes: list[Shape] = []
        self.last_message_text = ""
        self.pending_seq: Shape | None = None
        self.counter = 0
        self.text_color = theme_color(theme["text"]["color"]) or "000000"

    def style(self, key: str) -> Style:
        return theme_style(self.theme, key)

    def arrow(self, marker: str | None) -> str | None:
        role = MARKER_TO_ROLE.get(marker or "")
        if role is None:
            return None
        kind = self.theme["arrows"][role]
        return None if kind == "none" else kind

    def name(self, prefix: str, label: str = "") -> str:
        self.counter += 1
        label = re.sub(r"\s+", " ", label).strip()
        return f"{prefix} {self.counter}" + (f": {label[:40]}" if label else "")

    def convert(self, svg_path: Path) -> list[Shape]:
        root = ET.parse(svg_path).getroot()
        if root.get("aria-roledescription") not in (None, "sequence"):
            print(f"warning: {svg_path.name} はシーケンス図ではなさそうです"
                  f"（{root.get('aria-roledescription')}）。対応外の要素は落ちます", file=sys.stderr)
        self.walk(root, inherited_stroke=None)
        return self.shapes

    # -- 走査 ---------------------------------------------------------------
    def walk(self, el: ET.Element, inherited_stroke: str | None) -> None:
        t = tag(el)
        if t in ("defs", "style", "marker", "symbol", "title", "desc"):
            return
        cls = classes(el)
        if not self.mirror_actors and "actor-bottom" in cls:
            return
        if t == "g":
            stroke = parse_color(style_map(el).get("stroke")) or inherited_stroke
            if self.try_container(el):
                return
            for ch in el:
                self.walk(ch, stroke)
            return
        if t == "svg":
            for ch in el:
                self.walk(ch, inherited_stroke)
            return
        handler = getattr(self, f"on_{t}", None)
        if handler:
            handler(el)

    def try_container(self, g: ET.Element) -> bool:
        """参加者の箱・ノートのように「四角 1 つ + 文字」の g は、文字を四角の中に入れる。"""
        kids = list(g)
        rects = [k for k in kids if tag(k) == "rect"]
        texts = [k for k in kids if tag(k) == "text"]
        others = [k for k in kids if tag(k) not in ("rect", "text")]
        if len(rects) != 1 or not texts or others:
            return False
        rc = classes(rects[0])
        if not ({"actor", "note"} & rc):
            return False
        if not self.mirror_actors and "actor-bottom" in rc:
            return True
        shape = self.rect_shape(rects[0])
        lines: list[str] = []
        for tx in texts:
            lines.extend(text_lines(tx))
        shape.text = TextSpec(lines=lines, size_px=font_px(texts[0]), bold=is_bold(texts[0]),
                              color=self.text_color)
        label = " ".join(lines)
        shape.name = self.name("参加者" if "actor" in rc else "ノート", label)
        self.shapes.append(shape)
        return True

    # -- 要素ごと -------------------------------------------------------------
    def rect_shape(self, el: ET.Element) -> Shape:
        cls = classes(el)
        if "actor" in cls:
            st = self.style("actor")
        elif "note" in cls:
            st = self.style("note")
        elif any(c.startswith("activation") for c in cls):
            st = self.style("activation")
        else:  # rect ブロックの背景色など。SVG の属性をそのまま使う
            sm = style_map(el)
            st = Style(fill=parse_color(sm.get("fill") or el.get("fill")),
                       line=parse_color(sm.get("stroke") or el.get("stroke")), w=1.0)
        return Shape("rect", self.name("四角"), x=num(el.get("x")), y=num(el.get("y")),
                     w=num(el.get("width")), h=num(el.get("height")), style=st)

    def on_rect(self, el: ET.Element) -> None:
        s = self.rect_shape(el)
        if any(c.startswith("activation") for c in classes(el)):
            s.name = self.name("活性化")
        self.shapes.append(s)

    def on_circle(self, el: ET.Element) -> None:
        r = num(el.get("r"))
        cx, cy = num(el.get("cx")), num(el.get("cy"))
        self.shapes.append(Shape("ellipse", self.name("頭"), x=cx - r, y=cy - r, w=2 * r, h=2 * r,
                                 style=self.style("actor_man")))

    def on_line(self, el: ET.Element) -> None:
        x1, y1, x2, y2 = (num(el.get(k)) for k in ("x1", "y1", "x2", "y2"))
        cls = classes(el)
        start, end = marker_id(el.get("marker-start")), marker_id(el.get("marker-end"))

        # autonumber: 長さ 0 の線に丸のマーカーが付いている
        if start == "sequencenumber" or end == "sequencenumber":
            sn = self.theme["seq_number"]
            r = sn["r"]
            s = Shape("ellipse", self.name("番号"), x=x1 - r, y=y1 - r, w=2 * r, h=2 * r,
                      style=Style(fill=theme_color(sn["fill"]), line=None))
            self.shapes.append(s)
            self.pending_seq = s
            return
        if x1 == x2 and y1 == y2:
            return

        if "actor-line" in cls:
            st, name = self.style("lifeline"), self.name("ライフライン", el.get("name") or "")
        elif "messageLine1" in cls:
            st, name = self.style("reply"), self.name("応答", self.last_message_text)
        elif "messageLine0" in cls:
            st = self.style("message")
            if has_dash(el):
                st.dash = "sysDash"
            name = self.name("メッセージ", self.last_message_text)
        elif "loopLine" in cls:
            st = self.style("frame_sep" if has_dash(el) else "frame")
            name = self.name("枠線")
        else:
            st, name = self.style("default"), self.name("線")
            if has_dash(el):
                st.dash = "dash"
        s = Shape("line", name, x1=x1, y1=y1, x2=x2, y2=y2, style=st,
                  head=self.arrow(start), tail=self.arrow(end))
        self.shapes.append(s)
        if end == "crosshead":
            self.add_cross(x2, y2, st)

    def add_cross(self, x: float, y: float, st: Style) -> None:
        d = 5
        for (ax, ay, bx, by) in ((x - d, y - d, x + d, y + d), (x - d, y + d, x + d, y - d)):
            self.shapes.append(Shape("line", self.name("×"), x1=ax, y1=ay, x2=bx, y2=by,
                                     style=Style(line=st.line, w=st.w)))

    def on_path(self, el: ET.Element) -> None:
        segs = parse_path(el.get("d") or "")
        if not segs:
            return
        cls = classes(el)
        if "messageLine1" in cls:
            st = self.style("reply")
        elif "messageLine0" in cls:
            st = self.style("message")
        else:
            sm = style_map(el)
            st = Style(fill=parse_color(sm.get("fill") or el.get("fill")),
                       line=parse_color(sm.get("stroke") or el.get("stroke")) or self.style("default").line)
        end = marker_id(el.get("marker-end"))
        closed = any(op == "Z" for op, _ in segs)
        s = Shape("path", self.name("メッセージ" if cls & {"messageLine0", "messageLine1"} else "線",
                                    self.last_message_text), segments=segs, closed=closed, style=st,
                  tail=self.arrow(end))
        if not closed:
            s.style.fill = None
        self.shapes.append(s)

    def on_polygon(self, el: ET.Element) -> None:
        nums = [float(v) for v in re.findall(r"-?\d*\.?\d+", el.get("points") or "")]
        pts = list(zip(nums[0::2], nums[1::2]))
        if len(pts) < 2:
            return
        segs = [("M", [pts[0]])] + [("L", [p]) for p in pts[1:]] + [("Z", [])]
        st = self.style("label_box") if "labelBox" in classes(el) else Style(
            fill=parse_color(el.get("fill")), line=parse_color(el.get("stroke")))
        self.shapes.append(Shape("path", self.name("枠ラベル"), segments=segs, closed=True, style=st))

    def on_text(self, el: ET.Element) -> None:
        cls = classes(el)
        lines = [ln for ln in text_lines(el) if ln]
        if not lines:
            return
        size = font_px(el)
        if "sequenceNumber" in cls and self.pending_seq is not None:
            # 2 桁でも丸に収まる大きさにする
            fit = size * (0.8 if len(lines[0]) < 2 else 0.65)
            self.pending_seq.text = TextSpec(lines=lines, size_px=fit, bold=True,
                                             color=theme_color(self.theme["seq_number"]["text"]) or "FFFFFF")
            self.pending_seq.name = self.name("番号", lines[0])
            self.pending_seq = None
            return
        if "messageText" in cls:
            self.last_message_text = " ".join(lines)

        sm = style_map(el)
        anchor = sm.get("text-anchor") or el.get("text-anchor") or "start"
        baseline = el.get("dominant-baseline") or sm.get("dominant-baseline") or ""
        x, y = num(el.get("x")), num(el.get("y"))
        dy_raw = el.get("dy") or "0"
        dy = num(dy_raw) * size if dy_raw.endswith("em") else num(dy_raw)

        line_h = size * 1.25
        h = line_h * len(lines) + 2
        w = max(text_width_px(ln, size) for ln in lines) + 8
        if baseline in ("central", "middle"):
            cy = y + dy
        else:  # 英字のベースライン基準 → 文字の中央はおよそ 0.35em 上
            cy = y + dy - size * 0.35
        if "actor-man" in cls:  # 人型の名前は足の線に重なりやすいので少し下げる
            cy += size * 0.3
        if anchor == "middle":
            left, align = x - w / 2, "ctr"
        elif anchor == "end":
            left, align = x - w, "r"
        else:
            left, align = x, "l"
        prefix = "メッセージ文字" if "messageText" in cls else "文字"
        self.shapes.append(Shape(
            "textbox", self.name(prefix, " ".join(lines)), x=left, y=cy - h / 2, w=w, h=h,
            text=TextSpec(lines=lines, size_px=size, bold=is_bold(el), color=self.text_color, align=align)))


def parse_path(d: str) -> list[tuple[str, list[tuple[float, float]]]]:
    tokens = re.findall(r"[MLCZmlcz]|-?\d*\.?\d+(?:e-?\d+)?", d)
    segs: list[tuple[str, list[tuple[float, float]]]] = []
    i, op = 0, None
    cur = (0.0, 0.0)
    start = (0.0, 0.0)
    need = {"M": 1, "L": 1, "C": 3}
    while i < len(tokens):
        tk = tokens[i]
        if tk.isalpha():
            op = tk
            i += 1
            if op in "Zz":
                segs.append(("Z", []))
                cur = start
                continue
        if op is None:
            return []
        n = need[op.upper()]
        vals = tokens[i:i + n * 2]
        if len(vals) < n * 2:
            break
        i += n * 2
        pts = [(float(vals[k]), float(vals[k + 1])) for k in range(0, n * 2, 2)]
        if op.islower():
            pts = [(cur[0] + px, cur[1] + py) for px, py in pts]
        segs.append((op.upper(), pts))
        cur = pts[-1]
        if op.upper() == "M":
            start = cur
            op = "l" if op.islower() else "L"  # M の後ろに続く座標は L 扱い
    return segs


# ---------------------------------------------------------------------------
# 汎用 SVG → 中間表現（フローチャート・状態遷移図・ER 図・クラス図）
#
# シーケンス図以外は、図の意味を解釈せずに SVG の見た目をそのまま図形に写す。
#   - 座標: 入れ子の transform を行列で積み上げて絶対座標にする
#   - 色  : Mermaid は色を <style> の CSS で付けるので、簡易な CSS 解決で求める
#   - 矢じり: Excel の矢じりで表せる三角形だけ線の矢じりにし、
#            それ以外（白抜き三角・ひし形・鳥の足など）は記号の形を図形として置く
#   - 文字: ノードの箱・ラベルの背景に 1 つだけ文字があれば、図形の中に入れる
# ---------------------------------------------------------------------------
NUM_RE = r"[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?"
Matrix = tuple[float, float, float, float, float, float]  # a b c d e f（SVG の matrix と同じ）
IDENTITY: Matrix = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)

# Excel の矢じりで表す SVG マーカー（id の末尾）。三角形だけ。
SUPPORTED_GENERIC = {"flowchart-v2", "stateDiagram", "er", "class", "classDiagram"}
GENERIC_ARROW_MARKERS = ("pointEnd", "pointStart", "barbEnd", "arrowhead", "dependencyEnd",
                         "dependencyStart")
INHERITED = {"fill", "stroke", "stroke-width", "stroke-dasharray", "fill-opacity", "stroke-opacity",
             "font-size", "font-weight", "font-family", "text-anchor", "dominant-baseline",
             "visibility", "color", "marker-start", "marker-end"}
PRESENTATION = INHERITED | {"opacity", "display", "alignment-baseline"}


def mat_mul(m: Matrix, n: Matrix) -> Matrix:
    """m・n（n を先に適用する）。"""
    a, b, c, d, e, f = m
    A, B, C, D, E, F = n
    return (a * A + c * B, b * A + d * B, a * C + c * D, b * C + d * D,
            a * E + c * F + e, b * E + d * F + f)


def apply(m: Matrix, x: float, y: float) -> tuple[float, float]:
    a, b, c, d, e, f = m
    return a * x + c * y + e, b * x + d * y + f


def mat_scale(m: Matrix) -> float:
    return math.sqrt(abs(m[0] * m[3] - m[1] * m[2])) or 1.0


def axis_aligned(m: Matrix) -> bool:
    return abs(m[1]) < 1e-9 and abs(m[2]) < 1e-9 and m[0] > 0 and m[3] > 0


def parse_transform(s: str | None) -> Matrix:
    m = IDENTITY
    for name, args in re.findall(r"([a-zA-Z]+)\s*\(([^)]*)\)", s or ""):
        v = [float(x) for x in re.findall(NUM_RE, args)]
        if not v:
            continue
        if name == "translate":
            t = (1.0, 0.0, 0.0, 1.0, v[0], v[1] if len(v) > 1 else 0.0)
        elif name == "scale":
            t = (v[0], 0.0, 0.0, v[1] if len(v) > 1 else v[0], 0.0, 0.0)
        elif name == "rotate":
            r = math.radians(v[0])
            t = (math.cos(r), math.sin(r), -math.sin(r), math.cos(r), 0.0, 0.0)
            if len(v) == 3:
                t = mat_mul(mat_mul((1, 0, 0, 1, v[1], v[2]), t), (1, 0, 0, 1, -v[1], -v[2]))
        elif name == "matrix" and len(v) >= 6:
            t = tuple(v[:6])  # type: ignore[assignment]
        elif name == "skewX":
            t = (1.0, 0.0, math.tan(math.radians(v[0])), 1.0, 0.0, 0.0)
        elif name == "skewY":
            t = (1.0, math.tan(math.radians(v[0])), 0.0, 1.0, 0.0, 0.0)
        else:
            continue
        m = mat_mul(m, t)
    return m


Segs = list[tuple[str, list[tuple[float, float]]]]


def arc_to_cubic(x1, y1, rx, ry, phi, large, sweep, x2, y2) -> list[list[tuple[float, float]]]:
    """SVG の楕円弧を 3 次ベジェ（90 度ごと）に直す。"""
    if rx == 0 or ry == 0 or (x1 == x2 and y1 == y2):
        return [[(x1, y1), (x2, y2), (x2, y2)]]
    rx, ry = abs(rx), abs(ry)
    cp, sp = math.cos(math.radians(phi)), math.sin(math.radians(phi))
    dx, dy = (x1 - x2) / 2, (y1 - y2) / 2
    x1p, y1p = cp * dx + sp * dy, -sp * dx + cp * dy
    lam = (x1p / rx) ** 2 + (y1p / ry) ** 2
    if lam > 1:
        rx, ry = rx * math.sqrt(lam), ry * math.sqrt(lam)
    num = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p
    den = rx * rx * y1p * y1p + ry * ry * x1p * x1p
    coef = math.sqrt(max(0.0, num / den)) if den else 0.0
    if large == sweep:
        coef = -coef
    cxp, cyp = coef * rx * y1p / ry, -coef * ry * x1p / rx
    cx = cp * cxp - sp * cyp + (x1 + x2) / 2
    cy = sp * cxp + cp * cyp + (y1 + y2) / 2

    def ang(ux, uy, vx, vy):
        a = math.atan2(ux * vy - uy * vx, ux * vx + uy * vy)
        return a

    t1 = ang(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry)
    dt = ang((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry)
    if not sweep and dt > 0:
        dt -= 2 * math.pi
    elif sweep and dt < 0:
        dt += 2 * math.pi
    n = max(1, math.ceil(abs(dt) / (math.pi / 2)))
    step = dt / n
    k = 4 / 3 * math.tan(step / 4)
    out = []

    def pt(t):
        x, y = rx * math.cos(t), ry * math.sin(t)
        return cp * x - sp * y + cx, sp * x + cp * y + cy

    def deriv(t):
        x, y = -rx * math.sin(t), ry * math.cos(t)
        return cp * x - sp * y, sp * x + cp * y

    t = t1
    for _ in range(n):
        p0, p3 = pt(t), pt(t + step)
        d0, d3 = deriv(t), deriv(t + step)
        out.append([(p0[0] + k * d0[0], p0[1] + k * d0[1]),
                    (p3[0] - k * d3[0], p3[1] - k * d3[1]), p3])
        t += step
    out[-1][2] = (x2, y2)
    return out


def parse_path_full(d: str) -> Segs:
    """SVG のパスを M / L / C / Z だけの絶対座標に直す（H V S Q T A と相対指定に対応）。"""
    tokens = re.findall(r"[MmLlHhVvCcSsQqTtAaZz]|" + NUM_RE, d or "")
    need = {"M": 2, "L": 2, "H": 1, "V": 1, "C": 6, "S": 4, "Q": 4, "T": 2, "A": 7}
    segs: Segs = []
    i, cmd = 0, None
    cx = cy = sx = sy = 0.0
    last_c2 = last_q = None
    while i < len(tokens):
        tk = tokens[i]
        if tk.isalpha():
            cmd = tk
            i += 1
            if cmd in "Zz":
                segs.append(("Z", []))
                cx, cy = sx, sy
                last_c2 = last_q = None
                continue
        if cmd is None:
            break
        C, rel = cmd.upper(), cmd.islower()
        n = need[C]
        vals = tokens[i:i + n]
        if len(vals) < n or any(v.isalpha() for v in vals):
            break
        v = [float(x) for x in vals]
        i += n
        ox, oy = (cx, cy) if rel else (0.0, 0.0)
        if C == "M":
            cx, cy = v[0] + ox, v[1] + oy
            sx, sy = cx, cy
            segs.append(("M", [(cx, cy)]))
            cmd = "l" if rel else "L"
            last_c2 = last_q = None
        elif C in "LHV":
            if C == "L":
                cx, cy = v[0] + ox, v[1] + oy
            elif C == "H":
                cx = v[0] + (cx if rel else 0.0)
            else:
                cy = v[0] + (cy if rel else 0.0)
            segs.append(("L", [(cx, cy)]))
            last_c2 = last_q = None
        elif C in "CS":
            if C == "C":
                p1 = (v[0] + ox, v[1] + oy)
                p2, p3 = (v[2] + ox, v[3] + oy), (v[4] + ox, v[5] + oy)
            else:
                p1 = (2 * cx - last_c2[0], 2 * cy - last_c2[1]) if last_c2 else (cx, cy)
                p2, p3 = (v[0] + ox, v[1] + oy), (v[2] + ox, v[3] + oy)
            segs.append(("C", [p1, p2, p3]))
            last_c2, last_q = p2, None
            cx, cy = p3
        elif C in "QT":
            if C == "Q":
                q = (v[0] + ox, v[1] + oy)
                p3 = (v[2] + ox, v[3] + oy)
            else:
                q = (2 * cx - last_q[0], 2 * cy - last_q[1]) if last_q else (cx, cy)
                p3 = (v[0] + ox, v[1] + oy)
            p1 = (cx + 2 / 3 * (q[0] - cx), cy + 2 / 3 * (q[1] - cy))
            p2 = (p3[0] + 2 / 3 * (q[0] - p3[0]), p3[1] + 2 / 3 * (q[1] - p3[1]))
            segs.append(("C", [p1, p2, p3]))
            last_q, last_c2 = q, None
            cx, cy = p3
        elif C == "A":
            x2, y2 = v[5] + ox, v[6] + oy
            for c in arc_to_cubic(cx, cy, v[0], v[1], v[2], int(v[3]), int(v[4]), x2, y2):
                segs.append(("C", c))
            cx, cy = x2, y2
            last_c2 = last_q = None
    return segs


def transform_segs(segs: Segs, m: Matrix) -> Segs:
    return [(op, [apply(m, *p) for p in pts]) for op, pts in segs]


# -- CSS -------------------------------------------------------------------
_SIMPLE_RE = re.compile(r"(\*|[a-zA-Z][\w-]*)|#([\w-]+)|\.([\w-]+)|\[([^\]]+)\]"
                        r"|::?([\w-]+)(?:\(([^)]*)\))?")


def _parse_compound(s: str) -> list | None:
    parts, pos = [], 0
    while pos < len(s):
        m = _SIMPLE_RE.match(s, pos)
        if not m:
            return None
        tag_, id_, cls_, attr, pseudo, parg = m.groups()
        if tag_:
            parts.append(("tag", tag_))
        elif id_:
            parts.append(("id", id_))
        elif cls_:
            parts.append(("class", cls_))
        elif attr:
            am = re.match(r"\s*([\w-]+)\s*(?:([~^$*|]?=)\s*['\"]?([^'\"]*)['\"]?)?\s*$", attr)
            if not am:
                return None
            parts.append(("attr", am.groups()))
        elif pseudo:
            if pseudo == "not" and parg:
                inner = _parse_compound(parg.strip())
                if inner is None:
                    return None
                parts.append(("not", inner))
            else:
                return None  # :hover などは静的な見た目に関係しないので、その規則は使わない
        pos = m.end()
    return parts


def parse_selector(sel: str) -> list | None:
    """右から左へ [(compound, combinator)] の列にする。"""
    sel = re.sub(r"\s*([>+~])\s*", r" \1 ", sel.strip())
    toks = sel.split()
    chain, comb = [], " "
    for tk in toks:
        if tk in (">", "+", "~"):
            comb = tk
            continue
        comp = _parse_compound(tk)
        if comp is None:
            return None
        chain.append((comp, comb))
        comb = " "
    if not chain or any(c in "+~" for _, c in chain):
        return None  # 兄弟結合子は使われていないので扱わない
    return list(reversed(chain))


def specificity(chain: list) -> tuple[int, int, int]:
    a = b = c = 0

    def count(comp):
        nonlocal a, b, c
        for kind, val in comp:
            if kind == "id":
                a += 1
            elif kind in ("class", "attr"):
                b += 1
            elif kind == "tag" and val != "*":
                c += 1
            elif kind == "not":
                count(val)

    for comp, _ in chain:
        count(comp)
    return a, b, c


def _match_compound(comp: list, el: ET.Element) -> bool:
    for kind, val in comp:
        if kind == "tag":
            if val != "*" and tag(el) != val:
                return False
        elif kind == "id":
            if el.get("id") != val:
                return False
        elif kind == "class":
            if val not in classes(el):
                return False
        elif kind == "attr":
            name, op, want = val
            have = el.get(name)
            if have is None:
                return False
            if op == "=" and have != want:
                return False
            if op == "$=" and not have.endswith(want):
                return False
            if op == "^=" and not have.startswith(want):
                return False
            if op == "*=" and want not in have:
                return False
            if op == "~=" and want not in have.split():
                return False
        elif kind == "not":
            if _match_compound(val, el):
                return False
    return True


def match_selector(chain: list, el: ET.Element, ancestors: list[ET.Element]) -> bool:
    comp, _ = chain[0]
    if not _match_compound(comp, el):
        return False
    idx = len(ancestors)
    for k in range(1, len(chain)):
        comp, _ = chain[k]
        comb = chain[k - 1][1]
        if comb == ">":
            idx -= 1
            if idx < 0 or not _match_compound(comp, ancestors[idx]):
                return False
        else:
            idx -= 1
            while idx >= 0 and not _match_compound(comp, ancestors[idx]):
                idx -= 1
            if idx < 0:
                return False
    return True


def parse_css(text: str) -> list[tuple[list, tuple[int, int, int], int, dict[str, tuple[str, bool]]]]:
    text = re.sub(r"/\*.*?\*/", "", text, flags=re.S)
    rules = []
    order = 0
    for sel_text, body in re.findall(r"([^{}]+)\{([^{}]*)\}", text):
        decls = {}
        for part in body.split(";"):
            if ":" not in part:
                continue
            k, v = part.split(":", 1)
            v = v.strip()
            imp = "!important" in v
            decls[k.strip().lower()] = (v.replace("!important", "").strip(), imp)
        if not decls:
            continue
        for sel in re.split(r",(?![^(]*\))", sel_text):
            if sel.strip().startswith("@"):
                continue
            chain = parse_selector(sel)
            if chain is None:
                continue
            order += 1
            rules.append((chain, specificity(chain), order, decls))
    return rules


def css_color(v: str | None, current: str | None = None) -> tuple[str | None, float]:
    """CSS の色 → (RRGGBB, alpha)。塗りなしは (None, 0)。"""
    if v is None:
        return None, 0.0
    v = v.strip().lower()
    if v in ("none", "transparent", ""):
        return None, 0.0
    if v == "currentcolor":
        return css_color(current) if current else (None, 0.0)
    if m := re.fullmatch(r"#([0-9a-f]{6})([0-9a-f]{2})?", v):
        return m.group(1).upper(), (int(m.group(2), 16) / 255 if m.group(2) else 1.0)
    if m := re.fullmatch(r"#([0-9a-f]{3})", v):
        return "".join(c * 2 for c in m.group(1)).upper(), 1.0
    if m := re.fullmatch(r"rgba?\(([^)]*)\)", v):
        p = [x.strip() for x in re.split(r"[,\s/]+", m.group(1)) if x.strip()]
        rgb = [round(float(x[:-1]) * 2.55) if x.endswith("%") else round(float(x)) for x in p[:3]]
        a = float(p[3][:-1]) / 100 if len(p) > 3 and p[3].endswith("%") else (float(p[3]) if len(p) > 3 else 1)
        return "".join(f"{max(0, min(255, c)):02X}" for c in rgb), a
    if m := re.fullmatch(r"hsla?\(([^)]*)\)", v):
        p = [x.strip() for x in re.split(r"[,\s/]+", m.group(1)) if x.strip()]
        h = float(re.sub(r"deg$", "", p[0])) / 360
        s, l = float(p[1].rstrip("%")) / 100, float(p[2].rstrip("%")) / 100
        import colorsys
        r, g, b = colorsys.hls_to_rgb(h % 1, l, s)
        a = float(p[3].rstrip("%")) / (100 if p[3].endswith("%") else 1) if len(p) > 3 else 1
        return f"{round(r * 255):02X}{round(g * 255):02X}{round(b * 255):02X}", a
    named = {"black": "000000", "white": "FFFFFF", "red": "FF0000", "green": "008000",
             "blue": "0000FF", "gray": "808080", "grey": "808080", "yellow": "FFFF00",
             "orange": "FFA500", "lightgrey": "D3D3D3", "lightgray": "D3D3D3"}
    if v in named:
        return named[v], 1.0
    return None, 0.0


def css_px(v: str | None, font_px_: float = 16.0, default: float = 0.0) -> float:
    if v is None:
        return default
    v = v.strip()
    m = re.match(NUM_RE, v)
    if not m:
        return default
    n = float(m.group())
    unit = v[m.end():].strip()
    if unit == "em":
        return n * font_px_
    if unit == "pt":
        return n * 4 / 3
    return n


def dash_preset(v: str | None, width_px: float) -> str | None:
    if not v or v.strip() in ("none", "0", ""):
        return None
    nums = [float(x) for x in re.findall(NUM_RE, v)]
    if not nums or all(n == 0 for n in nums):
        return None
    if len(nums) >= 2 and nums[1] == 0:
        return None  # "1, 0" のように隙間 0 は実線
    seg = max(nums) / max(width_px, 0.5)
    return "sysDot" if seg <= 2.5 else ("sysDash" if seg <= 4.5 else "dash")


class GenericConverter:
    def __init__(self, theme: dict):
        self.theme = theme
        self.shapes: list[Shape] = []
        self.counter = 0

    def convert(self, svg_path: Path) -> list[Shape]:
        root = ET.parse(svg_path).getroot()
        css = "".join("".join(st.itertext()) for st in root.iter(SVG_NS + "style"))
        self.rules = parse_css(css)
        self.ids = {el.get("id"): el for el in root.iter() if el.get("id")}
        self.parent = {c: p for p in root.iter() for c in p}
        self.walk(root, IDENTITY, {"fill": "black", "font-size": "16px"}, [], 1.0)
        return self.shapes

    # -- 名前 ---------------------------------------------------------------
    def name(self, ancestors: list[ET.Element], el: ET.Element, label: str = "") -> str:
        prefix = "図形"
        for a in reversed([*ancestors, el]):
            c = classes(a)
            if tag(a) == "marker":
                prefix = "記号"
            elif "edgeLabel" in c or "edgeTerminals" in c:
                prefix = "ラベル"
            elif c & {"node", "statediagram-state"}:
                prefix = "ノード"
            elif c & {"cluster", "clusters"}:
                prefix = "グループ"
            elif "edgePaths" in c or "edgePath" in c or el.get("data-et") == "edge":
                prefix = "線"
            else:
                continue
            break
        if tag(el) == "text" and prefix in ("図形",):
            prefix = "文字"
        self.counter += 1
        label = re.sub(r"\s+", " ", label).strip()
        return f"{prefix} {self.counter}" + (f": {label[:40]}" if label else "")

    # -- スタイル -----------------------------------------------------------
    def computed(self, el: ET.Element, parent_style: dict, ancestors: list[ET.Element]) -> dict:
        st = {k: v for k, v in parent_style.items() if k in INHERITED}
        for k in PRESENTATION:
            if el.get(k) is not None:
                st[k] = el.get(k)
        normal, important = [], []
        for chain, spec, order, decls in self.rules:
            if match_selector(chain, el, ancestors):
                for k, (v, imp) in decls.items():
                    (important if imp else normal).append((spec, order, k, v))
        for _, _, k, v in sorted(normal):
            st[k] = v
        for k, v in style_map(el).items():
            if v and v != "undefined":
                st[k.lower()] = v
        for _, _, k, v in sorted(important):
            st[k] = v
        if "font-size" in st:  # em は親の文字の大きさが基準
            st["font-size"] = f"{css_px(st['font-size'], css_px(parent_style.get('font-size'), 16, 16), 16)}px"
        return st

    def shape_style(self, st: dict, m: Matrix, opacity: float, closed: bool) -> Style:
        fill, fa = css_color(st.get("fill"), st.get("color"))
        fa *= float(st.get("fill-opacity") or 1) * opacity
        line, la = css_color(st.get("stroke"), st.get("color"))
        la *= float(st.get("stroke-opacity") or 1) * opacity
        width_px = css_px(st.get("stroke-width"), default=1.0)
        if width_px <= 0:
            line = None
        if not closed:
            fill = None
        sw = width_px * mat_scale(m)
        return Style(fill=fill if fa > 0.001 else None, line=line if la > 0.001 else None,
                     w=sw * 0.75, dash=dash_preset(st.get("stroke-dasharray"), width_px),
                     fill_alpha=min(fa, 1.0), line_alpha=min(la, 1.0))

    # -- 走査 ---------------------------------------------------------------
    def walk(self, el: ET.Element, m: Matrix, parent_style: dict, ancestors: list[ET.Element],
             opacity: float) -> None:
        t = tag(el)
        if t in ("defs", "style", "marker", "symbol", "title", "desc", "foreignObject",
                 "clipPath", "mask", "linearGradient", "radialGradient", "pattern"):
            return
        st = self.computed(el, parent_style, ancestors)
        if st.get("display") == "none" or st.get("visibility") == "hidden":
            return
        op = opacity * float(st.get("opacity") or 1)
        m2 = mat_mul(m, parse_transform(el.get("transform")))
        if t in ("g", "svg", "a", "switch"):
            start = len(self.shapes)
            for ch in el:
                self.walk(ch, m2, st, [*ancestors, el], op)
            self.post_group(el, start)
            return
        handler = getattr(self, f"g_{t}", None)
        if handler:
            handler(el, m2, st, ancestors, op)

    def add(self, s: Shape, ancestors: list[ET.Element]) -> Shape:
        s.origin = id(ancestors[-1]) if ancestors else 0
        s.wrap = False
        self.shapes.append(s)
        return s

    def g_rect(self, el, m, st, anc, op):
        x, y = css_px(el.get("x")), css_px(el.get("y"))
        w, h = css_px(el.get("width")), css_px(el.get("height"))
        if w <= 0 or h <= 0:
            return
        rx = css_px(el.get("rx") or el.get("ry"))
        style = self.shape_style(st, m, op, True)
        if style.fill is None and style.line is None:
            return
        if axis_aligned(m):
            (x0, y0), (x1, y1) = apply(m, x, y), apply(m, x + w, y + h)
            kind = "roundrect" if rx > 0 else "rect"
            self.add(Shape(kind, self.name(anc, el), x=x0, y=y0, w=x1 - x0, h=y1 - y0, style=style,
                           radius=rx * mat_scale(m)), anc)
        else:
            pts = [(x, y), (x + w, y), (x + w, y + h), (x, y + h)]
            segs = [("M", [pts[0]])] + [("L", [p]) for p in pts[1:]] + [("Z", [])]
            self.add(Shape("path", self.name(anc, el), segments=transform_segs(segs, m), closed=True,
                           style=style), anc)

    def g_circle(self, el, m, st, anc, op):
        r = css_px(el.get("r"))
        self._ellipse(el, m, st, anc, op, css_px(el.get("cx")), css_px(el.get("cy")), r, r)

    def g_ellipse(self, el, m, st, anc, op):
        self._ellipse(el, m, st, anc, op, css_px(el.get("cx")), css_px(el.get("cy")),
                      css_px(el.get("rx")), css_px(el.get("ry")))

    def _ellipse(self, el, m, st, anc, op, cx, cy, rx, ry):
        if rx <= 0 or ry <= 0:
            return
        style = self.shape_style(st, m, op, True)
        if style.fill is None and style.line is None:
            return
        if axis_aligned(m):
            (x0, y0), (x1, y1) = apply(m, cx - rx, cy - ry), apply(m, cx + rx, cy + ry)
            self.add(Shape("ellipse", self.name(anc, el), x=x0, y=y0, w=x1 - x0, h=y1 - y0,
                           style=style), anc)
        else:
            d = (f"M{cx - rx},{cy} A{rx},{ry} 0 1 0 {cx + rx},{cy} A{rx},{ry} 0 1 0 {cx - rx},{cy} Z")
            self.add(Shape("path", self.name(anc, el), segments=transform_segs(parse_path_full(d), m),
                           closed=True, style=style), anc)

    def g_line(self, el, m, st, anc, op):
        p1 = apply(m, css_px(el.get("x1")), css_px(el.get("y1")))
        p2 = apply(m, css_px(el.get("x2")), css_px(el.get("y2")))
        if p1 == p2:
            return
        style = self.shape_style(st, m, op, False)
        if style.line is None:
            return
        s = self.add(Shape("line", self.name(anc, el), x1=p1[0], y1=p1[1], x2=p2[0], y2=p2[1],
                           style=style), anc)
        self.markers(s, el, st, anc, m)

    def g_polygon(self, el, m, st, anc, op, closed=True):
        nums = [float(v) for v in re.findall(NUM_RE, el.get("points") or "")]
        pts = list(zip(nums[0::2], nums[1::2]))
        if len(pts) < 2:
            return
        segs: Segs = [("M", [pts[0]])] + [("L", [p]) for p in pts[1:]]
        if closed:
            segs.append(("Z", []))
        self._path(el, transform_segs(segs, m), m, st, anc, op, closed)

    def g_polyline(self, el, m, st, anc, op):
        self.g_polygon(el, m, st, anc, op, closed=False)

    def g_path(self, el, m, st, anc, op):
        segs = parse_path_full(el.get("d") or "")
        if not any(op_ in "LC" for op_, _ in segs):
            return
        # SVG は閉じていないパスも塗る（端どうしを結んで塗る）。Excel も同じなので、塗りがあれば図形扱い
        closed = any(op_ == "Z" for op_, _ in segs) or css_color(st.get("fill"))[0] is not None
        self._path(el, transform_segs(segs, m), m, st, anc, op, closed)

    def _path(self, el, segs, m, st, anc, op, closed):
        style = self.shape_style(st, m, op, closed)
        if style.fill is None and style.line is None:
            return
        s = self.add(Shape("path", self.name(anc, el), segments=segs, closed=closed, style=style), anc)
        self.markers(s, el, st, anc, m)

    # -- マーカー（矢じり・多重度の記号など）--------------------------------
    def markers(self, s: Shape, el: ET.Element, st: dict, anc: list[ET.Element], m: Matrix) -> None:
        for which in ("start", "end"):
            ref = st.get(f"marker-{which}") or el.get(f"marker-{which}")
            mm = re.search(r"url\(#([^)]+)\)", ref or "")
            if not mm or mm.group(1) not in self.ids:
                continue
            mk = self.ids[mm.group(1)]
            vertex, angle = self.end_direction(s, which)
            if vertex is None:
                continue
            # 線の外側へ向かう向き（終点は進む向き、始点は逆向き）
            outward = angle if which == "end" else angle + math.pi
            base = re.sub(r"-margin$", "", mk.get("id") or "")
            orient = mk.get("orient") or "0"
            if orient == "auto-start-reverse" and which == "start":
                angle += math.pi
            elif orient not in ("auto", "auto-start-reverse"):
                angle = math.radians(num(orient))
            mk_m = self.marker_matrix(mk, vertex, angle, st, m)
            if any(base.endswith(k) for k in GENERIC_ARROW_MARKERS) and s.style.line:
                self.as_arrowhead(s, mk, mk_m, which, vertex, outward)
            else:
                mk_anc = self.ancestors_of(mk)
                mst = self.computed(mk, {"fill": "black", "font-size": "16px"}, mk_anc)
                for ch in mk:
                    self.walk(ch, mk_m, mst, [*mk_anc, mk], 1.0)

    def ancestors_of(self, el: ET.Element) -> list[ET.Element]:
        out = []
        while el in self.parent:
            el = self.parent[el]
            out.append(el)
        return list(reversed(out))

    def marker_matrix(self, mk: ET.Element, vertex, angle: float, st: dict, m: Matrix) -> Matrix:
        sw = 1.0
        if (mk.get("markerUnits") or "strokeWidth") == "strokeWidth":
            sw = css_px(st.get("stroke-width"), default=1.0)
        vb = [float(v) for v in re.findall(NUM_RE, mk.get("viewBox") or "")]
        sx = sy = 1.0
        if len(vb) == 4 and vb[2] > 0 and vb[3] > 0:
            sx = css_px(mk.get("markerWidth"), default=3) / vb[2]
            sy = css_px(mk.get("markerHeight"), default=3) / vb[3]
            sx = sy = min(sx, sy)
        k = sw * mat_scale(m)
        ref_x, ref_y = num(mk.get("refX")), num(mk.get("refY"))
        mt: Matrix = (1, 0, 0, 1, vertex[0], vertex[1])
        mt = mat_mul(mt, (math.cos(angle), math.sin(angle), -math.sin(angle), math.cos(angle), 0, 0))
        mt = mat_mul(mt, (k * sx, 0, 0, k * sy, 0, 0))
        return mat_mul(mt, (1, 0, 0, 1, -ref_x, -ref_y))

    def as_arrowhead(self, s: Shape, mk: ET.Element, mk_m: Matrix, which: str,
                     vertex: tuple[float, float], outward: float) -> None:
        """三角形の記号は Excel の矢じりにする。線の端を記号の先端まで伸ばす。"""
        pts = []
        for ch in mk.iter():
            if tag(ch) == "path":
                pts += [p for _, ps in parse_path_full(ch.get("d") or "") for p in ps]
            elif tag(ch) == "polygon":
                n = [float(v) for v in re.findall(NUM_RE, ch.get("points") or "")]
                pts += list(zip(n[0::2], n[1::2]))
        if pts:
            ux, uy = math.cos(outward), math.sin(outward)
            # 線の外側へ一番突き出た点が先端。線の端をそこまで伸ばす（Excel の矢じりは端が先端）
            reach = max((q[0] - vertex[0]) * ux + (q[1] - vertex[1]) * uy
                        for q in (apply(mk_m, *p) for p in pts))
            if reach > 0:
                self.move_end(s, which, (vertex[0] + ux * reach, vertex[1] + uy * reach))
        if which == "start":
            s.head = "triangle"
        else:
            s.tail = "triangle"
        # Excel の矢じりは線の太さに比例する。細い線では大きめにして Mermaid の見た目に寄せる
        s.arrow_size = "lg" if s.style.w < 1.0 else ("med" if s.style.w < 2.0 else "sm")

    @staticmethod
    def end_direction(s: Shape, which: str):
        if s.kind == "line":
            if which == "start":
                return (s.x1, s.y1), math.atan2(s.y1 - s.y2, s.x1 - s.x2) + math.pi
            return (s.x2, s.y2), math.atan2(s.y2 - s.y1, s.x2 - s.x1)
        pts = [p for op_, ps in s.segments if op_ != "Z" for p in ps]
        if len(pts) < 2:
            return None, 0.0
        if which == "start":
            v = pts[0]
            nxt = next((p for p in pts[1:] if math.dist(p, v) > 1e-6), None)
            return (v, math.atan2(nxt[1] - v[1], nxt[0] - v[0])) if nxt else (None, 0.0)
        v = pts[-1]
        prv = next((p for p in reversed(pts[:-1]) if math.dist(p, v) > 1e-6), None)
        return (v, math.atan2(v[1] - prv[1], v[0] - prv[0])) if prv else (None, 0.0)

    @staticmethod
    def move_end(s: Shape, which: str, p: tuple[float, float]) -> None:
        if s.kind == "line":
            if which == "start":
                s.x1, s.y1 = p
            else:
                s.x2, s.y2 = p
            return
        idx = 0 if which == "start" else max(i for i, (op_, _) in enumerate(s.segments) if op_ != "Z")
        op_, pts = s.segments[idx]
        pts = list(pts)
        pts[0 if which == "start" else -1] = p
        s.segments[idx] = (op_, pts)

    # -- 文字 ---------------------------------------------------------------
    def g_text(self, el, m, st, anc, op):
        fs = css_px(st.get("font-size"), default=16) or 16
        rows: list[tuple[str, float, float]] = []
        x = css_px(el.get("x"), fs)
        y = css_px(el.get("y"), fs) + css_px(el.get("dy"), fs)
        if el.text and el.text.strip():
            rows.append((el.text.strip(), x, y))
        for ts in el:
            if tag(ts) != "tspan":
                continue
            if ts.get("x") is not None:
                x = css_px(ts.get("x"), fs)
            if ts.get("y") is not None:
                y = css_px(ts.get("y"), fs)
            y += css_px(ts.get("dy"), fs)
            txt = "".join(ts.itertext()).strip()
            if txt:
                rows.append((txt, x, y))
        if not rows:
            return
        anchor = st.get("text-anchor") or "start"
        base = st.get("dominant-baseline") or st.get("alignment-baseline") or ""
        centered = base in ("central", "middle")
        lines = [r[0] for r in rows]
        cy = sum((r[2] if centered else r[2] - fs * 0.35) for r in rows) / len(rows)
        x0 = rows[0][1]
        w = max(text_width_px(ln, fs) for ln in lines) + 8
        h = fs * 1.2 * len(lines) + 2
        cx = x0 if anchor == "middle" else (x0 + w / 2 - 4 if anchor == "start" else x0 - w / 2 + 4)
        k = mat_scale(m)
        pcx, pcy = apply(m, cx, cy)
        color, _ = css_color(st.get("fill"), st.get("color"))
        align = {"middle": "ctr", "end": "r"}.get(anchor, "l")
        weight = st.get("font-weight") or "400"
        bold = weight == "bold" or num(weight, 400) >= 600
        self.add(Shape("textbox", self.name(anc, el, " ".join(lines)), x=pcx - w * k / 2,
                       y=pcy - h * k / 2, w=w * k, h=h * k,
                       text=TextSpec(lines=lines, size_px=fs * k, bold=bold, color=color or "333333",
                                     align=align)), anc)

    # -- グループ単位の後処理 ------------------------------------------------
    def post_group(self, g: ET.Element, start: int) -> None:
        mine = self.shapes[start:]
        if not mine:
            return
        # 1) Mermaid の「塗り用パス + 線用パス」の組を 1 つの図形にまとめる
        merged: list[Shape] = []
        for s in mine:
            prev = merged[-1] if merged else None
            # 同じ g の中で外形がほぼ同じ 2 本のパスは、塗り用と線用の組。CSS の当たり方で
            # どちらにも塗りと線が付くことがあるので、外形で判定する
            if (prev is not None and s.origin == prev.origin == id(g) and s.kind == prev.kind == "path"
                    and prev.text is None and not (s.head or s.tail or prev.head or prev.tail)
                    and all(abs(u - v) < 2 for u, v in zip(prev.bbox(), s.bbox()))):
                prev.closed = prev.closed or s.closed
                prev.style.fill = prev.style.fill or s.style.fill
                if s.style.line:
                    prev.style.line, prev.style.w = s.style.line, s.style.w
                    prev.style.dash, prev.style.line_alpha = s.style.dash, s.style.line_alpha
                continue
            merged.append(s)
        # 2) ノード・ラベルの文字を図形の中に入れる
        c = classes(g)
        if c & {"node", "edgeLabel"} or ("label" in c and "edgeLabel" in classes(self.parent.get(g, g))):
            texts = [s for s in merged if s.kind == "textbox"]
            bodies = [s for s in merged if s.kind != "textbox" and s.text is None
                      and s.kind != "line" and (s.closed or s.kind != "path")]
            if len(bodies) == 1 and texts:
                body = bodies[0]
                bx0, by0, bx1, by1 = body.bbox()
                inside = all(bx0 - 2 <= (t.x + t.w / 2) <= bx1 + 2 and by0 - 2 <= (t.y + t.h / 2) <= by1 + 2
                             for t in texts)
                if inside:
                    first = texts[0].text
                    body.text = TextSpec(lines=[ln for t in texts for ln in t.text.lines],
                                         size_px=first.size_px, bold=first.bold, color=first.color,
                                         align="ctr")
                    body.name = re.sub(r":.*$", "", body.name) + ": " + " ".join(body.text.lines)[:40]
                    merged = [s for s in merged if s not in texts]
        self.shapes[start:] = merged


# ---------------------------------------------------------------------------
# 中間表現 → DrawingML
# ---------------------------------------------------------------------------
class DrawingWriter:
    def __init__(self, theme: dict, scale: float):
        self.font = theme["text"]["font"]
        self.arrow_size = theme["arrows"]["size"]
        self.margin = float(theme["layout"]["margin"])
        self.scale = scale
        self.next_id = 2

    def emu(self, px: float) -> int:
        return int(round(px * EMU_PER_PX * self.scale))

    def nid(self) -> int:
        self.next_id += 1
        return self.next_id

    @staticmethod
    def color(val: str, alpha: float = 1.0) -> str:
        a = f'<a:alpha val="{int(round(alpha * 100000))}"/>' if alpha < 0.999 else ""
        return f'<a:solidFill><a:srgbClr val="{val}">{a}</a:srgbClr></a:solidFill>' if a else \
            f'<a:solidFill><a:srgbClr val="{val}"/></a:solidFill>'

    def ln(self, st: Style, head: str | None = None, tail: str | None = None,
           size: str | None = None) -> str:
        if st.line is None:
            return "<a:ln><a:noFill/></a:ln>"
        dash = f'<a:prstDash val="{st.dash}"/>' if st.dash else ""
        sz = size or self.arrow_size
        he = f'<a:headEnd type="{head}" w="{sz}" len="{sz}"/>' if head else ""
        te = f'<a:tailEnd type="{tail}" w="{sz}" len="{sz}"/>' if tail else ""
        return (f'<a:ln w="{int(st.w * EMU_PER_PT * self.scale)}">'
                f'{self.color(st.line, st.line_alpha)}{dash}'
                f'<a:round/>{he}{te}</a:ln>')

    def fill(self, st: Style) -> str:
        return self.color(st.fill, st.fill_alpha) if st.fill else "<a:noFill/>"

    def txbody(self, t: TextSpec | None, wrap: bool) -> str:
        if t is None:
            return ""
        sz = max(100, int(round(t.size_px * 0.75 * self.scale * 100)))
        b = ' b="1"' if t.bold else ""
        font = escape(self.font, {'"': "&quot;"})
        paras = []
        for line in t.lines:
            run = (f'<a:r><a:rPr lang="ja-JP" altLang="en-US" sz="{sz}"{b}>'
                   f'<a:solidFill><a:srgbClr val="{t.color}"/></a:solidFill>'
                   f'<a:latin typeface="{font}"/><a:ea typeface="{font}"/><a:cs typeface="{font}"/>'
                   f'</a:rPr><a:t>{escape(line)}</a:t></a:r>') if line else ""
            paras.append(f'<a:p><a:pPr algn="{t.align}"/>{run}'
                         f'<a:endParaRPr lang="ja-JP" sz="{sz}"/></a:p>')
        wrap_attr = "square" if wrap else "none"
        ins = 36000 if wrap else 0
        return (f'<xdr:txBody><a:bodyPr vertOverflow="overflow" horzOverflow="overflow" '
                f'wrap="{wrap_attr}" lIns="{ins}" tIns="0" rIns="{ins}" bIns="0" anchor="ctr" '
                f'rtlCol="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>{"".join(paras)}</xdr:txBody>')

    def xfrm(self, x: float, y: float, w: float, h: float, flip_h=False, flip_v=False) -> str:
        fh = ' flipH="1"' if flip_h else ""
        fv = ' flipV="1"' if flip_v else ""
        return (f'<a:xfrm{fh}{fv}><a:off x="{self.emu(x)}" y="{self.emu(y)}"/>'
                f'<a:ext cx="{max(self.emu(w), 1)}" cy="{max(self.emu(h), 1)}"/></a:xfrm>')

    def shape(self, s: Shape, ox: float, oy: float) -> str:
        name = escape(s.name, {'"': "&quot;"})
        if s.kind == "line":
            x1, y1, x2, y2 = s.x1 - ox, s.y1 - oy, s.x2 - ox, s.y2 - oy
            return (f'<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="{self.nid()}" name="{name}"/>'
                    f'<xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr>'
                    f'{self.xfrm(min(x1, x2), min(y1, y2), abs(x2 - x1), abs(y2 - y1), x2 < x1, y2 < y1)}'
                    f'<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>'
                    f'{self.ln(s.style, s.head, s.tail, s.arrow_size)}</xdr:spPr></xdr:cxnSp>')

        if s.kind == "path":
            bx0, by0, bx1, by1 = s.bbox()
            geom = self.custgeom(s, bx0, by0, bx1 - bx0, by1 - by0)
            x, y, w, h = bx0 - ox, by0 - oy, bx1 - bx0, by1 - by0
        else:
            prst = {"rect": "rect", "roundrect": "roundRect", "ellipse": "ellipse", "textbox": "rect"}[s.kind]
            av = "<a:avLst/>"
            if s.kind == "roundrect":  # 角の半径は短い辺に対する割合で指定する
                adj = int(min(50000, s.radius / max(min(s.w, s.h), 1e-6) * 100000))
                av = f'<a:avLst><a:gd name="adj" fmla="val {adj}"/></a:avLst>'
            geom = f'<a:prstGeom prst="{prst}">{av}</a:prstGeom>'
            x, y, w, h = s.x - ox, s.y - oy, s.w, s.h

        is_tb = s.kind == "textbox"
        style = Style() if is_tb else s.style
        tb_attr = ' txBox="1"' if is_tb else ""
        return (f'<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{self.nid()}" name="{name}"/>'
                f'<xdr:cNvSpPr{tb_attr}/></xdr:nvSpPr><xdr:spPr>{self.xfrm(x, y, w, h)}{geom}'
                f'{self.fill(style)}{self.ln(style, s.head, s.tail, s.arrow_size)}</xdr:spPr>'
                f'{self.txbody(s.text, wrap=s.wrap if s.wrap is not None else (not is_tb and s.kind != "ellipse"))}'
                '</xdr:sp>')

    def custgeom(self, s: Shape, bx: float, by: float, bw: float, bh: float) -> str:
        W, H = max(self.emu(bw), 1), max(self.emu(bh), 1)

        def pt(p):
            return f'<a:pt x="{self.emu(p[0] - bx)}" y="{self.emu(p[1] - by)}"/>'

        parts = []
        for op, pts in s.segments:
            if op == "M":
                parts.append(f"<a:moveTo>{pt(pts[0])}</a:moveTo>")
            elif op == "L":
                parts.append(f"<a:lnTo>{pt(pts[0])}</a:lnTo>")
            elif op == "C":
                parts.append(f"<a:cubicBezTo>{''.join(pt(p) for p in pts)}</a:cubicBezTo>")
            elif op == "Z":
                parts.append("<a:close/>")
        fill_attr = "" if s.closed else ' fill="none"'
        return ('<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>'
                '<a:rect l="0" t="0" r="r" b="b"/><a:pathLst>'
                f'<a:path w="{W}" h="{H}"{fill_attr}>{"".join(parts)}</a:path></a:pathLst></a:custGeom>')

    def drawing(self, shapes: list[Shape], title: str) -> str:
        boxes = [s.bbox() for s in shapes]
        m = self.margin
        ox = min(b[0] for b in boxes) - m
        oy = min(b[1] for b in boxes) - m
        W = max(b[2] for b in boxes) - ox + m
        H = max(b[3] for b in boxes) - oy + m
        cx, cy = self.emu(W), self.emu(H)
        body = "".join(self.shape(s, ox, oy) for s in shapes)
        title = escape(title, {'"': "&quot;"})
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff>'
            '<xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
            f'<xdr:ext cx="{cx}" cy="{cy}"/>'
            f'<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="2" name="{title}"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>'
            f'<xdr:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/>'
            f'<a:chOff x="0" y="0"/><a:chExt cx="{cx}" cy="{cy}"/></a:xfrm></xdr:grpSpPr>'
            f'{body}</xdr:grpSp><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>'
        )


# ---------------------------------------------------------------------------
# .xlsx パッケージ
# ---------------------------------------------------------------------------
def sheet_name(stem: str, used: set[str]) -> str:
    base = re.sub(r"[\[\]:*?/\\]", "_", stem)[:31] or "Sheet"
    name, n = base, 2
    while name in used:
        suffix = f"_{n}"
        name, n = base[: 31 - len(suffix)] + suffix, n + 1
    used.add(name)
    return name


def write_xlsx(out: Path, sheets: list[tuple[str, str]], font: str) -> None:
    ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" '
          'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          '<Override PartName="/xl/styles.xml" '
          'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>']
    wb_sheets, wb_rels = [], []
    files: dict[str, str] = {}
    for i, (name, drawing) in enumerate(sheets, 1):
        ct.append(f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
                  'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
        ct.append(f'<Override PartName="/xl/drawings/drawing{i}.xml" '
                  'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>')
        selected = ' tabSelected="1"' if i == 1 else ""
        name_x = escape(name, {'"': "&quot;"})
        wb_sheets.append(f'<sheet name="{name_x}" sheetId="{i}" r:id="rId{i}"/>')
        wb_rels.append(f'<Relationship Id="rId{i}" '
                       'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                       f'Target="worksheets/sheet{i}.xml"/>')
        files[f"xl/worksheets/sheet{i}.xml"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'
            f'<sheetViews><sheetView showGridLines="0" workbookViewId="0"{selected}/>'
            '</sheetViews><sheetData/>'
            '<pageMargins left="0.4" right="0.4" top="0.4" bottom="0.4" header="0.2" footer="0.2"/>'
            '<pageSetup paperSize="9" fitToWidth="1" fitToHeight="0"/>'  # A4・横 1 ページに収める
            '<drawing r:id="rId1"/></worksheet>')
        files[f"xl/worksheets/_rels/sheet{i}.xml.rels"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" '
            f'Target="../drawings/drawing{i}.xml"/></Relationships>')
        files[f"xl/drawings/drawing{i}.xml"] = drawing
    ct.append("</Types>")
    n = len(sheets)
    wb_rels.append(f'<Relationship Id="rId{n + 1}" '
                   'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
                   'Target="styles.xml"/>')
    font_x = escape(font, {'"': "&quot;"})
    files.update({
        "[Content_Types].xml": "".join(ct),
        "_rels/.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="xl/workbook.xml"/></Relationships>'),
        "xl/workbook.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<bookViews><workbookView/></bookViews><sheets>{"".join(wb_sheets)}</sheets></workbook>'),
        "xl/_rels/workbook.xml.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'{"".join(wb_rels)}</Relationships>'),
        "xl/styles.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<fonts count="1"><font><sz val="11"/><name val="{font_x}"/><family val="3"/>'
            '<charset val="128"/></font></fonts>'
            '<fills count="2"><fill><patternFill patternType="none"/></fill>'
            '<fill><patternFill patternType="gray125"/></fill></fills>'
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            '</styleSheet>'),
    })
    for path, xml in files.items():  # 壊れた XML を Excel に渡さない
        ET.fromstring(xml.encode("utf-8"))
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", files.pop("[Content_Types].xml"))
        for path, xml in files.items():
            z.writestr(path, xml)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def render_svg(mmd: Path, out_svg: Path, mmdc: str | None, mermaid_theme: str) -> None:
    # 文字を SVG の <text> で出させる（HTML ラベルのままだと文字の位置が取れない）
    config = {"theme": mermaid_theme, "htmlLabels": False,
              **{k: {"htmlLabels": False} for k in ("flowchart", "class", "state", "er")}}
    cfg = out_svg.with_suffix(".mermaid-config.json")
    cfg.write_text(json.dumps(config), encoding="utf-8")
    if mmdc:
        cmd = mmdc.split()
    elif shutil.which("mmdc"):
        cmd = ["mmdc"]
    else:
        cmd = ["npx", "-y", MMDC_PACKAGE]
    cmd = [*cmd, "-c", str(cfg), "-i", str(mmd), "-o", str(out_svg), "-q"]
    r = subprocess.run(cmd, capture_output=True, text=True)
    cfg.unlink(missing_ok=True)
    if r.returncode != 0 or not out_svg.exists():
        sys.exit(f"mmdc failed for {mmd}:\n{r.stderr or r.stdout}")


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Mermaid の図を Excel の図形に変換する")
    p.add_argument("inputs", nargs="+", type=Path, help=".mmd または mmdc で書き出した .svg")
    p.add_argument("-o", "--output", type=Path, help="出力する .xlsx（省略時は最初の入力名.xlsx）")
    p.add_argument("--theme", type=Path,
                   help="見た目の設定（TOML）。書いた項目だけ既定の theme.toml を上書きする")
    p.add_argument("--font", help="図形の文字のフォント（theme.toml の [text] font を上書き）")
    p.add_argument("--scale", type=float, default=1.0, help="拡大率（既定: 1.0 = SVG の 1px を 1px に）")
    p.add_argument("--no-mirror-actors", action="store_true", help="図の下側の参加者を描かない")
    p.add_argument("--keep-svg", type=Path, help="中間の SVG を残すディレクトリ")
    p.add_argument("--mmdc", help="mmdc のコマンド（既定: PATH の mmdc、無ければ npx で固定版）")
    a = p.parse_args(argv)

    try:
        theme = load_theme(a.theme)
    except ThemeError as e:
        sys.exit(str(e))
    if a.font:
        theme["text"]["font"] = a.font

    out = a.output or a.inputs[0].with_suffix(".xlsx")
    sheets, used = [], set()
    with tempfile.TemporaryDirectory() as tmp:
        for src in a.inputs:
            if not src.exists():
                sys.exit(f"not found: {src}")
            if src.suffix.lower() == ".svg":
                svg = src
            else:
                svg = (a.keep_svg or Path(tmp)) / f"{src.stem}.svg"
                svg.parent.mkdir(parents=True, exist_ok=True)
                render_svg(src, svg, a.mmdc, theme["mermaid"]["theme"])
            kind = ET.parse(svg).getroot().get("aria-roledescription") or ""
            if kind == "sequence":
                shapes = SvgConverter(theme, mirror_actors=not a.no_mirror_actors).convert(svg)
            else:
                if kind not in SUPPORTED_GENERIC:
                    print(f"warning: {src.name} は確認していない種類の図です（{kind or '不明'}）。"
                          "見た目は SVG のまま写しますが、崩れることがあります", file=sys.stderr)
                shapes = GenericConverter(theme).convert(svg)
            if not shapes:
                sys.exit(f"図形が 1 つも取れませんでした: {src}")
            name = sheet_name(src.stem, used)
            sheets.append((name, DrawingWriter(theme, a.scale).drawing(shapes, name)))
            print(f"{src} -> sheet '{name}' ({len(shapes)} shapes)")
    write_xlsx(out, sheets, theme["text"]["font"])
    print(f"wrote {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
