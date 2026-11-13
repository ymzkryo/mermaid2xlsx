#!/usr/bin/env python3
"""Mermaid のシーケンス図を、Excel のネイティブ図形で描いた .xlsx に変換する。

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

    known = STYLE_SECTIONS | {"layout", "text", "arrows", "seq_number"}
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

    allowed = {"layout": {"margin"}, "text": {"font", "color"},
               "arrows": {"sync", "async", "size"}, "seq_number": {"fill", "text", "r"}}
    for sec, keys in allowed.items():
        for k in t.get(sec, {}):
            if k not in keys:
                errors.append(f"[{sec}] {k}: 知らない項目（使えるのは {', '.join(sorted(keys))}）")

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


@dataclass
class TextSpec:
    lines: list[str]
    size_px: float = 16
    bold: bool = False
    color: str = "000000"
    align: str = "ctr"  # l / ctr / r


@dataclass
class Shape:
    kind: str  # rect / ellipse / textbox / line / path
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

    def ln(self, st: Style, head: str | None = None, tail: str | None = None) -> str:
        if st.line is None:
            return "<a:ln><a:noFill/></a:ln>"
        dash = f'<a:prstDash val="{st.dash}"/>' if st.dash else ""
        sz = self.arrow_size
        he = f'<a:headEnd type="{head}" w="{sz}" len="{sz}"/>' if head else ""
        te = f'<a:tailEnd type="{tail}" w="{sz}" len="{sz}"/>' if tail else ""
        return (f'<a:ln w="{int(st.w * EMU_PER_PT * self.scale)}">'
                f'<a:solidFill><a:srgbClr val="{st.line}"/></a:solidFill>{dash}'
                f'<a:round/>{he}{te}</a:ln>')

    @staticmethod
    def fill(st: Style) -> str:
        return f'<a:solidFill><a:srgbClr val="{st.fill}"/></a:solidFill>' if st.fill else "<a:noFill/>"

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
                    f'{self.ln(s.style, s.head, s.tail)}</xdr:spPr></xdr:cxnSp>')

        if s.kind == "path":
            bx0, by0, bx1, by1 = s.bbox()
            geom = self.custgeom(s, bx0, by0, bx1 - bx0, by1 - by0)
            x, y, w, h = bx0 - ox, by0 - oy, bx1 - bx0, by1 - by0
        else:
            prst = {"rect": "rect", "ellipse": "ellipse", "textbox": "rect"}[s.kind]
            geom = f'<a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>'
            x, y, w, h = s.x - ox, s.y - oy, s.w, s.h

        is_tb = s.kind == "textbox"
        style = Style() if is_tb else s.style
        tb_attr = ' txBox="1"' if is_tb else ""
        return (f'<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{self.nid()}" name="{name}"/>'
                f'<xdr:cNvSpPr{tb_attr}/></xdr:nvSpPr><xdr:spPr>{self.xfrm(x, y, w, h)}{geom}'
                f'{self.fill(style)}{self.ln(style, s.head, s.tail)}</xdr:spPr>'
                f'{self.txbody(s.text, wrap=not is_tb and s.kind != "ellipse")}</xdr:sp>')

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
def render_svg(mmd: Path, out_svg: Path, mmdc: str | None) -> None:
    if mmdc:
        cmd = mmdc.split()
    elif shutil.which("mmdc"):
        cmd = ["mmdc"]
    else:
        cmd = ["npx", "-y", MMDC_PACKAGE]
    cmd = [*cmd, "-i", str(mmd), "-o", str(out_svg), "-q"]
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0 or not out_svg.exists():
        sys.exit(f"mmdc failed for {mmd}:\n{r.stderr or r.stdout}")


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Mermaid のシーケンス図を Excel の図形に変換する")
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
                render_svg(src, svg, a.mmdc)
            shapes = SvgConverter(theme, mirror_actors=not a.no_mirror_actors).convert(svg)
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
