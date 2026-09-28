"""
venn_engine.py - set logic + pure-SVG renderers for Venn / Euler / UpSet plots.

No plotting library is used: every figure is hand-written SVG so text stays real
text (crisp at any zoom, editable in PowerPoint via "Convert to Shape") and
styling is fully controlled.  Only numpy is required.

PowerPoint-compatibility rules followed throughout:
  * presentation attributes only (no <style>/CSS classes, no mix-blend-mode)
  * no `dominant-baseline` (PPT ignores it) - vertical centring is done by
    offsetting y by ~0.35 em instead
  * sizes in px at 96 dpi, so 1 pt = 4/3 px and the figure lands on the slide
    at the physical size chosen in the app.
"""
from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from functools import lru_cache
from itertools import combinations
from xml.sax.saxutils import escape

import numpy as np

PT = 96 / 72  # px per point

# --------------------------------------------------------------------------- #
# Themes / palettes
# --------------------------------------------------------------------------- #
# Journal palettes follow the widely used ggsci definitions (NPG, AAAS, Lancet,
# NEJM, JAMA, JCO); "Cell-style" is a muted palette in the spirit of Cell Press
# figures.  They are *inspired by* journal house styles, not official assets.
THEMES: dict[str, dict] = {
    "Nature (NPG)": dict(
        palette=["#E64B35", "#4DBBD5", "#00A087", "#3C5488", "#F39B7F",
                 "#8491B4", "#91D1C2", "#DC0000"],
        font="Arial", opacity=0.55, style="Filled + outline", stroke="Darker shade",
        stroke_w=1.0, count_color="#1A1A1A", label_color="auto"),
    "Cell-style": dict(
        palette=["#3B6FB6", "#E8A33D", "#7FB069", "#C4515E", "#6C5B7B",
                 "#40A4B8", "#B08968", "#7F7F7F"],
        font="Arial", opacity=0.45, style="Filled + outline", stroke="Darker shade",
        stroke_w=1.4, count_color="#222222", label_color="auto"),
    "Science (AAAS)": dict(
        palette=["#3B4992", "#EE0000", "#008B45", "#631879", "#008280",
                 "#BB0021", "#5F559B", "#A20056"],
        font="Helvetica", opacity=0.45, style="Filled + outline", stroke="Darker shade",
        stroke_w=1.0, count_color="#111111", label_color="auto"),
    "Lancet": dict(
        palette=["#00468B", "#ED0000", "#42B540", "#0099B4", "#925E9F",
                 "#FDAF91", "#AD002A", "#ADB6B6"],
        font="Arial", opacity=0.45, style="Filled + outline", stroke="Darker shade",
        stroke_w=1.0, count_color="#111111", label_color="auto"),
    "NEJM": dict(
        palette=["#BC3C29", "#0072B5", "#E18727", "#20854E", "#7876B1",
                 "#6F99AD", "#FFDC91", "#EE4C97"],
        font="Arial", opacity=0.5, style="Filled + outline", stroke="Darker shade",
        stroke_w=1.0, count_color="#111111", label_color="auto"),
    "JAMA": dict(
        palette=["#374E55", "#DF8F44", "#00A1D5", "#B24745", "#79AF97",
                 "#6A6599", "#80796B", "#374E55"],
        font="Arial", opacity=0.5, style="Filled + outline", stroke="Darker shade",
        stroke_w=1.0, count_color="#111111", label_color="auto"),
    "JCO": dict(
        palette=["#0073C2", "#EFC000", "#868686", "#CD534C", "#7AA6DC",
                 "#003C67", "#8F7700", "#3B3B3B"],
        font="Arial", opacity=0.5, style="Filled + outline", stroke="Darker shade",
        stroke_w=1.0, count_color="#111111", label_color="auto"),
    "Okabe-Ito (colour-blind safe)": dict(
        palette=["#E69F00", "#56B4E9", "#009E73", "#F0E442", "#0072B2",
                 "#D55E00", "#CC79A7", "#000000"],
        font="Arial", opacity=0.5, style="Filled + outline", stroke="Darker shade",
        stroke_w=1.0, count_color="#111111", label_color="auto"),
    "Corporate navy (PPT)": dict(
        palette=["#1F3A68", "#2E75B6", "#5FA8D3", "#9DC3E6", "#F2A541",
                 "#7F7F7F", "#264653", "#A5A5A5"],
        font="Calibri", opacity=0.55, style="Filled, white edge", stroke="White",
        stroke_w=2.0, count_color="#FFFFFF", label_color="auto"),
    "Soft pastel": dict(
        palette=["#F4A6A6", "#A6C8F4", "#B5E3B0", "#F7D59C", "#CDB4DB",
                 "#A0E1E0", "#FFC8DD", "#D9D9D9"],
        font="Arial", opacity=0.7, style="Filled, no outline", stroke="None",
        stroke_w=0.0, count_color="#333333", label_color="#333333"),
    "Vivid gradient": dict(
        palette=["#FF5E62", "#36D1DC", "#8E54E9", "#F7B733", "#11998E",
                 "#FC4A1A", "#4776E6", "#A8A8A8"],
        font="Arial", opacity=0.55, style="Gradient", stroke="Darker shade",
        stroke_w=0.8, count_color="#1A1A1A", label_color="auto"),
    "Classic outline (print)": dict(
        palette=["#000000", "#000000", "#000000", "#000000", "#000000",
                 "#000000", "#000000", "#000000"],
        font="Arial", opacity=0.0, style="Outline only", stroke="Same colour",
        stroke_w=1.5, count_color="#000000", label_color="#000000"),
    "Grayscale": dict(
        palette=["#4D4D4D", "#8C8C8C", "#BFBFBF", "#262626", "#A6A6A6",
                 "#666666", "#D9D9D9", "#333333"],
        font="Arial", opacity=0.45, style="Filled + outline", stroke="Darker shade",
        stroke_w=1.0, count_color="#000000", label_color="#000000"),
}

FILL_STYLES = ["Filled + outline", "Filled, no outline", "Filled, white edge",
               "Gradient", "Outline only"]
STROKE_MODES = ["Darker shade", "Same colour", "White", "Black", "None"]
FONTS = ["Arial", "Helvetica", "Calibri", "Aptos", "Segoe UI", "Times New Roman",
         "Georgia"]


def hex_to_rgb(h: str) -> tuple[int, int, int]:
    h = h.lstrip("#")
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def rgb_to_hex(rgb) -> str:
    return "#{:02X}{:02X}{:02X}".format(*(max(0, min(255, int(round(v)))) for v in rgb))


def mix(c1: str, c2: str, t: float) -> str:
    a, b = hex_to_rgb(c1), hex_to_rgb(c2)
    return rgb_to_hex([a[i] * (1 - t) + b[i] * t for i in range(3)])


def luminance(c: str) -> float:
    r, g, b = (v / 255 for v in hex_to_rgb(c))
    f = lambda v: v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


# --------------------------------------------------------------------------- #
# Set logic
# --------------------------------------------------------------------------- #
_SPLIT = re.compile(r"[\r\n\t,;]+")


def split_text(text: str) -> list[str]:
    """Split pasted text on newlines, tabs, commas or semicolons."""
    return [t for t in (s.strip() for s in _SPLIT.split(text or "")) if t]


def normalise(item, case_sensitive: bool, numeric: bool) -> str | None:
    if item is None:
        return None
    if isinstance(item, float) and math.isnan(item):
        return None
    s = str(item).strip()
    if not s or s.lower() in {"nan", "none", "null"}:
        return None
    if numeric:
        try:
            f = float(s.replace(",", "")) if re.fullmatch(r"[-+]?[\d,]*\.?\d+(e[-+]?\d+)?", s, re.I) else None
        except ValueError:
            f = None
        if f is not None and math.isfinite(f):
            return str(int(f)) if f == int(f) and abs(f) < 1e15 else repr(f)
    return s if case_sensitive else s.casefold()


@dataclass
class SetData:
    names: list[str]
    sets: list[set]                      # normalised keys
    display: dict = field(default_factory=dict)   # key -> first-seen original text
    duplicates: list[int] = field(default_factory=list)

    @property
    def n(self) -> int:
        return len(self.names)

    @property
    def union(self) -> set:
        return set().union(*self.sets) if self.sets else set()

    def membership(self) -> dict:
        """key -> bitmask of the sets containing it."""
        m: dict = {}
        for i, s in enumerate(self.sets):
            for k in s:
                m[k] = m.get(k, 0) | (1 << i)
        return m

    def regions(self) -> dict[int, list[str]]:
        """Exclusive regions: bitmask -> sorted display items in exactly those sets."""
        out: dict[int, list] = {}
        for k, mask in self.membership().items():
            out.setdefault(mask, []).append(self.display.get(k, k))
        for mask in out:
            out[mask].sort(key=_natural_key)
        return out

    def region_name(self, mask: int, only: bool = True) -> str:
        ins = [self.names[i] for i in range(self.n) if mask >> i & 1]
        s = " ∩ ".join(ins)
        return s + (" only" if only and len(ins) < self.n else "")


def _natural_key(s: str):
    return [int(t) if t.isdigit() else t.casefold() for t in re.split(r"(\d+)", str(s))]


def build_sets(names: list[str], raw_lists: list[list], case_sensitive=False,
               numeric=True) -> SetData:
    sets, display, dups = [], {}, []
    for raw in raw_lists:
        s, seen = set(), 0
        for item in raw:
            k = normalise(item, case_sensitive, numeric)
            if k is None:
                continue
            seen += 1
            s.add(k)
            display.setdefault(k, str(item).strip())
        sets.append(s)
        dups.append(seen - len(s))
    return SetData(list(names), sets, display, dups)


# --------------------------------------------------------------------------- #
# Geometry
# --------------------------------------------------------------------------- #
# A shape is (cx, cy, rx, ry, angle_deg) in abstract units, SVG orientation
# (y grows downward).  4- and 5-set layouts were tuned numerically to maximise
# the smallest region's inscribed radius so every count fits.
Shape = tuple[float, float, float, float, float]


def classic_layout(n: int) -> list[Shape]:
    if n == 1:
        return [(0, 0, 1, 1, 0)]
    if n == 2:
        return [(-0.62, 0, 1, 1, 0), (0.62, 0, 1, 1, 0)]
    if n == 3:
        d = 1.22
        h = d * math.sqrt(3) / 2
        return [(-d / 2, -h / 3, 1, 1, 0), (d / 2, -h / 3, 1, 1, 0), (0, 2 * h / 3, 1, 1, 0)]
    if n == 4:
        x1, y1, x2, y2, rx, ry, a = 0.479, 0.200, 0.219, -0.154, 1.259, 0.655, 41.5
        return [(-x1, y1, rx, ry, a), (-x2, y2, rx, ry, a),
                (x2, y2, rx, ry, -a), (x1, y1, rx, ry, -a)]
    if n == 5:
        R, rx, ry, phi = 0.287, 0.995, 0.439, 19.8
        out = []
        for i in range(5):
            t = math.radians(-90 + 72 * i)
            out.append((R * math.cos(t), R * math.sin(t), rx, ry, math.degrees(t) + phi))
        return out
    raise ValueError("Classic Venn supports 1-5 sets")


def _lens_area(r1, r2, d):
    if d >= r1 + r2:
        return 0.0
    if d <= abs(r1 - r2):
        return math.pi * min(r1, r2) ** 2
    a = r1 * r1 * math.acos((d * d + r1 * r1 - r2 * r2) / (2 * d * r1))
    b = r2 * r2 * math.acos((d * d + r2 * r2 - r1 * r1) / (2 * d * r2))
    c = 0.5 * math.sqrt((-d + r1 + r2) * (d + r1 - r2) * (d - r1 + r2) * (d + r1 + r2))
    return a + b - c


def _distance_for_overlap(r1, r2, overlap):
    if overlap <= 0:
        return (r1 + r2) * 1.06
    full = math.pi * min(r1, r2) ** 2
    if overlap >= full - 1e-12:
        return abs(r1 - r2) * 0.45
    lo, hi = abs(r1 - r2), r1 + r2
    for _ in range(80):
        mid = (lo + hi) / 2
        if _lens_area(r1, r2, mid) > overlap:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


def proportional_layout(sizes: list[int], pair_overlap: dict) -> list[Shape]:
    """Area-proportional (Euler) circles for 2 or 3 sets.  Pairwise overlaps are
    exact; the triple region is approximate (as in BioVenn / eulerAPE-style tools)."""
    n = len(sizes)
    tot = max(1, max(sizes))
    r = [math.sqrt(max(s, tot * 1e-4) / tot / math.pi) for s in sizes]
    if n == 1:
        return [(0, 0, r[0], r[0], 0)]
    unit = lambda k: pair_overlap[k] / tot / 1.0
    if n == 2:
        d = _distance_for_overlap(r[0], r[1], unit((0, 1)) * 1.0)
        return [(-d / 2, 0, r[0], r[0], 0), (d / 2, 0, r[1], r[1], 0)]
    dab = _distance_for_overlap(r[0], r[1], unit((0, 1)))
    dac = _distance_for_overlap(r[0], r[2], unit((0, 2)))
    dbc = _distance_for_overlap(r[1], r[2], unit((1, 2)))
    A = np.array([0.0, 0.0])
    B = np.array([max(dab, 1e-6), 0.0])
    cosA = (dab ** 2 + dac ** 2 - dbc ** 2) / (2 * max(dab, 1e-6) * max(dac, 1e-6))
    cosA = max(-1.0, min(1.0, cosA))
    C = np.array([dac * cosA, dac * math.sqrt(1 - cosA ** 2)])
    pts = np.array([A, B, C])
    pts -= pts.mean(axis=0)
    return [(p[0], p[1], r[i], r[i], 0) for i, p in enumerate(pts)]


def _inside(shape: Shape, X, Y):
    cx, cy, rx, ry, a = shape
    c, s = math.cos(math.radians(a)), math.sin(math.radians(a))
    dx, dy = X - cx, Y - cy
    u, v = dx * c + dy * s, -dx * s + dy * c
    return (u / rx) ** 2 + (v / ry) ** 2 <= 1


def shapes_bbox(shapes):
    xs, ys = [], []
    for cx, cy, rx, ry, a in shapes:
        c, s = math.cos(math.radians(a)), math.sin(math.radians(a))
        hw = math.sqrt((rx * c) ** 2 + (ry * s) ** 2)
        hh = math.sqrt((rx * s) ** 2 + (ry * c) ** 2)
        xs += [cx - hw, cx + hw]
        ys += [cy - hh, cy + hh]
    return min(xs), min(ys), max(xs), max(ys)


@lru_cache(maxsize=64)
def region_anchors(shapes: tuple, res: int = 260):
    """For every non-empty region (bitmask) return (x, y, clearance) - the point
    furthest from any boundary (approx. pole of inaccessibility), found by
    rasterising the diagram and eroding all regions simultaneously."""
    x0, y0, x1, y1 = shapes_bbox(shapes)
    span = max(x1 - x0, y1 - y0)
    px = span / res
    xs = np.arange(x0, x1 + px, px)
    ys = np.arange(y0, y1 + px, px)
    X, Y = np.meshgrid(xs, ys)
    L = np.zeros(X.shape, dtype=np.int32)
    for i, sh in enumerate(shapes):
        L |= _inside(sh, X, Y).astype(np.int32) << i
    h, w = L.shape
    P = np.pad(L, 1, constant_values=-1)
    D = np.zeros(L.shape, dtype=np.int32)
    alive = L > 0
    it = 0
    while alive.any() and it < 400:
        it += 1
        A = np.pad(alive, 1)
        nbrs = [(0, 1), (1, 0), (0, -1), (-1, 0)]
        if it % 2:
            nbrs += [(1, 1), (1, -1), (-1, 1), (-1, -1)]
        new = alive.copy()
        for dy, dx in nbrs:
            new &= (P[1 + dy:1 + dy + h, 1 + dx:1 + dx + w] == L) & A[1 + dy:1 + dy + h, 1 + dx:1 + dx + w]
        alive = new
        D[alive] = it
    out = {}
    for m in np.unique(L):
        if m <= 0:
            continue
        mask = L == m
        dmax = D[mask].max()
        cand = mask & (D == dmax)
        cy_, cx_ = np.nonzero(cand)
        mx, my = cx_.mean(), cy_.mean()
        j = np.argmin((cx_ - mx) ** 2 + (cy_ - my) ** 2)
        out[int(m)] = (float(xs[cx_[j]]), float(ys[cy_[j]]), float((dmax + 0.7) * px))
    return out


def ellipse_point(shape: Shape, t: float):
    cx, cy, rx, ry, a = shape
    c, s = math.cos(math.radians(a)), math.sin(math.radians(a))
    u, v = rx * math.cos(t), ry * math.sin(t)
    return cx + u * c - v * s, cy + u * s + v * c


def set_label_anchor(shapes, i, prefer_top=False, classic_n=0):
    """Pick a boundary point of shape i that is far from the diagram centre and
    not covered by other shapes; return point + outward unit direction."""
    if classic_n in (2, 3):
        sh = shapes[i]
        up = not (classic_n == 3 and i == 2)
        return (sh[0], sh[1] - sh[3], 0.0, -1.0) if up else (sh[0], sh[1] + sh[3], 0.0, 1.0)
    if classic_n == 4:
        # outer ellipses: label beside their upper tip; inner ones: centred above it
        left = i < 2
        tx, ty = ellipse_point(shapes[i], math.pi if left else 0.0)
        if i in (0, 3):
            return (tx, ty, -1.0 if left else 1.0, -0.25)
        return (tx, ty, 0.0, -1.0)
    x0, y0, x1, y1 = shapes_bbox(shapes)
    gx, gy = (x0 + x1) / 2, (y0 + y1) / 2
    span = max(x1 - x0, y1 - y0)
    best, best_score = None, -1e9
    sh = shapes[i]
    for k in range(144):
        t = 2 * math.pi * k / 144
        px_, py_ = ellipse_point(sh, t)
        dx, dy = px_ - sh[0], py_ - sh[1]
        nrm = math.hypot(dx, dy) or 1
        ox, oy = px_ + dx / nrm * span * 0.02, py_ + dy / nrm * span * 0.02
        covered = sum(bool(_inside(o, np.array(ox), np.array(oy))) for j, o in enumerate(shapes) if j != i)
        score = math.hypot(px_ - gx, py_ - gy) / span - covered * 1.0
        if prefer_top:
            score += 0.25 * (-(py_ - gy) / span)
        if score > best_score:
            best_score, best = score, (px_, py_, dx / nrm, dy / nrm)
    return best


# --------------------------------------------------------------------------- #
# Venn / Euler SVG renderer
# --------------------------------------------------------------------------- #
@dataclass
class Style:
    width: int = 1000
    height: int = 700
    colors: list = field(default_factory=lambda: THEMES["Nature (NPG)"]["palette"])
    opacity: float = 0.5
    fill_style: str = "Filled + outline"
    stroke_mode: str = "Darker shade"
    stroke_w: float = 1.0
    font: str = "Arial"
    count_pt: float = 14
    label_pt: float = 16
    title_pt: float = 18
    title: str = ""
    count_color: str = "#1A1A1A"
    label_color: str = "auto"         # "auto" = darker set colour
    count_mode: str = "Count"          # Count | Count (%) | % | Items | None
    max_items: int = 6
    bold_counts: bool = False
    bold_labels: bool = True
    show_set_size: bool = True
    thousands: bool = True
    auto_contrast: bool = True
    emphasise_core: bool = False
    hide_zero: bool = False
    legend: str = "Labels"             # Labels | Legend | Both
    background: str = "transparent"    # transparent | white
    margin_pt: float = 12


def _fmt(n: int, thousands: bool) -> str:
    return f"{n:,}" if thousands else str(n)


def _text_w(s: str, px: float) -> float:
    """Rough text width (Arial-ish metrics) for layout/fit purposes."""
    w = 0.0
    for ch in s:
        if ch in "il.,:;|!'I1 ":
            w += 0.30
        elif ch in "mwMW@%":
            w += 0.85
        elif ch.isupper() or ch.isdigit():
            w += 0.64
        else:
            w += 0.53
    return w * px


def _font_attr(font: str) -> str:
    fb = "Times New Roman, serif" if font in ("Times New Roman", "Georgia") else "Arial, Helvetica, sans-serif"
    return f"{font}, {fb}" if font not in ("Arial",) else "Arial, Helvetica, sans-serif"


def _stroke_for(color: str, st: Style) -> str | None:
    fs = st.fill_style
    if fs == "Filled, no outline":
        return None
    if fs == "Filled, white edge":
        return "#FFFFFF"
    mode = st.stroke_mode
    if mode == "None":
        return None
    if mode == "White":
        return "#FFFFFF"
    if mode == "Black":
        return "#000000"
    if mode == "Same colour":
        return color
    return mix(color, "#000000", 0.28)


def _svg_open(st: Style, extra_defs: str = "") -> list[str]:
    out = [f'<svg xmlns="http://www.w3.org/2000/svg" version="1.1" '
           f'width="{st.width}" height="{st.height}" viewBox="0 0 {st.width} {st.height}">']
    if extra_defs:
        out.append(f"<defs>{extra_defs}</defs>")
    if st.background == "white":
        out.append(f'<rect x="0" y="0" width="{st.width}" height="{st.height}" fill="#FFFFFF"/>')
    return out


def _text(x, y, s, px, font, color, anchor="middle", bold=False, italic=False) -> str:
    fw = ' font-weight="bold"' if bold else ""
    fi = ' font-style="italic"' if italic else ""
    return (f'<text x="{x:.1f}" y="{y:.1f}" font-family="{escape(_font_attr(font))}" '
            f'font-size="{px:.1f}" fill="{color}" text-anchor="{anchor}"{fw}{fi}>{escape(s)}</text>')


def render_venn(data: SetData, st: Style, proportional: bool = False) -> tuple[str, list[str]]:
    """Return (svg, warnings)."""
    warnings: list[str] = []
    n = data.n
    regions = data.regions()
    union_n = max(1, len(data.union))
    sizes = [len(s) for s in data.sets]

    if proportional:
        if n > 3:
            raise ValueError("Area-proportional layout supports 2-3 sets")
        pairs = {(i, j): len(data.sets[i] & data.sets[j]) for i, j in combinations(range(n), 2)}
        shapes = proportional_layout(sizes, pairs)
    else:
        shapes = classic_layout(n)
    shapes = tuple(tuple(round(float(v), 6) for v in s) for s in shapes)
    anchors = region_anchors(shapes)

    label_px, count_px, title_px = st.label_pt * PT, st.count_pt * PT, st.title_pt * PT
    margin = st.margin_pt * PT
    top = margin + (title_px * 1.6 if st.title else 0)
    colors = [st.colors[i % len(st.colors)] for i in range(n)]

    # set-label text
    set_labels = []
    for i, nm in enumerate(data.names):
        set_labels.append(f"{nm} ({_fmt(sizes[i], st.thousands)})" if st.show_set_size else nm)

    # legend block size (placed on the right)
    legend_w = 0.0
    if st.legend in ("Legend", "Both"):
        legend_w = max(_text_w(s, label_px) for s in set_labels) + label_px * 2.2

    # --- fit: iterate scale so shapes + outside labels fit in canvas ---------
    x0, y0, x1, y1 = shapes_bbox(shapes)
    avail_w = st.width - 2 * margin - legend_w
    avail_h = st.height - top - margin
    label_anchor = [set_label_anchor(shapes, i, prefer_top=n <= 4,
                                     classic_n=0 if proportional else n) for i in range(n)]
    show_labels = st.legend in ("Labels", "Both")

    def layout(scale):
        boxes = []
        pad = label_px * 0.45
        for i in range(n):
            px_, py_, dx, dy = label_anchor[i]
            X, Y = px_ * scale, py_ * scale
            X += dx * pad
            Y += dy * pad
            w = _text_w(set_labels[i], label_px)
            if abs(dy) >= 0.5:
                anc, bx0 = "middle", X - w / 2
            elif dx > 0:
                anc, bx0 = "start", X
            else:
                anc, bx0 = "end", X - w
            if dy <= -0.5:
                base = Y - label_px * 0.15
            elif dy >= 0.5:
                base = Y + label_px * 0.85
            else:
                base = Y + label_px * 0.35 + dy * label_px * 0.6
            boxes.append([X, base, anc, bx0, base - label_px * 0.8, bx0 + w, base + label_px * 0.2])
        # push overlapping labels apart horizontally
        gap = label_px * 0.6
        for _ in range(20):
            moved = False
            for a in range(n):
                for b in range(a + 1, n):
                    A_, B_ = boxes[a], boxes[b]
                    if A_[4] < B_[6] and B_[4] < A_[6] and A_[3] < B_[5] + gap and B_[3] < A_[5] + gap:
                        left_, right_ = (A_, B_) if (A_[3] + A_[5]) <= (B_[3] + B_[5]) else (B_, A_)
                        ov = (left_[5] + gap - right_[3]) / 2 + 0.5
                        for bx, sgn in ((left_, -1), (right_, 1)):
                            bx[0] += sgn * ov
                            bx[3] += sgn * ov
                            bx[5] += sgn * ov
                        moved = True
            if not moved:
                break
        return boxes

    scale = min(avail_w / (x1 - x0), avail_h / (y1 - y0))
    for _ in range(6):
        boxes = layout(scale) if show_labels else []
        bx0 = min([x0 * scale] + [b[3] for b in boxes])
        by0 = min([y0 * scale] + [b[4] for b in boxes])
        bx1 = max([x1 * scale] + [b[5] for b in boxes])
        by1 = max([y1 * scale] + [b[6] for b in boxes])
        f = min(avail_w / (bx1 - bx0), avail_h / (by1 - by0))
        if abs(f - 1) < 0.005:
            break
        scale *= f
    boxes = layout(scale) if show_labels else []
    bx0 = min([x0 * scale] + [b[3] for b in boxes])
    by0 = min([y0 * scale] + [b[4] for b in boxes])
    bx1 = max([x1 * scale] + [b[5] for b in boxes])
    by1 = max([y1 * scale] + [b[6] for b in boxes])
    ox = margin + (avail_w - (bx1 - bx0)) / 2 - bx0
    oy = top + (avail_h - (by1 - by0)) / 2 - by0

    # --- defs (gradients) ----------------------------------------------------
    defs = ""
    if st.fill_style == "Gradient":
        for i, c in enumerate(colors):
            defs += (f'<radialGradient id="vg{i}" cx="40%" cy="35%" r="75%">'
                     f'<stop offset="0" stop-color="{mix(c, "#FFFFFF", 0.55)}"/>'
                     f'<stop offset="1" stop-color="{c}"/></radialGradient>')

    out = _svg_open(st, defs)
    if st.title:
        out.append(_text(st.width / 2, margin + title_px, st.title, title_px, st.font,
                         "#111111", bold=True))

    # --- shapes ---------------------------------------------------------------
    for i, (cx, cy, rx, ry, a) in enumerate(shapes):
        c = colors[i]
        X, Y = cx * scale + ox, cy * scale + oy
        if st.fill_style == "Outline only":
            fill, fop = "none", 0
        elif st.fill_style == "Gradient":
            fill, fop = f"url(#vg{i})", st.opacity
        else:
            fill, fop = c, st.opacity
        stroke = _stroke_for(c, st) if st.fill_style != "Outline only" else c
        sattr = (f' stroke="{stroke}" stroke-width="{max(st.stroke_w, 0.1) * PT:.2f}"'
                 if stroke and st.stroke_w > 0 else ' stroke="none"')
        fattr = f'fill="{fill}"' + (f' fill-opacity="{fop:.2f}"' if fill != "none" else "")
        if abs(rx - ry) < 1e-9:
            out.append(f'<circle cx="{X:.2f}" cy="{Y:.2f}" r="{rx * scale:.2f}" {fattr}{sattr}/>')
        else:
            out.append(f'<ellipse cx="{X:.2f}" cy="{Y:.2f}" rx="{rx * scale:.2f}" ry="{ry * scale:.2f}" '
                       f'transform="rotate({a:.2f} {X:.2f} {Y:.2f})" {fattr}{sattr}/>')

    # --- region labels -----------------------------------------------------------
    full = (1 << n) - 1
    for mask in range(1, full + 1):
        items = regions.get(mask, [])
        cnt = len(items)
        if cnt == 0 and st.hide_zero:
            continue
        if mask not in anchors:
            if cnt:
                warnings.append(f"Region '{data.region_name(mask)}' ({cnt}) cannot be drawn "
                                f"in the proportional layout - switch to classic to show it.")
            continue
        ax, ay, clear = anchors[mask]
        X, Y = ax * scale + ox, ay * scale + oy
        core = st.emphasise_core and mask == full and n > 1
        cpx = count_px * (1.25 if core else 1.0)
        room = clear * scale                      # inscribed radius in px
        tcol = _count_fill(st, mask, colors)
        if st.count_mode != "Items":
            txt = _count_text(cnt, union_n, st)
            if txt is None:
                continue
            fitpx = min(cpx, max(count_px * 0.55, 1.75 * room / max(0.6, _text_w(txt, 1.0))))
            out.append(_text(X, Y + fitpx * 0.35, txt, fitpx, st.font, tcol,
                             bold=st.bold_counts or core))
        else:
            k_max = max(1, st.max_items)
            lines = items[:k_max]
            if cnt > k_max:
                lines = (items[:k_max - 1] + [f"+{cnt - (k_max - 1)} more"]) if k_max > 1 else [f"{cnt} items"]
            if not lines:
                lines = ["0"]
            wmax = max(_text_w(ln, 1.0) for ln in lines)
            fs = min(count_px, 1.7 * room / (len(lines) * 1.12), 1.75 * room / max(0.6, wmax))
            fs = max(fs, count_px * 0.4)
            lh = fs * 1.12
            y_start = Y - (len(lines) - 1) * lh / 2 + fs * 0.35
            for k, ln in enumerate(lines):
                out.append(_text(X, y_start + k * lh, ln, fs, st.font, tcol, bold=st.bold_counts,
                                 italic=ln.startswith("+") and ln.endswith("more")))

    # --- set labels -----------------------------------------------------------
    if show_labels:
        for i, b in enumerate(boxes):
            X, base, anc = b[0] + ox, b[1] + oy, b[2]
            col = mix(colors[i], "#000000", 0.25) if st.label_color == "auto" else st.label_color
            if luminance(col) > 0.6:
                col = mix(col, "#000000", 0.45)
            out.append(_text(X, base, set_labels[i], label_px, st.font, col, anchor=anc,
                             bold=st.bold_labels))

    # --- legend ---------------------------------------------------------------
    if legend_w:
        lx = st.width - margin - legend_w + label_px * 0.6
        tot_h = n * label_px * 1.5
        ly = top + (avail_h - tot_h) / 2
        for i in range(n):
            y = ly + i * label_px * 1.5
            c = colors[i]
            stroke = _stroke_for(c, st) or c
            out.append(f'<rect x="{lx:.1f}" y="{y:.1f}" width="{label_px:.1f}" height="{label_px:.1f}" '
                       f'rx="{label_px * 0.2:.1f}" fill="{c}" fill-opacity="{max(st.opacity, 0.35) if st.fill_style != "Outline only" else 0:.2f}" '
                       f'stroke="{stroke}" stroke-width="{max(st.stroke_w, 0.8) * PT:.2f}"/>')
            out.append(_text(lx + label_px * 1.4, y + label_px * 0.82, set_labels[i], label_px,
                             st.font, "#222222", anchor="start"))
    out.append("</svg>")
    return "\n".join(out), warnings


def _count_text(cnt, union_n, st: Style):
    if st.count_mode == "None":
        return None
    pct = 100 * cnt / union_n
    pct_s = f"{pct:.1f}%" if pct < 10 else f"{pct:.0f}%"
    if st.count_mode == "Count":
        return _fmt(cnt, st.thousands)
    if st.count_mode == "Count (%)":
        return f"{_fmt(cnt, st.thousands)} ({pct_s})"
    if st.count_mode == "%":
        return pct_s
    return _fmt(cnt, st.thousands)


def _contrast(c1: str, c2: str) -> float:
    l1, l2 = sorted((luminance(c1), luminance(c2)), reverse=True)
    return (l1 + 0.05) / (l2 + 0.05)


def region_colour(st: Style, mask: int, colors) -> str:
    """Approximate the rendered colour of a region (alpha compositing over white)."""
    col = "#FFFFFF"
    if st.fill_style == "Outline only":
        return col
    op = st.opacity * (0.85 if st.fill_style == "Gradient" else 1.0)
    for i, c in enumerate(colors):
        if mask >> i & 1:
            col = mix(col, c, op)
    return col


def _count_fill(st: Style, mask, colors):
    """Chosen number colour, swapped to white/near-black if it would be unreadable."""
    base = "#1A1A1A" if st.count_color == "auto" else st.count_color
    if not st.auto_contrast:
        return base
    bg = region_colour(st, mask, colors)
    if _contrast(base, bg) >= 3.0:
        return base
    dark, light = "#1A1A1A", "#FFFFFF"
    return light if _contrast(light, bg) > _contrast(dark, bg) else dark


# --------------------------------------------------------------------------- #
# UpSet plot renderer (for any number of sets; best for >= 4)
# --------------------------------------------------------------------------- #
@dataclass
class UpsetOpts:
    sort_by: str = "Size"          # Size | Degree
    max_bars: int = 30
    min_size: int = 1
    bar_color: str = "#3C3C3C"
    color_singletons: bool = True
    show_values: bool = True
    show_set_sizes: bool = True


def render_upset(data: SetData, st: Style, uo: UpsetOpts) -> str:
    n = data.n
    regions = {m: len(v) for m, v in data.regions().items() if len(v) >= max(1, uo.min_size)}
    items = list(regions.items())
    deg = lambda m: bin(m).count("1")
    if uo.sort_by == "Degree":
        items.sort(key=lambda kv: (deg(kv[0]), -kv[1]))
    else:
        items.sort(key=lambda kv: (-kv[1], deg(kv[0])))
    items = items[:uo.max_bars]
    colors = [st.colors[i % len(st.colors)] for i in range(n)]
    sizes = [len(s) for s in data.sets]

    count_px, label_px, title_px = st.count_pt * PT, st.label_pt * PT, st.title_pt * PT
    font = st.font
    margin = st.margin_pt * PT
    top = margin + (title_px * 1.6 if st.title else 0)

    name_w = max(_text_w(nm, label_px) for nm in data.names) + label_px * 0.8
    setbar_w = (st.width - 2 * margin) * 0.18 if uo.show_set_sizes else 0
    left = margin + setbar_w + name_w
    right = st.width - margin
    nbars = max(1, len(items))
    col_w = (right - left) / nbars
    dot_r = min(col_w * 0.28, label_px * 0.55)
    row_h = max(dot_r * 2.6, label_px * 1.35)
    matrix_h = row_h * n
    bottom = st.height - margin
    matrix_top = bottom - matrix_h
    axis_px = count_px * 0.85
    chart_top = top + (count_px * 1.3 if uo.show_values else 0)
    chart_bottom = matrix_top - row_h * 0.35
    chart_h = max(10, chart_bottom - chart_top)
    maxv = max([v for _, v in items] + [1])
    tick = _nice_step(maxv)
    ymax = math.ceil(maxv / tick) * tick

    out = _svg_open(st)
    if st.title:
        out.append(_text(st.width / 2, margin + title_px, st.title, title_px, font, "#111111", bold=True))
    grid = "#D9D9D9"
    ink = "#333333"
    # y axis
    v = 0
    while v <= ymax + 1e-9:
        y = chart_bottom - chart_h * v / ymax
        out.append(f'<line x1="{left:.1f}" y1="{y:.1f}" x2="{right:.1f}" y2="{y:.1f}" stroke="{grid}" stroke-width="0.75"/>')
        out.append(_text(left - axis_px * 0.4, y + axis_px * 0.35, _fmt(int(v), st.thousands), axis_px,
                         font, "#555555", anchor="end"))
        v += tick
    out.append(_text(left - axis_px * 0.4, chart_top - count_px * 0.6, "Intersection size", axis_px,
                     font, "#555555", anchor="end", bold=True))
    # zebra rows
    for i in range(n):
        y = matrix_top + i * row_h
        if i % 2 == 0:
            out.append(f'<rect x="{left - name_w + 2:.1f}" y="{y:.1f}" width="{right - left + name_w - 2:.1f}" '
                       f'height="{row_h:.1f}" fill="#F2F2F2"/>')
    # bars + matrix
    bw = col_w * 0.62
    for k, (mask, v) in enumerate(items):
        cx = left + col_w * (k + 0.5)
        h = chart_h * v / ymax
        members = [i for i in range(n) if mask >> i & 1]
        c = colors[members[0]] if (uo.color_singletons and len(members) == 1) else uo.bar_color
        out.append(f'<rect x="{cx - bw / 2:.1f}" y="{chart_bottom - h:.1f}" width="{bw:.1f}" height="{h:.1f}" fill="{c}"/>')
        if uo.show_values:
            out.append(_text(cx, chart_bottom - h - count_px * 0.3, _fmt(v, st.thousands),
                             count_px * 0.85, font, ink, bold=st.bold_counts))
        ys = []
        for i in range(n):
            cy = matrix_top + (i + 0.5) * row_h
            on = i in members
            fill = (colors[i] if uo.color_singletons and len(members) == 1 else uo.bar_color) if on else "#D0D0D0"
            if on:
                ys.append(cy)
            out.append(f'<circle cx="{cx:.1f}" cy="{cy:.1f}" r="{dot_r:.1f}" fill="{fill}"/>')
        if len(ys) > 1:
            out.append(f'<line x1="{cx:.1f}" y1="{min(ys):.1f}" x2="{cx:.1f}" y2="{max(ys):.1f}" '
                       f'stroke="{uo.bar_color}" stroke-width="{dot_r * 0.55:.1f}"/>')
    # set names + set-size bars
    maxs = max(sizes + [1])
    for i in range(n):
        cy = matrix_top + (i + 0.5) * row_h
        out.append(_text(left - label_px * 0.4, cy + label_px * 0.35, data.names[i], label_px, font,
                         mix(colors[i], "#000000", 0.3), anchor="end", bold=st.bold_labels))
        if uo.show_set_sizes:
            bar_right = left - name_w - label_px * 0.2
            wbar = (setbar_w - label_px * 2.5) * sizes[i] / maxs
            out.append(f'<rect x="{bar_right - wbar:.1f}" y="{cy - row_h * 0.3:.1f}" width="{wbar:.1f}" '
                       f'height="{row_h * 0.6:.1f}" fill="{colors[i]}" fill-opacity="{max(st.opacity, 0.7):.2f}"/>')
            out.append(_text(bar_right - wbar - axis_px * 0.3, cy + axis_px * 0.35,
                             _fmt(sizes[i], st.thousands), axis_px, font, "#555555", anchor="end"))
    if uo.show_set_sizes:
        out.append(_text(left - name_w - label_px * 0.2, matrix_top - row_h * 0.3, "Set size",
                         axis_px, font, "#555555", anchor="end", bold=True))
    out.append("</svg>")
    return "\n".join(out)


def _nice_step(v: float) -> float:
    raw = v / 5
    mag = 10 ** math.floor(math.log10(raw)) if raw > 0 else 1
    for m in (1, 2, 2.5, 5, 10):
        if raw <= m * mag:
            return max(1, m * mag)
    return 10 * mag
