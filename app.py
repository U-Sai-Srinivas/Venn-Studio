"""
Venn Studio - publication- and PowerPoint-ready Venn / Euler / UpSet diagrams.

Run:  python -m streamlit run app.py
"""
from __future__ import annotations

import base64
import io
import json
import random

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

import venn_engine as ve

st.set_page_config(page_title="Venn Studio", page_icon="◎", layout="wide")

MAX_SETS = 8
LETTERS = "ABCDEFGH"

SIZE_PRESETS = {  # name -> (width in, height in, (label pt, count pt, title pt))
    "PowerPoint – content area (10 × 5.6 in)": (10.0, 5.6, (18, 16, 22)),
    "PowerPoint – full slide 16:9 (13.33 × 7.5 in)": (13.333, 7.5, (20, 18, 26)),
    "PowerPoint – half slide (6 × 5 in)": (6.0, 5.0, (14, 12, 16)),
    "Square (5 × 5 in)": (5.0, 5.0, (14, 12, 16)),
    "Nature – single column (89 mm)": (3.5, 3.0, (8, 7, 9)),
    "Nature – double column (183 mm)": (7.2, 4.5, (8, 7, 9)),
    "Cell – single column (85 mm)": (3.35, 3.0, (8, 7, 9)),
    "Cell – 1.5 column (114 mm)": (4.49, 3.5, (8, 7, 9)),
    "Cell – full width (174 mm)": (6.85, 4.5, (8, 7, 9)),
    "Custom": (8.0, 5.0, None),
}

GENE_POOL = """TP53 EGFR KRAS MYC PIK3CA PTEN BRAF CDKN2A RB1 APC VHL NRAS SMAD4 ERBB2 CTNNB1
IDH1 ARID1A NOTCH1 FBXW7 KMT2D STAT3 JAK2 MTOR AKT1 MAPK1 MAP2K1 HIF1A VEGFA TGFB1 IL6 TNF
NFKB1 RELA CXCL8 CCL2 SOX2 POU5F1 NANOG KLF4 GATA3 FOXA1 ESR1 AR PGR CDH1 VIM SNAI1 TWIST1
ZEB1 COL1A1 ACTA2 FN1 MMP2 MMP9 TIMP1 SERPINE1 CTGF SPP1 POSTN LOX FAP PDGFRB THY1 DCN LUM
BCL2 BAX CASP3 CASP8 MCL1 BIRC5 CDK4 CDK6 CCND1 E2F1 MDM2 ATM ATR CHEK1 BRCA1 BRCA2 PARP1""".split()


# --------------------------------------------------------------------------- #
# State helpers
# --------------------------------------------------------------------------- #
def ss_default(key, value):
    if key not in st.session_state:
        st.session_state[key] = value


def example_lists(n: int) -> tuple[list[str], list[list[str]]]:
    rng = random.Random(7)
    names = ["RNA-seq DEGs", "Proteomics", "CRISPR hits", "ATAC peaks", "scRNA markers",
             "Phospho-sites", "Metabolomics", "GWAS loci"][:n]
    core = rng.sample(GENE_POOL, 6)
    lists = []
    for i in range(n):
        pool = [g for g in GENE_POOL if g not in core]
        k = rng.randint(22, 34)
        lists.append(core + rng.sample(pool, k))
    return names, lists


def apply_theme():
    t = ve.THEMES[st.session_state.theme]
    for i in range(MAX_SETS):
        st.session_state[f"col{i}"] = t["palette"][i % len(t["palette"])]
    st.session_state.theme_ver = st.session_state.get("theme_ver", 0) + 1
    st.session_state.opacity = t["opacity"]
    st.session_state.fill_style = t["style"]
    st.session_state.stroke_mode = t["stroke"]
    st.session_state.stroke_w = t["stroke_w"]
    st.session_state.font = t["font"]
    st.session_state.count_color = t["count_color"] if t["count_color"] != "auto" else "#1A1A1A"
    st.session_state.label_auto = t["label_color"] == "auto"
    if t["label_color"] != "auto":
        st.session_state.label_color = t["label_color"]


def sync_colour(i: int):
    st.session_state[f"col{i}"] = st.session_state[f"cp{i}_{st.session_state.theme_ver}"]


def apply_size_preset():
    w, h, fonts = SIZE_PRESETS[st.session_state.size_preset]
    st.session_state.w_in, st.session_state.h_in = w, h
    if fonts:
        st.session_state.label_pt, st.session_state.count_pt, st.session_state.title_pt = fonts


def reset_grid(names: list[str], lists: list[list[str]]):
    n = len(names)
    st.session_state.n_sets = n
    for i, nm in enumerate(names):
        st.session_state[f"name{i}"] = nm
        st.session_state[f"txt{i}"] = "\n".join(lists[i])
    rows = max([len(l) for l in lists] + [15])
    df = pd.DataFrame({f"s{i}": (lists[i] + [None] * rows)[:rows] for i in range(n)})
    st.session_state.grid_df = df
    st.session_state.grid_ver = st.session_state.get("grid_ver", 0) + 1


def use_file_sets(gnames, lists):
    reset_grid(gnames, lists)
    st.session_state.input_mode = "Grid (paste from Excel)"


# defaults
ss_default("theme", "Nature (NPG)")
ss_default("size_preset", list(SIZE_PRESETS)[0])
if "col0" not in st.session_state:
    apply_theme()
    apply_size_preset()
if "grid_df" not in st.session_state:
    reset_grid(*example_lists(3))
ss_default("input_mode", "Grid (paste from Excel)")

# --------------------------------------------------------------------------- #
# Header
# --------------------------------------------------------------------------- #
st.markdown(
    """
    <style>
      .block-container {padding-top: 1.6rem;}
      .vs-sub {color:#6b7280; margin-top:-0.6rem; margin-bottom:0.8rem;}
      div[data-testid="stExpander"] details summary p {font-weight:600;}
    </style>
    """,
    unsafe_allow_html=True,
)
st.title("◎ Venn Studio")
st.markdown('<div class="vs-sub">Publication-quality Venn, Euler and UpSet diagrams · '
            'true vector SVG · journal themes · PowerPoint-ready</div>', unsafe_allow_html=True)

# --------------------------------------------------------------------------- #
# Sidebar – figure styling
# --------------------------------------------------------------------------- #
with st.sidebar:
    st.header("Figure style")
    st.selectbox("Theme", list(ve.THEMES), key="theme", on_change=apply_theme,
                 help="Palettes inspired by journal house styles (ggsci definitions). "
                      "Changing theme resets colours, fill and font below.")
    n_sets_now = st.session_state.n_sets

    with st.expander("Colours", expanded=True):
        cols = st.columns(4)
        for i in range(n_sets_now):
            with cols[i % 4]:
                # set colours live in non-widget keys (col{i}) so they survive while a
                # set is hidden; the picker key is versioned so a theme change resets it
                st.color_picker(LETTERS[i], value=st.session_state[f"col{i}"],
                                key=f"cp{i}_{st.session_state.theme_ver}",
                                on_change=sync_colour, args=(i,),
                                help=st.session_state.get(f"name{i}", f"Set {LETTERS[i]}"))
        st.slider("Fill opacity", 0.0, 1.0, key="opacity", step=0.05)
        st.selectbox("Fill style", ve.FILL_STYLES, key="fill_style")
        c1, c2 = st.columns(2)
        c1.selectbox("Outline", ve.STROKE_MODES, key="stroke_mode")
        c2.number_input("Outline (pt)", 0.0, 6.0, key="stroke_w", step=0.25)

    with st.expander("Labels & numbers", expanded=True):
        st.selectbox("Region labels", ["Count", "Count (%)", "%", "Items", "None"], key="count_mode",
                     help="'Items' writes the actual elements inside each region (best for small sets).")
        if st.session_state.count_mode == "Items":
            st.slider("Max items per region", 1, 15, 6, key="max_items")
        st.selectbox("Set names shown as", ["Labels", "Legend", "Both"], key="legend_mode")
        c1, c2 = st.columns(2)
        c1.checkbox("Show set size (n)", True, key="show_n")
        c2.checkbox("1,000 separators", True, key="thousands")
        c1.checkbox("Bold numbers", False, key="bold_counts")
        c2.checkbox("Bold set names", True, key="bold_labels")
        c1.checkbox("Hide zero regions", False, key="hide_zero")
        c2.checkbox("Emphasise core", False, key="emph_core",
                    help="Larger, bold count for the intersection of all sets.")
        c1, c2 = st.columns(2)
        c1.color_picker("Number colour", key="count_color")
        c1.checkbox("Auto-contrast", True, key="auto_contrast",
                    help="Numbers switch to white / black where the region colour would make them hard to read.")
        c2.checkbox("Name colour = set", key="label_auto")
        if not st.session_state.label_auto:
            ss_default("label_color", "#222222")
            c2.color_picker("Name colour", value=st.session_state.label_color, key="label_color_w",
                            on_change=lambda: st.session_state.update(
                                label_color=st.session_state.label_color_w))

    with st.expander("Typography", expanded=False):
        st.selectbox("Font", ve.FONTS, key="font",
                     help="Nature & Cell ask for Arial/Helvetica. Calibri/Aptos match Office decks.")
        c1, c2, c3 = st.columns(3)
        c1.number_input("Names pt", 4, 60, key="label_pt")
        c2.number_input("Numbers pt", 4, 60, key="count_pt")
        c3.number_input("Title pt", 4, 72, key="title_pt")
        st.text_input("Title (optional)", key="title")

    with st.expander("Size & export", expanded=True):
        st.selectbox("Canvas size", list(SIZE_PRESETS), key="size_preset", on_change=apply_size_preset,
                     help="Physical size when inserted into PowerPoint or a journal figure (96 px = 1 in).")
        c1, c2 = st.columns(2)
        c1.number_input("Width (in)", 1.0, 30.0, key="w_in", step=0.25)
        c2.number_input("Height (in)", 1.0, 30.0, key="h_in", step=0.25)
        st.radio("Background", ["Transparent", "White"], key="bg", horizontal=True)
        st.select_slider("PNG resolution", [150, 300, 600, 1200], value=300, key="dpi",
                         format_func=lambda v: f"{v} dpi")

# --------------------------------------------------------------------------- #
# Step 1 – data
# --------------------------------------------------------------------------- #
st.subheader("1 · Data")
top1, top2, top3, top4 = st.columns([1.1, 2.2, 1, 1])
with top1:
    n_sets = st.number_input("Number of sets", 2, MAX_SETS, key="n_sets")
with top2:
    st.segmented_control("Input", ["Grid (paste from Excel)", "Text boxes", "Upload file"],
                         key="input_mode", selection_mode="single")
with top3:
    st.write("")
    st.button("Load example", width="stretch",
              on_click=lambda: reset_grid(*example_lists(st.session_state.n_sets)))
with top4:
    st.write("")
    st.button("Clear all", width="stretch", on_click=lambda: reset_grid(
        [f"Set {LETTERS[i]}" for i in range(st.session_state.n_sets)],
        [[] for _ in range(st.session_state.n_sets)]))

for i in range(MAX_SETS):
    ss_default(f"name{i}", f"Set {LETTERS[i]}")
    ss_default(f"txt{i}", "")

name_cols = st.columns(n_sets)
for i in range(n_sets):
    name_cols[i].text_input(f"Name of set {i + 1}", key=f"name{i}", label_visibility="collapsed",
                            placeholder=f"Set {LETTERS[i]}")
names = [st.session_state[f"name{i}"].strip() or f"Set {LETTERS[i]}" for i in range(n_sets)]
if len(set(names)) < len(names):
    st.warning("Set names should be unique.")

mode = st.session_state.input_mode or "Grid (paste from Excel)"
raw_lists: list[list] = [[] for _ in range(n_sets)]

if mode == "Grid (paste from Excel)":
    st.caption("Click the top-left cell and **Ctrl+V** a block copied from Excel / Sheets – one set per "
               "column. The grid grows automatically. Use the boxes above for set names, or tick the "
               "header option if your pasted block includes them.")
    df = st.session_state.grid_df.copy()
    for i in range(n_sets):
        if f"s{i}" not in df:
            df[f"s{i}"] = None
    df = df[[f"s{i}" for i in range(n_sets)]].astype("object")
    cfg = {f"s{i}": st.column_config.TextColumn(f"{LETTERS[i]} · {names[i]}", width="medium")
           for i in range(n_sets)}
    edited = st.data_editor(df, column_config=cfg, num_rows="dynamic", hide_index=True,
                            width="stretch", height=320,
                            key=f"grid_{st.session_state.grid_ver}_{n_sets}")
    header_row = st.checkbox("First row of the grid contains set names", False)
    for i in range(n_sets):
        col = edited[f"s{i}"].tolist()
        if header_row and col:
            if col[0] not in (None, "") and str(col[0]).strip():
                names[i] = str(col[0]).strip()
            col = col[1:]
        raw_lists[i] = col

elif mode == "Text boxes":
    st.caption("Paste items separated by new lines, commas, tabs or semicolons.")
    tcols = st.columns(n_sets)
    for i in range(n_sets):
        tcols[i].text_area(names[i], key=f"txt{i}", height=260)
        raw_lists[i] = ve.split_text(st.session_state[f"txt{i}"])

else:
    up1, up2 = st.columns([2, 1])
    with up1:
        f = st.file_uploader("CSV, TSV or Excel", type=["csv", "tsv", "txt", "xlsx", "xls"])
    with up2:
        layout = st.radio("Layout", ["Wide – one column per set", "Long – item, set"],
                          help="Long: two columns, the item and the set it belongs to.")
    if f is not None:
        try:
            if f.name.lower().endswith((".xlsx", ".xls")):
                up = pd.read_excel(f, dtype=str)
            else:
                up = pd.read_csv(f, sep=None, engine="python", dtype=str)
        except Exception as e:  # noqa: BLE001
            st.error(f"Could not read file: {e}")
            up = None
        if up is not None:
            if layout.startswith("Long"):
                item_col, set_col = up.columns[:2]
                groups = list(dict.fromkeys(up[set_col].dropna()))[:MAX_SETS]
                upl = [up.loc[up[set_col] == g, item_col].tolist() for g in groups]
                gnames = [str(g) for g in groups]
            else:
                gnames = [str(c) for c in up.columns[:MAX_SETS]]
                upl = [up[c].tolist() for c in up.columns[:MAX_SETS]]
            if len(gnames) != n_sets:
                st.info(f"File has {len(gnames)} sets.")
                st.button(f"Use all {len(gnames)} sets from the file", on_click=use_file_sets,
                          args=(gnames, [[x for x in l if isinstance(x, str)] for l in upl]))
            for i in range(min(n_sets, len(gnames))):
                names[i] = gnames[i]
                raw_lists[i] = upl[i]
            st.dataframe(up.head(8), width="stretch", hide_index=True)
    else:
        st.info("No file yet – download the template below, fill it and upload, or just paste into the grid.")

o1, o2, o3, o4 = st.columns([1, 1, 1.2, 1.2])
case_sensitive = o1.toggle("Case-sensitive", False, help="Off: 'tp53' and 'TP53' are the same item.")
numeric = o2.toggle("Numeric matching", True, help="On: '5', '5.0' and '05' are the same item.")


def template_xlsx() -> bytes:
    names_t, lists_t = example_lists(4)
    lists_t = [l[:12] for l in lists_t]
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        pd.DataFrame({nm: l for nm, l in zip(names_t, lists_t)}).to_excel(xw, sheet_name="Sets", index=False)
        pd.DataFrame({"How to use": [
            "One set per column; the header row is the set name.",
            "Items may be text (gene symbols, IDs) or numbers; blanks are ignored.",
            "Columns may have different lengths.",
            "Copy the whole block (including headers) and paste into the Venn Studio grid, "
            "ticking 'First row contains set names' – or upload this file."]}).to_excel(
            xw, sheet_name="Instructions", index=False)
    return buf.getvalue()


o3.download_button("⬇ Excel template", template_xlsx(), "venn_template.xlsx",
                   width="stretch")
o4.download_button("⬇ CSV template",
                   "Set A,Set B,Set C\nTP53,TP53,EGFR\nEGFR,MYC,MYC\nKRAS,BRCA1,KRAS\n",
                   "venn_template.csv", "text/csv", width="stretch")

data = ve.build_sets(names, raw_lists, case_sensitive=case_sensitive, numeric=numeric)
sizes = [len(s) for s in data.sets]
summary = " · ".join(f"**{nm}** {sz:,}" for nm, sz in zip(names, sizes))
dups = sum(data.duplicates)
st.markdown(f"{summary} · union **{len(data.union):,}** unique"
            + (f" · <span style='color:#b45309'>{dups} duplicate entr{'y' if dups == 1 else 'ies'} "
               f"ignored</span>" if dups else ""), unsafe_allow_html=True)

# --------------------------------------------------------------------------- #
# Step 2 – diagram
# --------------------------------------------------------------------------- #
st.subheader("2 · Diagram")
kinds = []
if n_sets <= 5:
    kinds.append("Classic Venn")
if n_sets <= 3:
    kinds.append("Area-proportional (Euler)")
kinds.append("UpSet plot")
k1, k2 = st.columns([2.2, 3])
with k1:
    ss_default("kind", kinds[0])
    if st.session_state.kind not in kinds:
        st.session_state.kind = kinds[0]
    kind = st.segmented_control("Diagram type", kinds, key="kind", selection_mode="single") or kinds[0]
with k2:
    tips = {
        "Classic Venn": "Every possible region is drawn with equal room – the standard for 2–5 sets. "
                        "4–5 sets use the symmetric ellipse layouts seen in VennDiagram / jvenn.",
        "Area-proportional (Euler)": "Circle areas scale with set size and pairwise overlaps "
                                     "(BioVenn / eulerr style). The 3-way region is approximate.",
        "UpSet plot": "Recommended for ≥ 4 sets (UpSetR / Intervene style): bars show exclusive "
                      "intersection sizes, dots show which sets are involved.",
    }
    st.caption(tips[kind])

style = ve.Style(
    width=int(round(st.session_state.w_in * 96)),
    height=int(round(st.session_state.h_in * 96)),
    colors=[st.session_state[f"col{i}"] for i in range(n_sets)],
    opacity=st.session_state.opacity,
    fill_style=st.session_state.fill_style,
    stroke_mode=st.session_state.stroke_mode,
    stroke_w=st.session_state.stroke_w,
    font=st.session_state.font,
    count_pt=st.session_state.count_pt,
    label_pt=st.session_state.label_pt,
    title_pt=st.session_state.title_pt,
    title=st.session_state.get("title", ""),
    count_color=st.session_state.count_color,
    label_color="auto" if st.session_state.label_auto else st.session_state.get("label_color", "#222222"),
    count_mode=st.session_state.count_mode,
    max_items=st.session_state.get("max_items", 6),
    bold_counts=st.session_state.bold_counts,
    bold_labels=st.session_state.bold_labels,
    show_set_size=st.session_state.show_n,
    thousands=st.session_state.thousands,
    auto_contrast=st.session_state.auto_contrast,
    emphasise_core=st.session_state.emph_core,
    hide_zero=st.session_state.hide_zero,
    legend=st.session_state.legend_mode,
    background="white" if st.session_state.bg == "White" else "transparent",
)

warnings: list[str] = []
if len(data.union) == 0:
    st.info("Add some items to see the diagram.")
    st.stop()

if kind == "UpSet plot":
    with st.expander("UpSet options", expanded=False):
        u1, u2, u3 = st.columns(3)
        uo = ve.UpsetOpts(
            sort_by=u1.radio("Sort bars by", ["Size", "Degree"], horizontal=True),
            max_bars=u2.slider("Max bars", 3, 60, 25),
            min_size=u3.number_input("Min intersection size", 1, 10_000, 1),
            bar_color=u1.color_picker("Bar / dot colour", "#3C3C3C"),
            color_singletons=u2.checkbox("Colour single-set bars", True),
            show_values=u3.checkbox("Values on bars", True),
            show_set_sizes=u3.checkbox("Set-size bars", True),
        )
    svg = ve.render_upset(data, style, uo)
else:
    svg, warnings = ve.render_venn(data, style, proportional=kind.startswith("Area"))
for w in warnings:
    st.warning(w)


def preview_component(svg_text: str, w: int, h: int, dpi: int, white: bool, fname: str):
    b64 = base64.b64encode(svg_text.encode("utf-8")).decode()
    scale = dpi / 96
    box_h = 560
    html = f"""
    <div style="font-family:system-ui,-apple-system,Segoe UI,sans-serif">
      <div id="stage" style="height:{box_h}px;display:flex;align-items:center;justify-content:center;
           border:1px solid #e5e7eb;border-radius:10px;
           background-color:#fff;background-image:
             linear-gradient(45deg,#f3f4f6 25%,transparent 25%),linear-gradient(-45deg,#f3f4f6 25%,transparent 25%),
             linear-gradient(45deg,transparent 75%,#f3f4f6 75%),linear-gradient(-45deg,transparent 75%,#f3f4f6 75%);
           background-size:20px 20px;background-position:0 0,0 10px,10px -10px,-10px 0">
        <img id="fig" src="data:image/svg+xml;base64,{b64}"
             style="max-width:98%;max-height:{box_h - 16}px;"/>
      </div>
      <div style="display:flex;gap:8px;margin-top:10px;align-items:center;flex-wrap:wrap">
        <button onclick="dlPng()" style="{BTN}">⬇ PNG · {dpi} dpi</button>
        <button onclick="copyPng()" style="{BTN}">⧉ Copy PNG (paste into PPT)</button>
        <button onclick="dlSvg()" style="{BTN}">⬇ SVG (vector)</button>
        <span id="msg" style="color:#6b7280;font-size:13px"></span>
      </div>
    </div>
    <script>
      const W={w}, H={h}, S={scale}, WHITE={json.dumps(white)};
      const svgB64="{b64}";
      function msg(t){{document.getElementById('msg').textContent=t;}}
      function toCanvas(){{
        return new Promise((res,rej)=>{{
          const img=new Image();
          img.onload=()=>{{
            const c=document.createElement('canvas');
            c.width=Math.round(W*S); c.height=Math.round(H*S);
            const ctx=c.getContext('2d');
            if(WHITE){{ctx.fillStyle='#fff';ctx.fillRect(0,0,c.width,c.height);}}
            ctx.drawImage(img,0,0,c.width,c.height); res(c);
          }};
          img.onerror=rej; img.src='data:image/svg+xml;base64,'+svgB64;
        }});
      }}
      function save(blob,name){{
        const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=name;
        document.body.appendChild(a); a.click(); a.remove();
      }}
      async function dlPng(){{ const c=await toCanvas(); c.toBlob(b=>save(b,'{fname}.png'),'image/png'); msg('PNG saved'); }}
      function dlSvg(){{ save(new Blob([atob(svgB64)],{{type:'image/svg+xml'}}),'{fname}.svg'); msg('SVG saved'); }}
      async function copyPng(){{
        try{{ const c=await toCanvas();
          const blob=await new Promise(r=>c.toBlob(r,'image/png'));
          await navigator.clipboard.write([new ClipboardItem({{'image/png':blob}})]);
          msg('Copied – Ctrl+V into PowerPoint');
        }}catch(e){{ msg('Clipboard blocked by browser – use the PNG/SVG download instead'); }}
      }}
    </script>
    """
    if hasattr(st, "iframe"):
        st.iframe(html, height=box_h + 70)
    else:  # older Streamlit
        components.html(html, height=box_h + 70)


BTN = ("padding:7px 14px;border:1px solid #d1d5db;border-radius:8px;background:#fff;"
       "cursor:pointer;font-size:13px;font-weight:600;color:#111827")
fname = "venn_" + "_".join(n.replace(" ", "")[:10] for n in names)[:60]
preview_component(svg, style.width, style.height, st.session_state.dpi,
                  style.background == "white", fname)

d1, d2, d3 = st.columns([1, 1, 2.4])
d1.download_button("⬇ SVG for PowerPoint", svg.encode("utf-8"), f"{fname}.svg", "image/svg+xml",
                   width="stretch", type="primary")
try:
    import cairosvg  # optional

    d2.download_button("⬇ PDF (vector)", cairosvg.svg2pdf(bytestring=svg.encode()), f"{fname}.pdf",
                       "application/pdf", width="stretch")
except Exception:  # noqa: BLE001
    d2.caption("Vector PDF export: `pip install cairosvg` (optional).")
d3.info("**PowerPoint:** *Insert → Pictures → This Device → .svg*. It stays razor-sharp at any size. "
        "Right-click → **Convert to Shape** to edit text, colours and circles natively.", icon="💡")

# --------------------------------------------------------------------------- #
# Step 3 – intersections
# --------------------------------------------------------------------------- #
st.subheader("3 · Intersections")
regions = data.regions()
union_n = max(1, len(data.union))
rows = []
for mask in sorted(range(1, 1 << n_sets), key=lambda m: (-bin(m).count("1"), m)):
    items = regions.get(mask, [])
    rows.append({
        "Region": data.region_name(mask),
        **{nm: ("●" if mask >> i & 1 else "") for i, nm in enumerate(names)},
        "Degree": bin(mask).count("1"),
        "Count": len(items),
        "% of union": round(100 * len(items) / union_n, 2),
        "Items": "; ".join(items),
    })
reg_df = pd.DataFrame(rows)
t1, t2 = st.columns([3, 2])
with t1:
    show_empty = st.checkbox("Show empty regions", False)
    view = reg_df if show_empty else reg_df[reg_df["Count"] > 0]
    st.dataframe(view, width="stretch", hide_index=True, height=360,
                 column_config={"Items": st.column_config.TextColumn(width="large"),
                                "% of union": st.column_config.NumberColumn(format="%.1f%%")})
with t2:
    nonempty = [r["Region"] for r in rows if r["Count"] > 0]
    pick = st.selectbox("Items in region", nonempty) if nonempty else None
    if pick:
        its = reg_df.loc[reg_df["Region"] == pick, "Items"].iloc[0].split("; ")
        st.text_area(f"{len(its)} items – copy with Ctrl+A, Ctrl+C", "\n".join(its), height=280)

# membership matrix + pairwise overlaps
membership = data.membership()
mem_df = pd.DataFrame(
    [{"Item": data.display.get(k, k), **{nm: int(m >> i & 1) for i, nm in enumerate(names)}}
     for k, m in membership.items()]).sort_values("Item", key=lambda s: s.map(ve._natural_key))
pair = pd.DataFrame(index=names, columns=names, dtype=object)
jac = pd.DataFrame(index=names, columns=names, dtype=float)
for i in range(n_sets):
    for j in range(n_sets):
        inter = len(data.sets[i] & data.sets[j])
        uni = len(data.sets[i] | data.sets[j]) or 1
        pair.iloc[i, j] = inter
        jac.iloc[i, j] = round(inter / uni, 3)

with st.expander("Pairwise overlap & Jaccard index"):
    p1, p2 = st.columns(2)
    p1.caption("Shared items (any overlap, not exclusive)")
    p1.dataframe(pair, width="stretch")
    p2.caption("Jaccard index |A∩B| / |A∪B|")
    p2.dataframe(jac, width="stretch")


def results_xlsx() -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        reg_df.to_excel(xw, sheet_name="Intersections", index=False)
        mem_df.to_excel(xw, sheet_name="Membership matrix", index=False)
        pair.to_excel(xw, sheet_name="Pairwise overlap")
        jac.to_excel(xw, sheet_name="Jaccard")
        pd.DataFrame({nm: pd.Series(sorted((data.display[k] for k in s), key=ve._natural_key))
                      for nm, s in zip(names, data.sets)}).to_excel(xw, sheet_name="Input sets (clean)", index=False)
    return buf.getvalue()


e1, e2, _ = st.columns([1, 1, 2])
e1.download_button("⬇ All results (Excel)", results_xlsx(), "venn_results.xlsx",
                   width="stretch")
e2.download_button("⬇ Intersections (CSV)", reg_df.to_csv(index=False).encode(),
                   "venn_intersections.csv", "text/csv", width="stretch")
