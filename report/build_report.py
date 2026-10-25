"""Build the report, Report_Group8_DatasetC_6bit.docx (at the root of the submission), from the results files.

    python -m pip install python-docx
    python report/build_report.py

Every number in the report is read from results/ (tables, CSVs and the generated .md files)
or computed from data/; nothing is typed by hand.  Statements that depend on the data are
guarded by check(): if a rerun changes a result so that a sentence would become wrong, the
build stops and names the sentence instead of writing it.  Run run_all.py and the studies
first (README.md), then this script.

Markup in the text: **bold**, `code`, x_[sub] for a subscript and x^[sup] for a superscript.
"""
import csv
import os
import re
import sys

import numpy as np
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
RES, FIG, TAB = os.path.join(ROOT, "results"), os.path.join(ROOT, "figures"), os.path.join(ROOT, "results", "tables")
OUT = os.path.join(ROOT, "Report_Group8_DatasetC_6bit.docx")
ZIP = "CS611_A1_Group8_DatasetC_6bit.zip"
sys.path.insert(0, ROOT)
import selector as S                      # noqa: E402  (method name, roles)
import selector_suite as SUITE            # noqa: E402  (activation families)

NAMES = ["MXFP6_E3M2", "MXFP6_E2M3", "MXFP6_E4M1", "MXINT6", "BFP6"]
SHORT = {"MXFP6_E3M2": "E3M2", "MXFP6_E2M3": "E2M3", "MXFP6_E4M1": "E4M1", "MXINT6": "MXINT6", "BFP6": "BFP6"}
LONG = {"MXFP6_E3M2": "MXFP6 E3M2", "MXFP6_E2M3": "MXFP6 E2M3", "MXFP6_E4M1": "MXFP6 E4M1", "MXINT6": "MXINT6",
        "BFP6": "BFP6"}
MEMBERS = [("Aryan Kumar", "261110012"), ("Akash Kumar", "261110008"), ("Nikita Kumari Pandit", "261110078"),
           ("Siddhant Singh Rathoria", "261110062"), ("Anuj Chaudhary", "210167")]
REPO = "https://github.com/akashiitkanpur/cs611-g8"
INK, MUTED, ACCENT, METHOD_FILL, HEAD_FILL, NOTE_FILL = "1F2328", "57606A", "B3471D", "FCE9DC", "E8EEF6", "F4F4F2"
WC, XC, XT = "W_[C]", "X_[cal]", "X_[test]"


def check(cond, what):
    if not cond:
        raise SystemExit(f"build_report: the results no longer support this statement, rewrite it: {what}")


# --------------------------------------------------------------------------- reading results
def read_csv(path):
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def md_tables(name):
    """All pipe tables of a results .md file as (header, rows) with cells stripped of ** and `."""
    out, cur = [], None
    for line in open(os.path.join(RES, name), encoding="utf-8"):
        line = line.rstrip("\n")
        if line.startswith("|"):
            line = line.replace("\\|", "¦")
            line = re.sub(r"(?<=\d)\|(?=\d)", "¦", line)          # 16|4|5 inside a cell
            cells = [c.strip().replace("**", "").replace("`", "").replace("¦", "|") for c in line.strip("|").split("|")]
            if cur is None:
                cur = (cells, [])
            elif not all(set(c) <= set("-: ") for c in cells):
                cur[1].append(cells)
        elif cur is not None:
            out.append(cur)
            cur = None
    if cur is not None:
        out.append(cur)
    return out


def md_table(name, first, has=None, index=0):
    """The index-th table whose first header cell is `first` (and which has a column `has`), as dict rows."""
    hits = [(h, r) for h, r in md_tables(name) if h[0] == first and (has is None or has in h)]
    check(len(hits) > index, f"{name}: table starting with '{first}'" + (f" with column '{has}'" if has else ""))
    h, rows = hits[index]
    return h, [dict(zip(h, r)) for r in rows]


def num(s):
    m = re.search(r"[-+]?\d+(?:\.\d+)?", s.replace(",", ""))
    check(m is not None, f"a number in '{s}'")
    return float(m.group(0))


def ci(s):
    """'0.383 [0.325, 0.440]' -> (0.383, 0.325, 0.440)."""
    v = [float(x) for x in re.findall(r"[-+]?\d+\.\d+", s)]
    check(len(v) == 3, f"an interval in '{s}'")
    return tuple(v)


def f2(x):
    return f"{x:.2f}"


def f3(x):
    return f"{x:.3f}"


def pct(x, d=1):
    return f"{100 * x:.{d}f}%"


def winner_of(cell):
    """'E2M3 (10/10) +0.31 ± 0.00' -> 'E2M3'."""
    return cell.split(" ")[0]


# --------------------------------------------------------------------------- the data behind the report
T1 = {r["format"]: r for r in read_csv(os.path.join(TAB, "T1_main_C.csv"))}
T2 = {r["format"]: r for r in read_csv(os.path.join(TAB, "T2_storage.csv"))}
T3 = {r["variant"]: r for r in read_csv(os.path.join(TAB, "T3_e4m1_conventions.csv"))}
T4 = read_csv(os.path.join(TAB, "T4_bfp_family.csv"))
T6 = {r["format"]: r for r in read_csv(os.path.join(TAB, "T6_diagnostics_ocp.csv"))}
T7 = {(r["dataset"], r["mode"], r["policy"]): r for r in read_csv(os.path.join(TAB, "T7_datasets_A_to_E.csv"))}
T8 = {(r["dataset"], r["policy"]): r for r in read_csv(os.path.join(TAB, "T8_mxint6_sym_vs_tc.csv"))}
T9 = {r["policy"]: r for r in read_csv(os.path.join(TAB, "T9_wx_ranking.csv"))}
WX = "W+X held-out Y err % ({})"
E = {n: float(T1[n][WX.format("ocp")]) for n in NAMES}                 # primary: W and X quantized, class rule
EW = {n: float(T1[n]["W only held-out Y err % (ocp)"]) for n in NAMES}  # secondary: weights only
RANK, RANK_W = sorted(NAMES, key=E.get), sorted(NAMES, key=EW.get)
WIN = RANK[0]

W_C = np.load(os.path.join(ROOT, "data", "W_C.npy")).astype(np.float64)
X_CAL = np.load(os.path.join(ROOT, "data", "X_cal.npy")).astype(np.float64)
SIG = W_C.std()
KURT = float(((W_C - W_C.mean()) ** 4).mean() / W_C.var() ** 2)
MAXS = float(np.abs(W_C).max() / SIG)
XSIG = X_CAL.std()
XKURT = float(((X_CAL - X_CAL.mean()) ** 4).mean() / X_CAL.var() ** 2)
XMAXS = float(np.abs(X_CAL).max() / XSIG)
SV = np.linalg.svd(X_CAL, compute_uv=False)
SV_RATIO = float(SV[0] / SV[-1])

# scale study (weights only)
_, SC2A = md_table("scale_study.md", "format", has="bound_y")
SC2A = {r["format"].split(" ")[0]: r for r in SC2A if r["format"].split(" ")[0] in SHORT.values()}
_, SC_RANK = md_table("scale_study.md", "policy", has="verdict")
SC_RANK = {r["policy"]: r for r in SC_RANK}
_, SC_GAIN = md_table("scale_study.md", "format", has="opt_wide", index=1)
SC_GAIN = {r["format"].split(" ")[0]: r for r in SC_GAIN}
_, SC_SRC = md_table("scale_study.md", "format", has="clip")
SRC = {(r["format"].split(" ")[0], r["policy"]): r for r in SC_SRC}
_, SC_FIT = md_table("scale_study.md", "format", has="opt (even / odd)")
SC_FIT = {r["format"].split(" ")[0]: r for r in SC_FIT}

BENCH = {m: read_csv(os.path.join(RES, f"selector_bench_{m}.csv")) for m in ("wx", "w")}


def bench(suite, approach, xkind=None, mode="wx"):
    R = [r for r in BENCH[mode] if r["suite"] == suite and r["approach"] == approach
         and (xkind is None or r["matrix"].endswith("X=" + xkind))]
    check(R, f"selector_bench_{mode}.csv rows for {suite}/{approach}")
    rg = np.array([float(r["regret_y_db"]) for r in R])
    return dict(n=len(R), top1=np.mean([int(r["hit"]) for r in R]), regret=rg.mean(), p95=np.percentile(rg, 95),
                max=rg.max(), ms=1000 * np.median([float(r["seconds"]) for r in R]), rows=R,
                frac=np.median([float(r["sample_frac"]) for r in R]) if R[0]["sample_frac"] else None)


M = S.METHOD
BM, BMW = bench("random", M), bench("random", M, mode="w")
XKINDS = [k for k in SUITE.XKINDS if any(r["matrix"].endswith("X=" + k) for r in BENCH["wx"] if r["suite"] == "random")]
N_RANDOM = len({r["matrix"] for r in BENCH["wx"] if r["suite"] == "random"})
N_ADV = len({r["matrix"] for r in BENCH["wx"] if r["suite"] == "adversarial"})
DS = {r["matrix"].split()[-1]: r for r in bench("datasets", M)["rows"]}
DS_Y = {r["matrix"].split()[-1]: r for r in bench("datasets", "exhaustive_y")["rows"]}

_, SCALING = md_table("selector_scaling_wx.md", "size")
Q3 = "q3_study_wx.md"
_, OUT_T = md_table(Q3, "n, k, magnitudes")
OUT_T = {r["n, k, magnitudes"]: r for r in OUT_T}
_, MECH = md_table(Q3, "", has="E2M3")
MECH = {r[""]: r for r in MECH}
_, SPARSE = md_table(Q3, "zeros", has="winner ocp (seeds), margin pp")
SPARSE = {r["zeros"]: r for r in SPARSE}
_, SPARSE_M = md_table(Q3, "zeros", has="ρ32")
SPARSE_M = {r["zeros"]: r for r in SPARSE_M}
_, SCALEV = md_table(Q3, "spread")
OUT_W = {r["n, k, magnitudes"]: r for r in md_table("q3_study_w.md", "n, k, magnitudes")[1]}
SPARSE_W = {r["zeros"]: r for r in md_table("q3_study_w.md", "zeros", has="winner ocp (seeds), margin pp")[1]}
_, BLK = md_table("block_study_wx.md", "block length")
BLK = {r["block length"]: r for r in BLK}
_, BUDGET = md_table("block_study_wx.md", "budget (bits/elem)")
BUDGET = {r["budget (bits/elem)"]: r for r in BUDGET}
_, BLK_AE = md_table("block_study_wx.md", "dataset", has="class layouts winner")
_, MIXED = md_table("block_study_w.md", "dataset", has="per-block mix %")
BUDGET_W = {r["budget (bits/elem)"]: r for r in md_table("block_study_w.md", "budget (bits/elem)")[1]}


def x_stress(mode):
    name = f"x_stress_{mode}.md"
    per = {k: {r["X family"]: r for r in md_table(name, "X family", index=i)[1]} for i, k in enumerate("ABCDE")}
    h, shift = md_table(name, "X_test family")
    return per, h, shift


XS, SHIFT_H, SHIFT = x_stress("wx")
XS_W, SHIFT_H_W, SHIFT_W = x_stress("w")


def n_tests(path):
    return sum(1 for line in open(os.path.join(ROOT, path), encoding="utf-8") if line.startswith("def test_"))


# --------------------------------------------------------------------------- document helpers
d = Document()
sec = d.sections[0]
sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
for side in ("left_margin", "right_margin"):
    setattr(sec, side, Cm(2.0))
sec.top_margin, sec.bottom_margin = Cm(1.8), Cm(1.8)
st = d.styles["Normal"]
st.font.name, st.font.size = "Calibri", Pt(10.5)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
st.paragraph_format.space_after = Pt(4)
st.paragraph_format.line_spacing = 1.1
for lvl, size in ((1, 15), (2, 12), (3, 11)):
    hs = d.styles[f"Heading {lvl}"]
    hs.font.name, hs.font.size, hs.font.color.rgb = "Calibri", Pt(size), RGBColor.from_string(INK)
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):   # let "Calibri" apply
        hs.element.rPr.rFonts.attrib.pop(qn(attr), None)
    hs.paragraph_format.space_before, hs.paragraph_format.space_after = Pt(14 if lvl == 1 else 9), Pt(4)
FIGNO, TABNO, HEADS = [0], [0], []
TOKEN = re.compile(r"(\*\*|`[^`]+`|_\[[^\]]*\]|\^\[[^\]]*\])")


def rich(p, text, size=None, color=None, italic=False, bold=False, font=None):
    """Add text to paragraph p: **bold** toggles, `code`, x_[sub], x^[sup]."""
    b = bold
    for part in TOKEN.split(text):
        if not part:
            continue
        if part == "**":
            b = not b
            continue
        if part.startswith("`"):
            r = p.add_run(part[1:-1])
            r.font.name = "Consolas"
            r.font.size = Pt((size or 10.5) - 1)
        else:
            if part.startswith("_[") or part.startswith("^["):
                r = p.add_run(part[2:-1])
                r.font.subscript, r.font.superscript = part[0] == "_", part[0] == "^"
            else:
                r = p.add_run(part)
            if size:
                r.font.size = Pt(size)
            if font:
                r.font.name = font
                r._element.rPr.rFonts.set(qn("w:hAnsi"), font)
        if b:
            r.bold = True
        if color:
            r.font.color.rgb = RGBColor.from_string(color)
        if italic:
            r.italic = True
    return p


def P(text, **kw):
    return rich(d.add_paragraph(), text, **kw)


def H(text, level=1):
    p = d.add_heading("", level)
    rich(p, text)
    i = len(HEADS) + 1
    name = f"_Toc{i:06d}"
    s, e = OxmlElement("w:bookmarkStart"), OxmlElement("w:bookmarkEnd")
    s.set(qn("w:id"), str(i))
    s.set(qn("w:name"), name)
    e.set(qn("w:id"), str(i))
    p._p.pPr.addnext(s)
    p._p.append(e)
    HEADS.append((level, text, name))
    return p


def bullets(items, style="List Bullet"):
    for it in items:
        p = d.add_paragraph(style=style)
        rich(p, it)
        p.paragraph_format.space_after = Pt(2)


def eq(text):
    p = d.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(3), Pt(6)
    rich(p, text, font="Cambria Math", size=11)
    return p


def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def caption(kind, text):
    cnt = FIGNO if kind == "Figure" else TABNO
    cnt[0] += 1
    p = d.add_paragraph()
    rich(p, f"**{kind} {cnt[0]}.** {text}", size=9, color=MUTED)
    p.paragraph_format.space_after = Pt(8)
    if kind == "Table":
        p.paragraph_format.keep_with_next = True
    return cnt[0]


def table(header, rows, cap=None, highlight=(), widths=None, size=8.5, bold_first=False):
    """rows: list of lists of str.  highlight: row indices filled as 'the chosen method / the answer'."""
    n = caption("Table", cap) if cap else None
    t = d.add_table(rows=1 + len(rows), cols=len(header))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate([header] + rows):
        for j, v in enumerate(row):
            c = t.cell(i, j)
            c.text = ""
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            hl = i > 0 and i - 1 in highlight
            rich(p, str(v), size=size, bold=(i == 0 or hl or (bold_first and j == 0)))
            if i == 0:
                shade(c, HEAD_FILL)
            elif hl:
                shade(c, METHOD_FILL)
            if widths:
                c.width = Cm(widths[j])
    tr = t.rows[0]._tr.get_or_add_trPr()
    hdr = OxmlElement("w:tblHeader")
    hdr.set(qn("w:val"), "true")
    tr.append(hdr)
    d.add_paragraph().paragraph_format.space_after = Pt(2)
    return n


def figure(name, cap, width=16.5):
    path = os.path.join(FIG, name)
    check(os.path.exists(path), f"figure {name}")
    d.add_picture(path, width=Cm(width))
    d.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    d.paragraphs[-1].paragraph_format.keep_with_next = True
    return caption("Figure", cap)


def box(title, lines, fill=NOTE_FILL):
    t = d.add_table(rows=1, cols=1)
    t.style = "Table Grid"
    c = t.cell(0, 0)
    shade(c, fill)
    p = c.paragraphs[0]
    rich(p, f"**{title}**", size=11)
    for ln in lines:
        p = c.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        rich(p, ln, size=10)
    d.add_paragraph().paragraph_format.space_after = Pt(2)


def field_run(p, kind=None, instr=None, size=None):
    r = OxmlElement("w:r")
    if size:
        rpr = OxmlElement("w:rPr")
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str(int(size * 2)))
        rpr.append(sz)
        r.append(rpr)
    if kind:
        el = OxmlElement("w:fldChar")
        el.set(qn("w:fldCharType"), kind)
    else:
        el = OxmlElement("w:instrText")
        el.set(qn("xml:space"), "preserve")
        el.text = instr
    r.append(el)
    p._p.append(r)


def page_number_footer():
    p = sec.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rich(p, "CS611 Assignment 1 · Group 8 · Dataset C, 6-bit · page ", size=8, color=MUTED)
    field_run(p, "begin", size=8)
    field_run(p, instr="PAGE", size=8)
    field_run(p, "end", size=8)


def build_toc(anchor):
    """A Word TOC field over Heading 1-2.  Its cached entries (shown before Word updates the
    field) are the headings, linked to their bookmarks; Word adds page numbers on opening."""
    for i, (lvl, text, name) in enumerate(HEADS):
        p = d.add_paragraph()
        pf = p.paragraph_format
        pf.left_indent = Cm(0 if lvl == 1 else 0.7)
        pf.space_before, pf.space_after = Pt(5 if lvl == 1 else 0), Pt(1)
        if i == 0:
            field_run(p, "begin")
            field_run(p, instr=' TOC \\o "1-2" \\h \\z \\u ')
            field_run(p, "separate")
        hl = OxmlElement("w:hyperlink")
        hl.set(qn("w:anchor"), name)
        hl.set(qn("w:history"), "1")
        tmp = rich(d.add_paragraph(), text, size=10.5 if lvl == 1 else 10, bold=(lvl == 1))
        for r in tmp._p.findall(qn("w:r")):
            hl.append(r)
        tmp._p.getparent().remove(tmp._p)
        p._p.append(hl)
        if i == len(HEADS) - 1:
            field_run(p, "end")
        anchor._p.addprevious(p._p)
    anchor._p.getparent().remove(anchor._p)
    settings = d.settings.element                        # ask Word to fill in the page numbers on opening
    uf = OxmlElement("w:updateFields")
    uf.set(qn("w:val"), "true")
    later = {"hdrShapeDefaults", "footnotePr", "endnotePr", "compat", "docVars", "rsids", "mathPr", "attachedSchema",
             "themeFontLang", "clrSchemeMapping", "doNotIncludeSubdocsInStats", "doNotAutoCompressPictures",
             "forceUpgrade", "captions", "readModeInkLockDown", "smartTagType", "schemaLibrary", "shapeDefaults",
             "doNotEmbedSmartTags", "decimalSymbol", "listSeparator"}
    for child in settings:
        if child.tag.split("}")[-1] in later:
            child.addprevious(uf)
            break
    else:
        settings.append(uf)


page_number_footer()
for z in d.settings.element.findall(qn("w:zoom")):     # the default template's zoom lacks the required percent
    z.set(qn("w:percent"), "100")

# =========================================================================== title page
p = d.add_paragraph()
p.paragraph_format.space_before = Pt(36)
rich(p, "Choosing a 6-bit block format for a linear layer", size=22, bold=True, color=INK)
p = d.add_paragraph()
rich(p, "CS611 Assignment 1 · Group 8 · Dataset C · element width 6 bits", size=13, color=MUTED)
d.add_paragraph()
table(["Group member", "Roll number"], [list(m) for m in MEMBERS], widths=[7.0, 4.0], size=10)
P("**About this submission.** This report is the file `Report_Group8_DatasetC_6bit.docx` at the root of the "
  f"submitted zip `{ZIP}`. The zip also contains the code, the data and every result file that the report draws "
  "on. Paths written like `results/tables/T1_main_C.csv` are paths inside the zip, and `README.md` in the zip lists "
  "every file and the commands that regenerate it. All numbers in this report are read from those result files by "
  "`report/build_report.py`; none is typed by hand.", size=10)
P(f"The same code is also in the group's GitHub repository, {REPO} (secondary reference; private, access on request).",
  size=10, color=MUTED)
p = d.add_paragraph(style="TOC Heading")
rich(p, "Contents")
toc_anchor = d.add_paragraph()
d.add_page_break()

# =========================================================================== 1. task and setting
H("1. The task and the setting", 1)
H("1.1 The task", 2)
P(f"A linear layer computes **Y = W·X**. We are given the weight matrix {WC} (100 × 100) and a calibration input "
  f"{XC} (100 × {X_CAL.shape[1]}): each column of X is one input vector, and row j of X holds input channel j. The "
  f"graders evaluate on a separate input {XT} that we do not see. The layer is to be stored in one of five 6-bit "
  "block formats: **MXFP6 E3M2, MXFP6 E2M3, MXFP6 E4M1, MXINT6 and BFP6** (Section 2). The assignment asks three "
  "questions:")
bullets([
    f"**Q1.** Which of the five formats is most suitable for {WC}?",
    "**Q2.** Can an efficient method choose the format as well as exhaustive evaluation (quantizing with every "
    "candidate and measuring the error), without quantizing with every candidate?",
    "**Q3.** How do outliers, sparsity and variation in scale change the choice?",
])
H("1.2 The data", 2)
q = X_CAL.shape[0] / X_CAL.shape[1]
mp_ratio = (1 + np.sqrt(q)) / (1 - np.sqrt(q))          # largest / smallest singular value of white noise
check(SV_RATIO < 1.3 * mp_ratio and abs(XKURT - 3) < 0.3, "X_cal looks like white Gaussian noise")
P(f"{WC} has standard deviation σ = {SIG:.2f}, largest magnitude {np.abs(W_C).max():.2f} ({MAXS:.1f}σ) and kurtosis "
  f"{KURT:.2f} (a Gaussian has 3): its tails are heavier than a Gaussian's, but it has no extreme outliers "
  f"(Figure {FIGNO[0] + 1}). {XC} looks like white Gaussian noise: standard deviation {XSIG:.3f}, largest magnitude "
  f"{XMAXS:.1f}σ, kurtosis {XKURT:.2f}, and its singular values are spread no more than those of white noise of "
  f"the same shape (largest / smallest {SV_RATIO:.1f}, against {mp_ratio:.1f} expected), so its channels are "
  "nearly uncorrelated.")
figure("fig1_weight_distribution.png", f"Distribution of the weights of {WC} (log density) against a Gaussian with "
       "the same mean and standard deviation.", width=12)
H("1.3 Which operands are quantized", 2)
P("**Primary setting: W and X both quantized.** In a layer that runs in a block format, both operands of the "
  "matrix product are stored in that format: W in blocks along its rows, X in blocks down its columns. Both run "
  "along the reduction axis of Y = W·X (the index that the product sums over), which is how MX hardware computes "
  "dot products. The output error then depends on X in two ways: X decides how much each weight error matters, and "
  "X's own rounding error adds a second term. **All headline numbers in this report are for this setting**, "
  "written \"W+X\".")
P("**Secondary setting: weights only.** We also report the error when only W is quantized and X stays in full "
  "precision, for three reasons:")
bullets([
    "**It is how large language models are usually deployed.** Inference memory and bandwidth are dominated by the "
    "weights, so the weights are stored in a narrow format while the activations stay in 16-bit floating point "
    "(weight-only schemes such as GPTQ [6] and AWQ [5]). There, the weight format is the only choice that matters.",
    f"**{XT} is not available.** With weights only, X enters the error only as a weighting, estimated from {XC}. "
    f"With W+X, the rounding error of X itself is part of the result, and it is measured on {XC}, so that setting "
    f"relies more on {XT} resembling {XC}.",
    "**It separates the two effects.** The weight format's own behaviour (Q1's mechanism, Q3) is easier to see "
    "without the activation term on top.",
])
check(RANK[0] == RANK_W[0], "W+X and weights only have the same winner on C")
P(f"For dataset C both settings give the same answer (Section 4); where they differ elsewhere, we say so.")

# =========================================================================== 2. formats, assumptions, measurement
H("2. Formats, assumptions and how we measure", 1)
P("This section defines everything the results depend on: the formats, how the shared exponent is chosen, the "
  "assumptions we made where the class material leaves a choice open, the error measure, and the activation "
  "families used to test robustness.")
H("2.1 How a block format stores a matrix", 2)
P("All five candidates store a matrix in the same way. The values are cut into **blocks** of consecutive values "
  "along the reduction axis (along each row of W; down each column of X). All values of a block share one "
  "power-of-two **scale** 2^[e], and each value is stored as a 6-bit **element** code. A block is quantized in four "
  "steps:")
bullets([
    "Choose the shared exponent e from the largest magnitude in the block (Section 2.3 gives the rules we compare).",
    "Divide every value by 2^[e] and round it to the nearest point of the element grid; a value exactly halfway "
    "between two grid points goes to the one with an even code (\"round half to even\").",
    "A value beyond the largest grid point v_[max] is stored as ±v_[max] (saturation). There are no infinity or "
    "NaN codes.",
    "Store the 6-bit codes and the block's exponent.",
], style="List Number")
P("The candidates differ only in the **element grid** (which values a 6-bit code can represent) and the **block "
  "organisation** (how many values share a scale, and how many bits the scale has). A floating-point (FP) grid has "
  "points spaced proportionally to their size, so small values keep their relative precision; an integer grid has "
  "equally spaced points, so its precision is set by the block maximum.")
H("2.2 The five candidate formats", 2)
cand = [
    ["MXFP6 E3M2", "sign, 3 exponent bits, 2 mantissa bits (bias 3); normal exponents −2 … 4, with subnormals",
     "28", "32", "8-bit (E8M0)", f"{float(T2['MXFP6_E3M2']['bits/elem']):.2f}"],
    ["MXFP6 E2M3", "sign, 2 exponent bits, 3 mantissa bits (bias 1); normal exponents 0 … 2, with subnormals",
     "7.5", "32", "8-bit (E8M0)", f"{float(T2['MXFP6_E2M3']['bits/elem']):.2f}"],
    ["MXFP6 E4M1", "sign, 4 exponent bits, 1 mantissa bit (bias 7); normal exponents −6 … 8; every code finite",
     "384", "32", "8-bit (E8M0)", f"{float(T2['MXFP6_E4M1']['bits/elem']):.2f}"],
    ["MXINT6", "6-bit two's-complement integer k ∈ [−32, 31], value k·2^[−4] (so −2 … 1.9375)",
     "1.9375", "32", "8-bit (E8M0)", f"{float(T2['MXINT6']['bits/elem']):.2f}"],
    ["BFP6 (32|4|5)", "sign and a 5-bit magnitude M ∈ [0, 31], value ±M·2^[E−4]; E = shared exponent",
     "31 (×2^[E−4])", T2["BFP6"]["block"], "4-bit, E ∈ [−8, 7]", f"{float(T2['BFP6']['bits/elem']):.2f}"],
]
table(["Format", "Element (6 bits)", "v_[max] (element)", "Block", "Shared scale", "Bits/element"], cand,
      cap="The five candidates. Bits per element count the codes and the block scales of a 100 × 100 matrix "
          "(results/tables/T2_storage.csv).", widths=[2.6, 7.2, 1.8, 1.2, 2.4, 1.6], size=8.5)
P("**Storage.** A row of W has 100 values. With blocks of 32 it holds three full blocks and a **tail block of 4** "
  "values that gets its own scale (no padding), so 4 scales of 8 bits per row: (100·6 + 4·8)/100 = 6.32 bits per "
  f"element. BFP6 has the same {T2['BFP6']['blocks/row (K=100)']} groups per row, but with 4-bit exponents: "
  f"(100·6 + 4·4)/100 = {float(T2['BFP6']['bits/elem']):.2f}. For comparison we also report BFP with groups of 16 "
  f"(16|4|5, {T2['BFP6_g16']['blocks/row (K=100)']} groups per row, {float(T2['BFP6_g16']['bits/elem']):.2f} bits per "
  "element); it is not a candidate.")
H("2.3 How the shared exponent is chosen: scale policies", 2)
P("The shared exponent is the only free parameter of a block. Every policy below stores a power-of-two scale in the "
  "same field, so none of them changes the storage. e_[max] is the exponent of the largest binade of the element "
  "grid: 4 for E3M2, 2 for E2M3, 8 for E4M1 and 0 for MXINT6 and BFP6.")
pol_rows = [
    ["`ocp` (class rule)", "e = floor(log_[2] max|w|) − e_[max]: the block maximum lands in the top binade of the grid. "
     "This is the rule of the OCP MX specification [2] and the class notebook [1].", "no", "**all headline results**"],
    ["`ceil`", "the smallest power-of-two scale at which no value of the block saturates (for MXINT6 a negative "
     "block maximum may use the code −32)", "no", "comparison"],
    ["`opt`", "per block, try e_[ocp] + d for d ∈ {−2, −1, 0, +1} and keep the exponent with the smallest squared "
     "error of the block; costs four roundings", "no", "comparison"],
    ["`opt_out`", "the same search, but each value's error weighted by the energy of its input channel in the fit "
     "columns of X_[cal] (for X: by the energy of the matching column of W)", "fit columns", "comparison"],
    ["`model`", "the same search with the error predicted from a model instead of rounding", "no",
     "scale study only"],
    ["`oracle_y`", "all exponents of a row chosen together to minimise the output error on the fit columns, "
     "using the full matrix X·X^[T]", "fit columns", "scale study only"],
    ["`bound_y`", "as `oracle_y`, but fitted on the held-out columns themselves: a lower bound on what any exponent "
     "choice can reach, not a usable policy", "held-out columns", "scale study only"],
]
table(["Policy", "Rule for the shared exponent", "Uses X?", "Where used"], pol_rows,
      cap="Scale policies. When X is quantized too, it uses the same policy as W.", widths=[2.3, 10.2, 1.9, 2.4],
      size=8.5)
H("2.4 Assumptions and decisions", 2)
P("Where the class material leaves a choice open, we made the assumptions below. Each is stated so that it can be "
  "checked, and Section 4.4 shows how much the ones that change numbers matter.")
bfp32 = [r for r in T4 if r["format"] == "BFP6_g32e8"][0]
T4N = {r["format"]: r for r in T4}
bfp_eq_sym = T4N["BFP6"]["held-out Y err % (ocp)"] == T4N["MXINT6_sym"]["held-out Y err % (ocp)"]
assum = [
    ["A1", "Block length 32 for every candidate (MX blocks and BFP6 groups), along the rows of W (the reduction "
           "axis). With 100 columns a row has blocks of 32, 32, 32 and a tail of 4, and the tail has its own scale.",
     "The class convention and the OCP MX specification [2]; the course staff advised groups of 32 for this "
     "assignment. Not padding the tail wastes no storage."],
    ["A2", "MX shared scale: an 8-bit power of two (E8M0), exponent clamped to [−127, 127].", "OCP MX specification [2]."],
    ["A3", "Round half to even; saturate at ±v_[max]; no infinity or NaN codes.", "Class notebook [1]; OCP MX conversion rules [2]."],
    ["A4", "MXFP6 E4M1 reserves no codes: all 64 are finite and v_[max] = 384.",
     "The OCP 6-bit formats (E2M3, E3M2) reserve none; we treat E4M1 the same way. Two alternatives are tested in "
     "Section 4.4."],
    ["A5", "MXINT6 elements are two's-complement integers k ∈ [−32, 31] times 2^[−4].",
     "The standard integer encoding. The symmetric range [−31, 31] is tested in Section 4.4."],
    ["A6", "BFP6 = 32|4|5: groups of 32 values, a 4-bit shared exponent E = floor(log_[2] max) clamped to "
           "[−8, 7], and sign + 5-bit magnitude per value.",
     "The class defines BFP by group size | exponent bits | mantissa bits without fixing them; we use groups of "
     "32 (A1). With an 8-bit exponent BFP would be numerically identical to symmetric MXINT6; the 4-bit exponent "
     f"makes it a distinct format, with cheaper scales ({float(T2['BFP6']['bits/elem']):.2f} bits per element) and a "
     "limited range. BFP with groups of 16 is reported as a comparison row (Section 4.4)."],
    ["A7", "One format for the whole matrix; no mixing of formats between blocks.",
     "The question asks for one format per matrix. Section 7.4 checks what mixing would gain."],
    ["A8", "Primary setting W+X; weights only as a secondary setting.", "Section 1.3."],
    ["A9", "X is quantized with the same format as W, in blocks down its columns (one block = consecutive input "
           "channels of one input vector), with the same scale policy.",
     "Down the columns is the reduction axis of X in Y = W·X; one format for the layer."],
    ["A10", "Error measure: relative Frobenius error of the output (Section 2.5).",
     "The output is what the next layer sees; relative errors are comparable across datasets."],
    ["A11", f"The {X_CAL.shape[1]} columns of {XC} are split: even columns (\"fit columns\") for anything fitted to "
            "activations, odd columns (\"held-out columns\") for every reported error.",
     f"It imitates evaluation on the unseen {XT}: nothing is scored on the data it was tuned on."],
    ["A12", "Ties: two formats tie if the 95% interval of their error gap contains zero (paired bootstrap over the "
            "held-out columns, 2000 resamples).", "Avoids declaring a winner on a difference within sampling noise."],
    ["A13", "The data are float32; quantization and products are computed in float64.",
     "Float32 products agree to better than 10^[−6] relative, far below every gap we report."],
    ["A14", "Storage is counted in real bits per element, including the tail block's scale.", "Section 2.2."],
    ["A15", "Datasets A, B, D and E (other groups' matrices, same X_[cal]) are used only as context.",
     "They show whether a conclusion depends on the data."],
    ["A16", "Headline numbers use the class scale rule `ocp`; the other scale policies are reported alongside.",
     "It is the rule of the specification and the class. Section 5 shows that the ranking does not depend on it."],
    ["A17", "`ceil` is read literally: the smallest power-of-two scale at which nothing saturates, so for MXINT6 a "
            "negative block maximum may use the code −32.", "The literal meaning of \"clips nothing\" on a "
                                                            "two's-complement grid."],
]
table(["#", "Assumption or decision", "Reason"], assum, cap="Assumptions and decisions. Comments in the code cite "
      "them by number (for example `[A5]` in `formats.py`).", widths=[1.0, 8.2, 7.6], size=8.5)
H("2.5 How we measure", 2)
P("**Output error.** For a format f, W_[q] and X_[q] are W and X quantized with f. On the held-out columns of "
  f"{XC} we report the relative error of the layer output:")
eq("e_[Y] = ‖W·X − W_[q]·X_[q]‖_[F] / ‖W·X‖_[F]")
P("(with weights only, X_[q] = X). ‖·‖_[F] is the Frobenius norm, the square root of the sum of squares. We give "
  "e_[Y] in percent; gaps between formats are in percentage points (pp).")
bullets([
    "**Top-1 agreement** (Q2): a selection method is right on a matrix if its pick is the format with the smallest "
    "held-out error.",
    "**Regret** (Q2): when the pick is wrong, how much worse it is: 10·log_[10] of the error power of the pick over "
    "that of the best format, in dB. 0 dB means the pick is the best format; 0.1 dB is about 1.2% more error power.",
    "**Margin** (Q1, Q3): the runner-up's error minus the winner's, in pp. Where we repeat an experiment over random "
    "seeds, \"(7/10)\" means the format won in 7 of 10 seeds.",
])
H("2.6 Activation families used to test robustness", 2)
P(f"Because {XT} is unseen, we test whether the answers survive other activations. Besides the real {XC}, we use "
  "a bootstrap of it and nine synthetic families. Each synthetic X has one column per input vector and one row per "
  "input channel; Q1 uses 5 random seeds per family, and each matrix of the Q2 test suite gets one family at random.")
XDESC = {
    "X_cal (real)": ("the calibration input as given", "the real data"),
    "X_cal bootstrap": (f"the columns of {XC} resampled with replacement", f"another sample like {XC}"),
    "white": ("independent standard Gaussian entries", "the simplest baseline"),
    "chan_gain": ("white, with each channel multiplied by its own gain exp(N(0, 1))",
                  "channels of different scale, e.g. after normalisation layers with learned gains"),
    "outlier_chan": ("white, with 1/32 of the channels (3 of 100) multiplied by one gain between 10 and 100",
                     "a few much larger channels, as reported for LLM activations [7]"),
    "lowrank": ("a random product of rank 1 to 5, plus 0.1 × white noise", "strongly correlated channels"),
    "heavy_t3": ("independent Student-t entries with 3 degrees of freedom", "heavy-tailed tokens"),
    "ar1": ("channel j = 0.9 × channel j−1 + √0.19 × white noise", "neighbouring channels correlated (ρ = 0.9)"),
    "relu": ("max(white + 0.3, 0)", "rectified activations: non-negative, non-zero mean, 38% zeros"),
    "massive": ("white, with 1% of the channels (at least one) multiplied by 100",
                "the \"massive activations\" of large language models [8]"),
    "sparse_act": ("white, with each entry set to zero with probability 0.7", "sparse activations"),
}
XNAME = {"X_cal (real)": "X_[cal] (real)", "X_cal bootstrap": "X_[cal] bootstrap", "white": "white",
         "chan_gain": "channel gains", "outlier_chan": "outlier channels", "lowrank": "low rank",
         "heavy_t3": "Student-t3", "ar1": "AR(1)", "relu": "ReLU", "massive": "massive", "sparse_act": "70% zeros"}
check(set(XDESC) == set(XS["C"]) and set(XKINDS) <= set(XDESC), "the activation-family descriptions cover every family")
table(["Family", "How X is generated", "What it imitates"],
      [[XNAME[k], XDESC[k][0], XDESC[k][1]] for k in XDESC], cap="Activation families.", widths=[2.8, 7.6, 6.4], size=8.5)

# =========================================================================== 3. answer in brief
H("3. Answer in brief (more details below)", 1)
gap = T9["ocp"]
gap_v, gap_lo, gap_hi = (float(gap[k]) for k in ("gap winner vs runner-up (pp)", "95% CI low", "95% CI high"))
check(WIN == "MXFP6_E2M3", "E2M3 is the W+X held-out winner for C")
check(all(T9[p]["verdict"] == "MXFP6_E2M3" and float(T9[p]["95% CI low"]) > 0 for p in T9),
      "E2M3 wins W+X under every policy, interval above zero")
check(all(SC_RANK[p]["verdict"] == "E2M3" for p in SC_RANK), "E2M3 wins weights-only under every policy")
xs_c = XS["C"]
xs_lose = [f for f in xs_c if winner_of(xs_c[f]["winner"]) != "E2M3"]
check(all(num(xs_c[f]["margin pp"]) < 0.1 and winner_of(xs_c[f]["winner"]) in ("MXINT6", "BFP6") for f in xs_lose),
      "the families where E2M3 loses are near ties won by an integer format")
xs_min = min(num(xs_c[f]["margin pp"]) for f in xs_c)
out_share = OUT_T["400, 20σ, continuous"]["blocks with an outlier"]
seeds = lambda cell: int(re.search(r"\((\d+)/10\)", cell).group(1))                    # noqa: E731
margin = lambda cell: float(re.search(r"\+(\d+\.\d+)", cell).group(1))                 # noqa: E731
CMPC = "BFP6_g16 (cmp.)"                                   # the BFP 16|4|5 comparison column of the Q3 tables
check(all(winner_of(v["winner ocp (seeds), margin pp"]) == "E2M3" for v in OUT_T.values()),
      "E2M3 wins every outlier case")
hi = [v for k, v in OUT_T.items() if "continuous" in k and ", 100σ" in k]
check(all(num(v[CMPC]) < num(v["E2M3"]) for v in hi), "BFP 16|4|5 beats E2M3 at 100σ")
sp_win = [z for z in SPARSE if winner_of(SPARSE[z]["winner ocp (seeds), margin pp"]) == "E2M3"]
sp_lose = [z for z in SPARSE if z not in sp_win]
sp_all = max((z for z in sp_win if seeds(SPARSE[z]["winner ocp (seeds), margin pp"]) == 10), key=num)
check(sp_lose and min(num(z) for z in sp_lose) > max(num(z) for z in sp_win), "one crossover in the sparsity sweep")
sp_cross = min(sp_lose, key=num)
check(all(winner_of(SPARSE[z]["winner ocp (seeds), margin pp"]) == "MXINT6"
          and margin(SPARSE[z]["winner ocp (seeds), margin pp"]) < 0.05 for z in sp_lose),
      "past the crossover MXINT6 wins by a hair over BFP6")
row10k = [r for r in SCALEV if r["spread"].startswith("row") and "10^4" in r["spread"]][0]
col_rows = [r for r in SCALEV if r["spread"].startswith("column")]
CLAMP = "clamped blocks BFP6 / BFP6_g16"
clamp32 = lambda r: int(r[CLAMP].split("/")[0])            # noqa: E731
N_GROUPS = 100 * int(T2["BFP6"]["blocks/row (K=100)"])
check(clamp32(row10k) > 0 and num(row10k["BFP6"]) > 2 * num(row10k["E2M3"]), "BFP6 clamps at 10^4")
check(all(r["weak-25% winner"] in ("E3M2", "E4M1") for r in col_rows[1:]), "wide-range formats keep the weak channels")
c_col = [h for h in SHIFT_H if h.startswith("C ")][0]
shift_c = {r["X_test family"]: num(r[c_col].split("/")[1]) for r in SHIFT}
worst_fam = max(shift_c, key=shift_c.get)
check(DS["C"]["pick"] == WIN, "the chosen method picks the winner for C")
box("Q1. Which format is most suitable for W_[C] at 6 bits?", [
    f"**{LONG[WIN]}.** With W and X both quantized, its held-out output error is **{f3(E[WIN])}%** under the class "
    "scale rule, against " + ", ".join(f"{LONG[n]} {f3(E[n])}%" for n in RANK[1:]) + f". The gap to {LONG[RANK[1]]} "
    f"is {f3(gap_v)} pp, with 95% interval [{f3(gap_lo)}, {f3(gap_hi)}]: not a tie. Storage: "
    f"{float(T2[WIN]['bits/elem']):.2f} bits per element.",
    f"It is first under every scale policy, and also with weights only ({f3(EW[WIN])}% against "
    f"{f3(EW[RANK_W[1]])}% for " + " and ".join(LONG[n] for n in RANK_W[1:] if f3(EW[n]) == f3(EW[RANK_W[1]])) + ").",
    (f"Under {len(xs_c)} different activation families it stays first in {len(xs_c) - len(xs_lose)}; an integer "
     "format wins narrowly for " + " and ".join(f"{XNAME[f]} ({f2(num(xs_c[f]['margin pp']))} pp)" for f in xs_lose)
     + ". " if xs_lose else
     f"It is also first under all {len(xs_c)} activation families we tried, by at least {f2(xs_min)} pp. ")
    + (f"Choosing with {XC} costs nothing: whichever of these families {XT} comes from, E2M3 is still the best format for it."
       if shift_c[worst_fam] == 0 else
       f"Choosing with {XC} costs at most {f3(shift_c[worst_fam])} dB if {XT} comes from any of these families "
       f"(worst case: {XNAME[worst_fam]})."),
])
box("Q2. Can an efficient method match exhaustive evaluation without quantizing every candidate?", [
    "**Yes. The chosen method is the exact activation-weighted sketch**: `best, scores = selector.select_format(W, X_cal)` "
    "in `selector.py` (Section 6).",
    "The output error splits into a weight term (each weight's rounding error, weighted by the energy of its input "
    "channel) and an activation term (each activation's rounding error, weighted by the energy of its column of W). "
    "The method sorts the weights, and then the activations, into a fixed set of bins by their position relative "
    "to their block maximum, and adds up three sums per bin. A small table per format turns these sums into that "
    "format's error. Nothing is quantized.",
    "The result is **exact**: the scores equal the errors that quantizing with all five formats gives for the same "
    f"objective, and the method makes the same choice as that brute force on all {N_RANDOM} random test matrices. "
    f"It picks {SHORT[WIN]} for C. On the {N_RANDOM} matrices with {len(XKINDS)} activation families it picks the "
    f"held-out best format in **{pct(BM['top1'])}** of cases (mean regret {f3(BM['regret'])} dB), against "
    f"{pct(bench('random', 'exhaustive_y')['top1'])} for brute force on the full output. It takes {BM['ms']:.2f} ms "
    f"per matrix, against {bench('random', 'exhaustive_y')['ms']:.2f} ms for that brute force.",
    f"We compared it with {len(S.ROLE) - 1} other methods: brute-force references, variants of the method, "
    "extensions and alternative methods (Section 6.5).",
], fill=METHOD_FILL)
b632 = max((b for b in BUDGET if BUDGET[b]["best"] == "E2M3"), key=float)
check(all(BUDGET[b]["best"] == "E2M3" for b in BUDGET if 6.08 <= float(b) <= float(b632)), "E2M3 best up to b632")
box("Q3. How do outliers, sparsity and variation in scale affect the selection?", [
    f"**Outliers:** E2M3 wins in every case we tried, up to outliers of 100σ and an outlier in {out_share} of the "
    "blocks. Its lead shrinks at 100σ, where every format flushes the outlier's neighbours; BFP with groups of 16 "
    "(a comparison, not a candidate) would beat it there, because each outlier then coarsens fewer neighbours.",
    "**Sparsity:** zeros help the integer grids, whose precision is set by the block maximum; the FP error hardly "
    f"changes. E2M3 still wins up to {max(sp_win, key=num)} zeros; from {sp_cross} MXINT6 and BFP6 overtake it.",
    "**Variation in scale:** a different gain per output row is absorbed by the 8-bit MX scale, but BFP6's 4-bit "
    "exponent runs out of range at a spread of 10^[4]. Gains per input channel leave E2M3 best overall, but only "
    "the wide-range formats (E3M2, E4M1) keep the weakest channels accurate.",
    f"**Block organisation:** at equal storage, E2M3 with MX blocks stays best up to {b632} bits per element; with "
    "more storage, BFP with short groups wins.",
])

# =========================================================================== 4. Q1
H("4. Q1: which format is most suitable for W_[C]?", 1)
H("4.1 Result", 2)
rows = []
for n in RANK:
    r = T1[n]
    rows.append([LONG[n]] + [f3(float(r[WX.format(p)])) for p in ("ocp", "ceil", "opt", "opt_out")]
                + [f3(float(r["W only held-out Y err % (ocp)"])), f3(float(r["W only held-out Y err % (opt)"])),
                   f"{float(r['bits/elem']):.2f}"])
G16 = T1["BFP6_g16 (comparison)"]
rows.append(["BFP 16|4|5 (comparison, not a candidate)"] + [f3(float(G16[WX.format(p)])) for p in ("ocp", "ceil", "opt", "opt_out")]
            + [f3(float(G16["W only held-out Y err % (ocp)"])), f3(float(G16["W only held-out Y err % (opt)"])),
               f"{float(G16['bits/elem']):.2f}"])
t_res = table(["Format", "W+X `ocp` (class rule)", "W+X `ceil`", "W+X `opt`", "W+X `opt_out`", "Weights only `ocp`",
               "Weights only `opt`", "Bits/elem"], rows,
      cap="Dataset C: held-out output error e_[Y] (%) under each scale policy, sorted by the class rule. "
          "Primary columns: W and X quantized (results/tables/T1_main_C.csv). Last row: BFP with groups of 16, "
          "for comparison only.", highlight=(0,), size=8.5)
check(all(min(NAMES, key=lambda n: float(T1[n][WX.format(p)])) == WIN for p in ("ocp", "ceil", "opt", "opt_out"))
      and all(min(NAMES, key=lambda n: float(T1[n][f"W only held-out Y err % ({p})"])) == WIN for p in ("ocp", "opt")),
      "E2M3 is lowest in every column of the result table")
P(f"**{LONG[WIN]} is the most suitable format.** It has the lowest held-out error under every scale policy, with W "
  f"and X quantized and with weights only. The gap to {LONG[RANK[1]]} is clear of sampling noise under each policy "
  f"of Table {TABNO[0] + 1}.")
table(["Scale policy", "Ranking (best first)", "Gap winner vs runner-up (pp)", "95% interval"],
      [[f"`{p}`", " < ".join(SHORT[n] for n in T9[p]["ranking (best first)"].split(" < ")),
        f3(float(T9[p]["gap winner vs runner-up (pp)"])),
        f"[{f3(float(T9[p]['95% CI low']))}, {f3(float(T9[p]['95% CI high']))}]"] for p in T9],
      cap="W+X ranking and the gap between the top two, paired bootstrap over the held-out columns "
          "(results/tables/T9_wx_ranking.csv).")
H("4.2 Why E2M3 wins", 2)
check(int(T6["MXFP6_E2M3"]["zeroed"]) < min(int(T6["BFP6"]["zeroed"]), int(T6["MXINT6"]["zeroed"])), "zeroed counts order")
z_int = (f"{T6['MXINT6']['zeroed']} (MXINT6 and BFP6 alike)" if T6["MXINT6"]["zeroed"] == T6["BFP6"]["zeroed"]
         else f"{T6['BFP6']['zeroed']} (BFP6) and {T6['MXINT6']['zeroed']} (MXINT6)")
werr = {n: float(T1[n]["W err % (ocp)"]) for n in NAMES}
xerr = {n: float(T1[n]["X err % (ocp)"]) for n in NAMES}
check(min(werr, key=werr.get) == WIN, "E2M3 has the smallest weight error")
x_best = min(xerr, key=xerr.get)
fit_cols = X_CAL[:, ::2]
w_term = S.exhaustive(W_C, (fit_cols ** 2).sum(1))
x_term = S.exhaustive(fit_cols.T, (W_C ** 2).sum(0))
w_share = {n: w_term[n] / (w_term[n] + x_term[n]) for n in NAMES}
bullets([
    f"**It keeps the small weights.** E2M3 has 3 mantissa bits and subnormals, so small weights stay on a fine grid: "
    f"it flushes only {T6['MXFP6_E2M3']['zeroed']} of the 10,000 weights to zero. The integer grids flush "
    f"{z_int}, because their step is fixed by the block maximum. With weights only, flushed weights carry {SRC[('MXINT6', 'ocp')]['zero']} of MXINT6's squared error and "
    f"{SRC[('BFP6', 'ocp')]['zero']} of BFP6's, against {SRC[('E2M3', 'ocp')]['zero']} for E2M3. This is the known "
    "trade-off between floating-point and integer grids, studied for 8-bit formats in [10, 11].",
    f"**It does not pay for range that W_[C] does not need.** The largest weight is only {MAXS:.1f}σ, so E3M2 and E4M1 "
    "spend exponent bits on range and lose mantissa bits where the weights are. Under the class rule, clipping is "
    f"{SRC[('E3M2', 'ocp')]['clip']} of E3M2's weight error and {SRC[('E4M1', 'ocp')]['clip']} of E4M1's, because "
    "their grids end well below the top of their binade (Figure 2).",
    (f"**Both error terms favour it, and the weight term decides.** Quantized on its own, {XC} also has its smallest "
     f"error with E2M3 ({f3(xerr[WIN])}%, against {f3(xerr['MXINT6'])}% for MXINT6 and {f3(xerr['BFP6'])}% for BFP6), "
     f"but the gap is small: {XC} is Gaussian with a small range, which suits the integer grids too. The weight "
     f"errors differ much more ({f3(werr[WIN])}% against {f3(werr['BFP6'])}% for BFP6)"
     if x_best == WIN else
     f"**The weight error decides.** Quantized on its own, {XC} is handled slightly better by "
     f"{LONG[x_best]} ({f3(xerr[x_best])}% against {f3(xerr[WIN])}% for E2M3). But E2M3's weight error is much "
     f"smaller ({f3(werr[WIN])}% against {f3(werr['BFP6'])}% for BFP6)")
    + f", and the weight term is the larger part of the output error ({pct(w_share[WIN], 0)} of it for E2M3, "
    f"{pct(w_share['BFP6'], 0)} for BFP6; Section 6.1 defines the two terms).",
])
figure("fig2_error_model.png", "RMS relative rounding error of each element grid against a value's magnitude "
       "relative to its block maximum. E2M3 is finest over the top octaves, where most of W_[C]'s energy lies; the "
       "integer grids lose precision steadily below the block maximum.", width=14)
H("4.3 Does the answer depend on the activations?", 2)
P(f"We repeated the comparison under the {len(xs_c)} activation families of Section 2.6 (`x_stress.py`, "
  f"results/x_stress_wx.md). Table {TABNO[0] + 1} shows dataset C.")
table(["Activation family", "Winner (seeds)", "Margin pp", "E3M2", "E2M3", "E4M1", "MXINT6", "BFP6"],
      [[XNAME[f], xs_c[f]["winner"], xs_c[f]["margin pp"]] + [v.strip() for v in
       xs_c[f]["errors E3M2 / E2M3 / E4M1 / MXINT6 / BFP6"].split("/")] for f in xs_c],
      cap="Dataset C, W and X quantized, class rule: held-out e_[Y] (%) under each activation family "
          "(mean over 5 seeds for the synthetic families).", size=8.5)
xs_cw = XS_W["C"]
check(all(winner_of(xs_cw[f]["winner"]) == "E2M3" for f in xs_cw), "weights only: E2M3 wins C under every family")
b_col = [h for h in SHIFT_H if h.startswith("B ")][0]
shift_b = {r["X_test family"]: num(r[b_col].split("/")[1]) for r in SHIFT}
worst_b = max(shift_b, key=shift_b.get)
check(shift_b[worst_b] > 1, "dataset B's choice is fragile under other activations")
shift_cw = max(num(r[[h for h in SHIFT_H_W if h.startswith("C ")][0]].split("/")[1]) for r in SHIFT_W)
xs_tight = min(xs_c, key=lambda f: num(xs_c[f]["margin pp"]))
P((f"E2M3 wins under {len(xs_c) - len(xs_lose)} of the {len(xs_c)} families. The exceptions, "
   + " and ".join(XNAME[f] for f in xs_lose) + ", are near ties won by an integer format by less than 0.1 pp. "
   if xs_lose else
   f"E2M3 wins under all {len(xs_c)} families, in every seed. The closest family is {XNAME[xs_tight]} "
   f"({xs_c[xs_tight]['margin pp']} pp ahead of the runner-up); with massive activations BFP6 falls far behind "
   f"MXINT6 ({xs_c['massive']['errors E3M2 / E2M3 / E4M1 / MXINT6 / BFP6'].split('/')[4].strip()}% against "
   f"{xs_c['massive']['errors E3M2 / E2M3 / E4M1 / MXINT6 / BFP6'].split('/')[3].strip()}%): the largest values "
   "of a channel multiplied by 100 lie above 248, the largest value BFP6 can represent at the top of its 4-bit "
   "exponent range, so they saturate. ")
  + (f"What matters in practice is the cost of committing to E2M3 with {XC}. It is zero: for {XT} from any of the "
     "families, E2M3 remains the best format (results/x_stress_wx.md, last table). This is not automatic: for dataset "
     f"B, whose {XC} winner {SHORT[T7[('B', 'wx', 'ocp')]['winner']]} is a near tie, the same test costs up to "
     f"{f3(shift_b[worst_b])} dB ({XNAME[worst_b]})."
     if shift_c[worst_fam] == 0 else
     f"What matters in practice is the cost of committing to E2M3 with {XC}: on {XT} from any family it is at most "
     f"{f3(shift_c[worst_fam])} dB (worst case {XNAME[worst_fam]}; results/x_stress_wx.md, last table).")
  + (" With weights only, E2M3 also wins under every family, and committing to it costs nothing either."
     if shift_cw == 0 else
     f" With weights only, E2M3 wins under every family and the cost is at most {f3(shift_cw)} dB."))
H("4.4 Do our assumptions change the answer?", 2)
e4 = {k: f3(float(v["held-out Y err % (ocp)"])) for k, v in T3.items()}
check(all(float(v) > max(EW[n] for n in NAMES if n != "MXFP6_E4M1") for v in e4.values()), "E4M1 is last in every case")
bullets([
    f"**E4M1 without reserved codes (A4).** Our choice (v_[max] 384) gives {e4['MXFP6_E4M1']}% (weights only, class "
    "rule). An IEEE-like reserved top exponent (v_[max] 192) gives "
    + ("the same value, because under the class rule no block uses the top binade" if e4["MXFP6_E4M1_v192"] == e4["MXFP6_E4M1"]
       else f"{e4['MXFP6_E4M1_v192']}%")
    + f"; one NaN code (v_[max] 256) gives {e4['MXFP6_E4M1_nan1']}%. E4M1 is last in every case.",
    f"**MXINT6 two's complement against symmetric [−31, 31] (A5).** {f3(float(T8[('C', 'ocp')]['two' + chr(39) + 's compl. [-32,31]']))}% "
    f"for both under the class rule (weights only). Under `opt` two's complement is slightly better "
    f"({f3(float(T8[('C', 'opt')]['two' + chr(39) + 's compl. [-32,31]']))}% against "
    f"{f3(float(T8[('C', 'opt')]['symmetric [-31,31]']))}%), because a negative block maximum may use the code −32.",
    f"**BFP organisation (A6).** BFP with 32 values, an 8-bit exponent, the floor rule and round-half-to-even is "
    f"numerically identical to symmetric MXINT6 ({f3(float(bfp32['held-out Y err % (ocp)']))}%, a unit test checks "
    "this). Our BFP6 has the same grid and groups, so it differs from symmetric MXINT6 only where a group leaves the "
    + ("range of its 4-bit exponent; no group of W_[C] does, so on C the two give exactly the same error. "
       if bfp_eq_sym else "range of its 4-bit exponent. ")
    + f"Groups of 16 (16|4|5) lower BFP's error to {f3(float(G16[WX.format('ocp')]))}% with W and X quantized "
    f"({f3(float(G16['W only held-out Y err % (ocp)']))}% weights only), for "
    f"{float(G16['bits/elem']) - float(T2['BFP6']['bits/elem']):.2f} more bits per element"
    + (f"; E2M3 stays better ({f3(E[WIN])}%)." if float(G16[WX.format('ocp')]) > E[WIN] else "."),
    "**Range limits.** No BFP6 group of W_[C] reaches the limits of the 4-bit exponent (0 clamped groups, "
    "results/tables/T6_diagnostics_ocp.csv), so every candidate can represent W_[C].",
])
check(all(int(T6[n]["clamped blocks"]) == 0 for n in NAMES), "no block of C is clamped")
H("4.5 The other datasets", 2)
t7rows = []
for k in "ABCDE":
    a, b = T7[(k, "wx", "ocp")], T7[(k, "w", "ocp")]
    t7rows.append([k] + [f3(float(a[n])) for n in NAMES] + [SHORT[a["winner"]], f3(float(a["margin to 2nd (pp)"])),
                                                            SHORT[b["winner"]]])
table(["Dataset", "E3M2", "E2M3", "E4M1", "MXINT6", "BFP6", "Winner W+X", "Margin pp", "Winner weights only"], t7rows,
      cap="Context only: the five course datasets with the same X_[cal], W+X held-out e_[Y] (%), class rule "
          "(results/tables/T7_datasets_A_to_E.csv). Our dataset is C.", highlight=(2,))


def by_winner(mode):
    g = {}
    for k in "ABCDE":
        g.setdefault(SHORT[T7[(k, mode, "ocp")]["winner"]], []).append(k)
    return "; ".join(", ".join(v) + (" choose " if len(v) > 1 else " chooses ") + w
                     for w, v in sorted(g.items(), key=lambda kv: -len(kv[1])))


int_wins = [(k, m) for k in "ABCDE" for m in ("wx", "w") if T7[(k, m, "ocp")]["winner"] in ("MXINT6", "BFP6")]
check(int_wins and all(abs(float(T7[km + ("ocp",)]["MXINT6"]) - float(T7[km + ("ocp",)]["BFP6"])) < 0.01 for km in int_wins),
      "where an integer format wins, MXINT6 and BFP6 are within 0.01 pp")
check(all(T7[(k, "wx", "ocp")]["winner"] == "MXFP6_E2M3" for k in "CD"), "C and D choose E2M3 in W+X")
P(f"With W and X quantized, {by_winner('wx')}; with weights only, {by_winner('w')}. Wherever an integer format "
  "wins, MXINT6 and BFP6 are within 0.01 pp of each other, because inside its exponent range BFP6 has the same grid "
  "as symmetric MXINT6. So the best format does depend on the data, and C and D are the clearest cases for E2M3 "
  f"(margins {f3(float(T7[('C', 'wx', 'ocp')]['margin to 2nd (pp)']))} and "
  f"{f3(float(T7[('D', 'wx', 'ocp')]['margin to 2nd (pp)']))} pp).")

# =========================================================================== 5. scale optimization
H("5. Scale optimization", 1)
P("The shared exponent is the one free parameter of a block, so we studied how it should be chosen (Section 2.3 "
  "defines the policies). The study uses `scale_opt.py`, an implementation of the policies written independently "
  "of `formats.py`; the two agree bit for bit (results/scale_study.md §1).")
H("5.1 With W and X quantized", 2)
gain = {n: E[n] - float(T1[n][WX.format("opt")]) for n in NAMES}
ceil_better = [n for n in NAMES if float(T1[n][WX.format("ceil")]) < E[n]]
check(all(v > 0 for v in gain.values()), "opt helps every format in W+X")
check(set(ceil_better) == {"MXFP6_E3M2", "MXFP6_E4M1"}, "ceil helps only E3M2 and E4M1 in W+X")
check(all(T9[p]["verdict"] == WIN for p in T9), "the W+X ranking does not depend on the policy")
bullets([
    "**The per-block search (`opt`) helps every format, but never changes the ranking.** Gain over the class rule: "
    + "; ".join(f"{SHORT[n]} {f3(gain[n])} pp" for n in RANK) + f" (Table {t_res}).",
    "**For E3M2 and E4M1 the class rule clips too often.** For them `ceil`, the smallest scale that clips nothing, "
    "is better than the class rule. For E2M3, MXINT6 and BFP6 it is worse: raising the exponent by one doubles the "
    "step for every other value in the block.",
    f"**Weighting the search by the activations (`opt_out`)** gains little over `opt` (E2M3 "
    f"{f3(float(T1[WIN][WX.format('opt_out')]))}% against {f3(float(T1[WIN][WX.format('opt')]))}%).",
])
H("5.2 The full study, weights only", 2)
pol = ["ocp", "ceil", "opt", "model", "opt_out", "oracle_y", "bound_y"]
table(["Format"] + [f"`{p}`" for p in pol],
      [[SC2A[k]["format"]] + [SC2A[k][p] for p in pol] for k in ("E2M3", "BFP6", "MXINT6", "E3M2", "E4M1")],
      cap="Weights only: held-out e_[Y] (%) of C under every exponent policy (results/scale_study.md §2a).",
      highlight=(0,))
g = {k: ci(SC_GAIN[k]["opt"]) for k in SC_GAIN}
check(all(v[1] > 0 for v in g.values()), "opt helps every format (interval above 0)")
check(num(SC2A["E2M3"]["oracle_y"]) > num(SC2A["E2M3"]["opt"]), "oracle_y overfits on E2M3")
check(all(abs(num(SC2A[k]["model"]) - num(SC2A[k]["opt"])) < 0.02 for k in SC2A), "model is close to opt")
fit = SC_FIT["E2M3"]
bullets([
    "**The search gain is real.** Gain of `opt` over the class rule, pp [95% interval]: "
    + "; ".join(f"{k} {v[0]:+.3f} [{v[1]:+.3f}, {v[2]:+.3f}]" for k, v in g.items()) + ".",
    "**Predicting the error instead of rounding (`model`)** reaches the same result within 0.02 pp.",
    f"**Fitting the exponents to the activations overfits.** `oracle_y` is best on the columns it was fitted on (E2M3 "
    f"{fit['oracle_y (even / odd)'].split('/')[0].strip()}%) but worse on the held-out ones "
    f"({fit['oracle_y (even / odd)'].split('/')[1].strip()}% against {fit['opt (even / odd)'].split('/')[1].strip()}% "
    f"for `opt`). The lower bound for any exponent choice on the held-out columns is {SC2A['E2M3']['bound_y']}%, so "
    "`opt` is already close to the best possible.",
    "**Conclusion.** We report the class rule in every headline table, with `opt` alongside. The choice of format "
    "is the same under every policy.",
])
figure("fig5_scale_tradeoff.png", "Weights only: weight error of C as every block's exponent is moved by d from "
       "the class rule, split into clipped, rounded and flushed parts. The pink line is the per-block optimum.",
       width=17)

# =========================================================================== 6. Q2
H("6. Q2: the chosen selection method", 1)
box("The chosen method: the exact activation-weighted sketch", [
    "`best, scores = selector.select_format(W, X_cal)` in `selector.py`. It reads W and X once, quantizes nothing, "
    "and returns the exact activation-weighted error of every candidate and the format with the smallest one.",
    "Sections 6.1–6.4 explain and evaluate it. Section 6.5 compares it with the other methods we built.",
], fill=METHOD_FILL)
H("6.1 What error the method minimises", 2)
P("Write the quantized operands as W_[q] = W − ΔW and X_[q] = X − ΔX, where ΔW and ΔX are the rounding errors. "
  "The output error is then")
eq("W·X − W_[q]·X_[q] = ΔW·X + W·ΔX − ΔW·ΔX")
P("The last term is a product of two small errors (a few percent each), so we drop it. The rounding errors of W and "
  "of X are independent of each other, so on average the two remaining terms are uncorrelated and their squared "
  "sizes add:")
eq("‖W·X − W_[q]·X_[q]‖^[2] ≈ ‖ΔW·X‖^[2] + ‖W·ΔX‖^[2]")
P("**The weight term.** Row i of ΔW·X is Δ_[i]·X, where Δ_[i] is row i of ΔW. Its squared size is "
  "Δ_[i]·H·Δ_[i]^[T] with H = X·X^[T]. Rounding errors of different weights are close to independent with mean zero, "
  "so in expectation only the diagonal of H survives:")
eq("‖ΔW·X‖^[2] ≈ Σ_[i,j] h_[j]·ΔW_[ij]^[2],     h_[j] = Σ_[k] X_[jk]^[2]")
P("h_[j] is the **energy of input channel j**. So the weight term is a weighted sum of squared weight errors, in "
  "which every weight counts in proportion to the energy of the input channel it multiplies.")
P("**The activation term** has exactly the same form with the roles swapped: X^[T] takes the place of W, and the "
  "weight of input channel j is g_[j] = Σ_[i] W_[ij]^[2], the energy of column j of W:")
eq("‖W·ΔX‖^[2] ≈ Σ_[j,k] g_[j]·ΔX_[jk]^[2]")
P("With weights only, the activation term is absent. The idea that a weight's importance comes from its "
  "activations is the same one AWQ [5] uses; that the output error is close to such a weighted sum is also the "
  "basis of HIGGS [9].")
_, sc_c = S.select_format(W_C, fit_cols)
full_c = S.exhaustive_full_wx(W_C, fit_cols)
dev = max(abs(sc_c[n] / full_c[n] - 1) for n in NAMES)
check(dev < 0.02, "the two-term approximation is within 2% on C")
we = bench("random", "exhaustive_w")
wm = bench("random", "exhaustive_w", "massive")
P(f"**How good is the approximation?** On {WC} with the fit columns of {XC}, the sum of the two terms is within "
  f"{100 * dev:.1f}% of the true squared output error for every candidate. Leaving out the weighting is much worse: "
  f"choosing by the plain squared errors of W and X (brute force, unweighted) picks the held-out best format in "
  f"only {pct(we['top1'])} of our random test matrices, and in {pct(wm['top1'])} when a few activation channels "
  "are very large.")
H("6.2 How the method computes the error without quantizing", 2)
P("**The key observation.** Under the class rule, a block's scale is set by its largest value: it is 2^[s] with "
  "s = R − e_[max] and R = floor(log_[2] max|w|) (for BFP6, R is clamped to its 4-bit range). Write any non-zero "
  "value of the block as |w| = μ·2^[t] with 1 ≤ μ < 2, and let **d = R − t** be the number of octaves it sits below "
  "the block maximum. Divided by the scale, the value becomes μ·2^[t − s], and t − s = e_[max] − d depends only on "
  "d, not on R. So the rounded value is c_[f]·2^[t], where the factor c_[f](d, μ, sign) depends only on the format "
  "f and the value's position relative to its block maximum. Two values at the same relative position are rounded "
  "in exactly the same relative way in every block.")
P("The squared error of one value then splits into three sums that do not depend on the format:")
eq("(|w| − c_[f]·2^[t])^[2] = m_[2] − 2c_[f]·m_[1] + c_[f]^[2]·m_[0],     m_[0] = 4^[t],  m_[1] = 4^[t]·μ,  m_[2] = w^[2]")
P("**The algorithm**, for W (and then for X^[T] in the W+X setting):")
bullets([
    f"**Activation weights.** Compute h_[j] = Σ_[k] X_[jk]^[2] for every input channel from the calibration "
    "columns. (For the activation term: g_[j] = Σ_[i] W_[ij]^[2] for every column of W.)",
    "**Sketch.** Make one pass over the matrix for each block organisation (MX blocks with an 8-bit scale, BFP6 "
    "groups with a clamped 4-bit exponent; the two differ only where the clamp acts). For each "
    "non-zero value, find its bin: its sign, its octave offset d, and which of 64 equal slices of [1, 2) its "
    "mantissa μ falls in. Add h_[j]·m_[2], h_[j]·m_[1] and h_[j]·m_[0] to that bin. The sketch is these "
    f"2 × {S.DMAX - S.DMIN + 1} × {S.NB} × 3 sums per layout, the same for every format. (Offsets d run from "
    f"{S.DMIN} to {S.DMAX}: negative offsets occur only when a BFP6 group's maximum is above its exponent range, and "
    f"a value more than {S.DMAX} octaves below its block maximum is flushed to zero by every format, so such values "
    "share the last offset.)",
    "**Kernel.** For each format, a table gives c_[f] for every bin. It is computed once from the format's "
    "definition alone (its grid, v_[max], subnormals, sign handling, the −32 code of MXINT6, BFP6's exponent "
    "clamp); no data is needed.",
    "**Score.** Each format's error is the sum over bins of S_[2] − 2c_[f]·S_[1] + c_[f]^[2]·S_[0], where "
    "S_[0], S_[1], S_[2] are the bin's three sums. In the W+X setting, the format's score is its weight term plus "
    "its activation term.",
    "**Choose** the format with the smallest score.",
], style="List Number")
P("**Why it is exact, not an approximation.** For 6-bit grids, every rounding boundary and every saturation point, "
  "written as μ·2^[t], has μ at a multiple of 1/32. With 64 equal mantissa slices of [1, 2), no boundary falls "
  "inside a slice, so c_[f] is constant on each bin and the three sums give the error exactly. `test_selector.py` checks this to "
  "a relative 10^[−9] for all five candidates, on matrices built to hit every corner case: values exactly on grid "
  "points and on rounding midpoints, MXINT6's −32 code, both ends of BFP6's exponent range, all-zero blocks, the "
  "tail block of 4, 16 octaves of dynamic range and float32 data. It also checks that in the W+X setting the score "
  "equals the weight term plus the activation term computed by quantizing.")
P("**Properties that follow.**")
bullets([
    "**Cost.** One pass over the data, shared by all candidates. Scoring a format costs a fixed number of bins, "
    "whatever the size of the matrix; adding a format means adding its kernel, not reading the data again.",
    "**Additive.** Sketches of parts of a matrix (rows, layers, devices) add up, so a sketch can be built in a "
    "stream or in parallel.",
    "**Explains the choice.** Each format's error splits by octave offset and mantissa slice into flushed, clipped "
    "and rounded parts; Q3 uses this to explain why the winner changes.",
])
H("6.3 How we tested it", 2)
P(f"Besides the five course datasets, we built a permanent, seeded test suite (`selector_suite.py`):")
bullets([
    f"**{N_RANDOM} random 128 × 128 matrices.** Each starts from a Gaussian, Laplace, Student-t3 or log-uniform "
    "distribution; then a random share of weights (0–90%) is set to zero; outliers 3× to 1000× larger are added at "
    "random positions, in whole input channels, or one per block; and random gains are applied per row and, for "
    f"some matrices, per column. Each matrix gets one of the {len(XKINDS)} synthetic activation families of "
    "Section 2.6 (256 columns: even columns to fit, odd columns held out).",
    f"**{N_ADV} adversarial families**, each built to break one assumption: heavy tails, 50–95% zeros, outliers of "
    "20σ and 50σ, one outlier per block, 4 to 12 octaves of range, Pareto and log-normal magnitudes, few-level "
    "alphabets (ternary, binary, 6 levels), already-quantized data, column gains of 10^[4], half the rows zero, and "
    "matrices at both ends of BFP6's exponent range.",
    "**The truth** for every matrix is the held-out output error after quantizing with all five formats.",
])
H("6.4 Results", 2)
ds_rows = [[k, SHORT[DS[k]["pick"]], SHORT[DS[k]["true_y"]], f3(float(DS[k]["regret_y_db"]))] for k in "ABCDE"]
misses = [k for k in "ABCDE" if DS[k]["pick"] != DS[k]["true_y"]]
check(all(DS_Y[k]["pick"] == DS[k]["pick"] for k in misses), "brute force on the fit columns makes the same misses")
check(all(float(T7[(k, "wx", "ocp")]["margin to 2nd (pp)"]) < 0.05 for k in misses), "the misses are near ties")
table(["Dataset", "Chosen method picks", "Held-out winner", "Regret (dB)"], ds_rows,
      cap="The chosen method on the five course datasets, W+X (activation weights from the fit columns; truth on "
          "the held-out columns).", highlight=(2,))
P(f"It finds the held-out winner on {5 - len(misses)} of 5 datasets. "
  + (" ".join(f"On {k} it picks {SHORT[DS[k]['pick']]} where {SHORT[DS[k]['true_y']]} is better by "
              f"{f3(float(T7[(k, 'wx', 'ocp')]['margin to 2nd (pp)']))} pp; brute force on the full output of the fit "
              "columns makes the same pick, so the miss comes from the difference between the fit and held-out "
              "columns, not from the method." for k in misses) if misses else ""))
refs = [("exhaustive_h", "Brute force, same objective"), ("exhaustive_y", "Brute force, full output on fit columns"),
        ("exhaustive_w", "Brute force, unweighted")]
rows = [["Chosen method: exact sketch, activation-weighted", pct(BM["top1"]), f3(BM["regret"]), f2(BM["p95"]),
         f2(BM["max"]), f2(BM["ms"])]]
for a, lab in refs:
    b = bench("random", a)
    rows.append([lab, pct(b["top1"]), f3(b["regret"]), f2(b["p95"]), f2(b["max"]), f2(b["ms"])])
same = all(r1["pick"] == r2["pick"] for r1, r2 in zip(BM["rows"], bench("random", "exhaustive_h")["rows"]))
check(same, "the chosen method makes exactly the same picks as brute force with the same objective")
ey = bench("random", "exhaustive_y")
check(abs(ey["top1"] - BM["top1"]) < 0.01, "the method is within 1 point of full brute force")
table(["Method", "Top-1 agreement", "Mean regret dB", "95th pct. regret dB", "Max regret dB", "Median ms"], rows,
      cap=f"Random suite, {N_RANDOM} matrices, W+X (results/selector_bench_wx.md).", highlight=(0,))
P(f"**The chosen method makes exactly the same pick as brute force with the same objective on every one of the "
  f"{N_RANDOM} matrices**, without quantizing anything.")
miss_rg = [float(r["regret_y_db"]) for r in BM["rows"] if r["hit"] == "0"]
check(miss_rg and np.median(miss_rg) < 0.25, "the misses are mostly near ties")
P(f"**Why {pct(BM['top1'])} and not 100%: the misses come from the held-out columns, not from the method.** Every "
  "method, including brute force, chooses with the fit (even) columns of X, while the truth is measured on the "
  "held-out (odd) columns, which stand in for the unseen X_[test]. When two formats are close, the better one on "
  "the fit columns is sometimes not the better one on the held-out columns. Brute force with the same objective, "
  f"which quantizes every candidate, misses exactly the same {len(miss_rg)} matrices, because its picks are "
  "identical. Brute force on the full output error of the fit columns, which also captures correlations between "
  f"channels, is right in {pct(ey['top1'])} of cases: no better than the chosen method within the noise of the "
  "test. The misses are also mostly near ties: their median regret is "
  f"{np.median(miss_rg):.2f} dB, about {100 * (10 ** (np.median(miss_rg) / 10) - 1):.0f}% more error power. "
  "No method that sees only the fit columns can remove these misses; a method scored on the held-out columns "
  "themselves would reach 100% by definition.")
by = [[XNAME.get(k, k)] + [f"{pct(bench('random', a, k)['top1'], 0)} / {f3(bench('random', a, k)['regret'])}"
                           for a in (M, "exhaustive_y", "exhaustive_w")] for k in XKINDS]
table(["Activation family", "Chosen method", "Brute force, full output", "Brute force, unweighted"], by,
      cap="Random suite by activation family: top-1 agreement / mean regret (dB), W+X.")
fam_d = {k: bench("random", "exhaustive_y", k)["top1"] - bench("random", M, k)["top1"] for k in XKINDS}
ahead = sorted((k for k in XKINDS if fam_d[k] > 0.03), key=lambda k: -fam_d[k])
behind = sorted((k for k in XKINDS if fam_d[k] < -0.02), key=lambda k: fam_d[k])
check(ahead and set(ahead) <= {"relu", "lowrank"}, "full brute force is clearly ahead only where X has dominant directions")
P("Per family the two are within a few points of each other, in both directions. Full brute force is clearly ahead "
  "for " + " and ".join(f"{XNAME[k]} activations ({pct(bench('random', M, k)['top1'], 0)} against "
                        f"{pct(bench('random', 'exhaustive_y', k)['top1'], 0)})" for k in ahead)
  + ": these have a few dominant directions (for ReLU, the non-zero mean), so the error depends on how ΔW lines up "
  "with them, which a per-weight weighting cannot see. Section 6.5 lists an extension that handles it. "
  + ("The chosen method is ahead for " + " and ".join(f"{XNAME[k]} ({pct(bench('random', M, k)['top1'], 0)} against "
                                                       f"{pct(bench('random', 'exhaustive_y', k)['top1'], 0)})"
                                                       for k in behind)
     + ", where the off-diagonal part of X·X^[T] on 128 fit columns is mostly sampling noise that does not carry "
     "over to the held-out columns." if behind else ""))
figure("fig6_selector_accuracy_vs_cost.png", "Accuracy against cost of the chosen method (star) and of the methods "
       f"we compared it with, random suite, W+X. Colour: what each method scores (no activation weighting, "
       "activation-weighted, or the full output on the fit columns).", width=17)
big = SCALING[-1]
check(big["sketch agrees"] == "yes" and all(r["sketch agrees"] == "yes" for r in SCALING), "the sketch agrees at every size")
P(f"**Cost at scale** (Figure {FIGNO[0] + 1}; results/selector_scaling_wx.md). On matrices shaped like LLM layers up to "
  f"{big['size']} ({big['weights']} weights) the chosen method always agrees with brute force. In NumPy it takes "
  f"{big['exact sketch (s)']} s against {big['brute force (s)']} s: both are limited by memory traffic, so the gain "
  "is the single shared pass, not raw speed. The certified extension (Section 6.5) gives the same pick in "
  f"{big['certified IS (s)']} s, rounding only {big['blocks rounded by certified']} blocks.")
figure("fig7_selector_scaling.png", "Selection time against matrix size, W+X: brute force, the chosen method, and "
       "the certified (importance-sampled) extension. Labels: share of blocks the certified selector rounded.", width=12)
P(f"**Weights only.** In the secondary setting the chosen method picks the held-out best format in "
  f"{pct(BMW['top1'])} of the random matrices (mean regret {f3(BMW['regret'])} dB), again exactly as brute force "
  f"with the same objective, against {pct(bench('random', 'exhaustive_w', mode='w')['top1'])} for unweighted brute "
  "force (results/selector_bench_w.md, figures/fig6_selector_accuracy_vs_cost_w.png).")
H("6.5 The methods we compared it with", 2)
P(f"To choose the method, and to know how good it is, we implemented {len(S.ROLE) - 1} other ways of choosing the "
  "format behind the same interface (`selector.py`) and ran them all on the same matrices. They fall into four "
  "groups:")
bullets([
    "**Brute-force references** quantize with all five formats and measure an error. They are what the chosen "
    "method is measured against, and include how ANT [3] and MoFQ [4] choose a data type.",
    "**Variants of the chosen method** remove one ingredient: the activation weighting, mantissa resolution, or "
    "everything except two summary statistics.",
    "**Extensions** handle cases outside the assignment: very large matrices (sampling with a certificate), "
    "correlated activations (the full H on sampled rows), and the `opt` scale policy.",
    "**Alternative methods** take a different route: error models, a learned rule, distribution fits, and a "
    "shape rule from the literature.",
])
DESC = {
    "exhaustive_w": "Quantize with all five formats; pick the smallest plain squared error, with no activation weighting.",
    "exhaustive_h": "Quantize with all five formats; measure the same activation-weighted error the chosen method computes.",
    "exhaustive_opt": "Quantize with all five formats using the per-block exponent search (`opt`); pick the smallest error. This is how ANT and MoFQ select.",
    "exhaustive_y": "Quantize W and X with all five formats; measure the full output error on the fit columns (all of H, not only its diagonal).",
    "closed_form": "Two statistics, no rounding: the energy at block maxima (sets the integer grids' error) and the energy per binade (sets the FP grids' error), in textbook formulas without flushing or clipping.",
    "closed_form_h": "The same two statistics with each value weighted by its channel energy.",
    "sketch_nb1_h": "The chosen method with 1 mantissa slice per octave instead of 64 (not exact).",
    "sketch_nb4_h": "The chosen method with 4 mantissa slices per octave (not exact).",
    "sketch_nb16_h": "The chosen method with 16 mantissa slices per octave (not exact).",
    "sketch": "The chosen method without activation weighting: every value counts equally.",
    "sketch_s10_h": "The chosen method on a random 10% of the blocks.",
    "certified_h": "Round randomly sampled blocks, doubling the sample until the leading format is 3 standard errors ahead of every other.",
    "certified_is_h": "As above, but sample blocks in proportion to their weighted energy (importance sampling), with each block maximum read exactly.",
    "sketch_opt_h": "An exact sketch for the `opt` policy instead of the class rule.",
    "certified_rows_y": "Brute force with the full H on a sample of rows, enlarged until the winner is 3 standard errors ahead; sees correlated activations.",
    "learned_tree": "A small decision tree trained on separate random matrices, from six cheap features (share of zeros, kurtosis, block-peak and binade statistics, share of values far below their block maximum).",
    "hist_pred_h": "Histogram predictor (`predictor.py`): the expected rounding error of each magnitude bin, with each block maximum treated exactly.",
    "m1": "Per-value expected error from a smooth rounding model (step²/12, plus flushing and clipping), no rounding.",
    "m1_h": "The same, activation-weighted.",
    "hybrid_h": "The activation-weighted error model, with the top octave of every block rounded exactly.",
    "exh_rows10": "Brute force on a random 10% of the rows.",
    "ant_rule": "A rule in the spirit of ANT [3]: integer, mantissa-heavy or exponent-heavy format chosen by the kurtosis of W.",
    "fit_gauss": "Fit a Gaussian to W, quantize synthetic matrices drawn from it.",
    "fit_laplace": "Fit a Laplace distribution to W, quantize synthetic matrices drawn from it.",
    "fit_t": "Fit a Student-t distribution (degrees of freedom from the kurtosis, as in [12]), quantize synthetic draws.",
}
check(set(DESC) == set(S.ROLE) - {M}, "every compared method has a description")
RNAME = {"reference": "brute-force reference", "ablation": "variant", "extension": "extension", "alternative": "alternative"}
rk = {"reference": 0, "ablation": 1, "extension": 2, "alternative": 3}
cmp_rows = [["chosen method", "`sketch_h`", "Exact sketch, activation-weighted (Sections 6.1–6.2).", pct(BM["top1"]),
             f3(BM["regret"]), f2(BM["ms"])]]
for a in sorted((a for a in S.ROLE if a != M), key=lambda a: rk[S.ROLE[a]]):
    b = bench("random", a)
    extra = f" Reads {pct(b['frac'], 0)} of the blocks (median)." if b["frac"] else ""
    cmp_rows.append([RNAME[S.ROLE[a]], f"`{a}`", DESC[a] + extra, pct(b["top1"]), f3(b["regret"]), f2(b["ms"])])
table(["Group", "Name in `selector.py`", "What it does", "Top-1", "Mean regret dB", "Median ms"], cmp_rows,
      cap=f"The chosen method and every method we compared it with: random suite, {N_RANDOM} matrices, W+X. "
          "In the W+X setting every method scores W and X^[T] and adds the two.", highlight=(0,), size=7.5,
      widths=[2.0, 2.9, 7.6, 1.3, 1.6, 1.4])
se = np.sqrt(BM["top1"] * (1 - BM["top1"]) / N_RANDOM)
close = [a for a in S.ROLE if a != M and abs(bench("random", a)["top1"] - BM["top1"]) <= 2 * se]
better = [a for a in S.ROLE if a != M and bench("random", a)["top1"] > BM["top1"] + 2 * se]
check(not better, "no compared method is better than the chosen one by more than two standard errors")
COARSE = ("sketch_nb4_h", "sketch_nb16_h")
check(all(a in COARSE or bench("random", a)["ms"] > BM["ms"] for a in close), "every close method is slower or coarser")
NBS = ("sketch_nb1_h", "sketch_nb4_h", "sketch_nb16_h")
check(all(bench("random", a)["top1"] < BM["top1"] for a in NBS), "fewer slices are less accurate on the random suite")
close_coarse = [a for a in close if a in COARSE]
cheap = ("closed_form_h", "learned_tree", "fit_gauss", "ant_rule")
check(all(bench("random", a)["top1"] < BM["top1"] - 0.1 for a in cheap), "cheap shortcuts are clearly worse")
check(bench("random", "certified_rows_y", "lowrank")["top1"] > bench("random", M, "lowrank")["top1"],
      "certified_rows_y helps on low rank")
check(big["certified agrees"] == "yes", "certified agrees on the largest matrix")
w_oc, w_ma = bench("random", "exhaustive_w", "outlier_chan")["top1"], bench("random", "exhaustive_w", "massive")["top1"]
bullets([
    f"**The activation weighting matters most.** Without it (the unweighted brute force and sketch) agreement falls "
    f"from {pct(BM['top1'])} to {pct(we['top1'])}, and to {pct(w_oc, 0)} with outlier channels and {pct(w_ma, 0)} "
    "with massive activations.",
    "**Mantissa resolution matters.** With 1 slice per octave agreement drops to "
    f"{pct(bench('random', 'sketch_nb1_h')['top1'])}, with 4 to {pct(bench('random', 'sketch_nb4_h')['top1'])} and "
    f"with 16 to {pct(bench('random', 'sketch_nb16_h')['top1'])}, against {pct(BM['top1'])} with 64. With 64 slices every rounding midpoint and "
    "saturation point of the five candidates lies on a slice boundary (they all lie on multiples of 1/32 of an "
    "octave), which is what makes the sketch exact; with 16 or fewer, some slices straddle a midpoint.",
    f"**Several methods are as accurate, none is clearly better.** With {N_RANDOM} matrices, top-1 agreement has a "
    f"standard error of about {100 * se:.1f} points. Within two standard errors of the chosen method are "
    + ", ".join(f"`{a}` ({pct(bench('random', a)['top1'])}, {f2(bench('random', a)['ms'])} ms)" for a in close)
    + f". All of them are slower than the chosen method ({f2(BM['ms'])} ms)"
    + (" except the coarser sketch" + ("es" if len(close_coarse) > 1 else "") + " with "
       + " and ".join(a.split("nb")[1].split("_")[0] for a in close_coarse) + " slices, which "
       + ("are" if len(close_coarse) > 1 else "is") + " not exact." if close_coarse else "."),
    "**Cheap shortcuts are not good enough.** The two-statistic closed form, the learned tree, a Gaussian fit and "
    "the ANT-style shape rule are right in "
    + ", ".join(pct(bench('random', a)['top1'], 0) for a in cheap) + " of cases respectively.",
    f"**The extensions do what they are for.** `certified_rows_y` sees correlated activations "
    f"({pct(bench('random', 'certified_rows_y', 'lowrank')['top1'])} on low-rank X); `certified_is_h` gives the same "
    f"pick as the chosen method on the largest matrix while rounding {big['blocks rounded by certified']} blocks; "
    "`sketch_opt_h` predicts the `opt` policy exactly.",
])
P("**Why this method.** It is exact for its objective, so it has no modelling error to tune or explain; it is as "
  "accurate as any method we compared within the noise of the test; it is among the fastest (only less accurate or "
  "inexact methods are faster); it reads the data once; and its sketch also explains why a format wins.")
H("6.6 Limitations of the method", 2)
bullets([
    "It keeps only the diagonal of H = X·X^[T]. Strongly correlated activations (low rank) need the full-H "
    "extension `certified_rows_y`.",
    "It ignores the product of the two rounding errors, ΔW·ΔX. On W_[C] this costs less than "
    f"{100 * dev:.1f}% of the error.",
    "It scores the class scale rule, the rule of our headline tables; the `opt` policy needs `sketch_opt_h`.",
    f"BFP6's exponent clamp is modelled exactly for values below 2^[{7 + 1 - S.DMIN}], that is up to {-S.DMIN} "
    "octaves above the top of its exponent range; larger values would need more offset bins.",
    "Our test matrices and activation families are synthetic, as are the course datasets.",
])

# =========================================================================== 7. Q3
H("7. Q3: outliers, sparsity and variation in scale", 1)
P(f"We changed {WC} in controlled ways and measured the held-out error of all five formats with W and X quantized "
  "(`q3_study.py`, results/q3_study_wx.md). Each point is the mean of 10 random seeds. \"Margin\" is the "
  "runner-up's error minus the winner's in pp (mean ± standard deviation over seeds), and \"(7/10)\" counts the "
  "seeds the winner won. The same experiments with weights only (results/q3_study_w.md) lead to the same "
  "conclusions; we note where the numbers differ.")
H("7.1 Outliers", 2)
cont = {k: v for k, v in OUT_T.items() if "continuous" in k}
low = [v for k, v in cont.items() if any(f", {s}σ" in k for s in (5, 10, 20))]
check(all(v["winner ocp (seeds), margin pp"].startswith("E2M3 (10/10)") for v in low), "E2M3 wins every cell up to 20σ")
low_w = [v for k, v in OUT_W.items() if "continuous" in k and any(f", {s}σ" in k for s in (5, 10, 20))]
check(all(v["winner ocp (seeds), margin pp"].startswith("E2M3 (10/10)") for v in low_w), "weights only: same up to 20σ")
sel = ["100, 5σ, continuous", "100, 20σ, continuous", "100, 50σ, continuous", "100, 100σ, continuous",
       "400, 20σ, continuous", "100, 20σ, exact", "100, 50σ, exact"]
table(["n, k, magnitudes", "Blocks with an outlier", "E3M2", "E2M3", "MXINT6", "BFP6", "Winner (seeds), margin pp",
       "BFP 16|4|5 (comparison)"],
      [[s, OUT_T[s]["blocks with an outlier"], OUT_T[s]["E3M2"], OUT_T[s]["E2M3"], OUT_T[s]["MXINT6"], OUT_T[s]["BFP6"],
        OUT_T[s]["winner ocp (seeds), margin pp"], OUT_T[s][CMPC]] for s in sel],
      cap="n weights of W_[C] replaced by ±k·σ·U(0.7, 1.3) (continuous magnitudes); the last two rows use exact "
          "magnitudes k·σ. W+X held-out e_[Y] (%), class rule. Last column: BFP with groups of 16, for comparison "
          "only (not a candidate, never counted as the winner).", size=8.5)
mf = MECH["flushed to 0"]
ex20, ex50, co20 = OUT_T["100, 20σ, exact"], OUT_T["100, 50σ, exact"], OUT_T["100, 20σ, continuous"]
check(num(ex20["E3M2"]) < num(co20["E3M2"]) - 1 and num(ex50["E2M3"]) > num(OUT_T["100, 50σ, continuous"]["E2M3"]) + 0.3,
      "exact magnitudes favour E3M2 at 20σ and hurt E2M3 at 50σ")
w_ex = OUT_W["100, 20σ, exact"]
w_ex_e3 = num(w_ex["E3M2"]) < num(w_ex["E2M3"])
s100 = [seeds(v["winner ocp (seeds), margin pp"]) for v in hi]
m100 = [margin(v["winner ocp (seeds), margin pp"]) for v in hi]
m_peak = max(cont, key=lambda k: margin(cont[k]["winner ocp (seeds), margin pp"]))
c100 = OUT_T["100, 100σ, continuous"]
bullets([
    f"**E2M3 wins every outlier case**: up to 20σ in 10 of 10 seeds, even with an outlier in "
    f"{OUT_T['400, 20σ, continuous']['blocks with an outlier']} of the blocks, and at 100σ in {min(s100)} to "
    f"{max(s100)} of 10 seeds. Its lead first grows (largest at {m_peak}: "
    f"{margin(cont[m_peak]['winner ocp (seeds), margin pp']):.2f} pp) and then shrinks to {min(m100):.2f}–"
    f"{max(m100):.2f} pp at 100σ.",
    f"**Mechanism: an outlier sets its block's scale, so its neighbours get a coarse grid.** With 100 outliers at "
    f"20σ, E2M3 flushes {mf['E2M3']} weights to zero (against {T6['MXFP6_E2M3']['zeroed']} without outliers), "
    f"MXINT6 {mf['MXINT6']} and BFP6 {mf['BFP6']}. At 100σ every format flushes the outlier's neighbours, so the "
    "error is roughly the energy of the 31 neighbours that share its block, and the formats move closer together.",
    "**Shorter groups would change the answer here.** BFP with groups of 16 (16|4|5, the last column) has only 15 "
    f"neighbours per outlier, and at 100σ it beats E2M3 in every row (for example {c100[CMPC]}% against "
    f"{c100['E2M3']}% at 100 × 100σ). It costs {float(G16['bits/elem']) - float(T2['BFP6']['bits/elem']):.2f} more "
    "bits per element than our BFP6 and is not one of the candidates.",
    f"**E3M2 never wins with continuous magnitudes.** It keeps the most neighbours ({mf['E3M2']} flushed), but its "
    "top binade is twice as coarse (2 mantissa bits instead of 3), and the outliers themselves carry most of the "
    "energy.",
    f"**A grid artifact to avoid.** With exact magnitudes k·σ, E3M2's error at 100 × 20σ falls from {co20['E3M2']}% "
    f"to {ex20['E3M2']}%, only because 20σ happens to be exactly representable in E3M2"
    + (f" (with weights only this even makes E3M2 \"win\", {w_ex['E3M2']}% against {w_ex['E2M3']}%)" if w_ex_e3 else "")
    + f"; at 50σ the outliers sit exactly on an E2M3 rounding midpoint, and E2M3 jumps to {ex50['E2M3']}%. We "
    "therefore use continuous magnitudes.",
])
figure("fig8_q3_outliers.png", "W+X: gap of E3M2 and BFP6 to E2M3 against the share of blocks containing an "
       "outlier, for outliers of 5σ, 20σ and 100σ. Above zero, E2M3 is better.", width=12)
H("7.2 Sparsity", 2)
sp = [z for z in SPARSE_M if z in SPARSE]
table(["Zeros", "E3M2", "E2M3", "MXINT6", "BFP6", "Winner (seeds), margin pp", "ρ_[32]", "φ",
       "BFP 16|4|5 (comparison)", "ρ_[16]"],
      [[z, SPARSE[z]["E3M2"], SPARSE[z]["E2M3"], SPARSE[z]["MXINT6"], SPARSE[z]["BFP6"],
        SPARSE[z]["winner ocp (seeds), margin pp"], SPARSE_M[z]["ρ32"], SPARSE_M[z]["φ"], SPARSE[z][CMPC],
        SPARSE_M[z]["ρ16"]] for z in sp],
      cap="A share of W_[C]'s weights set to zero at random positions. W+X held-out e_[Y] (%), class rule; ρ and φ "
          "are the statistics explained below. The last two columns are BFP with groups of 16, for comparison only.",
      size=7.5)
sp_w_lose = [z for z in SPARSE_W if winner_of(SPARSE_W[z]["winner ocp (seeds), margin pp"]) != "E2M3"]
sp_w_cross = min(sp_w_lose, key=num)
check(all(margin(SPARSE_W[z]["winner ocp (seeds), margin pp"]) < 0.05 for z in sp_w_lose),
      "weights only: the integer formats win by a hair past the crossover")
cmp_cross = next(z for z in sorted(SPARSE, key=num) if num(SPARSE[z][CMPC]) < num(SPARSE[z]["E2M3"]))
check(all(num(SPARSE[z][CMPC]) < num(SPARSE[z]["E2M3"]) for z in SPARSE if num(z) >= num(cmp_cross)),
      "BFP 16|4|5 beats E2M3 from cmp_cross on")
sgap = lambda z: num(SPARSE[z]["E2M3"]) - num(SPARSE[z]["MXINT6"])                      # noqa: E731
bullets([
    f"**E2M3 stays almost flat** ({SPARSE['0%']['E2M3']}% → {SPARSE['95%']['E2M3']}%) while **the integer grids fall "
    f"steadily** (BFP6 {SPARSE['0%']['BFP6']}% → {SPARSE['95%']['BFP6']}%). E2M3 wins in every seed up to {sp_all} "
    f"zeros, but its lead shrinks from {margin(SPARSE['0%']['winner ocp (seeds), margin pp']):.2f} pp to "
    f"{margin(SPARSE[max(sp_win, key=num)]['winner ocp (seeds), margin pp']):.2f} pp at {max(sp_win, key=num)}. From "
    f"{sp_cross} zeros MXINT6 wins, {f2(sgap(sp_cross))} pp ahead of E2M3 at {sp_cross} and {f2(sgap('95%'))} pp at "
    "95%, with BFP6 within 0.01 pp of MXINT6. With weights only the crossover comes earlier, at "
    f"{sp_w_cross} zeros.",
    "**Mechanism.** An FP grid's error is relative to each value's own binade, so it is proportional to "
    "φ = Σ 4^[floor(log₂|w|)] / Σ w^[2], which does not depend on sparsity "
    f"({SPARSE_M['0%']['φ']} at every level). An integer grid's step is set by each block's maximum, so its error "
    "is proportional to ρ = Σ_[blocks] n_[nz]·4^[R] / Σ w^[2] (n_[nz] non-zero values in a block whose maximum lies "
    f"in binade R). The maximum of fewer non-zero values is smaller relative to their energy, so ρ falls with "
    f"sparsity: for blocks of 32 from {SPARSE_M['0%']['ρ32']} to {SPARSE_M['50%']['ρ32']} at 50% and "
    f"{SPARSE_M['95%']['ρ32']} at 95% zeros.",
    f"**Group length matters again.** Groups of 16 have smaller maxima, so ρ_[16] is lower at every level "
    f"({SPARSE_M['0%']['ρ16']} with no zeros), and the BFP 16|4|5 comparison beats E2M3 from {cmp_cross} zeros on.",
])
figure("fig9_q3_sparsity.png", "W+X: held-out error against the share of weights set to zero (mean and spread over "
       "10 seeds). BFP6 and MXINT6 lie on top of each other; the dashed line is BFP with groups of 16, for "
       "comparison only.", width=12)
H("7.3 Variation in scale", 2)
WEAK = next(k for k in SCALEV[0] if k.startswith("weak-25% error"))
weak = lambda r, i: r[WEAK].split("/")[i].strip()                                      # noqa: E731
table(["Gain spread", "E3M2", "E2M3", "E4M1", "MXINT6", "BFP6", "Winner (seeds)",
       "Weakest 25%: E3M2 / E2M3 / E4M1 / MXINT6 / BFP6", "BFP6 clamped groups"],
      [[r["spread"], r["E3M2"], r["E2M3"], r["E4M1"], r["MXINT6"], r["BFP6"], r["winner ocp (seeds), margin pp"],
        " / ".join(weak(r, i) for i in range(5)), str(clamp32(r))] for r in SCALEV],
      cap="Random gains 10^[U(−a, a)] per output row or per input column. W+X global held-out e_[Y] (%), and the "
          f"error of the part of Y that comes from the weakest 25% of rows or input channels. BFP6 has {N_GROUPS} "
          "groups.", size=7.5)
check(all(winner_of(r["winner ocp (seeds), margin pp"]) == "E2M3" for r in SCALEV), "E2M3 wins globally at every spread")
check(all(r["weak-25% winner"] == "E4M1" for r in col_rows[2:]), "E4M1 keeps the weakest channels from 10^3")
bullets([
    "**Gains per output row** are absorbed by the block scale, because every block lies inside one row: the MX "
    f"formats barely change. BFP6 does: at a spread of 10^[4], {clamp32(row10k)} of its {N_GROUPS} groups "
    "leave the range of its 4-bit exponent. They are the groups of the rows with the largest gains, whose values "
    "saturate at 248, and those rows carry most of Y, so BFP6's error jumps to "
    f"{row10k['BFP6']}%.",
    "**Gains per input column** change the values inside every block. The global error still selects E2M3 at every "
    f"spread, because with a nearly white X the weak channels contribute little to Y.",
    "**Judged on the weakest channels, the ranking flips** to the formats with many exponent bits. The block scale "
    f"follows the strongest channel in each block: at 10^[2], E3M2 keeps the weak part at {weak(col_rows[1], 0)}% "
    f"against {weak(col_rows[1], 1)}% for E2M3, and from 10^[3] on only E4M1 keeps it near "
    f"{weak(col_rows[2], 2)}%. If the activations amplified such channels, the chosen method would see it: the "
    "weighting by channel energy is exactly what moves its objective toward them.",
])
figure("fig10_q3_scale.png", "W+X, gains per input column: the global held-out error (left) and the error of the "
       "weakest 25% of input channels (right) against the gain spread.", width=16)
H("7.4 Block organisation", 2)
P("With 100 values per row, the storage cost is set by the number of blocks per row, and each format pays for its "
  "own scale width (8 bits for MX, 4 bits for BFP). So we compared formats at equal storage, letting the block "
  "length vary (`block_study.py`, results/block_study_wx.md).")
bud = [b for b in ("6.08", "6.16", "6.24", "6.28", "6.32", "6.40", "6.50", "6.64", "7.00") if b in BUDGET]
check(BUDGET["7.00"]["best"] == "BFP6", "BFP with short groups wins at 7 bits")
table(["Budget (bits/element)", "Best format", "Block length", "Bits used", "Held-out e_[Y] %", "Best E2M3"],
      [[b, BUDGET[b]["best"], BUDGET[b]["block length"], BUDGET[b]["bits used"], BUDGET[b]["held-out Y err %"],
        BUDGET[b]["best E2M3"]] for b in bud],
      cap="The best (format, block length) within each storage budget, dataset C, W+X, class rule.",
      highlight=tuple(i for i, b in enumerate(bud) if b == "6.32"))
b_w = max((b for b in BUDGET_W if BUDGET_W[b]["best"] == "E2M3"), key=float)
bullets([
    f"**E2M3 stays best up to {b632} bits per element** ({b_w} with weights only). With more storage, BFP with short "
    f"groups wins, because its 4-bit exponent buys twice as many blocks per bit: {BUDGET['7.00']['held-out Y err %']}% "
    "at 7.0 bits.",
    "**At equal block length, the integer grids gain most from shorter blocks**: ρ falls from "
    f"{BLK['100']['ρ']} with one block per row to {BLK['2']['ρ']} with blocks of 2, while E2M3 changes little.",
    "**Some datasets would prefer other block lengths at the same storage:** "
    + "; ".join(f"{r['dataset']}: {r['best (format, length) within 6.32 bits']}" for r in BLK_AE)
    + ". We keep the class layouts (A1, A6).",
    "**Mixing formats between blocks** is excluded (A7). As a check with weights only: letting each 32-block "
    "choose its own MX format, at 2 extra bits per block, would lower the error by only "
    + ", ".join(f"{r['dataset']} {r['relative reduction']}" for r in MIXED) + " (relative).",
])
figure("fig11_block_tradeoff.png", "W+X: held-out error against storage as the block length varies; circles mark "
       "the class layouts.", width=12)
H("7.5 What this means for the selection", 2)
sel_hit = [(r, num(r["selector picks the held-out winner"])) for r in list(OUT_T.values()) + list(SPARSE.values()) + SCALEV]
miss = [r for r, h in sel_hit if h < 100]
ms_sd = [tuple(float(v) for v in re.findall(r"\d+\.\d+", r["winner ocp (seeds), margin pp"].split(")")[1])) for r in miss]
check(all(m < 0.15 or m < 2 * s for m, s in ms_sd), "the chosen method misses only close cases in Q3")
P("The winner moves with three measurable quantities, and each is part of the sketch the chosen method builds:")
bullets([
    "how much energy sits at the block maxima (ρ, which falls with sparsity and favours the integer grids)",
    "how much energy is flushed to zero next to outliers",
    "whether any block leaves BFP6's 4-bit exponent range",
])
P(f"Across the {len(sel_hit)} Q3 experiments, the chosen method (fitted on the fit columns) picks the held-out winner "
  f"in every seed in {len(sel_hit) - len(miss)} of them. The others are close cases near a crossover: the mean "
  "margin is below 0.15 pp, or smaller than two seed-to-seed standard deviations, so the winner itself changes "
  "from seed to seed.")

# =========================================================================== 8. verification
H("8. Verification and reproducibility", 1)
bullets([
    "**Three independent implementations agree bit for bit** on every candidate, variant and scale policy: "
    "`formats.py` (used for all results), `scale_opt.py` (written separately, with a different rounding method) "
    "and an exact-fraction reference implementation (`verification/reference_impl.py`, cross-checked by "
    "`verification/crosscheck.py`). See results/scale_study.md §1.",
    f"**Unit tests:** `test_formats.py` ({n_tests('test_formats.py')} tests), `test_selector.py` "
    f"({n_tests('test_selector.py')}, including the exactness of the chosen method in both settings), "
    f"`test_scale_opt.py` ({n_tests('test_scale_opt.py')}) and `verification/test_reference_impl.py` "
    f"({n_tests('verification/test_reference_impl.py')}).",
    "**Reproduction:** `python run_all.py` regenerates the tables T1–T9 (results/tables/) and the figures, and "
    "records the data hashes in results/RUN_MANIFEST.md. Each study regenerates its own results file "
    "(`bench_selector.py`, `bench_scaling.py`, `x_stress.py`, `q3_study.py`, `block_study.py`, `scale_study.py`; "
    "add `--mode w` for weights only), and `report/build_report.py` rebuilds this report from them. README.md "
    "gives the exact commands.",
    "**Guarded statements:** every sentence of this report that depends on a result is checked by the build "
    "script; if a rerun changed the result, the build would stop instead of printing a wrong sentence.",
])

# =========================================================================== 9. limitations
H("9. Limitations", 1)
bullets([
    f"{XT} is unseen. Our evidence that the choice transfers is the held-out half of {XC}, a bootstrap of its "
    f"columns and {len(xs_c) - 2} synthetic activation families; none of these is the real {XT}.",
    "The chosen method uses the diagonal of H and the class scale rule; correlated activations and the `opt` "
    "policy need the extensions of Section 6.5.",
    "The `opt` search covers four exponents per block; wider windows gave the same result on C "
    "(results/scale_study.md).",
    "All synthetic test families are our own construction, and all timings are NumPy on one laptop.",
])

# =========================================================================== references
H("References", 1)
for i, ref in enumerate([
    "CS611 Lecture 19 notebook and board notes (MX, MXINT and BFP conventions).",
    "Open Compute Project, \"OCP Microscaling Formats (MX) Specification\", v1.0, 2023; B. D. Rouhani et al., "
    "\"Microscaling Data Formats for Deep Learning\", arXiv:2310.10537, 2023.",
    "C. Guo et al., \"ANT: Exploiting Adaptive Numerical Data Type for Low-bit Deep Neural Network Quantization\", MICRO 2022.",
    "Y. Zhang et al., \"Integer or Floating Point? New Outlooks for Low-Bit Quantization on Large Language Models\" (MoFQ), arXiv:2305.12356, 2023.",
    "J. Lin et al., \"AWQ: Activation-aware Weight Quantization for LLM Compression and Acceleration\", MLSys 2024.",
    "E. Frantar et al., \"GPTQ: Accurate Post-Training Quantization for Generative Pre-trained Transformers\", ICLR 2023.",
    "T. Dettmers et al., \"LLM.int8(): 8-bit Matrix Multiplication for Transformers at Scale\", NeurIPS 2022.",
    "M. Sun et al., \"Massive Activations in Large Language Models\", arXiv:2402.17762, 2024.",
    "V. Malinovskii et al., \"Pushing the Limits of Large Language Model Quantization via the Linearity Theorem\" (HIGGS), arXiv:2411.17525, 2024.",
    "A. Kuzmin et al., \"FP8 Quantization: The Power of the Exponent\", NeurIPS 2022.",
    "M. van Baalen et al., \"FP8 versus INT8 for Efficient Deep Learning Inference\", arXiv:2303.17951, 2023.",
    "J. Dotzel et al., \"Learning from Students: Applying t-Distributions to Explore Accurate and Efficient Formats for LLMs\", ICML 2024.",
], start=1):
    p = d.add_paragraph()
    p.paragraph_format.left_indent, p.paragraph_format.first_line_indent = Cm(0.8), Cm(-0.8)
    p.paragraph_format.space_after = Pt(2)
    rich(p, f"[{i}]\t{ref}", size=9)

build_toc(toc_anchor)
d.save(OUT)
print(f"wrote {os.path.relpath(OUT, ROOT)}: {FIGNO[0]} figures, {TABNO[0]} tables, {len(HEADS)} headings")
