"""Figures for the selector section (Aryan).   python make_selector_figures.py
Reads results/selector_bench.csv (bench_selector.py) and results/selector_scaling.md
(bench_scaling.py); writes figures/fig6_selector_accuracy_vs_cost.png and
figures/fig7_selector_scaling.png."""
import csv
import os
import re

import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt   # noqa: E402

import selector as S              # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
INK, INK2, GRID, SURF = "#0b0b0b", "#52514e", "#e4e3df", "#fcfcfb"
GRID2 = "#8f8d88"            # leader lines
# one colour per objective (<= 3 series in a scatter), reference palette slots 1-3
OBJ = {"weights only (way 1)": "#2a78d6", "H-weighted (way 2)": "#eb6834", "full Y on fit columns": "#1baf7a"}
LABEL = {"exhaustive_w": "brute force", "exhaustive_h": "brute force", "exhaustive_y": "brute force, full H",
         "sketch": "exact sketch", "sketch_h": "exact sketch", "sketch_nb1_h": "sketch 1 bin",
         "sketch_nb4_h": "4 bins", "sketch_nb16_h": "16 bins", "certified_is_h": "certified (IS)",
         "certified_h": "certified (uniform)", "sketch_s10_h": "sketch 10% sample", "closed_form": "closed form",
         "closed_form_h": "closed form", "akash_hist_h": "predictor.py", "m1": "M1", "m1_h": "M1",
         "hybrid_h": "hybrid", "exh_rows10": "brute 10% rows", "learned_tree": "learned tree",
         "ant_rule": "ANT shape rule", "fit_gauss": "Gaussian fit", "fit_laplace": "Laplace fit", "fit_t": "Student-t fit",
         "sketch_opt_h": "exact sketch (opt)", "exhaustive_opt": "brute force + scale\nsearch (ANT/MoFQ)",
         "certified_rows_y": "certified rows, full H"}


def objective(a):
    if a in ("exhaustive_y", "certified_rows_y"):
        return "full Y on fit columns"
    return "H-weighted (way 2)" if a.endswith("_h") else "weights only (way 1)"


def style(ax):
    ax.set_facecolor(SURF)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    for s in ("left", "bottom"):
        ax.spines[s].set_color(INK2)
    ax.tick_params(colors=INK2, labelsize=8)
    ax.grid(True, color=GRID, lw=0.8)
    ax.set_axisbelow(True)


def place_labels(fig, ax, pts, order):
    """Greedy label placement: for each point (most important first) try offsets on growing
    rings and keep the first whose text box overlaps no other label, no marker and stays inside
    the axes.  pts: {key: (x, y, text, fontkw)}."""
    fig.canvas.draw()
    r = fig.canvas.get_renderer()
    inv = ax.transData
    scale = fig.dpi / 72
    marks = {}
    for key, (x, y, _, _) in pts.items():
        px, py = inv.transform((x, y))
        m = (13 if key == S.METHOD else 4) * scale
        marks[key] = (px - m, py - m, px + m, py + m)
    axbox = ax.get_window_extent(r)
    placed = []
    if ax.get_legend():
        lb = ax.get_legend().get_window_extent(r)
        placed.append((lb.x0, lb.y0, lb.x1, lb.y1))
    dirs = []
    for deg in (0, 180, 90, 270, 45, 135, 315, 225, 22, 158, 338, 202, 68, 112, 292, 248):
        c, s = np.cos(np.radians(deg)), np.sin(np.radians(deg))
        dirs.append((c, s, "left" if c > 0.3 else "right" if c < -0.3 else "center",
                     "bottom" if s > 0.3 else "top" if s < -0.3 else "center"))
    for k in order:
        x, y, text, kw = pts[k]
        best = None                              # (overlap area, annotation, box)
        for rad in ((14, 18, 24) if k == S.METHOD else (6, 10, 15, 21, 28, 36, 46, 58)):
            for dx, dy, ha, va in dirs:
                t = ax.annotate(text, (x, y), xytext=(dx * rad, dy * rad), textcoords="offset points", ha=ha, va=va,
                                **kw, arrowprops=(dict(arrowstyle="-", color=GRID2, lw=0.7, shrinkA=1, shrinkB=3)
                                                  if rad > 6 else None))
                bb = t.get_window_extent(r).padded(1.5 * scale)
                box = (bb.x0, bb.y0, bb.x1, bb.y1)
                if not (bb.x0 >= axbox.x0 and bb.x1 <= axbox.x1 and bb.y0 >= axbox.y0 and bb.y1 <= axbox.y1):
                    t.remove()
                    continue
                over = sum(max(0, min(box[2], b[2]) - max(box[0], b[0])) * max(0, min(box[3], b[3]) - max(box[1], b[1]))
                           for b in placed + [v for q, v in marks.items() if q != k])
                if best is None or over < best[0]:
                    if best is not None:
                        best[1].remove()
                    best = (over, t, box)
                else:
                    t.remove()
                if over == 0:
                    break
            if best is not None and best[0] == 0:
                break
        if best is not None:                     # the first free spot, or else the least crowded one
            placed.append(best[2])


def fig6():
    rows = [r for r in csv.DictReader(open(os.path.join(HERE, "results", "selector_bench.csv"))) if r["suite"] == "random"]
    # closed_form and m1 without H sit on top of their H versions / of brute force: plot one of each
    names = [a for a in dict.fromkeys(r["approach"] for r in rows)
             if a not in ("closed_form", "m1", "fit_laplace", "sketch_opt_h", "certified_h")]
    stats = {}
    for a in names:
        A = [r for r in rows if r["approach"] == a]
        stats[a] = (1000 * np.median([float(r["seconds"]) for r in A]), 100 * np.mean([int(r["hit"]) for r in A]),
                    np.mean([float(r["regret_y_db"]) for r in A]))
    nx = len({r["matrix"].split("X=")[-1] for r in rows})
    fig, axs = plt.subplots(1, 2, figsize=(12, 5.4), facecolor=SURF)
    order = [S.METHOD] + sorted((a for a in names if a != S.METHOD), key=lambda a: -stats[a][1])
    allpts = []
    for ax, k, ylab in ((axs[0], 1, "top-1 agreement with held-out truth (%)"), (axs[1], 2, "mean regret (dB, log scale)")):
        style(ax)
        pts = {}
        for a in names:
            x, y = stats[a][0], max(stats[a][k], 1e-3)
            if a == S.METHOD:
                ax.scatter(x, y, s=260, marker="*", color=OBJ[objective(a)], edgecolor=INK, linewidth=1.0, zorder=5)
                pts[a] = (x, y, "OUR METHOD: exact sketch, H-weighted", dict(fontsize=8.5, color=INK, weight="bold"))
            else:
                ax.scatter(x, y, s=34, color=OBJ[objective(a)], edgecolor=SURF, linewidth=1.5, zorder=3, alpha=0.85)
                pts[a] = (x, y, LABEL[a], dict(fontsize=7, color=INK2))
        ax.set_xscale("log")
        if k == 2:
            ax.set_yscale("log")
        ax.set_xlim(0.2, 45)
        ax.set_xticks([0.3, 1, 3, 10, 30])
        ax.xaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _: f"{v:g}"))
        ax.set_ylim(*((0, 106) if k == 1 else (0.006, 14)))
        ax.set_xlabel("median time per 128x128 matrix (ms, log scale)", color=INK, fontsize=9)
        ax.set_ylabel(ylab, color=INK, fontsize=9)
        allpts.append((ax, pts))
    fig.suptitle(f"Accuracy vs cost of the format selector: 1000 random matrices, {nx} activation families",
                 color=INK, fontsize=11, x=0.01, ha="left", y=0.99, va="top")
    fig.text(0.01, 0.935, "\u2605 = our method, the one we submit: the same picks as brute force with the diagonal "
             "of H (exact, tested), from one pass over W and without quantizing it.\nEvery other point is an optional addition: brute-force "
             "references, ablations of the method, extensions and the alternatives we tried (docs/SELECTOR.md).",
             fontsize=8, color=INK2, ha="left", va="top", linespacing=1.5)
    handles = [plt.Line2D([], [], marker="*", ls="", color=OBJ["H-weighted (way 2)"], markeredgecolor=INK,
                          markersize=13, label="our method (H-weighted)")]
    handles += [plt.Line2D([], [], marker="o", ls="", color=c, markersize=6, label=l) for l, c in OBJ.items()]
    axs[1].legend(handles=handles, frameon=False, fontsize=8, loc="upper right", labelcolor=INK)
    fig.tight_layout(rect=(0, 0, 1, 0.89))
    for ax, pts in allpts:
        place_labels(fig, ax, pts, order)
    fig.savefig(os.path.join(HERE, "figures", "fig6_selector_accuracy_vs_cost.png"), dpi=160, facecolor=SURF)


def fig7():
    path = os.path.join(HERE, "results", "selector_scaling.md")
    if not os.path.exists(path):
        return
    pts = []
    for line in open(path):
        m = re.match(r"\| (\d+)x\d+ \| ([\d,]+) \| ([\d.]+) \| ([\d.]+) \| ([\d.]+) \| ([\d,]+) of ([\d,]+)", line)
        if m:
            pts.append([float(m.group(2).replace(",", ""))] + [float(m.group(i)) for i in (3, 4, 5)]
                       + [float(m.group(6).replace(",", "")) / float(m.group(7).replace(",", ""))])
    P = np.array(pts)[1:]                       # drop the 128x128 warm-up row (kernel tables are built there)
    fig, ax = plt.subplots(figsize=(6, 4), facecolor=SURF)
    style(ax)
    ser = (("brute force (quantize all 5)", 1, "#2a78d6"), ("exact sketch", 2, "#eb6834"), ("certified, importance-sampled", 3, "#1baf7a"))
    for lab, i, c in ser:
        ax.plot(P[:, 0], P[:, i], color=c, lw=2, marker="o", markersize=6, markeredgecolor=SURF, markeredgewidth=2, label=lab)
    for x, y, f in zip(P[:, 0], P[:, 3], P[:, 4]):
        ax.annotate(f"{100 * f:.2g}%", (x, y), xytext=(0, -11), textcoords="offset points", fontsize=7,
                    color=INK2, ha="center")
    ax.text(0.98, 0.04, "% = share of blocks the certified selector rounded", transform=ax.transAxes,
            fontsize=7, color=INK2, ha="right")
    ax.set_xscale("log")
    ax.set_yscale("log")
    ax.set_xlabel("weights in the matrix (log scale)", color=INK, fontsize=9)
    ax.set_ylabel("selection time (s, log scale)", color=INK, fontsize=9)
    ax.set_title("Selection cost vs matrix size (LLM-like weights, H-weighted)", color=INK, fontsize=10, loc="left")
    ax.legend(frameon=False, fontsize=8, loc="upper left", labelcolor=INK)
    ax.set_xlim(P[0, 0] / 1.6, P[-1, 0] * 1.6)
    ax.set_ylim(P[:, 1:4].min() / 2.5, P[:, 1:4].max() * 2)
    fig.tight_layout()
    fig.savefig(os.path.join(HERE, "figures", "fig7_selector_scaling.png"), dpi=160, facecolor=SURF)


if __name__ == "__main__":
    fig6()
    fig7()
