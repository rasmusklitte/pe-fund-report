"""Shared PPIM house style: palette, 16:9 slide helpers, branded table renderer."""
import numpy as np
import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.patches import Patch, Rectangle
from matplotlib.lines import Line2D
from matplotlib.transforms import blended_transform_factory
from pathlib import Path

PPIM_SAGE_DARK   = "#545B4A"   # dæmpet salviegrøn – primær accent / realiseret (DPI)
PPIM_SAGE_MID    = "#6E765E"   # salvie mellemtone
PPIM_SAGE_LIGHT  = "#C9CCB8"   # lys salvie – urealiseret (RVPI)
PPIM_OLIVE       = "#8B9377"   # støvet oliven
PPIM_BROWN       = "#5E4C3E"   # mørk jordbrun
PPIM_TAUPE       = "#8A7461"   # dæmpet taupe/brun
PPIM_TAUPE_LIGHT = "#B7A88F"   # lys taupe/tan
PPIM_TERRACOTTA  = "#A5734F"   # varm terracotta
PPIM_CREAM       = "#F4F2EB"   # cremet tone – kun zebra-rækker i tabellerne
PPIM_BEIGE       = "#E4DCC9"
PPIM_BLACK       = "#1E1E1E"
PPIM_RED_MUTED   = "#8C5A5A"   # dæmpet burgundy / tab
PPIM_RED_LIGHT   = "#C9A9A9"
PPIM_WHITE       = "#FFFFFF"

# 13,333" x 7,5" er PowerPoints "Widescreen (16:9)". save() gemmer UDEN bbox_inches="tight",
# så billedfilen altid har præcis 16:9 og kan lægges kant-til-kant på et slide.
SLIDE_W, SLIDE_H = 40 / 3, 7.5
SLIDE    = (SLIDE_W, SLIDE_H)
MX       = 0.055
SAVE_DPI = 240

BG = PPIM_WHITE
INK   = "#000000"
MUTED = "#000000"
LINE  = "#847E70"
GRID  = "#DBDCCF"

GILL       = "Gill Sans MT"
TITLE_FONT = GILL
BODY_FONT  = GILL
FONT_STACK = [GILL, "Segoe UI", "DejaVu Sans"]

import matplotlib.font_manager as _fm
if GILL not in {f.name for f in _fm.fontManager.ttflist}:
    print(f"ADVARSEL: '{GILL}' er ikke installeret – graferne falder tilbage til Segoe UI.")

plt.rcParams.update({
    "font.family":       FONT_STACK,
    "font.size":         10.5,
    "axes.titlesize":    13,
    "axes.labelsize":    11,
    "axes.labelcolor":   MUTED,
    "xtick.labelsize":   10,
    "ytick.labelsize":   10,
    "text.color":        INK,
    "xtick.color":       MUTED,
    "ytick.color":       MUTED,
    "axes.edgecolor":    LINE,
    "figure.figsize":    SLIDE,
    "figure.facecolor":  BG,
    "axes.facecolor":    BG,
    "savefig.facecolor": BG,
    "savefig.dpi":       SAVE_DPI,
    "figure.dpi":        110,
    "axes.axisbelow":    True,
    "axes.spines.top":   False,
    "axes.spines.right": False,
    "text.parse_math":   False,   # to "$" i samme tekst ("$1,2m ... $3,4m") er beløb, ikke en formel
})

TITLE  = dict(fontfamily=TITLE_FONT, color=INK)
KICKER = dict(fontfamily=BODY_FONT, color=INK)

# Sat af den importerende notebook efter dens config-/loader-celle, fordi add_header og
# source_line skal vise fondens navn og rapportdato uden at kende dem på importtidspunktet.
FUND = None
REPORT_DK = None
KILDE = None      # kildelinjens tekst; None -> "<FUND> track record" (PSCP-konventionen)
VALUTA = "EUR"    # ISO-kode for beløbene; styrer money() og tabellernes kildelinje
VALUTATEGN = {"EUR": "€", "USD": "$"}

HEADER_TOP_IN       = 1.62   # med undertitel
HEADER_TOP_IN_NOSUB = 1.32   # uden undertitel
BOTTOM_FRAC         = 0.075

HEADER_BAND = 1.35
MAX_ROW_H   = 0.40
MIN_ROW_H   = 0.235


def style_ax(ax):
    """Ren konsulent-akse: hvid flade, kun én hårfin bundlinje og lyse y-gitterlinjer."""
    ax.set_facecolor(BG)
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    ax.spines["bottom"].set_color(LINE)
    ax.spines["bottom"].set_linewidth(0.8)
    ax.tick_params(length=0, colors=INK)
    ax.grid(axis="y", color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)


def header_top_frac(subtitle=None, H=SLIDE_H):
    return 1 - (HEADER_TOP_IN if subtitle else HEADER_TOP_IN_NOSUB) / H


def add_header(fig, title, subtitle=None, x=None):
    x = MX if x is None else x
    H = fig.get_figheight()
    yf = lambda inch: 1 - inch / H
    fig.text(x, yf(0.34), f"{FUND}   ·   PETERSEN & PARTNERS", ha="left", va="top",
             fontsize=8.5, fontweight="bold", **KICKER)
    fig.text(x, yf(0.62), title, ha="left", va="top",
             fontsize=18, fontweight="bold", **TITLE)
    if subtitle:
        fig.text(x, yf(1.02), subtitle, ha="left", va="top", fontsize=11, color=MUTED)
        top_inch = HEADER_TOP_IN
    else:
        top_inch = HEADER_TOP_IN_NOSUB
    line_inch = top_inch - 0.28
    fig.add_artist(Line2D([x, x + 0.30], [yf(line_inch), yf(line_inch)], transform=fig.transFigure,
                          color=PPIM_SAGE_DARK, linewidth=1.2))
    return yf(top_inch)


def add_footer(fig, source, ax=None):
    left = ax.get_position().x0 if ax is not None else MX
    fig.text(left, 0.030, source, ha="left", va="bottom", fontsize=8.5, color=MUTED)
    fig.text(1 - MX, 0.030, "Petersen & Partners", ha="right", va="bottom",
             fontsize=9, color=MUTED, fontstyle="italic", fontfamily=TITLE_FONT)


def source_line(extra=""):
    return f"Kilde: {KILDE or f'{FUND} track record'} ({REPORT_DK})" + extra


def ref_pill(ax, y, text, accent=False):
    if accent:
        ax.axhline(y, color=PPIM_SAGE_DARK, linestyle="-", linewidth=1.5, zorder=3)
        fc, ec, tc = PPIM_SAGE_DARK, PPIM_SAGE_DARK, "white"
    else:
        ax.axhline(y, color=LINE, linestyle=(0, (4, 3)), linewidth=1.0, zorder=3)
        fc, ec, tc = BG, LINE, INK
    tr = blended_transform_factory(ax.transAxes, ax.transData)
    ax.text(1.008, y, text, transform=tr, va="center", ha="left",
            fontsize=9.5, color=tc, fontweight="bold", zorder=4, clip_on=False,
            bbox=dict(boxstyle="round,pad=0.32", fc=fc, ec=ec, linewidth=1.0))


def ref_pill_v(ax, x, text, accent=False):
    if accent:
        ax.axvline(x, color=PPIM_SAGE_DARK, linestyle="-", linewidth=1.5, zorder=3)
        fc, ec, tc = PPIM_SAGE_DARK, PPIM_SAGE_DARK, "white"
    else:
        ax.axvline(x, color=LINE, linestyle=(0, (4, 3)), linewidth=1.0, zorder=3)
        fc, ec, tc = BG, LINE, INK
    tr = blended_transform_factory(ax.transData, ax.transAxes)
    ax.text(x, 1.012, text, transform=tr, va="bottom", ha="center",
            fontsize=9.5, color=tc, fontweight="bold", zorder=4, clip_on=False,
            bbox=dict(boxstyle="round,pad=0.32", fc=fc, ec=ec, linewidth=1.0))


def legend_swatches(ax, entries, **kw):
    handles = [Patch(facecolor=c, edgecolor="none", label=l) for c, l in entries]
    lg = ax.legend(handles=handles, frameon=False, handlelength=1.1,
                   handleheight=1.1, borderpad=0.4, labelspacing=0.6, **kw)
    for t in lg.get_texts():
        t.set_color(INK)
    return lg


def save(fig, outfile):
    """Gem som et helt 16:9-slide. Bevidst UDEN bbox_inches="tight" (se house style docs)."""
    fig.set_size_inches(SLIDE_W, SLIDE_H)
    fig.savefig(outfile, dpi=SAVE_DPI, facecolor=fig.get_facecolor())
    print(f"gemt: {outfile}  ({round(SLIDE_W * SAVE_DPI)} x {round(SLIDE_H * SAVE_DPI)} px · 16:9)")


def _dk(v, dec=1):
    """Dansk talformat: komma-decimal, punktum-tusindtalsseparator."""
    return f"{v:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def money(v, dec=1):
    """Beløb i mio. i VALUTA: "€12,3m" / "$12,3m", ellers "12,3 mio. DKK"."""
    tegn = VALUTATEGN.get(VALUTA)
    fortegn = "-" if v < 0 else ""
    if tegn:
        return f"{fortegn}{tegn}{_dk(abs(v), dec)}m"
    return f"{fortegn}{_dk(abs(v), dec)} mio. {VALUTA}"


def mult(v):
    return _dk(v, 2) + "x"


def rows_per_slide(min_row_h=MIN_ROW_H):
    avail = (header_top_frac(True) - BOTTOM_FRAC) * SLIDE_H
    return max(1, int(avail / min_row_h - HEADER_BAND))


def _render_table_page(title, col_labels, rows, col_widths, aligns, outfile,
                       subtitle=None, highlight_row=None, source=None):
    w = np.asarray(col_widths, dtype=float)
    w = w / w.sum()
    edges = np.concatenate([[0.0], np.cumsum(w)])
    n = len(rows)

    fig, ax = plt.subplots(figsize=SLIDE)
    ax.set_axis_off()
    ax.set_xlim(0, 1); ax.set_ylim(0, 1)

    top_frac = add_header(fig, title, subtitle)
    fig.tight_layout(rect=[MX, BOTTOM_FRAC, 1 - MX, top_frac])
    add_footer(fig, source or source_line(f" · beløb i mio. {VALUTA}"), ax=ax)

    avail  = ax.get_position().height * SLIDE_H
    row_h  = min(MAX_ROW_H, avail / (n + HEADER_BAND))
    band   = row_h / avail
    ys     = np.concatenate([[1.0], 1.0 - HEADER_BAND * band - np.arange(n + 1) * band])

    fs_data = float(np.clip(row_h * 72 * 0.40, 6.6, 9.8))
    fs_head = fs_data + 0.7
    pad = 0.006

    def cx(j, ha):
        if ha == "left":
            return edges[j] + pad
        if ha == "right":
            return edges[j + 1] - pad
        return (edges[j] + edges[j + 1]) / 2

    yt, yb = ys[0], ys[1]
    ax.add_patch(Rectangle((0, yb), 1, yt - yb, transform=ax.transAxes,
                           facecolor=PPIM_SAGE_DARK, edgecolor="none", zorder=1))
    for j, lab in enumerate(col_labels):
        ax.text(cx(j, aligns[j]), (yt + yb) / 2, lab, transform=ax.transAxes,
                ha=aligns[j], va="center", color="white", fontsize=fs_head,
                fontweight="bold", zorder=3)

    for i, row in enumerate(rows):
        yt, yb = ys[i + 1], ys[i + 2]
        yc = (yt + yb) / 2
        is_hl = highlight_row is not None and i == highlight_row
        if is_hl:
            ax.add_patch(Rectangle((0, yb), 1, yt - yb, transform=ax.transAxes,
                                   facecolor=PPIM_SAGE_LIGHT, edgecolor="none", zorder=1))
            ax.add_line(Line2D([0, 1], [yt, yt], transform=ax.transAxes,
                               color=PPIM_SAGE_DARK, linewidth=1.2, zorder=2))
        elif i % 2 == 1:
            ax.add_patch(Rectangle((0, yb), 1, yt - yb, transform=ax.transAxes,
                                   facecolor=PPIM_CREAM, edgecolor="none", zorder=1))
        for j, val in enumerate(row):
            ax.text(cx(j, aligns[j]), yc, val, transform=ax.transAxes,
                    ha=aligns[j], va="center", color=INK, fontsize=fs_data,
                    fontweight="bold" if is_hl else "normal", zorder=3)

    save(fig, outfile)
    plt.show()


def render_table(title, col_labels, rows, col_widths, aligns, outfile,
                 subtitle=None, highlight_row=None, source=None, rows_per_page=None):
    """Tegner tabellen på ét eller flere 16:9-slides (`..._side1.png`, … hvis den ikke er på ét)."""
    per = rows_per_page or rows_per_slide()
    n = len(rows)
    pages = max(1, int(np.ceil(n / per)))
    per = int(np.ceil(n / pages))
    stem = outfile[:-4] if outfile.lower().endswith(".png") else outfile

    for p in range(pages):
        lo, hi = p * per, min((p + 1) * per, n)
        chunk = rows[lo:hi]
        hl = highlight_row - lo if highlight_row is not None and lo <= highlight_row < hi else None
        sub = subtitle
        if pages > 1:
            tag = f"side {p + 1} af {pages}"
            sub = f"{subtitle} · {tag}" if subtitle else tag
        _render_table_page(
            title, col_labels, chunk, col_widths, aligns,
            outfile if pages == 1 else f"{stem}_side{p + 1}.png",
            subtitle=sub, highlight_row=hl, source=source,
        )


def find_fil(navn, mapper=()):
    """Datafilen kopieres ikke nødvendigvis til den mappe, notebooken køres fra - prøv
    egen mappe, så hver ekstra kandidatmappe i `mapper` (fx en anden fonds datamappe)."""
    kandidater = [Path(navn)] + [Path(m) / navn for m in mapper]
    for p in kandidater:
        if p.exists():
            return p
    raise FileNotFoundError(f"Finder ikke {navn} - prøvede: {', '.join(str(p) for p in kandidater)}")
