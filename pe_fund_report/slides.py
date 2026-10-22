"""Fondssider som redigerbare PowerPoint-objekter (tabeller, tekst, diagram) på PPIM-masteren.

Opbygningen følger "Portfolio Fund Memos" i ATP PEP's kvartalsrapport: fire kolonner øverst
(beskrivelse, fondsidentitet, investors tilsagn, nøgletal) og nederst enten de underliggende
fonde eller - hvor fonden ikke er indlæst på fondsniveau - NAV-historik og værdibro.
"""
import io
import uuid
from contextlib import contextmanager
from copy import deepcopy
from math import ceil

from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import (XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION,
                             XL_TICK_LABEL_POSITION)
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Inches, Pt

from . import style
from .data import (BRANCHEGRUPPER, GICS_STRUKTUR, IKKE_KLASSIFICERET,
                   NAV_SYNLIG, REGIONER, SEGMENT_OEVRIGE, SEKTORER)
from .data import uden_selskabsform as _uden_selskabsform
from .deck import BROED_FONT, L_TEKST, _fjern_tomme_ph, _ny_slide, _saet_ph

IKKE_OPLYST = "n/a"
INGEN_TABELSTIL = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"   # "No Style, No Grid"

# Indholdsfladen på Layout_basic_2 (tommer)
X0, X1 = 0.66, 11.04
Y_UNDERTITEL, Y_TOP, Y_BUND, Y_NOTER = 0.86, 1.22, 2.95, 5.90

V, H, C = PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.CENTER


def _rgb(hexfarve):
    return RGBColor.from_string(hexfarve.lstrip("#"))


SAGE, SAGE_LYS, CREAM = _rgb(style.PPIM_SAGE_DARK), _rgb(style.PPIM_SAGE_LIGHT), _rgb(style.PPIM_CREAM)
SORT, HVID, STREG = _rgb(style.PPIM_BLACK), _rgb(style.PPIM_WHITE), _rgb(style.LINE)
# søjlefarver; samme rollefordeling som i _historik_graf, så decket er konsistent
SAGE_MID, TERRA, TAUPE = _rgb(style.PPIM_SAGE_MID), _rgb(style.PPIM_TERRACOTTA), _rgb(style.PPIM_TAUPE_LIGHT)
OLIVEN, ROED = _rgb(style.PPIM_OLIVE), _rgb(style.PPIM_RED_MUTED)


# ---------- talformater (dansk) ----------

# den anonymiserede udgave ganger alle beløb med én skjult faktor; procenter, multipler og IRR
# er uændrede. Alle beløb, decket viser, går derfor gennem tusind, million, kroner eller _kt.
_SKALA = 1.0


@contextmanager
def beloebsskala(faktor):
    """Vis alle beløb ganget med faktor, så længe blokken kører."""
    global _SKALA
    gammel, _SKALA = _SKALA, faktor
    try:
        yield
    finally:
        _SKALA = gammel


def kroner(v):
    return IKKE_OPLYST if v is None else style._dk(v * _SKALA, 0)


def _kt(v):
    """Et beløb i tusinder som tal til en diagramserie."""
    return round(v * _SKALA / 1000)


def tusind(v):
    return IKKE_OPLYST if v is None else style._dk(v * _SKALA / 1000, 0)


def million(v):
    """Hele fondens beløb fylder for meget i tusinder - virksomhedskortene viser mio."""
    return IKKE_OPLYST if v is None else style._dk(v * _SKALA / 1_000_000, 0)


def pct(v, dec=1):
    """v er en andel (0,123 -> 12,3 %)."""
    return IKKE_OPLYST if v is None else style._dk(v * 100, dec) + " %"


def gange(v):
    return IKKE_OPLYST if v is None else style.mult(v)


def kvartal(d):
    return f"Q{(d.month - 1) // 3 + 1} {d:%y}"


def aendring(v, fmt=tusind):
    """En ændring med fortegn: +1.234, -1.234 og 0 uden fortegn."""
    if v is None:
        return IKKE_OPLYST
    tekst = fmt(v)
    if tekst.lstrip("-") == fmt(0):        # -0,00x efter afrunding er 0
        return fmt(0)
    return tekst if tekst.startswith("-") else "+" + tekst


# ---------- byggeklodser ----------

def _skriv(tf, tekst, pt, fed=False, farve=SORT, just=V, kursiv=False):
    """Skriv tekst i første afsnit; "\n" giver et linjeskift i samme afsnit."""
    para = tf.paragraphs[0]
    para.alignment = just
    for i, linje in enumerate(str(tekst).split("\n")):
        if i:
            # linjeskiftet får egen størrelse; ellers regner PowerPoint linjen som 18 pt
            para._p.add_br().get_or_add_rPr().set("sz", str(round(pt * 100)))
        run = para.add_run()
        run.text = linje
        run.font.name, run.font.size = BROED_FONT, Pt(pt)
        run.font.bold, run.font.italic = fed, kursiv
        run.font.color.rgb = farve
    return para


def _tekst(slide, x, y, w, h, tekst, pt, fed=False, farve=SORT, just=V, kursiv=False,
           anker=MSO_ANCHOR.TOP):
    tf = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)).text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anker
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    _skriv(tf, tekst, pt, fed, farve, just, kursiv)
    return tf


def _titel(slide, tekst):
    """Titel uden masterens venstre-indryk (0,1"), så den flugter med undertitel og indhold."""
    ph = _saet_ph(slide, 0, tekst)
    ph.text_frame.margin_left = 0
    return ph


def _logo(slide, logo, x_hoejre=X1, y_midt=0.47, maks_b=1.55, maks_h=0.36, x_midt=None):
    """Forvalterens logo, højrestillet mod x_hoejre - eller centreret om x_midt.

    Logoerne har vidt forskellige proportioner, fra kvadratiske til meget brede. I en
    tabelkolonne ville højrestilling derfor skubbe de smalle logoer ud til kanten, så der
    centreres, når x_midt er sat. Logoer, der kun findes til mørk baggrund, lægges på en
    mørk plade.
    """
    if not logo:
        return
    luft = 0.07 if logo["moerk_baggrund"] else 0
    skala = min((maks_b - 2 * luft) / logo["bredde"], (maks_h - 2 * luft) / logo["hoejde"])
    b, h = logo["bredde"] * skala, logo["hoejde"] * skala
    x = (x_midt - b / 2) if x_midt is not None else (x_hoejre - luft - b)
    y = y_midt - h / 2
    if luft:
        plade = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x - luft),
                                       Inches(y - luft), Inches(b + 2 * luft), Inches(h + 2 * luft))
        plade.adjustments[0] = 0.12
        plade.fill.solid()
        plade.fill.fore_color.rgb = SAGE
        plade.line.fill.background()
        plade.shadow.inherit = False
    slide.shapes.add_picture(io.BytesIO(logo["billede"]), Inches(x), Inches(y), Inches(b), Inches(h))


def _streg(slide, x1, y1, x2, y2, farve=SAGE, pt=0.75, stiplet=False):
    s = slide.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    s.line.color.rgb = farve
    s.line.width = Pt(pt)
    if stiplet:
        s.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    return s


def _bjaelke(slide, x, y, w, h, farve):
    """Et fyldt rektangel uden kant - byggeklodsen i de tegnede bjælkediagrammer."""
    f = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(max(w, 0.004)),
                               Inches(max(h, 0.004)))
    f.fill.solid()
    f.fill.fore_color.rgb = farve
    f.line.fill.background()
    f.shadow.inherit = False
    f.text_frame.text = ""
    return f


def _soejler(slide, x, y, w, h, poster, maks, fmt=tusind):
    """Søjler på fælles grundlinje: (label, værdi, farve) pr. post.

    maks er fælles for alle rækker på sliden, så søjlerne kan sammenlignes på tværs.
    Værdien står over søjlen og labelen under grundlinjen.
    """
    bredde = w / (len(poster) * 2 - 1)        # søjle, mellemrum, søjle, ...
    for i, (label, vaerdi, farve) in enumerate(poster):
        sx = x + i * 2 * bredde
        sh = 0 if not maks else max(0.0, (vaerdi or 0) / maks) * h
        _bjaelke(slide, sx, y + h - sh, bredde, sh, farve)
        _tekst(slide, sx - 0.12, y + h - sh - 0.17, bredde + 0.24, 0.16, fmt(vaerdi), 7,
               just=C)
        _tekst(slide, sx - 0.12, y + h + 0.03, bredde + 0.24, 0.16, label, 6.5, farve=SAGE,
               just=C)
    _streg(slide, x - 0.04, y + h, x + w + 0.04, y + h, farve=SAGE, pt=0.5)


def _vandfald(slide, x, y, w, h, trin, maks, fmt=tusind):
    """Vandfald af trin fra _tilsagnsbro(): hvert trin er (label, værdi, bund, ud, farve).

    bund/ud er niveauer på den fælles skala: bund er søjlens nederste kant, og ud er det
    niveau, broen står på, når søjlen er passeret. Den stiplede linje tegnes vandret i ud,
    så den rammer næste søjles kant - broen går op, fordi resttilsagnet er tilsagnet minus
    kaldt kapital.
    """
    bredde = w / (len(trin) * 2 - 1)          # søjle, mellemrum, søjle, ...
    niveau = lambda v: y + h - (0 if not maks else min(1.0, max(0.0, v / maks)) * h)
    for i, (label, vaerdi, bund, ud, farve) in enumerate(trin):
        sx = x + i * 2 * bredde
        top, bot = niveau(bund + abs(vaerdi or 0)), niveau(bund)
        if vaerdi:
            _bjaelke(slide, sx, top, bredde, bot - top, farve)
        # værdien skrives altid - et trin uden beløb skal stå som 0, ikke som ingenting
        _tekst(slide, sx - 0.12, top - 0.17, bredde + 0.24, 0.16, fmt(vaerdi), 7, just=C)
        _tekst(slide, sx - 0.12, y + h + 0.03, bredde + 0.24, 0.16, label, 6.5, farve=SAGE,
               just=C)
        if i < len(trin) - 1:
            _streg(slide, sx + bredde, niveau(ud), sx + 2 * bredde, niveau(ud),
                   farve=STREG, pt=0.75, stiplet=True)
    _streg(slide, x - 0.04, y + h, x + w + 0.04, y + h, farve=SAGE, pt=0.5)


def _noegletalsbjaelke(slide, x, y, w, h, vaerdi, maks, tekst, farve, grundlinje=False):
    """Én søjle med værdien over - brugt til TVPI- og IRR-kolonnen i Lange tilsagn.

    Søjlen fylder højst 72 % af højden, så etiketten altid har plads over den.
    grundlinje: en tydelig sort linje under søjlen, lidt bredere end den, i stedet for den
    lyse hårstreg.
    """
    if grundlinje:
        _streg(slide, x - 0.06, y + h, x + w + 0.06, y + h, farve=SORT, pt=1.25)
    else:
        _streg(slide, x, y + h, x + w, y + h, farve=STREG, pt=0.5)
    if vaerdi is not None and maks:
        bh = min(1.0, abs(vaerdi) / maks) * h * 0.72
        _bjaelke(slide, x, y + h - bh, w, bh, farve)
        _tekst(slide, x - 0.1, y + h - bh - 0.19, w + 0.2, 0.17, tekst, 7.5, fed=True, just=C)
    else:
        _tekst(slide, x - 0.1, y + h - 0.19, w + 0.2, 0.17, tekst, 7.5, just=C)


def _tabel(slide, x, y, w, raekker, bredder, just, raekke_h, pt, hoved=False, total=False,
           zebra=False, grupper=(), totaler=()):
    """Tabel uden PowerPoints tabelstil. raekker er lister af færdigformaterede tekster.

    grupper: rækkeindeks, der er gruppeoverskrifter (fed skrift på cremet bund).
    totaler: rækkeindeks, der får totalrækkens udseende ud over den sidste ved total=True.
    """
    gf = slide.shapes.add_table(len(raekker), len(bredder), Inches(x), Inches(y),
                                Inches(w), Inches(raekke_h * len(raekker)))
    tbl = gf.table
    tbl.first_row = tbl.horz_banding = False
    tbl._tbl.tblPr.find(qn("a:tableStyleId")).text = INGEN_TABELSTIL
    for j, b in enumerate(bredder):
        tbl.columns[j].width = Inches(w * b / sum(bredder))
    sidste = len(raekker) - 1
    for i, raekke in enumerate(raekker):
        tbl.rows[i].height = Inches(raekke_h)
        er_hoved = hoved and i == 0
        er_total = (total and i == sidste) or i in totaler
        er_gruppe = i in grupper
        for j, vaerdi in enumerate(raekke):
            celle = tbl.cell(i, j)
            celle.margin_left = celle.margin_right = Inches(0.05)
            celle.margin_top = celle.margin_bottom = Inches(0.01)
            celle.vertical_anchor = MSO_ANCHOR.MIDDLE
            if er_hoved:
                celle.fill.solid(); celle.fill.fore_color.rgb = SAGE
            elif er_total:
                celle.fill.solid(); celle.fill.fore_color.rgb = SAGE_LYS
            elif er_gruppe or (zebra and (i - hoved) % 2 == 1):
                celle.fill.solid(); celle.fill.fore_color.rgb = CREAM
            else:
                celle.fill.background()
            para = _skriv(celle.text_frame, vaerdi, pt, fed=er_hoved or er_total or er_gruppe,
                          farve=HVID if er_hoved else SORT, just=just[j])
            # en tom celle får ellers PowerPoints standard på 18 pt og gør rækken højere
            para._p.get_or_add_endParaRPr().set("sz", str(round(pt * 100)))
    return tbl


def _blok(slide, x, w, overskrift, par, bredder=(1.15, 1)):
    """Én af de tre nøgletalskolonner: overskrift, streg og label/værdi-rækker."""
    _tekst(slide, x, Y_TOP, w, 0.22, overskrift, 10.5, fed=True, farve=SAGE)
    _streg(slide, x, Y_TOP + 0.25, x + w, Y_TOP + 0.25)
    _tabel(slide, x - 0.05, Y_TOP + 0.31, w + 0.10, [[a, b] for a, b in par],
           bredder=list(bredder), just=[V, H], raekke_h=0.265, pt=9.5)


def _historik_graf(slide, x, y, w, h, historik):
    data = CategoryChartData(number_format="#,##0")
    data.categories = [kvartal(p["dato"]) for p in historik]
    serier = [("NAV", "nav", style.PPIM_SAGE_DARK, 2.25),
              ("Akk. indbetalt", "indbetalt", style.PPIM_TERRACOTTA, 1.5),
              ("Akk. udloddet", "udloddet", style.PPIM_TAUPE_LIGHT, 1.5)]
    for navn, noegle, _, _ in serier:
        data.add_series(navn, [_kt(p[noegle]) for p in historik])
    diagram = slide.shapes.add_chart(XL_CHART_TYPE.LINE, Inches(x), Inches(y), Inches(w),
                                     Inches(h), data).chart
    diagram.font.name, diagram.font.size = BROED_FONT, Pt(8)
    diagram.font.color.rgb = SORT
    diagram.has_legend = True
    diagram.legend.position = XL_LEGEND_POSITION.BOTTOM
    diagram.legend.include_in_layout = False
    for serie, (_, _, farve, bredde) in zip(diagram.plots[0].series, serier):
        serie.smooth = False
        serie.format.line.color.rgb = _rgb(farve)
        serie.format.line.width = Pt(bredde)
    y_akse, x_akse = diagram.value_axis, diagram.category_axis
    y_akse.has_major_gridlines = True
    y_akse.major_gridlines.format.line.color.rgb = _rgb(style.GRID)
    y_akse.major_gridlines.format.line.width = Pt(0.75)
    y_akse.format.line.fill.background()
    y_akse.tick_labels.number_format, y_akse.tick_labels.number_format_is_linked = "#,##0", False
    x_akse.format.line.color.rgb = STREG
    x_akse.has_major_gridlines = False
    return diagram


def _beholdningstabel(s, y, beh, f, investor_navn, maerker=None):
    """Tabel over underliggende fonde eller porteføljeselskaber. Returnerer y under tabellen.

    maerker er {kolonnenavn: fodnotemærke}, fx ved Branche og Dagsværdi."""
    maerker = maerker or {}
    hvem = f"{investor_navn}s andel" if beh["skaleret"] else "hele fonden"
    r_ = beh["raekker"]
    sum_ = lambda k: sum(r[k] or 0 for r in r_)
    kost = sum_("kostpris")
    tal = lambda r: [tusind(r["kostpris"]), tusind(r["dagsvaerdi"]), tusind(r["udloddet"]),
                     tusind(r["total"]), gange(r["tvpi"])]
    total = [tusind(kost), tusind(sum_("dagsvaerdi")), tusind(sum_("udloddet")),
             tusind(sum_("total")), gange(sum_("total") / kost) if kost else IKKE_OPLYST]
    if beh["type"] == "fonde":
        titel = "Underliggende fonde"
        raekker = [["Fond", "Vintage", "Kostpris", "Dagsværdi", "Udloddet", "Total værdi",
                    "TVPI", "IRR"]]
        raekker += [[r["navn"], r["vintage"] or IKKE_OPLYST] + tal(r)
                    + [IKKE_OPLYST if r["irr"] is None else pct(r["irr"] / 100)] for r in r_]
        raekker.append(["I alt", ""] + total + [""])
        bredder, just = [3.2, 0.9, 1.2, 1.2, 1.2, 1.2, 0.9, 0.9], [V, C, H, H, H, H, H, H]
    else:
        titel = "Porteføljeselskaber"
        # GP'er uden closing-dato pr. selskab (fx Seed Capital) oplyser kun investeringsåret
        har_closing = any(r["closing"] for r in r_)
        # branchen er GICS-sektoren, som på Branchefordeling - ikke forvalterens egen betegnelse
        raekker = [["Selskab", "Branche", "Closing" if har_closing else "Vintage", "Kostpris",
                    "Dagsværdi", "Realiseret", "Total værdi", "TVPI"]]
        raekker += [[r["navn"], r.get("industri") or IKKE_OPLYST,
                     f"{r['closing']:%m.%Y}" if r["closing"] else r["vintage"] or IKKE_OPLYST]
                    + tal(r) for r in r_]
        raekker.append(["I alt", "", ""] + total)
        bredder, just = [3.2, 1.9, 0.9, 1.1, 1.1, 1.1, 1.1, 0.9], [V, V, C, H, H, H, H, H]
    raekker[0] = [k + maerker.get(k, "") for k in raekker[0]]
    raekke_h = 0.185 if len(raekker) > 8 else 0.215
    _tekst(s, X0, y, X1 - X0, 0.22, f"{titel} – {hvem} (t{f['valuta']})", 10.5, fed=True,
           farve=SAGE)
    _tabel(s, X0, y + 0.28, X1 - X0, raekker, bredder=bredder, just=just, raekke_h=raekke_h,
           pt=8, hoved=True, total=True, zebra=True)
    return y + 0.28 + raekke_h * len(raekker) + 0.2


def _industrinote(beh):
    """Om branchekolonnen: GICS-sektor, og hvor mange selskaber der er foreløbigt klassificeret."""
    if beh["type"] != "selskaber":
        return ""
    foreloebige = sum(1 for r in beh["raekker"] if r.get("industri_foreloebig"))
    return ("Branche er selskabets GICS-sektor.")


def _navnote(f):
    """Hvorfor NAV afviger fra dagsværdien i alt i tabellen, og hvad total værdi er.

    Delene kommer fra data.nav_afstemning(); en del udelades, når den ikke ses i tabellens
    tusinder eller er under 2 % af forskellen (afrundingsstøj ved siden af en stor forskel).
    """
    a, beh = f.get("nav_afstemning"), f["beholdninger"]
    v = f"t{f['valuta']}"
    tekst = ""
    if a:
        retning = "lavere" if a["forskel"] < 0 else "højere"
        tekst = f"NAV er {tusind(abs(a['forskel']))} {v} {retning} end dagsværdien i alt"
        if a["ejerandel"] is not None:
            graense = max(NAV_SYNLIG, 0.02 * abs(a["forskel"]))
            dele = []
            if a["carry"] is not None and abs(a["carry"]) >= graense:
                dele.append(f"hensat carried interest {aendring(a['carry'])}")
            if abs(a["oevrige"]) >= graense:
                dele.append(f"fondens likvider og øvrige poster {aendring(a['oevrige'])}")
            if abs(a["ejerandel"]) >= graense:
                dele.append(f"investors NAV er {pct(a['nav_andel'], 2)} af fondens NAV, mens "
                            f"ejerandelen efter tilsagn er {pct(a['andel'], 2)} "
                            f"({aendring(a['ejerandel'])})")
            tekst += (": " + "; ".join(dele) + ".") if dele else "."
        else:
            tekst += (". Forvalteren oplyser ikke fondens samlede NAV, så forskellen kan ikke "
                      "opdeles; dagsværdien er beholdningerne før fondens likvider, omkostninger "
                      "og carried interest.")
        if a["beholdninger_dato"] != a["nav_dato"]:
            tekst += (f" Beholdningerne er pr. {a['beholdninger_dato']:%d.%m.%Y}, NAV pr. "
                      f"{a['nav_dato']:%d.%m.%Y}.")
    if any(r["udloddet"] for r in beh["raekker"]):
        tekst += (" Total værdi er dagsværdi plus positionernes udlodninger og kan ikke "
                  "sammenlignes direkte med NAV.")
    return tekst.strip()


def _beholdningskilde(beh, f):
    """Kilden til beholdningstabellen; branche og NAV-afstemning har deres egne fodnoter."""
    if not beh["skaleret"]:
        dato = f" pr. {beh['dato']:%d.%m.%Y}" if beh["dato"] != f["nav_dato"] else ""
        return (f"Porteføljeselskaberne{dato} er fra forvalterens rapport; ejerandelen kan ikke "
                "beregnes, da fondens samlede tilsagn ikke er oplyst.")
    return ((f"Beholdningerne er pr. {beh['dato']:%d.%m.%Y}. "
             if beh["dato"] != f["nav_dato"] else "")
            + "Kostpris er tilsagn × indbetalt andel, hvor forvalteren ikke oplyser den"
            + ("; TVPI og IRR pr. fond er forvalterens egne." if beh["type"] == "fonde"
               else "."))


GL_TOP = 10               # positioner i feedertabellen; resten samles i "Øvrige"
GL_TABEL_B = 6.55         # tabellens bredde; fordelingen får resten
POSITIONSTYPER = {"SECONDARY": "Secondary", "CO_INVESTMENT": "Co-investment",
                  "PRIMARY": "Primær", "FUND": "Fond"}


# sidst i feederens bjælker: det, der ikke er en branche eller en bestemt delregion
FEEDER_SIDST = {IKKE_KLASSIFICERET, "Ikke klassificeret af forvalteren",
                "Diversificeret (fondsandele)"}


def _feederfordeling(beh, f):
    """Branche- og geografiblokken på en feeders side som [(overskrift, [{navn, andel}])].

    Branche er GICS-sektoren og geografi regionen, fordelt som på siderne
    Branchefordeling og Geografisk eksponering (f["brancher"], f["regioner"]). Uden dem
    bruges forvalterens egen fordeling eller fordelingen på typer af positioner.
    """
    blokke = []
    for art, overskrift in (("brancher", "Branche"), ("regioner", "Geografi")):
        b = f.get(art)
        if not b:
            continue
        andele = {}
        for x in b["linjer"]:
            navn = x["top"]
            andele[navn] = andele.get(navn, 0.0) + x["andel"]
        blokke.append((overskrift, [dict(navn=n, andel=a) for n, a in sorted(
            andele.items(), key=lambda x: (x[0] in FEEDER_SIDST, -x[1]))]))
    return blokke or beh.get("fordeling") or []


def _gennemlysningsblok(s, y, beh, f, investor_navn, bund=Y_NOTER):
    """En feeders side: den underliggende fonds største positioner og dens fordeling.

    Beløbene er investors NAV i feederen fordelt efter dagsværdi (data._feeder_beholdninger),
    så dagsværdien i i alt-rækken er feederens NAV.
    """
    via = _uden_selskabsform(beh["via"])
    _tekst(s, X0, y, X1 - X0, 0.22,
           f"Gennemlysning af {via} – {investor_navn}s andel (t{f['valuta']})", 10.5, fed=True,
           farve=SAGE)
    aktive = [r for r in beh["raekker"] if r["dagsvaerdi"] > 0]
    top, oevrige = aktive[:GL_TOP], aktive[GL_TOP:] + [r for r in beh["raekker"]
                                                       if r["dagsvaerdi"] <= 0]
    sum_ = lambda rk, k: sum(r[k] or 0 for r in rk)
    multiple = lambda rk: ((sum_(rk, "dagsvaerdi") + sum_(rk, "udloddet")) / sum_(rk, "kostpris")
                           if sum_(rk, "kostpris") else None)
    if beh["type"] == "fonde":
        raekker = [["Position", "Type", "Vintage", "Kostpris", "Dagsværdi", "Udloddet", "TVPI"]]
        raekker += [[r["navn"], POSITIONSTYPER.get(r["type"], r["type"] or ""),
                     r["vintage"] or IKKE_OPLYST, tusind(r["kostpris"]), tusind(r["dagsvaerdi"]),
                     tusind(r["udloddet"]), gange(multiple([r]))] for r in top]
        hale = lambda navn, rk: [navn, "", "", tusind(sum_(rk, "kostpris")),
                                 tusind(sum_(rk, "dagsvaerdi")), tusind(sum_(rk, "udloddet")),
                                 gange(multiple(rk))]
        bredder, just = [2.45, 0.95, 0.6, 0.75, 0.75, 0.7, 0.55], [V, V, C, H, H, H, H]
    else:
        raekker = [["Selskab", "Instrument", "Investeret", "Kostpris", "Dagsværdi", "Multiple"]]
        raekker += [[r["navn"], r.get("instrument") or IKKE_OPLYST,
                     f"{r['closing']:%m.%Y}" if r["closing"] else IKKE_OPLYST,
                     tusind(r["kostpris"]), tusind(r["dagsvaerdi"]), gange(multiple([r]))]
                    for r in top]
        hale = lambda navn, rk: [navn, "", "", tusind(sum_(rk, "kostpris")),
                                 tusind(sum_(rk, "dagsvaerdi")), gange(multiple(rk))]
        bredder, just = [1.75, 1.45, 0.8, 0.8, 0.8, 0.65], [V, V, C, H, H, H]
    if oevrige:
        raekker.append(hale(f"Øvrige ({len(oevrige)})", oevrige))
    raekker.append(hale("I alt", beh["raekker"]))
    ty = y + 0.30
    raekke_h = min(0.215, (bund - ty - 0.05) / len(raekker))
    _tabel(s, X0, ty, GL_TABEL_B, raekker, bredder=bredder, just=just, raekke_h=raekke_h, pt=8,
           hoved=True, total=True, zebra=True)

    # højre: GP'ens fordeling (branche og geografi) eller fordelingen på typer af positioner
    hx = X0 + GL_TABEL_B + 0.35
    hw = X1 - hx
    fy = ty
    fakta = beh.get("fakta") or {}
    linje = (f"{fakta['NUMBER_OF_FUNDS']:.0f} underliggende fonde · "
             f"{fakta['NUMBER_OF_COMPANIES']:.0f} selskaber"
             if fakta.get("NUMBER_OF_FUNDS") and fakta.get("NUMBER_OF_COMPANIES")
             else f"{len(aktive)} {'positioner' if beh['type'] == 'fonde' else 'selskaber'}")
    _tekst(s, hx, fy, hw, 0.18, linje, 9, fed=True, farve=SAGE)
    _streg(s, hx, fy + 0.22, X1, fy + 0.22)
    fy += 0.30
    blokke = _feederfordeling(beh, f)
    maks = max([g["andel"] or 0 for _, grupper in blokke for g in grupper] + [0.01])
    for overskrift, grupper in blokke:
        # to blokke à fire linjer plus Øvrige er, hvad der er plads til over noterne
        fy = _fordelingsbjaelker(s, hx, fy, hw, overskrift, grupper, maks, raekke_h=0.19,
                                 label_b=1.45, maks_grupper=4 if len(blokke) > 1 else 6) + 0.08


def _gennemlysningskilde(beh, f, investor_navn):
    via = _uden_selskabsform(beh["via"])
    tekst = (f"Positionerne er pr. {beh['dato']:%d.%m.%Y}." if beh["dato"] != f["nav_dato"]
             else "")
    if beh["type"] != "fonde":
        tekst += " Trancher i samme selskab er lagt sammen; multiple er (dagsværdi + udloddet) / kostpris."
    blokke = [b for b, _ in _feederfordeling(beh, f)]
    if f.get("brancher") or f.get("regioner"):
        tekst += (" Branche (GICS-sektor) og geografi er fordelt efter forvalterens fordeling "
                  "eller beholdningernes dagsværdi, i procent af dagsværdien.")
        # nævn kun en fordeling, der er opgjort på en anden dato end fondens NAV
        afvig = {navn: b["dato"] for navn, b in (("Branchefordelingen", f.get("brancher")),
                                                 ("Den geografiske fordeling", f.get("regioner")))
                 if b and b["dato"] != f["nav_dato"]}
        if len(afvig) == 2 and len(set(afvig.values())) == 1:
            tekst += (f" Branche- og geografifordelingen er pr. "
                      f"{next(iter(afvig.values())):%d.%m.%Y}.")
        else:
            tekst += "".join(f" {navn} er pr. {d:%d.%m.%Y}." for navn, d in afvig.items())
    elif "Type" in blokke:
        tekst += " Fordelingen på typer er beregnet af positionernes dagsværdi."
    elif blokke:
        tekst += " Branche- og geografifordelingen er forvalterens, i procent af dagsværdien."
    kredit = (beh.get("fakta") or {}).get("CREDIT_LINE_DRAWN")
    if kredit:
        tekst += (f" {via} har trukket {million(kredit)} mio. {f['valuta']} på sin "
                  "kreditfacilitet; den gearing er ikke vist som ekstra eksponering.")
    return tekst.strip()


BRO_X = 7.55              # værdibroens venstre kant på fondssiden


def _historik_og_bro(s, y, f, bund=Y_NOTER):
    """NAV-graf til venstre og værdibro til højre i den plads, der er tilbage over noterne."""
    _tekst(s, X0, y, BRO_X - X0 - 0.3, 0.22,
           f"Udvikling i NAV og kapitalstrømme (t{f['valuta']})", 10.5, fed=True, farve=SAGE)
    _historik_graf(s, X0 - 0.1, y + 0.22, BRO_X - X0 - 0.2, bund - y - 0.27, f["historik"])
    _vaerdibro(s, y, f, bund=bund)


def _vaerdibro(s, y, f, bro_x=BRO_X, bund=Y_NOTER):
    """Tabellen fra indbetalt kapital til NAV i højre side af fondssiden."""
    _tekst(s, bro_x, y, X1 - bro_x, 0.22, f"Fra indbetalt kapital til NAV (t{f['valuta']})",
           10.5, fed=True, farve=SAGE)
    bro = [[navn, tusind(v)] for navn, v in f["bro"]]
    raekke_h = min(0.27, (bund - y - 0.35) / len(bro))
    _tabel(s, bro_x, y + 0.28, X1 - bro_x, bro, bredder=[2.6, 1], just=[V, H],
           raekke_h=raekke_h, pt=9 if raekke_h >= 0.25 else 8, total=True, zebra=True)


EJD_MIN = 0.02            # ejendomme under 2 % af GAV samles i "Øvrige" i waterfallen


def _ejendomsskala(beh, f):
    """Investors andel af fondens ejendomme og gæld: investors NAV / fondens NAV.

    None, når fondens NAV ikke er oplyst, eller ejendommene er fra en anden dato end
    investors NAV; så viser waterfallen hele fonden.
    """
    fonds_nav = (beh.get("fakta") or {}).get("FUND_NAV")
    if not fonds_nav or beh["dato"] != f["nav_dato"]:
        return None
    return f["n"]["nav"] / fonds_nav


def _ejendomstrin(beh, skala):
    """Waterfallens trin fra ejendommene til NAV.

    Med ejendomme pr. linje (Home.Earth) bygger hver ejendom op til GAV; uden (CRED oplyser
    kun den samlede værdi) står GAV som én søjle. Kendes fondens aktiver i alt, deles vejen
    til NAV i likvider og øvrige aktiver op og gæld ned; ellers er det én nettopost.
    """
    r_ = beh["raekker"]
    fakta = beh.get("fakta") or {}
    gav = sum(r["dagsvaerdi"] for r in r_) if r_ else fakta.get("GROSS_ASSET_VALUE") or 0.0
    store = [r for r in r_ if r["dagsvaerdi"] >= EJD_MIN * gav]
    smaa = [r for r in r_ if r["dagsvaerdi"] < EJD_MIN * gav]
    trin = [(r["navn"], r["dagsvaerdi"] * skala, OLIVEN, False) for r in store]
    if smaa:
        trin.append((f"Øvrige ({len(smaa)})", sum(r["dagsvaerdi"] for r in smaa) * skala,
                     OLIVEN, False))
    trin.append(("Ejendomme (GAV)", gav * skala, SAGE_MID, True))
    fonds_nav, aktiver = fakta.get("FUND_NAV"), fakta.get("TOTAL_ASSETS")
    if fonds_nav is None:
        return trin
    if aktiver:
        trin += [("Likvider og øvrige aktiver", (aktiver - gav) * skala, OLIVEN, False),
                 ("Gæld", (fonds_nav - aktiver) * skala, ROED, False)]
    else:
        trin.append(("Gæld m.m., netto", (fonds_nav - gav) * skala, ROED, False))
    return trin + [("NAV", fonds_nav * skala, SAGE, True)]


def _ejendomsblok(s, y, beh, f, investor_navn, bund=Y_NOTER):
    """En ejendomsfonds side: waterfall fra ejendommene til NAV i hele bredden.

    Ejendommene er fondens værdier før gæld (data._ejendomme). Ganget med investors andel af
    fondens NAV ender waterfallen i investors NAV, så gearingen ses: ejendommene fylder mere
    end NAV, og gælden trækker forskellen fra. Værdibroen fra indbetalt kapital står i Excel.
    """
    skala = _ejendomsskala(beh, f)
    hvem = f"{investor_navn}s andel" if skala else "hele fonden"
    b = X1 - X0
    _tekst(s, X0, y, b, 0.22, f"Fra ejendomme til NAV – {hvem} (t{f['valuta']})", 10.5,
           fed=True, farve=SAGE)
    _streg(s, X0, y + 0.27, X0 + b, y + 0.27)
    trin = _ejendomstrin(beh, skala or 1.0)
    _nav_bro(s, X0, y + 0.6, b, bund - y - 1.05, trin, 0.0,
             max(v for _, v, _, er_niveau in trin if er_niveau), maks_bredde=1.0)


def _ejendomskilde(beh, f, investor_navn):
    skala = _ejendomsskala(beh, f)
    andel = "" if skala else ", hele fonden"
    fakta = beh.get("fakta") or {}
    gaeld = ("Gæld er fondens aktiver i alt fratrukket NAV." if fakta.get("TOTAL_ASSETS")
             else "Gæld m.m. er fondens NAV fratrukket ejendommenes værdi (GAV).")
    grundlag = "; ".join(f"{r['navn']} {r['sektor'].lower()}" for r in beh["raekker"]
                         if r["sektor"])
    dato = f" pr. {beh['dato']:%d.%m.%Y}" if beh["dato"] != f["nav_dato"] else ""
    return (f"Ejendommene og gælden er forvalterens værdier{dato}{andel}. "
            + gaeld
            + (f" Værdigrundlag: {grundlag}." if grundlag
               else " Forvalteren oplyser ikke værdien pr. ejendom."))


# ---------- fondssiden ----------

def fondsside(prs, f, investor_navn):
    """Én slide for én fond. f er en dict fra data.fond()."""
    n = f["n"]
    s = _ny_slide(prs, L_TEKST)
    _titel(s, f["fund_name"])
    _fjern_tomme_ph(s)
    _logo(s, f["logo"])
    _tekst(s, X0, Y_UNDERTITEL, X1 - X0, 0.22,
           f"{investor_navn}  ·  pr. {f['nav_dato']:%d.%m.%Y}  ·  beløb i t{f['valuta']}", 9.5)

    # øverste halvdel: beskrivelse + tre nøgletalsblokke adskilt af lodrette streger
    kol = [(X0, 3.25), (4.22, 2.45), (6.98, 1.95), (9.24, X1 - 9.24)]
    # mærkerne sættes, før noternes længde bestemmer, hvor indholdet skal slutte
    fod = Fodnoter("Kilde: Forvalterrapporter og egne beregninger.")
    forankrede = {"andel", "indbetalt", "resttilsagn", "nav", "irr", "beskrivelse"}
    m = lambda felt: fod.maerker(x for x in f["noter"] if getattr(x, "felt", None) == felt)
    beskrivelse_m = m("beskrivelse")
    _tekst(s, kol[0][0], Y_TOP, kol[0][1], 0.22, f"Kort beskrivelse af fonden{beskrivelse_m}",
           10.5, fed=True, farve=SAGE)
    _streg(s, kol[0][0], Y_TOP + 0.25, kol[0][0] + kol[0][1], Y_TOP + 0.25)
    _tekst(s, kol[0][0], Y_TOP + 0.34, kol[0][1], Y_BUND - Y_TOP - 0.34,
           f["beskrivelse"] or IKKE_OPLYST, 9.5)
    for x, _ in kol[1:]:
        _streg(s, x - 0.155, Y_TOP, x - 0.155, Y_BUND - 0.1, farve=SAGE_LYS)

    _blok(s, *kol[1], "Fondsidentitet", [
        ("Geografi", f["geografi"] or IKKE_OPLYST),
        ("Strategi", f["segment"] or IKKE_OPLYST),
        ("Fondsvaluta", f["valuta"]),
        ("Fondens NAV", tusind(n["fonds_nav"])) if n["fondsstoerrelse"] is None
        and n.get("fonds_nav") else ("Fondsstørrelse", tusind(n["fondsstoerrelse"])),
        ("Vintage", f["vintage"] or IKKE_OPLYST)], bredder=(1, 1.45))
    _blok(s, *kol[2], f"{investor_navn}s tilsagn", [
        ("Tilsagn", tusind(n["tilsagn"])),
        ("Ejerandel" + m("andel"), pct(n["andel"], 2)),
        ("Indbetalt kapital" + m("indbetalt"), tusind(n["indbetalt"])),
        ("PICC", pct(n["picc"])),
        ("Resttilsagn" + m("resttilsagn"), tusind(n["resttilsagn"]))])
    gp_m = ""
    if n["gp_irr"] is not None or n["gp_tvpi"] is not None:
        gp = [t for t in (f"TVPI {gange(n['gp_tvpi'])}" if n["gp_tvpi"] is not None else None,
                          f"IRR {pct(n['gp_irr'] / 100)}" if n["gp_irr"] is not None else None)
              if t]
        gp_m =fod.maerke("Forvalterens egne nøgletal: " + ", ".join(gp) + ".")
    _blok(s, *kol[3], f"Nøgletal{gp_m}", [
        ("Udlodninger", tusind(n["udloddet"])),
        ("NAV" + m("nav"), tusind(n["nav"])),
        ("DPI", gange(n["dpi"])),
        ("TVPI", gange(n["tvpi"])),
        ("IRR" + m("irr"), pct(n["irr"]))])

    beh = f["beholdninger"]
    tabel_m = {}
    if beh and beh.get("via"):
        fod.generel(_gennemlysningskilde(beh, f, investor_navn))
    elif beh and beh["type"] == "ejendomme":
        fod.generel(_ejendomskilde(beh, f, investor_navn))
    elif beh:
        fod.generel(_beholdningskilde(beh, f))
        tabel_m = {"Branche": fod.maerke(_industrinote(beh)),
                   "Dagsværdi": fod.maerke(_navnote(f)) if beh["skaleret"] else ""}
    else:
        fod.generel("Forvalteren oplyser ikke fondens porteføljeselskaber.")
    # resten står som generelle noter; en afvigende dato (fx primo for udviklingen) skal altid
    # stå i noterne, også hvor tallet, den gælder, ikke vises
    fod.generel([x for x in f["noter"] if getattr(x, "felt", None) not in forankrede])
    noter_y = _noter_y(fod.tekst())
    bund = noter_y - 0.05

    # nederste halvdel: beholdninger, hvor de findes; historik og værdibro, hvor der er plads
    y = Y_BUND + 0.05
    if beh and beh.get("via"):
        # en feeder: den underliggende fonds positioner i stedet for historik og værdibro
        _gennemlysningsblok(s, y, beh, f, investor_navn, bund=bund)
        y = bund
    elif beh and beh["type"] == "ejendomme":
        # en ejendomsfond: ejendommene før gæld ved siden af historik og værdibro
        _ejendomsblok(s, y, beh, f, investor_navn, bund=bund)
        y = bund
    elif beh:
        y = _beholdningstabel(s, y, beh, f, investor_navn, tabel_m)
    # en tabel over underliggende fonde står alene; graf og værdibro kun ved selskaber og fonde
    # uden beholdninger, og kun hvor der er plads
    if bund - y >= 1.5 and not (beh and beh["type"] == "fonde"):
        _historik_og_bro(s, y, f, bund=bund)
    fod.tegn(s)
    return s


NOTE_TEGN = 200           # tegn pr. linje i noterne ved 7 pt i hele bredden; hellere for lavt
NOTE_LINJE = 0.118        # linjehøjde ved 7 pt
NOTE_BUND = 6.28          # noterne skal slutte over Petersen & Partners-linjen og sidefoden


def _noter_y(tekst, y=Y_NOTER):
    """Hvor noterne skal starte: y, eller højere, hvis de ellers løber ned i logoet."""
    return min(y, NOTE_BUND - ceil(len(tekst) / NOTE_TEGN) * NOTE_LINJE)


SUPERSKRIFT = str.maketrans("0123456789", "⁰¹²³⁴⁵⁶⁷⁸⁹")


class Fodnoter:
    """Noterne nederst på én slide: generelle noter og nummererede fodnoter.

    En note om en bestemt kolonne, række, kasse eller et tal får et fodnotemærke dér
    (maerke() giver »¹«, og noten står som »¹ …«); en note om hele sliden (kilde, kurser,
    datoer) er generel og står uden nummer først. Numrene følger den rækkefølge, mærkerne
    sættes i, og samme tekst får samme nummer.
    """

    def __init__(self, kilde=None):
        self.generelle = []
        self.numre = {}
        self.generel(kilde)

    def maerke(self, tekst):
        """Mærket til tekst (»¹«), eller "" uden tekst."""
        if not tekst:
            return ""
        return str(self.numre.setdefault(tekst, len(self.numre) + 1)).translate(SUPERSKRIFT)

    def maerker(self, tekster):
        """Mærkerne til flere noter efter hinanden (»¹²«)."""
        return "".join(self.maerke(t) for t in tekster)

    def generel(self, *tekster):
        for t in tekster:
            if isinstance(t, (list, tuple)):
                self.generel(*t)
            elif t and t not in self.generelle:
                self.generelle.append(t)

    def tekst(self):
        nummererede = [f"{str(nr).translate(SUPERSKRIFT)} {t}"
                       for t, nr in sorted(self.numre.items(), key=lambda x: x[1])]
        return " ".join(self.generelle + nummererede)

    def tegn(self, s, y=Y_NOTER):
        """Noterne ved y, eller højere, hvis de ellers løber ned i logoet."""
        tekst = self.tekst()
        if tekst:
            _tekst(s, X0, _noter_y(tekst, y), X1 - X0, 0.5, tekst, 7, kursiv=True)


KORT_FAST = 1.84          # alt på et virksomhedskort undtagen beskrivelsen (tommer)
KORT_LINJE = 0.135        # linjehøjde for beskrivelsen ved 7,5 pt
KORT_TEGN = 42            # tegn pr. linje i et kort på en fjerdedel af indholdsbredden
KORT_KOL = 4              # kort pr. række; to rækker giver otte selskaber pr. slide
KORT_LUFT = 0.16


def _andel(v):
    """Ejerandel: 58 % frem for 58,0 %, men 20,8 % hvor decimalen betyder noget."""
    return pct(v, 0) if abs(v * 100 - round(v * 100)) < 0.05 else pct(v)


def _kort_hoejde(beskrivelser):
    """Korthøjde, så den længste beskrivelse har plads over søjlerne."""
    linjer = max([ceil(len(b or "") / KORT_TEGN) for b in beskrivelser] or [0])
    return KORT_FAST + max(0.22, KORT_LINJE * linjer)


def _profilkort(s, x, y, w, h, r, info, valuta, ejer_m=""):
    """Ét kort: navn, industri (GICS-sektor), closing, søjler for investering og dagsværdi, og valgfri tekst.

    Søjlerne har deres egen skala pr. kort. Fælles skala ville gøre alle andre kort ulæselige,
    når ét selskab er mange gange større end resten; TVPI under kortet giver forholdet. ejer_m
    er fodnotemærket efter ejerandelen.
    """
    _bjaelke(s, x, y, w, h, CREAM)
    indryk = 0.10
    tw = w - 2 * indryk
    _tekst(s, x + indryk, y + 0.08, tw, 0.20, r["navn"], 10, fed=True, farve=SAGE)
    linje2 = [t for t in (r.get("industri"), f"Closing {r['closing']:%m.%Y}" if r["closing"]
                          else (str(r["vintage"]) if r["vintage"] else None)) if t]
    _tekst(s, x + indryk, y + 0.28, tw, 0.30, " · ".join(linje2) or IKKE_OPLYST, 7.5)

    ejer = []
    if info.get("ejerandel") is not None:
        ejer.append(f"Ejerandel {_andel(info['ejerandel'])}")
    if info.get("stemmer") is not None:
        ejer.append(f"stemmer {_andel(info['stemmer'])}")
    if ejer:
        _tekst(s, x + indryk, y + 0.52, tw, 0.16, ", ".join(ejer) + ejer_m, 7.5, fed=True)

    if info.get("beskrivelse"):
        _tekst(s, x + indryk, y + 0.70, tw, h - KORT_FAST, info["beskrivelse"], 7.5)

    graf_h = 0.62
    gy = y + h - graf_h - 0.36
    _streg(s, x + indryk, gy - 0.16, x + w - indryk, gy - 0.16, farve=SAGE_LYS)
    maks = max(r["kostpris_fond"] or 0, r["dagsvaerdi_fond"] or 0) or 1
    _soejler(s, x + indryk + 0.20, gy, tw - 0.40, graf_h,
             [("Investering", r["kostpris_fond"], TAUPE),
              ("Dagsværdi", r["dagsvaerdi_fond"], SAGE)], maks, fmt=million)
    _tekst(s, x + indryk, y + h - 0.22, tw, 0.18, f"TVPI {gange(r['tvpi'])}", 8.5, fed=True,
           farve=SAGE, just=C)


def virksomhedsprofiler(prs, f, investor_navn, maks_pr_slide=KORT_KOL * 2):
    """Et kort pr. porteføljeselskab. Ingen slide for fonde, der ejer andre fonde.

    Beløbene er hele fondens, som GP'en rapporterer dem - ikke investors andel.
    """
    beh = f["beholdninger"]
    if not beh or beh["type"] != "selskaber" or beh.get("via"):
        return []
    raekker = [r for r in beh["raekker"] if r["dagsvaerdi_fond"] or r["kostpris_fond"]]
    if not raekker:
        return []
    info = f["virksomheder"]

    lavede = []
    sider = [raekker[i:i + maks_pr_slide] for i in range(0, len(raekker), maks_pr_slide)]
    for s_nr, side in enumerate(sider, 1):
        s = _ny_slide(prs, L_TEKST)
        suffiks = f" ({s_nr}:{len(sider)})" if len(sider) > 1 else ""
        _titel(s, f"Porteføljeselskaber{suffiks}")
        _fjern_tomme_ph(s)
        _logo(s, f["logo"])
        fod = Fodnoter("Kilde: Forvalterrapporter og egne beregninger.")
        beloeb_m = fod.maerke(
            "Beløbene er ikke investors andel. Hvert kort har sin egen søjleskala, så "
            "søjlehøjder kan ikke sammenlignes på tværs af kortene - det er TVPI, der viser "
            "forholdet.")
        ejer_m = fod.maerke("Ejerandel og stemmeandel er oplyst af forvalteren." if any(
            (info.get(r["navn"]) or {}).get(k) is not None for r in side
            for k in ("ejerandel", "stemmer")) else None)
        _tekst(s, X0, Y_UNDERTITEL, X1 - X0, 0.22,
               f"{f['fund_name']}  ·  {investor_navn}  ·  pr. {beh['dato']:%d.%m.%Y}  ·  "
               f"hele fonden, beløb i mio. {f['valuta']}{beloeb_m}", 9.5)

        kort_w = (X1 - X0 - KORT_LUFT * (KORT_KOL - 1)) / KORT_KOL
        n_raekker = ceil(len(side) / KORT_KOL)
        # højden følger den længste beskrivelse, men må ikke sprænge fladen over noterne
        plads = (Y_NOTER - Y_TOP - 0.15 - KORT_LUFT * (n_raekker - 1)) / n_raekker
        kort_h = min(plads,
                     _kort_hoejde([(info.get(r["navn"]) or {}).get("beskrivelse") for r in side]))
        for i, r in enumerate(side):
            _profilkort(s, X0 + (i % KORT_KOL) * (kort_w + KORT_LUFT),
                        Y_TOP + (i // KORT_KOL) * (kort_h + KORT_LUFT), kort_w, kort_h, r,
                        info.get(r["navn"]) or {}, f["valuta"], ejer_m)
        fod.tegn(s)
        lavede.append(s)
    return lavede


# ---------- oversigtssider (ATP-rapportens "Portfolio Reporting", side 5-7) ----------

_KALDT_LABEL = {False: "Kaldt", True: "Kaldt, netto"}

DAGGER = "†"        # markerer en fond, der ikke er opgjort pr. rapportdatoen


def _foraeldede(r):
    """{fondsnavn: opgørelsesdato} for de fonde, der halter efter rapportdatoen."""
    return {x["fond"]: x["dato"] for x in r["datotjek"]["foraeldede"]}


def _datomarkering(r, fund_name):
    """Dagger efter fondsnavnet i tabellerne, så afvigelsen ses i selve rækken."""
    return f" {DAGGER}" if fund_name in _foraeldede(r) else ""


def _datonoter(r):
    """Noter om uens opgørelsesdatoer. Tom liste, når alle fonde er opgjort samme dag."""
    tjek = r["datotjek"]
    if tjek["ens"]:
        return []
    noter = []
    if tjek["foraeldede"]:
        fonde = ", ".join(f"{x['fond']} pr. {x['dato']:%d.%m.%Y}" for x in tjek["foraeldede"])
        noter.append(f"{DAGGER} Ikke alle fonde er opgjort pr. "
                     f"{tjek['rapport_dato']:%d.%m.%Y}: {fonde}. Totaler og fordelinger "
                     "blander derfor opgørelsesdatoer.")
    for x in tjek["skaeve_beholdninger"]:
        noter.append(f"{x['fond']}: beholdningerne er opgjort pr. "
                     f"{x['beholdninger_dato']:%d.%m.%Y}, mens nøgletallene er pr. "
                     f"{x['nav_dato']:%d.%m.%Y}.")
    return noter


def _oversigtsslide(prs, r, titel, enhed=None):
    s = _ny_slide(prs, L_TEKST)
    _titel(s, titel)
    _fjern_tomme_ph(s)
    if enhed is None:
        enhed = f"beløb i t{r['valuta']}" if r["valuta"] else "beløb i tusinde, fondens valuta"
    _tekst(s, X0, Y_UNDERTITEL, X1 - X0, 0.22,
           f"{r['navn']}  ·  pr. {r['as_of']:%d.%m.%Y}  ·  {enhed}", 9.5)
    return s


def _raekke_h(antal, maks=0.30, mindst=0.19):
    """Rækkehøjde, så tabellen fylder fladen mellem undertitel og noter uden at løbe over."""
    return max(mindst, min(maks, (Y_NOTER - Y_TOP - 0.15) / antal))


# Kolonnerne i Lange tilsagn (tommer); summen er indholdsbredden X1 - X0
LT_KOL = [("Fond", 2.10, V), ("Forvalter", 1.20, C), ("Type", 0.95, C), ("Vintage", 0.62, C),
          ("Løbetid", 0.72, C), ("Valuta", 0.45, C), ("Fra tilsagn til resttilsagn", 2.82, C),
          ("TVPI", 0.76, C), ("IRR", 0.76, C)]
LT_BELOEB, LT_TVPI, LT_IRR = 6, 7, 8        # kolonneindeks for søjlerne


def _lt_x(i, kol=LT_KOL):
    """Venstrekant af kolonne i."""
    return X0 + sum(b for _, b, _ in kol[:i])


def _kolonnehoved(s, kol, hoved_h):
    """Kolonneoverskrifter med en streg under hver."""
    for i, (navn, b, just) in enumerate(kol):
        x = _lt_x(i, kol)
        _tekst(s, x, Y_TOP, b, hoved_h - 0.06, navn, 9, fed=True, farve=SAGE, just=just,
               anker=MSO_ANCHOR.BOTTOM)
        # stregen skal have samme midte som overskriften over den: en venstrestillet kolonne
        # starter ved tekstens kant, en centreret trækkes lige meget ind i begge sider
        ind_v, ind_h = (0.0, 0.10) if just == V else (0.05, 0.05)
        y_streg = Y_TOP + hoved_h - 0.02
        _streg(s, x + ind_v, y_streg, x + b - ind_h, y_streg)


def _fondscelle(s, r, f, kol, i, y, raekke_h, maerke=""):
    """Rækkens faste kolonner: fond, forvalterlogo og de af type, vintage, løbetid og valuta,
    som kol har (fundet ved kolonnenavnet). maerke er fodnotemærker efter fondsnavnet."""
    if i % 2 == 1:
        _bjaelke(s, X0 - 0.05, y, X1 - X0 + 0.1, raekke_h, CREAM)
    if i:
        _streg(s, X0, y, X1, y, farve=SAGE_LYS, pt=0.5)
    midt = y + raekke_h / 2
    if f["benchmark"]:
        _tekst(s, _lt_x(0, kol), y + 0.06, kol[0][1], 0.14, f"Benchmark: {f['benchmark']}", 6.5,
               farve=SAGE, kursiv=True)
    afvigende = _foraeldede(r).get(f["fund_name"])
    _tekst(s, _lt_x(0, kol), y + 0.20, kol[0][1] - 0.1,
           raekke_h - (0.46 if afvigende else 0.26),
           f["fund_name"] + _datomarkering(r, f["fund_name"]) + maerke, 9,
           anker=MSO_ANCHOR.MIDDLE)
    if afvigende:
        _tekst(s, _lt_x(0, kol), y + raekke_h - 0.26, kol[0][1], 0.16,
               f"opgjort pr. {afvigende:%d.%m.%Y}", 6.5, farve=TERRA)
    if f["logo"]:
        _logo(s, f["logo"], x_midt=_lt_x(1, kol) + kol[1][1] / 2, y_midt=midt,
              maks_b=min(1.05, kol[1][1] - 0.15), maks_h=min(0.34, raekke_h - 0.2))
    elif f.get("manager"):
        # uden logo (fx den anonymiserede udgave) står forvalterens navn i kolonnen
        _tekst(s, _lt_x(1, kol), y, kol[1][1], raekke_h, f["manager"], 8.5, just=C,
               anker=MSO_ANCHOR.MIDDLE)
    felter = {"Type": f["segment"], "Vintage": f["vintage"], "Løbetid": f["loebetid"],
              "Valuta": f["valuta"]}
    for j, (navn, _, _) in enumerate(kol):
        if navn not in felter:
            continue
        vaerdi = felter[navn]
        _tekst(s, _lt_x(j, kol), y, kol[j][1], raekke_h, vaerdi or IKKE_OPLYST, 8.5, just=C,
               anker=MSO_ANCHOR.MIDDLE)


def _tvpi_irr_soejler(s, kol, j_tvpi, y, raekke_h, n, maks_tvpi, maks_irr, grundlinje=False):
    """TVPI- og IRR-søjlen i kolonne j_tvpi og den næste.

    Med grundlinje står søjlerne på en linje tæt på rækkens bund og bruger næsten hele
    rækkehøjden; ellers får de den smalle bånd midt i rækken, som Lange tilsagn bruger.
    """
    top, h = (y + 0.10, raekke_h - 0.18) if grundlinje else (y + 0.24, raekke_h - 0.48)
    _noegletalsbjaelke(s, _lt_x(j_tvpi, kol) + 0.14, top, kol[j_tvpi][1] - 0.28,
                       h, n["tvpi"], maks_tvpi, gange(n["tvpi"]), SAGE_MID, grundlinje)
    j_irr = j_tvpi + 1
    _noegletalsbjaelke(s, _lt_x(j_irr, kol) + 0.14, top, kol[j_irr][1] - 0.28,
                       h, n["irr"], maks_irr, pct(n["irr"]),
                       OLIVEN if (n["irr"] or 0) >= 0 else ROED, grundlinje)


def kaldt_netto(n):
    """Kaldt kapital fratrukket genkaldelige udlodninger.

    Resttilsagnet er tilsagn - kaldt + genkaldelige udlodninger, hvor kaldt er forvalterens
    kaldte kapital (inkl. rabat på kald, uden udligningsrente). Trækkes de genkaldelige
    udlodninger fra den kaldte kapital i stedet, går tilsagn - kaldt op i resttilsagnet, og
    broen kan tegnes med to trin. Tallet er derfor lavere end "Indbetalt kapital" i tabellerne.
    """
    kaldt = n.get("kaldt")
    return (n["indbetalt"] if kaldt is None else kaldt or 0) - (n.get("recallable") or 0)


def _resttilsagnsnoter(f):
    """Noten til en fonds resttilsagn i tabellerne på tværs af fonde: de genkaldelige
    udlodninger, med fondens beløb. Hvad begrebet betyder, står i Definitioner. En rabat på kald
    har sin note ved den indbetalte kapital (data.rabat), da den indgår i den."""
    n, v = f["n"], f["valuta"]
    return ([f"Resttilsagnet indeholder genkaldelige udlodninger på "
             f"{tusind(n['recallable'])} t{v}."] if (n.get("recallable") or 0) else [])


def _rabatnote(fonde):
    """Note til tilsagnsslides, når en forvalter har givet rabat på kald. None ellers."""
    med = [f["fund_name"] for f in fonde if (f["n"].get("rabat") or 0) > 0]
    if not med:
        return None
    return ("Forvalteren har givet rabat på kapitalkald for " + ", ".join(med)
            + "; rabatten indgår i indbetalt kapital og som udlodning.")


def _tilsagnsbro(n, netto=None):
    """Vandfaldet tilsagn - kaldt = resttilsagn, som altid går op.

    netto styrer kun labelen, så alle rækker på en slide skriver det samme. Værdien nettes
    altid, hvilket ikke ændrer noget for en fond uden genkaldelige udlodninger.
    """
    tilsagn, kaldt = n["tilsagn"] or 0, kaldt_netto(n)
    rest = n["resttilsagn"] or 0
    netto = bool(n.get("recallable")) if netto is None else netto
    return [("Tilsagn", tilsagn, 0.0, tilsagn, SAGE),
            (_KALDT_LABEL[netto], kaldt, tilsagn - kaldt, tilsagn - kaldt, TERRA),
            ("Resttilsagn", rest, 0.0, rest, TAUPE)]


def lange_tilsagn(prs, r):
    """Visuel porteføljeoversigt: én række pr. fond med forvalterlogo og tegnede søjler."""
    s = _oversigtsslide(prs, r, "Lange tilsagn")
    fonde, t = r["fonde"], r["total"]
    n_raekker = len(fonde) + bool(t)
    hoved_h = 0.30
    raekke_h = min(0.95, (Y_NOTER - Y_TOP - hoved_h - 0.1) / max(n_raekker, 1))

    _kolonnehoved(s, LT_KOL, hoved_h)

    # labelen skal være ens på hele sliden, også på de fonde der intet har genkaldt
    nettet = any((f["n"].get("recallable") or 0) for f in fonde)

    # fælles beløbsskala kræver fælles valuta; ellers skaleres hver række for sig, så
    # søjlelængder ikke inviterer til at sammenligne DKK med EUR og USD
    samme_valuta = r["valuta"] is not None
    maks_belob = max([f["n"]["tilsagn"] or 0 for f in fonde] + [1]) if samme_valuta else None
    maks_tvpi = max([f["n"]["tvpi"] or 0 for f in fonde] + [1])
    maks_irr = max([abs(f["n"]["irr"] or 0) for f in fonde] + [0.01])

    y = Y_TOP + hoved_h
    for i, f in enumerate(fonde):
        n = f["n"]
        _fondscelle(s, r, f, LT_KOL, i, y, raekke_h)
        _vandfald(s, _lt_x(LT_BELOEB) + 0.14, y + 0.24, LT_KOL[LT_BELOEB][1] - 0.28,
                  raekke_h - 0.48, _tilsagnsbro(n, nettet), maks_belob or n["tilsagn"] or 1)
        _tvpi_irr_soejler(s, LT_KOL, LT_TVPI, y, raekke_h, n, maks_tvpi, maks_irr)
        y += raekke_h

    if t:
        _bjaelke(s, X0 - 0.05, y, X1 - X0 + 0.1, raekke_h, SAGE_LYS)
        _tekst(s, _lt_x(0), y, LT_KOL[0][1], raekke_h, "I alt", 9, fed=True,
               anker=MSO_ANCHOR.MIDDLE)
        _tekst(s, _lt_x(5), y, LT_KOL[5][1], raekke_h, r["valuta"] or "", 8.5, fed=True,
               just=C, anker=MSO_ANCHOR.MIDDLE)
        _vandfald(s, _lt_x(LT_BELOEB) + 0.14, y + 0.24, LT_KOL[LT_BELOEB][1] - 0.28,
                  raekke_h - 0.48, _tilsagnsbro(t, nettet), t["tilsagn"] or 1)
        for j, tekst_ in ((LT_TVPI, gange(t["tvpi"])), (LT_IRR, pct(t["irr"]))):
            _tekst(s, _lt_x(j), y, LT_KOL[j][1], raekke_h, tekst_, 8.5, fed=True, just=C,
                   anker=MSO_ANCHOR.MIDDLE)

    noter = ["Kilde: Forvalterrapporter og egne beregninger.",
             "Søjlerne går fra tilsagnet over den kaldte kapital til resttilsagnet."]
    if samme_valuta:
        noter.append("Alle fonde har samme skala"
                     + ("; i alt-rækken har sin egen." if t else "."))
    else:
        noter.append("Fondene har forskellige valutaer, så hver fond er skaleret for sig: "
                     "søjlelængder kan sammenlignes inden for en række, ikke på tværs af dem. "
                     "Der vises ingen samlet række.")
    if nettet:
        noter.append("Genkaldelige udlodninger er trukket fra den kaldte kapital, så søjlerne "
                     "går op i resttilsagnet. Kaldt, netto er derfor lavere end den indbetalte "
                     "kapital i porteføljeoversigten, som er den kapital, der faktisk er "
                     "betalt ind.")
    if _rabatnote(fonde):
        noter.append(_rabatnote(fonde))
    noter += _datonoter(r)
    fod = Fodnoter()
    fod.generel(noter)
    fod.tegn(s)
    return s


# Kolonnerne i Investeringer (tilsagn på tværs af fonde): Lange tilsagn med tal i stedet for vandfaldet
# (valutaen står på linjen over kolonnenavnet, som på fordelingssiderne)
TPT_KOL = [("Fond", 2.57, V), ("Forvalter", 1.20, C), ("Type", 0.95, C), ("Vintage", 0.58, C),
           ("Valuta", 0.45, C), ("Tilsagn", 0.73, C), ("PICC", 0.58, C),
           ("DKK\nResttilsagn", 0.88, C), ("DKK\nNAV", 0.88, C), ("TVPI", 0.78, C),
           ("IRR", 0.78, C)]
TPT_VALUTA, TPT_TILSAGN, TPT_PICC, TPT_REST_DKK, TPT_NAV_DKK, TPT_TVPI = 4, 5, 6, 7, 8, 9


def _kursnote(r):
    """Kurserne for de fremmede valutaer. Kilden (Danmarks Nationalbank) og at beløbene er i
    DKK, står allerede på sliden."""
    kurser = {v: k for v, k in (r.get("kurser") or {}).items() if v != "DKK" and k}
    if not kurser:
        return None
    dele = []
    for v, (kurs, dato) in sorted(kurser.items()):
        dele.append(f"{v} {style._dk(kurs, 4)}"
                    + (f" pr. {dato:%d.%m.%Y}" if dato != r["as_of"] else ""))
    return f"Kurser pr. {r['as_of']:%d.%m.%Y} (DKK pr. enhed): {'; '.join(dele)}."


def tilsagn_paa_tvaers(prs, r):
    """Som Lange tilsagn, men med tilsagn, PICC, resttilsagn i DKK og NAV i DKK som tal."""
    s = _oversigtsslide(prs, r, "Investeringer",
                        enhed="beløb i tusinde; tilsagn i fondens valuta, "
                              "resttilsagn og NAV i DKK")
    fonde, t, td = r["fonde"], r["total"], r.get("total_dkk")
    n_raekker = len(fonde) + bool(t or td)
    hoved_h = 0.44                                   # to linjer: valuta over kolonnenavn
    fod = Fodnoter("Kilde: Forvalterrapporter, Danmarks Nationalbank og egne beregninger.")
    hoved_m = {}
    kurs_m = fod.maerke(_kursnote(r))
    hoved_m |= {TPT_REST_DKK: kurs_m, TPT_NAV_DKK: kurs_m}
    # rækkernes mærker sættes, før noternes længde bestemmer rækkehøjden
    fond_m = {f["fund_name"]: (fod.maerker(_fondsnoter(f, "indbetalt")),
                               fod.maerker(_resttilsagnsnoter(f))) for f in fonde}
    i_alt_m = "" if t or not td else fod.maerke(
        "Fondene har forskellige valutaer, så i alt-rækken har kun beløbene i DKK samt TVPI og "
        "IRR i DKK.")
    fod.generel(_datonoter(r))
    raekke_h = min(0.95, (_noter_y(fod.tekst()) - Y_TOP - hoved_h - 0.1) / max(n_raekker, 1))
    _kolonnehoved(s, [(navn + hoved_m.get(j, ""), b, just)
                      for j, (navn, b, just) in enumerate(TPT_KOL)], hoved_h)

    maks_tvpi = max([f["n"]["tvpi"] or 0 for f in fonde] + [1])
    maks_irr = max([abs(f["n"]["irr"] or 0) for f in fonde] + [0.01])

    def tal(j, y, tekst, fed=False):
        _tekst(s, _lt_x(j, TPT_KOL), y, TPT_KOL[j][1], raekke_h, tekst, 8.5, fed=fed, just=C,
               anker=MSO_ANCHOR.MIDDLE)

    y = Y_TOP + hoved_h
    for i, f in enumerate(fonde):
        n = f["n"]
        _fondscelle(s, r, f, TPT_KOL, i, y, raekke_h, fond_m[f["fund_name"]][0])
        for j, tekst in ((TPT_TILSAGN, tusind(n["tilsagn"])), (TPT_PICC, pct(n.get("picc"))),
                         (TPT_REST_DKK, tusind(n.get("resttilsagn_dkk"))
                          + fond_m[f["fund_name"]][1]),
                         (TPT_NAV_DKK, tusind(n.get("nav_dkk")))):
            tal(j, y, tekst)
        _tvpi_irr_soejler(s, TPT_KOL, TPT_TVPI, y, raekke_h, n, maks_tvpi, maks_irr,
                          grundlinje=True)
        y += raekke_h

    if t or td:
        # DKK-beløbene kan altid lægges sammen; resten kun når fondene har samme valuta
        _bjaelke(s, X0 - 0.05, y, X1 - X0 + 0.1, raekke_h, SAGE_LYS)
        _tekst(s, _lt_x(0, TPT_KOL), y, TPT_KOL[0][1], raekke_h, "I alt" + i_alt_m, 9,
               fed=True, anker=MSO_ANCHOR.MIDDLE)
        celler = [(TPT_VALUTA, r["valuta"] or "")]
        if t:
            celler += [(TPT_TILSAGN, tusind(t["tilsagn"])), (TPT_PICC, pct(t["picc"])),
                       (TPT_TVPI, gange(t["tvpi"])), (TPT_TVPI + 1, pct(t["irr"]))]
        if td:
            celler += [(TPT_REST_DKK, tusind(td["resttilsagn_dkk"])),
                       (TPT_NAV_DKK, tusind(td["nav_dkk"]))]
            if not t and td["indbetalt_dkk"]:
                # som på fordelingssiderne: TVPI i DKK kan beregnes på tværs af valutaer
                celler.append((TPT_TVPI, gange((td["udloddet_dkk"] + td["nav_dkk"])
                                               / td["indbetalt_dkk"])))
            if not t:
                celler.append((TPT_TVPI + 1, pct((r.get("irr_dkk") or {}).get("irr"))))
        for j, tekst in celler:
            tal(j, y, tekst, fed=True)

    fod.tegn(s)
    return s


def portefoeljeoversigt(prs, r):
    """Én række pr. fond med nøgletal og en samlet række for porteføljen."""
    s = _oversigtsslide(prs, r, "Porteføljeoversigt")
    t = r["total"]
    fod = Fodnoter("Kilde: Forvalterrapporter og egne beregninger.")
    irr_m = ""
    if not t:
        fod.generel("Fondene har forskellige valutaer, så der vises ingen samlet række.")
    raekker = [["Fond", "Strategi", "Vintage", "Løbetid", "Valuta", "Tilsagn", "Indbetalt",
                "PICC", "Udloddet", "NAV", "DPI", "TVPI", f"IRR{irr_m}"]]
    for f in r["fonde"]:
        n = f["n"]
        indbetalt_m = fod.maerker(_fondsnoter(f, "indbetalt"))
        raekker.append([f["fund_name"] + _datomarkering(r, f["fund_name"]),
                        f["segment"] or IKKE_OPLYST, f["vintage"] or IKKE_OPLYST,
                        f["loebetid"] or IKKE_OPLYST, f["valuta"], tusind(n["tilsagn"]),
                        tusind(n["indbetalt"]) + indbetalt_m, pct(n["picc"]),
                        tusind(n["udloddet"]), tusind(n["nav"]), gange(n["dpi"]),
                        gange(n["tvpi"]), pct(n["irr"])])
    if t:
        raekker.append(["I alt", "", "", "", r["valuta"], tusind(t["tilsagn"]),
                        tusind(t["indbetalt"]), pct(t["picc"]), tusind(t["udloddet"]),
                        tusind(t["nav"]), gange(t["dpi"]), gange(t["tvpi"]), pct(t["irr"])])
    _tabel(s, X0, Y_TOP, X1 - X0, raekker,
           bredder=[2.65, 1.6, 0.7, 0.95, 0.65, 0.88, 0.88, 0.75, 0.88, 0.88, 0.65, 0.65, 0.7],
           just=[V, V, C, C, C, H, H, H, H, H, H, H, H], raekke_h=_raekke_h(len(raekker)), pt=9,
           hoved=True, total=bool(t), zebra=True)
    fod.generel(_datonoter(r))
    fod.tegn(s)
    return s


def _romertal(o):
    return bool(o) and set(o) <= set("IVX")


def _kort_fondsnavn(navn):
    """Kort navn til smalle kolonner.

    "K Alternativ Private Equity 2020 K/S" -> "KAPE 2020", "Petersen & Partners Private
    Credit II A/S" -> "PPPC II"; ellers de første to ord. Et navn, der slutter med et
    løbenummer (den anonymiserede udgaves "Private Equity-fond 1"), står helt.
    """
    uden = _uden_selskabsform(navn)
    ord_ = uden.split()
    if len(ord_) <= 2 or (ord_[-1].isdigit() and len(ord_[-1]) < 4):
        return uden
    aar = [o for o in ord_ if o.isdigit() and len(o) == 4]
    if aar:
        forbogstaver = "".join(o[0] for o in ord_ if o[0].isupper() and not o.isdigit())
        return f"{forbogstaver} {aar[0]}"
    nummer = next((i for i, o in enumerate(ord_) if i and _romertal(o)), None)
    if nummer:
        return "".join(o[0] for o in ord_[:nummer] if o[0].isupper()) + " " + ord_[nummer]
    return " ".join(ord_[:2])


def _fakta_flise(slide, x, y, w, h, tal, label):
    """Lille nøgletalsflise: stort tal over en label, på cremet bund."""
    _bjaelke(slide, x, y, w, h, CREAM)
    # tallet står på midten og labelen hænger under den, så de to tilsammen er centreret
    _tekst(slide, x, y, w, h * 0.58, str(tal), 16, fed=True, farve=SAGE, just=C,
           anker=MSO_ANCHOR.BOTTOM)
    _tekst(slide, x, y + h * 0.60, w, h * 0.35, label, 8, just=C)


POSITIONSTYPE = {"porteføljeselskab": "Porteføljeselskab",
                  "underliggende fond": "Underliggende fond", "fond": "Fond (ikke gennemlyst)"}


def stoerste_positioner(prs, r):
    """Porteføljens største positioner i DKK på tværs af fondene, med andel af investors NAV.

    Data fra data.stoerste_positioner(): selskaber og underliggende fonde som investors andel
    af dagsværdien, og fonde uden sammenlignelige beholdninger som én linje med deres NAV.
    None uden DKK-tal.
    """
    sp = r.get("stoerste_positioner")
    if not sp or not sp["linjer"]:
        return None
    s = _oversigtsslide(prs, r, "Største positioner", enhed="investors andel, beløb i tusinde DKK")

    def via(p):
        if len(p["fonde"]) > 2:
            return f"{len(p['fonde'])} fonde"
        return ", ".join(_kort_fondsnavn(f) + _datomarkering(r, f) for f in p["fonde"])

    fod = Fodnoter("Kilde: Forvalterrapporter, Danmarks Nationalbank og egne beregninger.")
    pos_m = fod.maerke("Positionerne er fra forvalternes beholdningsoversigter; samme position "
                       "i flere fonde er lagt sammen.")
    andel_m = fod.maerke("Beholdningerne summerer ikke til NAV, da fondene også har likvider, "
                         "hensat carried interest og andre poster.")
    fond_m = fod.maerke("Fonden har ingen beholdninger, der kan gennemlyses på samme måde, og "
                        "indgår med sin NAV som én position (lys bjælke)."
                        if sp["fondslinjer"] else None)
    linjer = sp["linjer"]
    raekker = [["", f"Position{pos_m}", "Type", "Via fond", "tDKK", f"Andel af NAV{andel_m}", ""]]
    raekker += [[str(i), p["navn"], POSITIONSTYPE[p["art"]] + (fond_m if p["art"] == "fond"
                                                               else ""),
                 "–" if p["art"] == "fond" else via(p), tusind(p["dkk"]), pct(p["andel"]), ""]
                for i, p in enumerate(linjer, 1)]
    raekker.append(["", f"De {len(linjer)} største i alt", "", "", tusind(sp["top_dkk"]),
                    pct(sp["top_andel"]), ""])
    bredder = [0.3, 3.05, 1.45, 2.6, 0.9, 0.95, 1.13]
    raekke_h = _raekke_h(len(raekker), maks=0.27)
    _tabel(s, X0, Y_TOP, X1 - X0, raekker, bredder=bredder, just=[H, V, V, V, H, H, V],
           raekke_h=raekke_h, pt=8.5, hoved=True, total=True, zebra=True)
    # bjælkerne i sidste kolonne, skaleret til den største position
    bx = X0 + (X1 - X0) * sum(bredder[:-1]) / sum(bredder) + 0.08
    bw = (X1 - X0) * bredder[-1] / sum(bredder) - 0.16
    maks = max(p["dkk"] for p in linjer) or 1
    for i, p in enumerate(linjer, 1):
        _bjaelke(s, bx, Y_TOP + i * raekke_h + raekke_h * 0.3, bw * max(p["dkk"], 0) / maks,
                 raekke_h * 0.4, SAGE_LYS if p["art"] == "fond" else SAGE)

    _afslut_noter(s, r, fod, Y_NOTER)
    return s


def gennemlysning(prs, r):
    """Hvad porteføljen reelt ejer: fondenes beholdninger samlet og rangeret efter dagsværdi."""
    g = r["gennemlysning"]
    if not g:
        return None
    s = _oversigtsslide(prs, r, "Gennemlysning af porteføljen")

    flise_w, flise_h = 1.70, 0.62
    for i, (tal, label) in enumerate([(g["antal_fonde"], "fonde i porteføljen"),
                                      (g["antal_underfonde"], "underliggende fonde"),
                                      (g["antal_selskaber"], "porteføljeselskaber")]):
        _fakta_flise(s, X0 + i * (flise_w + 0.14), Y_TOP, flise_w, flise_h, tal, label)

    top = g["linjer"][:10]
    oevrige = g["linjer"][10:]
    y = Y_TOP + flise_h + 0.30
    fod = Fodnoter("Kilde: Forvalterrapporter og egne beregninger. Fondenes beholdninger pr. "
                   f"{g['dato']:%d.%m.%Y} ganget med ejerandelen.")
    sum_m = fod.maerke(
        f"Summen af beholdningerne ({tusind(g['dagsvaerdi'])} t{r['valuta']}) afviger fra "
        f"fondenes samlede NAV ({tusind(g['nav'])} t{r['valuta']}), fordi likvider, "
        "omkostninger og carried interest på fondsniveau ikke indgår i forvalternes "
        "beholdningsoversigter.")
    _tekst(s, X0, y, X1 - X0, 0.22,
           f"Største eksponeringer – {r['navn']}s andel (t{r['valuta']}){sum_m}", 10.5,
           fed=True, farve=SAGE)
    y += 0.28

    # søjlediagram til venstre som rigtigt diagram, så tallene kan redigeres i PowerPoint
    graf_w = 5.55
    data = CategoryChartData(number_format="#,##0")
    data.categories = [x["navn"] for x in reversed(top)]       # største øverst i en bar-graf
    data.add_series("Dagsværdi", [_kt(x["dagsvaerdi"]) for x in reversed(top)])
    diagram = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(X0 - 0.08), Inches(y - 0.06),
                                 Inches(graf_w), Inches(Y_NOTER - y - 0.04), data).chart
    diagram.font.name, diagram.font.size = BROED_FONT, Pt(8)
    diagram.font.color.rgb = SORT
    diagram.has_legend = False
    diagram.has_title = False
    serie = diagram.plots[0].series[0]
    serie.format.fill.solid()
    serie.format.fill.fore_color.rgb = SAGE
    diagram.plots[0].gap_width = 60
    diagram.plots[0].has_data_labels = True
    etiketter = diagram.plots[0].data_labels
    etiketter.number_format, etiketter.number_format_is_linked = "#,##0", False
    etiketter.font.size, etiketter.font.name = Pt(7.5), BROED_FONT
    x_akse, y_akse = diagram.category_axis, diagram.value_axis
    x_akse.format.line.color.rgb = STREG
    x_akse.has_major_gridlines = False
    y_akse.has_major_gridlines = False
    y_akse.visible = False
    y_akse.format.line.fill.background()

    # tabellen til højre gentager de samme ti linjer med kostpris og TVPI
    tab_x = X0 + graf_w + 0.16
    via_m = fod.maerke("Linjerne blander underliggende fonde og porteføljeselskaber; kolonnen "
                       "»Via« viser hvilken fond eksponeringen kommer fra.")
    raekker = [["Eksponering", f"Via{via_m}", "Kostpris", "Dagsværdi", "TVPI"]]
    raekker += [[x["navn"], _kort_fondsnavn(x["fond"]), tusind(x["kostpris"]),
                 tusind(x["dagsvaerdi"]), gange(x["tvpi"])] for x in top]
    if oevrige:
        raekker.append([f"Øvrige ({len(oevrige)})", "", tusind(sum(x["kostpris"] or 0 for x in oevrige)),
                        tusind(sum(x["dagsvaerdi"] for x in oevrige)), ""])
    raekke_h = min(0.26, (Y_NOTER - y - 0.08) / len(raekker))
    _tabel(s, tab_x, y, X1 - tab_x, raekker, bredder=[1.85, 1.15, 0.85, 0.95, 0.6],
           just=[V, V, H, H, H], raekke_h=raekke_h, pt=7.5, hoved=True, total=bool(oevrige),
           zebra=True)

    if g["uden_beholdninger"]:
        fod.generel("Forvalteren oplyser ikke beholdninger for: "
                    + ", ".join(g["uden_beholdninger"]) + ".")
    fod.generel(_datonoter(r))
    fod.tegn(s)
    return s


FORD_MIN_H = 0.18     # laveste rækkehøjde på fordelingssiderne, før tabellen deles (tommer)


def _del_paa_sider(blokke, plads):
    """Fordel blokke af rækker på sider med plads til højst plads rækker hver.

    En blok er [(celler, art), ...] med art "total", "gruppe" eller "fond"; en gruppeblok
    starter med sin overskrift og slutter med sin i alt-række. Siderne fyldes op: en gruppe,
    der ikke kan nå at stå færdig, deles, når der er plads til overskriften og to fonde, og
    mindst én fond går med videre, så i alt-rækken aldrig står alene. Fortsættelsen får
    overskriften igen med "(fortsat)", så det altid kan ses, hvilken gruppe rækken hører til.
    """
    sider, side = [], []
    for blok in blokke:
        while len(side) + len(blok) > plads:
            n = min(plads - len(side), len(blok) - 2)
            if n >= 3 or not side:
                (celler, art), side = blok[0], side + blok[:n]
                blok = [([celler[0] + " (fortsat)"] + celler[1:], art)] + blok[n:]
            sider.append(side)
            side = []
        side += blok
    if side:
        sider.append(side)
    return sider


def fordeling(prs, r, noegle, titel):
    """Fondene grupperet på geografi eller segment, med andel af tilsagnet i DKK.

    Totalrækken står øverst; hver gruppe har en overskrift, sine fonde og en i alt-række.
    Er der flere rækker, end der er plads til, fortsætter tabellen på næste slide med
    kolonneoverskrifterne gentaget. Returnerer listen af slides.
    """
    fd = r["fordelinger"][noegle]

    def linje(tekst, valuta, x, andel):
        return [tekst, valuta or "", tusind(x["tilsagn"]) if valuta else IKKE_OPLYST,
                tusind(x["tilsagn_dkk"]), tusind(x["nav_dkk"]), tusind(x["resttilsagn_dkk"]),
                gange(x["tvpi"]), pct(andel)]

    t = fd["total"]

    def fodnoter():
        """Hver side får sine egne noter; mærkerne sættes i kolonnehovedets rækkefølge."""
        fod = Fodnoter("Kilde: Forvalterrapporter, Danmarks Nationalbank og egne beregninger.")
        tilsagn_m = fod.maerke(
            "Tilsagn i fondens valuta vises på i alt-rækkerne kun, hvor fondene har samme valuta."
            if t["valuta"] is None or any(g["sum"]["valuta"] is None for g in fd["grupper"])
            else None)
        kurs_m = fod.maerke(_kursnote(r))
        tvpi_m = andel_m = ""
        if any("foreløbig" in x for f in r["fonde"] for x in f.get("noter", [])):
            fod.generel(f"{VISTE_NAVNE[noegle]} er en foreløbig klassifikation.")
        fod.generel(_datonoter(r))
        # valutaen står over kolonnenavnet; beløbene er i tusinde, som undertitlen siger
        hoved = ["", "Valuta", f"Tilsagn{tilsagn_m}", f"DKK\nTilsagn{kurs_m}",
                 f"DKK\nNAV{kurs_m}", f"DKK\nResttilsagn{kurs_m}", f"TVPI{tvpi_m}",
                 f"Andel af\ntilsagn{andel_m}"]
        return fod, hoved
    blokke = [[(linje("I alt", t["valuta"], t, t["andel"]), "total")]]
    for g in fd["grupper"]:
        blokke.append(
            [([g["navn"]] + [""] * 7, "gruppe")]
            + [(linje("      " + f["fund_name"] + _datomarkering(r, f["fund_name"]),
                      f["valuta"], f["n"], fd["andele"][f["fund_name"]]), "fond")
               for f in g["fonde"]]
            + [(linje(f"I alt {g['navn']}", g["sum"]["valuta"], g["sum"], g["andel"]), "total")])

    plads = int((Y_NOTER - Y_TOP - 0.15) / FORD_MIN_H) - 2     # kolonnehovedet fylder to rækker
    sider = _del_paa_sider(blokke, plads)
    raekke_h = _raekke_h(max(len(x) for x in sider) + 2, mindst=FORD_MIN_H)

    lavede = []
    for nr, side in enumerate(sider):
        s = _oversigtsslide(prs, r, titel if nr == 0 else f"{titel} (fortsat)",
                            enhed="beløb i tusinde; tilsagn i fondens valuta, øvrige beløb i DKK")
        fod, hoved = fodnoter()
        raekker = [hoved] + [celler for celler, _ in side]
        _tabel(s, X0, Y_TOP, X1 - X0, raekker,
               bredder=[3.3, 0.65, 1.0, 1.1, 1.1, 1.25, 0.75, 0.95],
               just=[V, C, H, H, H, H, H, H], raekke_h=raekke_h, pt=8.5, hoved=True,
               grupper=[i for i, (_, art) in enumerate(side, 1) if art == "gruppe"],
               totaler=[i for i, (_, art) in enumerate(side, 1) if art == "total"])
        fod.tegn(s)
        lavede.append(s)
    return lavede


# ---------- branche- og geografifordeling ----------

BR_MIN = 0.02             # skiver under 2 % af NAV samles i »Øvrige« i cirkeldiagrammet
BR_ETIKET = 0.04          # skiver under 4 % får ingen etiket; tabellen har tallet
BR_RAEKKE_H = 0.155       # laveste rækkehøjde i tabellen (tommer)
BR_NOTER = Y_NOTER - 0.12  # noterne starter en linje højere: forklaringerne fylder fire linjer
BR_GRAF_W = 4.2           # cirkeldiagrammets bredde; tabellen får resten
# skiftevis mørk og lys, så nabo-skiver kan skelnes; beige sidst, da den næsten forsvinder
# mod de cremede sektorrækker, hvor farvefeltet står
BR_FARVER = [SAGE, TAUPE, TERRA, SAGE_LYS, _rgb(style.PPIM_BROWN), OLIVEN,
             _rgb(style.PPIM_RED_LIGHT), _rgb(style.PPIM_TAUPE), ROED, SAGE_MID,
             _rgb(style.PPIM_BEIGE)]
# geografien farves pr. region, så diagrammet viser regionen, og skiftevis mørk og lys
# inden for regionen, så nabo-delregionerne kan skelnes
# (den femte grønne skal være tydeligt lysere end den første, da begge typisk er store)
GEO_FARVER = {"EUROPE": [SAGE, _rgb("#AEB39C"), SAGE_MID, _rgb(style.GRID), OLIVEN],
              "NORTH_AMERICA": [TERRA, TAUPE, _rgb(style.PPIM_BROWN)],
              "ROW": [ROED, _rgb(style.PPIM_RED_LIGHT), _rgb(style.PPIM_TAUPE)]}
# PowerPoint ombryder etiketterne ved en fast bredde og deler lange navne midt i ordet
BR_ETIKET_ORDDELING = {"Diversificeret (fondsandele)": "Diversificeret\n(fondsandele)",
                       "Øvrige Europa og ikke opdelt": "Øvrige Europa\nog ikke opdelt",
                       "Storbritannien og Irland": "Storbritannien\nog Irland",
                       "Nordamerika, ikke opdelt": "Nordamerika,\nikke opdelt",
                       "Tyskland, Østrig og Schweiz": "Tyskland, Østrig\nog Schweiz",
                       "Norden ekskl. Danmark": "Norden\nekskl. Danmark",
                       "Asien og Stillehavsområdet": "Asien og Stille-\nhavsområdet",
                       "Mellemøsten og Afrika": "Mellemøsten\nog Afrika",
                       "Global, ikke opdelt": "Global,\nikke opdelt"}
BR_UDEN_GICS = _rgb("#A6A6A0")   # Diversificeret og Ikke klassificeret, når de får egen skive
BR_OEVRIGE = _rgb("#CFCFC9")    # kun »Øvrige« og »Øvrige fonde«
# én fast farve pr. GICS-sektor på alle slides og i alle dele, så samme farve altid er samme
# sektor; gråt er forbeholdt »Øvrige« og det, der ikke er GICS
SEKTOR_FARVER = {"FINANCIALS": TAUPE, "INDUSTRIALS": TERRA, "HEALTH_CARE": SAGE_LYS,
                 "INFORMATION_TECHNOLOGY": _rgb(style.PPIM_BROWN),
                 "CONSUMER_DISCRETIONARY": OLIVEN, "CONSUMER_STAPLES": _rgb(style.PPIM_RED_LIGHT),
                 "COMMUNICATION_SERVICES": SAGE, "MATERIALS": ROED, "UTILITIES": SAGE_MID,
                 "ENERGY": _rgb(style.PPIM_BEIGE),
                 # husfarverne er brugt af sektorerne og regionerne; en dæmpet petrol skiller
                 # sig ud fra ejendomssegmenternes blågrå skifer (EJD_SEGMENT_FARVER) på siden
                 # pr. aktivklasse uden at tage over som en klar farve ville
                 "REAL_ESTATE": _rgb("#4F7470")}


def _skiver(poster):
    """Skiverne i cirkeldiagrammet: (navn, dkk, andel, farve), og farven pr. post.

    poster er [(navn, dkk, andel, farve)] i tabellens rækkefølge. Poster under BR_MIN samles
    i én grå skive »Øvrige« (en enkelt lille post står under sit eget navn i sin egen farve).
    """
    skiver, farver, oevrige = [], {}, []
    for navn, dkk, andel, farve in poster:
        if (andel or 0) < BR_MIN:
            oevrige.append((navn, dkk, andel, farve))
            continue
        skiver.append((navn, dkk, andel, farve))
        farver[navn] = farve
    if len(oevrige) == 1:
        # en enkelt lille post står under sit eget navn og i sin egen farve; gråt er kun »Øvrige«
        navn, dkk, andel, farve = oevrige[0]
        skiver.append((navn, dkk, andel, farve))
        farver[navn] = farve
    elif oevrige:
        skiver.append(("Øvrige", sum(x[1] for x in oevrige), sum(x[2] or 0 for x in oevrige),
                       BR_OEVRIGE))
        farver |= {x[0]: BR_OEVRIGE for x in oevrige}
    return skiver, farver, [x[0] for x in oevrige]


def _brancheskiver(sektorer):
    """Én skive pr. sektor i sektorens faste farve (SEKTOR_FARVER); det, der ikke er GICS, i
    gråt."""
    return _skiver([(s["navn"], s["dkk"], s["andel"], SEKTOR_FARVER.get(s["kode"], BR_UDEN_GICS))
                    for s in sektorer])


def _geoskiver(regioner):
    """Én skive pr. delregion, farvet med regionens farvefamilie (GEO_FARVER)."""
    poster = []
    for g in regioner:
        familie, i = GEO_FARVER.get(g["kode"]), 0
        for u in g["under"] or [dict(navn=g["navn"], dkk=g["dkk"], andel=g["andel"])]:
            if familie is None:
                farve = BR_UDEN_GICS
            else:
                farve = familie[i % len(familie)]
                i += 1 if (u["andel"] or 0) >= BR_MIN else 0
            poster.append((u["navn"], u["dkk"], u["andel"], farve))
    return _skiver(poster)


def _cirkel(s, x, y, w, h, skiver, serie="NAV, tDKK"):
    """Cirkeldiagram som rigtigt diagram; etiketterne skrives med danske talformater."""
    cd = CategoryChartData(number_format="#,##0")
    cd.categories = [navn for navn, *_ in skiver]
    cd.add_series(serie, [_kt(dkk) for _, dkk, *_ in skiver])
    diagram = s.shapes.add_chart(XL_CHART_TYPE.PIE, Inches(x), Inches(y), Inches(w), Inches(h),
                                 cd).chart
    diagram.has_legend = diagram.has_title = False
    diagram.font.name, diagram.font.size = BROED_FONT, Pt(8)
    plot = diagram.plots[0]
    plot.vary_by_categories = True
    for punkt, (navn, _, andel, farve) in zip(plot.series[0].points, skiver):
        punkt.format.fill.solid()
        punkt.format.fill.fore_color.rgb = farve
        punkt.format.line.color.rgb = HVID
        punkt.format.line.width = Pt(1)
        if andel >= BR_ETIKET:
            etiket = punkt.data_label
            etiket.position = XL_LABEL_POSITION.OUTSIDE_END
            para = etiket.text_frame.paragraphs[0]
            linjer = BR_ETIKET_ORDDELING.get(navn, navn).split("\n") + [pct(andel)]
            for i, tekst in enumerate(linjer):
                if i:
                    para._p.add_br()
                run = para.add_run()
                run.text = tekst
                run.font.name, run.font.size = BROED_FONT, Pt(8)
                run.font.bold = i == len(linjer) - 1          # procenten fed, navnet ikke
                run.font.color.rgb = SORT
    return diagram


def _fordelingstabel(s, b, kolonne, linjer, skiver, ekstra=()):
    """Tabellen til højre for cirkeldiagrammet og diagrammet selv.

    linjer er [(celler, art, farve)] med art "gruppe" (fed på cremet bund) eller "under";
    farve sætter et farvefelt i første kolonne, så tabellen er diagrammets forklaring.
    ekstra er [(overskrift, i alt-tekst)] for kolonner efter Andel (benchmark); cellerne står
    allerede i linjerne.
    """
    hoejde = BR_NOTER - 0.1 - Y_TOP
    raekker = [["", kolonne, "NAV", "Andel"] + [o for o, _ in ekstra],
               ["", "I alt", tusind(b["total_dkk"]), pct(1.0)] + [i for _, i in ekstra]]
    raekker += [celler for celler, _, _ in linjer]
    raekke_h = min(0.24, hoejde / len(raekker))
    tab_x = X0 + BR_GRAF_W + 0.25
    bredder = [0.2, 3.4, 0.9, 0.8] if not ekstra else [0.2, 2.3, 0.85, 0.7] + [0.9] * len(ekstra)
    _tabel(s, tab_x, Y_TOP, X1 - tab_x, raekker, bredder=bredder,
           just=[C, V, H, H] + [H] * len(ekstra), raekke_h=raekke_h, pt=7.5, hoved=True,
           grupper=[i for i, (_, art, _) in enumerate(linjer, 2) if art == "gruppe"],
           totaler=[1])
    kant = 0.085
    for i, (_, _, farve) in enumerate(linjer, 2):
        if farve is not None:
            _bjaelke(s, tab_x + 0.07, Y_TOP + i * raekke_h + (raekke_h - kant) / 2, kant, kant,
                     farve)
    graf_h = min(hoejde, BR_GRAF_W)
    _cirkel(s, X0 - 0.1, Y_TOP + (hoejde - graf_h) / 2, BR_GRAF_W, graf_h, skiver)


def _navne(xs):
    return ", ".join(xs[:-1]) + " og " + xs[-1] if len(xs) > 1 else xs[0]


def _eksponeringsnoter(r, b, fordeling, fordelingen, klassifikationen):
    """Noterne, som branche- og geografisiden deler: placering efter strategi, datoer, fonde
    uden fordeling og den foreløbige klassifikation."""
    noter = []
    for fond, hvor in b["efter_strategi"]:
        noter.append(f"{fond} oplyser ingen {fordeling} og indgår samlet i {hvor}.")
    for fond, dato in b["andre_datoer"]:
        noter.append(f"{fordelingen} for {fond} er pr. {dato:%d.%m.%Y}.")
    if b["uden_fordeling"]:
        noter.append(f"Ingen {fordeling} for {_navne(b['uden_fordeling'])}; indgår som "
                     f"{IKKE_KLASSIFICERET.lower()}.")
    if b["foreloebig_andel"] and b["foreloebig_andel"] >= 0.0005:
        noter.append(f"{klassifikationen} for {pct(b['foreloebig_andel'])} af NAV er en "
                     "foreløbig klassifikation.")
    return noter


def _afslut_noter(s, r, fod, y=BR_NOTER):
    """Kursnoten og datonoterne som generelle noter, og noterne ved y. fod er Fodnoter eller en
    liste af generelle noter."""
    if not isinstance(fod, Fodnoter):
        noter, fod = fod, Fodnoter()
        fod.generel(noter)
    fod.generel(_kursnote(r), _datonoter(r))
    fod.tegn(s, y)


def _diversificeret_noter(b):
    """Forklaringen af »Diversificeret (fondsandele)«, når den findes."""
    return [f"»{x['navn']}« er andele i underliggende fonde, som forvalteren af "
            f"{_navne(x['fonde'])} ikke fordeler på brancher."
            for x in b["grupper"] if x["kode"] == "DIVERSIFIED"]


def brancher(prs, r):
    """Investors NAV fordelt på GICS-sektorer: cirkeldiagram og tabel.

    Tabellen er diagrammets forklaring: hver sektor har sin skives farve. Kun sektorerne
    vises; branchegrupperne står i dataudtrækket (fanen Brancher). None uden DKK-tal for alle
    fonde.
    """
    b = r.get("brancher")
    if not b:
        return None
    s = _oversigtsslide(prs, r, "Branchefordeling",
                        enhed="investors NAV, beløb i tusinde DKK")
    skiver, farver, smaa_sektorer = _brancheskiver(b["grupper"])
    indeks = (r.get("benchmark") or {}).get("indeks") or {}
    linjer = [(["", x["navn"], tusind(x["dkk"]), pct(x["andel"])]
               + [pct(v[x["kode"]]) if x["kode"] in v else IKKE_OPLYST for v in indeks.values()],
               "under", farver[x["navn"]]) for x in b["grupper"]]
    _fordelingstabel(s, b, "Sektor", linjer, skiver,
                     ekstra=[(navn, pct(sum(v.values()))) for navn, v in indeks.items()])

    noter = ["Kilde: Forvalterrapporter, Danmarks Nationalbank og egne beregninger. Investors "
             "NAV er fordelt efter forvalterens branchefordeling eller beholdningernes dagsværdi, "
             "med feederfondene gennemlyst; sektorerne følger GICS."]
    noter += _diversificeret_noter(b)
    noter += _eksponeringsnoter(r, b, "branchefordeling", "Branchefordelingen", "Branchen")
    if len(smaa_sektorer) > 1:
        noter.append(f"Diagrammet samler sektorerne under {pct(BR_MIN, 0)} (grå i tabellen) "
                     "i »Øvrige«.")
    noter += _benchmarknote(r)
    _afslut_noter(s, r, noter)
    return s


def _benchmarknote(r, art="brancher"):
    """Kilden til benchmarket, når der er et: sektorvægtene på branchesiderne, regionvægtene
    på geografisiderne."""
    bm = r.get("benchmark") or {}
    indeks = bm.get("indeks" if art == "brancher" else "regioner")
    if not indeks:
        return []
    dato = f" pr. {bm['dato']:%d.%m.%Y}" if bm.get("dato") else ""
    # indeksets navn står ved båndet, så noten nævner det kun, hvor der er flere
    navn = f" i {_navne(list(indeks))}" if len(indeks) > 1 else ""
    if art == "brancher":
        return [f"Sektorvægtene{navn}{dato}; diversificerede fondsandele har ingen sektor i "
                "indekset."]
    return [f"Regionvægtene{navn}{dato}."]


# lyse farver får sort tekst i branchekortet; de mørke får hvid
KORT_LYSE = {SAGE_LYS, TAUPE, CREAM, BR_OEVRIGE, _rgb(style.PPIM_BEIGE),
             _rgb(style.PPIM_RED_LIGHT), _rgb(style.GRID)}


KORT_MINDST = 0.03        # mindste areal i et kort (andel af fladen), så kassen kan bære en etiket


def _gulv(vaerdier, mindst=KORT_MINDST):
    """Værdierne med de mindste løftet til mindst x summen, og indeksene på de løftede.

    Kun arealet ændres; kasserne viser stadig de faktiske tal, og deres fodnote siger det
    (GULVNOTE).
    """
    gulv = mindst * sum(vaerdier)
    return ([max(v, gulv) if v > 0 else v for v in vaerdier],
            {i for i, v in enumerate(vaerdier) if 0 < v < gulv})


GULVNOTE = ""
GULV_PLADS = "\u2063"     # pladsholder for gulvmærket, mens etiketten prøves af i kassen


def _treemap(vaerdier, x, y, w, h):
    """Kasser (x, y, w, h) med areal i forhold til vaerdier, i samme rækkefølge som dem.

    Squarified treemap (Bruls, Huizing og van Wijk): de største værdier lægges først, i rækker
    langs fladens korte side, og en kasse kommer med i rækken, så længe rækkens dårligste
    side-forhold bliver bedre. Det giver kasser tæt på kvadrater, som er lette at sammenligne.
    """
    i_alt = sum(vaerdier)
    orden = sorted(range(len(vaerdier)), key=lambda i: -vaerdier[i])
    arealer = [vaerdier[i] / i_alt * w * h for i in orden]
    kasser = [None] * len(vaerdier)

    def vaerst(raekke, side):
        s = sum(raekke)
        return max(max(side * side * a / (s * s), s * s / (side * side * a)) for a in raekke)

    i = 0
    while i < len(arealer):
        side = min(w, h)
        raekke, j = [arealer[i]], i + 1
        while j < len(arealer) and vaerst(raekke + [arealer[j]], side) <= vaerst(raekke, side):
            raekke.append(arealer[j])
            j += 1
        tykkelse = sum(raekke) / side
        pos = 0.0
        for k, areal in enumerate(raekke):
            laengde = areal / tykkelse
            # rækken står langs den korte side: lodret i en bred flade, vandret i en høj
            kasser[orden[i + k]] = ((x, y + pos, tykkelse, laengde) if w >= h
                                    else (x + pos, y, laengde, tykkelse))
            pos += laengde
        if w >= h:
            x, w = x + tykkelse, w - tykkelse
        else:
            y, h = y + tykkelse, h - tykkelse
        i = j
    return kasser


def _kortkasse(s, kasse, navn, dkk, andel, farve, fod=None, maerke="", loeftet=False):
    """Én kasse i branche- og geografikortet: navn, beløb og andel, så stort som kassen tillader.

    maerke er fodnotemærker efter navnet (fx »Øvrige«). loeftet: kassen er forstørret (_gulv)
    og får GULVNOTE. En høj, smal kasse får navn og procent på tværs (lodret). Er der kun
    plads til procenten, får kassen én kort fodnote i fod med navn og beløb (og at den er
    forstørret). Returnerer, om navnet kom med.
    """
    fod = fod if fod is not None else Fodnoter()
    # gulvmærket sættes først, når det vides, om navnet kommer med; indtil da en pladsholder
    maerke += GULV_PLADS if loeftet else ""
    kx, ky, kw, kh = kasse
    f = _bjaelke(s, kx, ky, kw, kh, farve)
    f.line.fill.solid()
    f.line.color.rgb, f.line.width = HVID, Pt(2)          # hvide mellemrum mellem kasserne
    tekstfarve = SORT if farve in KORT_LYSE else HVID
    ind = 0.08 if min(kw, kh) >= 0.45 else 0.03     # små kasser: mindre luft, mere plads til tekst
    vist = navn + maerke
    # et navn, der er bredere end kassen, deles som i cirkeldiagrammet i stedet for midt i ordet
    passer = lambda pt: len(vist) * pt * 0.0085 <= kw - 2 * ind
    delt = BR_ETIKET_ORDDELING.get(navn, navn) + maerke

    def har_plads(linjer, bredde=kw, hoejde=kh):
        # et navn, der stadig ombrydes, fylder flere linjer; ellers løber procenten ud af kassen
        brugt = sum(ceil(len(del_) * pt * 0.0085 / max(bredde - 2 * ind, 0.05)) * pt * 1.2 / 72
                    for tekst, pt, _ in linjer for del_ in tekst.split("\n"))
        # og intet ord må være bredere end kassen, for så deles det midt i ordet
        ord_ = max((len(o) * pt * 0.0085 for tekst, pt, _ in linjer for o in tekst.split()),
                   default=0)
        return brugt <= hoejde - 2 * ind + 0.05 and ord_ <= bredde - 2 * ind + 0.02

    store = mellem = None
    uden = False
    if kw >= 1.5 and kh >= 0.85:
        store = [(vist if passer(13) else delt, 13, True),
                 (f"{tusind(dkk)} tDKK", 10.5, False), (pct(andel), 10.5, True)]
    if kw >= 0.95 and kh >= 0.6:
        mellem = [(vist if passer(9.5) else delt, 9.5, True),
                  (f"{tusind(dkk)} tDKK", 8.5, False), (pct(andel), 8.5, True)]
    if store and har_plads(store):
        linjer = store
    elif mellem and har_plads(mellem):
        linjer = mellem
    elif mellem and har_plads(mellem[:1] + [(f"{tusind(dkk)} tDKK  ·  {pct(andel)}", 8.5, True)]):
        linjer = mellem[:1] + [(f"{tusind(dkk)} tDKK  ·  {pct(andel)}", 8.5, True)]
    else:
        # en lille kasse skal stadig sige, hvad den er: navn og procent så småt som nødvendigt,
        # og på tværs, hvis kassen er høj og smal
        smaa = [[(vist, 8, True), (pct(andel), 8, True)], [(f"{vist} {pct(andel)}", 8, True)],
                [(vist, 7, True), (pct(andel), 7, True)], [(f"{vist} {pct(andel)}", 7, True)],
                [(f"{vist} {pct(andel)}", 6, True)]]
        linjer = next((x for x in smaa if har_plads(x)), None)
        if linjer is None and kh > kw:
            # på tværs helst navn og procent på hver sin linje; én linje kun med god margen, da
            # en drejet etiket, der lige akkurat ombrydes, sender procenten ud for sig selv
            linjer = (next((x for x in smaa if len(x) == 2 and har_plads(x, kh, kw)), None)
                      or next((x for x in smaa if len(x) == 1 and har_plads(x, kh * 0.85, kw)),
                              None))
            if linjer:
                return _lodret_etiket(s, kasse, ind, _med_gulvmaerke(linjer, fod), tekstfarve)
        if linjer is None:
            # kun procenten: én fodnote siger, hvilken kasse det er
            uden_m = fod.maerke(f"{navn}, {tusind(dkk)} tDKK"
                                + ("" if loeftet else "") + ".")
            linjer = [(f"{uden_m}{maerke.replace(GULV_PLADS, '')} {pct(andel)}".lstrip(), 7,
                       True)]
            uden = True
    linjer = _med_gulvmaerke(linjer, fod)
    tf = _tekst(s, kx + ind, ky + ind, max(kw - 2 * ind, 0.1), max(kh - 2 * ind, 0.1),
                linjer[0][0], linjer[0][1], fed=linjer[0][2], farve=tekstfarve)
    for tekst, pt, fed in linjer[1:]:
        para = tf.add_paragraph()
        run = para.add_run()
        run.text = tekst
        run.font.name, run.font.size, run.font.bold = BROED_FONT, Pt(pt), fed
        run.font.color.rgb = tekstfarve
    return not uden


def _med_gulvmaerke(linjer, fod):
    """Linjerne med gulvmærkets pladsholder skiftet ud med mærket til GULVNOTE."""
    if not any(GULV_PLADS in t for t, _, _ in linjer):
        return linjer
    m = fod.maerke(GULVNOTE)
    return [(t.replace(GULV_PLADS, m), pt, fed) for t, pt, fed in linjer]


def _lodret_etiket(s, kasse, ind, linjer, farve):
    """Etiketten i en høj, smal kasse, drejet en kvart omgang, så den læses nedefra og op."""
    kx, ky, kw, kh = kasse
    b, h = kh - 2 * ind, kw - 2 * ind
    # tekstboksen drejes om sit centrum, så den lægges med centrum i kassens centrum
    tf = _tekst(s, kx + kw / 2 - b / 2, ky + kh / 2 - h / 2, b, max(h, 0.1),
                linjer[0][0], linjer[0][1], fed=linjer[0][2], farve=farve)
    for tekst, pt, fed in linjer[1:]:
        run = tf.add_paragraph().add_run()
        run.text = tekst
        run.font.name, run.font.size, run.font.bold = BROED_FONT, Pt(pt), fed
        run.font.color.rgb = farve
    s.shapes[-1].rotation = 270
    return True


def _branchekasser(s, b, x, y, w, h, fod=None, base="NAV", del_navn=None):
    """Sektorerne i b som kasser i fladen (x, y, w, h).

    Returnerer (navnene i »Øvrige«, om små kasser blev gjort større). fod (Fodnoter) får
    noterne til »Øvrige« (sektorerne under BR_MIN af base, i delen del_navn), til
    Diversificeret og til de løftede kasser, med mærket i kassen.
    """
    fod = fod if fod is not None else Fodnoter()
    skiver, _, smaa = _brancheskiver(b["grupper"])
    vaerdier, loeftet = _gulv([max(dkk, 0.0) for _, dkk, _, _ in skiver])
    kasser = _treemap(vaerdier, x, y, w, h)
    kode = {g["navn"]: g["kode"] for g in b["grupper"]}
    diversificeret = dict(zip([g["navn"] for g in b["grupper"] if g["kode"] == "DIVERSIFIED"],
                              _diversificeret_noter(b)))
    for i, (kasse, (navn, dkk, andel, farve)) in enumerate(zip(kasser, skiver)):
        note = None
        if navn == "Øvrige" and len(smaa) > 1:
            # én note til »Øvrige«, med forklaringen af Diversificeret, hvis den ligger i den
            note = " ".join([f"»Øvrige«{f' i {del_navn}' if del_navn else ''} samler "
                             f"{_navne(smaa)}, hver under {pct(BR_MIN, 0)} af {base}."]
                            + [diversificeret[n] for n in smaa if n in diversificeret])
        elif kode.get(navn) == "DIVERSIFIED":
            note = diversificeret.get(navn)
        _kortkasse(s, kasse, navn, dkk, andel, farve, fod, fod.maerke(note), i in loeftet)
    return smaa, loeftet


# ---------- benchmarkbåndene under kortene ----------

BM_BAAND_H = 0.17         # højden af ét indeks' bånd (tommer)
BM_RAEKKE_LUFT = 0.05     # mellem båndene
BM_ETIKET_H = 0.14        # over det første og under det sidste bånd: etiketter til smalle stykker
BM_TOP_LUFT = 0.08        # mellem kortet og båndene
BM_NAVN_W = 0.85          # indeksets navn til venstre for båndet
BM_PT = 7
BM_NOTER_LINJE = 0.12    # noterne rykker en linje op på sider, der i forvejen er fulde


def _bm_hoejde(raekker):
    """Hvad båndene fylder under kortet; 0 uden benchmark."""
    if not raekker:
        return 0.0
    n = len(raekker)
    return BM_TOP_LUFT + 2 * BM_ETIKET_H + n * BM_BAAND_H + (n - 1) * BM_RAEKKE_LUFT


def _bm_tekstbredde(tekst, pt=BM_PT):
    return len(tekst) * pt * 0.0085 + 0.06


def _benchmarkbaand(s, y, raekker, maerke=""):
    """Ét 100 %-bånd pr. indeks under kortet, fra y: stykkerne i kategoriernes farver, så de
    kan sammenlignes med kasserne over dem.

    raekker er [(indeksnavn, [(navn, vægt, farve)])]; maerke står efter indeksets navn. Et
    stykke viser navn og vægt, eller kun
    vægten, når navnet ikke er plads til; er heller ikke den plads, står vægten over det første
    bånd eller under det sidste.
    """
    x, w = X0 + BM_NAVN_W, X1 - X0 - BM_NAVN_W
    by = y + BM_TOP_LUFT + BM_ETIKET_H
    for i, (indeks, stykker) in enumerate(raekker):
        _tekst(s, X0, by, BM_NAVN_W - 0.05, BM_BAAND_H, indeks + maerke, 7.5, fed=True,
               anker=MSO_ANCHOR.MIDDLE)
        total = sum(v for _, v, _ in stykker) or 1.0
        sx = x
        for navn, vaegt, farve in stykker:
            sw = w * vaegt / total
            f = _bjaelke(s, sx, by, sw, BM_BAAND_H, farve)
            f.line.fill.solid()
            f.line.color.rgb, f.line.width = HVID, Pt(1)
            tekstfarve = SORT if farve in KORT_LYSE else HVID
            etiket = next((t for t in (f"{navn} {pct(vaegt)}", pct(vaegt))
                           if _bm_tekstbredde(t) <= sw), None)
            if etiket:
                _tekst(s, sx, by, sw, BM_BAAND_H, etiket, BM_PT, fed=True, farve=tekstfarve,
                       just=C, anker=MSO_ANCHOR.MIDDLE)
            elif i in (0, len(raekker) - 1):
                # udenfor: over det første bånd, under det sidste
                over = i == 0
                ly = by - BM_ETIKET_H if over else by + BM_BAAND_H
                lw = _bm_tekstbredde(pct(vaegt), 6.5)
                _tekst(s, sx + sw / 2 - lw / 2, ly, lw, BM_ETIKET_H, pct(vaegt), 6.5,
                       just=C, anker=MSO_ANCHOR.BOTTOM if over else MSO_ANCHOR.TOP)
            sx += sw
        by += BM_BAAND_H + BM_RAEKKE_LUFT


def _viste_sektorer(fordelinger):
    """Sektorkoderne, der har egen kasse i mindst én af fordelingerne (ikke i »Øvrige«)."""
    viste = set()
    for b in fordelinger:
        _, farver, _ = _brancheskiver(b["grupper"])
        viste |= {g["kode"] for g in b["grupper"]
                  if farver.get(g["navn"], BR_OEVRIGE) != BR_OEVRIGE}
    return viste


def _bm_sektorer(r):
    """Båndene til branchekortene: [(indeks, [(sektor, vægt, farve)])], største først.

    Alle branchekort viser det samme bånd, så indeksets tal er ens på tværs af sliderne:
    indekset, som det er, og sektorer uden egen kasse i hele porteføljens kort samlet i
    »Øvrige« sidst, som i det kort.
    """
    viste = _viste_sektorer([r["brancher"]]) if r.get("brancher") else set()
    raekker = []
    for indeks, vaegte in ((r.get("benchmark") or {}).get("indeks") or {}).items():
        egne = [k for k in vaegte if k in viste]
        andre = [k for k in vaegte if k not in viste]
        if len(andre) == 1:              # én sektor alene står under sit eget navn, som i kortet
            egne, andre = egne + andre, []
        stykker = sorted(((SEKTORER.get(k, k), vaegte[k], SEKTOR_FARVER.get(k, BR_UDEN_GICS))
                          for k in egne), key=lambda x: -x[1])
        if andre:
            stykker.append(("Øvrige", sum(vaegte[k] for k in andre), BR_OEVRIGE))
        if stykker:
            raekker.append((indeks, stykker))
    return raekker


def _bm_regioner(r):
    """Båndene til geografikortet: [(indeks, [(region, vægt, farve)])] i regionernes farver."""
    return [(indeks, sorted(((REGIONER.get(k, k), v, (GEO_FARVER.get(k) or [BR_UDEN_GICS])[0])
                             for k, v in vaegte.items()), key=lambda x: -x[1]))
            for indeks, vaegte in ((r.get("benchmark") or {}).get("regioner") or {}).items()
            if vaegte]


def _bm_maerke(fod, noter):
    """Ét mærke ved indeksets navn til benchmarkets kilde og forklaringen af båndene."""
    return fod.maerke(" ".join(noter))


def _bm_baandnote(raekker, art, hele=True):
    """hele: sliden er hele porteføljens kort; ellers er båndet det samme som dér."""
    if not raekker:
        return []
    if art == "brancher" and not hele:
        return ["»Øvrige« er sektorerne uden egen kasse."]
    if art == "brancher":
        return ["Farverne er kassernes; »Øvrige« er sektorerne uden egen kasse."]
    return [""]


def branchekort(prs, r):
    """Branchefordelingen for hele porteføljen, ejendommene medregnet som sektoren Real Estate:
    én kasse pr. sektor over hele sliden, arealet er andelen.

    Kasserne har sektorernes faste farver og »Øvrige« for sektorer under BR_MIN. Hver kasse
    viser sektor, beløb i tDKK og andel af NAV. None uden DKK-tal for alle fonde.
    """
    b = r.get("brancher")
    if not b:
        return None
    s = _oversigtsslide(prs, r, "Branchefordeling – hele porteføljen",
                        enhed="investors NAV, beløb i tusinde DKK")
    raekker = _bm_sektorer(r)
    fod = Fodnoter("Kilde: Forvalterrapporter, Danmarks Nationalbank og egne beregninger. "
                   "Kassernes areal er sektorens andel, fordelt efter "
                   "forvalterens branchefordeling eller beholdningernes dagsværdi, med "
                   "feederfondene gennemlyst; sektorerne følger GICS, og ejendommene indgår som "
                   "Real Estate.")
    # noterne fylder en linje mere end på geografisiden (Øvrige, Diversificeret, benchmark)
    h = EJD_NOTER - 0.1 - Y_TOP - _bm_hoejde(raekker)
    _branchekasser(s, b, X0, Y_TOP, X1 - X0, h, fod)
    if raekker:
        _benchmarkbaand(s, Y_TOP + h, raekker, _bm_maerke(
            fod, _benchmarknote(r) + _bm_baandnote(raekker, "brancher")))
    fod.generel(_eksponeringsnoter(r, b, "branchefordeling", "Branchefordelingen", "Branchen"))
    _afslut_noter(s, r, fod, EJD_NOTER)
    return s


GEO_HOVED_H = 0.30        # regionens overskriftsbånd i geografikortet (tommer)
GEO_HOVED = _rgb("#3A3F33")   # båndets farve: en dyb salvie, mørkere end alle delregionerne


def _regionsramme(s, kasse):
    """En hvid ramme om en region, så regionerne skilles tydeligere end delregionerne."""
    kx, ky, kw, kh = kasse
    f = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(kx), Inches(ky), Inches(kw), Inches(kh))
    f.fill.background()
    f.line.color.rgb, f.line.width = HVID, Pt(3)
    f.shadow.inherit = False
    return f


def regionskort(prs, r):
    """Geografien for hele porteføljen, ejendommene medregnet: én kasse pr. region i regionens
    farve over hele sliden. None uden DKK-tal for alle fonde.
    """
    b = _kun_regioner(r.get("regioner"))
    if not b:
        return None
    s = _oversigtsslide(prs, r, "Geografisk eksponering – hele porteføljen",
                        enhed="investors NAV, beløb i tusinde DKK")
    fod = Fodnoter("Kilde: Forvalterrapporter, Danmarks Nationalbank og egne beregninger. "
                   "Kassernes areal er regionens andel, fordelt efter "
                   "forvalterens geografiske fordeling eller beholdningernes dagsværdi, med "
                   "feederfondene gennemlyst.")
    raekker = _bm_regioner(r)
    h = BR_NOTER - 0.1 - Y_TOP - _bm_hoejde(raekker)
    oevrige_noter, _ = _regionskasser(s, b, X0, Y_TOP, X1 - X0, h, fod)
    if raekker:
        _benchmarkbaand(s, Y_TOP + h, raekker, _bm_maerke(
            fod, _benchmarknote(r, "regioner") + _bm_baandnote(raekker, "regioner")))
    fod.generel(_regionsnoter(oevrige_noter),
                _eksponeringsnoter(r, b, "geografisk fordeling", "Den geografiske fordeling",
                                   "Geografien"))
    _afslut_noter(s, r, fod)
    return s


def _regionsnoter(oevrige_noter):
    noter = []
    if oevrige_noter:
        noter.append("; ".join(oevrige_noter) + f", hver under {pct(BR_MIN, 0)} af NAV.")
    return noter


def _regionskasser(s, b, x, y, w, h, fod=None):
    """Regionerne i b som kasser i fladen (x, y, w, h), hver delt i sine delregioner.

    Returnerer (»Øvrige«-noterne, om små kasser blev gjort større); fod (Fodnoter) får
    noterne til de løftede kasser og kasserne uden plads til navnet (_kortkasse).
    """
    fod = fod if fod is not None else Fodnoter()
    regioner_ = [g for g in b["grupper"] if g["dkk"] > 0]
    vaerdier, loeftet = _gulv([g["dkk"] for g in regioner_])
    kasser = _treemap(vaerdier, x, y, w, h)
    oevrige_noter = []
    for i, (kasse, g) in enumerate(zip(kasser, regioner_)):
        familie = GEO_FARVER.get(g["kode"]) or [BR_UDEN_GICS]
        kx, ky, kw, kh = kasse
        if not g["under"] or kw < 1.2 or kh < 0.9:
            _kortkasse(s, kasse, g["navn"], g["dkk"], g["andel"], familie[0], fod,
                       loeftet=i in loeftet)
            _regionsramme(s, kasse)
            continue
        # båndet øverst: kun regionens navn. Regionens beløb står allerede fordelt i
        # delregionerne under det; stod det også i båndet, ville sliden tælle det to gange og
        # summere til 200 %.
        baand = _bjaelke(s, kx, ky, kw, GEO_HOVED_H, GEO_HOVED)
        baand.line.fill.background()
        _tekst(s, kx + 0.08, ky, kw - 0.16, GEO_HOVED_H, g["navn"], 10, fed=True, farve=HVID,
               anker=MSO_ANCHOR.MIDDLE)
        # delregionerne med samme farver som i cirkeldiagrammet; de små samles i »Øvrige«
        store = [u for u in g["under"] if (u["andel"] or 0) >= BR_MIN]
        smaa = [u for u in g["under"] if (u["andel"] or 0) < BR_MIN]
        poster = [(u["navn"], u["dkk"], u["andel"], familie[i % len(familie)])
                  for i, u in enumerate(store)]
        if smaa:
            poster.append((smaa[0]["navn"] if len(smaa) == 1 else "Øvrige",
                           sum(u["dkk"] for u in smaa), sum(u["andel"] or 0 for u in smaa),
                           BR_OEVRIGE))
            if len(smaa) > 1:
                oevrige_noter.append(f"»Øvrige« under {g['navn']} samler "
                                     f"{_navne([u['navn'] for u in smaa])}")
        poster = [p for p in poster if p[1] > 0]
        vaerdier, indre_loeftet = _gulv([p[1] for p in poster])
        indre = _treemap(vaerdier, kx, ky + GEO_HOVED_H, kw, kh - GEO_HOVED_H)
        for j, (k, (navn, dkk, andel, farve)) in enumerate(zip(indre, poster)):
            _kortkasse(s, k, navn, dkk, andel, farve, fod, loeftet=j in indre_loeftet)
        loeftet = loeftet | {(i, j) for j in indre_loeftet}
        _regionsramme(s, kasse)
    return oevrige_noter, loeftet


# ---------- overblikkene med ejendommene for sig ----------

# ejendomssegmenterne i Real Estate-farvens familie (skifer), hvert med sin faste nuance, som
# delregionerne i deres regions familie; uden segment er »Øvrige« i samme lysegrå som andre
# »Øvrige«
EJD_SEGMENT_FARVER = {"RESIDENTIAL": _rgb("#4B565E"), "OFFICE": _rgb("#7F8B93"),
                      "LOGISTICS": _rgb("#3C464D"), "RETAIL": _rgb("#9AA4AA"),
                      "HOTEL": _rgb("#66727A"), "HEALTHCARE": _rgb("#B8C0C5"),
                      "MIXED": _rgb("#CDD3D6")}
KORT_LYSE |= {EJD_SEGMENT_FARVER[k] for k in ("RETAIL", "HEALTHCARE", "MIXED")}
EJD_MIN_W = 2.4           # ejendomsdelens mindste bredde (tommer), så navnene kan stå
EJD_LUFT = 0.15           # mellemrum mellem de to dele
EJD_NOTER = BR_NOTER - 0.12   # en linje mere til noterne: delene skal også forklares


def _delbaand(s, x, y, w, navn, dkk, andel):
    """Båndet over en del: navn, beløb og delens andel af hele NAV, så stort som bredden tillader."""
    baand = _bjaelke(s, x, y, w, GEO_HOVED_H, GEO_HOVED)
    baand.line.fill.background()
    tekst = f"{navn}  ·  {tusind(dkk)} tDKK  ·  {pct(andel)}"
    pt = next((pt for pt in (10, 9, 8) if len(tekst) * pt * 0.0075 <= w - 0.16), 7.5)
    _tekst(s, x + 0.08, y, w - 0.16, GEO_HOVED_H, tekst, pt, fed=True, farve=HVID,
           anker=MSO_ANCHOR.MIDDLE)


def _delbredder(andele, w, mindst):
    """Bredderne for dele med andelene: proportionale, men mindst `mindst` hver (lige brede,
    hvis der ikke er plads til det)."""
    if len(andele) * mindst >= w:
        return [w / len(andele)] * len(andele)
    faste = {}
    while True:
        ledig = w - sum(faste.values())
        rest = sum(a for i, a in enumerate(andele) if i not in faste) or 1.0
        for_smalle = {i: mindst for i, a in enumerate(andele)
                      if i not in faste and ledig * a / rest < mindst}
        if not for_smalle:
            return [faste.get(i, ledig * a / rest) for i, a in enumerate(andele)]
        faste |= for_smalle


def _delt_flade(s, dele, ejd, noter_y=EJD_NOTER, fod=None):
    """Fladen delt i delene fra venstre og ejendommene til højre, hver med sit bånd og sine
    kasser.

    dele er [(navn, total_dkk, andel af hele NAV, tegn)], hvor tegn(x, y, w, h) tegner delens
    kasser; ejd er ejendomsdelen fra data (None uden ejendomme), vist pr. ejendomssegment eller
    region (_ejendomskasser). Bredderne
    følger andelene, med EJD_MIN_W som mindste bredde. Ejendommenes noter (_ejendomsnoter) får
    deres mærke i Ejendomme-båndet; fod (Fodnoter) får også kassernes noter. Returnerer
    (ejendomsdelens øvrige noter, om små ejendomskasser blev gjort større, [hvad tegn
    returnerer pr. del]).
    """
    fod = fod if fod is not None else Fodnoter()
    alle = list(dele) + ([("Ejendomme" + fod.maerker(_ejendomsnoter(ejd)), ejd["total_dkk"],
                           ejd["andel"], None)] if ejd else [])
    w = X1 - X0 - EJD_LUFT * (len(alle) - 1)
    h = noter_y - 0.1 - Y_TOP
    ky, kh = Y_TOP + GEO_HOVED_H, h - GEO_HOVED_H
    x, resultater, ejd_noter, loeftet = X0, [], [], False
    for (navn, dkk, andel, tegn), bw in zip(alle, _delbredder([a[2] for a in alle], w,
                                                               EJD_MIN_W)):
        _delbaand(s, x, Y_TOP, bw, navn, dkk, andel)
        if tegn is not None:
            resultater.append(tegn(x, ky, bw, kh))
        else:
            ejd_noter, loeftet = _ejendomskasser(s, ejd, x, ky, bw, kh, fod)
        x += bw + EJD_LUFT
    return ejd_noter, loeftet, resultater


def _ejendomskasser(s, ejd, x, y, w, h, fod):
    """Ejendommene som kasser i fladen: i brancherne pr. ejendomssegment, i geografien pr.
    region i regionernes farver, som de andre dele. Returnerer (de øvrige noter, om små kasser
    blev gjort større); _ejendomsnoter står ved båndet (_delt_flade), og geografiens
    kildenoter (_eksponeringsnoter) tilføjer siden selv."""
    ford = ejd["fordeling"]
    if ejd["art"] == "regioner":
        oevrige_noter, loeftet = _regionskasser(s, _kun_regioner(ford), x, y, w, h, fod)
        return _paa_delens_base(_regionsnoter(oevrige_noter), "Ejendomme"), loeftet
    skiver, _, smaa = _skiver([(g["navn"], g["dkk"], g["andel"],
                                EJD_SEGMENT_FARVER.get(g["kode"], BR_OEVRIGE))
                               for g in ford["grupper"]])
    vaerdier, loeftet = _gulv([k[1] for k in skiver])
    for i, (kasse, (navn, dkk, andel, farve)) in enumerate(zip(_treemap(vaerdier, x, y, w, h),
                                                               skiver)):
        note = _ejendomsoevrige_note(ejd, smaa) if navn == SEGMENT_OEVRIGE else None
        _kortkasse(s, kasse, navn, dkk, andel, farve, fod, fod.maerke(note), i in loeftet)
    _regionsramme(s, (x, y, w, h))
    return [], loeftet


def _ejendomsoevrige_note(ejd, smaa):
    """Noten til »Øvrige« i ejendomsdelen: fondene uden ejendomssegment, der ikke kan
    klassificeres, og de små segmenter, der er lagt til."""
    uden, blandede = ejd["fordeling"]["uden_segment"], ejd["blandede"]
    dele = []
    if uden:
        hvad = "ejendomsandelen af " if set(uden) <= set(blandede) else ""
        dele.append(f"{hvad}{_navne(uden)}, som ikke kan klassificeres efter ejendomssegment")
    andre = [n for n in smaa if n != SEGMENT_OEVRIGE]
    if len(andre) > 1 or (andre and uden):
        dele.append(f"{_navne(andre)}, hver under {pct(BR_MIN, 0)} af ejendomsdelen")
    return ("»Øvrige« i ejendomsdelen er " + " samt ".join(dele) + ".") if dele else None


def _ejendomsopdelt(s, opd, tegn_resten, fod=None, noter_y=EJD_NOTER):
    """Fladen delt i de alternative investeringer (til venstre) og ejendomme (til højre).
    Returnerer (ejendomsnoterne, om små ejendomskasser blev gjort større, hvad tegn_resten
    returnerer)."""
    ejd, rest = opd["ejendomme"], opd["resten"]
    ejd_noter, loeftet, (resultat,) = _delt_flade(
        s, [("Alternative investeringer", rest["total_dkk"], 1 - ejd["andel"], tegn_resten)],
        ejd, noter_y=noter_y, fod=fod)
    return ejd_noter, loeftet, resultat


def _ejendomsnoter(ejd):
    """Hvad ejendomsdelen består af, og hvordan den er fordelt."""
    blandede = ejd["blandede"]
    uden = ejd["fordeling"]["uden_segment"] if ejd["art"] == "brancher" else []
    noter = ["Procenterne i kasserne er andel af NAV. "
             "Ejendomme er GICS-sektoren Real Estate på tværs af fondene"]
    if blandede and not set(blandede) <= set(uden):
        noter[0] += f", herunder ejendomsandelen af {_navne(blandede)}"
    if ejd["art"] == "brancher":
        # »Øvrige« har sin egen note i kassen (_ejendomsoevrige_note)
        noter[0] += ", fordelt efter fondenes ejendomssegment"
    noter[0] += "."
    if ejd["art"] == "regioner" and blandede:
        noter.append("Ejendomsandelen af en blandet fond antages at ligge geografisk som resten "
                     "af fonden.")
    return noter


def _ejendomsgeografinoter(r, ejd):
    """Kildenoterne for ejendommenes geografi (forvalterens fordeling, fokus osv.)."""
    if not ejd or ejd["art"] != "regioner":
        return []
    return _paa_delens_base(_eksponeringsnoter(r, _kun_regioner(ejd["fordeling"]),
                                               "geografisk fordeling",
                                               "Den geografiske fordeling", "Geografien"),
                            "Ejendomme")


def _paa_restens_base(noter):
    """Restens noter siger »af NAV« om andele, der her er af de alternative investeringer."""
    return [n.replace("af NAV", "af de alternative investeringer") for n in noter]


def branchekort_ejendomme(prs, r):
    """Branchekortet med ejendommene for sig: de alternative investeringer fordelt på
    sektorer og ejendomme pr. fond, hver med andele af sin egen del. None uden DKK-tal eller
    uden ejendomme."""
    opd = (r.get("uden_ejendomme") or {}).get("brancher")
    if not opd or not opd["resten"]:
        return None
    rest = opd["resten"]
    s = _oversigtsslide(prs, r, "Branchefordeling – ejendomme for sig",
                        enhed="investors NAV, beløb i tusinde DKK")
    raekker = _bm_sektorer(r)
    bund = EJD_NOTER - _bm_hoejde(raekker)
    fod = Fodnoter("Kilde: Forvalterrapporter, Danmarks Nationalbank og egne beregninger. "
                   "Fordelt efter forvalterens branchefordeling eller "
                   "beholdningernes dagsværdi, med feederfondene gennemlyst; sektorerne følger "
                   "GICS.")
    ejd_noter, _, _ = _ejendomsopdelt(s, opd, lambda x, y, w, h: _branchekasser(
        s, rest, x, y, w, h, fod, base="de alternative investeringer"), fod, noter_y=bund)
    if raekker:
        _benchmarkbaand(s, bund - 0.1, raekker,
                        _bm_maerke(fod, _benchmarknote(r)
                                   + _bm_baandnote(raekker, "brancher", hele=False)))
    fod.generel(ejd_noter, _paa_restens_base(_eksponeringsnoter(
        r, rest, "branchefordeling", "Branchefordelingen", "Branchen")))
    _afslut_noter(s, r, fod, EJD_NOTER)
    return s


def regionskort_ejendomme(prs, r):
    """Geografikortet med ejendommene for sig: de alternative investeringer fordelt på
    regioner og ejendomme pr. fond, hver med andele af sin egen del. None uden DKK-tal eller
    ejendomme."""
    opd = (r.get("uden_ejendomme") or {}).get("regioner")
    if not opd or not opd["resten"]:
        return None
    rest = _kun_regioner(opd["resten"])
    s = _oversigtsslide(prs, r, "Geografisk eksponering – ejendomme for sig",
                        enhed="investors NAV, beløb i tusinde DKK")
    fod = Fodnoter("Kilde: Forvalterrapporter, Danmarks Nationalbank og egne beregninger. "
                   "Fordelt efter forvalterens geografiske fordeling eller "
                   "beholdningernes dagsværdi, med feederfondene gennemlyst.")
    ejd_noter, _, (oevrige_noter, _) = _ejendomsopdelt(
        s, opd, lambda x, y, w, h: _regionskasser(s, rest, x, y, w, h, fod), fod)
    fod.generel(ejd_noter,
                _paa_restens_base(_regionsnoter(oevrige_noter)
                                  + _eksponeringsnoter(r, rest, "geografisk fordeling",
                                                       "Den geografiske fordeling", "Geografien")),
                _ejendomsgeografinoter(r, opd["ejendomme"]))
    _afslut_noter(s, r, fod, EJD_NOTER)
    return s


# Ejendomme forklares af _ejendomsnoter; her kun de to andre dele
AKTIVKLASSE_NOTE = ("Private Credit er fondene med private credit som strategi, Private Equity "
                    "de øvrige fonde, begge uden deres ejendomsandel.")
AK_NOTER = EJD_NOTER - 0.12   # endnu en linje til noterne: tre dele skal forklares
DELENS_BASE = "Procenterne i kasserne er andel af NAV."


def _paa_delens_base(noter, navn):
    """Delens noter siger »af NAV« om andele, der her er af delen."""
    return [n.replace("af NAV", f"af {navn}-delen") for n in noter]


def branchekort_aktivklasser(prs, r):
    """Branchekortet delt i Private Equity, Private Credit og ejendomme, hver med andele af sin
    egen del. None uden DKK-tal."""
    opd = (r.get("aktivklasser") or {}).get("brancher")
    if not opd or not opd["dele"]:
        return None
    s = _oversigtsslide(prs, r, "Branchefordeling pr. aktivklasse",
                        enhed="investors NAV, beløb i tusinde DKK; Private Equity, Private "
                              "Credit og ejendomme")
    raekker = _bm_sektorer(r)
    # med båndene fylder noterne en linje mere (båndenes kilde)
    noter_y = AK_NOTER - (BM_NOTER_LINJE if raekker else 0)
    bund = noter_y - _bm_hoejde(raekker)
    fod = Fodnoter("Kilde: Forvalterrapporter, Danmarks Nationalbank og egne beregninger. "
                   "Fordelt efter forvalterens branchefordeling eller "
                   "beholdningernes dagsværdi, med feederfondene gennemlyst; sektorerne følger "
                   "GICS.")
    if not opd["ejendomme"]:
        fod.generel(DELENS_BASE)
    del_m = fod.maerke(AKTIVKLASSE_NOTE)

    def tegner(d):
        return lambda x, y, w, h: _branchekasser(s, d["fordeling"], x, y, w, h, fod,
                                                 base="delen", del_navn=d["navn"])

    ejd_noter, _, _ = _delt_flade(
        s, [(d["navn"] + del_m, d["total_dkk"], d["andel"], tegner(d)) for d in opd["dele"]],
        opd["ejendomme"], noter_y=bund, fod=fod)
    if raekker:
        _benchmarkbaand(s, bund - 0.1, raekker,
                        _bm_maerke(fod, _benchmarknote(r)
                                   + _bm_baandnote(raekker, "brancher", hele=False)))
    fod.generel(ejd_noter)
    for d in opd["dele"]:
        fod.generel(_paa_delens_base(_eksponeringsnoter(r, d["fordeling"], "branchefordeling",
                                                        "Branchefordelingen", "Branchen"),
                                     d["navn"]))
    _afslut_noter(s, r, fod, noter_y)
    return s


def regionskort_aktivklasser(prs, r):
    """Geografikortet delt i Private Equity, Private Credit og ejendomme, hver med andele af
    sin egen del. None uden DKK-tal."""
    opd = (r.get("aktivklasser") or {}).get("regioner")
    if not opd or not opd["dele"]:
        return None
    s = _oversigtsslide(prs, r, "Geografisk eksponering pr. aktivklasse",
                        enhed="investors NAV, beløb i tusinde DKK; Private Equity, Private "
                              "Credit og ejendomme")
    dele = [d | {"fordeling": _kun_regioner(d["fordeling"])} for d in opd["dele"]]

    raekker = _bm_regioner(r)
    # med båndene fylder noterne en linje mere (båndenes kilde)
    noter_y = AK_NOTER - (BM_NOTER_LINJE if raekker else 0)
    bund = noter_y - _bm_hoejde(raekker)
    fod = Fodnoter("Kilde: Forvalterrapporter, Danmarks Nationalbank og egne beregninger. "
                   "Fordelt efter forvalterens geografiske fordeling eller "
                   "beholdningernes dagsværdi, med feederfondene gennemlyst.")
    if not opd["ejendomme"]:
        fod.generel(DELENS_BASE)
    del_m = fod.maerke(AKTIVKLASSE_NOTE)

    def tegner(b):
        return lambda x, y, w, h: _regionskasser(s, b, x, y, w, h, fod)

    ejd_noter, _, resultater = _delt_flade(
        s, [(d["navn"] + del_m, d["total_dkk"], d["andel"], tegner(d["fordeling"]))
            for d in dele], opd["ejendomme"], noter_y=bund, fod=fod)
    if raekker:
        _benchmarkbaand(s, bund - 0.1, raekker, _bm_maerke(
            fod, _benchmarknote(r, "regioner") + _bm_baandnote(raekker, "regioner")))
    fod.generel(ejd_noter)
    for d, (oevrige_noter, _) in zip(dele, resultater):
        fod.generel(_paa_delens_base(_regionsnoter(oevrige_noter)
                                     + _eksponeringsnoter(r, d["fordeling"],
                                                          "geografisk fordeling",
                                                          "Den geografiske fordeling",
                                                          "Geografien"), d["navn"]))
    fod.generel(_ejendomsgeografinoter(r, opd["ejendomme"]))
    _afslut_noter(s, r, fod, noter_y)
    return s


def _kun_regioner(b):
    """b med regionerne alene: sliderne viser ikke delregionerne (de står i Excel).

    Noten om fonde placeret efter strategi nævner derfor regionen, ikke delregionen.
    """
    if not b:
        return b
    region = {u["navn"]: g["navn"] for g in b["grupper"] for u in g["under"]}
    return {**b, "grupper": [{**g, "under": []} for g in b["grupper"]],
            "efter_strategi": [(fond, region.get(hvor, hvor))
                               for fond, hvor in b.get("efter_strategi", [])]}


def regioner(prs, r):
    """Investors NAV fordelt på regionerne Europa, Nordamerika og Resten af verden:
    cirkeldiagram og tabel, én skive og én række pr. region i regionens farve. None uden
    DKK-tal for alle fonde.
    """
    b = _kun_regioner(r.get("regioner"))
    if not b:
        return None
    s = _oversigtsslide(prs, r, "Geografisk eksponering",
                        enhed="investors NAV, beløb i tusinde DKK")
    skiver, farver, smaa = _geoskiver(b["grupper"])

    indeks = (r.get("benchmark") or {}).get("regioner") or {}
    linjer = [(["", g["navn"], tusind(g["dkk"]), pct(g["andel"])]
               + [pct(v[g["kode"]]) if g["kode"] in v else IKKE_OPLYST for v in indeks.values()],
               "gruppe", farver.get(g["navn"])) for g in b["grupper"]]
    _fordelingstabel(s, b, "Region", linjer, skiver,
                     ekstra=[(navn, pct(sum(v.values()))) for navn, v in indeks.items()])

    noter = ["Kilde: Forvalterrapporter, Danmarks Nationalbank og egne beregninger. Investors "
             "NAV er fordelt efter forvalterens geografiske fordeling eller beholdningernes "
             "dagsværdi, med feederfondene gennemlyst."]
    noter += _eksponeringsnoter(r, b, "geografisk fordeling", "Den geografiske fordeling",
                                "Geografien")
    if len(smaa) > 1:
        noter.append(f"Diagrammet samler regionerne under {pct(BR_MIN, 0)} (grå i tabellen) "
                     "i »Øvrige«.")
    noter += _benchmarknote(r, "regioner")
    _afslut_noter(s, r, noter)
    return s


# ---------- valutafordeling ----------

# samme valuta har samme farve i alle tre diagrammer; andre valutaer tager BR_FARVER i rækkefølge
VALUTA_FARVER = {"DKK": SAGE, "EUR": TERRA, "USD": TAUPE}
# (nøgle, navn i hullet, undertekst i hullet)
VALUTA_MAAL = [("eksponering", "Eksponering", "NAV + resttilsagn"), ("nav", "NAV", None),
               ("resttilsagn", "Resttilsagn", None)]
# ringene står i en næsten ligesidet trekant midt på sliden: højden er fast (fra Y_TOP til
# noterne), så afstanden mellem de nederste giver siderne. Ligesidet kan den ikke blive, for så
# ville de nederste ringe røre hinanden
VF_GRAF = 2.4             # diagrammets side (tommer); ringen fylder næsten hele
VF_HUL = 56               # hullets diameter i procent af ringens
VF_KANT = 0.02            # luft om ringen i diagrammet, som andel af siden
VF_BASIS = 3.5            # afstanden mellem de nederste ringes midter
VF_STREG_LUFT = 0.07      # luft mellem en forbindelseslinje og ringene


def _valutafarver(vf):
    """{valuta: farve} for alle valutaer i fordelingen, faste farver først."""
    valutaer = [x["valuta"] for x in vf["eksponering"] or vf["nav"] or vf["resttilsagn"]]
    ledige = iter([f for f in BR_FARVER if f not in VALUTA_FARVER.values()])
    return {v: VALUTA_FARVER.get(v) or next(ledige, BR_OEVRIGE) for v in valutaer}


def _fyld_diagramflade(diagram, kant):
    """Lad tegnefladen fylde diagrammet (minus kant på hver side) i stedet for PowerPoints
    automatiske layout, der lader en bred margen stå tom om en ring."""
    plot_area = diagram._chartSpace.find(".//" + qn("c:plotArea"))
    layout = plot_area.find(qn("c:layout"))
    if layout is None:
        layout = plot_area.makeelement(qn("c:layout"), {})
        plot_area.insert(0, layout)
    for barn in list(layout):
        layout.remove(barn)
    manuel = layout.makeelement(qn("c:manualLayout"), {})
    layout.append(manuel)
    for navn, vaerdi in (("layoutTarget", "inner"), ("xMode", "edge"), ("yMode", "edge"),
                         ("x", kant), ("y", kant), ("w", 1 - 2 * kant), ("h", 1 - 2 * kant)):
        manuel.append(manuel.makeelement(qn(f"c:{navn}"), {"val": str(vaerdi)}))


def _lys(farve):
    """Om en farve er så lys, at tekst på den skal være mørk."""
    r, g, b = farve
    return 0.299 * r + 0.587 * g + 0.114 * b > 150


def _ring(s, x, y, side, poster, serie):
    """Ringdiagram (doughnut) med valuta og andel skrevet på ringen; poster er
    [(navn, dkk, andel, farve)]. Returnerer navnene, der er for små til en etiket."""
    cd = CategoryChartData(number_format="#,##0")
    cd.categories = [navn for navn, *_ in poster]
    cd.add_series(serie, [_kt(dkk) for _, dkk, *_ in poster])
    diagram = s.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(x), Inches(y), Inches(side),
                                 Inches(side), cd).chart
    diagram.has_legend = diagram.has_title = False
    diagram.font.name, diagram.font.size = BROED_FONT, Pt(9)
    diagram._chartSpace.find(".//" + qn("c:holeSize")).set("val", str(VF_HUL))
    _fyld_diagramflade(diagram, VF_KANT)
    plot = diagram.plots[0]
    plot.vary_by_categories = True
    uden_etiket = []
    for punkt, (navn, _, andel, farve) in zip(plot.series[0].points, poster):
        punkt.format.fill.solid()
        punkt.format.fill.fore_color.rgb = farve
        punkt.format.line.color.rgb = HVID
        punkt.format.line.width = Pt(1.5)
        if andel < BR_ETIKET:
            uden_etiket.append(navn)
            continue
        para = punkt.data_label.text_frame.paragraphs[0]
        for i, tekst in enumerate((navn, pct(andel))):
            if i:
                para._p.add_br()
            run = para.add_run()
            run.text = tekst
            run.font.name, run.font.size, run.font.bold = BROED_FONT, Pt(9), bool(i)
            run.font.color.rgb = SORT if _lys(farve) else HVID
    return uden_etiket


def _forbindelser(s, midter):
    """Linjer mellem ringene, fra kant til kant: NAV og resttilsagn op til eksponeringen."""
    ring = VF_GRAF * (1 - 2 * VF_KANT) / 2 + VF_STREG_LUFT

    def linje(a, b, luft_a, luft_b):
        """Linjen fra a mod b, startende luft_a fra a og sluttende luft_b før b."""
        (x1, y1), (x2, y2) = a, b
        laengde = ((x2 - x1) ** 2 + (y2 - y1) ** 2) ** 0.5
        ux, uy = (x2 - x1) / laengde, (y2 - y1) / laengde
        _streg(s, x1 + ux * luft_a, y1 + uy * luft_a, x2 - ux * luft_b, y2 - uy * luft_b,
               farve=SAGE_MID, pt=1.25)

    nav, rest, eks = midter["nav"], midter["resttilsagn"], midter["eksponering"]
    linje(nav, eks, ring, ring)
    linje(rest, eks, ring, ring)


def valutafordeling(prs, r):
    """Tre ringdiagrammer i en trekant: eksponering øverst, NAV og resttilsagn under, hver
    fordelt på fondenes valuta i DKK, med målets navn og total i hullet. None uden DKK-tal."""
    vf = r.get("valutafordeling")
    if not vf:
        return None
    s = _oversigtsslide(prs, r, "Eksponering", enhed="beløb i tusinde DKK")
    farver = _valutafarver(vf)
    fod = Fodnoter("Kilde: Forvalterrapporter, Danmarks Nationalbank og egne beregninger.")
    fod.generel(_kursnote(r))
    hul_m = {"eksponering": fod.maerke(
                 "Hver fond indgår med hele sit beløb i fondens valuta; valutaen i fondenes "
                 "underliggende investeringer er ikke opgjort."),
             }
    mx, bund = (X0 + X1) / 2, Y_NOTER - 0.1 - VF_GRAF / 2
    midter = {"eksponering": (mx, Y_TOP + VF_GRAF / 2),
              "nav": (mx - VF_BASIS / 2, bund), "resttilsagn": (mx + VF_BASIS / 2, bund)}
    _forbindelser(s, midter)
    uden_etiket = set()
    for maal, navn, under in VALUTA_MAAL:
        mx, my = midter[maal]
        poster = vf[maal]
        if poster:
            uden_etiket |= set(_ring(
                s, mx - VF_GRAF / 2, my - VF_GRAF / 2, VF_GRAF,
                [(x["valuta"], x["dkk"], x["andel"], farver[x["valuta"]]) for x in poster],
                serie=f"{navn}, tDKK"))
        # i hullet: målets navn, eventuelt hvad det består af, og totalen
        i_alt = tusind(sum(x["dkk"] for x in poster))
        hul_w = VF_GRAF * (1 - 2 * VF_KANT) * VF_HUL / 100 * 0.92
        _tekst(s, mx - hul_w / 2, my - 0.32, hul_w, 0.3, navn + hul_m.get(maal, ""), 12,
               fed=True, farve=SAGE, just=C, anker=MSO_ANCHOR.BOTTOM)
        if under:
            _tekst(s, mx - hul_w / 2, my - 0.03, hul_w, 0.2, under, 8, farve=SAGE, just=C)
        _tekst(s, mx - hul_w / 2, my + (0.15 if under else -0.01), hul_w, 0.22,
               f"{i_alt} tDKK", 9.5, just=C)

    if uden_etiket:
        fod.generel(f"{_navne(sorted(uden_etiket))} er under {pct(BR_ETIKET, 0)} og har ingen "
                    "etiket i ringene.")
    fod.generel(_datonoter(r))
    fod.tegn(s)
    return s


# ---------- udvikling i kvartalet og året ----------

PERIODER_TEKST = {"kvartal": "kvartalet", "aar": "år til dato"}
VB_MAKS = 8               # linjer pr. halvdel på siden med værdibevægelser


def _viste_perioder(r):
    """Perioderne på udviklingssiderne; i første kvartal er kvartalet og året det samme."""
    pr = r["perioder"]
    return ["kvartal"] if pr["kvartal"] == pr["aar"] else ["kvartal", "aar"]


def _periodetitel(r, p):
    til = r["perioder"][p][1]
    if p == "kvartal":
        return (f"Kvartalet og år til dato ({kvartal(til)})" if _viste_perioder(r) == ["kvartal"]
                else f"Kvartalet ({kvartal(til)})")
    return f"År til dato ({til.year})"


def _fondsnoter(f, felt):
    """Fondens noter om ét tal (data.Note.felt), til sider på tværs af fondene; mærket står i
    fondens række, så noten nævner ikke fonden."""
    return [n_ for n_ in f.get("noter", []) if getattr(n_, "felt", None) == felt]


def _udviklingsnoter(f):
    """Fondens noter om afvigende primodato, som de står på fondssiden."""
    return _fondsnoter(f, "udvikling")


# hvordan valutakursændringen er regnet (Definitioner siger kun, hvad den er)
UDV_VALUTA = ("Valutakursændringen er forskellen, når NAV primo er omregnet med kursen på "
              "primodatoen, cash flows med kursen på betalingsdagen og værdiændring og NAV "
              "ultimo med kursen på rapportdatoen.")


def _nav_bro(slide, x, y, w, h, trin, gulv, loft, maks_bredde=None):
    """Bro fra NAV primo til ultimo: trin er (label, værdi, farve, er_niveau).

    Niveausøjlerne (primo og ultimo) står på grundlinjen; ændringerne svæver fra niveauet
    før til niveauet efter. Aksen går fra gulv til loft, så ændringerne kan ses ved siden af
    NAV; står gulv over 0, får niveausøjlerne et brud nederst.
    """
    plads = w / len(trin)
    # maks_bredde: få trin i hele sidens bredde giver ellers klodsede søjler
    bredde = min(plads * 0.62, maks_bredde or plads)
    niveau = lambda v: y + h - min(1.0, max(0.0, (v - gulv) / (loft - gulv))) * h
    stand = 0.0
    for i, (label, vaerdi, farve, er_niveau) in enumerate(trin):
        sx = x + i * plads + (plads - bredde) / 2
        v = vaerdi or 0
        if er_niveau:
            top, bund, stand = niveau(v), y + h, v
        else:
            top, bund = niveau(max(stand, stand + v)), niveau(min(stand, stand + v))
            stand += v
        _bjaelke(slide, sx, top, bredde, bund - top, farve)
        if er_niveau and gulv > 0:
            for dy in (0.10, 0.16):
                _streg(slide, sx - 0.03, y + h - dy + 0.03, sx + bredde + 0.03, y + h - dy - 0.03,
                       farve=HVID, pt=2)
        _tekst(slide, sx - 0.2, top - 0.18, bredde + 0.4, 0.16,
               tusind(vaerdi) if er_niveau else aendring(vaerdi), 8, fed=er_niveau, just=C)
        _tekst(slide, x + i * plads, y + h + 0.05, plads, 0.16, label, 7.5, farve=SAGE, just=C)
        if i < len(trin) - 1:
            _streg(slide, sx + bredde, niveau(stand), sx + plads, niveau(stand), farve=STREG,
                   pt=0.75, stiplet=True)
    _streg(slide, x, y + h, x + w, y + h, farve=SAGE, pt=0.5)


def _bro_trin(u):
    return [("NAV primo", u["primo"], SAGE, True),
            ("Kapitalkald", u["kald"], TERRA, False),
            ("Udlodninger", -u["udlodning"], TAUPE, False),
            ("Værdiændring", u["vaerdi"], OLIVEN if u["vaerdi"] >= 0 else ROED, False),
            ("Valutakurs", u["valuta"], SAGE_MID, False),
            ("NAV ultimo", u["ultimo"], SAGE, True)]


def _akse(broer):
    """Fælles gulv og loft for broerne, så de kan sammenlignes.

    Gulvet ligger halvanden gang broernes spænd under det laveste niveau, så primo- og
    ultimosøjlerne stadig fylder noget, men ændringerne ikke drukner i dem.
    """
    niveauer = []
    for trin in broer:
        stand = 0.0
        for _, v, _, er_niveau in trin:
            stand = v if er_niveau else stand + v
            niveauer.append(stand)
    lav, hoej = min(niveauer), max(niveauer)
    spaend = (hoej - lav) or abs(hoej) * 0.1 or 1.0
    return max(0.0, lav - spaend * 1.5), hoej


def _nullinje(x, w, neg, pos):
    """Nullinjens placering i x..x+w, så plads til fald og stigninger følger de største af dem."""
    return x + w * neg / (neg + pos) if neg + pos else x + w / 2


def _bidragsbjaelke(slide, x, w, nul, y, h, vaerdi, skala):
    """Vandret søjle fra nullinjen: grøn til højre, rød til venstre; skala er tommer pr. DKK."""
    if vaerdi:
        b = abs(vaerdi) * skala
        _bjaelke(slide, nul if vaerdi > 0 else nul - b, y, b, h, OLIVEN if vaerdi > 0 else ROED)


def udvikling(prs, r):
    """Porteføljens NAV fra primo til ultimo i kvartalet og året, og hvad hver fond bidrog med."""
    ud = r.get("udvikling") or {}
    perioder = [p for p in _viste_perioder(r) if ud.get(p)]
    if not perioder:
        return None
    s = _oversigtsslide(prs, r, "Udvikling i kvartalet og året", enhed="beløb i tusinde DKK")

    bro_b = 6.25
    fod = Fodnoter("Kilde: Forvalterrapporter, Danmarks Nationalbank og egne beregninger.")
    # i læserækkefølge: afkastet i overskriften, så broens trin fra venstre, så højre side;
    # alle mærker sættes, før noternes længde bestemmer, hvor indholdet skal slutte
    maerker = {"NAV primo": fod.maerke(
                   "Søjlerne for NAV primo og ultimo er afkortet, så ændringerne kan ses; "
                   + ("begge broer har samme skala." if len(perioder) > 1
                      else "aksen starter ikke ved 0.")),
               "Valutakurs": fod.maerke(UDV_VALUTA)}
    pr_fond_m = fod.maerke("Værdiændringen pr. fond indeholder fondens valutakursændring.")
    fond_m = {f["fund_name"]: fod.maerker(_udviklingsnoter(f)) for f in r["fonde"]}
    fod.generel(_datonoter(r))
    noter_y = _noter_y(fod.tekst())
    blok_h = (noter_y - Y_TOP - 0.1) / len(perioder)
    broer = [[(label + maerker.get(label, ""), *rest) for label, *rest in _bro_trin(ud[p])]
             for p in perioder]
    gulv, loft = _akse(broer)
    for i, (p, trin) in enumerate(zip(perioder, broer)):
        u, y = ud[p], Y_TOP + i * blok_h
        _tekst(s, X0, y, bro_b, 0.22, _periodetitel(r, p), 10.5, fed=True, farve=SAGE)
        _tekst(s, X0, y + 0.03, bro_b, 0.2, f"NAV {aendring(u['ultimo'] - u['primo'])} tDKK"
               f"  ·  afkast {pct(u['afkast'])}", 9, fed=True, just=H)
        _streg(s, X0, y + 0.27, X0 + bro_b, y + 0.27)
        _nav_bro(s, X0, y + 0.58, bro_b, blok_h - 0.98, trin, gulv, loft)

    # højre side: værdiændring inkl. valuta pr. fond, en kolonne pr. periode
    hx = X0 + bro_b + 0.45
    hw = X1 - hx
    _tekst(s, hx, Y_TOP, hw, 0.22, "Værdiændring pr. fond" + pr_fond_m, 10.5, fed=True,
           farve=SAGE)
    _streg(s, hx, Y_TOP + 0.27, X1, Y_TOP + 0.27)
    navn_b, tal_b = 1.45, 0.50
    kol_b = (hw - navn_b) / len(perioder)
    y = Y_TOP + 0.36
    for j, p in enumerate(perioder):
        _tekst(s, hx + navn_b + j * kol_b, y, kol_b, 0.18, PERIODER_TEKST[p].capitalize(), 8,
               fed=True, farve=SAGE, just=C)
    y += 0.24
    bidrag = {p: dict(ud[p]["bidrag"]) for p in perioder}
    alle = [v for p in perioder for v in bidrag[p].values()]
    neg, pos = max([-v for v in alle] + [0]), max(alle + [0])
    zone_b = kol_b - tal_b - 0.08
    skala = zone_b / (neg + pos) if neg + pos else 0
    fonde = r["fonde"]                  # i strategiorden som på de øvrige slides
    raekke_h = min(0.42, (noter_y - y - 0.1) / (len(fonde) + 1))
    for i, f in enumerate(fonde + [None]):
        ry = y + i * raekke_h
        if f is None:
            _bjaelke(s, hx - 0.05, ry, hw + 0.05, raekke_h, SAGE_LYS)
        elif i % 2 == 1:
            _bjaelke(s, hx - 0.05, ry, hw + 0.05, raekke_h, CREAM)
        navn = ("I alt" if f is None
                else _uden_selskabsform(f["fund_name"]) + _datomarkering(r, f["fund_name"])
                + fond_m[f["fund_name"]])
        _tekst(s, hx, ry, navn_b - 0.08, raekke_h, navn, 7.5, fed=f is None,
               anker=MSO_ANCHOR.MIDDLE)
        for j, p in enumerate(perioder):
            v = ud[p]["vaerdi"] + ud[p]["valuta"] if f is None else bidrag[p][f["fund_name"]]
            kx = hx + navn_b + j * kol_b
            if f is not None:
                _bidragsbjaelke(s, kx + 0.04, zone_b, _nullinje(kx + 0.04, zone_b, neg, pos),
                                ry + raekke_h * 0.3, raekke_h * 0.4, v, skala)
            _tekst(s, kx + kol_b - tal_b, ry, tal_b - 0.04, raekke_h, aendring(v), 8,
                   fed=f is None, just=H, anker=MSO_ANCHOR.MIDDLE)
    for j in range(len(perioder)):
        nul = _nullinje(hx + navn_b + j * kol_b + 0.04, zone_b, neg, pos)
        _streg(s, nul, y, nul, y + len(fonde) * raekke_h, farve=STREG, pt=0.5)

    fod.tegn(s)
    return s


def _fortegnsnote(r, fund_name, p):
    """Noten til et afkast med modsat fortegn af værdiændringen, eller None (fortegnsafvigelser)."""
    aarsag = next((a for fond, periode, a in r.get("fortegnsafvigelser") or []
                   if fond == fund_name and periode == p), None)
    if aarsag is None:
        return None
    f = next(f for f in r["fonde"] if f["fund_name"] == fund_name)
    # mærket står i fondens række og periodens kolonne, så noten nævner ingen af dem
    v = f["valuta"]
    if aarsag == "valuta":
        faldt = f["udvikling"][p]["afkast"] < 0
        return (f"Fondens værdi {'faldt' if faldt else 'steg'} i {v}, men {v} "
                f"{'steg' if faldt else 'faldt'} over for DKK, så værdiændringen i DKK og "
                "afkastet har modsat fortegn.")
    return ("Afkastet har modsat fortegn af værdiændringen, fordi Modified Dietz-afkastets "
            "kapitalbase (NAV primo plus de dagvægtede cash flows) er negativ efter store "
            "udlodninger i perioden.")


def aendringer_pr_fond(prs, r):
    """Tabel med hver fonds kapitalkald, udlodninger, værdiændring og afkast pr. periode."""
    ud = r.get("udvikling") or {}
    perioder = [p for p in _viste_perioder(r) if ud.get(p)]
    if not perioder:
        return None
    s = _oversigtsslide(prs, r, "Ændringer pr. fond",
                        enhed="beløb i tusinde DKK; afkast i fondens valuta")
    fod = Fodnoter("Kilde: Forvalterrapporter, Danmarks Nationalbank og egne beregninger.")
    vaerdi_m = fod.maerke(
        "Værdiændringen indeholder valutakursændringen for fonde i fremmed valuta, så i "
        "alt-rækken svarer til broerne på siden »Udvikling i kvartalet og året«.")
    afkast_m = fod.maerke("I alt-rækkens afkast er beregnet i DKK.")
    tvpi_m = fod.maerke(f"TVPI-ændringen er for {PERIODER_TEKST[perioder[-1]]}. TVPI falder, når "
                        "fonden kalder kapital, mens den står over 1x, også uden værditab.")
    hoved = ["Fond", "Valuta"]
    for _ in perioder:
        hoved += ["Kapital-\nkald", "Udlod-\nninger", f"Værdi-\nændring{vaerdi_m}",
                  f"Afkast{afkast_m}"]
    hoved += ["NAV", "TVPI", f"TVPI-\nændring{tvpi_m}"]
    raekker = [hoved]
    for f in r["fonde"]:
        raekke = [_uden_selskabsform(f["fund_name"]) + _datomarkering(r, f["fund_name"])
                  + fod.maerker(_udviklingsnoter(f)), f["valuta"]]
        for p in perioder:
            u = f["udvikling"][p]
            vaerdi = None if u.get("vaerdi_dkk") is None else u["vaerdi_dkk"] + u["valuta_dkk"]
            raekke += [tusind(u.get("kald_dkk")), tusind(u.get("udlodning_dkk")),
                       aendring(vaerdi),
                       pct(u["afkast"]) + fod.maerke(_fortegnsnote(r, f["fund_name"], p))]
        raekke += [tusind(f["n"].get("nav_dkk")), gange(f["n"]["tvpi"]),
                   aendring(f["udvikling"][perioder[-1]]["tvpi_aendring"], gange)]
        raekker.append(raekke)
    total = ["I alt", "DKK"]
    for p in perioder:
        u = ud[p]
        total += [tusind(u["kald"]), tusind(u["udlodning"]), aendring(u["vaerdi"] + u["valuta"]),
                  pct(u["afkast"])]
    total += [tusind(ud[perioder[0]]["ultimo"]), gange((r.get("sammendrag") or {}).get("tvpi")),
              ""]
    raekker.append(total)

    bredder = [2.45, 0.55] + [0.85, 0.85, 0.9, 0.62] * len(perioder) + [0.85, 0.6, 0.66]
    just = [V, C] + [H] * (len(bredder) - 2)
    w = X1 - X0
    x_kol = lambda j: X0 + w * sum(bredder[:j]) / sum(bredder)
    til = r["perioder"]["kvartal"][1]
    grupper = [(_periodetitel(r, p), 2 + 4 * i, 6 + 4 * i) for i, p in enumerate(perioder)]
    grupper.append((f"Pr. {til:%d.%m.%Y}", 2 + 4 * len(perioder), len(bredder)))
    for tekst, fra, til_ in grupper:
        _tekst(s, x_kol(fra), Y_TOP, x_kol(til_) - x_kol(fra), 0.2, tekst, 9, fed=True,
               farve=SAGE, just=C)
        _streg(s, x_kol(fra) + 0.06, Y_TOP + 0.24, x_kol(til_) - 0.06, Y_TOP + 0.24)
    tab_y = Y_TOP + 0.32
    # noterne kendes, før tabellen tegnes, så tabellen kan slutte over dem
    fod.generel(_datonoter(r))
    noter_y = _noter_y(fod.tekst())
    raekke_h = max(0.22, min(0.40, (noter_y - tab_y - 0.1) / len(raekker)))
    _tabel(s, X0, tab_y, w, raekker, bredder=bredder, just=just, raekke_h=raekke_h, pt=8.5,
           hoved=True, total=True, zebra=True)
    # lodrette streger mellem kvartalet, året og stillingen pr. dato, fra gruppeoverskriften
    # til tabellens bund
    for _, fra, _ in grupper:
        _streg(s, x_kol(fra), Y_TOP, x_kol(fra), tab_y + len(raekker) * raekke_h, farve=SAGE,
               pt=1.25)
    fod.tegn(s)
    return s


def vaerdibevaegelser(prs, r):
    """De underliggende fonde og selskaber, hvis værdi steg og faldt mest i kvartalet."""
    vb = r.get("vaerdibevaegelser")
    linjer = [x for x in (vb or {}).get("linjer", []) if x.get("kvartal") is not None]
    if not linjer:
        return None
    s = _oversigtsslide(prs, r, "Største værdibevægelser i porteføljen",
                        enhed="investors andel, beløb i tusinde DKK")
    stigninger = sorted((x for x in linjer if x["kvartal"] > 0), key=lambda x: -x["kvartal"])
    fald = sorted((x for x in linjer if x["kvartal"] < 0), key=lambda x: x["kvartal"])
    stigninger, fald = stigninger[:VB_MAKS], fald[:VB_MAKS]
    maks = max([abs(x["kvartal"]) for x in stigninger + fald] + [1])
    vis_aar = "aar" in _viste_perioder(r)

    luft = 0.40
    halv = (X1 - X0 - luft) / 2
    kol = [("Eksponering", 1.70, V), ("Via", 0.95, V),
           (PERIODER_TEKST["kvartal"].capitalize(), 1.72, C)]
    if vis_aar:
        kol.append(("År til dato", 0.62, H))
    skala = halv / sum(b for _, b, _ in kol)
    kol = [(n, b * skala, j) for n, b, j in kol]
    raekke_h = min(0.40, (Y_NOTER - Y_TOP - 0.75) / VB_MAKS)
    tal_b = 0.50
    for nr, (overskrift, raekker, farve) in enumerate(
            (("Største stigninger", stigninger, OLIVEN), ("Største fald", fald, ROED))):
        x0 = X0 + nr * (halv + luft)
        _tekst(s, x0, Y_TOP, halv, 0.22,
               f"{overskrift} i {kvartal(r['perioder']['kvartal'][1])}",
               10.5, fed=True, farve=SAGE)
        _streg(s, x0, Y_TOP + 0.27, x0 + halv, Y_TOP + 0.27)
        y = Y_TOP + 0.36
        x = x0
        for navn, b, j in kol:
            _tekst(s, x, y, b, 0.18, navn, 8, fed=True, farve=SAGE, just=j)
            x += b
        y += 0.26
        if not raekker:
            _tekst(s, x0, y + 0.05, halv, 0.2, "Ingen i kvartalet.", 8.5, kursiv=True)
        for i, linje in enumerate(raekker):
            ry = y + i * raekke_h
            if i % 2 == 1:
                _bjaelke(s, x0 - 0.05, ry, halv + 0.1, raekke_h, CREAM)
            x = x0
            _tekst(s, x, ry, kol[0][1] - 0.06, raekke_h, linje["navn"], 8.5,
                   anker=MSO_ANCHOR.MIDDLE)
            x += kol[0][1]
            _tekst(s, x, ry, kol[1][1] - 0.06, raekke_h,
                   _kort_fondsnavn(linje["fond"]) + _datomarkering(r, linje["fond"]), 7.5,
                   farve=SAGE, anker=MSO_ANCHOR.MIDDLE)
            x += kol[1][1]
            bjaelke_b = kol[2][1] - tal_b - 0.08
            _bjaelke(s, x + 0.04, ry + raekke_h * 0.3,
                     max(0.02, bjaelke_b * abs(linje["kvartal"]) / maks), raekke_h * 0.4, farve)
            _tekst(s, x + kol[2][1] - tal_b, ry, tal_b, raekke_h, aendring(linje["kvartal"]),
                   8.5, fed=True, just=H, anker=MSO_ANCHOR.MIDDLE)
            x += kol[2][1]
            if vis_aar:
                _tekst(s, x, ry, kol[3][1], raekke_h, aendring(linje.get("aar")), 8.5, just=H,
                       anker=MSO_ANCHOR.MIDDLE)

    noter = ["Kilde: Forvalterrapporter, Danmarks Nationalbank og egne beregninger.",
             "Værdibevægelsen er ændringen i dagsværdi + udloddet - kostpris i fondenes "
             "beholdningsoversigter, så nye investeringer og udlodninger ikke tæller med, "
             "ganget med den nuværende ejerandel og omregnet til DKK med kursen på "
             "rapportdatoen. Fondenes egne omkostninger og carried interest indgår ikke."]
    for p, fonde in vb["uaendrede"].items():
        noter.append(f"Ingen nye beholdningstal for {PERIODER_TEKST[p]}: {', '.join(fonde)}.")
    for p, x in vb["usammenlignelige"].items():
        noter.append(f"Kan ikke sammenlignes for {PERIODER_TEKST[p]} (afhændet, omdøbt eller "
                     f"købt før perioden): {', '.join(navn for _, navn in x)}.")
    if vb.get("ejendomsfonde"):
        # kort: noten er i forvejen tre linjer; ejendommenes værdier er før gæld
        noter.append("Ikke med: " + ", ".join(vb["ejendomsfonde"]) + " (før gæld).")
    if vb["uskalerede"]:
        noter.append("Ikke med, da ejerandelen ikke kendes: " + ", ".join(vb["uskalerede"]) + ".")
    if vb["uden_beholdninger"]:
        noter.append("Forvalteren oplyser ikke beholdninger for: "
                     + ", ".join(vb["uden_beholdninger"]) + ".")
    noter += _datonoter(r)
    fod = Fodnoter()
    fod.generel(noter)
    fod.tegn(s)
    return s


# ---------- samlede investeringer (likviditet) ----------

LQ_FLISE_Y, LQ_FLISE_H = 1.12, 0.52       # de fem tal øverst
LQ_KURVE_Y, LQ_KURVE_H = 1.72, 1.90       # J-kurven
LQ_BUND_Y = 3.70                          # overskrifterne over de to søjlediagrammer
LQ_LUFT = 0.45                            # mellem de to søjlediagrammer


def _mio(v):
    """Et beløb i mio. som tal til en diagramserie."""
    return round(v * _SKALA / 1_000_000, 1)


def _diagramakser(diagram, talformat):
    """Vandrette gitterlinjer, ingen lodret akselinje, kategorierne nederst."""
    diagram.font.name, diagram.font.size = BROED_FONT, Pt(8)
    diagram.font.color.rgb = SORT
    y_akse, x_akse = diagram.value_axis, diagram.category_axis
    y_akse.has_major_gridlines = True
    y_akse.major_gridlines.format.line.color.rgb = _rgb(style.GRID)
    y_akse.major_gridlines.format.line.width = Pt(0.75)
    y_akse.format.line.fill.background()
    y_akse.tick_labels.number_format, y_akse.tick_labels.number_format_is_linked = talformat, False
    x_akse.format.line.color.rgb = STREG
    x_akse.has_major_gridlines = False
    x_akse.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    diagram.has_legend = True
    diagram.legend.position = XL_LEGEND_POSITION.TOP
    diagram.legend.include_in_layout = False
    diagram.legend.font.size = Pt(8)


def _farv(serie, farve):
    serie.format.fill.solid()
    serie.format.fill.fore_color.rgb = farve


def _j_kurve(s, x, y, w, h, kvartaler):
    """Kvartalets ind- og udbetalinger som søjler og den akkumulerede netto cash flow som flade
    i ét diagram (et kombinationsdiagram), så det kan redigeres i PowerPoint."""
    cd = CategoryChartData(number_format="#,##0.0")
    cd.categories = [kvartal(q["dato"]) for q in kvartaler]
    cd.add_series("Indbetalinger i kvartalet", [_mio(q["kald"]) for q in kvartaler])
    cd.add_series("Udbetalinger i kvartalet", [_mio(q["udlodning"]) for q in kvartaler])
    cd.add_series("Akkumuleret netto cash flow", [_mio(q["akkumuleret"]) for q in kvartaler])
    diagram = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(x), Inches(y),
                                 Inches(w), Inches(h), cd).chart
    _diagramakser(diagram, '#,##0" mio."')
    soejler = diagram.plots[0]
    soejler.overlap, soejler.gap_width = 100, 30
    for serie, farve in zip(soejler.series, (TERRA, SAGE)):
        _farv(serie, farve)
        serie.invert_if_negative = False     # ellers bliver kaldene (negative) hvide
    # den akkumulerede serie flyttes over i en flade, der deler akserne med søjlerne; fladen
    # står først, så søjlerne tegnes oven på den
    bar = soejler._element
    ser = bar.findall(qn("c:ser"))[-1]
    bar.remove(ser)
    for tag in ("c:invertIfNegative", "c:spPr"):
        for el in ser.findall(qn(tag)):
            ser.remove(el)
    lys, kant = str(SAGE_LYS), str(SAGE)
    # kurven er en linje for sig: fladens egen kant ville også tegne lodret op til nul i
    # enderne, hvor fladen lukkes
    linje = deepcopy(ser)
    # lys salvie ved nul, der toner ud mod bunden, så søjlerne står tydeligt foran den
    ser.find(qn("c:tx")).addnext(parse_xml(
        f'<c:spPr {nsdecls("c", "a")}><a:gradFill rotWithShape="1"><a:gsLst>'
        f'<a:gs pos="0"><a:srgbClr val="{lys}"><a:alpha val="90000"/></a:srgbClr></a:gs>'
        f'<a:gs pos="100000"><a:srgbClr val="{lys}"><a:alpha val="15000"/></a:srgbClr></a:gs>'
        f'</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:ln><a:noFill/></a:ln>'
        f'</c:spPr>'))
    nr = str(len(soejler.series) + 1)
    linje.find(qn("c:idx")).set("val", nr)
    linje.find(qn("c:order")).set("val", nr)
    linje.find(qn("c:tx")).addnext(parse_xml(
        f'<c:spPr {nsdecls("c", "a")}><a:ln w="19050" cap="rnd"><a:solidFill>'
        f'<a:srgbClr val="{kant}"/></a:solidFill><a:round/></a:ln></c:spPr>'))
    linje.find(qn("c:spPr")).addnext(parse_xml(
        f'<c:marker {nsdecls("c")}><c:symbol val="none"/></c:marker>'))
    linje.append(parse_xml(f'<c:smooth {nsdecls("c")} val="0"/>'))
    akser = bar.findall(qn("c:axId"))
    for tag, ser_, foer in (("c:areaChart", ser, True), ("c:lineChart", linje, False)):
        plot = parse_xml(f'<{tag} {nsdecls("c")}><c:grouping val="standard"/>'
                         f'<c:varyColors val="0"/></{tag}>')
        plot.append(ser_)
        for akse in akser:
            plot.append(deepcopy(akse))
        # fladen bag søjlerne, linjen foran dem
        bar.addprevious(plot) if foer else bar.addnext(plot)
    # linjen og fladen er samme serie; kun fladen står i forklaringen
    diagram.legend._element.find(qn("c:legendPos")).addnext(parse_xml(
        f'<c:legendEntry {nsdecls("c")}><c:idx val="{nr}"/><c:delete val="1"/></c:legendEntry>'))
    return diagram


LQ_ETIKET_MIN = 0.08      # et stykke under 8 % af den højeste søjle er for lavt til sin etiket


def _etiket(punkt, tekst, farve, position=XL_LABEL_POSITION.CENTER, fed=False):
    etiket = punkt.data_label
    etiket.position = position
    etiket.text_frame.text = tekst
    run = etiket.text_frame.paragraphs[0].runs[0]
    run.font.size, run.font.bold, run.font.color.rgb = Pt(8), fed, farve


def _stablede_soejler(s, x, y, w, h, kategorier, serier, talformat, totaler=None):
    """Stablede søjler pr. år uden forklaring: hvert stykke bærer sin værdi som etiket.

    serier: [(navn, værdier, farve, etiketter)], nederst først; etiketter er en tekst pr.
    søjle (hvid på de mørke farver) eller None. totaler: en tekst pr. søjle oven på den.
    """
    cd = CategoryChartData(number_format=talformat)
    cd.categories = kategorier
    for navn, vaerdier, *_ in serier:
        cd.add_series(navn, vaerdier)
    if totaler:
        # en usynlig serie på 0 øverst i stablen bærer totalen som etiket ved sin fod
        cd.add_series("I alt", [0] * len(kategorier))
    diagram = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(x), Inches(y),
                                 Inches(w), Inches(h), cd).chart
    _diagramakser(diagram, talformat)
    diagram.has_legend = False
    plot = diagram.plots[0]
    plot.gap_width = 45
    hoejest = max(sum(v) for v in zip(*(vaerdier for _, vaerdier, *_ in serier))) or 1
    for serie, (_, vaerdier, farve, etiketter) in zip(plot.series, serier):
        _farv(serie, farve)
        for punkt, v, tekst in zip(serie.points, vaerdier, etiketter or []):
            if v >= LQ_ETIKET_MIN * hoejest:
                _etiket(punkt, tekst, SORT if farve in KORT_LYSE else HVID)
    if totaler:
        top = plot.series[-1]
        top.format.fill.background()
        for punkt, tekst in zip(top.points, totaler):
            _etiket(punkt, tekst, SORT, XL_LABEL_POSITION.INSIDE_BASE, fed=True)
    return diagram


def likviditet(prs, r, med_ejendomme=False):
    """Samlede investeringer: fem nøgletal, J-kurven af den akkumulerede netto cash flow,
    tilsagn og resttilsagn ved årets udgang og TVPI (DPI + RVPI) ved årets udgang. Uden
    ejendomsfondene (r["likviditet"]), eller med dem (r["likviditet_alle"]). Data fra
    data.likviditet(); None uden DKK-tal."""
    lq = r.get("likviditet_alle" if med_ejendomme else "likviditet")
    if not lq:
        return None
    s = _oversigtsslide(prs, r, "Samlede investeringer",
                        enhed=("hele porteføljen inkl. ejendomsfonde" if med_ejendomme
                               else "uden ejendomsfonde") + ", beløb i tusinde DKK")
    fod = Fodnoter("Kilde: Forvalterrapporter, Danmarks Nationalbank og egne beregninger.")
    fod.generel(f"Ejendomsfondene ({_navne(lq['udeladt'])}) er evergreen og indgår ikke."
                if lq["udeladt"] else None,
                "Ind- og udbetalinger er omregnet med Nationalbankens kurs på deres dato, "
                "tilsagn, resttilsagn og NAV med kursen ved årets udgang.")
    t, sidste = lq["total"], lq["aar"][-1]
    afkast = fod.maerke("Afkast er markedsværdi + udbetalinger - indbetalinger.")
    aar_maerke = (fod.maerke(f"{sidste['aar']} er pr. {sidste['dato']:%d.%m.%Y}.")
                  if (sidste["dato"].month, sidste["dato"].day) != (12, 31) else "")
    if lq.get("foer_start"):
        fod.generel(f"Ind- og udbetalinger før {lq['aar'][0]['aar']} "
                    f"({_navne(lq['foer_start'])}) indgår samlet i den akkumulerede netto cash "
                    f"flow fra start: {tusind(lq['primo'])} tDKK.")
    if lq["foer_nav"]:
        fod.generel(f"Før den første NAV indgår {_navne(lq['foer_nav'])} med det, der er "
                    "betalt ind netto.")
    fod.generel(_datonoter(r))
    noter_y = _noter_y(fod.tekst())

    # de fem tal
    fliser = [("Markedsværdi", t["nav"]), ("Afkast" + afkast, t["afkast"]),
              ("Indbetalinger", t["indbetalt"]), ("Udbetalinger", t["udloddet"]),
              ("Netto cash flow", t["netto"])]
    fw = (X1 - X0) / len(fliser)
    for i, (label, v) in enumerate(fliser):
        _tekst(s, X0 + i * fw, LQ_FLISE_Y, fw, 0.22, label, 11, just=C)
        _tekst(s, X0 + i * fw, LQ_FLISE_Y + 0.22, fw, 0.30, tusind(v), 16, fed=True,
               farve=SAGE, just=C)

    _j_kurve(s, X0, LQ_KURVE_Y, X1 - X0, LQ_KURVE_H, lq["kvartaler"])

    bw = (X1 - X0 - LQ_LUFT) / 2
    by, bh = LQ_BUND_Y + 0.24, noter_y - 0.05 - (LQ_BUND_Y + 0.24)
    kategorier = [str(a["aar"]) for a in lq["aar"]]
    kategorier[-1] += aar_maerke
    kaldt = _KALDT_LABEL[lq["genkaldelige"]]
    _tekst(s, X0, LQ_BUND_Y, bw, 0.22, "Kapitaltilsagn ved årets udgang", 11, fed=True)
    def mio(noegle, farve, navn):
        return (navn, [_mio(a[noegle]) for a in lq["aar"]], farve,
                [million(a[noegle]) + " mio." for a in lq["aar"]])

    _stablede_soejler(s, X0, by, bw, bh, kategorier,
                      [mio("resttilsagn", SAGE, "Resttilsagn"), mio("kaldt", SAGE_LYS, kaldt)],
                      '#,##0" mio."')
    x2 = X0 + bw + LQ_LUFT
    _tekst(s, x2, LQ_BUND_Y, bw, 0.22, "TVPI ved årets udgang", 11, fed=True)
    _stablede_soejler(s, x2, by, bw, bh, kategorier, [
        ("RVPI", [a["rvpi"] or 0 for a in lq["aar"]], SAGE, [gange(a["rvpi"]) for a in lq["aar"]]),
        ("DPI", [a["dpi"] or 0 for a in lq["aar"]], SAGE_LYS, None)],
        '0.00"x"', totaler=[gange(a["tvpi"]) for a in lq["aar"]])
    fod.tegn(s, noter_y)
    return s


# ---------- definitioner ----------

# (begreb, forklaring). Teksten skal følge beregningerne i data.py; ændres en beregning,
# skal forklaringen her følge med.
DEFINITIONER = [
    ("Tilsagn", "Det beløb, investor har forpligtet sig til at indskyde i fonden, i fondens "
                "valuta."),
    ("Indbetalt kapital", "Den kapital, forvalteren har kaldt, og som investor har betalt ind, "
                          "inkl. udligningsrente og rabat på kald."),
    ("Udlodninger", "Den kapital, fonden har betalt tilbage til investor."),
    ("Genkaldelige udlodninger", "Udlodninger, som forvalteren efter fondens vedtægter kan "
                                 "kalde ind igen."),
    ("Resttilsagn", "Tilsagn fratrukket kaldt kapital og tillagt "
                    "genkaldelige udlodninger: det, forvalteren fortsat kan kalde."),
    ("NAV", "Net asset value: investors andel af fondens indre værdi på opgørelsesdatoen, "
            "som forvalteren har opgjort den."),
    ("PICC", "Paid-in to committed capital: indbetalt kapital i forhold til tilsagnet."),
    ("DPI", "Distributions to paid-in: udlodninger i forhold til indbetalt kapital."),
    ("TVPI", "Total value to paid-in: (udlodninger + NAV) i forhold til indbetalt kapital. På "
             "i alt-rækker med flere valutaer er alle beløb omregnet til DKK."),
    ("IRR", "Den årlige interne rente af investors cash flows med NAV som sidste cash flow. "
            "Porteføljens IRR er regnet på alle fondes cash flows samlet, i DKK med dagens "
            "kurs. Kildeskat indgår ikke."),
    ("Vintage", "Det år, fonden begyndte at investere."),
    ("Løbetid", "Fondens løbetid i år efter vedtægterne. Hvert »+1« er en mulig forlængelse "
                "på ét år."),
    ("Ejerandel", "Investors tilsagn i forhold til fondens samlede tilsagn (evergreen: andel "
                  "af fondens NAV). Beholdninger ganges med ejerandelen."),
    ("Andel af tilsagn", "En fonds eller gruppes tilsagn i DKK i forhold til porteføljens "
                         "samlede tilsagn i DKK."),
    ("Beløb i DKK", "Nationalbankens kurs: NAV og tilsagn på rapportdatoen, ind- og "
                    "udbetalinger på deres egen dato."),
    ("Værdiskabelse", "Ændringen i fondens værdi ud over indbetalinger, udlodninger, "
                      "omkostninger og carried interest. Beregnet som residual, så "
                      "værdibroen går op; uden opdeling fra forvalteren efter begge."),
    ("Carried interest", "Forvalterens andel af fondens gevinst, når investorerne har fået "
                         "et aftalt minimumsafkast."),
    ("Gennemlysning", "Fondenes beholdninger af underliggende fonde og selskaber, opgjort som "
                      "investors andel. For en feeder fordeles investors NAV på den "
                      "underliggende fonds positioner efter dagsværdi."),
    ("Værdiændring", "Ændringen i NAV fra forrige kvartals- eller årsultimo, når kapitalkald "
                     "og udlodninger er trukket ud; efter omkostninger og carried interest."),
    ("Valutakursændring", "Den del af ændringen i DKK, der skyldes kursen for fonde i fremmed "
                          "valuta."),
    ("Afkast i perioden", "Modified Dietz: værdiændringen i forhold til NAV primo plus kapital"
                          "kald og udlodninger vægtet med den del af perioden, de var i fonden."),
    ("Værdibevægelse", "Ændringen i en beholdnings dagsværdi + udloddet - kostpris, opgjort som "
                       "investors andel."),
    (DAGGER, "Fonden er ikke opgjort pr. rapportdatoen; den faktiske dato står ved fonden."),
]
DEF_BEGREB_B = 1.35       # bredden af begrebskolonnen i hver halvdel (tommer)
DEF_TEGN = 68             # tegn pr. linje i forklaringen ved 8,5 pt; hellere for lavt end overlap
DEF_BEGREB_TEGN = 18      # tegn pr. linje i begrebet ved 9 pt fed
DEF_LINJE = 0.145         # linjehøjde ved 8,5 pt
DEF_LUFT = 0.13           # luft mellem to begreber


def _def_hoejde(begreb, tekst):
    """Højden af ét begreb: den højeste af begrebet og forklaringen."""
    return max(ceil(len(tekst) / DEF_TEGN), ceil(len(begreb) / DEF_BEGREB_TEGN)) * DEF_LINJE


def _def_halvdele(definitioner):
    """Del listen i to spalter, så den højeste spalte bliver så lav som muligt."""
    hoejder = [_def_hoejde(b, t) + DEF_LUFT for b, t in definitioner]
    i = min(range(len(hoejder) + 1),
            key=lambda k: max(sum(hoejder[:k]), sum(hoejder[k:])))
    return definitioner[:i], definitioner[i:]


def definitioner(prs, r):
    """Slutside med forklaringer på nøgletal og begreber i rapporten, i to spalter."""
    s = _ny_slide(prs, L_TEKST)
    _titel(s, "Definitioner")
    _fjern_tomme_ph(s)
    _tekst(s, X0, Y_UNDERTITEL, X1 - X0, 0.22,
           f"{r['navn']}  ·  nøgletal og begreber i rapporten", 9.5)
    luft_mellem = 0.40
    spalte_b = (X1 - X0 - luft_mellem) / 2
    for nr, spalte in enumerate(_def_halvdele(DEFINITIONER)):
        x = X0 + nr * (spalte_b + luft_mellem)
        y = Y_TOP
        for i, (begreb, tekst) in enumerate(spalte):
            if i:
                _streg(s, x, y - DEF_LUFT / 2, x + spalte_b, y - DEF_LUFT / 2,
                       farve=SAGE_LYS, pt=0.5)
            h = _def_hoejde(begreb, tekst)
            _tekst(s, x, y, DEF_BEGREB_B - 0.1, h, begreb, 9, fed=True, farve=SAGE)
            _tekst(s, x + DEF_BEGREB_B, y, spalte_b - DEF_BEGREB_B, h, tekst, 8.5)
            y += h + DEF_LUFT
    return s


# ---------- GICS-oversigt i appendiks ----------

# hvad hver GICS-sektor dækker, kort og på dansk; navnene står på engelsk som i resten af
# rapporten
GICS_BESKRIVELSER = {
    "ENERGY": "Olie, gas og kul: efterforskning, produktion, raffinering og transport samt "
              "udstyr og services til energisektoren.",
    "MATERIALS": "Kemikalier, byggematerialer, emballage, metaller og minedrift samt papir- og "
                 "skovprodukter.",
    "INDUSTRIALS": "Kapitalgoder som maskiner, byggeri og forsvar, erhvervsservice samt "
                   "transport og logistik.",
    "CONSUMER_DISCRETIONARY": "Varer og tjenester, forbrugerne kan udskyde: biler, holdbare "
                              "forbrugsgoder og tøj, hoteller, restauranter og uddannelse "
                              "samt detailhandel.",
    "CONSUMER_STAPLES": "Daglige fornødenheder: fødevarer, drikkevarer og tobak, husholdnings- "
                        "og plejeprodukter samt dagligvarehandel.",
    "HEALTH_CARE": "Sundhedsudstyr og -ydelser samt lægemidler, bioteknologi og "
                   "life science-værktøjer.",
    "FINANCIALS": "Banker, øvrige finansielle tjenester som betalinger, kapitalforvaltning og "
                  "børser samt forsikring.",
    "INFORMATION_TECHNOLOGY": "Software og IT-services, hardware og elektronisk udstyr samt "
                              "halvledere.",
    "COMMUNICATION_SERVICES": "Telekommunikation samt medier, underholdning og interaktive "
                              "medier som søgemaskiner og sociale netværk.",
    "UTILITIES": "Forsyning af el, gas og vand samt vedvarende og uafhængig elproduktion.",
    "REAL_ESTATE": "Ejendomsselskaber og REITs samt ejendomsadministration og -udvikling.",
}
GICS_KOL = 3
GICS_TEGN = 62            # tegn pr. linje i beskrivelsen ved 7,5 pt i en tredjedel af bredden
GICS_LINJE = 0.125        # linjehøjde ved 7,5 pt
GICS_NAVN_H = 0.22        # sektornavnet
GICS_LUFT = 0.14          # mellem to sektorer i samme spalte


def _gics_hoejde(kode):
    """Højden af én sektor: navn, beskrivelse og én linje pr. branchegruppe."""
    return (GICS_NAVN_H + ceil(len(GICS_BESKRIVELSER[kode]) / GICS_TEGN) * GICS_LINJE
            + len(GICS_STRUKTUR[kode]) * GICS_LINJE + 0.06)


def _gics_spalter(koder, n=GICS_KOL):
    """Sektorerne i GICS-rækkefølge delt i n spalter, så den højeste spalte bliver lavest."""
    if n == 1 or len(koder) <= 1:
        return [koder]
    bedst = None
    for i in range(1, len(koder)):
        resten = _gics_spalter(koder[i:], n - 1)
        spalter = [koder[:i]] + resten
        hoejde = max(sum(_gics_hoejde(k) + GICS_LUFT for k in sp) for sp in spalter)
        if bedst is None or hoejde < bedst[0]:
            bedst = (hoejde, spalter)
    return bedst[1]


def gics_oversigt(prs, r):
    """Appendiks: GICS-sektorerne med en kort beskrivelse og deres branchegrupper, i
    sektorernes faste farver."""
    s = _ny_slide(prs, L_TEKST)
    _titel(s, "GICS-sektorer og branchegrupper")
    _fjern_tomme_ph(s)
    _tekst(s, X0, Y_UNDERTITEL, X1 - X0, 0.22,
           f"{r['navn']}  ·  branchefordelingen i rapporten følger Global Industry "
           "Classification Standard", 9.5)
    luft_mellem = 0.30
    spalte_b = (X1 - X0 - luft_mellem * (GICS_KOL - 1)) / GICS_KOL
    for nr, spalte in enumerate(_gics_spalter(list(GICS_STRUKTUR))):
        x = X0 + nr * (spalte_b + luft_mellem)
        y = Y_TOP
        for kode in spalte:
            h = _gics_hoejde(kode)
            _bjaelke(s, x, y + 0.03, 0.07, h - 0.06, SEKTOR_FARVER.get(kode, BR_UDEN_GICS))
            tx, tw = x + 0.15, spalte_b - 0.15
            _tekst(s, tx, y, tw, GICS_NAVN_H, SEKTORER[kode], 9.5, fed=True, farve=SAGE)
            by = y + GICS_NAVN_H
            bh = ceil(len(GICS_BESKRIVELSER[kode]) / GICS_TEGN) * GICS_LINJE
            _tekst(s, tx, by, tw, bh, GICS_BESKRIVELSER[kode], 7.5)
            gy = by + bh + 0.03
            for gruppe in GICS_STRUKTUR[kode]:
                _tekst(s, tx, gy, tw, GICS_LINJE, f"·  {BRANCHEGRUPPER[gruppe]}", 7.5,
                       farve=SAGE)
                gy += GICS_LINJE
            y += h + GICS_LUFT
    fod = Fodnoter("Kilde: MSCI og S&P Dow Jones Indices, Global Industry Classification "
                   "Standard (GICS). Rapporten fordeler på sektorer; branchegrupperne er "
                   "sektorernes underkategorier. »Diversificeret (fondsandele)« og »Ikke "
                   "klassificeret« er rapportens egne grupper uden for GICS. Farverne er "
                   "sektorernes farver i rapportens branchefordeling.")
    fod.tegn(s)
    return s


# ---------- executive summary ----------

SAM_FLISE_H = 0.80        # højden af nøgletalsfliserne øverst
SAM_VENSTRE_B = 5.75      # bredden af højdepunkterne; fordelingen får resten
SAM_TEGN = 92             # tegn pr. linje i et højdepunkt ved 9,5 pt
SAM_LINJE = 0.17          # linjehøjde ved 9,5 pt
SAM_MAKS_GRUPPER = 5      # flere grupper end det samles i "Øvrige"


def _hoejdepunkter(r, fod=None):
    """Korte sætninger om porteføljen til executive summary; DKK-sætningerne kun med DKK-tal.

    fod (Fodnoter) får noterne om den største eksponering og fordelingerne, med mærket i
    sætningen."""
    fod = fod or Fodnoter()
    sm, fd = r["sammendrag"], r["fordelinger"]
    v = sm["vintage"]
    vintage = "" if not v else (f" med vintage {v[0]}" if v[0] == v[1]
                                else f" med vintage {v[0]}–{v[1]}")
    punkter = [f"Porteføljen består af {sm['antal_fonde']} fonde hos {sm['antal_forvaltere']} "
               f"forvaltere{vintage}."]
    if sm["indbetalt_dkk"] is not None:
        punkter += [
            f"Der er indbetalt {tusind(sm['indbetalt_dkk'])} tDKK, svarende til "
            f"{pct(sm['picc'])} af tilsagnet på {tusind(sm['tilsagn_dkk'])} tDKK, og fondene "
            f"har udloddet {tusind(sm['udloddet_dkk'])} tDKK.",
            f"Den samlede eksponering er {tusind(sm['eksponering_dkk'])} tDKK inklusive NAV, og "
            f"resttilsagnet på {tusind(sm['resttilsagn_dkk'])} tDKK kan stadig kaldes af "
            "forvalterne.",
            f"Porteføljens værdi, udlodninger og NAV, er {gange(sm['tvpi'])} den indbetalte "
            "kapital."]
    e = r.get("stoerste_eksponering")
    if e:
        m = fod.maerke("Eksponeringen er opgjort på forvalternes beholdningsoversigter"
                       + (f"; {_navne(e['udeladt'])} indgår ikke, da der ikke er "
                          "sammenlignelige beholdninger" if e["udeladt"] else "") + ".")
        punkter.append(f"Største enkelteksponering er {e['navn']} ({e['art']} via "
                       f"{e['fond']}) med {tusind(e['dkk'])} tDKK, svarende til "
                       f"{pct(e['andel'])} af NAV{m}.")
    u = (r.get("udvikling") or {}).get("kvartal")
    if u:
        nav = u["ultimo"] - u["primo"]
        punkter.append(
            f"I {kvartal(u['til'])} {'steg' if nav >= 0 else 'faldt'} NAV med "
            f"{tusind(abs(nav))} tDKK til {tusind(u['ultimo'])} tDKK: kapitalkald "
            f"{tusind(u['kald'])} tDKK, udlodninger {tusind(u['udlodning'])} tDKK og "
            f"værdiændring {aendring(u['vaerdi'] + u['valuta'])} tDKK (afkast "
            f"{pct(u['afkast'])}).")
    dele = []
    for noegle, (grupper, maal) in _sammendragsfordelinger(r).items():
        k = VISTE_NAVNE[noegle].lower()
        # segmenterne står i strategiorden, ikke efter størrelse
        med_andel = [g for g in grupper if g["andel"] is not None]
        if med_andel:
            g = max(med_andel, key=lambda g: g["andel"])
            dele.append(f"største {k} er {g['navn']} med {pct(g['andel'])} af {maal}")
    if dele:
        tekst = "; ".join(dele)
        punkter.append(tekst[0].upper() + tekst[1:] + ".")
    if _foraeldede(r):
        punkter.append(f"{DAGGER} Ikke alle fonde er opgjort pr. {r['as_of']:%d.%m.%Y}; se "
                       "noterne.")
    return punkter


# segment er dataens ord for fondens strategi (buyout, venture, private credit …); på
# sliderne hedder det investeringsstrategi, eller strategi, hvor der er smalt
VISTE_NAVNE = {"geografi": "Geografi", "segment": "Investeringsstrategi"}
# hvad fordelingerne på executive summary er målt på, efter geografiens mål
SAM_FORDELINGSNOTE = {
    "NAV": "Geografien er fordelt efter fondenes underliggende investeringer.",
    "tilsagnet": "Geografien er fondenes geografiske fokus."}


def _sammendragsfordelinger(r):
    """{"geografi": (grupper, mål), "segment": (...)} til executive summary.

    Nøglen "segment" er dataens navn; på sliden hedder den investeringsstrategi (VISTE_NAVNE).

    Geografien er de samme regioner som på siden Geografisk eksponering: investors NAV
    gennemlyst til fondenes investeringer. Uden den (en fond mangler DKK-kurs) bruges fondenes
    geografiske fokus efter tilsagn. Segmentet er fondenes strategi efter NAV i DKK, og uden
    DKK-tal efter tilsagn.
    """
    fd = r["fordelinger"]
    geo = ((r["regioner"]["grupper"], "NAV") if r.get("regioner")
           else (fd["geografi"]["grupper"], "tilsagnet"))
    seg = fd["segment"]
    seg = (([g | {"andel": g["andel_nav"]} for g in seg["grupper"]], "NAV")
           if (seg.get("total") or {}).get("andel_nav") else (seg["grupper"], "tilsagnet"))
    return {"geografi": geo, "segment": seg}


def _fordelingsbjaelker(s, x, y, w, overskrift, grupper, maks, raekke_h=0.23, label_b=1.75,
                        maks_grupper=None):
    """Gruppernes andel som vandrette bjælker. Returnerer y under blokken.

    maks er den største andel på sliden, så bjælkerne i begge blokke har samme skala. En
    gruppe, der allerede hedder "Øvrige", lægges i Øvrige-rækken.
    """
    _tekst(s, x, y, w, 0.18, overskrift, 9, fed=True)
    y += 0.24
    raekker = [(g["navn"], g["andel"]) for g in grupper if g["navn"] != "Øvrige"]
    oevrige = [(g["navn"], g["andel"]) for g in grupper if g["navn"] == "Øvrige"]
    maks_grupper = maks_grupper or SAM_MAKS_GRUPPER
    if len(raekker) + bool(oevrige) > maks_grupper:
        oevrige += raekker[maks_grupper - 1:]
        raekker = raekker[:maks_grupper - 1]
    if oevrige:
        raekker.append(("Øvrige", None if any(a is None for _, a in oevrige)
                        else sum(a for _, a in oevrige)))
    pct_b = 0.55
    bjaelke_b = w - label_b - pct_b
    for navn, andel in raekker:
        _tekst(s, x, y, label_b - 0.08, 0.20, navn, 8, anker=MSO_ANCHOR.MIDDLE)
        if andel:
            _bjaelke(s, x + label_b, y + 0.04, max(0.02, bjaelke_b * andel / maks), 0.12,
                     SAGE_MID)
        _tekst(s, x + w - pct_b, y, pct_b, 0.20, pct(andel), 8, just=H, anker=MSO_ANCHOR.MIDDLE)
        y += raekke_h
    return y


def sammendrag(prs, r):
    """Executive summary: nøgletal for hele porteføljen, højdepunkter og fordeling af NAV."""
    s = _oversigtsslide(prs, r, "Executive summary", enhed="beløb i tusinde DKK")
    sm = r["sammendrag"]

    fod = Fodnoter("Kilde: Forvalterrapporter, Danmarks Nationalbank og egne beregninger.")
    dkk = fod.maerke("Porteføljens nøgletal, også IRR, er beregnet på beløbene i DKK.")
    fliser = [(sm["antal_fonde"], "fonde i porteføljen"),
              (tusind(sm["tilsagn_dkk"]), "tilsagn, tDKK"),
              # fliserne er for smalle til nøgletallet på samme linje som labelen
              (tusind(sm["indbetalt_dkk"]), f"indbetalt, tDKK\nPICC {pct(sm['picc'])}"),
              (tusind(sm["resttilsagn_dkk"]), "resttilsagn, tDKK"),
              (tusind(sm["udloddet_dkk"]), f"udloddet, tDKK\nDPI {gange(sm['dpi'])}"),
              (tusind(sm["nav_dkk"]), "NAV, tDKK"),
              (tusind(sm.get("eksponering_dkk")), "eksponering, tDKK\nresttilsagn + NAV"),
              (gange(sm["tvpi"]), f"TVPI{dkk}"),
              (pct(sm.get("irr")), f"IRR{dkk}")]
    luft = 0.14
    flise_b = (X1 - X0 - luft * (len(fliser) - 1)) / len(fliser)
    for i, (tal, label) in enumerate(fliser):
        _fakta_flise(s, X0 + i * (flise_b + luft), Y_TOP, flise_b, SAM_FLISE_H, tal, label)

    y0 = Y_TOP + SAM_FLISE_H + 0.32
    _tekst(s, X0, y0, SAM_VENSTRE_B, 0.22, "Højdepunkter", 10.5, fed=True, farve=SAGE)
    _streg(s, X0, y0 + 0.27, X0 + SAM_VENSTRE_B, y0 + 0.27)
    y = y0 + 0.40
    for punkt in _hoejdepunkter(r, fod):
        h = ceil(len(punkt) / SAM_TEGN) * SAM_LINJE
        _tekst(s, X0, y, 0.15, SAM_LINJE, "•", 9.5, farve=SAGE)
        _tekst(s, X0 + 0.18, y, SAM_VENSTRE_B - 0.18, h, punkt, 9.5)
        y += h + 0.10

    hx = X0 + SAM_VENSTRE_B + 0.45
    hw = X1 - hx
    blokke = _sammendragsfordelinger(r)
    _tekst(s, hx, y0, hw, 0.22, "Fordeling", 10.5, fed=True, farve=SAGE)
    _streg(s, hx, y0 + 0.27, X1, y0 + 0.27)
    y = y0 + 0.40
    maks = max([g["andel"] or 0 for grupper, _ in blokke.values() for g in grupper] + [0.01])
    for noegle, (grupper, maal) in blokke.items():
        overskrift = (f"{VISTE_NAVNE[noegle]} – andel af {maal.removesuffix('et')}"
                      + (fod.maerke(SAM_FORDELINGSNOTE[maal]) if noegle == "geografi" else ""))
        y = _fordelingsbjaelker(s, hx, y, hw, overskrift, grupper, maks)
        y += 0.12

    fod.generel(_kursnote(r), _datonoter(r))
    fod.tegn(s)
    return s


# ---------- sidefod ----------

def sidefod(prs, investor_navn):
    """"Investor | Side n" nederst til højre på hver indholdsslide, i masterens slidenummerfelt.

    python-pptx kopierer ikke layoutets slidenummer-placeholder over på nye slides, så den
    klones her. Forside, skilleslides og slut-slide får ingen sidefod, men tæller med i
    nummereringen, så tallet er det samme som PowerPoints slidenummer. Nummeret er et
    slidenum-felt, så det følger med, hvis nogen flytter eller sletter slides bagefter.
    """
    layout = prs.slide_layouts[L_TEKST]
    ph_layout = next(ph for ph in layout.placeholders
                     if ph.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER)
    for nr, s in enumerate(prs.slides, start=1):
        if s.slide_layout != layout:
            continue
        s.shapes.clone_placeholder(ph_layout)
        ph = next(ph for ph in s.placeholders
                  if ph.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER)
        p = ph.text_frame.paragraphs[0]
        navn = p.add_run()
        navn.text, navn.font.bold = investor_navn, True
        p.add_run().text = " | Side "
        felt = p._p.makeelement(qn("a:fld"), {"id": f"{{{uuid.uuid4()}}}".upper(),
                                              "type": "slidenum"})
        felt.append(p._p.makeelement(qn("a:rPr"), {"lang": "da-DK"}))
        tekst = p._p.makeelement(qn("a:t"), {})
        tekst.text = str(nr)
        felt.append(tekst)
        p._p.append(felt)
