"""Fondssider som redigerbare PowerPoint-objekter (tabeller, tekst, diagram) på PPIM-masteren.

Opbygningen følger "Portfolio Fund Memos" i ATP PEP's kvartalsrapport: fire kolonner øverst
(beskrivelse, fondsidentitet, investors tilsagn, nøgletal) og nederst enten de underliggende
fonde eller - hvor fonden ikke er indlæst på fondsniveau - NAV-historik og værdibro.
"""
import io
from math import ceil

from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from . import style
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

def tusind(v):
    return IKKE_OPLYST if v is None else style._dk(v / 1000, 0)


def million(v):
    """Hele fondens beløb fylder for meget i tusinder - virksomhedskortene viser mio."""
    return IKKE_OPLYST if v is None else style._dk(v / 1_000_000, 0)


def pct(v, dec=1):
    """v er en andel (0,123 -> 12,3 %)."""
    return IKKE_OPLYST if v is None else style._dk(v * 100, dec) + " %"


def gange(v):
    return IKKE_OPLYST if v is None else style.mult(v)


def kvartal(d):
    return f"Q{(d.month - 1) // 3 + 1} {d:%y}"


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
        data.add_series(navn, [round(p[noegle] / 1000) for p in historik])
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


def _beholdningstabel(s, y, beh, f, investor_navn):
    """Tabel over underliggende fonde eller porteføljeselskaber. Returnerer y under tabellen."""
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
        raekker = [["Selskab", "Sektor", "Closing" if har_closing else "Vintage", "Kostpris",
                    "Dagsværdi", "Realiseret", "Total værdi", "TVPI"]]
        raekker += [[r["navn"], r["sektor"] or IKKE_OPLYST,
                     f"{r['closing']:%m.%Y}" if r["closing"] else r["vintage"] or IKKE_OPLYST]
                    + tal(r) for r in r_]
        raekker.append(["I alt", "", ""] + total)
        bredder, just = [2.2, 2.9, 0.9, 1.1, 1.1, 1.1, 1.1, 0.9], [V, V, C, H, H, H, H, H]
    raekke_h = 0.185 if len(raekker) > 8 else 0.215
    _tekst(s, X0, y, X1 - X0, 0.22, f"{titel} – {hvem} (t{f['valuta']})", 10.5, fed=True,
           farve=SAGE)
    _tabel(s, X0, y + 0.28, X1 - X0, raekker, bredder=bredder, just=just, raekke_h=raekke_h,
           pt=8, hoved=True, total=True, zebra=True)
    return y + 0.28 + raekke_h * len(raekker) + 0.2


def _beholdningskilde(beh):
    if not beh["skaleret"]:
        return (f"Porteføljeselskaber pr. {beh['dato']:%d.%m.%Y} er hele fondens tal fra "
                "forvalterens rapport; ejerandelen kan ikke beregnes, da fondens samlede "
                "tilsagn ikke er oplyst.")
    return (f"Fondens beholdninger pr. {beh['dato']:%d.%m.%Y} ganget med ejerandelen. "
            "Kostpris er tilsagn × indbetalt andel, hvor forvalteren ikke oplyser den; "
            "TVPI og IRR pr. fond er forvalterens egne.")


def _historik_og_bro(s, y, f):
    """NAV-graf til venstre og værdibro til højre i den plads, der er tilbage over noterne."""
    bro_x = 7.55
    _tekst(s, X0, y, bro_x - X0 - 0.3, 0.22,
           f"Udvikling i NAV og kapitalstrømme (t{f['valuta']})", 10.5, fed=True, farve=SAGE)
    _historik_graf(s, X0 - 0.1, y + 0.22, bro_x - X0 - 0.2, Y_NOTER - y - 0.27, f["historik"])
    _tekst(s, bro_x, y, X1 - bro_x, 0.22, f"Fra indbetalt kapital til NAV (t{f['valuta']})",
           10.5, fed=True, farve=SAGE)
    bro = [[navn, tusind(v)] for navn, v in f["bro"]]
    raekke_h = min(0.27, (Y_NOTER - y - 0.4) / len(bro))
    _tabel(s, bro_x, y + 0.28, X1 - bro_x, bro, bredder=[2.6, 1], just=[V, H],
           raekke_h=raekke_h, pt=9 if raekke_h >= 0.25 else 8, total=True, zebra=True)


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
    _tekst(s, kol[0][0], Y_TOP, kol[0][1], 0.22, "Kort beskrivelse af fonden", 10.5,
           fed=True, farve=SAGE)
    _streg(s, kol[0][0], Y_TOP + 0.25, kol[0][0] + kol[0][1], Y_TOP + 0.25)
    _tekst(s, kol[0][0], Y_TOP + 0.34, kol[0][1], Y_BUND - Y_TOP - 0.34,
           f["beskrivelse"] or IKKE_OPLYST, 9.5)
    for x, _ in kol[1:]:
        _streg(s, x - 0.155, Y_TOP, x - 0.155, Y_BUND - 0.1, farve=SAGE_LYS)

    _blok(s, *kol[1], "Fondsidentitet", [
        ("Geografi", f["geografi"] or IKKE_OPLYST),
        ("Segment", f["segment"] or IKKE_OPLYST),
        ("Fondsvaluta", f["valuta"]),
        ("Fondsstørrelse", tusind(n["fondsstoerrelse"])),
        ("Vintage", f["vintage"] or IKKE_OPLYST)], bredder=(1, 1.45))
    _blok(s, *kol[2], f"{investor_navn}s tilsagn", [
        ("Tilsagn", tusind(n["tilsagn"])),
        ("Ejerandel", pct(n["andel"], 2)),
        ("Indbetalt kapital", tusind(n["indbetalt"])),
        ("PICC", pct(n["picc"])),
        ("Resttilsagn", tusind(n["resttilsagn"]))])
    _blok(s, *kol[3], "Nøgletal", [
        ("Udlodninger", tusind(n["udloddet"])),
        ("NAV", tusind(n["nav"])),
        ("DPI", gange(n["dpi"])),
        ("TVPI", gange(n["tvpi"])),
        ("IRR", pct(n["irr"]))])

    # nederste halvdel: beholdninger, hvor de findes; historik og værdibro, hvor der er plads
    y = Y_BUND + 0.05
    beh = f["beholdninger"]
    kilder = []
    if beh:
        y = _beholdningstabel(s, y, beh, f, investor_navn)
        kilder.append(_beholdningskilde(beh))
    else:
        kilder.append("Forvalteren oplyser ikke fondens porteføljeselskaber.")
    if Y_NOTER - y >= 1.6:
        _historik_og_bro(s, y, f)
    kilde = " ".join(kilder)

    noter = [kilde] + f["noter"]
    if n["gp_irr"] is not None or n["gp_tvpi"] is not None:
        gp = [t for t in (f"TVPI {gange(n['gp_tvpi'])}" if n["gp_tvpi"] is not None else None,
                          f"IRR {pct(n['gp_irr'] / 100)}" if n["gp_irr"] is not None else None) if t]
        noter.append("Forvalterens egne nøgletal: " + ", ".join(gp) + ".")
    _tekst(s, X0, Y_NOTER, X1 - X0, 0.5, "Kilde: Forvalterrapporter og egne beregninger. "
           + " ".join(noter), 7, kursiv=True)
    return s


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


def _profilkort(s, x, y, w, h, r, info, valuta):
    """Ét kort: navn, sektor, closing, søjler for investering og dagsværdi, og valgfri tekst.

    Søjlerne har deres egen skala pr. kort. Fælles skala ville gøre alle andre kort ulæselige,
    når ét selskab er mange gange større end resten; TVPI under kortet giver forholdet.
    """
    _bjaelke(s, x, y, w, h, CREAM)
    indryk = 0.10
    tw = w - 2 * indryk
    _tekst(s, x + indryk, y + 0.08, tw, 0.20, r["navn"], 10, fed=True, farve=SAGE)
    linje2 = [t for t in (r["sektor"], f"Closing {r['closing']:%m.%Y}" if r["closing"]
                          else (str(r["vintage"]) if r["vintage"] else None)) if t]
    _tekst(s, x + indryk, y + 0.28, tw, 0.30, " · ".join(linje2) or IKKE_OPLYST, 7.5)

    ejer = []
    if info.get("ejerandel") is not None:
        ejer.append(f"Ejerandel {_andel(info['ejerandel'])}")
    if info.get("stemmer") is not None:
        ejer.append(f"stemmer {_andel(info['stemmer'])}")
    if ejer:
        _tekst(s, x + indryk, y + 0.52, tw, 0.16, ", ".join(ejer), 7.5, fed=True)

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
    if not beh or beh["type"] != "selskaber":
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
        _tekst(s, X0, Y_UNDERTITEL, X1 - X0, 0.22,
               f"{f['fund_name']}  ·  {investor_navn}  ·  pr. {beh['dato']:%d.%m.%Y}  ·  "
               f"hele fonden, beløb i mio. {f['valuta']}", 9.5)

        kort_w = (X1 - X0 - KORT_LUFT * (KORT_KOL - 1)) / KORT_KOL
        n_raekker = ceil(len(side) / KORT_KOL)
        # højden følger den længste beskrivelse, men må ikke sprænge fladen over noterne
        plads = (Y_NOTER - Y_TOP - 0.15 - KORT_LUFT * (n_raekker - 1)) / n_raekker
        kort_h = min(plads,
                     _kort_hoejde([(info.get(r["navn"]) or {}).get("beskrivelse") for r in side]))
        for i, r in enumerate(side):
            _profilkort(s, X0 + (i % KORT_KOL) * (kort_w + KORT_LUFT),
                        Y_TOP + (i // KORT_KOL) * (kort_h + KORT_LUFT), kort_w, kort_h, r,
                        info.get(r["navn"]) or {}, f["valuta"])

        noter = ["Kilde: Forvalterrapporter og egne beregninger.",
                 ("Beløbene er hele fondens investering og dagsværdi, ikke investors andel. "
                  "Hvert kort har sin egen søjleskala, så søjlehøjder kan ikke sammenlignes "
                  "på tværs af kortene - det er TVPI, der viser forholdet."),
                 "Ejerandel og stemmeandel er oplyst af forvalteren."]
        _tekst(s, X0, Y_NOTER, X1 - X0, 0.5, " ".join(noter), 7, kursiv=True)
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


def _fondscelle(s, r, f, kol, i, y, raekke_h):
    """Rækkens første seks kolonner: fond, forvalterlogo, type, vintage, løbetid og valuta."""
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
           f["fund_name"] + _datomarkering(r, f["fund_name"]), 9,
           anker=MSO_ANCHOR.MIDDLE)
    if afvigende:
        _tekst(s, _lt_x(0, kol), y + raekke_h - 0.26, kol[0][1], 0.16,
               f"opgjort pr. {afvigende:%d.%m.%Y}", 6.5, farve=TERRA)
    _logo(s, f["logo"], x_midt=_lt_x(1, kol) + kol[1][1] / 2, y_midt=midt,
          maks_b=min(1.05, kol[1][1] - 0.15), maks_h=min(0.34, raekke_h - 0.2))
    for j, vaerdi in ((2, f["segment"]), (3, f["vintage"]), (4, f["loebetid"]),
                      (5, f["valuta"])):
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

    Resttilsagnet er tilsagn - indbetalt + genkaldelige udlodninger. Trækkes de genkaldelige
    udlodninger fra den kaldte kapital i stedet, går tilsagn - kaldt op i resttilsagnet, og
    broen kan tegnes med to trin. Tallet er derfor lavere end "Indbetalt kapital" i
    tabellerne, som er den kapital, der faktisk er betalt ind.
    """
    return (n["indbetalt"] or 0) - (n.get("recallable") or 0)


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
    noter += _datonoter(r)
    _tekst(s, X0, Y_NOTER, X1 - X0, 0.5, " ".join(noter), 7, kursiv=True)
    return s


# Kolonnerne i Tilsagn på tværs af fonde: Lange tilsagn med tal i stedet for vandfaldet
# (valutaen står på linjen over kolonnenavnet, som på fordelingssiderne)
TPT_KOL = [("Fond", 2.15, V), ("Forvalter", 1.10, C), ("Type", 0.85, C), ("Vintage", 0.58, C),
           ("Løbetid", 0.62, C), ("Valuta", 0.45, C), ("Tilsagn", 0.73, C), ("PICC", 0.58, C),
           ("DKK\nResttilsagn", 0.88, C), ("DKK\nNAV", 0.88, C), ("TVPI", 0.78, C),
           ("IRR", 0.78, C)]
TPT_TILSAGN, TPT_PICC, TPT_REST_DKK, TPT_NAV_DKK, TPT_TVPI = 6, 7, 8, 9, 10


def _kursnote(r, beloeb="Resttilsagn og NAV"):
    """Kilden til DKK-omregningen, med kurserne for de fremmede valutaer."""
    kurser = {v: k for v, k in (r.get("kurser") or {}).items() if v != "DKK" and k}
    if not kurser:
        return None
    dele = []
    for v, (kurs, dato) in sorted(kurser.items()):
        dele.append(f"{v} {style._dk(kurs, 4)}"
                    + (f" pr. {dato:%d.%m.%Y}" if dato != r["as_of"] else ""))
    return (f"{beloeb} er omregnet til DKK med Danmarks Nationalbanks kurs pr. "
            f"{r['as_of']:%d.%m.%Y} (DKK pr. enhed: {'; '.join(dele)}).")


def tilsagn_paa_tvaers(prs, r):
    """Som Lange tilsagn, men med tilsagn, PICC, resttilsagn i DKK og NAV i DKK som tal."""
    s = _oversigtsslide(prs, r, "Tilsagn på tværs af fonde",
                        enhed="beløb i tusinde; tilsagn i fondens valuta, "
                              "resttilsagn og NAV i DKK")
    fonde, t, td = r["fonde"], r["total"], r.get("total_dkk")
    n_raekker = len(fonde) + bool(t or td)
    hoved_h = 0.44                                   # to linjer: valuta over kolonnenavn
    raekke_h = min(0.95, (Y_NOTER - Y_TOP - hoved_h - 0.1) / max(n_raekker, 1))
    _kolonnehoved(s, TPT_KOL, hoved_h)

    maks_tvpi = max([f["n"]["tvpi"] or 0 for f in fonde] + [1])
    maks_irr = max([abs(f["n"]["irr"] or 0) for f in fonde] + [0.01])

    def tal(j, y, tekst, fed=False):
        _tekst(s, _lt_x(j, TPT_KOL), y, TPT_KOL[j][1], raekke_h, tekst, 8.5, fed=fed, just=C,
               anker=MSO_ANCHOR.MIDDLE)

    y = Y_TOP + hoved_h
    for i, f in enumerate(fonde):
        n = f["n"]
        _fondscelle(s, r, f, TPT_KOL, i, y, raekke_h)
        for j, tekst in ((TPT_TILSAGN, tusind(n["tilsagn"])), (TPT_PICC, pct(n.get("picc"))),
                         (TPT_REST_DKK, tusind(n.get("resttilsagn_dkk"))),
                         (TPT_NAV_DKK, tusind(n.get("nav_dkk")))):
            tal(j, y, tekst)
        _tvpi_irr_soejler(s, TPT_KOL, TPT_TVPI, y, raekke_h, n, maks_tvpi, maks_irr,
                          grundlinje=True)
        y += raekke_h

    if t or td:
        # DKK-beløbene kan altid lægges sammen; resten kun når fondene har samme valuta
        _bjaelke(s, X0 - 0.05, y, X1 - X0 + 0.1, raekke_h, SAGE_LYS)
        _tekst(s, _lt_x(0, TPT_KOL), y, TPT_KOL[0][1], raekke_h, "I alt", 9, fed=True,
               anker=MSO_ANCHOR.MIDDLE)
        celler = [(5, r["valuta"] or "")]
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
        for j, tekst in celler:
            tal(j, y, tekst, fed=True)

    noter = ["Kilde: Forvalterrapporter, Danmarks Nationalbank og egne beregninger.",
             "PICC er indbetalt kapital i forhold til tilsagnet."]
    kursnote = _kursnote(r)
    if kursnote:
        noter.append(kursnote)
    if td and not t:
        noter.append("Fondene har forskellige valutaer, så i alt-rækken viser beløbene i DKK "
                     "og TVPI beregnet som (udloddet + NAV) / indbetalt, alt omregnet til DKK; "
                     "tilsagn, PICC og IRR vises ikke.")
    if any((f["n"].get("recallable") or 0) for f in fonde):
        noter.append("Resttilsagnet indeholder genkaldelige udlodninger, som forvalteren kan "
                     "kalde igen.")
    noter += _datonoter(r)
    _tekst(s, X0, Y_NOTER, X1 - X0, 0.5, " ".join(noter), 7, kursiv=True)
    return s


def portefoeljeoversigt(prs, r):
    """Én række pr. fond med nøgletal og en samlet række for porteføljen."""
    s = _oversigtsslide(prs, r, "Porteføljeoversigt")
    raekker = [["Fond", "Segment", "Vintage", "Valuta", "Tilsagn", "Indbetalt", "PICC",
                "Udloddet", "NAV", "DPI", "TVPI", "IRR"]]
    for f in r["fonde"]:
        n = f["n"]
        raekker.append([f["fund_name"] + _datomarkering(r, f["fund_name"]),
                        f["segment"] or IKKE_OPLYST, f["vintage"] or IKKE_OPLYST,
                        f["valuta"], tusind(n["tilsagn"]), tusind(n["indbetalt"]), pct(n["picc"]),
                        tusind(n["udloddet"]), tusind(n["nav"]), gange(n["dpi"]),
                        gange(n["tvpi"]), pct(n["irr"])])
    t = r["total"]
    if t:
        raekker.append(["I alt", "", "", r["valuta"], tusind(t["tilsagn"]), tusind(t["indbetalt"]),
                        pct(t["picc"]), tusind(t["udloddet"]), tusind(t["nav"]), gange(t["dpi"]),
                        gange(t["tvpi"]), pct(t["irr"])])
    _tabel(s, X0, Y_TOP, X1 - X0, raekker,
           bredder=[3.0, 1.75, 0.7, 0.65, 0.95, 0.95, 0.8, 0.95, 0.95, 0.7, 0.7, 0.75],
           just=[V, V, C, C, H, H, H, H, H, H, H, H], raekke_h=_raekke_h(len(raekker)), pt=9,
           hoved=True, total=bool(t), zebra=True)
    noter = ["Kilde: Forvalterrapporter og egne beregninger.",
             "IRR er beregnet på investors cash flows og NAV; samlet IRR er poolet over alle fonde."
             if t else "Fondene har forskellige valutaer, så der vises ingen samlet række."]
    noter += _datonoter(r)
    noter += [f"{f['fund_name']}: {n_}" for f in r["fonde"] for n_ in f["noter"]
              if n_.startswith("Indbetalt kapital")]
    _tekst(s, X0, Y_NOTER, X1 - X0, 0.5, " ".join(noter), 7, kursiv=True)
    return s


def _kort_fondsnavn(navn):
    """"K Alternativ Private Equity 2020 K/S" -> "KAPE 2020"; ellers de første to ord."""
    uden = navn.replace(" K/S", "").replace(" L.P.", "").replace(" P/S", "")
    ord_ = uden.split()
    if len(ord_) <= 2:
        return uden
    aar = [o for o in ord_ if o.isdigit() and len(o) == 4]
    forbogstaver = "".join(o[0] for o in ord_ if o[0].isupper() and not o.isdigit())
    return f"{forbogstaver} {aar[0]}" if aar else " ".join(ord_[:2])


def _fakta_flise(slide, x, y, w, h, tal, label):
    """Lille nøgletalsflise: stort tal over en label, på cremet bund."""
    _bjaelke(slide, x, y, w, h, CREAM)
    # tallet står på midten og labelen hænger under den, så de to tilsammen er centreret
    _tekst(slide, x, y, w, h * 0.58, str(tal), 16, fed=True, farve=SAGE, just=C,
           anker=MSO_ANCHOR.BOTTOM)
    _tekst(slide, x, y + h * 0.60, w, h * 0.35, label, 8, just=C)


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
    _tekst(s, X0, y, X1 - X0, 0.22,
           f"Største eksponeringer – {r['navn']}s andel (t{r['valuta']})", 10.5, fed=True,
           farve=SAGE)
    y += 0.28

    # søjlediagram til venstre som rigtigt diagram, så tallene kan redigeres i PowerPoint
    graf_w = 5.55
    data = CategoryChartData(number_format="#,##0")
    data.categories = [x["navn"] for x in reversed(top)]       # største øverst i en bar-graf
    data.add_series("Dagsværdi", [round(x["dagsvaerdi"] / 1000) for x in reversed(top)])
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
    raekker = [["Eksponering", "Via", "Kostpris", "Dagsværdi", "TVPI"]]
    raekker += [[x["navn"], _kort_fondsnavn(x["fond"]), tusind(x["kostpris"]),
                 tusind(x["dagsvaerdi"]), gange(x["tvpi"])] for x in top]
    if oevrige:
        raekker.append([f"Øvrige ({len(oevrige)})", "", tusind(sum(x["kostpris"] or 0 for x in oevrige)),
                        tusind(sum(x["dagsvaerdi"] for x in oevrige)), ""])
    raekke_h = min(0.26, (Y_NOTER - y - 0.08) / len(raekker))
    _tabel(s, tab_x, y, X1 - tab_x, raekker, bredder=[1.85, 1.15, 0.85, 0.95, 0.6],
           just=[V, V, H, H, H], raekke_h=raekke_h, pt=7.5, hoved=True, total=bool(oevrige),
           zebra=True)

    noter = [("Kilde: Forvalterrapporter og egne beregninger. Fondenes beholdninger pr. "
              f"{g['dato']:%d.%m.%Y} ganget med ejerandelen."),
             ("Linjerne blander underliggende fonde og porteføljeselskaber; kolonnen »Via« "
              "viser hvilken fond eksponeringen kommer fra."),
             (f"Summen af beholdningerne ({tusind(g['dagsvaerdi'])} t{r['valuta']}) afviger fra "
              f"fondenes samlede NAV ({tusind(g['nav'])} t{r['valuta']}), fordi likvider, "
              "omkostninger og carried interest på fondsniveau ikke indgår i "
              "forvalternes beholdningsoversigter.")]
    if g["uden_beholdninger"]:
        noter.append("Forvalteren oplyser ikke beholdninger for: "
                     + ", ".join(g["uden_beholdninger"]) + ".")
    noter += _datonoter(r)
    _tekst(s, X0, Y_NOTER, X1 - X0, 0.5, " ".join(noter), 7, kursiv=True)
    return s


FORD_MIN_H = 0.20     # laveste rækkehøjde på fordelingssiderne, før tabellen deles (tommer)


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
    # valutaen står over kolonnenavnet; beløbene er i tusinde, som undertitlen siger
    hoved = ["", "Valuta", "Tilsagn", "DKK\nTilsagn", "DKK\nNAV", "DKK\nResttilsagn", "TVPI",
             "Andel af\ntilsagn"]
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

    noter = ["Kilde: Forvalterrapporter, Danmarks Nationalbank og egne beregninger.",
             "Andel er tilsagnet i DKK i forhold til porteføljens samlede tilsagn i DKK."]
    noter.append("TVPI på i alt-rækkerne er (udloddet + NAV) / indbetalt, alt omregnet til DKK "
                 "med kursen på rapportdatoen.")
    if t["valuta"] is None or any(g["sum"]["valuta"] is None for g in fd["grupper"]):
        noter.append("Tilsagn i fondens valuta vises på i alt-rækkerne kun, hvor fondene har "
                     "samme valuta.")
    kursnote = _kursnote(r, "Fondenes beløb")
    if kursnote:
        noter.append(kursnote)
    if any("foreløbig" in x for f in r["fonde"] for x in f["noter"]):
        noter.append("Geografi og segment er en foreløbig klassifikation.")
    noter += _datonoter(r)

    lavede = []
    for nr, side in enumerate(sider):
        s = _oversigtsslide(prs, r, titel if nr == 0 else f"{titel} (fortsat)",
                            enhed="beløb i tusinde; tilsagn i fondens valuta, øvrige beløb i DKK")
        raekker = [hoved] + [celler for celler, _ in side]
        _tabel(s, X0, Y_TOP, X1 - X0, raekker,
               bredder=[3.3, 0.65, 1.0, 1.1, 1.1, 1.25, 0.75, 0.95],
               just=[V, C, H, H, H, H, H, H], raekke_h=raekke_h, pt=8.5, hoved=True,
               grupper=[i for i, (_, art) in enumerate(side, 1) if art == "gruppe"],
               totaler=[i for i, (_, art) in enumerate(side, 1) if art == "total"])
        _tekst(s, X0, Y_NOTER, X1 - X0, 0.5, " ".join(noter), 7, kursiv=True)
        lavede.append(s)
    return lavede


# ---------- definitioner ----------

# (begreb, forklaring). Teksten skal følge beregningerne i data.py; ændres en beregning,
# skal forklaringen her følge med.
DEFINITIONER = [
    ("Tilsagn", "Det beløb, investor har forpligtet sig til at indskyde i fonden, i fondens "
                "valuta."),
    ("Indbetalt kapital", "Den kapital, forvalteren har kaldt, og som investor har betalt ind. "
                          "Udligningsbeløb ved senere closings indgår ikke."),
    ("Udlodninger", "Den kapital, fonden har betalt tilbage til investor."),
    ("Genkaldelige udlodninger", "Udlodninger, som forvalteren efter fondens vedtægter kan "
                                 "kalde ind igen."),
    ("Resttilsagn", "Tilsagn fratrukket indbetalt kapital og tillagt genkaldelige "
                    "udlodninger: det beløb, forvalteren fortsat kan kalde."),
    ("NAV", "Net asset value: investors andel af fondens indre værdi på opgørelsesdatoen, "
            "som forvalteren har opgjort den."),
    ("PICC", "Paid-in to committed capital: indbetalt kapital i forhold til tilsagnet."),
    ("DPI", "Distributions to paid-in: udlodninger i forhold til indbetalt kapital."),
    ("TVPI", "Total value to paid-in: (udlodninger + NAV) i forhold til indbetalt kapital. På "
             "i alt-rækker med flere valutaer er alle beløb omregnet til DKK."),
    ("IRR", "Den årlige interne rente af investors cash flows med NAV som sidste cash flow. "
            "Porteføljens IRR er beregnet på alle fondes cash flows samlet, ikke som et "
            "gennemsnit. Kildeskat indgår ikke."),
    ("Vintage", "Det år, fonden begyndte at investere."),
    ("Løbetid", "Fondens løbetid i år efter vedtægterne. Hvert »+1« er en mulig forlængelse "
                "på ét år."),
    ("Ejerandel", "Investors tilsagn i forhold til fondens samlede tilsagn. Fondens "
                  "beholdninger ganges med ejerandelen for at vise investors andel."),
    ("Andel af tilsagn", "En fonds eller gruppes tilsagn i DKK i forhold til porteføljens "
                         "samlede tilsagn i DKK."),
    ("Beløb i DKK", "Omregnet med Danmarks Nationalbanks kurs på rapportdatoen eller den "
                    "seneste kurs før."),
    ("Værdiskabelse", "Ændringen i fondens værdi ud over indbetalinger, udlodninger, "
                      "omkostninger og carried interest. Beregnet som residual, så "
                      "værdibroen går op."),
    ("Carried interest", "Forvalterens andel af fondens gevinst, når investorerne har fået "
                         "et aftalt minimumsafkast."),
    ("Gennemlysning", "Fondenes beholdninger af underliggende fonde og selskaber, opgjort som "
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


# ---------- executive summary ----------

SAM_FLISE_H = 0.80        # højden af nøgletalsfliserne øverst
SAM_VENSTRE_B = 5.75      # bredden af højdepunkterne; fordelingen får resten
SAM_TEGN = 92             # tegn pr. linje i et højdepunkt ved 9,5 pt
SAM_LINJE = 0.17          # linjehøjde ved 9,5 pt
SAM_MAKS_GRUPPER = 5      # flere grupper end det samles i "Øvrige"


def _hoejdepunkter(r):
    """Korte sætninger om porteføljen til executive summary; DKK-sætningerne kun med DKK-tal."""
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
            f"Resttilsagnet på {tusind(sm['resttilsagn_dkk'])} tDKK kan endnu kaldes af "
            "forvalterne.",
            f"Porteføljens værdi, udlodninger og NAV, er {gange(sm['tvpi'])} den indbetalte "
            "kapital."]
    if sm["bedst"] and sm["svagest"]:
        b, s = sm["bedst"], sm["svagest"]
        punkter.append(f"Højeste TVPI har {b['fund_name']} med {gange(b['tvpi'])} (IRR "
                       f"{pct(b['irr'])}); laveste har {s['fund_name']} med {gange(s['tvpi'])} "
                       f"(IRR {pct(s['irr'])}).")
    stoerst = {k: fd[k]["grupper"][0] for k in ("geografi", "segment")
               if fd[k]["grupper"] and fd[k]["grupper"][0]["andel"] is not None}
    if stoerst:
        dele = [f"største {k} er {g['navn']} med {pct(g['andel'])} af tilsagnet"
                for k, g in stoerst.items()]
        tekst = "; ".join(dele)
        punkter.append(tekst[0].upper() + tekst[1:] + ".")
    if _foraeldede(r):
        punkter.append(f"{DAGGER} Ikke alle fonde er opgjort pr. {r['as_of']:%d.%m.%Y}; se "
                       "noterne.")
    return punkter


def _fordelingsbjaelker(s, x, y, w, overskrift, grupper, maks):
    """Gruppernes andel af tilsagnet som vandrette bjælker. Returnerer y under blokken.

    maks er den største andel på sliden, så bjælkerne i begge blokke har samme skala.
    """
    _tekst(s, x, y, w, 0.18, overskrift, 9, fed=True)
    y += 0.24
    raekker = [(g["navn"], g["andel"]) for g in grupper]
    if len(raekker) > SAM_MAKS_GRUPPER:
        rest = raekker[SAM_MAKS_GRUPPER - 1:]
        raekker = raekker[:SAM_MAKS_GRUPPER - 1] + [
            ("Øvrige", None if any(a is None for _, a in rest) else sum(a for _, a in rest))]
    label_b, pct_b = 1.75, 0.55
    bjaelke_b = w - label_b - pct_b
    for navn, andel in raekker:
        _tekst(s, x, y, label_b - 0.08, 0.20, navn, 8, anker=MSO_ANCHOR.MIDDLE)
        if andel:
            _bjaelke(s, x + label_b, y + 0.04, max(0.02, bjaelke_b * andel / maks), 0.12,
                     SAGE_MID)
        _tekst(s, x + w - pct_b, y, pct_b, 0.20, pct(andel), 8, just=H, anker=MSO_ANCHOR.MIDDLE)
        y += 0.23
    return y


def sammendrag(prs, r):
    """Executive summary: nøgletal for hele porteføljen, højdepunkter og fordeling af tilsagn."""
    s = _oversigtsslide(prs, r, "Executive summary", enhed="beløb i tusinde DKK")
    sm = r["sammendrag"]

    fliser = [(sm["antal_fonde"], "fonde i porteføljen"),
              (tusind(sm["tilsagn_dkk"]), "tilsagn, tDKK"),
              (tusind(sm["indbetalt_dkk"]), f"indbetalt, tDKK · PICC {pct(sm['picc'])}"),
              (tusind(sm["udloddet_dkk"]), f"udloddet, tDKK · DPI {gange(sm['dpi'])}"),
              (tusind(sm["nav_dkk"]), "NAV, tDKK"),
              (gange(sm["tvpi"]), "TVPI")]
    luft = 0.14
    flise_b = (X1 - X0 - luft * (len(fliser) - 1)) / len(fliser)
    for i, (tal, label) in enumerate(fliser):
        _fakta_flise(s, X0 + i * (flise_b + luft), Y_TOP, flise_b, SAM_FLISE_H, tal, label)

    y0 = Y_TOP + SAM_FLISE_H + 0.32
    _tekst(s, X0, y0, SAM_VENSTRE_B, 0.22, "Højdepunkter", 10.5, fed=True, farve=SAGE)
    _streg(s, X0, y0 + 0.27, X0 + SAM_VENSTRE_B, y0 + 0.27)
    y = y0 + 0.40
    for punkt in _hoejdepunkter(r):
        h = ceil(len(punkt) / SAM_TEGN) * SAM_LINJE
        _tekst(s, X0, y, 0.15, SAM_LINJE, "•", 9.5, farve=SAGE)
        _tekst(s, X0 + 0.18, y, SAM_VENSTRE_B - 0.18, h, punkt, 9.5)
        y += h + 0.10

    hx = X0 + SAM_VENSTRE_B + 0.45
    hw = X1 - hx
    _tekst(s, hx, y0, hw, 0.22, "Fordeling af tilsagn", 10.5, fed=True, farve=SAGE)
    _streg(s, hx, y0 + 0.27, X1, y0 + 0.27)
    y = y0 + 0.40
    blokke = (("geografi", "Geografi"), ("segment", "Segment"))
    maks = max([g["andel"] or 0 for k, _ in blokke for g in r["fordelinger"][k]["grupper"]]
               + [0.01])
    for noegle, overskrift in blokke:
        y = _fordelingsbjaelker(s, hx, y, hw, overskrift, r["fordelinger"][noegle]["grupper"],
                                maks)
        y += 0.12

    noter = ["Kilde: Forvalterrapporter, Danmarks Nationalbank og egne beregninger.",
             "Beløb, PICC, DPI og TVPI for porteføljen er beregnet på beløbene i DKK; IRR er "
             "pr. fond i fondens valuta."]
    kursnote = _kursnote(r, "Fondenes beløb")
    if kursnote:
        noter.append(kursnote)
    noter += _datonoter(r)
    _tekst(s, X0, Y_NOTER, X1 - X0, 0.5, " ".join(noter), 7, kursiv=True)
    return s
