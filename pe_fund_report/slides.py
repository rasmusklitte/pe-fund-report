"""Fondssider som redigerbare PowerPoint-objekter (tabeller, tekst, diagram) på PPIM-masteren.

Opbygningen følger "Portfolio Fund Memos" i ATP PEP's kvartalsrapport: fire kolonner øverst
(beskrivelse, fondsidentitet, investors tilsagn, nøgletal) og nederst enten de underliggende
fonde eller - hvor fonden ikke er indlæst på fondsniveau - NAV-historik og værdibro.
"""
import io

from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
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


# ---------- talformater (dansk) ----------

def tusind(v):
    return IKKE_OPLYST if v is None else style._dk(v / 1000, 0)


def pct(v, dec=1):
    """v er en andel (0,123 -> 12,3 %)."""
    return IKKE_OPLYST if v is None else style._dk(v * 100, dec) + " %"


def gange(v):
    return IKKE_OPLYST if v is None else style.mult(v)


def kvartal(d):
    return f"Q{(d.month - 1) // 3 + 1} {d:%y}"


# ---------- byggeklodser ----------

def _skriv(tf, tekst, pt, fed=False, farve=SORT, just=V, kursiv=False):
    para = tf.paragraphs[0]
    para.alignment = just
    run = para.add_run()
    run.text = str(tekst)
    run.font.name, run.font.size, run.font.bold, run.font.italic = BROED_FONT, Pt(pt), fed, kursiv
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


def _logo(slide, logo, x_hoejre=X1, y_midt=0.47, maks_b=1.55, maks_h=0.36):
    """Forvalterens logo øverst til højre, højrestillet mod indholdskanten.

    Logoer, der kun findes til mørk baggrund, lægges på en mørk plade.
    """
    if not logo:
        return
    luft = 0.07 if logo["moerk_baggrund"] else 0
    skala = min((maks_b - 2 * luft) / logo["bredde"], (maks_h - 2 * luft) / logo["hoejde"])
    b, h = logo["bredde"] * skala, logo["hoejde"] * skala
    x, y = x_hoejre - luft - b, y_midt - h / 2
    if luft:
        plade = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x - luft),
                                       Inches(y - luft), Inches(b + 2 * luft), Inches(h + 2 * luft))
        plade.adjustments[0] = 0.12
        plade.fill.solid()
        plade.fill.fore_color.rgb = SAGE
        plade.line.fill.background()
        plade.shadow.inherit = False
    slide.shapes.add_picture(io.BytesIO(logo["billede"]), Inches(x), Inches(y), Inches(b), Inches(h))


def _streg(slide, x1, y1, x2, y2, farve=SAGE, pt=0.75):
    s = slide.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    s.line.color.rgb = farve
    s.line.width = Pt(pt)
    return s


def _tabel(slide, x, y, w, raekker, bredder, just, raekke_h, pt, hoved=False, total=False,
           zebra=False, grupper=()):
    """Tabel uden PowerPoints tabelstil. raekker er lister af færdigformaterede tekster.

    grupper: rækkeindeks, der er gruppeoverskrifter (fed skrift på cremet bund).
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
        er_hoved, er_total = hoved and i == 0, total and i == sidste
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
            _skriv(celle.text_frame, vaerdi, pt, fed=er_hoved or er_total or er_gruppe,
                   farve=HVID if er_hoved else SORT, just=just[j])
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
                "GP-rapporten; ejerandelen kendes ikke, da fondsstørrelsen mangler.")
    return (f"Fondens beholdninger pr. {beh['dato']:%d.%m.%Y} ganget med ejerandelen. "
            "Kostpris er tilsagn × indbetalt andel, hvor GP ikke oplyser den; "
            "TVPI og IRR pr. fond er GP'ens.")


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
        kilder.append("Fondens porteføljeselskaber er ikke indlæst i databasen.")
    if Y_NOTER - y >= 1.6:
        _historik_og_bro(s, y, f)
    kilde = " ".join(kilder)

    noter = [kilde] + f["noter"]
    if n["gp_irr"] is not None or n["gp_tvpi"] is not None:
        gp = [t for t in (f"TVPI {gange(n['gp_tvpi'])}" if n["gp_tvpi"] is not None else None,
                          f"IRR {pct(n['gp_irr'] / 100)}" if n["gp_irr"] is not None else None) if t]
        noter.append("GP'ens egne nøgletal: " + ", ".join(gp) + ".")
    _tekst(s, X0, Y_NOTER, X1 - X0, 0.5, "Kilde: GP-rapporter og egne beregninger. "
           + " ".join(noter), 7, kursiv=True)
    return s


# ---------- oversigtssider (ATP-rapportens "Portfolio Reporting", side 5-7) ----------

def _oversigtsslide(prs, r, titel):
    s = _ny_slide(prs, L_TEKST)
    _titel(s, titel)
    _fjern_tomme_ph(s)
    enhed = f"beløb i t{r['valuta']}" if r["valuta"] else "beløb i tusinde, fondens valuta"
    _tekst(s, X0, Y_UNDERTITEL, X1 - X0, 0.22,
           f"{r['navn']}  ·  pr. {r['as_of']:%d.%m.%Y}  ·  {enhed}", 9.5)
    return s


def _raekke_h(antal, maks=0.30, mindst=0.19):
    """Rækkehøjde, så tabellen fylder fladen mellem undertitel og noter uden at løbe over."""
    return max(mindst, min(maks, (Y_NOTER - Y_TOP - 0.15) / antal))


def portefoeljeoversigt(prs, r):
    """Én række pr. fond med nøgletal og en samlet række for porteføljen."""
    s = _oversigtsslide(prs, r, "Porteføljeoversigt")
    raekker = [["Fond", "Segment", "Vintage", "Valuta", "Tilsagn", "Indbetalt", "PICC",
                "Udloddet", "NAV", "DPI", "TVPI", "IRR"]]
    for f in r["fonde"]:
        n = f["n"]
        raekker.append([f["fund_name"], f["segment"] or IKKE_OPLYST, f["vintage"] or IKKE_OPLYST,
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
    noter = ["Kilde: GP-rapporter og egne beregninger.",
             "IRR er beregnet på investors cash flows og NAV; samlet IRR er poolet over alle fonde."
             if t else "Fondene har forskellige valutaer, så der vises ingen samlet række."]
    noter += [f"{f['fund_name']}: {n_}" for f in r["fonde"] for n_ in f["noter"]
              if n_.startswith(("Indbetalt kapital er", "Seneste NAV"))]
    _tekst(s, X0, Y_NOTER, X1 - X0, 0.5, " ".join(noter), 7, kursiv=True)
    return s


def fordeling(prs, r, noegle, titel):
    """Fondene grupperet på geografi eller segment, med gruppens andel af indbetalt kapital."""
    s = _oversigtsslide(prs, r, titel)
    grupper = {}
    for f in r["fonde"]:
        grupper.setdefault(f[noegle] or "Ikke oplyst", []).append(f)
    sum_ = lambda fonde, k: sum(f["n"][k] or 0 for f in fonde)
    samme_valuta = r["total"] is not None
    i_alt = r["total"]["indbetalt"] if samme_valuta else None

    raekker, gruppe_idx = [["", "Valuta", "Tilsagn", "Indbetalt kapital", "NAV", "TVPI"]], []
    for navn, fonde in sorted(grupper.items(), key=lambda g: -sum_(g[1], "indbetalt")):
        gruppe_idx.append(len(raekker))
        if samme_valuta:
            ind = sum_(fonde, "indbetalt")
            raekker.append([f"{navn} ({pct(ind / i_alt) if i_alt else IKKE_OPLYST})", "",
                            tusind(sum_(fonde, "tilsagn")), tusind(ind), tusind(sum_(fonde, "nav")),
                            gange((sum_(fonde, "udloddet") + sum_(fonde, "nav")) / ind) if ind
                            else IKKE_OPLYST])
        else:
            raekker.append([navn, "", "", "", "", ""])
        for f in fonde:
            n = f["n"]
            raekker.append(["      " + f["fund_name"], f["valuta"], tusind(n["tilsagn"]),
                            tusind(n["indbetalt"]), tusind(n["nav"]), gange(n["tvpi"])])
    if samme_valuta:
        t = r["total"]
        raekker.append(["I alt", r["valuta"], tusind(t["tilsagn"]), tusind(t["indbetalt"]),
                        tusind(t["nav"]), gange(t["tvpi"])])
    _tabel(s, X0, Y_TOP, X1 - X0, raekker, bredder=[4.6, 0.9, 1.3, 1.5, 1.3, 0.9],
           just=[V, C, H, H, H, H], raekke_h=_raekke_h(len(raekker)), pt=9, hoved=True,
           total=samme_valuta, grupper=gruppe_idx)
    note = "Kilde: GP-rapporter og egne beregninger. Procenten er gruppens andel af indbetalt kapital."
    if any("udkast" in x for f in r["fonde"] for x in f["noter"]):
        note += " Geografi og segment er et udkast og ikke gennemgået."
    _tekst(s, X0, Y_NOTER, X1 - X0, 0.5, note, 7, kursiv=True)
    return s
