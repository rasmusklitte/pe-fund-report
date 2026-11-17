"""PPIM-slidemasteren fra 2026: layouts, pladsholdere og overgangen fra den gamle geometri.

deck.py er fælles med de andre PE-repos og beskriver den gamle master (ppim_slidemaster.pptx),
så den nye masters layout-indeks bor her. deck.py's hjælpere (_ny_slide, _saet_ph, _ryd_slides,
_fjern_tomme_ph) er uafhængige af masteren og bruges stadig.

slides.py tegner fortsat på den gamle flade, Layout_basic_2 på 11,693 x 6,575 tommer: X0/X1 og
Y_* og alle mål i tommer gælder den. Når et deck er bygget, flytter tilpas() hver figur, som
slides.py har tegnet, over på den nye flade på 13,333 x 7,5 tommer:

- vandret, så den gamle indholdsflade (X0-X1, 0,66-11,04) dækker den nye (0,75-12,58);
- lodret, så undertitlen (Y_UNDERTITEL, 0,86) står under den nye titel, der sidder lavere
  (rubrikken over den), og noternes bund (NOTE_BUND, 6,28) over den nye sidefod.

Den nye flade er 14 % bredere, men kun 3 % højere til indholdet, fordi titlen er rykket ned.
Skrifter, rækkehøjder og linjeafstande skaleres med den lodrette faktor, så tabeller og tekst
fylder det samme i højden som før; billeder (logoer), cirkler og cirkeldiagrammer beholder
deres form og centreres, hvor de stod.
"""

from lxml import etree
from pptx.oxml.ns import qn
from pptx.util import Inches

MASTER_NAVN = "PPIM_slidemaster_2026.pptx"

# Layout-indeks i PPIM_slidemaster_2026.pptx
L_FORSIDE = 0                 # Forside – foto: ph 0 titel, 1 undertitel, 13 mærkat, 14 dato
L_TEKST = 8                   # Kun titel – lys: ph 10 rubrik, 0 titel, 11 manchet, 12 slidenr.
L_SKILLE = [11, 12, 13, 14]   # Skilleside 1-4: ph 10 sektionsnummer, 0 titel, 11 manchet, 12 punkter
L_SLUT = 18                   # Slut-slide fondmægler

LAYOUT_NAVNE = {L_FORSIDE: "Forside – foto", L_TEKST: "Kun titel – lys",
                L_SKILLE[0]: "Skilleside 1", L_SKILLE[1]: "Skilleside 2",
                L_SKILLE[2]: "Skilleside 3", L_SKILLE[3]: "Skilleside 4",
                L_SLUT: "Slut-slide fondmægler"}

PH_RUBRIK = 10                # rubrikken over titlen på L_TEKST (versaler, spærret)
PH_SEKTIONSNUMMER = 10        # det store »01« på skillesiderne

# den gamle flade -> den nye (tommer)
GL_X0, NY_X0 = 0.66, 0.75               # indholdets venstre kant
GL_X1, NY_X1 = 11.04, 12.58             # indholdets højre kant
GL_Y0, NY_Y0 = 0.86, 1.36               # undertitlen: under den nye titel (0,78 + 34 pt)
GL_Y1, NY_Y1 = 6.28, 6.94               # noternes bund: over Petersen & Partners-linjen (7,03)
SX = (NY_X1 - NY_X0) / (GL_X1 - GL_X0)  # 1,140
SY = (NY_Y1 - NY_Y0) / (GL_Y1 - GL_Y0)  # 1,030
SKRIFT = SY                             # skrifter og rækkehøjder følger højden

SIDEFOD_B = 3.5               # sidefodens bredde: »Investor | Side n« er bredere end masterens felt

_EMU = 914400
_FORM = ("ellipse",)                    # figurer, der skal beholde deres form
_RUND_DIAGRAM = ("pieChart", "doughnutChart", "pie3DChart")


def _x(emu):
    return round((NY_X0 + (emu / _EMU - GL_X0) * SX) * _EMU)


def _y(emu):
    return round((NY_Y0 + (emu / _EMU - GL_Y0) * SY) * _EMU)


def _flyt(xfrm, fast_form=False):
    """Flyt og skaler én a:xfrm/p:xfrm fra den gamle flade til den nye."""
    off, ext = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
    if off is None or ext is None:
        return
    x, y, w, h = (int(off.get("x")), int(off.get("y")), int(ext.get("cx")), int(ext.get("cy")))
    if fast_form:
        # samme faktor begge veje; midtpunktet flyttes som alt andet
        mx, my = _x(x + w / 2), _y(y + h / 2)
        nw, nh = round(w * SY), round(h * SY)
        nx, ny = mx - nw // 2, my - nh // 2
    else:
        nx, ny, nw, nh = _x(x), _y(y), round(w * SX), round(h * SY)
    off.set("x", str(nx))
    off.set("y", str(ny))
    ext.set("cx", str(nw))
    ext.set("cy", str(nh))


def _skaler_tekst(rod):
    """Skriftstørrelser og linjeafstande i et XML-træ ganges med SKRIFT."""
    for el in rod.iter(qn("a:rPr"), qn("a:defRPr"), qn("a:endParaRPr")):
        if el.get("sz"):
            el.set("sz", str(max(100, round(int(el.get("sz")) * SKRIFT))))
    for el in rod.iter(qn("a:spcPts")):
        el.set("val", str(round(int(el.get("val")) * SKRIFT)))


def _skaler_indryk(rod):
    """Tekstfelternes og tabelcellernes indvendige margener (EMU)."""
    for el in rod.iter(qn("a:bodyPr")):
        for navn, f in (("lIns", SX), ("rIns", SX), ("tIns", SY), ("bIns", SY)):
            if el.get(navn):
                el.set(navn, str(round(int(el.get(navn)) * f)))
    for el in rod.iter(qn("a:tcPr")):
        for navn, f in (("marL", SX), ("marR", SX), ("marT", SY), ("marB", SY)):
            if el.get(navn):
                el.set(navn, str(round(int(el.get(navn)) * f)))


def _tabel(frame):
    tbl = frame.find(".//" + qn("a:tbl"))
    for kol in tbl.iter(qn("a:gridCol")):
        kol.set("w", str(round(int(kol.get("w")) * SX)))
    for r in tbl.iter(qn("a:tr")):
        r.set("h", str(round(int(r.get("h")) * SY)))


def _diagram(shape):
    """Diagrammets egen XML: skrifter og markørstørrelser. True for cirkel- og ringdiagrammer."""
    cs = shape.chart._chartSpace
    _skaler_tekst(cs)
    for st in cs.iter(qn("c:size")):
        if etree.QName(st.getparent()).localname == "marker":
            st.set("val", str(min(72, max(2, round(int(st.get("val")) * SKRIFT)))))
    plot = cs.find(".//" + qn("c:plotArea"))
    return any(plot.find(qn("c:" + t)) is not None for t in _RUND_DIAGRAM)


def tilpas_slide(slide):
    """Flyt alt, slides.py har tegnet på den gamle flade, over på den nye. Pladsholderne
    (titlen) står allerede, hvor den nye master placerer dem, og røres ikke."""
    for shape in slide.shapes:
        if shape.is_placeholder:
            continue
        el = shape._element
        tag = etree.QName(el).localname
        fast = False
        if tag == "graphicFrame":
            xfrm = el.find(qn("p:xfrm"))
            if shape.has_table:
                _tabel(el)
            elif shape.has_chart:
                fast = _diagram(shape)
        elif tag == "grpSp":
            xfrm = el.find(qn("p:grpSpPr")).find(qn("a:xfrm"))
        else:
            sppr = el.find(qn("p:spPr"))
            xfrm = sppr.find(qn("a:xfrm")) if sppr is not None else None
            geom = sppr.find(qn("a:prstGeom")) if sppr is not None else None
            fast = tag == "pic" or (geom is not None and geom.get("prst") in _FORM)
        if xfrm is not None:
            _flyt(xfrm, fast)
        _skaler_tekst(el)
        _skaler_indryk(el)


def tilpas(prs):
    """Alle indholdsslides i decket (layoutet L_TEKST); kaldes, når slides.py er færdig."""
    layout = prs.slide_layouts[L_TEKST]
    for s in prs.slides:
        if s.slide_layout == layout:
            tilpas_slide(s)


def rubrikker(prs, sektioner):
    """Rubrikken over titlen: sektionens nummer og navn (»01  Porteføljeoversigt«) på hver
    indholdsslide efter en skilleside. Slides før den første skilleside får ingen rubrik.

    sektioner er skillesidernes titler i rækkefølge; nummeret er det samme som på skillesiden.
    """
    tekst_layout = prs.slide_layouts[L_TEKST]
    ph_layout = next(ph for ph in tekst_layout.placeholders
                     if ph.placeholder_format.idx == PH_RUBRIK)
    skilleslides = [prs.slide_layouts[i] for i in L_SKILLE]
    nr = 0
    for s in prs.slides:
        if s.slide_layout in skilleslides:
            nr += 1
            continue
        if s.slide_layout != tekst_layout or not nr:
            continue
        s.shapes.clone_placeholder(ph_layout)
        ph = next(ph for ph in s.placeholders if ph.placeholder_format.idx == PH_RUBRIK)
        ph.text_frame.text = f"{nr:02d}  {sektioner[nr - 1]}"


FORSIDE_B, FORSIDE_H = 4.9, 2.3        # forsidens titelfelt (tommer); bundforankret under logoet
FORSIDE_PT = (58, 28)                  # masterens størrelse og den mindste, vi går ned til
TEGN_EM = 0.52                         # Baskerville Old Face: tegnbredde i em, rundet op
LINJE_EM = 0.9 * 1.2                   # titlens linjeafstand (90 %) i em


def forsidens_skrift(tekst):
    """Største skriftstørrelse, hvor titlen holder sig inden for forsidens titelfelt.

    Feltet er bundforankret, så en titel med for mange linjer vokser opad og ind over
    Petersen & Partners-logoet. PowerPoints autotilpasning virker først, når nogen redigerer
    feltet, så størrelsen sættes her ud fra et forsigtigt skøn over linjebrydningen.
    """
    ord_ = tekst.split()
    for pt in range(FORSIDE_PT[0], FORSIDE_PT[1] - 1, -2):
        pr_linje = int(FORSIDE_B * 72 / (TEGN_EM * pt))
        if max(len(o) for o in ord_) > pr_linje:
            continue
        linjer, laengde = 1, -1
        for o in ord_:
            if laengde + 1 + len(o) > pr_linje:
                linjer, laengde = linjer + 1, len(o)
            else:
                laengde += 1 + len(o)
        if linjer * pt * LINJE_EM / 72 <= FORSIDE_H:
            return pt
    return FORSIDE_PT[1]


def forside(slide, titel, undertitel):
    """Forsidens titel i en størrelse, der holder den fri af logoet, og undertitlen; mærkat
    og dato bruges ikke og fjernes."""
    from pptx.util import Pt

    from .deck import _fjern_tomme_ph, _saet_ph
    ph = _saet_ph(slide, 0, titel)
    for p in ph.text_frame.paragraphs:
        for run in p.runs:
            run.font.size = Pt(forsidens_skrift(titel))
    _saet_ph(slide, 1, undertitel)
    _fjern_tomme_ph(slide)


def skilleside(slide, nr, titel):
    """Skillesidens store nummer og titel; manchet og punkter står tomme og fjernes."""
    from .deck import _fjern_tomme_ph, _saet_ph
    _saet_ph(slide, PH_SEKTIONSNUMMER, f"{nr:02d}")
    _saet_ph(slide, 0, titel)
    _fjern_tomme_ph(slide)


def sidefod_bredde(ph):
    """Masterens slidenummerfelt er 1" bredt; »Investor | Side n« skal have mere, så feltet
    udvides mod venstre og holder sin højre kant."""
    # felterne arves fra layoutet; sættes ét, skal alle fire sættes
    x, y, w, h = ph.left, ph.top, ph.width, ph.height
    ph.left, ph.top = x + w - Inches(SIDEFOD_B), y
    ph.width, ph.height = Inches(SIDEFOD_B), h
