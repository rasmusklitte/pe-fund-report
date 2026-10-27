"""Formater og tabeller i slides.py, bygget på den rigtige slidemaster."""
import io
import re
from datetime import date
from itertools import pairwise

import pytest
from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches
from test_data import LQ_KURSER, _lq_fond

from pe_fund_report import data, slides
from pe_fund_report.deck import L_FORSIDE, L_SKILLE, L_SLUT, L_TEKST, _ny_slide, _ryd_slides
from pe_fund_report.paths import MASTER


def test_danske_talformater():
    assert slides.beloeb(1_234_567) == "1,2"
    assert slides.pct(0.123) == "12,3 %"
    assert slides.gange(1.5) == "1,50x"
    assert slides.beloeb(None) == slides.pct(None) == slides.gange(None) == "n/a"
    assert slides.kvartal(date(2026, 6, 30)) == "Q2 26"


def _selskab(navn, closing, vintage):
    return dict(navn=navn, sektor="Software", closing=closing, vintage=vintage, kostpris=1000.0,
                dagsvaerdi=1500.0, udloddet=0.0, total=1500.0, tvpi=1.5, irr=None)


def _tabel(raekker):
    prs = Presentation(str(MASTER))
    _ryd_slides(prs)
    s = _ny_slide(prs, L_TEKST)
    beh = dict(dato=date(2026, 6, 30), raekker=raekker, skaleret=True, type="selskaber")
    slides._beholdningstabel(s, 3.0, beh, dict(valuta="DKK"), "Investor A/S")
    tabel = next(sh.table for sh in s.shapes if sh.has_table)
    return [[c.text for c in r.cells] for r in tabel.rows]


def test_selskabstabel_viser_closing_naar_datoen_findes():
    r = _tabel([_selskab("A", date(2023, 1, 10), None)])
    assert r[0][2] == "Closing" and r[1][2] == "01.2023"


def test_selskabstabel_falder_tilbage_paa_vintage():
    r = _tabel([_selskab("A", None, 2022), _selskab("B", None, 2024)])
    assert r[0][2] == "Vintage" and [x[2] for x in r[1:3]] == ["2022", "2024"]
    assert r[0][-1] == "TVPI"


def _prs():
    prs = Presentation(str(MASTER))
    _ryd_slides(prs)
    return prs


def _tekster(slide):
    return [sh.text_frame.text for sh in slide.shapes if sh.has_text_frame]


def _n(tilsagn, indbetalt, resttilsagn, tvpi, irr, recallable=0.0):
    return dict(tilsagn=tilsagn, indbetalt=indbetalt, resttilsagn=resttilsagn, tvpi=tvpi,
                irr=irr, recallable=recallable)


def _rapport(fonde, total=None, foraeldede=(), skaeve=()):
    return dict(navn="Investor A/S", as_of=date(2026, 6, 30), valuta="DKK", fonde=fonde,
                total=total,
                datotjek=dict(rapport_dato=date(2026, 6, 30),
                              ens=not foraeldede and not skaeve,
                              foraeldede=list(foraeldede), skaeve_beholdninger=list(skaeve)))


def _lt_fond(navn, tilsagn, loebetid=None, benchmark=None, valuta="DKK"):
    return dict(fund_name=navn, logo=None, segment="Buyout", vintage=2021, loebetid=loebetid,
                benchmark=benchmark, valuta=valuta,
                n=_n(tilsagn, tilsagn / 2, tilsagn / 2, 1.5, 0.12))


def test_lange_tilsagn_viser_en_raekke_pr_fond_og_en_totalraekke():
    prs = _prs()
    r = _rapport([_lt_fond("Fond A", 100.0, "10+1+1"), _lt_fond("Fond B", 50.0)],
                 total=_n(150.0, 75.0, 75.0, 1.5, 0.12))
    t = _tekster(slides.lange_tilsagn(prs, r))
    assert "Fond A" in t and "Fond B" in t and "I alt" in t
    assert "10+1+1" in t                      # løbetid fra fund_info.yaml
    assert t.count(slides.IKKE_OPLYST) == 1   # Fond B mangler løbetid


def test_lange_tilsagn_uden_total_springer_totalraekken_over():
    prs = _prs()
    r = _rapport([_lt_fond("Fond A", 100.0)], total=None)
    assert "I alt" not in _tekster(slides.lange_tilsagn(prs, r))


def test_soejler_skalerer_mod_faelles_maksimum():
    prs = _prs()
    s = _ny_slide(prs, L_TEKST)
    foer = len(s.shapes)
    slides._soejler(s, 1.0, 1.0, 3.0, 1.0,
                    [("Hel", 100.0, slides.SAGE), ("Halv", 50.0, slides.TERRA),
                     ("Ingen", 0.0, slides.TAUPE)], maks=100.0)
    hoejder = [sh.height for sh in list(s.shapes)[foer:] if sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE]
    assert hoejder[0] == Inches(1.0)
    assert hoejder[1] == pytest.approx(Inches(0.5), rel=1e-6)
    assert hoejder[2] < Inches(0.01)          # nul tegnes som en hårfin streg, ikke en søjle


def test_virksomhedsprofiler_springer_fond_af_fonde_over():
    prs = _prs()
    f = dict(beholdninger=dict(dato=date(2026, 6, 30), raekker=[], skaleret=True, type="fonde"),
             virksomheder={}, fund_name="FoF", logo=None, valuta="DKK")
    assert slides.virksomhedsprofiler(prs, f, "Investor A/S") == []
    f["beholdninger"] = None
    assert slides.virksomhedsprofiler(prs, f, "Investor A/S") == []


def test_virksomhedsprofiler_tager_otte_selskaber_pr_slide():
    raekker = [dict(navn=f"Selskab {i}", sektor="Software", closing=None, vintage=2022,
                    kostpris_fond=10e6, dagsvaerdi_fond=20e6, tvpi=2.0) for i in range(9)]
    f = dict(beholdninger=dict(dato=date(2026, 6, 30), raekker=raekker, skaleret=True,
                               type="selskaber"),
             virksomheder={}, fund_name="Fond A", logo=None, valuta="DKK")
    sider = slides.virksomhedsprofiler(_prs(), f, "Investor A/S")
    assert len(sider) == 2
    assert "Porteføljeselskaber (1:2)" in _tekster(sider[0])
    assert "Selskab 7" in _tekster(sider[0])      # otte kort på første slide
    assert "Selskab 8" in _tekster(sider[1])


def _bro_gaar_op(trin):
    """Hvert trins bro-niveau skal ramme den næste søjles kant, ellers hopper den stiplede linje."""
    for (_, _, bund, ud, _), (_, naeste_v, naeste_bund, _, _) in pairwise(trin):
        kanter = (naeste_bund, naeste_bund + abs(naeste_v))
        assert min(kanter) - 1e-9 <= ud <= max(kanter) + 1e-9, (ud, kanter)
        assert bund >= -1e-9


def test_tilsagnsbro_netter_genkaldelige_udlodninger_ind_i_kaldt():
    # CataCap: 37.500 tilsagn - (20.512 kaldt - 1.026 genkaldt) = 18.014 resttilsagn
    n = _n(37500.0, 20512.0, 18014.0, 1.31, 0.174, recallable=1026.0)
    trin = slides._tilsagnsbro(n)
    assert [t[0] for t in trin] == ["Tilsagn", "Kaldt, netto", "Resttilsagn"]
    assert slides.kaldt_netto(n) == pytest.approx(19486.0)
    assert trin[1][1] == pytest.approx(19486.0)
    assert trin[-1][3] == pytest.approx(18014.0)      # broen ender i resttilsagnet
    _bro_gaar_op(trin)


def test_tilsagnsbro_bruger_kaldt_inkl_rabat():
    # PPPC2: 10.000 tilsagn - 3.200 kaldt (3.125 betalt + 75 rabat) = 6.800 resttilsagn
    n = dict(_n(10000.0, 3125.0, 6800.0, 1.11, 0.047), kaldt=3200.0)
    trin = slides._tilsagnsbro(n)
    assert trin[1][1] == pytest.approx(3200.0) and trin[-1][3] == pytest.approx(6800.0)
    _bro_gaar_op(trin)


def test_tilsagnsbro_uden_genkaldelige_udlodninger_bruger_bruttotallet():
    n = _n(25000.0, 21530.0, 3470.0, 1.37, 0.116)
    trin = slides._tilsagnsbro(n)
    assert [t[0] for t in trin] == ["Tilsagn", "Kaldt", "Resttilsagn"]
    assert trin[1][1] == pytest.approx(21530.0)       # intet at trække fra
    _bro_gaar_op(trin)


def test_lange_tilsagn_markerer_fond_med_afvigende_opgoerelsesdato():
    r = _rapport([_lt_fond("Fond A", 100.0), _lt_fond("Seed", 50.0)],
                 total=_n(150.0, 75.0, 75.0, 1.5, 0.12),
                 foraeldede=[dict(fond="Seed", dato=date(2025, 9, 30), kvartaler=3)])
    t = _tekster(slides.lange_tilsagn(_prs(), r))
    assert f"Seed {slides.DAGGER}" in t          # dagger i selve rækken
    assert "Fond A" in t                         # fonden uden afvigelse står urørt
    assert "opgjort pr. 30.09.2025" in t
    assert any("Totaler og fordelinger blander" in x for x in t)


def test_lange_tilsagn_uden_datoafvigelser_har_ingen_markering():
    r = _rapport([_lt_fond("Fond A", 100.0)], total=_n(100.0, 50.0, 50.0, 1.5, 0.12))
    t = _tekster(slides.lange_tilsagn(_prs(), r))
    assert "Fond A" in t
    assert not any(slides.DAGGER in x for x in t)


def _lt_fond_valuta(navn, tilsagn, valuta):
    return _lt_fond(navn, tilsagn, valuta=valuta)


def test_lange_tilsagn_deler_ikke_skala_paa_tvaers_af_valutaer():
    # en lille EUR-fond må ikke tegnes kort, fordi en stor DKK-fond sætter skalaen
    blandet = _rapport([_lt_fond_valuta("DKK-fond", 100.0, "DKK"),
                        _lt_fond_valuta("EUR-fond", 10.0, "EUR")])
    blandet["valuta"] = None
    t = _tekster(slides.lange_tilsagn(_prs(), blandet))
    assert any("forskellige valutaer" in x for x in t)
    assert not any("Alle fonde har samme skala" in x for x in t)

    ens = _rapport([_lt_fond("A", 100.0), _lt_fond("B", 10.0)])
    t = _tekster(slides.lange_tilsagn(_prs(), ens))
    assert any("Alle fonde har samme skala" in x for x in t)
    assert not any("forskellige valutaer" in x for x in t)


def test_lange_tilsagn_bruger_samme_kaldt_label_paa_hele_sliden():
    # én fond med genkaldelige udlodninger netter hele sliden, så søjlerne er sammenlignelige
    med = _lt_fond("Med", 100.0)
    med["n"] = _n(100.0, 60.0, 50.0, 1.5, 0.12, recallable=10.0)
    uden = _lt_fond("Uden", 80.0)
    uden["n"] = _n(80.0, 40.0, 40.0, 1.5, 0.12)
    t = _tekster(slides.lange_tilsagn(_prs(), _rapport([med, uden])))
    assert t.count("Kaldt, netto") == 2 and "Kaldt" not in t
    assert any("Genkaldelige udlodninger er trukket fra" in x for x in t)


def test_lange_tilsagn_uden_genkaldelige_udlodninger_skriver_bare_kaldt():
    r = _rapport([_lt_fond("A", 100.0)])
    t = _tekster(slides.lange_tilsagn(_prs(), r))
    assert "Kaldt" in t and "Kaldt, netto" not in t
    assert not any("Genkaldelige udlodninger" in x for x in t)


def test_lange_tilsagn_viser_valuta_pr_fond():
    dkk, eur = _lt_fond_valuta("DKK-fond", 100.0, "DKK"), _lt_fond_valuta("EUR-fond", 10.0, "EUR")
    blandet = _rapport([dkk, eur])
    blandet["valuta"] = None
    t = _tekster(slides.lange_tilsagn(_prs(), blandet))
    assert "DKK" in t and "EUR" in t          # valutaen står i selve rækken, ikke kun i noten


def test_lange_tilsagn_totalraekke_viser_portefoeljens_valuta():
    r = _rapport([_lt_fond("A", 100.0)], total=_n(100.0, 50.0, 50.0, 1.5, 0.12))
    t = _tekster(slides.lange_tilsagn(_prs(), r))
    assert t.count("DKK") == 2               # én pr. fond og én på i alt-rækken


def _logo_billede(bredde=2, hoejde=1):
    """Et rigtigt lille PNG; kun målene betyder noget for placeringen."""
    from PIL import Image
    buf = io.BytesIO()
    Image.new("RGB", (bredde, hoejde), "white").save(buf, format="PNG")
    return dict(billede=buf.getvalue(), bredde=bredde, hoejde=hoejde, moerk_baggrund=False)


def _logo_x(slide):
    pic = next(sh for sh in slide.shapes if sh.shape_type == MSO_SHAPE_TYPE.PICTURE)
    return Emu(pic.left).inches, Emu(pic.width).inches


def test_logo_centreres_om_x_midt():
    s = _ny_slide(_prs(), L_TEKST)
    slides._logo(s, _logo_billede(), x_midt=5.0, y_midt=3.0, maks_b=1.0, maks_h=1.0)
    venstre, bredde = _logo_x(s)
    assert venstre + bredde / 2 == pytest.approx(5.0, abs=0.01)


def test_logo_hoejrestilles_naar_x_midt_ikke_er_sat():
    s = _ny_slide(_prs(), L_TEKST)
    slides._logo(s, _logo_billede(), x_hoejre=5.0, y_midt=3.0, maks_b=1.0, maks_h=1.0)
    venstre, bredde = _logo_x(s)
    assert venstre + bredde == pytest.approx(5.0, abs=0.01)


def test_lange_tilsagn_centrerer_smalle_logoer_i_forvalterkolonnen():
    # et kvadratisk logo fylder langt mindre end kolonnen og ville højrestillet hænge i kanten
    fond = _lt_fond("Kvadratisk logo", 100.0)
    fond["logo"] = _logo_billede(1, 1)
    s = slides.lange_tilsagn(_prs(), _rapport([fond]))
    venstre, bredde = _logo_x(s)
    assert venstre + bredde / 2 == pytest.approx(
        slides._lt_x(1) + slides.LT_KOL[1][1] / 2, abs=0.02)


Y_HOVED = (1.20, 1.56)      # båndet, hvor kolonneoverskrifterne og deres streger ligger


def _i_hovedbaandet(sh):
    return sh.top is not None and Y_HOVED[0] <= Emu(sh.top).inches <= Y_HOVED[1]


MAERKER = "⁰¹²³⁴⁵⁶⁷⁸⁹"     # fodnotemærker efter en overskrift


def _uden_maerker(tekst):
    return tekst.translate({ord(c): None for c in MAERKER})


def _hoved(slide, navn):
    """Overskriftens tekstboks som (venstre, højre)."""
    for sh in slide.shapes:
        if (_i_hovedbaandet(sh) and sh.has_text_frame
                and sh.text_frame.text.replace("\v", "\n").strip().rstrip(MAERKER) == navn):
            return Emu(sh.left).inches, Emu(sh.left).inches + Emu(sh.width).inches
    return None


def test_lange_tilsagn_overskrifter_flugter_med_kolonnerne():
    s = slides.lange_tilsagn(_prs(), _rapport([_lt_fond("A", 100.0)]))
    for i, (navn, bredde, _) in enumerate(slides.LT_KOL):
        boks = _hoved(s, navn)
        assert boks is not None, navn
        x = slides._lt_x(i)
        assert boks == pytest.approx((x, x + bredde), abs=0.001), navn


def test_lange_tilsagn_streger_har_samme_midte_som_overskrifterne():
    s = slides.lange_tilsagn(_prs(), _rapport([_lt_fond("A", 100.0)]))
    streger = sorted(
        (Emu(sh.left).inches, Emu(sh.left).inches + Emu(sh.width).inches)
        for sh in s.shapes
        if sh.shape_type == MSO_SHAPE_TYPE.LINE and _i_hovedbaandet(sh))
    assert len(streger) == len(slides.LT_KOL)
    for (navn, bredde, just), (v, h) in zip(slides.LT_KOL, streger):
        if just == slides.C:
            i = [k for k, _, _ in slides.LT_KOL].index(navn)
            midt = slides._lt_x(i) + bredde / 2
            assert (v + h) / 2 == pytest.approx(midt, abs=0.001), navn


# ---------- Tilsagn på tværs af fonde ----------

def _tpt_fond(navn, tilsagn, valuta="DKK", kurs=1.0):
    f = _lt_fond(navn, tilsagn, valuta=valuta)
    f["n"].update(picc=0.5, nav=tilsagn * 0.75, resttilsagn_dkk=tilsagn / 2 * kurs,
                  nav_dkk=tilsagn * 0.75 * kurs)
    return f


def _tpt_rapport(fonde, total=None, total_dkk=None, kurser=None, **kw):
    r = _rapport(fonde, total=total, **kw)
    r.update(total_dkk=total_dkk, kurser=kurser or {"DKK": (1.0, date(2026, 6, 30))})
    return r


def test_tilsagn_paa_tvaers_har_egen_titel_og_dkk_kolonner():
    r = _tpt_rapport([_tpt_fond("Fond A", 100_000_000.0)])
    t = _tekster(slides.tilsagn_paa_tvaers(_prs(), r))
    assert "Investeringer" in t
    # kun DKK: ingen kursnote, så DKK-kolonnerne har intet mærke; PICC står i Definitioner
    assert "DKK\vResttilsagn" in t and "DKK\vNAV" in t and "PICC" in t
    assert "100,0" in t and "50,0 %" in t and "50,0" in t and "75,0" in t
    assert not any("Kaldt" in x for x in t)          # vandfaldet er væk


def test_tilsagn_paa_tvaers_summerer_dkk_paa_tvaers_af_valutaer():
    dkk = _tpt_fond("DKK-fond", 100_000_000.0)
    eur = _tpt_fond("EUR-fond", 10_000_000.0, valuta="EUR", kurs=7.46)
    r = _tpt_rapport([dkk, eur], total=None,
                     total_dkk=dict(resttilsagn_dkk=50_000_000 + 37_300_000,
                                    nav_dkk=75_000_000 + 55_950_000,
                                    indbetalt_dkk=100_000_000.0, udloddet_dkk=0.0),
                     kurser={"DKK": (1.0, date(2026, 6, 30)), "EUR": (7.46, date(2026, 6, 30))})
    r["valuta"] = None
    r["irr_dkk"] = dict(irr=0.0841, flows=[])
    t = _tekster(slides.tilsagn_paa_tvaers(_prs(), r))
    assert "I alt²" in t and "DKK\vNAV¹" in t
    assert "8,4 %" in t                               # porteføljens IRR i DKK
    assert "87,3" in t and "130,9" in t                   # DKK-summerne
    assert any("² Fondene har forskellige valutaer, så i alt-rækken har kun beløbene i DKK" in x
               for x in t)
    assert slides.gange((75_000 + 55_950) / 100_000) in t     # TVPI i DKK
    assert any("Kurser pr. 30.06.2026 (DKK pr. enhed): EUR 7,4600." in x for x in t)


def test_tilsagn_paa_tvaers_uden_fremmed_valuta_har_ingen_kursnote():
    r = _tpt_rapport([_tpt_fond("A", 100.0)])
    t = _tekster(slides.tilsagn_paa_tvaers(_prs(), r))
    assert not any("Nationalbanks kurs" in x for x in t)


def test_tilsagn_paa_tvaers_markerer_fond_med_afvigende_opgoerelsesdato():
    r = _tpt_rapport([_tpt_fond("Fond A", 100.0), _tpt_fond("Seed", 50.0)],
                     foraeldede=[dict(fond="Seed", dato=date(2025, 9, 30), kvartaler=3)])
    t = _tekster(slides.tilsagn_paa_tvaers(_prs(), r))
    assert f"Seed {slides.DAGGER}" in t and "opgjort pr. 30.09.2025" in t


def test_loebetid_staar_i_portefoeljeoversigten_ikke_i_tilsagn_paa_tvaers():
    f = _tpt_fond("Fond A", 100_000.0)
    f.update(loebetid="10+1+1", noter=[])
    f["n"].update(udloddet=0.0, dpi=0.0)
    r = _tpt_rapport([f])
    t = _tekster(slides.tilsagn_paa_tvaers(_prs(), r))
    assert "Løbetid" not in t and "10+1+1" not in t
    tabel = next(sh.table for sh in slides.portefoeljeoversigt(_prs(), r).shapes if sh.has_table)
    hoved = [c.text for c in tabel.rows[0].cells]
    assert hoved[hoved.index("Løbetid") - 1] == "Vintage"
    assert tabel.rows[1].cells[hoved.index("Løbetid")].text == "10+1+1"


def test_tilsagn_paa_tvaers_kolonner_fylder_indholdsbredden():
    assert sum(b for _, b, _ in slides.TPT_KOL) == pytest.approx(
        sum(b for _, b, _ in slides.LT_KOL))


def test_tilsagn_paa_tvaers_overskrifter_flugter_med_kolonnerne():
    s = slides.tilsagn_paa_tvaers(_prs(), _tpt_rapport([_tpt_fond("A", 100.0)]))
    for i, (navn, bredde, _) in enumerate(slides.TPT_KOL):
        boks = _hoved(s, navn)
        assert boks is not None, navn
        x = slides._lt_x(i, slides.TPT_KOL)
        assert boks == pytest.approx((x, x + bredde), abs=0.001), navn


def _ford_rapport():
    def fond(navn, geografi, valuta, kurs, tilsagn, indbetalt, nav):
        n = dict(tilsagn=tilsagn, indbetalt=indbetalt, udloddet=0.0, nav=nav,
                 tvpi=nav / indbetalt, tilsagn_dkk=tilsagn * kurs, nav_dkk=nav * kurs,
                 indbetalt_dkk=indbetalt * kurs, udloddet_dkk=0.0,
                 resttilsagn_dkk=(tilsagn - indbetalt) * kurs)
        return dict(fund_name=navn, geografi=geografi, segment="Buyout", valuta=valuta, n=n,
                    noter=[])
    fonde = [fond("Fond DK", "Danmark", "DKK", 1.0, 100_000, 50_000, 60_000),
             fond("Fond EU", "Europa", "EUR", 7.5, 40_000, 20_000, 30_000),
             fond("Fond EU2", "Europa", "DKK", 1.0, 200_000, 100_000, 110_000)]
    r = _rapport(fonde)
    r.update(valuta=None, kurser={"DKK": (1.0, date(2026, 6, 30)), "EUR": (7.5, date(2026, 6, 30))},
             fordelinger={"geografi": data.fordeling(fonde, "geografi")})
    return r


def _tabelraekker(slide):
    tbl = next(sh for sh in slide.shapes if sh.has_table).table
    return [[c.text for c in row.cells] for row in tbl.rows]


def test_fordeling_har_totalen_oeverst_og_en_i_alt_raekke_efter_hver_gruppe():
    sider = slides.fordeling(_prs(), _ford_rapport(), "geografi", "Geografisk fordeling")
    assert len(sider) == 1
    rk = _tabelraekker(sider[0])
    # valutaen på linjen over navnet, og fodnotemærker ved de kolonner, noterne forklarer
    assert rk[0] == ["", "Valuta", "Tilsagn¹", "DKK\vTilsagn²", "DKK\vNAV²", "DKK\vResttilsagn²",
                     "TVPI", "Andel af\vtilsagn"]   # TVPI og andel står i Definitioner
    assert [x[0].strip() for x in rk[1:]] == [
        "I alt", "Europa", "Fond EU", "Fond EU2", "I alt Europa",
        "Danmark", "Fond DK", "I alt Danmark"]
    assert rk[1][3] == slides.beloeb(600_000) and rk[1][7] == slides.pct(1.0)


def test_fordeling_i_alt_raekker_med_flere_valutaer_har_tvpi_men_ikke_tilsagn():
    rk = _tabelraekker(slides.fordeling(_prs(), _ford_rapport(), "geografi", "x")[0])
    europa, danmark = rk[5], rk[8]
    assert europa[1] == "" and europa[2] == slides.IKKE_OPLYST
    assert europa[6] == slides.gange((30_000 * 7.5 + 110_000) / (20_000 * 7.5 + 100_000))
    assert danmark[1] == "DKK" and danmark[6] == slides.gange(60_000 / 50_000)
    assert europa[7] == slides.pct(500_000 / 600_000)


def _blok(navn, antal):
    return ([([navn], "gruppe")] + [([f"{navn} {i}"], "fond") for i in range(antal)]
            + [([f"I alt {navn}"], "total")])


def test_del_paa_sider_bryder_mellem_grupperne():
    sider = slides._del_paa_sider([[(["I alt"], "total")], _blok("A", 3), _blok("B", 3)], 8)
    assert [[c[0] for c, _ in side] for side in sider] == [
        ["I alt", "A", "A 0", "A 1", "A 2", "I alt A"], ["B", "B 0", "B 1", "B 2", "I alt B"]]


def test_del_paa_sider_fylder_siden_op_frem_for_at_lade_den_staa_halvtom():
    sider = slides._del_paa_sider([_blok("A", 3), _blok("B", 6)], 10)
    tekster = [[c[0] for c, _ in side] for side in sider]
    assert tekster[0][5:] == ["B", "B 0", "B 1", "B 2", "B 3"]
    assert tekster[1] == ["B (fortsat)", "B 4", "B 5", "I alt B"]


def test_del_paa_sider_deler_en_for_lang_gruppe_med_fortsat_overskrift():
    sider = slides._del_paa_sider([_blok("A", 10)], 5)
    assert all(len(side) <= 5 for side in sider)
    tekster = [[c[0] for c, _ in side] for side in sider]
    assert tekster[0][0] == "A" and tekster[1][0] == "A (fortsat)"
    assert tekster[-1][-1] == "I alt A"
    fonde = [t for side in tekster for t in side if t.startswith("A ") and "fortsat" not in t]
    assert fonde == [f"A {i}" for i in range(10)]           # ingen fond tabt eller gentaget


def test_fordeling_fortsaetter_paa_ny_slide_naar_tabellen_er_for_lang():
    r = _ford_rapport()
    skabelon = r["fonde"][0]
    fonde = [skabelon | dict(fund_name=f"Fond {i:02d}", geografi=f"Region {i % 6}")
             for i in range(40)]
    r.update(fonde=fonde, fordelinger={"geografi": data.fordeling(fonde, "geografi")})
    sider = slides.fordeling(_prs(), r, "geografi", "Geografisk fordeling")
    assert len(sider) > 1
    assert "Geografisk fordeling (fortsat)" in _tekster(sider[1])
    alle = [rk for s in sider for rk in _tabelraekker(s)[1:]]
    assert sum(1 for rk in alle if rk[0].strip().startswith("Fond ")) == 40
    assert _tabelraekker(sider[1])[0][1] == "Valuta"          # kolonnehovedet gentages
    assert _tabelraekker(sider[0])[1][0] == "I alt"


def test_definitioner_viser_alle_begreber_inden_for_fladen():
    sider = slides.definitioner(_prs(), _rapport([]), supplerende=True)
    t = [x for s in sider for x in _tekster(s)]
    assert "Definitioner" in t
    assert ("Definitioner (fortsat)" in t) == (len(sider) > 1)
    for begreb, tekst in slides.DEFINITIONER:
        assert begreb in t and tekst in t, begreb
    for defs in slides._def_sider(slides.DEFINITIONER):
        for spalte in slides._def_halvdele(defs):
            hoejde = sum(slides._def_hoejde(b, t) + slides.DEF_LUFT for b, t in spalte)
            assert slides.Y_TOP + hoejde <= slides.Y_NOTER + 0.4     # over Petersen & Partners-linjen


def test_sammendrag_viser_noegletal_hoejdepunkter_og_fordeling():
    r = _ford_rapport()
    for f, m, v in zip(r["fonde"], ("X", "Y", "X"), (2020, 2022, 2023)):
        f.update(manager=m, vintage=v, fund_name=f["fund_name"])
        f["n"]["irr"] = 0.12
    r["fordelinger"]["segment"] = data.fordeling(r["fonde"], "segment")
    r["total_dkk"] = data.dkk_total(r["fonde"])
    r["sammendrag"] = data.sammendrag(r["fonde"], r["total_dkk"])
    s = slides.sammendrag(_prs(), r)
    t = _tekster(s)
    assert "Executive summary" in t
    for label in ("fonde i porteføljen", "tilsagn", "resttilsagn", "NAV",
                  "TVPI¹", "IRR¹"):
        assert label in t, label
    assert any(x.startswith("eksponering") and x.endswith("resttilsagn + NAV") for x in t)
    assert slides.beloeb(r["total_dkk"]["tilsagn_dkk"]) in t
    td = r["total_dkk"]
    assert slides.beloeb(td["resttilsagn_dkk"] + td["nav_dkk"]) in t
    assert any("3 fonde hos 2 forvaltere med vintage 2020–2023" in x for x in t)
    assert any(x.startswith("Største geografi er Europa") for x in t)
    assert any(f"Den samlede eksponering er {slides.beloeb(td['resttilsagn_dkk'] + td['nav_dkk'])} "
               "mio. DKK inklusive NAV, og resttilsagnet på "
               f"{slides.beloeb(td['resttilsagn_dkk'])} mio. DKK kan stadig kaldes af forvalterne."
               in x for x in t)
    # uden gennemlyste regioner falder geografien tilbage på fondenes fokus efter tilsagn;
    # strategien er efter NAV
    assert "Geografi – andel af tilsagn²" in t and "Investeringsstrategi – andel af NAV" in t
    assert any("² Geografien er fondenes geografiske fokus." in x for x in t)
    assert not any("egment" in x for x in t)
    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text not in ("", "Executive summary"):
            assert Emu(sh.top).inches <= slides.Y_NOTER + 0.01, sh.text_frame.text


# ---------- udvikling i kvartalet og året ----------

def test_aendring_skriver_fortegn():
    assert slides.aendring(1_234_567) == "+1,2"
    assert slides.aendring(-1_234_567) == "-1,2"
    assert slides.aendring(20_000) == "0,0"
    assert slides.aendring(-0.001, slides.gange) == "0,00x"
    assert slides.aendring(None) == "n/a"


def _udv(primo, kald, udl, vaerdi, valuta=0.0):
    return dict(fra=date(2026, 3, 31), til=date(2026, 6, 30), primo=primo, kald=kald,
                udlodning=udl, vaerdi=vaerdi, valuta=valuta,
                ultimo=primo + kald - udl + vaerdi + valuta, afkast=0.02)


def _udv_rapport(foraeldede=()):
    fonde = []
    for navn, valuta, kald, vaerdi in (("Alfa Fund I K/S", "DKK", 1e6, 2e6),
                                       ("Beta Fund II A/S", "USD", 0.0, -5e5)):
        u = dict(_udv(1e7, kald, 0.0, vaerdi), primo_dato=date(2026, 3, 31), flows=[],
                 tvpi_aendring=0.03, kald_dkk=kald, udlodning_dkk=0.0, vaerdi_dkk=vaerdi,
                 valuta_dkk=1e5)
        fonde.append(dict(fund_name=navn, valuta=valuta, noter=[], udvikling={"kvartal": u,
                          "aar": u}, n=dict(nav_dkk=1e7, tvpi=1.4)))
    r = _rapport(fonde, foraeldede=foraeldede)
    r["perioder"] = data.perioder(date(2026, 6, 30))
    r["udvikling"] = {p: dict(_udv(2e7, 1e6, 0.0, 1.5e6, 2e5),
                              bidrag=[("Alfa Fund I K/S", 2.1e6), ("Beta Fund II A/S", -4e5)])
                      for p in ("kvartal", "aar")}
    r["sammendrag"] = dict(tvpi=1.4)
    return r


def _indenfor_fladen(s, titel):
    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text not in ("", titel):
            assert Emu(sh.top).inches <= slides.Y_NOTER + 0.01, sh.text_frame.text
            assert Emu(sh.left + sh.width).inches <= slides.X1 + 0.3, sh.text_frame.text


def test_udvikling_viser_to_broer_og_bidrag_pr_fond():
    s = slides.udvikling(_prs(), _udv_rapport())
    t = _tekster(s)
    assert "Udvikling i kvartalet og året" in t
    assert "Kvartalet (Q2 26)" in t and "År til dato (2026)" in t
    # afkast og værdiændring står i Definitioner og har ingen fodnote
    assert t.count("NAV primo¹") == 2 and t.count("Valutakurs²") == 2
    assert t.count("Værdiændring") == 2
    assert "Alfa Fund I" in t and "+2,1" in t and "-0,4" in t
    assert "+1,7" in t                    # i alt: værdiændring + valuta
    assert not any("Modified Dietz" in x for x in t)
    _indenfor_fladen(s, "Udvikling i kvartalet og året")


def test_udvikling_i_foerste_kvartal_viser_en_bro():
    r = _udv_rapport()
    r["perioder"] = data.perioder(date(2026, 3, 31))
    t = _tekster(slides.udvikling(_prs(), r))
    assert t.count("NAV primo¹") == 1 and "Kvartalet og år til dato (Q1 26)" in t


def test_udvikling_springes_over_uden_dkk_tal():
    r = _udv_rapport()
    r["udvikling"] = {"kvartal": None, "aar": None}
    assert slides.udvikling(_prs(), r) is None
    assert slides.aendringer_pr_fond(_prs(), r) is None


def test_aendringer_pr_fond_har_en_raekke_pr_fond_og_i_alt():
    r = _udv_rapport(foraeldede=[dict(fond="Beta Fund II A/S", dato=date(2026, 3, 31),
                                      kvartaler=1)])
    s = slides.aendringer_pr_fond(_prs(), r)
    tabel = next(sh.table for sh in s.shapes if sh.has_table)
    raekker = [[c.text for c in rk.cells] for rk in tabel.rows]
    assert [x[0] for x in raekker[1:]] == ["Alfa Fund I", f"Beta Fund II {slides.DAGGER}", "I alt"]
    assert raekker[1][4] == "+2,1"          # værdiændring + valuta i DKK
    assert raekker[-1][4] == "+1,7"
    assert len(raekker[0]) == 2 + 4 * 2 + 3
    assert raekker[0][4] == "Værdi-\x0bændring¹" and raekker[0][5] == "Afkast²"


def _rt_rapport():
    r = _udv_rapport()

    def linje(navn, valuta, gebyrer, carry, kildeskat=0.0, rettet=None):
        return dict(fund_name=navn, valuta=valuta, primo=1e7, tilgang=1e6, afgang=-2e5,
                    vaerdi=5e5, kurs=1e5, ultimo=1.14e7, gebyrer=gebyrer, carry=carry,
                    resttilsagn=3e6, kildeskat=kildeskat, primo_rapporteret=rettet)
    linjer = [linje("Alfa Fund I K/S", "DKK", -25_000.0, -40_000.0, rettet=9.9e6),
              linje("Beta Fund II A/S", "USD", None, None, kildeskat=12_000.0)]
    i_alt = {k: sum(x[k] or 0 for x in linjer) for k in data.REGNSKABSFELTER}
    i_alt.update(gebyrer=-25_000.0, carry=-40_000.0)
    r["regnskabstal"] = {p: dict(fra=r["perioder"][p][0], til=r["perioder"][p][1], fonde=linjer,
                                 i_alt=i_alt, udeladte_gebyrer=["Beta Fund II A/S"],
                                 udeladte_carry=["Beta Fund II A/S"])
                         for p in ("kvartal", "aar")}
    return r


def test_regnskabstal_viser_bevaegelsesnoten_i_tusinder():
    r = _rt_rapport()
    s = slides.regnskabstal(_prs(), r, "aar")
    tabel = next(sh.table for sh in s.shapes if sh.has_table)
    raekker = [[c.text for c in rk.cells] for rk in tabel.rows]
    assert raekker[0][6] == "heraf\x0bgebyrer" and raekker[0][-1] == "Resttilsagn"
    alfa, beta, i_alt = raekker[1], raekker[2], raekker[3]
    assert alfa[:10] == ["Alfa Fund I", "DKK", "10.000¹", "+1.000", "-200", "+500", "-25", "-40",
                         "+100", "11.400"]
    assert beta[4] == "-200²" and beta[6:8] == ["n/a", "n/a"]
    assert i_alt[6:8] == ["-25³", "-40³"] and i_alt[-1] == "6.000"
    t = _tekster(s)
    assert "Regnskabstal – år til dato" in t and "År til dato (2026)" in t
    noter = next(x for x in t if x.startswith("Kilde"))
    assert "¹ Forvalteren har rettet NAV primo; oprindeligt rapporteret 9.900." in noter
    assert "² Kildeskat tilbageholdt af udlodningerne: 12." in noter
    assert "³ I alt summerer de fonde, der oplyser beløbet." in noter
    assert "eventualforpligtelse" in noter
    assert tabel.cell(1, 6).text_frame.paragraphs[0].runs[0].font.italic
    _indenfor_fladen(s, "Regnskabstal – år til dato")


def test_regnskabstal_i_foerste_kvartal_er_en_slide():
    r = _rt_rapport()
    assert slides.regnskabstal(_prs(), r, "kvartal") is not None
    r["perioder"] = data.perioder(date(2026, 3, 31))
    assert slides.regnskabstal(_prs(), r, "aar") is None
    s = slides.regnskabstal(_prs(), r, "kvartal")
    assert "Regnskabstal" in _tekster(s) and "Kvartalet og år til dato (Q1 26)" in _tekster(s)
    r["regnskabstal"] = {"kvartal": None, "aar": None}
    assert slides.regnskabstal(_prs(), r, "kvartal") is None


def test_aendringer_pr_fond_maerker_afkast_med_modsat_fortegn_og_skiller_perioderne():
    r = _udv_rapport()
    beta = r["fonde"][1]
    beta["udvikling"] = {p: dict(u) for p, u in beta["udvikling"].items()}
    beta["udvikling"]["aar"].update(afkast=-0.023, vaerdi_dkk=-1e5, valuta_dkk=4e5)
    r["fortegnsafvigelser"] = data.fortegnsafvigelser(r["fonde"])
    s = slides.aendringer_pr_fond(_prs(), r)
    tabel = next(sh.table for sh in s.shapes if sh.has_table)
    beta_raekke = [c.text for c in tabel.rows[2].cells]
    # kvartalet: værdien faldt også i USD, men afkastet er positivt (Dietz); året: kursen
    assert beta_raekke[5] == "2,0 %⁴" and beta_raekke[9] == "-2,3 %⁵"
    t = _tekster(s)
    # mærket står i fondens række og periodens kolonne, så noten nævner ingen af dem
    assert any("⁴ Afkastet har modsat fortegn af værdiændringen, fordi Modified Dietz" in x
               for x in t)
    assert any("⁵ Fondens værdi faldt i USD, men USD steg over for DKK" in x for x in t)
    assert not any("Beta Fund II, " in x for x in t)
    assert [c.text for c in tabel.rows[1].cells][5] == "2,0 %"      # Alfa uden afvigelse
    # tre lodrette streger: før kvartalet, før året og før stillingen pr. dato
    lodrette = [sh for sh in s.shapes if sh.shape_type == MSO_SHAPE_TYPE.LINE
                and sh.width == 0 and sh.height > 0]
    assert len(lodrette) == 3


def test_vaerdibevaegelser_viser_stigninger_og_fald():
    r = _udv_rapport()
    r["vaerdibevaegelser"] = dict(
        linjer=[dict(navn="Selskab A", fond="Alfa Fund I K/S", type="Selskab", kvartal=3e5,
                     aar=4e5),
                dict(navn="Fond B", fond="Alfa Fund I K/S", type="Secondary", kvartal=-1e5,
                     aar=None),
                dict(navn="Selskab C", fond="Alfa Fund I K/S", type="Selskab", aar=1e5)],
        uaendrede={"kvartal": ["Beta Fund II A/S"]},
        usammenlignelige={"aar": [("Alfa Fund I K/S", "Gammel")]},
        uskalerede=[], uden_beholdninger=[], dato=date(2026, 6, 30))
    s = slides.vaerdibevaegelser(_prs(), r)
    t = _tekster(s)
    assert "Største stigninger i Q2 26" in t and "Største fald i Q2 26" in t
    # typen af den underliggende investering i egen kolonne
    assert t.count("Type") == 2 and "Selskab" in t and "Secondary" in t
    assert any("Værdierne af porteføljeselskaberne er ikke opdateret i kvartalet" in x for x in t)
    assert "Selskab A" in t and "+0,3" in t and "Fond B" in t and "-0,1" in t
    assert "Selskab C" not in t           # ingen bevægelse i kvartalet
    noter = t[-1]
    assert "Beta Fund II A/S" in noter and "Gammel" in noter
    # én række med hele porteføljens værdiændring + valuta, som på Ændringer pr. fond
    assert _tabelraekker(s) == [["Værdiændring i alt, hele porteføljen",
                                 "Kvartalet (Q2 26): +1,7", "År til dato (2026): +1,7"]]
    _indenfor_fladen(s, "Største værdibevægelser i porteføljen")


def test_vaerdibevaegelser_uden_dkk_udvikling_har_ingen_totalraekke():
    r = _udv_rapport()
    r["udvikling"] = {"kvartal": None, "aar": None}
    r["vaerdibevaegelser"] = dict(
        linjer=[dict(navn="Selskab A", fond="Alfa Fund I K/S", kvartal=3e5)],
        uaendrede={}, usammenlignelige={}, uskalerede=[], uden_beholdninger=[],
        dato=date(2026, 6, 30))
    s = slides.vaerdibevaegelser(_prs(), r)
    assert not any(sh.has_table for sh in s.shapes)


# ---------- gennemlysning af feedere ----------

def _feeder():
    raekker = [dict(navn=f"Selskab {i}", type="UNREALISED", instrument="PIK", vintage=None,
                    closing=date(2025, 1, 1), kostpris=10.0 * i, dagsvaerdi=11.0 * i,
                    udloddet=0.0, total=11.0 * i, kostpris_fond=1.0, dagsvaerdi_fond=1.0,
                    udloddet_fond=0.0, sektor=None, tvpi=None, irr=None, document_id=1)
               for i in range(14, 0, -1)]
    beh = dict(dato=date(2026, 6, 30), raekker=raekker, skaleret=True, andel=0.01,
               type="selskaber", kode="PSCP5", via="Park Square Capital Partners V, SCSp",
               fordeling=[("Branche", [dict(navn="Industri", andel=0.6),
                                       dict(navn="Øvrige", andel=0.4)])],
               fakta={"CREDIT_LINE_DRAWN": 5e6})
    f = _lt_fond("Petersen & Partners Private Credit II A/S", 10_000.0, valuta="EUR")
    f.update(beholdninger=beh, manager="PP", geografi="Europa", fund_code="PPPC2",
             nav_dato=date(2026, 6, 30), beskrivelse="Feeder.", noter=[], historik=[],
             bro=[], virksomheder={})
    f["n"].update(andel=0.29, fondsstoerrelse=34_600.0, picc=0.3, udloddet=0.0, nav=3465.0,
                  dpi=0.0, gp_irr=None, gp_tvpi=None)
    return f


def test_feederside_viser_gennemlysning_i_stedet_for_historik():
    s = slides.fondsside(_prs(), _feeder(), "Investor A/S")
    t = _tekster(s)
    assert any(x.startswith("Gennemlysning af Park Square Capital Partners V") for x in t)
    assert not any(x.startswith("Udvikling i NAV") for x in t)
    assert not any(sh.has_chart for sh in s.shapes)
    tabel = next(sh.table for sh in s.shapes
                 if sh.has_table and sh.table.cell(0, 0).text == "Selskab")
    navne = [rk.cells[0].text for rk in tabel.rows]
    assert navne[1] == "Selskab 14" and navne[-2] == "Øvrige (4)" and navne[-1] == "I alt"
    assert "Industri" in t and any("kreditfacilitet" in x for x in t)


def test_sidefod_kun_paa_indholdsslides_med_investor_og_slidenummer():
    prs = _prs()
    for layout in (L_FORSIDE, L_TEKST, L_SKILLE[0], L_TEKST, L_SLUT):
        _ny_slide(prs, layout)
    slides.sidefod(prs, "A/S Test")

    def fod(s):
        return [ph for ph in s.placeholders if ph.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER]

    alle = list(prs.slides)
    assert [len(fod(s)) for s in alle] == [0, 1, 0, 1, 0]
    assert fod(alle[1])[0].text_frame.text == "A/S Test | Side 2"
    assert fod(alle[3])[0].text_frame.text == "A/S Test | Side 4"
    p = fod(alle[1])[0].text_frame.paragraphs[0]
    assert p.runs[0].font.bold and not p.runs[1].font.bold
    assert p._p.find(qn("a:fld")).get("type") == "slidenum"


def _branche_rapport(sektorer, foreloebig_andel=0.0, art="brancher"):
    total = sum(s["dkk"] for s in sektorer)
    for s in sektorer:
        s.setdefault("kode", "INDUSTRIALS")
        s.setdefault("under", s.pop("grupper", []))
        s.setdefault("stoerste_fond", "K Alternativ Private Equity 2021")
        s.setdefault("fonde", [s["stoerste_fond"]])
        s["andel"] = s["dkk"] / total
        for g in s["under"]:
            g["andel"] = g["dkk"] / total
            g.setdefault("stoerste_fond", s["stoerste_fond"])
    r = _rapport([])
    r[art] = dict(total_dkk=total, grupper=sektorer, foreloebig_dkk=0.0,
                  foreloebig_andel=foreloebig_andel, foreloebige_fonde=[],
                  andre_datoer=[], efter_strategi=[], uden_fordeling=[])
    r["kurser"] = {}
    return r


def _cirkel_vaerdier(s):
    diagram = next(sh.chart for sh in s.shapes if sh.has_chart)
    return list(diagram.plots[0].categories), list(diagram.plots[0].series[0].values)


def test_brancher_cirkel_summerer_til_tabellens_i_alt_og_samler_smaa_sektorer():
    r = _branche_rapport([dict(navn="Industri", dkk=600_000.0),
                          dict(navn="Finans", dkk=385_000.0, kode="FINANCIALS"),
                          dict(navn="Energi", dkk=10_000.0, kode="ENERGY"),
                          dict(navn="Forsyning", dkk=5_000.0, kode="UTILITIES")])
    s = slides.brancher(_prs(), r)
    kategorier, vaerdier = _cirkel_vaerdier(s)
    assert kategorier == ["Industri", "Finans", "Øvrige"]
    assert sum(vaerdier) == pytest.approx(1.0)        # mio. DKK, som tabellens i alt-række
    tabel = next(sh.table for sh in s.shapes if sh.has_table)
    tekst = [[c.text for c in row.cells] for row in tabel.rows]
    assert tekst[1][1:4] == ["I alt", "1,0", "100,0 %"]
    assert ["Energi", "Forsyning"] == [x[1] for x in tekst if x[1] in ("Energi", "Forsyning")]


def test_brancher_noterer_foreloebig_klassifikation_kun_naar_den_findes():
    sektorer = lambda: [dict(navn="Industri", dkk=900.0), dict(navn="Finans", dkk=100.0,
                                                                kode="FINANCIALS")]
    med = " ".join(_tekster(slides.brancher(_prs(), _branche_rapport(sektorer(), 0.25))))
    uden = " ".join(_tekster(slides.brancher(_prs(), _branche_rapport(sektorer()))))
    assert "25,0 % af NAV er en foreløbig klassifikation" in med
    assert "foreløbig" not in uden


def test_brancher_uden_data_springes_over():
    r = _rapport([])
    r["brancher"] = None
    assert slides.brancher(_prs(), r) is None


def test_brancher_har_ingen_bidragskolonne_og_forklarer_diversificeret():
    r = _branche_rapport([dict(navn="Industri", dkk=900.0, stoerste_fond="Seed Capital Denmark IV"),
                          dict(navn="Diversificeret (fondsandele)", dkk=100.0, kode="DIVERSIFIED",
                               stoerste_fond="Petersen & Partners Private Equity II")])
    s = slides.brancher(_prs(), r)
    tabel = next(sh.table for sh in s.shapes if sh.has_table)
    assert len(tabel.columns) == 4
    assert [c.text for c in tabel.rows[0].cells] == ["", "Sektor", "NAV", "Andel"]
    assert not any("Seed Capital Denmark IV" in x for x in _tekster(s))
    assert any("som forvalteren af Petersen & Partners Private Equity II ikke fordeler"
               in x for x in _tekster(s))


def test_brancher_viser_kun_sektorer():
    industri = dict(navn="Industri", dkk=900.0, grupper=[
        dict(navn="Kapitalgoder", dkk=600.0), dict(navn=data.IKKE_OPDELT, dkk=300.0)])
    s = slides.brancher(_prs(), _branche_rapport([industri, dict(navn="Finans", dkk=100.0,
                                                                 kode="FINANCIALS")]))
    tabel = next(sh.table for sh in s.shapes if sh.has_table)
    navne = [row.cells[1].text.strip() for row in tabel.rows]
    assert navne == ["Sektor", "I alt", "Industri", "Finans"]
    assert not any("Ikke opdelt" in x for x in _tekster(s))


def _overlapper(a, b):
    return (a[0] < b[0] + b[2] - 1e-9 and b[0] < a[0] + a[2] - 1e-9
            and a[1] < b[1] + b[3] - 1e-9 and b[1] < a[1] + a[3] - 1e-9)


def test_treemap_giver_arealer_i_forhold_til_vaerdierne_uden_overlap():
    vaerdier = [27.5, 17.3, 15.2, 14.8, 12.0, 7.0, 4.2, 2.0]
    kasser = slides._treemap(vaerdier, 1.0, 2.0, 10.0, 4.0)
    for v, (x, y, w, h) in zip(vaerdier, kasser):
        assert w * h == pytest.approx(v / sum(vaerdier) * 40.0)
        assert x >= 1.0 - 1e-9 and y >= 2.0 - 1e-9
        assert x + w <= 11.0 + 1e-9 and y + h <= 6.0 + 1e-9
    assert not any(_overlapper(a, b) for i, a in enumerate(kasser) for b in kasser[i + 1:])
    # squarified: ingen kasse bliver en tynd strimmel
    assert max(max(w / h, h / w) for _, _, w, h in kasser) < 4


def test_branchekort_viser_navn_beloeb_og_andel_og_samler_smaa_sektorer():
    r = _branche_rapport([dict(navn="Industrials", dkk=600_000.0),
                          dict(navn="Financials", dkk=385_000.0, kode="FINANCIALS"),
                          dict(navn="Energy", dkk=10_000.0, kode="ENERGY"),
                          dict(navn="Utilities", dkk=5_000.0, kode="UTILITIES")])
    s = slides.branchekort(_prs(), r)
    t = _tekster(s)
    assert "Branchefordeling – hele porteføljen" in t
    assert "Industrials\n0,6 mio. DKK\n60,0 %" in t
    # Industrials, Financials og én samlet kasse for de to små sektorer
    assert sum(sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE for sh in s.shapes) == 3
    # GICS-navnene er egennavne og står med stort
    assert any("»Øvrige« samler Energy og Utilities" in x for x in t)


def test_regioner_viser_kun_regionerne_i_regionens_farve():
    europa = dict(navn="Europa", kode="EUROPE", dkk=800_000.0,
                  stoerste_fond="Seed Capital Denmark IV",
                  grupper=[dict(navn="Danmark", dkk=600_000.0),
                           dict(navn="Øvrige Europa og ikke opdelt", dkk=200_000.0)])
    usa = dict(navn="Nordamerika", kode="NORTH_AMERICA", dkk=200_000.0,
               grupper=[dict(navn="USA", dkk=200_000.0)])
    s = slides.regioner(_prs(), _branche_rapport([europa, usa], art="regioner"))
    kategorier, vaerdier = _cirkel_vaerdier(s)
    assert kategorier == ["Europa", "Nordamerika"]
    assert sum(vaerdier) == pytest.approx(1.0)        # mio. DKK
    diagram = next(sh.chart for sh in s.shapes if sh.has_chart)
    farver = [p.format.fill.fore_color.rgb for p in diagram.plots[0].series[0].points]
    assert farver == [slides.GEO_FARVER["EUROPE"][0], slides.GEO_FARVER["NORTH_AMERICA"][0]]
    tabel = next(sh.table for sh in s.shapes if sh.has_table)
    navne = [row.cells[1].text.strip() for row in tabel.rows]
    assert "Europa" in navne and "Nordamerika" in navne
    assert "Danmark" not in navne and "USA" not in navne
    assert len(tabel.columns) == 4
    assert not any("ikke opdelt" in x for x in _tekster(s))


def test_regioner_uden_data_springes_over():
    r = _rapport([])
    r["regioner"] = None
    assert slides.regioner(_prs(), r) is None


def test_sammendrag_bruger_regionerne_fra_geografisk_eksponering():
    r = _ford_rapport()
    for f, m, v in zip(r["fonde"], ("X", "Y", "X"), (2020, 2022, 2023)):
        f.update(manager=m, vintage=v)
        f["n"]["irr"] = 0.12
    r["fordelinger"]["segment"] = data.fordeling(r["fonde"], "segment")
    r["total_dkk"] = data.dkk_total(r["fonde"])
    r["sammendrag"] = data.sammendrag(r["fonde"], r["total_dkk"])
    r["regioner"] = dict(grupper=[dict(navn="Europa", andel=0.75), dict(navn="USA", andel=0.25)])
    t = _tekster(slides.sammendrag(_prs(), r))
    assert "Geografi – andel af NAV²" in t and "Investeringsstrategi – andel af NAV" in t
    assert "Europa" in t and "USA" in t and "75,0 %" in t
    assert any("Største geografi er Europa med 75,0 % af NAV" in x for x in t)
    assert any("Geografien er fordelt efter fondenes underliggende investeringer." in x
               for x in t)
    assert not any("på siden Geografisk eksponering" in x for x in t)   # siden er ikke i decket


def test_sammendrag_viser_porteføljens_irr_i_dkk():
    r = _ford_rapport()
    for f, m, v in zip(r["fonde"], ("X", "Y", "X"), (2020, 2022, 2023)):
        f.update(manager=m, vintage=v)
        f["n"]["irr"] = 0.12
    r["fordelinger"]["segment"] = data.fordeling(r["fonde"], "segment")
    r["total_dkk"] = data.dkk_total(r["fonde"])
    r["sammendrag"] = data.sammendrag(r["fonde"], r["total_dkk"], 0.0841)
    t = _tekster(slides.sammendrag(_prs(), r))
    assert "IRR¹" in t and "8,4 %" in t          # egen flise med samme DKK-note som TVPI
    assert any("Porteføljens nøgletal, også IRR, er beregnet på beløbene i DKK." in x for x in t)


def test_sammendrag_naevner_det_stoerste_segment_ikke_det_foerste():
    r = _ford_rapport()
    for f, m, v in zip(r["fonde"], ("X", "Y", "X"), (2020, 2022, 2023)):
        f.update(manager=m, vintage=v)
        f["n"]["irr"] = 0.12
    r["total_dkk"] = data.dkk_total(r["fonde"])
    r["sammendrag"] = data.sammendrag(r["fonde"], r["total_dkk"])
    # segmenterne står i strategiorden, så det største er ikke nødvendigvis først
    r["fordelinger"]["segment"] = dict(
        total=dict(andel_nav=1.0),
        grupper=[dict(navn="Private equity", andel=0.6, andel_nav=0.3),
                 dict(navn="Ejendomme", andel=0.4, andel_nav=0.7)])
    t = _tekster(slides.sammendrag(_prs(), r))
    assert any("største investeringsstrategi er Ejendomme med 70,0 % af NAV" in x for x in t)


def test_sammendrag_viser_strategien_efter_tilsagn_uden_nav_i_dkk():
    r = _ford_rapport()
    for f, m, v in zip(r["fonde"], ("X", "Y", "X"), (2020, 2022, 2023)):
        f.update(manager=m, vintage=v)
        f["n"]["irr"] = 0.12
    r["total_dkk"] = data.dkk_total(r["fonde"])
    r["sammendrag"] = data.sammendrag(r["fonde"], r["total_dkk"])
    r["fordelinger"]["segment"] = dict(
        total=dict(andel_nav=None),
        grupper=[dict(navn="Private equity", andel=0.3, andel_nav=None),
                 dict(navn="Ejendomme", andel=0.7, andel_nav=None)])
    t = _tekster(slides.sammendrag(_prs(), r))
    assert "Investeringsstrategi – andel af tilsagn" in t
    assert any("største investeringsstrategi er Ejendomme med 70,0 % af tilsagnet" in x
               for x in t)


def test_sammendrag_naevner_stoerste_enkelteksponering():
    r = _ford_rapport()
    for f, m, v in zip(r["fonde"], ("X", "Y", "X"), (2020, 2022, 2023)):
        f.update(manager=m, vintage=v)
        f["n"]["irr"] = 0.12
    r["fordelinger"]["segment"] = data.fordeling(r["fonde"], "segment")
    r["total_dkk"] = data.dkk_total(r["fonde"])
    r["sammendrag"] = data.sammendrag(r["fonde"], r["total_dkk"])
    r["stoerste_eksponering"] = dict(navn="FLATPAY", art="porteføljeselskab",
                                     fond="Seed Capital Denmark IV", dkk=54_380_000.0,
                                     andel=0.171, udeladt=["Home.Earth"])
    t = _tekster(slides.sammendrag(_prs(), r))
    assert any("Største enkelteksponering er FLATPAY (porteføljeselskab via Seed Capital Denmark "
               "IV) med 54,4 mio. DKK, svarende til 17,1 % af NAV²." in x for x in t)
    assert any("Home.Earth indgår ikke" in x for x in t)


def test_selskabstabel_har_branchekolonne_med_gics_sektor():
    prs = Presentation(str(MASTER))
    _ryd_slides(prs)
    s = _ny_slide(prs, L_TEKST)
    raekker = [_selskab("A", date(2023, 1, 10), None) | dict(industri="Finans")]
    beh = dict(dato=date(2026, 6, 30), raekker=raekker, skaleret=True, type="selskaber")
    slides._beholdningstabel(s, 3.0, beh, dict(valuta="DKK"), "Investor A/S")
    tabel = next(sh.table for sh in s.shapes if sh.has_table)
    assert tabel.cell(0, 1).text == "Branche" and tabel.cell(1, 1).text == "Finans"


def test_branchekort_uden_data_springes_over():
    r = _rapport([])
    r["brancher"] = None
    assert slides.branchekort(_prs(), r) is None


def test_sammendrag_har_ikke_tvpi_punktet():
    r = _ford_rapport()
    for f, m, v in zip(r["fonde"], ("X", "Y", "X"), (2020, 2022, 2023)):
        f.update(manager=m, vintage=v)
        f["n"]["irr"] = 0.12
    r["fordelinger"]["segment"] = data.fordeling(r["fonde"], "segment")
    r["total_dkk"] = data.dkk_total(r["fonde"])
    r["sammendrag"] = data.sammendrag(r["fonde"], r["total_dkk"])
    assert not any("TVPI har" in x for x in _tekster(slides.sammendrag(_prs(), r)))


def _fond_med_underfonde(antal=2, afstemning=None):
    """En almindelig fondsside (ikke feeder) med en kort tabel over underliggende fonde."""
    f = _feeder()
    raekker = [dict(navn=f"Fond {i}", type="FUND", instrument=None, vintage=2021, closing=None,
                    kostpris=100.0, dagsvaerdi=120.0, udloddet=10.0, total=130.0,
                    kostpris_fond=1.0, dagsvaerdi_fond=1.0, udloddet_fond=0.0, sektor=None,
                    tvpi=1.3, irr=10.0, document_id=1) for i in range(antal)]
    f["beholdninger"] = dict(dato=date(2026, 6, 30), raekker=raekker, skaleret=True, andel=0.03,
                             type="fonde", kode="PD3LGT")
    f.update(fund_name="Private Debt III LGT", valuta="DKK", nav_afstemning=afstemning,
             historik=[dict(dato=date(2025, 12, 31), nav=1.0, indbetalt=1.0, udloddet=0.0),
                       dict(dato=date(2026, 6, 30), nav=2.0, indbetalt=1.0, udloddet=0.0)],
             bro=[("Indbetalt kapital", 1.0, True), ("NAV", 2.0, True)])
    return f


def test_fondsside_med_underliggende_fonde_har_hverken_graf_eller_vaerdibro():
    s = slides.fondsside(_prs(), _fond_med_underfonde(), "Investor A/S")
    assert not any(sh.has_chart for sh in s.shapes)
    assert not any(x.startswith("Fra indbetalt kapital til NAV") for x in _tekster(s))
    assert any(x.startswith("Underliggende fonde") for x in _tekster(s))


def test_fondsside_forklarer_forskellen_mellem_nav_og_dagsvaerdi():
    a = dict(nav=54_956_000.0, dagsvaerdi=62_260_000.0, forskel=-7_304_000.0,
             carry=-8_440_000.0, oevrige=1_136_000.0, ejerandel=0.0, andel=0.0287,
             nav_andel=0.0287, beholdninger_dato=date(2026, 6, 30), nav_dato=date(2026, 6, 30))
    noter = " ".join(_tekster(slides.fondsside(_prs(), _fond_med_underfonde(afstemning=a),
                                               "Investor A/S")))
    # NAV og dagsværdien i alt står allerede på siden; noten giver kun forskellen
    assert ("NAV er 7,3 mio. DKK lavere end dagsværdien i alt: hensat carried "
            "interest -8,4; fondens likvider og øvrige poster +1,1.") in noter
    assert "Total værdi er dagsværdi plus positionernes udlodninger" in noter
    uden = dict(a, carry=None, oevrige=None, ejerandel=None, nav_andel=None)
    noter = " ".join(_tekster(slides.fondsside(_prs(), _fond_med_underfonde(afstemning=uden),
                                               "Investor A/S")))
    assert "Forvalteren oplyser ikke fondens samlede NAV, så forskellen kan ikke opdeles" in noter


def test_fondsside_saetter_maerket_ved_det_tal_noten_handler_om():
    f = _fond_med_underfonde()
    f["noter"] = [data.Note("Indbetalt kapital indeholder 82.500 DKK i udligningsrente.",
                            "indbetalt"),
                  "Nationalbanken oplyser ingen kurs."]
    s = slides.fondsside(_prs(), f, "Investor A/S")
    t = _tekster(s)
    celler = [c.text for sh in s.shapes if sh.has_table for rk in sh.table.rows for c in rk.cells]
    assert "Indbetalt kapital¹" in celler and "Tilsagn" in celler
    noter = next(x for x in t if x.startswith("Kilde:"))
    assert "Nationalbanken oplyser ingen kurs." in noter                  # generel, uden nummer
    assert "¹ Indbetalt kapital indeholder 82.500 DKK i udligningsrente." in noter
    # NAV-afstemningen har sit mærke ved dagsværdien i beholdningstabellen
    assert "Dagsværdi²" in celler and "² Total værdi er dagsværdi plus" in noter


def test_gics_oversigt_viser_alle_sektorer_og_branchegrupper_inden_for_fladen():
    s = slides.gics_oversigt(_prs(), _rapport([]))
    t = _tekster(s)
    for kode, grupper in data.GICS_STRUKTUR.items():
        assert data.SEKTORER[kode] in t and slides.GICS_BESKRIVELSER[kode] in t
        for g in grupper:
            assert f"·  {data.BRANCHEGRUPPER[g]}" in t
    assert sum(x.startswith("·  ") for x in t) == 25
    for spalte in slides._gics_spalter(list(data.GICS_STRUKTUR)):
        hoejde = sum(slides._gics_hoejde(k) + slides.GICS_LUFT for k in spalte)
        assert slides.Y_TOP + hoejde <= slides.Y_NOTER + 0.1


def test_lange_noter_flyttes_op_over_logoet():
    assert slides._noter_y("x" * 300) == slides.Y_NOTER                 # to linjer: som altid
    lang = slides._noter_y("x" * 1000)                                  # fem linjer
    assert lang < slides.Y_NOTER
    assert lang + 5 * slides.NOTE_LINJE <= slides.NOTE_BUND + 1e-9


def test_feederside_viser_gics_sektorer_og_regioner():
    f = _feeder()
    linje = lambda top, under, andel: dict(top_kode=None, top=top, under_kode=None, under=under,
                                           andel=andel, foreloebig=0.0)
    f["brancher"] = dict(dato=date(2026, 6, 30), grundlag="HOLDINGS", gennem="PSCP5", linjer=[
        linje("Industri", "Kapitalgoder", 0.3), linje("Industri", "Erhvervsservice", 0.2),
        linje("Sundhedspleje", "Sundhedsudstyr og -ydelser", 0.5)])
    f["regioner"] = dict(dato=date(2025, 12, 31), grundlag="HOLDINGS", gennem="PSCP5", linjer=[
        linje("Nordamerika", "USA", 0.6), linje("Europa", "Frankrig", 0.4)])
    t = _tekster(slides.fondsside(_prs(), f, "Investor A/S"))
    assert "Industri" in t and "Sundhedspleje" in t and "50,0 %" in t     # sektor lagt sammen
    assert "Nordamerika" in t and "Europa" in t
    assert "USA" not in t and "Frankrig" not in t
    assert any("Branche (GICS-sektor) og geografi er fordelt efter forvalterens fordeling" in x for x in t)
    assert any("Den geografiske fordeling er pr. 31.12.2025." in x for x in t)
    assert not any("Branchefordelingen er pr." in x for x in t)      # samme dato som NAV


def _med_benchmark(r):
    r["benchmark"] = dict(dato=None, indeks={
        "MSCI ACWI": {"INDUSTRIALS": 0.10, "FINANCIALS": 0.16, "ENERGY": 0.74},
        "MSCI World": {"INDUSTRIALS": 0.11, "FINANCIALS": 0.15, "ENERGY": 0.74}})
    return r


def test_brancher_viser_benchmark_pr_sektor():
    r = _med_benchmark(_branche_rapport([
        dict(navn="Industri", dkk=900_000.0),
        dict(navn="Diversificeret (fondsandele)", dkk=100_000.0, kode="DIVERSIFIED")]))
    s = slides.brancher(_prs(), r)
    tabel = next(sh.table for sh in s.shapes if sh.has_table)
    rk = [[c.text for c in row.cells] for row in tabel.rows]
    assert rk[0][-2:] == ["MSCI ACWI", "MSCI World"]
    assert rk[2][1:] == ["Industri", "0,9", "90,0 %", "10,0 %", "11,0 %"]
    assert rk[3][-2:] == [slides.IKKE_OPLYST] * 2              # ingen sektor i indeksene
    assert any("Sektorvægtene i MSCI ACWI og MSCI World;" in x for x in _tekster(s))


def test_brancher_uden_benchmark_har_fire_kolonner():
    s = slides.brancher(_prs(), _branche_rapport([dict(navn="Industri", dkk=900.0)]))
    assert len(next(sh.table for sh in s.shapes if sh.has_table).columns) == 4
    assert not any("Benchmark" in x for x in _tekster(s))


def test_branchekort_viser_indeksene_i_baand_under_kortet():
    r = _med_benchmark(_branche_rapport([dict(navn="Industri", dkk=600_000.0),
                                         dict(navn="Finans", dkk=400_000.0, kode="FINANCIALS")]))
    t = _tekster(slides.branchekort(_prs(), r))
    assert "Industri\n0,6 mio. DKK\n60,0 %" in t                  # kassen uden benchmark
    assert not any("MSCI" in x and "mio. DKK" in x for x in t)
    # ét bånd pr. indeks, med mærket til benchmarknoten ved navnet
    assert "MSCI ACWI¹" in t and "MSCI World¹" in t
    assert any("¹ Sektorvægtene i MSCI ACWI og MSCI World" in x for x in t)
    # Energy har ingen egen kasse: alene står den under sit eget navn, som i kortet
    assert "Energy 74,0 %" in t and "Financials 16,0 %" in t
    assert any("Farverne er kassernes; »Øvrige« er sektorerne uden egen kasse." in x for x in t)


def test_bm_sektorer_samler_sektorer_uden_kasse_i_oevrige():
    r = _branche_rapport([dict(navn="Industri", dkk=900.0)])
    r["benchmark"] = dict(indeks={"MSCI ACWI": {"INDUSTRIALS": 0.2, "ENERGY": 0.3,
                                                "UTILITIES": 0.1, "REAL_ESTATE": 0.4}})
    ((navn, stykker),) = slides._bm_sektorer(r)
    assert navn == "MSCI ACWI"
    assert [(n, round(v, 3)) for n, v, _ in stykker] == [("Industrials", 0.2), ("Øvrige", 0.8)]


def test_branchekortene_viser_samme_benchmarkbaand():
    r = _aktivklasserapport()
    r["benchmark"] = dict(dato=None, indeks={"MSCI ACWI": {
        "INDUSTRIALS": 0.3, "FINANCIALS": 0.5, "REAL_ESTATE": 0.2}})

    def baand(s):
        # stykkernes etiketter i båndet, i rækkefølge fra venstre
        return [sh.text_frame.text for sh in sorted(s.shapes, key=lambda sh: sh.left)
                if sh.has_text_frame and sh.top > Inches(4.5) and " %" in sh.text_frame.text
                and not sh.text_frame.text.startswith("Kilde")]

    hele = baand(slides.branchekort(_prs(), r))
    assert hele and hele == baand(slides.branchekort_aktivklasser(_prs(), r))
    assert any("Real Estate 20,0 %" in x for x in hele)          # indekset, som det er


def test_regionskort_viser_en_kasse_pr_region():
    europa = dict(navn="Europa", kode="EUROPE", dkk=900_000.0, grupper=[
        dict(navn="Danmark", dkk=700_000.0), dict(navn="Frankrig", dkk=190_000.0),
        dict(navn="Benelux", dkk=6_000.0), dict(navn="Sydeuropa", dkk=4_000.0)])
    usa = dict(navn="Nordamerika", kode="NORTH_AMERICA", dkk=99_000.0,
               grupper=[dict(navn="USA", dkk=99_000.0)])
    row = dict(navn="Resten af verden", kode="ROW", dkk=1_000.0,
               grupper=[dict(navn="Asien og Stillehavsområdet", dkk=1_000.0)])
    s = slides.regionskort(_prs(), _branche_rapport([europa, usa, row], art="regioner"))
    t = _tekster(s)
    assert "Geografisk eksponering – hele porteføljen" in t
    assert "Europa\n0,9 mio. DKK\n90,0 %" in t
    assert "Nordamerika\n0,1 mio. DKK\n9,9 %" in t
    # den mindste region er gjort stor nok til en etiket i stedet for at stå i noten
    # forstørrelsen har ingen note (brugerens valg), så kassen har intet mærke
    assert any(x.startswith("Resten af verden\n") for x in t)
    assert not any("forstørret" in x for x in t)
    assert not any(x.startswith(("Danmark", "USA", "Asien")) for x in t)  # ingen delregioner
    # hvert beløb står én gang: kasserne giver 100 %
    procenter = [float(p.replace(",", ".")) for x in t if not x.startswith("Kilde")
                 for p in re.findall(r"(\d+,\d) %", x)]
    assert sum(procenter) == pytest.approx(100.0, abs=0.05)


def _regionsrapport_med_benchmark():
    europa = dict(navn="Europa", kode="EUROPE", dkk=600_000.0,
                  grupper=[dict(navn="Danmark", dkk=600_000.0)])
    usa = dict(navn="Nordamerika", kode="NORTH_AMERICA", dkk=400_000.0,
               grupper=[dict(navn="USA", dkk=400_000.0)])
    r = _branche_rapport([europa, usa], art="regioner")
    r["benchmark"] = dict(dato=None, indeks={}, regioner={
        "MSCI ACWI": {"NORTH_AMERICA": 0.67, "EUROPE": 0.15, "ROW": 0.18},
        "MSCI World": {"NORTH_AMERICA": 0.76, "EUROPE": 0.16, "ROW": 0.08}})
    return r


def test_regionskort_viser_indeksene_i_baand_under_kortet():
    t = _tekster(slides.regionskort(_prs(), _regionsrapport_med_benchmark()))
    assert "Europa\n0,6 mio. DKK\n60,0 %" in t
    assert "MSCI ACWI¹" in t and "MSCI World¹" in t
    assert "Nordamerika 67,0 %" in t and "Resten af verden 18,0 %" in t
    assert any("Regionvægtene i MSCI ACWI og MSCI World." in x for x in t)


def test_regioner_viser_benchmark_pr_region():
    s = slides.regioner(_prs(), _regionsrapport_med_benchmark())
    rk = [[c.text for c in row.cells] for row in next(sh.table for sh in s.shapes
                                                       if sh.has_table).rows]
    assert rk[0][-2:] == ["MSCI ACWI", "MSCI World"]
    assert ["Europa", "15,0 %", "16,0 %"] in [[x[1], *x[-2:]] for x in rk]


def test_regionskort_uden_benchmark_har_ingen_benchmarknote():
    r = _regionsrapport_med_benchmark()
    r["benchmark"] = dict(dato=None, indeks={"MSCI ACWI": {"ENERGY": 1.0}}, regioner={})
    assert not any("MSCI" in x for x in _tekster(slides.regionskort(_prs(), r)))


def test_regionskort_uden_data_springes_over():
    r = _rapport([])
    r["regioner"] = None
    assert slides.regionskort(_prs(), r) is None


def _ejendomsrapport():
    """Hele NAV 1.000.000: en ejendomsfond, en blandet fond med 10 % ejendomme og en PE-fond."""
    linje = lambda kode, top, under, andel: dict(top_kode=kode, top=top, under_kode=None,
                                                 under=under, andel=andel, foreloebig=0.0)
    fond = lambda navn, nav, brancher, regioner, re_sektor=None: dict(
        fund_name=navn, nav_dato=date(2026, 6, 30), n=dict(nav_dkk=nav), re_sektor=re_sektor,
        brancher=dict(grundlag="HOLDINGS", dato=date(2026, 6, 30), gennem=None, linjer=brancher),
        regioner=dict(grundlag="HOLDINGS", dato=date(2026, 6, 30), gennem=None, linjer=regioner))
    fonde = [
        fond("Home.Earth", 300_000.0, [linje("REAL_ESTATE", "Ejendomme", None, 1.0)],
             [linje("EUROPE", "Europa", "Danmark", 1.0)], "RESIDENTIAL"),
        fond("Commercial Real Estate Denmark", 140_000.0,
             [linje("REAL_ESTATE", "Ejendomme", None, 1.0)],
             [linje("EUROPE", "Europa", "Danmark", 1.0)], "OFFICE"),
        fond("K Alternativ Private Equity 2021", 100_000.0,
             [linje("INDUSTRIALS", "Industri", None, 0.95),
              linje("REAL_ESTATE", "Ejendomme", None, 0.05)],
             [linje("NORTH_AMERICA", "Nordamerika", "USA", 1.0)]),
        fond("Seed Capital Denmark IV", 460_000.0,
             [linje("FINANCIALS", "Finans", None, 0.6), linje("ENERGY", "Energi", None, 0.4)],
             [linje("EUROPE", "Europa", "Danmark", 0.7), linje("EUROPE", "Europa", "Frankrig", 0.3)])]
    r = _rapport([])
    r["kurser"] = {}
    r["brancher"] = data.eksponering(fonde, "brancher")
    r["regioner"] = data.eksponering(fonde, "regioner")
    r["uden_ejendomme"] = {art: data.opdel_ejendomme(fonde, art) for art in ("brancher", "regioner")}
    return r


def _procenter(tekster):
    return [float(p.replace(",", ".")) for x in tekster for p in re.findall(r"(\d+,\d) %", x)]


def test_branchekort_ejendomme_deler_sliden_og_regner_paa_hver_dels_base():
    s = slides.branchekort_ejendomme(_prs(), _ejendomsrapport())
    t = [_uden_maerker(x) for x in _tekster(s)]
    assert "Branchefordeling – ejendomme for sig" in t
    # båndene: delenes andel af hele NAV summerer til 100 %
    assert any(x.startswith("Ejendomme  ·  0,4 mio. DKK  ·  44,5 %") for x in t)
    assert any(x.startswith("Alternative investeringer  ·  0,6 mio. DKK  ·  55,5 %") for x in t)
    # de alternative investeringer til venstre, ejendommene til højre
    venstre = {_uden_maerker(sh.text_frame.text.split("  ·  ")[0]): sh.left for sh in s.shapes
               if sh.has_text_frame and "  ·  " in sh.text_frame.text}
    assert venstre["Alternative investeringer"] < venstre["Ejendomme"]
    # ejendomme pr. ejendomssegment, i procent af ejendommene; K Alternativs 5 tDKK har intet
    # segment og står som »Øvrige«, gjort stor nok til navn og procent
    assert "Residential\n0,3 mio. DKK\n67,4 %" in t
    assert any(x.startswith("Office\n0,1 mio. DKK\n31,5 %") for x in t)
    assert "Øvrige\n1,1 %" in t
    assert not any(x.startswith(("Home.Earth", "Commercial")) for x in t)   # ikke pr. fond
    # resten i procent af de alternative investeringer: Finans 276 af 555
    assert any(x.startswith("Finans\n0,3 mio. DKK\n49,7 %") for x in t)
    assert not any(x.startswith("Ejendomme\n") for x in t)
    # hver del summerer til 100 % af sig selv
    ejendomme = ("Residential", "Office", "Øvrige")
    assert sum(_procenter(x for x in t if x.startswith(ejendomme))) == pytest.approx(
        100.0, abs=0.15)
    assert sum(_procenter(x for x in t if "mio. DKK" in x and "  ·  " not in x
                          and not x.startswith((*ejendomme, "Kilde")))) == pytest.approx(
        100.0, abs=0.15)
    # »Øvrige« i ejendomsdelen har sin egen note, som siger, at det er ejendomsdelen
    assert any("»Øvrige« i ejendomsdelen er ejendomsandelen af K Alternativ Private Equity "
               "2021, som ikke kan klassificeres efter ejendomssegment." in x for x in t)
    assert not any("uden egen kasse" in x and "ejendomssegment;" in x for x in t)


def test_regionskort_ejendomme_fordeler_resten_paa_regioner():
    s = slides.regionskort_ejendomme(_prs(), _ejendomsrapport())
    t = _tekster(s)
    assert "Geografisk eksponering – ejendomme for sig" in t
    # resten: Seed 460 i Europa og K Alternativ 95 i Nordamerika, uden delregioner
    assert any(x.startswith("Europa\n0,5 mio. DKK\n82,9 %") for x in t)
    # ejendommene pr. region, ikke pr. fond: Home.Earth og CRED 440 i Europa, K Alternativs
    # ejendomsandel 5 i Nordamerika som resten af fonden
    assert any(x.startswith("Europa\n0,4 mio. DKK\n98,9 %") for x in t)
    assert any(x.startswith("Nordamerika") and "1,1 %" in x for x in t)
    assert not any(x.startswith(("Home.Earth", "Commercial")) for x in t)
    # begge dele summerer til 100 % af sig selv
    procenter = _procenter(x for x in t if "%" in x and "  ·  " not in x
                           and not x.startswith("Kilde"))
    assert sum(procenter) == pytest.approx(200.0, abs=0.3)
    assert not any(x.startswith(("Danmark", "Frankrig")) for x in t)
    assert any("antages at ligge geografisk som resten af fonden" in x for x in t)


def test_ejendomsopdelte_kort_uden_ejendomme_springes_over():
    r = _rapport([])
    r["uden_ejendomme"] = {"brancher": None, "regioner": None}
    assert slides.branchekort_ejendomme(_prs(), r) is None
    assert slides.regionskort_ejendomme(_prs(), r) is None


# ---------- Valutafordeling ----------

def _valuta_rapport():
    fond = lambda navn, valuta, nav, rest: dict(fund_name=navn, valuta=valuta,
                                                n=dict(nav_dkk=nav, resttilsagn_dkk=rest))
    fonde = [fond("A", "DKK", 600_000.0, 100_000.0), fond("B", "EUR", 300_000.0, 300_000.0),
             fond("C", "USD", 100_000.0, 0.0)]
    r = _rapport(fonde)
    r["kurser"] = {"DKK": (1.0, date(2026, 6, 30)), "EUR": (7.46, date(2026, 6, 30))}
    r["valutafordeling"] = data.valutafordeling(fonde)
    return r


def test_valutafordeling_viser_tre_cirkler_i_en_trekant_med_faste_farver():
    s = slides.valutafordeling(_prs(), _valuta_rapport())
    t = _tekster(s)
    assert s.shapes.title.text == "Eksponering"
    # navn og total står i hullet
    assert "Eksponering" in t and "NAV + resttilsagn" in t and "1,4 mio. DKK" in t
    assert "NAV" in t and "1,0 mio. DKK" in t and "Resttilsagn" in t and "0,4 mio. DKK" in t
    diagrammer = {sh.chart.plots[0].series[0].name.split(",")[0]: sh
                  for sh in s.shapes if sh.has_chart}
    assert set(diagrammer) == {"Eksponering", "NAV", "Resttilsagn"}
    assert all(sh.chart.chart_type == XL_CHART_TYPE.DOUGHNUT for sh in diagrammer.values())
    eks, nav, rest = (diagrammer[k] for k in ("Eksponering", "NAV", "Resttilsagn"))
    # trekanten: eksponering øverst, midt over de to andre
    assert eks.top < nav.top == rest.top
    assert nav.left < eks.left < rest.left
    # samme valuta har samme farve i alle tre
    for sh in diagrammer.values():
        plot = sh.chart.plots[0]
        farver = {k: p.format.fill.fore_color.rgb
                  for k, p in zip(plot.categories, plot.series[0].points)}
        assert farver["DKK"] == slides.VALUTA_FARVER["DKK"]
        assert farver["EUR"] == slides.VALUTA_FARVER["EUR"]
    # hullet siger »NAV + resttilsagn«; noten gentager det ikke
    assert not any("Eksponering er NAV plus resttilsagn" in x for x in t)
    assert any("Hver fond indgår med hele sit beløb i fondens valuta" in x for x in t)
    assert not any(sh.has_table for sh in s.shapes)
    # forbindelserne: kun linjerne fra NAV og resttilsagn op til eksponeringen, ingen base
    assert sum(sh.shape_type == MSO_SHAPE_TYPE.LINE for sh in s.shapes) == 2
    assert "+" not in t


def test_valutafordeling_uden_dkk_tal_springes_over():
    r = _rapport([])
    r["valutafordeling"] = None
    assert slides.valutafordeling(_prs(), r) is None


# ---------- Private Equity, Private Credit og ejendomme ----------

def _aktivklasse_fonde():
    """Hele NAV 1.000.000: en ejendomsfond, en PE-fond med 10 % ejendomme og en kreditfond."""
    linje = lambda kode, top, under, andel: dict(top_kode=kode, top=top, under_kode=None,
                                                 under=under, andel=andel, foreloebig=0.0)
    fond = lambda navn, segment, nav, brancher, regioner: dict(
        fund_name=navn, segment=segment, nav_dato=date(2026, 6, 30), n=dict(nav_dkk=nav),
        brancher=dict(grundlag="HOLDINGS", dato=date(2026, 6, 30), gennem=None, linjer=brancher),
        regioner=dict(grundlag="HOLDINGS", dato=date(2026, 6, 30), gennem=None, linjer=regioner))
    return [
        fond("Home.Earth", "Ejendomme", 200_000.0, [linje("REAL_ESTATE", "Ejendomme", None, 1.0)],
             [linje("EUROPE", "Europa", "Danmark", 1.0)]),
        fond("Seed Capital Denmark IV", "Venture", 500_000.0,
             [linje("FINANCIALS", "Finans", None, 0.6), linje("ENERGY", "Energi", None, 0.3),
              linje("REAL_ESTATE", "Ejendomme", None, 0.1)],
             [linje("EUROPE", "Europa", "Danmark", 0.7),
              linje("NORTH_AMERICA", "Nordamerika", "USA", 0.3)]),
        fond("Private Debt III LGT", "Private credit", 300_000.0,
             [linje("HEALTH_CARE", "Sundhedspleje", None, 0.5),
              linje("INDUSTRIALS", "Industri", None, 0.5)],
             [linje("NORTH_AMERICA", "Nordamerika", "USA", 1.0)])]


def _aktivklasserapport():
    fonde = _aktivklasse_fonde()
    r = _rapport([])
    r["kurser"] = {}
    r["brancher"] = data.eksponering(fonde, "brancher")
    r["regioner"] = data.eksponering(fonde, "regioner")
    r["aktivklasser"] = {art: data.opdel_aktivklasser(fonde, art)
                         for art in ("brancher", "regioner")}
    return r


def _baand(s):
    """{delens navn: (venstrekant, tekst)} for båndene over delene."""
    return {_uden_maerker(sh.text_frame.text.split("  ·  ")[0]):
            (sh.left, _uden_maerker(sh.text_frame.text))
            for sh in s.shapes if sh.has_text_frame and "  ·  " in sh.text_frame.text
            and " mio. DKK  ·  " in sh.text_frame.text}


def test_aktivklasser_deler_i_tre_dele_i_raekkefoelge():
    for byg in (slides.branchekort_aktivklasser, slides.regionskort_aktivklasser):
        s = byg(_prs(), _aktivklasserapport())
        baand = _baand(s)
        assert list(sorted(baand, key=lambda n: baand[n][0])) == [
            "Private Equity", "Private Credit", "Ejendomme"]
        # PE: Seed uden ejendomme 450; PC: 300; ejendomme: Home.Earth 200 + Seeds 50
        assert baand["Private Equity"][1].startswith("Private Equity  ·  0,5 mio. DKK  ·  45,0 %")
        assert baand["Private Credit"][1].startswith("Private Credit  ·  0,3 mio. DKK  ·  30,0 %")
        assert baand["Ejendomme"][1].startswith("Ejendomme  ·  0,2 mio. DKK  ·  25,0 %")
        assert any("Private Credit er fondene med private credit" in x for x in _tekster(s))


def test_branchekort_aktivklasser_regner_hver_del_paa_sin_egen_base():
    t = _tekster(slides.branchekort_aktivklasser(_prs(), _aktivklasserapport()))
    # PE: Finans 300 af 450, Energi 150 af 450; PC: Sundhedspleje og Industri 150 af 300 hver
    assert any(x.startswith("Finans\n0,3 mio. DKK\n66,7 %") for x in t)
    assert any(x.startswith("Sundhedspleje\n0,1 mio. DKK\n50,0 %") for x in t)
    # ejendommene pr. segment: Home.Earth (uden segment i fixturen) og Seeds 50 er »Øvrige«
    # »Øvrige« i ejendomsdelen har sit eget fodnotemærke
    assert any(re.fullmatch("Øvrige[¹²³⁴⁵⁶⁷⁸⁹]+\n0,2 mio. DKK\n100,0 %", x) for x in t)


def test_aktivklasser_uden_data_springes_over():
    r = _rapport([])
    r["aktivklasser"] = {"brancher": None, "regioner": None}
    assert slides.branchekort_aktivklasser(_prs(), r) is None
    assert slides.regionskort_aktivklasser(_prs(), r) is None


# ---------- faste farver og navne i kasserne ----------

def _kassefarver(s):
    """{første linje i kassens etiket: kassens farve} - etiketten ligger oven på kassen."""
    kasser = [sh for sh in s.shapes if sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
              and sh.fill.type == 1 and not sh.text_frame.text]        # 1: massiv fyldfarve
    return {t.text_frame.text.split("\n")[0]: k.fill.fore_color.rgb
            for t in s.shapes if t.shape_type == MSO_SHAPE_TYPE.TEXT_BOX and t.text_frame.text
            for k in kasser
            if k.left <= t.left < k.left + k.width and k.top <= t.top < k.top + k.height}


def test_samme_sektor_har_samme_faste_farve_i_alle_dele():
    fonde = _aktivklasse_fonde()
    # kreditfonden får også Finans, så sektoren står i både Private Equity og Private Credit
    fonde[2]["brancher"]["linjer"][1] |= dict(top_kode="FINANCIALS", top="Finans")
    r = _aktivklasserapport()
    r["aktivklasser"]["brancher"] = data.opdel_aktivklasser(fonde, "brancher")
    s = slides.branchekort_aktivklasser(_prs(), r)
    farver = [sh.fill.fore_color.rgb for sh in s.shapes
              if sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and sh.fill.type == 1]
    finans = slides.SEKTOR_FARVER["FINANCIALS"]
    assert farver.count(finans) == 2                    # én kasse i hver del
    assert slides.SEKTOR_FARVER["HEALTH_CARE"] in farver
    # ejendomssegmenterne har deres egne farver, ingen af dem en sektors
    assert not set(slides.EJD_SEGMENT_FARVER.values()) & set(slides.SEKTOR_FARVER.values())
    # gråt kun til »Øvrige«: ingen sektor har den farve
    assert slides.BR_OEVRIGE not in slides.SEKTOR_FARVER.values()
    assert len(set(slides.SEKTOR_FARVER.values())) == len(slides.SEKTOR_FARVER)


def test_en_enkelt_lille_sektor_beholder_sin_farve_og_kun_oevrige_er_graa():
    poster = [("Industrials", 900.0, 0.9, slides.SEKTOR_FARVER["INDUSTRIALS"]),
              ("Materials", 10.0, 0.01, slides.SEKTOR_FARVER["MATERIALS"])]
    skiver, _, _ = slides._skiver(poster)
    assert skiver[-1] == ("Materials", 10.0, 0.01, slides.SEKTOR_FARVER["MATERIALS"])
    to_smaa = poster + [("Energy", 5.0, 0.005, slides.SEKTOR_FARVER["ENERGY"])]
    skiver, _, _ = slides._skiver(to_smaa)
    assert skiver[-1] == ("Øvrige", 15.0, 0.015, slides.BR_OEVRIGE)


def test_kasse_uden_plads_til_navnet_faar_fodnotemaerke_og_note():
    s = _ny_slide(_prs(), L_TEKST)
    fod = slides.Fodnoter()
    assert not slides._kortkasse(s, (1.0, 1.0, 0.3, 0.2), "Communication Services", 1_321_000.0,
                                 0.025, slides.SAGE, fod)
    assert any(x == "¹ 2,5 %" for x in _tekster(s))
    assert fod.tekst() == "¹ Communication Services, 1,3 mio. DKK."
    # forstørret og uden navn: én note, ikke to mærker
    fod = slides.Fodnoter()
    slides._kortkasse(s, (1.0, 1.0, 0.3, 0.2), "Materials", 528_000.0, 0.01, slides.SAGE, fod,
                      loeftet=True)
    assert fod.tekst() == "¹ Materials, 0,5 mio. DKK."


def test_fodnoter_nummererer_i_raekkefoelge_og_genbruger_samme_tekst():
    fod = slides.Fodnoter("Kilde: X.")
    assert fod.maerke("A") == "¹" and fod.maerke("B") == "²" and fod.maerke("A") == "¹"
    assert fod.maerke(None) == "" and fod.maerker(["B", "C"]) == "²³"
    fod.generel("Generel.", ["Generel.", None])
    assert fod.tekst() == "Kilde: X. Generel. ¹ A ² B ³ C"
    for i in range(4, 12):
        fod.maerke(f"note {i}")
    assert fod.maerke("note 11") == "¹¹"


# ---------- største positioner ----------

def _positionsrapport():
    sektion = {"selskaber": "CURRENT", "fonde": "FUND", "ejendomme": "PROPERTY"}
    fond = lambda navn, valuta, nav, raekker, type_="selskaber": dict(
        fund_name=navn, valuta=valuta, nav_dato=date(2026, 6, 30), n=dict(nav_dkk=nav),
        beholdninger=dict(type=type_, skaleret=True, dato=date(2026, 6, 30), via=None,
                          raekker=[dict(navn=n, dagsvaerdi=v, type=sektion[type_])
                                   for n, v in raekker]))
    fonde = [fond("Seed Capital Denmark IV", "DKK", 500_000.0,
                  [("FLATPAY", 200_000.0), ("Landfolk", 50_000.0)]),
             fond("K Alternativ Private Equity 2021 K/S", "DKK", 300_000.0,
                  [("ACF VIII", 120_000.0)], type_="fonde"),
             fond("Home.Earth", "EUR", 200_000.0, [("Hus", 10_000.0)], type_="ejendomme")]
    r = _rapport(fonde)
    r["kurser"] = {}
    r["stoerste_positioner"] = data.stoerste_positioner(
        fonde, {"DKK": (1.0, None), "EUR": (7.46, None)})
    return r


def test_stoerste_positioner_viser_rangeret_tabel_med_fondslinjer():
    s = slides.stoerste_positioner(_prs(), _positionsrapport())
    t = _tekster(s)
    assert "Største positioner" in t
    tabel = next(sh.table for sh in s.shapes if sh.has_table)
    raekker = [[tabel.cell(i, j).text for j in range(6)] for i in range(len(tabel.rows))]
    # samme typer som på Største værdibevægelser (data.beholdningstype)
    assert raekker[1] == ["1", "FLATPAY", "Selskab", "SCD IV", "0,2", "20,0 %"]
    assert raekker[2][1:4] == ["Home.Earth", "Fond (ikke gennemlyst)³", "–"]
    assert raekker[3][1:4] == ["ACF VIII", "Fond", "KAPE 2021"]
    assert raekker[-1][1] == "De 4 største i alt" and raekker[-1][4] == "0,6"
    # fondene står i rækkerne, så noten nævner dem ikke
    assert any("Fonden har ingen beholdninger, der kan gennemlyses" in x for x in t)
    assert any("samme position i flere fonde er lagt sammen" in x for x in t)


def test_stoerste_positioner_uden_data_springes_over():
    r = _rapport([])
    r["stoerste_positioner"] = None
    assert slides.stoerste_positioner(_prs(), r) is None


def test_likviditet_viser_fem_tal_og_tre_diagrammer():
    lq = data.likviditet([_lq_fond("Alfa", "Private equity"), _lq_fond("Ejendom", "Ejendomme")],
                         lambda v, d: LQ_KURSER[d], date(2026, 6, 30))
    r = dict(navn="Investor", as_of=date(2026, 6, 30), likviditet=lq,
             datotjek=dict(ens=True, foraeldede=[], skaeve_beholdninger=[]))
    s = slides.likviditet(_prs(), r)
    t = _tekster(s)
    for label in ("Markedsværdi", "Afkast¹", "Indbetalinger", "Udbetalinger", "Netto cash flow"):
        assert label in t
    assert sum(sh.has_chart for sh in s.shapes) == 3
    noter = next(x for x in t if x.startswith("Kilde"))
    assert "Ejendomsfondene (Ejendom)" in noter and "2026 er pr. 30.06.2026" in noter
    # J-kurven: søjler og en flade, der deler akserne
    j = next(sh.chart for sh in s.shapes if sh.has_chart)
    plot_area = j._chartSpace.chart.plotArea
    assert plot_area.find(qn("c:areaChart")) is not None
    # kurven er en linje for sig, så fladens kant ikke tegner lodret op til nul i enderne
    assert plot_area.find(qn("c:lineChart")) is not None
    assert plot_area.find(qn("c:areaChart")).find(".//" + qn("a:ln")).find(qn("a:noFill")) is not None
    assert slides.likviditet(_prs(), r | dict(likviditet=None)) is None


def test_likviditet_med_ejendomme_siger_det_i_undertitlen():
    fonde = [_lq_fond("Alfa", "Private equity"), _lq_fond("Ejendom", "Ejendomme")]
    r = dict(navn="Investor", as_of=date(2026, 6, 30),
             likviditet_alle=data.likviditet(fonde, lambda v, d: LQ_KURSER[d], date(2026, 6, 30),
                                             uden=None),
             datotjek=dict(ens=True, foraeldede=[], skaeve_beholdninger=[]))
    t = _tekster(slides.likviditet(_prs(), r, med_ejendomme=True))
    assert any("hele porteføljen inkl. ejendomsfonde" in x for x in t)
    assert not any("indgår ikke" in x for x in t)


def test_likviditet_soejlerne_har_etiketter_og_forklaring():
    lq = data.likviditet([_lq_fond("Alfa", "Private equity")], lambda v, d: LQ_KURSER[d],
                         date(2026, 6, 30))
    r = dict(navn="Investor", as_of=date(2026, 6, 30), likviditet=lq,
             datotjek=dict(ens=True, foraeldede=[], skaeve_beholdninger=[]))
    with slides.beloebsskala(1_000_000):        # beløbene i mio., så de får etiketter
        s = slides.likviditet(_prs(), r)
    tilsagn, tvpi = [sh.chart for sh in s.shapes if sh.has_chart][1:]
    etiketter = lambda serie: [p.data_label.text_frame.text for p in serie.points]
    rvpi, _, total = tvpi.plots[0].series
    assert etiketter(total) == [slides.gange(a["tvpi"]) for a in lq["aar"]]
    assert etiketter(rvpi) == [slides.gange(a["rvpi"]) for a in lq["aar"]]
    # 2024: resttilsagn 60 EUR og kaldt 40 EUR til 7,2
    rest, kaldt = tilsagn.plots[0].series
    assert (etiketter(rest)[0], etiketter(kaldt)[0]) == ("432 mio.", "288 mio.")
    # forklaring på begge; den usynlige totalserie er slettet fra TVPI-forklaringen
    assert tilsagn.has_legend and tvpi.has_legend
    slettet = tvpi.legend._element.findall(qn("c:legendEntry"))
    assert [e.find(qn("c:idx")).get("val") for e in slettet] == ["2"]
    assert not tilsagn.legend._element.findall(qn("c:legendEntry"))


def _transaktionsrapport():
    fonde = [_lq_fond("Alfa", "Private equity"), _lq_fond("Beta Fund II", "Private credit")]
    kurs = lambda v, d: LQ_KURSER.get(d, 7.46)
    for f in fonde:
        f["tvpi_historik"] = [dict(dato=date(2025, 12, 31), tvpi=1.0, dpi=0.0, afvigende=False),
                              dict(dato=date(2026, 3, 31), tvpi=4.0, dpi=0.0, afvigende=True),
                              dict(dato=date(2026, 6, 30), tvpi=1.2, dpi=0.1, afvigende=False)]
        f["n"]["tvpi"] = 1.2
    return dict(navn="Investor", as_of=date(2026, 6, 30), fonde=fonde,
                kurser={"EUR": (7.46, date(2026, 6, 30))},
                aarlige_stroemme=data.aarlige_stroemme(fonde, kurs, date(2026, 6, 30)),
                kaldstempo=data.kaldstempo(fonde, kurs, date(2026, 6, 30)),
                datotjek=dict(ens=True, foraeldede=[], skaeve_beholdninger=[]))


def test_aarlige_stroemme_har_soejler_pr_fond_og_nettolinje():
    s = slides.aarlige_stroemme(_prs(), _transaktionsrapport())
    (diagram,) = [sh.chart for sh in s.shapes if sh.has_chart]
    plot_area = diagram._chartSpace.chart.plotArea
    assert plot_area.find(qn("c:lineChart")) is not None
    # én kaldserie og én udlodningsserie pr. fond; udlodningerne er skjult i forklaringen
    assert len(plot_area.find(qn("c:barChart")).findall(qn("c:ser"))) == 4
    assert len(diagram.legend._element.findall(qn("c:legendEntry"))) == 2
    assert slides.aarlige_stroemme(_prs(), _transaktionsrapport() | dict(
        aarlige_stroemme=None)) is None


def test_tvpi_over_tid_et_panel_pr_fond_uden_afvigende_kvartaler():
    s = slides.tvpi_over_tid(_prs(), _transaktionsrapport())
    diagrammer = [sh.chart for sh in s.shapes if sh.has_chart]
    assert len(diagrammer) == 2
    tvpi = diagrammer[0].plots[0].series[0]
    assert list(tvpi.values) == [1.0, None, 1.2]
    noter = next(x for x in _tekster(s) if x.startswith("Kilde"))
    assert "Ikke vist: Alfa pr. 31.03.2026 (4,00x)" in noter


def test_eksponering_over_tid_fliser_flader_linje_og_kald_pr_kvartal():
    kurs = lambda v, d: LQ_KURSER.get(d, 7.5)
    e = data.eksponering_over_tid([_lq_fond("Alfa", "Private equity")], kurs, date(2026, 6, 30))
    r = dict(navn="Investor", as_of=date(2026, 6, 30), eksponering_over_tid=e,
             datotjek=dict(ens=True, foraeldede=[], skaeve_beholdninger=[]))
    s = slides.eksponering_over_tid(_prs(), r)
    t = _tekster(s)
    for label in ("Eksponering", "NAV", "Resttilsagn", "Kaldt kapital i alt",
                  "Kaldt seneste 12 mdr."):
        assert label in t
    # den første flise er NAV + resttilsagn i dag
    assert t[t.index("Eksponering") + 1] == slides.beloeb(522.2 + 298.4)
    flader, kald = [sh.chart for sh in s.shapes if sh.has_chart]
    plot_area = flader._chartSpace.chart.plotArea
    # NAV og resttilsagn stablet, den kaldte kapital som linje på samme akser
    assert len(plot_area.find(qn("c:areaChart")).findall(qn("c:ser"))) == 2
    assert plot_area.find(qn("c:lineChart")) is not None
    assert len(list(kald.plots[0].series[0].values)) == len(e["kvartaler"]) == 10
    assert not kald.has_title and not kald.has_legend
    # en etiket pr. kvartal
    assert list(kald.plots[0].categories)[:5] == ["Q1 24", "Q2 24", "Q3 24", "Q4 24", "Q1 25"]
    noter = next(x for x in t if x.startswith("Kilde"))
    assert "Før den første NAV indgår Alfa" in noter
    _indenfor_fladen(s, "Kapitalkald, resttilsagn og eksponering over tid")
    assert slides.eksponering_over_tid(_prs(), r | dict(eksponering_over_tid=None)) is None


def test_kaldstempo_tabel_med_total():
    s = slides.kaldstempo(_prs(), _transaktionsrapport())
    (tabel,) = [sh.table for sh in s.shapes if sh.has_table]
    assert tabel.cell(len(tabel.rows) - 1, 0).text == "I alt"
    assert tabel.cell(1, 4).text == "n/a"           # ingen kald de seneste 12 måneder


# ---------- IRR mod Aleta ----------

def _aleta_fond(navn, valuta, aarsager, metode_ok=True):
    return {"fund_name": navn, "valuta": valuta, "aleta": {
        "irr": 0.10, "irr_dkk": 0.101, "irr_aletas_datoer": 0.0975, "aleta": 0.0974,
        "forskel": 0.0036, "aarsager": aarsager, "metode_ok": metode_ok,
        "dato": date(2026, 7, 1)}}


def test_irr_mod_aleta():
    aarsag = "2 betalinger har en tidligere dato hos Aleta"
    r = _rapport([_aleta_fond("Alfa Fund I K/S", "EUR", [aarsag]),
                  {"fund_name": "Beta Fund II A/S", "valuta": "DKK"}])
    s = slides.irr_mod_aleta(_prs(), r)
    raekker = _tabelraekker(s)
    assert len(raekker) == 2                # Beta har ingen Aleta-tal
    assert raekker[0] == ["Fond", "Valuta", "Rapporten,\vfondens valuta", "Rapporten,\vDKK",
                          "Aleta,\vDKK", "Forskel,\v%-point", "", "Årsag til forskellen"]
    assert raekker[1][:6] == ["Alfa Fund I K/S", "EUR", "10,00 %", "10,10 %", "9,74 %", "+0,36"]
    assert raekker[1][7].startswith("2 betalinger")
    t = " ".join(_tekster(s))
    assert "(eftervist for alle fonde)" in t and "kaldsbrev" not in t
    assert "datoen, som fremgår af banken" in t and "deposit notice og ejerbog" in t
    assert "forskellen er mellem de to IRR i DKK" in t
    assert ("For fondene i EUR er forskellen mellem rapportens IRR i fondens valuta og i DKK "
            "valutakursens udvikling" in t and "kronen er bundet til euroen" in t)
    assert "under 0,05 %-point for 1 af 1 fonde" in t
    _indenfor_fladen(s, "IRR sammenlignet med Aleta")


def test_aleta_afstemning_viser_status_med_irr_uden_vintage_og_transaktioner():
    status = {k: "ok" for k, _ in slides.ALETA_AFSTEM_KOL} | {
        "Vintage": "afviger", "NAV DKK": "lille afvigelse", "TVPI": "afviger"}
    alfa = _aleta_fond("Alfa Fund I K/S", "EUR", []) | {"aleta_kontrol": status}
    r = _rapport([alfa, {"fund_name": "Beta Fund II A/S", "valuta": "DKK"}])
    s = slides.aleta_afstemning(_prs(), r)
    raekker = _tabelraekker(s)
    assert len(raekker) == 2                # Beta har ingen Aleta-tal
    hoved = [x.replace("\v", " ") for x in raekker[0]]
    assert hoved[0] == "Fond" and len(hoved) == 12 and hoved[-1] == "IRR"
    assert not any(k in x for x in hoved for k in ("Vintage", "Transaktioner", "Valuta"))
    assert raekker[1][0] == "Alfa Fund I K/S"
    assert set(raekker[1][1:]) == {"●"}
    # prikkens farve viser status: NAV DKK (kolonne 2) er inden for, TVPI (sidste) uden for
    (tabel,) = [sh.table for sh in s.shapes if sh.has_table]
    farve = lambda j: tabel.cell(1, j).text_frame.paragraphs[0].runs[0].font.color.rgb
    assert farve(1) == slides.ALETA_MARKERING["ok"][0]
    assert farve(2) == slides.ALETA_MARKERING["lille afvigelse"][0]
    assert farve(10) == slides.ALETA_MARKERING["afviger"][0]
    t = " ".join(_tekster(s))
    assert "Aletas tal pr. 01.07.2026" in t and "NAV, tilsagn og resttilsagn i EUR" in t
    assert "Beløb" in t and "Nøgletal" in t                 # gruppeoverskrifterne
    # grænserne står ikke på sliden; valutapointen kun med en afvigelse i indbetalt/udloddet
    assert "0,5 %" not in t and "valutakurseffekter" not in t
    assert "Ens (9 af 11)" in t and "Større afvigelse (1 af 11)" in t
    alfa["aleta_kontrol"] = status | {"Indbetalt DKK": "lille afvigelse"}
    t = " ".join(_tekster(slides.aleta_afstemning(_prs(), r)))
    assert "valutakurseffekter i indbetalt og udloddet kapital" in t
    _indenfor_fladen(s, "Nøgletal afstemt med Aleta")
    assert slides.aleta_afstemning(_prs(), _rapport([{"fund_name": "Beta", "valuta": "DKK"}])) is None


def test_irr_mod_aleta_uden_aleta():
    assert slides.irr_mod_aleta(_prs(), _rapport([{"fund_name": "Beta", "valuta": "DKK"}])) is None


def _punkt(dato, fv, kost, udl=0.0):
    return dict(dato=dato, rapport=dato, dagsvaerdi=fv, kostpris=kost, udloddet=udl, tvpi=None)


def test_bevaegelser_i_dybden_et_kort_pr_bevaegelse():
    r = _udv_rapport()
    datoer = data.kvartalsdatoer(date(2026, 6, 30), 3)
    alfa = r["fonde"][0]
    alfa["beholdninger"] = dict(raekker=[dict(navn="Selskab A", closing=date(2023, 1, 15),
                                              vintage=None, sektor=None)])
    alfa["virksomheder"] = {"Selskab A": dict(beskrivelse="Laver maling.")}
    kort = [dict(navn="Selskab A", fond="Alfa Fund I K/S", type="Selskab", kvartal=3e5, aar=4e5,
                 bidrag=0.015, primo=_punkt(datoer[1], 1e6, 5e5),
                 ultimo=_punkt(datoer[2], 1.3e6, 5e5),
                 dele=dict(dagsvaerdi=3e5, udloddet=0.0, kostpris=0.0),
                 forloeb=[None, _punkt(datoer[1], 1e6, 5e5), _punkt(datoer[2], 1.3e6, 5e5)]),
            dict(navn="Fond B", fond="Alfa Fund I K/S", type="Fond", kvartal=-1e5, aar=None,
                 bidrag=None, primo=_punkt(datoer[1], 1e6, 8e5),
                 ultimo=_punkt(datoer[2], 1e6, 9e5),
                 dele=dict(dagsvaerdi=0.0, udloddet=0.0, kostpris=1e5),
                 forloeb=[_punkt(d, 1e6, 8e5) for d in datoer])]
    r["bevaegelser_i_dybden"] = dict(kort=kort, afkast=0.02, fra=datoer[1], til=datoer[2],
                                     datoer=datoer)
    s = slides.bevaegelser_i_dybden(_prs(), r)
    t = _tekster(s)
    assert "Bag de største værdibevægelser" in t
    assert "+0,3" in t and "-0,1" in t
    assert "Selskab · AF I · Closing 01.2023" in t
    assert any("Bidrag til porteføljens afkast: +1,50 procentpoint¹" in x for x in t)
    tekst_a = next(x for x in t if x.startswith("Forvalteren har opskrevet"))
    assert "30,0 % af dagsværdien primo." in tekst_a and tekst_a.endswith("Laver maling.")
    assert any("renset for investeringer og udlodninger" in x for x in t)
    # kvartalerne under søjlerne, også før beholdningen fandtes
    assert t.count("Q4 25") == 2
    _indenfor_fladen(s, "Bag de største værdibevægelser")
    assert slides.bevaegelser_i_dybden(_prs(), _udv_rapport()) is None


def test_koncentration_fliser_og_et_kort_pr_position():
    r = _udv_rapport()
    datoer = data.kvartalsdatoer(date(2026, 6, 30), 2)
    alfa = r["fonde"][0]
    alfa["beholdninger"] = dict(raekker=[dict(navn="Selskab A", closing=None, vintage=2022)])

    def kort(navn, art, fonde, dkk, andel, multiple, siden=None, **kw):
        return dict(navn=navn, art=art, type="Selskab", fonde=fonde, via=None, dkk=dkk,
                    andel=andel, kostpris=dkk / multiple, multiple=multiple, aendring_12m=0.1,
                    uaendret_siden=siden, uaendret_hele=False, fald=andel * 0.25,
                    andel_i_fonden=None if art == "fond" else 0.5,
                    forloeb=[_punkt(d, dkk, dkk / multiple) for d in datoer], **kw)
    r["koncentration"] = dict(
        kort=[kort("Selskab A", "porteføljeselskab", ["Alfa Fund I K/S"], 5e6, 0.25, 4.0,
                   siden=date(2025, 12, 31)),
              kort("Beta Fund II A/S", "fond", ["Beta Fund II A/S"], 2e6, 0.10, 1.2)],
        nav_dkk=2e7, datoer=datoer, stoerste=0.25, top5=0.5, antal_top5=5, top=0.7,
        antal_top=15, store=3)
    s = slides.koncentration(_prs(), r)
    t = _tekster(s)
    assert "Største positioner – koncentration og risiko" in t
    for label, v in (("Største position", "25,0 %"), ("5 største", "50,0 %"),
                     ("15 største", "70,0 %"), ("Positioner over 5 % af NAV", "3")):
        assert t[t.index(label) + 1] == v
    assert "Selskab · AF I · Vintage 2022" in t
    assert "Fond (ikke gennemlyst)¹" in t
    celler = [c.text for sh in s.shapes if sh.has_table for rk in sh.table.rows for c in rk.cells]
    assert "31.12.2025" in celler and "i kvartalet" in celler and "2,5 % af NAV" in celler
    assert "Andel af fonden²" in celler and celler.count("–") == 1     # ikke for fonden
    # en høj multiple: værdien er mest urealiseret gevinst
    assert sum("urealiseret gevinst" in x for x in t) == 1
    _indenfor_fladen(s, "Største positioner – koncentration og risiko")
    assert slides.koncentration(_prs(), _udv_rapport()) is None


def test_ejendomme_i_dybden_fliser_soejler_og_tabel_ned_til_nav():
    ekst = "property based on third party valuation"
    rapporter = [
        dict(dato=date(2025, 6, 30), fakta=dict(FUND_NAV=40e6, NAV_PER_UNIT=0.6), raekker=[
            dict(navn="Ejendom A", dagsvaerdi=60e6, sektor="Ekstern vurdering",
                 grundlag_koder=[ekst])]),
        dict(dato=date(2026, 6, 30), fakta=dict(FUND_NAV=50e6, NAV_PER_UNIT=0.66), raekker=[
            dict(navn="Ejendom A", dagsvaerdi=70e6, sektor="Ekstern vurdering",
                 grundlag_koder=[ekst]),
            dict(navn="Ejendom B", dagsvaerdi=30e6, sektor="Projektomkostninger",
                 grundlag_koder=["project costs to date"])])]
    f = dict(fund_name="Bolig K/S", valuta="EUR", nav_dato=date(2026, 6, 30),
             n=dict(andel=0.05), ejendomsforloeb=data.ejendomsforloeb(rapporter))
    r = _rapport([f])
    s = slides.ejendomme_i_dybden(_prs(), r)[0]
    t = _tekster(s)
    assert "Bolig K/S – ejendommene bag NAV" in t
    for label, v in (("GAV", "100,0"), ("Fondens NAV", "50,0"), ("NAV i % af GAV", "50,0 %"),
                     ("NAV pr. andel, 12 mdr.", "+10,0 %"),
                     ("Investor A/Ss andel af NAV", "5,00 %")):
        assert t[t.index(label) + 1] == v
    celler = [[c.text for c in rk.cells] for sh in s.shapes if sh.has_table
              for rk in sh.table.rows]
    assert celler[1] == ["Ejendom A", "Udlejning", "Ekstern vurdering", "70,0", "+10,0", "70,0 %"]
    assert celler[2][1] == "Under opførelse" and celler[2][4] == "Ny"
    assert celler[-3][:5] == ["Ejendomme (GAV)", "", "", "100,0", "+40,0"]
    assert celler[-2][0] == "Gæld m.m., netto²" and celler[-2][3] == "-50,0"
    assert celler[-1][:5] == ["Fondens NAV", "", "", "50,0", "+10,0"]
    diagram = next(sh.chart for sh in s.shapes if sh.has_chart)
    assert [x.name for p in diagram.plots for x in p.series] == [
        "Ejendom A", "Ejendom B", "Fondens NAV"]
    _indenfor_fladen(s, "Bolig K/S – ejendommene bag NAV")
    assert slides.ejendomme_i_dybden(_prs(), _rapport([dict(f, ejendomsforloeb=None)])) is None


def test_ejendomme_i_dybden_med_selskabsvaerdi_og_nettogaeld():
    ekst = "property based on third party valuation"

    def rapport(dato, a, gav, pev, nd, nav):
        return dict(dato=dato, raekker=[dict(navn="Ejendom A", dagsvaerdi=a,
                                             sektor="Ekstern vurdering", grundlag_koder=[ekst])],
                    fakta=dict(GROSS_ASSET_VALUE=gav, ENTITY_VALUE=pev, NET_DEBT=nd,
                               CASH=5e6, FUND_NAV=nav, NAV_PER_UNIT=0.7))
    e = data.ejendomsforloeb([rapport(date(2025, 6, 30), 80e6, 80e6, 10e6, 40e6, 50e6),
                              rapport(date(2026, 6, 30), 99.96e6, 100e6, 20e6, 60e6, 60e6)])
    f = dict(fund_name="Bolig K/S", valuta="EUR", nav_dato=date(2026, 6, 30),
             n=dict(andel=None), ejendomsforloeb=e)
    s = slides.ejendomme_i_dybden(_prs(), _rapport([f]))[0]
    t = _tekster(s)
    assert t[t.index("Belåningsgrad (LTV)¹") + 1] == "60,0 %"
    assert "Investor A/Ss andel af NAV" not in t
    celler = [[c.text for c in rk.cells] for sh in s.shapes if sh.has_table
              for rk in sh.table.rows]
    # forvalterens GAV; ejendommene summerer til 100,0 også, så ingen afrundingsnote
    assert [rk[0] for rk in celler[-4:]] == ["Ejendomme (GAV)", "Selskabsværdi (PEV)³",
                                            "Nettogæld⁴", "Fondens NAV"]
    assert [rk[3:5] for rk in celler[-4:]] == [["100,0", "+20,0"], ["20,0", "+10,0"],
                                              ["-60,0", "-20,0"], ["60,0", "+10,0"]]
    assert "fratrukket likvider (5,0)" in " ".join(t)
    diagram = next(sh.chart for sh in s.shapes if sh.has_chart)
    assert [x.name for p in diagram.plots for x in p.series][-2:] == ["Fondens NAV", "Nettogæld"]
    # en rest ud over afrunding står som Øvrige poster
    rest = slides._ejendomsbro(dict(e, fonds_nav=65e6), slides.Fodnoter())
    assert rest[-1] == ("Øvrige poster", 5e6, None)


def _pme_rapport():
    a = _tpt_fond("Fond A", 100.0)
    a["n"].update(ks_pme=1.18, direct_alpha=0.034, irr=0.12)
    a["pme"] = dict(ks_pme=1.18, direct_alpha=0.034, indeks_irr=0.09, merverdi=12_000_000.0)
    b = _tpt_fond("Fond B", 100.0)
    b["n"].update(ks_pme=0.80, direct_alpha=-0.06, irr=0.05)
    b["pme"] = dict(ks_pme=0.80, direct_alpha=-0.06, indeks_irr=0.14, merverdi=-3_000_000.0)
    gammel = _tpt_fond("Fond C", 100.0)
    gammel["n"].update(ks_pme=None, direct_alpha=None)
    gammel["pme"] = dict(grund="Fondens første cash flow (31.03.2005) ligger før indeksets start.")
    r = _tpt_rapport([gammel, b, a])
    r["pme"] = dict(indeks="MSCI ACWI (iShares ETF)", ks_pme=0.94, direct_alpha=-0.025,
                    irr=0.139, indeks_irr=0.168, merverdi=-19_800_000.0, udeladte=["Fond C"],
                    indeks_kvartal=0.142,
                    forloeb=[dict(dato=date(2026, 3, 31), nav=266e6, indekskonto=254e6,
                                  netto_indbetalt=193e6),
                             dict(dato=date(2026, 6, 30), nav=285e6, indekskonto=305e6,
                                  netto_indbetalt=205e6)])
    return r


def test_pme_fonde_sorteret_efter_direct_alpha_med_i_alt_nederst():
    s = slides.pme_fonde(_prs(), _pme_rapport())
    t = _tekster(s)
    assert "Fondene mod markedet" in t and "KS-PME" in t and "Merværdi\vmod ACWI³" in t
    assert "IRR: fonden mod de samme cash flows i ACWI²" in t
    raekker = [x for x in t if x.startswith(("Fond A", "Fond B", "Fond C", "I alt"))]
    assert [x[:6] for x in raekker] == ["Fond A", "Fond B", "Fond C", "I alt"]
    assert "Fond C¹" in t
    assert "1,18x" in t and "+3,4 %" in t and "+12,0 mio. DKK" in t and "-3,0 mio. DKK" in t
    assert "0,94x" in t and "-2,5 %" in t and "-19,8 mio. DKK" in t
    assert "12,0 %" in t and "9,0 %" in t and "13,9 %" in t and "16,8 %" in t
    noter = next(x for x in t if x.startswith("Kilde"))
    assert "MSCI ACWI (iShares ETF)" in noter and "før indeksets start" in noter
    assert "Indekset steg 14,2 % i DKK i seneste kvartal" in noter
    # siges én gang: ved rækken, ikke også som generel note
    assert "Fond C er ikke med" not in noter
    assert "heller ikke med i i alt-rækken" in noter


def test_pme_portefoelje_viser_tal_og_kurver():
    s = slides.pme_portefoelje(_prs(), _pme_rapport())
    t = _tekster(s)
    assert "Porteføljen mod markedet" in t
    assert "-19,8 mio. DKK" in t and "0,94x" in t and "-2,5 %" in t and "16,8 %" in t
    assert "Merværdi mod ACWI¹" in t
    noter = next(x for x in t if x.startswith("Kilde"))
    assert "¹ Merværdi mod ACWI: fondens NAV minus indekskontoens værdi" in noter
    assert "IRR med ACWI²" in t and "² IRR med ACWI er IRR af de samme kapitalkald" in noter
    assert "ikke forskellen mellem de to IRR'er" in noter
    assert "Fond C er ikke med i porteføljens tal" in noter
    diagram = next(sh.chart for sh in s.shapes if sh.has_chart)
    assert [x.name for x in diagram.plots[0].series] == ["NAV", "Indekskonto (ACWI)",
                                                          "Netto indbetalt"]


def test_pme_springes_over_uden_indeks():
    r = _pme_rapport()
    for f in r["fonde"]:
        f["n"]["ks_pme"] = None
    r["pme"] = {}
    assert slides.pme_fonde(_prs(), r) is None and slides.pme_portefoelje(_prs(), r) is None


def test_pme_kolonner_fylder_indholdsbredden():
    assert sum(b for _, b, _ in slides.PME_KOL) == pytest.approx(
        sum(b for _, b, _ in slides.LT_KOL))


# ---------- cash flow-prognosen ----------

def _prognoserapport():
    from test_prognose import KURSER, OPSAETNING, _fonde
    fonde = _fonde()
    return dict(navn="Investor", as_of=date(2026, 6, 30), fonde=fonde, kurser=KURSER,
                prognose=data.prognose_paa_fondene(fonde, KURSER, date(2026, 6, 30), OPSAETNING),
                datotjek=dict(ens=True, foraeldede=[], skaeve_beholdninger=[]))


def test_prognose_fliser_soejler_pr_fond_og_scenarier_som_linjer():
    r = _prognoserapport()
    s = slides.prognose(_prs(), r)
    t = _tekster(s)
    for label in ("Kald næste 12 mdr.", "Kald næste 3 år", "Udlodninger næste 3 år",
                  "Netto næste 3 år"):
        assert label in t
    assert t[t.index("Kald næste 12 mdr.") + 1] == slides.beloeb(
        r["prognose"]["naeste"]["basis"]["kald_12m"])
    (stroemme,) = [sh.chart for sh in s.shapes if sh.has_chart]
    plot_area = stroemme._chartSpace.chart.plotArea
    # netto pr. scenarie som linjer; det andet scenarie er stiplet
    linjer = plot_area.findall(qn("c:lineChart"))
    assert len(linjer) == 2
    assert sum(len(x.findall(".//" + qn("a:prstDash"))) for x in linjer) == 1
    assert list(stroemme.plots[0].categories)[:2] == ["2026¹", "2027"]
    noter = next(x for x in t if x.startswith("Kilde"))
    assert "Svagt marked: vækst −6 procentpoint, bow × 1,30" in noter
    assert "¹ 2026 er fra 01.07.2026." in noter
    _indenfor_fladen(s, "Forventede kapitalkald og udlodninger")
    assert slides.prognose(_prs(), r | dict(prognose=None)) is None


def test_prognose_nav_flader_scenarier_og_fliser():
    r = _prognoserapport()
    s = slides.prognose_nav(_prs(), r)
    t = _tekster(s)
    p = r["prognose"]
    # tallene ultimo 2029, tre år efter det første (halve) år
    assert t[t.index("NAV 2029, basis") + 1] == slides.beloeb(p["aar"]["basis"][3]["nav"])
    assert t[t.index("NAV 2029, svagt marked") + 1] == slides.beloeb(
        p["aar"]["svag"][3]["nav"])
    assert "Resttilsagn 2029" in t and "Kaldes ikke²" in t
    (nav,) = [sh.chart for sh in s.shapes if sh.has_chart]
    plot_area = nav._chartSpace.chart.plotArea
    # NAV og resttilsagn som flader, det svage markeds NAV som stiplet linje
    assert len(plot_area.find(qn("c:areaChart")).findall(qn("c:ser"))) == 2
    assert plot_area.find(qn("c:lineChart")).find(".//" + qn("a:prstDash")) is not None
    _indenfor_fladen(s, "Forventet NAV og resttilsagn")
    assert slides.prognose_nav(_prs(), r | dict(prognose=None)) is None


def test_prognose_pr_fond_tabel_og_profiler():
    r = _prognoserapport()
    r["fonde"][0]["prognose_afvigelser"] = {"bow": 1.0}
    r["prognose"] = data.prognose_paa_fondene(r["fonde"], r["kurser"], r["as_of"],
                                              __import__("test_prognose").OPSAETNING)
    s = slides.prognose_pr_fond(_prs(), r)
    fonde, profiler = [sh.table for sh in s.shapes if sh.has_table]
    hoved = [c.text for c in fonde.rows[0].cells]
    assert hoved[:4] == ["Fond", "Profil", "Resttilsagn", "Kald 2026¹"]
    assert hoved[-2:] == ["Ikke kaldt²", "TVPI³"]
    assert fonde.cell(1, 1).text == "Private equity⁴"       # egne antagelser
    assert fonde.cell(len(fonde.rows) - 1, 0).text == "I alt"
    assert [profiler.cell(i, 0).text for i in range(1, 3)] == ["Private equity", "Ejendomme"]
    assert profiler.cell(1, 1).text == "25 · 33 · 40 %"
    assert profiler.cell(2, 2).text == "ingen"              # evergreen kalder uden grænse
    noter = next(x for x in _tekster(s) if x.startswith("Kilde"))
    assert "⁴ Alfa har egne antagelser: bow 1,0." in noter
    _indenfor_fladen(s, "Prognose pr. fond")
