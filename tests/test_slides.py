"""Formater og tabeller i slides.py, bygget på den rigtige slidemaster."""
import io
from datetime import date
from itertools import pairwise

import pytest
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Emu, Inches

from pe_fund_report import data, slides
from pe_fund_report.deck import L_TEKST, _ny_slide, _ryd_slides
from pe_fund_report.paths import MASTER


def test_danske_talformater():
    assert slides.tusind(1_234_567) == "1.235"
    assert slides.pct(0.123) == "12,3 %"
    assert slides.gange(1.5) == "1,50x"
    assert slides.tusind(None) == slides.pct(None) == slides.gange(None) == "n/a"
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


def _hoved(slide, navn):
    """Overskriftens tekstboks som (venstre, højre)."""
    for sh in slide.shapes:
        if (_i_hovedbaandet(sh) and sh.has_text_frame
                and sh.text_frame.text.replace("\v", "\n").strip() == navn):
            return Emu(sh.left).inches, Emu(sh.left).inches + Emu(sh.width).inches
    return None


def test_lange_tilsagn_overskrifter_flugter_med_kolonnerne():
    s = slides.lange_tilsagn(_prs(), _rapport([_lt_fond("A", 100.0)]))
    for i, (navn, bredde, just) in enumerate(slides.LT_KOL):
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
    r = _tpt_rapport([_tpt_fond("Fond A", 100_000.0)])
    t = _tekster(slides.tilsagn_paa_tvaers(_prs(), r))
    assert "Tilsagn på tværs af fonde" in t
    assert "DKK\vResttilsagn" in t and "DKK\vNAV" in t and "PICC" in t
    assert "100" in t and "50,0 %" in t and "50" in t and "75" in t
    assert not any("Kaldt" in x for x in t)          # vandfaldet er væk


def test_tilsagn_paa_tvaers_summerer_dkk_paa_tvaers_af_valutaer():
    dkk = _tpt_fond("DKK-fond", 100_000.0)
    eur = _tpt_fond("EUR-fond", 10_000.0, valuta="EUR", kurs=7.46)
    r = _tpt_rapport([dkk, eur], total=None,
                     total_dkk=dict(resttilsagn_dkk=50_000 + 37_300, nav_dkk=75_000 + 55_950,
                                    indbetalt_dkk=100_000.0, udloddet_dkk=0.0),
                     kurser={"DKK": (1.0, date(2026, 6, 30)), "EUR": (7.46, date(2026, 6, 30))})
    r["valuta"] = None
    t = _tekster(slides.tilsagn_paa_tvaers(_prs(), r))
    assert "I alt" in t
    assert "87" in t and "131" in t                   # DKK-summerne
    assert any("i alt-rækken viser beløbene i DKK" in x for x in t)
    assert slides.gange((75_000 + 55_950) / 100_000) in t     # TVPI i DKK
    assert any("Nationalbanks kurs pr. 30.06.2026" in x and "EUR 7,4600" in x for x in t)


def test_tilsagn_paa_tvaers_uden_fremmed_valuta_har_ingen_kursnote():
    r = _tpt_rapport([_tpt_fond("A", 100.0)])
    t = _tekster(slides.tilsagn_paa_tvaers(_prs(), r))
    assert not any("Nationalbanks kurs" in x for x in t)


def test_tilsagn_paa_tvaers_markerer_fond_med_afvigende_opgoerelsesdato():
    r = _tpt_rapport([_tpt_fond("Fond A", 100.0), _tpt_fond("Seed", 50.0)],
                     foraeldede=[dict(fond="Seed", dato=date(2025, 9, 30), kvartaler=3)])
    t = _tekster(slides.tilsagn_paa_tvaers(_prs(), r))
    assert f"Seed {slides.DAGGER}" in t and "opgjort pr. 30.09.2025" in t


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
    assert rk[0] == ["", "Valuta", "Tilsagn", "DKK\vTilsagn", "DKK\vNAV", "DKK\vResttilsagn",
                     "TVPI", "Andel af\vtilsagn"]           # valutaen på linjen over navnet
    assert [x[0].strip() for x in rk[1:]] == [
        "I alt", "Europa", "Fond EU", "Fond EU2", "I alt Europa",
        "Danmark", "Fond DK", "I alt Danmark"]
    assert rk[1][3] == slides.tusind(600_000) and rk[1][7] == slides.pct(1.0)


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
    s = slides.definitioner(_prs(), _rapport([]))
    t = _tekster(s)
    assert "Definitioner" in t
    for begreb, tekst in slides.DEFINITIONER:
        assert begreb in t and tekst in t, begreb
    venstre, hoejre = slides._def_halvdele(slides.DEFINITIONER)
    for spalte in (venstre, hoejre):
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
    for label in ("fonde i porteføljen", "tilsagn, tDKK", "NAV, tDKK", "TVPI"):
        assert label in t, label
    assert slides.tusind(r["total_dkk"]["tilsagn_dkk"]) in t
    assert any("3 fonde hos 2 forvaltere med vintage 2020–2023" in x for x in t)
    assert any(x.startswith("Største geografi er Europa") for x in t)
    assert "Geografi" in t and "Segment" in t
    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text not in ("", "Executive summary"):
            assert Emu(sh.top).inches <= slides.Y_NOTER + 0.01, sh.text_frame.text
