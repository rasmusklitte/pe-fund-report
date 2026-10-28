"""Samlingen af decks (report.py): de supplerende slides og afstemningen med Aleta."""
from datetime import date

import pytest

from pe_fund_report import report, slides

SUPPLERENDE = ["regnskabstal", "likviditet", "pme_portefoelje", "pme_fonde", "aarlige_stroemme", "tvpi_over_tid", "kaldstempo",
               "eksponering_over_tid", "prognose", "prognose_nav", "prognose_pr_fond", "bevaegelser_i_dybden", "koncentration",
               "ejendomme_i_dybden", "ejendomsbidrag", "lange_tilsagn",
               "fordeling", "brancher", "regioner", "branchekort_ejendomme",
               "regionskort_ejendomme", "definitioner"]
ALETA = ["aleta_afstemning", "irr_mod_aleta"]


def test_supplerende_slides_rummer_dem_uden_for_rapporten(monkeypatch):
    kaldt = []
    for navn in SUPPLERENDE:
        monkeypatch.setattr(slides, navn,
                            lambda prs, r, *a, navn=navn, **k: kaldt.append((navn, a, k)))
    prs = report._byg_supplerende(dict(navn="Motortramp", as_of=date(2026, 6, 30)))
    # regnskabstallene for kvartalet og for året
    assert [navn for navn, *_ in kaldt] == ["regnskabstal"] + SUPPLERENDE
    assert kaldt[:2] == [("regnskabstal", ("kvartal",), {}), ("regnskabstal", ("aar",), {})]
    assert ("fordeling", ("geografi", "Geografisk fordeling efter tilsagn"), {}) in kaldt
    # likviditeten uden ejendomsfondene; med dem står i rapporten
    assert kaldt[2] == ("likviditet", (), {})
    # forside, Appendiks-skilleside og slutside; slidesne selv er erstattet ovenfor
    assert len(prs.slides._sldIdLst) == 3
    forside = [sh.text_frame.text for sh in prs.slides[0].shapes if sh.has_text_frame]
    assert "Motortramp – supplerende slides" in forside
    assert "Slides uden for porteføljerapporten pr. 30.06.2026" in forside


def test_supplerende_slides_er_ikke_i_rapporten(monkeypatch):
    kaldt = []
    monkeypatch.setattr(slides, "fordeling", lambda prs, r, noegle, titel: kaldt.append(noegle))
    for navn in SUPPLERENDE:
        if navn not in ("fordeling", "definitioner"):
            monkeypatch.setattr(slides, navn, lambda *a, navn=navn, **k: kaldt.append(navn))
    monkeypatch.setattr(slides, "likviditet", lambda prs, r, **k: kaldt.append(("likviditet", k)))
    # afstemningen med Aleta står i sit eget deck
    for navn in ALETA:
        monkeypatch.setattr(slides, navn, lambda *a, navn=navn: kaldt.append(navn))
    for navn in ("sammendrag", "udvikling", "tilsagn_paa_tvaers", "vaerdibevaegelser",
                 "aendringer_pr_fond",
                 "valutafordeling", "branchekort", "branchekort_aktivklasser", "regionskort",
                 "regionskort_aktivklasser", "stoerste_positioner", "gennemlysning",
                 "portefoeljeoversigt", "definitioner", "gics_oversigt",
                 "sidefod"):
        monkeypatch.setattr(slides, navn, lambda *a: None)
    report._byg_deck(dict(navn="Motortramp", as_of=date(2026, 6, 30), fonde=[]))
    # samlede investeringer med ejendomsfondene er rykket ind i rapporten
    assert kaldt == [("likviditet", dict(med_ejendomme=True)), "segment"]


def test_aleta_afstemningen_er_et_deck_for_sig(monkeypatch):
    kaldt = []
    for navn in ALETA:
        monkeypatch.setattr(slides, navn, lambda prs, r, navn=navn: kaldt.append(navn) or navn)
    monkeypatch.setattr(slides, "sidefod", lambda *a: kaldt.append("sidefod"))
    r = dict(navn="Motortramp", as_of=date(2026, 6, 30))
    prs = report._byg_aleta(r)
    assert kaldt == ALETA + ["sidefod"]
    # forside og slutside; slidesne selv er erstattet ovenfor
    assert len(prs.slides._sldIdLst) == 2
    forside = [sh.text_frame.text for sh in prs.slides[0].shapes if sh.has_text_frame]
    assert "Motortramp – afstemning med Aleta" in forside
    assert "Rapportens tal mod Aleta pr. 30.06.2026" in forside
    # uden Aletas udtræk springer begge slides over, og der er intet deck
    for navn in ALETA:
        monkeypatch.setattr(slides, navn, lambda prs, r: None)
    assert report._byg_aleta(r) is None


def test_aleta_afstemningen_laves_ikke_anonymiseret():
    with pytest.raises(ValueError, match="Aleta"):
        report.byg("MOTORTRAMP", anonym=True, med_aleta=True)
