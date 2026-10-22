"""Samlingen af decks (report.py): de supplerende slides."""
from datetime import date

from pe_fund_report import report, slides

SUPPLERENDE = ["likviditet", "likviditet", "lange_tilsagn", "fordeling", "brancher", "regioner",
               "branchekort_ejendomme", "regionskort_ejendomme", "vaerdibevaegelser",
               "definitioner"]


def test_supplerende_slides_rummer_dem_uden_for_rapporten(monkeypatch):
    kaldt = []
    for navn in SUPPLERENDE:
        monkeypatch.setattr(slides, navn,
                            lambda prs, r, *a, navn=navn, **k: kaldt.append((navn, a, k)))
    prs = report._byg_supplerende(dict(navn="Motortramp", as_of=date(2026, 6, 30)))
    assert [navn for navn, *_ in kaldt] == SUPPLERENDE
    assert ("fordeling", ("geografi", "Geografisk fordeling efter tilsagn"), {}) in kaldt
    # likviditeten først uden ejendomsfondene, så med
    assert kaldt[:2] == [("likviditet", (), {}), ("likviditet", (), dict(med_ejendomme=True))]
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
    for navn in ("sammendrag", "udvikling", "tilsagn_paa_tvaers", "aendringer_pr_fond",
                 "valutafordeling", "branchekort", "branchekort_aktivklasser", "regionskort",
                 "regionskort_aktivklasser", "stoerste_positioner", "gennemlysning",
                 "portefoeljeoversigt", "definitioner", "gics_oversigt", "sidefod"):
        monkeypatch.setattr(slides, navn, lambda *a: None)
    report._byg_deck(dict(navn="Motortramp", as_of=date(2026, 6, 30), fonde=[]))
    assert kaldt == ["segment"]
