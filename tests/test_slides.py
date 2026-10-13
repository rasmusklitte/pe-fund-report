"""Formater og tabeller i slides.py, bygget på den rigtige slidemaster."""
from datetime import date

from pptx import Presentation

from pe_fund_report import slides
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
