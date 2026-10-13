"""Beregninger i data.py, som ikke kræver databasen."""
from datetime import date

import pytest

from pe_fund_report import data


def test_xirr_et_aar_med_ti_procent():
    flows = [(date(2024, 1, 1), -100.0), (date(2025, 1, 1), 110.0)]
    assert data.xirr(flows) == pytest.approx(0.10, abs=1e-3)   # 366 dage, act/365


def test_xirr_uden_fortegnsskifte_er_none():
    assert data.xirr([(date(2024, 1, 1), -100.0), (date(2025, 1, 1), -5.0)]) is None
    assert data.xirr([]) is None


def _fond(valuta, tilsagn, indbetalt, udloddet, nav, flows):
    n = dict(tilsagn=tilsagn, indbetalt=indbetalt, udloddet=udloddet, nav=nav,
             resttilsagn=tilsagn - indbetalt)
    return dict(valuta=valuta, n=n, irr_flows=flows)


def test_total_laegger_fonde_i_samme_valuta_sammen():
    a = _fond("DKK", 100, 50, 10, 60, [(date(2024, 1, 1), -50), (date(2025, 1, 1), 70)])
    b = _fond("DKK", 200, 150, 0, 150, [(date(2024, 1, 1), -150), (date(2025, 1, 1), 150)])
    t = data.total([a, b])
    assert (t["tilsagn"], t["indbetalt"], t["udloddet"], t["nav"]) == (300, 200, 10, 210)
    assert t["picc"] == pytest.approx(200 / 300)
    assert t["dpi"] == pytest.approx(0.05)
    assert t["tvpi"] == pytest.approx(1.10)
    assert t["irr"] == pytest.approx(0.10, abs=1e-3)


def test_total_er_none_ved_blandede_valutaer():
    a = _fond("DKK", 100, 50, 0, 50, [])
    b = _fond("EUR", 100, 50, 0, 50, [])
    assert data.total([a, b]) is None


def test_fund_info_har_de_forventede_felter():
    for kode, info in data.fund_info().items():
        assert {"status", "beskrivelse", "geografi", "segment"} <= set(info), kode
