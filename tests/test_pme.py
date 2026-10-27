"""PME (pme.py) og koblingen til rapporten (data.pme_paa_fondene / pme_portefoelje)."""
from datetime import date

import pandas as pd
import pytest

from pe_fund_report import data, pme

T = date(2026, 6, 30)


def _serie(punkter):
    return pd.Series([v for _, v in punkter], index=[d for d, _ in punkter])


CF = [(date(2020, 1, 1), -100.0), (date(2022, 1, 1), -50.0), (date(2024, 1, 1), 60.0)]
NAV = [(T, 150.0)]


def test_fladt_indeks_giver_tvpi_og_irr():
    flad = lambda d: 100.0
    res = pme.pme(CF, NAV, T, flad)
    assert res["ks_pme"] == pytest.approx((60 + 150) / 150)
    assert res["direct_alpha"] == pytest.approx(data.xirr(CF + NAV))
    # indekskontoen står til nettoindbetalingen og har IRR 0
    assert res["indeks_nav"] == pytest.approx(90.0)
    assert res["merverdi"] == pytest.approx(60.0)
    assert res["indeks_irr"] == pytest.approx(0.0, abs=1e-6)


def test_fond_der_foelger_indekset_har_pme_1_og_alfa_0():
    # indekset vokser 10 % om året; fonden køber for 100 og er 100 * 1,1^t værd
    start = date(2020, 1, 1)
    indeks = lambda d: 1.1 ** ((d - start).days / 365)
    nav = 100 * indeks(T) - 40 * indeks(T) / indeks(date(2023, 1, 1))
    res = pme.pme([(start, -100.0), (date(2023, 1, 1), 40.0)], [(T, nav)], T, indeks)
    assert res["ks_pme"] == pytest.approx(1.0)
    assert res["direct_alpha"] == pytest.approx(0.0, abs=1e-6)
    assert res["merverdi"] == pytest.approx(0.0, abs=1e-6)
    assert res["indeks_nav"] == pytest.approx(nav)
    assert res["indeks_irr"] == pytest.approx(0.10, abs=1e-6)


def test_haandregnet_eksempel():
    # kald 100 ved niveau 100, udlodning 50 ved 125, NAV 80 ved 150
    indeks = pme.opslag(_serie([(date(2020, 1, 1), 100.0), (date(2022, 1, 1), 125.0),
                                (T, 150.0)]))
    res = pme.pme([(date(2020, 1, 1), -100.0), (date(2022, 1, 1), 50.0)], [(T, 80.0)], T,
                  indeks)
    assert res["ks_pme"] == pytest.approx((50 * 150 / 125 + 80) / (100 * 150 / 100))
    assert [x["faktor"] for x in res["linjer"]] == pytest.approx([1.5, 1.2, 1.0])
    assert [x["type"] for x in res["linjer"]] == ["Cash flow", "Cash flow", "NAV"]
    # kontoen: 100 * 1,5 - 50 * 1,2 = 90; NAV 80 er 10 under
    assert res["indeks_nav"] == pytest.approx(90.0)
    assert res["merverdi"] == pytest.approx(-10.0)


def test_indekskontoen_uden_vaerdi_har_ingen_irr():
    # stor udlodning, mens indekset faldt: kontoen er tømt
    indeks = pme.opslag(_serie([(date(2020, 1, 1), 100.0), (date(2022, 1, 1), 50.0),
                                (T, 50.0)]))
    res = pme.pme([(date(2020, 1, 1), -100.0), (date(2022, 1, 1), 80.0)], [(T, 10.0)], T,
                  indeks)
    assert res["indeks_nav"] < 0 and res["indeks_irr"] is None
    assert res["ks_pme"] == pytest.approx((80 + 10) / 50)


def test_indekskonto_kvartal_for_kvartal():
    indeks = pme.opslag(_serie([(date(2020, 1, 1), 100.0), (date(2020, 6, 30), 110.0),
                                (date(2020, 12, 31), 121.0)]))
    cf = [(date(2020, 1, 1), -100.0), (date(2020, 6, 30), 20.0)]
    konto = pme.indekskonto(cf, [date(2020, 3, 31), date(2020, 6, 30), date(2020, 12, 31)],
                            indeks)
    assert konto == pytest.approx([100.0, 90.0, 99.0])


def test_cash_flow_foer_indeksets_start_giver_none():
    indeks = pme.opslag(_serie([(date(2008, 3, 28), 35.0), (T, 150.0)]))
    assert pme.pme([(date(2005, 3, 31), -100.0)], [(T, 200.0)], T, indeks) is None


def test_opslag_tager_seneste_niveau_og_afviser_forældet_indeks():
    niveau = pme.opslag(_serie([(date(2026, 6, 26), 10.0), (date(2026, 6, 29), 11.0)]))
    assert niveau(date(2026, 6, 28)) == 10.0          # weekend: fredagens kurs
    assert niveau(date(2026, 6, 30)) == 11.0
    assert niveau(date(2026, 6, 25)) is None
    assert niveau(date(2026, 9, 30)) is None          # indekset slutter et kvartal før


def test_omregnet_til_anden_valuta():
    usd = lambda d: 100.0
    dkk_pr_usd, dkk_pr_eur = (lambda d: 6.5), (lambda d: 7.46)
    assert pme.omregnet(usd, dkk_pr_usd, None)(T) == pytest.approx(650.0)
    assert pme.omregnet(usd, dkk_pr_usd, dkk_pr_eur)(T) == pytest.approx(650 / 7.46)
    assert pme.omregnet(usd, lambda d: None, None)(T) is None


def _fond(kode, valuta, flows):
    return dict(fund_code=kode, fund_name=f"Fond {kode}", valuta=valuta, nav_dato=T,
                irr_flows=flows, n={})


def test_pme_paa_fondene_omregner_indekset_til_fondens_valuta():
    serie = _serie([(date(2010, 1, 1), 100.0), (T, 200.0)])
    # USD steg fra 5 til 6,5 DKK; for en DKK-fond voksede indekset altså 2 * 1,3 = 2,6x
    usd = _serie([(date(2010, 1, 1), 5.0), (T, 6.5)])
    dkk = _fond("A", "DKK", [(date(2010, 1, 1), -100.0), (T, 260.0)])
    usd_fond = _fond("B", "USD", [(date(2010, 1, 1), -100.0), (T, 260.0)])
    gammel = _fond("C", "DKK", [(date(2005, 1, 1), -100.0), (T, 260.0)])
    indeks_dkk = data.pme_paa_fondene([dkk, usd_fond, gammel], serie,
                                      {"USD": usd, "DKK": None}, "USD")
    assert dkk["n"]["ks_pme"] == pytest.approx(1.0)
    assert usd_fond["n"]["ks_pme"] == pytest.approx(1.3)
    assert gammel["n"]["ks_pme"] is None and "før indeksets start" in gammel["pme"]["grund"]
    assert indeks_dkk(T) == pytest.approx(1300.0)


def test_pme_paa_fondene_uden_indeks():
    f = _fond("A", "DKK", [(date(2010, 1, 1), -100.0), (T, 260.0)])
    data.pme_paa_fondene([f], None, {}, "USD")
    assert f["n"]["ks_pme"] is None and f["n"]["direct_alpha"] is None


def test_pme_portefoelje_udelader_fonde_foer_indeksets_start():
    indeks = pme.opslag(_serie([(date(2010, 1, 1), 100.0), (T, 100.0)]))
    a, c = _fond("A", "DKK", []), _fond("C", "DKK", [])
    flows = [dict(fund_code="C", fund_name="Fond C", dato=date(2005, 1, 1), dkk=-100.0,
                  type="Cash flow"),
             dict(fund_code="A", fund_name="Fond A", dato=date(2010, 1, 1), dkk=-100.0,
                  type="Cash flow"),
             dict(fund_code="C", fund_name="Fond C", dato=T, dkk=300.0,
                  type="NAV (sidste cash flow)"),
             dict(fund_code="A", fund_name="Fond A", dato=T, dkk=150.0,
                  type="NAV (sidste cash flow)")]
    p = data.pme_portefoelje([a, c], dict(irr=0.2, flows=flows), indeks, T)
    assert p["udeladte"] == ["Fond C"]
    assert p["ks_pme"] == pytest.approx(1.5)           # fladt indeks: TVPI for fond A
    assert p["irr"] == pytest.approx(data.xirr([(date(2010, 1, 1), -100.0), (T, 150.0)]))
    assert p["fonde"] == ["Fond A"] and p["merverdi"] == pytest.approx(50.0)
    assert {x["fund_name"] for x in p["linjer"]} == {"Fond A"}
    assert data.pme_portefoelje([a], None, indeks, T) is None


def test_pme_forloeb_pr_kvartal_til_rapportdatoen():
    indeks = pme.opslag(_serie([(date(2025, 1, 1), 100.0), (date(2026, 3, 31), 120.0),
                                (T, 150.0)]))
    linjer = [dict(dato=date(2025, 11, 15), beloeb=-100.0, type="Cash flow"),
              dict(dato=T, beloeb=130.0, type="NAV")]
    navs = {date(2025, 12, 31): 100.0, date(2026, 3, 31): 110.0, T: 130.0}
    fl = data.pme_forloeb(linjer, navs.get, indeks, T)
    assert [x["dato"] for x in fl] == list(navs)
    assert [x["nav"] for x in fl] == [100.0, 110.0, 130.0]
    assert [x["indekskonto"] for x in fl] == pytest.approx([100.0, 120.0, 150.0])
    assert all(x["netto_indbetalt"] == 100.0 for x in fl)
    assert data.pme_forloeb(linjer, lambda d: None, indeks, T) is None
