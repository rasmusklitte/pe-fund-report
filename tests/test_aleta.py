"""Rapporten mod Aletas tal for Motortramp (aleta.py). Læser databasen og springes over uden den.

    python -m pytest tests/test_aleta.py

Hver fond × kontrol er sin egen test, fx test_aleta[SEED4-NAV], og hver fonds kapitalstrømme mod
Aletas transaktioner er test_transaktioner[SEED4]. Kørslen skriver de samme linjer som
afstemningsdokument til exports/aleta_afstemning_<dato>.xlsx.
"""
from datetime import date

import pandas as pd
import pyodbc
import pytest
from aleta import (
    ALETA_DATO,
    AS_OF,
    FONDE,
    INVESTOR,
    KONTROLLER,
    TRANSAKTIONER,
    afstem,
    afstem_transaktioner,
    aleta_dato,
    laes_aleta,
    laes_transaktioner,
    skriv,
)
from sqlalchemy.exc import DBAPIError

from pe_fund_report.aleta import TRANSAKTIONSFONDE


@pytest.fixture(scope="module")
def rapport():
    from pe_fund_report import data
    try:
        return data.investor_rapport(INVESTOR, AS_OF.isoformat())
    except (DBAPIError, pyodbc.Error) as e:
        pytest.skip(f"databasen kan ikke nås: {e.__class__.__name__}")


@pytest.fixture(scope="module")
def transaktioner(rapport):
    if TRANSAKTIONER is None:
        return []
    return afstem_transaktioner(laes_transaktioner(), rapport)


@pytest.fixture(scope="module")
def linjer(rapport, transaktioner):
    from pe_fund_report import data, report
    from pe_fund_report.paths import EXPORTS
    # Aleta omregner beholdningerne med kursen på sin egen dato, den første i næste måned
    valutaer = {f["valuta"] for f in rapport["fonde"]}
    kurser_aleta = {v: data._dkk_kurs(v, ALETA_DATO) for v in valutaer}
    alle = afstem(laes_aleta(), rapport, kurser_aleta, transaktioner)
    sti = report._gem(skriv(alle, transaktioner=transaktioner), EXPORTS / f"aleta_afstemning_{AS_OF:%Y-%m-%d}.xlsx")
    print(f"\nafstemningsdokument: {sti}")
    return {(x["fond"], x["kontrol"]): x for x in alle}


@pytest.mark.db
def test_samme_fonde(rapport):
    aleta = {FONDE.get(a["securityCode"], a["securityCode"]) for a in laes_aleta()}
    assert aleta == {f["fund_code"] for f in rapport["fonde"]}


@pytest.mark.db
@pytest.mark.parametrize("fond", list(TRANSAKTIONSFONDE.values()))
def test_transaktioner(transaktioner, linjer, fond):
    if TRANSAKTIONER is None:
        pytest.skip("Aletas transaktionsudtræk (private_markets_transactions_*.xlsx) mangler")
    # en anden dato er en lille afvigelse; et andet beløb eller en strøm kun det ene sted afviger
    forkert = [t for t in transaktioner if t["fond"] == fond and t["status"] == "afviger"]
    assert not forkert, "; ".join(
        f"{t['art']} {t['type']}: Aleta {t['aleta']} pr. {t['aleta_dato']}, rapport {t['rapport']} "
        f"pr. {t['rapport_dato']} {t['ref']}".rstrip() for t in forkert)


@pytest.mark.db
@pytest.mark.parametrize("kontrol", [k[0] for k in KONTROLLER])
@pytest.mark.parametrize("fond", list(FONDE.values()))
def test_aleta(linjer, fond, kontrol):
    x = linjer[fond, kontrol]
    assert x["status"] != "afviger", (
        f"{fond} {kontrol}: Aleta {x['aleta']}, rapport {x['rapport']}, afvigelse "
        f"{x['afvigelse']}, tilladt {x['tilladt']} {x['bemaerkning']}".rstrip())


def _aleta(**k):
    return dict(dict(asset="Testfond", securityCode="KAPE2020", navCurrency="DKK", vintage=2020,
                     nav=1000.0, marketValue=1000.0, commitment=2000.0, unfundedCommitment=0.0,
                     drawdowns=800.0, distributions=100.0, calledPercent=0.4, dpi=0.125,
                     rvpi=1.25, tvpi=1.375, irr=0.1), **k)


def _rapport(valuta="DKK", **n):
    return dict(fonde=[dict(fund_code="KAPE20", valuta=valuta, vintage=2020, n=dict(dict(
        nav=1000.0, nav_dkk=1000.0, tilsagn=2000.0, tilsagn_dkk=2000.0, resttilsagn=0.0,
        resttilsagn_dkk=0.0, indbetalt=800.0,
        indbetalt_dkk=800.0, udloddet=100.0, udloddet_dkk=100.0, picc=0.4, dpi=0.125,
        tvpi=1.375, irr=0.1), **n))])


def test_aleta_dato():
    assert aleta_dato(date(2026, 6, 30)) == date(2026, 7, 1)
    assert aleta_dato(date(2025, 12, 31)) == date(2026, 1, 1)


def test_afstem_status():
    status = {x["kontrol"]: x["status"] for x in afstem([_aleta()], _rapport())}
    assert set(status.values()) == {"ok"}
    # 0,4 % ved siden af er en lille afvigelse, 0,6 % afviger; IRR har en absolut tolerance på
    # 0,25 pct.-point; en afvigelse, der forsvinder i afrundingen, er ok
    status = {x["kontrol"]: x["status"] for x in afstem(
        [_aleta()], _rapport(nav=1004.0, nav_dkk=1006.0, irr=0.103, udloddet_dkk=100.3))}
    assert (status["NAV"], status["NAV DKK"], status["IRR"]) == (
        "lille afvigelse", "afviger", "afviger")
    assert status["Udloddet DKK"] == "ok"
    assert status["RVPI"] == "lille afvigelse"      # regnes af NAV og indbetalt: 1004 / 800
    # under 1 % af den tilladte afvigelse er ok: NAV 1000,04 mod 1000 (tilladt 5, grænse 0,05)
    status = {x["kontrol"]: x["status"] for x in afstem([_aleta()], _rapport(nav=1000.04))}
    assert status["NAV"] == "ok"


def test_afstem_aletas_kurs():
    # beholdningerne omregnes med kursen på Aletas dato, kapitalstrømmene med transaktionskurserne
    linjer = {x["kontrol"]: x for x in afstem(
        [_aleta(navCurrency="EUR", marketValue=7470.0, commitment=14940.0, drawdowns=5960.0)],
        _rapport("EUR", nav_dkk=7460.0, tilsagn_dkk=14920.0, indbetalt=800.0, indbetalt_dkk=5960.0),
        {"EUR": (7.47, date(2026, 7, 1))})}
    assert linjer["NAV DKK"]["rapport"] == pytest.approx(7470.0)
    assert linjer["Tilsagn DKK"]["rapport"] == pytest.approx(14940.0)
    assert (linjer["NAV DKK"]["kurs"], linjer["NAV DKK"]["kursdato"]) == (7.47, date(2026, 7, 1))
    assert linjer["Indbetalt DKK"]["rapport"] == 5960.0
    assert linjer["Indbetalt DKK"]["kurs"] == pytest.approx(7.45)
    assert linjer["Indbetalt DKK"]["kursdato"] == "transaktionsdatoer"
    assert {linjer[k]["status"] for k in ("NAV DKK", "Tilsagn DKK", "Indbetalt DKK")} == {"ok"}
    # uden en kurs på Aletas dato kan beholdningen ikke sammenlignes
    x = {x["kontrol"]: x for x in afstem([_aleta()], _rapport(), {})}["NAV DKK"]
    assert (x["rapport"], x["status"]) == (None, "afviger")
    assert "ingen DKK-kurs" in x["bemaerkning"]


def test_afstem_tolerancer_og_valuta():
    # resttilsagnet måles mod tilsagnet (2.000 * 0,5 % = 10), ikke mod Aletas 0
    linjer = {x["kontrol"]: x for x in afstem([_aleta()], _rapport(resttilsagn_dkk=11.0))}
    assert linjer["Resttilsagn DKK"]["tilladt"] == 10.0
    assert linjer["Resttilsagn DKK"]["status"] == "afviger"
    # kapitalstrømmene sammenlignes i DKK, også for en fond i EUR
    linjer = {x["kontrol"]: x for x in afstem(
        [_aleta(navCurrency="EUR", drawdowns=746.0)],
        _rapport("EUR", indbetalt=100.0, indbetalt_dkk=746.0, udloddet=13.4))}
    assert linjer["Indbetalt DKK"]["status"] == linjer["Udloddet DKK"]["status"] == "ok"
    assert linjer["Valuta"]["status"] == "ok"
    # med kursen står beløbene også i fondens valuta; Aletas DKK regnes tilbage med kursen
    linjer = {x["kontrol"]: x for x in afstem(
        [_aleta(navCurrency="EUR", drawdowns=746.0)],
        _rapport("EUR", indbetalt=100.0, indbetalt_dkk=746.0) | dict(kurser={"EUR": (7.46, None)}))}
    x = linjer["Indbetalt DKK"]
    assert (x["valuta"], x["aleta_fv"], x["rapport_fv"]) == ("EUR", pytest.approx(100.0), 100.0)
    assert linjer["NAV"]["aleta_fv"] == 1000.0          # Aletas nav er allerede i fondens valuta
    assert linjer["IRR"]["aleta_fv"] is None
    # en fond, rapporten ikke har
    linjer = afstem([_aleta(securityCode="UKENDT")], _rapport())
    assert {x["status"] for x in linjer} == {"afviger"}
    assert linjer[0]["bemaerkning"] == "fonden er ikke i rapporten"



def test_afstem_aarsager():
    # Aletas NAV er forvalterens NAV før carry (CA890); RVPI følger med
    konto = pd.DataFrame({"period_end": [date(2026, 6, 30)] * 3, "basis": ["ITD"] * 3,
                          "account_code": ["CA890", "CA900", "CA110"],
                          "amount": [1200.0, 1000.0, 800.0]})
    rapport = _rapport(irr=0.08)
    rapport["fonde"][0] |= {"nav_dato": date(2026, 6, 30), "kilder": {"kapitalkonto": konto},
                            "irr_flows": [(date(2021, 6, 30), -800.0), (date(2026, 6, 30), 1000.0)]}
    linjer = {x["kontrol"]: x for x in afstem(
        [_aleta(nav=1200.0, marketValue=1200.0, rvpi=1.5, tvpi=1.625)], rapport)}
    assert "NAV før carried interest" in linjer["NAV"]["bemaerkning"]
    assert linjer["RVPI"]["bemaerkning"] == "Følger forskellen i NAV"
    assert linjer["IRR"]["bemaerkning"].startswith("Med NAV før carry giver rapportens strømme")
    # Aleta regner udligningsrenten ud af indbetalt; PICC følger med
    linjer = {x["kontrol"]: x for x in afstem([_aleta()], _rapport(
        indbetalt=820.0, indbetalt_dkk=820.0, udligning=20.0, picc=0.41))}
    assert linjer["Indbetalt DKK"]["bemaerkning"].startswith("Aleta regner udligningsrenten (20")
    assert linjer["PICC"]["bemaerkning"] == "Følger forskellen i indbetalt"
    # en afvigelse uden en årsag, der kan regnes efter, får ingen bemærkning
    linjer = {x["kontrol"]: x for x in afstem([_aleta()], _rapport(udloddet_dkk=110.0))}
    assert linjer["Udloddet DKK"]["bemaerkning"] == ""


def _tx(dato, beloeb, type_="kald", **k):
    return {"fond": "KAPE20", "id": f"id{dato}", "type": type_, "dato": dato, "beloeb": beloeb,
            "beloeb_dkk": beloeb, "kurs": 1.0, "i_noegletal": True, "vintage": date(2020, 1, 1),
            "kommentar": ""} | k


def test_afstem_transaktioner():
    rapport = _rapport()
    rapport["fonde"][0]["kapitalstroemme"] = [
        (date(2021, 1, 10), -100.0, "kald"), (date(2021, 3, 1), -150.0, "kald"),
        (date(2021, 3, 1), -50.0, "kald"),                    # to konti samme dag
        (date(2022, 5, 1), -300.0, "kald"), (date(2023, 1, 1), -80.0, "kald"),
        (date(2021, 6, 1), 40.0, "udlodning")]
    tx = [_tx(date(2021, 1, 10), 100.0), _tx(date(2021, 2, 26), 200.0),
          _tx(date(2022, 5, 1), 330.0), _tx(date(2024, 2, 1), 70.0),
          _tx(date(2021, 6, 1), 40.0, "udlodning"), _tx(date(2026, 8, 1), 90.0)]
    linjer = {(t["type"], t["art"]): t for t in afstem_transaktioner(tx, rapport)}
    assert set(linjer) == {("kald", "ens"), ("kald", "anden dato"), ("kald", "andet beløb"),
                           ("kald", "kun i Aleta"), ("kald", "kun i rapporten"),
                           ("kald", "efter rapportdatoen"), ("udlodning", "ens")}
    x = linjer["kald", "anden dato"]
    assert (x["rapport"], x["dage"], x["status"]) == (200.0, 3, "lille afvigelse")
    assert linjer["kald", "andet beløb"]["afvigelse"] == -30.0
    assert linjer["kald", "efter rapportdatoen"]["status"] == "efter rapportdatoen"
    # forklar() nævner beløbsforskellene ved indbetalt
    linjer = {x["kontrol"]: x for x in afstem(
        [_aleta(drawdowns=700.0)], _rapport(indbetalt_dkk=630.0), None,
        afstem_transaktioner(tx, rapport))}
    assert "pr. 01.05.2022 har Aleta 330, rapporten 300" in linjer["Indbetalt DKK"]["bemaerkning"]


def test_irr_sammenligning():
    from pe_fund_report.aleta import irr_sammenligning
    from pe_fund_report.data import xirr
    rapport = _rapport()
    f = rapport["fonde"][0]
    f["kapitalstroemme"] = [(date(2021, 3, 1), -100.0, "kald"), (date(2022, 5, 1), -300.0, "kald")]
    f["nav_dato"], f["n"]["nav"] = AS_OF, 600.0
    f["irr_flows"] = [(date(2021, 3, 1), -100.0), (date(2022, 5, 1), -300.0), (AS_OF, 600.0)]
    f["n"]["irr"] = xirr(f["irr_flows"])
    tx = [_tx(date(2021, 2, 26), 100.0), _tx(date(2022, 5, 1), 330.0)]
    aleta_irr = xirr([(date(2021, 2, 26), -100.0), (date(2022, 5, 1), -330.0), (ALETA_DATO, 600.0)])
    a = _aleta(irr=aleta_irr, marketValue=600.0, nav=600.0)
    x = irr_sammenligning(f, a, afstem_transaktioner(tx, rapport), ALETA_DATO, 1.0)
    assert x["metode_ok"] and x["aleta"] == aleta_irr and x["irr_dkk"] == f["n"]["irr"]
    # forskellen er mellem rapportens IRR i DKK og Aletas; i en DKK-fond er de to rapport-IRR ens
    assert x["forskel"] == pytest.approx(f["n"]["irr"] - aleta_irr)
    assert x["irr_aletas_datoer"] == pytest.approx(xirr(
        [(date(2021, 2, 26), -100.0), (date(2022, 5, 1), -300.0), (ALETA_DATO, 600.0)]))
    assert x["aarsager"] == ["1 betaling har en tidligere dato hos Aleta",
                             "kaldet pr. 01.05.2022 har et andet beløb hos Aleta"]


def test_irr_sammenligning_i_fremmed_valuta_sammenligner_i_dkk():
    from pe_fund_report.aleta import irr_sammenligning
    from pe_fund_report.data import xirr
    rapport = _rapport()
    f = rapport["fonde"][0]
    f["valuta"] = "USD"
    datoer = [date(2021, 3, 1), date(2022, 5, 1)]
    f["kapitalstroemme"] = [(datoer[0], -100.0, "kald"), (datoer[1], -300.0, "kald")]
    f["kilder"] = {"cash_flows": pd.DataFrame({"value_date": datoer, "fx_rate": [7.0, 7.0],
                                               "flow_ref": ["", ""], "notice_date": [None, None]})}
    f["nav_dato"], f["n"]["nav"] = AS_OF, 600.0
    f["irr_flows"] = [(datoer[0], -100.0), (datoer[1], -300.0), (AS_OF, 600.0)]
    f["n"]["irr"] = xirr(f["irr_flows"])
    # USD er faldet fra 7,0 til 6,0 kr., så IRR i DKK er lavere end i USD
    tx = [_tx(datoer[0], 100.0, beloeb_dkk=700.0, kurs=7.0),
          _tx(datoer[1], 300.0, beloeb_dkk=2100.0, kurs=7.0)]
    a = _aleta(irr=0.05, marketValue=3600.0, nav=600.0)
    x = irr_sammenligning(f, a, afstem_transaktioner(tx, rapport), ALETA_DATO, 6.0)
    irr_dkk = xirr([(datoer[0], -700.0), (datoer[1], -2100.0), (ALETA_DATO, 3600.0)])
    assert x["irr_dkk"] == pytest.approx(irr_dkk) and irr_dkk < f["n"]["irr"]
    # Aleta regner i DKK, så forskellen er mellem de to IRR i DKK, uden valutaeffekten
    assert x["forskel"] == pytest.approx(irr_dkk - 0.05)
    assert not any("USD" in a for a in x["aarsager"])


def test_find_kraever_udtraek_paa_aletas_dato(tmp_path):
    from pe_fund_report.aleta import find
    mappe = tmp_path / "INV"
    mappe.mkdir()
    assert find("INV", AS_OF, tmp_path) is None
    (mappe / "noegletal_2026-07-01.xlsx").touch()
    assert find("INV", AS_OF, tmp_path) == (mappe / "noegletal_2026-07-01.xlsx", None)
    for navn in ("pm_transactions_INV_2026-09-01.xlsx", "pm_transactions_INV_2026-10-22.xlsx"):
        (mappe / navn).touch()
    assert find("INV", AS_OF, tmp_path)[1].name == "pm_transactions_INV_2026-10-22.xlsx"
    # det nye navn; datoen sidst i navnet afgør, ikke præfikset
    (mappe / "private_markets_transactions_INV_2005-03-31_2026-10-26_2026-10-26.xlsx").touch()
    assert find("INV", AS_OF, tmp_path)[1].name.endswith("_2026-10-26.xlsx")
    (mappe / "pm_transactions_INV_2026-10-30.xlsx").touch()
    assert find("INV", AS_OF, tmp_path)[1].name == "pm_transactions_INV_2026-10-30.xlsx"
    assert find("INV", date(2026, 3, 31), tmp_path) is None
