"""Rapporten mod Aletas tal for Motortramp (aleta.py). Læser databasen og springes over uden den.

    python -m pytest tests/test_aleta.py

Hver fond × kontrol er sin egen test, fx test_aleta[SEED4-NAV], og kørslen skriver de samme
linjer som afstemningsdokument til exports/aleta_afstemning_<dato>.xlsx.
"""
import pyodbc
import pytest
from sqlalchemy.exc import DBAPIError

from aleta import AS_OF, FONDE, INVESTOR, KONTROLLER, afstem, laes_aleta, skriv


@pytest.fixture(scope="module")
def rapport():
    from pe_fund_report import data
    try:
        return data.investor_rapport(INVESTOR, AS_OF.isoformat())
    except (DBAPIError, pyodbc.Error) as e:
        pytest.skip(f"databasen kan ikke nås: {e.__class__.__name__}")


@pytest.fixture(scope="module")
def linjer(rapport):
    from pe_fund_report import report
    from pe_fund_report.paths import EXPORTS
    alle = afstem(laes_aleta(), rapport)
    sti = report._gem(skriv(alle), EXPORTS / f"aleta_afstemning_{AS_OF:%Y-%m-%d}.xlsx")
    print(f"\nafstemningsdokument: {sti}")
    return {(x["fond"], x["kontrol"]): x for x in alle}


@pytest.mark.db
def test_samme_fonde(rapport):
    aleta = {FONDE.get(a["securityCode"], a["securityCode"]) for a in laes_aleta()}
    assert aleta == {f["fund_code"] for f in rapport["fonde"]}


@pytest.mark.db
@pytest.mark.parametrize("kontrol", [k[0] for k in KONTROLLER])
@pytest.mark.parametrize("fond", list(FONDE.values()))
def test_aleta(linjer, fond, kontrol):
    x = linjer[fond, kontrol]
    assert x["status"] == "ok", (
        f"{fond} {kontrol}: Aleta {x['aleta']}, rapport {x['rapport']}, afvigelse "
        f"{x['afvigelse']}, tilladt {x['tilladt']} {x['bemaerkning']}".rstrip())


def _aleta(**k):
    return dict(dict(asset="Testfond", securityCode="KAPE2020", navCurrency="DKK", vintage=2020,
                     nav=1000.0, marketValue=1000.0, commitment=2000.0, unfundedCommitment=0.0,
                     drawdowns=800.0, distributions=100.0, calledPercent=0.4, dpi=0.125,
                     rvpi=1.25, tvpi=1.375, irr=0.1), **k)


def _rapport(valuta="DKK", **n):
    return dict(fonde=[dict(fund_code="KAPE20", valuta=valuta, vintage=2020, n=dict(dict(
        nav=1000.0, nav_dkk=1000.0, tilsagn_dkk=2000.0, resttilsagn_dkk=5.0, indbetalt=800.0,
        indbetalt_dkk=800.0, udloddet=100.0, udloddet_dkk=100.0, picc=0.4, dpi=0.125,
        tvpi=1.375, irr=0.1), **n))])


def test_afstem_status():
    status = {x["kontrol"]: x["status"] for x in afstem([_aleta()], _rapport())}
    assert set(status.values()) == {"ok"}
    # 0,4 % ved siden af er ok, 0,6 % afviger; IRR har en absolut tolerance på 0,25 pct.-point
    status = {x["kontrol"]: x["status"]
              for x in afstem([_aleta()], _rapport(nav=1004.0, nav_dkk=1006.0, irr=0.103))}
    assert (status["NAV"], status["NAV DKK"], status["IRR"]) == ("ok", "afviger", "afviger")
    assert status["RVPI"] == "ok"           # regnes af NAV og indbetalt: 1004 / 800


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
    # en fond, rapporten ikke har
    linjer = afstem([_aleta(securityCode="UKENDT")], _rapport())
    assert {x["status"] for x in linjer} == {"afviger"}
    assert linjer[0]["bemaerkning"] == "fonden er ikke i rapporten"
