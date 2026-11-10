"""Demorapporten: hele rapporten og de supplerende slides bygget af syntetiske data, uden
databasen. Fanger en slide, der ikke længere bygges, eller en forespørgsel, demodatabasen ikke
kan svare på."""
import io
import zipfile

import pytest
import yaml

from pe_fund_report import anonym, data, enhed, report
from pe_fund_report.demo import demo_kontekst, generator
from pe_fund_report.paths import FUND_INFO, PME_BENCHMARKS


@pytest.fixture(scope="module")
def demo():
    with demo_kontekst() as db:
        r = data.investor_rapport(generator.INVESTOR)
    return r, db


@pytest.fixture(scope="module")
def decks(demo):
    r, _ = demo
    return report._byg_deck(r, med_pme=True), report._byg_supplerende(r)


def _titler(prs):
    """Den første tekst på hver slide: titlen (forsiden: rapportens navn)."""
    ud = []
    for s in prs.slides:
        tekster = [sh.text_frame.text.strip() for sh in s.shapes
                   if sh.has_text_frame and len(sh.text_frame.text.strip()) > 3]
        ud.append(tekster[0] if tekster else "")
    return ud


RAPPORT = ["Executive summary", "Udvikling i kvartalet og året", "Likviditetsoverblik",
           "Porteføljen mod markedet", "Fondene mod markedet", "Fondene mod relevante indeks",
           "Udvikling efter kvartalets afslutning",
           "Udvikling efter kvartalets afslutning – pr. fond", "Investeringer",
           "Største værdibevægelser i porteføljen", "Ændringer pr. fond", "Eksponering",
           "Porteføljefordeling efter investeringsstrategi", "Branchefordeling – hele porteføljen",
           "Branchefordeling pr. aktivklasse", "Geografisk eksponering – hele porteføljen",
           "Geografisk eksponering pr. aktivklasse", "Største positioner", "Porteføljeoversigt",
           "Private Equity-fond 1", "Private Equity-fond 2", "Private Equity-fond 3",
           "Fund of funds 1", "Secondaries-fond 1", "Private Credit-fond 1", "Ejendomsfond 1",
           "Ejendomsfond 2", "Definitioner", "GICS-sektorer og branchegrupper"]
SUPPLERENDE = ["Regnskabstal – kvartalet", "Regnskabstal – år til dato", "Likviditetsoverblik",
               "Porteføljen mod markedet", "Fondene mod markedet", "Fondene mod relevante indeks",
               "Kapitalkald og udlodninger pr. år", "TVPI og DPI over tid",
               "Resttilsagn og kaldstempo", "Kapitalkald, resttilsagn og eksponering over tid",
               "Forventede kapitalkald og udlodninger", "Forventet NAV og resttilsagn",
               "Prognose pr. fond", "Bag de største værdibevægelser",
               "Største positioner – koncentration og risiko",
               "Ejendomsfond 1 – ejendommene bag NAV", "Ejendomsfond 1 – hvad har drevet afkastet",
               "Lange tilsagn", "Geografisk fordeling efter tilsagn", "Branchefordeling",
               "Geografisk eksponering", "Branchefordeling – ejendomme for sig",
               "Geografisk eksponering – ejendomme for sig", "Definitioner"]


def test_rapporten_har_alle_slides(decks):
    titler = _titler(decks[0])
    assert [t for t in RAPPORT if t not in titler] == []
    # virksomhedskortene for de tre fonde med porteføljeselskaber
    assert sum(t.startswith("Porteføljeselskaber") for t in titler) >= 3


def test_supplerende_har_alle_slides(decks):
    titler = _titler(decks[1])
    assert [t for t in SUPPLERENDE if t not in titler] == []


def test_porteføljen_er_realistisk(demo):
    r, _ = demo
    assert generator.kontrol(r) == []


def test_fond_bag_rapportdatoen(demo):
    r, _ = demo
    assert [x["fond"] for x in r["datotjek"]["foraeldede"]] == ["Private Equity-fond 2"]
    assert r["justeret_nav"] is not None


def test_ingen_rigtige_fondskoder(decks):
    with open(FUND_INFO, encoding="utf-8") as fh:
        koder = set(yaml.safe_load(fh) or {})
    with open(PME_BENCHMARKS, encoding="utf-8") as fh:
        koder |= set((yaml.safe_load(fh) or {}).get("fonde") or {})
    for prs in decks:
        assert anonym.laek(prs, sorted(koder)) == []


def test_demodatabasen_faar_kun_det_den_kan(demo):
    _, db = demo
    assert db.forespoergsler
    assert not any("OUTER APPLY" in sql.upper() for sql in db.forespoergsler)


def test_konteksten_rydder_op():
    foer = data.read_sql
    with demo_kontekst():
        assert data.read_sql is not foer
    assert data.read_sql is foer


def test_cli_demo(monkeypatch):
    from pe_fund_report import cli
    kald = {}
    monkeypatch.setattr(cli, "byg", lambda *a, **k: kald.update(a=a, k=k))
    cli.main(["--demo", "--med-pme"])
    assert kald["k"]["demo"] is True and kald["k"]["med_pme"] is True
    cli.main(["--demo", "--tusinder"])
    assert kald["k"]["tusinder"] is True
    with pytest.raises(SystemExit):
        cli.main(["--demo", "--aleta"])
    with pytest.raises(SystemExit):
        cli.main([])


def _xml_med(prs, tekst):
    """Delene i decket (slides, diagrammer, noter), hvor tekst står."""
    buf = io.BytesIO()
    prs.save(buf)
    with zipfile.ZipFile(buf) as z:
        return [n for n in z.namelist() if n.endswith(".xml")
                and tekst in z.read(n).decode("utf-8")]


def test_tusinder_uden_mio():
    """--tusinder: intet sted i nogen af de to decks må der stå »mio.« - en ny hårdkodet enhed
    fanges her."""
    with demo_kontekst(), enhed.beloebsenhed(enhed.TUSINDE):
        r = data.investor_rapport(generator.INVESTOR)
        decks = report._byg_deck(r, med_pme=True), report._byg_supplerende(r)
    for prs in decks:
        assert _xml_med(prs, "mio.") == []
        assert _xml_med(prs, "t.DKK")
