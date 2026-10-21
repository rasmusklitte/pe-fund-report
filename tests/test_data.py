"""Beregninger i data.py, som ikke kræver databasen."""
from datetime import date

import pandas as pd
import pytest

from pe_fund_report import data


def test_xirr_et_aar_med_ti_procent():
    flows = [(date(2024, 1, 1), -100.0), (date(2025, 1, 1), 110.0)]
    assert data.xirr(flows) == pytest.approx(0.10, abs=1e-3)   # 366 dage, act/365


def test_xirr_uden_fortegnsskifte_er_none():
    assert data.xirr([(date(2024, 1, 1), -100.0), (date(2025, 1, 1), -5.0)]) is None
    assert data.xirr([]) is None


def _fond(valuta, tilsagn, indbetalt, udloddet, nav, flows, recallable=0.0):
    n = dict(tilsagn=tilsagn, indbetalt=indbetalt, udloddet=udloddet, nav=nav,
             recallable=recallable, resttilsagn=tilsagn - indbetalt + recallable)
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


def test_total_er_none_naar_et_beloeb_mangler():
    a = _fond("DKK", 100, 50, 10, 60, [])
    b = _fond("DKK", 200, 150, 0, 150, [])
    b["n"].update(tilsagn=None, resttilsagn=None)
    t = data.total([a, b])
    # som 0 ville tilsagnet give PICC 200 % og et for lavt resttilsagn
    assert t["tilsagn"] is None and t["resttilsagn"] is None and t["picc"] is None
    assert t["indbetalt"] == 200 and t["tvpi"] == pytest.approx(220 / 200)


def _ka(*raekker):
    return pd.DataFrame(raekker, columns=["period_end", "basis", "account_code", "amount"])


def test_gp_indbetalt_bruger_kun_ca110_paa_nav_datoen():
    nav_dato = date(2026, 6, 30)
    ka = _ka((date(2026, 3, 31), "ITD", "CA110", 900.0), (nav_dato, "POINT", "MM110", 1000.0))
    # et ældre CA110 mangler de senere kald, og MM110 er kaldt kapital inkl. rabat
    assert data._gp_indbetalt(ka, nav_dato) is None
    ka = _ka((nav_dato, "ITD", "CA110", 950.0), (nav_dato, "POINT", "MM110", 1000.0))
    assert data._gp_indbetalt(ka, nav_dato) == 950.0


def test_gp_vinder_ikke_naar_forskellen_er_genkaldelige_udlodninger():
    # MIF3 pr. 30.06.2026: kaldene summerer til MM110, CA110 er fratrukket den genkaldelige
    # udlodning, som her tæller som udlodning - GP-tallet ville trække den fra to gange
    assert not data.gp_vinder(4_089_038.0, 3_698_585.0, 390_453.09)
    assert not data.gp_vinder(22_734_924.0, 22_734_924.0, 0.0)
    assert not data.gp_vinder(100.0, None, 0.0)
    # et manglende kald i databasen: forvalterens tal vinder
    assert data.gp_vinder(20_000_000.0, 21_000_000.0, 0.0)


def _cf(*beloeb):
    return pd.DataFrame({"amount": list(beloeb)})


def test_estimerede_cashflows_hoerer_til_investoren():
    info = {"estimerede_cashflows": [
        dict(investor="A", dato="2024-06-30", beloeb=-100, note="Kald 1"),
        dict(investor="B", dato="2024-06-30", beloeb=-200, note="Kald 2"),
        dict(investor="A", dato="2026-09-30", beloeb=-300, note="Efter NAV")]}
    est, advarsler = data.estimerede_cashflows(info, "A", _cf(-50.0), date(2026, 6, 30))
    assert [e["note"] for e in est] == ["Kald 1"] and advarsler == []


def test_registreret_skoen_medregnes_ikke_igen():
    info = {"estimerede_cashflows": [
        dict(investor="A", dato="2024-06-30", beloeb=-943382, note="Draw down 14")]}
    est, advarsler = data.estimerede_cashflows(info, "A", _cf(-943382.0), date(2026, 6, 30))
    assert est == [] and "Draw down 14" in advarsler[0]


def test_skoen_uden_investor_afvises():
    info = {"estimerede_cashflows": [dict(dato="2024-06-30", beloeb=-1, note="x")]}
    with pytest.raises(ValueError):
        data.estimerede_cashflows(info, "A", _cf(), date(2026, 6, 30))


def test_afstemning_mod_forvalteren():
    n = dict(tilsagn=1000.0, resttilsagn=500.0, tvpi=1.311, gp_tvpi=1.30)
    assert data.afstem_mod_forvalteren(n, 502.0, "EUR") == []     # afrunding
    advarsler = data.afstem_mod_forvalteren(n | {"gp_tvpi": 1.25}, 450.0, "EUR")
    assert len(advarsler) == 2 and "MM120" in advarsler[0] and "MM160" in advarsler[1]
    assert data.afstem_mod_forvalteren(n | {"gp_tvpi": None}, None, "EUR") == []


def test_fund_info_har_de_forventede_felter():
    for kode, info in data.fund_info().items():
        assert {"status", "beskrivelse"} <= set(info), kode
        # cash flows hører til investor × fond, så et skøn skal sige, hvis kald det er
        assert all("investor" in e for e in info.get("estimerede_cashflows", [])), kode
        # stamdata fra rpt.v_fund_master hører ikke hjemme her; to kilder kunne være uenige
        assert not {"geografi", "segment", "loebetid", "fondsstoerrelse",
                    "forvalter"} & set(info), kode


def test_loebetid_skriver_forlaengelser_som_etaarige():
    assert data.loebetid(10, 2) == "10+1+1"
    assert data.loebetid(10.0, 3.0) == "10+1+1+1"
    assert data.loebetid(12, float("nan")) == "12"
    assert data.loebetid(None, 2) is None
    assert data.loebetid(float("nan"), None) is None


def test_geografi_oversaetter_alle_databasens_koder():
    # kodelisten er check constraint'en på fund.fund.geography_focus
    for kode in ("DENMARK", "NORDICS", "EUROPE", "NORTH_AMERICA", "EUROPE_NORTH_AMERICA",
                 "GLOBAL"):
        assert data.geografi(kode) != kode
    assert data.geografi("EUROPE_NORTH_AMERICA") == "Europa og Nordamerika"
    assert data.geografi(None) is None


def test_strategi_noegle_sorterer_efter_strategi_og_saa_navn():
    fonde = [dict(fund_name=n, segment=s) for n, s in [
        ("A ejendom", "Ejendomme"), ("B kredit", "Private credit"), ("C ukendt", "Shipping"),
        ("D sek", "Secondaries"), ("E fof", "Fund of funds"), ("F pe", "Private equity"),
        ("A pe", "Private equity"), ("G intet", None)]]
    assert [f["fund_name"] for f in sorted(fonde, key=data.strategi_noegle)] == [
        "A pe", "F pe", "E fof", "D sek", "B kredit", "A ejendom", "C ukendt", "G intet"]


def test_segmentfordeling_staar_i_strategiorden_ikke_efter_stoerrelse():
    fond = lambda navn, segment, tilsagn: dict(
        fund_name=navn, segment=segment, valuta="DKK",
        n=dict(tilsagn=tilsagn, tilsagn_dkk=tilsagn, indbetalt=0.0, indbetalt_dkk=0.0,
               udloddet=0.0, udloddet_dkk=0.0, nav=0.0, nav_dkk=0.0, resttilsagn=tilsagn,
               resttilsagn_dkk=tilsagn, tvpi=None, irr=None))
    fd = data.fordeling([fond("Ejd", "Ejendomme", 900.0), fond("PE", "Private equity", 100.0)],
                        "segment")
    assert [g["navn"] for g in fd["grupper"]] == ["Private equity", "Ejendomme"]


def test_segment_bruger_struktur_foer_fokus():
    assert data.segment("FUND_OF_FUNDS", "BUYOUT", None) == "Fund of funds"
    assert data.segment("SECONDARY", "BUYOUT", "SMALL_MID") == "Secondaries"
    # buyout, venture og growth er samlet Private equity, uanset størrelsesfokus
    assert data.segment("DIRECT", "BUYOUT", "SMALL_MID") == "Private equity"
    assert data.segment("DIRECT", "BUYOUT", "ALL") == "Private equity"
    assert data.segment("DIRECT", "VENTURE", float("nan")) == "Private equity"
    assert data.segment("DIRECT", "GROWTH", None) == "Private equity"
    assert data.segment("DIRECT", "JUNIOR_DEBT", "MID_LARGE") == "Private credit"
    # private credit er sin egen aktivklasse, også som fund of funds
    assert data.segment("FUND_OF_FUNDS", "JUNIOR_DEBT", "SMALL_MID") == "Private credit"
    assert data.segment("DIRECT", "REAL_ESTATE", None) == "Ejendomme"
    assert data.segment(None, None, None) is None


def test_fondsstoerrelse_dropper_gammel_metric_for_stamdata():
    nav_dato = date(2026, 6, 30)
    frisk = (1_044.8e6, date(2026, 6, 30), 7)
    gammel = (52.9e6, date(2023, 6, 30), 8)
    assert data.fondsstoerrelse(frisk, 2_000e6, nav_dato) == (
        1_044.8e6, dict(tekst="Fondsrapport pr. 30.06.2026 (TOTAL_COMMITMENTS)", document_id=7))
    # Home.Earth: tilsagnene er vokset siden den daterede opgørelse
    assert data.fondsstoerrelse(gammel, 158.2e6, nav_dato)[0] == 158.2e6
    assert data.fondsstoerrelse(gammel, None, nav_dato)[0] == 52.9e6
    assert data.fondsstoerrelse(None, 158.2e6, nav_dato)[1]["document_id"] is None


def test_vaerdibro_uden_opdeling_viser_ikke_nuller():
    ka = pd.DataFrame(dict(period_end=[date(2026, 6, 30)] * 3, basis=["ITD"] * 3,
                           account_code=["CA110", "CA290", "CA900"],
                           amount=[6_045_470.0, 1_163_069.0, 7_208_539.0]))
    n = dict(indbetalt=6_045_470.0, udloddet=0.0, nav=7_208_539.0)
    bro = data._vaerdibro(ka, date(2026, 6, 30), n)
    assert [p for p, _ in bro] == ["Indbetalt kapital", "Udlodninger",
                                   "Værdiskabelse efter omkostninger og carry", "NAV"]
    assert sum(v for _, v in bro[:-1]) == pytest.approx(n["nav"])


def _beholdning(type_, raekker):
    return dict(dato=date(2026, 6, 30), raekker=raekker, skaleret=True, type=type_)


def _linje(navn, dagsvaerdi, kostpris=None, tvpi=None):
    return dict(navn=navn, kostpris=kostpris, dagsvaerdi=dagsvaerdi, udloddet=0.0,
                total=dagsvaerdi, tvpi=tvpi)


def _fond_med_beholdninger(valuta, navn, nav, beholdninger):
    f = _fond(valuta, 100, 50, 0, nav, [])
    f["fund_name"] = navn
    f["beholdninger"] = beholdninger
    return f


def test_gennemlysning_sorterer_og_taeller():
    a = _fond_med_beholdninger("DKK", "KAPE 2020", 60, _beholdning(
        "fonde", [_linje("ASF VII", 300.0), _linje("Charlesbank X", 100.0)]))
    b = _fond_med_beholdninger("DKK", "CataCap III", 40, _beholdning(
        "selskaber", [_linje("The Army Painter", 200.0)]))
    g = data.gennemlysning([a, b])
    assert [x["navn"] for x in g["linjer"]] == ["ASF VII", "The Army Painter", "Charlesbank X"]
    assert [x["fond"] for x in g["linjer"]][:2] == ["KAPE 2020", "CataCap III"]
    assert (g["antal_fonde"], g["antal_underfonde"], g["antal_selskaber"]) == (2, 2, 1)
    assert g["dagsvaerdi"] == pytest.approx(600.0)
    assert g["nav"] == pytest.approx(100.0)        # summen af fondenes NAV, ikke beholdningerne
    assert g["uden_beholdninger"] == []


def test_gennemlysning_er_none_ved_blandede_valutaer():
    a = _fond_med_beholdninger("DKK", "A", 10, _beholdning("fonde", [_linje("X", 1.0)]))
    b = _fond_med_beholdninger("EUR", "B", 10, _beholdning("fonde", [_linje("Y", 1.0)]))
    assert data.gennemlysning([a, b]) is None


def test_gennemlysning_er_none_uden_beholdninger():
    a = _fond("DKK", 100, 50, 0, 60, [])
    assert data.gennemlysning([a]) is None


def test_gennemlysning_noterer_fonde_uden_beholdninger():
    a = _fond_med_beholdninger("DKK", "Med", 60, _beholdning("fonde", [_linje("X", 5.0)]))
    b = _fond("DKK", 100, 50, 0, 40, [])
    b["fund_name"] = "Uden"
    g = data.gennemlysning([a, b])
    assert g["uden_beholdninger"] == ["Uden"] and g["antal_fonde"] == 2


def test_total_summerer_recallables():
    a = _fond("DKK", 100, 60, 10, 50, [], recallable=10)
    b = _fond("DKK", 200, 150, 0, 150, [])
    t = data.total([a, b])
    assert t["recallable"] == 10
    # resttilsagn = tilsagn - indbetalt + recallables, så vandfaldet går op på totalrækken
    assert t["resttilsagn"] == t["tilsagn"] - t["indbetalt"] + t["recallable"]


def _fond_med_datoer(navn, nav_dato, beholdninger_dato=None):
    f = _fond("DKK", 100, 50, 0, 60, [])
    f["fund_name"] = navn
    f["nav_dato"] = nav_dato
    f["beholdninger"] = (None if beholdninger_dato is None
                         else _beholdning("fonde", [_linje("X", 1.0)]) | {"dato": beholdninger_dato})
    return f


def test_datotjek_er_ens_naar_alle_fonde_er_opgjort_samme_dag():
    d = date(2026, 6, 30)
    t = data.datotjek([_fond_med_datoer("A", d, d), _fond_med_datoer("B", d, d)], d)
    assert t["ens"] and t["foraeldede"] == [] and t["skaeve_beholdninger"] == []


def test_datotjek_fanger_en_fond_der_halter():
    d, gammel = date(2025, 12, 31), date(2025, 9, 30)
    t = data.datotjek([_fond_med_datoer("A", d, d), _fond_med_datoer("Seed", gammel, gammel)], d)
    assert not t["ens"]
    assert t["foraeldede"] == [dict(fond="Seed", dato=gammel, kvartaler=1)]


def test_datotjek_maaler_afstanden_i_kvartaler_over_aarsskifte():
    t = data.datotjek([_fond_med_datoer("A", date(2025, 3, 31))], date(2026, 6, 30))
    assert t["foraeldede"][0]["kvartaler"] == 5


def test_datotjek_fanger_beholdninger_paa_anden_dato_end_nav():
    d = date(2026, 6, 30)
    t = data.datotjek([_fond_med_datoer("A", d, date(2026, 3, 31))], d)
    assert not t["ens"] and t["foraeldede"] == []
    assert t["skaeve_beholdninger"] == [
        dict(fond="A", nav_dato=d, beholdninger_dato=date(2026, 3, 31))]


def test_datotjek_ignorerer_fonde_uden_beholdninger():
    d = date(2026, 6, 30)
    assert data.datotjek([_fond_med_datoer("A", d, None)], d)["ens"]


class _FalskPrs:
    """Presentation-stub, der lader som om de første n filnavne er låst."""

    def __init__(self, laaste):
        self.laaste, self.gemt = set(laaste), None

    def save(self, sti):
        from pathlib import Path
        if Path(sti).name in self.laaste:
            raise PermissionError(13, "Permission denied", sti)
        self.gemt = sti


def test_gem_bruger_det_oenskede_navn_naar_filen_er_fri(tmp_path):
    from pe_fund_report.report import _gem
    prs = _FalskPrs([])
    sti = _gem(prs, tmp_path / "MOTORTRAMP_2026-06-30.pptx")
    assert sti.name == "MOTORTRAMP_2026-06-30.pptx" and prs.gemt == str(sti)


def test_gem_lægger_versionsnummer_paa_naar_filen_er_laast(tmp_path):
    from pe_fund_report.report import _gem
    prs = _FalskPrs(["MOTORTRAMP_2026-06-30.pptx", "MOTORTRAMP_2026-06-30_v2.pptx"])
    sti = _gem(prs, tmp_path / "MOTORTRAMP_2026-06-30.pptx")
    assert sti.name == "MOTORTRAMP_2026-06-30_v3.pptx"
    assert sti.parent == tmp_path          # versionen lander samme sted som originalen


def test_gem_giver_op_naar_alt_er_laast(tmp_path):
    from pe_fund_report.report import MAKS_VERSIONER, _gem
    alle = ["MOTORTRAMP.pptx"] + [f"MOTORTRAMP_v{i}.pptx" for i in range(2, MAKS_VERSIONER + 1)]
    with pytest.raises(PermissionError, match="luk de åbne udgaver"):
        _gem(_FalskPrs(alle), tmp_path / "MOTORTRAMP.pptx")


def test_fund_info_forvalter_peger_paa_en_kendt_forvalterkode():
    # 'forvalter' overstyrer logoet; en tastefejl her ville fjerne logoet uden at fejle,
    # så koden skal se ud som en manager_code og ikke som et fondsnavn
    for kode, info in data.fund_info().items():
        f = info.get("forvalter")
        assert f is None or (f.isupper() and " " not in f), (kode, f)


def _dkk_fond(rest, nav):
    plus = (lambda x: None) if rest is None else (lambda x: rest + x)
    return dict(n=dict(tilsagn_dkk=plus(500.0), indbetalt_dkk=plus(400.0), udloddet_dkk=plus(0.0),
                       resttilsagn_dkk=rest, nav_dkk=nav))


def test_dkk_total_summerer_paa_tvaers_af_valutaer():
    t = data.dkk_total([_dkk_fond(100.0, 200.0), _dkk_fond(746.0, 1492.0)])
    assert t == dict(tilsagn_dkk=1846.0, indbetalt_dkk=1646.0, udloddet_dkk=846.0,
                     resttilsagn_dkk=846.0, nav_dkk=1692.0)


def test_dkk_total_er_none_naar_en_fond_mangler_kurs():
    assert data.dkk_total([_dkk_fond(100.0, 200.0), _dkk_fond(None, None)]) is None


def test_valutafordeling_summerer_pr_valuta_og_eksponering_er_nav_plus_rest():
    fond = lambda valuta, nav, rest: dict(valuta=valuta, n=dict(nav_dkk=nav, resttilsagn_dkk=rest))
    vf = data.valutafordeling([fond("DKK", 600.0, 100.0), fond("EUR", 300.0, 300.0),
                               fond("DKK", 100.0, 0.0)])
    assert [(x["valuta"], x["dkk"]) for x in vf["nav"]] == [("DKK", 700.0), ("EUR", 300.0)]
    assert [(x["valuta"], x["dkk"]) for x in vf["resttilsagn"]] == [("EUR", 300.0), ("DKK", 100.0)]
    assert [(x["valuta"], x["dkk"]) for x in vf["eksponering"]] == [("DKK", 800.0), ("EUR", 600.0)]
    for maal in vf.values():
        assert sum(x["andel"] for x in maal) == pytest.approx(1.0)


def test_valutafordeling_er_none_naar_en_fond_mangler_dkk_tal():
    assert data.valutafordeling([dict(valuta="EUR", n=dict(nav_dkk=None,
                                                           resttilsagn_dkk=None))]) is None
    nul = data.valutafordeling([dict(valuta="DKK", n=dict(nav_dkk=5.0, resttilsagn_dkk=0.0))])
    assert nul["resttilsagn"] == []


def _aktiv_fond(navn, segment, nav, brancher):
    return dict(fund_name=navn, segment=segment, nav_dato=date(2026, 6, 30),
                n=dict(nav_dkk=nav),
                brancher=dict(grundlag="HOLDINGS", dato=date(2026, 6, 30), gennem=None, linjer=[
                    dict(top_kode=k, top=k.title(), under_kode=None, under=None, andel=a,
                         foreloebig=0.0) for k, a in brancher]))


def test_opdel_aktivklasser_summerer_til_nav_og_placerer_kredit_og_ejendomme():
    fonde = [_aktiv_fond("Ejendom", "Ejendomme", 200.0, [("REAL_ESTATE", 1.0)]),
             _aktiv_fond("PE", "Buyout", 500.0, [("FINANCIALS", 0.9), ("REAL_ESTATE", 0.1)]),
             _aktiv_fond("Kredit", data.FOKUS["JUNIOR_DEBT"], 300.0, [("INDUSTRIALS", 1.0)])]
    opd = data.opdel_aktivklasser(fonde, "brancher")
    dele = {d["navn"]: d for d in opd["dele"]}
    assert dele[data.PRIVATE_EQUITY]["total_dkk"] == pytest.approx(450.0)   # uden ejendomme
    assert dele[data.PRIVATE_CREDIT]["total_dkk"] == pytest.approx(300.0)
    assert opd["ejendomme"]["total_dkk"] == pytest.approx(250.0)            # inkl. PE's 10 %
    assert sum(d["andel"] for d in opd["dele"]) + opd["ejendomme"]["andel"] == pytest.approx(1.0)
    assert [g["kode"] for g in dele[data.PRIVATE_CREDIT]["fordeling"]["grupper"]] == ["INDUSTRIALS"]


def test_opdel_aktivklasser_udelader_tom_del_og_klarer_sig_uden_ejendomme():
    opd = data.opdel_aktivklasser([_aktiv_fond("PE", "Buyout", 500.0, [("FINANCIALS", 1.0)])],
                                  "brancher")
    assert opd["ejendomme"] is None
    assert [d["navn"] for d in opd["dele"]] == [data.PRIVATE_EQUITY]
    assert data.opdel_aktivklasser([dict(fund_name="X", segment=None, n=dict(nav_dkk=None))],
                                   "brancher") is None


def test_omregn_til_dkk_bruger_valutaens_kurs():
    f = dict(valuta="EUR", n=dict(tilsagn=30.0, indbetalt=20.0, udloddet=4.0, resttilsagn=10.0,
                                        nav=20.0), noter=[])
    data.omregn_til_dkk([f], {"EUR": (7.5, None)})
    assert f["n"]["resttilsagn_dkk"] == 75.0 and f["n"]["nav_dkk"] == 150.0
    assert f["n"]["tilsagn_dkk"] == 225.0 and f["n"]["udloddet_dkk"] == 30.0
    g = dict(valuta="USD", n=dict(tilsagn=30.0, indbetalt=20.0, udloddet=4.0, resttilsagn=10.0,
                                        nav=20.0), noter=[])
    data.omregn_til_dkk([g], {"USD": None})
    assert g["n"]["nav_dkk"] is None and g["noter"]


def _ford_fond(navn, geografi, valuta, kurs, tilsagn, indbetalt, udloddet, nav):
    n = dict(tilsagn=tilsagn, indbetalt=indbetalt, udloddet=udloddet, nav=nav,
             tvpi=(udloddet + nav) / indbetalt, tilsagn_dkk=tilsagn * kurs, nav_dkk=nav * kurs,
             indbetalt_dkk=indbetalt * kurs, udloddet_dkk=udloddet * kurs,
             resttilsagn_dkk=(tilsagn - indbetalt) * kurs)
    return dict(fund_name=navn, geografi=geografi, valuta=valuta, n=n)


def _ford_fonde():
    return [_ford_fond("A", "Danmark", "DKK", 1.0, 100.0, 50.0, 10.0, 60.0),
            _ford_fond("B", "Europa", "EUR", 7.5, 40.0, 20.0, 0.0, 30.0),
            _ford_fond("C", "Europa", "DKK", 1.0, 200.0, 100.0, 0.0, 110.0),
            _ford_fond("D", None, "DKK", 1.0, 0.0, 1.0, 0.0, 0.0)]


def test_fordeling_summerer_dkk_og_regner_andele_af_tilsagnet():
    fd = data.fordeling(_ford_fonde(), "geografi")
    assert [g["navn"] for g in fd["grupper"]] == ["Europa", "Danmark", "Ikke oplyst"]
    eu = fd["grupper"][0]
    assert eu["sum"]["tilsagn_dkk"] == 500.0 and eu["sum"]["nav_dkk"] == 335.0
    assert fd["total"]["tilsagn_dkk"] == 600.0 and fd["total"]["andel"] == 1.0
    assert eu["andel"] == pytest.approx(500 / 600)
    assert sum(g["andel"] for g in fd["grupper"]) == pytest.approx(1.0)
    assert fd["andele"]["B"] == pytest.approx(300 / 600)


def test_fordeling_viser_kun_tilsagn_i_fondens_valuta_naar_gruppen_har_een_valuta():
    fd = data.fordeling(_ford_fonde(), "geografi")
    eu, dk = fd["grupper"][0], fd["grupper"][1]
    assert eu["sum"]["valuta"] is None and eu["sum"]["tilsagn"] is None
    assert dk["sum"]["valuta"] == "DKK" and dk["sum"]["tilsagn"] == 100.0


def test_fordeling_regner_tvpi_i_dkk_paa_tvaers_af_valutaer():
    fd = data.fordeling(_ford_fonde(), "geografi")
    eu, dk = fd["grupper"][0], fd["grupper"][1]
    # Europa: B (EUR, kurs 7,5) og C (DKK); (udloddet + NAV) / indbetalt, alt i DKK
    assert eu["sum"]["tvpi"] == pytest.approx((0 + 30 * 7.5 + 110) / (20 * 7.5 + 100))
    assert dk["sum"]["tvpi"] == pytest.approx(70 / 50)        # én valuta: som før
    assert fd["total"]["tvpi"] == pytest.approx((10 + 225 + 110 + 60) / (50 + 150 + 100 + 1))


def test_fordeling_uden_dkk_kurs_har_ingen_andele():
    fonde = _ford_fonde()
    fonde[1]["n"]["tilsagn_dkk"] = None
    fd = data.fordeling(fonde, "geografi")
    assert fd["total"]["tilsagn_dkk"] is None
    assert fd["total"]["andel"] is None and all(g["andel"] is None for g in fd["grupper"])


def _sam_fonde():
    fonde = _ford_fonde()
    for f, (manager, vintage) in zip(fonde, [("X", 2020), ("Y", 2023), ("X", 2021), ("Z", None)]):
        f.update(manager=manager, vintage=vintage)
        f["n"]["irr"] = 0.1
    return fonde


def test_sammendrag_regner_portefoeljens_noegletal_i_dkk():
    fonde = _sam_fonde()
    sm = data.sammendrag(fonde, data.dkk_total(fonde))
    assert (sm["antal_fonde"], sm["antal_forvaltere"], sm["vintage"]) == (4, 3, (2020, 2023))
    assert sm["tilsagn_dkk"] == 600.0 and sm["nav_dkk"] == 395.0
    assert sm["picc"] == pytest.approx(301 / 600)
    assert sm["dpi"] == pytest.approx(10 / 301)
    assert sm["tvpi"] == pytest.approx((10 + 395) / 301)


def test_sammendrag_finder_hoejeste_og_laveste_tvpi():
    sm = data.sammendrag(_sam_fonde(), None)
    assert sm["bedst"]["fund_name"] == "B" and sm["svagest"]["fund_name"] == "D"


def test_sammendrag_uden_dkk_tal_har_ingen_portefoeljetal():
    sm = data.sammendrag(_sam_fonde(), None)
    assert sm["tilsagn_dkk"] is None and sm["tvpi"] is None and sm["picc"] is None


# ---------- udvikling i kvartalet og året ----------

def test_perioder_er_forrige_kvartals_og_aars_ultimo():
    p = data.perioder(date(2026, 6, 30))
    assert p["kvartal"] == (date(2026, 3, 31), date(2026, 6, 30))
    assert p["aar"] == (date(2025, 12, 31), date(2026, 6, 30))
    q1 = data.perioder(date(2026, 3, 31))
    assert q1["kvartal"] == q1["aar"] == (date(2025, 12, 31), date(2026, 3, 31))


def _historik(*punkter):
    """(dato, nav, indbetalt, udloddet) -> historik som i data.fond()."""
    return [dict(dato=d, nav=n, indbetalt=i, udloddet=u) for d, n, i, u in punkter]


def test_periodeaendring_gaar_op_og_regner_modified_dietz():
    h = _historik((date(2025, 12, 31), 100.0, 80.0, 0.0), (date(2026, 3, 31), 105.0, 80.0, 0.0),
                  (date(2026, 6, 30), 130.0, 100.0, 5.0))
    flows = [(date(2025, 6, 1), -80.0, "kald"), (date(2026, 5, 16), -20.0, "kald"),
             (date(2026, 5, 16), 5.0, "udlodning")]
    u = data._periodeaendring(h, flows, date(2026, 3, 31), date(2026, 6, 30))
    assert (u["primo"], u["kald"], u["udlodning"], u["ultimo"]) == (105.0, 20.0, 5.0, 130.0)
    assert u["primo"] + u["kald"] - u["udlodning"] + u["vaerdi"] == u["ultimo"]
    assert u["vaerdi"] == 10.0
    # nettoflowet på 15 er i fonden 45 af periodens 91 dage
    assert u["afkast"] == pytest.approx(10.0 / (105.0 + 15.0 * 45 / 91))
    assert u["tvpi_aendring"] == pytest.approx(135 / 100 - 105 / 80)


def test_periodeaendring_falder_tilbage_paa_seneste_nav_foer_perioden():
    f = dict(historik=_historik((date(2025, 9, 30), 50.0, 40.0, 0.0),
                                (date(2026, 3, 31), 60.0, 50.0, 0.0),
                                (date(2026, 6, 30), 62.0, 50.0, 0.0)),
             kapitalstroemme=[(date(2025, 11, 1), -10.0, "kald")], noter=[])
    data.udvikling(f, data.perioder(date(2026, 6, 30)))
    aar = f["udvikling"]["aar"]
    assert aar["primo_dato"] == date(2025, 9, 30)
    assert aar["kald"] == 10.0           # kaldet i Q4 25 tæller med, så broen går op
    assert aar["vaerdi"] == 2.0
    assert len(f["noter"]) == 1 and "30.09.2025" in f["noter"][0]
    assert f["udvikling"]["kvartal"]["primo_dato"] == date(2026, 3, 31)


def test_periodeaendring_uden_nav_foer_perioden_starter_fra_nul():
    h = _historik((date(2026, 6, 30), 9.0, 10.0, 0.0))
    u = data._periodeaendring(h, [(date(2026, 4, 15), -10.0, "kald")],
                              date(2026, 3, 31), date(2026, 6, 30))
    assert (u["primo"], u["kald"], u["vaerdi"]) == (0.0, 10.0, -1.0)


def _udv_fond(navn, valuta, primo, kald, ultimo, primo_dato=date(2026, 3, 31)):
    u = dict(fra=date(2026, 3, 31), til=date(2026, 6, 30), primo_dato=primo_dato, primo=primo,
             kald=kald, udlodning=0.0, vaerdi=ultimo - primo - kald, ultimo=ultimo,
             flows=[(date(2026, 5, 16), kald, "kald")] if kald else [], afkast=None)
    return dict(fund_name=navn, valuta=valuta, udvikling={"kvartal": u})


def test_omregning_til_dkk_giver_valutakursaendring_som_residual():
    kurser = {("EUR", date(2026, 3, 31)): 7.0, ("EUR", date(2026, 5, 16)): 7.2,
              ("EUR", date(2026, 6, 30)): 7.5}
    kurs = lambda v, d: 1.0 if v == "DKK" else kurser.get((v, d))
    a, b = _udv_fond("A", "DKK", 100.0, 10.0, 120.0), _udv_fond("B", "EUR", 10.0, 2.0, 13.0)
    data.omregn_udvikling_til_dkk([a, b], kurs)
    ua, ub = a["udvikling"]["kvartal"], b["udvikling"]["kvartal"]
    assert ua["valuta_dkk"] == 0.0 and ua["vaerdi_dkk"] == 10.0
    assert (ub["primo_dkk"], ub["kald_dkk"], ub["ultimo_dkk"]) == (70.0, 14.4, 97.5)
    assert ub["vaerdi_dkk"] == 7.5
    assert ub["valuta_dkk"] == pytest.approx(97.5 - 70.0 - 14.4 - 7.5)

    t = data.portefoelje_udvikling([a, b])["kvartal"]
    assert t["primo"] + t["kald"] - t["udlodning"] + t["vaerdi"] + t["valuta"] == \
        pytest.approx(t["ultimo"])
    assert [n for n, _ in t["bidrag"]] == ["B", "A"]      # 7,5 + 5,6 i valuta > 10


def test_fortegnsafvigelser_skelner_valuta_fra_dietz():
    def fond(navn, vaerdi, vist, afkast):
        u = dict(vaerdi=vaerdi, vaerdi_dkk=vist, valuta_dkk=0.0, afkast=afkast)
        return dict(fund_name=navn, udvikling={"aar": u})
    fonde = [fond("Valuta", -118.0, 297.0, -0.023),       # USD faldt i værdi, kursen vendte den
             fond("Dietz", 50.0, 50.0, -0.01),            # samme valuta, men modsat afkast
             fond("Ens", 10.0, 70.0, 0.02),
             fond("Nul", 0.0, 0.0, -0.01), fond("Uden kurs", 5.0, None, -0.01)]
    assert data.fortegnsafvigelser(fonde) == [("Valuta", "aar", "valuta"),
                                               ("Dietz", "aar", "dietz")]


def test_portefoelje_udvikling_er_none_uden_kurs():
    a = _udv_fond("A", "USD", 10.0, 0.0, 11.0)
    data.omregn_udvikling_til_dkk([a], lambda v, d: None)
    assert data.portefoelje_udvikling([a]) == {"kvartal": None}


def test_forkortelse_af_genkender_forvalternes_korte_navne():
    for kort, lang in [("ASF VII", "AlpInvest Secondaries Fund VII"),
                       ("ACF VIII", "AlpInvest Co-Investment Fund VIII"),
                       ("WPEF IX", "Waterland Private Equity Fund IX"),
                       ("Charlesbank X", "Charlesbank Equity Fund X"),
                       ("Advent Tech II", "Advent Global Technology II"),
                       ("Paragon IV", "The Paragon Fund IV")]:
        assert data._forkortelse_af(kort, lang), kort
    for kort, lang in [("Trivest VII", "Trivest Discovery Fund II"),
                       ("Waterland VIII", "Waterland Private Equity Fund IX"),
                       ("TA XIV", "TA XIV")]:
        assert not data._forkortelse_af(kort, lang), kort


def _beh_linje(navn, fv, kost, udl=0.0, vintage=2020, closing=None):
    return dict(navn=navn, dagsvaerdi=fv, kostpris=kost, udloddet=udl, vintage=vintage,
                closing=closing)


def test_vaerdibevaegelser_parrer_omdoebte_og_udelader_afhaendede():
    primo = _beholdning("fonde", [_beh_linje("AlpInvest Secondaries Fund VII", 100.0, 80.0),
                                  _beh_linje("Gammel Fond I", 50.0, 40.0)])
    nu = _beholdning("fonde", [_beh_linje("ASF VII", 90.0, 85.0, udl=5.0),
                               _beh_linje("Ny Fond I", 12.0, 10.0, closing=date(2026, 5, 1))])
    f = dict(fund_name="KAPE", valuta="DKK", beholdninger=nu,
             beholdninger_primo={"kvartal": primo},
             udvikling={"kvartal": dict(fra=date(2026, 3, 31))})
    vb = data.vaerdibevaegelser([f], {"DKK": (1.0, date(2026, 6, 30))})
    linjer = {x["navn"]: x["kvartal"] for x in vb["linjer"]}
    assert linjer == {"ASF VII": (90 + 5 - 85) - (100 - 80), "Ny Fond I": 2.0}
    assert vb["usammenlignelige"] == {"kvartal": [("KAPE", "Gammel Fond I")]}


def test_vaerdibevaegelser_markerer_fond_med_uaendrede_tal():
    beh = _beholdning("selskaber", [_beh_linje("A", 10.0, 5.0)])
    f = dict(fund_name="SEED", valuta="DKK", beholdninger=beh,
             beholdninger_primo={"kvartal": _beholdning("selskaber", [_beh_linje("A", 10.0, 5.0)])},
             udvikling={"kvartal": dict(fra=date(2026, 3, 31))})
    vb = data.vaerdibevaegelser([f], {"DKK": (1.0, date(2026, 6, 30))})
    assert vb["uaendrede"] == {"kvartal": ["SEED"]} and vb["linjer"] == []


def test_vaerdibevaegelser_udelader_ejendomsfond():
    # ejendommenes værdier er før gæld og uden kostpris, så de rangeres ikke
    beh = _beholdning("ejendomme", [_beh_linje("HTC", 38.0, None)]) | dict(skaleret=False)
    f = dict(fund_name="Home.Earth", valuta="EUR", beholdninger=beh,
             beholdninger_primo={"kvartal": _beholdning("ejendomme", [_beh_linje("HTC", 29.0, None)])},
             udvikling={"kvartal": dict(fra=date(2026, 3, 31))})
    vb = data.vaerdibevaegelser([f], {"EUR": (7.47, date(2026, 6, 30))})
    assert vb["linjer"] == [] and vb["ejendomsfonde"] == ["Home.Earth"]
    assert vb["uskalerede"] == []


# ---------- rabat på kald ----------

def test_kaldt_bruger_forvalterens_tal_ved_rabat():
    # PPPC2: kaldt 3.200.000 EUR, betalt 3.125.000 EUR (25 % rabat på første kald)
    assert data.kaldt(3_125_000.0, 3_200_000.0) == 3_200_000.0
    assert data.kaldt(20_511_966.0, 20_511_963.0) == 20_511_966.0      # afrunding, ikke rabat
    assert data.kaldt(22_734_924.0, None) == 22_734_924.0              # ingen MM110
    assert data.kaldt(100.0, 90.0) == 100.0                            # mindre kaldt: ikke rabat


# ---------- gennemlysning af feedere ----------

def test_selskab_fjerner_instrumentord():
    assert data._selskab("STADA Common") == "STADA"
    assert data._selskab("Pavion Junior Pref") == "Pavion"
    assert data._selskab("VPP Pref Equity") == "VPP"
    assert data._selskab("Equity") == "Equity"           # første ord bliver altid
    assert data._selskab("Summa Equity Fund III (No.2) AB") == "Summa Equity Fund III (No.2) AB"


def _holdings(rows, instrument=None, industri=None, forslag=None):
    """rows: (section, name, invested, fair_value) som i rpt.v_portfolio_holding."""
    n = len(rows)
    return pd.DataFrame(dict(
        report_date=[date(2026, 6, 30)] * n, section=[r[0] for r in rows],
        name=[r[1] for r in rows], segment=[None] * n, trade_date=[date(2025, 1, 1)] * n,
        vintage=[None] * n, instrument=instrument or [None] * n, commitment=[None] * n,
        invested=[r[2] for r in rows], distributions=[0.0] * n, fair_value=[r[3] for r in rows],
        paid_in_pct=[None] * n, multiple=[None] * n, irr_pct=[None] * n, document_id=[7] * n,
        industry_code=industri or [None] * n, industry_suggested=forslag or [None] * n))


def test_beholdninger_samler_trancher_og_skalerer_efter_nav(monkeypatch):
    h = _holdings([("UNREALISED", "STADA", 60.0, 66.0), ("UNREALISED", "STADA Common", 40.0, 54.0),
                   ("UNREALISED", "Opella", 100.0, 80.0)], instrument=["PIK", "Equity", "PIK"])
    monkeypatch.setattr(data, "read_sql", lambda *a, **k: h)
    beh = data._beholdninger("PSCP5", date(2026, 6, 30), None, nav=20.0)
    assert beh["kode"] == "PSCP5" and beh["skaleret"] and beh["type"] == "selskaber"
    assert [r["navn"] for r in beh["raekker"]] == ["STADA", "Opella"]
    stada = beh["raekker"][0]
    assert (stada["kostpris_fond"], stada["dagsvaerdi_fond"]) == (100.0, 120.0)
    assert stada["instrument"] == "PIK, Equity" and stada["tvpi"] == pytest.approx(1.2)
    assert beh["andel"] == pytest.approx(20.0 / 200.0)
    assert sum(r["dagsvaerdi"] for r in beh["raekker"]) == pytest.approx(20.0)   # = NAV


def test_beholdninger_hos_secondaryfond_er_underliggende_fonde(monkeypatch):
    h = _holdings([("SECONDARY", "Fund A", 10.0, 30.0), ("CO_INVESTMENT", "Selskab", 5.0, 10.0)])
    monkeypatch.setattr(data, "read_sql", lambda *a, **k: h)
    assert data._beholdninger("NEP6", date(2026, 6, 30), None, nav=4.0)["type"] == "fonde"


def test_vaerdibevaegelser_matcher_uden_hensyn_til_store_bogstaver():
    primo = _beholdning("selskaber", [_beh_linje("DiversiTech", 10.0, 8.0)])
    nu = _beholdning("selskaber", [_beh_linje("Diversitech", 12.0, 8.0)])
    f = dict(fund_name="PPPC2", valuta="DKK", beholdninger=nu,
             beholdninger_primo={"kvartal": primo},
             udvikling={"kvartal": dict(fra=date(2026, 3, 31))})
    vb = data.vaerdibevaegelser([f], {"DKK": (1.0, date(2026, 6, 30))})
    assert [x["kvartal"] for x in vb["linjer"]] == [2.0] and vb["usammenlignelige"] == {}


def test_udligningsrente_i_indbetalt_og_udloddet():
    # betalt udligningsrente er en indbetaling, modtaget en udlodning; kildeskat ingen af dem
    cf = pd.DataFrame(dict(
        flow_class=["CONTRIBUTION", "EQUALISATION", "DISTRIBUTION", "EQUALISATION",
                    "WITHHOLDING_TAX"],
        amount=[-1000.0, -17.0, 300.0, 5.0, -2.0]))
    assert -cf[data._er_indbetaling(cf)].amount.sum() == 1017.0
    assert cf[data._er_udlodning(cf)].amount.sum() == 305.0


def _branchelinje(sektor_kode, gruppe, andel, foreloebig=0.0):
    sektor = data.SEKTORER.get(sektor_kode, sektor_kode or data.IKKE_KLASSIFICERET)
    return dict(top_kode=sektor_kode, top=sektor, under_kode=None, under=gruppe,
                andel=andel, foreloebig=foreloebig)


def _branchefond(navn, nav_dkk, linjer, grundlag="GP_BREAKDOWN", dato=date(2026, 6, 30)):
    return dict(fund_name=navn, nav_dato=date(2026, 6, 30), n=dict(nav_dkk=nav_dkk),
                brancher=None if linjer is None else dict(grundlag=grundlag, dato=dato,
                                                          gennem=None, linjer=linjer))


def _branchefonde():
    return [
        _branchefond("A", 600.0, [_branchelinje("INDUSTRIALS", "Kapitalgoder", 0.5),
                                  _branchelinje("INDUSTRIALS", data.IKKE_OPDELT, 0.25),
                                  _branchelinje("HEALTH_CARE", "Sundhedsudstyr og -ydelser", 0.25, 1.0)]),
        _branchefond("B", 300.0, [_branchelinje("INDUSTRIALS", "Erhvervsservice", 1.0, 0.5)],
                     grundlag="HOLDINGS", dato=date(2025, 12, 31)),
        _branchefond("C", 100.0, None),
    ]


def test_eksponering_brancher_summerer_til_porteføljens_nav_i_dkk():
    b = data.eksponering(_branchefonde(), "brancher")
    assert b["total_dkk"] == 1000.0
    assert sum(s["dkk"] for s in b["grupper"]) == pytest.approx(1000.0)
    assert sum(s["andel"] for s in b["grupper"]) == pytest.approx(1.0)
    industri = b["grupper"][0]
    assert industri["navn"] == "Industrials" and industri["dkk"] == pytest.approx(750.0)
    assert sum(g["dkk"] for g in industri["under"]) == pytest.approx(industri["dkk"])
    assert industri["stoerste_fond"] == "A" and industri["fonde"] == ["A", "B"]


def test_eksponering_brancher_sorterer_ikke_opdelt_og_uklassificeret_sidst():
    b = data.eksponering(_branchefonde(), "brancher")
    assert [s["navn"] for s in b["grupper"]] == ["Industrials", "Health Care",
                                                  data.IKKE_KLASSIFICERET]
    assert [g["navn"] for g in b["grupper"][0]["under"]] == [
        "Kapitalgoder", "Erhvervsservice", data.IKKE_OPDELT]


def test_eksponering_brancher_fond_uden_fordeling_er_ikke_klassificeret():
    b = data.eksponering(_branchefonde(), "brancher")
    assert b["uden_fordeling"] == ["C"]
    assert b["grupper"][-1]["dkk"] == 100.0 and b["grupper"][-1]["under"] == []


def test_eksponering_brancher_vaegter_foreloebig_klassifikation_med_nav():
    b = data.eksponering(_branchefonde(), "brancher")
    # A: 25 % af 600 helt foreløbig; B: halvdelen af 300
    assert b["foreloebig_dkk"] == pytest.approx(150.0 + 150.0)
    assert b["foreloebig_andel"] == pytest.approx(0.3)
    assert b["foreloebige_fonde"] == ["A", "B"]


def test_eksponering_brancher_noterer_afvigende_dato_og_strategiplacering():
    fonde = _branchefonde()
    fonde.append(_branchefond("D", 0.0, [_branchelinje("REAL_ESTATE", data.IKKE_OPDELT, 1.0)],
                              grundlag="FUND_DEFAULT"))
    b = data.eksponering(fonde, "brancher")
    assert b["andre_datoer"] == [("B", date(2025, 12, 31))]
    assert b["efter_strategi"] == [("D", "Real Estate")]


def test_eksponering_brancher_uden_dkk_kurs_er_none():
    fonde = _branchefonde()
    fonde[1]["n"]["nav_dkk"] = None
    assert data.eksponering(fonde, "brancher") is None


def _ejendomsfonde():
    """En ren ejendomsfond, en blandet fond med 10 % ejendomme og en fond uden ejendomme."""
    fonde = [
        _branchefond("Ejd", 400.0, [_branchelinje("REAL_ESTATE", data.IKKE_OPDELT, 1.0)],
                     grundlag="FUND_DEFAULT"),
        _branchefond("Bland", 500.0, [_branchelinje("INDUSTRIALS", "Kapitalgoder", 0.9),
                                      _branchelinje("REAL_ESTATE", data.IKKE_OPDELT, 0.1)]),
        _branchefond("PE", 100.0, [_branchelinje("FINANCIALS", "Forsikring", 1.0)])]
    regioner = {"Ejd": [("EUROPE", "DENMARK", 1.0)], "Bland": [("EUROPE", "FRANCE", 0.6),
                                                               ("NORTH_AMERICA", "US", 0.4)],
                "PE": [("EUROPE", "DENMARK", 1.0)]}
    for f in fonde:
        f["regioner"] = _regionfond(f["fund_name"], 0.0, regioner[f["fund_name"]])["regioner"]
    return fonde


def test_opdel_ejendomme_fordeler_ejendommene_paa_segment_og_region():
    fonde = _ejendomsfonde()
    fonde[0]["re_sektor"] = "RESIDENTIAL"
    # brancher: den rene ejendomsfond under sit segment, den blandede fonds 50 uden segment sidst
    seg = data.opdel_ejendomme(fonde, "brancher")["ejendomme"]["fordeling"]
    assert [(g["navn"], g["dkk"]) for g in seg["grupper"]] == [
        ("Residential", 400.0), (data.SEGMENT_OEVRIGE, pytest.approx(50.0))]
    assert seg["uden_segment"] == ["Bland"]
    assert sum(g["andel"] for g in seg["grupper"]) == pytest.approx(1.0)
    # regioner: ejendomsandelen af hver fond fordelt som fonden; Bland 60/40 Europa/Nordamerika
    geo = data.opdel_ejendomme(fonde, "regioner")["ejendomme"]["fordeling"]
    assert {g["kode"]: g["dkk"] for g in geo["grupper"]} == {
        "EUROPE": pytest.approx(430.0), "NORTH_AMERICA": pytest.approx(20.0)}


def test_opdel_ejendomme_tager_sektoren_fra_alle_fonde_og_giver_resten_egen_base():
    opd = data.opdel_ejendomme(_ejendomsfonde(), "brancher")
    ejd, rest = opd["ejendomme"], opd["resten"]
    assert ejd["total_dkk"] == pytest.approx(450.0) and ejd["andel"] == pytest.approx(0.45)
    assert [(x["navn"], x["dkk"]) for x in ejd["fonde"]] == [("Ejd", 400.0),
                                                             ("Bland", pytest.approx(50.0))]
    assert sum(x["andel"] for x in ejd["fonde"]) == pytest.approx(1.0)
    assert ejd["blandede"] == ["Bland"]
    # resten: ingen ejendomme, og andelene er af resten (550), ikke af hele NAV
    assert rest["total_dkk"] == pytest.approx(550.0)
    assert ejd["total_dkk"] + rest["total_dkk"] == pytest.approx(1000.0)
    assert [(g["navn"], g["dkk"]) for g in rest["grupper"]] == [
        ("Industrials", pytest.approx(450.0)), ("Financials", pytest.approx(100.0))]
    assert sum(g["andel"] for g in rest["grupper"]) == pytest.approx(1.0)
    assert rest["efter_strategi"] == []           # den rene ejendomsfond er helt ude


def test_opdel_ejendomme_regioner_fordeler_den_blandede_fond_forholdsmaessigt():
    opd = data.opdel_ejendomme(_ejendomsfonde(), "regioner")
    rest = opd["resten"]
    assert rest["total_dkk"] == pytest.approx(550.0)
    europa, usa = rest["grupper"]
    # Bland: 450 tilbage, 60/40; PE: 100 i Danmark; ejendomsfondens Danmark er væk
    assert europa["dkk"] == pytest.approx(270.0 + 100.0)
    assert usa["dkk"] == pytest.approx(180.0)
    assert sum(g["andel"] for g in rest["grupper"]) == pytest.approx(1.0)


def test_opdel_ejendomme_uden_ejendomme_eller_dkk_er_none():
    fonde = _ejendomsfonde()[2:]
    assert data.opdel_ejendomme(fonde, "brancher") is None
    fonde = _ejendomsfonde()
    fonde[0]["n"]["nav_dkk"] = None
    assert data.opdel_ejendomme(fonde, "regioner") is None


def test_branchegrupper_daekker_alle_gics_grupper():
    assert len(data.BRANCHEGRUPPER) == 25
    assert set(data.SEKTORER) >= {"ENERGY", "MATERIALS", "INDUSTRIALS", "CONSUMER_DISCRETIONARY",
                                  "CONSUMER_STAPLES", "HEALTH_CARE", "FINANCIALS",
                                  "INFORMATION_TECHNOLOGY", "COMMUNICATION_SERVICES",
                                  "UTILITIES", "REAL_ESTATE"}


def _regionfond(navn, nav_dkk, linjer, grundlag="HOLDINGS"):
    return dict(fund_name=navn, nav_dato=date(2026, 6, 30), n=dict(nav_dkk=nav_dkk),
                regioner=dict(grundlag=grundlag, dato=date(2026, 6, 30), gennem=None, linjer=[
                    dict(top_kode=top, top=data.REGIONER[top], under_kode=under,
                         under=data.DELREGIONER[under], andel=andel, foreloebig=0.0)
                    for top, under, andel in linjer]))


def test_eksponering_regioner_summerer_og_saetter_ikke_opdelt_sidst():
    fonde = [_regionfond("A", 800.0, [("EUROPE", "DENMARK", 0.5), ("EUROPE", "EUROPE_OTHER", 0.25),
                                      ("NORTH_AMERICA", "US", 0.25)]),
             _regionfond("B", 200.0, [("EUROPE", "FRANCE", 0.5), ("NORTH_AMERICA", "USA_OTHER", 0.5)],
                         grundlag="GP_BREAKDOWN")]
    b = data.eksponering(fonde, "regioner")
    assert sum(g["dkk"] for g in b["grupper"]) == pytest.approx(1000.0)
    europa, usa = b["grupper"]
    assert europa["navn"] == "Europa" and europa["dkk"] == pytest.approx(700.0)
    assert [u["navn"] for u in europa["under"]] == ["Danmark", "Frankrig",
                                                    "Øvrige Europa og ikke opdelt"]
    assert [u["navn"] for u in usa["under"]] == ["USA", "Nordamerika, ikke opdelt"]
    assert europa["under"][1]["stoerste_fond"] == "B"
    assert sum(u["dkk"] for u in usa["under"]) == pytest.approx(usa["dkk"])


def test_delregioner_daekker_databasens_koder():
    assert set(data.DELREGIONER) >= data.IKKE_OPDELTE_REGIONER
    assert len(data.DELREGIONER) == 15
    assert set(data.REGIONER) == {"NORTH_AMERICA", "EUROPE", "ROW"}


def test_beholdninger_viser_gics_sektoren_som_industri(monkeypatch):
    h = _holdings([("UNREALISED", "FLATPAY", 10.0, 90.0), ("UNREALISED", "Landfolk", 5.0, 4.0),
                   ("UNREALISED", "Ukendt", 1.0, 1.0)],
                  industri=["FINANCIALS", "CONSUMER_DISCRETIONARY", None],
                  forslag=[True, False, None])
    monkeypatch.setattr(data, "read_sql", lambda *a, **k: h)
    beh = data._beholdninger("SEED4", date(2026, 6, 30), 0.5)
    r = {x["navn"]: x for x in beh["raekker"]}
    assert r["FLATPAY"]["industri"] == "Financials" and r["FLATPAY"]["industri_foreloebig"]
    assert r["Landfolk"]["industri"] == "Consumer Discretionary" and not r["Landfolk"]["industri_foreloebig"]
    assert r["Ukendt"]["industri"] is None and not r["Ukendt"]["industri_foreloebig"]


def _eksponeringsfond(navn, valuta, nav_dkk, raekker, type_="selskaber", skaleret=True):
    return dict(fund_name=navn, valuta=valuta, n=dict(nav_dkk=nav_dkk),
                beholdninger=dict(type=type_, skaleret=skaleret, dato=date(2026, 6, 30),
                                  raekker=[dict(navn=n, dagsvaerdi=v) for n, v in raekker]))


def test_stoerste_eksponering_sammenligner_i_dkk_og_udelader_ejendomme():
    fonde = [_eksponeringsfond("DKK-fond", "DKK", 500.0, [("A", 100.0), ("B", 50.0)]),
             _eksponeringsfond("EUR-fond", "EUR", 300.0, [("C", 20.0)], type_="fonde"),
             _eksponeringsfond("Ejendomme", "DKK", 200.0, [("Hus", 999.0)], type_="ejendomme"),
             _eksponeringsfond("Uskaleret", "DKK", 0.0, [("Hele fonden", 5000.0)], skaleret=False)]
    e = data.stoerste_eksponering(fonde, {"DKK": (1.0, None), "EUR": (7.5, None)})
    # C: 20 EUR = 150 DKK slår A's 100 DKK; ejendomme og uskalerede tal er ikke med
    assert e["navn"] == "C" and e["dkk"] == 150.0 and e["art"] == "underliggende fond"
    assert e["andel"] == pytest.approx(150.0 / 1000.0)
    assert e["udeladt"] == ["Ejendomme", "Uskaleret"]


def _positionsfonde():
    fonde = [_eksponeringsfond("DKK-fond", "DKK", 500.0, [("A", 100.0), ("C", 30.0)]),
             _eksponeringsfond("EUR-fond", "EUR", 300.0, [("c", 20.0)], type_="fonde"),
             _eksponeringsfond("Ejendomme", "DKK", 200.0, [("Hus", 999.0)], type_="ejendomme"),
             _eksponeringsfond("Uskaleret", "DKK", 40.0, [("Hele fonden", 5000.0)], skaleret=False)]
    for f in fonde:
        f["nav_dato"] = date(2026, 6, 30)
    return fonde


def test_stoerste_positioner_laegger_samme_navn_sammen_og_tager_fonde_uden_gennemlysning_med():
    sp = data.stoerste_positioner(_positionsfonde(), {"DKK": (1.0, None), "EUR": (7.5, None)})
    linjer = {p["navn"]: p for p in sp["linjer"]}
    # C i DKK-fonden (30) og c i EUR-fonden (20 EUR = 150 DKK) er én position
    assert linjer["C"]["dkk"] == pytest.approx(180.0)
    assert linjer["C"]["fonde"] == ["DKK-fond", "EUR-fond"]
    # ejendomsfonden og den uskalerede fond indgår med deres NAV, ikke deres beholdninger
    assert linjer["Ejendomme"]["art"] == "fond" and linjer["Ejendomme"]["dkk"] == 200.0
    assert linjer["Uskaleret"]["dkk"] == 40.0 and "Hele fonden" not in linjer
    assert sp["fondslinjer"] == ["Ejendomme", "Uskaleret"]
    assert [p["navn"] for p in sp["linjer"]] == ["Ejendomme", "C", "A", "Uskaleret"]
    assert sp["nav_dkk"] == 1040.0
    assert linjer["A"]["andel"] == pytest.approx(100.0 / 1040.0)
    assert sp["top_andel"] == pytest.approx(520.0 / 1040.0)


def test_stoerste_positioner_begraenser_antallet_og_kraever_dkk():
    kurser = {"DKK": (1.0, None), "EUR": (7.5, None)}
    sp = data.stoerste_positioner(_positionsfonde(), kurser, antal=2)
    assert [p["navn"] for p in sp["linjer"]] == ["Ejendomme", "C"]
    fonde = _positionsfonde()
    fonde[0]["n"]["nav_dkk"] = None
    assert data.stoerste_positioner(fonde, kurser) is None


def test_stoerste_eksponering_er_stoerste_gennemlyste_position():
    kurser = {"DKK": (1.0, None), "EUR": (7.5, None)}
    e = data.stoerste_eksponering(_positionsfonde(), kurser)
    # samme sammenlægning som på siden Største positioner; fondslinjerne tæller ikke
    assert e["navn"] == "C" and e["dkk"] == pytest.approx(180.0) and e["fond"] == "DKK-fond"


def test_sektorerne_har_engelske_gics_navne():
    assert data.SEKTORER["FINANCIALS"] == "Financials"
    assert data.SEKTORER["REAL_ESTATE"] == "Real Estate"
    assert data.BRANCHEGRUPPER[4510] == "Software & Services"
    assert data.SEKTORER["DIVERSIFIED"].startswith("Diversificeret")    # ikke GICS


def test_stoerste_eksponering_uden_beholdninger_er_none():
    fonde = [dict(fund_name="X", valuta="DKK", n=dict(nav_dkk=1.0), beholdninger=None)]
    assert data.stoerste_eksponering(fonde, {"DKK": (1.0, None)}) is None


def _afstemningsbeh(dagsvaerdi_fond, andel=0.01, type_="selskaber"):
    return dict(type=type_, skaleret=True, andel=andel, dato=date(2026, 6, 30),
                raekker=[dict(dagsvaerdi_fond=v, dagsvaerdi=v * andel) for v in dagsvaerdi_fond])


def test_nav_afstemning_deler_forskellen_i_carry_oevrige_og_ejerandel():
    # fonden: beholdninger 2.168, NAV før carry 2.208, NAV 1.914 (mio.); investor 1 % + lidt
    beh = _afstemningsbeh([1_500e6, 668e6])
    a = data.nav_afstemning(19.3e6, date(2026, 6, 30), beh, 1_914e6, 2_208e6)
    assert a["forskel"] == pytest.approx(19.3e6 - 21.68e6)
    assert a["carry"] == pytest.approx(0.01 * (1_914e6 - 2_208e6))
    assert a["oevrige"] == pytest.approx(0.01 * (2_208e6 - 2_168e6))
    assert a["ejerandel"] == pytest.approx(19.3e6 - 0.01 * 1_914e6)
    assert a["carry"] + a["oevrige"] + a["ejerandel"] == pytest.approx(a["forskel"])
    assert a["nav_andel"] == pytest.approx(19.3e6 / 1_914e6)


def test_nav_afstemning_uden_fondens_nav_giver_kun_forskellen():
    a = data.nav_afstemning(9e6, date(2026, 6, 30), _afstemningsbeh([1_000e6]))
    assert a["forskel"] == pytest.approx(-1e6)
    assert a["carry"] is None and a["oevrige"] is None and a["ejerandel"] is None


def test_nav_afstemning_er_none_uden_synlig_forskel_feeder_eller_ejendomme():
    beh = _afstemningsbeh([1_000e6])
    assert data.nav_afstemning(10e6 + 400, date(2026, 6, 30), beh) is None
    assert data.nav_afstemning(1.0, date(2026, 6, 30), beh | dict(via="X")) is None
    assert data.nav_afstemning(1.0, date(2026, 6, 30), beh | dict(type="ejendomme")) is None
    assert data.nav_afstemning(1.0, date(2026, 6, 30), beh | dict(skaleret=False)) is None


def test_benchmark_har_gics_sektorer_der_summerer_til_hundrede():
    bm = data.benchmark()
    assert set(bm["indeks"]) == {"MSCI ACWI"}
    for vaegte in bm["indeks"].values():
        assert set(vaegte) <= set(data.SEKTORER) - data.UDEN_GICS
        assert sum(vaegte.values()) == pytest.approx(1.0, abs=0.001)
    assert bm["indeks"]["MSCI ACWI"]["INFORMATION_TECHNOLOGY"] == pytest.approx(0.3321)


def test_gics_struktur_har_alle_sektorer_og_hver_branchegruppe_en_gang():
    assert set(data.GICS_STRUKTUR) == set(data.SEKTORER) - data.UDEN_GICS
    grupper = [g for gs in data.GICS_STRUKTUR.values() for g in gs]
    assert sorted(grupper) == sorted(data.BRANCHEGRUPPER) and len(grupper) == 25
    assert data.GICS_STRUKTUR["FINANCIALS"] == [4010, 4020, 4030]


def test_benchmark_uden_fil_er_none(tmp_path):
    assert data.benchmark(tmp_path / "findes_ikke.yaml") is None


def test_benchmark_regioner_har_regionskoderne_og_summerer_til_hundrede():
    bm = data.benchmark()
    assert set(bm["regioner"]) == set(bm["indeks"])
    assert bm["regioner"]["MSCI ACWI"]["NORTH_AMERICA"] == pytest.approx(0.6664)
    for vaegte in bm["regioner"].values():
        assert set(vaegte) <= set(data.REGIONER)
        assert sum(vaegte.values()) == pytest.approx(1.0, abs=0.001)


def test_benchmark_laeser_regionvaegte(tmp_path):
    sti = tmp_path / "bm.yaml"
    sti.write_text("regioner:\n  MSCI ACWI:\n    NORTH_AMERICA: 67.11\n    EUROPE: 20\n"
                   "    ROW: 12.89\n", encoding="utf-8")
    bm = data.benchmark(sti)
    assert bm["indeks"] == {}
    assert bm["regioner"]["MSCI ACWI"]["NORTH_AMERICA"] == pytest.approx(0.6711)
