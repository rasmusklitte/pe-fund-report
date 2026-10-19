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


def test_fund_info_har_de_forventede_felter():
    for kode, info in data.fund_info().items():
        assert {"status", "beskrivelse"} <= set(info), kode
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


def test_segment_bruger_struktur_foer_fokus():
    assert data.segment("FUND_OF_FUNDS", "BUYOUT", None) == "Fund of funds"
    assert data.segment("SECONDARY", "BUYOUT", "SMALL_MID") == "Secondary fund of funds"
    assert data.segment("DIRECT", "BUYOUT", "SMALL_MID") == "Small/mid-cap buyout"
    assert data.segment("DIRECT", "BUYOUT", "ALL") == "Buyout"
    assert data.segment("DIRECT", "VENTURE", float("nan")) == "Venture"
    assert data.segment("DIRECT", "JUNIOR_DEBT", "MID_LARGE") == "Private credit"
    assert data.segment(None, None, None) is None


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
