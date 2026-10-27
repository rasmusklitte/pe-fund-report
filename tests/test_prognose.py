"""Cash flow-prognosen (prognose.py) og koblingen til rapporten (data.prognose_paa_fondene)."""
from datetime import date

import pytest
from test_data import _lq_fond

from pe_fund_report import data, prognose

T = date(2026, 6, 30)
SLUT = date(2036, 12, 31)
PE = dict(kaldsrater=[0.25, 0.33, 0.40], kald_til_aar=8, vaekst=0.12, bow=2.5, udbytte=0.0,
          loebetid_aar=12)
PROFILER = {"Private equity": PE, "standard": PE | dict(vaekst=0.10),
            "Ejendomme": dict(kaldsrater=[0.5], vaekst=0.06, bow=2.5, udbytte=0.04,
                              loebetid_aar=15)}


def test_profil_efter_segment_med_standard_og_fondens_afvigelser():
    assert prognose.profil("Private equity", PROFILER) == ("Private equity", PE)
    assert prognose.profil("Shipping", PROFILER)[0] == "standard"
    navn, p = prognose.profil("Private equity", PROFILER, {"profil": "Ejendomme", "bow": 1.0})
    assert (navn, p["bow"], p["udbytte"]) == ("Ejendomme", 1.0, 0.04)
    # en stavefejl bliver ikke stille til standardprofilen
    with pytest.raises(ValueError, match="ukendte felter"):
        prognose.profil("Private equity", PROFILER, {"boow": 1.0})
    with pytest.raises(ValueError, match="ukendt prognoseprofil"):
        prognose.profil("Private equity", PROFILER, {"profil": "Venture"})


def test_scenarie_laegger_vaekst_til_og_ganger_resten():
    s = prognose.scenarie(PE, dict(vaekst=-0.06, bow=1.3, kaldsrater=3.0))
    assert s["vaekst"] == pytest.approx(0.06) and s["bow"] == pytest.approx(3.25)
    assert s["kaldsrater"] == [0.75, 0.99, 1.0]          # højst 1
    assert prognose.scenarie(PE, None) == PE


def test_kvartalsslut():
    assert prognose.kvartalsslut(T, date(2027, 3, 31)) == [
        date(2026, 9, 30), date(2026, 12, 31), date(2027, 3, 31)]
    # midt i et kvartal: først det kvartals ultimo
    assert prognose.kvartalsslut(date(2026, 5, 15), date(2026, 9, 30)) == [
        date(2026, 6, 30), date(2026, 9, 30)]


def test_kaeden_gaar_op_hvert_kvartal():
    kv = prognose.fremskriv(T, SLUT, 40.0, 70.0, date(2024, 3, 1), 12, PE)
    assert len(kv) == 42
    for q, naeste in zip(kv, kv[1:]):
        assert naeste["rest_primo"] == pytest.approx(q["resttilsagn"])
        assert naeste["nav_primo"] == pytest.approx(q["nav"])
    for q in kv:
        assert q["nav"] == pytest.approx(q["nav_primo"] * q["vaekstfaktor"] + q["kald"]
                                         - q["udlodning"])
        assert q["resttilsagn"] == pytest.approx(q["rest_primo"] - q["kald"])
        assert q["kald"] >= 0 and q["udlodning"] >= 0
    assert sum(q["kald"] for q in kv) <= 40.0


def test_kald_stopper_ved_kald_til_aar_og_alt_udloddes_ved_loebetidens_udgang():
    kv = prognose.fremskriv(T, SLUT, 40.0, 70.0, date(2024, 3, 1), 12, PE)
    efter = [q for q in kv if q["alder"] >= 8]
    assert efter and all(q["kald"] == 0 for q in efter)
    assert kv[-1]["resttilsagn"] > 0                    # resten kaldes aldrig
    udloeb = next(q for q in kv if q["dato"] >= date(2036, 3, 1))
    assert udloeb["nav"] == pytest.approx(0.0, abs=1e-9)
    assert all(q["nav"] == pytest.approx(0.0, abs=1e-9) for q in kv if q["dato"] > udloeb["dato"])


def test_aarets_kald_er_jaevnt_fordelt_over_aaret():
    p = PE | dict(kaldsrater=[1.0], kald_til_aar=None)
    kv = prognose.fremskriv(T, date(2027, 6, 30), 40.0, 0.0, date(2026, 6, 30), 12, p)
    # hele resttilsagnet på et år, ikke i første kvartal
    assert [round(q["kald"], 1) for q in kv] == [10.1, 10.1, 9.9, 10.0]
    # et år på 365 dage er 0,9993 år på 365,25 dage; resten kaldes året efter
    assert kv[-1]["resttilsagn"] == pytest.approx(0.0, abs=0.05)


def test_fond_efter_loebetiden_udlodder_alt_med_det_samme():
    kv = prognose.fremskriv(T, SLUT, 5.0, 70.0, date(2010, 1, 1), 12, PE)
    assert kv[0]["udlodning"] == pytest.approx(70.0 * kv[0]["vaekstfaktor"])
    assert sum(q["kald"] for q in kv) == 0


def test_evergreen_afvikles_aldrig():
    p = PROFILER["Ejendomme"]
    kv = prognose.fremskriv(T, SLUT, 0.0, 100.0, date(2005, 1, 1), None, p)
    assert kv[-1]["nav"] > 100.0
    assert kv[0]["udlodningsrate"] == pytest.approx(1 - 0.96 ** kv[0]["aar"])


def test_svagt_marked_udlodder_mindre_end_basis_og_staerkt_mere():
    def udloddet(aendring, til=SLUT):
        kv = prognose.fremskriv(T, til, 40.0, 70.0, date(2024, 3, 1), 12,
                                prognose.scenarie(PE, aendring))
        return sum(q["udlodning"] for q in kv)
    svag, basis, staerk = (dict(vaekst=-0.06, bow=1.3), None, dict(vaekst=0.04, bow=0.8))
    for til in (date(2029, 6, 30), SLUT):
        assert udloddet(svag, til) < udloddet(basis, til) < udloddet(staerk, til)


def test_pr_aar_og_summer():
    kv = prognose.fremskriv(T, date(2027, 12, 31), 40.0, 70.0, date(2024, 3, 1), 12, PE)
    aar = prognose.pr_aar(kv, kurs=2.0)
    assert [a["aar"] for a in aar] == [2026, 2027]
    assert aar[0]["kald"] == pytest.approx(2 * sum(q["kald"] for q in kv[:2]))
    assert aar[1]["nav"] == pytest.approx(2 * kv[-1]["nav"])
    assert aar[0]["netto"] == pytest.approx(aar[0]["udlodning"] - aar[0]["kald"])
    to = prognose.summer([aar, aar])
    assert to[1]["resttilsagn"] == pytest.approx(2 * aar[1]["resttilsagn"])


OPSAETNING = dict(horisont_aar=10, profiler=PROFILER,
                  scenarier=dict(basis=dict(navn="Basis"),
                                 svag=dict(navn="Svagt marked", vaekst=-0.06, bow=1.3)))
KURSER = {"EUR": (7.46, T)}


def _fonde():
    alfa = _lq_fond("Alfa", "Private equity") | dict(loebetid_aar=10.0, evergreen=False)
    ejendom = _lq_fond("Ejendom", "Ejendomme") | dict(loebetid_aar=None, evergreen=True)
    return [alfa, ejendom]


def test_prognose_paa_fondene():
    fonde = _fonde()
    p = data.prognose_paa_fondene(fonde, KURSER, T, OPSAETNING)
    assert p["scenarier"] == [("basis", "Basis"), ("svag", "Svagt marked")]
    assert p["til"] == SLUT and p["uden_kurs"] == []
    alfa, ejendom = fonde
    assert alfa["prognose"]["profil"] == "Private equity"
    # alderen regnes fra første kald, løbetiden fra databasen
    assert alfa["prognose"]["start"] == date(2024, 3, 1)
    assert alfa["prognose"]["loebetid"] == 10.0 and ejendom["prognose"]["loebetid"] is None
    # porteføljen er fondenes kvartaler i DKK lagt sammen år for år
    basis = p["aar"]["basis"]
    kald = sum(q["kald"] for f in fonde for q in f["prognose"]["scenarier"]["basis"]["kvartaler"])
    assert sum(a["kald"] for a in basis) == pytest.approx(kald * 7.46)
    assert p["naeste"]["basis"]["netto_3aar"] == pytest.approx(
        p["naeste"]["basis"]["udlodning_3aar"] - p["naeste"]["basis"]["kald_3aar"])
    assert p["fonde"][0]["resttilsagn"] == pytest.approx(40.0 * 7.46)
    # TVPI: (udloddet + kommende udlodninger + slut-NAV) / (indbetalt + kommende kald)
    x = alfa["prognose"]["scenarier"]["basis"]
    kv = x["kvartaler"]
    assert x["tvpi"] == pytest.approx((10 + sum(q["udlodning"] for q in kv) + kv[-1]["nav"])
                                      / (60 + sum(q["kald"] for q in kv)))
    assert p["tvpi"]["svag"] < p["tvpi"]["basis"]
    assert list(p["profiler"]) == ["Private equity", "Ejendomme"]


def test_prognose_med_fondens_afvigelser_og_uden_kurs():
    fonde = _fonde()
    fonde[0]["prognose_afvigelser"] = {"bow": 1.0}
    p = data.prognose_paa_fondene(fonde, {"EUR": None}, T, OPSAETNING)
    assert p is None                                    # ingen fond med DKK-kurs
    p = data.prognose_paa_fondene(fonde, KURSER, T, OPSAETNING)
    assert fonde[0]["prognose"]["parametre"]["bow"] == 1.0
    assert p["fonde"][0]["afvigelser"] == ["bow"]
    assert data.prognose_paa_fondene(fonde, KURSER, T, None) is None


def test_profilfilen_har_en_profil_pr_segment():
    o = data.prognoseprofiler()
    for segment in [*data.STRATEGI_ORDEN, prognose.STANDARD]:
        p = o["profiler"][segment]
        assert {"kaldsrater", "vaekst", "bow", "udbytte", "loebetid_aar"} <= set(p), segment
        assert set(p) <= set(prognose.FELTER), segment
        assert all(0 < r <= 1 for r in p["kaldsrater"]), segment
    # basis står først og ændrer intet
    noegle, basis = next(iter(o["scenarier"].items()))
    assert set(basis) == {"navn"}, noegle


def test_fund_info_prognose_er_gyldig():
    o = data.prognoseprofiler()
    for info in data.fund_info().values():
        if info.get("prognose"):
            prognose.profil(None, o["profiler"], info["prognose"])     # fejler ved stavefejl
