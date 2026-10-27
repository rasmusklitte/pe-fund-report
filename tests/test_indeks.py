"""Benchmarkindeksets cache (indeks.py): yfinance kaldes kun efter et kvartalsskifte."""
from datetime import date, timedelta

import pandas as pd

from pe_fund_report import indeks


def _hent(kaldt, slut=None):
    """En falsk yfinance: daglige kurser fra 2024 til et stykke efter til."""
    def hent(ticker, til):
        kaldt.append(til)
        datoer = pd.date_range("2024-01-01", slut or til + timedelta(days=20), freq="B")
        return pd.Series(range(1, len(datoer) + 1), index=[d.date() for d in datoer], dtype=float)
    return hent


def _fejl(ticker, til):
    raise ConnectionError("intet net")


def test_seneste_kvartalsslut():
    assert indeks.seneste_kvartalsslut(date(2026, 10, 27)) == date(2026, 9, 30)
    assert indeks.seneste_kvartalsslut(date(2026, 9, 30)) == date(2026, 6, 30)
    assert indeks.seneste_kvartalsslut(date(2026, 1, 5)) == date(2025, 12, 31)
    assert indeks.seneste_kvartalsslut(date(2026, 4, 1)) == date(2026, 3, 31)


def test_uden_csv_hentes_til_kvartalsultimo_og_gemmes(tmp_path):
    kaldt = []
    s = indeks.indeksserie("ACWI", date(2026, 10, 27), _hent(kaldt), tmp_path)
    assert kaldt == [date(2026, 9, 30)]
    assert s.index[-1] == date(2026, 9, 30)               # intet efter ultimo
    assert (tmp_path / "ACWI.csv").exists()
    gemt = indeks.indeksserie("ACWI", date(2026, 10, 27), _fejl, tmp_path)
    assert gemt.index[-1] == date(2026, 9, 30) and gemt.iloc[-1] == s.iloc[-1]


def test_csv_til_kvartalsultimo_henter_ikke_igen(tmp_path):
    kaldt = []
    indeks.indeksserie("ACWI", date(2026, 10, 1), _hent(kaldt), tmp_path)
    # resten af kvartalet og på selve ultimo: ingen ny hentning
    for d in (date(2026, 11, 15), date(2026, 12, 31)):
        indeks.indeksserie("ACWI", d, _hent(kaldt), tmp_path)
    assert kaldt == [date(2026, 9, 30)]


def test_ultimo_i_weekenden_henter_ikke_igen(tmp_path):
    # 28.06.2026 er en søndag: sidste kurs er fredag den 26.
    kaldt = []
    indeks.indeksserie("ACWI", date(2026, 7, 1), _hent(kaldt, slut=date(2026, 6, 26)), tmp_path)
    indeks.indeksserie("ACWI", date(2026, 7, 20), _hent(kaldt), tmp_path)
    assert len(kaldt) == 1


def test_csv_et_kvartal_gammel_hentes_igen(tmp_path):
    kaldt = []
    indeks.indeksserie("ACWI", date(2026, 8, 1), _hent(kaldt), tmp_path)
    s = indeks.indeksserie("ACWI", date(2026, 10, 2), _hent(kaldt), tmp_path)
    assert kaldt == [date(2026, 6, 30), date(2026, 9, 30)]
    assert s.index[-1] == date(2026, 9, 30)


def test_fejl_ved_hentning_bruger_csv_eller_giver_none(tmp_path, capsys):
    assert indeks.indeksserie("ACWI", date(2026, 8, 1), _fejl, tmp_path) is None
    indeks.indeksserie("ACWI", date(2026, 8, 1), _hent([]), tmp_path)
    s = indeks.indeksserie("ACWI", date(2026, 10, 2), _fejl, tmp_path)
    assert s.index[-1] == date(2026, 6, 30)
    assert "ADVARSEL" in capsys.readouterr().out
