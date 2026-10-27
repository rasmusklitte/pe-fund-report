"""Benchmarkindekset til PME: kurserne fra yfinance, gemt som CSV i data/indeks.

yfinance kaldes kun, når et kvartalsskifte er passeret siden sidste hentning: filen slutter
altid ved et kvartalsultimo, så den er ajour, indtil det næste kvartal er omme. Uden net
bruges filen, som den er.
"""
from datetime import date, timedelta

import pandas as pd

from .paths import INDEKS

BENCHMARK_TICKER = "ACWI"                       # iShares MSCI ACWI ETF
BENCHMARK_NAVN = "MSCI ACWI (iShares ETF)"
BENCHMARK_VALUTA = "USD"
HELLIGDAGE = 5      # dage, kvartalets sidste handelsdag kan ligge før ultimo (weekend, helligdag)
START = date(2000, 1, 1)


def seneste_kvartalsslut(i_dag):
    """Det seneste kvartalsultimo før i_dag (på selve ultimo er kvartalets lukkekurs der endnu
    ikke)."""
    m = (i_dag.month - 1) // 3 * 3
    if m == 0:
        return date(i_dag.year - 1, 12, 31)
    return date(i_dag.year, m + 1, 1) - timedelta(days=1)


def _hent_yfinance(ticker, til):
    """Hele kurshistorikken til og med til, justeret for udbytte (auto_adjust), så kursen er
    et total return-indeks. pd.Series dato -> kurs."""
    import yfinance as yf      # kun når der faktisk hentes
    df = yf.Ticker(ticker).history(start=START.isoformat(),
                                   end=(til + timedelta(days=1)).isoformat(), auto_adjust=True)
    if df.empty:
        raise ValueError(f"yfinance gav ingen kurser for {ticker}")
    return pd.Series(df["Close"].to_numpy(dtype=float),
                     index=[pd.Timestamp(d).date() for d in df.index])


def _laes(sti):
    df = pd.read_csv(sti, parse_dates=["dato"])
    return pd.Series(df["kurs"].to_numpy(dtype=float), index=[d.date() for d in df["dato"]])


def _skriv(sti, serie):
    sti.parent.mkdir(parents=True, exist_ok=True)
    pd.DataFrame({"dato": [d.isoformat() for d in serie.index], "kurs": serie.to_numpy()}) \
        .to_csv(sti, index=False, float_format="%.6f")


def indeksserie(ticker=BENCHMARK_TICKER, i_dag=None, hent=_hent_yfinance, mappe=INDEKS):
    """Indeksets kurser som pd.Series (dato -> kurs, stigende), eller None.

    Fra data/indeks/<ticker>.csv, når filen når frem til det seneste kvartalsultimo; ellers
    hentes historikken til og med det ultimo og gemmes. Mislykkes hentningen, bruges filen,
    hvis den findes (med en advarsel i byggeloggen); ellers None.
    """
    sti = mappe / f"{ticker}.csv"
    ultimo = seneste_kvartalsslut(i_dag or date.today())
    gemt = _laes(sti) if sti.exists() else None
    if gemt is not None and not gemt.empty \
            and gemt.index[-1] >= ultimo - timedelta(days=HELLIGDAGE):
        return gemt
    try:
        serie = hent(ticker, ultimo)
    except Exception as e:      # net, yfinance's API: rapporten bygges stadig
        print(f"ADVARSEL: kunne ikke hente {ticker} fra yfinance ({e})"
              + (f"; bruger kurserne til {gemt.index[-1]:%d.%m.%Y}" if gemt is not None
                 else "; PME vises ikke"))
        return gemt
    serie = serie[[d <= ultimo for d in serie.index]].sort_index()
    serie = serie[~serie.index.duplicated(keep="last")]
    _skriv(sti, serie)
    print(f"hentet: {ticker} fra yfinance til {serie.index[-1]:%d.%m.%Y} ({len(serie)} kurser)")
    return serie
