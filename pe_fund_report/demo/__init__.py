"""Demorapporten: en fiktiv investor (DEMO), bygget uden databasen.

demo_kontekst() erstatter data.read_sql med en SQLite-database i hukommelsen, som har
databasens skemaer og tabeller (rpt-viewene som almindelige tabeller) fyldt med syntetiske
data fra generator.py. Så kører data.py's egne forespørgsler uændret, og alt fra fond() til
slides og Excel regnes, som det gør på de rigtige data. En ny forespørgsel mod en ny tabel eller
kolonne skal også have den i generator.py, ellers fejler demoen (og tests/test_demo.py).

SQLite forstår ikke al T-SQL; _til_sqlite() oversætter det, data.py bruger: SELECT TOP 1
bliver LIMIT 1 og ISNULL(a, b) IFNULL(a, b). OUTER APPLY (forvalterens fordeling,
GP_BREAKDOWN) bruges ikke af demofondene og afvises med en tydelig fejl.
"""
import re
import sqlite3
from contextlib import contextmanager
from datetime import date, datetime, timedelta

import pandas as pd

from .. import data, db, indeks
from . import generator
from .generator import INVESTOR  # noqa: F401  (report.byg bruger den)

SKEMAER = ("fund", "rpt", "dbo", "portfolio")
SERVER = "demo (syntetiske data)"
_DATO = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def _til_sqlite(sql):
    """T-SQL -> SQLite for de konstruktioner, data.py bruger."""
    if re.search(r"\bOUTER\s+APPLY\b", sql, re.IGNORECASE):
        raise NotImplementedError("demodatabasen kan ikke OUTER APPLY (GP_BREAKDOWN); giv "
                                  "demofonden HOLDINGS eller FUND_DEFAULT i generator.py")
    sql = re.sub(r"\bISNULL\s*\(", "IFNULL(", sql, flags=re.IGNORECASE)
    if re.search(r"\bSELECT\s+TOP\s+1\b", sql, re.IGNORECASE):
        sql = re.sub(r"\bSELECT\s+TOP\s+1\b", "SELECT", sql, count=1, flags=re.IGNORECASE)
        sql = sql.rstrip().rstrip(";") + " LIMIT 1"
    return sql


def _parameter(v):
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    return v


def _datoer_tilbage(df):
    """Datoer kommer ud af SQLite som tekst; giv dem tilbage som datetime.date, som pyodbc gør."""
    for k in df.columns:
        if df[k].dtype == object or pd.api.types.is_string_dtype(df[k]):
            vaerdier = df[k].dropna()
            if not vaerdier.empty and all(isinstance(v, str) and _DATO.match(v)
                                          for v in vaerdier):
                df[k] = pd.Series([None if v is None or pd.isna(v) else date.fromisoformat(v)
                                   for v in df[k]], index=df.index, dtype=object)
    return df


def _mangler(v):
    return v is None or (isinstance(v, float) and v != v)


def _python(v):
    """numpy-tal som Python-tal, så sqlite3 kan gemme dem."""
    return v.item() if hasattr(v, "item") else v


def _til_tabel(df):
    ud = df.copy()
    for k in ud.columns:
        if ud[k].dtype == object:
            ud[k] = [_parameter(v) for v in ud[k]]
    return ud


def forbindelse():
    """En SQLite-forbindelse i hukommelsen med demodataene i databasens skemaer."""
    tabeller, _ = generator.demodata()
    con = sqlite3.connect(":memory:", check_same_thread=False)
    for s in SKEMAER:
        con.execute(f"ATTACH DATABASE ':memory:' AS {s}")
    for navn, df in tabeller.items():
        # pandas' to_sql kender ikke SQLite's vedhæftede skemaer, så tabellen laves her
        df = _til_tabel(df)
        con.execute(f"CREATE TABLE {navn} ({', '.join(df.columns)})")
        con.executemany(f"INSERT INTO {navn} VALUES ({', '.join('?' * len(df.columns))})",
                        [tuple(None if _mangler(v) else _python(v) for v in r)
                         for r in df.itertuples(index=False)])
    # tabeller, data.py læser, men demoen ikke fylder: tomme, med kolonnerne
    for navn, kolonner in TOMME.items():
        skema, tabel = navn.split(".")
        if navn not in tabeller:
            con.execute(f"CREATE TABLE {skema}.{tabel} ({', '.join(kolonner)})")
    return con


TOMME = {
    "rpt.v_fund_logo": ["fund_code", "background", "width_px", "height_px", "image"],
    "rpt.v_portfolio_breakdown": ["fund_code", "report_date", "dimension", "category",
                                  "share_pct", "amount", "document_id", "industry_code",
                                  "industry_group_gics", "region_code", "subregion_code"],
    "portfolio.industry_mapping": ["fund_id", "label", "industry_code", "industry_group_gics",
                                   "is_suggested"],
    "portfolio.geo_mapping": ["fund_id", "label", "subregion_code", "is_suggested"],
}


class Demodatabase:
    """read_sql mod demodatabasen; husker forespørgslerne (til tests)."""

    def __init__(self):
        self.con = forbindelse()
        self.forespoergsler = []

    def read_sql(self, sql, **params):
        if sql.lstrip().split(None, 1)[0].upper() not in ("SELECT", "WITH"):
            raise ValueError("Kun SELECT/WITH er tilladt mod databasen")
        self.forespoergsler.append(sql)
        df = pd.read_sql(_til_sqlite(sql), self.con,
                         params={k: _parameter(v) for k, v in params.items()})
        return _datoer_tilbage(df)


def _uden_hentning(ticker, til):
    raise RuntimeError("demoen henter ikke kurser")


@contextmanager
def demo_kontekst():
    """Kør bygningen mod demodataene: database, fund_info og indekskurserne (fra data/indeks,
    aldrig fra yfinance). Giver Demodatabasen."""
    demodb = Demodatabase()
    _, info = generator.demodata()
    gamle = (data.read_sql, data.fund_info, indeks.indeksserie, db.SERVER, db.DATABASE)
    oprindelig_serie = indeks.indeksserie

    def indeksserie(ticker=indeks.BENCHMARK_TICKER, i_dag=None, hent=None, mappe=indeks.INDEKS):
        # som om det var dagen efter demodatoen: kurserne i data/indeks går dertil
        return oprindelig_serie(ticker, i_dag=generator.DATO + timedelta(days=1),
                                hent=_uden_hentning, mappe=mappe)

    data.read_sql = demodb.read_sql
    data.fund_info = lambda: info
    indeks.indeksserie = indeksserie
    db.SERVER, db.DATABASE = SERVER, "DEMO"
    try:
        yield demodb
    finally:
        (data.read_sql, data.fund_info, indeks.indeksserie, db.SERVER, db.DATABASE) = gamle


def demorapport():
    """Hele rapport-dicten for demoinvestoren (data.investor_rapport mod demodataene)."""
    with demo_kontekst():
        return data.investor_rapport(generator.INVESTOR)
