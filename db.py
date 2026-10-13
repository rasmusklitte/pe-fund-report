"""Læseadgang til PE-databasen (se analysis-data-guide.md, afsnit 1).

Standard er sql1\\ppim / PE med Windows-godkendelse. Kan overstyres med miljøvariablerne
PE_DB_SERVER, PE_DB_DATABASE og PE_DB_DRIVER. Repoet skriver aldrig til databasen -
read_sql afviser alt andet end SELECT/WITH.
"""
import os
from functools import lru_cache

import pandas as pd
import sqlalchemy as sa
from sqlalchemy.engine import URL

SERVER   = os.environ.get("PE_DB_SERVER", r"sql1\ppim")
DATABASE = os.environ.get("PE_DB_DATABASE", "PE")
DRIVER   = os.environ.get("PE_DB_DRIVER", "ODBC Driver 18 for SQL Server")


@lru_cache(maxsize=1)
def engine():
    # Driver 18 krypterer som standard, og det fejler mod en on-prem-server med
    # selvsigneret certifikat - derfor Encrypt=no.
    odbc = (f"DRIVER={{{DRIVER}}};SERVER={SERVER};DATABASE={DATABASE};"
            "Trusted_Connection=yes;Encrypt=no;TrustServerCertificate=yes;")
    return sa.create_engine(URL.create("mssql+pyodbc", query={"odbc_connect": odbc}))


def read_sql(sql, **params):
    """Kør en SELECT og returnér en DataFrame. Parametre bindes som :navn."""
    if sql.lstrip().split(None, 1)[0].upper() not in ("SELECT", "WITH"):
        raise ValueError("pe_analysis læser kun fra databasen - kun SELECT/WITH er tilladt")
    with engine().connect() as conn:
        return pd.read_sql(sa.text(sql), conn, params=params)
