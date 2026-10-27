"""Stier til de filer i repoet, som pakken læser og skriver.

Pakken køres fra et checkout (eller en editable installation), så repo-roden er mappen
over pakken.
"""
import os
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FUND_INFO = ROOT / "config" / "fund_info.yaml"       # manuelle stamdata pr. fond
BENCHMARK = ROOT / "config" / "benchmark.yaml"       # MSCI-indeksenes sektorvægte
PROGNOSE = ROOT / "config" / "cashflow_profiler.yaml"  # profilerne til cash flow-prognosen
ALETA = ROOT / "data" / "aleta"                       # Aletas udtræk pr. investor
INDEKS = ROOT / "data" / "indeks"                     # benchmarkindeksets kurser fra yfinance
MASTER = ROOT / "assets" / "ppim_slidemaster.pptx"   # PPIM-slidemasteren
EXPORTS = ROOT / "exports"                           # færdige rapporter (ikke i git)
LOG = ROOT / "log"                                   # byggeloggen, én fil pr. bygning (ikke i git)
# kildedokumenterne; fund.document.rel_path er relativ til denne mappe
DOKUMENTER = Path(os.environ.get("PE_DOCUMENTS",
                                 ROOT.parent / "pe-external-fund-data" / "documents"))
