"""Stier til de filer i repoet, som pakken læser og skriver.

Pakken køres fra et checkout (eller en editable installation), så repo-roden er mappen
over pakken.
"""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FUND_INFO = ROOT / "config" / "fund_info.yaml"       # manuelle stamdata pr. fond
MASTER = ROOT / "assets" / "ppim_slidemaster.pptx"   # PPIM-slidemasteren
EXPORTS = ROOT / "exports"                           # færdige rapporter (ikke i git)
