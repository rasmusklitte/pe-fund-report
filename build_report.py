"""Bygger investorrapporten som pptx på PPIM-masteren.

    python build_report.py MOTORTRAMP                 # seneste NAV-dato
    python build_report.py MOTORTRAMP --as-of 2025-12-31

Resultat: exports/<INVESTOR>_<dato>.pptx med forside, porteføljeoversigt, én side pr. fond og slut-slide.
"""
import argparse
from pathlib import Path

from pptx import Presentation

import data
import slides
from deck import L_FORSIDE, L_SKILLE, L_SLUT, _find_master, _ny_slide, _ryd_slides, _saet_ph


def byg(investor_code, as_of=None, outfile=None, master=None):
    r = data.investor_rapport(investor_code, as_of)
    prs = Presentation(str(_find_master(master)))
    _ryd_slides(prs)

    s = _ny_slide(prs, L_FORSIDE)
    _saet_ph(s, 0, f"{r['navn']} – private equity")
    _saet_ph(s, 1, f"Fondsrapport pr. {r['as_of']:%d.%m.%Y}")

    s = _ny_slide(prs, L_SKILLE[0])
    _saet_ph(s, 0, "Porteføljeoversigt")
    slides.portefoeljeoversigt(prs, r)
    slides.fordeling(prs, r, "geografi", "Geografisk fordeling efter indbetalt kapital")
    slides.fordeling(prs, r, "segment", "Segmentfordeling efter indbetalt kapital")

    s = _ny_slide(prs, L_SKILLE[1])
    _saet_ph(s, 0, "Fondsoversigter")
    for f in r["fonde"]:
        slides.fondsside(prs, f, r["navn"])

    _ny_slide(prs, L_SLUT)

    outfile = Path(outfile or f"exports/{r['investor']}_{r['as_of']:%Y-%m-%d}.pptx")
    outfile.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(outfile))
    print(f"gemt: {outfile}  ({len(r['fonde'])} fonde, {len(prs.slides._sldIdLst)} slides)")
    return outfile


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("investor", help="investor_code i fund.investor, fx MOTORTRAMP")
    ap.add_argument("--as-of", help="rapportdato (ÅÅÅÅ-MM-DD); standard er seneste NAV-dato")
    ap.add_argument("--out", help="sti til pptx-filen")
    a = ap.parse_args()
    byg(a.investor, a.as_of, a.out)
