"""Samler investorrapporten: forside, porteføljeoversigt, én side pr. fond og slut-slide."""
from pathlib import Path

from pptx import Presentation

from . import data, slides
from .deck import L_FORSIDE, L_SKILLE, L_SLUT, _find_master, _ny_slide, _ryd_slides, _saet_ph
from .paths import EXPORTS, MASTER


def byg(investor_code, as_of=None, outfile=None, master=None):
    """Byg rapporten for én investor og returnér stien til pptx-filen.

    as_of: rapportdato (ÅÅÅÅ-MM-DD); standard er investors seneste NAV-dato.
    outfile: standard er exports/<INVESTOR>_<dato>.pptx.
    """
    r = data.investor_rapport(investor_code, as_of)
    prs = Presentation(str(_find_master(master or MASTER)))
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

    outfile = Path(outfile or EXPORTS / f"{r['investor']}_{r['as_of']:%Y-%m-%d}.pptx")
    outfile.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(outfile))
    print(f"gemt: {outfile}  ({len(r['fonde'])} fonde, {len(prs.slides._sldIdLst)} slides)")
    return outfile
