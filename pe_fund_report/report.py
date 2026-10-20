"""Samler investorrapporten: forside, porteføljeoversigt, én side pr. fond og slut-slide."""
from pathlib import Path

from pptx import Presentation

from . import data, excel, slides
from .deck import L_FORSIDE, L_SKILLE, L_SLUT, _find_master, _ny_slide, _ryd_slides, _saet_ph
from .paths import EXPORTS, MASTER


def byg(investor_code, as_of=None, outfile=None, master=None, med_data=True):
    """Byg rapporten for én investor og returnér stien til pptx-filen.

    as_of: rapportdato (ÅÅÅÅ-MM-DD); standard er investors seneste NAV-dato.
    outfile: standard er exports/<INVESTOR>_<dato>.pptx.
    med_data: gem også dataudtrækket <outfile>_data.xlsx med kilder til hvert tal.
    """
    r = data.investor_rapport(investor_code, as_of)
    prs = Presentation(str(_find_master(master or MASTER)))
    _ryd_slides(prs)

    s = _ny_slide(prs, L_FORSIDE)
    # porteføljen rummer også private credit og ejendomme, så ikke "private equity"
    _saet_ph(s, 0, f"{r['navn']} – alternative investeringer")
    _saet_ph(s, 1, f"Fondsrapport pr. {r['as_of']:%d.%m.%Y}")
    slides.sammendrag(prs, r)
    slides.udvikling(prs, r)            # springes over uden DKK-kurser for alle fonde

    s = _ny_slide(prs, L_SKILLE[0])
    _saet_ph(s, 0, "Porteføljeoversigt")
    slides.tilsagn_paa_tvaers(prs, r)
    slides.aendringer_pr_fond(prs, r)
    # geografien efter tilsagn vises ikke længere; Geografisk eksponering afløser den
    slides.fordeling(prs, r, "segment", "Porteføljefordeling efter investeringsstrategi")
    slides.brancher(prs, r)             # springes over uden DKK-kurser for alle fonde
    slides.branchekort(prs, r)          # samme fordeling som kasser, der fylder sliden
    slides.branchekort_ejendomme(prs, r)    # ejendomme for sig, resten på sin egen base
    slides.regioner(prs, r)             # springes over uden DKK-kurser for alle fonde
    slides.regionskort(prs, r)          # samme geografi som kasser, der fylder sliden
    slides.regionskort_ejendomme(prs, r)    # ejendomme for sig, resten på sin egen base
    slides.gennemlysning(prs, r)        # springes over, hvis fondene har forskellige valutaer
    slides.vaerdibevaegelser(prs, r)    # springes over, hvis ingen fond har beholdninger

    s = _ny_slide(prs, L_SKILLE[1])
    _saet_ph(s, 0, "Fondsoversigter")
    slides.portefoeljeoversigt(prs, r)  # oversigten over fondene indleder fondssiderne
    for f in r["fonde"]:
        slides.fondsside(prs, f, r["navn"])
        slides.virksomhedsprofiler(prs, f, r["navn"])   # kun fonde med porteføljeselskaber

    slides.definitioner(prs, r)
    _ny_slide(prs, L_SLUT)
    slides.sidefod(prs, r["navn"])      # til sidst, så sidetallene er endelige

    outfile = _gem(prs, Path(outfile or EXPORTS / f"{r['investor']}_{r['as_of']:%Y-%m-%d}.pptx"))
    print(f"gemt: {outfile}  ({len(r['fonde'])} fonde, {len(prs.slides._sldIdLst)} slides)")
    if med_data:
        # dataudtrækket ved siden af decket: samme navn med _data, til intern afstemning
        xlsx = _gem(excel.skriv(r), outfile.with_name(f"{outfile.stem}_data.xlsx"))
        print(f"gemt: {xlsx}  (data og kildedokumenter)")
    _advar_om_datoer(r["datotjek"])
    _advar_om_klassifikation(r["brancher"], "branchen")
    _advar_om_klassifikation(r["regioner"], "geografien")
    if r["valuta"] is None:
        valutaer = ", ".join(sorted({f["valuta"] for f in r["fonde"]}))
        print(f"ADVARSEL: fondene har forskellige valutaer ({valutaer}) - beløb lægges ikke "
              "sammen: ingen samlet række, ingen poolet IRR, og gennemlysningen springes over")
    return outfile


MAKS_VERSIONER = 50


def _gem(prs, outfile):
    """Gem decket (eller projektmappen; begge har .save) og returnér den sti, det endte på.

    PowerPoint låser filen, mens den er åben. I stedet for at fejle lægges et versionsnummer
    på - MOTORTRAMP_2026-06-30_v2.pptx og så videre - så en bygning aldrig går tabt, fordi
    nogen sad med den forrige udgave åben.
    """
    outfile.parent.mkdir(parents=True, exist_ok=True)
    for version in range(1, MAKS_VERSIONER + 1):
        sti = (outfile if version == 1
               else outfile.with_name(f"{outfile.stem}_v{version}{outfile.suffix}"))
        try:
            prs.save(str(sti))
        except PermissionError:
            print(f"{sti.name} er låst (åben i PowerPoint?) - prøver næste versionsnummer")
            continue
        return sti
    raise PermissionError(f"Kunne ikke skrive {outfile.name} i {MAKS_VERSIONER} versioner - "
                          "luk de åbne udgaver i PowerPoint")


def _advar_om_klassifikation(b, hvad):
    """Skriv den del af NAV, hvis branche eller geografi ikke er gennemgået, så den kan rettes
    før udsendelse."""
    if b and b["foreloebig_andel"] and b["foreloebig_andel"] >= 0.0005:
        print(f"ADVARSEL: {hvad} for {b['foreloebig_andel']:.1%} af NAV er en foreløbig "
              f"klassifikation (is_suggested) - {', '.join(b['foreloebige_fonde'])}; "
              "se fanerne Brancher og Geografi i dataudtrækket")


def _advar_om_datoer(tjek):
    """Skriv uens opgørelsesdatoer ud ved bygningen, så de ses før decket sendes videre."""
    if tjek["ens"]:
        print(f"datotjek: alle fonde er opgjort pr. {tjek['rapport_dato']:%d.%m.%Y}")
        return
    for x in tjek["foraeldede"]:
        print(f"ADVARSEL: {x['fond']} er opgjort pr. {x['dato']:%d.%m.%Y}, "
              f"{x['kvartaler']} kvartal(er) før rapportdatoen "
              f"{tjek['rapport_dato']:%d.%m.%Y} - totaler blander opgørelsesdatoer")
    for x in tjek["skaeve_beholdninger"]:
        print(f"ADVARSEL: {x['fond']} har beholdninger pr. "
              f"{x['beholdninger_dato']:%d.%m.%Y}, men nøgletal pr. {x['nav_dato']:%d.%m.%Y}")
