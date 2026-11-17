"""Kommandolinjen: pe-fund-report MOTORTRAMP [--as-of ÅÅÅÅ-MM-DD] [--out sti] [--uden-data]
[--uden-supplerende] [--anonym] [--aleta] [--med-pme] [--tusinder]; pe-fund-report --demo bygger demorapporten."""
import argparse

from .report import byg


def main(argv=None):
    ap = argparse.ArgumentParser(
        prog="pe-fund-report",
        description="Bygger investorrapporten som pptx på PPIM-masteren.")
    ap.add_argument("investor", nargs="?",
                    help="investor_code i fund.investor, fx MOTORTRAMP (udelades med --demo)")
    ap.add_argument("--as-of", help="rapportdato (ÅÅÅÅ-MM-DD); standard er seneste NAV-dato")
    ap.add_argument("--out", help="sti til pptx-filen; standard er exports/<INVESTOR>_<dato>.pptx")
    ap.add_argument("--master", help="anden slidemaster end assets/PPIM_slidemaster_2026.pptx")
    ap.add_argument("--uden-data", action="store_true",
                    help="spring dataudtrækket <pptx>_data.xlsx over")
    ap.add_argument("--uden-supplerende", action="store_true",
                    help="spring <pptx>_supplerende_slides.pptx over (slides uden for rapporten)")
    ap.add_argument("--anonym", action="store_true",
                    help="anonymiseret udgave: ingen navne, logoer eller beskrivelser, beløb "
                         "skaleret med en skjult faktor, intet dataudtræk")
    ap.add_argument("--aleta", action="store_true",
                    help="byg også <pptx>_aleta_afstemning.pptx: rapportens nøgletal og IRR mod "
                         "Aletas udtræk (rapporten selv har ingen Aleta-sider)")
    ap.add_argument("--med-pme", action="store_true",
                    help="udgave af rapporten med PME-slidesene efter Samlede investeringer, "
                         "gemt som <INVESTOR>_<dato>_med_pme.pptx")
    ap.add_argument("--tusinder", action="store_true",
                    help="beløb i tusinder (54.400 t.DKK) i stedet for mio. (54,4 mio. DKK); "
                         "gemt som <...>_tusinder.pptx")
    ap.add_argument("--demo", action="store_true",
                    help="demorapporten: en fiktiv investor med syntetiske data, uden databasen")
    a = ap.parse_args(argv)
    if not a.investor and not a.demo:
        ap.error("angiv en investor, fx MOTORTRAMP, eller --demo")
    if a.demo and a.aleta:
        ap.error("--aleta kan ikke bruges med --demo: demoinvestoren har ingen Aleta-udtræk")
    if a.anonym and a.aleta:
        ap.error("--aleta kan ikke bruges med --anonym: Aleta røber investoren")
    byg(a.investor, a.as_of, a.out, a.master, med_data=not a.uden_data, anonym=a.anonym,
        med_supplerende=not a.uden_supplerende, med_aleta=a.aleta, med_pme=a.med_pme,
        demo=a.demo, tusinder=a.tusinder)


if __name__ == "__main__":
    main()
