"""Kommandolinjen: pe-fund-report MOTORTRAMP [--as-of ÅÅÅÅ-MM-DD] [--out sti] [--uden-data]
[--uden-supplerende] [--anonym]."""
import argparse

from .report import byg


def main(argv=None):
    ap = argparse.ArgumentParser(
        prog="pe-fund-report",
        description="Bygger investorrapporten som pptx på PPIM-masteren.")
    ap.add_argument("investor", help="investor_code i fund.investor, fx MOTORTRAMP")
    ap.add_argument("--as-of", help="rapportdato (ÅÅÅÅ-MM-DD); standard er seneste NAV-dato")
    ap.add_argument("--out", help="sti til pptx-filen; standard er exports/<INVESTOR>_<dato>.pptx")
    ap.add_argument("--master", help="anden slidemaster end assets/ppim_slidemaster.pptx")
    ap.add_argument("--uden-data", action="store_true",
                    help="spring dataudtrækket <pptx>_data.xlsx over")
    ap.add_argument("--uden-supplerende", action="store_true",
                    help="spring <pptx>_supplerende_slides.pptx over (slides uden for rapporten)")
    ap.add_argument("--anonym", action="store_true",
                    help="anonymiseret udgave: ingen navne, logoer eller beskrivelser, beløb "
                         "skaleret med en skjult faktor, intet dataudtræk")
    a = ap.parse_args(argv)
    byg(a.investor, a.as_of, a.out, a.master, med_data=not a.uden_data, anonym=a.anonym,
        med_supplerende=not a.uden_supplerende)


if __name__ == "__main__":
    main()
