"""Kommandolinjen: pe-fund-report MOTORTRAMP [--as-of ÅÅÅÅ-MM-DD] [--out sti]."""
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
    a = ap.parse_args(argv)
    byg(a.investor, a.as_of, a.out, a.master)


if __name__ == "__main__":
    main()
