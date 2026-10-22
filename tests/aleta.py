"""Afstemning af Motortramp-rapporten mod Aletas tal (motortramp_aleta_data.xlsx).

Aleta er et andet system med de samme fonde; udtrækket har én række pr. fond. afstem() sammenligner
rapportens nøgletal med Aletas, kontrol for kontrol. test_aleta.py og afstemningsdokumentet bruger
begge afstem(), så de altid er enige:

    python tests/aleta.py       # skriver exports/aleta_afstemning_<dato>.xlsx
"""
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

ALETA = Path(__file__).with_name("motortramp_aleta_data.xlsx")
INVESTOR = "MOTORTRAMP"
AS_OF = date(2026, 6, 30)       # Aleta skriver 01.07.2026, dagen efter kvartalsslut
TOLERANCE = 0.005               # som data.GP_TOLERANCE

# Aletas securityCode -> vores fund_code; navnene ligner ikke altid hinanden (Park Square er vores
# feeder PPPC2, Newbury PPPE2, "API Property DK" er CRED)
FONDE = {
    "APIPropertyDK(CRED)": "CRED",
    "KAPE2020": "KAPE20",
    "KAPE2021": "KAPE21",
    "KAPE2022": "KAPE22",
    "Seed Capital": "SEED4",
    "Home.Earth": "HOME",
    "Private Debt III": "PD3LGT",
    "Catacap III": "CATACAP3",
    "Park Square - Private Credit II": "PPPC2",
    "Newbury - Private Equity II": "PPPE2",
}

IKKE_SAMMENLIGNELIG = "ikke sammenlignelig"
KUN_DKK = "Aleta omregner kapitalstrømmene til DKK til egne kurser; rapporten har dem i fondens valuta"


def _rvpi(n):
    return n["nav"] / n["indbetalt"] if n["indbetalt"] else None


# (kontrol, Aletas kolonne, rapportens tal, tolerance, kun DKK-fonde)
# tolerance: "eksakt", ("relativ", x), ("af_tilsagn", x) eller ("absolut", x)
KONTROLLER = [
    ("Valuta", "navCurrency", lambda f: f["valuta"], "eksakt", False),
    ("Vintage", "vintage", lambda f: f["vintage"], "eksakt", False),
    ("NAV", "nav", lambda f: f["n"]["nav"], ("relativ", TOLERANCE), False),
    ("NAV DKK", "marketValue", lambda f: f["n"]["nav_dkk"], ("relativ", TOLERANCE), False),
    ("Tilsagn DKK", "commitment", lambda f: f["n"]["tilsagn_dkk"], ("relativ", TOLERANCE), False),
    ("Resttilsagn DKK", "unfundedCommitment", lambda f: f["n"]["resttilsagn_dkk"],
     ("af_tilsagn", TOLERANCE), False),
    ("Indbetalt", "drawdowns", lambda f: f["n"]["indbetalt"], ("relativ", TOLERANCE), True),
    ("Udloddet", "distributions", lambda f: f["n"]["udloddet"], ("relativ", TOLERANCE), True),
    ("PICC", "calledPercent", lambda f: f["n"]["picc"], ("absolut", 0.005), False),
    ("DPI", "dpi", lambda f: f["n"]["dpi"], ("absolut", 0.01), False),
    ("RVPI", "rvpi", lambda f: _rvpi(f["n"]), ("absolut", 0.01), False),
    ("TVPI", "tvpi", lambda f: f["n"]["tvpi"], ("absolut", 0.01), False),
    ("IRR", "irr", lambda f: f["n"]["irr"], ("absolut", 0.0025), False),
]


def laes_aleta(sti=ALETA):
    """Aletas rækker som dicts med kolonnenavnene som nøgler; vintage som årstal."""
    ws = load_workbook(sti, read_only=True, data_only=True).worksheets[0]
    raekker = ws.iter_rows(values_only=True)
    hoved = next(raekker)
    ud = []
    for r in raekker:
        if not any(r):
            continue
        x = dict(zip(hoved, r, strict=True))
        if hasattr(x["vintage"], "year"):
            x["vintage"] = x["vintage"].year
        ud.append(x)
    return ud


def _tal(v):
    """np.float64 og int som float, så sammenligning og Excel får almindelige tal."""
    return float(v) if isinstance(v, int | float) and not isinstance(v, bool) else v


def _vurder(aleta, rapport, tolerance, tilsagn):
    """(afvigelse, tilladt afvigelse, status) for ét tal."""
    if aleta is None or rapport is None:
        return None, None, "afviger"
    if tolerance == "eksakt":
        return None, None, "ok" if aleta == rapport else "afviger"
    art, x = tolerance
    afvigelse = rapport - aleta
    tilladt = {"relativ": abs(aleta) * x, "af_tilsagn": abs(tilsagn or 0) * x, "absolut": x}[art]
    return afvigelse, tilladt, "ok" if abs(afvigelse) <= tilladt else "afviger"


def afstem(aleta, r):
    """Én række pr. fond og kontrol: fond, kontrol, aleta, rapport, afvigelse, tilladt, status,
    bemaerkning. En fond fra Aleta, der ikke er i rapporten, får status afviger i alle kontroller."""
    fonde = {f["fund_code"]: f for f in r["fonde"]}
    ud = []
    for a in aleta:
        kode = FONDE.get(a["securityCode"], a["securityCode"])
        f = fonde.get(kode)
        for kontrol, kolonne, hent, tolerance, kun_dkk in KONTROLLER:
            linje = dict(fond=kode, aleta_navn=a["asset"], kontrol=kontrol,
                         aleta=_tal(a[kolonne]), rapport=None, afvigelse=None, tilladt=None,
                         status="afviger", bemaerkning="")
            if f is None:
                linje["bemaerkning"] = "fonden er ikke i rapporten"
            else:
                linje["rapport"] = _tal(hent(f))
                if kun_dkk and f["valuta"] != "DKK":
                    linje.update(status=IKKE_SAMMENLIGNELIG, bemaerkning=KUN_DKK)
                else:
                    linje["afvigelse"], linje["tilladt"], linje["status"] = _vurder(
                        linje["aleta"], linje["rapport"], tolerance, _tal(a["commitment"]))
            ud.append(linje)
    return ud


# --- afstemningsdokumentet -----------------------------------------------------------------------

GROEN = PatternFill("solid", fgColor="C6EFCE")
ROED = PatternFill("solid", fgColor="FFC7CE")
GRAA = PatternFill("solid", fgColor="E7E6E6")
HOVED = PatternFill("solid", fgColor="4F6B5A")
HVID_FED = Font(bold=True, color="FFFFFF")
MARKERING = {"ok": ("✓", GROEN), "afviger": ("✗", ROED), IKKE_SAMMENLIGNELIG: ("–", GRAA)}


def _format(kontrol):
    if kontrol in ("PICC", "IRR"):
        return "0.00%"
    if kontrol in ("DPI", "RVPI", "TVPI"):
        return "0.000"
    return "#,##0"


def skriv(linjer, as_of=AS_OF):
    """Projektmappe med fanerne Oversigt (fond × kontrol) og Detaljer (én række pr. kontrol)."""
    wb = Workbook()
    ov = wb.active
    ov.title = "Oversigt"
    antal = {s: sum(x["status"] == s for x in linjer) for s in MARKERING}
    ov["A1"] = f"Motortramp: rapporten mod Aleta pr. {as_of:%d.%m.%Y}"
    ov["A1"].font = Font(bold=True, size=14)
    ov["A2"] = (f"{antal['ok']} ok, {antal['afviger']} afviger, "
                f"{antal[IKKE_SAMMENLIGNELIG]} ikke sammenlignelige - se fanen Detaljer")
    kontroller = [k[0] for k in KONTROLLER]
    fonde = list(dict.fromkeys(x["fond"] for x in linjer))
    for j, navn in enumerate(["Fond", *kontroller], start=1):
        c = ov.cell(4, j, navn)
        c.fill, c.font, c.alignment = HOVED, HVID_FED, Alignment(horizontal="center")
    opslag = {(x["fond"], x["kontrol"]): x for x in linjer}
    for i, fond in enumerate(fonde, start=5):
        ov.cell(i, 1, fond).font = Font(bold=True)
        for j, kontrol in enumerate(kontroller, start=2):
            tegn, fyld = MARKERING[opslag[fond, kontrol]["status"]]
            c = ov.cell(i, j, tegn)
            c.fill, c.alignment = fyld, Alignment(horizontal="center")
    ov.column_dimensions["A"].width = 12
    for j in range(2, len(kontroller) + 2):
        ov.column_dimensions[ov.cell(4, j).column_letter].width = 13
    ov.freeze_panes = "B5"

    de = wb.create_sheet("Detaljer")
    kolonner = [("Fond", "fond", 11), ("Navn i Aleta", "aleta_navn", 34), ("Kontrol", "kontrol", 16),
                ("Aleta", "aleta", 15), ("Rapport", "rapport", 15), ("Afvigelse", "afvigelse", 13),
                ("Tilladt afvigelse", "tilladt", 15), ("Status", "status", 18),
                ("Bemærkning", "bemaerkning", 60)]
    for j, (navn, _, bredde) in enumerate(kolonner, start=1):
        c = de.cell(1, j, navn)
        c.fill, c.font = HOVED, HVID_FED
        de.column_dimensions[c.column_letter].width = bredde
    for i, x in enumerate(linjer, start=2):
        for j, (_, noegle, _) in enumerate(kolonner, start=1):
            c = de.cell(i, j, x[noegle])
            if noegle in ("aleta", "rapport", "afvigelse", "tilladt") and x["kontrol"] != "Vintage":
                c.number_format = _format(x["kontrol"])
        de.cell(i, 8).fill = MARKERING[x["status"]][1]
    de.auto_filter.ref = de.dimensions
    de.freeze_panes = "A2"
    return wb


if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    from pe_fund_report import data, report
    from pe_fund_report.paths import EXPORTS

    linjer = afstem(laes_aleta(), data.investor_rapport(INVESTOR, AS_OF.isoformat()))
    sti = report._gem(skriv(linjer), EXPORTS / f"aleta_afstemning_{AS_OF:%Y-%m-%d}.xlsx")
    print(f"gemt: {sti}")
    for x in linjer:
        if x["status"] == "afviger":
            print(f"AFVIGER: {x['fond']} {x['kontrol']}: Aleta {x['aleta']}, rapport {x['rapport']}")
