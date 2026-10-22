"""Afstemning af Motortramp-rapporten mod Aletas tal (motortramp_aleta_data.xlsx).

Aleta er et andet system med de samme fonde; udtrækket har én række pr. fond. afstem() sammenligner
rapportens nøgletal med Aletas, kontrol for kontrol. Aletas beløb er i DKK (undtagen nav, der er i
fondens valuta), så rapportens DKK-beløb (n["..._dkk"] fra data.omregn_til_dkk) sammenlignes med
dem: NAV, tilsagn og resttilsagn er omregnet med Nationalbankens kurs på rapportdatoen, indbetalt og
udloddet med kursen på hver kapitalstrøms egen dato, som Aleta gør. test_aleta.py kører
kontrollerne og skriver afstemningsdokumentet (skriv()) med de samme linjer, så de altid er enige:

    python -m pytest tests/test_aleta.py     # testene + exports/aleta_afstemning_<dato>.xlsx
"""
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

def _rvpi(n):
    return n["nav"] / n["indbetalt"] if n["indbetalt"] else None


# (kontrol, Aletas kolonne, rapportens tal, tolerance)
# tolerance: "eksakt", ("relativ", x), ("af_tilsagn", x) eller ("absolut", x)
KONTROLLER = [
    ("Valuta", "navCurrency", lambda f: f["valuta"], "eksakt"),
    ("Vintage", "vintage", lambda f: f["vintage"], "eksakt"),
    ("NAV", "nav", lambda f: f["n"]["nav"], ("relativ", TOLERANCE)),
    ("NAV DKK", "marketValue", lambda f: f["n"]["nav_dkk"], ("relativ", TOLERANCE)),
    ("Tilsagn DKK", "commitment", lambda f: f["n"]["tilsagn_dkk"], ("relativ", TOLERANCE)),
    ("Resttilsagn DKK", "unfundedCommitment", lambda f: f["n"]["resttilsagn_dkk"],
     ("af_tilsagn", TOLERANCE)),
    ("Indbetalt DKK", "drawdowns", lambda f: f["n"]["indbetalt_dkk"], ("relativ", TOLERANCE)),
    ("Udloddet DKK", "distributions", lambda f: f["n"]["udloddet_dkk"], ("relativ", TOLERANCE)),
    ("PICC", "calledPercent", lambda f: f["n"]["picc"], ("absolut", 0.005)),
    ("DPI", "dpi", lambda f: f["n"]["dpi"], ("absolut", 0.01)),
    ("RVPI", "rvpi", lambda f: _rvpi(f["n"]), ("absolut", 0.01)),
    ("TVPI", "tvpi", lambda f: f["n"]["tvpi"], ("absolut", 0.01)),
    ("IRR", "irr", lambda f: f["n"]["irr"], ("absolut", 0.0025)),
]

# beløbskontrollerne -> rapportens tal i fondens valuta, til fanen Detaljer; Aletas DKK-beløb
# regnes tilbage med samme kurs (NAV er i fondens valuta hos begge)
FONDSVALUTA = {"NAV": "nav", "NAV DKK": "nav", "Tilsagn DKK": "tilsagn",
               "Resttilsagn DKK": "resttilsagn", "Indbetalt DKK": "indbetalt",
               "Udloddet DKK": "udloddet"}


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
    bemaerkning og, for beløbene, valuta, kurs og aleta_fv / rapport_fv / afvigelse_fv i fondens
    valuta; Kurs er rapportens kurs (for indbetalt og udloddet gennemsnittet af
    transaktionskurserne, vægtet med beløbene). Status afgøres af DKK-tallene. En fond fra Aleta, der ikke er i rapporten, får status afviger i alle kontroller."""
    fonde = {f["fund_code"]: f for f in r["fonde"]}
    kurser = r.get("kurser", {})
    ud = []
    for a in aleta:
        kode = FONDE.get(a["securityCode"], a["securityCode"])
        f = fonde.get(kode)
        for kontrol, kolonne, hent, tolerance in KONTROLLER:
            linje = dict(fond=kode, aleta_navn=a["asset"], kontrol=kontrol,
                         aleta=_tal(a[kolonne]), rapport=None, afvigelse=None, tilladt=None,
                         status="afviger", bemaerkning="", valuta=None, kurs=None,
                         aleta_fv=None, rapport_fv=None, afvigelse_fv=None)
            if f is None:
                linje["bemaerkning"] = "fonden er ikke i rapporten"
            else:
                linje["rapport"] = _tal(hent(f))
                k = (kurser.get(f["valuta"]) or (None,))[0]
                noegle = FONDSVALUTA.get(kontrol)
                if kontrol in ("Indbetalt DKK", "Udloddet DKK") and f["n"].get(noegle):
                    # rapportens egen omregning med transaktionskurserne, som én gennemsnitskurs
                    k = f["n"][f"{noegle}_dkk"] / f["n"][noegle]
                linje["afvigelse"], linje["tilladt"], linje["status"] = _vurder(
                    linje["aleta"], linje["rapport"], tolerance, _tal(a["commitment"]))
                if kontrol in FONDSVALUTA and k:
                    aleta_fv = linje["aleta"] if kontrol == "NAV" else linje["aleta"] / k
                    rapport_fv = _tal(f["n"].get(FONDSVALUTA[kontrol]))
                    linje.update(valuta=f["valuta"], kurs=_tal(k), aleta_fv=aleta_fv,
                                 rapport_fv=rapport_fv,
                                 afvigelse_fv=None if rapport_fv is None else rapport_fv - aleta_fv)
            ud.append(linje)
    return ud


# --- afstemningsdokumentet -----------------------------------------------------------------------

GROEN = PatternFill("solid", fgColor="C6EFCE")
ROED = PatternFill("solid", fgColor="FFC7CE")
HOVED = PatternFill("solid", fgColor="4F6B5A")
HVID_FED = Font(bold=True, color="FFFFFF")
MARKERING = {"ok": ("✓", GROEN), "afviger": ("✗", ROED)}


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
    ov["A2"] = (f"{antal['ok']} ok, {antal['afviger']} afviger - se fanen Detaljer. Rapportens "
                "beløb i EUR og USD er omregnet til DKK med Nationalbankens kurs: NAV og tilsagn "
                "med kursen på rapportdatoen, indbetalt og udloddet med kursen på hver "
                "transaktionsdato. I Detaljer står beløbene også i fondens valuta; Aletas "
                "DKK-beløb er regnet tilbage med samme kurs")
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
                ("Valuta", "valuta", 8), ("Kurs", "kurs", 9),
                ("Aleta i fondsvaluta", "aleta_fv", 17), ("Rapport i fondsvaluta", "rapport_fv", 17),
                ("Afvigelse i fondsvaluta", "afvigelse_fv", 15),
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
            elif noegle in ("aleta_fv", "rapport_fv", "afvigelse_fv"):
                c.number_format = "#,##0"
            elif noegle == "kurs":
                c.number_format = "0.0000"
            elif noegle == "status":
                c.fill = MARKERING[x["status"]][1]
    de.auto_filter.ref = de.dimensions
    de.freeze_panes = "A2"
    return wb

