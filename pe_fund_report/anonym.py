"""Den anonymiserede udgave: samme rapport uden navne på investor, fonde og investeringer.

Rapporten bygges af de rigtige data, så alle beregninger og parringen af beholdninger mellem
kvartaler er de samme; anonymiser() omskriver bagefter den færdige rapport-dict. Beløbene
skaleres ikke her, men i visningen (slides.beloebsskala) - kun beløb, data.py allerede har
skrevet ind i noternes tekst, ganges med faktoren her. laek() er sikkerhedsnettet: den finder
de oprindelige navne i den færdige præsentation, så bygningen kan stoppe, før filen gemmes.
"""
import re

from .data import Note, uden_selskabsform

INVESTOR = "Investor"
BESKRIVELSE = "Beskrivelsen er udeladt i den anonymiserede udgave."
# fondenes nye navne efter segment; et ukendt segment bliver blot "Fond"
FONDSNAVNE = {"Private equity": "Private Equity-fond", "Fund of funds": "Fund of funds",
              "Secondaries": "Secondaries-fond", "Private credit": "Private Credit-fond",
              "Ejendomme": "Ejendomsfond"}
# beholdningslinjer efter type (section i rpt.v_portfolio_holding)
FONDSTYPER = {"FUND", "SECONDARY", "PRIMARY"}
MINDST = 3                # kortere navne søges ikke i teksten (de ligner almindelige ord)
# felter, der kun er til Excel-udtrækket eller afslører fonden direkte
FJERN = ("kilder", "dokumenter", "kursopslag")
# GP'ens antal fonde og selskaber kan slås op; kreditfaciliteten er et beløb og skaleres
FJERN_FAKTA = ("NUMBER_OF_FUNDS", "NUMBER_OF_COMPANIES")
# et beløb i en notes tekst: "17.574 DKK", "1.026 tDKK", "4,5 mio. EUR"
BELOEB = re.compile(r"(?<![\d.,])(\d{1,3}(?:\.\d{3})*(?:,\d+)?)( t| mio\. | )([A-Z]{3})\b")
SELSKABSFORM = re.compile(r"^(A/S|ApS|K/S|P/S)\s+|\s+(A/S|ApS|K/S|P/S|L\.P\.|SCSp)$")


def _varianter(navn):
    """Navnet, som det står, uden selskabsform foran eller bagved."""
    navn = (navn or "").strip()
    return {v for v in (navn, uden_selskabsform(navn), SELSKABSFORM.sub("", navn).strip()) if v}


class _Navne:
    """Oprindelige navne -> pseudonymer, uden hensyn til store og små bogstaver."""

    def __init__(self):
        self.kort = {}                  # lille navn -> pseudonym
        self.koder = {}                 # fondskoder: kun hele strenge, med store bogstaver
        self.taellere = {}
        self.forvaltere = {}            # forvalterens navn -> "Forvalter A", "Forvalter B" ...

    def ny(self, prefiks):
        self.taellere[prefiks] = self.taellere.get(prefiks, 0) + 1
        return f"{prefiks} {self.taellere[prefiks]}"

    def saet(self, navn, pseudonym):
        for v in _varianter(navn):
            self.kort.setdefault(v.lower(), pseudonym)

    def har(self, navn):
        return (navn or "").strip().lower() in self.kort


def _navnekort(r):
    """Byg alle pseudonymer: investor, fonde, de underliggende fonde, beholdninger og forvaltere."""
    nv = _Navne()
    nv.saet(r["investor"], INVESTOR)
    nv.saet(r["navn"], INVESTOR)
    for f in r["fonde"]:
        navn = nv.ny(FONDSNAVNE.get(f["segment"], "Fond"))
        nv.saet(f["fund_name"], navn)
        nv.saet(f["juridisk_navn"], navn)
        nv.koder[f["fund_code"]] = navn
    # feederens underliggende fond først, så de får de laveste numre
    for f in r["fonde"]:
        beh = f["beholdninger"] or {}
        if beh.get("via"):
            navn = nv.ny("Underliggende fond")
            nv.saet(beh["via"], navn)
            nv.koder.setdefault(beh.get("via_kode"), navn)
    for f in r["fonde"]:
        for beh in [f["beholdninger"], *(f.get("beholdninger_primo") or {}).values()]:
            for x in (beh or {}).get("raekker", []):
                if not nv.har(x["navn"]):
                    nv.saet(x["navn"], nv.ny(
                        "Ejendom" if x["type"] == "PROPERTY"
                        else "Underliggende fond" if x["type"] in FONDSTYPER else "Selskab"))
    # forvalterne får bogstaver, så det stadig ses, hvilke fonde der har samme forvalter
    for f in r["fonde"]:
        if f.get("manager") and f["manager"] not in nv.forvaltere:
            bogstav = f"Forvalter {chr(ord('A') + len(nv.forvaltere))}"
            nv.forvaltere[f["manager"]] = bogstav
            if not nv.har(f["manager"]):
                nv.saet(f["manager"], bogstav)
    return nv


def _fjern_identitet(f):
    """Det, der afslører fonden uden at være et navn: logo, beskrivelse, fondsstørrelse og
    ejerandel (fondsstørrelsen kan regnes tilbage af ejerandel og tilsagn)."""
    f.update(logo=None, beskrivelse=BESKRIVELSE, virksomheder={}, benchmark=None,
             nav_afstemning=None)
    f["n"].update(fondsstoerrelse=None, fonds_nav=None, andel=None)
    for beh in [f["beholdninger"], *(f.get("beholdninger_primo") or {}).values()]:
        if not beh:
            continue
        if beh.get("fakta"):
            beh["fakta"] = {k: v for k, v in beh["fakta"].items() if k not in FJERN_FAKTA}
        for x in beh.get("raekker", []):
            if x["type"] != "PROPERTY":
                x["sektor"] = None      # forvalterens egen betegnelse; ejendommenes er værdigrundlaget


def anonymiser(r, skala):
    """(den anonymiserede rapport, de oprindelige navne). r ændres ikke.

    Fondene hedder efter segment ("Private Equity-fond 1"), beholdningerne "Selskab n",
    "Underliggende fond n" og "Ejendom n", forvalterne "Forvalter A", investoren "Investor".
    Logoer, beskrivelser,
    fondsstørrelse, ejerandel og virksomhedsoplysninger fjernes; beløb i noternes tekst
    ganges med skala. originaler er alle oprindelige navne og koder til laek().
    """
    nv = _navnekort(r)
    tekstnavne = sorted((n for n in nv.kort if len(n) >= MINDST), key=len, reverse=True)
    moenster = (re.compile(r"(?<!\w)(" + "|".join(map(re.escape, tekstnavne)) + r")(?!\w)",
                           re.IGNORECASE) if tekstnavne else None)

    def skaler(m):
        tal = float(m.group(1).replace(".", "").replace(",", "."))
        dec = 1 if "," in m.group(1) else 0
        tekst = f"{tal * skala:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return f"{tekst}{m.group(2)}{m.group(3)}"

    def streng(s):
        if s in nv.koder:
            return nv.koder[s]
        if s.strip().lower() in nv.kort:
            return nv.kort[s.strip().lower()]
        ny = moenster.sub(lambda m: nv.kort[m.group(1).lower()], s) if moenster else s
        return BELOEB.sub(skaler, ny) if ny != s or any(c.isdigit() for c in s) else ny

    # samme objekt bliver ét objekt i kopien: fondene går igen i fordelingerne, og det, der
    # fjernes fra r["fonde"] nedenfor, skal også være væk dér
    kopier = {}

    def gaa(x):
        if isinstance(x, Note):
            return Note(streng(x), x.felt)
        if isinstance(x, str):
            return streng(x)
        if id(x) in kopier:
            return kopier[id(x)]
        if isinstance(x, dict):
            ny = kopier[id(x)] = {}
            ny.update((gaa(k), gaa(v)) for k, v in x.items())
            return ny
        if isinstance(x, list):
            ny = kopier[id(x)] = []
            ny.extend(gaa(v) for v in x)
            return ny
        if isinstance(x, tuple):
            return tuple(gaa(v) for v in x)
        return x

    kopi = gaa({k: v for k, v in r.items() if k not in FJERN})
    for rigtig, f in zip(r["fonde"], kopi["fonde"], strict=True):
        f.pop("kilder", None)
        _fjern_identitet(f)
        f["manager"] = nv.forvaltere.get(rigtig.get("manager"))
    pseudonymer = {n.lower() for n in [INVESTOR, *nv.kort.values()]}
    originaler = sorted(n for n in (set(nv.kort) - pseudonymer) | set(nv.koder)
                        if n and len(n) >= MINDST)
    return kopi, originaler


def tekster(prs):
    """Al tekst i præsentationen: figurer, grupper, tabeller, diagrammer og notesider."""
    def figurer(shapes):
        for sh in shapes:
            if sh.shape_type == 6:          # MSO_SHAPE_TYPE.GROUP
                yield from figurer(sh.shapes)
                continue
            if sh.has_text_frame:
                yield sh.text_frame.text
            if getattr(sh, "has_table", False) and sh.has_table:
                for raekke in sh.table.rows:
                    for celle in raekke.cells:
                        yield celle.text
            if getattr(sh, "has_chart", False) and sh.has_chart:
                for plot in sh.chart.plots:
                    yield from (str(c) for c in plot.categories)
                    yield from (s.name or "" for s in plot.series)

    for slide in prs.slides:
        yield from figurer(slide.shapes)
        if slide.has_notes_slide:
            yield slide.notes_slide.notes_text_frame.text


def laek(prs, originaler):
    """De oprindelige navne, der stadig står i præsentationen (tom liste, hvis ingen)."""
    tekst = "\n".join(tekster(prs))
    return [n for n in originaler
            if re.search(r"(?<!\w)" + re.escape(n) + r"(?!\w)", tekst, re.IGNORECASE)]
