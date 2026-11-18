"""Syntetiske data til demorapporten: en lille, fiktiv investor med otte fonde.

Tabellerne har databasens navne og kolonner (fund.*, rpt.*, dbo.*, portfolio.*), så data.py's
egne forespørgsler kan køre mod dem (demo/__init__.py lægger dem i SQLite). Alt er seedet og
dermed ens fra kørsel til kørsel.

Fondene simuleres nedefra: hver beholdning har sine træk (investeringer), sin værdiudvikling
(egen drift, et fælles marked og støj) og sine udlodninger (exit, løbende eller renter).
Fondens kald er trækkene plus gebyrer, udlodningerne er beholdningernes, og NAV er
dagsværdien plus likvider minus hensat carry; investors tal er ejerandelen af fondens. Så
hænger NAV, cash flows, kapitalkonto og beholdninger sammen, som de gør hos en rigtig GP.
Fondens drift kalibreres, så TVPI rammer et realistisk niveau for strategi og alder.

Navnene er generiske som i den anonymiserede udgave: "Private Equity-fond 1", "Selskab 3",
"Ejendom 2", "Forvalter A".
"""
from dataclasses import dataclass, field
from datetime import date, timedelta
from functools import lru_cache

import numpy as np
import pandas as pd

from .. import data

SEED = 20260930
DATO = date(2026, 9, 30)            # demorapportens dato
INVESTOR = "DEMO"
INVESTOR_NAVN = "Demo-investor"
FX_START = date(2012, 1, 2)


# ---------------------------------------------------------------------------------------------
# datoer

def kvartalsslut(d):
    m = (d.month - 1) // 3 * 3 + 3
    return (date(d.year, m + 1, 1) if m < 12 else date(d.year + 1, 1, 1)) - timedelta(days=1)


def kvartaler(fra, til):
    """Kvartalsultimoerne fra fra's kvartal til og med til."""
    ud, q = [], kvartalsslut(fra)
    while q <= til:
        ud.append(q)
        q = kvartalsslut(q + timedelta(days=1))
    return ud


def hverdag(d):
    while d.weekday() >= 5:
        d += timedelta(days=1)
    return d


def _i_kvartalet(q, maaned, dag):
    """En hverdag i q's kvartal: maaned 1-3 i kvartalet."""
    return hverdag(date(q.year, q.month - 3 + maaned, dag))


# ---------------------------------------------------------------------------------------------
# markedet og valutakurserne

# kvartalsvist log-afkast i markedet, hvor det afveg fra det normale (+1 %)
MARKED = {(2015, 3): -0.06, (2016, 1): -0.03, (2018, 4): -0.08, (2020, 1): -0.15,
          (2020, 2): 0.10, (2020, 3): 0.06, (2020, 4): 0.08, (2021, 1): 0.04, (2021, 2): 0.04,
          (2021, 3): 0.03, (2021, 4): 0.04, (2022, 1): -0.04, (2022, 2): -0.10,
          (2022, 3): -0.03, (2022, 4): 0.03, (2023, 1): 0.02, (2023, 4): 0.04,
          (2025, 1): -0.03, (2025, 2): 0.04, (2026, 1): 0.0, (2026, 2): 0.02, (2026, 3): 0.005}


def marked(q):
    return MARKED.get((q.year, (q.month - 1) // 3 + 1), 0.01)


# USD/DKK nogenlunde som den gik; EUR ligger fast omkring centralkursen
USD_ANKRE = [(date(2012, 1, 2), 5.75), (date(2013, 6, 30), 5.70), (date(2014, 6, 30), 5.45),
             (date(2015, 3, 31), 6.95), (date(2016, 12, 31), 7.05), (date(2017, 12, 31), 6.25),
             (date(2018, 12, 31), 6.50), (date(2019, 9, 30), 6.80), (date(2020, 12, 31), 6.05),
             (date(2021, 12, 31), 6.55), (date(2022, 9, 30), 7.60), (date(2023, 6, 30), 6.85),
             (date(2024, 6, 30), 6.95), (date(2024, 12, 31), 7.15), (date(2025, 6, 30), 6.40),
             (DATO, 6.35)]


def valutakurser(rng):
    """{valuta: pd.Series(dato -> DKK pr. enhed)} for hver bankdag fra FX_START til DATO."""
    dage = [d.date() for d in pd.bdate_range(FX_START, DATO)]
    x = np.array([(d - FX_START).days for d in dage], dtype=float)
    ax = np.array([(d - FX_START).days for d, _ in USD_ANKRE], dtype=float)
    usd = np.interp(x, ax, [k for _, k in USD_ANKRE])
    stoej = np.zeros(len(dage))
    for i in range(1, len(dage)):      # middelreverterende støj omkring forløbet
        stoej[i] = 0.97 * stoej[i - 1] + rng.normal(0, 0.004)
    usd = usd * np.exp(stoej)
    eur = np.clip(7.445 + np.cumsum(rng.normal(0, 0.0015, len(dage))) * 0.1, 7.43, 7.47)
    eur = 7.445 + (eur - 7.445) * 0.8
    return {"USD": pd.Series(np.round(usd, 4), index=dage),
            "EUR": pd.Series(np.round(eur, 4), index=dage)}


def kurs_paa(kurser, valuta, d):
    if valuta == "DKK":
        return 1.0
    s = kurser[valuta]
    return float(s[s.index <= d].iloc[-1])


# ---------------------------------------------------------------------------------------------
# GICS og geografi

# branchegrupper pr. sektor, de almindeligste først, og GP'ens egen betegnelse for dem
GRUPPER = {
    "INFORMATION_TECHNOLOGY": [(4510, "Software"), (4510, "IT-services"),
                               (4520, "Hardware"), (4530, "Halvledere")],
    "HEALTH_CARE": [(3510, "Medtech"), (3520, "Life science"), (3510, "Sundhedsydelser")],
    "INDUSTRIALS": [(2010, "Industri"), (2020, "Erhvervsservice"), (2030, "Transport")],
    "CONSUMER_DISCRETIONARY": [(2530, "Forbrugerservice"), (2550, "Detail"),
                               (2520, "Forbrugsgoder")],
    "FINANCIALS": [(4020, "Finansielle tjenester"), (4030, "Forsikring")],
    "COMMUNICATION_SERVICES": [(5020, "Medier"), (5010, "Telekommunikation")],
    "CONSUMER_STAPLES": [(3020, "Fødevarer"), (3030, "Husholdningsprodukter")],
    "MATERIALS": [(1510, "Materialer")],
    "ENERGY": [(1010, "Energi")],
    "UTILITIES": [(5510, "Forsyning")],
    "REAL_ESTATE": [(6020, "Ejendomme")],
}
GRUPPENAVNE = {4510: "Software & Services", 4520: "Technology Hardware & Equipment",
               4530: "Semiconductors & Semiconductor Equipment",
               3510: "Health Care Equipment & Services",
               3520: "Pharmaceuticals, Biotechnology & Life Sciences", 2010: "Capital Goods",
               2020: "Commercial & Professional Services", 2030: "Transportation",
               2530: "Consumer Services", 2550: "Consumer Discretionary Distribution & Retail",
               2520: "Consumer Durables & Apparel", 4020: "Financial Services",
               4030: "Insurance", 5020: "Media & Entertainment",
               5010: "Telecommunication Services", 3020: "Food, Beverage & Tobacco",
               3030: "Household & Personal Products", 1510: "Materials", 1010: "Energy",
               5510: "Utilities", 6010: "Equity REITs",
               6020: "Real Estate Management & Development"}
DELREGION_REGION = {"US": "NORTH_AMERICA", "CANADA": "NORTH_AMERICA", "DENMARK": "EUROPE",
                    "NORDICS": "EUROPE", "UK_IRELAND": "EUROPE", "DACH": "EUROPE",
                    "FRANCE": "EUROPE", "BENELUX": "EUROPE", "SOUTHERN_EUROPE": "EUROPE",
                    "ASIA_PACIFIC": "ROW", "LATIN_AMERICA": "ROW", "MIDDLE_EAST_AFRICA": "ROW"}
SUBREGION_ENGELSK = {"US": "United States", "CANADA": "Canada", "DENMARK": "Denmark",
                     "NORDICS": "Nordics excl. Denmark", "UK_IRELAND": "UK & Ireland",
                     "DACH": "DACH", "FRANCE": "France", "BENELUX": "Benelux",
                     "SOUTHERN_EUROPE": "Southern Europe", "ASIA_PACIFIC": "Asia Pacific",
                     "LATIN_AMERICA": "Latin America", "MIDDLE_EAST_AFRICA": "Middle East & Africa"}

# en kort dansk beskrivelse pr. GP-betegnelse til virksomhedskortene
BESKRIVELSER = {
    "Software": ["Udvikler abonnementsbaseret software til økonomistyring i mellemstore "
                 "virksomheder.",
                 "Leverer en cloudplatform til planlægning og drift i logistikbranchen.",
                 "Software til compliance og risikostyring for banker og forsikringsselskaber."],
    "IT-services": ["Driver og sikrer IT-infrastruktur for kommuner og regioner.",
                    "Konsulenthus inden for dataplatforme og kunstig intelligens."],
    "Hardware": ["Producerer sensorer og styreenheder til industriel automation."],
    "Halvledere": ["Designer chips til strømstyring i elbiler og vindmøller."],
    "Medtech": ["Udvikler og sælger engangsudstyr til hospitaler i Nordeuropa.",
                "Producent af høreapparater og tilbehør med egen distribution."],
    "Life science": ["Kontraktproducent af biologiske lægemidler i klinisk fase.",
                     "Udvikler diagnostiske test til tidlig opsporing af infektioner."],
    "Sundhedsydelser": ["Kæde af tandklinikker med over 60 klinikker.",
                        "Driver private fysioterapi- og genoptræningscentre."],
    "Industri": ["Producent af ventiler og pumper til vand- og spildevandsanlæg.",
                 "Specialiseret leverandør af komponenter til vindmølleindustrien."],
    "Erhvervsservice": ["Teknisk inspektion, test og certificering for industrien.",
                        "Facility services med fokus på rengøring og ejendomsdrift."],
    "Transport": ["Kølelogistik og distribution af fødevarer i Norden."],
    "Forbrugerservice": ["Kæde af fitnesscentre med medlemskab uden binding.",
                         "Udbyder af efteruddannelse og certificeringskurser online."],
    "Detail": ["Onlineforhandler af sports- og outdoorudstyr."],
    "Forbrugsgoder": ["Designer og producent af møbler til hjem og kontor."],
    "Finansielle tjenester": ["Betalingsløsninger til små og mellemstore webshops.",
                              "Administration af pensionsordninger for virksomheder."],
    "Forsikring": ["Specialforsikring inden for erhverv og transport."],
    "Medier": ["Digital medievirksomhed med abonnementer og events."],
    "Telekommunikation": ["Fibernetværk i mindre byer og landdistrikter."],
    "Fødevarer": ["Producent af plantebaserede fødevarer til detailhandlen."],
    "Husholdningsprodukter": ["Producent af miljømærkede rengøringsmidler."],
    "Materialer": ["Producent af emballage af genanvendt pap og plast."],
    "Energi": ["Service og vedligehold af energiinfrastruktur."],
    "Forsyning": ["Udvikler og driver solcelleparker i Sydeuropa."],
}


# ---------------------------------------------------------------------------------------------
# fondene

@dataclass
class Fond:
    kode: str
    navn: str                       # juridisk navn; decket fjerner selskabsformen
    forvalter: str
    valuta: str
    vintage: int
    foerste_kald: date
    tilsagn: float                  # investors tilsagn
    andel: float                    # investors ejerandel (tilsagn / fondens tilsagn)
    fokus: str                      # strategy_focus
    struktur: str | None            # fund_structure
    geografi: str                   # geography_focus
    art: str                        # selskaber, fonde, secondaries, kredit, ejendomme, balance
    antal: int                      # beholdninger
    sektorer: dict                  # sektor -> vægt
    regioner: dict                  # delregion -> vægt
    tvpi: float                     # den TVPI, driften kalibreres til
    beskrivelse: str
    benchmark: str | None = None
    loebetid: tuple = (10, 2)
    nav_til: date = DATO
    gebyr: float = 0.0175           # årligt gebyr af fondens tilsagn i investeringsperioden
    re_sektor: str | None = None
    evergreen: bool = False
    underliggende: str | None = None    # feederens masterfond
    ekstra: dict = field(default_factory=dict)


FORVALTERE = {"PE1": "Forvalter A", "PE2": "Forvalter B", "PE3": "Forvalter C",
              "FOF1": "Forvalter D", "SEK1": "Forvalter E", "PCM1": "Forvalter F",
              "PC1": "Demo Feeder ApS", "EJD1": "Forvalter G", "EJD2": "Forvalter H"}

FONDE = [
    Fond("PE1", "Private Equity-fond 1 SCSp", "Forvalter A", "EUR", 2018, date(2018, 9, 14),
         15_000_000, 0.025, "BUYOUT", None, "EUROPE", "selskaber", 11,
         {"INDUSTRIALS": 28, "INFORMATION_TECHNOLOGY": 22, "HEALTH_CARE": 20,
          "CONSUMER_DISCRETIONARY": 12, "FINANCIALS": 10, "CONSUMER_STAPLES": 8},
         {"DACH": 20, "NORDICS": 20, "UK_IRELAND": 15, "FRANCE": 10, "BENELUX": 10, "US": 25},
         1.75,
         "Europæisk buyoutfond, der køber majoriteter i mellemstore, ofte "
         "familieejede virksomheder i Nordeuropa. Forvalteren skaber værdi gennem "
         "opkøbsstrategier, professionalisering og internationalisering.",
         benchmark="MSCI Europe Small Cap",
         ekstra=dict(udligning=-61_250.0, genkaldelig=True, omvurderet=date(2025, 12, 31))),
    Fond("PE2", "Private Equity-fond 2, L.P.", "Forvalter B", "USD", 2021, date(2021, 11, 9),
         12_000_000, 0.012, "GROWTH", None, "NORTH_AMERICA", "selskaber", 10,
         {"INFORMATION_TECHNOLOGY": 45, "HEALTH_CARE": 20, "FINANCIALS": 15,
          "COMMUNICATION_SERVICES": 10, "CONSUMER_DISCRETIONARY": 10},
         {"US": 70, "CANADA": 10, "UK_IRELAND": 10, "ASIA_PACIFIC": 10},
         1.18,
         "Nordamerikansk growth equity-fond, der tager minoritets- og majoritetsposter i "
         "profitable software- og teknologivirksomheder i vækst.",
         benchmark="MSCI USA Small Cap", nav_til=date(2026, 6, 30), gebyr=0.02),
    Fond("PE3", "Private Equity-fond 3 K/S", "Forvalter C", "DKK", 2016, date(2016, 5, 20),
         60_000_000, 0.08, "VENTURE", None, "NORDICS", "selskaber", 14,
         {"INFORMATION_TECHNOLOGY": 45, "HEALTH_CARE": 32, "INDUSTRIALS": 13,
          "CONSUMER_DISCRETIONARY": 10},
         {"DENMARK": 65, "NORDICS": 25, "US": 10},
         2.35,
         "Nordisk venturefond, der investerer i teknologi- og life science-virksomheder fra "
         "seed til vækstfasen og følger op i de senere runder.",
         benchmark="OMX Copenhagen Small Cap", loebetid=(10, 3), gebyr=0.025),
    Fond("FOF1", "Fund of funds 1, L.P.", "Forvalter D", "USD", 2014, date(2014, 10, 8),
         9_000_000, 0.03, "BUYOUT", "FUND_OF_FUNDS", "GLOBAL", "fonde", 12,
         {"INFORMATION_TECHNOLOGY": 28, "INDUSTRIALS": 20, "HEALTH_CARE": 15,
          "CONSUMER_DISCRETIONARY": 15, "FINANCIALS": 10, "COMMUNICATION_SERVICES": 7,
          "ENERGY": 5},
         {"US": 55, "CANADA": 5, "UK_IRELAND": 10, "DACH": 8, "FRANCE": 7, "ASIA_PACIFIC": 10,
          "LATIN_AMERICA": 5},
         1.95,
         "Global fund of funds med tilsagn til omkring tolv buyout- og growthfonde i USA, "
         "Europa og Asien, fordelt over tre vintageår.",
         benchmark="MSCI ACWI", loebetid=(12, 2), gebyr=0.0075),
    Fond("SEK1", "Secondaries-fond 1 SCSp", "Forvalter E", "EUR", 2020, date(2020, 7, 15),
         10_000_000, 0.008, "BUYOUT", "SECONDARY", "EUROPE_NORTH_AMERICA", "secondaries", 11,
         {"INDUSTRIALS": 24, "INFORMATION_TECHNOLOGY": 22, "HEALTH_CARE": 15,
          "CONSUMER_DISCRETIONARY": 10, "FINANCIALS": 10, "CONSUMER_STAPLES": 9,
          "MATERIALS": 5, "UTILITIES": 5},
         {"US": 50, "UK_IRELAND": 15, "DACH": 13, "FRANCE": 7, "NORDICS": 5,
          "ASIA_PACIFIC": 10},
         1.48,
         "Secondaryfond, der køber LP-andele i modne buyoutfonde med rabat og supplerer med "
         "co-investeringer sammen med førende forvaltere.",
         benchmark="MSCI World Small Cap", loebetid=(8, 2), gebyr=0.01),
    Fond("PC1", "Private Credit-fond 1 K/S", "Demo Feeder ApS", "EUR", 2022, date(2022, 6, 22),
         12_000_000, 0.024, "JUNIOR_DEBT", None, "EUROPE", "kredit", 16,
         {"INFORMATION_TECHNOLOGY": 25, "HEALTH_CARE": 25, "INDUSTRIALS": 25,
          "COMMUNICATION_SERVICES": 10, "CONSUMER_STAPLES": 10, "MATERIALS": 5},
         {"UK_IRELAND": 20, "FRANCE": 20, "DACH": 25, "BENELUX": 10, "SOUTHERN_EUROPE": 5,
          "US": 20},
         1.13,
         "Feeder til en europæisk direct lending-fond, der yder unitranche- og "
         "juniorlån til PE-ejede mellemstore virksomheder med variabel rente.",
         benchmark="Morningstar European Leveraged Loan", loebetid=(7, 1), gebyr=0.01,
         underliggende="PCM1", ekstra=dict(rabat=0.25)),
    Fond("EJD1", "Ejendomsfond 1 K/S", "Forvalter G", "DKK", 2019, date(2019, 9, 20),
         45_000_000, 0.0, "REAL_ESTATE", None, "DENMARK", "ejendomme", 7,
         {"REAL_ESTATE": 100}, {"DENMARK": 100}, 1.24,
         "Dansk evergreen-ejendomsfond, der udvikler og ejer bæredygtige boligejendomme i "
         "de større danske byer og løbende optager ny kapital.",
         re_sektor="RESIDENTIAL", evergreen=True, gebyr=0.0),
    Fond("EJD2", "Ejendomsfond 2 SCSp", "Forvalter H", "EUR", 2017, date(2017, 4, 11),
         8_000_000, 0.04, "REAL_ESTATE", None, "EUROPE", "balance", 1,
         {"REAL_ESTATE": 100}, {"DACH": 50, "BENELUX": 30, "NORDICS": 20}, 1.32,
         "Lukket europæisk logistikfond, der ejer moderne lager- og distributionsejendomme "
         "tæt på de store transportkorridorer i Tyskland, Benelux og Norden.",
         re_sektor="LOGISTICS", loebetid=(8, 2), gebyr=0.012),
]
MASTER = dict(kode="PCM1", navn="Private Credit-masterfond", forvalter="Forvalter F")


# ---------------------------------------------------------------------------------------------
# beholdningerne

@dataclass
class Beholdning:
    navn: str
    sektion: str                    # COMPANY, FUND, SECONDARY, CO_INVESTMENT, LOAN
    sektor: str
    gruppe: int
    gp_betegnelse: str
    delregion: str
    start: int                      # kvartalets indeks for første træk
    traek: dict                     # kvartalsindeks -> beløb (hele fondens)
    drift: float                    # årlig log-drift
    vol: float
    beta: float
    maade: str                      # exit, loebende, rente
    exit: int | None = None
    rente: float = 0.0
    tilsagn: float | None = None    # underliggende fondes tilsagn
    instrument: str | None = None
    vintage: int | None = None


def _kvoter(antal, vaegte, stoerrelser):
    """Fordel antal beholdninger (størst først) på kategorierne, så den størrelsesvægtede
    fordeling ligger tæt på vaegte: hver beholdning går til kategorien med størst underskud."""
    i_alt = sum(vaegte.values())
    maal = {k: v / i_alt for k, v in vaegte.items()}
    tildelt = {k: 0.0 for k in vaegte}
    samlet = sum(stoerrelser)
    ud = [None] * antal
    for i in sorted(range(antal), key=lambda i: -stoerrelser[i]):
        k = max(maal, key=lambda k: maal[k] * samlet - tildelt[k])
        ud[i] = k
        tildelt[k] += stoerrelser[i]
    return ud


class Navne:
    """Fortløbende generiske navne på tværs af fondene, som i den anonymiserede udgave."""

    def __init__(self):
        self.t = {}

    def ny(self, prefiks):
        self.t[prefiks] = self.t.get(prefiks, 0) + 1
        return f"{prefiks} {self.t[prefiks]}"


def _beholdninger(f, rng, n_kv, fondsstoerrelse, navne):
    """Fondens beholdninger med træk, drift og udlodningsmåde."""
    deployeres = fondsstoerrelse * {"kredit": 0.88, "selskaber": 0.78}.get(f.art, 0.82)
    stoerrelser = list(rng.lognormal(0, 0.55, f.antal))
    if f.art == "selskaber" and f.fokus == "VENTURE":
        stoerrelser = list(rng.lognormal(0, 0.8, f.antal))
    sektorer = _kvoter(f.antal, f.sektorer, stoerrelser)
    regioner = _kvoter(f.antal, f.regioner, list(rng.permutation(stoerrelser)))
    s = sum(stoerrelser)
    investeringsperiode = 14 if f.art != "secondaries" else 12
    ud = []
    for i in range(f.antal):
        beloeb = deployeres * stoerrelser[i] / s
        start = int(round(i * investeringsperiode / f.antal + rng.uniform(0, 1.5)))
        gruppe, betegnelse = GRUPPER[sektorer[i]][int(rng.integers(len(GRUPPER[sektorer[i]])))]
        faelles = dict(sektor=sektorer[i], gruppe=gruppe, gp_betegnelse=betegnelse,
                       delregion=regioner[i], start=start)
        if f.art == "selskaber":
            venture = f.fokus == "VENTURE"
            # venture: første runde og to opfølgninger; buyout og growth: hele beløbet på én gang
            traek = ({start: beloeb * 0.4, start + 4: beloeb * 0.35, start + 9: beloeb * 0.25}
                     if venture else {start: beloeb})
            ud.append(Beholdning(navne.ny("Selskab"), "COMPANY", traek=traek,
                                 drift=float(np.clip(rng.normal(0.10, 0.20), -0.35, 0.30)) if venture
                                 else rng.normal(0.10, 0.10),
                                 vol=0.06 if venture else 0.035, beta=0.5, maade="exit",
                                 # venture holdes længere end buyout
                                 exit=start + int(rng.integers(24, 44) if venture
                                                  else rng.integers(18, 30)),
                                 **faelles))
        elif f.art == "fonde":
            traek = {start + k: beloeb / 12 for k in range(12)}
            ud.append(Beholdning(navne.ny("Underliggende fond"), "FUND", traek=traek,
                                 drift=rng.normal(0.13, 0.05), vol=0.02, beta=0.6,
                                 maade="loebende", tilsagn=beloeb / 0.95,
                                 vintage=(f.foerste_kald.year + start // 4), **faelles))
        elif f.art == "secondaries":
            if i % 3 == 2:      # hver tredje er en co-investering
                ud.append(Beholdning(navne.ny("Selskab"), "CO_INVESTMENT", traek={start: beloeb},
                                     drift=rng.normal(0.12, 0.08), vol=0.035, beta=0.5,
                                     maade="exit", exit=start + int(rng.integers(14, 22)),
                                     **faelles))
            else:
                traek = {start: beloeb * 0.85, start + 2: beloeb * 0.1, start + 5: beloeb * 0.05}
                ud.append(Beholdning(navne.ny("Underliggende fond"), "SECONDARY", traek=traek,
                                     drift=rng.normal(0.10, 0.04), vol=0.02, beta=0.6,
                                     maade="loebende", tilsagn=beloeb * 1.05,
                                     vintage=int(f.foerste_kald.year - rng.integers(4, 9)),
                                     **faelles))
        elif f.art == "kredit":
            ud.append(Beholdning(navne.ny("Selskab"), "LOAN", traek={start: beloeb},
                                 drift=rng.normal(0.0, 0.015), vol=0.01, beta=0.1, maade="rente",
                                 exit=start + int(rng.integers(12, 20)),
                                 rente=rng.uniform(0.085, 0.105),
                                 instrument="Unitranche" if rng.random() < 0.7 else "Second lien",
                                 **faelles))
        else:       # balance: én kunstig beholdning, der kun driver NAV
            traek = {k: beloeb / 10 for k in range(10)}
            ud.append(Beholdning("portefølje", "PROPERTY", traek=traek, drift=0.05, vol=0.015,
                                 beta=0.15, maade="loebende", **faelles))
    return [b for b in ud if min(b.traek) < n_kv]


def _udlodningsrate(alder, art):
    """Kvartalsvis udlodning af dagsværdien for løbende udloddende beholdninger."""
    if art == "secondaries":
        return 0.0 if alder < 3 else min(0.012 + 0.003 * alder, 0.06)
    if art == "balance":
        return 0.012 if alder < 26 else 0.06    # husleje, derefter salg af ejendommene
    return 0.0 if alder < 14 else min(0.02 + 0.004 * (alder - 14), 0.09)


def _simuler(f, beholdninger, datoer, rng_stoej, ekstra_drift):
    """Kør beholdningerne kvartal for kvartal. Giver pr. kvartal fondens træk og udlodninger og
    pr. beholdning (akkumuleret kostpris, udloddet, dagsværdi)."""
    n = len(datoer)
    stoej = {b.navn: rng_stoej.normal(0, 1, n) for b in beholdninger}
    traek, udl = np.zeros(n), np.zeros(n)
    forloeb = {b.navn: [] for b in beholdninger}
    for b in beholdninger:
        fv = kost = udloddet = 0.0
        for t in range(n):
            if t > b.start and fv > 0:
                # værdien skabes mest i de første fire år; modne beholdninger vokser langsommere
                alder = 1.4 if t - b.start <= 16 else 0.5
                r = (b.drift + ekstra_drift) * alder / 4 + b.beta * marked(datoer[t]) \
                    + b.vol * stoej[b.navn][t]
                if t <= b.start + 1 and b.sektion in ("COMPANY", "CO_INVESTMENT", "LOAN"):
                    r = 0.0             # står til kostpris det første halve år
                fv *= float(np.exp(r))
            d = b.traek.get(t, 0.0)
            fv += d
            kost += d
            traek[t] += d
            u = 0.0
            if b.maade == "exit" and t == b.exit:      # exit i to trancher: 60 %, så resten
                u = fv * 0.6
                fv -= u
            elif b.maade == "exit" and t == b.exit + 3:
                u, fv = fv, 0.0
            elif b.maade == "loebende" and fv > 0:
                u = fv * _udlodningsrate(t - b.start, f.art)
                fv -= u
            elif b.maade == "rente" and fv > 0:
                u = kost * b.rente / 4 if t > b.start else 0.0
                if t == b.exit:
                    u, fv = u + fv, 0.0
            udloddet += u
            udl[t] += u
            forloeb[b.navn].append((kost, udloddet, fv))
    return traek, udl, forloeb


def _fondsforloeb(f, fondsstoerrelse, datoer, traek, udl, forloeb):
    """Fondens kvartalstal: kald (træk + gebyrer), udlodninger, dagsværdi, NAV før og efter carry
    og gebyrerne akkumuleret."""
    n = len(datoer)
    fv = np.array([sum(forloeb[k][t][2] for k in forloeb) for t in range(n)])
    # fuldt gebyr af tilsagnet i de første fem år, derefter et faldende gebyr af det investerede
    gebyr = np.array([fondsstoerrelse * f.gebyr / 4 * (1 if t < 20
                                                       else 0.6 * max(0.2, 1 - 0.1 * (t - 20)))
                      for t in range(n)])
    if f.art in ("ejendomme",):
        gebyr[:] = 0.0
    kald = traek + gebyr
    likvider = fv * 0.03
    nav_ex = fv + likvider
    indbetalt, udloddet = np.cumsum(kald), np.cumsum(udl)
    gevinst = nav_ex + udloddet - indbetalt
    # carry hensættes, når fonden er over sin hurdle (1,3x), og vokser gradvist ind
    carry = np.where((nav_ex + udloddet) > 1.3 * indbetalt, 0.2 * np.maximum(gevinst, 0) * 0.8, 0)
    carry = np.minimum(carry, nav_ex * 0.25)
    return dict(kald=kald, udl=udl, fv=fv, nav_ex=nav_ex, nav=nav_ex - carry, carry=carry,
                gebyrer=np.cumsum(gebyr))


def _tvpi(ff, andel, nav_t):
    ind = (ff["kald"][:nav_t + 1]).sum() * andel
    return ((ff["udl"][:nav_t + 1]).sum() * andel + ff["nav"][nav_t] * andel) / ind if ind else 0


@dataclass
class Resultat:
    fond: Fond
    datoer: list
    fondsstoerrelse: float
    beholdninger: list
    forloeb: dict
    ff: dict


def simuler_fond(f, i, navne_tilstand):
    """Simulér én fond og kalibrér driften, så investors TVPI ved NAV-datoen rammer f.tvpi."""
    datoer = kvartaler(f.foerste_kald, DATO)
    nav_t = max(t for t, d in enumerate(datoer) if d <= f.nav_til)
    andel = f.andel or 0.04
    fondsstoerrelse = f.tilsagn / andel
    navne = Navne()
    navne.t = dict(navne_tilstand.t)

    def kør(ekstra):
        rng = np.random.default_rng([SEED, i])
        navne_kopi = Navne()
        navne_kopi.t = dict(navne.t)
        beh = _beholdninger(f, rng, len(datoer), fondsstoerrelse, navne_kopi)
        traek, udl, forl = _simuler(f, beh, datoer, np.random.default_rng([SEED, i, 1]), ekstra)
        ff = _fondsforloeb(f, fondsstoerrelse, datoer, traek, udl, forl)
        return beh, forl, ff, navne_kopi

    lo, hi = -0.4, 0.6
    for _ in range(30):
        mid = (lo + hi) / 2
        _, _, ff, _ = kør(mid)
        lo, hi = (mid, hi) if _tvpi(ff, andel, nav_t) < f.tvpi else (lo, mid)
    beh, forl, ff, navne_kopi = kør((lo + hi) / 2)
    navne_tilstand.t = navne_kopi.t
    return Resultat(f, datoer, fondsstoerrelse, beh, forl, ff)


# ---------------------------------------------------------------------------------------------
# ejendomsfonden med ejendomme (evergreen)

def simuler_ejendomme(f, rng, navne):
    """Ejendomsfonden ejendom for ejendom: købte ejendomme til aftalt pris og derefter ekstern
    vurdering, projekter til omkostningerne og derefter vurdering med udviklingsgevinst.
    Fondens NAV = GAV + selskabsværdi - nettogæld; nye andele udstedes til NAV pr. andel."""
    datoer = kvartaler(f.foerste_kald, DATO)
    n = len(datoer)
    ejendomme = []
    plan = [(0, "koeb", 420), (2, "projekt", 380), (5, "koeb", 310), (8, "projekt", 460),
            (12, "koeb", 280), (16, "projekt", 350), (21, "projekt", 300),
            (24, "projekt", 270), (27, "koeb", 190)]      # et projekt under opførelse, et nyt køb
    for start, art, mio in plan:
        ejendomme.append(dict(navn=navne.ny("Ejendom"), start=start, art=art,
                              pris=mio * 1e6 * rng.uniform(0.9, 1.1),
                              byggetid=int(rng.integers(6, 9)), drift=rng.normal(0.06, 0.01)))
    vaerdier = {e["navn"]: [] for e in ejendomme}      # pr. kvartal: [(grundlag, værdi)]
    for e in ejendomme:
        v = None
        for t in range(n):
            linjer = []
            if t < e["start"]:
                pass
            elif e["art"] == "koeb":
                if v is None:
                    v = e["pris"]
                    linjer = [("property based on agreed purchase price", v)]
                else:
                    v *= float(np.exp(e["drift"] / 4 + 0.6 * _rentechok(datoer[t])
                                      + rng.normal(0, 0.01)))
                    linjer = [("property based on third party valuation", v)]
            else:
                alder = t - e["start"]
                if alder < e["byggetid"]:
                    omk = e["pris"] * (alder + 1) / e["byggetid"]
                    linjer = [("project costs to date", omk)]
                    if alder >= e["byggetid"] - 3:       # gevinsten indregnes mod slutningen
                        linjer.append(("development profit",
                                       e["pris"] * 0.02 * (alder - e["byggetid"] + 4)))
                else:
                    if v is None:
                        v = e["pris"] * 1.08
                    else:
                        v *= float(np.exp(e["drift"] / 4 + 0.6 * _rentechok(datoer[t])
                                          + rng.normal(0, 0.01)))
                    linjer = [("property based on third party valuation", v)]
            vaerdier[e["navn"]].append(linjer)
    gav = np.array([sum(v for linjer in (vaerdier[e["navn"]][t] for e in ejendomme)
                        for _, v in linjer) for t in range(n)])
    # belåningsgraden glider langsomt fra 40 % til 43 %
    gaeld = gav * (np.linspace(0.40, 0.43, n) + rng.normal(0, 0.001, n))
    likvider = gav * 0.02
    pev = gav * 0.015
    nav = gav + pev - gaeld
    # ny egenkapital er 60 % af stigningen i hver ejendoms kostprisdel (køb og
    # projektomkostninger); pr. ejendom, så et færdigt projekt ikke skjuler et nyt køb
    def kostdel(e, t):
        return sum(v for g, v in vaerdier[e["navn"]][t] if g in data.KOSTGRUNDLAG)

    indbetalt = np.zeros(n)
    for t in range(n):
        ny = sum(max(kostdel(e, t) - (kostdel(e, t - 1) if t else 0.0), 0.0) for e in ejendomme)
        indbetalt[t] = (indbetalt[t - 1] if t else 0.0) + (ny * 0.6 if t else 0.0)
    # andele: udstedes til NAV pr. andel primo; investor deltager i de første rejsninger
    andele = np.zeros(n)
    inv_andele = np.zeros(n)
    npu = np.zeros(n)
    inv_kald = np.zeros(n)
    rest = f.tilsagn
    for t in range(n):
        forrige_npu = npu[t - 1] if t else 100.0
        ny_kapital = indbetalt[t] - (indbetalt[t - 1] if t else 0.0)
        forrige_andele = andele[t - 1] if t else 0.0
        if t == 0:
            ny_kapital = nav[0]
        andele[t] = forrige_andele + ny_kapital / forrige_npu
        npu[t] = nav[t] / andele[t]
        kald = 0.0
        if t > 0 and t % 2 == 1 and rest > 0 and t < 22:
            kald = min(rest, f.tilsagn * 0.09)
        rest -= kald
        inv_kald[t] = kald
        inv_andele[t] = (inv_andele[t - 1] if t else 0.0) + kald / forrige_npu
    inv_nav = inv_andele * npu
    return dict(datoer=datoer, ejendomme=ejendomme, vaerdier=vaerdier, gav=gav, gaeld=gaeld,
                likvider=likvider, pev=pev, nav=nav, indbetalt=indbetalt, npu=npu,
                inv_kald=inv_kald, inv_nav=inv_nav)


def _rentechok(q):
    """Ejendomsværdierne faldt med renten i 2022-23."""
    return {(2022, 3): -0.03, (2022, 4): -0.04, (2023, 1): -0.02, (2023, 2): -0.01}.get(
        (q.year, (q.month - 1) // 3 + 1), 0.0)


# ---------------------------------------------------------------------------------------------
# tabellerne

class Tabeller:
    """Rækkerne til hver tabel, med fortløbende id'er."""

    def __init__(self):
        self.r = {}
        self.dok_id = 1000

    def add(self, tabel, **kol):
        self.r.setdefault(tabel, []).append(kol)

    def dokument(self, fund_id, doc_type, period_end=None, value_date=None, note=None):
        self.dok_id += 1
        d = period_end or value_date
        self.add("fund.document", document_id=self.dok_id, fund_id=fund_id, doc_type=doc_type,
                 period_end=period_end, value_date=value_date,
                 notice_date=None if value_date is None else value_date - timedelta(days=10),
                 external_ref=f"DEMO-{self.dok_id}",
                 file_name=f"{doc_type.lower()}_{d:%Y%m%d}.pdf",
                 rel_path=f"demo/{doc_type.lower()}_{d:%Y%m%d}_{self.dok_id}.pdf",
                 page_count=12 if doc_type == "QUARTERLY_REPORT" else 2, status="APPROVED",
                 extraction_method="DEMO", approved_by="demo",
                 approved_at=(d + timedelta(days=45)), note=note)
        return self.dok_id

    def frames(self):
        return {t: pd.DataFrame(r) for t, r in self.r.items()}


KONTI = {"CA110": "Contributions", "CA210": "Management fees", "CA220": "Fund expenses",
         "CA230": "Financing costs", "CA290": "Net result", "CA400": "Carried interest",
         "MM110": "Called capital", "MM120": "Unfunded commitment", "MM150": "Net IRR (%)",
         "MM160": "Net TVPI"}


def byg():
    """{tabelnavn: DataFrame} til demodatabasen og fund_info til demoen."""
    tb = Tabeller()
    kurser = valutakurser(np.random.default_rng([SEED, 99]))
    tb.add("fund.investor", investor_id=1, investor_code=INVESTOR, legal_name=INVESTOR_NAVN)
    for kode, k in KONTI.items():
        tb.add("fund.account", account_id=len(tb.r.get("fund.account", [])) + 1,
               account_code=kode, name=k)
    for v, s in kurser.items():
        for d, k in s.items():
            tb.add("dbo.fx_rates_nationalbanken", currency=v, rate_date=d, dkk_per_1=k)
    forvaltere = {}
    for navn in FORVALTERE.values():
        if navn not in forvaltere:
            forvaltere[navn] = len(forvaltere) + 1
            tb.add("fund.manager", manager_id=forvaltere[navn],
                   manager_code=f"M{forvaltere[navn]}", name=navn)
    for s, r in DELREGION_REGION.items():
        tb.add("portfolio.geo_subregion", subregion_code=s, region_code=r)

    navne = Navne()
    info = {}
    fund_ids = {}
    for i, f in enumerate(FONDE, 1):
        fund_ids[f.kode] = i
    fund_ids[MASTER["kode"]] = len(FONDE) + 1     # masterfonden får et id til sidst
    tb.add("fund.fund", fund_id=fund_ids["PCM1"], fund_code="PCM1", name=MASTER["navn"],
           strategy="Private credit", currency="EUR", vintage_year=2021,
           manager_id=forvaltere[MASTER["forvalter"]])
    tb.add("rpt.v_fund_master", fund_code="PCM1", fund_name=MASTER["navn"],
           manager_name=MASTER["forvalter"], manager_code=f"M{forvaltere[MASTER['forvalter']]}",
           geography_focus="EUROPE", fund_structure=None, strategy_focus="JUNIOR_DEBT",
           size_focus=None, re_sector=None, term_years=7, extension_years=1,
           total_commitments=500_000_000, underlying_fund_code=None)

    for i, f in enumerate(FONDE, 1):
        fid = fund_ids[f.kode]
        tb.add("fund.fund", fund_id=fid, fund_code=f.kode, name=f.navn,
               strategy=data.FOKUS.get(f.fokus), currency=f.valuta, vintage_year=f.vintage,
               manager_id=forvaltere[f.forvalter])
        tb.add("fund.investment", investment_id=100 + i, investor_id=1, fund_id=fid, is_active=1)
        if f.art == "ejendomme":
            sim = simuler_ejendomme(f, np.random.default_rng([SEED, i]), navne)
            _ejendomsfond(tb, f, fid, 100 + i, sim, kurser)
            stoerrelse = None
        else:
            res = simuler_fond(f, i, navne)
            _fond(tb, f, fid, 100 + i, res, kurser, fund_ids)
            stoerrelse = res.fondsstoerrelse
            if f.art == "selskaber":
                info_v = {}
                brugt = {}
                for b in res.beholdninger:
                    tekster = BESKRIVELSER[b.gp_betegnelse]
                    j = brugt.get(b.gp_betegnelse, 0)
                    brugt[b.gp_betegnelse] = j + 1
                    aar = f.foerste_kald.year + (f.foerste_kald.month + 3 * b.start) // 12
                    beskr = f"Investeret i {aar}. " + tekster[j % len(tekster)]
                    v = dict(beskrivelse=beskr)
                    if f.fokus == "BUYOUT":
                        ejer = round(float(np.random.default_rng([SEED, i, j]).uniform(0.45, 0.9)),
                                     2)
                        v.update(ejerandel=ejer, stemmer=round(min(ejer + 0.05, 1.0), 2))
                    info_v[b.navn] = v
                info.setdefault(f.kode, {})["virksomheder"] = info_v
        tb.add("rpt.v_fund_master", fund_code=f.kode, fund_name=f.navn, manager_name=f.forvalter,
               manager_code=f"M{forvaltere[f.forvalter]}", geography_focus=f.geografi,
               fund_structure=f.struktur, strategy_focus=f.fokus, size_focus=None,
               re_sector=f.re_sektor, term_years=f.loebetid[0], extension_years=f.loebetid[1],
               total_commitments=stoerrelse, underlying_fund_code=f.underliggende)
        d = info.setdefault(f.kode, {})
        d.update(status="ok", beskrivelse=f.beskrivelse)
        if f.benchmark:
            d["benchmark"] = f.benchmark
        if f.evergreen:
            d["evergreen"] = True
    return tb.frames(), info


def _cash_flow(tb, iid, fid, f, d, beloeb, klasse, kurser, recallable=0, konto=None):
    # databasens konti: kald og udlodninger uden opdeling, equalization efter fortegnet
    konti = {"CONTRIBUTION": ("CF199", "Capital call - not split"),
             "DISTRIBUTION": (("CF230", "Distribution - recallable") if recallable
                              else ("CF295", "Distribution - not split")),
             "EQUALISATION": (("CF130", "Capital call - equalisation / late-closing interest")
                              if beloeb < 0 else ("CF240", "Distribution - equalisation received")),
             "WITHHOLDING_TAX": ("CF280", "Withholding tax deducted")}
    kode, navn = konto or konti[klasse]
    dok = tb.dokument(fid, "CAPITAL_CALL" if beloeb < 0 else "DISTRIBUTION", value_date=d)
    k = kurs_paa(kurser, f.valuta, d)
    tb.add("rpt.v_cash_flow", investment_id=iid, value_date=d, notice_date=d - timedelta(days=10),
           account_code=kode, account_name=navn, flow_class=klasse, is_recallable=recallable,
           amount=round(beloeb, 2), fx_rate=k, amount_dkk=round(beloeb * k, 2),
           flow_ref=f"{f.kode}-{d:%Y%m%d}-{kode}", document_id=dok)


def _kapitalkonto(tb, iid, q, dok, konti):
    for kode, beloeb in konti.items():
        if beloeb is None:
            continue
        tb.add("fund.capital_account", investment_id=iid, period_end=q, basis="ITD",
               account_id=list(KONTI).index(kode) + 1, amount=round(float(beloeb), 2),
               document_id=dok, is_active=1)


def _exponering(tb, iid, kode, q, linjer, grundlag, fid=None):
    """Investors eksponering pr. sektor/branchegruppe og region/delregion i de to rpt-views."""
    for (sektor, gruppe, delregion), beloeb in linjer.items():
        if beloeb <= 0:
            continue
        tb.add("rpt.v_industry_exposure", investment_id=iid, period_end=q,
               look_through_fund=kode, split_date=q, basis=grundlag, industry_code=sektor,
               industry_group_gics=gruppe, industry_group=GRUPPENAVNE.get(gruppe),
               nav_amount=beloeb)
        tb.add("rpt.v_geography_exposure", investment_id=iid, period_end=q,
               look_through_fund=kode, split_date=q, basis=grundlag,
               region_code=DELREGION_REGION[delregion], subregion_code=delregion,
               subregion=SUBREGION_ENGELSK[delregion], nav_amount=beloeb)


def _fond(tb, f, fid, iid, res, kurser, fund_ids):
    """Rækkerne for en fond, der ikke er en ejendomsfond med ejendomme."""
    ff, datoer, andel = res.ff, res.datoer, f.andel
    nav_t = max(t for t, d in enumerate(datoer) if d <= f.nav_til)
    tb.add("fund.commitment_event", investment_id=iid,
           event_date=f.foerste_kald - timedelta(days=60), event_type="INITIAL",
           amount=f.tilsagn, document_id=tb.dokument(fid, "SUBSCRIPTION",
                                                     value_date=f.foerste_kald - timedelta(days=60)),
           note=None)
    # investors cash flows: én samlet kald og én samlet udlodning pr. kvartal
    flows = []
    kaldt = 0.0
    rabat = f.ekstra.get("rabat", 0.0)
    rng = np.random.default_rng([SEED, fid, 7])
    for t, q in enumerate(datoer):
        kald = ff["kald"][t] * andel
        if kald > 1:
            d = _i_kvartalet(q, 2, int(rng.integers(5, 25)))
            kaldt += kald
            betalt = kald * (1 - rabat) if rabat and t == 0 else kald
            flows.append((d, -betalt, "CONTRIBUTION", 0))
        udl = ff["udl"][t] * andel
        if udl > 1:
            d = _i_kvartalet(q, 3, int(rng.integers(5, 22)))
            # den første udlodning er genkaldelig (et hurtigt exit, der kan kaldes igen)
            genkaldelig = int(bool(f.ekstra.get("genkaldelig"))
                              and not any(k == "DISTRIBUTION" for _, _, k, _ in flows))
            flows.append((d, udl, "DISTRIBUTION", genkaldelig))
            if f.art == "secondaries" and q.year >= 2025:
                flows.append((d, -round(udl * 0.015, 2), "WITHHOLDING_TAX", 0))
    if f.ekstra.get("udligning"):
        flows.append((_i_kvartalet(datoer[2], 2, 10), f.ekstra["udligning"], "EQUALISATION", 0))
    for d, beloeb, klasse, genk in flows:
        _cash_flow(tb, iid, fid, f, d, beloeb, klasse, kurser, recallable=genk)

    # kvartalsrapporterne: NAV, kapitalkonto, nøgletal og beholdninger
    for t, q in enumerate(datoer[:nav_t + 1]):
        dok = tb.dokument(fid, "QUARTERLY_REPORT", period_end=q)
        nav = ff["nav"][t] * andel
        omvurderet = f.ekstra.get("omvurderet") == q
        tb.add("rpt.v_nav", investment_id=iid, period_end=q,
               nav=round(nav * 0.988, 2) if omvurderet else round(nav, 2),
               restated_nav=round(nav, 2) if omvurderet else None,
               restatement="Restated in the following quarter's report" if omvurderet else None,
               document_id=dok)
        betalt = [(d, a) for d, a, k, _ in flows if d <= q and k == "CONTRIBUTION"]
        indbetalt = -sum(a for _, a in betalt)
        kaldt_q = sum(ff["kald"][:t + 1]) * andel
        konti = dict(CA110=indbetalt, CA210=-ff["gebyrer"][t] * andel * 0.8,
                     CA220=-ff["gebyrer"][t] * andel * 0.2,
                     CA230=-ff["gebyrer"][t] * andel * 0.03, CA400=-ff["carry"][t] * andel)
        if t == nav_t:
            genk = sum(a for d, a, k, g in flows if g and d <= q)
            udl = sum(a for d, a, k, _ in flows if d <= q and k == "DISTRIBUTION")
            irr = data.xirr([(d, a) for d, a, k, _ in flows
                             if d <= q and k in ("CONTRIBUTION", "DISTRIBUTION")]
                            + [(q, nav)])
            # en rabat tæller som indbetalt og udloddet, som i rapporten
            rabat_beloeb = kaldt_q - indbetalt if rabat else 0.0
            udligning = -sum(a for d, a, k, _ in flows if d <= q and k == "EQUALISATION")
            ind_i_alt = indbetalt + rabat_beloeb + udligning
            konti.update(MM110=kaldt_q, MM120=f.tilsagn - kaldt_q + genk,
                         MM150=None if irr is None else round(irr * 100, 1),
                         MM160=round((udl + rabat_beloeb + nav) / ind_i_alt, 2))
        _kapitalkonto(tb, iid, q, dok, konti)
        kode = f.underliggende or f.kode
        fonds_nav, fonds_nav_ex = ff["nav"][t], ff["nav_ex"][t]
        if not f.underliggende:
            tb.add("rpt.v_fund_metric", fund_code=f.kode, metric_code="TOTAL_COMMITMENTS",
                   value=res.fondsstoerrelse, report_date=q, document_id=dok)
            if f.art == "balance":
                gav = fonds_nav / 0.58
                for m, v in (("GROSS_ASSET_VALUE", gav), ("TOTAL_ASSETS", gav * 1.05),
                             ("FUND_NAV", fonds_nav)):
                    tb.add("rpt.v_fund_metric", fund_code=f.kode, metric_code=m, value=v,
                           report_date=q, document_id=dok)
                continue
            for m, v in (("FUND_NAV", fonds_nav), ("FUND_NAV_EXCL_CARRY", fonds_nav_ex)):
                tb.add("rpt.v_fund_metric", fund_code=f.kode, metric_code=m, value=v,
                       report_date=q, document_id=dok)
        else:
            tb.add("rpt.v_fund_metric", fund_code=f.kode, metric_code="TOTAL_COMMITMENTS",
                   value=res.fondsstoerrelse * 0.3, report_date=q, document_id=dok)
            # masterfonden: investors andel er feederens del af masteren
            dok_m = tb.dokument(fund_ids[kode], "QUARTERLY_REPORT", period_end=q)
            levende = sum(1 for b in res.beholdninger if res.forloeb[b.navn][t][2] > 0)
            for m, v in (("FUND_NAV", fonds_nav), ("NUMBER_OF_COMPANIES", levende),
                         ("CREDIT_LINE_DRAWN", fonds_nav * 0.12)):
                tb.add("rpt.v_fund_metric", fund_code=kode, metric_code=m, value=v,
                       report_date=q, document_id=dok_m)
            dok = dok_m
        if f.art == "balance":
            continue
        for b in res.beholdninger:
            kost, udl_b, fv = res.forloeb[b.navn][t]
            if kost <= 0:
                continue
            tb.add("rpt.v_portfolio_holding", fund_code=kode, report_date=q,
                   section=b.sektion, name=b.navn, segment=b.gp_betegnelse,
                   trade_date=_i_kvartalet(datoer[b.start], 2, 3), vintage=b.vintage,
                   instrument=b.instrument,
                   commitment=b.tilsagn,
                   invested=round(kost, 2), distributions=round(udl_b, 2),
                   fair_value=round(fv, 2),
                   paid_in_pct=round(100 * kost / b.tilsagn, 1) if b.tilsagn else None,
                   multiple=round((fv + udl_b) / kost, 2), irr_pct=None, document_id=dok,
                   industry_code=b.sektor, industry_group_gics=b.gruppe,
                   industry_suggested=0,
                   region_code=DELREGION_REGION[b.delregion], subregion_code=b.delregion,
                   geography_suggested=0)

    # eksponeringen på NAV-datoen: beholdningernes dagsværdi, eller fondens strategi
    q = datoer[nav_t]
    linjer = {}
    if f.art == "balance":
        for delregion, vaegt in f.regioner.items():
            linjer[("REAL_ESTATE", 6020, delregion)] = vaegt
        for delregion in f.regioner:
            tb.add("portfolio.geo_mapping", fund_id=fid, label=None, subregion_code=delregion,
                   is_suggested=0)
        tb.add("portfolio.industry_mapping", fund_id=fid, label=None, industry_code="REAL_ESTATE",
               industry_group_gics=6020, is_suggested=0)
        _exponering(tb, iid, f.kode, q, linjer, "FUND_DEFAULT")
        return
    for b in res.beholdninger:
        fv = res.forloeb[b.navn][nav_t][2]
        k = (b.sektor, b.gruppe, b.delregion)
        linjer[k] = linjer.get(k, 0.0) + fv
    _exponering(tb, iid, f.underliggende or f.kode, q, linjer, "HOLDINGS")


def _ejendomsfond(tb, f, fid, iid, sim, kurser):
    datoer = sim["datoer"]
    tb.add("fund.commitment_event", investment_id=iid,
           event_date=f.foerste_kald - timedelta(days=45), event_type="INITIAL",
           amount=f.tilsagn * 2 / 3,
           document_id=tb.dokument(fid, "SUBSCRIPTION",
                                   value_date=f.foerste_kald - timedelta(days=45)), note=None)
    tb.add("fund.commitment_event", investment_id=iid, event_date=date(2021, 10, 1),
           event_type="INCREASE", amount=f.tilsagn / 3,
           document_id=tb.dokument(fid, "SUBSCRIPTION", value_date=date(2021, 10, 1)),
           note="Forhøjelse af tilsagnet")
    rng = np.random.default_rng([SEED, fid, 7])
    for t, q in enumerate(datoer):
        if sim["inv_kald"][t] > 0:
            _cash_flow(tb, iid, fid, f, _i_kvartalet(q, 2, int(rng.integers(5, 25))),
                       -sim["inv_kald"][t], "CONTRIBUTION", kurser)
    indbetalt = np.cumsum(sim["inv_kald"])
    for t, q in enumerate(datoer):
        dok = tb.dokument(fid, "QUARTERLY_REPORT", period_end=q)
        nav = sim["inv_nav"][t]
        if nav <= 0:
            continue
        tb.add("rpt.v_nav", investment_id=iid, period_end=q, nav=round(nav, 2),
               restated_nav=None, restatement=None, document_id=dok)
        konti = dict(CA110=indbetalt[t], CA290=nav - indbetalt[t])
        if q == datoer[-1]:
            konti.update(MM110=indbetalt[t], MM120=f.tilsagn - indbetalt[t],
                         MM160=round(nav / indbetalt[t], 2))
        _kapitalkonto(tb, iid, q, dok, konti)
        for m, v in (("GROSS_ASSET_VALUE", sim["gav"][t]), ("NET_DEBT", sim["gaeld"][t]),
                     ("ENTITY_VALUE", sim["pev"][t]), ("CASH", sim["likvider"][t]),
                     ("FUND_NAV", sim["nav"][t]), ("NAV_PER_UNIT", sim["npu"][t]),
                     ("PAID_IN", sim["indbetalt"][t] + sim["nav"][0]),
                     ("TOTAL_ASSETS", sim["gav"][t] + sim["likvider"][t] + sim["pev"][t])):
            tb.add("rpt.v_fund_metric", fund_code=f.kode, metric_code=m, value=round(v, 2),
                   report_date=q, document_id=dok)
        for e in sim["ejendomme"]:
            for grundlag, v in sim["vaerdier"][e["navn"]][t]:
                tb.add("rpt.v_portfolio_holding", fund_code=f.kode, report_date=q,
                       section="PROPERTY", name=e["navn"], segment=grundlag, trade_date=None,
                       vintage=None, instrument=None, commitment=None, invested=None,
                       distributions=None, fair_value=round(v, 2), paid_in_pct=None,
                       multiple=None, irr_pct=None, document_id=dok, industry_code="REAL_ESTATE",
                       industry_group_gics=6020, industry_suggested=0, region_code="EUROPE", subregion_code="DENMARK",
                       geography_suggested=0)
    q = datoer[-1]
    _exponering(tb, iid, f.kode, q, {("REAL_ESTATE", 6020, "DENMARK"): sim["inv_nav"][-1]},
                "HOLDINGS")


@lru_cache(maxsize=1)
def demodata():
    """(tabeller, fund_info), bygget én gang pr. proces."""
    return byg()


# ---------------------------------------------------------------------------------------------
# kontrol af, at porteføljen stadig ligner en dansk fonds

# segmentet efter tilsagn i DKK: mål og tolerance
SEGMENTMAAL = {"Private equity": 0.45, "Fund of funds": 0.10, "Secondaries": 0.12,
               "Private credit": 0.15, "Ejendomme": 0.18}
SEGMENT_TOLERANCE = 0.05
# regionerne efter NAV (gennemlyst): (mindst, højst)
REGIONSGRAENSER = {"Europa": (0.55, 0.80), "Nordamerika": (0.15, 0.40),
                   "Resten af verden": (0.02, 0.12)}


def kontrol(r):
    """Afvigelser fra en realistisk fordeling i den færdige rapport-dict (tom liste: i orden).

    Fanger, at en ændring af seed eller parametre gør porteføljen urealistisk: strategierne
    efter tilsagn, valutaerne og regionerne efter NAV, sektorerne gennemlyst og fondenes TVPI.
    """
    ud = []
    segmenter = {g["navn"]: g["andel"] for g in r["fordelinger"]["segment"]["grupper"]}
    for navn, maal in SEGMENTMAAL.items():
        if abs(segmenter.get(navn, 0.0) - maal) > SEGMENT_TOLERANCE:
            ud.append(f"segment {navn}: {segmenter.get(navn, 0.0):.1%} af tilsagnet, mål {maal:.0%}")
    valutaer = {x["valuta"]: x["andel"] for x in r["valutafordeling"]["nav"]}
    if set(valutaer) != {"DKK", "EUR", "USD"} or min(valutaer.values()) < 0.15:
        ud.append(f"valutaerne efter NAV: {valutaer}")
    if max(valutaer, key=valutaer.get) != "EUR":
        ud.append("EUR er ikke den største valuta")
    regioner = {g["navn"]: g["andel"] for g in r["regioner"]["grupper"]}
    for navn, (lav, hoej) in REGIONSGRAENSER.items():
        if not lav <= regioner.get(navn, 0.0) <= hoej:
            ud.append(f"region {navn}: {regioner.get(navn, 0.0):.1%} af NAV, ventet "
                      f"{lav:.0%}-{hoej:.0%}")
    sektorer = {g["navn"]: g["andel"] for g in r["brancher"]["grupper"]}
    if len(sektorer) < 9:
        ud.append(f"kun {len(sektorer)} sektorer")
    uden_ejd = {k: v for k, v in sektorer.items() if k != "Real Estate"}
    if max(uden_ejd, key=uden_ejd.get) != "Information Technology":
        ud.append("Information Technology er ikke den største sektor uden for ejendomme")
    if not 0.12 <= sektorer.get("Real Estate", 0.0) <= 0.30:
        ud.append(f"Real Estate: {sektorer.get('Real Estate', 0.0):.1%} af NAV")
    for f in r["fonde"]:
        spec = next((x for x in FONDE if x.kode == f["fund_code"]), None)
        if spec and spec.art != "ejendomme" and abs(f["n"]["tvpi"] - spec.tvpi) > 0.03:
            ud.append(f"{f['fund_name']}: TVPI {f['n']['tvpi']:.2f}, mål {spec.tvpi:.2f}")
    return ud
