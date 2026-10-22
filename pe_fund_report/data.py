"""Henter og beregner alt, hvad investorrapporten viser, for én investor.

Reglerne følger docs/analysis-data-guide.md: rpt-views først, LP-fortegn på cash flows
(betalt < 0), betalt udligningsrente (EQUALISATION) tæller med i indbetalt og modtaget i
udloddet, men ikke i kaldt kapital og resttilsagn; en rabat på kald tæller som indbetalt og
udloddet (som i Aleta); WITHHOLDING_TAX holdes ude af IRR,
og basistabeller filtreres på APPROVED + is_active = 1. Investor- og fondsniveau blandes
ikke: beholdningerne er hele fondens tal og skaleres med investors ejerandel.
"""
from datetime import date, timedelta

import pandas as pd
import yaml

from .db import read_sql
from .paths import BENCHMARK, FUND_INFO

GP_TOLERANCE = 0.005      # afvigelse mellem cash flows og GP'ens indbetalte kapital, før GP-tallet bruges


class Note(str):
    """En note om ét bestemt tal på fondssiden: felt er nøglen i n (indbetalt, irr, andel,
    resttilsagn, nav) eller udvikling/beskrivelse. Sliden sætter fodnotemærket ved det tal;
    som str virker noten alle andre steder som før (Excel, tests)."""

    def __new__(cls, tekst, felt=None):
        note = super().__new__(cls, tekst)
        note.felt = felt
        return note

# Geografi og segment kommer fra rpt.v_fund_master som koder; databasens check constraints
# på fund.fund lukker kodelisterne, så hver kode har præcis én dansk tekst her. Den
# geografiske fordeling grupperer på teksten, så to stavemåder ville blive to grupper.
GEOGRAFIER = {"DENMARK": "Danmark", "NORDICS": "Norden", "EUROPE": "Europa",
              "NORTH_AMERICA": "Nordamerika", "EUROPE_NORTH_AMERICA": "Europa og Nordamerika",
              "GLOBAL": "Global"}
STRUKTURER = {"FUND_OF_FUNDS": "Fund of funds", "SECONDARY": "Secondaries"}
# buyout, venture og growth står samlet som Private equity
FOKUS = {"BUYOUT": "Private equity", "VENTURE": "Private equity", "GROWTH": "Private equity",
         "JUNIOR_DEBT": "Private credit", "SHIPPING": "Shipping",
         "MULTI_STRATEGY": "Multistrategi", "REAL_ESTATE": "Ejendomme"}
# rækkefølgen, fondene og segmenterne vises i; andre segmenter kommer sidst
STRATEGI_ORDEN = ["Private equity", "Fund of funds", "Secondaries", "Private credit", "Ejendomme"]


def strategi_noegle(f):
    """Sorteringsnøgle for fondene: efter strategi i STRATEGI_ORDEN, så efter navn."""
    return _strategi_plads(f.get("segment")), f["fund_name"]


def _strategi_plads(segment):
    return STRATEGI_ORDEN.index(segment) if segment in STRATEGI_ORDEN else len(STRATEGI_ORDEN)


SELSKABSFORMER = (" K/S", " A/S", " P/S", ", L.P.", " L.P.", ", SCSp", " SCSp")


def uden_selskabsform(navn):
    """Fondsnavnet uden selskabsform, som decket viser det: "Home.Earth K/S" -> "Home.Earth"."""
    for suffiks in SELSKABSFORMER:
        navn = navn.replace(suffiks, "")
    return navn


def _kode(v):
    return None if v is None or pd.isna(v) else str(v)


def geografi(geography_focus):
    """Dansk tekst for geography_focus. En ukendt kode vises som koden, None som None."""
    kode = _kode(geography_focus)
    return GEOGRAFIER.get(kode, kode)


def loebetid(term_years, extension_years):
    """Løbetiden som GP'erne skriver den: "10+1+1" for 10 år og to års forlængelse.

    Databasen har kun det samlede antal forlængelsesår; de skrives som etårige
    forlængelser, som i fondsvilkårene for vores fonde. None, når løbetiden ikke kendes.
    """
    if term_years is None or pd.isna(term_years):
        return None
    ekstra = 0 if extension_years is None or pd.isna(extension_years) else int(extension_years)
    return str(int(term_years)) + "+1" * ekstra


def segment(fund_structure, strategy_focus, size_focus):
    """Segmentet, som Type-kolonnen og segmentfordelingen viser det.

    Private credit er sin egen aktivklasse og står som "Private credit", også når fonden er
    en fund of funds (Private Debt III LGT); ellers beskrives fund of funds og secondaries ved
    strukturen og de direkte fonde ved fokus, hvor buyout, venture og growth er Private
    equity. size_focus indgår ikke længere.
    """
    struktur, fokus = _kode(fund_structure), _kode(strategy_focus)
    if fokus == "JUNIOR_DEBT":
        return FOKUS[fokus]
    if struktur in STRUKTURER:
        return STRUKTURER[struktur]
    if fokus is None:
        return None
    return FOKUS.get(fokus, fokus)


def _f(v):
    """Decimal/NaN -> float eller None."""
    return None if v is None or pd.isna(v) else float(v)


def _dato(v):
    return pd.Timestamp(v).date()


def xirr(flows):
    """flows: [(dato, beløb)], act/365, bisektion. None hvis der ikke er fortegnsskifte."""
    flows = [(d, a) for d, a in flows if a]
    if not flows or min(a for _, a in flows) >= 0 or max(a for _, a in flows) <= 0:
        return None
    t0 = min(d for d, _ in flows)
    f = lambda r: sum(a / (1 + r) ** ((d - t0).days / 365) for d, a in flows)
    lo, hi = -0.99, 10.0
    if f(lo) * f(hi) > 0:
        return None
    for _ in range(200):
        mid = (lo + hi) / 2
        lo, hi = (mid, hi) if f(lo) * f(mid) > 0 else (lo, mid)
    return (lo + hi) / 2


def fund_info():
    with open(FUND_INFO, encoding="utf-8") as fh:
        return yaml.safe_load(fh) or {}


def benchmark(sti=BENCHMARK):
    """MSCI-indeksenes vægte: {dato, indeks: {navn: {sektorkode: andel}},
    regioner: {navn: {regionskode: andel}}}, eller None.

    Andelene er brøker (33,21 % -> 0,3321). Uden filen vises branche- og geografisiderne uden
    benchmark; uden afsnittet regioner kun geografisiderne.
    """
    try:
        with open(sti, encoding="utf-8") as fh:
            b = yaml.safe_load(fh) or {}
    except FileNotFoundError:
        return None

    def vaegte(afsnit):
        return {navn: {kode: float(v) / 100 for kode, v in (vgt or {}).items()}
                for navn, vgt in (b.get(afsnit) or {}).items()}

    indeks, regioner = vaegte("indeks"), vaegte("regioner")
    if not indeks and not regioner:
        return None
    return dict(dato=_dato(b["dato"]) if b.get("dato") else None, indeks=indeks,
                regioner=regioner)


def investor_navn(investor_code):
    df = read_sql("SELECT legal_name FROM fund.investor WHERE investor_code = :c", c=investor_code)
    if df.empty:
        raise ValueError(f"Investoren '{investor_code}' findes ikke i fund.investor")
    return df.legal_name[0]


def _investeringer(investor_code):
    return read_sql("""
        SELECT x.investment_id, i.investor_code, f.fund_code, f.name AS fund_name, f.strategy,
               f.currency,
               f.vintage_year,
               -- vores feeder-selskaber står med os selv som forvalter; vis den forvalter,
               -- der forvalter den underliggende fond, som pengene reelt ligger i
               COALESCE(u.manager_name, m.name) AS manager_name,
               COALESCE(u.manager_code, m.manager_code) AS manager_code,
               fm.geography_focus, fm.fund_structure, fm.strategy_focus, fm.size_focus,
               fm.re_sector,
               fm.term_years, fm.extension_years, fm.total_commitments,
               fm.underlying_fund_code, u.fund_name AS underlying_fund_name
        FROM fund.investment x
        JOIN fund.investor i ON i.investor_id = x.investor_id
        JOIN fund.fund f ON f.fund_id = x.fund_id
        JOIN fund.manager m ON m.manager_id = f.manager_id
        LEFT JOIN rpt.v_fund_master fm ON fm.fund_code = f.fund_code
        LEFT JOIN rpt.v_fund_master u ON u.fund_code = fm.underlying_fund_code
        WHERE i.investor_code = :c AND x.is_active = 1
        ORDER BY f.name""", c=investor_code)


def _logo(fund_code):
    """Fondens logo fra rpt.v_fund_logo; helst udgaven til lys baggrund. None, hvis intet.

    Viewet giver fondens eget logo, hvor fonden har et (PD3LGT: LGT, selv om forvalteren er
    AEP), og ellers forvalterens.
    """
    df = read_sql("""
        SELECT background, width_px, height_px, image FROM rpt.v_fund_logo
        WHERE fund_code = :f
        ORDER BY CASE background WHEN 'LIGHT' THEN 0 ELSE 1 END""", f=fund_code)
    if df.empty:
        return None
    r = df.iloc[0]
    return dict(billede=bytes(r.image), bredde=int(r.width_px), hoejde=int(r.height_px),
                moerk_baggrund=r.background == "DARK")


def _kapitalkonto(investment_id, as_of):
    return read_sql("""
        SELECT ca.period_end, ca.basis, a.account_code, a.name AS account_name, ca.amount,
               ca.document_id
        FROM fund.capital_account ca
        JOIN fund.document d ON d.document_id = ca.document_id AND d.status = 'APPROVED'
        JOIN fund.account a ON a.account_id = ca.account_id
        WHERE ca.is_active = 1 AND ca.investment_id = :i AND ca.period_end <= :d""",
        i=investment_id, d=as_of)


def _gp_indbetalt(ka, nav_dato):
    """GP'ens akkumulerede indbetalinger: CA110 ITD på NAV-datoen, ellers None.

    Kun et tal pr. NAV-datoen: et ældre tal mangler de senere kald og ville trække den
    indbetalte kapital ned. MM110 bruges ikke: det er kaldt kapital, som ved rabat på kald
    (Petersen & Partners' feedere) er mere end det, investor har betalt.
    """
    itd = ka[(ka.account_code == "CA110") & (ka.basis == "ITD") & (ka.period_end == nav_dato)]
    return _f(itd.amount.iloc[0]) if not itd.empty else None


def gp_vinder(cf_indbetalt, gp_indbetalt, genkaldelige):
    """Skal forvalterens indbetalte kapital bruges frem for summen af kapitalkaldene?

    Kun når de afviger med mere end GP_TOLERANCE, og afvigelsen ikke forklares af de
    genkaldelige udlodninger: nogle forvaltere (MIF3) trækker dem fra i CA110, mens de her
    tæller som udlodninger, så GP-tallet ville trække dem fra to gange.
    """
    if not gp_indbetalt:
        return False
    graense = GP_TOLERANCE * gp_indbetalt
    return (abs(gp_indbetalt - cf_indbetalt) > graense
            and abs(gp_indbetalt + genkaldelige - cf_indbetalt) > graense)


def estimerede_cashflows(info, investor_code, cf, nav_dato):
    """(skønnede cash flows til og med nav_dato, advarsler) for én investors andel i fonden.

    Skønnene står i fund_info.yaml under fonden, hver med den investor, kaldet tilhører
    (cash flows hører til investor × fond). Et skøn, hvis beløb nu findes blandt de
    registrerede cash flows, er afløst af den rigtige meddelelse og udelades, så kaldet
    ikke tælles to gange; advarslen beder om at slette det fra filen.
    """
    est, advarsler = [], []
    for e in info.get("estimerede_cashflows", []):
        if "investor" not in e:
            raise ValueError(f"estimerede_cashflows mangler investor: {e}")
        if e["investor"] != investor_code or _dato(e["dato"]) > nav_dato:
            continue
        if (abs(cf.amount - float(e["beloeb"])) < 1).any():
            advarsler.append(f"det skønnede cash flow ({e['note']}) er nu registreret og "
                             "medregnes ikke igen - slet det fra fund_info.yaml")
            continue
        est.append(e)
    return est, advarsler


def kaldt(indbetalt, gp_kaldt):
    """Kapital kaldt mod tilsagnet: forvalterens MM110, når den overstiger det indbetalte.

    Har forvalteren givet rabat på et kald (Petersen & Partners' feedere: 25 % af første
    kald), er der kaldt mere af tilsagnet, end investor har betalt. Rabatten mindsker
    resttilsagnet; se rabat() for, hvordan den tæller i indbetalt og udloddet. Små afvigelser
    (afrunding) ignoreres.
    """
    if gp_kaldt and gp_kaldt - indbetalt > GP_TOLERANCE * gp_kaldt:
        return gp_kaldt
    return indbetalt


def rabat(betalt, kaldt_, cf):
    """(rabat, dato) for en rabat på kald, eller (0.0, None).

    Rabatten er det kaldte mod tilsagnet ud over det betalte. Den tæller som indbetalt kapital
    og som udlodning (brugerens valg, som i Aleta): investor har fået kaldet fuldt, og
    forvalteren har betalt rabatten tilbage. Datoen er det første kapitalkald, som rabatten
    blev givet på; da begge ben ligger samme dag, ændrer den ikke IRR.
    """
    beloeb = kaldt_ - betalt
    kald = cf[cf.flow_class == "CONTRIBUTION"]
    if beloeb <= 0 or kald.empty:
        return 0.0, None
    return float(beloeb), min(kald.value_date)


TVPI_TOLERANCE = 0.015    # forvalterne oplyser TVPI med to decimaler


def afstem_mod_forvalteren(n, gp_resttilsagn, valuta):
    """Advarsler, hvor resttilsagn eller TVPI afviger fra forvalterens eget tal (MM120, MM160).

    Kun til bygningen: tallene i decket regnes af cash flows og NAV, men en afvigelse
    betyder typisk et manglende kald, en manglende udlodning eller en udløbet genkaldelig
    udlodning, som skal ses efter før udsendelse.
    """
    ud = []
    if (gp_resttilsagn is not None and n["resttilsagn"] is not None and n["tilsagn"]
            and abs(n["resttilsagn"] - gp_resttilsagn) > GP_TOLERANCE * n["tilsagn"]):
        ud.append(f"resttilsagn {n['resttilsagn']:,.0f} {valuta}, forvalteren (MM120) "
                  f"{gp_resttilsagn:,.0f} {valuta}")
    if (n["gp_tvpi"] is not None and n["tvpi"] is not None
            and abs(n["tvpi"] - n["gp_tvpi"]) > TVPI_TOLERANCE):
        ud.append(f"TVPI {n['tvpi']:.3f}, forvalteren (MM160) {n['gp_tvpi']:.2f}")
    return ud


def _memo(ka, nav_dato, konto):
    r = ka[(ka.account_code == konto) & (ka.period_end == nav_dato)]
    return _f(r.amount.iloc[0]) if not r.empty else None


def _vaerdibro(ka, nav_dato, n):
    """Fra indbetalt til NAV. Værdiskabelsen er residualen, så broen altid stemmer."""
    nu = ka[ka.period_end == nav_dato]
    itd = nu[nu.basis == "ITD"].set_index("account_code").amount
    if "CA290" in itd.index and not any(k in itd.index for k in ("CA210", "CA220", "CA230",
                                                                   "CA400")):
        # resultatet er ikke opdelt (Home.Earth): gebyrer og carry er ikke oplyst, og nuller
        # ville ligne oplyste tal
        return [("Indbetalt kapital", n["indbetalt"]), ("Udlodninger", 0.0 - n["udloddet"]),
                ("Værdiskabelse efter omkostninger og carry",
                 n["nav"] - n["indbetalt"] + n["udloddet"]),
                ("NAV", n["nav"])]
    if not itd.empty:
        gebyrer = sum(_f(itd.get(k)) or 0 for k in ("CA210", "CA220", "CA230"))
        carry = _f(itd.get("CA400")) or 0
    else:
        # kun YTD (CataCap): carry summeres over årsultimoerne + indeværende år; gebyrer er ikke udskilt
        ytd = ka[(ka.basis == "YTD") & (ka.account_code == "CA400")]
        aar = ytd[(pd.to_datetime(ytd.period_end).dt.month == 12) | (ytd.period_end == nav_dato)]
        gebyrer, carry = None, float(aar.amount.sum())
    vaerdi = n["nav"] - n["indbetalt"] + n["udloddet"] - (gebyrer or 0) - carry
    bro = [("Indbetalt kapital", n["indbetalt"]),
           ("Udlodninger", 0.0 - n["udloddet"]),
           ("Værdiskabelse før carry" if gebyrer is not None
            else "Værdiskabelse efter omkostninger, før carry", vaerdi)]
    if gebyrer is not None:
        bro.append(("Forvaltningsgebyr og omkostninger", gebyrer))
    bro += [("Carried interest", carry), ("NAV", n["nav"])]
    return bro


# Afsluttende ord i Park Squares tranchenavne ("STADA Common", "Pavion Junior Pref"): de
# beskriver instrumentet, ikke selskabet, og skifter stavemåde fra kvartal til kvartal
INSTRUMENTORD = {"common", "pik", "pref", "preferred", "preffered", "junior", "senior", "second",
                 "delayed", "draw", "equity", "lien", "jr", "sr"}


def _selskab(navn):
    """Selskabet bag et tranchenavn: "Pavion Junior Pref" -> "Pavion". Første ord bliver altid."""
    ord_ = navn.split()
    while len(ord_) > 1 and ord_[-1].lower() in INSTRUMENTORD:
        ord_.pop()
    return " ".join(ord_)


def _saml(linjer):
    """Læg linjer med samme nøgle sammen: beløb summeres, TVPI regnes igen af beløbene."""
    if len(linjer) == 1:
        return dict(linjer[0])
    stoerst = max(linjer, key=lambda x: x["dagsvaerdi_fond"])
    s = lambda k: sum(x[k] or 0 for x in linjer)
    kost, fv, udl = s("kostpris_fond"), s("dagsvaerdi_fond"), s("udloddet_fond")
    datoer = [x["closing"] for x in linjer if x["closing"]]
    vintager = [x["vintage"] for x in linjer if x["vintage"]]
    instrumenter = list(dict.fromkeys(x["instrument"] for x in linjer if x["instrument"]))
    return dict(stoerst, kostpris_fond=kost, dagsvaerdi_fond=fv, udloddet_fond=udl,
                closing=min(datoer) if datoer else None, vintage=min(vintager) if vintager else None,
                instrument=", ".join(instrumenter) or None,
                tvpi=(fv + udl) / kost if kost else None, irr=None)


# Ejendomsfondenes værdigrundlag pr. linje (segment i rpt.v_portfolio_holding), på dansk
VAERDIGRUNDLAG = {
    "property based on third party valuation": "ekstern vurdering",
    "building rights based on third party valuation": "byggeret, ekstern vurdering",
    "property based on agreed purchase price": "aftalt købspris",
    "project costs to date": "projektomkostninger",
    "acquisition costs to date": "anskaffelsesomkostninger",
    "development profit": "udviklingsgevinst",
    "development fee": "udviklingshonorar",
    "assets measured at cost": "kostpris",
    "valued at cost": "kostpris",
    "at cost": "kostpris",
}


def _ejendomme(fund_code, h):
    """En ejendomsfonds ejendomme som hele fondens værdier, før gæld.

    Ejendommene summerer til fondens bruttoværdi (GAV), mens investors andel er NAV efter
    gæld; at gange dem med ejerandelen ville give mere end investors NAV, så de skaleres
    ikke (skaleret = False) og vises som fordelingen af fondens ejendomsværdi. Linjer for
    samme ejendom (Nærheden: ekstern vurdering + kostpris) lægges sammen, og
    værdigrundlagene skrives efter hinanden, det største først.
    """
    dato = _dato(h.report_date.iloc[0])
    grupper = {}
    for r in h.itertuples():
        grupper.setdefault(r.name, []).append(r)
    raekker = []
    for navn, linjer in grupper.items():
        linjer.sort(key=lambda r: -(_f(r.fair_value) or 0))
        vaerdi = sum(_f(r.fair_value) or 0.0 for r in linjer)
        if not vaerdi:
            continue
        grundlag = list(dict.fromkeys(
            VAERDIGRUNDLAG.get(str(r.segment).strip().lower(), str(r.segment).strip().lower())
            for r in linjer if r.segment))
        tekst = " + ".join(grundlag)
        raekker.append(dict(
            navn=navn, sektor=tekst[:1].upper() + tekst[1:] or None, vintage=None, closing=None,
            instrument=None, kostpris_fond=None, dagsvaerdi_fond=vaerdi, udloddet_fond=0.0,
            kostpris=None, dagsvaerdi=vaerdi, udloddet=0.0, total=vaerdi, tvpi=None, irr=None,
            type="PROPERTY", document_id=None if pd.isna(linjer[0].document_id)
            else int(linjer[0].document_id)))
    raekker.sort(key=lambda x: -x["dagsvaerdi"])
    m = read_sql("""
        SELECT metric_code, value FROM rpt.v_fund_metric
        WHERE fund_code = :f AND report_date = :d
          AND metric_code IN ('GROSS_ASSET_VALUE', 'TOTAL_ASSETS', 'FUND_NAV')""",
                 f=fund_code, d=dato)
    return dict(dato=dato, raekker=raekker, skaleret=False, andel=1.0, kode=fund_code,
                type="ejendomme", fakta={r.metric_code: _f(r.value) for r in m.itertuples()})


def _ejendomsbalance(inv, dato):
    """En ejendomsfond uden ejendomme pr. linje: kun fondens GAV, aktiver i alt og NAV.

    CRED oplyser ejendommenes samlede værdi, men ikke værdien pr. ejendom; waterfallen på
    fondssiden går så fra GAV til NAV uden opdeling. None for andre fonde, eller når GAV eller
    fondens NAV mangler på datoen.
    """
    if _kode(inv.strategy_focus) != "REAL_ESTATE":
        return None
    m = read_sql("""
        SELECT metric_code, value FROM rpt.v_fund_metric
        WHERE fund_code = :f AND report_date = :d
          AND metric_code IN ('GROSS_ASSET_VALUE', 'TOTAL_ASSETS', 'FUND_NAV')""",
                 f=inv.fund_code, d=dato)
    fakta = {r.metric_code: _f(r.value) for r in m.itertuples()}
    if not fakta.get("GROSS_ASSET_VALUE") or not fakta.get("FUND_NAV"):
        return None
    return dict(dato=dato, raekker=[], skaleret=False, andel=1.0, kode=inv.fund_code,
                type="ejendomme", fakta=fakta)


def _beholdninger(fund_code, as_of, andel, nav=None):
    """Fondens beholdninger på seneste rapportdato <= as_of.

    Skaleres til investors andel, når ejerandelen kendes; ellers vises hele fondens tal
    (skaleret = False), så siden ikke står tom, blot fordi fondsstørrelsen mangler. Med nav
    (gennemlysning af en feeder) er andelen nav / fondens samlede dagsværdi, så investors
    NAV fordeles på positionerne efter dagsværdi. Linjer med samme navn lægges sammen, og i
    tranchetabeller (instrument udfyldt) samles trancherne pr. selskab.
    """
    h = read_sql("""
        SELECT h.report_date, h.section, h.name, h.segment, h.trade_date, h.vintage,
               h.instrument, h.commitment, h.invested, h.distributions,
               h.fair_value, h.paid_in_pct, h.multiple, h.irr_pct, h.document_id,
               h.industry_code, h.industry_suggested
        FROM rpt.v_portfolio_holding h
        WHERE h.fund_code = :f AND h.report_date =
              (SELECT MAX(report_date) FROM rpt.v_portfolio_holding
               WHERE fund_code = :f AND report_date <= :d)""", f=fund_code, d=as_of)
    if h.empty:
        return None
    if (h.section == "PROPERTY").all():
        return _ejendomme(fund_code, h)
    tranchevis = h.instrument.notna().any()
    grupper = {}
    for r in h.itertuples():
        kost = _f(r.invested)
        if kost is None and _f(r.commitment) is not None and _f(r.paid_in_pct) is not None:
            kost = _f(r.commitment) * _f(r.paid_in_pct) / 100
        navn = _selskab(r.name) if tranchevis else r.name
        # *_fond er hele fondens tal; virksomhedsprofilerne viser dem, som GP'en rapporterer dem,
        # mens beholdningstabellen viser investors andel.
        grupper.setdefault(navn.lower(), []).append(dict(
            navn=navn, vintage=None if pd.isna(r.vintage) else int(r.vintage),
            sektor=r.segment, closing=None if pd.isna(r.trade_date) else _dato(r.trade_date),
            # GICS-sektoren (engelsk navn); sektor ovenfor er forvalterens egen betegnelse
            industri=SEKTORER.get(_kode(r.industry_code)),
            industri_foreloebig=bool(r.industry_suggested) and not pd.isna(r.industry_suggested),
            instrument=None if pd.isna(r.instrument) else r.instrument,
            kostpris_fond=kost, dagsvaerdi_fond=_f(r.fair_value) or 0.0,
            udloddet_fond=_f(r.distributions) or 0.0, tvpi=_f(r.multiple), irr=_f(r.irr_pct),
            type=r.section, document_id=None if pd.isna(r.document_id) else int(r.document_id)))
    rows = [_saml(g) for g in grupper.values()]
    if nav is not None:
        samlet = sum(x["dagsvaerdi_fond"] for x in rows)
        andel = nav / samlet if samlet else None
    skaleret, andel = andel is not None, andel or 1.0
    for x in rows:
        kost, fv, udl = x["kostpris_fond"], x["dagsvaerdi_fond"], x["udloddet_fond"]
        x.update(kostpris=None if kost is None else kost * andel, dagsvaerdi=fv * andel,
                 udloddet=udl * andel, total=(fv + udl) * andel)
    rows.sort(key=lambda x: -x["total"])
    # underliggende fonde: ATP/KAPE (section FUND) og LP-andele hos en secondaryfond (NEP);
    # resten er porteføljeselskaber
    fondsandel = sum(x["dagsvaerdi_fond"] for x in rows
                     if x["type"] in ("FUND", "SECONDARY", "PRIMARY"))
    samlet = sum(x["dagsvaerdi_fond"] for x in rows)
    return dict(dato=_dato(h.report_date.iloc[0]), raekker=rows, skaleret=skaleret, andel=andel,
                kode=fund_code,
                type="fonde" if (h.section == "FUND").all() or (samlet and fondsandel > samlet / 2)
                else "selskaber")


def perioder(rapport_dato):
    """{"kvartal": (fra, til), "aar": (fra, til)}: kvartalet og året til og med rapportdatoen.

    fra er ultimo af forrige kvartal og forrige år; i første kvartal er de to perioder ens.
    """
    kvartal_start = date(rapport_dato.year, 3 * ((rapport_dato.month - 1) // 3) + 1, 1)
    return dict(kvartal=(kvartal_start - timedelta(days=1), rapport_dato),
                aar=(date(rapport_dato.year - 1, 12, 31), rapport_dato))


PERIODENAVNE = {"kvartal": "kvartalet", "aar": "år til dato"}


def _periodeaendring(historik, flows, fra, til):
    """Fra NAV primo til NAV ultimo for én fond i perioden (fra, til].

    historik: fondens kvartalshistorik (dato, nav, indbetalt, udloddet); flows: [(dato,
    beløb, klasse)] med LP-fortegn (betalt < 0), klasse "kald" eller "udlodning". Primo er
    seneste NAV på eller før fra; mangler den, bruges den seneste før, og cash flows tælles
    fra den dato, så broen stadig går op. Uden NAV før perioden er primo 0. Ultimo er
    fondens seneste NAV. Værdiændringen er residualen, og afkastet er Modified Dietz:
    værdiændringen over primo plus de dagvægtede cash flows.
    """
    foer = [h for h in historik if h["dato"] <= fra]
    h_primo, h_ultimo = (foer[-1] if foer else None), historik[-1]
    primo_dato, primo = (h_primo["dato"], h_primo["nav"]) if h_primo else (fra, 0.0)
    slut = h_ultimo["dato"]
    i_perioden = [(d, -a, k) for d, a, k in flows if primo_dato < d <= slut]  # ind i fonden > 0
    kald = sum(a for _, a, k in i_perioden if k == "kald")
    udlodning = -sum(a for _, a, k in i_perioden if k == "udlodning")
    ultimo = h_ultimo["nav"]
    vaerdi = ultimo - primo - kald + udlodning
    dage = (slut - primo_dato).days
    naevner = primo + (sum(a * (slut - d).days / dage for d, a, _ in i_perioden) if dage else 0)
    tvpi = lambda h: (h["udloddet"] + h["nav"]) / h["indbetalt"] if h and h["indbetalt"] else None
    t_primo, t_ultimo = tvpi(h_primo), tvpi(h_ultimo)
    return dict(fra=fra, til=til, primo_dato=primo_dato, primo=primo, kald=kald,
                udlodning=udlodning, vaerdi=vaerdi, ultimo=ultimo,
                afkast=vaerdi / naevner if naevner > 0 else None, naevner=naevner,
                tvpi_aendring=(t_ultimo - t_primo if t_primo is not None and t_ultimo is not None
                               else None),
                flows=i_perioden)


def udvikling(f, rapport_perioder):
    """Sæt f["udvikling"] = {periode: _periodeaendring(...)} og noter om afvigende primo."""
    f["udvikling"] = {p: _periodeaendring(f["historik"], f["kapitalstroemme"], fra, til)
                      for p, (fra, til) in rapport_perioder.items()}
    for p, u in f["udvikling"].items():
        if u["primo"] and u["primo_dato"] < u["fra"]:
            f["noter"].append(Note(f"Udviklingen for {PERIODENAVNE[p]} regnes fra NAV pr. "
                                   f"{u['primo_dato']:%d.%m.%Y}, da forvalteren ikke har "
                                   f"opgjort NAV pr. {u['fra']:%d.%m.%Y}.", "udvikling"))


# Park Squares kategorier, som de står i v_portfolio_breakdown
KATEGORIER = {"United Kingdom": "Storbritannien", "United States": "USA", "France": "Frankrig",
              "Nordics": "Norden", "Other": "Øvrige", "Healthcare & Life Sciences":
              "Sundhed og life science", "Industrials": "Industri", "Software & IT":
              "Software og IT", "Business Services": "Erhvervsservice", "Education": "Uddannelse",
              "Financials": "Finans", "Building Products": "Byggematerialer",
              "SECONDARY": "Secondary", "CO_INVESTMENT": "Co-investment", "PRIMARY": "Primær"}
FORDELINGER = (("INDUSTRY", "Branche"), ("GEOGRAPHY", "Geografi"))


def _feeder_beholdninger(inv, nav_dato, nav):
    """Gennemlysning af en feeder: den underliggende fonds positioner med investors NAV fordelt.

    Feederen (fx PPPC2) ejer kun én fond (PSCP5); dens egen NAV fordeles på den underliggende
    fonds positioner efter dagsværdi. fordeling er GP'ens branche- og geografifordeling, hvor
    den findes, ellers fordelingen på typer af positioner; fakta er GP'ens antal fonde og
    selskaber og trukket kreditfacilitet. None for andre fonde, eller uden positioner.
    """
    kode = _kode(inv.underlying_fund_code)
    if kode is None:
        return None
    beh = _beholdninger(kode, nav_dato, None, nav=nav)
    if beh is None:
        return None
    # nogle forvaltere oplyser kun beløb (NEP); andelen er så beløbet af dimensionens sum
    bd = read_sql("""
        SELECT dimension, category,
               COALESCE(share_pct, 100.0 * amount
                        / NULLIF(SUM(amount) OVER (PARTITION BY dimension), 0)) AS share_pct
        FROM rpt.v_portfolio_breakdown
        WHERE fund_code = :f AND dimension IN ('INDUSTRY', 'GEOGRAPHY') AND report_date =
              (SELECT MAX(report_date) FROM rpt.v_portfolio_breakdown
               WHERE fund_code = :f AND report_date <= :d)
        ORDER BY dimension, share_pct DESC""", f=kode, d=nav_dato)
    fordeling = [(titel, [dict(navn=KATEGORIER.get(r.category, r.category), andel=r.share_pct / 100)
                          for r in bd[bd.dimension == dim].itertuples()])
                 for dim, titel in FORDELINGER if (bd.dimension == dim).any()]
    if not fordeling:
        samlet = sum(x["dagsvaerdi"] for x in beh["raekker"]) or 1
        typer = {}
        for x in beh["raekker"]:
            typer[x["type"]] = typer.get(x["type"], 0) + x["dagsvaerdi"] / samlet
        fordeling = [("Type", [dict(navn=KATEGORIER.get(t, t), andel=a)
                               for t, a in sorted(typer.items(), key=lambda x: -x[1])])]
    m = read_sql("""
        SELECT metric_code, value FROM rpt.v_fund_metric
        WHERE fund_code = :f AND report_date = :d
          AND metric_code IN ('NUMBER_OF_FUNDS', 'NUMBER_OF_COMPANIES', 'CREDIT_LINE_DRAWN',
                              'FUND_NAV')""", f=kode, d=beh["dato"])
    fakta = {r.metric_code: _f(r.value) for r in m.itertuples()}
    return beh | dict(via=inv.underlying_fund_name, via_kode=kode, fordeling=fordeling,
                      fakta=fakta)


# GICS-sektorer (portfolio.industry) og -branchegrupper (portfolio.industry_group) med deres
# officielle engelske navne, som decket viser dem. DIVERSIFIED og OTHER er databasens egne
# koder uden for GICS og står på dansk.
SEKTORER = {"ENERGY": "Energy", "MATERIALS": "Materials", "INDUSTRIALS": "Industrials",
            "CONSUMER_DISCRETIONARY": "Consumer Discretionary",
            "CONSUMER_STAPLES": "Consumer Staples", "HEALTH_CARE": "Health Care",
            "FINANCIALS": "Financials", "INFORMATION_TECHNOLOGY": "Information Technology",
            "COMMUNICATION_SERVICES": "Communication Services", "UTILITIES": "Utilities",
            "REAL_ESTATE": "Real Estate", "DIVERSIFIED": "Diversificeret (fondsandele)",
            "OTHER": "Ikke klassificeret af forvalteren"}
UDEN_GICS = {"DIVERSIFIED", "OTHER", None}      # har ingen branchegrupper
IKKE_KLASSIFICERET = "Ikke klassificeret"
IKKE_OPDELT = "Ikke opdelt af forvalteren"      # sektor oplyst, branchegruppe ikke
BRANCHEGRUPPER = {
    1010: "Energy", 1510: "Materials", 2010: "Capital Goods",
    2020: "Commercial & Professional Services", 2030: "Transportation",
    2510: "Automobiles & Components", 2520: "Consumer Durables & Apparel",
    2530: "Consumer Services", 2550: "Consumer Discretionary Distribution & Retail",
    3010: "Consumer Staples Distribution & Retail", 3020: "Food, Beverage & Tobacco",
    3030: "Household & Personal Products", 3510: "Health Care Equipment & Services",
    3520: "Pharmaceuticals, Biotechnology & Life Sciences", 4010: "Banks",
    4020: "Financial Services", 4030: "Insurance", 4510: "Software & Services",
    4520: "Technology Hardware & Equipment",
    4530: "Semiconductors & Semiconductor Equipment", 5010: "Telecommunication Services",
    5020: "Media & Entertainment", 5510: "Utilities", 6010: "Equity REITs",
    6020: "Real Estate Management & Development"}
# GICS-hierarkiet: de to første cifre i en branchegruppes kode er sektoren
_SEKTOR_CIFRE = {10: "ENERGY", 15: "MATERIALS", 20: "INDUSTRIALS", 25: "CONSUMER_DISCRETIONARY",
                 30: "CONSUMER_STAPLES", 35: "HEALTH_CARE", 40: "FINANCIALS",
                 45: "INFORMATION_TECHNOLOGY", 50: "COMMUNICATION_SERVICES", 55: "UTILITIES",
                 60: "REAL_ESTATE"}
GICS_STRUKTUR = {sektor: [k for k in sorted(BRANCHEGRUPPER) if k // 100 == cifre]
                 for cifre, sektor in _SEKTOR_CIFRE.items()}


def _heltal(v):
    return None if v is None or pd.isna(v) else int(v)


# Geografiske regioner (portfolio.geo_region) og delregioner (portfolio.geo_subregion) på
# dansk. *_OTHER er forvalterens tal, når de ikke er delt længere ned end regionen.
REGIONER = {"NORTH_AMERICA": "Nordamerika", "EUROPE": "Europa", "ROW": "Resten af verden"}
DELREGIONER = {
    "US": "USA", "CANADA": "Canada", "USA_OTHER": "Nordamerika, ikke opdelt",
    "DENMARK": "Danmark", "NORDICS": "Norden ekskl. Danmark",
    "UK_IRELAND": "Storbritannien og Irland", "DACH": "Tyskland, Østrig og Schweiz",
    "FRANCE": "Frankrig", "BENELUX": "Benelux", "SOUTHERN_EUROPE": "Sydeuropa",
    "EUROPE_OTHER": "Øvrige Europa og ikke opdelt", "LATIN_AMERICA": "Latinamerika",
    "ASIA_PACIFIC": "Asien og Stillehavsområdet", "MIDDLE_EAST_AFRICA": "Mellemøsten og Afrika",
    "ROW_OTHER": "Global, ikke opdelt"}
IKKE_OPDELTE_REGIONER = {"USA_OTHER", "EUROPE_OTHER", "ROW_OTHER"}

# Branche- og geografifordelingen læses og lægges sammen på samme måde: et rpt-view med
# investors eksponering pr. fond, to niveauer (sektor/branchegruppe, region/delregion) og
# den del, der hviler på en klassifikation, som ikke er gennemgået. Kolonnenavnene her er
# faste tekster, aldrig input, så de kan sættes direkte ind i SQL'en.
ARTER = {
    "brancher": dict(view="rpt.v_industry_exposure", top="industry_code",
                     under="industry_group_gics", under_navn="industry_group",
                     forslag="industry_suggested", dimension="INDUSTRY",
                     mapning="portfolio.industry_mapping", udeladt="industry_code"),
    "regioner": dict(view="rpt.v_geography_exposure", top="region_code",
                     under="subregion_code", under_navn="subregion",
                     forslag="geography_suggested", dimension="GEOGRAPHY",
                     mapning="portfolio.geo_mapping", udeladt="subregion_code"),
}


def _under_kode(art, v):
    return _heltal(v) if art == "brancher" else _kode(v)


def _navne(art, top, under, under_engelsk):
    """(top, under) som decket viser dem. under er None, hvor niveauet ikke deles (Diversificeret o.l.)."""
    if art == "brancher":
        if top in UDEN_GICS:
            return SEKTORER.get(top, top or IKKE_KLASSIFICERET), None
        return (SEKTORER.get(top, top),
                BRANCHEGRUPPER.get(under, under_engelsk) if under else IKKE_OPDELT)
    if top is None:
        return IKKE_KLASSIFICERET, None
    return REGIONER.get(top, top), DELREGIONER.get(under, under_engelsk or under)


def _eksponering(iid, nav_dato, fund_code, art):
    """Fondens branche- eller geografifordeling som andele, der summerer til 1, eller None.

    Andelene kommer fra rpt-viewet i ARTER, som vælger forvalterens fordeling,
    beholdningernes dagsværdi eller fondens strategi (FUND_DEFAULT) og gennemlyser feederne.
    Viewets beløb bruges kun til vægte: investors NAV i DKK sættes på i eksponering(), så
    kurs og NAV er de samme som på de andre sider. Hver linje får også den andel, der
    hviler på en klassifikation, som endnu ikke er gennemgået (is_suggested).
    """
    a = ARTER[art]
    e = read_sql(f"""
        SELECT look_through_fund, split_date, basis, {a['top']} AS niveau1, {a['under']} AS niveau2,
               {a['under_navn']} AS under_navn, SUM(nav_amount) AS beloeb
        FROM {a['view']}
        WHERE investment_id = :i AND period_end = :d AND basis IS NOT NULL
        GROUP BY look_through_fund, split_date, basis, {a['top']}, {a['under']},
                 {a['under_navn']}""", i=iid, d=nav_dato)
    i_alt = _f(e.beloeb.sum()) if not e.empty else None
    if not i_alt:
        return None
    kode, dato, grundlag = e.look_through_fund[0], _dato(e.split_date[0]), e.basis[0]
    if grundlag == "HOLDINGS":
        k = read_sql(f"""
            SELECT {a['top']} AS niveau1, {a['under']} AS niveau2, document_id,
                   SUM(fair_value) AS vaegt,
                   SUM(CASE WHEN {a['forslag']} = 1 THEN fair_value ELSE 0 END) AS foreloebig
            FROM rpt.v_portfolio_holding
            WHERE fund_code = :f AND report_date = :d AND fair_value > 0
            GROUP BY {a['top']}, {a['under']}, document_id""", f=kode, d=dato)
    elif grundlag == "GP_BREAKDOWN":
        # samme opslag som v_portfolio_breakdown: fondens egen mapning før den generelle
        k = read_sql(f"""
            SELECT b.{a['top']} AS niveau1, b.{a['under']} AS niveau2, b.document_id,
                   SUM(COALESCE(b.amount, b.share_pct)) AS vaegt,
                   SUM(CASE WHEN m.is_suggested = 1 THEN COALESCE(b.amount, b.share_pct)
                            ELSE 0 END) AS foreloebig
            FROM rpt.v_portfolio_breakdown b
            JOIN fund.fund f ON f.fund_code = b.fund_code
            OUTER APPLY (SELECT TOP 1 m.is_suggested FROM {a['mapning']} m
                         WHERE m.label = b.category AND (m.fund_id = f.fund_id OR m.fund_id IS NULL)
                         ORDER BY CASE WHEN m.fund_id IS NULL THEN 1 ELSE 0 END) m
            WHERE b.fund_code = :f AND b.report_date = :d AND b.dimension = '{a['dimension']}'
              AND ISNULL(b.{a['udeladt']}, '') <> 'EXCLUDED'
            GROUP BY b.{a['top']}, b.{a['under']}, b.document_id""", f=kode, d=dato)
    elif art == "brancher":
        k = read_sql("""
            SELECT m.industry_code AS niveau1, m.industry_group_gics AS niveau2,
                   NULL AS document_id, 1 AS vaegt, CAST(m.is_suggested AS int) AS foreloebig
            FROM portfolio.industry_mapping m JOIN fund.fund f ON f.fund_id = m.fund_id
            WHERE f.fund_code = :f AND m.label IS NULL""", f=kode)
    else:
        k = read_sql("""
            SELECT s.region_code AS niveau1, m.subregion_code AS niveau2,
                   NULL AS document_id, 1 AS vaegt, CAST(m.is_suggested AS int) AS foreloebig
            FROM portfolio.geo_mapping m JOIN fund.fund f ON f.fund_id = m.fund_id
            JOIN portfolio.geo_subregion s ON s.subregion_code = m.subregion_code
            WHERE f.fund_code = :f AND m.label IS NULL""", f=kode)
    vaegte = {}
    for x in k.itertuples():
        v = vaegte.setdefault((_kode(x.niveau1), _under_kode(art, x.niveau2)), [0.0, 0.0])
        v[0] += _f(x.vaegt) or 0
        v[1] += _f(x.foreloebig) or 0
    linjer = []
    for x in e.itertuples():
        top, under = _kode(x.niveau1), _under_kode(art, x.niveau2)
        vaegt, foreloebig = vaegte.get((top, under), (0, 0))
        top_navn, under_navn = _navne(art, top, under, _kode(x.under_navn))
        linjer.append(dict(top_kode=top, top=top_navn, under_kode=under, under=under_navn,
                           andel=_f(x.beloeb) / i_alt,
                           foreloebig=foreloebig / vaegt if vaegt else 0.0))
    return dict(grundlag=grundlag, dato=dato, gennem=kode if kode != fund_code else None,
                linjer=linjer,
                dokumenter=sorted({int(x) for x in k.document_id.dropna()}))


def fondsstoerrelse(metric, stamdata, nav_dato):
    """Fondens samlede tilsagn og kilden: (beløb, dict(tekst, document_id)).

    metric er GP'ens daterede TOTAL_COMMITMENTS som (beløb, dato, document_id) eller None. Den
    vinder, når den højst er et år ældre end NAV-datoen; ellers stamdataenes fondsstørrelse,
    for en fond med løbende tilsagn (Home.Earth: 52,9 mio. EUR i 2023, 158,2 mio. EUR i 2026)
    vokser efter den sidste daterede opgørelse. Uden stamdata bruges den gamle metric.
    """
    if metric and ((nav_dato - metric[1]).days <= 366 or not stamdata):
        beloeb, dato, dok = metric
        return beloeb, dict(tekst=f"Fondsrapport pr. {dato:%d.%m.%Y} (TOTAL_COMMITMENTS)",
                            document_id=dok)
    return stamdata, dict(tekst="Stamdata (rpt.v_fund_master)", document_id=None)


def _er_indbetaling(cf):
    """Kapitalkald plus betalt udligningsrente (EQUALISATION med LP-fortegn < 0)."""
    return (cf.flow_class == "CONTRIBUTION") | ((cf.flow_class == "EQUALISATION")
                                                & (cf.amount < 0))


def _er_udlodning(cf):
    """Udlodninger plus modtaget udligningsrente (EQUALISATION > 0)."""
    return (cf.flow_class == "DISTRIBUTION") | ((cf.flow_class == "EQUALISATION")
                                                & (cf.amount > 0))


def _datoer(cf):
    return ", ".join(f"{d:%d.%m.%Y}" for d in sorted(set(cf.value_date)))


def fond(inv, as_of, info):
    """Alle tal til én fondsside. inv er en række fra _investeringer()."""
    iid, kode = int(inv.investment_id), inv.fund_code
    info = info.get(kode, {}) or {}
    noter = []


    nav = read_sql("""SELECT period_end, COALESCE(restated_nav, nav) AS nav,
                             nav AS nav_rapporteret, restatement, document_id FROM rpt.v_nav
                      WHERE investment_id = :i AND period_end <= :d ORDER BY period_end""",
                   i=iid, d=as_of)
    if nav.empty:
        return None
    nav["period_end"] = nav.period_end.map(_dato)
    nav_dato = nav.period_end.iloc[-1]

    cf = read_sql("""SELECT value_date, notice_date, account_code, account_name, flow_class,
                            is_recallable, amount, fx_rate, amount_dkk, flow_ref, document_id
                     FROM rpt.v_cash_flow WHERE investment_id = :i
                     ORDER BY value_date, account_code""", i=iid)
    cf["value_date"] = cf.value_date.map(_dato)
    cf["amount"] = cf.amount.astype(float)
    # intet efter rapportdatoen: en rapport pr. en tidligere dato skal se ud, som den gjorde da
    cf = cf[cf.value_date <= as_of]
    alle_cf = cf
    senere, cf = cf[cf.value_date > nav_dato], cf[cf.value_date <= nav_dato]

    tilsagn_haendelser = read_sql("""SELECT event_date, event_type, amount, document_id, note
                                     FROM fund.commitment_event
                                     WHERE investment_id = :i AND event_date <= :d
                                     ORDER BY event_date""", i=iid, d=nav_dato)
    tilsagn = _f(tilsagn_haendelser.amount.sum()) if not tilsagn_haendelser.empty else None
    # udligningsrente ved senere closings er betalt ud over tilsagnet: den betalte tæller
    # med i indbetalt og den modtagne i udloddet, men den nedbringer ikke tilsagnet, så
    # kaldt kapital og resttilsagn regnes uden (som forvalterens MM110)
    betalt_udl = _er_indbetaling(cf) & (cf.flow_class == "EQUALISATION")
    modtaget_udl = _er_udlodning(cf) & (cf.flow_class == "EQUALISATION")
    udligning_betalt = -cf[betalt_udl].amount.sum()
    udligning_modtaget = cf[modtaget_udl].amount.sum()
    cf_indbetalt = -cf[cf.flow_class == "CONTRIBUTION"].amount.sum()
    udloddet = cf[_er_udlodning(cf)].amount.sum()
    recallable = cf[(cf.flow_class == "DISTRIBUTION") & (cf.is_recallable == 1)].amount.sum()

    ka = _kapitalkonto(iid, nav_dato)
    ka["period_end"] = ka.period_end.map(_dato)
    gp_indbetalt = _gp_indbetalt(ka, nav_dato)
    mod_tilsagn = cf_indbetalt
    if gp_vinder(cf_indbetalt, gp_indbetalt, recallable):
        mod_tilsagn = gp_indbetalt
        # forvalterens tal står i nøgletallene; noten siger kun, hvad kaldene summerer til
        noter.append(Note(f"Indbetalt kapital følger forvalterens opgørelse; de modtagne "
                          f"kapitalkald summerer til {cf_indbetalt / 1000:,.0f} t{inv.currency}."
                          .replace(",", "."), "indbetalt"))

    estimerede, advarsler = estimerede_cashflows(info, inv.investor_code, cf, nav_dato)
    est = [(_dato(e["dato"]), float(e["beloeb"])) for e in estimerede]
    if est:
        noter.append(Note("IRR og historik medregner kapitalkald med skønnet dato, hvor "
                          "meddelelsen ikke er modtaget: "
                          + "; ".join(e["note"] for e in estimerede) + ".", "irr"))

    nav_nu = float(nav.nav.iloc[-1])
    irr_flows = [(r.value_date, r.amount) for r in cf.itertuples()
                 if r.flow_class != "WITHHOLDING_TAX"]
    irr_flows = irr_flows + est + [(nav_dato, nav_nu)]
    irr = xirr(irr_flows)

    tc = read_sql("""SELECT TOP 1 value, report_date, document_id FROM rpt.v_fund_metric
                     WHERE fund_code = :f AND metric_code = 'TOTAL_COMMITMENTS' AND report_date <= :d
                     ORDER BY report_date DESC""", f=kode, d=nav_dato)
    fondsstr, fondsstr_kilde = fondsstoerrelse(
        None if tc.empty else (_f(tc.value[0]), _dato(tc.report_date[0]),
                               int(tc.document_id[0])),
        _f(inv.total_commitments), nav_dato)
    andel = tilsagn / fondsstr if tilsagn and fondsstr else None
    evergreen = bool(info.get("evergreen"))
    fonds_nav = None
    if evergreen:
        # en evergreen-fond udsteder nye andele til den aktuelle kurs, så tilsagn / samlede
        # tilsagn undervurderer ejerskabet for de tidlige investorer; NAV-andelen er den rigtige
        fn = read_sql("""SELECT value FROM rpt.v_fund_metric
                         WHERE fund_code = :f AND metric_code = 'FUND_NAV' AND report_date = :d""",
                      f=kode, d=nav_dato)
        fonds_nav = None if fn.empty else _f(fn.value[0])
        if fonds_nav:
            # Definitioner forklarer ejerandelen for en evergreen-fond, så der er ingen note
            andel = float(nav.nav.iloc[-1]) / fonds_nav

    # hele beløb: rabat og udligningsrente er typisk skæve beløb (82.500), som tusinder
    # runder forkert
    hel = lambda v: f"{v:,.0f}".replace(",", ".") + f" {inv.currency}"
    if udligning_betalt or udligning_modtaget:
        dele = []
        if udligning_betalt:
            dele.append(f"Indbetalt kapital indeholder {hel(udligning_betalt)} i "
                        f"udligningsrente ({_datoer(cf[betalt_udl])})")
        if udligning_modtaget:
            dele.append(f"udlodningerne {hel(udligning_modtaget)} i modtaget udligningsrente "
                        f"({_datoer(cf[modtaget_udl])})")
        tekst = " og ".join(dele)
        # hvad udligningsrente er, står i Definitioner; noten giver beløbet og følgen
        noter.append(Note(tekst[0].upper() + tekst[1:] + "; den indgår ikke i kaldt kapital og "
                          "resttilsagn.", "indbetalt"))
    kaldt_ = kaldt(mod_tilsagn, _memo(ka, nav_dato, "MM110"))
    rabat_, rabat_dato = rabat(mod_tilsagn, kaldt_, cf)
    indbetalt = mod_tilsagn + rabat_ + udligning_betalt
    udloddet = udloddet + rabat_
    if rabat_:
        # den indbetalte kapital og udlodningerne står i nøgletallene
        noter.append(Note(f"Forvalteren har givet {hel(rabat_)} i rabat på kapitalkaldet "
                          f"{rabat_dato:%d.%m.%Y}; rabatten indgår i indbetalt kapital og som "
                          "udlodning.", "indbetalt"))
    n = dict(tilsagn=tilsagn, indbetalt=indbetalt, kaldt=kaldt_, rabat=rabat_,
             udligning=float(udligning_betalt),
             udloddet=udloddet, nav=nav_nu, recallable=float(recallable),
             resttilsagn=(tilsagn - kaldt_ + recallable) if tilsagn else None,
             picc=indbetalt / tilsagn if tilsagn else None,
             dpi=udloddet / indbetalt if indbetalt else None,
             tvpi=(udloddet + nav_nu) / indbetalt if indbetalt else None,
             irr=irr, andel=andel, fondsstoerrelse=fondsstr, fonds_nav=fonds_nav,
             gp_irr=_memo(ka, nav_dato, "MM150"), gp_tvpi=_memo(ka, nav_dato, "MM160"))

    # kvartalshistorik: NAV og akkumulerede cash flows (inkl. de skønnede kapitalkald)
    ind = [(d, -a) for d, a in zip(cf[_er_indbetaling(cf)].value_date,
                                   cf[_er_indbetaling(cf)].amount)]
    ind += [(d, -a) for d, a in est]
    udl = list(zip(cf[_er_udlodning(cf)].value_date, cf[_er_udlodning(cf)].amount))
    if rabat_:
        ind.append((rabat_dato, rabat_))
        udl.append((rabat_dato, rabat_))
    historik = [dict(dato=r.period_end, nav=float(r.nav),
                     indbetalt=sum(a for d, a in ind if d <= r.period_end),
                     udloddet=sum(a for d, a in udl if d <= r.period_end))
                for r in nav.itertuples()]

    if recallable:
        # fondssiden viser tilsagn, indbetalt og resttilsagn side om side; uden denne note
        # ser de tre tal ud til ikke at gå op
        # tusindseparatoren laves for sig; .replace(",", ".") på hele sætningen ville også
        # ramme kommaet foran "som"
        beloeb = f"{recallable / 1000:,.0f}".replace(",", ".")
        noter.append(Note(f"Resttilsagnet indeholder genkaldelige udlodninger på {beloeb} "
                          f"t{inv.currency}.", "resttilsagn"))

    if not senere.empty:
        dele = []
        for kl, tekst in (("CONTRIBUTION", "kapitalkald"), ("DISTRIBUTION", "udlodning")):
            s = senere[senere.flow_class == kl]
            if not s.empty:
                dele.append(f"{tekst} {abs(s.amount.sum()) / 1000:,.0f} t{inv.currency}"
                            .replace(",", ".")
                            + f" ({', '.join(f'{d:%d.%m.%Y}' for d in sorted(set(s.value_date)))})")
        if dele:
            noter.append(Note("Efter NAV-datoen: " + " og ".join(dele) + " – ikke medregnet.",
                              "nav"))

    beh = (_beholdninger(kode, nav_dato, andel)
           or _feeder_beholdninger(inv, nav_dato, nav_nu)
           or _ejendomsbalance(inv, nav_dato))
    # fondens samlede NAV med og uden hensat carry, så forskellen mellem investors NAV og
    # beholdningernes dagsværdi kan forklares på fondssiden
    fonds_navs = {}
    if beh and not beh.get("via") and beh["type"] != "ejendomme" and beh["skaleret"]:
        m = read_sql("""SELECT metric_code, value, document_id FROM rpt.v_fund_metric
                        WHERE fund_code = :f AND report_date = :d
                          AND metric_code IN ('FUND_NAV', 'FUND_NAV_EXCL_CARRY')""",
                     f=beh["kode"], d=beh["dato"])
        fonds_navs = {r.metric_code: (_f(r.value), None if pd.isna(r.document_id)
                                      else int(r.document_id)) for r in m.itertuples()}

    return dict(
        # decket viser navnet uden selskabsform; det juridiske navn står i Excel-stamdataene
        fund_code=kode, fund_name=uden_selskabsform(inv.fund_name), juridisk_navn=inv.fund_name,
        manager=inv.manager_name,
        # en feeder viser logoet for den underliggende fond, som pengene reelt ligger i
        logo=_logo(_kode(inv.underlying_fund_code) or kode), valuta=inv.currency,
        strategi=inv.strategy, vintage=None if pd.isna(inv.vintage_year) else int(inv.vintage_year),
        beskrivelse=info.get("beskrivelse"), geografi=geografi(inv.geography_focus),
        segment=segment(inv.fund_structure, inv.strategy_focus, inv.size_focus),
        re_sektor=_kode(inv.re_sector),     # ejendomsfondenes segment (RESIDENTIAL, OFFICE ...)
        loebetid="Evergreen" if evergreen else loebetid(inv.term_years, inv.extension_years),
        evergreen=evergreen,
        benchmark=info.get("benchmark"), virksomheder=info.get("virksomheder") or {},
        nav_dato=nav_dato, n=n, historik=historik,
        bro=_vaerdibro(ka, nav_dato, n),
        beholdninger=beh,
        nav_afstemning=nav_afstemning(nav_nu, nav_dato, beh,
                                      (fonds_navs.get("FUND_NAV") or (None,))[0],
                                      (fonds_navs.get("FUND_NAV_EXCL_CARRY") or (None,))[0]),
        brancher=_eksponering(iid, nav_dato, kode, "brancher"),
        regioner=_eksponering(iid, nav_dato, kode, "regioner"),
        irr_flows=irr_flows, noter=noter,
        # kapitalstroemme er de cash flows, historikken bygger på: kald inkl. de skønnede
        kapitalstroemme=[(d, -a, "kald") for d, a in ind] + [(d, a, "udlodning") for d, a in udl],
        # rækkerne bag tallene, med dokument-id, til dataudtrækket i Excel
        kilder=dict(nav=nav, cash_flows=alle_cf, kapitalkonto=ka, tilsagn=tilsagn_haendelser,
                    fondsstoerrelse=fondsstr_kilde, fonds_nav=fonds_navs,
                    indbetalt="forvalteren" if mod_tilsagn != cf_indbetalt else "kapitalkald",
                    estimerede=estimerede, rabat=(rabat_dato, rabat_) if rabat_ else None),
        # til bygningen, ikke til decket: forhold, der skal ses på før udsendelse
        advarsler=advarsler + afstem_mod_forvalteren(n, _memo(ka, nav_dato, "MM120"),
                                                     inv.currency)
        + (["beskrivelserne i fund_info.yaml er ikke gennemgået (status: udkast)"]
           if info.get("status") == "udkast" else []))


def total(fonde):
    """Porteføljens samlede nøgletal; IRR er poolet over alle fondes cash flows.

    Beløb i forskellige valutaer lægges ikke sammen: None, hvis fondene ikke har samme valuta.
    """
    if len({f["valuta"] for f in fonde}) != 1:
        return None
    # et ukendt beløb gør summen ukendt; som 0 ville det fx give for høj PICC
    s = lambda k: (None if any(f["n"][k] is None for f in fonde)
                   else sum(f["n"][k] for f in fonde))
    tilsagn, indbetalt, udloddet, nav = s("tilsagn"), s("indbetalt"), s("udloddet"), s("nav")
    return dict(tilsagn=tilsagn, indbetalt=indbetalt, udloddet=udloddet, nav=nav,
                recallable=s("recallable"), resttilsagn=s("resttilsagn"),
                picc=indbetalt / tilsagn if tilsagn and indbetalt is not None else None,
                dpi=udloddet / indbetalt if indbetalt and udloddet is not None else None,
                tvpi=((udloddet + nav) / indbetalt
                      if indbetalt and udloddet is not None and nav is not None else None),
                irr=xirr([x for f in fonde for x in f["irr_flows"]]))


def gennemlysning(fonde):
    """Porteføljens beholdninger samlet i én liste, opgjort som investors andel.

    Linjerne blandes på tværs af fondene, så underliggende fonde og porteføljeselskaber står
    side om side; "type" på hver linje siger hvad den er. None, hvis fondene ikke har samme
    valuta, eller ingen fond har beholdninger.
    """
    if len({f["valuta"] for f in fonde}) != 1:
        return None
    med = [f for f in fonde if f.get("beholdninger")
           and f["beholdninger"]["type"] != "ejendomme"]
    if not med:
        return None
    linjer = [dict(navn=r["navn"], fond=f["fund_name"], type=b["type"], kostpris=r["kostpris"],
                   dagsvaerdi=r["dagsvaerdi"], udloddet=r["udloddet"], total=r["total"],
                   tvpi=r["tvpi"])
              for f in med for b in [f["beholdninger"]] for r in b["raekker"]]
    linjer.sort(key=lambda x: -x["dagsvaerdi"])
    tael = lambda t: sum(len(f["beholdninger"]["raekker"]) for f in med
                         if f["beholdninger"]["type"] == t)
    return dict(linjer=linjer, antal_fonde=len(fonde), antal_underfonde=tael("fonde"),
                antal_selskaber=tael("selskaber"),
                dato=max(f["beholdninger"]["dato"] for f in med),
                dagsvaerdi=sum(x["dagsvaerdi"] for x in linjer),
                nav=sum(f["n"]["nav"] or 0 for f in fonde),
                uden_beholdninger=[f["fund_name"] for f in fonde if not f.get("beholdninger")])


NAV_SYNLIG = 500          # mindste forskel mellem NAV og dagsværdi, der ses i tusinder


def nav_afstemning(nav, nav_dato, beh, fonds_nav=None, fonds_nav_ex_carry=None):
    """Hvorfor investors NAV afviger fra beholdningernes dagsværdi i alt, eller None.

    Beholdningsoversigten er fondens positioner; fondens NAV rummer også hensat carried
    interest, likvider og øvrige poster, og investors NAV er ikke nødvendigvis præcis
    ejerandelen (tilsagn / fondens tilsagn) af fondens NAV. Med fondens NAV deles forskellen
    i de tre dele, som summerer til forskellen:
        carry     = andel · (fondens NAV - fondens NAV før carry)
        oevrige   = andel · (fondens NAV før carry - beholdningernes dagsværdi i fonden)
        ejerandel = investors NAV - andel · fondens NAV
    Uden fondens NAV er delene None. Kun skalerede beholdninger af selskaber eller fonde; en
    feeder er skaleret til NAV, og en ejendomsfond går til NAV i sin waterfall.
    """
    if not beh or beh.get("via") or beh["type"] == "ejendomme" or not beh["skaleret"]:
        return None
    dagsvaerdi = sum(r["dagsvaerdi"] for r in beh["raekker"])
    forskel = nav - dagsvaerdi
    if abs(forskel) < NAV_SYNLIG:
        return None
    a = beh["andel"]
    carry = oevrige = ejerandel = None
    if fonds_nav:
        fv_fond = sum(r["dagsvaerdi_fond"] for r in beh["raekker"])
        carry = a * (fonds_nav - fonds_nav_ex_carry) if fonds_nav_ex_carry else None
        oevrige = a * ((fonds_nav_ex_carry or fonds_nav) - fv_fond)
        ejerandel = nav - a * fonds_nav
    return dict(nav=nav, dagsvaerdi=dagsvaerdi, forskel=forskel, carry=carry, oevrige=oevrige,
                ejerandel=ejerandel, andel=a, nav_andel=nav / fonds_nav if fonds_nav else None,
                beholdninger_dato=beh["dato"], nav_dato=nav_dato)


def stoerste_eksponering(fonde, kurser):
    """Den største enkelte beholdning på tværs af fondene, i DKK, eller None.

    Kandidaterne er fondenes beholdningslinjer - porteføljeselskaber og underliggende fonde,
    for en feeder den underliggende fonds positioner - opgjort som investors andel og
    omregnet med kursen på rapportdatoen. Ejendomme (hele fondens værdier før gæld) og
    uskalerede beholdninger (hele fondens tal) kan ikke sammenlignes og er udeladt; de står i
    "udeladt" sammen med fonde uden beholdninger.
    """
    positioner, udeladt = _positionskandidater(fonde, kurser)
    if not positioner:
        return None
    p = positioner[0]
    nav = sum(x["n"].get("nav_dkk") or 0 for x in fonde)
    return dict(navn=p["navn"], fond=p["fonde"][0], via=p["via"], dkk=p["dkk"], art=p["art"],
                andel=p["dkk"] / nav if nav else None, dato=p["dato"], udeladt=udeladt)


def _positionskandidater(fonde, kurser):
    """Beholdningslinjerne på tværs af fondene som investors andel i DKK, størst først.

    Samme navn (uden hensyn til store og små bogstaver) i flere fonde lægges sammen til én
    position. Returnerer (positioner, udeladte fonde): ejendomme, uskalerede beholdninger og
    fonde uden beholdninger eller kurs kan ikke sammenlignes og er udeladt.
    """
    positioner, udeladt = {}, []
    for f in fonde:
        b = f.get("beholdninger")
        kurs = (kurser.get(f["valuta"]) or (None,))[0]
        if not b or b["type"] == "ejendomme" or not b["skaleret"] or kurs is None:
            udeladt.append(f["fund_name"])
            continue
        art = "underliggende fond" if b["type"] == "fonde" else "porteføljeselskab"
        for r in b["raekker"]:
            p = positioner.setdefault(r["navn"].casefold(), dict(
                navn=r["navn"], art=art, fonde=[], via=b.get("via"), dkk=0.0, dato=b["dato"]))
            p["dkk"] += r["dagsvaerdi"] * kurs
            if f["fund_name"] not in p["fonde"]:
                p["fonde"].append(f["fund_name"])
            p["dato"] = max(p["dato"], b["dato"])
    return sorted(positioner.values(), key=lambda p: -p["dkk"]), udeladt


POSITIONER_ANTAL = 15


def stoerste_positioner(fonde, kurser, antal=POSITIONER_ANTAL):
    """Porteføljens største positioner i DKK som andel af investors samlede NAV, eller None.

    Positionerne er beholdningslinjerne som i stoerste_eksponering() (investors andel af
    dagsværdien, feederne gennemlyst, samme navn i flere fonde lagt sammen). En fond uden
    sammenlignelige beholdninger - ejendomsfondene, uskalerede beholdninger og fonde uden
    beholdninger - indgår som én position med sin NAV (art "fond"). None uden DKK-tal.
    """
    if not fonde or any(f["n"].get("nav_dkk") is None for f in fonde):
        return None
    positioner, udeladt = _positionskandidater(fonde, kurser)
    fondslinjer = [f for f in fonde if f["fund_name"] in udeladt]
    positioner += [dict(navn=f["fund_name"], art="fond", fonde=[f["fund_name"]], via=None,
                        dkk=f["n"]["nav_dkk"], dato=f["nav_dato"]) for f in fondslinjer]
    positioner.sort(key=lambda p: -p["dkk"])
    nav = sum(f["n"]["nav_dkk"] for f in fonde)
    linjer = [p | {"andel": p["dkk"] / nav if nav else None} for p in positioner[:antal]]
    top = sum(p["dkk"] for p in linjer)
    return dict(linjer=linjer, nav_dkk=nav, top_dkk=top, top_andel=top / nav if nav else None,
                fondslinjer=[f["fund_name"] for f in fondslinjer])


def datotjek(fonde, rapport_dato):
    """Kontrollér at alle fonde er opgjort pr. samme dato.

    To ting kan skride: en fond kan mangle den seneste NAV, så dens tal er ældre end resten
    af rapporten, og dens beholdninger kan være opgjort på en anden dato end dens egen NAV.
    Begge dele gør, at totaler og fordelinger blander opgørelsesdatoer, så de skal frem på
    sliden i stedet for at ligge i en fodnote.
    """
    gamle = [dict(fond=f["fund_name"], dato=f["nav_dato"],
                  kvartaler=_kvartaler_mellem(f["nav_dato"], rapport_dato))
             for f in fonde if f["nav_dato"] < rapport_dato]
    skaeve = [dict(fond=f["fund_name"], nav_dato=f["nav_dato"],
                   beholdninger_dato=f["beholdninger"]["dato"])
              for f in fonde
              if f["beholdninger"] and f["beholdninger"]["dato"] != f["nav_dato"]]
    return dict(rapport_dato=rapport_dato, ens=not gamle and not skaeve,
                foraeldede=gamle, skaeve_beholdninger=skaeve)


def _kvartaler_mellem(fra, til):
    return (til.year - fra.year) * 4 + (til.month - fra.month) // 3


def _dkk_kurs(valuta, dato):
    """(DKK pr. 1 enhed, kursdato) fra Nationalbanken, seneste kurs på eller før dato.

    DKK har ingen rækker i kurstabellen og får kurs 1. None, hvis der ingen kurs findes.
    """
    if valuta == "DKK":
        return 1.0, dato
    df = read_sql("""SELECT TOP 1 rate_date, dkk_per_1 FROM dbo.fx_rates_nationalbanken
                     WHERE currency = :v AND rate_date <= :d ORDER BY rate_date DESC""",
                  v=valuta, d=dato)
    if df.empty:
        return None
    return _f(df.dkk_per_1[0]), _dato(df.rate_date[0])


def omregn_til_dkk(fonde, kurser):
    """Sæt beløbene omregnet til DKK (tilsagn_dkk, nav_dkk, …) på hver fonds nøgletal.

    None uden kurs.
    """
    for f in fonde:
        kurs = (kurser.get(f["valuta"]) or (None,))[0]
        n = f["n"]
        for k in ("tilsagn", "indbetalt", "udloddet", "resttilsagn", "nav"):
            n[f"{k}_dkk"] = n[k] * kurs if kurs is not None and n[k] is not None else None
        if kurs is None:
            f["noter"].append(f"Nationalbanken oplyser ingen {f['valuta']}-kurs, så fonden "
                              "kan ikke omregnes til DKK.")


DKK_FELTER = ("tilsagn_dkk", "indbetalt_dkk", "udloddet_dkk", "resttilsagn_dkk", "nav_dkk")


def dkk_total(fonde):
    """DKK-beløbene summeret på tværs af valutaer.

    None, hvis en fond mangler et tal.
    """
    total = {}
    for k in DKK_FELTER:
        vaerdier = [f["n"].get(k) for f in fonde]
        if not vaerdier or any(v is None for v in vaerdier):
            return None
        total[k] = sum(vaerdier)
    return total


def valutafordeling(fonde):
    """NAV, resttilsagn og eksponering (NAV + resttilsagn) i DKK fordelt på fondenes valuta.

    {"nav": [...], "resttilsagn": [...], "eksponering": [...]}, hver [{valuta, dkk, andel}]
    med den største først; tom, når målet summerer til 0. None, hvis en fond mangler et
    DKK-tal.
    """
    if not fonde or any(f["n"].get(k) is None for f in fonde
                        for k in ("nav_dkk", "resttilsagn_dkk")):
        return None
    resultat = {}
    for maal, beloeb in (("nav", lambda n: n["nav_dkk"]),
                         ("resttilsagn", lambda n: n["resttilsagn_dkk"]),
                         ("eksponering", lambda n: n["nav_dkk"] + n["resttilsagn_dkk"])):
        pr_valuta = {}
        for f in fonde:
            pr_valuta[f["valuta"]] = pr_valuta.get(f["valuta"], 0.0) + beloeb(f["n"])
        i_alt = sum(pr_valuta.values())
        resultat[maal] = [] if not i_alt else [
            dict(valuta=v, dkk=d, andel=d / i_alt)
            for v, d in sorted(pr_valuta.items(), key=lambda x: -x[1])]
    return resultat


def _gruppesum(fonde):
    """Summen af en gruppe fonde til fordelingssiderne.

    DKK-beløbene lægges altid sammen, tilsagn i fondens valuta kun, når alle fonde i gruppen
    har samme valuta. TVPI er (udloddet + NAV) / indbetalt i DKK, alt omregnet med kursen på
    rapportdatoen, så den kan beregnes på tværs af valutaer. For en gruppe i én valuta er
    den lig TVPI i fondens valuta, fordi kursen går ud.
    """
    valutaer = {f["valuta"] for f in fonde}
    valuta = valutaer.pop() if len(valutaer) == 1 else None
    dkk = {k: (None if any(f["n"].get(k) is None for f in fonde)
               else sum(f["n"][k] for f in fonde)) for k in DKK_FELTER}
    ind = dkk["indbetalt_dkk"]
    tvpi = ((dkk["udloddet_dkk"] + dkk["nav_dkk"]) / ind
            if ind and dkk["udloddet_dkk"] is not None and dkk["nav_dkk"] is not None else None)
    return dict(valuta=valuta, tilsagn=sum(f["n"]["tilsagn"] or 0 for f in fonde) if valuta
                else None, tvpi=tvpi, **dkk)


def fordeling(fonde, noegle):
    """Fondene grupperet på noegle ("geografi" eller "segment") med andel af tilsagnet i DKK.

    {total, grupper: [{navn, fonde, sum, andel, andel_nav}], andele: {fund_name: andel}}.
    andel_nav er gruppens andel af NAV i DKK (executive summary). Grupperne sorteres efter
    tilsagn i DKK. Andelene er None, hvis en fond mangler en DKK-kurs.
    """
    total = _gruppesum(fonde)
    i_alt, nav_i_alt = total["tilsagn_dkk"], total["nav_dkk"]
    andel = lambda v: v / i_alt if i_alt and v is not None else None
    andel_nav = lambda v: v / nav_i_alt if nav_i_alt and v is not None else None
    grupper = {}
    for f in fonde:
        grupper.setdefault(f[noegle] or "Ikke oplyst", []).append(f)
    resultat = [dict(navn=navn, fonde=g, sum=(s := _gruppesum(g)), andel=andel(s["tilsagn_dkk"]),
                     andel_nav=andel_nav(s["nav_dkk"]))
                for navn, g in grupper.items()]
    if noegle == "segment":
        resultat.sort(key=lambda g: _strategi_plads(g["navn"]))
    else:
        resultat.sort(key=lambda g: -(g["sum"]["tilsagn_dkk"] or 0))
    return dict(total=total | {"andel": andel(i_alt), "andel_nav": andel_nav(nav_i_alt)},
                grupper=resultat,
                andele={f["fund_name"]: andel(f["n"].get("tilsagn_dkk")) for f in fonde})


# hvad der står sidst i hver art: det uklassificerede og det, forvalteren ikke deler op
SIDST_TOP = {"brancher": UDEN_GICS, "regioner": {None}}


def _sidst_under(art, navn):
    if art == "brancher":
        return navn == IKKE_OPDELT
    return navn in {DELREGIONER[k] for k in IKKE_OPDELTE_REGIONER}


def eksponering(fonde, art):
    """Investors NAV i DKK fordelt på tværs af fondene: sektor/branchegruppe eller region/delregion.

    art er "brancher" eller "regioner"; hver fonds andele (f[art] fra _eksponering) ganges
    med fondens NAV i DKK, så grupperne summerer til porteføljens NAV. En fond uden fordeling
    står samlet som "Ikke klassificeret". Grupperne sorteres efter beløb med det
    uklassificerede sidst, undergrupperne efter beløb med det ikke opdelte sidst. None, hvis
    en fond mangler NAV i DKK.
    """
    if not fonde or any(f["n"].get("nav_dkk") is None for f in fonde):
        return None
    grupper, foreloebig = {}, 0.0
    for f in fonde:
        b = f.get(art)
        linjer = b["linjer"] if b else [dict(top_kode=None, top=IKKE_KLASSIFICERET, under=None,
                                             andel=1.0, foreloebig=0.0)]
        for x in linjer:
            dkk = x["andel"] * f["n"]["nav_dkk"]
            g = grupper.setdefault(x["top"], dict(kode=x["top_kode"], navn=x["top"], dkk=0.0,
                                                  under={}, fonde={}))
            g["dkk"] += dkk
            g["fonde"][f["fund_name"]] = g["fonde"].get(f["fund_name"], 0.0) + dkk
            if x["under"]:
                u = g["under"].setdefault(x["under"], {})
                u[f["fund_name"]] = u.get(f["fund_name"], 0.0) + dkk
            foreloebig += x["foreloebig"] * dkk
    total = sum(f["n"]["nav_dkk"] for f in fonde)
    andel = lambda v: v / total if total else None
    resultat = []
    for g in sorted(grupper.values(), key=lambda g: (g["kode"] in SIDST_TOP[art], -g["dkk"])):
        under = sorted(((n, sum(u.values()), max(u, key=u.get)) for n, u in g["under"].items()),
                       key=lambda u: (_sidst_under(art, u[0]), -u[1]))
        resultat.append(dict(kode=g["kode"], navn=g["navn"], dkk=g["dkk"], andel=andel(g["dkk"]),
                             stoerste_fond=max(g["fonde"], key=g["fonde"].get),
                             fonde=sorted(g["fonde"], key=lambda n: -g["fonde"][n]),
                             under=[dict(navn=n, dkk=v, andel=andel(v), stoerste_fond=fond)
                                    for n, v, fond in under]))
    med = [f for f in fonde if f.get(art)]
    return dict(
        total_dkk=total, grupper=resultat, foreloebig_dkk=foreloebig,
        foreloebig_andel=andel(foreloebig),
        foreloebige_fonde=[f["fund_name"] for f in med
                           if any(x["foreloebig"] for x in f[art]["linjer"])],
        # fordelinger opgjort på en anden dato end fondens NAV
        andre_datoer=[(f["fund_name"], f[art]["dato"]) for f in med
                      if f[art]["dato"] != f["nav_dato"]],
        # fonde, som forvalteren ikke fordeler; de står samlet efter deres strategi
        efter_strategi=[(f["fund_name"], f[art]["linjer"][0]["top"] if art == "brancher"
                         else f[art]["linjer"][0]["under"])
                        for f in med if f[art]["grundlag"] == "FUND_DEFAULT"],
        uden_fordeling=[f["fund_name"] for f in fonde if not f.get(art)])


EJENDOMSSEKTOR = "REAL_ESTATE"


def ejendomsandel(f):
    """Den del af fondens NAV, der ligger i GICS-sektoren Real Estate (0 uden branchefordeling)."""
    b = f.get("brancher")
    return sum(x["andel"] for x in b["linjer"] if x["top_kode"] == EJENDOMSSEKTOR) if b else 0.0


def opdel_ejendomme(fonde, art):
    """Investors NAV delt i ejendomme og resten, hver med sin egen base for andelene.

    Ejendomme er GICS-sektoren Real Estate på tværs af alle fonde: ejendomsfondene helt og
    ejendomsandelen af de øvrige. {ejendomme: {total_dkk, andel (af hele NAV), fonde: [{navn,
    dkk, andel (af ejendommene)}]}, resten: eksponering(...) af NAV uden ejendomme}. I
    brancherne fjernes sektoren og fondens øvrige sektorer skaleres op; i regionerne antages
    ejendomsandelen at ligge som resten af fonden, da fordelingerne ikke krydser branche og
    geografi. None uden DKK-tal for alle fonde eller uden ejendomme.
    """
    if not fonde or any(f["n"].get("nav_dkk") is None for f in fonde):
        return None
    ejendomme, resten = _uden_ejendomme(fonde, art)
    if ejendomme is None:
        return None
    return dict(ejendomme=ejendomme, resten=eksponering(resten, art))


def _uden_ejendomme(fonde, art):
    """(ejendomsdelen, fondene uden deres ejendomsandel) til opdelingerne.

    Ejendomsdelen er {total_dkk, andel (af hele NAV), fonde: [{navn, dkk, andel (af
    ejendommene)}], blandede}, eller None uden ejendomme. Resten er fondene med NAV uden
    ejendomsandelen; i brancherne er sektoren fjernet og de øvrige skaleret op.
    """
    andele = {f["fund_name"]: ejendomsandel(f) for f in fonde}
    ejendomme = sorted(((f["fund_name"], andele[f["fund_name"]] * f["n"]["nav_dkk"])
                        for f in fonde if andele[f["fund_name"]] > 0), key=lambda x: -x[1])
    i_alt = sum(dkk for _, dkk in ejendomme)
    resten = []
    for f in fonde:
        a = andele[f["fund_name"]]
        if a >= 1 - 1e-9:
            continue                    # en ren ejendomsfond: intet tilbage
        rest = f | {"n": f["n"] | {"nav_dkk": f["n"]["nav_dkk"] * (1 - a)}}
        if art == "brancher" and a > 0:
            b = f["brancher"]
            rest["brancher"] = b | {"linjer": [x | {"andel": x["andel"] / (1 - a)}
                                               for x in b["linjer"]
                                               if x["top_kode"] != EJENDOMSSEKTOR]}
        resten.append(rest)
    if i_alt <= 0:
        return None, resten
    total = sum(f["n"]["nav_dkk"] for f in fonde)
    # ejendomsandelen af hver fond som sin egen "fond", til ejendommenes fordeling
    ejd_fonde = [f | {"n": f["n"] | {"nav_dkk": andele[f["fund_name"]] * f["n"]["nav_dkk"]}}
                 for f in fonde if andele[f["fund_name"]] > 0]
    return dict(total_dkk=i_alt, andel=i_alt / total if total else None, art=art,
                fonde=[dict(navn=n, dkk=dkk, andel=dkk / i_alt) for n, dkk in ejendomme],
                # fonde, der kun delvis er ejendomme; deres andel nævnes i noten
                blandede=[n for n, _ in ejendomme if andele[n] < 1 - 1e-9],
                # i brancherne fordelt på ejendomssegment, i geografien som fondenes eksponering
                fordeling=(ejendomssegmenter(ejd_fonde) if art == "brancher"
                           else eksponering(ejd_fonde, art))), resten


# Ejendomsfondenes segment (fund.fund.re_sector, lukket af CK_fund_re_sector), på engelsk
# som GICS-navnene
EJENDOMSSEGMENTER = {"RESIDENTIAL": "Residential", "OFFICE": "Office", "LOGISTICS": "Logistics",
                     "RETAIL": "Retail", "HOTEL": "Hotel", "HEALTHCARE": "Healthcare",
                     "MIXED": "Mixed"}
SEGMENT_OEVRIGE = "Øvrige"     # ejendomsandelen af en fond uden ejendomssegment


def ejendomssegmenter(ejd_fonde):
    """Ejendommene fordelt på fondenes ejendomssegment (re_sector i rpt.v_fund_master).

    ejd_fonde er fondene med ejendomsandelen af deres NAV som nav_dkk. En fond uden segment
    (fx en buyout-fond med lidt ejendomme) står under SEGMENT_OEVRIGE, sidst.
    {total_dkk, grupper: [{kode, navn, dkk, andel, fonde}], uden_segment: [fondsnavne]}.
    """
    total = sum(f["n"]["nav_dkk"] for f in ejd_fonde)
    grupper = {}
    for f in ejd_fonde:
        kode = f.get("re_sektor")
        g = grupper.setdefault(kode, dict(kode=kode, navn=EJENDOMSSEGMENTER.get(
            kode, kode) if kode else SEGMENT_OEVRIGE, dkk=0.0, fonde=[]))
        g["dkk"] += f["n"]["nav_dkk"]
        g["fonde"].append(f["fund_name"])
    resultat = sorted(grupper.values(), key=lambda g: (g["kode"] is None, -g["dkk"]))
    for g in resultat:
        g["andel"] = g["dkk"] / total if total else None
    return dict(total_dkk=total, grupper=resultat,
                uden_segment=[f["fund_name"] for f in ejd_fonde if not f.get("re_sektor")])


PRIVATE_EQUITY, PRIVATE_CREDIT = "Private Equity", "Private Credit"


def opdel_aktivklasser(fonde, art):
    """Investors NAV delt i Private Equity, Private Credit og ejendomme, hver med sin egen base.

    Ejendomme er som i opdel_ejendomme(); Private Credit er resten af fondene med private
    credit som strategi (segmentet), Private Equity resten af de øvrige. {ejendomme: dict eller
    None, dele: [{navn, total_dkk, andel (af hele NAV), fordeling: eksponering(...)}]}; en del
    uden NAV udelades. None uden DKK-tal for alle fonde.
    """
    if not fonde or any(f["n"].get("nav_dkk") is None for f in fonde):
        return None
    ejendomme, resten = _uden_ejendomme(fonde, art)
    total = sum(f["n"]["nav_dkk"] for f in fonde)
    dele = []
    for navn, med in ((PRIVATE_EQUITY, lambda f: f.get("segment") != FOKUS["JUNIOR_DEBT"]),
                      (PRIVATE_CREDIT, lambda f: f.get("segment") == FOKUS["JUNIOR_DEBT"])):
        fonde_ = [f for f in resten if med(f)]
        i_alt = sum(f["n"]["nav_dkk"] for f in fonde_)
        if i_alt > 0:
            dele.append(dict(navn=navn, total_dkk=i_alt, andel=i_alt / total,
                             fordeling=eksponering(fonde_, art)))
    return dict(ejendomme=ejendomme, dele=dele)


def irr_dkk(fonde, kurs):
    """Porteføljens IRR i DKK, poolet over alle fondes cash flows, eller None.

    Hver cash flow omregnes med kursen på sin egen dato (som amount_dkk i rpt.v_cash_flow),
    og NAV er fondens NAV i DKK (kursen på rapportdatoen) på fondens NAV-dato. Så kan fonde i
    forskellige valutaer regnes sammen, og afkastet er det, investor har haft i kroner,
    valutakursændringer medregnet. kurs(valuta, dato) -> DKK pr. enhed eller None; mangler en
    kurs, er IRR None. {irr, flows: [{fund_code, fund_name, dato, valuta, beloeb, kurs, dkk,
    type}]} med flows i datoorden, til kontrollen i Excel.
    """
    flows = []
    for f in fonde:
        # irr_flows slutter med NAV på NAV-datoen (data.fond)
        *cf, (nav_dato, _) = f["irr_flows"]
        nav_dkk = f["n"].get("nav_dkk")
        if nav_dkk is None:
            return None
        for d, a in cf:
            k = kurs(f["valuta"], d)
            if k is None:
                return None
            flows.append(dict(fund_code=f["fund_code"], fund_name=f["fund_name"], dato=d,
                              valuta=f["valuta"], beloeb=a, kurs=k, dkk=a * k, type="Cash flow"))
        flows.append(dict(fund_code=f["fund_code"], fund_name=f["fund_name"], dato=nav_dato,
                          valuta=f["valuta"], beloeb=f["n"]["nav"],
                          kurs=nav_dkk / f["n"]["nav"] if f["n"]["nav"] else None,
                          dkk=nav_dkk, type="NAV (sidste cash flow)"))
    if not flows:
        return None
    flows.sort(key=lambda x: (x["dato"], x["type"] != "Cash flow"))
    return dict(irr=xirr([(x["dato"], x["dkk"]) for x in flows]), flows=flows)


def sammendrag(fonde, total_dkk, irr=None):
    """Porteføljens hovedtal til executive summary.

    Beløb og nøgletal for porteføljen er i DKK (total_dkk), så fonde i forskellige valutaer kan
    lægges sammen; uden DKK-tal er de None. irr er porteføljens IRR i DKK fra irr_dkk().
    """
    td = total_dkk or {}
    ind, udl, nav = td.get("indbetalt_dkk"), td.get("udloddet_dkk"), td.get("nav_dkk")
    vintager = [f["vintage"] for f in fonde if f.get("vintage")]
    med_tvpi = sorted((f for f in fonde if f["n"]["tvpi"] is not None),
                      key=lambda f: f["n"]["tvpi"])
    fond_ = lambda f: dict(fund_name=f["fund_name"], tvpi=f["n"]["tvpi"], irr=f["n"]["irr"])
    return dict(
        antal_fonde=len(fonde),
        antal_forvaltere=len({f["manager"] for f in fonde if f.get("manager")}),
        vintage=(min(vintager), max(vintager)) if vintager else None,
        **{k: td.get(k) for k in DKK_FELTER},
        picc=ind / td["tilsagn_dkk"] if ind is not None and td.get("tilsagn_dkk") else None,
        dpi=udl / ind if ind and udl is not None else None,
        tvpi=(udl + nav) / ind if ind and udl is not None and nav is not None else None,
        irr=irr,
        # det, investor samlet har stående i fondene: værdien i dag plus det, der kan kaldes
        eksponering_dkk=(td["resttilsagn_dkk"] + nav
                         if td.get("resttilsagn_dkk") is not None and nav is not None else None),
        bedst=fond_(med_tvpi[-1]) if med_tvpi else None,
        svagest=fond_(med_tvpi[0]) if len(med_tvpi) > 1 else None)


UDVIKLING_FELTER = ("primo", "kald", "udlodning", "vaerdi", "valuta", "ultimo", "naevner")


def omregn_udvikling_til_dkk(fonde, kurs):
    """Sæt DKK-beløbene (primo_dkk, kald_dkk, …, valuta_dkk) på hver fonds udvikling.

    kurs(valuta, dato) -> DKK pr. enhed eller None. Primo omregnes med kursen på primodatoen,
    hver cash flow med kursen på sin egen dato, værdiændringen og ultimo med kursen på
    rapportdatoen (som NAV i DKK på de øvrige sider). Valutakursændringen er residualen, så
    broen i DKK går op; for en DKK-fond er den 0. Mangler en kurs, er beløbene None.
    """
    for f in fonde:
        for u in f["udvikling"].values():
            kurser = [kurs(f["valuta"], d) for d in [u["primo_dato"], u["til"]]
                      + [d for d, _, _ in u["flows"]]]
            if any(k is None for k in kurser):
                u.update({f"{k}_dkk": None for k in UDVIKLING_FELTER})
                continue
            k_primo, k_til, k_flow = kurser[0], kurser[1], kurser[2:]
            dkk = [(d, a * k, kl) for (d, a, kl), k in zip(u["flows"], k_flow)]
            kald = sum(a for _, a, kl in dkk if kl == "kald")
            udlodning = -sum(a for _, a, kl in dkk if kl == "udlodning")
            primo, ultimo, vaerdi = u["primo"] * k_primo, u["ultimo"] * k_til, u["vaerdi"] * k_til
            dage = (u["til"] - u["primo_dato"]).days
            u.update(primo_dkk=primo, kald_dkk=kald, udlodning_dkk=udlodning, vaerdi_dkk=vaerdi,
                     ultimo_dkk=ultimo, valuta_dkk=ultimo - primo - kald + udlodning - vaerdi,
                     naevner_dkk=primo + (sum(a * (u["til"] - d).days / dage for d, a, _ in dkk)
                                          if dage else 0))


def fortegnsafvigelser(fonde):
    """[(fund_name, periode, årsag)], hvor den viste værdiændring og afkastet har modsat fortegn.

    Værdiændringen vises i DKK med valutakursændringen, afkastet er Modified Dietz i fondens
    valuta. årsag er "valuta", når værdiændringen i fondens valuta har samme fortegn som
    afkastet (valutakursen vender den), ellers "dietz" (en lille eller negativ kapitalbase,
    fx efter store udlodninger). Uden DKK-beløb eller med et nul springes perioden over.
    """
    ud = []
    for f in fonde:
        for p, u in f["udvikling"].items():
            vist = (None if u.get("vaerdi_dkk") is None
                    else u["vaerdi_dkk"] + u["valuta_dkk"])
            afkast = u.get("afkast")
            if not vist or not afkast or (vist > 0) == (afkast > 0):
                continue
            ud.append((f["fund_name"], p,
                       "valuta" if u["vaerdi"] and (u["vaerdi"] > 0) == (afkast > 0)
                       else "dietz"))
    return ud


def portefoelje_udvikling(fonde):
    """Porteføljens bro fra NAV primo til ultimo i DKK pr. periode, og hver fonds bidrag.

    {periode: {primo, kald, udlodning, vaerdi, valuta, ultimo, afkast, bidrag}} hvor bidrag
    er [(fund_name, værdiændring + valuta i DKK)] sorteret med det største først. Afkastet er
    Modified Dietz i DKK, poolet over fondene. En periode er None, hvis en fond mangler
    DKK-beløb.
    """
    resultat = {}
    for p in fonde[0]["udvikling"] if fonde else ():
        us = [f["udvikling"][p] for f in fonde]
        if any(u.get("primo_dkk") is None for u in us):
            resultat[p] = None
            continue
        s = {k: sum(u[f"{k}_dkk"] for u in us) for k in UDVIKLING_FELTER}
        bidrag = sorted(((f["fund_name"], f["udvikling"][p]["vaerdi_dkk"]
                          + f["udvikling"][p]["valuta_dkk"]) for f in fonde), key=lambda x: -x[1])
        naevner = s.pop("naevner")
        resultat[p] = dict(fra=us[0]["fra"], til=us[0]["til"], bidrag=bidrag,
                           afkast=(s["vaerdi"] + s["valuta"]) / naevner if naevner > 0 else None,
                           **s)
    return resultat


def _forkortelse_af(kort, lang):
    """Er kort en forkortelse af lang? "WPEF IX" af "Waterland Private Equity Fund IX".

    Forvalterne skriver samme fond forskelligt fra kvartal til kvartal ("Charlesbank X" /
    "Charlesbank Equity Fund X"). Sidste ord (fondsnummeret) skal være ens; de øvrige ord i
    kort skal i rækkefølge være begyndelsen af et ord i lang ("Tech" af "Technology") eller
    forbogstaverne i en række ord ("ACF" af "AlpInvest Co-Investment Fund"). Ord i lang må
    springes over.
    """
    k, l_ = kort.lower().split(), lang.lower().split()
    if len(k) < 2 or len(l_) < 2 or k[-1] != l_[-1] or k == l_:
        return False
    k, l_ = k[:-1], l_[:-1]

    def passer(i, j):
        if i == len(k):
            return True
        for start in range(j, len(l_)):
            if l_[start].startswith(k[i]) and passer(i + 1, start + 1):
                return True
            slut = start + len(k[i])
            if (len(k[i]) > 1 and slut <= len(l_)
                    and "".join(o[0] for o in l_[start:slut]) == k[i] and passer(i + 1, slut)):
                return True
        return False

    return passer(0, 0)


def _par(nu, primo):
    """{navn nu: navn primo} for beholdninger, der er den samme under to navne.

    Kun beholdninger uden navnefælle på den anden dato parres, med samme vintage, og kun hvor
    præcis én kandidat passer, så et tvivlsomt par hellere står som ny og afhændet.
    """
    rest_nu = [r for r in nu if r["navn"] not in {x["navn"] for x in primo}]
    rest_primo = [r for r in primo if r["navn"] not in {x["navn"] for x in nu}]
    par = {}
    for a in rest_nu:
        kandidater = [b for b in rest_primo
                      if (a["vintage"] is None or b["vintage"] is None or a["vintage"] == b["vintage"])
                      and (_forkortelse_af(a["navn"], b["navn"])
                           or _forkortelse_af(b["navn"], a["navn"]))]
        if len(kandidater) == 1 and kandidater[0]["navn"] not in par.values():
            par[a["navn"]] = kandidater[0]["navn"]
    return par


def _vaerdiskabt(r):
    """Gevinsten på en beholdning: dagsværdi + udloddet - kostpris (investors andel)."""
    return r["dagsvaerdi"] + r["udloddet"] - (r["kostpris"] or 0)


def _taller(r):
    return (r["dagsvaerdi"], r["udloddet"], r["kostpris"])


def vaerdibevaegelser(fonde, kurser):
    """Beholdningernes værdibevægelse i kvartalet og året, i DKK på rapportdatoens kurs.

    Bevægelsen er ændringen i dagsværdi + udloddet - kostpris mellem beholdningerne primo
    (f["beholdninger_primo"][periode]) og nu, så nye investeringer og udlodninger ikke tæller
    som værdi. Kun beholdninger, der findes på begge datoer (evt. under to navne, se _par),
    og nye beholdninger med closing i perioden får en bevægelse; resten - afhændede, omdøbte
    og nye uden closing i perioden - kan ikke sammenlignes og står i usammenlignelige, for
    en afhændet beholdning forsvinder fra listen med sine udlodninger. Fonde, hvis
    beholdninger ikke er skaleret til investors andel, står i uskalerede; fonde uden nye
    beholdninger siden periodens start, eller hvor alle tal er uændrede, står i uaendrede.
    None, hvis ingen fond har beholdninger.
    """
    med = [f for f in fonde if f.get("beholdninger")]
    if not med:
        return None
    linjer, uaendrede, usammenlignelige = {}, {}, {}
    ejendomsfonde = [f for f in med if f["beholdninger"]["type"] == "ejendomme"]
    brugte = [f for f in med if f["beholdninger"]["skaleret"]
              and (kurser.get(f["valuta"]) or (None,))[0] is not None]
    for f in brugte:
        kurs = kurser[f["valuta"]][0]
        nu_r = f["beholdninger"]["raekker"]
        for p, primo_beh in f["beholdninger_primo"].items():
            fra = f["udvikling"][p]["fra"]
            primo_r = (primo_beh or {}).get("raekker", [])
            # linjen bærer det nyeste navn; primo-linjer under et ældre navn omdøbes til det
            nyt_navn = {gl: ny for ny, gl in _par(nu_r, primo_r).items()}
            # store og små bogstaver skifter også mellem rapporterne ("DiversiTech"/"Diversitech")
            primo = {nyt_navn.get(r["navn"], r["navn"]).lower(): r for r in primo_r}
            nu_navne = {r["navn"].lower() for r in nu_r}
            faelles = [r for r in nu_r if r["navn"].lower() in primo]
            if f["beholdninger"]["dato"] <= fra or (
                    faelles and len(faelles) == len(nu_r) == len(primo)
                    and all(_taller(r) == _taller(primo[r["navn"].lower()]) for r in faelles)):
                uaendrede.setdefault(p, []).append(f["fund_name"])
                continue
            for r in nu_r:
                linje = linjer.setdefault((f["fund_name"], r["navn"]),
                                          dict(navn=r["navn"], fond=f["fund_name"]))
                if r["navn"].lower() in primo:
                    linje[p] = (_vaerdiskabt(r) - _vaerdiskabt(primo[r["navn"].lower()])) * kurs
                elif r["closing"] and r["closing"] > fra:
                    linje[p] = _vaerdiskabt(r) * kurs
                else:
                    usammenlignelige.setdefault(p, []).append((f["fund_name"], r["navn"]))
            usammenlignelige.setdefault(p, []).extend(
                (f["fund_name"], x["navn"]) for n, x in primo.items() if n not in nu_navne)
    return dict(linjer=list(linjer.values()), uaendrede=uaendrede,
                usammenlignelige={p: sorted(set(x)) for p, x in usammenlignelige.items() if x},
                dato=max(f["beholdninger"]["dato"] for f in med),
                uskalerede=[f["fund_name"] for f in med
                            if f not in brugte and f not in ejendomsfonde],
                ejendomsfonde=[f["fund_name"] for f in ejendomsfonde],
                uden_beholdninger=[f["fund_name"] for f in fonde if not f.get("beholdninger")])


def _dokument_ider(fonde):
    """Alle dokument-id'er bag fondenes tal: NAV, cash flows, kapitalkonto, tilsagn, beholdninger."""
    ider = set()
    for f in fonde:
        k = f["kilder"]
        for df in (k["nav"], k["cash_flows"], k["kapitalkonto"], k["tilsagn"]):
            ider |= {int(x) for x in df.document_id.dropna()}
        ider |= {d for _, d in (k.get("fonds_nav") or {}).values() if d}
        if k["fondsstoerrelse"]["document_id"]:
            ider.add(k["fondsstoerrelse"]["document_id"])
        for beh in [f["beholdninger"], *(f.get("beholdninger_primo") or {}).values()]:
            ider |= {r["document_id"] for r in (beh or {}).get("raekker", []) if r["document_id"]}
        for art in ARTER:
            ider |= set((f.get(art) or {}).get("dokumenter", []))
    return sorted(ider)


def dokumenter(ider):
    """Kildedokumenterne fra fund.document, ét pr. id: type, fil, sti, status og godkendelse."""
    if not ider:
        return pd.DataFrame(columns=["document_id"])
    return read_sql(f"""
        SELECT d.document_id, f.fund_code, d.doc_type, d.period_end, d.value_date, d.notice_date,
               d.external_ref, d.file_name, d.rel_path, d.page_count, d.status,
               d.extraction_method, d.approved_by, d.approved_at, d.note
        FROM fund.document d
        LEFT JOIN fund.fund f ON f.fund_id = d.fund_id
        WHERE d.document_id IN ({", ".join(str(int(i)) for i in ider)})
        ORDER BY f.fund_code, d.doc_type, COALESCE(d.period_end, d.value_date)""")


def investor_rapport(investor_code, as_of=None):
    """{investor, navn, as_of, fonde: [...]} — as_of er som standard seneste NAV-dato."""
    as_of = _dato(as_of) if as_of else date.today()
    info = fund_info()
    fonde = [f for f in (fond(inv, as_of, info) for inv in _investeringer(investor_code).itertuples())
             if f is not None]
    if not fonde:
        raise ValueError(f"Ingen NAV-data for '{investor_code}' pr. {as_of}")
    fonde.sort(key=strategi_noegle)     # alle slides og Excel viser fondene i denne rækkefølge
    rapport_dato = max(f["nav_dato"] for f in fonde)
    for f in fonde:
        if f["nav_dato"] < rapport_dato:
            f["noter"].insert(0, Note(f"Seneste NAV er pr. {f['nav_dato']:%d.%m.%Y}.", "nav"))
    kurser = {v: _dkk_kurs(v, rapport_dato) for v in sorted({f["valuta"] for f in fonde})}
    omregn_til_dkk(fonde, kurser)
    total_dkk = dkk_total(fonde)

    rapport_perioder = perioder(rapport_dato)
    kurscache = {}

    def kurs(valuta, dato):
        if (valuta, dato) not in kurscache:
            kurscache[valuta, dato] = _dkk_kurs(valuta, dato)
        return (kurscache[valuta, dato] or (None,))[0]

    for f in fonde:
        udvikling(f, rapport_perioder)
        # primo med samme andel som ultimo; for en feeder er det den underliggende fond
        beh = f["beholdninger"]
        f["beholdninger_primo"] = {
            p: _beholdninger(beh["kode"], fra, beh["andel"] if beh["skaleret"] else None)
            if beh else None
            for p, (fra, _) in rapport_perioder.items()}
    omregn_udvikling_til_dkk(fonde, kurs)
    irr_i_dkk = irr_dkk(fonde, kurs)
    return dict(investor=investor_code, navn=investor_navn(investor_code),
                as_of=rapport_dato, fonde=fonde, total=total(fonde),
                kurser=kurser, total_dkk=total_dkk, irr_dkk=irr_i_dkk,
                sammendrag=sammendrag(fonde, total_dkk, irr_i_dkk and irr_i_dkk["irr"]),
                valutafordeling=valutafordeling(fonde),
                fordelinger={k: fordeling(fonde, k) for k in ("geografi", "segment")},
                brancher=eksponering(fonde, "brancher"), benchmark=benchmark(),
                regioner=eksponering(fonde, "regioner"),
                # samme fordelinger med ejendommene for sig og resten på sin egen base
                uden_ejendomme={art: opdel_ejendomme(fonde, art)
                                for art in ("brancher", "regioner")},
                # og delt i Private Equity, Private Credit og ejendomme
                aktivklasser={art: opdel_aktivklasser(fonde, art)
                              for art in ("brancher", "regioner")},
                gennemlysning=gennemlysning(fonde),
                stoerste_eksponering=stoerste_eksponering(fonde, kurser),
                stoerste_positioner=stoerste_positioner(fonde, kurser),
                perioder=rapport_perioder, udvikling=portefoelje_udvikling(fonde),
                fortegnsafvigelser=fortegnsafvigelser(fonde),
                vaerdibevaegelser=vaerdibevaegelser(fonde, kurser),
                datotjek=datotjek(fonde, rapport_dato),
                valuta=fonde[0]["valuta"] if len({f["valuta"] for f in fonde}) == 1 else None,
                # (valuta, ønsket dato) -> (kurs, kursdato) for hver kurs, rapporten har brugt
                kursopslag={**{(v, rapport_dato): k for v, k in kurser.items()}, **kurscache},
                dokumenter=dokumenter(_dokument_ider(fonde)))
