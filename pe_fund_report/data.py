"""Henter og beregner alt, hvad investorrapporten viser, for én investor.

Reglerne følger docs/analysis-data-guide.md: rpt-views først, LP-fortegn på cash flows
(betalt < 0), EQUALISATION holdes ude af indbetalt/udloddet, WITHHOLDING_TAX ude af IRR,
og basistabeller filtreres på APPROVED + is_active = 1. Investor- og fondsniveau blandes
ikke: beholdningerne er hele fondens tal og skaleres med investors ejerandel.
"""
from datetime import date

import pandas as pd
import yaml

from .db import read_sql
from .paths import FUND_INFO

GP_TOLERANCE = 0.005      # afvigelse mellem cash flows og GP'ens indbetalte kapital, før GP-tallet bruges

# Geografi og segment kommer fra rpt.v_fund_master som koder; databasens check constraints
# på fund.fund lukker kodelisterne, så hver kode har præcis én dansk tekst her. Den
# geografiske fordeling grupperer på teksten, så to stavemåder ville blive to grupper.
GEOGRAFIER = {"DENMARK": "Danmark", "NORDICS": "Norden", "EUROPE": "Europa",
              "NORTH_AMERICA": "Nordamerika", "EUROPE_NORTH_AMERICA": "Europa og Nordamerika",
              "GLOBAL": "Global"}
STRUKTURER = {"FUND_OF_FUNDS": "Fund of funds", "SECONDARY": "Secondary fund of funds"}
FOKUS = {"BUYOUT": "Buyout", "VENTURE": "Venture", "GROWTH": "Growth",
         "JUNIOR_DEBT": "Private credit", "SHIPPING": "Shipping",
         "MULTI_STRATEGY": "Multistrategi"}
STOERRELSER = {"SMALL_MID": "Small/mid-cap", "MID_LARGE": "Mid/large-cap"}


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

    Fund of funds og secondaries beskrives ved strukturen; de direkte fonde ved fokus,
    og buyout-fonde får størrelsesfokus foran ("Small/mid-cap buyout").
    """
    struktur, fokus, stoerrelse = _kode(fund_structure), _kode(strategy_focus), _kode(size_focus)
    if struktur in STRUKTURER:
        return STRUKTURER[struktur]
    if fokus is None:
        return None
    tekst = FOKUS.get(fokus, fokus)
    if fokus == "BUYOUT" and stoerrelse in STOERRELSER:
        return f"{STOERRELSER[stoerrelse]} buyout"
    return tekst


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


def investor_navn(investor_code):
    df = read_sql("SELECT legal_name FROM fund.investor WHERE investor_code = :c", c=investor_code)
    if df.empty:
        raise ValueError(f"Investoren '{investor_code}' findes ikke i fund.investor")
    return df.legal_name[0]


def _investeringer(investor_code):
    return read_sql("""
        SELECT x.investment_id, f.fund_code, f.name AS fund_name, f.strategy, f.currency,
               f.vintage_year,
               -- vores feeder-selskaber står med os selv som forvalter; vis den forvalter,
               -- der forvalter den underliggende fond, som pengene reelt ligger i
               COALESCE(u.manager_name, m.name) AS manager_name,
               COALESCE(u.manager_code, m.manager_code) AS manager_code,
               fm.geography_focus, fm.fund_structure, fm.strategy_focus, fm.size_focus,
               fm.term_years, fm.extension_years, fm.total_commitments
        FROM fund.investment x
        JOIN fund.investor i ON i.investor_id = x.investor_id
        JOIN fund.fund f ON f.fund_id = x.fund_id
        JOIN fund.manager m ON m.manager_id = f.manager_id
        LEFT JOIN rpt.v_fund_master fm ON fm.fund_code = f.fund_code
        LEFT JOIN rpt.v_fund_master u ON u.fund_code = fm.underlying_fund_code
        WHERE i.investor_code = :c AND x.is_active = 1
        ORDER BY f.name""", c=investor_code)


def _logo(manager_code):
    """Forvalterens logo fra rpt.v_manager_logo; helst udgaven til lys baggrund. None, hvis intet."""
    df = read_sql("""
        SELECT background, width_px, height_px, image FROM rpt.v_manager_logo
        WHERE manager_code = :c
        ORDER BY CASE background WHEN 'LIGHT' THEN 0 ELSE 1 END""", c=manager_code)
    if df.empty:
        return None
    r = df.iloc[0]
    return dict(billede=bytes(r.image), bredde=int(r.width_px), hoejde=int(r.height_px),
                moerk_baggrund=r.background == "DARK")


def _kapitalkonto(investment_id, as_of):
    return read_sql("""
        SELECT ca.period_end, ca.basis, a.account_code, ca.amount
        FROM fund.capital_account ca
        JOIN fund.document d ON d.document_id = ca.document_id AND d.status = 'APPROVED'
        JOIN fund.account a ON a.account_id = ca.account_id
        WHERE ca.is_active = 1 AND ca.investment_id = :i AND ca.period_end <= :d""",
        i=investment_id, d=as_of)


def _gp_indbetalt(ka, nav_dato):
    """GP'ens akkumulerede indbetalinger: CA110 ITD på NAV-datoen, ellers seneste MM110."""
    itd = ka[(ka.account_code == "CA110") & (ka.basis == "ITD") & (ka.period_end == nav_dato)]
    if not itd.empty:
        return _f(itd.amount.iloc[0])
    mm = ka[ka.account_code == "MM110"].sort_values("period_end")
    return _f(mm.amount.iloc[-1]) if not mm.empty else None


def _memo(ka, nav_dato, konto):
    r = ka[(ka.account_code == konto) & (ka.period_end == nav_dato)]
    return _f(r.amount.iloc[0]) if not r.empty else None


def _vaerdibro(ka, nav_dato, n):
    """Fra indbetalt til NAV. Værdiskabelsen er residualen, så broen altid stemmer."""
    nu = ka[ka.period_end == nav_dato]
    itd = nu[nu.basis == "ITD"].set_index("account_code").amount
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
           ("Udlodninger", -n["udloddet"]),
           ("Værdiskabelse før carry" if gebyrer is not None
            else "Værdiskabelse efter omkostninger, før carry", vaerdi)]
    if gebyrer is not None:
        bro.append(("Forvaltningsgebyr og omkostninger", gebyrer))
    bro += [("Carried interest", carry), ("NAV", n["nav"])]
    return bro


def _beholdninger(fund_code, as_of, andel):
    """Fondens beholdninger på seneste rapportdato <= as_of.

    Skaleres til investors andel, når ejerandelen kendes; ellers vises hele fondens tal
    (skaleret = False), så siden ikke står tom, blot fordi fondsstørrelsen mangler.
    """
    h = read_sql("""
        SELECT h.report_date, h.section, h.name, h.segment, h.trade_date, h.vintage,
               h.commitment, h.invested, h.distributions,
               h.fair_value, h.paid_in_pct, h.multiple, h.irr_pct
        FROM rpt.v_portfolio_holding h
        WHERE h.fund_code = :f AND h.report_date =
              (SELECT MAX(report_date) FROM rpt.v_portfolio_holding
               WHERE fund_code = :f AND report_date <= :d)""", f=fund_code, d=as_of)
    if h.empty:
        return None
    skaleret, andel = andel is not None, andel or 1.0
    rows = []
    for r in h.itertuples():
        kost = _f(r.invested)
        if kost is None and _f(r.commitment) is not None and _f(r.paid_in_pct) is not None:
            kost = _f(r.commitment) * _f(r.paid_in_pct) / 100
        fv, udl = _f(r.fair_value) or 0.0, _f(r.distributions) or 0.0
        # *_fond er hele fondens tal; virksomhedsprofilerne viser dem, som GP'en rapporterer dem,
        # mens beholdningstabellen viser investors andel.
        rows.append(dict(navn=r.name, vintage=None if pd.isna(r.vintage) else int(r.vintage),
                         sektor=r.segment, closing=None if pd.isna(r.trade_date) else _dato(r.trade_date),
                         kostpris=None if kost is None else kost * andel,
                         dagsvaerdi=fv * andel, udloddet=udl * andel, total=(fv + udl) * andel,
                         kostpris_fond=kost, dagsvaerdi_fond=fv, udloddet_fond=udl,
                         tvpi=_f(r.multiple), irr=_f(r.irr_pct)))
    rows.sort(key=lambda x: -x["total"])
    # ATP/KAPE-typen (section FUND) er underliggende fonde; resten er porteføljeselskaber
    return dict(dato=_dato(h.report_date.iloc[0]), raekker=rows, skaleret=skaleret,
                type="fonde" if (h.section == "FUND").all() else "selskaber")


def fond(inv, as_of, info):
    """Alle tal til én fondsside. inv er en række fra _investeringer()."""
    iid, kode = int(inv.investment_id), inv.fund_code
    info = info.get(kode, {}) or {}
    noter = []


    nav = read_sql("""SELECT period_end, COALESCE(restated_nav, nav) AS nav FROM rpt.v_nav
                      WHERE investment_id = :i AND period_end <= :d ORDER BY period_end""",
                   i=iid, d=as_of)
    if nav.empty:
        return None
    nav["period_end"] = nav.period_end.map(_dato)
    nav_dato = nav.period_end.iloc[-1]

    cf = read_sql("""SELECT value_date, flow_class, is_recallable, amount, flow_ref
                     FROM rpt.v_cash_flow WHERE investment_id = :i""", i=iid)
    cf["value_date"] = cf.value_date.map(_dato)
    cf["amount"] = cf.amount.astype(float)
    senere, cf = cf[cf.value_date > nav_dato], cf[cf.value_date <= nav_dato]

    tilsagn = _f(read_sql("""SELECT SUM(amount) AS s FROM fund.commitment_event
                             WHERE investment_id = :i AND event_date <= :d""", i=iid, d=nav_dato).s[0])
    cf_indbetalt = -cf[cf.flow_class == "CONTRIBUTION"].amount.sum()
    udloddet = cf[cf.flow_class == "DISTRIBUTION"].amount.sum()
    recallable = cf[(cf.flow_class == "DISTRIBUTION") & (cf.is_recallable == 1)].amount.sum()

    ka = _kapitalkonto(iid, nav_dato)
    ka["period_end"] = ka.period_end.map(_dato)
    gp_indbetalt = _gp_indbetalt(ka, nav_dato)
    indbetalt = cf_indbetalt
    if gp_indbetalt and abs(gp_indbetalt - cf_indbetalt) > GP_TOLERANCE * gp_indbetalt:
        indbetalt = gp_indbetalt
        noter.append(f"Indbetalt kapital følger forvalterens opgørelse "
                     f"({gp_indbetalt / 1000:,.0f} t{inv.currency}); de modtagne kapitalkald "
                     f"summerer til {cf_indbetalt / 1000:,.0f} t{inv.currency}."
                     .replace(",", "."))

    est = [(_dato(e["dato"]), float(e["beloeb"])) for e in info.get("estimerede_cashflows", [])
           if _dato(e["dato"]) <= nav_dato]
    if est:
        noter.append("IRR og historik medregner kapitalkald med skønnet dato, hvor "
                     "meddelelsen ikke er modtaget: "
                     + "; ".join(e["note"] for e in info["estimerede_cashflows"]) + ".")

    nav_nu = float(nav.nav.iloc[-1])
    irr_flows = [(r.value_date, r.amount) for r in cf.itertuples()
                 if r.flow_class != "WITHHOLDING_TAX"]
    irr_flows = irr_flows + est + [(nav_dato, nav_nu)]
    irr = xirr(irr_flows)

    # ejerandel: tilsagn / fondens samlede tilsagn; GP'ens dateret tal (TOTAL_COMMITMENTS),
    # ellers stamdataenes fondsstørrelse i rpt.v_fund_master
    tc = read_sql("""SELECT TOP 1 value FROM rpt.v_fund_metric
                     WHERE fund_code = :f AND metric_code = 'TOTAL_COMMITMENTS' AND report_date <= :d
                     ORDER BY report_date DESC""", f=kode, d=nav_dato)
    fondsstr = _f(tc.value[0]) if not tc.empty else _f(inv.total_commitments)
    andel = tilsagn / fondsstr if tilsagn and fondsstr else None

    n = dict(tilsagn=tilsagn, indbetalt=indbetalt, udloddet=udloddet, nav=nav_nu,
             recallable=float(recallable),
             resttilsagn=(tilsagn - indbetalt + recallable) if tilsagn else None,
             picc=indbetalt / tilsagn if tilsagn else None,
             dpi=udloddet / indbetalt if indbetalt else None,
             tvpi=(udloddet + nav_nu) / indbetalt if indbetalt else None,
             irr=irr, andel=andel, fondsstoerrelse=fondsstr,
             gp_irr=_memo(ka, nav_dato, "MM150"), gp_tvpi=_memo(ka, nav_dato, "MM160"))

    # kvartalshistorik: NAV og akkumulerede cash flows (inkl. de skønnede kapitalkald)
    ind = [(d, -a) for d, a in zip(cf[cf.flow_class == "CONTRIBUTION"].value_date,
                                   cf[cf.flow_class == "CONTRIBUTION"].amount)]
    ind += [(d, -a) for d, a in est]
    udl = list(zip(cf[cf.flow_class == "DISTRIBUTION"].value_date,
                   cf[cf.flow_class == "DISTRIBUTION"].amount))
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
        noter.append(f"Resttilsagnet indeholder genkaldelige udlodninger på {beloeb} "
                     f"t{inv.currency}, som forvalteren kan kalde igen.")

    if not senere.empty:
        dele = []
        for kl, tekst in (("CONTRIBUTION", "kapitalkald"), ("DISTRIBUTION", "udlodning")):
            s = senere[senere.flow_class == kl]
            if not s.empty:
                dele.append(f"{tekst} {abs(s.amount.sum()) / 1000:,.0f} t{inv.currency}"
                            .replace(",", ".")
                            + f" ({', '.join(f'{d:%d.%m.%Y}' for d in sorted(set(s.value_date)))})")
        if dele:
            noter.append("Efter NAV-datoen: " + " og ".join(dele) + " – ikke medregnet.")
    if info.get("status") == "udkast":
        noter.append("Fondsbeskrivelsen er foreløbig.")

    return dict(
        fund_code=kode, fund_name=inv.fund_name, manager=inv.manager_name,
        logo=_logo(inv.manager_code), valuta=inv.currency,
        strategi=inv.strategy, vintage=None if pd.isna(inv.vintage_year) else int(inv.vintage_year),
        beskrivelse=info.get("beskrivelse"), geografi=geografi(inv.geography_focus),
        segment=segment(inv.fund_structure, inv.strategy_focus, inv.size_focus),
        loebetid=loebetid(inv.term_years, inv.extension_years),
        benchmark=info.get("benchmark"), virksomheder=info.get("virksomheder") or {},
        nav_dato=nav_dato, n=n, historik=historik,
        bro=_vaerdibro(ka, nav_dato, n), beholdninger=_beholdninger(kode, nav_dato, andel),
        irr_flows=irr_flows, noter=noter)


def total(fonde):
    """Porteføljens samlede nøgletal; IRR er poolet over alle fondes cash flows.

    Beløb i forskellige valutaer lægges ikke sammen: None, hvis fondene ikke har samme valuta.
    """
    if len({f["valuta"] for f in fonde}) != 1:
        return None
    s = lambda k: sum(f["n"][k] or 0 for f in fonde)
    tilsagn, indbetalt, udloddet, nav = s("tilsagn"), s("indbetalt"), s("udloddet"), s("nav")
    return dict(tilsagn=tilsagn, indbetalt=indbetalt, udloddet=udloddet, nav=nav,
                recallable=s("recallable"), resttilsagn=s("resttilsagn"),
                picc=indbetalt / tilsagn if tilsagn else None,
                dpi=udloddet / indbetalt if indbetalt else None,
                tvpi=(udloddet + nav) / indbetalt if indbetalt else None,
                irr=xirr([x for f in fonde for x in f["irr_flows"]]))


def gennemlysning(fonde):
    """Porteføljens beholdninger samlet i én liste, opgjort som investors andel.

    Linjerne blandes på tværs af fondene, så underliggende fonde og porteføljeselskaber står
    side om side; "type" på hver linje siger hvad den er. None, hvis fondene ikke har samme
    valuta, eller ingen fond har beholdninger.
    """
    if len({f["valuta"] for f in fonde}) != 1:
        return None
    med = [f for f in fonde if f.get("beholdninger")]
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

    {total, grupper: [{navn, fonde, sum, andel}], andele: {fund_name: andel}}. Grupperne
    sorteres efter tilsagn i DKK. Andelene er None, hvis en fond mangler en DKK-kurs.
    """
    total = _gruppesum(fonde)
    i_alt = total["tilsagn_dkk"]
    andel = lambda v: v / i_alt if i_alt and v is not None else None
    grupper = {}
    for f in fonde:
        grupper.setdefault(f[noegle] or "Ikke oplyst", []).append(f)
    resultat = [dict(navn=navn, fonde=g, sum=(s := _gruppesum(g)), andel=andel(s["tilsagn_dkk"]))
                for navn, g in grupper.items()]
    resultat.sort(key=lambda g: -(g["sum"]["tilsagn_dkk"] or 0))
    return dict(total=total | {"andel": andel(i_alt)}, grupper=resultat,
                andele={f["fund_name"]: andel(f["n"].get("tilsagn_dkk")) for f in fonde})


def sammendrag(fonde, total_dkk):
    """Porteføljens hovedtal til executive summary.

    Beløb og nøgletal for porteføljen er i DKK (total_dkk), så fonde i forskellige valutaer kan
    lægges sammen; uden DKK-tal er de None. IRR for porteføljen beregnes ikke her: der findes
    ingen cash flow-serie i DKK, så IRR vises kun pr. fond.
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
        bedst=fond_(med_tvpi[-1]) if med_tvpi else None,
        svagest=fond_(med_tvpi[0]) if len(med_tvpi) > 1 else None)


def investor_rapport(investor_code, as_of=None):
    """{investor, navn, as_of, fonde: [...]} — as_of er som standard seneste NAV-dato."""
    as_of = _dato(as_of) if as_of else date.today()
    info = fund_info()
    fonde = [f for f in (fond(inv, as_of, info) for inv in _investeringer(investor_code).itertuples())
             if f is not None]
    if not fonde:
        raise ValueError(f"Ingen NAV-data for '{investor_code}' pr. {as_of}")
    rapport_dato = max(f["nav_dato"] for f in fonde)
    for f in fonde:
        if f["nav_dato"] < rapport_dato:
            f["noter"].insert(0, f"Seneste NAV er pr. {f['nav_dato']:%d.%m.%Y}.")
    kurser = {v: _dkk_kurs(v, rapport_dato) for v in sorted({f["valuta"] for f in fonde})}
    omregn_til_dkk(fonde, kurser)
    total_dkk = dkk_total(fonde)
    return dict(investor=investor_code, navn=investor_navn(investor_code),
                as_of=rapport_dato, fonde=fonde, total=total(fonde),
                kurser=kurser, total_dkk=total_dkk, sammendrag=sammendrag(fonde, total_dkk),
                fordelinger={k: fordeling(fonde, k) for k in ("geografi", "segment")},
                gennemlysning=gennemlysning(fonde),
                datotjek=datotjek(fonde, rapport_dato),
                valuta=fonde[0]["valuta"] if len({f["valuta"] for f in fonde}) == 1 else None)
