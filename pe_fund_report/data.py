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
               f.vintage_year, m.name AS manager_name, m.manager_code
        FROM fund.investment x
        JOIN fund.investor i ON i.investor_id = x.investor_id
        JOIN fund.fund f ON f.fund_id = x.fund_id
        JOIN fund.manager m ON m.manager_id = f.manager_id
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
        rows.append(dict(navn=r.name, vintage=None if pd.isna(r.vintage) else int(r.vintage),
                         sektor=r.segment, closing=None if pd.isna(r.trade_date) else _dato(r.trade_date),
                         kostpris=None if kost is None else kost * andel,
                         dagsvaerdi=fv * andel, udloddet=udl * andel, total=(fv + udl) * andel,
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
        noter.append(f"Indbetalt kapital er GP'ens opgørelse ({gp_indbetalt / 1000:,.0f} tDKK)"
                     f"; de registrerede kapitalkald summerer til {cf_indbetalt / 1000:,.0f} tDKK."
                     .replace(",", "."))

    est = [(_dato(e["dato"]), float(e["beloeb"])) for e in info.get("estimerede_cashflows", [])
           if _dato(e["dato"]) <= nav_dato]
    if est:
        noter.append("IRR og historik medregner kapitalkald med skønnet dato, hvor meddelelsen "
                     "ikke er modtaget: "
                     + "; ".join(e["note"] for e in info["estimerede_cashflows"]) + ".")

    nav_nu = float(nav.nav.iloc[-1])
    irr_flows = [(r.value_date, r.amount) for r in cf.itertuples()
                 if r.flow_class != "WITHHOLDING_TAX"]
    irr_flows = irr_flows + est + [(nav_dato, nav_nu)]
    irr = xirr(irr_flows)

    # ejerandel: tilsagn / fondens samlede tilsagn (databasen, ellers config/fund_info.yaml)
    tc = read_sql("""SELECT TOP 1 value FROM rpt.v_fund_metric
                     WHERE fund_code = :f AND metric_code = 'TOTAL_COMMITMENTS' AND report_date <= :d
                     ORDER BY report_date DESC""", f=kode, d=nav_dato)
    fondsstr = _f(tc.value[0]) if not tc.empty else _f(info.get("fondsstoerrelse"))
    andel = tilsagn / fondsstr if tilsagn and fondsstr else None

    n = dict(tilsagn=tilsagn, indbetalt=indbetalt, udloddet=udloddet, nav=nav_nu,
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

    if not senere.empty:
        dele = []
        for kl, tekst in (("CONTRIBUTION", "kapitalkald"), ("DISTRIBUTION", "udlodning")):
            s = senere[senere.flow_class == kl]
            if not s.empty:
                dele.append(f"{tekst} {abs(s.amount.sum()) / 1000:,.0f} tDKK".replace(",", ".")
                            + f" ({', '.join(f'{d:%d.%m.%Y}' for d in sorted(set(s.value_date)))})")
        if dele:
            noter.append("Efter NAV-datoen: " + " og ".join(dele) + " – ikke medregnet.")
    if info.get("status") == "udkast":
        noter.append("Fondsbeskrivelse, geografi og segment er et udkast og ikke gennemgået.")

    return dict(
        fund_code=kode, fund_name=inv.fund_name, manager=inv.manager_name,
        logo=_logo(inv.manager_code), valuta=inv.currency,
        strategi=inv.strategy, vintage=None if pd.isna(inv.vintage_year) else int(inv.vintage_year),
        beskrivelse=info.get("beskrivelse"), geografi=info.get("geografi"),
        segment=info.get("segment"), nav_dato=nav_dato, n=n, historik=historik,
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
                resttilsagn=s("resttilsagn"),
                picc=indbetalt / tilsagn if tilsagn else None,
                dpi=udloddet / indbetalt if indbetalt else None,
                tvpi=(udloddet + nav) / indbetalt if indbetalt else None,
                irr=xirr([x for f in fonde for x in f["irr_flows"]]))


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
    return dict(investor=investor_code, navn=investor_navn(investor_code),
                as_of=rapport_dato, fonde=fonde, total=total(fonde),
                valuta=fonde[0]["valuta"] if len({f["valuta"] for f in fonde}) == 1 else None)
