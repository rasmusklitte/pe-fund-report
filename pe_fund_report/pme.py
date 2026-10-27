"""Public market equivalent: fondens cash flows målt mod et børsindeks (rene funktioner).

Alle mål bygger på samme idé, en indekskonto: hvert kapitalkald køber indekset, og hver
udlodning sælger det, så kontoen har præcis fondens ind- og udbetalinger.
- KS-PME (Kaplan og Schoar 2005): udlodninger og NAV forrentet med indekset til
  opgørelsesdatoen i forhold til kapitalkaldene forrentet på samme måde. Over 1 har fonden
  slået indekset.
- Direct Alpha (Gredil, Griffiths og Stucke 2014): IRR af de samme forrentede cash flows,
  dvs. det årlige merafkast over indekset.
- Indekskontoens værdi og IRR (Long og Nickels 1996): det, fondens NAV og IRR skal
  sammenlignes med. NAV minus kontoens værdi er merværdien mod indekset i kroner.

Et indeks er her en funktion dato -> niveau (eller None), så valutaomregningen kan lægges
uden på. Ingen database, intet yfinance.
"""
from bisect import bisect_right


def opslag(serie, tolerance=5):
    """Seneste niveau på eller før datoen fra en pd.Series (dato -> niveau, stigende).

    None før seriens start og mere end tolerance dage efter dens slutning, så et forældet
    indeks ikke stille bruges som dagens.
    """
    if serie is None or serie.empty:
        return lambda d: None
    datoer, niveauer = list(serie.index), [float(v) for v in serie.to_numpy()]

    def niveau(d):
        i = bisect_right(datoer, d)
        if i == 0 or (d - datoer[-1]).days > tolerance:
            return None
        return niveauer[i - 1]
    return niveau


def omregnet(indeks, kurs_fra, kurs_til):
    """Indekset i en anden valuta. kurs_fra/kurs_til: dato -> DKK pr. enhed af indeksets
    valuta / målvalutaen; None betyder DKK (kurs 1) eller samme valuta."""
    def niveau(d):
        v = indeks(d)
        a = 1.0 if kurs_fra is None else kurs_fra(d)
        b = 1.0 if kurs_til is None else kurs_til(d)
        if v is None or a is None or not b:
            return None
        return v * a / b
    return niveau


def forrentede(flows, slut, indeks, type_="Cash flow"):
    """Hver cash flow forrentet med indekset til slut: [{dato, beloeb, niveau, faktor,
    forrentet, type}], eller None, hvis indekset mangler på en af datoerne."""
    i_slut = indeks(slut)
    if not i_slut:
        return None
    linjer = []
    for d, a in flows:
        i_t = indeks(d)
        if not i_t:
            return None
        linjer.append(dict(dato=d, beloeb=a, niveau=i_t, faktor=i_slut / i_t,
                           forrentet=a * i_slut / i_t, type=type_))
    return linjer


def indekskonto(flows, datoer, indeks):
    """Indekskontoens værdi på hver dato: kald (< 0) købes, udlodninger (> 0) sælges, til
    indeksets niveau på dagen. Liste med én værdi pr. dato, None hvor indekset mangler."""
    ud = []
    for d in datoer:
        i_d = indeks(d)
        tidligere = [(t, a) for t, a in flows if t <= d]
        niveauer = [indeks(t) for t, _ in tidligere]
        if not i_d or not all(niveauer):
            ud.append(None)
            continue
        ud.append(-sum(a * i_d / i_t for (_, a), i_t in zip(tidligere, niveauer, strict=True)))
    return ud


def pme(flows, navs, slut, indeks):
    """PME for cash flows med LP-fortegn (kald < 0) og NAV'er [(dato, beløb)] (én pr. fond),
    alt forrentet til slut. {ks_pme, direct_alpha, merverdi, indeks_nav, indeks_irr,
    slutniveau, linjer}, eller None, hvis indekset mangler på en dato eller der intet er kaldt.

    merverdi er NAV minus indekskontoens værdi pr. slut (når alle NAV'er er pr. slut); indeks_irr
    er None, når udlodningerne har tømt kontoen (værdi <= 0), så der ingen IRR er at regne.
    """
    from .data import xirr      # her og ikke øverst: data importerer dette modul
    cf = forrentede(flows, slut, indeks)
    nv = forrentede(navs, slut, indeks, "NAV")
    if cf is None or nv is None:
        return None
    linjer = sorted(cf + nv, key=lambda x: (x["dato"], x["type"] == "NAV"))
    ud = sum(x["forrentet"] for x in linjer if x["forrentet"] > 0)
    ind = -sum(x["forrentet"] for x in linjer if x["forrentet"] < 0)
    if not ind:
        return None
    konto = -sum(x["forrentet"] for x in cf)
    return dict(ks_pme=ud / ind,
                direct_alpha=xirr([(x["dato"], x["forrentet"]) for x in linjer]),
                merverdi=sum(x["forrentet"] for x in linjer), indeks_nav=konto,
                indeks_irr=xirr(list(flows) + [(slut, konto)]) if konto > 0 else None,
                slutniveau=indeks(slut), linjer=linjer)
