"""Cash flow-prognose pr. fond efter Takahashi og Alexander (Yale, 2002) (rene funktioner).

En fond fremskrives kvartal for kvartal fra rapportdatoen ud fra sin profil
(config/cashflow_profiler.yaml):
- Kald: resttilsagnet ved starten af fondens leveår × årets kaldsrate, fordelt jævnt over
  året. Efter kald_til_aar (og løbetiden) kaldes intet; det, der står tilbage, kaldes aldrig.
- Vækst: NAV vokser med (1 + vaekst) om året.
- Udlodning: NAV efter vækst × max(udbytte, (alder / løbetid)^bow) om året; ved løbetidens
  udgang udloddes resten. En evergreen-fond har ingen løbetid og udlodder kun udbyttet.
- NAV = NAV efter vækst + kald - udlodning.

Beløbene er i fondens valuta, og kald er positive (det kaldte beløb). Ingen database.
"""
from datetime import date, timedelta

FELTER = ("kaldsrater", "kald_til_aar", "vaekst", "bow", "udbytte", "loebetid_aar")
STANDARD = "standard"     # profilen for et segment uden sin egen
AAR = 365.25


def profil(segment, profiler, afvigelser=None):
    """(profilens navn, parametre) for et segment, med fondens egne afvigelser lagt ovenpå.

    afvigelser er fund_info.yaml's prognose: {profil: <navn>, <felt>: <værdi>}. Et ukendt felt
    eller en ukendt profil er en fejl, så en stavefejl ikke stille bliver standardprofilen.
    """
    afvigelser = dict(afvigelser or {})
    navn = afvigelser.pop("profil", None)
    ukendte = set(afvigelser) - set(FELTER)
    if ukendte:
        raise ValueError(f"ukendte felter i prognose: {', '.join(sorted(ukendte))}")
    if navn is not None and navn not in profiler:
        raise ValueError(f"ukendt prognoseprofil: {navn}")
    navn = navn or (segment if segment in profiler else STANDARD)
    return navn, {**profiler[navn], **afvigelser}


def scenarie(p, aendring):
    """Profilen i et scenarie: vaekst + aendring, bow, udbytte og kaldsrater × aendring
    (kaldsraterne højst 1). Felter, scenariet ikke nævner, er uændrede."""
    a = aendring or {}
    return {**p,
            "vaekst": p["vaekst"] + a.get("vaekst", 0.0),
            "bow": p["bow"] * a.get("bow", 1.0),
            "udbytte": p["udbytte"] * a.get("udbytte", 1.0),
            "kaldsrater": [min(1.0, r * a.get("kaldsrater", 1.0)) for r in p["kaldsrater"]]}


def kvartalsslut(efter, til):
    """Kvartalsultimoerne efter datoen efter, til og med til."""
    ud = []
    aar, kv = efter.year, (efter.month - 1) // 3 + 1
    d = _kvartalsultimo(aar, kv)        # ultimo af kvartalet, datoen ligger i
    while d <= til:
        if d > efter:
            ud.append(d)
        aar, kv = (aar + 1, 1) if kv == 4 else (aar, kv + 1)
        d = _kvartalsultimo(aar, kv)
    return ud


def _kvartalsultimo(aar, kv):
    return date(aar, 12, 31) if kv == 4 else date(aar, 3 * kv + 1, 1) - timedelta(days=1)


def _alder(start, d):
    return max(0.0, (d - start).days / AAR)


def fremskriv(fra, til, resttilsagn, nav, start, loebetid, p):
    """Fondens kvartaler fra (ekskl.) til til (inkl.).

    start: fondens første kapitalkald (alderen regnes derfra); loebetid: år, None for en
    evergreen-fond. [{dato, aar, alder, nyt_aar, kaldsrate, rest_aarsstart, vaekstfaktor,
    udlodningsrate, rest_primo, nav_primo, kald, udlodning, nav, resttilsagn}]; satserne er
    dem, der er brugt i kvartalet (kaldsraten er årlig, udlodningsraten for kvartalet), og
    aar er kvartalets længde i år.
    """
    rest, nav = max(0.0, resttilsagn or 0.0), max(0.0, nav or 0.0)
    forrige, aarsstart, leveaar = fra, rest, None
    ud = []
    for d in kvartalsslut(fra, til):
        dt = (d - forrige).days / AAR
        alder, alder_slut = _alder(start, forrige), _alder(start, d)
        nyt_aar = int(alder) != leveaar
        if nyt_aar:
            leveaar, aarsstart = int(alder), rest
        rater = p["kaldsrater"]
        slut = min(x for x in (loebetid, p.get("kald_til_aar"), float("inf")) if x is not None)
        kaldsrate = 0.0 if alder >= slut else rater[min(leveaar, len(rater) - 1)]
        kald = min(rest, aarsstart * kaldsrate * dt)
        vaekst = (1 + p["vaekst"]) ** dt
        if loebetid is None:
            aarlig = p["udbytte"]
        elif alder_slut >= loebetid:
            aarlig = 1.0
        else:
            aarlig = max(p["udbytte"], (alder / loebetid) ** p["bow"])
        rate = 1.0 if aarlig >= 1 else 1 - (1 - aarlig) ** dt
        udlodning = nav * vaekst * rate
        ud.append(dict(dato=d, aar=dt, alder=alder, nyt_aar=nyt_aar, kaldsrate=kaldsrate,
                       rest_aarsstart=aarsstart, vaekstfaktor=vaekst, udlodningsrate=rate,
                       rest_primo=rest, nav_primo=nav, kald=kald, udlodning=udlodning,
                       nav=nav * vaekst + kald - udlodning, resttilsagn=rest - kald))
        rest, nav, forrige = rest - kald, nav * vaekst + kald - udlodning, d
    return ud


def pr_aar(kvartaler, kurs=1.0):
    """Kvartalerne samlet pr. kalenderår, ganget med kurs: [{aar, kald, udlodning, netto, nav,
    resttilsagn}], hvor nav og resttilsagn er ved årets sidste kvartal og netto er udlodning -
    kald."""
    aar = {}
    for q in kvartaler:
        a = aar.setdefault(q["dato"].year, dict(aar=q["dato"].year, kald=0.0, udlodning=0.0))
        a["kald"] += q["kald"] * kurs
        a["udlodning"] += q["udlodning"] * kurs
        a["nav"], a["resttilsagn"] = q["nav"] * kurs, q["resttilsagn"] * kurs
    for a in aar.values():
        a["netto"] = a["udlodning"] - a["kald"]
    return [aar[k] for k in sorted(aar)]


def summer(raekker):
    """Flere fondes pr_aar() lagt sammen år for år (samme år i alle)."""
    ud = {}
    for linjer in raekker:
        for a in linjer:
            s = ud.setdefault(a["aar"], dict(aar=a["aar"], kald=0.0, udlodning=0.0, netto=0.0,
                                             nav=0.0, resttilsagn=0.0))
            for k in ("kald", "udlodning", "netto", "nav", "resttilsagn"):
                s[k] += a[k]
    return [ud[k] for k in sorted(ud)]
