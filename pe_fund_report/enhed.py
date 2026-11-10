"""Beløbsenheden, decket viser: mio. med én decimal (standard) eller tusinder uden decimaler
(--tusinder). Sættes for hele bygningen med beloebsenhed(); slides.py og data.py's notetekster
læser den, når de skriver et beløb eller dets enhed."""
from contextlib import contextmanager
from dataclasses import dataclass


@dataclass(frozen=True)
class Enhed:
    kort: str           # foran valutaen: »mio. DKK« / »t.DKK«
    divisor: int
    decimaler: int
    mellemrum: str      # mellem enheden og valutaen: »mio. DKK«, »t.DKK«
    diagram: str        # efter tallet på diagramakser og -etiketter: »120 mio.«; tusinder
                        # står uden (»120.000«), så etiketterne kan stå i søjlerne
    ord: str            # uden valuta efter: »beløb i mio., fondens valuta« / »i tusinder, ...«


MIO = Enhed("mio.", 1_000_000, 1, " ", " mio.", "mio.")
TUSINDE = Enhed("t.", 1_000, 0, "", "", "tusinder")

_AKTUEL = MIO


def aktuel():
    return _AKTUEL


@contextmanager
def beloebsenhed(e):
    """Vis alle beløb i enheden e, så længe blokken kører."""
    global _AKTUEL
    gammel, _AKTUEL = _AKTUEL, e
    try:
        yield
    finally:
        _AKTUEL = gammel


def kort():
    """Enheden uden valuta efter: »mio.« / »tusinder«."""
    return _AKTUEL.ord


def tekst(valuta):
    """Enheden foran valutaen: »mio. DKK« / »t.DKK«."""
    return f"{_AKTUEL.kort}{_AKTUEL.mellemrum}{valuta}"


def diagramsuffiks():
    """Det, der står efter et beløb på en diagramakse eller -etiket: » mio.« eller intet."""
    return _AKTUEL.diagram


def aksefmt():
    """Talformatet på en diagramakse med beløb."""
    return f'#,##0"{_AKTUEL.diagram}"' if _AKTUEL.diagram else "#,##0"


def seriefmt():
    """Talformatet på en diagramserie af beløb (slides._mio)."""
    return "#,##0" + ("." + "0" * _AKTUEL.decimaler if _AKTUEL.decimaler else "")
