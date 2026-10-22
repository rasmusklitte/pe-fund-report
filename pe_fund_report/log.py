"""Byggeloggen: én fil i log/ pr. bygning, med hvem, hvornår, hvilken kode og hvad der blev skrevet.

Alt, hvad bygningen skriver i terminalen (gemt, datotjek, ADVARSEL …), kommer også i loggen, så
advarslerne kan ses bagefter. Loggen er intern: den nævner investoren og de rigtige fonde, også
for den anonymiserede udgave, og ligger derfor uden for git ligesom exports/. Den skjulte faktor
i den anonymiserede udgave skrives aldrig.
"""
import getpass
import io
import platform
import subprocess
import sys
import time
import traceback
from contextlib import contextmanager
from datetime import datetime

from . import db
from .paths import LOG, ROOT


class _Tee(io.TextIOBase):
    """Skriver både til terminalen og til en buffer."""

    def __init__(self, ud):
        self.ud, self.buffer_ = ud, io.StringIO()

    def write(self, s):
        self.ud.write(s)
        return self.buffer_.write(s)

    def flush(self):
        self.ud.flush()


class Byggelog:
    """Det, bygningen fortæller loggen undervejs: de gemte filer og rapporten."""

    def __init__(self):
        self.filer, self.r = [], None


def _git():
    """'9869321' eller '9869321 med ikke-committede ændringer'; 'ukendt' uden git."""
    try:
        kald = lambda *a: subprocess.run(["git", *a], cwd=ROOT, capture_output=True, text=True,
                                         timeout=10, check=True).stdout.strip()
        commit = kald("rev-parse", "--short", "HEAD")
        return commit + (" med ikke-committede ændringer" if kald("status", "--porcelain") else "")
    except (OSError, subprocess.SubprocessError):
        return "ukendt"


def _dk(v, dec=0):
    return f"{v:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _portefoelje(r):
    """'10 fonde, NAV 317.699.673 DKK, IRR (DKK) 12,3 %' ud fra rapporten."""
    dele = [f"{len(r['fonde'])} fonde"]
    if (r.get("total_dkk") or {}).get("nav_dkk") is not None:
        dele.append(f"NAV {_dk(r['total_dkk']['nav_dkk'])} DKK")
    if (r.get("irr_dkk") or {}).get("irr") is not None:
        dele.append(f"IRR (DKK) {_dk(r['irr_dkk']['irr'] * 100, 1)} %")
    return ", ".join(dele)


def skriv_log(mappe, navn, oplysninger, tekst):
    """Skriv loggen som mappe/navn: oplysningerne som 'Felt: værdi' og derefter udskriften.
    Returnerer stien."""
    mappe.mkdir(parents=True, exist_ok=True)
    bredde = max(len(k) for k, _ in oplysninger) + 2
    linjer = [f"{k + ':':<{bredde}}{v}" for k, v in oplysninger]
    sti = mappe / navn
    sti.write_text("\n".join(linjer) + "\n\n--- Udskrift ---\n" + tekst, encoding="utf-8")
    return sti


@contextmanager
def byggelog(investor, as_of, anonym, mappe=LOG):
    """Kør bygningen inde i denne; loggen skrives bagefter, også når bygningen fejler (status
    FEJL med traceback, og fejlen går videre). Giver en Byggelog, som bygningen lægger de gemte
    filer og rapporten i."""
    start, tid = time.monotonic(), datetime.now()
    lg, tee, gammel = Byggelog(), _Tee(sys.stdout), sys.stdout
    sys.stdout, status = tee, "OK"
    try:
        yield lg
    except BaseException:
        status = "FEJL"
        tee.write(traceback.format_exc())
        raise
    finally:
        sys.stdout = gammel
        dato = lg.r["as_of"] if lg.r else as_of
        udgave = "anonymiseret" if anonym else "normal"
        navn = (f"{tid:%Y-%m-%d_%H%M%S}_{investor}_{dato or 'seneste'}"
                + ("_anonym" if anonym else "") + ".log")
        oplysninger = [
            ("Byggelog", f"{tid:%d.%m.%Y %H:%M:%S}"),
            ("Status", status),
            ("Bruger", f"{getpass.getuser()} på {platform.node()}"),
            ("Kommando", " ".join(["python -m pe_fund_report", *sys.argv[1:]])),
            ("Investor", investor),
            ("Rapportdato", f"{dato:%d.%m.%Y}" if hasattr(dato, "year") else dato or "seneste"),
            ("Udgave", udgave),
            ("Kode", f"git {_git()}"),
            ("Database", f"{db.SERVER} / {db.DATABASE}"),
            ("Filer", ", ".join(str(f) for f in lg.filer) or "ingen"),
            ("Varighed", f"{time.monotonic() - start:.0f} s")]
        if lg.r:
            oplysninger.append(("Porteføljen", _portefoelje(lg.r)))
        sti = skriv_log(mappe, navn, oplysninger, tee.buffer_.getvalue())
        print(f"log: {sti}")
