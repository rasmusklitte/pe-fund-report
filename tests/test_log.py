"""Byggeloggen (log.py): én fil pr. bygning, også når bygningen fejler."""
from datetime import date

import pytest

from pe_fund_report import log


def test_skriv_log(tmp_path):
    sti = log.skriv_log(tmp_path / "log", "x.log", [("Status", "OK"), ("Investor", "MOTORTRAMP")],
                        "gemt: a.pptx\nADVARSEL: noget\n")
    tekst = sti.read_text(encoding="utf-8")
    assert tekst.startswith("Status:   OK\nInvestor: MOTORTRAMP\n")
    assert tekst.endswith("--- Udskrift ---\ngemt: a.pptx\nADVARSEL: noget\n")


def test_byggelog_skriver_udskrift_og_oplysninger(tmp_path, capsys):
    with log.byggelog("MOTORTRAMP", None, False, mappe=tmp_path) as lg:
        print("ADVARSEL: KAPE22 er opgjort pr. 31.03.2026")
        lg.filer.append("exports/MOTORTRAMP_2026-06-30.pptx")
        lg.r = dict(as_of=date(2026, 6, 30), fonde=[{}, {}],
                    total_dkk=dict(nav_dkk=317_699_673.4), irr_dkk=dict(irr=0.1234))
    # udskriften står stadig i terminalen, og loggen er nævnt til sidst
    ud = capsys.readouterr().out
    assert "ADVARSEL: KAPE22" in ud and ud.rstrip().endswith(".log")
    (sti,) = tmp_path.iterdir()
    assert sti.name.endswith("_MOTORTRAMP_2026-06-30.log")
    tekst = sti.read_text(encoding="utf-8")
    for linje in ("Status:      OK", "Rapportdato: 30.06.2026", "Udgave:      normal",
                  "Filer:       exports/MOTORTRAMP_2026-06-30.pptx",
                  "Porteføljen: 2 fonde, NAV 317.699.673 DKK, IRR (DKK) 12,3 %",
                  "ADVARSEL: KAPE22 er opgjort pr. 31.03.2026"):
        assert linje in tekst, linje
    assert "Kode:        git " in tekst


def test_byggelog_ved_fejl(tmp_path, capsys):
    with pytest.raises(ValueError, match="nævner stadig"):
        with log.byggelog("MOTORTRAMP", "2026-06-30", True, mappe=tmp_path):
            print("noget før fejlen")
            raise ValueError("Den anonymiserede udgave nævner stadig: catacap iii")
    (sti,) = tmp_path.iterdir()
    assert sti.name.endswith("_MOTORTRAMP_2026-06-30_anonym.log")
    tekst = sti.read_text(encoding="utf-8")
    assert "Status:      FEJL" in tekst and "Udgave:      anonymiseret" in tekst
    assert "Filer:       ingen" in tekst and "Porteføljen" not in tekst
    assert "noget før fejlen" in tekst and "Traceback" in tekst
