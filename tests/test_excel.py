"""Dataudtrækket i excel.py, bygget på en lille rapport uden database."""
from datetime import date

import pandas as pd

from pe_fund_report import data, excel

NAV_DATO = date(2026, 6, 30)


def _fond(dok_mappe):
    nav = pd.DataFrame(dict(period_end=[date(2025, 12, 31), date(2026, 3, 31), NAV_DATO],
                            nav=[100.0, 105.0, 130.0], nav_rapporteret=[100.0, 105.0, 130.0],
                            restatement=[None, None, None], document_id=[1, 2, 3]))
    cf = pd.DataFrame(dict(
        value_date=[date(2025, 6, 1), date(2026, 5, 16), date(2026, 5, 16), date(2026, 8, 1)],
        notice_date=[None] * 4, account_code=["CF100", "CF100", "CF200", "CF100"],
        account_name=["Kald"] * 2 + ["Udlodning", "Kald"],
        flow_class=["CONTRIBUTION", "CONTRIBUTION", "DISTRIBUTION", "CONTRIBUTION"],
        is_recallable=[0, 0, 0, 0], amount=[-80.0, -20.0, 5.0, -7.0], fx_rate=[None] * 4,
        amount_dkk=[None] * 4, flow_ref=["a", "b", "c", "d"], document_id=[10, 11, 12, 13]))
    ka = pd.DataFrame(dict(period_end=[NAV_DATO], basis=["ITD"], account_code=["CA110"],
                           account_name=["Contributions"], amount=[100.0], document_id=[3]))
    tilsagn = pd.DataFrame(dict(event_date=[date(2025, 1, 1)], event_type=["INITIAL"],
                                amount=[200.0], document_id=[None], note=["Aftale"]))
    historik = [dict(dato=d, nav=v, indbetalt=i, udloddet=u) for d, v, i, u in
                [(date(2025, 12, 31), 100.0, 80.0, 0.0), (date(2026, 3, 31), 105.0, 80.0, 0.0),
                 (NAV_DATO, 130.0, 100.0, 5.0)]]
    f = dict(fund_code="ALFA", fund_name="Alfa Fund I K/S", valuta="DKK", nav_dato=NAV_DATO,
             manager="Alfa", geografi="Danmark", segment="Buyout", vintage=2021, loebetid="10",
             benchmark=None, noter=[], historik=historik,
             kapitalstroemme=[(d, a, "kald" if a < 0 else "udlodning")
                              for d, a in zip(cf.value_date[:3], cf.amount[:3])],
             n=dict(tilsagn=200.0, indbetalt=100.0, udloddet=5.0, recallable=0.0, nav=130.0,
                    irr=0.12, fondsstoerrelse=2000.0),
             irr_flows=[(date(2025, 6, 1), -80.0), (date(2026, 5, 16), -20.0),
                        (date(2026, 5, 16), 5.0), (NAV_DATO, 130.0)],
             bro=[("Indbetalt kapital", 100.0), ("Udlodninger", -5.0),
                  ("Værdiskabelse før carry", 35.0), ("NAV", 130.0)],
             beholdninger=dict(dato=NAV_DATO, skaleret=True, andel=0.1, type="selskaber",
                               raekker=[dict(navn="Selskab", type="COMPANY", sektor="IT",
                                             vintage=2022, closing=None, kostpris_fond=50.0,
                                             dagsvaerdi_fond=80.0, udloddet_fond=0.0,
                                             document_id=3)]),
             beholdninger_primo={"kvartal": None, "aar": None},
             kilder=dict(nav=nav, cash_flows=cf, kapitalkonto=ka, tilsagn=tilsagn,
                         fondsstoerrelse=dict(tekst="Stamdata", document_id=None),
                         indbetalt="kapitalkald", estimerede=[]))
    data.udvikling(f, data.perioder(NAV_DATO))
    data.omregn_udvikling_til_dkk([f], lambda v, d: 1.0)
    (dok_mappe / "ALFA").mkdir()
    (dok_mappe / "ALFA" / "q2.pdf").write_bytes(b"%PDF")
    return f


def _rapport(tmp_path, monkeypatch):
    monkeypatch.setattr(excel, "DOKUMENTER", tmp_path)
    f = _fond(tmp_path)
    dok = pd.DataFrame(dict(
        document_id=[1, 2, 3, 10, 11, 12, 13], fund_code=["ALFA"] * 7,
        doc_type=["CAPITAL_STATEMENT"] * 3 + ["CAPITAL_CALL"] * 2 + ["DISTRIBUTION", "CAPITAL_CALL"],
        period_end=[None] * 7, value_date=[None] * 7, notice_date=[None] * 7,
        external_ref=[None] * 7,
        file_name=["q4.pdf", "q1.pdf", "q2.pdf", "k1.pdf", "k2 - PLACEHOLDER.pdf", "u1.pdf", "k3.pdf"],
        rel_path=["ALFA/q4.pdf", "ALFA/q1.pdf", "ALFA/q2.pdf", "ALFA/k1.pdf", "ALFA/k2.pdf",
                  "ALFA/u1.pdf", "ALFA/k3.pdf"],
        page_count=[1] * 7, status=["APPROVED"] * 7, extraction_method=["MANUAL"] * 7,
        approved_by=["x"] * 7, approved_at=[None] * 7, note=[None] * 7))
    r = dict(navn="Investor A/S", as_of=NAV_DATO, fonde=[f], perioder=data.perioder(NAV_DATO),
             kurser={"DKK": (1.0, NAV_DATO)}, kursopslag={("DKK", NAV_DATO): (1.0, NAV_DATO)},
             udvikling=data.portefoelje_udvikling([f]), vaerdibevaegelser=None, dokumenter=dok,
             datotjek=dict(ens=True, foraeldede=[], skaeve_beholdninger=[]))
    f["n"]["nav_dkk"] = f["n"]["nav"]
    r["irr_dkk"] = data.irr_dkk([f], lambda v, d: 1.0)
    f["n"].update(tilsagn_dkk=200.0, resttilsagn_dkk=100.0, indbetalt_dkk=100.0, udloddet_dkk=5.0)
    r["likviditet"] = data.likviditet([f], lambda v, d: 1.0, NAV_DATO)
    return excel.skriv(r)


def _kolonne(ws, navn):
    return next(c.column for c in ws[1] if c.value == navn)


def test_projektmappen_har_faner_i_fast_raekkefoelge(tmp_path, monkeypatch):
    wb = _rapport(tmp_path, monkeypatch)
    assert wb.sheetnames[:4] == ["Læs mig", "Nøgletal", "Udvikling", "Værdibro"]
    assert wb.sheetnames[-1] == "Dokumenter"


def test_noegletal_er_formler_med_kontrol_mod_kapitalstroemmene(tmp_path, monkeypatch):
    ws = _rapport(tmp_path, monkeypatch)["Nøgletal"]
    tvpi = ws.cell(2, _kolonne(ws, "TVPI")).value
    assert tvpi.startswith("=IFERROR((") and "/F2" in tvpi
    kontrol = ws.cell(2, _kolonne(ws, "Kontrol: indbetalt iflg. kapitalkald")).value
    assert "SUMIFS('Kapitalstrømme'!$I:$I" in kontrol and '"CONTRIBUTION"' in kontrol
    assert ws.cell(2, _kolonne(ws, "Kontrol: IRR med XIRR")).value.startswith("=IFERROR(XIRR(IRR!")


def test_irr_i_dkk_har_omregning_som_formel_og_xirr_kontrol(tmp_path, monkeypatch):
    ws = _rapport(tmp_path, monkeypatch)["IRR DKK"]
    assert ws.cell(2, _kolonne(ws, "Beløb DKK")).value == "=E2*F2"
    assert ws.cell(2, _kolonne(ws, "Kontrol: IRR med XIRR")).value.startswith("=IFERROR(XIRR(")
    assert ws.cell(ws.max_row, _kolonne(ws, "Type")).value == "NAV (sidste cash flow)"


def test_kapitalstroemme_efter_nav_datoen_er_ikke_i_noegletal(tmp_path, monkeypatch):
    ws = _rapport(tmp_path, monkeypatch)["Kapitalstrømme"]
    i = _kolonne(ws, "I nøgletal")
    vaerdier = [ws.cell(r, i).value for r in range(2, ws.max_row + 1)]
    assert vaerdier == ["Ja", "Ja", "Ja", "Nej"]


def test_kilder_linker_til_pdf_og_til_dokumentfanen(tmp_path, monkeypatch):
    wb = _rapport(tmp_path, monkeypatch)
    nav = wb["NAV"]
    ultimo = nav.cell(4, _kolonne(nav, "Kildedokument"))
    assert ultimo.value == "q2.pdf" and ultimo.hyperlink.target.startswith("file:")
    assert nav.cell(2, _kolonne(nav, "Kildedokument")).hyperlink is None   # filen findes ikke
    assert nav.cell(4, _kolonne(nav, "Dokument-id")).hyperlink.location.startswith("Dokumenter!")  # hop i projektmappen
    dok = wb["Dokumenter"]
    plads = {dok.cell(r, 1).value: dok.cell(r, _kolonne(dok, "Pladsholder")).value
             for r in range(2, dok.max_row + 1)}
    assert plads[11] == "Ja" and plads[10] == ""


def test_udvikling_regner_vaerdiaendringen_som_formel(tmp_path, monkeypatch):
    ws = _rapport(tmp_path, monkeypatch)["Udvikling"]
    assert ws.cell(2, _kolonne(ws, "Periode")).value == "Kvartalet"
    assert ws.cell(2, _kolonne(ws, "Værdiændring")).value == "=J2-G2-H2+I2"
    assert ws.cell(3, 1).value == "I alt"


def test_likviditet_har_kvartaler_og_aarsskifter_med_formler(tmp_path, monkeypatch):
    ws = _rapport(tmp_path, monkeypatch)["Likviditet"]
    assert ws.cell(2, 1).value == "Kvartal" and ws.cell(2, _kolonne(ws, "Netto DKK")).value == "=D2-C2"
    assert ws.cell(3, _kolonne(ws, "Akkumuleret netto DKK")).value == "=F2+E3"
    aar = next(rad for rad in range(2, ws.max_row + 1) if ws.cell(rad, 1).value == "Årsskifte")
    assert ws.cell(aar, _kolonne(ws, "TVPI")).value == f'=IFERROR((K{aar}+L{aar})/J{aar},"")'
