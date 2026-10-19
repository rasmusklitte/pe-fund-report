# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
pip install -e .[dev]                       # installs the pe-fund-report entry point too

python -m pe_fund_report MOTORTRAMP          # build the deck (latest NAV date)
python -m pe_fund_report MOTORTRAMP --as-of 2025-12-31 --out exports/test.pptx

pytest                                      # whole suite (no database needed)
pytest tests/test_slides.py::test_danske_talformater -q   # single test
ruff check .
```

The positional argument is `investor_code` in `fund.investor`. Output defaults to
`exports/<INVESTOR>_<date>.pptx`. PowerPoint locks the file while it is open, so `report._gem()`
falls back to `_v2`, `_v3`, … rather than failing — don't reach for `--out` to dodge a lock.

`ruff check .` currently reports 14 pre-existing findings (C408 `dict()` literals, DTZ011
`date.today()`) in `data.py` and the tests — that is the baseline, not a regression you caused.

## Architecture

A one-way pipeline, no intermediate files: **database → one nested dict → PowerPoint shapes**.

- `data.py` owns every database read and every calculation. `investor_rapport()` returns the whole
  report as one dict: `{investor, navn, as_of, valuta, total, gennemlysning, fonde: [...]}`, where
  each fund dict carries `n` (the key figures), `historik` (quarterly NAV + cumulative flows),
  `bro` (the value bridge), `beholdninger` and `noter`. Slide code never queries; it only formats
  this dict. `total()` and `gennemlysning()` are pure functions over the fund list — put portfolio
  aggregates there, not in a slide builder, so they stay testable without the database.
- `slides.py` builds each page as native PowerPoint tables, text boxes, charts and shapes (no
  images), so the deck stays editable. All geometry comes from the module constants `X0/X1` and
  `Y_UNDERTITEL/Y_TOP/Y_BUND/Y_NOTER` (inches on `Layout_basic_2`, a 11.693 × 6.575 in canvas —
  not the 13.333 × 7.5 of a standard 16:9 deck) — change those, not per-shape numbers, to move
  content. Real chart objects (`_historik_graf`, the look-through bar chart) are used where the
  data should stay editable in PowerPoint; the in-row bars on *Lange tilsagn* and the company
  cards are drawn rectangles (`_bjaelke`, `_soejler`, `_vandfald`, `_noegletalsbjaelke`), because
  one chart object per row is unwieldy.
- **Slide notes go to the client.** Keep them to sources and method; no file paths, view names or
  internal workflow states. `data.py` phrases caveats the same way — "forvalteren oplyser ikke …",
  not "ikke indlæst i databasen".
- `report.py` assembles the slide order and saves; `cli.py` is only argument parsing.
- The *Executive summary* (slide 2) formats `data.sammendrag()`: portfolio figures in DKK
  (from `total_dkk`), highest/lowest TVPI fund and the largest geography/segment from
  `r["fordelinger"]`. There is no portfolio IRR across currencies; IRR is only shown per fund.
- The *Definitioner* slide before the end slide prints `slides.DEFINITIONER`, plain text that
  describes how `data.py` calculates each figure. Change a calculation, update its definition.
- `deck.py` + `style.py` are the shared PPIM house-style modules, **kept byte-identical to the
  other PE repos** (`pe-analysis`, `pe-quarterly-report`, …). They are excluded from ruff in
  `pyproject.toml`. Do not refactor them here; port a change to the other repos or not at all.
  `deck.py` also holds the slide-master layout indices (`L_FORSIDE`, `L_TEKST`, `L_SKILLE`, …).

`data.datotjek()` guards the one assumption the whole deck rests on: that every fund is stated
at the same date. Funds report on their own schedule and the NAV series have gaps, so
`investor_rapport()` takes the **latest** NAV date across funds as the report date, and any fund
that lags silently contributes older figures to every total and distribution. The check returns
`foraeldede` (funds behind the report date) and `skaeve_beholdninger` (holdings dated differently
from that fund's own NAV); `byg()` prints it at build time, and the cross-fund slides mark the
row with a dagger and the fund's real date. Keep new cross-fund slides wired into
`slides._datonoter(r)` / `_datomarkering(r, navn)` — a slide that aggregates across funds without
them will mix dates with no visible trace.

`noter` is how data-quality caveats reach the deck: `data.py` appends a Danish sentence, and
`slides.py` prints them in the small italic source line at `Y_NOTER`. Add a caveat by appending to
`noter`, not by writing text in the slide builder.

### Code conventions

Identifiers, docstrings and comments are **Danish** (`fond`, `tilsagn`, `indbetalt`, `udloddet`,
`beholdninger`, `_saet_ph`); only the shared `style.py`/`deck.py` API and database columns are
English. Report text is Danish, numbers use Danish formatting via `style._dk` (wrapped by
`slides.tusind/pct/gange`), and missing values render as `"n/a"` (`slides.IKKE_OPLYST`), never as
a blank or a zero.

## Data rules

`docs/analysis-data-guide.md` is the contract for the `PE` database and takes precedence over
intuition; `docs/existing-database.md` describes the legacy `dbo` objects, which this project does
not use. Non-negotiables:

- Read through the `rpt` views. Base tables need `document.status = 'APPROVED'` and `is_active = 1`.
- **Never add fund-level amounts to investor-level amounts.** `rpt.v_nav` is our share;
  `rpt.v_portfolio_holding` is the whole fund and is scaled by the ownership share in
  `_beholdninger()`. When fund size is unknown the holdings are shown unscaled and labelled so.
- Cash flows carry the LP sign (paid < 0). `EQUALISATION` stays out of paid-in and distributions,
  `WITHHOLDING_TAX` out of IRR.
- IRR is `data.xirr` (act/365 bisection) over the investor's cash flows with NAV as the final flow;
  the portfolio IRR is pooled over all funds, not averaged. `data.total()` returns `None` when the
  funds are not all in the same currency — the slides then drop the total row.
- If the GP's own paid-in figure differs from the registered capital calls by more than
  `GP_TOLERANCE` (0.5 %), the GP figure wins and a note is added.
- DKK conversion (*Tilsagn på tværs af fonde*: resttilsagn and NAV in DKK) uses one
  Nationalbanken rate per currency at the report date, read by `data._dkk_kurs()` from
  `dbo.fx_rates_nationalbanken` (latest `rate_date` on or before; DKK has no rows and gets 1).
  `data.dkk_total()` sums DKK amounts across currencies; it is `None` if a fund lacks a rate.
  The geography/segment slides use the same conversion via `data.fordeling()`: the share is
  commitment in DKK over total commitment in DKK; TVPI on the total/subtotal rows is
  (distributions + NAV) / paid-in, all in DKK at the report-date rate (equal to the
  fund-currency TVPI for a single-currency group); fund-currency commitment is shown on those
  rows only when all their funds share one currency.
- `db.read_sql` rejects anything but `SELECT`/`WITH`. Keep it that way: this repo is read-only
  against the database. Connection defaults to `sql1\ppim` / `PE` with Windows auth, overridable
  via `PE_DB_SERVER`, `PE_DB_DATABASE`, `PE_DB_DRIVER`.

## Manual master data

Fund master data comes from `rpt.v_fund_master`, joined in `data._investeringer()`:
- geography and segment (`geography_focus`, `fund_structure`, `strategy_focus`, `size_focus`),
  translated to Danish by `data.geografi()` / `data.segment()`. The code lists are closed by
  check constraints on `fund.fund`; a new code needs a line in `data.GEOGRAFIER` /
  `STRUKTURER` / `FOKUS` / `STOERRELSER`.
- løbetid from `term_years` + `extension_years` via `data.loebetid()`, which writes the
  extension years as one-year steps ("10+1+1"); the database holds only the total.
- fund size: the dated `TOTAL_COMMITMENTS` metric first, else `total_commitments`.
- manager and logo: for our feeder companies (`underlying_fund_code` set) the underlying
  fund's manager, not our own.

`config/fund_info.yaml` holds only what the database does not: `beskrivelse` and
`estimerede_cashflows` for notices that never arrived. Optional: `benchmark` (shown on
*Lange tilsagn*) and `virksomheder` (per-company description and ownership for the profile
cards; the key must match `name` in `rpt.v_portfolio_holding` exactly). `status: udkast` marks
text written by Claude and not yet reviewed; it puts a footnote on the slide until someone
changes it to `ok`. `tests/test_data.py` asserts every entry has `status` and `beskrivelse`,
and that none carries a field the master view now owns.

## Layout reference

`docs/reference/2026 Q2 ATP PEP IX B - Quarterly Internal Report.pdf` is the model for the layout
and the metric set — consult it before changing what a fund page shows.
`assets/ppim_slidemaster.pptx` is the slide master; `tests/test_slides.py` builds against the real
file, so tests break if its layouts move.
