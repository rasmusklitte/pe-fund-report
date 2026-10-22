# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
pip install -e .[dev]                       # installs the pe-fund-report entry point too

python -m pe_fund_report MOTORTRAMP          # build the deck (latest NAV date)
python -m pe_fund_report MOTORTRAMP --as-of 2025-12-31 --out exports/test.pptx

pytest                                      # whole suite; only test_aleta.py reads the database
pytest -m "not db"                          # without the database tests
python -m pytest tests/test_aleta.py        # the report vs Aleta, also writes exports/aleta_afstemning_<date>.xlsx
pytest tests/test_slides.py::test_danske_talformater -q   # single test
ruff check .
```

The positional argument is `investor_code` in `fund.investor`. Output defaults to
`exports/<INVESTOR>_<date>.pptx`. PowerPoint locks the file while it is open, so `report._gem()`
falls back to `_v2`, `_v3`, … rather than failing — don't reach for `--out` to dodge a lock.

`tests/test_aleta.py` compares the Motortramp report at 30.06.2026 with Aleta's export
(`tests/motortramp_aleta_data.xlsx`), one test per fund × check (`tests/aleta.py`: `FONDE` maps
Aleta's `securityCode` to our fund code, `KONTROLLER` holds the tolerances). It is marked `db` and
skips when the database can't be reached. Failures there are differences to look into, not
necessarily bugs. Aleta's amounts are in DKK (except
`nav`), so ours are compared in DKK: NAV and commitments at the report-date rate (`n["..._dkk"]`),
paid-in and distributions at each cash flow's own date (`aleta.transaktionskurs`), as Aleta does.

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
  internal workflow states. **Nothing is said twice** (the user's rule): a note never repeats what the
  slide already shows — the subtitle (DKK, investors NAV, »hele fonden«), a heading or column name
  (»Andel af NAV«, »NAV + resttilsagn«), a figure in a table or block (NAV, ejerandel, fondens NAV,
  indbetalt), the fund in the row a mark sits in, or the index name by the bar — nor what another
  note on the same slide says, nor anything `slides.DEFINITIONER` already explains (PICC, TVPI,
  IRR, afkast, værdiændring, ejerandel, gennemlysning, genkaldelige udlodninger, udligningsrente …;
  a note gives the fund's own amount, not the definition again) (the source line names Danmarks Nationalbank, so `_kursnote` gives
  only the rates). Dates equal to the NAV date are left out, but a date that differs (an older
  primo NAV, holdings or a breakdown at another date, a fund behind the report date) is always
  noted, also on slides that do not show the figure it affects (the user's rule). `data.py` phrases caveats the same way — "forvalteren oplyser ikke …",
  not "ikke indlæst i databasen".
- `report.py` assembles the slide order and saves; `cli.py` is only argument parsing. Order:
  cover, executive summary, udvikling | *Porteføljeoversigt* section: *Investeringer*
  (`tilsagn_paa_tvaers`), ændringer pr. fond, *Eksponering* (`valutafordeling`), segment,
  *Branchefordeling – hele porteføljen* (`branchekort`), branche pr. aktivklasse,
  *Geografisk eksponering – hele porteføljen* (`regionskort`), geografi pr. aktivklasse,
  største positioner, gennemlysning |
  *Fondsoversigter* section: the portfolio overview table, then each fund page and its company
  cards | *Appendiks* section: definitions, the GICS overview (`slides.gics_oversigt`: the 11
  sectors with a Danish description in `GICS_BESKRIVELSER` and their industry groups from
  `data.GICS_STRUKTUR`, in the sector colours) | end slide. The geography-by-commitment slide is no longer built (the data
  stays: the executive summary falls back on it). Nor are the pie slides *Branchefordeling*
  and *Geografisk eksponering* (`slides.brancher` / `regioner`) or the two "ejendomme for sig"
  treemaps (`branchekort_ejendomme` / `regionskort_ejendomme`, replaced by the whole-portfolio
  treemaps at the user's request): the code and tests are kept for later use. *Største værdibevægelser*
  (`slides.vaerdibevaegelser`) is likewise kept but not built; its data stays in Excel. Client
  notes must therefore not refer to those pages by name.
- *Eksponering* (`slides.valutafordeling`, after *Ændringer pr. fond*): three doughnuts
  (`slides._ring`) in a near-equilateral triangle in the middle of the slide (exposure on top,
  NAV and resttilsagn below, `VF_BASIS` apart; a true equilateral would make the lower rings
  touch), with `_forbindelser` drawing one line from each lower ring up to the exposure (no
  line or »+« between NAV and resttilsagn); no table. Currency and % sit on the ring, the measure's name and total in the hole.
  `_fyld_diagramflade` replaces PowerPoint's auto layout so the ring fills its chart box. Data
  from `data.valutafordeling()` (pure) in `r["valutafordeling"]`: NAV, resttilsagn and exposure (= NAV + resttilsagn) in DKK
  grouped by the fund's currency; `None` if a fund lacks a DKK figure. Colours are fixed per
  currency in `VALUTA_FARVER`.
- Page footer: `slides.sidefod()` runs last in `report.byg()` and clones the master's slide-number
  placeholder onto every `L_TEKST` slide as "**<investor>** | Side n" (n is a `slidenum` field).
  Cover, dividers and end slide get none but still count. *Lange tilsagn* (`slides.lange_tilsagn`,
  the only user of `benchmark`) is kept and tested but no longer in the deck - the appendix was dropped.
- **Feeder look-through.** PPPC2 and PPPE2 are our feeders into PSCP5 and NEP6
  (`underlying_fund_code`). When a feeder has no holdings of its own,
  `data._feeder_beholdninger()` takes the underlying fund's holdings and **scales by NAV**:
  `andel = feeder NAV / underlying total fair value`, so the positions add up to the
  investor's NAV (the user chose this). Scaling through the commitment chain overstates the
  position 1.7-2.4x. `beh["via"]` makes `fondsside` show `_gennemlysningsblok` (top 10
  positions + the GP's breakdown from `rpt.v_portfolio_breakdown`) instead of the NAV chart and
  bridge, and makes `virksomhedsprofiler` skip the fund. `_beholdninger()` adds up lines with
  the same name, and in tranche tables (`instrument` set) it groups by company via `_selskab()`,
  which strips instrument words. Matching between quarters ignores case.
- **Property funds.** Home.Earth reports properties (`section = 'PROPERTY'`), not companies.
  `data._ejendomme()` keeps them as whole-fund values before debt (they sum to GAV, while the
  investor's share is NAV after debt), so they are not scaled by ownership, and translates the
  valuation basis via `VAERDIGRUNDLAG`. `fondsside` shows `_ejendomsblok` instead of the NAV
  chart and value bridge: a full-width waterfall (`_nav_bro`) from each property up to GAV,
  then (with `TOTAL_ASSETS`) other assets up and debt down to NAV, scaled by investor NAV /
  fund NAV (`FUND_NAV` metric) so it ends at the investor's NAV. A real-estate fund without
  property lines (CRED) gets the same page from its GAV/TOTAL_ASSETS/FUND_NAV metrics via
  `data._ejendomsbalance()`. They stay out of `vaerdibevaegelser` and `gennemlysning`, and
  get no company cards.
- **Evergreen funds** (`evergreen: true` in `fund_info.yaml`, since the database has no flag):
  løbetid shows "Evergreen", and the ownership share is investor NAV / fund NAV rather than
  commitment / fund commitments, because new units are issued at the current price.
- Fund names in the deck drop the company form (`data.uden_selskabsform`, applied to
  `fund_name` in `data.fond()`); the legal name is `juridisk_navn`, shown in Excel's Stamdata.
- `excel.py` writes `<deck>_data.xlsx` next to the deck (skip it with `--uden-data`). This is
  the **internal** reconciliation workbook, not for the client: every database row carries its
  `document_id` (an in-workbook link to the *Dokumenter* sheet) and a file link to the PDF
  under `paths.DOKUMENTER` (`../pe-external-fund-data/documents`, override with `PE_DOCUMENTS`;
  `fund.document.rel_path` is relative to it). Derived figures are Excel formulas, and the
  *Kontrol* columns recompute them from the other sheets (SUMIFS over *Kapitalstrømme*, XIRR
  over *IRR*). The source rows come from `f["kilder"]` and `r["dokumenter"]` /
  `r["kursopslag"]` in `data.py`. When you add a read to `data.fond()`, keep its
  `document_id` and put the rows in `kilder`, so the figure can be traced.
- The *Executive summary* (slide 2) formats `data.sammendrag()`: portfolio figures in DKK
  (from `total_dkk`), highest/lowest TVPI fund and the largest geography/segment. Geography is the
  same regions as *Geografisk eksponering* (`r["regioner"]`, NAV look-through; falls back to the
  funds' geography focus by commitment when that is `None`); segment is the funds' strategy by
  NAV in DKK (`andel_nav` from `r["fordelinger"]`, the user's choice), by commitment when a fund
  lacks a DKK figure. The segment slide itself stays by commitment. `slides._sammendragsfordelinger()` picks them, and the block headings say
  which measure each uses. One bullet names the largest single exposure from
  `data.stoerste_eksponering()` (pure): the largest look-through line of
  `_positionskandidater()` — company or underlying fund, feeders looked through, the same name in
  several funds added up — as the investor's share in DKK; property and unscaled holdings are
  left out and named in the note. It is the same ranking as *Største positioner*.
- *Største positioner* (`slides.stoerste_positioner`, after geografi pr. aktivklasse): the top
  `POSITIONER_ANTAL` (15) positions from `data.stoerste_positioner()` (pure), as a table with a
  drawn bar and an "i alt" row. A position is the investor's share of the holding's fair value
  (ownership × fair value, as on the fund pages) in DKK at the report-date rate, shown as % of the
  investor's total NAV (the user's choice; holdings don't sum to NAV because of the funds' cash
  and carry). Funds without comparable holdings — property funds, unscaled or no holdings — are
  ranked as one line at their NAV ("Fond (ikke gennemlyst)", light bar), also the user's choice.
- **Portfolio IRR in DKK** (`data.irr_dkk()`, pure, in `r["irr_dkk"]`): every fund's IRR cash
  flows converted at the Nationalbanken rate on their own date (same as `amount_dkk` in
  `rpt.v_cash_flow`) with `nav_dkk` as the final flow, pooled with `xirr`. It is the IRR on the
  executive summary tile and on the *Investeringer* total row when currencies differ; a
  single-currency total row keeps the fund-currency IRR from `total()`. The Excel sheet
  *IRR DKK* lists the flows with the DKK amount as a formula and an XIRR check.
- **Industry and geography slides** (`slides.brancher` / `slides.regioner`, after the segment
  slide): one shared data path, configured per `art` in `data.ARTER`. `data._eksponering()` reads
  each fund's shares from `rpt.v_industry_exposure` / `rpt.v_geography_exposure` (GP breakdown, else
  holdings at fair value, else `FUND_DEFAULT` from the fund's strategy; feeders look through) and
  only uses the view's amounts as weights. `eksponering(fonde, art)` (pure) multiplies by the fund's
  own `nav_dkk`, so the groups sum to the portfolio NAV in DKK; keys are neutral (`grupper` →
  `under`: sector → industry group, region → subregion). Names come from `SEKTORER` /
  `BRANCHEGRUPPER` (the official **English** GICS names, as the user asked; DIVERSIFIED/OTHER are
  not GICS and stay Danish, and the asset class »Ejendomme« stays Danish) and `REGIONER` /
  `DELREGIONER`; `*_OTHER` subregions are "ikke opdelt" and are
  explained in the slide note. Classifications with `is_suggested = 1` go into `foreloebig_andel`:
  the slide says "foreløbig klassifikation" and `byg()` warns. The geography pie has one slice per
  region, in the first colour of its family (`GEO_FARVER`). Excel sheets *Brancher* and *Geografi*
  recompute each line as share × NAV. The method is in the slides' source notes, not in
  `DEFINITIONER`, which is full. The geography slides (pie, treemaps, feeder pages) show only the
  three regions: `slides._kun_regioner()` empties `under`, so the subregions live on only in the
  data and the Excel *Geografi* sheet. Don't alias SQL columns `top` (reserved word in T-SQL). The
  region codes are the database's (`NORTH_AMERICA`, `EUROPE`, `ROW`); a new code in
  `portfolio.geo_region`/`geo_subregion` needs a line in `REGIONER`/`DELREGIONER`.
- **Industry slides.** *Branchefordeling* shows GICS sectors only (pie + table); the industry
  groups stay in the data and in the Excel *Brancher* sheet. *Branchefordeling – hele porteføljen*
  (`slides.branchekort`) draws the same sectors as a squarified treemap (`slides._treemap`, pure,
  areas proportional to the values) of native rectangles, with »Øvrige« for
  sectors under `BR_MIN`; each box shows name, tDKK and %, shrinking to name + % or % only when
  small. *Geografisk eksponering – hele porteføljen* (`slides.regionskort`) is the same idea with one
  box per region in the pie's colours. Every box must say what it is: `slides._gulv` lifts
  the area of boxes under `KORT_MINDST` (3 %) of their treemap so they can carry a label (the
  figures stay the actual ones; `_gulvnote` says so), and `_kortkasse` falls back to smaller type
  and then to a label turned 90° in tall, narrow boxes (name and % on two lines first). A box
  that only has room for its percentage gets a footnote mark (`FODNOTETEGN`, "¹ 2,5 %") and is
  named in the note by `_uden_navn_note()`; the slide builder passes one `uden_navn` list through
  `_branchekasser` / `_regionskasser` / `_delt_flade`, so the marks are unique per slide.
- **One colour per category.** Each GICS sector has a fixed colour in `slides.SEKTOR_FARVER`,
  the same on every slide and in every part (the user's rule: a colour always means the same
  category). Grey is reserved: `BR_OEVRIGE` only for »Øvrige« / »Øvrige fonde«, `BR_UDEN_GICS`
  for non-GICS rows. Real Estate is a muted teal (#4F7470; the user found plum too loud), so
  it is not confused with the blue-grey property-segment slates on *pr. aktivklasse* (the house colours are taken by sectors and
  regions), and the property segments are fixed slate shades in `EJD_SEGMENT_FARVER`.
  Regions keep `GEO_FARVER`.
- **Treemaps with property for itself** (`slides.branchekort_ejendomme` / `regionskort_ejendomme`,
  kept but not in the deck). The area is split into »Alternative investeringer« (left,
  drawn by the same `_branchekasser` / `_regionskasser` helpers as the overviews) and Ejendomme
  (right), widths proportional to NAV. The Ejendomme part is not shown per fund (the user's
  choice): `ejendomme["fordeling"]` splits it on the industry slide by the funds' property
  segment (`re_sector` in `rpt.v_fund_master` → `data.EJENDOMSSEGMENTER`, English like GICS;
  the property share of a fund without a segment is »Øvrige«, in the »Øvrige« grey), via
  `data.ejendomssegmenter()`, and on the geography slide by region, as an ordinary
  `eksponering()` over each fund's property NAV, drawn by `_regionskasser` in the region
  colours. Both via `slides._ejendomskasser`. The bands show each part's share of total NAV; the
  box percentages are of the part's own NAV, so each part sums to 100 %. Ejendomme is the GICS
  sector `REAL_ESTATE` across all funds (property funds wholly, plus the real-estate share of
  mixed funds), from `data.opdel_ejendomme()` (pure) in `r["uden_ejendomme"][art]`; `resten` is
  an ordinary `eksponering()` result over the funds' non-property NAV. The data has no industry ×
  geography cross, so on the geography slide a mixed fund's property share is assumed to sit like
  the rest of the fund (said in the note). The industry slide shows ACWI ex real estate,
  rescaled to 100 %.
- **Treemaps per asset class** (`slides.branchekort_aktivklasser` / `regionskort_aktivklasser`,
  each right after its whole-portfolio treemap): the same idea in three parts, Private Equity |
  Private Credit | Ejendomme, from `data.opdel_aktivklasser()` (pure) in
  `r["aktivklasser"][art]`. Ejendomme is the same GICS-sector cut as `opdel_ejendomme()` (both
  use `data._uden_ejendomme()`); Private Credit is the rest of the funds whose segment is private
  credit (`JUNIOR_DEBT`), Private Equity the rest of the others. `slides._delt_flade` draws any
  number of parts left to right with Ejendomme last; `_delbredder` keeps each part at least
  `EJD_MIN_W` wide. `_ejendomsopdelt` is its one-part wrapper.
- **MSCI benchmark.** `config/benchmark.yaml` holds the MSCI ACWI sector weights (World is
  commented out: the user only wants ACWI for now; every index listed gets its own bar), keyed
  by `industry_code`, and under `regioner:` their region weights keyed by `NORTH_AMERICA`/`EUROPE`/`ROW`
  (summed from ETF country weights; unlisted countries go to ROW; the slides do not explain
  this, the user's choice).
  `data.benchmark()` loads it into `r["benchmark"]` (`indeks`, `regioner`). Every industry
  treemap shows the same sector bar (`slides._bm_sektorer(r)`: the index as it is, grouped by
  the whole-portfolio map), so the index figures are identical across slides (the user's rule). The treemaps do not
  print the index in the boxes: `slides._benchmarkbaand` draws one 100 % stacked bar per index
  under the map, in the categories' colours (sectors without their own box go into »Øvrige«;
  on *pr. aktivklasse* the sector index is without Real Estate, rescaled), and the map shrinks by
  `_bm_hoejde()`. *Branchefordeling* / *Geografisk eksponering* (pies, not in the deck) add one
  table column per index. Without the file the slides are as before.
- **Fund-page notes.** `data.nav_afstemning()` (pure) explains why the investor's NAV differs from
  the holdings' fair value in the table: with `FUND_NAV`/`FUND_NAV_EXCL_CARRY` from
  `rpt.v_fund_metric` it splits the gap into accrued carry, the fund's cash and other items, and
  the investor's NAV share vs. the ownership share (the parts sum to the gap); without them the
  note says it cannot be split. Feeders (scaled to NAV) and property funds get no such note.
  `slides._navnote` writes it; parts under 2 % of the gap are left out. A page with an
  underlying-funds table shows no NAV chart or value bridge. `slides._noter_y()` moves long notes
  up (and the content's bottom edge with them) so they stay clear of the logo and footer.
- **Feeder pages.** The Branche/Geografi bars use the same GICS sectors and subregions as the
  cross-fund slides (`slides._feederfordeling` from `f["brancher"]`/`f["regioner"]`); the GP's
  own categories (`KATEGORIER`) are only a fallback.
- **Branche column.** Company tables (header "Branche") and company cards show the GICS sector (`SEKTORER`) from
  `rpt.v_portfolio_holding.industry_code` as `industri`; `sektor` is the GP's own label, kept for
  Excel and the property funds' valuation basis. Unreviewed classifications are counted in the
  fund page's source note.
- **Quarter and year to date.** `data.perioder()` gives the quarter and year windows ending at the
  report date. `data.udvikling()` sets `f["udvikling"][periode]` per fund from `historik` and
  `kapitalstroemme`: NAV primo (latest NAV on or before the start; a gap falls back to an older
  NAV with a note), calls, distributions, value change as the residual, Modified Dietz return.
  `omregn_udvikling_til_dkk()` converts with the rate on each date, and the currency effect is
  the DKK residual. `portefoelje_udvikling()` (pure) sums to `r["udvikling"]`, and
  `vaerdibevaegelser()` (pure) compares holdings with `beholdninger_primo`. Holding names are
  not stable between GP reports, so `_par()`/`_forkortelse_af()` pair renamed lines, and
  anything unpaired goes in `usammenlignelige` instead of being counted as a gain or loss. The
  slides are `udvikling` (after the executive summary) and `aendringer_pr_fond`;
  `vaerdibevaegelser` is kept but not in the deck. *Ændringer pr. fond* has vertical lines between
  the quarter, the year and "Pr. <date>". Its value change is in DKK incl. FX while the return is
  Modified Dietz in the fund's currency, so the two can have opposite signs:
  `data.fortegnsafvigelser()` (pure) finds those cells in `r["fortegnsafvigelser"]` with the cause
  (`valuta`: the fund-currency value change agrees with the return; `dietz`: a negative capital
  base), and `slides._fortegnsnote` footnotes the return cell.
- The *Definitioner* slide before the end slide prints `slides.DEFINITIONER`, plain text that
  describes how `data.py` calculates each figure. Change a calculation, update its definition.
- `deck.py` + `style.py` are the shared PPIM house-style modules, **kept byte-identical to the
  other PE repos** (`pe-analysis`, `pe-quarterly-report`, …). They are excluded from ruff in
  `pyproject.toml`. Do not refactor them here; port a change to the other repos or not at all.
  `deck.py` also holds the slide-master layout indices (`L_FORSIDE`, `L_TEKST`, `L_SKILLE`, …).

**Anonymized deck** (`--anonym`, `report.byg(anonym=True)`): the same report built from the real
data, then rewritten by `anonym.anonymiser()` (pure), so every calculation and the
quarter-to-quarter pairing run on the real names. Funds are named by segment ("Private Equity-fond
1"), holdings "Selskab n" / "Underliggende fond n" / "Ejendom n", managers "Forvalter A", the
investor "Investor"; logos, descriptions, company info, fund size, ownership share and the NAV
reconciliation are removed (the user's choice; vintage and term stay). Amounts are multiplied by
one hidden random factor away from 1 (`report._skjult_faktor`) **at the display layer**
(`slides.beloebsskala`), so ratios, TVPI, IRR and % are exact: every place that prints an amount
must go through `tusind` / `million` / `kroner` / `_kt`, and amounts `data.py` writes into note
text are scaled by regex in `anonym`. No Excel workbook is written. `anonym.laek()` searches all
text in the finished deck (tables, charts, notes pages) for the original names and codes, and
`byg()` raises without saving if any is left — a new name field in `data.fond()` must be added to
`anonym._navnekort`. The PPIM house logo in the master stays.

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
`noter`, not by writing text in the slide builder. A caveat about one figure is a `data.Note`
(a `str` with `felt`: `indbetalt`, `irr`, `andel`, `resttilsagn`, `nav`, `udvikling`,
`beskrivelse`), so the slide can put its footnote mark next to that figure.

**Footnotes.** Every slide collects its notes in a `slides.Fodnoter`: `generel()` for notes about
the whole slide (source, FX rates, date mix; unnumbered, first), `maerke(tekst)` for a note about
one column, row, box or figure — it returns the superscript mark (»¹«) to put there, numbers in
the order marks are set, and reuses the number for the same text. `tegn(s, y)` prints it all
(moving up when long, `_noter_y`). Where the layout depends on the notes' height, set every mark
before drawing and size the content to `_noter_y(fod.tekst())` (see `udvikling`,
`tilsagn_paa_tvaers`, `aendringer_pr_fond`, `fondsside`). Treemap boxes get their marks from
`_kortkasse` (»Øvrige«, Diversificeret, an enlarged box, a box with room only for its %); bands
and the benchmark bar carry the notes for their part. The † date mark stays as it is.

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
- Cash flows carry the LP sign (paid < 0). `WITHHOLDING_TAX` stays out of IRR.
- `EQUALISATION` (late-closing interest) deviates from the data guide by the user's choice:
  paid equalisation counts in paid-in and received in distributions (`data._er_indbetaling` /
  `_er_udlodning`), with a note on the fund page. It is outside the commitment, so called
  capital, discount and resttilsagn are computed without it (like the GP's MM110).
- IRR is `data.xirr` (act/365 bisection) over the investor's cash flows with NAV as the final flow;
  the portfolio IRR is pooled over all funds, not averaged. `data.total()` returns `None` when the
  funds are not all in the same currency — the slides then drop the total row.
- **Call discounts** (Petersen & Partners' feeders: 25 % off the first call) count as paid-in and
  as a distribution on the first call date, as Aleta books them (the user's choice):
  `data.rabat()` gives MM110 − the paid amount, `fond()` adds it to `indbetalt`, `udloddet`,
  `historik` and `kapitalstroemme` (not to the IRR flows: the two legs net out the same day), and
  Excel lists both legs in *Kapitalstrømme*. Resttilsagn is still tilsagn − MM110.
- If the GP's own paid-in figure (CA110 ITD **at the NAV date only**; MM110 is called capital,
  which includes call discounts) differs from the registered capital calls by more than
  `GP_TOLERANCE` (0.5 %), the GP figure wins and a note is added (`data.gp_vinder`) — unless
  the gap equals the recallable distributions: some GPs (MIF3) net those out of CA110, while
  here they count as distributions, so the GP figure would subtract them twice.
- Cash flows after `as_of` are never read (`fond()` filters them), so a report for an earlier
  date is reproducible. `data.afstem_mod_forvalteren` compares resttilsagn with MM120 and
  TVPI with MM160 and adds a build warning (`f["advarsler"]`, printed by `byg()`, not in the
  deck) when they differ; unreviewed `status: udkast` text is warned about the same way.
- DKK conversion (*Investeringer*: resttilsagn and NAV in DKK) uses one
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
- geography and segment (`geography_focus`, `fund_structure`, `strategy_focus`), translated by
  `data.geografi()` / `data.segment()`. Segment: `JUNIOR_DEBT` is "Private credit" (even as a
  fund of funds), then the structure ("Fund of funds", "Secondaries"), then the focus, where
  buyout, venture and growth are all "Private equity" (`size_focus` is not used). The code
  lists are closed by check constraints on `fund.fund`; a new code needs a line in
  `data.GEOGRAFIER` / `STRUKTURER` / `FOKUS`.
- **Fund order.** `investor_rapport()` sorts `r["fonde"]` by `data.strategi_noegle` —
  `STRATEGI_ORDEN` (Private equity, Fund of funds, Secondaries, Private credit, Ejendomme;
  anything else last), then name — so every slide, the fund pages and Excel list funds in that
  order. The segment groups in `fordeling()` follow the same order (not size), so the executive
  summary picks the largest segment with `max`, not the first group.
- løbetid from `term_years` + `extension_years` via `data.loebetid()`, which writes the
  extension years as one-year steps ("10+1+1"); the database holds only the total.
- fund size: the dated `TOTAL_COMMITMENTS` metric first, else `total_commitments`.
- manager and logo: for our feeder companies (`underlying_fund_code` set) the underlying
  fund's manager and logo, not our own. Logos come from `rpt.v_fund_logo` by fund code: the
  fund's own logo where it has one (PD3LGT: LGT, while its manager is AEP), else the manager's.

`config/fund_info.yaml` holds only what the database does not: `beskrivelse` and
`estimerede_cashflows` for notices that never arrived. Each estimate names its `investor`
(cash flows belong to investor × fund; `data.estimerede_cashflows` raises without it), and an
estimate whose amount now appears among the registered cash flows is dropped with a build
warning, so a notice that arrives later is never counted twice. Optional: `benchmark` (shown on
*Lange tilsagn*) and `virksomheder` (per-company description and ownership for the profile
cards; the key must match `name` in `rpt.v_portfolio_holding` exactly). `status: udkast` marks
text written by Claude and not yet reviewed, until someone changes it to `ok`; it is for
internal tracking only and puts no note in the deck (the user's choice). `tests/test_data.py` asserts every entry has `status` and `beskrivelse`,
and that none carries a field the master view now owns.

## Layout reference

`docs/reference/2026 Q2 ATP PEP IX B - Quarterly Internal Report.pdf` is the model for the layout
and the metric set — consult it before changing what a fund page shows.
`assets/ppim_slidemaster.pptx` is the slide master; `tests/test_slides.py` builds against the real
file, so tests break if its layouts move.
