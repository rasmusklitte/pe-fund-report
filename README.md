# pe-fund-report

Builds a PowerPoint report of the private equity funds held by one investor, on the PPIM slide
master. The layout and metrics are modelled on ATP PEP's quarterly internal report
(`docs/reference/`).

The deck contains:

1. Front page
2. Portfolio overview: one row per fund with commitment, paid-in, PICC, distributions, NAV, DPI,
   TVPI and IRR, and a total row
3. Distribution by geography and by segment
4. One page per fund: description, fund identity, the investor's commitment, key figures, the
   manager's logo, and either the fund's holdings or NAV history with a value bridge

All tables, text and charts are native PowerPoint objects and can be edited in the deck. The
report text is Danish.

## Usage

```
python -m pe_fund_report MOTORTRAMP                      # latest NAV date
python -m pe_fund_report MOTORTRAMP --as-of 2025-12-31
python -m pe_fund_report MOTORTRAMP --out exports/test.pptx
```

The argument is the `investor_code` in `fund.investor`. The result is written to
`exports/<INVESTOR>_<date>.pptx`. If that file is open in PowerPoint, the build writes
`..._v2.pptx` (then `_v3`, …) instead of failing, and says so.

## Setup

Requires Python 3.11+, the "ODBC Driver 18 for SQL Server" and Windows access to the PE database.

```
pip install -e .[dev]
```

This also installs the `pe-fund-report` command, equivalent to `python -m pe_fund_report`.

The database defaults to `sql1\ppim` / `PE` with Windows authentication and can be overridden with
`PE_DB_SERVER`, `PE_DB_DATABASE` and `PE_DB_DRIVER`. The repo only reads: `db.read_sql` rejects
anything but `SELECT` / `WITH`.

## Layout

```
pe_fund_report/        the package
  cli.py               command line
  report.py            assembles the deck
  data.py              everything read from the database, and the calculations (XIRR, TVPI, ...)
  slides.py            slide builders: fund page, portfolio overview, distributions
  db.py                read-only database access
  paths.py             locations of config, slide master and exports
  deck.py, style.py    shared PPIM house style (kept identical to the other PE repos)
config/fund_info.yaml  master data the database does not hold
assets/                PPIM slide master
docs/                  database guide, and the reference report in docs/reference/
tests/                 tests that run without the database
exports/               built reports (not in git)
```

## Manual master data

`config/fund_info.yaml` holds what the database does not: fund description, geography, segment,
fund size where no `TOTAL_COMMITMENTS` is loaded, and estimated cash flows for notices that are
missing. Entries marked `status: udkast` are unreviewed drafts and put a footnote on the slide;
change the status to `ok` once reviewed.

## Data rules

The calculations follow `docs/analysis-data-guide.md`:

- `rpt` views first; base tables only with `status = 'APPROVED'` and `is_active = 1`.
- Cash flows carry the LP sign (paid < 0). Equalisation is left out of paid-in and distributions,
  withholding tax out of IRR.
- IRR is XIRR (act/365) on the investor's cash flows with NAV as the final flow. The portfolio
  IRR is pooled over all funds.
- If the GP's own paid-in figure differs from the registered capital calls by more than 0.5 %,
  the GP figure is used and the slide gets a footnote.
- Holdings are whole-fund figures scaled by the investor's ownership share; they are shown
  unscaled and labelled as such when the fund size is unknown.
- Manager logos come from `rpt.v_manager_logo`; a logo made for dark backgrounds is placed on a
  dark plate.

## Tests

```
pytest
```
