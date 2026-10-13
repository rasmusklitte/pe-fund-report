# PE fund data — guide for analysis

A self-contained description of the data this repository (`pe-external-fund-data`) loads
into SQL Server, for use in other projects that analyse it. It covers what the tables and
views hold, the conventions you must respect (signs, units, currencies, active rows), and
example queries.

State as of **2026-09-30** (migrations V001–V009).

---

## 1. Connection

| | |
|---|---|
| Server | `sql1\ppim` (SQL Server 2019) |
| Database | `PE` |
| Auth | Windows authentication |
| Collation | `Danish_Norwegian_CI_AS` (case-insensitive: `'PetVet' = 'Petvet'`) |
| Schemas to use | **`rpt`** (views, start here), `portfolio` and `fund` (tables) |
| Do not use | `dbo.*`. It belongs to another application (the legacy `pe-database`). The only `dbo` object the views read is `dbo.fx_rates_nationalbanken`, for DKK conversion. |

Treat everything as **read-only** from an analysis repo. All writes go through the `pefund`
CLI in this repository.

```python
# e.g. pyodbc / pandas
import pyodbc, pandas as pd
conn = pyodbc.connect(
    "DRIVER={ODBC Driver 17 for SQL Server};SERVER=sql1\\ppim;DATABASE=PE;Trusted_Connection=yes;"
)
nav = pd.read_sql("SELECT * FROM rpt.v_nav", conn)
```

Money columns are `decimal(19,2)`, and pyodbc returns them as Python `Decimal`. Convert to
float only at the analysis stage.

---

## 2. The big picture

The data comes from PDFs sent by the general partners (GPs) of external PE / private-credit
funds. There are **two strictly separate levels**:

| Level | Perspective | Where | Source documents |
|---|---|---|---|
| **Investor (LP) level** | *Our* investor's position in a fund: cash paid/received, our NAV | `fund.cash_flow`, `fund.capital_account` → `rpt.v_cash_flow`, `rpt.v_nav`, `rpt.v_fund_summary` | capital calls, distribution notices, capital account statements |
| **Fund level** | The *whole fund* and its portfolio | schema `portfolio` → `rpt.v_fund_metric`, `rpt.v_fund_report_summary`, `rpt.v_portfolio_holding`, `rpt.v_portfolio_breakdown`, `rpt.v_portfolio_company` | fund financial statements / quarterly reports |

**Never add fund-level amounts to investor-level amounts.** A fund's NAV in
`v_fund_report_summary` is the whole fund's NAV (often hundreds of millions), while
`v_nav.nav` is our share of it.

Some funds are fund-level only (none of our investors is in them): Capidea III and IV
(`CAP3`, `CAP4`). They have fund-report data but no investor-level data.

### Row validity (applies to every table)

- The GP documents go through a workflow:
  `REGISTERED → EXTRACTED → VALIDATED → APPROVED`, or `REJECTED` / `SUPERSEDED`.
- Facts are only counted when their document is `APPROVED` and the row has `is_active = 1`.
  When a GP restates a document, the new version supersedes the old one, and the old rows
  stay in the table with `is_active = 0`.
- **The `rpt` views apply both filters already.** If you query the base tables directly, add:
  ```sql
  JOIN fund.document d ON d.document_id = x.document_id AND d.status = 'APPROVED'
  WHERE x.is_active = 1
  ```
- Fund-level data: for each fund and report date, only the **latest approved report** is
  active. So the views give one version per fund and date, even if the same report arrived
  through several investors.

---

## 3. Master data (schema `fund`)

| Table | Key | What |
|---|---|---|
| `fund.manager` | `manager_code` | The GP / management company |
| `fund.fund` | `fund_code` | A fund: `strategy`, `currency` (fund currency), `vintage_year` (often NULL) |
| `fund.investor` | `investor_code` | Our investor entities (the LPs) |
| `fund.investment` | `investment_id` | Investor × fund (one investor can be in many funds, and a fund can have several of our investors) |
| `fund.commitment_event` | | Dated commitment changes per investment: `INITIAL`, `INCREASE` (amount +; decreases −). The commitment at date D is `SUM(amount) WHERE event_date <= D`. |
| `fund.account` | `account_code` | Chart of accounts (section 6) |
| `fund.document` | `document_id` | Every source PDF: `doc_type`, `status`, `period_end`, `file_name`, `note` |

`strategy`: `PE` private equity, `PC` private credit, `FOF` fund of funds, `RA` real assets.

### Current funds and investments

| Manager | Fund code | Fund | Strategy | Ccy | Our investor(s) |
|---|---|---|---|---|---|
| NEP (Newbury Partners) | NEP5 | Newbury Equity Partners V L.P. | PE (secondaries) | USD | PPPE |
| NEP | NEP6 | Newbury Equity Partners VI L.P. | PE (secondaries) | USD | PPPE2 |
| PSCP (Park Square Capital) | PSCP4 | Park Square Capital Partners IV, SCSp | PC | EUR | PPPC |
| PSCP | PSCP5 | Park Square Capital Partners V, SCSp | PC | EUR | PPPC2 |
| NAVIGARE | MIF3 | Maritime Investment Fund III K/S | RA | USD | PPMAR |
| ATPPEP | ATPPEP9 | ATP Private Equity Partners IX B K/S | FOF | EUR | PPFOF |
| CATACAP | CATACAP3 | CataCap III K/S | PE | DKK | MOTORTRAMP |
| SEED | SEED4 | Seed Capital Denmark IV K/S | PE | DKK | MOTORTRAMP |
| CAPIDEA | CAP3, CAP4 | Capidea Kapital III / IV K/S | PE | DKK | — (fund level only) |

Investors: `PPPE` Petersen & Partners Private Equity A/S, `PPPE2` … Private Equity II A/S,
`PPPC` … Private Credit A/S, `PPPC2` … Private Credit II A/S, `PPMAR` … Maritime A/S,
`PPFOF` Petersen & Partners FoF K/S, `MOTORTRAMP` A/S Motortramp.

Query the master tables rather than hard-coding this list, since it grows.

---

## 4. Investor-level views

### 4.1 `rpt.v_cash_flow` — every cash flow of our investors

One row per cash-flow component (e.g. a capital call split into investments, management
fee and expenses).

| Column | Meaning |
|---|---|
| `fund_code`, `fund_name`, `manager_code`, `strategy`, `investor_code`, `investment_id` | who / where |
| `value_date` | date the cash moves (use this for XIRR and time series) |
| `notice_date` | date of the notice |
| `account_code`, `account_name` | component (`CF1xx` calls, `CF2xx` distributions; section 6) |
| `flow_class` | `CONTRIBUTION`, `DISTRIBUTION`, `WITHHOLDING_TAX`, `EQUALISATION` (below) |
| `is_recallable` | 1 for recallable distributions (`CF230`) |
| `amount`, `currency` | **LP cash perspective: paid to the fund < 0, received > 0**, in fund currency |
| `fx_rate`, `amount_dkk` | DKK per 1 unit of currency (Danmarks Nationalbank, latest rate on/before `value_date`) and the DKK amount |
| `flow_ref` | GP's notice reference (e.g. `PSCP4 Drawdown 27 2025-03-20`) |
| `document_id`, `file_name` | source |

`flow_class` rules:
- **CONTRIBUTION**: `CF100`–`CF199` except `CF130`. This includes `CF150` refunds, which are positive. Paid-in = `-SUM(amount)`.
- **DISTRIBUTION**: `CF200`–`CF299` except `CF240` and `CF280`. **Gross**, as the GPs report it.
- **WITHHOLDING_TAX**: `CF280`, tax withheld from distributions (negative). It is paid on the investor's behalf, and usually a tax credit. It is not deducted from DPI.
- **EQUALISATION**: `CF130` / `CF240`, late-closing interest paid / received. It is outside the commitment, so leave it out of paid-in and distributions.

### 4.2 `rpt.v_nav` — our NAV per statement date

One row per investment and capital-statement date (quarter ends).

| Column | Meaning |
|---|---|
| `period_end` | statement date |
| `nav` | closing NAV (account `CA900`, net of carried interest) as reported |
| `restatement` | year-end restatement (`MM180`), only on year-end dates (see below) |
| `restated_nav` | `nav + restatement`, the audited year-end NAV |
| `commitment` | commitment in force at `period_end` |
| `fx_rate`, `nav_dkk`, `restated_nav_dkk` | DKK conversion at `period_end` |

**Year-end restatements.** GPs revalue after the audit, so the next year's opening NAV can
differ from the Q4 closing NAV that was reported first. The difference is stored as
`restatement` on the prior year-end date. The reported figures are never changed. Use
`restated_nav` for the audited view, and `nav` for "as first reported".

### 4.3 `rpt.v_fund_summary` — key figures per investment (latest NAV date)

One row per investment, as of its latest NAV date (`as_of`). Only cash flows up to
`as_of` count.

| Column | Meaning |
|---|---|
| `commitment`, `paid_in`, `distributions`, `recallable` | fund currency, positive numbers |
| `unfunded` | `commitment − paid_in + recallable` |
| `nav` | NAV at `as_of` |
| `equalisation_interest`, `withholding_tax` | net amounts (LP sign) |
| `dpi`, `rvpi`, `tvpi` | distributions / paid-in, NAV / paid-in, (distributions + NAV) / paid-in |
| `paid_in_dkk`, `distributions_dkk`, `nav_dkk` | DKK at each flow's own date / at `as_of` |
| `first_cash_flow` | first value date |

**IRR is not in the views** (it needs iteration). The standard definition used here: XIRR
over `v_cash_flow` rows with `flow_class <> 'WITHHOLDING_TAX'` and `value_date <= as_of`,
plus the NAV as a final positive flow on `as_of` (actual/365). Example in section 9.

### 4.4 Base tables behind them (for detail)

- `fund.cash_flow`: `investment_id, document_id, account_id, notice_date, value_date, amount (LP sign), currency, flow_ref, is_active`.
- `fund.capital_account`: the capital-account statement lines per `investment_id, period_end, basis, account_id, amount, currency, is_active`.
  - `basis`: `QTD`, `YTD`, `ITD` (as printed in the statement), or `POINT` (memo balances and restatements at a date).
  - Sign: movements as they affect the capital account (contributions +, distributions −, losses −), so
    `CA100 opening + SUM(accounts with in_nav_rollforward = 1) = CA900 closing` for each basis.
  - Use it for NAV bridges (value creation vs. fees vs. flows) and for the GP-reported memo figures (`MM1xx`).
  - There is no view: filter it yourself (see section 2).

---

## 5. Fund-level views (schema `portfolio`)

All amounts are in **full units** (millions and thousands are already scaled) and belong
to the **whole fund**. `report_date` is the report's quarter end.

### 5.1 `rpt.v_fund_metric` — key figures from fund reports (long format)

`fund_code, fund_name, report_date, metric_code, value, unit, currency, source_label, document_id`

- `unit`:
  - `CCY`: an amount in `currency`
  - `PCT`: percent (`9.4` = 9.4 %)
  - `X`: a multiple (`1.14` = 1.14x)
  - `COUNT`
- Metrics differ per GP, because each reports its own set. Coverage (2026-09-30):

| metric_code | unit | funds |
|---|---|---|
| FUND_NAV | CCY | ATPPEP9, CAP3, CAP4, MIF3, NEP5, NEP6 |
| TOTAL_COMMITMENTS, PAID_IN, DISTRIBUTED, UNFUNDED | CCY | ATPPEP9, CAP3, CAP4, MIF3 |
| RECALLABLE_DISTRIBUTIONS | CCY | ATPPEP9, MIF3 |
| PORTFOLIO_FAIR_VALUE | CCY | ATPPEP9, CAP3, CAP4, MIF3 |
| NET_IRR | PCT | ATPPEP9, CAP3, CAP4, MIF3, PSCP4, PSCP5 |
| NET_TVPI, NET_DPI | X | ATPPEP9, CAP3, CAP4, MIF3, PSCP4, PSCP5 |
| NET_RVPI | X | CAP3, CAP4, MIF3 |
| GROSS_IRR | PCT | CAP3, CAP4, PSCP4, PSCP5 |
| GROSS_MOIC | X | CAP3, CAP4 |
| ASSET_IRR | PCT | PSCP4, PSCP5 |
| PICC | PCT (ATP, Capidea) / **X (MIF3)** | ATPPEP9, CAP3, CAP4, MIF3 |
| ACCRUED_CARRY, CASH | CCY | CAP3, CAP4 |
| CASH_AND_OTHER_NET | CCY | ATPPEP9, CAP3, CAP4 |
| CREDIT_LINE_DRAWN | CCY | MIF3, NEP5, NEP6 |
| CALLED_FROM_LPS, COMMITTED_TO_DATE, SECONDARY_VALUE, UNDERLYING_UNFUNDED | CCY | NEP5, NEP6 |
| CALLED_PCT | PCT | NEP5, NEP6 |
| DISTRIBUTED_PCT_OF_CALLED, SECONDARY_FUNDED_PCT | PCT | NEP5 |
| NUMBER_OF_FUNDS, NUMBER_OF_COMPANIES | COUNT | NEP5, NEP6 |
| INVESTED_CAPITAL, REALISED_PROCEEDS, UNREALISED_VALUE, TOTAL_VALUE, INVESTMENT_PROFIT | CCY | PSCP4, PSCP5 (track-record totals) |
| NUMBER_OF_INVESTMENTS | COUNT | PSCP4, PSCP5 |
| AVG_INVESTMENT_SIZE, WA_EBITDA, WA_ENTERPRISE_VALUE | CCY | PSCP4, PSCP5 |
| WA_EQUITY_CUSHION, WA_YIELD_TO_3Y_CALL | PCT | PSCP4, PSCP5 |
| DEBT_TO_EBITDA | X | PSCP4, PSCP5 |
| ALLOCATED_COMMITMENT, UNALLOCATED_COMMITMENT, CALLED_FOR_INVESTMENTS, CALLED_FOR_OPERATIONS, CUMULATIVE_NET_PROFIT, CURRENT_LIABILITIES, RECEIVABLES_AND_OTHER | CCY | MIF3 |
| NET_IRR_EXCL_CARRY | PCT | MIF3 |

Watch the `unit` column (PICC is a percent for some GPs and a multiple for MIF3).

### 5.2 `rpt.v_fund_report_summary` — one row per fund (latest report)

A wide pivot of the latest report date's metrics: `total_commitments, paid_in, distributed,
unfunded, recallable, portfolio_fair_value, accrued_carry, fund_nav, net_dpi, net_rvpi,
net_tvpi, picc_pct, net_irr_pct, gross_irr_pct, gross_moic`, plus `manager_code, strategy,
vintage_year, currency, report_date`. `our_investors` lists our investor codes in the fund,
and is NULL for fund-level-only funds (Capidea). A metric the GP doesn't report is NULL.

### 5.3 `rpt.v_portfolio_holding` — the fund's holdings per report date

One row per holding line in the GP's portfolio table.

| Column | Meaning |
|---|---|
| `section` | NEP: `SECONDARY` / `CO_INVESTMENT` / `PRIMARY`. PSCP: `REALISED` / `UNREALISED` (track record; one row per tranche). Capidea: `CURRENT` / `REALISED`. MIF3: `CURRENT`. ATP: `FUND` (underlying funds). |
| `name` | holding (company, fund or SPV) as printed |
| `sponsor`, `transaction_ref`, `instrument`, `segment`, `geography`, `vintage`, `trade_date`, `exit_date` | where the GP gives them (PSCP: `instrument`, `trade_date`, `exit_date` (month, stored as first of month) and `local_currency`; ATP: `segment`, `vintage`) |
| `currency` | currency of the amounts: the fund currency, **except ATP**, whose holdings are in each fund's **local currency** (`local_currency` = `currency`) |
| `total_capitalization, commitment, invested, distributions, fair_value, total_value, unfunded, gain_loss` | amounts, full units |
| `paid_in_pct`, `irr_pct` | percent |
| `dpi`, `multiple` | multiples (x) |

Coverage: NEP5 2020-06 → 2026-06, NEP6 2024-03 →, PSCP4 2019-12 →, PSCP5 2023-12 →,
CAP3 2019-09 → 2026-03, CAP4 2023-12 → 2026-03, MIF3 2025-06 →, ATPPEP9 2025-12 →.

Holding names are as printed and can vary between reports. PSCP's track record has one
row **per tranche**, so group by `name` for company totals.

### 5.4 `rpt.v_portfolio_breakdown` — allocation splits per report date

`fund_code, report_date, dimension, category, share_pct, amount, currency, basis`

- `dimension`:
  - PSCP: `INDUSTRY`, `GEOGRAPHY`, `CURRENCY`, `TOP_ISSUER` (with `amount`)
  - ATPPEP9: `GEOGRAPHY`, `SEGMENT`
- `share_pct`: percent. The shares of one dimension on one date sum to ≈100, except `TOP_ISSUER`, which covers only the top 10.
- `basis` says what the share is of (e.g. "share of portfolio fair value").
- Some reports print their charts as images. Those dates have no breakdown rows (e.g. PSCP Q1 2026).

### 5.5 `rpt.v_portfolio_company` — portfolio-company key statistics (PSCP)

One row per portfolio company per report, from the "Key statistics" boxes in Park
Square's reports (PSCP4 2019-12 → 2026-06, 27 reports; PSCP5 2023-12 → 2026-06, 10
reports; 755 rows).

| Column | Meaning |
|---|---|
| `company` | **use this to group over time**: one name per company (renames and spellings are merged via `portfolio.company_alias`) |
| `company_name` | as printed in that report |
| `headquarters`, `sponsor` | as printed |
| `tranche` | the fund's instruments in the company, as printed: `"PIK(€) / Equity(€)"` |
| `invested_amount`, `invested_currency` | the fund's amount invested in the company (all tranches), EUR |
| `kpi_basis` | `LTM` or `FY` |
| `kpi_period_end` | month end of the company figures' period, usually 1–2 quarters before `report_date` |
| `kpi_currency` | currency of revenue / EBITDA / EV (the company's own: EUR, USD, GBP, SEK, CHF) |
| `revenue`, `ebitda`, `enterprise_value` | the **company's** figures, full units of `kpi_currency` |
| `ebitda_margin_pct` | `100 × ebitda / revenue` |
| `ev_to_ebitda` | `enterprise_value / ebitda` |
| `net_leverage` | net debt / EBITDA (x), as printed |
| `net_senior_leverage` | net senior debt / EBITDA, used instead of `net_leverage` in PSCP IV Q4 2025 and PSCP V Q1 2026 for some companies ("Net Sr Debt"). Use `COALESCE(net_leverage, net_senior_leverage)` only if you accept mixing the two. |
| `leverage_at_investment` | from the GP's text "Total net leverage is 7.7x, down from 9.0x at investment" (NULL where not stated) |
| `commentary` | the GP's description and quarterly commentary for the company |

Notes:
- The GP seems to set EV at a constant multiple of EBITDA per company (e.g. Barentz 13.0x
  every quarter), so `ev_to_ebitda` reflects the GP's valuation method more than the market.
- A figure the GP prints as "n.a." is NULL (e.g. Dechra's net debt in 2023).
- PSCP V Q2 2025 has no rows (its company pages are scanned images).
- Company figures are not in EUR: convert with `kpi_currency` before summing across companies.
- The base table `portfolio.company_statistic` also has `raw_text`, the whole box as printed.
  For 2019–2021 reports this includes per-tranche Total Value / MM / entry price / margin.
- Aliases in `portfolio.company_alias` (`alias_name → company_name`): Pharmazell → Axplora,
  Bosch Packaging → Syntegon, BASF → MBCC Group, Hansen Colors → Oterra,
  Thyssenkrupp / Thyssenkrupp Elevators → TKE, IntraFi → IntraFi Network,
  PetVet and DiversiTech spelling.

---

## 6. Chart of accounts (`fund.account`)

`account_group`:
- `CASHFLOW` (`cash_direction`: IN / OUT)
- `CAPACC` (capital account; `in_nav_rollforward` = 1 for movements)
- `MEMO`

| Code | Name |
|---|---|
| CF100 | Capital call – investments |
| CF110 | Capital call – management fee |
| CF120 | Capital call – fund and organisational expenses |
| CF130 | Capital call – equalisation / late-closing interest |
| CF150 | Contribution refund – equalisation true-up (reduces paid-in) |
| CF190 | Capital call – other |
| CF199 | Capital call – not split |
| CF200 | Distribution – return of capital |
| CF210 | Distribution – realised gain |
| CF220 | Distribution – income (interest / dividend) |
| CF230 | Distribution – recallable |
| CF240 | Distribution – equalisation received |
| CF280 | Withholding tax deducted |
| CF290 | Distribution – other |
| CF295 | Distribution – not split |
| CA100 | Opening NAV |
| CA110 | Contributions |
| CA120 | Distributions |
| CA200 | Net investment / operating income |
| CA210 | Management fees |
| CA220 | Fund expenses |
| CA230 | Financial items / financing costs |
| CA290 | Net result from operations – not split |
| CA300 | Realised gain / (loss) |
| CA310 | Change in unrealised gain / (loss) |
| CA490 | Other capital account movements |
| CA400 | Carried interest adjustment (a roll-forward movement; NEP, MIF) |
| CA890 | NAV before carried interest adjustment |
| CA900 | **Closing NAV (net of carried interest)** |
| MM100 | Commitment (as stated in the document) |
| MM110 | Cumulative contributions |
| MM120 | Unfunded commitment |
| MM130 | Cumulative distributions |
| MM140 | Recallable distributions available |
| MM150 | Fund-reported net IRR (%) |
| MM160 | Fund-reported TVPI (x) |
| MM170 | Commitment reserved (secured commitments / guarantees; Maritime) |
| MM180 | Restatement of prior year-end NAV |
| MM200 | Accrued carried-interest balance (memo; PSCP, ATP, where NAV is already net of carry) |

Split granularity differs by GP: some split calls into CF100/110/120 and others only give
CF199 ("not split"). Compare across funds at `flow_class` level, not account level.

---

## 7. Currency

- Investor-level amounts are in the **fund currency** (`currency` column). `*_dkk`
  columns convert with Danmarks Nationalbank rates (`dbo.fx_rates_nationalbanken`:
  `rate_date, currency, dkk_per_1`), using the latest rate on or before the date.
- Fund-level amounts: `currency` on each row. That is the fund currency, except ATP
  holdings (local currency) and the company KPIs (`kpi_currency`).
- To convert other things yourself:
  ```sql
  OUTER APPLY (SELECT TOP 1 r.dkk_per_1 FROM dbo.fx_rates_nationalbanken r
               WHERE r.currency = x.currency AND r.rate_date <= x.some_date
               ORDER BY r.rate_date DESC) fx
  ```
  For cross rates (e.g. GBP → EUR) divide the two `dkk_per_1` values.
- **DKK has no rows in the rate table**, so for the DKK funds (CATACAP3, SEED4, CAP3, CAP4)
  `fx_rate` and every `*_dkk` column are **NULL**. Use `COALESCE(amount_dkk, amount)` when
  `currency = 'DKK'` (rate 1). This also applies to `v_fund_summary.*_dkk` for those funds.

---

## 8. Coverage and known gaps (2026-09-30)

- 8 investments with investor-level data. The quarterly NAV series have no gaps, except:
  - CataCap III: no reports Q3 2023 – Q1 2024
  - Seed Capital IV: no reports for 2023 Q4, 2024 Q2 and 2025 Q4; draw downs 14 and 16 are missing
- NAV series start: NEP5 2020-06, PSCP4 2021-03, SEED4 2022-03, CATACAP3 2023-06,
  PSCP5 2023-12, NEP6 2024-03, MIF3 2025-09, ATPPEP9 2025-12.
- Cash flows start earlier than the NAV series for some funds (e.g. PSCP4 from 2021-02,
  NEP5 from 2020-03). Latest cash flows can be after the latest NAV date, and
  `v_fund_summary` ignores them.
- Fund reports: 81 approved for our funds (full history, no gaps in any fund's
  report-date series), plus Capidea III (33 quarters from Q1 2018) and IV (12 quarters
  from Q2 2023).
- Some GPs switch reporting templates, so a metric can appear or disappear between dates.
- `document.note` records any manual decisions: e.g. approval despite a GP table that
  doesn't add up (NEP5 Mar 2022, PSCP4 Dec 2025).

---

## 9. Example queries

**NAV time series (restated) with DKK**
```sql
SELECT fund_code, investor_code, period_end, restated_nav, currency, restated_nav_dkk
FROM rpt.v_nav
ORDER BY fund_code, investor_code, period_end;
```

**Paid-in, distributions and TVPI over time for one investment**
```sql
SELECT n.fund_code, n.period_end, n.nav,
       -SUM(CASE WHEN c.flow_class = 'CONTRIBUTION' THEN c.amount END) AS paid_in,
        SUM(CASE WHEN c.flow_class = 'DISTRIBUTION' THEN c.amount END) AS distributions,
       (COALESCE(SUM(CASE WHEN c.flow_class = 'DISTRIBUTION' THEN c.amount END), 0) + n.nav)
         / NULLIF(-SUM(CASE WHEN c.flow_class = 'CONTRIBUTION' THEN c.amount END), 0) AS tvpi
FROM rpt.v_nav n
JOIN rpt.v_cash_flow c ON c.investment_id = n.investment_id AND c.value_date <= n.period_end
WHERE n.fund_code = 'NEP5'
GROUP BY n.fund_code, n.period_end, n.nav
ORDER BY n.period_end;
```

**XIRR per investment (Python)**
```python
from datetime import date
flows = pd.read_sql("SELECT investment_id, value_date, amount FROM rpt.v_cash_flow "
                    "WHERE flow_class <> 'WITHHOLDING_TAX'", conn)
summary = pd.read_sql("SELECT investment_id, fund_code, as_of, nav FROM rpt.v_fund_summary", conn)

def xirr(cfs):  # cfs: list of (date, amount); actual/365, bisection
    t0 = min(d for d, _ in cfs)
    f = lambda r: sum(a / (1 + r) ** ((d - t0).days / 365) for d, a in cfs)
    lo, hi = -0.99, 10.0
    for _ in range(200):
        mid = (lo + hi) / 2
        lo, hi = (mid, hi) if f(lo) * f(mid) > 0 else (lo, mid)
    return (lo + hi) / 2

for s in summary.itertuples():
    cf = flows[(flows.investment_id == s.investment_id) & (flows.value_date <= s.as_of)]
    series = [(d, float(a)) for d, a in zip(cf.value_date, cf.amount)] + [(s.as_of, float(s.nav))]
    print(s.fund_code, round(xirr(series) * 100, 2))
```

**Fund-level metrics as a wide time series**
```sql
SELECT fund_code, report_date,
       MAX(CASE WHEN metric_code = 'NET_IRR'  THEN value END) AS net_irr_pct,
       MAX(CASE WHEN metric_code = 'NET_TVPI' THEN value END) AS net_tvpi,
       MAX(CASE WHEN metric_code = 'NET_DPI'  THEN value END) AS net_dpi
FROM rpt.v_fund_metric
GROUP BY fund_code, report_date
ORDER BY fund_code, report_date;
```

**Portfolio-company leverage and margin over time (PSCP)**
```sql
SELECT fund_code, company, report_date, kpi_period_end, kpi_currency,
       revenue / 1e6 AS revenue_m, ebitda / 1e6 AS ebitda_m, ebitda_margin_pct,
       net_leverage, leverage_at_investment,
       net_leverage - leverage_at_investment AS leverage_change_since_entry
FROM rpt.v_portfolio_company
WHERE company = 'Barentz'
ORDER BY fund_code, report_date;
```

**LTM EBITDA growth per company between consecutive reports**
```sql
SELECT fund_code, company, report_date, kpi_period_end, ebitda,
       100.0 * (ebitda / NULLIF(LAG(ebitda) OVER (PARTITION BY fund_code, company
                                                ORDER BY report_date), 0) - 1) AS ebitda_growth_pct
FROM rpt.v_portfolio_company
WHERE kpi_basis = 'LTM';
```
The KPI period doesn't always move every quarter (the same LTM figure can be repeated), so
compare on `kpi_period_end` changes if you need clean growth rates.

**Concentration: share of fair value in the top 5 holdings (latest date per fund)**
```sql
WITH latest AS (SELECT fund_code, MAX(report_date) d FROM rpt.v_portfolio_holding GROUP BY fund_code),
h AS (
  SELECT h.fund_code, h.name, SUM(h.fair_value) fv
  FROM rpt.v_portfolio_holding h JOIN latest l ON l.fund_code = h.fund_code AND l.d = h.report_date
  WHERE h.fair_value IS NOT NULL
  GROUP BY h.fund_code, h.name)
SELECT fund_code,
       100.0 * SUM(CASE WHEN rk <= 5 THEN fv END) / SUM(fv) AS top5_share_pct
FROM (SELECT *, ROW_NUMBER() OVER (PARTITION BY fund_code ORDER BY fv DESC) rk FROM h) x
GROUP BY fund_code;
```
(ATP holdings are in local currencies, so convert before summing across currencies.)

**Capital-account bridge (value creation vs. fees) for a year**
```sql
SELECT f.fund_code, ca.period_end, a.account_code, a.name, ca.amount
FROM fund.capital_account ca
JOIN fund.document d ON d.document_id = ca.document_id AND d.status = 'APPROVED'
JOIN fund.account a ON a.account_id = ca.account_id
JOIN fund.investment x ON x.investment_id = ca.investment_id
JOIN fund.fund f ON f.fund_id = x.fund_id
WHERE ca.is_active = 1 AND ca.basis = 'YTD' AND a.account_group = 'CAPACC'
  AND f.fund_code = 'PSCP4' AND ca.period_end = '2025-12-31'
ORDER BY a.account_code;
```

---

## 10. Pitfalls checklist

- [ ] Use `rpt` views, or filter base tables on `APPROVED` + `is_active = 1`.
- [ ] LP level (`v_nav`, `v_cash_flow`, `v_fund_summary`) ≠ fund level (`portfolio`, `v_fund_*`, `v_portfolio_*`). Don't mix.
- [ ] Cash-flow sign: paid < 0, received > 0. Paid-in = `-SUM(CONTRIBUTION)`.
- [ ] Leave out `EQUALISATION` from paid-in / distributions; leave `WITHHOLDING_TAX` out of DPI and XIRR.
- [ ] Use `restated_nav` for audited year-ends.
- [ ] Check `unit` in `v_fund_metric` (PCT is percent, X is a multiple, and PICC varies).
- [ ] Amounts are full units (not thousands or millions); convert currencies before summing.
- [ ] `*_dkk` columns are NULL for DKK funds (no DKK rate row): use the amount itself.
- [ ] Company KPIs are in `kpi_currency` for the period ending `kpi_period_end`, not `report_date`.
- [ ] Group portfolio companies on `company`, not `company_name`.
- [ ] `vintage_year` is often NULL; don't rely on it.
- [ ] Money is `Decimal`; don't do accounting in float.
