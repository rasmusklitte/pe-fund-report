# Existing content in the `PE` database

Read-only inspection, 2026-09-29. **Nothing here is used by this project**, except that the `rpt` views
may read `dbo.fx_rates_nationalbanken` for FX (see CLAUDE.md).

## Server

| | |
|---|---|
| Server | `sql1\ppim` (reports as `SQLSERVER\PPIM`) |
| Version | SQL Server 2019 Standard, 15.0.2190.7, on Windows Server 2022 |
| Database | `PE`, compatibility level 140, collation `Danish_Norwegian_CI_AS` |
| Auth | Windows (`PEOPLEGROUP\rka`), ODBC Driver 18 |
| Rights | CREATE TABLE / VIEW / PROCEDURE / SCHEMA, ALTER, CONTROL |

Other databases on the server (`FundMarket`, `PPIM`, `manager`, `BONDS`, `Equity`, `JOLAR`, …) are
unrelated to this project.

## Objects (all in `dbo`, ~96)

They come from the `pe-database` app (menu-driven data entry, Danish column names). Characteristics:
funds are keyed by a free-text `FOND` column, NAV is stored as `float`, and there are `_old` / `_test`
copies of the main tables.

**Tables:**

| Table | Rows | Content |
|---|---:|---|
| `tbl_pe_transactions` (+ `_old`, `_test`) | 1,417 | Investor-level transactions (COMMITMENT, DRAWDOWN, DISTRIBUTION, …) |
| `tbl_pe_fund_transactions` (+ `_old`, `_underlying`) | 284 | Fund-level transactions: `FOND`, `DATO`, `EVENT`, `BELOEB` |
| `tbl_pe_fund_commitments` | 42 | `FOND`, `DATO`, `EVENT`, `TILSAGN` |
| `tbl_pe_fund_nav` | 79 | `DATO`, `FOND`, `EGENKAPITAL_VAERDI` (float), `VALUTA`, `VALUTA_KURS` |
| `tbl_pe_stamdata_fund` | 6 | Fund master data (see below) |
| `tbl_pe_stamdata`, `tbl_pe_subscription`, `tbl_pe_emails` | 115 / 177 / 189 | End-investor master data and subscriptions |
| `tbl_pe_faelles_events` | 115 | Shared events |
| `tbl_pe_transaction_log`, `tbl_pe_transaction_extended_log` | ~100 | Audit logs |
| `tbl_pe_fond_mapping` | 4 | `FOND` → identifier |
| `pe_irr` | 3,802 | IRR cache |
| `fx_rates_nationalbanken` | 20,875 | Daily FX from Nationalbanken: `rate_date`, `currency`, `dkk_per_100`, `dkk_per_1`, `source`, `ingested_at_utc` — **live feed, allowed for read** |
| `tbl_Posteringer_UC`, `Uniconta Posteringer PPPE` | 839 / 0 | 2021 accounting postings |

**Code objects:**
- About 35 procedures (`usp_insert_*`, `usp_update_*`, `usp_pe_prognosis_*`, `usp_refresh_pe_irr_cache`, …).
- 11 views (`vw_pe_dim_fund`, `vw_pe_fact_nav`, `vw_pe_events`, `vw_pe_irr`, …).
- The function `fn_IRR`.

## Fund structure seen in `tbl_pe_stamdata_fund`

Our LP vehicles and the external funds they invest in. These became the V001 seed rows; no amounts
were copied.

| LP vehicle (investor_code) | External fund (fund_code) | Ccy |
|---|---|---|
| Petersen & Partners Private Equity A/S (PPPE) | Newbury Equity Partners V L.P. (NEP5) | USD |
| Petersen & Partners Private Equity II A/S (PPPE2) | Newbury Equity Partners VI L.P. (NEP6) | USD |
| Petersen & Partners Private Credit A/S (PPPC) | Park Square Capital Partners IV, SCSp (PSCP4) | EUR |
| Petersen & Partners Private Credit II A/S (PPPC2) | Park Square Capital Partners V, SCSp (PSCP5) | EUR |
| Petersen & Partners FoF K/S (PPFOF) | ATP Private Equity Partners IX B K/S (ATPPEP9) | EUR |
| Petersen & Partners Maritime A/S (PPMAR) | Maritime Investment Fund III K/S (MIF3) | USD |
