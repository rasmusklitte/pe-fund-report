# PDF inventory — what the fund documents contain and how to extract them

Reconnaissance of 56 sample PDFs (2026-09-29): 2–3 recent capital accounts, capital calls and
financial statements / quarterly reports per fund. **Every file has a text layer — no OCR needed.**

## Summary

| Fund | Document | Pages | Layout stable? | Method | Notes |
|---|---|---:|---|---|---|
| NEP5 / NEP6 | Capital account ("Reconciliation of Partners' Capital Account") | 1 | Yes | **Profile** | YTD + cumulative; read **our** column, see §NEP |
| NEP5 / NEP6 | Capital call letter | 1–2 | Two variants (letter; "Bridge" platform from 2026 for NEP VI) | **Profile** (regex) | One amount, no split |
| PSCP4 / PSCP5 | Partner's capital statement (ILPA template) | 6–7 | Yes | **Profile** | QTD/YTD/ITD, full fee detail, cash-flow history, GP multiples |
| PSCP4 / PSCP5 | Drawdown notice (incl. "partially deemed") | 3 | Yes | **Profile** (regex) | Gross contribution, netted distribution, net payment |
| ATPPEP9 | Capital account | 1–2 | Two variants (wide table ≤ Q1 2026; vertical from Q2 2026) | **Profile** | **EUR thousands**, Danish number format |
| ATPPEP9 | Net capital call | 2 | Yes (distribution section optional) | **Profile** | Full component split, Danish number format |
| MIF3 | Investor statement | 2 | Yes (net-NAV label changed after Q3 2025) | **Profile** (cell engine) | QTD/YTD/ITD, dated cash-flow list |
| MIF3 | Deposit notice | 3 | Yes | **Profile** (regex) | Component split with "Your share" |
| All | Financial statements / quarterly reports | 20–66 | Fund-specific | **Claude** (later) | Fund level only; portfolio holdings |

**Conclusion:** all LP-level data (capital accounts, calls) can be extracted deterministically.
Claude is only needed for the fund-level reports (portfolio overview, fund key figures), which are not
required for our NAV and cash-flow reporting.

## Extraction engine requirements learned from the samples

1. **Two text back-ends.**
   - `pypdf` gives clean numbers for NEP, where pdfplumber splits digits ("5 5,000,000"). But pypdf loses
     row order on the Maritime statement.
   - pdfplumber **words with x-positions** reconstruct any table correctly. The primary engine is therefore:
     words → rows (by `top`) → cells (by column x-ranges taken from header words) → join tokens → parse.
   - This also fixes detached minus signs (Maritime Q3 2025: `- 2 91,474` = −291,474, while a lone `-` = 0).
2. **Letters use regex** on whitespace-collapsed text. NEP VI splits tokens over lines
   ("January \n14\n, \n2026"), so collapse the text first.
3. **Number format per profile.**
   - ATP uses Danish format (`1.632`, `5.435.346`, `4,9%`); everyone else uses English format.
     `parse_amount` guesses, but a single-dot value like `1.632` is ambiguous, so the profile must state
     `number_format: da|en`.
4. **Unit per profile.** The ATP capital account is in **EUR thousands** (`unit: 1000`); the rest are in
   whole currency units. The rounding (±500 per figure) needs a wider validation tolerance for ATP.
5. **Column selection by header, not position.** The investor column is found by matching the investor
   alias in the header (NEP: "Petersen & Partners" vs "Total Partnership"; PSCP: "LP's Allocation of
   Total Fund"; MIF: "Petersen & Partners Maritime A/S"; ATP: the row starting with the investor name).
6. **Layout variants per profile.** A profile may hold several variants, each recognised by anchor
   texts (ATP capital account; NEP VI call; MIF net-NAV label).
7. **Investor aliases seen:**
   - "Petersen & Partners Private Equity AS", without the slash (NEP V calls)
   - "Petersen & Partners Private Equity II A/S"
   - "Petersen & Partners FoF K/S"
   - "Petersen & Partners FoF KS" (in file names)

## Per manager

### Newbury (NEP5, NEP6) — manager `NEP`

**Capital account.** Two numeric columns, whose headers wrap over two lines.
- **NEP V:** the columns are "Petersen & Partners" (our LP: commitment 123,800,000) and
  "Total Partnership" (the whole Ontario feeder: 163,800,000). Read the **first** column.
  - ⚠ The existing Claude skill `nep-capital-account-update-v2` says to read "Total". That is only
    harmless for NEP VI, where both columns are identical.
- **NEP VI:** "Petersen & Partners Private Equity II A/S" and "Total Partnership" are identical.
- Lines:
  - **ITD:** Partnership commitments, capital contributions since inception, unfunded, distributions since inception.
  - **YTD (Jan 1 → period end):** Partners' capital Jan 1, contributions, net increase/(decrease) from
    operations (**not split** into income/gains/fees), distributions, partners' capital end (gross),
    current-year adjustment to carried interest – unrealized, partners' capital net of carried interest.
- Roll-forward holds: 100,983,610 + 0 − 3,166,446 − 999,895 = 96,817,269; + 4,089,155 = 100,906,424.
- ⚠ **NEP V: the Jan 1, 2026 opening (100,983,610) ≠ the Dec 31, 2025 closing in the Q4 statement**
  (101,463,036 net / 101,991,739 gross). The Q4 statement is unaudited, so year-end was probably
  restated after audit. The "opening = prior closing" validation should be a **WARN**, and the difference
  shown as a restatement.

**Capital call.**
- The letter gives: notice date, % called, commitment, **funding requirement $X on or before DATE**,
  % contributed, and remaining unfunded.
- There is no component split, so the whole amount maps to *capital call – not split*.
- From January 2026, NEP VI notices come from a new admin platform ("Bridge Partner ID"). Same fields,
  different layout; tokens break across lines.

**Financial statements (quarterly).**
- Ontario feeder FS (balance sheet, operations, changes in partners' capital, cash flows).
- LP report (narrative, distributions received).
- Schedule of investments per underlying fund.
- Master fund FS.

Fund-level data only → Claude, later.

### Park Square (PSCP4, PSCP5) — manager `PSCP`

**Partner's capital statement** — the ILPA reporting template:
- 9 numeric columns: LP's allocation / Total Fund / FLP's allocation × Quarter to / Year to / Inception to.
  **Ours = columns 1–3.** A viewer watermark ("Allan Nielsen") appears in the header rows only.
- Fixed ILPA labels:
  - Beginning NAV, contributions, distributions
  - management fees (gross / waivers / offsets / net)
  - ~25 expense lines with subtotals
  - interest and dividend income
  - total net operating income
  - net realised and unrealised gain
  - Ending NAV – net of carry
  - carried interest roll-forward
  - Ending NAV – gross of carry
- **The NAV is already net of carry.** The carry lines are an accrued *balance* (memo), not a
  roll-forward movement. Roll-forward: 17,918,793 − 136,666 − 26,770 + 334,062 = 18,089,419 ✓.
- A.2 commitment reconciliation: total commitment, beginning unfunded, contributions,
  **recallable distributions**, ending unfunded.
- Page 4: **complete dated cash-flow history since 2021** (contributions, true-up interest,
  distributions). Page 5: GP's DPI / RVPI / TVPI / net IRR.
- ITD contributions (22.7M) exceed the commitment (14.8M) because of recycling (recallable distributions).

**Drawdown notice.**
- Issue date, commitment, aggregate fund call, **your Capital Contribution €X**, drawdown date,
  remaining commitment.
- "Partially deemed" variant: gross contribution €522,476.67, *netted against Distribution No. 21*,
  net payment €415,615.21. This is stored as two cash-flow rows (gross call; distribution), whose sum is
  the payment.
- The component breakdown ("Schedule I") is **not included** in the PDFs, so calls map to
  *capital call – not split*.

**Financial statements** (60+ pages, consolidated accounts) → Claude, later.

**Quarter-end equalisation accruals (PSCP V, booked 2026-10-08 on the user's instruction).**
The statements' "Cash flow schedule and net IRR calculation table" counts the equalisation of
a subsequent closing as returned at the quarter end and drawn again the next day, before the
cash settles with a later notice. No notice exists for these rows, so they are booked from
four **placeholder documents** (blank page + MANUAL extraction, doc_type OTHER, whole EUR as
printed): 2024-03-31 `CF150` +69,639 and `CF240` +2,522; 2024-04-01 `CF100` −69,639 and
`CF130` −2,522; 2024-12-31 `CF150` +1,630,473; 2025-01-01 `CF100` −1,630,473. Each pair nets
to zero (paid-in unchanged, 11,726,857.52 at Q2 2026); only the IRR at those two quarter
ends changes (Q1 2024 −23.81% → −11.17%, Q4 2024 −19.07% → 5.42%). The statements print no
IRR for PSCP V, so there is no GP figure to compare with. Look for new pairs in the schedule
when a statement arrives after a subsequent closing.

### ATP PEP (ATPPEP9) — manager `ATPPEP`

**Capital account** (Danish format, **EUR thousands**):
- **Old layout (Q4 2025, Q1 2026):** a wide table, one row per investor, for current quarter and since
  inception. Columns: capital account start, paid in, distributions, portfolio funds P/L, mgmt fee,
  operating expenses, financial items, capital account end, carry accrued, capital account net.
  Commitment table on page 2.
- **New layout (Q2 2026):** vertical, with commitments first, then the capital account with
  Quarterly / YTD / Since inception columns.
- Recallable distributions are shown in the commitment table.

**Net capital call** (whole EUR, Danish format). Two columns: fund and our share (use the second).
- Components: capital paid-in to the portfolio, Danske Bank fees and interest, audit fee, expenses,
  management fee, excess cash.
- Optional distribution section (distributions from portfolio), then total net capital call.
- Then commitment, paid-in including this notice, remaining commitment, recallable, **value date**.
- Appendix 1 lists the underlying funds.

**Quarterly internal report / annual report (årsrapport)** — portfolio companies by FMV, notes,
key figures → Claude, later.

### Navigare (MIF3) — manager `NAVIGARE`

**Investor statement** (USD):
- 9 numeric columns: our LP / Navigare Partners Invest III K/S / Total Fund × current quarter / YTD /
  since inception. Ours = columns 1, 4, 7.
- Lines:
  - opening balance, capital contributions, capital distributions, net contributions
  - net movement in unrealised gain, realised gain
  - **"Distribution from investments"** (a P&L add-back equal to the distributions)
  - management fees, other admin expenses, financial income/(expenses)
  - NAV, unrealised carried interest, **NAV after carried interest** (the reported NAV)
- Page 2:
  - commitment and ownership %
  - dated list of calls, equalisations, recallable distributions and distributions
  - remaining commitment, net IRR, drawn bridge facility
- Our LP entered via a late closing ("Equalization 6", 24 Sep 2025) → equalisation flows.

**Deposit notice.**
- Notice date, value date, total amount.
- Specification with fund total and "Your share": investment cost (per project), management fee,
  other partnership expenses.
- Commitment, contributed before and after, undrawn after, recallable capital.

**Quarterly report** (42–45 page presentation) → Claude, later.

### Capidea (CAP3, CAP4) — manager `CAPIDEA`, fund-level only (added 2026-09-30)

None of our investors is in these funds: the reports live in `documents/<fund>/inbox/` (no
investor level; V006 `fund.document.fund_id`) and go to schema `portfolio` only. DKK.
CAP3 = Capidea Kapital III K/S (vintage 2018, DKK 844m, reports Q4 2017 – Q1 2026);
CAP4 = Capidea Kapital IV K/S (vintage 2023, DKK 1,251.5m, Q2 2023 – Q1 2026).

**Quarterly report** (15–78 pages): "Fund Performance Summary" + a portfolio companies table
(name, industry, entry date, valuation/exit date, invested / realised / unrealised / total in
DKKm, gross IRR, multiple; Total row over realised + current). Four template variants:
- `fund_report` (Q4 2025 on): summary in whole DKK, columns quarter / YTD / since inception /
  prior year-end → the third number. Distributions printed negative.
- `fund_report_dkkm` (to Q3 2025): DKKm, one column per year, "Acc." last → flows and NAV =
  last number, balances (investments, cash, carry) = first column, because the "Acc."
  column is stale in CAP4 Q3 2024 and CAP3 Q2 2024. Ratios in "OTHER INFORMATION".
- `fund_report_dkkm_da` (CAP3 Q1 2018 – Q2 2019, Q1 2020): Danish decimals ("736,3"), ratios
  per year with "Acc." last; Q1 2020's portfolio table has decimal points.
- `fund_report_dkkm_ratios` (CAP3 Q2 2023): the summary table is not in the text layer → only
  unfunded, IRRs, ratios and holdings.

Name and industry cannot be told apart in the portfolio table, so the industries are listed
in every profile: a new company with a new industry fails as "row not understood" → add it.
No portfolio table (key figures only): CAP3 2018 – Q2 2019 (none printed), Q1 2021 and Q2 2022
(not in the text layer). Not captured: fund-level drawdown lists, P&L lines, other LPs' capital
accounts. Approved with `--override`: CAP3 Q2 2024 (GP's Maki Toys row: 186.0 ≠ 34.8 + 148.1).
Rejected: CAP3 Q4 2017 (a welcome letter, nothing drawn; MANUAL extraction without records).

### CataCap (CATACAP3) — manager `CATACAP`, investor MOTORTRAMP (added 2026-09-30)

CataCap III K/S, DKK, vintage 2022 (LPA 30 Jun 2022). A/S Motortramp: commitment 37,500,000
(derived from Equalization Drawdown 1: 36,372,945 remaining after 1,127,055; the 17,574
"additional amount" does not reduce the commitment), first close 21 Aug 2023.

**Draw down notice** (`capital_call`, pypdf). Letter: number ("EQDD1/2023", "8/2026"), dated,
drawdown date, amount, remaining commitment. Page 3: one row per LP with sections (c) before,
(d) this notice, (e) after, each closed by a "%" column. The component columns differ per
notice and the number format alternates (1,234 / 1.234) → strategy `group_table`: names from
the header, values from our row, both split at the "%" columns; format from the row's %.
Mapping: management fee and its payroll taxes → CF110; fund expenses, establishment, broken
deal costs → CF120; investments → CF100; "additional amount" (equalisation) → CF130; memo
commitment MM100, remaining commitment MM120. The (e) group of DD5 omits two columns (GP).
DD6: page 3 is an image → `extracted/…DD6….json` (CLAUDE): CF199 277,061 from the letter;
the derived split (DD7 before − DD5 after) is in its note.

**Distribution notice** (`distribution`): fully recallable ("all recyclable, LPA 8.3.1") →
CF230. A distribution that is not (all) recallable matches no profile yet.

**Quarterly report** (37–45 pages): per-LP capital accounts in the appendix (pdfplumber; pypdf
puts the names on lines of their own from 2024): since inception, quarter, year to date.
`capital_statement` (7 columns, no carry: 2023 Q2 – 2025 Q2) and `capital_statement_carry`
(9 columns: + before carry CA890, unrealised carried interest CA400). Booked: YTD (the QTD
opening is the 1 January balance in Q2 2026, so QTD does not roll forward) + memo from the
since-inception row (commitment, contributions, outstanding, recallable, TVPI).
"CC III 2025 Q2 - Revised Capital Accounts" supersedes the original Q2 2025 (the revised
adds carry): load the original first. No reports for Q3 2023 – Q1 2024 (NAV gap).
Rejected: 5 ESG reports, the AGM notice, the 2024 annual report (MANUAL records without
figures); its identical copy is in processed/duplicates.

### Seed Capital (SEED4) — manager `SEED`, investor MOTORTRAMP (added 2026-09-30)

Seed Capital Denmark IV K/S, DKK, vintage 2020 (final closing Sep 2021). A/S Motortramp came in
at the 2nd closing (letter 15 Jul 2021): commitment 30,000,000. The annexes print the LP as
"Motortramp" (alias added).

**Draw down letter** (`drawdown`): "Amount: DKK 2,020,154", "Value date: 28 June 2023" → CF199.
**Annex 1 & 2** (`drawdown_annex` / `_en`, detail documents like PSCP's Schedule I): our row =
commitment, %, contributed to date, unpaid, amount less management fee (CF100), management fee
(CF110), total, contributed after, unpaid after; replaces the letter's CF199 on the same value
date when the totals agree. Number format per annex (DD9: 1,234; the rest: 1.234).
**Distribution letter** (`distribution`): exit proceeds, no split → CF295. Its annex (fund total
by company) is rejected: nothing beyond the letter.

**Quarterly report** (`capital_statement` / `_carry`, pdfplumber: pypdf puts the headings after
the rows): per-partner capital accounts for the quarter and since inception. Booked: since
inception (ITD, opening 0) → CA900; the quarter table is before carry and the quarter's carry
movement is not printed. From Q4 2024 "CAPITAL ACCOUNT END PERIOD: <date>" ("Adjusted for" /
"Deducting unrealised carried interest"): before carry CA890, carry balance CA400 (= the ITD
movement), net CA900. Number format per report (`number_format: auto`, on the capital-account
pages). Period from "CURRENT REPORTING PERIOD: Qn yyyy (…)"; the Q1 2023 report's since-inception
heading says "31 MARCH 2022" (GP typo). The file "Quarterly Report Q4 2025" is the Q1 2026 report.

Missing: draw downs 14 (≈943,382, Apr–Sep 2024) and 16 (≈2,172,990, Sep 2024–Nov 2025), implied by
the annexes' "contributed to date" — until loaded, SEED4 paid-in / TVPI / IRR from cash flows
are overstated; reports for 2023 Q4, 2024 Q2 and 2025 Q4. Rejected: the distribution annex, the
2024 annual report and its specifications.

### Alternative Equity Partners (KAPE22) — manager `AEP`, investor MOTORTRAMP (added 2026-10-05)

K Alternativ Private Equity 2022 K/S ("KAPE 2022"), DKK, a fund of funds (11 underlying funds,
8 managers; AlpInvest manages the portfolio), fund size 841.4m (522.9m until the last closing,
31 Aug 2022). A/S Motortramp: commitment 14,000,000 = 1.66%, closing date 2021-09-30 (user).
The reports name investors by ID only: ours is **A1615** ("Investor ID (AEP)" on the notices,
next to the name), added as an alias of MOTORTRAMP.

**Kapitalkald** (`capital_call`, pypdf, Danish numbers; also "Kapitalkald nr. 4 samt udlodning
nr. 1"). Page 2, Sektion C has our share per component; the rows wrap over 2–3 text lines, so
they are read from the collapsed text (`items.region`). Mapping: Investering i … → CF100;
Forvaltningsgebyr → CF110; Administrationsomkostninger, AlpInvest øvrige omkostninger,
Tegningsprovision → CF120; Arbejdskapital → CF190, and "(Returnering af) tidligere kaldt
arbejdskapital" → CF190 negative (user decision 2026-10-05: it increases the resttilsagn, so it
is a negative call, not a distribution); Udlodning fra … → CF295 (no effect on the
resttilsagn: not recallable). Check: components = "Samlet beløb til indbetaling" (the wire,
net of the distribution). Memo: tilsagn MM100, kald inkl. this notice MM110, resttilsagn MM120.

**Kvartalsrapport** (14–17 pages) is two things at once, and one document carries both
(`doc_type` CAPITAL_STATEMENT with fund-level `records`):
- *Investorkapitalkonti*: one row per investor ID, since inception (ITD, opening 0): kaldt
  kapital CA110, udloddet CA120, andel af gevinst/(tab) CA290, tegningsprovision CA220,
  forvaltningsgebyr + skyldig forvaltningsgebyr CA210, kapitalkonto før hensat succeshonorar
  CA890, hensat succeshonorar CA400, kapitalkonto efter CA900; memo MM100/110/120, TVPI MM160,
  IRR MM150. The row is read with a regex (`fields`), because zero cells are sometimes left
  out (Q2/Q3 2024); checks: roll-forward (tolerance 3) and før + hensat = efter.
  Six layouts, one profile each: `capital_statement` (from Q1 2026), `_2024` (Q2 2024 – Q4 2025,
  same row, older fund table), `_2023q4` ("Transfer regulering"), `_2023` (Q4 2022 – Q3 2023),
  `_2022q3`, `_2022h1`. Before the first call (Apr 2023) the account is slightly negative or
  an unrealised gain on nothing paid in (−18,972 … 209,441): loaded as printed.
- *Fund level* → schema `portfolio`: key figures from "Fondsopsummering" (mio. DKK, one
  decimal: paid-in, distributed, NAV; net IRR, DPI, RVPI, TVPI; total commitment) and
  "Porteføljeafkast", one holding per underlying fund (section FUND: tilsagn, kaldt, kaldt %,
  udloddet, DPI, aktuel værdi, værdiforøgelse, TVPI, IRR), tied to "Samlet porteføljeafkast".
  From Q1 2026 the table has no called amount (only %) and shorter fund names ("Advent GPE X"
  for "Advent International GPE X"); the net DKK TVPI / IRR are kept. Repairs
  (`holdings.substitute`): the rating bullet "⬤", an empty ΔTVPI cell (Q1 2026), and negative
  values of funds not yet called whose minus sign is detached (Q1–Q2 2023). The Q1 2022
  totals row has one token less: no tie-out for that report. Not captured: sector / geography
  / top-10 shares (chart labels, scrambled in the text), the fund's P&L and balance sheet,
  the other investors' rows.

pdfplumber splits the digits in these PDFs ("1 4.000.000"), so everything is read with pypdf.

Missing: the report for Q1 2024, and the notices for Kapitalkald nr. 5 and 6 — the user does
not have them. On the user's instruction (2026-10-05) the two calls are registered through
**placeholder documents** ("Kapitalkald nr. N - PLACEHOLDER (no notice received).pdf": a blank
page + a MANUAL extraction, note in `fields.note`): nr. 5 = 754,330.00 (Samlet kaldt kapital
Q3 → Q4 2024) and nr. 6 = 1,374,211.34 (5,635,206.32 before call 7 less what was called
before), both CF199, value dates **estimated** as the quarter ends 2024-12-31 and 2025-06-30.
Paid-in then equals the GP's (6,879,799); IRR is slightly off (17.9% against the GP's 17.6%).
If a real notice arrives, reject its placeholder: the value date differs, so it is not
superseded automatically.
Loaded 2026-10-05: 5 notices (the second copy of Kapitalkald 1 is in processed/duplicates) and
17 reports, Q1 2022 – Q2 2026. Rejected: the 2024 annual report (MANUAL record without
figures; no investor figures in it).

#### KAPE21 — K Alternativ Private Equity 2021 K/S (added 2026-10-05)

Same manager, templates and investor ID (A1615). DKK, fund size 558.4m, first closing
5 Oct 2020, 7 underlying funds. A/S Motortramp: commitment 25,000,000 = 4.48%, closing date
2020-10-05 (the fund's first closing, as for KAPE22). What differs from KAPE22:

- **Kapitalkald 4** (Sep 2021) is a letter (`capital_call_letter`, English numbers, Danish
  dates → `parse_date` knows Danish month names): our share per component on page 1,
  commitment and paid-in before / after on page 3.
- **Kapitalkald 5** (Oct 2021): Sektion C without the K/S column, "Tranaktion" in the header,
  and "Rebalancering i relation til nye investorer" (−274,842.40, increases the resttilsagn)
  → CF150 (contribution refund, equalisation). The call rows are matched on the transaction
  type ("Kald: Investeringer", "Kald: Arbejdskapital - Tidligere kaldt" …), not on the
  description, which varies ("Tilbageførsel af …", "Returnering af …").
- **Udlodning nr. 3** (Aug 2026, 2,652,668.34): a distribution without a call
  (`distribution`, CF295; total "Samlet beløb til udlodning"). Udlodning 1 and 2 are on
  Kapitalkald 11 and 13.
- **Kvartalsrapport 2022 – 2023**: "rådgivningsgebyr" columns and a succeshonorar column from
  the start (`capital_statement_adv`, `_adv_2023q4`); "Udloddet" is printed positive there and
  deducted (sign −1). From Q2 2024 the layouts are KAPE22's.
- **Fund totals** (paid-in, distributed, NAV, TVPI, IRR) in the 2024+ profiles come from the
  "Total" row of Investorkapitalkonti (whole DKK, the sum of the investors' accounts), not
  from "Fondsopsummering": that page's text is scrambled in KAPE21 Q2 2026 (the first load
  read NAV 21.4m for 688.4m; corrected the same day with `docs reextract`, profiles v2, for
  the 18 reports of both funds). The 2022–2023 profiles still read the summary (mio. DKK).

Not received: the notices for Kapitalkald 1–3 and 14, and the report for Q1 2024. As for
KAPE22 they are booked from **placeholders** (CF199, estimated value dates): 773,463.00 at
2021-02-28 and 160,064.97 at 2021-03-31 (paid-in per the advisor's reports of those dates;
933,527.97 before Kapitalkald 4 per its letter), and Kapitalkald 14 = 1,542,726.99 at
2024-12-31 (22,388,657.53 on Udlodning 3 less 20,845,930.54 after Kapitalkald 13).
Rejected: the 2024 annual report and the advisor's (Spektrum) portfolio reports of Feb / Mar
2021 — not fund documents; they also show a K Alternativ Private Equity 2020 K/S
(25,000,000), not loaded. Loaded: 12 notices + 3 placeholders, 17 reports Q1 2022 – Q2 2026;
the second copy of Kapitalkald 11 is in processed/duplicates.

#### KAPE20 — K Alternativ Private Equity 2020 K/S (added 2026-10-05)

Same manager, templates and investor ID (A1615). DKK, fund size 389.0m, first closing
26 Feb 2020, last 5 Oct 2020, 7 underlying funds. A/S Motortramp: commitment 25,000,000 =
6.43%, closing date 2020-02-26 (the fund's first closing, as for the other two). The 2021
letters are from KAPE FAIF ApS (Spektrum), the manager at the time. What is new here:

- **Kapitalkald 4–6** are letters (`capital_call_letter` v2): nr. 5 says "andel af Investering
  i … er:" and "modtaget senest den"; rows "Stiftelsesomkostninger" and "Justering af
  tidligere kaldte administrationsomkostninger" (negative) → CF120; "Regulering af tidligere
  kaldt kapital" (nr. 6, −46,446.12) → CF150, like the "Rebalancering" on KAPE21 nr. 5.
- **Kapitalkald 8** (Jan 2022): the investment rows carry no transaction type, only the
  underlying fund's name → CF100 by a rule for untyped rows naming a fund.
- **Distribution rows** named by the fund only ("Charlesbank Equity Fund X §8 Udlodning") →
  CF295 via the type at the end of the label.
- **Kapitalkald 16 & udlodning 5**: the distribution exceeds the call, so the total is
  "Samlet beløb til udlodning" (negative, a net receipt of 1,478,149.10).
- **Kvartalsrapport 2022 – Q3 2023**: the `capital_statement_adv` columns (the profile is
  recognised by "Kaldt rådgivnings-"), and an empty ΔIRR cell in the Q1 2022 fund table.
- The fund's DPI is only on "Fondsopsummering"; it is taken only when the previous quarter's
  value follows it, so a scrambled page (Q2 2026) leaves it out instead of reading the change
  column (first load: 0.04; corrected with `docs reextract`, profiles v3).

Not received: the notices for Kapitalkald 1–3 (110,321.42 in total per the letter for nr. 4;
booked from one placeholder, CF199, value date estimated 2021-02-28) and the report for
Q1 2024; no reports before Q1 2022. Every later call and all six distributions are there:
paid-in 21,529,567 and distributions 6,371,274 tie to the GP's figures. Rejected: the 2024
annual report and three Secure Spectrum newsletters (commentary only). Loaded: 16 notices +
1 placeholder, 17 reports Q1 2022 – Q2 2026; the second copy of Kapitalkald 11 is in
processed/duplicates.

## Distribution notices (added 2026-09-29, 15 files)

All are **Profile** (regex on the notice text).

- **NEP V** (letter)
  - Fields: total proceeds, Ontario feeder share, **"Your Share: $X"**, tax withheld, **net wire amount**,
    payment date, and a *tax* estimate of the split (% return of capital / gain / interest / dividends).
  - That % split is an estimate for tax purposes, so it is stored only as a note, and the amount is
    booked to `CF295` (not split). Withholding tax, if any, goes to `CF280`.
- **PSCP** (text only readable with **pypdf**: pdfplumber drops every line containing "€")
  - *Distribution Notice:* amount, value date ("on or around"), the **redrawable** portion (goes to
    `CF230`; the rest to `CF295`), and contributions outstanding and remaining commitment afterwards.
  - *Deemed Distribution Notice:* same fields, "no payment is due; netted against Drawdown No. X".
    It is booked from its own notice, with `flow_ref = 'PSCP4 Distribution 21'`. The netting drawdown
    notice then books only its gross contribution.
  - *True-Up Notice – Return of Capital* (subsequent closings): **Capital Contribution 5.2(a)** refunded
    (goes to `CF150`, which reduces paid-in) plus **Additional Payment 10.2(b)** (equalisation interest,
    goes to `CF240`); net amount and value date.
- **MIF** (same layout as the deposit notices; specification with "Your share")
  - *Distribution Notice*: e.g. "Operational Profit from investments" (non-recallable) goes to `CF220`.
  - *Net Distribution Notice*: a recallable re-financing distribution (goes to `CF230`) **netted against a
    deposit notice** (investment cost `CF100`, management fee `CF110`, other expenses `CF120`). One
    document yields both inflows and outflows, and the deposit-notice lines carry `flow_ref = 'MIF3 DN15'`.

**General rule:** a netted notice produces one cash-flow row per component, with gross amounts. The
sum of the rows equals the net payment, which is validated. `flow_ref` stops the same component being
loaded from two documents.

## Consequences for the data model (applied to V001)

- Accounts added:
  - `CF150` contribution refund at a true-up (PSCP V; reduces paid-in, not a distribution)
  - `CA290` net result from operations, not split (NEP)
  - `CA230` financial items / financing costs (ATP, MIF, PSCP facility interest)
  - `CF199` capital call, not split (NEP, PSCP)
  - `CF295` distribution, not split
  - `MM200` accrued carried-interest balance (PSCP, ATP; memo, not a roll-forward movement)
  - `MM150` fund-reported net IRR
  - `MM160` fund-reported TVPI (GP figures kept for cross-checking our calculations)
- `CA400` is now explicitly the carry **movement** that bridges gross and net NAV (NEP, MIF).
  `CA890` holds the NAV before carry adjustment.
- `document.external_ref` holds the notice number ("Drawdown Notice 27", "DN16", "Call 31").
- `cash_flow.flow_ref` holds the reference of the flow itself. It is used for de-duplication, e.g. to
  stop "Distribution No. 21" being loaded both from a deemed drawdown and from its own distribution notice.

## Decisions (user, 2026-09-29)

1. Distribution notices were added where available (NEP V, PSCP IV/V, MIF III; none for NEP VI or ATP).
2. **Cash flows come from notices.** The statement cash-flow lists are used for backfill and reconciliation only.
3. The NEP V Dec 2025 → Jan 2026 difference (479,426) is **not known** to the user. It stays a WARN on
   load and should be raised with Newbury.
4. MIF "Distribution from investments" maps to `CA300` realised gain.
5. The `nep-capital-account-update-v2` skill is to read the investor's own column. A corrected SKILL.md was prepared; it must be uploaded to the claude.ai plugin, because the local copy is synced.

## Findings from the first end-to-end dry run (2026-09-29)

All 52 LP-level documents were staged, validated and approved in one rolled-back transaction.
The "opening = prior year-end closing" rule found **year-end restatements** in three funds. The GPs
changed the Dec 31, 2025 capital after the Q4 statement was issued:

| Fund | Dec 31, 2025 closing (Q4 statement) | Jan 1, 2026 opening (2026 statements) | Difference |
|---|---:|---:|---:|
| NEP5 | 101,463,036 | 100,983,610 | −479,426 |
| NEP6 | 23,899,320 | 23,984,692 | +85,372 |
| MIF3 | 3,162,149 | 3,194,497 | +32,348 |

These are WARNs, not ERRORs. They are worth confirming with the GPs (audited year-end figures).

## Full history batch (2026-09-29, 280 new files)

All LP-level documents extract with every check passing, except as noted:

| Manager | New document types / variants handled |
|---|---|
| PSCP | **Schedule I** (`schedule.yaml`): the component split behind the notices due on a date (incl. combined "Deemed Distribution No. 21 and Drawdown No. 27"). Direction comes from the "impact on remaining commitment" column (Decreases = we pay; Increases / No impact = we receive), since one schedule prints the net column with the wrong sign. The 2016 ILPA statements 2021-23 ("capital statement for <investor>"). Old notice headers (TO/FAO/DATE). True-ups where **we pay** (`true_up_payment.yaml`; 2021 "Drawdown 1 and 2" and 2025 "True Up Notice 11/12"), interest-only true-ups, "Distributable Cash" wording, and "required to make to make". |
| NEP | 2020 inception-year statements (singular labels, no carry → closing = NAV), statements whose pypdf text separates numbers from labels (re-read with word positions), NEP VI on the **Bridge platform** (`partner_package.yaml`, `partner_statement.yaml`), capital-call letters with split words (pdfplumber first) and a commitment **increase** ("additional capital commitment"). |
| NAVIGARE | **Equalization notices** (`equalization.yaml`): the specification page is letter-spaced ("E q u a liza tio n"), so those fields use `nospace`. |

Scanned PDFs (no text layer) were read by Claude and saved as `extracted/*.json` with
`method: CLAUDE` (used by `pefund load` when no profile matches and the sha256 is the same):
PSCP IV Deemed Drawdown 9, Deemed Distribution 4 and their Schedule I (2022-06-29), and the
PSCP V statement for 30 Sep 2024 (QTD and ITD only; YTD left out because one figure could not
be read with certainty).

Not loaded, deliberately:
- Fund-level reports before Q2 2026 stay in `inbox` (see "Fund reports" below).
- `MIF3 260309-DN2`: one page, "Cancellation Deposit Notice 2" (a guarantee released; no cash).

Learned rules:
- Notice numbers are not unique (true-ups reuse "Drawdown Notice No. 1"): a notice is identified by
  number **and** value date, and flow references include the value date.
- A Schedule I replaces the cash flows of all notices of the investment due on the same date when
  the totals agree; equalisation interest ("Additional Payment", CF130/CF240) is not in a Schedule I
  and is kept from the notice.
- Commitments change over time (PSCP IV €8.1M → €14.8M in 2021; PSCP V €19M → €32.5M → €36.3M;
  NEP V increased in 2021). The notices carry the commitment at each date (MM100 memo lines).

## Fund reports (added 2026-09-29): schema `portfolio`, full history

The user wants fund-level data, kept apart from the investor level, with full history (decided 2026-09-29).
There is one `fund_report.yaml` per manager. For Q2 2026 each fund loads with every check passing:

| Manager | Document | Holdings | Key figures | Breakdowns |
|---|---|---|---|---|
| NEP | Ontario feeder package: "Report to the Limited Partners" and the Secondary / Co-Investment / Primary Portfolio Summary ($000). pypdf, with pdfplumber fallback | per section, with a totals-row check (NEP5 111 rows, NEP6 ~25) | NAV (rounded, $m), number of funds and companies, underlying unfunded, called %, DPI %, credit line drawn (master fund balance sheet) | – |
| PSCP | "Financial Report and Unaudited Accounts": Track record (€m, one row per tranche; Realised/Unrealised) and Portfolio overview | tranche rows (instrument, currency, trade/exit month), totals check | 17: invested/realised/unrealised, gross/net/asset IRR, TVPI/DPI, W.A. EBITDA/EV/equity cushion/yield... | top 10 issuers, industry, geography, currency |
| ATPPEP | "Quarterly Internal Report" (EUR thousands, Danish numbers) | underlying funds: segment, vintage, **local-currency** commitment, PICC, DPI, TVPI, IRR (n/m → NULL) | fund overview: commitments, paid in, distributed, recallable, unfunded, portfolio FV, NAV, net IRR, DPI, TVPI, PICC | geography, segment (share of invested capital) |
| NAVIGARE | "Unaudited quarterly report" deck (USD) | Investment portfolio summary: one row per SPV (cost, distributions, FV, gains, TVPI, DPI), totals check | Key figures page, "Balance Qn" column (middle number), IRR and multiples, credit facility drawn | – (allocation charts are images) |

Known gaps (older reports, not needed now):
- NEP5 Dec 2025: 6 secondary-total checks fail.
- NEP6 Sep 2025 FINAL and the "QUARTERLY FUND PACKAGE" format match no profile.
- PSCP Q1 2026: the breakdown charts are images, so they are recorded as notes.
- MIF Q4 2025 and the ATP annual report ("årsrapport") match no profile.

Text-layer quirks handled:
- pypdf splits numbers before a comma ("10 ,244"), which the holdings strategy repairs.
- There is a zero-width space in "€146.5​m"; invisible characters are now stripped on all pages.
- "0.00x" appears in a Danish table; ratios fall back to the other decimal mark.
- pdfplumber splits words ("C SOV", "C ontainer"); `segment` values are re-joined.
- "Q2 2026" / "2026 Q2" parse as the quarter's last day.

### Full history load (2026-09-29)

79 reports are approved: NEP5 2020-06 to 2026-06 (24 dates), NEP6 2024-03 onward, PSCP4 2019-12 onward, PSCP5 2023-12 onward, MIF3 2025-06 onward, and ATPPEP9 2025-12 onward. Drafts are SUPERSEDED where a final report exists for the same quarter.

Four reports are staged but not approved (status EXTRACTED):
- NEP5 Mar 2022 (Ontario and L.P. copies): the secondary-portfolio unfunded total is $97k off the sum of the rows. All rows are captured and every other column matches, so the report's own table doesn't add up.
- PSCP4 Sep 2019: the track-record table has no text layer.
- PSCP4 Dec 2025 (audited): the realised total includes LGC Pref (€34.6m) and another row, but neither is in any text layer.

Not loaded: NEP6 Sep 2025 drafts (broken font encoding in the tables; the FINAL of that quarter is loaded) and the ATP 2025 annual report (no profile).

Variants handled for the history:
- NEP:
  - a name wrapped onto its own line, whose numbers-only line is not a totals row
  - the "QUARTERLY FUND PACKAGE" layout, with a pdfplumber retry
  - 2020 reports: the date comes from "Overview As of", and the primary table is optional
  - split numbers "10,7 65", "$ 8 1,706" and "$ 3 89"
- PSCP:
  - track records up to 2024 without instrument and currency
  - "Feb -20" and "(CAD)Mar-22"
  - FY 2022 "Realized/Unrealized"
  - "31 Match 2025" (the date is taken from the file name)
  - key figures optional
- Maritime: the Q4 deck is an "Audited quarterly report"; key figures optional.

### Portfolio company key statistics (PSCP, added 2026-09-30)

From ~page 20 on, each PSCP report has a "Portfolio company overview": per company a
heading, a description and a "Key statistics" box (Revenue / EBITDA / EV / Net Debt with
LTM or FY period, Headquarters, Sponsor, Tranche, Invested Amount), three per page. The
`companies` strategy stores one row per company and report in `portfolio.company_statistic`
(V008). A dry run over the 38 approved PSCP reports gives 4–30 companies per report, about
755 rows, in line with the unrealised holdings of the same report.

Layout variants handled:
- up to 2021, the box also has Total Value / Total MM / Entry price / Margin (or Rate / Fair
  Value at) per tranche, with tranches separated by spaces: those rows are kept in `raw_text` only
- PSCP IV FY 2024: every character has a slightly skewed text matrix, so pdfplumber's
  `extract_words` splits words into letters; words are built from the characters instead
- PSCP IV Q1 2024 p.33: a second slide is stacked over the first, and white panels hide
  the lower one. Characters under a later fill at least 40% of the page wide are dropped
  (logo frames next to the boxes are narrower and overlap real figures)
- "Key statistics" / "Key Statistics"; "Net Sr Debt" instead of "Net Debt" in PSCP IV
  Q4 2025 and PSCP V Q1 2026 (own column `net_senior_leverage`)
- empty template boxes ("#N/A", or VPP after its exit: "Revenue (LTM Jan-00) m") are skipped
  with a warning
- no currency on a revenue figure (Oterra, Petvet, Veonet, Thyssenkrupp in 2023): taken
  from the box's other figures, with a warning. "n.a." values stay NULL (Dechra net debt,
  Flow Control revenue in 2023)

Gap: PSCP V Q2 2025 — the company pages are one image each (no text layer), so no rows.
Company names are stored as printed; they vary (DiversiTech / Diversitech, PetVet / Petvet,
IntraFi / IntraFi Network, Thyssenkrupp / Thyssenkrupp Elevators / TKE) and GPs rename
(Pharmazell → Axplora, Hansen Colors → Oterra, Bosch Packaging → Syntegon, BASF → MBCC Group).
`portfolio.company_alias` (V009, confirmed by the user 2026-09-30) maps them to one name, the
view's `company` column: 58 printed names → 51 companies, one row per company and report.

Loaded 2026-09-30 with `docs reextract` (all 38 reports; PSCP4 Dec 2025 again with
`--override`, same reason as before): PSCP4 599 rows over 27 reports, PSCP5 156 over 10.

## Left in the inbox on purpose (2026-09-30)

| File | Why |
|---|---|
| PSCP4 `inbox/2019/… 30-09-2019.pdf` | the track-record table has no text layer; the series starts Dec 2019 |
| NEP6 `… September 30, 2025 - DRAFT.pdf`, `Ontario VI report 20250930 -DRAFT.pdf` | broken font encoding in the tables; the FINAL of Sep 2025 is loaded |
| MIF3 `260309-DN2- Petersen & Partners .pdf` | "Cancellation Deposit Notice 2": a guarantee released, no cash |
| ATPPEP9 `2025 - årsrapport - PEP IX B KS.pdf` | no profile yet (Q4 2025 comes from the quarterly report) |
| SEED4 — (not received) | draw downs 14 and 16: asked the user for the letters (2026-09-30) |
| KAPE22 — (not received) | quarterly report Q1 2024; Kapitalkald nr. 5 and 6 are booked from placeholders (user, 2026-10-05) |
| KAPE21 — (not received) | quarterly report Q1 2024 and the reports before Q1 2022; Kapitalkald nr. 1–3 and 14 are booked from placeholders |
| KAPE20 — (not received) | quarterly report Q1 2024 and the reports before Q1 2022; Kapitalkald nr. 1–3 are booked from a placeholder |
| NEP5 `12.31.21 … Capital Account.pdf`, NEP5/NEP6 `12.31.24 … Feeder FS and NAV.pdf` | investor-level, not reviewed yet; the NAV for those dates is already loaded from other statements |

Approved with `--override` (failed totals, reason in `document.note`):
- NEP5 Mar 2022 (both copies): the GP's secondary unfunded total is $97k off.
- PSCP4 Dec 2025: realised rows are missing from the PDF text (€45.9m invested, incl. LGC Pref).

NEP VI Bridge packages from Q2 2024 on have a 2-column Partner Statement
(`partner_package_ytd.yaml`). Docs #497/#509/#512/#515 were re-read with `docs reextract`.
