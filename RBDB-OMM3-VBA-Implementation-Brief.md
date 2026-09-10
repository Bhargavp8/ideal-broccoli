# RBDB Email Automation — OMM3 Migration Implementation Brief

**For:** Claude Code, working on `modAnomalyAlert` (v22) in the RBDB Access database
**From:** Power BI side of the same project
**Date:** 10 September 2026

---

## 0. Read this first

You have the VBA repo. You do **not** have the Power BI model, and you cannot
see it. This document contains every definition you need, stated exactly, so
you can implement the VBA side without guessing.

**The single most important constraint in this document:** the email and the
Power BI dashboard must produce identical numbers. They read the same source
files and must apply the same keys, the same filters, and the same proration.
If they diverge, two systems disagree in front of stakeholders and neither can
be trusted. Every definition below is normative — port it exactly, do not
"improve" it.

Where you need a decision from the human, **stop and ask**. A list of open
questions is at the end. Do not guess your way past them.

---

## 1. What exists today (unchanged, do not break)

The database runs a daily/weekly email alert via CDO. It compares ZPT receipts
against an HWA (historical weighted average) forecast and flags surges/drops.

Confirmed VBA ↔ DAX parity was established earlier and must be preserved:

| VBA | DAX |
|---|---|
| `MonthBasisMap()` | `CM Week Expected (N-Month Basis)` |
| `WeekBasisMap()` | `CM Week Expected (N-Week Basis)` |
| `ExpectedMap()` | `CM Bucket Expected` / `CM MTD Expected` |
| `ActMap(mode:=2)` | `CM MTD Actual` |
| `SurgeTest()` mode 2 | `CM MTD Z-Score` |
| `WeekVolatilityMap()` | `CM Week Volatility (N-Week Basis)` |
| `MTDVolatilityMap()` | `CM MTD Volatility` |
| `DueWeek()` / `DueEvalDate()` | `_Due Week` / `_Due Eval Date` |
| `MonthWeekOf()` | `Calendar[MonthWeek]` |

**None of the HWA logic changes in this work.** Leave it alone.

### 1.1 Known pre-existing defects (fix in this release)

**a) `m_effTestMode` has no DAX equivalent.**
In VBA, when `NWeeks < MIN_VOL_WEEKS (3)` while `SurgeTestMode = 2`, the run
falls back to the percent-deviation test. DAX has no `CM Effective Test Mode`.
Dormant while `NWeeks = 12`, but it is a silent divergence waiting to happen.
Flag it to the human; the DAX side needs building too.

**b) `tbl_Config` defaults disagree with the live dashboard.**

| Parameter | `tbl_Config` default | Live dashboard |
|---|---|---|
| NWeeks | 4 | 12 |
| DollarFloor | 0 | 10,000 |
| SurgeThreshold | (differs) | 0.15 |

Also in use on the dashboard: NMonths = 6, SeasonalWeight = 0.4,
SurgeZThreshold = 1.5.

These must be reconciled. Ask the human which side is authoritative before
changing anything.

---

## 2. What is changing

Four workstreams. **A and B ship together. C is independent. D comes last.**

| | Workstream | Status |
|---|---|---|
| A | Drop MML plan, adopt OMM3 at engine grain | Build |
| B | Remove plan from Sales-Grp / Part-Type breakdowns | Build |
| C | Order-count Z-score alongside dollar Z-score | Build, needs decisions |
| D | GTF part-level pipeline | Build last (viable, 71% coverage) |

---

## 3. Workstream A + B — the plan source swap

### 3.1 The structural constraint

The MML plan (`lnk_PlanReceipts` → `tbl_Plan_Local`) is keyed on **Site +
PlanMonth**. The OMM3 plan is keyed on **Customer + Engine Model**, monthly,
for 2026.

OMM3 has **no Sales Grp column and no Part Type column**. It cannot be broken
down that way. An earlier attempt on the Power BI side borrowed a Sales Grp
label from ZPT receipts history via `MAX()`, and this was rejected: a
customer+engine spanning multiple Sales Grps got an arbitrary label, and one
with no receipts history got a blank. Planned dollars were being silently
misattributed to the wrong bucket.

**Consequence:** the email must split into two sections that do not mix.

### 3.2 Required email structure

**Section 1 — HWA anomaly detection.** Unchanged. Plant → Sales Grp →
Part Type breakdowns, actuals vs historical weighted average. **No plan
figures anywhere in this section.** Remove every reference to
`lnk_PlanReceipts` / `tbl_Plan_Local` from Sales-Grp-grained output.

**Section 2 — Plan attainment (new).** Customer + Engine Family grain only.
OMM3 plan vs ZPT actuals. Never broken down by Sales Grp or Part Type.

Anything that currently places a plan number on a Sales-Grp row is removed,
not migrated.

### 3.3 Objects to add

Mirror the existing `lnk_*` → `tbl_*_Local` pattern:

- `lnk_OMM3` — linked to the OMM3 source (see 3.4)
- `tbl_OMM3_Local` — local snapshot, refreshed at run start like
  `tbl_ZPT_Local`

**Do not delete `lnk_PlanReceipts` or `tbl_Plan_Local` yet.** Keep them
dormant for at least one full cycle so old and new output can be diffed.

### 3.4 OMM3 source definition

**File:** `O:\Materials\Restricted\Department\Material Planning\2026\Copy of Sales Forecast 2026 OMM3.xlsx`
**Sheets:** `Summary - PWCS` (plant 5610), `Summary - CAS` (plant 5650)

Both sheets are structured the same way, with these traps:

1. **Column count differs.** PWCS has 38 raw columns, CAS has 34. Any
   positional column reference must be written per-sheet.
2. **Real data ends at a row containing the literal text `"MCE Total"`.**
   Everything after it is non-customer content (GTF duplicate blocks,
   `SI Forecast (AAL) Volume`, `SI Chamber Total`, `Broker sales`,
   `Backlog recovery`, `Go Get Total`, `Grand Total`). Find the position of
   `"MCE Total"` in the Customer column and keep only rows before it.
3. **The customer column is `Sold-to-party Name`.** On the Power BI side this
   string caused repeated parser errors and was renamed to `Customer`
   immediately after header promotion. Do the same, or reference it carefully.
4. **Month columns are dates:** `1/1/2026` through `12/1/2026`, holding the
   monthly plan value.
5. **Dashes.** A `-` in a month cell means "nothing planned" and is converted
   to `0`. This is significant: **60 keys are listed in the file with dashes
   in every month.** They appear in the forecast but are not actually
   forecast. Any "is this customer planned?" test must check
   `SUM(PlannedAmt) > 0`, **not** row existence.
6. **Rows where `Engine Model` is null are dropped.**
7. Columns to keep: Customer, Engine Model, the twelve month columns.
   Plant is assigned per sheet ("5610" / "5650").

The unpivoted result has one row per Customer + Engine Model + Month, with
`PlannedAmt`, `MonthNum` (1–12), `PlanYear` (2026), and `Plant`.

**There is no ActualAmt in OMM3. It is plan-only, by design.** Actuals come
from ZPT.

---

## 4. The key architecture — port this exactly

This is the part most likely to silently diverge. Read carefully.

### 4.1 The join key

```
CustEngPlantKey = <normalised customer> | <engine family> | <plant code>
```

Built identically on both sides:

- **ZPT:** `[Cust Name]`, `[Eng Model]`, `[Plant]` ("5610"/"5650")
- **OMM3:** `[Customer]`, `[Engine Model]`, `[Plant]` ("5610"/"5650")

### 4.2 Text normalisation (`fnNormText`)

Applied to every key component. In Power Query:

```m
t0 = if input = null then "" else Text.From(input),
t1 = Text.Replace(t0, Character.FromNumber(160), " "),   // NBSP -> space
t2 = Text.Upper(Text.Trim(Text.Clean(t1))),
t3 = Text.Combine(List.Select(Text.Split(t2, " "), each _ <> ""), " ")
```

In order: null-safe, replace non-breaking space (U+00A0) with a regular
space, strip control characters, trim, uppercase, then **collapse internal
runs of whitespace to a single space**.

The NBSP and internal-whitespace steps matter. `Text.Clean` and `Text.Trim`
alone do not handle either, and both occur in this data.

### 4.3 Customer name mapping (`fnCustName`)

Applied after normalisation. Currently one entry:

```
"ALL NIPPON AIRWAYS CO.,LTD."  ->  "ALL NIPPON AIRWAYS CO LTD"
```

Same airline, two spellings across the two systems, ~1.74M of plan split
between them.

**Suspected but unconfirmed, do not add without verification:** `SETNA I0`
(letter O vs zero?), `JET MIDWEST. INC`.

### 4.4 Engine family mapping (`fnEngFamily`)

Applied after normalisation. **This is a curated map, not a rule.** Do not
substitute a "strip everything after the first hyphen" algorithm — that would
merge `CFM56-5C` into `CFM56-7B`, which are genuinely different engines.

```
"PW4000-94"       -> "PW4000"
"PW4000-100"      -> "PW4000"
"PW4000-112"      -> "PW4000"
"PW4000-94/100"   -> "PW4000"
"PW4000-100/112"  -> "PW4000"
"V2500_A1"        -> "V2500"
"V2500_A5/D5"     -> "V2500"
"CFM56-5B"        -> "CFM56"
"CFM56-7"         -> "CFM56"
"F117-PW-100"     -> "F117"
"PW150A"          -> "PW150"
```

Anything not in the map passes through unchanged. That includes `PW1000`,
`PW1100G-JM`, `PW1500G`, `PW1900G`, `PW100`, `PW2000`, `GP7000`.

**Why this exists.** OMM3 forecasts `PW4000-94`; ZPT books the receipts as
bare `PW4000`, or as `PW4000-94/100`. Before this map, plan and actuals for
the same commercial relationship landed on different keys and both sides
reported as failures. Merging at family level moved measured attainment from
**0.43 to 1.04** — the earlier figure was a key-matching artifact, not
business performance.

**The one genuinely ambiguous merge is CFM56.** ZPT distinguishes `-5B` and
`-7`; OMM3 does not split them at all. Merging is the only way to compare, but
it should be stated explicitly whenever the number is presented.

### 4.5 Governance — read this

There are now **two independent implementations** of these three functions:
Power Query (M) and VBA (yours). If one gains a mapping and the other does
not, keys drift and nobody notices for months.

**Strongly recommended:** drive both from an Access table, e.g.
`tbl_EngFamilyMap(SourceValue, FamilyValue)` and
`tbl_CustNameMap(SourceValue, MappedValue)`, which Power Query reads via a
linked connection. One row added, both systems updated.

If the human declines that, add a loud comment at both map definitions naming
the other location, and raise it as a known risk.

---

## 5. The measures — exact semantics to reproduce

These are the live DAX definitions. Reproduce the **behaviour**, not the
syntax.

### 5.1 Anchor chain (already exists in VBA — verify only)

```dax
_Due Eval Date =
IF( DAY(TODAY()) <= 7, DATE(YEAR(TODAY()), MONTH(TODAY()), 1) - 1, TODAY() )

_Due Week =
VAR vDay = DAY(TODAY())
RETURN SWITCH( TRUE(), vDay <= 7, 4, vDay <= 14, 1, vDay <= 21, 2, 3 )

_CM Anchor  = [_Due Eval Date]
_CM Start   = DATE( YEAR([_CM Anchor]), MONTH([_CM Anchor]), 1 )
_CM End     = EOMONTH( [_CM Anchor], 0 )
_CM Cutoff  = IF( [CM Weeks Complete] >= 4, [_CM End],
                  DATE( YEAR([_CM Start]), MONTH([_CM Start]),
                        [CM Weeks Complete] * 7 ) )
```

Note the behaviour on the 1st–7th of a month: the anchor jumps back to the
last day of the **previous** month and weeks = 4. On the 8th it jumps forward
to today with weeks = 1. Numbers change sharply on the 8th. This is existing,
intended behaviour — do not alter it, but be aware when comparing runs.

### 5.2 Actuals

```dax
CM YTD Actual =
VAR vStart = DATE( YEAR([_CM Anchor]), 1, 1 )
VAR vTo    = [_CM Cutoff]
RETURN CALCULATE(
    SUM('ZPT_Database (Combined)'[Adjusted Value]),
    REMOVEFILTERS('Calendar'),
    'ZPT_Database (Combined)'[Rcvd Date] >= vStart,
    'ZPT_Database (Combined)'[Rcvd Date] <= vTo )
```

Sum of `Adjusted Value` from 1 January of the anchor year to the cutoff date
inclusive. `CM MTD Actual` is identical but starts at `_CM Start`.

### 5.3 Plan — YTD

```dax
YTD Planned (OMM3) =
VAR vYear  = YEAR( [_CM Anchor] )
VAR vMonth = MONTH( [_CM Anchor] )
VAR vWeeks = [CM Weeks Complete]
VAR vKeys  = VALUES( CustEngDim[CustEngPlantKey] )
VAR vPrior =
    CALCULATE( SUM( OMM3_Sales_Combined[PlannedAmt] ),
        REMOVEFILTERS('Calendar'),
        TREATAS( vKeys, OMM3_Sales_Combined[CustEngPlantKey] ),
        OMM3_Sales_Combined[PlanYear] = vYear,
        OMM3_Sales_Combined[MonthNum] < vMonth )
VAR vCurr =
    CALCULATE( SUM( OMM3_Sales_Combined[PlannedAmt] ),
        REMOVEFILTERS('Calendar'),
        TREATAS( vKeys, OMM3_Sales_Combined[CustEngPlantKey] ),
        OMM3_Sales_Combined[PlanYear] = vYear,
        OMM3_Sales_Combined[MonthNum] = vMonth ) * DIVIDE( vWeeks, 4 )
RETURN IF( ISBLANK(vPrior) && ISBLANK(vCurr), BLANK(),
           COALESCE(vPrior,0) + COALESCE(vCurr,0) )
```

**In words:** every completed month from January up to (not including) the
anchor month, counted in full; plus the anchor month multiplied by
`weeksComplete / 4`. Future months excluded entirely.

The `weeks/4` proration is the existing house convention and matches
`CM Plan Month Total`. It is deliberately not day-based. On the 8th of a
month it counts 25% of the month against ~26% elapsed — a known ~1pp skew,
accepted for consistency with VBA.

### 5.4 Plan — full year

```dax
FY Plan Total (OMM3) =
VAR vYear = YEAR( [_CM Anchor] )
VAR vKeys = VALUES( CustEngDim[CustEngPlantKey] )
RETURN CALCULATE( SUM( OMM3_Sales_Combined[PlannedAmt] ),
    REMOVEFILTERS('Calendar'),
    TREATAS( vKeys, OMM3_Sales_Combined[CustEngPlantKey] ),
    OMM3_Sales_Combined[PlanYear] = vYear )
```

All twelve months, no cutoff, no proration.

### 5.5 Derived comparisons

```dax
YTD Gap (Act - Plan) =
VAR vPlan = [YTD Planned (OMM3)]
RETURN IF( vPlan = 0, BLANK(), COALESCE([CM YTD Actual],0) - vPlan )

YTD Plan Attainment % =
VAR vPlan = [YTD Planned (OMM3)]
RETURN IF( vPlan = 0, BLANK(), DIVIDE( COALESCE([CM YTD Actual],0), vPlan ) )

YTD Pct Change (OMM3 Plan) =
VAR vPlan = [YTD Planned (OMM3)]
VAR vAct  = COALESCE( [CM YTD Actual], 0 )
RETURN IF( vPlan = 0, BLANK(), DIVIDE( vAct - vPlan, vPlan ) )

FY Plan Consumed % =
VAR vFY = [FY Plan Total (OMM3)]
RETURN IF( vFY = 0, BLANK(), DIVIDE( COALESCE([CM YTD Actual],0), vFY ) )
```

**Note the `vPlan = 0` test, not `ISBLANK(vPlan)`.** An all-dash key sums to
0, not blank. Testing for blank would report such a key as being massively
ahead of plan. All four must test `= 0`.

### 5.6 Status banding

```dax
YTD Plan Status =
VAR vPlan = [YTD Planned (OMM3)]
VAR vAct  = COALESCE( [CM YTD Actual], 0 )
VAR vGap  = vAct - vPlan
VAR vPct  = DIVIDE( vGap, vPlan )
VAR vThr  = [Surge Threshold Value]     // percentage, default 0.15
VAR vFlr  = [Dollar Floor Value]        // dollars, default 5000
RETURN SWITCH( TRUE(),
    ISBLANK(vPlan),          BLANK(),
    vPlan = 0 && vAct > vFlr, "Unplanned",
    ABS(vGap) < vFlr,        "On plan",
    vPct >= vThr,            "Over plan",
    vPct <= -vThr,           "Under plan",
    "On plan" )
```

Uses `Surge Threshold` (a percentage), **not** `Surge Z Threshold` (standard
deviations, HWA-only). Do not confuse them.

### 5.7 Match classification

```dax
Match Status =
VAR k = CustEngDim[CustEngPlantKey]
VAR vPlan = CALCULATE( SUM(OMM3_Sales_Combined[PlannedAmt]), ALL(),
                OMM3_Sales_Combined[CustEngPlantKey] = k ) > 0
VAR vAct  = CALCULATE( COUNTROWS('ZPT_Database (Combined)'), ALL(),
                'ZPT_Database (Combined)'[CustEngPlantKey] = k ) > 0
RETURN SWITCH( TRUE(),
    vPlan && vAct, "Both",
    vAct,          "ZPT only (no plan)",
    vPlan,         "Plan only (no receipts)",
    "Neither" )
```

Plan side tests **dollars** (`SUM > 0`), which excludes all-dash keys. Actual
side tests **row existence** across all history, not 2026 — so a key can be
"Both" while having zero YTD actual, meaning a customer forecast for this year
who has not ordered yet. That is intentional and is a finding worth reporting.

The email's plan section should report on `"Both"` keys by default.

---

## 6. Validated reference figures (as at 2026-09-10)

Reproduce these on the VBA side. Any deviation means the port is wrong.

```
anchor=2026-09-10  weeks=1  cutoff=2026-09-07

KEYS (Cust|Family|Plant)
  zpt=375   omm3=160   intersect=155
  dimRows=380   matchStatusBoth=95
  omm3 keys with dollars=100   allDash=60

PLAN
  yr2026 total = 257,848,882
  matched      = 236,257,132   (91.6%)
  plant 5610   = 110,673,297
  plant 5650   = 147,175,585

POPULATION (of 95 matched)
  ytdAct=86   ytdPlan>0=95
  bothYTD=86  planNoAct=9  actNoPlan=0

ADDITIVITY   ytdPlan 0.0%  fyPlan 0.0%  ytdAct 0.0%
PRORATION    ytdPlan/fyPlan = 0.682   (expected 0.688)

RATIOS
  matched:        act=166,479,303  plan=161,127,364  ratio=1.033
  pairedOnly(86): act=166,479,303  plan=160,138,532  ratio=1.040
```

**Headline: ~104% of plan through eight months.** The `pairedOnly` ratio is
the presentable figure.

Proration reads 0.682 against a theoretical 0.688 because the OMM3 plan is
genuinely ~1% back-loaded. Tolerance is 0.01, not 0.002.

---

## 7. Workstream C — order-count Z-score

### 7.1 Intent

Currently volatility and Z-score run on `Adjusted Value` (dollars). Add a
parallel track on **count of distinct sales orders**, so a change in order
*pattern* is detected even when dollar value looks normal.

### 7.2 Blocking prerequisite

`tbl_ZPT_Local` does not currently carry `[Sales Order]`. The union query is:

```sql
SELECT '5610' AS Plant, [Sales Grp], [Cust Name], [Part Type Desc],
       [Eng Model], [Rcvd Date], [Adjusted Value], [Qty Rcvd]
FROM lnk_5610
UNION ALL
SELECT '5650' AS Plant, ... FROM lnk_5650;
```

`[Sales Order]` must be added to both branches, and to the `tbl_ZPT_Local`
schema. Both `lnk_5610` and `lnk_5650` have the column, so this is safe —
but it changes the local table's shape, so check every consumer.

### 7.3 New procedures

Mirror the existing dollar-based ones exactly:

- `OrderCountMap()` — distinct count of `[Sales Order]` per bucket,
  parallel to `ActMap()`
- `OrderVolatilityMap()` — parallel to `WeekVolatilityMap()`, including the
  outlier-trim tie-breaking logic
- `MTDOrderVolatilityMap()` — parallel to `MTDVolatilityMap()`
- Z-score branch in `SurgeTest()`, or a sibling `SurgeTestOrders()`

Reuse the existing structure wherever possible. Do not re-derive the
volatility maths.

### 7.4 Decisions required before building

1. **Both tracks, or replace dollars?** Recommendation: run both, and flag
   when they disagree — normal dollars with anomalous order count means
   average order size moved, which is its own signal.
2. **Orders spanning months.** If a sales order has receipts in three
   different months, does it count once (first receipt) or once per month?
   Distinct-within-bucket is simplest but inflates the annual total.
3. **Minimum-sample floor.** Does `MIN_VOL_WEEKS = 3` apply to the order
   track too?
4. **Threshold.** Reuse `SurgeZThreshold` (1.5), or a separate parameter?
   Order counts are small integers with a different distribution; 1.5 sigma
   does not mean the same thing.
5. **Dollar floor.** `DollarFloor` is meaningless for counts. Does the order
   track need an equivalent minimum-orders floor to suppress noise on
   low-volume groups?

**Do not proceed past this list without answers.**

### 7.5 Parity requirement

Any measure built here must also be built in DAX, or the two systems diverge
again. Produce the DAX definitions alongside the VBA and hand them to the
human for the Power BI side.

---

## 8. Workstream D — GTF: viable, build after A/B

**Status changed 10 September 2026.** An earlier assessment marked this
blocked on a 5-of-75 exact-match result. That test was measuring a
contaminated source and an unnormalised key. After fixing both, **71.1% of the
PWCS GTF plan sits on parts that match ZPT receipts.** Build it — but after
workstreams A and B are shipped and stable, not alongside them.

### 8.1 What GTF is

Part-level forecast: Part Description + Engine + Sales Grp. **No Customer
field** — this is the structural difference from OMM3, and it means GTF gets
its own page and its own dimension table. It cannot reuse `CustEngDim`.

**File:** `O:\Materials\Restricted\Department\Material Planning\2026\GTF 2025 LRP (OMM3 rev).xlsx`
**Sheet:** `GTF forecast (2025 LRP)`
**Scope:** PWCS only (see 8.3).

### 8.2 Source traps — all four must be handled

The sheet is not tabular. Four separate structures are embedded in it, and
each one corrupts the totals if not removed.

**a) Duplicated header block.** Volume Phasing (units) appears first, Sales $
Phasing (dollars) second. After double header promotion, the dollar columns
are `Jan_16` through `Dec_27`. Selecting the first Jan–Dec block gives unit
counts, not money. The engine column is literally named `"Engine "` with a
trailing space.

**b) MIC block.** Rows labelled `MIC number of sets` and
`MIC (number of pieces per month)` introduce a repeat of the S1 parts in
**units**. Truncate the table at the first of those labels.

**c) 24K block.** A row containing the literal `24K`, followed by a repeat of
the part names above it. **Confirmed with the file owner: 24K is the total of
the first block — a subtotal, not a variant.** Including it double-counts the
plan. Truncate at the `24K` row.

An earlier attempt tagged 24K as a variant column rather than removing it.
That was wrong, and it inflated both the plan total and the matched value.

**d) Section headers.** `PWCS` and `CAS` appear as labels in the Part
Description column, dividing the sheet. `PWCS` is consumed by header
promotion and never appears as a data value, so fill-down alone leaves the
first block null — default nulls to `"5610"` after filling.

Also excluded as label rows: `TOTAL`, `AWARDED`, `NOT AWARDED`,
`PART DESCRIPTION`, `2028`, `2029`, and any purely numeric value.

Dashes convert to `0`, as in OMM3.

**Validated result after all four fixes:** 561 rows, 47 distinct parts,
engine value `PW1100` only.

### 8.3 CAS is excluded

With the 24K truncation in place, the CAS block falls outside the retained
range and is dropped. This is correct and matches the original project
decision.

Measured before exclusion: **CAS total 822,807 with 24,497 matched — 3.0%
coverage, 1.8% of total plan.** Nothing of value is lost.

**GTF is a PWCS-only comparison.** State this on any page or email section.

### 8.4 The part signature — how the vocabularies are bridged

GTF and ZPT describe the same parts in different words and different word
order:

| GTF | ZPT |
|---|---|
| `Part Group S1 - Seal Ring Assembly - HPC 4th Stage Outer Air` | `PW1100G HPC-Stage 4 Outer Air Seal Assy` |
| `Part Group S11 - Stator - HPC Exit` | `PW1100G HPC Exit Stator` |
| `Part Group S1 - Seal Ring Assy - HPC 7th Stage Outer Air` | `PW1100G HPC Stg 7 OAS Ring Assembly` |

`fnPartSig` normalises both sides to a **sorted token signature**. Sorting
removes word-order differences while keeping stage numbers significant.

Steps, in order:

1. Uppercase.
2. If the string starts with `PART GROUP`, drop that phrase **and the group
   code token that follows** (`S1`, `R6A`, `S11`…). Do not search for `" - "`
   — the sheet contains `Part Group S7- Case Assembly` and
   `Part Group S14 -Seal Assembly` with inconsistent spacing.
3. Replace every character that is not A–Z, 0–9 or space with a space. **Use
   a keep-list, not a strip-list.** The sheet contains at least three dash
   variants including U+2212; a strip-list missed one and left bare `-`
   tokens on the highest-value parts, blocking ~16M of matches.
4. Apply synonyms: `STG`→`STAGE`, `1ST`→`1` through `8TH`→`8`, `NO`→`NUM`,
   `ASSY`/`ASSLY`/`ASSEM`/`AS`→`ASSEMBLY`, `BRG`→`BEARING`,
   `SUPP`/`SUPPT`→`SUPPORT`, `SEG`→`SEGMENT`, `OAS`→`OUTER AIR SEAL`,
   `RAS`→`REAR AIR SEAL`, `TUBES`→`TUBE`, `VANES`→`VANE`.
5. Drop noise tokens: engine names (`PW1100G`, `PW1100`, `PW4000`, `PW2000`,
   `V2500`, `CFM56`), `GTF`, `OF`, `THE`, `AND`, `RING`.
6. Deduplicate, sort alphabetically, join with spaces.

**Do not add `"A"` to the noise list.** It collapses `Ring-A/O,Locating` to a
lone `O`. (Tested: removing it does not change coverage, since ZPT has no
equivalent part — but the signature is wrong either way.)

**Never merge parts whose numeric tokens differ.** Because tokens are sorted,
digits land first, so comparing token 1 compares the stage number. Stage 5 and
stage 6 are different parts regardless of how similar the rest reads. An
earlier candidate-matcher without this guard proposed stage-4-to-stage-7 and
bearing-1-to-bearing-3 pairs.

Two pairs that look mergeable and are not: `HPC OUTER … STAGE 6` versus
`HPC REAR … STAGE 6` (OAS vs RAS — different seals), and
`HPT STAGE 1 VANE` versus `HPT STAGE 1 VANE SUPPORT` (a vane and its support).

### 8.5 Measured result

```
PWCS (5610)  total plan = 46,015,455 (pre-24K-fix)
             matched    = 32,677,467
             coverage   = 71.1%
CAS  (5650)  coverage   = 3.0%  -> excluded
```

Progression, for the record: 8.6% (wrong engine filter, CAS included) →
30.4% (CAS split out) → **71.1%** (punctuation keep-list, 24K subtotal
removed).

**Matches were audited by eye** against raw descriptions on both sides. The
top 15 by value are all genuinely the same part, with stage numbers aligned
throughout and no cross-stage collisions.

### 8.6 The unmatched 29%

~13.3M of PWCS plan has no matching ZPT receipts. The largest items:

- `Part Group S7 - MTF (TIC)` and `TIC (MTF)` — note GTF itself writes this
  part two ways
- `Part Group S13 - Support Assembly - No. 4 Bearing Front Face Seal` and
  its near-twin
- `Ring-A/O,Locating,stg1/2/3,upper,lower` — no ZPT equivalent exists
- `1st/2nd/3rd Stg Vane Support (Upper & Lower)`

These are mostly parts ZPT has not booked receipts against, not names that
could be bridged. **Report the 29% gap alongside any GTF attainment figure.**
Do not present GTF attainment as covering the whole forecast.

### 8.7 Build instruction

Build GTF **after** workstreams A and B are shipped and stable. It is a
separate page and a separate email section, not an extension of the OMM3 work.

1. Port `fnPartSig` to VBA in `modKeys`, alongside the other three functions.
   Same governance risk as §4.5 — two implementations, one meaning.
2. Build `PartDim` on the part signature, mirroring `CustEngDim`.
3. Plan measures with `TREATAS` on the signature key, following the §5
   definitions exactly — same anchor, same `weeks/4` proration.
4. Validate with a GTF DIAG before anything reaches the email.
5. GTF has a `Sales Grp` column of its own, unlike OMM3. Confirm with the
   human whether GTF may be broken down by it — the §3.1 objection does not
   apply here, since the Sales Grp is native to the plan rather than borrowed
   from receipts.

### 8.8 Still open on GTF

- **Row count.** 561 rows ÷ 12 months = 46.75 against 47 distinct parts. Does
  not divide cleanly. Either one part is missing a month row or two parts
  share a description. Investigate before building.
- **Part Number.** GTF has one; ZPT reportedly does not. If ZPT has an
  equivalent field anywhere, a part-number join would beat signature matching
  outright and should replace it. Verify rather than assume.
- **ZPT engine families.** `PW1000`, `PW1100G-JM`, `PW1500G`, `PW1900G` are
  separate families in ZPT; GTF writes `PW1100` and `PW1500`. The coverage
  query names all four explicitly. Whether they should be consolidated in
  `fnEngFamily` is undecided — they may be genuinely different engines.
- **ZPT part descriptions under the GTF engine family include other engines'
  parts** (`PW4000 100" 1.6-2-3 STG STATOR`, `PW2000 FNG`,
  `V2500 NO.3 BEARING FACE SEAL ASSY`). Either miscoded receipts or
  unreliable naming. Flag to the human; it does not block the build but it
  bounds how much the remaining 29% can improve.

---

## 9. Build sequence

Do not reorder. Each step is verifiable before the next.

1. **`modKeys`** — `NormText()`, `MapCustName()`, `MapEngFamily()`, plus the
   backing tables if the human approves 4.5. Write a test procedure that
   feeds known inputs and asserts the outputs, including: NBSP handling,
   double-space collapse, `PW4000-94/100` → `PW4000`,
   `ALL NIPPON AIRWAYS CO.,LTD.` → `ALL NIPPON AIRWAYS CO LTD`.
2. **`lnk_OMM3` / `tbl_OMM3_Local`** — with the MCE Total truncation, dash
   handling, null-engine filter, and per-sheet column counts.
3. **Key-parity check** — a procedure that outputs distinct
   `CustEngPlantKey` counts for ZPT and OMM3 and their intersection. Must
   reproduce **375 / 160 / 155** (allowing for data refreshed since
   2026-09-10).
4. **Plan measures** — YTD, FY, gap, attainment, status. Verify against §6.
5. **Email restructure** — two sections, plan removed from Sales-Grp output.
6. **Parallel run** — one full cycle producing both old and new output. Diff
   them. Only then retire MML.
7. **Order-count Z-score** — after §7.4 is answered.
8. **`CM Effective Test Mode`** — close the pre-existing gap.
9. **GTF** — per §8.7, only once A/B are stable.

---

## 10. Open questions — ask the human, do not guess

**Scope and presentation**

1. Should plan attainment appear in the **daily** email at all? A YTD
   attainment figure moves very little day to day. Monthly may be more
   appropriate.
2. How should the plant split be presented? Plant is in the key, so PWCS and
   CAS can be reported separately — but a customer with plan in one plant and
   receipts in the other shows as under-plan in one and unplanned in the
   other, reconciling only in the combined view. Is that mismatch a finding
   to surface, or should the email report combined only?
3. Should these be reported as their own findings?
   - **60 keys** listed in the OMM3 file with dashes in every month
   - **9 keys** with 2026 plan and historical receipts but nothing this year
   - **21.6M** of plan (8.4%) on customers ZPT has never seen

**Configuration**

4. `tbl_Config` vs live dashboard mismatch (§1.1b) — which is authoritative?
5. Should the engine-family and customer-name maps be table-driven so Power
   BI and VBA share one source (§4.5)?

**Order-count Z-score**

6. All five questions in §7.4.

**GTF**

7. May GTF be broken down by its own `Sales Grp` column? Unlike OMM3, GTF
   carries one natively, so the §3.1 objection does not apply — but confirm
   before building it into a page or email section.
8. Does ZPT hold a part number anywhere? A part-number join would replace
   signature matching entirely (§8.8).

**Plan versioning (known, unresolved by design)**

9. Both YTD and FY plan measures always read the *current* OMM3 file, for all
   months including closed ones. If the file is revised mid-year, YTD figures
   for already-closed months shift retroactively and silently. There is no
   snapshot of what the plan said at the time. Should the VBA snapshot
   `tbl_OMM3_Local` monthly to create one?

---

## 11. Things that must not change

- All HWA logic and its VBA ↔ DAX parity (§1)
- `CustEngKey` and `PartEngKey` on the ZPT side — other pages and measures
  depend on them
- `[Eng Model]` — the raw column is preserved; family mapping goes in a
  **new** column
- The existing PWCS/CAS plant-level pages and
  `MTD Planned Receipts by Sales Group (using EST ratio)`
- The `weeks/4` proration convention
- The CDO send mechanism and Task Scheduler wiring
- `fnPartSig`'s numeric-token guard (§8.4) — parts whose stage or bearing
  numbers differ must never merge, whatever their similarity score
