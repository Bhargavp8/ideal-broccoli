# RBDB Email Automation — Operator Runbook

**What this covers:** getting the right data into the Access database, and
getting the scheduled email to send.

**Who this is for:** whoever owns the database day to day, including whoever
inherits it.

---

## 0. What the system actually does

A `.accdb` on your machine (or a share) links — **read only** — to three
Excel/data sources on the O: drive. Once a day, Task Scheduler opens it, an
AutoExec macro calls `RBDB_RunAlert()`, and the code:

1. copies the linked data into **local** tables (`tbl_ZPT_Local`,
   `tbl_Plan_Local`), rebuilding them from scratch;
2. works out whether an email is **due** today, from the calendar;
3. if one is, checks the staged data is current enough to report on;
4. computes the anomaly analysis and sends via CDO through
   `mailhub.utc.com:25`.

> **Nothing in this system ever writes to the O: drive.** Every source is
> opened read-only. If someone tells you the automation corrupted a source
> workbook, that did not come from here.

---

## 1. One-time setup

Do these **in order**. Steps 1.1–1.3 are prerequisites; the code raises a
specific error if any is missing.

### 1.1 Add the VBA reference

In the VBA editor: **Tools → References →** tick
**Microsoft Office _xx.x_ Access Database Engine Object Library**.

Without it, `Dim db As DAO.Database` will not compile and *nothing* runs.

### 1.2 Create the linked tables

**External Data → New Data Source → From File → Excel → Link to the data
source by creating a linked table.**

| Linked table | Points at | Must expose |
|---|---|---|
| `lnk_5610` | ZPT receipts, plant 5610 (PWCS) | Sales Grp, Cust Name, Part Type Desc, Eng Model, Rcvd Date, Adjusted Value, Qty Rcvd |
| `lnk_5650` | ZPT receipts, plant 5650 (CAS) | same as above |
| `lnk_PlanReceipts` | `O:\Finance\Common\Receipts & Backlog Automation\PBI\Receipts Plan and Reported.xlsx`, **Sheet1** | Site, Month, Planned Receipts, Reported Receipts |

**Trap in `lnk_PlanReceipts`:** the `Site` column is matched with a
*contains* test — `InStr(Site, 'PWCS')`, then `InStr(Site, 'CAS')`. Anything
matching neither is staged as `'Unknown'` and **silently contributes nothing
to either plant's plan**. If the plan columns in the email read as zero, check
this first:

```sql
SELECT PlantCd, Count(*) FROM tbl_Plan_Local GROUP BY PlantCd;
```

Any rows under `Unknown` are plan dollars that have gone missing.

### 1.3 Create the union query

Create a query named exactly **`qry_ZPT_Combined`**:

```sql
SELECT '5610' AS Plant, [Sales Grp], [Cust Name], [Part Type Desc],
       [Eng Model], [Rcvd Date], [Adjusted Value], [Qty Rcvd]
FROM lnk_5610
UNION ALL
SELECT '5650' AS Plant, [Sales Grp], [Cust Name], [Part Type Desc],
       [Eng Model], [Rcvd Date], [Adjusted Value], [Qty Rcvd]
FROM lnk_5650;
```

The literal `Plant` column is **required**. Without it the run stops with
*"qry_ZPT_Combined has no [Plant] column"* — the two plants cannot be reported
separately and every plan share would be wrong.

### 1.4 Run the setup procedure — **once**

In the VBA editor's Immediate window:

```
RBDB_Setup()
```

> ### ⚠ `RBDB_Setup()` is a FIRST-RUN-ONLY procedure
>
> It **drops and recreates `tbl_Config` and `tbl_AlertLog`**. Running it on a
> live database throws away every tuned threshold and the entire send history.
> There is no undo and no confirmation prompt.
>
> If you only want to re-stage the data, run `RBDB_Test()` instead.

What it does: creates `tbl_Config` with defaults, creates `tbl_AlertLog`,
stages the local tables, and **seeds the log so every week up to today is
marked as already reported** — that is deliberate, so setting the system up
mid-month doesn't fire three catch-up emails at people.

It finishes with a message box showing the staged row counts per plant. Check
the **"unmatched"** count is 0. Anything there is a receipt row whose Plant is
neither 5610 nor 5650, and it will be invisible in the email.

### 1.5 Check the settings

```
RBDB_Parity()
```

Prints every setting the Power BI dashboard depends on and flags the ones
that would make the email disagree with it. `[ok]` on all lines means the two
should tie out.

Run this **first** whenever anyone reports that the email and the dashboard
disagree. It is faster than reconciling two numbers by hand and it is usually
the answer.

### 1.6 Set the recipients

Open `tbl_Config` and find the row where `ParamName = 'Recipients'`.

- The addresses go in the **`Notes`** column, **not** `ParamValue`.
- Separate multiple addresses with **semicolons**.

```
Notes: someone@prattwhitney.com; someone.else@prattwhitney.com; shared.mailbox@prattwhitney.com
```

> ### ⚠ Do this before anyone leaves
>
> If `Notes` is empty the code falls back to a **single hard-coded named
> address** — and the log records the send as successful either way. On the
> day that account is deprovisioned, the email stops arriving and
> `tbl_AlertLog` keeps saying it was sent. **Add at least one colleague and a
> shared mailbox.**

### 1.7 Wire up the schedule

1. **AutoExec macro** in the database: a single `RunCode` action calling
   `RBDB_RunAlert()`.
2. **Task Scheduler**: a daily task that opens the `.accdb`. Daily is correct
   — the code itself decides which days actually send (see §3).

> **Never rename `RBDB_Setup`, `RBDB_Test` or `RBDB_RunAlert`.** Task
> Scheduler and AutoExec call them by name. Rename one and the scheduled run
> stops firing — silently.

### 1.8 Send a test

```
RBDB_Test()
```

> ### ⚠ `RBDB_Test()` sends a real email to the real recipients
>
> It is not a dry run. It forces the current week to send even if the log
> already records it as sent. Point `Recipients` at yourself first if you
> don't want the distribution list to receive it.

If nothing is outstanding it shows a summary box instead (trigger mode, pace
mode, bucket due, window end, data reach) without sending.

---

## 2. What has to be true every day

The automation is only as current as the workbooks behind it.

| Source | Who refreshes it | What breaks if it's stale |
|---|---|---|
| ZPT receipts (`lnk_5610`, `lnk_5650`) | the upstream extract | The whole email. An empty window reads as **every group at −100%** — a false alarm that looks catastrophic. |
| `lnk_PlanReceipts` | Finance, monthly | Plan columns read as zero or against a stale month. |

The code protects against the first case rather than trusting it:

- **`DataLagGraceDays`** (default 3) — data is accepted as current if its
  latest date is within this many days of *today*, even if it stops short of
  the window end. The last days of a bucket genuinely have no receipts
  sometimes.
- **`MaxDeferDays`** (default 3) — inside this, a due email **defers** and
  tomorrow's run tries again. Past it, the email **sends anyway** with an
  amber banner naming the date the data actually reaches, and a note in the
  subject line.

**It never silently withholds, and it never silently lies.** But a deferral
is only visible in `tbl_AlertLog` — nobody is emailed to say "your report was
held". If the report stops arriving, that log is the first place to look.

---

## 3. When it sends

The scheduler fires **daily**. The code sends on the first day of each bucket,
reporting the bucket that just closed:

| Day of month | Reports |
|---|---|
| 1–7 | **week 4 of the previous month** |
| 8–14 | week 1 of this month |
| 15–21 | weeks 1–2 |
| 22+ | weeks 1–3 |

`tbl_AlertLog` is the once-only guard: a row with matching `EvalYM` +
`EvalWeek` and `Sent = True` means that bucket is closed.

**Catch-up:** a bucket missed because the machine was off, Access crashed, or
a run errored is still unsent in the log, so the next daily run picks it up.
If several buckets went by, every one is logged as *superseded* and **only the
newest is emailed** — four catch-up emails arriving together get deleted
unread.

**Quiet weeks still close.** A bucket where nothing breached logs
`Sent = True` with "no groups flagged", so it doesn't re-evaluate every day.

---

## 4. Routine maintenance

| When | Do this |
|---|---|
| After **any** edit to `tbl_Config` | `RBDB_Parity()` |
| After adding a customer or engine mapping | `RBDB_TestKeys()`, then `RBDB_ExportMaps()` and paste into the `.pbix` |
| Monthly | Scan `tbl_AlertLog` for `Sent = False` and for rows whose `Detail` starts with `ERR` |
| When someone says the numbers disagree with the dashboard | `RBDB_Parity()` first, then `Diag "5610", "<sales group>"` |

### Checking the log

```sql
SELECT RunDate, EvalYM, EvalWeek, GroupsFlagged, Sent, Detail
FROM tbl_AlertLog
ORDER BY RunDate DESC;
```

`RBDB_RunAlert()` **swallows errors into this table** rather than throwing a
dialog at an unattended machine. A row with `Sent = False` and a `Detail`
beginning `ERR` is a failed run that nobody was told about. This is the only
place it shows up.

---

## 5. When something goes wrong

| Symptom | Cause | Fix |
|---|---|---|
| *"qry_ZPT_Combined has no [Plant] column"* | The union query is missing the `'5610' AS Plant` literal | §1.3 |
| *"Linked table 'lnk_PlanReceipts' not found"* | Link broken or renamed | External Data → Linked Table Manager, re-point it |
| *"the expression you entered has a function name that Microsoft Access can't find"* | The VBA project will not compile, so Access cannot resolve `PlanNum()` inside the staging query | **Debug → Compile** in the VBA editor, fix the compile error, re-run |
| *"table already exists"* on a staging table | The table is open in Datasheet or Design view, so the drop failed | Close it and re-run |
| Email arrives with an **amber banner** | Data was short and `MaxDeferDays` expired — it sent anyway rather than withhold | Refresh the source workbook; the next run reports in full |
| No email at all, no error | Either nothing breached (check `tbl_AlertLog` for "no groups flagged"), or the run deferred, or the task didn't fire | Log first, then Task Scheduler history |
| Plan columns are all `-` or zero | `Site` values didn't match PWCS/CAS | Run the `tbl_Plan_Local` check in §1.2 |
| Every group shows −100% | The staged data is empty for the window | Check `RefDateGlobal` — `SELECT Max(RcvdDate) FROM tbl_ZPT_Local;` |

### Two editing hazards, if you go into the code

- **Never declare a variable called `nZ`.** VBA is case-insensitive and the
  editor auto-capitalises, so a local `nZ` silently rewrites every `Nz()` call
  in the module into `nZ()` and the project stops compiling.
- **Key expressions passed to `AggMap` are raw SQL.** Column names, the `&`
  operator and string literals only. A function the query engine cannot
  resolve there fails the entire run.

---

## 6. Known risks at handover

These are flagged in the code itself and are still open:

1. **`Recipients` is one named person.** See §1.6. Highest-priority fix.
2. **The dashboard link in the email points at a `.pbix` file.** After the
   Report Server migration that URL downloads a multi-hundred-MB file that
   only opens in Power BI Desktop, instead of opening the report. Replace
   `REPORT_URL` with the Report Server report URL.
3. **`MAIL_FROM` is registered with the relay.** Changing
   `RBDB_Anomaly_Alert@prattwhitney.com` requires IT, not a code edit.
4. **`tbl_Config` defaults disagree with the live dashboard** — `NWeeks` (4 vs
   12), `DollarFloor` (0 vs 10,000), `SurgeThreshold` (0.25 vs 0.15). Until
   this is settled, the email and the dashboard are answering slightly
   different questions. `RBDB_Parity()` currently asserts `NWeeks = 4`, so
   changing it to match the dashboard will make the parity check report a
   failure until that expectation is updated too.

---

## 7. Additional setup for the OMM3 plan work

**Not yet required.** Do these only when the OMM3 changes are being installed;
until then the steps above are the whole system.

### 7.1 Create the key map tables

```
RBDB_SetupKeyMaps()
RBDB_TestKeys()
```

`RBDB_SetupKeyMaps()` creates `tbl_CustNameMap` and `tbl_EngFamilyMap` and
seeds them. It is **safe to re-run** — an existing table is left completely
alone, so mappings you added by hand are never discarded. (`RBDB_ReseedKeyMaps()`
is the deliberate destructive version, and it asks first.)

`RBDB_TestKeys()` must report **all checks passed** before anything is built
on these keys. Every downstream plan figure depends on them.

### 7.2 Keep Power BI in step

The customer-name and engine-family mappings exist **twice** — once here, once
in Power Query. If one side gains a mapping and the other doesn't, keys drift
apart and nobody notices for months: the totals still add up, they are just
measuring the wrong thing.

So whenever you add a mapping:

1. Add the row to `tbl_CustNameMap` or `tbl_EngFamilyMap`.
2. Run `RBDB_TestKeys()`.
3. Run `RBDB_ExportMaps()` — it writes `RBDB_KeyMaps.m` next to the `.accdb`.
4. Paste that into the `.pbix`.

**All four steps.** Access is the master; the Power Query side is
*regenerated*, never hand-edited.

**Source values must be stored already-normalised** (uppercased, trimmed,
internal double-spaces collapsed). `RBDB_TestKeys()` checks every stored row
for this and names any that are wrong — a source value that isn't normalised
can never match anything, and nothing else would ever tell you.

### 7.3 Link the OMM3 forecast

**File:** `O:\Materials\Restricted\Department\Material Planning\2026\Copy of Sales Forecast 2026 OMM3.xlsx`

Two sheets, and **they need two linked tables** — Access links one worksheet
per linked table, so the single `lnk_OMM3` the design brief describes becomes:

| Linked table | Sheet | Plant |
|---|---|---|
| `lnk_OMM3_5610` | `Summary - PWCS` | 5610 |
| `lnk_OMM3_5650` | `Summary - CAS` | 5650 |

They are then unioned the same way `qry_ZPT_Combined` unions the two ZPT
links. **The two sheets have different column counts** (PWCS 38, CAS 34), so
each leg of that union has to be written separately — do not copy one and
change the table name.

**Before linking, open the file and check which row holds the header**
(the one containing `Sold-to-party Name`):

- **Row 1** → link the sheet directly, with *First Row Contains Column
  Headings* ticked.
- **Any other row** → create a **named range** in the workbook covering the
  header row and the data below it, and link the *named range* instead.
  Access cannot skip leading rows on a sheet link, and getting this wrong
  gives you a table whose column names are `F1`, `F2`, `F3`…

Reference the customer column as `[Sold-to-party Name]` — the hyphen means it
always needs the square brackets.

### 7.4 Add `[Sales Order]` to the union query

Both `lnk_5610` and `lnk_5650` already carry the column. Add it to **both**
legs of `qry_ZPT_Combined`:

```sql
SELECT '5610' AS Plant, [Sales Grp], [Cust Name], [Part Type Desc],
       [Eng Model], [Rcvd Date], [Adjusted Value], [Qty Rcvd], [Sales Order]
FROM lnk_5610
UNION ALL
SELECT '5650' AS Plant, [Sales Grp], [Cust Name], [Part Type Desc],
       [Eng Model], [Rcvd Date], [Adjusted Value], [Qty Rcvd], [Sales Order]
FROM lnk_5650;
```

This is needed for the order-count analysis later. Adding it now is
deliberate: it is one schema change instead of two, and the column sits
harmlessly unused until that work is built.

### 7.5 Two things to confirm on the Power BI side

Both affect whether the email and the dashboard agree, and neither can be
checked from the Access side:

1. **Is `[CM Weeks Complete]` derived from the calendar or from the latest
   receipt date?** It must be the calendar (`_Due Week`). If it reads the
   data, the dashboard silently freezes on a stale window whenever the
   workbook stops refreshing — the exact failure this database was rebuilt to
   avoid.
2. **Which dollar floor does `YTD Plan Status` actually read?** The
   documentation says 5,000 in one place and 10,000 in another. Whichever the
   model resolves at runtime is the one the email has to match.
