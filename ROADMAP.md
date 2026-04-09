# Pharmacy Tools Add-In — Development Roadmap

## Current Version: 2.3.3

---

## Version 2.3.3 — Completed (2026-04-09)

### Bugfix: m_LastHttpStatus Scope
`m_LastHttpStatus` was incorrectly declared inside `HttpPost()` in v2.3.2,
making it invisible to the batch loop's 400-retry logic. Moved to module level.

---

## Version 2.3.2 — Completed (2026-04-09)

### Distribution UUID Auto-Discovery
Added `FetchDistUUID()`, which queries the CMS DKAN metastore at
`/api/1/metastore/schemas/dataset/items/{dsID}?show-reference-ids=true`
before each year's batch run to retrieve the live distribution UUID.
Hardcoded entries in `KnownDistUUIDs()` now serve only as a session-start
fast-path; UUID rotations that previously caused silent 400 failures are
handled automatically.

UUID rotation events are written to the log:
`UUID rotated for YYYY: <old UUID> -> <new UUID>`

### POST 400 Auto-Retry
If a batch POST returns HTTP 400, the add-in re-fetches the distribution UUID
once from the metastore and retries the batch before falling back to GET.
`m_LastHttpStatus` (module-level Long) carries the HTTP status out of
`HttpPost()` to the calling loop.

---

## Version 2.3.1 — Completed (2026-04-09)

### 2026 Dataset Support
Added 2026 dataset UUID (`fbb83258-11c7-47f5-8b18-5f8e79f7e704`) and
distribution UUID (`8b801945-2507-5057-8f97-eb7586889f3d`) to the hardcoded
tables. POST now works for 2026 dispense dates without requiring a manual
Manage Datasets entry.

---

## Version 2.3 — Completed

### 1. Full Dataset Manager (replaces Manage Datasets wizard)
Replaced the 4-step MsgBox wizard with a MsgBox/InputBox-driven manager
showing all configured years (built-in and custom) in a single scrollable
view. Users can view, add, edit, and delete dataset entries. Built-in entries
are labelled "Built-in"; user overrides are labelled "Built-in*"; custom
entries are labelled "Custom". Registry-backed persistence.

### 2. Row Colour Reset on Successful Re-Lookup
At the start of Pass 2, all data-row interior colours are reset to xlNone
before red highlights are re-applied. Highlights now reflect the current run
only, not accumulated history.

---

## Version 2.3.5 — In Progress

### Dataset Manager UI Replacement

**Background:**
The v2.3 Dataset Manager was implemented using MsgBox/InputBox chains — an
acknowledged shortcut that does not meet the original requirement for a proper
form-based interface. v2.3.5 replaces it with the correct implementation.

**Planned behaviour:**
A modeless floating sidebar UserForm (Option C), launched from the existing
Manage Datasets ribbon button. Styled similarly to Excel's built-in Name
Manager. Stays open while the user works.

The sidebar displays a grid of all configured years with the following columns:
- Year
- Dataset ID (UUID)
- Distribution UUID
- Type (Built-in / Custom / Override)

Actions available:
- Add (new year)
- Edit (selected row)
- Delete / Reset (selected row)
- Close

A read-only detail panel below the list expands the selected row's full UUIDs.
Window position is persisted in the registry so the sidebar reopens where the
user left it.

---

## Version 2.4 — Planned

### 1. Multi-Year Batch Optimization

**Current behaviour:**
All rows are bucketed by year and each year is queried separately in
sequence. Batches within a year are efficient (15 NDCs per POST), but
years themselves are processed one at a time with no parallelism.

**Planned behaviour:**
Where a spreadsheet contains rows spanning multiple years, process all
years' batch queues back-to-back in a single pass rather than waiting
for each year to complete before starting the next. Results are merged
into the shared API cache before Pass 2 begins, reducing total runtime
proportionally to the number of years present.

Note: VBA is single-threaded — "concurrent" here means interleaved
sequential dispatch, not true parallelism.

---

### 2. Per-Row Lookup Status Column

**Planned behaviour:**
Insert a "NADAC Status" column immediately after NADAC Effective Date.
Populate it on every run with one of the following standardised codes:

| Code        | Meaning                                              |
|-------------|------------------------------------------------------|
| OK          | Rate found and written successfully                  |
| NOT_FOUND   | NDC not present in the NADAC dataset                 |
| NO_RATE     | NDC found but no rate on or before the dispense date |
| BAD_NDC     | NDC value is missing, malformed, or could not be cleaned |
| BAD_DATE    | Dispense date is missing or could not be parsed      |
| NO_DATASET  | Annual dataset could not be resolved for this row's year |

The column should be formatted as plain text and reset on each execution.

---

### 3. Unfound Rows Moved to Separate Sheet

**Planned behaviour:**
At the end of each lookup run, any row with a status of NOT_FOUND or
NO_RATE should be:

- Removed from the main data sheet (iterate in reverse to avoid index drift)
- Copied in full to a new sheet named "NADAC Not Found"
- The sheet is created fresh on each run (not appended)
- A summary line at the top shows run date/time and row count
- The main sheet retains only rows with status OK

---

## Version 2.5 — Planned

### Margin Column

**Calculation:**
  Margin = Paid Amount - (NADAC Per Unit x Quantity)

**Requirements:**
- Auto-detect or prompt for the Paid Amount column and Quantity column
  (same pattern as existing NDC/Date column detection)
- Insert a "Margin" column immediately after NADAC Effective Date
- Format as currency ($#,##0.00)
- Negative margins (paid below NADAC) highlighted in light orange
- Recalculates on every lookup run alongside the NADAC values
- If Quantity column cannot be found, prompt the user to identify it

---

## Version 3.0 — Planned

### Dataset Manager — Full VBA UserForm (Modal Dialog)

**Background:**
v2.3.5 delivers the modeless sidebar (Option C) as the near-term solution.
v3.0 replaces it with a full modal VBA UserForm (Option A) — a proper dialog
with a grid, Add/Edit/Delete buttons, and native event handlers. This is the
originally intended implementation and provides the most polished, native
Excel feel.

**Planned behaviour:**
- Modal dialog launched from the Manage Datasets ribbon button
- Grid (ListBox or MSForms DataGrid) showing all configured years
- Inline or panel-based editing
- Full event handler architecture (no MsgBox/InputBox substitutes)
- Distinct visual treatment for Built-in vs Custom vs Override entries

---

## Future Considerations (backlog, unscheduled)

- Additional tools in the Pharmacy Tools ribbon (TBD)
- Export lookup log summary to a new sheet
- Scheduled/automated refresh via Windows Task Scheduler

---

*Last updated: 2026-04-09*
*Contact: Jeremiah Maudlin | jmaudlin@omegaim.com | +1 (641) 660-6367*
# Pharmacy Tools Add-In — Development Roadmap

## Current Version: 2.3

---

## Version 2.3 — Completed

### 1. Full Dataset Manager (replaces Manage Datasets wizard)
Replaced the original 4-step MsgBox wizard with a streamlined MsgBox/InputBox
chain showing all configured years (built-in and custom) in a single scrollable
view. Users can view, add, edit, and delete dataset entries. Built-in entries
are labelled "Built-in"; user overrides are labelled "Built-in*"; user-added
years are labelled "Custom". Registry-backed persistence.

### 2. Row Colour Reset on Successful Re-Lookup
At the start of Pass 2, all data-row interior colours are reset to xlNone
before red highlights are re-applied. Highlights now reflect the current run
only, not accumulated history.

---

## Version 2.3.5 — In Progress

### Dataset Manager UI Replacement

**Background:**
The v2.3 Dataset Manager was implemented using MsgBox/InputBox chains — an
acknowledged shortcut that does not meet the original requirement for a proper
form-based interface. v2.3.5 replaces it with the correct implementation.
This is an interim solution; v3.0 delivers the final full modal UserForm.

**Planned behaviour:**
A modeless floating sidebar UserForm (Option C), launched from the existing
Manage Datasets ribbon button. Styled similarly to Excel's built-in Name
Manager. Stays open while the user works.

The sidebar displays a grid of all configured years with the following columns:
- Year
- Dataset ID (UUID)
- Distribution UUID
- Type (Built-in / Custom / Override)

Actions available:
- Add (new year)
- Edit (selected row)
- Delete / Reset (selected row)
- Close

A read-only detail panel below the list expands the selected row's full UUIDs.
Window position is persisted in the registry so the sidebar reopens where the
user left it.

---

## Version 2.4 — Planned

### 1. Multi-Year Batch Optimization

**Current behaviour:**
All rows are bucketed by year and each year is queried separately in
sequence. Batches within a year are efficient (15 NDCs per POST), but
years themselves are processed one at a time.

**Planned behaviour:**
Where a spreadsheet contains rows spanning multiple years, pre-queue all
years' batch lists before making any API calls, then process them in an
interleaved sequence so results from all years are merged into the shared
API cache before Pass 2 begins. This reduces total runtime proportionally
to the number of years present.

> **Note:** VBA is single-threaded. "Concurrent" here means interleaved
> pre-queued batches merged before Pass 2 — not OS-level threading.

---

### 2. Per-Row Lookup Status Column

**Planned behaviour:**
Insert a "NADAC Status" column immediately after NADAC Effective Date.
Populate it on every run with one of the following standardised codes:

| Code        | Meaning                                              |
|-------------|------------------------------------------------------|
| OK          | Rate found and written successfully                  |
| NOT_FOUND   | NDC not present in the NADAC dataset                 |
| NO_RATE     | NDC found but no rate on or before the dispense date |
| BAD_NDC     | NDC value is missing, malformed, or could not be cleaned |
| BAD_DATE    | Dispense date is missing or could not be parsed      |
| NO_DATASET  | Annual dataset could not be resolved for this row's year |

The column should be formatted as plain text and reset on each execution.

---

### 3. Unfound Rows Moved to Separate Sheet

**Planned behaviour:**
At the end of each lookup run, any row with a status of NOT_FOUND or
NO_RATE should be:

- Removed from the main data sheet
- Copied in full to a new sheet named "NADAC Not Found"
- The sheet is created fresh on each run (not appended)
- A summary line at the top shows run date/time and row count
- The main sheet retains only rows with status OK

> **Implementation note:** Delete rows in reverse order (bottom-up iteration)
> to avoid row-index drift when rows are removed mid-pass. Alternatively,
> collect all row indices first, then delete in a single reverse-sorted pass.

---

## Version 2.5 — Planned

### Margin Column

**Calculation:**
  Margin = Paid Amount - (NADAC Per Unit x Quantity)

**Requirements:**
- Auto-detect or prompt for the Paid Amount column and Quantity column
  (same pattern as existing NDC/Date column detection)
- Insert a "Margin" column immediately after NADAC Effective Date
- Format as currency ($#,##0.00)
- Negative margins (paid below NADAC) highlighted in light orange
- Recalculates on every lookup run alongside the NADAC values
- If Quantity column cannot be found, prompt the user to identify it

---

## Version 3.0 — Planned

### Dataset Manager — Full VBA UserForm (Modal Dialog)

**Background:**
v2.3.5 delivers the modeless sidebar (Option C) as the near-term solution.
v3.0 replaces it with a full modal VBA UserForm (Option A) — a proper dialog
with a grid, Add/Edit/Delete buttons, and native event handlers. This is the
originally intended implementation and provides the most polished, native
Excel feel. v3.0 supersedes the v2.3.5 sidebar; both will not coexist.

**Planned behaviour:**
- Modal dialog launched from the Manage Datasets ribbon button
- Grid (ListBox or MSForms DataGrid) showing all configured years
- Inline or panel-based editing
- Full event handler architecture (no MsgBox/InputBox substitutes)
- Distinct visual treatment for Built-in vs Custom vs Override entries

---

## Known Issues / Tech Debt

| Item | Notes |
|------|-------|
| `ShowDatasetManager` Dim hoisting | Variables like `dsID`, `distID`, `eType` are declared inside a `For` loop body. VBA hoists all `Dim` to procedure scope so they do not reset per iteration. Works correctly but is misleading — resolved when v2.3.5 UserForm replaces this code. |
| `BestRecord()` sort assumption | Returns the first record where `effective_date <= dispDt`, relying on the API returning results in descending order. The POST body requests `desc` and the GET fallback requests `ORDER DESC` — assumption is valid but implicit. |
| `DiscoverDatasetID()` exact title match | Auto-discovery matches the exact CMS dataset title string. If CMS changes their naming convention, discovery silently fails. A fuzzy-match fallback is desirable long-term. |
| Stale `.SYNOPSIS`/`.DESCRIPTION` | Installer comment block still references "frmDatasetManager UserForm" — corrected in v2.3.5 release. |

---

## Future Considerations (backlog, unscheduled)

- Additional tools in the Pharmacy Tools ribbon (TBD)
- Export lookup log summary to a new sheet
- Scheduled/automated refresh via Windows Task Scheduler

---

*Last updated: 2026-04-06*
*Contact: Jeremiah Maudlin | jmaudlin@omegaim.com | +1 (641) 660-6367*
