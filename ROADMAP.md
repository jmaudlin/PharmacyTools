# Pharmacy Tools Add-In — Development Roadmap

## Current Version: 2.3

---

## Version 2.3 — Completed

### 1. Full Dataset Manager (replaces Manage Datasets wizard)
Replaced the 4-step MsgBox wizard with a modeless sidebar UserForm showing
all configured years (built-in and custom) in a single view. Users can view,
add, edit, and delete dataset entries. Built-in entries are labelled
"Built-in"; user overrides are labelled "Custom". Registry-backed persistence.

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
Where a spreadsheet contains rows spanning multiple years, fetch all
years concurrently rather than sequentially. Each year's batch queue
runs independently, and results are merged into the shared API cache
before Pass 2 begins. This will reduce total runtime proportionally to
the number of years present.

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

*Last updated: 2026-04-06*
*Contact: Jeremiah Maudlin | jmaudlin@omegaim.com | +1 (641) 660-6367*
