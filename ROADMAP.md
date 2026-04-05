# Pharmacy Tools Add-In — Development Roadmap

## Current Version: 2.3

---

## Version 2.3 — Completed

### 1. Full Dataset Manager (replaces Manage Datasets wizard)
Replaced the 4-step MsgBox wizard with a proper VBA UserForm showing all
configured years (built-in and custom) in a sortable grid.  Users can:
- View all years with Dataset ID, Distribution UUID, and entry type
- Edit any entry (built-in or custom) — overrides stored in registry
- Delete custom entries; reset built-in overrides to factory defaults
- Add new years via the existing 4-step guided wizard (preserved inline)
Built-in entries display as "Built-in"; overridden built-ins as "Built-in*";
user-added years as "Custom".

### 2. Row Colour Reset on Re-Lookup
At the start of Pass 2, all data-row interior colours are reset to xlNone
before writing.  Salmon-red is then re-applied only to rows that fail the
*current* run, so repeated lookups never accumulate stale highlights.

### 3. About Screen Updated to v2.3
Describes both NADAC Lookup and Manage Datasets tools.

---

## Version 2.4 — Planned

### 1. Multi-Year Batch Optimization
Fetch all calendar years concurrently rather than sequentially.  Each
year's batch queue runs independently; results merge into the shared API
cache before Pass 2 begins.  Expected runtime reduction proportional to
the number of distinct years in the spreadsheet.

### 2. Per-Row Status Column
Add a "NADAC Status" column (inserted after NADAC Effective Date) with
standardized codes:
- `OK`            — rate found and written
- `NO_RATE`       — NDC found but no rate on or before the dispense date
- `NOT_FOUND`     — NDC not present in the NADAC dataset
- `INVALID`       — missing or unparseable date or NDC
- `NO_DATASET`    — no NADAC dataset located for the dispense-date year

### 3. Unfound Rows Sheet
Move (or copy) rows where NADAC Status = NOT_FOUND to a separate
"NADAC Not Found" sheet for easier follow-up, rather than leaving them
inline with a red highlight.

---

## Version 2.5 — Planned

### 1. Margin Column
**Calculation:**
  Margin = Paid Amount − (NADAC Per Unit × Quantity)

**Requirements:**
- Auto-detect or prompt for the Paid Amount column and Quantity column
  (same detection pattern as NDC/Date columns)
- Insert a "Margin" column immediately after NADAC Effective Date
- Format as currency ($#,##0.00)
- Negative margins highlighted in light orange to flag below-NADAC
  reimbursements
- Recalculates on every lookup run
- If Quantity column not found, prompt the user to identify it

---
