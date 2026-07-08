# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repo is

A Senior Care Center roster/lunch-scheduling tool, implemented **three times independently** on three
different platforms, each a fully self-contained deliverable with no shared code or build system:

- **`Code.gs`** — Google Apps Script, bound to a Google Sheet ("Staff Settings" / "Roster" / "Dashboard" tabs).
- **`RosterModule.bas`** + **`SeniorCareRoster.xlsm`** — VBA macro module and its companion Excel workbook
  (same three-sheet structure plus a "Home" tab with macro buttons).
- **`web/index.html`** — a single self-contained HTML/CSS/JS file (no build step, no dependencies), storing
  state in `localStorage`. Also published as a Claude Artifact.

There is no package manager, no bundler, and no shared module between the three — **the same scheduling
logic and validation rules are hand-duplicated in all three files**. Any change to the business rules
(shift times, lunch-slot rules, role list, validation messages) must be made in all three places to keep
them consistent. `SeniorCareRoster.xlsm` intentionally ships with **no embedded VBA project** — the user
pastes `RosterModule.bas` in manually via Alt+F11. Do not embed macros back into that file.

## Running / developing each version

There is no build, lint, or automated test command in this repo. To work on each version:

- **Google Sheets**: paste `Code.gs` into the Apps Script editor bound to a spreadsheet; it adds a
  "Center Roster" menu (`onOpen`) with Generate Optimal Roster / Refresh Dashboard / Initialize Sheet
  Templates.
- **Excel**: open `SeniorCareRoster.xlsm`, `Alt+F11` → Insert Module → paste `RosterModule.bas`, save as
  `.xlsm`. Buttons on the "Home" sheet call `GenerateOptimalRoster` / `RefreshDashboard` /
  `InitializeSheetTemplates`.
- **Web**: open `web/index.html` directly in a browser — no server needed.

### Verifying changes (no test suite exists — verify by direct execution)

- **VBA**: never test against `SeniorCareRoster.xlsm` directly. Copy it to a scratch file, temporarily set
  `HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM = 1` (needed for `VBProject.VBComponents`
  access), import a **renamed** copy of the module via Excel COM automation (`New-Object -ComObject
  Excel.Application`, `Workbooks.Open`, `VBComponents.Import`, `Application.Run`), then revert the
  registry key and delete the scratch file/kill the Excel process afterward. Renaming the imported module
  matters: VBA is case-insensitive, so a local variable like `Dim roles: roles = Roles()` silently shadows
  a same-named module function and throws "Subscript out of range" — this is why the module's lookup
  functions are named `RoleNames()` / `LunchFillOrder()`-style rather than matching common local var names.
  Also replace `MsgBox` calls with a logging stub when testing headlessly, since a real `MsgBox` blocks
  forever with no user present to click it.
- **Apps Script**: no local runner — test inside an actual bound Google Sheet.
- **Web**: the pure logic (`buildRoster`, `solveLunchSlots`, `scanValidation`) has no DOM dependency and
  can be copy-extracted into a throwaway Node script for fast checks. For layout/responsive checks, headless
  Edge/Chrome screenshots work, but note: on Windows, headless Chromium enforces a minimum window width
  (~500px) that `--window-size` cannot go below — you cannot screenshot true small-phone widths directly.
  Verify narrow-viewport CSS via an injected script reading `getComputedStyle`/`getBoundingClientRect`
  instead of trusting a cropped screenshot.

Also watch for orphaned `EXCEL.EXE`/`msedge.exe` processes from interrupted automation — kill them before
re-running, and re-verify `Workbook.HasVBProject === False` on `SeniorCareRoster.xlsm` after any VBA testing
session, since a crashed COM session can silently leave VBA components embedded in the real file.

## Core domain rules (must stay identical across all three implementations)

- Shift is 07:00–16:00 in 30-minute blocks (18 blocks). Roles are fixed: `Nurse`, `Care Staff`, `Kitchen`,
  `Escort` — order matters (Nurse first) as the tie-break/priority order when sorting employees.
- Every active employee gets exactly one 60-minute lunch, starting at one of 4 fixed slots: `11:30`,
  `12:00`, `12:30`, `13:00`.
- **Facility-wide constraint**: no two employees, of any role, may start lunch in the same slot. This is a
  hard cap of 4 active employees total (not per-role) — exceeding it throws a `RosterError`
  (`assignAllLunches` / `AssignAllLunches` / `buildRoster`) and generation aborts *before* touching the
  Roster sheet/grid.
- **The scheduling algorithm** (`solveLunchSlots` / `SolveLunchSlotsRecurse`) is a small backtracking
  search, not a fixed greedy fill order — the search space is tiny (≤4 employees over 4 slots, ≤24
  permutations), so it finds the true optimum. It minimizes *same-role adjacent-slot pairs*: adjacent slots
  (e.g. 12:00 & 12:30) overlap 30 minutes, so if they're the only two people in a role, that role goes
  fully unstaffed for that half hour. A fixed sequential fill order can miss a zero-violation arrangement
  that does exist (e.g. 2 Nurses + 2 Care Staff can always split into two non-overlapping pairs); the
  backtracking search never does.
- Employees are sorted by role-priority-then-name before solving, for deterministic, reproducible output.
- **Validation always re-derives state from the actual grid**, never from generation-time memory
  (`scanRosterForValidation` / `ScanRosterForValidation` / `scanValidation`) — this is what allows manual
  edits (spreadsheet cell edits, or the web app's clickable lunch cells) to be re-validated via a
  "Refresh Dashboard" action, using the grid as the single source of truth.
- Validation reports two independent things per role/block:
  1. **Duplicate lunch start** (ERROR, zero tolerance) — two+ employees inferred to start lunch at the
     same slot. Should never happen after correct generation; only arises from a manual edit.
  2. **Floor violation** — a role fully unstaffed for some half-hour. ERROR if the role has 2+ members
     (the solver formally guarantees this can't happen from correct generation, so seeing it means a
     manual edit broke something); INFO if it's a 1-person role (an expected, unavoidable gap while that
     one person eats).

## Platform-specific gotchas

- **Writing "HH:MM" strings into spreadsheet cells** (Google Sheets or Excel) without first forcing text
  format causes silent auto-conversion to a time serial number — this breaks any later string comparison
  against header labels. Both `Code.gs` and `RosterModule.bas` explicitly set the cell format to text
  (`setNumberFormat`/`NumberFormat = "@"`) before writing these headers; preserve this when touching that
  code.
- **Google Sheets Dashboard rebuild**: `Range.clear()` does not reliably un-merge previously merged cells,
  and the Dashboard's row layout shifts every time the warning count changes — `breakApart()`/`UnMerge`
  must run before `clear()` on every rebuild, or a merge-conflict error appears on the second run.
- **Web app**: single HTML file, no backend, state in `localStorage` (key `senior_care_roster_v1`). Layout
  uses a CSS Grid shell (`.app`) with a left sidebar on desktop and a fixed bottom tab bar on mobile
  (`@media (max-width: 880px)`); grid items need explicit `min-width: 0` or wide inner content (e.g. a
  table) can blow out the whole page width. The global success/error banner (`#globalBanner`) is
  deliberately placed outside the per-view `<section>` elements so it stays visible across a `switchView()`
  call (e.g. right after Generate Roster auto-switches to the Roster tab).
