# roster

Senior Care Center roster and lunch-scheduling planner, available in three independent, self-contained
versions:

- **Google Sheets** — [`Code.gs`](Code.gs), Google Apps Script bound to a spreadsheet.
- **Excel** — [`RosterModule.bas`](RosterModule.bas) (VBA macro, pasted in manually) plus the companion
  workbook [`SeniorCareRoster.xlsm`](SeniorCareRoster.xlsm).
- **Web** — [`web/index.html`](web/index.html), a single self-contained HTML/CSS/JS file with no build
  step or dependencies (state persists in the browser's local storage).

Each version generates a shift roster for four roles (Nurse, Care Staff, Kitchen, Escort) across a
07:00–16:00 shift, assigns each active employee a one-hour lunch from four fixed slots
(11:30 / 12:00 / 12:30 / 13:00) with no two employees starting lunch at the same time, and validates the
resulting grid for staffing-floor and overlap issues.

See [`CLAUDE.md`](CLAUDE.md) for architecture notes and how to develop/verify changes to each version.
