/**
 * SENIOR CARE CENTER — ROSTER PLANNING TOOL
 * ==========================================
 * Sheets required (auto-created if missing): "Staff Settings", "Roster", "Dashboard"
 *
 * Menu: Center Roster > Generate Optimal Roster / Refresh Dashboard / Initialize Sheet Templates
 */

// ----------------------------- CONFIG -----------------------------

const SHEET_STAFF = 'Staff Settings';
const SHEET_ROSTER = 'Roster';
const SHEET_DASHBOARD = 'Dashboard';

const ROLES = ['Nurse', 'Care Staff', 'Kitchen', 'Escort'];

const SHIFT_START_MIN = 7 * 60;   // 07:00
const SHIFT_END_MIN = 16 * 60;    // 16:00
const BLOCK_MIN = 30;             // 30-minute grid
const ROSTER_FIRST_TIME_COL = 3;  // Column C

// The four permitted lunch start times, per spec. Each lunch is exactly 60 minutes
// (two consecutive 30-minute blocks). Window: 11:30 start (earliest) .. 14:00 end (latest),
// which satisfies "must end by 2:30 PM".
const LUNCH_SLOTS = [
  { id: '11:30', start: 11 * 60 + 30 },
  { id: '12:00', start: 12 * 60 },
  { id: '12:30', start: 12 * 60 + 30 },
  { id: '13:00', start: 13 * 60 }
];
const LUNCH_DURATION_MIN = 60;

// Facility-wide rule: no two employees, of any role, may ever start lunch in the same
// slot. Each of the 4 slots is used by at most one person, so at most 4 employees total
// can be scheduled for lunch across the whole center.
const MAX_TOTAL_LUNCH_CAPACITY = LUNCH_SLOTS.length; // 4

const COLOR_CHARCOAL = '#263238';
const COLOR_CHARCOAL_TEXT = '#FFFFFF';
const COLOR_WORK_BG = '#DCF0DC';
const COLOR_WORK_TEXT = '#256029';
const COLOR_LUNCH_BG = '#FFE7CC';
const COLOR_LUNCH_TEXT = '#B75E00';
const COLOR_OK = '#DCF0DC';
const COLOR_WARNING = '#FFF3CD';
const COLOR_ERROR = '#FADBD8';

/** Thrown for constraint violations that must halt generation before the sheet is touched. */
class RosterError extends Error {}

// ----------------------------- MENU -----------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Center Roster')
    .addItem('Generate Optimal Roster', 'generateOptimalRoster')
    .addItem('Refresh Dashboard', 'refreshDashboard')
    .addSeparator()
    .addItem('Initialize Sheet Templates', 'initializeSheetTemplates')
    .addToUi();
}

// ----------------------------- MAIN ENTRY POINTS -----------------------------

function generateOptimalRoster() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureTemplateSheets(ss);

    const staffSheet = ss.getSheetByName(SHEET_STAFF);
    const { employees, warnings: readWarnings } = readActiveStaff(staffSheet);

    if (employees.length === 0) {
      ui.alert(
        'No Active Employees',
        'No active employees were found in "Staff Settings". Add staff and mark "Is Active" as TRUE before generating a roster.',
        ui.ButtonSet.OK
      );
      return;
    }

    try {
      assignAllLunches(employees);
    } catch (err) {
      if (err instanceof RosterError) {
        ui.alert('Cannot Generate Roster — Constraint Violation', err.message, ui.ButtonSet.OK);
        return; // Abort before touching the Roster sheet.
      }
      throw err;
    }

    const rosterSheet = ss.getSheetByName(SHEET_ROSTER);
    writeRosterSheet(rosterSheet, employees);
    formatRosterSheet(rosterSheet, employees.length);

    refreshDashboardInternal(ss, readWarnings);

    ui.alert(
      'Roster Generated',
      `Roster successfully generated for ${employees.length} active employee(s).\nSee the Dashboard sheet for coverage metrics and validation warnings.`,
      ui.ButtonSet.OK
    );
  } catch (err) {
    Logger.log(err && err.stack ? err.stack : err);
    SpreadsheetApp.getUi().alert(
      'Unexpected Error',
      'Something went wrong while generating the roster:\n' + (err && err.message ? err.message : err),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/** Re-validates and redraws the Dashboard from whatever is currently in the Roster sheet,
 *  without touching Staff Settings or Roster. Useful after manual edits to the roster grid. */
function refreshDashboard() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureTemplateSheets(ss);
    refreshDashboardInternal(ss, []);
    ui.alert('Dashboard Refreshed', 'Dashboard metrics and validation warnings have been updated.', ui.ButtonSet.OK);
  } catch (err) {
    Logger.log(err && err.stack ? err.stack : err);
    ui.alert('Unexpected Error', 'Could not refresh the dashboard:\n' + (err && err.message ? err.message : err), ui.ButtonSet.OK);
  }
}

function initializeSheetTemplates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureTemplateSheets(ss);
  SpreadsheetApp.getUi().alert('Sheet templates are ready. Fill in "Staff Settings", then run Center Roster > Generate Optimal Roster.');
}

// ----------------------------- SHEET SETUP -----------------------------

function ensureTemplateSheets(ss) {
  let staffSheet = ss.getSheetByName(SHEET_STAFF);
  if (!staffSheet) {
    staffSheet = ss.insertSheet(SHEET_STAFF);
  }
  if (staffSheet.getLastRow() === 0) {
    staffSheet.getRange(1, 1, 1, 3).setValues([['Employee Name', 'Role', 'Is Active']]);
    staffSheet.getRange(2, 1, 3, 3).setValues([
      ['Jane Smith', 'Nurse', true],
      ['Alex Chen', 'Care Staff', true],
      ['Maria Lopez', 'Kitchen', true]
    ]);
  }
  formatStaffSettingsSheet(staffSheet);

  let rosterSheet = ss.getSheetByName(SHEET_ROSTER);
  if (!rosterSheet) {
    rosterSheet = ss.insertSheet(SHEET_ROSTER);
  }
  if (rosterSheet.getLastRow() === 0) {
    writeRosterHeaderOnly(rosterSheet);
  }

  let dashboardSheet = ss.getSheetByName(SHEET_DASHBOARD);
  if (!dashboardSheet) {
    dashboardSheet = ss.insertSheet(SHEET_DASHBOARD);
  }
}

function formatStaffSettingsSheet(sheet) {
  sheet.getRange(1, 1, 1, 3)
    .setBackground(COLOR_CHARCOAL)
    .setFontColor(COLOR_CHARCOAL_TEXT)
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, 1, 200);
  sheet.setColumnWidths(2, 1, 140);
  sheet.setColumnWidths(3, 1, 100);

  const maxRows = Math.max(sheet.getMaxRows(), 200);
  if (sheet.getMaxRows() < maxRows) sheet.insertRowsAfter(sheet.getMaxRows(), maxRows - sheet.getMaxRows());

  // Role dropdown
  const roleRange = sheet.getRange(2, 2, maxRows - 1, 1);
  const roleRule = SpreadsheetApp.newDataValidation().requireValueInList(ROLES, true).setAllowInvalid(false).build();
  roleRange.setDataValidation(roleRule);

  // Is Active checkbox
  const activeRange = sheet.getRange(2, 3, maxRows - 1, 1);
  activeRange.insertCheckboxes();
}

function writeRosterHeaderOnly(sheet) {
  const headers = buildRosterHeaders();
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function buildRosterHeaders() {
  const headers = ['Employee Name', 'Role'];
  for (let t = SHIFT_START_MIN; t < SHIFT_END_MIN; t += BLOCK_MIN) {
    headers.push(minutesToLabel(t));
  }
  return headers;
}

// ----------------------------- READ STAFF -----------------------------

function readActiveStaff(staffSheet) {
  const lastRow = staffSheet.getLastRow();
  const warnings = [];
  if (lastRow < 2) return { employees: [], warnings };

  const data = staffSheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const employees = [];

  data.forEach((row, i) => {
    const rawName = row[0];
    const rawRole = row[1];
    const rawActive = row[2];
    const rowNum = i + 2;

    const name = (rawName || '').toString().trim();
    if (!name) return; // skip blank rows silently

    const isActive = rawActive === true || /^(true|yes|active)$/i.test((rawActive || '').toString().trim());
    if (!isActive) return;

    const normalizedRole = ROLES.find(r => r.toLowerCase() === (rawRole || '').toString().trim().toLowerCase());
    if (!normalizedRole) {
      warnings.push(`[WARNING] Row ${rowNum} ("${name}") has an unrecognized role "${rawRole}" and was excluded from the roster. Valid roles: ${ROLES.join(', ')}.`);
      return;
    }

    employees.push({ name, role: normalizedRole });
  });

  return { employees, warnings };
}

// ----------------------------- LUNCH ASSIGNMENT -----------------------------

/**
 * Assigns every active employee (across all roles) to one of the 4 lunch slots so that no
 * two employees ever start lunch in the same slot. Since only 4 slots exist, at most 4
 * employees total can be scheduled — this is a hard facility-wide cap, not a per-role one.
 * Mutates each employee object with lunchStart / lunchSlotLabel.
 * Throws RosterError if there are more than MAX_TOTAL_LUNCH_CAPACITY active employees.
 */
function assignAllLunches(employees) {
  const n = employees.length;
  if (n === 0) return;

  if (n > MAX_TOTAL_LUNCH_CAPACITY) {
    throw new RosterError(
      `There are ${n} active employees, but only ${MAX_TOTAL_LUNCH_CAPACITY} lunch slots exist and each slot may be ` +
      `used by only one person — no two staff may start lunch at the same time. Reduce the active headcount to ` +
      `${MAX_TOTAL_LUNCH_CAPACITY} or fewer, mark extra staff inactive, or add more lunch slots, then try again.`
    );
  }

  // Sort by role priority then name so the solver's placement is deterministic and roles
  // that matter most for coverage (Nurse first, per spec) are considered first when ties occur.
  const sorted = employees.slice().sort((a, b) => {
    if (a.role !== b.role) return ROLES.indexOf(a.role) - ROLES.indexOf(b.role);
    return a.name.localeCompare(b.name);
  });

  const assignment = solveLunchSlots(sorted);
  sorted.forEach((emp, i) => {
    emp.lunchStart = assignment[i].start;
    emp.lunchSlotLabel = assignment[i].id;
  });
}

/**
 * Finds the assignment of employees to distinct lunch slots that minimizes same-role
 * adjacent-slot pairs. Adjacent slots (e.g. 12:00 and 12:30) overlap by 30 minutes, so if
 * they're the only two people in that role, the role goes fully unstaffed for that half
 * hour. The search space is tiny (at most 4 employees over 4 slots, <=24 permutations), so
 * a small backtracking search finds the true optimum instead of relying on a fixed greedy
 * fill order, which can miss a zero-violation arrangement that does exist (e.g. 2 Nurses +
 * 2 Care Staff can always be split into two non-overlapping pairs, but a naive sequential
 * fill can accidentally leave the second role with an adjacent, overlapping pair).
 */
function solveLunchSlots(employees) {
  const n = employees.length;
  const used = new Array(LUNCH_SLOTS.length).fill(false);
  const current = new Array(n);
  let best = null;
  let bestViolations = Infinity;

  const isAdjacent = (a, b) => Math.abs(a - b) === 1;

  function recurse(i, violations) {
    if (violations >= bestViolations) return; // can't beat the best solution found so far
    if (i === n) {
      bestViolations = violations;
      best = current.slice();
      return;
    }
    for (let s = 0; s < LUNCH_SLOTS.length; s++) {
      if (used[s]) continue;
      let added = 0;
      for (let j = 0; j < i; j++) {
        if (employees[j].role === employees[i].role && isAdjacent(current[j], s)) added++;
      }
      used[s] = true;
      current[i] = s;
      recurse(i + 1, violations + added);
      used[s] = false;
      if (bestViolations === 0) return; // already optimal, stop searching
    }
  }

  recurse(0, 0);
  return best.map(slotIdx => LUNCH_SLOTS[slotIdx]);
}

// ----------------------------- ROSTER SHEET -----------------------------

function writeRosterSheet(sheet, employees) {
  sheet.clear();

  const sorted = employees.slice().sort((a, b) => {
    if (a.role !== b.role) return ROLES.indexOf(a.role) - ROLES.indexOf(b.role);
    return a.name.localeCompare(b.name);
  });

  const headers = buildRosterHeaders();
  const blockStarts = [];
  for (let t = SHIFT_START_MIN; t < SHIFT_END_MIN; t += BLOCK_MIN) blockStarts.push(t);

  const rows = sorted.map(emp => {
    const cells = blockStarts.map(t => {
      const onLunch = emp.lunchStart !== undefined && t >= emp.lunchStart && t < emp.lunchStart + LUNCH_DURATION_MIN;
      return onLunch ? 'LUNCH' : 'WORK';
    });
    return [emp.name, emp.role, ...cells];
  });

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}

function formatRosterSheet(sheet, numEmployees) {
  const totalCols = buildRosterHeaders().length;

  // Header styling
  const headerRange = sheet.getRange(1, 1, 1, totalCols);
  headerRange
    .setBackground(COLOR_CHARCOAL)
    .setFontColor(COLOR_CHARCOAL_TEXT)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);

  sheet.setColumnWidth(1, 170);
  sheet.setColumnWidth(2, 110);
  for (let c = ROSTER_FIRST_TIME_COL; c <= totalCols; c++) {
    sheet.setColumnWidth(c, 55);
  }

  if (numEmployees > 0) {
    const dataRange = sheet.getRange(1, 1, numEmployees + 1, totalCols);
    dataRange.setBorder(true, true, true, true, true, true, '#D0D0D0', SpreadsheetApp.BorderStyle.SOLID);

    sheet.getRange(2, 1, numEmployees, 1).setFontWeight('bold');
    sheet.getRange(2, 1, numEmployees, totalCols).setHorizontalAlignment('center');
    sheet.getRange(2, 1, numEmployees, 1).setHorizontalAlignment('left');

    const timeRange = sheet.getRange(2, ROSTER_FIRST_TIME_COL, numEmployees, totalCols - ROSTER_FIRST_TIME_COL + 1);
    const workRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('WORK')
      .setBackground(COLOR_WORK_BG)
      .setFontColor(COLOR_WORK_TEXT)
      .setRanges([timeRange])
      .build();
    const lunchRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('LUNCH')
      .setBackground(COLOR_LUNCH_BG)
      .setFontColor(COLOR_LUNCH_TEXT)
      .setRanges([timeRange])
      .build();
    sheet.setConditionalFormatRules([workRule, lunchRule]);
  }

  sheet.setRowHeight(1, 28);
}

// ----------------------------- DASHBOARD -----------------------------

function refreshDashboardInternal(ss, extraNotes) {
  const rosterSheet = ss.getSheetByName(SHEET_ROSTER);
  const dashboardSheet = ss.getSheetByName(SHEET_DASHBOARD);

  const validation = scanRosterForValidation(rosterSheet);
  const allWarnings = (extraNotes || []).concat(validation.warnings);

  writeDashboard(dashboardSheet, validation, allWarnings);
}

const LUNCH_BLOCK_LABELS = ['11:30', '12:00', '12:30', '13:00', '13:30'];

/** Reads the Roster sheet's WORK/LUNCH grid directly (single source of truth) and
 *  recomputes per-role, per-block lunch coverage plus rule violations. */
function scanRosterForValidation(rosterSheet) {
  const lastRow = rosterSheet.getLastRow();
  const lastCol = rosterSheet.getLastColumn();
  const result = { roleTotals: {}, coverage: {}, warnings: [] };
  ROLES.forEach(r => { result.roleTotals[r] = 0; result.coverage[r] = {}; LUNCH_BLOCK_LABELS.forEach(b => result.coverage[r][b] = 0); });

  if (lastRow < 2 || lastCol < ROSTER_FIRST_TIME_COL) {
    result.warnings.push('[INFO] Roster has not been generated yet. Run "Generate Optimal Roster" first.');
    return result;
  }

  const values = rosterSheet.getRange(1, 1, lastRow, lastCol).getValues();
  const header = values[0];
  const blockColIndex = {}; // label -> 0-based column index in `values` rows
  LUNCH_BLOCK_LABELS.forEach(label => {
    const idx = header.findIndex(h => h === label);
    if (idx >= 0) blockColIndex[label] = idx;
  });

  // label -> names of employees whose lunch appears to START in that slot (first LUNCH
  // block encountered, chronologically), used for the "no two people start at once" rule.
  const startedAt = {};
  LUNCH_BLOCK_LABELS.forEach(l => (startedAt[l] = []));

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const name = row[0];
    const role = row[1];
    if (!name || !ROLES.includes(role)) continue;
    result.roleTotals[role]++;

    let inferredStart = null;
    LUNCH_BLOCK_LABELS.forEach(label => {
      const colIdx = blockColIndex[label];
      if (colIdx !== undefined && row[colIdx] === 'LUNCH') {
        result.coverage[role][label]++;
        if (inferredStart === null) inferredStart = label;
      }
    });
    if (inferredStart !== null) startedAt[inferredStart].push(name);
  }

  LUNCH_BLOCK_LABELS.forEach(label => {
    if (startedAt[label].length > 1) {
      result.warnings.push(
        `[ERROR] ${startedAt[label].length} employees are all starting lunch at ${label} (${startedAt[label].join(', ')}) ` +
        `— only one person may start lunch at a given time. Reassign one of them or regenerate the roster.`
      );
    }
  });

  ROLES.forEach(role => {
    const total = result.roleTotals[role];
    if (total === 0) return;
    LUNCH_BLOCK_LABELS.forEach(label => {
      const onLunch = result.coverage[role][label];
      const working = total - onLunch;

      if (working === 0) {
        if (total >= 2) {
          result.warnings.push(
            `[ERROR] Role "${role}" is completely unstaffed during the ${label} half-hour — minimum staffing floor ` +
            `violated. With ${total} employees in this role, the roster generator should never produce this on its ` +
            `own; the roster was likely edited manually — consider regenerating.`
          );
        } else {
          result.warnings.push(`[INFO] Role "${role}" has only one staff member, so it is unstaffed during their own lunch (${label}). This is expected for single-person roles.`);
        }
      }
    });
  });

  return result;
}

function writeDashboard(sheet, validation, warnings) {
  // Break apart any merged cells from a previous run before clearing — clear() does not
  // reliably unmerge, and the row layout below shifts as the warning count changes.
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.clear();
  sheet.clearConditionalFormatRules();

  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 140);
  for (let c = 3; c <= 8; c++) sheet.setColumnWidth(c, 90);

  sheet.getRange(1, 1, 1, 6).merge()
    .setValue('Senior Care Center — Roster Dashboard')
    .setBackground(COLOR_CHARCOAL)
    .setFontColor(COLOR_CHARCOAL_TEXT)
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('left');
  sheet.setRowHeight(1, 32);

  // Role headcount (live formulas against Staff Settings)
  sheet.getRange(3, 1, 1, 2).setValues([['Role', 'Active Staff']])
    .setBackground(COLOR_CHARCOAL).setFontColor(COLOR_CHARCOAL_TEXT).setFontWeight('bold');

  ROLES.forEach((role, i) => {
    const row = 4 + i;
    sheet.getRange(row, 1).setValue(role);
    sheet.getRange(row, 2).setFormula(
      `=COUNTIFS('${SHEET_STAFF}'!B:B,"${role}",'${SHEET_STAFF}'!C:C,TRUE)`
    );
  });
  const totalRow = 4 + ROLES.length;
  sheet.getRange(totalRow, 1).setValue('Total Active Staff').setFontWeight('bold');
  sheet.getRange(totalRow, 2).setFormula(`=COUNTIF('${SHEET_STAFF}'!C:C,TRUE)`).setFontWeight('bold');

  // Status + timestamp
  const statusLevel = worstLevel(warnings);
  const statusRow = totalRow + 2;
  sheet.getRange(statusRow, 1).setValue('Overall Roster Status:').setFontWeight('bold');
  const statusCell = sheet.getRange(statusRow, 2).setValue(statusLevel);
  statusCell.setBackground(statusLevel === 'ERROR' ? COLOR_ERROR : statusLevel === 'WARNING' ? COLOR_WARNING : COLOR_OK);
  statusCell.setFontWeight('bold');

  sheet.getRange(statusRow + 1, 1).setValue('Last Refreshed:').setFontWeight('bold');
  sheet.getRange(statusRow + 1, 2).setValue(new Date());

  // Warnings list
  const warnRow = statusRow + 3;
  sheet.getRange(warnRow, 1, 1, 6).merge().setValue('Validation Warnings')
    .setBackground(COLOR_CHARCOAL).setFontColor(COLOR_CHARCOAL_TEXT).setFontWeight('bold');

  let nextRow = warnRow + 1;
  if (warnings.length === 0) {
    sheet.getRange(nextRow, 1, 1, 6).merge().setValue('No issues detected.').setBackground(COLOR_OK);
    nextRow++;
  } else {
    warnings.forEach(w => {
      const level = w.match(/^\[(ERROR|WARNING|INFO)\]/);
      const bg = level && level[1] === 'ERROR' ? COLOR_ERROR : level && level[1] === 'WARNING' ? COLOR_WARNING : '#EEEEEE';
      sheet.getRange(nextRow, 1, 1, 6).merge().setValue(w).setBackground(bg).setWrap(true);
      nextRow++;
    });
  }

  // Lunch coverage mini table
  nextRow += 1;
  sheet.getRange(nextRow, 1, 1, 6).merge().setValue('Lunch-Hour Coverage (headcount on lunch per 30-min block)')
    .setBackground(COLOR_CHARCOAL).setFontColor(COLOR_CHARCOAL_TEXT).setFontWeight('bold');
  nextRow++;

  const tableHeader = ['Role', ...LUNCH_BLOCK_LABELS];
  sheet.getRange(nextRow, 1, 1, tableHeader.length).setValues([tableHeader]).setFontWeight('bold').setBackground('#ECEFF1');
  nextRow++;

  ROLES.forEach(role => {
    if (!validation.roleTotals[role]) return;
    const row = [role, ...LUNCH_BLOCK_LABELS.map(l => validation.coverage[role][l])];
    sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);
    nextRow++;
  });

  sheet.getRange(1, 1, nextRow, 8).setBorder(true, true, true, true, false, false, '#D0D0D0', SpreadsheetApp.BorderStyle.SOLID);
}

function worstLevel(warnings) {
  if (warnings.some(w => w.startsWith('[ERROR]'))) return 'ERROR';
  if (warnings.some(w => w.startsWith('[WARNING]'))) return 'WARNING';
  return 'OK';
}

// ----------------------------- UTIL -----------------------------

function minutesToLabel(totalMinutes) {
  const h = Math.floor(totalMinutes / 60);
  const m = totalMinutes % 60;
  return (h < 10 ? '0' : '') + h + ':' + (m < 10 ? '0' : '') + m;
}
