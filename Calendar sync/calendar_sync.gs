/**
 * Calendar Sync for Google Sheets (Google Apps Script)
 *
 * Builds one sheet per month with a calendar grid and a details table, and keeps
 * both in sync with one or more Google Calendars WITHOUT overwriting manual edits.
 *
 * Highlights
 *  - Incremental sync: adds new events, refreshes date/title/time of existing ones
 *  - Stable matching by event ID + date, so recurring events stay distinct
 *  - Never deletes rows automatically; flags events removed from the calendar
 *  - Custom menu, dropdowns (data validation) and conditional formatting
 *  - Handles daylight saving time via a named time zone
 *
 * Setup: create two calendars named as in CALENDAR_CONFIGS below (with a few
 * demo events), paste this file into Extensions > Apps Script, reload the sheet,
 * then use the "Calendar Sync" menu.
 */
// ============================================================
//  CONFIG
// ============================================================
const CALENDAR_CONFIGS = [
  { name: "Demo Training Calendar", label: "🥋 Training" },
  { name: "Demo Class Calendar", label: "📖 Classes" }
];

const START_MONTH = 9;   // September
const START_YEAR = 2026;
const END_MONTH = 12;    // December
const END_YEAR = 2026;

const MAX_DETAIL_ROWS = 100; // max events per month in the details table
const DETAIL_COLS = 10;      // A..J

// Time zone used for the Time column (named zone, follows daylight saving automatically)
const DISPLAY_TZ = "Europe/Helsinki";

// Zero-based column positions inside a details row
const IDX = { DATE: 0, TITLE: 1, TIME: 2, TODO: 3, DETAILS: 4, STATUS: 5, STATUS_DETAILS: 6, SOURCE: 7, SYNC: 8, KEY: 9 };

const HEADERS = ['📅 Date', '📚 Class Name', '⏰ Time', '📝 TO DO', '✍️ Details', '✅ Status', '📊 Status Details', '📌 Source', '🔄 Sync', 'Event Key'];
const MONTH_EMOJIS = ['🎉', '❄️', '🌸', '🌼', '🌻', '☀️', '🏖️', '🍂', '🍁', '🎃', '🦃', '🎄'];

// Adds a "Calendar Sync" menu to the sheet
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Calendar Sync')
    .addItem('Sync new events now', 'createMonthlyCalendarWithEvents')
    .addItem('Clean up deleted events', 'cleanupDeletedEvents')
    .addToUi();
}

// ============================================================
//  MAIN: incremental sync (never clears manual edits)
// ============================================================
function createMonthlyCalendarWithEvents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();

  const calendars = CALENDAR_CONFIGS.map(cfg => {
    const found = CalendarApp.getCalendarsByName(cfg.name);
    if (found.length === 0) {
      throw new Error("Calendar with name '" + cfg.name + "' not found!");
    }
    return { calendar: found[0], label: cfg.label };
  });

  let year = START_YEAR;
  let month = START_MONTH;

  while (year < END_YEAR || (year === END_YEAR && month <= END_MONTH)) {
    const layout = getLayout(year, month);
    const sheetName = `${getMonthName(month)} ${year}`;

    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      buildNewSheet(sheet, year, month, layout);
    }

    ensureDetailsHeader(sheet, layout);
    refreshGridRule(sheet, year, month, layout);
    syncEvents(sheet, year, month, layout, calendars, tz);
    renderGrid(sheet, year, month, layout, tz);

    month++;
    if (month > 12) {
      month = 1;
      year++;
    }
  }
}

// ============================================================
//  LAYOUT (same row positions as the original script)
// ============================================================
function getLayout(year, month) {
  const firstDay = new Date(year, month - 1, 1);
  const daysInMonth = new Date(year, month, 0).getDate();
  const startCol = firstDay.getDay() + 1;

  let row = 3;
  let col = startCol;
  for (let d = 1; d <= daysInMonth; d++) {
    col++;
    if (col > 7) { col = 1; row++; }
  }
  const gridLastRow = (col === 1) ? row - 1 : row;

  return { daysInMonth, startCol, gridLastRow, detailsHeaderRow: row + 2 };
}

// ============================================================
//  ONE-TIME BUILD of a brand-new month sheet (styling, dropdowns, rules)
// ============================================================
function buildNewSheet(sheet, year, month, layout) {
  // Month title
  sheet.getRange('A1:G1').merge()
    .setValue(`${MONTH_EMOJIS[month - 1]} ${getMonthName(month)} ${year}`)
    .setBackground('#2C3E50')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setFontSize(16);
  sheet.setRowHeight(1, 40);

  // Day-of-week header
  const daysOfWeek = ['☀️ Sun', '🌙 Mon', '🔥 Tue', '🌟 Wed', '🎯 Thu', '🎨 Fri', '🎮 Sat'];
  sheet.getRange(2, 1, 1, 7).setValues([daysOfWeek])
    .setBackground('#4A90E2')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setFontSize(11);

  // Calendar grid cell styling
  for (let d = 1; d <= layout.daysInMonth; d++) {
    const idx = layout.startCol - 1 + (d - 1);
    const r = 3 + Math.floor(idx / 7);
    const c = 1 + (idx % 7);
    sheet.getRange(r, c)
      .setVerticalAlignment('top')
      .setHorizontalAlignment('left')
      .setFontWeight('bold')
      .setWrap(true)
      .setBackground((c === 1 || c === 7) ? '#F5F5F5' : '#FFFFFF')
      .setBorder(true, true, true, true, false, false, '#D3D3D3', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setRowHeight(r, 80);
  }

  // Column widths (shared by grid + details table)
  const widths = [110, 200, 130, 140, 250, 120, 250];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  const firstDataRow = layout.detailsHeaderRow + 1;

  // Dropdowns
  const yesNo = SpreadsheetApp.newDataValidation()
    .requireValueInList(['✅ Yes', '❌ No'], true)
    .setAllowInvalid(false)
    .build();
  const todoRange = sheet.getRange(firstDataRow, IDX.TODO + 1, MAX_DETAIL_ROWS, 1);
  const statusRange = sheet.getRange(firstDataRow, IDX.STATUS + 1, MAX_DETAIL_ROWS, 1);
  todoRange.setDataValidation(yesNo);
  statusRange.setDataValidation(yesNo);

  // Details area basic styling (cols A..G)
  sheet.getRange(firstDataRow, 1, MAX_DETAIL_ROWS, 7)
    .setVerticalAlignment('top')
    .setWrap(true)
    .setFontSize(9);

  // Conditional formatting
  const todoRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('✅ Yes').setBackground('#9B59B6').setFontColor('#FFFFFF')
    .setRanges([todoRange]).build();
  const statusYesRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('✅ Yes').setBackground('#27AE60').setFontColor('#FFFFFF')
    .setRanges([statusRange]).build();
  const statusNoRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('❌ No').setBackground('#E74C3C').setFontColor('#FFFFFF')
    .setRanges([statusRange]).build();
  const pastDateRule = SpreadsheetApp.newConditionalFormatRule()
    .whenDateBefore(SpreadsheetApp.RelativeDate.TODAY)
    .setBackground('#D4EDDA').setFontColor('#155724')
    .setRanges([sheet.getRange(firstDataRow, 1, MAX_DETAIL_ROWS, 1)]).build();

  const calendarGridRange = sheet.getRange(3, 1, layout.gridLastRow - 2, 7);
  const pastCalendarDateRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(getGridFormula(year, month))
    .setBackground('#D4EDDA').setFontColor('#155724')
    .setRanges([calendarGridRange]).build();

  sheet.setConditionalFormatRules([todoRule, statusYesRule, statusNoRule, pastDateRule, pastCalendarDateRule]);
}

// Green only when the cell holds a day AND that day is over (before today).
// The A3<>"" check matters: blank cells would otherwise be read as "day 0" and turn green.
function getGridFormula(year, month) {
  return `=AND(A3<>"", DATE(${year},${month},VALUE(LEFT(A3,FIND(CHAR(10),A3&CHAR(10))-1)))<TODAY())`;
}

// ============================================================
//  Keeps the "past days turn green" rule covering the WHOLE grid.
//  Also repairs sheets made by the original script, whose rule skipped the last week row.
// ============================================================
function refreshGridRule(sheet, year, month, layout) {
  const gridRange = sheet.getRange(3, 1, layout.gridLastRow - 2, 7);

  // Keep every rule except the old grid "past date" rule
  const kept = sheet.getConditionalFormatRules().filter(rule => {
    const cond = rule.getBooleanCondition();
    if (!cond || cond.getCriteriaType() !== SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) return true;
    return String(cond.getCriteriaValues()[0]).indexOf('FIND(CHAR(10)') === -1;
  });

  const gridRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(getGridFormula(year, month))
    .setBackground('#D4EDDA')
    .setFontColor('#155724')
    .setRanges([gridRange])
    .build();

  kept.push(gridRule);
  sheet.setConditionalFormatRules(kept);
}

// ============================================================
//  Details header + extra columns (Source / Sync / hidden Event Key)
//  Safe to run on old sheets too: upgrades them without touching data.
// ============================================================
function ensureDetailsHeader(sheet, layout) {
  const headerRow = layout.detailsHeaderRow;
  const needsSetup = sheet.getRange(headerRow, IDX.KEY + 1).getValue() === '';

  sheet.getRange(headerRow, 1, 1, DETAIL_COLS).setValues([HEADERS])
    .setBackground('#2C3E50')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setFontSize(10);

  if (needsSetup) {
    sheet.setColumnWidth(IDX.SOURCE + 1, 130);
    sheet.setColumnWidth(IDX.SYNC + 1, 170);
    sheet.getRange(headerRow + 1, IDX.SOURCE + 1, MAX_DETAIL_ROWS, 3)
      .setVerticalAlignment('top')
      .setWrap(true)
      .setFontSize(9);
    sheet.hideColumns(IDX.KEY + 1); // hidden column J holds the event key used for matching
  }
}

// ============================================================
//  SYNC: add only NEW events, refresh Date/Class/Time of existing ones.
//  Never touches TO DO, Details, Status, Status Details.
// ============================================================
function syncEvents(sheet, year, month, layout, calendars, tz) {
  const firstDataRow = layout.detailsHeaderRow + 1;
  const startDate = new Date(year, month - 1, 1);
  const endDate = new Date(year, month, 0, 23, 59, 59);

  // 1) Fetch events from all calendars
  const fetched = [];
  calendars.forEach(({ calendar, label }) => {
    calendar.getEvents(startDate, endDate).forEach(event => {
      fetched.push({ event, label, key: makeKey(event, tz) });
    });
  });
  fetched.sort((a, b) => a.event.getStartTime() - b.event.getStartTime());

  // 2) Read what's already in the sheet
  const data = sheet.getRange(firstDataRow, 1, MAX_DETAIL_ROWS, DETAIL_COLS).getValues();
  const byKey = {};
  data.forEach((r, i) => {
    if (isBlankRow(r)) return;
    if (r[IDX.KEY]) byKey[r[IDX.KEY]] = i;
  });

  // 3) Backfill keys for rows created by older versions of the script
  //    (match on same day + same title)
  data.forEach((r, i) => {
    if (isBlankRow(r) || r[IDX.KEY] || !(r[IDX.DATE] instanceof Date)) return;
    const day = Utilities.formatDate(r[IDX.DATE], tz, 'yyyy-MM-dd');
    const match = fetched.find(f =>
      byKey[f.key] === undefined &&
      f.event.getTitle() === r[IDX.TITLE] &&
      Utilities.formatDate(f.event.getStartTime(), tz, 'yyyy-MM-dd') === day
    );
    if (match) {
      r[IDX.KEY] = match.key;
      byKey[match.key] = i;
    }
  });

  // 4) Update existing rows / collect new ones
  const newRows = [];
  const fetchedKeys = {};

  fetched.forEach(f => {
    fetchedKeys[f.key] = true;
    const start = f.event.getStartTime();
    const timeText = formatTime(start) + ' - ' + formatTime(f.event.getEndTime()) + ' ' + getTzLabel(start);

    if (byKey[f.key] !== undefined) {
      const r = data[byKey[f.key]];
      r[IDX.DATE] = start;
      r[IDX.TITLE] = f.event.getTitle();
      r[IDX.TIME] = timeText;
      if (!r[IDX.SOURCE]) r[IDX.SOURCE] = f.label;
      r[IDX.SYNC] = '';
    } else {
      newRows.push([start, f.event.getTitle(), timeText, '❌ No', '', '❌ No', '', f.label, '', f.key]);
    }
  });

  // 5) Flag rows whose event no longer exists in the calendar.
  //    Never removed here - use the "Clean up deleted events" menu command for that.
  data.forEach(r => {
    if (r[IDX.KEY] && !fetchedKeys[r[IDX.KEY]]) {
      r[IDX.SYNC] = '⚠️ Not in calendar anymore';
    }
  });

  // 6) Combine existing (non-blank) rows with newly-fetched events and sort by date
  const existingRows = data.filter(r => !isBlankRow(r));
  const allRows = existingRows.concat(newRows);
  allRows.sort((a, b) => {
    const ta = (a[IDX.DATE] instanceof Date) ? a[IDX.DATE].getTime() : 0;
    const tb = (b[IDX.DATE] instanceof Date) ? b[IDX.DATE].getTime() : 0;
    return ta - tb;
  });

  if (allRows.length > MAX_DETAIL_ROWS) {
    throw new Error('Too many events in ' + sheet.getName() + ' (limit ' + MAX_DETAIL_ROWS + '). Increase MAX_DETAIL_ROWS.');
  }

  rebuildDetailsBlock(sheet, firstDataRow, allRows);
}

// ============================================================
//  MENU COMMAND: remove rows flagged "not in calendar anymore" that are
//  still untouched (TO DO/Status at default, Details/Status Details empty).
//  Run this yourself whenever you want; the regular sync never deletes rows.
// ============================================================
function cleanupDeletedEvents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let year = START_YEAR;
  let month = START_MONTH;
  let totalRemoved = 0;
  const perSheet = [];

  while (year < END_YEAR || (year === END_YEAR && month <= END_MONTH)) {
    const sheetName = `${getMonthName(month)} ${year}`;
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const layout = getLayout(year, month);
      const firstDataRow = layout.detailsHeaderRow + 1;
      const data = sheet.getRange(firstDataRow, 1, MAX_DETAIL_ROWS, DETAIL_COLS).getValues();

      const kept = [];
      let removed = 0;
      data.forEach(r => {
        if (isBlankRow(r)) return;
        const flagged = r[IDX.SYNC] === '⚠️ Not in calendar anymore';
        const untouched = (r[IDX.TODO] === '' || r[IDX.TODO] === '❌ No') &&
                           (r[IDX.STATUS] === '' || r[IDX.STATUS] === '❌ No') &&
                           r[IDX.DETAILS] === '' &&
                           r[IDX.STATUS_DETAILS] === '';
        if (flagged && untouched) {
          removed++;
          return;
        }
        kept.push(r);
      });

      if (removed > 0) {
        rebuildDetailsBlock(sheet, firstDataRow, kept);
        renderGrid(sheet, year, month, layout, Session.getScriptTimeZone());
        totalRemoved += removed;
        perSheet.push(removed + ' in ' + sheetName);
      }
    }
    month++;
    if (month > 12) { month = 1; year++; }
  }

  const msg = totalRemoved > 0
    ? 'Removed ' + totalRemoved + ' row(s): ' + perSheet.join(', ')
    : 'Nothing to remove - no untouched rows are flagged as deleted.';
  SpreadsheetApp.getUi().alert(msg);
}

// ============================================================
//  Clears and rewrites the details block from a plain array of rows,
//  padding with blanks and reapplying date format + alternating stripes.
// ============================================================
function rebuildDetailsBlock(sheet, firstDataRow, rows) {
  const blockRange = sheet.getRange(firstDataRow, 1, MAX_DETAIL_ROWS, DETAIL_COLS);
  blockRange.clearContent();

  const padded = rows.slice();
  while (padded.length < MAX_DETAIL_ROWS) padded.push(new Array(DETAIL_COLS).fill(''));
  blockRange.setValues(padded);

  if (rows.length > 0) {
    sheet.getRange(firstDataRow, 1, rows.length, 1).setNumberFormat('MMM dd, yyyy');
  }

  const stripes = [];
  for (let i = 0; i < MAX_DETAIL_ROWS; i++) {
    const color = (i < rows.length) ? ((i + 1) % 2 === 0 ? '#F9F9F9' : '#FFFFFF') : '#FFFFFF';
    stripes.push(new Array(DETAIL_COLS).fill(color));
  }
  blockRange.setBackgrounds(stripes);
}

// ============================================================
//  GRID: rebuilt from the details table every run (it's derived data)
// ============================================================
function renderGrid(sheet, year, month, layout, tz) {
  const firstDataRow = layout.detailsHeaderRow + 1;
  const gridRows = layout.gridLastRow - 2;
  const detail = sheet.getRange(firstDataRow, 1, MAX_DETAIL_ROWS, DETAIL_COLS).getValues();

  // Group detail rows by day of month
  const byDay = {};
  detail.forEach((r, i) => {
    if (!(r[IDX.DATE] instanceof Date)) return;
    if (Utilities.formatDate(r[IDX.DATE], tz, 'yyyy-M') !== year + '-' + month) return;
    const day = parseInt(Utilities.formatDate(r[IDX.DATE], tz, 'd'), 10);
    if (!byDay[day]) byDay[day] = [];
    byDay[day].push({ row: firstDataRow + i, title: String(r[IDX.TITLE]), label: r[IDX.SOURCE] || '📚' });
  });

  const values = [];
  const sizes = [];
  for (let r = 0; r < gridRows; r++) {
    values.push(new Array(7).fill(''));
    sizes.push(new Array(7).fill(10));
  }

  for (let d = 1; d <= layout.daysInMonth; d++) {
    const idx = layout.startCol - 1 + (d - 1);
    const r = Math.floor(idx / 7);
    const c = idx % 7;
    const events = byDay[d];

    if (!events) {
      values[r][c] = d;
    } else {
      const lines = events
        .map(e => e.label + ' ' + e.title.replace(/"/g, '""'))
        .join('\n');
      values[r][c] = '=HYPERLINK("#gid=' + sheet.getSheetId() + '&range=A' + events[0].row + '", "' + d + '\n' + lines + '")';
      sizes[r][c] = 9;
    }
  }

  const gridRange = sheet.getRange(3, 1, gridRows, 7);
  gridRange.setValues(values);
  gridRange.setFontSizes(sizes);
}

// ============================================================
//  HELPERS
// ============================================================
// Event id + day => stable across title/time edits, distinct per recurring instance
function makeKey(event, tz) {
  return event.getId() + '|' + Utilities.formatDate(event.getStartTime(), tz, 'yyyy-MM-dd');
}

function isBlankRow(r) {
  return r[IDX.DATE] === '' && r[IDX.TITLE] === '';
}

function getMonthName(monthNumber) {
  return ['January','February','March','April','May','June','July','August','September','October','November','December'][monthNumber-1];
}

function formatTime(date) {
  return Utilities.formatDate(date, DISPLAY_TZ, "HH:mm");
}

// Returns e.g. "GMT+3" in summer or "GMT+2" in winter, based on the event's own date
function getTzLabel(date) {
  const z = Utilities.formatDate(date, DISPLAY_TZ, "Z"); // e.g. "+0300"
  const sign = z.charAt(0);
  const hours = parseInt(z.substr(1, 2), 10);
  const mins = z.substr(3, 2);
  return 'GMT' + sign + hours + (mins !== '00' ? ':' + mins : '');
}