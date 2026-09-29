# Calendar Sync for Google Sheets

A Google Apps Script that turns Google Calendar events into a monthly planning workbook:
a calendar grid plus a details table with dropdown status fields, kept in sync with the calendars
without overwriting anything you type by hand.

## What it demonstrates
- **Idempotent, incremental sync:** new events are added, existing ones refreshed, manual notes never touched
- **Stable identity:** each row is keyed by calendar event ID plus date, so recurring events stay distinct
- **Safe deletion:** removed events are flagged, and only untouched rows are cleaned up on request from a menu
- **Spreadsheet features:** data validation dropdowns, conditional formatting (past days turn green), custom menu, in-sheet hyperlinks from the grid to details
- **Correct time handling:** a named time zone, so daylight saving changes are handled automatically

## Files
- `calendar_sync.gs`: the script (calendar names are placeholders)
- `Calendar_Sync_Demo_Workbook.xlsx`: an example of the output using invented events

## Setup
1. Create two Google Calendars named `Demo Training Calendar` and `Demo Class Calendar` and add a few events, or edit `CALENDAR_CONFIGS`.
2. Open a new Google Sheet, then Extensions > Apps Script, and paste in `calendar_sync.gs`.
3. Set `START_MONTH`, `START_YEAR`, `END_MONTH`, `END_YEAR` and `DISPLAY_TZ`.
4. Reload the sheet and run **Calendar Sync > Sync new events now** (approve the permissions once).

All data in the demo workbook is invented.
