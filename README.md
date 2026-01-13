# Shift-Manager-Google-App-Script
Google Apps Script to generate and maintain two-week staff schedules with previews, stats, manual edits, and rotation tracking from a Google Sheet.

## Shift Manager Script Guide

### Overview
Google Apps Script that builds two-week schedules from a “Staff Info” sheet, writes them to a “Schedule” sheet, and keeps counts/rotation data in sync. It also provides previews, statistics, manual edits, and reset utilities.

### Prerequisites
- Host sheet with these tabs (names can be customized in Settings): `Settings`, `Staff Info`, `Schedule`, `Logs`, `Count Backups`.
- Apps Script services: SpreadsheetApp, HtmlService, MailApp, Session, Utilities.

### Configure Settings
In the `Settings` tab (col A=Key, col B=Value), define:
- `Schedule Sheet Name`, `Staff Info Sheet Name`, `Logs Sheet Name`, `Count Backups Sheet Name`
- `Staff Req Weekday Regular`, `Staff Req Friday Evening`, `Staff Req Saturday`, `Staff Req Sunday`
- `On Leave Status Text`

### Prepare Staff Info Sheet
Required columns (exact headers):
- `Players`, `Rotation Count`, `Weekday Shifts`, `Weekend Shifts`, `Total Shifts`, `Region`, `Requests`, `Last Extra Shift`, `Last Weekend Shift`, `Status`
Optional but recommended: `Level`, `Email`, `Contact number`, `Given`, `Preferred Slots`.

### Key Actions
- **Generate schedule:** `generateSchedule()`  
  - Loads settings/staff, builds schedule, shows preview modal, applies to `Schedule`, updates counts, checks rotation completion.
- **View statistics:** `displayShiftStatistics()`  
  - Calculates stats from the current `Schedule`, shows modal with CSV download and email-send.
- **Manual edit:** `showManualShiftEditor()`  
  - UI to edit a specific shift, updates counts accordingly.
- **Reset counts:** `showResetConfirmation()`  
  - Backs up counts to `Count Backups`, then resets rotation/shift counts.

### Schedule Output Format
- Weekday grid: time slots as rows (two rows per slot for staff1/staff2), days as columns.
- Weekend list: time slot + staff1/staff2 + day.

### Usage Tips
- Keep headers exact; validation will fail fast on missing/renamed columns.
- Use ISO date format in counts (`YYYY-MM-DD`) for consistency.
- For CSV/email stats, ensure FileSaver is available or rely on the built-in fallback download.

### Error Handling
- Critical failures log to `Logs` and surface alerts.
- Missing sheets throw errors early (via `getSheetOrThrow`).
- Fallback scheduling stages try to avoid empty slots and log when constraints are relaxed.


