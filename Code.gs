/**
 * @OnlyCurrentDoc
 */

// =================================================================
// SECTION 1: CONSTANTS, SETTINGS, HELPERS, STAFF INFO HANDLING
// =================================================================

/**
 * Defines keys expected in the 'Settings' sheet (Column A = Key, Column B = Value).
 * Requirement interpretations:
 * REQ_WEEKDAY: Staff needed PER weekday slot (Mon-Fri 8-12, 12-4, 4-8). Typically 2.
 * REQ_FRI_EVE: Staff needed for Fri 4pm-8pm slot. Typically 2.
 * REQ_SATURDAY: Staff needed PER Saturday slot (9-1, 1-5). Typically 2.
 * REQ_SUNDAY: Staff needed PER Sunday slot (9-1, 1-5). Typically 2.
 */
const SETTINGS_KEYS = {
  // Sheet Names
  SCHEDULE_SHEET: 'Schedule Sheet Name',
  STAFF_INFO_SHEET: 'Staff Info Sheet Name',
  LOGS_SHEET: 'Logs Sheet Name',
  BACKUPS_SHEET: 'Count Backups Sheet Name',
  // Staff Requirements (Defaults used if setting missing)
  REQ_WEEKDAY: 'Staff Req Weekday Regular',
  REQ_FRI_EVE: 'Staff Req Friday Evening',
  REQ_SATURDAY: 'Staff Req Saturday', 
  REQ_SUNDAY: 'Staff Req Sunday',     
  // Other settings
  ON_LEAVE_STATUS_TEXT: 'On Leave Status Text' // Text used in Status column to indicate leave
};

/**
 * Defines standard time slots for scheduling.
 */
const WEEKDAY_TIME_SLOTS = {
  REGULAR: ['8am - 12pm', '12pm - 4pm', '4pm - 8pm'],
  FRIDAY_EVENING: ['8am - 12pm', '12pm - 4pm', '4pm - 8pm'] 
};

const WEEKEND_TIME_SLOTS = {
  SATURDAY: ['9am - 1pm', '1pm - 5pm'],
  SUNDAY: ['9am - 1pm', '1pm - 5pm']
};

/**
 * Defines arrays for days of the week.
 */
const WEEKDAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
const WEEKEND_DAYS = ['Saturday', 'Sunday'];
const ALL_DAYS = [...WEEKDAYS, ...WEEKEND_DAYS];

/**
 * Header names expected in the 'Staff Info' sheet.
 * It's CRUCIAL that these strings EXACTLY match the headers in your sheet.
 */
const STAFF_INFO_HEADERS = {
  PLAYER_NAME: 'Players',
  ROTATION_COUNT: 'Rotation Count', // Numerical counter for completed extra-shift rotations (used for extra shift fairness)
  WEEKDAY_SHIFTS: 'Weekday Shifts', // Historical total weekday shifts worked
  WEEKEND_SHIFTS: 'Weekend Shifts', // Historical total weekend shifts worked (used for weekend fairness)
  TOTAL_SHIFTS: 'Total Shifts',     // Historical total shifts (sum of weekday and weekend, used for overall fairness tie-breaking)
  REGION: 'Region',
  LEVEL: 'Level',
  REQUESTS: 'Requests',             // Availability requests string
  EMAIL: 'Email',
  CONTACT: 'Contact number',
  ROLE: 'Given',
  LAST_EXTRA_SHIFT: 'Last Extra Shift', // Date string (when they last received an *extra* weekday shift)
  LAST_WEEKEND_SHIFT: 'Last Weekend Shift', // Date string (when they last worked *any* weekend shift) - also used for weekend fairness
  STATUS: 'Status',                 // e.g., "Active", "On Leave"
  PREFERRED_SLOTS: 'Preferred Slots' // e.g., "Mon 8-12, Sat 9-1"
};

// --- Helper Functions ---

/**
 * Helper function to get the previous day in the WEEKDAYS array.
 * @param {string} day - Current day (e.g., 'Tuesday').
 * @return {string|null} Previous day or null if it's the first day (Monday).
 */
function getPreviousDay(day) {
  const index = WEEKDAYS.indexOf(day);
  return index > 0 ? WEEKDAYS[index - 1] : null;
}

/**
 * Builds a CSV string from the stats object for download or email.
 * @param {Object} stats
 * @return {string}
 */
function buildStatsCsv(stats) {
  if (!stats || !stats.playerStats) return '';
  const escapeCsvCell = (cellData) => {
    if (cellData === null || cellData === undefined) return '';
    const str = String(cellData);
    if (/[",\n]/.test(str)) {
      return `"${str.replace(/"/g, '""')}"`;
    }
    return str;
  };

  const csvContent = [];
  csvContent.push('Staff Member,Total Shifts,Weekday Shifts,Weekend Shifts,Utilization %,Level,Region,Status');
  Object.keys(stats.playerStats).sort().forEach(name => {
    const staff = stats.playerStats[name];
    csvContent.push([
      escapeCsvCell(name),
      escapeCsvCell(staff.totalShifts),
      escapeCsvCell(staff.weekdayShifts),
      escapeCsvCell(staff.weekendShifts),
      escapeCsvCell((staff.utilizationPercentage || 0).toFixed(1)),
      escapeCsvCell(staff.level),
      escapeCsvCell(staff.region),
      escapeCsvCell(staff.status)
    ].join(','));
  });

  csvContent.push('');
  csvContent.push('Overall Statistics,Value');
  csvContent.push('Total Shifts,' + escapeCsvCell(stats.totalShifts));
  csvContent.push('Average Shifts Per Staff,' + escapeCsvCell((stats.averageShiftsPerPlayer || 0).toFixed(1)));

  if (stats.imbalances && stats.imbalances.length > 0) {
    csvContent.push('');
    csvContent.push('Potential Imbalances');
    stats.imbalances.forEach(imbalance => csvContent.push(escapeCsvCell(imbalance)));
  }

  return csvContent.join('\n');
}

/**
 * Sends statistics via email to the active/effective user with CSV attachment.
 * @param {string} statsJson - JSON string of stats object.
 */
function sendStatisticsEmail(statsJson) {
  if (!statsJson) throw new Error('Stats payload missing');
  let statsObj;
  try {
    statsObj = JSON.parse(statsJson);
  } catch (e) {
    throw new Error(`Failed to parse stats JSON: ${e.message}`);
  }
  const recipient = Session.getActiveUser()?.getEmail() || Session.getEffectiveUser()?.getEmail();
  if (!recipient) throw new Error('No recipient email available.');

  const subject = 'Shift Statistics';
  const body = 'Attached are the latest shift statistics in CSV format.';
  const csvString = buildStatsCsv(statsObj);
  const blob = Utilities.newBlob(csvString, 'text/csv', 'shift_statistics.csv');

  MailApp.sendEmail({
    to: recipient,
    subject,
    body,
    attachments: [blob]
  });
}

/**
 * Helper function to get the next day in the WEEKDAYS array.
 * @param {string} day - Current day (e.g., 'Thursday').
 * @return {string|null} Next day or null if it's the last day (Friday).
 */
function getNextDay(day) {
  const index = WEEKDAYS.indexOf(day);
  return index < WEEKDAYS.length - 1 ? WEEKDAYS[index + 1] : null;
}

/**
 * Checks if a staff member is already assigned to any shift on a given day
 * within a specific week's schedule object.
 * @param {string} staff - Staff member's name.
 * @param {string} day - Day to check (e.g., 'Monday').
 * @param {Object} weekSchedule - The schedule object for a single week (e.g., schedule.week1).
 * @return {boolean} True if staff is assigned to the day, false otherwise.
 */
function isStaffAssignedToDay(staff, day, weekSchedule) {
  if (!staff || staff === '-' || !weekSchedule || !weekSchedule[day]) {
    return false;
  }
  const daySlots = weekSchedule[day];
  for (const slot in daySlots) {
    const assignment = daySlots[slot];
    // Check both staff slots if they exist
    if (assignment && (assignment.staff1 === staff || assignment.staff2 === staff)) {
        return true;
    }
  }
  return false;
}


/**
 * Checks if a staff member is available on a given day based on their request string.
 * Case-insensitive checks for common request patterns.
 * Does NOT override requests - if a request makes them unavailable, they are unavailable.
 * @param {string} staffName - Staff member's name (for logging/debugging).
 * @param {string} day - Day to check availability for (e.g., 'Wednesday').
 * @param {string} requests - Staff member's requests string (lowercase).
 * @return {boolean} True if staff is available, false otherwise.
 */
function isStaffAvailableOnDay(staffName, day, requests) {
  if (!requests) return true; // No requests means available

  const dayLower = day.toLowerCase();
  const requestsLower = requests.toLowerCase();

  // Check for explicit day exclusions (e.g., "not Monday", "not mondays")
  if (requestsLower.includes(`not to be shifted ${dayLower}`) || requestsLower.includes(`not to be shifted ${dayLower}s`)) return false;
  if (dayLower === 'friday' && requestsLower.includes('not to be shifted fridays')) return false;
  if (dayLower === 'wednesday' && requestsLower.includes('not to be shifted wednesdays')) return false;
  if ((dayLower === 'wednesday' || dayLower === 'friday') && requestsLower.includes('not to be shifted on wednesdays and fridays')) return false;
  if (WEEKEND_DAYS.map(d => d.toLowerCase()).includes(dayLower) && requestsLower.includes('not to be shifted weekends')) return false;

  // Check for 'only' days (must match one of the allowed 'only' days if 'only' is present)
  // Example: "only tuesdays and thursdays" -> onlyDays = ['tuesdays', 'thursdays']
  const onlyMatch = requestsLower.match(/only\s+(.*)$/);
  if (onlyMatch && onlyMatch[1]) {
      const onlyDaysFragment = onlyMatch[1];
      const onlyDays = onlyDaysFragment.split(/[,\s&]+/) // Split by comma, space, and &
                                       .map(d => d.trim().toLowerCase()) // Trim and lowercase each part
                                       .filter(d => d.length > 0 && d !== 'and'); // Remove empty parts and 'and'

      if (onlyDays.length > 0) {
           // Check if the current day (full name or abbreviation) matches any of the 'only' days
           const isAllowed = onlyDays.some(onlyDayStr => {
               // Simple check: does the request string part contain the day name or its first 3 letters?
               // This is an approximation; more complex parsing might be needed for "only weekend" etc.
               return onlyDayStr.includes(dayLower) || dayLower.includes(onlyDayStr) || // "monday".includes("mon")
                      onlyDayStr.startsWith(dayLower.substring(0, 3)); // Check abbreviation start
           });
           if (!isAllowed) {
               // console.log(`${staffName} unavailable on ${day} due to 'only' request: "${requests}"`);
               return false; // If 'only' is present, and current day isn't one of the 'only' days, they are unavailable.
           }
      }
      // If 'only' is present but no valid days are listed after it, perhaps treat as unavailable?
      // For now, if onlyDays is empty after parsing, this check passes, meaning availability relies only on exclusions.
  }

   // Check for specific slot exclusions within requests if format is like "no 8am-12pm Mon" or similar.
   // The current isStaffAvailableOnDay is day-level only. Slot-level preference check is in isSlotPreferred.
   // If you need hard *exclusions* by slot in the request string, this function would need expansion,
   // but the prompt indicates "meeting them halfway" via requests primarily means day-level.

  return true; // Available if passed all exclusion checks and matched 'only' criteria if present
}

/**
 * Returns an ISO-8601 date string (YYYY-MM-DD) for deterministic storage.
 */
function toIsoDateString(date = new Date()) {
  return date.toISOString().slice(0, 10);
}

/**
 * Ensures we always retrieve an existing sheet or fail fast with context.
 * @param {string} name - Sheet name.
 * @param {string} [context=''] - Optional context for error clarity.
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getSheetOrThrow(name, context = '') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    const prefix = context ? `${context}: ` : '';
    throw new Error(`${prefix}Sheet "${name}" not found.`);
  }
  return sheet;
}

/**
 * Performs a lightweight deep clone for plain data objects.
 * @param {Object} obj
 * @return {Object}
 */
function safeJsonClone(obj) {
  return obj ? JSON.parse(JSON.stringify(obj)) : obj;
}

/**
 * Initializes an empty week schedule object structure matching the applyScheduleToSheet expectations (staff1/staff2 per slot).
 * This structure is used internally during schedule generation.
 * @return {Object} An empty week schedule object with days and time slots, defaulting to space for 2 staff placeholders.
 */
function initializeWeek() {
    const weekStructure = {};
    ALL_DAYS.forEach(day => {
        weekStructure[day] = {};
        let timeSlots = [];
        // Use the timeslot constants directly as initialize only builds the structure
        if (WEEKDAYS.includes(day)) {
          timeSlots = (day === 'Friday') ? WEEKDAY_TIME_SLOTS.FRIDAY_EVENING : WEEKDAY_TIME_SLOTS.REGULAR;
          timeSlots.forEach(slot => {
             weekStructure[day][slot] = { staff1: '', staff2: '' }; // Weekdays expect staff1/staff2
          });
        } else { // Weekend Days
          timeSlots = (day === 'Saturday') ? WEEKEND_TIME_SLOTS.SATURDAY : WEEKEND_TIME_SLOTS.SUNDAY;
           timeSlots.forEach(slot => {
             // All slots now support staff1/staff2 internally.
             weekStructure[day][slot] = { staff1: '', staff2: '' };
           });
        }
    });
  return weekStructure;
}

// --- Settings Function ---

/**
 * Reads settings from the 'Settings' sheet (assumed Col A = Key, Col B = Value).
 * Provides default values if settings are missing.
 * Validates numeric settings are non-negative integers and sheet names are non-empty.
 * Note: This function does NOT call logAction directly anymore to prevent recursion issues.
 * Callers of getSettings are responsible for catching errors and logging them via logAction.
 * @return {Object} An object containing the settings.
 * @throws {Error} If settings sheet is found but errors occur while reading, or if critical settings are invalid.
 */
function getSettings() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheetName = 'Settings'; // Keep default name hardcoded here
    const sheet = ss.getSheetByName(settingsSheetName);
    const settings = {};

    // Defaults with descriptions aligned to interpretation
    // UPDATED: All requirements default to 2 as per user request.
    const defaults = {
        [SETTINGS_KEYS.SCHEDULE_SHEET]: 'Schedule',
        [SETTINGS_KEYS.STAFF_INFO_SHEET]: 'Staff Info',
        [SETTINGS_KEYS.LOGS_SHEET]: 'Logs',
        [SETTINGS_KEYS.BACKUPS_SHEET]: 'Count Backups',
        [SETTINGS_KEYS.REQ_WEEKDAY]: 2,    // Default: 2 staff per regular weekday slot
        [SETTINGS_KEYS.REQ_FRI_EVE]: 2,    // Default: 2 staff for Fri 4pm-8pm slot
        [SETTINGS_KEYS.REQ_SATURDAY]: 2,   // Default: 2 staff PER Saturday slot
        [SETTINGS_KEYS.REQ_SUNDAY]: 2,     // Default: 2 staff PER Sunday slot
        [SETTINGS_KEYS.ON_LEAVE_STATUS_TEXT]: 'On Leave'
    };

    if (!sheet) {
        console.warn(`Settings sheet "${settingsSheetName}" not found. Using default settings.`);
        // Return defaults immediately if sheet is missing. Caller must log.
        return defaults;
    }

    try {
        // Read specific range (assuming keys in Col A, values in Col B)
        const lastRow = sheet.getLastRow();
        if (lastRow === 0) {
             console.warn(`Settings sheet "${settingsSheetName}" is empty. Using default settings.`);
             return defaults; // Return defaults if sheet is empty. Caller must log.
        }
        const data = sheet.getRange(1, 1, lastRow, 2).getValues(); // Read A1:B[lastRow]

        data.forEach(row => {
            const key = row[0] ? String(row[0]).trim() : null;
            const value = row[1]; // Keep original type (number, string, boolean, date)
            if (key && Object.values(SETTINGS_KEYS).includes(key)) { // Only include keys defined in SETTINGS_KEYS constant
                 settings[key] = (value === undefined || value === null) ? '' : value;
            } else if (key) {
                // Optional: log if a setting key is in the sheet but not recognized
                // console.log(`Unrecognized setting key in sheet: "${key}" at row ${data.indexOf(row) + 1}.`);
            }
        });
    } catch (e) {
         // Catch errors during reading data from the sheet. Caller must log.
         console.error(`Error reading data from Settings sheet "${settingsSheetName}": ${e}`);
         // Re-throw the error so the caller's catch block can handle and log it.
         throw new Error(`Error reading data from Settings sheet "${settingsSheetName}": ${e.message}`);
    }

    // Merge defaults with read settings (settings from sheet override defaults)
    const finalSettings = { ...defaults, ...settings };

    // Find the keys from SETTINGS_KEYS object based on the values from the sheet
    const finalSettingsByKeyName = {};
    Object.keys(SETTINGS_KEYS).forEach(keyName => {
        const sheetKey = SETTINGS_KEYS[keyName];
        finalSettingsByKeyName[keyName] = finalSettings[sheetKey];
    });


    // Type Conversion and Validation for Numeric Settings
    const numericKeys = ['REQ_WEEKDAY', 'REQ_FRI_EVE', 'REQ_SATURDAY', 'REQ_SUNDAY'];
    numericKeys.forEach(keyName => {
        const sheetKey = SETTINGS_KEYS[keyName];
        const value = finalSettingsByKeyName[keyName];
        const defaultValue = defaults[sheetKey];
        const parsedValue = Number(value);
        const isValid = Number.isInteger(parsedValue) && parsedValue >= 0;
        const numValue = isValid ? parsedValue : defaultValue;

        if (!isValid) {
            if (value === '' || value === null || value === undefined) {
                console.warn(`Missing setting for "${sheetKey}". Using default: ${defaultValue}.`);
            } else {
                console.warn(`Invalid numeric setting for "${sheetKey}" (${value}). Must be a non-negative integer. Using default: ${defaultValue}.`);
            }
        }
        finalSettingsByKeyName[keyName] = numValue; // Use validated number (or fallback)
    });

    // Type Conversion and Validation for Sheet Name Settings
    const sheetNameKeys = ['SCHEDULE_SHEET', 'STAFF_INFO_SHEET', 'LOGS_SHEET', 'BACKUPS_SHEET'];
     sheetNameKeys.forEach(keyName => {
         const sheetKey = SETTINGS_KEYS[keyName];
         finalSettingsByKeyName[keyName] = String(finalSettingsByKeyName[keyName] || defaults[sheetKey]).trim();

         if (finalSettingsByKeyName[keyName] === '') {
             // This is a critical configuration error if a required sheet name is empty
             const errorMsg = `Critical: Setting "${sheetKey}" is empty (""). Please correct the Settings sheet or script defaults.`;
             console.error(errorMsg);
             // Re-throw here as this will cause later functions to fail
             throw new Error(errorMsg);
         }
     });

     // Ensure status text is a trimmed string
     finalSettingsByKeyName['ON_LEAVE_STATUS_TEXT'] = String(finalSettingsByKeyName['ON_LEAVE_STATUS_TEXT'] || defaults[SETTINGS_KEYS.ON_LEAVE_STATUS_TEXT]).trim();
     if (finalSettingsByKeyName['ON_LEAVE_STATUS_TEXT'] === '') {
          console.warn(`Setting "${SETTINGS_KEYS.ON_LEAVE_STATUS_TEXT}" is empty. Staff status will not be correctly identified as 'On Leave'. Using default: "${defaults[SETTINGS_KEYS.ON_LEAVE_STATUS_TEXT]}".`);
          finalSettingsByKeyName['ON_LEAVE_STATUS_TEXT'] = defaults[SETTINGS_KEYS.ON_LEAVE_STATUS_TEXT];
     }

    console.log("Settings loaded:", finalSettingsByKeyName);
    return finalSettingsByKeyName; // Return the final validated settings object keyed by constant name
}


// --- Staff Info Sheet Handling ---

/**
 * Validates the structure and critical data of the Staff Info sheet.
 * Checks for required columns and basic data integrity (e.g., numeric fields, non-empty names/status).
 * @param {Sheet} sheet - The Google Sheet object for 'Staff Info'.
 * @param {Object} settings - The loaded settings object.
 * @return {string[]} An array of validation error messages. Empty if no errors.
 */
function validateStaffInfoSheet(sheet, settings) {
  const errors = [];
  const staffInfoSheetName = settings.STAFF_INFO_SHEET;

  if (!sheet) {
    errors.push(`Staff Info sheet "${staffInfoSheetName}" not found.`);
    return errors;
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    errors.push(`Staff Info sheet "${staffInfoSheetName}" has no data rows (only header or empty).`);
    return errors;
  }

  const headers = data[0].map(h => String(h).trim());
  const requiredCoreColumns = [
    STAFF_INFO_HEADERS.PLAYER_NAME,
    STAFF_INFO_HEADERS.ROTATION_COUNT,
    STAFF_INFO_HEADERS.WEEKDAY_SHIFTS,
    STAFF_INFO_HEADERS.WEEKEND_SHIFTS,
    STAFF_INFO_HEADERS.TOTAL_SHIFTS,
    STAFF_INFO_HEADERS.REGION,
    STAFF_INFO_HEADERS.REQUESTS,
    STAFF_INFO_HEADERS.LAST_EXTRA_SHIFT,
    STAFF_INFO_HEADERS.LAST_WEEKEND_SHIFT,
    STAFF_INFO_HEADERS.STATUS
  ];

  const missingColumns = requiredCoreColumns.filter(col => !headers.includes(col));
  if (missingColumns.length > 0) {
    errors.push(`Missing critical required columns in "${staffInfoSheetName}": ${missingColumns.join(', ')}. Please ensure headers match exactly.`);
  }

  // Optional columns - warn if missing but don't block execution
  if (!headers.includes(STAFF_INFO_HEADERS.LEVEL)) { console.warn(`Optional column "${STAFF_INFO_HEADERS.LEVEL}" not found in "${staffInfoSheetName}".`); }
  if (!headers.includes(STAFF_INFO_HEADERS.EMAIL)) { console.warn(`Optional column "${STAFF_INFO_HEADERS.EMAIL}" not found in "${staffInfoSheetName}".`); }
  if (!headers.includes(STAFF_INFO_HEADERS.CONTACT)) { console.warn(`Optional column "${STAFF_INFO_HEADERS.CONTACT}" not found in "${staffInfoSheetName}".`); }
  if (!headers.includes(STAFF_INFO_HEADERS.ROLE)) { console.warn(`Optional column "${STAFF_INFO_HEADERS.ROLE}" not found in "${staffInfoSheetName}".`); }
  if (!headers.includes(STAFF_INFO_HEADERS.PREFERRED_SLOTS)) { console.warn(`Optional column "${STAFF_INFO_HEADERS.PREFERRED_SLOTS}" not found in "${staffInfoSheetName}". Preference sorting will be less effective.`); }


  // Get column indices after checking for missing ones
  const playerColIndex = headers.indexOf(STAFF_INFO_HEADERS.PLAYER_NAME);
  const rotationColIndex = headers.indexOf(STAFF_INFO_HEADERS.ROTATION_COUNT);
  const weekdayColIndex = headers.indexOf(STAFF_INFO_HEADERS.WEEKDAY_SHIFTS);
  const weekendColIndex = headers.indexOf(STAFF_INFO_HEADERS.WEEKEND_SHIFTS);
  const totalColIndex = headers.indexOf(STAFF_INFO_HEADERS.TOTAL_SHIFTS);
  const regionColIndex = headers.indexOf(STAFF_INFO_HEADERS.REGION);
  const statusColIndex = headers.indexOf(STAFF_INFO_HEADERS.STATUS);
  const lastExtraShiftColIndex = headers.indexOf(STAFF_INFO_HEADERS.LAST_EXTRA_SHIFT);
  const lastWeekendShiftColIndex = headers.indexOf(STAFF_INFO_HEADERS.LAST_WEEKEND_SHIFT);


  // Validate data rows (starting from row 2, index 1)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowNum = i + 1; // Sheet row number (1-based)
    const staffName = (playerColIndex !== -1 && row[playerColIndex]) ? String(row[playerColIndex]).trim() : '';

    // Check for missing Player Name (critical)
    if (playerColIndex === -1 || staffName === '') {
      errors.push(`Row ${rowNum}: ${STAFF_INFO_HEADERS.PLAYER_NAME} is empty or column is missing.`);
      continue; // Cannot validate rest of row without a name
    }

    // Check for missing Status (critical for filtering)
    if (statusColIndex !== -1) {
        const statusValue = String(row[statusColIndex] || '').trim();
        if (statusValue === '') {
             errors.push(`Row ${rowNum} (${staffName}): ${STAFF_INFO_HEADERS.STATUS} is empty.`);
        }
    } else if (requiredCoreColumns.includes(STAFF_INFO_HEADERS.STATUS)) {
         errors.push(`Row ${rowNum} (${staffName}): ${STAFF_INFO_HEADERS.STATUS} column is missing.`);
    }

    // Check numeric columns for non-negative integer values (or empty)
    const numericColsToCheck = [
      { name: STAFF_INFO_HEADERS.ROTATION_COUNT, index: rotationColIndex },
      { name: STAFF_INFO_HEADERS.WEEKDAY_SHIFTS, index: weekdayColIndex },
      { name: STAFF_INFO_HEADERS.WEEKEND_SHIFTS, index: weekendColIndex },
      { name: STAFF_INFO_HEADERS.TOTAL_SHIFTS, index: totalColIndex }
    ];
    numericColsToCheck.forEach(col => {
      if (col.index !== -1) {
        const value = row[col.index];
        if (value !== '' && value !== null) {
            const num = Number(value);
            if (typeof value === 'string' && value.trim() === '') {
                // Okay, treat empty string as valid
            } else if (typeof num !== 'number' || !isFinite(num) || num < 0 || !Number.isInteger(num)) {
                 errors.push(`Row ${rowNum} (${staffName}): Invalid value "${value}" in "${col.name}". Must be a non-negative integer or empty.`);
            }
        }
      } else if (requiredCoreColumns.includes(col.name)) {
          // Already reported as missing column
      }
    });

    // Check for missing Region (critical)
     if (regionColIndex !== -1) {
         const regionValue = String(row[regionColIndex] || '').trim();
         if (regionValue === '') {
             errors.push(`Row ${rowNum} (${staffName}): ${STAFF_INFO_HEADERS.REGION} is empty.`);
         }
     } else if (requiredCoreColumns.includes(STAFF_INFO_HEADERS.REGION)) {
         // Already reported as missing column
     }
  }

  return errors;
}


/**
 * Retrieves staff information from the Staff Info sheet into an object format.
 * Includes Status and Preferred Slots. Calculates historical Total Shifts as sum of components.
 * @param {Sheet} sheet - The Google Sheet object for 'Staff Info'.
 * @param {Object} settings - The loaded settings object.
 * @return {Object} Staff information object keyed by player name. Throws error if validation fails.
 */
function getStaffInfo(sheet, settings) {
  // validateStaffInfoSheet throws error if critical validation fails
  const validationErrors = validateStaffInfoSheet(sheet, settings);
  if (validationErrors.length > 0) {
    const errorMsg = `Staff Info Sheet Validation Failed for "${settings.STAFF_INFO_SHEET}":\n- ${validationErrors.join('\n- ')}`;
    console.error(errorMsg);
    // Log validation errors using logAction, passing settings
    logAction('Staff Info Validation Failed', errorMsg, 'ERROR', settings);
    throw new Error(errorMsg); // Re-throw to stop execution
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const colIndices = {};
  for (const key in STAFF_INFO_HEADERS) { colIndices[key] = headers.indexOf(STAFF_INFO_HEADERS[key]); }

  const staffInfo = {};
  const onLeaveStatusText = settings.ON_LEAVE_STATUS_TEXT;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = colIndices.PLAYER_NAME !== -1 ? String(row[colIndices.PLAYER_NAME] || '').trim() : '';
    if (!name) continue;

    const getValue = (colKey, defaultValue = '') => {
        const index = colIndices[colKey];
        return (index !== -1 && row[index] !== undefined && row[index] !== null && String(row[index]).trim() !== '') ? String(row[index]).trim() : defaultValue;
    };

    const getNumericValue = (colKey, defaultValue = 0) => {
        const index = colIndices[colKey];
        if (index === -1) return defaultValue;
        const val = row[index];
        if (val === '' || val === null) return defaultValue;
        const num = Number(val);
        return (typeof num === 'number' && isFinite(num)) ? num : defaultValue;
    };

    const weekdayShifts = getNumericValue('WEEKDAY_SHIFTS', 0);
    const weekendShifts = getNumericValue('WEEKEND_SHIFTS', 0);

    staffInfo[name] = {
      name: name,
      level: getValue('LEVEL'),
      region: getValue('REGION'),
      requests: getValue('REQUESTS'),
      email: getValue('EMAIL'),
      contact: getValue('CONTACT'),
      role: getValue('ROLE'),
      rotationCount: getNumericValue('ROTATION_COUNT', 0),
      weekdayShifts: weekdayShifts,
      weekendShifts: weekendShifts,
      shiftCount: weekdayShifts + weekendShifts, // Calculated sum

      lastExtraShiftDate: getValue('LAST_EXTRA_SHIFT', ''),
      lastWeekendShiftDate: getValue('LAST_WEEKEND_SHIFT', ''),

      status: getValue('STATUS', 'Active'),
      preferredSlots: getValue('PREFERRED_SLOTS', ''),

      // Current Rotation Tracking (Initialized)
      currentRotationWeekdayShifts: 0,
      currentRotationWeekendShifts: 0,
      currentRotationTotalShifts: 0,
      extraShiftGivenThisRotation: false
    };

     // Optional: Sanity check against the sheet's Total Shifts value if column exists
     const sheetTotal = getNumericValue('TOTAL_SHIFTS', staffInfo[name].shiftCount); // Default to calculated if sheet col missing/invalid
     if (sheetTotal !== staffInfo[name].shiftCount) {
          console.warn(`Staff Info: Row ${i+1} (${name}) Total Shifts mismatch. Sheet: ${sheetTotal}, Calculated: ${staffInfo[name].shiftCount}. Using calculated sum internally.`);
     }
  }

  if (Object.keys(staffInfo).length === 0) {
    const errorMsg = `Staff Info Sheet Validation Failed: No valid staff data rows found in "${settings.STAFF_INFO_SHEET}". Ensure '${STAFF_INFO_HEADERS.PLAYER_NAME}' column is present and has names.`;
     console.error(errorMsg);
     logAction('Staff Info Load Error', errorMsg, 'ERROR', settings);
     throw new Error(errorMsg);
  }
  console.log(`Successfully loaded info for ${Object.keys(staffInfo).length} staff members from "${settings.STAFF_INFO_SHEET}".`);
  return staffInfo;
}


/**
 * Updates the Staff Info sheet with counts and last shift dates using batch update.
 * Does NOT update Status, Region, Level, Requests, Email, Contact, Role, or Preferred Slots.
 * @param {Object} staffInfo - The updated staff information object (after scheduling).
 * @param {Object} settings - The loaded settings object (to get sheet name).
 * @throws {Error} If the Staff Info sheet is not found or required update columns are missing.
 */
function updateStaffInfoSheet(staffInfo, settings) {
  console.log("--- Starting updateStaffInfoSheet (Batch Mode) ---");
  const sheetName = settings.STAFF_INFO_SHEET;
  let sheet;
  try {
      sheet = getSheetOrThrow(sheetName, 'updateStaffInfoSheet');
  } catch (e) {
      const errorMsg = e.message || `Update failed: Staff Info sheet "${sheetName}" not found.`;
      console.error(errorMsg);
      logAction('Update Staff Info Error', errorMsg, 'ERROR', settings);
      throw new Error(errorMsg);
  }

  const range = sheet.getDataRange();
  const data = range.getValues();
  if (data.length < 2) {
       console.log(`Staff Info sheet "${sheetName}" has no data rows to update.`);
       return;
  }

  const headers = data[0].map(h => String(h).trim());

  const colIndicesToUpdate = {
    name: headers.indexOf(STAFF_INFO_HEADERS.PLAYER_NAME),
    rotationCount: headers.indexOf(STAFF_INFO_HEADERS.ROTATION_COUNT),
    weekdayShifts: headers.indexOf(STAFF_INFO_HEADERS.WEEKDAY_SHIFTS),
    weekendShifts: headers.indexOf(STAFF_INFO_HEADERS.WEEKEND_SHIFTS),
    totalShifts: headers.indexOf(STAFF_INFO_HEADERS.TOTAL_SHIFTS),
    lastExtraShift: headers.indexOf(STAFF_INFO_HEADERS.LAST_EXTRA_SHIFT),
    lastWeekendShift: headers.indexOf(STAFF_INFO_HEADERS.LAST_WEEKEND_SHIFT)
  };

  const requiredUpdateCols = [
      STAFF_INFO_HEADERS.PLAYER_NAME, STAFF_INFO_HEADERS.ROTATION_COUNT,
      STAFF_INFO_HEADERS.WEEKDAY_SHIFTS, STAFF_INFO_HEADERS.WEEKEND_SHIFTS,
      STAFF_INFO_HEADERS.TOTAL_SHIFTS, STAFF_INFO_HEADERS.LAST_EXTRA_SHIFT,
      STAFF_INFO_HEADERS.LAST_WEEKEND_SHIFT
  ];
  const missingUpdateCols = requiredUpdateCols.filter(colName => headers.indexOf(colName) === -1);


  if (missingUpdateCols.length > 0) {
      const errorMsg = `Cannot update Staff Info sheet "${sheetName}". Missing required columns: ${missingUpdateCols.join(', ')}. Please ensure headers match exactly.`;
      console.error(errorMsg);
      logAction('Update Staff Info Error', errorMsg, 'ERROR', settings);
      throw new Error(errorMsg);
  }

  let updatedCount = 0;
  for (let i = 1; i < data.length; i++) {
    const staffNameInSheet = String(data[i][colIndicesToUpdate.name] || '').trim();
    if (!staffNameInSheet) continue;

    const staffMemberData = staffInfo[staffNameInSheet];

    if (staffMemberData) {
        const newTotalShifts = (staffMemberData.weekdayShifts ?? 0) + (staffMemberData.weekendShifts ?? 0);

        // Check if any value we *intend* to update is different
        const rotationChanged = data[i][colIndicesToUpdate.rotationCount] != (staffMemberData.rotationCount ?? 0);
        const weekdayChanged = data[i][colIndicesToUpdate.weekdayShifts] != (staffMemberData.weekdayShifts ?? 0);
        const weekendChanged = data[i][colIndicesToUpdate.weekendShifts] != (staffMemberData.weekendShifts ?? 0);
        const totalChanged = data[i][colIndicesToUpdate.totalShifts] != newTotalShifts; // Compare against the calculated total
        const lastExtraChanged = String(data[i][colIndicesToUpdate.lastExtraShift] || '').trim() !== String(staffMemberData.lastExtraShiftDate || '').trim();
        const lastWeekendChanged = String(data[i][colIndicesToUpdate.lastWeekendShift] || '').trim() !== String(staffMemberData.lastWeekendShiftDate || '').trim();


        if (rotationChanged || weekdayChanged || weekendChanged || totalChanged || lastExtraChanged || lastWeekendChanged) {
             // Update the array row with the new values
            data[i][colIndicesToUpdate.rotationCount] = staffMemberData.rotationCount ?? 0;
            data[i][colIndicesToUpdate.weekdayShifts] = staffMemberData.weekdayShifts ?? 0;
            data[i][colIndicesToUpdate.weekendShifts] = staffMemberData.weekendShifts ?? 0;
            data[i][colIndicesToUpdate.totalShifts] = newTotalShifts; // Write the calculated total
            data[i][colIndicesToUpdate.lastExtraShift] = staffMemberData.lastExtraShiftDate || '';
            data[i][colIndicesToUpdate.lastWeekendShift] = staffMemberData.lastWeekendShiftDate || '';
            updatedCount++;
        }
    } else {
        console.warn(`Staff "${staffNameInSheet}" found in sheet "${sheetName}" but not in the provided staffInfo object for updating. Row ${i+1}.`);
    }
  }

  if (updatedCount > 0) {
    range.setValues(data);
    console.log(`--- Finished updateStaffInfoSheet: ${updatedCount} staff records updated in sheet "${sheetName}". ---`);
  } else {
    console.log(`--- Finished updateStaffInfoSheet: No staff records required updating in sheet "${sheetName}". ---`);
  }

  SpreadsheetApp.flush();
}

// =================================================================
// END OF SECTION 1
// =================================================================

// =================================================================
// SECTION 2: SCHEDULE GENERATION LOGIC (REFACTORED AND CORRECTED)
// =================================================================

/**
 * Calculates the total number of weekday shifts for a two-week period based on settings.
 * @param {Object} settings - The loaded settings object.
 * @returns {number}
 */
function calculateTotalWeekdayShifts(settings) {
  let totalSingleWeek = 0;
  WEEKDAYS.forEach(day => {
    const timeSlots = (day === 'Friday') ? WEEKDAY_TIME_SLOTS.FRIDAY_EVENING : WEEKDAY_TIME_SLOTS.REGULAR;
    timeSlots.forEach(timeSlot => {
      // Use settings for requirements
      const req = (day === 'Friday' && timeSlot === '4pm - 8pm') ? settings.REQ_FRI_EVE : settings.REQ_WEEKDAY;
      totalSingleWeek += req;
    });
  });
  return totalSingleWeek * 2;
}

/**
 * Calculates the total number of weekend shifts for a two-week period based on settings.
 * @param {Object} settings - The loaded settings object.
 * @returns {number}
 */
function calculateTotalWeekendShifts(settings) {
    // Assumes settings define per-slot requirements for weekends
  const satShiftsPerWeek = WEEKEND_TIME_SLOTS.SATURDAY.length * settings.REQ_SATURDAY;
  const sunShiftsPerWeek = WEEKEND_TIME_SLOTS.SUNDAY.length * settings.REQ_SUNDAY;
  return (satShiftsPerWeek + sunShiftsPerWeek) * 2;
}

/**
 * Initializes the shift tracking object for all staff members.
 * @param {Object} staffInfo - Staff information object.
 * @returns {Object}
 */
function initializeShiftTracking(staffInfo) {
  const tracking = {};
  const staffMembers = Object.keys(staffInfo);
  staffMembers.forEach(staff => {
    tracking[staff] = {
      weekdayShifts: 0,
      weekendShifts: 0,
      totalShifts: 0,
      extraShiftAllowed: false
    };
  });
  return tracking;
}

/**
 * Gets eligible weekday staff for a shift, enforcing all standard rules.
 * @param {Object} staffInfo
 * @param {Object} shiftTracking
 * @param {string} day
 * @param {string} timeSlot
 * @param {Object} currentWeek
 * @param {number} baseShifts
 * @param {string[]} staffPool
 * @returns {string[]}
 */
function getEligibleWeekdayStaff(staffInfo, shiftTracking, day, timeSlot, currentWeek, baseShifts, staffPool) {
    return staffPool.filter(staff => {
        const staffData = staffInfo[staff];
        const trackingData = shiftTracking[staff];
        if (!trackingData || !staffData) return false;

        const currentWeekdayShifts = trackingData.weekdayShifts;
        const maxAllowedShifts = trackingData.extraShiftAllowed ? baseShifts + 1 : baseShifts;
        if (currentWeekdayShifts >= maxAllowedShifts) return false;

        if (isStaffAssignedToDay(staff, day, currentWeek)) return false;
        
        if (!isStaffAvailableOnDay(staff, day, staffData.requests || '')) return false;

        // Check for adjacent shift conflict (previous day's last slot)
        const firstSlotToday = (day === 'Friday' ? WEEKDAY_TIME_SLOTS.FRIDAY_EVENING : WEEKDAY_TIME_SLOTS.REGULAR)[0];
        if (timeSlot === firstSlotToday) {
            const prevDay = getPreviousDay(day);
            if (prevDay) {
                 const prevSlots = (prevDay === 'Friday' ? WEEKDAY_TIME_SLOTS.FRIDAY_EVENING : WEEKDAY_TIME_SLOTS.REGULAR);
                 const prevLastSlot = prevSlots[prevSlots.length - 1];
                 if (currentWeek[prevDay]?.[prevLastSlot]?.staff1 === staff || currentWeek[prevDay]?.[prevLastSlot]?.staff2 === staff) {
                     return false;
                 }
            }
        }
        return true;
    });
}

/**
 * Gets eligible weekend staff for a shift, enforcing all standard rules.
 * @param {Object} staffInfo
 * @param {Object} shiftTracking
 * @param {string} day
 * @param {string} timeSlot
 * @param {Object} currentWeek
 * @param {Set} week1WeekendStaff
 * @param {boolean} isWeek2
 * @returns {string[]}
 */
function getEligibleWeekendStaff(staffInfo, shiftTracking, day, timeSlot, currentWeek, week1WeekendStaff, isWeek2) {
  const staffPool = Object.keys(staffInfo);
  return staffPool.filter(staff => {
      const staffData = staffInfo[staff];
      const trackingData = shiftTracking[staff];
      if (!trackingData || !staffData) return false;

      // Rule: Can't work weekends in both weeks
      if (isWeek2 && week1WeekendStaff.has(staff)) return false;
      // Rule: Can't work twice on the same day
      if (isStaffAssignedToDay(staff, day, currentWeek)) return false;
      // Rule: Must be available based on requests
      if (!isStaffAvailableOnDay(staff, day, staffData.requests || '')) return false;
      // Rule: Max 1 weekend shift per person for the entire 2-week rotation
      if (trackingData.weekendShifts >= 1) return false; 
      
      return true;
  });
}

/**
 * Selects staff for a shift based on fairness criteria.
 * @param {string[]} eligibleStaff
 * @param {number} staffNeeded
 * @param {Object} shiftTracking
 * @param {Object} staffInfo
 * @returns {string[]}
 */
function selectStaffForShift(eligibleStaff, staffNeeded, shiftTracking, staffInfo) {
    if (!eligibleStaff || eligibleStaff.length === 0) {
        return Array(staffNeeded).fill('');
    }
    const selected = [];
    let remaining = [...eligibleStaff];

    while (selected.length < staffNeeded && remaining.length > 0) {
        // Sort by least current rotation shifts, then least historical total shifts, then alphabetically
        remaining.sort((a, b) => {
            const currentTotalA = shiftTracking[a]?.totalShifts ?? Infinity;
            const currentTotalB = shiftTracking[b]?.totalShifts ?? Infinity;
            if (currentTotalA !== currentTotalB) return currentTotalA - currentTotalB;

            const historicalTotalA = staffInfo[a]?.shiftCount ?? Infinity;
            const historicalTotalB = staffInfo[b]?.shiftCount ?? Infinity;
            if (historicalTotalA !== historicalTotalB) return historicalTotalA - historicalTotalB;
            
            return a.localeCompare(b);
        });
        
        const nextStaff = remaining.shift();
        selected.push(nextStaff);
    }
    // Fill remaining spots with empty strings if not enough staff were found
    while (selected.length < staffNeeded) {
        selected.push('');
    }
    return selected;
}

/**
 * Helper to check for adjacent shift conflicts, used in fallback logic.
 * @param {string} staff
 * @param {string} day
 * @param {string} timeSlot
 * @param {Object} currentWeek
 * @returns {boolean} True if there IS a conflict.
 */
function isAdjacentConflict(staff, day, timeSlot, currentWeek) {
    const firstSlotToday = (day === 'Friday' ? WEEKDAY_TIME_SLOTS.FRIDAY_EVENING : WEEKDAY_TIME_SLOTS.REGULAR)[0];
    if (timeSlot === firstSlotToday) {
        const prevDay = getPreviousDay(day);
        if (prevDay) {
             const prevSlots = (prevDay === 'Friday' ? WEEKDAY_TIME_SLOTS.FRIDAY_EVENING : WEEKDAY_TIME_SLOTS.REGULAR);
             const prevLastSlot = prevSlots[prevSlots.length - 1];
             if (currentWeek[prevDay]?.[prevLastSlot]?.staff1 === staff || currentWeek[prevDay]?.[prevLastSlot]?.staff2 === staff) {
                 return true; // Conflict found
             }
        }
    }
    return false; // No conflict
}

/**
 * Centralized basic eligibility checks used across weekday/weekend selection and fallbacks.
 * @param {Object} params
 * @param {string} params.staff
 * @param {string} params.day
 * @param {string} params.timeSlot
 * @param {Object} params.currentWeek
 * @param {string} params.requests
 * @param {boolean} [params.preventSameDay=true]
 * @param {boolean} [params.preventAdjacent=false]
 * @return {boolean}
 */
function basicEligibilityCheck({
    staff,
    day,
    timeSlot,
    currentWeek,
    requests,
    preventSameDay = true,
    preventAdjacent = false
}) {
    if (!isStaffAvailableOnDay(staff, day, requests || '')) return false;
    if (preventSameDay && isStaffAssignedToDay(staff, day, currentWeek)) return false;
    if (preventAdjacent && isAdjacentConflict(staff, day, timeSlot, currentWeek)) return false;
    return true;
}


/**
 * Creates an optimized schedule based on staff information and requirements from settings.
 * INCLUDES ROBUST FALLBACK LOGIC to prevent empty shifts.
 * @param {Object} staffInfo - The staff information object.
 * @param {Object} settings - The loaded settings object.
 * @return {Object} The generated schedule object.
 */
function createOptimizedSchedule(staffInfo, settings) {
    console.log("--- Starting createOptimizedSchedule (with Robust Fallback Logic) ---");
    const schedule = { week1: initializeWeek(), week2: initializeWeek() };
    // Work on a cloned copy to avoid mutating the incoming reference.
    const staffState = safeJsonClone(staffInfo);
    
    // Filter for active staff to build the schedule
    const activeStaff = Object.values(staffState).filter(s => s.status !== settings.ON_LEAVE_STATUS_TEXT);
    const activeStaffNames = activeStaff.map(s => s.name);
    
    if (activeStaffNames.length === 0) {
        throw new Error("No active staff members found to generate a schedule.");
    }

    const totalWeekdayShifts = calculateTotalWeekdayShifts(settings);
    const baseWeekdayShifts = Math.floor(totalWeekdayShifts / activeStaffNames.length);
    const extraShiftsNeeded = totalWeekdayShifts % activeStaffNames.length;
    console.log(`Base Weekday Shifts per Active Staff: ${baseWeekdayShifts}, Extra Shifts to Assign: ${extraShiftsNeeded}`);

    const week1WeekendStaff = new Set();
    const shiftTracking = initializeShiftTracking(staffState);

    // Determine who gets extra shifts based on fairness (lowest rotation count, then lowest historical shifts)
    const staffForExtraShifts = [...activeStaffNames]
        .sort((a, b) => {
            const rotationCountA = staffState[a].rotationCount || 0; 
            const rotationCountB = staffState[b].rotationCount || 0;
            if (rotationCountA !== rotationCountB) return rotationCountA - rotationCountB;
            return (staffState[a].shiftCount || 0) - (staffState[b].shiftCount || 0);
        }).slice(0, extraShiftsNeeded);
        
    staffForExtraShifts.forEach(staff => { if (shiftTracking[staff]) shiftTracking[staff].extraShiftAllowed = true; });
    console.log(`Staff marked for extra weekday shifts (${staffForExtraShifts.length}): ${staffForExtraShifts.join(', ') || 'None'}`);
    
    // --- Main Scheduling Loop ---
    ['week1', 'week2'].forEach((weekKey, weekIndex) => {
        const isWeek2 = (weekIndex === 1);
        console.log(`--- Processing ${weekKey.toUpperCase()} ---`);

        // --- WEEKDAY SHIFTS ---
        WEEKDAYS.forEach(day => {
            const timeSlots = day === 'Friday' ? WEEKDAY_TIME_SLOTS.FRIDAY_EVENING : WEEKDAY_TIME_SLOTS.REGULAR;
            timeSlots.forEach(timeSlot => {
                const staffNeeded = (day === 'Friday' && timeSlot === '4pm - 8pm') ? settings.REQ_FRI_EVE : settings.REQ_WEEKDAY;
                let selectedForSlot = [];

                for (let i = 0; i < staffNeeded; i++) {
                    let selectedName = null;
                    const alreadySelectedInSlot = [...selectedForSlot]; // Staff already picked for this specific slot

                    // === STAGE 1: Standard Eligibility ===
                    let eligiblePool = getEligibleWeekdayStaff(staffState, shiftTracking, day, timeSlot, schedule[weekKey], baseWeekdayShifts, activeStaffNames)
                                         .filter(s => !alreadySelectedInSlot.includes(s));
                    
                    if (eligiblePool.length > 0) {
                        selectedName = selectStaffForShift(eligiblePool, 1, shiftTracking, staffState)[0];
                    }

                    // === STAGE 2: Fallback - Relax Max Shifts Rule ===
                    if (!selectedName) {
                        console.warn(`[FALLBACK-W1] Triggered for ${weekKey} ${day} ${timeSlot}: Relaxing max shift rule.`);
                        eligiblePool = activeStaffNames.filter(staff => {
                            return basicEligibilityCheck({
                                staff,
                                day,
                                timeSlot,
                                currentWeek: schedule[weekKey],
                                requests: staffState[staff].requests,
                                preventSameDay: true,
                                preventAdjacent: true
                            }) && !alreadySelectedInSlot.includes(staff);
                        });
                        if (eligiblePool.length > 0) {
                           selectedName = selectStaffForShift(eligiblePool, 1, shiftTracking, staffState)[0];
                        }
                    }

                    // === STAGE 3: Fallback - Relax Same Day Rule ===
                    if (!selectedName) {
                        console.warn(`[FALLBACK-W2] Triggered for ${weekKey} ${day} ${timeSlot}: Relaxing same-day assignment rule.`);
                         eligiblePool = activeStaffNames.filter(staff => {
                            return basicEligibilityCheck({
                                staff,
                                day,
                                timeSlot,
                                currentWeek: schedule[weekKey],
                                requests: staffState[staff].requests,
                                preventSameDay: false,
                                preventAdjacent: true
                            }) && !alreadySelectedInSlot.includes(staff);
                        });
                        if (eligiblePool.length > 0) {
                           selectedName = selectStaffForShift(eligiblePool, 1, shiftTracking, staffState)[0];
                        }
                    }
                    
                    // Assign if found, and update tracking
                    if (selectedName) {
                        selectedForSlot.push(selectedName);
                        if (shiftTracking[selectedName]) {
                            shiftTracking[selectedName].weekdayShifts++;
                            shiftTracking[selectedName].totalShifts++;
                        }
                    } else {
                        console.error(`[FINAL FAILURE] Could not find any staff for ${weekKey} ${day} ${timeSlot}. The slot will be empty.`);
                    }
                } // End loop for staffNeeded

                // Update schedule object for the slot
                schedule[weekKey][day][timeSlot] = { staff1: selectedForSlot[0] || '', staff2: selectedForSlot[1] || '' };

            });
        });

        // --- WEEKEND SHIFTS ---
        WEEKEND_DAYS.forEach(day => {
            const timeSlots = day === 'Saturday' ? WEEKEND_TIME_SLOTS.SATURDAY : WEEKEND_TIME_SLOTS.SUNDAY;
            timeSlots.forEach(timeSlot => {
                const staffNeeded = (day === 'Saturday') ? settings.REQ_SATURDAY : settings.REQ_SUNDAY;
                let selectedForSlot = [];
                
                for (let i = 0; i < staffNeeded; i++) {
                    let selectedName = null;
                    const alreadySelectedInSlot = [...selectedForSlot];

                    // === STAGE 1: Standard Eligibility ===
                    let eligiblePool = getEligibleWeekendStaff(staffState, shiftTracking, day, timeSlot, schedule[weekKey], week1WeekendStaff, isWeek2)
                                         .filter(s => !alreadySelectedInSlot.includes(s));

                    if (eligiblePool.length > 0) {
                        selectedName = selectStaffForShift(eligiblePool, 1, shiftTracking, staffState)[0];
                    }

                    // === STAGE 2: Fallback (Week 2 only) - Relax "No Repeat Weekend" Rule ===
                    if (!selectedName && isWeek2) {
                        console.warn(`[FALLBACK-E1] Triggered for ${weekKey} ${day} ${timeSlot}: Allowing staff from Week 1 to work again.`);
                        // Get eligible staff, but IGNORE the week1WeekendStaff check
                        eligiblePool = activeStaffNames.filter(staff => {
                            return basicEligibilityCheck({
                                staff,
                                day,
                                timeSlot,
                                currentWeek: schedule[weekKey],
                                requests: staffState[staff].requests,
                                preventSameDay: true,
                                preventAdjacent: false
                            }) &&
                            (shiftTracking[staff].weekendShifts < 1) &&
                            !alreadySelectedInSlot.includes(staff);
                        });
                        if (eligiblePool.length > 0) {
                            selectedName = selectStaffForShift(eligiblePool, 1, shiftTracking, staffState)[0];
                        }
                    }
                    
                    // === STAGE 3: Fallback (All Weekends) - Relax "Max 1 Weekend Shift" Rule ===
                    if (!selectedName) {
                         console.warn(`[FALLBACK-E2] Triggered for ${weekKey} ${day} ${timeSlot}: Allowing staff to work a second weekend shift.`);
                         eligiblePool = activeStaffNames.filter(staff => {
                             // Ignore week1 check AND max weekend shifts check
                            return basicEligibilityCheck({
                                staff,
                                day,
                                timeSlot,
                                currentWeek: schedule[weekKey],
                                requests: staffState[staff].requests,
                                preventSameDay: true,
                                preventAdjacent: false
                            }) && !alreadySelectedInSlot.includes(staff);
                         });
                         if (eligiblePool.length > 0) {
                            selectedName = selectStaffForShift(eligiblePool, 1, shiftTracking, staffState)[0];
                        }
                    }

                    // Assign if found, and update tracking
                    if (selectedName) {
                        selectedForSlot.push(selectedName);
                        if (shiftTracking[selectedName]) {
                            shiftTracking[selectedName].weekendShifts++; 
                            shiftTracking[selectedName].totalShifts++;
                            staffState[selectedName].lastWeekendShiftDate = toIsoDateString();
                            if (weekIndex === 0) { 
                                week1WeekendStaff.add(selectedName); 
                            }
                        }
                    } else {
                        console.error(`[FINAL FAILURE] Could not find any staff for ${weekKey} ${day} ${timeSlot}. The slot will be empty.`);
                    }
                } // End loop for staffNeeded

                // Update schedule object for the slot
                schedule[weekKey][day][timeSlot] = { staff1: selectedForSlot[0] || '', staff2: selectedForSlot[1] || '' };
            });
        });
    });

    // Final update of staffInfo object with historical and rotation counts
    Object.keys(shiftTracking).forEach(staff => {
        if (staffState[staff]) {
            staffState[staff].weekdayShifts += shiftTracking[staff].weekdayShifts;
            staffState[staff].weekendShifts += shiftTracking[staff].weekendShifts;
            
            const didTheyGetExtra = shiftTracking[staff].extraShiftAllowed && (shiftTracking[staff].weekdayShifts > baseWeekdayShifts);
            if (didTheyGetExtra) {
                 staffState[staff].rotationCount++;
                 staffState[staff].lastExtraShiftDate = toIsoDateString();
            }
            staffState[staff].shiftCount = staffState[staff].weekdayShifts + staffState[staff].weekendShifts;
        }
    });

    console.log("--- Finished createOptimizedSchedule ---");
    return { schedule, staffInfo: staffState };
}

// =================================================================
// END OF SECTION 2
// =================================================================

// =================================================================
// SECTION 3: SCHEDULE APPLICATION, STATISTICS, UI (FINAL v2 - Fixes)
// =================================================================

/**
 * Applies the generated schedule to the 'Schedule' sheet using batch writes for performance.
 * Handles formatting after data is written. Ensures all rows have the same number of columns.
 * Uses settings for sheet name. Ensures Sunday now shows 2 staff members.
 * @param {Object} newSchedule - The generated schedule object ({ week1: {...}, week2: {...} }).
 * @param {Object} settings - The loaded settings object.
 */
function applyScheduleToSheet(newSchedule, settings) {
  console.log("--- Starting applyScheduleToSheet (Batch Mode) ---");
  const sheetName = settings.SCHEDULE_SHEET;
  let sheet;
  try {
    sheet = getSheetOrThrow(sheetName, 'applyScheduleToSheet');
  } catch (e) {
    const errorMsg = e.message || `Schedule sheet "${sheetName}" not found. Cannot apply schedule.`;
    console.error(errorMsg);
    logAction('Apply Schedule Error', errorMsg, 'ERROR', settings);
    throw new Error(errorMsg);
  }

  const maxDataColumns = 1 + WEEKDAYS.length;
  const outputData = [];
  const formatting = { merges: [], backgrounds: [], fontWeights: [], fontSizes: [], alignments: [], borders: [] };

  function padRow(rowData, length) {
    const currentLength = rowData.length;
    if (currentLength < length) {
      for (let i = currentLength; i < length; i++) { rowData.push(''); }
    }
    return rowData.slice(0, length);
  }

  function addFormat(type, details) {
      if (!formatting[type]) formatting[type] = [];
      formatting[type].push(details);
  }

  let currentRow = 1;

  outputData.push(padRow(['UPDATED NEW SCHEDULE'], maxDataColumns));
  addFormat('fontWeights', { rangeA1: `A${currentRow}`, weight: 'bold' });
  addFormat('fontSizes', { rangeA1: `A${currentRow}`, size: 14 });
  addFormat('merges', `A${currentRow}:${String.fromCharCode(64 + maxDataColumns)}${currentRow}`);
  addFormat('alignments', { rangeA1: `A${currentRow}`, horizontal: 'center'});
  currentRow++;
  outputData.push(padRow([], maxDataColumns)); currentRow++;

  ['week1', 'week2'].forEach((weekKey, weekIndex) => {
    const weekData = newSchedule[weekKey];
     if (!weekData) {
        console.warn(`applyScheduleToSheet: Data for ${weekKey} missing.`);
        outputData.push(padRow([`${weekKey.toUpperCase()} - DATA MISSING`], maxDataColumns));
        const missingWeekHeaderRow = currentRow; const missingWeekHeaderRange = `A${missingWeekHeaderRow}:${String.fromCharCode(64 + maxDataColumns)}${missingWeekHeaderRow}`;
        addFormat('merges', missingWeekHeaderRange); addFormat('alignments', { rangeA1: missingWeekHeaderRange, horizontal: 'center'}); addFormat('fontWeights', { rangeA1: missingWeekHeaderRange, weight: 'bold' });
        currentRow++; outputData.push(padRow([], maxDataColumns)); currentRow++; return;
    }

    outputData.push(padRow([`Week ${weekIndex + 1}`], maxDataColumns));
    const weekHeaderRow = currentRow; const weekHeaderFormatRange = `A${weekHeaderRow}:${String.fromCharCode(64 + maxDataColumns)}${weekHeaderRow}`;
    addFormat('fontWeights', { rangeA1: weekHeaderFormatRange, weight: 'bold' }); addFormat('backgrounds', { rangeA1: weekHeaderFormatRange, color: '#e2e8f0' }); addFormat('merges', weekHeaderFormatRange); addFormat('alignments', { rangeA1: weekHeaderFormatRange, horizontal: 'center' });
    currentRow++; outputData.push(padRow([], maxDataColumns)); currentRow++;

    outputData.push(padRow(['Weekday Schedule'], maxDataColumns));
    const weekdaySubHeaderRow = currentRow; const weekdaySubHeaderRange = `A${weekdaySubHeaderRow}:${String.fromCharCode(64 + maxDataColumns)}${weekdaySubHeaderRow}`;
    addFormat('fontWeights', { rangeA1: `A${weekdaySubHeaderRow}`, weight: 'bold' }); addFormat('backgrounds', { rangeA1: `A${weekdaySubHeaderRow}`, color: '#f1f5f9' }); addFormat('merges', weekdaySubHeaderRange); addFormat('alignments', { rangeA1: `A${weekdaySubHeaderRow}`, horizontal: 'center' });
    currentRow++;

    const weekdayHeaderRowData = ['Time Slot', ...WEEKDAYS];
    outputData.push(padRow(weekdayHeaderRowData, maxDataColumns));
    const weekdayHeaderSheetRow = currentRow; const weekdayHeaderRange = `A${weekdayHeaderSheetRow}:${String.fromCharCode(64 + maxDataColumns)}${weekdayHeaderSheetRow}`;
    addFormat('backgrounds', { rangeA1: weekdayHeaderRange, color: '#f8fafc' }); addFormat('fontWeights', { rangeA1: weekdayHeaderRange, weight: 'bold' }); addFormat('alignments', { rangeA1: `A${weekdayHeaderSheetRow}`, horizontal: 'center'});
     WEEKDAYS.forEach((_, index) => { addFormat('alignments', { rangeA1: `${String.fromCharCode(66 + index)}${weekdayHeaderSheetRow}`, horizontal: 'center'}); });
    currentRow++;

    const allWeekdayTimeSlots = WEEKDAY_TIME_SLOTS.REGULAR; // This includes '4pm - 8pm'
    allWeekdayTimeSlots.forEach(timeSlot => {
      const staff1Row = [timeSlot]; const staff2Row = [''];
      WEEKDAYS.forEach(day => {
        const assignment = weekData[day]?.[timeSlot] || { staff1: '', staff2: '' };
        staff1Row.push(assignment.staff1 || '-');
        
        // This logic correctly uses settings.REQ_FRI_EVE to determine if staff2 should be N/A or the actual staff name
        const staff2Value = (day === 'Friday' && timeSlot === '4pm - 8pm' && settings.REQ_FRI_EVE < 2) ? '-' : (assignment.staff2 || '-');
        staff2Row.push(staff2Value);
      });
      outputData.push(padRow(staff1Row, maxDataColumns));
      const staff1SheetRow = currentRow; addFormat('backgrounds', { rangeA1: `A${staff1SheetRow}`, color: '#f1f5f9' }); addFormat('alignments', { rangeA1: `A${staff1SheetRow}`, horizontal: 'center'});
       WEEKDAYS.forEach((_, index) => { addFormat('alignments', { rangeA1: `${String.fromCharCode(66 + index)}${staff1SheetRow}`, horizontal: 'center'}); });
      currentRow++;
      outputData.push(padRow(staff2Row, maxDataColumns));
      const staff2SheetRow = currentRow;
       WEEKDAYS.forEach((_, index) => { addFormat('alignments', { rangeA1: `${String.fromCharCode(66 + index)}${staff2SheetRow}`, horizontal: 'center'}); });
      currentRow++;
      const timeSlotMergeRange = `A${staff1SheetRow}:A${staff2SheetRow}`; addFormat('merges', timeSlotMergeRange); addFormat('alignments', { rangeA1: `A${staff1SheetRow}`, vertical: 'middle'});
    });
    outputData.push(padRow([], maxDataColumns)); currentRow++;

    outputData.push(padRow(['Weekend Schedule'], maxDataColumns));
    const weekendSubHeaderRow = currentRow; const weekendSubHeaderRange = `A${weekendSubHeaderRow}:${String.fromCharCode(64 + maxDataColumns)}${weekendSubHeaderRow}`;
    addFormat('fontWeights', { rangeA1: `A${weekendSubHeaderRow}`, weight: 'bold' }); addFormat('backgrounds', { rangeA1: `A${weekendSubHeaderRow}`, color: '#f1f5f9' }); addFormat('merges', weekendSubHeaderRange); addFormat('alignments', { rangeA1: `A${weekendSubHeaderRow}`, horizontal: 'center' });
    currentRow++;

    const weekendHeaderRowData = ['Time Slot', 'Staff', 'Day'];
    outputData.push(padRow(weekendHeaderRowData, maxDataColumns)); // Corrected: maxDataColumns, not 3
    const weekendHeaderSheetRow = currentRow; const weekendHeaderRange = `A${weekendHeaderSheetRow}:C${weekendHeaderSheetRow}`; // Header only spans 3 cols
    addFormat('backgrounds', { rangeA1: weekendHeaderRange, color: '#f8fafc' }); addFormat('fontWeights', { rangeA1: weekendHeaderRange, weight: 'bold' }); addFormat('alignments', { rangeA1: `A${weekendHeaderSheetRow}:C${weekendHeaderSheetRow}`, horizontal: 'center'});
    currentRow++;

    // CORRECTED: Both Saturday and Sunday now use the same logic to display 2 staff members
    WEEKEND_DAYS.forEach(day => {
        const timeSlots = (day === 'Saturday') ? WEEKEND_TIME_SLOTS.SATURDAY : WEEKEND_TIME_SLOTS.SUNDAY;
        const startRowForMerges = currentRow;
        timeSlots.forEach((timeSlot, index) => {
            const assignment = weekData[day]?.[timeSlot] || { staff1: '', staff2: '' };
            // Ensure rowData is padded to maxDataColumns for sheet consistency, even though weekend data is only in first 3.
            const rowDataStaff1 = padRow([(index === 0 ? timeSlot : ''), assignment.staff1 || '-', (index === 0 ? day : '')], maxDataColumns);
            const rowDataStaff2 = padRow(['', assignment.staff2 || '-', ''], maxDataColumns);
            
            outputData.push(rowDataStaff1);
            const staff1SheetRow = currentRow; 
            if (index === 0) addFormat('backgrounds', { rangeA1: `A${staff1SheetRow}`, color: '#f1f5f9' }); 
            addFormat('alignments', { rangeA1: `A${staff1SheetRow}:C${staff1SheetRow}`, horizontal: 'center'}); // Align content in first 3 cols
            currentRow++;
            
            outputData.push(rowDataStaff2);
            const staff2SheetRow = currentRow; 
            addFormat('alignments', { rangeA1: `B${staff2SheetRow}:C${staff2SheetRow}`, horizontal: 'center'}); // Align content in first 3 cols for staff2
            currentRow++;
        });
        const endRowForMerges = currentRow - 1;
        if (endRowForMerges >= startRowForMerges && timeSlots.length > 0) {
            addFormat('merges', `A${startRowForMerges}:A${endRowForMerges}`); 
            addFormat('alignments', { rangeA1: `A${startRowForMerges}`, vertical: 'middle'});
            addFormat('merges', `C${startRowForMerges}:C${endRowForMerges}`); 
            addFormat('alignments', { rangeA1: `C${startRowForMerges}`, vertical: 'middle'});
        }
    });

    outputData.push(padRow([], maxDataColumns)); currentRow++;
    if (weekIndex === 0) { outputData.push(padRow([], maxDataColumns)); currentRow++; outputData.push(padRow([], maxDataColumns)); currentRow++; }
  }); // End week loop

  sheet.clearContents().clearFormats();

  if (outputData.length > 0) {
    const requiredRows = outputData.length; const requiredCols = maxDataColumns;
    if (requiredRows > sheet.getMaxRows()) { sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows()); }
    if (requiredCols > sheet.getMaxColumns()) { sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredCols - sheet.getMaxColumns()); }
    // Only delete rows/cols if current max exceeds required AND required is > 0
    if (sheet.getMaxRows() > requiredRows && requiredRows > 0) { sheet.deleteRows(requiredRows + 1, sheet.getMaxRows() - requiredRows); }
    if (sheet.getMaxColumns() > requiredCols && requiredCols > 0) { sheet.deleteColumns(requiredCols + 1, sheet.getMaxColumns() - requiredCols); }


    const dataRange = sheet.getRange(1, 1, requiredRows, requiredCols); dataRange.setValues(outputData);
    console.log("Applying formatting...");
    formatting.merges.forEach(rangeA1 => { try { sheet.getRange(rangeA1).merge(); } catch(e){ console.error(`Merge fail: ${rangeA1}: ${e}`)} });
    formatting.backgrounds.forEach(f => { try { sheet.getRange(f.rangeA1).setBackground(f.color); } catch(e){ console.error(`BG fail: ${f.rangeA1}: ${e}`)} });
    formatting.fontWeights.forEach(f => { try { sheet.getRange(f.rangeA1).setFontWeight(f.weight); } catch(e){ console.error(`Weight fail: ${f.rangeA1}: ${e}`)} });
    formatting.fontSizes.forEach(f => { try { if (f.size) sheet.getRange(f.rangeA1).setFontSize(f.size); } catch(e){ console.error(`Size fail: ${f.rangeA1}: ${e}`)} });
    formatting.alignments.forEach(f => { try { const range = sheet.getRange(f.rangeA1); if (f.horizontal) range.setHorizontalAlignment(f.horizontal); if (f.vertical) range.setVerticalAlignment(f.vertical); } catch(e){ console.error(`Align fail: ${f.rangeA1}: ${e}`)} });

     try {
        let firstHeaderRowIndex = -1;
        for(let r=0; r<outputData.length; r++){ if(outputData[r][0] === 'Time Slot' && outputData[r][1] === 'Monday') { firstHeaderRowIndex = r; break; } }
        if (firstHeaderRowIndex !== -1) {
            const firstDataRow = firstHeaderRowIndex + 2; // Start borders below header and the first staff1 row
            const lastDataRow = outputData.length;
            if (lastDataRow >= firstDataRow) {
                const borderRange = sheet.getRange(firstDataRow -1, 1, lastDataRow - (firstDataRow -1) + 1, maxDataColumns); // Adjusted to include the staff1 row
                borderRange.setBorder(true, true, true, true, true, true, '#d1d5db', SpreadsheetApp.BorderStyle.SOLID_THIN);
            } else { console.warn("Border application range invalid."); }
        } else { console.warn("Could not find start of data area for borders."); }
     } catch (e) { console.error(`Border application fail: ${e}`)}

    try {
        sheet.setColumnWidth(1, 120); // Time Slot / Day
        for (let i = 2; i <= maxDataColumns; i++) { sheet.setColumnWidth(i, 150); } // Weekday names / Staff
        // If weekend data is in col B & C (cols 2 & 3 one-based)
        if (maxDataColumns >= 3) { // Ensure at least 3 columns exist
           sheet.setColumnWidth(2, 150); // Staff (for weekend)
           sheet.setColumnWidth(3, 100); // Day (for weekend) or Wednesday
        }

    } catch (e) { console.error(`Column width fail: ${e}`)}

    SpreadsheetApp.flush();
    console.log("--- Finished applyScheduleToSheet ---");
  } else {
      console.warn(`applyScheduleToSheet: No data generated for "${sheetName}". Sheet cleared.`);
      logAction('Apply Schedule Warning', `No schedule data generated for "${sheetName}". Sheet was cleared.`, 'WARN', settings);
  }
}


/**
 * Shows a preview of the schedule in a HTML modal dialog.
 * CORRECTED: Sunday now renders 2 staff slots, like Saturday.
 * @param {Object} newSchedule - The generated schedule object.
 * @param {Object} settings - The loaded settings object.
 */
function showSchedulePreview(newSchedule, settings) {
  if (!newSchedule || (!newSchedule.week1 && !newSchedule.week2)) {
       SpreadsheetApp.getUi().alert("Cannot show preview: No schedule data generated.");
       return;
  }

  const htmlTemplateString = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js" integrity="sha512-qZvrmS2ekKPF2mSznTQsxqPgnpkI4DNTlrdUmTzrDgektczlKNRRhy5X5AAOnx5S09ydFYWWNSfcEqDTTHgtNA==" crossorigin="anonymous" referrerpolicy="no-referrer"></script>
      <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js" integrity="sha512-BNaRQnYJYiPSqHHDb58B0yaPfCu+Wgds8Gp/gU33kqBtgNS4tSPHuGibyoeqMV/TJlSKda6FXzoEyYGjTe+vXA==" crossorigin="anonymous" referrerpolicy="no-referrer"></script>
      <style>
        /* --- Styles (same as before) --- */
        * { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif; box-sizing: border-box; }
        body { padding: 20px; background-color: #f5f7fa; line-height: 1.4; color: #333; }
        h2 { color: #1a1a1a; font-weight: 600; margin-bottom: 24px; text-align: center; font-size: 1.5em; }
        h3 { color: #2c3e50; font-weight: 600; margin-top: 32px; margin-bottom: 16px; border-bottom: 1px solid #e2e8f0; padding-bottom: 6px; font-size: 1.2em; }
        h4 { color: #34495e; font-weight: 600; margin-top: 24px; margin-bottom: 12px; font-size: 1em; }
        .schedule-container { background: white; border-radius: 8px; box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1); padding: 24px; margin-bottom: 24px; overflow-x: auto; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; border-radius: 6px; overflow: hidden; border: 1px solid #e2e8f0; }
        th, td { border: 1px solid #e2e8f0; padding: 10px 12px; text-align: center; font-size: 13px; vertical-align: middle; }
        th { background-color: #f8fafc; font-weight: 600; color: #475569; text-transform: uppercase; font-size: 0.85rem; letter-spacing: 0.03em; }
        td { background-color: #ffffff; }
        tr:nth-child(even) td { background-color: #fcfcfd; }
        td.time-slot-label, th.time-slot-header { background-color: #f1f5f9; font-weight: 500; color: #475569; text-align: center; vertical-align: middle; }
        td.day-label { text-align: center; vertical-align: middle; font-weight: 500; background-color: #f8fafc; }
        .staff-name { color: #1d4ed8; font-weight: 500; }
        .empty-slot { color: #94a3b8; font-style: italic; }
        .not-applicable { color: #cbd5e1; font-style: italic; font-size: 0.9em;}
        .week-header { background-color: #e2e8f0; font-weight: bold; text-align: center; padding: 10px; margin-bottom: 15px; border-radius: 4px; }
        .button-container { display: flex; justify-content: flex-end; margin-bottom: 20px; gap: 10px; }
        .download-btn, .close-btn { background-color: #2563eb; color: white; padding: 10px 20px; border: none; border-radius: 6px; cursor: pointer; font-weight: 500; transition: background-color 0.2s ease; font-size: 13px; }
        .download-btn:hover { background-color: #1d4ed8; }
        .download-btn:disabled { background-color: #94a3b8; cursor: not-allowed; opacity: 0.7; }
        .close-btn { background-color: #64748b; }
        .close-btn:hover { background-color: #475569; }
      </style>
    </head>
    <body>
      <div id="schedule-content">
        <div class="button-container">
          <button id="pdf-button" class="download-btn" onclick="downloadPDF()">Download PDF</button>
          <button class="close-btn" onclick="google.script.host.close()">Close</button>
        </div>
        <h2>Preview of New Schedule</h2>

        <? /* --- Template Scriptlets Start --- */ ?>
        <? if (typeof newSchedule !== 'undefined' && newSchedule) { ?>
          <? Object.keys(newSchedule).sort().forEach(function(weekKey) { ?>
            <? var weekData = newSchedule[weekKey]; ?>
             <? if (!weekData) return; /* Skip rendering if week data is missing */ ?>
            <div class="schedule-container">
              <div class="week-header"><?= weekKey.charAt(0).toUpperCase() + weekKey.slice(1).replace('week', 'Week ') ?></div>
              <h4>Weekday Schedule</h4>
              <table>
                <thead>
                  <tr>
                    <th class="time-slot-header">Time</th>
                    <? if (typeof WEEKDAYS !== 'undefined') { WEEKDAYS.forEach(function(day) { ?> <th><?= day ?></th> <? }); } ?>
                  </tr>
                </thead>
                <tbody>
                <? if (typeof WEEKDAY_TIME_SLOTS !== 'undefined' && WEEKDAY_TIME_SLOTS.REGULAR) { ?>
                  <? WEEKDAY_TIME_SLOTS.REGULAR.forEach(function(timeSlot) { ?>
                    <tr>
                      <td rowspan="2" class="time-slot-label"><?= timeSlot ?></td>
                      <? /* Staff 1 row */ ?>
                      <? WEEKDAYS.forEach(function(day) { ?>
                        <? var assignment = weekData?.[day]?.[timeSlot] || {}; ?>
                        <td class="<?= assignment.staff1 ? 'staff-name' : 'empty-slot' ?>"><?= assignment.staff1 || '-' ?></td>
                      <? }); ?>
                    </tr>
                    <tr>
                      <? /* Staff 2 row */ ?>
                      <? WEEKDAYS.forEach(function(day) { ?>
                        <? var assignment = weekData?.[day]?.[timeSlot] || {}; ?>
                        <? var isFriEveSingle = (day === 'Friday' && timeSlot === '4pm - 8pm' && settings.REQ_FRI_EVE < 2); ?>
                        <? var staff2ToDisplay = isFriEveSingle ? undefined : assignment.staff2; ?>
                        <td class="<?= staff2ToDisplay ? 'staff-name' : (staff2ToDisplay === undefined ? 'not-applicable' : 'empty-slot') ?>">
                          <?= staff2ToDisplay || (staff2ToDisplay === undefined ? 'N/A' : '-') ?>
                        </td>
                      <? }); ?>
                    </tr>
                  <? }); ?>
                <? } else { ?>
                  <tr><td colspan="<?= (typeof WEEKDAYS !== 'undefined' ? WEEKDAYS.length : 0) + 1 ?>">Error: Weekday slots data missing.</td></tr>
                <? } ?>
                </tbody>
              </table>

              <h4>Weekend Schedule</h4>
              <table>
                 <thead>
                  <tr> <th class="time-slot-header">Time Slot</th> <th>Staff</th> <th>Day</th> </tr>
                </thead>
                <tbody>
                  <!-- Saturday -->
                  <? if (typeof WEEKEND_TIME_SLOTS !== 'undefined' && WEEKEND_TIME_SLOTS.SATURDAY && WEEKEND_TIME_SLOTS.SATURDAY.length > 0) { var satSlots = WEEKEND_TIME_SLOTS.SATURDAY; satSlots.forEach(function(timeSlot, index) { var assignment = weekData?.['Saturday']?.[timeSlot] || {}; ?>
                       <tr> <? if(index === 0) { ?> <td rowspan="<?= satSlots.length * 2 ?>" class="time-slot-label"><?= satSlots.join('<br>') ?></td> <? } ?> <td class="<?= assignment.staff1 ? 'staff-name' : 'empty-slot' ?>"><?= assignment.staff1 || '-' ?></td> <? if(index === 0) { ?> <td rowspan="<?= satSlots.length * 2 ?>" class="day-label">Saturday</td> <? } ?> </tr>
                       <tr> <td class="<?= assignment.staff2 ? 'staff-name' : 'empty-slot' ?>"><?= assignment.staff2 || '-' ?></td> </tr>
                  <? }); } else { ?> <tr><td colspan="3">No Saturday slots defined or data missing.</td></tr> <? } ?>
                  
                  <!-- Sunday - CORRECTED to show 2 staff -->
                   <? if (typeof WEEKEND_TIME_SLOTS !== 'undefined' && WEEKEND_TIME_SLOTS.SUNDAY && WEEKEND_TIME_SLOTS.SUNDAY.length > 0) { var sunSlots = WEEKEND_TIME_SLOTS.SUNDAY; sunSlots.forEach(function(timeSlot, index) { var assignment = weekData?.['Sunday']?.[timeSlot] || {}; ?>
                      <tr> <? if(index === 0) { ?> <td rowspan="<?= sunSlots.length * 2 ?>" class="time-slot-label"><?= sunSlots.join('<br>') ?></td> <? } ?> <td class="<?= assignment.staff1 ? 'staff-name' : 'empty-slot' ?>"><?= assignment.staff1 || '-' ?></td> <? if(index === 0) { ?> <td rowspan="<?= sunSlots.length * 2 ?>" class="day-label">Sunday</td> <? } ?> </tr>
                      <tr> <td class="<?= assignment.staff2 ? 'staff-name' : 'empty-slot' ?>"><?= assignment.staff2 || '-' ?></td> </tr>
                  <? }); } else { ?> <tr><td colspan="3">No Sunday slots defined or data missing.</td></tr> <? } ?>
                </tbody>
              </table>
            </div>
          <? }); ?>
        <? } else { ?> <p style="text-align:center; padding: 20px;">No schedule data available.</p> <? } ?>
        <? /* --- Template Scriptlets End --- */ ?>
      </div>

      <script>
        // PDF download script is unchanged and should work as is.
        async function downloadPDF() {
          const content = document.getElementById('schedule-content'); const btn = document.getElementById('pdf-button'); const closeBtn = btn?.nextElementSibling;
          if (!btn) { console.error("PDF Button not found"); return; }
          if (typeof html2canvas === 'undefined' || typeof window.jspdf === 'undefined') { alert("PDF generation libraries did not load correctly."); return; }
          const { jsPDF } = window.jspdf;
          try {
            btn.textContent = 'Generating PDF...'; btn.disabled = true; if (closeBtn) closeBtn.disabled = true;
            const buttonsContainer = content.querySelector('.button-container'); const originalDisplay = buttonsContainer ? buttonsContainer.style.display : 'flex'; if(buttonsContainer) buttonsContainer.style.display = 'none';
            const canvas = await html2canvas(content, { scale: 2, useCORS: true, backgroundColor: '#ffffff', logging: false, windowWidth: content.scrollWidth, windowHeight: content.scrollHeight });
            if(buttonsContainer) buttonsContainer.style.display = originalDisplay;
            const imgData = canvas.toDataURL('image/png'); const pdfWidth = 297 - 20; const pdfHeight = 210 - 20; const imgWidth = pdfWidth; const imgHeight = (canvas.height * imgWidth) / canvas.width; let heightLeft = imgHeight; const pdf = new jsPDF('landscape', 'mm', 'a4'); let position = 10;
            pdf.addImage(imgData, 'PNG', 10, position, imgWidth, imgHeight); heightLeft -= pdfHeight;
            while (heightLeft > 0) { position = heightLeft - imgHeight + 10; pdf.addPage(); pdf.addImage(imgData, 'PNG', 10, position, imgWidth, imgHeight); heightLeft -= pdfHeight; }
            pdf.save('schedule_preview.pdf');
          } catch (error) { console.error('PDF generation failed:', error); alert('Failed to generate PDF.');
          } finally { btn.textContent = 'Download PDF'; btn.disabled = false; if (closeBtn) closeBtn.disabled = false; }
        }
      </script>
    </body>
    </html>
  `;

  const template = HtmlService.createTemplate(htmlTemplateString);
  template.newSchedule = newSchedule;
  template.settings = settings; // Pass settings to template
  template.WEEKDAYS = WEEKDAYS; 
  template.WEEKEND_DAYS = WEEKEND_DAYS;
  template.WEEKDAY_TIME_SLOTS = WEEKDAY_TIME_SLOTS; 
  template.WEEKEND_TIME_SLOTS = WEEKEND_TIME_SLOTS;

  const html = template.evaluate() .setWidth(1200) .setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, 'Schedule Preview');
}


/**
 * Calculates schedule statistics directly from a given schedule object.
 * CORRECTED: Now properly counts staff2 for both Saturday and Sunday.
 */
function getCurrentScheduleStats(schedule, staffInfo, settings) {
    console.log("--- Calculating Statistics from Schedule Object ---");
    const stats = { playerStats: {}, totalShifts: 0, totalWeekdayShifts: 0, totalWeekendShifts: 0, timeOfDayShifts: { Morning: 0, Afternoon: 0, Evening: 0, WeekendMorning: 0, WeekendAfternoon: 0 }, highestShifts: { staff: 'N/A', count: 0 }, lowestShifts: { staff: 'N/A', count: Infinity }, averageShiftsPerPlayer: 0, weekdayPercentage: 0, weekendPercentage: 0, imbalances: [], errors: [] };
    const onLeaveStatusText = settings.ON_LEAVE_STATUS_TEXT || 'On Leave';
    const allStaffNames = Object.keys(staffInfo || {});
    if (allStaffNames.length === 0) { stats.errors.push("No staff information provided."); return stats; }

    allStaffNames.forEach(name => { if (staffInfo[name]) { stats.playerStats[name] = { totalShifts: 0, weekdayShifts: 0, weekendShifts: 0, utilizationPercentage: 0, level: staffInfo[name].level || 'N/A', region: staffInfo[name].region || 'N/A', status: staffInfo[name].status || 'N/A' }; } });
    
    if (!schedule || typeof schedule !== 'object') { stats.errors.push("Invalid schedule data object."); return stats; }
    
    const incrementShift = (staffName, isWeekend) => {
      if (!staffName || staffName === '-') return; 
      if (!stats.playerStats[staffName]) { 
        console.warn(`Stats Calc: Staff "${staffName}" found in schedule but missing from Staff Info. Adding to stats.`); 
        stats.playerStats[staffName] = { totalShifts: 0, weekdayShifts: 0, weekendShifts: 0, utilizationPercentage: 0, level: 'N/A', region: 'N/A', status: 'Unknown' }; 
      }
      stats.playerStats[staffName].totalShifts++; 
      stats.totalShifts++; 
      if (isWeekend) { 
        stats.playerStats[staffName].weekendShifts++; stats.totalWeekendShifts++; 
      } else { 
        stats.playerStats[staffName].weekdayShifts++; stats.totalWeekdayShifts++; 
      } 
    };
    
    ['week1', 'week2'].forEach(weekKey => { 
        const weekData = schedule[weekKey]; 
        if (!weekData || typeof weekData !== 'object') return; 

        WEEKDAYS.forEach(day => { 
            const dayData = weekData[day]; 
            if (!dayData || typeof dayData !== 'object') return; 
            const timeSlots = (day === 'Friday') ? WEEKDAY_TIME_SLOTS.FRIDAY_EVENING : WEEKDAY_TIME_SLOTS.REGULAR; 
            timeSlots.forEach(timeSlot => { 
                const assignment = dayData[timeSlot]; 
                if (!assignment || typeof assignment !== 'object') return; 
                incrementShift(assignment.staff1, false); 
                incrementShift(assignment.staff2, false); 
                if (assignment.staff1 && assignment.staff1 !== '-') { 
                    if (timeSlot.includes('8am')) stats.timeOfDayShifts.Morning++; 
                    else if (timeSlot.includes('12pm')) stats.timeOfDayShifts.Afternoon++; 
                    else if (timeSlot.includes('4pm')) stats.timeOfDayShifts.Evening++; 
                } 
                if (assignment.staff2 && assignment.staff2 !== '-') { // This correctly counts staff2 for Friday evening if present
                    if (timeSlot.includes('8am')) stats.timeOfDayShifts.Morning++; 
                    else if (timeSlot.includes('12pm')) stats.timeOfDayShifts.Afternoon++; 
                    else if (timeSlot.includes('4pm')) stats.timeOfDayShifts.Evening++;
                }
            }); 
        }); 

        WEEKEND_DAYS.forEach(day => { 
            const dayData = weekData[day]; 
            if (!dayData || typeof dayData !== 'object') return; 
            const timeSlots = (day === 'Saturday') ? WEEKEND_TIME_SLOTS.SATURDAY : WEEKEND_TIME_SLOTS.SUNDAY; 
            timeSlots.forEach(timeSlot => { 
                const assignment = dayData[timeSlot]; 
                if (!assignment || typeof assignment !== 'object') return; 
                incrementShift(assignment.staff1, true); 
                incrementShift(assignment.staff2, true); // CORRECTED: Always check for staff2 for all weekend days
                if (assignment.staff1 && assignment.staff1 !== '-') { 
                    if (timeSlot.includes('9am')) stats.timeOfDayShifts.WeekendMorning++; 
                    else if (timeSlot.includes('1pm')) stats.timeOfDayShifts.WeekendAfternoon++; 
                } 
                if (assignment.staff2 && assignment.staff2 !== '-') { // CORRECTED: Count staff2 for time-of-day stats for all weekend days
                    if (timeSlot.includes('9am')) stats.timeOfDayShifts.WeekendMorning++; 
                    else if (timeSlot.includes('1pm')) stats.timeOfDayShifts.WeekendAfternoon++; 
                } 
            }); 
        }); 
    });
    
    if (stats.totalShifts > 0) { 
        const activeStaffInStats = Object.values(stats.playerStats).filter(p => p.status !== onLeaveStatusText); 
        const activeStaffCount = activeStaffInStats.length; 
        stats.averageShiftsPerPlayer = activeStaffCount > 0 ? stats.totalShifts / activeStaffCount : 0; 
        stats.weekdayPercentage = (stats.totalWeekdayShifts / stats.totalShifts) * 100; 
        stats.weekendPercentage = (stats.totalWeekendShifts / stats.totalShifts) * 100; 
        const imbalanceThreshold = 1.5; 
        Object.entries(stats.playerStats).forEach(([name, pStats]) => { 
            pStats.utilizationPercentage = (pStats.totalShifts / stats.totalShifts) * 100; 
            if (pStats.totalShifts > stats.highestShifts.count) { 
                stats.highestShifts = { staff: name, count: pStats.totalShifts }; 
            } 
            if (pStats.totalShifts < stats.lowestShifts.count) { 
                stats.lowestShifts = { staff: name, count: pStats.totalShifts }; 
            } 
            if (pStats.status !== onLeaveStatusText && activeStaffCount > 1) { 
                const deviation = Math.abs(pStats.totalShifts - stats.averageShiftsPerPlayer); 
                if (deviation > imbalanceThreshold) { 
                    stats.imbalances.push(`${name}: ${pStats.totalShifts} shifts (Avg Active: ${stats.averageShiftsPerPlayer.toFixed(1)}, Deviation: ${deviation.toFixed(1)})`); 
                } 
            } 
        }); 
        if (stats.lowestShifts.count !== Infinity) { 
            const lowestCount = stats.lowestShifts.count; 
            const staffAtLowest = Object.entries(stats.playerStats).filter(([n, p]) => p.totalShifts === lowestCount).map(([n,p]) => n); 
            if (staffAtLowest.length > 3) { 
                stats.lowestShifts.staff = `${staffAtLowest.length} staff`; 
            } else if (staffAtLowest.length >= 1) { 
                stats.lowestShifts.staff = staffAtLowest.join(', '); 
            } else { 
                stats.lowestShifts = { staff: 'N/A', count: 0 }; 
            } 
        } else { 
            stats.lowestShifts = { staff: 'N/A', count: 0 }; 
        } 
    } else { 
        stats.lowestShifts = { staff: 'N/A', count: 0 }; 
        if (stats.errors.length === 0) { console.log("Stats Calc: No shifts found in schedule data."); } 
    }
    console.log(`--- Finished Statistics Calculation. Total Shifts Found: ${stats.totalShifts} ---`);
    return stats;
}


/**
 * Parses the raw data from the Schedule Sheet into a structured schedule object.
 * CORRECTED: Now properly parses staff2 for both Saturday and Sunday.
 */
function parseScheduleSheetToObject(scheduleData, sheetName) {
    const schedule = { week1: initializeWeek(), week2: initializeWeek() }; const errors = []; let currentWeekKey = null; let parsingSection = null; let weekdayDayHeaders = []; let weekendLastDayParsed = null; let weekendLastTimeSlotParsed = null; let weekendIsStaff1Row = true; const WEEKDAY_SECTION_HEADER = 'Weekday Schedule'; const WEEKEND_SECTION_HEADER = 'Weekend Schedule'; const WEEKDAY_HEADER_MARKER = 'Time Slot'; const WEEKEND_HEADER_MARKER = 'Time Slot'; const WEEKEND_STAFF_HEADER = 'Staff'; const WEEKEND_DAY_HEADER = 'Day';
    
    for (let i = 0; i < scheduleData.length; i++) { 
        const row = scheduleData[i]; const rowNum = i + 1; 
        if (!row || row.every(cell => String(cell || '').trim() === '')) continue; 
        const colA = String(row[0] || '').trim(); const colB = String(row[1] || '').trim(); const colC = String(row[2] || '').trim(); 
        
        try { 
            if (colA.startsWith('Week 1')) { currentWeekKey = 'week1'; parsingSection = null; weekdayDayHeaders = []; weekendLastDayParsed = null; weekendLastTimeSlotParsed = null; continue; } 
            if (colA.startsWith('Week 2')) { currentWeekKey = 'week2'; parsingSection = null; weekdayDayHeaders = []; weekendLastDayParsed = null; weekendLastTimeSlotParsed = null; continue; } 
            if (!currentWeekKey || !schedule[currentWeekKey]) continue; 
            if (colA === WEEKDAY_SECTION_HEADER) { parsingSection = 'weekday'; weekdayDayHeaders = []; continue; } 
            if (colA === WEEKEND_SECTION_HEADER) { parsingSection = 'weekend'; weekendLastDayParsed = null; weekendLastTimeSlotParsed = null; weekendIsStaff1Row = true; continue; }
    
            if (parsingSection === 'weekday') { 
                if (colA === WEEKDAY_HEADER_MARKER && WEEKDAYS.includes(colB)) { 
                    weekdayDayHeaders = row.slice(1).map(h => String(h).trim()); continue; 
                } 
                const timeSlotMatch = [...WEEKDAY_TIME_SLOTS.REGULAR, ...WEEKDAY_TIME_SLOTS.FRIDAY_EVENING].find(slot => slot === colA); // Check both sets
                if (timeSlotMatch && weekdayDayHeaders.length > 0) { 
                    weekdayDayHeaders.forEach((day, dayIndex) => { 
                        if (WEEKDAYS.includes(day)) { 
                            const staffName = String(row[dayIndex + 1] || '').trim(); 
                            if (staffName && staffName !== '-') { 
                                if (!schedule[currentWeekKey][day]) schedule[currentWeekKey][day] = {}; 
                                if (!schedule[currentWeekKey][day][timeSlotMatch]) schedule[currentWeekKey][day][timeSlotMatch] = { staff1: '', staff2: ''}; 
                                schedule[currentWeekKey][day][timeSlotMatch].staff1 = staffName; 
                            } 
                        } 
                    }); 
                } else if (colA === '' && i > 0 && weekdayDayHeaders.length > 0) { // This handles staff2 for weekdays
                    const prevRow = scheduleData[i-1]; const prevRowTimeSlot = String(prevRow[0] || '').trim(); 
                    const prevSlotMatch = [...WEEKDAY_TIME_SLOTS.REGULAR, ...WEEKDAY_TIME_SLOTS.FRIDAY_EVENING].find(slot => slot === prevRowTimeSlot); 
                    if (prevSlotMatch) { 
                        weekdayDayHeaders.forEach((day, dayIndex) => { 
                            if (WEEKDAYS.includes(day)) { 
                                const staffName = String(row[dayIndex + 1] || '').trim(); 
                                if (staffName && staffName !== '-') { 
                                    if (!schedule[currentWeekKey][day]) schedule[currentWeekKey][day] = {}; 
                                    if (!schedule[currentWeekKey][day][prevSlotMatch]) schedule[currentWeekKey][day][prevSlotMatch] = { staff1: '', staff2: ''}; 
                                    schedule[currentWeekKey][day][prevSlotMatch].staff2 = staffName; 
                                } 
                            } 
                        }); 
                    } 
                } 
            } else if (parsingSection === 'weekend') { 
                if (colA === WEEKEND_HEADER_MARKER && colB === WEEKEND_STAFF_HEADER && colC === WEEKEND_DAY_HEADER) continue; 
                const staffName = String(row[1] || '').trim(); // Staff name is in Col B (index 1)
                let day = String(row[2] || '').trim();       // Day is in Col C (index 2)
                let timeSlot = String(row[0] || '').trim();  // Time slot is in Col A (index 0)
                
                // Persist day and timeslot if they are merged (empty in subsequent rows)
                if (day) weekendLastDayParsed = day; else day = weekendLastDayParsed;
                if (timeSlot) weekendLastTimeSlotParsed = timeSlot; else timeSlot = weekendLastTimeSlotParsed;

                if (!day || !timeSlot || !WEEKEND_DAYS.includes(day)) {
                    // console.warn(`Row ${rowNum}: Skipping weekend row due to missing day/timeslot: Day='${day}', TimeSlot='${timeSlot}'`);
                    continue;
                }
                
                // Determine if this row is for staff1 or staff2
                // A row with an empty Column A AND an empty Column C (after the first staff1 row for a merged block) is a staff2 row
                const isStaff2Row = String(row[0] || '').trim() === '' && String(row[2] || '').trim() === '';
                const isStaff1Row = String(row[0] || '').trim() !== ''; // A non-empty Col A is staff1 for a new timeslot
                
                if (staffName && staffName !== '-') {
                    if (!schedule[currentWeekKey][day]) schedule[currentWeekKey][day] = {};
                    if (!schedule[currentWeekKey][day][timeSlot]) schedule[currentWeekKey][day][timeSlot] = { staff1: '', staff2: '' };
                    
                    if (isStaff1Row || !schedule[currentWeekKey][day][timeSlot].staff1) { // If it's a new timeslot row or staff1 is not yet filled
                        schedule[currentWeekKey][day][timeSlot].staff1 = staffName;
                    } else { // Otherwise, it must be staff2 for that timeslot
                        schedule[currentWeekKey][day][timeSlot].staff2 = staffName;
                    }
                }
            } 
        } catch (e) { 
            const errorMsg = `Error parsing row ${rowNum}: ${e.message}.`; console.error(errorMsg, e.stack); errors.push(errorMsg); break; 
        } 
    }
    if (errors.length > 0) {
        console.warn(`parseScheduleSheetToObject completed with ${errors.length} errors. Review logs.`);
    }
    return { schedule, errors };
}

/**
 * Displays shift statistics in a HTML modal dialog with Download/Email options.
 * This function remains largely the same but will now display corrected data.
 */
function displayShiftStatistics(preCalculatedStats = null) {
    const functionStartTime = new Date();
    logAction('Display Shift Statistics Attempt', 'Opening statistics modal.');
    const ui = SpreadsheetApp.getUi();
    let settings, stats, staffInfo;

    try {
        settings = getSettings();
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const staffInfoSheetName = settings.STAFF_INFO_SHEET;
        const staffInfoSheet = ss.getSheetByName(staffInfoSheetName);
        if (!staffInfoSheet) throw new Error(`Staff Info sheet "${staffInfoSheetName}" not found.`);
        staffInfo = getStaffInfo(staffInfoSheet, settings);

        if (preCalculatedStats) {
            stats = preCalculatedStats;
        } else {
            const scheduleSheetName = settings.SCHEDULE_SHEET; 
            const scheduleSheet = ss.getSheetByName(scheduleSheetName); 
            if (!scheduleSheet) throw new Error(`Schedule sheet "${scheduleSheetName}" not found.`); 
            const scheduleData = scheduleSheet.getDataRange().getValues(); 
            const parsingResult = parseScheduleSheetToObject(scheduleData, scheduleSheetName); 
            const currentScheduleFromSheet = parsingResult.schedule; 
            const parsingErrors = parsingResult.errors; 
            if (parsingErrors.length > 0) console.warn("Parsing Errors:", parsingErrors);
            stats = getCurrentScheduleStats(currentScheduleFromSheet, staffInfo, settings); 
            if (!stats.errors) stats.errors = []; 
            if (parsingErrors.length > 0) stats.errors.unshift(...parsingErrors.map(err => `Parsing Error: ${err}`));
        }
        
        if (!stats || typeof stats !== 'object') throw new Error("Stats object invalid.");
        stats.playerStats = stats.playerStats || {}; 
        stats.errors = stats.errors || []; 
        stats.imbalances = stats.imbalances || []; 
        stats.timeOfDayShifts = stats.timeOfDayShifts || {}; 
        stats.highestShifts = stats.highestShifts || { staff: 'N/A', count: 0 }; 
        stats.lowestShifts = stats.lowestShifts || { staff: 'N/A', count: 0 };

        // Pre-serialized payloads to embed directly in the client without extra parsing layers.
        const statsPayload = JSON.stringify(stats);
        const staffInfoPayload = JSON.stringify(staffInfo);
        const settingsPayload = JSON.stringify(settings);
        const onLeaveText = settings.ON_LEAVE_STATUS_TEXT || 'On Leave';

        const statsHtmlString = `
        <!DOCTYPE html>
        <html>
        <head>
          <base target="_top">
          <script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js" integrity="sha512-Qlv6VSKh1gDKGoJbnyA5Nw7esAgUZ5FFUYWfVejTlgNPaMLMIUWFGwBzBCKbUR1BlochQSX6DgaMQAGClmMCmLQA==" crossorigin="anonymous" referrerpolicy="no-referrer"></script>
          <style>
            * { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif; box-sizing: border-box; }
            body { padding: 20px; background-color: #f8fafc; line-height: 1.5; color: #334155; }
            h2 { color: #1e293b; font-weight: 600; margin-bottom: 10px; text-align: center; }
            h3 { color: #334155; font-weight: 600; margin-top: 25px; margin-bottom: 15px; border-bottom: 1px solid #e2e8f0; padding-bottom: 8px; }
            h4 { color: #475569; font-weight: 600; margin-top: 20px; margin-bottom: 10px; }
            .subtitle { text-align: center; color: #64748b; margin-top: 0; margin-bottom: 20px; }
            .stats-container { background: white; border-radius: 8px; box-shadow: 0 2px 6px rgba(0, 0, 0, 0.1); padding: 24px; max-width: 1200px; margin: 0 auto; }
            .kpi-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 16px; margin-bottom: 24px; }
            .kpi-card { background-color: #f8fafc; border-radius: 8px; padding: 16px; box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05); }
            .kpi-title { font-size: 14px; color: #64748b; margin-bottom: 8px; font-weight: 500; }
            .kpi-value { font-size: 24px; font-weight: 600; color: #334155; margin-bottom: 4px; }
            .kpi-unit { font-size: 12px; color: #94a3b8; }
            .stats-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(240px, 1fr)); gap: 16px; margin-bottom: 24px; }
            .player-card { background-color: #f8fafc; border-radius: 8px; padding: 16px; box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05); }
            .player-name { font-weight: 600; font-size: 16px; color: #334155; margin-bottom: 10px; padding-bottom: 8px; border-bottom: 1px solid #e2e8f0; }
            .player-stats { display: grid; grid-template-columns: repeat(2, 1fr); gap: 8px; }
            .player-stat-item { margin-bottom: 8px; }
            .player-stat-label { font-size: 12px; color: #64748b; display: block; }
            .player-stat-value { font-size: 14px; color: #334155; font-weight: 500; }
            .alert { background-color: #fff4e6; border-left: 4px solid #f97316; padding: 12px 16px; margin-bottom: 24px; color: #7c2d12; border-radius: 4px; }
            .imbalance-list { background-color: #f0f9ff; border-left: 4px solid #0ea5e9; padding: 12px 16px; margin-bottom: 24px; color: #0c4a6e; border-radius: 4px; }
            .imbalance-list ul { margin-top: 8px; margin-bottom: 0; padding-left: 20px; }
            .status-active { color: #16a34a; }
            .status-onleave { color: #ea580c; }
            .status-unknown { color: #64748b; }
            .button-row { display: flex; justify-content: center; gap: 16px; margin-top: 32px; }
            .action-btn { background-color: #2563eb; color: white; padding: 10px 20px; border: none; border-radius: 6px; cursor: pointer; font-weight: 500; transition: background-color 0.2s ease; }
            .action-btn:hover { background-color: #1d4ed8; }
            .action-btn:disabled { background-color: #cbd5e1; cursor: not-allowed; }
            .action-btn.email { background-color: #0891b2; }
            .action-btn.email:hover { background-color: #0e7490; }
          </style>
        </head>
        <body>
          <div class="stats-container" id="stats-content">
            <h2>Staff Shift Statistics</h2>
            <p class="subtitle">(Based on the current schedule data)</p>
            
            <? if (stats && stats.errors && stats.errors.length > 0) { ?>
              <div class="alert">
                <strong>Issues encountered:</strong>
                <ul>
                  <? stats.errors.forEach(function(err) { ?>
                    <li><?= err ?></li>
                  <? }); ?>
                </ul>
              </div>
            <? } ?>
            
            <? if (typeof stats !== 'undefined' && stats) { ?>
              <h3>Overall Summary</h3>
              <div class="kpi-grid">
                <div class="kpi-card">
                  <div class="kpi-title">Total Shifts</div>
                  <div class="kpi-value"><?= stats.totalShifts ?></div>
                </div>
                <div class="kpi-card">
                  <div class="kpi-title">Weekday Shifts</div>
                  <div class="kpi-value"><?= stats.totalWeekdayShifts ?></div>
                  <div class="kpi-unit"><?= (stats.weekdayPercentage || 0).toFixed(1) ?>% of total</div>
                </div>
                <div class="kpi-card">
                  <div class="kpi-title">Weekend Shifts</div>
                  <div class="kpi-value"><?= stats.totalWeekendShifts ?></div>
                  <div class="kpi-unit"><?= (stats.weekendPercentage || 0).toFixed(1) ?>% of total</div>
                </div>
                <div class="kpi-card">
                  <div class="kpi-title">Avg. Shifts Per Active Staff</div>
                  <div class="kpi-value"><?= (stats.averageShiftsPerPlayer || 0).toFixed(1) ?></div>
                </div>
                <div class="kpi-card">
                  <div class="kpi-title">Highest Shifts</div>
                  <div class="kpi-value"><?= stats.highestShifts.count ?></div>
                  <div class="kpi-unit"><?= stats.highestShifts.staff ?></div>
                </div>
                <div class="kpi-card">
                  <div class="kpi-title">Lowest Shifts</div>
                  <div class="kpi-value"><?= stats.lowestShifts.count ?></div>
                  <div class="kpi-unit"><?= stats.lowestShifts.staff ?></div>
                </div>
              </div>
              
              <? if (stats.imbalances && stats.imbalances.length > 0) { ?>
                <div class="imbalance-list">
                  <h4>Potential Imbalances (Active Staff):</h4>
                  <ul>
                    <? stats.imbalances.forEach(function(imbalanceMsg) { ?>
                      <li><?= imbalanceMsg ?></li>
                    <? }); ?>
                  </ul>
                </div>
              <? } ?>
              
              <h3>Staff Breakdown</h3>
              <div class="stats-grid">
                <? var playerStats = stats.playerStats || {}; ?>
                <? if (Object.keys(playerStats).length > 0) { ?>
                  <? Object.keys(playerStats).sort().forEach(function(name) { ?>
                    <? var pStats = playerStats[name] || {}; ?>
                    <? var statusClass = pStats.status === onLeaveText ? 'status-onleave' : (pStats.status === 'Active' ? 'status-active' : 'status-unknown'); ?>
                    <div class="player-card">
                      <div class="player-name">
                        <?= name ?>
                        <span style="font-size:0.8em; font-weight:normal;" class="<?= statusClass ?>">(<?= pStats.status || 'N/A' ?>)</span>
                      </div>
                      <div class="player-stats">
                        <div class="player-stat-item">
                          <span class="player-stat-label">Total Shifts</span>
                          <span class="player-stat-value"><?= pStats.totalShifts ?></span>
                        </div>
                        <div class="player-stat-item">
                          <span class="player-stat-label">Utilization</span>
                          <span class="player-stat-value"><?= (pStats.utilizationPercentage || 0).toFixed(1) ?>%</span>
                        </div>
                        <div class="player-stat-item">
                          <span class="player-stat-label">Weekday</span>
                          <span class="player-stat-value"><?= pStats.weekdayShifts ?></span>
                        </div>
                        <div class="player-stat-item">
                          <span class="player-stat-label">Weekend</span>
                          <span class="player-stat-value"><?= pStats.weekendShifts ?></span>
                        </div>
                        <div class="player-stat-item">
                          <span class="player-stat-label">Level</span>
                          <span class="player-stat-value"><?= pStats.level || 'N/A' ?></span>
                        </div>
                        <div class="player-stat-item">
                          <span class="player-stat-label">Region</span>
                          <span class="player-stat-value"><?= pStats.region || 'N/A' ?></span>
                        </div>
                      </div>
                    </div>
                  <? }); ?>
                <? } else { ?>
                  <p>No individual staff statistics available.</p>
                <? } ?>
              </div>
              
              <h3>Shift Times Breakdown</h3>
              <div class="kpi-grid">
                <? var timeShifts = stats.timeOfDayShifts || {}; ?>
                <div class="kpi-card">
                  <div class="kpi-title">Morning Shifts</div>
                  <div class="kpi-value"><?= timeShifts.Morning || 0 ?></div>
                </div>
                <div class="kpi-card">
                  <div class="kpi-title">Afternoon Shifts</div>
                  <div class="kpi-value"><?= timeShifts.Afternoon || 0 ?></div>
                </div>
                <div class="kpi-card">
                  <div class="kpi-title">Evening Shifts</div>
                  <div class="kpi-value"><?= timeShifts.Evening || 0 ?></div>
                </div>
                <div class="kpi-card">
                  <div class="kpi-title">Weekend Morning</div>
                  <div class="kpi-value"><?= timeShifts.WeekendMorning || 0 ?></div>
                </div>
                <div class="kpi-card">
                  <div class="kpi-title">Weekend Afternoon</div>
                  <div class="kpi-value"><?= timeShifts.WeekendAfternoon || 0 ?></div>
                </div>
              </div>
              
              <div class="button-row">
                <button class="action-btn" id="downloadCsvBtn">Download Stats (CSV)</button>
                <button class="action-btn email" id="emailStatsBtn">Email Statistics</button>
              </div>
            <? } else { ?>
              <div class="alert"><strong>Error:</strong> Could not display statistics data.</div>
            <? } ?>
          </div>
          
          <script>
            // Client-side script remains unchanged.
            // Pre-embedded JSON objects (safe, not double-encoded)
            let statsData = <?!= statsPayload ?>;
            let staffInfoData = <?!= staffInfoPayload ?>;
            let settingsData = <?!= settingsPayload ?>;
            let onLeaveText = settingsData['ON_LEAVE_STATUS_TEXT'] || 'On Leave';
            function escapeCsvCell(cellData) { return cellData === null || cellData === undefined ? '' : String(cellData).includes(',') || String(cellData).includes('\\n') || String(cellData).includes('"') ? '"' + String(cellData).replace(/"/g, '""') + '"' : String(cellData); }
            function downloadStatsCSV() {
              if (!statsData) { alert("Error preparing CSV."); return; }
              try {
                let csvContent = [];
                csvContent.push("Staff Member,Total Shifts,Weekday Shifts,Weekend Shifts,Utilization %,Level,Region,Status");
                Object.keys(statsData.playerStats || {}).sort().forEach(name => {
                  const staff = statsData.playerStats[name];
                  csvContent.push([
                    escapeCsvCell(name),
                    escapeCsvCell(staff.totalShifts),
                    escapeCsvCell(staff.weekdayShifts),
                    escapeCsvCell(staff.weekendShifts),
                    escapeCsvCell((staff.utilizationPercentage || 0).toFixed(1)),
                    escapeCsvCell(staff.level),
                    escapeCsvCell(staff.region),
                    escapeCsvCell(staff.status)
                  ].join(','));
                });
                csvContent.push('');
                csvContent.push('Overall Statistics,Value');
                csvContent.push('Total Shifts,' + escapeCsvCell(statsData.totalShifts));
                csvContent.push('Average Shifts Per Staff,' + escapeCsvCell((statsData.averageShiftsPerPlayer || 0).toFixed(1)));
                if (statsData.imbalances && statsData.imbalances.length > 0) {
                  csvContent.push('');
                  csvContent.push('Potential Imbalances');
                  statsData.imbalances.forEach(imbalance => { csvContent.push(escapeCsvCell(imbalance)); });
                }
                const csvString = csvContent.join("\\n");
                const blob = new Blob(["\uFEFF" + csvString], { type: 'text/csv;charset=utf-8;' });
                if (typeof saveAs === 'function') {
                  saveAs(blob, "shift_statistics.csv");
                } else {
                  // Fallback download without FileSaver
                  const url = URL.createObjectURL(blob);
                  const a = document.createElement('a');
                  a.href = url;
                  a.download = "shift_statistics.csv";
                  document.body.appendChild(a);
                  a.click();
                  a.remove();
                  URL.revokeObjectURL(url);
                }
              } catch (e) {
                console.error("CSV generation error:", e);
                alert("Error creating CSV.");
              }
            }
            function triggerEmailStats() {
              if (!statsData) { alert("Error preparing email."); return; }
              const btn = document.getElementById('emailStatsBtn');
              btn.disabled = true; btn.textContent = 'Sending...';
              google.script.run.withSuccessHandler(() => {
                  alert('Statistics email sent!');
                  btn.disabled = false; btn.textContent = 'Email Statistics';
              }).withFailureHandler((error) => {
                  alert('Failed to send email: ' + error.message);
                  btn.disabled = false; btn.textContent = 'Email Statistics';
              }).sendStatisticsEmail(JSON.stringify(statsData));
            }
            
            var downloadBtnElem = document.getElementById('downloadCsvBtn');
            if (downloadBtnElem) downloadBtnElem.addEventListener('click', downloadStatsCSV);
            
            var emailBtnElem = document.getElementById('emailStatsBtn');
            if (emailBtnElem) emailBtnElem.addEventListener('click', triggerEmailStats);

          </script>
        </body>
        </html>`;

        const template = HtmlService.createTemplate(statsHtmlString);
        template.stats = stats;
        template.staffInfo = staffInfo;
        template.settings = settings;
        template.onLeaveText = onLeaveText;
        template.statsPayload = statsPayload;
        template.staffInfoPayload = staffInfoPayload;
        template.settingsPayload = settingsPayload;

        const html = template.evaluate().setWidth(1000).setHeight(800);
        ui.showModalDialog(html, 'Staff Shift Statistics');
    } catch (e) {
        const errorMsg = `Error displaying shift statistics: ${e.message}`;
        console.error(errorMsg, e.stack);
        logAction('Display Stats CRITICAL Error', errorMsg, 'ERROR', settings);
        ui.alert('Error', `Failed to display statistics:\n${e.message}`, ui.ButtonSet.OK);
    }
}


// =================================================================
// END OF SECTION 3
// =================================================================

// =================================================================
// SECTION 4: ORCHESTRATION, MANUAL EDITS, RESETS, LOGGING, MENU
// =================================================================

// --- Rotation Cycle Logic ---

/**
 * Checks if the rotation cycle is complete.
 * A cycle is considered complete when *all* active staff members have a rotation count > 0.
 * @param {Object} staffInfo - The staff information object.
 * @param {Object} settings - The loaded settings object.
 * @return {boolean} True if the rotation cycle is complete for all active staff, false otherwise.
 */
function checkRotationCycleCompletion(staffInfo, settings) {
  if (!staffInfo || Object.keys(staffInfo).length === 0) {
      console.log("checkRotationCycleCompletion: No staff info available.");
      return false;
  }

  const onLeaveStatusText = settings.ON_LEAVE_STATUS_TEXT;
  const activeStaff = Object.values(staffInfo).filter(staff => staff.status !== onLeaveStatusText);

  if (activeStaff.length === 0) {
      console.log("checkRotationCycleCompletion: No active staff found. Cycle cannot be complete.");
      return false; 
  }

  const allActiveHaveRotated = activeStaff.every(staff => (staff.rotationCount || 0) > 0);

  console.log(`Rotation cycle completion check (for ${activeStaff.length} active staff): ${allActiveHaveRotated ? 'Complete' : 'Not Complete'}`);
  return allActiveHaveRotated;
}

/**
 * Creates a backup row in the 'Count Backups' sheet for each staff member.
 * @param {Sheet} staffInfoSheet - The Staff Info sheet object.
 * @param {Object} settings - The loaded settings object.
 * @return {number} The number of staff records backed up.
 */
function createRotationCountBackup(staffInfoSheet, settings) {
    console.log("--- Creating Rotation Count Backup ---");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const backupSheetName = settings.BACKUPS_SHEET;
    let backupSheet = ss.getSheetByName(backupSheetName);
    const backupTimestamp = new Date();

    if (!backupSheet) {
        backupSheet = ss.insertSheet(backupSheetName);
        backupSheet.appendRow([
            'Backup Timestamp',
            STAFF_INFO_HEADERS.PLAYER_NAME,
            STAFF_INFO_HEADERS.ROTATION_COUNT,
            STAFF_INFO_HEADERS.WEEKDAY_SHIFTS,
            STAFF_INFO_HEADERS.WEEKEND_SHIFTS,
            STAFF_INFO_HEADERS.TOTAL_SHIFTS,
            STAFF_INFO_HEADERS.LAST_EXTRA_SHIFT,
            STAFF_INFO_HEADERS.LAST_WEEKEND_SHIFT
        ]);
        console.log(`Created sheet: "${backupSheetName}". Added headers.`);
    }

    const data = staffInfoSheet.getDataRange().getValues();
    if (data.length < 2) {
        console.log("No data rows found in Staff Info sheet to backup.");
        return 0;
    }

    const headers = data[0].map(h => String(h).trim());
    const colIndices = {
        name: headers.indexOf(STAFF_INFO_HEADERS.PLAYER_NAME),
        rotation: headers.indexOf(STAFF_INFO_HEADERS.ROTATION_COUNT),
        weekday: headers.indexOf(STAFF_INFO_HEADERS.WEEKDAY_SHIFTS),
        weekend: headers.indexOf(STAFF_INFO_HEADERS.WEEKEND_SHIFTS),
        total: headers.indexOf(STAFF_INFO_HEADERS.TOTAL_SHIFTS),
        lastExtra: headers.indexOf(STAFF_INFO_HEADERS.LAST_EXTRA_SHIFT),
        lastWeekend: headers.indexOf(STAFF_INFO_HEADERS.LAST_WEEKEND_SHIFT)
    };

    const backupData = [];
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const staffName = String(row[colIndices.name] || '').trim();
        if (!staffName) continue;

        backupData.push([
            backupTimestamp,
            staffName,
            row[colIndices.rotation],
            row[colIndices.weekday],
            row[colIndices.weekend],
            row[colIndices.total],
            row[colIndices.lastExtra],
            row[colIndices.lastWeekend]
        ]);
    }

    if (backupData.length > 0) {
        backupSheet.getRange(backupSheet.getLastRow() + 1, 1, backupData.length, backupData[0].length).setValues(backupData);
        console.log(`Backed up ${backupData.length} staff records to "${backupSheetName}".`);
    } else {
        console.log("No staff data found in Staff Info sheet to backup.");
    }
    SpreadsheetApp.flush();
    return backupData.length;
}

/**
 * Resets rotation counts (and related historical shift counts/dates) to 0/empty for all staff members.
 * @param {Sheet} staffInfoSheet - The Staff Info sheet object.
 * @return {number} The number of staff records whose counts were reset.
 */
function resetRotationCounts(staffInfoSheet) {
    console.log("--- Resetting Rotation Counts and Historical Shifts ---");
    const range = staffInfoSheet.getDataRange();
    const data = range.getValues();
    if (data.length < 2) {
        console.log("No data rows in Staff Info sheet to reset.");
        return 0;
    }

    const headers = data[0].map(h => String(h).trim());
    const colIndicesToReset = {
        rotationCount: headers.indexOf(STAFF_INFO_HEADERS.ROTATION_COUNT),
        weekdayShifts: headers.indexOf(STAFF_INFO_HEADERS.WEEKDAY_SHIFTS),
        weekendShifts: headers.indexOf(STAFF_INFO_HEADERS.WEEKEND_SHIFTS),
        totalShifts: headers.indexOf(STAFF_INFO_HEADERS.TOTAL_SHIFTS),
        lastExtraShift: headers.indexOf(STAFF_INFO_HEADERS.LAST_EXTRA_SHIFT),
        lastWeekendShift: headers.indexOf(STAFF_INFO_HEADERS.LAST_WEEKEND_SHIFT)
    };
    
    let recordsModified = 0;
    for (let i = 1; i < data.length; i++) {
        const staffName = String(data[i][headers.indexOf(STAFF_INFO_HEADERS.PLAYER_NAME)] || '').trim();
        if (!staffName) continue;

        let rowModified = false;
        Object.values(colIndicesToReset).forEach(colIndex => {
            if (colIndex !== -1) {
                // Reset to 0 for counts, empty string for dates
                const resetValue = headers[colIndex].includes('Shift') && !headers[colIndex].includes('Count') ? '' : 0;
                 if (data[i][colIndex] !== resetValue) {
                     data[i][colIndex] = resetValue;
                     rowModified = true;
                 }
            }
        });
        if (rowModified) recordsModified++;
    }

    if (recordsModified > 0) {
        range.setValues(data);
        console.log(`Reset counts for ${recordsModified} staff records in "${staffInfoSheet.getName()}".`);
    } else {
        console.log("No counts or dates needed resetting in Staff Info sheet.");
    }
    SpreadsheetApp.flush();
    return recordsModified;
}

// --- Logging ---

/**
 * Logs script actions to the 'Logs' sheet.
 * @param {string} action - Description of the action performed.
 * @param {string} [details=''] - Optional details.
 * @param {string} [status='INFO'] - Status level ('INFO', 'WARN', 'ERROR').
 * @param {Object} [settings=null] - Loaded settings object.
 */
function logAction(action, details = '', status = 'INFO', settings = null) {
  let logsSheetName = 'Logs'; 
  let logsSheet = null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    if (settings && settings.LOGS_SHEET) {
        logsSheetName = settings.LOGS_SHEET;
    }
    logsSheet = ss.getSheetByName(logsSheetName);

  } catch (e) {
    console.error(`Log Action Failed: Could not get Logs sheet "${logsSheetName}". Error: ${e}`);
    console.log(`ACTION (Fallback): ${status.toUpperCase()} - ${action}${details ? ' | Details: ' + details : ''}`);
    return;
  }

  if (!logsSheet) {
    try {
         logsSheet = ss.insertSheet(logsSheetName);
         logsSheet.appendRow(['Timestamp', 'User', 'Status', 'Action', 'Details']);
    } catch (createError) {
        console.error(`CRITICAL Log Action Failed: Could not create Logs sheet "${logsSheetName}". Error: ${createError}`);
        console.log(`ACTION (CRITICAL Fallback): ${status.toUpperCase()} - ${action}${details ? ' | Details: ' + details : ''}`);
        return;
    }
  }

  try {
    const timestamp = new Date();
    let user = Session.getActiveUser()?.getEmail() || Session.getEffectiveUser()?.getEmail() || 'unknown';
    logsSheet.appendRow([timestamp, user, status.toUpperCase(), action, details]);
    SpreadsheetApp.flush();
  } catch (writeError) {
    console.error(`CRITICAL Log Action Failed: Cannot write to sheet "${logsSheetName}". Error: ${writeError}.`);
  }
}


// --- Manual Shift Editor ---

/**
 * Shows the manual shift editor UI.
 */
function showManualShiftEditor() {
    logAction('Show Manual Editor Attempt', 'Opening manual shift editor modal.');
    const ui = SpreadsheetApp.getUi();

    try {
        const settings = getSettings();
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const staffInfoSheet = ss.getSheetByName(settings.STAFF_INFO_SHEET);
        if (!staffInfoSheet) throw new Error(`Staff Info sheet "${settings.STAFF_INFO_SHEET}" not found.`);

        const staffInfoData = getStaffInfo(staffInfoSheet, settings); // staffInfoData is the object

        const template = HtmlService.createTemplateFromFile('ManualShiftEditor');
        
        // Pass the staffInfo object directly
        template.staffInfo = staffInfoData; 

        // Pass all necessary constants and settings explicitly
        template.ALL_DAYS_DATA = ALL_DAYS;
        template.WEEKDAYS_DATA = WEEKDAYS;
        template.WEEKEND_DAYS_DATA = WEEKEND_DAYS;
        template.WEEKDAY_TIME_SLOTS_DATA = WEEKDAY_TIME_SLOTS;
        template.WEEKEND_TIME_SLOTS_DATA = WEEKEND_TIME_SLOTS;
        template.settingsFromServer = settings; // Pass the whole settings object

        const html = template.evaluate().setWidth(650).setHeight(600);
        ui.showModalDialog(html, 'Manual Shift Editor');

    } catch (e) {
        console.error("Error showing manual editor: " + e.message, e.stack);
        logAction('Show Manual Editor Error', e.message, 'ERROR');
        ui.alert("Error opening Manual Shift Editor:", `Failed:\n${e.message}`, ui.ButtonSet.OK);
    }
}

/**
 * Handles the request from the Manual Shift Editor UI to update a specific shift assignment.
 * CORRECTED: Now supports updating Sunday staff2.
 * @param {Object} formData - Shift details from the client-side form.
 * @return {string} A success message.
 */
function updateShiftAssignment(formData) {
    console.log("--- Manual Shift Update Attempt ---", formData);
    logAction('Manual Shift Update Attempt', `Data: ${JSON.stringify(formData)}`);

    const settings = getSettings(); // Load settings to check REQ values if needed
    const scheduleSheet = getSheetOrThrow(settings.SCHEDULE_SHEET, 'updateShiftAssignment');
    const staffInfoSheet = getSheetOrThrow(settings.STAFF_INFO_SHEET, 'updateShiftAssignment');
    
    const scheduleData = scheduleSheet.getDataRange().getValues();
    let targetRowIndex = -1, targetColIndex = -1, weekStartRowIndex = -1;

    const weekHeaderText = formData.week === 'week1' ? 'Week 1' : 'Week 2';
    weekStartRowIndex = scheduleData.findIndex(row => String(row[0]).trim().startsWith(weekHeaderText)); // Use startsWith for flexibility
    if (weekStartRowIndex === -1) throw new Error(`Could not find "${weekHeaderText}" header in sheet.`);

    const isWeekend = WEEKEND_DAYS.includes(formData.day);

    try {
        if (!isWeekend) { // Weekday including Friday
            let weekdaySectionStartIndex = scheduleData.findIndex((row, i) => i > weekStartRowIndex && String(row[0]).trim() === 'Weekday Schedule');
            if (weekdaySectionStartIndex === -1) throw new Error("Could not find 'Weekday Schedule' header after " + weekHeaderText);
            
            let dayHeaderRowIndex = weekdaySectionStartIndex + 1;
            if (dayHeaderRowIndex >= scheduleData.length || String(scheduleData[dayHeaderRowIndex][0]).trim() !== 'Time Slot') {
                 throw new Error("Weekday 'Time Slot' header row not found where expected.");
            }
            
            const dayHeaders = scheduleData[dayHeaderRowIndex].map(h => String(h).trim());
            targetColIndex = dayHeaders.indexOf(formData.day); // This is 0-based from the dayHeaders array
            if (targetColIndex === -1) throw new Error(`Day "${formData.day}" column not found in weekday headers.`);
            // targetColIndex needs to be 1-based for sheet.getRange(), and dayHeaders[0] is 'Time Slot'
            // So if formData.day is 'Monday' (index 1 in dayHeaders), actual sheet col is 2 (B).
            // If formData.day is dayHeaders[targetColIndex], then sheet col is targetColIndex + 1
            
            let timeSlotStartRowIndex = -1;
            for (let i = dayHeaderRowIndex + 1; i < scheduleData.length; i++) {
                if(String(scheduleData[i][0]).trim() === formData.timeSlot) {
                    timeSlotStartRowIndex = i;
                    break;
                }
                if(String(scheduleData[i][0]).trim().startsWith('Week')) break; // Stop if next week is reached
            }
            if (timeSlotStartRowIndex === -1) throw new Error(`Time slot "${formData.timeSlot}" not found for ${formData.day}.`);

            targetRowIndex = timeSlotStartRowIndex + (formData.staffPosition === 'staff2' ? 1 : 0);
        } else { // Weekend (Saturday or Sunday)
            targetColIndex = 1; // Staff is always in Col B (index 1) for weekends in the 3-column layout
            let weekendSectionStartIndex = scheduleData.findIndex((row, i) => i > weekStartRowIndex && String(row[0]).trim() === 'Weekend Schedule');
            if (weekendSectionStartIndex === -1) throw new Error("Could not find 'Weekend Schedule' header after " + weekHeaderText);
            
            let dayBlockStartIndex = -1, timeSlotBlockStartIndex = -1;
            
            // Find the start of the specific day block (e.g., "Saturday")
            // Day is in Column C (index 2) for the first row of a weekend day block
            for (let i = weekendSectionStartIndex + 2; i < scheduleData.length; i++) { // Start search after header
                if (String(scheduleData[i][2]).trim() === formData.day) { 
                    dayBlockStartIndex = i;
                    break;
                }
                if(String(scheduleData[i][0]).trim().startsWith('Week')) break; // Stop if next week is reached
            }
            if (dayBlockStartIndex === -1) throw new Error(`Day block for "${formData.day}" not found in weekend section.`);

            // Find the specific time slot within that day block
            // Time slot is in Column A (index 0) for the first row of a time slot
            for (let i = dayBlockStartIndex; i < scheduleData.length; i++) {
                // If we encounter a new Day in Col C that is not our target day, or a new section, stop.
                const currentRowDay = String(scheduleData[i][2] || '').trim();
                if (currentRowDay && currentRowDay !== formData.day && WEEKEND_DAYS.includes(currentRowDay)) break;
                if (String(scheduleData[i][0]).trim().startsWith('Week')) break;


                if (String(scheduleData[i][0]).trim() === formData.timeSlot) { 
                    timeSlotBlockStartIndex = i;
                    break;
                }
            }
             if (timeSlotBlockStartIndex === -1) throw new Error(`Time slot "${formData.timeSlot}" not found for "${formData.day}" in weekend section.`);

            targetRowIndex = timeSlotBlockStartIndex + (formData.staffPosition === 'staff2' ? 1 : 0);
        }

    } catch (parseError) {
        logAction('Manual Shift Update Error', `Finding cell: ${parseError.message}`, 'ERROR', settings);
        throw new Error(`Error finding cell: ${parseError.message}. Check sheet format.`);
    }

    if (targetRowIndex === -1 || targetColIndex === -1) {
        logAction('Manual Shift Update Error', `Could not determine target cell. Row: ${targetRowIndex}, Col: ${targetColIndex}`, 'ERROR', settings);
        throw new Error("Could not determine target cell.");
    }
    
    // Check if the target cell is valid (within sheet bounds)
    if (targetRowIndex >= scheduleData.length || targetColIndex >= (scheduleData[targetRowIndex] ? scheduleData[targetRowIndex].length : 0) ) {
        logAction('Manual Shift Update Error', `Target cell [${targetRowIndex+1}, ${targetColIndex+1}] is out of bounds.`, 'ERROR', settings);
        throw new Error(`Calculated target cell [${targetRowIndex+1}, ${targetColIndex+1}] is out of bounds for the sheet.`);
    }


    const currentStaff = String(scheduleData[targetRowIndex][targetColIndex] || '').trim();
    const newStaff = (formData.staffMember === '--REMOVE--' || formData.staffMember === '-') ? '' : String(formData.staffMember).trim();

    if (currentStaff === newStaff || (currentStaff === '-' && newStaff === '')) {
        return "No change needed. Assignment is already correct.";
    }

    // Update Schedule Sheet
    scheduleSheet.getRange(targetRowIndex + 1, targetColIndex + 1).setValue(newStaff === '' ? '-' : newStaff);

    // Update Staff Info Counts
    const staffInfoRange = staffInfoSheet.getDataRange();
    const staffInfoData = staffInfoRange.getValues();
    const staffHeaders = staffInfoData[0].map(h => String(h).trim());
    const nameIndex = staffHeaders.indexOf(STAFF_INFO_HEADERS.PLAYER_NAME);
    const weekdayIndex = staffHeaders.indexOf(STAFF_INFO_HEADERS.WEEKDAY_SHIFTS);
    const weekendIndex = staffHeaders.indexOf(STAFF_INFO_HEADERS.WEEKEND_SHIFTS);
    const totalIndex = staffHeaders.indexOf(STAFF_INFO_HEADERS.TOTAL_SHIFTS);

    if (nameIndex === -1 || weekdayIndex === -1 || weekendIndex === -1 || totalIndex === -1) {
        throw new Error("One or more required columns missing from Staff Info sheet for count update.");
    }

    const countCol = isWeekend ? weekendIndex : weekdayIndex;

    // Decrement old staff count
    if (currentStaff && currentStaff !== '-') {
        const oldStaffRowArray = staffInfoData.find(row => String(row[nameIndex]).trim() === currentStaff);
        if (oldStaffRowArray) {
            oldStaffRowArray[countCol] = Math.max(0, (Number(oldStaffRowArray[countCol]) || 0) - 1);
            oldStaffRowArray[totalIndex] = (Number(oldStaffRowArray[weekdayIndex]) || 0) + (Number(oldStaffRowArray[weekendIndex]) || 0);
        }
    }
    // Increment new staff count
    if (newStaff) {
        const newStaffRowArray = staffInfoData.find(row => String(row[nameIndex]).trim() === newStaff);
        if (newStaffRowArray) {
            newStaffRowArray[countCol] = (Number(newStaffRowArray[countCol]) || 0) + 1;
            newStaffRowArray[totalIndex] = (Number(newStaffRowArray[weekdayIndex]) || 0) + (Number(newStaffRowArray[weekendIndex]) || 0);
        } else {
             throw new Error(`Selected new staff "${newStaff}" not found in Staff Info sheet. Counts not updated.`);
        }
    }

    staffInfoRange.setValues(staffInfoData);
    SpreadsheetApp.flush();

    logAction('Manual Shift Edit Success', `Changed ${formData.day} ${formData.timeSlot} (Pos: ${formData.staffPosition}) from "${currentStaff}" to "${newStaff}".`, 'INFO', settings);
    return `Shift updated successfully.`;
}


// --- Reset Functionality ---

/**
 * Shows a confirmation dialog before resetting all historical counts.
 */
function showResetConfirmation() {
  const settings = getSettings();
  const ui = SpreadsheetApp.getUi();

  const confirmMessage = `WARNING: This will BACKUP current counts from "${settings.STAFF_INFO_SHEET}" to "${settings.BACKUPS_SHEET}", then RESET ALL counts and dates in "${settings.STAFF_INFO_SHEET}" to zero/empty.\n\nThis cannot be undone.\n\nProceed?`;
  const response = ui.alert('Confirm Reset All Shift Counts', confirmMessage, ui.ButtonSet.YES_NO);

  if (response === ui.Button.YES) {
    try {
      const result = resetAllCounts(settings);
      logAction('Reset Counts Success', `Reset ${result.recordsReset} records. Backup created: ${result.backupRecords}.`, 'INFO', settings);
      ui.alert('Reset Successful', `All shift counts in "${settings.STAFF_INFO_SHEET}" have been reset.\nA backup was created in "${settings.BACKUPS_SHEET}".`, ui.ButtonSet.OK);
    } catch (error) {
      logAction('Reset Counts Error', error.message, 'ERROR', settings);
      ui.alert('Reset Failed', `Failed to reset counts:\n${error.message}`, ui.ButtonSet.OK);
    }
  } else {
    logAction('Reset Counts Cancelled', 'User cancelled the reset operation.', 'INFO', settings);
    ui.alert('Reset cancelled.');
  }
}

/**
 * Performs the full reset process: backup and then reset.
 * @param {Object} settings - The loaded settings object.
 * @return {Object} Result object: { recordsReset: number, backupRecords: number }.
 */
function resetAllCounts(settings) {
  console.log("--- Starting Reset All Counts Process ---");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staffSheet = ss.getSheetByName(settings.STAFF_INFO_SHEET);

  if (!staffSheet) throw new Error(`Reset failed: Staff Info sheet "${settings.STAFF_INFO_SHEET}" not found.`);
  
  const validationErrors = validateStaffInfoSheet(staffSheet, settings);
  if (validationErrors.length > 0) {
      throw new Error(`Reset failed due to validation errors in "${settings.STAFF_INFO_SHEET}":\n- ${validationErrors.join('\n- ')}`);
  }

  const backedUpRecords = createRotationCountBackup(staffSheet, settings);
  const recordsModified = resetRotationCounts(staffSheet);

  console.log("--- Reset All Counts Process Finished ---");
  return { recordsReset: recordsModified, backupRecords: backedUpRecords };
}


// --- Main Generation Function ---

/**
 * Orchestrates the entire schedule generation process from start to finish.
 */
function generateSchedule() {
  const startTime = new Date();
  logAction('Generate Schedule Attempt', 'Starting process.', 'INFO');
  const ui = SpreadsheetApp.getUi();
  let settings; // Declare settings here to be available in catch block

  try {
    // 1. Load Settings and Staff Info
    console.log("Loading settings and staff info...");
    settings = getSettings(); // Assign here
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const staffInfoSheet = ss.getSheetByName(settings.STAFF_INFO_SHEET);
    if (!staffInfoSheet) throw new Error(`Staff Info sheet "${settings.STAFF_INFO_SHEET}" not found.`);
    
    // getStaffInfo now includes validation and will throw an error on failure
    let staffInfo = getStaffInfo(staffInfoSheet, settings);
    logAction('Staff Info Loaded', `Loaded data for ${Object.keys(staffInfo).length} staff.`, 'INFO', settings);

    // 2. Generate Schedule
    console.log("Creating optimized schedule...");
    const { schedule: newSchedule, staffInfo: updatedStaffInfo } = createOptimizedSchedule(staffInfo, settings); // Pass settings
    logAction('Schedule Generation Complete', 'In-memory schedule created.', 'INFO', settings);

    // 3. Show Preview
    console.log("Showing Schedule Preview...");
    showSchedulePreview(newSchedule, settings); // Pass settings

    // 4. Confirm Application
    const response = ui.alert(
      'Apply New Schedule?',
      `Review the preview. Apply this schedule to "${settings.SCHEDULE_SHEET}" and update counts in "${settings.STAFF_INFO_SHEET}"?`,
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      logAction('User Confirmation', 'User confirmed YES to apply schedule.', 'INFO', settings);

      // 5. Apply to Sheet and Update Counts
      applyScheduleToSheet(newSchedule, settings); // Pass settings
      updateStaffInfoSheet(updatedStaffInfo, settings); // Pass settings
      logAction('Schedule Applied & Staff Updated', 'Schedule written to sheet and counts updated.', 'INFO', settings);

      // 6. Check for Rotation Cycle Completion
      const rotationCycleComplete = checkRotationCycleCompletion(updatedStaffInfo, settings); // Pass settings
      if (rotationCycleComplete) {
        logAction('Rotation Cycle Complete', 'All active staff have rotated. Prompting reset.', 'INFO', settings);
        const resetConfirm = ui.alert(
            'Rotation Cycle Complete!',
            `All active staff have completed a rotation.\n\nDo you want to BACKUP and RESET all counts now?`,
            ui.ButtonSet.YES_NO
        );
        if (resetConfirm === ui.Button.YES) {
            resetAllCounts(settings); // Pass settings
            logAction('Rotation Reset Success', 'User confirmed reset after cycle completion.', 'INFO', settings);
            ui.alert('Rotation counts have been reset.');
        } else {
            logAction('Rotation Reset Skipped', 'User chose not to reset counts.', 'INFO', settings);
        }
      }

      const duration = (new Date().getTime() - startTime.getTime()) / 1000;
      ui.alert('Success!', `New schedule applied and staff info updated.\n(Duration: ${duration.toFixed(1)}s)`, ui.ButtonSet.OK);
      logAction('Generate Schedule Success', `Process completed successfully in ${duration.toFixed(1)}s.`, 'INFO', settings);

    } else {
      logAction('Generate Schedule Cancelled', 'User chose NO after preview.', 'INFO', settings);
      ui.alert('Schedule application cancelled.');
    }

  } catch (error) {
    console.error('FATAL ERROR in generateSchedule:', error.stack);
    // Use the settings object if it was loaded, otherwise, log without it.
    logAction('Generate Schedule FATAL Error', error.message, 'ERROR', settings || null); 
    ui.alert('Error', `Process failed:\n${error.message}`, ui.ButtonSet.OK);
   }
}


// --- Custom Menu ---

/**
 * Creates a custom menu in the Google Sheet UI when the spreadsheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Shift Manager Pro')
      .addItem('➡️ Generate New Schedule', 'generateSchedule')
      .addSeparator()
      .addItem('📊 View Shift Statistics', 'displayShiftStatistics')
      .addItem('✏️ Manual Shift Editor', 'showManualShiftEditor')
      .addSeparator()
      .addItem('🔄 Reset All Counts...', 'showResetConfirmation')
      .addToUi();
}

// =================================================================
// END OF SECTION 4
// =================================================================
