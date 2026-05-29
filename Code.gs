const SPREADSHEET_ID = '1Ze6Fm1ViaiAE98tjjDCgY6oCNPT_F-y3';
const SOURCE_SHEET_GID = 1858323035;
const SOURCE_COLUMN_INDEX = 5; // E
const SUMMARY_SHEET_NAME = 'Сводка причин отказа';
const EMPTY_REASON_LABEL = 'причина не указана';
const STAGE_MARKER = 'закрыто и нереализовано';

/**
 * Builds a summary table of refusal reasons for closed-unrealized deals.
 * Source data is read from column E in the sheet with SOURCE_SHEET_GID.
 * Result is written to SUMMARY_SHEET_NAME as:
 *   причина отказа | количество
 */
function buildRefusalReasonSummary() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sourceSheet = findSheetByGid_(spreadsheet, SOURCE_SHEET_GID);
  const summarySheet = getOrCreateSheet_(spreadsheet, SUMMARY_SHEET_NAME);

  const lastRow = sourceSheet.getLastRow();
  const summaryRows = [['причина отказа', 'количество']];

  if (lastRow > 1) {
    const stageValues = sourceSheet
      .getRange(2, SOURCE_COLUMN_INDEX, lastRow - 1, 1)
      .getDisplayValues()
      .map((row) => String(row[0] || '').trim());

    const reasonCounts = countReasons_(stageValues);
    Object.entries(reasonCounts)
      .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], 'ru'))
      .forEach(([reason, count]) => summaryRows.push([reason, count]));
  }

  summarySheet.clearContents();
  summarySheet
    .getRange(1, 1, summaryRows.length, summaryRows[0].length)
    .setValues(summaryRows);
  summarySheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  summarySheet.autoResizeColumns(1, 2);
}

/**
 * Counts refusal reasons for values that include STAGE_MARKER.
 * @param {string[]} stageValues
 * @return {Object<string, number>}
 */
function countReasons_(stageValues) {
  const counts = {};

  stageValues.forEach((value) => {
    const normalized = value.toLowerCase();
    if (!normalized.includes(STAGE_MARKER)) {
      return;
    }

    const reason = extractReason_(value);
    counts[reason] = (counts[reason] || 0) + 1;
  });

  return counts;
}

/**
 * Extracts reason from first parentheses block, otherwise fallback label.
 * @param {string} value
 * @return {string}
 */
function extractReason_(value) {
  const match = value.match(/\(([^()]*)\)/);
  if (!match) {
    return EMPTY_REASON_LABEL;
  }

  const reason = String(match[1] || '').trim();
  return reason || EMPTY_REASON_LABEL;
}

/**
 * Finds a sheet in spreadsheet by numeric sheet ID (gid).
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @param {number} gid
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function findSheetByGid_(spreadsheet, gid) {
  const sheet = spreadsheet.getSheets().find((item) => item.getSheetId() === gid);
  if (!sheet) {
    throw new Error(`Лист с gid=${gid} не найден`);
  }

  return sheet;
}

/**
 * Returns existing sheet by name or creates a new one.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @param {string} sheetName
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateSheet_(spreadsheet, sheetName) {
  const existingSheet = spreadsheet.getSheetByName(sheetName);
  if (existingSheet) {
    return existingSheet;
  }

  return spreadsheet.insertSheet(sheetName);
}
