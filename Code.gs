const SPREADSHEET_ID = '1_G3iskoNiluovIPPW8X4QzK0GnT_F9TQLbdQjHJBATQ';
const SOURCE_SHEET_NAME = '';
const SOURCE_COLUMN_INDEX = 5; // E
const SUMMARY_SHEET_NAME = 'Сводка причин отказа';
const EMPTY_REASON_LABEL = 'причина не указана';
const STAGE_MARKERS = ['закрыто и не реализовано', 'закрыто и нереализовано'];

/**
 * Builds a summary table of refusal reasons for closed-unrealized deals.
 * Source data is read from column E in detected source sheet.
 * Result is written to SUMMARY_SHEET_NAME as:
 *   причина отказа | количество
 */
function buildRefusalReasonSummary() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sourceSheet = getSourceSheet_(spreadsheet);
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
    const normalized = normalizeText_(value);
    if (!isClosedUnrealizedStage_(normalized)) {
      return;
    }

    const reason = extractReason_(value);
    counts[reason] = (counts[reason] || 0) + 1;
  });

  return counts;
}

/**
 * Checks whether stage value is "closed and unrealized".
 * Supports both variants: "не реализовано" and "нереализовано".
 * @param {string} normalizedValue
 * @return {boolean}
 */
function isClosedUnrealizedStage_(normalizedValue) {
  return STAGE_MARKERS.some((marker) => normalizedValue.includes(marker));
}

/**
 * Normalizes text for stable matching.
 * @param {string} value
 * @return {string}
 */
function normalizeText_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
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

/**
 * Resolves source sheet:
 * 1) by SOURCE_SHEET_NAME (if specified),
 * 2) by header in E1 containing "этап",
 * 3) first sheet that is not summary sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getSourceSheet_(spreadsheet) {
  if (SOURCE_SHEET_NAME) {
    const namedSheet = spreadsheet.getSheetByName(SOURCE_SHEET_NAME);
    if (!namedSheet) {
      throw new Error(`Лист "${SOURCE_SHEET_NAME}" не найден`);
    }

    return namedSheet;
  }

  const sheets = spreadsheet
    .getSheets()
    .filter((sheet) => sheet.getName() !== SUMMARY_SHEET_NAME);

  const byHeader = sheets.find((sheet) => {
    const headerValue = String(sheet.getRange(1, SOURCE_COLUMN_INDEX).getDisplayValue() || '')
      .trim()
      .toLowerCase();
    return headerValue.includes('этап');
  });

  if (byHeader) {
    return byHeader;
  }

  if (sheets.length > 0) {
    return sheets[0];
  }

  throw new Error('Не найден лист-источник для чтения столбца E');
}
