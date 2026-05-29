const SPREADSHEET_ID = '1_G3iskoNiluovIPPW8X4QzK0GnT_F9TQLbdQjHJBATQ';
const SOURCE_SHEET_NAME = '';
const ID_COLUMN_INDEX = 1; // A
const SOURCE_COLUMN_INDEX = 5; // E
const BUDGET_COLUMN_INDEX = 6; // F
const SUMMARY_SHEET_NAME = 'Сводка причин отказа';
const PROFIT_SUMMARY_SHEET_NAME = 'Сводка причин отказа и прибыли';
const EMPTY_REASON_LABEL = 'причина не указана';
const STAGE_MARKERS = ['закрыто и не реализовано', 'закрыто и нереализовано'];

/**
 * Builds a summary table of refusal reasons for closed-unrealized deals.
 * Source data is read from column E in detected source sheet.
 * Result is written to:
 * - SUMMARY_SHEET_NAME: причина отказа | количество | процент от общего числа сделок
 * - PROFIT_SUMMARY_SHEET_NAME:
 *   причина отказа | количество | недополученная прибыль | процент от общего числа сделок
 */
function buildRefusalReasonSummary() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sourceSheet = getSourceSheet_(spreadsheet);
  const summarySheet = getOrCreateSheet_(spreadsheet, SUMMARY_SHEET_NAME);
  const profitSummarySheet = getOrCreateSheet_(spreadsheet, PROFIT_SUMMARY_SHEET_NAME);

  const lastRow = sourceSheet.getLastRow();
  const summaryRows = [['причина отказа', 'количество', 'процент от общего числа сделок']];
  const profitSummaryRows = [
    ['причина отказа', 'количество', 'недополученная прибыль', 'процент от общего числа сделок'],
  ];

  if (lastRow > 1) {
    const sourceRows = sourceSheet
      .getRange(2, ID_COLUMN_INDEX, lastRow - 1, BUDGET_COLUMN_INDEX - ID_COLUMN_INDEX + 1)
      .getValues();

    const totalDeals = countDealsById_(sourceRows);
    const reasonStats = collectReasonStats_(sourceRows);
    const sortedStats = Object.entries(reasonStats).sort(
      (a, b) =>
        b[1].count - a[1].count ||
        b[1].lostProfit - a[1].lostProfit ||
        a[0].localeCompare(b[0], 'ru'),
    );

    sortedStats.forEach(([reason, stats]) => {
      const dealShare = totalDeals > 0 ? stats.count / totalDeals : 0;
      summaryRows.push([reason, stats.count, dealShare]);
      profitSummaryRows.push([reason, stats.count, stats.lostProfit, dealShare]);
    });
  }

  writeTable_(summarySheet, summaryRows);
  writeTable_(profitSummarySheet, profitSummaryRows);
  summarySheet.getRange(2, 3, Math.max(summaryRows.length - 1, 1), 1).setNumberFormat('0.00%');
  profitSummarySheet.getRange(2, 3, Math.max(profitSummaryRows.length - 1, 1), 1).setNumberFormat('#,##0.00');
  profitSummarySheet.getRange(2, 4, Math.max(profitSummaryRows.length - 1, 1), 1).setNumberFormat('0.00%');
}

/**
 * Collects per-reason count and budget sum for closed-unrealized deals.
 * @param {Array<Array<*>>} sourceRows rows from columns A:F
 * @return {Object<string, {count:number, lostProfit:number}>}
 */
function collectReasonStats_(sourceRows) {
  const statsByReason = {};

  sourceRows.forEach((row) => {
    const stageValue = String(row[SOURCE_COLUMN_INDEX - ID_COLUMN_INDEX] || '').trim();
    const normalized = normalizeText_(stageValue);
    if (!isClosedUnrealizedStage_(normalized)) {
      return;
    }

    const reason = extractReason_(stageValue);
    const budget = parseBudget_(row[BUDGET_COLUMN_INDEX - ID_COLUMN_INDEX]);
    if (!statsByReason[reason]) {
      statsByReason[reason] = {count: 0, lostProfit: 0};
    }

    statsByReason[reason].count += 1;
    statsByReason[reason].lostProfit += budget;
  });

  return statsByReason;
}

/**
 * Counts total number of deals by non-empty ID in column A.
 * @param {Array<Array<*>>} sourceRows rows from columns A:F
 * @return {number}
 */
function countDealsById_(sourceRows) {
  return sourceRows.filter((row) => {
    const idValue = row[0];
    if (typeof idValue === 'number') {
      return Number.isFinite(idValue);
    }

    return String(idValue || '').trim() !== '';
  }).length;
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
 * Parses budget value from sheet cell.
 * Empty or invalid values are treated as zero.
 * @param {*} value
 * @return {number}
 */
function parseBudget_(value) {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value;
  }

  const normalized = String(value || '')
    .replace(/\s+/g, '')
    .replace(',', '.');
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
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
 * Clears and writes a 2D table to sheet with bold header and autoresize.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Array<Array<*>>} rows
 */
function writeTable_(sheet, rows) {
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
  sheet.autoResizeColumns(1, rows[0].length);
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
    .filter(
      (sheet) =>
        sheet.getName() !== SUMMARY_SHEET_NAME && sheet.getName() !== PROFIT_SUMMARY_SHEET_NAME,
    );

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
