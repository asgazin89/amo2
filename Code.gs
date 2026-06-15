const SPREADSHEET_ID = '19cqiVBuTDFdCUAyAw_1SkuLPEX5kzTOkO6NWDR_xHP8';
const OUTPUT_SHEET_NAME = 'Компании с одной сделкой';

const REQUIRED_COLUMNS = {
  dealName: 'Название сделки',
  budget: 'Бюджет',
  stage: 'Этап сделки',
  closeDate: 'Дата закрытия',
};

/**
 * Builds a report with deals where company is mentioned only once.
 * Company is extracted from "Название сделки".
 */
function buildSingleDealCompaniesReport() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sourceSheet = findSourceSheet_(spreadsheet);
  const outputSheet = getOrCreateSheet_(spreadsheet, OUTPUT_SHEET_NAME);

  const lastRow = sourceSheet.getLastRow();
  const lastColumn = sourceSheet.getLastColumn();

  const headers = sourceSheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  const sourceRows =
    lastRow > 1 ? sourceSheet.getRange(2, 1, lastRow - 1, lastColumn).getDisplayValues() : [];

  const columnIndexes = resolveRequiredIndexes_(headers);
  const contactIndexes = findContactColumnIndexes_(headers);

  const companyCounts = {};
  const preparedRows = sourceRows.map((row) => {
    const dealName = String(row[columnIndexes.dealName] || '').trim();
    const companyName = extractCompanyNameFromDeal_(dealName);
    companyCounts[companyName] = (companyCounts[companyName] || 0) + 1;

    return {row, companyName};
  });

  const outputHeaders = [
    REQUIRED_COLUMNS.dealName,
    REQUIRED_COLUMNS.budget,
    REQUIRED_COLUMNS.stage,
    REQUIRED_COLUMNS.closeDate,
  ].concat(contactIndexes.map((index) => headers[index]));

  const outputRows = preparedRows
    .filter((entry) => companyCounts[entry.companyName] === 1)
    .map((entry) => {
      const row = entry.row;
      return [
        row[columnIndexes.dealName],
        row[columnIndexes.budget],
        row[columnIndexes.stage],
        row[columnIndexes.closeDate],
      ].concat(contactIndexes.map((index) => row[index]));
    });

  const resultRows = [outputHeaders].concat(outputRows);
  writeTable_(outputSheet, resultRows);
}

/**
 * Finds source sheet containing required columns.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function findSourceSheet_(spreadsheet) {
  const requiredHeaderNames = Object.values(REQUIRED_COLUMNS);
  const sheet = spreadsheet.getSheets().find((candidate) => {
    const lastColumn = candidate.getLastColumn();
    if (lastColumn === 0) {
      return false;
    }

    const headers = candidate.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
    return requiredHeaderNames.every((headerName) => headers.includes(headerName));
  });

  if (!sheet) {
    throw new Error('Не найден лист-источник с нужными колонками');
  }

  return sheet;
}

/**
 * Resolves indexes of required columns.
 * @param {string[]} headers
 * @return {{dealName:number, budget:number, stage:number, closeDate:number}}
 */
function resolveRequiredIndexes_(headers) {
  const indexes = {
    dealName: headers.indexOf(REQUIRED_COLUMNS.dealName),
    budget: headers.indexOf(REQUIRED_COLUMNS.budget),
    stage: headers.indexOf(REQUIRED_COLUMNS.stage),
    closeDate: headers.indexOf(REQUIRED_COLUMNS.closeDate),
  };

  if (Object.values(indexes).some((index) => index < 0)) {
    throw new Error('В исходном листе отсутствуют обязательные колонки');
  }

  return indexes;
}

/**
 * Finds contact columns by header patterns.
 * @param {string[]} headers
 * @return {number[]}
 */
function findContactColumnIndexes_(headers) {
  const explicitHeaders = new Set(['Основной контакт', 'Компания контакта']);
  const contactPattern = /(контакт|email|e-mail|телефон|phone|telegram|whatsapp|факс)/i;

  const indexes = [];
  headers.forEach((header, index) => {
    const headerText = String(header || '').trim();
    if (!headerText) {
      return;
    }

    if (explicitHeaders.has(headerText) || contactPattern.test(headerText)) {
      indexes.push(index);
    }
  });

  return indexes;
}

/**
 * Extract company name from deal name.
 * Example:
 *   "664К. ФКП \"НПЦ \"Дельта\"" -> "ФКП \"НПЦ \"Дельта\""
 * @param {string} dealName
 * @return {string}
 */
function extractCompanyNameFromDeal_(dealName) {
  const normalizedDealName = String(dealName || '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!normalizedDealName) {
    return 'компания не указана';
  }

  const companyName = normalizedDealName
    .replace(/^\s*\d+\s*[a-zа-яё-]*\s*[.)\-–—:]?\s*(?:\d+\s*[a-zа-яё-]*\s*[.)\-–—:]?\s*)*/i, '')
    .trim();
  return companyName || normalizedDealName;
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
 * Clears and writes a 2D table to sheet with header formatting.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Array<Array<*>>} rows
 */
function writeTable_(sheet, rows) {
  const safeRows = rows.length > 0 ? rows : [['нет данных']];

  sheet.clearContents();
  sheet.getRange(1, 1, safeRows.length, safeRows[0].length).setValues(safeRows);
  sheet.getRange(1, 1, 1, safeRows[0].length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, safeRows[0].length);
}
