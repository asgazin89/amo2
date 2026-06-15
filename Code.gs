var SPREADSHEET_ID = '19cqiVBuTDFdCUAyAw_1SkuLPEX5kzTOkO6NWDR_xHP8';
var OUTPUT_SHEET_NAME = 'Компании с одной сделкой';
var NO_COMPANY_LABEL = 'компания не указана';

/**
 * Главная функция: строит отчет по компаниям с одной сделкой.
 */
function buildSingleDealCompaniesReport() {
  var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sourceSheet = findSourceSheet(spreadsheet);
  var outputSheet = getOrCreateSheet(spreadsheet, OUTPUT_SHEET_NAME);

  var lastRow = sourceSheet.getLastRow();
  var lastColumn = sourceSheet.getLastColumn();
  if (lastColumn === 0) {
    throw new Error('Исходный лист пустой');
  }

  var headers = sourceSheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  var rows = [];
  if (lastRow > 1) {
    rows = sourceSheet.getRange(2, 1, lastRow - 1, lastColumn).getDisplayValues();
  }

  var idxDeal = headers.indexOf('Название сделки');
  var idxBudget = headers.indexOf('Бюджет');
  var idxStage = headers.indexOf('Этап сделки');
  var idxCloseDate = headers.indexOf('Дата закрытия');

  if (idxDeal < 0 || idxBudget < 0 || idxStage < 0 || idxCloseDate < 0) {
    throw new Error('Не найдены обязательные колонки: Название сделки / Бюджет / Этап сделки / Дата закрытия');
  }

  var contactIndexes = getContactIndexes(headers);
  var companyCounts = {};
  var prepared = [];
  var i;

  for (i = 0; i < rows.length; i++) {
    var dealName = String(rows[i][idxDeal] || '').trim();
    var companyName = extractCompanyName(dealName);
    companyCounts[companyName] = (companyCounts[companyName] || 0) + 1;
    prepared.push({
      row: rows[i],
      companyName: companyName,
    });
  }

  var outputHeaders = ['Название сделки', 'Бюджет', 'Этап сделки', 'Дата закрытия'];
  for (i = 0; i < contactIndexes.length; i++) {
    outputHeaders.push(headers[contactIndexes[i]]);
  }

  var outputRows = [outputHeaders];
  for (i = 0; i < prepared.length; i++) {
    var item = prepared[i];
    if (companyCounts[item.companyName] !== 1) {
      continue;
    }

    var out = [
      item.row[idxDeal],
      item.row[idxBudget],
      item.row[idxStage],
      item.row[idxCloseDate],
    ];

    var j;
    for (j = 0; j < contactIndexes.length; j++) {
      out.push(item.row[contactIndexes[j]]);
    }
    outputRows.push(out);
  }

  if (outputRows.length === 1) {
    outputRows.push(['нет данных', '', '', '']);
  }

  writeTable(outputSheet, outputRows);
}

/**
 * Ищет лист с нужными заголовками.
 */
function findSourceSheet(spreadsheet) {
  var sheets = spreadsheet.getSheets();
  var i;
  for (i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var lastColumn = sheet.getLastColumn();
    if (lastColumn === 0) {
      continue;
    }

    var headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
    if (
      headers.indexOf('Название сделки') >= 0 &&
      headers.indexOf('Бюджет') >= 0 &&
      headers.indexOf('Этап сделки') >= 0 &&
      headers.indexOf('Дата закрытия') >= 0
    ) {
      return sheet;
    }
  }
  throw new Error('Не найден лист-источник с нужными колонками');
}

/**
 * Возвращает индексы всех контактных колонок.
 */
function getContactIndexes(headers) {
  var indexes = [];
  var i;

  for (i = 0; i < headers.length; i++) {
    var header = String(headers[i] || '').trim();
    if (!header) {
      continue;
    }

    if (header === 'Основной контакт' || header === 'Компания контакта') {
      indexes.push(i);
      continue;
    }

    if (
      /контакт/i.test(header) ||
      /email/i.test(header) ||
      /e-mail/i.test(header) ||
      /телефон/i.test(header) ||
      /phone/i.test(header) ||
      /telegram/i.test(header) ||
      /whatsapp/i.test(header) ||
      /факс/i.test(header)
    ) {
      indexes.push(i);
    }
  }

  return indexes;
}

/**
 * Извлекает название компании из "Название сделки".
 */
function extractCompanyName(dealName) {
  var normalized = String(dealName || '').replace(/\s+/g, ' ').trim();
  if (!normalized) {
    return NO_COMPANY_LABEL;
  }

  var company = normalized
    .replace(/^\s*\d+\s*[a-zа-яё-]*\s*[.)\-–—:]?\s*(?:\d+\s*[a-zа-яё-]*\s*[.)\-–—:]?\s*)*/i, '')
    .trim();

  return company || normalized;
}

/**
 * Возвращает лист или создает новый.
 */
function getOrCreateSheet(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (sheet) {
    return sheet;
  }
  return spreadsheet.insertSheet(sheetName);
}

/**
 * Записывает таблицу на лист.
 */
function writeTable(sheet, rows) {
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, rows[0].length);
}
