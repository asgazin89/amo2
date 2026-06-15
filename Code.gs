var SPREADSHEET_ID = '19cqiVBuTDFdCUAyAw_1SkuLPEX5kzTOkO6NWDR_xHP8';
var OUTPUT_SHEET_NAME = 'Компании с одной сделкой';
var NO_COMPANY_LABEL = 'компания не указана';

function buildSingleDealCompaniesReport() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  }

  var sourceSheet = findSourceSheet_(spreadsheet);
  var outputSheet = getOrCreateSheet_(spreadsheet, OUTPUT_SHEET_NAME);

  var lastRow = sourceSheet.getLastRow();
  var lastColumn = sourceSheet.getLastColumn();
  if (lastColumn < 1) {
    throw new Error('Исходный лист пустой');
  }

  var headers = sourceSheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  var data = [];
  if (lastRow > 1) {
    data = sourceSheet.getRange(2, 1, lastRow - 1, lastColumn).getDisplayValues();
  }

  var idxDeal = headers.indexOf('Название сделки');
  var idxBudget = headers.indexOf('Бюджет');
  var idxStage = headers.indexOf('Этап сделки');
  var idxClose = headers.indexOf('Дата закрытия');

  if (idxDeal < 0 || idxBudget < 0 || idxStage < 0 || idxClose < 0) {
    throw new Error('Не найдены обязательные колонки: Название сделки / Бюджет / Этап сделки / Дата закрытия');
  }

  var contactIndexes = getContactIndexes_(headers);
  var companyCounts = {};
  var prepared = [];
  var i;

  for (i = 0; i < data.length; i++) {
    var row = data[i];
    var companyName = extractCompanyNameFromDeal_(row[idxDeal]);
    companyCounts[companyName] = (companyCounts[companyName] || 0) + 1;
    prepared.push({
      row: row,
      companyName: companyName,
    });
  }

  var outHeaders = ['Название сделки', 'Бюджет', 'Этап сделки', 'Дата закрытия'];
  for (i = 0; i < contactIndexes.length; i++) {
    outHeaders.push(headers[contactIndexes[i]]);
  }

  var outRows = [outHeaders];
  for (i = 0; i < prepared.length; i++) {
    if (companyCounts[prepared[i].companyName] !== 1) {
      continue;
    }

    var out = [
      prepared[i].row[idxDeal],
      prepared[i].row[idxBudget],
      prepared[i].row[idxStage],
      prepared[i].row[idxClose],
    ];

    var j;
    for (j = 0; j < contactIndexes.length; j++) {
      out.push(prepared[i].row[contactIndexes[j]]);
    }

    outRows.push(out);
  }

  if (outRows.length === 1) {
    outRows.push(['нет данных', '', '', '']);
  }

  writeTable_(outputSheet, outRows);
}

function findSourceSheet_(spreadsheet) {
  var sheets = spreadsheet.getSheets();
  var i;

  for (i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var lastColumn = sheet.getLastColumn();
    if (lastColumn < 1) {
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

function getContactIndexes_(headers) {
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

function extractCompanyNameFromDeal_(dealName) {
  var text = String(dealName || '').replace(/\s+/g, ' ').trim();
  if (!text) {
    return NO_COMPANY_LABEL;
  }

  var company = text.replace(/^\s*\d+\s*[A-Za-zА-Яа-яЁё-]*\s*[.)\-:]?\s*(?:\d+\s*[A-Za-zА-Яа-яЁё-]*\s*[.)\-:]?\s*)*/, '').trim();
  return company || text;
}

function getOrCreateSheet_(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (sheet) {
    return sheet;
  }
  return spreadsheet.insertSheet(sheetName);
}

function writeTable_(sheet, rows) {
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, rows[0].length);
}
