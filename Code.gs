var SPREADSHEET_ID = '19cqiVBuTDFdCUAyAw_1SkuLPEX5kzTOkO6NWDR_xHP8';
var OUTPUT_SHEET_NAME = 'Компании с одной сделкой';
var NO_COMPANY_LABEL = 'компания не указана';
var LEGAL_FORMS = ['ООО', 'АО', 'ПАО', 'ЗАО', 'ОАО', 'ИП', 'ФКП', 'ФГУП', 'ГБУ', 'ГКУ', 'МУП'];

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
  var idxOpen = headers.indexOf('Дата создания');
  var idxClose = headers.indexOf('Дата закрытия');

  if (idxDeal < 0 || idxBudget < 0 || idxStage < 0 || idxOpen < 0 || idxClose < 0) {
    throw new Error(
      'Не найдены обязательные колонки: Название сделки / Бюджет / Этап сделки / Дата создания / Дата закрытия',
    );
  }

  var contactIndexes = getContactIndexes_(headers);
  var companyCounts = {};
  var companyByProjectKey = {};
  var projectCompanyCounts = {};
  var prepared = [];
  var i;

  for (i = 0; i < data.length; i++) {
    var row = data[i];
    var dealName = row[idxDeal];
    var companyName = extractCompanyNameFromDeal_(dealName);
    var projectKey = extractProjectKey_(dealName);

    if (projectKey && companyName !== NO_COMPANY_LABEL) {
      if (!projectCompanyCounts[projectKey]) {
        projectCompanyCounts[projectKey] = {};
      }
      projectCompanyCounts[projectKey][companyName] = (projectCompanyCounts[projectKey][companyName] || 0) + 1;
    }

    prepared.push({
      row: row,
      companyName: companyName,
      projectKey: projectKey,
    });
  }

  // Build projectKey -> dominant company map (e.g. 70К.25 -> ООО СПТ).
  var projectKeys = Object.keys(projectCompanyCounts);
  for (i = 0; i < projectKeys.length; i++) {
    var key = projectKeys[i];
    var byCompany = projectCompanyCounts[key];
    var names = Object.keys(byCompany);
    if (names.length === 0) {
      continue;
    }

    var bestName = names[0];
    var j;
    for (j = 1; j < names.length; j++) {
      if (byCompany[names[j]] > byCompany[bestName]) {
        bestName = names[j];
      }
    }
    companyByProjectKey[key] = bestName;
  }

  // Final company assignment + counting.
  for (i = 0; i < prepared.length; i++) {
    if (
      prepared[i].companyName === NO_COMPANY_LABEL &&
      prepared[i].projectKey &&
      companyByProjectKey[prepared[i].projectKey]
    ) {
      prepared[i].companyName = companyByProjectKey[prepared[i].projectKey];
    }

    companyCounts[prepared[i].companyName] = (companyCounts[prepared[i].companyName] || 0) + 1;
  }

  var outHeaders = ['Название сделки', 'Бюджет', 'Этап сделки', 'Дата открытия', 'Дата закрытия'];
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
      prepared[i].row[idxOpen],
      prepared[i].row[idxClose],
    ];

    var j;
    for (j = 0; j < contactIndexes.length; j++) {
      out.push(prepared[i].row[contactIndexes[j]]);
    }

    outRows.push(out);
  }

  if (outRows.length === 1) {
    var emptyRow = [];
    for (i = 0; i < outHeaders.length; i++) {
      emptyRow.push('');
    }
    emptyRow[0] = 'нет данных';
    outRows.push(emptyRow);
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
      headers.indexOf('Дата создания') >= 0 &&
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

  var parts = text.split(' ');
  if (parts.length > 0 && isCodeToken_(parts[0])) {
    parts.shift();
    text = parts.join(' ').trim();
  }
  if (!text) {
    return NO_COMPANY_LABEL;
  }

  var formsPattern = '(?:' + LEGAL_FORMS.join('|') + ')';

  // Cases like "НИЛ АП, ООО".
  var matchCommaForm = text.match(new RegExp('^(.+?,\\s*' + formsPattern + '\\b)', 'i'));
  if (matchCommaForm && matchCommaForm[1]) {
    return normalizeCompanyName_(matchCommaForm[1]);
  }

  // Find first legal form in string and build company name from there.
  var matchForm = text.match(new RegExp('\\b' + formsPattern + '\\b', 'i'));
  if (matchForm) {
    var formStart = text.search(new RegExp('\\b' + formsPattern + '\\b', 'i'));
    var candidate = text.substring(formStart).trim();

    // Best case: quoted legal entity name.
    var matchQuoted = candidate.match(
      new RegExp(
        '^(' + formsPattern + '(?:\\s+[A-Za-zА-Яа-яЁё0-9.-]+){0,3}\\s*[«"]\\s*[^"»«]+\\s*[»"]?)',
        'i',
      ),
    );
    if (matchQuoted && matchQuoted[1]) {
      return normalizeCompanyName_(matchQuoted[1]);
    }

    // Individual entrepreneur without quotes.
    if (/^ИП\b/i.test(candidate)) {
      var matchIp = candidate.match(/^(ИП(?:\s+[A-Za-zА-Яа-яЁё.-]+){1,4})/i);
      if (matchIp && matchIp[1]) {
        return normalizeCompanyName_(matchIp[1]);
      }
    }

    // Generic company name capture until descriptive tail.
    var stopSplit = candidate.split(/\s[-–—/]\s|\s\(|,\s*(?:поставка|монтаж|сборка|жгут|кабел|плата|корпус|камера)/i);
    var companyPart = stopSplit[0];
    var matchOrg = companyPart.match(
      new RegExp('^(' + formsPattern + '(?:\\s+[A-Za-zА-Яа-яЁё0-9.-]+){0,8})', 'i'),
    );
    if (matchOrg && matchOrg[1]) {
      return normalizeCompanyName_(matchOrg[1]);
    }

    var normalizedCandidate = normalizeCompanyName_(candidate);
    return isCompanyNameValid_(normalizedCandidate) ? normalizedCandidate : NO_COMPANY_LABEL;
  }

  var fallback = normalizeCompanyName_(text);
  return isCompanyNameValid_(fallback) ? fallback : NO_COMPANY_LABEL;
}

function extractProjectKey_(dealName) {
  var text = String(dealName || '').replace(/\s+/g, ' ').trim();
  if (!text) {
    return '';
  }

  var firstToken = text.split(' ')[0];
  if (!isCodeToken_(firstToken)) {
    return '';
  }

  var parts = firstToken.split('.');
  if (parts.length >= 2) {
    return parts[0] + '.' + parts[1];
  }

  return firstToken;
}

function isCodeToken_(token) {
  var text = String(token || '').trim();
  if (!text) {
    return false;
  }

  if (!/^[A-Za-zА-Яа-яЁё0-9._\-\/]+$/.test(text)) {
    return false;
  }

  return /\d/.test(text);
}

function normalizeCompanyName_(value) {
  return String(value || '')
    .replace(/[«»"]/g, '')
    .replace(/\s+/g, ' ')
    .replace(/^[\s\-.,;:]+|[\s\-.,;:]+$/g, '')
    .trim();
}

function isCompanyNameValid_(name) {
  var text = String(name || '').trim();
  if (!text) {
    return false;
  }

  if (/^\(.*\)$/.test(text)) {
    return false;
  }

  if (!/[A-Za-zА-Яа-яЁё]/.test(text)) {
    return false;
  }

  return true;
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
