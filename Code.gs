const SPREADSHEET_ID = '1_G3iskoNiluovIPPW8X4QzK0GnT_F9TQLbdQjHJBATQ';
const SOURCE_SHEET_NAME = '';
const ID_COLUMN_INDEX = 1; // A
const SOURCE_COLUMN_INDEX = 5; // E
const BUDGET_COLUMN_INDEX = 6; // F
const DATE_COLUMN_INDEX = 7; // G (Дата создания)
const SUMMARY_SHEET_NAME = 'Сводка причин отказа';
const PROFIT_SUMMARY_SHEET_NAME = 'Сводка причин отказа и прибыли';
const MONTHLY_SUMMARY_SHEET_NAME = 'Сводка причин по месяцам';
const CHARTS_SHEET_NAME = 'Графики причин отказа';
const EMPTY_REASON_LABEL = 'причина не указана';
const STAGE_MARKERS = ['закрыто и не реализовано', 'закрыто и нереализовано'];

/**
 * Builds a summary table of refusal reasons for closed-unrealized deals.
 * Source data is read from column E in detected source sheet.
 * Result is written to:
 * - SUMMARY_SHEET_NAME: причина отказа | количество | процент от общего числа сделок
 * - PROFIT_SUMMARY_SHEET_NAME:
 *   причина отказа | количество | недополученная прибыль | процент от общего числа сделок
 * - MONTHLY_SUMMARY_SHEET_NAME:
 *   отдельная таблица на каждую причину отказа с помесячными метриками
 * - CHARTS_SHEET_NAME:
 *   2 больших графика (количество и недополученная прибыль) по всем причинам
 */
function buildRefusalReasonSummary() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sourceSheet = getSourceSheet_(spreadsheet);
  const summarySheet = getOrCreateSheet_(spreadsheet, SUMMARY_SHEET_NAME);
  const profitSummarySheet = getOrCreateSheet_(spreadsheet, PROFIT_SUMMARY_SHEET_NAME);
  const monthlySummarySheet = getOrCreateSheet_(spreadsheet, MONTHLY_SUMMARY_SHEET_NAME);
  const chartsSheet = getOrCreateSheet_(spreadsheet, CHARTS_SHEET_NAME);

  const lastRow = sourceSheet.getLastRow();
  const summaryRows = [['причина отказа', 'количество', 'процент от общего числа сделок']];
  const profitSummaryRows = [
    ['причина отказа', 'количество', 'недополученная прибыль', 'процент от общего числа сделок'],
  ];

  if (lastRow > 1) {
    const sourceRows = sourceSheet
      .getRange(2, ID_COLUMN_INDEX, lastRow - 1, DATE_COLUMN_INDEX - ID_COLUMN_INDEX + 1)
      .getValues();

    const totalDeals = countDealsById_(sourceRows);
    const totalDealsByMonth = countDealsByMonth_(sourceRows);
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

    const monthlyRows = buildMonthlySummaryRows_(sortedStats, totalDealsByMonth);
    writeTable_(monthlySummarySheet, monthlyRows);
    applyMonthlySummaryFormats_(monthlySummarySheet, monthlyRows);
    rebuildChartsSheet_(chartsSheet, sortedStats, totalDealsByMonth);
  } else {
    writeTable_(monthlySummarySheet, [['нет данных', '', '', '']]);
    rebuildChartsSheet_(chartsSheet, [], {});
  }

  writeTable_(summarySheet, summaryRows);
  writeTable_(profitSummarySheet, profitSummaryRows);
  summarySheet.getRange(2, 3, Math.max(summaryRows.length - 1, 1), 1).setNumberFormat('0.00%');
  profitSummarySheet.getRange(2, 3, Math.max(profitSummaryRows.length - 1, 1), 1).setNumberFormat('#,##0.00');
  profitSummarySheet.getRange(2, 4, Math.max(profitSummaryRows.length - 1, 1), 1).setNumberFormat('0.00%');
}

/**
 * Collects per-reason totals and per-month stats for closed-unrealized deals.
 * @param {Array<Array<*>>} sourceRows rows from columns A:G
 * @return {Object<string, {count:number, lostProfit:number, months:Object<string, {count:number, lostProfit:number}>}>}
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
    const monthKey = parseMonthKey_(row[DATE_COLUMN_INDEX - ID_COLUMN_INDEX]);

    if (!statsByReason[reason]) {
      statsByReason[reason] = {count: 0, lostProfit: 0, months: {}};
    }

    statsByReason[reason].count += 1;
    statsByReason[reason].lostProfit += budget;

    if (!monthKey) {
      return;
    }

    if (!statsByReason[reason].months[monthKey]) {
      statsByReason[reason].months[monthKey] = {count: 0, lostProfit: 0};
    }

    statsByReason[reason].months[monthKey].count += 1;
    statsByReason[reason].months[monthKey].lostProfit += budget;
  });

  return statsByReason;
}

/**
 * Counts total number of deals by non-empty ID in column A.
 * @param {Array<Array<*>>} sourceRows rows from columns A:G
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
 * Counts total deals by month using non-empty ID in column A and creation date in G.
 * @param {Array<Array<*>>} sourceRows rows from columns A:G
 * @return {Object<string, number>}
 */
function countDealsByMonth_(sourceRows) {
  const totals = {};

  sourceRows.forEach((row) => {
    const idValue = row[0];
    const hasId =
      (typeof idValue === 'number' && Number.isFinite(idValue)) ||
      String(idValue || '').trim() !== '';
    if (!hasId) {
      return;
    }

    const monthKey = parseMonthKey_(row[DATE_COLUMN_INDEX - ID_COLUMN_INDEX]);
    if (!monthKey) {
      return;
    }

    totals[monthKey] = (totals[monthKey] || 0) + 1;
  });

  return totals;
}

/**
 * Builds rows for monthly summary sheet.
 * @param {Array<[string, {count:number, lostProfit:number, months:Object<string, {count:number, lostProfit:number}>}]>} sortedStats
 * @param {Object<string, number>} totalDealsByMonth
 * @return {Array<Array<*>>}
 */
function buildMonthlySummaryRows_(sortedStats, totalDealsByMonth) {
  const monthKeys = Object.keys(totalDealsByMonth).sort();
  if (sortedStats.length === 0 || monthKeys.length === 0) {
    return [['нет данных', '', '', '']];
  }

  const rows = [];

  sortedStats.forEach(([reason, stats], reasonIndex) => {
    rows.push([`Причина отказа: ${reason}`, '', '', '']);
    rows.push(['месяц', 'количество', 'процент от общего числа сделок', 'недополученная прибыль']);

    monthKeys.forEach((monthKey) => {
      const monthStats = stats.months[monthKey] || {count: 0, lostProfit: 0};
      const monthTotalDeals = totalDealsByMonth[monthKey] || 0;
      const monthShare = monthTotalDeals > 0 ? monthStats.count / monthTotalDeals : 0;

      rows.push([formatMonthKey_(monthKey), monthStats.count, monthShare, monthStats.lostProfit]);
    });

    if (reasonIndex < sortedStats.length - 1) {
      rows.push(['', '', '', '']);
    }
  });

  return rows;
}

/**
 * Applies number formats for monthly summary columns.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Array<Array<*>>} rows
 */
function applyMonthlySummaryFormats_(sheet, rows) {
  if (rows.length === 0) {
    return;
  }

  sheet.getRange(1, 3, rows.length, 1).setNumberFormat('0.00%');
  sheet.getRange(1, 4, rows.length, 1).setNumberFormat('#,##0.00');

  rows.forEach((row, index) => {
    const isReasonTitle =
      String(row[0] || '').startsWith('Причина отказа:') &&
      String(row[1] || '') === '' &&
      String(row[2] || '') === '' &&
      String(row[3] || '') === '';
    const isTableHeader = String(row[0] || '').toLowerCase() === 'месяц';

    if (isReasonTitle || isTableHeader) {
      sheet.getRange(index + 1, 1, 1, 4).setFontWeight('bold');
    }
  });
}

/**
 * Rebuilds chart sheet with two large charts by refusal reason:
 * count and lost profit by month.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Array<[string, {count:number, lostProfit:number, months:Object<string, {count:number, lostProfit:number}>}]>} sortedStats
 * @param {Object<string, number>} totalDealsByMonth
 */
function rebuildChartsSheet_(sheet, sortedStats, totalDealsByMonth) {
  sheet.clear();
  sheet.getCharts().forEach((chart) => sheet.removeChart(chart));

  const monthKeys = Object.keys(totalDealsByMonth).sort();
  if (sortedStats.length === 0 || monthKeys.length === 0) {
    sheet.getRange(1, 1).setValue('Нет данных для построения графиков');
    return;
  }

  const reasons = sortedStats.map(([reason]) => reason);
  const countTable = [['месяц'].concat(reasons)];
  const profitTable = [['месяц'].concat(reasons)];

  const statsByReason = {};
  sortedStats.forEach(([reason, stats]) => {
    statsByReason[reason] = stats;
  });

  monthKeys.forEach((monthKey) => {
    const monthLabel = formatMonthKey_(monthKey);
    const countRow = [monthLabel];
    const profitRow = [monthLabel];

    reasons.forEach((reason) => {
      const monthStats = (statsByReason[reason].months || {})[monthKey] || {count: 0, lostProfit: 0};
      countRow.push(monthStats.count);
      profitRow.push(monthStats.lostProfit);
    });

    countTable.push(countRow);
    profitTable.push(profitRow);
  });

  const tableStartColumn = 20; // T
  const countRange = sheet.getRange(1, tableStartColumn, countTable.length, countTable[0].length);
  const profitStartRow = countTable.length + 4;
  const profitRange = sheet.getRange(profitStartRow, tableStartColumn, profitTable.length, profitTable[0].length);

  countRange.setValues(countTable);
  profitRange.setValues(profitTable);
  countRange.getCell(1, 1).offset(0, 0, 1, countTable[0].length).setFontWeight('bold');
  profitRange.getCell(1, 1).offset(0, 0, 1, profitTable[0].length).setFontWeight('bold');
  sheet
    .getRange(2, tableStartColumn + 1, Math.max(countTable.length - 1, 1), Math.max(countTable[0].length - 1, 1))
    .setNumberFormat('0');
  sheet
    .getRange(
      profitStartRow + 1,
      tableStartColumn + 1,
      Math.max(profitTable.length - 1, 1),
      Math.max(profitTable[0].length - 1, 1),
    )
    .setNumberFormat('#,##0.00');

  const countChart = sheet
    .newChart()
    .asLineChart()
    .addRange(countRange)
    .setNumHeaders(1)
    .setPosition(1, 1, 0, 0)
    .setOption('title', 'Количество сделок по причинам отказа')
    .setOption('legend', {position: 'right'})
    .setOption('hAxis', {title: 'Месяц'})
    .setOption('vAxis', {title: 'Количество'})
    .setOption('width', 1400)
    .setOption('height', 520)
    .setOption('chartArea', {left: 90, top: 50, width: '62%', height: '72%'})
    .build();

  const profitChart = sheet
    .newChart()
    .asLineChart()
    .addRange(profitRange)
    .setNumHeaders(1)
    .setPosition(28, 1, 0, 0)
    .setOption('title', 'Недополученная прибыль по причинам отказа')
    .setOption('legend', {position: 'right'})
    .setOption('hAxis', {title: 'Месяц'})
    .setOption('vAxis', {title: 'Недополученная прибыль'})
    .setOption('width', 1400)
    .setOption('height', 520)
    .setOption('chartArea', {left: 90, top: 50, width: '62%', height: '72%'})
    .build();

  sheet.insertChart(countChart);
  sheet.insertChart(profitChart);
}

/**
 * Parses month key (YYYY-MM) from date value.
 * Supports Date object and string like "dd.mm.yyyy hh:mm:ss".
 * @param {*} value
 * @return {string}
 */
function parseMonthKey_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !Number.isNaN(value.getTime())) {
    return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, '0')}`;
  }

  const raw = String(value || '').trim();
  if (!raw) {
    return '';
  }

  const datePart = raw.split(' ')[0];
  const match = datePart.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);
  if (!match) {
    return '';
  }

  const month = Number(match[2]);
  const year = Number(match[3]);
  if (!Number.isFinite(month) || !Number.isFinite(year) || month < 1 || month > 12) {
    return '';
  }

  return `${year}-${String(month).padStart(2, '0')}`;
}

/**
 * Formats month key YYYY-MM to human-readable Russian label.
 * @param {string} monthKey
 * @return {string}
 */
function formatMonthKey_(monthKey) {
  const [yearText, monthText] = String(monthKey || '').split('-');
  const month = Number(monthText);
  const year = Number(yearText);
  const monthNames = [
    'январь',
    'февраль',
    'март',
    'апрель',
    'май',
    'июнь',
    'июль',
    'август',
    'сентябрь',
    'октябрь',
    'ноябрь',
    'декабрь',
  ];

  if (!Number.isFinite(month) || !Number.isFinite(year) || month < 1 || month > 12) {
    return monthKey;
  }

  return `${monthNames[month - 1]} ${year}`;
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
        sheet.getName() !== SUMMARY_SHEET_NAME &&
        sheet.getName() !== PROFIT_SUMMARY_SHEET_NAME &&
        sheet.getName() !== MONTHLY_SUMMARY_SHEET_NAME &&
        sheet.getName() !== CHARTS_SHEET_NAME,
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
