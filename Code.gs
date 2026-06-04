const SPREADSHEET_ID = '1_G3iskoNiluovIPPW8X4QzK0GnT_F9TQLbdQjHJBATQ';
const SOURCE_SHEET_NAME = '';
const ID_COLUMN_INDEX = 1; // A
const DEAL_NAME_COLUMN_INDEX = 2; // B
const COMPANY_COLUMN_INDEX = 3; // C
const SOURCE_COLUMN_INDEX = 5; // E
const BUDGET_COLUMN_INDEX = 6; // F
const DATE_COLUMN_INDEX = 7; // G (Дата создания)
const SUMMARY_SHEET_NAME = 'Сводка причин отказа';
const PROFIT_SUMMARY_SHEET_NAME = 'Сводка причин отказа и прибыли';
const MONTHLY_SUMMARY_SHEET_NAME = 'Сводка причин по месяцам';
const SUCCESS_SUMMARY_SHEET_NAME = 'Сводка успешных сделок';
const SUCCESS_COMPANY_BUDGET_SHEET_NAME = 'Успешные сделки по компаниям (сумма)';
const SUCCESS_COMPANY_COUNT_SHEET_NAME = 'Успешные сделки по компаниям (количество)';
const CONTROL_SHEET_NAME = '⚡ Обновление отчетов';
const CONTROL_BUTTON_CELL = 'B4';
const EMPTY_COMPANY_LABEL = 'компания не указана';
const EMPTY_REASON_LABEL = 'причина не указана';
const STAGE_MARKERS = ['закрыто и не реализовано', 'закрыто и нереализовано'];
const SUCCESS_STATUS_ORDER = ['успешно реализовано', 'отправка', 'закрытие договора'];
const AUTO_REFRESH_HANDLER = 'refreshAllReports';
const AUTO_REFRESH_CHANGE_HANDLER = 'refreshAllReportsOnChange';
const MIN_AUTO_REFRESH_INTERVAL_MS = 30000;
const LAST_AUTO_REFRESH_KEY = 'lastAutoRefreshTs';

/**
 * Adds custom menu buttons to spreadsheet UI.
 */
function onOpen() {
  setupRefreshControlSheet_();

  SpreadsheetApp.getUi()
    .createMenu('Отчеты CRM')
    .addItem('Обновить все отчеты', 'runRefreshFromUi')
    .addSeparator()
    .addItem('Включить автообновление', 'installAutoRefreshTriggers')
    .addItem('Отключить автообновление', 'removeAutoRefreshTriggers')
    .addToUi();
}

/**
 * Manual helper to recreate the visible refresh control sheet.
 * Use this if the control sheet is missing.
 */
function setupRefreshButton() {
  setupRefreshControlSheet_();
}

/**
 * Builds a summary table of refusal reasons for closed-unrealized deals.
 * Source data is read from column E in detected source sheet.
 * Result is written to:
 * - SUMMARY_SHEET_NAME: причина отказа | количество | процент от общего числа сделок
 * - PROFIT_SUMMARY_SHEET_NAME:
 *   причина отказа | количество | недополученная прибыль | процент от общего числа сделок
 * - MONTHLY_SUMMARY_SHEET_NAME:
 *   отдельная таблица на каждую причину отказа с помесячными метриками
 *   + 2 больших графика (количество и недополученная прибыль) по всем причинам
 */
function buildRefusalReasonSummary() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SPREADSHEET_ID);
  const sourceSheet = getSourceSheet_(spreadsheet);
  const summarySheet = getOrCreateSheet_(spreadsheet, SUMMARY_SHEET_NAME);
  const profitSummarySheet = getOrCreateSheet_(spreadsheet, PROFIT_SUMMARY_SHEET_NAME);
  const monthlySummarySheet = getOrCreateSheet_(spreadsheet, MONTHLY_SUMMARY_SHEET_NAME);
  const successSummarySheet = getOrCreateSheet_(spreadsheet, SUCCESS_SUMMARY_SHEET_NAME);
  const successCompanyBudgetSheet = getOrCreateSheet_(spreadsheet, SUCCESS_COMPANY_BUDGET_SHEET_NAME);
  const successCompanyCountSheet = getOrCreateSheet_(spreadsheet, SUCCESS_COMPANY_COUNT_SHEET_NAME);

  const lastRow = sourceSheet.getLastRow();
  const sourceRows =
    lastRow > 1
      ? sourceSheet
          .getRange(2, ID_COLUMN_INDEX, lastRow - 1, DATE_COLUMN_INDEX - ID_COLUMN_INDEX + 1)
          .getValues()
      : [];
  const summaryRows = [['причина отказа', 'количество', 'процент от общего числа сделок']];
  const profitSummaryRows = [
    ['причина отказа', 'количество', 'недополученная прибыль', 'процент от общего числа сделок'],
  ];

  if (sourceRows.length > 0) {
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
    rebuildChartsOnMonthlySheet_(monthlySummarySheet, sortedStats, totalDealsByMonth, monthlyRows.length);
  } else {
    writeTable_(monthlySummarySheet, [['нет данных', '', '', '']]);
    rebuildChartsOnMonthlySheet_(monthlySummarySheet, [], {}, 1);
  }

  const successRows = buildSuccessfulSummaryRows_(sourceRows);
  writeTable_(successSummarySheet, successRows);
  successSummarySheet
    .getRange(2, 3, Math.max(successRows.length - 1, 1), 1)
    .setNumberFormat('#,##0.00');
  successSummarySheet
    .getRange(2, 4, Math.max(successRows.length - 1, 1), 1)
    .setNumberFormat('0.00%');
  successSummarySheet
    .getRange(successRows.length, 1, 1, successRows[0].length)
    .setFontWeight('bold');

  const companySuccessStats = collectSuccessfulCompanyStats_(sourceRows);
  const companyBudgetRows = buildCompanySuccessfulRows_(companySuccessStats, 'budget');
  const companyCountRows = buildCompanySuccessfulRows_(companySuccessStats, 'count');
  writeTable_(successCompanyBudgetSheet, companyBudgetRows);
  writeTable_(successCompanyCountSheet, companyCountRows);
  successCompanyBudgetSheet
    .getRange(2, 2, Math.max(companyBudgetRows.length - 1, 1), 4)
    .setNumberFormat('#,##0.00');
  successCompanyCountSheet
    .getRange(2, 2, Math.max(companyCountRows.length - 1, 1), 4)
    .setNumberFormat('0');

  writeTable_(summarySheet, summaryRows);
  writeTable_(profitSummarySheet, profitSummaryRows);
  summarySheet.getRange(2, 3, Math.max(summaryRows.length - 1, 1), 1).setNumberFormat('0.00%');
  profitSummarySheet.getRange(2, 3, Math.max(profitSummaryRows.length - 1, 1), 1).setNumberFormat('#,##0.00');
  profitSummarySheet.getRange(2, 4, Math.max(profitSummaryRows.length - 1, 1), 1).setNumberFormat('0.00%');
}

/**
 * Main public entrypoint to recalculate all report sheets.
 * Can be used in triggers and manual runs.
 */
function refreshAllReports() {
  runRefreshWithLock_('manual');
}

/**
 * UI action for manual report refresh from spreadsheet menu.
 */
function runRefreshFromUi() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SPREADSHEET_ID);
  spreadsheet.toast('Идет обновление отчетов...', 'Отчеты CRM', 4);

  refreshAllReports();

  spreadsheet.toast('Отчеты успешно обновлены', 'Отчеты CRM', 5);
}

/**
 * Simple trigger that runs on direct user edits in the spreadsheet.
 * Works without explicit trigger installation.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function onEdit(e) {
  if (isManualRefreshButtonEdit_(e)) {
    runRefreshFromControlButton_(e);
    return;
  }

  runRefreshWithLock_('onEdit', e);
}

/**
 * Installable trigger handler for spreadsheet change events.
 * @param {GoogleAppsScript.Events.SheetsOnChange} e
 */
function refreshAllReportsOnChange(e) {
  if (shouldSkipChangeEvent_(e)) {
    return;
  }

  runRefreshWithLock_('onChange', e);
}

/**
 * Installs auto-refresh triggers:
 * - onEdit for instant updates after changes
 * - hourly time trigger as fallback refresh
 *
 * Run once manually to enable automatic updates.
 */
function installAutoRefreshTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => {
    const handlerName = trigger.getHandlerFunction();
    if (handlerName === AUTO_REFRESH_HANDLER || handlerName === AUTO_REFRESH_CHANGE_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger(AUTO_REFRESH_HANDLER).forSpreadsheet(SPREADSHEET_ID).onEdit().create();
  ScriptApp.newTrigger(AUTO_REFRESH_CHANGE_HANDLER).forSpreadsheet(SPREADSHEET_ID).onChange().create();
  ScriptApp.newTrigger(AUTO_REFRESH_HANDLER).timeBased().everyMinutes(5).create();
}

/**
 * Removes report auto-refresh triggers created by installAutoRefreshTriggers().
 */
function removeAutoRefreshTriggers() {
  ScriptApp.getProjectTriggers()
    .filter((trigger) => {
      const handlerName = trigger.getHandlerFunction();
      return handlerName === AUTO_REFRESH_HANDLER || handlerName === AUTO_REFRESH_CHANGE_HANDLER;
    })
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));
}

/**
 * Handles refresh from a visible checkbox button on control sheet.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function runRefreshFromControlButton_(e) {
  if (!e || !e.range) {
    return;
  }

  const spreadsheet = e.range.getSheet().getParent();
  spreadsheet.toast('Идет обновление отчетов...', 'Отчеты CRM', 4);
  try {
    runRefreshWithLock_('manualButton', e);
    spreadsheet.toast('Отчеты успешно обновлены', 'Отчеты CRM', 5);
  } catch (error) {
    spreadsheet.toast(`Ошибка обновления: ${error.message}`, 'Отчеты CRM', 8);
    throw error;
  } finally {
    e.range.setValue(false);
  }
}

/**
 * Refresh wrapper with lock and lightweight debounce for frequent edits.
 * @param {string} source
 * @param {Object=} eventObject
 */
function runRefreshWithLock_(source, eventObject) {
  if (source === 'onEdit' && shouldSkipAutoRefresh_()) {
    return;
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    return;
  }

  try {
    buildRefusalReasonSummary();
    markAutoRefreshTimestamp_();
  } catch (error) {
    console.error(`refresh failed from ${source}`, error);
    if (eventObject) {
      throw error;
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * Returns true when edited cell is the visible refresh checkbox button.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 * @return {boolean}
 */
function isManualRefreshButtonEdit_(e) {
  if (!e || !e.range || String(e.value || '').toUpperCase() !== 'TRUE') {
    return false;
  }

  const sheet = e.range.getSheet();
  return (
    sheet.getName() === CONTROL_SHEET_NAME &&
    e.range.getA1Notation() === CONTROL_BUTTON_CELL
  );
}

/**
 * Creates/updates a visible control sheet with manual refresh button.
 */
function setupRefreshControlSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet_(spreadsheet, CONTROL_SHEET_NAME);

  sheet.clear();
  sheet.setFrozenRows(1);
  sheet.getRange('A1:E1').merge();
  sheet.getRange('A1').setValue('Управление обновлением отчетов');
  sheet.getRange('A2:E2').merge();
  sheet.getRange('A2').setValue('Нажми на чекбокс в B4, чтобы запустить полный пересчет всех отчетов.');
  sheet.getRange('A4').setValue('обновить данные');
  sheet.getRange(CONTROL_BUTTON_CELL).insertCheckboxes();
  sheet.getRange(CONTROL_BUTTON_CELL).setValue(false);
  sheet.getRange('C4:E4').merge();
  sheet.getRange('C4').setValue('⬅ обновить данные');

  sheet.getRange('A1').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
  sheet.getRange('A2').setWrap(true);
  sheet.getRange('A4').setFontWeight('bold').setFontSize(12);
  sheet.getRange(CONTROL_BUTTON_CELL).setBackground('#34a853').setFontColor('#ffffff').setFontWeight('bold');
  sheet.getRange('C4').setFontWeight('bold').setFontColor('#1a73e8');
  sheet.autoResizeColumns(1, 5);
  sheet.setColumnWidths(1, 5, 180);
  sheet.setRowHeights(1, 4, 38);
}

/**
 * Returns true when previous refresh happened too recently.
 * Helps avoid trigger flooding on rapid edits.
 * @return {boolean}
 */
function shouldSkipAutoRefresh_() {
  const props = PropertiesService.getScriptProperties();
  const now = Date.now();
  const prev = Number(props.getProperty(LAST_AUTO_REFRESH_KEY) || 0);
  return Number.isFinite(prev) && now - prev < MIN_AUTO_REFRESH_INTERVAL_MS;
}

/**
 * Persists successful refresh timestamp.
 */
function markAutoRefreshTimestamp_() {
  PropertiesService.getScriptProperties().setProperty(LAST_AUTO_REFRESH_KEY, String(Date.now()));
}

/**
 * Filters noisy onChange events to avoid recursive reruns from formatting/charts.
 * @param {GoogleAppsScript.Events.SheetsOnChange=} e
 * @return {boolean}
 */
function shouldSkipChangeEvent_(e) {
  if (!e || !e.changeType) {
    return false;
  }

  const skipTypes = ['FORMAT', 'OTHER', 'INSERT_GRID', 'REMOVE_GRID'];
  return skipTypes.includes(String(e.changeType));
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
 * Rebuilds two large charts on monthly summary sheet by refusal reason:
 * count and lost profit by month.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Array<[string, {count:number, lostProfit:number, months:Object<string, {count:number, lostProfit:number}>}]>} sortedStats
 * @param {Object<string, number>} totalDealsByMonth
 * @param {number} monthlyRowsCount
 */
function rebuildChartsOnMonthlySheet_(sheet, sortedStats, totalDealsByMonth, monthlyRowsCount) {
  sheet.getCharts().forEach((chart) => sheet.removeChart(chart));

  const monthKeys = Object.keys(totalDealsByMonth).sort();
  if (sortedStats.length === 0 || monthKeys.length === 0) {
    sheet.getRange(1, 6).setValue('Нет данных для построения графиков');
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

  const tableStartColumn = 8; // H
  const countRange = sheet.getRange(1, tableStartColumn, countTable.length, countTable[0].length);
  const profitStartRow = countTable.length + 3;
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
    .setPosition(monthlyRowsCount + 2, 1, 0, 0)
    .setOption('title', 'Количество сделок по причинам отказа')
    .setOption('legend', {position: 'right'})
    .setOption('hAxis', {title: 'Месяц'})
    .setOption('vAxis', {title: 'Количество'})
    .setOption('width', 1500)
    .setOption('height', 560)
    .setOption('chartArea', {left: 90, top: 50, width: '62%', height: '72%'})
    .build();

  const profitChart = sheet
    .newChart()
    .asLineChart()
    .addRange(profitRange)
    .setNumHeaders(1)
    .setPosition(monthlyRowsCount + 31, 1, 0, 0)
    .setOption('title', 'Недополученная прибыль по причинам отказа')
    .setOption('legend', {position: 'right'})
    .setOption('hAxis', {title: 'Месяц'})
    .setOption('vAxis', {title: 'Недополученная прибыль'})
    .setOption('width', 1500)
    .setOption('height', 560)
    .setOption('chartArea', {left: 90, top: 50, width: '62%', height: '72%'})
    .build();

  sheet.insertChart(countChart);
  sheet.insertChart(profitChart);
}

/**
 * Builds summary rows for successful deal statuses.
 * @param {Array<Array<*>>} sourceRows rows from columns A:G
 * @return {Array<Array<*>>}
 */
function buildSuccessfulSummaryRows_(sourceRows) {
  const statsByStatus = {};
  SUCCESS_STATUS_ORDER.forEach((status) => {
    statsByStatus[status] = {count: 0, budget: 0};
  });

  sourceRows.forEach((row) => {
    const stageValue = normalizeSuccessfulStatus_(row[SOURCE_COLUMN_INDEX - ID_COLUMN_INDEX]);
    const statusKey = resolveSuccessfulStatusKey_(stageValue);
    if (!statusKey) {
      return;
    }

    statsByStatus[statusKey].count += 1;
    statsByStatus[statusKey].budget += parseBudget_(row[BUDGET_COLUMN_INDEX - ID_COLUMN_INDEX]);
  });

  const totalDeals = countDealsById_(sourceRows);
  let successfulDealsTotal = 0;
  let successfulBudgetTotal = 0;
  const rows = [['статус', 'количество сделок', 'общая сумма бюджета', 'процент от общего числа сделок']];
  SUCCESS_STATUS_ORDER.forEach((statusKey) => {
    const statusCount = statsByStatus[statusKey].count;
    const statusBudget = statsByStatus[statusKey].budget;
    const statusShare = totalDeals > 0 ? statusCount / totalDeals : 0;

    rows.push([statusKey, statusCount, statusBudget, statusShare]);
    successfulDealsTotal += statusCount;
    successfulBudgetTotal += statusBudget;
  });
  rows.push([
    'итого успешные сделки (конверсия из лида)',
    successfulDealsTotal,
    successfulBudgetTotal,
    totalDeals > 0 ? successfulDealsTotal / totalDeals : 0,
  ]);

  return rows;
}

/**
 * Collects successful-deal stats by company and status.
 * @param {Array<Array<*>>} sourceRows rows from columns A:G
 * @return {Object<string, {statusBudgets:Object<string, number>, statusCounts:Object<string, number>, totalBudget:number, totalCount:number}>}
 */
function collectSuccessfulCompanyStats_(sourceRows) {
  const statsByCompany = {};

  sourceRows.forEach((row) => {
    const stageValue = normalizeSuccessfulStatus_(row[SOURCE_COLUMN_INDEX - ID_COLUMN_INDEX]);
    const statusKey = resolveSuccessfulStatusKey_(stageValue);
    if (!statusKey) {
      return;
    }

    const companyName = resolveCompanyName_(
      row[COMPANY_COLUMN_INDEX - ID_COLUMN_INDEX],
      row[DEAL_NAME_COLUMN_INDEX - ID_COLUMN_INDEX],
    );
    const budget = parseBudget_(row[BUDGET_COLUMN_INDEX - ID_COLUMN_INDEX]);
    if (!statsByCompany[companyName]) {
      statsByCompany[companyName] = {
        statusBudgets: {},
        statusCounts: {},
        totalBudget: 0,
        totalCount: 0,
      };
      SUCCESS_STATUS_ORDER.forEach((status) => {
        statsByCompany[companyName].statusBudgets[status] = 0;
        statsByCompany[companyName].statusCounts[status] = 0;
      });
    }

    statsByCompany[companyName].statusBudgets[statusKey] += budget;
    statsByCompany[companyName].statusCounts[statusKey] += 1;
    statsByCompany[companyName].totalBudget += budget;
    statsByCompany[companyName].totalCount += 1;
  });

  return statsByCompany;
}

/**
 * Builds successful-deal company table rows for budget or count metric.
 * @param {Object<string, {statusBudgets:Object<string, number>, statusCounts:Object<string, number>, totalBudget:number, totalCount:number}>} statsByCompany
 * @param {'budget' | 'count'} metric
 * @return {Array<Array<*>>}
 */
function buildCompanySuccessfulRows_(statsByCompany, metric) {
  const rows = [['компания', 'успешно реализовано', 'отправка', 'закрытие договора', 'итого']];
  const companies = Object.entries(statsByCompany).sort((a, b) => {
    const aTotal = metric === 'budget' ? a[1].totalBudget : a[1].totalCount;
    const bTotal = metric === 'budget' ? b[1].totalBudget : b[1].totalCount;
    return bTotal - aTotal || a[0].localeCompare(b[0], 'ru');
  });

  companies.forEach(([companyName, stats]) => {
    const statusValues = SUCCESS_STATUS_ORDER.map((status) =>
      metric === 'budget' ? stats.statusBudgets[status] : stats.statusCounts[status],
    );
    const totalValue = metric === 'budget' ? stats.totalBudget : stats.totalCount;
    rows.push([companyName].concat(statusValues, [totalValue]));
  });

  if (rows.length === 1) {
    rows.push([EMPTY_COMPANY_LABEL, 0, 0, 0, 0]);
  }

  return rows;
}

/**
 * Resolves company name from C column, fallback to parsed value from B.
 * @param {*} companyCellValue
 * @param {*} dealNameCellValue
 * @return {string}
 */
function resolveCompanyName_(companyCellValue, dealNameCellValue) {
  const companyFromC = normalizeCompanyName_(companyCellValue);
  if (companyFromC) {
    return companyFromC;
  }

  const dealName = normalizeCompanyName_(dealNameCellValue);
  if (!dealName) {
    return EMPTY_COMPANY_LABEL;
  }

  const fromDealName = normalizeCompanyName_(
    dealName.replace(/^\s*\d+\s*[a-zа-яё-]*\s*[.)\-–—:]?\s*/i, ''),
  );
  return fromDealName || dealName || EMPTY_COMPANY_LABEL;
}

/**
 * Normalizes company text.
 * @param {*} value
 * @return {string}
 */
function normalizeCompanyName_(value) {
  return String(value || '')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Normalizes successful status value for matching.
 * @param {*} value
 * @return {string}
 */
function normalizeSuccessfulStatus_(value) {
  return normalizeText_(value).replace('реализованно', 'реализовано');
}

/**
 * Maps normalized stage to one of supported successful statuses.
 * @param {string} normalizedStage
 * @return {string}
 */
function resolveSuccessfulStatusKey_(normalizedStage) {
  if (normalizedStage === 'успешно реализовано') {
    return 'успешно реализовано';
  }

  if (normalizedStage === 'отправка') {
    return 'отправка';
  }

  if (normalizedStage === 'закрытие договора') {
    return 'закрытие договора';
  }

  return '';
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
        sheet.getName() !== SUCCESS_SUMMARY_SHEET_NAME &&
        sheet.getName() !== SUCCESS_COMPANY_BUDGET_SHEET_NAME &&
        sheet.getName() !== SUCCESS_COMPANY_COUNT_SHEET_NAME &&
        sheet.getName() !== CONTROL_SHEET_NAME,
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
