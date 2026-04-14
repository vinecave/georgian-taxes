var SHEET_NAME = 'Income';
var HEADERS = ['Month', 'Date Filled', 'Income Date', 'Currency', 'Gross Revenue', 'Amount GEL', 'Running Sum', 'Tax 1%', 'Exchange Rate'];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Tax Calculator')
    .addItem('Add Income', 'showAddIncomeDialog')
    .addToUi();
}

function getOrCreateSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getLastDataRow_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;
  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== '') return i + 2;
  }
  return 0;
}

function checkIfFirstRow() {
  var sheet = getOrCreateSheet_();
  return getLastDataRow_(sheet) === 0;
}

function showAddIncomeDialog() {
  var html = HtmlService.createHtmlOutputFromFile('Dialog')
    .setWidth(420)
    .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Income');
}

function formatDot(value) {
  return Number(value).toFixed(2);
}

function getExchangeRate(currency, dateString) {
  if (currency === 'GEL') {
    return { rate: 1.0, quantity: 1 };
  }
  var url = 'https://nbg.gov.ge/gw/api/ct/monetarypolicy/currencies/en/json/?currencies=' +
    encodeURIComponent(currency) + '&date=' + encodeURIComponent(dateString);
  var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    throw new Error('NBG API returned status ' + response.getResponseCode());
  }
  var data = JSON.parse(response.getContentText());
  if (!data || !data[0] || !data[0].currencies || !data[0].currencies[0]) {
    throw new Error('No exchange rate data found for ' + currency + ' on ' + dateString);
  }
  var entry = data[0].currencies[0];
  return { rate: entry.rate, quantity: entry.quantity };
}

function addIncomeRow(formData) {
  var sheet = getOrCreateSheet_();
  var lastDataRow = getLastDataRow_(sheet);

  var rateInfo = getExchangeRate(formData.currency, formData.incomeDate);
  var ratePerUnit = rateInfo.rate / rateInfo.quantity;
  var grossRevenue = parseFloat(formData.grossRevenue);
  var amountGel = (formData.currency === 'GEL') ? grossRevenue : grossRevenue * ratePerUnit;
  var tax = amountGel * 0.01;

  var runningSum;
  if (lastDataRow === 0) {
    var initial = parseFloat(formData.initialRunningSum) || 0;
    runningSum = initial + amountGel;
  } else {
    var prevRunning = parseFloat(sheet.getRange(lastDataRow, 7).getValue()) || 0;
    runningSum = prevRunning + amountGel;
  }

  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var newRow = lastDataRow === 0 ? 2 : lastDataRow + 1;

  var numericCols = [6, 7, 8, 9]; // F, G, H, I
  numericCols.forEach(function(col) {
    sheet.getRange(newRow, col).setNumberFormat('@');
  });

  sheet.getRange(newRow, 1).setValue(formData.month);
  sheet.getRange(newRow, 2).setValue(today);
  sheet.getRange(newRow, 3).setValue(formData.incomeDate);
  sheet.getRange(newRow, 4).setValue(formData.currency);
  sheet.getRange(newRow, 5).setNumberFormat('@').setValue(formatDot(grossRevenue));
  sheet.getRange(newRow, 6).setValue(formatDot(amountGel));
  sheet.getRange(newRow, 7).setValue(formatDot(runningSum));
  sheet.getRange(newRow, 8).setValue(formatDot(tax));
  sheet.getRange(newRow, 9).setValue(formatDot(ratePerUnit));

  applyThresholdColor_(sheet, newRow, runningSum);

  return {
    success: true,
    amountGel: formatDot(amountGel),
    runningSum: formatDot(runningSum),
    tax: formatDot(tax),
    exchangeRate: formatDot(ratePerUnit)
  };
}

function applyThresholdColor_(sheet, row, runningSum) {
  var cell = sheet.getRange(row, 7);
  if (runningSum >= 500000) {
    cell.setBackground('#cc0000').setFontColor('#ffffff');
  } else if (runningSum >= 450000) {
    cell.setBackground('#ff9900').setFontColor('#000000');
  }
}
