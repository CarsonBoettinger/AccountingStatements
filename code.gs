function getIncomeStatement(symbol) {
  var apiKey = 'YOURAPIKEY';  // Replace with your Alpha Vantage API key
  var url = `https://www.alphavantage.co/query?function=INCOME_STATEMENT&symbol=${symbol}&apikey=${apiKey}`;
  var response = UrlFetchApp.fetch(url);
  var data = JSON.parse(response.getContentText());
  var incomeStatement = data["annualReports"];
  
  if (incomeStatement) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.clear();
    var headers = Object.keys(incomeStatement[0]);
    sheet.appendRow(headers);
    
    incomeStatement.forEach(function(report) {
      var row = headers.map(function(header) {
        return report[header];
      });
      sheet.appendRow(row);
    });
  } else {
    Logger.log("No data found for the symbol: " + symbol);
  }
}

function getBalanceSheet(symbol) {
  var apiKey = 'YOURAPIKEY';  // Replace with your Alpha Vantage API key
  var url = `https://www.alphavantage.co/query?function=BALANCE_SHEET&symbol=${symbol}&apikey=${apiKey}`;
  var response = UrlFetchApp.fetch(url);
  var data = JSON.parse(response.getContentText());
  var balanceSheet = data["annualReports"];
  
  if (balanceSheet) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.clear();
    var headers = Object.keys(balanceSheet[0]);
    sheet.appendRow(headers);
    
    balanceSheet.forEach(function(report) {
      var row = headers.map(function(header) {
        return report[header];
      });
      sheet.appendRow(row);
    });
  } else {
    Logger.log("No data found for the symbol: " + symbol);
  }
}

function getCashFlow(symbol) {
  var apiKey = 'YOURAPIKEY';  // Replace with your Alpha Vantage API key
  var url = `https://www.alphavantage.co/query?function=CASH_FLOW&symbol=${symbol}&apikey=${apiKey}`;
  var response = UrlFetchApp.fetch(url);
  var data = JSON.parse(response.getContentText());
  var cashFlow = data["annualReports"];
  
  if (cashFlow) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.clear();
    var headers = Object.keys(cashFlow[0]);
    sheet.appendRow(headers);
    
    cashFlow.forEach(function(report) {
      var row = headers.map(function(header) {
        return report[header];
      });
      sheet.appendRow(row);
    });
  } else {
    Logger.log("No data found for the symbol: " + symbol);
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Alpha Vantage')
    .addItem('Get Income Statement', 'promptForIncomeStatement')
    .addItem('Get Balance Sheet', 'promptForBalanceSheet')
    .addItem('Get Cash Flow', 'promptForCashFlow')
    .addToUi();
}

function promptForIncomeStatement() {
  var symbol = Browser.inputBox('Enter stock symbol for Income Statement:');
  getIncomeStatement(symbol);
}

function promptForBalanceSheet() {
  var symbol = Browser.inputBox('Enter stock symbol for Balance Sheet:');
  getBalanceSheet(symbol);
}

function promptForCashFlow() {
  var symbol = Browser.inputBox('Enter stock symbol for Cash Flow:');
  getCashFlow(symbol);
}

