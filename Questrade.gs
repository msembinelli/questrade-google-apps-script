var QuestradeApiSession = function(refreshToken) {

  var data = {
    'grant_type' : 'refresh_token',
    'refresh_token' : refreshToken
  };
  var options = {
    'method' : 'POST',
    'payload' : data
  };
  
  var url = 'https://login.questrade.com/oauth2/token'
  var authData = JSON.parse(UrlFetchApp.fetch(url, options).getContentText());
  this.apiServer = authData['api_server'];
  this.accessToken = authData['access_token'];
  this.refreshToken = authData['refresh_token'];
  this.authHeader = {
    'Authorization' : 'Bearer ' + this.accessToken
  };
  
  options = {
    'method' : 'GET',
    'headers' : this.authHeader
  };
  var url = this.apiServer + 'v1/accounts';
  var accountsData  = JSON.parse(UrlFetchApp.fetch(url, options).getContentText());
  this.accounts = accountsData['accounts'];
  
  this.getRefreshToken = function() {
    return this.refreshToken;
  }
  
  this.getPositions = function() {
    var options = {
      'method' : 'GET',
      'headers' : this.authHeader
    };
    for(var i = 0; i < this.accounts.length; i++)
    {
      var url = this.apiServer + 'v1/accounts/' + this.accounts[i]['number'] + '/positions';
      writeJSONtoSheet(JSON.parse(UrlFetchApp.fetch(url, options).getContentText())['positions'],"Positions");
    }
  }
  
  this.getBalances = function() {
    var options = {
      'method' : 'GET',
      'headers' : this.authHeader
    };
    for(var i = 0; i < this.accounts.length; i++)
    {
      var url = this.apiServer + 'v1/accounts/' + this.accounts[i]['number'] + '/balances';
      writeJSONtoSheet(JSON.parse(UrlFetchApp.fetch(url, options).getContentText())['perCurrencyBalances'], "Balances");
    }
  }

  this.authRevoke = function() {
    var data = {
      'token' : this.accessToken
    };
    var options = {
      'method' : 'POST',
      'payload' : data
    };
    var url = 'https://login.questrade.com/oauth2/revoke'
    UrlFetchApp.fetch(url, options);
  }
}

function writeJSONtoSheet(json, sheetname) {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = doc.getSheetByName(sheetname);
  var keys = Object.keys(json).sort();
  if (keys.length < 1)
  {
    // Nothing to do, return
    return;
  }
  
  var last = sheet.getLastColumn();
  var header = [];
  if (last != 0) {
    header = sheet.getRange(1, 1, 1, last).getValues()[0];
  }
  var newCols = [];
 
  if (header.length == 0) {
    for (var k = 0; k < keys.length; k++) {
      var header_keys = Object.keys(json[keys[k]]);
      for (var h = 0; h < header_keys.length; h++) {
        if (newCols.indexOf(header_keys[h]) === -1 && header.indexOf(header_keys[h]) === -1) {
          newCols.push(header_keys[h]);
        }
      }
    }
    sheet.appendRow(newCols);
    header = newCols;
  }

 
  var rows = [];
 
  for (var i = 0; i < keys.length; i++) {
    var row = [];
    for (var h = 0; h < header.length; h++) {
      row.push(header[h] in json[keys[i]] ? json[keys[i]][header[h]] : "");
    }
    if (row.length > 0) {
      rows.push(row);
    }
  }
  
  // We want to erase everything below the headers so that we get new data
  sheet.deleteRows(2, sheet.getLastRow());

  if (rows.length > 0) {
    for (var j = 0; j < rows.length; j++) {
      sheet.appendRow(rows[j]);
    }
  }
  
}

var SheetDbName = 'DataStore';

function initSheetDb(password) {
  // Sheet that stores refresh token
  var Db = new SimpleSheetDb(SheetDbName, password);
  
  // Init QT session, read out positions and balances to sheet
  try {
    var qt = new QuestradeApiSession(Db.simpleSheetDbRead("token"));
    getPositionsAndBalances(qt);
    // Write out the new refresh token to the DB
    Db.simpleSheetDbWrite("token", qt.getRefreshToken());
  }
  catch(e) {
    // Connection failed or bad token, prompt user for new token/password
    doGet(true);
  }

};

function initSheetDbNewToken(password, token) {
  var Db = new SimpleSheetDb(SheetDbName, password);
  
  try {
    var qt = new QuestradeApiSession(token);
    getPositionsAndBalances(qt);
    // Write out the new refresh token to the DB
    Db.simpleSheetDbWrite("token", qt.getRefreshToken());
  }
  catch(e) {
    // Something went wrong
    Logger.log("Connection still failed even with new token...");
  }
}

function getPositionsAndBalances(qt)
{
  qt.getPositions();
  qt.getBalances();
  Logger.log("Wrote balances/positions!");
}

function doGet(needNewToken) {
  if(needNewToken) {
    var html = HtmlService.createHtmlOutputFromFile('PasswordAndToken');
    SpreadsheetApp.getUi().showModalDialog(html, 'Enter Password and Questrade API Token');
  }
  else {
    var html = HtmlService.createHtmlOutputFromFile('Password');
    SpreadsheetApp.getUi().showModalDialog(html, 'Enter Password');
  }
}