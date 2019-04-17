function authorization() {
    var storedTokenData = {
        refresh_token: ''
    };
    try {
        storedTokenData = JSON.parse(PropertiesService.getUserProperties().getProperty('qt'));
    } catch (e) {
        getNewToken();
        return undefined;
    }
    if (storedTokenData == null) {
        getNewToken();
        return undefined;
    }

    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
        'method': 'post',
        'payload': {
            'refresh_token': storedTokenData.refresh_token,
            'grant_type': 'refresh_token',
        }
    };

    const url = PropertiesService.getScriptProperties().getProperty('tokenUrl')
    try {
        var tokenObj = JSON.parse(UrlFetchApp.fetch(url, options).getContentText());
    } catch (e) {
        getNewToken();
        return undefined;
    }
    PropertiesService.getUserProperties().setProperty('qt', JSON.stringify(tokenObj));
    return tokenObj;
}

function getNewToken() {
    var html = HtmlService.createHtmlOutputFromFile('EnterToken');
    SpreadsheetApp.getUi().showModalDialog(html, 'Enter Questrade API Token');
}

function saveNewToken(refresh_token) {
    var storeData = {
        'refresh_token': refresh_token
    }
    PropertiesService.getUserProperties().setProperty('qt', JSON.stringify(storeData));
    run();
}

var QuestradeApiSession = function () {
    this.authData = authorization();
    if (this.authData == undefined) {
        return;
    }
    this.authHeader = {
        'Authorization': 'Bearer ' + this.authData.access_token
    };

    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
        'method': 'get',
        'headers': this.authHeader
    };
    var url = this.authData.api_server + 'v1/accounts';
    var accountsData = JSON.parse(UrlFetchApp.fetch(url, options).getContentText());
    this.accounts = accountsData['accounts'];

    this.getPositions = function () {
        const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
            'method': 'get',
            'headers': this.authHeader
        };
        var row = 1;
        for (var i = 0; i < this.accounts.length; i++) {
            var url = this.authData.api_server + 'v1/accounts/' + this.accounts[i]['number'] + '/positions';
            row = writeJSONtoSheet(JSON.parse(UrlFetchApp.fetch(url, options).getContentText())['positions'], "Positions", this.accounts[i], row);
        }
    }

    this.getBalances = function () {
        const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
            'method': 'get',
            'headers': this.authHeader
        };
        var row = 1;
        for (var i = 0; i < this.accounts.length; i++) {
            var url = this.authData.api_server + 'v1/accounts/' + this.accounts[i]['number'] + '/balances';
            row = writeJSONtoSheet(JSON.parse(UrlFetchApp.fetch(url, options).getContentText())['perCurrencyBalances'], "Balances", this.accounts[i], row);
        }
    }
}

function writeJSONtoSheet(json, sheetname, account, currentRow) {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getSheetByName(sheetname);
    var keys = Object.keys(json).sort();

    if (sheet == null) {
        sheet = doc.insertSheet(sheetname);
    }

    if (keys.length < 1) {
        console.error("Nothing to do, return");
        return currentRow;
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var header = [];
    for (var k = 0; k < keys.length; k++) {
        var header_keys = Object.keys(json[keys[k]]);
        for (var h = 0; h < header_keys.length; h++) {
            if (header.indexOf(header_keys[h]) === -1) {
                header.push(header_keys[h]);
            }
        }
    }
    if (currentRow == 1 && lastRow) {
        sheet.getRange(1, 1, lastRow, lastCol).clear();
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
    sheet.getRange(currentRow, 1, 1, 2).setValues([[account['type'], account['number']]]);
    currentRow++;
    sheet.getRange(currentRow, 1, 1, header.length).setValues([header]);
    currentRow++;
    if (rows.length > 0) {
        var range = sheet.getRange(currentRow, 1, keys.length, header.length);
        range.setValues(rows);
        doc.setNamedRange(account['type'] + "_" + account['number'] + "_" + sheetname, range);
    }
    currentRow += keys.length;
    currentRow++;
    return currentRow;
}

function getPositionsAndBalances(qt) {
    qt.getPositions();
    qt.getBalances();
    Logger.log("Wrote balances/positions!");
}

function run() {
    // Init QT session, read out positions and balances to sheet
    try {
        getPositionsAndBalances(new QuestradeApiSession());
    }
    catch (e) {
        // Connection failed or bad token, prompt user for new token/password
        console.log(e);
    }
};

function onOpen(e) {
    // Add a custom menu to the spreadsheet.
    SpreadsheetApp.getUi() // Or DocumentApp, SlidesApp, or FormApp.
        .createMenu('Questrade')
        .addItem('Run', 'run')
        .addToUi();
}

