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
        for (var i = 0; i < this.accounts.length; i++) {
            var url = this.authData.api_server + 'v1/accounts/' + this.accounts[i]['number'] + '/positions';
            writeJSONtoSheet(JSON.parse(UrlFetchApp.fetch(url, options).getContentText())['positions'], "Positions");
        }
    }

    this.getBalances = function () {
        const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
            'method': 'get',
            'headers': this.authHeader
        };
        for (var i = 0; i < this.accounts.length; i++) {
            var url = this.authData.api_server + 'v1/accounts/' + this.accounts[i]['number'] + '/balances';
            writeJSONtoSheet(JSON.parse(UrlFetchApp.fetch(url, options).getContentText())['combinedBalances'], "Balances");
        }
    }
}

function writeJSONtoSheet(json, sheetname) {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getSheetByName(sheetname);
    var keys = Object.keys(json).sort();
    if (keys.length < 1) {
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
    sheet.deleteRows(2, sheet.getLastRow() - 1);

    if (rows.length > 0) {
        for (var j = 0; j < rows.length; j++) {
            sheet.appendRow(rows[j]);
        }
    }

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