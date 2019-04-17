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
    this.accounts = accountsData.accounts;

    this.getPositions = function () {
        const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
            'method': 'get',
            'headers': this.authHeader
        };
        let table = {
            rows: [],
            namedRanges: [],
        };
        const sheetName = "Positions";
        this.accounts.forEach(account => {
            var url = this.authData.api_server + 'v1/accounts/' + account.number + '/positions';
            table = writeJsonToTable(JSON.parse(UrlFetchApp.fetch(url, options).getContentText())['positions'], sheetName, table, account);
        });
        writeTableToSheet(sheetName, table);
        writeNamedRangesToSheet(sheetName, table);
    }

    this.getBalances = function () {
        const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
            'method': 'get',
            'headers': this.authHeader
        };
        let table = {
            rows: [],
            namedRanges: [],
        };
        const sheetName = "Balances";
        this.accounts.forEach(account => {
            var url = this.authData.api_server + 'v1/accounts/' + account.number + '/balances';
            table = writeJsonToTable(JSON.parse(UrlFetchApp.fetch(url, options).getContentText())['perCurrencyBalances'], sheetName, table, account);
        });
        writeTableToSheet(sheetName, table);
        writeNamedRangesToSheet(sheetName, table);
    }
}

function objectValues(obj) {
    const keys = Object.keys(obj);
    return keys.map(key => obj[key]);
}

function writeJsonToTable(json, sheetName, table, account) {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getSheetByName(sheetName);

    if (sheet == null) {
        sheet = doc.insertSheet(sheetName);
    }

    if (json === undefined || json.length == 0) {
        return table;
    }

    const entries = json.sort();

    // Get account row
    table.rows.push(objectValues(account));

    // Header for one object should be the same for all
    const header = Object.keys(entries[0]);
    table.rows.push(header);

    // Get values of all entries in a 2D array
    entries.forEach(entry => table.rows.push(objectValues(entry)));

    // Insert empty row for spacing
    table.rows.push([]);

    const lastNamedRange = table.namedRanges[table.namedRanges.length - 1];
    let startRow = 3;
    let startCol = 1;
    const numRows = entries.length;
    const numCols = header.length;
    if (lastNamedRange !== undefined && table.namedRanges.length > 0) {
        // If a previous named range exists, start range at next section (new account)
        startRow = startRow + lastNamedRange.startRow + lastNamedRange.numRows;
    }

    table.namedRanges.push({
        name: account.type + "_" + account.number + "_" + sheetName,
        startRow,
        numRows,
        range: sheet.getRange(startRow, startCol, numRows, numCols),
    });

    return table;
}

function writeTableToSheet(sheetName, table) {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getSheetByName(sheetName);

    sheet.clear();

    let maxColumnLength = 0;
    table.rows.forEach(row => {
        if (row.length > maxColumnLength) {
            maxColumnLength = row.length;
        }
    });

    // Normalize all rows so they are the same length (required by setValues)
    table.rows.forEach(row => {
        const difference = maxColumnLength - row.length;
        if (difference > 0) {
            let i;
            for (i = 0; i < difference; i++) {
                row.push("");
            }
        }
    });

    // Write table to sheet
    var range = sheet.getRange(1, 1, table.rows.length, maxColumnLength);
    range.setValues(table.rows);
}

function writeNamedRangesToSheet(sheetName, table) {
    var doc = SpreadsheetApp.getActiveSpreadsheet();

    table.namedRanges.forEach(namedRange =>
        doc.setNamedRange(namedRange.name, namedRange.range)
    )
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
        .addItem('Pull', 'run')
        .addToUi();
}

