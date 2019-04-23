function reset() {
    var service = getService();
    service.reset();
}
function getService() {
    return OAuth2.createService('questrade')
        .setAuthorizationBaseUrl('https://login.questrade.com/oauth2/authorize')
        .setTokenUrl('https://login.questrade.com/oauth2/token')

        // Set the client ID.
        .setClientId(PropertiesService.getScriptProperties().getProperty('customerKey'))

        // No secret provided by QT. Use dummy one to make oauth2 lib happy.
        .setClientSecret('secret')

        // Set the name of the callback function in the script referenced
        // above that should be invoked to complete the OAuth flow.
        .setCallbackFunction('authCallback')

        // Set the property store where authorized tokens should be persisted.
        .setPropertyStore(PropertiesService.getUserProperties())

        // Set the scopes to request.
        .setScope('read_acc')

        .setParam('response_type', 'code');
}
function authCallback(request) {
    var service = getService();
    var authorized = service.handleCallback(request);
    if (authorized) {
        return HtmlService.createHtmlOutput('Success! <script>setTimeout(function() { top.window.close() }, 1);</script>');
    } else {
        return HtmlService.createHtmlOutput('Denied.');
    }
}
function logRedirectUri() {
    console.log(OAuth2.getRedirectUri());
}
var QuestradeApiSession = function () {
    this.service = getService();
    if (!this.service.hasAccess())
    {
        var authorizationUrl = this.service.getAuthorizationUrl();
        var template = HtmlService.createTemplate(
            '<a href="<?= authorizationUrl ?>" target="_blank">Authorize</a>. ' +
            'Pull again when the authorization is complete.');
        template.authorizationUrl = authorizationUrl;
        var page = template.evaluate();
        SpreadsheetApp.getUi().showSidebar(page);
        return;
    }
    this.authHeader = {
        'Authorization': 'Bearer ' + this.service.getAccessToken()
    };
    this.apiServer = this.service.getToken().api_server;
    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
        'method': 'get',
        'headers': this.authHeader
    };
    var url = this.apiServer + 'v1/accounts';
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
            var url = this.apiServer + 'v1/accounts/' + account.number + '/positions';
            table = writeJsonToTable(JSON.parse(UrlFetchApp.fetch(url, options).getContentText())['positions'], sheetName, table, account);
        });
        writeTableToSheet(sheetName, table);
        writeNamedRangesToSheet(sheetName, table);
    }

    this.getBalances = function (method) {
        const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
            'method': 'get',
            'headers': this.authHeader
        };
        let table = {
            rows: [],
            namedRanges: [],
        };
        const sheetName = method;
        this.accounts.forEach(account => {
            var url = this.apiServer + 'v1/accounts/' + account.number + '/balances';
            table = writeJsonToTable(JSON.parse(UrlFetchApp.fetch(url, options).getContentText())[method], sheetName, table, account);
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
    qt.getBalances('perCurrencyBalances');
    qt.getBalances('combinedBalances');
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

