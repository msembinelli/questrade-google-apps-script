# questrade-google-apps-script
Class to interface with questrade's web API via Google Apps Scripts. I developed this to pull positions/balances into a Google Sheet in order to help balance/manage my portfolio.

## Install
Install clasp globally with npm:
`npm i -g clasp`

Install typescript and apps script ts types:
`npm install`

## Deploy
In `.clasp.json` edit the script ID to contain the ID of your script created within your google sheet (Can be found in File > Project Properties > Info Tab):
```
{
  "scriptId": "<your-script-id-here>",
  "rootDir": "src/",
  "fileExtension": "ts"
}
```

With the same script ID set the redirect address of your Questrade Personal Apps to:
`https://script.google.com/macros/d/{SCRIPT ID}/usercallback`

Login to clasp:
`clasp login` or `sudo clasp login`

Push to remote project:
`clasp push`

Open the project:
`clasp open`

Navigate to script properties (File > Project Properties > Script Properties) and add the following key value pair:
`customerKey <QUESTRADE_PERSONAL_APP_CUSTOMER_KEY>`

Add [apps-script-oauth2](https://github.com/gsuitedevs/apps-script-oauth2) library:
1. Click on the menu item "Resources > Libraries..."
2. In the "Find a Library" text box, enter the script ID
   `1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF` and click the
   "Select" button.
3. Choose a version in the dropdown box (usually best to pick the latest
   version).
4. Click the "Save" button.

A menu named Questrade will be added within a few second when opening the sheet. Select option `Pull` to get data from Questrade.
If no valid credential is stored, a link to authorize the script will appear on the side. Once authorize use `Questrade->Pull` again.
