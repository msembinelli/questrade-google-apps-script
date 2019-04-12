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
Login to clasp:
`clasp login` or `sudo clasp login`

Push to remote project:
`clasp push`

Open the project:
`clasp open`

Navigate to script properties (File > Project Properties > Script Properties) and add the following key value pair:
tokenUrl https://login.questrade.com/oauth2/token

Now attach the `run()` command to a button in google sheets or set the `run()` function to run when the sheet opens.

Upon the first run, you will get a prompt asking for your questrade refresh key. You will have to retreive this from your questrade account via the app hub.
