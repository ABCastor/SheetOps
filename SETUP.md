# Setup

SheetOps needs a Google Cloud Platform project with three APIs enabled and a set of OAuth 2.0 Desktop app credentials. This is a one-time setup.

## Step 1: Create a GCP project

Go to [console.cloud.google.com](https://console.cloud.google.com/), click the project selector, then New Project. Name it anything (`sheetops` works fine). Note the Project Number from the dashboard; you will need it later when linking Apps Script projects.

## Step 2: Enable APIs

With your project selected, enable these APIs:

- [Google Sheets API](https://console.cloud.google.com/apis/library/sheets.googleapis.com)
- [Google Drive API](https://console.cloud.google.com/apis/library/drive.googleapis.com)
- [Apps Script API](https://console.cloud.google.com/apis/library/script.googleapis.com)

Click each link, select your project, click Enable.

## Step 3: Configure the OAuth consent screen

Go to APIs & Services > OAuth consent screen. Choose External (or Internal for Google Workspace). Fill in the app name and your email address, then add these scopes on the Scopes screen:

```
https://www.googleapis.com/auth/spreadsheets
https://www.googleapis.com/auth/drive
https://www.googleapis.com/auth/script.projects
https://www.googleapis.com/auth/userinfo.email
```

Add your own email as a Test user and save.

## Step 4: Create OAuth credentials

Go to APIs & Services > Credentials > Create Credentials > OAuth client ID. Set the application type to Desktop app. Click Create, then copy the Client ID and Client Secret.

## Step 5: Authenticate SheetOps

```bash
node bin/sheetops.js auth \
  --client-id YOUR_CLIENT_ID \
  --client-secret YOUR_CLIENT_SECRET
```

A browser window opens. Sign in with the account you added as a Test user. Credentials go to `~/.sheetops-creds.json` and are never committed to git.

## Step 6: Link Apps Script to GCP (for `run-script`)

The `run-script` command uses the Apps Script Execution API. Each Apps Script project must be linked to your GCP project once. Run:

```bash
node bin/sheetops.js setup-gcp --project YOUR_PROJECT_NAME
```

Follow the printed steps. This is per Apps Script project, not per account.

## Verify

```bash
node bin/sheetops.js init
```

The output shows which tools are available, the active Google account, and all registered projects.
