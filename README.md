# Technicians Scheduled Work Report

Node.js 20 project that creates a new Excel workbook in SharePoint for Zoho FSM scheduled Service Appointments. It fetches Service Appointments first, filters by today's business date, and only fetches Work Orders when an appointment is linked to one.

The report intentionally does not use Time Sheets, actual start/end times, or actual duration. It uses Scheduled Start, Scheduled End, and Scheduled Duration only.

## Local Setup

```bash
npm install
cp .env.example .env
```

Fill in `.env` with Zoho OAuth values and Microsoft Graph app-only credentials.

## Zoho OAuth Setup

Start the local callback helper:

```bash
node src/oauthCallback.js
```

Open this authorization URL after replacing placeholders:

```text
{ZOHO_ACCOUNTS_URL}/oauth/v2/auth?scope=ZohoFSM.modules.ServiceAppointments.READ,ZohoFSM.modules.WorkOrders.READ,ZohoFSM.users.READ&client_id={ZOHO_CLIENT_ID}&response_type=code&access_type=offline&redirect_uri=http://localhost:3000/oauth/callback&prompt=consent
```

The callback page and terminal will show the authorization code.

Exchange the authorization code for tokens:

```text
POST {ZOHO_ACCOUNTS_URL}/oauth/v2/token
grant_type=authorization_code
client_id=ZOHO_CLIENT_ID
client_secret=ZOHO_CLIENT_SECRET
redirect_uri=http://localhost:3000/oauth/callback
code=AUTHORIZATION_CODE
```

Store the returned refresh token as `ZOHO_REFRESH_TOKEN`.

Daily runs use the refresh-token flow:

```text
POST {ZOHO_ACCOUNTS_URL}/oauth/v2/token
grant_type=refresh_token
client_id=ZOHO_CLIENT_ID
client_secret=ZOHO_CLIENT_SECRET
refresh_token=ZOHO_REFRESH_TOKEN
```

Required Zoho scopes:

```text
ZohoFSM.modules.ServiceAppointments.READ
ZohoFSM.modules.WorkOrders.READ
ZohoFSM.users.READ
```

## Microsoft Setup

Use an Azure App Registration with Microsoft Graph application permissions and admin consent:

```text
Files.ReadWrite.All
Sites.ReadWrite.All
```

Set these values in `.env`:

```env
MICROSOFT_TENANT_ID=
MICROSOFT_CLIENT_ID=
MICROSOFT_CLIENT_SECRET=
MICROSOFT_SITE_ID=
MICROSOFT_DRIVE_ID=
MICROSOFT_FOLDER_PATH=/FSM Scheduled Reports
```

`MICROSOFT_SITE_ID` is kept for reference. The upload uses `MICROSOFT_DRIVE_ID` and `MICROSOFT_FOLDER_PATH`.

## Run Locally

```bash
node src/index.js
```

On success, the SharePoint Excel workbook URL is printed.

## Railway Deployment

1. Create a Railway service from this repository.
2. Set the environment variables from `.env.example`.
3. Use the Dockerfile or Railway's Node builder.
4. Run `node src/index.js` as the scheduled command.

The script exits with code `0` on success and non-zero on failure.

## Environment Variables

```text
ZOHO_ACCOUNTS_URL
ZOHO_FSM_API_DOMAIN
ZOHO_CLIENT_ID
ZOHO_CLIENT_SECRET
ZOHO_REFRESH_TOKEN
BUSINESS_TIMEZONE=Asia/Dubai
MICROSOFT_TENANT_ID
MICROSOFT_CLIENT_ID
MICROSOFT_CLIENT_SECRET
MICROSOFT_SITE_ID
MICROSOFT_DRIVE_ID
MICROSOFT_FOLDER_PATH=/FSM Scheduled Reports
REPORT_TITLE_PREFIX=Technicians Scheduled Work
COMPANY_NAME=TPH Group
REPORT_DATE_OVERRIDE optional, YYYY-MM-DD for local backtesting only
ALLOW_REPORT_DATE_OVERRIDE=false
```

For Railway production, leave `REPORT_DATE_OVERRIDE` empty and keep `ALLOW_REPORT_DATE_OVERRIDE=false`. The script will generate the report for the current day in `BUSINESS_TIMEZONE`.

## Railway Cron

Railway cron schedules are UTC. To run every day at 7:00 AM UAE time, use:

```text
0 3 * * *
```

This is already set in `railway.json`.

## Output

Every run creates a new Excel workbook named:

```text
Technicians Scheduled Work - YYYY-MM-DD.xlsx
```

If a file with that name already exists, SharePoint is asked to rename the new upload instead of replacing the existing workbook.

The first tab is `Scheduled Work` and includes title rows, grouped technician rows, subtotal rows per Service Resource, and a Grand Total row. If no appointments are found, the workbook is still created with the title/header rows and:

```text
No scheduled appointments found for YYYY-MM-DD.
```
