# Step-by-Step Execution Guide

This guide explains how to run the publication development report in Microsoft
Edge using saved website logins. Website discovery stays within each
publication's own domain; the automation does not search Google or Bing.

## 1. Open PowerShell

Open Windows PowerShell and change to the automation folder:

```powershell
Set-Location "C:\Users\vijay\OneDrive\Documents\AI Search - 3\publication-research-automation"
```

Keep this PowerShell window open until the report finishes.

## 2. Confirm the input spreadsheet is present

Run:

```powershell
Test-Path ".\input\Publications for AI Search.xlsx"
```

PowerShell should display:

```text
True
```

If it displays `False`, copy the current publication spreadsheet into the
`input` folder and name it exactly:

```text
Publications for AI Search.xlsx
```

The automation reads publication name, website, region, county, and city. It
does not read or send the spreadsheet's username or password columns.

## 3. Set the OpenAI API key

The API key must be set in every new PowerShell session used for a live run:

```powershell
$secureKey = Read-Host "OpenAI API key" -AsSecureString
$env:OPENAI_API_KEY = [System.Net.NetworkCredential]::new("", $secureKey).Password
Remove-Variable secureKey
```

Confirm that the variable exists without displaying the secret:

```powershell
if ($env:OPENAI_API_KEY) { "API key is set" } else { "API key is missing" }
```

Never enter the API key in the spreadsheet, source code, Task Scheduler
arguments, screenshots, or email.

## 4. Close previous automation browser windows

Before starting a browser run, close every Microsoft Edge window previously
opened by this automation.

Only one automation run may use `browser-profile` at a time. Regular Edge
windows using your normal Edge profile may remain open.

## 5. Save a website login for the first time

Set up one publication at a time. Example:

```powershell
powershell -ExecutionPolicy Bypass -File .\run_report.ps1 `
  -LoginSetup `
  -Site napavalleyregister.com
```

When Edge opens:

1. Allow the publication homepage to finish loading.
2. Navigate to the publication's login page if needed.
3. Sign in using your authorized subscription.
4. Complete any consent, multifactor, or human-verification step.
5. Open a subscriber article and confirm its text is readable.
6. Return to PowerShell and press Enter.

The browser closes after setup. Cookies and local storage are retained in:

```text
browser-profile
```

Repeat this step for each login-required publication. For example:

```powershell
powershell -ExecutionPolicy Bypass -File .\run_report.ps1 `
  -LoginSetup `
  -Site theregistrysf.com
```

Login sessions normally survive closing PowerShell, closing Edge, and restarting
the computer. Repeat login setup only when a site expires or invalidates its
session.

## 6. Test one publication without using AI tokens

Use a dry run to confirm browser discovery and extraction:

```powershell
powershell -ExecutionPolicy Bypass -File .\run_report.ps1 `
  -Period weekly `
  -Browser `
  -Site napavalleyregister.com `
  -DryRun
```

During internal publication searches, PowerShell displays:

```text
Waiting for internal search results: development
Waiting for internal search results: housing apartments
Waiting for internal search results: retail restaurant
```

The automation waits for background network activity, JavaScript results, and
lazy-loaded links before continuing.

`-DryRun` performs browser extraction and creates a test workbook, but it skips
AI analysis and PDF downloading.

## 7. Run one publication with AI analysis

After the dry run succeeds:

```powershell
powershell -ExecutionPolicy Bypass -File .\run_report.ps1 `
  -Period weekly `
  -Browser `
  -Site napavalleyregister.com
```

Do not close Edge or PowerShell while the publication is being reviewed.

Normal report execution never pauses for login, Continue, CAPTCHA, or
verification. Affected pages are recorded in `Source Audit` and skipped
automatically. Only the separate `-LoginSetup` command asks you to press Enter.

The automation does not bypass paywalls, CAPTCHA, multifactor authentication, or
website access controls.

## 8. Run all publications weekly

Use this command for the current day and previous six days:

```powershell
powershell -ExecutionPolicy Bypass -File .\run_report.ps1 `
  -Period weekly `
  -Browser
```

The run can take considerable time when the spreadsheet contains many
publications. Each website is searched only within its own domain.

## 9. Run all publications monthly

Use this command for the rolling 30-day period ending today:

```powershell
powershell -ExecutionPolicy Bypass -File .\run_report.ps1 `
  -Period monthly `
  -Browser
```

## 10. Run an exact date range

Dates are inclusive and must use `YYYY-MM-DD`:

```powershell
powershell -ExecutionPolicy Bypass -File .\run_report.ps1 `
  -StartDate 2026-07-01 `
  -EndDate 2026-07-31 `
  -Browser
```

## 11. Process a small batch

To test only the first three distinct publication domains:

```powershell
powershell -ExecutionPolicy Bypass -File .\run_report.ps1 `
  -Period weekly `
  -Browser `
  -MaxSites 3
```

Use `-Site` when you need a specific publication. Use `-MaxSites` when you need
the first few publications from the spreadsheet.

## 12. Monitor the run

Normal PowerShell progress resembles:

```text
[1/60] Browser reviewing Publication Name
    Opening publication homepage
    Discovering candidate articles
    Sitemap discovery found ...
    Waiting for internal search results: development
    Publication search found ...
    Homepage discovery found ...
    Browser article 1/10 ...
    Extracted ... candidate article(s)
    ... qualifying item(s)
```

Do not start another run while this process is active.

To stop a run safely, click the PowerShell window and press:

```text
Ctrl+C
```

Then close the automation Edge window before restarting.

## 13. Find the completed report

Every run creates a new timestamped folder under:

```text
runs
```

Example:

```text
runs\
  2026-07-24_to_2026-07-30_2026-07-30T...\
    Development Research 2026-07-24 to 2026-07-30.xlsx
    run-metadata.json
    pdf\
    preview\
```

Open the Excel workbook and review:

- `Summary`: report period, totals, and scope.
- `Developments`: qualifying projects, addresses, cross streets, units,
  descriptions, and article URLs.
- `Source Audit`: every processed website, including blocked and no-result sites.
- `PDF Index`: downloaded or failed building-plan PDF links.

Downloaded plan files are stored in the `pdf` folder and named:

```text
City - Address - YYYY-MM-DD.pdf
```

## 14. Run unattended after logins are established

Only use unattended mode after visible browser runs work reliably:

```powershell
powershell -ExecutionPolicy Bypass -File .\run_report.ps1 `
  -Period weekly `
  -Browser `
  -Headless `
  -NonInteractive
```

An unattended run cannot complete an expired login, CAPTCHA, multifactor prompt,
or new consent screen. It records and skips those sites, then continues through
the complete publication list. Review `Source Audit` afterward and refresh
affected sessions with `-LoginSetup`.

## 15. Schedule a weekly run

In Windows Task Scheduler:

1. Select **Create Task**.
2. Choose the Windows account that owns `browser-profile`.
3. Create a weekly trigger.
4. Add an action with:

   **Program/script**

   ```text
   powershell.exe
   ```

   **Arguments**

   ```text
   -ExecutionPolicy Bypass -File "C:\Users\vijay\OneDrive\Documents\AI Search - 3\publication-research-automation\run_report.ps1" -Period weekly -Browser -Headless -NonInteractive
   ```

   **Start in**

   ```text
   C:\Users\vijay\OneDrive\Documents\AI Search - 3\publication-research-automation
   ```

5. Ensure the scheduled account can obtain `OPENAI_API_KEY` from an approved
   secure user-level secret mechanism.
6. Do not place the API key directly in the Task Scheduler arguments.
7. Test the task manually and review the generated `Source Audit`.

## 16. Common problems

### `OPENAI_API_KEY is not set`

Repeat Step 3 in the current PowerShell window.

### Edge opens but immediately closes

Close every automation Edge window and retry. Another run may be locking
`browser-profile`.

### The site is not logged in

Repeat Step 5 for that domain. Press Enter only after a subscriber article is
fully readable.

### `Challenge failed: Bot detected`

Stop the run. Do not repeatedly retry or attempt to bypass the challenge. Confirm
that old automation Edge windows are closed and repeat login setup once. If the
site continues rejecting automated browser access, review it manually.

### Internal search results are still loading

The default minimum wait is 5 seconds, with up to 12 additional seconds for
relevant links. These values can be increased in `config.json`:

```json
"searchMinimumWaitMilliseconds": 5000,
"searchResultsWaitMilliseconds": 12000
```

### A website has no qualifying result

Check `Source Audit`. The website may have had no matching article in the date
window, no usable internal search results, expired login, or blocked access.

### A PDF failed to download

Open `PDF Index`, follow the recorded PDF URL, and download it manually using
your authorized browser session.
