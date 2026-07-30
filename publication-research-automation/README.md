# Publication Development Research Automation

This folder contains a repeatable weekly/monthly pipeline for the publication list in `input/Publications for AI Search.xlsx`.

For complete operating instructions, see
[`STEP_BY_STEP_GUIDE.md`](STEP_BY_STEP_GUIDE.md).

The script:

- reads publication name, URL, region, county, and city from the workbook;
- deliberately ignores and never sends workbook usernames or passwords;
- searches each distinct publication domain for the selected date window;
- keeps demolitions, ground-up retail or restaurant buildings, retail-center developments, residential developments, and qualifying mixed-use developments;
- extracts source-supported addresses, cross streets, unit counts, square footage, developers, project status, and source URLs;
- downloads direct plan PDFs when found and labels them `City - Address - YYYY-MM-DD.pdf`;
- creates a formatted Excel workbook with `Summary`, `Developments`, `Source Audit`, and `PDF Index` sheets;
- creates one timestamped run folder, so previous reports are not overwritten.

## One-time setup

1. Put the current publication workbook at:

   `input/Publications for AI Search.xlsx`

2. Set your OpenAI API key in the PowerShell session that will run the report:

   ```powershell
   $env:OPENAI_API_KEY = [System.Net.NetworkCredential]::new(
     "",
     (Read-Host "OpenAI API key" -AsSecureString)
   ).Password
   ```

   Do not put the key in the workbook, source code, or a committed file.

3. Review `config.json`. The default model is `gpt-5.6-terra`, which balances research quality and cost. You can change it to `gpt-5.6-sol` for maximum capability or `gpt-5.6-luna` for lower-cost, high-volume runs.

## Run weekly

```powershell
powershell -ExecutionPolicy Bypass -File .\run_report.ps1 -Period weekly
```

The weekly window is today plus the previous six days.

## Run monthly

```powershell
powershell -ExecutionPolicy Bypass -File .\run_report.ps1 -Period monthly
```

The monthly window is the rolling 30-day period ending today.

## Browser mode for blocked or login-required publications

Browser mode uses a visible Microsoft Edge window and a dedicated persistent
profile in `browser-profile/`. It searches for candidate articles, opens the
rendered pages with your saved cookies, extracts the visible article evidence and
PDF links, and then sends that evidence to the AI for classification.

The first time, open each subscribed site and sign in:

```powershell
powershell -ExecutionPolicy Bypass -File .\run_report.ps1 -LoginSetup
```

For each website, finish any authorized login or verification in Edge and press
Enter in PowerShell. You can use `-MaxSites 5` to set up a smaller initial batch.
The profile is excluded from Git and the workbook password columns are never read
or sent to the browser or AI.

Run a visible weekly browser review:

```powershell
powershell -ExecutionPolicy Bypass -File .\run_report.ps1 `
  -Period weekly `
  -Browser
```

Normal report runs never pause for user action. If a page requests login,
Continue, CAPTCHA, multifactor verification, or another access step, the
automation records the issue in `Source Audit`, skips it, and continues. Use
`-LoginSetup` separately to refresh a website session.

After the sessions have been established, an unattended browser run can use:

```powershell
powershell -ExecutionPolicy Bypass -File .\run_report.ps1 `
  -Period weekly `
  -Browser `
  -Headless `
  -NonInteractive
```

An unattended run cannot solve an expired login, CAPTCHA, multifactor prompt, or
new consent screen. Those sites are recorded as blocked and skipped so the
remaining websites can finish.

To test browser extraction on one site without making an AI request:

```powershell
powershell -ExecutionPolicy Bypass -File .\run_report.ps1 `
  -Period weekly `
  -Browser `
  -Headless `
  -NonInteractive `
  -MaxSites 1 `
  -DryRun
```

To run one publication by domain or name:

```powershell
powershell -ExecutionPolicy Bypass -File .\run_report.ps1 `
  -Period weekly `
  -Browser `
  -Site napavalleyregister.com
```

Discovery remains entirely within the publication website. It uses publication
sitemaps, the publication's internal search page, its homepage, and article
pages. The automation does not open Google or Bing.

## Run an exact date range

```powershell
powershell -ExecutionPolicy Bypass -File .\run_report.ps1 `
  -StartDate 2026-07-01 `
  -EndDate 2026-07-31
```

Dates are inclusive and must use `YYYY-MM-DD`.

## Test without API calls

```powershell
powershell -ExecutionPolicy Bypass -File .\run_report.ps1 -Period weekly -DryRun
```

This verifies workbook reading, output generation, formatting, and previews. It does not search websites or download PDFs.

To test a small live sample:

```powershell
powershell -ExecutionPolicy Bypass -File .\run_report.ps1 -Period weekly -MaxSites 3
```

## Output

Each run is saved under:

```text
runs/
  2026-07-22_to_2026-07-28_2026-07-28T.../
    Development Research 2026-07-22 to 2026-07-28.xlsx
    run-metadata.json
    pdf/
    preview/
```

`Source Audit` always includes every distinct website processed, including blocked sites, inactive/archived sites, duplicates, errors, and sites with no qualifying result.

## Scheduling

Windows Task Scheduler can run:

```text
powershell.exe
```

with arguments:

```text
-ExecutionPolicy Bypass -File "C:\Users\vijay\OneDrive\Documents\AI Search - 3\publication-research-automation\run_report.ps1" -Period weekly
```

For unattended scheduling, store `OPENAI_API_KEY` in a secure user-level secret mechanism available to the scheduled account. Do not place the key directly in Task Scheduler arguments.

## Important limitations

- Paywalled or robots-blocked publications may return limited results; the audit sheet records this.
- Browser mode uses only content your normal browser session is authorized to
  display. It does not bypass paywalls, robots enforcement, CAPTCHA, multifactor
  authentication, or other access controls.
- The workflow relies on public web search. It does not sign in to publication accounts.
- Cross streets and plan links remain blank when no reliable source provides them.
- Some sites block automated PDF downloads even when the PDF is visible in search. Those failures remain in `PDF Index` with the source URL so they can be downloaded manually.
- Web research and tool calls incur OpenAI API usage charges. Reduce `concurrency`, use a shorter date window, or select `gpt-5.6-luna` if cost is a concern.
