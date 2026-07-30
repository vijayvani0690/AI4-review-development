# Guide for Any AI Coding Assistant

This repository is a standalone Node.js application. Codex is not required to
install, run, troubleshoot, or modify it. This guide can be given to ChatGPT,
Claude, Gemini, GitHub Copilot, or another coding assistant as project context.

## Copy-and-paste starting prompt

Use this prompt when beginning a new conversation with an AI assistant:

```text
Please read README.md, STEP_BY_STEP_GUIDE.md, and AI_ASSISTANT_GUIDE.md before
making changes. This is a Windows Node.js 20+ application that researches
publication websites using Microsoft Edge and creates an Excel development
report. Preserve saved browser profiles, never commit credentials or
node_modules, keep discovery within each publication website in browser mode,
and make report runs non-interactive so a failed website is audited and skipped
instead of stopping the full run. Before handing back changes, run the syntax
checks and the three-site dry run listed in AI_ASSISTANT_GUIDE.md.
```

Then describe the specific change or problem.

## Application purpose

The application reads publication websites from
`input/Publications for AI Search.xlsx`, researches a selected weekly, monthly,
or explicit date window, and reports:

- demolitions;
- new retail and restaurant buildings;
- retail-center developments;
- residential developments;
- qualifying mixed-use developments;
- supported street addresses and cross streets;
- linked building-plan PDFs.

The result is a timestamped Excel workbook with `Summary`, `Developments`,
`Source Audit`, and `PDF Index` worksheets. Downloaded PDFs use:

```text
City - Address - YYYY-MM-DD.pdf
```

## Technical architecture

| File | Responsibility |
| --- | --- |
| `run_report.ps1` | User-facing PowerShell launcher and argument forwarding |
| `run.mjs` | Date windows, AI requests, classification, PDF downloads, audit records, and orchestration |
| `browser_research.mjs` | Persistent Edge profile, internal-site discovery, rendered-page extraction, timeouts, and skip behavior |
| `standalone_workbook.mjs` | ExcelJS input reading and formatted report generation |
| `config.json` | Model, timeouts, concurrency, browser limits, and research scope |
| `setup.ps1` | Installs locked local Node.js dependencies |
| `package.json` / `package-lock.json` | Reproducible standalone dependencies |

Runtime data is intentionally excluded from Git:

- `node_modules/`
- `browser-profile/`
- `runs/`
- `input/*.xlsx`
- `.env`
- `*.log`

## Required behavior to preserve

1. Normal report runs must never wait for keyboard input.
2. A blocked, broken, slow, or inaccessible website must be recorded in
   `Source Audit`, skipped, and followed by the next website.
3. Only `-LoginSetup` may ask the user to sign in or complete verification.
4. Browser-mode discovery must remain within the current publication's domain.
   Do not use Google or Bing.
5. Do not bypass CAPTCHA, paywalls, access controls, robots enforcement, or
   multifactor authentication.
6. Reuse the dedicated `browser-profile` so authorized sessions can persist on
   the same computer.
7. Never read, transmit, print, or commit workbook usernames and passwords.
8. Do not invent addresses, cross streets, project facts, or PDF links. Leave
   unsupported values blank.
9. Preserve the PDF naming format and ISO date format.
10. Never commit API keys, browser cookies, input workbooks, generated reports,
    or `node_modules`.

## Setup on a new Windows computer

Requirements:

- Node.js 20 or newer;
- Microsoft Edge;
- PowerShell;
- an OpenAI API key for live analysis.

Run:

```powershell
git clone https://github.com/vijayvani0690/AI4-review-development.git
cd .\AI4-review-development\publication-research-automation
powershell -ExecutionPolicy Bypass -File .\setup.ps1
```

Copy the workbook to:

```text
input/Publications for AI Search.xlsx
```

Set the API key only in the current PowerShell session:

```powershell
$env:OPENAI_API_KEY = [System.Net.NetworkCredential]::new(
  "",
  (Read-Host "OpenAI API key" -AsSecureString)
).Password
```

On a new computer, run login setup once because browser profiles and cookies are
not stored in Git:

```powershell
powershell -ExecutionPolicy Bypass -File .\run_report.ps1 -LoginSetup
```

## Common commands

```powershell
# Fast report-generation test without AI calls
powershell -ExecutionPolicy Bypass -File .\run_report.ps1 -Period weekly -DryRun -MaxSites 3

# Browser extraction test without AI calls
powershell -ExecutionPolicy Bypass -File .\run_report.ps1 -Period weekly -Browser -DryRun -MaxSites 1

# Normal weekly browser run
powershell -ExecutionPolicy Bypass -File .\run_report.ps1 -Period weekly -Browser

# Normal monthly browser run
powershell -ExecutionPolicy Bypass -File .\run_report.ps1 -Period monthly -Browser

# One publication
powershell -ExecutionPolicy Bypass -File .\run_report.ps1 -Period weekly -Browser -Site "napavalleyregister.com"

# Explicit inclusive date range
powershell -ExecutionPolicy Bypass -File .\run_report.ps1 -StartDate 2026-07-01 -EndDate 2026-07-31 -Browser
```

## Validation required after code changes

Run these from `publication-research-automation`:

```powershell
node --check .\run.mjs
node --check .\browser_research.mjs
node --check .\standalone_workbook.mjs
npm ls --depth=0
powershell -ExecutionPolicy Bypass -File .\run_report.ps1 -Period weekly -DryRun -MaxSites 3
```

For browser changes, also run the one-site browser dry run. Confirm that it
finishes without asking for input and that failures appear in `Source Audit`.

For workbook changes, open the generated `.xlsx` and visually inspect all four
worksheets. Confirm that dates display as `YYYY-MM-DD`, filters work, hyperlinks
open, headers are readable, and no formula or rendering errors appear.

## Using a different AI coding assistant

No application change is needed. Give the assistant this repository and the
starting prompt above. The assistant only needs normal file and terminal access.
It does not need Codex-specific tools.

## Using a different AI model provider at runtime

The current live-analysis implementation calls the OpenAI Responses API in
`run.mjs` and reads `OPENAI_API_KEY`. That is separate from which coding
assistant maintains the repository.

To add another runtime provider, introduce a provider adapter instead of mixing
provider-specific code into browser extraction or workbook generation. The
adapter must return the same normalized development records currently consumed
by `run.mjs`, including:

- publication and article metadata;
- development type, status, and description;
- city, state, county, address, and cross street;
- unit and square-foot values;
- demolition details and developer;
- plan PDF URLs;
- confidence and evidence notes.

Preserve the existing JSON validation, source URLs, retry/timeout behavior,
auditing, and skip-on-error behavior. Add a configuration field for the provider
and use a provider-specific environment variable. Never silently send data to a
different provider.

## Safe Git handoff

Before committing:

```powershell
git status --short
git diff --check
```

Stage only source code and documentation. Confirm that `browser-profile`,
`input`, `runs`, `node_modules`, credentials, and unrelated ZIP/output files are
not staged.
