[README.md](https://github.com/user-attachments/files/30032592/README.md)
# Cygnus Playwright POC

Playwright automation for the Cygnus Power BI dashboard — exports report data for
regression testing, and validates that a report migrated from Power BI **Import
mode** to **Direct Lake** still shows the same data.

---

## 1. First-Time Setup Guide

### One-command installer (recommended)

### One file, double-click (new machine, easiest)

Share **`bootstrap.cmd`** (email it, Teams/Slack it, drop it on a shared
drive — it's a single standalone file, the rest of the repo doesn't need to
be there yet). On the new machine, double-click it. It downloads the
installer and runs it, which:
- clones/updates the repo to `%USERPROFILE%\cygnus-playwright-poc`
- runs `install.ps1` automatically

A console window opens so you can watch progress and answer the credential
prompts; it stays open at the end so you can read the result before closing.
If something fails partway, fix the reported issue and double-click it again
— steps already completed are skipped.

### One link / bootstrap command (alternative — if you'd rather paste a command)

If the machine does not even have the repo yet, share this single command:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "$u='https://raw.githubusercontent.com/chakshugoel1/cygnus-playwright-poc/master/bootstrap-install.ps1'; $p=Join-Path $env:TEMP 'bootstrap-install.ps1'; iwr -UseBasicParsing $u -OutFile $p; powershell -NoProfile -ExecutionPolicy Bypass -File $p"
```

What this does:
- clones/updates the repo to `%USERPROFILE%\\cygnus-playwright-poc`
- runs `install.ps1` automatically

Optional example (skip parity smoke check):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "$u='https://raw.githubusercontent.com/chakshugoel1/cygnus-playwright-poc/master/bootstrap-install.ps1'; $p=Join-Path $env:TEMP 'bootstrap-install.ps1'; iwr -UseBasicParsing $u -OutFile $p; powershell -NoProfile -ExecutionPolicy Bypass -File $p -SkipParity"
```

If your org blocks `Invoke-WebRequest`, copy `bootstrap-install.ps1` (or
`bootstrap.cmd`, which needs no arguments) to the new machine and run it
directly.

> **First-run security prompt:** since both files are downloaded from the
> internet, Windows SmartScreen will show "Windows protected your PC" the
> first time you run either one. Click **More info → Run anyway** — this is
> normal for any unsigned script/file shared outside the Microsoft Store, not
> a sign something is wrong.

For a fresh machine, run from repo root:

```powershell
.\install.ps1
```

Or if script execution is blocked:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1
```

What it does in phase 1:
- checks/installs Node LTS and Git via `winget` (if missing)
- installs root npm dependencies
- installs Playwright Chromium
- creates `%USERPROFILE%\Power_BI_report_validation_credentials`
- creates `%USERPROFILE%\Power_BI_report_validation_credentials\.env` and `pbi-service-principal.json` from templates if missing
- runs `npm run test:setup` once if credentials exist and auth session is missing
- runs `npm run parity` as a smoke check (unless skipped)
- optionally installs `app-desktop` dependencies

Useful installer modes:

```powershell
.\install.ps1 -Mode verify            # check-only, no installs/changes
.\install.ps1 -Mode fast              # skips npm install/browsers/parity
.\install.ps1 -Mode full -SkipParity  # full setup but skip parity smoke check
```

NPM shortcuts:

```powershell
npm run install:setup
npm run install:verify
npm run install:fast
```

---

Follow these steps in order the first time you set this project up on a machine.
Steps 5–6 (credentials) only need to be done once per machine, not once per run.

### Prerequisites

- **Node.js** (LTS, v20+) — [nodejs.org](https://nodejs.org)
- **Git**
- Access to the Cygnus Power BI workspace with a valid login, plus the two
  credential files described in Step 5 (ask your tech lead / IT if you don't
  have these yet)

### Step 1 — Clone the repo

```powershell
git clone https://github.com/chakshugoel1/cygnus-playwright-poc.git
cd cygnus-playwright-poc
```

### Step 2 — Install dependencies

```powershell
npm install
```

### Step 3 — Install the Playwright browser

```powershell
npm run install-browsers
```

This downloads Chromium into Playwright's local cache. It's a one-time step per
machine (not per clone) — if you've already run Playwright on this machine
before, it may already be cached.

### Step 4 — Create the secrets folder

Credentials live **outside** this project folder, at a fixed path, so they're
never accidentally committed or synced to OneDrive/Git:

```
%USERPROFILE%\Power_BI_report_validation_credentials\
```

Create that folder if it doesn't exist yet.

### Step 5 — Add the two credential files

**`%USERPROFILE%\Power_BI_report_validation_credentials\.env`**
```
E2E_USERNAME=your.name@sopra-steria.com
E2E_PASSWORD=your-password
```
This is a dedicated test account used to log in to Cygnus automatically — not
your personal MFA-protected account.

**`%USERPROFILE%\Power_BI_report_validation_credentials\pbi-service-principal.json`**
```json
{
  "tenantId": "<from IT>",
  "clientId": "<from IT>",
  "clientSecret": "<from IT>"
}
```
This is used to call the Power BI REST API directly (report metadata, embed
tokens). Get these three values from whoever administers your Power BI tenant.

> Both files are already excluded by `.gitignore` and will never show up in
> `git status` — that's expected, not a mistake.

### Step 6 — Log in once and verify

```powershell
npm run test:setup
```
This opens a browser, logs in with the credentials from Step 5, and saves the
session so future runs don't need to log in again. **Do this before running
anything else** — `parity` and `main:run` (below) expect this session to
already exist and won't trigger it automatically.

Then run one flow to confirm everything works end to end:
```powershell
npm run parity
```
If it finishes and prints a summary (pass or "differences found" — either is
fine, it means the automation itself worked), setup is complete. If it errors
out, see the **Troubleshooting** note at the end of this file.

---

## 2. POC Explanation

### What this does

This project drives a real Power BI report inside a Playwright-controlled
browser, pulls the actual rendered data out of every visual (via Power BI's
`exportData()` API — not scraped text), and writes it to Excel so it can be
compared. It supports two separate use cases:

| Flow | Command | What it checks |
|---|---|---|
| **Main Run** | `npm run main:run` | Does the Cygnus report still show the same data it showed last time? (regression check against a saved baseline) |
| **Report Parity** | `npm run parity` | Does a report migrated from **Import mode** to **Direct Lake** still show the same data as the original? |

### How it embeds the report

Power BI's embed SDK expects the parent page to be on the `app.powerbi.com`
origin. `harness/harness.html` is intercepted and served locally by Playwright
at a fake `app.powerbi.com` URL (`page.route()`), which makes the embedded
report iframe same-origin — so `postMessage` calls between them work without
tripping Power BI's origin checks. See `tests/helpers/harness.helpers.ts` for
the details.

### Report Parity — how the comparison works

1. The **source** (Import mode) and **target** (Direct Lake) report identities
   are swapped in and out via `applyReportIdentity()` — same report embed
   mechanism, different dataset/report IDs each time.
2. Every visual on every configured page is exported to `expected.xlsx`
   (source) and `actual.xlsx` (target).
3. `excel-compare.helpers.ts` re-groups each sheet back into its individual
   Power BI visuals (not just raw rows) and compares them **visual to
   visual** — so a table that moved to a different position, or a visual
   whose column got renamed, doesn't get treated as a real difference.
4. `parity-summary.xlsx` is generated with four sheets:
   - **Parity Summary** — overall verdict and counts
   - **Page Comparison** — one row per report page
   - **Visual Comparison** — one row per visual, with an exact status
     (identical / header text differs / duplicated / values differ / missing
     / ambiguous)
   - **Differences** — sample rows behind any genuine value differences

The Playwright test itself only fails on a real automation error (login
failed, a report wouldn't embed, a file didn't get written). Finding that the
two reports' *data* differs is a normal, expected result — it shows up as a
console message and an annotation in the HTML report, not a failure. Set
`FAIL_ON_PARITY_DIFF=1` before running if you ever want a data mismatch to
fail the test outright (e.g. wiring this into a CI gate).

Full details on configuring report pairs, pages, and interpreting results:
see **`REPORT-PARITY-README.md`**.

### Desktop GUI (optional, no terminal needed)

`app-desktop/` is an Electron wrapper around the same setup/parity commands,
for anyone who'd rather click a button than use PowerShell. See
**`app-desktop/README.md`** for details.

### Command reference

| Command | Purpose |
|---|---|
| `npm run test:setup` | Log in and save the session (run first, and again if the session expires) |
| `npm run main:run` | Export the current report and compare against the saved baseline |
| `npm run parity` | Export source + target and compare (Report Parity) |
| `npm run parity:source` | Export only the source (refresh the baseline) |
| `npm run parity:target` | Export only the target, compare against the existing baseline |
| `npm run compare:excel` | Standalone Excel-to-Excel comparison (no browser) |
| `npm run show:report` | Open the last Playwright HTML report |

To select a specific report pair (when more than one is configured in
`tests/helpers/comparison-config.helpers.ts`):
```powershell
$env:PAIR='Cygnus'; npm run parity
```

### Where results land

- `playwright-report/` — Main Run HTML report
- `playwright-report-parity/<pair-name>/` — `expected.xlsx`, `actual.xlsx`,
  `parity-summary.xlsx`
- `test-results/` — traces, screenshots, and videos for failed runs

### Troubleshooting

- **`npm run test:setup` fails immediately** — check `.env` exists at
  `%USERPROFILE%\Power_BI_report_validation_credentials\.env` with `E2E_USERNAME` and
  `E2E_PASSWORD` set.
- **A parity run fails with an HTTP 401/403 fetching report metadata** —
  usually the service principal credentials in `pbi-service-principal.json`
  are missing, wrong, or don't have access to the workspace.
- **Chromium not found** — re-run `npm run install-browsers`.
