# Cygnus Playwright POC

Playwright automation for the Cygnus Power BI dashboard — exports report data for
regression testing, and validates that a report migrated from Power BI **Import
mode** to **Direct Lake** still shows the same data.

---

## 1. First-Time Setup Guide

### One file, double-click (new machine, easiest)

Share **`bootstrap.cmd`** (email it, Teams/Slack it, drop it on a shared
drive — it's a single standalone file, the rest of the repo doesn't need to
be there yet). On the new machine, double-click it.

A console window opens and explains what it's about to do, then:
- clones/updates the repo to `%USERPROFILE%\cygnus-playwright-poc`
- installs Node/Git via `winget` if missing
- runs `npm install` for both the root project and the desktop app, and
  installs the Playwright browser
- prompts once for your `E2E_USERNAME` / `E2E_PASSWORD` (a dedicated test
  account, not your personal MFA-protected login) and saves them to
  `%USERPROFILE%\Power_BI_report_validation_credentials\.env`
- creates a desktop shortcut to the app

It does **not** log in or run a comparison automatically — those open a real
browser / take several minutes, so they're left for you to trigger when
you're ready (see "Desktop GUI" below). The window stays open at the end and
tells you exactly which shortcut to double-click next.

If something fails partway, fix the reported issue and double-click
`bootstrap.cmd` again — steps already completed are skipped.

### One link / bootstrap command (alternative — if you'd rather paste a command)

If the machine does not even have the repo yet, share this single command:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "$u='https://raw.githubusercontent.com/chakshugoel1/cygnus-playwright-poc/master/bootstrap-install.ps1'; $p=Join-Path $env:TEMP 'bootstrap-install.ps1'; iwr -UseBasicParsing $u -OutFile $p; powershell -NoProfile -ExecutionPolicy Bypass -File $p"
```

If your org blocks `Invoke-WebRequest`, copy `bootstrap-install.ps1` (or
`bootstrap.cmd`, which needs no arguments) to the new machine and run it
directly.

> **First-run security prompt:** since both files are downloaded from the
> internet, Windows SmartScreen will show "Windows protected your PC" the
> first time you run either one. Click **More info → Run anyway** — this is
> normal for any unsigned script/file shared outside the Microsoft Store, not
> a sign something is wrong.

### Running the installer directly (already have the repo)

```powershell
.\install.ps1
```

Or if script execution is blocked:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1
```

What it does, in order:
1. **Prerequisites** — Node/Git via `winget` if missing, `npm install` for
   both the root project and `app-desktop` (always, not optional — the
   desktop shortcut depends on it), Playwright browser install.
2. **Secrets** — prompts for `E2E_USERNAME`/`E2E_PASSWORD`, saves to
   `%USERPROFILE%\Power_BI_report_validation_credentials\.env`. Also creates
   a `pbi-service-principal.json` template (optional — only needed for the
   DAX-based `main:run` flow; the parity flow works without it).
3. **Auth session** — reports whether a saved login session already exists.
   Does **not** open a browser to log in automatically (pass `-RunAuthSetup`
   to opt into that).
4. **Parity smoke check** — reports readiness. Does **not** run automatically
   (pass `-RunParity` to opt into that, or use the desktop app).
5. **Desktop shortcut** — creates a `.lnk` shortcut to `app-desktop\launch-desktop.bat`.
   Asks where to put it (defaults to your Desktop).

Useful installer flags:

```powershell
.\install.ps1 -Mode verify                # check-only, no installs/changes
.\install.ps1 -Mode fast                  # skips npm install/browsers
.\install.ps1 -Mode full -NonInteractive  # no prompts (unattended)
.\install.ps1 -RunAuthSetup -RunParity    # opt into full automation (no desktop app needed)
```

NPM shortcuts (run these from an already-cloned repo):

```powershell
npm run install:setup
npm run install:verify
npm run install:fast
```

---

### Manual setup (if you're not using the installer)

Follow these steps in order the first time you set this project up by hand.
Steps 4–5 (credentials) only need to be done once per machine, not once per run.

#### Prerequisites

- **Node.js** (LTS, v20+) — [nodejs.org](https://nodejs.org)
- **Git**
- Access to the Cygnus Power BI workspace with a valid login, plus the
  credential file described in Step 5 (ask your tech lead / IT if you don't
  have these yet)

#### Step 1 — Clone the repo

```powershell
git clone https://github.com/chakshugoel1/cygnus-playwright-poc.git
cd cygnus-playwright-poc
```

#### Step 2 — Install dependencies

```powershell
npm install
```

#### Step 3 — Install the Playwright browser

```powershell
npm run install-browsers
```

This downloads Chromium into Playwright's local cache. It's a one-time step per
machine (not per clone) — if you've already run Playwright on this machine
before, it may already be cached.

#### Step 4 — Add your credentials

Credentials live **outside** this project folder, at a fixed path, so they're
never accidentally committed or synced to OneDrive/Git:

**`%USERPROFILE%\Power_BI_report_validation_credentials\.env`**
```
E2E_USERNAME=your.name@sopra-steria.com
E2E_PASSWORD=your-password
```
This is a dedicated test account used to log in to Cygnus automatically — not
your personal MFA-protected account. Create the folder and file yourself if
you're skipping the installer.

**`%USERPROFILE%\Power_BI_report_validation_credentials\pbi-service-principal.json`**
(optional — only needed for the DAX-based `main:run` flow; the parity flow
works without it, falling back to your signed-in user token)
```json
{
  "tenantId": "<from IT>",
  "clientId": "<from IT>",
  "clientSecret": "<from IT>"
}
```

> Both files live outside the repo and are also covered by `.gitignore` —
> they will never show up in `git status`, that's expected.

**Credentials are never entered anywhere in the desktop app's UI** — the
`.env` file above is the single, exclusive source for them.

#### Step 5 — Log in once and verify

```powershell
npm run test:setup
```
This opens a browser, logs in with the credentials from Step 4, and saves the
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

### Desktop GUI

`app-desktop/` is an Electron app for running everything without a terminal —
this is the primary way to use the project day to day, not just an optional
extra. Open it via the shortcut the installer created, or manually:

```powershell
cd app-desktop
npm install
npm start
```

What it does:
- **Report identities** — enter Source/Target Tenant ID / Group ID / Report ID
  / Dataset ID (no code file to edit). Nothing is hardcoded here; this
  overrides `tests/helpers/comparison-config.helpers.ts` for the run.
- **Discover Filters** — reads report-level filters instantly, and reads
  page-level slicers too. Leave the Pages field blank and it scans pages in
  order, crawling only the first one that actually has slicers — usually
  enough, since most reports repeat the same fields across pages, and this
  avoids a second full-report embed. Type specific page names to crawl those
  instead.
- **Run Authentication / Run Parity / Run Auth + Parity** — Run Parity checks
  for a saved login session first and transparently falls back to Auth +
  Parity if one doesn't exist yet, so a cold start never fails partway
  through with a confusing token error.
- **Credentials are never entered in this UI** — they live only in
  `%USERPROFILE%\Power_BI_report_validation_credentials\.env`, set via the
  installer's prompt or by editing that file directly.

See **`app-desktop/README.md`** for more detail.

### Command reference

| Command | Purpose |
|---|---|
| `npm run test:setup` | Log in and save the session (run first, and again if the session expires) |
| `npm run main:run` | Export the current report and compare against the saved baseline |
| `npm run parity` | Export source + target and compare (Report Parity) |
| `npm run parity:source` | Export only the source (refresh the baseline) |
| `npm run parity:target` | Export only the target, compare against the existing baseline |
| `npm run compare:excel` | Standalone Excel-to-Excel comparison (no browser) |
| `npm run discover:slicers` | Discover report-level filters and (optionally) page slicers — see the header comment in `tests/specs/discover-slicers.spec.ts` for all the `DISCOVER_*` env vars |
| `npm run show:report` | Open the last Playwright HTML report |
| `npm run test:unit` | Run the browser-free unit tests (filter planning, scenario validation) |
| `npm run typecheck` | `tsc --noEmit` across the whole project |

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

All of the above are local-only output folders (git-ignored) — safe to
delete anytime; they're recreated on the next run.

### Troubleshooting

- **`npm run test:setup` fails immediately** — check `.env` exists at
  `%USERPROFILE%\Power_BI_report_validation_credentials\.env` with
  `E2E_USERNAME` and `E2E_PASSWORD` set.
- **Desktop app: "Run Parity" ran Authentication first, unexpectedly** —
  that's intentional: it means no saved login session was found, so it
  automatically ran Auth + Parity instead of failing partway through. Not an
  error.
- **A parity run fails with an HTTP 401/403 fetching report metadata** —
  usually the service principal credentials in `pbi-service-principal.json`
  are missing, wrong, or don't have access to the workspace. This file is
  optional for the parity flow, so a missing/incomplete one isn't itself the
  problem — check this only if the error specifically mentions the service
  principal.
- **Chromium not found** — re-run `npm run install-browsers`.
