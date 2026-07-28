# How to use the Power BI Report Validator

A first-time walkthrough of the desktop app. The order is always the same:
**Run Authentication → Discover Filters → Run Parity.**

## Before you start

- You have the **report IDs** for both reports you're comparing (the Import-mode
  source and the Direct-Lake target): Tenant, Group/Workspace, Report, Dataset.
- A **test account** is configured in the `.env` file inside
  `Power_BI_report_validation_credentials` (set once by the installer, or edited
  directly). Credentials are never typed into the app itself.

## The flow

```text
┌──────────────────────────────────────────────────────────────────┐
│          START — open the Power BI Report Validator app          │
└──────────────────────────────────────────────────────────────────┘
                                  │
                                  ▼
┌──────────────────────────────────────────────────────────────────┐
│ STEP 1 · Enter report identities                      (Card 1)   │
│                                                                  │
│   Source · Expected (Import)      Target · Actual (Direct Lake)  │
│     • Tenant ID                     • Tenant ID                  │
│     • Group / Workspace ID          • Group / Workspace ID       │
│     • Report ID                     • Report ID                  │
│     • Dataset ID                    • Dataset ID                 │
│                                                                  │
│   • Pair name  (e.g. Cygnus)                                     │
│   • Case-insensitive compare  (optional)                         │
└──────────────────────────────────────────────────────────────────┘
                                  │
                                  ▼
┌──────────────────────────────────────────────────────────────────┐
│ STEP 2 · Run Authentication                    (Card 2 button)   │
│                                                                  │
│   Opens a browser → sign in with the test account.               │
│   Saves your login session so the next steps don't re-ask.       │
│   Credentials come from the .env file, never typed in the app.   │
└──────────────────────────────────────────────────────────────────┘
                                  │
                                  ▼
                          ╱ Signed in? ╲
                          ◇            ◇──── No ──►  fix .env / try again
                          ╲            ╱
                                  │ Yes
                                  ▼
┌──────────────────────────────────────────────────────────────────┐
│ STEP 3 · Discover Filters                      (Card 3 button)   │
│                                                                  │
│   Read-only crawl of the report. It finds:                       │
│     • all report PAGES                                           │
│     • every SLICER / filter on those pages (and its values)      │
│                                                                  │
│   Then, in the results panel below:                              │
│     • tick which PAGES to include                                │
│     • pick the SLICER values you want applied                    │
│   Your picks are remembered for the next Run Parity.             │
└──────────────────────────────────────────────────────────────────┘
                                  │
                                  ▼
┌──────────────────────────────────────────────────────────────────┐
│ STEP 4 · Run Parity                            (Card 2 button)   │
│                                                                  │
│   Exports both reports with your chosen filters:                 │
│     Source (Import)        →  expected.xlsx                      │
│     Target (Direct Lake)   →  actual.xlsx                        │
│   Compares them and shows, right in the app:                     │
│     • PASS / REVIEW / FAIL verdict                               │
│     • plain-English summary + per-page breakdown                 │
│   Buttons open: expected.xlsx · actual.xlsx · parity-summary.xlsx │
└──────────────────────────────────────────────────────────────────┘
                                  │
                                  ▼
┌──────────────────────────────────────────────────────────────────┐
│       DONE — review the verdict; open the files for detail       │
└──────────────────────────────────────────────────────────────────┘
```

## Step notes

**1 · Report identities.** Fill in all four IDs for both the Source (Expected /
Import mode) and Target (Actual / Direct Lake) reports, plus a Pair name. Turn on
*Case-insensitive compare* only if you want labels that differ just by
case/spacing to be treated as equal.

**2 · Run Authentication.** Click **Run Authentication**. A browser opens; sign in
with the test account. The session is saved, so you won't be asked again on the
next runs. This must be done before Discover Filters — the app will warn you if
no session exists yet.

**3 · Discover Filters.** Click **Discover Filters**. It reads the report
(without changing anything) and lists every page and every slicer/filter with its
values. In the results panel, tick the pages you want and pick the slicer values
you want applied. Those choices are remembered and used by the next Run Parity.

**4 · Run Parity.** Click **Run Parity**. Both reports are exported with your
chosen filters (source → `expected.xlsx`, target → `actual.xlsx`), compared, and
the verdict (PASS / REVIEW / FAIL) plus a per-page breakdown appears in the app.
Use the buttons to open `expected.xlsx`, `actual.xlsx`, or `parity-summary.xlsx`
for the full detail.

## Shortcut for repeat runs

**Run Auth + Parity** does steps 2 and 4 back-to-back in one click — handy once
you've already used Discover Filters to set the pages and slicer values you want.
