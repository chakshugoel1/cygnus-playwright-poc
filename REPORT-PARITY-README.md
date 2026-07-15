# Report Parity (Migration) Validator

Validates that a Power BI report migrated from **Import mode** → **Direct Lake mode**
still shows **identical data**.

```
source (Import mode)  → exported → expected.xlsx   ← EXPECTED values
target (Direct Lake)  → exported → actual.xlsx     ← ACTUAL values
                        compare  → parity-summary.xlsx
```

---

## 1. Configure your report pair (the only file you need to edit)

Open **`tests/helpers/comparison-config.helpers.ts`** and fill in the **target**
GUIDs. The Cygnus **source** is already seeded with your current report:

```ts
export const PAIRS: ReportPair[] = [
  {
    name: 'Cygnus',
    source: {   // Import mode — EXPECTED (already filled in)
      tenantId:  '8b87af7d-...',
      groupId:   '3f3f8c93-...',
      reportId:  '2ba857e0-...',
      datasetId: 'fbb612a3-...',
      rlsRole:   'DynamicRoles',
    },
    target: {   // Direct Lake — ACTUAL  ←←← REPLACE THESE
      tenantId:  '8b87af7d-...',                        // usually the same tenant
      groupId:   'REPLACE_WITH_TARGET_WORKSPACE_GUID',
      reportId:  'REPLACE_WITH_TARGET_REPORT_GUID',
      datasetId: 'REPLACE_WITH_TARGET_DATASET_GUID',
      rlsRole:   'DynamicRoles',
    },
  },
];
```

**Where to find the GUIDs:** open the migrated report in Power BI. The URL is
`app.powerbi.com/groups/<groupId>/reports/<reportId>/...`. The `datasetId` is on
the report's *Settings → Semantic model* page (or via the workspace's dataset list).

If you leave the placeholders in, the test **fails immediately with a clear
message** before opening a browser — it will not silently compare the wrong thing.

### Optional settings

```ts
pages: ['Employee', 'Manager'],       // only validate these tabs (default: ALL tabs)
pageMap: { 'DU Head': 'DU-Head' },    // only if a tab was RENAMED during migration
                                      // (source name → target name)
```

By default all pages are auto-discovered and matched by name — no config needed
if the tab names are unchanged.

---

## 2. Run it

```powershell
npm run parity          # export source + target, then compare   ← normal use
npm run parity:source   # export ONLY the source → refresh the expected baseline
npm run parity:target   # export ONLY the target → compare against existing baseline
```

`parity:target` is the fast loop: once you have a baseline, re-validate the
migrated report without re-exporting the Import-mode report each time. If no
baseline exists yet, it fails with an instruction to run `parity:source` first.

Multiple pairs configured? Select one:

```powershell
$env:PAIR='Cygnus'; npm run parity     # PowerShell
```

---

## 3. Output

Everything lands in **`playwright-report-parity/<PairName>/`**:

| File | Contents |
|---|---|
| `expected.xlsx` | Source (Import mode) data — one sheet per report page |
| `actual.xlsx` | Target (Direct Lake) data — one sheet per report page |
| `parity-summary.xlsx` | **The deliverable.** Verdict + per-page diffs + sample differing rows |

`parity-summary.xlsx` has three sheets:

- **Parity Summary** — PASS/FAIL verdict, which reports were compared, counts
- **Page Comparison** — per page: rows in source, rows in target, rows only on
  one side, and status (Identical / Different / Missing / Only in target)
- **Differences** — the actual differing rows, so you can see *what* changed

The test **fails** (red in Playwright) when the data does not match, so it works
in CI as a migration gate.

---

## Notes

- **Comparison is order-independent.** Rows and visuals are matched as a
  multiset, so a different row order or visual order will *not* be reported as a
  difference — only genuinely different data is.
- **Only data-bearing visuals are compared.** Images, shapes, text boxes and
  buttons are skipped (they carry no data).
- **Switching pairs?** Delete the old `playwright-report-parity/<Pair>/` folder,
  or the stale `expected.xlsx` will be reused.
- **The existing Cygnus flow is unchanged.** `npm run main:run` behaves exactly
  as before — the parity feature only activates on the `parity*` scripts.
- Auth is shared with the existing setup. If the token can't be found, run
  `npm run test:setup` first.
