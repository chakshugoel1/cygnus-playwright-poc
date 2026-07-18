/**
 * slicer-config.helpers.ts
 * ============================================================================
 *  DEFINES WHICH SLICER FILTERS TO APPLY DURING A RUN.
 * ============================================================================
 *
 * A "scenario" is a named set of slicer selections applied to specific pages
 * before data is exported. Example: { name: "Sales", pages: { CXO: [ ... ] } }.
 *
 * Two ways to provide scenarios:
 *
 *   1. Inline in code (SCENARIOS below) — version-controlled, good for a fixed
 *      regression set.
 *   2. A JSON file written by the Discover step / desktop GUI — good for ad-hoc,
 *      pick-and-run use. This is enabled by the same env-var override pattern
 *      the parity report-identity override already uses:
 *        CYGNUS_SLICER_OVERRIDE=1
 *        CYGNUS_SLICER_CONFIG_PATH=.runtime/slicer-scenarios.json
 *
 * If NEITHER is provided, runs behave exactly as before — no filters applied.
 * "None" is always represented by simply not listing a slicer for that page.
 *
 * CASCADING / HIERARCHY: selections within a page are an ORDERED list. They are
 * applied top-to-bottom, waiting for the report to redraw between each, so a
 * dependent slicer (e.g. Department under a BU/DU hierarchy) is set only after
 * its parents. You do not declare the dependency anywhere — you just list the
 * parent selection before the child, and Power BI narrows the child for you.
 */

import * as fs   from 'fs';
import * as path from 'path';

// ── One slicer selection on a page ───────────────────────────────────────────
export interface SlicerSelection {
  /**
   * Power BI visual NAME of a specific slicer — REQUIRED for hierarchy fields
   * (BU/DU/DU Detail style), since those can only be set via that exact
   * visual. OPTIONAL for flat fields: when omitted, the field is applied via
   * page.setFilters() using `targets` alone — no slicer visual needed at all,
   * which is what lets a field discovered on one page apply cleanly to a
   * different page that may not even have its own visible slicer for it.
   */
  visualName?: string;
  /** Human-readable slicer title — for logs/output only, not used to target. */
  title?: string;
  /** Value(s) to select. Multiple values = OR (Power BI "In"). */
  values: string[];
  /** True if this is a tree/hierarchy slicer (BU/DU style). REQUIRES visualName. */
  isHierarchy?: boolean;
  /**
   * The field(s) this slicer is bound to (from discovery's targetLabel data).
   * REQUIRED when visualName is omitted (field-based application). Optional,
   * supplementary info otherwise — also what enables global-filter batching
   * (see planFilterApplication).
   */
  targets?: SlicerTarget[];
}

/** A slicer's binding target — matches harness.helpers.ts's SlicerTarget shape. */
export interface SlicerTarget {
  table?:  string;
  column?: string;
  [key: string]: unknown;
}

// ── A named scenario: a set of page → ordered selections ─────────────────────
export interface SlicerScenario {
  /** Friendly name — used for output subfolder + summary labelling. */
  name: string;
  /**
   * Map of page display name → ordered slicer selections for that page.
   * A page not listed here (or listed with an empty array) gets no filter —
   * i.e. "None". Selections are applied in array order (cascade-safe) WITHIN
   * each kind — but report-parity.spec.ts's perPageHook applies ALL
   * visual-based selections (visualName set) before ANY field-based ones
   * (targets only, no visualName) on a given page, regardless of how they're
   * interleaved in this array. This is safe for the common case (a page's
   * field-based selections are independent of its own hierarchy cascade —
   * that's the whole point of the field-based path), but do NOT rely on a
   * field-based selection being applied strictly BETWEEN two visual-based
   * ones on the same page — it won't be.
   *
   * A selection can be either visual-based (`visualName` set — required for
   * hierarchy fields) or field-based (`targets` set, no `visualName` — flat
   * fields only, applied via page.setFilters(), no slicer visual needed on
   * that page at all). To apply one field's filter to several pages, repeat
   * the same {title, targets, values} entry under each page's key.
   */
  pages: Record<string, SlicerSelection[]>;
}

/* =============================================================================
 *  INLINE SCENARIOS (optional). Leave empty to rely solely on the JSON override
 *  / desktop GUI. Add entries here for a fixed, version-controlled regression
 *  set. The commented example shows a cascading CXO selection: BU/DU first
 *  (hierarchy), then Department (its child), then a flat Status slicer.
 * =============================================================================
 */
export const SCENARIOS: SlicerScenario[] = [
  // {
  //   name: 'CXO-BU1-Sales',
  //   pages: {
  //     CXO: [
  //       { visualName: '<bu-du-slicer-visual-name>', title: 'BU / DU', values: ['BU1'], isHierarchy: true },
  //       { visualName: '<department-slicer-visual-name>', title: 'Department', values: ['Sales'] },
  //       { visualName: '<status-slicer-visual-name>', title: 'Status', values: ['Active'] },
  //     ],
  //   },
  // },
];

// ── JSON override (Discover step / GUI writes this) ──────────────────────────

interface RuntimeSlicerConfig {
  scenarios?: SlicerScenario[];
}

function readRuntimeSlicerConfig(): SlicerScenario[] | null {
  const enabled = (process.env['CYGNUS_SLICER_OVERRIDE'] ?? '').trim() === '1';
  if (!enabled) return null;

  const configPath = (process.env['CYGNUS_SLICER_CONFIG_PATH'] ?? '').trim();
  if (!configPath) return null;

  const resolved = path.isAbsolute(configPath)
    ? configPath
    : path.join(process.cwd(), configPath);

  if (!fs.existsSync(resolved)) {
    throw new Error(`Slicer scenario config file not found: ${resolved}`);
  }

  const raw = fs.readFileSync(resolved, 'utf8');
  let parsed: RuntimeSlicerConfig;
  try {
    parsed = JSON.parse(raw) as RuntimeSlicerConfig;
  } catch (e) {
    throw new Error(`Slicer scenario config is not valid JSON: ${resolved}`);
  }

  const scenarios = parsed.scenarios ?? [];
  if (!Array.isArray(scenarios)) {
    throw new Error(`Slicer scenario config "scenarios" must be an array: ${resolved}`);
  }
  validateScenarios(scenarios, resolved);
  return scenarios;
}

/** Throws a clear error if the scenario list is malformed. */
export function validateScenarios(scenarios: SlicerScenario[], sourceLabel = 'inline SCENARIOS'): void {
  // Two scenarios that collide on name OR output slug would silently overwrite
  // each other's expected/actual/summary files mid-run (they share an output
  // folder — see scenarioSlug() below) — catch that here, before any browser work.
  const seenNames = new Map<string, number>(); // lowercased name -> first index
  const seenSlugs = new Map<string, number>();  // slug -> first index

  scenarios.forEach((sc, i) => {
    if (!sc.name || !sc.name.trim()) {
      throw new Error(`Scenario #${i + 1} in ${sourceLabel} is missing a "name".`);
    }

    const nameKey = sc.name.trim().toLowerCase();
    const firstNameIdx = seenNames.get(nameKey);
    if (firstNameIdx !== undefined) {
      throw new Error(
        `Scenario #${i + 1} ("${sc.name}") in ${sourceLabel} duplicates the name of scenario #${firstNameIdx + 1} — ` +
        `each scenario name must be unique (they share an output subfolder).`,
      );
    }
    seenNames.set(nameKey, i);

    const slug = scenarioSlug(sc.name);
    const firstSlugIdx = seenSlugs.get(slug);
    if (firstSlugIdx !== undefined) {
      throw new Error(
        `Scenario #${i + 1} ("${sc.name}") in ${sourceLabel} produces the same output folder ("${slug}") as ` +
        `scenario #${firstSlugIdx + 1} — rename one of them so their output doesn't collide.`,
      );
    }
    seenSlugs.set(slug, i);

    if (!sc.pages || typeof sc.pages !== 'object') {
      throw new Error(`Scenario "${sc.name}" in ${sourceLabel} has no "pages" object.`);
    }
    for (const [pageName, selections] of Object.entries(sc.pages)) {
      if (!Array.isArray(selections)) {
        throw new Error(`Scenario "${sc.name}", page "${pageName}": selections must be an array.`);
      }
      selections.forEach((sel, j) => {
        const label = sel.title ?? sel.visualName ?? `selection #${j + 1}`;
        if (!sel.visualName && (!sel.targets || sel.targets.length === 0)) {
          throw new Error(
            `Scenario "${sc.name}", page "${pageName}", ${label}: needs either "visualName" ` +
            `(visual-based) or "targets" (field-based) — neither was provided.`,
          );
        }
        if (sel.isHierarchy && !sel.visualName) {
          throw new Error(
            `Scenario "${sc.name}", page "${pageName}", ${label}: hierarchy selections require ` +
            `"visualName" — a hierarchy field can only be set through its own slicer visual.`,
          );
        }
        if (!Array.isArray(sel.values) || sel.values.length === 0) {
          throw new Error(
            `Scenario "${sc.name}", page "${pageName}", ${label}: "values" must be a non-empty array.`,
          );
        }
      });
    }
  });
}

// ── Public accessors ─────────────────────────────────────────────────────────

/**
 * Returns the active list of scenarios: the JSON override if enabled, else the
 * inline SCENARIOS, else empty. An empty result means "run unfiltered", which
 * is the original behaviour.
 */
export function getScenarios(): SlicerScenario[] {
  const override = readRuntimeSlicerConfig();
  if (override) return override;
  validateScenarios(SCENARIOS);
  return SCENARIOS;
}

/**
 * True if any filtering is configured at all. When false, callers should run
 * exactly the original single unfiltered pass.
 */
export function hasScenarios(): boolean {
  return getScenarios().length > 0;
}

/**
 * Returns the ordered slicer selections for a given page within a scenario,
 * or an empty array (meaning "None" for that page).
 */
/**
 * Case-insensitive lookup into a Record keyed by page display name. A page
 * name that doesn't exactly case-match (e.g. "cxo" vs "CXO") would otherwise
 * silently return nothing — no error, filter just never applies. Shared by
 * selectionsForPage below and by report-parity.spec.ts's lookup into a
 * FilterApplicationPlan's perPage map, which has the exact same shape/risk.
 */
export function getForPageCI<T>(record: Record<string, T[]>, pageDisplayName: string): T[] {
  const wanted = pageDisplayName.toLowerCase();
  const key = Object.keys(record).find(k => k.toLowerCase() === wanted);
  return key ? record[key] : [];
}

export function selectionsForPage(scenario: SlicerScenario, pageDisplayName: string): SlicerSelection[] {
  return getForPageCI(scenario.pages, pageDisplayName);
}

/** Filesystem-safe version of a scenario name, for output subfolders. */
export function scenarioSlug(name: string): string {
  return name.trim().replace(/[^a-zA-Z0-9._-]+/g, '_').replace(/^_+|_+$/g, '') || 'scenario';
}

// ── Global-filter batching plan ──────────────────────────────────────────────
//
// Opt-in (off by default — see CYGNUS_SLICER_GLOBAL_FILTERS below). When a
// scenario asks for the SAME field, with the SAME values, on 2+ pages, that
// can be applied ONCE via report.setFilters() instead of once per page via
// each page's own slicer visual — meaningful time savings on scenarios that
// repeat a filter across many pages. This function only PLANS the split; it
// does no I/O, so it's fully unit-testable without a live report.
//
// A selection is eligible for the global path only if:
//   - it has `targets` populated (discovery captures this; hand-written
//     configs that omit it just stay on the safe, always-correct per-page path)
//   - it is NOT a hierarchy field (Power BI's Hierarchy filter type is
//     slicer-only — report.setFilters() rejects it; confirmed against SDK source)
//   - the exact same (targets, values) pair appears on 2+ pages
//
// Anything not eligible is left untouched in the per-page plan — this
// function can only ever move work FROM per-page TO global, never drop it.

/** Whether the global-filter fast path is enabled for this run. Off by default. */
export function globalFiltersEnabled(): boolean {
  return (process.env['CYGNUS_SLICER_GLOBAL_FILTERS'] ?? '').trim() === '1';
}

export interface GlobalFilterGroup {
  targets: SlicerTarget[];
  values:  string[];
  pages:   string[]; // page display names this group covers
}

export interface FilterApplicationPlan {
  /** Field+value groups to apply once, at the report level. */
  global: GlobalFilterGroup[];
  /** Remaining selections, unchanged, to apply per-page as before. */
  perPage: Record<string, SlicerSelection[]>;
}

function targetKey(targets: SlicerTarget[]): string {
  // Order-independent-enough for our purposes: targets within one slicer are
  // already returned in a stable order by Power BI, so a straight join is fine.
  return targets.map(t => `${t.table ?? ''}.${t.column ?? ''}`).join('|');
}

function selectionKey(targets: SlicerTarget[], values: string[]): string {
  return targetKey(targets) + '::' + [...values].sort().join(',');
}

export function planFilterApplication(scenario: SlicerScenario): FilterApplicationPlan {
  // First pass: bucket every (targets, values) combo by its key, tracking
  // which pages want it and under which selection object.
  const groups = new Map<string, { targets: SlicerTarget[]; values: string[]; pages: string[] }>();

  for (const [pageName, selections] of Object.entries(scenario.pages)) {
    for (const sel of selections) {
      if (!sel.targets || sel.targets.length === 0) continue; // no target info — stays per-page
      if (sel.isHierarchy || sel.targets.length > 1) continue; // hierarchy — must stay per-page

      const key = selectionKey(sel.targets, sel.values);
      const existing = groups.get(key);
      if (existing) {
        if (!existing.pages.includes(pageName)) existing.pages.push(pageName);
      } else {
        groups.set(key, { targets: sel.targets, values: sel.values, pages: [pageName] });
      }
    }
  }

  // Only groups spanning 2+ pages are worth batching — a single-page field
  // gets no benefit from the global path (one application either way) and
  // stays on the simpler, more-tested per-page code path.
  const global: GlobalFilterGroup[] = [];
  const globalKeys = new Set<string>();
  for (const [key, g] of groups) {
    if (g.pages.length >= 2) {
      global.push({ targets: g.targets, values: g.values, pages: g.pages });
      globalKeys.add(key);
    }
  }

  // Second pass: build the per-page remainder, excluding anything now handled globally.
  const perPage: Record<string, SlicerSelection[]> = {};
  for (const [pageName, selections] of Object.entries(scenario.pages)) {
    perPage[pageName] = selections.filter(sel => {
      if (!sel.targets || sel.targets.length === 0) return true;
      if (sel.isHierarchy || sel.targets.length > 1) return true;
      return !globalKeys.has(selectionKey(sel.targets, sel.values));
    });
  }

  return { global, perPage };
}