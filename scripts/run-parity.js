/**
 * run-parity.js
 *
 * Cross-platform launcher for the report-parity (migration) validator.
 * Windows PowerShell does not support `MODE=source npx ...` inline env syntax,
 * and this project has no cross-env dependency — so we set the env var here in
 * Node and spawn Playwright. Works identically in PowerShell, CMD and bash.
 *
 * Usage (via npm scripts):
 *   npm run parity           → MODE=both
 *   npm run parity:source    → MODE=source
 *   npm run parity:target    → MODE=target
 *
 * Select a configured pair:
 *   PowerShell:  $env:PAIR='Cygnus'; npm run parity
 *   bash:        PAIR=Cygnus npm run parity
 */

const { spawnSync } = require('child_process');

const VALID_MODES = ['both', 'source', 'target'];

const mode = (process.argv[2] || 'both').toLowerCase();
if (!VALID_MODES.includes(mode)) {
  console.error(`[run-parity] Invalid mode "${mode}". Expected one of: ${VALID_MODES.join(', ')}`);
  process.exit(1);
}

// Any extra args after the mode are passed straight through to Playwright.
const passthrough = process.argv.slice(3);

const args = [
  'playwright', 'test',
  '--project=report-parity',
  '--no-deps',
  '--reporter=list,html',
  ...passthrough,
];

console.log(`[run-parity] MODE=${mode}${process.env.PAIR ? ` PAIR=${process.env.PAIR}` : ''}`);
console.log(`[run-parity] npx ${args.join(' ')}\n`);

const result = spawnSync('npx', args, {
  stdio: 'inherit',
  shell: true,                                   // required for npx on Windows
  env: { ...process.env, MODE: mode },
});

process.exit(result.status ?? 1);
