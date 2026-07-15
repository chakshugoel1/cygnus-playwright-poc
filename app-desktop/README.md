# Cygnus Desktop Runner

Desktop wrapper for existing scripts:
- `npm run test:setup`
- `npm run parity`

## Key guarantees

- Login credentials are NOT collected in UI.
- Existing auth flow remains `.env`-driven.
- Existing CLI flow remains unchanged.
- Runtime report overrides are opt-in only and activated only when:
  - `CYGNUS_UI_RUNTIME_OVERRIDE=1`
  - `CYGNUS_UI_RUNTIME_CONFIG_PATH` points to JSON config.

## Run (dev)

```powershell
cd app-desktop
npm install
npm start
```

## Launch without npm start

Double-click [launch-desktop.bat](launch-desktop.bat) to open the UI directly.

To create a desktop shortcut:

```powershell
cd app-desktop
npm run make-shortcut
```

## Reversibility

To disable UI-driven overrides immediately:

1. Stop using desktop app and run CLI commands directly.
2. Ensure env vars are not set:
   - `CYGNUS_UI_RUNTIME_OVERRIDE`
   - `CYGNUS_UI_RUNTIME_CONFIG_PATH`

With those vars unset, `tests/helpers/comparison-config.helpers.ts` falls back to existing static `PAIRS` behavior.
