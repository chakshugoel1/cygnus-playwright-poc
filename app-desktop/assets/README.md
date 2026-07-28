# App icon — single source of truth

**`cygnus.ico` is the one icon file the whole app uses.** Three places reference
it, all pointing at this same file:

| Where it shows up | Referenced from |
|-------------------|-----------------|
| Desktop shortcut  | `install.ps1` and `app-desktop/CreateDesktopShortcut.ps1` (`IconLocation`) |
| App window & taskbar (running app) | `app-desktop/main.js` (`BrowserWindow({ icon })`) |
| Packaged `.exe`   | `app-desktop/package.json` (`build.win.icon`) |

## To change the icon

Pick **one** of these — you never edit more than one place:

1. **Have a new picture?** Replace `cygnus-source.png` (a square PNG — 512×512
   recommended, transparent background) and run:
   ```
   cd app-desktop
   npm run make-icon
   ```
   That regenerates `cygnus.ico` (16 / 32 / 48 / 256 px) from the PNG. Done —
   every place above now uses the new icon.

2. **Already have a `.ico`?** Just overwrite `cygnus.ico` with it. Done.

> `npm run make-icon` needs Pillow once: `python -m pip install Pillow`.

## When each surface updates

- **App window / taskbar** — next time you launch the app.
- **Desktop shortcut** — next time the shortcut is created (rerun the installer's
  shortcut step, or `npm run make-shortcut`). Existing shortcuts don't
  auto-refresh; Windows also caches icons, so a sign-out/in may be needed.
- **Packaged `.exe`** — next `npm run dist`.

## Files here

- `cygnus.ico` — the icon everything uses (generated; committed).
- `cygnus-source.png` — the editable source picture the `.ico` is built from.
- `make-icon.py` — the PNG → ICO converter behind `npm run make-icon`.
