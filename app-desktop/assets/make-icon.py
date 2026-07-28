#!/usr/bin/env python3
"""
Regenerates cygnus.ico from cygnus-source.png.

cygnus.ico is the SINGLE source of truth for the app's icon: the desktop
shortcut, the Electron app window/taskbar, and the packaged .exe all point at
that one file. To change the icon, replace cygnus-source.png (a square PNG,
512x512 recommended, with transparency) and run this script — or drop a
ready-made cygnus.ico here directly.

Usage (from app-desktop/):   npm run make-icon
Or directly:                 python assets/make-icon.py

Requires Pillow:             python -m pip install Pillow
"""
import os
import sys

try:
    from PIL import Image
except ImportError:
    sys.exit("Pillow is required. Install it with:  python -m pip install Pillow")

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "cygnus-source.png")
OUT = os.path.join(HERE, "cygnus.ico")

# Sizes Windows actually asks for: 16 (taskbar/small), 32 & 48 (desktop),
# 256 (large-icons / high-DPI). Each is rendered from the source with
# high-quality LANCZOS resampling and bundled into one multi-resolution .ico.
SIZES = [256, 48, 32, 16]


def main() -> None:
    if not os.path.exists(SRC):
        sys.exit(f"Source image not found: {SRC}")

    img = Image.open(SRC).convert("RGBA")
    if img.width != img.height:
        print(f"Warning: source is {img.width}x{img.height} (not square) — the "
              "icon may look stretched.", file=sys.stderr)

    frames = [img.resize((s, s), Image.LANCZOS) for s in SIZES]
    frames[0].save(OUT, format="ICO", sizes=[(s, s) for s in SIZES],
                   append_images=frames[1:])

    print(f"Wrote {OUT} from {os.path.basename(SRC)} "
          f"({img.width}x{img.height}) — sizes: {SIZES}")


if __name__ == "__main__":
    main()
