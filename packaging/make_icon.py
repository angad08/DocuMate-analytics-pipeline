"""
Turn logo\\documate.ico.png into a real logo\\documate.ico.

Windows will not accept a PNG as an exe icon, whatever the file is called,
and PyInstaller passes the file straight through - so a PNG named .ico.png
is rejected at build time rather than quietly working.

A .ico is a container holding the same picture at several sizes, and
Windows picks the one it needs: 16 for the taskbar, 32 for the window
corner, 256 for the large view in Explorer. Giving it only one size leaves
Windows to scale, which looks soft at small sizes.

Run it only when the logo changes:

    python packaging\\make_icon.py

Requirements:  pip install pillow
"""

import os
import sys


PACKAGING = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(PACKAGING)

SOURCE = os.path.join(ROOT, "logo", "documate.ico.png")
TARGET = os.path.join(ROOT, "logo", "documate.ico")

SIZES = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]


def main():
    try:
        from PIL import Image
    except ImportError:
        raise SystemExit("This needs Pillow:  pip install pillow")

    if not os.path.exists(SOURCE):
        raise SystemExit("No logo at " + SOURCE)

    # RGBA so a logo with transparency keeps it; a flat PNG is unharmed.
    image = Image.open(SOURCE).convert("RGBA")
    image.save(TARGET, format="ICO", sizes=SIZES)

    print("Wrote " + TARGET)
    print("Sizes: " + ", ".join(str(width) for width, _ in SIZES))


if __name__ == "__main__":
    sys.exit(main())
