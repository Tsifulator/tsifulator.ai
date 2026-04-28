#!/usr/bin/env python3
"""
Generate the menu bar template icon for tsifl Helper.app.

macOS menu bar icons are 'template images' — black + transparent only,
auto-tinted by the OS based on dark/light mode. 22x22 @1x, 44x44 @2x.

We render a chunky lowercase 't' in a bold sans font. Recognizable at
small sizes, has visual weight (won't get clipped behind the notch like
a thin glyph).

Run with:  python3 _make_icon.py
Output:    icon.png (44x44, black-on-transparent)
"""

from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

OUTPUT = Path(__file__).resolve().parent / "icon.png"

# 44x44 = retina @2x. macOS auto-downscales to 22x22 on @1x displays.
SIZE = 44


def _find_font():
    """Find a bold system sans font. Falls back through plausible paths."""
    candidates = [
        "/System/Library/Fonts/Helvetica.ttc",
        "/System/Library/Fonts/HelveticaNeue.ttc",
        "/System/Library/Fonts/Avenir Next.ttc",
        "/System/Library/Fonts/SFNS.ttf",
        "/Library/Fonts/Arial Bold.ttf",
    ]
    for p in candidates:
        if Path(p).exists():
            try:
                return ImageFont.truetype(p, SIZE - 6, index=0)
            except Exception:
                continue
    # Last resort
    return ImageFont.load_default()


def _draw_t(fill):
    """Render a centered bold 't' on a transparent SIZE×SIZE canvas with the
    given fill color tuple. Returns the PIL.Image object."""
    img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    font = _find_font()
    text = "t"
    bbox = draw.textbbox((0, 0), text, font=font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]
    x = (SIZE - text_w) / 2 - bbox[0]
    y = (SIZE - text_h) / 2 - bbox[1]
    draw.text((x, y), text, font=font, fill=fill)
    return img


def main():
    # 1. Black template icon — for the menu bar (macOS auto-tints based on
    # light/dark mode).
    black_img = _draw_t((0, 0, 0, 255))
    black_img.save(OUTPUT, "PNG")
    print(f"Wrote {OUTPUT} ({SIZE}x{SIZE}) — black template (menu bar)")

    # 2. Blue icon — for the floating shortcut panel logo. Uses tsifl's
    # brand blue (#0D5EAF). The panel fills its content with the blue
    # version directly — no template tinting needed.
    blue_path = OUTPUT.parent / "icon_blue.png"
    blue_img = _draw_t((13, 94, 175, 255))
    blue_img.save(blue_path, "PNG")
    print(f"Wrote {blue_path} ({SIZE}x{SIZE}) — tsifl blue (panel logo)")


if __name__ == "__main__":
    main()
