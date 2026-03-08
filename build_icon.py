"""
build_icon.py
-------------
Generates a default icon.ico for FetchPro if no icon.ico exists.

To use a CUSTOM icon:
    Replace (or place) your icon.ico file next to this script before running
    the build. The build will use it automatically.

Supported input formats (auto-converted):
    icon.ico  – used directly
    icon.png  – converted to ICO
    icon.jpg  – converted to ICO
"""

import sys
from pathlib import Path

ICON_PATH = Path("icon.ico")
SIZES = [16, 32,48, 64, 128, 256]


def _build_default_icon() -> None:
    """Create a simple FetchPro logo as ICO using Pillow."""
    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError:
        print("[build_icon] Pillow not installed – skipping icon generation.")
        print("             Install with: pip install pillow")
        return

    images = []
    for size in SIZES:
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # Background circle – deep blue
        margin = max(1, size // 16)
        draw.ellipse(
            [margin, margin, size - margin, size - margin],
            fill=(30, 100, 220, 255),
        )

        # Arrow pointing down (download symbol)
        cx = size // 2
        arrow_w = size * 0.35
        arrow_h = size * 0.22
        shaft_w = size * 0.12

        # Shaft
        draw.rectangle(
            [cx - shaft_w / 2, size * 0.18, cx + shaft_w / 2, size * 0.58],
            fill=(255, 255, 255, 255),
        )
        # Arrowhead
        draw.polygon(
            [
                (cx - arrow_w / 2, size * 0.52),
                (cx + arrow_w / 2, size * 0.52),
                (cx, size * 0.52 + arrow_h),
            ],
            fill=(255, 255, 255, 255),
        )
        # Baseline (tray)
        tray_h = shaft_w
        draw.rectangle(
            [cx - arrow_w / 2, size * 0.74, cx + arrow_w / 2, size * 0.74 + tray_h],
            fill=(255, 255, 255, 255),
        )

        images.append(img)

    images[0].save(
        ICON_PATH,
        format="ICO",
        append_images=images[1:],
        sizes=[(s, s) for s in SIZES],
    )
    print(f"[build_icon] Default icon created → {ICON_PATH}")


def _convert_image_to_ico(src: Path) -> None:
    """Convert an existing PNG/JPG to ICO."""
    try:
        from PIL import Image
    except ImportError:
        print("[build_icon] Pillow not installed – cannot convert image.")
        sys.exit(1)

    img = Image.open(src).convert("RGBA")
    imgs = [img.resize((s, s), Image.LANCZOS) for s in SIZES]
    imgs[0].save(
        ICON_PATH,
        format="ICO",
        append_images=imgs[1:],
        sizes=[(s, s) for s in SIZES],
    )
    print(f"[build_icon] Converted {src} → {ICON_PATH}")


def main() -> None:
    if ICON_PATH.exists():
        print(f"[build_icon] Found existing {ICON_PATH} – using it as-is.")
        return

    for ext in ("png", "jpg", "jpeg"):
        candidate = Path(f"icon.{ext}")
        if candidate.exists():
            _convert_image_to_ico(candidate)
            return

    print("[build_icon] No icon file found – generating default icon.")
    _build_default_icon()


if __name__ == "__main__":
    main()
