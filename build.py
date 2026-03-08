"""
build.py – Local EXE builder for FetchPro
==========================================
Run from the project root:

    python build.py                 # builds with default / existing icon
    python build.py --icon my.ico   # builds with a specific ICO file
    python build.py --icon my.png   # auto-converts PNG → ICO, then builds

Requirements:
    pip install pyinstaller pillow
"""

import argparse
import shutil
import subprocess
import sys
from pathlib import Path


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _check_tool(name: str) -> None:
    if shutil.which(name) is None:
        print(f"[build] ERROR: '{name}' not found. Install with: pip install {name}")
        sys.exit(1)


def _resolve_icon(icon_arg: str | None) -> Path:
    """Return the ICO file to use for the build."""
    icon_path = Path("icon.ico")

    if icon_arg:
        src = Path(icon_arg)
        if not src.exists():
            print(f"[build] ERROR: icon file not found: {src}")
            sys.exit(1)
        if src.suffix.lower() != ".ico":
            print(f"[build] Converting {src} → icon.ico …")
            _convert_to_ico(src, icon_path)
        else:
            icon_path = src
    else:
        # Delegate to build_icon.py (handles existing file, PNG, or default)
        subprocess.run([sys.executable, "build_icon.py"], check=True)

    return icon_path


def _convert_to_ico(src: Path, dst: Path) -> None:
    try:
        from PIL import Image
    except ImportError:
        print("[build] ERROR: Pillow is required for image conversion.")
        print("        Install with: pip install pillow")
        sys.exit(1)

    sizes = [16, 32, 48, 64, 128, 256]
    img = Image.open(src).convert("RGBA")
    imgs = [img.resize((s, s), Image.LANCZOS) for s in sizes]
    imgs[0].save(dst, format="ICO", append_images=imgs[1:],
                 sizes=[(s, s) for s in sizes])
    print(f"[build] Saved {dst}")


# ---------------------------------------------------------------------------
# Main build
# ---------------------------------------------------------------------------

def build(icon: Path, output_name: str = "FetchPro") -> None:
    _check_tool("pyinstaller")

    dist_dir = Path("dist")
    build_dir = Path("build")

    cmd = [
        "pyinstaller",
        "--onefile",
        "--windowed",
        f"--icon={icon}",
        f"--name={output_name}",
        "--add-data", f"{icon};.",
        "fetchpro.py",
    ]

    print("[build] Running PyInstaller …")
    print("        " + " ".join(cmd))
    result = subprocess.run(cmd)

    if result.returncode != 0:
        print("[build] ERROR: PyInstaller failed.")
        sys.exit(result.returncode)

    exe = dist_dir / f"{output_name}.exe"
    if exe.exists():
        print(f"\n[build] ✓ Build successful: {exe.resolve()}")
    else:
        print("[build] WARNING: EXE not found at expected path.")

    # Clean up PyInstaller temp files
    spec_file = Path(f"{output_name}.spec")
    if spec_file.exists():
        spec_file.unlink()
    if build_dir.exists():
        shutil.rmtree(build_dir)
    print("[build] Cleaned up build artefacts.")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Build FetchPro EXE with a custom icon."
    )
    parser.add_argument(
        "--icon",
        metavar="FILE",
        help="Path to an .ico or .png/.jpg file to use as the application icon. "
             "If omitted, icon.ico (or a generated default) is used.",
    )
    parser.add_argument(
        "--name",
        default="FetchPro",
        help="Output EXE filename without extension (default: FetchPro).",
    )
    args = parser.parse_args()

    icon_path = _resolve_icon(args.icon)
    build(icon_path, output_name=args.name)


if __name__ == "__main__":
    main()
