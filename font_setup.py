#!/usr/bin/env python3
"""
font_setup.py  —  Cross-platform font path resolver
Run this ONCE after cloning to verify fonts are available.
Usage:  python font_setup.py
"""

import os, sys, platform

FONT_CANDIDATES = {
    "linux": "/usr/share/fonts/truetype/liberation/",
    "darwin": "/Library/Fonts/",          # macOS
    "win32":  "C:/Windows/Fonts/",        # Windows
}

REQUIRED = [
    "LiberationSans-Regular.ttf",
    "LiberationSans-Bold.ttf",
    "LiberationSerif-Regular.ttf",
    "LiberationSerif-Bold.ttf",
    "LiberationSerif-Italic.ttf",
]

def find_font_dir():
    sys_key = sys.platform  # 'linux', 'darwin', 'win32'
    for key, path in FONT_CANDIDATES.items():
        if key in sys_key and os.path.isdir(path):
            # Check at least one required font exists
            if any(os.path.exists(os.path.join(path, f)) for f in REQUIRED):
                return path
    return None

def check_fonts():
    d = find_font_dir()
    if not d:
        print("❌  Liberation fonts NOT found on this system.")
        print()
        print("  Please install them:")
        if sys.platform == "darwin":
            print("    brew install --cask font-liberation")
        elif sys.platform == "win32":
            print("    Download from: https://github.com/liberationfonts/liberation-fonts/releases")
            print("    Then right-click each .ttf → Install")
        else:
            print("    sudo apt-get install fonts-liberation")
        sys.exit(1)

    missing = [f for f in REQUIRED if not os.path.exists(os.path.join(d, f))]
    if missing:
        print(f"⚠️  Font directory found: {d}")
        print(f"   Missing files: {missing}")
        sys.exit(1)

    print(f"✅  Liberation fonts found at: {d}")
    return d

def patch_font_dir(font_dir: str):
    """Patch FONT_DIR in invoice engine files to match this OS."""
    targets = ["invoice_infinitegz_v7.py", "invoice_ast.py"]
    for fname in targets:
        if not os.path.exists(fname):
            continue
        with open(fname, "r") as f:
            src = f.read()
        # Replace any existing FONT_DIR assignment
        import re
        new_src = re.sub(
            r'FONT_DIR\s*=\s*"[^"]*"',
            f'FONT_DIR = "{font_dir}"',
            src
        )
        if new_src != src:
            with open(fname, "w") as f:
                f.write(new_src)
            print(f"   ✏️  Patched FONT_DIR in {fname}")
        else:
            print(f"   ✔  {fname} already correct")

if __name__ == "__main__":
    print("=== Invoice System — Font Setup ===")
    font_dir = check_fonts()
    patch_font_dir(font_dir)
    print()
    print("All fonts OK. You can now run:  streamlit run app.py")
