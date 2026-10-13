"""Check that drawing-translate can run here; print what to install if not.

Usage: python check_deps.py            (check only)
       python check_deps.py --install  (pip install the missing packages into the user site)
       python check_deps.py --install --cn   (same, through the Tsinghua PyPI mirror; use in mainland China)
"""
import importlib.util
import os
import subprocess
import sys

PKGS = {"pymupdf": "pymupdf", "rapidocr_onnxruntime": "rapidocr_onnxruntime", "openpyxl": "openpyxl"}
NEEDED = {"pymupdf": "required", "rapidocr_onnxruntime": "for scans / drawings without a text layer",
          "openpyxl": "for the review spreadsheet"}
FONTS = [r"C:\Windows\Fonts\Deng.ttf", r"C:\Windows\Fonts\simhei.ttf"]  # same list as overlay.py
MIRROR = "https://pypi.tuna.tsinghua.edu.cn/simple"

missing = [m for m in PKGS if importlib.util.find_spec(m) is None]
for m in PKGS:
    print(f"{'OK     ' if m not in missing else 'MISSING'} {m:22} ({NEEDED[m]})")
font = next((f for f in FONTS if os.path.exists(f)), None)
print(f"{'OK     ' if font else 'WARN   '} Chinese font: {font or 'none found, falls back to the PDF built-in CJK font'}")

if missing and "--install" in sys.argv:
    cmd = [sys.executable, "-m", "pip", "install", "--user", *[PKGS[m] for m in missing]]
    if "--cn" in sys.argv:
        cmd += ["-i", MIRROR, "--timeout", "60"]
    print("running:", " ".join(cmd))
    sys.exit(subprocess.call(cmd))
elif missing:
    print("\nInstall with:\n  python -m pip install --user " + " ".join(PKGS[m] for m in missing)
          + f"\n(mainland China, faster:  add  -i {MIRROR})")
    sys.exit(1)
