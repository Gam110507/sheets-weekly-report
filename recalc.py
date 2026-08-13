"""
Recalculate the workbook so its formulas carry cached values.

    python3 recalc.py

Why this step exists, and why it is not optional:

openpyxl writes formulas as text. It does not evaluate them, because it is not
a spreadsheet engine — it has no idea what `SUMIFS` means. So the file that
comes out of `build_workbook.py` contains 11,270 correct formulas and not one
number.

That is fine for Excel and Google Sheets, which calculate on open. It is fatal
for `verify_workbook.py`, which reads the file with `data_only=True` and
therefore sees only cached results. With no cache, every check reads `None`,
and the verifier reports seventeen failures against a workbook that is
perfectly correct.

So this hands the file to LibreOffice, which is a real spreadsheet engine,
tells it to recalculate everything, and saves the results back. After this the
workbook holds both the formulas AND the values they evaluate to, which is also
what a client wants: a file that shows its numbers the moment it opens, rather
than one that flashes zeros while it thinks.

The order is: build -> recalc -> verify. Skipping the middle step is the single
easiest way to conclude this project is broken when it is not.

Requires LibreOffice:
    Ubuntu/Debian   sudo apt install libreoffice-calc
    macOS           brew install --cask libreoffice
    Windows         https://www.libreoffice.org/download/
"""

import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

WB = Path("weekly-report-template.xlsx")

# LibreOffice ships under three names depending on the platform and packaging.
CANDIDATES = ["soffice", "libreoffice"]
MAC_DEFAULT = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")


def find_soffice() -> str:
    for name in CANDIDATES:
        found = shutil.which(name)
        if found:
            return found
    if MAC_DEFAULT.exists():
        return str(MAC_DEFAULT)
    sys.exit(
        "ERROR: LibreOffice not found.\n"
        "  This step needs a real spreadsheet engine to evaluate the formulas;\n"
        "  openpyxl cannot do it. Install LibreOffice and run this again:\n"
        "    Ubuntu/Debian  sudo apt install libreoffice-calc\n"
        "    macOS          brew install --cask libreoffice\n"
        "    Windows        https://www.libreoffice.org/download/\n"
        "  Alternatively: open the file in Excel or Google Sheets and save it.\n"
        "  That has the same effect — the formulas get calculated and cached.")


def main() -> int:
    if not WB.exists():
        sys.exit(f"ERROR: {WB} not found. Run `python3 build_workbook.py` first.")

    soffice = find_soffice()

    # Convert into a scratch directory rather than in place. LibreOffice writes
    # the output next to where it is told, and pointing it at the source folder
    # risks it clobbering the input halfway through and leaving nothing behind.
    with tempfile.TemporaryDirectory() as tmp:
        proc = subprocess.run(
            [soffice, "--headless", "--norestore",
             "--convert-to", "xlsx:Calc MS Excel 2007 XML",
             "--outdir", tmp, str(WB)],
            capture_output=True, text=True, timeout=300)

        produced = Path(tmp) / WB.name
        if proc.returncode != 0 or not produced.exists():
            print(proc.stdout, proc.stderr, sep="\n", file=sys.stderr)
            sys.exit("ERROR: LibreOffice could not recalculate the workbook.\n"
                     "  If it is already running, close it and try again — it "
                     "refuses to start a second instance in headless mode.")

        shutil.copyfile(produced, WB)

    # Prove it worked rather than assuming. A silent no-op here would push the
    # failure downstream into verify_workbook.py, where it looks like the
    # numbers are wrong rather than absent.
    from openpyxl import load_workbook
    cached = load_workbook(WB, data_only=True)["Weekly Report"]["B6"].value
    if cached is None:
        sys.exit("ERROR: the workbook still has no cached values after recalc.\n"
                 "  Open it in Excel or Google Sheets, save, and try verify again.")

    print(f"recalculated {WB}")
    print(f"  revenue, latest week now cached as {cached}")
    print("  next: python3 verify_workbook.py")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
