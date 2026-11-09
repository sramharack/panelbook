"""shared v1.0 (panelbook 2.0.0). Recalculate every formula in an .xlsx with headless LibreOffice, then report formula errors.

usage: python shared/recalc.py in.xlsx out.xlsx
Exit code 1 if any cell evaluates to an Excel error (#REF!, #VALUE!, #N/A, ...).
Needs LibreOffice (`soffice` on PATH, or set SOFFICE=/path/to/soffice).
"""
import os
import pathlib
import shutil
import subprocess
import sys
import tempfile

import openpyxl

# Fresh LibreOffice profile that always recalculates on load (default is "never" for xlsx).
XCU = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>"""
ERRORS = ("#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "#N/A", "#NUM!", "#NULL!")


def soffice():
    exe = os.environ.get("SOFFICE") or shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        for p in [r"C:\Program Files\LibreOffice\program\soffice.exe",
                  "/Applications/LibreOffice.app/Contents/MacOS/soffice"]:
            if os.path.exists(p):
                exe = p
    if not exe:
        sys.exit("LibreOffice not found: install it or set SOFFICE to the soffice executable.")
    return exe


def recalc(src, dst):
    src, dst = pathlib.Path(src).resolve(), pathlib.Path(dst).resolve()
    with tempfile.TemporaryDirectory() as tmp:
        prof = pathlib.Path(tmp, "profile")
        (prof / "user").mkdir(parents=True)
        (prof / "user" / "registrymodifications.xcu").write_text(XCU, encoding="utf-8")
        out = pathlib.Path(tmp, "out")
        out.mkdir()
        cmd = [soffice(), f"-env:UserInstallation={prof.as_uri()}", "--headless", "--calc",
               "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(out), str(src)]
        subprocess.run(cmd, check=True, capture_output=True, timeout=900)
        shutil.move(str(out / (src.stem + ".xlsx")), dst)
    return dst


def find_errors(path, limit=20):
    wb = openpyxl.load_workbook(path, data_only=True)
    found = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value in ERRORS:
                    found.append(f"{ws.title}!{c.coordinate} {c.value}")
    return found[:limit], len(found)


if __name__ == "__main__":
    recalc(sys.argv[1], sys.argv[2])
    errs, n = find_errors(sys.argv[2])
    print(f"recalculated -> {sys.argv[2]}  ({n} formula errors)")
    for e in errs:
        print("  ", e)
    sys.exit(1 if n else 0)
