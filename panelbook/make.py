"""Build the panel schedule workbook end to end.

    python panelbook/make.py            -> dist/Panel_Schedule_Tool.xlsx
    python panelbook/make.py --test     -> also adds the regression panels (not for release)

Steps: generate with openpyxl -> recalc with LibreOffice -> copy cached values back
into the openpyxl file (keeps its formatting; Excel still recalculates on open).
"""
import argparse
import pathlib
import subprocess
import sys

HERE = pathlib.Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE / "src"))
sys.path.insert(0, str(HERE / "tests"))
sys.path.insert(0, str(ROOT / "shared"))

import build  # noqa: E402
import recalc  # noqa: E402


def make(out, test=False):
    out = pathlib.Path(out).resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    raw = out.with_name(out.stem + "_raw.xlsx")
    calc = out.with_name(out.stem + "_calc.xlsx")
    extra = None
    if test:
        from test_panels import TEST_PANELS
        extra = TEST_PANELS
    build.main(str(raw), extra=extra)
    recalc.recalc(raw, calc)
    errs, n = recalc.find_errors(calc)
    subprocess.run([sys.executable, str(ROOT / "shared" / "inject.py"), str(raw), str(calc), str(out)], check=True)
    raw.unlink()
    print(f"{out.name}: {n} formula errors")
    for e in errs:
        print("  ", e)
    return out, calc, n


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=str(ROOT / "dist" / "Panel_Schedule_Tool.xlsx"))
    ap.add_argument("--test", action="store_true", help="include regression panels")
    a = ap.parse_args()
    _, calc, n = make(a.out, a.test)
    calc.unlink()
    sys.exit(1 if n else 0)
